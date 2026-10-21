/*---------------------------------------------------------------------------------------------
 *  Copyright (c) Microsoft Corporation. All rights reserved.
 *  Licensed under the MIT License. See License.txt in the project root for license information.
 *--------------------------------------------------------------------------------------------*/

import * as vscode from 'vscode';
import { randomBytes, createHash, randomUUID } from 'node:crypto';
import { basename } from 'node:path';
import { writeFile, rename, unlink } from 'node:fs/promises';
import { EngineClient } from './client.js';
import { activityRegion, type ActivityRegion } from './agentActivity.js';
import type { Kind, State, Command, DocumentReview, DocumentReviewChanges, DraftViewSection } from './types.js';
import { registerDocumentTools, type DocumentToolInput } from './documentTools.js';
import { activeAgentRequest, activeAgentLabel, DocumentHistory, withAgentToolRequest } from './agentRequest.js';
import { prepareWordFormatting, canRebaseTextEdits } from './wordFormatting.js';
import { DocumentVersions } from './documentVersions.js';
import { DocumentDrafts } from './documentDrafts.js';
import { discoverCommands, boundedCommandResult, projectCommandResult } from './documentCapabilities.js';

interface View { state: State; pages: number[][]; commands?: Command[] }
class Document implements vscode.CustomDocument {
	readonly client: EngineClient;
	readonly panels = new Set<vscode.WebviewPanel>();
	readonly history = new DocumentHistory();
	readonly drafts = new DocumentDrafts();
	draftAgentLabel = 'Agent';
	draftWorkingId?: string;
	draftWorkingUntil = 0;
	draftRequest?: string;
	reviewRequest?: string;
	reviewVersion = 0;
	review?: DocumentReview;
	readonly versions?: DocumentVersions;
	readonly versionRequests = new Set<string>();
	view!: View;
	latestBytes!: Uint8Array;
	agentActivity?: { region?: ActivityRegion; updated: number };
	commandResult?: { id: string; version: string; value: unknown };
	readonly commandResults = new Map<string, { id: string; version: string; value: unknown; size: number }>();
	private snapshotRevision = -1;
	password?: string;
	baseHash = '';
	dirty = false;
	saving = false;
	serial = Promise.resolve();
	instance = randomUUID();
	activity?: { id: string; ids: string[]; placements: Record<string, 'before' | 'after'>; insertions: DraftViewSection[]; revision: number; request?: string; label: string; intent: string; expiresAt: number; rects: State['rects']; timer: ReturnType<typeof setTimeout> };
	constructor(readonly uri: vscode.Uri, readonly kind: Kind, readonly context: vscode.ExtensionContext, private readonly onDispose: () => void) {
		this.client = new EngineClient(vscode.Uri.joinPath(context.extensionUri, 'out', 'engineWorker.cjs').fsPath);
		if (kind === 'word' && context.globalStorageUri) { this.versions = new DocumentVersions(context.globalStorageUri.fsPath, uri.toString()); }
	}
	async load(bytes: Uint8Array) {
		this.instance = randomUUID();
		this.history.clear();
		this.commandResult = undefined; this.commandResults.clear();
		this.drafts.discard();
		this.reviewRequest = undefined;
		this.reviewVersion++;
		this.review = undefined;
		this.versionRequests.clear();
		this.view = await this.client.request<View>('open', { kind: this.kind, name: basename(this.uri.path), bytes, extensionsRoot: vscode.Uri.joinPath(this.context.extensionUri, '..').fsPath, password: this.password, sourceUri: this.uri.toString(), artifactRoot: this.context.globalStorageUri ? vscode.Uri.joinPath(this.context.globalStorageUri, 'pdf-artifacts').fsPath : undefined });
		this.latestBytes = bytes;
		this.snapshotRevision = this.view.state.revision;
	}
	enqueue<T>(operation: () => Promise<T>): Promise<T> {
		const task = this.serial.then(operation);
		this.serial = task.then(() => { }, () => { });
		return task;
	}
	async saved() {
		if (this.kind !== 'pdf') { return; }
		this.view = { ...this.view, ...await this.client.request<View>('saved') };
		this.snapshotRevision = this.view.state.revision;
		await this.broadcast();
	}
	async snapshot(): Promise<Uint8Array> {
		if (this.kind === 'word' && this.snapshotRevision !== this.view.state.revision) {
			this.latestBytes = await this.client.request<Uint8Array>('export');
			this.snapshotRevision = this.view.state.revision;
		}
		return this.latestBytes;
	}
	async broadcast(command?: string) {
		await Promise.all([...this.panels].filter(panel => this.kind !== 'pdf' || panel.visible !== false).map(panel => panel.webview.postMessage({ type: 'state', kind: this.kind, ...this.view, ...(this.kind === 'pdf' ? { state: { ...this.view.state, pdfNewArtifacts: undefined, pdfArtifacts: this.view.state.pdfArtifacts?.slice(0, 20), pdfArtifactCount: this.view.state.pdfArtifacts?.length || 0 } } : {}), name: basename(this.uri.path), command })));
		if (this.kind === 'word') {
			const changes = this.review?.changes.length ? await this.client.request<DocumentReview['changes']>('reviewLayout', { changes: this.review.changes }) : [];
			await Promise.all([...this.panels].filter(panel => this.kind !== 'pdf' || panel.visible !== false).map(panel => panel.webview.postMessage({ type: 'review', review: { ...this.review, version: this.version, reviewVersion: this.reviewVersion, changes } })));
			const drafts = await Promise.all(this.drafts.pending().map(async section => ({ ...section, anchorRects: await this.client.request<State['rects']>('agentRects', { blockId: section.anchor.id }).catch(() => []) })));
			await Promise.all([...this.panels].filter(panel => this.kind !== 'pdf' || panel.visible !== false).map(panel => panel.webview.postMessage({ type: 'draft', sections: drafts, label: this.draftAgentLabel, request: this.draftRequest, workingId: this.draftWorkingUntil > Date.now() ? this.draftWorkingId || drafts[0]?.id : undefined, expiresAt: this.draftWorkingUntil })));
		}
		if (this.activity) { await this.sendActivity(this.activity.rects); }
	}
	async sendActivity(rects: State['rects']) {
		const activity = this.activity;
		if (activity && activity.revision !== this.view.state.revision) {
			const revision = this.view.state.revision;
			const located = await this.locateActivity(activity.ids, activity.placements, activity.intent);
			rects = located.rects;
			if (this.activity !== activity || revision !== this.view.state.revision) { return; }
			activity.rects = rects; activity.insertions = located.insertions; activity.revision = revision;
		}
		await Promise.all([...this.panels].filter(panel => this.kind !== 'pdf' || panel.visible !== false).map(panel => panel.webview.postMessage({ type: 'agentActivity', rects: rects || [], label: activity?.label || activeAgentLabel(), ...(activity ? { insertions: activity.insertions, intent: activity.intent, expiresAt: activity.expiresAt, request: activity.request, revision: activity.revision } : {}) })));
	}
	async stopActivity(force = false) { if (!force && activeAgentRequest() && this.activity?.request && this.activity.request !== activeAgentRequest()) { return; } if (this.activity) { clearTimeout(this.activity.timer); this.activity = undefined; } await this.sendActivity([]); }
	async renewActivity(request: string) {
		if (this.activity?.request === request) {
			clearTimeout(this.activity.timer); this.activity.expiresAt = Date.now() + 30000;
			this.activity.timer = setTimeout(() => { void this.stopActivity(); }, 30000);
			await this.sendActivity(this.activity.rects);
		}
		if (this.draftRequest === request && this.draftWorkingUntil > 0 && this.drafts.pending().length) { this.draftWorkingUntil = Date.now() + 30000; await this.broadcast(); }
	}
	async locateActivity(ids: string[], placements: Record<string, 'before' | 'after'>, intent: string) {
		const rects: NonNullable<State['rects']> = [];
		const insertions: DraftViewSection[] = [];
		for (const id of ids) {
			const anchorRects = await this.client.request<NonNullable<State['rects']>>('agentRects', { blockId: id }).catch(() => []);
			const position = placements[id];
			if (position) { insertions.push({ id: `activity:${id}:${position}`, title: '', intent, anchor: { id, position }, anchorRects }); }
			else { rects.push(...anchorRects); }
		}
		return { rects: rects.slice(0, 200), insertions };
	}
	async startActivity(id: string, intent?: string, targets = [id], placements: Record<string, 'before' | 'after'> = {}) {
		const preview = intent ?? (this.activity?.ids.includes(id) && this.activity.request === activeAgentRequest() ? this.activity.intent : 'Updating this block…');
		const ids = [...new Set(targets)].slice(0, 200);
		const { rects, insertions } = await this.locateActivity(ids, placements, preview);
		await this.stopActivity(true);
		this.activity = { id, ids, placements, insertions, revision: this.view.state.revision, request: activeAgentRequest(), label: activeAgentLabel(), intent: preview, expiresAt: Date.now() + 30000, rects, timer: setTimeout(() => { void this.stopActivity(); }, 30000) };
		await this.sendActivity(rects);
	}
	get version() { return `${this.instance}:${this.view.state.revision}`; }
	/**
	 * Reports agent activity over the region a command touched. The region is
	 * derived internally from the command and the engine state, so agents never
	 * pass ids or coordinates. A short window keeps the previous region while an
	 * agent runs a sequence of commands on the same block.
	 */
	async announceActivity(command: string, args: Record<string, unknown>, before: State, after: State) {
		const now = Date.now();
		let region = activityRegion(this.kind, args, before, after);
		if (!region && this.agentActivity?.region && now - this.agentActivity.updated < 3000) { region = this.agentActivity.region; }
		this.agentActivity = { region, updated: now };
		if (!region || !this.panels.size) { return; }
		await Promise.all([...this.panels].filter(panel => this.kind !== 'pdf' || panel.visible !== false).map(panel => panel.webview.postMessage({ type: 'agent', label: command, region })));
	}
	dispose() { if (this.activity) { clearTimeout(this.activity.timer); } this.onDispose(); this.client.dispose(); }
}
const hash = (bytes: Uint8Array) => createHash('sha256').update(bytes).digest('hex');
function checkAgentAccess(token?: vscode.CancellationToken) {
	if (token?.isCancellationRequested) { throw new vscode.CancellationError(); }
	if (!vscode.workspace.isTrusted) { throw new Error('Trust this workspace before allowing agent edits.'); }
}
async function read(uri: vscode.Uri) {
	if ((await vscode.workspace.fs.stat(uri)).size > 128 * 1024 * 1024) { throw new Error('Document exceeds the 128 MB limit.'); }
	const bytes = await vscode.workspace.fs.readFile(uri);
	if (bytes.length > 128 * 1024 * 1024) { throw new Error('Document exceeds the 128 MB limit.'); }
	return bytes;
}
/** Resolve image files in the host; binary data never needs to pass through model context. */
async function prepareImageEdits(input: DocumentToolInput): Promise<DocumentToolInput> {
	const edits = input.operation === 'batch' ? input.edits as Record<string, unknown>[] : [input];
	const cache = new Map<string, string>(); let total = 0;
	const prepared = [];
	for (const edit of edits) {
		const params = edit.params as Record<string, unknown> | undefined;
		if (edit.operation !== 'command' || !['insert.picture', 'picture.change'].includes(String(edit.command)) || params?.fileUri === undefined) { prepared.push(edit); continue; }
		if (typeof params.fileUri !== 'string' || params.data !== undefined || params.path !== undefined) { throw new Error('Use one image fileUri, without data or path.'); }
		let data = cache.get(params.fileUri);
		if (!data) {
			const uri = vscode.Uri.parse(params.fileUri);
			if (!['file', 'vscode-remote'].includes(uri.scheme)) { throw new Error('Image fileUri must identify a local or workspace file, not a web address.'); }
			if ((await vscode.workspace.fs.stat(uri)).size > 16 * 1024 * 1024) { throw new Error('Image exceeds the 16 MB limit.'); }
			const bytes = await vscode.workspace.fs.readFile(uri); total += bytes.length;
			if (total > 16 * 1024 * 1024) { throw new Error('Image batch exceeds the 16 MB limit.'); }
			data = Buffer.from(bytes).toString('base64'); cache.set(params.fileUri, data);
		}
		const { fileUri: _fileUri, ...rest } = params; prepared.push({ ...edit, params: { ...rest, data } });
	}
	return input.operation === 'batch' ? { ...input, edits: prepared } : prepared[0] as DocumentToolInput;
}
async function writeAtomic(uri: vscode.Uri, bytes: Uint8Array) {
	// workspace.fs.rename(overwrite) emits a DELETE operation for the destination,
	// which disposes VS Code's open custom-document model and its undo history.
	if (uri.scheme !== 'file') { await vscode.workspace.fs.writeFile(uri, bytes); return; }
	const temp = vscode.Uri.joinPath(uri, '..', '.' + basename(uri.path) + '.' + randomBytes(8).toString('hex') + '.tmp');
	try { await writeFile(temp.fsPath, bytes, { flag: 'wx' }); await rename(temp.fsPath, uri.fsPath); }
	finally { await unlink(temp.fsPath).catch(() => { }); }
}
export function registerEditors(context: vscode.ExtensionContext, kinds: Kind[]) {
	const documents = new Map<string, Document>();
	const reviewData = async (doc: Document): Promise<DocumentReview> => {
		const result = await doc.client.request<DocumentReviewChanges>('reviewRead');
		return { uri: doc.uri.toString(), version: doc.version, reviewVersion: doc.reviewVersion, changes: result.changes || [] };
	};
	const captureVersion = async (doc: Document, label: string) => {
		if (!doc.versions) { return; }
		try {
			const bytes = await doc.snapshot();
			const projection = await doc.client.request<string>('reviewProjection');
			const review = await doc.client.request<string>('reviewExport');
			await doc.versions.capture(bytes, projection, label, doc.review?.changes.length ? { review, request: doc.reviewRequest } : undefined);
		} catch { void vscode.window.showErrorMessage?.('The document is still open, but its local version history could not be saved. Check available disk space.'); }
	};
	const beforeRequest = async (doc: Document, request?: string) => {
		if (!request || doc.versionRequests.has(request)) { return; }
		await captureVersion(doc, 'Before Agent Edits'); doc.versionRequests.add(request);
		if (doc.versionRequests.size > 30) { doc.versionRequests.delete(doc.versionRequests.values().next().value!); }
	};
	const publishReview = async (doc: Document, request?: string) => {
		if (doc.kind !== 'word' || !request) { return; }
		doc.reviewRequest = request;
		doc.reviewVersion++;
		const data = await reviewData(doc);
		doc.review = data;
		const changes = await doc.client.request<DocumentReview['changes']>('reviewLayout', { changes: data.changes });
		await Promise.all([...doc.panels].map(panel => panel.webview.postMessage({ type: 'review', review: { ...data, changes } })));
		await captureVersion(doc, data.changes.length ? 'Agent Edits' : 'Review Completed');
		try { await vscode.commands.executeCommand('aiOffice.registerDocumentReview', request, data); }
		catch { void vscode.window.showErrorMessage?.('Document changes were applied, but change review could not open. Reload Sangam One before continuing.'); }
	};
	const agentEdits = new Map<Kind, (doc: Document, input: DocumentToolInput, token: vscode.CancellationToken, draftSection?: string) => Promise<unknown>>();
	const reviewEdits = new Map<Kind, (doc: Document, before: State) => void>();
	// The file an agent edits when none is named: the document whose editor is
	// active, or the one most recently active among the open documents.
	const lastActive = new Map<Kind, Document>();
	const requestDocuments = new Map<string, string>();
	const focusedReviewDocument = () => {
		const input = vscode.window.tabGroups?.activeTabGroup?.activeTab?.input as { uri?: vscode.Uri; modified?: vscode.Uri } | undefined;
		const resource = input?.uri || input?.modified;
		return resource?.scheme === 'ai-office-review' ? documents.get(resource.query) : undefined;
	};
	for (const kind of kinds) {
		const toolbarPanels = new Map<string, { doc: Document; panel: vscode.WebviewPanel }>();
		context.subscriptions.push(vscode.commands.registerCommand(`_aiOffice.${kind}.requestToolbar`, async (uri: vscode.Uri) => {
			const target = [...toolbarPanels.values()].find(({ doc, panel }) => doc.uri.toString() === uri.toString() && panel.active);
			if (target) { await target.panel.webview.postMessage({ type: 'publishToolbar' }); }
		}), vscode.commands.registerCommand(`_aiOffice.${kind}.toolbarAction`, async (uri: vscode.Uri, panelId: string, instance: string, context: string, action: object) => {
			const target = toolbarPanels.get(panelId);
			if (!target || target.doc.uri.toString() !== uri.toString() || target.doc.instance !== instance || !target.panel.active) { return false; }
			return target.doc.enqueue(async () => {
				if (!target.panel.active || target.doc.instance !== instance) { return false; }
				if (kind === 'word') { target.panel.reveal(target.panel.viewColumn, false); }
				return target.panel.webview.postMessage({ type: 'toolbarAction', context, instance, action });
			});
		}));
		const changed = new vscode.EventEmitter<vscode.CustomDocumentEditEvent<Document>>();
		context.subscriptions.push(changed);
		const recordEdit = (doc: Document, before: State, label: string, request?: string): void => {
			if (request) { doc.history.checkpoint(request); }
			const steps = Math.max(1, (doc.view.state.undoDepth ?? 0) - (before.undoDepth ?? 0));
			doc.history.record(steps, request);
			const position = doc.history.cursor;
			changed.fire({
				document: doc, label,
				undo: () => doc.enqueue(async () => {
					for (let step = 0; step < steps; step++) { await execute(doc, kind === 'pdf' ? 'edit_undo' : 'edit.undo', {}, false); }
					doc.history.cursor = position - 1;
				}),
				redo: () => doc.enqueue(async () => {
					for (let step = 0; step < steps; step++) { await execute(doc, kind === 'pdf' ? 'edit_redo' : 'edit.redo', {}, false); }
					doc.history.cursor = position;
				})
			});
		};
		reviewEdits.set(kind, (doc, before) => recordEdit(doc, before, 'Undo agent change'));
		const execute = async (doc: Document, command: string, args: Record<string, unknown>, notify = true, agent = false, token?: vscode.CancellationToken) => {
			const request = agent ? activeAgentRequest() : undefined;
			await doc.stopActivity();
			if (agent) { checkAgentAccess(token); }
			await beforeRequest(doc, request);
			if (agent) { checkAgentAccess(token); }
			const beforeState = doc.view.state;
			const before = beforeState.revision;
			const result = await doc.client.request<View>('execute', { command, args, review: kind === 'word' && agent && !!request });
			if (kind !== 'word' && result.state.revision !== before) { doc.latestBytes = await doc.client.request<Uint8Array>('export'); }
			doc.view = { ...doc.view, ...result };
			if (result.state.revision !== before) {
				doc.dirty = true;
				if (notify) { recordEdit(doc, beforeState, command, request); }
			}
			await doc.broadcast(command);
			if (request || /(?:undo|redo)$/.test(command)) { await publishReview(doc, request || doc.reviewRequest); }
			if (agent) { await doc.announceActivity(command, args, beforeState, result.state); }
			return result;
		};
		agentEdits.set(kind, async (doc, input, token, draftSection) => {
			const request = activeAgentRequest();
			if (!vscode.workspace.isTrusted) { throw new Error('Trust this workspace before allowing agent edits.'); }
			const edits = input.operation === 'batch' ? input.edits as Record<string, unknown>[] : [input];
			if (!Array.isArray(edits) || !edits.length || edits.length > 200 || edits.some(edit => !edit || typeof edit.id !== 'string' && edit.operation !== 'append' && !(edit.operation === 'comment' && edit.action !== 'add'))) { throw new Error('Read the document to obtain a target block ID for every edit (1-200 edits).'); }
			const observed = typeof input.version === 'string' && input.version.startsWith(doc.instance + ':') ? Number(input.version.slice(doc.instance.length + 1)) : -1;
			const rebased = input.version !== doc.version;
			if (rebased && (!Number.isSafeInteger(observed) || observed < 0 || observed > doc.view.state.revision || !canRebaseTextEdits(edits))) { throw new Error('Document changed. Read the target again before editing. Structural operations need a fresh version.'); }
			const tail = edits.some(edit => edit.operation === 'append') ? await doc.client.request<{ blocks: { id: string }[] }>('agentRead', { args: { operation: 'tail' } }) : undefined;
			const targetId = typeof edits[0].id === 'string' ? edits[0].id : tail?.blocks[0]?.id;
			if (token.isCancellationRequested) { throw new vscode.CancellationError(); }
			await beforeRequest(doc, request);
			const beforeState = doc.view.state;
			const before = beforeState.revision;
			const idsBefore = targetId ? await doc.client.request<{ blocks: { id: string; parentId: string }[] }>('agentRead', { args: { id: targetId } }) : undefined;
			const parentId = idsBefore?.blocks[0]?.parentId;
			const announced = doc.activity?.request === request && !!doc.activity;
			// Draft fills already have an exact pending view zone. Masking their anchor
			// would also shimmer an unrelated paragraph (often the document title).
			if (draftSection) { await doc.stopActivity(); }
			else if (targetId) {
				const placements: Record<string, 'before' | 'after'> = {};
				for (const edit of edits) {
					if (edit.operation === 'insert' && typeof edit.id === 'string') { placements[edit.id] = edit.position === 'before' ? 'before' : 'after'; }
					if (edit.operation === 'append' && tail?.blocks[0]?.id) { placements[tail.blocks[0].id] = 'after'; }
				}
				await doc.startActivity(targetId, undefined, [targetId, ...edits.flatMap(edit => typeof edit.id === 'string' ? [edit.id] : []), ...Object.keys(placements)], placements);
			}
			const started = Date.now();
			let accepted = false;
			try {
				checkAgentAccess(token);
				const result = await doc.client.request<View>('agentEdit', { args: { ...input, revision: rebased ? observed : before }, review: !!request });
				accepted = true;
				doc.view = { ...doc.view, ...result };
				if (result.state.revision !== before) {
					doc.dirty = true;
					recordEdit(doc, beforeState, 'Agent document edit', request);
				}
				// Let the preview paint before publishing the replacement. The native edit stays atomic.
				await new Promise(resolve => setTimeout(resolve, Math.max(0, 900 - (Date.now() - started))));
				// Remove the editor-only preview before publishing committed content and its reveal.
				await doc.stopActivity().catch(() => {});
				// The native mutation is committed. Optional UI/read failures must not invite a duplicate edit.
				const warnings: string[] = [];
				try { await doc.broadcast(); } catch { warnings.push('Content was applied, but the editor display could not refresh. Reopen the view before continuing.'); }
				try { await publishReview(doc, request); } catch { warnings.push('Content was applied, but change review could not refresh. Do not repeat this edit.'); }
				const created = result.state.agentCreatedBlocks || [];
				let content: Record<string, unknown> = {};
				try { content = await doc.client.request<Record<string, unknown>>('agentRead', { args: input.operation === 'comment' ? { operation: 'comments' } : ['insert', 'append', 'replaceBlocks'].includes(String(input.operation)) && created.length ? { id: created[0].id, limit: 1 } : ['insert', 'delete', 'replaceBlocks'].includes(String(input.operation)) ? { parentId, limit: 20 } : { id: targetId } }); }
				catch { warnings.push('Content was applied; its preview could not be read. Request a small read to inspect it, without repeating the edit.'); }
				const commandResult = JSON.stringify(result.state.agentCommandResult ?? null).length <= 8000 ? result.state.agentCommandResult : { truncated: true, guidance: 'Read the relevant document/comment blocks in pages.' };
				if (result.state.revision !== before) {
					const ids = [...new Set([...created.map(block => block.id), ...edits.filter(edit => edit.operation !== 'delete' && typeof edit.id === 'string').map(edit => String(edit.id))])].slice(0, 30);
					const rects = (await Promise.all(ids.map(blockId => doc.client.request<NonNullable<State['rects']>>('agentRects', { blockId }).catch(() => [])))).flat().slice(0, 200);
					await Promise.all([...doc.panels].map(panel => panel.webview.postMessage({ type: 'agentCommit', revision: doc.view.state.revision, rects, label: activeAgentLabel(), request: activeAgentRequest() }))).catch(() => {});
				}

				const structuralPlaceholders = edits.every(edit => edit.operation === 'delete') ? created : [];
				return { uri: doc.uri.toString(), name: basename(doc.uri.path), version: doc.version, completed: true, ...(input.operation === 'append' ? { placement: 'documentEnd' } : {}), ...(warnings.length ? { warnings } : {}), rebased, ...content, created: structuralPlaceholders.length ? [] : created, ...(structuralPlaceholders.length ? { retainedStructuralParagraphs: structuralPlaceholders.length, guidance: 'The requested deletion completed. The document engine retained required empty structural paragraphs (after tables or in empty containers). These are not unfinished deletions; do not delete them again.' } : {}), commandResult, appliedEdits: result.state.agentAppliedEdits || edits.map(edit => ({ id: edit.id, operation: edit.operation, command: edit.command, params: edit.params })), saved: false, undoable: result.state.revision !== before };
			} finally {
				if (accepted || !announced || token.isCancellationRequested) { await doc.stopActivity().catch(() => { /* Cleanup cannot turn a committed edit into a failure. */ }); }
			}
		});
		const provider: vscode.CustomEditorProvider<Document> = {
			onDidChangeCustomDocument: changed.event,
			async openCustomDocument(uri, openContext) {
				const doc = new Document(uri, kind, context, () => {
					documents.delete(uri.toString());
					if (lastActive.get(kind) === doc) { lastActive.delete(kind); }
				});
				try {
					const original = await read(uri);
					doc.baseHash = hash(original);
					const bytes = openContext.backupId ? await read(vscode.Uri.parse(openContext.backupId)) : original;
					for (; ;) {
						try { await doc.load(bytes); break; } catch (error) {
							if (!/password|encrypted/i.test(String(error))) { throw error; }
							doc.password = await vscode.window.showInputBox({ password: true, prompt: 'Enter the document password' });
							if (doc.password === undefined) { throw error; }
						}
					}
					doc.dirty = !!openContext.backupId;
					if (doc.versions) {
						try {
							const pending = await doc.versions.restoreReview(bytes);
							if (pending) {
								await doc.client.request('reviewImport', { value: pending.review }); doc.reviewRequest = pending.request;
								doc.review = await reviewData(doc);
							}
						} catch { void vscode.window.showErrorMessage('Local change review could not be restored. Document versions are still available from History.'); }
					}
					documents.set(uri.toString(), doc);
					return doc;
				} catch (error) { doc.dispose(); throw error; }
			},
			async resolveCustomEditor(doc, panel) {
				const panelId = randomUUID();
				let appToolbar = false;
				try { appToolbar = await vscode.commands.executeCommand('_sangamOne.documentToolbar.supported') === true; } catch { /* Standalone extension hosts retain their local toolbar. */ }
				if (appToolbar) { toolbarPanels.set(panelId, { doc, panel }); }
				doc.panels.add(panel);
				if (panel.active) { lastActive.set(kind, doc); }
				panel.onDidChangeViewState(() => { if (kind === 'pdf' && panel.visible) { void doc.broadcast(); } if (panel.active) { lastActive.set(kind, doc); if (appToolbar) { void panel.webview.postMessage({ type: 'publishToolbar' }); } } });
				const output = vscode.Uri.joinPath(context.extensionUri, 'out');
				panel.webview.options = { enableScripts: true, localResourceRoots: [output] };
				const nonce = randomBytes(24).toString('hex');
				const resource = (name: string) => panel.webview.asWebviewUri(vscode.Uri.joinPath(output, name));
				panel.webview.html = `<!DOCTYPE html><html><head><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1"><meta http-equiv="Content-Security-Policy" content="default-src 'none'; script-src 'nonce-${nonce}'; style-src ${panel.webview.cspSource} 'unsafe-inline'; img-src data: blob:;"><link rel="stylesheet" href="${resource('editor.css')}"></head><body data-document-instance="${doc.instance}" class="${appToolbar ? 'app-toolbar-hosted' : ''}"><header id="toolbar" aria-label="Document formatting"></header><div id="info"><span id="filename"></span><span id="hint"></span></div><div id="error" role="alert" hidden></div><main id="pages"></main><textarea id="keyboard" aria-label="Document text input" spellcheck="false"></textarea><dialog id="dialog"><form method="dialog"><label id="prompt-label" for="prompt"></label><input id="prompt"><div><button value="cancel">Cancel</button><button id="confirm" value="ok">Apply</button></div></form></dialog><script nonce="${nonce}" src="${resource('editor.js')}"></script></body></html>`;
				const listener = panel.webview.onDidReceiveMessage(message => {
					doc.serial = doc.serial.then(async () => {
						try {
							switch (message.type) {
								case 'toolbarState':
									if (appToolbar && panel.active) { await vscode.commands.executeCommand('_sangamOne.documentToolbar.update', doc.uri, { kind, panel: panelId, instance: doc.instance, context: message.context, controls: message.controls, paragraphStyle: message.paragraphStyle }); }
									break;
								case 'ready': await doc.broadcast(); break;
								case 'tablePreview': {
									// Preview requests never change selection, history, or the document.
									let result;
									if (kind === 'word' && message.args?.revision === doc.view.state.revision) {
										try { result = await doc.client.request('tablePreview', { args: message.args }); } catch { /* A stale drag is discarded. */ }
									}
									await panel.webview.postMessage({ type: 'tablePreview', request: message.request, revision: message.args?.revision, result }); break;
								}
								case 'render': {
									const result = await doc.client.request('render', { page: message.page, scale: message.scale, ...(kind === 'word' ? { frame: message.request, base: message.base } : {}) });
									await panel.webview.postMessage({ type: 'rendered', request: message.request, page: message.page, result }); break;
								}
								case 'pdfPickResource': {
									if (kind !== 'pdf' || typeof message.request !== 'string' || message.request.length > 100) { break; }
									const picked = await vscode.window.showOpenDialog({ canSelectFiles: message.folders !== true, canSelectFolders: message.folders === true, canSelectMany: false });
									await panel.webview.postMessage({ type: 'pdfResource', request: message.request, uri: picked?.[0]?.toString() }); break;
								}
								case 'pdfArtifactList': {
									if (kind !== 'pdf') { break; }
									const chosen = await vscode.window.showQuickPick((doc.view.state.pdfArtifacts || []).map(artifact => ({ label: artifact.name, description: `${Math.ceil(artifact.size / 1024)} KB`, artifact })), { title: 'Generated PDF Files' });
									if (chosen) { await vscode.commands.executeCommand('vscode.open', vscode.Uri.parse(chosen.artifact.uri)); } break;
								}
								case 'pdfOpenArtifact': {
									if (kind === 'pdf' && typeof message.uri === 'string' && doc.view.state.pdfArtifacts?.some(artifact => artifact.uri === message.uri)) { await vscode.commands.executeCommand('vscode.open', vscode.Uri.parse(message.uri)); } break;
								}
								case 'execute': { if (message.revision !== undefined && message.revision !== doc.view.state.revision) { throw new Error('Document changed. Reopen the PDF tools before applying.'); } const result = await execute(doc, message.command, message.args || {}); if (kind === 'pdf' && message.inspect === true) { await panel.webview.postMessage({ type: 'pdfInspection', result: result.state.agentCommandResult }); } break; }
								case 'reviewDecide':
									void vscode.commands.executeCommand('aiOffice.word.reviewDecide', doc.uri.toString(), message.version, message.reviewVersion, message.action, message.ids).then(undefined, async error => {
										await panel.webview.postMessage({ type: 'error', error: String(error) });
										await doc.enqueue(async () => { doc.review = await reviewData(doc); await doc.broadcast(); });
									}); break;
								case 'versions': void vscode.commands.executeCommand('aiOffice.word.versions', doc.uri.toString()); break;
								case 'draftDiscard': doc.drafts.discard(message.id); await doc.broadcast(); break;
								case 'undo': void vscode.commands.executeCommand('_aiOffice.documentHistory', doc.uri, 'undo').then(undefined, error => panel.webview.postMessage({ type: 'error', error: String(error) })); break;
								case 'redo': void vscode.commands.executeCommand('_aiOffice.documentHistory', doc.uri, 'redo').then(undefined, error => panel.webview.postMessage({ type: 'error', error: String(error) })); break;
								case 'save':
									void vscode.commands.executeCommand('workbench.action.files.save').then(() => panel.webview.postMessage(doc.dirty ? { type: 'error', error: 'The document was not saved. Your edits remain open; retry Save or use Save As.' } : { type: 'saved', pendingSections: doc.drafts.status().pending }), error => panel.webview.postMessage({ type: 'error', error: String(error) })); break;
								case 'image': {
									const files = await vscode.window.showOpenDialog({ canSelectMany: false, filters: { Images: ['png', 'jpg', 'jpeg', 'gif', 'webp'] } });
									if (!files?.[0]) { break; }
									const bytes = await read(files[0]);
									if (bytes.length > 16 * 1024 * 1024) { throw new Error('Image exceeds the 16 MB limit.'); }
									const data = Buffer.from(bytes).toString('base64');
									await execute(doc, kind === 'word' ? message.replace ? 'picture.change' : 'insert.picture' : kind === 'pdf' ? 'image.insert' : 'insert.image', { ...(kind === 'word' ? message.replace ? { _image: message.target } : { width: 300 } : { page: message.page }), data, name: basename(files[0].path) });
									break;
								}
							}
						} catch (error) { await panel.webview.postMessage({ type: 'error', error: String(error) }); }
					});
				});
				const watcher = vscode.workspace.createFileSystemWatcher(new vscode.RelativePattern(vscode.Uri.joinPath(doc.uri, '..'), basename(doc.uri.path)));
				const reload = watcher.onDidChange(() => {
					if (doc.saving) { return; }
					doc.serial = doc.serial.then(async () => {
						const bytes = await read(doc.uri);
						if (hash(bytes) === doc.baseHash) { return; }
						if (doc.dirty) { await panel.webview.postMessage({ type: 'error', error: 'This file changed outside Sangam One. Save As to keep your edits before reopening it.' }); return; }
						await doc.load(bytes); doc.baseHash = hash(bytes); await doc.broadcast();
					}).catch(error => { void panel.webview.postMessage({ type: 'error', error: String(error) }); });
				});
				panel.onDidDispose(() => { toolbarPanels.delete(panelId); doc.panels.delete(panel); listener.dispose(); reload.dispose(); watcher.dispose(); });
			},
			async saveCustomDocument(doc) {
				return doc.enqueue(async () => {
					const bytes = await doc.snapshot();
					if (hash(await read(doc.uri)) !== doc.baseHash) { throw new Error('File changed on disk. Use Save As to keep your edits.'); }
					doc.saving = true;
					try { await writeAtomic(doc.uri, bytes); doc.baseHash = hash(bytes); doc.dirty = false; }
					finally { doc.saving = false; }
					try { await doc.saved(); } catch { void vscode.window.showErrorMessage?.('The file was saved, but PDF session metadata could not refresh. Reopen the document before continuing.'); }
					await captureVersion(doc, 'Saved Version');
				});
			},
			async saveCustomDocumentAs(doc, destination) {
				const extension = { word: 'docx', pdf: 'pdf', slides: 'pptx' }[kind];
				if (!destination.path.toLowerCase().endsWith('.' + extension)) { throw new Error(`Save this document with the .${extension} extension.`); }
				await doc.enqueue(async () => writeAtomic(destination, await doc.snapshot()));
			},
			async revertCustomDocument(doc) {
				await doc.stopActivity();
				await doc.serial;
				const bytes = await read(doc.uri);
				await doc.load(bytes); doc.baseHash = hash(bytes); doc.dirty = false; await doc.broadcast();
			},
			async backupCustomDocument(doc, backup) {
				await doc.enqueue(async () => vscode.workspace.fs.writeFile(backup.destination, await doc.snapshot()));
				return { id: backup.destination.toString(), delete: async () => { try { await vscode.workspace.fs.delete(backup.destination); } catch { /* Backup may already have been consumed. */ } } };
			}
		};
		context.subscriptions.push(vscode.window.registerCustomEditorProvider('aiOffice.' + kind, provider, { webviewOptions: { retainContextWhenHidden: true }, supportsMultipleEditorsPerDocument: false }));
		/**
		 * The live command mirrors the user's mental model: it edits the file the
		 * user is looking at. A file name, a path fragment or a URI targets another
		 * open document; omitting the target selects the active one, so agent calls
		 * never need a workspace path.
		 */
		const open = () => [...documents.values()].filter(document => document.kind === kind);
		const list = (docs: Document[]) => docs.map(document => basename(document.uri.path)).join(', ') || 'none';
		const resolveDocument = (target?: string): Document => {
			const candidates = open();
			if (!target) {
				const input = vscode.window.tabGroups.activeTabGroup?.activeTab?.input as { uri?: vscode.Uri; viewType?: string } | undefined;
				const focused = input?.viewType === 'aiOffice.' + kind && input.uri ? documents.get(input.uri.toString()) : undefined;
				const reviewed = focusedReviewDocument();
				const document = focused ?? (reviewed?.kind === kind ? reviewed : undefined) ?? lastActive.get(kind);
				if (document) { return document; }
				if (candidates.length === 1) { return candidates[0]; }
				if (!candidates.length) { throw new Error(`No ${kind} document is open in Sangam One. Open the file first.`); }
				throw new Error(`No active ${kind} document. Open: ${list(candidates)}. Pass the file name to edit.`);
			}
			const uri = /^[a-z][a-z0-9+.-]+:/i.test(target) ? vscode.Uri.parse(target) : undefined;
			if (uri) {
				const document = documents.get(uri.toString()) ?? candidates.find(candidate => candidate.uri.fsPath === uri.fsPath);
				if (document) { return document; }
				throw new Error(`"${target}" is not open in Sangam One. Open: ${list(candidates)}.`);
			}
			const needle = target.replace(/\\/g, '/').toLowerCase();
			const byName = candidates.filter(candidate => basename(candidate.uri.path).toLowerCase() === basename(needle));
			const matches = byName.length ? byName : candidates.filter(candidate => candidate.uri.path.toLowerCase().endsWith('/' + needle));
			if (matches.length === 1) { return matches[0]; }
			if (matches.length > 1) { throw new Error(`"${target}" matches more than one open document: ${list(matches)}.`); }
			throw new Error(`No open ${kind} document matches "${target}". Open: ${list(candidates)}.`);
		};
		context.subscriptions.push(vscode.commands.registerCommand('aiOffice.' + kind + '.capabilities', async (input: { uri?: string; command?: string; category?: string; offset?: number; limit?: number; version?: string; params?: Record<string, unknown>; run?: boolean; operation?: string; resultId?: string; resultPath?: string }, token: vscode.CancellationToken) => {
			checkAgentAccess(token);
			if (!input.run && input.operation === 'list') { return { documents: open().map(doc => ({ uri: doc.uri.toString(), name: basename(doc.uri.path), kind, version: doc.version })) }; }
			if (!input.run && input.operation === 'open') {
				if (!input.uri) { throw new Error('Opening a document requires its full file URI.'); }
				const uri = vscode.Uri.parse(input.uri);
				if (!['file', 'vscode-remote'].includes(uri.scheme) || !uri.path.toLowerCase().endsWith(kind === 'pdf' ? '.pdf' : '.docx')) { throw new Error('Open requires a document URI matching the selected format.'); }
				if (!documents.has(uri.toString())) { await vscode.commands.executeCommand('vscode.openWith', uri, 'aiOffice.' + kind); }
			}
			if (!input.run && input.operation && !['open', 'list', 'discover'].includes(input.operation)) { throw new Error('Use discover, open or list.'); }
			const request = activeAgentRequest(), bound = request ? requestDocuments.get(request) : undefined;
			const doc = resolveDocument(input.uri || bound);
			if (doc.kind !== kind) { throw new Error('The requested document format does not match this command.'); }
			if (request && !bound) { requestDocuments.set(request, doc.uri.toString()); if (requestDocuments.size > 256) { requestDocuments.delete(requestDocuments.keys().next().value!); } }
			if (input.run && kind === 'pdf' && ['edit_undo', 'edit_redo'].includes(input.command || '')) { return history(doc, { operation: input.command === 'edit_redo' ? 'redo' : 'undo', version: input.version }, token); }
			return doc.enqueue(async () => {
				checkAgentAccess(token);
				if (!input.run && input.resultId) { const receipt = doc.commandResults.get(input.resultId); if (!receipt) { throw new Error('This result is no longer cached. Use a current receipt.'); } return { uri: doc.uri.toString(), sourceVersion: receipt.version, resultId: input.resultId, ...projectCommandResult(receipt.value, input.resultPath, input.offset, input.limit) }; }
				if (!input.run) { return { uri: doc.uri.toString(), kind, version: doc.version, ...discoverCommands(doc.view.commands || [], input) }; }
				if (kind !== 'pdf') { throw new Error('Use the guarded document tools for Word. Generic command execution currently supports PDF only.'); }
				const definition = doc.view.commands?.find(command => command.id === input.command);
				if (!definition) { throw new Error('Discover the command schema before running a supported command.'); }
				if (!definition.readOnly && input.version !== doc.version) { throw new Error('Document changed. Discover its current version before editing.'); }
				if (!input.params || typeof input.params !== 'object' || Array.isArray(input.params) || JSON.stringify(input.params).length > 100000) { throw new Error('params must be an object within 100,000 characters.'); }
				const params = { ...input.params };
				if (input.command === 'text_extract' && (!Array.isArray(params.pages) || !params.pages.length || params.pages.length > 10)) { throw new Error('Read text in explicit page ranges of 1–10 pages.'); }
				const beforeRevision = doc.view.state.revision;
				const result = await execute(doc, definition.id, params, true, true, token);
				doc.commandResult = { id: randomUUID(), version: doc.version, value: { result: result.state.agentCommandResult, artifacts: result.state.pdfNewArtifacts || [] } };
				doc.commandResults.set(doc.commandResult.id, { ...doc.commandResult, size: JSON.stringify(doc.commandResult.value).length });
				while (doc.commandResults.size > 1 && (doc.commandResults.size > 4 || [...doc.commandResults.values()].reduce((size, receipt) => size + receipt.size, 0) > 4 * 1024 * 1024)) { doc.commandResults.delete(doc.commandResults.keys().next().value!); }
				return { uri: doc.uri.toString(), kind, version: doc.version, completed: true, modified: result.state.revision !== beforeRevision, saved: false, pages: result.pages.length, result: boundedCommandResult(result.state.agentCommandResult), resultId: doc.commandResult.id, artifacts: (result.state.pdfNewArtifacts || []).slice(0, 20), artifactCount: result.state.pdfNewArtifacts?.length || 0, guidance: result.state.revision !== beforeRevision ? 'Applied to the open document through native undo. Verify the affected pages; save normally when ready.' : undefined };
			});
		}));
		context.subscriptions.push(vscode.commands.registerCommand('aiOffice.' + kind + '.execute', async (uriOrName: string | undefined, command: string, args: Record<string, unknown> = {}, token?: vscode.CancellationToken) => {
			checkAgentAccess(token);
			if (typeof command !== 'string' || !command) { throw new Error(`A command id is required: 'aiOffice.${kind}.execute(fileNameOrUri, command, params)'.`); }
			const doc = resolveDocument(typeof uriOrName === 'string' && uriOrName.trim() ? uriOrName.trim() : undefined);
			if (['edit.undo', 'edit.redo', 'edit_undo', 'edit_redo'].includes(command)) {
				await doc.serial;
				checkAgentAccess(token);
				await vscode.commands.executeCommand('_aiOffice.documentHistory', doc.uri, command.endsWith('redo') ? 'redo' : 'undo');
				return { ...doc.view, name: basename(doc.uri.path), uri: doc.uri.toString() };
			}
			let result: View | undefined;
			const operation = doc.serial.then(async () => { result = await execute(doc, command, args, true, true, token); });
			doc.serial = operation.catch(() => { });
			await operation;
			return result && { ...result, name: basename(doc.uri.path), uri: doc.uri.toString() };
		}));
	}
	// One tool serves every Sangam One kind: the model names the command and,
	// optionally, the file. Without a file it edits the document the user is
	// looking at, so agents never ask the user for a path. The word/pdf
	// extension owns the tool and dispatches to the slides command when needed.
	if (kinds.includes('word')) {
		const kindOfTarget = (file?: string): Kind | undefined => {
			const name = file?.toLowerCase() ?? '';
			if (name.endsWith('.docx')) { return 'word'; }
			if (name.endsWith('.pdf')) { return 'pdf'; }
			if (name.endsWith('.pptx')) { return 'slides'; }
			const viewType = (vscode.window.tabGroups.activeTabGroup?.activeTab?.input as { viewType?: string } | undefined)?.viewType;
			return viewType?.startsWith('aiOffice.') ? viewType.slice('aiOffice.'.length) as Kind : undefined;
		};
		for (const tool of ['tools', 'run'] as const) {
			context.subscriptions.push(vscode.lm.registerTool<DocumentToolInput>('office_document_' + tool, {
				prepareInvocation: () => ({ invocationMessage: tool === 'tools' ? 'Discovering document tools' : 'Working on the PDF document' }),
				async invoke(options, token) {
					try {
						const input = options.input;
						const kind = input.kind !== undefined ? input.kind as Kind : kindOfTarget(input.uri);
						if (!kind || !kinds.includes(kind)) { throw new Error('Open a supported document or specify its URI and kind.'); }
						const result = await withAgentToolRequest(options.toolInvocationToken, async () => vscode.commands.executeCommand('aiOffice.' + kind + '.capabilities', { ...input, run: tool === 'run' }, token));
						return new vscode.LanguageModelToolResult([new vscode.LanguageModelTextPart(JSON.stringify(result))]);
					} catch (error) {
						if (token.isCancellationRequested) { throw error; }
						return new vscode.LanguageModelToolResult([new vscode.LanguageModelTextPart(JSON.stringify({ error: String(error), guidance: 'Use discovery for the current version and exact command schema. Do not repeat an unchanged refused call.' }))]);
					}
				}
			}));
		}
		context.subscriptions.push(vscode.lm.registerTool<{ command: string; file?: string; params?: Record<string, unknown> }>('aiOfficeEdit', {
			invoke: async (options, token) => {
				checkAgentAccess(token);
				const input: { command?: string; file?: string; params?: Record<string, unknown> } = options.input ?? {};
				if (typeof input.command !== 'string' || !input.command) { throw new Error('A command id is required, e.g. content.replace.'); }
				const file = typeof input.file === 'string' && input.file.trim() ? input.file.trim() : undefined;
				const kind = kindOfTarget(file);
				if (!kind) { throw new Error('Open the document to edit in Sangam One first, or pass its file name (DOCX, PDF or PPTX).'); }
					const result = await withAgentToolRequest(options.toolInvocationToken, async () => vscode.commands.executeCommand<View & { name: string; uri: string } | undefined>('aiOffice.' + kind + '.execute', file, input.command, input.params ?? {}, token));
				return new vscode.LanguageModelToolResult([new vscode.LanguageModelTextPart(JSON.stringify(result ?? {}))]);
			}
		}));
	}
	const historyStatus = (doc: Document) => ({
		undo: !!doc.view.state.undo, redo: !!doc.view.state.redo,
		checkpoints: [...doc.history.checkpoints].map(([id, position]) => ({
			id, stepsToRestore: doc.history.cursor - position,
			canRestore: position <= doc.history.cursor && doc.history.nativeStepsTo(position) <= (doc.view.state.undoDepth ?? doc.history.cursor)
		}))
	});
	const history = async (doc: Document, input: DocumentToolInput, token?: vscode.CancellationToken) => {
		await doc.serial;
		if (input.operation === 'status' || !input.operation) { return { uri: doc.uri.toString(), version: doc.version, ...historyStatus(doc) }; }
		if (token?.isCancellationRequested) { throw new vscode.CancellationError(); }
		if (token && input.version !== doc.version) { throw new Error('Document changed. Read history before changing it.'); }
		if (input.operation === 'checkpoint') {
			const checkpoint = randomUUID();
			doc.history.checkpoint(checkpoint);
			return { uri: doc.uri.toString(), version: doc.version, checkpoint, ...historyStatus(doc) };
		}
		let steps = 1;
		const action = input.operation === 'redo' ? 'redo' : 'undo';
		if (input.operation === 'restore') {
			const position = doc.history.checkpoints.get(String(input.checkpoint));
			if (position === undefined || position > doc.history.cursor) { throw new Error('This checkpoint is unavailable. Use Redo to reapply undone edits.'); }
			steps = doc.history.cursor - position;
			if (doc.history.nativeStepsTo(position) > (doc.view.state.undoDepth ?? doc.history.cursor)) { throw new Error('This checkpoint is older than the available undo history.'); }
		} else if (input.operation !== 'undo' && input.operation !== 'redo') { throw new Error('Unknown history operation'); }
		for (let step = 0; step < steps; step++) {
			if (token?.isCancellationRequested) { throw new vscode.CancellationError(); }
			const position = doc.history.cursor;
			if (!await vscode.commands.executeCommand('_aiOffice.documentHistory', doc.uri, action)) { throw new Error('No further document history is available.'); }
			if (doc.history.cursor !== position + (action === 'redo' ? 1 : -1)) { throw new Error('Document history changed during restoration. Read its current status before continuing.'); }
		}
		return { uri: doc.uri.toString(), version: doc.version, ...historyStatus(doc), completed: true };
	};
	if (kinds.includes('word')) {
		if (context.globalStorageUri) {
			const versionText = new Map<string, string>();
			context.subscriptions.push(vscode.workspace.registerTextDocumentContentProvider('ai-office-version', {
				provideTextDocumentContent: uri => versionText.get(uri.toString()) || ''
			}));
			context.subscriptions.push(vscode.workspace.onDidCloseTextDocument(document => versionText.delete(document.uri.toString())));
			context.subscriptions.push(vscode.commands.registerCommand('aiOffice.word.versions', async (uri: string) => {
				try {
					const doc = documents.get(uri); if (!doc?.versions) { throw new Error('Open the Word document to view its versions.'); }
					const versions = await doc.enqueue(async () => { await captureVersion(doc, 'Current Version'); return doc.versions!.list(); });
					const chosen = await vscode.window.showQuickPick(versions.map((version, index) => ({
						label: `${index === 0 ? '$(circle-filled) ' : '$(history) '}${version.label}`,
						description: new Date(version.time).toLocaleString(), detail: 'Compare this document version with the current document', version
					})), { title: 'Document Version History', placeHolder: 'Choose a version to review its differences from the current document' });
					if (!chosen) { return; }
					const original = await doc.versions.projection(chosen.version.id);
					const current = await doc.enqueue(() => doc.client.request<string>('reviewProjection'));
					const key = randomUUID(), name = basename(doc.uri.path);
					const before = vscode.Uri.from({ scheme: 'ai-office-version', path: `/${name} — Previous.md`, query: key });
					const after = vscode.Uri.from({ scheme: 'ai-office-version', path: `/${name} — Current.md`, query: key });
					versionText.set(before.toString(), original); versionText.set(after.toString(), current);
					await vscode.commands.executeCommand('vscode.diff', before, after, `${name} · ${new Date(chosen.version.time).toLocaleString()} ↔ Current`, { preview: true });
				} catch (error) { void vscode.window.showErrorMessage(String(error)); }
			}));
		}
		context.subscriptions.push(vscode.commands.registerCommand('aiOffice.word.reviewStatus', async (uri: string) => {
			const doc = documents.get(uri);
			if (!doc) { throw new Error('Open the document before reviewing its changes.'); }
			return doc.enqueue(() => reviewData(doc));
		}));
		context.subscriptions.push(vscode.commands.registerCommand('aiOffice.word.saveDocument', async (uri: string) => {
			if (!documents.has(uri)) { throw new Error('Document is closed.'); }
			await vscode.commands.executeCommand('workbench.action.files.save', vscode.Uri.parse(uri));
		}));
		context.subscriptions.push(vscode.commands.registerCommand('aiOffice.word.reviewDecide', async (uri: string, version: string, reviewVersion: number, action: 'keep' | 'undo', ids: string[]) => {
			const doc = documents.get(uri);
			if (!doc || doc.kind !== 'word') { throw new Error('Open the Word document before reviewing its changes.'); }
			return doc.enqueue(async () => {
				if (doc.version !== version || doc.reviewVersion !== reviewVersion) { throw new Error('Document changed. Refresh the review before deciding.'); }
				if (!['keep', 'undo'].includes(action) || !Array.isArray(ids) || ids.length > 2000 || ids.some(id => typeof id !== 'string')) { throw new Error('Invalid review changes'); }
				const before = doc.view.state;
				const result = await doc.client.request<View>('reviewDecide', { revision: before.revision, action, ids });
				doc.view = { ...doc.view, ...result };
				if (result.state.revision !== before.revision) {
					doc.dirty = true;
					// User review is a normal undoable edit, not another agent proposal.
					reviewEdits.get('word')!(doc, before);
				}
				await doc.broadcast();
				await publishReview(doc, doc.reviewRequest);
				return reviewData(doc);
			});
		}));
		context.subscriptions.push(vscode.commands.registerCommand('aiOffice.word.requestCheckpoints', (request: string) =>
			[...documents.values()].filter(doc => doc.history.checkpoints.has(request)).map(doc => ({ uri: doc.uri.toString(), name: basename(doc.uri.path), checkpoint: request }))));
		context.subscriptions.push(vscode.commands.registerCommand('aiOffice.word.navigateChatHistory', async (requests: { requestId: string; applied: boolean }[]) => {
			const plans: { doc: Document; position: number }[] = [];
			for (const doc of documents.values()) {
				await doc.serial;
				const relevant = requests.filter(request => doc.history.checkpoints.has(request.requestId) && doc.history.requestEnd(request.requestId) !== undefined);
				if (!relevant.length) { continue; }
				const undone = relevant.filter(request => !request.applied).map(request => doc.history.checkpoints.get(request.requestId)!);
				// Redo stops at this chat's last edit, preserving unrelated later edits
				// that are already applied without replaying ones that remain undone.
				const position = undone.length ? Math.min(...undone) : Math.max(doc.history.cursor, ...relevant.map(request => doc.history.requestEnd(request.requestId)!));
				if (position < doc.history.cursor && doc.history.nativeStepsTo(position) > (doc.view.state.undoDepth ?? doc.history.cursor)) {
					throw new Error('This document checkpoint is older than the available undo history.');
				}
				plans.push({ doc, position });
			}
			for (const { doc, position } of plans) {
				while (doc.history.cursor !== position) {
					await history(doc, { operation: doc.history.cursor < position ? 'redo' : 'undo' });
				}
			}
		}));
	}
	if (kinds.includes('word')) {
		context.subscriptions.push(vscode.commands.registerCommand('_aiOffice.word.renewActivity', (token: unknown, uris: string[]) => withAgentToolRequest(token, async () => {
			const request = activeAgentRequest(); if (!request || !Array.isArray(uris)) { return; }
			await Promise.all(uris.slice(0, 200).map(uri => { const doc = documents.get(uri); return doc?.enqueue(() => doc.renewActivity(request)); }));
		})));
		context.subscriptions.push(vscode.commands.registerCommand('_aiOffice.document.activeDocument', () => {
			const candidates = [...documents.values()];
			const doc = focusedReviewDocument() || candidates.find(doc => [...doc.panels].some(panel => panel.active)) || (candidates.length === 1 ? candidates[0] : undefined);
			return doc ? { uri: doc.uri.toString(), name: basename(doc.uri.path), kind: doc.kind } : undefined;
		}));
		context.subscriptions.push(vscode.commands.registerCommand('_aiOffice.word.activeDocument', () => {
			const candidates = [...documents.values()].filter(doc => doc.kind === 'word');
			const doc = focusedReviewDocument() || candidates.find(doc => [...doc.panels].some(panel => panel.active)) || (candidates.includes(lastActive.get('word')!) ? lastActive.get('word') : undefined) || (candidates.length === 1 ? candidates[0] : undefined);
			return doc ? { uri: doc.uri.toString(), name: basename(doc.uri.path) } : undefined;
		}));
		registerDocumentTools(context, {
			async invoke(tool, input, token) {
				if (!input || typeof input !== 'object' || JSON.stringify(input).length > 160000) { throw new Error('Document tool input is invalid or too large.'); }
				if (tool === 'help' && input.guide !== undefined) {
					if (!['content', 'commands', 'identity', 'drafting', 'comments'].includes(String(input.guide))) { throw new Error('Unknown document guide.'); }
					return { guide: new TextDecoder().decode(await vscode.workspace.fs.readFile(vscode.Uri.joinPath(context.extensionUri, 'skills/ai-office-docs', String(input.guide) + '.md'))) };
				}
				if (tool === 'read' && input.operation === 'open') {
					if (typeof input.uri !== 'string') { throw new Error('Opening a document requires its full file URI.'); }
					const uri = vscode.Uri.parse(input.uri);
					if (!['file', 'vscode-remote'].includes(uri.scheme) || !uri.path.toLowerCase().endsWith('.docx')) { throw new Error('Open accepts a local or workspace DOCX URI.'); }
					if (!documents.has(uri.toString())) { await vscode.commands.executeCommand('vscode.openWith', uri, 'aiOffice.word'); }
				}
				const wordDocuments = [...documents.values()].filter(doc => doc.kind === 'word');
				if (tool === 'read' && input.operation === 'list') {
					return { documents: wordDocuments.map(doc => ({ uri: doc.uri.toString(), name: basename(doc.uri.path), version: doc.version, active: [...doc.panels].some(panel => panel.active) })) };
				}
				const request = activeAgentRequest(), bound = request ? requestDocuments.get(request) : undefined;
				const doc = input.uri ? documents.get(vscode.Uri.parse(input.uri).toString()) : bound ? documents.get(bound) : focusedReviewDocument() || wordDocuments.find(doc => [...doc.panels].some(panel => panel.active)) || lastActive.get('word') || (wordDocuments.length === 1 ? wordDocuments[0] : undefined);
				if (!doc || doc.kind !== 'word') { throw new Error('Open the Word document first. Use office_document_read with operation=list to identify open documents.'); }
				if (request && !bound) { requestDocuments.set(request, doc.uri.toString()); if (requestDocuments.size > 256) { requestDocuments.delete(requestDocuments.keys().next().value!); } }
				if (tool === 'history') { return history(doc, input, token); }
				const task = doc.serial.then(async () => {
					if (token.isCancellationRequested) { throw new vscode.CancellationError(); }
					if (tool === 'comments') {
						const action = input.operation ?? 'list';
						if (!['list', 'add', 'reply', 'update', 'resolve', 'delete'].includes(String(action))) { throw new Error('Use list, add, reply, update, resolve or delete for comments.'); }
						if (action === 'list') { return { uri: doc.uri.toString(), version: doc.version, ...await doc.client.request<Record<string, unknown>>('agentRead', { args: { ...input, operation: 'comments' } }) }; }
						if (action === 'update' && typeof input.resolved === 'boolean') {
							const result = await agentEdits.get('word')!(doc, { operation: 'batch', version: input.version, edits: [
								{ operation: 'comment', action: 'update', commentId: input.commentId, text: input.text },
								{ operation: 'comment', action: 'resolve', commentId: input.commentId, resolved: input.resolved }
							] }, token);
							return { ...result as Record<string, unknown>, ...await doc.client.request<Record<string, unknown>>('agentRead', { args: { operation: 'comments', commentId: input.commentId } }) };
						}
						return agentEdits.get('word')!(doc, { ...input, operation: 'comment', action }, token);
					}
					if (tool === 'draft') {
						const operation = input.operation ?? 'status';
						if (operation === 'status') { return { uri: doc.uri.toString(), version: doc.version, ...doc.drafts.status() }; }
						checkAgentAccess(token);
						if (operation === 'discard') { doc.drafts.discard(typeof input.sectionId === 'string' ? input.sectionId : undefined); await doc.broadcast(); return { uri: doc.uri.toString(), version: doc.version, ...doc.drafts.status() }; }
						if (input.version !== doc.version) { throw new Error('Document changed. Read its current version before planning or filling a section.'); }
						if (operation === 'plan') {
							if (typeof input.id !== 'string' || input.position !== undefined && !['before', 'after'].includes(String(input.position))) { throw new Error('A draft needs a current paragraph/table ID and position before or after.'); }
							const read = await doc.client.request<{ blocks: { id: string; kind: string }[] }>('agentRead', { args: { id: input.id } });
							if (!read.blocks.length || ['row', 'cell'].includes(read.blocks[0].kind)) { throw new Error('Plan sections relative to a current paragraph or table.'); }
							checkAgentAccess(token);
							doc.drafts.plan({ id: input.id, position: input.position === 'before' ? 'before' : 'after' }, input.sections as { title: string; intent: string; level?: number }[]);
							doc.draftAgentLabel = activeAgentLabel(); doc.draftRequest = activeAgentRequest(); doc.draftWorkingUntil = Date.now() + 30000; doc.draftWorkingId = undefined;
							await doc.broadcast(); return { uri: doc.uri.toString(), version: doc.version, ...doc.drafts.status() };
						}
						if (operation !== 'fill' || typeof input.sectionId !== 'string') { throw new Error('Use plan, status, fill or discard for drafting.'); }
						const section = doc.drafts.section(input.sectionId), anchor = doc.drafts.location(input.sectionId);
						doc.draftRequest = activeAgentRequest(); doc.draftWorkingId = section.id; doc.draftWorkingUntil = Date.now() + 30000; await doc.broadcast();
						if (!Array.isArray(input.blocks) || !input.blocks.length || input.blocks.length > 49) { throw new Error('Fill a section with 1-49 content blocks; its heading is added automatically.'); }
						const result = await agentEdits.get('word')!(doc, { operation: 'insert', id: anchor.id, position: anchor.position, version: input.version, blocks: [{ kind: 'heading', text: section.title, style: 'Heading' + section.level }, ...input.blocks] }, token, section.id) as { created: { id: string }[] };
						doc.drafts.complete(section.id, result.created.map(block => block.id)); doc.draftWorkingId = undefined; doc.draftWorkingUntil = Date.now() + 30000; await doc.broadcast();
						return { ...result, filledSectionId: section.id, draft: doc.drafts.status() };
					}
					if (tool === 'format') {
						if (input.ids !== undefined && input.version !== doc.version || input.version !== undefined && input.version !== doc.version) { throw new Error('Document changed. Read the target again before formatting IDs. Query-only formatting can resolve fresh targets without a prior read.'); }
						const plan = await prepareWordFormatting(input, args => doc.client.request('agentRead', { args }));
						const result = await agentEdits.get('word')!(doc, { operation: 'batch', version: doc.version, edits: plan.edits }, token) as Record<string, unknown>;
						return {
							uri: result.uri, version: result.version, saved: result.saved, undoable: result.undoable, appliedEdits: result.appliedEdits,
							verifiedScope: true, matchedBlocks: plan.matchedBlocks, matchedRanges: plan.matchedRanges, scope: plan.scope, target: plan.target, commands: plan.commands,
							guidance: 'The complete resolved formatting batch succeeded atomically. These counts verify the selected scope. Do not repeat it or reread unchanged blocks; only perform any still-missing user requirements.'
						};
					}
					if (tool === 'edit') { return agentEdits.get('word')!(doc, await prepareImageEdits(input), token); }
					if (tool === 'activity') {
						if (input.operation === 'stop') { if (!activeAgentRequest() || !doc.draftRequest || doc.draftRequest === activeAgentRequest()) { doc.draftWorkingUntil = 0; } await doc.stopActivity(); if (doc.drafts.pending().length) { await doc.broadcast(); } }
						else {
							if (input.version !== doc.version || typeof input.id !== 'string') { throw new Error('Read the target again before starting work.'); }
							if (input.intent !== undefined && (typeof input.intent !== 'string' || !input.intent.trim() || input.intent.length > 240)) { throw new Error('Activity intent must be 1–240 characters of plain text.'); }
							if (input.ids !== undefined && (!Array.isArray(input.ids) || !input.ids.length || input.ids.length > 200 || input.ids.some(id => typeof id !== 'string' || !id) || new Set([input.id, ...input.ids]).size > 200)) { throw new Error('Activity ids must contain 1–200 current block IDs.'); }
							if (input.position !== undefined && input.position !== 'before' && input.position !== 'after') { throw new Error('Activity position must be before or after.'); }
							await doc.startActivity(input.id, typeof input.intent === 'string' ? input.intent.trim() : undefined, [input.id, ...(input.ids as string[] || [])], input.position ? Object.fromEntries([input.id, ...(input.ids as string[] || [])].map(id => [id, input.position as 'before' | 'after'])) : {});
						}
						return { uri: doc.uri.toString(), active: !!doc.activity, located: !!doc.activity && (!!doc.activity.rects?.length || doc.activity.insertions.some(section => section.anchorRects?.length)), expiresInSeconds: doc.activity ? 30 : 0 };
					}
					if (tool === 'help') {
						const commands = doc.view.commands || [];
						const category = typeof input.category === 'string' ? input.category : '';
						const matches = commands.filter(command => input.command ? command.id === input.command : category ? command.id.startsWith(category + '.') : true);
						return input.command || category ? { commands: matches, ...(matches.length ? {} : { error: 'No exact command matched. Use category with one of the available prefixes.', categories: [...new Set(commands.map(command => command.id.split('.')[0]))] }), targeted: 'Use office_document_edit operation=command with a block ID. caret.*, select.* and edit.* are UI-only; other enabled mutation commands run on the target. Read-only list.get and styles.inspect use office_document_read operation=inspect with command and paragraph id; never edit to inspect.' } : { categories: [...new Set(commands.map(command => command.id.split('.')[0]))], operations: ['insert', 'update', 'replaceText', 'replaceBlocks', 'delete', 'command', 'batch'], guides: ['content', 'commands', 'identity', 'drafting', 'comments'], tools: ['read', 'edit', 'format', 'draft', 'comments', 'history', 'activity', 'help'], history: historyStatus(doc), guidance: 'Read the ai-office-docs skill. Request one category or command for its exact parameters; IDs are internal and must not appear in user-facing messages.' };
					}
					return { uri: doc.uri.toString(), version: doc.version, history: historyStatus(doc), draft: doc.drafts.status(), ...await doc.client.request<Record<string, unknown>>('agentRead', { args: input }) };
				});
				doc.serial = task.then(() => { }, () => { });
				return task;
			}
		});
	}
}
