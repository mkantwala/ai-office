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
import type { Kind, State, Command, DocumentReview, DocumentReviewChanges } from './types.js';
import { registerDocumentTools, type DocumentToolInput } from './documentTools.js';
import { activeAgentRequest, DocumentHistory, withAgentToolRequest } from './agentRequest.js';
import { prepareWordFormatting } from './wordFormatting.js';

interface View { state: State; pages: number[][]; commands?: Command[] }
class Document implements vscode.CustomDocument {
	readonly client: EngineClient;
	readonly panels = new Set<vscode.WebviewPanel>();
	readonly history = new DocumentHistory();
	reviewRequest?: string;
	reviewVersion = 0;
	view!: View;
	latestBytes!: Uint8Array;
	agentActivity?: { region?: ActivityRegion; updated: number };
	private snapshotRevision = -1;
	password?: string;
	baseHash = '';
	dirty = false;
	saving = false;
	serial = Promise.resolve();
	instance = randomUUID();
	activity?: { id: string; rects: State['rects']; timer: ReturnType<typeof setTimeout> };
	constructor(readonly uri: vscode.Uri, readonly kind: Kind, readonly context: vscode.ExtensionContext, private readonly onDispose: () => void) {
		this.client = new EngineClient(vscode.Uri.joinPath(context.extensionUri, 'out', 'engineWorker.cjs').fsPath);
	}
	async load(bytes: Uint8Array) {
		this.instance = randomUUID();
		this.history.clear();
		this.reviewRequest = undefined;
		this.reviewVersion++;
		this.view = await this.client.request<View>('open', { kind: this.kind, name: basename(this.uri.path), bytes, extensionsRoot: vscode.Uri.joinPath(this.context.extensionUri, '..').fsPath, password: this.password });
		this.latestBytes = bytes;
		this.snapshotRevision = this.view.state.revision;
	}
	enqueue<T>(operation: () => Promise<T>): Promise<T> {
		const task = this.serial.then(operation);
		this.serial = task.then(() => { }, () => { });
		return task;
	}
	async snapshot(): Promise<Uint8Array> {
		if (this.kind === 'word' && this.snapshotRevision !== this.view.state.revision) {
			this.latestBytes = await this.client.request<Uint8Array>('export');
			this.snapshotRevision = this.view.state.revision;
		}
		return this.latestBytes;
	}
	async broadcast(command?: string) {
		await Promise.all([...this.panels].map(panel => panel.webview.postMessage({ type: 'state', kind: this.kind, ...this.view, name: basename(this.uri.path), command })));
		if (this.activity) { await this.sendActivity(this.activity.rects); }
	}
	async sendActivity(rects: State['rects']) {
		await Promise.all([...this.panels].map(panel => panel.webview.postMessage({ type: 'agentActivity', rects: rects || [] })));
	}
	async stopActivity() { if (this.activity) { clearTimeout(this.activity.timer); this.activity = undefined; } await this.sendActivity([]); }
	async startActivity(id: string) {
		const rects = await this.client.request<State['rects']>('agentRects', { blockId: id });
		await this.stopActivity();
		this.activity = { id, rects, timer: setTimeout(() => { void this.stopActivity(); }, 30000) };
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
		await Promise.all([...this.panels].map(panel => panel.webview.postMessage({ type: 'agent', label: command, region })));
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
	const publishReview = async (doc: Document, request?: string) => {
		if (doc.kind !== 'word' || !request) { return; }
		doc.reviewRequest = request;
		doc.reviewVersion++;
		const data = await reviewData(doc);
		try { await vscode.commands.executeCommand('aiOffice.registerDocumentReview', request, data); }
		catch { void vscode.window.showErrorMessage?.('Document changes were applied, but change review could not open. Reload AI Office before continuing.'); }
	};
	const agentEdits = new Map<Kind, (doc: Document, input: DocumentToolInput, token: vscode.CancellationToken) => Promise<unknown>>();
	const reviewEdits = new Map<Kind, (doc: Document, before: State) => void>();
	// The file an agent edits when none is named: the document whose editor is
	// active, or the one most recently active among the open documents.
	const lastActive = new Map<Kind, Document>();
	const focusedReviewDocument = () => {
		const input = vscode.window.tabGroups?.activeTabGroup?.activeTab?.input as { uri?: vscode.Uri; modified?: vscode.Uri } | undefined;
		const resource = input?.uri || input?.modified;
		return resource?.scheme === 'ai-office-review' ? documents.get(resource.query) : undefined;
	};
	for (const kind of kinds) {
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
		agentEdits.set(kind, async (doc, input, token) => {
			const request = activeAgentRequest();
			if (!vscode.workspace.isTrusted) { throw new Error('Trust this workspace before allowing agent edits.'); }
			if (input.version !== doc.version) { throw new Error('Document changed. Read the target again before editing.'); }
			const edits = input.operation === 'batch' ? input.edits as Record<string, unknown>[] : [input];
			if (!Array.isArray(edits) || !edits.length || edits.length > 200 || edits.some(edit => !edit || typeof edit.id !== 'string')) { throw new Error('Read the document to obtain a target block ID for every edit (1-200 edits).'); }
			const targetId = edits[0].id as string;
			if (token.isCancellationRequested) { throw new vscode.CancellationError(); }
			const beforeState = doc.view.state;
			const before = beforeState.revision;
			const idsBefore = await doc.client.request<{ blocks: { id: string; parentId: string }[] }>('agentRead', { args: { id: targetId } });
			const parentId = idsBefore.blocks[0]?.parentId;
			await doc.startActivity(targetId);
			const started = Date.now();
			try {
				if (token.isCancellationRequested) { throw new vscode.CancellationError(); }
				const result = await doc.client.request<View>('agentEdit', { args: { ...input, revision: before }, review: !!request });
				doc.view = { ...doc.view, ...result };
				if (result.state.revision !== before) {
					doc.dirty = true;
					recordEdit(doc, beforeState, 'Agent document edit', request);
				}
				if (input.operation !== 'delete' && doc.activity) {
					// Some text commands merge/remove the target; the edit still succeeded.
					doc.activity.rects = await doc.client.request<State['rects']>('agentRects', { blockId: targetId }).catch(() => []);
				}
				await doc.broadcast();
				await publishReview(doc, request);
				const content = await doc.client.request<Record<string, unknown>>('agentRead', { args: input.operation === 'insert' || input.operation === 'delete' ? { parentId, limit: 20 } : { id: targetId } });
				const commandResult = JSON.stringify(result.state.agentCommandResult ?? null).length <= 8000 ? result.state.agentCommandResult : { truncated: true, guidance: 'Read the relevant document/comment blocks in pages.' };
				return { uri: doc.uri.toString(), version: doc.version, ...content, created: result.state.agentCreatedBlocks || [], commandResult, appliedEdits: result.state.agentAppliedEdits || edits.map(edit => ({ id: edit.id, operation: edit.operation, command: edit.command, params: edit.params })), saved: false, undoable: result.state.revision !== before };
			} finally {
				// Give fast operations enough time to paint, without delaying the engine itself.
				await new Promise(resolve => setTimeout(resolve, Math.max(0, 400 - (Date.now() - started))));
				await doc.stopActivity();
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
					documents.set(uri.toString(), doc);
					return doc;
				} catch (error) { doc.dispose(); throw error; }
			},
			async resolveCustomEditor(doc, panel) {
				doc.panels.add(panel);
				if (panel.active) { lastActive.set(kind, doc); }
				panel.onDidChangeViewState(() => { if (panel.active) { lastActive.set(kind, doc); } });
				const output = vscode.Uri.joinPath(context.extensionUri, 'out');
				panel.webview.options = { enableScripts: true, localResourceRoots: [output] };
				const nonce = randomBytes(24).toString('hex');
				const resource = (name: string) => panel.webview.asWebviewUri(vscode.Uri.joinPath(output, name));
				panel.webview.html = `<!DOCTYPE html><html><head><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1"><meta http-equiv="Content-Security-Policy" content="default-src 'none'; script-src 'nonce-${nonce}'; style-src ${panel.webview.cspSource} 'unsafe-inline'; img-src data: blob:;"><link rel="stylesheet" href="${resource('editor.css')}"></head><body><header id="toolbar" aria-label="Document formatting"></header><div id="info"><span id="filename"></span><span id="hint"></span></div><div id="error" role="alert" hidden></div><main id="pages"></main><textarea id="keyboard" aria-label="Document text input" spellcheck="false"></textarea><dialog id="dialog"><form method="dialog"><label id="prompt-label" for="prompt"></label><input id="prompt"><div><button value="cancel">Cancel</button><button id="confirm" value="ok">Apply</button></div></form></dialog><script nonce="${nonce}" src="${resource('editor.js')}"></script></body></html>`;
				const listener = panel.webview.onDidReceiveMessage(message => {
					doc.serial = doc.serial.then(async () => {
						try {
							switch (message.type) {
								case 'ready': await doc.broadcast(); break;
								case 'render': {
									const result = await doc.client.request('render', { page: message.page, scale: message.scale, ...(kind === 'word' ? { frame: message.request, base: message.base } : {}) });
									await panel.webview.postMessage({ type: 'rendered', request: message.request, page: message.page, result }); break;
								}
								case 'execute': await execute(doc, message.command, message.args || {}); break;
								case 'undo': void vscode.commands.executeCommand('_aiOffice.documentHistory', doc.uri.toString(), 'undo').then(undefined, error => panel.webview.postMessage({ type: 'error', error: String(error) })); break;
								case 'redo': void vscode.commands.executeCommand('_aiOffice.documentHistory', doc.uri.toString(), 'redo').then(undefined, error => panel.webview.postMessage({ type: 'error', error: String(error) })); break;
								case 'save':
									void vscode.commands.executeCommand('workbench.action.files.save').then(() => panel.webview.postMessage({ type: 'saved' }), error => panel.webview.postMessage({ type: 'error', error: String(error) })); break;
								case 'image': {
									const files = await vscode.window.showOpenDialog({ canSelectMany: false, filters: { Images: ['png', 'jpg', 'jpeg', 'gif', 'webp'] } });
									if (!files?.[0]) { break; }
									const bytes = await read(files[0]);
									if (bytes.length > 16 * 1024 * 1024) { throw new Error('Image exceeds the 16 MB limit.'); }
									const data = Buffer.from(bytes).toString('base64');
									await execute(doc, kind === 'word' ? 'insert.picture' : kind === 'pdf' ? 'image.insert' : 'insert.image', { ...(kind === 'word' ? { width: 300 } : { page: message.page }), data, name: basename(files[0].path) });
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
						if (doc.dirty) { await panel.webview.postMessage({ type: 'error', error: 'This file changed outside AI Office. Save As to keep your edits before reopening it.' }); return; }
						await doc.load(bytes); doc.baseHash = hash(bytes); await doc.broadcast();
					}).catch(error => { void panel.webview.postMessage({ type: 'error', error: String(error) }); });
				});
				panel.onDidDispose(() => { doc.panels.delete(panel); listener.dispose(); reload.dispose(); watcher.dispose(); });
			},
			async saveCustomDocument(doc) {
				return doc.enqueue(async () => {
					const bytes = await doc.snapshot();
					if (hash(await read(doc.uri)) !== doc.baseHash) { throw new Error('File changed on disk. Use Save As to keep your edits.'); }
					doc.saving = true;
					try { await writeAtomic(doc.uri, bytes); doc.baseHash = hash(bytes); doc.dirty = false; }
					finally { doc.saving = false; }
				});
			},
			async saveCustomDocumentAs(doc, destination) {
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
				if (!candidates.length) { throw new Error(`No ${kind} document is open in AI Office. Open the file first.`); }
				throw new Error(`No active ${kind} document. Open: ${list(candidates)}. Pass the file name to edit.`);
			}
			const uri = /^[a-z][a-z0-9+.-]+:/i.test(target) ? vscode.Uri.parse(target) : undefined;
			if (uri) {
				const document = documents.get(uri.toString()) ?? candidates.find(candidate => candidate.uri.fsPath === uri.fsPath);
				if (document) { return document; }
				throw new Error(`"${target}" is not open in AI Office. Open: ${list(candidates)}.`);
			}
			const needle = target.replace(/\\/g, '/').toLowerCase();
			const byName = candidates.filter(candidate => basename(candidate.uri.path).toLowerCase() === basename(needle));
			const matches = byName.length ? byName : candidates.filter(candidate => candidate.uri.path.toLowerCase().endsWith('/' + needle));
			if (matches.length === 1) { return matches[0]; }
			if (matches.length > 1) { throw new Error(`"${target}" matches more than one open document: ${list(matches)}.`); }
			throw new Error(`No open ${kind} document matches "${target}". Open: ${list(candidates)}.`);
		};
		context.subscriptions.push(vscode.commands.registerCommand('aiOffice.' + kind + '.execute', async (uriOrName: string | undefined, command: string, args: Record<string, unknown> = {}, token?: vscode.CancellationToken) => {
			checkAgentAccess(token);
			if (typeof command !== 'string' || !command) { throw new Error(`A command id is required: 'aiOffice.${kind}.execute(fileNameOrUri, command, params)'.`); }
			const doc = resolveDocument(typeof uriOrName === 'string' && uriOrName.trim() ? uriOrName.trim() : undefined);
			if (['edit.undo', 'edit.redo', 'edit_undo', 'edit_redo'].includes(command)) {
				await doc.serial;
				checkAgentAccess(token);
				await vscode.commands.executeCommand('_aiOffice.documentHistory', doc.uri.toString(), command.endsWith('redo') ? 'redo' : 'undo');
				return { ...doc.view, name: basename(doc.uri.path), uri: doc.uri.toString() };
			}
			let result: View | undefined;
			const operation = doc.serial.then(async () => { result = await execute(doc, command, args, true, true, token); });
			doc.serial = operation.catch(() => { });
			await operation;
			return result && { ...result, name: basename(doc.uri.path), uri: doc.uri.toString() };
		}));
	}
	// One tool serves every AI Office kind: the model names the command and,
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
		context.subscriptions.push(vscode.lm.registerTool<{ command: string; file?: string; params?: Record<string, unknown> }>('aiOfficeEdit', {
			invoke: async (options, token) => {
				checkAgentAccess(token);
				const input: { command?: string; file?: string; params?: Record<string, unknown> } = options.input ?? {};
				if (typeof input.command !== 'string' || !input.command) { throw new Error('A command id is required, e.g. content.replace.'); }
				const file = typeof input.file === 'string' && input.file.trim() ? input.file.trim() : undefined;
				const kind = kindOfTarget(file);
				if (!kind) { throw new Error('Open the document to edit in AI Office first, or pass its file name (DOCX, PDF or PPTX).'); }
				const result = await withAgentToolRequest(options.toolInvocationToken, () => vscode.commands.executeCommand<View & { name: string; uri: string } | undefined>('aiOffice.' + kind + '.execute', file, input.command, input.params ?? {}, token));
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
			if (!await vscode.commands.executeCommand('_aiOffice.documentHistory', doc.uri.toString(), action)) { throw new Error('No further document history is available.'); }
			if (doc.history.cursor !== position + (action === 'redo' ? 1 : -1)) { throw new Error('Document history changed during restoration. Read its current status before continuing.'); }
		}
		return { uri: doc.uri.toString(), version: doc.version, ...historyStatus(doc), completed: true };
	};
	if (kinds.includes('word')) {
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
				if (!Array.isArray(ids) || ids.length > 2000 || ids.some(id => typeof id !== 'string')) { throw new Error('Invalid review changes'); }
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
		registerDocumentTools(context, {
			async invoke(tool, input, token) {
				if (!input || typeof input !== 'object' || JSON.stringify(input).length > 160000) { throw new Error('Document tool input is invalid or too large.'); }
				if (tool === 'help' && input.guide !== undefined) {
					if (!['content', 'commands', 'identity'].includes(String(input.guide))) { throw new Error('Unknown document guide.'); }
					return { guide: new TextDecoder().decode(await vscode.workspace.fs.readFile(vscode.Uri.joinPath(context.extensionUri, 'skills/ai-office-docs', String(input.guide) + '.md'))) };
				}
				const wordDocuments = [...documents.values()].filter(doc => doc.kind === 'word');
				if (tool === 'read' && input.operation === 'list') {
					return { documents: wordDocuments.map(doc => ({ uri: doc.uri.toString(), name: basename(doc.uri.path), version: doc.version, active: [...doc.panels].some(panel => panel.active) })) };
				}
					const doc = input.uri ? documents.get(vscode.Uri.parse(input.uri).toString()) : focusedReviewDocument() || wordDocuments.find(doc => [...doc.panels].some(panel => panel.active)) || lastActive.get('word') || (wordDocuments.length === 1 ? wordDocuments[0] : undefined);
				if (!doc || doc.kind !== 'word') { throw new Error('Open the Word document first. Use office_document_read with operation=list to identify open documents.'); }
				if (tool === 'history') { return history(doc, input, token); }
				const task = doc.serial.then(async () => {
					if (token.isCancellationRequested) { throw new vscode.CancellationError(); }
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
					if (tool === 'edit') { return agentEdits.get('word')!(doc, input, token); }
					if (tool === 'activity') {
						if (input.operation === 'stop') { await doc.stopActivity(); }
						else {
							if (input.version !== doc.version || typeof input.id !== 'string') { throw new Error('Read the target again before starting work.'); }
							await doc.startActivity(input.id);
						}
						return { uri: doc.uri.toString(), active: !!doc.activity, expiresInSeconds: doc.activity ? 30 : 0 };
					}
					if (tool === 'help') {
						const commands = doc.view.commands || [];
						const category = typeof input.category === 'string' ? input.category : '';
						const matches = commands.filter(command => input.command ? command.id === input.command : category ? command.id.startsWith(category + '.') : true);
						return input.command || category ? { commands: matches, ...(matches.length ? {} : { error: 'No exact command matched. Use category with one of the available prefixes.', categories: [...new Set(commands.map(command => command.id.split('.')[0]))] }), targeted: 'Use office_document_edit operation=command with a block ID. caret.*, select.* and edit.* are UI-only; other enabled commands run on the target.' } : { categories: [...new Set(commands.map(command => command.id.split('.')[0]))], operations: ['insert', 'update', 'delete', 'command', 'batch'], history: historyStatus(doc), guidance: 'Read the ai-office-docs skill. Request one category or command for its exact parameters; IDs are internal and must not appear in user-facing messages.' };
					}
					return { uri: doc.uri.toString(), version: doc.version, history: historyStatus(doc), ...await doc.client.request<Record<string, unknown>>('agentRead', { args: input }) };
				});
				doc.serial = task.then(() => { }, () => { });
				return task;
			}
		});
	}
}
