/*---------------------------------------------------------------------------------------------
 *  Copyright (c) Microsoft Corporation. All rights reserved.
 *  Licensed under the MIT License. See License.txt in the project root for license information.
 *--------------------------------------------------------------------------------------------*/

import * as vscode from 'vscode';
import { basename, extname } from 'node:path';
import { randomBytes } from 'node:crypto';
import { hash, read, writeAtomic } from './provider.js';

interface Reply { data: { ok: boolean; error?: string; result?: unknown; revision: number; documentId: number }; bytes: Uint8Array }
class WordDocument implements vscode.CustomDocument {
	panel?: vscode.WebviewPanel;
	password?: string;
	dirty = false;
	saving = false;
	baseHash: string;
	revision = 0;
	documentId?: number;
	needsSaveAs = false;
	nextId = 0;
	serial = Promise.resolve();
	private readyResolve!: () => void;
	readonly ready = new Promise<void>(resolve => this.readyResolve = resolve);
	private openedResolve!: () => void;
	private readonly opened = new Promise<void>(resolve => this.openedResolve = resolve);
	private openError?: Error;
	readonly pending = new Map<number, { resolve(reply: Reply): void; reject(error: Error): void; timer: ReturnType<typeof setTimeout> }>();
	constructor(readonly uri: vscode.Uri, public bytes: Uint8Array, original: Uint8Array, private readonly onDispose: () => void) { this.baseHash = hash(original); }
	loaded() { this.readyResolve(); }
	initialized(error?: Error) { this.openError = error; this.openedResolve(); }
	async request(method: string, args: Record<string, unknown> = {}, bytes?: Uint8Array): Promise<Reply> {
		await this.ready;
		if (method !== 'open') { await this.opened; if (this.openError) { throw this.openError; } }
		if (!this.panel) { throw new Error('Open the Word document first.'); }
		return new Promise((resolve, reject) => {
			const id = ++this.nextId;
			const timer = setTimeout(() => { this.pending.delete(id); reject(new Error('Wordcraft did not respond. Reopen the document; the last snapshot is available for recovery.')); }, 120000);
			this.pending.set(id, { resolve, reject, timer });
			void this.panel!.webview.postMessage({ type: 'request', id, method, args, bytes: bytes ? new Uint8Array(bytes) : undefined });
		});
	}
	reply(message: Reply & { id: number }) {
		const pending = this.pending.get(message.id); if (!pending) { return; }
		this.pending.delete(message.id); clearTimeout(pending.timer);
		if (message.data.ok) { pending.resolve(message); } else { pending.reject(new Error(message.data.error || 'Wordcraft command failed.')); }
	}
	dispose() { this.panel = undefined; this.loaded(); this.initialized(new Error('Word document closed.')); this.onDispose(); for (const pending of this.pending.values()) { clearTimeout(pending.timer); pending.reject(new Error('Word document closed.')); } this.pending.clear(); }
}

export function registerFullWord(context: vscode.ExtensionContext) {
	const documents = new Map<string, WordDocument>();
	const changed = new vscode.EventEmitter<vscode.CustomDocumentContentChangeEvent<WordDocument>>();
	context.subscriptions.push(changed);
	async function open(doc: WordDocument) {
		for (; ;) {
			try { const reply = await doc.request('open', { name: basename(doc.uri.path), password: doc.password, dirty: doc.dirty }, doc.bytes); doc.revision = reply.data.revision; doc.documentId = reply.data.documentId; doc.needsSaveAs = false; doc.initialized(); return; }
			catch (error) {
				if (!/password|encrypted/i.test(String(error))) { throw error; }
				doc.password = await vscode.window.showInputBox({ password: true, prompt: 'Enter the Word document password' });
				if (doc.password === undefined) { throw error; }
			}
		}
	}
	async function save(doc: WordDocument, target: vscode.Uri, conflictCheck: boolean) {
		if (doc.needsSaveAs && target.toString() === doc.uri.toString()) { throw new Error('This is a newly generated document. Use Save As and choose a new file name.'); }
		if (conflictCheck && hash(await read(doc.uri)) !== doc.baseHash) { throw new Error('File changed on disk. Use Save As to keep your edits.'); }
		const snapshot = await doc.request('export', { name: basename(target.path), stamp: true });
		doc.saving = true;
		let success = false;
		try {
			await writeAtomic(target, snapshot.bytes);
			success = true;
			if (target.toString() === doc.uri.toString()) { doc.baseHash = hash(snapshot.bytes); if (doc.revision === snapshot.data.revision) { doc.bytes = snapshot.bytes; doc.dirty = false; } }
		} finally {
			doc.saving = false;
			await doc.request('saved', { revision: snapshot.data.revision, success });
		}
	}
	async function action(doc: WordDocument, command: string, params: Record<string, unknown> = {}) {
		switch (command) {
			case 'file.open': {
				const files = await vscode.window.showOpenDialog({ canSelectMany: false, filters: { Documents: ['docx'] } });
				if (files?.[0]) { await vscode.commands.executeCommand('vscode.openWith', files[0], 'aiOffice.word'); } break;
			}
			case 'file.close': void vscode.commands.executeCommand('workbench.action.closeActiveEditor'); break;
			case 'file.save': void vscode.commands.executeCommand(doc.needsSaveAs ? 'workbench.action.files.saveAs' : 'workbench.action.files.save'); break;
			case 'file.saveAs': void vscode.commands.executeCommand('workbench.action.files.saveAs'); break;
			case 'file.exportPdf': case 'file.exportPng': case 'file.printPdf': {
				const extension = command === 'file.exportPng' ? 'png' : 'pdf';
				const destination = await vscode.window.showSaveDialog({ defaultUri: vscode.Uri.joinPath(doc.uri, '..', basename(doc.uri.path, extname(doc.uri.path)) + '.' + extension), filters: { [extension.toUpperCase()]: [extension] } });
				if (destination) { const snapshot = await doc.request('export', { ...params, name: basename(destination.path) }); await writeAtomic(destination, snapshot.bytes); } break;
			}
		}
	}
	const provider: vscode.CustomEditorProvider<WordDocument> = {
		onDidChangeCustomDocument: changed.event,
		async openCustomDocument(uri, openContext) {
			const original = await read(uri);
			const bytes = openContext.backupId ? await read(vscode.Uri.parse(openContext.backupId)) : original;
			const doc = new WordDocument(uri, bytes, original, () => documents.delete(uri.toString()));
			doc.dirty = !!openContext.backupId; documents.set(uri.toString(), doc); return doc;
		},
		async resolveCustomEditor(doc, panel) {
			doc.panel = panel;
			const output = vscode.Uri.joinPath(context.extensionUri, 'out'), engine = vscode.Uri.joinPath(context.extensionUri, 'media/word-app');
			panel.webview.options = { enableScripts: true, localResourceRoots: [output, engine] };
			const resource = (directory: vscode.Uri, name: string) => panel.webview.asWebviewUri(vscode.Uri.joinPath(directory, name)).toString();
			const nonce = randomBytes(24).toString('hex');
			const config = JSON.stringify({ module: resource(engine, 'wordcraft.js'), wasm: resource(engine, 'wordcraft_bg.wasm') }).replaceAll('<', '\\u003c');
			panel.webview.html = `<!DOCTYPE html><html><head><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1"><meta http-equiv="Content-Security-Policy" content="default-src 'none'; script-src ${panel.webview.cspSource} 'nonce-${nonce}' 'wasm-unsafe-eval'; connect-src ${panel.webview.cspSource}; style-src ${panel.webview.cspSource} 'unsafe-inline'; img-src data: blob:; frame-src blob:;"><link rel="stylesheet" href="${resource(output, 'wordFull.css')}"></head><body><canvas id="wordcraft_canvas" tabindex="0" aria-label="Wordcraft document editor"></canvas><div id="loading">Loading Wordcraft...</div><div id="failure" role="alert" hidden></div><script id="config" type="application/json">${config}</script><script nonce="${nonce}" type="module" src="${resource(output, 'wordFull.js')}"></script></body></html>`;
			const listener = panel.webview.onDidReceiveMessage(message => {
				if (message.type === 'response') { doc.reply(message); return; }
				if (message.type === 'fatal') {
					const error = new Error(message.data?.error || message.error || 'Wordcraft failed to start.');
					doc.loaded(); doc.initialized(error);
					for (const pending of doc.pending.values()) { clearTimeout(pending.timer); pending.reject(error); } doc.pending.clear();
					void panel.webview.postMessage({ type: 'error', error: error.message }); void vscode.window.showErrorMessage(error.message); return;
				}
				if (message.type === 'edit') {
					if (doc.documentId !== undefined && message.data.documentId !== doc.documentId) { doc.needsSaveAs = true; }
					doc.bytes = new Uint8Array(message.bytes); doc.revision = message.data.revision; doc.dirty = true; changed.fire({ document: doc }); return;
				}
				doc.serial = doc.serial.then(async () => {
					try {
						if (message.type === 'ready') { doc.loaded(); await open(doc); }
						else if (message.type === 'action') { await action(doc, message.data.command, message.data.params); }
						else if (message.type === 'paste') {
							const text = await vscode.env.clipboard.readText();
							if (text.length > 16 * 1024 * 1024) { throw new Error('Clipboard text exceeds 16 MB.'); }
							await doc.request('control', { method: 'engine.execute', params: { command: message.data.command, params: { text } } });
						}
						else if (message.type === 'new') {
							const target = await vscode.window.showSaveDialog({ defaultUri: vscode.Uri.joinPath(doc.uri, '..', 'Untitled.docx'), filters: { Word: ['docx'] } });
							if (target) { if (target.toString() === doc.uri.toString()) { throw new Error('Choose a different name for the new document.'); } await writeAtomic(target, new Uint8Array(message.bytes)); await vscode.commands.executeCommand('vscode.openWith', target, 'aiOffice.word'); }
						} else if (message.type === 'pick') {
							if (message.data.purpose === 'document') { await action(doc, 'file.open'); return; }
							const files = await vscode.window.showOpenDialog({ canSelectMany: false });
							if (files?.[0]) { await doc.request('inbox', { name: basename(files[0].path) }, await read(files[0])); }
						} else if (message.type === 'error') { throw new Error(message.data?.error || message.error); }
					} catch (error) { void vscode.window.showErrorMessage(String(error)); if (message.type === 'ready') { doc.loaded(); doc.initialized(new Error(String(error))); await panel.webview.postMessage({ type: 'error', error: String(error) }); } }
				});
			});
			const watcher = vscode.workspace.createFileSystemWatcher(new vscode.RelativePattern(vscode.Uri.joinPath(doc.uri, '..'), basename(doc.uri.path)));
			const reload = watcher.onDidChange(() => {
				if (doc.saving) { return; }
				doc.serial = doc.serial.then(async () => {
					try {
						const bytes = await read(doc.uri); if (hash(bytes) === doc.baseHash) { return; }
						if (doc.dirty) { throw new Error('This Word document changed outside AI Office. Use Save As to keep your edits.'); }
						const previous = doc.bytes;
						doc.bytes = bytes;
						try { await open(doc); doc.baseHash = hash(bytes); } catch (error) { doc.bytes = previous; throw error; }
					} catch (error) { void vscode.window.showErrorMessage(String(error)); }
				});
			});
			panel.onDidDispose(() => { listener.dispose(); reload.dispose(); watcher.dispose(); doc.panel = undefined; });
		},
		async saveCustomDocument(doc) { await doc.serial; await save(doc, doc.uri, true); },
		async saveCustomDocumentAs(doc, target) { await doc.serial; await save(doc, target, false); },
		async revertCustomDocument(doc) {
			await doc.serial; const previous = doc.bytes, dirty = doc.dirty;
			doc.bytes = await read(doc.uri); doc.dirty = false;
			try { await open(doc); doc.baseHash = hash(doc.bytes); } catch (error) { doc.bytes = previous; doc.dirty = dirty; throw error; }
		},
		async backupCustomDocument(doc, backup) {
			await vscode.workspace.fs.writeFile(backup.destination, doc.bytes);
			return { id: backup.destination.toString(), delete: async () => { try { await vscode.workspace.fs.delete(backup.destination); } catch { /* Backup was consumed. */ } } };
		}
	};
	context.subscriptions.push(vscode.window.registerCustomEditorProvider('aiOffice.word', provider, { webviewOptions: { retainContextWhenHidden: true }, supportsMultipleEditorsPerDocument: false }));
	context.subscriptions.push(vscode.commands.registerCommand('aiOffice.word.control', async (uri: string, method: string, params: Record<string, unknown> = {}) => {
		const doc = documents.get(vscode.Uri.parse(uri).toString()); if (!doc) { throw new Error('Open the Word document first.'); }
		const reply = await doc.request('control', { method, params }); return reply.data.result;
	}));
	context.subscriptions.push(vscode.commands.registerCommand('aiOffice.word.execute', async (uri: string, command: string, params: Record<string, unknown> = {}) => {
		const doc = documents.get(vscode.Uri.parse(uri).toString()); if (!doc) { throw new Error('Open the Word document first.'); }
		const reply = await doc.request('control', { method: 'engine.execute', params: { command, params } });
		const result = reply.data.result as { ok: boolean; error?: string; result?: unknown };
		if (!result.ok) { throw new Error(result.error || 'Word command failed.'); } return result.result;
	}));
}
