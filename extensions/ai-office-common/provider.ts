/*---------------------------------------------------------------------------------------------
 *  Copyright (c) Microsoft Corporation. All rights reserved.
 *  Licensed under the MIT License. See License.txt in the project root for license information.
 *--------------------------------------------------------------------------------------------*/

import * as vscode from 'vscode';
import { randomBytes, createHash } from 'node:crypto';
import { basename } from 'node:path';
import { writeFile, rename, unlink } from 'node:fs/promises';
import { EngineClient } from './client.js';
import type { Kind, State, Command } from './types.js';

interface View { state: State; pages: number[][]; commands?: Command[] }
class Document implements vscode.CustomDocument {
	readonly client: EngineClient;
	readonly panels = new Set<vscode.WebviewPanel>();
	view!: View;
	latestBytes!: Uint8Array;
	private snapshotRevision = -1;
	password?: string;
	baseHash = '';
	dirty = false;
	saving = false;
	serial = Promise.resolve();
	constructor(readonly uri: vscode.Uri, readonly kind: Kind, readonly context: vscode.ExtensionContext, private readonly onDispose: () => void) {
		this.client = new EngineClient(vscode.Uri.joinPath(context.extensionUri, 'out', 'engineWorker.cjs').fsPath);
	}
	async load(bytes: Uint8Array) {
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
	}
	dispose() { this.onDispose(); this.client.dispose(); }
}
const hash = (bytes: Uint8Array) => createHash('sha256').update(bytes).digest('hex');
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
	for (const kind of kinds) {
		const changed = new vscode.EventEmitter<vscode.CustomDocumentEditEvent<Document>>();
		context.subscriptions.push(changed);
		const execute = async (doc: Document, command: string, args: Record<string, unknown>, notify = true) => {
			const before = doc.view.state.revision;
			const result = await doc.client.request<View>('execute', { command, args });
			if (kind !== 'word' && result.state.revision !== before) { doc.latestBytes = await doc.client.request<Uint8Array>('export'); }
			doc.view = { ...doc.view, ...result };
			if (notify && result.state.revision !== before) {
				doc.dirty = true;
				changed.fire({
					document: doc, label: command,
					undo: async () => { const task = doc.serial.then(() => execute(doc, kind === 'pdf' ? 'edit_undo' : 'edit.undo', {}, false)); doc.serial = task.then(() => { }, () => { }); await task; },
					redo: async () => { const task = doc.serial.then(() => execute(doc, kind === 'pdf' ? 'edit_redo' : 'edit.redo', {}, false)); doc.serial = task.then(() => { }, () => { }); await task; }
				});
			}
			await doc.broadcast(command);
			return result;
		};
		const provider: vscode.CustomEditorProvider<Document> = {
			onDidChangeCustomDocument: changed.event,
			async openCustomDocument(uri, openContext) {
				const doc = new Document(uri, kind, context, () => documents.delete(uri.toString()));
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
								case 'undo': void vscode.commands.executeCommand('undo'); break;
								case 'redo': void vscode.commands.executeCommand('redo'); break;
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
		context.subscriptions.push(vscode.commands.registerCommand('aiOffice.' + kind + '.execute', async (uri: string, command: string, args: Record<string, unknown> = {}) => {
			const doc = documents.get(vscode.Uri.parse(uri).toString());
			if (!doc) { throw new Error('Open the document in AI Office first.'); }
			let result: View | undefined;
			const operation = doc.serial.then(async () => { result = await execute(doc, command, args); });
			doc.serial = operation.catch(() => { });
			await operation; return result;
		}));
	}
}
