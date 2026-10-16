/*---------------------------------------------------------------------------------------------
 *  Copyright (c) Microsoft Corporation. All rights reserved.
 *  Licensed under the MIT License. See License.txt in the project root for license information.
 *--------------------------------------------------------------------------------------------*/
const { test } = require('node:test');
const assert = require('node:assert/strict');
const Module = require('node:module');
const path = require('node:path');

test('full Word synchronizes binary documents, releases failed saves, recovers and protects conflicts', async () => {
	const files = new Map([['/memo.docx', Buffer.from('original')]]);
	const commands = new Map(), messages = [], changes = [];
	let provider, receive, failWrite = false;
	const uri = value => ({ path: value, toString: () => value });
	const disposable = { dispose() { } };
	const vscode = {
		Uri: { parse: uri, joinPath: (base, ...parts) => uri(path.posix.resolve(base.path, ...parts)) },
		RelativePattern: class { },
		EventEmitter: class { event = () => { }; fire(value) { changes.push(value); } dispose() { } },
		window: { registerCustomEditorProvider: (_id, value) => { provider = value; return disposable; }, showErrorMessage: error => { throw new Error(error); } },
		commands: { registerCommand: (id, fn) => { commands.set(id, fn); return disposable; } },
		workspace: {
			createFileSystemWatcher: () => ({ ...disposable, onDidChange: () => disposable }),
			fs: { writeFile: async (target, bytes) => files.set(target.path, bytes), delete: async target => files.delete(target.path) }
		}
	};
	const helpers = {
		hash: bytes => Buffer.from(bytes).toString('hex'), read: async target => files.get(target.path),
		writeAtomic: async (target, bytes) => { if (failWrite) { throw new Error('Disk full'); } files.set(target.path, bytes); }
	};
	const originalLoad = Module._load;
	try {
		Module._load = (id, ...args) => id === 'vscode' ? vscode : id === './provider.js' ? helpers : originalLoad(id, ...args);
		require('../../.build/office-word-provider.cjs').registerFullWord({ extensionUri: uri('/extension'), subscriptions: [] });
	} finally { Module._load = originalLoad; }
	const doc = await provider.openCustomDocument(uri('/memo.docx'), {});
	const panel = {
		webview: {
			cspSource: 'test:', asWebviewUri: value => value, onDidReceiveMessage: fn => { receive = fn; return disposable; },
			postMessage: async message => {
				messages.push(message);
				if (message.type !== 'request') { return true; }
				if (message.method === 'open') { assert.equal(message.bytes.constructor, Uint8Array); assert.equal(Buffer.from(message.bytes).toString(), 'original'); }
				receive({ type: 'response', id: message.id, data: { ok: true, revision: 2, documentId: 1, result: {} }, bytes: new Uint8Array(Buffer.from('edited')) });
				return true;
			}
		}, onDidDispose: () => disposable
	};
	await provider.resolveCustomEditor(doc, panel);
	receive({ type: 'ready' }); await doc.serial;
	receive({ type: 'edit', data: { revision: 2, documentId: 1 }, bytes: Buffer.from('edited') });
	assert.equal(changes.length, 1);
	const backup = await provider.backupCustomDocument(doc, { destination: uri('/backup') });
	assert.equal(Buffer.from(files.get('/backup')).toString(), 'edited');
	failWrite = true;
	await assert.rejects(provider.saveCustomDocument(doc), /Disk full/);
	assert.equal(messages.at(-1).method, 'saved'); assert.equal(messages.at(-1).args.success, false); assert.equal(doc.dirty, true);
	failWrite = false; await provider.saveCustomDocument(doc);
	assert.equal(messages.at(-1).args.success, true); assert.equal(doc.dirty, false);
	files.set('/memo.docx', Buffer.from('external'));
	await assert.rejects(provider.saveCustomDocument(doc), /changed on disk/);
	await provider.saveCustomDocumentAs(doc, uri('/copy.docx'));
	assert.equal(Buffer.from(files.get('/copy.docx')).toString(), 'edited');
	receive({ type: 'edit', data: { revision: 3, documentId: 2 }, bytes: Buffer.from('generated') });
	await assert.rejects(provider.saveCustomDocument(doc), /newly generated/);
	await backup.delete(); doc.dispose();
	await assert.rejects(commands.get('aiOffice.word.execute')('/memo.docx', 'text.insert'), /Open the Word document/);
});
