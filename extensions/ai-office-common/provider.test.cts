/*---------------------------------------------------------------------------------------------
 *  Copyright (c) Microsoft Corporation. All rights reserved.
 *  Licensed under the MIT License. See License.txt in the project root for license information.
 *--------------------------------------------------------------------------------------------*/
const { test } = require('node:test');
const assert = require('node:assert/strict');
const Module = require('node:module');
const path = require('node:path');

test('custom editor saves, backs up, undoes and rejects external conflicts without poisoning its queue', async () => {
	const files = new Map([['/memo.docx', Buffer.from('original document bytes')]]);
	const commands = new Map(); const edits = []; let provider;
	const uri = value => ({ scheme: 'file', path: value, fsPath: value, toString: () => value });
	const nativeFiles = {
		writeFile: async (name, bytes, options) => { assert.equal(options.flag, 'wx'); assert.equal(files.has(name), false); files.set(name, bytes); },
		rename: async (from, to) => { files.set(to, files.get(from)); files.delete(from); },
		unlink: async name => files.delete(name)
	};
	const vscode = {
		Uri: { parse: uri, joinPath: (base, ...parts) => uri(path.posix.resolve(base.path, ...parts)) },
		EventEmitter: class { event = () => {}; fire(value) { edits.push(value); } dispose() {} },
		window: { registerCustomEditorProvider: (_id, value) => { provider = value; return { dispose() {} }; } },
		commands: { registerCommand: (id, fn) => { commands.set(id, fn); return { dispose() {} }; } },
		workspace: { fs: {
			stat: async value => ({ size: files.get(value.path).length }),
			readFile: async value => files.get(value.path),
			writeFile: async (value, bytes) => files.set(value.path, bytes),
			rename: async () => { throw new Error('workspace rename would dispose the open custom-document model'); },
			delete: async value => files.delete(value.path)
		} }
	};
	class EngineClient {
		text = ''; revision = 0; past = []; future = [];
		async request(method, args) {
			if (method === 'open') { this.text = Buffer.from(args.bytes).toString(); this.past = []; this.future = []; }
			if (method === 'execute') {
				if (args.command === 'bad') { throw new Error('Unsupported command'); }
				if (args.command === 'edit.undo') { this.future.push(this.text); this.text = this.past.pop(); }
				else if (args.command === 'edit.redo') { this.past.push(this.text); this.text = this.future.pop(); }
				else { this.past.push(this.text); this.future = []; this.text = args.args.text; }
				this.revision++;
			}
			if (method === 'export') { return Buffer.from(this.text); }
			return { pages: [[612, 792]], state: { revision: this.revision }, commands: [] };
		}
		dispose() {}
	}
	const originalLoad = Module._load;
	try {
		Module._load = (id, ...args) => id === 'vscode' ? vscode : id === './client.js' ? { EngineClient } : id === 'node:fs/promises' ? nativeFiles : originalLoad(id, ...args);
		require('../../.build/office-provider.cjs').registerEditors({ extensionUri: uri('/extension'), subscriptions: [] }, ['word']);
	} finally { Module._load = originalLoad; }
	const doc = await provider.openCustomDocument(uri('/memo.docx'), {});
	const execute = commands.get('aiOffice.word.execute');
	await assert.rejects(execute('/memo.docx', 'bad'), /Unsupported/);
	await execute('/memo.docx', 'text.insert', { text: 'first edit' });
	await execute('/memo.docx', 'text.insert', { text: 'second edit' });
	await edits.at(-1).undo();
	await provider.saveCustomDocument(doc);
	assert.equal(Buffer.from(files.get('/memo.docx')).toString(), 'first edit');
	await edits.at(-1).redo();
	const backup = await provider.backupCustomDocument(doc, { destination: uri('/backup') });
	assert.equal(Buffer.from(files.get('/backup')).toString(), 'second edit');
	files.set('/memo.docx', Buffer.from('external agent changed it'));
	await assert.rejects(provider.saveCustomDocument(doc), /changed on disk/);
	await provider.saveCustomDocumentAs(doc, uri('/copy.docx'));
	assert.equal(Buffer.from(files.get('/copy.docx')).toString(), 'second edit');
	await provider.revertCustomDocument(doc);
	assert.equal(Buffer.from(doc.latestBytes).toString(), 'external agent changed it');
	await backup.delete(); doc.dispose();
	await assert.rejects(execute('/memo.docx', 'text.insert', { text: 'closed' }), /Open the document/);
});
