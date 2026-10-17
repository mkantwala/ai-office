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
	let exports = 0;
	const uri = value => ({ scheme: 'file', path: value, fsPath: value, toString: () => value });
	const nativeFiles = {
		writeFile: async (name, bytes, options) => { assert.equal(options.flag, 'wx'); assert.equal(files.has(name), false); files.set(name, bytes); },
		rename: async (from, to) => { files.set(to, files.get(from)); files.delete(from); },
		unlink: async name => files.delete(name)
	};
	const vscode = {
		lm: { registerTool: () => ({ dispose() {} }) },
		Uri: { parse: uri, joinPath: (base, ...parts) => uri(path.posix.resolve(base.path, ...parts)) },
		EventEmitter: class { event = () => { }; fire(value) { edits.push(value); } dispose() { } },
		window: { registerCustomEditorProvider: (_id, value) => { provider = value; return { dispose() { } }; } },
		commands: { registerCommand: (id, fn) => { commands.set(id, fn); return { dispose() { } }; } },
		workspace: {
			fs: {
				stat: async value => ({ size: files.get(value.path).length }),
				readFile: async value => files.get(value.path),
				writeFile: async (value, bytes) => files.set(value.path, bytes),
				rename: async () => { throw new Error('workspace rename would dispose the open custom-document model'); },
				delete: async value => files.delete(value.path)
			}
		}
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
			if (method === 'export') { exports++; return Buffer.from(this.text); }
			return { pages: [[612, 792]], state: { revision: this.revision }, commands: [] };
		}
		dispose() { }
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
	assert.equal(exports, 0, 'Typing and undo must not ZIP the document before displaying changes');
	await provider.saveCustomDocument(doc);
	assert.equal(exports, 1);
	assert.equal(Buffer.from(files.get('/memo.docx')).toString(), 'first edit');
	await edits.at(-1).redo();
	const backup = await provider.backupCustomDocument(doc, { destination: uri('/backup') });
	assert.equal(Buffer.from(files.get('/backup')).toString(), 'second edit');
	assert.equal(exports, 2, 'Recovery captures the latest engine state');
	files.set('/memo.docx', Buffer.from('external agent changed it'));
	await assert.rejects(provider.saveCustomDocument(doc), /changed on disk/);
	await provider.saveCustomDocumentAs(doc, uri('/copy.docx'));
	assert.equal(exports, 2, 'An unchanged snapshot is reused');
	assert.equal(Buffer.from(files.get('/copy.docx')).toString(), 'second edit');
	await provider.revertCustomDocument(doc);
	assert.equal(Buffer.from(doc.latestBytes).toString(), 'external agent changed it');
	await backup.delete(); doc.dispose();
	await assert.rejects(execute('/memo.docx', 'text.insert', { text: 'closed' }), /Open the document/);
});

test('Word tools serialize edits, reject stale versions, clear activity, and respect cancellation and trust', async () => {
	const tools = new Map(); const edits = []; const messages = [];
	let provider;
	const uri = value => ({ scheme: 'file', path: value, fsPath: value, toString: () => value });
	const vscode = {
		Uri: { parse: uri, joinPath: (base, ...parts) => uri(path.posix.resolve(base.path, ...parts)) },
		EventEmitter: class { event = () => {}; fire(value) { edits.push(value); } dispose() {} },
		CancellationError: class extends Error { constructor() { super('Cancelled'); } },
		LanguageModelTextPart: class { constructor(value) { this.value = value; } },
		LanguageModelToolResult: class { constructor(content) { this.content = content; } },
		lm: { registerTool: (name, value) => { tools.set(name, value); return { dispose() {} }; } },
		window: { registerCustomEditorProvider: (_id, value) => { provider = value; return { dispose() {} }; } },
		commands: { registerCommand: () => ({ dispose() {} }) },
		workspace: { isTrusted: true, fs: { stat: async () => ({ size: 32 }), readFile: async () => Buffer.from('Initial content') } }
	};
	class EngineClient {
		text = 'Initial content'; revision = 0; past = [];
		async request(method, args) {
			if (method === 'agentRead') { return { blocks: [{ id: 'internal-id', parentId: 'body', text: this.text }] }; }
			if (method === 'agentRects') { return [{ page: 0, x: 72, y: 72, width: 120, height: 20 }]; }
			if (method === 'agentEdit') {
				assert.equal(args.args.revision, this.revision);
				if (args.args.text === 'fail') { throw new Error('Invalid edit'); }
				this.past.push(this.text); this.text = args.args.text; this.revision++;
			}
			if (method === 'execute' && args.command === 'edit.undo') { this.text = this.past.pop(); this.revision++; }
			if (method === 'export') { return Buffer.from(this.text); }
			return { pages: [[612, 792]], state: { revision: this.revision }, commands: [{ id: 'format.bold', label: 'Bold', params: { value: 'boolean' } }] };
		}
		dispose() {}
	}
	const originalLoad = Module._load;
	try {
		Module._load = (id, ...args) => id === 'vscode' ? vscode : id === './client.js' ? { EngineClient } : originalLoad(id, ...args);
		delete require.cache[require.resolve('../../.build/office-provider.cjs')];
		require('../../.build/office-provider.cjs').registerEditors({ extensionUri: uri('/extension'), subscriptions: [] }, ['word']);
	} finally { Module._load = originalLoad; }
	const doc = await provider.openCustomDocument(uri('/memo.docx'), {});
	const token = { isCancellationRequested: false };
	const invoke = async (name, input, cancellation = token) => JSON.parse((await tools.get('office_document_' + name).invoke({ input }, cancellation)).content[0].value);
	try {
		doc.panels.add({ active: true, webview: { postMessage: async message => { messages.push(message); } } });
		assert.equal(tools.size, 4);
		const initial = await invoke('read', {});
		assert.equal(initial.blocks[0].text, 'Initial content');
		const first = invoke('edit', { operation: 'update', id: 'internal-id', version: initial.version, text: 'First change' });
		const stale = invoke('edit', { operation: 'update', id: 'internal-id', version: initial.version, text: 'Stale overwrite' });
		assert.equal((await first).saved, false);
		assert.match((await stale).error, /Document changed/);
		assert.equal(edits.length, 1);
		assert.equal(Buffer.from(doc.latestBytes).toString(), 'Initial content', 'Agent edits also defer DOCX serialization');
		assert.equal((await invoke('read', {})).blocks[0].text, 'First change');
		assert.equal(Buffer.from(await doc.enqueue(() => doc.snapshot())).toString(), 'First change');
		assert.ok(messages.some(message => message.type === 'agentActivity' && message.rects.length));
		assert.deepEqual(messages.at(-1), { type: 'agentActivity', rects: [] });
		await edits[0].undo();
		assert.equal((await invoke('read', {})).blocks[0].text, 'Initial content');
		assert.equal(Buffer.from(await doc.enqueue(() => doc.snapshot())).toString(), 'Initial content');
		let current = await invoke('read', {});
		const failure = await invoke('edit', { operation: 'update', id: 'internal-id', version: current.version, text: 'fail' });
		assert.match(failure.error, /Invalid edit/);
		assert.equal((await invoke('read', {})).version, current.version);
		assert.deepEqual(messages.at(-1), { type: 'agentActivity', rects: [] });
		await assert.rejects(invoke('edit', { operation: 'update', id: 'internal-id', version: current.version, text: 'Cancelled edit' }, { isCancellationRequested: true }), /Cancelled/);
		vscode.workspace.isTrusted = false;
		assert.match((await invoke('edit', { operation: 'update', id: 'internal-id', version: current.version, text: 'Untrusted edit' })).error, /Trust this workspace/);
		vscode.workspace.isTrusted = true;
		assert.deepEqual((await invoke('help', {})).categories, ['format']);
		assert.equal((await invoke('help', { command: 'format.bold' })).commands.length, 1);
		await invoke('activity', { operation: 'start', id: 'internal-id', version: current.version });
		assert.ok(doc.activity);
		await invoke('activity', { operation: 'stop' });
		assert.equal(doc.activity, undefined);
		await provider.revertCustomDocument(doc);
		current = await invoke('read', {});
		assert.notEqual(current.version, initial.version, 'Reopen/revert must invalidate old request versions even if engine revision resets');
	} finally { doc.dispose(); }
});
