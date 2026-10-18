/*---------------------------------------------------------------------------------------------
 *  Copyright (c) Microsoft Corporation. All rights reserved.
 *  Licensed under the MIT License. See License.txt in the project root for license information.
 *--------------------------------------------------------------------------------------------*/
const { test } = require('node:test');
const assert = require('node:assert/strict');
const Module = require('node:module');
const path = require('node:path');
const { withAgentRequest } = require('./agentRequest.ts');

test('custom editor saves, backs up, undoes and rejects external conflicts without poisoning its queue', async () => {
	const files = new Map([['/memo.docx', Buffer.from('original document bytes')]]);
	const commands = new Map(); const tools = new Map(); const edits = []; let provider;
	let exports = 0;
	const uri = value => ({ scheme: 'file', path: value, fsPath: value, toString: () => value });
	const nativeFiles = {
		writeFile: async (name, bytes, options) => { assert.equal(options.flag, 'wx'); assert.equal(files.has(name), false); files.set(name, bytes); },
		rename: async (from, to) => { files.set(to, files.get(from)); files.delete(from); },
		unlink: async name => files.delete(name)
	};
	const vscode = {
		Uri: { parse: uri, joinPath: (base, ...parts) => uri(path.posix.resolve(base.path, ...parts)) },
		CancellationError: class extends Error { constructor() { super('Cancelled'); } },
		EventEmitter: class { event = () => { }; fire(value) { edits.push(value); } dispose() { } },
		LanguageModelTextPart: class { constructor(value) { this.value = value; } },
		LanguageModelToolResult: class { constructor(content) { this.content = content; } },
		window: {
			registerCustomEditorProvider: (_id, value) => { provider = value; return { dispose() { } }; },
			tabGroups: { activeTabGroup: undefined }
		},
		lm: { registerTool: (id, tool) => { tools.set(id, tool); return { dispose() { } }; } },
		commands: {
			registerCommand: (id, fn) => { commands.set(id, fn); return { dispose() { } }; },
			executeCommand: (id, ...args) => commands.get(id)(...args)
		},
		workspace: {
			isTrusted: true,
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
			return { pages: [[612, 792]], state: { revision: this.revision, ...(args.command === 'select.text' ? { rects: [{ page: 0, x: 1, y: 2, width: 30, height: 8 }] } : {}) }, commands: [] };
		}
		dispose() { }
	}
	const originalLoad = Module._load;
	try {
		Module._load = (id, ...args) => id === 'vscode' ? vscode : id === './client.js' ? { EngineClient } : id === './agentRequest.js' ? require('./agentRequest.ts') : id === 'node:fs/promises' ? nativeFiles : originalLoad(id, ...args);
		require('../../.build/office-provider.cjs').registerEditors({ extensionUri: uri('/extension'), subscriptions: [] }, ['word']);
	} finally { Module._load = originalLoad; }
	const doc = await provider.openCustomDocument(uri('/memo.docx'), {});
	const execute = commands.get('aiOffice.word.execute');
	await assert.rejects(execute('/memo.docx', 'bad'), /Unsupported/);
	const messages = [];
	doc.panels.add({ webview: { postMessage: async message => { messages.push(message); } } });
	await execute('/memo.docx', 'select.text', { text: 'first' });
	assert.ok(messages.some(message => message.type === 'agent' && message.label === 'select.text' && message.region?.rects?.length === 1), 'Agent commands report the touched region to the editor');
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

	// Agents never carry document ids or workspace paths: with no target the
	// command edits the document whose editor is active, a bare file name
	// selects an open document and the result names the file that was edited.
	vscode.window.tabGroups = { activeTabGroup: { activeTab: { input: { viewType: 'aiOffice.word', uri: uri('/memo.docx') } } } };
	const active = await execute(undefined, 'text.insert', { text: 'via the active document' });
	assert.equal(active.name, 'memo.docx');
	assert.equal(active.uri, '/memo.docx');
	const named = await execute('memo.docx', 'text.insert', { text: 'via a file name' });
	assert.equal(named.name, 'memo.docx');
	await assert.rejects(execute('missing.docx', 'text.insert', { text: 'x' }), /No open word document matches "missing\.docx"/);

	// The chat tool uses the same resolution and reports the file it edited.
	const tool = tools.get('aiOfficeEdit');
	const toolResult = await tool.invoke({ input: { command: 'text.insert', params: { text: 'via the chat tool' } } }, {});
	assert.ok(toolResult.content[0].value.includes('"name":"memo.docx"'), 'The tool result reports the resolved file name');
	const beforeRejectedCalls = { text: doc.client.text, revision: doc.client.revision, edits: edits.length };
	const input = { command: 'text.insert', params: { text: 'Rejected agent edit' } };
	await assert.rejects(tool.invoke({ input }, { isCancellationRequested: true }), /Cancelled/);
	vscode.workspace.isTrusted = false;
	await assert.rejects(tool.invoke({ input }, { isCancellationRequested: false }), /Trust this workspace/);
	await assert.rejects(execute('memo.docx', input.command, input.params), /Trust this workspace/);
	vscode.workspace.isTrusted = true;

	// A call can be cancelled or lose trust while waiting behind document work.
	for (const reason of ['cancel', 'trust']) {
		const cancellation = { isCancellationRequested: false };
		let release;
		doc.serial = new Promise(resolve => { release = resolve; });
		const pending = tool.invoke({ input }, cancellation);
		if (reason === 'cancel') { cancellation.isCancellationRequested = true; }
		else { vscode.workspace.isTrusted = false; }
		release();
		await assert.rejects(pending, reason === 'cancel' ? /Cancelled/ : /Trust this workspace/);
		vscode.workspace.isTrusted = true;
	}
	assert.deepEqual({ text: doc.client.text, revision: doc.client.revision, edits: edits.length }, beforeRejectedCalls, 'Rejected calls never reach the engine or add undo events');

	await backup.delete(); doc.dispose();
	await assert.rejects(execute('/memo.docx', 'text.insert', { text: 'closed' }), /No open word document matches/);
});

test('Word tools serialize edits, reject stale versions, clear activity, and respect cancellation and trust', async () => {
	const tools = new Map(); const edits = []; const messages = []; const commands = new Map();
	let provider;
	const uri = value => ({ scheme: 'file', path: value, fsPath: value, toString: () => value });
	const vscode = {
		Uri: { parse: uri, joinPath: (base, ...parts) => uri(path.posix.resolve(base.path, ...parts)) },
		EventEmitter: class { event = () => { }; fire(value) { edits.push(value); } dispose() { } },
		CancellationError: class extends Error { constructor() { super('Cancelled'); } },
		LanguageModelTextPart: class { constructor(value) { this.value = value; } },
		LanguageModelToolResult: class { constructor(content) { this.content = content; } },
		lm: { registerTool: (name, value) => { tools.set(name, value); return { dispose() { } }; } },
		window: { registerCustomEditorProvider: (_id, value) => { provider = value; return { dispose() { } }; } },
		commands: {
			registerCommand: (id, fn) => { commands.set(id, fn); return { dispose() { } }; }, executeCommand: async (id, ...args) => {
				if (id === '_aiOffice.documentHistory') {
					if (args[1] === 'undo') { await edits[doc.history.cursor - 1].undo(); } else { await edits[doc.history.cursor].redo(); }
					return true;
				}
				return commands.get(id)(...args);
			}
		},
		workspace: { isTrusted: true, fs: { stat: async () => ({ size: 32 }), readFile: async () => Buffer.from('Initial content') } }
	};
	class EngineClient {
		text = 'Initial content'; revision = 0; past = []; future = []; failNextRead = false;
		async request(method, args) {
			if (method === 'agentRead') { if (this.failNextRead) { this.failNextRead = false; throw new Error('Preview read failed'); } return { blocks: [{ id: 'internal-id', parentId: 'body', text: this.text }] }; }
			if (method === 'agentRects') { return [{ page: 0, x: 72, y: 72, width: 120, height: 20 }]; }
			if (method === 'agentEdit') {
				assert.equal(args.args.revision, this.revision);
				if (args.args.text === 'fail') { throw new Error('Invalid edit'); }
				this.past.push(this.text); this.future = []; this.text = args.args.text; this.revision++;
				this.failNextRead = this.text === 'Applied despite preview failure';
			}
			if (method === 'execute' && args.command === 'edit.undo') { this.future.push(this.text); this.text = this.past.pop(); this.revision++; }
			if (method === 'execute' && args.command === 'edit.redo') { this.past.push(this.text); this.text = this.future.pop(); this.revision++; }
			if (method === 'export') { return Buffer.from(this.text); }
			return { pages: [[612, 792]], state: { revision: this.revision, undo: this.past.length > 0, redo: this.future.length > 0, undoDepth: this.past.length }, commands: [{ id: 'format.bold', label: 'Bold', params: { value: 'boolean' } }] };
		}
		dispose() { }
	}
	const originalLoad = Module._load;
	try {
		Module._load = (id, ...args) => id === 'vscode' ? vscode : id === './client.js' ? { EngineClient } : id === './agentRequest.js' ? require('./agentRequest.ts') : originalLoad(id, ...args);
		delete require.cache[require.resolve('../../.build/office-provider.cjs')];
		require('../../.build/office-provider.cjs').registerEditors({ extensionUri: uri('/extension'), subscriptions: [] }, ['word']);
	} finally { Module._load = originalLoad; }
	const doc = await provider.openCustomDocument(uri('/memo.docx'), {});
	const token = { isCancellationRequested: false };
	const invoke = async (name, input, cancellation = token) => JSON.parse((await tools.get('office_document_' + name).invoke({ input }, cancellation)).content[0].value);
	try {
		doc.panels.add({ active: true, webview: { postMessage: async message => { messages.push(message); } } });
		assert.equal(tools.size, 9);
		const initial = await invoke('read', {});
		assert.equal(initial.blocks[0].text, 'Initial content');
		const other = await provider.openCustomDocument(uri('/other.docx'), {});
		vscode.window.tabGroups = { activeTabGroup: { activeTab: { input: { uri: { scheme: 'ai-office-review', query: other.uri.toString() } } } } };
		assert.equal((await invoke('read', {})).uri, '/other.docx', 'The active review targets its source Word document even with several DOCX files open');
		vscode.window.tabGroups = undefined;
		other.dispose();
		const first = withAgentRequest('test-request', () => invoke('edit', { operation: 'update', id: 'internal-id', version: initial.version, text: 'First change' }));
		const stale = invoke('edit', { operation: 'update', id: 'internal-id', version: initial.version, text: 'Stale overwrite' });
		assert.equal((await first).saved, false);
		assert.equal((await invoke('history', { operation: 'status' })).checkpoints[0].id, 'test-request');
		assert.match((await stale).error, /Document changed/);
		assert.equal(edits.length, 1);
		assert.equal(Buffer.from(doc.latestBytes).toString(), 'Initial content', 'Agent edits also defer DOCX serialization');
		assert.equal((await invoke('read', {})).blocks[0].text, 'First change');
		assert.equal(Buffer.from(await doc.enqueue(() => doc.snapshot())).toString(), 'First change');
		assert.ok(messages.some(message => message.type === 'agentActivity' && message.rects.length));
		assert.deepEqual(messages.at(-1), { type: 'agentActivity', rects: [] });
		await edits[0].undo();
		assert.equal((await invoke('history', { operation: 'status' })).redo, true);
		await edits[0].redo();
		assert.equal((await invoke('read', {})).blocks[0].text, 'First change');
		const status = await invoke('history', { operation: 'status' });
		await invoke('history', { operation: 'restore', checkpoint: 'test-request', version: status.version });
		assert.equal((await invoke('read', {})).blocks[0].text, 'Initial content');
		assert.equal(Buffer.from(await doc.enqueue(() => doc.snapshot())).toString(), 'Initial content');
		const navigate = commands.get('aiOffice.word.navigateChatHistory');
		await navigate([{ requestId: 'test-request', applied: true }]);
		assert.equal((await invoke('read', {})).blocks[0].text, 'First change');
		await withAgentRequest('second-request', async () => {
			let state = await invoke('read', {});
			await invoke('edit', { operation: 'update', id: 'internal-id', version: state.version, text: 'Second change' });
			state = await invoke('read', {});
			await invoke('edit', { operation: 'update', id: 'internal-id', version: state.version, text: 'Third change' });
		});
		await navigate([{ requestId: 'test-request', applied: true }]);
		assert.equal((await invoke('read', {})).blocks[0].text, 'Third change', 'Navigating an already-applied chat preserves later edits');
		await navigate([{ requestId: 'test-request', applied: true }, { requestId: 'second-request', applied: false }]);
		assert.equal((await invoke('read', {})).blocks[0].text, 'First change', 'Checkpoint restores every operation in the selected request');
		await navigate([{ requestId: 'test-request', applied: true }, { requestId: 'second-request', applied: true }]);
		assert.equal((await invoke('read', {})).blocks[0].text, 'Third change', 'Default timeline Redo reapplies all operations');
		await navigate([{ requestId: 'test-request', applied: false }, { requestId: 'second-request', applied: false }]);
		assert.equal((await invoke('read', {})).blocks[0].text, 'Initial content', 'Start Over restores all requests');
		await navigate([{ requestId: 'test-request', applied: true }]);
		assert.equal((await invoke('read', {})).blocks[0].text, 'First change', 'Redo of one chat must not replay another chat\'s later edits');
		await navigate([{ requestId: 'test-request', applied: true }, { requestId: 'second-request', applied: true }]);
		let latest = await invoke('read', {});
		await invoke('edit', { operation: 'update', id: 'internal-id', version: latest.version, text: 'Later independent edit' });
		latest = await invoke('history', { operation: 'status' });
		await invoke('history', { operation: 'undo', version: latest.version });
		await navigate([{ requestId: 'test-request', applied: true }, { requestId: 'second-request', applied: true }]);
		assert.equal((await invoke('read', {})).blocks[0].text, 'Third change', 'Chat Redo leaves later independent edits undone');
		await navigate([{ requestId: 'test-request', applied: false }, { requestId: 'second-request', applied: false }]);
		await navigate([{ requestId: 'unrelated-session', applied: false }]);
		assert.equal((await invoke('read', {})).blocks[0].text, 'Initial content');
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
		assert.match((await invoke('edit', { operation: 'replaceText', id: 'internal-id', ifHash: 'a'.repeat(64), version: initial.version, find: 'Initial', text: 'Stale' })).error, /Document changed/);
		const exportBefore = Buffer.from(await doc.enqueue(() => doc.snapshot())).toString();
		const plan = await invoke('draft', { operation: 'plan', version: current.version, id: 'internal-id', sections: [{ title: 'Pending A', intent: 'Research A' }, { title: 'Pending B', intent: 'Research B' }] });
		assert.equal(plan.pending, 2); assert.equal(plan.savedInDocument, false);
		assert.equal((await invoke('draft', { operation: 'status' })).pending, 2);
		assert.equal(Buffer.from(await doc.enqueue(() => doc.snapshot())).toString(), exportBefore);
		assert.ok(messages.some(message => message.type === 'draft' && message.sections.length === 2));
		const cancelled = { isCancellationRequested: true };
		await assert.rejects(invoke('draft', { operation: 'fill', version: current.version, sectionId: plan.sections[0].id, blocks: [{ kind: 'paragraph', text: 'Never apply' }] }, cancelled), /Cancelled/);
		assert.equal((await invoke('draft', { operation: 'status' })).pending, 2);
		assert.match((await invoke('draft', { operation: 'discard', sectionId: 42 })).error, /sectionId/);
		assert.equal((await invoke('draft', { operation: 'status' })).pending, 2);
		vscode.workspace.isTrusted = false;
		assert.match((await invoke('draft', { operation: 'discard' })).error, /Trust this workspace/);
		vscode.workspace.isTrusted = true;
		assert.equal((await invoke('draft', { operation: 'discard', sectionId: plan.sections[0].id })).pending, 1);
		assert.equal((await invoke('draft', { operation: 'discard' })).pending, 0);
		assert.match((await invoke('comments', { operation: 'resolve', version: current.version, commentId: 0 })).error, /explicit boolean/);
		assert.match((await invoke('comments', { operation: 'reply', version: current.version, commentId: -1, text: 'No' })).error, /commentId/);
		const applied = await invoke('edit', { operation: 'update', id: 'internal-id', version: current.version, text: 'Applied despite preview failure' });
		assert.equal(applied.completed, true); assert.equal(applied.error, undefined); assert.match(applied.warnings[0], /Content was applied/);
		assert.equal((await invoke('read', {})).blocks[0].text, 'Applied despite preview failure');
		await invoke('history', { operation: 'undo', version: applied.version });
		assert.equal((await invoke('read', {})).blocks[0].text, 'Initial content');


	} finally { doc.dispose(); }
});
