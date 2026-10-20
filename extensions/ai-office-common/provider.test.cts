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
	await assert.rejects(provider.saveCustomDocumentAs(doc, uri('/copy.js')), /\.docx extension/);
	await assert.rejects(provider.saveCustomDocumentAs(doc, uri('/copy.pdf')), /\.docx extension/);
	assert.equal(files.has('/copy.pdf'), false, 'Save As must not write DOCX bytes under another format');
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
					assert.equal(args[0], doc.uri, 'History must cross the extension boundary as a URI object so remote schemes translate');
					if (args[1] === 'undo') { await edits[doc.history.cursor - 1].undo(); } else { await edits[doc.history.cursor].redo(); }
					return true;
				}
				return commands.get(id)(...args);
			}
		},
		workspace: { isTrusted: true, fs: { stat: async () => ({ size: 32 }), readFile: async () => Buffer.from('Initial content') } }
	};
	class EngineClient {
		text = 'Initial content'; revision = 0; past = []; future = []; failNextRead = false; comments = [{ commentId: 0, text: 'Original comment', resolved: true }];
		async request(method, args) {
			if (method === 'agentRead' && args.args.operation === 'comments') { return { comments: this.comments.map(comment => ({ ...comment })) }; }
			if (method === 'agentRead') { if (this.failNextRead) { this.failNextRead = false; throw new Error('Preview read failed'); } return { blocks: [{ id: 'internal-id', parentId: 'body', text: this.text }] }; }
			if (method === 'agentRects') { return [{ page: 0, x: 72, y: 72, width: 120, height: 20 }]; }
			if (method === 'agentEdit' && args.args.operation === 'batch' && args.args.edits.every(edit => edit.operation === 'comment')) {
				assert.equal(args.args.revision, this.revision);
				for (const edit of args.args.edits) { if (edit.action === 'update') { this.comments[0].text = edit.text; } else if (edit.action === 'resolve') { this.comments[0].resolved = edit.resolved; } }
				this.past.push(this.text); this.revision++;
				return { pages: [[612, 792]], state: { revision: this.revision, undo: true, undoDepth: this.past.length, agentAppliedEdits: args.args.edits } };
			}
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
		await withAgentRequest('test-request', () => invoke('activity', { operation: 'start', id: 'internal-id', version: initial.version, intent: 'Clarifying this paragraph' }));
		const first = withAgentRequest('test-request', () => invoke('edit', { operation: 'update', id: 'internal-id', version: initial.version, text: 'First change' }));
		const stale = invoke('edit', { operation: 'update', id: 'internal-id', version: initial.version, text: 'Stale overwrite' });
		assert.equal((await first).saved, false);
		assert.equal((await invoke('history', { operation: 'status' })).checkpoints[0].id, 'test-request');
		assert.match((await stale).error, /Document changed/);
		assert.equal(edits.length, 1);
		assert.ok(messages.filter(message => message.type === 'agentActivity' && message.intent === 'Clarifying this paragraph').length >= 2, 'Automatic edit activity preserves the announced intent on the same block');
		const committedFrame = messages.findIndex(message => message.type === 'state' && message.state.revision === doc.view.state.revision);
		assert.ok(messages.slice(0, committedFrame).some(message => message.type === 'agentActivity' && !message.rects.length), 'The preview clears before committed content is published');
		assert.equal(Buffer.from(doc.latestBytes).toString(), 'Initial content', 'Agent edits also defer DOCX serialization');
		assert.equal((await invoke('read', {})).blocks[0].text, 'First change');
		assert.equal(Buffer.from(await doc.enqueue(() => doc.snapshot())).toString(), 'First change');
		assert.ok(messages.some(message => message.type === 'agentActivity' && message.rects.length));
		assert.ok(messages.some(message => message.type === 'agentCommit' && message.rects.length && message.revision === doc.view.state.revision), 'Committed edits publish bounded geometry for visual reveal');
		assert.deepEqual(messages.at(-1), { type: 'agentActivity', rects: [], label: 'Agent' });
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
		assert.deepEqual(messages.at(-1), { type: 'agentActivity', rects: [], label: 'Agent' });
		await assert.rejects(invoke('edit', { operation: 'update', id: 'internal-id', version: current.version, text: 'Cancelled edit' }, { isCancellationRequested: true }), /Cancelled/);
		vscode.workspace.isTrusted = false;
		assert.match((await invoke('edit', { operation: 'update', id: 'internal-id', version: current.version, text: 'Untrusted edit' })).error, /Trust this workspace/);
		vscode.workspace.isTrusted = true;
		assert.deepEqual((await invoke('help', {})).categories, ['format']);
		assert.equal((await invoke('help', { command: 'format.bold' })).commands.length, 1);
		const previewBytes = Buffer.from(await doc.enqueue(() => doc.snapshot())).toString();
		await invoke('activity', { operation: 'start', id: 'internal-id', version: current.version, intent: 'Clarifying the figures' });
		assert.ok(doc.activity);
		assert.equal(messages.at(-1).intent, 'Clarifying the figures');
		assert.equal(Buffer.from(await doc.enqueue(() => doc.snapshot())).toString(), previewBytes, 'Preview must not change document bytes');
		const previewVersion = doc.version;
		assert.match((await invoke('activity', { operation: 'start', id: 'internal-id', version: current.version, intent: 'x'.repeat(241) })).error, /intent/);
		assert.match((await invoke('activity', { operation: 'start', id: 'internal-id', version: current.version, intent: 7 })).error, /intent/);
		assert.equal(doc.version, previewVersion);
		assert.equal(doc.activity.intent, 'Clarifying the figures', 'Invalid previews leave the existing activity unchanged');
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
		assert.ok(messages.some(message => message.type === 'draft' && message.sections.length === 2 && message.workingId === plan.sections[0].id && message.expiresAt > Date.now()));
		await invoke('activity', { operation: 'stop' });
		assert.equal(messages.at(-1).type, 'draft');
		assert.equal(messages.at(-1).workingId, undefined, 'Stopping a run ends its draft shimmer');
		assert.equal(messages.at(-1).sections.length, 2, 'Stopping keeps the planned sections available');
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
		const commentState = await invoke('comments', {});
		const commentEdit = await invoke('comments', { operation: 'update', version: commentState.version, commentId: 0, text: 'Updated comment', resolved: false });
		assert.equal(commentEdit.completed, true); assert.deepEqual(commentEdit.comments, [{ commentId: 0, text: 'Updated comment', resolved: false }]);
		assert.deepEqual(commentEdit.appliedEdits.map(edit => edit.action), ['update', 'resolve']);
		assert.match((await invoke('comments', { operation: 'update', version: commentEdit.version, commentId: 0, text: 'Bad', resolved: 'false' })).error, /explicit boolean/);
		assert.match((await invoke('comments', { operation: 'reply', version: commentEdit.version, commentId: 0, text: 'Bad', find: 'ignored anchor' })).error, /does not accept find/);

		const parallelDoc = await provider.openCustomDocument(uri('/other.docx'), {});
		const panel = [...doc.panels][0];
		try {
			await withAgentRequest('bound-request', async () => {
				assert.equal((await invoke('read', {})).uri, '/memo.docx');
				panel.active = false; parallelDoc.panels.add({ active: true, webview: { postMessage: async () => {} } });
				assert.equal((await invoke('read', {})).uri, '/memo.docx', 'Switching tabs does not redirect an existing request');
				assert.equal((await invoke('read', { uri: '/other.docx' })).uri, '/other.docx', 'An explicit URI still supports multi-file work');
				assert.equal((await invoke('read', {})).uri, '/memo.docx');
			});
			const a = await invoke('read', { uri: '/memo.docx' }), b = await invoke('read', { uri: '/other.docx' });
			const results = await Promise.all([
				withAgentRequest('parallel-A', () => invoke('edit', { uri: '/memo.docx', operation: 'update', id: 'internal-id', version: a.version, text: 'File A only' })),
				withAgentRequest('parallel-B', () => invoke('edit', { uri: '/other.docx', operation: 'update', id: 'internal-id', version: b.version, text: 'File B only' }))
			]);
			assert.equal(results[0].uri, '/memo.docx'); assert.equal(results[1].uri, '/other.docx');
			assert.equal((await invoke('read', { uri: '/memo.docx' })).blocks[0].text, 'File A only');
			assert.equal((await invoke('read', { uri: '/other.docx' })).blocks[0].text, 'File B only');
			assert.ok(doc.history.checkpoints.has('parallel-A')); assert.equal(doc.history.checkpoints.has('parallel-B'), false);
			assert.ok(parallelDoc.history.checkpoints.has('parallel-B')); assert.equal(parallelDoc.history.checkpoints.has('parallel-A'), false);
			const now = await invoke('read', { uri: '/memo.docx' });
			await withAgentRequest('owner-A', () => invoke('activity', { uri: '/memo.docx', operation: 'start', id: 'internal-id', version: now.version }));
			await withAgentRequest('owner-B', () => invoke('activity', { uri: '/memo.docx', operation: 'start', id: 'internal-id', version: now.version }));
			await withAgentRequest('owner-A', () => invoke('activity', { uri: '/memo.docx', operation: 'stop' }));
			assert.equal(doc.activity.request, 'owner-B', 'An older request cannot clear another request’s activity');
			const timer = doc.activity.timer;
			await commands.get('_aiOffice.word.renewActivity')({ requestId: 'owner-A' }, ['/memo.docx']);
			assert.equal(doc.activity.timer, timer);
			await commands.get('_aiOffice.word.renewActivity')({ requestId: 'owner-B' }, ['/memo.docx']);
			assert.notEqual(doc.activity.timer, timer);
			await withAgentRequest('owner-B', () => invoke('activity', { uri: '/memo.docx', operation: 'stop' }));
			assert.equal(doc.activity, undefined);
			await commands.get('_aiOffice.word.renewActivity')({ requestId: 'owner-B' }, ['/memo.docx']);
			assert.equal(doc.activity, undefined, 'Heartbeat never resurrects a stopped preview');
		} finally { panel.active = true; parallelDoc.dispose(); }

		let opened;
		commands.set('vscode.openWith', async (uri, viewType) => { assert.equal(viewType, 'aiOffice.word'); opened = await provider.openCustomDocument(uri, {}); });
		try {
			const result = await invoke('read', { operation: 'open', uri: '/new.docx' });
			assert.equal(result.uri, '/new.docx'); assert.equal(result.blocks[0].text, 'Initial content');
			assert.match((await invoke('read', { operation: 'open', uri: '/other.pdf' })).error, /DOCX URI/);
		} finally { opened?.dispose(); }

	} finally { doc.dispose(); }
});
