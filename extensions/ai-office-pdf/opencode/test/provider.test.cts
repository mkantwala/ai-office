/*---------------------------------------------------------------------------------------------
 *  Copyright (c) Microsoft Corporation. All rights reserved.
 *  Licensed under the MIT License. See License.txt in the project root for license information.
 *--------------------------------------------------------------------------------------------*/

const assert = require('node:assert/strict');
const { test } = require('node:test');
const Module = require('node:module');
const path = require('node:path');

function harness(saved = [], trusted = true) {
	const disposables = [], commands = new Map(), connections = [], calls = [], items = new Map();
	let provider, controller;
	class Uri {
		constructor(value) { Object.assign(this, value); }
		static from(value) { return new Uri(value); }
		static file(fsPath) { return new Uri({ scheme: 'file', path: fsPath, fsPath }); }
		static joinPath(base, ...parts) { return Uri.file(path.join(base.fsPath, ...parts)); }
		toString() { return this.scheme + ':' + this.path; }
	}
	class CancellationTokenSource {
		constructor() {
			this.listeners = new Set();
			this.token = { isCancellationRequested: false, onCancellationRequested: fn => { this.listeners.add(fn); return { dispose: () => this.listeners.delete(fn) }; } };
		}
		cancel() { this.token.isCancellationRequested = true; this.listeners.forEach(fn => fn()); }
		dispose() { this.listeners.clear(); }
	}
	class Text { constructor(value) { this.value = value; } }
	class RequestTurn { constructor(prompt) { this.prompt = prompt; } }
	class ResponseTurn { constructor(response) { this.response = response; } }
	const vscode = {
		Uri, CancellationTokenSource, LanguageModelTextPart: Text, ChatRequestTurn2: RequestTurn, ChatResponseTurn2: ResponseTurn, ChatResponseMarkdownPart: Text,
		ChatToolInvocationPart: class { constructor(toolName, toolCallId) { Object.assign(this, { toolName, toolCallId }); } },
		ThemeIcon: class { constructor(id) { this.id = id; } }, TabInputCustom: class {},
		ChatSessionStatus: { InProgress: 2, Completed: 1, Failed: 0, NeedsInput: 3 },
		l10n: { t: value => value },
		workspace: { isTrusted: trusted, workspaceFolders: [{ uri: Uri.file('/workspace/example') }], getConfiguration: () => ({ get: () => 'opencode' }) },
		window: { tabGroups: { activeTabGroup: {} }, showQuickPick: async options => options.find(option => option.id === 'allow') },
		commands: { registerCommand: (id, handler) => { commands.set(id, handler); return { dispose() {} }; }, executeCommand: async (id, ...args) => calls.push({ command: id, args }) },
		lm: {
			tools: [{ name: 'office_document_edit', description: 'Edit', inputSchema: { type: 'object' } }, { name: 'office_document_activity', description: 'Activity', inputSchema: { type: 'object' } }],
			invokeTool: async (name, options, token) => { assert.ok(!token?.isCancellationRequested); calls.push({ name, options }); return { content: [new Text(JSON.stringify(name === 'office_document_activity' ? { uri: 'file:/document.docx' } : { appliedEdits: [{ id: 'paragraph-1' }], version: 'updated' }))] }; }
		},
		chat: {
			createChatSessionItemController: (_type, refreshHandler) => controller = {
				items: { get: uri => items.get(uri.toString()), add: item => items.set(item.resource.toString(), item), replace: values => { items.clear(); values.forEach(item => items.set(item.resource.toString(), item)); } },
				createChatSessionItem: (resource, label) => ({ resource, label }), refreshHandler, dispose() {}
			},
			createChatParticipant: () => ({ dispose() {} }),
			registerChatSessionContentProvider: (_type, value) => { provider = value; return { dispose() {} }; }
		}
	};
	const entry = path.resolve('.build/opencode-provider-test.cjs');
	const previous = Module._load;
	Module._load = function (request, parent, isMain) { return request === 'vscode' ? vscode : previous.call(this, request, parent, isMain); };
	let register;
	try { delete require.cache[entry]; register = require(entry).registerOpenCodeAgent; } finally { Module._load = previous; }
	let stored = saved;
	const configOptions = [{ id: 'model', name: 'Model', type: 'select', currentValue: 'first', options: [{ value: 'first', name: 'First' }, { value: 'second', name: 'Second' }] }];
	const factory = (executable, cwd, client) => {
		const connection = { executable, cwd, client, alive: false, disposed: false, configuration: [],
			async open(servers, nativeId) {
				this.alive = true; this.servers = servers; this.loaded = nativeId;
				if (nativeId) {
					client.sessionUpdate({ sessionId: nativeId, update: { sessionUpdate: 'user_message_chunk', content: { type: 'text', text: 'Previous request' } } });
					client.sessionUpdate({ sessionId: nativeId, update: { sessionUpdate: 'agent_message_chunk', content: { type: 'text', text: 'Previous reply' } } });
				}
				return { sessionId: nativeId || 'native-id', configOptions };
			},
			async configure(id, value) { this.configuration.push({ id, value }); return { configOptions: configOptions.map(option => ({ ...option, currentValue: value })) }; },
			async prompt(text) {
				this.promptText = text;
				if (text.endsWith('wait')) { return new Promise(resolve => { this.finish = resolve; }); }
				const env = Object.fromEntries(this.servers[0].env.map(pair => [pair.name, pair.value]));
				const invoke = async (name, input) => {
					const response = await fetch(env.SANGAM_DOCS_URL, { method: 'POST', headers: { Authorization: 'Bearer ' + env.SANGAM_DOCS_TOKEN }, body: JSON.stringify({ id: 1, method: 'tools/call', params: { name, arguments: input } }) });
					return response.json();
				};
				await invoke('office_document_activity', { operation: 'start', id: 'paragraph-1' });
				const edited = await invoke('office_document_edit', { id: 'paragraph-1', ifHash: 'guard', text: 'Revised' });
				assert.match(edited.result.content[0].text, /appliedEdits/);
				const permission = await client.requestPermission({ sessionId: 'native-id', toolCall: { title: 'Write document', rawInput: { file: 'document.docx' } }, options: [{ optionId: 'allow', name: 'Allow Once', kind: 'allow_once' }, { optionId: 'deny', name: 'Deny', kind: 'reject_once' }] });
				this.permission = permission;
				client.sessionUpdate({ sessionId: 'native-id', update: { sessionUpdate: 'agent_message_chunk', content: { type: 'text', text: 'Document updated.' } } });
				return { stopReason: 'end_turn' };
			},
			async cancel() { this.finish?.({ stopReason: 'cancelled' }); },
			dispose() { this.disposed = true; this.alive = false; }
		};
		connections.push(connection); return connection;
	};
	register({ subscriptions: disposables, extensionUri: Uri.file(path.resolve('extensions/ai-office-pdf')), globalState: { get: () => stored, update: async (_key, value) => { stored = JSON.parse(JSON.stringify(value)); } } }, factory);
	const input = () => ({ groups: [] });
	const token = () => new CancellationTokenSource();
	const stream = () => { const text = []; return { text, progress() {}, push() {}, thinkingProgress() {}, markdown: value => text.push(value) }; };
	const newSession = async () => { const item = await controller.newChatSessionItemHandler({ request: { prompt: 'Edit a document' }, inputState: input() }); controller.items.add(item); return item; };
	return { vscode, provider, controller, newSession, input, token, stream, connections, calls, items, commands, saved: () => stored, dispose: () => disposables.reverse().forEach(value => value.dispose()) };
}

test('OpenCode registers in the native picker and keeps empty sessions lightweight', async () => {
	const h = harness();
	try {
		const session = await h.provider.provideChatSessionContent(h.vscode.Uri.from({ scheme: 'opencode', path: '/untitled-preview' }), h.token().token, { inputState: h.input() });
		assert.equal(session.history.length, 0); assert.equal(h.connections.length, 0);
		await h.commands.get('aiOffice.opencode.newSession')();
		assert.equal(h.calls[0].command, 'workbench.action.chat.openNewSessionSidebar.opencode');
	} finally { h.dispose(); }
});

test('OpenCode uses request-scoped document tools, clears shimmer and exposes model choices', async () => {
	const h = harness();
	try {
		const item = await h.newSession(), input = h.input();
		const session = await h.provider.provideChatSessionContent(item.resource, h.token().token, { inputState: input });
		const stream = h.stream(); const token = h.token();
		const invocationToken = { requestId: 'verified-request' };
		assert.deepEqual(await session.requestHandler({ prompt: 'Update the document', references: [], toolInvocationToken: invocationToken }, {}, stream, token.token), {});
		const edit = h.calls.find(call => call.name === 'office_document_edit');
		assert.deepEqual(edit.options, { input: { id: 'paragraph-1', ifHash: 'guard', text: 'Revised' }, toolInvocationToken: invocationToken });
		assert.deepEqual(h.connections[0].permission, { outcome: { outcome: 'selected', optionId: 'allow' } });
		assert.equal(h.calls.at(-1).options.input.operation, 'stop');
		assert.deepEqual(stream.text, ['Document updated.']);
		assert.equal(input.groups[0].items[1].id, 'second');
		assert.equal(h.saved()[0].nativeId, 'native-id');
		const env = Object.fromEntries(h.connections[0].servers[0].env.map(pair => [pair.name, pair.value]));
		const count = h.calls.length;
		const late = await fetch(env.SANGAM_DOCS_URL, { method: 'POST', headers: { Authorization: 'Bearer ' + env.SANGAM_DOCS_TOKEN }, body: JSON.stringify({ id: 2, method: 'tools/call', params: { name: 'office_document_edit', arguments: { text: 'Late edit' } } }) });
		assert.equal((await late.json()).result.isError, true, 'Idle bridge calls cannot mutate the document');
		assert.equal(h.calls.length, count);
		input.groups[0].selected = input.groups[0].items[1];
		await session.requestHandler({ prompt: 'Update again', references: [], toolInvocationToken: invocationToken }, {}, h.stream(), token.token);
		assert.deepEqual(h.connections[0].configuration, [{ id: 'model', value: 'second' }]);
	} finally { h.dispose(); }
});

test('OpenCode reload resumes native history and cancellation stops the active turn', async () => {
	const h = harness([{ id: 'saved-id', nativeId: 'native-id', cwd: '/workspace/example', title: 'Saved', created: 1, modified: 2 }]);
	try {
		const item = [...h.items.values()][0], token = h.token();
		const session = await h.provider.provideChatSessionContent(item.resource, token.token, { inputState: h.input() });
		assert.deepEqual(session.history.map(turn => turn.prompt || turn.response[0].value), ['Previous request', 'Previous reply']);
		const pending = session.requestHandler({ prompt: 'wait', references: [], toolInvocationToken: { requestId: 'cancel-me' } }, {}, h.stream(), token.token);
		while (!h.connections[0].finish) { await new Promise(resolve => setImmediate(resolve)); }
		token.cancel(); assert.deepEqual(await pending, {});
		assert.equal(item.status, h.vscode.ChatSessionStatus.Completed);
	} finally { h.dispose(); }
});

test('OpenCode refuses to launch in untrusted workspaces', async () => {
	const h = harness([], false);
	try {
		const item = await h.newSession();
		const session = await h.provider.provideChatSessionContent(item.resource, h.token().token, { inputState: h.input() });
		const result = await session.requestHandler({ prompt: 'Edit', references: [] }, {}, h.stream(), h.token().token);
		assert.match(result.errorDetails.message, /Trust this workspace/);
		assert.equal(h.connections.length, 0);
	} finally { h.dispose(); }
});
