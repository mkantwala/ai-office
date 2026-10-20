/*---------------------------------------------------------------------------------------------
 *  Copyright (c) Microsoft Corporation. All rights reserved.
 *  Licensed under the MIT License. See License.txt in the project root for license information.
 *--------------------------------------------------------------------------------------------*/

const assert = require('node:assert/strict');
const { test } = require('node:test');
const Module = require('node:module');
const path = require('node:path');
const fs = require('node:fs');
const documentTools = JSON.parse(fs.readFileSync('extensions/ai-office-pdf/package.json', 'utf8')).contributes.languageModelTools.filter(tool => tool.name.startsWith('office_document_'));

function harness(saved = [], trusted = true, activeDocument = undefined) {
	const disposables = [], commands = new Map(), connections = [], calls = [], items = new Map();
	let provider, controller, participant;
	const statusButton = { show() { this.visible = true; }, dispose() { this.disposed = true; } };
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
		StatusBarAlignment: { Right: 2 },
		workspace: { isTrusted: trusted, workspaceFolders: [{ uri: Uri.file('/workspace/example') }], getConfiguration: () => ({ get: () => 'opencode' }) },
		window: { createStatusBarItem: (id, alignment, priority) => Object.assign(statusButton, { id, alignment, priority }), tabGroups: { activeTabGroup: {} }, showQuickPick: async options => options.find(option => option.id === 'allow') },
		commands: { getCommands: async () => ['workbench.action.chat.openNewSessionSidebar.opencode', 'workbench.action.chat.openNewSessionSidebar.codex', 'workbench.action.chat.openNewSessionSidebar.claude-code'], registerCommand: (id, handler) => { commands.set(id, handler); return { dispose() {} }; }, executeCommand: async (id, ...args) => { calls.push({ command: id, args }); return id === '_aiOffice.word.activeDocument' ? activeDocument : undefined; } },
		lm: {
			tools: documentTools.map(tool => ({ name: tool.name, description: tool.modelDescription, inputSchema: tool.inputSchema })),
			invokeTool: async (name, options, token) => { assert.ok(!token?.isCancellationRequested); calls.push({ name, options }); return { content: [new Text(JSON.stringify(options.input.simulateError ? { error: 'Document changed', retry: 'Read the target again' } : name === 'office_document_activity' ? { uri: 'file:/document.docx' } : { appliedEdits: [{ id: 'paragraph-1' }], version: 'updated' }))] }; }
		},
		chat: {
			createChatSessionItemController: (_type, refreshHandler) => controller = {
				items: { get: uri => items.get(uri.toString()), add: item => items.set(item.resource.toString(), item), replace: values => { items.clear(); values.forEach(item => items.set(item.resource.toString(), item)); } },
				createChatSessionInputState: groups => ({ groups, managed: true }),
				createChatSessionItem: (resource, label) => ({ resource, label }), refreshHandler, dispose() {}
			},
			createChatParticipant: (id, handler) => participant = { id, handler, dispose() {} },
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
				if (text.endsWith('tool-error')) { const result = await invoke('office_document_edit', { simulateError: true }); this.structuredError = result.result; return { stopReason: 'end_turn' }; }
				if (text.endsWith('all-tools')) {
					const listing = await fetch(env.SANGAM_DOCS_URL, { method: 'POST', headers: { Authorization: 'Bearer ' + env.SANGAM_DOCS_TOKEN }, body: JSON.stringify({ id: 2, method: 'tools/list' }) });
					assert.equal(listing.status, 200);
					this.listedTools = (await listing.json()).result.tools;
					for (const tool of documentTools) { await invoke(tool.name, { operation: 'test', uri: 'file:/document.docx', version: 'version', ifHash: 'guard' }); }
					return { stopReason: 'end_turn' };
				}
				await invoke('office_document_activity', { operation: 'start', id: 'paragraph-1' });
				const edited = await invoke('office_document_edit', { id: 'paragraph-1', ifHash: 'guard', text: 'Revised' });
				assert.match(edited.result.content[0].text, /appliedEdits/);
				const permission = await client.requestPermission({ sessionId: 'native-id', toolCall: { title: 'Write document', rawInput: { file: 'document.docx' } }, options: [{ optionId: 'allow', name: 'Allow Once', kind: 'allow_once' }, { optionId: 'deny', name: 'Deny', kind: 'reject_once' }] });
				this.permission = permission;
				client.sessionUpdate({ sessionId: 'native-id', update: { sessionUpdate: 'tool_call', toolCallId: 'doc-edit', title: 'Edit document', kind: 'edit', status: 'in_progress' } });
				client.sessionUpdate({ sessionId: 'native-id', update: { sessionUpdate: 'tool_call_update', toolCallId: 'doc-edit', status: 'completed' } });
				client.sessionUpdate({ sessionId: 'native-id', update: { sessionUpdate: 'agent_message_chunk', content: { type: 'text', text: 'Document updated.' } } });
				return { stopReason: 'end_turn' };
			},
			async cancel() { this.finish?.({ stopReason: 'cancelled' }); },
			dispose() { this.disposed = true; this.alive = false; }
		};
		connections.push(connection); return connection;
	};
	register({ subscriptions: disposables, extensionUri: Uri.file(path.resolve('extensions/ai-office-pdf')), globalState: { get: () => stored, update: async (_key, value) => { stored = JSON.parse(JSON.stringify(value)); } } }, factory);
	const input = () => controller.getChatSessionInputState(undefined, { previousInputState: { groups: [] } });
	const token = () => new CancellationTokenSource();
	const stream = () => { const text = [], parts = []; return { text, parts, progress() {}, push: part => parts.push(part), thinkingProgress() {}, markdown: value => text.push(value) }; };
	const newSession = async () => { const item = await controller.newChatSessionItemHandler({ request: { prompt: 'Edit a document' }, inputState: input() }); controller.items.add(item); return item; };
	return { vscode, statusButton, provider, participant, controller, newSession, input, token, stream, connections, calls, items, commands, saved: () => stored, dispose: () => disposables.reverse().forEach(value => value.dispose()) };
}

test('OpenCode switches back to the original agent explicitly and lists installed native agents', async () => {
	const h = harness();
	try {
		let options;
		h.vscode.window.showQuickPick = async items => { options = items; return items.find(item => item.label === 'Built-in Agent'); };
		await h.commands.get('aiOffice.agents.switch')();
		assert.deepEqual(options.map(item => item.label), ['Built-in Agent', 'OpenCode', 'Codex', 'Claude Code']);
		assert.equal(h.calls.at(-1).command, 'aiOffice.opencode.returnToAgent');
		await h.commands.get('aiOffice.opencode.returnToAgent')();
		assert.deepEqual(h.calls.at(-1), { command: 'workbench.action.chat.newLocalChat', args: [{ agentMode: true }] });
		assert.equal(h.connections.length, 0, 'Switching must not start a CLI or delete an existing conversation');
		h.vscode.window.showQuickPick = async () => undefined;
		const count = h.calls.length; await h.commands.get('aiOffice.agents.switch')(); assert.equal(h.calls.length, count, 'Escape leaves the agent unchanged');
	} finally { h.dispose(); }
});

test('OpenCode manifest enables native session creation commands', () => {
	const manifest = JSON.parse(fs.readFileSync('extensions/ai-office-pdf/package.json', 'utf8'));
	const contribution = manifest.contributes.chatSessions.find(session => session.type === 'opencode');
	assert.equal(contribution.canDelegate, true, 'The workbench only registers native new-session commands for delegating providers');
	assert.equal(contribution.requiresCopilotSignIn, false);
	assert.ok(!manifest.contributes.chatParticipants.some(participant => participant.id === contribution.type), 'The session contribution owns agent registration; a duplicate participant declaration breaks activation');
});

test('OpenCode registers in the native picker and keeps empty sessions lightweight', async () => {
	const h = harness();
	try {
		const session = await h.provider.provideChatSessionContent(h.vscode.Uri.from({ scheme: 'opencode', path: '/untitled-preview' }), h.token().token, { inputState: h.input() });
		assert.equal(session.history.length, 0); assert.equal(h.connections.length, 0);
		await h.commands.get('aiOffice.opencode.newSession')();
		assert.equal(h.calls[0].command, 'workbench.action.chat.openNewSessionSidebar.opencode');
		assert.equal(h.statusButton.alignment, h.vscode.StatusBarAlignment.Right);
		assert.equal(h.statusButton.command, 'aiOffice.agents.switch');
		assert.equal(h.statusButton.visible, true);
		assert.equal(h.participant.id, 'opencode', 'Native sessions invoke their session type as the agent ID');
		const previousInputState = { groups: [{ id: 'model', items: [{ id: 'selected-model' }], selected: { id: 'selected-model' } }] };
		const state = h.controller.getChatSessionInputState(undefined, { previousInputState });
		assert.equal(state.managed, true);
		assert.deepEqual(state.groups, previousInputState.groups);
	} finally { h.dispose(); }
});

test('OpenCode uses request-scoped document tools, clears shimmer and exposes model choices', async () => {
	const h = harness();
	try {
		const item = await h.newSession(), input = h.input();
		await h.provider.provideChatSessionContent(item.resource, h.token().token, { inputState: input });
		const session = { requestHandler: (request, _context, stream, token) => h.participant.handler(request, { history: [], chatSessionContext: { chatSessionItem: item, inputState: input } }, stream, token) };
		const stream = h.stream(); const token = h.token();
		const invocationToken = { requestId: 'verified-request' };
		assert.deepEqual(await session.requestHandler({ prompt: 'Update the document', references: [], toolInvocationToken: invocationToken }, {}, stream, token.token), {});
		const edit = h.calls.find(call => call.name === 'office_document_edit');
		assert.deepEqual(edit.options, { input: { id: 'paragraph-1', ifHash: 'guard', text: 'Revised' }, toolInvocationToken: invocationToken });
		assert.deepEqual(h.connections[0].permission, { outcome: { outcome: 'selected', optionId: 'allow' } });
		assert.equal(h.calls.at(-1).options.input.operation, 'stop');
		assert.equal(h.calls.at(-1).options.toolInvocationToken, invocationToken, 'Cleanup retains ownership and cannot clear another request’s shimmer');
		assert.deepEqual(stream.text, ['Document updated.']);
		assert.deepEqual(stream.parts.map(part => ({ id: part.toolCallId, live: part.enablePartialUpdate, complete: part.isComplete })), [
			{ id: 'doc-edit', live: true, complete: false }, { id: 'doc-edit', live: true, complete: true }
		], 'Initial and final ACP chunks must update one live card, never create a denied legacy card');
		assert.equal(stream.parts[1].invocationMessage, undefined, 'Status-only updates preserve the original tool label');
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
		const pending = h.participant.handler({ prompt: 'wait', references: [], toolInvocationToken: { requestId: 'cancel-me' } }, { history: [], chatSessionContext: { chatSessionItem: item, inputState: h.input() } }, h.stream(), token.token);
		while (!h.connections[0].finish) { await new Promise(resolve => setImmediate(resolve)); }
		token.cancel(); assert.deepEqual(await pending, {});
		assert.equal(item.status, h.vscode.ChatSessionStatus.Completed);
	} finally { h.dispose(); }
});

test('OpenCode refuses to launch in untrusted workspaces', async () => {
	const h = harness([], false);
	try {
		const item = await h.newSession();
		const result = await h.participant.handler({ prompt: 'Edit', references: [] }, { history: [], chatSessionContext: { chatSessionItem: item, inputState: h.input() } }, h.stream(), h.token().token);
		assert.match(result.errorDetails.message, /Trust this workspace/);
		assert.equal(h.connections.length, 0);
	} finally { h.dispose(); }
});


test('Native OpenCode requests expose every document tool with unchanged schemas and request context', async () => {
	const h = harness();
	try {
		const item = await h.newSession(), invocationToken = { requestId: 'all-tools' };
		const result = await h.participant.handler({ prompt: 'all-tools', references: [], toolInvocationToken: invocationToken }, { history: [], chatSessionContext: { chatSessionItem: item, inputState: h.input() } }, h.stream(), h.token().token);
		assert.deepEqual(result, {});
		assert.deepEqual(h.connections[0].listedTools.map(tool => ({ name: tool.name, inputSchema: tool.inputSchema })), documentTools.map(tool => ({ name: tool.name, inputSchema: tool.inputSchema })));
		for (const tool of documentTools) {
			const call = h.calls.find(call => call.name === tool.name && call.options.input.operation === 'test');
			assert.deepEqual(call.options, { input: { operation: 'test', uri: 'file:/document.docx', version: 'version', ifHash: 'guard' }, toolInvocationToken: invocationToken });
		}
	} finally { h.dispose(); }
});

test('Native OpenCode request failures are returned to chat and sessions can retry', async () => {
	const h = harness();
	try {
		const item = await h.newSession(), context = { history: [], chatSessionContext: { chatSessionItem: item, inputState: h.input() } };
		await h.participant.handler({ prompt: 'all-tools', references: [] }, context, h.stream(), h.token().token);
		h.connections[0].prompt = async () => { throw new Error('Provider unavailable'); };
		const result = await h.participant.handler({ prompt: 'retry', references: [] }, context, h.stream(), h.token().token);
		assert.match(result.errorDetails.message, /Provider unavailable/);
		assert.equal(item.status, h.vscode.ChatSessionStatus.Failed);
		h.connections[0].prompt = async () => ({ stopReason: 'end_turn' });
		assert.deepEqual(await h.participant.handler({ prompt: 'retry', references: [] }, context, h.stream(), h.token().token), {});
	} finally { h.dispose(); }
});

	test('OpenCode binds omitted tool URIs to the editor captured at request start', async () => {
	const h = harness([], true, { uri: 'file:/active.docx', name: 'active.docx' });
	try {
		const item = await h.newSession();
		const result = await h.participant.handler({ prompt: 'default-target', references: [], toolInvocationToken: { requestId: 'active-binding' } }, { history: [], chatSessionContext: { chatSessionItem: item, inputState: h.input() } }, h.stream(), h.token().token);
		assert.deepEqual(result, {});
		const edit = h.calls.find(call => call.name === 'office_document_edit');
		assert.equal(edit.options.input.uri, 'file:/active.docx');
		assert.ok(h.connections[0].promptText.includes('active.docx'));
	} finally { h.dispose(); }
	});

test('OpenCode treats structured document refusals as failed MCP tool calls', async () => {
	const h = harness();
	try {
		const item = await h.newSession();
		await h.participant.handler({ prompt: 'tool-error', references: [] }, { history: [], chatSessionContext: { chatSessionItem: item, inputState: h.input() } }, h.stream(), h.token().token);
		assert.equal(h.connections[0].structuredError.isError, true);
		assert.match(h.connections[0].structuredError.content[0].text, /Document changed/);
	} finally { h.dispose(); }
});
