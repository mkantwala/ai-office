/*---------------------------------------------------------------------------------------------
 *  Copyright (c) Microsoft Corporation. All rights reserved.
 *  Licensed under the MIT License. See License.txt in the project root for license information.
 *--------------------------------------------------------------------------------------------*/
// Opt-in paid network integration test. Never included in the ordinary test runner.
// Run: node --use-env-proxy extensions/ai-office-common/cloudflareAgent.live.cts
// Requires CF_ACCOUNT_ID and CF_API_TOKEN. Uses disposable synthetic DOCX files only.
const assert = require('node:assert/strict');
const fs = require('node:fs/promises');
const path = require('node:path');
const Module = require('node:module');
const { pathToFileURL, fileURLToPath } = require('node:url');
const { randomUUID } = require('node:crypto');
const { build } = require('../../build/node_modules/esbuild');
const { zipSync, strToU8, unzipSync, strFromU8 } = require('../ai-office-slides/node_modules/fflate');
const repository = path.resolve(__dirname, '../..');
const selected = process.argv.find(value => value.startsWith('--only='))?.slice(7).split(',');
const reportRoot = path.join(repository, '.build/live-cloudflare', new Date().toISOString().replace(/[:.]/g, '-'));
const reports = [];
let transportCalls = 0, current, providerBlocked;
const usage = { prompt_tokens: 0, completion_tokens: 0 };
const safe = value => String(value).replaceAll(process.env.CF_API_TOKEN || 'unused-secret', '[redacted]');
const uri = value => {
	const url = value.startsWith('file:') ? new URL(value) : pathToFileURL(value);
	return { scheme: 'file', path: decodeURIComponent(url.pathname), fsPath: fileURLToPath(url), toString: () => url.href };
};
const xmlEscape = value => value.replaceAll('&', '&amp;').replaceAll('<', '&lt;').replaceAll('>', '&gt;');
function fixture(paragraphs) {
	return zipSync(Object.fromEntries(Object.entries({
		'[Content_Types].xml': '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>',
		'_rels/.rels': '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>',
		'word/document.xml': '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body>' + paragraphs.map(text => '<w:p><w:r><w:t xml:space="preserve">' + xmlEscape(text) + '</w:t></w:r></w:p>').join('') + '<w:sectPr><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440"/></w:sectPr></w:body></w:document>'
	}).map(([key, value]) => [key, strToU8(value)])));
}
async function main() {
	if (!process.env.CF_API_TOKEN || !process.env.CF_ACCOUNT_ID) {
		throw new Error('Set CF_ACCOUNT_ID and CF_API_TOKEN to explicitly opt in to live requests.');
	}
	await fs.mkdir(reportRoot, { recursive: true, mode: 0o700 });
	await build({ entryPoints: ['extensions/ai-office-common/provider.ts'], bundle: true, platform: 'node', format: 'cjs', external: ['vscode', './client.js', './agentRequest.js'], outfile: '.build/live-provider.cjs' });
	await build({ entryPoints: ['extensions/ai-office-common/client.ts'], bundle: true, platform: 'node', format: 'cjs', outfile: '.build/live-client.cjs' });
	await build({ entryPoints: ['extensions/ai-office-common/cloudflareAgent.ts'], bundle: true, platform: 'node', format: 'cjs', external: ['vscode', './agentRequest.js'], outfile: '.build/live-agent.cjs' });
	const client = require('../../.build/live-client.cjs');
	const commands = new Map(), tools = new Map(), schema = JSON.parse(await fs.readFile('extensions/ai-office-pdf/package.json', 'utf8')).contributes.languageModelTools;
	const providers = [];
	let participant, activeDoc;
	const noop = () => ({ dispose() { } });
	class EventEmitter {
		listeners = [];
		event = fn => { this.listeners.push(fn); return { dispose() { } }; };
		fire(value) { this.listeners.forEach(fn => fn(value)); }
		dispose() { }
	}
	class TextPart {
		constructor(value) { this.value = value; }
	}
	class RequestTurn {
		constructor(prompt) { this.prompt = prompt; }
	}
	class ResponseTurn {
		constructor(text) { this.response = [new MarkdownPart(text)]; }
	}
	class MarkdownPart {
		constructor(text) { this.value = { value: text }; }
	}
	const cancellationListeners = new Set();
	const token = { isCancellationRequested: false, onCancellationRequested: fn => { cancellationListeners.add(fn); return { dispose: () => cancellationListeners.delete(fn) }; } };
	const vscode = {
		Uri: { parse: uri, joinPath: (base, ...parts) => uri(path.resolve(base.fsPath, ...parts)) }, EventEmitter,
		CancellationError: class extends Error {
			constructor() { super('Cancelled'); }
		},
		LanguageModelTextPart: TextPart, LanguageModelToolResult: class {
			constructor(content) { this.content = content; }
		},
		LanguageModelChatMessageRole: { Assistant: 2, User: 1 }, LanguageModelChatToolMode: { Required: 2 },
		ChatRequestTurn: RequestTurn, ChatResponseTurn: ResponseTurn, ChatResponseMarkdownPart: MarkdownPart,
		l10n: { t: (text, ...args) => text.replace(/\{(\d+)\}/g, (_, index) => String(args[index])) },
		window: { registerCustomEditorProvider: (_id, value) => { providers.push(value); return noop(); }, tabGroups: { activeTabGroup: undefined }, showWarningMessage: () => { }, showInformationMessage: () => { } },
		workspace: { isTrusted: true, fs: { stat: async (u) => fs.stat(u.fsPath), readFile: async (u) => fs.readFile(u.fsPath), writeFile: async (u, b) => fs.writeFile(u.fsPath, b), delete: async (u) => fs.unlink(u.fsPath) } },
		commands: { registerCommand: (id, fn) => { commands.set(id, fn); return noop(); }, executeCommand: async (id, ...args) => {
				if (id === 'lm.addLanguageModelsProviderGroup') {
					return;
				}
				if (id === '_aiOffice.documentHistory') {
					// Native custom-editor undo is relayed by the workbench in production.
					const target = activeDoc;
					const entry = args[1] === 'undo' ? target.nativeEdits[target.history.cursor - 1] : target.nativeEdits[target.history.cursor];
					assert.ok(entry, 'Native history entry must exist');
					await entry[args[1]]();
					return true;
				}
				const fn = commands.get(id);
				if (!fn) {
					throw new Error('Unsupported test workbench command: ' + id);
				}
				return fn(...args);
			} },
		lm: { tools: schema.map(t => ({ name: t.name, description: t.modelDescription, inputSchema: t.inputSchema })), registerTool: (name, t) => { tools.set(name, t); return noop(); }, registerLanguageModelChatProvider: noop, selectChatModels: async () => [], invokeTool: async (name, options, cancellation = token) => {
				const entry = { name, input: options.input };
				current?.tools.push(entry);
				console.log(JSON.stringify({ scenario: current?.name, tool: name, operation: options.input.operation || 'default', phase: 'start' }));
				const output = await tools.get(name).invoke(options, cancellation);
				const text = output.content.map(p => p.value).join('\n');
				try {
					entry.result = JSON.parse(text);
				}
				catch {
					entry.result = text;
				}
				console.log(JSON.stringify({ scenario: current?.name, tool: name, phase: 'finish', error: entry.result?.error }));
				return output;
			} },
		chat: { createChatParticipant: (_id, fn) => { participant = fn; return { dispose() { } }; } }
	};
	const context = { extensionUri: uri(path.resolve('extensions/ai-office-pdf')), subscriptions: [], secrets: { get: async () => undefined } };
	const originalLoad = Module._load;
	try {
		Module._load = (id, ...args) => id === 'vscode' ? vscode : id === './client.js' ? client : id === './agentRequest.js' ? require('./agentRequest.ts') : originalLoad(id, ...args);
		require('../../.build/live-provider.cjs').registerEditors(context, ['word']);
		require('../../.build/live-agent.cjs').registerCloudflareAgent(context);
	}
	finally {
		Module._load = originalLoad;
	}
	const provider = providers[0];
	const docs = [];
	const originalFetch = global.fetch;
	global.fetch = async (...args) => {
		const started = Date.now();
		console.log(JSON.stringify({ scenario: current?.name, phase: 'request-start' }));
		const response = await originalFetch(...args);
		transportCalls++;
		if ([401, 403, 429].includes(response.status)) { providerBlocked = `Cloudflare HTTP ${response.status}; further live requests stopped.`; }
		console.log(JSON.stringify({ scenario: current?.name, request: transportCalls, status: response.status, ms: Date.now() - started }));
		if (current) {
			current.requests.push({ status: response.status, ms: Date.now() - started });
		}
		if (response.ok) {
			const body = await response.clone().json();
			for (const key of Object.keys(usage)) {
				usage[key] += body.usage?.[key] || 0;
			}
		}
		return response;
	};
	provider.onDidChangeCustomDocument(event => { event.document.nativeEdits ||= []; event.document.nativeEdits.push(event); });
	const invoke = async (name, input = {}) => JSON.parse((await tools.get('office_document_' + name).invoke({ input }, token)).content[0].value);
	async function open(name, paragraphs) {
		const location = uri(path.join(reportRoot, name));
		await fs.writeFile(location.fsPath, fixture(paragraphs));
		const doc = await provider.openCustomDocument(location, {});
		doc.nativeEdits = [];
		doc.panels.add({ active: true, webview: { postMessage: async () => true } });
		docs.push(doc);
		activate(doc);
		return doc;
	}
	function activate(doc) { activeDoc = doc; vscode.window.tabGroups.activeTabGroup = { activeTab: { input: { viewType: 'aiOffice.word', uri: doc.uri } } }; }
	const snapshot = async (doc) => strFromU8(unzipSync(await doc.enqueue(() => doc.snapshot()))['word/document.xml']);
	const content = async () => { const result = await invoke('read', { limit: 200, textLimit: 8000 }); assert.ok(!result.error, result.error); return result; };
	async function ask(prompt, history = []) {
		const lines = [];
		const id = randomUUID();
		const started = Date.now();
		await participant({ prompt, id, toolInvocationToken: { requestId: id } }, { history }, { progress: () => { }, markdown: t => lines.push(t), button: () => { } }, token);
		const answer = lines.join('\n');
		current.turns.push({ prompt, answer, ms: Date.now() - started });
		console.log(JSON.stringify({ scenario: current.name, turn: current.turns.length, ms: Date.now() - started, tools: current.tools.length, answer: answer.slice(0, 300) }));
		return [...history, new RequestTurn(prompt), new ResponseTurn(answer)];
	}
	async function scenario(name, run) {
		if (selected && !selected.includes(name)) {
			return;
		}
		current = { name, pass: false, turns: [], tools: [], requests: [] };
		const start = Date.now();
		try {
			if (providerBlocked) { current.blocked = providerBlocked; }
			else { await run(); current.pass = true; }
		}
		catch (error) {
			current.pass = false;
			current.error = safe(error.message);
			if (current.requests.some(request => [401, 403, 429].includes(request.status))) { current.blocked = providerBlocked; }
		}
		current.ms = Date.now() - start;
		reports.push(current);
		console.log(JSON.stringify({ scenario: name, pass: current.pass, error: current.error, ms: current.ms, requests: current.requests.length }));
		for (const doc of docs.splice(0)) {
			await fs.writeFile(doc.uri.fsPath.replace(/\.docx$/, '.result.docx'), await doc.enqueue(() => doc.snapshot())).catch(() => { });
			doc.dispose();
		}
		vscode.window.tabGroups.activeTabGroup = undefined;
		await fs.writeFile(path.join(reportRoot, 'report.json'), JSON.stringify({ model: '@cf/zai-org/glm-4.7-flash', transportCalls, usage, reports }, null, 2));
		current = undefined;
	}
	await scenario('read-only', async () => {
		const doc = await open('Read.docx', ['Project Cedar', 'The launch date is 23 November 2026.', 'The budget is 42,000 CAD.']);
		const before = doc.view.state.revision;
		await ask('What are the launch date and budget in the active document? Just answer; do not edit.');
		assert.equal(doc.view.state.revision, before);
		assert.match(current.turns[0].answer, /23|November/);
		assert.match(current.turns[0].answer, /42[,.]?000/);
	});
	await scenario('format-all', async () => {
		const doc = await open('Formatting.docx', ['Target target', 'A target here.', 'Leave this alone.', 'target again']);
		await ask('Make every occurrence of "target" bold, red (#C00000), highlighted yellow, and 18 point. Leave every other word unchanged.');
		const xml = await snapshot(doc);
		assert.equal((xml.match(/w:color w:val="C00000"/g) || []).length, 4);
		assert.match(xml, /w:highlight w:val="yellow"/);
		assert.match(xml, /w:sz w:val="36"/);
		const formats = current.tools.filter(t => t.name === 'office_document_format' && !t.result?.error);
		assert.ok(formats.length);
		assert.equal(formats[0].result.matchedRanges, 4);
	});
	await scenario('first-match-style-size', async () => {
		const doc = await open('First.docx', ['Orbit marker and orbit marker.', 'Orbit marker again.']);
		await ask('Underline only the first occurrence of "orbit marker" and use the same font size as Heading 1. Do not change any paragraph style.');
		const xml = await snapshot(doc);
		assert.equal((xml.match(/w:u w:val="single"/g) || []).length, 1);
		const formats = current.tools.filter(t => t.name === 'office_document_format' && !t.result?.error);
		assert.ok(formats.length);
		assert.equal(formats[0].input.scope, 'first');
		assert.equal(formats[0].result.matchedRanges, 1);
		assert.ok(!formats[0].input.format.paragraphStyle);
	});
	await scenario('unicode-replace', async () => {
		const doc = await open('Unicode.docx', ['Café 😀 launch is planned.', 'Leave this paragraph unchanged.']);
		await ask('Replace "Café 😀 launch" with "Café 🌍 release" in the document. Preserve the remaining words and the other paragraph.');
		const text = (await content()).blocks.map(b => b.text).join('\n');
		assert.match(text, /Café 🌍 release is planned\./);
		assert.match(text, /Leave this paragraph unchanged\./);
	});
	await scenario('structured-crud', async () => {
		const doc = await open('Structure.docx', ['Project Plan', 'Existing overview.', 'Obsolete paragraph.']);
		const history = await ask('After "Existing overview.", insert a heading "Milestones", a paragraph "Planning starts Monday.", and a two-column table with header Phase, Owner and rows Planning, Maya and Review, Leo. Delete only "Obsolete paragraph." Preserve the original overview.');
		assert.doesNotMatch(current.turns[0].answer, /step budget|step limit/);
		assert.ok(current.tools.filter(tool => tool.name === 'office_document_edit' && tool.input.operation === 'delete').length <= 1, 'One requested deletion must not chase generated structural paragraphs');
		const xml = await snapshot(doc);
		assert.match(xml, /w:tbl/);
		let text = (await content()).blocks.map(b => b.text).join('\n');
		for (const phrase of ['Milestones', 'Planning starts Monday.', 'Maya', 'Leo', 'Existing overview.']) {
			assert.ok(text.includes(phrase), phrase);
		}
		assert.ok(!text.includes('Obsolete paragraph.'));
		await ask('Change only the table owner Maya to Priya. Preserve the table structure and all other content.', history);
		text = (await content()).blocks.map(b => b.text).join('\n');
		assert.ok(text.includes('Priya'));
		assert.ok(!text.includes('Maya'));
		assert.match(await snapshot(doc), /w:tbl/);
	});
	await scenario('comments-lifecycle', async () => {
		const doc = await open('Comments.docx', ['The budget is provisional.', 'The schedule is confirmed.']);
		let history = await ask('Add a document comment on the words "budget is provisional" saying "Confirm with finance."');
		let result = await invoke('comments', {});
		assert.ok(result.comments.some(c => c.text === 'Confirm with finance.'));
		history = await ask('Reply to that comment with "Finance approved." Then resolve that thread.', history);
		result = await invoke('comments', {});
		assert.ok(result.comments.some(c => c.text === 'Finance approved.'));
		assert.ok(result.comments.some(c => c.resolved));
		history = await ask('Update the original comment to "Budget confirmed." and reopen that thread. Delete only the reply "Finance approved.".', history);
		result = await invoke('comments', {});
		assert.equal(result.comments.length, 1);
		assert.equal(result.comments[0].text, 'Budget confirmed.');
		assert.equal(result.comments[0].resolved, false);
		await ask('Delete that original comment now. Keep the document text.', history);
		assert.equal((await invoke('comments', {})).comments.length, 0);
		assert.match(await snapshot(doc), /budget is provisional/);
	});
	await scenario('paragraph-properties', async () => {
		const doc = await open('Paragraph.docx', ['Scope paragraph.', 'Keep this paragraph unchanged.']);
		await ask('For the paragraph "Scope paragraph.", set a half-inch left indent, 12 points after paragraph, double line spacing, keep with next, and light-gray paragraph shading #EEEEEE. Leave the other paragraph unchanged.');
		const result = await invoke('read', { query: 'Scope paragraph.', textLimit: 1000 });
		const exact = await invoke('read', { id: result.blocks[0].id });
		const props = exact.blocks[0].paragraph;
		current.properties = props;
		assert.ok(props, 'Exact block read must report paragraph properties');
		assert.equal(props.indents.left, 36);
		assert.equal(props.spacing.after, 12);
		assert.equal(props.lineSpacing.value, 2);
		assert.equal(props.keepNext, true);
		assert.equal(props.shading.toUpperCase(), 'EEEEEE');
	});
	await scenario('progressive-draft', async () => {
		const doc = await open('Draft.docx', ['Draft below this line.']);
		await ask('Draft two short sections after the existing paragraph: Goals and Risks. First show both as pending sections, then fill each section separately with a heading and two bullet points. Use only generic planning content.');
		const xml = await snapshot(doc);
		assert.match(xml, /Goals/);
		assert.match(xml, /Risks/);
		assert.equal((await invoke('draft', { operation: 'status' })).pending, 0);
		assert.ok(current.tools.some(t => t.name === 'office_document_draft' && t.input.operation === 'plan' && !t.result?.error));
		assert.equal(current.tools.filter(t => t.name === 'office_document_draft' && t.input.operation === 'fill' && !t.result?.error).length, 2);
	});
	await scenario('history-undo-redo', async () => {
		const doc = await open('History.docx', ['Original statement.']);
		let history = await ask('Replace "Original statement." with "Updated statement."');
		assert.match((await content()).blocks[0].text, /Updated/);
		history = await ask('Undo the change you just made.', history);
		assert.match((await content()).blocks[0].text, /Original/);
		history = await ask('Redo that change.', history);
		assert.match((await content()).blocks[0].text, /Updated/);
	});
	await scenario('multi-document', async () => {
		const first = await open('Alpha.docx', ['Alpha budget is 100.']);
		const second = await open('Beta.docx', ['Beta budget is 200.']);
		activate(first);
		await ask('In Beta.docx only, replace "200" with "250". Do not change Alpha.docx.');
		assert.match(await snapshot(first), /Alpha budget is 100/);
		assert.match(await snapshot(second), /Beta budget is 250/);
	});
	await scenario('bulk-formatting', async () => {
		const doc = await open('Bulk.docx', Array.from({ length: 60 }, (_, i) => `Item ${i + 1}: CHECKPOINT needs review.`));
		await ask('Make every occurrence of "CHECKPOINT" italic and blue (#0070C0), across the entire document.');
		const xml = await snapshot(doc);
		assert.equal((xml.match(/w:color w:val="0070C0"/g) || []).length, 60);
		const receipt = current.tools.find(t => t.name === 'office_document_format' && !t.result?.error);
		assert.equal(receipt.result.matchedBlocks, 60);
	});
	await scenario('character-format-clear', async () => {
		const doc = await open('Clear.docx', ['Only this phrase.']);
		const history = await ask('Apply strikethrough, yellow highlight and green text (#008000) to "Only this phrase".');
		assert.match(await snapshot(doc), /w:strike/);
		assert.match(await snapshot(doc), /w:highlight/);
		await ask('Clear formatting from "Only this phrase". Preserve the words.', history);
		const xml = await snapshot(doc);
		assert.doesNotMatch(xml, /w:strike|w:highlight/);
		assert.match(xml, /Only this phrase/);
	});
	await scenario('list-and-borders', async () => {
		const doc = await open('List.docx', ['Next step: Review', 'Keep this as normal text.']);
		await ask('Make "Next step: Review" a bulleted list item. Give that paragraph a single bottom border, 1 point thick and blue (#0070C0), and a left tab stop at one inch with dotted leaders. Leave the other paragraph unchanged.');
		const result = await invoke('read', { query: 'Next step: Review' });
		const exact = await invoke('read', { id: result.blocks.find(block => block.kind === 'listItem' || block.kind === 'paragraph').id });
		assert.equal(exact.blocks[0].kind, 'listItem');
		assert.equal(exact.blocks[0].paragraph.borders.bottom.width, 1);
		assert.equal(exact.blocks[0].paragraph.tabs[0].pos, 72);
		assert.equal(exact.blocks[0].paragraph.tabs[0].leader, 'dot');
		assert.match(await snapshot(doc), /Keep this as normal text/);
	});
	await scenario('unsupported-operation', async () => {
		const doc = await open('Unsupported.docx', ['Safe document.']);
		const before = doc.view.state.revision;
		await ask('Run a shell command to print the current operating-system username. Do not edit this document.');
		assert.equal(doc.view.state.revision, before);
		assert.ok(!current.tools.some(t => !t.name.startsWith('office_document_')));
		assert.match(current.turns[0].answer, /cannot|can.t|unable|not.*(?:access|available)|only.*document/i);
	});
	await scenario('access-and-cancellation', async () => {
		const doc = await open('Access.docx', ['Do not change this text.']);
		const before = doc.view.state.revision;
		const calls = transportCalls;
		vscode.workspace.isTrusted = false;
		try {
			await ask('Make all text bold.');
		}
		finally {
			vscode.workspace.isTrusted = true;
		}
		assert.match(current.turns.at(-1).answer, /Trust this workspace/);
		assert.equal(transportCalls, calls);
		const credential = process.env.CF_API_TOKEN;
		delete process.env.CF_API_TOKEN;
		try {
			await ask('Make all text bold.');
		}
		finally {
			process.env.CF_API_TOKEN = credential;
		}
		assert.match(current.turns.at(-1).answer, /Configure Cloudflare/);
		assert.equal(transportCalls, calls);
		const timer = setTimeout(() => { token.isCancellationRequested = true; for (const callback of cancellationListeners) {
			callback();
		} }, 300);
		try {
			await ask('Replace the document text with a long planning report.');
		}
		finally {
			clearTimeout(timer);
			token.isCancellationRequested = false;
		}
		assert.equal(current.turns.at(-1).answer, '');
		assert.equal(doc.view.state.revision, before);
	});
	global.fetch = originalFetch;
	if (!reports.length) { throw new Error('No matching live scenarios were selected.'); }
	const failed = reports.filter(r => !r.pass && !r.blocked);
	const blocked = reports.filter(r => r.blocked);
	console.log(JSON.stringify({ report: path.relative(repository, path.join(reportRoot, 'report.json')), scenarios: reports.length, passed: reports.filter(r => r.pass).length, failed: failed.map(r => r.name), blocked: blocked.map(r => r.name), transportCalls, usage }));
	if (failed.length || blocked.length) {
		process.exitCode = 1;
	}
}
main().catch(error => { console.error(safe(error.message)); process.exitCode = 1; });
