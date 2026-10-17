/*---------------------------------------------------------------------------------------------
 *  Copyright (c) Microsoft Corporation. All rights reserved.
 *  Licensed under the MIT License. See License.txt in the project root for license information.
 *--------------------------------------------------------------------------------------------*/
import { test } from 'node:test';
import * as assert from 'node:assert/strict';
import { completeOfficeChat, officeModel, runOfficeAgent, type OfficeMessage, type OfficeTool } from './cloudflareClient.js';

// Every request uses an injected local fake. No account, real token or network is used.
const credentials = { accountId: '0'.repeat(32), token: 'fake-local-unit-test-token' };
const tools: OfficeTool[] = [{ type: 'function', function: { name: 'office_document_edit', description: 'Edit', parameters: { type: 'object' } } }];
const call = (name: string, args: string): OfficeMessage => ({ role: 'assistant', content: null, tool_calls: [{ id: 'call-1', type: 'function', function: { name, arguments: args } }] });

test('Cloudflare transport fixes model and endpoint, disables thinking, and keeps credentials out of body', async () => {
	let captured: { url: string; body: Record<string, unknown>; auth: string } | undefined;
	const reply = await completeOfficeChat(credentials, [{ role: 'user', content: 'Write a heading' }], tools, new AbortController().signal, async (url, init) => {
		captured = { url: String(url), body: JSON.parse(String(init!.body)), auth: (init!.headers as Record<string, string>).Authorization };
		return Response.json({ choices: [{ finish_reason: 'tool_calls', message: call('office_document_edit', '{"operation":"insert"}') }] });
	});
	assert.deepEqual(captured, { url: `https://api.cloudflare.com/client/v4/accounts/${credentials.accountId}/ai/v1/chat/completions`, auth: 'Bearer ' + credentials.token, body: { model: officeModel, messages: [{ role: 'user', content: 'Write a heading' }], stream: false, max_completion_tokens: 4096, temperature: 0.3, chat_template_kwargs: { enable_thinking: false }, tools, tool_choice: 'auto', parallel_tool_calls: false } });
	assert.equal(reply.tool_calls?.[0].function.name, 'office_document_edit');
});

test('Cloudflare failures and cancellation never echo response bodies or retry', async () => {
	for (const status of [401, 403, 429, 500]) {
		let calls = 0;
		await assert.rejects(completeOfficeChat(credentials, [], [], new AbortController().signal, async () => { calls++; return new Response('private upstream text', { status }); }), error => error instanceof Error && !error.message.includes('private upstream text'));
		assert.equal(calls, 1);
	}
	const cancelled = new AbortController(); cancelled.abort();
	await assert.rejects(completeOfficeChat(credentials, [], [], cancelled.signal, async () => { assert.fail('Cancelled requests must not reach transport'); }));
	await assert.rejects(completeOfficeChat(credentials, [], [], new AbortController().signal, async () => Response.json({ choices: [{ finish_reason: 'length', message: call('office_document_edit', '{}') }] })), /incomplete/);
});

test('Word agent only invokes supplied tools, returns errors to model, and stops after cancellation', async () => {
	const invoked: string[] = []; let round = 0;
	const result = await runOfficeAgent({ messages: [], tools, signal: new AbortController().signal, progress: () => {},
		complete: async messages => {
			round++;
			if (round === 1) { return call('execute_shell', '{}'); }
			if (round === 2) { assert.match(messages.at(-1)!.content!, /unavailable/); return call('office_document_edit', '{"id":"internal-only"}'); }
			assert.equal(messages.at(-1)?.role, 'tool'); return { role: 'assistant', content: 'Updated the paragraph.' };
		}, invoke: async name => { invoked.push(name); return '{"ok":true}'; }
	});
	assert.deepEqual({ result, invoked }, { result: 'Updated the paragraph.', invoked: ['office_document_edit'] });
	const cancel = new AbortController();
	await assert.rejects(runOfficeAgent({ messages: [], tools, signal: cancel.signal, progress: () => {}, complete: async () => { cancel.abort(); return call('office_document_edit', '{}'); }, invoke: async () => { assert.fail('No edit after cancellation'); } }));
});

test('Word agent has a bounded step budget', async () => {
	let edits = 0;
	const result = await runOfficeAgent({ messages: [], tools, signal: new AbortController().signal, progress: () => {}, complete: async () => call('office_document_edit', '{}'), invoke: async () => { edits++; return '{}'; } });
	assert.equal(edits, 16); assert.match(result, /step limit/);
});

test('Word agent preserves oversized edit success and new IDs to prevent duplicate edits', async () => {
	let round = 0; let edits = 0;
	const created = [{ id: 'new-block', kind: 'paragraph', parentId: 'body' }];
	await runOfficeAgent({ messages: [], tools, signal: new AbortController().signal, progress: () => {}, complete: async messages => {
		if (++round === 1) { return call('office_document_edit', '{}'); }
		const result = JSON.parse(messages.at(-1)!.content!);
		assert.deepEqual({ completed: result.completed, version: result.version, created: result.created, truncated: result.contentTruncated }, { completed: true, version: 'fresh-version', created, truncated: true });
		return { role: 'assistant', content: 'Updated.' };
	}, invoke: async () => { edits++; return JSON.stringify({ uri: 'file:///memo.docx', version: 'fresh-version', created, undoable: true, saved: false, text: 'x'.repeat(25000) }); } });
	assert.equal(edits, 1);
});
