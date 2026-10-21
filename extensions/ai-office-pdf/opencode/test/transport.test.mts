/*---------------------------------------------------------------------------------------------
 *  Copyright (c) Microsoft Corporation. All rights reserved.
 *  Licensed under the MIT License. See License.txt in the project root for license information.
 *--------------------------------------------------------------------------------------------*/

import assert from 'node:assert/strict';
import { after, before, test } from 'node:test';
import { mkdtemp, writeFile, rm } from 'node:fs/promises';
import { join } from 'node:path';
import { tmpdir } from 'node:os';
import { OpenCodeConnection } from '../connection.ts';
import { OpenCodeDocumentBridge } from '../documentBridge.ts';

let directory: string;
let executable: string;
before(async () => {
	directory = await mkdtemp(join(tmpdir(), 'sangam-acp-'));
	executable = join(directory, 'peer.cjs');
	await writeFile(executable, `#!/usr/bin/env node
const readline = require('node:readline');
const send = value => process.stdout.write(JSON.stringify(value) + '\\n');
const result = (id, value) => send({jsonrpc:'2.0',id,result:value});
const update = text => send({jsonrpc:'2.0',method:'session/update',params:{sessionId:'native-session',update:{sessionUpdate:'agent_message_chunk',content:{type:'text',text}}}});
let waiting, permission;
readline.createInterface({input:process.stdin}).on('line', line => {
 const message = JSON.parse(line); const p = message.params || {};
 if (message.id === 700) { result(permission,{stopReason: message.result.outcome.outcome === 'selected' ? 'end_turn' : 'cancelled'}); return; }
 if (message.method === 'initialize') result(message.id,{protocolVersion:1,agentCapabilities:{loadSession:true},authMethods:[]});
 else if (message.method === 'session/new') result(message.id,{sessionId:'native-session',configOptions:[{id:'model',name:'Model',type:'select',currentValue:'fake/first',options:[{value:'fake/first',name:'First'},{value:'fake/second',name:'Second'}]}]});
 else if (message.method === 'session/load') {update('Restored history');result(message.id,{});}
 else if (message.method === 'session/set_config_option') result(message.id,{configOptions:[{id:p.configId,name:'Model',type:'select',currentValue:p.value,options:[{value:p.value,name:'Selected'}]}]});
 else if (message.method === 'session/cancel') {if(waiting!==undefined) result(waiting,{stopReason:'cancelled'});}
 else if (message.method === 'session/prompt') {
  const text=p.prompt[0].text;
  if(text==='crash') process.exit(7);
  else if(text==='wait') {waiting=message.id;update('Working');}
  else if(text==='permission') {permission=message.id;send({jsonrpc:'2.0',id:700,method:'session/request_permission',params:{sessionId:'native-session',toolCall:{toolCallId:'call-1',title:'Write file'},options:[{optionId:'once',name:'Allow once',kind:'allow_once'},{optionId:'deny',name:'Deny',kind:'reject_once'}]}});}
  else {update('Hello from OpenCode');result(message.id,{stopReason:'end_turn'});}
 }
});
`, { mode: 0o755 });
});
after(async () => { await rm(directory, { recursive: true, force: true }); });

test('ACP negotiates capabilities, streams messages, switches models and resumes durable IDs', { skip: process.platform === 'win32' }, async () => {
	const updates: string[] = [];
	const connection = new OpenCodeConnection(executable, directory, { sessionUpdate: update => { if (update.update.sessionUpdate === 'agent_message_chunk' && update.update.content.type === 'text') { updates.push(update.update.content.text); } }, requestPermission: async () => ({ outcome: { outcome: 'cancelled' } }) });
	try {
		assert.equal((await connection.start()).agentCapabilities?.loadSession, true);
		assert.equal((await connection.open([])).sessionId, 'native-session');
		assert.equal((await connection.configure('model', 'fake/second')).configOptions[0].currentValue, 'fake/second');
		assert.equal((await connection.prompt('hello')).stopReason, 'end_turn');
		assert.equal((await connection.open([], 'native-session')).sessionId, 'native-session');
		assert.deepEqual(updates, ['Hello from OpenCode', 'Restored history']);
	} finally { connection.dispose(); }
});

test('ACP forwards explicit permission choices and cancels a running turn', { skip: process.platform === 'win32' }, async () => {
	let permissions = 0;
	let working: () => void;
	const started = new Promise<void>(resolve => { working = resolve; });
	const connection = new OpenCodeConnection(executable, directory, { sessionUpdate: () => working(), requestPermission: async request => { permissions++; return { outcome: { outcome: 'selected', optionId: request.options[1].optionId } }; } });
	try {
		await connection.open([]);
		assert.equal((await connection.prompt('permission')).stopReason, 'end_turn');
		const pending = connection.prompt('wait'); await started; await connection.cancel();
		assert.deepEqual({ permissions, reason: (await pending).stopReason }, { permissions: 1, reason: 'cancelled' });
	} finally { connection.dispose(); }
});

test('ACP reports missing executables and unexpected process exits without hanging', { skip: process.platform === 'win32' }, async () => {
	const client = { sessionUpdate: () => {}, requestPermission: async () => ({ outcome: { outcome: 'cancelled' as const } }) };
	const missing = new OpenCodeConnection(join(directory, 'missing'), directory, client);
	await assert.rejects(missing.start(), /Could not start OpenCode/);
	const connection = new OpenCodeConnection(executable, directory, client);
	try { await connection.open([]); await assert.rejects(connection.prompt('crash')); }
	finally { connection.dispose(); }
});

test('MCP only accepts authenticated loopback calls and preserves tool arguments/results', async () => {
	const calls: object[] = [];
	const bridge = new OpenCodeDocumentBridge(() => [{ name: 'office_document_edit', description: 'Edit', inputSchema: { type: 'object' } }, { name: 'shell', description: 'Not a document tool', inputSchema: {} }], async (name, input) => { calls.push({ name, input }); return { content: [{ type: 'text', text: '{"appliedEdits":["p1"]}' }] }; });
	const address = await bridge.start();
	const post = (body: object, authorization = `Bearer ${address.token}`) => fetch(address.url, { method: 'POST', headers: { Authorization: authorization, 'Content-Type': 'application/json' }, body: JSON.stringify(body) });
	try {
		assert.equal((await post({ id: 1, method: 'tools/list' }, 'Bearer invalid')).status, 401);
		const listed = await (await post({ id: 1, method: 'tools/list' })).json();
		assert.deepEqual(listed.result.tools.map((tool: { name: string }) => tool.name), ['office_document_edit']);
		const rejected = await (await post({ id: 2, method: 'tools/call', params: { name: 'shell' } })).text();
		assert.equal(calls.length, 0); assert.equal(rejected, '');
		const result = await (await post({ id: 3, method: 'tools/call', params: { name: 'office_document_edit', arguments: { id: 'p1', ifHash: 'guard', text: 'Revised' } } })).json();
		assert.deepEqual({ calls, result: result.result }, { calls: [{ name: 'office_document_edit', input: { id: 'p1', ifHash: 'guard', text: 'Revised' } }], result: { content: [{ type: 'text', text: '{"appliedEdits":["p1"]}' }] } });
	} finally { bridge.dispose(); }
});
