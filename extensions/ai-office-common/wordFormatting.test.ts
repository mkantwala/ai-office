/*---------------------------------------------------------------------------------------------
 *  Copyright (c) Microsoft Corporation. All rights reserved.
 *  Licensed under the MIT License. See License.txt in the project root for license information.
 *--------------------------------------------------------------------------------------------*/

import { test } from 'node:test';
import assert from 'node:assert/strict';
import { readFile } from 'node:fs/promises';
import { paragraphPropertiesSchema } from './paragraphProperties.js';
import { prepareWordFormatting } from './wordFormatting.js';
import { DocumentHistory } from './agentRequest.js';

test('Word first-match formatting remains bounded when hundreds of blocks match', async () => {
	const reads: Record<string, unknown>[] = [];
	const read = async (args: Record<string, unknown>) => {
		reads.push(args);
		return { total: 500, nextOffset: 1, blocks: [{ id: 'first', kind: 'paragraph', text: '😀 target target', textLength: 15 }] };
	};
	const plan = await prepareWordFormatting({ query: 'target', scope: 'first', format: { bold: true } }, read);
	assert.deepEqual({ reads, edits: plan.edits, matchedBlocks: plan.matchedBlocks, matchedRanges: plan.matchedRanges }, {
		reads: [{ query: 'target', offset: 0, limit: 1, textLimit: 8000 }],
		edits: [{ operation: 'command', id: 'first', command: 'format.bold', params: { value: true }, start: 2, end: 8 }],
		matchedBlocks: 1, matchedRanges: 1
	});
	await assert.rejects(prepareWordFormatting({ query: 'target', scope: 'all', format: { bold: true } }, read), /More than 200 blocks/);
});

test('Word request end boundaries follow branching, checkpoint retention and reload', () => {
	const history = new DocumentHistory();
	history.checkpoint('first'); history.record(1, 'first'); history.record(1, 'first');
	history.checkpoint('second'); history.record(1, 'second');
	history.cursor = 1;
	history.record();
	assert.deepEqual({ first: history.requestEnd('first'), second: history.requestEnd('second'), hasSecond: history.checkpoints.has('second') }, { first: 1, second: undefined, hasSecond: false });
	for (let index = 0; index < 31; index++) {
		const id = 'request-' + index;
		history.checkpoint(id); history.record(1, id);
	}
	assert.equal(history.requestEnd('first'), undefined, 'Evicting a checkpoint also removes its request boundary');
	assert.equal(history.requestEnd('request-0'), undefined);
	history.clear();
	assert.deepEqual({ cursor: history.cursor, end: history.end, checkpoints: history.checkpoints.size, last: history.requestEnd('request-30') }, { cursor: 0, end: 0, checkpoints: 0, last: undefined });
});

test('Word paragraph patches target native text blocks and reject malformed properties before editing', async () => {
	const read = async () => ({ blocks: [{ id: 'paragraph-1', kind: 'paragraph', text: 'Keep this text', textLength: 14 }] });
	const paragraph = { indents: { left: 36, firstLine: -18 }, spacing: { after: 8 }, tabs: [{ pos: 144, leader: 'dot' }], keepNext: true };
	const plan = await prepareWordFormatting({ ids: ['paragraph-1'], format: { paragraph } }, read);
	assert.deepEqual(plan.edits, [{ operation: 'command', id: 'paragraph-1', command: 'para.properties', params: paragraph }]);
	for (const patch of [{ unsupported: true }, { indents: { left: NaN } }, { keepNext: 'yes' }, { lineSpacing: { rule: 'multiple', value: 20 } }, { tabs: [{ pos: 36 }, { pos: 36 }] }, { borders: { bottom: { color: 'red' } } }]) {
		await assert.rejects(prepareWordFormatting({ ids: ['paragraph-1'], format: { paragraph: patch } }, read));
	}
	await assert.rejects(prepareWordFormatting({ ids: ['table-1'], format: { paragraph } }, async () => ({ blocks: [{ id: 'table-1', kind: 'table', text: '', textLength: 0 }] })), /require paragraph/);
});


test('Word paragraph tool schema stays aligned with the shared typed property contract', async () => {
	const manifest = JSON.parse(await readFile('extensions/ai-office-pdf/package.json', 'utf8'));
	const tool = manifest.contributes.languageModelTools.find((tool: { name: string }) => tool.name === 'office_document_format');
	const { description: _description, ...schema } = tool.inputSchema.properties.format.properties.paragraph;
	assert.deepEqual(schema, paragraphPropertiesSchema);
});

test('Word formatting canonicalizes model hex colors without mutating inputs or weakening native patches', async () => {
	let reads = 0;
	const read = async () => { reads++; return { blocks: [{ id: 'p1', kind: 'paragraph', text: 'target', textLength: 6 }] }; };
	const format = { color: '#c00000', paragraph: { shading: '#EEEEEE', borders: { bottom: { color: '#123456', width: 1 } } } };
	const plan = await prepareWordFormatting({ ids: ['p1'], format }, read);
	assert.equal(plan.edits[0].params.color, 'c00000');
	assert.deepEqual(plan.edits[1].params, { shading: 'EEEEEE', borders: { bottom: { color: '123456', width: 1 } } });
	assert.equal(format.color, '#c00000'); assert.equal(format.paragraph.shading, '#EEEEEE');
	const before = reads;
	for (const format of [{ color: '#BAD' }, { color: '#1234567' }, { paragraph: { shading: '#BAD' } }, { paragraph: { borders: { bottom: { color: '#123456', extra: true } } } }]) {
		await assert.rejects(prepareWordFormatting({ ids: ['p1'], format }, read));
	}
	assert.equal(reads, before, 'Invalid colors and unknown fields fail before target reads or edits');
});

test('Word formatting repairs name the correct payload fields and ignore an explicit false clear flag', async () => {
	const read = async () => ({ blocks: [{ id: 'p1', kind: 'paragraph', text: 'target', textLength: 6 }] });
	const plan = await prepareWordFormatting({ query: 'target', format: { bold: true, clear: false } }, read);
	assert.deepEqual(plan.commands, ['format.bold']);
	await assert.rejects(prepareWordFormatting({ query: 'target', format: { shading: '#EEEEEE' } }, read), /format\.paragraph\.shading/);
});
