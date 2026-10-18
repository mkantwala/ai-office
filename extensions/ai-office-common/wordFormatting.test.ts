/*---------------------------------------------------------------------------------------------
 *  Copyright (c) Microsoft Corporation. All rights reserved.
 *  Licensed under the MIT License. See License.txt in the project root for license information.
 *--------------------------------------------------------------------------------------------*/

import { test } from 'node:test';
import assert from 'node:assert/strict';
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
