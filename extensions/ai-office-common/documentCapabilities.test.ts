/*---------------------------------------------------------------------------------------------
 *  Copyright (c) Microsoft Corporation. All rights reserved.
 *  Licensed under the MIT License. See License.txt in the project root for license information.
 *--------------------------------------------------------------------------------------------*/

import { test } from 'node:test';
import assert from 'node:assert/strict';
import { discoverCommands, boundedCommandResult } from './documentCapabilities.js';
import { pdfPageRange } from './pdfTools.js';

test('Capability discovery keeps schemas out of catalogs and paginates command summaries', () => {
	const commands = Array.from({ length: 55 }, (_, index) => ({ id: 'page_' + index, label: 'Page action', params: { largeSchema: 'x'.repeat(2000) }, readOnly: index === 0 }));
	assert.deepEqual(discoverCommands(commands, {}).categories, [{ category: 'pages', count: 55 }]);
	const first = discoverCommands(commands, { category: 'pages', limit: 10 });
	assert.ok(first.mode === 'commands');
	assert.equal(first.commands?.length, 10); assert.equal(first.nextOffset, 10);
	assert.ok(JSON.stringify(first).length < 1500);
	assert.deepEqual(discoverCommands(commands, { command: 'page_22' }).command, commands[22]);
	const last = discoverCommands(commands, { category: 'pages', offset: 50 });
	assert.ok(last.mode === 'commands'); assert.equal(last.nextOffset, undefined);
	assert.throws(() => discoverCommands(commands, { limit: 1000 }), /limit/);
	assert.throws(() => discoverCommands(commands, { command: 'unknown' }), /Unknown command/);
	assert.throws(() => discoverCommands(commands, { category: 'unknown' }), /Unknown category/);
});

test('Large command results are explicitly omitted without truncating JSON', () => {
	assert.deepEqual(boundedCommandResult({ text: 'small' }), { text: 'small' });
	const result = boundedCommandResult({ text: 'x'.repeat(25000) }) as { omitted: boolean; characters: number };
	assert.equal(result.omitted, true); assert.ok(result.characters > 25000);
});

test('PDF page ranges are bounded, ordered and reject ambiguous partial input', () => {
	assert.deepEqual(pdfPageRange('3, 1–2, 2', 5), [1, 2, 3]);
	assert.deepEqual(pdfPageRange('all', 3), [1, 2, 3]);
	for (const value of ['0', '1-1000000', '3-1', '1,', 'x', '', '-1', '1.5']) { assert.throws(() => pdfPageRange(value, 5)); }
	assert.throws(() => pdfPageRange('all', 10001));
});

test('Cached receipts support bounded projections without re-executing native operations', async () => {
	const { projectCommandResult } = await import('./documentCapabilities.js');
	const receipt = { artifacts: Array.from({ length: 120 }, (_, index) => ({ uri: 'file:///result-' + index })), result: [{ text: 'Existing output' }] };
	const page = projectCommandResult(receipt, '/artifacts', 50, 20); assert.ok(page.kind === 'array'); assert.equal(page.total, 120); assert.equal(page.nextOffset, 70); assert.deepEqual(page.items, receipt.artifacts.slice(50, 70));
	const scalar = projectCommandResult(receipt, '/result/0/text'); assert.ok(scalar.kind === 'value'); assert.deepEqual(scalar.value, 'Existing output');
	assert.throws(() => projectCommandResult(receipt, '/__proto__'), /does not exist/);
	assert.throws(() => projectCommandResult(receipt, '/missing'), /does not exist/);
	assert.throws(() => projectCommandResult(receipt, '/artifacts', 0, 101), /limit/);
});

test('Word command discovery groups formatting separately without loading its schemas', () => {
	const commands = [{ id: 'format.spacing', label: 'Character Spacing', params: '{points:number}' }, { id: 'para.properties', label: 'Paragraph' }, { id: 'table.borders', label: 'Borders' }];
	assert.deepEqual(discoverCommands(commands, {}).categories, [{ category: 'formatting', count: 1 }, { category: 'paragraphs', count: 1 }, { category: 'tables', count: 1 }]);
	const list = discoverCommands(commands, { category: 'formatting' });
	assert.ok(list.mode === 'commands'); assert.equal(Object.hasOwn(list.commands[0], 'params'), false);
	assert.equal(discoverCommands(commands, { command: 'format.spacing' }).command?.params, '{points:number}');
});
