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
