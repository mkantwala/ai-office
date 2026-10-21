/*---------------------------------------------------------------------------------------------
 *  Copyright (c) Microsoft Corporation. All rights reserved.
 *  Licensed under the MIT License. See License.txt in the project root for license information.
 *--------------------------------------------------------------------------------------------*/

import type { Command } from './types.js';

const wordCategories: Record<string, string> = { format: 'formatting', para: 'paragraphs', styles: 'styles', list: 'lists', table: 'tables', layout: 'pages', references: 'references', equation: 'equations', picture: 'images', arrange: 'images', insert: 'insert' };

export function commandCategory(id: string): string {
	const word = /^(format|para|styles|list|table|layout|references|equation|picture|arrange|insert)\./.exec(id);
	if (word) { return wordCategories[word[1]]; }
	if (/^bookmark_/.test(id)) { return 'bookmarks'; }
	if (/^accessibility_/.test(id)) { return 'accessibility'; }
	if (/^measure_/.test(id)) { return 'measurements'; }
	if (/^(sign_|fill_sign)/.test(id)) { return 'signing'; }
	if (/^(printers|printer_|doc_print)/.test(id)) { return 'printing'; }
	if (/^(js_|action_)/.test(id)) { return 'scripting'; }
	if (/^(pdfa_|print_readiness)/.test(id)) { return 'standards'; }
	if (/^doc_(export|import|create|combine|split|reduce|optimize|compare)/.test(id)) { return 'files'; }

	if (/^page_add_/.test(id)) { return 'content'; }
	if (/^(page_|select\.page)/.test(id)) { return 'pages'; }
	if (/^(comment_|comments_|review\.)/.test(id)) { return 'comments'; }
	if (/^(form_|fill_sign)/.test(id)) { return 'forms'; }
	if (/^(redact_|doc_(hidden_info|remove_hidden|audit_redactions|protect|unprotect))/.test(id)) { return 'security'; }
	if (/^(doc_(header_footer|watermark|background|remove_marks))/.test(id)) { return 'marks'; }
	if (/^(edit[._])/.test(id)) { return 'history'; }
	if (/^(text_|content[._]|object_|image[._]|page_add_)/.test(id)) { return 'content'; }
	if (/^(link_)/.test(id)) { return 'links'; }
	return 'document';
}

/** Schemas are returned only for explicitly selected commands; catalog pages contain short summaries. */
export function discoverCommands(commands: Command[], input: { command?: string; category?: string; offset?: number; limit?: number }) {
	if (input.command !== undefined) {
		const command = commands.find(command => command.id === input.command);
		if (!command) { throw new Error('Unknown command. Discover a category before requesting its schema.'); }
		return { mode: 'schema' as const, command };
	}
	const offset = input.offset ?? 0, limit = input.limit ?? 20;
	if (!Number.isSafeInteger(offset) || offset < 0 || !Number.isSafeInteger(limit) || limit < 1 || limit > 40) { throw new Error('Use offset >= 0 and limit from 1 to 40.'); }
	const categories = [...new Set(commands.map(command => commandCategory(command.id)))].sort();
	if (!input.category) { return { mode: 'categories' as const, categories: categories.map(category => ({ category, count: commands.filter(command => commandCategory(command.id) === category).length })), guidance: 'Select a category, then request the schema for the command you need. Do not load all schemas.' }; }
	if (!categories.includes(input.category)) { throw new Error('Unknown category. Discover available categories first.'); }
	const matches = commands.filter(command => commandCategory(command.id) === input.category);
	return { mode: 'commands' as const, category: input.category, total: matches.length, commands: matches.slice(offset, offset + limit).map(({ id, label, readOnly, destructive }) => ({ id, label, readOnly: !!readOnly, destructive: !!destructive })), ...(offset + limit < matches.length ? { nextOffset: offset + limit } : {}) };
}

/** Never return a partial JSON fragment or a huge document in a tool response. */
export function boundedCommandResult(value: unknown, budget = 24000) {
	const serialized = JSON.stringify(value);
	return serialized && serialized.length > budget ? { omitted: true, characters: serialized.length, guidance: 'Result exceeded the context budget. Retry with a smaller page range or a narrower query.' } : value;
}

/** Read a cached receipt without repeating an expensive read or a document mutation. */
export function projectCommandResult(value: unknown, pointer = '', offset = 0, limit = 40) {
	if (pointer.length > 1000 || pointer !== '' && !pointer.startsWith('/') || !Number.isSafeInteger(offset) || offset < 0 || !Number.isSafeInteger(limit) || limit < 1 || limit > 100) { throw new Error('Use a JSON pointer, offset >= 0 and limit from 1 to 100.'); }
	let selected = value;
	const keys = pointer === '' ? [] : pointer.slice(1).split('/').map(key => key.replaceAll('~1', '/').replaceAll('~0', '~'));
	if (keys.length > 16) { throw new Error('Result projection is too deep.'); }
	for (const key of keys) {
		if (!selected || typeof selected !== 'object' || !Object.hasOwn(selected, key) || ['__proto__', 'constructor', 'prototype'].includes(key)) { throw new Error('Result path does not exist. Inspect the parent keys first.'); }
		selected = (selected as Record<string, unknown>)[key];
	}
	if (Array.isArray(selected)) { return { kind: 'array' as const, path: pointer, total: selected.length, items: boundedCommandResult(selected.slice(offset, offset + limit)), ...(offset + limit < selected.length ? { nextOffset: offset + limit } : {}) }; }
	if (selected && typeof selected === 'object') { const entries = Object.entries(selected); return { kind: 'object' as const, path: pointer, keys: entries.slice(offset, offset + limit).map(([key, value]) => ({ key, type: Array.isArray(value) ? 'array' : typeof value, ...(Array.isArray(value) ? { length: value.length } : {}) })), ...(offset + limit < entries.length ? { nextOffset: offset + limit } : {}) }; }
	return { kind: 'value' as const, path: pointer, value: boundedCommandResult(selected) };
}
