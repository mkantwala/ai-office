/*---------------------------------------------------------------------------------------------
 *  Copyright (c) Microsoft Corporation. All rights reserved.
 *  Licensed under the MIT License. See License.txt in the project root for license information.
 *--------------------------------------------------------------------------------------------*/

import type { Command } from './types.js';

export function commandCategory(id: string): string {
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
