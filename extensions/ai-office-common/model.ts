/*---------------------------------------------------------------------------------------------
 *  Copyright (c) Microsoft Corporation. All rights reserved.
 *  Licensed under the MIT License. See License.txt in the project root for license information.
 *--------------------------------------------------------------------------------------------*/

import { readFile } from 'node:fs/promises';
import { join } from 'node:path';
import { pathToFileURL } from 'node:url';
import { randomUUID } from 'node:crypto';
import { readWordIdentities, writeWordIdentities } from './wordIdentity.js';
import { SlidesEngine } from './slides.js';
import type { Engine, Kind, State, Command } from './types.js';

interface WasmDocument {
	pages(): string; state(): string; commands(): string; execute(id: string, args: string): string;
	tables?(page: number): string;
	export(): Uint8Array; render(page: number, scale: number): Uint8Array;
	image?(page: number, name: string, bytes: Uint8Array): string;
	width: number; height: number; free(): void;
	agent_init?(prefix: string, metadata: string): void;
	agent_metadata?(): string;
	agent_read?(args: string): string;
	agent_edit?(args: string): string;
	agent_rects?(id: string): string;
	review_start?(): void; review_finish?(): void; review_read?(): string;
	review_decide?(revision: bigint, action: string, ids: string): string;
	review_export?(): string; review_import?(value: string): void; review_projection?(): string;
}

/**
 * Agent-facing create, read, update and delete operations that target content by
 * text. The engine-specific target (Word selection, PDF paragraph, PowerPoint
 * shape or table cell) is resolved internally, so callers never carry document
 * or shape ids between inspect and run.
 */
const contentCommands: Command[] = [
	{ id: 'content.find', label: 'Find text', params: { find: 'string', occurrence: 'match number (default 1)', page: 'PDF/slides page (default current)' } },
	{ id: 'content.replace', label: 'Replace text', params: { find: 'string', text: 'replacement string', occurrence: 'match number (default 1)', all: 'replace every match (default false)', page: 'PDF/slides page (default current)' } },
	{ id: 'content.delete', label: 'Delete text', params: { find: 'string', occurrence: 'match number (default 1)', all: 'delete every match (default false)', page: 'PDF/slides page (default current)' } },
	{ id: 'content.append', label: 'Append text', params: { text: 'string', page: 'PDF/slides page (default last)' } }
];

function textArgument(value: unknown, name: string): string {
	if (typeof value !== 'string' || !value) { throw new Error(name + ' is required.'); }
	if (value.length > 100000) { throw new Error(name + ' is limited to 100,000 characters.'); }
	return value;
}

function pageArgument(value: unknown): number {
	const page = Number(value);
	if (!Number.isInteger(page) || page < 1 || page > 1000000) { throw new Error('page must be a positive integer.'); }
	return page;
}

function occurrenceArgument(value: unknown): number {
	if (value === undefined) { return 1; }
	const occurrence = Number(value);
	if (!Number.isInteger(occurrence) || occurrence < 1 || occurrence > 10000) { throw new Error('occurrence must be an integer from 1 to 10000.'); }
	return occurrence;
}

function missing(find: string): never {
	throw new Error('`' + find + '` not found.');
}

function countMatches(text: string, find: string): number {
	let count = 0, index = 0;
	for (; ;) {
		const at = text.indexOf(find, index);
		if (at < 0) { return count; }
		count++; index = at + find.length;
	}
}

function replaceFirst(text: string, find: string, replacement: string): string {
	const at = text.indexOf(find);
	return at < 0 ? text : text.slice(0, at) + replacement + text.slice(at + find.length);
}

function replaceEvery(text: string, find: string, replacement: string): string {
	return text.split(find).join(replacement);
}

function withContentCommands(kind: Kind, base: Engine): Engine {
	const run = (id: string, args: Record<string, unknown>) => base.execute(id, args);

	/** Selects the requested 1-based page so the returned state describes it. */
	async function pageState(page?: unknown): Promise<State> {
		if (page === undefined) { return base.state(); }
		const number = pageArgument(page);
		const total = base.pages().length;
		if (number > total) { throw new Error('Page ' + number + ' does not exist. The document has ' + total + ' page' + (total === 1 ? '' : 's') + '.'); }
		return run(kind === 'pdf' ? 'select.page' : 'select.shape', { page: number });
	}

	async function find(args: Record<string, unknown>): Promise<State> {
		const query = textArgument(args.find, 'find');
		const occurrence = occurrenceArgument(args.occurrence);
		if (kind === 'word') { return run('select.text', { text: query, occurrence }); }
		const current = await pageState(args.page);
		const matches = (current.blocks || []).filter(block => block.text.includes(query)).slice(0, 100);
		if (!matches.length) { missing(query); }
		return { ...current, matches };
	}

	async function edit(args: Record<string, unknown>, remove: boolean): Promise<State> {
		const query = textArgument(args.find, 'find');
		const replacement = remove ? '' : textArgument(args.text, 'text');
		const all = args.all === true;
		const occurrence = occurrenceArgument(args.occurrence);
		if (kind === 'word') {
			if (all) {
				const full = (await run('select.all', {})).selectedText || '';
				const count = countMatches(full, query);
				if (!count) { missing(query); }
				if (count > 500) { throw new Error('More than 500 matches; use occurrence or a longer find text.'); }
				// Replace from the last match backwards so earlier occurrence numbers stay valid.
				let result: State | undefined;
				for (let index = count; index >= 1; index--) {
					await run('select.text', { text: query, occurrence: index });
					result = replacement ? await run('text.insert', { text: replacement }) : await run('text.delete', {});
				}
				return result!;
			}
			await run('select.text', { text: query, occurrence });
			return replacement ? run('text.insert', { text: replacement }) : run('text.delete', {});
		}
		const current = await pageState(args.page);
		const matches = (current.blocks || []).filter(block => block.text.includes(query));
		if (!matches.length) { missing(query); }
		let targets: typeof matches;
		if (all) { targets = matches; }
		else {
			const target = matches[occurrence - 1];
			if (!target) { throw new Error('Occurrence ' + occurrence + ' of `' + query + '` was not found.'); }
			targets = [target];
		}
		const page = (current.page || 0) + 1;
		let result: State | undefined;
		// Edit from the last match backwards: deleting a PDF paragraph renumbers the paragraphs after it.
		for (const block of [...targets].reverse()) {
			const text = all ? replaceEvery(block.text, query, replacement) : replaceFirst(block.text, query, replacement);
			result = await run(kind === 'pdf' ? 'text_edit' : 'text.replace', kind === 'pdf'
				? { page, paragraph: block.id, text }
				: { page, shape: block.id, ...(block.row !== undefined ? { row: block.row, col: block.col } : {}), text });
		}
		return result!;
	}

	async function append(args: Record<string, unknown>): Promise<State> {
		const text = textArgument(args.text, 'text');
		if (kind === 'word') {
			await run('caret.docEnd', {});
			await run('text.newParagraph', {});
			return run('text.insert', { text });
		}
		const sizes = base.pages();
		const last = sizes.length;
		const page = args.page === undefined ? last : pageArgument(args.page);
		if (page > last) { throw new Error('Page ' + page + ' does not exist. The document has ' + last + ' page' + (last === 1 ? '' : 's') + '.'); }
		const current = await pageState(page);
		const blocks = current.blocks || [];
		const [width, pageHeight] = sizes[page - 1];
		const bottom = blocks.reduce((value, block) => Math.max(value, block.rect[3]), 0);
		if (kind === 'pdf') {
			// Keep the new paragraph separated from an existing one; a smaller gap is
			// grouped into the paragraph above by the PDF line editor.
			const line = blocks.reduce((value, block) => block.rect[3] === bottom ? Math.max(value, block.rect[3] - block.rect[1]) : value, 0);
			const y = Math.min(Math.max(blocks.length ? Math.round(bottom + Math.max(28, line * 1.5)) : 72, 72), pageHeight - 36);
			return run('page_add_text', { page, at: [72, y], width: Math.max(width - 144, 72), text });
		}
		const y = Math.min(Math.max(bottom / 72 + .25, 1), Math.max(pageHeight / 72 - 1.4, 1));
		return run('insert.text', { page, text, at: [1, y] });
	}

	return {
		agent: base.agent,
		pages: () => base.pages(),
		state: () => base.state(),
		commands: () => [...base.commands(), ...contentCommands],
		execute: async (id, args) => {
			switch (id) {
				case 'content.find': return find(args);
				case 'content.replace': return edit(args, false);
				case 'content.delete': return edit(args, true);
				case 'content.append': return append(args);
				default:
					if (id.startsWith('content.')) { throw new Error('Unknown content command: ' + id); }
					return base.execute(id, args);
			}
		},
		export: () => base.export(),
		render: (page, scale) => base.render(page, scale),
		dispose: () => base.dispose()
	};
}

export async function openEngine(kind: Kind, name: string, bytes: Uint8Array, extensionsRoot: string, password?: string): Promise<Engine> {
	if (bytes.length > 128 * 1024 * 1024) { throw new Error('Document exceeds the 128 MB limit.'); }
	if (bytes.length < 22) { throw new Error('This file is empty or incomplete. Create a valid document in its original application, then open it here.'); }
	if (kind === 'slides') { return withContentCommands(kind, await SlidesEngine.open(bytes)); }
	const project = kind === 'word' ? 'wordcraft' : 'pdfcraft';
	const directory = join(extensionsRoot, 'ai-office-pdf/media', kind === 'word' ? 'word-engine' : 'engine');
	// Generated engine loaders are selected at runtime and kept outside the host bundle.
	// eslint-disable-next-line no-restricted-syntax
	const module = await import(pathToFileURL(join(directory, project + '.js')).href);
	await module.default({ module_or_path: await readFile(join(directory, project + '_bg.wasm')) });
	const doc: WasmDocument = new module.OfficeDocument(name, bytes, password);
	if (kind === 'word' && doc.agent_init) {
		// Encrypted files are not ZIP packages until the engine decrypts them.
		doc.agent_init(randomUUID(), JSON.stringify(password ? [] : readWordIdentities(bytes)));
	}
	const state = (serialized = doc.state()): State => {
		const value = JSON.parse(serialized);
		if (kind === 'pdf') { value.blocks = value.blocks.map((block: { paragraph: number; text: string; editor_text: string; rect: number[]; font: string; size: number }) => ({ id: block.paragraph, text: block.editor_text || block.text, rect: block.rect, font: block.font, size: block.size })); }
		return value;
	};
	const engine: Engine = {
		agent: kind === 'word' && doc.agent_read && doc.agent_edit && doc.agent_rects ? {
			read: args => JSON.parse(doc.agent_read!(JSON.stringify(args))),
			edit: async args => JSON.parse(doc.agent_edit!(JSON.stringify(args))),
			rects: id => JSON.parse(doc.agent_rects!(id)),
			reviewStart: () => { if (!doc.review_start) { throw new Error('Rebuild the Word engine to enable change review.'); } doc.review_start(); },
			reviewFinish: () => doc.review_finish?.(),
			reviewRead: () => JSON.parse(doc.review_read?.() || '{"changes":[]}'),
			reviewDecide: (revision, action, ids) => { if (!doc.review_decide) { throw new Error('Rebuild the Word engine to enable change review.'); } return state(doc.review_decide(BigInt(revision), action, JSON.stringify(ids))); },
			reviewExport: () => doc.review_export?.() || '',
			reviewImport: value => doc.review_import?.(value),
			reviewProjection: () => doc.review_projection?.() || ''
		} : undefined,
		pages: () => JSON.parse(doc.pages()), state,
		commands: () => {
			const commands: Command[] = JSON.parse(doc.commands());
			if (kind === 'pdf') { commands.push({ id: 'image.insert', label: 'Image', params: { page: '1-based page', data: 'base64 image', name: 'filename' } }, { id: 'select.page', label: 'Select page', params: { page: '1-based page' } }); }
			return commands;
		},
		execute: async (id, args) => {
			if (kind === 'word') { return state(doc.execute(id, JSON.stringify(args))); }
			if (id === 'image.insert' && doc.image) {
				if (typeof args.data !== 'string' || args.data.length > 24 * 1024 * 1024) { throw new Error('Image data is missing or too large.'); }
				const page = Number(args.page);
				if (!Number.isInteger(page) || page < 1) { throw new Error('Page must be a positive integer.'); }
				doc.image(page - 1, typeof args.name === 'string' ? args.name : 'image.png', Buffer.from(args.data, 'base64'));
			} else { doc.execute(id, JSON.stringify(args)); }
			return state();
		},
		export: async () => {
			const bytes = doc.export();
			return kind === 'word' && doc.agent_metadata && !password ? writeWordIdentities(bytes, JSON.parse(doc.agent_metadata())) : bytes;
		},
		render: (page, scale) => ({ pixels: doc.render(page, scale), width: doc.width, height: doc.height, ...(doc.tables ? { tables: JSON.parse(doc.tables(page)) } : {}) }),
		dispose: () => doc.free()
	};
	return withContentCommands(kind, engine);
}
