/*---------------------------------------------------------------------------------------------
 *  Copyright (c) Microsoft Corporation. All rights reserved.
 *  Licensed under the MIT License. See License.txt in the project root for license information.
 *--------------------------------------------------------------------------------------------*/

import { normalizeParagraphProperties, paragraphPropertiesSchema } from './paragraphProperties.js';

interface WordBlock { id: string; kind: string; text: string; textLength: number; nextTextOffset?: number | null }
interface WordRead { blocks?: WordBlock[]; total?: number; nextOffset?: number | null; styles?: { id: string; name: string; size: number }[] }
type ReadWord = (args: Record<string, unknown>) => Promise<WordRead>;
interface Edit { operation: string; id: string; command: string; params: Record<string, unknown>; start?: number; end?: number }

const fields = ['paragraphStyle', 'font', 'size', 'sizeFromStyle', 'bold', 'italic', 'underline', 'alignment', 'lineSpacing', 'paragraph', 'strikethrough', 'color', 'highlight', 'clear'];

/** Resolve a complete formatting operation inside the document's serial queue. */
export async function prepareWordFormatting(input: Record<string, unknown>, read: ReadWord) {
	const query = input.query;
	if (query !== undefined && (typeof query !== 'string' || !query || query.length > 500)) { throw new Error('query must contain 1-500 literal characters.'); }
	const ids = input.ids;
	if (ids !== undefined && (!Array.isArray(ids) || !ids.length || ids.length > 200 || ids.some(id => typeof id !== 'string' || !id))) { throw new Error('ids must contain 1-200 block IDs from a current read.'); }
	if (!query && !ids) { throw new Error('Provide query to find text, or ids from a current read.'); }
	if (input.scope !== undefined && input.scope !== 'all' && input.scope !== 'first') { throw new Error('scope must be all or first.'); }
	if (input.target !== undefined && input.target !== 'matches' && input.target !== 'paragraphs') { throw new Error('target must be matches or paragraphs.'); }
	const target = input.target ?? (query ? 'matches' : 'paragraphs');
	if (target === 'matches' && !query) { throw new Error('A literal query is required to format matching text.'); }
	if (!input.format || typeof input.format !== 'object' || Array.isArray(input.format)) { throw new Error('format must be an object with supported formatting properties.'); }
	const format = input.format as Record<string, unknown>;
	for (const key of Object.keys(format)) {
		if (!fields.includes(key)) {
			if (Object.hasOwn(paragraphPropertiesSchema.properties, key)) { throw new Error(`Use format.paragraph.${key}, not format.${key}. Paragraph layout, shading, borders, tabs and pagination belong inside format.paragraph. Preserve the other requested properties in the corrected call.`); }
			throw new Error(`Unsupported format property ${key}. Supported: ${fields.join(', ')}. To replace words, use office_document_edit with operation=replaceText, id and version from read, find=old words, and text=new words. Formatting never changes text.`);
		}
	}
	if (format.size !== undefined && format.sizeFromStyle !== undefined) { throw new Error('Use either size or sizeFromStyle, not both.'); }
	const commands: { command: string; params: Record<string, unknown> }[] = [];
	for (const key of ['paragraphStyle', 'font', 'sizeFromStyle']) {
		if (format[key] !== undefined && (typeof format[key] !== 'string' || !format[key] || format[key].length > 200)) { throw new Error(`${key} must be a non-empty string of at most 200 characters.`); }
	}
	if (typeof format.paragraphStyle === 'string') { commands.push({ command: 'para.style', params: { style: format.paragraphStyle } }); }
	if (typeof format.font === 'string') { commands.push({ command: 'format.font', params: { name: format.font } }); }
	let size = format.size;
	if (typeof format.sizeFromStyle === 'string') {
		const response = await read({ operation: 'styles', query: format.sizeFromStyle });
		const name = format.sizeFromStyle.replace(/\s/g, '').toLowerCase();
		const style = response.styles?.find(style => [style.id, style.name].some(value => value.replace(/\s/g, '').toLowerCase() === name));
		if (!style) { throw new Error('The requested style was not found. Read operation=styles to choose an existing style.'); }
		size = style.size;
	}
	if (size !== undefined) {
		if (typeof size !== 'number' || !Number.isFinite(size) || size < 1 || size > 1638) { throw new Error('size must be between 1 and 1638 points.'); }
		commands.push({ command: 'format.size', params: { size } });
	}
	for (const key of ['bold', 'italic', 'underline', 'strikethrough']) {
		if (format[key] !== undefined) {
			if (typeof format[key] !== 'boolean') { throw new Error(`${key} must be true or false.`); }
			commands.push({ command: 'format.' + key, params: { value: format[key] } });
		}
	}
	if (format.clear !== undefined && format.clear !== false) {
		if (format.clear !== true || Object.keys(format).length !== 1) { throw new Error('clear must be true and used alone.'); }
		commands.push({ command: 'format.clear', params: {} });
	}
	if (format.color !== undefined) {
		if (typeof format.color !== 'string' || !/^(auto|#?[0-9a-fA-F]{6})$/.test(format.color)) { throw new Error('color must be auto or six hexadecimal digits (RRGGBB), with an optional #.'); }
		commands.push({ command: 'format.color', params: { color: format.color.replace(/^#/, '') } });
	}
	if (format.highlight !== undefined) {
		if (typeof format.highlight !== 'string' || !['none', 'yellow', 'brightGreen', 'turquoise', 'pink', 'blue', 'red', 'darkBlue', 'teal', 'green', 'violet', 'darkRed', 'darkYellow', 'gray50', 'gray25', 'black'].includes(format.highlight)) { throw new Error('Unsupported highlight color.'); }
		commands.push({ command: 'format.highlight', params: { color: format.highlight } });
	}
	if (format.alignment !== undefined) {
		if (!['left', 'center', 'right', 'justify', 'distribute', 'start', 'end'].includes(String(format.alignment))) { throw new Error('Unsupported paragraph alignment.'); }
		commands.push({ command: 'para.align', params: { value: format.alignment } });
	}
	if (format.lineSpacing !== undefined) {
		if (typeof format.lineSpacing !== 'number' || !Number.isFinite(format.lineSpacing) || format.lineSpacing < .5 || format.lineSpacing > 10) { throw new Error('lineSpacing must be between 0.5 and 10.'); }
		commands.push({ command: 'para.lineSpacing', params: { value: format.lineSpacing } });
	}
	if (format.paragraph !== undefined) {
		const paragraph = normalizeParagraphProperties(format.paragraph);
		if (format.alignment !== undefined && paragraph.alignment !== undefined || format.lineSpacing !== undefined && paragraph.lineSpacing !== undefined) { throw new Error('Specify alignment and line spacing only once.'); }
		commands.push({ command: 'para.properties', params: paragraph as Record<string, unknown> });
	}
	if (!commands.length) { throw new Error('Provide at least one supported format property.'); }
	const blocks: WordBlock[] = [];
	if (Array.isArray(ids)) {
		for (const id of new Set(ids)) {
			const response = await read({ id, textLimit: 8000 });
			const block = response.blocks?.find(block => block.id === id);
			if (!block) { throw new Error('A target no longer exists. Read the document again.'); }
			blocks.push(block);
		}
	} else {
		let offset = 0;
		for (; ;) {
			const response = await read({ query, offset, limit: input.scope === 'first' ? 1 : 20, textLimit: 8000 });
			blocks.push(...response.blocks || []);
			if ((input.scope !== 'first' && (response.total ?? blocks.length) > 200) || blocks.length > 200) { throw new Error('More than 200 blocks match. Narrow the query or provide smaller groups of IDs. No formatting was applied.'); }
			if ((response.nextOffset === null || response.nextOffset === undefined) || input.scope === 'first' && blocks.length) { break; }
			if (response.nextOffset <= offset) { throw new Error('Document search did not advance. No formatting was applied.'); }
			offset = response.nextOffset;
		}
	}
	const edits: Edit[] = [];
	const targets = new Set<string>();
	let ranges = 0, characters = 0;
	for (const block of blocks) {
		if (format.paragraph !== undefined && !['paragraph', 'heading', 'listItem'].includes(block.kind)) { throw new Error('Paragraph properties require paragraph, heading or list-item IDs.'); }
		const matches: { start: number; end: number }[] = [];
		if (query) {
			if (block.kind === 'table' || block.kind === 'row' || block.kind === 'cell') { throw new Error('For text matching, target paragraph IDs or use a query without structural IDs.'); }
			let text = block.text;
			let offset = block.nextTextOffset;
			while (offset !== null && offset !== undefined) {
				if (text.length > 100000 || characters + text.length > 1000000) { throw new Error('Text exceeds the formatting search limit. Narrow the target.'); }
				const response = await read({ id: block.id, textOffset: offset, textLimit: 8000 });
				const part = response.blocks?.[0];
				if (!part || part.nextTextOffset !== null && part.nextTextOffset !== undefined && part.nextTextOffset <= offset) { throw new Error('Text search did not advance. No formatting was applied.'); }
				text += part.text; offset = part.nextTextOffset;
			}
			characters += text.length;
			if (characters > 1000000) { throw new Error('Text exceeds the formatting search limit. Narrow the target.'); }
			const expression = new RegExp(String(query).replace(/[.*+?^${}()|[\]\\]/g, '\\$&'), 'giu');
			for (const match of text.matchAll(expression)) {
				const start = [...text.slice(0, match.index)].length;
				matches.push({ start, end: start + [...match[0]].length });
				if (input.scope === 'first') { break; }
			}
			if (!matches.length) { continue; }
		} else { matches.push({ start: 0, end: block.textLength }); }
		ranges += matches.length; targets.add(block.id);
		for (const command of commands) {
			const selections = target === 'matches' && command.command.startsWith('format.') ? matches : [undefined];
			for (const range of selections) {
				edits.push({ operation: 'command', id: block.id, ...command, ...range });
				if (edits.length > 200) { throw new Error('The complete change exceeds 200 formatting actions. Use smaller groups of IDs or a narrower query. No formatting was applied.'); }
			}
		}
		if (input.scope === 'first') { break; }
	}
	if (!edits.length) { throw new Error('No matching text was found. No formatting was applied.'); }
	return { edits, matchedBlocks: targets.size, matchedRanges: ranges, scope: input.scope ?? 'all', target, commands: commands.map(command => command.command) };
}

/** Do not silently ignore a formatting command attached to a text replacement. */
export function validateDocumentEdits(input: Record<string, unknown>) {
	const edits = input.operation === 'batch' ? input.edits : [input];
	if (!Array.isArray(edits)) { throw new Error('operation=batch requires an edits array.'); }
	const common = ['operation', 'id', 'uri', 'version', 'ifHash', 'allowDetach'];
	const properties: Record<string, string[]> = { append: ['blocks'], insert: ['position', 'blocks'], update: ['text', 'start', 'end'], replaceText: ['find', 'nth', 'text'], replaceBlocks: ['blocks'], delete: [], command: ['command', 'params', 'start', 'end'] };
	for (const edit of edits) {
		if (!edit || typeof edit !== 'object' || Array.isArray(edit)) { throw new Error('Each edit must be an object.'); }
		const operation = edit.operation;
		if (operation === 'append' && (edit.id !== undefined || edit.ifHash !== undefined || edit.position !== undefined)) { throw new Error('Append resolves the document end internally. Omit id, position and ifHash.'); }
		if (typeof operation !== 'string' || !Object.hasOwn(properties, operation)) { throw new Error('Use append, insert, update, replaceText, replaceBlocks, delete or command for each edit.'); }
		if (edit.ifHash !== undefined && (typeof edit.ifHash !== 'string' || !/^[a-f0-9]{64}$/i.test(edit.ifHash))) { throw new Error('ifHash must be the full block hash from a read.'); }
		if (edit.allowDetach !== undefined && typeof edit.allowDetach !== 'boolean') { throw new Error('allowDetach must be an explicit boolean.'); }
		for (const key of Object.keys(edit)) {
			if (![...common, ...properties[operation]].includes(key)) { throw new Error(`The ${operation} operation does not accept ${key}. For formatting, use office_document_format with query or ids and format properties. update only replaces text.`); }
		}
		if (operation === 'replaceText' && (typeof edit.find !== 'string' || !edit.find || edit.find.length > 2000 || edit.nth !== undefined && (!Number.isInteger(edit.nth) || edit.nth < 1))) { throw new Error('replaceText requires literal find text and optional positive nth.'); }
		if (['update', 'replaceText'].includes(operation) && typeof edit.text !== 'string') { throw new Error(`${operation} requires text containing the NEW replacement words. Use {operation:"replaceText", id:<block ID from read>, version:<version from read>, find:"old words", text:"new words"}. Use office_document_format only for appearance changes.`); }
	}
}

/** Only unchanged text targets can bypass unrelated document revisions, never structural operations. */
export function canRebaseTextEdits(edits: Record<string, unknown>[]) {
	return edits.length > 0 && edits.every(edit => ['update', 'replaceText'].includes(String(edit.operation)) && typeof edit.ifHash === 'string' && /^[a-f0-9]{64}$/i.test(edit.ifHash));
}
