/*---------------------------------------------------------------------------------------------
 *  Copyright (c) Microsoft Corporation. All rights reserved.
 *  Licensed under the MIT License. See License.txt in the project root for license information.
 *--------------------------------------------------------------------------------------------*/

import * as ppt from '../ai-office-slides/node_modules/@office-kit/pptx/dist/index.js';
import { renderSlideToSvg } from '../ai-office-slides/node_modules/@office-kit/pptx-preview/dist/index.js';
import { checkPresentationZip } from '../ai-office-slides/zipLimits.js';
import type { Engine, State, Command } from './types.js';

const definitions: Command[] = [
	{ id: 'select.shape', label: 'Select text box', params: { page: '1-based slide', shape: 'shape id' } },
	{ id: 'text.replace', label: 'Edit text', params: { page: '1-based slide', shape: 'shape id', text: 'string' } },
	{ id: 'insert.text', label: 'Insert text', params: { page: '1-based slide', text: 'string', at: '[x, y] inches (optional, default [1, 1])' } },
	{ id: 'format.text', label: 'Format text', params: { page: 'slide', shape: 'id', range: '{start,end} UTF-16 offsets (optional)', font: 'name', size: 'points', bold: 'boolean', italic: 'boolean', underline: 'boolean' } },
	{ id: 'para.align', label: 'Align', params: { value: 'left|center|right|justify', page: 'slide', shape: 'id' } },
	{ id: 'para.list', label: 'List', params: { value: 'bullet|number|none', page: 'slide', shape: 'id' } },
	{ id: 'para.spacing', label: 'Spacing', params: { value: 'multiple', page: 'slide', shape: 'id' } },
	{ id: 'insert.table', label: 'Table', params: { page: 'slide', rows: '1..20', cols: '1..10' } },
	{ id: 'insert.image', label: 'Image', params: { page: 'slide', data: 'base64' } },
	{ id: 'insert.link', label: 'Link', params: { page: 'slide', shape: 'id', url: 'https URL' } },
	{ id: 'insert.comment', label: 'Comment', params: { page: 'slide', text: 'string' } },
	{ id: 'edit.undo', label: 'Undo' }, { id: 'edit.redo', label: 'Redo' }
];
const num = (value: unknown, min: number, max: number) => {
	if (typeof value !== 'number' || !Number.isFinite(value) || value < min || value > max) { throw new Error('Number outside allowed range: ' + min + '..' + max); }
	return value;
};
export class SlidesEngine implements Engine {
	private page = 0;
	private revision = 0;
	private past: Uint8Array[] = [];
	private future: Uint8Array[] = [];
	private constructor(private deck: ppt.PresentationData) { }
	static async open(bytes: Uint8Array) {
		checkPresentationZip(bytes);
		const engine = new SlidesEngine(await ppt.loadPresentation(bytes));
		if (ppt.getSlides(engine.deck).length > 2000) { throw new Error('Presentation exceeds 2,000 slides.'); }
		return engine;
	}
	commands() { return definitions; }
	pages() {
		const size = ppt.getSlideSize(this.deck);
		return ppt.getSlides(this.deck).map(() => [Number(size?.width || 9144000) / 12700, Number(size?.height || 5143500) / 12700]);
	}
	state(): State {
		const slide = ppt.getSlides(this.deck)[this.page];
		return {
			revision: this.revision, page: this.page, undo: !!this.past.length, redo: !!this.future.length,
			comments: slide ? ppt.getSlideComments(slide).map(c => ({ author: ppt.getCommentAuthor(c).name, text: ppt.getCommentText(c) })) : [],
			blocks: slide ? ppt.getSlideShapes(slide).flatMap(shape => {
				const rect = ppt.getShapeBoundsResolved(this.deck, shape);
				if (!rect) { return []; }
				if (ppt.isTableShape(shape)) {
					const widths = ppt.getTableColumnWidths(shape).map(Number), heights = ppt.getTableRowHeights(shape).map(Number);
					return ppt.getTableCells(shape).flatMap((row, r) => row.map((cell, c) => {
						const x = Number(rect.x) + widths.slice(0, c).reduce((a, b) => a + b, 0), y = Number(rect.y) + heights.slice(0, r).reduce((a, b) => a + b, 0);
						return { id: ppt.getShapeId(shape), row: r, col: c, text: ppt.getTableCellText(cell), rect: [x / 12700, y / 12700, (x + widths[c]) / 12700, (y + heights[r]) / 12700], size: 14 };
					}));
				}
				if (!ppt.hasShapeText(shape)) { return []; }
				return [{ id: ppt.getShapeId(shape), text: ppt.getShapeText(shape), rect: [Number(rect.x) / 12700, Number(rect.y) / 12700, Number(rect.x + rect.w) / 12700, Number(rect.y + rect.h) / 12700], size: 18 }];
			}) : []
		};
	}
	async execute(id: string, args: Record<string, unknown>): Promise<State> {
		if (!definitions.some(item => item.id === id)) { throw new Error('Unknown slide command: ' + id); }
		if (id === 'edit.undo' || id === 'edit.redo') {
			const from = id === 'edit.undo' ? this.past : this.future;
			const to = id === 'edit.undo' ? this.future : this.past;
			const bytes = from.at(-1);
			if (!bytes) { throw new Error('Nothing to ' + id.slice(5)); }
			const restored = await ppt.loadPresentation(bytes);
			to.push(await this.export()); from.pop(); this.deck = restored; this.revision++; return this.state();
		}
		const page = args.page === undefined ? this.page : num(args.page, 1, ppt.getSlides(this.deck).length) - 1;
		if (!Number.isInteger(page)) { throw new Error('Slide must be an integer.'); }
		this.page = page;
		const slide = ppt.getSlides(this.deck)[page];
		if (!slide) { throw new Error('Slide does not exist.'); }
		const shape = ppt.getSlideShapes(slide).find(item => ppt.getShapeId(item) === args.shape);
		if (id === 'select.shape') { return this.state(); }
		const before = await this.export();
		try {
			if (id === 'insert.text') {
				if (typeof args.text !== 'string' || args.text.length > 100000) { throw new Error('Text missing or too large.'); }
				let at = { x: 1, y: 1 };
				if (args.at !== undefined) {
					if (!Array.isArray(args.at) || args.at.length !== 2) { throw new Error('at must be an [x, y] pair in inches.'); }
					at = { x: num(args.at[0], 0, 100), y: num(args.at[1], 0, 100) };
				}
				ppt.addSlideTextBox(slide, { x: ppt.inches(at.x), y: ppt.inches(at.y), w: ppt.inches(6), h: ppt.inches(1), text: args.text });
			} else if (id === 'insert.table') {
				const rows = num(args.rows, 1, 20), cols = num(args.cols, 1, 10);
				if (!Number.isInteger(rows) || !Number.isInteger(cols)) { throw new Error('Table dimensions must be integers.'); }
				ppt.addSlideTable(slide, { x: ppt.inches(1), y: ppt.inches(1), w: ppt.inches(6), h: ppt.inches(Math.min(rows * .4, 4)), rows: Array.from({ length: rows }, () => Array.from({ length: cols }, () => '')) });
			} else if (id === 'insert.image') {
				if (typeof args.data !== 'string' || args.data.length > 24 * 1024 * 1024) { throw new Error('Image missing or too large.'); }
				ppt.addSlideImage(slide, Buffer.from(args.data, 'base64'), { x: ppt.inches(1), y: ppt.inches(1), w: ppt.inches(4), h: ppt.inches(3), fit: 'contain' });
			} else if (id === 'insert.comment') {
				if (typeof args.text !== 'string' || !args.text.trim()) { throw new Error('Enter comment text.'); }
				ppt.addSlideComment(slide, { author: { name: 'AI Office' }, text: args.text });
			} else {
				if (!shape) { throw new Error('Select a text box first.'); }
				const cell = ppt.isTableShape(shape) && args.row !== undefined && args.col !== undefined ? ppt.getTableCell(shape, num(args.row, 0, 1000), num(args.col, 0, 1000)) : undefined;
				if (id === 'text.replace') {
					if (typeof args.text !== 'string' || args.text.length > 100000) { throw new Error('Text missing or too large.'); }
					if (cell) { ppt.setTableCellText(cell, args.text, { preserveFormatting: true }); }
					else { ppt.setShapeText(shape, args.text, { preserveFormatting: true }); }
				} else if (id === 'format.text') {
					const format: ppt.TextFormat = {};
					for (const key of ['bold', 'italic', 'underline'] as const) { if (typeof args[key] === 'boolean') { format[key] = args[key]; } }
					if (typeof args.font === 'string' && args.font.length <= 100) { format.font = args.font; }
					if (args.size !== undefined) { format.size = num(args.size, 1, 400); }
					let range: { start: number; end: number } | undefined;
					if (args.range && typeof args.range === 'object') {
						const value = args.range as { start: number; end: number };
						const length = cell ? ppt.getTableCellText(cell).length : ppt.getShapeText(shape).length;
						const start = num(value.start, 0, length), end = num(value.end, start, length);
						if (end > start) { range = { start, end }; }
					}
					if (cell) { ppt.setTableCellTextFormat(cell, format, { range }); }
					else { ppt.setShapeTextFormat(shape, format, { range }); }
				} else if (id === 'para.align') {
					if (!['left', 'center', 'right', 'justify'].includes(String(args.value))) { throw new Error('Invalid alignment.'); }
					if (cell) { ppt.setTableCellAlignment(cell, args.value as ppt.ParagraphAlignment); }
					else { ppt.setShapeAlignment(shape, args.value as ppt.ParagraphAlignment); }
				} else if (id === 'para.list') {
					if (!['bullet', 'number', 'none'].includes(String(args.value))) { throw new Error('Invalid list style.'); }
					if (cell) { for (let i = 0; i < ppt.getTableCellText(cell).split('\n').length; i++) { ppt.setParagraphBullet(cell, i, args.value as ppt.BulletStyle); } }
					else { ppt.setShapeBulletStyle(shape, args.value as ppt.BulletStyle); }
				} else if (id === 'para.spacing') {
					const value = num(args.value, .5, 3);
					for (let i = 0; i < (cell ? ppt.getTableCellText(cell) : ppt.getShapeText(shape)).split('\n').length; i++) { ppt.setParagraphLineSpacing(cell || shape, i, { kind: 'pct', value: value * 100 }); }
				} else if (id === 'insert.link') {
					const url = new URL(String(args.url));
					if (!['https:', 'http:', 'mailto:'].includes(url.protocol)) { throw new Error('Use an http, https or mailto URL.'); }
					ppt.setShapeHyperlink(shape, url.href);
				}
			}
		} catch (error) { this.deck = await ppt.loadPresentation(before); throw error; }
		this.past.push(before);
		while (this.past.length > 50 || (this.past.length > 1 && this.past.reduce((n, b) => n + b.length, 0) > 128 * 1024 * 1024)) { this.past.shift(); }
		this.future = []; this.revision++;
		return this.state();
	}
	async export() { return ppt.savePresentation(this.deck); }
	render(page: number) {
		const slide = ppt.getSlides(this.deck)[page];
		if (!slide) { throw new Error('Slide does not exist.'); }
		return { svg: renderSlideToSvg(this.deck, slide, { textLayout: 'svg' }) };
	}
	dispose() { this.past = []; this.future = []; }
}
