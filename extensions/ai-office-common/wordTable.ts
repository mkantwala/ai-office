/*---------------------------------------------------------------------------------------------
 *  Copyright (c) Microsoft Corporation. All rights reserved.
 *  Licensed under the MIT License. See License.txt in the project root for license information.
 *--------------------------------------------------------------------------------------------*/

import type { TableGeometry, TablePreview } from './types.js';

type Point = { page: number; x: number; y: number };
type Edge = { kind: 'column' | 'row' | 'corner'; index: number; distance: number };
interface Gesture {
	pointer: number; page: HTMLElement; from: Point; start: Point; command: string; reach: number; revision: number;
	table: TableGeometry; latest?: Record<string, unknown>; pending?: number; moved: boolean;
}

/** Page-local affordances; preview and commit use the same native table calculations. */
export class WordTable {
	private readonly geometry = new Map<number, { revision: number; tables: TableGeometry[] }>();
	private readonly controller = new AbortController();
	private readonly menu = document.createElement('div');
	private readonly handles = document.createElement('div');
	private readonly guide = document.createElement('div');
	private revision = -1;
	private gesture?: Gesture;
	private hoverKey = '';
	private request = 0;
	private animation = 0;
	private ghost?: HTMLCanvasElement;
	constructor(private readonly pages: HTMLElement,
		private readonly point: (event: PointerEvent, page: HTMLElement) => Point,
		private readonly place: (node: HTMLElement, page: number, x: number, y: number, w: number, h: number) => void,
		private readonly execute: (command: string, args?: Record<string, unknown>) => void,
		private readonly ask: (label: string, initial?: string) => Promise<string | undefined>,
		private readonly focus: () => void,
		private readonly scale: (page: HTMLElement) => number,
		private readonly requestPreview: (request: number, args: Record<string, unknown>) => void,
		private readonly selected: (at: Point) => boolean) {
		const signal = this.controller.signal;
		this.menu.className = 'word-table-context'; this.menu.popover = 'manual'; this.menu.setAttribute('role', 'menu'); this.menu.setAttribute('aria-label', 'Table options'); document.body.append(this.menu);
		this.handles.className = 'word-table-handles'; this.guide.className = 'word-table-guide'; this.guide.setAttribute('aria-hidden', 'true');
		pages.addEventListener('pointermove', event => this.move(event), { signal, capture: true });
		pages.addEventListener('pointerdown', event => this.down(event), { signal, capture: true });
		pages.addEventListener('pointerup', event => this.up(event), { signal, capture: true });
		pages.addEventListener('pointercancel', () => this.cancel(), { signal });
		pages.addEventListener('lostpointercapture', () => this.cancel(), { signal });
		pages.addEventListener('pointerleave', () => { if (!this.gesture && !this.menu.matches(':popover-open')) { this.clearHover(); } }, { signal });
		pages.addEventListener('contextmenu', event => this.context(event), { signal });
		pages.addEventListener('dblclick', event => {
			const target = this.target(event); const edge = target && this.edge(target);
			if (edge?.kind === 'column' && edge.index > 0) {
				event.preventDefault(); event.stopPropagation(); this.execute('table.fitColumn', { ...target!.at, reach: 4 / this.scale(target!.page) }); this.focus();
			}
		}, { signal, capture: true });
		document.addEventListener('pointerdown', event => { if (!this.menu.contains(event.target as Node)) { this.close(); } }, { signal });
		document.addEventListener('keydown', event => {
			if (event.key === 'Escape' && (this.gesture || this.menu.matches(':popover-open'))) {
				event.preventDefault(); event.stopPropagation(); this.cancel(); this.close(); this.focus();
			}
		}, { signal, capture: true });
		pages.parentElement?.addEventListener('scroll', () => { if (!this.gesture) { this.close(); this.clearHover(); } }, { signal });
		window.addEventListener('resize', () => { this.cancel(); this.close(); }, { signal });
		window.addEventListener('blur', () => this.cancel(), { signal });
		this.menu.addEventListener('keydown', event => {
			if (!['ArrowDown', 'ArrowUp', 'Home', 'End'].includes(event.key)) { return; }
			event.preventDefault(); const items = [...this.menu.querySelectorAll<HTMLButtonElement>('button')].filter(item => !item.disabled && item.getClientRects().length > 0);
			const index = items.indexOf(document.activeElement as HTMLButtonElement);
			items[event.key === 'Home' ? 0 : event.key === 'End' ? items.length - 1 : (index + (event.key === 'ArrowDown' ? 1 : -1) + items.length) % items.length]?.focus();
		}, { signal });
	}
	update(revision: number) { if (revision !== this.revision) { this.revision = revision; this.geometry.clear(); this.close(); this.cancel(); } }
	forget(page: number) { this.geometry.delete(page); if (this.handles.parentElement === this.pages.children[page] && !this.gesture) { this.clearHover(); } }
	render(page: number, revision: number, tables: TableGeometry[]) { if (revision === this.revision) { this.geometry.set(page, { revision, tables }); } }
	private close() { if (this.menu.matches(':popover-open')) { this.menu.hidePopover(); } }
	private clearHover() { this.hoverKey = ''; this.handles.remove(); this.pages.querySelectorAll<HTMLElement>('.page.word').forEach(page => page.style.cursor = ''); }
	private cancel() {
		const gesture = this.gesture; this.gesture = undefined; cancelAnimationFrame(this.animation); this.animation = 0;
		if (gesture?.page.hasPointerCapture(gesture.pointer)) { gesture.page.releasePointerCapture(gesture.pointer); }
		this.guide.remove(); this.guide.replaceChildren(); this.ghost = undefined; this.clearHover(); this.pages.style.cursor = '';
	}
	private target(event: PointerEvent | MouseEvent) {
		const page = (event.target as HTMLElement).closest<HTMLElement>('.page.word'); if (!page) { return; }
		const at = this.point(event as PointerEvent, page); const geometry = this.geometry.get(at.page); if (geometry?.revision !== this.revision) { return; }
		const margin = Math.max(18, 22 / this.scale(page));
		const table = [...geometry.tables].reverse().find(t => t.rows.length && at.x >= t.cols[0] - margin && at.x <= t.cols[t.cols.length - 1] + margin && at.y >= t.rows[0].top - margin && at.y <= t.rows[t.rows.length - 1].bottom + margin);
		if (!table) { return; }
		const row = table.rows.find(r => at.y < r.bottom) || table.rows[table.rows.length - 1];
		const col = Math.max(0, Math.min(table.cols.length - 2, table.cols.filter(x => x <= at.x).length - 1));
		return { page, at, table, row, col };
	}
	private edge(target: NonNullable<ReturnType<WordTable['target']>>): Edge | undefined {
		const { at, table, page } = target; const reach = 4 / this.scale(page); let best: Edge | undefined;
		const offer = (edge: Edge) => { if (edge.distance <= reach && (!best || edge.distance < best.distance - .01)) { best = edge; } };
		const row = table.rows.find(r => at.y >= r.top && at.y <= r.bottom);
		if (row) { table.cols.forEach((x, index) => { if ((!row.edges || row.edges.includes(index)) && (index > 0 || at.x <= x + 1)) { offer({ kind: 'column', index, distance: Math.abs(x - at.x) }); } }); }
		if (at.x >= table.cols[0] && at.x <= table.cols.at(-1)!) { table.rows.forEach((r, index) => offer({ kind: 'row', index, distance: Math.abs(r.bottom - at.y) })); }
		return best;
	}
	private move(event: PointerEvent) {
		if (this.gesture) {
			event.preventDefault(); event.stopPropagation(); const g = this.gesture;
			const to = this.point(event, g.page); g.moved ||= Math.hypot(to.x - g.start.x, to.y - g.start.y) * this.scale(g.page) > 3;
			g.latest = { command: g.command, revision: g.revision, ...g.from, toX: g.from.x + to.x - g.start.x, toY: g.from.y + to.y - g.start.y, reach: g.reach, shift: event.shiftKey, ctrl: event.ctrlKey || event.metaKey };
			if (g.moved && !this.animation) { this.animation = requestAnimationFrame(() => { this.animation = 0; this.sendPreview(); }); } return;
		}
		if ((event.target as HTMLElement).closest('.word-table-handles') || this.menu.matches(':popover-open')) { return; }
		const target = this.target(event); if (!target) { this.clearHover(); return; }
		const { table, page, at, row, col } = target; const edge = this.edge(target);
		page.style.cursor = edge?.kind === 'column' ? 'col-resize' : edge?.kind === 'row' ? 'row-resize' : '';
		const left = table.cols[0], right = table.cols.at(-1)!, top = table.rows[0].top, bottom = table.rows.at(-1)!.bottom;
		const insertRow = at.x < left && [row.top, row.bottom].find(y => Math.abs(y - at.y) < 4);
		const insertCol = at.y < top && table.cols.find(x => Math.abs(x - at.x) < 4);
		const rowOptions = at.x < left + 8 / this.scale(page), columnOptions = at.y < top + 4 / this.scale(page);
		const key = JSON.stringify([at.page, left, top, row.row, col, insertRow, insertCol, rowOptions, columnOptions]); if (key === this.hoverKey) { return; } this.hoverKey = key;
		this.handles.replaceChildren(); page.append(this.handles);
		const icon = (path: string) => { const svg = document.createElementNS('http://www.w3.org/2000/svg', 'svg'); svg.setAttribute('viewBox', '0 0 24 24'); svg.setAttribute('aria-hidden', 'true'); const p = document.createElementNS(svg.namespaceURI, 'path'); p.setAttribute('d', path); svg.append(p); return svg; };
		const handle = (label: string, kind: string, path: string, x: number, y: number, command?: string, nativeX = x, nativeY = y) => {
			const button = document.createElement('button'); button.type = 'button'; button.title = label; button.setAttribute('aria-label', label); button.append(icon(path));
			button.dataset.handle = kind; button.dataset.command = command || ''; button.dataset.x = String(nativeX); button.dataset.y = String(nativeY);
			this.place(button, at.page, x, y, 0, 0); button.style.width = button.style.height = ''; this.handles.append(button);
			if (!command) { button.addEventListener('click', e => { e.stopPropagation(); this.open(e.clientX, e.clientY, target, kind === 'row' ? 'row' : kind === 'column' ? 'column' : 'cell'); }); }
			if (command === 'table.insertAt') { button.addEventListener('click', e => { e.stopPropagation(); this.execute(command, { page: at.page, x: nativeX, y: nativeY, reach: 4 / this.scale(page) }); this.focus(); }); }
		};
		const dots = 'M5 12h.01M12 12h.01M19 12h.01';
		if (insertRow !== false && insertRow !== undefined) { handle('Insert row here', 'insert', 'M12 5v14M5 12h14', left - 10, insertRow, 'table.insertAt'); }
		else if (rowOptions) { handle('Row options', 'row', 'M9 5h.01M15 5h.01M9 12h.01M15 12h.01M9 19h.01M15 19h.01', left - 10, (row.top + row.bottom) / 2); }
		if (insertCol !== false && insertCol !== undefined) { handle('Insert column here', 'insert', 'M12 5v14M5 12h14', insertCol, top - 10, 'table.insertAt'); }
		else if (columnOptions) { handle('Column options', 'column', dots, (table.cols[col] + table.cols[col + 1]) / 2, top - 6); }
		handle('Table options', 'options', dots, right - 6, top - 6);
		if (table.rows[0].row === 0 && table.movable !== false) { handle('Drag to move table', 'move', 'M12 3v18M3 12h18M8 7l4-4 4 4M8 17l4 4 4-4M7 8l-4 4 4 4M17 8l4 4-4 4', left - 10, top - 10, 'table.moveDrag', left - 6, top - 6); }
		if (table.endsHere) { handle('Drag to resize table', 'resize', 'M8 16 18 6M10 6h8v8', right + 5, bottom + 5, 'table.resizeDrag', right + 4, bottom + 4); }
	}
	private down(event: PointerEvent) {
		if (event.button !== 0) { return; }
		const button = (event.target as HTMLElement).closest<HTMLButtonElement>('.word-table-handles button');
		if (button && (!button.dataset.command || button.dataset.command === 'table.insertAt')) { event.preventDefault(); event.stopPropagation(); return; }
		const target = this.target(event); if (!target) { return; }
		const edge = this.edge(target); let command = button?.dataset.command; let from = { ...target.at };
		if (button && command) { from = { page: from.page, x: Number(button.dataset.x), y: Number(button.dataset.y) }; }
		else if (edge) { command = 'table.resizeDrag'; if (edge.kind === 'column') { from.x = target.table.cols[edge.index]; } else { from.y = target.table.rows[edge.index].bottom; } }
		if (!command) { return; }
		event.preventDefault(); event.stopPropagation(); this.close(); this.focus();
		this.gesture = { page: target.page, pointer: event.pointerId, from, start: target.at, command, reach: 4 / this.scale(target.page), revision: this.revision, table: target.table, moved: false };
		target.page.setPointerCapture(event.pointerId); target.page.style.cursor = command === 'table.moveDrag' ? 'grabbing' : button ? 'nwse-resize' : edge?.kind === 'column' ? 'col-resize' : 'row-resize';
		this.pages.style.cursor = target.page.style.cursor;
		if (command === 'table.moveDrag') { this.makeGhost(target.page, target.table); }
	}
	private makeGhost(page: HTMLElement, table: TableGeometry) {
		const source = page.querySelector('canvas'); if (!source || source.width <= 1) { return; }
		const scale = source.width / (page.getBoundingClientRect().width / this.scale(page));
		const x = Math.max(0, table.cols[0] * scale), y = Math.max(0, table.rows[0].top * scale);
		const width = Math.min(source.width - x, (table.cols.at(-1)! - table.cols[0]) * scale), height = Math.min(source.height - y, (table.rows.at(-1)!.bottom - table.rows[0].top) * scale);
		if (width <= 0 || height <= 0) { return; }
		this.ghost = document.createElement('canvas'); this.ghost.width = Math.ceil(width); this.ghost.height = Math.ceil(height);
		this.ghost.getContext('2d')?.drawImage(source, x, y, width, height, 0, 0, this.ghost.width, this.ghost.height);
	}
	private sendPreview() {
		const g = this.gesture; if (!g?.latest || g.pending !== undefined) { return; }
		const args = g.latest; g.latest = undefined; g.pending = ++this.request; this.requestPreview(g.pending, args);
	}
	preview(request: number, revision: number, result?: TablePreview) {
		const g = this.gesture; if (!g || g.pending !== request || g.revision !== revision || revision !== this.revision) { return; } g.pending = undefined;
		if (result) {
			this.guide.replaceChildren(); this.guide.dataset.kind = result.kind; g.page.append(this.guide);
			for (const rect of result.rects) {
				const outline = document.createElement('div'); outline.className = 'word-table-preview'; this.place(outline, g.from.page, rect.x, rect.y, rect.width, rect.height);
				if (this.ghost && result.kind === 'move') { outline.append(this.ghost); } this.guide.append(outline);
			}
		}
		this.sendPreview();
	}
	private up(event: PointerEvent) {
		const g = this.gesture; if (!g || event.pointerId !== g.pointer) { return; }
		this.move(event); event.preventDefault(); event.stopPropagation(); const to = this.point(event, g.page); const moved = g.moved; this.cancel();
		if (moved) { this.execute(g.command, { ...g.from, toX: g.from.x + to.x - g.start.x, toY: g.from.y + to.y - g.start.y, reach: g.reach, shift: event.shiftKey, ctrl: event.ctrlKey || event.metaKey }); }
		else if (g.command === 'table.moveDrag') { this.execute('table.selectAt', { ...g.from, reach: g.reach }); }
		this.focus();
	}
	private context(event: MouseEvent) { const target = this.target(event); if (!target) { return; } event.preventDefault(); event.stopPropagation(); this.open(event.clientX, event.clientY, target); }
	private open(x: number, y: number, target: NonNullable<ReturnType<WordTable['target']>>, scope: 'row' | 'column' | 'cell' = 'cell') {
		this.close(); this.menu.replaceChildren();
		const at = { page: target.at.page, x: (target.table.cols[target.col] + target.table.cols[target.col + 1]) / 2, y: (target.row.top + target.row.bottom) / 2 };
		if (scope !== 'cell' || !this.selected(target.at)) { this.execute('caret.set', at); }
		if (scope !== 'cell') { this.execute(scope === 'row' ? 'table.selectRow' : 'table.selectColumn'); }
		const revision = this.revision;
		const run = (command: string, params: Record<string, unknown> = {}) => { this.close(); if (revision === this.revision) { this.execute(command, params); } this.focus(); };
		const item = (label: string, command: string, params?: Record<string, unknown>, action?: () => void) => {
			const button = document.createElement('button'); button.type = 'button'; button.textContent = label; button.setAttribute('role', 'menuitem');
			const svg = document.createElementNS('http://www.w3.org/2000/svg', 'svg'); svg.setAttribute('viewBox', '0 0 24 24'); svg.setAttribute('aria-hidden', 'true');
			const path = document.createElementNS(svg.namespaceURI, 'path');
			const column = command.includes('Column');
			path.setAttribute('d', command.startsWith('table.insert') ? (column ? 'M11 4v16m5-16v16M2 12h6M5 9v6' : 'M4 11h16M4 16h16M2 5h6M5 2v6') : command.startsWith('table.delete') ? (column ? 'M11 4v16m5-16v16M2 12h6' : 'M4 11h16M4 16h16M2 5h6') : command === 'table.sort' ? (params?.descending ? 'M12 3v18m-6-6 6 6 6-6' : 'M12 21V3m-6 6 6-6 6 6') : command === 'table.cellAlign' ? (params?.value === 'topCenter' ? 'M4 5h16M7 10h10M4 15h16M7 20h10' : params?.value === 'topRight' ? 'M4 5h16M10 10h10M4 15h16M10 20h10' : 'M4 5h16M4 10h10M4 15h16M4 20h10') : command === 'table.columnWidth' ? 'M4 4v16m16-16v16M7 12h10m-7-3-3 3 3 3m4-6 3 3-3 3' : 'M4 12h16m-4-4 4 4-4 4M8 8l-4 4 4 4');
			svg.append(path); button.prepend(svg);
			button.addEventListener('click', action || (() => run(command, params))); this.menu.append(button); return button;
		};
		const separator = () => { const line = document.createElement('hr'); line.setAttribute('role', 'separator'); this.menu.append(line); };
		item('Insert row above', 'table.insertRowAbove'); item('Insert row below', 'table.insertRowBelow'); item('Insert column left', 'table.insertColumnLeft'); item('Insert column right', 'table.insertColumnRight');
		const moves = document.createElement('details'); const summary = document.createElement('summary'); summary.textContent = 'Move row or column'; moves.append(summary); this.menu.append(moves);
		for (const [label, axis, forward] of [['Move row up', 'row', false], ['Move row down', 'row', true], ['Move column left', 'column', false], ['Move column right', 'column', true]] as const) { moves.append(item(label, 'table.moveAxis', { axis, forward })); }
		// A prompt uses the same validated native width command; cancellation is a no-op.
		item('Set column width…', 'table.columnWidth', undefined, async () => { this.close(); let value = await this.ask('Column width in points (72 pt = 1 inch)', String(Math.round(target.table.cols[target.col + 1] - target.table.cols[target.col]))); while (value !== undefined) { const n = Number(value); if (value.trim() && Number.isFinite(n) && n >= 6 && n <= 1584) { run('table.columnWidth', { width: n }); return; } value = await this.ask('Enter a width from 6 to 1584 points', value); } });
		item('Set row height…', 'table.rowHeight', undefined, async () => {
			this.close(); let value = await this.ask('Minimum row height in points (72 pt = 1 inch)', String(Math.round(target.row.bottom - target.row.top)));
			while (value !== undefined) { const height = Number(value); if (value.trim() && Number.isFinite(height) && height >= 1 && height <= 1584) { run('table.rowHeight', { height, rule: 'atLeast' }); return; } value = await this.ask('Enter a height from 1 to 1584 points', value); }
			this.focus();
		});
		const layout = document.createElement('details'); const layoutTitle = document.createElement('summary'); layoutTitle.textContent = 'Size and layout'; layout.append(layoutTitle); this.menu.append(layout);
		for (const [label, command, params] of [
			['AutoFit to contents', 'table.autofit', { mode: 'contents' }], ['Fit to page width', 'table.autofit', { mode: 'window' }],
			['Fixed column widths', 'table.autofit', { mode: 'fixed' }], ['Distribute columns evenly', 'table.distributeColumns', {}],
			['Distribute rows evenly', 'table.distributeRows', {}], ['Repeat header row', 'table.repeatHeader', { value: true }],
			['Stop repeating header row', 'table.repeatHeader', { value: false }]
		] as [string, string, Record<string, unknown>][]) { layout.append(item(label, command, params)); }
		const cells = document.createElement('details'); const cellsTitle = document.createElement('summary'); cellsTitle.textContent = 'Cells and borders'; cells.append(cellsTitle); this.menu.append(cells);
		cells.append(item('Select cell', 'table.selectCell'), item('Select table', 'table.selectTable'));
		cells.append(item('Merge selected cells', 'table.merge'), item('Split cell into two', 'table.split', { columns: 2 }));
		for (const [label, kind] of [['All borders', 'all'], ['Outside borders', 'outside'], ['Inside borders', 'inside'], ['No borders', 'none']]) { cells.append(item(label, 'table.borders', { kind })); }
		separator(); item('Delete row', 'table.deleteRow'); item('Delete column', 'table.deleteColumn'); item('Delete table', 'table.deleteTable').classList.add('destructive');
		separator(); item('Sort column ascending', 'table.sort', { column: target.col, descending: false }); item('Sort column descending', 'table.sort', { column: target.col, descending: true });
		separator(); for (const [label, value] of [['Align left', 'topLeft'], ['Align center', 'topCenter'], ['Align right', 'topRight']]) { item(label, 'table.cellAlign', { value }); }
		separator(); const palette = document.createElement('div'); palette.className = 'word-table-palette'; palette.setAttribute('aria-label', 'Cell background color'); this.menu.append(palette);
		for (const [label, color] of [['No fill', null], ['Gray', '262626'], ['Red', '4C1613'], ['Orange', '54290F'], ['Yellow', '4D350B'], ['Green', '183A21'], ['Blue', '0B3155'], ['Purple', '302052']]) {
			const button = item(label!, 'table.shading', { color }); button.title = label!; button.setAttribute('aria-label', label!); button.textContent = ''; if (!color) { button.classList.add('no-fill'); } if (color) { button.style.backgroundColor = '#' + color; } palette.append(button);
		}
		this.menu.showPopover(); this.menu.style.left = Math.max(8, Math.min(x, innerWidth - this.menu.offsetWidth - 8)) + 'px'; this.menu.style.top = Math.max(8, Math.min(y, innerHeight - this.menu.offsetHeight - 8)) + 'px'; this.menu.querySelector<HTMLButtonElement>('button')?.focus();
	}
	dispose() { this.controller.abort(); this.cancel(); this.close(); this.menu.remove(); this.handles.remove(); this.geometry.clear(); }
}
