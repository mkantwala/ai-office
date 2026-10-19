/*---------------------------------------------------------------------------------------------
 *  Copyright (c) Microsoft Corporation. All rights reserved.
 *  Licensed under the MIT License. See License.txt in the project root for license information.
 *--------------------------------------------------------------------------------------------*/

import type { TableGeometry } from './types.js';

/** Page-local affordances over the native canvas. All mutations use native commands/history. */
export class WordTable {
	private readonly geometry = new Map<number, { revision: number; tables: TableGeometry[] }>();
	private readonly controller = new AbortController();
	private readonly menu = document.createElement('div');
	private readonly handles = document.createElement('div');
	private revision = -1;
	private gesture?: { pointer: number; page: HTMLElement; from: { page: number; x: number; y: number }; command: string };
	constructor(private readonly pages: HTMLElement,
		private readonly point: (event: PointerEvent, page: HTMLElement) => { page: number; x: number; y: number },
		private readonly place: (node: HTMLElement, page: number, x: number, y: number, w: number, h: number) => void,
		private readonly execute: (command: string, args?: Record<string, unknown>) => void,
		private readonly ask: (label: string, initial?: string) => Promise<string | undefined>,
		private readonly focus: () => void) {
		const signal = this.controller.signal;
		this.menu.className = 'word-table-context'; this.menu.popover = 'manual'; this.menu.setAttribute('role', 'menu'); this.menu.setAttribute('aria-label', 'Table options'); document.body.append(this.menu);
		this.handles.className = 'word-table-handles';
		pages.addEventListener('pointermove', event => this.move(event), { signal });
		pages.addEventListener('pointerdown', event => this.down(event), { signal, capture: true });
		pages.addEventListener('pointerup', event => this.up(event), { signal, capture: true });
		pages.addEventListener('pointercancel', () => { this.gesture = undefined; this.handles.remove(); }, { signal });
		pages.addEventListener('pointerleave', () => { if (!this.gesture) { this.handles.remove(); pages.querySelectorAll<HTMLElement>('.page.word').forEach(page => page.style.cursor = ''); } }, { signal });
		pages.addEventListener('contextmenu', event => this.context(event), { signal });
		document.addEventListener('pointerdown', event => { if (!this.menu.contains(event.target as Node)) { this.close(); } }, { signal });
		document.addEventListener('keydown', event => {
			if (event.key === 'Escape' && (this.gesture || this.menu.matches(':popover-open'))) { event.preventDefault(); event.stopPropagation(); this.gesture = undefined; this.close(); this.handles.remove(); this.focus(); }
		}, { signal });
		pages.parentElement?.addEventListener('scroll', () => { this.close(); this.handles.remove(); }, { signal });
		window.addEventListener('resize', () => { this.close(); this.handles.remove(); }, { signal });
		this.menu.addEventListener('keydown', event => {
			if (!['ArrowDown', 'ArrowUp', 'Home', 'End'].includes(event.key)) { return; }
			event.preventDefault(); const items = [...this.menu.querySelectorAll<HTMLButtonElement>('button')].filter(item => item.getClientRects().length > 0);
			const index = items.indexOf(document.activeElement as HTMLButtonElement);
			items[event.key === 'Home' ? 0 : event.key === 'End' ? items.length - 1 : (index + (event.key === 'ArrowDown' ? 1 : -1) + items.length) % items.length]?.focus();
		}, { signal });
	}
	update(revision: number) { if (revision !== this.revision) { this.revision = revision; this.geometry.clear(); this.close(); this.handles.remove(); this.gesture = undefined; } }
	forget(page: number) { this.geometry.delete(page); if (this.handles.parentElement === this.pages.children[page]) { this.handles.remove(); } }
	render(page: number, revision: number, tables: TableGeometry[]) { if (revision === this.revision) { this.geometry.set(page, { revision, tables }); } }
	private close() { if (this.menu.matches(':popover-open')) { this.menu.hidePopover(); } }
	private target(event: PointerEvent | MouseEvent) {
		const page = (event.target as HTMLElement).closest<HTMLElement>('.page.word'); if (!page) { return; }
		const at = this.point(event as PointerEvent, page); const geometry = this.geometry.get(at.page); if (geometry?.revision !== this.revision) { return; }
		// Last geometry is the innermost table when tables are nested.
		const table = [...geometry.tables].reverse().find(t => t.rows.length && at.x >= t.cols[0] - 18 && at.x <= t.cols[t.cols.length - 1] + 6 && at.y >= t.rows[0].top - 14 && at.y <= t.rows[t.rows.length - 1].bottom + 6);
		if (!table) { return; }
		const row = table.rows.find(r => at.y <= r.bottom) || table.rows[table.rows.length - 1];
		const col = Math.max(0, Math.min(table.cols.length - 2, table.cols.filter(x => x <= at.x).length - 1));
		return { page, at, table, row, col };
	}
	private move(event: PointerEvent) {
		if (this.gesture) { return; }
		if ((event.target as HTMLElement).closest('.word-table-handles')) { return; }
		const target = this.target(event); this.handles.remove(); const pageUnderPointer = (event.target as HTMLElement).closest<HTMLElement>('.page.word'); if (pageUnderPointer) { pageUnderPointer.style.cursor = ''; } if (!target) { return; }
		const { table, page, at, row, col } = target;
		this.handles.replaceChildren(); page.append(this.handles);
		const handle = (label: string, text: string, x: number, y: number, command?: string) => {
			const button = document.createElement('button'); button.type = 'button'; button.title = label; button.setAttribute('aria-label', label); button.textContent = text;
			if (command === 'table.resizeDrag') {
				const svg = document.createElementNS('http://www.w3.org/2000/svg', 'svg'); svg.setAttribute('viewBox', '0 0 24 24'); svg.setAttribute('aria-hidden', 'true');
				const path = document.createElementNS(svg.namespaceURI, 'path'); path.setAttribute('d', 'M6 18 18 6M10 6h8v8M6 10v8h8'); svg.append(path); button.replaceChildren(svg);
			}
			button.dataset.handle = label === 'Row options' ? 'row' : label === 'Column options' ? 'column' : label === 'Table options' ? 'options' : command === 'table.moveDrag' ? 'move' : 'resize';
			button.dataset.command = command || ''; button.dataset.x = String(x); button.dataset.y = String(y);
			this.place(button, at.page, x, y, 0, 0); this.handles.append(button);
			button.style.width = button.style.height = ''; // fixed CSS hit targets, independent of zoom
			if (!command) { button.addEventListener('click', e => { e.stopPropagation(); this.open(e.clientX, e.clientY, target, label === 'Row options' ? 'row' : label === 'Column options' ? 'column' : 'cell'); }); }
		};
		handle('Row options', '⠿', table.cols[0] - 10, (row.top + row.bottom) / 2);
		handle('Column options', '⋯', (table.cols[col] + table.cols[col + 1]) / 2, table.rows[0].top - 8);
		handle('Table options', '⋯', table.cols[table.cols.length - 1] - 12, table.rows[0].top + 12);
		if (table.rows[0].row === 0) { handle('Drag to move table', '⠿', table.cols[0] - 6, table.rows[0].top - 6, 'table.moveDrag'); }
		if (table.endsHere) { handle('Drag to resize table', '↘', table.cols[table.cols.length - 1] + 3, table.rows[table.rows.length - 1].bottom + 3, 'table.resizeDrag'); }
		const vertical = table.cols.find(x => Math.abs(x - at.x) < 3);
		const horizontal = table.rows.find(r => Math.abs(r.bottom - at.y) < 3);
		page.style.cursor = vertical !== undefined ? 'col-resize' : horizontal ? 'row-resize' : '';
	}
	private down(event: PointerEvent) {
		if (event.button !== 0) { return; }
		const button = (event.target as HTMLElement).closest<HTMLButtonElement>('.word-table-handles button');
		if (button && !button.dataset.command) { event.preventDefault(); event.stopPropagation(); return; }
		const target = this.target(event); if (!target) { return; }
		let from = target.at; let command = button?.dataset.command;
		if (button && command) { from = { page: from.page, x: Number(button.dataset.x), y: Number(button.dataset.y) }; }
		if (!command && (target.table.cols.some(x => Math.abs(x - from.x) < 3) || target.table.rows.some(r => Math.abs(r.bottom - from.y) < 3))) { command = 'table.resizeDrag'; }
		if (!command) { return; }
		event.preventDefault(); event.stopPropagation(); this.close();
		this.gesture = { page: target.page, pointer: event.pointerId, from, command }; target.page.setPointerCapture(event.pointerId);
	}
	private up(event: PointerEvent) {
		const gesture = this.gesture; if (!gesture || event.pointerId !== gesture.pointer) { return; }
		event.preventDefault(); event.stopPropagation(); const to = this.point(event, gesture.page); this.gesture = undefined;
		if (gesture.page.hasPointerCapture(event.pointerId)) { gesture.page.releasePointerCapture(event.pointerId); }
		if (Math.abs(to.x - gesture.from.x) + Math.abs(to.y - gesture.from.y) > 1) {
			this.execute(gesture.command, { ...gesture.from, toX: to.x, toY: to.y, reach: 3, shift: event.shiftKey, ctrl: event.ctrlKey });
		}
		this.focus();
	}
	private context(event: MouseEvent) { const target = this.target(event); if (!target) { return; } event.preventDefault(); event.stopPropagation(); this.open(event.clientX, event.clientY, target); }
	private open(x: number, y: number, target: NonNullable<ReturnType<WordTable['target']>>, scope: 'row' | 'column' | 'cell' = 'cell') {
		this.close(); this.menu.replaceChildren();
		const at = { page: target.at.page, x: (target.table.cols[target.col] + target.table.cols[target.col + 1]) / 2, y: (target.row.top + target.row.bottom) / 2 };
		this.execute('caret.set', at);
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
		separator(); item('Delete row', 'table.deleteRow'); item('Delete column', 'table.deleteColumn'); item('Delete table', 'table.deleteTable').classList.add('destructive');
		separator(); item('Sort column ascending', 'table.sort', { column: target.col, descending: false }); item('Sort column descending', 'table.sort', { column: target.col, descending: true });
		separator(); for (const [label, value] of [['Align left', 'topLeft'], ['Align center', 'topCenter'], ['Align right', 'topRight']]) { item(label, 'table.cellAlign', { value }); }
		separator(); const palette = document.createElement('div'); palette.className = 'word-table-palette'; palette.setAttribute('aria-label', 'Cell background color'); this.menu.append(palette);
		for (const [label, color] of [['No fill', null], ['Gray', '262626'], ['Red', '4C1613'], ['Orange', '54290F'], ['Yellow', '4D350B'], ['Green', '183A21'], ['Blue', '0B3155'], ['Purple', '302052']]) {
			const button = item(label!, 'table.shading', { color }); button.title = label!; button.setAttribute('aria-label', label!); button.textContent = ''; if (!color) { button.classList.add('no-fill'); } if (color) { button.style.backgroundColor = '#' + color; } palette.append(button);
		}
		this.menu.showPopover(); this.menu.style.left = Math.max(8, Math.min(x, innerWidth - this.menu.offsetWidth - 8)) + 'px'; this.menu.style.top = Math.max(8, Math.min(y, innerHeight - this.menu.offsetHeight - 8)) + 'px'; this.menu.querySelector<HTMLButtonElement>('button')?.focus();
	}
	dispose() { this.controller.abort(); this.close(); this.menu.remove(); this.handles.remove(); this.geometry.clear(); }
}
