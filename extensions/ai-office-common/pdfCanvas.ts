/*---------------------------------------------------------------------------------------------
 *  Copyright (c) Microsoft Corporation. All rights reserved.
 *  Licensed under the MIT License. See License.txt in the project root for license information.
 *--------------------------------------------------------------------------------------------*/
import type { Block, PdfObject, State } from './types.js';

type Mode = 'select' | 'text' | 'image' | 'link' | 'note' | 'highlight' | 'rectangle' | 'oval' | 'arrow' | 'form' | 'redact' | 'crop';
type Grip = 'move' | 'left' | 'right' | 'corner';
interface Snapshot { state: State; dimensions: number[][] }
interface Host {
	snapshot(): Snapshot;
	execute(command: string, args: Record<string, unknown>, revision: number): void;
	form(command?: string, initial?: Record<string, unknown>): void;
	select(block?: Block, object?: PdfObject): void;
	page(index: number): void;
	message(text: string): void;
	organize(): void;
}
interface Gesture { page: number; revision: number; start: number[]; end: number[]; grip: Grip; objects: PdfObject[]; pointer: number; target: HTMLElement; moved: boolean; mode: Mode; clicked?: PdfObject; extend: boolean; aspect: boolean; previews: HTMLElement[] }
const key = (object: PdfObject) => object.kind + ':' + object.index;
const validRect = (rect: number[]) => rect.length === 4 && rect.every(Number.isFinite) && rect[2] > rect[0] && rect[3] > rect[1];

/** Canvas gestures use the same native commands and revision guards as agent edits. */
export class PdfCanvas {
	private readonly events = new AbortController();
	private chosen = new Set<string>();
	private gesture?: Gesture;
	private editor?: { node: HTMLTextAreaElement; object: PdfObject; page: number; revision: number; original: string };
	private menu?: HTMLElement;
	private palette?: HTMLDialogElement;
	private mode: Mode = 'select';
	private revision = -1;
	private page = -1;
	private waiting = false;
	private lastClick?: { key: string; time: number };
	constructor(private readonly pages: HTMLElement, private readonly host: Host) {
		const signal = this.events.signal;
		pages.addEventListener('pointerdown', event => this.down(event), { signal, capture: true });
		pages.addEventListener('pointermove', event => this.move(event), { signal });
		pages.addEventListener('pointerup', event => this.up(event), { signal });
		pages.addEventListener('pointercancel', () => this.cancel(), { signal });
		pages.addEventListener('lostpointercapture', () => { if (this.gesture) { this.cancel(); } }, { signal });
		pages.addEventListener('contextmenu', event => this.context(event), { signal });
		pages.addEventListener('keydown', event => this.keydown(event), { signal });
		window.addEventListener('blur', () => this.cancel(), { signal });
	}
	dispose() { this.events.abort(); this.cancel(); this.menu?.remove(); this.palette?.remove(); this.editor?.node.remove(); }
	update() {
		const { state } = this.host.snapshot();
		if (state.revision !== this.revision || state.page !== this.page) {
			this.cancel(); this.chosen.clear(); this.menu?.remove(); this.host.select();
			if (this.editor) { if (this.editor.node.value !== this.editor.original) { this.host.message('The PDF changed while you were typing. Copy your draft before reopening this text; Escape discards it.'); } else { this.editor.node.remove(); this.editor = undefined; } }
		}
		this.revision = state.revision; this.page = state.page || 0; this.waiting = false; this.render();
	}
	failed() { this.waiting = false; this.cancel(); this.render(); }
	private objects() { return (this.host.snapshot().state.pdfObjects || []).filter(object => validRect(object.rect)); }
	private selected() { return this.objects().filter(object => this.chosen.has(key(object))); }
	private at(event: MouseEvent, page: HTMLElement) {
		const bounds = page.getBoundingClientRect(), index = Number(page.dataset.page), size = this.host.snapshot().dimensions[index];
		return [(event.clientX - bounds.left) * size[0] / bounds.width, (event.clientY - bounds.top) * size[1] / bounds.height];
	}
	private place(node: HTMLElement, page: number, rect: number[]) {
		const target = this.pages.children[page] as HTMLElement, size = this.host.snapshot().dimensions[page]; if (!target || !size) { return; }
		node.style.left = rect[0] / size[0] * 100 + '%'; node.style.top = rect[1] / size[1] * 100 + '%';
		node.style.width = (rect[2] - rect[0]) / size[0] * 100 + '%'; node.style.height = (rect[3] - rect[1]) / size[1] * 100 + '%'; target.append(node);
	}
	private choose(object: PdfObject, extend = false) {
		if (!extend) { this.chosen.clear(); } if (extend && this.chosen.has(key(object))) { this.chosen.delete(key(object)); } else { this.chosen.add(key(object)); }
		const only = this.selected(); const one = only.length === 1 ? only[0] : undefined;
		const block = one?.kind === 'text' ? this.host.snapshot().state.blocks?.find(block => block.id === one.index) : one?.type === 'text' ? { id: one.index, text: one.text || '', rect: one.rect } : undefined;
		this.host.select(block, one); this.render();
	}
	private render() {
		this.pages.querySelectorAll('.pdf-object,.pdf-object-actions').forEach(node => node.remove());
		const { state } = this.host.snapshot(); if (state.page === undefined) { return; }
		for (const object of this.objects().sort((a, b) => Number(a.type === 'text') - Number(b.type === 'text'))) {
			const node = document.createElement('div'); node.className = 'pdf-object'; node.tabIndex = 0; node.role = 'button'; node.dataset.object = key(object); node.dataset.selected = String(this.chosen.has(key(object)));
			node.setAttribute('aria-label', `${object.type === 'text' ? 'Text' : 'Artwork'}: ${object.text?.slice(0, 60) || object.index}`); node.setAttribute('aria-pressed', String(this.chosen.has(key(object))));
			node.title = 'Drag to move · Double-click to edit · Shift-click to select together';
			this.place(node, state.page, object.rect);
			if (this.chosen.size === 1 && this.chosen.has(key(object))) {
				const grips: Grip[] = object.kind === 'text' ? (state.pdfRotation === 0 ? ['left', 'right'] : []) : object.kind === 'added' || object.fullRect ? ['corner'] : [];
				for (const grip of grips) { const handle = document.createElement('span'); handle.className = 'pdf-grip pdf-grip-' + grip; handle.dataset.grip = grip; handle.title = grip === 'corner' ? 'Resize' : 'Drag to rewrap text'; node.append(handle); }
			}
		}
		const selected = this.selected(); if (selected.length !== 1 || this.editor) { return; }
		const object = selected[0], actions = document.createElement('div'); actions.className = 'pdf-object-actions'; actions.setAttribute('role', 'toolbar'); actions.setAttribute('aria-label', 'Selected PDF object');
		const action = (label: string, run: () => void) => { const button = document.createElement('button'); button.type = 'button'; button.textContent = label; button.title = label; button.addEventListener('pointerdown', event => event.stopPropagation()); button.addEventListener('click', run); actions.append(button); };
		if (object.type === 'text') { action('Edit text', () => this.edit(object)); }
		action('Properties', () => this.properties(object));
		if (object.kind === 'image' || object.type === 'image') { action('Rotate', () => this.run(object.kind === 'image' ? 'image_edit' : 'content_update', object.kind === 'image' ? { image: object.index, action: 'rotate', quarters: 1 } : { index: object.index, rotate: 90 })); }
		action('Delete', () => this.remove(object));
		const surface = this.pages.children[state.page] as HTMLElement, size = this.host.snapshot().dimensions[state.page], scale = surface.clientWidth / size[0];
		const below = object.rect[1] * scale < 44, top = below ? object.rect[3] : object.rect[1];
		this.place(actions, state.page, [object.rect[0], top, object.rect[0] + 1, top + 1]);
		if (below) { actions.style.transform = 'translateY(8px)'; }
		actions.style.left = Math.max(4, Math.min(object.rect[0] * scale, surface.clientWidth - actions.offsetWidth - 4)) + 'px';
	}
	private down(event: PointerEvent) {
		const target = event.target as HTMLElement, page = target.closest<HTMLElement>('.page.pdf'); if (!page || target.closest('textarea,button,.pdf-object-actions,.pdf-context')) { return; }
		if (event.button !== 0 || this.waiting) { return; }
		const node = target.closest<HTMLElement>('.pdf-object'), object = this.objects().find(object => key(object) === node?.dataset.object);
		if (this.mode === 'select' && !object) { this.flush(); this.chosen.clear(); this.host.select(); this.render(); return; }
		event.preventDefault(); event.stopPropagation(); this.menu?.remove();
		if (this.editor) { this.flush(); return; }
		const index = Number(page.dataset.page); if (this.mode === 'select' && index !== this.host.snapshot().state.page) { this.host.page(index); return; }
		if (object && this.mode === 'select') { if (event.shiftKey || !this.chosen.has(key(object))) { this.choose(object, event.shiftKey); } }
		const point = this.at(event, page), grip = (target.dataset.grip || 'move') as Grip;
		this.gesture = { page: index, revision: this.host.snapshot().state.revision, start: point, end: point, grip, objects: this.mode === 'select' ? this.selected() : [], pointer: event.pointerId, target: page, moved: false, mode: this.mode, clicked: object, extend: event.shiftKey, aspect: !event.shiftKey, previews: [] };
		page.setPointerCapture(event.pointerId); page.focus({ preventScroll: true });
	}
	private preview(gesture: Gesture, object: PdfObject) {
		const [dx, dy] = gesture.end.map((v, i) => v - gesture.start[i]); const rect = [...object.rect];
		if (gesture.grip === 'left') { rect[0] = Math.min(rect[2] - 12, rect[0] + dx); }
		else if (gesture.grip === 'right') { rect[2] = Math.max(rect[0] + 12, rect[2] + dx); }
		else if (gesture.grip === 'corner') { rect[2] = Math.max(rect[0] + 12, rect[2] + dx); rect[3] = Math.max(rect[1] + 12, rect[3] + dy); if (gesture.aspect && object.type !== 'text') { const scale = Math.max((rect[2] - rect[0]) / (object.rect[2] - object.rect[0]), (rect[3] - rect[1]) / (object.rect[3] - object.rect[1])); rect[2] = rect[0] + (object.rect[2] - object.rect[0]) * scale; rect[3] = rect[1] + (object.rect[3] - object.rect[1]) * scale; } }
		else { rect[0] += dx; rect[2] += dx; rect[1] += dy; rect[3] += dy; } return rect;
	}
	private move(event: PointerEvent) {
		const gesture = this.gesture; if (!gesture || gesture.pointer !== event.pointerId) { return; }
		gesture.end = this.at(event, gesture.target); gesture.aspect = !event.shiftKey; const scale = gesture.target.clientWidth / this.host.snapshot().dimensions[gesture.page][0];
		if (Math.hypot(gesture.end[0] - gesture.start[0], gesture.end[1] - gesture.start[1]) * scale < 4 && !gesture.moved) { return; }
		gesture.moved = true; event.preventDefault();
		const rects = gesture.mode === 'select' ? gesture.objects.map(object => this.preview(gesture, object)) : [this.drawRect(gesture)];
		for (const [index, rect] of rects.entries()) {
			let preview = gesture.previews[index];
			if (!preview) {
				preview = document.createElement('div'); preview.className = 'pdf-drag-preview'; gesture.previews[index] = preview;
				const object = gesture.objects[index], source = gesture.target.querySelector('canvas');
				if (source && source.width > 1 && object) {
					const image = document.createElement('canvas'), size = this.host.snapshot().dimensions[gesture.page], box = object.rect;
					image.width = Math.max(1, Math.round((box[2] - box[0]) / size[0] * source.width)); image.height = Math.max(1, Math.round((box[3] - box[1]) / size[1] * source.height));
					image.getContext('2d')?.drawImage(source, box[0] / size[0] * source.width, box[1] / size[1] * source.height, image.width, image.height, 0, 0, image.width, image.height); preview.append(image);
				} else { preview.textContent = object?.text?.slice(0, 300) || ''; }
			}
			this.place(preview, gesture.page, rect);
		}
		gesture.target.style.cursor = gesture.grip === 'move' ? 'grabbing' : gesture.grip === 'corner' ? 'nwse-resize' : 'ew-resize';
	}
	private drawRect(gesture: Gesture) { const [x, y] = gesture.start, [ex, ey] = gesture.end; return [Math.min(x, ex), Math.min(y, ey), Math.max(x, ex), Math.max(y, ey)]; }
	private up(event: PointerEvent) {
		const gesture = this.gesture; if (!gesture || gesture.pointer !== event.pointerId) { return; }
		this.move(event); this.cancel();
		if (gesture.revision !== this.host.snapshot().state.revision) { this.host.message('The PDF changed during the gesture. Select the object again.'); return; }
		if (gesture.mode !== 'select') { this.draw(gesture); return; }
		if (!gesture.moved || !gesture.objects.length) { const object = !gesture.extend ? gesture.clicked : undefined; if (object && this.chosen.size > 1) { this.choose(object); } if (!object) { this.lastClick = undefined; } if (object) { const now = performance.now(); if (this.lastClick?.key === key(object) && now - this.lastClick.time < 450) { this.lastClick = undefined; this.edit(object); } else { this.lastClick = { key: key(object), time: now }; } } return; }
		const object = gesture.objects[0];
		if (gesture.grip === 'move') { this.run('object_move', { objects: gesture.objects.map(({ kind, index }) => ({ kind, index })), offset: gesture.end.map((value, i) => value - gesture.start[i]), generation: gesture.revision }); }
		else if (object.kind === 'text') { const rect = this.preview(gesture, object); this.run('text_edit', { paragraph: object.index, width: rect[2] - rect[0], dx: rect[0] - object.rect[0] }); }
		else if (object.kind === 'added') { this.run('content_update', { index: object.index, rect: this.preview(gesture, object) }); }
		else if (object.kind === 'image' && object.fullRect) {
			const visible = object.rect, preview = this.preview(gesture, object), sx = (preview[2] - preview[0]) / (visible[2] - visible[0]), sy = (preview[3] - preview[1]) / (visible[3] - visible[1]);
			const rect = object.fullRect.map((value, index) => preview[index % 2] + (value - visible[index % 2]) * (index % 2 ? sy : sx));
			this.run('image_edit', { image: object.index, action: 'move', rect });
		}
	}
	private cancel() { const gesture = this.gesture; this.gesture = undefined; if (gesture) { gesture.target.style.cursor = ''; if (gesture.target.hasPointerCapture(gesture.pointer)) { gesture.target.releasePointerCapture(gesture.pointer); } } this.pages.querySelectorAll('.pdf-drag-preview').forEach(node => node.remove()); }
	private run(command: string, args: Record<string, unknown>) { if (this.waiting) { return; } this.waiting = true; this.host.execute(command, { page: this.host.snapshot().state.page! + 1, ...args }, this.host.snapshot().state.revision); }
	private edit(object: PdfObject) {
		if (object.type !== 'text') { this.properties(object); return; } if (this.waiting) { return; }
		if (this.editor) { this.flush(); return; }
		this.choose(object); const { state, dimensions } = this.host.snapshot(), page = state.page || 0;
		const block = state.blocks?.find(block => object.kind === 'text' && block.id === object.index), text = block?.text ?? object.text ?? '';
		const input = document.createElement('textarea'); input.className = 'block-editor pdf-text-editor'; input.setAttribute('aria-label', 'Edit PDF text'); input.value = text;
		this.place(input, page, [object.rect[0], object.rect[1], object.rect[2], Math.max(object.rect[3], object.rect[1] + 32)]);
		input.style.fontSize = (block?.size || 12) * (this.pages.children[page] as HTMLElement).clientWidth / dimensions[page][0] + 'px';
		this.editor = { node: input, object, page, revision: state.revision, original: text };
		input.addEventListener('keydown', event => { if (event.key === 'Escape') { event.preventDefault(); event.stopPropagation(); input.remove(); this.editor = undefined; this.render(); } else if (event.key === 'Enter' && (event.ctrlKey || event.metaKey)) { event.preventDefault(); event.stopPropagation(); this.flush(); } });
		input.focus(); input.setSelectionRange(text.length, text.length); this.render();
	}
	flush() {
		const editor = this.editor; if (!editor) { return; } if (editor.revision !== this.host.snapshot().state.revision && editor.node.value !== editor.original) { this.host.message('The PDF changed. Copy your draft before reopening this text; Escape discards it.'); return; } this.editor = undefined; const value = editor.node.value; editor.node.remove();
		if (value !== editor.original) { if (editor.revision !== this.host.snapshot().state.revision) { this.host.message('The document changed. Reopen the text before editing.'); } else { this.run(editor.object.kind === 'added' ? 'content_update' : 'text_edit', { ...(editor.object.kind === 'added' ? { index: editor.object.index } : { paragraph: editor.object.index }), text: value }); } }
		this.render();
	}
	format(params: Record<string, unknown>) {
		if (this.waiting) { return; }
		const editor = this.editor, object = editor?.object || this.selected()[0];
		if (!object || object.type !== 'text' || this.selected().length > 1) { this.host.message('Select one text box first.'); return; }
		if (editor && editor.revision !== this.host.snapshot().state.revision) { this.host.message('The PDF changed. Copy your draft before reopening this text.'); return; }
		if (object.kind === 'added' && Object.hasOwn(params, 'line_spacing')) { this.host.message('Line spacing is available for existing PDF paragraphs; this added text box uses its native spacing.'); return; }
		if (editor) { this.editor = undefined; editor.node.remove(); }
		this.run(object.kind === 'added' ? 'content_update' : 'text_edit', { ...(object.kind === 'added' ? { index: object.index } : { paragraph: object.index }), ...(editor ? { text: editor.node.value } : {}), ...params });
	}
	private properties(object: PdfObject) { this.host.form(object.kind === 'image' ? 'image_edit' : object.kind === 'added' ? 'content_update' : 'text_edit', { page: this.page + 1, ...(object.kind === 'image' ? { image: object.index, action: 'move', rect: object.fullRect || object.rect } : object.kind === 'added' ? { index: object.index } : { paragraph: object.index }) }); }
	private remove(object: PdfObject) { this.run(object.kind === 'image' ? 'image_edit' : object.kind === 'added' ? 'content_delete' : 'text_edit', object.kind === 'image' ? { image: object.index, action: 'delete' } : object.kind === 'added' ? { index: object.index } : { paragraph: object.index, text: '' }); }
	private keydown(event: KeyboardEvent) {
		if ((event.target as HTMLElement).closest('textarea,input,button')) { return; }
		if (event.key === 'Escape') { this.cancel(); this.setMode('select'); this.menu?.remove(); return; }
		const object = this.selected()[0], node = (event.target as HTMLElement).closest<HTMLElement>('.pdf-object');
		if (node && (event.key === 'Enter' || event.key === ' ')) { event.preventDefault(); const hit = this.objects().find(item => key(item) === node.dataset.object); if (hit) { this.edit(hit); } return; }
		if (!object || this.waiting) { return; }
		const delta: Record<string, number[]> = { ArrowLeft: [-1, 0], ArrowRight: [1, 0], ArrowUp: [0, -1], ArrowDown: [0, 1] };
		if (delta[event.key]) { event.preventDefault(); event.stopPropagation(); this.run('object_move', { objects: this.selected().map(({ kind, index }) => ({ kind, index })), offset: delta[event.key].map(value => value * (event.shiftKey ? 10 : 1)), generation: this.revision }); }
		if (event.key === 'Delete' || event.key === 'Backspace') { event.preventDefault(); event.stopPropagation(); if (this.selected().length === 1) { this.remove(object); } }
	}
	private context(event: MouseEvent) {
		const node = (event.target as HTMLElement).closest<HTMLElement>('.pdf-object'); if (!node) { return; } event.preventDefault();
		const object = this.objects().find(item => key(item) === node.dataset.object); if (!object) { return; } this.choose(object); this.menu?.remove();
		const menu = this.menu = document.createElement('div'); menu.className = 'pdf-context'; menu.setAttribute('role', 'menu'); menu.tabIndex = -1;
		for (const [label, run] of [['Edit / Properties', () => this.edit(object)], ['Properties', () => this.properties(object)], ['Delete', () => this.remove(object)]] as [string, () => void][]) { const button = document.createElement('button'); button.textContent = label; button.role = 'menuitem'; button.addEventListener('click', () => { menu.remove(); run(); }); menu.append(button); }
		menu.style.left = Math.min(event.clientX, innerWidth - 210) + 'px'; menu.style.top = Math.min(event.clientY, innerHeight - 150) + 'px'; document.body.append(menu); menu.addEventListener('keydown', event => { if (event.key === 'Escape') { menu.remove(); } }); menu.addEventListener('focusout', event => { if (!menu.contains(event.relatedTarget as Node)) { menu.remove(); } }); menu.querySelector('button')?.focus();
	}
	setMode(mode: Mode) { this.flush(); this.cancel(); this.mode = mode; this.pages.dataset.pdfMode = mode; this.host.message(mode === 'select' ? '' : `Drag on the page to place ${mode === 'redact' ? 'a redaction mark (apply separately)' : mode}. Escape cancels.`); }
	private draw(gesture: Gesture) {
		let rect = this.drawRect(gesture); if (!gesture.moved) { rect = [rect[0], rect[1], rect[0] + 180, rect[1] + 36]; }
		this.setMode('select'); if (gesture.page !== this.host.snapshot().state.page) { this.host.page(gesture.page); } const page = gesture.page + 1;
		const form = (command: string, args: Record<string, unknown>) => this.host.form(command, { page, ...args });
		switch (gesture.mode) {
			case 'text': form('page_add_text', { at: rect.slice(0, 2), width: rect[2] - rect[0] }); break;
			case 'image': form('page_add_image', { rect }); break;
			case 'link': form('link_add', { rect }); break;
			case 'form': form('form_add_field', { type: 'text', rect }); break;
			case 'crop': this.host.form('page_set_box', { pages: [page], box: 'crop', rect }); break;
			case 'redact': this.run('redact_mark', { page, rect }); break;
			case 'note': form('comment_add', { type: 'note', at: rect.slice(0, 2) }); break;
			case 'arrow': this.run('comment_add', { page, type: 'arrow', from: gesture.start, to: gesture.end }); break;
			case 'highlight': this.run('comment_add', { page, type: 'highlight', quads: [[rect[0], rect[1], rect[2], rect[1], rect[0], rect[3], rect[2], rect[3]]] }); break;
			case 'rectangle': case 'oval': this.run('comment_add', { page, type: gesture.mode, rect }); break;
		}
	}
	openTools() {
		this.flush(); this.palette?.remove(); const dialog = this.palette = document.createElement('dialog'); dialog.className = 'pdf-tools-dialog pdf-toolbox'; dialog.setAttribute('aria-label', 'Edit PDF');
		const title = document.createElement('h2'); title.textContent = 'Edit PDF'; dialog.append(title);
		const hint = document.createElement('p'); hint.textContent = 'Select and drag objects. Double-click text to edit. Shift-click selects objects together.'; dialog.append(hint);
		const add = (label: string, run: () => void) => { const button = document.createElement('button'); button.textContent = label; button.title = label; button.addEventListener('click', () => { dialog.close(); run(); }); dialog.append(button); };
		for (const [mode, label] of [['select', 'Select / Move'], ['text', 'Add text'], ['image', 'Add image'], ['link', 'Add link'], ['note', 'Sticky note'], ['highlight', 'Highlight area'], ['rectangle', 'Rectangle'], ['oval', 'Oval'], ['arrow', 'Arrow'], ['form', 'Form field'], ['crop', 'Crop page'], ['redact', 'Mark redaction']] as [Mode, string][]) { add(label, () => this.setMode(mode)); }
		add('Organize pages', () => this.host.organize()); add('Apply redactions', () => this.host.form('redact_apply')); add('All PDF tools', () => this.host.form()); add('Close', () => {}); document.body.append(dialog); dialog.showModal();
	}
}
