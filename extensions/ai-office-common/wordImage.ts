/*---------------------------------------------------------------------------------------------
 *  Copyright (c) Microsoft Corporation. All rights reserved.
 *  Licensed under the MIT License. See License.txt in the project root for license information.
 *--------------------------------------------------------------------------------------------*/
import type { ImageGeometry } from './types.js';
import { wordCursor } from './wordCursor.js';
type Point = { page: number; x: number; y: number };
type Place = (node: HTMLElement, page: number, x: number, y: number, width: number, height: number) => void;
type Gesture = { id: number; page: HTMLElement; image: ImageGeometry; start: Point; handle: string; rect: Point & { width: number; height: number }; ghost: HTMLCanvasElement; moved: boolean };

/** Contextual image controls. Drags preview locally; the drop is one native edit. */
export class WordImage {
	private revision = -1;
	private image?: ImageGeometry;
	private readonly geometry = new Map<number, ImageGeometry[]>();
	private readonly frame = document.createElement('div');
	private readonly menu = document.createElement('div');
	private readonly controller = new AbortController();
	private gesture?: Gesture;
	private menuTarget?: ImageGeometry;
	constructor(private readonly pages: HTMLElement, private readonly point: (event: PointerEvent | MouseEvent, page: HTMLElement) => Point, private readonly place: Place,
		private readonly execute: (command: string, args: Record<string, unknown>) => void, private readonly replace: (target: Record<string, unknown>) => void, private readonly focus: () => void,
		private readonly scale: (page: HTMLElement) => number) {
		this.frame.className = 'word-image-frame'; this.menu.className = 'word-image-menu'; this.menu.popover = 'manual'; this.menu.setAttribute('role', 'dialog'); this.menu.setAttribute('aria-label', 'Image options'); document.body.append(this.menu);
		const signal = this.controller.signal;
		pages.addEventListener('pointerdown', event => this.down(event), { capture: true, signal });
		pages.addEventListener('pointermove', event => this.move(event), { capture: true, signal });
		pages.addEventListener('pointerup', event => this.up(event), { capture: true, signal });
		pages.addEventListener('pointercancel', () => this.cancel(), { signal });
		pages.addEventListener('lostpointercapture', () => this.cancel(), { signal });
		pages.addEventListener('contextmenu', event => { const image = this.hit(event); if (image) { event.preventDefault(); event.stopImmediatePropagation(); this.select(image); this.showOptions(image); } }, { capture: true, signal });
		pages.addEventListener('dblclick', event => { const image = this.hit(event); if (image) { event.preventDefault(); event.stopImmediatePropagation(); this.showOptions(image); } }, { capture: true, signal });
		document.addEventListener('keydown', event => {
			if (this.menu.matches(':popover-open')) { if (event.key === 'Escape') { event.preventDefault(); event.stopImmediatePropagation(); this.close(); this.focus(); } return; }
			if (event.key === 'Escape' && this.gesture) { event.preventDefault(); event.stopImmediatePropagation(); this.cancel(); return; }
			if (!this.image || event.target instanceof HTMLInputElement || event.target instanceof HTMLTextAreaElement && event.target.id !== 'keyboard') { return; }
			if (event.key === 'Delete' || event.key === 'Backspace') { event.preventDefault(); event.stopImmediatePropagation(); this.run('picture.delete'); }
			else if (event.key.startsWith('Arrow') && this.image.wrap !== 'inline') { event.preventDefault(); event.stopImmediatePropagation(); const d = event.shiftKey ? 10 : 1; this.run('arrange.nudge', { dx: event.key === 'ArrowLeft' ? -d : event.key === 'ArrowRight' ? d : 0, dy: event.key === 'ArrowUp' ? -d : event.key === 'ArrowDown' ? d : 0 }); }
		}, { capture: true, signal });
		document.addEventListener('pointerdown', event => { if (!this.menu.contains(event.target as Node) && !this.frame.contains(event.target as Node)) { this.close(); } }, { signal });
		window.addEventListener('blur', () => this.cancel(), { signal }); window.addEventListener('resize', () => { this.cancel(); this.close(); }, { signal });
	}
	private target(image = this.image) { return image ? { revision: this.revision, parentId: image.parentId, offset: image.offset } : {}; }
	private run(command: string, args: Record<string, unknown> = {}) { this.close(); this.execute(command, { ...args, _image: this.target() }); this.focus(); }
	update(revision: number, image?: ImageGeometry | null) {
		if (revision !== this.revision) { this.cancel(); this.close(); this.geometry.clear(); }
		this.revision = revision; this.image = image || undefined; this.paint();
	}
	render(page: number, revision: number, images: ImageGeometry[]) { if (revision === this.revision) { this.geometry.set(page, images); this.paint(); } }
	forget(page: number) { this.geometry.delete(page); if (this.image?.page === page) { this.frame.remove(); } }
	private paint() {
		this.frame.remove(); const image = this.image; if (!image || this.gesture || !this.pages.children[image.page]) { return; }
		this.frame.replaceChildren(); this.place(this.frame, image.page, image.x, image.y, image.width, image.height); this.frame.style.transform = `rotate(${image.rotation}deg)`;
		for (const handle of ['nw', 'n', 'ne', 'e', 'se', 's', 'sw', 'w']) { const button = document.createElement('button'); button.className = 'word-image-handle ' + handle; button.dataset.handle = handle; button.setAttribute('aria-label', 'Resize image ' + handle); button.title = 'Drag to resize · Shift unlocks proportions'; button.style.cursor = wordCursor(handle === 'n' || handle === 's' ? 'row' : handle === 'e' || handle === 'w' ? 'column' : handle === 'ne' || handle === 'sw' ? 'opposite' : 'corner'); this.frame.append(button); }
		const options = document.createElement('button'); options.className = 'word-image-options'; options.textContent = '•••'; options.title = 'Image options'; options.setAttribute('aria-label', 'Image options'); options.addEventListener('pointerdown', e => { e.preventDefault(); e.stopPropagation(); }); options.addEventListener('click', e => { e.preventDefault(); e.stopPropagation(); this.showOptions(image); }); this.frame.append(options);
	}
	private hit(event: MouseEvent) {
		const page = (event.target as HTMLElement).closest<HTMLElement>('.page.word'); if (!page) { return; } const p = this.point(event, page);
		return this.geometry.get(p.page)?.find(i => { const angle = -i.rotation * Math.PI / 180, dx = p.x - i.x - i.width / 2, dy = p.y - i.y - i.height / 2, x = dx * Math.cos(angle) - dy * Math.sin(angle), y = dx * Math.sin(angle) + dy * Math.cos(angle); return Math.abs(x) <= i.width / 2 && Math.abs(y) <= i.height / 2; });
	}
	private select(image: ImageGeometry) { if (this.image?.parentId !== image.parentId || this.image.offset !== image.offset || this.image.page !== image.page) { this.execute('picture.select', { page: image.page, x: image.x + image.width / 2, y: image.y + image.height / 2, revision: this.revision }); } }
	private down(event: PointerEvent) {
		if (event.button !== 0 || (event.target as HTMLElement).closest('.word-image-options')) { return; }
		const handle = (event.target as HTMLElement).dataset.handle || 'move', image = this.hit(event) || (handle !== 'move' ? this.image : undefined); if (!image) { return; }
		event.preventDefault(); event.stopImmediatePropagation(); this.close(); this.select(image); this.focus();
		const page = this.pages.children[image.page] as HTMLElement, source = page.querySelector('canvas'); if (!source || !source.width) { return; }
		// Rotated pictures can be resized through precise properties; keep canvas drags axis-aligned.
		if (image.rotation % 360 !== 0) { return; }
		const ghost = document.createElement('canvas'); ghost.className = 'word-image-ghost';
		const scale = this.scale(page), bounds = page.getBoundingClientRect();
		const ratio = source.width / (bounds.width / scale), x = image.x * ratio, y = image.y * ratio;
		const w = image.width * ratio, h = image.height * ratio;
		ghost.width = Math.max(1, Math.min(1600, Math.round(w))); ghost.height = Math.max(1, Math.min(1600, Math.round(h))); ghost.getContext('2d')?.drawImage(source, x, y, w, h, 0, 0, ghost.width, ghost.height);
		this.gesture = { id: event.pointerId, page, image: { ...image }, start: this.point(event, page), handle, rect: { page: image.page, x: image.x, y: image.y, width: image.width, height: image.height }, ghost, moved: false }; page.setPointerCapture(event.pointerId);
	}
	private move(event: PointerEvent) {
		const g = this.gesture; if (!g) { const page = (event.target as HTMLElement).closest<HTMLElement>('.page.word'); if (page) { page.style.cursor = this.hit(event) ? wordCursor('move') : ''; } return; } event.preventDefault(); event.stopImmediatePropagation(); const at = this.point(event, g.page), dx = at.x - g.start.x, dy = at.y - g.start.y; if (Math.hypot(dx, dy) < 3 && !g.moved) { return; } g.moved = true;
		const i = g.image; let x = i.x, y = i.y, width = i.width, height = i.height;
		if (g.handle === 'move') { x = Math.max(0, x + dx); y = Math.max(0, y + dy); }
		else {
			if (g.handle.includes('e')) { width += dx; } if (g.handle.includes('s')) { height += dy; } if (g.handle.includes('w')) { width -= dx; } if (g.handle.includes('n')) { height -= dy; }
			if (!event.shiftKey) { if (g.handle === 'n' || g.handle === 's') { width = height * i.width / i.height; } else { height = width * i.height / i.width; } }
			width = Math.min(4000, Math.max(1, width)); height = Math.min(4000, Math.max(1, height)); if (g.handle.includes('w')) { x += i.width - width; } if (g.handle.includes('n')) { y += i.height - height; }
		}
		if (g.handle !== 'move' && i.wrap === 'inline') { x = i.x; y = i.y; }
		g.rect = { page: i.page, x, y, width, height }; this.frame.remove(); this.place(g.ghost, i.page, x, y, width, height);
	}
	private up(event: PointerEvent) {
		const g = this.gesture; if (!g) { return; } event.preventDefault(); event.stopImmediatePropagation(); const { moved, rect, image, handle } = g; this.cancel();
		if (moved) { if (handle === 'move' || image.wrap !== 'inline') { this.execute('arrange.bounds', { ...rect, _image: this.target(image) }); } else { this.execute('picture.size', { width: rect.width, height: rect.height, lockAspect: false, _image: this.target(image) }); } } this.focus();
	}
	private cancel() { const g = this.gesture; this.gesture = undefined; g?.ghost.remove(); if (g?.page.hasPointerCapture(g.id)) { g.page.releasePointerCapture(g.id); } this.paint(); }
	private close() { if (this.menu.matches(':popover-open')) { this.menu.hidePopover(); } this.menuTarget = undefined; }
	private showOptions(image: ImageGeometry) {
		this.close(); this.menuTarget = image; this.menu.replaceChildren();
		const title = document.createElement('strong'); title.textContent = 'Image options'; this.menu.append(title);
		const apply = (command: string, args: Record<string, unknown>) => { const target = this.menuTarget; if (!target) { return; } this.execute(command, { ...args, _image: this.target(target) }); this.close(); this.focus(); };
		const row = (name: string, values: [string, number][], command: string, extra: Record<string, unknown> = {}, factor = 1) => {
			const group = document.createElement('fieldset'), legend = document.createElement('legend'); legend.textContent = name; group.append(legend); const fields = new Map<string, HTMLInputElement>();
			for (const [key, value] of values) { const label = document.createElement('label'); label.textContent = key[0].toUpperCase() + key.slice(1); const input = document.createElement('input'); input.type = 'number'; input.value = String(Math.round(value * 100) / 100); input.step = '0.01'; if (command === 'picture.size') { input.min = String(1 / 72); input.max = String(4000 / 72); } if (command === 'picture.crop') { input.min = '0'; input.max = '45'; } input.required = true; label.append(input); group.append(label); fields.set(key, input); }
			let aspect: HTMLInputElement | undefined;
			if (command === 'picture.size') { const label = document.createElement('label'); aspect = document.createElement('input'); aspect.type = 'checkbox'; aspect.checked = true; label.append(aspect, 'Lock proportions'); group.append(label); for (const [key, input] of fields) { input.addEventListener('input', () => { if (aspect?.checked && Number.isFinite(input.valueAsNumber)) { const other = fields.get(key === 'width' ? 'height' : 'width')!; other.value = String(Math.round(input.valueAsNumber * (key === 'width' ? image.height / image.width : image.width / image.height) * 100) / 100); } }); } }
			const button = document.createElement('button'); button.textContent = 'Apply ' + name.toLowerCase(); button.addEventListener('click', () => { if ([...fields.values()].some(i => !i.reportValidity() || !Number.isFinite(i.valueAsNumber))) { return; } apply(command, { ...extra, ...(aspect ? { lockAspect: aspect.checked } : {}), ...Object.fromEntries([...fields].map(([k, v]) => [k, v.valueAsNumber * factor])) }); }); group.append(button); this.menu.append(group);
		};
		row('Size (inches)', [['width', image.width / 72], ['height', image.height / 72]], 'picture.size', { lockAspect: false }, 72);
		const reset = document.createElement('button'); reset.textContent = 'Restore original size'; reset.addEventListener('click', () => apply('picture.size', { original: true })); this.menu.append(reset);
		row('Crop (%)', [['left', image.crop[0] * 100], ['top', image.crop[1] * 100], ['right', image.crop[2] * 100], ['bottom', image.crop[3] * 100]], 'picture.crop', {}, .01);
		row('Rotation (degrees)', [['degrees', image.rotation]], 'arrange.rotation');
		const select = (label: string, options: string[], value: string, command: string, key: string) => { const wrap = document.createElement('label'); wrap.textContent = label; const input = document.createElement('select'); const placeholder = document.createElement('option'); placeholder.value = ''; placeholder.textContent = 'Choose…'; placeholder.disabled = true; input.append(placeholder); for (const option of options) { const node = document.createElement('option'); node.value = option; node.textContent = option.replace(/([A-Z])/g, ' $1'); input.append(node); } input.value = value; input.addEventListener('change', () => apply(command, { [key]: input.value })); wrap.append(input); this.menu.append(wrap); };
		select('Text wrapping', ['inline', 'square', 'tight', 'through', 'topAndBottom', 'behindText', 'inFrontOfText'], image.wrap, 'arrange.wrap', 'wrap');
		select('Align', ['left', 'center', 'right'], '', 'arrange.align', 'value');
		row('Position (points)', [['x', image.x], ['y', image.y]], 'arrange.bounds', { page: image.page });
		for (const [label, direction] of [['Flip horizontally', 'flipHorizontal'], ['Flip vertically', 'flipVertical']]) { const b = document.createElement('button'); b.textContent = label; b.addEventListener('click', () => apply('arrange.rotate', { direction })); this.menu.append(b); }
		const alt = document.createElement('label'); alt.textContent = 'Alt text'; const text = document.createElement('textarea'); text.value = image.alt; text.maxLength = 10000; alt.append(text); this.menu.append(alt);
		const decorative = document.createElement('label'); const check = document.createElement('input'); check.type = 'checkbox'; check.checked = image.decorative; decorative.append(check, 'Decorative image'); this.menu.append(decorative);
		const saveAlt = document.createElement('button'); saveAlt.textContent = 'Apply alt text'; saveAlt.addEventListener('click', () => apply('object.altText', { text: text.value, decorative: check.checked })); this.menu.append(saveAlt);
		const caption = document.createElement('label'); caption.textContent = 'Caption'; const captionText = document.createElement('input'); captionText.maxLength = 2000; caption.append(captionText); this.menu.append(caption);
		const position = document.createElement('select'); for (const value of ['below', 'above']) { const option = document.createElement('option'); option.value = value; option.textContent = 'Caption ' + value; position.append(option); } this.menu.append(position);
		const addCaption = document.createElement('button'); addCaption.textContent = 'Add numbered caption'; addCaption.addEventListener('click', () => apply('references.caption', { label: 'Figure', text: captionText.value, position: position.value })); this.menu.append(addCaption);
		for (const [label, action] of [['Replace image', () => { this.replace(this.target(image)); this.close(); }], ['Delete image', () => apply('picture.delete', {})]] as const) { const b = document.createElement('button'); b.textContent = label; b.addEventListener('click', action); this.menu.append(b); }
		this.menu.showPopover(); this.menu.querySelector('input')?.focus();
	}
	dispose() { this.controller.abort(); this.cancel(); this.close(); this.frame.remove(); this.menu.remove(); }
}
