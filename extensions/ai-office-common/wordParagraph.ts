/*---------------------------------------------------------------------------------------------
 *  Copyright (c) Microsoft Corporation. All rights reserved.
 *  Licensed under the MIT License. See License.txt in the project root for license information.
 *--------------------------------------------------------------------------------------------*/

import type { State } from './types.js';
import { validateParagraphProperties, type ParagraphProperties, type ParagraphTab } from './paragraphProperties.js';

/** Paragraph controls share the native property patch used by guarded block edits. */
export class WordParagraph {
	private readonly controller = new AbortController();
	private readonly dialog = document.createElement('dialog');
	private readonly ruler = document.createElement('div');
	private readonly observer: ResizeObserver;
	private state: State = { revision: -1 };
	private lastParagraph?: State['paragraph'];
	private dimensions: number[][] = [];
	private key = '';
	private origin?: { revision: number; selection: string };
	private formController?: AbortController;
	private gesture?: { pointer: number; x: number; width: number; scale: number; start: State; marker: HTMLElement; tab?: number; property?: 'left' | 'right' | 'firstLine'; value: number; remove?: boolean };

	constructor(private readonly pages: HTMLElement, private readonly apply: (patch: ParagraphProperties, state: State) => void, private readonly focusDocument: () => void) {
		this.dialog.className = 'word-paragraph-dialog'; this.dialog.setAttribute('aria-label', 'Paragraph Properties'); document.body.append(this.dialog);
		this.ruler.className = 'word-ruler'; this.ruler.title = 'Drag indents or tab stops. Click to add a tab; drag it away to remove it. Arrow keys adjust the focused marker.'; this.ruler.hidden = true; this.ruler.setAttribute('aria-label', 'Paragraph Ruler · Inches'); this.ruler.setAttribute('role', 'group'); pages.before(this.ruler);
		const signal = this.controller.signal;
		this.dialog.addEventListener('close', () => { this.formController?.abort(); this.origin = undefined; this.focusDocument(); }, { signal });
		this.ruler.addEventListener('pointerdown', event => this.pointerDown(event), { signal });
		this.ruler.addEventListener('pointermove', event => this.pointerMove(event), { signal });
		this.ruler.addEventListener('pointerup', event => this.pointerUp(event), { signal });
		this.ruler.addEventListener('pointercancel', () => { this.gesture = undefined; this.key = ''; this.layout(); }, { signal });
		this.ruler.addEventListener('keydown', event => this.rulerKey(event), { signal });
		this.observer = new ResizeObserver(() => { this.key = ''; this.layout(); }); this.observer.observe(pages);
	}

	update(state: State, dimensions: number[][]) {
		this.state = state; this.dimensions = dimensions;
		if (state.paragraph?.ruler) { this.lastParagraph = state.paragraph; }
		if (this.origin && (state.revision !== this.origin.revision || JSON.stringify(state.selection) !== this.origin.selection)) {
			const error = this.dialog.querySelector<HTMLElement>('[role="alert"]'); if (error) { error.textContent = 'Document or selection changed. Close and reopen Paragraph Properties.'; }
			const apply = this.dialog.querySelector<HTMLButtonElement>('[type="submit"]'); if (apply) { apply.disabled = true; }
		}
		if (this.gesture && (state.revision !== this.gesture.start.revision || JSON.stringify(state.selection) !== JSON.stringify(this.gesture.start.selection))) { this.gesture = undefined; this.key = ''; }
		this.layout();
	}

	open() {
		if (this.dialog.open || !this.state.paragraph) { return; }
		this.formController?.abort(); this.formController = new AbortController(); const signal = this.formController.signal;
		const start = this.state; const properties = start.paragraph!.properties; const mixed = start.paragraph!.mixed;
		this.origin = { revision: start.revision, selection: JSON.stringify(start.selection) };
		const patch: ParagraphProperties = {};
		const form = document.createElement('form'); const title = document.createElement('h2'); title.textContent = 'Paragraph Properties'; form.append(title);
		const note = document.createElement('p'); note.textContent = 'Applies to selected paragraphs. Measurements are in points; blank values are mixed.'; form.append(note);
		const error = document.createElement('p'); error.setAttribute('role', 'alert'); form.append(error);
		const group = (name: string) => { const fieldset = document.createElement('fieldset'); const legend = document.createElement('legend'); legend.textContent = name; fieldset.append(legend); form.append(fieldset); return fieldset; };
		const label = (parent: HTMLElement, name: string, input: HTMLElement) => { const element = document.createElement('label'); const text = document.createElement('span'); text.textContent = name; element.append(text, input); parent.append(element); };
		const numeric = (parent: HTMLElement, name: string, value: number | undefined, min: number, max: number, changed: (value: number) => void, eventSignal = signal) => {
			const input = document.createElement('input'); input.type = 'number'; input.min = String(min); input.max = String(max); input.step = 'any'; input.value = value === undefined ? '' : String(Math.round(value * 100) / 100); input.placeholder = 'Mixed';
			input.addEventListener('input', () => { input.setCustomValidity(input.value === '' ? 'Enter a value, or cancel to keep existing formatting.' : ''); if (input.validity.valid) { changed(input.valueAsNumber); } }, { signal: eventSignal }); label(parent, name, input); return input;
		};
		const select = (parent: HTMLElement, name: string, options: [string, string][], value: string | undefined, changed: (value: string) => void, eventSignal = signal) => {
			const input = document.createElement('select'); if (value === undefined) { input.add(new Option('Mixed', '')); }
			for (const [text, id] of options) { input.add(new Option(text, id)); } input.value = value ?? '';
			input.addEventListener('change', () => { if (input.value) { changed(input.value); } }, { signal: eventSignal }); label(parent, name, input); return input;
		};
		const layout = group('Alignment and Indentation');
		select(layout, 'Alignment', ['left', 'center', 'right', 'justify', 'distribute'].map(value => [value[0].toUpperCase() + value.slice(1), value]), mixed.includes('alignment') ? undefined : properties.alignment, value => { patch.alignment = value as ParagraphProperties['alignment']; });
		for (const [key, name] of [['left', 'Left Indent'], ['right', 'Right Indent'], ['firstLine', 'First Line · Negative for Hanging']] as const) {
			numeric(layout, name, mixed.includes('indents') ? undefined : properties.indents?.[key], -1584, 1584, value => { patch.indents = { ...patch.indents, [key]: value }; });
		}
		const spacing = group('Spacing');
		for (const [key, name] of [['before', 'Before Paragraph'], ['after', 'After Paragraph']] as const) { numeric(spacing, name, mixed.includes('spacing') ? undefined : properties.spacing?.[key], 0, 1584, value => { patch.spacing = { ...patch.spacing, [key]: value }; }); }
		const lines = properties.lineSpacing ?? { rule: 'multiple', value: 1 };
		const lineValue = numeric(spacing, 'Line Spacing Value', mixed.includes('lineSpacing') ? undefined : lines.value, .5, 1584, value => { patch.lineSpacing = { rule: lineRule.value as typeof lines.rule, value }; });
		const lineRule = select(spacing, 'Line Spacing Rule', [['Multiple', 'multiple'], ['Exactly · Points', 'exactly'], ['At Least · Points', 'atLeast']], mixed.includes('lineSpacing') ? undefined : lines.rule, value => {
			lineValue.min = value === 'multiple' ? '.5' : '1'; lineValue.max = value === 'multiple' ? '10' : '1584';
			lineValue.value = value === 'multiple' ? '1' : '12'; lineValue.setCustomValidity(''); patch.lineSpacing = { rule: value as typeof lines.rule, value: Number(lineValue.value) };
		});
		lineValue.min = lines.rule === 'multiple' ? '.5' : '1'; lineValue.max = lines.rule === 'multiple' ? '10' : '1584';
		const pagination = group('Line and Page Breaks');
		for (const [key, name] of [['keepNext', 'Keep with Next'], ['keepLines', 'Keep Lines Together'], ['widowControl', 'Widow/Orphan Control'], ['pageBreakBefore', 'Page Break Before']] as const) {
			const input = document.createElement('input'); input.type = 'checkbox'; input.checked = !!properties[key]; input.indeterminate = mixed.includes(key);
			input.addEventListener('change', () => { patch[key] = input.checked; }, { signal }); label(pagination, name, input);
		}
		const appearance = group('Borders and Shading');
		const borderWidth = numeric(appearance, 'Border Width', .5, .25, 6, () => setBorder());
		const borderColor = document.createElement('input'); borderColor.type = 'color'; borderColor.value = '#333333'; label(appearance, 'Border Color', borderColor);
		const borderStyle = select(appearance, 'Border Style', [['Single', 'single'], ['Double', 'double'], ['Dotted', 'dotted'], ['Dashed', 'dashed']], 'single', () => setBorder());
		const setBorder = () => {
			if (!borderKind.value || borderKind.value === 'keep') { return; }
			const b = { width: borderWidth.valueAsNumber, color: borderColor.value.slice(1), style: borderStyle.value as 'single' | 'double' | 'dotted' | 'dashed' };
			patch.borders = borderKind.value === 'none' || borderKind.value === 'all' ? Object.fromEntries(['top', 'left', 'bottom', 'right', 'between'].map(side => [side, borderKind.value === 'all' && side !== 'between' ? b : null])) : { [borderKind.value]: b };
		};
		const borderKind = select(appearance, 'Apply Borders', [['Keep Existing', 'keep'], ['None', 'none'], ['Outside', 'all'], ['Bottom', 'bottom'], ['Top', 'top'], ['Left', 'left'], ['Right', 'right']], 'keep', () => { if (borderKind.value === 'keep') { delete patch.borders; } else { setBorder(); } });
		borderColor.addEventListener('input', () => { if (borderKind.value !== 'keep') { setBorder(); } }, { signal });
		const fill = document.createElement('input'); fill.type = 'color'; fill.value = '#' + (properties.shading || 'ffffff'); fill.addEventListener('input', () => { patch.shading = fill.value.slice(1); }, { signal }); label(appearance, 'Paragraph Fill · White for White Paper', fill);
		const tabs = group('Tab Stops'); const tabRows = document.createElement('div'); tabRows.className = 'word-tab-rows'; tabs.append(tabRows);
		const stops: ParagraphTab[] = structuredClone(properties.tabs || []);
		const tabsChanged = () => { patch.tabs = stops; };
		let tabController = new AbortController();
		signal.addEventListener('abort', () => tabController.abort(), { once: true });
		const drawTabs = () => {
			tabController.abort(); tabController = new AbortController(); const tabSignal = tabController.signal;
			tabRows.replaceChildren();
			for (const [index, stop] of stops.entries()) {
				const row = document.createElement('div'); row.className = 'word-tab-row'; tabRows.append(row);
				const pos = numeric(row, 'Position', stop.pos, 0, 1584, value => { stop.pos = value; tabsChanged(); }, tabSignal); pos.required = true;
				select(row, 'Alignment', ['left', 'center', 'right', 'decimal', 'bar'].map(value => [value, value]), stop.align ?? 'left', value => { stop.align = value as ParagraphTab['align']; tabsChanged(); }, tabSignal);
				select(row, 'Leader', [['None', 'none'], ['Dots', 'dot'], ['Dashes', 'hyphen'], ['Underline', 'underscore'], ['Middle Dots', 'middleDot']], stop.leader ?? 'none', value => { stop.leader = value as ParagraphTab['leader']; tabsChanged(); }, tabSignal);
				const remove = document.createElement('button'); remove.type = 'button'; remove.textContent = 'Remove'; remove.setAttribute('aria-label', `Remove Tab Stop ${index + 1}`); remove.addEventListener('click', () => { stops.splice(index, 1); tabsChanged(); drawTabs(); }, { signal: tabSignal }); row.append(remove);
			}
		}; drawTabs();
		const add = document.createElement('button'); add.type = 'button'; add.textContent = 'Add Tab Stop'; add.addEventListener('click', () => { if (stops.length >= 64) { error.textContent = 'At most 64 tab stops are supported.'; return; } let pos = 36; while (stops.some(t => Math.abs(t.pos - pos) < .01) && pos < 1584) { pos += 36; } stops.push({ pos: Math.min(pos, 1584), align: 'left', leader: 'none' }); tabsChanged(); drawTabs(); }, { signal }); tabs.append(add);
		if (mixed.includes('tabs')) { const note = document.createElement('p'); note.textContent = 'Tab stops differ. Editing this list replaces them on every selected paragraph.'; tabs.append(note); }
		const actions = document.createElement('div'); actions.className = 'word-paragraph-actions';
		const cancel = document.createElement('button'); cancel.type = 'button'; cancel.textContent = 'Cancel'; cancel.addEventListener('click', () => this.dialog.close(), { signal });
		const apply = document.createElement('button'); apply.type = 'submit'; apply.textContent = 'Apply'; actions.append(cancel, apply); form.append(actions);
		form.addEventListener('submit', event => { event.preventDefault(); try { if (this.state.revision !== start.revision || JSON.stringify(this.state.selection) !== JSON.stringify(start.selection)) { throw new Error('Document or selection changed. Reopen Paragraph Properties.'); } validateParagraphProperties(patch); this.apply(patch, start); this.dialog.close(); } catch (reason) { error.textContent = String(reason instanceof Error ? reason.message : reason); } }, { signal });
		this.dialog.replaceChildren(form); this.dialog.showModal();
	}

	private layout() {
		if (this.gesture) { return; }
		const paragraph = this.state.paragraph ?? this.lastParagraph; const page = this.state.caret?.page ?? 0;
		const node = this.pages.children[page] as HTMLElement | undefined; const width = this.dimensions[page]?.[0];
		const body = this.state.paragraph?.ruler ?? this.state.pageRuler ?? this.lastParagraph?.ruler;
		this.ruler.hidden = !body || !node || !width; if (this.ruler.hidden) { return; }
		const mixed = !this.state.paragraph?.ruler || paragraph!.mixed.includes('indents');
		const key = JSON.stringify([paragraph, body, mixed, node!.clientWidth, width, page]); if (key === this.key) { return; } this.key = key;
		const focus = document.activeElement instanceof HTMLElement && this.ruler.contains(document.activeElement) ? document.activeElement.dataset.marker : undefined;
		const scale = node!.clientWidth / width; this.ruler.style.width = node!.clientWidth + 'px'; this.ruler.replaceChildren();
		const track = document.createElement('div'); track.className = 'word-ruler-body'; track.style.left = body!.left * scale + 'px'; track.style.width = body!.width * scale + 'px'; this.ruler.append(track);
		const unit = document.createElement('span'); unit.className = 'word-ruler-unit'; unit.textContent = 'in'; unit.setAttribute('aria-hidden', 'true'); this.ruler.append(unit);
		const indent = paragraph?.properties.indents ?? {};
		// Match the upstream ruler: zero is the text margin, and graduations follow page scale.
		// Omit fine divisions and crowded labels at small widths; never rescale the document units.
		const subdivision = scale * 9 >= 5 ? 9 : 18;
		const labelStep = scale * 72 >= 44 ? 72 : 144;
		const handles = mixed ? [] : [indent.left || 0, (indent.left || 0) + (indent.firstLine || 0), body!.width - (indent.right || 0), ...(paragraph?.properties.tabs || []).map(tab => tab.pos)];
		for (let pos = Math.ceil(-body!.left / subdivision) * subdivision; pos <= width - body!.left; pos += subdivision) {
			const x = (body!.left + pos) * scale;
			if (x < 3 || x > node!.clientWidth - 3) { continue; }
			const tick = document.createElement('span'); tick.className = 'word-ruler-tick'; tick.setAttribute('aria-hidden', 'true'); tick.style.left = x + 'px';
			tick.dataset.kind = pos % 72 === 0 ? 'major' : pos % 36 === 0 ? 'half' : 'minor'; this.ruler.append(tick);
			if (pos > 0 && pos < body!.width && pos % labelStep === 0 && handles.every(handle => Math.abs(handle - pos) * scale > 12)) {
				const label = document.createElement('span'); label.className = 'word-ruler-number'; label.style.left = x + 'px'; label.textContent = String(pos / 72); label.setAttribute('aria-hidden', 'true'); this.ruler.append(label);
			}
		}
		const marker = (key: string, name: string, pos: number, value: number, align = 'left') => {
			const b = document.createElement('button'); b.type = 'button'; b.dataset.marker = key; b.title = name;
			b.setAttribute('aria-label', name); b.setAttribute('role', 'slider'); b.setAttribute('aria-orientation', 'horizontal');
			b.setAttribute('aria-valuemin', key.startsWith('tab:') ? '0' : '-1584'); b.setAttribute('aria-valuemax', '1584');
			b.style.left = (body!.left + pos) * scale + 'px';
			const svg = document.createElementNS('http://www.w3.org/2000/svg', 'svg'); svg.setAttribute('viewBox', '0 0 12 10'); svg.setAttribute('aria-hidden', 'true');
			const path = document.createElementNS(svg.namespaceURI, 'path');
			const tabPaths: Record<string, string> = { left: 'M3 1v7h7', right: 'M9 1v7H2', center: 'M6 1v7M2 8h8', decimal: 'M6 1v7M2 8h8M10 4h.01', bar: 'M6 1v8' };
			path.setAttribute('d', key.startsWith('tab:') ? tabPaths[align] || tabPaths.left : key === 'firstLine' ? 'M1 1h10L6 7Z' : 'M1 7 6 1l5 6ZM1 8h10'); svg.append(path); b.append(svg);
			const hint = document.createElement('span'); hint.className = 'word-ruler-value'; b.append(hint); this.markerValue(b, value);
			this.ruler.append(b);
		};
		if (mixed) { return; } // Keep the page scale visible without suggesting a shared indent.
		marker('left', 'Left Indent · Arrow Keys Adjust', indent.left || 0, indent.left || 0);
		marker('firstLine', 'First-Line Indent · Arrow Keys Adjust', (indent.left || 0) + (indent.firstLine || 0), indent.firstLine || 0);
		marker('right', 'Right Indent · Arrow Keys Adjust', body!.width - (indent.right || 0), indent.right || 0);
		if (!paragraph!.mixed.includes('tabs')) { for (const [index, stop] of (paragraph!.properties.tabs || []).entries()) { marker(`tab:${index}`, `Tab Stop ${index + 1} · Delete Removes`, stop.pos, stop.pos, stop.align); } }
		if (focus) { Array.from(this.ruler.querySelectorAll<HTMLElement>('[data-marker]')).find(e => e.dataset.marker === focus)?.focus(); }
	}

	private markerValue(marker: HTMLElement, value: number) {
		const inches = `${Number((value / 72).toFixed(2))} in`;
		marker.setAttribute('aria-valuenow', String(value)); marker.setAttribute('aria-valuetext', `${inches} (${value} points)`);
		const hint = marker.querySelector('.word-ruler-value'); if (hint) { hint.textContent = inches; }
	}

	private pointerDown(event: PointerEvent) {
		if (event.button !== 0 || !this.state.paragraph?.ruler || this.state.paragraph.mixed.includes('indents')) { return; }
		const marker = (event.target as HTMLElement).closest<HTMLElement>('[data-marker]');
		const rect = this.ruler.getBoundingClientRect(); const page = this.state.caret?.page ?? 0; const scale = rect.width / this.dimensions[page][0]; const body = this.state.paragraph.ruler;
		if (!marker) {
			if (event.target === this.ruler || (event.target as HTMLElement).classList.contains('word-ruler-body')) {
				const pos = Math.round(((event.clientX - rect.left) / scale - body.left) / 6) * 6;
				if (pos >= 0 && pos <= body.width && !this.state.paragraph.mixed.includes('tabs')) { const tabs = this.state.paragraph.properties.tabs || []; if (tabs.length < 64 && !tabs.some(t => Math.abs(t.pos - pos) < .01)) { this.apply({ tabs: [...tabs, { pos, align: 'left', leader: 'none' }] }, this.state); } }
			} return;
		}
		event.preventDefault(); marker.focus({ preventScroll: true }); marker.dataset.dragging = 'true'; const key = marker.dataset.marker!;
		this.gesture = { pointer: event.pointerId, x: rect.left, width: rect.width, scale, start: this.state, marker, value: Number(marker.getAttribute('aria-valuenow')), ...(key.startsWith('tab:') ? { tab: Number(key.slice(4)) } : { property: key as 'left' | 'right' | 'firstLine' }) };
		this.ruler.setPointerCapture(event.pointerId);
	}
	private pointerMove(event: PointerEvent) {
		const g = this.gesture; if (!g || g.pointer !== event.pointerId) { return; }
		const paragraph = g.start.paragraph!; const body = paragraph.ruler!; const indents = paragraph.properties.indents || {};
		const pos = (event.clientX - g.x) / g.scale - body.left;
		const raw = g.property === 'right' ? body.width - pos : g.property === 'firstLine' ? pos - (indents.left || 0) : pos;
		g.value = Math.max(g.tab === undefined ? -1584 : 0, Math.min(1584, event.altKey ? Math.round(raw * 4) / 4 : Math.round(raw / 6) * 6));
		const rect = this.ruler.getBoundingClientRect(); g.remove = g.tab !== undefined && (event.clientY < rect.top - 20 || event.clientY > rect.bottom + 20);
		g.marker.style.opacity = g.remove ? '.3' : '1'; this.markerValue(g.marker, g.value);
		g.marker.style.left = (body.left + (g.property === 'right' ? body.width - g.value : g.property === 'firstLine' ? (indents.left || 0) + g.value : g.value)) * g.scale + 'px';
	}
	private pointerUp(event: PointerEvent) {
		const g = this.gesture; if (!g || g.pointer !== event.pointerId) { return; } this.pointerMove(event); this.gesture = undefined;
		if (this.ruler.hasPointerCapture(event.pointerId)) { this.ruler.releasePointerCapture(event.pointerId); }
		if (g.tab !== undefined) { const tabs = structuredClone(g.start.paragraph!.properties.tabs || []); if (g.remove) { tabs.splice(g.tab, 1); } else { tabs[g.tab].pos = g.value; } if (!tabs.some((t, i) => tabs.some((other, j) => j !== i && Math.abs(t.pos - other.pos) < .01))) { this.apply({ tabs }, g.start); } }
		else if (g.property) { this.apply({ indents: { [g.property]: g.value } }, g.start); }
		this.key = ''; this.layout(); this.focusDocument();
	}
	private rulerKey(event: KeyboardEvent) {
		const marker = (event.target as HTMLElement).closest<HTMLElement>('[data-marker]'); if (!marker || !this.state.paragraph) { return; }
		const key = marker.dataset.marker!; const isTab = key.startsWith('tab:');
		if (!['ArrowLeft', 'ArrowRight', 'Delete', 'Backspace'].includes(event.key) || !isTab && ['Delete', 'Backspace'].includes(event.key)) { return; }
		event.preventDefault(); const value = Number(marker.getAttribute('aria-valuenow')) + (event.key === 'ArrowLeft' ? -1 : 1) * (event.shiftKey ? 10 : 1);
		if (isTab) { const tabs = structuredClone(this.state.paragraph.properties.tabs || []); const index = Number(key.slice(4)); if (event.key === 'Delete' || event.key === 'Backspace') { tabs.splice(index, 1); } else { tabs[index].pos = Math.max(0, Math.min(1584, value)); } try { validateParagraphProperties({ tabs }); this.apply({ tabs }, this.state); } catch { /* Keep the current stops when a move would collide. */ } }
		else { this.apply({ indents: { [key]: Math.max(-1584, Math.min(1584, value)) } }, this.state); }
	}

	dispose() { this.observer.disconnect(); this.controller.abort(); this.formController?.abort(); this.dialog.remove(); this.ruler.remove(); }
}
