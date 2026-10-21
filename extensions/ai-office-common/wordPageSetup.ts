/*---------------------------------------------------------------------------------------------
 *  Copyright (c) Microsoft Corporation. All rights reserved.
 *  Licensed under the MIT License. See License.txt in the project root for license information.
 *--------------------------------------------------------------------------------------------*/

import type { State } from './types.js';
import type { PageBorder, PageBorderSides, PagePropertyPatch } from './pageProperties.js';

/** Page controls apply guarded patches; they never replace the full section or its header references. */
export class WordPageSetup {
	private readonly dialog = document.createElement('dialog');
	private readonly controller = new AbortController();
	private formController?: AbortController;
	private state: State = { revision: -1 };
	private origin?: State;
	constructor(private readonly apply: (patch: PagePropertyPatch, state: State) => void, private readonly focusDocument: () => void) {
		this.dialog.className = 'word-paragraph-dialog word-page-dialog'; document.body.append(this.dialog);
		this.dialog.addEventListener('close', () => { this.formController?.abort(); this.origin = undefined; this.focusDocument(); }, { signal: this.controller.signal });
	}
	update(state: State) {
		this.state = state;
		if (this.origin && !this.matches(this.origin)) {
			this.dialog.querySelector<HTMLElement>('[role="alert"]')!.textContent = 'Document or selection changed. Close and reopen Page Settings.';
			this.dialog.querySelector<HTMLButtonElement>('[type="submit"]')!.disabled = true;
		}
	}
	private matches(state: State) { return this.state.revision === state.revision && this.state.pageSetup?.page === state.pageSetup?.page && JSON.stringify(this.state.selection) === JSON.stringify(state.selection); }
	open(mode: 'margins' | 'borders') {
		const start = this.state, setup = start.pageSetup;
		if (this.dialog.open || !setup) { return; }
		this.formController?.abort(); this.formController = new AbortController(); const signal = this.formController.signal; this.origin = start;
		const form = document.createElement('form');
		const title = document.createElement('h2'); title.textContent = mode === 'margins' ? 'Page Margins' : 'Page Borders'; this.dialog.setAttribute('aria-label', title.textContent); form.append(title);
		const note = document.createElement('p'); note.textContent = mode === 'margins' ? 'Margins are measured in inches.' : 'Choose individual sides. Border width and distance from the page edge are measured in points.'; form.append(note);
		const error = document.createElement('p'); error.setAttribute('role', 'alert'); form.append(error);
		const fieldset = document.createElement('fieldset'); form.append(fieldset);
		const label = (name: string, input: HTMLElement, parent: HTMLElement = fieldset) => { input.setAttribute('aria-label', name); const element = document.createElement('label'); const span = document.createElement('span'); span.textContent = name; element.append(span, input); parent.append(element); };
		const select = (name: string, options: [string, string][], value: string, parent: HTMLElement = fieldset) => {
			const input = document.createElement('select'); for (const [text, id] of options) { input.add(new Option(text, id)); } input.value = value; label(name, input, parent); return input;
		};
		const number = (name: string, value: number, min: number, max: number, step: number) => {
			const input = document.createElement('input'); input.type = 'number'; input.required = true; input.value = String(Math.round(value * 1000) / 1000); input.min = String(min); input.max = String(max); input.step = String(step); label(name, input); return input;
		};
		let collect: () => Partial<PagePropertyPatch>;
		if (mode === 'margins') {
			const preset = select('Preset', [['Custom', 'custom'], ['Normal · 1 inch', 'normal'], ['Narrow · 0.5 inch', 'narrow'], ['Moderate', 'moderate'], ['Wide', 'wide']], 'custom');
			const inputs = { top: number('Top', setup.margins.top / 72, 0, 22, .01), bottom: number('Bottom', setup.margins.bottom / 72, 0, 22, .01), left: number('Left', setup.margins.left / 72, 0, 22, .01), right: number('Right', setup.margins.right / 72, 0, 22, .01), gutter: number('Gutter', setup.margins.gutter / 72, 0, 22, .01) };
			for (const input of Object.values(inputs)) { input.step = 'any'; }
			const presets: Record<string, number[]> = { normal: [1, 1, 1, 1], narrow: [.5, .5, .5, .5], moderate: [1, 1, .75, .75], wide: [1, 1, 2, 2] };
			preset.addEventListener('change', () => { const values = presets[preset.value]; if (values) { for (const [i, key] of (['top', 'bottom', 'left', 'right'] as const).entries()) { inputs[key].value = String(values[i]); } inputs.gutter.value = '0'; } }, { signal });
			for (const input of Object.values(inputs)) { input.addEventListener('input', () => { preset.value = 'custom'; }, { signal }); }
			collect = () => {
				const margins = { top: inputs.top.valueAsNumber * 72, bottom: inputs.bottom.valueAsNumber * 72, left: inputs.left.valueAsNumber * 72, right: inputs.right.valueAsNumber * 72, gutter: inputs.gutter.valueAsNumber * 72 };
				if (margins.left + margins.right + margins.gutter > setup.width - 36 || margins.top + margins.bottom > setup.height - 36) { throw new Error('Margins leave no room for text. Reduce the margins or gutter.'); }
				return { margins };
			};
		} else {
			const first = Object.values(setup.borders).find((side): side is PageBorder => !!side);
			const style = select('Style', ['single', 'double', 'dotted', 'dashed', 'thick', 'triple', 'dotDash', 'wave'].map(value => [value === 'dotDash' ? 'Dot Dash' : value[0].toUpperCase() + value.slice(1), value]), first?.style || 'single');
			const width = number('Width · pt', first?.width || 1, .25, 6, .25);
			const space = number('Distance from Page Edge · pt', first?.space ?? 24, 0, 31, 1);
			const color = document.createElement('input'); color.type = 'color'; color.value = '#' + (first?.color || '000000'); label('Color', color);
			const automatic = document.createElement('input'); automatic.type = 'checkbox'; automatic.checked = !first?.color; label('Automatic Color', automatic);
			const dirty = new Set<keyof PageBorder>();
			for (const [input, key] of [[style, 'style'], [width, 'width'], [space, 'space'], [color, 'color'], [automatic, 'color']] as const) { const mark = () => { dirty.add(key); if (input === color) { automatic.checked = false; } }; input.addEventListener('change', mark, { signal }); input.addEventListener('input', mark, { signal }); }
			const sideGroup = document.createElement('fieldset'); const legend = document.createElement('legend'); legend.textContent = 'Border Sides'; sideGroup.append(legend); form.append(sideGroup);
			const sideInput = (side: 'top' | 'bottom' | 'left' | 'right') => { const input = document.createElement('input'); input.type = 'checkbox'; input.checked = !!setup.borders[side]; label(side[0].toUpperCase() + side.slice(1) + ' Border', input, sideGroup); return input; };
			const sides = { top: sideInput('top'), bottom: sideInput('bottom'), left: sideInput('left'), right: sideInput('right') };
			const shortcuts = document.createElement('div'); shortcuts.className = 'word-page-border-actions';
			for (const [name, checked] of [['All Sides', true], ['No Border', false]] as const) { const button = document.createElement('button'); button.type = 'button'; button.textContent = name; button.addEventListener('click', () => { Object.values(sides).forEach(input => { input.checked = checked; }); }, { signal }); shortcuts.append(button); } form.append(shortcuts);
			collect = () => {
				const borders: PageBorderSides = {};
				const current: PageBorder = { style: style.value as PageBorder['style'], width: width.valueAsNumber, space: space.valueAsNumber, color: automatic.checked ? null : color.value.slice(1).toUpperCase() };
				for (const side of ['top', 'bottom', 'left', 'right'] as const) {
					if (!sides[side].checked) { borders[side] = null; continue; }
					const value = { ...(setup.borders[side] || current) }; for (const key of dirty) { Object.assign(value, { [key]: current[key] }); } borders[side] = value;
				}
				return { borders: Object.values(borders).some(Boolean) ? borders : null };
			};
		}
		const scopeArea = document.createElement('div'); scopeArea.className = 'word-page-scope'; form.append(scopeArea);
		const scope = select('Apply To', [['This Section', 'section'], ['Whole Document', 'document']], 'section', scopeArea);
		const actions = document.createElement('div'); actions.className = 'word-paragraph-actions';
		const cancel = document.createElement('button'); cancel.type = 'button'; cancel.textContent = 'Cancel'; cancel.addEventListener('click', () => this.dialog.close(), { signal });
		const apply = document.createElement('button'); apply.type = 'submit'; apply.textContent = 'Apply'; actions.append(cancel, apply); form.append(actions);
		form.addEventListener('submit', event => { event.preventDefault(); try { if (!this.matches(start)) { throw new Error('Document or selection changed. Reopen Page Settings.'); } this.apply({ page: setup.page, applyTo: scope.value as 'section' | 'document', ...collect() }, start); this.dialog.close(); } catch (reason) { error.textContent = reason instanceof Error ? reason.message : String(reason); } }, { signal });
		this.dialog.replaceChildren(form); this.dialog.showModal(); this.dialog.querySelector<HTMLInputElement | HTMLSelectElement>('input,select')?.focus();
	}
	dispose() { this.controller.abort(); this.formController?.abort(); this.dialog.remove(); }
}
