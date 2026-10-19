/*---------------------------------------------------------------------------------------------
 *  Copyright (c) Microsoft Corporation. All rights reserved.
 *  Licensed under the MIT License. See License.txt in the project root for license information.
 *--------------------------------------------------------------------------------------------*/

import { wordMenus } from './wordMenus.js';
import { wordToolbarFeedback } from './wordToolbarFeedback.js';

interface WordToolbarControls {
	text?: HTMLButtonElement;
	undo: HTMLButtonElement;
	redo: HTMLButtonElement;
	save: HTMLButtonElement;
	versions?: HTMLButtonElement;
	properties?: HTMLButtonElement;
	style: HTMLSelectElement;
	font: HTMLSelectElement;
	size: HTMLSelectElement;
	emphasis: HTMLButtonElement[];
	strike?: HTMLButtonElement;
	color?: HTMLSelectElement;
	highlight?: HTMLSelectElement;
	clear?: HTMLButtonElement;
	align: HTMLSelectElement;
	list: HTMLSelectElement;
	spacing: HTMLSelectElement;
	table: HTMLButtonElement;
	image: HTMLButtonElement;
	link: HTMLButtonElement;
	comment: HTMLButtonElement;
}

// Original line icons. Keep the assets local to the webview and inherit theme colors.
const paths = {
	versions: 'M3 11a9 9 0 1 1 2 7M3 4v7h7M12 7v5l3 2',
	properties: 'M4 5h16M4 10h10M4 15h10M4 20h16M18 9v7m-3-3 3 3 3-3',
	color: 'm6 17 6-14 6 14M8 12h8M4 21h16',
	highlight: 'm8 15 9-11 4 4-9 11-4-4Zm0 0-4 4h8M3 22h18',
	clear: 'm4 14 9-11 8 7-9 11H9l-5-5v-2Zm5-6 8 7M14 21h8',
	undo: 'M9 5 4 10l5 5M4 10h10a6 6 0 0 1 0 12',
	redo: 'm15 5 5 5-5 5m5-5H10a6 6 0 0 0 0 12',
	save: 'M5 3h12l4 4v14H3V3h2m2 0v6h10V3M7 21v-8h10v8',
	align: 'M4 5h16M4 10h10M4 15h16M4 20h10',
	list: 'M9 5h11M9 12h11M9 19h11M4 5h.01M4 12h.01M4 19h.01',
	spacing: 'm3 7 3-3 3 3M6 4v16m-3-3 3 3 3-3M13 5h8M13 12h8M13 19h8',
	table: 'M3 4h18v16H3V4m0 5h18M3 14h18M9 4v16M15 4v16',
	image: 'M3 3h18v18H3V3m0 14 6-6 4 4 3-3 5 5M16 7h.01',
	link: 'm10 13 4-4m-6 7-1 1a4 4 0 0 1-6-6l4-4a4 4 0 0 1 6 0m2 10a4 4 0 0 0 6 0l4-4a4 4 0 0 0-6-6l-1 1',
	comment: 'M21 15a3 3 0 0 1-3 3H8l-5 4V5a3 3 0 0 1 3-3h12a3 3 0 0 1 3 3v10M8 10h8m-4-4v8',
	more: 'M5 12h.01M12 12h.01M19 12h.01',
};

function icon(name: keyof typeof paths, ownerDocument: Document): SVGSVGElement {
	const svg = ownerDocument.createElementNS('http://www.w3.org/2000/svg', 'svg');
	svg.setAttribute('viewBox', '0 0 24 24');
	svg.setAttribute('aria-hidden', 'true');
	svg.setAttribute('focusable', 'false');
	const path = ownerDocument.createElementNS(svg.namespaceURI, 'path');
	path.setAttribute('d', paths[name]);
	svg.append(path);
	return svg;
}

function iconButton(button: HTMLButtonElement, name: keyof typeof paths): HTMLButtonElement {
	const ownerDocument = button.ownerDocument;
	const label = ownerDocument.createElement('span');
	label.className = 'word-action-label';
	label.textContent = button.textContent;
	button.classList.add('word-icon-button');
	button.replaceChildren(icon(name, button.ownerDocument), label);
	return button;
}

/** Reorganize the existing controls without duplicating document commands or selection state. */
export function arrangeWordToolbar(toolbar: HTMLElement, controls: WordToolbarControls, insertTable: (rows: number, cols: number) => void, focusDocument: () => void): () => void {
	const ownerDocument = toolbar.ownerDocument;
	const targetWindow = ownerDocument.defaultView!;
	const events = new AbortController();
	const listen = { signal: events.signal };
	toolbar.classList.add('word-toolbar');

	toolbar.setAttribute('role', 'toolbar');
	toolbar.setAttribute('aria-orientation', 'horizontal');
	toolbar.setAttribute('aria-label', 'Document formatting');
	const row = ownerDocument.createElement('div');
	row.className = 'word-toolbar-row';
	const group = (label: string, children: HTMLElement[]) => {
		const element = ownerDocument.createElement('div');
		element.className = 'word-toolbar-group';
		element.setAttribute('role', 'group');
		element.setAttribute('aria-label', label);
		element.dataset.label = label;
		element.append(...children);
		return element;
	};
	controls.style.className = 'word-style-select';
	// The current bridge does not report selection formatting. Show action labels,
	// rather than pretending that the caret's style or font is known.
	controls.style.options[0].textContent = 'Paragraph Style';
	controls.style.title = 'Paragraph Style';
	controls.style.setAttribute('aria-label', 'Paragraph Style');
	controls.font.className = 'word-font-select';
	controls.size.className = 'word-size-select';
	controls.font.title = 'Font Family';
	controls.font.setAttribute('aria-label', 'Font Family');
	controls.size.title = 'Font Size';
	controls.size.setAttribute('aria-label', 'Font Size');
	for (const [index, label] of ['Bold', 'Italic', 'Underline'].entries()) {
		controls.emphasis[index].title = `${label} (${navigator.platform.includes('Mac') ? '⌘' : 'Ctrl+'}${['B', 'I', 'U'][index]})`;
		controls.emphasis[index].setAttribute('aria-label', label);
		controls.emphasis[index].classList.add('word-emphasis-button');
	}
	controls.save.classList.add('word-save-button');
	const history = group('Document', [iconButton(controls.undo, 'undo'), iconButton(controls.redo, 'redo'), iconButton(controls.save, 'save'), ...(controls.versions ? [iconButton(controls.versions, 'versions')] : [])]);
	const host = toolbar.closest<HTMLElement>('.monaco-workbench') || ownerDocument.body;
	const menus = wordMenus(controls.style, controls.table, insertTable, focusDocument, host, () => more.focus());
	const style = group('Styles', [menus.style]);
	const dropdowns = [menus.selector(controls.font, 'Font'), menus.selector(controls.size, 'Size'), menus.selector(controls.align, icon('align', ownerDocument)), menus.selector(controls.spacing, icon('spacing', ownerDocument)), menus.selector(controls.list, icon('list', ownerDocument))];
	const font = group('Font', dropdowns.slice(0, 2).map(item => item.element));
	const emphasis = group('Text Formatting', [...controls.emphasis, ...(controls.strike ? [controls.strike] : [])]);
	if (controls.strike) { controls.strike.classList.add('word-emphasis-button'); controls.strike.style.textDecoration = 'line-through'; }
	const colors = [controls.color, controls.highlight].filter((item): item is HTMLSelectElement => !!item).map(select => { select.dataset.palette = 'true'; return menus.selector(select, icon(select === controls.color ? 'color' : 'highlight', ownerDocument)); });
	dropdowns.push(...colors);
	emphasis.append(...colors.map(item => item.element));
	const paragraph = group('Paragraph', [...dropdowns.slice(2, 5).map(item => item.element), ...(controls.properties ? [iconButton(controls.properties, 'properties')] : [])]);
	const insert = group('Insert', [iconButton(controls.link, 'link'), iconButton(controls.comment, 'comment'), iconButton(controls.image, 'image')]);
	const table = group('Insert', [...(controls.text ? [controls.text] : []), iconButton(controls.table, 'table')]);
	const groups = [history, style, font, emphasis, paragraph, insert];
	const more = ownerDocument.createElement('button');
	more.type = 'button';
	more.textContent = 'More';
	more.title = 'More Formatting';
	more.setAttribute('aria-label', 'More Formatting');
	more.setAttribute('aria-expanded', 'false');
	iconButton(more, 'more');
	more.classList.add('word-more-button');
	const overflow = ownerDocument.createElement('div');
	overflow.id = 'word-toolbar-overflow';
	overflow.className = 'word-toolbar-overflow';
	overflow.popover = 'auto';
	overflow.setAttribute('role', 'group');
	overflow.setAttribute('aria-label', 'More Formatting');
	more.setAttribute('aria-controls', overflow.id);
	more.popoverTargetElement = overflow;
	if (controls.clear) { table.append(iconButton(controls.clear, 'clear')); }
	overflow.append(table);
	const allControls = [controls.undo, controls.redo, controls.save, controls.versions, controls.properties, controls.text, menus.styleButton, ...dropdowns.map(item => item.button), ...controls.emphasis, controls.strike, controls.clear, controls.table, controls.image, controls.link, controls.comment, more].filter((control): control is HTMLButtonElement => !!control);
	row.append(...groups, more);
	toolbar.replaceChildren(row);
	host.append(overflow);

	const positionOverflow = () => {
		const rect = more.getBoundingClientRect();
		overflow.style.left = `${Math.max(8, Math.min(rect.right - overflow.offsetWidth, targetWindow.innerWidth - overflow.offsetWidth - 8))}px`;
		overflow.style.top = `${rect.bottom + 8}px`;
		overflow.style.maxHeight = `${Math.max(80, targetWindow.innerHeight - rect.bottom - 16)}px`;
	};
	overflow.addEventListener('toggle', () => {
		const open = overflow.matches(':popover-open');
		more.setAttribute('aria-expanded', String(open));
		if (open) {
			positionOverflow();
			allControls.find(control => overflow.contains(control) && !control.disabled)?.focus({ preventScroll: true });
		}
	}, listen);
	// Hide before a command opens a modal prompt, retaining the engine selection.
	overflow.addEventListener('click', event => {
		if (event.target instanceof targetWindow.Element && event.target.closest('button') && !event.target.closest('[data-word-menu]')) { overflow.hidePopover(); }
	}, { ...listen, capture: true });
	overflow.addEventListener('change', () => overflow.hidePopover(), listen);
	overflow.addEventListener('keydown', event => {
		if (event.key === 'Escape') {
			event.preventDefault(); event.stopPropagation(); overflow.hidePopover(); more.focus();
		}
	}, listen);
	row.addEventListener('keydown', event => {
		if (!(event.target instanceof targetWindow.HTMLButtonElement) || !['ArrowLeft', 'ArrowRight', 'Home', 'End'].includes(event.key)) { return; }
		const items = allControls.filter(item => row.contains(item) && !item.disabled && item.getClientRects().length > 0);
		const index = items.indexOf(event.target);
		const next = event.key === 'Home' ? 0 : event.key === 'End' ? items.length - 1 : (index + (event.key === 'ArrowRight' ? 1 : -1) + items.length) % items.length;
		event.preventDefault(); items[next]?.focus();
	}, listen);
	const layout = () => {
		if (overflow.matches(':popover-open')) { overflow.hidePopover(); }
		const focused = ownerDocument.activeElement;
		for (const element of groups) { row.insertBefore(element, more); }
		// Move real controls, preserving their listeners, disabled state and values.
		for (const element of [insert, paragraph, font, style, emphasis, history]) {
			if (row.scrollWidth <= row.clientWidth) { break; }
			overflow.insertBefore(element, table);
		}
		if (focused instanceof targetWindow.HTMLElement && groups.some(element => element.contains(focused))) {
			if (overflow.contains(focused)) { more.focus({ preventScroll: true }); }
			else { focused.focus({ preventScroll: true }); }
		}
	};
	let width = row.clientWidth;
	const observer = new ResizeObserver(() => { if (row.clientWidth !== width) { width = row.clientWidth; layout(); } });
	observer.observe(row);
	layout();
	const disposeFeedback = wordToolbarFeedback(host);
	return () => {
		observer.disconnect(); events.abort(); disposeFeedback(); menus.dispose(); overflow.remove();
		toolbar.classList.remove('word-toolbar');
	};
}
