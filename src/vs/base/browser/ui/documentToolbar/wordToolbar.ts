/*---------------------------------------------------------------------------------------------
 *  Copyright (c) Microsoft Corporation. All rights reserved.
 *  Licensed under the MIT License. See License.txt in the project root for license information.
 *--------------------------------------------------------------------------------------------*/

import { wordMenus } from './wordMenus.js';
import { wordToolbarFeedback } from './wordToolbarFeedback.js';

interface WordToolbarControls {
	text?: HTMLButtonElement;
	margins?: HTMLButtonElement;
	pageBorders?: HTMLButtonElement;
	header?: HTMLButtonElement;
	footer?: HTMLButtonElement;
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

// Original line icons. Keep the assets local to the webview and inherit theme colors;
// palette tools carry a themed guide bar (text color, highlight) as in Docs.
interface IconSegment {
	d: string;
	stroke?: string;
}
const paths = {
	versions: [{ d: 'M3 11a9 9 0 1 1 2 7M3 4v7h7M12 7v5l3 2' }],
	properties: [{ d: 'M4 5h16M4 10h10M4 15h10M4 20h16M18 9v7m-3-3 3 3 3-3' }],
	color: [{ d: 'm6 17 6-14 6 14M8 12h8' }, { d: 'M4 21h16', stroke: 'var(--word-red, #d93025)' }],
	highlight: [{ d: 'm8 15 9-11 4 4-9 11-4-4Zm0 0-4 4h8' }, { d: 'M3 22h18', stroke: 'var(--word-yellow, #f9cc0d)' }],
	clear: [{ d: 'm4 14 9-11 8 7-9 11H9l-5-5v-2Zm5-6 8 7M14 21h8' }],
	undo: [{ d: 'M9 5 4 10l5 5M4 10h10a6 6 0 0 1 0 12' }],
	redo: [{ d: 'm15 5 5 5-5 5m5-5H10a6 6 0 0 0 0 12' }],
	save: [{ d: 'M5 3h12l4 4v14H3V3h2m2 0v6h10V3M7 21v-8h10v8' }],
	paintFormat: [{ d: 'M5 3h10v6H5zM15 4h3a2 2 0 0 1 2 2v2a2 2 0 0 1-2 2h-5M9 9v3a2 2 0 0 0 2 2h1v7' }],
	align: [{ d: 'M4 5h16M4 10h10M4 15h16M4 20h10' }],
	list: [{ d: 'M9 5h11M9 12h11M9 19h11M4 5h.01M4 12h.01M4 19h.01' }],
	spacing: [{ d: 'm3 7 3-3 3 3M6 4v16m-3-3 3 3 3-3M13 5h8M13 12h8M13 19h8' }],
	table: [{ d: 'M3 4h18v16H3V4m0 5h18M3 14h18M9 4v16M15 4v16' }],
	image: [{ d: 'M3 3h18v18H3V3m0 14 6-6 4 4 3-3 5 5M16 7h.01' }],
	link: [{ d: 'm10 13 4-4m-6 7-1 1a4 4 0 0 1-6-6l4-4a4 4 0 0 1 6 0m2 10a4 4 0 0 0 6 0l4-4a4 4 0 0 0-6-6l-1 1' }],
	comment: [{ d: 'M21 15a3 3 0 0 1-3 3H8l-5 4V5a3 3 0 0 1 3-3h12a3 3 0 0 1 3 3v10M8 10h8m-4-4v8' }],
	more: [{ d: 'M5 12h.01M12 12h.01M19 12h.01' }],
} satisfies Record<string, IconSegment[]>;

function icon(name: keyof typeof paths, ownerDocument: Document): SVGSVGElement {
	const svg = ownerDocument.createElementNS('http://www.w3.org/2000/svg', 'svg');
	svg.setAttribute('viewBox', '0 0 24 24');
	svg.setAttribute('aria-hidden', 'true');
	svg.setAttribute('focusable', 'false');
	for (const segment of paths[name]) {
		const path = ownerDocument.createElementNS(svg.namespaceURI, 'path');
		path.setAttribute('d', segment.d);
		if (segment.stroke) { path.style.stroke = segment.stroke; }
		svg.append(path);
	}
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
	// Formatting values are supplied by the native selection state.
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
	// Docs-style row order: history, paint format, style, font+size, emphasis,
	// color+highlight, reset, spacing/align/list, inserts, more, suggestions.
	const more = ownerDocument.createElement('button');
	more.type = 'button';
	more.textContent = 'More';
	more.title = 'More Formatting';
	more.setAttribute('aria-label', 'More Formatting');
	more.setAttribute('aria-expanded', 'false');
	iconButton(more, 'more');
	more.classList.add('word-more-button');
	const host = toolbar.closest<HTMLElement>('.monaco-workbench') || ownerDocument.body;
	const menus = wordMenus(controls.style, controls.table, insertTable, focusDocument, host, () => more.focus());
	const dropdowns = [menus.selector(controls.font, 'Font'), menus.selector(controls.size, 'Size'), menus.selector(controls.align, icon('align', ownerDocument)), menus.selector(controls.spacing, icon('spacing', ownerDocument)), menus.selector(controls.list, icon('list', ownerDocument))];
	const decrease = ownerDocument.createElement('button');
	const increase = ownerDocument.createElement('button');
	for (const [button, label, text, delta] of [[decrease, 'Decrease Font Size', '−', -1], [increase, 'Increase Font Size', '+', 1]] as const) {
		button.type = 'button'; button.textContent = text; button.title = label; button.setAttribute('aria-label', label);
		button.classList.add('word-size-step');
		button.addEventListener('pointerdown', event => event.preventDefault(), listen);
		button.addEventListener('click', () => {
			const current = Number(controls.size.dataset.currentSize || controls.size.dataset.lastSize || 11);
			const value = String(Math.max(1, Math.min(400, current + delta)));
			for (const option of [...controls.size.options]) { if (option.dataset.step) { option.remove(); } }
			if (![...controls.size.options].some(option => option.value === value)) { const option = new targetWindow.Option(value, value); option.dataset.step = 'true'; controls.size.add(option); }
			controls.size.value = value; controls.size.dataset.lastSize = value; controls.size.dataset.currentSize = value;
			controls.size.dispatchEvent(new targetWindow.Event('change', { bubbles: true }));
		}, listen);
	}
	controls.size.addEventListener('change', () => { if (controls.size.value) { controls.size.dataset.lastSize = controls.size.value; controls.size.dataset.currentSize = controls.size.value; } }, { ...listen, capture: true });
	const syncFont = () => {
		const fontName = controls.font.dataset.currentFont ?? controls.font.dataset.lastFont ?? 'Font';
		dropdowns[0].value.textContent = fontName || 'Mixed fonts';
		dropdowns[0].button.title = `Font Family · ${fontName || 'Mixed fonts'}`;
		dropdowns[1].value.textContent = controls.size.dataset.currentSize !== undefined ? controls.size.dataset.currentSize || '—' : controls.size.dataset.lastSize || '11';
		const numericSize = Number(controls.size.dataset.currentSize || controls.size.dataset.lastSize || 11);
		decrease.disabled = controls.size.disabled || numericSize <= 1; increase.disabled = controls.size.disabled || numericSize >= 400;
	};
	controls.font.addEventListener('change', () => { if (controls.font.value) { controls.font.dataset.lastFont = controls.font.value; controls.font.dataset.currentFont = controls.font.value; } syncFont(); }, { ...listen, capture: true });
	controls.font.addEventListener('formatstate', syncFont, listen);
	controls.size.addEventListener('change', syncFont, listen);
	controls.size.addEventListener('formatstate', syncFont, listen);
	const fontObserver = new targetWindow.MutationObserver(syncFont); fontObserver.observe(controls.size, { attributes: true, attributeFilter: ['disabled'] });
	syncFont();
	const history = group('Document', [iconButton(controls.undo, 'undo'), iconButton(controls.redo, 'redo')]);
	// Visual affordance only; its command mapping follows the snapshot work.
	const paintButton = iconButton(ownerDocument.createElement('button'), 'paintFormat');
	paintButton.type = 'button'; paintButton.title = 'Paint Format'; paintButton.setAttribute('aria-label', 'Paint Format');
	paintButton.classList.add('word-paint-button');
	paintButton.addEventListener('pointerdown', event => event.preventDefault());
	const paintGroup = group('Formatting', [paintButton]);
	const style = group('Styles', [menus.style]);
	const font = group('Font', [dropdowns[0].element, decrease, dropdowns[1].element, increase]);
	const emphasis = group('Text Formatting', [...controls.emphasis, ...(controls.strike ? [controls.strike] : [])]);
	if (controls.strike) { controls.strike.classList.add('word-emphasis-button'); controls.strike.style.textDecoration = 'line-through'; }
	const colors = [controls.color, controls.highlight].filter((item): item is HTMLSelectElement => !!item).map(select => { select.dataset.palette = 'true'; return menus.selector(select, icon(select === controls.color ? 'color' : 'highlight', ownerDocument)); });
	dropdowns.push(...colors);
	const colorsGroup = group('Colors', colors.map(item => item.element));
	const clearGroup = controls.clear ? group('Reset', [iconButton(controls.clear, 'clear')]) : undefined;
	const paragraph = group('Paragraph', [dropdowns[3].element, dropdowns[2].element, dropdowns[4].element, ...(controls.properties ? [iconButton(controls.properties, 'properties')] : [])]);
	const commentButton = iconButton(controls.comment, 'comment'); commentButton.classList.add('word-comment-badge');
	const insert = group('Insert', [...(controls.text ? [controls.text] : []), iconButton(controls.image, 'image'), iconButton(controls.table, 'table'), iconButton(controls.link, 'link'), commentButton]);
	const groups: HTMLElement[] = [history, paintGroup, style, font, emphasis, colorsGroup, paragraph, insert];
	if (clearGroup) { groups.splice(6, 0, clearGroup); }
	const overflow = ownerDocument.createElement('div');
	overflow.id = 'word-toolbar-overflow';
	overflow.className = 'word-toolbar-overflow';
	overflow.popover = 'auto';
	overflow.setAttribute('role', 'group');
	overflow.setAttribute('aria-label', 'More Formatting');
	more.setAttribute('aria-controls', overflow.id);
	more.popoverTargetElement = overflow;
	const extras = ownerDocument.createElement('div');
	extras.className = 'word-toolbar-group word-overflow-extras';
	extras.setAttribute('role', 'group');
	extras.setAttribute('aria-label', 'Extended formatting');
	// Placeholders for menus that follow: clicking one must not collapse the popup.
	extras.setAttribute('data-word-menu', '');
	const extraButton = (text: string, label: string, dropdown: boolean) => {
		const button = ownerDocument.createElement('button');
		button.type = 'button'; button.textContent = text; button.title = label; button.setAttribute('aria-label', label);
		button.classList.add('word-extra');
		if (dropdown) { button.classList.add('word-dropdown-button'); }
		return button;
	};
	extras.append(
		extraButton('A²', 'Superscript', false),
		extraButton('A₂', 'Subscript', false),
		extraButton('Ag', 'Small Caps', true),
		extraButton('Ab', 'Letter Spacing', true),
		extraButton('¶', 'Paragraph Mark', false),
		extraButton('⇥', 'Indent', true),
		extraButton('Aa', 'Line Spacing Options', true),
		extraButton('”', 'Block Quote', true),
	);
	overflow.append(extras);
	const headerFooter = [controls.margins, controls.pageBorders, controls.header, controls.footer].filter((control): control is HTMLButtonElement => !!control);
	if (headerFooter.length) { overflow.append(group('Page Layout', headerFooter)); }
	const allControls = [...headerFooter, controls.undo, controls.redo, decrease, increase, controls.properties, controls.text, menus.styleButton, ...dropdowns.map(item => item.button), ...controls.emphasis, controls.strike, controls.clear, controls.table, controls.image, controls.link, controls.comment, more].filter((control): control is HTMLButtonElement => !!control);
	const suggestions = ownerDocument.createElement('button');
	suggestions.type = 'button';
	suggestions.title = 'Suggestions';
	suggestions.setAttribute('aria-label', 'Suggestions');
	suggestions.className = 'word-suggestions-chip';
	suggestions.addEventListener('pointerdown', event => event.preventDefault());
	suggestions.replaceChildren(icon('comment', ownerDocument), ownerDocument.createTextNode('Suggestions'));
	row.append(...groups, more, suggestions);
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
		// Paint format and the suggestions chip stay pinned to the row.
		for (const element of [insert, paragraph, clearGroup, colorsGroup, emphasis, style, history, font]) {
			if (row.scrollWidth <= row.clientWidth) { break; }
			if (!element) { continue; }
			overflow.appendChild(element);
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
		observer.disconnect(); fontObserver.disconnect(); events.abort(); disposeFeedback(); menus.dispose(); overflow.remove();
		toolbar.classList.remove('word-toolbar');
	};
}
