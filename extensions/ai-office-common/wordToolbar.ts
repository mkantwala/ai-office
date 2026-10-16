/*---------------------------------------------------------------------------------------------
 *  Copyright (c) Microsoft Corporation. All rights reserved.
 *  Licensed under the MIT License. See License.txt in the project root for license information.
 *--------------------------------------------------------------------------------------------*/

interface WordToolbarControls {
	undo: HTMLButtonElement;
	redo: HTMLButtonElement;
	save: HTMLButtonElement;
	style: HTMLSelectElement;
	font: HTMLSelectElement;
	size: HTMLSelectElement;
	emphasis: HTMLButtonElement[];
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

function icon(name: keyof typeof paths): SVGSVGElement {
	const svg = document.createElementNS('http://www.w3.org/2000/svg', 'svg');
	svg.setAttribute('viewBox', '0 0 24 24');
	svg.setAttribute('aria-hidden', 'true');
	svg.setAttribute('focusable', 'false');
	const path = document.createElementNS(svg.namespaceURI, 'path');
	path.setAttribute('d', paths[name]);
	svg.append(path);
	return svg;
}

function iconButton(button: HTMLButtonElement, name: keyof typeof paths): HTMLButtonElement {
	const label = document.createElement('span');
	label.className = 'word-action-label';
	label.textContent = button.textContent;
	button.classList.add('word-icon-button');
	button.replaceChildren(icon(name), label);
	return button;
}

function iconSelect(select: HTMLSelectElement, name: 'align' | 'list' | 'spacing'): HTMLElement {
	const wrapper = document.createElement('span');
	wrapper.className = 'word-icon-select';
	const label = document.createElement('span');
	label.className = 'word-action-label';
	label.textContent = select.title;
	// The real select stays focusable and retains native keyboard/menu behavior.
	wrapper.append(icon(name), label, select);
	return wrapper;
}

/** Reorganize the existing controls without duplicating document commands or selection state. */
export function arrangeWordToolbar(toolbar: HTMLElement, controls: WordToolbarControls): () => void {
	const events = new AbortController();
	const listen = { signal: events.signal };
	toolbar.classList.add('word-toolbar');
	document.body.classList.add('word-document');
	toolbar.setAttribute('role', 'toolbar');
	toolbar.setAttribute('aria-orientation', 'horizontal');
	const row = document.createElement('div');
	row.className = 'word-toolbar-row';
	const group = (label: string, children: HTMLElement[]) => {
		const element = document.createElement('div');
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
		controls.emphasis[index].title = label;
		controls.emphasis[index].setAttribute('aria-label', label);
		controls.emphasis[index].classList.add('word-emphasis-button');
	}
	const history = group('Document', [iconButton(controls.undo, 'undo'), iconButton(controls.redo, 'redo'), iconButton(controls.save, 'save')]);
	const style = group('Styles', [controls.style]);
	const font = group('Font', [controls.font, controls.size]);
	const emphasis = group('Text Formatting', controls.emphasis);
	const paragraph = group('Paragraph', [iconSelect(controls.align, 'align'), iconSelect(controls.spacing, 'spacing'), iconSelect(controls.list, 'list')]);
	const insert = group('Insert', [iconButton(controls.link, 'link'), iconButton(controls.comment, 'comment'), iconButton(controls.image, 'image')]);
	const table = group('Table', [iconButton(controls.table, 'table')]);
	const groups = [history, style, font, emphasis, paragraph, insert];
	const more = document.createElement('button');
	more.type = 'button';
	more.textContent = 'More';
	more.title = 'More Formatting';
	more.setAttribute('aria-label', 'More Formatting');
	more.setAttribute('aria-expanded', 'false');
	iconButton(more, 'more');
	more.classList.add('word-more-button');
	const overflow = document.createElement('div');
	overflow.id = 'word-toolbar-overflow';
	overflow.className = 'word-toolbar-overflow';
	overflow.popover = 'auto';
	overflow.setAttribute('role', 'group');
	overflow.setAttribute('aria-label', 'More Formatting');
	more.setAttribute('aria-controls', overflow.id);
	more.popoverTargetElement = overflow;
	overflow.append(table);
	row.append(...groups, more);
	toolbar.replaceChildren(row);
	document.body.append(overflow);

	const positionOverflow = () => {
		const rect = more.getBoundingClientRect();
		overflow.style.left = `${Math.max(8, Math.min(rect.right - overflow.offsetWidth, innerWidth - overflow.offsetWidth - 8))}px`;
		overflow.style.top = `${rect.bottom + 8}px`;
		overflow.style.maxHeight = `${Math.max(80, innerHeight - rect.bottom - 16)}px`;
	};
	overflow.addEventListener('toggle', () => {
		const open = overflow.matches(':popover-open');
		more.setAttribute('aria-expanded', String(open));
		if (open) {
			positionOverflow();
			overflow.querySelector<HTMLElement>('button:not(:disabled), select:not(:disabled)')?.focus({ preventScroll: true });
		}
	}, listen);
	// Hide before a command opens a modal prompt, retaining the engine selection.
	overflow.addEventListener('click', event => {
		if (event.target instanceof Element && event.target.closest('button')) { overflow.hidePopover(); }
	}, { ...listen, capture: true });
	overflow.addEventListener('change', () => overflow.hidePopover(), listen);
	overflow.addEventListener('keydown', event => {
		if (event.key === 'Escape') {
			event.preventDefault(); event.stopPropagation(); overflow.hidePopover(); more.focus();
		}
	}, listen);
	row.addEventListener('keydown', event => {
		if (!(event.target instanceof HTMLButtonElement) || !['ArrowLeft', 'ArrowRight', 'Home', 'End'].includes(event.key)) { return; }
		const items = [...row.querySelectorAll<HTMLElement>('button:not(:disabled), select:not(:disabled)')];
		const index = items.indexOf(event.target);
		const next = event.key === 'Home' ? 0 : event.key === 'End' ? items.length - 1 : (index + (event.key === 'ArrowRight' ? 1 : -1) + items.length) % items.length;
		event.preventDefault(); items[next]?.focus();
	}, listen);
	const layout = () => {
		if (overflow.matches(':popover-open')) { overflow.hidePopover(); }
		const focused = document.activeElement;
		for (const element of groups) { row.insertBefore(element, more); }
		// Move real controls, preserving their listeners, disabled state and values.
		for (const element of [insert, paragraph, font, style, emphasis]) {
			if (row.scrollWidth <= row.clientWidth) { break; }
			overflow.insertBefore(element, table);
		}
		if (focused instanceof HTMLElement && groups.some(element => element.contains(focused))) {
			if (overflow.contains(focused)) { more.focus({ preventScroll: true }); }
			else { focused.focus({ preventScroll: true }); }
		}
	};
	const observer = new ResizeObserver(layout);
	observer.observe(row);
	layout();
	return () => {
		observer.disconnect(); events.abort(); overflow.remove();
		toolbar.classList.remove('word-toolbar'); document.body.classList.remove('word-document');
	};
}
