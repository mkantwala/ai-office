/*---------------------------------------------------------------------------------------------
 *  Copyright (c) Microsoft Corporation. All rights reserved.
 *  Licensed under the MIT License. See License.txt in the project root for license information.
 *--------------------------------------------------------------------------------------------*/

/** Word-only visual menus. Commands still flow through the existing engine controls. */
export function wordMenus(select: HTMLSelectElement, table: HTMLButtonElement, insert: (rows: number, cols: number) => void, focusDocument: () => void, host: HTMLElement = select.ownerDocument.body, focusOverflow: () => void = focusDocument) {
	const ownerDocument = select.ownerDocument;
	const targetWindow = ownerDocument.defaultView!;
	const events = new AbortController();
	const listen = { signal: events.signal };
	const menus: HTMLElement[] = [];
	const styleButton = ownerDocument.createElement('button');
	styleButton.type = 'button'; styleButton.disabled = select.disabled; styleButton.className = 'word-style-select';
	styleButton.title = 'Paragraph Style'; styleButton.setAttribute('aria-label', 'Paragraph Style');
	styleButton.textContent = 'Paragraph style';
	select.hidden = true;
	const wrapper = ownerDocument.createElement('span'); wrapper.className = 'word-style-control'; wrapper.append(styleButton, select);
	const create = (button: HTMLButtonElement, id: string, label: string, initialFocus: () => HTMLElement | undefined) => {
		button.dataset.wordMenu = 'true'; button.setAttribute('aria-haspopup', 'dialog'); button.setAttribute('aria-expanded', 'false'); button.setAttribute('aria-controls', id);
		const menu = ownerDocument.createElement('div'); menu.id = id; menu.className = 'word-visual-menu'; menu.popover = 'manual';
		menu.setAttribute('role', 'dialog'); menu.setAttribute('aria-label', label); host.append(menu); menus.push(menu);
		const close = (restore = false) => { if (menu.matches(':popover-open')) { menu.hidePopover(); } button.setAttribute('aria-expanded', 'false'); if (restore) { const overflow = button.closest<HTMLElement>('.word-toolbar-overflow'); if (overflow && !overflow.matches(':popover-open')) { focusOverflow(); } else { button.focus(); } } };
		button.addEventListener('click', () => {
			if (menu.matches(':popover-open')) { close(true); return; }
			for (const other of menus) { if (other.matches(':popover-open')) { other.hidePopover(); } }
			const rect = button.getBoundingClientRect();
			const overflow = button.closest<HTMLElement>('.word-toolbar-overflow'); if (overflow?.matches(':popover-open')) { overflow.hidePopover(); }
			menu.showPopover(); button.setAttribute('aria-expanded', 'true');
			menu.style.left = `${Math.max(8, Math.min(rect.left, targetWindow.innerWidth - menu.offsetWidth - 8))}px`;
			menu.style.top = `${Math.max(8, Math.min(rect.bottom + 6, targetWindow.innerHeight - menu.offsetHeight - 8))}px`;
			initialFocus()?.focus({ preventScroll: true });
		}, listen);
		menu.addEventListener('keydown', event => { if (event.key === 'Escape') { event.preventDefault(); event.stopPropagation(); close(true); } }, listen);
		ownerDocument.addEventListener('pointerdown', event => { if (event.target instanceof targetWindow.Node && !menu.contains(event.target) && !button.contains(event.target)) { close(); } }, listen);
		ownerDocument.addEventListener('focusin', event => { if (event.target instanceof targetWindow.Node && !menu.contains(event.target) && !button.contains(event.target)) { close(); } }, listen);
		targetWindow.addEventListener('resize', () => close(), listen);
		menu.addEventListener('toggle', () => button.setAttribute('aria-expanded', String(menu.matches(':popover-open'))), listen);
		return { menu, close };
	};
	const styles = create(styleButton, 'word-style-menu', 'Paragraph styles', () => styleItems[0]);
	const styleItems: HTMLButtonElement[] = [];
	for (const option of [...select.options].filter(option => option.value)) {
		const item = ownerDocument.createElement('button'); item.type = 'button'; item.className = 'word-style-preview'; item.dataset.style = option.value;
		item.setAttribute('role', 'menuitemradio'); item.setAttribute('aria-checked', 'false');
		const check = ownerDocument.createElement('span'); check.className = 'word-style-check'; check.setAttribute('aria-hidden', 'true'); check.textContent = '✓';
		const sample = ownerDocument.createElement('span'); sample.textContent = option.textContent;
		item.append(check, sample); styles.menu.append(item); styleItems.push(item);
		item.addEventListener('click', () => { styles.close(); select.value = option.value; select.dispatchEvent(new Event('change', { bubbles: true })); focusDocument(); }, listen);
	}
	styles.menu.setAttribute('role', 'menu');
	select.addEventListener('formatstate', () => { styleButton.disabled = select.disabled;
		const current = select.dataset.currentStyle;
		styleButton.textContent = [...select.options].find(option => option.value === current)?.textContent || 'Paragraph style';
		for (const item of styleItems) { item.setAttribute('aria-checked', String(item.dataset.style === current)); }
	}, listen);
	styles.menu.addEventListener('keydown', event => {
		if (!['ArrowDown', 'ArrowUp', 'Home', 'End'].includes(event.key)) { return; }
		event.preventDefault(); const index = styleItems.indexOf(ownerDocument.activeElement as HTMLButtonElement);
		styleItems[event.key === 'Home' ? 0 : event.key === 'End' ? styleItems.length - 1 : (index + (event.key === 'ArrowDown' ? 1 : -1) + styleItems.length) % styleItems.length]?.focus();
	}, listen);
	const tables = create(table, 'word-table-menu', 'Insert table', () => cells[0]);
	const heading = ownerDocument.createElement('div'); heading.className = 'word-menu-heading'; heading.textContent = 'Insert table';
	const grid = ownerDocument.createElement('div'); grid.className = 'word-table-grid'; grid.setAttribute('role', 'group'); grid.setAttribute('aria-label', 'Choose table size');
	const size = ownerDocument.createElement('div'); size.className = 'word-table-size'; size.setAttribute('aria-live', 'polite');
	const cells: HTMLButtonElement[] = [];
	const highlight = (row: number, col: number) => {
		size.textContent = `${col} × ${row}`;
		cells.forEach((cell, index) => cell.classList.toggle('selected', Math.floor(index / 10) < row && index % 10 < col));
	};
	for (let row = 1; row <= 8; row++) {
		for (let col = 1; col <= 10; col++) {
			const cell = ownerDocument.createElement('button'); cell.type = 'button'; cell.className = 'word-table-cell'; cell.tabIndex = row === 1 && col === 1 ? 0 : -1;
			cell.setAttribute('aria-label', `${col} columns, ${row} rows`);
			cell.addEventListener('pointerenter', () => highlight(row, col), listen);
			cell.addEventListener('focus', () => { cells.forEach(other => other.tabIndex = other === cell ? 0 : -1); highlight(row, col); }, listen);
			cell.addEventListener('click', () => { tables.close(); insert(row, col); focusDocument(); }, listen);
			cells.push(cell); grid.append(cell);
		}
	}
	grid.addEventListener('keydown', event => {
		if (!['ArrowLeft', 'ArrowRight', 'ArrowUp', 'ArrowDown', 'Home', 'End'].includes(event.key)) { return; }
		event.preventDefault();
		const index = Math.max(0, cells.indexOf(ownerDocument.activeElement as HTMLButtonElement));
		const row = Math.max(1, Math.min(8, Math.floor(index / 10) + 1 + (event.key === 'ArrowDown' ? 1 : event.key === 'ArrowUp' ? -1 : 0)));
		const col = event.key === 'Home' ? 1 : event.key === 'End' ? 10 : Math.max(1, Math.min(10, index % 10 + 1 + (event.key === 'ArrowRight' ? 1 : event.key === 'ArrowLeft' ? -1 : 0)));
		cells[(row - 1) * 10 + col - 1].focus();
	}, listen);

	highlight(1, 1); tables.menu.append(heading, grid, size);
	return { style: wrapper, styleButton, dispose: () => { events.abort(); menus.forEach(menu => menu.remove()); } };
}
