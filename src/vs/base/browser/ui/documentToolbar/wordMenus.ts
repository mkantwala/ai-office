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
	styleButton.textContent = 'Aa';
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
		menu.addEventListener('keydown', event => { if (event.key === 'Escape') { event.preventDefault(); event.stopPropagation(); close(true); } else if (event.key === 'Tab') { close(); } }, listen);
		ownerDocument.addEventListener('pointerdown', event => { if (event.target instanceof targetWindow.Node && !menu.contains(event.target) && !button.contains(event.target)) { close(); } }, listen);
		ownerDocument.addEventListener('focusin', event => { if (event.target instanceof targetWindow.Node && !menu.contains(event.target) && !button.contains(event.target)) { close(); } }, listen);
		targetWindow.addEventListener('resize', () => close(), listen);
		ownerDocument.addEventListener('scroll', event => { if (event.target instanceof targetWindow.Node && !menu.contains(event.target)) { close(); } }, { ...listen, capture: true });
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
	styles.menu.setAttribute('role', 'menu'); styleButton.setAttribute('aria-haspopup', 'menu');
	select.addEventListener('formatstate', () => { styleButton.disabled = select.disabled;
		const current = select.dataset.currentStyle;
		styleButton.textContent = 'Aa';
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

	// Button menus share the original select's command and disabled state.
	const selectors: { select: HTMLSelectElement; button: HTMLButtonElement }[] = [];
	const selector = (select: HTMLSelectElement, content: string | SVGSVGElement) => {
		const palette = !!select.dataset.palette;
		const textColor = palette && Array.from(select.options).some(option => option.value === 'auto');
		if (textColor) {
			const rows = [
				'000000 434343 666666 999999 B7B7B7 CCCCCC D9D9D9 EFEFEF F3F3F3 FFFFFF',
				'980000 FF0000 FF9900 FFFF00 00FF00 00FFFF 4A86E8 0000FF 9900FF FF00FF',
				'E6B8AF F4CCCC FCE5CD FFF2CC D9EAD3 D0E0E3 C9DAF8 CFE2F3 D9D2E9 EAD1DC',
				'DD7E6B EA9999 F9CB9C FFE599 B6D7A8 A2C4C9 A4C2F4 9FC5E8 B4A7D6 D5A6BD',
				'CC4125 E06666 F6B26B FFD966 93C47D 76A5AF 6D9EEB 6FA8DC 8E7CC3 C27BA0',
				'A61C00 CC0000 E69138 F1C232 6AA84F 45818E 3C78D8 3D85C6 674EA7 A64D79',
				'85200C 990000 B45F06 BF9000 38761D 134F5C 1155CC 0B5394 351C75 741B47',
				'5B0F00 660000 783F04 7F6000 274E13 0C343D 1C4587 073763 20124D 4C1130',
			];
			const previous = Array.from(select.options);
			const values = rows.join(' ').split(' ');
			select.replaceChildren(...previous.filter(option => !option.value || option.value === 'auto'));
			for (const value of values) { select.add(new targetWindow.Option(previous.find(option => option.value === value)?.textContent || `#${value}`, value)); }
			for (const option of previous) { if (option.value && option.value !== 'auto' && !values.includes(option.value)) { select.add(option); } }
		}
		const button = ownerDocument.createElement('button'); button.type = 'button'; button.className = 'word-dropdown-button';
		button.title = select.title; button.setAttribute('aria-label', select.getAttribute('aria-label') || select.title);
		const value = ownerDocument.createElement('span'); value.className = 'word-control-value';
		if (typeof content === 'string') { value.textContent = content; button.append(value); } else { button.append(content); }
		const wrapper = ownerDocument.createElement('span'); wrapper.className = 'word-button-control'; select.hidden = true; wrapper.append(button, select);
		const items: HTMLButtonElement[] = [];
		const current = () => select.dataset.currentFont ?? select.dataset.currentSize ?? select.dataset.currentValue;
		const popup = create(button, `word-option-menu-${selectors.length}`, select.title, () => items.find(item => item.dataset.value === current()) || items[0]); popup.menu.classList.add('word-option-menu'); popup.menu.setAttribute('role', 'menu'); button.setAttribute('aria-haspopup', 'menu');
		if (palette) { popup.menu.classList.add('word-color-menu'); }
		for (const option of [...select.options].filter(option => option.value)) {
			const item = ownerDocument.createElement('button'); item.type = 'button'; item.textContent = option.textContent; item.setAttribute('role', typeof content === 'string' ? 'menuitemradio' : 'menuitem'); item.dataset.value = option.value; item.setAttribute('aria-label', option.textContent || option.value);
			if (typeof content === 'string') { item.setAttribute('aria-checked', String(option.value === current())); }
			if (select.classList.contains('word-font-select')) { item.style.fontFamily = option.value; }
			if (select.dataset.palette) {
				item.title = option.textContent || option.value;
				item.textContent = '';
				if (option.value === 'auto' || option.value === 'none') { item.textContent = option.textContent; item.classList.add('word-color-reset'); }
				const swatch = ownerDocument.createElement('span'); swatch.className = 'word-color-swatch';
				const colors: Record<string, string> = { yellow: '#ffff00', brightGreen: '#00ff00', turquoise: '#00ffff', pink: '#ff00ff', blue: '#0000ff', red: '#ff0000', darkBlue: '#000080', teal: '#008080', green: '#008000', violet: '#800080', darkRed: '#800000', darkYellow: '#808000', gray50: '#808080', gray25: '#c0c0c0', black: '#000000' };
				swatch.style.backgroundColor = /^[0-9a-f]{6}$/i.test(option.value) ? `#${option.value}` : colors[option.value] || 'transparent'; item.prepend(swatch);
			}
			item.addEventListener('click', () => { popup.close(); select.value = option.value; select.dispatchEvent(new targetWindow.Event('change', { bubbles: true })); focusDocument(); }, listen);
			items.push(item); popup.menu.append(item);
		}
		if (textColor) {
			const custom = ownerDocument.createElement('label'); custom.className = 'word-custom-color'; custom.textContent = 'Custom';
			const input = ownerDocument.createElement('input'); input.type = 'color'; input.setAttribute('aria-label', 'Custom Text Color'); input.title = 'Choose a custom text color';
			const applyColor = (hex: string) => {
				if (!/^#[0-9a-f]{6}$/i.test(hex)) { return; }
				const value = hex.slice(1).toUpperCase();
				if (!Array.from(select.options).some(option => option.value === value)) { select.add(new targetWindow.Option(`#${value}`, value)); }
				popup.close(); select.value = value; select.dispatchEvent(new targetWindow.Event('change', { bubbles: true })); focusDocument();
			};
			input.addEventListener('change', () => applyColor(input.value), listen);
			custom.append(input); popup.menu.append(custom);
			const EyeDropper = (targetWindow as Window & { EyeDropper?: new () => { open(options: { signal: AbortSignal }): Promise<{ sRGBHex: string }> } }).EyeDropper;
			if (EyeDropper) {
				const dropper = ownerDocument.createElement('button'); dropper.type = 'button'; dropper.className = 'word-color-reset'; dropper.textContent = 'Pick Screen Color';
				dropper.addEventListener('click', async () => {
					try { applyColor((await new EyeDropper().open({ signal: events.signal })).sRGBHex); } catch { /* Canceling a screen color selection leaves formatting unchanged. */ }
				}, listen);
				popup.menu.append(dropper);
			}
		}
		const syncSelection = () => { for (const item of items) { if (item.getAttribute('role') === 'menuitemradio') { item.setAttribute('aria-checked', String(item.dataset.value === current())); } } };
		button.addEventListener('click', syncSelection, listen);
		select.addEventListener('formatstate', syncSelection, listen); select.addEventListener('change', () => { select.dataset.currentValue = select.value; syncSelection(); }, { ...listen, capture: true });
		popup.menu.addEventListener('keydown', event => {
			if (event.target instanceof targetWindow.HTMLInputElement || !['ArrowDown', 'ArrowUp', 'ArrowLeft', 'ArrowRight', 'Home', 'End'].includes(event.key)) { return; }
			event.preventDefault(); const index = items.indexOf(ownerDocument.activeElement as HTMLButtonElement);
			const step = palette && ['ArrowDown', 'ArrowUp'].includes(event.key) ? 10 : 1;
			items[event.key === 'Home' ? 0 : event.key === 'End' ? items.length - 1 : (index + (['ArrowDown', 'ArrowRight'].includes(event.key) ? step : -step) + items.length) % items.length]?.focus();
		}, listen);
		selectors.push({ select, button }); button.disabled = select.disabled; observer.observe(select, { attributes: true, attributeFilter: ['disabled'] });
		return { element: wrapper, button, value };
	};
	const update = () => { for (const entry of selectors) { entry.button.disabled = entry.select.disabled; } styleButton.disabled = select.disabled; };
	const observer = new targetWindow.MutationObserver(update);
	observer.observe(select, { attributes: true, attributeFilter: ['disabled'] });

	highlight(1, 1); tables.menu.append(heading, grid, size);
	return { style: wrapper, styleButton, selector, dispose: () => { observer.disconnect(); events.abort(); menus.forEach(menu => menu.remove()); } };
}
