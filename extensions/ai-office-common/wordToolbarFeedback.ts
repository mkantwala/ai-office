/*---------------------------------------------------------------------------------------------
 *  Copyright (c) Microsoft Corporation. All rights reserved.
 *  Licensed under the MIT License. See License.txt in the project root for license information.
 *--------------------------------------------------------------------------------------------*/

/** One delegated tooltip for all toolbar/popover controls; no per-control timers or dependencies. */
export function wordToolbarFeedback(): () => void {
	const events = new AbortController(), listen = { signal: events.signal };
	const tooltip = document.createElement('div'); tooltip.id = 'word-control-tooltip'; tooltip.className = 'word-control-tooltip'; tooltip.popover = 'manual'; tooltip.setAttribute('role', 'tooltip'); document.body.append(tooltip);
	let target: HTMLElement | undefined;
	let timer: ReturnType<typeof setTimeout> | undefined;
	const titles = new Map<HTMLElement, string>();
	const control = (node: EventTarget | null) => node instanceof Element ? node.closest<HTMLElement>('.word-toolbar button,.word-toolbar select,.word-toolbar-overflow button,.word-toolbar-overflow select,.word-visual-menu button') ?? undefined : undefined;
	const hide = () => {
		clearTimeout(timer); timer = undefined;
		if (tooltip.matches(':popover-open')) { tooltip.hidePopover(); }
		if (target) {
			const ids = (target.getAttribute('aria-describedby') || '').split(' ').filter(id => id && id !== tooltip.id);
			if (ids.length) { target.setAttribute('aria-describedby', ids.join(' ')); } else { target.removeAttribute('aria-describedby'); }
		}
		target = undefined;
	};
	const show = (next: HTMLElement | undefined, delay: number) => {
		if (next === target) { return; } hide();
		if (!next) { return; }
		const title = next.getAttribute('title');
		if (title) { titles.set(next, title); next.removeAttribute('title'); }
		const label = titles.get(next) || next.getAttribute('aria-label');
		if (!label) { return; } target = next;
		timer = setTimeout(() => {
			if (!next.isConnected || next !== target) { hide(); return; }
			tooltip.textContent = label; tooltip.showPopover();
			const box = next.getBoundingClientRect(), tip = tooltip.getBoundingClientRect();
			tooltip.style.left = `${Math.max(8, Math.min(box.left + (box.width - tip.width) / 2, innerWidth - tip.width - 8))}px`;
			tooltip.style.top = `${box.bottom + tip.height + 14 < innerHeight ? box.bottom + 7 : Math.max(8, box.top - tip.height - 7)}px`;
			next.setAttribute('aria-describedby', [next.getAttribute('aria-describedby'), tooltip.id].filter(Boolean).join(' '));
		}, delay);
	};
	document.addEventListener('pointerover', event => { if (event.pointerType !== 'touch') { show(control(event.target), 350); } }, listen);
	document.addEventListener('pointerout', event => { if (control(event.relatedTarget) !== target) { hide(); } }, listen);
	document.addEventListener('focusin', event => show(control(event.target), 100), listen);
	document.addEventListener('focusout', hide, listen);
	document.addEventListener('pointerdown', hide, listen);
	document.addEventListener('keydown', event => { if (event.key === 'Escape') { hide(); } }, listen);
	document.addEventListener('scroll', hide, { ...listen, capture: true });
	window.addEventListener('resize', hide, listen);
	window.addEventListener('blur', hide, listen);
	return () => { hide(); events.abort(); tooltip.remove(); for (const [element, title] of titles) { element.title = title; } titles.clear(); };
}
