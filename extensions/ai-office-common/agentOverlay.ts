/*---------------------------------------------------------------------------------------------
 *  Copyright (c) Microsoft Corporation. All rights reserved.
 *  Licensed under the MIT License. See License.txt in the project root for license information.
 *--------------------------------------------------------------------------------------------*/

import type { State } from './types.js';

type Place = (node: HTMLElement, page: number, x: number, y: number, width: number, height: number) => void;

/** Screen blending lights the canvas glyphs while leaving the white document paper unchanged. */
export function renderAgentActivity(pages: HTMLElement, rects: NonNullable<State['rects']>, place: Place, label = 'Agent', intent?: string) {
	pages.querySelectorAll('.agent-shimmer').forEach(node => node.remove());
	const bounds = new Map<number, { left: number; top: number; right: number; bottom: number }>();
	for (const rect of rects.slice(0, 200)) {
		if (!Number.isInteger(rect.page) || rect.page < 0 || !pages.children[rect.page] || ![rect.x, rect.y, rect.width, rect.height].every(Number.isFinite) || rect.width <= 0 || rect.height <= 0) { continue; }
		const node = document.createElement('div'); node.className = intent ? 'agent-shimmer agent-preview-anchor' : 'agent-shimmer agent-shimmer-text'; node.setAttribute('aria-hidden', 'true');
		place(node, rect.page, rect.x, rect.y, rect.width, rect.height);
		// Use placed coordinates so the outline includes any inline review view zones.
		const left = parseFloat(node.style.left), top = parseFloat(node.style.top);
		const right = left + parseFloat(node.style.width), bottom = top + parseFloat(node.style.height);
		const previous = bounds.get(rect.page);
		bounds.set(rect.page, previous ? { left: Math.min(left, previous.left), top: Math.min(top, previous.top), right: Math.max(right, previous.right), bottom: Math.max(bottom, previous.bottom) } : { left, top, right, bottom });
	}
	for (const [page, box] of bounds) {
		if (intent) {
			const preview = document.createElement('div'); preview.className = 'agent-shimmer agent-edit-preview'; preview.setAttribute('aria-hidden', 'true');
			Object.assign(preview.style, { left: `${box.left}px`, top: `${box.top}px`, width: `${box.right - box.left}px`, height: `${box.bottom - box.top}px` });
			const text = document.createElement('p'); text.textContent = '[' + intent.slice(0, 240) + ']';
			text.style.webkitLineClamp = String(Math.max(1, Math.floor((box.bottom - box.top) / 19)));
			preview.append(text); pages.children[page].append(preview);
		}
		const node = document.createElement('div'); node.className = 'agent-shimmer agent-shimmer-border'; node.setAttribute('aria-hidden', 'true');
		node.style.left = `${box.left - 5}px`; node.style.top = `${box.top - 4}px`;
		node.style.width = `${box.right - box.left + 10}px`; node.style.height = `${box.bottom - box.top + 8}px`;
		pages.children[page].append(node);
		const cursor = document.createElement('div'); cursor.className = 'agent-shimmer agent-cursor'; cursor.setAttribute('aria-hidden', 'true');
		cursor.style.left = `${box.left - 2}px`; cursor.style.top = `${box.top}px`; cursor.style.height = '20px';
		const tag = document.createElement('span'); tag.textContent = label.slice(0, 40); cursor.append(tag); pages.children[page].append(cursor);
	}
	return bounds.size > 0;
}

/** Reveal committed raster text quickly. This never types into or mutates the native document. */
export class WordAgentReveal {
	private pending?: { revision: number; rects: NonNullable<State['rects']> };
	private timer?: ReturnType<typeof setTimeout>;
	constructor(private readonly pages: HTMLElement, private readonly place: Place, private readonly revision: () => number) { }
	clear() { clearTimeout(this.timer); this.timer = undefined; this.pending = undefined; this.pages.querySelectorAll('.agent-reveal').forEach(node => node.remove()); }
	queue(revision: number, rects: NonNullable<State['rects']>) {
		this.clear();
		if (revision !== this.revision() || matchMedia('(prefers-reduced-motion: reduce)').matches) { return; }
		this.pending = { revision, rects: rects.slice(0, 80) }; this.paint();
		this.timer = setTimeout(() => this.clear(), 900);
	}
	paint() {
		const pending = this.pending; if (!pending) { return; }
		if (pending.revision !== this.revision()) { this.clear(); return; }
		const remaining: NonNullable<State['rects']> = [];
		for (const [index, rect] of pending.rects.entries()) {
			const page = this.pages.children[rect.page] as HTMLElement | undefined;
			if (!page || ![rect.x, rect.y, rect.width, rect.height].every(Number.isFinite) || rect.width <= 0 || rect.height <= 0) { continue; }
			const bounds = page.getBoundingClientRect(); if (bounds.bottom < 0 || bounds.top > innerHeight) { continue; }
			const canvas = page.querySelector('canvas');
			if (Number(canvas?.dataset.revision) !== pending.revision) { remaining.push(rect); continue; }
			const mask = document.createElement('div'); mask.className = 'agent-reveal'; mask.setAttribute('aria-hidden', 'true');
			this.place(mask, rect.page, rect.x, rect.y, rect.width, rect.height);
			mask.style.animationDelay = `${Math.min(index * 18, 240)}ms`;
			mask.addEventListener('animationend', () => mask.remove(), { once: true });
		}
		pending.rects = remaining;
	}
}
