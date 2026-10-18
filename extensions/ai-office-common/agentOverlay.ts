/*---------------------------------------------------------------------------------------------
 *  Copyright (c) Microsoft Corporation. All rights reserved.
 *  Licensed under the MIT License. See License.txt in the project root for license information.
 *--------------------------------------------------------------------------------------------*/

import type { State } from './types.js';

type Place = (node: HTMLElement, page: number, x: number, y: number, width: number, height: number) => void;

/** Screen blending lights the canvas glyphs while leaving the white document paper unchanged. */
export function renderAgentActivity(pages: HTMLElement, rects: NonNullable<State['rects']>, place: Place) {
	pages.querySelectorAll('.agent-shimmer').forEach(node => node.remove());
	const bounds = new Map<number, { left: number; top: number; right: number; bottom: number }>();
	for (const rect of rects) {
		if (!Number.isInteger(rect.page) || rect.page < 0 || !pages.children[rect.page] || ![rect.x, rect.y, rect.width, rect.height].every(Number.isFinite) || rect.width <= 0 || rect.height <= 0) { continue; }
		const node = document.createElement('div'); node.className = 'agent-shimmer agent-shimmer-text'; node.setAttribute('aria-hidden', 'true');
		place(node, rect.page, rect.x, rect.y, rect.width, rect.height);
		// Use placed coordinates so the outline includes any inline review view zones.
		const left = parseFloat(node.style.left), top = parseFloat(node.style.top);
		const right = left + parseFloat(node.style.width), bottom = top + parseFloat(node.style.height);
		const previous = bounds.get(rect.page);
		bounds.set(rect.page, previous ? { left: Math.min(left, previous.left), top: Math.min(top, previous.top), right: Math.max(right, previous.right), bottom: Math.max(bottom, previous.bottom) } : { left, top, right, bottom });
	}
	for (const [page, box] of bounds) {
		const node = document.createElement('div'); node.className = 'agent-shimmer agent-shimmer-border'; node.setAttribute('aria-hidden', 'true');
		node.style.left = `${box.left - 5}px`; node.style.top = `${box.top - 4}px`;
		node.style.width = `${box.right - box.left + 10}px`; node.style.height = `${box.bottom - box.top + 8}px`;
		pages.children[page].append(node);
	}
	return bounds.size > 0;
}
