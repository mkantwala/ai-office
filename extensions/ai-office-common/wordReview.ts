/*---------------------------------------------------------------------------------------------
 *  Copyright (c) Microsoft Corporation. All rights reserved.
 *  Licensed under the MIT License. See License.txt in the project root for license information.
 *--------------------------------------------------------------------------------------------*/

import type { DocumentReview, DocumentReviewChange, DraftViewSection } from './types.js';

interface Zone { at: number; height: number; node: HTMLElement }
/** Like code-editor view zones: insert deleted text without changing document layout or bytes. */
export class WordReview {
	private data?: DocumentReview;
	private drafts: DraftViewSection[] = [];
	private readonly zones = new Map<number, Zone[]>();
	private busy = false;
	constructor(private readonly pages: HTMLElement, private readonly decide: (data: DocumentReview, action: 'keep' | 'undo', ids: string[]) => void, private readonly onLayout: () => void, private readonly discardDraft: (id: string) => void) { }
	updateDrafts(sections: DraftViewSection[]) { this.drafts = sections; }
	update(data: DocumentReview) { this.data = data; this.busy = false; }
	get count() { return this.data?.changes.length || 0; }
	decideAll(action: 'keep' | 'undo') { if (this.data && !this.busy) { this.decide(this.data, action, this.data.changes.map(change => change.id)); } }
	toView(page: number, y: number) { return y + (this.zones.get(page) || []).filter(zone => zone.at <= y).reduce((sum, zone) => sum + zone.height, 0); }
	toDocument(page: number, y: number) {
		let offset = 0;
		for (const zone of this.zones.get(page) || []) {
			if (y < zone.at + offset) { break; }
			if (y < zone.at + offset + zone.height) { return zone.at; }
			offset += zone.height;
		}
		return y - offset;
	}
	private controls(change: DocumentReviewChange) {
		const row = document.createElement('div'); row.className = 'word-review-controls';
		const label = document.createElement('span'); label.textContent = change.label; row.append(label);
		for (const action of ['keep', 'undo'] as const) {
			const button = document.createElement('button'); button.textContent = action === 'keep' ? 'Keep' : 'Undo';
			button.setAttribute('aria-label', `${button.textContent} This Change`); button.disabled = this.busy;
			button.addEventListener('pointerdown', event => event.preventDefault());
			button.addEventListener('click', () => {
				if (this.busy || !this.data) { return; }
				this.busy = true;
				this.pages.querySelectorAll<HTMLButtonElement>('.word-review-controls button').forEach(item => item.disabled = true);
				this.decide(this.data, action, [change.id]);
			}); row.append(button);
		}
		return row;
	}
	/** Rebuild only after state/resize; raster updates reuse the measured view zones. */
	layout(dimensions: number[][]) {
		this.zones.clear();
		for (const [page, element] of Array.from(this.pages.children).entries()) {
			const node = element as HTMLElement;
			node.querySelectorAll('.word-review-zone,.word-review-insert,.word-review-slice').forEach(child => child.remove());
			node.classList.remove('has-word-review'); node.style.height = ''; node.style.aspectRatio = dimensions[page]?.join(' / ') || '';
			const source = node.querySelector('canvas'); if (source) { source.style.height = ''; }
		}
		if (!this.data?.changes.length && !this.drafts.length) { return; }
		for (const change of this.data?.changes || []) {
			const anchors = change.anchorRects?.length ? change.anchorRects : change.rects;
			const anchor = change.anchor?.after ? anchors?.at(-1) : anchors?.[0];
			const page = anchor?.page ?? 0, element = this.pages.children[page] as HTMLElement | undefined;
			if (!element || !dimensions[page]) { continue; }
			const scale = element.clientWidth / dimensions[page][0];
			const at = anchor ? (anchor.y + (change.anchor?.after ? anchor.height : 0)) * scale : 40 * scale;
			const zone = document.createElement('section'); zone.className = 'word-review-zone'; zone.setAttribute('aria-label', 'Proposed Document Change');
			zone.style.left = '0'; zone.style.width = '100%';
			zone.style.paddingInline = `${Math.max(16, (anchor?.x || 40) * scale)}px`;
			zone.addEventListener('pointerdown', event => event.stopPropagation());
			zone.append(this.controls(change));
			if (change.beforeText) {
				const deleted = document.createElement('div'); deleted.className = 'word-review-deleted';
				const marker = document.createElement('span'); marker.className = 'word-review-sign'; marker.textContent = '−'; marker.setAttribute('aria-hidden', 'true');
				const text = document.createElement('del'); text.textContent = change.beforeText; deleted.append(marker, text); zone.append(deleted);
			}
			// Formatting and structure remain inspectable even when text did not change.
			if (change.beforeText === change.afterText || change.id === 'document' || change.label === 'Table change') {
				const details = document.createElement('details'), summary = document.createElement('summary'); summary.textContent = 'Formatting and Structure Changes'; details.append(summary);
				for (const [value, className] of [[change.before, 'word-review-deleted'], [change.after, 'word-review-added']]) {
					const text = document.createElement('pre'); text.className = className; text.textContent = value; details.append(text);
				}
				details.addEventListener('toggle', () => { this.remeasure(page, dimensions); this.onLayout(); }); zone.append(details);
			}
			element.append(zone);
			const zones = this.zones.get(page) || []; zones.push({ at, height: zone.offsetHeight, node: zone }); this.zones.set(page, zones);
		}
		for (const section of this.drafts) {
			const anchor = section.anchor.position === 'after' ? section.anchorRects?.at(-1) : section.anchorRects?.[0];
			const page = anchor?.page ?? 0, element = this.pages.children[page] as HTMLElement | undefined;
			if (!element || !dimensions[page]) { continue; }
			const scale = element.clientWidth / dimensions[page][0];
			const at = anchor ? (anchor.y + (section.anchor.position === 'after' ? anchor.height : 0)) * scale : 40 * scale;
			const zone = document.createElement('section'); zone.className = 'word-review-zone word-draft-zone'; zone.setAttribute('aria-label', 'Pending Document Section');
			zone.style.left = '0'; zone.style.width = '100%'; zone.style.paddingInline = `${Math.max(16, (anchor?.x || 40) * scale)}px`;
			zone.addEventListener('pointerdown', event => event.stopPropagation());
			const label = document.createElement('small'); label.textContent = 'Pending section · Not saved in DOCX';
			const title = document.createElement('strong'); title.textContent = section.title;
			const intent = document.createElement('p'); intent.textContent = anchor ? section.intent : 'Anchor changed. Replan this section.';
			const discard = document.createElement('button'); discard.textContent = 'Discard'; discard.addEventListener('pointerdown', event => event.preventDefault()); discard.setAttribute('aria-label', `Discard pending section: ${section.title}`); discard.addEventListener('click', () => this.discardDraft(section.id));
			zone.append(label, title, intent, discard); element.append(zone);
			const zones = this.zones.get(page) || []; zones.push({ at, height: zone.offsetHeight, node: zone }); this.zones.set(page, zones);
		}
		for (const [page, zones] of this.zones) {
			zones.sort((a, b) => a.at - b.at); this.remeasure(page, dimensions);
		}
		for (const change of this.data?.changes || []) {
			if (!change.afterText) { continue; }
			for (const rect of change.rects || []) {
				const element = this.pages.children[rect.page] as HTMLElement, scale = element.clientWidth / dimensions[rect.page][0];
				const highlight = document.createElement('div'); highlight.className = 'word-review-insert'; highlight.setAttribute('aria-label', 'Added Content');
				highlight.dataset.documentY = String(rect.y * scale);
				highlight.style.cssText = `left:${rect.x * scale}px;top:${this.toView(rect.page, rect.y * scale)}px;width:${Math.max(20, rect.width * scale)}px;height:${rect.height * scale}px`;
				element.append(highlight);
			}
		}
	}
	private remeasure(page: number, dimensions: number[][]) {
		const element = this.pages.children[page] as HTMLElement, zones = this.zones.get(page) || [];
		const scale = element.clientWidth / dimensions[page][0], baseHeight = dimensions[page][1] * scale;
		let offset = 0;
		for (const zone of zones) { zone.height = zone.node.offsetHeight; zone.node.style.top = `${zone.at + offset}px`; offset += zone.height; }
		element.classList.add('has-word-review'); element.style.aspectRatio = ''; element.style.height = `${baseHeight + offset}px`;
		const source = element.querySelector('canvas')!; source.style.height = `${baseHeight}px`;
		for (const highlight of element.querySelectorAll<HTMLElement>('.word-review-insert')) { highlight.style.top = `${this.toView(page, Number(highlight.dataset.documentY))}px`; }
		this.paint(page, dimensions);
	}
	paint(page: number, dimensions: number[][]) {
		const zones = this.zones.get(page); if (!zones?.length) { return; }
		const element = this.pages.children[page] as HTMLElement, source = element.querySelector('canvas')!;
		element.querySelectorAll('.word-review-slice').forEach(node => node.remove());
		if (source.width <= 1 || source.height <= 1) { return; }
		const scale = element.clientWidth / dimensions[page][0], baseHeight = dimensions[page][1] * scale, ratio = source.height / baseHeight;
		let start = 0, offset = 0;
		for (const boundary of [...zones, { at: baseHeight, height: 0 }]) {
			const end = Math.min(baseHeight, Math.max(start, boundary.at));
			if (end > start) {
				const slice = document.createElement('canvas'); slice.className = 'word-review-slice';
				slice.width = source.width; slice.height = Math.max(1, Math.round((end - start) * ratio));
				slice.style.cssText = `top:${start + offset}px;height:${end - start}px`;
				slice.getContext('2d')!.drawImage(source, 0, start * ratio, source.width, (end - start) * ratio, 0, 0, slice.width, slice.height);
				element.append(slice);
			}
			start = end; offset += boundary.height;
		}
	}
}
