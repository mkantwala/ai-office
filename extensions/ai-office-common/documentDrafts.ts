/*---------------------------------------------------------------------------------------------
 *  Copyright (c) Microsoft Corporation. All rights reserved.
 *  Licensed under the MIT License. See License.txt in the project root for license information.
 *--------------------------------------------------------------------------------------------*/

import { randomUUID } from 'node:crypto';

export interface DraftSection { id: string; title: string; intent: string; level: number; roots: string[] }
export interface DraftAnchor { id: string; position: 'before' | 'after' }

/** Pending intentions are session-local editor state, never placeholder paragraphs in DOCX. */
export class DocumentDrafts {
	private anchor?: DraftAnchor;
	private sections: DraftSection[] = [];
	plan(anchor: DraftAnchor, input: { title: string; intent: string; level?: number }[]) {
		if (this.sections.some(section => !section.roots.length)) { throw new Error('A draft already has unfinished sections. Fill or discard them before planning another draft.'); }
		if (!Array.isArray(input) || !input.length || input.length > 20 || input.some(section => !section || typeof section.title !== 'string' || !section.title.trim() || section.title.length > 120 || typeof section.intent !== 'string' || !section.intent.trim() || section.intent.length > 300 || section.level !== undefined && ![1, 2, 3].includes(section.level))) { throw new Error('Plan 1-20 sections with a title (1-120 characters), intent (1-300 characters), and optional level 1-3.'); }
		this.anchor = anchor;
		this.sections = input.map(section => ({ id: randomUUID(), title: section.title, intent: section.intent, level: section.level ?? 2, roots: [] }));
		return this.status();
	}
	status() { return { sections: this.sections.map(section => ({ ...section, state: section.roots.length ? 'filled' : 'pending' })), pending: this.sections.filter(section => !section.roots.length).length, savedInDocument: false }; }
	section(id: string) {
		const section = this.sections.find(section => section.id === id);
		if (!section || section.roots.length) { throw new Error('This pending section no longer exists. Read draft status before continuing.'); }
		return section;
	}
	location(id: string): DraftAnchor {
		const section = this.section(id), index = this.sections.indexOf(section);
		const previous = this.sections.slice(0, index).reverse().find(section => section.roots.length);
		if (previous) { return { id: previous.roots.at(-1)!, position: 'after' }; }
		const next = this.sections.slice(index + 1).find(section => section.roots.length);
		if (next) { return { id: next.roots[0], position: 'before' }; }
		if (!this.anchor) { throw new Error('Draft has no document anchor.'); }
		return this.anchor;
	}
	complete(id: string, roots: string[]) {
		if (!roots.length) { throw new Error('The section did not create document blocks.'); }
		this.section(id).roots = roots;
	}
	discard(id?: string) {
		if (id) { const section = this.section(id); this.sections.splice(this.sections.indexOf(section), 1); }
		else { this.sections = []; this.anchor = undefined; }
		return this.status();
	}
	pending() { return this.sections.filter(section => !section.roots.length).map(section => ({ id: section.id, title: section.title, intent: section.intent, anchor: this.location(section.id) })); }
}
