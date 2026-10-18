/*---------------------------------------------------------------------------------------------
 *  Copyright (c) Microsoft Corporation. All rights reserved.
 *  Licensed under the MIT License. See License.txt in the project root for license information.
 *--------------------------------------------------------------------------------------------*/

import DOMPurify from '../ai-office-slides/node_modules/dompurify/dist/purify.es.mjs';
import type { Kind, State, Block } from './types.js';
import { arrangeWordToolbar } from './wordToolbar.js';
import { renderAgentActivity } from './agentOverlay.js';
import { WordReview } from './wordReview.js';
declare function acquireVsCodeApi(): { postMessage(message: unknown): void };
const api = acquireVsCodeApi();
const el = (id: string) => document.getElementById(id)!;
const pages = el('pages');
const wordReview = new WordReview(pages, (data, action, ids) => send('reviewDecide', { version: data.version, reviewVersion: data.reviewVersion, action, ids }), () => { overlay(); agentOverlay(); }, id => send('draftDiscard', { id }));
const reviewSummary = document.createElement('div'); reviewSummary.className = 'word-review-summary'; reviewSummary.hidden = true; pages.before(reviewSummary);
function showReviewSummary() {
	reviewSummary.replaceChildren(); reviewSummary.hidden = !wordReview.count;
	if (!wordReview.count) { return; }
	const text = document.createElement('span'); text.textContent = `${wordReview.count} proposed ${wordReview.count === 1 ? 'change' : 'changes'} · Before Agent → Current`;
	reviewSummary.append(text);
	for (const action of ['keep', 'undo'] as const) {
		const button = document.createElement('button'); button.textContent = action === 'keep' ? 'Keep All' : 'Undo All'; button.setAttribute('aria-label', `${button.textContent} Changes in This Document`);
		button.addEventListener('click', () => { wordReview.decideAll(action); }); reviewSummary.append(button);
	}
}
const toolbar = el('toolbar');
const comments = document.createElement('details');
comments.className = 'comments'; pages.before(comments);
const keyboard = el('keyboard') as HTMLTextAreaElement;
let kind: Kind;
let state: State = { revision: -1 };
let dimensions: number[][] = [];
let selected: Block | undefined;
let currentPage = 0;
let activeEditor: HTMLTextAreaElement | undefined;
let pendingText = false;
let renderId = 0;
const visible = new Set<number>();
const rendered = new Set<number>();
const pendingRender = new Map<number, { revision: number; page: number; draft: boolean }>();
let renderBusy = false;
let composing = false;
let drag = false;
let pointerFrame = 0;
let lastPointer: { page: number; x: number; y: number } | undefined;
let initialized = false;
let disposeToolbar: (() => void) | undefined;
let agentRects: NonNullable<State['rects']> = [];
let legacyAgentRects: NonNullable<State['rects']> = [];
const agentStatus = document.createElement('span');
agentStatus.className = 'agent-status'; agentStatus.setAttribute('role', 'status'); agentStatus.setAttribute('aria-live', 'polite');
el('info').append(agentStatus);
function agentOverlay() {
	const active = renderAgentActivity(pages, agentRects.length ? agentRects : legacyAgentRects, place);
	agentStatus.textContent = active ? 'Agent is working on this content' : '';
	pages.setAttribute('aria-busy', String(active));
}

let wordText = '';
let wordTextTimer: ReturnType<typeof setTimeout> | undefined;
let wordTypingUntil = 0;
let wordRefineTimer: ReturnType<typeof setTimeout> | undefined;

function wordTyping() {
	wordTypingUntil = performance.now() + 200;
	clearTimeout(wordRefineTimer);
	wordRefineTimer = setTimeout(() => { rendered.clear(); renderNext(); }, 210);
}
let revealWordCaret = false;
function flushWordInput() {
	clearTimeout(wordTextTimer); wordTextTimer = undefined;
	if (!wordText) { return; }
	const text = wordText; wordText = '';
	api.postMessage({ type: 'execute', command: 'text.insert', args: { text } });
}
function typeWordText(text: string) {
	wordTyping();
	wordText += text;
	if (!wordTextTimer) { wordTextTimer = setTimeout(flushWordInput, 0); }
}
window.addEventListener('pagehide', () => { flushWordInput(); clearTimeout(wordRefineTimer); disposeToolbar?.(); clearAgent(); }, { once: true });
window.addEventListener('blur', flushWordInput);
function send(type: string, data: Record<string, unknown> = {}) { if (type !== 'render' && type !== 'ready') { flushWordInput(); } api.postMessage({ type, ...data }); }
function execute(command: string, args: Record<string, unknown> = {}) {
	if (kind === 'word' && command.startsWith('text.')) { wordTyping(); }
	el('error').hidden = true;
	send('execute', { command, args });
}
function message(error: unknown) { el('error').hidden = false; el('error').textContent = String(error); }
function selectionArgs() {
	return kind === 'pdf' ? { page: currentPage + 1, paragraph: selected?.id } :
		{ page: currentPage + 1, shape: selected?.id, row: selected?.row, col: selected?.col, ...(activeEditor && activeEditor.selectionEnd > activeEditor.selectionStart ? { range: { start: activeEditor.selectionStart, end: activeEditor.selectionEnd } } : {}) };
}
function flushText() {
	if (!pendingText || !activeEditor || !selected) { return; }
	const text = activeEditor.value; pendingText = false;
	execute(kind === 'pdf' ? 'text_edit' : 'text.replace', { ...selectionArgs(), text });
	selected.text = text;
}
function closeText() { flushText(); activeEditor?.remove(); activeEditor = undefined; }
async function ask(label: string, initial = ''): Promise<string | undefined> {
	const dialog = el('dialog') as HTMLDialogElement;
	const input = el('prompt') as HTMLInputElement;
	if (dialog.open) { input.focus(); return undefined; }
	el('prompt-label').textContent = label; input.value = initial; dialog.returnValue = 'cancel';
	dialog.showModal(); input.focus(); input.select();
	return new Promise(resolve => dialog.addEventListener('close', () => resolve(dialog.returnValue === 'ok' ? input.value : undefined), { once: true }));
}
function button(label: string, title: string, action: () => void, disabled = false) {
	const b = document.createElement('button'); b.textContent = label; b.title = title; b.setAttribute('aria-label', title); b.disabled = disabled;
	b.addEventListener('pointerdown', event => event.preventDefault());
	b.addEventListener('click', action); toolbar.append(b); return b;
}
function menu(label: string, items: [string, string][], action: (value: string) => void, disabled = false) {
	const select = document.createElement('select'); select.title = label; select.setAttribute('aria-label', label); select.disabled = disabled;
	const placeholder = document.createElement('option'); placeholder.textContent = label; placeholder.value = ''; select.append(placeholder);
	for (const [value, text] of items) { const option = document.createElement('option'); option.value = value; option.textContent = text; select.append(option); }
	select.addEventListener('change', () => { if (!select.value) { return; } action(select.value); select.value = ''; if (kind === 'word') { keyboard.focus({ preventScroll: true }); } });
	toolbar.append(select); return select;
}
function needsBox() { if (kind !== 'word' && !selected) { message('Click a text box or paragraph first.'); return true; } return false; }
function format(property: string, value: unknown) {
	if (needsBox()) { return; } flushText();
	if (kind === 'word') { execute('format.' + property, { [property === 'font' ? 'name' : property === 'size' ? 'size' : 'value']: value }); keyboard.focus({ preventScroll: true }); }
	else { execute(kind === 'pdf' ? 'text_edit' : 'format.text', { ...selectionArgs(), [property]: value }); }
}
function buildToolbar() {
	disposeToolbar?.();
	toolbar.replaceChildren();
	button('Text', 'Insert text', () => { if (kind === 'word') { keyboard.focus(); return; } void (async () => { const text = await ask('New text'); if (!text) { return; } execute(kind === 'pdf' ? 'page_add_text' : 'insert.text', { page: currentPage + 1, text, ...(kind === 'pdf' ? { at: [50, 50], width: 300 } : {}) }); })(); });
	const style = menu('Normal text', [['Normal', 'Normal text'], ...(kind === 'word' ? [['Title', 'Title'], ['Subtitle', 'Subtitle']] as [string, string][] : []), ['Heading1', 'Heading 1'], ['Heading2', 'Heading 2'], ['Heading3', 'Heading 3']], style => execute('para.style', { style }), kind !== 'word');
	style.id = 'paragraph-style';
	const font = menu('Font', (kind === 'pdf' ? [['helvetica', 'Sans'], ['times', 'Serif'], ['courier', 'Mono']] : [['Arial', 'Arial'], ['Times New Roman', 'Times New Roman'], ['Courier New', 'Courier New']]) as [string, string][], font => format('font', font));
	const size = menu('Size', ['10', '11', '12', '14', '16', '18', '24', '32', '48'].map(x => [x, x]), size => format('size', Number(size)));
	const emphasis: HTMLButtonElement[] = [];
	for (const [label, property] of [['B', 'bold'], ['I', 'italic'], ['U', 'underline']]) {
		const b = button(label, property, () => {
			if (kind === 'word') { execute('format.' + property); keyboard.focus({ preventScroll: true }); return; }
			const value = b.getAttribute('aria-pressed') !== 'true'; format(property, value); b.setAttribute('aria-pressed', String(value));
		});
		emphasis.push(b);
		b.style.fontWeight = property === 'bold' ? '700' : ''; b.style.fontStyle = property === 'italic' ? 'italic' : ''; b.style.textDecoration = property === 'underline' ? 'underline' : '';
	}
	const align = menu('Align', ['left', 'center', 'right', 'justify'].map(x => [x, x]), value => {
		if (needsBox()) { return; } flushText(); execute(kind === 'pdf' ? 'text_edit' : 'para.align', kind === 'word' ? { value } : { ...selectionArgs(), [kind === 'pdf' ? 'align' : 'value']: value });
	});
	const list = menu('List', [['bullet', 'Bullets'], ['number', 'Numbered'], ['none', 'No list']], value => {
		if (needsBox()) { return; } flushText();
		execute(kind === 'word' ? (value === 'number' ? 'para.numbering' : 'para.bullets') : 'para.list', kind === 'word' ? { off: value === 'none' } : { ...selectionArgs(), value });
	}, kind === 'pdf');
	const spacing = menu('Spacing', [['1', 'Single'], ['1.5', '1.5 lines'], ['2', 'Double']], value => {
		if (needsBox()) { return; } flushText(); execute(kind === 'word' ? 'para.lineSpacing' : kind === 'pdf' ? 'text_edit' : 'para.spacing', kind === 'word' ? { value: Number(value) } : { ...selectionArgs(), [kind === 'pdf' ? 'line_spacing' : 'value']: Number(value) });
	});
	const divider = document.createElement('span'); divider.className = 'divider'; toolbar.append(divider);
	const table = button('Table', kind === 'pdf' ? 'PDF table insertion is not available' : 'Insert table', () => {
		if (kind === 'word') { return; }
		void (async () => {
			const answer = await ask('Table size: rows x columns', '3 x 3'); if (!answer) { return; }
			const match = /^(\d+)\s*[xX×]\s*(\d+)$/.exec(answer);
			if (!match || Number(match[1]) > 20 || Number(match[2]) > 10 || Number(match[1]) < 1 || Number(match[2]) < 1) { message('Use 1-20 rows and 1-10 columns.'); return; }
			flushText(); execute('insert.table', { rows: Number(match[1]), cols: Number(match[2]), page: currentPage + 1 });
		})();
	}, kind === 'pdf');
	const image = button('Image', 'Insert image', () => { flushText(); send('image', { page: currentPage + 1 }); });
	const link = button('Link', 'Insert link', () => {
		void (async () => {
			if (needsBox()) { return; } const url = await ask('Link address', 'https://'); if (!url) { return; }
			try { if (!['https:', 'http:', 'mailto:'].includes(new URL(url).protocol)) { throw new Error(); } } catch { message('Enter an http, https or mailto address.'); return; }
			flushText(); execute(kind === 'pdf' ? 'link_add' : 'insert.link', kind === 'word' ? { url } : { ...selectionArgs(), ...(kind === 'pdf' ? { paragraph: undefined, rect: selected?.rect } : {}), url });
		})();
	});
	const comment = button('Comment', 'Add comment', () => {
		void (async () => {
			const text = await ask('Comment'); if (!text?.trim()) { return; } flushText();
			execute(kind === 'word' ? 'review.newComment' : kind === 'pdf' ? 'comment_add' : 'insert.comment',
				kind === 'word' ? { text } : kind === 'pdf' ? { page: currentPage + 1, type: 'note', at: selected?.rect.slice(0, 2) || [40, 40], contents: text } : { page: currentPage + 1, text });
		})();
	});
	const spacer = document.createElement('span'); spacer.className = 'spacer'; toolbar.append(spacer);
	const undo = button('Undo', 'Undo (Ctrl+Z)', () => { closeText(); send('undo'); }); undo.id = 'undo';
	const redo = button('Redo', 'Redo (Ctrl+Shift+Z)', () => { closeText(); send('redo'); }); redo.id = 'redo';
	const save = button('Save', 'Save (Ctrl+S)', () => { flushText(); send('save'); });
	const versions = kind === 'word' ? button('History', 'Document Version History', () => send('versions')) : undefined;
	if (kind === 'word') {
		disposeToolbar = arrangeWordToolbar(toolbar, { undo, redo, save, versions, style, font, size, emphasis, align, list, spacing, table, image, link, comment },
			(rows, cols) => execute('insert.table', { rows, cols, borders: true }), () => keyboard.focus({ preventScroll: true }));
	}
}
function point(event: PointerEvent, page: HTMLElement) {
	const bounds = page.getBoundingClientRect(); const index = Number(page.dataset.page);
	const scale = bounds.width / dimensions[index][0];
	return { page: index, x: (event.clientX - bounds.left) / scale, y: (kind === 'word' ? wordReview.toDocument(index, event.clientY - bounds.top) : event.clientY - bounds.top) / scale };
}
function makePages() {
	observer.disconnect(); visible.clear(); rendered.clear(); pages.replaceChildren();
	for (const [index, size] of dimensions.entries()) {
		const page = document.createElement('section'); page.className = 'page ' + kind; page.dataset.page = String(index); page.style.aspectRatio = String(size[0] / size[1]);
		page.setAttribute('aria-label', (kind === 'slides' ? 'Slide ' : 'Page ') + (index + 1));
		const content = document.createElement(kind === 'slides' ? 'img' : 'canvas'); page.append(content); pages.append(page);
		page.addEventListener('pointerdown', event => {
			if ((event.target as HTMLElement).closest('.word-review-zone')) { return; }
			if (event.target instanceof HTMLTextAreaElement || event.target instanceof HTMLButtonElement) { return; }
			currentPage = index;
			if (kind === 'word') {
				event.preventDefault(); drag = true; page.setPointerCapture(event.pointerId); keyboard.focus({ preventScroll: true });
				execute('caret.set', { ...point(event, page), extend: event.shiftKey });
			} else { closeText(); selected = undefined; execute(kind === 'pdf' ? 'select.page' : 'select.shape', { page: index + 1 }); }
		});
		page.addEventListener('pointermove', event => {
			if (!drag || kind !== 'word') { return; } lastPointer = point(event, page);
			if (!pointerFrame) { pointerFrame = requestAnimationFrame(() => { pointerFrame = 0; if (lastPointer) { execute('caret.set', { ...lastPointer, extend: true }); } }); }
		});
		page.addEventListener('pointerup', () => { drag = false; });
		page.addEventListener('dblclick', () => { if (kind === 'word') { execute('select.word'); } });
		observer.observe(page);
	}
}
const observer = new IntersectionObserver(entries => {
	for (const entry of entries) {
		const index = Number((entry.target as HTMLElement).dataset.page);
		if (entry.isIntersecting) { visible.add(index); }
		else { visible.delete(index); rendered.delete(index); const canvas = entry.target.querySelector('canvas'); if (canvas) { canvas.width = canvas.height = 1; } entry.target.querySelectorAll('.word-review-slice').forEach(node => node.remove()); }
	}
	renderNext();
}, { rootMargin: '450px' });
function renderNext() {
	if (renderBusy) { return; }
	const caretPage = kind === 'word' ? state.caret?.page : undefined;
	const page = caretPage !== undefined && visible.has(caretPage) && !rendered.has(caretPage) ? caretPage : [...visible].find(index => !rendered.has(index)); if (page === undefined) { return; }
	const node = pages.children[page] as HTMLElement;
	const draft = kind === 'word' && performance.now() < wordTypingUntil;
	const request = ++renderId; renderBusy = true; pendingRender.set(request, { revision: state.revision, page, draft });
	const canvas = node.querySelector('canvas');
	const scale = node.clientWidth * Math.min(devicePixelRatio, 2) / dimensions[page][0];
	// Use a smaller raster while keys are arriving; restore full display resolution
	// after the burst. Layout and saved document quality are always unchanged.
	send('render', { request, page, scale: draft ? Math.min(scale, 1) : scale, ...(kind === 'word' && canvas && canvas.width > 1 ? { base: Number(canvas.dataset.frame) } : {}) });
}
function place(node: HTMLElement, page: number, x: number, y: number, w: number, h: number) {
	const element = pages.children[page] as HTMLElement; if (!element || !dimensions[page]?.[0] || ![x, y, w, h].every(Number.isFinite)) { return; }
	const scale = element.clientWidth / dimensions[page][0];
	node.style.left = x * scale + 'px'; node.style.top = (kind === 'word' ? wordReview.toView(page, y * scale) : y * scale) + 'px';
	node.style.width = w ? w * scale + 'px' : '2px'; node.style.height = h * scale + 'px'; element.append(node);
}
function overlay() {
	pages.querySelectorAll('.selection,.caret,.block').forEach(node => node.remove());

	if (kind === 'word') {
		for (const r of state.rects || []) { const node = document.createElement('div'); node.className = 'selection'; place(node, r.page, r.x, r.y, r.width, r.height); }
		if (state.caret && !(state.rects?.length)) {
			const c = state.caret; const node = document.createElement('div'); node.className = 'caret'; place(node, c.page, c.x, c.y, c.width || 0, c.height || 0);
			if (c.width && !c.height) { node.style.height = '1.5px'; } else { node.style.width = '1.5px'; }
			if (revealWordCaret) { node.scrollIntoView({ block: 'nearest', inline: 'nearest' }); revealWordCaret = false; }
			const box = node.getBoundingClientRect(); keyboard.style.left = box.x + 'px'; keyboard.style.top = box.y + 'px';
		}
	} else {
		for (const block of state.blocks || []) {
			const node = document.createElement('button'); node.className = 'block'; node.setAttribute('aria-label', 'Edit ' + (block.text.slice(0, 50) || 'text box'));
			place(node, state.page || 0, block.rect[0], block.rect[1], block.rect[2] - block.rect[0], block.rect[3] - block.rect[1]);
			node.addEventListener('click', event => {
				event.stopPropagation(); closeText(); selected = block; currentPage = state.page || 0;
				const input = document.createElement('textarea'); input.className = 'block-editor'; input.setAttribute('aria-label', 'Edit text'); input.value = block.text;
				place(input, currentPage, block.rect[0], block.rect[1], block.rect[2] - block.rect[0], Math.max(block.rect[3] - block.rect[1], 32));
				input.style.fontSize = ((block.size || 12) * (pages.children[currentPage] as HTMLElement).clientWidth / dimensions[currentPage][0]) + 'px';
				input.addEventListener('input', () => { pendingText = true; flushText(); });
				input.addEventListener('keydown', event => { if (event.key === 'Escape') { closeText(); } if (event.key === 'Enter' && (event.ctrlKey || event.metaKey)) { closeText(); } });
				activeEditor = input; input.focus();
			});
		}
	}
}
function defaultHint() { el('hint').textContent = kind === 'word' ? 'Click to type · Ctrl+S to save' : 'Click text to edit · Ctrl+Enter to finish · Ctrl+S to save'; }
let agentTimer: ReturnType<typeof setTimeout> | undefined;
function clearAgent() {
	clearTimeout(agentTimer); agentTimer = undefined;
	legacyAgentRects = []; agentOverlay();
}
/**
 * Shimmers over the region an agent is currently working on. The region comes
 * from the extension host, which derives it from the executed command; agents
 * never send ids or coordinates.
 */
function showAgent(region: { page: number; rects: { x: number; y: number; width: number; height: number }[] }, label: string) {
	clearAgent();
	const element = pages.children[region.page] as HTMLElement | undefined;
	if (!element || !region.rects.length) { return; }
	const bounds = element.getBoundingClientRect();
	if (bounds.bottom < 80 || bounds.top > innerHeight - 40) { element.scrollIntoView({ block: 'nearest' }); }
	legacyAgentRects = region.rects.map(rect => ({ ...rect, page: region.page })); agentOverlay();
	el('hint').textContent = 'Agent: ' + label;
	agentTimer = setTimeout(() => { clearAgent(); defaultHint(); }, 2000);
}
keyboard.addEventListener('compositionstart', () => { composing = true; });
keyboard.addEventListener('compositionend', event => { composing = false; if (event.data) { typeWordText(event.data); } keyboard.value = ''; });
keyboard.addEventListener('input', event => {
	if (composing || (event as InputEvent).inputType === 'insertCompositionText') { return; }
	if (keyboard.value) { typeWordText(keyboard.value); } keyboard.value = '';
});
keyboard.addEventListener('paste', event => { event.preventDefault(); const text = event.clipboardData?.getData('text/plain'); if (text) { execute('text.insert', { text }); } });
document.addEventListener('copy', event => { if (document.activeElement === keyboard && state.selectedText) { event.preventDefault(); event.clipboardData?.setData('text/plain', state.selectedText); } });
document.addEventListener('cut', event => { if (document.activeElement === keyboard && state.selectedText) { event.preventDefault(); event.clipboardData?.setData('text/plain', state.selectedText); execute('text.delete'); } });
document.addEventListener('keydown', event => {
	if ((el('dialog') as HTMLDialogElement).open || composing) { return; }
	const mod = event.ctrlKey || event.metaKey;
	if (mod && event.key.toLowerCase() === 's') { event.preventDefault(); flushText(); send('save'); return; }
	if (mod && event.key.toLowerCase() === 'z') { event.preventDefault(); closeText(); send(event.shiftKey ? 'redo' : 'undo'); return; }
	if (mod && event.key.toLowerCase() === 'y') { event.preventDefault(); closeText(); send('redo'); return; }
	if (kind !== 'word' || document.activeElement !== keyboard) { return; }
	const nav: Record<string, string> = { ArrowLeft: 'left', ArrowRight: 'right', ArrowUp: 'up', ArrowDown: 'down', Home: mod ? 'docStart' : 'home', End: mod ? 'docEnd' : 'end' };
	if (nav[event.key]) { event.preventDefault(); execute('caret.' + nav[event.key], { extend: event.shiftKey }); return; }
	if (mod) {
		const command: Record<string, string> = { a: 'select.all', b: 'format.bold', i: 'format.italic', u: 'format.underline' };
		if (command[event.key.toLowerCase()]) { event.preventDefault(); execute(command[event.key.toLowerCase()]); } return;
	}
	const commands: Record<string, string> = { Enter: event.shiftKey ? 'text.lineBreak' : 'text.newParagraph', Backspace: 'text.backspace', Delete: 'text.delete', Tab: event.shiftKey ? 'text.backTab' : 'text.tab' };
	if (commands[event.key]) { event.preventDefault(); execute(commands[event.key]); }
});
window.addEventListener('message', event => {
	if (event.origin !== location.origin) { return; } const data = event.data;
	if (data.type === 'agentActivity') { agentRects = Array.isArray(data.rects) ? data.rects : []; agentOverlay(); return; }
	if (data.type === 'draft') { wordReview.updateDrafts(data.sections); wordReview.layout(dimensions); overlay(); agentOverlay(); return; }
	if (data.type === 'review') { wordReview.update(data.review); showReviewSummary(); wordReview.layout(dimensions); overlay(); agentOverlay(); return; }
	if (data.type === 'error') { message(data.error); renderBusy = false; return; }
	if (data.type === 'saved') { el('hint').textContent = data.pendingSections ? `Saved · ${data.pendingSections} pending sections are editor-only` : 'Saved'; return; }
	if (data.type === 'agent') { if (data.region) { showAgent(data.region, data.label); } return; }
	if (data.type === 'state') {
		kind = data.kind; const changed = state.revision !== data.state.revision;
		const repaginated = JSON.stringify(dimensions) !== JSON.stringify(data.pages);
		state = data.state; dimensions = data.pages;
		revealWordCaret = kind === 'word' && (data.command?.startsWith('text.') || data.command === 'insert.table');
		const commentsOpen = comments.open;
		comments.replaceChildren(); comments.open = commentsOpen; comments.hidden = !state.comments?.length;
		const summary = document.createElement('summary'); summary.textContent = `Comments (${state.comments?.length || 0})`; comments.append(summary);
		for (const comment of state.comments || []) { const p = document.createElement('p'); p.textContent = (comment.resolved ? 'Resolved · ' : '') + (comment.author || 'Comment') + ': ' + comment.text; if (typeof comment.parentId === 'number') { p.style.marginInlineStart = '24px'; } comments.append(p); }
		if (!initialized) { initialized = true; buildToolbar(); }
		if (kind === 'word') {
			const style = el('paragraph-style'); style.dataset.currentStyle = state.paragraphStyle || ''; style.dispatchEvent(new Event('formatstate'));
		}
		el('filename').textContent = data.name;
		defaultHint();
		(el('undo') as HTMLButtonElement).disabled = !state.undo; (el('redo') as HTMLButtonElement).disabled = !state.redo;
		if (repaginated || !pages.children.length) { closeText(); makePages(); }
		if (changed) { rendered.clear(); renderNext(); }
		if (kind === 'word') { wordReview.layout(dimensions); }
		overlay(); agentOverlay();
	} else if (data.type === 'rendered') {
		renderBusy = false;
		const request = pendingRender.get(data.request); pendingRender.delete(data.request);
		if (!request || (kind !== 'word' && request.revision !== state.revision) || !pages.children[data.page]) { renderNext(); return; }
		const node = pages.children[data.page]; const result = data.result;
		if (result.svg) {
			const clean = DOMPurify.sanitize(result.svg, { USE_PROFILES: { svg: true, svgFilters: true }, FORBID_TAGS: ['a', 'script', 'foreignObject', 'animate', 'set'] });
			const parsed = new DOMParser().parseFromString(clean, 'image/svg+xml');
			for (const element of parsed.querySelectorAll('[href],[*|href]')) {
				for (const attr of Array.from(element.attributes)) { if (attr.localName === 'href' && !attr.value.startsWith('#') && !/^data:image\/(png|jpeg|gif|webp);base64,/i.test(attr.value)) { element.removeAttributeNode(attr); } }
			}
			(node.querySelector('img') as HTMLImageElement).src = 'data:image/svg+xml;charset=utf-8,' + encodeURIComponent(new XMLSerializer().serializeToString(parsed));
		} else {
			const canvas = node.querySelector('canvas')!;
			const height = result.canvasHeight ?? result.height;
			if (canvas.width !== result.width || canvas.height !== height) { canvas.width = result.width; canvas.height = height; }
			if (result.height > 0) {
				const pixels = result.pixels as Uint8Array<ArrayBuffer>;
				canvas.getContext('2d')!.putImageData(new ImageData(new Uint8ClampedArray(pixels.buffer, pixels.byteOffset, pixels.byteLength), result.width, result.height), 0, result.y ?? 0);
			}
			canvas.dataset.frame = String(data.request);
			canvas.dataset.revision = String(request.revision);
			if (kind === 'word') { wordReview.paint(data.page, dimensions); }
		}
		// Show completed Word frames during continuous typing instead of discarding
		// every frame until the user stops. A newer revision still requests a refresh.
		if (request.revision === state.revision && (!request.draft || performance.now() < wordTypingUntil)) { rendered.add(data.page); } renderNext();
	}
});
let resizeTimer: ReturnType<typeof setTimeout>;
window.addEventListener('resize', () => { clearTimeout(resizeTimer); resizeTimer = setTimeout(() => { rendered.clear(); renderNext(); if (kind === 'word') { wordReview.layout(dimensions); } overlay(); agentOverlay(); }, 200); });
send('ready');
