/*---------------------------------------------------------------------------------------------
 *  Copyright (c) Microsoft Corporation. All rights reserved.
 *  Licensed under the MIT License. See License.txt in the project root for license information.
 *--------------------------------------------------------------------------------------------*/

import { arrangeWordToolbar } from './wordToolbar.js';

export interface DocumentToolbarControl {
	key: string;
	currentValue?: string;
	type: 'button' | 'select';
	label: string;
	text?: string;
	disabled: boolean;
	pressed?: boolean;
	options?: { value: string; label: string }[];
}
export interface DocumentToolbarAction { control: string; value?: string; rows?: number; cols?: number }
export interface DocumentToolbarSnapshot {
	uri: string; kind: 'word' | 'pdf' | 'slides'; panel: string; instance: string; context: string;
	controls: DocumentToolbarControl[]; paragraphStyle?: string;
}

/** Render the same compact controls in app chrome without owning document selection or mutations. */
export function createDocumentToolbar(element: HTMLElement, descriptors: DocumentToolbarControl[], action: (action: DocumentToolbarAction) => void, focusDocument: () => void) {
	const ownerDocument = element.ownerDocument;
	const targetWindow = ownerDocument.defaultView!;
	const controls = new Map<string, HTMLButtonElement | HTMLSelectElement>();
	for (const descriptor of descriptors) {
		const control = ownerDocument.createElement(descriptor.type === 'select' ? 'select' : 'button');
		control.title = descriptor.label; control.setAttribute('aria-label', descriptor.label); control.dataset.officeControl = descriptor.key;
		if (control instanceof targetWindow.HTMLSelectElement) {
			for (const option of descriptor.options || []) { control.add(new targetWindow.Option(option.label, option.value)); }
			control.addEventListener('change', () => { if (control.value) { action({ control: descriptor.key, value: control.value }); control.value = ''; } });
		} else {
			control.type = 'button'; control.textContent = descriptor.text || descriptor.label;
			control.addEventListener('pointerdown', event => event.preventDefault());
			control.addEventListener('click', () => action({ control: descriptor.key }));
		}
		if (descriptor.currentValue !== undefined) { control.dataset[descriptor.key === 'font' ? 'currentFont' : 'currentSize'] = descriptor.currentValue; }
		control.disabled = descriptor.disabled; controls.set(descriptor.key, control);
	}
	const button = (key: string) => controls.get(key) as HTMLButtonElement;
	const select = (key: string) => controls.get(key) as HTMLSelectElement;
	const emphasis = ['bold', 'italic', 'underline'].map(key => button(key));
	for (const [index, control] of emphasis.entries()) { control.textContent = ['B', 'I', 'U'][index]; control.style.fontWeight = index === 0 ? '700' : ''; control.style.fontStyle = index === 1 ? 'italic' : ''; control.style.textDecoration = index === 2 ? 'underline' : ''; }
	const dispose = arrangeWordToolbar(element, {
		text: button('text'), undo: button('undo'), redo: button('redo'), save: button('save'), versions: button('versions'), properties: button('properties'),
		strike: button('strike'), clear: button('clear'), color: select('color'), highlight: select('highlight'),
		style: select('style'), font: select('font'), size: select('size'), emphasis,
		align: select('align'), list: select('list'), spacing: select('spacing'), table: button('table'), image: button('image'), link: button('link'), comment: button('comment')
	}, (rows, cols) => action({ control: 'table', rows, cols }), focusDocument);
	return {
		update(descriptors: DocumentToolbarControl[], paragraphStyle?: string) {
			for (const descriptor of descriptors) {
				const control = controls.get(descriptor.key); if (!control) { continue; }
				control.disabled = descriptor.disabled;
				if (descriptor.currentValue !== undefined) { control.dataset[descriptor.key === 'font' ? 'currentFont' : 'currentSize'] = descriptor.currentValue; control.dispatchEvent(new targetWindow.Event('formatstate')); }
				if (descriptor.pressed !== undefined) { control.setAttribute('aria-pressed', String(descriptor.pressed)); }
			}
			const style = select('style'); style.dataset.currentStyle = paragraphStyle || ''; style.dispatchEvent(new Event('formatstate'));
		},
		dispose
	};
}

/** Cross-process toolbar messages contain bounded UI metadata, never document content. */
export function isDocumentToolbarSnapshot(value: unknown): value is DocumentToolbarSnapshot {
	if (!value || typeof value !== 'object') { return false; }
	const snapshot = value as DocumentToolbarSnapshot;
	if (!['word', 'pdf', 'slides'].includes(snapshot.kind) || ['uri', 'panel', 'instance', 'context'].some(key => typeof snapshot[key as keyof DocumentToolbarSnapshot] !== 'string' || String(snapshot[key as keyof DocumentToolbarSnapshot]).length > 2000)) { return false; }
	if (snapshot.paragraphStyle !== undefined && (typeof snapshot.paragraphStyle !== 'string' || snapshot.paragraphStyle.length > 200)) { return false; }
	if (!Array.isArray(snapshot.controls) || snapshot.controls.length > 28) { return false; }
	const selectKeys = ['style', 'font', 'size', 'align', 'list', 'spacing', 'color', 'highlight'];
	const buttonKeys = ['text', 'bold', 'italic', 'underline', 'table', 'image', 'link', 'comment', 'undo', 'redo', 'save', 'properties', 'versions', 'strike', 'clear'];
	const seen = new Set<string>();
	for (const control of snapshot.controls) {
		if (!control || seen.has(control.key) || ![...selectKeys, ...buttonKeys].includes(control.key) || control.type !== (selectKeys.includes(control.key) ? 'select' : 'button') || typeof control.label !== 'string' || control.label.length > 160 || typeof control.disabled !== 'boolean') { return false; }
		if (control.type === 'select' && (!Array.isArray(control.options) || control.options.length > 40 || control.options.some(option => !option || typeof option.value !== 'string' || option.value.length > 200 || typeof option.label !== 'string' || option.label.length > 200))) { return false; }
		if (control.currentValue !== undefined && (!['font', 'size'].includes(control.key) || typeof control.currentValue !== 'string' || control.currentValue.length > 200)) { return false; }
		seen.add(control.key);
	}
	return [...selectKeys.filter(key => !['color', 'highlight'].includes(key)), ...buttonKeys.filter(key => !['text', 'properties', 'versions', 'strike', 'clear'].includes(key))].every(key => seen.has(key));
}
