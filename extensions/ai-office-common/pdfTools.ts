/*---------------------------------------------------------------------------------------------
 *  Copyright (c) Microsoft Corporation. All rights reserved.
 *  Licensed under the MIT License. See License.txt in the project root for license information.
 *--------------------------------------------------------------------------------------------*/

/** Parse human page ranges without silently accepting partial or unbounded selections. */
export function pdfPageRange(text: string, total: number): number[] {
	if (!Number.isSafeInteger(total) || total < 1 || total > 10000) { throw new Error('The document has an unsupported page count.'); }
	if (text.trim().toLowerCase() === 'all') { return Array.from({ length: total }, (_, index) => index + 1); }
	if (text.length > 2000 || !text.trim()) { throw new Error('Enter pages such as 1, 3–5, or all.'); }
	const pages = new Set<number>();
	for (const token of text.split(',')) {
		const match = /^\s*(\d+)(?:\s*[-–]\s*(\d+))?\s*$/.exec(token);
		if (!match) { throw new Error('Use page numbers and ranges, such as 1, 3–5.'); }
		const start = Number(match[1]), end = Number(match[2] || start);
		if (start < 1 || end < start || end > total) { throw new Error(`Pages must be between 1 and ${total}.`); }
		for (let page = start; page <= end; page++) { pages.add(page); }
	}
	return [...pages].sort((a, b) => a - b);
}

export class PdfTools {
	private dialog?: HTMLDialogElement;
	constructor(private readonly snapshot: () => { total: number; page: number; revision: number }, private readonly execute: (command: string, params: Record<string, unknown>, revision: number) => void) { }
	dispose() { this.dialog?.remove(); }
	open() {
		if (this.dialog?.open) { this.dialog.focus(); return; }
		this.dialog?.remove();
		const snapshot = { ...this.snapshot() };
		const dialog = this.dialog = document.createElement('dialog'); dialog.className = 'pdf-tools-dialog'; dialog.setAttribute('aria-labelledby', 'pdf-tools-title');
		const form = document.createElement('form'); form.method = 'dialog';
		const title = document.createElement('h2'); title.id = 'pdf-tools-title'; title.textContent = 'PDF Tools';
		const intro = document.createElement('p'); intro.textContent = `${snapshot.total} pages · Changes use document Undo and are saved with the file.`;
		const field = (label: string, input: HTMLElement) => { const row = document.createElement('label'); row.textContent = label; row.append(input); form.append(row); };
		const operation = document.createElement('select'); operation.setAttribute('aria-label', 'PDF operation');
		for (const [value, label] of [['page_rotate', 'Rotate pages'], ['page_duplicate', 'Duplicate pages'], ['page_delete', 'Delete pages'], ['page_move', 'Move pages'], ['page_insert_blank', 'Insert blank page'], ['page_set_box', 'Crop pages'], ['doc_watermark', 'Add text watermark'], ['doc_header_footer', 'Add page numbers'], ['doc_background', 'Page background']]) { operation.add(new Option(label, value)); }
		form.append(title, intro); field('Operation', operation);
		const range = document.createElement('input'); range.value = String(snapshot.page + 1); range.placeholder = '1, 3–5, or all'; range.setAttribute('aria-label', 'Pages'); field('Pages', range);
		const options = document.createElement('div'); options.className = 'pdf-tool-options'; form.append(options);
		const error = document.createElement('p'); error.className = 'pdf-tool-error'; error.setAttribute('role', 'alert'); error.hidden = true;
		let value: HTMLInputElement | HTMLSelectElement;
		const render = () => {
			options.replaceChildren(); error.hidden = true; range.disabled = operation.value === 'page_insert_blank';
			const label = document.createElement('label');
			if (operation.value === 'page_rotate') {
				const select = document.createElement('select'); select.setAttribute('aria-label', 'Rotation'); for (const [degrees, text] of [['90', '90° clockwise'], ['180', '180°'], ['270', '90° counterclockwise']]) { select.add(new Option(text, degrees)); } value = select; label.textContent = 'Rotation';
			} else {
				const input = document.createElement('input'); value = input;
				const type = operation.value;
				if (['page_move', 'page_insert_blank', 'page_set_box'].includes(type)) { input.type = 'number'; input.min = type === 'page_set_box' ? '0' : '1'; input.max = String(type === 'page_insert_blank' ? snapshot.total + 1 : snapshot.total); input.value = type === 'page_set_box' ? '18' : String(snapshot.page + 1); if (type === 'page_set_box') { input.max = '7200'; input.step = '0.5'; } }
				if (type === 'doc_watermark') { input.value = 'DRAFT'; input.maxLength = 240; }
				if (type === 'doc_header_footer') { input.value = '<<Page 1 of n>>'; input.maxLength = 240; }
				if (type === 'doc_background') { input.type = 'color'; input.value = '#ffffff'; }
				label.textContent = type === 'page_move' ? 'Destination position' : type === 'page_insert_blank' ? 'Insert at page' : type === 'page_set_box' ? 'Crop from every edge (points)' : type === 'doc_watermark' ? 'Watermark text' : type === 'doc_header_footer' ? 'Footer text · <<1>> is the page number' : 'Background color';
				if (['page_delete', 'page_duplicate'].includes(type)) { label.hidden = true; }
			}
			label.append(value); options.append(label);
			const destructive = operation.value === 'page_delete'; apply.textContent = destructive ? 'Delete Selected Pages' : 'Apply'; apply.classList.toggle('destructive', destructive);
		};
		const footer = document.createElement('footer'), cancel = document.createElement('button'), apply = document.createElement('button'); cancel.type = 'button'; cancel.textContent = 'Cancel'; cancel.addEventListener('click', () => dialog.close()); apply.type = 'submit'; apply.textContent = 'Apply'; footer.append(cancel, apply); form.append(error, footer);
		operation.addEventListener('change', render); render();
		form.addEventListener('submit', event => {
			event.preventDefault();
			try {
				if (snapshot.revision !== this.snapshot().revision) { throw new Error('The document changed. Close and reopen PDF Tools before applying.'); }
				const command = operation.value, pages = command === 'page_insert_blank' ? [] : pdfPageRange(range.value, snapshot.total);
				let params: Record<string, unknown> = { pages };
				if (command === 'page_delete' && pages.length >= snapshot.total) { throw new Error('Keep at least one page in the PDF.'); }
				if (command === 'page_rotate') { params.degrees = Number(value.value); }
				if (command === 'page_insert_blank' || command === 'page_move') { const position = Number(value.value); if (!Number.isSafeInteger(position) || position < 1 || position > snapshot.total + (command === 'page_insert_blank' ? 1 : 0)) { throw new Error('Choose a valid page position.'); } if (command === 'page_insert_blank') { params = { at: position }; } else { params.to = position; } }
				if (command === 'page_set_box') { const margin = Number(value.value); if (!Number.isFinite(margin) || margin < 0 || margin > 7200) { throw new Error('Choose a crop margin from 0 to 7200 points.'); } params.box = 'crop'; params.margins = [margin, margin, margin, margin]; }
				if (command === 'doc_watermark' || command === 'doc_header_footer') { if (!value.value.trim()) { throw new Error('Enter text before applying.'); } if (command === 'doc_watermark') { params.text = value.value; params.opacity = 0.25; params.rotation = 45; } else { params.footer_center = value.value; params.font_size = 10; } }
				if (command === 'doc_background') { params.color = value.value; }
				this.execute(command, params, snapshot.revision); dialog.close();
			} catch (reason) { error.textContent = reason instanceof Error ? reason.message : String(reason); error.hidden = false; }
		});
		dialog.append(form); document.body.append(dialog); dialog.showModal(); operation.focus();
	}
}
