/*---------------------------------------------------------------------------------------------
 *  Copyright (c) Microsoft Corporation. All rights reserved.
 *  Licensed under the MIT License. See License.txt in the project root for license information.
 *--------------------------------------------------------------------------------------------*/

import { test } from 'node:test';
import assert from 'node:assert/strict';
import { resolve } from 'node:path';
import { mkdir, writeFile, readFile } from 'node:fs/promises';
import { spawnSync } from 'node:child_process';
import { deflateSync } from 'node:zlib';
import { zipSync, strToU8, unzipSync, strFromU8 } from '../ai-office-slides/node_modules/fflate/esm/index.mjs';
import * as ppt from '../ai-office-slides/node_modules/@office-kit/pptx/dist/index.js';
import { openEngine } from './model.js';
import { WordFrame } from './wordFrame.js';

const root = resolve('extensions');
const demo = resolve('.build/office-edit-demo');
// Original single blue pixel; no external artwork or font assets.
function imageFixture() {
	const chunk = (name: string, data: Buffer) => {
		const payload = Buffer.concat([Buffer.from(name), data]); let crc = 0xffffffff;
		for (const byte of payload) { crc ^= byte; for (let bit = 0; bit < 8; bit++) { crc = (crc >>> 1) ^ (crc & 1 ? 0xedb88320 : 0); } }
		const length = Buffer.alloc(4), checksum = Buffer.alloc(4); length.writeUInt32BE(data.length); checksum.writeUInt32BE((crc ^ 0xffffffff) >>> 0);
		return Buffer.concat([length, payload, checksum]);
	};
	const header = Buffer.alloc(13); header.writeUInt32BE(1); header.writeUInt32BE(1, 4); header[8] = 8; header[9] = 6;
	return Buffer.concat([Buffer.from([137, 80, 78, 71, 13, 10, 26, 10]), chunk('IHDR', header), chunk('IDAT', deflateSync(Buffer.from([0, 40, 80, 220, 255]))), chunk('IEND', Buffer.alloc(0))]).toString('base64');
}
const docx = zipSync(Object.fromEntries(Object.entries({
	'[Content_Types].xml': '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>',
	'_rels/.rels': '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>',
	'word/document.xml': '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:p><w:r><w:t>AI Office editing</w:t></w:r></w:p><w:p><w:r><w:t>Select this paragraph and try the toolbar.</w:t></w:r></w:p><w:sectPr><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440"/></w:sectPr></w:body></w:document>'
}).map(([name, value]) => [name, strToU8(value)])));

function pdfFixture() {
	const stream = 'BT /F1 20 Tf 72 700 Td (AI Office editing) Tj ET\n';
	const objects = ['<< /Type /Catalog /Pages 2 0 R >>', '<< /Type /Pages /Kids [3 0 R] /Count 1 >>', '<< /Type /Page /Parent 2 0 R /MediaBox [0 0 612 792] /Resources << /Font << /F1 5 0 R >> >> /Contents 4 0 R >>', `<< /Length ${stream.length} >>\nstream\n${stream}endstream`, '<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica >>'];
	let value = '%PDF-1.7\n'; const offsets = [0];
	objects.forEach((obj, i) => { offsets.push(Buffer.byteLength(value)); value += `${i + 1} 0 obj\n${obj}\nendobj\n`; });
	const xref = Buffer.byteLength(value);
	value += `xref\n0 ${offsets.length}\n0000000000 65535 f \n` + offsets.slice(1).map(n => String(n).padStart(10, '0') + ' 00000 n \n').join('') + `trailer\n<< /Size ${offsets.length} /Root 1 0 R >>\nstartxref\n${xref}\n%%EOF\n`;
	return Buffer.from(value);
}

test('Word: native editing, toolbar operations, undo, export and reopen', async () => {
	await mkdir(demo, { recursive: true }); await writeFile(resolve(demo, 'Memo.docx'), docx);
	const engine = await openEngine('word', 'Memo.docx', docx, root);
	try {
		await engine.execute('select.text', { text: 'AI Office' });
		await engine.execute('text.insert', { text: 'Interactive office' });
		await engine.execute('edit.undo', {}); await engine.execute('select.all', {});
		assert.match(engine.state().selectedText!, /AI Office/);
		await engine.execute('edit.redo', {}); await engine.execute('select.text', { text: 'Interactive office' });
		for (const command of ['format.bold', 'format.italic', 'format.underline']) { await engine.execute(command, { value: true }); }
		await engine.execute('format.font', { name: 'Arial' }); await engine.execute('format.size', { size: 18 });
		await engine.execute('para.align', { value: 'center' }); await engine.execute('para.style', { style: 'Heading1' });
		await engine.execute('para.bullets', {}); await engine.execute('para.numbering', {}); await engine.execute('para.lineSpacing', { value: 1.5 });
		await engine.execute('insert.link', { url: 'https://example.com', text: 'Office link' });
		await engine.execute('review.newComment', { text: 'Review this memo' });
		assert.equal(engine.state().comments?.[0].text, 'Review this memo');
		await engine.execute('caret.docEnd', {}); await engine.execute('text.newParagraph', {}); await engine.execute('insert.table', { rows: 2, cols: 2 });
		await engine.execute('insert.picture', { data: imageFixture(), width: 36 });
		const bytes = await engine.export();
		const reopened = await openEngine('word', 'saved.docx', bytes, root);
		try { await reopened.execute('select.all', {}); assert.match(reopened.state().selectedText!, /Interactive office/); assert.equal(reopened.state().comments?.[0].text, 'Review this memo'); assert.ok(reopened.render(0, .5).pixels!.length > 100); } finally { reopened.dispose(); }
		const xml = strFromU8(unzipSync(bytes)['word/document.xml']); assert.match(xml, /w:tbl/); assert.match(xml, /w:b/);
		await assert.rejects(engine.execute('file.save', {}));
	} finally { engine.dispose(); }
});

test('Word visual styles and visible table borders survive save and reopen', async () => {
	const engine = await openEngine('word', 'table.docx', docx, root);
	try {
		for (const style of ['Title', 'Subtitle', 'Heading1', 'Heading2', 'Heading3', 'Normal']) {
			await engine.execute('para.style', { style }); assert.equal(engine.state().paragraphStyle, style);
		}
		await engine.execute('caret.docEnd', {}); await engine.execute('text.newParagraph', {});
		await engine.execute('insert.table', { rows: 3, cols: 7, borders: true });
		const bytes = await engine.export();
		const xml = strFromU8(unzipSync(bytes)['word/document.xml']);
		assert.match(xml, /w:tblBorders/); assert.match(xml, /w:insideH/); assert.match(xml, /w:insideV/);
		const reopened = await openEngine('word', 'table.docx', bytes, root);
		try {
			const raster = reopened.render(0, 1.5); let longest = 0;
			for (let y = 0; y < raster.height!; y++) {
				let run = 0;
				for (let x = 0; x < raster.width!; x++) {
					const offset = (y * raster.width! + x) * 4;
					run = raster.pixels![offset] < 230 && raster.pixels![offset + 1] < 230 && raster.pixels![offset + 2] < 230 ? run + 1 : 0;
					longest = Math.max(longest, run);
				}
			}
			assert.ok(longest > 500, 'An empty table must paint continuous visible horizontal borders');
		} finally { reopened.dispose(); }
	} finally { engine.dispose(); }
});

test('Word imports missing built-in styles without replacing custom styles', async () => {
	const files = unzipSync(docx);
	files['word/styles.xml'] = strToU8('<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:style w:type="paragraph" w:styleId="Normal"><w:name w:val="Normal"/><w:rPr><w:sz w:val="26"/></w:rPr></w:style><w:style w:type="paragraph" w:styleId="Heading1"><w:name w:val="Heading 1"/><w:rPr><w:sz w:val="44"/></w:rPr></w:style></w:styles>');
	files['word/_rels/document.xml.rels'] = strToU8('<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="styles" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/></Relationships>');
	const engine = await openEngine('word', 'import.docx', zipSync(files), root);
	try {
		await engine.execute('para.style', { style: 'Title' });
		assert.equal(engine.state().paragraphStyle, 'Title');
		await engine.execute('edit.undo', {});
		assert.doesNotMatch(strFromU8(unzipSync(await engine.export())['word/styles.xml']), /w:styleId="Title"/);
		await engine.execute('edit.redo', {});
		for (const style of ['Subtitle', 'Heading1', 'Heading2', 'Heading3', 'Title']) {
			await engine.execute('para.style', { style }); assert.equal(engine.state().paragraphStyle, style);
		}
		await assert.rejects(engine.execute('para.style', { style: 'Missing custom style' }));
		const bytes = await engine.export();
		const styles = strFromU8(unzipSync(bytes)['word/styles.xml']);
		assert.match(styles, /w:sz w:val="26"/); assert.match(styles, /w:sz w:val="44"/);
		const reopened = await openEngine('word', 'saved.docx', bytes, root);
		try { assert.equal(reopened.state().paragraphStyle, 'Title'); assert.ok(reopened.render(0, .5).pixels!.length); } finally { reopened.dispose(); }
	} finally { engine.dispose(); }
});

test('Word frame patches reconstruct full pixels and reject stale bases', () => {
	const cache = new WordFrame();
	const initial = { pixels: new Uint8Array(400).fill(255), width: 10, height: 10 };
	cache.update(initial, 0, 1, 1);
	const edited = { ...initial, pixels: initial.pixels.slice() }; edited.pixels[125] = 0;
	const patch = cache.update(edited, 0, 1, 2, 1);
	assert.equal(patch.height, 1); assert.equal(patch.pixels.length, 40);
	assert.equal(patch.y, 3);
	const painted = initial.pixels.slice(); painted.set(patch.pixels, patch.y! * 40);
	assert.deepEqual(painted, edited.pixels);
	assert.equal(cache.update(edited, 0, 1, 3, 2).height, 0);
	assert.equal(cache.update(edited, 0, 1, 4, 1).height, 10);
	assert.equal(cache.update(edited, 1, 1, 5, 4).height, 10);
	assert.equal(cache.update(edited, 1, 2, 6, 5).height, 10);
	assert.equal(cache.update(edited, 1, 2, 7).height, 10);
});

test('PDF: edit paragraph, format, comments, links, undo and reopen', async () => {
	await mkdir(demo, { recursive: true }); const source = pdfFixture(); await writeFile(resolve(demo, 'Contract.pdf'), source);
	const engine = await openEngine('pdf', 'Contract.pdf', source, root);
	try {
		assert.match(engine.state().blocks![0].text, /AI Office/);
		await engine.execute('text_edit', { page: 1, paragraph: 1, text: 'Editable PDF' });
		assert.match(engine.state().blocks![0].text, /Editable PDF/);
		await engine.execute('edit_undo', {}); assert.match(engine.state().blocks![0].text, /AI Office/);
		await engine.execute('edit_redo', {});
		await engine.execute('text_edit', { page: 1, paragraph: 1, bold: true, italic: true, underline: true, font: 'times', size: 18, align: 'center', line_spacing: 1.5 });
		await engine.execute('comment_add', { page: 1, type: 'note', at: [50, 50], contents: 'PDF review' });
		assert.ok(engine.state().comments?.some(c => c.text === 'PDF review'));
		await engine.execute('link_add', { page: 1, rect: [72, 72, 200, 100], url: 'https://example.com' });
		await engine.execute('page_add_text', { page: 1, at: [72, 180], width: 300, text: 'Additional text' });
		await engine.execute('image.insert', { page: 1, data: imageFixture(), name: 'pixel.png' });
		const reopened = await openEngine('pdf', 'saved.pdf', await engine.export(), root);
		try { assert.match(reopened.state().blocks!.map(b => b.text).join(' '), /Editable PDF/); assert.ok(reopened.state().comments?.some(c => c.text === 'PDF review')); assert.ok(reopened.render(0, .5).pixels!.length > 100); } finally { reopened.dispose(); }
	} finally { engine.dispose(); }
});

test('Slides: text and table cells, toolbar, undo, export and reopen', async () => {
	const deck = ppt.createPresentation(); const slide = ppt.addSlide(deck, { layout: ppt.getSlideLayouts(deck)[0] });
	ppt.addSlideTextBox(slide, { x: ppt.inches(1), y: ppt.inches(1), w: ppt.inches(8), h: ppt.inches(1), text: 'AI Office presentation' });
	const source = await ppt.savePresentation(deck); await mkdir(demo, { recursive: true }); await writeFile(resolve(demo, 'Board.pptx'), source);
	const engine = await openEngine('slides', 'Board.pptx', source, root);
	try {
		const shape = engine.state().blocks!.find(b => b.text.includes('AI Office'))!.id;
		await engine.execute('text.replace', { page: 1, shape, text: 'Editable slides' });
		await engine.execute('edit.undo', {}); assert.ok(engine.state().blocks!.some(b => b.text === 'AI Office presentation'));
		await engine.execute('edit.redo', {});
		await engine.execute('format.text', { page: 1, shape, bold: true, italic: true, underline: true, font: 'Arial', size: 24, range: { start: 0, end: 8 } });
		for (const [command, value] of [['para.align', 'center'], ['para.list', 'number'], ['para.spacing', 1.5]] as const) { await engine.execute(command, { page: 1, shape, value }); }
		await engine.execute('insert.link', { page: 1, shape, url: 'https://example.com' });
		await engine.execute('insert.comment', { page: 1, text: 'Slide review' });
		await engine.execute('insert.image', { page: 1, data: imageFixture() });
		await engine.execute('insert.table', { page: 1, rows: 2, cols: 2 });
		const cell = engine.state().blocks!.find(b => b.row === 0 && b.col === 0)!;
		assert.ok(cell); await engine.execute('text.replace', { page: 1, shape: cell.id, row: 0, col: 0, text: 'Revenue' });
		const bytes = await engine.export(); const reopened = await openEngine('slides', 'saved.pptx', bytes, root);
		try { assert.ok(reopened.state().blocks!.some(b => b.text === 'Revenue')); assert.equal(reopened.state().comments?.[0].text, 'Slide review'); assert.match(reopened.render(0, 1).svg!, /Editable/); } finally { reopened.dispose(); }
		await assert.rejects(engine.execute('format.text', { page: 1, shape, size: -1 }));
		assert.ok(engine.state().blocks!.some(b => b.text === 'Revenue'));
	} finally { engine.dispose(); }
});

test('CLI runs agent scripts for all formats and rejects paths outside its root', async () => {
	const cli = resolve('extensions/ai-office-pdf/out/cli.cjs');
	for (const [file, command, params] of [
		['Memo.docx', 'text.insert', { text: 'CLI Word ' }],
		['Contract.pdf', 'text_edit', { page: 1, paragraph: 1, text: 'CLI PDF' }],
		['Board.pptx', 'text.replace', { page: 1, shape: 2, text: 'CLI Slides' }]
	] as const) {
		const script = resolve(demo, 'changes.json'), output = resolve(demo, 'cli-' + file);
		await writeFile(script, JSON.stringify([{ command, params }]));
		const run = spawnSync(process.execPath, [cli, 'run', resolve(demo, file), '--script', script, '--output', output, '--root', demo], { encoding: 'utf8' });
		assert.equal(run.status, 0, run.stderr); assert.equal(JSON.parse(run.stdout).commands, 1);
		const kind = file.endsWith('docx') ? 'word' : file.endsWith('pdf') ? 'pdf' : 'slides';
		const reopened = await openEngine(kind, file, await readFile(output), root);
		try {
			if (kind === 'word') { await reopened.execute('select.all', {}); assert.match(reopened.state().selectedText!, /CLI Word/); }
			else { assert.match(reopened.state().blocks!.map(b => b.text).join(' '), /CLI PDF|CLI Slides/); }
		} finally { reopened.dispose(); }
	}
	await writeFile(resolve(demo, '../office-outside.docx'), docx);
	const denied = spawnSync(process.execPath, [cli, 'inspect', '../office-outside.docx', '--root', demo], { encoding: 'utf8' });
	assert.notEqual(denied.status, 0); assert.match(denied.stderr, /outside --root/);
});
