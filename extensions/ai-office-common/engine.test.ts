/*---------------------------------------------------------------------------------------------
 *  Copyright (c) Microsoft Corporation. All rights reserved.
 *  Licensed under the MIT License. See License.txt in the project root for license information.
 *--------------------------------------------------------------------------------------------*/

import { test } from 'node:test';
import assert from 'node:assert/strict';
import { resolve } from 'node:path';
import { mkdir, writeFile, readFile, mkdtemp, rm } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import { spawnSync } from 'node:child_process';
import { deflateSync } from 'node:zlib';
import { zipSync, strToU8, unzipSync, strFromU8 } from '../ai-office-slides/node_modules/fflate/esm/index.mjs';
import * as ppt from '../ai-office-slides/node_modules/@office-kit/pptx/dist/index.js';
import { openEngine } from './model.js';
import { prepareWordFormatting, validateDocumentEdits } from './wordFormatting.js';
import { WordFrame } from './wordFrame.js';
import { activityRegion } from './agentActivity.js';
import { DocumentDrafts } from './documentDrafts.js';
import { DocumentVersions } from './documentVersions.js';
import { activeAgentRequest, activeAgentLabel, withAgentRequest } from './agentRequest.js';
import type { Block, State } from './types.js';

interface WordAgentBlock { id: string; text: string; kind: string }

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

test('Word caret follows the active font rather than line spacing or a neighboring large run', async () => {
	const engine = await openEngine('word', 'Caret.docx', docx, root);
	try {
		await engine.execute('select.text', { text: 'AI Office' }); await engine.execute('format.size', { size: 12 });
		await engine.execute('caret.docStart', {}); const normal = engine.state().caret!;
		await engine.execute('para.lineSpacing', { value: 3 }); const spaced = engine.state().caret!;
		assert.ok(normal.height > 8 && normal.height < 24); assert.ok(Math.abs(normal.height - spaced.height) < .1, 'Leading must not enlarge the caret');
		await engine.execute('select.text', { text: 'editing' }); await engine.execute('format.size', { size: 48 });
		await engine.execute('caret.docStart', {}); assert.ok(Math.abs(engine.state().caret!.height - normal.height) < .1, 'Neighboring large text must not enlarge a small-text caret');
		await engine.execute('select.text', { text: 'editing' }); const large = engine.state().caret!; assert.ok(large.height > normal.height * 2);
	} finally { engine.dispose(); }
});

test('Word page ruler remains available in lists, tables and after select-all', async () => {
	const engine = await openEngine('word', 'ruler.docx', docx, root);
	try {
		const original = engine.state().pageRuler;
		assert.ok(original && original.width > 0);
		await engine.execute('para.bullets', {});
		assert.deepEqual({ page: engine.state().pageRuler, handles: engine.state().paragraph?.ruler }, { page: original, handles: null });
		await engine.execute('select.all', {});
		assert.deepEqual(engine.state().pageRuler, original);
		await engine.execute('caret.docEnd', {});
		await engine.execute('text.newParagraph', {});
		await engine.execute('insert.table', { rows: 2, cols: 2 });
		assert.deepEqual(engine.state().pageRuler, original);
	} finally { engine.dispose(); }
});

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

test('Word paragraph properties preserve block identity, comments, native review and DOCX formatting', async () => {
	const engine = await openEngine('word', 'Paragraph.docx', docx, root);
	try {
		const agent = engine.agent!;
		const read = () => agent.read({ parentId: 'body' }) as { blocks: { id: string; text: string; kind: string }[] };
		const first = read().blocks[0];
		await engine.execute('select.text', { text: 'AI Office' });
		await engine.execute('review.newComment', { text: 'Retain this anchor' });
		first.text = read().blocks[0].text;
		const patch = { indents: { left: 36, right: 12, firstLine: -18 }, spacing: { before: 6, after: 10 }, lineSpacing: { rule: 'exactly', value: 18 }, tabs: [{ pos: 144, align: 'right', leader: 'dot' }], borders: { bottom: { width: 1.5, color: '2255AA', style: 'dashed' } }, shading: 'EEF4FF', keepNext: true, keepLines: true, widowControl: false, pageBreakBefore: false };
		agent.reviewStart!();
		await agent.edit({ revision: engine.state().revision, operation: 'command', id: first.id, command: 'para.properties', params: patch });
		agent.reviewFinish!();
		const current = agent.read({ id: first.id }) as { blocks: { id: string; text: string; paragraph: typeof patch }[] };
		assert.deepEqual({ id: current.blocks[0].id, text: current.blocks[0].text, indent: current.blocks[0].paragraph.indents, spacing: current.blocks[0].paragraph.spacing, tabs: current.blocks[0].paragraph.tabs }, { id: first.id, text: first.text, indent: patch.indents, spacing: patch.spacing, tabs: patch.tabs });
		assert.equal(engine.state().comments?.[0].text, 'Retain this anchor');
		const changes = agent.reviewRead!().changes; assert.ok(changes.length);
		await agent.reviewDecide!(engine.state().revision, 'undo', changes.map(change => change.id));
		assert.equal((agent.read({ id: first.id }) as typeof current).blocks[0].paragraph.indents.left, 0);
		assert.equal(engine.state().comments?.[0].text, 'Retain this anchor');
		await agent.edit({ revision: engine.state().revision, operation: 'command', id: first.id, command: 'para.properties', params: patch });
		await agent.edit({ revision: engine.state().revision, operation: 'command', id: first.id, command: 'para.properties', params: { borders: { top: { color: 'AA5522' } } } });
		const bytes = await engine.export(); const xml = strFromU8(unzipSync(bytes)['word/document.xml']);
		for (const pattern of [/w:ind/, /w:spacing/, /w:tabs/, /w:leader="dot"/, /w:pBdr/, /w:shd/, /w:keepNext/, /w:keepLines/, /w:widowControl/]) { assert.match(xml, pattern); }
		const reopened = await openEngine('word', 'Reopened.docx', bytes, root);
		try {
			const saved = reopened.agent!.read({ id: first.id }) as { blocks: { text: string; paragraph: { indents: object; tabs: object[]; shading: string; lineSpacing: object; keepNext: boolean; widowControl: boolean; borders: { top: object; bottom: object } } }[] };
			assert.deepEqual({ text: saved.blocks[0].text, indents: saved.blocks[0].paragraph.indents, tabs: saved.blocks[0].paragraph.tabs, shading: saved.blocks[0].paragraph.shading, lineSpacing: saved.blocks[0].paragraph.lineSpacing, keepNext: saved.blocks[0].paragraph.keepNext, widow: saved.blocks[0].paragraph.widowControl }, { text: first.text, indents: patch.indents, tabs: patch.tabs, shading: patch.shading, lineSpacing: patch.lineSpacing, keepNext: true, widow: false });
			assert.ok(saved.blocks[0].paragraph.borders.top && saved.blocks[0].paragraph.borders.bottom, 'Patching top border preserves bottom border');
		} finally { reopened.dispose(); }
		await agent.edit({ revision: engine.state().revision, operation: 'command', id: first.id, command: 'para.properties', params: { borders: { bottom: { color: '112233' }, top: null } } });
		const borderRead = agent.read({ id: first.id }) as { blocks: { paragraph: { borders: { bottom: { width: number; color: string; style: string }; top: null } } }[] };
		assert.deepEqual(borderRead.blocks[0].paragraph.borders, { bottom: { width: 1.5, color: '112233', style: 'dashed' }, top: null, left: null, right: null, between: null });
		await agent.edit({ revision: engine.state().revision, operation: 'command', id: first.id, command: 'para.properties', params: { tabs: [] } });
		assert.deepEqual((agent.read({ id: first.id }) as { blocks: { paragraph: { tabs: object[] } }[] }).blocks[0].paragraph.tabs, []);
		await agent.edit({ revision: engine.state().revision, operation: 'insert', id: first.id, position: 'after', blocks: [{ kind: 'table', rows: [['Cell paragraph']] }] });
		const table = read().blocks.find(block => block.kind === 'table')!;
		await assert.rejects(agent.edit({ revision: engine.state().revision, operation: 'command', id: table.id, command: 'para.properties', params: { keepNext: true } }), /require a paragraph/);
		const cellParagraph = (agent.read({ query: 'Cell paragraph' }) as { blocks: { id: string; kind: string }[] }).blocks.find(block => block.kind === 'paragraph')!;
		await agent.edit({ revision: engine.state().revision, operation: 'command', id: cellParagraph.id, command: 'para.properties', params: { spacing: { after: 12 }, keepLines: true } });
		assert.equal((agent.read({ id: cellParagraph.id }) as { blocks: { paragraph: { keepLines: boolean } }[] }).blocks[0].paragraph.keepLines, true);
		const snapshot = engine.state(); const before = agent.reviewProjection!();
		await engine.execute('caret.docEnd', {});
		await assert.rejects(engine.execute('para.properties', { keepNext: false, ifRevision: snapshot.revision, ifSelection: snapshot.selection }), /selection changed/);
		await assert.rejects(agent.edit({ revision: engine.state().revision, operation: 'batch', edits: [{ operation: 'command', id: first.id, command: 'para.properties', params: { keepNext: false } }, { operation: 'command', id: first.id, command: 'para.properties', params: { tabs: [{ pos: -1 }] } }] }));
		assert.equal(agent.reviewProjection!(), before, 'Invalid batch restores all native formatting');
	} finally { engine.dispose(); }
});

test('Word: manual typing and multiline paste remain addressable by guarded agent edits', async () => {
	const engine = await openEngine('word', 'Manual.docx', docx, root);
	type Node = { id: string; kind: string; text: string; parentId: string };
	const read = () => (engine.agent!.read({ limit: 100 }) as { blocks: Node[] }).blocks.map(({ id, kind, text, parentId }) => ({ id, kind, text, parentId }));
	try {
		const original = read();
		await engine.execute('caret.docEnd', {});
		await engine.execute('text.newParagraph', {});
		await engine.execute('text.insert', { text: 'Manually typed 😀' });
		const typed = read().at(-1)!;
		const staleRevision = engine.state().revision;
		await engine.execute('text.insert', { text: ' and continued' });
		assert.deepEqual(read().at(-1), { ...typed, text: 'Manually typed 😀 and continued' });
		await assert.rejects(engine.agent!.edit({ operation: 'update', id: typed.id, revision: staleRevision, text: 'Stale replacement' }), /Document changed/);
		await engine.execute('text.lineBreak', {});
		await engine.execute('text.insert', { text: 'Same paragraph' });
		assert.equal(read().at(-1)!.id, typed.id, 'Shift+Enter keeps the paragraph identity');
		const beforePaste = read();
		await engine.execute('edit.pasteText', { text: '\r\nPasted second 😀\n\nPasted fourth\rLast' });
		const pasted = read();
		assert.deepEqual(pasted.slice(original.length).map(block => block.text), [
			'Manually typed 😀 and continued\nSame paragraph', 'Pasted second 😀', '', 'Pasted fourth', 'Last'
		]);
		assert.equal(new Set(pasted.map(block => block.id)).size, pasted.length);
		assert.equal(pasted[original.length].id, typed.id, 'Paste preserves the destination paragraph identity');
		await engine.execute('edit.undo', {});
		assert.deepEqual(read(), beforePaste, 'A multiline paste is one undo step');
		await engine.execute('edit.redo', {});
		assert.deepEqual(read(), pasted, 'Redo restores the paragraph identities');
		const target = pasted.find(block => block.text === 'Pasted second 😀')!;
		engine.agent!.reviewStart!();
		await engine.agent!.edit({ operation: 'update', id: target.id, revision: engine.state().revision, text: 'Agent edited this paragraph' });
		engine.agent!.reviewFinish!();
		assert.deepEqual(read(), pasted.map(block => block.id === target.id ? { ...block, text: 'Agent edited this paragraph' } : block), 'Agent edit touches only the manually pasted paragraph');
		const reopened = await openEngine('word', 'Manual.docx', await engine.export(), root);
		try { assert.deepEqual((reopened.agent!.read({ limit: 100 }) as { blocks: Node[] }).blocks.map(({ id, kind, text, parentId }) => ({ id, kind, text, parentId })), read(), 'Manual and agent content retain IDs through DOCX save and reopen'); }
		finally { reopened.dispose(); }
		await engine.execute('select.text', { text: 'Pasted fourth' });
		await engine.execute('edit.pasteText', { text: 'Replacement\nNew paragraph' });
		assert.deepEqual(read().slice(-3).map(block => block.text), ['Replacement', 'New paragraph', 'Last'], 'Paste replaces the current selection and preserves following paragraphs');
		const beforeSplit = read();
		await engine.execute('caret.docEnd', {});
		await engine.execute('text.newParagraph', {});
		assert.equal(read().length, beforeSplit.length + 1);
		await engine.execute('text.backspace', {});
		assert.deepEqual(read(), beforeSplit, 'Backspace joins paragraphs without changing surviving IDs');
		await engine.agent!.edit({ operation: 'insert', id: original[0].id, position: 'after', revision: engine.state().revision, blocks: [{ kind: 'table', rows: [['Cell seed', 'Neighbor cell']] }] });
		await engine.execute('select.text', { text: 'Cell seed' });
		await engine.execute('edit.pasteText', { text: 'Cell first\nCell second' });
		const cellParagraphs = read().filter(block => ['Cell first', 'Cell second'].includes(block.text));
		assert.equal(cellParagraphs.length, 2);
		assert.equal(cellParagraphs[0].parentId, cellParagraphs[1].parentId, 'Pasted cell paragraphs stay in the same cell');
		assert.notEqual(cellParagraphs[0].parentId, 'body');
		await engine.agent!.edit({ operation: 'update', id: cellParagraphs[1].id, revision: engine.state().revision, text: 'Agent edited cell paragraph' });
		assert.ok(read().some(block => block.text === 'Neighbor cell'), 'Agent edit preserves the neighboring cell');
	} finally { engine.dispose(); }
});

test('Word headers and footers use page-specific stories, stable blocks, native undo and DOCX parts', async () => {
	const files = unzipSync(docx);
	files['word/document.xml'] = strToU8(strFromU8(files['word/document.xml']).replace('<w:sectPr>', '<w:p><w:r><w:br w:type="page"/></w:r><w:r><w:t>Second page</w:t></w:r></w:p><w:sectPr><w:titlePg/>'));
	const engine = await openEngine('word', 'Headers.docx', zipSync(files), root);
	type Node = { id: string; text: string; parentId: string };
	const read = () => (engine.agent!.read({}) as { blocks: Node[] }).blocks;
	try {
		assert.equal(engine.pages().length, 2);
		const body = read().filter(block => block.parentId === 'body');
		await engine.execute('insert.editHeader', { page: 0 });
		assert.equal(engine.state().editingRegion, 'header');
		await engine.execute('text.insert', { text: 'First page header' });
		const first = read().find(block => block.text === 'First page header')!;
		await engine.execute('insert.closeHeader', {});
		assert.equal(engine.state().editingRegion, 'body');
		await engine.execute('insert.editHeader', { page: 1 });
		await engine.execute('text.insert', { text: 'Following page header' });
		const following = read().find(block => block.text === 'Following page header')!;
		assert.notEqual(first.parentId, following.parentId, 'Different First Page edits the story actually shown on each page');
		const revision = engine.state().revision;
		await engine.execute('insert.editHeader', { page: 0 });
		assert.equal(engine.state().revision, revision, 'Entering an existing header does not create a content revision');
		const editingRaster = engine.render(0, 1).pixels!.slice();
		await engine.execute('insert.closeHeader', {});
		const bodyRaster = engine.render(0, 1).pixels!;
		assert.notDeepEqual(bodyRaster, editingRaster, 'Header editing emphasizes the active story');
		await engine.execute('insert.editHeader', { page: 0 });
		assert.equal(engine.state().caret?.page, 0);
		await engine.execute('insert.editFooter', { page: 0 });
		await engine.execute('text.insert', { text: 'Footer text' });
		const footer = read().find(block => block.text === 'Footer text')!;
		assert.equal(engine.state().editingRegion, 'footer');
		await engine.execute('edit.undo', {});
		assert.ok(!read().some(block => block.text === 'Footer text'));
		await engine.execute('edit.redo', {});
		engine.agent!.reviewStart!();
		await engine.agent!.edit({ operation: 'update', id: footer.id, revision: engine.state().revision, text: 'Agent edited footer' });
		engine.agent!.reviewFinish!();
		const review = engine.agent!.reviewRead!();
		assert.equal(review.changes.length, 1, 'Footer agent edits participate in native Keep/Undo');
		assert.deepEqual(read().filter(block => block.parentId === 'body'), body, 'Header/footer typing and agent edits preserve body blocks');
		await engine.execute('insert.closeHeader', {});
		await assert.rejects(engine.execute('insert.editHeader', { page: 99 }), /page/);
		await assert.rejects(engine.execute('insert.editFooter', { page: -1 }), /page/);
		const bytes = await engine.export();
		const parts = unzipSync(bytes);
		assert.ok(Object.keys(parts).some(name => /word\/header.*\.xml$/.test(name)));
		assert.ok(Object.keys(parts).some(name => /word\/footer.*\.xml$/.test(name)));
		const reopened = await openEngine('word', 'Headers.docx', bytes, root);
		try { assert.deepEqual((reopened.agent!.read({}) as { blocks: Node[] }).blocks, read()); }
		finally { reopened.dispose(); }
	} finally { engine.dispose(); }
});

test('Word overflowing tables expand the editor raster without changing saved paper size', async () => {
	const engine = await openEngine('word', 'Overflow.docx', docx, root);
	try {
		await engine.execute('caret.docEnd', {});
		await engine.execute('text.newParagraph', {});
		await engine.execute('insert.table', { rows: 2, cols: 2, borders: true });
		await engine.execute('table.autofit', { mode: 'fixed' });
		await engine.execute('table.columnWidth', { width: 900 });
		const paper = engine.state().pageRegions![0];
		const display = engine.pages()[0];
		assert.ok(display[0] > paper.width, 'Wide table gets a wider scrollable editor surface');
		const raster = engine.render(0, 1);
		assert.equal(raster.width, Math.ceil(display[0]));
		const tableRight = Math.max(...raster.tables!.flatMap(table => table.cols));
		assert.ok(tableRight > paper.width && tableRight < display[0], 'The final column and resize handle fit inside the raster');
		const documentXml = strFromU8(unzipSync(await engine.export())['word/document.xml']);
		assert.match(documentXml, /w:pgSz[^>]*w:w="12240"/, 'Scrolling does not change the saved page width');
		await engine.execute('edit.undo', {});
		assert.equal(engine.pages()[0][0], paper.width, 'Undo shrinks the editor back to the physical page');
	} finally { engine.dispose(); }
});

test('Word page margins and borders patch native sections atomically and survive DOCX export', async () => {
	const engine = await openEngine('word', 'Page Settings.docx', docx, root);
	try {
		await engine.execute('insert.editHeader', { page: 0 });
		await engine.execute('text.insert', { text: 'Preserve this header' });
		await engine.execute('insert.closeHeader', {});
		const original = engine.agent!.read({});
		const before = engine.state();
		const side = { style: 'double', width: 1.5, color: '286FB5', space: 18 };
		await engine.execute('layout.pageProperties', { page: 0, applyTo: 'section', margins: { top: 36, bottom: 36, left: 54, right: 54, gutter: 6 }, borders: { top: side, bottom: side, left: side, right: side }, ifRevision: before.revision, ifSelection: before.selection });
		const expected = engine.state().pageSetup!;
		assert.deepEqual(expected.margins, { top: 36, bottom: 36, left: 54, right: 54, gutter: 6 });
		assert.deepEqual(expected.borders.top, side);
		assert.deepEqual(engine.agent!.read({}), original, 'Page formatting preserves body and header blocks');
		await engine.execute('edit.undo', {});
		assert.deepEqual(engine.state().pageSetup, before.pageSetup, 'Combined margins and borders are one undo step');
		await engine.execute('edit.redo', {});
		assert.deepEqual(engine.state().pageSetup, expected);
		await engine.execute('select.text', { text: 'AI Office' });
		await assert.rejects(engine.execute('layout.pageProperties', { margins: { top: 40 }, ifRevision: before.revision, ifSelection: before.selection }), /changed/);
		const guarded = engine.state();
		await engine.execute('caret.docEnd', {});
		await assert.rejects(engine.execute('layout.pageProperties', { margins: { top: 40 }, ifRevision: guarded.revision, ifSelection: guarded.selection }), /selection changed/);
		const projection = engine.agent!.reviewProjection!();
		for (const patch of [
			{ margins: { left: 600 }, borders: null }, { margins: { top: -1 } }, { margins: { typo: 10 } },
			{ borders: { top: { color: 'red' } } }, { borders: { top: { style: 'invented' } } },
			{ borders: { top: { width: 20 } } }, { borders: { top: { space: 32 } } }, { applyTo: true, borders: null }
		]) {
			await assert.rejects(engine.execute('layout.pageProperties', patch));
			assert.equal(engine.agent!.reviewProjection!(), projection, 'Invalid page patches are atomic');
		}
		const target = (original as { blocks: { id: string }[] }).blocks[0];
		await engine.agent!.edit({ operation: 'command', id: target.id, revision: engine.state().revision, command: 'layout.pageProperties', params: { borders: { top: { color: 'AA5522' }, left: null } } });
		assert.deepEqual(engine.state().pageSetup!.borders, { top: { ...side, color: 'AA5522' }, bottom: side, left: null, right: side }, 'Agent partial border edits preserve omitted side properties');
		const bytes = await engine.export();
		const xml = strFromU8(unzipSync(bytes)['word/document.xml']);
		assert.match(xml, /w:pgBorders[^>]*w:offsetFrom="page"/, 'Export uses the same page-edge distances as the renderer');
		assert.match(xml, /w:pgMar[^>]*w:top="720"/);
		const reopened = await openEngine('word', 'Page Settings.docx', bytes, root);
		try {
			assert.deepEqual(reopened.state().pageSetup, engine.state().pageSetup);
			assert.deepEqual(reopened.agent!.read({}), original);
		} finally { reopened.dispose(); }
		await engine.execute('layout.pageProperties', { borders: null });
		assert.ok(Object.values(engine.state().pageSetup!.borders).every(side => side === null));
	} finally { engine.dispose(); }
});

test('Word whole-document page settings validate every section before applying a patch', async () => {
	const files = unzipSync(docx);
	let xml = strFromU8(files['word/document.xml']).replace('w:w="12240"', 'w:w="5760"');
	xml = xml.replace('<w:p>', '<w:p><w:pPr><w:sectPr><w:type w:val="nextPage"/><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440"/></w:sectPr></w:pPr>');
	files['word/document.xml'] = strToU8(xml);
	const engine = await openEngine('word', 'Sections.docx', zipSync(files), root);
	try {
		assert.equal(engine.pages().length, 2);
		const before = engine.agent!.reviewProjection!();
		await assert.rejects(engine.execute('layout.pageProperties', { page: 0, applyTo: 'document', margins: { left: 150, right: 150 }, borders: { top: { width: 1 } } }), /Margins leave no room/);
		assert.equal(engine.agent!.reviewProjection!(), before, 'A narrow later section cannot cause a partial update');
		await engine.execute('layout.pageProperties', { page: 0, applyTo: 'document', margins: { left: 36, right: 36 }, borders: { top: { width: 1, color: '286FB5' } } });
		const saved = strFromU8(unzipSync(await engine.export())['word/document.xml']);
		assert.equal(saved.match(/<w:pgBorders/g)?.length, 2);
		assert.equal(saved.match(/w:left="720"/g)?.length, 2);
		await engine.execute('edit.undo', {});
		assert.equal(engine.agent!.reviewProjection!(), before, 'One Undo restores all sections');
	} finally { engine.dispose(); }
});

test('Word: internal IDs, atomic CRUD, Unicode ranges, revision guards and persistence', async () => {
	const engine = await openEngine('word', 'Agent.docx', docx, root);
	type Node = { id: string; kind: string; text: string; parentId: string; nextTextOffset?: number };
	try {
		assert.ok(engine.agent, 'Rebuild the Word WASM engine before running agent tests');
		const guide = await readFile(resolve(root, 'ai-office-pdf/skills/ai-office-docs/commands.md'), 'utf8');
		for (const command of engine.commands()) { assert.ok(guide.includes('`' + command.id + '`'), `Skill lacks coverage for ${command.id}`); }
		const skill = await readFile(resolve(root, 'ai-office-pdf/skills/ai-office-docs/SKILL.md'), 'utf8');
		assert.ok(skill.split(/\s+/).length < 500, 'Keep the always-loaded skill concise');
		const read = (args: Record<string, unknown> = {}) => engine.agent!.read(args) as { blocks: Node[]; nextOffset?: number };
		const edit = (args: Record<string, unknown>) => engine.agent!.edit({ revision: engine.state().revision, ...args });
		const original = read().blocks;
		assert.equal(original.length, 2);
		assert.notEqual(original[0].id, original[1].id);
		const id = original[0].id;
		const inserted = await edit({
			operation: 'insert', id, position: 'before', blocks: [
				{ kind: 'heading', text: 'Agent proposal', style: 'Heading1' },
				{ kind: 'paragraph', text: original[0].text },
				{ kind: 'table', rows: [['Phase', 'Duration'], ['Planning', '1 week']] }
			]
		});
		assert.equal(inserted.agentCreatedBlocks?.length, 3, 'Return only created root IDs for follow-up edits');
		assert.equal(read({ id }).blocks[0].text, original[0].text, 'Insertion must not retarget existing IDs');
		const all = read().blocks;
		assert.equal(new Set(all.map(block => block.id)).size, all.length);
		assert.ok(all.some(block => block.kind === 'heading'));
		assert.ok(all.some(block => block.kind === 'cell'));
		const table = all.find(block => block.kind === 'table')!;
		assert.ok(engine.agent.rects(table.id)!.length > 0);
		await engine.execute('edit.undo', {});
		assert.deepEqual(read().blocks.map(block => block.id), original.map(block => block.id), 'A whole multi-block insert is one undo step');
		await engine.execute('edit.redo', {});
		assert.deepEqual(read().blocks.map(block => block.id), all.map(block => block.id), 'Redo restores the same new identities');
		await edit({ operation: 'update', id, text: 'A😀 café' });
		await edit({ operation: 'update', id, start: 1, end: 2, text: '🌍' });
		assert.equal(read({ id }).blocks[0].text, 'A🌍 café');
		const chunk = read({ id, textLimit: 2 }).blocks[0];
		assert.equal(chunk.text, 'A🌍');
		assert.equal(chunk.nextTextOffset, 2);
		const revision = engine.state().revision;
		await assert.rejects(engine.agent.edit({ operation: 'delete', id, revision: revision - 1 }), /Document changed/);
		await assert.rejects(edit({ operation: 'insert', id, blocks: [{ kind: 'paragraph', text: 'Valid' }, { kind: 'unknown', text: 'Invalid' }] }), /Supported new blocks/);
		await assert.rejects(edit({ operation: 'update', id, start: -1, end: 1, text: 'Invalid' }), /index|offset|range/i);
		await assert.rejects(edit({ operation: 'update', id, start: 4294967296, end: 4294967297, text: 'Invalid' }), /index|offset|range/i);
		assert.equal(engine.state().revision, revision, 'Failures leave revision and content unchanged');
		assert.equal(read({ id }).blocks[0].text, 'A🌍 café');
		await edit({ operation: 'command', id, command: 'format.bold', params: { value: true } });
		assert.equal(read({ id }).blocks[0].id, id);
		const comment = await edit({ operation: 'command', id, command: 'review.newComment', params: { text: 'Check this paragraph' } });
		assert.ok(comment.agentCommandResult && typeof comment.agentCommandResult === 'object' && Object.hasOwn(comment.agentCommandResult, 'id'), 'Targeted commands return follow-up handles');
		await engine.execute('edit.undo', {});
		const child = read({ query: 'Planning' }).blocks[0];
		await edit({ operation: 'update', id: child.id, text: 'Research' });
		assert.equal(read({ id: child.id }).blocks[0].text, 'Research');
		const beforeDelete = read().blocks;
		await edit({ operation: 'delete', id: table.id });
		assert.equal(read({ id: child.id }).blocks.length, 0, 'Deleting a table removes its descendants');
		await engine.execute('edit.undo', {});
		assert.deepEqual(read().blocks, beforeDelete);
		const bytes = await engine.export();
		const reopened = await openEngine('word', 'Agent.docx', bytes, root);
		try { assert.deepEqual(reopened.agent!.read({}), read(), 'IDs and content survive a saved DOCX round-trip'); } finally { reopened.dispose(); }
		const files = unzipSync(bytes);
		files['word/document.xml'] = strToU8(strFromU8(files['word/document.xml']).replace('Research', 'External edit'));
		const external = await openEngine('word', 'Agent.docx', zipSync(files), root);
		try {
			const current = external.agent!.read({}) as { blocks: Node[] };
			assert.ok(current.blocks.every(block => !beforeDelete.some(old => old.id === block.id)), 'External edits invalidate private ID mappings');
		} finally { external.dispose(); }
	} finally { engine.dispose(); }
});

test('Word batches paginated targets and character formatting atomically with undo and redo', async () => {
	const files = unzipSync(docx);
	const paragraphs = Array.from({ length: 26 }, (_, i) => `<w:p><w:r><w:t>\u{1f680} Scope ${i} ? preserve this description</w:t></w:r></w:p>`).join('');
	files['word/document.xml'] = strToU8(`<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body>${paragraphs}<w:sectPr/></w:body></w:document>`);
	const engine = await openEngine('word', 'batch.docx', zipSync(files), root);
	try {
		type Result = { blocks: { id: string }[]; total: number; nextOffset: number | null };
		const targets: { id: string }[] = []; let offset = 0;
		do {
			const page = engine.agent!.read({ query: 'Scope', offset }) as Result;
			assert.equal(page.total, 26); targets.push(...page.blocks); offset = page.nextOffset ?? -1;
		} while (offset >= 0);
		const styles = engine.agent!.read({ operation: 'styles', query: 'Heading1' }) as { styles: { size: number }[] };
		const size = styles.styles[0].size; assert.ok(size > 0);
		const edits = targets.flatMap(({ id }) => [
			{ operation: 'command', id, start: 0, end: 7, command: 'format.underline', params: { value: true } },
			{ operation: 'command', id, start: 0, end: 7, command: 'format.size', params: { size } }
		]);
		const changed = await engine.agent!.edit({ operation: 'batch', edits, revision: engine.state().revision });
		assert.equal(changed.agentAppliedEdits!.length, 52); assert.equal(changed.undoDepth, 1);
		const xml = strFromU8(unzipSync(await engine.export())['word/document.xml']);
		assert.equal((xml.match(/<w:u /g) || []).length, 26);
		assert.equal(xml.split('Scope</w:t>').length - 1, 26);
		assert.equal((xml.match(/preserve this description/g) || []).length, 26);
		await engine.execute('edit.undo', {}); assert.equal(engine.state().redo, true);
		const revision = engine.state().revision;
		await assert.rejects(engine.agent!.edit({ operation: 'batch', revision, edits: [edits[0], { ...edits[1], id: 'missing' }] }), /no longer exists/);
		assert.equal(engine.state().revision, revision); assert.equal(engine.state().redo, true);
		assert.doesNotMatch(strFromU8(unzipSync(await engine.export())['word/document.xml']), /<w:u /);
		await engine.execute('edit.redo', {});
		assert.equal((strFromU8(unzipSync(await engine.export())['word/document.xml']).match(/<w:u /g) || []).length, 26);
		const reopened = await openEngine('word', 'batch.docx', await engine.export(), root);
		try { assert.equal((reopened.agent!.read({ query: 'Scope' }) as Result).total, 26); } finally { reopened.dispose(); }
	} finally { engine.dispose(); }
});

test('Word formatting tool resolves every literal match, pagination, Unicode ranges and style sizes in one batch', async () => {
	const files = unzipSync(docx);
	const paragraphs = Array.from({ length: 26 }, (_, i) => `<w:p><w:r><w:t>${i ? 'milestone [a] with description' : String.fromCodePoint(0x1f9ed) + ' Milestone [A] ' + 'x'.repeat(9000) + ' Milestone [A]'}</w:t></w:r></w:p>`).join('');
	files['word/document.xml'] = strToU8(`<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body>${paragraphs}<w:sectPr/></w:body></w:document>`);
	const engine = await openEngine('word', 'format.docx', zipSync(files), root);
	try {
		const read = async (args: Record<string, unknown>) => engine.agent!.read(args) as Awaited<ReturnType<Parameters<typeof prepareWordFormatting>[1]>>;
		const plan = await prepareWordFormatting({ query: 'Milestone [A]', format: { underline: true, sizeFromStyle: 'Heading 1' } }, read);
		assert.equal(plan.matchedBlocks, 26); assert.equal(plan.matchedRanges, 27); assert.equal(plan.edits.length, 54);
		const style = await read({ operation: 'styles', query: 'Heading1' });
		assert.ok(plan.edits.filter(edit => edit.command === 'format.size').every(edit => edit.params.size === style.styles![0].size));
		const result = await engine.agent!.edit({ operation: 'batch', edits: plan.edits, revision: engine.state().revision });
		assert.equal(result.undoDepth, 1);
		const xml = strFromU8(unzipSync(await engine.export())['word/document.xml']);
		assert.equal((xml.match(/<w:u /g) || []).length, 27);
		await engine.execute('edit.undo', {}); assert.equal(engine.state().redo, true);
		const first = await prepareWordFormatting({ query: 'Milestone [A]', scope: 'first', format: { bold: true } }, read);
		assert.equal(first.edits.length, 1); assert.equal(first.matchedRanges, 1);
		assert.throws(() => validateDocumentEdits({ operation: 'update', id: 'opaque', text: 'unchanged', command: 'format.underline' }), /does not accept command/);
		await assert.rejects(prepareWordFormatting({ query: 'Milestone [A]', format: { unknown: true } }, read), /Unsupported format property/);
		assert.equal(engine.state().redo, true, 'Preparation errors never change document history');
	} finally { engine.dispose(); }
});

test('Word review keeps one paragraph and undoes another without replacing the document', async () => {
	const engine = await openEngine('word', 'review.docx', docx, root);
	try {
		const agent = engine.agent!;
		const original = agent.read({ limit: 20 }) as { blocks: { id: string; text: string }[] };
		const [first, second] = original.blocks;
		agent.reviewStart!();
		await agent.edit({ operation: 'batch', revision: engine.state().revision, edits: [
			{ operation: 'update', id: first.id, text: 'First proposal' },
			{ operation: 'update', id: second.id, text: 'Second proposal' }
		] });
		agent.reviewFinish!();
		const changes = agent.reviewRead!().changes;
		assert.equal(changes.length, 2);
		assert.ok(changes.every(change => !change.before.includes(first.id) && !change.after.includes(second.id)), 'Private block identities are absent from the readable diff');
		const kept = changes.find(change => change.after.includes('First proposal'))!;
		const rejected = changes.find(change => change.after.includes('Second proposal'))!;
		agent.reviewDecide!(engine.state().revision, 'keep', [kept.id]);
		assert.equal(agent.reviewRead!().changes.length, 1);
		const revision = engine.state().revision;
		assert.throws(() => agent.reviewDecide!(revision - 1, 'undo', [rejected.id]), /Document changed/);
		agent.reviewDecide!(revision, 'undo', [rejected.id]);
		assert.equal(agent.reviewRead!().changes.length, 0);
		assert.deepEqual((agent.read({ limit: 20 }) as typeof original).blocks.map(block => block.text), ['First proposal', second.text]);
		await engine.execute('edit.undo', {});
		assert.deepEqual((agent.read({ limit: 20 }) as typeof original).blocks.map(block => block.text), ['First proposal', 'Second proposal'], 'Review Undo participates in normal document history');
	} finally { engine.dispose(); }
});

test('Word review exposes formatting changes and rejects conflicting manual edits atomically', async () => {
	const engine = await openEngine('word', 'review-format.docx', docx, root);
	try {
		const agent = engine.agent!;
		const blocks = (agent.read({ limit: 20 }) as { blocks: { id: string }[] }).blocks;
		agent.reviewStart!();
		await agent.edit({ operation: 'command', revision: engine.state().revision, id: blocks[0].id, command: 'format.bold', params: { value: true } });
		agent.reviewFinish!();
		const change = agent.reviewRead!().changes[0];
		assert.notEqual(change.before, change.after);
		assert.match(change.after, /bold: true/);
		await agent.edit({ operation: 'update', revision: engine.state().revision, id: blocks[0].id, text: 'My later manual content' });
		const revision = engine.state().revision;
		assert.throws(() => agent.reviewDecide!(revision, 'undo', [change.id]), /overwrite your edits/);
		assert.equal(engine.state().revision, revision);
		agent.reviewDecide!(revision, 'keep', [change.id]);
		assert.equal((agent.read({ id: blocks[0].id }) as { blocks: { text: string }[] }).blocks[0].text, 'My later manual content');
		assert.equal(agent.reviewRead!().changes.length, 0);
	} finally { engine.dispose(); }
});

test('Word file review restores inserted and deleted blocks while preserving unrelated user edits', async () => {
	const engine = await openEngine('word', 'review-structure.docx', docx, root);
	try {
		const agent = engine.agent!;
		const original = agent.read({ limit: 20 }) as { blocks: { id: string; text: string }[] };
		agent.reviewStart!();
		await agent.edit({ revision: engine.state().revision, operation: 'batch', edits: [
			{ operation: 'insert', id: original.blocks[0].id, blocks: [{ kind: 'paragraph', text: 'Agent insertion' }] },
			{ operation: 'delete', id: original.blocks[1].id }
		] });
		agent.reviewFinish!();
		const pending = agent.reviewRead!().changes;
		assert.equal(pending.length, 2);
		await agent.edit({ revision: engine.state().revision, operation: 'update', id: original.blocks[0].id, text: 'Unrelated manual content' });
		agent.reviewDecide!(engine.state().revision, 'undo', pending.map(change => change.id));
		assert.deepEqual((agent.read({ limit: 20 }) as typeof original).blocks.map(block => block.text), ['Unrelated manual content', original.blocks[1].text]);
		assert.equal(agent.reviewRead!().changes.length, 0);
		const reopened = await openEngine('word', 'review-structure.docx', await engine.export(), root);
		try { assert.deepEqual((reopened.agent!.read({ limit: 20 }) as typeof original).blocks.map(block => block.text), ['Unrelated manual content', original.blocks[1].text]); }
		finally { reopened.dispose(); }
	} finally { engine.dispose(); }
});

test('Word inline review preserves text, deleted-block anchors and decisions across native reopen', async () => {
	const engine = await openEngine('word', 'persistent-review.docx', docx, root);
	try {
		const agent = engine.agent!, blocks = (agent.read({ limit: 20 }) as { blocks: { id: string; text: string }[] }).blocks;
		agent.reviewStart!();
		await agent.edit({ revision: engine.state().revision, operation: 'batch', edits: [
			{ operation: 'update', id: blocks[0].id, text: 'Revised paragraph 🌍' },
			{ operation: 'delete', id: blocks[1].id }
		] }); agent.reviewFinish!();
		const pending = agent.reviewRead!().changes;
		assert.deepEqual(pending.map(change => [change.beforeText, change.afterText, change.anchor]), [
			[blocks[0].text, 'Revised paragraph 🌍', { id: blocks[0].id, after: false }],
			[blocks[1].text, '', { id: blocks[0].id, after: true }]
		]);
		const reopened = await openEngine('word', 'persistent-review.docx', await engine.export(), root);
		try {
			reopened.agent!.reviewImport!(agent.reviewExport!());
			assert.deepEqual(reopened.agent!.reviewRead!().changes, pending);
			reopened.agent!.reviewDecide!(reopened.state().revision, 'undo', pending.map(change => change.id));
			assert.deepEqual((reopened.agent!.read({ limit: 20 }) as { blocks: { text: string }[] }).blocks.map(block => block.text), blocks.map(block => block.text));
		} finally { reopened.dispose(); }
	} finally { engine.dispose(); }
});

test('Word document versions deduplicate snapshots, retain review only for matching bytes and prune old versions', async () => {
	const directory = await mkdtemp(resolve(tmpdir(), 'office-versions-'));
	try {
		const versions = new DocumentVersions(directory, 'file:///private/memo.docx');
		await versions.capture(Buffer.from('original'), 'Before', 'Before Agent Edits');
		await versions.capture(Buffer.from('proposal'), 'After', 'Agent Edits', { review: '{"private":"native snapshot"}', request: 'request' });
		await versions.capture(Buffer.from('proposal'), 'After', 'Saved Version', { review: '{"private":"native snapshot"}', request: 'request' });
		assert.equal((await versions.list()).length, 2);
		const reopened = new DocumentVersions(directory, 'file:///private/memo.docx');
		assert.equal((await reopened.restoreReview(Buffer.from('proposal')))?.request, 'request');
		assert.equal(await reopened.restoreReview(Buffer.from('external edit')), undefined);
		const original = (await reopened.list()).at(-1)!;
		assert.equal(await reopened.projection(original.id), 'Before');
		await assert.rejects(reopened.projection('../index'), /no longer available/);
		for (let index = 0; index < 31; index++) { await reopened.capture(Buffer.from(`version${index}`), `Text ${index}`, 'Saved Version'); }
		assert.equal((await reopened.list()).length, 30);
		await assert.rejects(reopened.projection(original.id), /no longer available/);
		assert.equal(await reopened.restoreReview(Buffer.from('version30')), undefined);
	} finally { await rm(directory, { recursive: true, force: true }); }
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
		const target = (engine.agent!.read({ limit: 1 }) as { blocks: { id: string }[] }).blocks[0];
		await engine.agent!.edit({ revision: engine.state().revision, operation: 'insert', id: target.id, blocks: [{ kind: 'heading', text: 'Agent imported heading', style: 'Heading2' }] });
		assert.equal(engine.state().paragraphStyle, 'Heading2');
		await engine.execute('edit.undo', {});
		assert.doesNotMatch(strFromU8(unzipSync(await engine.export())['word/styles.xml']), /w:styleId="Heading2"/);
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

test('Word cached typing pixels match full repaint through edits, formatting and undo', async () => {
	const engine = await openEngine('word', 'Raster.docx', docx, root);
	try {
		for (const scale of [1, 1.375, 2]) {
			engine.render(0, scale);
			for (const [command, args] of [
				['caret.docStart', {}], ['text.insert', { text: 'Sharp accents: é Å 漢字 ' }],
				['text.backspace', {}], ['text.newParagraph', {}],
				['text.insert', { text: 'Wrapping text across lines. '.repeat(14) }],
				['edit.undo', {}], ['edit.redo', {}], ['select.all', {}],
				['format.size', { size: 28 }], ['format.bold', { value: true }],
				['format.underline', { value: true }], ['edit.undo', {}],
				['caret.docEnd', {}], ['insert.table', { rows: 2, cols: 2 }], ['text.insert', { text: 'Cell typing' }]
			] as [string, Record<string, unknown>][]) {
				await engine.execute(command, args);
				const cached = engine.render(0, scale);
				// A different scale replaces this page's baseline and forces a complete
				// repaint at the original scale, without changing the document or fonts.
				engine.render(0, scale + .125);
				const full = engine.render(0, scale);
				assert.equal(cached.width, full.width); assert.equal(cached.height, full.height);
				assert.ok(Buffer.from(cached.pixels!).equals(Buffer.from(full.pixels!)), `Cached pixels differ after ${command} at ${scale}x`);
			}
		}
	} finally { engine.dispose(); }
});

test('Word frame cache retains adjacent pages and evicts older baselines', () => {
	const cache = new WordFrame();
	const raster = { pixels: new Uint8Array(400).fill(255), width: 10, height: 10 };
	cache.update(raster, 0, 2, 1); cache.update(raster, 1, 2, 2);
	const changed = { ...raster, pixels: raster.pixels.slice() }; changed.pixels[44] = 0;
	assert.equal(cache.update(changed, 0, 2, 3, 1).height, 1, 'Neighboring page must not force a full transfer');
	cache.update(raster, 2, 2, 4);
	assert.equal(cache.update(raster, 1, 2, 5, 2).height, 10, 'Evicted page must receive a complete raster');
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

test('PDF page operations keep page geometry, native Undo and saved output synchronized', async () => {
	const engine = await openEngine('pdf', 'Pages.pdf', pdfFixture(), root);
	try {
		const rotate = engine.commands().find(command => command.id === 'page_rotate');
		assert.ok(rotate && rotate.params && rotate.readOnly === false);
		assert.equal((rotate.params as { properties: Record<string, unknown> }).properties.doc, undefined, 'Session IDs stay inside the engine');
		await engine.execute('page_insert_blank', { at: 2, width: 400, height: 600 });
		assert.deepEqual(engine.pages(), [[612, 792], [400, 600]]);
		await engine.execute('page_rotate', { pages: [2], degrees: 90 });
		assert.deepEqual(engine.pages()[1], [600, 400]);
		await engine.execute('page_duplicate', { pages: [1] });
		assert.equal(engine.pages().length, 3);
		await engine.execute('page_move', { pages: [3], to: 1 });
		assert.deepEqual(engine.pages()[0], [600, 400]);
		await engine.execute('page_set_box', { pages: [1], box: 'crop', margins: [12, 12, 12, 12] });
		assert.deepEqual(engine.pages()[0], [576, 376]);
		await engine.execute('page_delete', { pages: [2] }); assert.equal(engine.pages().length, 2);
		await engine.execute('edit_undo', {}); assert.equal(engine.pages().length, 3);
		await engine.execute('edit_redo', {}); assert.equal(engine.pages().length, 2);
		const version = engine.state().revision;
		const extracted = await engine.execute('text_extract', { pages: [2] });
		assert.equal(extracted.revision, version, 'Read-only commands never change history');
		assert.match(JSON.stringify(extracted.agentCommandResult), /AI Office editing/);
		const reopened = await openEngine('pdf', 'Saved-pages.pdf', await engine.export(), root);
		try { assert.deepEqual(reopened.pages(), engine.pages()); assert.ok((await reopened.renderAsync!(0, .5)).pixels!.length); } finally { reopened.dispose(); }
		const original = await engine.export();
		await assert.rejects(engine.execute('page_rotate', { pages: [99], degrees: 90 }), /page/i);
		await assert.rejects(engine.execute('doc_watermark', { text: 'X', file: '/secret.png' }), /explicit file:\/\/ URIs/);
		await assert.rejects(engine.execute('doc_save', { path: '/tmp/not-allowed.pdf' }), /managed by the application|not enabled/);
		assert.deepEqual(await engine.export(), original, 'Refused edits leave document bytes unchanged');
	} finally { engine.dispose(); }
});

test('PDF watermarks, printed page numbers and backgrounds survive save and reopen', async () => {
	const engine = await openEngine('pdf', 'Marks.pdf', pdfFixture(), root);
	try {
		await engine.execute('doc_watermark', { pages: [1], text: 'SYNTHETIC DRAFT', opacity: .25, rotation: 45 });
		await engine.execute('doc_header_footer', { pages: [1], footer_center: '<<Page 1 of n>>', font_size: 10 });
		await engine.execute('doc_background', { pages: [1], color: '#fafafa' });
		const reopened = await openEngine('pdf', 'Saved-marks.pdf', await engine.export(), root);
		try {
			const text = JSON.stringify((await reopened.execute('text_extract', { pages: [1] })).agentCommandResult);
			assert.match(text, /SYNTHETIC DRAFT/); assert.match(text, /Page 1 of 1/);
			assert.match(text, /AI Office editing/); assert.ok((await reopened.renderAsync!(0, .5)).pixels!.length);
		} finally { reopened.dispose(); }
	} finally { engine.dispose(); }
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
		try { assert.match(reopened.state().blocks!.map(b => b.text).join(' '), /Editable PDF/); assert.ok(reopened.state().comments?.some(c => c.text === 'PDF review')); assert.ok((await reopened.renderAsync!(0, .5)).pixels!.length > 100); } finally { reopened.dispose(); }
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

test('Agent content commands create, update and delete by text without ids', async () => {
	const word = await openEngine('word', 'Memo.docx', docx, root);
	try {
		await word.execute('content.append', { text: 'Closing note' });
		await word.execute('content.replace', { find: 'AI Office', text: 'Quarterly report' });
		await word.execute('content.delete', { find: 'toolbar' });
		const found = await word.execute('content.find', { find: 'Quarterly report' });
		assert.ok((found.rects || []).length, 'Word find selects the matched text');
		await word.execute('content.replace', { find: 'e', text: 'E', all: true });
		await word.execute('select.all', {});
		const text = word.state().selectedText!;
		assert.match(text, /QuartErly rEport Editing/);
		assert.match(text, /Closing notE/);
		assert.doesNotMatch(text, /toolbar|e/);
		const bytes = await word.export();
		const reopened = await openEngine('word', 'agent.docx', bytes, root);
		try { await reopened.execute('select.all', {}); assert.match(reopened.state().selectedText!, /Closing notE/); } finally { reopened.dispose(); }
	} finally { word.dispose(); }

	const pdf = await openEngine('pdf', 'Agent.pdf', pdfFixture(), root);
	try {
		await pdf.execute('content.replace', { find: 'AI Office', text: 'AI Office team' });
		assert.equal(pdf.state().blocks![0].text, 'AI Office team editing');
		const found = await pdf.execute('content.find', { find: 'team' });
		assert.equal(found.matches?.length, 1);
		await pdf.execute('content.append', { text: 'Signed copy' });
		assert.ok(pdf.state().blocks!.some(block => block.text === 'Signed copy'), 'PDF append adds a separate paragraph');
		await pdf.execute('content.delete', { find: 'Signed copy' });
		assert.ok(!pdf.state().blocks!.some(block => block.text.includes('Signed copy')));
		await pdf.execute('content.delete', { find: 'AI Office team editing' });
		assert.equal(pdf.state().blocks!.length, 0, 'Deleting the full text removes the paragraph');
	} finally { pdf.dispose(); }

	const deck = ppt.createPresentation(); const slide = ppt.addSlide(deck, { layout: ppt.getSlideLayouts(deck)[0] });
	ppt.addSlideTextBox(slide, { x: ppt.inches(1), y: ppt.inches(1), w: ppt.inches(8), h: ppt.inches(1), text: 'Agenda' });
	const slides = await openEngine('slides', 'Agent.pptx', await ppt.savePresentation(deck), root);
	try {
		await slides.execute('content.replace', { find: 'Agenda', text: 'Agenda and notes' });
		assert.ok(slides.state().blocks!.some(block => block.text === 'Agenda and notes'), 'Slide shapes resolve by text');
		await slides.execute('content.append', { text: 'Questions' });
		const initial = slides.state().blocks![0], appended = slides.state().blocks!.find(block => block.text === 'Questions')!;
		assert.ok(appended.rect[3] > initial.rect[3], 'Appended slide text sits below the existing box');
		await slides.execute('content.delete', { find: 'Questions' });
		assert.ok(!slides.state().blocks!.some(block => block.text.includes('Questions')));
	} finally { slides.dispose(); }
});

test('Agent activity regions resolve without caller supplied ids', () => {
	assert.deepEqual(activityRegion('word', {}, { revision: 0 }, { revision: 1, rects: [{ page: 0, x: 10, y: 20, width: 30, height: 8 }] }), { page: 0, rects: [{ x: 10, y: 20, width: 30, height: 8 }] });
	const block = (id: number, text: string, rect: number[]): Block => ({ id, text, rect });
	assert.deepEqual(activityRegion('slides', { shape: 2 }, undefined, { revision: 1, page: 1, blocks: [block(2, 'Agenda', [10, 20, 40, 40])] }), { page: 1, rects: [{ x: 10, y: 20, width: 30, height: 20 }] });
	const before: State = { revision: 0, page: 0, blocks: [block(1, 'Old', [0, 0, 10, 10])] };
	const after: State = { revision: 1, page: 0, blocks: [block(1, 'Old', [0, 0, 10, 10]), block(2, 'New paragraph', [0, 20, 100, 32])] };
	assert.deepEqual(activityRegion('pdf', { text: 'New paragraph' }, before, after), { page: 0, rects: [{ x: 0, y: 20, width: 100, height: 12 }] });
	assert.deepEqual(activityRegion('pdf', {}, before, after), { page: 0, rects: [{ x: 0, y: 20, width: 100, height: 12 }] }, 'Added blocks are reported when args carry no target');
	assert.equal(activityRegion('pdf', {}, before, before), undefined);
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

test('CLI content commands replace text without ids in every format', async () => {
	const cli = resolve('extensions/ai-office-pdf/out/cli.cjs'); await mkdir(demo, { recursive: true });
	for (const [file, text] of [['Memo.docx', 'CLI content Word'], ['Contract.pdf', 'CLI content PDF'], ['Board.pptx', 'CLI content Slides']] as const) {
		const script = resolve(demo, 'content-changes.json'), output = resolve(demo, 'content-' + file);
		await writeFile(script, JSON.stringify([{ command: 'content.replace', params: { find: 'AI Office', text } }]));
		const run = spawnSync(process.execPath, [cli, 'run', resolve(demo, file), '--script', script, '--output', output, '--root', demo], { encoding: 'utf8' });
		assert.equal(run.status, 0, run.stderr); assert.equal(JSON.parse(run.stdout).commands, 1);
		const kind = file.endsWith('docx') ? 'word' : file.endsWith('pdf') ? 'pdf' : 'slides';
		const reopened = await openEngine(kind, file, await readFile(output), root);
		try {
			if (kind === 'word') { await reopened.execute('select.all', {}); assert.match(reopened.state().selectedText!, /CLI content Word/); }
			else { assert.match(reopened.state().blocks!.map(block => block.text).join(' '), /CLI content (PDF|Slides)/); }
		} finally { reopened.dispose(); }
	}
});


test('Word block guards preserve unrelated work and roll back conflicting batches', async () => {
	const engine = await openEngine('word', 'Guards.docx', docx, root);
	type Node = { id: string; text: string; ifHash: string };
	const read = () => (engine.agent!.read({}) as { blocks: Node[] }).blocks;
	const edit = (args: Record<string, unknown>) => engine.agent!.edit({ revision: engine.state().revision, ...args });
	try {
		const [first, second] = read(), revision = engine.state().revision;
		assert.match(first.ifHash, /^[a-f0-9]{64}$/);
		await edit({ operation: 'update', id: second.id, text: 'User changed another paragraph' });
		await edit({ operation: 'replaceText', id: first.id, ifHash: first.ifHash, revision, find: 'AI Office', text: 'Office agent' });
		assert.equal(read()[1].text, 'User changed another paragraph');
		assert.equal(read()[0].text, 'Office agent editing');
		await assert.rejects(edit({ operation: 'update', id: first.id, ifHash: first.ifHash, revision, text: 'Overwrite user' }), /Block changed/);
		await assert.rejects(edit({ operation: 'delete', id: second.id, ifHash: read()[1].ifHash, revision }), /fresh document version/);
		const current = read(), before = engine.state().revision;
		await assert.rejects(edit({ operation: 'batch', edits: [
			{ operation: 'update', id: first.id, ifHash: current[0].ifHash, text: 'Must roll back' },
			{ operation: 'update', id: second.id, ifHash: second.ifHash, text: 'Conflict' }
		] }), /Block changed/);
		assert.equal(engine.state().revision, before); assert.deepEqual(read(), current);
		await engine.execute('select.text', { text: 'Office agent' }); await engine.execute('format.bold', { value: true });
		await assert.rejects(edit({ operation: 'replaceText', id: first.id, ifHash: current[0].ifHash, find: 'editing', text: 'changed' }), /Block changed/);
	} finally { engine.dispose(); }
});

test('Word section outlines include nested headings and custom outline styles without leaking peer sections', async () => {
	const engine = await openEngine('word', 'Outline.docx', docx, root);
	type Node = { id: string; kind: string; text: string; headingLevel: number };
	const read = (args: Record<string, unknown> = {}) => engine.agent!.read(args) as { blocks: Node[]; nextOffset: number | null; summary: unknown };
	try {
		const id = read().blocks[0].id;
		await engine.agent!.edit({ revision: engine.state().revision, operation: 'insert', id, position: 'before', blocks: [
			{ kind: 'heading', text: 'Scope A', style: 'Heading1' }, { kind: 'paragraph', text: 'A body' },
			{ kind: 'heading', text: 'Nested', style: 'Heading2' }, { kind: 'paragraph', text: 'Nested body' },
			{ kind: 'table', rows: [['Only A']] }, { kind: 'heading', text: 'Scope B', style: 'Heading1' }, { kind: 'paragraph', text: 'B body' }
		] });
		const outline = read({ operation: 'outline' });
		assert.deepEqual(outline.blocks.map(n => n.text), ['Scope A', 'Nested', '', 'Scope B']);
		assert.deepEqual(outline.blocks.filter(n => n.kind === 'heading').map(n => n.headingLevel), [1, 2, 1]);
		const sectionId = outline.blocks[0].id, section = read({ sectionId });
		assert.ok(section.blocks.some(n => n.text === 'Nested body')); assert.ok(section.blocks.some(n => n.text === 'Only A'));
		assert.ok(!section.blocks.some(n => n.text.includes('Scope B') || n.text === 'B body'));
		const page = read({ sectionId, limit: 2 }); assert.equal(page.nextOffset, 2);
		assert.equal(read({ sectionId, offset: 2, limit: 2 }).blocks[0].text, 'Nested');
		assert.throws(() => read({ sectionId: id }), /heading/);
		const exported = unzipSync(await engine.export());
		exported['word/document.xml'] = strToU8(strFromU8(exported['word/document.xml']).replace('<w:pPr>', '<w:pPr><w:outlineLvl w:val="2"/>').replace('w:val="Heading1"', 'w:val="CustomHeading"'));
		const reopened = await openEngine('word', 'Custom.docx', zipSync(exported), root);
		try { assert.equal((reopened.agent!.read({ operation: 'outline' }) as { blocks: Node[] }).blocks[0].headingLevel, 3); } finally { reopened.dispose(); }
	} finally { engine.dispose(); }
});

test('Word literal edits preserve formatting and anchored comment threads through DOCX and review', async () => {
	const engine = await openEngine('word', 'Comments.docx', docx, root);
	type Comment = { commentId: number; parentId: number | null; text: string; orphaned: boolean; resolved: boolean; anchor: { start: { id: string; offset: number }; end: { id: string; offset: number } } };
	const read = () => (engine.agent!.read({}) as { blocks: { id: string; text: string }[] }).blocks;
	const comments = () => (engine.agent!.read({ operation: 'comments', textLimit: 4000 }) as { comments: Comment[] }).comments;
	const edit = (args: Record<string, unknown>) => engine.agent!.edit({ revision: engine.state().revision, ...args });
	try {
		const id = read()[0].id;
		await edit({ operation: 'update', id, text: '😀 hello hello end' });
		await assert.rejects(edit({ operation: 'replaceText', id, find: 'hello', text: 'hi' }), /more than once/);
		await edit({ operation: 'replaceText', id, find: 'hello', nth: 2, text: 'world' });
		assert.equal(read()[0].text, '😀 hello world end');
		await engine.execute('select.text', { text: 'hello' }); await engine.execute('format.bold', { value: true });
		const added = await edit({ operation: 'comment', action: 'add', id, find: 'world', text: 'Confirm this word' });
		const rootId = comments()[0].commentId; assert.equal((added.agentCommandResult as { id: number }).id, rootId);
		assert.equal(comments()[0].orphaned, false);
		await edit({ operation: 'comment', action: 'reply', commentId: rootId, text: 'Confirmed' });
		const replyId = comments().find(c => c.parentId === rootId)!.commentId;
		await assert.rejects(edit({ operation: 'comment', action: 'reply', commentId: replyId, text: 'Nested' }), /root threadId/);
		await edit({ operation: 'comment', action: 'update', commentId: replyId, text: 'Updated reply' });
		await edit({ operation: 'comment', action: 'resolve', commentId: rootId, resolved: true });
		assert.equal(comments().find(c => c.commentId === rootId)!.resolved, true);
		const start = comments()[0].anchor.start.offset;
		engine.agent!.reviewStart!();
		await edit({ operation: 'replaceText', id, find: 'end', text: 'finish' });
		engine.agent!.reviewFinish!();
		assert.equal(comments()[0].anchor.start.offset, start);
		assert.equal(comments()[1].text, 'Updated reply');
		assert.ok(engine.agent!.reviewRead!().changes.length);
		await assert.rejects(edit({ operation: 'update', id, text: 'Would detach' }), /inline objects|comment anchors/i);
		await assert.rejects(edit({ operation: 'delete', id }), /detach comment anchors/);
		await assert.rejects(edit({ operation: 'replaceBlocks', id, blocks: [{ kind: 'paragraph', text: 'Replace' }] }), /detach comment anchors/);
		const bytes = await engine.export(), xml = strFromU8(unzipSync(bytes)['word/document.xml']); assert.match(xml, /w:b/); assert.match(xml, /commentRangeStart/);
		const reopened = await openEngine('word', 'Roundtrip.docx', bytes, root);
		try { const cs = (reopened.agent!.read({ operation: 'comments', textLimit: 4000 }) as { comments: Comment[] }).comments; assert.equal(cs.length, 2); assert.ok(cs.every(c => !c.orphaned)); assert.ok(cs.some(c => c.text === 'Updated reply')); } finally { reopened.dispose(); }
		await edit({ operation: 'delete', id, allowDetach: true }); assert.ok(comments().every(c => c.orphaned));
		await engine.execute('edit.undo', {}); assert.ok(comments().every(c => !c.orphaned));
		await edit({ operation: 'comment', action: 'delete', commentId: replyId }); assert.equal(comments().length, 1);
		await edit({ operation: 'comment', action: 'delete', commentId: rootId }); assert.equal(comments().length, 0);
	} finally { engine.dispose(); }
});

test('Word progressive draft intentions never serialize and out-of-order fills retain planned order with Undo', async () => {
	const engine = await openEngine('word', 'Draft.docx', docx, root), draft = new DocumentDrafts();
	const read = () => (engine.agent!.read({ parentId: 'body' }) as { blocks: { id: string; text: string }[] }).blocks;
	try {
		const before = await engine.export(), id = read()[0].id;
		const plan = draft.plan({ id, position: 'after' }, [{ title: 'First section', intent: 'Research A' }, { title: 'Second section', intent: 'Research B' }]);
		assert.equal(plan.pending, 2); assert.deepEqual(await engine.export(), before);
		assert.throws(() => draft.plan({ id, position: 'after' }, [{ title: 'Other', intent: 'Other' }]), /unfinished/);
		for (const section of [plan.sections[1], plan.sections[0]]) {
			const anchor = draft.location(section.id);
			const result = await engine.agent!.edit({ revision: engine.state().revision, operation: 'insert', ...anchor, blocks: [{ kind: 'heading', text: section.title, style: 'Heading2' }, { kind: 'paragraph', text: section.intent }] });
			draft.complete(section.id, result.agentCreatedBlocks!.map(b => b.id));
		}
		assert.deepEqual(read().map(b => b.text).slice(1, 5), ['First section', 'Research A', 'Second section', 'Research B']);
		assert.equal(draft.status().pending, 0); draft.discard(); assert.equal(read().length, 6);
		await engine.execute('edit.undo', {}); assert.ok(!read().some(b => b.text === 'First section')); assert.ok(read().some(b => b.text === 'Second section'));
		const exported = strFromU8(unzipSync(await engine.export())['word/document.xml']); assert.doesNotMatch(exported, /pending|Research A/);
		const remaining = draft.plan({ id, position: 'before' }, [{ title: 'Unwritten', intent: 'Must stay editor-only' }]); draft.discard(remaining.sections[0].id); assert.equal(draft.pending().length, 0);
	} finally { engine.dispose(); }
});


test('Word explicit structured replacement retires IDs and restores the original with Undo', async () => {
	const engine = await openEngine('word', 'Replace.docx', docx, root);
	const read = () => (engine.agent!.read({ parentId: 'body' }) as { blocks: { id: string; kind: string; text: string }[] }).blocks;
	try {
		const original = read(), revision = engine.state().revision;
		const result = await engine.agent!.edit({ revision, operation: 'replaceBlocks', id: original[0].id, blocks: [{ kind: 'heading', text: 'New structure', style: 'Heading1' }, { kind: 'table', rows: [['Phase', 'State'], ['One', 'Ready']] }] });
		assert.equal(result.agentCreatedBlocks?.length, 2); assert.ok(!read().some(b => b.id === original[0].id));
		assert.equal(read()[1].kind, 'table'); assert.equal(read()[2].id, original[1].id);
		await engine.execute('edit.undo', {}); assert.deepEqual(read(), original);
		await engine.execute('edit.redo', {}); assert.equal(read()[0].text, 'New structure');
	} finally { engine.dispose(); }
});


test('Word review undoes progressive fills and their dependent comment as one safe proposal', async () => {
	const engine = await openEngine('word', 'Combined.docx', docx, root);
	const read = () => (engine.agent!.read({ parentId: 'body' }) as { blocks: { id: string; text: string }[] }).blocks;
	try {
		const initial = read(), first = initial[0].id;
		for (const title of ['Goals', 'Rollout']) {
			engine.agent!.reviewStart!(); await engine.agent!.edit({ revision: engine.state().revision, operation: 'insert', id: first, blocks: [{ kind: 'heading', text: title, style: 'Heading2' }, { kind: 'paragraph', text: title + ' body' }] }); engine.agent!.reviewFinish!();
		}
		engine.agent!.reviewStart!(); await engine.agent!.edit({ revision: engine.state().revision, operation: 'comment', action: 'add', id: read().find(b => b.text === 'Goals body')!.id, find: 'Goals', text: 'Check this' }); engine.agent!.reviewFinish!();
		assert.equal(engine.agent!.reviewRead!().changes[0].id, 'document');
		engine.agent!.reviewDecide!(engine.state().revision, 'undo', ['document']);
		assert.deepEqual(read(), initial); assert.equal(engine.state().comments!.length, 0); assert.equal(engine.agent!.reviewRead!().changes.length, 0);
	} finally { engine.dispose(); }
});

test('Word agent color, highlight and strike preserve block identity and round-trip with undo', async () => {
	const engine = await openEngine('word', 'Colors.docx', docx, root);
	try {
		const agent = engine.agent!;
		const read: Parameters<typeof prepareWordFormatting>[1] = async args => agent.read(args) as Awaited<ReturnType<Parameters<typeof prepareWordFormatting>[1]>>;
		const plan = await prepareWordFormatting({ query: 'AI Office', format: { color: 'C00000', highlight: 'yellow', strikethrough: true } }, read);
		const id = plan.edits[0].id;
		await agent.edit({ operation: 'batch', revision: engine.state().revision, edits: plan.edits });
		const xml = strFromU8(unzipSync(await engine.export())['word/document.xml']);
		assert.match(xml, /w:color w:val="C00000"/); assert.match(xml, /w:highlight w:val="yellow"/); assert.match(xml, /w:strike/);
		assert.equal((await read({ id })).blocks?.[0].id, id);
		await engine.execute('edit.undo', {});
		assert.doesNotMatch(strFromU8(unzipSync(await engine.export())['word/document.xml']), /w:highlight|w:strike/);
		await engine.execute('edit.redo', {});
		const clear = await prepareWordFormatting({ query: 'AI Office', format: { clear: true } }, read);
		await agent.edit({ operation: 'batch', revision: engine.state().revision, edits: clear.edits });
		assert.doesNotMatch(strFromU8(unzipSync(await engine.export())['word/document.xml']), /w:highlight|w:strike/);
		await assert.rejects(prepareWordFormatting({ query: 'AI', format: { color: '#BAD' } }, read), /color must/);
		await assert.rejects(prepareWordFormatting({ query: 'AI', format: { highlight: ['yellow'] } }, read), /Unsupported highlight/);
		await assert.rejects(prepareWordFormatting({ query: 'AI', format: { clear: true, bold: true } }, read), /used alone/);
	} finally { engine.dispose(); }
});

test('Word agent refuses repeated deletion of a required trailing paragraph without changing IDs or history', async () => {
	const engine = await openEngine('word', 'Structural paragraph.docx', docx, root);
	try {
		const agent = engine.agent!;
		const read = () => agent.read({ parentId: 'body' }) as { blocks: { id: string; kind: string; text: string }[] };
		const last = read().blocks.at(-1)!;
		const edit = (input: Record<string, unknown>) => agent.edit({ ...input, revision: engine.state().revision });
		await edit({ operation: 'insert', id: last.id, position: 'before', blocks: [{ kind: 'table', rows: [['Owner'], ['Maya']] }] });
		await edit({ operation: 'delete', id: last.id });
		const before = read().blocks; const blank = before.at(-1)!;
		assert.equal(blank.kind, 'paragraph'); assert.equal(blank.text, '');
		const revision = engine.state().revision; const depth = engine.state().undoDepth;
		await assert.rejects(edit({ operation: 'delete', id: blank.id }), /required after a table|empty container/);
		assert.equal(engine.state().revision, revision); assert.equal(engine.state().undoDepth, depth);
		assert.deepEqual(read().blocks.map(block => block.id), before.map(block => block.id));
	} finally { engine.dispose(); }
});

test('Word comment text update and reopening can commit atomically with one Undo step', async () => {
	const engine = await openEngine('word', 'Comment update.docx', docx, root);
	try {
		const agent = engine.agent!;
		const id = (agent.read({ parentId: 'body' }) as { blocks: { id: string }[] }).blocks[0].id;
		const edit = (input: Record<string, unknown>) => agent.edit({ ...input, revision: engine.state().revision });
		await edit({ operation: 'comment', action: 'add', id, text: 'Original comment' });
		const read = () => (agent.read({ operation: 'comments' }) as { comments: { commentId: number; text: string; resolved: boolean }[] }).comments[0];
		const commentId = read().commentId;
		await edit({ operation: 'comment', action: 'resolve', commentId, resolved: true });
		await edit({ operation: 'batch', edits: [
			{ operation: 'comment', action: 'update', commentId, text: 'Updated comment' },
			{ operation: 'comment', action: 'resolve', commentId, resolved: false }
		] });
		assert.deepEqual({ text: read().text, resolved: read().resolved }, { text: 'Updated comment', resolved: false });
		await engine.execute('edit.undo', {});
		assert.deepEqual({ text: read().text, resolved: read().resolved }, { text: 'Original comment', resolved: true });
	} finally { engine.dispose(); }
});

test('Word table previews match native resize and move plans without creating edits', async () => {
	const engine = await openEngine('word', 'Table gestures.docx', docx, root);
	try {
		await engine.execute('caret.docEnd', {}); await engine.execute('text.newParagraph', {});
		await engine.execute('insert.table', { rows: 3, cols: 3, borders: true });
		const geometry = () => engine.render(0, .5).tables![0];
		let table = geometry();
		const from = { page: 0, x: table.cols[1], y: (table.rows[0].top + table.rows[0].bottom) / 2, reach: 3 };
		const before = engine.state();
		const preview = engine.tablePreview!({ ...from, command: 'table.resizeDrag', revision: before.revision, toX: from.x + 20, toY: from.y });
		assert.equal(preview.kind, 'column'); assert.deepEqual(engine.state(), before, 'Preview must not alter selection, revision or history');
		await engine.execute('table.resizeDrag', { ...from, toX: from.x + 20, toY: from.y });
		assert.ok(Math.abs(geometry().cols[1] - preview.rects[0].x) < .1);
		assert.equal(engine.state().undoDepth, before.undoDepth! + 1);
		await engine.execute('edit.undo', {}); assert.deepEqual(geometry().cols, table.cols);
		assert.throws(() => engine.tablePreview!({ ...from, revision: before.revision, command: 'table.resizeDrag', toX: from.x, toY: from.y }), /changed/);
		table = geometry(); const move = { page: 0, x: table.cols[0] - 6, y: table.rows[0].top - 6, reach: 3 };
		const moved = engine.tablePreview!({ ...move, command: 'table.moveDrag', revision: engine.state().revision, toX: move.x + 18, toY: move.y + 24 });
		await engine.execute('table.moveDrag', { ...move, toX: move.x + 18, toY: move.y + 24 });
		assert.ok(Math.abs(geometry().cols[0] - moved.rects[0].x) < .1); assert.ok(Math.abs(geometry().rows[0].top - moved.rects[0].y) < .1);
		await engine.execute('edit.undo', {});
		table = geometry(); const row = table.rows[0];
		const rowArgs = { page: 0, x: (table.cols[0] + table.cols[1]) / 2, y: row.bottom, toX: (table.cols[0] + table.cols[1]) / 2, toY: row.bottom + 18, reach: 3 };
		const rowPreview = engine.tablePreview!({ ...rowArgs, command: 'table.resizeDrag', revision: engine.state().revision });
		await engine.execute('table.resizeDrag', rowArgs); assert.ok(Math.abs(geometry().rows[0].bottom - rowPreview.rects[0].y) < .1);
		await engine.execute('edit.undo', {});
		table = geometry(); const corner = { page: 0, x: table.cols.at(-1)! + 4, y: table.rows.at(-1)!.bottom + 4, reach: 3 };
		const cornerArgs = { ...corner, toX: corner.x + 20, toY: corner.y + 30 };
		const cornerPreview = engine.tablePreview!({ ...cornerArgs, command: 'table.resizeDrag', revision: engine.state().revision });
		await engine.execute('table.resizeDrag', cornerArgs); table = geometry();
		assert.ok(Math.abs(table.cols.at(-1)! - table.cols[0] - cornerPreview.rects[0].width) < .1);
		await engine.execute('edit.undo', {});
		await engine.execute('caret.set', { page: 0, x: table.cols[0] + 10, y: table.rows[0].top + 5 });
		for (const [command, args] of [['table.rowHeight', { height: 30 }], ['table.autofit', { mode: 'fixed' }], ['table.distributeColumns', {}], ['table.distributeRows', {}], ['table.repeatHeader', { value: true }], ['table.borders', { kind: 'outside', width: 1, color: '286FB5' }], ['table.selectRow', {}], ['table.merge', {}]] as [string, Record<string, unknown>][]) { await engine.execute(command, args); }
		table = geometry(); assert.ok(!table.rows[0].edges!.includes(1), 'Merged cells must not expose a draggable internal border');
		await engine.execute('table.split', { columns: 3 });
		const saved = await engine.export(); const reopened = await openEngine('word', 'Saved gestures.docx', saved, root);
		try { const savedTable = reopened.render(0, .5).tables![0]; assert.equal(savedTable.rows.length, 3); assert.equal(savedTable.cols.length, 4); } finally { reopened.dispose(); }
	} finally { engine.dispose(); }
});

test('Word table context commands preserve block identities, native undo and DOCX properties', async () => {
	const engine = await openEngine('word', 'Context table.docx', docx, root);
	try {
		const agent = engine.agent!;
		const first = (agent.read({}) as { blocks: { id: string }[] }).blocks[0];
		await agent.edit({ revision: engine.state().revision, operation: 'insert', id: first.id, position: 'after', blocks: [{ kind: 'table', rows: [['Name', 'Value'], ['Zoe', '2'], ['Amy', '1']] }] });
		const target = (agent.read({ query: 'Zoe' }) as { blocks: { id: string; kind: string }[] }).blocks.find(b => b.kind === 'paragraph')!;
		const command = (id: string, params = {}) => agent.edit({ revision: engine.state().revision, operation: 'command', id: target.id, command: id, params });
		const geometry = engine.render(0, 1).tables!; assert.equal(geometry.length, 1); assert.equal(geometry[0].cols.length, 3); assert.equal(geometry[0].rows.length, 3);
		await command('table.moveAxis', { axis: 'row', forward: true });
		let xml = strFromU8(unzipSync(await engine.export())['word/document.xml']); assert.ok(xml.indexOf('Amy') < xml.indexOf('Zoe'));
		assert.equal((agent.read({ id: target.id }) as { blocks: { text: string }[] }).blocks[0].text, 'Zoe');
		await engine.execute('edit.undo', {}); xml = strFromU8(unzipSync(await engine.export())['word/document.xml']); assert.ok(xml.indexOf('Zoe') < xml.indexOf('Amy'));
		await command('table.moveAxis', { axis: 'column', forward: true });
		xml = strFromU8(unzipSync(await engine.export())['word/document.xml']); assert.ok(xml.indexOf('Value') < xml.indexOf('Name'));
		await engine.execute('edit.undo', {});
		await command('table.sort', { column: 0, header: true, descending: false });
		xml = strFromU8(unzipSync(await engine.export())['word/document.xml']); assert.ok(xml.indexOf('Amy') < xml.indexOf('Zoe'));
		await command('table.columnWidth', { width: 120 }); await command('table.shading', { color: 'D9EAD3' }); await command('table.cellAlign', { value: 'topCenter' });
		xml = strFromU8(unzipSync(await engine.export())['word/document.xml']); assert.match(xml, /D9EAD3/); assert.match(xml, /w:jc w:val="center"/);
		await command('table.insertRowAbove'); assert.equal(engine.render(0, 1).tables![0].rows.length, 4);
		await engine.execute('table.deleteRow', {}); assert.equal(engine.render(0, 1).tables![0].rows.length, 3);
		await command('table.insertColumnRight'); assert.equal(engine.render(0, 1).tables![0].cols.length, 4);
		await engine.execute('table.deleteColumn', {}); assert.equal(engine.render(0, 1).tables![0].cols.length, 3);
		const bytes = await engine.export(); const reopened = await openEngine('word', 'Reopened.docx', bytes, root);
		try { assert.equal((reopened.agent!.read({ id: target.id }) as { blocks: { text: string }[] }).blocks[0].text, 'Zoe'); assert.match(strFromU8(unzipSync(await reopened.export())['word/document.xml']), /D9EAD3/); } finally { reopened.dispose(); }
		await command('table.deleteTable'); assert.equal(engine.render(0, 1).tables!.length, 0);
		await engine.execute('edit.undo', {}); assert.equal(engine.render(0, 1).tables!.length, 1);
	} finally { engine.dispose(); }
});

test('Word images support exact selection, guarded agent edits, layout, crop, captions and DOCX round trips', async () => {
	const engine = await openEngine('word', 'Images.docx', docx, root);
	try {
		await engine.execute('caret.docEnd', {}); await engine.execute('text.newParagraph', {});
		await engine.execute('insert.picture', { data: imageFixture(), width: 144 });
		const image = engine.render(0, 1).images![0]; assert.ok(image);
		const select = async () => { const i = engine.render(0, 1).images![0]; await engine.execute('picture.select', { page: i.page, x: i.x + i.width / 2, y: i.y + i.height / 2, revision: engine.state().revision }); return engine.state().image!; };
		const selected = await select(); assert.ok(selected.width > 0);
		const target = { revision: engine.state().revision, parentId: selected.parentId, offset: selected.offset };
		await engine.execute('picture.size', { width: 108, height: 72, lockAspect: false, _image: target });
		assert.equal(engine.state().image!.width, 108); assert.equal(engine.state().image!.height, 72);
		await assert.rejects(engine.execute('picture.size', { width: 300, _image: target }), /changed/);
		const edit = async (command: string, params: Record<string, unknown> = {}) => {
			const read = engine.agent!.read({ id: image.parentId }) as { blocks: { id: string; images: { offset: number }[] }[] };
			const block = read.blocks[0], offset = block.images[0].offset;
			await engine.agent!.edit({ operation: 'command', id: block.id, revision: engine.state().revision, start: offset, end: offset + 1, command, params });
		};
		await assert.rejects(engine.agent!.edit({ operation: 'command', id: image.parentId, revision: engine.state().revision, command: 'picture.size', params: { width: 300 } }), /start\/end/);
		await edit('picture.crop', { left: .1, top: .05 });
		await edit('arrange.wrap', { wrap: 'square' });
		await edit('arrange.bounds', { x: 100, y: 180, page: 0 });
		await edit('arrange.rotation', { degrees: 90 });
		await edit('object.altText', { text: 'Blue verification square', title: 'Test image', decorative: false });
		await edit('picture.change', { data: imageFixture() });
		await edit('references.caption', { label: 'Figure', text: 'An editable image', position: 'below' });
		const bytes = await engine.export(), files = unzipSync(bytes), xml = strFromU8(files['word/document.xml']);
		assert.match(xml, /Blue verification square/); assert.match(xml, /Test image/); assert.match(xml, /SEQ.*Figure/); assert.match(xml, /An editable image/); assert.match(xml, /srcRect/); assert.match(xml, /wrapSquare/);
		const reopened = await openEngine('word', 'Images.docx', bytes, root);
		try {
			const read = reopened.agent!.read({ id: image.parentId }) as { blocks: { images: { offset: number; width: number; height: number; rotation: number; alt: string; crop: number[] }[] }[] };
			const i = read.blocks[0].images[0]; assert.equal(i.width, 108); assert.equal(i.height, 72); assert.equal(i.alt, 'Blue verification square'); assert.equal(i.rotation, 90); assert.ok(Math.abs(i.crop[0] - .1) < .0001);
			await reopened.agent!.edit({ operation: 'command', id: image.parentId, revision: reopened.state().revision, start: i.offset, end: i.offset + 1, command: 'picture.delete', params: {} });
			assert.equal(reopened.render(0, 1).images!.length, 0); await reopened.execute('edit.undo', {}); assert.equal(reopened.render(0, 1).images!.length, 1);
		} finally { reopened.dispose(); }
	} finally { engine.dispose(); }
});

test('Word append resolves the actual body end beyond pagination and remains atomic through Undo and DOCX', async () => {
	const files = unzipSync(docx);
	const more = Array.from({ length: 73 }, (_, i) => `<w:p><w:r><w:t>Body paragraph ${i}</w:t></w:r></w:p>`).join('');
	files['word/document.xml'] = strToU8(strFromU8(files['word/document.xml']).replace('<w:sectPr>', more + '<w:sectPr>'));
	const engine = await openEngine('word', 'Append.docx', zipSync(files), root);
	type Read = { blocks: { id: string; text: string; parentId: string; ifHash: string }[]; total: number };
	const tail = () => (engine.agent!.read({ operation: 'tail' }) as Read).blocks[0];
	try {
		const first = engine.agent!.read({ limit: 20 }) as Read;
		assert.equal(first.blocks.length, 20); assert.ok(first.total > 20); assert.equal(tail().text, 'Body paragraph 72');
		const revision = engine.state().revision;
		await engine.agent!.edit({ operation: 'append', revision, blocks: [{ kind: 'paragraph', text: 'hello world' }] });
		assert.equal(tail().text, 'hello world'); assert.equal(tail().parentId, 'body');
		await assert.rejects(engine.agent!.edit({ operation: 'append', revision, blocks: [{ text: 'Stale append' }] }), /fresh/);
		await assert.rejects(engine.agent!.edit({ operation: 'append', revision: engine.state().revision, id: first.blocks[19].id, blocks: [{ text: 'Wrong anchor' }] }), /Omit id/);
		await engine.agent!.edit({ operation: 'batch', revision: engine.state().revision, edits: [{ operation: 'append', blocks: [{ kind: 'paragraph', text: 'second' }] }, { operation: 'append', blocks: [{ kind: 'paragraph', text: 'third' }] }] });
		assert.equal(tail().text, 'third');
		await engine.execute('edit.undo', {}); assert.equal(tail().text, 'hello world');
		await engine.execute('edit.redo', {}); assert.equal(tail().text, 'third');
		assert.deepEqual((engine.agent!.read({ limit: 20 }) as Read).blocks.map(b => b.ifHash), first.blocks.map(b => b.ifHash));
		const final = tail(); const reopened = await openEngine('word', 'Append.docx', await engine.export(), root);
		try { const last = (reopened.agent!.read({ operation: 'tail' }) as Read).blocks[0]; assert.equal(last.text, 'third'); assert.equal(last.id, final.id); } finally { reopened.dispose(); }
	} finally { engine.dispose(); }
});

test('Word concurrent agent request contexts retain their own identity and label across awaits', async () => {
	const seen: string[] = [];
	await Promise.all(['A', 'B'].map((id, i) => withAgentRequest(id, async () => {
		for (let n = 0; n < 4; n++) {
			await new Promise(resolve => setTimeout(resolve, i + 1));
			assert.equal(activeAgentRequest(), id); assert.equal(activeAgentLabel(), id + ' agent'); seen.push(id);
		}
		await assert.rejects(withAgentRequest('nested', async () => { assert.equal(activeAgentRequest(), 'nested'); throw new Error('nested failure'); }), /nested failure/);
		assert.equal(activeAgentRequest(), id);
	}, id + ' agent')));
	assert.equal(seen.length, 8); assert.equal(activeAgentRequest(), undefined); assert.equal(activeAgentLabel(), 'Agent');
});

test('Word review groups consecutive insertion runs at separate anchors without merging paragraph identities', async () => {
	const engine = await openEngine('word', 'Grouped review.docx', docx, root);
	try {
		const agent = engine.agent!, original = (agent.read({ parentId: 'body', limit: 20 }) as { blocks: { id: string; text: string }[] }).blocks;
		agent.reviewStart!();
		await agent.edit({ operation: 'batch', revision: engine.state().revision, edits: [
			{ operation: 'insert', id: original[0].id, position: 'after', blocks: ['Discovery', 'Design', 'Delivery', 'Review'].map(text => ({ kind: 'paragraph', text })) },
			{ operation: 'insert', id: original[1].id, position: 'after', blocks: ['Owner', 'Deadline', 'Status'].map(text => ({ kind: 'paragraph', text })) }
		] });
		agent.reviewFinish!();
		let changes = agent.reviewRead!().changes;
		assert.equal(changes.length, 7);
		assert.equal(new Set(changes.map(change => change.id)).size, 7, 'Every paragraph retains its own identity');
		assert.equal(new Set(changes.map(change => change.groupId)).size, 2, 'Existing content divides meaningful insertion groups');
		assert.equal(changes.filter(change => change.groupId === changes[0].groupId).length, 4);
		const rejected = changes.filter(change => change.groupId === changes[0].groupId).map(change => change.id);
		agent.reviewDecide!(engine.state().revision, 'undo', rejected);
		changes = agent.reviewRead!().changes;
		assert.equal(changes.length, 3);
		const content = (agent.read({ parentId: 'body', limit: 20 }) as { blocks: { id: string; text: string }[] }).blocks;
		assert.deepEqual(content.filter(block => original.some(before => before.id === block.id)), original);
		assert.ok(content.some(block => block.text === 'Owner'));
		assert.ok(!content.some(block => block.text === 'Discovery'));
		await engine.execute('edit.undo', {});
		assert.equal((agent.read({ parentId: 'body', limit: 20 }) as { blocks: { id: string }[] }).blocks.length, original.length + 7, 'Native Undo restores every paragraph in the rejected group');
		const restored = await openEngine('word', 'Grouped review.docx', await engine.export(), root);
		try {
			const after = (restored.agent!.read({ parentId: 'body', limit: 20 }) as { blocks: { id: string; text: string }[] }).blocks;
			assert.deepEqual(after.map(block => [block.id, block.text]), (agent.read({ parentId: 'body', limit: 20 }) as { blocks: { id: string; text: string }[] }).blocks.map(block => [block.id, block.text]));
		} finally { restored.dispose(); }
	} finally { engine.dispose(); }
});

const wordParityFormats: [string, Record<string, unknown>, RegExp][] = [
	['doubleUnderline', {}, /w:u[^>]*double/], ['wordUnderline', {}, /w:u[^>]*words/],
	['border', { value: true }, /w:bdr/], ['doubleStrikethrough', { value: true }, /w:dstrike/],
	['subscript', { value: true }, /w:vertAlign[^>]*subscript/], ['superscript', { value: true }, /w:vertAlign[^>]*superscript/],
	['allCaps', { value: true }, /w:caps/], ['smallCaps', { value: true }, /w:smallCaps/], ['hidden', { value: true }, /w:vanish/],
	['outline', { value: true }, /w:outline/], ['shadow', { value: true }, /w:shadow/], ['emboss', { value: true }, /w:emboss/], ['engrave', { value: true }, /w:imprint/],
	['growFont', {}, /w:sz/], ['shrinkFont', {}, /w:sz/], ['growFont1', {}, /w:sz/], ['shrinkFont1', {}, /w:sz/],
	['shading', { color: 'DDEEFF' }, /w:shd[^>]*DDEEFF/], ['changeCase', { mode: 'upper' }, /AI OFFICE EDITING/],
	['resetChar', {}, /AI Office editing/], ['spacing', { points: 1.5 }, /w:spacing[^>]*30/], ['scale', { percent: 125 }, /w:w[^>]*125/],
	['position', { points: 2 }, /w:position[^>]*4/], ['charStyle', { style: 'Emphasis' }, /w:rStyle[^>]*Emphasis/],
	['set', { props: { bold: true, spacing: 1.25 } }, /w:spacing[^>]*25/],
	['enclose', { shape: 'circle' }, / eq /i], ['phonetic', { ruby: 'ei' }, /w:ruby/], ['textEffects', { style: 'shadowed' }, /w14:shadow/]
];
for (const [name, params, expected] of wordParityFormats) {
	test(`Word parity format.${name} targets a block, survives DOCX reopen and remains undoable`, async () => {
		const engine = await openEngine('word', 'Formatting parity.docx', docx, root);
		try {
			const agent = engine.agent!;
			const before = (agent.read({ parentId: 'body' }) as { blocks: WordAgentBlock[] }).blocks;
			const id = before[0].id;
			assert.ok(engine.commands().some(command => command.id === 'format.' + name));
			if (name === 'resetChar') { await agent.edit({ revision: engine.state().revision, operation: 'command', id, command: 'format.bold', params: { value: true } }); }
			const original = (agent.read({ parentId: 'body' }) as { blocks: WordAgentBlock[] }).blocks;
			const revision = engine.state().revision;
			const args = { operation: 'command', revision, id, command: 'format.' + name, params, ...(name === 'enclose' || name === 'phonetic' ? { start: 0, end: name === 'enclose' ? 1 : 2 } : {}) };
			await agent.edit(args);
			assert.ok(engine.state().revision > revision);
			const current = (agent.read({ parentId: 'body' }) as { blocks: WordAgentBlock[] }).blocks;
			assert.equal(current[0].id, id);
			assert.deepEqual(current[1], before[1], 'The neighboring paragraph stays untouched');
			const saved = await engine.export(), xml = strFromU8(unzipSync(saved)['word/document.xml']);
			assert.match(xml, expected);
			const reopened = await openEngine('word', 'Reopened.docx', saved, root);
			try {
				const content = (blocks: WordAgentBlock[]) => blocks.map(block => Object.fromEntries(Object.entries(block).filter(([key]) => key !== 'ifHash')));
				assert.deepEqual(content((reopened.agent!.read({ parentId: 'body' }) as { blocks: WordAgentBlock[] }).blocks), content(current));
				assert.equal(strFromU8(unzipSync(await reopened.export())['word/document.xml']), xml, 'Native DOCX formatting is stable after import normalization; hashes belong to a fresh read');
			} finally { reopened.dispose(); }
			await assert.rejects(agent.edit(args), /changed|version|guard/i);
			await engine.execute('edit.undo', {});
			assert.deepEqual((agent.read({ parentId: 'body' }) as { blocks: WordAgentBlock[] }).blocks, original);
			await engine.execute('edit.redo', {});
			assert.deepEqual((agent.read({ parentId: 'body' }) as { blocks: WordAgentBlock[] }).blocks, current);
		} finally { engine.dispose(); }
	});
}

test('Word phonetic commands require explicit content instead of opening an upstream dialog', async () => {
	const engine = await openEngine('word', 'Phonetic.docx', docx, root);
	try {
		const agent = engine.agent!, revision = engine.state().revision;
		const id = (agent.read({ parentId: 'body' }) as { blocks: WordAgentBlock[] }).blocks[0].id;
		await assert.rejects(agent.edit({ revision, operation: 'command', id, command: 'format.phonetic', params: {} }), /params.ruby/);
		assert.equal(engine.state().revision, revision);
		assert.ok(!engine.commands().some(command => command.id === 'format.fontDialog'));
	} finally { engine.dispose(); }
});

test('Word command inventory matches the built bridge and distinguishes pending reference commands', async () => {
	const engine = await openEngine('word', 'Catalog.docx', docx, root);
	try {
		const audit = JSON.parse(await readFile(resolve(root, 'ai-office-pdf/WORD_CAPABILITIES.json'), 'utf8')) as { commands: { command: string; integration: string }[] };
		assert.equal(new Set(audit.commands.map(command => command.command)).size, audit.commands.length);
		assert.deepEqual(engine.commands().filter(command => !command.id.startsWith('content.')).map(command => command.id).sort(), audit.commands.filter(command => command.integration !== 'pending').map(command => command.command).sort());
		assert.ok(audit.commands.some(command => command.command === 'format.fontDialog' && command.integration === 'pending'));
		assert.ok(audit.commands.some(command => command.command === 'file.save' && command.integration === 'pending'));
	} finally { engine.dispose(); }
});

test('Word extended formatting batches preserve comments and use one review and native Undo step', async () => {
	const engine = await openEngine('word', 'Batch formatting.docx', docx, root);
	try {
		const agent = engine.agent!;
		await engine.execute('select.text', { text: 'AI Office' });
		await engine.execute('review.newComment', { text: 'Keep this comment' });
		const original = (agent.read({ parentId: 'body' }) as { blocks: WordAgentBlock[] }).blocks;
		const id = original[0].id;
		const edit = (command: string, params: Record<string, unknown>) => ({ operation: 'command', id, start: 0, end: 9, command, params });
		const edits = [edit('format.smallCaps', { value: true }), edit('format.spacing', { points: 1.5 }), edit('format.textEffects', { style: 'shadowed' })];
		agent.reviewStart!();
		await agent.edit({ revision: engine.state().revision, operation: 'batch', edits });
		agent.reviewFinish!();
		const changed = (agent.read({ parentId: 'body' }) as { blocks: WordAgentBlock[] }).blocks;
		assert.equal(agent.reviewRead!().changes.length, 1);
		assert.equal(engine.state().comments?.[0].text, 'Keep this comment');
		await engine.execute('edit.undo', {});
		assert.deepEqual((agent.read({ parentId: 'body' }) as { blocks: WordAgentBlock[] }).blocks, original);
		await engine.execute('edit.redo', {});
		assert.deepEqual((agent.read({ parentId: 'body' }) as { blocks: WordAgentBlock[] }).blocks, changed);
		await agent.reviewDecide!(engine.state().revision, 'undo', agent.reviewRead!().changes.map(change => change.id));
		assert.deepEqual((agent.read({ parentId: 'body' }) as { blocks: WordAgentBlock[] }).blocks, original);
		assert.equal(engine.state().comments?.[0].text, 'Keep this comment');
		const revision = engine.state().revision;
		await assert.rejects(agent.edit({ revision, operation: 'batch', edits: [edits[0], edit('format.enclose', { shape: 'invalid' })] }));
		assert.equal(engine.state().revision, revision);
		assert.deepEqual((agent.read({ parentId: 'body' }) as { blocks: WordAgentBlock[] }).blocks, original);
	} finally { engine.dispose(); }
});

test('Word effective formatting inspection is Unicode-aware and never changes selection or history', async () => {
	const engine = await openEngine('word', 'Read formatting.docx', docx, root);
	try {
		const agent = engine.agent!;
		const id = (agent.read({ parentId: 'body' }) as { blocks: WordAgentBlock[] }).blocks[0].id;
		await agent.edit({ revision: engine.state().revision, operation: 'update', id, text: '😀 AI Office unchanged' });
		await agent.edit({ revision: engine.state().revision, operation: 'batch', edits: [
			{ operation: 'command', id, start: 2, end: 11, command: 'format.smallCaps', params: { value: true } },
			{ operation: 'command', id, start: 2, end: 11, command: 'format.spacing', params: { points: 1.5 } }
		] });
		await engine.execute('caret.docEnd', {});
		const state = engine.state(), revision = state.revision;
		const inspect = (start: number, end: number) => agent.read({ operation: 'formatting', id, start, end }) as { formatting: { smallCaps: boolean; spacing: number | null } };
		assert.deepEqual({ smallCaps: inspect(2, 11).formatting.smallCaps, spacing: inspect(2, 11).formatting.spacing }, { smallCaps: true, spacing: 1.5 });
		assert.equal(inspect(12, 21).formatting.smallCaps, false);
		assert.equal(inspect(0, 21).formatting.spacing, null, 'A mixed range reports null rather than a guessed spacing');
		assert.deepEqual(engine.state(), state, 'Read-only inspection leaves caret, revision and native Undo untouched');
		assert.throws(() => agent.read({ operation: 'formatting', id, start: 50, end: 55 }), /range/);
		assert.throws(() => agent.read({ operation: 'formatting', id, start: 4, end: 3 }), /range/);
		assert.throws(() => agent.read({ operation: 'formatting' }), /paragraph ID/);
		assert.equal(engine.state().revision, revision);
		assert.match(String(engine.commands().find(command => command.id === 'format.smallCaps')?.params), /value/);
	} finally { engine.dispose(); }
});

const paragraphParityCases: [string, Record<string, unknown>, RegExp, string?][] = [
	['alignLeft', {}, /w:jc[^>]*left/], ['alignCenter', {}, /w:jc[^>]*center/], ['alignRight', {}, /w:jc[^>]*right/], ['justify', {}, /w:jc[^>]*both/], ['distribute', {}, /w:jc[^>]*distribute/],
	['rtl', { value: true }, /w:bidi/], ['ltr', { value: true }, /w:bidi/, 'rtl'],
	['indent', {}, /w:left="720"/], ['outdent', {}, /w:left="0"/, 'indent'], ['hangingIndent', {}, /w:hanging="720"/], ['removeHanging', {}, /w:firstLine="0"/, 'hangingIndent'],
	['indents', { left: 24, right: 12, firstLine: -6 }, /w:hanging="120"/], ['single', {}, /w:line="240"/], ['double', {}, /w:line="480"/], ['oneAndHalf', {}, /w:line="360"/],
	['spacing', { before: 6, after: 9 }, /w:after="180"/], ['addSpaceBefore', {}, /w:before="240"/], ['removeSpaceAfter', {}, /w:after="0"/], ['reset', {}, /AI Office editing/, 'hangingIndent'],
	['heading1', {}, /w:pStyle[^>]*Heading1/], ['heading2', {}, /w:pStyle[^>]*Heading2/], ['heading3', {}, /w:pStyle[^>]*Heading3/], ['normal', {}, /w:pStyle[^>]*Normal/, 'heading1'],
	['multilevel', { kind: 'legal' }, /w:numPr/], ['listLevel', { level: 1 }, /w:ilvl[^>]*1/, 'numbering'], ['restartNumbering', {}, /w:numPr/, 'numbering'],
	['moveUp', {}, /AI Office editing/], ['moveDown', {}, /AI Office editing/], ['promote', {}, /w:pStyle[^>]*Heading1/, 'heading2'], ['demote', {}, /w:pStyle[^>]*Heading2/, 'heading1'],
	['shading', { color: 'DDEEFF' }, /w:shd[^>]*DDEEFF/], ['borders', { kind: 'bottom', width: 1, color: '2255AA' }, /w:bottom[^>]*2255AA/],
	['keepNext', { value: true }, /w:keepNext/], ['keepLines', { value: true }, /w:keepLines/], ['pageBreakBefore', { value: true }, /w:pageBreakBefore/], ['widowControl', { value: false }, /w:widowControl/],
	['asianTypography', { kinsoku: false, autoSpaceDE: false }, /w:kinsoku/], ['tabs', { tabs: [{ pos: 144, align: 'right', leader: 'dot' }] }, /w:tab[^>]*w:pos="2880"/],
	['outlineLevel', { level: 2 }, /w:outlineLvl[^>]*2/], ['set', { props: { keepNext: true, indentLeft: 24 } }, /w:left="480"/],
	['defineBullet', { char: '◆' }, /w:numPr/], ['defineNumber', { format: 'upperLetter', text: '%1)', start: 3 }, /w:numPr/], ['setNumberingValue', { value: 7 }, /w:numPr/, 'numbering']
];
for (const [name, params, expected, prepare] of paragraphParityCases) {
	test(`Word paragraph parity para.${name} preserves identities, scoped edits and native history`, async () => {
		const engine = await openEngine('word', 'Paragraph parity.docx', docx, root);
		try {
			const agent = engine.agent!, blocks = () => (agent.read({ parentId: 'body' }) as { blocks: WordAgentBlock[] }).blocks;
			const initial = blocks(), id = initial[name === 'moveUp' ? 1 : 0].id;
			if (prepare) { await agent.edit({ revision: engine.state().revision, operation: 'command', id, command: 'para.' + prepare, params: {} }); }
			const before = blocks(), revision = engine.state().revision;
			await agent.edit({ revision, operation: 'command', id, command: 'para.' + name, params });
			assert.ok(engine.state().revision > revision);
			const current = blocks(); assert.ok(current.some(block => block.id === id));
			assert.deepEqual(current.find(block => block.id !== id), before.find(block => block.id !== id));
			if (name === 'moveUp') { assert.equal(current[0].id, id); }
			if (name === 'moveDown') { assert.equal(current[1].id, id); }
			const bytes = await engine.export(), xml = strFromU8(unzipSync(bytes)['word/document.xml']); assert.match(xml, expected);
			const reopened = await openEngine('word', 'Reopened.docx', bytes, root);
			try { assert.deepEqual((reopened.agent!.read({ parentId: 'body' }) as { blocks: WordAgentBlock[] }).blocks.map(({ id, text }) => ({ id, text })), current.map(({ id, text }) => ({ id, text }))); assert.match(strFromU8(unzipSync(await reopened.export())['word/document.xml']), expected); } finally { reopened.dispose(); }
			await engine.execute('edit.undo', {}); assert.deepEqual(blocks(), before);
			await engine.execute('edit.redo', {}); assert.deepEqual(blocks(), current);
		} finally { engine.dispose(); }
	});
}

for (const action of ['create', 'modify', 'updateToMatch', 'delete', 'setVisibility', 'addToGallery', 'rename', 'inspectorClear']) {
	test(`Word style parity styles.${action} preserves block IDs, comments and shared review`, async () => {
		const engine = await openEngine('word', 'Style parity.docx', docx, root);
		try {
			const agent = engine.agent!, blocks = () => (agent.read({ parentId: 'body' }) as { blocks: WordAgentBlock[] }).blocks;
			const id = blocks()[0].id;
			await engine.execute('select.text', { text: 'AI Office' }); await engine.execute('review.newComment', { text: 'Preserve this style comment' });
			const run = (command: string, params: Record<string, unknown>) => agent.edit({ revision: engine.state().revision, operation: 'command', id, command, params });
			const created = await run('styles.create', { name: 'Parity Style', fromSelection: true });
			const style = (created.agentCommandResult as { id: string }).id;
			assert.ok(style);
			await run('format.bold', { value: true });
			await run('para.indents', { left: 24 });
			const params: Record<string, Record<string, unknown>> = {
				create: { name: 'Second Parity', fromSelection: true }, modify: { style, chr: { italic: true, size: 14 }, para: { align: 'center', spaceBefore: 7 } },
				updateToMatch: { style }, delete: { style }, setVisibility: { style, gallery: false, hidden: true }, addToGallery: { style, value: false }, rename: { style, name: 'Renamed Parity' }, inspectorClear: { level: 'paragraphFormatting' }
			};
			const before = blocks(); const previous = unzipSync(await engine.export());
			agent.reviewStart!(); await run('styles.' + action, params[action]); agent.reviewFinish!();
			const current = blocks(); assert.deepEqual(current.map(block => block.id), before.map(block => block.id));
			assert.equal(engine.state().comments?.[0].text, 'Preserve this style comment');
			const bytes = await engine.export(), files = unzipSync(bytes), styles = strFromU8(files['word/styles.xml']);
			if (action === 'rename') { assert.match(styles, /Renamed Parity/); }
			if (action === 'create') { assert.match(styles, /Second Parity/); }
			if (action === 'modify') { assert.match(styles, /w:sz[^>]*28/); assert.match(styles, /w:jc[^>]*center/); }
			if (action === 'delete') { assert.doesNotMatch(styles, /w:name[^>]*Parity Style/); }
			if (action === 'setVisibility') { assert.match(styles, /w:semiHidden/); }
			if (action === 'updateToMatch') { assert.match(styles, /w:b(?:\s|\/|>)/); }
			if (action === 'inspectorClear') { assert.doesNotMatch(strFromU8(files['word/document.xml']), /w:left="480"/); }
			assert.ok(agent.reviewRead!().changes.length);
			const reopened = await openEngine('word', 'Reopened.docx', bytes, root);
			try { assert.deepEqual((reopened.agent!.read({ parentId: 'body' }) as { blocks: WordAgentBlock[] }).blocks.map(({ id, text }) => ({ id, text })), current.map(({ id, text }) => ({ id, text }))); } finally { reopened.dispose(); }
			await engine.execute('edit.undo', {}); assert.deepEqual(blocks(), before);
			await engine.execute('edit.redo', {}); assert.deepEqual(blocks(), current);
			await agent.reviewDecide!(engine.state().revision, 'undo', agent.reviewRead!().changes.map(change => change.id));
			assert.deepEqual(blocks(), before);
			assert.equal(strFromU8(unzipSync(await engine.export())['word/styles.xml']), strFromU8(previous['word/styles.xml']));
		} finally { engine.dispose(); }
	});
}

test('Word multilevel lists and style inspection are native, bounded and read-only', async () => {
	const engine = await openEngine('word', 'List parity.docx', docx, root);
	try {
		const agent = engine.agent!, id = (agent.read({ parentId: 'body' }) as { blocks: WordAgentBlock[] }).blocks[0].id;
		await agent.edit({ revision: engine.state().revision, operation: 'command', id, command: 'list.define', params: { name: 'Parity List', levels: [{ format: 'upperLetter', text: '%1)', start: 3, alignedAt: 0, indentAt: 24, follow: 'space', bold: true }, { format: 'lowerRoman', text: '%1.%2)', indentAt: 48 }], apply: true } });
		const state = engine.state();
		const read = () => agent.read({ operation: 'inspect', id, command: 'list.get' }) as { result: { num: number; levels: { format: string; text: string; start: number; bold: boolean }[] } };
		const list = read().result; assert.equal(list.levels.length, 9); assert.equal(list.levels[0].format, 'upperLetter'); assert.equal(list.levels[0].start, 3); assert.equal(list.levels[1].format, 'lowerRoman');
		const inspector = agent.read({ operation: 'inspect', id, command: 'styles.inspect' }) as { result: { paragraph: { direct: unknown[] } } };
		assert.ok(inspector.result.paragraph); assert.doesNotMatch(JSON.stringify(inspector.result), /officeId/);
		assert.deepEqual(engine.state(), state);
		const bytes = await engine.export(); assert.match(strFromU8(unzipSync(bytes)['word/numbering.xml']), /Parity List/);
		const reopened = await openEngine('word', 'Reopened.docx', bytes, root);
		try { assert.deepEqual((reopened.agent!.read({ operation: 'inspect', id, command: 'list.get' }) as { result: typeof list }).result, list); } finally { reopened.dispose(); }
		await assert.rejects(agent.edit({ revision: engine.state().revision, operation: 'command', id, command: 'list.get', params: {} }), /operation=inspect/);
		assert.throws(() => agent.read({ operation: 'inspect', id, command: 'styles.delete', params: { style: 'Normal' } }), /mutations require guarded edits/);
		assert.deepEqual(engine.state(), state);
	} finally { engine.dispose(); }
});

test('Word numbering restarts include only the contiguous list and preserve later content', async () => {
	const engine = await openEngine('word', 'Numbering scope.docx', docx, root);
	try {
		const agent = engine.agent!, read = () => (agent.read({ parentId: 'body' }) as { blocks: WordAgentBlock[] }).blocks;
		const ids = read().map(block => block.id);
		for (const id of ids) { await agent.edit({ revision: engine.state().revision, operation: 'command', id, command: 'para.numbering', params: {} }); }
		await agent.edit({ revision: engine.state().revision, operation: 'append', blocks: [{ kind: 'paragraph', text: 'Outside the list' }] });
		const before = read(), revision = engine.state().revision;
		const list = (id: string) => (agent.read({ operation: 'inspect', id, command: 'list.get' }) as { result: { num: number } }).result.num;
		assert.equal(list(ids[0]), list(ids[1]));
		await agent.edit({ revision, operation: 'command', id: ids[0], command: 'para.setNumberingValue', params: { value: 7 } });
		assert.equal(list(ids[0]), list(ids[1]));
		assert.deepEqual(read().at(-1), before.at(-1));
		assert.deepEqual(read().map(block => block.id), before.map(block => block.id));
		const xml = strFromU8(unzipSync(await engine.export())['word/numbering.xml']); assert.match(xml, /w:startOverride[^>]*7/);
		await engine.execute('edit.undo', {}); assert.deepEqual(read(), before);
	} finally { engine.dispose(); }
});

test('Word shared style and list batches roll back on failure and reject stale definitions', async () => {
	const engine = await openEngine('word', 'Shared resources.docx', docx, root);
	try {
		const agent = engine.agent!, id = (agent.read({ parentId: 'body' }) as { blocks: WordAgentBlock[] }).blocks[0].id;
		const revision = engine.state().revision;
		const before = unzipSync(await engine.export());
		await assert.rejects(agent.edit({ revision, operation: 'batch', edits: [
			{ operation: 'command', id, command: 'styles.create', params: { name: 'Must not survive' } },
			{ operation: 'command', id, command: 'list.define', params: { levels: [{ text: '%1)' }] } },
			{ operation: 'command', id, command: 'styles.modify', params: { style: 'Missing style', chr: { bold: true } } }
		] }));
		assert.equal(engine.state().revision, revision);
		const after = unzipSync(await engine.export());
		assert.equal(strFromU8(after['word/styles.xml']), strFromU8(before['word/styles.xml']));
		assert.equal(strFromU8(after['word/document.xml']), strFromU8(before['word/document.xml']));
		assert.deepEqual(after['word/numbering.xml'], before['word/numbering.xml']);
		await agent.edit({ revision, operation: 'command', id, command: 'styles.create', params: { name: 'Valid style' } });
		await assert.rejects(agent.edit({ revision, operation: 'command', id, command: 'list.define', params: { levels: [{}] } }), /fresh document version/);
	} finally { engine.dispose(); }
});

test('Word style inspector clears an exact Unicode character range without clearing its neighbors', async () => {
	const engine = await openEngine('word', 'Clear range.docx', docx, root);
	try {
		const agent = engine.agent!, id = (agent.read({ parentId: 'body' }) as { blocks: WordAgentBlock[] }).blocks[0].id;
		await agent.edit({ revision: engine.state().revision, operation: 'update', id, text: '😀 Alpha Beta' });
		await agent.edit({ revision: engine.state().revision, operation: 'command', id, command: 'format.bold', params: { value: true } });
		await agent.edit({ revision: engine.state().revision, operation: 'command', id, start: 2, end: 7, command: 'styles.inspectorClear', params: { level: 'characterFormatting' } });
		const read = (start: number, end: number) => (agent.read({ operation: 'formatting', id, start, end }) as { formatting: { bold: boolean } }).formatting.bold;
		assert.equal(read(2, 7), false); assert.equal(read(8, 12), true);
		await assert.rejects(agent.edit({ revision: engine.state().revision, operation: 'command', id, start: 2, end: 7, command: 'styles.inspectorClear', params: { level: 'paragraphFormatting' } }), /Ranges require/);
		await engine.execute('edit.undo', {}); assert.equal(read(2, 7), true);
	} finally { engine.dispose(); }
});

test('Word paragraph moves refuse to split comment spans and leave all anchors and history intact', async () => {
	const engine = await openEngine('word', 'Move comments.docx', docx, root);
	try {
		const agent = engine.agent!;
		await engine.execute('select.all', {}); await engine.execute('review.newComment', { text: 'Comment across paragraphs' });
		const before = agent.read({ parentId: 'body' }) as { blocks: WordAgentBlock[] }, state = engine.state();
		await assert.rejects(agent.edit({ revision: state.revision, operation: 'command', id: before.blocks[0].id, command: 'para.moveDown', params: {} }), /comment span/);
		assert.deepEqual(agent.read({ parentId: 'body' }), before); assert.deepEqual(engine.state(), state);
		await engine.execute('caret.docStart', {});
		await assert.rejects(engine.execute('para.moveDown', {}), /comment span/);
		assert.equal(engine.state().revision, state.revision);
	} finally { engine.dispose(); }
});

test('Word native paragraph moves can move a complete commented selection as one unit', async () => {
	const engine = await openEngine('word', 'Move complete comment.docx', docx, root);
	try {
		const agent = engine.agent!, read = () => (agent.read({ parentId: 'body' }) as { blocks: WordAgentBlock[] }).blocks;
		await engine.execute('select.all', {}); await engine.execute('review.newComment', { text: 'Move this entire range' });
		await agent.edit({ revision: engine.state().revision, operation: 'append', blocks: [{ kind: 'paragraph', text: 'Outside annotation' }] });
		const before = read();
		await engine.execute('select.range', { anchor: { story: 'body', path: [0], off: 0 }, focus: { story: 'body', path: [2], off: 0 } });
		await engine.execute('para.moveDown', {});
		assert.deepEqual(read().map(block => block.id), [before[2].id, before[0].id, before[1].id]);
		assert.equal(engine.state().comments?.[0].text, 'Move this entire range');
		const reopened = await openEngine('word', 'Reopened.docx', await engine.export(), root);
		try { assert.equal(reopened.state().comments?.[0].text, 'Move this entire range'); } finally { reopened.dispose(); }
		await engine.execute('edit.undo', {}); assert.deepEqual(read(), before);
	} finally { engine.dispose(); }
});
