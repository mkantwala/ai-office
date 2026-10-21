/*---------------------------------------------------------------------------------------------
 *  Copyright (c) Microsoft Corporation. All rights reserved.
 *  Licensed under the MIT License. See License.txt in the project root for license information.
 *--------------------------------------------------------------------------------------------*/

import { test } from 'node:test';
import assert from 'node:assert/strict';
import { readFile, writeFile, mkdtemp, rm, mkdir, symlink } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import { resolve, join, win32, posix } from 'node:path';
import { fileURLToPath, pathToFileURL } from 'node:url';
import { randomUUID } from 'node:crypto';
import { openEngine } from './model.js';
import { pdfResourcePath } from './nativePdf.js';
import { unzipSync, strFromU8 } from '../ai-office-slides/node_modules/fflate/esm/index.mjs';
import type { Engine } from './types.js';

const root = resolve('extensions');
const fixtureTooltip = 'Response';
function fixture() {
	const stream = 'BT /F1 20 Tf 72 700 Td (Synthetic PDF content) Tj ET\n';
	const objects = ['<< /Type /Catalog /Pages 2 0 R >>', '<< /Type /Pages /Kids [3 0 R] /Count 1 >>', '<< /Type /Page /Parent 2 0 R /MediaBox [0 0 612 792] /Resources << /Font << /F1 5 0 R >> >> /Contents 4 0 R >>', `<< /Length ${stream.length} >>\nstream\n${stream}endstream`, '<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica >>'];
	let value = '%PDF-1.7\n'; const offsets = [0]; objects.forEach((object, index) => { offsets.push(Buffer.byteLength(value)); value += `${index + 1} 0 obj\n${object}\nendobj\n`; }); const xref = Buffer.byteLength(value);
	value += `xref\n0 ${offsets.length}\n0000000000 65535 f \n` + offsets.slice(1).map(offset => String(offset).padStart(10, '0') + ' 00000 n \n').join('') + `trailer\n<< /Size ${offsets.length} /Root 1 0 R >>\nstartxref\n${xref}\n%%EOF\n`; return Buffer.from(value);
}
async function setup() {
	const work = await mkdtemp(join(tmpdir(), 'sangam-pdf-test-')), source = join(work, 'Original.pdf'); await writeFile(source, fixture());
	const engine = await openEngine('pdf', 'Original.pdf', fixture(), root, undefined, { artifactRoot: join(work, 'artifacts'), sourceUri: pathToFileURL(source).href, resourceRoot: work });
	return { work, source, engine, close: async () => { engine.dispose(); await rm(work, { recursive: true, force: true }); } };
}
const result = (engine: Engine) => engine.state().agentCommandResult as Record<string, unknown>[];

test('PDF native catalog accounts for every reference command and excludes OCR including batch/action bypasses', async () => {
	const { engine, close } = await setup();
	try {
		const audit = JSON.parse(await readFile(resolve(root, 'ai-office-pdf/PDF_CAPABILITIES.json'), 'utf8')) as { commands: { command: string; integration: string; reason?: string }[] };
		for (const entry of audit.commands) { assert.equal(engine.commands().some(command => command.id === entry.command), entry.integration === 'native', entry.command); if (entry.integration !== 'native') { assert.ok(entry.reason); } }
		assert.equal(audit.commands.length, 171); assert.equal(engine.commands().filter(command => audit.commands.some(entry => entry.command === command.id && entry.integration === 'native')).length, 159);
		await assert.rejects(engine.execute('ocr_recognize', {}), /excluded/);
		await assert.rejects(engine.execute('doc_batch', { paths: [], steps: ['ocr'] }), /OCR is excluded/);
		await assert.rejects(engine.execute('action_run', { paths: [], steps: [{ step: 'recognize_text' }] }), /OCR is excluded/);
		await engine.execute('action_list', {}); assert.ok(!JSON.stringify(result(engine)).includes('recognize_text'));
	} finally { await close(); }
});

test('PDF native forms, scripts, bookmarks, transitions and measurements persist through export', async () => {
	const { engine, close } = await setup();
	try {
		await engine.execute('bookmark_add', { title: 'Start', page: 1 }); await engine.execute('bookmark_rename', { path: [1], title: 'Introduction' });
		await engine.execute('page_transition_set', { pages: [1], style: 'Fade', duration: 1 });
		await engine.execute('form_add_field', { name: 'Answer', page: 1, type: 'text', rect: [72, 180, 300, 210] });
		await engine.saved!(); await engine.execute('doc_info', {}); assert.equal((result(engine)[0].document as { dirty: boolean }).dirty, false);
		await engine.execute('form_fill', { values: { Answer: '42' } }); await engine.execute('form_set_props', { field: 'Answer', tooltip: fixtureTooltip, required: true });
		await engine.execute('js_enabled', { enabled: true }); await engine.execute('js_run', { script: `this.getField('Answer').value='43'` });
		assert.equal(result(engine)[0].error, null);
		await engine.execute('measure_distance', { page: 1, points: [[72, 250], [172, 250]] }); await engine.execute('measure_export', {});
		const csv = String(result(engine)[0].path); assert.match(await readFile(fileURLToPath(csv), 'utf8'), /Distance|distance/);
		const reopened = await openEngine('pdf', 'Saved.pdf', await engine.export(), root);
		try {
			await reopened.execute('form_fields', {}); assert.match(JSON.stringify(result(reopened)), /43/); assert.match(JSON.stringify(result(reopened)), /Response/);
			await reopened.execute('bookmark_list', {}); assert.match(JSON.stringify(result(reopened)), /Introduction/);
			await reopened.execute('doc_info', {}); assert.match(JSON.stringify(result(reopened)), /Fade/);
			assert.ok((await reopened.renderAsync!(0, .5)).pixels!.length);
		} finally { reopened.dispose(); }
	} finally { await close(); }
});

test('PDF native merge, split, extraction, comparison, exports and batch artifacts preserve source files', async () => {
	const { engine, source, work, close } = await setup(); const uri = pathToFileURL(source).href;
	try {
		await engine.execute('page_insert_blank', { at: 2 });
		await engine.execute('doc_split', { every: 1 }); const parts = result(engine)[0].files as { path: string }[]; assert.equal(parts.length, 2);
		for (const part of parts) { const reopened = await openEngine('pdf', 'part.pdf', await readFile(fileURLToPath(part.path)), root); assert.equal(reopened.pages().length, 1); reopened.dispose(); }
		await engine.execute('page_extract', { pages: [1] }); assert.ok(result(engine).some(row => typeof row.created_document === 'string'));
		await engine.execute('doc_combine', { paths: [uri, uri] }); const combined = result(engine).find(row => row.created_document)!;
		const combinedEngine = await openEngine('pdf', 'Combined.pdf', await readFile(fileURLToPath(String(combined.created_document))), root); assert.equal(combinedEngine.pages().length, 4, 'The active source URI uses current unsaved bytes'); combinedEngine.dispose();
		await engine.execute('doc_compare', { other: uri }); assert.ok(Array.isArray(result(engine)[0].changes));
		await engine.execute('doc_compare_report', { other: uri }); assert.match(String(result(engine)[0].path), /^file:/);
		for (const format of ['docx', 'pptx', 'html', 'rtf']) {
			await engine.execute('doc_export_office', { format }); const bytes = await readFile(fileURLToPath(String(result(engine)[0].path)));
			if (format === 'docx' || format === 'pptx') { const files = unzipSync(bytes); assert.ok(Object.keys(files).some(name => name.startsWith(format === 'docx' ? 'word/' : 'ppt/'))); if (format === 'docx') { assert.match(strFromU8(files['word/document.xml']), /Synthetic PDF content/); } }
			else { assert.match(bytes.toString(), /Synthetic PDF content/); }
		}
		await engine.execute('doc_export_images', { pages: [1], dpi: 72 }); const png = (result(engine)[0].files as string[])[0]; assert.equal((await readFile(fileURLToPath(png))).readUInt32BE(0), 0x89504e47);
		await engine.execute('accessibility_report', {}); assert.match(await readFile(fileURLToPath(String(result(engine)[0].path)), 'utf8'), /html/i);
		await engine.execute('doc_optimize', {}); const optimized = await openEngine('pdf', 'optimized.pdf', await readFile(fileURLToPath(String(result(engine)[0].path))), root); assert.equal(optimized.pages().length, 2); optimized.dispose();
		await engine.execute('doc_batch', { paths: [uri], steps: ['optimize'] }); assert.equal(result(engine)[0].failed, 0);
		await engine.execute('action_run', { paths: [uri], steps: [{ step: 'add_watermark', arg: 'SYNTHETIC' }] }); assert.ok((result(engine)[0].files as { output: string }[])[0].output);
		assert.deepEqual(await readFile(source), fixture(), 'Every export leaves the source on disk intact');
		assert.ok((engine.state().pdfArtifacts || []).every(artifact => artifact.uri.startsWith(pathToFileURL(join(work, 'artifacts')).href)));
	} finally { await close(); }
});

test('PDF native redaction removes recoverable text; security and signatures create valid independent artifacts', async () => {
	const { engine, close } = await setup();
	try {
		await engine.execute('redact_mark', { find: 'Synthetic PDF' }); await engine.execute('redact_apply', {});
		const redacted = await openEngine('pdf', 'Redacted.pdf', await engine.export(), root);
		try { await redacted.execute('text_find', { query: 'Synthetic PDF' }); assert.equal(result(redacted)[0].count, 0); } finally { redacted.dispose(); }
		await engine.execute('edit_undo', {}); await engine.execute('edit_undo', {});
		const password = 'synthetic-' + randomUUID();
		await engine.execute('doc_protect', { open_password: password, permissions_password: password + '-owner', copy: true }); const encrypted = await engine.export();
		await assert.rejects(openEngine('pdf', 'Protected.pdf', encrypted, root), /password/i);
		const unlocked = await openEngine('pdf', 'Protected.pdf', encrypted, root, password + '-owner');
		try { await unlocked.execute('doc_unprotect', {}); const plain = await openEngine('pdf', 'Plain.pdf', await unlocked.export(), root); assert.equal(plain.pages().length, 1); plain.dispose(); } finally { unlocked.dispose(); }
		await engine.execute('edit_undo', {}); const before = await engine.export(), revision = engine.state().revision;
		await engine.execute('sign_id_create', { name: 'Synthetic Test', password, key: 'p256' }); const id = String(result(engine)[0].path);
		await engine.execute('sign_document', { id, password, page: 1, rect: [72, 500, 300, 560] }); const signed = String(result(engine)[0].path);
		assert.equal(engine.state().revision, revision); assert.deepEqual(await engine.export(), before, 'Signing a copy preserves the current document and native history');
		assert.ok(!JSON.stringify(result(engine)).includes(password));
		const reopened = await openEngine('pdf', 'Signed.pdf', await readFile(fileURLToPath(signed)), root);
		try { await reopened.execute('sign_list', {}); assert.match(JSON.stringify(result(reopened)), /Synthetic Test/); assert.match(JSON.stringify(result(reopened)), /not been modified|intact|valid/i); } finally { reopened.dispose(); }
	} finally { await close(); }
});

test('PDF native URI workflows enforce root confinement, preserve existing outputs and keep exports beyond engine disposal', async () => {
	const { engine, source, work, close } = await setup(); const outside = await mkdtemp(join(tmpdir(), 'sangam-pdf-outside-'));
	try {
		await assert.rejects(engine.execute('doc_export_text', { path: pathToFileURL(source).href }), /new file|already exists/);
		await assert.rejects(engine.execute('page_insert_file', { at: 1, path: pathToFileURL(join(outside, 'not-a-file.pdf')).href }), /outside --root/);
		await assert.rejects(engine.execute('doc_export_text', { path: pathToFileURL(join(outside, 'not-written.txt')).href }), /outside --root/);
		if (process.platform !== 'win32') { const linked = join(work, 'linked'); await symlink(outside, linked); await assert.rejects(engine.execute('page_insert_file', { at: 1, path: pathToFileURL(join(linked, 'no.pdf')).href }), /outside --root/); }
		const output = join(work, 'result.txt'); await engine.execute('doc_export_text', { path: pathToFileURL(output).href }); assert.match(await readFile(output, 'utf8'), /Synthetic PDF content/);
		await assert.rejects(engine.execute('doc_export_text', { path: pathToFileURL(output).href }), /already exists/);
		await engine.execute('doc_export_office', {}); const artifact = String(result(engine)[0].path); engine.dispose(); assert.ok((await readFile(fileURLToPath(artifact))).length);
		assert.deepEqual(await readFile(source), fixture());
	} finally { await close(); await rm(outside, { recursive: true, force: true }); }
});

test('PDF native image editing, camera cleanup and PDF/A report actual conversion results', async () => {
	const { engine, work, close } = await setup();
	try {
		await engine.execute('doc_export_images', { pages: [1], dpi: 72 }); const image = (result(engine)[0].files as string[])[0];
		await engine.execute('doc_create', { from: 'images', paths: [image] }); const created = String(result(engine).find(row => row.created_document)!.created_document);
		const picture = await openEngine('pdf', 'Image.pdf', await readFile(fileURLToPath(created)), root, undefined, { artifactRoot: join(work, 'artifacts'), resourceRoot: work });
		try {
			await picture.execute('page_images', { page: 1 }); assert.ok(JSON.stringify(result(picture)).includes('image'));
			await picture.execute('image_edit', { page: 1, image: 1, action: 'flip_horizontal' });
			await picture.execute('image_save', { page: 1, image: 1 }); const exported = await readFile(fileURLToPath(String(result(picture)[0].path))); assert.ok(exported.length);
			await picture.execute('page_enhance_camera', { page: 1, mode: 'greyscale', level: 30, corners: [0, 0, picture.pages()[0][0], 0, picture.pages()[0][0], picture.pages()[0][1], 0, picture.pages()[0][1]] });
			await picture.execute('pdfa_convert', { level: '2b' }); assert.ok(result(picture)[0]);
			await picture.execute('pdfa_verify', { level: '2b' }); assert.ok(result(picture)[0]);
			const reopened = await openEngine('pdf', 'Saved-image.pdf', await picture.export(), root); assert.ok((await reopened.renderAsync!(0, .5)).pixels!.length); reopened.dispose();
		} finally { picture.dispose(); }
	} finally { await close(); }
});

test('PDF renders and exports when the system temporary folder is a symlink', async () => {
	const work = await mkdtemp(join(tmpdir(), 'sangam-pdf-temp-alias-'));
	const directory = join(work, 'physical'), alias = join(work, 'alias');
	const keys = process.platform === 'win32' ? ['TMP', 'TEMP'] : ['TMPDIR'];
	const original = keys.map(key => process.env[key]);
	let engine: Engine | undefined;
	try {
		await mkdir(directory); await symlink(directory, alias, process.platform === 'win32' ? 'junction' : 'dir');
		for (const key of keys) { process.env[key] = alias; }
		engine = await openEngine('pdf', 'Alias.pdf', fixture(), root);
		const raster = await engine.renderAsync!(0, .5);
		assert.ok(raster.pixels && raster.width && raster.height);
		assert.equal(raster.pixels.length, raster.width * raster.height * 4);
		assert.equal(Buffer.from(await engine.export()).subarray(0, 4).toString(), '%PDF');
		await engine.execute('page_rotate', { pages: [1], degrees: 90 });
		assert.ok((await engine.renderAsync!(0, .5)).pixels?.length);
		assert.equal(Buffer.from(await engine.export()).subarray(0, 4).toString(), '%PDF');
	} finally {
		keys.forEach((key, index) => { if (original[index] === undefined) { delete process.env[key]; } else { process.env[key] = original[index]; } });
		engine?.dispose(); await rm(work, { recursive: true, force: true });
	}
});

test('PDF private resources accept equivalent Windows namespace paths and reject different files', () => {
	const directory = String.raw`C:\Users\Reviewer\AppData\Local\Temp\sangam-pdf-test`;
	assert.equal(pdfResourcePath(directory, String.raw`\\?\C:\Users\Reviewer\AppData\Local\Temp\sangam-pdf-test\raster.raw`, 'raster.raw', win32), win32.join(directory, 'raster.raw'));
	assert.equal(pdfResourcePath(win32.toNamespacedPath(directory), win32.join(directory, 'raster.raw'), 'raster.raw', win32), win32.toNamespacedPath(win32.join(directory, 'raster.raw')));
	assert.equal(pdfResourcePath(directory, String.raw`\\?\c:\users\reviewer\appdata\local\temp\sangam-pdf-test\snapshot.pdf`, 'snapshot.pdf', win32), win32.join(directory, 'snapshot.pdf'));
	const unc = String.raw`\\server\share\Temp\sangam-pdf-test`;
	assert.equal(pdfResourcePath(unc, String.raw`\\?\UNC\server\share\Temp\sangam-pdf-test\raster.raw`, 'raster.raw', win32), win32.join(unc, 'raster.raw'));
	assert.equal(pdfResourcePath('/private/tmp/sangam-pdf-test', '/private/tmp/sangam-pdf-test/raster.raw', 'raster.raw', posix), '/private/tmp/sangam-pdf-test/raster.raw');
	for (const reported of [String.raw`C:\Other\raster.raw`, win32.join(directory, 'input.pdf'), win32.join(directory, 'nested', 'raster.raw'), undefined, null, '', 'raster.raw\0other']) {
		assert.throws(() => pdfResourcePath(directory, reported, 'raster.raw', win32), /Invalid PDF raster resource/);
	}
	assert.throws(() => pdfResourcePath('/tmp/session', '/tmp/other/snapshot.pdf', 'snapshot.pdf', posix), /Invalid PDF export resource/);
});
