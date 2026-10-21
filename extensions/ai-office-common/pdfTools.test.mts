/*---------------------------------------------------------------------------------------------
 *  Copyright (c) Microsoft Corporation. All rights reserved.
 *  Licensed under the MIT License. See License.txt in the project root for license information.
 *--------------------------------------------------------------------------------------------*/

import { test } from 'node:test';
import assert from 'node:assert/strict';
import { chromium } from 'playwright';
import { createRequire } from 'node:module';
import { readFile } from 'node:fs/promises';
const { build } = createRequire(new URL('../../build/package.json', import.meta.url))('esbuild');

test('PDF tools share exact native parameters, reject stale dialogs and retain keyboard dismissal', async () => {
	const bundle = await build({ entryPoints: ['extensions/ai-office-common/pdfTools.ts'], bundle: true, platform: 'browser', format: 'iife', globalName: 'OfficePDF', write: false });
	const browser = await chromium.launch({ executablePath: process.env.OFFICE_BROWSER_PATH, headless: true, args: ['--no-sandbox'] });
	try {
		const page = await browser.newPage();
		await page.setContent('<button id="open">PDF Tools</button>');
		await page.addStyleTag({ content: await readFile('extensions/ai-office-common/editor.css', 'utf8') });
		await page.addScriptTag({ content: bundle.outputFiles[0].text });
		await page.evaluate(`window.snapshot={total:6,page:4,revision:7};window.calls=[];window.tools=new OfficePDF.PdfTools(()=>snapshot,(command,params,revision)=>calls.push({command,params,revision}));document.getElementById('open').onclick=()=>tools.open()`);
		await page.locator('#open').click();
		await page.getByLabel('Pages', { exact: true }).fill('2, 4–6');
		await page.getByRole('button', { name: 'Apply', exact: true }).click();
		assert.deepEqual(await page.evaluate('calls'), [{ command: 'page_rotate', params: { pages: [2, 4, 5, 6], degrees: 90 }, revision: 7 }]);
		await page.locator('#open').click(); await page.getByLabel('PDF operation').selectOption('page_delete'); await page.getByLabel('Pages', { exact: true }).fill('all');
		await page.getByRole('button', { name: 'Delete Selected Pages' }).click(); assert.match(await page.getByRole('alert').textContent(), /Keep at least one/);
		await page.getByLabel('PDF operation').selectOption('doc_watermark');
		await page.evaluate('snapshot.revision++'); await page.getByRole('button', { name: 'Apply', exact: true }).click();
		assert.match(await page.getByRole('alert').textContent(), /document changed/);
		assert.equal(await page.evaluate('calls.length'), 1);
		await page.keyboard.press('Escape'); assert.equal(await page.getByRole('dialog').isVisible(), false);
		await page.locator('#open').click(); await page.getByLabel('PDF operation').selectOption('doc_header_footer'); await page.getByRole('button', { name: 'Apply', exact: true }).click();
		assert.deepEqual(await page.evaluate('calls.at(-1)'), { command: 'doc_header_footer', params: { pages: [5], footer_center: '<<Page 1 of n>>', font_size: 10 }, revision: 8 });
		await page.evaluate('tools.dispose()');
	} finally { await browser.close(); }
});

test('All PDF Tools renders native schemas, optional typed properties and stale guards without JSON editing', async () => {
	const bundle = await build({ entryPoints: ['extensions/ai-office-common/pdfCommandForm.ts', 'extensions/ai-office-common/pdfInspection.ts'], bundle: true, platform: 'browser', format: 'iife', globalName: 'OfficePDFForm', write: false, outdir: '.build/pdf-browser-tests' });
	const browser = await chromium.launch({ executablePath: process.env.OFFICE_BROWSER_PATH, headless: true, args: ['--no-sandbox'] });
	try {
		const page = await browser.newPage(); await page.setContent('<button id="open">All PDF Tools</button>');
		await page.addStyleTag({ content: await readFile('extensions/ai-office-common/editor.css', 'utf8') });
		await page.addScriptTag({ content: bundle.outputFiles.find(file => file.path.endsWith('pdfCommandForm.js'))!.text });
		await page.evaluate(`window.snapshot={total:3,revision:4,commands:[{id:'form_set_props',label:'Field Properties',params:{type:'object',properties:{field:{type:'string'},appearance:{type:'object'}},required:['field'],additionalProperties:false}}]};window.calls=[];window.tools=new OfficePDFForm.PdfCommandForm(()=>snapshot,(command,params,revision)=>calls.push({command,params,revision}));document.getElementById('open').onclick=()=>tools.open()`);
		await page.locator('#open').click(); await page.getByLabel('Field', { exact: false }).first().fill('Answer');
		await page.getByRole('checkbox', { name: 'Set Appearance' }).check(); await page.getByRole('button', { name: 'Add Entry' }).click(); await page.getByPlaceholder('Name', { exact: true }).fill('width');
		await page.locator('.pdf-schema-entry select').selectOption('1'); await page.getByRole('spinbutton').fill('2'); await page.getByRole('button', { name: 'Apply', exact: true }).click();
		assert.deepEqual(await page.evaluate('calls'), [{ command: 'form_set_props', params: { field: 'Answer', appearance: { width: 2 } }, revision: 4 }]);
		await page.locator('#open').click(); await page.getByLabel('Field', { exact: false }).first().fill('Answer'); await page.evaluate('snapshot={...snapshot,revision:5}'); await page.getByRole('button', { name: 'Apply', exact: true }).click();
		assert.match(await page.getByRole('alert').textContent(), /document changed/); assert.equal(await page.evaluate('calls.length'), 1); await page.keyboard.press('Escape');
		await page.addScriptTag({ content: bundle.outputFiles.find(file => file.path.endsWith('pdfInspection.js'))!.text });
		await page.evaluate(`window.inspection=document.createElement('div');document.body.append(inspection);OfficePDFForm.renderPdfInspection(inspection,[{doc:3,text:'<img src=x onerror=alert(1)>',items:Array.from({length:10000},(_,i)=>({value:i}))}])`);
		assert.equal(await page.locator('img').count(), 0); assert.ok(await page.locator('details').count() < 200); assert.match(await page.locator('body').textContent(), /Showing 40 of 10000/);
		await page.evaluate('tools.dispose()');
	} finally { await browser.close(); }
});

test('PDF contextual tools prefill page, target and drawn geometry in the native form', async () => {
	const bundle = await build({ entryPoints: ['extensions/ai-office-common/pdfCommandForm.ts'], bundle: true, platform: 'browser', format: 'iife', globalName: 'OfficePDFForm', write: false });
	const browser = await chromium.launch({ executablePath: process.env.OFFICE_BROWSER_PATH, headless: true, args: ['--no-sandbox'] });
	try {
		const page = await browser.newPage(); page.setDefaultTimeout(8000); await page.setContent('<div></div>'); await page.addScriptTag({ content: bundle.outputFiles[0].text });
		await page.evaluate(`window.calls=[];window.tools=new OfficePDFForm.PdfCommandForm(()=>({total:4,revision:6,commands:[{id:'form_add_field',label:'Form field',params:{type:'object',properties:{page:{type:'integer'},type:{type:'string',enum:['text','checkbox']},rect:{type:'array',items:{type:'number'},minItems:4,maxItems:4},name:{type:'string'}},required:['page','type','rect']}}]}),(command,args,revision)=>calls.push({command,args,revision}));tools.open('form_add_field',{page:3,type:'checkbox',rect:[10,20,40,50]})`);
		await page.getByRole('button', { name: 'Apply', exact: true }).click();
		assert.deepEqual(await page.evaluate('calls'), [{ command: 'form_add_field', args: { page: 3, type: 'checkbox', rect: [10, 20, 40, 50] }, revision: 6 }]);
	} finally { await browser.close(); }
});
