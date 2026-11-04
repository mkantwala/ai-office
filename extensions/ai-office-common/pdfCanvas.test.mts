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

async function setup() {
	const bundle = await build({ entryPoints: ['extensions/ai-office-common/pdfCanvas.ts'], bundle: true, platform: 'browser', format: 'iife', globalName: 'PDFCanvas', write: false });
	const browser = await chromium.launch({ executablePath: process.env.OFFICE_BROWSER_PATH, headless: true, args: ['--no-sandbox'] });
	const page = await browser.newPage({ viewport: { width: 1200, height: 1000 } });
	page.setDefaultTimeout(8000);
	await page.setContent('<div id="pages"><section class="page pdf" data-page="0" tabindex="-1" style="position:relative;width:900px;height:1200px"></section></div>');
	await page.addStyleTag({ content: await readFile('extensions/ai-office-common/editor.css', 'utf8') });
	await page.addScriptTag({ content: bundle.outputFiles[0].text });
	await page.evaluate(`window.calls=[];window.forms=[];window.messages=[];window.snapshot={dimensions:[[600,800]],state:{revision:3,page:0,pdfRotation:0,blocks:[{id:1,text:'Original paragraph',size:12,rect:[60,80,260,120]}],pdfObjects:[{kind:'text',index:1,type:'text',text:'Original paragraph',rect:[60,80,260,120]},{kind:'added',index:1,type:'image',rect:[300,200,400,300]}]}};window.canvas=new PDFCanvas.PdfCanvas(document.getElementById('pages'),{snapshot:()=>snapshot,execute:(command,args,revision)=>calls.push({command,args,revision}),form:(command,args)=>forms.push({command,args}),select:()=>{},page:()=>{},message:text=>messages.push(text),organize:()=>{}});canvas.update()`);
	return { browser, page };
}

test('PDF canvas drags at display scale, previews locally, resizes text and cancels stale or escaped gestures', async () => {
	const { browser, page } = await setup();
	try {
		const box = await page.locator('[data-object="text:1"]').boundingBox(); assert.ok(box);
		await page.mouse.move(box.x + 30, box.y + 20); await page.mouse.down(); await page.mouse.move(box.x + 60, box.y + 65, { steps: 5 });
		assert.equal(await page.locator('.pdf-drag-preview').count(), 1); assert.equal(await page.evaluate('calls.length'), 0);
		await page.mouse.up();
		assert.deepEqual(await page.evaluate('calls[0]'), { command: 'object_move', args: { page: 1, objects: [{ kind: 'text', index: 1 }], offset: [20, 30], generation: 3 }, revision: 3 });
		await page.evaluate('canvas.update()'); await page.locator('[data-object="text:1"]').click();
		const grip = await page.locator('.pdf-grip-right').boundingBox(); assert.ok(grip);
		await page.mouse.move(grip.x + grip.width / 2, grip.y + grip.height / 2); await page.mouse.down(); await page.mouse.move(grip.x + grip.width / 2 + 60, grip.y + grip.height / 2); await page.mouse.up();
		assert.deepEqual(await page.evaluate('calls.at(-1)'), { command: 'text_edit', args: { page: 1, paragraph: 1, width: 240, dx: 0 }, revision: 3 });
		await page.evaluate('canvas.update()'); const count = await page.evaluate('calls.length');
		await page.mouse.move(box.x + 30, box.y + 20); await page.mouse.down(); await page.mouse.move(box.x + 70, box.y + 20); await page.keyboard.press('Escape'); await page.mouse.up(); assert.equal(await page.evaluate('calls.length'), count);
		await page.mouse.move(box.x + 30, box.y + 20); await page.mouse.down(); await page.mouse.move(box.x + 70, box.y + 20); await page.evaluate('snapshot.state.revision++'); await page.mouse.up(); assert.equal(await page.evaluate('calls.length'), count); assert.match(await page.evaluate('messages.at(-1)'), /changed/);
	} finally { await browser.close(); }
});

test('PDF canvas groups moves, edits on double click and creates page-positioned tools', async () => {
	const { browser, page } = await setup();
	try {
		await page.locator('[data-object="text:1"]').click(); await page.locator('[data-object="added:1"]').click({ modifiers: ['Shift'] });
		const box = await page.locator('[data-object="added:1"]').boundingBox(); assert.ok(box);
		await page.mouse.move(box.x + 30, box.y + 30); await page.mouse.down(); await page.mouse.move(box.x + 45, box.y + 45); await page.mouse.up();
		assert.equal(await page.evaluate('calls[0].args.objects.length'), 2);
		await page.evaluate('canvas.update()'); await page.locator('[data-object="text:1"]').dblclick();
		await page.getByLabel('Edit PDF text').fill('Changed text'); assert.equal(await page.evaluate('calls.length'), 1);
		await page.getByLabel('Edit PDF text').press('Control+Enter'); assert.equal(await page.evaluate('calls.at(-1).args.text'), 'Changed text');
		await page.evaluate('canvas.update();canvas.setMode("form")'); const surface = await page.locator('.page').boundingBox(); assert.ok(surface);
		await page.mouse.move(surface.x + 150, surface.y + 500); await page.mouse.down(); await page.mouse.move(surface.x + 450, surface.y + 560); await page.mouse.up();
		assert.deepEqual(await page.evaluate('forms.at(-1)'), { command: 'form_add_field', args: { page: 1, type: 'text', rect: [100, 500 / 1.5, 300, 560 / 1.5] } });
		await page.evaluate('canvas.setMode("redact")'); await page.mouse.move(surface.x + 150, surface.y + 500); await page.mouse.down(); await page.mouse.move(surface.x + 450, surface.y + 560); await page.mouse.up();
		assert.equal(await page.evaluate('calls.at(-1).command'), 'redact_mark', 'Drawing must never apply permanent redaction');
	} finally { await browser.close(); }
});

test('PDF text drafts combine formatting atomically and survive concurrent agent changes', async () => {
	const { browser, page } = await setup();
	try {
		await page.locator('[data-object="text:1"]').dblclick(); await page.getByLabel('Edit PDF text').fill('Draft with formatting');
		await page.evaluate('canvas.format({bold:true})');
		assert.deepEqual(await page.evaluate('calls'), [{ command: 'text_edit', args: { page: 1, paragraph: 1, text: 'Draft with formatting', bold: true }, revision: 3 }]);
		await page.evaluate('snapshot.state.revision++;canvas.update()'); await page.locator('[data-object="text:1"]').dblclick(); await page.getByLabel('Edit PDF text').fill('Concurrent draft');
		await page.evaluate('snapshot.state.revision++;canvas.update();canvas.flush()');
		assert.equal(await page.getByLabel('Edit PDF text').inputValue(), 'Concurrent draft'); assert.equal(await page.evaluate('calls.length'), 1);
		assert.match(await page.evaluate('messages.at(-1)'), /Copy your draft/); await page.getByLabel('Edit PDF text').press('Escape'); assert.equal(await page.getByLabel('Edit PDF text').count(), 0);
	} finally { await browser.close(); }
});

test('PDF image resize uses full placement and preserves the visible crop geometry', async () => {
	const { browser, page } = await setup();
	try {
		await page.evaluate(`snapshot.state.pdfObjects=[{kind:'image',index:1,type:'image',rect:[100,100,200,200],fullRect:[80,80,220,220]}];snapshot.state.revision++;canvas.update()`);
		await page.locator('[data-object="image:1"]').click(); const handle = await page.locator('.pdf-grip-corner').boundingBox(); assert.ok(handle);
		await page.mouse.move(handle.x + 4, handle.y + 4); await page.mouse.down(); await page.mouse.move(handle.x + 79, handle.y + 79); await page.mouse.up();
		assert.deepEqual(await page.evaluate('calls.at(-1)'), { command: 'image_edit', args: { page: 1, image: 1, action: 'move', rect: [70, 70, 280, 280] }, revision: 4 });
	} finally { await browser.close(); }
});

test('PDF selection actions stay on the paper near its top and right edges', async () => {
	const { browser, page } = await setup();
	try {
		await page.evaluate(`snapshot.state.pdfObjects=[{kind:'text',index:1,type:'text',text:'Near edge',rect:[550,2,598,22]}];snapshot.state.revision++;canvas.update()`);
		await page.locator('[data-object="text:1"]').click(); const actions = await page.getByRole('toolbar', { name: 'Selected PDF object' }).boundingBox(), surface = await page.locator('.page').boundingBox(); assert.ok(actions && surface);
		assert.ok(actions.x >= surface.x && actions.x + actions.width <= surface.x + surface.width); assert.ok(actions.y >= surface.y);
		await page.getByRole('button', { name: 'Properties', exact: true }).click(); assert.equal(await page.evaluate('forms.at(-1).args.paragraph'), 1);
	} finally { await browser.close(); }
});
