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
