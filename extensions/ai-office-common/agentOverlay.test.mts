/*---------------------------------------------------------------------------------------------
 *  Copyright (c) Microsoft Corporation. All rights reserved.
 *  Licensed under the MIT License. See License.txt in the project root for license information.
 *--------------------------------------------------------------------------------------------*/

import { test } from 'node:test';
import assert from 'node:assert/strict';
import { createRequire } from 'node:module';
import { readFile } from 'node:fs/promises';
import { chromium } from 'playwright';
const { build } = createRequire(new URL('../../build/package.json', import.meta.url))('esbuild');

test('Word activity labels, bounded reveals and reduced motion never change document content', async () => {
	const bundle = await build({ entryPoints: ['extensions/ai-office-common/agentOverlay.ts'], bundle: true, platform: 'browser', format: 'iife', globalName: 'OfficeActivity', write: false });
	const browser = await chromium.launch({ ...(process.env.OFFICE_BROWSER_PATH ? { executablePath: process.env.OFFICE_BROWSER_PATH } : {}), headless: true, args: ['--no-sandbox'] });
	try {
		const page = await browser.newPage();
		await page.setContent('<main id="pages"><section style="position:relative;width:600px;height:800px"><canvas data-revision="3"></canvas><p>Native document content</p></section></main>');
		await page.addStyleTag({ content: await readFile('extensions/ai-office-common/editor.css', 'utf8') });
		await page.addScriptTag({ content: bundle.outputFiles[0].text });
		const active = await page.evaluate(`(() => {
			const pages = document.getElementById('pages');
			window.activityPlace = (node, page, x, y, width, height) => { Object.assign(node.style, {left:x+'px',top:y+'px',width:width+'px',height:height+'px'}); pages.children[page].append(node); };
			window.activityRect = {page:0,x:40,y:50,width:180,height:20}; window.activityRevision = 3;
			OfficeActivity.renderAgentActivity(pages,[activityRect,{...activityRect,page:-1}],activityPlace,'OpenCode');
			window.activityReveal = new OfficeActivity.WordAgentReveal(pages,activityPlace,()=>activityRevision);
			activityReveal.queue(3,[activityRect]);
			return {label:document.querySelector('.agent-cursor span').textContent,shimmer:document.querySelectorAll('.agent-shimmer-text').length,masks:document.querySelectorAll('.agent-reveal').length,text:document.querySelector('p').textContent};
		})()`);
		assert.deepEqual(active, { label: 'OpenCode', shimmer: 1, masks: 1, text: 'Native document content' });
		await page.waitForTimeout(450);
		assert.equal(await page.locator('.agent-reveal').count(), 0, 'The fast reveal removes its mask');
		assert.equal(await page.evaluate(`(() => { activityReveal.queue(3,[activityRect]); activityRevision=4;activityReveal.paint();return document.querySelectorAll('.agent-reveal').length; })()`), 0, 'A later edit cancels stale masks');
		await page.emulateMedia({ reducedMotion: 'reduce' });
		assert.equal(await page.evaluate(`(() => { activityRevision=3;activityReveal.queue(3,[activityRect]);return document.querySelectorAll('.agent-reveal').length; })()`), 0);
		assert.equal(await page.locator('p').textContent(), 'Native document content');
		await page.evaluate(`OfficeActivity.renderAgentActivity(document.getElementById('pages'),[],activityPlace);activityReveal.clear()`);
		assert.equal(await page.locator('.agent-shimmer,.agent-reveal').count(), 0);
	} finally { await browser.close(); }
});

test('Only the current draft shimmers and an expired or stopped lease clears its cursor', async () => {
	const bundle = await build({ entryPoints: ['extensions/ai-office-common/wordReview.ts'], bundle: true, platform: 'browser', format: 'iife', globalName: 'OfficeReview', write: false });
	const browser = await chromium.launch({ ...(process.env.OFFICE_BROWSER_PATH ? { executablePath: process.env.OFFICE_BROWSER_PATH } : {}), headless: true, args: ['--no-sandbox'] });
	try {
		const page = await browser.newPage();
		await page.setContent('<main id="pages"><section style="position:relative;width:600px;height:800px"><canvas></canvas></section></main>');
		await page.addStyleTag({ content: await readFile('extensions/ai-office-common/editor.css', 'utf8') });
		await page.addScriptTag({ content: bundle.outputFiles[0].text });
		await page.evaluate(`(() => {
			window.draftReview = new OfficeReview.WordReview(document.getElementById('pages'),()=>{},()=>{},()=>{});
			window.draftSections = ['one','two'].map(id=>({id,title:id,intent:'Preparing '+id,anchor:{id:'anchor',position:'after'},anchorRects:[{page:0,x:40,y:50,width:180,height:20}]}));
			draftReview.updateDrafts(draftSections,'OpenCode','one',Date.now()+350);draftReview.layout([[600,800]]);
		})()`);
		assert.equal(await page.locator('.word-draft-zone').count(), 2);
		assert.equal(await page.locator('.word-draft-zone[data-working="true"]').count(), 1);
		assert.equal(await page.locator('.word-draft-zone[data-working="true"] small').textContent(), 'OpenCode');
		assert.equal(await page.locator('.word-draft-zone[data-working="false"] strong').evaluate(node => getComputedStyle(node).animationName), 'none');
		await page.waitForTimeout(400);
		assert.equal(await page.locator('.word-draft-zone[data-working="true"]').count(), 0);
		await page.evaluate(`draftReview.updateDrafts(draftSections,'OpenCode','two',0);draftReview.layout([[600,800]])`);
		assert.equal(await page.locator('.word-draft-zone[data-working="true"]').count(), 0);
		await page.evaluate(`draftReview.dispose()`);
	} finally { await browser.close(); }
});
