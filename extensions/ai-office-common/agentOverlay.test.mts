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
		const preview = await page.evaluate(`(() => {
			OfficeActivity.renderAgentActivity(document.getElementById('pages'),[activityRect],activityPlace,'OpenCode','Clarifying <img src=x> and checking figures');
			const preview=document.querySelector('.agent-edit-preview');
			return {text:preview.textContent,images:preview.querySelectorAll('img').length,background:getComputedStyle(preview).backgroundColor,label:document.querySelector('.agent-cursor span').textContent,original:document.querySelector('p').textContent};
		})()`);
		assert.deepEqual(preview, { text: '[Clarifying <img src=x> and checking figures]', images: 0, background: 'rgb(255, 255, 255)', label: 'OpenCode', original: 'Native document content' });
		await page.evaluate(`OfficeActivity.renderAgentActivity(document.getElementById('pages'),[activityRect],activityPlace,'OpenCode','Shortening the paragraph')`);
		assert.equal(await page.locator('.agent-edit-preview').count(), 1, 'Repainting replaces the preview without stacking masks');

		await page.waitForTimeout(450);
		assert.equal(await page.locator('.agent-reveal').count(), 0, 'The fast reveal removes its mask');
		assert.equal(await page.evaluate(`(() => { activityReveal.queue(3,[activityRect]); activityRevision=4;activityReveal.paint();return document.querySelectorAll('.agent-reveal').length; })()`), 0, 'A later edit cancels stale masks');
		await page.emulateMedia({ reducedMotion: 'reduce' });
		assert.equal(await page.evaluate(`(() => { activityRevision=3;activityReveal.queue(3,[activityRect]);return document.querySelectorAll('.agent-reveal').length; })()`), 0);
		assert.equal(await page.locator('.agent-edit-preview p').evaluate(node => getComputedStyle(node).animationName), 'none');
		assert.equal(await page.locator('#pages > section > p').textContent(), 'Native document content');
		await page.evaluate(`OfficeActivity.renderAgentActivity(document.getElementById('pages'),[],activityPlace);activityReveal.clear()`);
		assert.equal(await page.locator('.agent-shimmer,.agent-reveal,.agent-edit-preview').count(), 0);
		assert.equal(await page.locator('#pages > section > p').textContent(), 'Native document content', 'Stopping restores the original display without editing content');
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

test('Review never fabricates a page-one location when geometry is missing', async () => {
	const bundle = await build({ entryPoints: ['extensions/ai-office-common/wordReview.ts'], bundle: true, platform: 'browser', format: 'iife', globalName: 'OfficeReview', write: false });
	const browser = await chromium.launch({ ...(process.env.OFFICE_BROWSER_PATH ? { executablePath: process.env.OFFICE_BROWSER_PATH } : {}), headless: true, args: ['--no-sandbox'] });
	try {
		const page = await browser.newPage();
		await page.setContent('<main id="pages"><section data-page="0" style="position:relative;width:600px;height:800px"><canvas></canvas></section><section data-page="1" style="position:relative;width:600px;height:800px"><canvas></canvas></section></main>');
		await page.addScriptTag({ content: bundle.outputFiles[0].text });
		await page.evaluate(`(() => {window.review=new OfficeReview.WordReview(document.getElementById('pages'),()=>{},()=>{},()=>{});window.change={id:'body:last',label:'Paragraph change',before:'',after:'hello world',beforeText:'',afterText:'hello world',rects:[]};review.update({version:'1',reviewVersion:1,changes:[change]});review.layout([[600,800],[600,800]]);})()`);
		assert.equal(await page.locator('#pages .word-review-zone').count(), 0);
		assert.match(await page.locator('.word-unplaced-review').textContent(), /hello world/);
		await page.evaluate(`review.update({version:'2',reviewVersion:2,changes:[{...change,rects:[{page:1,x:40,y:500,width:100,height:20}]}]});review.layout([[600,800],[600,800]])`);
		assert.equal(await page.locator('[data-page="0"] .word-review-zone').count(), 0);
		assert.equal(await page.locator('[data-page="1"] .word-review-zone').count(), 1);
		assert.equal(await page.locator('.word-unplaced-review').isVisible(), false);
		await page.evaluate(`review.dispose()`);
	} finally { await browser.close(); }
});

test('Activity covers each batch region without masking unrelated content between them', async () => {
	const bundle = await build({ entryPoints: ['extensions/ai-office-common/agentOverlay.ts'], bundle: true, platform: 'browser', format: 'iife', globalName: 'OfficeActivity', write: false });
	const browser = await chromium.launch({ ...(process.env.OFFICE_BROWSER_PATH ? { executablePath: process.env.OFFICE_BROWSER_PATH } : {}), headless: true, args: ['--no-sandbox'] });
	try {
		const page = await browser.newPage();
		await page.setContent('<main id="pages"><section style="position:relative;width:600px;height:800px"><p>Untouched content</p></section></main>');
		await page.addScriptTag({ content: bundle.outputFiles[0].text });
		const result = await page.evaluate(`(() => {
			const pages=document.getElementById('pages');
			OfficeActivity.renderAgentActivity(pages,[{page:0,x:40,y:50,width:200,height:20},{page:0,x:40,y:70,width:180,height:20},{page:0,x:40,y:400,width:200,height:20}],(node,page,x,y,width,height)=>{Object.assign(node.style,{left:x+'px',top:y+'px',width:width+'px',height:height+'px'});pages.children[page].append(node);},'OpenCode','Updating related items');
			return {regions:[...pages.querySelectorAll('.agent-edit-preview')].map(node=>[node.style.top,node.style.height]),cursors:pages.querySelectorAll('.agent-cursor').length,text:pages.querySelector('p').textContent};
		})()`);
		assert.deepEqual(result, { regions: [['50px', '40px'], ['400px', '20px']], cursors: 2, text: 'Untouched content' });
	} finally { await browser.close(); }
});

test('Related block changes share one Keep or Undo action while independent regions stay separate', async () => {
	const bundle = await build({ entryPoints: ['extensions/ai-office-common/wordReview.ts'], bundle: true, platform: 'browser', format: 'iife', globalName: 'OfficeReview', write: false });
	const browser = await chromium.launch({ ...(process.env.OFFICE_BROWSER_PATH ? { executablePath: process.env.OFFICE_BROWSER_PATH } : {}), headless: true, args: ['--no-sandbox'] });
	try {
		const page = await browser.newPage();
		await page.setContent('<main id="pages"><section style="position:relative;width:600px;height:800px"><canvas></canvas></section></main>');
		await page.addScriptTag({ content: bundle.outputFiles[0].text });
		await page.evaluate(`(() => {
			window.calls=[];window.review=new OfficeReview.WordReview(document.getElementById('pages'),(data,action,ids)=>calls.push({action,ids}),()=>{},()=>{});
			window.changes=Array.from({length:7},(_,i)=>({id:'block-'+i,groupId:i<4?'insert:section-one':'insert:section-two',label:'Paragraph change',before:'',after:'Item '+i,beforeText:'',afterText:'Item '+i,rects:[{page:0,x:40,y:50+i*30,width:180,height:20}]}));
			review.update({version:'1',reviewVersion:1,changes});review.layout([[600,800]]);
		})()`);
		assert.equal(await page.locator('.word-review-zone').count(), 2);
		assert.equal(await page.evaluate('review.count'), 2);
		await page.locator('.word-review-zone').first().getByRole('button', { name: 'Keep This Change' }).click();
		assert.deepEqual(await page.evaluate('calls'), [{ action: 'keep', ids: ['block-0', 'block-1', 'block-2', 'block-3'] }]);
		await page.evaluate(`review.update({version:'2',reviewVersion:2,changes:changes.slice(4)});review.layout([[600,800]])`);
		await page.getByRole('button', { name: 'Undo This Change' }).click();
		assert.deepEqual(await page.evaluate('calls[1]'), { action: 'undo', ids: ['block-4', 'block-5', 'block-6'] });
		assert.equal(await page.locator('.word-review-insert').count(), 3);
		await page.evaluate('review.dispose()');
	} finally { await browser.close(); }
});
