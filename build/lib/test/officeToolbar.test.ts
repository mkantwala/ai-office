/*---------------------------------------------------------------------------------------------
 *  Copyright (c) Microsoft Corporation. All rights reserved.
 *  Licensed under the MIT License. See License.txt in the project root for license information.
 *--------------------------------------------------------------------------------------------*/

import assert from 'node:assert/strict';
import { once } from 'node:events';
import { readFile } from 'node:fs/promises';
import { createServer } from 'node:http';
import path from 'node:path';
import { after, before, suite, test } from 'node:test';
import { build } from 'esbuild';
import { chromium, type Browser } from 'playwright';

const root = path.resolve(import.meta.dirname, '../../..');

suite('Office toolbar in a browser', () => {
	let browser: Browser;
	let address: string;
	let script: string;
	let appScript: string;
	let css: string;
	const server = createServer((request, response) => {
		response.setHeader('Content-Type', request.url === '/editor.js' ? 'text/javascript' : request.url === '/editor.css' ? 'text/css' : 'text/html');
		response.end(request.url === '/editor.js' ? script : request.url === '/editor.css' ? css : `<!DOCTYPE html><html><head>
		<style>:root { --vscode-font-family: sans-serif; --vscode-foreground: #202124; --vscode-editor-foreground: #202124; --vscode-editor-background: #fff; --vscode-sideBar-background: #f8fafd; --vscode-focusBorder: #4c78ca; --vscode-editorWidget-background: #fff; }</style>
		<link rel="stylesheet" href="/editor.css"></head><body${request.url?.includes('hosted=1') ? ' class="app-toolbar-hosted"' : ''}>
		<header id="toolbar" aria-label="Document Formatting"></header><div id="info"><span id="filename"></span><span id="hint"></span></div>
		<div id="error" hidden></div><main id="pages"></main><textarea id="keyboard" aria-label="Document text input"></textarea>
		<dialog id="dialog"><form method="dialog"><label id="prompt-label" for="prompt"></label><input id="prompt"><button value="cancel">Cancel</button><button id="confirm" value="ok">Apply</button></form></dialog>
		<script>
		window.officeMessages = [];
		const kind = new URL(location.href).searchParams.get('kind') || 'word';
		window.acquireVsCodeApi = () => ({ postMessage(message) {
			window.officeMessages.push(message);
			if (message.type === 'ready') setTimeout(() => window.dispatchEvent(new MessageEvent('message', { origin: location.origin, data: { type: 'state', kind, name: 'Preview.docx', pages: [], state: { revision: 1, undo: true, redo: false, selectedText: 'Selected passage' } } })), 0);
		} });
		</script><script src="/editor.js"></script></body></html>`);
	});
	before(async () => {
		const result = await build({ absWorkingDir: root, entryPoints: ['extensions/ai-office-common/editor.ts'], bundle: true, platform: 'browser', write: false });
		script = result.outputFiles[0].text;
		appScript = (await build({ absWorkingDir: root, entryPoints: ['src/vs/base/browser/ui/documentToolbar/documentToolbar.ts'], bundle: true, platform: 'browser', format: 'iife', globalName: 'AppDocumentToolbar', write: false })).outputFiles[0].text;
		css = (await readFile(path.join(root, 'src/vs/base/browser/ui/documentToolbar/documentToolbar.css'), 'utf8')) + '\n' + (await readFile(path.join(root, 'extensions/ai-office-common/editor.css'), 'utf8'));
		server.listen(0, '127.0.0.1');
		await once(server, 'listening');
		const location = server.address();
		assert.ok(location && typeof location !== 'string');
		address = `http://127.0.0.1:${location.port}`;
		browser = await chromium.launch({ executablePath: process.env.PLAYWRIGHT_CHROMIUM_EXECUTABLE_PATH, args: ['--no-sandbox'] });
	});
	after(async () => {
		await browser?.close();
		server.close();
	});
	test('Word keeps one row at narrow widths and restores controls when expanded', async () => {
		const page = await browser.newPage();
		try {
			await page.goto(address);
			await page.getByRole('button', { name: 'More Formatting' }).waitFor();
			const controls = page.locator('#toolbar button, #toolbar select, #word-toolbar-overflow button, #word-toolbar-overflow select');
			const controlCount = await controls.count();
			for (const width of [1280, 640, 360, 240, 1280]) {
				await page.setViewportSize({ width, height: 800 });
				await page.waitForFunction(() => {
					const row = document.querySelector('.word-toolbar-row')!;
					return row.scrollWidth <= row.clientWidth;
				});
				const bounds = await page.locator('#toolbar').boundingBox();
				assert.ok(bounds && bounds.height <= 48, `Toolbar wrapped at ${width}px`);
				assert.equal(await controls.count(), controlCount);
			}
			assert.equal(await page.locator('#toolbar').getByRole('button', { name: 'Font Family' }).isVisible(), true);
		} finally { await page.close(); }
	});
	test('Overflow retains formatting, insertion, save and keyboard behavior', async () => {
		const page = await browser.newPage({ viewport: { width: 360, height: 800 } });
		try {
			await page.goto(address);
			const more = page.getByRole('button', { name: 'More Formatting' });
			await more.click();
			await page.getByRole('button', { name: 'Font Family' }).click();
			await page.getByRole('menuitem', { name: 'Arial', exact: true }).click();
			await page.waitForFunction(() => !document.querySelector('#word-toolbar-overflow')!.matches(':popover-open'));
			await more.focus();
			await page.keyboard.press('Enter');
			await page.getByRole('button', { name: 'Paragraph Style', exact: true }).click();
			await page.getByRole('menuitemradio', { name: 'Heading 1' }).click();
			await more.click();
			await page.getByRole('button', { name: 'Insert table', exact: true }).click();
			await page.getByRole('button', { name: '3 columns, 2 rows', exact: true }).click();
			if (!await page.getByRole('button', { name: 'Bold', exact: true }).isVisible()) { await page.getByRole('button', { name: 'More Formatting' }).click(); }
			await page.getByRole('button', { name: 'Bold', exact: true }).click();
			await page.getByRole('button', { name: 'Save (Ctrl+S)', exact: true }).click();
			await page.getByRole('button', { name: 'Undo (Ctrl+Z)', exact: true }).click();
			const messages = await page.evaluate(() => Reflect.get(window, 'officeMessages'));
			assert.deepStrictEqual(messages.filter((message: { type: string }) => message.type !== 'ready'), [
				{ type: 'execute', command: 'format.font', args: { name: 'Arial' } },
				{ type: 'execute', command: 'para.style', args: { style: 'Heading1' } },
				{ type: 'execute', command: 'insert.table', args: { rows: 2, cols: 3, borders: true } },
				{ type: 'execute', command: 'format.bold', args: {} },
				{ type: 'save' }, { type: 'undo' },
			]);
			await more.focus(); await page.keyboard.press('Enter');
			await page.waitForFunction(() => document.querySelector('#word-toolbar-overflow')!.matches(':popover-open'));
			await page.keyboard.press('Escape');
			assert.equal(await more.evaluate(element => element === document.activeElement), true);
			assert.equal(await page.getByRole('button', { name: 'Redo (Ctrl+Shift+Z)', exact: true }).isDisabled(), true);
		} finally { await page.close(); }
	});
	test('Word shows style previews and a table grid with hover and keyboard selection', async () => {
		const page = await browser.newPage({ viewport: { width: 1280, height: 800 } });
		try {
			await page.goto(address);
			await page.getByRole('button', { name: 'Paragraph Style', exact: true }).click();
			const normal = page.getByRole('menuitemradio', { name: 'Normal text' });
			const heading = page.getByRole('menuitemradio', { name: 'Heading 1' });
			assert.ok(await heading.evaluate(el => parseFloat(getComputedStyle(el).fontSize)) > await normal.evaluate(el => parseFloat(getComputedStyle(el).fontSize)));
			await page.screenshot({ path: path.join(root, '.build/word-style-picker.png') });
			await page.getByRole('menuitemradio', { name: 'Subtitle', exact: true }).click();
			await page.getByRole('button', { name: 'More Formatting' }).click();
			await page.getByRole('button', { name: 'Insert table', exact: true }).click();
			await page.getByRole('button', { name: '7 columns, 3 rows', exact: true }).hover();
			assert.equal(await page.locator('.word-table-cell.selected').count(), 21);
			assert.equal(await page.locator('.word-table-size').innerText(), '7 × 3');
			await page.screenshot({ path: path.join(root, '.build/word-table-picker.png') });
			await page.getByRole('button', { name: '1 columns, 1 rows', exact: true }).focus();
			await page.keyboard.press('ArrowRight'); await page.keyboard.press('ArrowDown'); await page.keyboard.press('Enter');
			assert.deepStrictEqual(await page.evaluate(() => Reflect.get(window, 'officeMessages').at(-1)), { type: 'execute', command: 'insert.table', args: { rows: 2, cols: 2, borders: true } });
			assert.equal(await page.locator('#dialog').evaluate(el => (el as HTMLDialogElement).open), false);
			await page.screenshot({ path: path.join(root, '.build/word-menus-browser.png') });
		} finally { await page.close(); }
	});
	test('Word batches text without reordering a following formatting command', async () => {
		const page = await browser.newPage();
		try {
			await page.goto(address);
			await page.getByRole('button', { name: 'Paragraph Style', exact: true }).waitFor();
			await page.evaluate(() => {
				const input = document.getElementById('keyboard') as HTMLTextAreaElement;
				for (const value of ['a', 'b', 'c']) { input.value = value; input.dispatchEvent(new InputEvent('input', { inputType: 'insertText', data: value })); }
				(document.querySelector('[aria-label="Bold"]') as HTMLButtonElement).click();
			});
			assert.deepStrictEqual(await page.evaluate(() => Reflect.get(window, 'officeMessages').slice(-2)), [
				{ type: 'execute', command: 'text.insert', args: { text: 'abc' } },
				{ type: 'execute', command: 'format.bold', args: {} }
			]);
		} finally { await page.close(); }
	});
	test('PDF and slides retain their existing controls', async () => {
		for (const kind of ['pdf', 'slides']) {
			const page = await browser.newPage();
			try {
				await page.goto(`${address}?kind=${kind}`);
				await page.getByRole('button', { name: 'Insert text', exact: true }).waitFor();
				assert.equal(await page.locator('.word-toolbar').count(), 0);
				assert.equal(await page.getByRole('combobox', { name: 'Normal text', exact: true }).isDisabled(), true);
				await page.getByRole('button', { name: 'Insert image', exact: true }).click();
				assert.deepStrictEqual(await page.evaluate(() => Reflect.get(window, 'officeMessages').at(-1)), { type: 'image', page: 1 });
			} finally { await page.close(); }
		}
	});
	test('Agent shimmer follows page coordinates, clears and respects reduced motion', async () => {
		const page = await browser.newPage({ reducedMotion: 'reduce' });
		try {
			await page.goto(address);
			await page.getByRole('button', { name: 'More Formatting' }).waitFor();
			await page.evaluate(() => {
				const send = (data: unknown) => window.dispatchEvent(new MessageEvent('message', { origin: location.origin, data }));
				send({ type: 'state', kind: 'word', name: 'Preview.docx', pages: [[612, 792]], state: { revision: 2 } });
				send({ type: 'agentActivity', rects: [{ page: 0, x: 72, y: 100, width: 300, height: 24 }] });
			});
			const shimmer = page.locator('.agent-shimmer-text');
			assert.equal(await page.locator('.agent-shimmer-border').count(), 1);
			assert.equal(await shimmer.count(), 1);
			assert.equal(await shimmer.evaluate(element => getComputedStyle(element).pointerEvents), 'none');
			assert.equal(await shimmer.evaluate(element => getComputedStyle(element).animationName), 'none');
			assert.equal(await page.locator('#pages').getAttribute('aria-busy'), 'true');
			assert.match(await page.getByRole('status').innerText(), /Agent is working/);
			const before = await shimmer.boundingBox();
			await page.setViewportSize({ width: 700, height: 900 });
			await page.waitForTimeout(250);
			const after = await shimmer.boundingBox();
			assert.ok(before && after && after.width < before.width);
			await page.evaluate(() => window.dispatchEvent(new MessageEvent('message', { origin: location.origin, data: { type: 'agentActivity', rects: [] } })));
			assert.equal(await shimmer.count(), 0);
			assert.equal(await page.locator('#pages').getAttribute('aria-busy'), 'false');
			assert.equal(await page.locator('.agent-status').innerText(), '');
		} finally { await page.close(); }
	});
	test('Paragraph panel patches only changed fields, guards selection and exposes accessible ruler controls', async () => {
		const page = await browser.newPage({ viewport: { width: 1400, height: 1000 } });
		try {
			await page.goto(address); await page.getByRole('button', { name: 'More Formatting' }).waitFor();
			const state = { revision: 4, undo: true, selection: { anchor: 1, focus: 1 }, caret: { page: 0, x: 72, y: 72, height: 12 }, paragraph: { id: 'internal-1', properties: { alignment: 'left', indents: { left: 0, right: 0, firstLine: 0 }, spacing: { before: 0, after: 8 }, lineSpacing: { rule: 'multiple', value: 1.15 }, tabs: [{ pos: 144, align: 'right', leader: 'dot' }], keepNext: false, keepLines: false, widowControl: true, pageBreakBefore: false }, mixed: [], ruler: { left: 72, width: 468 } } };
			const update = async (value: typeof state) => page.evaluate(state => window.dispatchEvent(new MessageEvent('message', { origin: location.origin, data: { type: 'state', kind: 'word', name: 'Paragraph.docx', pages: [[612, 792]], state } })), value);
			await update(state);
			const properties = page.getByRole('button', { name: /Paragraph Properties ·/ });
			if (!await properties.isVisible()) { await page.getByRole('button', { name: 'More Formatting' }).click(); }
			await properties.click();
			const dialog = page.getByRole('dialog', { name: 'Paragraph Properties' });
			await dialog.getByLabel('Left Indent', { exact: true }).fill('36');
			await dialog.getByLabel('Keep with Next', { exact: true }).check();
			await dialog.getByRole('button', { name: 'Apply', exact: true }).click();
			const edits = await page.evaluate(() => (window as unknown as { officeMessages: { command?: string; args?: object }[] }).officeMessages.filter(m => m.command === 'para.properties'));
			assert.deepEqual(edits[0].args, { indents: { left: 36 }, keepNext: true, ifRevision: 4, ifSelection: state.selection });
			const left = page.getByRole('slider', { name: 'Left Indent · Arrow Keys Adjust', exact: true }); await left.focus(); await page.keyboard.press('ArrowRight');
			assert.equal(await page.evaluate(() => (window as unknown as { officeMessages: { command?: string; args?: { indents?: { left?: number } } }[] }).officeMessages.filter(m => m.command === 'para.properties').at(-1)?.args?.indents?.left), 1);
			if (!await properties.isVisible()) { await page.getByRole('button', { name: 'More Formatting' }).click(); }
			await properties.click();
			await update({ ...state, revision: 5 });
			assert.equal(await dialog.getByRole('button', { name: 'Apply', exact: true }).isDisabled(), true);
			assert.match(await dialog.getByRole('alert').innerText(), /selection changed/);
			await page.keyboard.press('Escape'); assert.equal(await dialog.isVisible(), false);
			await page.setViewportSize({ width: 420, height: 800 });
			await page.getByRole('button', { name: 'More Formatting' }).click(); await properties.click();
			const bounds = await dialog.boundingBox(); assert.ok(bounds && bounds.width <= 420);
			await dialog.getByRole('button', { name: 'Cancel', exact: true }).click();
		} finally { await page.close(); }
	});

	test('Word button palettes and tooltips dispatch native formatting commands', async () => {
		const page = await browser.newPage({ viewport: { width: 1600, height: 1000 } });
		page.setDefaultTimeout(10000);
		try {
			await page.goto(address);
			await page.evaluate(() => { document.body.style.setProperty('--vscode-foreground', '#cccccc'); document.body.style.setProperty('--vscode-editorWidget-background', '#252526'); });
			await page.getByRole('button', { name: 'Text Color', exact: true }).click();
			assert.equal(await page.getByRole('menuitem', { name: 'Red', exact: true }).evaluate(element => getComputedStyle(element).backgroundColor), 'rgba(0, 0, 0, 0)');
			await page.getByRole('menuitem', { name: 'Red', exact: true }).click();
			await page.getByRole('button', { name: 'Text Highlight', exact: true }).click();
			await page.getByRole('menuitem', { name: 'Yellow', exact: true }).click();
			await page.getByRole('button', { name: 'Strikethrough', exact: true }).click();
			await page.getByRole('button', { name: 'Font Size', exact: true }).click();
			await page.keyboard.press('End'); await page.keyboard.press('Enter');
			const commands = await page.evaluate(() => Reflect.get(window, 'officeMessages').filter((message: { type: string }) => message.type === 'execute'));
			assert.deepEqual(commands.map((message: { command: string; args: object }) => [message.command, message.args]), [
				['format.color', { color: 'C00000' }], ['format.highlight', { color: 'yellow' }], ['format.strikethrough', {}], ['format.size', { size: 48 }]
			]);
			await page.getByRole('button', { name: 'Bold', exact: true }).hover();
			await page.getByRole('tooltip').waitFor({ state: 'visible' });
			assert.match(await page.getByRole('tooltip').innerText(), /Bold/);
			await page.keyboard.press('Escape');
			await page.getByRole('tooltip').waitFor({ state: 'hidden' });
			await page.emulateMedia({ reducedMotion: 'reduce' });
			assert.equal(await page.getByRole('button', { name: 'Bold', exact: true }).evaluate(element => getComputedStyle(element).transitionDuration), '0s');
		} finally { await page.close(); }
	});

	test('App toolbar reuses document actions while the in-tab toolbar stays hidden and stale selections are refused', async () => {
		const page = await browser.newPage({ viewport: { width: 1500, height: 1000 } });
		try {
			await page.goto(address + '/?hosted=1');
			await page.waitForFunction(() => Reflect.get(window, 'officeMessages').some((message: { type: string }) => message.type === 'toolbarState'));
			await page.addScriptTag({ content: appScript });
			await page.evaluate(() => {
				const snapshot = Reflect.get(window, 'officeMessages').findLast((message: { type: string }) => message.type === 'toolbarState');
				const host = document.createElement('div'); host.className = 'sangam-document-toolbar'; host.style.display = 'flex'; const toolbar = document.createElement('header'); host.append(toolbar); document.body.prepend(host);
				Reflect.get(window, 'AppDocumentToolbar').createDocumentToolbar(toolbar, snapshot.controls, (action: object) => window.dispatchEvent(new MessageEvent('message', { origin: location.origin, data: { type: 'toolbarAction', context: snapshot.context, action } })), () => {});
			});
			assert.equal(await page.locator('#toolbar').isVisible(), false);
			assert.equal((await page.locator('.sangam-document-toolbar > header').boundingBox())?.height, 42);
			assert.equal(await page.getByRole('toolbar', { name: 'Document formatting' }).count(), 1);
			if (!await page.getByRole('button', { name: 'Bold', exact: true }).isVisible()) { await page.getByRole('button', { name: 'More Formatting' }).click(); }
			await page.getByRole('button', { name: 'Bold', exact: true }).click();
			await page.getByRole('button', { name: 'Paragraph Style', exact: true }).click();
			await page.getByRole('menuitemradio', { name: 'Heading 2' }).click();
			const before = await page.evaluate(() => Reflect.get(window, 'officeMessages').filter((message: { type: string }) => message.type === 'execute'));
			assert.deepEqual(before.map((message: { command: string }) => message.command), ['format.bold', 'para.style']);
			await page.evaluate(() => window.dispatchEvent(new MessageEvent('message', { origin: location.origin, data: { type: 'state', kind: 'word', name: 'Preview.docx', pages: [], state: { revision: 2, undo: true } } })));
			await page.getByRole('button', { name: 'Italic', exact: true }).click();
			assert.match(await page.locator('#error').innerText(), /selection changed/);
			assert.equal(await page.evaluate(() => Reflect.get(window, 'officeMessages').filter((message: { type: string }) => message.type === 'execute').length), before.length);
			await page.setViewportSize({ width: 400, height: 800 });
			const more = page.getByRole('button', { name: 'More Formatting' });
			await page.waitForFunction(() => { const host = document.querySelector('.sangam-document-toolbar')!; const more = document.querySelector('.word-more-button')!; return more.getBoundingClientRect().right <= host.getBoundingClientRect().right; });
			await more.click();
			assert.equal(await page.getByRole('button', { name: 'Font Family' }).isVisible(), true);
		} finally { await page.close(); }
	});

});
