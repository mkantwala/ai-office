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
	test('Word typing keeps display resolution without a blurry draft or delayed refinement', async () => {
		const page = await browser.newPage({ viewport: { width: 1600, height: 1000 }, deviceScaleFactor: 2 });
		page.setDefaultTimeout(10000);
		try {
			await page.goto(address);
			await page.getByRole('button', { name: 'More Formatting' }).waitFor();
			await page.evaluate(() => {
				const state = (revision: number) => window.dispatchEvent(new MessageEvent('message', { origin: location.origin, data: {
					type: 'state', kind: 'word', name: 'Typing.docx', pages: [[612, 792]], command: 'text.insert',
					state: { revision, undo: true, redo: false, caret: { page: 0, x: 72, y: 72, height: 12, width: 1 } }
				} }));
				Reflect.set(window, 'typingState', state); state(2);
			});
			await page.waitForFunction(() => Reflect.get(window, 'officeMessages').some((m: { type: string }) => m.type === 'render'));
			const initial = await page.evaluate(() => Reflect.get(window, 'officeMessages').find((m: { type: string }) => m.type === 'render'));
			assert.ok(initial.scale > 1);
			await page.evaluate(request => window.dispatchEvent(new MessageEvent('message', { origin: location.origin, data: {
				type: 'rendered', request: request.request, page: 0, result: { width: 1, height: 1, pixels: new Uint8Array([255, 255, 255, 255]) }
			} })), initial);
			await page.locator('#keyboard').focus(); await page.keyboard.type('Sharp');
			await page.evaluate(() => Reflect.get(window, 'typingState')(3));
			await page.waitForFunction(() => Reflect.get(window, 'officeMessages').filter((m: { type: string }) => m.type === 'render').length === 2);
			const typing = await page.evaluate(() => Reflect.get(window, 'officeMessages').filter((m: { type: string }) => m.type === 'render').at(-1));
			assert.equal(typing.scale, initial.scale, 'Typing must retain the physical-pixel scale');
			await page.evaluate(request => window.dispatchEvent(new MessageEvent('message', { origin: location.origin, data: {
				type: 'rendered', request: request.request, page: 0, result: { width: 1, height: 1, pixels: new Uint8Array([255, 255, 255, 255]) }
			} })), typing);
			await page.waitForTimeout(300);
			assert.equal(await page.evaluate(() => Reflect.get(window, 'officeMessages').filter((m: { type: string }) => m.type === 'render').length), 2);
		} finally { await page.close(); }
	});

	test('Word keeps one row at narrow widths and restores controls when expanded', async () => {
		const page = await browser.newPage();
		page.setDefaultTimeout(10000);
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
	test('Overflow retains formatting, insertion and keyboard behavior', async () => {
		const page = await browser.newPage({ viewport: { width: 360, height: 800 } });
		page.setDefaultTimeout(10000);
		try {
			await page.goto(address);
			const more = page.getByRole('button', { name: 'More Formatting' });
			await more.click();
			await page.getByRole('button', { name: 'Font Family' }).click();
			await page.getByRole('menuitemradio', { name: 'Arial', exact: true }).click();
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
			assert.equal(await page.getByRole('button', { name: 'Save (Ctrl+S)', exact: true }).count(), 0);
			await page.keyboard.press('Control+s');
			if (!await page.getByRole('button', { name: 'Undo (Ctrl+Z)', exact: true }).isVisible()) { await more.click(); }
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
			assert.equal(await page.getByRole('button', { name: 'Redo (Ctrl+Shift+Z)', exact: true, includeHidden: true }).isDisabled(), true);
		} finally { await page.close(); }
	});
	test('Word shows style previews and a table grid with hover and keyboard selection', async () => {
		const page = await browser.newPage({ viewport: { width: 1280, height: 800 } });
		page.setDefaultTimeout(10000);
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
		page.setDefaultTimeout(10000);
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
		page.setDefaultTimeout(10000);
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
		page.setDefaultTimeout(10000);
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
		page.setDefaultTimeout(10000);
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

	test('Ruler matches the page at narrow widths and high DPI, uses inch graduations and retains point-based edits', async () => {
		const page = await browser.newPage({ viewport: { width: 1280, height: 1000 }, deviceScaleFactor: 2 });
		page.setDefaultTimeout(10000);
		try {
			await page.goto(address); await page.getByRole('button', { name: 'More Formatting' }).waitFor();
			await page.evaluate(() => window.dispatchEvent(new MessageEvent('message', { origin: location.origin, data: {
				type: 'state', kind: 'word', name: 'Ruler.docx', pages: [[612, 792]], state: { revision: 4, selection: { anchor: 1, focus: 1 }, caret: { page: 0, x: 72, y: 72, height: 12 }, paragraph: { id: 'p1', properties: { indents: { left: 0, right: 0, firstLine: 0 }, tabs: [] }, mixed: [], ruler: { left: 72, width: 468 } } }
			} })));
			const ruler = page.getByRole('group', { name: 'Paragraph Ruler · Inches', exact: true });
			for (const width of [1280, 420, 800]) {
				await page.setViewportSize({ width, height: 1000 });
				await page.waitForFunction(() => Math.abs(document.querySelector('.word-ruler')!.getBoundingClientRect().width - document.querySelector('.page')!.getBoundingClientRect().width) < 1);
				const bounds = await ruler.boundingBox(); const paper = await page.locator('.page').boundingBox();
				assert.ok(bounds && paper && Math.abs(bounds.x - paper.x) < 1 && bounds.height <= 30);
				assert.ok(await ruler.locator('.word-ruler-tick').count() > 20);
				assert.doesNotMatch((await ruler.locator('.word-ruler-number').allTextContents()).join(','), /72|144|216/);
				const first = await ruler.getByRole('slider', { name: 'First-Line Indent · Arrow Keys Adjust', exact: true }).boundingBox();
				assert.ok(first && bounds && Math.abs(first.x + first.width / 2 - (bounds.x + 72 * bounds.width / 612)) < 1);
			}
			const left = ruler.getByRole('slider', { name: 'Left Indent · Arrow Keys Adjust', exact: true });
			const handle = await left.boundingBox(); const bounds = await ruler.boundingBox(); assert.ok(handle && bounds);
			await page.mouse.move(handle.x + handle.width / 2, handle.y + handle.height / 2); await page.mouse.down();
			await page.mouse.move(handle.x + handle.width / 2 + 36 * bounds.width / 612, handle.y + handle.height / 2);
			assert.equal(await left.getAttribute('aria-valuetext'), '0.5 in (36 points)');
			await page.mouse.up();
			const edits = await page.evaluate(() => Reflect.get(window, 'officeMessages').filter((message: { command?: string }) => message.command === 'para.properties'));
			assert.equal(edits.at(-1).args.indents.left, 36);
			assert.equal(edits.at(-1).args.ifRevision, 4);
		} finally { await page.close(); }
	});

	test('Long font names preserve the toolbar layout, selected options and reduced-motion behavior', async () => {
		const page = await browser.newPage({ viewport: { width: 1600, height: 1000 }, reducedMotion: 'no-preference' });
		page.setDefaultTimeout(10000);
		try {
			await page.goto(address);
			const font = page.getByRole('button', { name: 'Font Family', exact: true });
			const before = await font.boundingBox(); assert.ok(before);
			await font.click();
			assert.equal(await page.getByRole('menu', { name: 'Font Family', exact: true }).evaluate(e => getComputedStyle(e).animationName), 'word-menu-open');
			await page.getByRole('menuitemradio', { name: 'Times New Roman', exact: true }).click();
			const after = await font.boundingBox(); assert.ok(after); assert.equal(after.width, before.width);
			assert.equal(await font.locator('.word-control-value').innerText(), 'Times New Roman');
			assert.equal(await font.locator('.word-control-value').evaluate(e => getComputedStyle(e).textOverflow), 'ellipsis');
			await font.click(); assert.equal(await page.getByRole('menuitemradio', { name: 'Times New Roman', exact: true }).getAttribute('aria-checked'), 'true');
			await page.keyboard.press('Escape');
			for (const width of [640, 360, 1600]) {
				await page.setViewportSize({ width, height: 800 });
				await page.waitForFunction(() => { const row = document.querySelector('.word-toolbar-row')!; return row.scrollWidth <= row.clientWidth; });
			}
			await page.emulateMedia({ reducedMotion: 'reduce' }); await font.click();
			assert.equal(await page.getByRole('menu', { name: 'Font Family', exact: true }).evaluate(e => getComputedStyle(e).animationName), 'none');
			assert.equal(await page.getByRole('menuitemradio', { name: 'Times New Roman', exact: true }).getAttribute('aria-checked'), 'true');
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
				['format.color', { color: 'C00000' }], ['format.highlight', { color: 'yellow' }], ['format.strikethrough', {}], ['format.size', { size: 72 }]
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

	test('Direct toolbar messages reject stale document instances and preserve typing order', async () => {
		const page = await browser.newPage(); page.setDefaultTimeout(10000);
		try {
			await page.goto(address + '?hosted=1');
			await page.waitForFunction(() => Reflect.get(window, 'officeMessages').some((m: { type: string }) => m.type === 'toolbarState'));
			await page.evaluate(() => {
				document.body.dataset.documentInstance = 'current-instance';
				const context = Reflect.get(window, 'officeMessages').filter((m: { type: string }) => m.type === 'toolbarState').at(-1).context;
				window.dispatchEvent(new MessageEvent('message', { origin: location.origin, data: { type: 'toolbarAction', instance: 'old-instance', context, action: { control: 'bold' } } }));
			});
			assert.equal(await page.evaluate(() => Reflect.get(window, 'officeMessages').filter((m: { type: string }) => m.type === 'execute').length), 0);
			await page.evaluate(() => {
				const context = Reflect.get(window, 'officeMessages').filter((m: { type: string }) => m.type === 'toolbarState').at(-1).context;
				window.dispatchEvent(new MessageEvent('message', { origin: location.origin, data: { type: 'toolbarAction', instance: 'current-instance', context, action: { control: 'bold' } } }));
			});
			await page.keyboard.insertText('After action');
			await page.waitForFunction(() => Reflect.get(window, 'officeMessages').filter((m: { type: string }) => m.type === 'execute').length === 2);
			assert.deepEqual(await page.evaluate(() => Reflect.get(window, 'officeMessages').filter((m: { type: string }) => m.type === 'execute').map((m: { command: string }) => m.command)), ['format.bold', 'text.insert']);
		} finally { await page.close(); }
	});

	test('App toolbar reuses document actions while the in-tab toolbar stays hidden and stale selections are refused', async () => {
		const page = await browser.newPage({ viewport: { width: 1500, height: 1000 } });
		page.setDefaultTimeout(10000);
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

	test('Font size steppers retain selection commands and Save/History stay out of the toolbar', async () => {
		const page = await browser.newPage({ viewport: { width: 1600, height: 1000 } });
		page.setDefaultTimeout(10000);
		try {
			await page.goto(address);
			await page.getByRole('button', { name: 'Increase Font Size', exact: true }).click();
			await page.getByRole('button', { name: 'Increase Font Size', exact: true }).click();
			await page.getByRole('button', { name: 'Decrease Font Size', exact: true }).click();
			const commands = await page.evaluate(() => Reflect.get(window, 'officeMessages').filter((m: { command?: string }) => m.command === 'format.size'));
			assert.deepEqual(commands.map((m: { args: { size: number } }) => m.args.size), [12, 13, 12]);
			assert.equal(await page.locator('#keyboard').evaluate(e => e === document.activeElement), true);
			await page.keyboard.type('after size');
			assert.equal(await page.getByRole('button', { name: 'Save (Ctrl+S)', exact: true }).count(), 0);
			assert.equal(await page.getByRole('button', { name: /Document Versions/ }).count(), 0);
		} finally { await page.close(); }
	});

	test('Table hover handles, context actions and border drags use native page coordinates and dismiss safely', async () => {
		const page = await browser.newPage({ viewport: { width: 1200, height: 1000 } });
		page.setDefaultTimeout(10000);
		try {
			await page.goto(address);
			await page.getByRole('button', { name: 'More Formatting' }).waitFor();
			await page.evaluate(() => window.dispatchEvent(new MessageEvent('message', { origin: location.origin, data: { type: 'state', kind: 'word', name: 'Table.docx', pages: [[612, 792]], state: { revision: 2, undo: true } } })));
			await page.waitForFunction(() => Reflect.get(window, 'officeMessages').some((m: { type: string }) => m.type === 'render'));
			await page.evaluate(() => {
				const request = Reflect.get(window, 'officeMessages').filter((m: { type: string }) => m.type === 'render').at(-1).request;
				window.dispatchEvent(new MessageEvent('message', { origin: location.origin, data: { type: 'rendered', request, page: 0, result: { width: 1, height: 1, pixels: new Uint8Array([255,255,255,255]), tables: [{ cols: [72, 250, 540], rows: [{ row: 0, top: 100, bottom: 140, edges: [0, 2] }, { row: 1, top: 140, bottom: 180, edges: [0, 1, 2] }], endsHere: true }] } } }));
			});
			const paper = await page.locator('.page.word').boundingBox(); assert.ok(paper); const scale = paper.width / 612;
			await page.mouse.move(paper.x + 400 * scale, paper.y + 160 * scale);
			assert.equal(await page.getByRole('button', { name: 'Row options', exact: true }).count(), 0, 'Edge controls stay quiet over cell text');
			await page.mouse.move(paper.x + 64 * scale, paper.y + 160 * scale);
			assert.equal(await page.getByRole('button', { name: 'Row options', exact: true }).isVisible(), true);
			await page.mouse.click(paper.x + 400 * scale, paper.y + 160 * scale, { button: 'right' });
			await page.getByRole('menuitem', { name: 'Sort column descending', exact: true }).click();
			let commands = await page.evaluate(() => Reflect.get(window, 'officeMessages').filter((m: { type: string }) => m.type === 'execute'));
			assert.deepEqual(commands.at(-1), { type: 'execute', command: 'table.sort', args: { column: 1, descending: true } });
			await page.mouse.move(paper.x + 250 * scale, paper.y + 120 * scale);
			assert.equal(await page.locator('.page.word').evaluate(e => getComputedStyle(e).cursor), 'text', 'Merged cells have no internal resize edge');
			await page.mouse.move(paper.x + 250 * scale, paper.y + 160 * scale);
			assert.equal(await page.locator('.page.word').evaluate(e => getComputedStyle(e).cursor), 'col-resize');
			await page.mouse.down(); assert.equal(await page.locator('#keyboard').evaluate(e => e === document.activeElement), true); await page.mouse.move(paper.x + 280 * scale, paper.y + 160 * scale);
			await page.waitForFunction(() => Reflect.get(window, 'officeMessages').some((m: { type: string }) => m.type === 'tablePreview'));
			await page.evaluate(() => {
				const m = Reflect.get(window, 'officeMessages').filter((m: { type: string }) => m.type === 'tablePreview').at(-1);
				window.dispatchEvent(new MessageEvent('message', { origin: location.origin, data: { type: 'tablePreview', request: m.request, revision: m.args.revision, result: { kind: 'column', rects: [{ x: 280, y: 100, width: 0, height: 80 }] } } }));
			});
			assert.equal(await page.locator('.word-table-preview').count(), 1);
			assert.equal(await page.evaluate(() => Reflect.get(window, 'officeMessages').filter((m: { command?: string }) => m.command === 'table.resizeDrag').length), 0, 'Dragging must not mutate until release');
			await page.mouse.up();
			assert.equal(await page.locator('.word-table-preview').count(), 0);
			assert.equal(await page.locator('#keyboard').evaluate(e => e === document.activeElement), true);
			await page.mouse.move(paper.x + 400 * scale, paper.y + 140 * scale);
			assert.equal(await page.locator('.page.word').evaluate(e => getComputedStyle(e).cursor), 'row-resize');
			commands = await page.evaluate(() => Reflect.get(window, 'officeMessages').filter((m: { type: string }) => m.type === 'execute'));
			assert.equal(commands.at(-1).command, 'table.resizeDrag'); assert.ok(Math.abs(commands.at(-1).args.toX - 280) < 1);
			await page.mouse.move(paper.x + 250 * scale, paper.y + 160 * scale); await page.mouse.down(); await page.mouse.move(paper.x + 300 * scale, paper.y + 160 * scale);
			await page.keyboard.press('Escape'); await page.mouse.up();
			assert.equal(await page.evaluate(() => Reflect.get(window, 'officeMessages').filter((m: { command?: string }) => m.command === 'table.resizeDrag').length), 1, 'Escape cancels without an edit');
			await page.mouse.dblclick(paper.x + 250 * scale, paper.y + 160 * scale);
			commands = await page.evaluate(() => Reflect.get(window, 'officeMessages').filter((m: { type: string }) => m.type === 'execute'));
			assert.equal(commands.at(-1).command, 'table.fitColumn');
			await page.mouse.click(paper.x + 400 * scale, paper.y + 160 * scale, { button: 'right' });
			await page.getByRole('menuitem', { name: 'Set row height…', exact: true }).click();
			await page.locator('#prompt').fill('30'); await page.locator('#confirm').click();
			assert.equal(await page.locator('#keyboard').evaluate(e => e === document.activeElement), true);
			await page.keyboard.type('continue');
			await page.waitForFunction(() => Reflect.get(window, 'officeMessages').some((m: { command?: string }) => m.command === 'text.insert'));
			await page.mouse.click(paper.x + 400 * scale, paper.y + 160 * scale, { button: 'right' });
			await page.keyboard.press('Escape'); assert.equal(await page.getByRole('menu', { name: 'Table options', exact: true }).isVisible(), false);
		} finally { await page.close(); }
	});

});
