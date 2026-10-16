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
	let css: string;
	const server = createServer((request, response) => {
		response.setHeader('Content-Type', request.url === '/editor.js' ? 'text/javascript' : request.url === '/editor.css' ? 'text/css' : 'text/html');
		response.end(request.url === '/editor.js' ? script : request.url === '/editor.css' ? css : `<!DOCTYPE html><html><head>
		<style>:root { --vscode-font-family: sans-serif; --vscode-foreground: #202124; --vscode-editor-foreground: #202124; --vscode-editor-background: #fff; --vscode-sideBar-background: #f8fafd; --vscode-focusBorder: #4c78ca; --vscode-editorWidget-background: #fff; }</style>
		<link rel="stylesheet" href="/editor.css"></head><body>
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
		css = await readFile(path.join(root, 'extensions/ai-office-common/editor.css'), 'utf8');
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
			assert.equal(await page.locator('#toolbar').getByRole('combobox', { name: 'Font Family' }).isVisible(), true);
		} finally { await page.close(); }
	});
	test('Overflow retains formatting, insertion, save and keyboard behavior', async () => {
		const page = await browser.newPage({ viewport: { width: 360, height: 800 } });
		try {
			await page.goto(address);
			const more = page.getByRole('button', { name: 'More Formatting' });
			await more.click();
			await page.getByRole('combobox', { name: 'Font Family' }).selectOption('Arial');
			await page.waitForFunction(() => !document.querySelector('#word-toolbar-overflow')!.matches(':popover-open'));
			await more.focus();
			await page.keyboard.press('Enter');
			await page.getByRole('combobox', { name: 'Paragraph Style' }).selectOption('Heading1');
			await more.click();
			await page.getByRole('button', { name: 'Insert table', exact: true }).click();
			await page.getByLabel('Table size: rows x columns').fill('2 x 3');
			await page.getByRole('button', { name: 'Apply', exact: true }).click();
			await page.getByRole('button', { name: 'Bold', exact: true }).click();
			await page.getByRole('button', { name: 'Save (Ctrl+S)', exact: true }).click();
			await page.getByRole('button', { name: 'Undo (Ctrl+Z)', exact: true }).click();
			const messages = await page.evaluate(() => Reflect.get(window, 'officeMessages'));
			assert.deepStrictEqual(messages.filter((message: { type: string }) => message.type !== 'ready'), [
				{ type: 'execute', command: 'format.font', args: { name: 'Arial' } },
				{ type: 'execute', command: 'para.style', args: { style: 'Heading1' } },
				{ type: 'execute', command: 'insert.table', args: { rows: 2, cols: 3 } },
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
});
