/*---------------------------------------------------------------------------------------------
 *  Copyright (c) Microsoft Corporation. All rights reserved.
 *  Licensed under the MIT License. See License.txt in the project root for license information.
 *--------------------------------------------------------------------------------------------*/

import assert from 'assert';
import fs from 'fs';
import path from 'path';
import { suite, test } from 'node:test';
import { isBuiltinExtensionEnabled, isBuiltinExtensionBuildEnabled } from '../builtinExtensionPolicy.ts';
import { scanBuiltinExtensions } from '../extensions.ts';

const root = path.resolve(import.meta.dirname, '../../..');
const manifests = fs.readdirSync(path.join(root, 'extensions')).flatMap(directory => {
	const file = path.join(root, 'extensions', directory, 'package.json');
	if (!fs.existsSync(file)) {
		return [];
	}
	const manifest: { publisher: string; name: string; extensionDependencies?: string[] } = JSON.parse(fs.readFileSync(file, 'utf8'));
	return [{ directory, manifest }];
});

suite('AI Office built-in extensions', () => {
	test('ships document, agent and shell essentials instead of programming tools', () => {
		const expected = [
			'ai-office-pdf', 'ai-office-slides', 'configuration-editing', 'copilot', 'diff',
			'git', 'git-base', 'github', 'github-authentication', 'json', 'json-language-features',
			'log', 'markdown-basics', 'markdown-language-features', 'markdown-math', 'media-preview',
			'mermaid-markdown-features', 'microsoft-authentication', 'prompt-basics', 'search-result',
			'simple-browser', 'theme-abyss', 'theme-defaults', 'theme-kimbie-dark', 'theme-modern-icons',
			'theme-monokai', 'theme-monokai-dimmed', 'theme-quietlight', 'theme-red', 'theme-seti',
			'theme-solarized-dark', 'theme-solarized-light', 'theme-tomorrow-night-blue', 'yaml',
		];
		assert.deepStrictEqual(manifests.filter(({ manifest }) => isBuiltinExtensionEnabled(manifest)).map(({ directory }) => directory).sort(), expected.sort());
	});

	test('kept extensions have all their declared extension dependencies', () => {
		const enabled = manifests.filter(({ manifest }) => isBuiltinExtensionEnabled(manifest));
		const ids = new Set(enabled.map(({ manifest }) => `${manifest.publisher}.${manifest.name}`.toLowerCase()));
		const missing = enabled.flatMap(({ manifest }) => (manifest.extensionDependencies ?? [])
			.filter(id => !ids.has(id.toLowerCase()))
			.map(id => `${manifest.publisher}.${manifest.name} -> ${id}`));
		assert.deepStrictEqual(missing, []);
	});

	test('nested desktop, browser and media builds obey the packaging policy', () => {
		for (const { directory, manifest } of manifests) {
			for (const input of ['tsconfig.json', 'client/tsconfig.json', 'esbuild.browser.mts', 'preview-src/esbuild.webview.mts']) {
				const buildPath = path.join('extensions', directory, input);
				assert.strictEqual(isBuiltinExtensionBuildEnabled(buildPath), isBuiltinExtensionEnabled(manifest), buildPath);
				assert.strictEqual(isBuiltinExtensionBuildEnabled(path.join(root, buildPath)), isBuiltinExtensionEnabled(manifest), buildPath);
			}
		}
		assert.strictEqual(isBuiltinExtensionBuildEnabled('.vscode/extensions/vscode-extras/tsconfig.json'), true);
	});

	test('extension IDs are case insensitive', () => {
		assert.strictEqual(isBuiltinExtensionEnabled({ publisher: 'VSCode', name: 'PYTHON' }), false);
		assert.strictEqual(isBuiltinExtensionEnabled({ publisher: 'GitHub', name: 'copilot-chat' }), true);
	});

	test('browser build discovery excludes programming extensions even when their source is present', () => {
		const extensions = scanBuiltinExtensions(path.join(root, 'extensions'));
		assert.ok(extensions.some(extension => extension.extensionPath === 'markdown-language-features'));
		const enabledDirectories = new Set(manifests.filter(({ manifest }) => isBuiltinExtensionEnabled(manifest)).map(({ directory }) => directory));
		assert.deepStrictEqual(extensions.filter(extension => !enabledDirectories.has(extension.extensionPath)).map(extension => extension.extensionPath), []);
	});
});
