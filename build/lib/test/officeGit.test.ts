/*---------------------------------------------------------------------------------------------
 *  Copyright (c) Microsoft Corporation. All rights reserved.
 *  Licensed under the MIT License. See License.txt in the project root for license information.
 *--------------------------------------------------------------------------------------------*/

import assert from 'assert';
import path from 'path';
import { test } from 'node:test';
import { build } from 'esbuild';

const root = path.resolve(import.meta.dirname, '../../..');
for (const [entryPoint, replacement] of [
	['src/vs/platform/agentHost/node/agentHostServices.ts', 'src/vs/platform/agentHost/common/disabledAgentHostGitService.ts'],
	['src/vs/code/electron-utility/sharedProcess/sharedProcessMain.ts', 'src/vs/platform/git/common/disabledLocalGitService.ts'],
	['src/vs/platform/agentHost/node/shared/serverToolGroups.ts', 'src/vs/platform/agentHost/node/shared/artifactServerTools.ts'],
]) {
	test(`${entryPoint} excludes Git process implementations`, async () => {
		const result = await build({
			absWorkingDir: root, entryPoints: [entryPoint], bundle: true, platform: 'node', format: 'esm',
			packages: 'external', write: false, metafile: true, outfile: '.build/office-git-check.js',
			loader: { '.css': 'empty', '.svg': 'dataurl', '.png': 'dataurl', '.ttf': 'dataurl', '.wasm': 'binary', '.sh': 'text', '.ps1': 'text', '.scm': 'text' },
		});
		const inputs = new Set(Object.keys(result.metafile.inputs).map(file => file.replaceAll('\\', '/')));
		assert.ok(inputs.has(replacement));
		assert.deepStrictEqual([
			'src/vs/platform/git/node/localGitService.ts',
			'src/vs/platform/agentHost/node/agentHostGitService.ts',
		].filter(file => inputs.has(file)), []);
		if (entryPoint.endsWith('/serverToolGroups.ts')) {
			assert.deepStrictEqual([
				'src/vs/platform/agentHost/node/shared/sessionIsolationTools.ts',
				'src/vs/platform/agentHost/node/shared/agentMergeServerTools.ts',
			].filter(file => inputs.has(file)), []);
		}
	});
}
