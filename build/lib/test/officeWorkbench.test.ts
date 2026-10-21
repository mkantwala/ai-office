/*---------------------------------------------------------------------------------------------
 *  Copyright (c) Microsoft Corporation. All rights reserved.
 *  Licensed under the MIT License. See License.txt in the project root for license information.
 *--------------------------------------------------------------------------------------------*/

import assert from 'assert';
import path from 'path';
import { suite, test } from 'node:test';
import { build } from 'esbuild';

const root = path.resolve(import.meta.dirname, '../../..');
const excludedContributions = [
	'notebook/browser/notebook.contribution.ts',
	'interactive/browser/interactive.contribution.ts',
	'replNotebook/browser/repl.contribution.ts',
	'debug/browser/debug.contribution.ts',
	'debug/browser/debug.service.contribution.ts',
	'testing/browser/testing.contribution.ts',
	'tasks/browser/task.contribution.ts',
	'tasks/browser/taskService.ts',
	'tasks/electron-browser/taskService.ts',
	'mergeEditor/browser/mergeEditor.contribution.ts',
	'mergeEditor/electron-browser/mergeEditor.contribution.ts',
	'issue/browser/issue.contribution.ts',
	'issue/electron-browser/issue.contribution.ts',
	'remote/common/remote.contribution.ts',
	'remote/browser/remote.contribution.ts',
	'remote/electron-browser/remote.contribution.ts',
	'remote/browser/remoteStartEntry.contribution.ts',
	'remoteTunnel/electron-browser/remoteTunnel.contribution.ts',
	'snippets/browser/snippets.contribution.ts',
	'snippets/browser/snippets.service.contribution.ts',
	'git/browser/git.contributions.ts',
	'scm/browser/scm.contribution.ts',
	'scm/browser/scm.service.contribution.ts',
	'scm/browser/quickDiff.contribution.ts',
	'github/browser/githubLinkPresentation.contribution.ts',
].map(file => `src/vs/workbench/contrib/${file}`);

suite('AI Office workbench dependencies', () => {
	for (const application of ['workbench', 'sessions']) {
		for (const target of ['desktop', 'web']) {
			test(`${application} ${target} excludes developer UI and retains document and agent infrastructure`, async () => {
				const result = await build({
					absWorkingDir: root,
					entryPoints: [`src/vs/${application}/${application}.${target}.main.ts`],
					bundle: true,
					platform: 'browser',
					format: 'esm',
					packages: 'external',
					write: false,
					metafile: true,
					outfile: '.build/office-workbench-check.js',
					loader: { '.css': 'empty', '.svg': 'dataurl', '.png': 'dataurl', '.ttf': 'dataurl', '.wasm': 'binary', '.sh': 'text', '.ps1': 'text', '.scm': 'text' },
					external: ['node:*'],
				});
				const inputs = new Set(Object.keys(result.metafile.inputs).map(file => file.replaceAll('\\', '/')));
				assert.deepStrictEqual(excludedContributions.filter(file => inputs.has(file)), []);
				const required = [
					'customEditor/browser/customEditor.contribution.ts',
					'webview/browser/webview.contribution.ts',
					'chat/browser/chat.shared.contribution.ts',
					'mcp/browser/mcp.contribution.ts',
					'multiDiffEditor/browser/multiDiffEditor.contribution.ts',
					'comments/browser/comments.contribution.ts',
					'timeline/browser/timeline.service.contribution.ts',
					'localHistory/browser/localHistory.contribution.ts',
					'scm/browser/scm.office.contribution.ts',
					'git/browser/git.office.contribution.ts',
					'chat/browser/chatEditing/chatEditingOfficeDocumentEntry.ts',
					'notebook/browser/notebook.service.contribution.ts',
					'remote/common/remote.service.contribution.ts',
					'tasks/common/tasks.office.contribution.ts',
				].map(file => `src/vs/workbench/contrib/${file}`);
				if (application === 'workbench') {
					required.push(...[
						'timeline/browser/timeline.contribution.ts',
						'markers/browser/markers.contribution.ts',
					].map(file => `src/vs/workbench/contrib/${file}`));
				}
				required.push('src/vs/platform/markers/common/markerService.ts');
				assert.deepStrictEqual(required.filter(file => !inputs.has(file)), []);
				const excludedGit = [
					'src/vs/workbench/contrib/chat/browser/pluginGitCommandService.ts',
					'src/vs/workbench/contrib/chat/electron-browser/pluginGitCommandService.ts',
					'src/vs/sessions/contrib/changes/browser/changes.contribution.ts',
					'src/vs/sessions/contrib/changes/browser/sessionSyncChanges.ts',
					'src/vs/sessions/contrib/codeReview/browser/codeReview.contributions.ts',
					'src/vs/sessions/contrib/github/browser/github.contribution.ts',
					'src/vs/sessions/contrib/applyCommitsToParentRepo/browser/applyChangesToParentRepo.ts',
				];
				assert.deepStrictEqual(excludedGit.filter(file => inputs.has(file)), []);
			});
		}
	}
});
