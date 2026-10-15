/*---------------------------------------------------------------------------------------------
 *  Copyright (c) Microsoft Corporation. All rights reserved.
 *  Licensed under the MIT License. See License.txt in the project root for license information.
 *--------------------------------------------------------------------------------------------*/

import assert from 'assert';
import { TreeVisibility } from '../../../../../base/browser/ui/tree/tree.js';
import { Event } from '../../../../../base/common/event.js';
import { joinPath } from '../../../../../base/common/resources.js';
import { mock } from '../../../../../base/test/common/mock.js';
import { ensureNoDisposablesAreLeakedInTestSuite } from '../../../../../base/test/common/utils.js';
import { TestConfigurationService } from '../../../../../platform/configuration/test/common/testConfigurationService.js';
import { TestWorkspace } from '../../../../../platform/workspace/test/common/testWorkspace.js';
import { EditorInput } from '../../../../common/editor/editorInput.js';
import { IEditorService } from '../../../../services/editor/common/editorService.js';
import { workbenchInstantiationService } from '../../../../test/browser/workbenchTestServices.js';
import { NullFilesConfigurationService, TestFileService } from '../../../../test/common/workbenchTestServices.js';
import { IExplorerService } from '../../browser/files.js';
import { FilesFilter } from '../../browser/views/explorerViewer.js';
import { ExplorerItem, NewExplorerItem } from '../../common/explorerModel.js';

suite('Files - AI Office Explorer filter', () => {
	const store = ensureNoDisposablesAreLeakedInTestSuite();
	let filter: FilesFilter;
	let root: ExplorerItem;
	let fileService: TestFileService;
	let configurationService: TestConfigurationService;
	let editable: ExplorerItem | undefined;
	let visibleEditors: EditorInput[];

	setup(() => {
		editable = undefined;
		visibleEditors = [];
		fileService = store.add(new TestFileService());
		configurationService = new TestConfigurationService({
			files: { exclude: { '**/private': true, '**/hidden.docx': true } },
			explorer: { excludeGitIgnore: false }
		});
		root = new ExplorerItem(TestWorkspace.folders[0].uri, fileService, configurationService, NullFilesConfigurationService, undefined, true);
		const instantiationService = workbenchInstantiationService({
			fileService: () => fileService,
			configurationService: () => configurationService,
		}, store);
		instantiationService.stub(IExplorerService, {
			getEditableData: item => item === editable ? { validationMessage: () => null, onFinish: async () => { } } : undefined,
			findClosestRoot: () => root,
		});
		instantiationService.stub(IEditorService, { visibleEditors, onDidVisibleEditorsChange: Event.None });
		filter = store.add(instantiationService.createInstance(FilesFilter));
	});

	function item(name: string, isDirectory = false, parent = root): ExplorerItem {
		const child = new ExplorerItem(joinPath(parent.resource, name), fileService, configurationService, NullFilesConfigurationService, parent, isDirectory);
		parent.addChild(child);
		return child;
	}

	test('shows only office documents and Markdown, including mixed-case extensions', () => {
		const names = ['Memo.docx', 'Budget.XLSX', 'Contract.PdF', 'style.md', 'Notes.MARKDOWN', 'Deck.PpTx', 'code.ts', 'readme.txt', 'image.png', 'archive.docx.zip', 'LICENSE', '.gitignore'];
		assert.deepStrictEqual(names.filter(name => filter.filter(item(name), TreeVisibility.Visible)), names.slice(0, 6));
	});

	test('keeps folders and nested documents navigable', () => {
		const folder = item('Reports', true);
		const empty = item('Empty', true);
		const dotted = item('archive.zip', true);
		const document = item('Board.docx', false, folder);
		assert.deepStrictEqual([root, folder, empty, dotted, document].map(stat => filter.filter(stat, TreeVisibility.Visible)), [true, true, true, true, true]);
	});

	test('keeps exclusions and hidden parents effective', () => {
		const folder = item('private', true);
		assert.deepStrictEqual([
			filter.filter(folder, TreeVisibility.Visible),
			filter.filter(item('Memo.docx', false, folder), TreeVisibility.Hidden),
			filter.filter(item('hidden.docx'), TreeVisibility.Visible),
		], [false, false, false]);
	});

	test('keeps the workspace error placeholder visible', () => {
		const placeholder = new ExplorerItem(root.resource, fileService, configurationService, NullFilesConfigurationService, undefined);
		placeholder.error = new Error('Workspace unavailable');
		assert.strictEqual(filter.filter(placeholder, TreeVisibility.Visible), true);
	});

	test('does not reveal unsupported files when their editor is open', () => {
		const code = item('code.ts');
		const document = item('hidden.docx');
		for (const stat of [code, document]) {
			visibleEditors.push(new class extends mock<EditorInput>() {
				override get resource() { return stat.resource; }
			});
		}
		assert.deepStrictEqual([code, document].map(stat => filter.filter(stat, TreeVisibility.Visible)), [false, true]);
	});

	test('keeps a new filename input visible until editing ends', () => {
		const draft = new NewExplorerItem(fileService, configurationService, NullFilesConfigurationService, root, false);
		editable = draft;
		const duringEdit = filter.filter(draft, TreeVisibility.Visible);
		editable = undefined;
		assert.deepStrictEqual([duringEdit, filter.filter(draft, TreeVisibility.Visible)], [true, false]);
	});

	test('applies the same file types to Explorer Find in unloaded folders', () => {
		const names = ['nested/Memo.DOCX', 'nested/Budget.xlsx', 'nested/Contract.pdf', 'nested/notes.md', 'nested/notes.markdown', 'nested/main.js', 'nested/Deck.pptx'];
		assert.deepStrictEqual(names.filter(name => !filter.isIgnored(joinPath(root.resource, name), root.resource, false)), [...names.slice(0, 5), 'nested/Deck.pptx']);
		assert.strictEqual(filter.isIgnored(joinPath(root.resource, 'folder.txt'), root.resource, true), false);
	});
});
