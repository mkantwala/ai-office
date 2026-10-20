/*---------------------------------------------------------------------------------------------
 *  Copyright (c) Microsoft Corporation. All rights reserved.
 *  Licensed under the MIT License. See License.txt in the project root for license information.
 *--------------------------------------------------------------------------------------------*/

import assert from 'assert';
import { DeferredPromise } from '../../../../../base/common/async.js';
import { URI } from '../../../../../base/common/uri.js';
import { ensureNoDisposablesAreLeakedInTestSuite } from '../../../../../base/test/common/utils.js';
import { TestDialogService } from '../../../../../platform/dialogs/test/common/testDialogService.js';
import { NotificationMessage } from '../../../../../platform/notification/common/notification.js';
import { TestNotificationService } from '../../../../../platform/notification/test/common/testNotificationService.js';
import { UndoRedoElementType } from '../../../../../platform/undoRedo/common/undoRedo.js';
import { UndoRedoService } from '../../../../../platform/undoRedo/common/undoRedoService.js';
import { runCustomEditorUndoRedo } from '../../common/customEditorUndoRedo.js';

suite('Custom editor undo/redo ordering', () => {
	ensureNoDisposablesAreLeakedInTestSuite();

	test('rapid undo and redo finish in order without lock warnings or lost operations', async () => {
		const warnings: NotificationMessage[] = [];
		const notifications = new class extends TestNotificationService {
			override warn(message: NotificationMessage) { warnings.push(message); return super.warn(message); }
		};
		const service = new UndoRedoService(new TestDialogService(), notifications);
		const resource = URI.file('/document.docx');
		const started = new DeferredPromise<void>();
		const finish = new DeferredPromise<void>();
		const actions: string[] = [];
		for (const label of ['format.bold', 'text.insert']) {
			service.pushElement({
				type: UndoRedoElementType.Resource, resource, label, code: label,
				undo: async () => {
					actions.push(`undo:${label}`);
					if (label === 'text.insert') { started.complete(); await finish.p; }
				},
				redo: async () => { actions.push(`redo:${label}`); },
			});
		}
		const first = runCustomEditorUndoRedo(service, resource, 'undo');
		await started.p;
		const second = runCustomEditorUndoRedo(service, resource, 'undo');
		const third = runCustomEditorUndoRedo(service, resource, 'redo');
		finish.complete();
		const results = await Promise.all([first, second, third]);
		assert.deepStrictEqual({ results, actions, warnings }, {
			results: [true, true, true],
			actions: ['undo:text.insert', 'undo:format.bold', 'redo:format.bold'], warnings: [],
		});
	});

	test('one busy document does not block other documents and empty history returns false', async () => {
		const service = new UndoRedoService(new TestDialogService(), new TestNotificationService());
		const first = URI.file('/first.docx');
		const second = URI.file('/second.docx');
		const started = new DeferredPromise<void>();
		const finish = new DeferredPromise<void>();
		service.pushElement({ type: UndoRedoElementType.Resource, resource: first, label: 'First', code: 'edit', undo: async () => { started.complete(); await finish.p; }, redo: () => {} });
		service.pushElement({ type: UndoRedoElementType.Resource, resource: second, label: 'Second', code: 'edit', undo: () => {}, redo: () => {} });
		const pending = runCustomEditorUndoRedo(service, first, 'undo');
		await started.p;
		try {
			assert.deepStrictEqual(await Promise.all([
				runCustomEditorUndoRedo(service, second, 'undo'),
				runCustomEditorUndoRedo(service, second, 'undo'),
			]), [true, false]);
		} finally { finish.complete(); await pending; }
	});
});
