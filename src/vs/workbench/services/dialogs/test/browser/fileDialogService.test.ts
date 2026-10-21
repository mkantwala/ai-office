/*---------------------------------------------------------------------------------------------
 *  Copyright (c) Microsoft Corporation. All rights reserved.
 *  Licensed under the MIT License. See License.txt in the project root for license information.
 *--------------------------------------------------------------------------------------------*/

import assert from 'assert';
import { URI } from '../../../../../base/common/uri.js';
import { ensureNoDisposablesAreLeakedInTestSuite } from '../../../../../base/test/common/utils.js';
import { workbenchInstantiationService } from '../../../../test/browser/workbenchTestServices.js';
import { FileDialogService } from '../../browser/fileDialogService.js';

class DocumentFileDialogService extends FileDialogService {
	public saveOptions(uri: URI) { return this.getPickFileToSaveDialogOptions(uri); }
}

suite('Document File Dialog', () => {
	const disposables = ensureNoDisposablesAreLeakedInTestSuite();
	test('Document Save As offers the native format instead of programming languages', () => {
		const instantiation = workbenchInstantiationService(undefined, disposables);
		const service = instantiation.createInstance(DocumentFileDialogService);
		for (const [extension, name] of [['docx', 'Word Document'], ['pdf', 'PDF Document'], ['pptx', 'PowerPoint Presentation'], ['xlsx', 'Excel Workbook'], ['md', 'Markdown Document']]) {
			assert.deepStrictEqual(service.saveOptions(URI.file('/memo.' + extension.toUpperCase())).filters, [{ name, extensions: [extension] }]);
		}
	});
});
