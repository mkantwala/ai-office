/*---------------------------------------------------------------------------------------------
 *  Copyright (c) Microsoft Corporation. All rights reserved.
 *  Licensed under the MIT License. See License.txt in the project root for license information.
 *--------------------------------------------------------------------------------------------*/

import { localize } from '../../../../nls.js';
import { InstantiationType, registerSingleton } from '../../../../platform/instantiation/common/extensions.js';
import { IWorkbenchIssueService } from './issue.js';

/** Keep extension management and Settings Sync injectable without loading the issue reporter. */
class UnsupportedIssueService implements IWorkbenchIssueService {
	declare readonly _serviceBrand: undefined;

	async openReporter(): Promise<void> {
		throw new Error(localize('issueReporter.unavailable', "The issue reporter is not available in AI Office."));
	}
}

registerSingleton(IWorkbenchIssueService, UnsupportedIssueService, InstantiationType.Delayed);
