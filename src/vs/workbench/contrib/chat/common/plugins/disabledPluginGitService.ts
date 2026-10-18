/*---------------------------------------------------------------------------------------------
 *  Copyright (c) Microsoft Corporation. All rights reserved.
 *  Licensed under the MIT License. See License.txt in the project root for license information.
 *--------------------------------------------------------------------------------------------*/

import { localize } from '../../../../../nls.js';
import { IPluginGitService } from './pluginGitService.js';

/** Local agent skills remain available; Git plugin repositories cannot clone, fetch or update. */
export class DisabledPluginGitService implements IPluginGitService {
	declare readonly _serviceBrand: undefined;
	private async unavailable(): Promise<never> { throw new Error(localize('git.unavailable', "Git is not available in Sangam One. Use local agent skills instead.")); }
	cloneRepository() { return this.unavailable(); }
	pull() { return this.unavailable(); }
	checkout() { return this.unavailable(); }
	checkoutCommit() { return this.unavailable(); }
	revParse() { return this.unavailable(); }
	fetch() { return this.unavailable(); }
	fetchRepository() { return this.unavailable(); }
	revListCount() { return this.unavailable(); }
}
