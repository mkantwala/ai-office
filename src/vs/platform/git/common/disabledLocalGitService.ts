/*---------------------------------------------------------------------------------------------
 *  Copyright (c) Microsoft Corporation. All rights reserved.
 *  Licensed under the MIT License. See License.txt in the project root for license information.
 *--------------------------------------------------------------------------------------------*/

import { localize } from '../../../nls.js';
import { ILocalGitService } from './localGitService.js';

/** Shared-process compatibility without loading a Git process implementation. */
export class DisabledLocalGitService implements ILocalGitService {
	declare readonly _serviceBrand: undefined;
	private async unavailable(): Promise<never> { throw new Error(localize('git.unavailable', "Git is not available in AI Office.")); }
	clone() { return this.unavailable(); }
	pull() { return this.unavailable(); }
	checkout() { return this.unavailable(); }
	checkoutCommit() { return this.unavailable(); }
	revParse() { return this.unavailable(); }
	getRemoteUrl() { return this.unavailable(); }
	fetch() { return this.unavailable(); }
	revListCount() { return this.unavailable(); }
	async cancel(): Promise<void> { }
}
