/*---------------------------------------------------------------------------------------------
 *  Copyright (c) Microsoft Corporation. All rights reserved.
 *  Licensed under the MIT License. See License.txt in the project root for license information.
 *--------------------------------------------------------------------------------------------*/

import { Disposable } from '../../../../base/common/lifecycle.js';
import { InstantiationType, registerSingleton } from '../../../../platform/instantiation/common/extensions.js';
import { IGitService, IGitRepository } from '../common/gitService.js';

/** No extension activation, repository discovery, delegate wait or Git commands. */
class OfficeGitService implements IGitService {
	declare readonly _serviceBrand: undefined;
	readonly repositories: IGitRepository[] = [];
	setDelegate() { return Disposable.None; }
	async openRepository(): Promise<undefined> { return undefined; }
}

registerSingleton(IGitService, OfficeGitService, InstantiationType.Delayed);
