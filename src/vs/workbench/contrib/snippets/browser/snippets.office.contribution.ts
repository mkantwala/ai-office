/*---------------------------------------------------------------------------------------------
 *  Copyright (c) Microsoft Corporation. All rights reserved.
 *  Licensed under the MIT License. See License.txt in the project root for license information.
 *--------------------------------------------------------------------------------------------*/

import { InstantiationType, registerSingleton } from '../../../../platform/instantiation/common/extensions.js';
import { ISnippetsService } from './snippets.js';
import type { Snippet, SnippetFile } from './snippetsFile.js';

/** Preserve consumers of the snippets API without loading user snippets or completion providers. */
class OfficeSnippetsService implements ISnippetsService {
	declare readonly _serviceBrand: undefined;

	async getSnippetFiles(): Promise<Iterable<SnippetFile>> { return []; }
	async getSnippets(): Promise<Snippet[]> { return []; }
	getSnippetsSync(): Snippet[] { return []; }
	isEnabled(): boolean { return false; }
	updateEnablement(): void { }
	updateUsageTimestamp(): void { }
}

registerSingleton(ISnippetsService, OfficeSnippetsService, InstantiationType.Delayed);
