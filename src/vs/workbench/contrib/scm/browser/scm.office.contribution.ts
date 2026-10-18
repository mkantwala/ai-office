/*---------------------------------------------------------------------------------------------
 *  Copyright (c) Microsoft Corporation. All rights reserved.
 *  Licensed under the MIT License. See License.txt in the project root for license information.
 *--------------------------------------------------------------------------------------------*/

import { Event } from '../../../../base/common/event.js';
import { constObservable } from '../../../../base/common/observable.js';
import { localize } from '../../../../nls.js';
import { InstantiationType, registerSingleton } from '../../../../platform/instantiation/common/extensions.js';
import { ISCMRepository, ISCMRepositorySelectionMode, ISCMService, ISCMViewService } from '../common/scm.js';
import { IQuickDiffService } from '../common/quickDiff.js';
import { QuickDiffService } from '../common/quickDiffService.js';
import { IQuickDiffModelService, QuickDiffModelService } from './quickDiffModel.js';

/** Compatibility for shared extension/chat APIs, without providers, repository discovery or SCM UI. */
class OfficeSCMService implements ISCMService {
	declare readonly _serviceBrand: undefined;
	readonly onDidAddRepository = Event.None;
	readonly onDidRemoveRepository = Event.None;
	readonly repositories: ISCMRepository[] = [];
	readonly repositoryCount = 0;
	registerSCMProvider(): never { throw new Error(localize('scm.unavailable', "Source Control is not available in AI Office.")); }
	getRepository(): undefined { return undefined; }
}

class OfficeSCMViewService implements ISCMViewService {
	declare readonly _serviceBrand: undefined;
	readonly menus = { getRepositoryMenus(): never { throw new Error(localize('scm.unavailable', "Source Control is not available in AI Office.")); } };
	readonly selectionModeConfig = constObservable(ISCMRepositorySelectionMode.Single);
	readonly explorerEnabledConfig = constObservable(false);
	readonly graphShowIncomingChangesConfig = constObservable(false);
	readonly graphShowOutgoingChangesConfig = constObservable(false);
	readonly repositories: ISCMRepository[] = [];
	readonly onDidChangeRepositories = Event.None;
	readonly didFinishLoadingRepositories = constObservable(true);
	readonly visibleRepositories = [];
	readonly onDidChangeVisibleRepositories = Event.None;
	readonly focusedRepository = undefined;
	readonly onDidFocusRepository = Event.None;
	readonly activeRepository = constObservable(undefined);
	isVisible(): boolean { return false; }
	toggleVisibility(): void { }
	toggleSortKey(): void { }
	toggleSelectionMode(): void { }
	focus(): void { }
	pinActiveRepository(): void { }
}

registerSingleton(ISCMService, OfficeSCMService, InstantiationType.Delayed);
registerSingleton(ISCMViewService, OfficeSCMViewService, InstantiationType.Delayed);
// Generic saved-document comparisons still serve local history and chat. No SCM decorators or commands.
registerSingleton(IQuickDiffService, QuickDiffService, InstantiationType.Delayed);
registerSingleton(IQuickDiffModelService, QuickDiffModelService, InstantiationType.Delayed);
