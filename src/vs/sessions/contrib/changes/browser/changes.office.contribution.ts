/*---------------------------------------------------------------------------------------------
 *  Copyright (c) Microsoft Corporation. All rights reserved.
 *  Licensed under the MIT License. See License.txt in the project root for license information.
 *--------------------------------------------------------------------------------------------*/

import { InstantiationType, registerSingleton } from '../../../../platform/instantiation/common/extensions.js';
import { IChangesViewService } from '../common/changesViewService.js';
import { ChangesViewService } from './changesViewService.js';
import { ISessionChangesService, SessionChangesService } from './sessionChangesService.js';
import { ICodeReviewService, CodeReviewService } from '../../codeReview/browser/codeReviewService.js';
import { IGitHubService, GitHubService } from '../../github/browser/githubService.js';
import { IPullRequestIconCache, PullRequestIconCache } from '../../github/browser/pullRequestIconCache.js';

// Shared session consumers still require these services. Repository UI, commit,
// sync, pull-request polling and Git-based code-review actions are not registered.
registerSingleton(IChangesViewService, ChangesViewService, InstantiationType.Delayed);
registerSingleton(ISessionChangesService, SessionChangesService, InstantiationType.Delayed);
registerSingleton(ICodeReviewService, CodeReviewService, InstantiationType.Delayed);
registerSingleton(IGitHubService, GitHubService, InstantiationType.Delayed);
registerSingleton(IPullRequestIconCache, PullRequestIconCache, InstantiationType.Delayed);
