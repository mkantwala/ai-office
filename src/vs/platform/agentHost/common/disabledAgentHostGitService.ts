/*---------------------------------------------------------------------------------------------
 *  Copyright (c) Microsoft Corporation. All rights reserved.
 *  Licensed under the MIT License. See License.txt in the project root for license information.
 *--------------------------------------------------------------------------------------------*/

import { constObservable } from '../../../base/common/observable.js';
import { localize } from '../../../nls.js';
import { IAgentHostGitService } from './agentHostGitService.js';

/** Agent sessions use folders and document snapshots, without repository probes or Git processes. */
export class DisabledAgentHostGitService implements IAgentHostGitService {
	declare readonly _serviceBrand: undefined;
	private async unavailable(): Promise<never> { throw new Error(localize('git.unavailable', "Git is not available in AI Office.")); }
	async getCurrentBranch(): Promise<undefined> { return undefined; }
	async getCurrentBranchName(): Promise<undefined> { return undefined; }
	async getDefaultBranch(): Promise<undefined> { return undefined; }
	async getRefs(): Promise<never[]> { return []; }
	async getBranches(): Promise<never[]> { return []; }
	async getBranch(): Promise<undefined> { return undefined; }
	hasGitRoot() { return constObservable(false); }
	async getRepositoryRoot(): Promise<undefined> { return undefined; }
	async getWorktreeRoots(): Promise<never[]> { return []; }
	addWorktree() { return this.unavailable(); }
	copyWorktreeIncludeFiles() { return this.unavailable(); }
	symlinkWorktreeFolders() { return this.unavailable(); }
	addExistingWorktree() { return this.unavailable(); }
	removeWorktree() { return this.unavailable(); }
	async branchExists(): Promise<boolean> { return false; }
	createBranch() { return this.unavailable(); }
	checkout() { return this.unavailable(); }
	async hasUncommittedChanges(): Promise<boolean> { return false; }
	createStash() { return this.unavailable(); }
	commitAll() { return this.unavailable(); }
	mergeBranch() { return this.unavailable(); }
	restore() { return this.unavailable(); }
	async hasUpstream(): Promise<boolean> { return false; }
	fetch() { return this.unavailable(); }
	pull() { return this.unavailable(); }
	push() { return this.unavailable(); }
	async getSessionGitState(): Promise<undefined> { return undefined; }
	async getFetchRemoteUrls(): Promise<undefined> { return undefined; }
	async getFetchRemotes(): Promise<undefined> { return undefined; }
	async getUntrackedPaths(): Promise<undefined> { return undefined; }
	async computeSessionFileDiffs(): Promise<undefined> { return undefined; }
	async resolveBranchBaselineCommit(): Promise<undefined> { return undefined; }
	async showBlob(): Promise<undefined> { return undefined; }
	async captureWorkingTreeAsTree(): Promise<undefined> { return undefined; }
	commitTree() { return this.unavailable(); }
	updateRef() { return this.unavailable(); }
	deleteRefs() { return this.unavailable(); }
	async revParse(): Promise<undefined> { return undefined; }
	async listRefNamesWithOids(): Promise<never[]> { return []; }
	overlayPathIntoTree() { return this.unavailable(); }
	async diffTreePaths(): Promise<undefined> { return undefined; }
	async computeFileDiffsBetweenRefs(): Promise<undefined> { return undefined; }
	async getBranchDiffSafetyInfo(): Promise<undefined> { return undefined; }
	async getDiffPatchBetweenRefs(): Promise<undefined> { return undefined; }
}
