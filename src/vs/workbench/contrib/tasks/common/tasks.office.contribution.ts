/*---------------------------------------------------------------------------------------------
 *  Copyright (c) Microsoft Corporation. All rights reserved.
 *  Licensed under the MIT License. See License.txt in the project root for license information.
 *--------------------------------------------------------------------------------------------*/

import { Action } from '../../../../base/common/actions.js';
import { Event } from '../../../../base/common/event.js';
import { Disposable } from '../../../../base/common/lifecycle.js';
import { localize } from '../../../../nls.js';
import { InstantiationType, registerSingleton } from '../../../../platform/instantiation/common/extensions.js';
import { ITaskService, IWorkspaceFolderTaskResult } from './taskService.js';
import { Task, TaskSorter } from './tasks.js';

/** Keep the extension-host API injectable without task discovery, execution, commands or watchers. */
class OfficeTaskService implements ITaskService {
	declare readonly _serviceBrand: undefined;
	readonly onDidStateChange = Event.None;
	readonly onDidChangeTaskProviders = Event.None;
	readonly onDidReconnectToTasks = Event.None;
	readonly onDidChangeTaskSystemInfo = Event.None;
	readonly onDidChangeTaskConfig = Event.None;
	readonly isReconnected = true;
	readonly supportsMultipleTaskExecutions = false;
	readonly hasTaskSystemInfo = false;

	configureAction(): Action { return new Action('workbench.action.tasks.configureTaskRunner', localize('tasks.unavailable', "Tasks are not available in Sangam One."), undefined, false); }
	async run(): Promise<never> { throw new Error(localize('tasks.unavailable', "Tasks are not available in Sangam One.")); }
	async terminate(): Promise<never> { throw new Error(localize('tasks.unavailable', "Tasks are not available in Sangam One.")); }
	inTerminal(): boolean { return false; }
	async getActiveTasks(): Promise<Task[]> { return []; }
	async getBusyTasks(): Promise<Task[]> { return []; }
	async tasks(): Promise<Task[]> { return []; }
	async getKnownTasks(): Promise<Task[]> { return []; }
	async getSavedTasks(): Promise<Task[]> { return []; }
	taskTypes(): string[] { return []; }
	async getWorkspaceTasks(): Promise<Map<string, IWorkspaceFolderTaskResult>> { return new Map(); }
	async getTask(): Promise<undefined> { return undefined; }
	async tryResolveTask(): Promise<undefined> { return undefined; }
	createSorter(): TaskSorter { return new TaskSorter([]); }
	getTaskDescription(): undefined { return undefined; }
	getTerminalsForTasks(): undefined { return undefined; }
	getTaskProblems(): undefined { return undefined; }
	async openConfig(): Promise<boolean> { return false; }
	async customize(): Promise<void> { }
	async extensionCallbackTaskComplete(): Promise<void> { }
	registerTaskProvider() { return Disposable.None; }
	registerTaskSystem(): void { }
	registerSupportedExecutions(): void { }
	removeRecentlyUsedTask(): void { }
	rerun(): void { }
}

registerSingleton(ITaskService, OfficeTaskService, InstantiationType.Delayed);
