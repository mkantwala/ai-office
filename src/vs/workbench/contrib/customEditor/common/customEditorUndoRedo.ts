/*---------------------------------------------------------------------------------------------
 *  Copyright (c) Microsoft Corporation. All rights reserved.
 *  Licensed under the MIT License. See License.txt in the project root for license information.
 *--------------------------------------------------------------------------------------------*/

import { SequencerByKey } from '../../../../base/common/async.js';
import { URI } from '../../../../base/common/uri.js';
import { IUndoRedoService } from '../../../../platform/undoRedo/common/undoRedo.js';

const queues = new WeakMap<IUndoRedoService, SequencerByKey<string>>();

/** Share ordering between editor keybindings, document tools and toolbar history actions. */
export function runCustomEditorUndoRedo(service: IUndoRedoService, resource: URI, action: 'undo' | 'redo'): Promise<boolean> {
	let queue = queues.get(service);
	if (!queue) { queue = new SequencerByKey<string>(); queues.set(service, queue); }
	return queue.queue(service.getUriComparisonKey(resource), async () => {
		if (action === 'undo' ? !service.canUndo(resource) : !service.canRedo(resource)) { return false; }
		await service[action](resource);
		return true;
	});
}
