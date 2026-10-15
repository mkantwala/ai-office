/*---------------------------------------------------------------------------------------------
 *  Copyright (c) Microsoft Corporation. All rights reserved.
 *  Licensed under the MIT License. See License.txt in the project root for license information.
 *--------------------------------------------------------------------------------------------*/

import { parentPort } from 'node:worker_threads';
import { openEngine } from './model.js';
import type { Engine, Kind } from './types.js';
let engine: Engine | undefined;
let queue = Promise.resolve();
parentPort?.on('message', message => {
	queue = queue.then(async () => {
		try {
			let result: unknown;
			if (message.method === 'open') {
				const opened = await openEngine(message.kind as Kind, message.name, message.bytes, message.extensionsRoot, message.password);
				engine?.dispose(); engine = opened;
				result = { pages: engine.pages(), state: engine.state(), commands: engine.commands() };
			} else {
				if (!engine) { throw new Error('Document is not open.'); }
				switch (message.method) {
					case 'execute': result = { state: await engine.execute(message.command, message.args), pages: engine.pages() }; break;
					case 'export': result = await engine.export(); break;
					case 'render': result = engine.render(message.page, message.scale); break;
					default: throw new Error('Unknown engine request.');
				}
			}
			parentPort?.postMessage({ id: message.id, result });
		} catch (error) { parentPort?.postMessage({ id: message.id, error: String(error) }); }
	});
});
