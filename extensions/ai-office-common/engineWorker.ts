/*---------------------------------------------------------------------------------------------
 *  Copyright (c) Microsoft Corporation. All rights reserved.
 *  Licensed under the MIT License. See License.txt in the project root for license information.
 *--------------------------------------------------------------------------------------------*/

import { parentPort } from 'node:worker_threads';
import { openEngine } from './model.js';
import type { Engine, Kind } from './types.js';
import { WordFrame } from './wordFrame.js';
let engine: Engine | undefined;
let wordFrame: WordFrame | undefined;
let queue = Promise.resolve();
parentPort?.on('message', message => {
	queue = queue.then(async () => {
		try {
			let result: unknown;
			let transfer: ArrayBuffer[] = [];
			if (message.method === 'open') {
				const opened = await openEngine(message.kind as Kind, message.name, message.bytes, message.extensionsRoot, message.password);
				engine?.dispose(); engine = opened;
				wordFrame = message.kind === 'word' ? new WordFrame() : undefined;
				result = { pages: engine.pages(), state: engine.state(), commands: engine.commands() };
			} else {
				if (!engine) { throw new Error('Document is not open.'); }
				switch (message.method) {
					case 'execute': result = { state: await engine.execute(message.command, message.args), pages: engine.pages() }; break;
					case 'export': result = await engine.export(); break;
					case 'render': {
						const raster = engine.render(message.page, message.scale);
						if (wordFrame && raster.pixels && raster.width && raster.height && typeof message.frame === 'number') {
							const update = wordFrame.update({ pixels: raster.pixels, width: raster.width, height: raster.height }, message.page, message.scale, message.frame, message.base);
							result = update; transfer = [update.pixels.buffer as ArrayBuffer];
						} else { result = raster; }
						break;
					}
					default: throw new Error('Unknown engine request.');
				}
			}
			parentPort?.postMessage({ id: message.id, result }, transfer);
		} catch (error) { parentPort?.postMessage({ id: message.id, error: String(error) }); }
	});
});
