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
					case 'reviewRead': result = engine.agent?.reviewRead?.() || { changes: [] }; break;
					case 'reviewExport': result = engine.agent?.reviewExport?.() || ''; break;
					case 'reviewImport': engine.agent?.reviewImport?.(message.value); result = true; break;
					case 'reviewProjection': result = engine.agent?.reviewProjection?.() || ''; break;
					case 'reviewLayout': result = message.changes.map((change: { id: string; anchor?: { id: string } }) => {
						const rects = (id: string) => { try { return engine?.agent?.rects(id) || []; } catch { return []; } };
						return { ...change, rects: change.id === 'document' ? [] : rects(change.id.slice(change.id.indexOf(':') + 1)), anchorRects: change.anchor ? rects(change.anchor.id) : [] };
					}); break;
					case 'reviewDecide':
						if (!engine.agent?.reviewDecide) { throw new Error('Word change review is unavailable.'); }
						result = { state: engine.agent.reviewDecide(message.revision, message.action, message.ids), pages: engine.pages() }; break;
					case 'agentRead':
						if (!engine.agent) { throw new Error('Rebuild the Word engine to enable document tools.'); }
						result = engine.agent.read(message.args); break;
					case 'agentRects':
						if (!engine.agent) { throw new Error('Document tools are unavailable.'); }
						result = engine.agent.rects(message.blockId); break;
					case 'agentEdit':
						if (!engine.agent) { throw new Error('Document tools are unavailable.'); }
						if (message.review) { engine.agent.reviewStart?.(); }
						result = { state: await engine.agent.edit(message.args), pages: engine.pages() };
						if (message.review) { engine.agent.reviewFinish?.(); } break;
					case 'execute':
						if (message.review) { engine.agent?.reviewStart?.(); }
						result = { state: await engine.execute(message.command, message.args), pages: engine.pages() };
						if (message.review) { engine.agent?.reviewFinish?.(); } break;
					case 'tablePreview': result = engine.tablePreview?.(message.args); break;
					case 'export': result = await engine.export(); break;
					case 'render': {
						const raster = engine.render(message.page, message.scale);
						if (wordFrame && raster.pixels && raster.width && raster.height && typeof message.frame === 'number') {
							const update = wordFrame.update({ pixels: raster.pixels, width: raster.width, height: raster.height }, message.page, message.scale, message.frame, message.base);
							result = { ...update, tables: raster.tables }; transfer = [update.pixels.buffer as ArrayBuffer];
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
