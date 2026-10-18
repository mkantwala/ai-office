/*---------------------------------------------------------------------------------------------
 *  Copyright (c) Microsoft Corporation. All rights reserved.
 *  Licensed under the MIT License. See License.txt in the project root for license information.
 *--------------------------------------------------------------------------------------------*/

import { Worker } from 'node:worker_threads';
export class EngineClient {
	private readonly worker: Worker;
	private nextId = 0;
	private stopped?: Error;
	private readonly pending = new Map<number, { resolve(value: unknown): void; reject(error: Error): void; timer: ReturnType<typeof setTimeout> }>();
	constructor(path: string) {
		this.worker = new Worker(path);
		this.worker.on('message', message => {
			const item = this.pending.get(message.id); if (!item) { return; }
			this.pending.delete(message.id); clearTimeout(item.timer);
			if (message.error) { item.reject(new Error(message.error)); } else { item.resolve(message.result); }
		});
		this.worker.on('error', error => this.fail(error));
		this.worker.on('exit', code => this.fail(new Error('Document engine stopped (' + code + '). Reopen the tab.')));
	}
	request<T>(method: string, args: Record<string, unknown> = {}): Promise<T> {
		if (this.stopped) { return Promise.reject(this.stopped); }
		return new Promise((resolve, reject) => {
			const id = ++this.nextId;
			const timer = setTimeout(() => { this.fail(new Error('Document engine timed out. Reopen the tab.')); void this.worker.terminate(); }, 120000);
			this.pending.set(id, { resolve: value => resolve(value as T), reject, timer });
			try { this.worker.postMessage({ id, method, ...args }); }
			catch (error) {
				clearTimeout(timer); this.pending.delete(id);
				reject(error instanceof Error ? error : new Error('Document request could not be sent.'));
			}
		});
	}
	private fail(error: Error) { this.stopped = error; for (const item of this.pending.values()) { clearTimeout(item.timer); item.reject(error); } this.pending.clear(); }
	dispose() { this.fail(new Error('Document closed.')); void this.worker.terminate(); }
}
