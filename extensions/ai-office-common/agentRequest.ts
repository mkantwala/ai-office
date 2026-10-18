/*---------------------------------------------------------------------------------------------
 *  Copyright (c) Microsoft Corporation. All rights reserved.
 *  Licensed under the MIT License. See License.txt in the project root for license information.
 *--------------------------------------------------------------------------------------------*/

let current: string | undefined;
export function activeAgentRequest() { return current; }
export async function withAgentRequest<T>(id: string, run: () => Promise<T>): Promise<T> {
	if (current) { throw new Error('A document agent request is already running. Wait for it to finish or stop it first.'); }
	current = id;
	try { return await run(); } finally { current = undefined; }
}

export class DocumentHistory {
	cursor = 0;
	end = 0;
	private readonly steps = new Map<number, number>();
	readonly checkpoints = new Map<string, number>();
	checkpoint(id: string) {
		if (this.checkpoints.has(id)) { return; }
		this.checkpoints.set(id, this.cursor);
		if (this.checkpoints.size > 30) { this.checkpoints.delete(this.checkpoints.keys().next().value!); }
	}
	record(steps = 1) {
		for (const [id, position] of this.checkpoints) { if (position > this.cursor) { this.checkpoints.delete(id); } }
		for (const position of this.steps.keys()) { if (position > this.cursor) { this.steps.delete(position); } }
		this.end = ++this.cursor;
		this.steps.set(this.cursor, steps);
		const oldest = this.cursor - 2048;
		for (const position of this.steps.keys()) { if (position <= oldest) { this.steps.delete(position); } }
		for (const [id, position] of this.checkpoints) { if (position < oldest) { this.checkpoints.delete(id); } }
	}
	nativeStepsTo(position: number) {
		let steps = 0;
		for (let index = position + 1; index <= this.cursor; index++) { steps += this.steps.get(index) ?? 1; }
		return steps;
	}
	clear() { this.cursor = this.end = 0; this.checkpoints.clear(); this.steps.clear(); }
}
