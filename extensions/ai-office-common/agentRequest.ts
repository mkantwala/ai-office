/*---------------------------------------------------------------------------------------------
 *  Copyright (c) Microsoft Corporation. All rights reserved.
 *  Licensed under the MIT License. See License.txt in the project root for license information.
 *--------------------------------------------------------------------------------------------*/

import { AsyncLocalStorage } from 'node:async_hooks';

const requests = new AsyncLocalStorage<{ id: string; label: string }>();
export function activeAgentLabel() { return requests.getStore()?.label || 'Agent'; }
export function activeAgentRequest() { return requests.getStore()?.id; }
export function withAgentRequest<T>(id: string, run: () => Promise<T>, label = 'Agent'): Promise<T> {
	return requests.run({ id, label }, run);
}

/** This built-in fork supplies the verified request identity in its opaque tool token. Never take it from model arguments. */
export function withAgentToolRequest<T>(token: unknown, run: () => Promise<T>): Promise<T> {
	const id = (token as { requestId?: string } | undefined)?.requestId;
	if (!id || id === activeAgentRequest()) { return run(); }
	const scheme = (token as { sessionResource?: { scheme?: string } })?.sessionResource?.scheme || '';
	const label = /opencode/i.test(scheme) ? 'OpenCode' : /claude/i.test(scheme) ? 'Claude' : /codex/i.test(scheme) ? 'Codex' : 'Agent';
	return withAgentRequest(id, run, label);
}

export class DocumentHistory {
	cursor = 0;
	end = 0;
	private readonly steps = new Map<number, number>();
	private readonly requestEnds = new Map<string, number>();
	readonly checkpoints = new Map<string, number>();
	private forgetCheckpoint(id: string) { this.checkpoints.delete(id); this.requestEnds.delete(id); }
	checkpoint(id: string) {
		if (this.checkpoints.has(id)) { return; }
		this.checkpoints.set(id, this.cursor);
		if (this.checkpoints.size > 30) { this.forgetCheckpoint(this.checkpoints.keys().next().value!); }
	}
	requestEnd(id: string) { return this.requestEnds.get(id); }
	record(steps = 1, request?: string) {
		for (const [id, position] of this.checkpoints) {
			if (position > this.cursor) { this.forgetCheckpoint(id); }
			else if ((this.requestEnds.get(id) ?? 0) > this.cursor) { this.requestEnds.set(id, this.cursor); }
		}
		for (const position of this.steps.keys()) { if (position > this.cursor) { this.steps.delete(position); } }
		this.end = ++this.cursor;
		this.steps.set(this.cursor, steps);
		if (request) { this.requestEnds.set(request, this.cursor); }
		const oldest = this.cursor - 2048;
		for (const position of this.steps.keys()) { if (position <= oldest) { this.steps.delete(position); } }
		for (const [id, position] of this.checkpoints) { if (position < oldest) { this.forgetCheckpoint(id); } }
	}
	nativeStepsTo(position: number) {
		let steps = 0;
		for (let index = position + 1; index <= this.cursor; index++) { steps += this.steps.get(index) ?? 1; }
		return steps;
	}
	clear() { this.cursor = this.end = 0; this.checkpoints.clear(); this.requestEnds.clear(); this.steps.clear(); }
}
