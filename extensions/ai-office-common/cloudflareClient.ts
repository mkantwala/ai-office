/*---------------------------------------------------------------------------------------------
 *  Copyright (c) Microsoft Corporation. All rights reserved.
 *  Licensed under the MIT License. See License.txt in the project root for license information.
 *--------------------------------------------------------------------------------------------*/

export const officeModel = '@cf/zai-org/glm-4.7-flash';
export interface CloudflareCredentials { accountId: string; token: string }
export interface OfficeTool { type: 'function'; function: { name: string; description: string; parameters: object } }
export interface OfficeToolCall { id: string; type: 'function'; function: { name: string; arguments: string } }
export interface OfficeMessage {
	role: 'system' | 'user' | 'assistant' | 'tool';
	content: string | null;
	tool_calls?: OfficeToolCall[];
	tool_call_id?: string;
}

/** A single fixed endpoint/model. Never log credentials, document content or upstream error bodies. */
export async function completeOfficeChat(credentials: CloudflareCredentials, messages: OfficeMessage[], tools: OfficeTool[], signal: AbortSignal, request: typeof fetch = fetch, required = false): Promise<OfficeMessage> {
	if (!/^[a-f0-9]{32}$/i.test(credentials.accountId) || !credentials.token.trim()) {
		throw new Error('Configure Cloudflare in AI Office before sending a message.');
	}
	signal.throwIfAborted();
	if (Buffer.byteLength(JSON.stringify({ messages, tools }), 'utf8') > 120000) { throw new Error('This conversation is full. Start a new chat and continue with the relevant document section.'); }
	let response: Response;
	try {
		response = await request(`https://api.cloudflare.com/client/v4/accounts/${credentials.accountId}/ai/v1/chat/completions`, {
			method: 'POST', redirect: 'error', signal: AbortSignal.any([signal, AbortSignal.timeout(90000)]),
			headers: { Authorization: `Bearer ${credentials.token}`, 'Content-Type': 'application/json' },
			body: JSON.stringify({ model: officeModel, messages, stream: false, max_completion_tokens: 4096, temperature: 0.3, chat_template_kwargs: { enable_thinking: false }, ...(tools.length ? { tools, tool_choice: required ? 'required' : 'auto', parallel_tool_calls: false } : {}) })
		});
	} catch {
		signal.throwIfAborted();
		throw new Error('Cloudflare could not be reached or the request timed out. No automatic retry was made.');
	}
	if (!response.ok) {
		await response.body?.cancel();
		throw new Error(response.status === 401 || response.status === 403 ? 'Cloudflare rejected the credentials. Check the account and Workers AI permissions.' : response.status === 429 ? 'Cloudflare rate limit reached. Try again later.' : `Cloudflare request failed (HTTP ${response.status}).`);
	}
	// Bound even a malformed provider response before parsing it.
	const reader = response.body?.getReader();
	if (!reader) { throw new Error('Cloudflare returned an empty response.'); }
	const chunks: Uint8Array[] = []; let size = 0;
	try {
		while (true) {
			const { done, value } = await reader.read();
			if (done) { break; }
			size += value.byteLength;
			if (size > 1000000) { throw new Error('Cloudflare response exceeded the size limit.'); }
			chunks.push(value);
		}
	} finally { await reader.cancel(); reader.releaseLock(); }
	let message: OfficeMessage;
	try {
		const body = JSON.parse(Buffer.concat(chunks).toString('utf8'));
		if (body.choices?.[0]?.finish_reason === 'length') { throw new Error(); }
		message = body.choices?.[0]?.message;
		if (!message || message.role !== 'assistant' || (message.content !== null && message.content !== undefined && typeof message.content !== 'string')) { throw new Error(); }
		if (message.tool_calls !== undefined && (!Array.isArray(message.tool_calls) || message.tool_calls.length > 8 || message.tool_calls.some(call => !call || typeof call.id !== 'string' || call.type !== 'function' || typeof call.function?.name !== 'string' || typeof call.function?.arguments !== 'string'))) { throw new Error(); }
	} catch { throw new Error('Cloudflare returned an incomplete or invalid response. No edit from that response was applied.'); }
	return { role: 'assistant', content: message.content ?? null, ...(message.tool_calls?.length ? { tool_calls: message.tool_calls } : {}) };
}

/** Bounded, sequential tool loop; results are retained only for this request. */
export async function runOfficeAgent(options: {
	messages: OfficeMessage[];
	tools: OfficeTool[];
	signal: AbortSignal;
	complete: (messages: OfficeMessage[], tools: OfficeTool[]) => Promise<OfficeMessage>;
	invoke: (name: string, input: Record<string, unknown>) => Promise<string>;
	progress: (name: string) => void;
}): Promise<string> {
	const messages = [...options.messages];
	const allowed = new Set(options.tools.map(tool => tool.function.name));
	const searches = new Map<string, { uri: string; total: number; ids: Set<string>; pending: boolean }>();
	const applied = new Map<string, { uri: string; label: string; ids: Set<string>; verified?: boolean }>();
	let reviewed = false;
	const repeatedErrors = new Map<string, number>();
	const repeatedReads = new Map<string, number>();
	let lastError: string | undefined;
	const canonical = (value: unknown): string => JSON.stringify(value, (_key, entry) => entry && typeof entry === 'object' && !Array.isArray(entry) ? Object.fromEntries(Object.keys(entry).sort().map(key => [key, entry[key]])) : entry);
	const gaps = () => [...searches.values()].flatMap(search => [...applied.values()].filter(action => action.uri === search.uri && !action.verified).flatMap(action => {
		const done = [...search.ids].filter(id => action.ids.has(id)).length;
		return done > 0 && (done < search.total || search.pending) ? [{ action: action.label, applied: done, found: search.total, moreResults: search.pending }] : [];
	}));

	for (let round = 0; round < 16; round++) {
		options.signal.throwIfAborted();
		const reply = await options.complete(messages, options.tools);
		options.signal.throwIfAborted();
		messages.push(reply);
		if (!reply.tool_calls?.length) {
			const remaining = gaps();
			if (applied.size && !reviewed) {
				reviewed = true;
				messages.push({ role: 'system', content: 'Completion check from successful tool results: ' + JSON.stringify({ actions: [...applied.values()].map(action => ({ action: action.label, targets: action.ids.size })), coverageGaps: remaining }) + '. Compare every clause of the user request against these actual actions and targets. If an intended action or target is missing, finish it and follow pagination before claiming completion. Do not repeat completed edits. If a subset was intended, state that scope accurately. Successful formatting receipts already verify their full selected scope and exact actions. Do not read blocks or history merely to confirm those receipts; only resolve still-missing user requirements.' });
				continue;
			}
			if (remaining.length) { return 'Completed document operations: ' + remaining.map(item => `${item.action}: ${item.applied} of ${item.found} matching blocks`).join('; ') + '. Other matches were not changed for these operations. The broader change is incomplete.'; }
			return reply.content || 'The model returned no answer. Please try a more specific request.';
		}
		for (const call of reply.tool_calls) {
			options.signal.throwIfAborted();
			let result: string;
			try {
				if (!allowed.has(call.function.name)) { throw new Error('This tool is unavailable. Use only the supplied Word tools.'); }
				if (call.function.arguments.length > 160000) { throw new Error('Tool input is too large.'); }
				const input = JSON.parse(call.function.arguments);
				if (!input || typeof input !== 'object' || Array.isArray(input)) { throw new Error('Tool input must be an object.'); }
				options.progress(call.function.name);
				result = await options.invoke(call.function.name, input);
				try {
					const value = JSON.parse(result);
					if (value.error) {
						lastError = String(value.error);
						const count = (repeatedErrors.get(lastError) ?? 0) + 1;
						repeatedErrors.set(lastError, count);
						if (count >= 3) {
							return 'The agent repeated the same invalid operation three times and stopped. ' + (applied.size ? 'Some operations completed; inspect document history for recovery options. ' : 'No successful document edit was recorded. ') + 'Last error: ' + lastError;
						}
					}
					if (!value.error && call.function.name === 'office_document_read' && typeof value.version === 'string') {
						const key = canonical([value.uri, value.version, { ...input, operation: input.operation ?? 'blocks', version: undefined }]);
						const count = (repeatedReads.get(key) ?? 0) + 1;
						repeatedReads.set(key, count);
						if (count > 1) {
							result = JSON.stringify({ ...value, guidance: 'This unchanged document result was already returned. Reuse its IDs and version. For formatting, call office_document_format with query and all format properties; it resolves targets and style sizes itself. Do not reread the same blocks or rewrite unchanged text.' });
						}
					}

					if (!value.error && call.function.name === 'office_document_read' && typeof input.query === 'string' && typeof value.uri === 'string' && Array.isArray(value.blocks) && typeof value.total === 'number') {
						const key = JSON.stringify([value.uri, input.query, input.parentId]);
						const entry = searches.get(key) || { uri: value.uri, total: value.total, ids: new Set<string>(), pending: false };
						for (const block of value.blocks) { if (typeof block.id === 'string') { entry.ids.add(block.id); } }
						entry.total = value.total; entry.pending = entry.ids.size < value.total;
						searches.set(key, entry);
					}
					if (!value.error && ['office_document_edit', 'office_document_format'].includes(call.function.name) && typeof value.uri === 'string' && Array.isArray(value.appliedEdits)) {
						for (const edit of value.appliedEdits) {
							if (typeof edit.id !== 'string') { continue; }
							const key = JSON.stringify([value.uri, edit.operation, edit.command, edit.params]);
							const entry = applied.get(key) || { uri: value.uri, label: edit.command || edit.operation, ids: new Set<string>(), verified: false };
							entry.verified ||= value.verifiedScope === true;
							entry.ids.add(edit.id); applied.set(key, entry);
						}
					}
				} catch { /* Ordinary non-JSON results are still returned to the model. */ }
			} catch (error) {
				options.signal.throwIfAborted();
				result = JSON.stringify({ error: error instanceof Error ? error.message : 'Tool failed.' });
			}
			if (result.length > 24000) {
				// Never turn a completed mutation into an apparent failure: that invites duplicate edits.
				let compact: string | undefined;
				try {
					const value = JSON.parse(result);
					if (['office_document_edit', 'office_document_format'].includes(call.function.name) && typeof value.version === 'string' && !value.error) {
						compact = JSON.stringify({ completed: true, uri: value.uri, version: value.version, undoable: value.undoable, saved: value.saved, appliedEdits: value.appliedEdits, ...(JSON.stringify(value.created || []).length <= 8000 ? { created: value.created } : {}), contentTruncated: true, guidance: 'The edit completed. Read a smaller page or text chunk to inspect it. Do not repeat the completed edit.' });
					}
				} catch { /* Non-JSON results use the ordinary context limit error. */ }
				result = compact || JSON.stringify({ error: 'Result exceeds the context budget. Request a smaller page, text chunk, guide or exact command.' });
			}
			messages.push({ role: 'tool', tool_call_id: call.id, content: result });
		}
	}
	const completed = [...applied.values()].map(action => `${action.label}: ${action.ids.size} blocks`).join('; ');
	return completed ? 'Completed operations: ' + completed + '. Check document history for available Undo/Redo. The agent used its step budget before confirming any remaining requirements.' : 'The agent reached its step limit without a confirmed document edit.' + (lastError ? ' Last error: ' + lastError : '');
}
