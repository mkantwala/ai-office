/*---------------------------------------------------------------------------------------------
 *  Copyright (c) Microsoft Corporation. All rights reserved.
 *  Licensed under the MIT License. See License.txt in the project root for license information.
 *--------------------------------------------------------------------------------------------*/

import * as vscode from 'vscode';
import { validateDocumentEdits } from './wordFormatting.js';
import { withAgentToolRequest } from './agentRequest.js';

export interface DocumentToolInput {
	uri?: string;
	operation?: string;
	id?: string;
	version?: string;
	[key: string]: unknown;
}
export interface DocumentToolBackend {
	invoke(tool: 'read' | 'edit' | 'format' | 'help' | 'activity' | 'history' | 'draft' | 'comments', input: DocumentToolInput, token: vscode.CancellationToken): Promise<unknown>;
}

/** Keep the model-facing surface small; detailed command schemas are fetched through help. */
export function registerDocumentTools(context: vscode.ExtensionContext, backend: DocumentToolBackend) {
	for (const tool of ['read', 'edit', 'format', 'help', 'activity', 'history', 'draft', 'comments'] as const) {
		context.subscriptions.push(vscode.lm.registerTool<DocumentToolInput>('office_document_' + tool, {
			prepareInvocation: () => ({ invocationMessage: tool === 'format' ? 'Formatting the Word document' : tool === 'edit' ? 'Editing the Word document' : tool === 'read' ? 'Reading the Word document' : tool === 'help' ? 'Checking document editing capabilities' : tool === 'history' ? 'Checking document history' : tool === 'draft' ? 'Drafting document sections' : tool === 'comments' ? 'Working with document comments' : 'Working on the selected document content' }),
			async invoke(options, token) {
				try {
					if (token.isCancellationRequested) { throw new vscode.CancellationError(); }
					if (tool === 'edit') { validateDocumentEdits(options.input); }
					if (tool === 'draft' || tool === 'comments') { validateDocumentAuxiliary(tool, options.input); }
					const value = await withAgentToolRequest(options.toolInvocationToken, () => backend.invoke(tool, options.input, token));
					return new vscode.LanguageModelToolResult([new vscode.LanguageModelTextPart(JSON.stringify(value))]);
				} catch (error) {
					if (token.isCancellationRequested) { throw error; }
					const message = String(error);
					const stale = /Document changed|Block changed|no longer exists|target again|not found/i.test(message);
					return new vscode.LanguageModelToolResult([new vscode.LanguageModelTextPart(JSON.stringify({ error: message, retry: stale ? 'Read the target once to obtain its current version, then retry the corrected edit.' : 'Correct the arguments using the tool schema. For formatting use office_document_format. Do not repeat this unchanged call or reread the document to fix an argument error.' }))]);
				}
			}
		}));
	}
}


/** Validate conditional fields even when a non-model caller bypasses JSON-schema validation. */
function validateDocumentAuxiliary(tool: 'draft' | 'comments', input: DocumentToolInput) {
	const operation = input.operation ?? (tool === 'draft' ? 'status' : 'list');
	const allowed = tool === 'draft' ? ['plan', 'status', 'fill', 'discard'] : ['list', 'add', 'reply', 'update', 'resolve', 'delete'];
	if (!allowed.includes(operation)) { throw new Error(`Unsupported ${tool} operation.`); }
	if (tool === 'comments') {
		const fields: Record<string, string[]> = {
			list: ['commentId', 'offset', 'limit', 'textOffset', 'textLimit'], add: ['id', 'text', 'find', 'nth', 'start', 'end'],
			reply: ['commentId', 'text'], update: ['commentId', 'text', 'resolved'], resolve: ['commentId', 'resolved'], delete: ['commentId']
		};
		for (const key of Object.keys(input)) {
			if (!['operation', 'uri', 'version', ...fields[operation]].includes(key)) { throw new Error(`${operation} comments does not accept ${key}. Replies, updates, resolve and delete use commentId from a comment list; only add accepts a text anchor.`); }
		}
		if (input.resolved !== undefined && typeof input.resolved !== 'boolean') { throw new Error('resolved must be an explicit boolean.'); }
	}
	for (const key of ['uri', 'version', 'id', 'sectionId']) {
		if (input[key] !== undefined && (typeof input[key] !== 'string' || !input[key])) { throw new Error(`${key} must be a non-empty string.`); }
	}
	const mutation = !['status', 'list', 'discard'].includes(operation);
	if (mutation && typeof input.version !== 'string') { throw new Error('A fresh document version is required.'); }
	if (tool === 'draft') {
		if (operation === 'fill' && typeof input.sectionId !== 'string') { throw new Error('Fill requires a pending sectionId.'); }
		return;
	}
	if (operation !== 'list' && operation !== 'add' && (!Number.isInteger(input.commentId) || Number(input.commentId) < 0 || Number(input.commentId) > 0xffffffff)) { throw new Error('A valid commentId from a comment read is required.'); }
	if (operation === 'add' && typeof input.id !== 'string') { throw new Error('Add requires a paragraph ID.'); }
	if (['add', 'reply', 'update'].includes(operation) && (typeof input.text !== 'string' || !input.text || input.text.length > 10000)) { throw new Error('Comment text requires 1-10000 characters.'); }
	if (operation === 'resolve' && typeof input.resolved !== 'boolean') { throw new Error('resolved must be an explicit boolean.'); }
	for (const key of ['commentId', 'start', 'end', 'offset', 'limit', 'textOffset', 'textLimit', 'nth']) {
		if (input[key] !== undefined && (!Number.isSafeInteger(input[key]) || Number(input[key]) < (['limit', 'textLimit', 'nth'].includes(key) ? 1 : 0))) { throw new Error(`${key} must be a valid nonnegative integer (limits and nth start at 1).`); }
	}
}
