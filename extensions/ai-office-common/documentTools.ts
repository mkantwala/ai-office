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
	invoke(tool: 'read' | 'edit' | 'format' | 'help' | 'activity' | 'history', input: DocumentToolInput, token: vscode.CancellationToken): Promise<unknown>;
}

/** Keep the model-facing surface small; detailed command schemas are fetched through help. */
export function registerDocumentTools(context: vscode.ExtensionContext, backend: DocumentToolBackend) {
	for (const tool of ['read', 'edit', 'format', 'help', 'activity', 'history'] as const) {
		context.subscriptions.push(vscode.lm.registerTool<DocumentToolInput>('office_document_' + tool, {
			prepareInvocation: () => ({ invocationMessage: tool === 'format' ? 'Formatting the Word document' : tool === 'edit' ? 'Editing the Word document' : tool === 'read' ? 'Reading the Word document' : tool === 'help' ? 'Checking document editing capabilities' : tool === 'history' ? 'Checking document history' : 'Working on the selected document content' }),
			async invoke(options, token) {
				try {
					if (token.isCancellationRequested) { throw new vscode.CancellationError(); }
					if (tool === 'edit') { validateDocumentEdits(options.input); }
					const value = await withAgentToolRequest(options.toolInvocationToken, () => backend.invoke(tool, options.input, token));
					return new vscode.LanguageModelToolResult([new vscode.LanguageModelTextPart(JSON.stringify(value))]);
				} catch (error) {
					if (token.isCancellationRequested) { throw error; }
					const message = String(error);
					const stale = /Document changed|no longer exists|target again/i.test(message);
					return new vscode.LanguageModelToolResult([new vscode.LanguageModelTextPart(JSON.stringify({ error: message, retry: stale ? 'Read the target once to obtain its current version, then retry the corrected edit.' : 'Correct the arguments using the tool schema. For formatting use office_document_format. Do not repeat this unchanged call or reread the document to fix an argument error.' }))]);
				}
			}
		}));
	}
}
