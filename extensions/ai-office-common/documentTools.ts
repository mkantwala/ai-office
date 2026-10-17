/*---------------------------------------------------------------------------------------------
 *  Copyright (c) Microsoft Corporation. All rights reserved.
 *  Licensed under the MIT License. See License.txt in the project root for license information.
 *--------------------------------------------------------------------------------------------*/

import * as vscode from 'vscode';

export interface DocumentToolInput {
	uri?: string;
	operation?: string;
	id?: string;
	version?: string;
	[key: string]: unknown;
}
export interface DocumentToolBackend {
	invoke(tool: 'read' | 'edit' | 'help' | 'activity', input: DocumentToolInput, token: vscode.CancellationToken): Promise<unknown>;
}

/** Keep the model-facing surface small; detailed command schemas are fetched through help. */
export function registerDocumentTools(context: vscode.ExtensionContext, backend: DocumentToolBackend) {
	for (const tool of ['read', 'edit', 'help', 'activity'] as const) {
		context.subscriptions.push(vscode.lm.registerTool<DocumentToolInput>('office_document_' + tool, {
			prepareInvocation: () => ({ invocationMessage: tool === 'edit' ? 'Editing the Word document' : tool === 'read' ? 'Reading the Word document' : tool === 'help' ? 'Checking document editing capabilities' : 'Working on the selected document content' }),
			async invoke(options, token) {
				try {
					if (token.isCancellationRequested) { throw new vscode.CancellationError(); }
					const value = await backend.invoke(tool, options.input, token);
					return new vscode.LanguageModelToolResult([new vscode.LanguageModelTextPart(JSON.stringify(value))]);
				} catch (error) {
					if (token.isCancellationRequested) { throw error; }
					return new vscode.LanguageModelToolResult([new vscode.LanguageModelTextPart(JSON.stringify({ error: String(error), retry: 'Read the target again if the document changed. Never guess block IDs or unsupported commands.' }))]);
				}
			}
		}));
	}
}
