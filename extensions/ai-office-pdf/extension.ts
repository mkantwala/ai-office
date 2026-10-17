/*---------------------------------------------------------------------------------------------
 *  Copyright (c) Microsoft Corporation. All rights reserved.
 *  Licensed under the MIT License. See License.txt in the project root for license information.
 *--------------------------------------------------------------------------------------------*/

import * as vscode from 'vscode';
import { registerEditors } from '../ai-office-common/provider.js';
import { registerCloudflareAgent } from '../ai-office-common/cloudflareAgent.js';
export function activate(context: vscode.ExtensionContext) {
	registerEditors(context, ['pdf', 'word']);
	registerCloudflareAgent(context);
}
