/*---------------------------------------------------------------------------------------------
 *  Copyright (c) Microsoft Corporation. All rights reserved.
 *  Licensed under the MIT License. See License.txt in the project root for license information.
 *--------------------------------------------------------------------------------------------*/

import { createInterface } from 'node:readline';

// OpenCode starts this small stdio MCP adapter; the credential stays out of process arguments.
const url = process.env.SANGAM_DOCS_URL;
const token = process.env.SANGAM_DOCS_TOKEN;
if (!url || !token || !/^http:\/\/127\.0\.0\.1:\d+\/mcp$/.test(url)) { process.exit(1); }
const lines = createInterface({ input: process.stdin });
lines.on('line', line => {
	void (async () => {
		if (Buffer.byteLength(line) > 1024 * 1024) { process.exit(1); }
		const message = JSON.parse(line) as { id?: string | number };
		try {
			const response = await fetch(url!, { method: 'POST', headers: { Authorization: `Bearer ${token}`, 'Content-Type': 'application/json' }, body: line, signal: AbortSignal.timeout(120000) });
			if (message.id === undefined) { return; }
			if (!response.ok) { throw new Error('The document tool connection is unavailable.'); }
			process.stdout.write(JSON.stringify(await response.json()) + '\n');
		} catch (error) {
			if (message.id !== undefined) { process.stdout.write(JSON.stringify({ jsonrpc: '2.0', id: message.id, error: { code: -32603, message: error instanceof Error ? error.message : 'Document tool failed' } }) + '\n'); }
		}
	})().catch(() => process.exit(1));
});
