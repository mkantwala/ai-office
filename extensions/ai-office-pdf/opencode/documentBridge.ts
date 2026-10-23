/*---------------------------------------------------------------------------------------------
 *  Copyright (c) Microsoft Corporation. All rights reserved.
 *  Licensed under the MIT License. See License.txt in the project root for license information.
 *--------------------------------------------------------------------------------------------*/

import { createServer, type Server } from 'node:http';
import { randomBytes, timingSafeEqual } from 'node:crypto';

export const documentToolNames = new Set(['office_document_tools', 'office_document_run', 'office_document_read', 'office_document_edit', 'office_document_help', 'office_document_format', 'office_document_history', 'office_document_draft', 'office_document_comments', 'office_document_activity']);
export interface DocumentTool { name: string; description: string; inputSchema: object }
export interface DocumentToolResult { content: { type: 'text'; text: string }[]; isError?: boolean }

/** An authenticated loopback MCP bridge exposes the same guarded document tools as the built-in agent. */
export class OpenCodeDocumentBridge {
	private readonly token = randomBytes(32).toString('hex');
	private server?: Server;
	constructor(private readonly tools: () => DocumentTool[], private readonly invoke: (name: string, input: object) => Promise<DocumentToolResult>) { }

	async start(): Promise<{ url: string; token: string }> {
		if (this.server) { throw new Error('Document bridge is already running.'); }
		const expected = Buffer.from(`Bearer ${this.token}`);
		const server = this.server = createServer(async (request, response) => {
			const supplied = Buffer.from(request.headers.authorization || '');
			if (supplied.length !== expected.length || !timingSafeEqual(supplied, expected)) { response.writeHead(401).end(); return; }
			if (request.method !== 'POST' || request.url !== '/mcp') { response.writeHead(405).end(); return; }
			try {
				let size = 0; const chunks: Buffer[] = [];
				for await (const chunk of request) { size += chunk.length; if (size > 1024 * 1024) { response.writeHead(413).end(); return; } chunks.push(chunk); }
				const message = JSON.parse(Buffer.concat(chunks).toString('utf8')) as { id?: string | number; method?: string; params?: { name?: string; arguments?: object } };
				if (message.id === undefined) { response.writeHead(202).end(); return; }
				let result: object;
				if (message.method === 'initialize') { result = { protocolVersion: '2024-11-05', serverInfo: { name: 'sangam-docs', version: '0.1.0' }, capabilities: { tools: {} } }; }
				else if (message.method === 'ping') { result = {}; }
				else if (message.method === 'tools/list') { result = { tools: this.tools().filter(tool => documentToolNames.has(tool.name)) }; }
				else if (message.method === 'tools/call') {
					const name = message.params?.name;
					if (!name || !documentToolNames.has(name)) { throw new Error('Unknown document tool.'); }
					const input = message.params?.arguments ?? {};
					if (!input || typeof input !== 'object' || Array.isArray(input)) { throw new Error('Document tool arguments must be an object.'); }
					try { result = await this.invoke(name, input); } catch (error) { result = { isError: true, content: [{ type: 'text', text: error instanceof Error ? error.message : 'Document tool failed.' }] }; }
				} else { response.writeHead(200, { 'Content-Type': 'application/json' }).end(JSON.stringify({ jsonrpc: '2.0', id: message.id, error: { code: -32601, message: 'Unknown MCP method' } })); return; }
				response.writeHead(200, { 'Content-Type': 'application/json' }).end(JSON.stringify({ jsonrpc: '2.0', id: message.id, result }));
			} catch { response.writeHead(400).end(); }
		});
		server.requestTimeout = 30000;
		await new Promise<void>((resolve, reject) => { server.once('error', reject); server.listen(0, '127.0.0.1', resolve); });
		const address = server.address();
		if (!address || typeof address === 'string') { throw new Error('Could not open the document tool bridge.'); }
		return { url: `http://127.0.0.1:${address.port}/mcp`, token: this.token };
	}

	dispose() { this.server?.close(); this.server?.closeAllConnections(); this.server = undefined; }
}
