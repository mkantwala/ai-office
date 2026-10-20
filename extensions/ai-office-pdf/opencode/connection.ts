/*---------------------------------------------------------------------------------------------
 *  Copyright (c) Microsoft Corporation. All rights reserved.
 *  Licensed under the MIT License. See License.txt in the project root for license information.
 *--------------------------------------------------------------------------------------------*/

import { spawn, type ChildProcessWithoutNullStreams } from 'node:child_process';
import { Readable, Writable } from 'node:stream';
import { ClientSideConnection, ndJsonStream, PROTOCOL_VERSION, type Client, type InitializeResponse, type McpServer, type NewSessionResponse } from '@agentclientprotocol/sdk';

/** One ACP process per conversation. OpenCode owns authentication, tools and durable session IDs. */
export class OpenCodeConnection {
	private child?: ChildProcessWithoutNullStreams;
	private connection?: ClientSideConnection;
	private capabilities?: InitializeResponse;
	private stopped = false;
	private sessionId?: string;

	constructor(private readonly executable: string, private readonly cwd: string, private readonly client: Client) { }

	get alive() { return !!this.connection && !this.stopped && !this.connection.signal.aborted; }

	private async deadline<T>(operation: Promise<T>, milliseconds = 30000): Promise<T> {
		let timer: ReturnType<typeof setTimeout> | undefined;
		try {
			return await Promise.race([operation, new Promise<never>((_, reject) => {
				timer = setTimeout(() => { this.dispose(); reject(new Error('OpenCode did not respond. Check the executable and provider configuration, then retry.')); }, milliseconds);
			})]);
		} finally { clearTimeout(timer); }
	}

	async start(): Promise<InitializeResponse> {
		if (this.capabilities && this.alive) { return this.capabilities; }
		if (this.stopped) { throw new Error('This OpenCode connection has closed. Start it again to resume the session.'); }
		const child = this.child = spawn(this.executable, ['acp'], { cwd: this.cwd, env: { ...process.env, OPENCODE_CLIENT: 'sangam-one' }, stdio: 'pipe', windowsHide: true });
		// Drain diagnostics without collecting provider tokens or unbounded process output.
		child.stderr.resume();
		const started = new Promise<void>((resolve, reject) => {
			child.once('spawn', resolve);
			child.once('error', error => reject(new Error(`Could not start OpenCode (${error.message}). Install OpenCode or set its executable in OpenCode: Configure Connection.`)));
		});
		try {
			await started;
			const connection = this.connection = new ClientSideConnection(() => this.client, ndJsonStream(Writable.toWeb(child.stdin) as WritableStream<Uint8Array>, Readable.toWeb(child.stdout) as ReadableStream<Uint8Array>, { maxMessageBytes: 8 * 1024 * 1024 }));
			void connection.closed.catch(() => {}).finally(() => { this.stopped = true; });
			this.capabilities = await this.deadline(connection.initialize({ protocolVersion: PROTOCOL_VERSION, clientInfo: { name: 'sangam-one', title: 'Sangam One', version: '0.1.0' }, clientCapabilities: { fs: { readTextFile: false, writeTextFile: false }, terminal: false } }));
			return this.capabilities;
		} catch (error) { this.dispose(); throw error; }
	}

	async open(mcpServers: McpServer[], sessionId?: string): Promise<NewSessionResponse> {
		const capabilities = await this.start();
		if (sessionId) {
			if (!capabilities.agentCapabilities?.loadSession) { throw new Error('This OpenCode version cannot resume sessions. Update OpenCode or start a new session.'); }
			this.sessionId = sessionId;
			const loaded = await this.deadline(this.connection!.loadSession({ sessionId, cwd: this.cwd, mcpServers }));
			return { ...loaded, sessionId };
		}
		const created = await this.deadline(this.connection!.newSession({ cwd: this.cwd, mcpServers }));
		this.sessionId = created.sessionId;
		return created;
	}

	async configure(id: string, value: string) {
		if (!this.alive || !this.sessionId) { throw new Error('The OpenCode session is not connected.'); }
		return this.deadline(this.connection!.setSessionConfigOption({ sessionId: this.sessionId, configId: id, value }));
	}

	async prompt(text: string) {
		if (!this.alive || !this.sessionId) { throw new Error('The OpenCode session is not connected.'); }
		return this.connection!.prompt({ sessionId: this.sessionId, prompt: [{ type: 'text', text }] });
	}

	async cancel() {
		if (!this.alive || !this.sessionId) { return; }
		await this.connection!.cancel({ sessionId: this.sessionId });
	}

	dispose() {
		if (this.stopped && !this.child) { return; }
		this.stopped = true;
		const child = this.child; this.child = undefined;
		if (child && child.exitCode === null && child.signalCode === null) {
			child.kill();
			const timer = setTimeout(() => { if (child.exitCode === null && child.signalCode === null) { child.kill('SIGKILL'); } }, 2000);
			timer.unref(); child.once('exit', () => clearTimeout(timer));
		}
	}
}
