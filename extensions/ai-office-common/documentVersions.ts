/*---------------------------------------------------------------------------------------------
 *  Copyright (c) Microsoft Corporation. All rights reserved.
 *  Licensed under the MIT License. See License.txt in the project root for license information.
 *--------------------------------------------------------------------------------------------*/

import { createHash, randomUUID } from 'node:crypto';
import { mkdir, readFile, writeFile, rename, unlink } from 'node:fs/promises';
import { join } from 'node:path';

export interface DocumentVersion { id: string; time: string; label: string; hash: string; size: number }
interface Index { schema: 1; versions: DocumentVersion[]; pending?: { hash: string; review: string; request?: string } }
const digest = (value: Uint8Array | string) => createHash('sha256').update(value).digest('hex');

/** Local document versions, never Git objects or files beside the user's document. */
export class DocumentVersions {
	private readonly directory: string;
	constructor(root: string, uri: string) { this.directory = join(root, 'document-versions', digest(uri)); }
	private async index(): Promise<Index> {
		try {
			const bytes = await readFile(join(this.directory, 'index.json'));
			if (bytes.length > 64 * 1024 * 1024) { throw new Error('Document review exceeds local storage limit.'); }
			const data: Index = JSON.parse(bytes.toString());
			if (data.schema !== 1 || !Array.isArray(data.versions)) { throw new Error('Unsupported document version history.'); }
			return data;
		} catch (error) { if ((error as NodeJS.ErrnoException).code === 'ENOENT') { return { schema: 1, versions: [] }; } throw error; }
	}
	private async write(index: Index) {
		await mkdir(this.directory, { recursive: true, mode: 0o700 });
		const temporary = join(this.directory, `${randomUUID()}.tmp`);
		try { await writeFile(temporary, JSON.stringify(index), { mode: 0o600 }); await rename(temporary, join(this.directory, 'index.json')); }
		finally { await unlink(temporary).catch(() => { }); }
	}
	async capture(bytes: Uint8Array, projection: string, label: string, pending?: { review: string; request?: string }) {
		const index = await this.index(), hash = digest(bytes);
		const removed: DocumentVersion[] = [];
		if (index.versions.at(-1)?.hash !== hash) {
			await mkdir(this.directory, { recursive: true, mode: 0o700 });
			const version: DocumentVersion = { id: randomUUID(), time: new Date().toISOString(), label, hash, size: bytes.length + Buffer.byteLength(projection) };
			await writeFile(join(this.directory, `${version.id}.docx`), bytes, { mode: 0o600 });
			await writeFile(join(this.directory, `${version.id}.md`), projection, { mode: 0o600 });
			index.versions.push(version);
			while (index.versions.length > 30 || (index.versions.length > 1 && index.versions.reduce((sum, item) => sum + item.size, 0) > 64 * 1024 * 1024)) { removed.push(index.versions.shift()!); }
		}
		index.pending = pending ? { hash, ...pending } : undefined;
		await this.write(index);
		for (const item of removed) { for (const suffix of ['docx', 'md']) { await unlink(join(this.directory, `${item.id}.${suffix}`)).catch(() => { }); } }
	}
	async restoreReview(bytes: Uint8Array) {
		const index = await this.index();
		return index.pending?.hash === digest(bytes) ? index.pending : undefined;
	}
	async list() { return (await this.index()).versions.slice().reverse(); }
	async projection(id: string) {
		if (!(await this.index()).versions.some(item => item.id === id)) { throw new Error('Document version is no longer available.'); }
		return readFile(join(this.directory, `${id}.md`), 'utf8');
	}
}
