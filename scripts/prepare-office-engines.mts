/*---------------------------------------------------------------------------------------------
 *  Copyright (c) Microsoft Corporation. All rights reserved.
 *  Licensed under the MIT License. See License.txt in the project root for license information.
 *--------------------------------------------------------------------------------------------*/

import { readFile, writeFile, mkdir, access } from 'node:fs/promises';
import { createHash } from 'node:crypto';
import { spawnSync } from 'node:child_process';
import { dirname, resolve, join } from 'node:path';
import { fileURLToPath } from 'node:url';

const root = resolve(dirname(fileURLToPath(import.meta.url)), '..');
const sources = JSON.parse(await readFile(join(root, 'engines/sources.json'), 'utf8'));
const destination = join(root, '.build/engines');
const exists = async (path: string) => access(path).then(() => true, () => false);
function git(cwd: string, args: string[]) {
	const result = spawnSync('git', ['-C', cwd, ...args], { encoding: 'utf8' });
	if (result.error || result.status !== 0) { throw result.error || new Error(result.stderr || result.stdout); }
	return result.stdout.trim();
}
await mkdir(destination, { recursive: true });
for (const [name, source] of Object.entries(sources)) {
	const patch = join(root, 'engines', source.patch);
	const digest = createHash('sha256').update(await readFile(patch)).digest('hex');
	if (digest !== source.sha256) { throw new Error(`Patch checksum mismatch: ${name}`); }
	const checkout = join(destination, name);
	const stamp = join(checkout, '.ai-office-source.json');
	if (await exists(stamp)) {
		const installed = JSON.parse(await readFile(stamp, 'utf8'));
		if (installed.sha256 !== digest || git(checkout, ['rev-parse', 'HEAD']) !== source.revision) {
			throw new Error(`${checkout} has a different source version. Move it aside before preparing again; local edits will not be overwritten.`);
		}
		console.log(`${name}: already prepared (local changes preserved)`);
		continue;
	}
	if (await exists(checkout)) { throw new Error(`${checkout} already exists. Move it aside before preparing; it will not be overwritten.`); }
	await mkdir(checkout);
	console.log(`${name}: fetching ${source.revision}`);
	git(checkout, ['init']);
	git(checkout, ['remote', 'add', 'origin', source.repository]);
	git(checkout, ['fetch', '--depth=1', 'origin', source.revision]);
	git(checkout, ['checkout', '--detach', 'FETCH_HEAD']);
	git(checkout, ['apply', '--check', patch]);
	git(checkout, ['apply', patch]);
	await writeFile(stamp, JSON.stringify(source, null, 2) + '\n');
	console.log(`${name}: AI Office changes applied`);
}
