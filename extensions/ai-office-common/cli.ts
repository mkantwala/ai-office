/*---------------------------------------------------------------------------------------------
 *  Copyright (c) Microsoft Corporation. All rights reserved.
 *  Licensed under the MIT License. See License.txt in the project root for license information.
 *--------------------------------------------------------------------------------------------*/

import { readFile, writeFile, rename, unlink, realpath, stat } from 'node:fs/promises';
import { resolve, dirname, extname, basename, relative, isAbsolute, join } from 'node:path';
import { pathToFileURL } from 'node:url';
import { randomBytes, createHash } from 'node:crypto';
import { openEngine } from './model.js';
import type { Kind } from './types.js';

async function readBounded(path: string, limit: number) {
	if ((await stat(path)).size > limit) { throw new Error('Input exceeds the ' + limit / 1024 / 1024 + ' MB limit.'); }
	const bytes = await readFile(path);
	if (bytes.length > limit) { throw new Error('Input exceeds the size limit.'); }
	return bytes;
}

async function main() {
	const args = process.argv.slice(2);
	const action = args.shift();
	const file = args.shift();
	if (!file || !['commands', 'inspect', 'run'].includes(action || '')) {
		process.stdout.write('Sangam One CLI\n  node cli.cjs commands FILE\n  node cli.cjs inspect FILE\n  node cli.cjs run FILE --script commands.json --output result.docx [--root DIR]\nUse --in-place instead of --output to overwrite the input after a conflict check.\nCommand script: [{"command":"select.text","params":{"text":"Hello"}},{"command":"format.bold","params":{}}]\n');
		return;
	}
	const option = (name: string) => { const index = args.indexOf(name); return index >= 0 ? args[index + 1] : undefined; };
	const root = option('--root') ? await realpath(resolve(option('--root')!)) : undefined;
	async function confined(path: string, output = false) {
		const absolute = resolve(root || process.cwd(), path);
		const actual = output ? join(await realpath(dirname(absolute)), basename(absolute)) : await realpath(absolute);
		if (root) { const rel = relative(root, actual); if (rel === '..' || rel.startsWith('..' + (process.platform === 'win32' ? '\\' : '/')) || isAbsolute(rel)) { throw new Error('Path is outside --root.'); } }
		return actual;
	}
	const input = await confined(file);
	const extension = extname(input).toLowerCase();
	const kind: Kind = extension === '.docx' ? 'word' : extension === '.pdf' ? 'pdf' : extension === '.pptx' ? 'slides' : (() => { throw new Error('Use a DOCX, PDF or PPTX file.'); })();
	const original = await readBounded(input, 128 * 1024 * 1024);
	const engine = await openEngine(kind, basename(input), original, resolve(__dirname, '../..'), process.env.AI_OFFICE_PASSWORD, { resourceRoot: root, artifactRoot: root ? join(root, '.sangam-pdf-artifacts') : undefined, sourceUri: pathToFileURL(input).href });
	try {
		if (action === 'commands') { process.stdout.write(JSON.stringify(engine.commands(), null, 2) + '\n'); return; }
		if (action === 'inspect') {
			if (kind === 'word') { await engine.execute('select.all', {}); }
			process.stdout.write(JSON.stringify({ pages: engine.pages(), state: engine.state() }, null, 2) + '\n'); return;
		}
		const script = option('--script');
		if (!script) { throw new Error('--script is required.'); }
		const scriptBytes = await readBounded(await confined(script), 32 * 1024 * 1024);
		if (scriptBytes.length > 32 * 1024 * 1024) { throw new Error('Command script exceeds 32 MB.'); }
		const commands = JSON.parse(scriptBytes.toString('utf8'));
		if (!Array.isArray(commands) || commands.length > 10000) { throw new Error('Expected at most 10,000 commands.'); }
		const results = [];
		for (const step of commands) {
			if (!step || typeof step.command !== 'string' || (step.params !== undefined && (typeof step.params !== 'object' || step.params === null || Array.isArray(step.params)))) { throw new Error('Each step needs command and optional params object.'); }
			results.push(await engine.execute(step.command, step.params || {}));
		}
		const destination = args.includes('--in-place') ? input : option('--output');
		if (!destination) { throw new Error('Choose --output or --in-place. No file was changed.'); }
		const output = await confined(destination, true);
		if (extname(output).toLowerCase() !== extension) { throw new Error('Output must use the same document format.'); }
		const same = output.toLowerCase() === input.toLowerCase();
		if (same && !args.includes('--in-place')) { throw new Error('Use --in-place to overwrite the input.'); }
		if (same) {
			const digest = (bytes: Uint8Array) => createHash('sha256').update(bytes).digest('hex');
			if (digest(await readBounded(input, 128 * 1024 * 1024)) !== digest(original)) { throw new Error('Input changed during the operation; no output written.'); }
		}
		const bytes = await engine.export();
		const temp = join(dirname(output), '.' + basename(output) + '.' + randomBytes(8).toString('hex') + '.tmp');
		try { await writeFile(temp, bytes, { flag: 'wx' }); await rename(temp, output); }
		finally { await unlink(temp).catch(() => { }); }
		process.stdout.write(JSON.stringify({ output, commands: results.length, state: engine.state() }, null, 2) + '\n');
	} finally { engine.dispose(); }
}
void main().catch(error => { process.stderr.write(JSON.stringify({ error: String(error) }) + '\n'); process.exitCode = 1; });
