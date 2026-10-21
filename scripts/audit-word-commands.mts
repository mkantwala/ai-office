/*---------------------------------------------------------------------------------------------
 *  Copyright (c) Microsoft Corporation. All rights reserved.
 *  Licensed under the MIT License. See License.txt in the project root for license information.
 *--------------------------------------------------------------------------------------------*/
import { readFile, writeFile } from 'node:fs/promises';
import { resolve } from 'node:path';
import { pathToFileURL } from 'node:url';

// Run after building Word. Inspect the actual registry, including generated commands;
// source-text matching is insufficient for a capability inventory.
const directory = resolve('extensions/ai-office-pdf/media/word-engine');
const engine = await import(pathToFileURL(resolve(directory, 'wordcraft.js')).href);
await engine.default({ module_or_path: await readFile(resolve(directory, 'wordcraft_bg.wasm')) });
const document = new engine.OfficeDocument('Audit.txt', new TextEncoder().encode('Synthetic audit'), undefined);
try {
	const source = JSON.parse(await readFile('engines/sources.json', 'utf8')).wordcraft;
	const { commands } = JSON.parse(document.command_catalog());
	const inventory = {
		repository: source.repository, revision: source.revision,
		commands: commands.map(({ id, label, integrated, readOnly }: { id: string; label: string; integrated: boolean; readOnly: boolean }) => ({
			command: id, label, readOnly, integration: integrated ? /^(edit\.|caret\.|select\.)|^picture\.select$/.test(id) ? 'editor' : 'block-command' : 'pending'
		})).sort((a: { command: string }, b: { command: string }) => a.command.localeCompare(b.command, 'en'))
	};
	const path = 'extensions/ai-office-pdf/WORD_CAPABILITIES.json';
	const text = JSON.stringify(inventory, null, 2) + '\n';
	if (process.argv.includes('--write')) { await writeFile(path, text); }
	else if (await readFile(path, 'utf8') !== text) { throw new Error('Word capability inventory differs from the built registry. Run with --write after reviewing the changes.'); }
	console.log(`${commands.length} registry commands; ${commands.filter((command: { integrated: boolean }) => command.integrated).length} integrated; inventory verified.`);
} finally { document.free(); }
