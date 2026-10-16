/*---------------------------------------------------------------------------------------------
 *  Copyright (c) Microsoft Corporation. All rights reserved.
 *  Licensed under the MIT License. See License.txt in the project root for license information.
 *--------------------------------------------------------------------------------------------*/
import { spawnSync } from 'node:child_process';
import { mkdir, copyFile, readFile, readdir, writeFile } from 'node:fs/promises';
import { dirname, resolve, join } from 'node:path';
import { homedir } from 'node:os';
import { fileURLToPath } from 'node:url';
import { build } from 'esbuild';
import { existsSync } from 'node:fs';

const extension = dirname(fileURLToPath(import.meta.url));
const root = resolve(extension, '../..');
const fullWord = process.argv.includes('--word-full');
const word = fullWord || process.argv.includes('--word');
const project = word ? 'wordcraft' : 'pdfcraft';
const prepared = resolve(root, '.build/engines', project);
const source = (word ? process.env.WORDCRAFT_SOURCE : process.env.PDFCRAFT_SOURCE) || (existsSync(prepared) ? prepared : resolve(root, `../${project}`));
const target = process.env.CARGO_TARGET_DIR || (existsSync(prepared) ? resolve(root, '.build/office-rust-target') : resolve(root, '../pdfcraft/target'));
const output = join(extension, fullWord ? 'media/word-app' : word ? 'media/word-engine' : 'media/engine');
await Promise.all([
	build({ entryPoints: [join(extension, '../ai-office-common/engineWorker.ts')], bundle: true, platform: 'node', format: 'cjs', outfile: join(extension, 'out/engineWorker.cjs') }),
	build({ entryPoints: [join(extension, '../ai-office-common/cli.ts')], bundle: true, platform: 'node', format: 'cjs', outfile: join(extension, 'out/cli.cjs') }),
	build({ entryPoints: [join(extension, 'extension.ts')], bundle: true, platform: 'node', format: 'cjs', external: ['vscode'], outfile: join(extension, 'out/extension.js'), sourcemap: true }),
	build({ entryPoints: [join(extension, '../ai-office-common/editor.ts')], bundle: true, platform: 'browser', outfile: join(extension, 'out/editor.js'), sourcemap: true }),
	build({ entryPoints: [join(extension, '../ai-office-common/wordFull.ts')], bundle: true, platform: 'browser', format: 'esm', outfile: join(extension, 'out/wordFull.js'), sourcemap: true })
]);
await copyFile(join(extension, '../ai-office-common/editor.css'), join(extension, 'out/editor.css'));
await copyFile(join(extension, '../ai-office-common/wordFull.css'), join(extension, 'out/wordFull.css'));
await mkdir(join(extension, 'out/licenses'), { recursive: true });
for (const [name, file] of [['pptx', '@office-kit/pptx/LICENSE'], ['pptx-preview', '@office-kit/pptx-preview/LICENSE'], ['dompurify', 'dompurify/LICENSE'], ['fflate', 'fflate/LICENSE']]) {
	await copyFile(join(extension, '../ai-office-slides/node_modules', file), join(extension, `out/licenses/${name}.txt`));
}
if (process.argv.includes('--ui-only')) { process.exit(0); }
function run(command: string, args: string[], cwd = source) {
	const result = spawnSync(command, args, { cwd, stdio: 'inherit' });
	if (result.error || result.status !== 0) { throw result.error || new Error(`${command} exited with ${result.status}`); }
}
if (!process.argv.includes('--package-only')) {
	run('cargo', ['build', '-p', fullWord ? 'wordcraft-web' : `${project}-preview`, ...(fullWord ? ['--features', 'ai-office'] : []), '--locked', '--target-dir', target, '--target', 'wasm32-unknown-unknown', '--profile', 'low-disk', '-j', '2']);
}
await mkdir(output, { recursive: true });
await writeFile(join(output, 'package.json'), '{"type":"module"}\n');
const bindgen = process.env.WASM_BINDGEN || join(root, '.build/pdf-tools/bin', process.platform === 'win32' ? 'wasm-bindgen.exe' : 'wasm-bindgen');
run(bindgen, ['--target', 'web', '--out-dir', output, '--out-name', project, join(target, `wasm32-unknown-unknown/low-disk/${fullWord ? 'wordcraft-web' : project + '_preview'}.wasm`)]);
for (const name of ['LICENSE-MIT', 'LICENSE-APACHE', 'NOTICE', 'ATTRIBUTION.md', ...(word ? [] : ['ATTRIBUTION.toml'])]) {
	await copyFile(join(source, name), join(output, name));
}
const attribution = word ? '' : await readFile(join(source, 'ATTRIBUTION.toml'), 'utf8');
const assetLicences = new Set([...attribution.matchAll(/^licence_file\s*=\s*"([^"]+)"/gm)].map(match => match[1]));
if (word) {
	for (const file of await readdir(join(source, 'assets/fonts'))) {
		if (/license|ofl|copyright/i.test(file)) { assetLicences.add(`assets/fonts/${file}`); }
	}
}
for (const name of assetLicences) {
	// Cargo dependency notices are included in NOTICE; local asset licences keep their paths.
	if (name.includes('://') || name.startsWith('..')) { continue; }
	try {
		await mkdir(dirname(join(output, name)), { recursive: true });
		await copyFile(join(source, name), join(output, name));
	} catch (error) {
		if (!(error instanceof Error) || !('code' in error) || error.code !== 'ENOENT') { throw error; }
	}
}
const lock = await readFile(join(source, 'Cargo.lock'), 'utf8');
const registry = join(process.env.CARGO_HOME || join(homedir(), '.cargo'), 'registry/src');
for (const [crate, folder] of [['epaint_default_fonts', 'fonts'], ['hayro-cmap', 'assets']]) {
	const entry = lock.split('[[package]]').find(value => value.includes(`name = "${crate}"`));
	if (!entry) { continue; }
	const version = entry?.match(/version = "([^"]+)"/)?.[1];
	let copied = false;
	for (const index of await readdir(registry)) {
		const directory = join(registry, index, `${crate}-${version}`, folder);
		let files;
		try { files = await readdir(directory); } catch (error) { if (error instanceof Error && 'code' in error && error.code === 'ENOENT') { continue; } throw error; }
		const destination = join(output, 'dependency-licenses', `${crate}-${version}`);
		await mkdir(destination, { recursive: true });
		for (const file of files.filter(file => file.endsWith('.txt'))) { await copyFile(join(directory, file), join(destination, file)); }
		copied = true;
		break;
	}
	if (!copied) { throw new Error(`Missing bundled asset licences for ${crate}-${version} in the Cargo registry.`); }
}
const revision = spawnSync('git', ['-c', `safe.directory=${source.replaceAll('\\', '/')}`, 'rev-parse', 'HEAD'], { cwd: source, encoding: 'utf8' });
await writeFile(join(output, 'BUILD.txt'), `${project} ${revision.stdout.trim()} with AI Office integration.\nRust wasm32-unknown-unknown; profile low-disk; ${fullWord ? 'full Wordcraft web interface' : 'headless document engine'}.\n`);
console.log(`${project} engine installed in ${output}`);
