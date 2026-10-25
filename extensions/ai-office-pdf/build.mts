/*---------------------------------------------------------------------------------------------
 *  Copyright (c) Microsoft Corporation. All rights reserved.
 *  Licensed under the MIT License. See License.txt in the project root for license information.
 *--------------------------------------------------------------------------------------------*/
import { spawnSync } from 'node:child_process';
import { mkdir, copyFile, readFile, readdir, writeFile, chmod } from 'node:fs/promises';
import { dirname, resolve, join } from 'node:path';
import { homedir } from 'node:os';
import { fileURLToPath } from 'node:url';
import { build } from 'esbuild';
import { existsSync } from 'node:fs';

const extension = dirname(fileURLToPath(import.meta.url));
const root = resolve(extension, '../..');
const word = process.argv.includes('--word');
const project = word ? 'wordcraft' : 'pdfcraft';
const prepared = resolve(root, '.build/engines', project);
const vendored = resolve(root, 'engines/wordcraft');
/* The Wordcraft view engine is vendored in-repo for the DOCX viewer; the fetch+patch
 * path only applies to PdfCraft (and to a WORDCRAFT_SOURCE override for experiments). */
const source = (word
	? (process.env.WORDCRAFT_SOURCE || (existsSync(vendored) ? vendored : (existsSync(prepared) ? prepared : resolve(root, `../${project}`))))
	: (process.env.PDFCRAFT_SOURCE || (existsSync(prepared) ? prepared : resolve(root, `../${project}`))));
const target = process.env.CARGO_TARGET_DIR || (existsSync(prepared) ? resolve(root, '.build/office-rust-target') : resolve(root, '../pdfcraft/target'));
const native = !word;
const output = native ? join(extension, 'media/native', process.platform + '-' + process.arch) : join(extension, word ? 'media/word-engine' : 'media/engine');
await Promise.all([
	build({ entryPoints: [join(extension, '../ai-office-common/engineWorker.ts')], bundle: true, platform: 'node', format: 'cjs', outfile: join(extension, 'out/engineWorker.cjs') }),
	build({ entryPoints: [join(extension, '../ai-office-common/cli.ts')], bundle: true, platform: 'node', format: 'cjs', outfile: join(extension, 'out/cli.cjs') }),
	build({ entryPoints: [join(extension, 'extension.ts')], bundle: true, platform: 'node', format: 'cjs', external: ['vscode'], outfile: join(extension, 'out/extension.js'), sourcemap: true }),
	build({ entryPoints: [join(extension, '../ai-office-common/editor.ts')], bundle: true, platform: 'browser', outfile: join(extension, 'out/editor.js'), sourcemap: true }),
	build({ entryPoints: [join(extension, 'opencode/mcpProxy.ts')], bundle: true, platform: 'node', format: 'cjs', outfile: join(extension, 'out/opencodeMcpProxy.cjs') })
]);
await writeFile(join(extension, 'out/editor.css'), (await readFile(join(root, 'src/vs/base/browser/ui/documentToolbar/documentToolbar.css'), 'utf8')) + '\n' + (await readFile(join(extension, '../ai-office-common/editor.css'), 'utf8')));
await mkdir(join(extension, 'out/licenses'), { recursive: true });
await copyFile(join(extension, 'node_modules/@agentclientprotocol/sdk/LICENSE'), join(extension, 'out/licenses/acp-sdk.txt'));
await copyFile(join(extension, 'node_modules/zod/LICENSE'), join(extension, 'out/licenses/zod.txt'));
for (const [name, file] of [['pptx', '@office-kit/pptx/LICENSE'], ['pptx-preview', '@office-kit/pptx-preview/LICENSE'], ['dompurify', 'dompurify/LICENSE'], ['fflate', 'fflate/LICENSE']]) {
	await copyFile(join(extension, '../ai-office-slides/node_modules', file), join(extension, `out/licenses/${name}.txt`));
}
if (process.argv.includes('--ui-only')) { process.exit(0); }
function run(command: string, args: string[], cwd = source) {
	const result = spawnSync(command, args, { cwd, stdio: 'inherit' });
	if (result.error || result.status !== 0) { throw result.error || new Error(`${command} exited with ${result.status}`); }
}
if (!process.argv.includes('--package-only')) {
	if (native) { run('cargo', ['build', '-p', 'pdfcraft-preview', '--bin', 'sangam-pdf', '--locked', '--target-dir', target, '--profile', 'low-disk', '-j', '2']); }
	else { run('cargo', ['build', '-p', `${project}-preview`, '--locked', '--target-dir', target, '--target', 'wasm32-unknown-unknown', '--profile', 'low-disk', '-j', '2']); }
}
await mkdir(output, { recursive: true });
if (native) {
	const executable = process.platform === 'win32' ? 'sangam-pdf.exe' : 'sangam-pdf';
	await copyFile(join(target, 'low-disk', executable), join(output, executable));
	if (process.platform !== 'win32') { await chmod(join(output, executable), 0o755); const strip = spawnSync('strip', [process.platform === 'darwin' ? '-S' : '--strip-debug', join(output, executable)]); if (strip.error || strip.status !== 0) { throw strip.error || new Error('Could not strip native PDF debug symbols. Install the platform build tools.'); } }
} else {
	await writeFile(join(output, 'package.json'), '{"type":"module"}\n');
	const bindgen = process.env.WASM_BINDGEN || join(root, '.build/pdf-tools/bin', process.platform === 'win32' ? 'wasm-bindgen.exe' : 'wasm-bindgen');
	run(bindgen, ['--target', 'web', '--out-dir', output, '--out-name', project, join(target, `wasm32-unknown-unknown/low-disk/${project}_preview.wasm`)]);
}
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
const engineLabel = (word && source === vendored) ? 'vendored engines/wordcraft (DOCX view engine)' : revision.stdout.trim();
await writeFile(join(output, 'BUILD.txt'), `${project} ${engineLabel} with local content-only preview crate.\nRust ${native ? 'native ' + process.platform + '-' + process.arch : 'wasm32-unknown-unknown'}; profile low-disk; ${native ? 'persistent confined automation engine (OCR excluded)' : 'content-only preview renderer'}.\n`);
console.log(`${project} engine installed in ${output}`);
