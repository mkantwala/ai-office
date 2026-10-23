/*---------------------------------------------------------------------------------------------
 *  Copyright (c) Microsoft Corporation. All rights reserved.
 *  Licensed under the MIT License. See License.txt in the project root for license information.
 *--------------------------------------------------------------------------------------------*/

import { spawn, type ChildProcessWithoutNullStreams } from 'node:child_process';
import { readFile, writeFile, mkdir, mkdtemp, stat, lstat, readdir, copyFile, link, unlink, rm, realpath } from 'node:fs/promises';
import { join, basename, dirname, resolve, relative, extname, isAbsolute } from 'node:path';
import * as paths from 'node:path';
import { tmpdir } from 'node:os';
import { fileURLToPath, pathToFileURL } from 'node:url';
import { randomUUID } from 'node:crypto';
import type { Engine, Command, State, Raster } from './types.js';

export interface PdfArtifact { uri: string; name: string; mime: string; size: number }
interface NativeView { pages: number[][]; state: State }

/** Rust canonicalizes Windows paths to the verbatim namespace. Read only our known resource,
 * allowing equivalent path spellings without trusting a worker-selected file location. */
export function pdfResourcePath(root: string, reported: unknown, name: 'snapshot.pdf' | 'raster.raw', api: Pick<typeof paths, 'resolve' | 'relative' | 'toNamespacedPath'> = paths): string {
	const invalid = () => new Error(name === 'raster.raw' ? 'Invalid PDF raster resource.' : 'Invalid PDF export resource.');
	if (typeof reported !== 'string' || !reported || reported.length > 32768 || reported.includes('\0')) { throw invalid(); }
	const expected = api.resolve(root, name), candidate = api.resolve(root, reported);
	if (api.relative(api.toNamespacedPath(expected), api.toNamespacedPath(candidate)) !== '') { throw invalid(); }
	return expected;
}

const fileKeys = new Set(['path', 'paths', 'file', 'files', 'folder', 'out', 'out_dir', 'output', 'directories', 'image', 'certificate', 'certificates']);
const outputKeys: Record<string, string[]> = {
	doc_batch: ['out_dir'], page_extract: ['out', 'out_dir'], doc_split: ['out_dir'], doc_create_multiple: ['out', 'out_dir'], doc_combine: ['out'], doc_create: ['out'],
	doc_export_images: ['folder'], doc_export_all_images: ['folder'], doc_export_text: ['path'], doc_export_office: ['path'], doc_export_data: ['path'],
	doc_reduce: ['path'], doc_optimize: ['path'], doc_compare_report: ['path'], measure_export: ['out'], accessibility_report: ['path'], accessibility_review_record: ['path'], accessibility_review_save: ['path'], accessibility_review_html: ['path'],
	sign_id_create: ['path'], sign_document: ['out'], doc_print: ['path'], form_merge_data: ['path'], action_run: ['folder'], comments_summarize: ['out'], image_save: ['path']
};
const outputNames: Record<string, string> = { doc_export_office: 'Document', doc_export_text: 'Extracted Text', doc_export_data: 'Form Data', doc_reduce: 'Compressed', doc_optimize: 'Optimized', doc_compare_report: 'Comparison', measure_export: 'Measurements', accessibility_report: 'Accessibility Report', accessibility_review_record: 'Accessibility Review', accessibility_review_html: 'Accessibility Review', accessibility_review_save: 'Review Decisions', sign_id_create: 'Signing Identity', sign_document: 'Signed Document', doc_print: 'Print Layout', image_save: 'Image' };
const outputExtensions: Record<string, string> = { doc_export_text: '.txt', doc_export_data: '.xfdf', doc_export_office: '.docx', measure_export: '.csv', accessibility_review_save: '.json', accessibility_review_html: '.html', sign_id_create: '.p12', form_merge_data: '.csv', image_save: '.png', accessibility_report: '.html', accessibility_review_record: '.json' };
const mimeTypes: Record<string, string> = { '.pdf': 'application/pdf', '.png': 'image/png', '.jpg': 'image/jpeg', '.jpeg': 'image/jpeg', '.tif': 'image/tiff', '.tiff': 'image/tiff', '.docx': 'application/vnd.openxmlformats-officedocument.wordprocessingml.document', '.pptx': 'application/vnd.openxmlformats-officedocument.presentationml.presentation', '.xlsx': 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet', '.html': 'text/html', '.txt': 'text/plain', '.csv': 'text/csv', '.json': 'application/json', '.xfdf': 'application/vnd.adobe.xfdf' };

function sameResource(a: string, b: string): boolean {
	try { const left = resolve(fileURLToPath(a)), right = resolve(fileURLToPath(b)); return process.platform === 'win32' ? left.toLowerCase() === right.toLowerCase() : left === right; } catch { return false; }
}

function resource(value: string): string {
	if (!value.startsWith('file:')) { throw new Error('File parameters require explicit file:// URIs. Binary data and raw paths do not belong in agent parameters.'); }
	return fileURLToPath(value);
}

/** A persistent native session preserves PdfCraft history while confining engine file access to a private staging directory. */
export class NativePdfEngine implements Engine {
	private child!: ChildProcessWithoutNullStreams;
	private sequence = 0;
	private buffer = '';
	private stopped?: Error;
	private readonly pending = new Map<number, { resolve(value: unknown): void; reject(error: Error): void; timer: ReturnType<typeof setTimeout> }>();
	private view!: NativeView;
	private definitions: Command[] = [];
	private readonly published = new Map<string, string>();
	private readonly sourceUris = new Map<string, string>();
	private readonly cachedRasters = new Map<string, Raster>();
	private artifactDirectory = '';
	private artifactRoot = '';
	private generated: PdfArtifact[] = [];
	private constructor(private readonly root: string, private readonly sourceUri?: string, private readonly resourceRoot?: string) { }

	static async open(executable: string, name: string, bytes: Uint8Array, password?: string, options?: { artifactRoot?: string; sourceUri?: string; resourceRoot?: string }): Promise<NativePdfEngine> {
		// Match Rust's canonical path convention for every private IPC resource, including artifact receipts.
		const root = paths.toNamespacedPath(await realpath(await mkdtemp(join(tmpdir(), 'sangam-pdf-'))));
		const engine = new NativePdfEngine(root, options?.sourceUri, options?.resourceRoot);
		try {
			await mkdir(join(root, 'input')); await mkdir(join(root, 'output'));
			const artifacts = options?.artifactRoot || join(tmpdir(), 'sangam-pdf-artifacts');
			engine.artifactRoot = artifacts;
			const input = join(root, 'input', basename(name)); await writeFile(input, bytes, { mode: 0o600 });
			if (options?.sourceUri) { engine.sourceUris.set(input, options.sourceUri); }
			engine.child = spawn(executable, [root], { stdio: ['pipe', 'pipe', 'pipe'], windowsHide: true });
			engine.child.stdout.setEncoding('utf8');
			engine.child.stdout.on('data', (chunk: string) => {
				engine.buffer += chunk;
				if (engine.buffer.length > 64 * 1024 * 1024) { engine.fail(new Error('PDF engine response exceeded its memory budget.')); return; }
				let at: number;
				while ((at = engine.buffer.indexOf('\n')) >= 0) {
					const line = engine.buffer.slice(0, at); engine.buffer = engine.buffer.slice(at + 1);
					try {
						const message = JSON.parse(line); const request = engine.pending.get(message.id); if (!request) { continue; }
						engine.pending.delete(message.id); clearTimeout(request.timer);
						if (message.error) { request.reject(new Error(String(message.error).replaceAll(root, '[private PDF workspace]'))); } else { request.resolve(message.result); }
					} catch { engine.fail(new Error('PDF engine returned an invalid response.')); }
				}
			});
			// Consume diagnostics without recording PDF text, paths, passwords or certificate details.
			engine.child.stderr.on('data', () => {});
			engine.child.on('error', () => engine.fail(new Error('The native PDF engine could not start. Rebuild or reinstall the PDF extension.')));
			engine.child.on('exit', () => engine.fail(new Error('The native PDF engine stopped. Reopen the document.')));
			engine.view = engine.normalize(await engine.request<NativeView>('open', { path: input, password }));
			engine.definitions = await engine.request<Command[]>('commands');
			for (const command of engine.definitions) { engine.describeResources(command); }
			engine.definitions.push({ id: 'select.page', label: 'Select Page', readOnly: true, params: { type: 'object', properties: { page: { type: 'integer', minimum: 1 } }, required: ['page'], additionalProperties: false } });
			return engine;
		} catch (error) { engine.dispose(); throw error; }
	}
	private normalize(view: NativeView): NativeView {
		view.state.blocks = (view.state.blocks || []).map(block => {
			const source = block as unknown as { paragraph: number; text: string; editor_text?: string; rect: number[]; font?: string; size?: number };
			return { id: source.paragraph, text: source.editor_text ?? source.text, rect: source.rect, font: source.font, size: source.size };
		});
		return view;
	}
	private request<T>(method: string, args: object = {}): Promise<T> {
		if (this.stopped) { return Promise.reject(this.stopped); }
		return new Promise((resolve, reject) => {
			const id = ++this.sequence;
			let payload: string; try { payload = JSON.stringify({ id, method, args }) + '\n'; } catch { reject(new Error('PDF request is not serializable.')); return; }
			const timer = setTimeout(() => this.fail(new Error('PDF operation timed out. Reopen the document before retrying.')), 115000);
			this.pending.set(id, { resolve: value => resolve(value as T), reject, timer });
			this.child.stdin.write(payload, error => { if (error) { this.fail(new Error('The native PDF connection closed.')); } });
		});
	}
	private fail(error: Error) { if (this.stopped) { return; } this.stopped = error; for (const request of this.pending.values()) { clearTimeout(request.timer); request.reject(error); } this.pending.clear(); this.child?.kill(); }
	private describeResources(command: Command) {
		const schema = command.params as { properties?: Record<string, { type?: string; description?: string; items?: { type?: string }; enum?: string[] }>; required?: string[] };
		for (const [key, property] of Object.entries(schema.properties || {})) {
			if (command.id === 'sign_document' && key === 'id') { property.description = 'Input digital-ID file URI, or a keychain:/windows: identity returned by this host.'; }
			else if (key === 'other' || ['before', 'after'].includes(key) && command.id.startsWith('accessibility_review_')) { property.type = 'string'; property.description = 'File URI of the older PDF to compare. The host loads it without exposing engine IDs.'; }
			else if (fileKeys.has(key) && (property.type === 'string' || property.type === 'array' && property.items?.type === 'string')) { property.description = (outputKeys[command.id]?.includes(key) ? 'Optional output file/directory URI. Omit to create a managed artifact. Existing files are never overwritten. ' : 'Input file/directory URI; the host stages its bytes. ') + (property.description || ''); }
		}
		if (['doc_export_office', 'doc_export_data'].includes(command.id) && schema.properties) { schema.properties.format = { type: 'string', enum: command.id === 'doc_export_office' ? ['docx', 'pptx', 'html', 'rtf'] : ['xfdf', 'fdf', 'xml', 'xdp', 'csv', 'txt'], description: 'Output format for the managed artifact. An explicit output URI must use the same extension.' }; }
		const defaults = outputKeys[command.id]; if (defaults && schema.required) { schema.required = schema.required.filter(key => !defaults.includes(key)); }
		if (['sign_windows_ids', 'sign_keychain_ids', 'printers', 'printer_options', 'doc_print'].includes(command.id)) { command.description = `${command.description} Availability depends on this host's operating system, configured printers or certificate store.`; }
	}
	private async confined(path: string) {
		if (!this.resourceRoot) { return; }
		const root = await realpath(this.resourceRoot);
		const inside = (candidate: string) => { const value = relative(root, candidate); return value !== '..' && !value.startsWith('..' + (process.platform === 'win32' ? '\\' : '/')) && !isAbsolute(value); };
		if (!inside(resolve(path))) { throw new Error('Resource is outside --root.'); }
		let existing = resolve(path);
		while (true) {
			try { const actual = await realpath(existing); if (!inside(actual)) { throw new Error('Resource is outside --root.'); } return; }
			catch (error) { if ((error as NodeJS.ErrnoException).code !== 'ENOENT' || dirname(existing) === existing) { throw error; } existing = dirname(existing); }
		}
	}
	private async stage(uri: string, destination: string, budget: { bytes: number; files: number }, depth = 0): Promise<void> {
		const path = resource(uri); await this.confined(path); const info = await lstat(path);
		if (info.isSymbolicLink() || depth > 8 || ++budget.files > 1000) { throw new Error('Input folders must contain at most 1,000 files, without links, within eight levels.'); }
		if (info.isDirectory()) {
			await mkdir(destination, { recursive: true });
			for (const entry of await readdir(path)) { await this.stage(pathToFileURL(join(path, entry)).href, join(destination, entry), budget, depth + 1); }
		} else {
			budget.bytes += info.size; if (!info.isFile() || info.size > 128 * 1024 * 1024 || budget.bytes > 256 * 1024 * 1024) { throw new Error('Input files exceed the 128 MB per-file or 256 MB combined budget.'); }
			if (this.sourceUri && sameResource(this.sourceUri, uri)) { const bytes = await this.export(); await writeFile(destination, bytes, { mode: 0o600 }); } else { await copyFile(path, destination); } this.sourceUris.set(destination, uri);
		}
	}
	private async prepare(command: string, params: Record<string, unknown>, job: string) {
		const targets = new Map<string, string>(), sources: number[] = [], inputs: string[] = [], budget = { bytes: 0, files: 0 };
		const outputs = outputKeys[command] || [];
		const walk = async (value: unknown, key: string): Promise<unknown> => {
			if (typeof value === 'string' && (fileKeys.has(key) || key === 'to_pdf' || key === 'other' || ['before', 'after'].includes(key) && command.startsWith('accessibility_review_') || key === 'id' && command === 'sign_document' && value.startsWith('file:'))) {
				if (outputs.includes(key) || key === 'to_pdf') {
					const isFolder = ['folder', 'out_dir'].includes(key);
					const destination = resource(value); await this.confined(destination);
					if (this.sourceUri && sameResource(pathToFileURL(destination).href, this.sourceUri)) { throw new Error('Export to a new file so the open document and its history stay intact.'); }
					if (!isFolder) { try { await stat(destination); throw new Error('Output already exists. Choose a new file name.'); } catch (error) { if ((error as NodeJS.ErrnoException).code !== 'ENOENT') { throw error; } } }
					const path = join(job, isFolder ? key : basename(destination)); if (isFolder) { await mkdir(path, { recursive: true }); }
					targets.set(path, destination); return path;
				}
				const source = resource(value), folder = join(this.root, 'input', randomUUID()); await mkdir(folder); const destination = join(folder, basename(source));
				await this.stage(value, destination, budget); inputs.push(destination);
				if (key === 'other' || ['before', 'after'].includes(key) && command.startsWith('accessibility_review_')) { const docs = await this.request<{ doc: number }[]>('open_source', { path: destination }); const id = docs[0]?.doc; if (!id) { throw new Error('Comparison source did not open.'); } sources.push(id); return id; }
				return destination;
			}
			if (Array.isArray(value)) { return Promise.all(value.map(item => walk(item, key))); }
			if (value && typeof value === 'object') { return Object.fromEntries(await Promise.all(Object.entries(value).map(async ([child, item]) => [child, await walk(item, child)]))); }
			return value;
		};
		const prepared = await walk(params, '') as Record<string, unknown>;
		const defaults: Record<string, string> = { doc_batch: 'out_dir', doc_split: 'out_dir', doc_create_multiple: 'out_dir', doc_export_images: 'folder', doc_export_all_images: 'folder', sign_document: 'out', doc_export_text: 'path', doc_export_office: 'path', doc_export_data: 'path', doc_reduce: 'path', doc_optimize: 'path', doc_compare_report: 'path', measure_export: 'out', accessibility_report: 'path', accessibility_review_record: 'path', accessibility_review_save: 'path', accessibility_review_html: 'path', sign_id_create: 'path', action_run: 'folder', image_save: 'path' };
		const output = command === 'doc_print' && !params.printer ? 'path' : defaults[command];
		if (output && prepared[output] === undefined) {
			const folder = ['folder', 'out_dir'].includes(output);
			let extension = outputExtensions[command] || '.pdf';
			if (command === 'doc_export_office' && typeof params.format === 'string') { extension = '.' + params.format; }
			if (command === 'doc_export_data' && typeof params.format === 'string') { extension = '.' + params.format; }
			prepared[output] = folder ? join(job, 'files') : join(job, (outputNames[command] || 'PDF Output') + extension);
			if (folder) { await mkdir(prepared[output] as string, { recursive: true }); }
		}
		if (['doc_export_office', 'doc_export_data'].includes(command)) { if (params.format && extname(String(prepared.path)).slice(1) !== params.format) { throw new Error('Output format and file extension must match.'); } delete prepared.format; }
		return { params: prepared, targets, sources, inputs };
	}
	private async artifacts(targets: Map<string, string>): Promise<{ artifacts: PdfArtifact[]; paths: Map<string, string> }> {
		const paths = new Map<string, string>(), artifacts: PdfArtifact[] = []; let bytes = 0, count = 0;
		const visit = async (path: string) => {
			for (const entry of await readdir(path, { withFileTypes: true })) {
				const source = join(path, entry.name); if (entry.isSymbolicLink()) { throw new Error('The PDF engine returned an unexpected link.'); }
				if (entry.isDirectory()) { await visit(source); continue; }
				const info = await stat(source), stamp = `${info.mtimeMs}:${info.size}`;
				if (this.published.get(source) === stamp) { continue; }
				if (++count > 1000 || (bytes += info.size) > 256 * 1024 * 1024) { throw new Error('PDF outputs exceeded the 1,000-file / 256 MB artifact budget.'); }
				if (!this.artifactDirectory) { await mkdir(this.artifactRoot, { recursive: true }); this.artifactDirectory = await mkdtemp(join(this.artifactRoot, 'session-')); }
				let destination = join(this.artifactDirectory, randomUUID() + '-' + entry.name);
				for (const [root, target] of targets) { const suffix = relative(root, source); if (suffix === '' || suffix !== '..' && !suffix.startsWith('..' + (process.platform === 'win32' ? '\\' : '/')) && !isAbsolute(suffix)) { destination = suffix ? join(target, suffix) : target; break; } }
				await mkdir(dirname(destination), { recursive: true });
				const staging = join(dirname(destination), '.' + randomUUID() + '.tmp');
				try { await writeFile(staging, await readFile(source), { mode: 0o600, flag: 'wx' }); await link(staging, destination); } finally { await unlink(staging).catch(() => {}); }
				const uri = pathToFileURL(destination).href; paths.set(source, uri); paths.set(relative(this.root, source).replaceAll('\\', '/'), uri);
				artifacts.push({ uri, name: targets.size ? basename(destination) : entry.name, mime: mimeTypes[extname(destination).toLowerCase()] || 'application/octet-stream', size: info.size }); this.published.set(source, stamp);
			}
		};
		await visit(join(this.root, 'output')); return { artifacts, paths };
	}
	private project(value: unknown, paths: Map<string, string>): unknown {
		if (typeof value === 'string') { return paths.get(value) || this.sourceUris.get(value) || (value.startsWith(this.root) ? '[private engine resource]' : value); }
		if (Array.isArray(value)) { return value.map(item => this.project(item, paths)); }
		if (value && typeof value === 'object') { return Object.fromEntries(Object.entries(value).map(([key, item]) => [key, this.project(item, paths)])); }
		return value;
	}
	pages() { return this.view.pages; }
	state() { return this.view.state; }
	commands() { return this.definitions; }
	async execute(command: string, params: Record<string, unknown>): Promise<State> {
		if (command !== 'image.insert' && !this.definitions.some(definition => definition.id === command)) { throw new Error('Command is not enabled here; application lifecycle and OCR commands use separate controls or are excluded.'); }
		const job = join(this.root, 'output', randomUUID()); await mkdir(job);
		if (command === 'image.insert') {
			if (typeof params.data !== 'string' || params.data.length > 24 * 1024 * 1024) { throw new Error('Image data is missing or too large.'); }
			const image = join(this.root, 'input', randomUUID() + '.png'); await writeFile(image, Buffer.from(params.data, 'base64'));
			command = 'page_add_image'; params = { page: params.page, path: pathToFileURL(image).href };
		}
		const prepared = await this.prepare(command, params, job);
		let view: NativeView;
		try { view = this.normalize(await this.request<NativeView>('execute', { command, params: prepared.params })); }
		finally { for (const id of prepared.sources) { await this.request('close_source', { doc: id }).catch(() => {}); } if (command !== 'text_font_import' && command !== 'text_font_files') { await Promise.all(prepared.inputs.map(path => rm(path, { recursive: true, force: true }).catch(() => {}))); } }
		this.view = view; this.cachedRasters.clear();
		try {
			const { artifacts, paths } = await this.artifacts(prepared.targets);
			view.state.agentCommandResult = this.project(view.state.agentCommandResult, paths); this.generated = [...artifacts, ...this.generated].slice(0, 1000); view.state.pdfArtifacts = this.generated; view.state.pdfNewArtifacts = artifacts;
			await rm(join(this.root, 'output'), { recursive: true, force: true }); await mkdir(join(this.root, 'output')); this.published.clear();
		} catch (error) {
			// Content may already be committed: retain its revision and prevent blind duplicate edits.
			view.state.agentCommandResult = { nativeResult: view.state.agentCommandResult, warning: 'The operation completed, but exporting its artifacts failed: ' + String(error) };
		}
		return view.state;
	}
	async saved() { const previous = this.view.state; this.view = this.normalize(await this.request<NativeView>('saved')); this.view.state.agentCommandResult = previous.agentCommandResult; this.view.state.pdfArtifacts = this.generated; this.view.state.pdfNewArtifacts = []; this.cachedRasters.clear(); return this.view.state; }
	async export() { const result = await this.request<{ path: string }>('export'); return new Uint8Array(await readFile(pdfResourcePath(this.root, result.path, 'snapshot.pdf'))); }
	render(page: number, scale: number): Raster { const cached = this.cachedRasters.get(`${page}:${scale}`); if (!cached) { throw new Error('Use the asynchronous PDF renderer for an uncached page.'); } return cached; }
	async renderAsync(page: number, scale: number): Promise<Raster> {
		const key = `${page}:${scale}`, cached = this.cachedRasters.get(key); if (cached) { return cached; }
		const result = await this.request<{ path: string; width: number; height: number }>('render', { page, scale });
		const path = pdfResourcePath(this.root, result.path, 'raster.raw');
		const raster = { width: result.width, height: result.height, pixels: new Uint8Array(await readFile(path)) };
		if (raster.pixels.length !== raster.width * raster.height * 4) { throw new Error('Incomplete PDF raster.'); }
		if (this.cachedRasters.size >= 2) { this.cachedRasters.delete(this.cachedRasters.keys().next().value!); } this.cachedRasters.set(key, raster); return raster;
	}
	dispose() { this.fail(new Error('PDF document closed.')); void rm(this.root, { recursive: true, force: true }).catch(() => {}); }
}
