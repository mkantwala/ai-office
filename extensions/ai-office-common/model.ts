/*---------------------------------------------------------------------------------------------
 *  Copyright (c) Microsoft Corporation. All rights reserved.
 *  Licensed under the MIT License. See License.txt in the project root for license information.
 *--------------------------------------------------------------------------------------------*/

import { readFile } from 'node:fs/promises';
import { join } from 'node:path';
import { pathToFileURL } from 'node:url';
import { SlidesEngine } from './slides.js';
import type { Engine, Kind, State, Command } from './types.js';

interface WasmDocument {
	pages(): string; state(): string; commands(): string; execute(id: string, args: string): string;
	export(): Uint8Array; render(page: number, scale: number): Uint8Array;
	image?(page: number, name: string, bytes: Uint8Array): string;
	width: number; height: number; free(): void;
}
export async function openEngine(kind: Kind, name: string, bytes: Uint8Array, extensionsRoot: string, password?: string): Promise<Engine> {
	if (bytes.length > 128 * 1024 * 1024) { throw new Error('Document exceeds the 128 MB limit.'); }
	if (bytes.length < 22) { throw new Error('This file is empty or incomplete. Create a valid document in its original application, then open it here.'); }
	if (kind === 'slides') { return SlidesEngine.open(bytes); }
	const project = kind === 'word' ? 'wordcraft' : 'pdfcraft';
	const directory = join(extensionsRoot, 'ai-office-pdf/media', kind === 'word' ? 'word-engine' : 'engine');
	// Generated engine loaders are selected at runtime and kept outside the host bundle.
	// eslint-disable-next-line no-restricted-syntax
	const module = await import(pathToFileURL(join(directory, project + '.js')).href);
	await module.default({ module_or_path: await readFile(join(directory, project + '_bg.wasm')) });
	const doc: WasmDocument = new module.OfficeDocument(name, bytes, password);
	const state = (): State => {
		const value = JSON.parse(doc.state());
		if (kind === 'pdf') { value.blocks = value.blocks.map((block: { paragraph: number; text: string; editor_text: string; rect: number[]; font: string; size: number }) => ({ id: block.paragraph, text: block.editor_text || block.text, rect: block.rect, font: block.font, size: block.size })); }
		return value;
	};
	return {
		pages: () => JSON.parse(doc.pages()), state,
		commands: () => {
			const commands: Command[] = JSON.parse(doc.commands());
			if (kind === 'pdf') { commands.push({ id: 'image.insert', label: 'Image', params: { page: '1-based page', data: 'base64 image', name: 'filename' } }, { id: 'select.page', label: 'Select page', params: { page: '1-based page' } }); }
			return commands;
		},
		execute: async (id, args) => {
			if (id === 'image.insert' && doc.image) {
				if (typeof args.data !== 'string' || args.data.length > 24 * 1024 * 1024) { throw new Error('Image data is missing or too large.'); }
				const page = Number(args.page);
				if (!Number.isInteger(page) || page < 1) { throw new Error('Page must be a positive integer.'); }
				doc.image(page - 1, typeof args.name === 'string' ? args.name : 'image.png', Buffer.from(args.data, 'base64'));
			} else { doc.execute(id, JSON.stringify(args)); }
			return state();
		},
		export: async () => doc.export(),
		render: (page, scale) => ({ pixels: doc.render(page, scale), width: doc.width, height: doc.height }),
		dispose: () => doc.free()
	};
}
