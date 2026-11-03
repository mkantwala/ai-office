/*---------------------------------------------------------------------------------------------
 *  Copyright (c) Microsoft Corporation. All rights reserved.
 *  Licensed under the MIT License. See License.txt in the project root for license information.
 *--------------------------------------------------------------------------------------------*/

/**
 * Font family names installed on this machine, offered to the font family menu beside the
 * engine's own families. Only names are read out of the machine's existing font files; no
 * font data is bundled, and nothing here panics — unreadable files and directories are
 * skipped so the menu never depends on the disk layout.
 */
import { readdirSync, readFileSync } from "node:fs";
import { join } from "node:path";

let cache: string[] | undefined;

const STYLE_SUFFIX = / (?:regular|thin|extralight|light|semilight|medium|semibold|demibold|bold|extrabold|black|heavy)(?: italic)?$/i;

const utf16be = (buffer: Buffer) => {
	let text = "";
	for (let i = 0; i + 1 < buffer.length; i += 2) {
		text += String.fromCharCode(buffer.readUInt16BE(i));
	}
	return text;
};

/** Family names (name id 16, falling back to 1) from one OpenType font face. */
function faceFamily(buffer: Buffer, start: number): string {
	const numTables = buffer.readUInt16BE(start + 4);
	for (let t = 0; t < numTables; t++) {
		const record = start + 12 + t * 16;
		if (buffer.toString("latin1", record, record + 4) !== "name") {
			continue;
		}
		// OpenType table offsets are always from the beginning of the file, for
		// single fonts and for faces inside a collection alike.
		const offset = buffer.readUInt32BE(record + 8);
		const count = buffer.readUInt16BE(offset + 2);
		const strings = offset + buffer.readUInt16BE(offset + 4);
		let typographic = "";
		let legacy = "";
		for (let i = 0; i < count; i++) {
			const entry = offset + 6 + i * 12;
			const id = buffer.readUInt16BE(entry + 6);
			if (id !== 1 && id !== 16) {
				continue;
			}
			const length = buffer.readUInt16BE(entry + 8);
			const at = strings + buffer.readUInt16BE(entry + 10);
			if (length === 0 || at + length > buffer.length) {
				continue;
			}
			const platform = buffer.readUInt16BE(entry);
			const raw = (
				platform === 1
					? buffer.toString("latin1", at, at + length)
					: utf16be(buffer.subarray(at, at + length))
			);
			// Family names never carry control characters; a hit here means the
			// offset landed outside the real string data, so ignore that entry.
			if (!raw || /[\u0000-\u001f]/.test(raw)) {
				continue;
			}
			const text = raw.replace(/\s+/g, " ").trim();
			if (!text) {
				continue;
			}
			if (id === 16 && !typographic) {
				typographic = text;
			} else if (id === 1 && !legacy) {
				legacy = text;
			}
		}
		const family = typographic || legacy.replace(STYLE_SUFFIX, "");
		if (family) {
			return family;
		}
	}
	return "";
}

function collectionFamilies(buffer: Buffer): string[] {
	const count = Math.min(buffer.readUInt32BE(4), 64);
	const families: string[] = [];
	for (let i = 0; i < count && i < 64; i++) {
		try {
			const face = faceFamily(buffer, buffer.readUInt32BE(12 + i * 4));
			if (face) {
				families.push(face);
			}
		} catch {
			/* A torn collection entry stays out of the menu; the rest still load. */
		}
	}
	return families;
}

function fontDirectories(): string[] {
	const directories: string[] = [];
	const add = (directory: string | undefined, deep = false) => {
		if (directory) {
			directories.push(deep ? `${directory}?2` : directory);
		}
	};
	add(process.env.windir && join(process.env.windir, "Fonts"));
	add(
		process.env.LOCALAPPDATA &&
			join(process.env.LOCALAPPDATA, "Microsoft", "Windows", "Fonts"),
	);
	add("/Library/Fonts");
	add("/System/Library/Fonts", true);
	add("/usr/share/fonts", true);
	add("/usr/local/share/fonts", true);
	add(process.env.XDG_DATA_HOME && join(process.env.XDG_DATA_HOME, "fonts"), true);
	add(process.env.HOME && join(process.env.HOME, ".local", "share", "fonts"), true);
	add(process.env.HOME && join(process.env.HOME, "Library", "Fonts"));
	return directories;
}

/** Installed family names, sorted case-insensitively; empty when nothing could be read. */
export function systemFontFamilies(): string[] {
	if (!cache) {
		const seen = new Map<string, string>();
		for (const directory of fontDirectories()) {
			const deep = directory.endsWith("?2");
			const root = deep ? directory.slice(0, -2) : directory;
			const folders: string[] = [root];
			try {
				if (deep) {
					for (const entry of readdirSync(root, { withFileTypes: true })) {
						if (entry.isDirectory()) {
							folders.push(join(root, entry.name));
						}
					}
				}
			} catch {
				continue;
			}
			for (const folder of folders) {
				let files: string[] = [];
				try {
					files = readdirSync(folder);
				} catch {
					continue;
				}
				for (const file of files) {
					if (!/\.(ttf|otf|ttc|otc)$/i.test(file)) {
						continue;
					}
					try {
						const buffer = readFileSync(join(folder, file));
						const head = Buffer.from(buffer.subarray(0, 4)).toString("latin1");
						const names = head === "ttcf" ? collectionFamilies(buffer) : [faceFamily(buffer, 0)];
						for (const name of names) {
							if (!name || seen.has(name.toLowerCase())) {
								continue;
							}
							seen.set(name.toLowerCase(), name);
						}
					} catch {
						/* Unreadable font file: skip it, keep the rest of the menu. */
					}
					if (seen.size >= 600) {
						break;
					}
				}
			}
		}
		cache = [...seen.values()].sort((a, b) =>
			a.localeCompare(b, undefined, { sensitivity: "base" }),
		);
	}
	return cache;
}
