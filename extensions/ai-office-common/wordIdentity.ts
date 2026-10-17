/*---------------------------------------------------------------------------------------------
 *  Copyright (c) Microsoft Corporation. All rights reserved.
 *  Licensed under the MIT License. See License.txt in the project root for license information.
 *--------------------------------------------------------------------------------------------*/

import { createHash } from 'node:crypto';
import { strFromU8, strToU8, unzipSync, zipSync } from '../ai-office-slides/node_modules/fflate/esm/index.mjs';

const part = 'customXml/ai-office-blocks.xml';
const maxExpandedBytes = 256 * 1024 * 1024;

function unpack(bytes: Uint8Array) {
	let total = 0;
	return unzipSync(bytes, { filter: entry => {
		total += entry.originalSize;
		if (total > maxExpandedBytes) { throw new Error('Document exceeds the expanded package size limit.'); }
		return true;
	} });
}

function fingerprint(files: Record<string, Uint8Array>): string {
	const hash = createHash('sha256');
	for (const name of Object.keys(files).filter(name => name.startsWith('word/')).sort()) {
		hash.update(name).update('\0').update(files[name]);
	}
	return hash.digest('hex');
}

/** IDs are private package metadata. External edits invalidate the mapping instead of mis-targeting content. */
export function readWordIdentities(bytes: Uint8Array): unknown {
	const files = unpack(bytes);
	if (!files[part] || files[part].length > 8 * 1024 * 1024) { return []; }
	try {
		const encoded = /<identities xmlns="urn:ai-office:block-identities:v1">([A-Za-z0-9+/=]+)<\/identities>/.exec(strFromU8(files[part]))?.[1];
		if (!encoded) { return []; }
		const data = JSON.parse(Buffer.from(encoded, 'base64').toString('utf8'));
		return data.version === 1 && data.fingerprint === fingerprint(files) && Array.isArray(data.nodes) ? data.nodes : [];
	} catch { return []; }
}

export function writeWordIdentities(bytes: Uint8Array, nodes: unknown): Uint8Array {
	const files = unpack(bytes);
	const data = Buffer.from(JSON.stringify({ version: 1, fingerprint: fingerprint(files), nodes })).toString('base64');
	if (data.length > 8 * 1024 * 1024) { throw new Error('Document identity metadata exceeds the size limit.'); }
	files[part] = strToU8(`<?xml version="1.0" encoding="UTF-8"?><identities xmlns="urn:ai-office:block-identities:v1">${data}</identities>`);
	const types = strFromU8(files['[Content_Types].xml']);
	if (!types.includes('/' + part)) {
		files['[Content_Types].xml'] = strToU8(types.replace('</Types>', `<Override PartName="/${part}" ContentType="application/xml"/></Types>`));
	}
	// A distinct relationship avoids claiming the metadata is user-authored custom XML data.
	const rels = strFromU8(files['_rels/.rels']);
	if (!rels.includes('Target="' + part + '"')) {
		let id = 'aiOfficeBlockIdentities';
		while (rels.includes(`Id="${id}"`)) { id += '_'; }
		files['_rels/.rels'] = strToU8(rels.replace('</Relationships>', `<Relationship Id="${id}" Type="urn:ai-office:block-identities" Target="${part}"/></Relationships>`));
	}
	return zipSync(files);
}
