/*---------------------------------------------------------------------------------------------
 *  Copyright (c) Microsoft Corporation. All rights reserved.
 *  Licensed under the MIT License. See License.txt in the project root for license information.
 *--------------------------------------------------------------------------------------------*/

/** Bound declared ZIP expansion before the OOXML parser allocates its document parts. */
export function checkPresentationZip(bytes: Uint8Array): void {
	if (bytes.length < 22) { throw new Error('PowerPoint file is empty or incomplete.'); }
	if (bytes.length > 64 * 1024 * 1024) { throw new Error('Presentation exceeds the 64 MB limit.'); }
	const view = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);
	let end = bytes.length - 22;
	const first = Math.max(0, end - 65535);
	while (end >= first && view.getUint32(end, true) !== 0x06054b50) { end--; }
	if (end < first) { throw new Error('Presentation ZIP directory is missing.'); }
	const count = view.getUint16(end + 10, true);
	let offset = view.getUint32(end + 16, true);
	let expanded = 0;
	if (count > 20000) { throw new Error('Presentation contains too many parts.'); }
	for (let i = 0; i < count; i++) {
		if (offset + 46 > end || view.getUint32(offset, true) !== 0x02014b50) { throw new Error('Invalid presentation ZIP directory.'); }
		const size = view.getUint32(offset + 24, true);
		expanded += size;
		if (size > 64 * 1024 * 1024 || expanded > 256 * 1024 * 1024) { throw new Error('Expanded presentation exceeds the memory limit.'); }
		offset += 46 + view.getUint16(offset + 28, true) + view.getUint16(offset + 30, true) + view.getUint16(offset + 32, true);
	}
}
