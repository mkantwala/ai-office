/*---------------------------------------------------------------------------------------------
 *  Copyright (c) Microsoft Corporation. All rights reserved.
 *  Licensed under the MIT License. See License.txt in the project root for license information.
 *--------------------------------------------------------------------------------------------*/

const shapes = {
	text: { path: 'M11 5h10M16 5v22M11 27h10', fallback: 'text' },
	column: { path: 'M3 16h26M3 16l5-5M3 16l5 5M29 16l-5-5M29 16l-5 5M13 7v18M19 7v18', fallback: 'col-resize' },
	row: { path: 'M16 3v26M16 3l-5 5M16 3l5 5M16 29l-5-5M16 29l5-5M7 13h18M7 19h18', fallback: 'row-resize' },
	corner: { path: 'M6 6l20 20M6 6h8M6 6v8M26 26h-8M26 26v-8', fallback: 'nwse-resize' },
};
const cursors = new Map<keyof typeof shapes, string>();

/** Page pointers use a dark stroke and white halo independently of the app/OS theme. */
export function wordCursor(kind: keyof typeof shapes): string {
	let cursor = cursors.get(kind);
	if (!cursor) {
		const { path, fallback } = shapes[kind];
		const svg = `<svg xmlns="http://www.w3.org/2000/svg" width="32" height="32" viewBox="0 0 32 32"><path d="${path}" fill="none" stroke="white" stroke-width="4" stroke-linejoin="round"/><path d="${path}" fill="none" stroke="#171b26" stroke-width="1.5" stroke-linejoin="round"/></svg>`;
		cursor = `url("data:image/svg+xml,${encodeURIComponent(svg)}") 16 16, ${fallback}`;
		cursors.set(kind, cursor);
	}
	return cursor;
}
