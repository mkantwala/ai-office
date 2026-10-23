/*---------------------------------------------------------------------------------------------
 *  Copyright (c) Microsoft Corporation. All rights reserved.
 *  Licensed under the MIT License. See License.txt in the project root for license information.
 *--------------------------------------------------------------------------------------------*/

/** Bounded, text-only presentation of a user-requested inspection. Agent reads never open it. */
export function renderPdfInspection(container: HTMLElement, value: unknown) {
	container.replaceChildren(); let remaining = 200;
	const label = (name: string) => name.replaceAll('_', ' ').replace(/\b\w/g, char => char.toUpperCase());
	const paint = (parent: HTMLElement, name: string, value: unknown, depth: number) => {
		if (--remaining < 0 || depth > 8 || ['doc', 'path', 'object_generation'].includes(name)) { return; }
		if (value && typeof value === 'object') {
			const details = document.createElement('details'), summary = document.createElement('summary'); summary.textContent = label(name) + (Array.isArray(value) ? ` (${value.length})` : ''); details.open = depth < 2; details.append(summary);
			const entries = Array.isArray(value) ? value.map((item, index) => [String(index + 1), item] as const) : Object.entries(value);
			for (const [key, item] of entries.slice(0, 40)) { paint(details, key, item, depth + 1); }
			if (entries.length > 40) { const more = document.createElement('p'); more.textContent = `Showing 40 of ${entries.length} entries. Narrow the inspection for more detail.`; details.append(more); }
			parent.append(details);
		} else if (value !== undefined && value !== null) {
			const row = document.createElement('p'), key = document.createElement('strong'); key.textContent = label(name) + ': '; row.append(key, document.createTextNode(String(value).slice(0, 4000))); parent.append(row);
		}
	};
	paint(container, 'Results', value, 0);
}
