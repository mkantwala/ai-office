/*---------------------------------------------------------------------------------------------
 *  Copyright (c) Microsoft Corporation. All rights reserved.
 *  Licensed under the MIT License. See License.txt in the project root for license information.
 *--------------------------------------------------------------------------------------------*/

export { };
declare function acquireVsCodeApi(): { postMessage(value: unknown): void };
const api = acquireVsCodeApi();
interface WordModule { default(options: { module_or_path: Uint8Array }): Promise<unknown>; office_request(id: number, method: string, params: string, bytes: Uint8Array): void }
let engine: WordModule | undefined;
const failure = document.getElementById('failure')!;
function showError(error: unknown) { failure.hidden = false; failure.textContent = String(error); }
(globalThis as typeof globalThis & { aiOfficeWordEvent: (kind: string, id: number, data: string, bytes: Uint8Array) => void }).aiOfficeWordEvent = (kind, id, data, bytes) => {
	try {
		if (kind === 'ready') { document.getElementById('loading')!.hidden = true; }
		api.postMessage({ type: kind, id, data: JSON.parse(data), bytes: bytes.slice() });
	} catch (error) { showError(error); }
};
window.addEventListener('message', event => {
	if (event.origin !== location.origin) { return; }
	const message = event.data;
	if (message.type === 'error') { showError(message.error); return; }
	if (message.type !== 'request') { return; }
	try {
		if (!engine) { throw new Error('Wordcraft is loading.'); }
		engine.office_request(message.id, message.method, JSON.stringify(message.args || {}), new Uint8Array(message.bytes?.data || message.bytes || []));
	} catch (error) { api.postMessage({ type: 'response', id: message.id, data: { ok: false, error: String(error) } }); }
});
const config = JSON.parse(document.getElementById('config')!.textContent!);
try {
	// Generated WebAssembly bindings are a separate resource selected by the extension host.
	// eslint-disable-next-line no-restricted-syntax
	engine = await import(config.module);
	await engine!.default({ module_or_path: new Uint8Array(await (await fetch(config.wasm)).arrayBuffer()) });
} catch (error) { showError(error); api.postMessage({ type: 'fatal', error: String(error) }); }
