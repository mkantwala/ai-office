/*---------------------------------------------------------------------------------------------
 *  Copyright (c) Microsoft Corporation. All rights reserved.
 *  Licensed under the MIT License. See License.txt in the project root for license information.
 *--------------------------------------------------------------------------------------------*/

interface Raster { pixels: Uint8Array; width: number; height: number }

// Keep up to two page frames in the worker. Send only changed scanlines across the
// extension host/webview bridge, but only when the receiver still has that frame.
export class WordFrame {
	private readonly frames = new Map<number, Raster & { scale: number; frame: number }>();

	update(raster: Raster, page: number, scale: number, frame: number, base?: number) {
		const previous = this.frames.get(page);
		this.frames.delete(page);
		this.frames.set(page, { ...raster, scale, frame });
		while (this.frames.size > 1 && (this.frames.size > 2 || [...this.frames.values()].reduce((sum, value) => sum + value.pixels.byteLength, 0) > 32 * 1024 * 1024)) {
			this.frames.delete(this.frames.keys().next().value!);
		}
		if (!previous || previous.frame !== base || previous.scale !== scale || previous.width !== raster.width || previous.height !== raster.height) {
			return { ...raster, pixels: raster.pixels.slice() };
		}
		const stride = raster.width * 4;
		let start = 0, end = raster.pixels.length;
		while (start < end && raster.pixels[start] === previous.pixels[start]) { start++; }
		while (end > start && raster.pixels[end - 1] === previous.pixels[end - 1]) { end--; }
		const y = Math.floor(start / stride);
		const bottom = Math.ceil(end / stride);
		return { pixels: raster.pixels.slice(y * stride, bottom * stride), width: raster.width, height: bottom - y, y, canvasHeight: raster.height };
	}
}
