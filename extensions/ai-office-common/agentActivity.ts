/*---------------------------------------------------------------------------------------------
 *  Copyright (c) Microsoft Corporation. All rights reserved.
 *  Licensed under the MIT License. See License.txt in the project root for license information.
 *--------------------------------------------------------------------------------------------*/

import type { Kind, State, Block } from './types.js';

/** A rectangle in document points, relative to its page. */
export interface ActivityRect { x: number; y: number; width: number; height: number }
/** The page region an agent command is currently working on. */
export interface ActivityRegion { page: number; rects: ActivityRect[] }

const key = (block: Block) => block.id + ':' + (block.row ?? '') + ':' + (block.col ?? '');

const rect = (block: Block): ActivityRect => ({
	x: block.rect[0], y: block.rect[1],
	width: Math.max(block.rect[2] - block.rect[0], 4),
	height: Math.max(block.rect[3] - block.rect[1], 4)
});

function union(blocks: readonly Block[]): ActivityRect {
	const left = Math.min(...blocks.map(block => block.rect[0]));
	const top = Math.min(...blocks.map(block => block.rect[1]));
	const right = Math.max(...blocks.map(block => block.rect[2]));
	const bottom = Math.max(...blocks.map(block => block.rect[3]));
	return { x: left, y: top, width: Math.max(right - left, 4), height: Math.max(bottom - top, 4) };
}

function targetBlock(blocks: readonly Block[], args: Record<string, unknown>): Block | undefined {
	const id = typeof args.shape === 'number' ? args.shape : typeof args.paragraph === 'number' ? args.paragraph : undefined;
	if (id !== undefined) {
		const match = blocks.find(block => block.id === id && (args.row === undefined || (block.row === args.row && block.col === args.col)));
		if (match) { return match; }
	}
	const text = typeof args.text === 'string' && args.text ? args.text : typeof args.find === 'string' && args.find ? args.find : undefined;
	if (text) { return blocks.find(block => block.text === text) ?? blocks.find(block => block.text.includes(text)); }
	return undefined;
}

/**
 * Resolves the region an agent command worked on without any id supplied by the
 * caller: the current (or previous) selection for Word, the targeted block for
 * PDF/slides, or the blocks a command added.
 */
export function activityRegion(kind: Kind, args: Record<string, unknown>, before: State | undefined, after: State): ActivityRegion | undefined {
	if (kind === 'word') {
		const rects = (after.rects?.length ? after.rects : before?.rects) ?? [];
		if (rects.length) {
			const page = rects[0].page;
			return { page, rects: rects.filter(item => item.page === page).map(item => ({ x: item.x, y: item.y, width: Math.max(item.width, 2), height: Math.max(item.height, 4) })) };
		}
		const caret = after.caret ?? before?.caret;
		if (caret) { return { page: caret.page, rects: [{ x: caret.x, y: caret.y, width: 2, height: Math.max(caret.height, 4) }] }; }
		return undefined;
	}
	const page = after.page ?? before?.page ?? 0;
	const target = targetBlock(after.blocks ?? [], args) ?? (before ? targetBlock(before.blocks ?? [], args) : undefined);
	if (target) { return { page, rects: [rect(target)] }; }
	const known = new Set((before?.blocks ?? []).map(key));
	const fresh = (after.blocks ?? []).filter(block => !known.has(key(block)));
	if (fresh.length) { return { page, rects: [union(fresh)] }; }
	return undefined;
}
