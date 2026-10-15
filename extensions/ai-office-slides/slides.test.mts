/*---------------------------------------------------------------------------------------------
 *  Copyright (c) Microsoft Corporation. All rights reserved.
 *  Licensed under the MIT License. See License.txt in the project root for license information.
 *--------------------------------------------------------------------------------------------*/

import { test } from 'node:test';
import assert from 'node:assert/strict';
import { createPresentation, addSlide, getSlideLayouts, addSlideTextBox, inches, getShapeText, getSlides, getSlideShapes, loadPresentation, savePresentation } from '@office-kit/pptx';
import { renderSlideToSvg } from '@office-kit/pptx-preview';
import { checkPresentationZip } from './zipLimits.ts';

test('renders original slide text after loading PPTX bytes', async () => {
	const deck = createPresentation();
	const slide = addSlide(deck, { layout: getSlideLayouts(deck)[0] });
	addSlideTextBox(slide, { x: inches(1), y: inches(1), w: inches(8), h: inches(1), text: 'AI Office presentation' });
	assert.match(renderSlideToSvg(deck, slide, { textLayout: 'svg' }), /AI Office/);
	const bytes = await savePresentation(deck);
	checkPresentationZip(bytes);
	const reopened = await loadPresentation(bytes);
	assert.ok(getSlideShapes(getSlides(reopened)[0]).some(shape => getShapeText(shape) === 'AI Office presentation'));
});

test('rejects malformed ZIPs and a declared expansion bomb before parsing', () => {
	assert.throws(() => checkPresentationZip(new Uint8Array(30)), /directory/);
	const bytes = new Uint8Array(68);
	const view = new DataView(bytes.buffer);
	view.setUint32(0, 0x02014b50, true);
	view.setUint32(24, 0xffffffff, true);
	view.setUint32(46, 0x06054b50, true);
	view.setUint16(56, 1, true);
	assert.throws(() => checkPresentationZip(bytes), /memory limit/);
});
