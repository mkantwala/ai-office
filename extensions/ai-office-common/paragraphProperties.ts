/*---------------------------------------------------------------------------------------------
 *  Copyright (c) Microsoft Corporation. All rights reserved.
 *  Licensed under the MIT License. See License.txt in the project root for license information.
 *--------------------------------------------------------------------------------------------*/

export interface ParagraphTab { pos: number; align?: 'left' | 'center' | 'right' | 'decimal' | 'bar'; leader?: 'none' | 'dot' | 'hyphen' | 'underscore' | 'middleDot' }
export interface ParagraphBorder { width?: number; color?: string; style?: 'single' | 'double' | 'dotted' | 'dashed' }
export interface ParagraphProperties {
	alignment?: 'left' | 'center' | 'right' | 'justify' | 'distribute';
	indents?: { left?: number; right?: number; firstLine?: number };
	spacing?: { before?: number; after?: number };
	lineSpacing?: { rule: 'multiple' | 'exactly' | 'atLeast'; value: number };
	tabs?: ParagraphTab[];
	borders?: Partial<Record<'top' | 'left' | 'bottom' | 'right' | 'between', ParagraphBorder | null>>;
	shading?: string;
	keepNext?: boolean; keepLines?: boolean; pageBreakBefore?: boolean; widowControl?: boolean;
}

const numberSchema = (minimum: number, maximum: number) => ({ type: 'number', minimum, maximum });
const enumSchema = (...values: string[]) => ({ type: 'string', enum: values });
const objectSchema = (properties: Record<string, object>, required?: string[]) => ({ type: 'object', properties, additionalProperties: false, minProperties: 1, ...(required ? { required } : {}) });
const colorSchema = { type: 'string', pattern: '^[a-fA-F0-9]{6}$' };
/** Shared source for the strict agent schema; distances are points. */
export const paragraphPropertiesSchema = objectSchema({
	alignment: enumSchema('left', 'center', 'right', 'justify', 'distribute'),
	indents: objectSchema({ left: numberSchema(-1584, 1584), right: numberSchema(-1584, 1584), firstLine: numberSchema(-1584, 1584) }),
	spacing: objectSchema({ before: numberSchema(0, 1584), after: numberSchema(0, 1584) }),
	lineSpacing: objectSchema({ rule: enumSchema('multiple', 'exactly', 'atLeast'), value: numberSchema(.5, 1584) }, ['rule', 'value']),
	tabs: { type: 'array', maxItems: 64, items: objectSchema({ pos: numberSchema(0, 1584), align: enumSchema('left', 'center', 'right', 'decimal', 'bar'), leader: enumSchema('none', 'dot', 'hyphen', 'underscore', 'middleDot') }, ['pos']) },
	borders: objectSchema(Object.fromEntries(['top', 'left', 'bottom', 'right', 'between'].map(side => [side, { anyOf: [{ type: 'null' }, objectSchema({ width: numberSchema(.25, 6), color: colorSchema, style: enumSchema('single', 'double', 'dotted', 'dashed') })] }]))),
	shading: colorSchema,
	keepNext: { type: 'boolean' }, keepLines: { type: 'boolean' }, pageBreakBefore: { type: 'boolean' }, widowControl: { type: 'boolean' }
});

/** Validate before resolving targets; the native engine independently validates the same patch. */
export function validateParagraphProperties(value: unknown): asserts value is ParagraphProperties {
	const object = (v: unknown, keys: string[]): Record<string, unknown> => {
		if (!v || typeof v !== 'object' || Array.isArray(v) || !Object.keys(v).length) { throw new Error('Paragraph properties require a non-empty object.'); }
		for (const key of Object.keys(v)) { if (!keys.includes(key)) { throw new Error(`Unsupported paragraph property: ${key}.`); } }
		return v as Record<string, unknown>;
	};
	const number = (v: unknown, min: number, max: number) => { if (typeof v !== 'number' || !Number.isFinite(v) || v < min || v > max) { throw new Error(`Paragraph value must be between ${min} and ${max}.`); } };
	const choice = (v: unknown, values: string[]) => { if (typeof v !== 'string' || !values.includes(v)) { throw new Error('Unsupported paragraph property value.'); } };
	const color = (v: unknown) => { if (typeof v !== 'string' || !/^[a-f\d]{6}$/i.test(v)) { throw new Error('Color must be six hexadecimal digits.'); } };
	const patch = object(value, Object.keys(paragraphPropertiesSchema.properties));
	for (const [key, v] of Object.entries(patch)) {
		if (key === 'alignment') { choice(v, ['left', 'center', 'right', 'justify', 'distribute']); }
		else if (key === 'indents' || key === 'spacing') { const fields = object(v, key === 'indents' ? ['left', 'right', 'firstLine'] : ['before', 'after']); for (const n of Object.values(fields)) { number(n, key === 'indents' ? -1584 : 0, 1584); } }
		else if (key === 'lineSpacing') { const fields = object(v, ['rule', 'value']); choice(fields.rule, ['multiple', 'exactly', 'atLeast']); number(fields.value, fields.rule === 'multiple' ? .5 : 1, fields.rule === 'multiple' ? 10 : 1584); }
		else if (key === 'tabs') {
			if (!Array.isArray(v) || v.length > 64) { throw new Error('Use at most 64 tab stops.'); }
			const positions: number[] = [];
			for (const tab of v) { const t = object(tab, ['pos', 'align', 'leader']); number(t.pos, 0, 1584); if (positions.some(pos => Math.abs(pos - Number(t.pos)) < .01)) { throw new Error('Tab positions must be distinct.'); } positions.push(Number(t.pos)); if (t.align !== undefined) { choice(t.align, ['left', 'center', 'right', 'decimal', 'bar']); } if (t.leader !== undefined) { choice(t.leader, ['none', 'dot', 'hyphen', 'underscore', 'middleDot']); } }
		} else if (key === 'borders') {
			for (const border of Object.values(object(v, ['top', 'left', 'bottom', 'right', 'between']))) { if (border === null) { continue; } const b = object(border, ['width', 'color', 'style']); if (b.width !== undefined) { number(b.width, .25, 6); } if (b.color !== undefined) { color(b.color); } if (b.style !== undefined) { choice(b.style, ['single', 'double', 'dotted', 'dashed']); } }
		} else if (key === 'shading') { color(v); }
		else if (typeof v !== 'boolean') { throw new Error(`${key} must be true or false.`); }
	}
}
