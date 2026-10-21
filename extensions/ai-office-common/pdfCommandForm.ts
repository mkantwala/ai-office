/*---------------------------------------------------------------------------------------------
 *  Copyright (c) Microsoft Corporation. All rights reserved.
 *  Licensed under the MIT License. See License.txt in the project root for license information.
 *--------------------------------------------------------------------------------------------*/

import type { Command } from './types.js';
import { commandCategory } from './documentCapabilities.js';
import { pdfPageRange } from './pdfTools.js';

interface Schema {
	type?: string | string[]; title?: string; description?: string; enum?: (string | number | boolean)[];
	properties?: Record<string, Schema>; required?: string[]; items?: Schema; additionalProperties?: Schema | boolean;
	minimum?: number; maximum?: number; minItems?: number; maxItems?: number; minLength?: number; maxLength?: number;
	oneOf?: Schema[]; anyOf?: Schema[];
}
interface Field { element: HTMLElement; read(): unknown }
const humanize = (name: string) => name.replaceAll('_', ' ').replace(/\b\w/g, letter => letter.toUpperCase());

/** All native operations share a generated visual form; command schemas stay internal UI data. */
export class PdfCommandForm {
	private dialog?: HTMLDialogElement;
	constructor(private readonly snapshot: () => { commands: Command[]; total: number; revision: number }, private readonly execute: (command: string, params: Record<string, unknown>, revision: number) => void, private readonly pickFile?: (folders: boolean) => Promise<string | undefined>) { }
	dispose() { this.dialog?.remove(); }
	open() {
		if (this.dialog?.open) { this.dialog.focus(); return; } this.dialog?.remove();
		const snapshot = this.snapshot();
		const dialog = this.dialog = document.createElement('dialog'); dialog.className = 'pdf-tools-dialog pdf-command-dialog'; dialog.setAttribute('aria-label', 'All PDF Tools');
		const title = document.createElement('h2'); title.textContent = 'All PDF Tools';
		const search = document.createElement('input'); search.type = 'search'; search.placeholder = 'Search tools · forms, bookmarks, export…'; search.setAttribute('aria-label', 'Search PDF tools');
		const select = document.createElement('select'); select.setAttribute('aria-label', 'PDF tool');
		const description = document.createElement('p'), fields = document.createElement('div'); fields.className = 'pdf-command-fields';
		const form = document.createElement('form'); form.noValidate = true;
		const error = document.createElement('p'); error.setAttribute('role', 'alert'); error.className = 'pdf-tool-error'; error.hidden = true;
		const footer = document.createElement('footer'), cancel = document.createElement('button'), apply = document.createElement('button'); cancel.type = 'button'; cancel.textContent = 'Close'; cancel.addEventListener('click', () => dialog.close()); apply.type = 'submit'; apply.textContent = 'Apply'; footer.append(cancel, apply);
		let root: Field | undefined, definition: Command | undefined;
		const render = () => {
			definition = snapshot.commands.find(command => command.id === select.value); fields.replaceChildren(); error.hidden = true;
			if (!definition) { description.textContent = 'No matching tools.'; root = undefined; apply.disabled = true; return; }
			description.textContent = definition.description || definition.label; apply.textContent = definition.readOnly ? 'Inspect' : 'Apply'; apply.disabled = false;
			root = this.field(definition.params as Schema, 'Parameters', true, snapshot.total, 0); fields.append(root.element);
		};
		const filter = () => {
			const previous = select.value, query = search.value.trim().toLowerCase(); select.replaceChildren();
			const commands = snapshot.commands.filter(command => !['edit_undo', 'edit_redo', 'select.page'].includes(command.id) && (!query || (command.label + ' ' + command.id + ' ' + commandCategory(command.id)).toLowerCase().includes(query)));
			for (const category of [...new Set(commands.map(command => commandCategory(command.id)))].sort()) {
				const group = document.createElement('optgroup'); group.label = humanize(category);
				for (const command of commands.filter(command => commandCategory(command.id) === category)) { group.append(new Option(command.label, command.id)); } select.append(group);
			}
			if (commands.some(command => command.id === previous)) { select.value = previous; } render();
		};
		search.addEventListener('input', filter); select.addEventListener('change', render); filter();
		form.append(title, search, select, description, fields, error, footer);
		form.addEventListener('submit', event => {
			event.preventDefault(); error.hidden = true;
			try {
				if (!definition || !root) { return; }
				if (!definition.readOnly && snapshot.revision !== this.snapshot().revision) { throw new Error('The document changed. Reopen this tool before applying.'); }
				this.execute(definition.id, root.read() as Record<string, unknown>, snapshot.revision); dialog.close();
			} catch (reason) { error.textContent = reason instanceof Error ? reason.message : String(reason); error.hidden = false; }
		});
		dialog.append(form); document.body.append(dialog); dialog.showModal(); search.focus();
	}
	private field(schema: Schema, name: string, required: boolean, total: number, depth: number): Field {
		if (depth > 8) { throw new Error('This tool exceeds the supported visual nesting depth. Use the same operation through the document agent.'); }
		if (!schema.type && !schema.properties && !schema.enum && !schema.oneOf && !schema.anyOf) { schema = { ...schema, oneOf: [{ type: 'string', title: 'Text' }, { type: 'number', title: 'Number' }, { type: 'boolean', title: 'Yes / No' }, { type: 'object', title: 'Properties', additionalProperties: true }, { type: 'array', title: 'Text List', items: { type: 'string' } }, { type: 'array', title: 'Number List', items: { type: 'number' } }] }; }
		if (schema.type === 'object' && schema.properties === undefined && schema.additionalProperties === undefined) { schema = { ...schema, additionalProperties: true }; }

		const container = document.createElement('div'); container.className = 'pdf-schema-field';
		const type = Array.isArray(schema.type) ? schema.type.find(type => type !== 'null') : schema.type;
		if (type === 'null') { container.textContent = 'No value'; return { element: container, read: () => null }; }
		const label = document.createElement('label'); label.textContent = humanize(name) + (required ? ' *' : ''); label.title = schema.description || ''; container.append(label);
		if ((schema.oneOf || schema.anyOf) && !type && !schema.properties) {
			const variants = schema.oneOf || schema.anyOf || [], select = document.createElement('select'), content = document.createElement('div'); let child: Field | undefined;
			variants.forEach((variant, index) => select.add(new Option(variant.title || String(variant.type || 'Option ' + (index + 1)), String(index))));
			const update = () => { child = this.field(variants[Number(select.value)], name, required, total, depth + 1); content.replaceChildren(child.element); }; select.addEventListener('change', update); label.append(select); container.append(content); update(); return { element: container, read: () => child?.read() };
		}
		if (type === 'object' || schema.properties || schema.additionalProperties) {
			const group = document.createElement('fieldset'); group.className = 'pdf-schema-group'; const children: { name: string; field: Field }[] = [];
			let included = required; if (!required) { const toggle = document.createElement('input'); toggle.type = 'checkbox'; toggle.setAttribute('aria-label', 'Set ' + humanize(name)); toggle.addEventListener('change', () => { included = toggle.checked; group.hidden = !included; }); label.append(toggle); group.hidden = true; }
			for (const [key, child] of Object.entries(schema.properties || {})) { const field = this.field(child, key, !!schema.required?.includes(key), total, depth + 1); group.append(field.element); children.push({ name: key, field }); }
			const entries: { key: HTMLInputElement; value: Field; row: HTMLElement }[] = [];
			if (schema.additionalProperties) {
				const add = document.createElement('button'); add.type = 'button'; add.textContent = 'Add Entry'; add.addEventListener('click', () => { if (entries.length >= 200) { return; } const row = document.createElement('div'); row.className = 'pdf-schema-entry'; const key = document.createElement('input'); key.placeholder = 'Name'; key.setAttribute('aria-label', humanize(name) + ' entry name'); const value = this.field(typeof schema.additionalProperties === 'object' ? schema.additionalProperties : {}, 'Value', true, total, depth + 1); const remove = document.createElement('button'); remove.type = 'button'; remove.textContent = 'Remove'; const entry = { key, value, row }; remove.addEventListener('click', () => { entries.splice(entries.indexOf(entry), 1); row.remove(); }); row.append(key, value.element, remove); add.before(row); entries.push(entry); key.focus(); }); group.append(add);
			}
			container.append(group); return { element: container, read: () => {
				if (!included) { return undefined; } const result: Record<string, unknown> = {};
				for (const child of children) { const value = child.field.read(); if (value !== undefined) { result[child.name] = value; } }
				for (const entry of entries) { const key = entry.key.value.trim(); if (!key || ['__proto__', 'constructor', 'prototype'].includes(key) || Object.hasOwn(result, key)) { throw new Error('Use a unique entry name.'); } result[key] = entry.value.read(); } return result;
			} };
		}
		if (type === 'array' && (schema.items?.type === 'object' || schema.items?.type === 'array' || schema.items?.properties)) {
			const list = document.createElement('div'), entries: Field[] = [], add = document.createElement('button'); add.type = 'button'; add.textContent = 'Add Item'; add.addEventListener('click', () => { if (entries.length >= Math.min(200, schema.maxItems || 200)) { return; } const field = this.field(schema.items || {}, 'Item ' + (entries.length + 1), true, total, depth + 1), row = document.createElement('div'), remove = document.createElement('button'); remove.type = 'button'; remove.textContent = 'Remove'; remove.addEventListener('click', () => { entries.splice(entries.indexOf(field), 1); row.remove(); }); row.append(field.element, remove); list.append(row); entries.push(field); }); container.append(list, add); return { element: container, read: () => { if (!entries.length && !required) { return undefined; } if (entries.length < (schema.minItems || 0)) { throw new Error(`Add at least ${schema.minItems} items to ${name}.`); } return entries.map(entry => entry.read()); } };
		}
		if (schema.enum || type === 'boolean') {
			const select = document.createElement('select'); if (!required) { select.add(new Option('Unchanged / Not Set', '')); }
			const choices = schema.enum || [true, false]; choices.forEach((value, index) => select.add(new Option(String(value), String(index)))); label.append(select);
			return { element: container, read: () => select.value === '' ? undefined : choices[Number(select.value)] };
		}
		const input = document.createElement(type === 'array' || /text|script|contents/i.test(name) ? 'textarea' : 'input'); let touched = false;
		input.setAttribute('aria-label', humanize(name)); input.addEventListener('input', () => { touched = true; });
		if (input instanceof HTMLInputElement && /password/i.test(name)) { input.type = 'password'; input.autocomplete = 'off'; }
		if (input instanceof HTMLInputElement && ['integer', 'number'].includes(type || '')) { input.type = 'number'; input.step = type === 'integer' ? '1' : 'any'; if (schema.minimum !== undefined) { input.min = String(schema.minimum); } if (schema.maximum !== undefined) { input.max = String(schema.maximum); } }
		if (schema.maxLength) { input.maxLength = Math.min(schema.maxLength, 100000); }
		input.placeholder = type === 'array' ? (name === 'pages' ? '1, 3–5, or all' : schema.items?.type === 'string' ? 'One value per line' : 'Values separated by commas') : schema.description?.includes('URI') ? 'file:///…' : '';
		label.append(input);
		if (this.pickFile && schema.description?.includes('URI') && type === 'string' && !schema.description.startsWith('Optional output')) { const choose = document.createElement('button'); choose.type = 'button'; choose.textContent = 'Choose File…'; choose.addEventListener('click', () => { void this.pickFile!(['folder', 'directories'].includes(name)).then(uri => { if (uri) { input.value = uri; touched = true; } }); }); container.append(choose); }
		return { element: container, read: () => {
			const value = input.value; if (!value && !required && !touched) { return undefined; }
			if (type === 'array') {
				if (name === 'pages' && schema.items?.type === 'integer') { return pdfPageRange(value, total); }
				const values = (schema.items?.type === 'string' || Array.isArray(schema.items?.type) && schema.items.type.includes('string')) ? value.split(/\r?\n/).filter(Boolean).map(item => item === 'null' && Array.isArray(schema.items?.type) && schema.items.type.includes('null') ? null : item) : value.split(/[,\s]+/).filter(Boolean).map(item => { const number = Number(item); if (!Number.isFinite(number)) { throw new Error(`${name} requires numeric values.`); } return number; });
				if (values.length < (schema.minItems || 0) || values.length > (schema.maxItems || 10000)) { throw new Error(`Check the number of items in ${name}.`); } return values;
			}
			if (['integer', 'number'].includes(type || '')) { const number = Number(value); if (!value || !Number.isFinite(number) || type === 'integer' && !Number.isInteger(number) || schema.minimum !== undefined && number < schema.minimum || schema.maximum !== undefined && number > schema.maximum) { throw new Error(`Enter a valid ${name}.`); } return number; }
			if (!schema.type && !schema.properties && /^-?\d+(?:\.\d+)?$/.test(value)) { return Number(value); }
			if (value.length < (schema.minLength || 0)) { throw new Error(`${name} requires at least ${schema.minLength} characters.`); } return value;
		} };
	}
}
