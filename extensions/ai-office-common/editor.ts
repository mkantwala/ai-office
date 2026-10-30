/*---------------------------------------------------------------------------------------------
 *  Copyright (c) Microsoft Corporation. All rights reserved.
 *  Licensed under the MIT License. See License.txt in the project root for license information.
 *--------------------------------------------------------------------------------------------*/

import DOMPurify from "../ai-office-slides/node_modules/dompurify/dist/purify.es.mjs";
import type { Kind, State, Block, Command, PdfObject } from "./types.js";
import { arrangeWordToolbar } from "./wordToolbar.js";
import { renderAgentActivity, WordAgentReveal } from "./agentOverlay.js";
import type {
	DocumentToolbarAction,
	DocumentToolbarControl,
} from "../../src/vs/base/browser/ui/documentToolbar/documentToolbar.js";
import { renderPdfInspection } from "./pdfInspection.js";
import { PdfCommandForm } from "./pdfCommandForm.js";
import { PdfTools } from "./pdfTools.js";
import { PdfCanvas } from "./pdfCanvas.js";
import { WordImage } from "./wordImage.js";
import { WordTable } from "./wordTable.js";
import { WordParagraph } from "./wordParagraph.js";
import { WordPageSetup } from "./wordPageSetup.js";
import { WordReview } from "./wordReview.js";
import { wordCursor } from "./wordCursor.js";
declare function acquireVsCodeApi(): { postMessage(message: unknown): void };
const api = acquireVsCodeApi();
const el = (id: string) => document.getElementById(id)!;
const pages = el("pages");
const agentReveal = new WordAgentReveal(pages, place, () => state.revision);
pages.style.setProperty("--word-text-cursor", wordCursor("text"));
pages.style.setProperty("--word-corner-cursor", wordCursor("corner"));
const wordReview = new WordReview(
	pages,
	(data, action, ids) =>
		send("reviewDecide", {
			version: data.version,
			reviewVersion: data.reviewVersion,
			action,
			ids,
		}),
	() => {
		wordReview.layout(dimensions);
		overlay();
		agentOverlay();
	},
	(id) => send("draftDiscard", { id }),
);
const reviewSummary = document.createElement("div");
reviewSummary.className = "word-review-summary";
reviewSummary.hidden = true;
pages.before(reviewSummary);
function showReviewSummary() {
	reviewSummary.replaceChildren();
	reviewSummary.hidden = !wordReview.count;
	if (!wordReview.count) {
		return;
	}
	const text = document.createElement("span");
	text.textContent = `${wordReview.count} proposed ${wordReview.count === 1 ? "change" : "changes"} · Before Agent → Current`;
	reviewSummary.append(text);
	for (const action of ["keep", "undo"] as const) {
		const button = document.createElement("button");
		button.textContent = action === "keep" ? "Keep All" : "Undo All";
		button.setAttribute(
			"aria-label",
			`${button.textContent} Changes in This Document`,
		);
		button.addEventListener("click", () => {
			wordReview.decideAll(action);
		});
		reviewSummary.append(button);
	}
}
const toolbar = el("toolbar");
const appToolbar = document.body.classList.contains("app-toolbar-hosted");
let toolbarControls: Record<string, HTMLButtonElement | HTMLSelectElement> = {};
let toolbarContext = "";
function publishToolbar() {
	if (!appToolbar || !initialized) {
		return;
	}
	toolbarContext = JSON.stringify([
		state.revision,
		state.selection,
		selected?.id,
		kind === "pdf" ? selectedPdfObject?.kind : undefined,
		state.page,
	]);
	const controls: DocumentToolbarControl[] = Object.entries(
		toolbarControls,
	).map(([key, control]) => ({
		key,
		...(["font", "size"].includes(key)
			? {
					currentValue: String(
						kind === "pdf"
							? ((key === "font" ? selected?.font : selected?.size) ?? "")
							: (state.formatting?.[key as "font" | "size"] ?? ""),
					),
				}
			: {}),
		type: control instanceof HTMLSelectElement ? "select" : "button",
		label:
			control.getAttribute("aria-label") ||
			control.title ||
			control.textContent ||
			key,
		...(control instanceof HTMLButtonElement
			? { text: control.textContent || "" }
			: {}),
		disabled: control.disabled,
		...(control.hasAttribute("aria-pressed")
			? { pressed: control.getAttribute("aria-pressed") === "true" }
			: {}),
		...(control instanceof HTMLSelectElement
			? {
					options: Array.from(control.options).map((option) => ({
						value: option.value,
						label: option.textContent || "",
					})),
				}
			: {}),
	}));
	send("toolbarState", {
		controls,
		context: toolbarContext,
		paragraphStyle: state.paragraphStyle,
		painter: !!state.painter?.active,
	});
}
function toolbarAction(
	action: DocumentToolbarAction,
	context: string,
	instance?: string,
) {
	if (
		(document.body.dataset.documentInstance &&
			instance !== document.body.dataset.documentInstance) ||
		context !== toolbarContext
	) {
		publishToolbar();
		message(
			"The document selection changed. Choose the formatting action again.",
		);
		return;
	}
	if (
		action.control === "table" &&
		kind === "word" &&
		Number.isInteger(action.rows) &&
		Number.isInteger(action.cols) &&
		action.rows! >= 1 &&
		action.rows! <= 20 &&
		action.cols! >= 1 &&
		action.cols! <= 10
	) {
		execute("insert.table", {
			rows: action.rows,
			cols: action.cols,
			borders: true,
		});
		focusWordInput();
		return;
	}
	if (
		action.control === "size" &&
		(kind === "word" || kind === "pdf") &&
		typeof action.value === "string" &&
		/^\d+(?:\.\d+)?$/.test(action.value) &&
		Number(action.value) >= 1 &&
		Number(action.value) <= 400 &&
		!toolbarControls.size?.disabled
	) {
		if (kind === "pdf") {
			pdfCanvas.format({ size: Number(action.value) });
		} else {
			execute("format.size", { size: Number(action.value) });
		}
		focusWordInput();
		return;
	}
	if (action.control === "paintFormat" && kind === "word") {
		execute(
			"edit.formatPainter",
			action.value === "sticky" ? { sticky: true } : {},
		);
		focusWordInput();
		return;
	}
	if (
		action.control === "color" &&
		(kind === "word" || kind === "pdf") &&
		typeof action.value === "string" &&
		/^[0-9a-f]{6}$/i.test(action.value) &&
		!toolbarControls.color?.disabled
	) {
		if (kind === "pdf") {
			pdfCanvas.format({ color: "#" + action.value });
		} else {
			execute("format.color", { color: action.value });
		}
		focusWordInput();
		return;
	}
	const control = toolbarControls[action.control];
	if (!control || control.disabled) {
		return;
	}
	if (control instanceof HTMLSelectElement) {
		if (
			!Array.from(control.options).some(
				(option) => option.value === action.value && option.value,
			)
		) {
			return;
		}
		control.value = action.value!;
		control.dispatchEvent(new Event("change"));
	} else {
		control.click();
	}
	focusWordInput();
}
const comments = document.createElement("details");
comments.className = "comments";
comments.hidden = true;
pages.before(comments);
const keyboard = el("keyboard") as HTMLTextAreaElement;
function focusWordInput() {
	if (kind !== "word") {
		return;
	}
	const dialog = document.querySelector<HTMLDialogElement>("dialog[open]");
	if (dialog) {
		if (!dialog.contains(document.activeElement)) {
			dialog
				.querySelector<HTMLElement>("input, select, textarea, button")
				?.focus({ preventScroll: true });
		}
		return;
	}
	// An open Find bar keeps its input's focus; other focus requests go to the document.
	if (!findBar.hidden && findBar.contains(document.activeElement)) {
		return;
	}
	if (
		!document.querySelector(
			".word-table-context:popover-open,.word-image-menu:popover-open",
		)
	) {
		keyboard.focus({ preventScroll: true });
	}
}
// Document Find & Replace over the engine's inbuilt search commands (`edit.find`,
// `edit.findNext`/`findPrevious`, `edit.replace`, `edit.replaceAll`); matches become the
// selection the engine already renders.
const findBar = document.createElement("div");
findBar.className = "find-bar";
findBar.hidden = true;
const findInput = document.createElement("input");
findInput.type = "text";
findInput.placeholder = "Find in document";
findInput.setAttribute("aria-label", "Find in document");
const findCount = document.createElement("span");
findCount.className = "find-count";
findCount.setAttribute("aria-live", "polite");
const findPrevious = document.createElement("button");
findPrevious.type = "button";
findPrevious.textContent = "↑";
findPrevious.title = "Previous Match (Shift+Enter)";
findPrevious.setAttribute("aria-label", "Previous Match");
const findNextButton = document.createElement("button");
findNextButton.type = "button";
findNextButton.textContent = "↓";
findNextButton.title = "Next Match (Enter)";
findNextButton.setAttribute("aria-label", "Next Match");
const findClose = document.createElement("button");
findClose.type = "button";
findClose.textContent = "✕"; // allow-any-unicode-next-line
findClose.title = "Close Find and Replace (Escape)";
findClose.setAttribute("aria-label", "Close Find and Replace");
const replaceInput = document.createElement("input");
replaceInput.type = "text";
replaceInput.placeholder = "Replace with";
replaceInput.setAttribute("aria-label", "Replace with");
const replaceOneButton = document.createElement("button");
replaceOneButton.type = "button";
replaceOneButton.textContent = "Replace";
replaceOneButton.title = "Replace This Match";
replaceOneButton.setAttribute("aria-label", "Replace This Match");
const replaceAllButton = document.createElement("button");
replaceAllButton.type = "button";
replaceAllButton.textContent = "Replace All";
replaceAllButton.title = "Replace All Matches";
replaceAllButton.setAttribute("aria-label", "Replace All Matches");
const findRow = document.createElement("div");
findRow.className = "find-row";
const replaceRow = document.createElement("div");
replaceRow.className = "find-row";
findRow.append(findInput, findCount, findPrevious, findNextButton, findClose);
replaceRow.append(replaceInput, replaceOneButton, replaceAllButton);
findBar.append(findRow, replaceRow);
document.body.append(findBar);
let findQuery = "";
let findTimer: ReturnType<typeof setTimeout> | undefined;
function openFind() {
	if (kind !== "word") {
		return;
	}
	findBar.hidden = false;
	findInput.focus();
	findInput.select();
	if (findQuery) {
		execute("edit.find", { text: findQuery });
	}
}
function closeFind() {
	clearTimeout(findTimer);
	findTimer = undefined;
	if (findBar.hidden) {
		return;
	}
	findBar.hidden = true;
	focusWordInput();
}
function stepFind(command: string) {
	if (!findQuery) {
		return;
	}
	// An empty result set: the engine's step reports "no matches" as an error, so
	// show the count instead of raising it over the document.
	if (!state.find || !state.find.count) {
		findCount.textContent = "No matches";
		return;
	}
	execute(command);
}
findInput.addEventListener("input", () => {
	findQuery = findInput.value;
	clearTimeout(findTimer);
	findTimer = setTimeout(() => {
		findTimer = undefined;
		execute("edit.find", { text: findQuery });
	}, 250);
});
findInput.addEventListener("keydown", (event) => {
	if (event.key === "Enter") {
		event.preventDefault();
		event.stopPropagation();
		stepFind(event.shiftKey ? "edit.findPrevious" : "edit.findNext");
	} else if (event.key === "Escape") {
		event.preventDefault();
		event.stopPropagation();
		closeFind();
	}
});
findPrevious.addEventListener("click", () => stepFind("edit.findPrevious"));
findNextButton.addEventListener("click", () => stepFind("edit.findNext"));
findClose.addEventListener("click", closeFind);
replaceOneButton.addEventListener("click", () => {
	if (findQuery) {
		execute("edit.replace", { text: findQuery, with: replaceInput.value });
	}
});
replaceAllButton.addEventListener("click", () => {
	if (findQuery) {
		execute("edit.replaceAll", { text: findQuery, with: replaceInput.value });
	}
});
const wordTable = new WordTable(
	pages,
	point,
	place,
	execute,
	ask,
	focusWordInput,
	(page) =>
		page.getBoundingClientRect().width /
		dimensions[Number(page.dataset.page)][0],
	(request, args) => send("tablePreview", { request, args }),
	(at) =>
		!!state.rects?.some(
			(r) =>
				r.page === at.page &&
				at.x >= r.x &&
				at.x <= r.x + r.width &&
				at.y >= r.y &&
				at.y <= r.y + r.height,
		),
);
const wordImage = new WordImage(
	pages,
	point,
	place,
	execute,
	(target) => send("image", { replace: true, target }),
	focusWordInput,
	(page) => page.clientWidth / dimensions[Number(page.dataset.page)][0],
);

let documentCommands: Command[] = [];
const resourcePickers = new Map<string, (uri: string | undefined) => void>();
const pdfCommandForm = new PdfCommandForm(
	() => ({
		commands: documentCommands,
		total: dimensions.length,
		revision: state.revision,
	}),
	(command, args, revision) =>
		send("execute", { command, args, revision, inspect: true }),
	(folders) =>
		new Promise((resolve) => {
			const request = String(Date.now()) + ":" + resourcePickers.size;
			resourcePickers.set(request, resolve);
			send("pdfPickResource", { request, folders });
		}),
);
const pdfTools = new PdfTools(
	() => ({
		total: dimensions.length,
		page: currentPage,
		revision: state.revision,
	}),
	(command, args, revision) => send("execute", { command, args, revision }),
	() => pdfCommandForm.open(),
);
let selectedPdfObject: PdfObject | undefined;
const pdfCanvas = new PdfCanvas(pages, {
	snapshot: () => ({ state, dimensions }),
	execute: (command, args, revision) =>
		send("execute", { command, args, revision }),
	form: (command, initial) => pdfCommandForm.open(command, initial),
	select: (block, object) => {
		selected = block;
		selectedPdfObject = object;
		publishToolbar();
	},
	page: (page) => {
		currentPage = page;
		execute("select.page", { page: page + 1 });
	},
	message: (text) => {
		if (text) {
			el("hint").textContent = text;
		} else {
			defaultHint();
		}
	},
	organize: () => pdfTools.open(),
});
const pdfInspection = document.createElement("details");
pdfInspection.className = "pdf-inspection";
pdfInspection.hidden = true;
pages.before(pdfInspection);
const pdfArtifacts = document.createElement("details");
pdfArtifacts.className = "pdf-artifacts";
pdfArtifacts.hidden = true;
pages.before(pdfArtifacts);

let wordParagraph: WordParagraph | undefined;
let kind: Kind;
let state: State = { revision: -1 };
let dimensions: number[][] = [];
let selected: Block | undefined;
let currentPage = 0;
let activeEditor: HTMLTextAreaElement | undefined;
let pendingText = false;
let renderId = 0;
const visible = new Set<number>();
const rendered = new Set<number>();
const pendingRender = new Map<
	number,
	{ revision: number; page: number; region?: State["editingRegion"] }
>();
let renderBusy = false;
let composing = false;
let drag = false;
let pointerFrame = 0;
let lastPointer: { page: number; x: number; y: number } | undefined;
let initialized = false;
let wordPageSetup: WordPageSetup | undefined;
let disposeToolbar: { dispose(): void } | undefined;
let agentLabel = "Agent";
let agentIntent: string | undefined;
let followRequest: string | undefined;
let followAgent = true;
function followAgentChange(request: string | undefined, selector: string) {
	if (request && request !== followRequest) {
		followRequest = request;
		followAgent = true;
	}
	if (!followAgent) {
		return;
	}
	const target = pages.querySelector<HTMLElement>(selector);
	if (!target) {
		return;
	}
	const bounds = target.getBoundingClientRect();
	if (bounds.top < 150 || bounds.bottom > innerHeight - 40) {
		target.scrollIntoView({ block: "center", behavior: "instant" });
	}
}
let activityTimer: ReturnType<typeof setTimeout> | undefined;
let agentRects: NonNullable<State["rects"]> = [];
let legacyAgentRects: NonNullable<State["rects"]> = [];
const agentStatus = document.createElement("span");
agentStatus.className = "agent-status";
agentStatus.setAttribute("role", "status");
agentStatus.setAttribute("aria-live", "polite");
el("info").append(agentStatus);
function agentOverlay() {
	const active = renderAgentActivity(
		pages,
		agentRects.length ? agentRects : legacyAgentRects,
		place,
		agentLabel,
		agentRects.length ? agentIntent : undefined,
	);
	agentStatus.textContent = active
		? agentLabel + ": " + (agentIntent || "Working on this content")
		: "";
	pages.setAttribute("aria-busy", String(active));
}

let wordText = "";
let wordTextTimer: ReturnType<typeof setTimeout> | undefined;
let revealWordCaret = false;
function flushWordInput() {
	clearTimeout(wordTextTimer);
	wordTextTimer = undefined;
	if (!wordText) {
		return;
	}
	const text = wordText;
	wordText = "";
	api.postMessage({ type: "execute", command: "text.insert", args: { text } });
}
function typeWordText(text: string) {
	wordText += text;
	if (!wordTextTimer) {
		wordTextTimer = setTimeout(flushWordInput, 0);
	}
}
window.addEventListener(
	"pagehide",
	() => {
		flushWordInput();
		disposeToolbar?.dispose();
		wordParagraph?.dispose();
		wordPageSetup?.dispose();
		wordReview.dispose();
		wordTable.dispose();
		wordImage.dispose();
		pdfCanvas.dispose();
		pdfTools.dispose();
		pdfCommandForm.dispose();
		for (const resolve of resourcePickers.values()) {
			resolve(undefined);
		}
		resourcePickers.clear();
		clearTimeout(activityTimer);
		agentReveal.clear();
		clearAgent();
	},
	{ once: true },
);
window.addEventListener("blur", flushWordInput);
function clearActivityPreview() {
	if (
		!agentRects.length &&
		agentIntent === undefined &&
		activityTimer === undefined
	) {
		return;
	}
	clearTimeout(activityTimer);
	activityTimer = undefined;
	agentIntent = undefined;
	agentRects = [];
	agentOverlay();
	wordReview.updateActivity([]);
	wordReview.layout(dimensions);
	overlay();
}
window.addEventListener("pointerdown", () => {
	followAgent = false;
	agentReveal.clear();
});
window.addEventListener(
	"wheel",
	() => {
		followAgent = false;
	},
	{ passive: true },
);
window.addEventListener("focus", () => {
	if (
		initialized &&
		(document.activeElement === document.body ||
			document.activeElement === keyboard)
	) {
		focusWordInput();
	}
});
function send(type: string, data: Record<string, unknown> = {}) {
	if (type !== "render" && type !== "ready") {
		flushWordInput();
	}
	api.postMessage({ type, ...data });
}
function execute(command: string, args: Record<string, unknown> = {}) {
	el("error").hidden = true;
	send("execute", { command, args });
}
function message(error: unknown) {
	el("error").hidden = false;
	el("error").textContent = String(error);
}
function selectionArgs() {
	return kind === "pdf"
		? {
				page: currentPage + 1,
				...(selectedPdfObject?.kind === "added"
					? { index: selectedPdfObject.index }
					: { paragraph: selected?.id }),
			}
		: {
				page: currentPage + 1,
				shape: selected?.id,
				row: selected?.row,
				col: selected?.col,
				...(activeEditor &&
				activeEditor.selectionEnd > activeEditor.selectionStart
					? {
							range: {
								start: activeEditor.selectionStart,
								end: activeEditor.selectionEnd,
							},
						}
					: {}),
			};
}
function flushText() {
	if (kind === "pdf") {
		pdfCanvas.flush();
		return;
	}
	if (!pendingText || !activeEditor || !selected) {
		return;
	}
	const text = activeEditor.value;
	pendingText = false;
	execute("text.replace", { ...selectionArgs(), text });
	selected.text = text;
}
function closeText() {
	flushText();
	activeEditor?.remove();
	activeEditor = undefined;
}
async function ask(label: string, initial = ""): Promise<string | undefined> {
	const dialog = el("dialog") as HTMLDialogElement;
	const input = el("prompt") as HTMLInputElement;
	if (dialog.open) {
		input.focus();
		return undefined;
	}
	el("prompt-label").textContent = label;
	input.value = initial;
	dialog.returnValue = "cancel";
	dialog.showModal();
	input.focus();
	input.select();
	return new Promise((resolve) => {
		const controller = new AbortController();
		const signal = controller.signal;
		const finish = () => {
			controller.abort();
			focusWordInput();
			resolve(dialog.returnValue === "ok" ? input.value : undefined);
		};
		dialog.querySelector("form")!.addEventListener(
			"submit",
			(event) => {
				event.preventDefault();
				dialog.close(
					(event.submitter as HTMLButtonElement | null)?.value === "ok"
						? "ok"
						: "cancel",
				);
				finish();
			},
			{ signal },
		);
		dialog.addEventListener(
			"cancel",
			(event) => {
				event.preventDefault();
				dialog.close("cancel");
				finish();
			},
			{ signal },
		);
		dialog.addEventListener("close", finish, { signal, once: true });
	});
}
function button(
	label: string,
	title: string,
	action: () => void,
	disabled = false,
) {
	const b = document.createElement("button");
	b.textContent = label;
	b.title = title;
	b.setAttribute("aria-label", title);
	b.disabled = disabled;
	b.addEventListener("pointerdown", (event) => event.preventDefault());
	b.addEventListener("click", () => {
		action();
		focusWordInput();
	});
	toolbar.append(b);
	return b;
}
function menu(
	label: string,
	items: [string, string][],
	action: (value: string) => void,
	disabled = false,
) {
	const select = document.createElement("select");
	select.title = label;
	select.setAttribute("aria-label", label);
	select.disabled = disabled;
	const placeholder = document.createElement("option");
	placeholder.textContent = label;
	placeholder.value = "";
	select.append(placeholder);
	for (const [value, text] of items) {
		const option = document.createElement("option");
		option.value = value;
		option.textContent = text;
		select.append(option);
	}
	select.addEventListener("change", () => {
		if (!select.value) {
			return;
		}
		action(select.value);
		select.value = "";
		if (kind === "word") {
			keyboard.focus({ preventScroll: true });
		}
	});
	toolbar.append(select);
	return select;
}
function needsBox() {
	if (kind !== "word" && !selected) {
		message("Click a text box or paragraph first.");
		return true;
	}
	return false;
}
function format(property: string, value: unknown) {
	if (kind === "pdf") {
		pdfCanvas.format({ [property]: value });
		return;
	}
	if (needsBox()) {
		return;
	}
	flushText();
	if (kind === "word") {
		execute("format." + property, {
			[property === "font" ? "name" : property === "size" ? "size" : "value"]:
				value,
		});
		keyboard.focus({ preventScroll: true });
	} else {
		execute("format.text", { ...selectionArgs(), [property]: value });
	}
}
function buildToolbar() {
	disposeToolbar?.dispose();
	toolbar.replaceChildren();
	const text = button("Text", "Insert text", () => {
		if (kind === "word") {
			keyboard.focus();
			return;
		}
		if (kind === "pdf") {
			pdfCanvas.setMode("text");
			return;
		}
		void (async () => {
			const text = await ask("New text");
			if (!text) {
				return;
			}
			execute("insert.text", { page: currentPage + 1, text });
		})();
	});
	const style = menu(
		"Normal text",
		[
			["Normal", "Normal text"],
			...(kind === "word"
				? ([
						["Title", "Title"],
						["Subtitle", "Subtitle"],
					] as [string, string][])
				: []),
			["Heading1", "Heading 1"],
			["Heading2", "Heading 2"],
			["Heading3", "Heading 3"],
		],
		(style) => execute("para.style", { style }),
		kind !== "word",
	);
	style.id = "paragraph-style";
	const font = menu(
		"Font",
		(kind === "pdf"
			? [
					["helvetica", "Sans"],
					["times", "Serif"],
					["courier", "Mono"],
				]
			: [
					"Aptos",
					"Arial",
					"Calibri",
					"Cambria",
					"Georgia",
					"Times New Roman",
					"Verdana",
					"Courier New",
				].map((font) => [font, font])) as [string, string][],
		(font) => format("font", font),
	);
	const size = menu(
		"Size",
		[
			"8",
			"9",
			"10",
			"11",
			"12",
			"14",
			"16",
			"18",
			"20",
			"22",
			"24",
			"26",
			"28",
			"32",
			"36",
			"48",
			"72",
		].map((x) => [x, x]),
		(size) => format("size", Number(size)),
	);
	const emphasis: HTMLButtonElement[] = [];
	for (const [label, property] of [
		["B", "bold"],
		["I", "italic"],
		["U", "underline"],
	]) {
		const b = button(label, property, () => {
			if (kind === "word") {
				execute("format." + property);
				keyboard.focus({ preventScroll: true });
				return;
			}
			const value = b.getAttribute("aria-pressed") !== "true";
			format(property, value);
			b.setAttribute("aria-pressed", String(value));
		});
		emphasis.push(b);
		b.style.fontWeight = property === "bold" ? "700" : "";
		b.style.fontStyle = property === "italic" ? "italic" : "";
		b.style.textDecoration = property === "underline" ? "underline" : "";
	}
	const strike =
		kind === "word"
			? button("S", "Strikethrough", () => execute("format.strikethrough"))
			: undefined;
	const clear =
		kind === "word"
			? button("Clear", "Clear Formatting", () => execute("format.clear"))
			: undefined;
	const color =
		kind !== "slides"
			? menu(
					"Text Color",
					[
						...(kind === "word"
							? ([["auto", "Automatic"]] as [string, string][])
							: []),
						["000000", "Black"],
						["FFFFFF", "White"],
						["C00000", "Red"],
						["E97132", "Orange"],
						["FFFF00", "Yellow"],
						["008000", "Green"],
						["0070C0", "Blue"],
						["7030A0", "Purple"],
						["808080", "Gray"],
					],
					(color) =>
						kind === "pdf"
							? pdfCanvas.format({ color: "#" + color })
							: execute("format.color", { color }),
				)
			: undefined;
	const highlight =
		kind === "word"
			? menu(
					"Text Highlight",
					[
						["none", "No highlight"],
						["yellow", "Yellow"],
						["brightGreen", "Bright green"],
						["turquoise", "Turquoise"],
						["pink", "Pink"],
						["blue", "Blue"],
						["red", "Red"],
						["gray25", "Light gray"],
						["gray50", "Gray"],
					],
					(color) => execute("format.highlight", { color }),
				)
			: undefined;
	const align = menu(
		"Align",
		["left", "center", "right", "justify"].map((x) => [x, x]),
		(value) => {
			if (kind === "pdf") {
				pdfCanvas.format({ align: value });
				return;
			}
			if (needsBox()) {
				return;
			}
			flushText();
			execute(
				"para.align",
				kind === "word" ? { value } : { ...selectionArgs(), value },
			);
		},
	);
	const list = menu(
		"List",
		[
			["bullet", "Bullets"],
			["number", "Numbered"],
			["none", "No list"],
		],
		(value) => {
			if (needsBox()) {
				return;
			}
			flushText();
			execute(
				kind === "word"
					? value === "number"
						? "para.numbering"
						: "para.bullets"
					: "para.list",
				kind === "word"
					? { off: value === "none" }
					: { ...selectionArgs(), value },
			);
		},
		kind === "pdf",
	);
	const spacing = menu(
		"Spacing",
		[
			["1", "Single"],
			["1.5", "1.5 lines"],
			["2", "Double"],
		],
		(value) => {
			if (kind === "pdf") {
				pdfCanvas.format({ line_spacing: Number(value) });
				return;
			}
			if (needsBox()) {
				return;
			}
			flushText();
			execute(
				kind === "word" ? "para.lineSpacing" : "para.spacing",
				kind === "word"
					? { value: Number(value) }
					: { ...selectionArgs(), value: Number(value) },
			);
		},
	);
	const divider = document.createElement("span");
	divider.className = "divider";
	toolbar.append(divider);
	const table = button(
		"Table",
		kind === "pdf" ? "PDF table insertion is not available" : "Insert table",
		() => {
			if (kind === "word") {
				return;
			}
			void (async () => {
				const answer = await ask("Table size: rows x columns", "3 x 3");
				if (!answer) {
					return;
				}
				const match = /^(\d+)\s*[xX×]\s*(\d+)$/.exec(answer);
				if (
					!match ||
					Number(match[1]) > 20 ||
					Number(match[2]) > 10 ||
					Number(match[1]) < 1 ||
					Number(match[2]) < 1
				) {
					message("Use 1-20 rows and 1-10 columns.");
					return;
				}
				flushText();
				execute("insert.table", {
					rows: Number(match[1]),
					cols: Number(match[2]),
					page: currentPage + 1,
				});
			})();
		},
		kind === "pdf",
	);
	const image = button("Image", "Insert image", () => {
		if (kind === "pdf") {
			pdfCanvas.setMode("image");
			return;
		}
		flushText();
		send("image", { page: currentPage + 1 });
	});
	const link = button("Link", "Insert link", () => {
		if (kind === "pdf") {
			pdfCanvas.setMode("link");
			return;
		}
		void (async () => {
			if (needsBox()) {
				return;
			}
			const url = await ask("Link address", "https://");
			if (!url) {
				return;
			}
			try {
				if (!["https:", "http:", "mailto:"].includes(new URL(url).protocol)) {
					throw new Error();
				}
			} catch {
				message("Enter an http, https or mailto address.");
				return;
			}
			flushText();
			execute(
				"insert.link",
				kind === "word" ? { url } : { ...selectionArgs(), url },
			);
		})();
	});
	const comment = button("Comment", "Add comment", () => {
		if (kind === "pdf") {
			pdfCanvas.setMode("note");
			return;
		}
		void (async () => {
			const text = await ask("Comment");
			if (!text?.trim()) {
				return;
			}
			flushText();
			execute(
				kind === "word" ? "review.newComment" : "insert.comment",
				kind === "word" ? { text } : { page: currentPage + 1, text },
			);
		})();
	});
	const spacer = document.createElement("span");
	spacer.className = "spacer";
	toolbar.append(spacer);
	const undo = button("Undo", "Undo (Ctrl+Z)", () => {
		closeText();
		send("undo");
	});
	undo.id = "undo";
	const redo = button("Redo", "Redo (Ctrl+Shift+Z)", () => {
		closeText();
		send("redo");
	});
	redo.id = "redo";
	const save = button("Save", "Save (Ctrl+S)", () => {
		flushText();
		send("save");
	});
	const margins =
		kind === "word"
			? button("Margins", "Page Margins", () => wordPageSetup?.open("margins"))
			: undefined;
	const pageBorders =
		kind === "word"
			? button("Page Borders", "Page Borders", () =>
					wordPageSetup?.open("borders"),
				)
			: undefined;
	const header =
		kind === "word"
			? button("Header", "Edit Header", () =>
					execute("insert.editHeader", { page: currentPage }),
				)
			: undefined;
	const footer =
		kind === "word"
			? button("Footer", "Edit Footer", () =>
					execute("insert.editFooter", { page: currentPage }),
				)
			: undefined;
	const versions =
		kind === "word"
			? button("History", "Document Version History", () => send("versions"))
			: undefined;
	let properties: HTMLButtonElement | undefined =
		kind === "pdf"
			? button(
					"PDF Tools",
					"PDF Tools · Organize Pages, Crop, Watermarks and Page Numbers",
					() => {
						closeText();
						pdfCanvas.openTools();
					},
				)
			: undefined;
	if (kind === "word") {
		document.body.classList.add("word-document");
		wordPageSetup = new WordPageSetup(
			(patch, snapshot) =>
				execute("layout.pageProperties", {
					...patch,
					ifRevision: snapshot.revision,
					ifSelection: snapshot.selection,
				}),
			focusWordInput,
		);
		wordParagraph = new WordParagraph(
			pages,
			(patch, snapshot) => {
				execute("para.properties", {
					...patch,
					ifRevision: snapshot.revision,
					ifSelection: snapshot.selection,
				});
			},
			() => keyboard.focus({ preventScroll: true }),
		);
		properties = button(
			"Paragraph",
			"Paragraph Properties · Indents, Spacing, Borders and Pagination",
			() => wordParagraph?.open(),
		);
		if (!appToolbar) {
			disposeToolbar = arrangeWordToolbar(
				toolbar,
				{
					undo,
					redo,
					save,
					versions,
					properties,
					header,
					footer,
					margins,
					pageBorders,
					style,
					font,
					size,
					emphasis,
					strike,
					clear,
					color,
					highlight,
					align,
					list,
					spacing,
					table,
					image,
					link,
					comment,
				},
				(rows, cols) => execute("insert.table", { rows, cols, borders: true }),
				() => keyboard.focus({ preventScroll: true }),
				(sticky) =>
					execute("edit.formatPainter", sticky ? { sticky: true } : {}),
			);
		}
	}
	toolbarControls = {
		...(margins ? { margins } : {}),
		...(pageBorders ? { pageBorders } : {}),
		...(header ? { header } : {}),
		...(footer ? { footer } : {}),
		...(strike ? { strike } : {}),
		...(clear ? { clear } : {}),
		...(color ? { color } : {}),
		...(highlight ? { highlight } : {}),
		text,
		style,
		font,
		size,
		bold: emphasis[0],
		italic: emphasis[1],
		underline: emphasis[2],
		align,
		list,
		spacing,
		table,
		image,
		link,
		comment,
		undo,
		redo,
		save,
		...(properties ? { properties } : {}),
		...(versions ? { versions } : {}),
	};
	if (appToolbar) {
		toolbar.hidden = true;
	}
}
function point(event: MouseEvent, page: HTMLElement) {
	const bounds = page.getBoundingClientRect();
	const index = Number(page.dataset.page);
	const scale = bounds.width / dimensions[index][0];
	return {
		page: index,
		x: (event.clientX - bounds.left) / scale,
		y:
			(kind === "word"
				? wordReview.toDocument(index, event.clientY - bounds.top)
				: event.clientY - bounds.top) / scale,
	};
}
function makePages() {
	observer.disconnect();
	visible.clear();
	rendered.clear();
	pages.replaceChildren();
	for (const [index, size] of dimensions.entries()) {
		const page = document.createElement("section");
		page.className = "page " + kind;
		page.dataset.page = String(index);
		page.style.aspectRatio = String(size[0] / size[1]);
		if (kind === "word") {
			page.style.width = `${(size[0] * 4) / 3}px`;
			const paper = state.pageRegions?.[index];
			if (paper && (size[0] > paper.width || size[1] > paper.height)) {
				page.dataset.overflow = "true";
				page.style.setProperty("--paper-width", `${(paper.width * 4) / 3}px`);
				page.style.setProperty("--paper-height", `${(paper.height * 4) / 3}px`);
			}
		}
		if (kind === "pdf") {
			page.tabIndex = -1;
		}
		page.setAttribute(
			"aria-label",
			(kind === "slides" ? "Slide " : "Page ") + (index + 1),
		);
		const content = document.createElement(
			kind === "slides" ? "img" : "canvas",
		);
		page.append(content);
		pages.append(page);
		page.addEventListener("pointerdown", (event) => {
			if (
				(event.target as HTMLElement).closest(
					".word-review-zone,.word-header-footer-controls",
				)
			) {
				return;
			}
			if (event.button !== 0) {
				return;
			}
			if (
				event.target instanceof HTMLTextAreaElement ||
				event.target instanceof HTMLButtonElement
			) {
				return;
			}
			currentPage = index;
			if (kind === "word") {
				event.preventDefault();
				drag = true;
				page.setPointerCapture(event.pointerId);
				keyboard.focus({ preventScroll: true });
				execute("caret.set", { ...point(event, page), extend: event.shiftKey });
			} else {
				closeText();
				selected = undefined;
				execute(kind === "pdf" ? "select.page" : "select.shape", {
					page: index + 1,
				});
			}
		});
		page.addEventListener("pointermove", (event) => {
			if (!drag || kind !== "word") {
				return;
			}
			lastPointer = point(event, page);
			if (!pointerFrame) {
				pointerFrame = requestAnimationFrame(() => {
					pointerFrame = 0;
					if (lastPointer) {
						execute("caret.set", { ...lastPointer, extend: true });
					}
				});
			}
		});
		page.addEventListener("pointerup", () => {
			if (kind !== "word" || !drag) {
				return;
			}
			drag = false;
			// Upstream's canvas: an armed format painter applies its picked formatting to the
			// selection when the drag ends; a plain click (collapsed caret) does not apply.
			if (state.painter?.active && state.rects?.length) {
				execute("edit.pasteFormat");
			}
		});
		page.addEventListener("dblclick", (event) => {
			if (
				kind !== "word" ||
				event.defaultPrevented ||
				(event.target as HTMLElement).closest(".word-header-footer-controls")
			) {
				return;
			}
			const at = point(event, page),
				region = state.pageRegions?.[index];
			const target =
				region && at.y >= 0 && at.y < region.top
					? "header"
					: region && at.y > region.bottom && at.y <= region.height
						? "footer"
						: "body";
			if (target !== "body") {
				execute(
					target === "header" ? "insert.editHeader" : "insert.editFooter",
					{ page: index },
				);
				execute("caret.set", at);
			} else if (state.editingRegion && state.editingRegion !== "body") {
				execute("insert.closeHeader");
				execute("caret.set", at);
			} else {
				execute("select.word");
			}
		});
		observer.observe(page);
	}
}
const observer = new IntersectionObserver(
	(entries) => {
		for (const entry of entries) {
			const index = Number((entry.target as HTMLElement).dataset.page);
			if (entry.isIntersecting) {
				visible.add(index);
			} else {
				visible.delete(index);
				rendered.delete(index);
				wordTable.forget(index);
				wordImage.forget(index);
				const canvas = entry.target.querySelector("canvas");
				if (canvas) {
					canvas.width = canvas.height = 1;
				}
				entry.target
					.querySelectorAll(".word-review-slice")
					.forEach((node) => node.remove());
			}
		}
		renderNext();
	},
	{ rootMargin: "450px" },
);
function renderNext() {
	if (renderBusy) {
		return;
	}
	const caretPage = kind === "word" ? state.caret?.page : undefined;
	const page =
		caretPage !== undefined &&
		visible.has(caretPage) &&
		!rendered.has(caretPage)
			? caretPage
			: [...visible].find((index) => !rendered.has(index));
	if (page === undefined) {
		return;
	}
	const node = pages.children[page] as HTMLElement;
	const request = ++renderId;
	renderBusy = true;
	pendingRender.set(request, {
		revision: state.revision,
		page,
		region: state.editingRegion,
	});
	const canvas = node.querySelector("canvas");
	const scale =
		(node.clientWidth * Math.min(devicePixelRatio, 2)) / dimensions[page][0];
	// Keep the same physical-pixel resolution while typing; native rendering caches
	// repaint changed text bands without downsampling the page.
	send("render", {
		request,
		page,
		scale,
		...(kind === "word" && canvas && canvas.width > 1
			? { base: Number(canvas.dataset.frame) }
			: {}),
	});
}
function place(
	node: HTMLElement,
	page: number,
	x: number,
	y: number,
	w: number,
	h: number,
) {
	const element = pages.children[page] as HTMLElement;
	if (
		!element ||
		!dimensions[page]?.[0] ||
		![x, y, w, h].every(Number.isFinite)
	) {
		return;
	}
	const scale = element.clientWidth / dimensions[page][0];
	node.style.left = x * scale + "px";
	node.style.top =
		(kind === "word" ? wordReview.toView(page, y * scale) : y * scale) + "px";
	node.style.width = w ? w * scale + "px" : "2px";
	node.style.height = h * scale + "px";
	element.append(node);
}
function overlay() {
	pages
		.querySelectorAll(".selection,.caret,.block,.word-header-footer-controls")
		.forEach((node) => node.remove());

	if (kind === "word") {
		const activePage = state.caret?.page ?? currentPage;
		currentPage = activePage;
		const region = state.pageRegions?.[activePage];
		if (region && state.editingRegion && state.editingRegion !== "body") {
			const controls = document.createElement("div");
			controls.className = "word-header-footer-controls";
			const label = document.createElement("span");
			label.textContent =
				state.editingRegion === "header" ? "Header" : "Footer";
			const close = document.createElement("button");
			close.type = "button";
			close.textContent = "Close Header and Footer";
			close.addEventListener("pointerdown", (event) => {
				event.preventDefault();
				event.stopPropagation();
			});
			close.addEventListener("click", (event) => {
				event.stopPropagation();
				execute("insert.closeHeader");
				focusWordInput();
			});
			controls.append(label, close);
			place(
				controls,
				activePage,
				0,
				state.editingRegion === "header" ? region.top : region.bottom,
				region.width,
				0,
			);
		}
		for (const r of state.rects || []) {
			const node = document.createElement("div");
			node.className = "selection";
			place(node, r.page, r.x, r.y, r.width, r.height);
		}
		if (state.caret && !state.rects?.length) {
			const c = state.caret;
			const node = document.createElement("div");
			node.className = "caret";
			place(node, c.page, c.x, c.y, c.width || 0, c.height || 0);
			if (c.width && !c.height) {
				node.style.height = "1.5px";
			} else {
				node.style.width = "1.5px";
			}
			if (revealWordCaret) {
				node.scrollIntoView({ block: "nearest", inline: "nearest" });
				revealWordCaret = false;
			}
			const box = node.getBoundingClientRect();
			keyboard.style.left = box.x + "px";
			keyboard.style.top = box.y + "px";
		}
	} else if (kind !== "pdf") {
		for (const block of state.blocks || []) {
			const node = document.createElement("button");
			node.className = "block";
			node.setAttribute(
				"aria-label",
				"Edit " + (block.text.slice(0, 50) || "text box"),
			);
			place(
				node,
				state.page || 0,
				block.rect[0],
				block.rect[1],
				block.rect[2] - block.rect[0],
				block.rect[3] - block.rect[1],
			);
			node.addEventListener("click", (event) => {
				event.stopPropagation();
				closeText();
				selected = block;
				currentPage = state.page || 0;
				const input = document.createElement("textarea");
				input.className = "block-editor";
				input.setAttribute("aria-label", "Edit text");
				input.value = block.text;
				place(
					input,
					currentPage,
					block.rect[0],
					block.rect[1],
					block.rect[2] - block.rect[0],
					Math.max(block.rect[3] - block.rect[1], 32),
				);
				input.style.fontSize =
					((block.size || 12) *
						(pages.children[currentPage] as HTMLElement).clientWidth) /
						dimensions[currentPage][0] +
					"px";
				input.addEventListener("input", () => {
					pendingText = true;
					flushText();
				});
				input.addEventListener("keydown", (event) => {
					if (event.key === "Escape") {
						closeText();
					}
					if (event.key === "Enter" && (event.ctrlKey || event.metaKey)) {
						closeText();
					}
				});
				activeEditor = input;
				input.focus();
				publishToolbar();
			});
		}
	}
	if (kind === "word") {
		reportViewedPage();
	}
}
/**
 * The page being looked at (upstream #3880): the caret's page while the caret is
 * on screen, else the page at the top of the window. `topPage` is the first page
 * whose bottom is still below the window's top; the last page past the end.
 */
function viewedPage() {
	// Everything is measured in viewport coordinates: the window spans 0..innerHeight.
	const top = 0,
		bottom = window.innerHeight;
	const caretPage = state.caret?.page;
	const pageEl =
		caretPage !== undefined ? pages.children[caretPage] : undefined;
	if (pageEl && caretPage !== undefined && dimensions[caretPage]?.[0]) {
		const scale =
			pageEl.getBoundingClientRect().width / dimensions[caretPage][0];
		if (scale > 0) {
			// The caret as a thin rect, or the first selection rect on its page.
			const source =
				(state.rects || []).find((r) => r.page === caretPage) ||
				(state.caret
					? {
							page: caretPage,
							x: state.caret.x,
							y: state.caret.y,
							width: 0,
							height: state.caret.height,
						}
					: undefined);
			if (source) {
				const from = pageEl.getBoundingClientRect().top + source.y * scale;
				const to = from + Math.max(source.height, 1) * scale;
				if (to >= top && from <= bottom) {
					return caretPage;
				}
			}
		}
	}
	const children = pages.children;
	for (let index = 0; index < children.length; index++) {
		if ((children[index] as HTMLElement).getBoundingClientRect().bottom > top) {
			return index;
		}
	}
	return Math.max(0, children.length - 1);
}
let reportedPage: number | undefined;
function reportViewedPage() {
	if (kind !== "word") {
		return;
	}
	const page = viewedPage();
	if (page !== reportedPage) {
		reportedPage = page;
		send("pageChanged", { page });
	}
}
let scrollFrame = 0;
function scheduleViewedPage() {
	if (!scrollFrame) {
		scrollFrame = requestAnimationFrame(() => {
			scrollFrame = 0;
			reportViewedPage();
		});
	}
}
window.addEventListener("scroll", scheduleViewedPage, { passive: true });
window.addEventListener("resize", scheduleViewedPage, { passive: true });
// Wheel scrolling and scrollbar drags happen entirely inside the webview; the window
// toolbar's menus can't see them, so tell the host to dismiss any open menu.
let scrollNotice = 0;
window.addEventListener(
	"scroll",
	() => {
		if (!scrollNotice) {
			scrollNotice = requestAnimationFrame(() => {
				scrollNotice = 0;
				send("canvasActivity", {});
			});
		}
	},
	{ passive: true },
);
function defaultHint() {
	el("hint").textContent =
		kind !== "slides"
			? ""
			: "Click text to edit · Ctrl+Enter to finish · Ctrl+S to save";
}
let agentTimer: ReturnType<typeof setTimeout> | undefined;
function clearAgent() {
	clearTimeout(agentTimer);
	agentTimer = undefined;
	legacyAgentRects = [];
	agentOverlay();
}
/**
 * Shimmers over the region an agent is currently working on. The region comes
 * from the extension host, which derives it from the executed command; agents
 * never send ids or coordinates.
 */
function showAgent(
	region: {
		page: number;
		rects: { x: number; y: number; width: number; height: number }[];
	},
	label: string,
) {
	clearAgent();
	const element = pages.children[region.page] as HTMLElement | undefined;
	if (!element || !region.rects.length) {
		return;
	}
	const bounds = element.getBoundingClientRect();
	if (bounds.bottom < 80 || bounds.top > innerHeight - 40) {
		element.scrollIntoView({ block: "nearest" });
	}
	legacyAgentRects = region.rects.map((rect) => ({
		...rect,
		page: region.page,
	}));
	agentOverlay();
	el("hint").textContent = "Agent: " + label;
	agentTimer = setTimeout(() => {
		clearAgent();
		defaultHint();
	}, 2000);
}
keyboard.addEventListener("compositionstart", () => {
	agentReveal.clear();
	clearActivityPreview();
	composing = true;
});
keyboard.addEventListener("compositionend", (event) => {
	composing = false;
	if (event.data) {
		typeWordText(event.data);
	}
	keyboard.value = "";
});
keyboard.addEventListener("input", (event) => {
	agentReveal.clear();
	clearActivityPreview();
	if (
		composing ||
		(event as InputEvent).inputType === "insertCompositionText"
	) {
		return;
	}
	if (keyboard.value) {
		typeWordText(keyboard.value);
	}
	keyboard.value = "";
});
keyboard.addEventListener("paste", (event) => {
	event.preventDefault();
	const text = event.clipboardData?.getData("text/plain");
	if (text) {
		execute(kind === "word" ? "edit.pasteText" : "text.insert", { text });
	}
});
document.addEventListener("copy", (event) => {
	if (document.activeElement === keyboard && state.selectedText) {
		event.preventDefault();
		event.clipboardData?.setData("text/plain", state.selectedText);
	}
});
document.addEventListener("cut", (event) => {
	if (document.activeElement === keyboard && state.selectedText) {
		event.preventDefault();
		event.clipboardData?.setData("text/plain", state.selectedText);
		execute("text.delete");
	}
});
document.addEventListener("keydown", (event) => {
	agentReveal.clear();
	clearActivityPreview();
	if (
		document.querySelector("dialog[open]") ||
		(event.target instanceof HTMLElement &&
			event.target.closest(".word-ruler")) ||
		composing
	) {
		return;
	}
	// The webview forwards bubbling keys to workbench keybindings even after preventDefault.
	const consume = () => {
		event.preventDefault();
		event.stopPropagation();
	};
	const mod = event.ctrlKey || event.metaKey;
	if (
		event.key === "Escape" &&
		kind === "word" &&
		state.editingRegion &&
		state.editingRegion !== "body"
	) {
		consume();
		execute("insert.closeHeader");
		focusWordInput();
		return;
	}
	if (mod && event.key.toLowerCase() === "s") {
		consume();
		flushText();
		send("save");
		return;
	}
	if (mod && event.key.toLowerCase() === "z") {
		consume();
		closeText();
		send(event.shiftKey ? "redo" : "undo");
		return;
	}
	if (mod && event.key.toLowerCase() === "y") {
		consume();
		closeText();
		send("redo");
		return;
	}
	if (mod && event.key.toLowerCase() === "f" && kind === "word") {
		consume();
		openFind();
		return;
	}
	if (event.key === "Escape" && kind === "word" && !findBar.hidden) {
		consume();
		closeFind();
		return;
	}
	if (kind !== "word" || document.activeElement !== keyboard) {
		return;
	}
	const nav: Record<string, string> = {
		ArrowLeft: "left",
		ArrowRight: "right",
		ArrowUp: "up",
		ArrowDown: "down",
		Home: mod ? "docStart" : "home",
		End: mod ? "docEnd" : "end",
	};
	if (nav[event.key]) {
		consume();
		execute("caret." + nav[event.key], { extend: event.shiftKey });
		return;
	}
	if (mod) {
		const command: Record<string, string> = {
			a: "select.all",
			b: "format.bold",
			i: "format.italic",
			u: "format.underline",
		};
		if (command[event.key.toLowerCase()]) {
			consume();
			execute(command[event.key.toLowerCase()]);
		}
		return;
	}
	const commands: Record<string, string> = {
		Enter: event.shiftKey ? "text.lineBreak" : "text.newParagraph",
		Backspace: "text.backspace",
		Delete: "text.delete",
		Tab: event.shiftKey ? "text.backTab" : "text.tab",
	};
	if (commands[event.key]) {
		consume();
		execute(commands[event.key]);
	}
});
window.addEventListener("message", (event) => {
	if (event.origin !== location.origin) {
		return;
	}
	const data = event.data;
	if (data.type === "focusDocument") {
		if (
			!document.body.dataset.documentInstance ||
			data.instance === document.body.dataset.documentInstance
		) {
			focusWordInput();
		}
		return;
	}
	if (data.type === "openFind") {
		openFind();
		return;
	}
	if (data.type === "tablePreview") {
		wordTable.preview(data.request, data.revision, data.result);
		return;
	}
	if (data.type === "publishToolbar") {
		publishToolbar();
		return;
	}
	if (data.type === "toolbarAction") {
		toolbarAction(data.action, data.context, data.instance);
		return;
	}
	if (data.type === "agentCommit") {
		const rects = Array.isArray(data.rects) ? data.rects : [];
		if (
			rects.length &&
			(!data.request || data.request !== followRequest || followAgent)
		) {
			const marker = document.createElement("div");
			marker.className = "agent-commit-target";
			const rect = rects[0];
			if (pages.children[rect.page]) {
				place(marker, rect.page, rect.x, rect.y, rect.width, rect.height);
				followAgentChange(data.request, ".agent-commit-target");
				marker.remove();
			}
		}
		agentReveal.queue(data.revision, rects);
		renderNext();
		return;
	}
	if (data.type === "pdfInspection") {
		pdfInspection.hidden = false;
		pdfInspection.open = true;
		const summary = document.createElement("summary");
		summary.textContent = "PDF Inspection";
		pdfInspection.replaceChildren(summary);
		const body = document.createElement("div");
		renderPdfInspection(body, data.result);
		pdfInspection.append(body);
		return;
	}
	if (data.type === "pdfResource") {
		const resolve = resourcePickers.get(data.request);
		resourcePickers.delete(data.request);
		resolve?.(data.uri);
		return;
	}
	if (data.type === "agentActivity") {
		if (data.revision !== undefined && data.revision !== state.revision) {
			return;
		}
		clearTimeout(activityTimer);
		activityTimer = undefined;
		agentLabel =
			typeof data.label === "string" ? data.label.slice(0, 40) : "Agent";
		agentIntent =
			typeof data.intent === "string" ? data.intent.slice(0, 240) : undefined;
		agentRects = Array.isArray(data.rects) ? data.rects : [];
		const insertions = Array.isArray(data.insertions) ? data.insertions : [];
		wordReview.updateActivity(insertions, agentLabel, data.expiresAt);
		wordReview.layout(dimensions);
		overlay();
		if (agentRects.length) {
			activityTimer = setTimeout(
				clearActivityPreview,
				Math.max(
					0,
					Math.min(30000, (data.expiresAt || Date.now() + 30000) - Date.now()),
				),
			);
		}
		agentOverlay();
		if (insertions.length) {
			followAgentChange(data.request, ".word-activity-zone");
		} else if (agentRects.length) {
			followAgentChange(data.request, ".agent-shimmer-border");
		}
		return;
	}
	if (data.type === "draft") {
		wordReview.updateDrafts(
			data.sections,
			data.label,
			data.workingId,
			data.expiresAt,
		);
		wordReview.layout(dimensions);
		overlay();
		agentOverlay();
		followAgentChange(data.request, '.word-draft-zone[data-working="true"]');
		return;
	}
	if (data.type === "review") {
		wordReview.update(data.review);
		showReviewSummary();
		wordReview.layout(dimensions);
		overlay();
		agentOverlay();
		return;
	}
	if (data.type === "error") {
		if (kind === "pdf") {
			pdfCanvas.failed();
		}
		message(data.error);
		renderBusy = false;
		return;
	}
	if (data.type === "saved") {
		el("hint").textContent = data.pendingSections
			? `Saved · ${data.pendingSections} pending sections are editor-only`
			: "Saved";
		return;
	}
	if (data.type === "agent") {
		if (data.region) {
			showAgent(data.region, data.label);
		}
		return;
	}
	if (data.type === "state") {
		const firstState = !initialized;
		kind = data.kind;
		const changed = state.revision !== data.state.revision;
		const regionChanged = state.editingRegion !== data.state.editingRegion;
		if (changed) {
			agentReveal.clear();
			clearActivityPreview();
		}
		const repaginated =
			JSON.stringify(dimensions) !== JSON.stringify(data.pages);
		state = data.state;
		dimensions = data.pages;
		if (Array.isArray(data.commands)) {
			documentCommands = data.commands;
		}
		if (kind === "pdf" && state.pdfArtifacts?.length) {
			pdfArtifacts.hidden = false;
			pdfArtifacts.replaceChildren();
			const summary = document.createElement("summary");
			summary.textContent = `Generated Files (${state.pdfArtifactCount || state.pdfArtifacts.length})`;
			pdfArtifacts.append(summary);
			for (const artifact of state.pdfArtifacts) {
				const open = document.createElement("button");
				open.textContent = artifact.name;
				open.addEventListener("click", () =>
					send("pdfOpenArtifact", { uri: artifact.uri }),
				);
				pdfArtifacts.append(open);
			}
			if ((state.pdfArtifactCount || 0) > state.pdfArtifacts.length) {
				const more = document.createElement("button");
				more.textContent = "Show All Generated Files";
				more.addEventListener("click", () => send("pdfArtifactList"));
				pdfArtifacts.append(more);
			}
		}
		if (kind === "pdf") {
			currentPage = Math.min(
				state.page ?? currentPage,
				Math.max(0, dimensions.length - 1),
			);
		}
		wordTable.update(state.revision);
		wordImage.update(state.revision, state.image);
		revealWordCaret =
			kind === "word" &&
			(data.command?.startsWith("text.") ||
				data.command === "insert.table" ||
				data.command === "insert.editHeader" ||
				data.command === "insert.editFooter" ||
				data.command === "edit.goto" ||
				data.command === "edit.find" ||
				data.command?.startsWith("edit.findNext") ||
				data.command?.startsWith("edit.findPrevious") ||
				data.command?.startsWith("edit.replace") ||
				data.command === "edit.pasteFormat");
		const commentsOpen = comments.open;
		comments.replaceChildren();
		comments.open = commentsOpen;
		comments.hidden = kind === "word" || !state.comments?.length;
		const summary = document.createElement("summary");
		summary.textContent = `Comments (${state.comments?.length || 0})`;
		comments.append(summary);
		for (const comment of state.comments || []) {
			const p = document.createElement("p");
			p.textContent =
				(comment.resolved ? "Resolved · " : "") +
				(comment.author || "Comment") +
				": " +
				comment.text;
			if (typeof comment.parentId === "number") {
				p.style.marginInlineStart = "24px";
			}
			comments.append(p);
		}
		if (!initialized) {
			initialized = true;
			if (kind === "pdf") {
				document.body.classList.add("pdf-document");
			}
			buildToolbar();
		}
		if (kind === "word") {
			for (const [key, value] of Object.entries({
				bold: state.formatting?.bold,
				italic: state.formatting?.italic,
				underline: state.formatting?.underline,
				strike: state.formatting?.strike,
			})) {
				if (typeof value === "boolean") {
					toolbarControls[key]?.setAttribute("aria-pressed", String(value));
				}
			}
			const size = toolbarControls.size;
			if (size) {
				size.dataset.currentSize = String(state.formatting?.size ?? "");
				toolbarControls.font.dataset.currentFont = state.formatting?.font || "";
				size.dispatchEvent(new Event("formatstate"));
			}
			const style = el("paragraph-style");
			style.dataset.currentStyle = state.paragraphStyle || "";
			style.dispatchEvent(new Event("formatstate"));
			if (state.find) {
				findCount.textContent = state.find.count
					? `${state.find.current + 1} of ${state.find.count}`
					: "No matches";
				findPrevious.disabled = findNextButton.disabled = !state.find.count;
			} else {
				findCount.textContent =
					findQuery && !findBar.hidden ? "No matches" : "";
				findPrevious.disabled = findNextButton.disabled = true;
			}
		}
		el("filename").textContent = data.name;
		defaultHint();
		(el("undo") as HTMLButtonElement).disabled = !state.undo;
		(el("redo") as HTMLButtonElement).disabled = !state.redo;
		if (repaginated || !pages.children.length) {
			closeText();
			makePages();
		}
		if (changed || regionChanged) {
			rendered.clear();
			renderNext();
		}
		if (kind === "word") {
			wordReview.layout(dimensions);
			wordParagraph?.update(state, dimensions);
			wordPageSetup?.update(state);
		}
		if (kind === "pdf") {
			pdfCanvas.update();
		}
		overlay();
		agentOverlay();
		publishToolbar();
		if (firstState && document.hasFocus()) {
			focusWordInput();
		}
	} else if (data.type === "rendered") {
		renderBusy = false;
		const request = pendingRender.get(data.request);
		pendingRender.delete(data.request);
		if (
			!request ||
			(kind !== "word" && request.revision !== state.revision) ||
			(kind === "word" && request.region !== state.editingRegion) ||
			!pages.children[data.page]
		) {
			renderNext();
			return;
		}
		const node = pages.children[data.page];
		const result = data.result;
		if (result.svg) {
			const clean = DOMPurify.sanitize(result.svg, {
				USE_PROFILES: { svg: true, svgFilters: true },
				FORBID_TAGS: ["a", "script", "foreignObject", "animate", "set"],
			});
			const parsed = new DOMParser().parseFromString(clean, "image/svg+xml");
			for (const element of parsed.querySelectorAll("[href],[*|href]")) {
				for (const attr of Array.from(element.attributes)) {
					if (
						attr.localName === "href" &&
						!attr.value.startsWith("#") &&
						!/^data:image\/(png|jpeg|gif|webp);base64,/i.test(attr.value)
					) {
						element.removeAttributeNode(attr);
					}
				}
			}
			(node.querySelector("img") as HTMLImageElement).src =
				"data:image/svg+xml;charset=utf-8," +
				encodeURIComponent(new XMLSerializer().serializeToString(parsed));
		} else {
			const canvas = node.querySelector("canvas")!;
			const height = result.canvasHeight ?? result.height;
			if (canvas.width !== result.width || canvas.height !== height) {
				canvas.width = result.width;
				canvas.height = height;
			}
			if (result.height > 0) {
				const pixels = result.pixels as Uint8Array<ArrayBuffer>;
				canvas
					.getContext("2d")!
					.putImageData(
						new ImageData(
							new Uint8ClampedArray(
								pixels.buffer,
								pixels.byteOffset,
								pixels.byteLength,
							),
							result.width,
							result.height,
						),
						0,
						result.y ?? 0,
					);
			}
			canvas.dataset.frame = String(data.request);
			canvas.dataset.revision = String(request.revision);
			canvas.dataset.region = request.region || "body";
			if (kind === "word") {
				wordReview.paint(data.page, dimensions);
				agentReveal.paint();
				if (request.revision === state.revision) {
					wordTable.render(data.page, request.revision, result.tables || []);
					wordImage.render(data.page, request.revision, result.images || []);
				}
			}
		}
		// Show completed Word frames during continuous typing instead of discarding
		// every frame until the user stops. A newer revision still requests a refresh.
		if (request.revision === state.revision) {
			rendered.add(data.page);
		}
		renderNext();
	}
});
let resizeTimer: ReturnType<typeof setTimeout>;
window.addEventListener("resize", () => {
	agentReveal.clear();
	clearTimeout(resizeTimer);
	resizeTimer = setTimeout(() => {
		rendered.clear();
		renderNext();
		if (kind === "word") {
			wordReview.layout(dimensions);
		}
		overlay();
		agentOverlay();
	}, 200);
});
send("ready");
