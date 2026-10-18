---
name: ai-office-docs
description: Read, create, update and delete content in AI Office documents (DOCX, PDF, PPTX). Use when an agent must write document contents, inspect Word/PDF/PowerPoint text, replace or delete document text, append paragraphs or slides text, or drive an open AI Office editor.
---

# AI Office documents

AI Office edits DOCX, PDF and PPTX files. Agents change document contents without ever handling document, shape or paragraph ids: commands target content by text and resolve the engine target internally.

## Surfaces

- **Files (headless):** `scripts\ai-office-cli.cmd` from the repository root (`node extensions/ai-office-pdf/out/cli.cjs` on any platform). Operates on a copy; the input file is untouched until a successful rename.
- **Open editor (live):** VS Code command `aiOffice.<kind>.execute(fileNameOrUri?, command, params)` with `kind` = `word`, `pdf` or `slides`, applies to the already open document session and participates in its undo/save flow. Without a file argument the command edits the current active file; a bare file name (`'memo.docx'`) or a path fragment selects another open document. The result carries `name` and `uri` of the file that was edited. While it runs, the editor shimmers over the region the command touched.
- **Open editor (chat tool):** the `aiOfficeEdit` language model tool (`#aiOffice`) takes `{ command, file?, params? }` and runs the same engine commands on an open document. It is the in-app agent's surface: with no `file` it edits the document the user is looking at — the user never has to name a file, and the agent must never ask for one. The tool result reports the resolved file `name`, `uri` and the new document state.

Prefer the `content.*` commands below. Use engine commands (further down) only for formatting, caret work or inserts that need explicit parameters.

## Workflow (files)

1. `ai-office-cli.cmd inspect FILE` — read contents. Word: all text in `selectedText`. PDF: `blocks` (`id` = paragraph). PPTX: `blocks` (`id` = shape; table cells carry `row`/`col`).
2. Write a JSON array of steps: `[{ "command": "...", "params": { ... } }]`.
3. `ai-office-cli.cmd run FILE --script changes.json --output result.docx` — or `--in-place` to overwrite the input after a conflict check. Add `--root DIR` to confine all paths. Output must keep the input format.
4. `inspect` the output to verify. Errors are JSON on stderr with exit code 1; nothing is written on failure.

## Content commands (no ids, every format)

| Command | Params | Behavior |
|---|---|---|
| `content.find` | `{ find, occurrence?, page? }` | Word: selects the match and returns its `rects`. PDF/PPTX: returns matches in `matches`. Missing text is an error. |
| `content.replace` | `{ find, text, occurrence?, all?, page? }` | Replaces the matched text. `all: true` replaces every non-overlapping match. |
| `content.delete` | `{ find, occurrence?, all?, page? }` | Deletes the matched text; deleting a whole PDF paragraph removes it. |
| `content.append` | `{ text, page? }` | New paragraph at the document end (Word), text block below the lowest block (PDF), text box below existing boxes (PPTX). Defaults to the last page/slide. |

- `find` matches a substring inside a block/shape/cell; `occurrence` (1-based, default 1) picks the nth match in document order. Word `all` handles at most 500 matches.
- Word replacement replaces the selection; PDF preserves existing paragraph formatting; PPTX preserves shape/cell formatting.
- `page` is a 1-based page (PDF) or slide (PPTX); it is ignored by Word.

```json
[
  { "command": "content.replace", "params": { "find": "Q3 revenue", "text": "$4.2M" } },
  { "command": "content.append", "params": { "text": "Signed: Finance" } }
]
```

## Word commands

| Commands | Params | Purpose |
|---|---|---|
| `text.insert` `text.newParagraph` `text.lineBreak` `text.tab` `text.backTab` `text.backspace` `text.delete` | `text.insert {text}`; others none | Typing at the caret/selection |
| `caret.left` `caret.right` `caret.up` `caret.down` `caret.home` `caret.end` `caret.docStart` `caret.docEnd` | `{extend?}` | Move caret / extend selection |
| `caret.set` | `{pos:{page,x,y} points from page top-left, extend?}` | Position the caret |
| `select.all` `select.word` `select.collapse` `select.text` `select.range` | `select.text {text, occurrence?}`; `select.range {anchor, focus}` | Selection |
| `edit.undo` `edit.redo` | none | History |
| `format.bold` `format.italic` `format.underline` `format.font` `format.size` | `{value?}`; underline `{value?, style?}`; font `{name?, eastAsia?}`; size `{size}` | Character formatting on the selection |
| `para.align` `para.lineSpacing` `para.style` `para.bullets` `para.numbering` | align `{value: left\|center\|right\|justify\|distribute\|start\|end}`; lineSpacing `{value\|atLeast\|exactly}`; style `{style: name or id}`; bullets/numbering `{kind?, off?}` | Paragraph formatting |
| `insert.table` `insert.picture` `insert.link` | table `{rows, cols, style?, borders?}`; picture `{path?\|data: base64, width?, alt?}`; link `{url, text?}` | Insert at the caret |
| `review.newComment` `review.comments` | `{text}` | Comments |

## PowerPoint commands (per slide, pages 1-based)

| Commands | Params | Purpose |
|---|---|---|
| `select.shape` | `{page, shape}` | Focus a shape (rarely needed; `content.*` resolves by text) |
| `text.replace` | `{page, shape, row?, col?, text}` | Set a text box or table cell text |
| `insert.text` | `{page, text, at?: [x, y] inches}` | New text box |
| `format.text` | `{page, shape, row?, col?, font?, size?, bold?, italic?, underline?, range?: {start, end} UTF-16 offsets}` | Format a shape, cell or text range |
| `para.align` `para.list` `para.spacing` | `{page, shape, row?, col?, value}` (align: left/center/right/justify; list: bullet/number/none; spacing: 0.5-3) | Paragraph layout |
| `insert.table` `insert.image` `insert.link` `insert.comment` | table `{page, rows 1-20, cols 1-10}`; image `{page, data: base64}`; link `{page, shape, url}`; comment `{page, text}` | Inserts |
| `edit.undo` `edit.redo` | none | History |

## PDF commands (pages and paragraphs 1-based, points from page top-left)

| Commands | Params | Purpose |
|---|---|---|
| `text_edit` | `{page, paragraph, text?, bold?, italic?, underline?, font?, size?, align?, line_spacing?}` | Edit a paragraph; omit formatting fields to keep existing formatting, empty text removes the paragraph |
| `page_add_text` | `{page, at: [x, y], width, text}` | Add a text block |
| `content_update` | `{page, index, ...}` | Edit previously added content |
| `link_add` `comment_add` | link `{page, rect: [x1, y1, x2, y2], url}`; comment `{page, type, at, contents}` | Annotations |
| `image.insert` | `{page, data: base64, name}` | Insert an image (16 MB limit) |
| `select.page` | `{page}` | Make a page current |
| `edit_undo` `edit_redo` | none | History |

## Rules

- `ai-office-cli.cmd commands FILE` prints the exact parameter schema for that file's format — use it instead of guessing engine parameters.
- Scripts hold up to 10,000 steps; text arguments are limited to 100,000 characters.
- Encrypted Word/PDF files: pass the password in the `AI_OFFICE_PASSWORD` environment variable.
- Do not run the CLI against a file with unsaved changes in the app; save first or use `aiOffice.<kind>.execute` / the `aiOfficeEdit` tool while it is open.
- The live command and the `aiOfficeEdit` tool resolve the file themselves: with no file argument they edit the current active file, and unresolved targets come back as errors listing the open documents. Never invent a path and never ask the user for one.
- In the app, one agent command can add several undo steps (composed PDF/PPTX edits are grouped by the engine; Word sequences create one step per engine command).
- Example live edit (chat tool): `#aiOffice` with `{ "command": "content.append", "params": { "text": "Approved" } }` appends to the active file. Add `"file": "memo.docx"` (or a URI) to target another open document; the raw command form is `executeCommand('aiOffice.word.execute', undefined, 'content.append', { text: 'Approved' })`.
