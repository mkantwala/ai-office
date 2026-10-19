# AI Office editing commands

Word uses a visual paragraph-style menu and a 10-column by 8-row table picker. Table insertion sends `insert.table` with `rows`, `cols` and `borders: true`, ensuring visible borders even when an imported document lacks a usable TableGrid style. Title, Subtitle and Heading previews dispatch the same `para.style` command used by agents.

Word input batches adjacent text for up to 16 ms, flushes before subsequent commands, and displays completed render frames during continuous typing. DOCX serialization occurs on save or recovery backup instead of on each keystroke. PDF and PowerPoint retain their existing toolbars and input handling.

The toolbar, extension-host workers and CLI share these engine adapters. No network server is started. Word calls Wordcraft's command registry; PDF calls Pdfcraft's automation table; PowerPoint calls office-kit through the shared adapter. File/image pickers collect parameters for headless document commands.

Word's compact toolbar groups history/save, paragraph style, font, emphasis, paragraph layout and insertion controls in a single row. More Formatting always contains table insertion and receives additional groups when the editor narrows. Controls keep their native keyboard behavior and document actions when moved. Escape closes the popover and returns focus to More Formatting. PDF and PowerPoint retain their existing toolbar layouts.

The visual references are Google's [refreshed Docs interface](https://workspaceupdates.googleblog.com/2023/03/refreshed-ui-google-drive-docs-sheets-slides.html) and [formatting toolbar documentation](https://support.google.com/docs/answer/1663349). Icons are local, original line drawings and colors follow the workbench theme. Font/style selectors describe actions because the current engine bridge does not expose the selection's formatting; Word emphasis buttons ask the engine to toggle its actual selection state.

Build the Word/PDF webview with `node extensions/ai-office-pdf/build.mts --ui-only`. Run the browser toolbar checks from the repository root with `node --test build/lib/test/officeToolbar.test.ts` after installing Playwright Chromium, or set `PLAYWRIGHT_CHROMIUM_EXECUTABLE_PATH` to an installed Chromium executable. The checks cover narrow/wide layouts, overflow actions, keyboard dismissal and the existing PDF/PowerPoint toolbar controls.

From the VS Code repository:

```powershell
.\scripts\ai-office-cli.cmd commands C:\documents\memo.docx
.\scripts\ai-office-cli.cmd inspect C:\documents\slides.pptx
.\scripts\ai-office-cli.cmd run C:\documents\memo.docx --script changes.json --output C:\documents\memo-edited.docx
```

`commands` prints supported IDs and parameter descriptions. `inspect` prints page dimensions, Word text (in `selectedText`), current PDF paragraphs/PowerPoint text boxes and comments. It selects all Word text in memory without saving the file. `run` starts Word with a caret; use `select.text` to target known text or `select.all` in a script. It prints the final state. Scripts are JSON arrays; output retains the input format. Use `--in-place` instead of `--output` to overwrite the input with a conflict check. Writes use a temporary sibling file and rename. Use `--root C:\documents` to confine input, script and output paths. Encrypted Word/PDF passwords can be passed through the `AI_OFFICE_PASSWORD` environment variable. Errors are JSON on stderr with a nonzero exit code.

Example `changes.json` for Word:

```json
[
  { "command": "select.text", "params": { "text": "Quarterly report" } },
  { "command": "format.bold", "params": { "value": true } },
  { "command": "format.size", "params": { "size": 24 } },
  { "command": "review.newComment", "params": { "text": "Check the figures" } }
]
```

PDF example (page/paragraph numbers start at 1; coordinates use points from the top left):

```json
[
  { "command": "text_edit", "params": { "page": 1, "paragraph": 1, "text": "Updated paragraph", "bold": true, "size": 14 } },
  { "command": "comment_add", "params": { "page": 1, "type": "note", "at": [40, 40], "contents": "Review this" } }
]
```

PowerPoint example (find shape IDs with `inspect`; slide numbers start at 1):

```json
[
  { "command": "text.replace", "params": { "page": 1, "shape": 2, "text": "Updated title" } },
  { "command": "format.text", "params": { "page": 1, "shape": 2, "size": 32, "bold": true } },
  { "command": "para.align", "params": { "page": 1, "shape": 2, "value": "center" } }
]
```

PowerPoint cell operations add zero-based `row` and `col` alongside the table's shape ID. `format.text` accepts optional `range: {start,end}` with UTF-16 offsets. Word's `caret.set` uses a zero-based page and point coordinates. Use the CLI's `commands` output for engine-specific parameters.

Agents can skip shape and paragraph ids entirely with the `content.*` commands: `content.find {find, occurrence?, page?}`, `content.replace {find, text, occurrence?, all?, page?}`, `content.delete {find, occurrence?, all?, page?}` and `content.append {text, page?}`. `find` matches a substring inside a block/shape/cell; the engine target (Word selection, PDF paragraph, PowerPoint shape or table cell) is resolved internally. `all: true` replaces or deletes every non-overlapping match (Word handles at most 500). The same commands are available for open documents through the `aiOffice.<kind>.execute` VS Code command, which takes `(fileNameOrUri?, commandId, params)`, and through the `aiOfficeEdit` language model tool, which chat agents use. Both resolve the file themselves: without a file argument they edit the current active file, and a bare file name or path fragment selects another open document. The result reports the `name` and `uri` of the edited file. While an agent command runs, the editor shimmers over the region it touched. See `.agents/skills/ai-office-docs/SKILL.md` for the compact agent skill.

```json
[
  { "command": "content.replace", "params": { "find": "Q3 revenue", "text": "$4.2M" } },
  { "command": "content.append", "params": { "text": "Signed: Finance" } }
]
```

| Toolbar | Word | PDF | PowerPoint |
|---|---|---|---|
| Text | `text.insert` | `page_add_text` / `text_edit` | `insert.text` / `text.replace` |
| Style | `para.style` | Disabled | Disabled |
| Font / size / B I U | `format.font/size/bold/italic/underline` | `text_edit` | `format.text` |
| Align | `para.align` | `text_edit` | `para.align` |
| List | `para.bullets/numbering` | Disabled | `para.list` |
| Spacing | `para.lineSpacing` | `text_edit` | `para.spacing` |
| Table | `insert.table` | Disabled | `insert.table` |
| Image (base64 bytes) | `insert.picture` | `image.insert` | `insert.image` |
| Link | `insert.link` | `link_add` | `insert.link` |
| Comment | `review.newComment` | `comment_add` | `insert.comment` |
| Undo / redo | `edit.undo/redo` | `edit_undo/edit_redo` | `edit.undo/redo` |
| Save | CLI output flags | CLI output flags | CLI output flags |

Build both Rust engines and both extensions before running. Editing round-trip and save/conflict tests use original synthetic fixtures:

```powershell
node extensions/ai-office-common/test.mts
```

The automated tests create synthetic documents in `.build/office-edit-demo`. Run them from the repository root after changing the provider or the engines:

```powershell
node extensions/node_modules/esbuild/bin/esbuild extensions/ai-office-common/provider.ts --bundle --platform=node --format=cjs --external:vscode --external:./client.js --outfile=.build/office-provider.cjs
node --test extensions/ai-office-common/provider.test.cts
node extensions/node_modules/esbuild/bin/esbuild extensions/ai-office-common/engine.test.ts --bundle --platform=node --format=esm --outfile=.build/office-engine.test.mjs
node --test .build/office-engine.test.mjs
```

Word rendering sends changed pixel rows when the webview still has the matching previous frame. Typing uses a temporary raster capped at one pixel per document point; full display resolution returns after a 200 ms pause. This affects only the preview, not layout or saved files. Imported DOCX files can apply missing built-in paragraph styles through the same `para.style` command used by the CLI; existing style definitions are preserved.

Live Word agents use `office_document_read` (outline/section/search/chunk reads), `office_document_edit` (native blocks and guarded literal patches), `office_document_format`, `office_document_draft` (temporary pending sections and atomic fills), `office_document_comments` (anchored native threads), `office_document_history`, `office_document_activity` and `office_document_help`. The compact skill is `extensions/ai-office-pdf/skills/ai-office-docs/SKILL.md`; detailed guides load on demand. Block hashes permit text edits across unrelated revisions while protecting changed formatting/objects; structural edits require a fresh document-instance version. Pending cards are session-local UI only and never serialized; completed fills use the normal DOCX history and Keep/Undo review. Git is not involved.

Word hardening: caret geometry uses the active font’s ascent/descent rather than paragraph line spacing; the webview keeps a thin display stroke. Review view zones are reused across caret-only updates. Toolbar feedback uses one delegated tooltip, CSS press/menu transitions, reduced-motion overrides and disposable listeners. Responsive popovers close their parent overflow and restore visible keyboard focus. A successful native agent mutation returns a successful receipt even if a later preview/review refresh fails; warnings explain the unavailable follow-up without inviting duplicate edits. A cancelled/failed Save does not report “Saved” while the document remains dirty. Worker request serialization failures clear their pending timers and leave the worker usable. These are verified behaviors, not a guarantee against every platform, document or deployment failure.

Paragraph properties share a bounded native `para.properties` patch between the user controls and `office_document_format` (`format.paragraph`). Exact paragraph/heading/list-item reads include effective `paragraph` properties; summary reads stay compact. Omitted properties remain unchanged. Text, inline objects, comments and stable block IDs survive formatting. Agent edits require the existing fresh document version and enter normal native review. Structural table/row/cell IDs cannot masquerade as paragraph targets; use the paragraph IDs inside cells.

The Paragraph toolbar button opens alignment, indents, before/after spacing, multiple/exact/at-least line spacing, tabs with leaders, per-side borders, explicit shading and pagination controls. Mixed properties are shown blank or indeterminate; only edited fields are submitted. The modal closes without applying on Cancel/Escape, and rejects stale document/selection state. Tab editing replaces custom stops (empty clears custom stops, not document-default intervals); border-side patches preserve other sides and omitted side attributes. Shading uses an explicit six-digit RGB color; white fill is `FFFFFF`.

The horizontal ruler uses document points and page scaling, supports draggable left/right/first-line indents and tab stops, snaps to six points (Alt: quarter points), adds tabs by clicking, and removes a tab by dragging it off the ruler. Focused markers support arrow keys (Shift: ten points), and tab markers support Delete/Backspace. Each completed drag is one native Undo step. The ruler appears for ordinary, unnumbered, horizontal, single-column body paragraphs with consistent indents. Complex containers, numbered lists, RTL and mirrored indents use the properties dialog instead of guessed ruler geometry; their paragraph formatting remains available. Ruler positioning is outside the document canvas, so inline review view zones do not change its coordinate system.

Sangam One hosts the compact document toolbar in the active workbench editor group, below its tabs. Word, PDF and slide webviews keep their command handlers and advanced dialogs while hiding their local formatting header. The workbench reuses the shared controls in `src/vs/base/browser/ui/documentToolbar`, including style/table menus, overflow and tooltips. It shows controls for the active supported custom editor and removes the toolbar for other editor types.

Toolbar messages contain bounded control metadata and a selection context, not document text. The route pins the active URI, panel, document instance and context; both sides refuse stale or inactive targets. URI objects cross the extension-host boundary so remote `vscode-remote` resources map to their native file URI correctly. Existing editor actions still flush pending typing before formatting, use native document history, and enter the same agent review flows. Standalone extension hosts without the workbench toolbar capability retain the compact local toolbar. The Word ruler stays with the document; advanced property dialogs open on demand.

Formatting uses real compact buttons with keyboard-accessible option menus, local line icons, hover tooltips and reduced-motion-aware transitions. The Aa button opens style previews; font/size, alignment, lists and spacing use button menus. Text color and highlight show named swatches; strikethrough and Clear Formatting call native commands. The same operations are available through `office_document_format` as `strikethrough`, `color` (RRGGBB or auto), `highlight` (Word color or none) and `clear:true` (alone). The native bridge allowlist and on-demand agent guide include these commands.

The Word comments summary bar is temporarily hidden; native comment data and agent comment commands remain available.

The paragraph ruler displays inches from the text margin, with adaptive subdivisions, shaded page margins and vector indent/tab handles. Its width follows the rendered page in CSS pixels, including narrow views and high-DPI displays. Native measurements and guarded agent edits remain in points; focused/dragged markers announce both units. Tab handle shapes distinguish left, center, right, decimal and bar stops.

## Opt-in live Cloudflare evaluation

`cloudflareAgent.live.cts` exercises the registered chat participant, actual document tool backend, request history and native Word worker against synthetic DOCX files. It replaces only the VS Code UI/service shell. It makes real, billable GLM-4.7-Flash requests and is deliberately excluded from the normal test runner. Supply `CF_ACCOUNT_ID` and `CF_API_TOKEN` through the process environment; no credentials belong in the repository. Build the extension and Word engine first, then run from the repository root:

```sh
node --use-env-proxy extensions/ai-office-common/cloudflareAgent.live.cts
```

Use `--only=format-all,history-undo-redo` to rerun selected scenarios. The Node proxy flag is useful in managed environments; ordinary direct-network installations can omit it. JSON traces and synthetic result DOCX files are written under ignored `.build/live-cloudflare`. The checks inspect actual native content, DOCX XML, formatting and history rather than accepting the model's completion claim. Results are model evaluations and can vary between runs; a failed assertion is retained as a failure. Missing PDF engine artifacts can block the separate all-format suite without affecting this Word evaluation.

Formatting tools accept both `RRGGBB` and `#RRGGBB`, canonicalizing colors before native edits. Paragraph distances are explicitly points in the tool schema. Deletion receipts distinguish required structural paragraphs from requested new content, and the native agent refuses repeated deletion of those paragraphs. An initial tool-free model answer receives an evidence check before it can be used as the document agent's answer.

The live harness stops subsequent network scenarios after HTTP 401/403/429 and reports them as blocked, never passed. Comment text updates can include an explicit resolved boolean; both changes commit atomically, and operation-specific fields are validated instead of silently ignored.
