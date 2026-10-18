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
