# AI Office editing commands

The toolbar, extension-host workers and CLI share these engine adapters. No network server is started. Word calls Wordcraft's command registry; PDF calls Pdfcraft's automation table; PowerPoint calls office-kit through the shared adapter. File/image pickers collect parameters for headless document commands.

Word's full interface dispatches the same native engine registry. The headless Word adapter exposes over 600 commands, excluding host-owned open/save/export/close operations; use CLI output flags to save. Query command results appear in the final state's `result`. For open documents, `aiOffice.word.execute(uriString, commandId, params)` returns the native result. `aiOffice.word.control(uriString, method, params)` exposes Wordcraft's structured control protocol, including `engine.commands`, `engine.execute`, `document.inspect` and `ui.inspect`, returning its `{ok,result}` envelope. Native ribbon and UI commands require an open editor. File requests open host dialogs and must not be used for unattended CLI execution.

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

The tests create a small demo folder at `.build/office-edit-demo`. User documents are never used as test fixtures.
