# Discover exact capabilities on demand

Prefer `office_document_format` for common character and paragraph formatting. It resolves all literal matches and actual style sizes in one atomic operation. Use the native commands below for other capabilities.

`office_document_help` without arguments returns categories. Pass `category` for a small related set or `command` for a single exact parameter schema from the running engine. The running registry is authoritative. This index covers all commands currently enabled by the Word bridge; availability can change with engine versions.

| Category | Enabled commands | Agent usage |
|---|---|---|
| Text | `text.insert`, `edit.pasteText`, `text.newParagraph`, `text.lineBreak`, `text.tab`, `text.backTab`, `text.backspace`, `text.delete` | Prefer structured insert/update/delete. `edit.pasteText` inserts clipboard text as paragraphs in one undo step; `text.lineBreak` stays within a paragraph. Targeted text commands act at the end of the target paragraph range; backspace can join paragraphs. |
| Character formatting | `format.bold`, `format.italic`, `format.underline`, `format.font`, `format.size`, `format.strikethrough`, `format.color`, `format.highlight`, `format.clear` | Use targeted command; send explicit values where supported rather than toggling unknown formatting. |
| Paragraph formatting | `para.properties`, `para.align`, `para.style`, `para.bullets`, `para.numbering`, `para.lineSpacing` | Prefer `office_document_format` with `format.paragraph` for a typed partial patch; `para.properties` accepts that same object on a paragraph/heading/list-item ID. Read the exact block for its effective `paragraph` properties. List items retain their paragraph identity. |
| Page settings | `layout.pageProperties` | Guarded partial margins and per-side borders in points; use `page` (zero-based) and `applyTo: "section"` or `"document"`. Margins: top/bottom/left/right/gutter. Borders: top/bottom/left/right objects (style, width 0.25–6, color RRGGBB or null, space 0–31), null clears a side; borders null clears all. Omitted properties, header/footer references and block IDs stay intact. |
| Headers and footers | `insert.editHeader`, `insert.editFooter`, `insert.closeHeader` | Header/footer navigation uses a zero-based `page`. Missing stories are created with native Undo. Then read the new part and edit its paragraph IDs; body IDs remain unchanged. |
| Insertion | `insert.table`, `insert.picture`, `insert.link` | Use targeted command; prefer structured table insertion when populating cells. Pictures accept host-resolved image fileUri and native size parameters; get help. Inline images/links are addressed through their containing paragraph, not independent block IDs. |
| Review | `review.newComment`, `review.comments`, `review.wordCount` | Add comments through targeted commands. Comments can also be inspected through their document story. These commands do not implement acceptance/rejection of agent changes. `review.wordCount` is read-only and returns `pages`, `words`, `characters` (excluding spaces), `charactersWithSpaces`, `paragraphs` and `lines`. |
| Caret (editor only) | `caret.set`, `caret.left`, `caret.right`, `caret.up`, `caret.down`, `caret.home`, `caret.end`, `caret.docStart`, `caret.docEnd`, `edit.goto` | The UI uses these. Agent tools use IDs and ranges instead; targeted command rejects them. `edit.goto` moves the caret to a page, paragraph or bookmark: `{"page": n}` is one-based and clamped to the document, `{"pageBy": n}` moves relatively, `{"paragraph": n}` and `{"bookmark": name}` target positions; it changes the selection only, never the text. |
| Selection (editor only) | `select.all`, `select.range`, `select.word`, `select.text`, `select.collapse` | Read `operation:selection` for user context. Targeted command rejects these to avoid losing the target. |
| History (editor only) | `edit.undo`, `edit.redo` | Normal editor Undo/Redo handles complete agent operations. These commands are excluded from targeted edits because they cannot be scoped to a block. Use the separate `office_document_history` tool for status, undo, redo or restoring a request checkpoint. |
| Format painter (editor only) | `edit.formatPainter`, `edit.pasteFormat` | The UI uses these. `edit.formatPainter` copies the selection's character and paragraph formatting; running it again turns it off, `{"sticky": true}` keeps it armed across uses. `edit.pasteFormat` applies the copied formatting to the current selection. |
| Find and replace (editor only) | `edit.find`, `edit.findNext`, `edit.findPrevious`, `edit.replace`, `edit.replaceAll` | The UI uses these. `edit.find` searches from the caret with `{"text": q}` and selects the first match; it accepts Word's match options (`matchCase`, `wholeWord`, `wildcards`, `regex`, …) and an empty text clears results. `edit.findNext`/`edit.findPrevious` step through the stored query. `edit.replace` replaces the selected match and moves on; `edit.replaceAll` replaces every match in one undo step with `{"text": q, "with": r}`. |

Table-border color/thickness, row/column insertion/deletion, file creation, automatic save, and Keep/Undo review are not exposed by these tools. Do not promise those capabilities. Protection and native Track Changes disable agent block mutation. Discover missing capabilities rather than attempting raw XML or filesystem workarounds.

All ordinary document toolbar commands remain available in the editor. The tool's smaller surface intentionally provides deterministic targeting and bounded output. Large image payloads can exceed the tool input limit; do not inline oversized base64 data into model context.

The shared document adapter also exposes `content.find`, `content.replace`, `content.delete`, and `content.append` through the CLI and the `aiOfficeEdit` tool. They locate content by text and support Word, PDF and PowerPoint. The Word tools described above use internal block identities and revision checks for edits to the live document.

## Tables

Use `office_document_help` with the exact command to load its parameters. Run table commands through `office_document_edit` (`operation:command`) on a current paragraph ID **inside the intended cell**. The target selects that cell; a table ID selects all its content. Structural commands need a fresh document version. Native commands preserve block IDs for surviving content and create one Undo step.

- Insert: `table.insertRowAbove`, `table.insertRowBelow`, `table.insertColumnLeft`, `table.insertColumnRight`.
- Delete: `table.deleteRow`, `table.deleteColumn`, `table.deleteTable`. Confirm the requested scope from the user's instruction.
- Move: `table.moveAxis` with `axis:row|column`, `forward:true|false` moves one position. Merged cells must be unmerged first; edge moves are refused without mutation.
- Size: `table.columnWidth` with `width` in points (72 points = 1 inch); `table.rowHeight` with `height` and `rule:atLeast|exact`.
- Layout: `table.autofit` with `mode:contents|window|fixed`; `table.distributeColumns`, `table.distributeRows`; `table.repeatHeader` with explicit `value:true|false`.
- Cells: `table.merge` uses the current cell selection; `table.split` accepts `columns` (default 2). Do not merge without a deliberate cell selection.
- Borders: `table.borders` with `kind:all|outside|inside|none|top|bottom|left|right`, optional `width` in points and six-digit `color`.
- Sort: `table.sort` with zero-based `column`, `descending` and optional `header`. Supply `header:true` to protect a header row when it is not marked as repeating in DOCX.
- Cell formatting: `table.cellAlign` with `value:topLeft|topCenter|topRight|centerLeft|center|centerRight|bottomLeft|bottomCenter|bottomRight`; `table.shading` with six-digit `color` or `null` to clear.
- Selection: `table.selectCell`, `table.selectRow`, `table.selectColumn`, `table.selectTable` are interactive commands. For agent edits, address the paragraph in the intended cell directly.
- Pointer gestures: `table.selectAt`, `table.insertAt`, `table.fitColumn` use `page,x,y,reach`; `table.resizeDrag`, `table.moveDrag` accept zero-based `page` and page-point `x,y,toX,toY`. These support editor handles; prefer semantic commands for agents, rather than guessing coordinates.

Table handles and context menus only expose geometry for rendered pages. They do not create an HTML copy of the table or bypass native review/history.

## Images

Image inspection returns `images` on its containing paragraph: `offset` is a Unicode character offset, not a persistent standalone ID. Use a fresh paragraph ID and version with `operation:command`, `start:offset`, `end:offset+1`. This exact range is required for image edits and prevents modifying the wrong picture when a paragraph contains several.

- `picture.size`, `picture.crop`, `picture.change`, `picture.delete`, `picture.reset`: resize, crop, replace, remove or reset the exact selected image. Inspect help for parameters. For insertion/replacement, `params.fileUri` resolves a local/workspace image file in the host without putting base64 in model context; omit `data` and `path`. Image files and each batch are limited to 16 MB; smaller files are preferable.
- `picture.altText`, `object.altText`, `picture.border`, `picture.style`: accessibility and appearance. Alt text supports description, title and decorative status through `object.altText`.
- `arrange.wrap`, `arrange.wrapText`, `arrange.moveWithText`, `arrange.position`, `arrange.align`, `arrange.bounds`, `arrange.nudge`, `arrange.rotation`, `arrange.rotate`, `arrange.bringForward`, `arrange.sendBackward`: wrapping, anchor behaviour, positioning, alignment, size and rotation. These use native points. Moving an inline image with `arrange.bounds` makes it floating.
- `references.caption`: a native numbered Figure/Table caption; use `label`, `text` and `position:above|below`. Captions are normal editable document content with native numbering fields.
- `picture.select` is an editor-only hit-test command. Agents use paragraph ID and exact image start/end, never screen coordinates.

At the end of a document, use `office_document_edit` with `operation:append`, blocks and a fresh version, without id or position. The engine resolves the true body end. Use `office_document_read` with `operation:tail` to inspect the final top-level block. Never guess the end from a paginated read. Keep the returned document URI for subsequent calls and verify the inserted block's text in that same document before reporting success.
