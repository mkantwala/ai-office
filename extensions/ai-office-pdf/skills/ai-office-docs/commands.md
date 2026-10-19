# Discover exact capabilities on demand

Prefer `office_document_format` for common character and paragraph formatting. It resolves all literal matches and actual style sizes in one atomic operation. Use the native commands below for other capabilities.

`office_document_help` without arguments returns categories. Pass `category` for a small related set or `command` for a single exact parameter schema from the running engine. The running registry is authoritative. This index covers all commands currently enabled by the Word bridge; availability can change with engine versions.

| Category | Enabled commands | Agent usage |
|---|---|---|
| Text | `text.insert`, `text.newParagraph`, `text.lineBreak`, `text.tab`, `text.backTab`, `text.backspace`, `text.delete` | Prefer structured insert/update/delete. Targeted text commands act at the end of the target paragraph range; backspace can join paragraphs. |
| Character formatting | `format.bold`, `format.italic`, `format.underline`, `format.font`, `format.size`, `format.strikethrough`, `format.color`, `format.highlight`, `format.clear` | Use targeted command; send explicit values where supported rather than toggling unknown formatting. |
| Paragraph formatting | `para.properties`, `para.align`, `para.style`, `para.bullets`, `para.numbering`, `para.lineSpacing` | Prefer `office_document_format` with `format.paragraph` for a typed partial patch; `para.properties` accepts that same object on a paragraph/heading/list-item ID. Read the exact block for its effective `paragraph` properties. List items retain their paragraph identity. |
| Insertion | `insert.table`, `insert.picture`, `insert.link` | Use targeted command; prefer structured table insertion when populating cells. Pictures require image data and engine-specific size parameters; get help. Inline images/links are addressed through their containing paragraph, not independent block IDs. |
| Review | `review.newComment`, `review.comments` | Add comments through targeted commands. Comments can also be inspected through their document story. These commands do not implement acceptance/rejection of agent changes. |
| Caret (editor only) | `caret.set`, `caret.left`, `caret.right`, `caret.up`, `caret.down`, `caret.home`, `caret.end`, `caret.docStart`, `caret.docEnd` | The UI uses these. Agent tools use IDs and ranges instead; targeted command rejects them. |
| Selection (editor only) | `select.all`, `select.range`, `select.word`, `select.text`, `select.collapse` | Read `operation:selection` for user context. Targeted command rejects these to avoid losing the target. |
| History (editor only) | `edit.undo`, `edit.redo` | Normal editor Undo/Redo handles complete agent operations. These commands are excluded from targeted edits because they cannot be scoped to a block. Use the separate `office_document_history` tool for status, undo, redo or restoring a request checkpoint. |

Table-border color/thickness, row/column insertion/deletion, standalone image resizing, file creation, automatic save, and Keep/Undo review are not exposed by these tools. Do not promise those capabilities. Protection and native Track Changes disable agent block mutation. Discover missing capabilities rather than attempting raw XML or filesystem workarounds.

All ordinary document toolbar commands remain available in the editor. The tool's smaller surface intentionally provides deterministic targeting and bounded output. Large image payloads can exceed the tool input limit; do not inline oversized base64 data into model context.

The shared document adapter also exposes `content.find`, `content.replace`, `content.delete`, and `content.append` through the CLI and the `aiOfficeEdit` tool. They locate content by text and support Word, PDF and PowerPoint. The Word tools described above use internal block identities and revision checks for edits to the live document.

## Tables

Use `office_document_help` with the exact command to load its parameters. Run table commands through `office_document_edit` (`operation:command`) on a current paragraph ID **inside the intended cell**. The target selects that cell; a table ID selects all its content. Structural commands need a fresh document version. Native commands preserve block IDs for surviving content and create one Undo step.

- Insert: `table.insertRowAbove`, `table.insertRowBelow`, `table.insertColumnLeft`, `table.insertColumnRight`.
- Delete: `table.deleteRow`, `table.deleteColumn`, `table.deleteTable`. Confirm the requested scope from the user's instruction.
- Move: `table.moveAxis` with `axis:row|column`, `forward:true|false` moves one position. Merged cells must be unmerged first; edge moves are refused without mutation.
- Size: `table.columnWidth` with `width` in points (72 points = 1 inch).
- Sort: `table.sort` with zero-based `column`, `descending` and optional `header`. Supply `header:true` to protect a header row when it is not marked as repeating in DOCX.
- Cell formatting: `table.cellAlign` with `value:topLeft|topCenter|topRight|centerLeft|center|centerRight|bottomLeft|bottomCenter|bottomRight`; `table.shading` with six-digit `color` or `null` to clear.
- Selection: `table.selectRow`, `table.selectColumn`, `table.selectTable` are interactive commands. For agent edits, address the paragraph in the intended cell directly.
- Pointer gestures: `table.resizeDrag`, `table.moveDrag` accept zero-based `page` and page-point `x,y,toX,toY`. These support editor handles; prefer semantic commands for agents, rather than guessing coordinates.

Table handles and context menus only expose geometry for rendered pages. They do not create an HTML copy of the table or bypass native review/history.
