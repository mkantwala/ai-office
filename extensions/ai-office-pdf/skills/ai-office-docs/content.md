# Content operations

Read results contain `uri`, `version`, and paginated blocks with `id`, `kind`, `parentId`, `style`, a text slice, `textLength`, `textOffset`, `nextTextOffset`, and child count. `nextOffset` continues the block list. Selection reads return `content`, `selection`, and bounded `selectedText`. A selection may be backwards or span multiple blocks; never assume its start/end describe one paragraph.

## Format

Prefer `office_document_format` for formatting: `{query:"literal text",scope:"all",format:{underline:true,sizeFromStyle:"Heading 1"}}`. Queries are literal and case-insensitive; `scope:"first"` selects only the first occurrence. One call resolves paginated targets, Unicode ranges and actual style sizes, then commits every property atomically. No preliminary read is needed for a query. `target:"matches"` formats matching text; `target:"paragraphs"` formats containing paragraphs. Current block `ids` can narrow targets but require a fresh version.

Supported properties: `bold`, `italic`, `underline` (explicit booleans), `font`, `size` (points), `sizeFromStyle`, `paragraphStyle`, `alignment`, `lineSpacing`. Use either size or sizeFromStyle. Copying size does not apply the paragraph style. Successful receipts verify scope; do not repeat formatting or reread unchanged targets. Unsupported or oversized plans fail before editing.

## Create

Use `operation: insert`, an existing paragraph/table `id`, `position: before|after`, and a `blocks` array. New IDs are assigned by the application. Existing IDs survive insertion. To insert inside a cell, target one of its paragraphs.

- Paragraph: `{kind:"paragraph", text:"…", style:"Normal"}`.
- Heading: `{kind:"heading", text:"…", style:"Heading1"}` (Heading1–Heading3 are standard choices).
- Table: `{kind:"table", rows:[["Phase","Duration"],["Planning","1 week"]]}`.

Use separate blocks for paragraphs, not newline-separated text. Insert at most 50 blocks per call; tables allow 1–20 rows and 1–10 columns with equal cell counts. Table text can be edited afterward by reading its row, cell, and paragraph children. Use paragraph formatting commands afterward to make list items. Unsupported styles must not be guessed.

## Read

Default reads return small snippets. Use `query` for text search or `parentId:"body"` for top-level content. Use `id` for an exact block and follow `nextTextOffset` for the rest. To inspect a table, read its children, then a row's cells, then a cell's paragraphs. Search results are individual text blocks, so a paragraph inside a table retains its parent cell ID. Secondary stories have parents like `part-…`; do not mix body text with headers, footers or comments.

## Update

Use `operation: update`, a paragraph/heading/list-item ID, and `text`. Without offsets this replaces the paragraph's text. Supply `start` and `end` for a smaller replacement. Offsets count Unicode code points, not UTF-16 code units or bytes; start is inclusive and end exclusive. An empty range inserts text. An empty replacement deletes that range. Preserve inline objects and formatting by targeting only the intended text range; whole-paragraph replacement also replaces its inline contents.

Text updates accept no formatting fields. Prefer `office_document_format` for supported formatting. Use `operation: command` for other formatting, images, links or comments. Fetch the exact schema from help, then provide `command` and `params`. Character formatting accepts optional Unicode `start`/`end` offsets to target a substring. Without them, formatting/review commands target the block's paragraph range; insert/text commands run at the end of that range. Current commands cannot target an arbitrary subset of cells as one noncontiguous selection.

## Delete

Use `operation: delete` for a paragraph or complete table. Deleting a table also deletes its descendants. Read first and verify that this matches the user's intent. The engine retains an empty paragraph when a container needs one. Row and cell IDs support navigation and targeting; structural row/cell deletion is not currently exposed.

Edits require the latest `version`. After each edit, use the returned version or reread. For long documents, write coherent sections in separate calls and check cancellation between calls. Use `operation:batch`, a fresh `version`, and `edits:[{operation,id,...},...]` for 1?200 edits. The entire batch is validated on a private copy and commits as one Undo/Redo step. Nested batches are rejected. Read all matching pages before constructing it; every requested target and action must be represented. `appliedEdits` confirms which operations succeeded, including unchanged targets. Separate tool calls remain separate undo steps. Request checkpoints restore all subsequent document edits, with Redo available afterward. Keep/Undo review is not implemented.

Use `office_document_read` with `operation:styles` and an optional name/ID query to discover the document?s real style sizes. Use `office_document_history` status for available Undo/Redo and request checkpoints; mutations require its fresh version. History changes affect the open document and do not save it.

History also supports `operation:checkpoint` with a fresh version to record the current position without editing the document. Later `restore` takes its internal checkpoint token. New edits after Undo discard the abandoned redo branch; checkpoints beyond the retained native undo history are unavailable.
