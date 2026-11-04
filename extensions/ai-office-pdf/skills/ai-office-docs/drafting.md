# Progressive drafting

Use `office_document_draft` for a new multi-section passage in an open Word document. Read an outline and the relevant insertion target first; never recreate a document or rewrite unrelated content.

1. `operation:plan`, fresh `version`, existing paragraph/table `id`, optional `position:before|after` (default after), and `sections:[{title,intent,level?}]`. Plan 1–20 concise sections; heading levels 1–3, default 2. Intent says what is coming or what you must research. The user sees pending cards immediately.
2. Research only what the next section needs. Briefly tell the user which section you are working on.
3. `operation:fill`, fresh `version`, returned `sectionId`, and 1–49 body `blocks` using the content guide’s paragraph/heading/table schema. The planned heading is inserted automatically; do not repeat it. One section per call, one atomic undoable change. Use the receipt’s version/created IDs afterward.
4. Repeat as sections are ready. Use `status` after interruption; do not repeat completed fills. Cancellation or a rejected fill leaves its pending card in place and does not write partial blocks.

Sections may be filled out of order; the app keeps their planned order using neighboring filled blocks. A deleted/undone anchor makes the fill fail safely: read current content, discard remaining pending cards, and replan. Do not guess an anchor.

Pending cards are temporary editor state, not DOCX paragraphs and not versioned file content. Save/export contains only filled content. Closing/reloading the document clears pending intentions. `discard` with sectionId removes one pending card; without it clears the plan. It never removes already written content. Use normal editing/history for content changes. Completed content participates in shimmer, local history and per-change/file/all-files Keep/Undo.

Use native document structures, not a Markdown roundtrip. A short requested edit needs no draft plan. Choose tables for compared attributes and diagrams/charts only when supported capabilities justify them; do not invent drawing tools. Keep IDs private.
