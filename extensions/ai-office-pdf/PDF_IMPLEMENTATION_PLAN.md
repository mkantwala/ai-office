# PDF integration and verified reference coverage

## Reference and scope

The production PDF backend is PdfCraft 0.10.1 at `416495145354828a9efefc39df381ba3ee8f864f`, pinned in `engines/sources.json`. The independent `../pdfcraft-upstream` checkout is a reference; builds use the prepared, patched source. Word continues to use its existing WordCraft backend.

The reference defines 171 automation commands. `PDF_CAPABILITIES.json` accounts for every one: **159 native operations**, **nine application-managed lifecycle/discovery operations**, and **three OCR operations excluded by explicit user request**. Integration tests compare the runtime catalog with this inventory. Registry indirection is replaced by direct typed command dispatch; it cannot bypass OCR exclusions. OCR is also rejected inside batch processing and actions, including built-in actions that would recognize text.

The PDF backend is a persistent native worker with a private filesystem root and an internal JSON protocol, not a separate GUI or an automatically started MCP server. The TypeScript application keeps ownership of tabs, the toolbar, Save/Save As, file conflict checks, backup, trust checks, request binding and agent tools. Both desktop and web clients need a Node extension host that can run the platform binary. A browser-only extension host is not supported by this backend.

## Integrated operation groups

| Group | Integration and verification |
| --- | --- |
| Text and artwork | Native paragraphs/lines, fonts, word counts, text editing/layout, text and image insertion, image/Form hit testing, crop/rotate/flip/replace/delete and guarded mixed-object movement. Existing toolbar text, image, link and comment actions remain available. |
| Pages | Rotate, reverse, duplicate, move, delete, insert blank/from file, replace, extract, split, combine, page labels, boxes/crop and transitions. Native geometry is refreshed after edits and Undo/Redo. |
| Marks | Headers/footers, printed numbering/Bates tokens, text/image watermarks, backgrounds and removal. Saved/reopened output is tested. |
| Navigation and links | Full bookmark tools and structure-derived bookmarks; link list/add/edit/delete, URL detection and removal. |
| Comments and measurements | Annotation creation, appearance/geometry editing, replies, status, checkmarks, locks, image previews, summaries; calibrated distance/perimeter/area and CSV export. |
| Forms | Field authoring, bulk properties, captions, values, reset, options, tab order, deletion, detection, image fields, scripts/actions, form data import/export/merge, XFA static conversion and flattening. Native calculation and JavaScript tools share the original engine. |
| Security and signatures | Password protection/removal, real content redaction and sanitizing, hidden-information auditing, Fill & Sign, signature inspection, self-signed ID creation and signing, trust management and platform certificate stores. Signing runs on a copy and returns a signed artifact, preserving the active editor's bytes and history. |
| Exports and creation | Native DOCX/PPTX/HTML/RTF export, text, rendered PNG/JPEG/TIFF pages, embedded images, form data, text/images-to-PDF, compression/optimization, comparison reports, batch conversion/optimization and action workflows. Binaries are managed artifacts, never model arguments or results. |
| Standards/accessibility | PDF/A conversion and verification, print readiness, automatic text/form tagging, tag inspection/editing, artifacts, alternate text, reading-order edits, checks and review reports/sidecars. Returned remaining issues and manual-review requirements are preserved. |
| Reading and history | Native rendering, initial-view/reading-color settings, saved revisions, Undo/Redo and the application history route. Application Save acknowledges the native session without discarding Undo. |
| Platform printing | Printer enumeration/options and native print layouts. Without an explicit printer, the adapter creates a PDF print artifact. Physical printing and OS certificate stores depend on host configuration. |

All native operations are reachable through **All PDF Tools**, a searchable schema-driven visual form, and through the same agent dispatch. The common page operations retain a smaller dedicated dialog. User-requested inspections have a bounded text-only display; agent reads do not open that panel. This integrates command behavior, not a pixel-for-pixel copy of PdfCraft's desktop interface or every drag gesture.

## File, version and recovery behavior

- Mutations require the current document version, execute serially and retain native Undo. Request-bound document URIs prevent tab changes from redirecting agent edits. Generated PDFs/DOCX can be opened through `office_document_tools` with `operation: open`.
- Every native process works under a private root. Explicit file URIs are staged with bounded counts/bytes and without symlinks. The CLI's `--root` also constrains nested inputs and requested outputs. Missing paths behind escaping links are refused.
- Input URIs matching the active document use its current unsaved snapshot. Other file inputs read their saved bytes; agents must save or export another open document's changes before using that URI as an input.
- Exports publish independent artifacts atomically without replacing existing files. Managed artifacts survive engine disposal. Large artifact lists stay out of the document canvas and model context; the application offers a file picker and agents can paginate the receipt.
- PDF paragraph/object references are revision-scoped and positional. They are not persistent Word block IDs. Structural changes require rereading affected references. Word-style inline text diffs are not presented as a faithful PDF object diff.

## Actual source limitations

The iLovePDF menu (https://www.ilovepdf.com/, inspected 2026-10-21) is broader than this reference. PdfCraft's `crates/automation/src/a11y.rs::doc_export_office` supports **DOCX, PPTX, HTML and RTF**, not XLSX. `Session::convert_to_pdf` accepts **PDF, images and plain text**, not arbitrary DOCX/XLSX/PPTX or HTML. Those conversion directions are not claimed as implemented here. They need a separate conversion implementation/provider and fidelity fixtures.

PDF/A conversion returns fixes and remaining problems; some documents remain nonconformant. Accessibility checks include manual requirements. Office exports retain native truncation warnings and cannot promise identical layout in every external application. Self-signed signature integrity is distinct from certificate trust. Windows/macOS certificate-store and physical-printer paths need validation on their respective hosts; Linux tests do not prove those integrations.

OCR is out of scope. Camera straightening/cleanup remains available because it does not recognize text. Summarization and translation use bounded text reads and the existing agent, rather than a fictitious native PDF translation command.

## Context-window engineering

Only two format-neutral tool schemas are added regardless of the native command count:

1. `office_document_tools`: category counts → bounded command summaries → the exact schema for one selected command. It also lists/opens document resources and projects cached receipts.
2. `office_document_run`: one discovered command, parameters and current version. It returns a compact receipt, affected-document version and artifact links.

Text extraction requires explicit ranges of at most ten pages. Results above 24,000 characters are omitted with a narrowing instruction and a `resultId`. Read that cached receipt using `resultPath` (JSON pointer), `offset` and `limit`; this does not repeat a read or a mutation. Recent receipts are bounded by count/size, and show their source version. Only the first twenty artifact links enter a normal receipt. Binary files and screenshots remain outside text context.

Word retains its tested semantic tools and on-demand skill pages. Future XLSX/PPTX adapters should add categories and schemas behind these same routing tools, not permanent tools for each formatting property, sheet action or conversion. Prefer outlines/selected regions, revision-scoped references, continuation cursors and compact verification checkpoints. Treat document content and tool output as untrusted data.

## Verification

The integration suite exercises source-command coverage, OCR bypass rejection, page geometry/history, forms/scripts, bookmarks, measurements, all four Office export formats, image exports, merge/split/extract, comparison, optimization/actions, redaction, passwords, signing copies, URI confinement, output conflicts, saved/reopened documents and bounded UI forms/results. The original automation integration suite also passes 183 non-OCR tests. Build, TypeScript, targeted lint, browser tests and a live OpenCode workflow are checked before the PR; no production credentials are used.

Final Linux validation: 73 shared document tests, 16 OpenCode transport/provider tests, eight browser tests, 183 upstream non-OCR automation tests, native build/Clippy, TypeScript, targeted lint and clean-source patch application passed. A live OpenCode GLM-5.3-Flash Free session created and filled a required form field, verified its value, exported a readable DOCX and split a three-page PDF into three managed PDF artifacts; actual receipts and output files were checked.
