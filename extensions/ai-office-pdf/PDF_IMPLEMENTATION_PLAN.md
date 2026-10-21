# PDF integration and format-neutral agent tools

## Goal and reference sources

Integrate PdfCraft into Sangam One while retaining the application toolbar, worker isolation, native document Undo, guarded agent requests, file conflict checks and deferred saving used by Word. PDF is a fixed-layout format: paragraphs and artwork use page geometry and revision-scoped references; do not manufacture Word block identities or reflow semantics.

Reference checkout: `../pdfcraft-upstream` (https://github.com/storytold/pdfcraft). The build continues to use the revision and patch in `engines/sources.json`; the reference checkout is not a build input. PdfCraft is MIT OR Apache-2.0; retain its notices and attribution. Do not copy ArtCraft trademarks.

Feature inventory source: https://www.ilovepdf.com/ (opened 2026-10-21). Its current menu includes merge, split/extract/remove/organize pages, compression, repair, OCR, scans, image and Office conversions, HTML conversion, PDF/A, editing, forms, signing, watermarks, rotation, numbering, crop, unlock/protect, comparison, summarization, translation, Markdown and custom workflows. This inventory is a target, not a claim of completed parity.

## Delivery sequence

| Area | Current integration | Next verification or implementation |
| --- | --- | --- |
| Rendering, text editing, image insertion, comments, links, save/reopen, Undo/Redo | Existing PDF integration; retained | Continue PDF regression coverage alongside Word |
| Rotation, duplication, deletion, moving, blank pages, crop | First integration milestone | Native round trips, dimensions, history, invalid scopes, browser controls |
| Watermarks, backgrounds, printed page numbers/header/footer | First integration milestone | Export/reopen and native geometry checks |
| Native metadata, forms, comments, links, object edits, redaction, hidden information | Scoped engine commands exposed through discovery and guarded run | Dedicated visual tools and command-specific end-to-end tests; do not claim all UI workflows complete |
| Merge, split, extract, replace pages | Engine supports them; existing file-based automation cannot directly access host files from WASM | Add host-mediated URI input/output and byte-based engine adapters. Preserve bookmarks, forms, links and layers. No document bytes through model context |
| Compress and optimize | Native engine has optimizer/export support | Return an output artifact through host-managed export, compare sizes, retain the original if output grows |
| Password protection/unlock | Engine supports both; not enabled by the new generic command surface | Passwords through secret input; verify encrypted exports and permission enforcement without echoing secrets |
| OCR and scan cleanup | Native engine has OCR and camera enhancement | Confirm WASM/runtime compatibility and language-model assets, download budgets, cancellation and scanned-page round trips |
| PDF to/from images, Word, PowerPoint, Excel, HTML, Markdown | Some native exporters exist; spreadsheet editor not yet integrated | Add explicit capability checks per direction, fidelity fixtures and optional conversion providers. Never represent extracted text as a faithful Office conversion |
| PDF/A and repair | Engine has preflight and repair behavior | Independently validate conformance and repaired output; expose unsupported cases honestly |
| Fill/sign and certificate signatures | Native form/signing support | Separate drawn/typed marks from cryptographic signatures; OS/keystore signing needs host adapters |
| Compare | Native comparison exists | Host-mediated two-file inputs and page-level review; preserve user edits |
| Summarize, translate, reusable workflows | Agent layer can orchestrate discovered capabilities | Bounded text reads, citations/page references, explicit translation outputs and durable cancellable jobs |
| PDF Keep/Undo review and live previews | Existing native Undo and agent activity retained; Word inline review remains Word-specific | PDF page/object before-and-after review with stable revision-scoped anchors; do not reuse Word text diffs for flattened PDF artwork |

Each milestone must pass focused engine export/reopen, command, cancellation/trust/conflict and visual UI tests before being marked complete. Operations exposed by the engine remain subject to its own format and permission limitations.

## Context-window engineering

Use **two format-neutral tool schemas** as the extensible boundary:

1. `office_document_tools`: return the document URI, version and category counts. A selected category returns a bounded page of short command summaries. Only an explicitly requested command returns its JSON schema.
2. `office_document_run`: execute that command with its parameters and current version. Large results are omitted with a narrowing instruction, never returned as broken or truncated JSON.

PDF engine command definitions remain internal data, not hundreds of registered LLM tools. Mutations are serialized, version guarded, trust checked, cancellable before execution and use native history. No filename guessing or silent format fallback. Word retains its tested semantic tools in this milestone; generic run deliberately refuses Word mutations.

For XLSX/PPTX, implement adapters under this same discovery contract with format-specific categories and typed parameters. Do not register a tool per cell, formatting property or conversion. A future compact adapter can route Word semantic operations through the two-entry surface while retaining the current typed implementations and tests.

Keep skills as a short routing index with reference pages loaded only when needed. Read PDF text by explicit page ranges (at most ten per call); use search to identify relevant pages. For tables/sheets, return bounded regions and continuation tokens. For presentations, return slide outlines then selected objects. Revision-scoped IDs and versions, not whole-document rereads, carry state between calls.

Keep binary inputs, rendered pages and generated exports in host-managed artifacts referenced by URI. Limit screenshots to changed pages for appearance verification. Checkpoint completed work with compact receipts (URI, revision, changed pages/object references, verification and artifact links), not repeated document text. Tool results and document content remain untrusted data.

OpenCode and the built-in agent receive the compact discovery/run entry points alongside existing Word tools. A later schema-on-demand transport can advertise only routing tools for all formats once Word semantic dispatch and mixed-format session migration are verified. This milestone reduces growth from PDF capabilities; it does not claim that existing Word schemas have already been removed from context.
