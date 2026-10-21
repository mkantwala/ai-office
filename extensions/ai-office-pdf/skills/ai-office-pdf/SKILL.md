---
name: ai-office-pdf
description: Read and edit PDFs in Sangam One using bounded discovery and the native PdfCraft command engine.
---

Use `office_document_tools` to discover the active PDF and current version. Select a category, then request only the command schema needed. Execute through `office_document_run`; reuse its returned version for the next mutation. Never load every schema or guess payload shapes. Use `operation: list` to identify open files and `operation: open` with a PDF/DOCX URI to open an artifact.

PDF paragraphs and artwork are positional, revision-scoped references, not persistent Word IDs. Reread them after structural changes. Text extraction takes explicit ranges of at most ten pages. Oversized results return a `resultId`: use discovery with that ID, a `resultPath`, offset and limit. Never repeat a mutation to retrieve its output.

Pass file URIs in file parameters; the host stages bytes and returns managed output artifacts. Omit output locations for generated artifacts. Existing outputs are never replaced. Inputs matching the active PDF include unsaved changes; other inputs use saved files. Keep binaries out of context.

Edits remain unsaved until normal Save and support native Undo. Signing produces a separate signed copy. Report remaining PDF/A/accessibility/export warnings accurately; signature integrity does not imply certificate trust. OCR is excluded. Platform printers and certificate stores require host support. Native Office export supports DOCX/PPTX/HTML/RTF; do not invent XLSX export or Office/HTML-to-PDF conversion.

Verify affected pages and artifacts before claiming completion. Keep IDs internal. Use the existing Word skill for DOCX block edits. See `../../PDF_IMPLEMENTATION_PLAN.md` for the audited capability inventory and platform limits.
