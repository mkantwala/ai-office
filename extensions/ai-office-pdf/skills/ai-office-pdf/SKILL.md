---
name: ai-office-pdf
description: Read and edit an open PDF in Sangam One through bounded capability discovery and native commands.
---

Use `office_document_tools` to discover the active PDF, its version and category counts. Select only the relevant category, then request the exact command schema. Execute with `office_document_run`; use the returned current version for the next mutation. Never load every schema or guess a command shape.

PDF paragraphs, images, forms and page geometry are revision-scoped fixed-layout content. Read them fresh after structural changes. They are not persistent Word paragraph IDs. Text reads use explicit ranges of at most ten pages. Narrow oversized results; do not repeatedly request the whole file.

Native edits remain unsaved until normal Save and support document Undo. Never claim Word-style inline Keep/Undo review for PDFs unless the tool confirms it exists. Verify affected pages and exported/reopened output before claiming format fidelity. Keep binaries out of model context; use host-managed URIs for future file workflows. Unsupported operations must be reported, not simulated with destructive rewrites.

The integration roadmap and known gaps are in `../../PDF_IMPLEMENTATION_PLAN.md`. Use the existing Word skill only for DOCX documents.
