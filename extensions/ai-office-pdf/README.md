# AI Office documents

DOCX and PDF open directly on a page canvas with a compact editing toolbar. Word uses the Rust Wordcraft engine for hit testing, selections, typing, pagination and formatting. PDF uses the Rust Pdfcraft engine and its headless automation commands for paragraph editing and annotations. Engines run in extension-host worker threads; the webview contains only the document surface and controls.

Click in Word to position the caret; drag to select text. In PDF, click a paragraph to edit it and click outside or press Ctrl+Enter to finish. Ctrl+S saves, Ctrl+Z undoes, and Ctrl+Shift+Z redoes. Comments appear in the Comments disclosure above the document. VS Code provides save, Save As, revert and crash-recovery backups. Changes made externally reload clean documents; conflicting saves are refused so you can use Save As.

Word supports heading styles, font, size, bold, italic, underline, alignment, bullets, numbering, spacing, tables, images, links and comments. PDF supports text insertion, paragraph text/font/size/emphasis/alignment/spacing, images, links and comments. PDF lists, heading styles and tables are disabled. PDF formatting applies to the paragraph; scanned pages need OCR before their text can be edited. This initial surface does not add OCR or a full office ribbon.

All toolbar document operations are available through the [shared CLI](../ai-office-common/README.md). The CLI reads and writes files; the registered `aiOffice.word.execute` and `aiOffice.pdf.execute` VS Code commands operate on open documents using `(uriString, commandId, params)` and participate in the same undo/save flow. A CLI edit to a file with unsaved in-app changes requires resolving the resulting conflict.

Build from the VS Code repository using Node 24:

```powershell
node extensions/ai-office-pdf/build.mts
node extensions/ai-office-pdf/build.mts --word
```

Requires sibling pdfcraft and wordcraft checkouts, Rust wasm32-unknown-unknown and wasm-bindgen 0.2.129 at `.build/pdf-tools/bin` (or WASM_BINDGEN). The shared worker and CLI also require the dependencies installed in `extensions/ai-office-slides`. PDFCRAFT_SOURCE, WORDCRAFT_SOURCE and CARGO_TARGET_DIR override source/target paths. `--ui-only` rebuilds TypeScript; `--package-only` packages existing Rust builds. Notices accompany generated engines. Opening a document does not save it.
