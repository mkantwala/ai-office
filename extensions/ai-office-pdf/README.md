# AI Office documents

DOCX opens in Wordcraft's full Rust web application: ribbon tabs, dialogs, rulers, editable document canvas and native undo/redo. PDF uses the Rust Pdfcraft engine with a compact toolbar. Word's full UI runs as WebAssembly in the webview; PDF and the headless Word CLI use extension-host worker engines.

Click in Word to position the caret; drag to select text. In PDF, click a paragraph to edit it and click outside or press Ctrl+Enter to finish. Ctrl+S saves, Ctrl+Z undoes, and Ctrl+Shift+Z redoes. Comments appear in the Comments disclosure above the document. VS Code provides save, Save As, revert and crash-recovery backups. Changes made externally reload clean documents; conflicting saves are refused so you can use Save As.

Word exposes the existing Wordcraft web features across Home, Insert, Draw, Design, Layout, References, Mailings, Review, View and Developer. It retains upstream feature and file-fidelity limitations. Host file operations and recovery use VS Code; Print PDF exports a PDF. Ribbon Paste reads host clipboard text. Desktop-only integrations remain unavailable in the web application. PDF supports text insertion, paragraph formatting, images, links and comments. PDF lists, heading styles and tables are disabled; scanned pages need OCR before their text can be edited.

All toolbar document operations are available through the [shared CLI](../ai-office-common/README.md). The CLI reads and writes files; the registered `aiOffice.word.execute` and `aiOffice.pdf.execute` VS Code commands operate on open documents using `(uriString, commandId, params)` and participate in the same undo/save flow. A CLI edit to a file with unsaved in-app changes requires resolving the resulting conflict.

Build from the VS Code repository using Node 24:

```powershell
node extensions/ai-office-pdf/build.mts
node extensions/ai-office-pdf/build.mts --word
node extensions/ai-office-pdf/build.mts --word-full
```

Prepare pinned sources with `node scripts/prepare-office-engines.mts` (sibling checkouts also work). Requires Rust wasm32-unknown-unknown and wasm-bindgen 0.2.129 at `.build/pdf-tools/bin` (or WASM_BINDGEN). The shared worker and CLI also require dependencies installed in `extensions/ai-office-slides`. PDFCRAFT_SOURCE, WORDCRAFT_SOURCE and CARGO_TARGET_DIR override paths. `--ui-only` rebuilds TypeScript; `--package-only` packages existing Rust builds. Notices accompany generated engines. Opening a document does not save it.
