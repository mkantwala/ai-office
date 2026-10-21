# AI Office documents

DOCX and PDF open directly on a page canvas with a compact editing toolbar. Word uses the Rust Wordcraft engine for hit testing, selections, typing, pagination and formatting. PDF uses the Rust Pdfcraft engine and its headless automation commands for paragraph editing and annotations. Word runs in an extension-host worker thread and PDF in a persistent native worker; the webview contains only the document surface and controls.

Click in Word to position the caret; drag to select text. In PDF, click a paragraph to edit it and click outside or press Ctrl+Enter to finish. Ctrl+S saves, Ctrl+Z undoes, and Ctrl+Shift+Z redoes. Comments appear in the Comments disclosure above the document. VS Code provides save, Save As, revert and crash-recovery backups. Changes made externally reload clean documents; conflicting saves are refused so you can use Save As.

Word supports heading styles, font, size, bold, italic, underline, alignment, bullets, numbering, spacing, tables, images, links and comments. PDF supports text insertion, paragraph text/font/size/emphasis/alignment/spacing, images, links and comments. PDF lists, heading styles and tables are disabled. PDF formatting applies to the paragraph. OCR is intentionally excluded. The PDF Tools dialog exposes the remaining native editing and file commands.

All toolbar document operations are available through the [shared CLI](../ai-office-common/README.md). The CLI reads and writes files; the registered `aiOffice.word.execute` and `aiOffice.pdf.execute` VS Code commands operate on open documents using `(fileNameOrUri, commandId, params)` and participate in the same undo/save flow, and the `aiOfficeEdit` language model tool exposes the same commands to chat agents (it is contributed by this extension and dispatches to any AI Office kind). Without a file argument both edit the active document, and a bare file name targets another open document. A CLI edit to a file with unsaved in-app changes requires resolving the resulting conflict.

Build from the VS Code repository using Node 24:

```powershell
node extensions/ai-office-pdf/build.mts
node extensions/ai-office-pdf/build.mts --word
```

Run `node scripts/prepare-office-engines.mts` first to prepare the pinned sources. PDF requires a native Rust toolchain; Word requires Rust wasm32-unknown-unknown and wasm-bindgen 0.2.129 at `.build/pdf-tools/bin` (or WASM_BINDGEN). The shared worker and CLI also require the dependencies installed in `extensions/ai-office-slides`. PDFCRAFT_SOURCE, WORDCRAFT_SOURCE and CARGO_TARGET_DIR override source/target paths. `--ui-only` rebuilds TypeScript; `--package-only` packages existing Rust builds. Notices accompany generated engines. Opening a document does not save it.

## Word agent

The existing chat panel uses **AI Office** and one Cloudflare Workers AI model, `@cf/zai-org/glm-4.7-flash`. No GitHub/Copilot login is needed. Run **AI Office: Configure Cloudflare** to store your account ID and Workers AI token in VS Code SecretStorage, or set `CF_ACCOUNT_ID` and `CF_API_TOKEN` in the environment before launching the app. Stored credentials take precedence. Configuration, activation and model discovery never send a test request; inference starts only when you send a chat message. Credentials are never workspace settings, model input, browser/webview content or tracked files. The remove-credentials command clears SecretStorage; unset environment variables separately.

GLM-4.7-Flash supports tool calling. [Cloudflare's published pricing](https://developers.cloudflare.com/workers-ai/models/glm-4.7-flash/) is $0.0605/M input tokens and $0.40/M output tokens as checked October 17, 2026. Thinking is disabled, output is capped at 4,096 tokens per request, and a user turn is capped at 16 model rounds. No automatic retries or model fallback occur. Conversation input is bounded to six prior text turns plus relevant tool results, with a 120,000-byte request budget. Stop cancels pending inference and prevents subsequent edits; already completed edits remain undoable.

Open a DOCX and ask the agent to write or edit it. Four structured tools operate on the live editor: `office_document_read`, `office_document_edit`, `office_document_help`, and `office_document_activity`. Paragraphs (including headings/list items), tables, rows and cells have internal IDs. Inserts return the new root IDs; reads expose children and paged text. Updates use Unicode code point ranges. Each edit requires a fresh document version and creates one ordinary Undo step. Changes use normal Save/Auto Save and backup handling. A shimmer marks the target while the agent works and clears after edits, cancellation or expiry.

The [internal document skill](skills/ai-office-docs/SKILL.md) is loaded automatically by this participant. Its short root describes the workflow; detailed CRUD, command and identity guides are loaded through `office_document_help` only as needed. Exact engine command schemas are discovered on demand. Other compatible agents can use the registered language-model tools and contributed skill.

IDs survive normal save/reopen and Undo/Redo via private DOCX metadata. External package changes invalidate that mapping conservatively. Encrypted files currently retain IDs only within the open session. Inline images and links use the containing paragraph/range; they do not yet have standalone block IDs. Arbitrary file creation and unsupported table border/row/column operations are not implemented by these tools.

Agent edits appear in the standard chat changed-file list, with native **Keep**, **Undo**, **Keep All**, and **Undo All** controls. Open a changed file to review text and formatting differences, then decide on each paragraph/table change, the file, or all files. Word also shows red struck-through deletions, green additions and per-change Keep/Undo directly in the original document, using view zones that do not alter its saved pagination. Shared structural/annotation changes are reviewed together as one coherent document change, with expandable formatting details. Conflicting later manual edits prevent an unsafe Undo. Private native snapshots are stored in local extension storage, outside the document and repository; a saved document or matching backup restores its pending inline review after reload. A mismatched external file never receives an old review. **History** opens timestamped document versions for comparison with the current version; the normal diff editor supports inline and side-by-side views. Versions retain up to 30 snapshots, pruning oldest entries above 64 MB (a single latest snapshot may exceed that limit). No Git repository is required. Kept content still uses normal Save/Auto Save. Checkpoint restore remains a separate request-history operation.

Run `node extensions/ai-office-common/test.mts --word` from the repository root for the Word/provider/Cloudflare checks without requiring the PDF engine bundle.

Validation: synthetic DOCX integration tests cover CRUD, identity persistence, atomic rollback, stale requests and Unicode ranges; provider tests cover cancellation, workspace trust and undo; browser tests cover shimmer geometry and reduced motion. Cloudflare transport and tool-loop tests inject local fake responses and never use real credentials or the network.

## OpenCode agent sessions

Select **OpenCode** in the agent/session picker, or run **OpenCode: New Agent Session** from the Command Palette. OpenCode runs through its official [Agent Client Protocol](https://opencode.ai/docs/acp/) (`opencode acp`) in the existing chat UI. The removed terminal panel is not required.

Install OpenCode on the machine hosting the extension. It uses your existing OpenCode providers, authentication and default model; configure those once with OpenCode itself (`opencode auth login` in a system terminal when needed). Sangam One/GitHub login is optional. **OpenCode: Configure Connection** sets `aiOffice.opencode.path` if the executable is not on PATH. The executable setting is machine-scoped and cannot be supplied by a workspace.

Responses, reasoning and tool activity stream into chat. OpenCode's permission requests show the actual options and tool input; cancellation dismisses pending permission choices and stops the turn. A process that fails to stop is terminated. Model and agent-mode choices come from OpenCode's session configuration after connecting, including its configured model providers. Workspace trust is required before starting the local agent.

The `sangam-docs` MCP bridge exposes the existing document tools plus the two bounded PDF discovery/execution entry points. Calls carry the current chat request's opaque tool token, preserving guarded block edits, native history, request-level Keep/Undo and agent shimmer. Tools cannot run through the bridge while the chat is idle or cancelled. The bridge binds only to loopback and uses an ephemeral credential passed through the child environment. No model credentials are copied into this extension or persisted by it.

OpenCode owns durable conversation history. Sangam One stores only its session ID, folder and display metadata, then reloads the conversation through ACP. Empty sessions start no process; idle processes and tool bridges shut down after two minutes and reconnect on the next request. Session metadata is bounded to the latest 200 entries.

Build the extension with `node extensions/ai-office-pdf/build.mts --ui-only` from the repository root, or through the normal built-in extension esbuild pipeline. The build includes the small stdio MCP adapter and dependency licenses. Run offline integration checks with `node extensions/ai-office-pdf/opencode/test.mts`; they cover the wire protocol, approvals, cancellation, failures, native session registration, document-tool routing, model selection and restoration.

## PDF integration

The PDF Tools button in the application formatting toolbar opens scoped page operations: rotate, duplicate, delete, move, insert a blank page, crop, text watermark, printed page numbering and background color. They use the same PdfCraft automation commands exposed to agents and retain native document Undo/Redo. A dialog captured before another edit refuses to apply to the changed document.

`office_document_tools` lists category counts, bounded command summaries and one requested command schema. `office_document_run` executes a discovered PDF command using the current document version. The engine command table stays out of the permanent model tool list. Text reads require explicit ranges of at most ten pages; oversized results return a narrowing instruction. OpenCode and the built-in agent can use both entry points. Existing Word semantic tools remain available.

The reference checkout is `../pdfcraft-upstream`. Builds use the pinned source and checked patch in `engines/sources.json`. See [PDF implementation plan](PDF_IMPLEMENTATION_PLAN.md) for the iLovePDF feature inventory, conversion and file-workflow gaps, PDF review work and the format-neutral tool design for future spreadsheets and presentations. This milestone does not implement complete iLovePDF parity.

Validation:

```sh
node extensions/ai-office-pdf/build.mts
node extensions/ai-office-common/test.mts
OFFICE_BROWSER_PATH=/usr/bin/chromium node --experimental-transform-types --test extensions/ai-office-common/pdfTools.test.mts
```

The complete integration supersedes the first milestone's seven/55-command WASM bridge: [PDF capability inventory](PDF_CAPABILITIES.json) accounts for all 171 reference commands (159 native, nine application-managed, three excluded OCR commands). **All PDF Tools** exposes the native schemas through searchable visual forms. Bounded user inspections and managed artifact links use the same engine as agents.

Build the platform worker with `node extensions/ai-office-pdf/build.mts --native` after `node scripts/prepare-office-engines.mts`. It stays in the Node extension host; there is no second application window or additional MCP server. Generated binaries are ignored. See the implementation plan for URI staging, saved/unsaved source semantics, printer/certificate platform requirements and conversion directions absent from the reference itself.
