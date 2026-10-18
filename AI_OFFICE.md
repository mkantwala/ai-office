# AI Office

AI Office is a VS Code fork with an office document explorer, direct Word/PDF/PowerPoint editing, a compact toolbar and a Word document agent backed by Cloudflare Workers AI. The extension marketplace is disabled. Markdown uses the existing editor.

Word and PDF use Rust engines compiled to WebAssembly. PowerPoint uses `@office-kit/pptx` and `@office-kit/pptx-preview`. The engines' native application interfaces are not embedded in the editor.

Developer feature entry points and their associated built-ins are excluded from the office workbench. See the [removal map](AI_OFFICE_REMOVAL_MAP.md) for the removed UI, retained API dependencies and validation commands.

Timeline/local history, markers, SCM/Git, multi-diff and comments are retained for document workflows. The [reuse map](AI_OFFICE_REPURPOSE_MAP.md) records their current capabilities, the missing Word adapters and the implementation order.

## Source and build

All AI Office application changes are in this repository. Pinned Wordcraft/Pdfcraft revisions and patches are stored under [engines](engines/README.md), so a fresh checkout does not depend on unpublished sibling-repository changes.

Install the normal VS Code development prerequisites and dependencies described in the upstream contributor documentation. Install the slide extension's dependencies with `npm install --ignore-scripts` inside `extensions/ai-office-slides`. The office extension builds require Node 24, Rust's `wasm32-unknown-unknown` target, and wasm-bindgen CLI 0.2.129.

From the repository root:

```powershell
node scripts/prepare-office-engines.mts
rustup target add wasm32-unknown-unknown
cargo install wasm-bindgen-cli --version 0.2.129 --locked --root .build/pdf-tools
node extensions/ai-office-pdf/build.mts
node extensions/ai-office-pdf/build.mts --word
node extensions/ai-office-slides/build.mts
```

After building the VS Code client, launch with `scripts\code.bat`. For an already prepared development build, set `VSCODE_SKIP_PRELAUNCH=1` to skip the normal preparation step.

## Editing and agents

Click in Word to type and drag to select. Click a PDF paragraph or PowerPoint text box/table cell to edit it; press Ctrl+Enter or click outside to finish. Ctrl+S saves; Ctrl+Z and Ctrl+Shift+Z undo and redo.

For Word, the existing chat panel uses GLM-4.7-Flash without Copilot sign-in. Configure credentials through **AI Office: Configure Cloudflare**. The agent uses stable internal block IDs, structured live-document CRUD, undo, and a shimmer over its current target. The [document agent guide](extensions/ai-office-pdf/README.md#word-agent) explains configuration, the compact internal skill, boundaries and local-only tests.

See [editing commands and CLI examples](extensions/ai-office-common/README.md). On Windows, `scripts\ai-office-cli.cmd` invokes the bundled CLI. All toolbar document operations use the same engine adapters as the CLI. PDF tables/lists and PowerPoint heading styles are disabled in this first editing surface.

Agents create, read, update and delete document content by text with the `content.*` commands (`find`, `replace`, `delete`, `append`), available both to CLI scripts and to open documents through the `aiOffice.<kind>.execute` VS Code command and the `aiOfficeEdit` language model tool that the chat agent uses. Both live surfaces resolve the file themselves: with no file argument they edit the current active file, and a file name or path fragment targets another open document. Targets are resolved internally, so agents never carry shape or paragraph ids between inspect and run, and the editor shimmers over the region an agent command touches. The compact agent skill lives at `.agents/skills/ai-office-docs/SKILL.md`.

GitHub Copilot sign-in is not part of AI Office. Chat and agent surfaces are backed by the language models you configure yourself — for example an OpenAI-compatible endpoint added through Manage Models — and the sign-in entry points (chat title bar, status bar, accounts menu, setup dialogs and the Agents window gate) are removed. The [removal map](AI_OFFICE_REMOVAL_MAP.md) lists the removed surfaces.

Validate the office engines and custom-document save/undo flow with `node extensions/ai-office-common/test.mts` after building the extensions and Rust engines. Generated files and local test documents stay under ignored build directories.

Word agent edits support atomic batches with Unicode character ranges and explicit applied-operation receipts. The agent checks scope before claiming completion and can inspect actual document style sizes. Document history is addressed by file URI, so Undo/Redo works while chat has focus; request checkpoints use that same history and can be restored through chat or the history tool. Checkpoints are session-local and follow ordinary redo-branch and undo-retention limits.

Common Word formatting uses `office_document_format`: literal queries resolve matching blocks and Unicode ranges internally, apply all properties atomically, and return verified scope counts. `sizeFromStyle` copies the actual document style size. Invalid text-update/formatting combinations are rejected; repeated identical errors stop early with an honest result.

Word request checkpoints use the standard chat **Restore Checkpoint**, **Start Over**, and **Redo** controls and confirmation settings. Custom-document history follows the same request boundaries as the native chat timeline; no separate restore button is added to agent responses.
