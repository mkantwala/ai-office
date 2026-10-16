# AI Office

AI Office is a VS Code fork with an office document explorer, direct Word/PDF/PowerPoint editing, a compact toolbar and shared CLI commands for agents. The extension marketplace is disabled. Markdown uses the existing editor.

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

See [editing commands and CLI examples](extensions/ai-office-common/README.md). On Windows, `scripts\ai-office-cli.cmd` invokes the bundled CLI. All toolbar document operations use the same engine adapters as the CLI. PDF tables/lists and PowerPoint heading styles are disabled in this first editing surface.

Validate the office engines and custom-document save/undo flow with `node extensions/ai-office-common/test.mts` after building the extensions and Rust engines. Generated files and local test documents stay under ignored build directories.
