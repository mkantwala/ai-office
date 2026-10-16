# AI Office

AI Office is a VS Code fork with an office document explorer, Word/PDF/PowerPoint editing and shared CLI commands for agents. Word embeds Wordcraft's full web interface; PDF and PowerPoint use compact editing toolbars. The extension marketplace is disabled. Markdown uses the existing editor.

Word and PDF use Rust engines compiled to WebAssembly. Wordcraft runs its egui ribbon, dialogs, rulers and editable document canvas inside a VS Code custom editor. The host handles file pickers, saves, conflicts and recovery. PowerPoint uses `@office-kit/pptx` and `@office-kit/pptx-preview`.

Word opens with Page Width zoom so the document follows the available editor pane. Narrow panes group additional ribbon tabs under More and document actions under Actions; resizing the window or side panels does not enlarge the embedded page beyond its pane. Manual zoom remains available under View.

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
node extensions/ai-office-pdf/build.mts --word-full
node extensions/ai-office-slides/build.mts
```

After building the VS Code client, launch with `scripts\code.bat`. For an already prepared development build, set `VSCODE_SKIP_PRELAUNCH=1` to skip the normal preparation step.

## Editing and agents

Click in Word to type and drag to select. Click a PDF paragraph or PowerPoint text box/table cell to edit it; press Ctrl+Enter or click outside to finish. Ctrl+S saves; Ctrl+Z and Ctrl+Shift+Z undo and redo.

See [editing commands and CLI examples](extensions/ai-office-common/README.md). On Windows, `scripts\ai-office-cli.cmd` invokes the bundled CLI. Word's native engine commands are available through both the live editor and headless CLI; UI commands and host file dialogs belong to the live editor. PDF tables/lists and PowerPoint heading styles are disabled in this first editing surface.

This embeds Wordcraft's web capabilities, including its upstream feature and DOCX fidelity limitations. Desktop-only operating-system integrations are not provided by WebAssembly. The ribbon Paste action reads text from the host clipboard; browser keyboard paste retains Wordcraft's browser support. Print PDF offers a PDF export. Native Wordcraft owns document undo/redo; use its ribbon or keyboard shortcuts while the document canvas is focused.

Validate the office engines and custom-document save/undo flow with `node extensions/ai-office-common/test.mts` after building the extensions and Rust engines. Generated files and local test documents stay under ignored build directories.
