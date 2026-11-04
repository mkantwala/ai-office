# AI Office (Sangam One)

An AI-native office documents editor built on **Code - OSS**: direct DOCX, PDF and PPTX editing on page-shaped canvases, a compact document toolbar, and built-in `vscode.lm` agent tools over VS Code's own language-model API. No extension marketplace, no Copilot sign-in, no external agent bridges.

Start here: [AI Office overview](AI_OFFICE.md) · [Engine sources](engines/README.md) · [Agent feature plan](AI_OFFICE_AGENT_PLAN.md) · [Editing commands & CLI](extensions/ai-office-common/README.md)

## What works today

**Word (DOCX) — the reference integration**

The Word engine (Wordcraft) is vendored at `engines/wordcraft/` and compiled to a headless WebAssembly module that renders and edits full-fidelity documents in a canvas webview:

- Direct typing, selection, toolbar formatting, outline and tables
- Every one of the engine's **622 registered commands** is agent-reachable: `office_document_help` discovers any command schema, and `office_document_run` / `aiOfficeEdit` execute it (caret/select commands establish the engine caret)
- Guarded semantic tools: `office_document_read/edit/format/draft/comments/history/activity` operate on stable internal block IDs with version gates and native undo
- Agent changes get inline red/green review with per-change **Keep/Undo**, file-level and all-files review, plus **History** (up to 30 local versions, no Git required)
- Standard DOCX in/out verified against ECMA-376 with round trips through LibreOffice and python-docx (see [DOCX compatibility](extensions/ai-office-pdf/DOCX_COMPATIBILITY.md))

**PDF**

Pinned Pdfcraft native engine (see [PDF implementation plan](extensions/ai-office-pdf/PDF_IMPLEMENTATION_PLAN.md)): complete file workflows, forms, signing, merge/split, redaction, measurements and exports; direct text/shape editing on the canvas. `office_document_run` executes discovered PDF commands on the open document.

**PowerPoint (PPTX)**

Text, table and cell editing through `@office-kit/pptx` in the `Sangam One Slides` editor. No agent tools in this milestone.

## Future planning (I own this code)

1. **I own the engines inside this repository.** Engine behavior changes land directly in `engines/wordcraft/` with a regression test — never as an external runtime patch. Upstream reference clones (`../wordcraft-upstream`) stay read-only mirrors for comparison.
2. **Agent integration is mine to (re)build on built-in tools only.** The OpenCode MCP bridge was removed (`extensions/ai-office-pdf/opencode/`); the agent surface is `vscode.lm` document tools plus `aiOfficeEdit`. No loopback adapters, no credential passing.
3. **Kind-scoped agent tools.** When an agent is attached, a DOCX target exposes only the Word tools and a PPTX target only the PPT tools: the model's tool list follows the *active document kind* each turn. Discovery is already kind-aware (each engine exposes only its own command registry), and mismatched tool calls are rejected at the run gate.
4. **PDF and PPT agent parity** follow the exact DOCX recipe — engine vendoring, headless wasm app, guarded tools, discovery, run gate, tests. Written down as TODOs in [AI_OFFICE_AGENT_PLAN.md](AI_OFFICE_AGENT_PLAN.md) with the upstream engines:
   - PDF: [storytold/pdfcraft](https://github.com/storytold/pdfcraft)
   - PPT: [storytold/deckcraft](https://github.com/storytold/deckcraft) (future vendored engine, mirroring the wordcraft approach)
5. **Word first.** Finish the Word command surface on this branch; open dedicated branches for PDF and PPT agent work using the plan document as the checklist.

## Documentation map

| Document | Contents |
| --- | --- |
| [AI_OFFICE.md](AI_OFFICE.md) | Overview, setup, build, editing and agent behavior |
| [AI_OFFICE_AGENT_PLAN.md](AI_OFFICE_AGENT_PLAN.md) | The DOCX agent recipe described step by step, plus PDF/PPT TODOs |
| [AI_OFFICE_REMOVAL_MAP.md](AI_OFFICE_REMOVAL_MAP.md) | Removed VS Code surfaces and retained API dependencies |
| [AI_OFFICE_REPURPOSE_MAP.md](AI_OFFICE_REPURPOSE_MAP.md) | Reused VS Code workbench capabilities and missing Word adapters |
| [engines/README.md](engines/README.md) | Engine vendoring model (wordcraft vendored, pdfcraft pinned+patched) |
| [extensions/ai-office-pdf/README.md](extensions/ai-office-pdf/README.md) | Workspace, agent tools and validation details |
| [extensions/ai-office-pdf/DOCX_COMPATIBILITY.md](extensions/ai-office-pdf/DOCX_COMPATIBILITY.md) | DOCX fidelity verification record |
| [extensions/ai-office-common/README.md](extensions/ai-office-common/README.md) | Document CLI commands and examples |

## Building

```powershell
npm install
node scripts/prepare-office-engines.mts          # fetches the pinned pdfcraft revision
rustup target add wasm32-unknown-unknown
cargo install wasm-bindgen-cli --version 0.2.129 --locked --root .build/pdf-tools
node extensions/ai-office-pdf/build.mts           # PDF native engine
node extensions/ai-office-pdf/build.mts --word    # Word wasm (release profile)
node extensions/ai-office-slides/build.mts        # PPTX extension
```

After building the VS Code client, launch with `scripts\code.bat` (a prepared development build can set `VSCODE_SKIP_PRELAUNCH=1`). Test suite: `node extensions/ai-office-common/test.mts` (~76 tests; `--word` runs the Word subset only). Engine crates run their own upstream test suites (`cargo test -p wordcraft-docx --tests`).

## Acknowledgements and licenses

- **Code - OSS** — © Microsoft Corporation, MIT license. The upstream VS Code project and its license (`LICENSE.txt`) are retained; AI Office is an independent product built on this source.
- **Wordcraft** — [storytold/wordcraft](https://github.com/storytold/wordcraft), vendored at `engines/wordcraft/` for Word. Its `LICENSE-MIT`, `LICENSE-APACHE`, `NOTICE` and `ATTRIBUTION.md` ship with the editor build under `extensions/ai-office-pdf/media/word-engine/`.
- **Pdfcraft** — [storytold/pdfcraft](https://github.com/storytold/pdfcraft), pinned revision plus `engines/pdfcraft.patch`; upstream notices, attribution and asset licenses are packaged with PDF builds.
- **@office-kit/pptx / pptx-preview / dompurify / fflate** — licenses copied into `extensions/ai-office-pdf/out/licenses/` by the build.
