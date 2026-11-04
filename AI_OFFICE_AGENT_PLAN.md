# Agent feature plan — the DOCX recipe, then PDF and PPT TODOs

Every document kind should end up with the same agent capability shape that Word already has.
This doc records **exactly what is implemented for DOCX today** so a future PDF agent branch or
PPT agent branch can replicate it step by step — reusing the shared plumbing in
`extensions/ai-office-common/provider.ts`, not writing a second one.

Upstream engines referenced by the TODOs (do not clone until the branch starts):

- PDF: <https://github.com/storytold/pdfcraft> — already pinned in `engines/sources.json` with `engines/pdfcraft.patch`
- PPT: <https://github.com/storytold/deckcraft> — not yet vendored anywhere

---

## 1. The DOCX implementation (current, working — the recipe)

### Engine layer

1. `engines/wordcraft/` is a vendored, trimmed upstream checkout (upstream commit `b4da59f0` of
   [storytold/wordcraft](https://github.com/storytold/wordcraft)). Workspace members are
   `crates/*` + `apps/ai-office-preview` only; `ui-egui`, `mcp`, `control-key`, `zotero`, apps
   and `xtask` are **not** vendored. License/attribution files (`LICENSE-MIT`, `LICENSE-APACHE`,
   `NOTICE`, `ATTRIBUTION.md`) ship with the build.
2. `apps/ai-office-preview` is the only app we added: a **headless** wasm entry (no UI) exposing
   `Session` + `execute(id, args)` for every `CommandSpec` in `crates/engine/src/cmd/*.rs`
   (622 ids: para, table, format, styles, references, fields, caret, select, …), DOCX I/O
   dispatch, and rendering to RGBA pixels. `cargo test -p wordcraft-docx --tests` covers the
   DOCX import/export regressions.
3. Build with `node extensions/ai-office-pdf/build.mts --word` → `--profile release`
   (opt-level 3, thin LTO) → wasm-bindgen → `extensions/ai-office-pdf/media/word-engine/`
   (gitignored) with a generated `BUILD.txt` label. Do not fall back to a dev profile for any
   editor-facing build.

### Host layer (`extensions/ai-office-common/provider.ts`, `registerEditors(context, kinds)`)

4. Eleven language-model tools are registered once for all kinds
   (`registerDocumentTools`): `office_document_read`, `office_document_edit`,
   `office_document_format`, `office_document_draft`, `office_document_comments`,
   `office_document_history`, `office_document_activity`, `office_document_help`,
   `office_document_tools`, `office_document_run`, and the generic unfiltered `aiOfficeEdit`.
   Tool *names* are shared across kinds; each invoke resolves the **active document** and routes
   to the matching kind internally (PDF and PPT endpoints differ).
5. Guarded semantic edits flow through the engine's native agent surface: `agentEdit` /
   `agentRead` with the document revision, internal block/stable IDs, request checkpoints
   (`withAgentRequest`), activity shimmer (`OfficeActivity.renderAgentActivity`), `recordEdit`,
   and inline red/green review with per-change **Keep/Undo** (`publishReview`).
6. The run gate in `aiOffice.<kind>.capabilities` (input.run): the command must be discovered
   first (`definition` lookup), mutations require `input.version === doc.version` and the word
   branch accepts omitted params (object ≤ 100k chars when present). The PDF branch keeps the
   1–10 page `text_extract` rule and command-result receipts/artifacts. Both branches call
   `execute(doc, id, params, true, true, token)` — the same passthrough used by
   `aiOfficeEdit`; undo/redo ids are special-cased only for edit-review bookkeeping.
7. Discovery is unfiltered end to end: `model.ts` hands the full engine `commands()` list to
   `documentCapabilities.ts` (`discoverCommands` paginates; nothing is hidden), and
   `office_document_help` returns categories/per-command schemas with guidance pointing at
   `office_document_edit` / `office_document_run` / `aiOfficeEdit`.
8. Every agent path checks `checkAgentAccess(token)` (cancellation + workspace trust).

### Agent loop, today

9. `chatParticipants` entry `aiOffice.wordAgent` (`extensions/ai-office-common/cloudflareAgent.ts`)
   drives a Word document agent over `vscode.lm` model + tools; its system prompt restricts the
   model to the document tools; the compact skill ships via `chatSkills`
   (`extensions/ai-office-pdf/skills/ai-office-docs/SKILL.md`).
10. Tests: `node extensions/ai-office-common/test.mts` — 76 tests, including the run gate
    (`provider.test.cts`: Word run with/without version, stale rejection, Keep/Undo review,
    cancellation, trust), plus `cargo test -p wordcraft-docx` and `.build/menu-probe.cjs` for
    visual rendering.

### What "same for the next kind" means

A new kind needs: an engine that (a) parses/serializes that file format and (b) exposes a
command registry + headless `execute`, a wasm/native preview app, `registerEditors` in its
extension, and then **everything above in host code already works** — the run gate, discovery,
tools and guards are kind-generic. The only kind-specific code paths today are the PDF branch
in the run gate and per-kind semantic tool payloads.

---

## 2. TODO — PPT agent parity (upstream: [storytold/deckcraft](https://github.com/storytold/deckcraft))

Current state: `extensions/ai-office-slides` registers the `aiOffice.slides` custom editor
through the shared `registerEditors(context, ['slides'])`; text/table/cell editing is
`@office-kit/pptx`-based in JS. The shared tools already register, but there is no Rust engine,
no command registry, no guarded block edits, no agent loop.

- [ ] **Vendor Deckcraft into `engines/deckcraft/`** exactly like wordcraft: clone upstream,
      trim the workspace to engine crates + a headless `ai-office-preview`-style wasm app, keep
      `LICENSE`/`NOTICE`/`ATTRIBUTION` files, commit the tree. Reference clones stay beside the
      repo (`../deckcraft` / `../deckcraft-upstream`).
- [ ] **Extend `extensions/ai-office-pdf/build.mts`** with a `--ppt`/slides profile branch
      mirroring the word branch: wasm release profile, wasm-bindgen, output under
      `media/deck-engine/`, BUILD.txt label. Keep the PDF native build untouched.
- [ ] **Switch the slides document client** from the JS editing model to the engine worker
      (`engineWorker.ts` + deck wasm), preserving `aiOffice.slides` custom editor, undo, and
      save formats.
- [ ] **Command registry**: confirm `model.ts` passthrough picks up the deck command list; add
      engine `CommandSpec` ids covering the PPTX feature set (text frames, tables, shapes,
      themes) with `crates/engine/src/cmd/*.rs` tests.
- [ ] **Guarded block edits**: map slide shapes/runs to stable internal block IDs; wire
      `agentEdit`/`agentRead` equivalents; shimmer + request checkpoints come free from the
      shared host code.
- [ ] **Run gate**: the word branch of the run gate is already generic; only PDF-specific
      validations live in the pdf branch — verify slides needs no new branch.
- [ ] **Skill + guidance**: add a slides variant of the `ai-office-docs` skill and extend
      `office_document_help` guidance for slide-specific flow (placeholders, masters, layouts).
- [ ] **Tests**: provider-level slides tool tests (mirror `provider.test.cts`), PPTX round-trip
      fidelity suite (python-pptx / LibreOffice Impress comparisons like the DOCX one), and a
      slides menu-probe for visual verification.
- [ ] **DOCX-compatibility-style record**: `extensions/ai-office-slides/PPTX_COMPATIBILITY.md`
      documenting round trips before claiming parity.

## 3. TODO — PDF agent parity (upstream: [storytold/pdfcraft](https://github.com/storytold/pdfcraft))

Current state: the strongest non-Word base — native engine pinned at `engines/sources.json`
(`pdfcraft.patch`, checksummed) builds `sangam-pdf`; `office_document_run` already executes
discovered PDF commands; discovery is complete; command-result receipts + artifacts exist.

- [ ] **Vendoring decision**: either keep the pinned+patch model (reproducible, documented in
      [engines/README.md](engines/README.md)) or vendor like wordcraft. If vendoring: copy the
      prepared checkout into `engines/pdfcraft/`, fold `pdfcraft.patch` into the files, delete
      the pin from `sources.json`, and drop the `prepare-office-engines.mts` pdfcraft path.
- [ ] **Semantic tool parity**: make the eight guarded tools first-class for PDF objects the
      way they are for Word blocks — text runs/shapes on a page as stable IDs through the
      native agent surface, instead of canvas-only gestures; PDF native undo is single-step, so
      decide whether Keep/Undo review zones apply or remain Word-only.
- [ ] **Tool list scoping for the model** — see §4; PDF's read (1–10 range rule in the run
      gate) is already enforced.
- [ ] **Tests**: extend `provider.test.cts` with the same run-gate assertions used for Word
      (success, stale, cancellation, trust) against a synthetic PDF, and add a PDF round-trip
      record section similar to `DOCX_COMPATIBILITY.md`.

## 4. Cross-cutting: kind-scoped agent tool lists

The eleven tool names are shared across kinds, so scoping happens in the **agent loop**:

- Each turn, the loop exposes only the active document kind's tool set: a DOCX target → Word
  tools, PPTX → slides tools, PDF → PDF tools. (Optional: `vscode.lm` dynamic
  register/dispose on active-editor change to keep the built-in picker aligned; the model list
  refreshes between turns.)
- Discovery is already kind-aware (each engine exposes only its own commands), and the run gate
  rejects mismatched targets ("Use the guarded document tools…" / format-check on `vscode.open`
  URIs), so a leaked tool call cannot mutate the wrong kind.
- If an external agent is ever attached again, it must enforce the same discipline in its own
  host; `checkAgentAccess` remains the trust/cancellation backstop either way.

## 5. Property of the recipe

Keep it small and layered: **engine = upstream + committed changes with tests**, **host =
generic plumbing in `extensions/ai-office-common`**, **kind differences = separate wasm/native
engine + a bump token in the run gate**. If a second implementation of run/review/discovery
appears, merge it back into the shared provider instead.
