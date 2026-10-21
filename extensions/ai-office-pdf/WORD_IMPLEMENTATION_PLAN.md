# WordCraft command parity

The app keeps WordCraft's pinned `159b84f54c74e9e82fb3109b758ae7dd6f4e3ab0` engine and existing WASM rendering path. The independent `../wordcraft-upstream` checkout is reference material, not a build input. No engine upgrade is bundled with this command expansion.

## Audited status

`WORD_CAPABILITIES.json` is generated from the **running registry**, including generated entries, using `node scripts/audit-word-commands.mts --write` after building Word. Running the script without `--write` checks for drift. The development-only `command_catalog` export is not a model tool and does not place the registry in chat context.

The pinned registry contains 622 commands. The formatting and paragraph/list/style batches raise the integrated bridge from 99 to **181** commands. The inventory distinguishes editor controls from targeted block commands and explicitly marks 441 remaining entries as pending. Pending includes application-owned operations such as file Save and UI dialog requests; the count does not mean 495 missing document mutations. **Full command parity is not complete.**

The inspected reference at `d588e0b8` declares 633 commands. Comparing its command definitions with the pin found additional work including bookmark deletion, empty-header removal, multilevel-list presets, recipient sorting/filtering and cross-document object copying. These need individual adaptations or a separately verified engine upgrade; exposing names without implementation would misrepresent support.

## Landed: native character formatting

28 additional commands cover double/word-only underline, double strike, character borders, sub/superscript, caps/small caps, hidden text, outline/shadow/emboss/engrave, font-size steps, shading, Change Case, direct-format reset, character spacing/scale/position, character styles, native CharProps, enclosed characters, phonetic guides and text effects. They reuse the original engine implementations, native DOCX serializer and layout.

Agents discover exact parameters through `office_document_help` or bounded `office_document_tools`, then use `office_document_edit` with operation `command`, a block ID, a fresh version and an optional Unicode character range. Related commands can be batched into one atomic Undo/review unit. Generic `office_document_run` remains PDF-only so it cannot bypass Word's targeting and guard checks.

Phonetic commands require explicit ruby text or null, rather than silently opening an upstream dialog. Ruby and enclosure operations introduce inline objects, so agents must reread character offsets afterward. Block IDs survive save/reopen. Hashes belong to fresh reads: native import may normalize internal defaults while producing identical DOCX XML.

The formatting-state and font-dialog registry commands are not exposed as edits: native effective formatting is read through `office_document_read` operation `formatting` (paragraph ID and optional Unicode range), and existing application font controls serve the dialog role. The read preserves caret, document revision and native history. Ten native toggles have their supported explicit value parameter documented in bridge metadata, avoiding accidental reversal when formatting is already applied. The compact root skill points to the detailed command guide without preloading its schemas. Word formatting, paragraph and table discovery use separate categories.

## Landed: paragraphs, lists and styles

The second batch adds 54 commands: 43 paragraph operations, eight style mutations, multilevel list definition and two native read-only inspections. It covers alignment/direction, indents, spacing, headings, list levels/restarts/value overrides, paragraph movement, shading/borders, pagination, Asian typography, tabs and native paragraph-property patches. Custom styles can be created, modified, updated to match, deleted, renamed and shown/hidden in the gallery; the inspector can clear one formatting level. Character-level clearing accepts exact Unicode ranges and preserves neighboring runs.

Native resets and inspector clearing now retain block IDs. Formatting-only paragraph snapshots strip internal IDs before they can enter shared style definitions; Update to Match preserves the target's identity. Comments, style resources and numbering definitions participate in native Undo and coherent Keep/Undo review. Paragraph movement refuses to split an existing comment span or move content across its boundary; complete annotated selections remain movable through native editor commands.

Read-only `office_document_read` operation `inspect` supports `list.get` and `styles.inspect` on a paragraph ID, optional Unicode range and native inspection parameters. It clones the inspection session, returns a bounded result and never changes the actual caret, revision or history. Mutation commands are refused on this read route; inspection commands are refused on the edit route. Paragraph/list/style discovery remains behind existing tools, with no new permanent model tool.

Styles and existing list definitions are shared resources: changing them can affect other paragraphs. Numbering restart/value changes include subsequent contiguous items of the same list. Default tab changes are document-wide. These scopes are documented in the command guide and verified where applicable. Dialog-only paragraph/style-pane commands, cross-file style copying and multi-paragraph sorting remain pending rather than falsely returning success.

## Next integration stages

1. Finish paragraph/style integration: add explicit contiguous-block targeting for sorting and a host adapter for cross-file style copying. Map remaining dialog-only actions to application controls.
2. Tables, fields, bookmarks and references: validate table selection semantics, cross-references and story changes; add identity and DOCX round-trip fixtures.
3. Review: distinguish native Track Changes from agent proposal review. Never treat Keep/Undo as accepting arbitrary document revisions; preserve comment anchors and current protection guards.
4. Equations, shapes, drawings, charts, content controls and mail merge: adapt typed inputs, referenced resources and selection targets. Use the existing renderer/serializer, not parallel HTML document content.
5. Application and file operations: reuse Save/Save As, trust, cancellation, checkpoints and managed artifacts. Filesystem, printer, speech and native-dialog commands need host adapters; they cannot be enabled blindly in WASM.
6. Upstream-only additions: upgrade the source pin independently with regression checks for typing performance, pagination, block identities, comments, review and external DOCX fidelity.

Each stage must update the running inventory, implement actual command routing, document limitations and verify affected behavior before claiming parity. No extra permanent model tool is needed for a new command.

## Validation

Every newly exposed formatting command has a targeted-block integration test for untouched neighboring content, stable IDs, native Undo/Redo, stale-version refusal and saved/reopened DOCX XML. Additional tests cover phonetic-dialog refusal, inventory drift, bounded discovery, coherent review, comments and atomic rollback. Native DOCX round trips do not prove identical rendering in Microsoft Word or Google Docs; external fidelity testing remains necessary for broader interoperability claims.

The initial integration passes 106 shared tests, including one real-engine check per new formatting command, coherent review/Undo, and read-only Unicode formatting inspection. The Word WASM build, TypeScript, targeted lint, running-inventory check and clean-source patch check pass. A live OpenCode GLM-5.3-Flash Free session applied small caps and 1.5-point spacing to one selected phrase in a single atomic batch; the receipts confirmed one edit and untouched neighboring content.

The paragraph/list/style batch passes 163 shared document/provider tests and a focused Rust identity regression test. The two comment/range regressions were reproduced before their fixes and pass afterward. TypeScript, targeted ESLint, the native Word WASM build, the running 181-command inventory and clean-source patch application were verified. The always-loaded skill is 388 words; detailed command guidance stays on demand.
