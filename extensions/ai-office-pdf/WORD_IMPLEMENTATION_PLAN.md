# WordCraft command parity

The app keeps WordCraft's pinned `159b84f54c74e9e82fb3109b758ae7dd6f4e3ab0` engine and existing WASM rendering path. The independent `../wordcraft-upstream` checkout is reference material, not a build input. No engine upgrade is bundled with this command expansion.

## Audited status

`WORD_CAPABILITIES.json` is generated from the **running registry**, including generated entries, using `node scripts/audit-word-commands.mts --write` after building Word. Running the script without `--write` checks for drift. The development-only `command_catalog` export is not a model tool and does not place the registry in chat context.

The pinned registry contains 622 commands. This first batch raises the integrated bridge from 99 to **127** commands. The inventory distinguishes editor controls from targeted block commands and explicitly marks 495 remaining entries as pending. Pending includes application-owned operations such as file Save and UI dialog requests; the count does not mean 495 missing document mutations. **Full command parity is not complete.**

The inspected reference at `d588e0b8` declares 633 commands. Comparing its command definitions with the pin found additional work including bookmark deletion, empty-header removal, multilevel-list presets, recipient sorting/filtering and cross-document object copying. These need individual adaptations or a separately verified engine upgrade; exposing names without implementation would misrepresent support.

## Landed: native character formatting

28 additional commands cover double/word-only underline, double strike, character borders, sub/superscript, caps/small caps, hidden text, outline/shadow/emboss/engrave, font-size steps, shading, Change Case, direct-format reset, character spacing/scale/position, character styles, native CharProps, enclosed characters, phonetic guides and text effects. They reuse the original engine implementations, native DOCX serializer and layout.

Agents discover exact parameters through `office_document_help` or bounded `office_document_tools`, then use `office_document_edit` with operation `command`, a block ID, a fresh version and an optional Unicode character range. Related commands can be batched into one atomic Undo/review unit. Generic `office_document_run` remains PDF-only so it cannot bypass Word's targeting and guard checks.

Phonetic commands require explicit ruby text or null, rather than silently opening an upstream dialog. Ruby and enclosure operations introduce inline objects, so agents must reread character offsets afterward. Block IDs survive save/reopen. Hashes belong to fresh reads: native import may normalize internal defaults while producing identical DOCX XML.

The formatting-state and font-dialog registry commands are not exposed as edits: native effective formatting is read through `office_document_read` operation `formatting` (paragraph ID and optional Unicode range), and existing application font controls serve the dialog role. The read preserves caret, document revision and native history. Ten native toggles have their supported explicit value parameter documented in bridge metadata, avoiding accidental reversal when formatting is already applied. The compact root skill points to the detailed command guide without preloading its schemas. Word formatting, paragraph and table discovery use separate categories.

## Next integration stages

1. Paragraphs, lists and styles: classify executable mutations versus dialog-only entry points; adapt richer style/list definitions and preserve block targeting.
2. Tables, fields, bookmarks and references: validate table selection semantics, cross-references and story changes; add identity and DOCX round-trip fixtures.
3. Review: distinguish native Track Changes from agent proposal review. Never treat Keep/Undo as accepting arbitrary document revisions; preserve comment anchors and current protection guards.
4. Equations, shapes, drawings, charts, content controls and mail merge: adapt typed inputs, referenced resources and selection targets. Use the existing renderer/serializer, not parallel HTML document content.
5. Application and file operations: reuse Save/Save As, trust, cancellation, checkpoints and managed artifacts. Filesystem, printer, speech and native-dialog commands need host adapters; they cannot be enabled blindly in WASM.
6. Upstream-only additions: upgrade the source pin independently with regression checks for typing performance, pagination, block identities, comments, review and external DOCX fidelity.

Each stage must update the running inventory, implement actual command routing, document limitations and verify affected behavior before claiming parity. No extra permanent model tool is needed for a new command.

## Validation

Every newly exposed formatting command has a targeted-block integration test for untouched neighboring content, stable IDs, native Undo/Redo, stale-version refusal and saved/reopened DOCX XML. Additional tests cover phonetic-dialog refusal, inventory drift, bounded discovery, coherent review, comments and atomic rollback. Native DOCX round trips do not prove identical rendering in Microsoft Word or Google Docs; external fidelity testing remains necessary for broader interoperability claims.

The initial integration passes 106 shared tests, including one real-engine check per new formatting command, coherent review/Undo, and read-only Unicode formatting inspection. The Word WASM build, TypeScript, targeted lint, running-inventory check and clean-source patch check pass. A live OpenCode GLM-5.3-Flash Free session applied small caps and 1.5-point spacing to one selected phrase in a single atomic batch; the receipts confirmed one edit and untouched neighboring content.
