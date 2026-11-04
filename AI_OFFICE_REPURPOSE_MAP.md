# AI Office reuse map

These modules are retained for document workflows, with Word/DOCX as the first implementation target. This is the implementation map for the cleanup, not a claim that the workflows below are already connected to Word. The supplied line counts are estimates, not measured savings or estimates of implementation effort.

| Retained module | Document purpose | Available now | Required Word adapter |
|---|---|---|---|
| `timeline` and `localHistory` | Document version history | Timeline providers, history storage, retention settings, entry commands | Capture custom-document saves; open immutable DOCX snapshots in Word; compare and restore through the live document session |
| `markers` | Document consistency checks | Diagnostic storage, severity, grouping, filtering and Problems view | Extract document structure, run consistency rules, and navigate findings to document anchors |
| Native Word snapshots and local versions | Document versioning without Git | Per-block Keep/Undo, saved-review recovery, timestamped DOCX snapshots and History comparisons | Further author-attributed revision history and version restore UI |
| `multiDiffEditor` | Review of agent changes | Multi-file text comparison, source resolvers and review UI infrastructure | Read-only semantic document comparisons, formatting/structure changes, and proposal acceptance with revision checks |
| `comments` | Document comments | Workbench Comments view and extension comment controllers; Word persists anchored threads with stable comment IDs, replies and resolved/orphaned state through agent tools | Workbench Comments view bridge and navigation between that view and native Word anchors |

## Current boundaries

The document workbench retains document history, markers, multi-diff and comments. Git/SCM UI and built-in Git extensions are excluded from both application windows. Shared extension/chat consumers receive empty SCM/Git compatibility services, while the agent-host and shared-process Git implementations are disabled. The Agents window does not register repository Changes, commit/sync, PR polling or apply-commit actions. `build/lib/test/officeWorkbench.test.ts`, `officeGit.test.ts` and the built-in extension policy tests protect these boundaries.

Word is implemented in [`extensions/ai-office-pdf`](extensions/ai-office-pdf), with the shared session and save lifecycle in [`provider.ts`](extensions/ai-office-common/provider.ts), the engine adapter in [`model.ts`](extensions/ai-office-common/model.ts), and the webview in [`editor.ts`](extensions/ai-office-common/editor.ts). The Word engine is vendored at [`engines/wordcraft`](engines/README.md) (trimmed upstream tree; DOCX view only).

The editor owns a live engine session, serializes operations, exports bytes, supports undo/redo and protects saves against external file changes. Chat-associated Word commands apply to the live session and register pending review snapshots. Native chat Keep/Undo controls now review readable text and formatting differences at paragraph/table, file, or all-file scope; shared structural resources are grouped coherently. Checkpoint restoration remains separate. The standalone CLI opens a separate engine session; it is not the owner of an already open editor.

## Document version history

Reuse [`LocalHistoryTimeline`](src/vs/workbench/contrib/localHistory/browser/localHistoryTimeline.ts) and [`IWorkingCopyHistoryService`](src/vs/workbench/services/workingCopy/common/workingCopyHistory.ts) for the history list and storage lifecycle.

There is a specific missing connection: [`WorkingCopyHistoryTracker`](src/vs/workbench/services/workingCopy/common/workingCopyHistoryTracker.ts) accepts only save events with file metadata (`stat`). [`MainThreadCustomEditorModel`](src/vs/workbench/api/browser/mainThreadCustomEditors.ts) emits successful saves with `reason` and `source`, without `stat`. Consequently, retaining the existing tracker does not automatically record Word saves. Crash-recovery backups and the engine's in-memory revision counter are also not durable version history.

The implementation needs a workbench-side custom-document save adapter, or an equivalent explicit history integration, that records the exact successfully saved bytes and metadata. Internal workbench services must stay behind that boundary rather than being imported into the Office extension. Capture an initial baseline and subsequent saved versions, identify each immutable version by a content hash, and record its time and source. Keep the existing enable/exclude/retention policies and decide an appropriate size limit for office documents.

Opening a history entry must use a read-only Word snapshot. Restoring must run through the active document lifecycle, check unsaved and externally changed content, retain the version being replaced, and refresh the editor and undo state. A raw file replacement behind an open engine session is insufficient. Rename and Save As need explicit history identity behavior.

Completion checks: save and reopen a document across app restarts; see distinct versions; avoid duplicate versions for unchanged bytes; compare and restore a version; preserve edits when a save fails or a restore conflicts; exercise local desktop and server-hosted Word sessions.

## Document consistency checks

Reuse [`IMarkerService`](src/vs/platform/markers/common/markers.ts) and the [`markers` view](src/vs/workbench/contrib/markers/browser/markersView.ts). They supply the diagnostic infrastructure, not the document analysis rules.

Start with deterministic checks against structured Word content, such as skipped heading levels, inconsistent configured terminology and inconsistent styles for equivalent headings. Each finding needs a rule ID, severity, explanation, document version and anchor. Fixes must use engine commands and participate in undo. Debounce analysis, cancel obsolete runs, and discard results whose base revision no longer matches the editor.

The existing Problems navigation opens a text-editor line/column selection. Introduce an adapter that resolves a finding to a Word paragraph, run or table cell and reveals it in the webview. A virtual text projection can provide diagnostic ranges, provided it retains an exact mapping back to document anchors. Do not present arbitrary page numbers or invented line numbers as valid text ranges.

Completion checks: a known inconsistent sample produces the expected findings; clicking a finding highlights its document location; fixing it removes the finding; undo restores it; late analysis cannot overwrite diagnostics from a newer revision.

## Document versioning without Git

Git and Source Control have been removed from the application. Word uses private native baseline/proposal snapshots for per-change decisions and immutable DOCX versions for history. Versions are content-hashed, timestamped and stored in local extension storage outside the document and repository. The History control compares a selected version with the current readable text/formatting projection. Saved documents and matching backups recover pending inline review after reload; mismatched external files never receive stale review state.

Source Control and Git are not a future document-versioning dependency. Preserve the shared text diff editor, multi-file review and native chat Keep/Undo independently of repository services.

## Review of agent changes

Reuse [`multiDiffEditor`](src/vs/workbench/contrib/multiDiffEditor) as the review container and source-resolution infrastructure. Its [`input model`](src/vs/workbench/contrib/multiDiffEditor/browser/multiDiffEditorInput.ts) resolves text models. Feeding it zipped DOCX bytes will not produce a document review.

Create a proposal from an immutable base version and a separate proposed engine state. Store the base content hash and live revision, the proposed commands, and the resulting snapshot. Provide a deterministic read-only text/structure projection for comparison, including explicit formatting and table changes; a plain paragraph-text diff would miss those changes. Map review entries back to Word locations and provide page previews where useful.

Accept and reject must operate on the proposal. Rejection leaves the live document unchanged. Acceptance checks the base revision and disk identity, then applies a validated operation through the document's serialized session with coherent undo behavior. The current immediate `execute` command must gain a separate proposal lifecycle before it can support review. Partial acceptance requires independent, validated operations; dependent changes should be accepted together or regenerated against the current version.

Completion checks: review text, formatting and table changes; reject without modifying the live document; accept and undo an entire proposal; detect stale proposals after user edits or external changes; preserve unsupported DOCX content on export.

## Document comments

Reuse the [`Comments` view](src/vs/workbench/contrib/comments) and extension comment controller APIs as the shared thread list. Word already supports `review.newComment`, exports/reopens comments, and shows a simple list inside its webview. That list is not connected to the workbench Comments view.

The shared [`State`](extensions/ai-office-common/types.ts) currently exposes comment author and text only. Extend the engine bridge with stable comment IDs, anchors and supported thread metadata before synchronizing controllers. Treat the document's comments as the persisted source of truth, and map workbench thread IDs to engine comment IDs. Replies and resolved state require explicit engine/import/export support rather than a UI-only implementation.

The workbench's current location labels and reveal behavior assume text ranges. Route comment selection to a document anchor and synchronize focus/highlighting with the webview. Preserve comments and their anchors through edits, undo, save and reopen. If an anchor cannot be resolved, show that state explicitly instead of silently attaching the comment to another paragraph.

Completion checks: create a comment from a Word selection; see it in both surfaces; navigate back to the selection; save and reopen with identity and location preserved; verify undo and anchor behavior after surrounding text changes.

## Shared foundation and implementation order

All five workflows need a document adapter that provides immutable byte snapshots, saved-version hashes, live revisions, structured content and resolvable anchors. Page coordinates are useful for rendering but cannot be durable identifiers because pagination changes. Engine paragraph/run indexes also need validation or rebasing after structural edits. Version every projection and anchor mapping against the snapshot from which it was produced.

Implement in this order:

1. Saved Word snapshots, the custom-save history connection, read-only historical opening and restore.
2. Structured extraction, document anchors and semantic comparisons; use them for local document history and revision checkpoints.
3. Agent proposals with revision validation and accept/reject/undo.
4. Deterministic consistency diagnostics with document navigation and fixes.
5. Workbench comment integration using stable IDs and the same anchor resolver.

These stages have separate completion checks. Retaining imports, changing panel labels or passing dependency tests does not establish that a document workflow is implemented.

## Implemented Word review and local versions

Word agent changes now have native chat review entries and inline Word page view zones. Red deleted text and green proposed text appear at the affected paragraph/table, with per-change and document Keep/Undo; native chat keeps file and all-file decisions. Private native baseline/proposal snapshots and timestamped DOCX versions are kept in extension storage, without Git. History compares a selected version with the current text/formatting projection. Saved files and matching backups restore pending inline review on reload; content hashes reject mismatched external documents. The earlier proposal discussion above describes future work, rather than the immediate-apply review behavior now used by Word, which follows VS Code chat editing: Keep resolves the proposal and Undo reverses only its selected blocks.
