# Document inspection and agent verification plan

Use one guarded document service for editing and inspection. Keep stable internal block IDs, native history, pending drafts, comments, activity overlays and Keep/Undo review. Manual input and agent input already modify the same native document; there is no separate conversion into blocks.

## What exists

| Capability | Current integration | Remaining work |
| --- | --- | --- |
| Semantic structure | `../ai-office-common/documentTools.ts` registers eight tools with `vscode.lm.registerTool`; `provider.ts` serializes reads and edits with manual operations | Strengthen operation-specific TypeScript request/result types; keep runtime validation at transport boundaries |
| Guarded commands | Native `agent.rs` exposes IDs, content hashes and revisions; the host adds document-instance versions and request checkpoints | Reuse this path for inspector actions; never bypass it with a second editing session |
| Geometry | Native `agent.rects` gives block rectangles; page sizes and table geometry already support the editor | Expose bounded, revision-tagged geometry to agents, including coordinate units and page numbering |
| Pixels | Native `OfficeDocument.render` rasterizes pages; `engineWorker.ts` returns RGBA pixels to the editor | Encode selected pages as PNG and support image tool results; current document tools return text only |
| Errors | Command failures and stale-version failures reach tool callers | Give inspector findings structured codes, severity, target IDs and evidence |
| External agent transport | `opencode/documentBridge.ts` is a thin authenticated loopback MCP adapter; its callback uses `vscode.lm.invokeTool` | Preserve this shared implementation and cancellation/request attribution |

Upstream WordCraft provides `inspect_document`, `get_text` and `render_page` in `crates/mcp/src/tools.rs`. Its headless backend renders without a browser. `screenshot`, `click`, `key` and `ui_inspect` need a connected WordCraft desktop UI; those desktop controls are not automatically available in Sangam One's webview.

## Architecture decisions

- Use VS Code's existing tool registration for agents hosted by the app. A typed internal API should sit behind those tools, not replace their registration or permission/request context.
- Keep MCP for external processes such as OpenCode. Stdio cannot run directly in a browser, but MCP also has HTTP transports. Removing MCP does not itself make this extension browser-compatible: the current worker and host still use Node APIs.
- Treat transport latency separately from model inference, layout and rendering. Measure before claiming a performance improvement.
- Read structure to understand, use guarded semantic commands to edit, and request pixels when appearance matters. Keep raw mouse/keyboard automation for UI testing and unsupported interactions.
- Every inspection response must identify the same document instance and revision. A render or geometry result from an older revision must be marked stale before it can justify another edit.

## Delivery order

1. **Manual input correctness.** Route clipboard text through native `edit.pasteText`, preserving paragraph boundaries and one-step undo. Verify typing, Enter, Shift+Enter, paste, selection replacement, stale guards, subsequent agent edits, and DOCX save/reopen. Preserve normal paragraph-cache behavior.
2. **Typed shared contracts.** Introduce discriminated request/result types around the existing document backend and worker. Keep operation schemas, runtime validation and lazy-loaded help aligned. Test equivalent results through built-in tools and the OpenCode adapter, including cancellation and pending review.
3. **Bounded inspection.** Add a read-only inspection surface for outline/text, requested block geometry, page metadata and selected-page images. Require URI/version, page/block filters and strict size limits. Return document coordinates with a stated origin, units and page-index convention. Include headers, footers, tables and other stories when requested; flag unsupported coverage explicitly.
4. **Deterministic verification.** Capture targeted semantic snapshots before and after an operation; compare content, formatting, IDs, comments and unrelated nodes. Save/reopen and compare again, allowing IDs to change when external edits legitimately invalidate private mappings. Track affected pages and layout changes by revision. Report overflow, table splits and collisions as findings with context; intentional wrapping and overlaps must not automatically fail verification.
5. **Optional visual review.** Render only affected pages plus neighboring pages when pagination moves content. Use the same revision for geometry and pixels. Add numbered block overlays to a separate review image, map labels to internal IDs, and retain the unmodified image. Vision findings are suggestions requiring structural confirmation before another edit.
6. **Interoperability checks.** Add isolated export fixtures rendered by pinned LibreOffice with known fonts, bounded execution and no mutation of the user's source file. Compare semantic content and visual differences with tolerances. For Word fidelity claims, test representative exports in Microsoft Word as well.

## Acceptance criteria

- Manual typing and paste produce fresh readable block content and unique stable IDs; agent edits touch only the intended targets.
- A paste, agent batch and review decision retain their expected undo granularity. Pending agent review continues to protect later user edits.
- Inspector output is bounded, cancellable and revision-consistent; large documents do not require whole-document text or screenshots for a local edit.
- Inspection itself does not modify document content, selection, history or dirty state.
- Round-trip fixtures cover Unicode, paragraph formatting, tables, images, comments and review data. Unsupported checks are reported as unavailable, never as passed.
- A page-count change, overlap or renderer difference is evaluated against the requested edit and supported layout semantics.
- WordCraft rendering demonstrates WordCraft appearance. LibreOffice supplies an independent compatibility signal, not proof of identical Microsoft Word pagination. Upstream documents its layout gaps in `docs/layout-parity.md`.
