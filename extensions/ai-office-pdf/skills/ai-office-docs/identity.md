# Identity and concurrency

Paragraphs (including headings and list items), tables, rows, and cells have opaque internal IDs. Parent IDs describe containment. IDs are independent of text and position and survive in-app edits, undo/redo, and ordinary DOCX save/reopen. Inline objects are contents of their paragraph, so use that paragraph's ID and an appropriate range/command. User-facing UI and chat should never display these identifiers.

For ordinary DOCX packages, the app stores identities in a private XML part with a fingerprint of the Word package. If another application changes that package, the mapping is discarded rather than applied to potentially different content. Password-protected files currently have session-local identities. Always discover IDs again after opening a document.

Every tool response includes a document-instance version. Edits and activity-start require a matching version, checked inside the document's serial operation queue. A manual edit, Undo/Redo, or reopen invalidates stale requests. The version covers the whole document: reread even when another block changed. The engine prepares an edit on a copy and commits only after validation; failed operations do not partially replace the document.

Do not treat ID validity as permission to alter a block. Read its content and ensure it still matches the user's intended target. Keep the document URI, latest version, target ID, and a short task summary in working context; fetch content and command details as needed instead of retaining the entire document or registry.

Shimmer is a temporary activity indicator, not an edit lock or a tracked-change highlight. It must not imply that work continues after the agent finishes or stops. User interaction can clear it; a 30-second lease prevents abandoned highlights. Changes are unsaved until normal Save. Ordinary Undo is available, but selective agent-change rejection is a separate future review feature.
