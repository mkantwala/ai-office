# VS Code Agents Instructions

This file provides instructions for AI coding agents working with the VS Code codebase.

For detailed project overview, architecture, coding guidelines, and validation steps, see the [Copilot Instructions](.github/copilot-instructions.md).

## Wordcraft upstream reference

Before implementing a Word feature, inspect the original Wordcraft implementation at `../wordcraft-upstream` (repository: https://github.com/storytold/wordcraft) or `../wordcraft`. These separate checkouts are references, not application build inputs. The Word engine is vendored in `engines/wordcraft/`: copy the upstream code you need into that vendored tree and commit it, instead of patching or pinning upstream. Keep the vendored workspace trimmed to what the DOCX viewer needs (no egui, MCP or desktop apps), preserve upstream LICENSE/NOTICE/ATTRIBUTION files, and verify adapted behavior with focused tests, including DOCX round trips when serialization changes.

Reuse upstream document, layout, rendering and DOCX behavior in Rust where practical. Use TypeScript for VS Code integration, toolbar/property controls and agent orchestration. Check both engine support and upstream UI/command wiring before reimplementing a feature. Preserve internal block identities, guarded edits, native history, comments and Keep/Undo review when adapting upstream functionality.

Commit Word engine adaptations directly in the vendored `engines/wordcraft/` tree; there is no Wordcraft patch or pin anymore. Changes only in the ignored `.build/engines/wordcraft` checkout or a sibling reference clone are not deliverable. Review upstream license and attribution requirements when copying code.

## Pdfcraft upstream reference

Before implementing a PDF feature, inspect `../pdfcraft-upstream` (https://github.com/storytold/pdfcraft). Keep it separate from the pinned build checkout. Reuse native PDF operations in Rust; use TypeScript for application controls, file workflows and agent orchestration. Persist adaptations in `engines/pdfcraft.patch` and update its checksum in `engines/sources.json`. Verify saved/reopened PDFs, permissions and page geometry. Follow `extensions/ai-office-pdf/PDF_IMPLEMENTATION_PLAN.md` for parity gaps and bounded agent capability discovery; do not advertise unimplemented conversions or Word-style block behavior for PDF objects.
