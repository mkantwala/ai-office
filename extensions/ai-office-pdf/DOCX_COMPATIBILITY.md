# DOCX compatibility

Sangam One saves standard DOCX packages. Agent intent previews, shimmer and cursor labels are editor-only; they never become document paragraphs or change DOCX formatting. Committed edits use the existing native document engine and tracked review.

## Verification on 2026-10-20

- All 231 tests in the patched engine's `wordcraft-docx` package passed across 14 suites (`cargo test -p wordcraft-docx --tests`). These cover import/export and preservation regressions, including relationships, styles, numbering, fields and unsupported-part passthrough.
- All 62 focused document, OpenCode and browser tests passed, including preview cleanup and unchanged document bytes while previewing an edit.
- Four synthetic DOCX round trips passed 148 package and formatting checks. Two documents were generated independently with python-docx; two were rewritten by LibreOffice before importing them into Sangam One. Each pair covered formatting and multilingual content. A guarded agent text replacement was applied to each formatting document before export.
- LibreOfficeDev 26.8 independently opened and rendered every original and exported sample. Rendered text and page counts matched, allowing the requested replacement. The exported formatting sample was visually inspected.
- Checked properties included direct fonts, sizes, emphasis and colours, paragraph alignment and spacing, margins, page dimensions, headers/footers, table text, merged cells, repeating table headers, images and page borders. Exported ZIP entries, XML and internal relationship targets were checked. Preview intent text was absent from exported documents.

## Limits

These checks establish compatibility for the tested cases, not lossless preservation of every possible DOCX. Microsoft Word and Google Docs were not directly exercised in this environment. Different fonts and layout engines can change line wrapping, pagination and object positioning; Google Docs can also convert unsupported features on import.

The underlying Wordcraft engine has known preservation limits for unknown elements and attributes inside document stories, style definitions and numbering definitions, and for regenerated themes. Complex or business-critical templates therefore still require opening the exported file in the intended Word/Google Docs version and reviewing its layout. See the upstream `docs/file-format-parity.md` in the vendored Wordcraft tree at `engines/wordcraft` (and in the reference clones beside this repository) for feature coverage; this fork does not claim complete Word parity.
