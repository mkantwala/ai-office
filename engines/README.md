# Rust engine sources

AI Office uses Wordcraft for Word documents and Pdfcraft for PDF documents. `sources.json` pins their upstream revisions. The two binary-capable Git patches contain all local engine changes, including the new AI Office WebAssembly crates. Upstream source histories and license notices remain in their original repositories.

From the AI Office root, run `node scripts/prepare-office-engines.mts` to fetch those exact revisions and apply the patches into the ignored `.build/engines/` directory. The build script prefers those prepared sources, then falls back to the sibling checkouts used during development. `WORDCRAFT_SOURCE`, `PDFCRAFT_SOURCE` and `CARGO_TARGET_DIR` remain available as overrides.

The setup script verifies patch checksums and refuses to overwrite an existing checkout with a different version. Generated WebAssembly, dependencies, local documents, build output and profiles are not source files and are not committed. The original sibling working trees are preserved.
