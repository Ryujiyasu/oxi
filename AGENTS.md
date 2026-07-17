# Oxi Development Guide

## Project Overview

Oxi is an OSS document processing suite built with Rust + WebAssembly.
The goal is to parse, render, and eventually edit .docx / .xlsx / .pptx files natively in the browser.

## Architecture

- **oxidocs-common**: Shared OOXML utilities (ZIP, XML, relationships) — renamed from oxi-common 2026-07-07 (crates.io oxidocs-* unification; the oxi-* prefix is squatted by an unrelated active project)
- **oxidocs-core**: .docx engine — parser, IR, layout, font metrics
- **oxicells-core**: .xlsx engine — parser, IR
- **oxislides-core**: .pptx engine — parser, IR
- **oxidocs-wasm**: WebAssembly bindings via wasm-bindgen (renamed from oxi-wasm 2026-07-07; web/ artifacts are oxidocs_wasm.js / oxidocs_wasm_bg.wasm)
- **web/**: Vanilla JS + Canvas editor

## IR Design Principles

The Intermediate Representation (IR) must be language-agnostic and NOT depend on Word-specific internals.
Structure: Document → Page → Block (Paragraph | Table | Image) → Run

## Font Metrics

Font files are NEVER committed to the repository. Only pre-computed metrics tables are included.
Metrics are measured on GitHub Actions Windows runners and stored as data tables.

## Japanese Typography (Kinsoku)

Priority order:
1. Kinsoku processing (line-start/line-end prohibited characters)
2. Character spacing (justification)
3. Ruby (furigana)
4. Vertical writing (basics only)

Reference: JIS X 4051

## Testing

- Golden tests: render .docx with Oxi, compare pixel-by-pixel against Word screenshots
- Test fixtures go in tests/fixtures/
- CI: `cargo test`, `cargo clippy`, `wasm-pack build`

## Build Commands

```bash
cargo build                          # Build all
cargo test                           # Run tests
cargo clippy                         # Lint
cd crates/oxidocs-wasm && wasm-pack build --target web  # Wasm build (browser editor)
cd tools/oxi-gdi-renderer && cargo build --release  # GDI renderer (verify pipeline)
```

**Three render paths exist** — all must be rebuilt when source changes:
- **WASM** (`crates/oxidocs-wasm/pkg/`) — used by browser editor at `web/`
- **Native GDI renderer** (`tools/oxi-gdi-renderer/target/release/`) — used by
  `measure_pagination_oxi.py` and as fallback when `OXI_USE_GDI=1`
- **Native DWrite renderer** (`tools/oxi-dwrite-renderer/target/release/`) —
  **DEFAULT for `pipeline.verify`** since Session 50 (commit 04cc22d).
  DirectWrite matches Word's text engine glyph metrics

`wasm-pack build` does NOT rebuild either native renderer. Skipping a rebuild
causes verify to run with a stale pre-patch binary — pre-patch output is
compared to pre-patch baseline, producing false-positive ~0 net Δ even when
the patch has large impact. **2026-04-26 incident**: e4b8734 shipped with
claimed +0.0006 net Δ; rebuild revealed -0.0911 actual; reverted via 2e0e1f0.
**2026-05-07 incident**: Session 56 first Finding 3 verify run reported "0
improved 0 regressed" because DWrite was stale (cargo build in oxi-gdi-renderer
dir does NOT trigger oxi-dwrite-renderer rebuild). Always rebuild **BOTH**
GDI and DWrite before pipeline.verify.

## Session memory (IMPORTANT — read first)

Ra loop procedure, current state, and all derived-specification notes live in
**Codex.local.md** (gitignored, this directory). Read it BEFORE starting any
layout/Ra work. Research documents under docs/spec/ are likewise local-only
(untracked). Do not move their content into tracked files.

Ship-notes / Domain Status entries / derivation findings are appended to
**Codex.local.md, never to this file**. This file is public; the research
memory is not.

## License

Three-layer model (decided 2026-06-13):
- **Core engine crates** (oxidocs, oxidocs-common, oxidocs-core, oxicells-core, oxislides-core, oxipdf-core, oxihanko, oxidocs-cli, oxi-desktop): **MPL-2.0**. Every `.rs` file in these crates must carry the MPL-2.0 Exhibit A header (file-level copyleft makes per-file headers load-bearing — add the header to any NEW source file)
- **Bindings** (oxidocs-wasm, oxidocs-python): **MIT OR Apache-2.0** (Rust ecosystem dual-license convention)
- **Conformance corpus** (self-authored repros in tools/golden-test/repros/): **CC BY-SA 4.0**; measurement code is MPL-2.0

Contributions: DCO only (`git commit -s`), no CLA. All third-party crate licenses must be MPL-2.0-compatible (MIT, Apache-2.0, BSD).
