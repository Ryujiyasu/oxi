# Oxi Development Guide

## Project Overview

Oxi is an OSS document processing suite built with Rust + WebAssembly.
The goal is to parse, render, and eventually edit .docx / .xlsx / .pptx files natively in the browser.

## Architecture

- **oxi-common**: Shared OOXML utilities (ZIP, XML, relationships)
- **oxidocs-core**: .docx engine — parser, IR, layout, font metrics
- **oxicells-core**: .xlsx engine — parser, IR
- **oxislides-core**: .pptx engine — parser, IR
- **oxi-wasm**: WebAssembly bindings via wasm-bindgen
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
cd crates/oxi-wasm && wasm-pack build --target web  # Wasm build
```

## Ra: Autonomous Word Specification Analysis Loop

At the start of each session, check the current state and continue autonomous specification analysis.

### State Management
- Specification: `docs/spec/word_layout_spec_ra.md`
- Measurement data: `pipeline_data/ra_manual_measurements.json`
- SSIM baseline: `pipeline_data/ssim_baseline.json`

### Autonomous Loop Procedure
1. Read `docs/spec/word_layout_spec_ra.md`, identify unresolved questions
2. Select the highest-impact unresolved question
3. Create Python COM measurement script in `tools/metrics/`
4. Execute and append results to `pipeline_data/ra_manual_measurements.json`
5. Analyze results and update specification
6. Implement confirmed specifications in Rust
7. Run `python -m pipeline.verify` for SSIM regression check
8. If net positive → commit; if negative → revert
9. Return to step 1

### Domain Status (2026-03-28)
- **char_width**: Fallback implemented (MS UI Gothic). No effect on current test documents
- **page_break**: widow/orphan, keepNext/keepTogether implemented. Mid-paragraph page break fixed (net +0.041)
- **spacing**: Collapse (max(sa,sb)) implemented. net +0.71
- **line_height**: Table cell reset implemented. net +0.66
- **grid_snap**: Implemented
- **justify**: docDefaults jc=both inheritance fixed. Justify enabled for all documents
- **SSIM: 0.7496 → 0.8083 (+0.059)** Baseline: 147 documents, 399 pages
- **char_width (2026-03-30)**: Twips-based width calculation (round(advance*fontSize*20/UPM)/20). Matches Word line breaks
- **GDI width overrides**: 9 fonts with complete GDI width tables (1055KB)
- **GDI renderer**: Pipeline switched to oxi-gdi-renderer (TextOutW) for pixel-accurate comparison
- **DML diff tools**: word_dml_extract.py + dml_diff.py for structural layout comparison
- **Remaining improvements**: table row height (0.25pt border overhead), 1ec overflow, heading line height

### Measurement Template
Correct method for measuring line height is "Y coordinate difference between 2 paragraphs":
```python
y1 = doc.Paragraphs(1).Range.Information(6)  # wdVerticalPositionRelativeToPage
y2 = doc.Paragraphs(2).Range.Information(6)
gap = y2 - y1  # = line_height + spacing
```
`Format.LineSpacing` returns the setting value only, not the actual rendered height.

### DML-Driven Improvement Loop

DML (layout structure) match is a **prerequisite** for pixel match.
Pixel comparison (SSIM) is only for final verification. The improvement driver is DML comparison.

**Tools:**
- `tools/metrics/word_dml_extract.py` — Extract Word COM positions → JSON cache (`pipeline_data/word_dml/`)
- `tools/metrics/dml_diff.py` — Structural diff: Oxi layout_json vs Word DML cache
- `tools/oxi-gdi-renderer/` — GDI renderer for pixel-level SSIM verification

**Loop:**
1. `word_dml_extract.py` to cache Word positions (once)
2. Code fix
3. `dml_diff.py` for structural diff (1 min, no rendering needed)
4. When diff improves → next issue
5. After batch of fixes → `pipeline.verify` for full SSIM check

**Rationale:** Pixel loop (GDI render → SSIM) takes 10 min/cycle. DML diff takes 1 min/cycle = 10x faster.
DML diff also gives precise, actionable feedback ("table row 3 Y off by 0.25pt") vs SSIM ("something looks different").

### Critical Rules
- No DLL disassembly. Black-box measurement via COM API only
- Never implement from speculation. Always confirm values via COM measurement first
- Revert any change that decreases SSIM (net positive rule)

### No Excuses by Design
Ra is built on the premise that there are no valid excuses for layout differences.
- Word's layout is **deterministic** — same input always produces the same output
- Every value is **measurable via COM API** — Y coordinates, line heights, character widths, paragraph spacing
- Any difference = unimplemented specification = identifiable via COM measurement → fixable
- **Specifications are finite. Measurement results are permanent assets.** Once measured, never needs re-derivation
- Fixing one specification gap improves multiple documents simultaneously (convergent structure)
- Not "cannot do" but "not yet done" — purely a matter of measurement count and implementation time

## License

MIT. All third-party crate licenses must be MIT-compatible.
