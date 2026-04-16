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

### Autonomous Loop Procedure (Tightened 2026-04-09)
1. Read `docs/spec/word_layout_spec_ra.md`, identify unresolved questions
2. Select the highest-impact unresolved question
3. Create Python COM measurement script in `tools/metrics/`
4. Execute on **≥3 distinct documents** + **a minimal repro doc you author yourself**
   that isolates only the spec under test. Append all results to
   `pipeline_data/ra_manual_measurements.json`
5. Analyze. Spec moves to **"hypothesis"** until 3+ docs + minimal repro all agree.
   Only then promote to **"confirmed"**. Single-doc observations are NEVER confirmed
6. Implement confirmed specifications in Rust
7. **Rebuild caches**: `wasm-pack build` + copy to web/ + delete `pipeline_data/oxi_png/`
   (and DML cache if structure changed). Stale caches = fake SSIM (see baseline-drift)
8. **Primary verification**: Word EMF vs Oxi GDI pixel diff via `pipeline.verify`
   on the full baseline. DML diff is **secondary only** (Information(6) is line-box
   top, not glyph top — see com-info6-caveat)
9. **Bottom-N floor rule** (2026-04-16 FINAL, phase-based):
   - Compute `pre_b3 = sum of 3 lowest min(SSIM) per doc` from pre-fix baseline
   - Phase 1 (`pre_b3 ≤ 1.5`): N=3. Merge if post-fix bottom-3 sum > pre_b3
   - Phase 2 (`pre_b3 > 1.5`): N=5. Merge if post-fix bottom-5 sum > pre-fix bottom-5 sum
   - Phase切替は不可逆。rank ≥(N+1) は無視。Averages and net Δ are NOT the gate
   - Phase 2 への自動拡大は whack-a-mole nesting（rank 4-5 の隠れregressionが後で噴出）を抑える
10. If an EXCEPTION must be carved out for a doc → the original spec is **wrong**,
    not incomplete. Re-derive it from a richer input space (PANOSE, proportional
    flag, etc.) before re-implementing. Do not stack exceptions
11. Return to step 1

### Domain Status (2026-03-28)
- **char_width**: Fallback implemented (MS UI Gothic). No effect on current test documents
- **page_break**: widow/orphan, keepNext/keepTogether implemented. Mid-paragraph page break fixed (net +0.041)
- **spacing**: Collapse (max(sa,sb)) implemented. net +0.71
- **line_height**: Table cell reset implemented. net +0.66
- **grid_snap**: Implemented
- **justify**: docDefaults jc=both inheritance fixed. Justify enabled for all documents
- **SSIM: 0.7496 → 0.8584** Baseline: 177 documents, 352 pages (GDI)
- **w:w text scale (2026-04-14)**: Parse + layout support for character width percentage (w:w val="80" = 80% width)
- **fontTable.xml (2026-04-14)**: Parse PANOSE, charset, family, pitch for font substitution strategy
- **bracketPair Bézier (2026-04-14)**: Curved bracket rendering with κ=0.5523 quarter-circle approximation
- **char_width (2026-03-30)**: Twips-based width calculation (round(advance*fontSize*20/UPM)/20). Matches Word line breaks
- **GDI width overrides**: 9 fonts with complete GDI width tables (1055KB)
- **GDI renderer**: Pipeline switched to oxi-gdi-renderer (TextOutW) for pixel-accurate comparison
- **DML diff tools**: word_dml_extract.py + dml_diff.py for structural layout comparison
- **margin fix (2026-04-10)**: Exact twip margins (removed round_10tw), empty para CJK font, hangingChars parse
- **is_fullwidth fix (2026-04-10)**: Added 7 Unicode blocks (Arrows, Math Operators, Letterlike Symbols, etc.) to CJK fullwidth table. Fixes → overlap
- **twip-priority indent (2026-04-10)**: When both twip and *Chars indent values exist, twip takes priority (pre-computed by Word)
- **LM2 offset fix (2026-04-13)**: Removed centering offset from cursor_y start (cursor starts at topMargin, centering via text_y_offset)
- **VML bracket (2026-04-13)**: VML shape type 185 (double bracket 〔〕) parsed and rendered
- **table x/border (2026-04-13)**: Table border x = margin-padding-border/2. Row height excludes inside-H border overhead
- **bottom margin fix (2026-04-13)**: Exact bottom margin (no 10tw round). Top margin rounds for content Y, bottom stays exact for page break limit
- **Multiple spacing CEIL (2026-04-13)**: LM0 multiple spacing cumulative round uses CEIL not ROUND (MS Mincho 10.5pt×1.15: 310.5tw→320→16.0pt)
- **empty para ppr_rpr (2026-04-14)**: Empty paragraph raw_spaced_tw uses ppr_rpr font (matching line_height_for_line_inner)
- **10tw char width (2026-04-14)**: Character widths rounded to 10tw (0.5pt) matching Word. COM-confirmed 13 font/size, 181 chars
- **cumul raw model (2026-04-14)**: Multiple spacing cumulative raw position model. COM-confirmed 8/9 positions
- **baseline_adjust disabled (2026-04-14)**: No per-fragment baseline shift. CJK and Latin digits share same Y
- **Remaining improvements**: table Y start position, charGrid文字詰め, table cell floating shapes, textbox charGrid, 1ec overflow

### Measurement Template
Correct method for measuring line height is "Y coordinate difference between 2 paragraphs":
```python
y1 = doc.Paragraphs(1).Range.Information(6)  # wdVerticalPositionRelativeToPage
y2 = doc.Paragraphs(2).Range.Information(6)
gap = y2 - y1  # = line_height + spacing
```
`Format.LineSpacing` returns the setting value only, not the actual rendered height.

### Pixel-Driven Improvement Loop (Revised 2026-04-09)

**Word EMF vs Oxi GDI pixel diff** is the primary improvement signal.
DML diff was the primary signal previously, but COM `Information(6)` returns
line-box top (not glyph top), so DML |dy| does not validate `text_y_offset` and
can mask real pixel regressions. DML diff is now **secondary** — useful for
narrowing down *which block* differs, not for confirming a fix.

**Tools:**
- `tools/oxi-gdi-renderer/` — GDI renderer (TextOutW) for Oxi side
- Word EMF path: CopyAsPicture → PlayEnhMetaFile (see ssim_progress)
- `pipeline.verify` — full-baseline pixel diff
- `tools/metrics/dml_diff.py` — secondary, block-level diagnosis only
- `tools/metrics/word_dml_extract.py` — Word COM position cache
  (regenerate whenever layout shape changes)

**Loop:**
1. Pick a single document where Oxi vs Word EMF differs
2. Author a **minimal repro** that isolates the suspected spec
3. COM-measure the repro on ≥3 variants → spec hypothesis
4. Implement
5. Rebuild WASM + clear `pipeline_data/oxi_png/`
6. Pixel diff the repro: must match Word EMF exactly. If not, spec is wrong
7. Run `pipeline.verify` on full baseline
8. **Bottom-N floor check**: per-doc min(page SSIMs), sort ascending.
   N=3 if pre-fix bottom-3 sum ≤ 1.5 else N=5. Merge if post-fix bottom-N sum
   > pre-fix bottom-N sum. Rank ≥(N+1) docs irrelevant
9. Commit when bottom-N sum rises; revert only if it falls or stays equal

### Critical Rules
- No DLL disassembly. Black-box measurement via COM API only
- Never implement from speculation. Always confirm values via COM measurement first
- **Bottom-N floor rule** (2026-04-16 FINAL, phase-based):
  N=3 in Phase 1 (pre-fix bottom-3 sum ≤ 1.5), N=5 in Phase 2 (>1.5).
  Merge if bottom-N sum strictly increases. Rank ≥(N+1) docs outside the gate.
  Phase 2 expansion to bottom-5 prevents whack-a-mole: rank 4-5 regressions
  no longer go silent and re-emerge as future bottom-3 targets
- **3-doc + minimal-repro rule**: a spec is "hypothesis" until 3 distinct real
  docs AND a self-authored minimal repro all agree. Single-doc → never confirmed
- **No EXCEPTION stacking**: if a confirmed spec needs a per-font / per-doc
  carve-out, the spec itself is wrong. Re-derive from a richer input space
- **Cache hygiene**: rebuild WASM + delete `pipeline_data/oxi_png/` (and DML
  cache when shapes change) before every verify run. Stale caches = fake SSIM
- **Information(6) is not glyph top**: do not use DML |dy| as the merge gate.
  Pixel diff (Word EMF vs Oxi GDI) is the only ground truth

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
