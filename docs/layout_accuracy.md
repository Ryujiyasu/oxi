# Layout Accuracy — Detailed Progress

Oxi's layout engine is measured against Microsoft Word using two complementary metrics:

1. **SSIM (pixel-level)** — 177 real-world .docx documents (352 pages). Word GDI EMF (150dpi) vs Oxi GDI renderer (TextOutW, 150dpi).
2. **DML structural diff** — paragraph Y positions and line-break positions compared via Word COM API.

All specifications are derived from COM API black-box measurements — no DLL disassembly.

## Targeted Test Suite (49 documents)

| Metric | Value |
|---|---|
| **Average SSIM** | **0.9788** |
| Pages >= 0.95 | **48/49** (98%) |
| **DML perfect** (P\|dy\|=0, \|dch\|=0) | **35/49** (71%) |
| Average paragraph Y deviation | **0.02pt** |
| Average char-count deviation per line | **0.13** |

## 177-document Baseline Progress

| Date | avg SSIM | Pages >= 0.90 | Key Changes |
|------|----------|---------------|-------------|
| 2026-03-28 | 0.7884 | — | Baseline: 147 docs, grid snap, spacing collapse, justify, twips char width, GDI height ppem round |
| 2026-03-30 | 0.8083 | — | DML-driven improvement loop, GDI renderer pipeline |
| 2026-03-31 | 0.8152 | 79/157 (50%) | ceil_10tw line height, text_y_offset, table cell lineSpacing |
| 2026-04-01 | 0.8191 | 121/415 (29%) | pPr/rPr empty paragraph font, tab_stops, linesAndChars table row snap |
| 2026-04-02 | 0.8194 | 133/424 (31%) | Table border overhead fix, pixel perfect proof (GDI TextOutW), GDI width tables ppem 7-50 |
| 2026-04-03 | 0.8212 | — | CJK 83/64 eighth-pt floor, charGrid pitch, charSpace 1/4096pt |
| 2026-04-04 | 0.8286 | 150/437 (34%) | pBdr border overhead, bullet marker size, docDefaults lineSpacing table cell reset |
| 2026-04-05 | 0.8305 | 155/437 (35%) | Multiple spacing cumulative ceil, beforeLines/afterLines grid snap |
| 2026-04-06 | **0.8430** | 168/438 (38%) | LM0 line height formula, docGrid no-type, font alias, eastAsia fallback |
| 2026-04-07 | — | — | autoSpaceDE, font mapping, mixed-font line height, bold metrics |
| 2026-04-09 | 0.8292 | — | linesAndChars cumulative round, textbox v-anchor, table cell padding |
| 2026-04-10 | **0.8520** | — | leftChars indent, fullwidth symbols, font unification |
| 2026-04-12 | 0.8528 | — | Numbering, font widths, cumul carry/skip |
| 2026-04-13 | **0.8567** | — | Bottom margin fix, Multiple spacing ROUND, empty para ppr_rpr |
| 2026-04-14 | **0.8584** | — | 12 new OOXML elements parsed, 10tw char width rounding, cumulative raw model |

## COM-Confirmed Specifications (key examples)

| Specification | Behavior |
|---|---|
| **autoSpaceDE** | Adds 2.5pt only between Latin alphanumerics and CJK ideographs/kana. CJK punctuation does NOT trigger auto-space |
| **LayoutMode=0 multiple spacing** | Uses ROUND to 0.5pt (not CEIL); cumulative includes last line |
| **LayoutMode>=1 multiple spacing** | Uses CEIL to 0.5pt; cumulative excludes last line |
| **CJK 83/64 line height** | `(winAsc + winDes) * fontSize * 83/64`, no 1/8pt floor; round at final step |
| **linesAndChars grid lines** | `gl(n) = ((margin_tw + n*pitch_tw) / 10 + 1) * 10` — anchored to topMargin, ceil to 10tw |
| **doNotExpandShiftReturn** | Soft-break (Shift+Enter) lines are NOT justified in jc=both paragraphs |
| **Character width rounding** | Word rounds all character advance widths to 10tw (0.5pt) units |
| **Half-width Japanese font names** | "MS Mincho" -> Yu Mincho metrics, "MS PGothic" -> Yu Gothic (GDI fallback) |
| **Theme ea="" resolution** | Falls back to docDefaults eastAsia, then to system default CJK font |
| **Mixed-font line height** | When line has Latin text in a CJK 83/64 ascii font, that font's CJK height is included in max |
| **General Punctuation fullwidth** | Specific chars only: - ' dagger double-dagger ... per-mille prime etc. |
| **Bold-aware metrics** | When run is bold, lookup `{family} Bold` or `{family} Demibold` variant |

## OOXML Elements Parsed (04-14)

12 previously unparsed elements added:

| Element | Docs | Impact |
|---------|------|--------|
| `w:tblStylePr` | 163 | Table conditional formatting (shading, borders, bold, color) |
| `w:wordWrap` | 34 | CJK line break control |
| `w:adjustRightInd` | 26 | CJK grid right indent |
| `w:outlineLvl` | 184 | Outline level for TOC |
| `w:framePr` | 13 | Drop caps, positioned paragraphs |
| `w:tblPrEx` | 7 | Row-level table property exceptions |
| `w:textDirection` | 4 | Cell text direction |
| `w:textAlignment` | 4 | Line text alignment |
| `w:position` | 2 | Run vertical offset |
| `w:em` | 1 | Emphasis marks |
| `w:doNotExpandShiftReturn` | 62 | Compat: soft-break justification |
| `fontTable.xml` | 226 | PANOSE-1, charset, family, pitch |

## Method

- **Rendering**: Word GDI EMF (CopyAsPicture -> PlayEnhMetaFile) vs Oxi GDI (TextOutW), both at 150dpi
- **Pixel comparison**: SSIM (Structural Similarity Index) per page
- **Structural comparison**: Word COM `Information(6)` Y positions + line break analysis via `dml_diff.py`
- **Zero-regression rule**: any page that gets worse = revert. Net averages are informational only


---

## SSIM Progress Table (moved from README 2026-07-13)

The date-by-date progress table formerly in README.md. The chart in README.md
plots the same series; this table carries the per-date change notes.

| Date | avg SSIM | gate / Phase | Key Changes |
|------|----------|--------------|-------------|
| 2026-03-28 | 0.7884 | bottom-5 sum | Baseline: grid snap, spacing collapse, justify, GDI metrics |
| 2026-04-06 | **0.8430** | bottom-5 sum | LM0 line height, docGrid no-type, font alias, eastAsia fallback |
| 2026-04-10 | **0.8520** | bottom-5 sum | leftChars indent, fullwidth symbols, font unification |
| 2026-04-14 | **0.8584** | bottom-5 sum 2.8035 | 12 new OOXML elements, 10tw char width, cumulative raw model |
| 2026-04-18 | **0.8597** | bottom-5 sum 3.0597 | 4-agent parallel session: CJK wrap strict overflow, empty-br stub, hanging-indent, row-height, yakumono compat15 gate (+0.2562 bottom-5) |
| 2026-04-21 | **0.8625** | bottom-5 sum 3.2451 | LM0 first-line centering (Bug A, `(line_h - fontSize)/2` offset scales with font size — 46 gen2_* Title docs no longer -4.32pt shifted); footer first-type phantom fix |
| 2026-04-28 | **0.8699** | **Phase 1** (pagination) | Methodology redesign: gate moves from bottom-5 SSIM sum to per-doc page-match correctness. Bottom-5 cascade plateau (R21-R34) revealed SSIM single-gate cannot move past structural mismatches. Phase 1 measures whether each Word paragraph lands on the same page as Oxi |
| 2026-05-08 | **0.8855** | Phase 1 37/55 | Day 14 leading-ws absorbs indent (+0.0098), Day 16 cs inheritance (+0.1066), Day 18 broad merge_run_style (+0.0252), Day 26 table row grid-snap removal (+0.2138), Day 28 adjustLineHeightInTable flag-conditional cell snap rule, Day 29f Times New Roman space data fix |
| 2026-05-12 | **0.8893** | **Phase 1 43/55** | Day 33 part 57 wrap_w uses cell_w (191cb + 1636 PASS), Day 33 part 59 page-break order fix preserves current-line text on hard break (cb8be7 PASS). Phase A+B (cursor advance + page-break decision precision) refactor commitment producing first concrete wins after 7+ sessions of investigation |
| 2026-05-12 | 0.8893 | Phase 1 42/55 (corrected) | Day 33 part 62 measurer fix: `measure_pagination_oxi.py` text concat now sorts by (y, x) instead of x-only — previously, multi-line wrapped paragraphs had their characters interleaved across lines, making the matcher unable to align them with Word. Fix tipped 04b88e to PASS, but revealed 3 docs (31420af, b837808, db9ca18) that previously appeared PASS due to insufficient matches. Methodology correction, not a layout regression |
| 2026-05-12 | **0.8892** | **Phase 1 43/55** | Day 33 part 65 (R7.18): body page-break check now uses natural line height (ascent+descent, no grid leading) instead of full grid line height — Word allows the leading portion of a grid-snapped line to extend into the bottom margin. COM-verified via db9ca18 paragraph 37: Word fits a line at y=758 whose grid bottom is 776 (5.25pt past pgBot 771). Companion fix: widow_control inheritance now propagates the explicit flag through the style chain, so widowControl=0 set on Normal correctly disables the orphan check for descendants. db9ca18 FAIL→PASS (3 pages matching Word). 0 PASS→FAIL transitions |
| 2026-05-12 | **0.8895** | **Phase 1 44/55** | Day 33 part 69 (R7.24): preserve fixed-layout (`<w:tblLayout w:type="fixed"/>`) table column widths — previously Oxi shrank the last column when grid_columns sum exceeded content_width (correct for autofit, wrong for fixed). 7-session a47e6 investigation: 21.1pt wrap-width loss caused "fullwidth+年月日" paragraph to overflow by 0.55pt → wrap to 2 lines, +25pt row 0 over-pump, +1.4pt at pi=2 → +1 page. 1-line fix tipped a47e6 to PASS (0.5 → 1.0) and improved d4d126 (0.8 → 0.857). Methodology lesson: estimate-vs-render diagnostic ≠ real cause; render-vs-render is the correct layer |
| 2026-05-12 | **0.8932** | **Phase 1 46/55** | Day 33 part 71 (R7.28): vMerge=restart cells excluded from row height (both estimate and render-side max). Word distributes a vMerge=restart cell's content across the entire vMerge span; the restart row's own height is set by non-merged cells. COM-verified on de6e t5 row 13 (Word 33pt, Oxi was 238pt → now 32.35pt). Previously vMerge=continue was already excluded; this commit extends the same rule to "restart". Phase 1: 31420af FAIL→PASS (0.8→1.0), 6514 FAIL→PASS (0.529→1.0); a1d6 0.556→0.875 (still FAIL, 1 outlier from PASS). 0 PASS→FAIL. Mean SSIM net +0.0037 across 410 pages. Two-line fix at mod.rs:5707 + 6404 |
| 2026-05-29 | — | **Phase 1 54/55** | Continued pagination fixes (cell-paragraph spacing collapse S427, etc.) lift Phase 1 to 54/55. The only remaining FAIL is `3a4f9f`, a split-table document Oxi paginates to 94 pages vs Word's 8 |
| 2026-05-29 | — | **Phase 2** (element IoU) | Gate moves to per-element bounding-box IoU; plateaus at mean IoU ≈ **0.9692**. Phase 2's median-dy subtraction absorbs uniform per-table offsets, so the IoU ≥ 0.99 entry bar proved structurally unreachable — which is precisely why the real remaining error (a uniform table-top offset, visible only in pixels) was invisible to it. See [CLAUDE.md](CLAUDE.md) |
| 2026-05-30 | per-page **0.8862** · per-doc **0.9235** | **Phase 3** (SSIM) | Primary gate switches back to pixel SSIM (mean ≥ 0.99 + bottom-N floor) on a freshly recomputed baseline. SSIM is the only metric that sees the uniform table-top offset Phase 2 hid. Phase 1 (54/55) and Phase 2 (0.9692) are kept as regression sentinels |
| 2026-06-03 | per-page **0.9126** | **Phase 3** · Phase 1 54/55 | R35 yakumono capacity-budget line breaking (S475/S476, docGrid `lines`+`linesAndChars`), then a 36-doc correctness sweep shipping localized coverage fixes: floating-textbox z-order (S478), 144 pt footnote separator (S479), dash-dot art borders (S480), explicit nil-cell-border suppression (S482), Word "final" revision view (S483), **upright CJK vertical writing** (S489), ellipse ○ option-markers (S490) |
| 2026-06-12 | per-page **0.9189** | **Phase 3** · Phase 1 54/55 | Two weeks of COM-measured spec re-derivations (S495-S548): `lineRule=exact` text bottom-aligns in its box (S495), cell inline images (S533), inline drawing canvases (S535/S537), three justification bugs — style-chain `jc` inheritance, explicit `jc=left` vs style default, jc-left natural breaks (S539/S540) — demand *oikomi* with a line-total fs/2 budget under Word-2010 compat (S543-S546), **character-width trio**: UPM-256 halfwidth = fs/2 exact, autoSpaceDE/DN = fs/4 true-space, yakumono pair-halving gated on `w:kern` with the full 26×26 pair table (S546/S547), compat-15 oidashi-not-burasagari + exact-line page-break threshold (S548). The single Phase-1 FAIL (`3a4f9f`) is down to 3 paragraphs (one page early), all traced to the inline-image text-line model |
| 2026-06-30 | per-page **0.9253** · per-doc **0.9487** | **Phase 3** · Phase 1 86/87 | Corpus expanded to 87 body-text docs (Phase 1) / 238 SSIM-scored docs. A long pagination + fidelity run (S559-S707): the *char-budget wall* (per-line 約物 demand-compression model, derived from a controlled synthetic dataset + Word-PDF render-truth — gate, mechanism, half-em/0.32em caps, ぶら下げ), the **form-family row-height re-derivation** (drift is a cell-tcBorder border-box overhead, not CJK line-height — S648/S660/S661/S666), the **gen2/Latin vertical & horizontal stack** (no-type-docGrid hhea line height S671, render-x word separation S672, glyph-centering S614/S670, DWrite kerning-off S668), **font substitution** (Latin-only eastAsia → MS Mincho S634, embedded Zen Old Mincho S612z), multi-column + vertical/bidi section flow (S637/S638/S678/S679), and a coverage sweep graduating **vertical writing, tate-chu-yoko (縦中横), ruby, warichu (割注), emphasis marks, run/paragraph shading, and character borders** (S654-S707). Found increasingly via a Word-vs-LibreOffice bug-finder (pages where LibreOffice ≈ Word but Oxi ≠ Word = a fixable Oxi bug) and a feature-injection perturbation harness. The sole remaining Phase-1 FAIL (`tokyoshugyo`) is the legacy compat≤14 約物-oikomi body wrap |
| 2026-07-03 | — | **Phase 1 87/87 = 100%** | **Pagination COMPLETE**: every paragraph of every corpus document lands on the same page as Microsoft Word (S713-S722 closed the last doc, `tokyoshugyo`). The same day, an **adversarial probe harness** (74 self-authored feature probes gated against real Word) opened the next frontier: 11 fixes shipped in one day (multi-column paragraph split S723, gate hardening S724, paragraph-tail tolerance S725, tall-footer/footnote/header reservations S726/S727/S731, tblHeader repeat S728, continuous-section margins + zero-height break mark S729/S730, evenPage/oddPage parity blank pages S732, column breaks S733) — all with zero corpus regressions |
| 2026-07-12 | per-page **0.9370** · per-doc **0.9587** | **Phase 3** · Phase 1 87/87 | Row-height border-box cursor model (ROWBOX2 bundle), derived cell char-budget (CELLPAIR), SAMPLE watermark (WordArt em/advance-fit), Latin underline, English real-document corpus opened (6 UK/US government docs added to the pagination gate; 3 already PASS) — stale-LRPB root-caused, justified space-shrink, numbering-indent precedence, en/em-dash font class (S758-S801) |

**Phase-based gate** (since 2026-04-28): the merge gate is currently **Phase 3 — pixel SSIM** (mean ≥ 0.99 + bottom-N floor), active since 2026-05-30. Earlier phases are kept as regression sentinels: **Phase 1** pagination correctness (per-paragraph page match, 54/55) and **Phase 2** element IoU (mean 0.9692). Phases 1 and 2 each plateaued below their entry bars for structural reasons — pagination on one split-table outlier, IoU because its median-dy subtraction hides uniform table offsets — so the gate advanced to the metric that can see the remaining pixel error. The phase-based methodology is documented in [CLAUDE.md](CLAUDE.md) under "Merge gate".

---

## Ra: Empirical Convergence (moved from README 2026-07-13)

### Ra: Empirical Convergence

Oxi's Word compatibility is built on empirical reverse engineering, not speculation. The premise was tightened after Sessions 38-45 falsified some of the founding axioms (R30 measurement bug, R33 41-page regression, R21 plateau). What remains:

- Word's layout is **deterministic** — same input always produces the same output. This is the basis for measurement-driven specification derivation
- Hypotheses are **falsifiable via COM measurement or pixel diff**. Speculation is not a basis for layout changes
- Word output is the ground truth for the **fidelity goal** (matching Word's render). OOXML spec is the ground truth for the **correctness goal** (parser, IR semantics). When the two disagree (undocumented Word quirks), fidelity wins for rendering, correctness wins for parser/IR
- The merge gate is **phase-based** — Phase 1 (pagination correctness), Phase 2 (element IoU ≥ 0.99), Phase 3 (SSIM ≥ 0.99 + bottom-N floor). SSIM remains the long-term goal but is gated only at Phase 3; tracked at every phase as a regression sentinel

This is not "best effort." It is a measurement-driven convergence loop, where each merge moves the gate's primary metric or is rejected. The same loop transfers to ODF parity once the v2 baseline is in place — the reference renderer changes (LibreOffice headless), the methodology does not.

### Implementation Gap: ODF Rendering Parity

The most critical task for v2. Oxi's current Ra loop targets SSIM = 1.0 against Microsoft Word for .docx. The EU public-sector market needs the same fidelity for .odt. The methodology transfers — deterministic reference output, measurement-driven specification derivation, phase-based merge gate — but every layout-engine entry point currently presupposes OOXML structures.

The work splits into three:
1. **ODF parser** — `.odt` → IR. The IR is already format-agnostic; the parser is additive
2. **ODF-specific layout rules** — paragraph / list / table semantics that differ from OOXML need explicit branches, not silent OOXML defaults
3. **ODF reference baseline** — pick a deterministic reference renderer (LibreOffice headless is the obvious candidate, given its 20-year status as the ODF reference implementation) for the SSIM gate
