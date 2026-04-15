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
