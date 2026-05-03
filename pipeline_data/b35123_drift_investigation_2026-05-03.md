# b35123fe8efc_tokumei_08_01 p.1 vertical drift — investigation handoff

**Date**: 2026-05-03
**Branch**: session50-visual
**SSIM**: baseline 0.6664, current 0.6661 (essentially unchanged by main's 0.667+pair-rule fix)
**Status**: #1 worst doc by p.1 baseline SSIM in the entire 177-doc corpus
**Doc shape**: 2 pages, table-heavy form (匿名データの適正管理措置の内容)

---

## 1. Measurement method

`c:/tmp/measure_b35123_v2.py` (already on disk; not committed).

- Word side: `Range.Information(WD_VPOS=6)` on collapsed start range (R30 fix)
  for each paragraph until 3 consecutive p.2 paragraphs seen → 57 p.1
  paragraphs captured.
- Oxi side: `oxi-dwrite-renderer --dump-layout=...` then group `text`
  elements by **rounded Y (0.5pt bucket)**, taking concatenated text per Y.
  46 unique-Y "lines" extracted. (Note: para_idx grouping yielded only 4
  groups because the form's table cells share paragraph index — Y bucketing
  is the right granularity for this doc.)
- Pair: for each Word paragraph, find Oxi line with nearest Y; report
  `(word_y, oxi_y, oxi_y − word_y)` and per-step jump.

Raw output: `c:/tmp/b35123_v2.log`.

## 2. Drift trajectory (oscillating, NOT monotonic)

Unlike ed025 (which had 4 monotonic jumps with cumulative drift), b35123
shows **oscillating drift between adjacent paragraphs**, with several
±10pt swings:

| paragraph range | drift range | pattern |
|---|---:|---|
| 1 – 11 | +2.0 to +3.5pt | small monotonic increase |
| 12 – 21 | +0.5 to +2.5pt | unstable (cell oscillation) |
| 22 – 28 | −3.0 to +0.5pt | **Oxi above Word** |
| **29 – 31** | **+7.0 to +8.0pt** | **single +10.5pt jump** |
| **32** | **−9.0pt** | **−17pt swing** ("ーブ" line) |
| 33 – 35 | +8.0pt | swings back +17pt |
| 36 – 38 | +7 to −8pt | bouncing per cell |
| 39 – 45 | +0.5 to +5.5pt | |
| 46 – 57 | −2 to 0pt | end of page, Oxi above |

Key feature: **drift signs flip multiple times** (positive → negative → positive),
indicating that this isn't a single accumulating bug but per-row mismatches.

## 3. Per-jump diagnoses

### 3.1 Para 29 jump (+10.5pt) and Para 32 swing (−17pt)

Around paragraphs 29–35 there's a +10.5pt jump up, then −17pt swing down,
then +17pt back up. The +10.5pt+17pt+(−17pt)+17pt pattern is consistent
with "Oxi places one row 8pt taller than Word, then a stray fragment
appears in a different cell" or similar table-cell row-height divergence.

Specifically:
- Para 29 Word_y=381.50, Oxi_y=389.50 (Oxi 8pt below)
- Para 30, 31 same +7 to +8pt drift (consistent for the next 2 paras)
- Para 32 ("ーブ" — 2-character fragment, likely wrap continuation):
  Word_y=380.50 (same row as 29!), Oxi_y=371.50 (Oxi 9pt ABOVE Word)
- Para 33 onwards back to +8pt

Working hypothesis: there's a table row around y=380 in Word containing
multiple paragraphs (`para 29 = "...対する規"`, `para 32 = "...ーブ"`).
Word renders them at different X within the same row (same Y), but Oxi
puts them at different Y values (treating them as sequentially stacked).

The +8pt extra height in Oxi suggests Oxi adds an extra body line
(~8pt at 11pt body) per affected row.

### 3.2 Repeated text paragraphs (paras 25–28, 33–35, etc.)

Paras 25, 26, 27, 28 all start with `（）人の場合、匿名データを…` —
identical openings. Same for paras 33–35 (`部局の／匿名データ…`).

This is consistent with **multiple cells in the same table row containing
similar text** (e.g., a 「区分」 cell + 「内容」 cell + extra cell).
Word and Oxi place them in different sequence due to cell traversal order.

This is NOT a Y bug per se — it's a measurement artifact. The script picks
"nearest Oxi Y" but the Word/Oxi para sequence is misaligned.

### 3.3 Final section (paras 46–57): mild drift, nearly aligned

End of page settles to drift ≈ −1pt. The bottom of p.1 is mostly
correct. The damage is concentrated in the middle table region (paras
22–42).

## 4. Comparison with ed025

| feature | ed025 | b35123 |
|---|---|---|
| drift profile | monotonic 4 jumps | oscillating ±9pt |
| cumulative final drift | +145pt | varies |
| primary bug | empty-para line height (+16pt) | per-row table cell height |
| secondary bug | table-cell paragraph stacking (+91pt) | (similar?) |
| fix complexity | 1-2 spots | many table rows |

**b35123 is harder to fix** than ed025. ed025 had 1-2 specific bugs that
cascade everywhere; b35123 has many small per-row mismatches that don't
share a single root.

## 5. Pixel-level table row pairing (UPDATED 2026-05-03)

Tool: `tools/metrics/detect_table_rows.py` (added on main, commit `199a31c`).
Detects horizontal dark lines spanning ≥20% page width = table cell borders.
Pairs Word vs Oxi by closest Y (≤30pt threshold). Authoritative for row
structure since it works on rendered pixels, not paragraph indices.

Result on b35123 p.1:

| word_y | oxi_y | diff | span_w | span_o | note |
|---:|---:|---:|---:|---:|---|
| 123.4 | 124.3 | +0.96 | 76% | 76% | top header row |
| 186.2 | 185.3 | −0.96 | 66% | 66% | |
| **248.6** | **NO MATCH** | — | 66% | — | **Oxi missing this row border** |
| 311.0 | 308.6 | −2.40 | 66% | 66% | |
| **432.0** | **438.2** | **+6.24** ⚠ | 66% | 76% | **biggest single drift** |
| 505.9 | 508.8 | +2.88 | 66% | 66% | |
| 601.4 | 599.0 | −2.40 | 76% | 76% | |
| 707.5 | 705.1 | −2.40 | 66% | 66% | |
| 769.9 | 765.6 | −4.32 | 76% | 76% | bottom |

Plus Oxi has an EXTRA border at y=550.1 with no Word match within 30pt
(span 66%, looks like a cell border Oxi inserts where Word draws none).

### Key findings (clearer than the per-paragraph analysis above)

1. **Row 3 (Word y=248.6) is missing in Oxi.** Word has 9 row borders;
   Oxi has 9 row borders too, but the structural mapping is off — one of
   Word's middle rows has no counterpart, and Oxi inserts an extra border
   at y=550 that Word doesn't have. This is a **table row count /
   structure** divergence, not just per-row height.

2. **Row 5 (y=432 in Word, 438 in Oxi) drifts +6.24pt.** Largest single
   row offset; likely the cause of the para-29-onwards +8pt drift seen
   in the per-paragraph data.

3. **Spans match closely** (66% / 76% — Word and Oxi see the same border
   thickness/extent), so the borders being detected are the same kind of
   element, just placed at different Y.

4. **Cumulative pattern: drifts oscillate** from +0.96 → −0.96 → −2.40 →
   +6.24 → +2.88 → −2.40 → −2.40 → −4.32. This confirms b35123 isn't
   just one bug accumulating — multiple small mis-alignments per row.

### Updated diagnosis priority

1. **Primary: Find why Oxi misses Word's row at y=248.6.** This is the
   only structural mismatch (count vs y-offset). Could be:
   - A `<w:tr>` with `<w:cantSplit/>` or `<w:hideMark/>` that Oxi handles
     differently
   - A row whose `<w:trHeight>` is 0 or the row has merged cells
     making the "border" invisible
   - A `<w:tcBorders>` setting Oxi misinterprets

2. **Secondary: Row 5 +6.24pt over-tall.** Why does Oxi make this
   specific row 6pt taller than Word? Inspect the cell content of row 5
   — likely contains a paragraph with line-height/spacing Oxi mis-computes.

3. **Tertiary: Oxi extra border at 550.1.** Possibly Oxi splits a Word
   merged-cell into two rows.

## §6.5 Update 2026-05-03 — vMerge row height investigation (Day 1)

**Direct measurement of Oxi row heights via debug instrumentation**
(`OXI_DBG_ROWS=1` env var, temporary print at `mod.rs:row height calc`):

```
Oxi page 1 (b35123):
  R0 (header):              y=124.50, h=13.12,  cont=0/2
  R1 (組織的管理措置 restart): y=142.50, h=25.75,  cont=0/2
  R2 continue:              y=185.62, h=51.00,  cont=1/2
  R3 continue:              y=248.12, h=51.00,  cont=1/2
  R4 continue:              y=308.75, h=38.38,  cont=1/2
  R5 (人的管理措置 restart):   y=350.00, h=22.65,  cont=0/2
  R6 continue:              y=385.50, h=36.38,  cont=1/2
  R7 (物理的管理措置 restart): y=438.50, h=51.00,  cont=0/2  ← +6.5pt drift vs Word 432
  R8 continue:              y=509.00, h=38.38,  cont=1/2
  R9 continue:              y=550.25, h=38.38,  cont=1/2
  R10 (技術的管理措置 restart):y=599.25, h=38.38,  cont=0/2
  R11 continue:             y=652.25, h=38.38,  cont=1/2
  R12 continue:             y=705.25, h=51.00,  cont=1/2
```

**Key revisions to §5 pixel analysis**:

The pixel-detection "missing row at y=248.6 / extra border at 550.1" was
**a false positive of the >20% page-width threshold**. Internal vMerge
boundaries (where the LEFT label cell continues but the RIGHT measure
cell transitions) draw a partial-width border on the right column only.
That partial line is below my pixel detection threshold, so it appears
"missing" — but it's actually drawn correctly.

The REAL bug is **content-height mismatch within vMerge groups**:

| vMerge group | Word total height | Oxi total height | drift |
|---|---:|---:|---:|
| Header | 62.8pt | 18.0pt (R0 only) | -44.8pt ← Oxi too short |
| 組織的 (4 rows: R1-R4) | 245pt | 207.5pt | -37.5pt |
| 人的 (2 rows: R5-R6) | 95pt | 89pt | -6pt |
| 物理的 (3 rows: R7-R9) | 169pt | 161pt | -8pt |
| 技術的 (3 rows: R10-R12) | 168pt | 157pt | -11pt |
| **TOTAL** | 740pt | 632pt | **-108pt** |

Wait — the cumulative drift goes the OTHER direction in pixel measurement.
Let me re-check by comparing observed Oxi y vs Word y:

| boundary | Word y | Oxi y | diff |
|---|---:|---:|---:|
| Header top | 123.4 | 124.5 | +1.1 |
| 組織的 top | 186.2 | 185.6 | -0.6 |
| 人的 top | ~311 | 350.0 | +39pt!? |
| 物理的 top | 432.0 | 438.5 | +6.5 |
| 技術的 top | 601.4 | 599.25 | -2.15 |
| Last row top | 707.5 | 705.25 | -2.25 |
| Table bottom | 769.9 | ~756.25 | -13.65 |

The 人的 transition (Word at ~311, Oxi at 350) is +39pt off — Oxi puts 人的
much later. This is the dominant drift source.

Why? Looking at Oxi data: R4 ends at 308.75 + 38.38 = 347.13. R5 (人的 restart)
starts at 350.00. So the gap from R0 (header, 124.5) to R5 (人的 350.0) =
225.5pt covers 4 measure rows of 組織的.

In Word: header=123.4, 人的~311 → 187.6pt covers same 4 measure rows. So
Word's 組織的 group is MUCH SHORTER (187pt vs 226pt). Oxi adds 39pt extra
to 組織的 group.

That ~39pt overshoot is the bug. Likely R2/R3 each at h=51pt are too tall
(Word may use ~38pt each).

**Root cause hypothesis**: `estimate_para_height` for the cell content
returns 51pt for these checkbox+text cells, but Word renders them at
~38pt. Possibly a wrap-decision (Oxi wraps to 3 lines, Word to 2) or
line-height multiplier mismatch.

R2/R3 content is 「□ 匿名データを取り扱う者の権限及び責務並びに業務を…」
— a long checkbox+text line that wraps. Cell width drives wrap count,
which drives row height.

**Next step (multi-day work)**: Diagnose why Oxi wraps these long
checkbox+text rows to MORE lines than Word. Likely related to:
- inner_w calc (cell_w - pad_l - pad_r - border)
- character width for □ + Japanese text
- Latin-CJK word break rules

## §6.7 Update 2026-05-03 — Word COM cell measurements (Day 1.5)

Word COM (`Range.Cells` iteration vMerge-safe) gives actual Word cell Y:

| label | Word y | Oxi y | drift |
|---|---:|---:|---:|
| R1 hdr | 126.00 | 124.50 | -1.5 |
| R2 組織restart | 144.00 | 142.50 | -1.5 |
| R3 cont | 189.00 | 185.62 | -3.4 |
| R4 cont | 251.00 | 248.12 | -2.9 |
| R5 cont | 313.50 | 308.75 | -4.75 |
| **R6 人的restart** | **358.50** | **350.00** | **-9.0** ← swing |
| R7 cont | 381.50 | 385.50 | **+4.0** |
| R8 物理restart | 434.50 | 438.50 | +4.0 |
| R9 cont | 508.50 | 509.00 | +0.5 |
| R10 cont | 553.50 | 550.25 | -3.25 |
| R11 技術restart | 604.00 | 599.25 | -4.75 |
| R12 cont | 657.00 | 652.25 | -4.75 |
| R13 cont | 710.00 | 705.25 | -4.75 |

R5→R6 transition: Word advance 45pt, Oxi advance 41.25pt (drift -3.75pt).
R6→R7 transition: Word advance 23pt, Oxi advance 35.5pt (drift +12.5pt).
R6 (人的 restart) is the dominant drift source.

## §6.8 Root cause — 2-pass row_height revision

`crates/oxidocs-core/src/layout/mod.rs:6189-6209` has a 2-pass row height:
1. Initial `row_height` from `estimate_para_height` per cell (line 5526)
2. After rendering, `max_actual_cell_h > row_height` triggers `row_height
   = max_actual_cell_h` (revised upward)

For R6 (人的 restart, `<w:trHeight w:val="453"/>` = 22.65pt atLeast):
- Declared row_height = 22.65pt (matches trHeight)
- Actual rendered cell content height ≈ 35.5pt
- → `row_height` revised to 35.5pt (+12.85pt vs declared)

cursor_y advances by the REVISED row_height.

Word renders the same row at exactly 23pt (matches trHeight + content).

**Why is Oxi's actual rendered height larger than Word's?**

Hypothesis: cell content `<w:p>` paragraphs render differently between
estimate and actual. Estimate uses wrap heuristic that matches Word
(content fits in 1 line). Actual rendering wraps to 2+ lines somewhere.

Likely culprits:
1. **Line height multiplier** — actual rendering uses 1.5× line spacing
   while estimate uses 1.0×, growing each line by 50%.
2. **Paragraph spacing** — actual adds `space_before/after` between
   consecutive cell paragraphs while estimate doesn't.
3. **Cell padding** — actual adds `pad_t + pad_b` separately while
   estimate already accounts for it.

To fix (multi-day):
- Add debug output for `actual_cell_h` vs `cell_content_h` (estimate)
  per cell, identify the discrepancy source.
- Reconcile estimate and actual rendering paths so they match.
- Ideal: rendering path CALLS estimate_para_height (single source).

**Predicted impact**: If R6 row drops from 35.5pt to 23pt (-12.5pt), the
post-R6 drift collapses by ~12pt. b35123 likely improves by ~+0.05 SSIM.
Same fix may help other vMerge+trHeight docs (1ec1 has similar pattern,
1636d28e too — could be a class fix benefiting multiple bottom docs).

## §6.9 Day 2 — char_width discrepancy between estimate and actual rendering

Per-cell debug instrumentation (`OXI_DBG_CELL=1` env) gives both estimate
and actual content_h:

| row | cell | text | est | act | gap |
|---|---|---|---:|---:|---:|
| R0 hdr | C0 区分 (2字) | 1 para | 13.12 | 17.50 | +4.4 |
| R0 hdr | C1 匿名... (29字) | 1 para | 13.12 | 17.50 | +4.4 |
| R1 組織restart | C0 組織的管理措置 (6字) | 1 para | 25.75 | **35.00** | +9.3 |
| R1 | C1 (2 paras) | | 25.75 | 42.62 | +16.9 |
| **R5 人的restart** | **C0 人的管理措置 (6字)** | **1 para** | **13.12** | **35.00** | **+21.9** ← 主犯 |
| R5 | C1 □ 法人... (1 para) | | 16.50 | 16.00 | -0.5 |
| R6 | C1 (3 paras) | | 36.38 | 52.50 | +16.1 |
| R7 物理restart | C0 物理的管理措置 | 1 para | 25.75 | 35.00 | +9.3 |

**Pattern**: ALL label cells (C0 with names like 組織的管理措置, 人的管理
措置, etc., 6 chars × 10.5pt) show est=13-25pt vs act=35pt — Oxi actual
rendering wraps these to 2 lines, estimate computes 1 line.

**count_cell_lines (estimate path) returns 1 line** — meaning chars FIT
in 59.85pt cell width per estimate's char-width calc.

**actual rendering wraps to 2 lines** — meaning chars OVERFLOW 59.85pt
per the actual path's char-width calc.

→ **Different `char_width_pt` values returned by estimate vs actual paths**.

For 6 chars × 10.5pt = 63pt:
- Cell width 59.85pt
- estimate: ~9.85pt/char × 6 = 59.1pt (fits)
- actual: ~10.5pt/char × 6 = 63pt (wraps)

Word side: same character "人的管理措置" renders at h=23pt (1 line) in
its 59.85pt cell — matches estimate's behavior. Actual rendering is
wrong.

## §6.10 Fix candidate

Two paths:

**A. Make actual rendering use estimate's char_width** — find the
`char_width_pt_*` function used in actual cell rendering (around
mod.rs:5993 or break_into_lines internals) and reconcile with
`char_width_pt_with_fallback` used by `count_cell_lines`.

**B. Allow overflow in actual rendering** — Word may not strictly wrap
when content barely overflows cell width. Add a small tolerance (e.g.
1pt) before wrap decision.

Option A is more correct (matches Word). Option B is a heuristic.

**Predicted impact**: ~7-12 label cells per b35123 page each save 12-22pt
of unnecessary row height. b35123 row drift collapses, SSIM likely
+0.05 to +0.10. Same fix benefits any doc with narrow vMerge label
columns (tokumei series, kyodokenkyuyoushiki series, b35 forms).

## Multi-day Day 3+ work plan

1. **Identify char_width fn used in actual cell rendering** — break_into_lines
   or its sub-call. Compare with `char_width_pt_with_fallback`.
2. **Diff the implementations** — find why one returns 9.85pt and the
   other returns 10.5pt for the same CJK char in MS Mincho theme font.
3. **Decide fix direction** — change actual to match estimate (likely),
   or estimate to match actual.
4. **Implement + canary** on b35123, 1ec1, 1636d28e, and bottom-15.
5. **Ship** if net positive.

## 7. Files / data

- Measurement script: `c:/tmp/measure_b35123_v2.py`
- Measurement output: `c:/tmp/b35123_v2.log` (raw paragraph-by-paragraph table)
- Visual diff: `c:/tmp/b35123_diff.png` (red = oxi extra, blue = word extra)
- Source docx: `tools/golden-test/documents/docx/b35123fe8efc_tokumei_08_01.docx`
- Word PNG: `pipeline_data/word_png/b35123fe8efc_tokumei_08_01/page_0001.png`
- Oxi PNG: `pipeline_data/oxi_png/b35123fe8efc_tokumei_08_01/page_0001.png`
  (DWrite, post-merge; current SSIM 0.6661)
- Layout dump: `c:/tmp/b35123_layout.json`
- Sister doc handoff: `pipeline_data/ed025_drift_investigation_2026-05-03.md`
- Pre-existing partial work: main commit `bce8715`
  (`research(b35123): per-char COM measurement scripts (Inv-X, partial)`).
  These per-char scripts (under `tools/metrics/measure_b35123_per_char*.py`)
  measure HORIZONTAL position; this investigation measures VERTICAL.
  Both axes likely contribute to the 0.6661 SSIM.
- Pixel-level row detector: `tools/metrics/detect_table_rows.py`
  (main commit `199a31c`). Hardcodes b35123 paths; Oxi PNG must be at
  `pipeline_data/oxi_png/b35123fe8efc_tokumei_08_01/page_p1.png` (note
  filename: `page_p1.png`, not `page_0001.png` — currently a copy was
  made manually as workaround).

## 7. Recommended next investigation steps

1. **Identify the source of Oxi's missing row at 248.6 and extra at 550**.
   Compare `<w:tbl>` XML structure for the rows around those Y positions.
2. **Inspect row 5's cell content**: which `<w:p>` elements are inside,
   what are their explicit/inherited line/spacing values, and how does
   Oxi compute the cell height.
3. **Once row structure aligns, re-run** `detect_table_rows.py` to
   confirm; remaining ~2pt per-row diffs are the per-cell line-height
   issue (similar to ed025 §3.1).

## 8. Why b35123 ranks #1 worst (and why it's resistant)

- Form-style document with dense table layout (matches "tokumei" series
  of forms — see also 6514f214e482, 1636d28e2c46, a1d6e4efa2e7,
  de6e32b5960b, all in bottom-10).
- Each row contains identical headers + checkbox + body text →
  even small row-height drift produces large pixel mismatch.
- The repeated-text structure means small Y mis-mapping cascades visually.
- Table cell row-height is one of the trickiest layout sub-systems to
  reproduce (cellMargin, vAlign, line-spacing-in-cell, vMerge, all
  interact).

**For aggregate baseline SSIM improvement**, the entire tokumei form-doc
class may share the same root cause. Fixing one row-height bug could
move 5+ docs out of the bottom-10 simultaneously.
