# Analysis — table_border_overhead (2026-04-09, Round 2)

Source data: `measurements.json` (17 variants, Word 2024 / Windows COM)

## Raw data — 17 variants

| variant            | rows | top  | bot  | l    | r    | ih   | table_h | row heights (pt)                            |
|--------------------|------|------|------|------|------|------|---------|---------------------------------------------|
| 1row_none          |   1  | —    | —    | —    | —    | —    | 50.5    | [25.5]                                      |
| 1row_top4_only     |   1  | 0.5  | —    | —    | —    | —    | 51.0    | [25.5]                                      |
| 1row_bot4_only     |   1  | —    | 0.5  | —    | —    | —    | 51.0    | [26.0]                                      |
| 1row_outer4        |   1  | 0.5  | 0.5  | 0.5  | 0.5  | —    | 51.5    | [26.0]                                      |
| 1row_topbot8       |   1  | 1.0  | 1.0  | —    | —    | —    | 52.5    | [26.5]                                      |
| 1row_outer8        |   1  | 1.0  | 1.0  | 1.0  | 1.0  | —    | 52.5    | [26.5]                                      |
| 1row_outer16       |   1  | 2.0* | 2.0* | 2.0* | 2.0* | —    | 54.5    | [27.5]                                      |
| 2row_outer4        |   2  | 0.5  | 0.5  | 0.5  | 0.5  | —    | 77.0    | [25.5, 26.0]                                |
| 2row_outer4_ih4    |   2  | 0.5  | 0.5  | 0.5  | 0.5  | 0.5  | 77.5    | [26.0, 26.0]                                |
| 2row_outer4_ih8    |   2  | 0.5  | 0.5  | 0.5  | 0.5  | 1.0  | 78.0    | [26.5, 26.0]                                |
| 2row_outer4_ih16   |   2  | 0.5  | 0.5  | 0.5  | 0.5  | 2.0  | 79.0    | [27.5, 26.0]                                |
| 3row_outer4_ih4    |   3  | 0.5  | 0.5  | 0.5  | 0.5  | 0.5  | 104.0   | [26.0, 26.5, 26.0]                          |
| 3row_outer4_ih16   |   3  | 0.5  | 0.5  | 0.5  | 0.5  | 2.0  | 107.0   | [27.5, 28.0, 26.0]                          |
| 3row_topbot4_ih4   |   3  | 0.5  | 0.5  | —    | —    | 0.5  | 104.0   | [26.0, 26.5, 26.0]                          |
| 5row_outer4_ih4    |   5  | 0.5  | 0.5  | 0.5  | 0.5  | 0.5  | 156.0   | [26.0, 26.5, 26.0, 26.0, 26.0]              |
| 10row_outer4_ih4   |  10  | 0.5  | 0.5  | 0.5  | 0.5  | 0.5  | 287.0   | [26.0, 26.5, 26.0, 26.0, 26.5, 26.0, 26.0, 26.5, 26.0, 26.0] |
| 20row_outer4_ih4   |  20  | 0.5  | 0.5  | 0.5  | 0.5  | 0.5  | 548.5   | (3 rows of 26.5pt sprinkled among 17 of 26.0pt) |

\*sz=16: declared OOXML 2.0pt, but COM `Border.LineWidth` reports 18 (2.25pt)
because the API snaps to wdLineWidth enum quanta. The **declared OOXML value
(2.0pt) is what layout consumes** — verified by the 1row_outer16 fit below.

## Constants (calibrated from singletons)

- `M` (marker paragraph contribution to measured table_h delta) = **25.0pt**
- `R` (bare row content height for "あ" 10.5pt MS Mincho, no padding, no borders) = **25.5pt**

Derived from `1row_none`: M + R = 50.5pt.

## Spec 1 — Border overhead

```
overhead_pt = top_border + bot_border + (num_rows - 1) * insideH
                          (left/right borders do NOT contribute)
```

Validated against all 17 variants:

| evidence                                    | conclusion |
|---------------------------------------------|------------|
| 1row_none / 1row_top4_only / 1row_bot4_only | top and bot contribute independently, each `+sz`. |
| 1row_outer4 vs 1row_topbot8                 | l/r contribute zero. (51.5−51.0=0.5 from bot border alone. l/r addition is +0pt.) |
| 1row_outer8 vs 1row_topbot8 (both 52.5)     | l/r contribute zero in larger sz too. |
| 1row_outer16 = 54.5                         | uses **declared** 2.0pt (not COM-reported 2.25). Confirms layout source. |
| 2row_outer4 vs 1row_outer4                  | row 2 adds R + bot only (+25.5). With ih=0 the inter-row gap is 0. |
| 2row_outer4_ih{4,8,16}                      | (N−1)·ih contributes linearly (Δih=0.5 → Δheight=0.5; Δih=2.0 → Δheight=2.0 vs ih4). |
| 3row_outer4_ih4 vs 3row_topbot4_ih4         | identical 104.0pt — l/r truly ignored even in multi-row+ih. |
| 3row_outer4_ih16                            | N=3, ih=2.0 → predicted 25 + 76.5 + 1 + 4 = 106.5pt. Measured 107.0. **+0.5 residual** (see Spec 2). |

The border overhead spec is **clean** for the 1-row case and **clean modulo a
sub-pt residual** for multi-row.

## Spec 2 — Multi-row content height drift (NEW hypothesis, separate phenomenon)

The +0.5pt residual appearing in N≥3 rows is **not** a border-accounting bug.
It is an accumulation artifact in the per-row content height itself.

Evidence:

- 3row_outer4_ih4 vs 3row_outer4_ih16: residual is **the same +0.5pt**
  regardless of ih width. → not driven by ih.
- 3row_outer4_ih4 vs 3row_topbot4_ih4: identical 104.0pt → not driven by l/r.
- N-scaling of residual:

  | N  | predicted (Spec 1) | measured | residual |
  |----|---------------------|----------|----------|
  |  2 | 77.5                | 77.5     | 0.0      |
  |  3 | 103.5               | 104.0    | +0.5     |
  |  5 | 155.5               | 156.0    | +0.5     |
  | 10 | 285.5               | 287.0    | +1.5     |
  | 20 | 545.5               | 548.5    | +3.0     |

  Per-row residual: 0, 0.17, 0.10, 0.15, 0.15. Asymptotically ~0.15pt/row.
- The residual lands on specific rows (rows 2, 5, 8, 11, 15, 18 in N=20).
  These are 26.5pt rows interspersed in a sea of 26.0pt rows. Total extras
  match the residual exactly: 6 × 0.5pt = 3.0pt for N=20.

**Interpretation:** the *true* per-row content height in Word's float math is
slightly more than 25.5pt — call it `R_true ≈ 25.65pt`. Each row's top Y is
snapped to a 0.5pt grid (matching the COM `Information(6)` resolution). When
the cumulative drift `(k × R_true) mod 0.5` crosses 0.5, that row's snapped
delta jumps from 26.0 to 26.5.

This is **the same quantization Word applies to grid snap and line height**
elsewhere, so it is not a new mechanism — it's the existing
`floor(font_size × 83/64 × 8) / 8` machinery (already documented in
`memory/line_height_eighth_pt.md`) interacting with table cell content.

**Status of Spec 2:** new hypothesis. Confirmation requires:

1. Reverse-engineer R_true for "あ" 10.5pt MS Mincho. Likely candidates:
   - `floor(10.5 × 83/64 × 8) / 8 = 13.5pt` × 2 lines? No (only 1 line of text)
   - `25.5 + (10.5 × something / 64)` ?
   - Direct measurement via float-precision GDI rendering (`tools/oxi-gdi-renderer/`)
2. Confirm via varied font sizes (10, 11, 12, 14pt) — does R_true scale linearly?
3. Confirm via varied content (2 lines per cell, mixed content) — separates
   line-height drift from cell-padding drift.
4. Look at Oxi's existing `eighth_pt` floor and see whether table-cell content
   already uses it (mod.rs). If yes, the formula is already there but applied
   to the wrong granularity.

This is a **separate spec from border overhead** and should be re-derived
in its own minimal repro directory (proposed: `tests/fixtures/minimal_repro/
table_row_height_drift/`).

## Real-doc cross-check (1-row case) — and a critical finding

Ran `tools/metrics/measure_3doc_table_borders.py` (Word side) +
`cargo run --release --example layout_json -- ... --structure` (Oxi side):

| doc                              | table | borders   | Word h | Oxi h | gap   | Spec 1 prediction | Oxi after Spec 1 | vs Word |
|----------------------------------|-------|-----------|--------|-------|-------|--------------------|-------------------|---------|
| 683f...contract_addon_00         |   1   | sz=4 all  | 55.00  | 54.00 | −1.00 | +1.00pt overhead   | 55.00             | **0.0**  ✓ |
| 683f...contract_addon_00         |   2   | sz=4 all  | 41.50  | 40.50 | −1.00 | +1.00pt overhead   | 41.50             | **0.0**  ✓ |
| 4a36...kyodokenkyuyoushiki10     |   1   | sz=4 all  | 609.50 | 609.00 | −0.50 | +1.00pt overhead   | 610.00            | **+0.50** ✗ |
| e201...tokumei_08_05             |   1   | sz=4 all  | 15.00  | 20.50 | **+5.50** | +1.00pt overhead | 21.50             | **+6.50** ✗ |

### Critical reading

- **683f** (×2): Spec 1 fits perfectly. The −1.0pt gap = top 0.5 + bot 0.5
  exactly. Both independent tables in the same document agree.
- **kyodoken10**: Word and Oxi differ by only −0.5pt, not −1.0pt. The cell
  contains ~25 lines of form-letter content. **Spec 2's per-line drift
  (~0.15pt × N) has cancelled half of Spec 1's overhead.** Implementing
  Spec 1 alone here would push Oxi from 609.0 → 610.0, a 0.5pt regression
  vs the current state.
- **tokumei_08_05**: Oxi computes the cell as 20.5pt vs Word's 15.0pt.
  This is a +5.5pt error in the **opposite direction** and is unrelated to
  border overhead. Probably a table-cell minimum-height bug for short cells
  (single 4-char line "申出番号"). **Out of scope for Spec 1.**

### What this means for Spec 1 promotion

We have **2 confirming real docs** (both tables in 683f), not 3. The other
two real docs are blocked by other phenomena:

- kyodoken10 → blocked by Spec 2 (multi-line content drift)
- tokumei_08_05 → blocked by an orthogonal short-cell sizing bug

**Spec 1 alone would regress kyodoken10 by 0.5pt.** Under the new
zero-regression rule, Spec 1 cannot ship without Spec 2 (or without finding
two more clean 1-row docs with short content where Spec 2 doesn't activate).

This is **exactly what the discipline is for**. Under the old net-positive
rule, the kyodoken10 0.5pt regression would have been hidden by the
683f gains. The new gate caught it before any code was touched.

## Decision

**Halt Spec 1 implementation.** Three viable next steps, in order of merit:

1. **Derive Spec 2 first**, then ship 1+2 atomically. This is the
   structurally correct path because the two specs are demonstrably
   entangled in real documents. Open
   `tests/fixtures/minimal_repro/table_row_height_drift/` and isolate
   per-line drift in single cells with known line counts.
2. Find ≥2 more real 1-row-table docs where the cell content is **short
   enough** that Spec 2 drift is < 0.25pt (so it snap-rounds to 0). Then
   Spec 1 has its 3-doc backing and can ship alone — but kyodoken10 will
   show a 0.5pt regression that we'll have to revert or re-fix.
3. Investigate the tokumei_08_05 short-cell bug as its own minimal repro
   (separate spec). Likely a faster win than Spec 2 and may unblock other
   small-cell documents.

**Anti-patterns explicitly avoided:**

- ❌ Implementing only the 1-row carve-out — already established as
  EXCEPTION stacking and not implementable.
- ❌ Implementing Spec 1 with `OXI_ALLOW_REGRESSION=1` to force the merge
  past the new gate. The gate exists exactly to prevent this.
- ❌ Treating the kyodoken10 +0.5pt as "good enough average" — that's the
  old net-positive rule the new discipline replaces.

## Spec 2 (row content drift) — what's known so far

From the 17 minimal repro variants:

- Per-row content height in Word's float math is `R_true ≈ 25.65pt` for
  "あ" 10.5pt MS Mincho with no padding/borders, even though
  `Information(6)` reports each individual row as 25.5pt or 26.0pt due to
  0.5pt Y quantization.
- Each row's snapped position = `round(k × R_true × 2) / 2`. The fractional
  drift (0.15pt/row) accumulates and snaps up roughly every 3-4 rows.
- Validated against N = 2, 3, 5, 10, 20 (residuals 0, +0.5, +0.5, +1.5, +3.0).

This drift is per **row**, but within a single cell with multi-line content
the same drift accumulates per **line**, which is what we see in kyodoken10.

Likely root cause: Oxi rounds line height to 0.5pt at each line, while Word
keeps a float and snaps at render time. The fix is probably the same as
`feedback/cumul_round_cross_para` (cumulative carry across paragraphs) but
applied within table cells too.

A separate minimal repro is required: `tests/fixtures/minimal_repro/
table_row_height_drift/` with:
- 1 cell × N lines for N = 1, 2, 5, 10, 20, 50
- Various font sizes (10, 10.5, 11, 12pt)
- Various line spacing modes (auto / multiple / exact)

## Status

**HYPOTHESIS — partial.** Border overhead formula is well-supported by 17
minimal variants and 2 real-doc tables, but **blocked from implementation**
by entanglement with Spec 2 in real-world multi-line single-cell tables.

Implementation would currently fix 683f but introduce a 0.5pt regression
on kyodoken10 — the zero-regression rule says no.
