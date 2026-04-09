# Analysis — table_border_overhead (2026-04-09)

Source data: `measurements.json` (10 variants, Word 2024 / Windows COM)

## Raw observations

| variant            | rows | top  | bot  | left | right | ih   | table_h | row_h list      |
|--------------------|------|------|------|------|-------|------|---------|------------------|
| 1row_none          |  1   | —    | —    | —    | —     | —    | 50.5    | [25.5]           |
| 1row_top4_only     |  1   | 0.5  | —    | —    | —     | —    | 51.0    | [25.5]           |
| 1row_bot4_only     |  1   | —    | 0.5  | —    | —     | —    | 51.0    | [26.0]           |
| 1row_outer4        |  1   | 0.5  | 0.5  | 0.5  | 0.5   | —    | 51.5    | [26.0]           |
| 1row_topbot8       |  1   | 1.0  | 1.0  | —    | —     | —    | 52.5    | [26.5]           |
| 1row_outer8        |  1   | 1.0  | 1.0  | 1.0  | 1.0   | —    | 52.5    | [26.5]           |
| 1row_outer16       |  1   | 2.0* | 2.0* | 2.0* | 2.0*  | —    | 54.5    | [27.5]           |
| 2row_outer4        |  2   | 0.5  | 0.5  | 0.5  | 0.5   | —    | 77.0    | [25.5, 26.0]     |
| 2row_outer4_ih4    |  2   | 0.5  | 0.5  | 0.5  | 0.5   | 0.5  | 77.5    | [26.0, 26.0]     |
| 3row_outer4_ih4    |  3   | 0.5  | 0.5  | 0.5  | 0.5   | 0.5  | 104.0   | [26.0, 26.5, 26.0] |

*sz=16: OOXML `w:sz="16"` declares 2.0pt, but `Border.LineWidth` reports 18 (2.25pt)
because Word's COM API snaps to wdLineWidth enum quanta (2, 4, 6, 8, 12, 18, 24).
The **actual layout consumption is 2.0pt** (the declared value), not 2.25pt — see
1row_outer16 row delta vs 1row_none.

## Constants derived

Letting `M` = marker-paragraph height above the table, `R` = bare row content height:

- `1row_none`: `M + R = 50.5`
- `2row_outer4`: `M + 2R + top + bot = M + 2R + 1.0 = 77.0` → `M + 2R = 76.0` → **`R = 25.5`, `M = 25.0`**
- Cross-check `1row_outer4`: `25 + 25.5 + 0.5 + 0.5 = 51.5` ✓
- Cross-check `1row_outer16`: `25 + 25.5 + 2.0 + 2.0 = 54.5` ✓

## Closed-form hypothesis

Border contribution to **total table height**:

```
overhead_pt = top_border_pt
            + bot_border_pt
            + (num_rows - 1) * insideH_pt          (only when has_inside_h && num_rows > 1)
```

Left/right borders do **not** affect row height (only column width). Verified by
the fact that `1row_topbot8` (no l/r) and `1row_outer8` (with l/r) have identical
table_height = 52.5pt.

### Per-row decomposition (where it lands)

- **Above row 1**: `top_border_pt` extra space (the cell top Y shifts down)
- **Below row N**: `bot_border_pt` extra space (the after-Y shifts down)
- **Between row i and row i+1**: `insideH_pt` extra space

Mapped to the script's `row_height_pt` metric (= `next_top_y − this_top_y`):

- For row i (i < N): `row_height = R + (insideH or 0)`
- For row N (last): `row_height = R + bot_border`
- The `top_border` is **not** part of any `row_height`; it's above row 1's cell top.

## Validation against the 10 variants

| variant            | predicted overhead | predicted table_h | measured | Δ    |
|--------------------|--------------------|--------------------|----------|------|
| 1row_none          | 0                  | 50.5               | 50.5     | 0    |
| 1row_top4_only     | 0.5                | 51.0               | 51.0     | 0    |
| 1row_bot4_only     | 0.5                | 51.0               | 51.0     | 0    |
| 1row_outer4        | 1.0                | 51.5               | 51.5     | 0    |
| 1row_topbot8       | 2.0                | 52.5               | 52.5     | 0    |
| 1row_outer8        | 2.0                | 52.5               | 52.5     | 0    |
| 1row_outer16       | 4.0                | 54.5               | 54.5     | 0    |
| 2row_outer4        | 1.0                | 77.0               | 77.0     | 0    |
| 2row_outer4_ih4    | 1.5                | 77.0               | 77.5     | **+0.5** |
| 3row_outer4_ih4    | 2.0                | 103.5              | 104.0    | **+0.5** |

**9/10 exact. 2 variants off by exactly +0.5pt.**

## Unexplained 0.5pt residual on multi-row insideH cases

Both `2row_outer4_ih4` and `3row_outer4_ih4` measure 0.5pt taller than the
formula predicts. Possible causes (NOT yet disambiguated):

1. **Word Y quantization**: COM `Information(6)` returns Y in 0.5pt steps
   (HIMETRIC / 360-th of an inch internally? unclear). One row in the chain
   could be a 0.25pt sub-pixel that snaps up. The 3-row case isolates this
   to row 2 (`row_height_pt = 26.5` vs neighbours 26.0).
2. **Off-by-one insideH counting**: maybe insideH is counted `N` times not
   `N−1` for tables with insideH, with the extra applied somewhere
   (e.g., adding to row N's bottom). But this contradicts the per-row data:
   row N = 26.0 in both 2row and 3row variants — it does NOT carry an extra IH.
3. **Border collapse / overlap rule**: when bot border meets insideH meets
   outer-bot, maybe Word adds an extra padding to disambiguate.
4. **Hidden cell-margin default that activates only with insideH**: e.g.,
   adjustLineHeightInTable interaction, even though we forced tcMar=0.

The 1-row formula is **clean** and fully explains the 683f regression
(`overhead = top + bot = 1.0pt`, exactly the −1.0pt Oxi gap). Implementing
just the 1-row part would fix 683f without touching the multi-row mystery.

But per the no-EXCEPTION-stacking rule, we must NOT promote a partial spec to
"confirmed". The hypothesis stays hypothesis until the multi-row residual is
resolved.

## Next measurement variants needed

To disambiguate the 0.5pt residual:

1. **5-row, 10-row, 20-row × outer4 + ih4**: if residual is constant 0.5pt
   regardless of N, it's a per-table fixed extra. If it grows with N, it's
   per-row. If it stays at exactly 0.5 in 3, 5, 10-row, the source is fixed.
2. **Vary insideH width**: 2row_ih8, 2row_ih16. Does residual scale with IH?
3. **Vary content height**: bigger fontSize → does residual change?
4. **Disable left/right borders entirely on multi-row**: 2row_topbot4_ih4.
   Rules out a left/right-driven artifact.
5. **3row with NO outer top/bot, only insideH**: isolates insideH from outer
   border interaction.
6. **Look at the underlying XML in the saved variants**: confirm python-docx
   wrote `tcMar=0` correctly and Word didn't override.

## Cross-check against real docs (still required for confirmation)

Even when the multi-row formula is solid, the spec needs ≥3 real-doc
agreements before promotion. Candidates:

- ✅ `683ffcab86e2_*.docx` — 1-row formula predicts +1.0pt overhead, matching
  the observed −1.0pt Oxi gap (54.0 → 55.0pt expected). **1 of 3.**
- ⏳ Need a multi-row real doc with insideH (e.g., `tokumei_08_*`)
- ⏳ Need a multi-row real doc WITHOUT insideH

## Status

**HYPOTHESIS — partial.** 1-row case fully explained; multi-row insideH case
has a 0.5pt residual of unknown source. **Do NOT implement** until the
residual is resolved AND ≥3 real docs agree.

Implementation would currently fix 683f but risk regressing multi-row
tables — exactly what the zero-regression gate exists to catch.
