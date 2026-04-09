# table_row_height_drift — minimal repro

## Spec under test

How does Word advance Y inside a **table cell** with multiple lines of content?

Round 2 of `table_border_overhead/` discovered that real-world multi-line cells
(e.g. kyodoken10 with ~25 form lines in one cell) accumulate ~0.15pt of
fractional Y per line, which snaps up roughly every 3-4 lines on the 0.5pt
COM `Information(6)` grid. This residual is what blocks shipping the border
overhead spec alone — implementing border overhead while leaving the per-line
drift unfixed regresses kyodoken10 by +0.5pt.

## Relation to existing cumulative-round fix

The body-paragraph version of this drift is already fixed via thread-local
`CUM_DRIFT_Y` in `crates/oxidocs-core/src/layout/mod.rs` (commit 4bdff3e,
documented in `memory/cumulative_round_cross_para.md`). That fix:

- applies to single-line body paragraphs only
- in LayoutMode=0 only
- for non-exact/atLeast rules
- carries `raw_height - rounded_height` per paragraph and snaps the cursor
  by ±0.5pt when the carry crosses ±0.25pt

Hypothesis: the same mechanism is needed inside table cells, applied
**per-line** instead of per-paragraph. Multi-line table cells currently
round each line height independently and lose the fractional pt.

## Status

**HYPOTHESIS** (2026-04-09). Derived from `table_border_overhead/` round 2
analysis, not yet measured directly in isolation.

## Variants to author

Single-cell, single-paragraph tables with N lines of identical content,
varying:

| variant                | rows | cols | lines/cell | font           | sz    | rule    |
|------------------------|------|------|-----------|-----------------|-------|---------|
| 1cell_1line_mincho10p5 |  1   |  1   |    1      | MS Mincho       | 10.5  | auto    |
| 1cell_2line_mincho10p5 |  1   |  1   |    2      | MS Mincho       | 10.5  | auto    |
| 1cell_3line_mincho10p5 |  1   |  1   |    3      | MS Mincho       | 10.5  | auto    |
| 1cell_5line_mincho10p5 |  1   |  1   |    5      | MS Mincho       | 10.5  | auto    |
| 1cell_10line_mincho10p5|  1   |  1   |   10      | MS Mincho       | 10.5  | auto    |
| 1cell_20line_mincho10p5|  1   |  1   |   20      | MS Mincho       | 10.5  | auto    |
| 1cell_50line_mincho10p5|  1   |  1   |   50      | MS Mincho       | 10.5  | auto    |
| 1cell_10line_mincho10  |  1   |  1   |   10      | MS Mincho       | 10    | auto    |
| 1cell_10line_mincho11  |  1   |  1   |   10      | MS Mincho       | 11    | auto    |
| 1cell_10line_mincho12  |  1   |  1   |   10      | MS Mincho       | 12    | auto    |
| 1cell_10line_calibri11 |  1   |  1   |   10      | Calibri         | 11    | auto    |
| 1cell_10line_yumin10p5 |  1   |  1   |   10      | Yu Mincho       | 10.5  | auto    |

Each cell holds N copies of `"あ\n"` (forced line breaks within one paragraph
via Shift+Enter, OOXML `<w:br/>`). No padding overrides, no borders, no
inter-row interaction.

## Measurements wanted

For each variant:

- Y of paragraph above table (anchor), Y of paragraph below table
- Y of each line within the cell (via `Range.Sentences` or `.Words`/Information(6))
- `cell.Range.Information(6)` for the first character of each line
- Cell content height = `Y(after-table) − Y(cell first line)`
- Per-line delta (this gives the snapped advance)
- Sum of per-line deltas vs total content height (drift detector)

## Goal

Reverse-engineer:

```
line_y[k] = Y0 + round_to_0.5pt(k * R_true(font, sz, rule))
```

Specifically, find `R_true` for each (font, sz) tuple. Compare against:

- Oxi's current rounded line height (`floor(sz × 83/64 × 8) / 8` for CJK)
- The body-para `CUM_DRIFT_Y` formula

If the table-cell drift matches the body-para mechanism exactly, the fix is
to extend the existing thread-local carry to apply per-line inside cells
(not per-paragraph). Else there's a separate cell-specific quantization.

## Files

- `generate.py` — author N-line cell variants
- `measure.py` — extract line Y positions via Word COM
- `measurements.json` — output
- `analysis.md` — derivation (after measurements come in)

## Anti-patterns

- ❌ "Multi-line cell adds 0.15×N overhead" carve-out — that hides the
  per-line snap mechanism
- ❌ Implementing the fix in cell layout without first checking it's
  the same mechanism as cross-para drift (otherwise we add a parallel
  code path that will rot)
- ❌ Confirming Spec 2 from one font alone — we already know mixed-font
  drift went the wrong way once (`mixed_font_line_height` regression in
  the cross-para fix)
