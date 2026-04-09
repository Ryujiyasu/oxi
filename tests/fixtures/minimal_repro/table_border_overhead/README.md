# table_border_overhead — minimal repro

## Spec under test

How does Word add border thickness to **table row height**?

Current Oxi implementation ([crates/oxidocs-core/src/layout/mod.rs:3054-3061](../../../../crates/oxidocs-core/src/layout/mod.rs#L3054-L3061)):

```rust
// COM-confirmed (2026-04-02): only insideH border adds row height overhead.
// Outer borders (top/bottom/left/right) do NOT affect row height.
let border_overhead = if table.style.has_inside_h {
    table.style.border_width.unwrap_or(...)
} else {
    0.0
};
```

This rule was confirmed against **a single document** in 2026-04-02 and is now contradicted by 683f data:

- 683f Table 1 (1 row, no real insideH rendering): Word 55.0pt, Oxi 54.0pt → −1.0pt
- 683f Table 2 (1 row, no real insideH rendering): Word 41.5pt, Oxi 40.5pt → −1.0pt
- diff = exactly 0.5pt + 0.5pt = top border + bottom border (sz=4 → 0.5pt each)

The "insideH only" rule is therefore at best incomplete, possibly wrong. Per the
no-EXCEPTION-stacking rule (`feedback_ra_loop_tightening`), we re-derive instead
of carving out a 1-row special case.

## Status

**HYPOTHESIS** (2026-04-09). The 2026-04-02 confirmation has been demoted.
Re-derivation in progress.

## Variants to author

To reverse-engineer `row_height = f(content, padding, top_border, bot_border, insideH, num_rows, sz)`,
generate `.docx` files spanning these dimensions:

| variant         | rows | top  | bot  | left | right | insideH | sz   |
|-----------------|------|------|------|------|-------|---------|------|
| 1row_none       |   1  | none | none | none | none  | none    | —    |
| 1row_outer4     |   1  | sng4 | sng4 | sng4 | sng4  | none    | 4    |
| 1row_outer8     |   1  | sng8 | sng8 | sng8 | sng8  | none    | 8    |
| 1row_outer16    |   1  | sng16| sng16| sng16| sng16 | none    | 16   |
| 2row_outer4     |   2  | sng4 | sng4 | sng4 | sng4  | none    | 4    |
| 2row_outer4_ih4 |   2  | sng4 | sng4 | sng4 | sng4  | sng4    | 4    |
| 3row_outer4_ih4 |   3  | sng4 | sng4 | sng4 | sng4  | sng4    | 4    |
| 1row_top4_only  |   1  | sng4 | none | none | none  | none    | 4    |
| 1row_bot4_only  |   1  | none | sng4 | none | none  | none    | 4    |
| 1row_topbot8    |   1  | sng8 | sng8 | none | none  | none    | 8    |

Each cell holds **identical** content (`"あ"` 10.5pt, MS Mincho, no padding override),
so the only varying input is the border configuration. This isolates the spec.

## Real-document cross-check (need ≥3)

Candidates already known to involve table borders:

- `683ffcab86e2*.docx` (the doc that broke the old spec) — 1-row outer-only tables
- `tokumei_08_*.docx` — multi-row, has insideH
- `LOD_Handbook.docx` — multi-row tables
- (TBD) — find a 1-row table with NO borders for the null case

The minimal repro must match COM-measured row heights from **all** of these
before the spec is promoted to confirmed.

## Files

- `generate.py` — author the variant `.docx` files via python-docx
- `measure.py` — for each variant, open in Word COM and dump the row Y
  coordinates and computed heights to `measurements.json`
- `measurements.json` — output of measure.py (gitignored if large; small here)
- `analysis.md` — summary of the derived formula once measurements are in

## Next actions

1. Run `python generate.py` to produce variants (Windows + python-docx)
2. Run `python measure.py` (Windows + Word + pywin32)
3. Inspect `measurements.json`, derive a closed-form `border_overhead(top, bot, ih, num_rows, sz)`
4. Cross-check the formula against the 3 real docs above
5. If everything agrees: implement, run `pipeline.verify` (zero-regression gate),
   then promote this spec to confirmed in `manifest.json` and the relevant memory
6. If anything disagrees: stay in hypothesis, expand variants, do not implement

## Anti-patterns to avoid

- ❌ "1-row tables get +1pt overhead" carve-out — that is EXCEPTION stacking
- ❌ Confirming after measuring only 683f again — need 3 distinct sources
- ❌ Dropping the existing `has_inside_h` field as "obviously unused" before
  the new formula is proven — it may still be a component of the right answer
