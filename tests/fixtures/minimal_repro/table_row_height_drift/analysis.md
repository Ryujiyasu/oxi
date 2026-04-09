# Analysis — table_row_height_drift (2026-04-09)

Source data: `measurements.json` (12 variants v2 with separate `<w:p>` per
"line", Word 2024 / Windows COM)

## Two findings, in order of discovery

### Finding A: `<w:br/>` soft line breaks are not rendered by Oxi (v1)

The first attempt at this minimal repro used `<w:br/>` soft line breaks
inside a single paragraph (one paragraph per cell, N "あ" runs separated by
`<w:br/>`). Oxi laid out **all N characters horizontally on the same Y
position** instead of breaking lines. The 3-line variant rendered as:

```
TEXT  90.000   96.500  ...  T あ
TEXT 105.750   96.500  ...  T あ
TEXT 121.500   96.500  ...  T あ
```

Same Y, three different X. So `<w:br/>` is a **completely separate parser
gap**: `crates/oxidocs-core/src/parser/ooxml.rs` only handles `<w:br
type="page"/>` (as `\x0C` for inline page break), not the bare `<w:br/>`
soft line break that ECMA-376 §17.3.3.1 defines.

This is documented as a separate issue in
`memory/parser_gap_w_br_soft_break.md`. It is **not** the kyodoken10 bug,
just an unrelated find-while-fishing.

### Finding B: this minimal repro tests a different scenario than kyodoken10

After pivoting v2 to use separate `<w:p>` paragraphs (matching kyodoken10's
structure), measurements look clean and show drift behaviour very similar
to what the body-paragraph cumulative-round fix already handles:

| variant                | N  | th     | span   | per-line deltas (grouped)                 | avg     |
|------------------------|----|--------|--------|-------------------------------------------|---------|
| 1cell_2line_mincho10p5 |  2 |  76.0  |  25.5  | 25.5×1                                    | 25.500  |
| 1cell_3line_mincho10p5 |  3 | 102.0  |  51.5  | 25.5×1 26.0×1                             | 25.750  |
| 1cell_5line_mincho10p5 |  5 | 153.0  | 102.5  | 25.5×1 26.0×1 25.5×2                      | 25.625  |
| 1cell_10line_mincho10p5| 10 | 281.5  | 231.0  | 25.5×1 26.0×1 25.5×2 26.0×1 25.5×2 ... 25.5×1 | 25.667 |
| 1cell_20line_mincho10p5| 20 | 538.0  | 487.5  | (alternation)                             | 25.658  |

The alternation between 25.5pt and 26.0pt deltas is exactly the
"snap-to-0.5pt cumulative drift" pattern from `cumul_round_cross_para`.
For N=20, span/19 = 25.658pt, which corresponds to a float `R_true`.

Decomposing 25.66pt:

- 10.5pt × 83/64 = 13.617pt  (CJK natural single-line height)
- × 1.15 (`line=276` from docDefaults `pPrDefault/spacing`) = 15.659pt
- + 10pt (`after=200` from same docDefaults) = **25.659pt** ✓

So the per-paragraph advance Word produces here is **{line height ×
spacing multiplier} + SpaceAfter**, accumulated as float and snapped per
paragraph to 0.5pt.

### What Oxi produces for the same variants

`cargo run --example layout_json -- ... --structure`:

```
1cell_2line_mincho10p5 :  LINE y=96.50  LINE y=110.00                 (Δ=13.50)
1cell_5line_mincho10p5 :  Δ=13.50 ×4
1cell_10line_mincho10p5:  Δ=13.50 ×9
1cell_20line_mincho10p5:  Δ=13.50 ×19
```

Oxi's per-paragraph delta is a **constant 13.5pt** (= floor(10.5×83/64×8)/8).
Word's is ~25.66pt with drift. The gap per paragraph is ~12.16pt, which
breaks down as:

- ~2.16pt missing line multiplier (`line=276` not applied → bare 13.5
  instead of 13.5 × 1.15 = 15.525 ≈ 15.66 with eighth-pt floor)
- ~10pt missing SpaceAfter (`after=200` not applied)

In other words, **inside our generated table cells, Oxi is dropping both
the line multiplier and SpaceAfter from `pPrDefault` inheritance**.

For an 20-paragraph cell, that's 20 × 12.16 = 243pt of missing height.
**This is not the kyodoken10 bug.** kyodoken10's gap was −0.5pt, not −243pt.

### Why our minimal repro is not equivalent to kyodoken10

Inspecting `kyodokenkyuyoushiki10.docx`:

```xml
<!-- styles.xml -->
<w:docDefaults>
  <w:rPrDefault><w:rPr>... (rFonts only) ...</w:rPr></w:rPrDefault>
  <w:pPrDefault/>     <!-- EMPTY: no spacing line=, no after= -->
</w:docDefaults>

<!-- document.xml: cell paragraphs (28 of them in 1 cell) -->
<w:p>
  <w:pPr>
    <w:snapToGrid w:val="0"/>      <!-- KEY: disables grid snap -->
    <w:rPr><w:color w:val="000000"/></w:rPr>
  </w:pPr>
  ...
</w:p>
```

Two structural differences from our minimal repro:

1. **kyodoken10's pPrDefault is empty.** No line multiplier, no SpaceAfter.
   python-docx auto-generates `line=276 after=200` for any new doc, so our
   minimal repro inherits values that kyodoken10 explicitly does not have.
2. **Each kyodoken10 cell paragraph has `snapToGrid=false`.** This takes
   the layout out of the docGrid pitch entirely and uses natural font
   height (per `memory/line_height_research`). Our minimal repro has the
   default `snapToGrid=true`.

These differences put kyodoken10 on a **completely different code path**
in Oxi. The −0.5pt gap there cannot be reproduced or characterised by
the variants we generated here.

## Status

**HYPOTHESIS REFUTED — repro irrelevant.** The minimal repro as built
(v1 with `<w:br/>` or v2 with separate `<w:p>` and python-docx defaults)
does not isolate the same drift that affects kyodoken10. Two findings
remain valuable as side effects:

1. **Bug: `<w:br/>` soft line break not parsed.** Separate parser gap.
   Should get its own minimal repro and fix.
2. **Bug: docDefaults `pPr/spacing/line` and `pPr/spacing/after` not
   inherited into table cell paragraphs (in our generated docs).** Caught
   this before any wrong implementation. Worth verifying whether this is
   a real Oxi bug or an artifact of how python-docx writes docDefaults.

Neither is the same phenomenon as the kyodoken10 −0.5pt gap.

## Next steps

To actually nail Spec 2 (the kyodoken10 entanglement):

1. **Re-author the minimal repro to match kyodoken10's structure exactly:**
   - Empty `pPrDefault`
   - `snapToGrid="0"` on every cell paragraph
   - 25-30 paragraphs of "あ" in a single cell with sz=4 borders
   - Should reproduce a sub-pixel gap, ideally exactly −0.5pt at N=25
2. Once reproduced, instrument Oxi's code path for `snapToGrid=false`
   table cell paragraphs and find where the 0.5pt accumulates.
3. **Or:** redirect entirely. The kyodoken10 gap is small (−0.5pt over a
   609.5pt table), and the user-visible improvement is marginal compared
   to the 683f case (+1.0pt over 55pt). One option is to ship Spec 1
   (border overhead) anyway and explicitly accept the 0.5pt regression on
   kyodoken10 — but **only** with explicit user approval, since this
   violates the new zero-regression rule.

## Side findings to log separately

- `parser_gap_w_br_soft_break.md` — `<w:br/>` not handled
- `unverified_pPrDefault_inheritance_in_cells.md` — docDefaults sa/line
  not flowing into table cell paras in our repro (verify this is a real
  bug, not a python-docx XML quirk)
