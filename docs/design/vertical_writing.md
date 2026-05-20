# Vertical Writing Implementation Design

**Session 130, 2026-05-20** — Initial design document.
**Status**: Design + COM measurement complete. Implementation pending.

## Motivation

Layout code does not consult `text_direction` (parsed and stored in IR
but never read). Cells with `<w:textDirection w:val="tbRlV"/>` are
rendered as horizontal text, causing cascading layout failures:

- **2ea81a** (SSIM p.1=0.787, p.2=0.737): vertical-label form rows
  render as vertically-stacked horizontal lines instead of horizontal
  cells with rotated labels.
- **459f05f** (SSIM 0.81-0.85): same form-row pattern, repeats across
  contact-info fields.
- **7ead52** (SSIM 0.919, 1 page): one row affected.
- **ed025c** (Phase 1 FAIL, 1 paragraph delta=-1): partial — vertical
  writing is one of several drift sources.

Total impact: 4 docs out of 267 (1.5%), of which 1 is a Phase 1 failure.
Expected: 2ea81a + 459f05f confirmed clean fixes (5 pages of SSIM gain
from ~0.79 to baseline mean ~0.89), ed025c Phase 1 recovery possible.

CLAUDE.md prioritizes vertical writing at #4 ("basics only"). This
design addresses the minimum needed to correctly render tbRlV vertical
text in table cells, not full vertical-layout document support.

## OOXML Spec Background

`<w:textDirection val="..."/>` inside `<w:tcPr>` defines text direction
for cell content:

| Value | Direction | Reading order |
|---|---|---|
| `lrTb`  | Left-to-right, top-to-bottom (default) | Western/CJK horizontal |
| `tbRl`  | Top-to-bottom, right-to-left | CJK vertical, characters upright |
| `tbRlV` | Top-to-bottom, right-to-left, **letters rotated 90° CW** | CJK vertical with Western chars rotated to fit vertical flow |
| `btLr`  | Bottom-to-top, left-to-right | Rotated headers (rare) |
| `lrTbV` | Left-to-right, top-to-bottom, letters rotated | Rotated headers (rare) |

**This design covers `tbRlV` only.** All 4 affected docs use `tbRlV`
(none use `tbRl`, `btLr`, `lrTbV` per OOXML scan of 267-doc baseline).

## COM Measurements

`tools/metrics/com_measure_vert_writing.py` measured 5 minimal repros
+ 4 real docs. Key data points (Word as ground truth):

### Cell geometry (V1 repro: 1 row, 2 cells, vert text 6 chars × 8pt, vAlign=center)

- Vert cell: `x=83.25, y=88.50, w=35.45pt` (`w` = 709dxa content +
  ~10.8pt padding)
- Adj cell: `x=107.25, y=88.50, w=432.35pt`
- **Both cells share `y_top=88.50`** ← parallel rows, NOT stacked.
- Cell height reported as `9999999` (Word's "Auto" sentinel) when
  trHeight is absent. Actual rendered height = max content height.

### Cell geometry (2ea81a real doc, tbl=1 row=8 — the failing row)

- Vert cell: `x=75.00, y=500.25, w=35.45, h=113.15pt`
- Adj cell: `x=97.50, y=478.50, w=432.35pt`
- **Vert cell has 2 paragraphs in vertical flow**: 予納する理由 (6
  chars) + （いずれかを選択） (8 chars) = 14 chars × 8pt = 112pt
  natural vertical extent.
- Cell height 113.15pt ≈ 14 chars × 8pt + small inter-para gap.
- vAlign=center: vert text START y = 500.25 = cell_top (478.50) + 21.75pt
  offset (centered within 113.15pt cell since text fits exactly).
- Adj cell paragraphs at y = 478.50, 482.25, 508.50, 521.25, 547.50
  (horizontal paragraphs flowing top-down within cell), starting at
  row top.

### Effect of vAlign on vertical text

V1 (vAlign=center): cell_x = 83.25 (text centered along x-axis of cell)
V4 (no vAlign), V5 (vAlign=top): cell_x = 94.50 (text at right edge of
cell, since tbRl reads right-to-left)

→ In `tbRlV`, vAlign maps to the **perpendicular (x) axis**:
  - `top` → text anchored at right edge of cell (writing start)
  - `center` → text centered horizontally in cell
  - `bottom` → text anchored at left edge

The natural reading start is the RIGHT side of the cell (because tbRl
= right-to-left for new lines).

## Algorithm

### Phase A — Cell dimensions

For a cell `tc` with `text_direction == "tbRlV"`:

```
def vert_cell_writing_length(tc) -> f32:
    """Natural extent along vertical (writing) direction.

    Sum over each paragraph p:
        sum_chars(p) * font_size(p) + spacing_after(p)
    Then add inter-paragraph spacing.
    """

def vert_cell_writing_breadth(tc) -> f32:
    """Extent perpendicular to writing direction (= cell width along x).

    Defaults to max(font_size(p) for p in tc.paragraphs) for single-line
    vert text. Multi-line vert text (line wrap) adds more breadth.
    For tbRlV: subsequent lines flow LEFT (right-to-left).
    """

def cell_height(tc) -> f32:
    if tc.text_direction in (tbRlV, tbRl):
        # Writing direction is vertical → writing length determines height
        return vert_cell_writing_length(tc)
    else:
        # Default: horizontal text, sum of line heights
        return sum_line_heights(tc)

def row_height(tr) -> f32:
    return max(cell_height(tc) for tc in tr.cells)
```

**Key**: For tbRlV cells, cell HEIGHT (page y-extent) = sum of character
heights along writing direction, NOT a function of cell WIDTH (page x-
extent). Conversely, cell WIDTH is given by `tcW` (fixed grid column).

This is the OPPOSITE of horizontal text where cell width is the
"writing length" axis and content wraps within it.

### Phase B — Cell layout (row top y assignment)

Same as horizontal: row top y is determined by previous row's bottom +
border. All cells in a row share the same y_top regardless of text
direction. The fix is only that cell HEIGHT calculation considers vert
writing.

### Phase C — Text positioning within vert cell

```
def position_vert_text(tc, cell_top_y, cell_left_x):
    # vAlign maps to PERPENDICULAR axis for vertical writing
    valign = tc.v_align or "top"

    # Writing direction (y-axis) start: top of cell + vAlign-perpendicular padding
    # Actually for tbRlV, "writing direction" IS y-axis (top-down),
    # so vAlign affects x-axis position (perpendicular)

    writing_length = vert_cell_writing_length(tc)

    # y-positioning: text starts at cell_top_y (writing starts at top)
    # If text shorter than cell height, where does it go?
    # In Word: text starts at cell_top regardless. cell_h is FORCED by
    # vert text length (Phase A). So text fills cell exactly.
    text_y_start = cell_top_y

    # x-positioning: vAlign-controlled
    if valign == "top":
        # Text anchored at right edge of cell (tbRl reading start)
        text_x_anchor = cell_left_x + cell_width
    elif valign == "center":
        # Centered along cell width
        text_x_anchor = cell_left_x + cell_width / 2
    elif valign == "bottom":
        # Anchored at left edge (tbRl reading end)
        text_x_anchor = cell_left_x

    # For each paragraph, place chars vertically starting at text_y_start
    # going down, with chars rotated 90° CW for tbRlV variant
```

### Phase D — Rendering

For tbRlV, each glyph is rotated 90° CW around its baseline anchor.
The rotated glyph's:
- top edge becomes the right edge
- right edge becomes the bottom edge
- bottom edge becomes the left edge
- left edge becomes the top edge

#### GDI (oxi-gdi-renderer)
Use `lfEscapement` on the `LOGFONT`: set to `-900` (tenths of a degree
CW) for tbRlV. This rotates the entire text run. Then `TextOutW` at
the rotated baseline position.

Alternative: pre-rotate via SetWorldTransform (GDI affine matrix).

#### DirectWrite (oxi-dwrite-renderer)
Two options:
1. **Glyph rotation via transform**: use IDWriteTextRenderer with a
   `DWRITE_MATRIX` rotation applied per glyph run.
2. **Vertical layout**: IDWriteTextFormat supports
   `SetReadingDirection(DWRITE_READING_DIRECTION_RIGHT_TO_LEFT)` and
   `SetFlowDirection(DWRITE_FLOW_DIRECTION_TOP_TO_BOTTOM)` but this
   does NOT rotate Latin glyphs — that needs the transform approach.

Recommend (1) for tbRlV (Latin chars also rotated).

## Implementation Plan (4-6 sessions)

### Session 130 (this session) — Design
- [x] Build minimal repros (V1-V5)
- [x] COM measure on repros + 4 real docs
- [x] Write design doc (this file)
- [x] Save memory

### Session 131 — Layout (cell height + cell positioning)
- Modify `crates/oxidocs-core/src/layout/` to:
  - Detect `cell.text_direction == "tbRlV"` (or `"tbRl"`)
  - Compute cell height via `vert_cell_writing_length`
  - Position adjacent horizontal cells at row top (same y) without
    being affected by vert cell's narrow width
  - Skip horizontal line wrap for vert cells (text flows vertically,
    not wrapped on width)
- Add `OXI_VERT_WRITING=1` env var gate (research mode, default OFF)
- Unit tests against COM measurements
- Verify on V1-V5 repros: dump-layout should show vert text Y positions
  matching Word

### Session 132 — Renderer (GDI rotation)
- Modify `tools/oxi-gdi-renderer/` to handle vert text:
  - Detect text element flagged as vertical
  - Set `lfEscapement = -900` in `LOGFONT`
  - Adjust origin coords so rotated baseline lands at correct page x/y
- Rebuild GDI renderer
- Visual verify on V1 repro (open Word + render Oxi side-by-side)

### Session 133 — DWrite + verify
- Modify `tools/oxi-dwrite-renderer/` similarly (glyph transform matrix)
- Rebuild DWrite renderer (DEFAULT for pipeline.verify since S50)
- Run full pipeline.verify with `OXI_VERT_WRITING=1`
  - Expect: 2ea81a + 459f05f improve SSIM
  - Expect: ed025c partial improvement (Phase 1 fail may resolve)
- If Phase 1 doesn't regress and SSIM improves, flip default ON
- Refresh ssim_baseline.json (per S123 lesson)

### Session 134 (optional) — Edge cases
- `vAlign=top` / `vAlign=bottom` mapping (default = top)
- Multi-paragraph vert cell (2ea81a real case)
- Vert text overflow (text length > cell height — does Word truncate
  or extend cell? COM showed cell h auto-extends.)
- `tbRl` variant (chars NOT rotated, only flow direction)

## Risk Register

| Risk | Mitigation |
|---|---|
| Vert cell width-change cascades: changing cell height for vert cells could affect row heights of adj horizontal cells → ripple to pagination | Phase 1 sentinel + env var gate. Refactor S2 (S131) checks Phase 1 OFF and ON. |
| GDI/DWrite escapement coordinate conventions differ | Test on V1 repro first; visual side-by-side with Word screenshot |
| Multi-paragraph in vert cell: paragraph spacing direction | COM measurement (2ea81a) shows paragraphs flow top-down within vert cell. Inter-para spacing measured at ~1pt (113.15 - 112 = 1.15pt for 2 paras) |
| ed025c partial fix: other drift sources may regress when vert writing fixed | Run pagination_diff after S133; if ed025c worsens, isolate fix to docs that purely benefit |
| Other docs not in our 4-doc list might use textDirection in subtle ways | OOXML scan confirmed only 4 docs in 267-baseline use vertical textDirection (1.5%). Cross-impact ≤ 4 docs. |

## Success Criteria

1. **Phase 1 PASS rate ≥ 53/55** (no pagination regression) with env
   var OFF.
2. **Phase 1 PASS rate ≥ 53/55** with env var ON (gate works).
3. **SSIM improves on 2ea81a, 459f05f, 7ead52** vs baseline with env
   var ON (≥ +0.05 mean across these 3 docs' pages).
4. **No SSIM regression > 0.005** on other 263 docs.
5. **ed025c**: either Phase 1 recovers (53→54/55) OR explicit ack that
   it has additional non-vert-writing drift sources.

If 1-4 pass, flip default ON. If only 1-2-3 pass but 4 has regression,
keep env var gate.

## References

- COM measurements: `pipeline_data/vert_writing_measurements_S130.json`
- Minimal repros: `tools/golden-test/repros/vert_writing_S130/`
- Drift signature (2ea81a): `[[session129-2ea81a-vertical-writing-trigger]]`
- IR field: `crates/oxidocs-core/src/ir/types.rs:438`
  (`CellProperties.text_direction`)
- Parser: `crates/oxidocs-core/src/parser/ooxml.rs:5300-5306`
- OOXML spec: ECMA-376 Part 1 §17.4.66 textDirection
