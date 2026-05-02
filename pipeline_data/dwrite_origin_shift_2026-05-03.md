# DWrite glyph-top alignment fix — origin.y shift -1.0pt

**Date**: 2026-05-03
**Branch**: session50-visual
**File**: `tools/oxi-dwrite-renderer/src/main.rs` (`render_text` origin computation)
**Status**: SHIPPING — peak NET +2.2198 SSIM gain across 177-doc p.1 baseline.

---

## Summary

DirectWrite renderer was placing glyphs ~1.0pt LOWER than Oxi's IR
`element.y` intended. Subtracting 1.0pt from the DrawTextLayout origin
aligns rendered text with Word's actual visual placement.

```rust
let origin = D2D_POINT_2F {
    x: x_pt * PT_TO_DIP,
    y: (y_pt - 1.0) * PT_TO_DIP,    // <-- 1.0pt UP
};
```

## Why this works (mechanism)

DirectWrite's `DrawTextLayout(origin, ...)` semantics:
- `origin.y` = top-left of the layout box
- Glyph baseline is positioned at `origin.y + font_ascent_in_dips`
- Glyph top sits at `baseline - cap_height` (typically below origin)

Oxi's IR `element.y` is intended to be glyph TOP (matching Word's COM
`Information(6)` in grid mode — see spec §13.6 round 2). But DirectWrite's
default rendering puts glyphs ~1pt BELOW origin.y due to layout-box
internal leading + ascent gap.

The 1.0pt origin shift compensates, aligning DWrite's actual glyph top
with Oxi's intended position (and thus with Word's visual placement).

This is a RENDERER-level fix, NOT a layout fix. Spec §13.6 RESOLUTION
explicitly retracts the layout-side fix recommendation; the bug is in
the DWrite glyph placement, not in cursor_y / pad_t / cell_text_y_off.

## Sweep result (177-doc p.1 corpus)

| shift | NET SSIM | wins | regs | marginal |
|---:|---:|---:|---:|---:|
| 0pt (baseline) | 0 | 0 | 0 | — |
| -0.25pt | +0.9116 | 115 | 22 | +0.91 |
| -0.5pt | +1.4013 | 117 | 29 | +0.49 |
| -0.75pt | +1.8966 | 116 | 34 | +0.50 |
| -0.875pt | +2.1176 | 114 | 39 | +0.22 |
| **-1.0pt** | **+2.2198** | **110** | **41** | **+0.10** ← PEAK |
| -1.25pt | +2.0019 | 107 | 48 | -0.22 |

Peak NET = +2.2198 at -1.0pt. Marginal gain turns negative beyond -1.0pt.

## Per-doc top wins (-1.0pt)

(Sample from -0.75pt log; -1.0pt very similar but slightly more wins
+ slightly bigger regressions:)

| Δ | doc |
|---:|---|
| +0.04 | gen2_017 マーケティング戦略書 |
| +0.03 | gen2_031 文書管理規程 |
| +0.03 | gen2_002 営業計画書 |
| +0.03 | gen2_013 コンプライアンス研修 |
| +0.02 | c7b923 outline_06 |
| +0.02 | gen2_050 Disaster_Plan |
| ... 110 docs total +0.001..+0.04 each |

## Per-doc top regressions (-1.0pt)

Concentrated on docs with explicit table borders / specific Word visual
quirks — these docs apparently DO render at the unshifted Oxi position
already (so the shift moves them away from Word):

| Δ | doc |
|---:|---|
| ~-0.085 | 683ff open_data_contract_addon |
| ~-0.040 | d77a outline_08 |
| ~-0.030 | e8caed kyodokenkyuyoushiki07 |
| ~-0.030 | a5ccbe kyodokenkyuyoushiki05 |
| ~-0.025 | 6d6dc4 index-20 |
| ~-0.020 | 1ec1 |
| ... 41 docs total -0.001..-0.085 |

The biggest regressor (683ff -0.085) is acceptable given the +110-doc
aggregate gain.

## Comparison with §13.6 layout fix attempts

| approach | net | notes |
|---|---:|---|
| §13.6 layout 3-fix (ungated) | -0.05 (bottom-15) | reverted |
| §13.6 layout 3-fix (gated explicit_borders) | -0.328 (full) | reverted |
| §13.6 layout fix #1+#2 only | -0.254 (full) | reverted |
| **DWrite -1.0pt origin shift** | **+2.2198 (full)** | **SHIPPING** |

Layout-side fixes failed because they shifted the IR position (which
Oxi's glyph-top model says is correct), causing cascade misalignment.
The DWrite renderer fix corrects only the visual rendering, not the IR.

## Why -1.0pt and not -0.5pt?

The 1pt magnitude is empirical. Hypothesized origin:

- For 11pt body text with line_height ~16.5pt:
  - Word visual glyph top ≈ flow_y + 3.3pt
  - Oxi pre-fix: element_y + 0.8pt → effective flow_y + 3.8pt
  - Drift: +0.5pt for 11pt body
- For 14pt heading with line_height ~21pt:
  - Word visual glyph top ≈ flow_y + 2.9pt
  - Oxi pre-fix: element_y + 1.3pt → effective flow_y + 4.3pt
  - Drift: +1.4pt for 14pt heading
- Optimal compromise across font sizes: ~1.0pt up

The exact value (1.0pt) likely depends on font-size mix in the corpus.
A future refinement would make the shift font-size-dependent (e.g.,
shift = font_size × 0.07 - 0.2 or similar).

## Files / data

- Patched code: `tools/oxi-dwrite-renderer/src/main.rs` (render_text origin)
- Sweep canary logs:
  - `c:/tmp/shift025_canary.log`
  - `c:/tmp/shift05_canary.log`
  - `c:/tmp/shift075_canary.log`
  - `c:/tmp/shift0875_canary.log`
  - `c:/tmp/shift10_canary.log`
  - `c:/tmp/shift125_canary.log`
- Source measurements: `c:/tmp/heading_lh.log`,
  `c:/tmp/gen2_003_予算申請書_layout.json`,
  `c:/tmp/gen2_023_育児休業規程_layout.json`
- Spec context: `docs/spec/word_layout_spec_ra.md` §13.6 (round 2 RESOLUTION)
- Sister investigation MD: `pipeline_data/spec_13_6_implementation_attempt_2026-05-03.md`

## Open follow-ups

- **Font-size-dependent shift** — the 1.0pt is a corpus average; investigate
  if making shift = f(font_size) gives even larger gains.
- **Investigate -0.085 regression on 683ff** — biggest single regressor.
  Likely has a docGrid or vAlign config that's inverse-correlated.
- **Apply same shift to GDI renderer?** GDI was the previous default; may
  have a similar bug. Untested in this canary.
