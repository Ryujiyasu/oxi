// This Source Code Form is subject to the terms of the Mozilla Public
// License, v. 2.0. If a copy of the MPL was not distributed with this
// file, You can obtain one at https://mozilla.org/MPL/2.0/.

//! Ruby (furigana) layout helpers.
//!
//! Pure-function module deriving geometric values for ruby annotations
//! from the parsed `Ruby` IR. Empirical formulas calibrated against Word
//! COM measurements; see `docs/spec/word_layout_spec_ra.md` §18 for the
//! measurement basis (fixtures RUBY_V1 .. RUBY_V10).
//!
//! All callers should use these helpers rather than re-deriving the
//! formula at the call site.

use crate::ir::Ruby;

/// Word's default ruby raise (height above base baseline) when
/// `<w:hpsRaise>` is omitted, in pt — calibrated for 10.5pt MS Mincho
/// base. Confirmed empirically: V6 `RUBY_V6_hpsRaise.docx` showed
/// `raise=None` and `raise=18` (= 9pt) produce identical line-height
/// expansion at 10.5pt base. See spec §18.4.
pub const DEFAULT_HPS_RAISE_PT: f32 = 9.0;

/// V13 (round 9, 2026-04-27) generalised default-raise: when
/// `<w:hpsRaise>` is omitted, Word's default depends on both base
/// and hps (no clean closed-form, see spec §18.7). The dominant
/// pattern: for the typical "hps = base/2" ruby case, default_raise
/// ≈ base − 1pt across the V13 grid (base ∈ {9, 11, 12, 14}pt with
/// derived defaults 8.34, 9.55, 10.79, 13.25 — Δ from base−1 ∈
/// [−0.45, +0.25]pt, well within Word's ±0.5pt rounding noise).
///
/// R75 (2026-04-29) implements option (c) from spec §18.9: when the
/// (currently-set) hps is approximately base/2, scale the default
/// raise as `base_pt − 1.0`. Falls back to the 10.5pt constant 9.0pt
/// otherwise — the non-typical `hps ≈ base` case empirically maps
/// to `default_raise ≈ base + 0.5pt` but isn't covered until a
/// fixture surfaces it.
fn default_hps_raise_pt(base_pt: f32, hps_pt: f32) -> f32 {
    // Round-8/V6 calibrated invariant: at base = 10.5pt the empirical
    // default_raise is exactly 9pt (Word measurement V6, hps-independent
    // across tested range). The V13 derived "base − 1pt" rule gives 9.5pt
    // there, off by 0.5pt — small but enough to break the pre-R74 test
    // suite. Keep the legacy constant for the 10.5pt anchor and apply
    // V13 scaling only outside that anchor.
    if (base_pt - 10.5).abs() < 0.01 {
        return DEFAULT_HPS_RAISE_PT;
    }
    if base_pt <= 0.0 {
        return DEFAULT_HPS_RAISE_PT;
    }
    // R75: typical "hps = base/2" case → default_raise ≈ base − 1pt.
    // R85 (2026-04-29): extreme "hps = base" case → default_raise ≈
    // base + 0.5pt. V13 grid (round 9, 2026-04-27) derived defaults at
    // base ∈ {9, 11, 12, 14}pt × hps=base = {9.46, 11.43, 12.29, 14.50},
    // all within ±0.5pt of `base + 0.5`. Spec §18.7 records this as
    // the "hps doubled" empirical pattern; R85 wires option (c) for
    // the extreme too, in addition to R75's typical case.
    //
    // The dispatch is by hps_pt vs base/2 closeness:
    //   hps ≈ base/2  → R75 rule (typical ruby)
    //   hps ≈ base    → R85 rule (extreme; e.g. CJK-base-sized ruby)
    //   else          → fall back to 10.5pt anchor (DEFAULT_HPS_RAISE_PT)
    let hps_is_half_base = (hps_pt - base_pt / 2.0).abs() < 0.5;
    let hps_is_full_base = (hps_pt - base_pt).abs() < 0.5;
    if hps_is_half_base {
        (base_pt - 1.0).max(0.0)
    } else if hps_is_full_base {
        base_pt + 0.5
    } else {
        DEFAULT_HPS_RAISE_PT
    }
}

/// Word's typical line-box ascent above the base baseline at 10.5pt
/// MS Mincho, in pt. Empirical anchor for the expansion formula —
/// V3-derived when base = 10.5pt; V13 (round 9, 2026-04-27) confirmed
/// the formula generalises linearly via `base_pt × 9/10.5` across
/// {9, 10.5, 11, 12, 14}pt MS Mincho with explicit hpsRaise (spec §18.7).
pub const LINE_BOX_ASCENT_PT_AT_10_5: f32 = 9.0;

/// V13-confirmed scaling: line-box ascent at arbitrary base size, in pt.
/// Formula: `base_pt × 9 / 10.5`. Returns the constant 9.0 when base = 10.5pt
/// (the legacy round-8 calibration), and scales linearly for other sizes.
fn line_box_ascent_pt(base_pt: f32) -> f32 {
    base_pt * (LINE_BOX_ASCENT_PT_AT_10_5 / 10.5)
}

/// Approximate ratio of ruby ascent to ruby font size (how much of
/// the ruby's height extends above its baseline). Empirical fit.
/// See spec §18.4.
pub const HPS_ASCENT_RATIO: f32 = 0.75;

/// Compute the line-height expansion (paragraph-tail) added by a
/// ruby annotation, in pt.
///
/// Formula: `max(0, hpsRaise_pt + 0.75 × hps_pt − base_pt × 9/10.5)`.
/// Calibrated against V3 fixtures (round 8) for 10.5pt MS Mincho base,
/// then generalised by V13 (round 9, 2026-04-27) to base ∈
/// {9, 10.5, 11, 12, 14}pt MS Mincho. The `base_pt × 9/10.5` term
/// reduces to the original 9pt constant when base = 10.5pt, so all
/// pre-existing 10.5pt tests stay green. See spec §18.4 and §18.7.
///
/// `base_pt` is the base text font size, used to derive defaults
/// when `hps_halfpt` is unset and as the scaling input for the
/// line-box ascent.
pub fn ruby_expansion_pt(ruby: &Ruby, base_pt: f32) -> f32 {
    let hps_pt = ruby.hps_halfpt
        .map(|h| h as f32 / 2.0)
        .unwrap_or(base_pt / 2.0);
    let hps_raise_pt = ruby.hps_raise_halfpt
        .map(|h| h as f32 / 2.0)
        .unwrap_or_else(|| default_hps_raise_pt(base_pt, hps_pt));
    let raw = hps_raise_pt + HPS_ASCENT_RATIO * hps_pt - line_box_ascent_pt(base_pt);
    raw.max(0.0)
}

/// Compute the inline width of a ruby field — the horizontal advance
/// it consumes in the line box.
///
/// Per V2 measurement (spec §18.5): `field_w = max(base_w, ruby_w)`,
/// invariant across all 5 `rubyAlign` modes.
///
/// `base_text_w_pt` and `ruby_text_w_pt` should both be measured at
/// their respective font sizes.
#[allow(dead_code)]
pub fn ruby_field_width_pt(base_text_w_pt: f32, ruby_text_w_pt: f32) -> f32 {
    base_text_w_pt.max(ruby_text_w_pt)
}

/// Position the ruby annotation horizontally relative to the base text,
/// per ECMA-376 §17.3.3.26 (CT_RubyAlign).
///
/// Returns `(ruby_x_offset, char_spacing_pt)` where:
/// - `ruby_x_offset` is the X offset from `base_x` to the start of the
///   ruby annotation (so ruby_screen_x = base_x + ruby_x_offset)
/// - `char_spacing_pt` is per-character extra spacing for the ruby text
///   (used by `distributeLetter` and `distributeSpace` to spread the
///   annotation across the base width).
///
/// Inputs:
/// - `base_w_pt`: rendered width of the full base text
/// - `ruby_w_pt`: rendered width of the ruby annotation (sum of glyph
///   advances at hps font size, no extra spacing)
/// - `ruby_char_count`: number of ruby characters (used to compute the
///   per-char spacing for distribute modes)
/// - `align`: rubyAlign mode (None defaults to Center)
pub fn ruby_position(
    base_w_pt: f32,
    ruby_w_pt: f32,
    ruby_char_count: usize,
    align: Option<crate::ir::RubyAlign>,
) -> (f32, f32) {
    use crate::ir::RubyAlign;
    let mode = align.unwrap_or(RubyAlign::Center);
    match mode {
        RubyAlign::Center | RubyAlign::RightVertical => {
            // Center the ruby horizontally over the base.
            ((base_w_pt - ruby_w_pt) / 2.0, 0.0)
        }
        RubyAlign::Left => {
            // Both base and ruby left-aligned at field origin.
            (0.0, 0.0)
        }
        RubyAlign::Right => {
            // Both base and ruby right-aligned within the field.
            // Field width = max(base_w, ruby_w). Ruby ends at field's
            // right edge, which equals base_w when base is wider, or
            // ruby_w (= base_x + ruby_w) otherwise. Either way, ruby
            // start = base_w - ruby_w (which is negative when ruby is
            // wider, indicating overhang to the left).
            (base_w_pt - ruby_w_pt, 0.0)
        }
        RubyAlign::DistributeLetter => {
            // Ruby chars distributed evenly across the BASE width, with
            // half-padding at each end. Total spacing across N chars is
            // `base_w - ruby_w`; each gap (between chars + ends) gets
            // an equal share. With N chars we have N+1 gaps if we count
            // half-padding at both ends, but the ECMA-376 spec says the
            // ruby starts and ends with half-spacing — so per-char extra
            // = (base_w - ruby_w) / N.
            // Each char's start offset cumulates by per-char extra.
            // First char's left edge has half-spacing as offset.
            if ruby_char_count == 0 {
                ((base_w_pt - ruby_w_pt) / 2.0, 0.0)
            } else {
                let extra_total = base_w_pt - ruby_w_pt;
                let per_char = extra_total / ruby_char_count as f32;
                (per_char / 2.0, per_char)
            }
        }
        RubyAlign::DistributeSpace => {
            // Like distributeLetter but distributes across the FIELD
            // (max(base_w, ruby_w)) with extra padding at the ends. For
            // base_w >= ruby_w (typical), this matches distributeLetter
            // semantically since field_w = base_w. We treat them the
            // same here; refinement may be needed when ruby_w > base_w.
            if ruby_char_count == 0 {
                ((base_w_pt - ruby_w_pt) / 2.0, 0.0)
            } else {
                let field_w = base_w_pt.max(ruby_w_pt);
                let extra_total = field_w - ruby_w_pt;
                let per_char = extra_total / (ruby_char_count + 1) as f32;
                let start_offset = per_char + (base_w_pt - field_w);
                (start_offset, per_char)
            }
        }
    }
}

/// Return the maximum ruby expansion across all runs in a paragraph.
/// 0.0 if no run contains a ruby annotation.
///
/// Used for paragraph-tail line-height expansion. Per V7 measurement,
/// the expansion is paragraph-level (applied at paragraph end), not
/// per-line.
pub fn paragraph_ruby_expansion_pt(runs: &[crate::ir::Run], para_font_size_pt: f32) -> f32 {
    let mut max_exp: f32 = 0.0;
    for run in runs {
        if let Some(ref ruby) = run.ruby {
            let base_pt = run.style.font_size.unwrap_or(para_font_size_pt);
            let exp = ruby_expansion_pt(ruby, base_pt);
            if exp > max_exp {
                max_exp = exp;
            }
        }
    }
    max_exp
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::ir::{Ruby, RubyAlign};

    fn ruby(hps: Option<u32>, raise: Option<u32>) -> Ruby {
        Ruby {
            base: "漢字".into(),
            text: "かんじ".into(),
            font_size: hps.map(|h| h as f32 / 2.0),
            align: None,
            hps_halfpt: hps,
            hps_raise_halfpt: raise,
            hps_base_text_halfpt: Some(21),
            lang: None,
        }
    }

    /// V3 single-axis hps measurements at default raise (=9pt) on
    /// 10.5pt MS Mincho base. The empirical fit is the formula
    /// max(0, 9 + 0.75 × hps_pt − 9) = 0.75 × hps_pt; this is within
    /// ±0.5pt of measured values, reflecting Word's internal rounding.
    #[test]
    fn ruby_expansion_default_raise_matches_v3_v8() {
        // (hps_halfpt, hps_pt, expected_predicted_pt within tolerance)
        let cases = [
            (5,  2.5, 1.875),
            (7,  3.5, 2.625),
            (9,  4.5, 3.375),
            (11, 5.5, 4.125),
            (13, 6.5, 4.875),
            (15, 7.5, 5.625),
        ];
        for (hps, hps_pt, expected) in cases {
            let r = ruby(Some(hps), None);
            let exp = ruby_expansion_pt(&r, 10.5);
            assert!(
                (exp - expected).abs() < 0.001,
                "hps={hps_pt} expected {expected:.3} got {exp:.3}"
            );
        }
    }

    /// V6 explicit hpsRaise observations on hps=11 base 10.5pt.
    /// The formula matches measured within 0.5pt for raise ∈ {6, 9, 12, 18}.
    #[test]
    fn ruby_expansion_explicit_raise_matches_v6() {
        // (raise_halfpt, raise_pt, measured_pt) — formula prediction = max(0, raise + 0.75*5.5 - 9)
        let cases = [
            (6, 3.0, 0.0),    // pred max(0, 3 + 4.125 - 9) = 0
            (12, 6.0, 1.5),   // pred max(0, 6 + 4.125 - 9) = 1.125; measured 1.5
            (18, 9.0, 4.0),   // pred 4.125; measured 4.0
            (24, 12.0, 7.0),  // pred 7.125; measured 7.0
            (36, 18.0, 13.0), // pred 13.125; measured 13.0
        ];
        for (raise, _raise_pt, measured) in cases {
            let r = ruby(Some(11), Some(raise));
            let exp = ruby_expansion_pt(&r, 10.5);
            // Measured-to-formula tolerance: 0.5pt (Word internal rounding)
            assert!(
                (exp - measured).abs() <= 0.5,
                "raise={raise} measured {measured} got {exp}"
            );
        }
    }

    /// V13 (round 9, 2026-04-27): the line-box ascent constant scales
    /// linearly with base_pt as `base × 9/10.5`. R74 (2026-04-29) wires
    /// this into `ruby_expansion_pt`. Check against V13's grid spanning
    /// base ∈ {9, 11, 12, 14}pt × raise ∈ {6, 12}pt — predicted values
    /// must reproduce the formula exactly; measured values agree within
    /// ±0.5pt Word-rounding tolerance (see spec §18.7).
    #[test]
    fn ruby_expansion_v13_grid_matches_predicted_and_measured() {
        // (base_pt, raise_halfpt, hps_halfpt, predicted, measured)
        // Predicted = base × 9/10.5 line-box ascent, formula in §18.7.
        // Measured = Word COM (V13 fixtures), tolerance 0.5pt.
        let cases = [
            ( 9.0_f32,  12,  9, 1.661_f32, 1.50_f32),  // base=9, raise=6,  hps=4.5
            ( 9.0_f32,  24,  9, 7.661_f32, 7.50_f32),  // base=9, raise=12, hps=4.5
            (11.0_f32,  12, 11, 0.696_f32, 0.25_f32),  // base=11, raise=6,  hps=5.5
            (11.0_f32,  24, 11, 6.696_f32, 6.75_f32),  // base=11, raise=12, hps=5.5
            (12.0_f32,  12, 12, 0.214_f32, 0.00_f32),  // base=12, raise=6,  hps=6.0
            (12.0_f32,  24, 12, 6.214_f32, 6.00_f32),  // base=12, raise=12, hps=6.0
            (14.0_f32,  12, 14, 0.000_f32, 0.00_f32),  // base=14, raise=6,  hps=7.0  (clamped to 0)
            (14.0_f32,  24, 14, 5.250_f32, 5.50_f32),  // base=14, raise=12, hps=7.0
        ];
        for (base_pt, raise_halfpt, hps_halfpt, predicted, measured) in cases {
            let r = Ruby {
                base: "x".into(),
                text: "y".into(),
                font_size: Some(hps_halfpt as f32 / 2.0),
                align: None,
                hps_halfpt: Some(hps_halfpt),
                hps_raise_halfpt: Some(raise_halfpt),
                hps_base_text_halfpt: Some((base_pt * 2.0) as u32),
                lang: None,
            };
            let exp = ruby_expansion_pt(&r, base_pt);
            assert!(
                (exp - predicted).abs() < 0.01,
                "V13 base={base_pt} raise_hp={raise_halfpt} hps_hp={hps_halfpt}: \
                 expected predicted {predicted:.3}, got {exp:.3}"
            );
            assert!(
                (exp - measured).abs() <= 0.5,
                "V13 measured tolerance: base={base_pt} raise_hp={raise_halfpt} \
                 measured {measured:.2}, got {exp:.3}"
            );
        }
    }

    /// 10.5pt base must still produce identical results to the legacy
    /// constant=9 formula — V13 generalisation collapses to `9` when
    /// base = 10.5. Guards the round-7-era ship from R74's refactor.
    #[test]
    fn ruby_expansion_10pt5_base_unchanged_by_v13_generalisation() {
        // Round-8 ship case: base=10.5, hps=11 (=5.5pt), raise=18 (=9pt).
        // Pre-R74 formula: max(0, 9 + 0.75×5.5 − 9) = 4.125pt.
        let r = ruby(Some(11), Some(18));
        let exp = ruby_expansion_pt(&r, 10.5);
        assert!(
            (exp - 4.125).abs() < 0.001,
            "10.5pt base: expected 4.125pt, got {exp}"
        );
    }

    /// R85 (2026-04-29): the "hps = base" extreme ruby case. V13 grid
    /// derived defaults {9.46, 11.43, 12.29, 14.50} at base ∈
    /// {9, 11, 12, 14}pt × hps=base. Implementation rule:
    /// `default_raise ≈ base + 0.5pt` (spec §18.7), within ±0.5pt
    /// vs Word measured. Sister test to the R75 hps=base/2 grid.
    #[test]
    fn ruby_default_raise_v13_grid_hps_equals_base() {
        // (base_pt, hps_halfpt, V13_measured_default_raise)
        // hps_halfpt = base × 2 (= base in pt, so hps_pt == base_pt)
        let cases = [
            ( 9.0_f32, 18,  9.46_f32),
            (11.0_f32, 22, 11.43_f32),
            (12.0_f32, 24, 12.29_f32),
            (14.0_f32, 28, 14.50_f32),
        ];
        for (base_pt, hps_halfpt, measured) in cases {
            let hps_pt = hps_halfpt as f32 / 2.0;
            // Helper picks the R85 branch (`base + 0.5`):
            assert!(
                (default_hps_raise_pt(base_pt, hps_pt) - (base_pt + 0.5)).abs() < 0.001,
                "base={base_pt} hps={hps_pt} (= base): expected base+0.5, got {}",
                default_hps_raise_pt(base_pt, hps_pt)
            );
            // End-to-end via ruby_expansion_pt with raise=None: prediction
            // matches V13 measured within ±0.5pt §18.9 (c) tolerance.
            let r = Ruby {
                base: "x".into(),
                text: "y".into(),
                font_size: Some(hps_pt),
                align: None,
                hps_halfpt: Some(hps_halfpt),
                hps_raise_halfpt: None, // exercises default_hps_raise_pt
                hps_base_text_halfpt: Some((base_pt * 2.0) as u32),
                lang: None,
            };
            let predicted = (base_pt + 0.5) + 0.75 * hps_pt - base_pt * (9.0 / 10.5);
            let predicted = predicted.max(0.0);
            let exp = ruby_expansion_pt(&r, base_pt);
            assert!(
                (exp - predicted).abs() < 0.01,
                "base={base_pt}: predicted {predicted:.3}, got {exp:.3}"
            );
            // V13 measured expansion = derived_default_raise + 0.75×hps − base×9/10.5
            let v13_expansion = measured + 0.75 * hps_pt - base_pt * (9.0 / 10.5);
            let v13_expansion = v13_expansion.max(0.0);
            assert!(
                (exp - v13_expansion).abs() <= 0.5,
                "V13 measured tolerance: base={base_pt} v13 {v13_expansion:.3}, got {exp:.3}"
            );
        }
    }

    /// R75 (2026-04-29): default_hps_raise_pt scales with base for the
    /// typical "hps = base/2" ruby case. V13 grid (round 9, 2026-04-27)
    /// derived these defaults at base ∈ {9, 11, 12, 14}pt with hps = base/2.
    /// Implementation uses the simpler `default_raise = base − 1pt`
    /// approximation per spec §18.9 (c), accepting ±0.5pt vs Word.
    #[test]
    fn ruby_default_raise_scales_with_base_v13_grid() {
        // (base_pt, hps_halfpt, V13_measured_default_raise)
        // hps_halfpt = base × 2 / 2 (= base/2 in pt when halfpt-encoded)
        let cases = [
            ( 9.0_f32,  9, 8.34_f32),  // base=9, hps=4.5 (= 9/2)
            (11.0_f32, 11, 9.55_f32),  // base=11, hps=5.5
            (12.0_f32, 12, 10.79_f32), // base=12, hps=6.0
            (14.0_f32, 14, 13.25_f32), // base=14, hps=7.0
        ];
        for (base_pt, hps_halfpt, measured) in cases {
            let hps_pt = hps_halfpt as f32 / 2.0;
            // Direct helper check against the simpler "base − 1pt" rule:
            assert!(
                (default_hps_raise_pt(base_pt, hps_pt) - (base_pt - 1.0)).abs() < 0.001,
                "base={base_pt} hps={hps_pt}: expected base−1, \
                 got {}", default_hps_raise_pt(base_pt, hps_pt)
            );
            // End-to-end: ruby_expansion_pt with raise=None should now use
            // the scaled default. Compute predicted expansion via the same
            // formula and compare against V13's measured value within
            // ±0.5pt rounding (the spec §18.9 (c) tolerance).
            let r = Ruby {
                base: "x".into(),
                text: "y".into(),
                font_size: Some(hps_pt),
                align: None,
                hps_halfpt: Some(hps_halfpt),
                hps_raise_halfpt: None, // ← exercises default_hps_raise_pt
                hps_base_text_halfpt: Some((base_pt * 2.0) as u32),
                lang: None,
            };
            let predicted = (base_pt - 1.0) + 0.75 * hps_pt - base_pt * (9.0 / 10.5);
            let predicted = predicted.max(0.0);
            let exp = ruby_expansion_pt(&r, base_pt);
            assert!(
                (exp - predicted).abs() < 0.01,
                "base={base_pt}: predicted {predicted:.3}, got {exp:.3}"
            );
            // V13 measured is hps × 0.75 + (base−1) − base × 9/10.5; the
            // V13 paper itself notes the cell-by-cell measurement:
            //   base=9 / hps=4.5: predicted (8 + 3.375 - 7.714) = 3.66
            //                     (= measured (8.34 + 3.375 - 7.714) = 4.0 within 0.5pt)
            // The V13 column gives expansion (computed with explicit raise);
            // for the default-raise case the relevant comparison is
            // expansion vs no_ruby_LH measurement which §18.7 derives back
            // to default_raise. Cell-by-cell back-derivation:
            //   default_raise_derived = measured_exp + base × 9/10.5 − 0.75 × hps
            // Re-deriving from above: measured_default_raise was the value;
            // expansion = (default_raise_derived) + 0.75 × hps − base × 9/10.5.
            // So measured expansion = default_raise_derived + 0.75 × hps − base × 9/10.5.
            // If our formula gives base−1 instead of derived, the expansion
            // diff equals (base − 1) − default_raise_derived.
            // Within ±0.5pt across V13 grid (max 0.45pt at base=11).
            let v13_expansion = measured + 0.75 * hps_pt - base_pt * (9.0 / 10.5);
            let v13_expansion = v13_expansion.max(0.0);
            assert!(
                (exp - v13_expansion).abs() <= 0.5,
                "V13 measured tolerance: base={base_pt} v13_expansion {v13_expansion:.3}, got {exp:.3}"
            );
        }
    }

    /// Default raise is treated as 9pt regardless of hps (V6 finding).
    #[test]
    fn ruby_default_raise_is_9pt() {
        let no_raise = ruby(Some(11), None);
        let raise_18 = ruby(Some(11), Some(18));
        let exp_default = ruby_expansion_pt(&no_raise, 10.5);
        let exp_explicit = ruby_expansion_pt(&raise_18, 10.5);
        assert!(
            (exp_default - exp_explicit).abs() < 0.001,
            "default raise should equal raise=18 (=9pt), got default {exp_default} vs explicit {exp_explicit}"
        );
    }

    /// Field width = max(base_w, ruby_w) regardless of alignment (V2 finding).
    #[test]
    fn ruby_field_width_takes_max() {
        // V2 case: base 21pt (2 chars × 10.5pt), ruby 22pt (4 chars × 5.5pt)
        assert_eq!(ruby_field_width_pt(21.0, 22.0), 22.0);
        assert_eq!(ruby_field_width_pt(22.0, 21.0), 22.0);
        // Equal widths: max returns one of them
        assert_eq!(ruby_field_width_pt(20.0, 20.0), 20.0);
        // Negative (degenerate) — should still pick larger
        assert_eq!(ruby_field_width_pt(0.0, 5.0), 5.0);
    }

    /// `paragraph_ruby_expansion_pt` returns 0 when no run has ruby.
    #[test]
    fn paragraph_ruby_expansion_zero_for_no_ruby() {
        let runs: Vec<crate::ir::Run> = vec![];
        assert_eq!(paragraph_ruby_expansion_pt(&runs, 10.5), 0.0);
    }

    #[test]
    fn ruby_position_center_centers_above_base() {
        // V2 case: base 21pt, ruby 22pt. Center: ruby_x = (21 - 22) / 2 = -0.5
        let (x, sp) = ruby_position(21.0, 22.0, 4, Some(RubyAlign::Center));
        assert_eq!(x, -0.5);
        assert_eq!(sp, 0.0);
        // Symmetric case: base 30, ruby 10 (3 chars). Centered offset = 10.
        let (x, sp) = ruby_position(30.0, 10.0, 3, Some(RubyAlign::Center));
        assert_eq!(x, 10.0);
        assert_eq!(sp, 0.0);
    }

    #[test]
    fn ruby_position_left_aligns_to_field_origin() {
        let (x, sp) = ruby_position(30.0, 10.0, 3, Some(RubyAlign::Left));
        assert_eq!(x, 0.0);
        assert_eq!(sp, 0.0);
    }

    #[test]
    fn ruby_position_right_aligns_to_field_end() {
        // base 30, ruby 10 → ruby starts at 30-10=20.
        let (x, sp) = ruby_position(30.0, 10.0, 3, Some(RubyAlign::Right));
        assert_eq!(x, 20.0);
        assert_eq!(sp, 0.0);
        // ruby wider than base: 21 base, 22 ruby → x=-1 (overhangs left).
        let (x, _) = ruby_position(21.0, 22.0, 4, Some(RubyAlign::Right));
        assert_eq!(x, -1.0);
    }

    #[test]
    fn ruby_position_distribute_letter_spreads_across_base() {
        // base 30, ruby 10, 5 chars: extra = 20, per-char = 4. Start = 2.
        let (x, sp) = ruby_position(30.0, 10.0, 5, Some(RubyAlign::DistributeLetter));
        assert_eq!(x, 2.0);
        assert_eq!(sp, 4.0);
    }

    #[test]
    fn ruby_position_default_to_center() {
        // align=None should fall back to Center
        let (x, sp) = ruby_position(30.0, 10.0, 3, None);
        assert_eq!(x, 10.0);
        assert_eq!(sp, 0.0);
    }

    /// `paragraph_ruby_expansion_pt` returns max across multiple ruby runs.
    #[test]
    fn paragraph_ruby_expansion_takes_max_across_runs() {
        use crate::ir::{Run, RunStyle};
        let r1 = Ruby {
            base: "a".into(),
            text: "x".into(),
            font_size: Some(5.5),
            align: Some(RubyAlign::Center),
            hps_halfpt: Some(11),
            hps_raise_halfpt: Some(18),  // 9pt → expansion 4.125
            hps_base_text_halfpt: Some(21),
            lang: None,
        };
        let r2 = Ruby {
            hps_raise_halfpt: Some(36),  // 18pt → expansion much larger
            ..r1.clone()
        };
        let mk_run = |ruby: Option<Ruby>| Run {
            text: "base".into(),
            style: RunStyle::default(),
            url: None,
            footnote_ref: None,
            endnote_ref: None,
            comment_range_start: Vec::new(),
            comment_range_end: Vec::new(),
            comment_references: Vec::new(),
            tracked_change: None,
            rpr_change: None,
            ruby,
            bookmark_name: None,
            is_math: false,
            field_type: None,
            has_last_rendered_page_break: false,
        };
        let runs = vec![mk_run(Some(r1)), mk_run(None), mk_run(Some(r2))];
        let max_exp = paragraph_ruby_expansion_pt(&runs, 10.5);
        // r2 has bigger raise → bigger expansion; max picks it.
        let r2_only = ruby_expansion_pt(
            &Ruby {
                base: "a".into(), text: "x".into(), font_size: Some(5.5),
                align: Some(RubyAlign::Center),
                hps_halfpt: Some(11), hps_raise_halfpt: Some(36),
                hps_base_text_halfpt: Some(21), lang: None,
            },
            10.5,
        );
        assert!((max_exp - r2_only).abs() < 0.001);
    }
}
