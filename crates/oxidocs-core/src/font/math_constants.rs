//! Cambria Math OpenType MATH table constants, bundled at compile time.
//!
//! The JSON dump (`data/cambria_math_constants.json`) is extracted from
//! `C:/Windows/Fonts/cambria.ttc` subfont 1 via `fontTools`. All values
//! are in design units; UPM = 2048.
//!
//! See `docs/spec/omml_notes.md` / `docs/spec/omml_phase1_summary.md`
//! for the complete set of 56 constants and their usage in math layout.

use serde::{Deserialize, Serialize};
use std::collections::HashMap;

/// Full MATH table constants (subset used for layout).
///
/// Fields match names from ECMA-376 §22.1 / OpenType MATH spec.
#[derive(Debug, Clone, Default, Serialize, Deserialize)]
#[serde(default)]
#[allow(non_snake_case)]
pub struct MathConstants {
    // Superscript geometry
    pub SuperscriptShiftUp: i32,
    pub SuperscriptShiftUpCramped: i32,
    pub SuperscriptBottomMin: i32,
    pub SuperscriptBaselineDropMax: i32,
    pub SuperscriptBottomMaxWithSubscript: i32,

    // Subscript geometry
    pub SubscriptShiftDown: i32,
    pub SubscriptTopMax: i32,
    pub SubscriptBaselineDropMin: i32,
    pub SubSuperscriptGapMin: i32,
    pub SpaceAfterScript: i32,

    // Fraction (inline)
    pub FractionNumeratorShiftUp: i32,
    pub FractionDenominatorShiftDown: i32,
    pub FractionNumeratorGapMin: i32,
    pub FractionDenominatorGapMin: i32,
    pub FractionRuleThickness: i32,

    // Fraction (display)
    pub FractionNumeratorDisplayStyleShiftUp: i32,
    pub FractionDenominatorDisplayStyleShiftDown: i32,
    pub FractionNumDisplayStyleGapMin: i32,
    pub FractionDenomDisplayStyleGapMin: i32,

    // Skewed fractions
    pub SkewedFractionHorizontalGap: i32,
    pub SkewedFractionVerticalGap: i32,

    // Radical
    pub RadicalVerticalGap: i32,
    pub RadicalDisplayStyleVerticalGap: i32,
    pub RadicalRuleThickness: i32,
    pub RadicalExtraAscender: i32,
    pub RadicalKernBeforeDegree: i32,
    pub RadicalKernAfterDegree: i32,
    pub RadicalDegreeBottomRaisePercent: i32,

    // Nary operator limits
    pub UpperLimitGapMin: i32,
    pub UpperLimitBaselineRiseMin: i32,
    pub LowerLimitGapMin: i32,
    pub LowerLimitBaselineDropMin: i32,

    // Stack (equation array)
    pub StackTopShiftUp: i32,
    pub StackTopDisplayStyleShiftUp: i32,
    pub StackBottomShiftDown: i32,
    pub StackBottomDisplayStyleShiftDown: i32,
    pub StackGapMin: i32,
    pub StackDisplayStyleGapMin: i32,

    // Stretch stack
    pub StretchStackTopShiftUp: i32,
    pub StretchStackBottomShiftDown: i32,
    pub StretchStackGapAboveMin: i32,
    pub StretchStackGapBelowMin: i32,

    // Overbar / Underbar
    pub OverbarVerticalGap: i32,
    pub OverbarRuleThickness: i32,
    pub OverbarExtraAscender: i32,
    pub UnderbarVerticalGap: i32,
    pub UnderbarRuleThickness: i32,
    pub UnderbarExtraDescender: i32,

    // Miscellaneous
    pub AxisHeight: i32,
    pub AccentBaseHeight: i32,
    pub FlattenedAccentBaseHeight: i32,
    pub MathLeading: i32,
    pub ScriptPercentScaleDown: i32,     // e.g., 73 = 73%
    pub ScriptScriptPercentScaleDown: i32, // e.g., 60 = 60%
    pub DelimitedSubFormulaMinHeight: i32,
    pub DisplayOperatorMinHeight: i32,
}

/// Full MATH table payload: constants + metadata.
#[allow(dead_code)]
#[derive(Debug, Clone, Deserialize)]
struct MathConstantsFile {
    font: String,
    upm: u32,
    math_constants: MathConstants,
    #[serde(default)]
    glyph_info: HashMap<String, serde_json::Value>,
    #[serde(default)]
    variants: HashMap<String, serde_json::Value>,
}

/// Loaded Cambria Math MATH table constants + UPM for scaling.
#[derive(Debug, Clone)]
pub struct MathTable {
    pub upm: u32,
    pub constants: MathConstants,
}

impl MathTable {
    /// Load the bundled Cambria Math constants. Parses at first call.
    /// Uses `include_str!` so the JSON is baked into the binary at build.
    pub fn cambria_math() -> &'static Self {
        use std::sync::OnceLock;
        static TABLE: OnceLock<MathTable> = OnceLock::new();
        TABLE.get_or_init(|| {
            let json = include_str!("data/cambria_math_constants.json");
            let file: MathConstantsFile = serde_json::from_str(json)
                .expect("cambria_math_constants.json parse");
            MathTable { upm: file.upm, constants: file.math_constants }
        })
    }

    /// Convert a design-unit value to points at the given font size.
    ///
    /// `pts = value * font_size / upm`
    #[inline]
    pub fn du_to_pt(&self, du: i32, font_size: f32) -> f32 {
        du as f32 * font_size / self.upm as f32
    }

    /// Script scale factor at depth 1 (sub/sup). Typically 0.73 for Cambria Math.
    #[inline]
    pub fn script_scale(&self) -> f32 {
        self.constants.ScriptPercentScaleDown as f32 / 100.0
    }

    /// Script-script scale factor at depth 2+ (nested scripts). Typically 0.60.
    #[inline]
    pub fn script_script_scale(&self) -> f32 {
        self.constants.ScriptScriptPercentScaleDown as f32 / 100.0
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn cambria_math_loads() {
        let t = MathTable::cambria_math();
        assert_eq!(t.upm, 2048);
        // A few sanity-check constants we measured in Phase 1
        assert_eq!(t.constants.SuperscriptShiftUp, 750);
        assert_eq!(t.constants.SubscriptShiftDown, 418);
        assert_eq!(t.constants.FractionRuleThickness, 133);
        assert_eq!(t.constants.AxisHeight, 585);
        assert_eq!(t.constants.ScriptPercentScaleDown, 73);
        assert_eq!(t.constants.ScriptScriptPercentScaleDown, 60);
    }

    #[test]
    fn du_to_pt_scaling() {
        let t = MathTable::cambria_math();
        // SuperscriptShiftUp at 10.5pt: 750 * 10.5 / 2048 ≈ 3.845pt
        let v = t.du_to_pt(750, 10.5);
        assert!((v - 3.845).abs() < 0.01, "got {v}");

        // FractionRuleThickness at 10.5pt: 133 * 10.5 / 2048 ≈ 0.682pt
        let v = t.du_to_pt(133, 10.5);
        assert!((v - 0.682).abs() < 0.01, "got {v}");

        // AxisHeight at 12pt: 585 * 12 / 2048 ≈ 3.428pt
        let v = t.du_to_pt(585, 12.0);
        assert!((v - 3.428).abs() < 0.01, "got {v}");
    }

    #[test]
    fn script_scale_factors() {
        let t = MathTable::cambria_math();
        assert!((t.script_scale() - 0.73).abs() < 0.001);
        assert!((t.script_script_scale() - 0.60).abs() < 0.001);
    }

    #[test]
    fn bundled_display_constants() {
        let t = MathTable::cambria_math();
        // Display-style constants are taller than inline-style.
        assert!(t.constants.FractionNumeratorDisplayStyleShiftUp
            > t.constants.FractionNumeratorShiftUp);
        assert!(t.constants.FractionDenominatorDisplayStyleShiftDown
            > t.constants.FractionDenominatorShiftDown);
        assert!(t.constants.RadicalDisplayStyleVerticalGap
            > t.constants.RadicalVerticalGap);
    }
}
