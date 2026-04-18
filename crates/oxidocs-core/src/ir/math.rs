//! OMML (Office Math Markup Language) IR types.
//!
//! Implements the type system for ECMA-376 Part 1 §22.1 Math Markup Language.
//!
//! This module defines the Intermediate Representation for parsed OMML content
//! (`<m:oMath>` inline math and `<m:oMathPara>` display math).
//!
//! # Design
//!
//! - `MathBlock`: top-level container (inline vs display)
//! - `MathExpr`: recursive expression tree covering all OMML primitives
//! - `MathStyle`: DisplayStyle/Text/Script/ScriptScript (per ECMA-376 §22.1.3.2)
//! - `MathRunStyle`: per-run style (font, bold/italic) for leaf text
//!
//! See also:
//! - `docs/spec/omml_notes.md` — ECMA-376 §22.1 reference
//! - `docs/spec/omml_phase1_summary.md` — research consolidation
//! - `tools/metrics/output/cambria_math_constants.json` — layout constants
//! - `tools/metrics/output/cambria_math_glyph_tables.json` — per-glyph data

use serde::{Deserialize, Serialize};

use super::types::RunStyle;

/// Top-level math block: either inline (flows with text) or display (own paragraph).
///
/// ECMA-376 §22.1.2.77 `oMath` (inline), §22.1.2.78 `oMathPara` (display).
#[derive(Debug, Clone, Serialize, Deserialize)]
pub enum MathBlock {
    /// Inline math, flows with surrounding body text. Uses Text style.
    Inline(Vec<MathExpr>),

    /// Display math, standalone paragraph. Uses Display style.
    Display {
        content: Vec<MathExpr>,
        /// Horizontal alignment of the display equation within its paragraph.
        /// COM-verified: defaults to "center" if `<m:oMathParaPr>/<m:jc>` absent.
        jc: MathAlignment,
    },
}

/// Math paragraph alignment (§22.1.2.40 `jc`).
#[derive(Debug, Clone, Copy, Serialize, Deserialize, PartialEq, Eq)]
pub enum MathAlignment {
    /// Centered (default for `<m:oMathPara>`).
    Center,
    /// Left-aligned.
    Left,
    /// Right-aligned.
    Right,
    /// Center-group: each equation in a group centered at the widest equation.
    CenterGroup,
}

impl Default for MathAlignment {
    fn default() -> Self { MathAlignment::Center }
}

/// Recursive math expression tree. Covers all OMML primitives.
///
/// Ordered by implementation priority (Phase 3). Common constructs first;
/// edge cases last.
#[derive(Debug, Clone, Serialize, Deserialize)]
pub enum MathExpr {
    // ---- Leaf nodes ----

    /// Plain math text (from `<m:t>`). Substitution to italic-math variants
    /// happens at layout/render time, not during IR construction.
    Text(String),

    /// Math run with run-level style overrides (`<m:r>` with `<m:rPr>`).
    Run {
        text: String,
        style: MathRunStyle,
    },

    // ---- Structural primitives (ordered by frequency in academic docs) ----

    /// Fraction (§22.1.2.28 `f`). Numerator over denominator with bar.
    Fraction {
        num: Box<MathExpr>,
        den: Box<MathExpr>,
        bar_type: FracBarType,
    },

    /// Superscript (§22.1.2.108 `sSup`). `base^sup`.
    Superscript {
        base: Box<MathExpr>,
        sup: Box<MathExpr>,
    },

    /// Subscript (§22.1.2.105 `sSub`). `base_sub`.
    Subscript {
        base: Box<MathExpr>,
        sub: Box<MathExpr>,
    },

    /// Combined sub+superscript (§22.1.2.107 `sSubSup`). `base_sub^sup`.
    SubSuperscript {
        base: Box<MathExpr>,
        sub: Box<MathExpr>,
        sup: Box<MathExpr>,
    },

    /// Pre-sub and pre-superscript (§22.1.2.106 `sPre`). Like `^sup_sub base`
    /// (e.g., isotope notation `^14_6 C`).
    PreScript {
        base: Box<MathExpr>,
        sub: Box<MathExpr>,
        sup: Box<MathExpr>,
    },

    /// Radical (§22.1.2.101 `rad`). Square root (degree=None or empty) or
    /// nth root (degree=Some(expr)).
    Radical {
        /// None or empty = square root. Some(expr) = nth root.
        degree: Option<Box<MathExpr>>,
        radicand: Box<MathExpr>,
    },

    /// N-ary operator (§22.1.2.70 `nary`). Sum, integral, product, etc.
    /// with optional lower/upper limits.
    Nary {
        /// Operator character (e.g., '∑', '∫', '∏'). Parsed from `<m:chr>`.
        op: char,
        /// Lower limit (sub or under, depending on `lim_loc`).
        sub: Option<Box<MathExpr>>,
        /// Upper limit (sup or over).
        sup: Option<Box<MathExpr>>,
        /// Operand expression.
        operand: Box<MathExpr>,
        /// Position of limits: subSup (as scripts) or undOvr (over/under).
        lim_loc: LimLoc,
        /// Whether to force grown-variant glyph (`<m:grow>`).
        grow: bool,
    },

    /// Delimiter (§22.1.2.16 `d`). Bracketed content `(content)`, `[content]`, etc.
    Delimiter {
        beg: char,
        end: char,
        /// Optional separator for multi-element delimiter (like `|a|b|`).
        sep: Option<char>,
        content: Box<MathExpr>,
    },

    /// Function application (§22.1.2.35 `func`). Function name + argument.
    /// Used for `sin(x)`, `log x`, `lim_{x→0}`, etc.
    Function {
        name: Box<MathExpr>,
        arg: Box<MathExpr>,
    },

    /// Matrix (§22.1.2.61 `m`). 2D array of math cells.
    Matrix {
        /// Rows × columns (row-major).
        rows: Vec<Vec<MathExpr>>,
        /// Number of columns (redundant with rows[0].len() but explicit).
        cols: usize,
        /// Column alignment (per `<m:mcJc>` in `<m:mcPr>`).
        col_align: MathAlignment,
    },

    /// Accent above base (§22.1.2.1 `acc`). Hat, tilde, macron, vector, etc.
    Accent {
        /// Combining accent character (e.g., U+0302 for hat, U+20D7 for vector).
        accent: char,
        base: Box<MathExpr>,
    },

    /// Bar / overline / underline (§22.1.2.8 `bar`).
    Bar {
        pos: BarPos,
        base: Box<MathExpr>,
    },

    /// Lower limit (§22.1.2.53 `limLow`) or upper limit (§22.1.2.54 `limUpp`).
    /// Places the limit below (or above) the base (e.g., `lim_{x→0}`).
    Limit {
        base: Box<MathExpr>,
        lim: Box<MathExpr>,
        pos: LimitPos,
    },

    /// Group character (§22.1.2.38 `groupChr`). Underbrace, overbrace, etc.
    GroupChar {
        chr: char,
        pos: BarPos,
        base: Box<MathExpr>,
    },

    /// Equation array (§22.1.2.22 `eqArr`). Stacked equations.
    EqArray(Vec<MathExpr>),

    /// Boxed expression (§22.1.2.15 `box`). Visual grouping, no bar.
    BoxExpr(Box<MathExpr>),

    /// Bordered box (§22.1.2.11 `borderBox`). Rectangle around expression.
    BorderBox {
        base: Box<MathExpr>,
        /// Which sides of the border are drawn (top/bot/left/right).
        sides: BoxBorders,
    },

    /// Phantom (§22.1.2.85 `phant`). Reserves space without visible ink.
    Phantom(Box<MathExpr>),

    /// Sequence of expressions at the same level (used inside `<m:e>`).
    /// Acts like a concatenation: `a+b` is Seq([Run("a"), Run("+"), Run("b")]).
    Seq(Vec<MathExpr>),
}

/// Fraction bar styles (§22.1.2.30 `type` attribute).
#[derive(Debug, Clone, Copy, Serialize, Deserialize, PartialEq, Eq)]
pub enum FracBarType {
    /// Default: horizontal line between numerator and denominator.
    Bar,
    /// No bar (for binomial coefficients).
    NoBar,
    /// Linear inline: rendered as `num/den`.
    Linear,
    /// Skewed: diagonal line (seldom used).
    Skewed,
}

impl Default for FracBarType {
    fn default() -> Self { FracBarType::Bar }
}

/// Location of limits on n-ary operator (§22.1.2.52 `limLoc`).
#[derive(Debug, Clone, Copy, Serialize, Deserialize, PartialEq, Eq)]
pub enum LimLoc {
    /// Limits rendered as subscript/superscript (typical for inline math).
    SubSup,
    /// Limits rendered under/over the operator (typical for display math).
    UndOvr,
}

/// Bar position (§22.1.2.9 `pos` for bar, also used for groupChr).
#[derive(Debug, Clone, Copy, Serialize, Deserialize, PartialEq, Eq)]
pub enum BarPos {
    /// Above the base (overline).
    Top,
    /// Below the base (underline).
    Bot,
}

/// Limit position (for `<m:limLow>` vs `<m:limUpp>`).
#[derive(Debug, Clone, Copy, Serialize, Deserialize, PartialEq, Eq)]
pub enum LimitPos {
    /// Limit placed below the base (§22.1.2.53 `limLow`).
    Lower,
    /// Limit placed above the base (§22.1.2.54 `limUpp`).
    Upper,
}

/// Border sides for `<m:borderBox>`.
#[derive(Debug, Clone, Copy, Serialize, Deserialize, PartialEq, Eq, Default)]
pub struct BoxBorders {
    pub top: bool,
    pub bot: bool,
    pub left: bool,
    pub right: bool,
    /// Diagonal strikethrough.
    pub strikeh: bool,
    pub strikev: bool,
    pub strikebltr: bool,
    pub striketlbr: bool,
}

/// Math run style (for `<m:r>` with `<m:rPr>`). Math runs typically use
/// Cambria Math font with italic-math substitutions applied to Latin/Greek
/// letters; specific rPr attributes override or adjust.
#[derive(Debug, Clone, Default, Serialize, Deserialize)]
pub struct MathRunStyle {
    /// Script style: plain, double-struck, fraktur, sans-serif, monospace.
    /// From `<m:rPr>/<m:scr>` (§22.1.2.104).
    #[serde(default)]
    pub script: Option<MathScript>,

    /// Math style variation: plain, bold, italic, bold-italic.
    /// From `<m:rPr>/<m:sty>` (§22.1.2.116).
    #[serde(default)]
    pub math_style: Option<MathStyleVariant>,

    /// Whether to suppress italic-math substitution.
    /// From `<m:rPr>/<m:nor>` (§22.1.2.79), when true letters stay upright.
    #[serde(default)]
    pub literal: bool,

    /// Nested run style for font/color/size override. Maps to `<w:rPr>`
    /// embedded within `<m:rPr>`.
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub run_style: Option<RunStyle>,
}

/// Mathematical alphanumeric scripts (§22.1.2.104 `scr`).
#[derive(Debug, Clone, Copy, Serialize, Deserialize, PartialEq, Eq)]
pub enum MathScript {
    /// Roman (default). Italic applied to letters by default.
    Roman,
    /// Script (calligraphic) — U+1D49C range.
    Script,
    /// Fraktur — U+1D504 range.
    Fraktur,
    /// Double-struck (blackboard bold) — U+1D538 range.
    DoubleStruck,
    /// Sans-serif — U+1D5A0 range.
    SansSerif,
    /// Monospace — U+1D670 range.
    Monospace,
}

/// Math style variation (§22.1.2.116 `sty`).
#[derive(Debug, Clone, Copy, Serialize, Deserialize, PartialEq, Eq)]
pub enum MathStyleVariant {
    /// Plain upright (no italic, no bold).
    Plain,
    /// Bold (not italic).
    Bold,
    /// Italic (the default for math letters).
    Italic,
    /// Bold italic.
    BoldItalic,
}

/// Rendering style context (§22.1.3.2). Controls which MATH constants are
/// selected (inline vs display) and script-size cascade.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum MathStyle {
    /// Display-style: `<m:oMathPara>` top-level. Uses *DisplayStyle* MATH constants.
    Display,
    /// Text/inline-style: `<m:oMath>` in running text. Uses base MATH constants.
    Text,
    /// Script-level: subscripts, superscripts, limits. Scale = 73% (ScriptPercentScaleDown).
    Script,
    /// Script-script level: nested scripts. Scale = 60% (ScriptScriptPercentScaleDown).
    ScriptScript,
}

impl MathStyle {
    /// Font scale factor relative to base math font size.
    /// From Cambria Math: ScriptPercentScaleDown=73, ScriptScriptPercentScaleDown=60.
    pub fn scale_factor(&self) -> f32 {
        match self {
            MathStyle::Display | MathStyle::Text => 1.0,
            MathStyle::Script => 0.73,
            MathStyle::ScriptScript => 0.60,
        }
    }

    /// Descend into a script context (e.g., rendering a superscript).
    /// Script → ScriptScript → ScriptScript (doesn't descend further).
    pub fn script_style(&self) -> MathStyle {
        match self {
            MathStyle::Display | MathStyle::Text => MathStyle::Script,
            MathStyle::Script => MathStyle::ScriptScript,
            MathStyle::ScriptScript => MathStyle::ScriptScript,
        }
    }

    /// Whether this is display-style (selects DisplayStyle* constants).
    pub fn is_display(&self) -> bool {
        matches!(self, MathStyle::Display)
    }

    /// Initial style from MathBlock container.
    pub fn from_block(block: &MathBlock) -> MathStyle {
        match block {
            MathBlock::Inline(_) => MathStyle::Text,
            MathBlock::Display { .. } => MathStyle::Display,
        }
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn math_style_scale_factor() {
        assert_eq!(MathStyle::Display.scale_factor(), 1.0);
        assert_eq!(MathStyle::Text.scale_factor(), 1.0);
        assert_eq!(MathStyle::Script.scale_factor(), 0.73);
        assert_eq!(MathStyle::ScriptScript.scale_factor(), 0.60);
    }

    #[test]
    fn math_style_cascade() {
        let d = MathStyle::Display;
        let s1 = d.script_style();
        let s2 = s1.script_style();
        let s3 = s2.script_style();
        assert_eq!(s1, MathStyle::Script);
        assert_eq!(s2, MathStyle::ScriptScript);
        assert_eq!(s3, MathStyle::ScriptScript); // stays at ScriptScript
    }

    #[test]
    fn mathblock_initial_style() {
        let inline = MathBlock::Inline(vec![]);
        let display = MathBlock::Display { content: vec![], jc: MathAlignment::Center };
        assert_eq!(MathStyle::from_block(&inline), MathStyle::Text);
        assert_eq!(MathStyle::from_block(&display), MathStyle::Display);
    }

    #[test]
    fn default_alignment_is_center() {
        let a: MathAlignment = Default::default();
        assert_eq!(a, MathAlignment::Center);
    }

    #[test]
    fn math_expr_roundtrip_json() {
        // Construct a small expression (fraction a/b) and roundtrip via JSON.
        let expr = MathExpr::Fraction {
            num: Box::new(MathExpr::Text("a".to_string())),
            den: Box::new(MathExpr::Text("b".to_string())),
            bar_type: FracBarType::Bar,
        };
        let json = serde_json::to_string(&expr).expect("serialize");
        let back: MathExpr = serde_json::from_str(&json).expect("deserialize");
        match back {
            MathExpr::Fraction { num, den, bar_type } => {
                assert!(matches!(*num, MathExpr::Text(ref s) if s == "a"));
                assert!(matches!(*den, MathExpr::Text(ref s) if s == "b"));
                assert_eq!(bar_type, FracBarType::Bar);
            }
            _ => panic!("wrong variant"),
        }
    }
}
