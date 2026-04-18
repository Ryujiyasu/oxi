//! OMML math layout — bounding-box and position computation.
//!
//! This module defines the interface for Phase 3 math rendering. It
//! consumes a `MathBlock` tree and produces a bounding box + positioned
//! glyph list. Current state: leaf-only implementation (Text/Run) with
//! stubs for the recursive primitives.
//!
//! Layout flow:
//! 1. `layout_math_block(&block, font_size) -> MathLayout`
//! 2. For each `MathExpr` in the block:
//!    - Apply `math_substitute` to each character
//!    - Query `MathTable::cambria_math()` for MATH constants
//!    - Query `MathGlyphTables::cambria_math()` for per-glyph data
//!    - Recursively compose children's bboxes according to primitive rules
//! 3. Returns absolute positions + final bbox
//!
//! Coordinate convention: local to the math block's origin. Bbox `y=0`
//! is the math baseline. Positive y goes DOWN (matches Oxi overall).

use crate::font::{MathTable, MathGlyphTables, math_substitute};
use crate::ir::{MathBlock, MathExpr, MathStyle};

/// Bounding box for a math fragment. All values in points, relative to
/// a math baseline at y=0. Width extends rightward from origin x=0.
///
/// Think of it like a glyph metric: advance_width + above-baseline (asc)
/// + below-baseline (desc).
#[derive(Debug, Clone, Copy, Default, PartialEq)]
pub struct MathBBox {
    /// Horizontal advance (content width, including italic correction).
    pub advance: f32,
    /// Height above baseline (ascent) in points. Always ≥ 0.
    pub ascent: f32,
    /// Depth below baseline (descent) in points. Always ≥ 0.
    pub descent: f32,
    /// Italic correction in points (extra space before a superscript).
    pub italic_correction: f32,
}

impl MathBBox {
    /// Total vertical extent (ascent + descent).
    #[inline]
    pub fn height(&self) -> f32 { self.ascent + self.descent }

    /// Union two bboxes horizontally (side-by-side). Used for `Seq`.
    pub fn hstack(&self, rhs: &MathBBox) -> MathBBox {
        MathBBox {
            advance: self.advance + rhs.advance,
            ascent: self.ascent.max(rhs.ascent),
            descent: self.descent.max(rhs.descent),
            italic_correction: rhs.italic_correction, // last char's italic correction
        }
    }

    /// Stack two bboxes vertically (top on top). Used for fractions, stacks.
    /// `gap` is the inter-element gap in points.
    pub fn vstack(top: &MathBBox, bot: &MathBBox, gap: f32) -> MathBBox {
        MathBBox {
            advance: top.advance.max(bot.advance),
            ascent: top.height() + gap / 2.0,
            descent: bot.height() + gap / 2.0,
            italic_correction: 0.0, // vertical stacks don't carry italic correction
        }
    }
}

/// Layout context: font size + math style (for constant selection).
#[derive(Debug, Clone, Copy)]
pub struct MathLayoutContext {
    pub font_size: f32,
    pub style: MathStyle,
}

impl MathLayoutContext {
    /// Effective font size at this style level.
    pub fn effective_font_size(&self) -> f32 {
        self.font_size * self.style.scale_factor()
    }

    /// Descend into script style (sub/sup).
    pub fn descend_script(&self) -> MathLayoutContext {
        MathLayoutContext {
            font_size: self.font_size,
            style: self.style.script_style(),
        }
    }
}

/// Estimated bbox for a single character in Cambria Math at the given
/// effective font size.
///
/// Uses a simple heuristic: width = fontSize × 0.5 (math italic letters
/// average ~0.5em wide); ascent/descent approximate 0.7 / 0.2 em.
/// Refined in Phase 3 with actual Cambria Math horizontal advance tables.
pub fn leaf_char_bbox(c: char, ctx: &MathLayoutContext) -> MathBBox {
    let eff = ctx.effective_font_size();
    let sub = math_substitute(c);
    let tables = MathGlyphTables::cambria_math();
    let table = MathTable::cambria_math();
    let italic_corr = tables.italic_correction(sub)
        .map(|du| table.du_to_pt(du, eff))
        .unwrap_or(0.0);
    MathBBox {
        advance: eff * 0.5,
        ascent: eff * 0.7,
        descent: eff * 0.2,
        italic_correction: italic_corr,
    }
}

/// Bounding box for a leaf Text/Run (concatenation of chars).
pub fn leaf_text_bbox(text: &str, ctx: &MathLayoutContext) -> MathBBox {
    let mut acc = MathBBox::default();
    for c in text.chars() {
        let b = leaf_char_bbox(c, ctx);
        acc = acc.hstack(&b);
    }
    acc
}

/// Top-level: compute the bbox for a full MathBlock.
///
/// In Phase 3 this will also emit positioned glyph lists; currently
/// returns only the bbox for leaf Text/Run content. Non-leaf primitives
/// return a zero bbox (their recursive layout is TODO for Phase 3).
pub fn layout_math_block(block: &MathBlock, font_size: f32) -> MathBBox {
    let ctx = MathLayoutContext {
        font_size,
        style: MathStyle::from_block(block),
    };
    let exprs: &[MathExpr] = match block {
        MathBlock::Inline(xs) => xs,
        MathBlock::Display { content, .. } => content,
    };
    let mut acc = MathBBox::default();
    for e in exprs {
        let b = layout_expr(e, &ctx);
        acc = acc.hstack(&b);
    }
    acc
}

/// Dispatch bbox computation by MathExpr variant. Phase 2 implements
/// only leaf cases; Phase 3 adds the full primitive set.
pub fn layout_expr(expr: &MathExpr, ctx: &MathLayoutContext) -> MathBBox {
    match expr {
        MathExpr::Text(s) => leaf_text_bbox(s, ctx),
        MathExpr::Run { text, .. } => leaf_text_bbox(text, ctx),
        MathExpr::Seq(children) => {
            let mut acc = MathBBox::default();
            for c in children {
                acc = acc.hstack(&layout_expr(c, ctx));
            }
            acc
        }
        // Phase 3: full recursive layout for these primitives.
        MathExpr::Fraction { num, den, .. } => {
            let nb = layout_expr(num, &ctx.descend_script());
            let db = layout_expr(den, &ctx.descend_script());
            let table = MathTable::cambria_math();
            let gap = table.du_to_pt(table.constants.FractionRuleThickness, ctx.font_size);
            MathBBox::vstack(&nb, &db, gap)
        }
        MathExpr::Superscript { base, sup } => {
            let bb = layout_expr(base, ctx);
            let sb = layout_expr(sup, &ctx.descend_script());
            let table = MathTable::cambria_math();
            let shift_up = table.du_to_pt(table.constants.SuperscriptShiftUp, ctx.font_size);
            MathBBox {
                advance: bb.advance + bb.italic_correction + sb.advance,
                ascent: bb.ascent.max(sb.height() + shift_up),
                descent: bb.descent,
                italic_correction: sb.italic_correction,
            }
        }
        MathExpr::Subscript { base, sub } => {
            let bb = layout_expr(base, ctx);
            let sb = layout_expr(sub, &ctx.descend_script());
            let table = MathTable::cambria_math();
            let shift_down = table.du_to_pt(table.constants.SubscriptShiftDown, ctx.font_size);
            MathBBox {
                advance: bb.advance + sb.advance,
                ascent: bb.ascent,
                descent: bb.descent.max(sb.height() + shift_down),
                italic_correction: sb.italic_correction,
            }
        }
        MathExpr::SubSuperscript { base, sub, sup } => {
            let bb = layout_expr(base, ctx);
            let super_b = layout_expr(sup, &ctx.descend_script());
            let sub_b = layout_expr(sub, &ctx.descend_script());
            let table = MathTable::cambria_math();
            let sup_shift = table.du_to_pt(table.constants.SuperscriptShiftUp, ctx.font_size);
            let sub_shift = table.du_to_pt(table.constants.SubscriptShiftDown, ctx.font_size);
            MathBBox {
                advance: bb.advance + bb.italic_correction
                    + super_b.advance.max(sub_b.advance),
                ascent: bb.ascent.max(super_b.height() + sup_shift),
                descent: bb.descent.max(sub_b.height() + sub_shift),
                italic_correction: 0.0,
            }
        }
        MathExpr::Radical { radicand, .. } => {
            let rb = layout_expr(radicand, ctx);
            let table = MathTable::cambria_math();
            let gap = table.du_to_pt(table.constants.RadicalVerticalGap, ctx.font_size);
            let thk = table.du_to_pt(table.constants.RadicalRuleThickness, ctx.font_size);
            MathBBox {
                advance: rb.advance + ctx.font_size * 0.5, // radical sign width estimate
                ascent: rb.ascent + gap + thk,
                descent: rb.descent,
                italic_correction: 0.0,
            }
        }
        // Primitives not yet implemented — return zero bbox.
        // Phase 3 fills in Nary / Matrix / Delimiter / Accent / etc.
        _ => MathBBox::default(),
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::ir::{MathAlignment, FracBarType};

    #[test]
    fn leaf_char_bbox_has_italic_correction_for_integral() {
        let ctx = MathLayoutContext { font_size: 10.5, style: MathStyle::Text };
        // ∫ has italic correction 415 DU in Cambria Math
        let b = leaf_char_bbox('∫', &ctx);
        // 415 * 10.5 / 2048 ≈ 2.13 pt
        assert!(b.italic_correction > 2.0 && b.italic_correction < 2.3,
                "got {}", b.italic_correction);
    }

    #[test]
    fn text_bbox_accumulates_advance() {
        let ctx = MathLayoutContext { font_size: 10.5, style: MathStyle::Text };
        let b_one = leaf_text_bbox("x", &ctx);
        let b_three = leaf_text_bbox("xxx", &ctx);
        // Three chars should have ~3× the advance of one
        assert!((b_three.advance - 3.0 * b_one.advance).abs() < 0.01);
    }

    #[test]
    fn empty_inline_block_is_zero_bbox() {
        let block = MathBlock::Inline(vec![]);
        let b = layout_math_block(&block, 10.5);
        assert_eq!(b, MathBBox::default());
    }

    #[test]
    fn display_style_is_selected_for_display_block() {
        let block = MathBlock::Display {
            content: vec![MathExpr::Text("a".to_string())],
            jc: MathAlignment::Center,
        };
        let b = layout_math_block(&block, 12.0);
        assert!(b.advance > 0.0);
    }

    #[test]
    fn fraction_bbox_stacks_vertically() {
        let frac = MathExpr::Fraction {
            num: Box::new(MathExpr::Text("a".to_string())),
            den: Box::new(MathExpr::Text("b".to_string())),
            bar_type: FracBarType::Bar,
        };
        let ctx = MathLayoutContext { font_size: 10.5, style: MathStyle::Text };
        let b = layout_expr(&frac, &ctx);
        // Height should be larger than either child alone
        let a_only = leaf_char_bbox('a', &ctx.descend_script());
        assert!(b.height() > a_only.height() * 1.5);
    }

    #[test]
    fn superscript_ascent_grows() {
        // x^2: base ascent + superscript lifted above
        let sup = MathExpr::Superscript {
            base: Box::new(MathExpr::Text("x".to_string())),
            sup: Box::new(MathExpr::Text("2".to_string())),
        };
        let ctx = MathLayoutContext { font_size: 10.5, style: MathStyle::Text };
        let b = layout_expr(&sup, &ctx);
        let x_alone = leaf_char_bbox('x', &ctx);
        assert!(b.ascent > x_alone.ascent);
    }

    #[test]
    fn script_context_scales_down() {
        let ctx = MathLayoutContext { font_size: 10.5, style: MathStyle::Text };
        let ctx_s = ctx.descend_script();
        assert!((ctx_s.effective_font_size() - 10.5 * 0.73).abs() < 0.01);
        let ctx_ss = ctx_s.descend_script();
        assert!((ctx_ss.effective_font_size() - 10.5 * 0.60).abs() < 0.01);
    }

    #[test]
    fn bbox_hstack_accumulates() {
        let a = MathBBox { advance: 5.0, ascent: 7.0, descent: 2.0, italic_correction: 0.5 };
        let b = MathBBox { advance: 3.0, ascent: 6.0, descent: 3.0, italic_correction: 0.0 };
        let u = a.hstack(&b);
        assert_eq!(u.advance, 8.0);
        assert_eq!(u.ascent, 7.0);   // max
        assert_eq!(u.descent, 3.0);  // max
        assert_eq!(u.italic_correction, 0.0); // rhs's
    }
}
