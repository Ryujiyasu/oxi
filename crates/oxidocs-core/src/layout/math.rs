//! OMML math layout — bounding-box and position computation, plus
//! Phase 3 MVP: emit flat `LayoutElement::Text` entries for the GDI
//! renderer to draw Cambria Math glyphs.
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
use crate::layout::{LayoutElement, LayoutContent};

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
            let sub_ctx = if ctx.style.is_display() { *ctx } else { ctx.descend_script() };
            let nb = layout_expr(num, &sub_ctx);
            let db = layout_expr(den, &sub_ctx);
            let table = MathTable::cambria_math();
            let fs = ctx.font_size;
            let (num_shift_du, den_shift_du) = if ctx.style.is_display() {
                (table.constants.FractionNumeratorDisplayStyleShiftUp,
                 table.constants.FractionDenominatorDisplayStyleShiftDown)
            } else {
                (table.constants.FractionNumeratorShiftUp,
                 table.constants.FractionDenominatorShiftDown)
            };
            let num_shift_up = table.du_to_pt(num_shift_du, fs);
            let den_shift_down = table.du_to_pt(den_shift_du, fs);
            MathBBox {
                advance: nb.advance.max(db.advance),
                ascent: num_shift_up + nb.ascent,
                descent: den_shift_down + db.descent,
                italic_correction: 0.0,
            }
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
            let fs = ctx.font_size;
            let gap_du = if ctx.style.is_display() {
                table.constants.RadicalDisplayStyleVerticalGap
            } else {
                table.constants.RadicalVerticalGap
            };
            let gap = table.du_to_pt(gap_du, fs);
            let thk = table.du_to_pt(table.constants.RadicalRuleThickness, fs);
            let extra = table.du_to_pt(table.constants.RadicalExtraAscender, fs);
            MathBBox {
                advance: rb.advance + fs * 0.6,
                ascent: rb.ascent + gap + thk + extra,
                descent: rb.descent,
                italic_correction: 0.0,
            }
        }
        MathExpr::Delimiter { content, .. } => {
            let cb = layout_expr(content, ctx);
            let fs = ctx.font_size;
            let delim_w = fs * 0.45;
            MathBBox {
                advance: cb.advance + 2.0 * delim_w,
                ascent: cb.ascent.max(fs * 0.8),
                descent: cb.descent.max(fs * 0.2),
                italic_correction: 0.0,
            }
        }
        MathExpr::Bar { pos, base } => {
            let bb = layout_expr(base, ctx);
            let table = MathTable::cambria_math();
            let fs = ctx.font_size;
            let (gap, thick, extra) = match pos {
                crate::ir::BarPos::Top => (
                    table.du_to_pt(table.constants.OverbarVerticalGap, fs),
                    table.du_to_pt(table.constants.OverbarRuleThickness, fs),
                    table.du_to_pt(table.constants.OverbarExtraAscender, fs),
                ),
                crate::ir::BarPos::Bot => (
                    table.du_to_pt(table.constants.UnderbarVerticalGap, fs),
                    table.du_to_pt(table.constants.UnderbarRuleThickness, fs),
                    table.du_to_pt(table.constants.UnderbarExtraDescender, fs),
                ),
            };
            let mut bbox = bb;
            match pos {
                crate::ir::BarPos::Top => bbox.ascent += gap + thick + extra,
                crate::ir::BarPos::Bot => bbox.descent += gap + thick + extra,
            }
            bbox
        }
        MathExpr::Accent { base, .. } => {
            let bb = layout_expr(base, ctx);
            let table = MathTable::cambria_math();
            let fs = ctx.font_size;
            let gap = table.du_to_pt(table.constants.OverbarVerticalGap, fs);
            MathBBox {
                advance: bb.advance,
                ascent: bb.ascent + gap + fs * 0.9,
                descent: bb.descent,
                italic_correction: 0.0,
            }
        }
        MathExpr::Limit { base, lim, pos } => {
            let bb = layout_expr(base, ctx);
            let lim_ctx = ctx.descend_script();
            let lb = layout_expr(lim, &lim_ctx);
            let table = MathTable::cambria_math();
            let fs = ctx.font_size;
            let common = bb.advance.max(lb.advance);
            match pos {
                crate::ir::LimitPos::Lower => {
                    let gap = table.du_to_pt(table.constants.LowerLimitGapMin, fs);
                    let drop = table.du_to_pt(table.constants.LowerLimitBaselineDropMin, fs);
                    MathBBox {
                        advance: common,
                        ascent: bb.ascent,
                        descent: bb.descent + gap + drop + lb.ascent + lb.descent,
                        italic_correction: 0.0,
                    }
                }
                crate::ir::LimitPos::Upper => {
                    let gap = table.du_to_pt(table.constants.UpperLimitGapMin, fs);
                    let rise = table.du_to_pt(table.constants.UpperLimitBaselineRiseMin, fs);
                    MathBBox {
                        advance: common,
                        ascent: bb.ascent + gap + rise + lb.ascent + lb.descent,
                        descent: bb.descent,
                        italic_correction: 0.0,
                    }
                }
            }
        }
        MathExpr::Matrix { rows, .. } => {
            if rows.is_empty() { return MathBBox::default(); }
            let n_cols = rows.iter().map(|r| r.len()).max().unwrap_or(0);
            let mut col_widths = vec![0.0_f32; n_cols];
            let mut row_heights = vec![0.0_f32; rows.len()];
            for (i, row) in rows.iter().enumerate() {
                for (j, cell) in row.iter().enumerate() {
                    let bb = layout_expr(cell, ctx);
                    if bb.advance > col_widths[j] { col_widths[j] = bb.advance; }
                    let h = bb.ascent + bb.descent;
                    if h > row_heights[i] { row_heights[i] = h; }
                }
            }
            let table = MathTable::cambria_math();
            let fs = ctx.font_size;
            let gap = table.du_to_pt(table.constants.MathLeading, fs);
            let axis_h = table.du_to_pt(table.constants.AxisHeight, fs);
            let total_h: f32 = row_heights.iter().sum::<f32>()
                + gap * rows.len().saturating_sub(1) as f32;
            let total_w: f32 = col_widths.iter().sum::<f32>()
                + gap * n_cols.saturating_sub(1) as f32;
            MathBBox {
                advance: total_w,
                ascent: total_h / 2.0 + axis_h,
                descent: total_h / 2.0 - axis_h,
                italic_correction: 0.0,
            }
        }
        MathExpr::Nary { sub, sup, operand, .. } => {
            let fs = ctx.font_size;
            let op_size = if ctx.style.is_display() { fs * 1.6 } else { fs * 1.2 };
            let op_w = op_size * 0.6;
            let lim_ctx = ctx.descend_script();
            let sub_b = sub.as_ref().map(|s| layout_expr(s, &lim_ctx));
            let sup_b = sup.as_ref().map(|s| layout_expr(s, &lim_ctx));
            let op_bbox = layout_expr(operand, ctx);
            let limits_w = op_w
                .max(sub_b.as_ref().map(|b| b.advance).unwrap_or(0.0))
                .max(sup_b.as_ref().map(|b| b.advance).unwrap_or(0.0));
            MathBBox {
                advance: limits_w + fs * 0.1 + op_bbox.advance,
                ascent: (op_size * 0.8)
                    .max(sup_b.as_ref().map(|b| op_size + b.height()).unwrap_or(0.0))
                    .max(op_bbox.ascent),
                descent: (op_size * 0.2)
                    .max(sub_b.as_ref().map(|b| op_size + b.height()).unwrap_or(0.0))
                    .max(op_bbox.descent),
                italic_correction: 0.0,
            }
        }
        MathExpr::Function { name, arg } => {
            let nb = layout_expr(name, ctx);
            let ab = layout_expr(arg, ctx);
            let gap = ctx.font_size * 0.15;
            MathBBox {
                advance: nb.advance + gap + ab.advance,
                ascent: nb.ascent.max(ab.ascent),
                descent: nb.descent.max(ab.descent),
                italic_correction: ab.italic_correction,
            }
        }
        MathExpr::GroupChar { pos, base, .. } => {
            let bb = layout_expr(base, ctx);
            let table = MathTable::cambria_math();
            let fs = ctx.font_size;
            let gap = table.du_to_pt(table.constants.StretchStackGapAboveMin, fs);
            let chr_h = fs * 0.8;
            let mut bbox = bb;
            match pos {
                crate::ir::BarPos::Top => bbox.ascent += gap + chr_h,
                crate::ir::BarPos::Bot => bbox.descent += gap + chr_h,
            }
            bbox
        }
        MathExpr::EqArray(items) => {
            if items.is_empty() { return MathBBox::default(); }
            let bbs: Vec<MathBBox> = items.iter().map(|e| layout_expr(e, ctx)).collect();
            let table = MathTable::cambria_math();
            let fs = ctx.font_size;
            let gap = table.du_to_pt(table.constants.StackGapMin, fs);
            let total_h: f32 = bbs.iter().map(|b| b.height()).sum::<f32>()
                + gap * items.len().saturating_sub(1) as f32;
            let axis = table.du_to_pt(table.constants.AxisHeight, fs);
            let common_w = bbs.iter().map(|b| b.advance).fold(0.0_f32, f32::max);
            MathBBox {
                advance: common_w,
                ascent: total_h / 2.0 + axis,
                descent: total_h / 2.0 - axis,
                italic_correction: 0.0,
            }
        }
        MathExpr::PreScript { base, sub, sup } => {
            let s_ctx = ctx.descend_script();
            let bb = layout_expr(base, ctx);
            let sb = layout_expr(sub, &s_ctx);
            let pb = layout_expr(sup, &s_ctx);
            let pre_w = sb.advance.max(pb.advance);
            let table = MathTable::cambria_math();
            let fs = ctx.font_size;
            let sup_shift = table.du_to_pt(table.constants.SuperscriptShiftUp, fs);
            let sub_shift = table.du_to_pt(table.constants.SubscriptShiftDown, fs);
            MathBBox {
                advance: pre_w + bb.advance,
                ascent: bb.ascent.max(sup_shift + pb.ascent),
                descent: bb.descent.max(sub_shift + sb.descent),
                italic_correction: bb.italic_correction,
            }
        }
        // Primitives not yet implemented — return zero bbox.
        _ => MathBBox::default(),
    }
}

/// Extract all text content from a MathExpr tree, applying
/// `math_substitute` to each character. Returns a flat string suitable
/// for Phase 3 MVP rendering as a single line via `LayoutElement::Text`.
///
/// Structural chars are inserted for fractions ("/"), radicals ("√"),
/// delimiters (their `beg`/`end` chars), etc., to give human-readable
/// approximation. Proper stacked layout comes in later Phase 3 commits.
pub fn extract_flat_text(expr: &MathExpr) -> String {
    let mut out = String::new();
    append_flat(&mut out, expr);
    out
}

fn append_flat(out: &mut String, expr: &MathExpr) {
    match expr {
        MathExpr::Text(s) | MathExpr::Run { text: s, .. } => {
            for c in s.chars() { out.push(math_substitute(c)); }
        }
        MathExpr::Seq(children) => {
            for c in children { append_flat(out, c); }
        }
        MathExpr::Fraction { num, den, bar_type } => {
            use crate::ir::FracBarType;
            match bar_type {
                FracBarType::NoBar => {
                    append_flat(out, num);
                    out.push(' ');
                    append_flat(out, den);
                }
                FracBarType::Linear => {
                    append_flat(out, num);
                    out.push('/');
                    append_flat(out, den);
                }
                _ => {
                    append_flat(out, num);
                    out.push('/');
                    append_flat(out, den);
                }
            }
        }
        MathExpr::Superscript { base, sup } => {
            append_flat(out, base);
            out.push('^');
            append_flat(out, sup);
        }
        MathExpr::Subscript { base, sub } => {
            append_flat(out, base);
            out.push('_');
            append_flat(out, sub);
        }
        MathExpr::SubSuperscript { base, sub, sup } => {
            append_flat(out, base);
            out.push('_');
            append_flat(out, sub);
            out.push('^');
            append_flat(out, sup);
        }
        MathExpr::PreScript { base, sub, sup } => {
            out.push('_');
            append_flat(out, sub);
            out.push('^');
            append_flat(out, sup);
            append_flat(out, base);
        }
        MathExpr::Radical { degree, radicand } => {
            if let Some(d) = degree {
                out.push('^');
                append_flat(out, d);
            }
            out.push('√');
            append_flat(out, radicand);
        }
        MathExpr::Nary { op, sub, sup, operand, .. } => {
            out.push(*op);
            if let Some(s) = sub { out.push('_'); append_flat(out, s); }
            if let Some(s) = sup { out.push('^'); append_flat(out, s); }
            out.push(' ');
            append_flat(out, operand);
        }
        MathExpr::Delimiter { beg, end, content, .. } => {
            out.push(*beg);
            append_flat(out, content);
            out.push(*end);
        }
        MathExpr::Function { name, arg } => {
            append_flat(out, name);
            out.push(' ');
            append_flat(out, arg);
        }
        MathExpr::Matrix { rows, .. } => {
            // Matrix itself has no brackets; Delimiter wraps it when needed.
            for (i, row) in rows.iter().enumerate() {
                if i > 0 { out.push(';'); out.push(' '); }
                for (j, cell) in row.iter().enumerate() {
                    if j > 0 { out.push(' '); }
                    append_flat(out, cell);
                }
            }
        }
        MathExpr::Accent { accent, base } => {
            append_flat(out, base);
            out.push(*accent);
        }
        MathExpr::Bar { base, .. } => {
            out.push('‾');
            append_flat(out, base);
        }
        MathExpr::Limit { base, lim, pos } => {
            use crate::ir::LimitPos;
            append_flat(out, base);
            match pos {
                LimitPos::Lower => out.push('_'),
                LimitPos::Upper => out.push('^'),
            }
            append_flat(out, lim);
        }
        MathExpr::GroupChar { chr, base, .. } => {
            append_flat(out, base);
            out.push(*chr);
        }
        MathExpr::EqArray(children) => {
            for (i, c) in children.iter().enumerate() {
                if i > 0 { out.push_str("; "); }
                append_flat(out, c);
            }
        }
        MathExpr::BoxExpr(inner) | MathExpr::Phantom(inner) => {
            append_flat(out, inner);
        }
        MathExpr::BorderBox { base, .. } => {
            append_flat(out, base);
        }
    }
}

// ============================================================================
// Phase 3: Positioned LayoutElement emission
// ============================================================================

/// Emit a `LayoutElement::Text` for substituted characters at a given
/// baseline y. `x` is the left edge. Uses Cambria Math font.
fn emit_text_at(
    text: String,
    x: f32,
    baseline_y: f32,
    font_size: f32,
) -> LayoutElement {
    // Text element y in oxi convention = top of line-box, not baseline.
    // Approximation: top = baseline - ascent, where ascent ≈ 0.8 × font_size.
    let ascent_approx = font_size * 0.8;
    let top = baseline_y - ascent_approx;
    let char_count = text.chars().count() as f32;
    let approx_width = char_count * font_size * 0.55;
    LayoutElement::new(
        x,
        top,
        approx_width,
        font_size * 1.2,
        LayoutContent::Text {
            text,
            font_size,
            font_family: Some("Cambria Math".to_string()),
            bold: false,
            italic: false,
            underline: false,
            underline_style: None,
            strikethrough: false,
            double_strikethrough: false,
            color: None,
            highlight: None,
            field_type: None,
            character_spacing: 0.0,
            text_scale: 100.0,
            is_vertical: false,
        },
    )
}

/// Emit positioned LayoutElements for a single math expression.
/// Returns (elements, bbox). All elements use absolute page coordinates.
///
/// Origin: element is rendered with its LEFT edge at `x` and its BASELINE
/// at `baseline_y`. Bbox returned describes the rendered content size.
pub fn emit_expr(
    expr: &MathExpr,
    x: f32,
    baseline_y: f32,
    ctx: &MathLayoutContext,
) -> (Vec<LayoutElement>, MathBBox) {
    let eff_size = ctx.effective_font_size();
    match expr {
        MathExpr::Text(s) | MathExpr::Run { text: s, .. } => {
            if s.is_empty() {
                return (vec![], MathBBox::default());
            }
            // Apply italic-math substitution per-char.
            let subbed: String = s.chars().map(math_substitute).collect();
            let bbox = leaf_text_bbox(s, ctx);
            let el = emit_text_at(subbed, x, baseline_y, eff_size);
            (vec![el], bbox)
        }
        MathExpr::Seq(children) => {
            let mut elems = Vec::new();
            let mut cur_x = x;
            let mut total = MathBBox::default();
            for child in children {
                let (e, b) = emit_expr(child, cur_x, baseline_y, ctx);
                elems.extend(e);
                cur_x += b.advance;
                total = total.hstack(&b);
            }
            (elems, total)
        }
        MathExpr::Fraction { num, den, bar_type } => {
            emit_fraction(num, den, *bar_type, x, baseline_y, ctx)
        }
        MathExpr::Superscript { base, sup } => {
            emit_superscript(base, sup, x, baseline_y, ctx)
        }
        MathExpr::Subscript { base, sub } => {
            emit_subscript(base, sub, x, baseline_y, ctx)
        }
        MathExpr::SubSuperscript { base, sub, sup } => {
            emit_subsuperscript(base, sub, sup, x, baseline_y, ctx)
        }
        MathExpr::Radical { degree, radicand } => {
            emit_radical(degree.as_deref(), radicand, x, baseline_y, ctx)
        }
        MathExpr::Matrix { rows, col_align, .. } => {
            emit_matrix(rows, *col_align, x, baseline_y, ctx)
        }
        MathExpr::Delimiter { beg, end, content, .. } => {
            emit_delimiter(*beg, *end, content, x, baseline_y, ctx)
        }
        MathExpr::Bar { pos, base } => {
            emit_bar(*pos, base, x, baseline_y, ctx)
        }
        MathExpr::Accent { accent, base } => {
            emit_accent(*accent, base, x, baseline_y, ctx)
        }
        MathExpr::Limit { base, lim, pos } => {
            emit_limit(base, lim, *pos, x, baseline_y, ctx)
        }
        MathExpr::Nary { op, sub, sup, operand, lim_loc, .. } => {
            emit_nary(*op, sub.as_deref(), sup.as_deref(), operand, *lim_loc, x, baseline_y, ctx)
        }
        MathExpr::Function { name, arg } => {
            emit_function(name, arg, x, baseline_y, ctx)
        }
        MathExpr::GroupChar { chr, pos, base } => {
            emit_group_chr(*chr, *pos, base, x, baseline_y, ctx)
        }
        MathExpr::EqArray(items) => {
            emit_eq_array(items, x, baseline_y, ctx)
        }
        MathExpr::PreScript { base, sub, sup } => {
            emit_prescript(base, sub, sup, x, baseline_y, ctx)
        }
        // Other primitives: fall back to flat text via extract_flat_text.
        _ => {
            let flat = extract_flat_text(expr);
            if flat.is_empty() {
                return (vec![], MathBBox::default());
            }
            let bbox = layout_expr(expr, ctx);
            let el = emit_text_at(flat, x, baseline_y, eff_size);
            (vec![el], bbox)
        }
    }
}

/// Emit fraction with num above bar, den below bar. Bar drawn as TableBorder.
fn emit_fraction(
    num: &MathExpr,
    den: &MathExpr,
    bar_type: crate::ir::FracBarType,
    x: f32,
    baseline_y: f32,
    ctx: &MathLayoutContext,
) -> (Vec<LayoutElement>, MathBBox) {
    use crate::ir::FracBarType;
    let table = MathTable::cambria_math();
    let fs = ctx.font_size;

    // Scale sub-expressions at script style if this is an inline fraction.
    // (Display style keeps parent size for num/den.)
    let sub_ctx = if ctx.style.is_display() { *ctx } else { ctx.descend_script() };

    // Compute num and den bboxes without emission first.
    let num_bbox = layout_expr(num, &sub_ctx);
    let den_bbox = layout_expr(den, &sub_ctx);

    // Fraction dimensions from MATH constants.
    let (num_shift_du, den_shift_du, rule_thick_du) = if ctx.style.is_display() {
        (
            table.constants.FractionNumeratorDisplayStyleShiftUp,
            table.constants.FractionDenominatorDisplayStyleShiftDown,
            table.constants.FractionRuleThickness,
        )
    } else {
        (
            table.constants.FractionNumeratorShiftUp,
            table.constants.FractionDenominatorShiftDown,
            table.constants.FractionRuleThickness,
        )
    };
    let num_shift_up = table.du_to_pt(num_shift_du, fs);
    let den_shift_down = table.du_to_pt(den_shift_du, fs);
    let rule_thick = table.du_to_pt(rule_thick_du, fs);
    let axis_height = table.du_to_pt(table.constants.AxisHeight, fs);

    // Common width: max of num and den advances.
    let common_w = num_bbox.advance.max(den_bbox.advance);
    let num_x = x + (common_w - num_bbox.advance) / 2.0;
    let den_x = x + (common_w - den_bbox.advance) / 2.0;

    // Num baseline: above baseline_y by num_shift_up.
    let num_baseline = baseline_y - num_shift_up;
    // Den baseline: below baseline_y by den_shift_down.
    let den_baseline = baseline_y + den_shift_down;
    // Bar center y: at math axis (baseline_y - axis_height).
    let bar_y = baseline_y - axis_height;

    let mut elems = Vec::new();
    let (ne, _nb) = emit_expr(num, num_x, num_baseline, &sub_ctx);
    let (de, _db) = emit_expr(den, den_x, den_baseline, &sub_ctx);
    elems.extend(ne);
    elems.extend(de);

    // Emit the fraction bar as TableBorder (horizontal line) unless NoBar/Skewed.
    if !matches!(bar_type, FracBarType::NoBar) && !matches!(bar_type, FracBarType::Linear) {
        elems.push(LayoutElement::new(
            x, bar_y - rule_thick / 2.0,
            common_w, rule_thick,
            LayoutContent::TableBorder {
                x1: x,
                y1: bar_y,
                x2: x + common_w,
                y2: bar_y,
                color: None,
                width: rule_thick,
                style: None,
            },
        ));
    }

    let bbox = MathBBox {
        advance: common_w,
        ascent: num_shift_up + num_bbox.ascent,
        descent: den_shift_down + den_bbox.descent,
        italic_correction: 0.0,
    };
    (elems, bbox)
}

/// Emit superscript: base followed by raised sup at script size.
fn emit_superscript(
    base: &MathExpr,
    sup: &MathExpr,
    x: f32,
    baseline_y: f32,
    ctx: &MathLayoutContext,
) -> (Vec<LayoutElement>, MathBBox) {
    let table = MathTable::cambria_math();
    let fs = ctx.font_size;
    let (mut base_elems, base_bbox) = emit_expr(base, x, baseline_y, ctx);

    let sup_ctx = ctx.descend_script();
    let shift_up = table.du_to_pt(table.constants.SuperscriptShiftUp, fs);
    let sup_x = x + base_bbox.advance + base_bbox.italic_correction;
    let sup_baseline = baseline_y - shift_up;
    let (sup_elems, sup_bbox) = emit_expr(sup, sup_x, sup_baseline, &sup_ctx);

    base_elems.extend(sup_elems);
    let bbox = MathBBox {
        advance: base_bbox.advance + base_bbox.italic_correction + sup_bbox.advance,
        ascent: base_bbox.ascent.max(shift_up + sup_bbox.ascent),
        descent: base_bbox.descent,
        italic_correction: sup_bbox.italic_correction,
    };
    (base_elems, bbox)
}

/// Emit subscript: base followed by lowered sub at script size.
fn emit_subscript(
    base: &MathExpr,
    sub: &MathExpr,
    x: f32,
    baseline_y: f32,
    ctx: &MathLayoutContext,
) -> (Vec<LayoutElement>, MathBBox) {
    let table = MathTable::cambria_math();
    let fs = ctx.font_size;
    let (mut base_elems, base_bbox) = emit_expr(base, x, baseline_y, ctx);

    let sub_ctx = ctx.descend_script();
    let shift_down = table.du_to_pt(table.constants.SubscriptShiftDown, fs);
    let sub_x = x + base_bbox.advance;
    let sub_baseline = baseline_y + shift_down;
    let (sub_elems, sub_bbox) = emit_expr(sub, sub_x, sub_baseline, &sub_ctx);

    base_elems.extend(sub_elems);
    let bbox = MathBBox {
        advance: base_bbox.advance + sub_bbox.advance,
        ascent: base_bbox.ascent,
        descent: base_bbox.descent.max(shift_down + sub_bbox.descent),
        italic_correction: sub_bbox.italic_correction,
    };
    (base_elems, bbox)
}

/// Emit radical: √ sign + overline over radicand. Optional degree for nth-root.
fn emit_radical(
    degree: Option<&MathExpr>,
    radicand: &MathExpr,
    x: f32,
    baseline_y: f32,
    ctx: &MathLayoutContext,
) -> (Vec<LayoutElement>, MathBBox) {
    let table = MathTable::cambria_math();
    let fs = ctx.font_size;

    // Radicand bbox (at the same style — not script).
    let rad_bbox = layout_expr(radicand, ctx);

    // MATH constants (select display vs inline gap).
    let v_gap_du = if ctx.style.is_display() {
        table.constants.RadicalDisplayStyleVerticalGap
    } else {
        table.constants.RadicalVerticalGap
    };
    let v_gap = table.du_to_pt(v_gap_du, fs);
    let rule_thick = table.du_to_pt(table.constants.RadicalRuleThickness, fs);
    let extra_asc = table.du_to_pt(table.constants.RadicalExtraAscender, fs);

    // Radical sign "√" (U+221A) rendered as text. Its ascent + descent should
    // cover the radicand's full height + vertical gap + rule thickness.
    let sign_width = fs * 0.6;  // approximate √ glyph advance

    // Radicand inner left edge (after √ sign).
    let radicand_x = x + sign_width;

    // Overbar y: above radicand top, gap above.
    let radicand_top_y = baseline_y - rad_bbox.ascent;
    let bar_y = radicand_top_y - v_gap - rule_thick / 2.0;
    let bar_width = rad_bbox.advance;

    // Render the √ sign. Its top should roughly align with bar.
    let mut elems = Vec::new();
    // Substitute √ char (U+221A, not in our substitution table).
    elems.push(emit_text_at(
        '\u{221A}'.to_string(),
        x,
        baseline_y,
        fs,
    ));

    // Render radicand.
    let (rad_elems, _rb) = emit_expr(radicand, radicand_x, baseline_y, ctx);
    elems.extend(rad_elems);

    // Render the horizontal overbar.
    elems.push(LayoutElement::new(
        radicand_x,
        bar_y - rule_thick / 2.0,
        bar_width,
        rule_thick,
        LayoutContent::TableBorder {
            x1: radicand_x,
            y1: bar_y,
            x2: radicand_x + bar_width,
            y2: bar_y,
            color: None,
            width: rule_thick,
            style: None,
        },
    ));

    // Optional degree: small nth-root index to upper-left of √.
    if let Some(deg_expr) = degree {
        let deg_ctx = ctx.descend_script().descend_script(); // ScriptScript
        let raise_du = table.constants.RadicalDegreeBottomRaisePercent; // percent
        let raise_frac = raise_du as f32 / 100.0;
        let deg_baseline = baseline_y - fs * raise_frac;
        let deg_x = x - table.du_to_pt(
            -table.constants.RadicalKernAfterDegree.abs(), fs,
        ).abs().max(fs * 0.15);
        let (de, _db) = emit_expr(deg_expr, deg_x, deg_baseline, &deg_ctx);
        elems.extend(de);
    }

    let bbox = MathBBox {
        advance: sign_width + rad_bbox.advance,
        ascent: rad_bbox.ascent + v_gap + rule_thick + extra_asc,
        descent: rad_bbox.descent,
        italic_correction: 0.0,
    };
    (elems, bbox)
}

/// Emit matrix: 2D grid of cells with per-column alignment and per-row heights.
fn emit_matrix(
    rows: &[Vec<MathExpr>],
    col_align: crate::ir::MathAlignment,
    x: f32,
    baseline_y: f32,
    ctx: &MathLayoutContext,
) -> (Vec<LayoutElement>, MathBBox) {
    use crate::ir::MathAlignment;
    let table = MathTable::cambria_math();
    let fs = ctx.font_size;

    if rows.is_empty() {
        return (vec![], MathBBox::default());
    }
    let n_cols = rows.iter().map(|r| r.len()).max().unwrap_or(0);
    if n_cols == 0 {
        return (vec![], MathBBox::default());
    }

    // Pre-compute bbox for each cell.
    let mut cell_bboxes: Vec<Vec<MathBBox>> = Vec::with_capacity(rows.len());
    for row in rows.iter() {
        let row_bb: Vec<MathBBox> = row.iter()
            .map(|e| layout_expr(e, ctx))
            .collect();
        cell_bboxes.push(row_bb);
    }

    // Column widths: max advance per column.
    let mut col_widths = vec![0.0_f32; n_cols];
    for row in cell_bboxes.iter() {
        for (j, bb) in row.iter().enumerate() {
            if bb.advance > col_widths[j] {
                col_widths[j] = bb.advance;
            }
        }
    }
    // Row heights: max (ascent + descent) per row.
    let row_heights: Vec<f32> = cell_bboxes.iter()
        .map(|row| row.iter()
            .map(|bb| bb.ascent + bb.descent)
            .fold(0.0_f32, f32::max))
        .collect();
    let row_ascents: Vec<f32> = cell_bboxes.iter()
        .map(|row| row.iter()
            .map(|bb| bb.ascent)
            .fold(0.0_f32, f32::max))
        .collect();

    // Inter-column gap ~ MathLeading (from MATH constants).
    let col_gap = table.du_to_pt(table.constants.MathLeading, fs);
    // Inter-row gap: same order of magnitude as col_gap.
    let row_gap = table.du_to_pt(table.constants.MathLeading, fs);

    // Matrix origin y: top of first row = baseline_y - axis_height - half_height.
    // For simplicity, center vertically on the math axis.
    let axis_h = table.du_to_pt(table.constants.AxisHeight, fs);
    let total_height: f32 = row_heights.iter().sum::<f32>()
        + row_gap * (rows.len().saturating_sub(1)) as f32;
    let matrix_top_y = baseline_y - axis_h - total_height / 2.0;

    // Compute column x positions.
    let col_xs: Vec<f32> = {
        let mut xs = Vec::with_capacity(n_cols);
        let mut cur = x;
        for (i, w) in col_widths.iter().enumerate() {
            xs.push(cur);
            cur += *w;
            if i + 1 < n_cols { cur += col_gap; }
        }
        xs
    };

    // Emit each cell.
    let mut elems = Vec::new();
    let mut cur_y = matrix_top_y;
    for (i, row) in rows.iter().enumerate() {
        let row_baseline = cur_y + row_ascents[i];
        for (j, cell) in row.iter().enumerate() {
            if j >= n_cols { break; }
            let cell_bb = &cell_bboxes[i][j];
            let col_w = col_widths[j];
            let col_x = col_xs[j];
            // Align cell within column.
            let cell_x = match col_align {
                MathAlignment::Left => col_x,
                MathAlignment::Right => col_x + col_w - cell_bb.advance,
                _ => col_x + (col_w - cell_bb.advance) / 2.0,  // center or centerGroup
            };
            let (ce, _) = emit_expr(cell, cell_x, row_baseline, ctx);
            elems.extend(ce);
        }
        cur_y += row_heights[i] + row_gap;
    }

    let total_width: f32 = col_widths.iter().sum::<f32>()
        + col_gap * (n_cols.saturating_sub(1)) as f32;

    let bbox = MathBBox {
        advance: total_width,
        ascent: total_height / 2.0 + axis_h,
        descent: total_height / 2.0 - axis_h,
        italic_correction: 0.0,
    };
    (elems, bbox)
}

/// Emit n-ary operator with sub/sup limits.
/// limLoc=undOvr: limits stacked above/below operator.
/// limLoc=subSup: limits as scripts to the right.
fn emit_nary(
    op: char,
    sub: Option<&MathExpr>,
    sup: Option<&MathExpr>,
    operand: &MathExpr,
    lim_loc: crate::ir::LimLoc,
    x: f32,
    baseline_y: f32,
    ctx: &MathLayoutContext,
) -> (Vec<LayoutElement>, MathBBox) {
    use crate::ir::LimLoc;
    let table = MathTable::cambria_math();
    let fs = ctx.font_size;

    // Operator glyph: render larger if grow or display.
    let op_size = if ctx.style.is_display() { fs * 1.6 } else { fs * 1.2 };
    let op_w = op_size * 0.6;

    let mut elems = Vec::new();
    let mut cur_x = x;

    let lim_ctx = ctx.descend_script();
    let sub_bbox = sub.map(|s| layout_expr(s, &lim_ctx));
    let sup_bbox = sup.map(|s| layout_expr(s, &lim_ctx));

    match lim_loc {
        LimLoc::UndOvr => {
            // Center operator and limits on common column.
            let common_w = op_w
                .max(sub_bbox.as_ref().map(|b| b.advance).unwrap_or(0.0))
                .max(sup_bbox.as_ref().map(|b| b.advance).unwrap_or(0.0));
            let op_x = cur_x + (common_w - op_w) / 2.0;
            elems.push(emit_text_at(op.to_string(), op_x, baseline_y, op_size));

            if let (Some(s_expr), Some(s_bb)) = (sup, sup_bbox.as_ref()) {
                let sup_x = cur_x + (common_w - s_bb.advance) / 2.0;
                let rise = table.du_to_pt(table.constants.UpperLimitBaselineRiseMin, fs);
                let gap = table.du_to_pt(table.constants.UpperLimitGapMin, fs);
                let sup_baseline = baseline_y - op_size * 0.8 - gap - rise;
                let (e, _) = emit_expr(s_expr, sup_x, sup_baseline, &lim_ctx);
                elems.extend(e);
            }
            if let (Some(s_expr), Some(s_bb)) = (sub, sub_bbox.as_ref()) {
                let sub_x = cur_x + (common_w - s_bb.advance) / 2.0;
                let drop = table.du_to_pt(table.constants.LowerLimitBaselineDropMin, fs);
                let gap = table.du_to_pt(table.constants.LowerLimitGapMin, fs);
                let sub_baseline = baseline_y + op_size * 0.2 + gap + drop;
                let (e, _) = emit_expr(s_expr, sub_x, sub_baseline, &lim_ctx);
                elems.extend(e);
            }
            cur_x += common_w;
        }
        LimLoc::SubSup => {
            // Operator at baseline, sub/sup as regular scripts to the right.
            elems.push(emit_text_at(op.to_string(), cur_x, baseline_y, op_size));
            cur_x += op_w;
            if let (Some(s_expr), Some(s_bb)) = (sup, sup_bbox.as_ref()) {
                let sup_x = cur_x;
                let shift_up = table.du_to_pt(table.constants.SuperscriptShiftUp, fs);
                let (e, _) = emit_expr(s_expr, sup_x, baseline_y - shift_up, &lim_ctx);
                elems.extend(e);
                cur_x += s_bb.advance;
            }
            if let (Some(s_expr), Some(s_bb)) = (sub, sub_bbox.as_ref()) {
                let shift_down = table.du_to_pt(table.constants.SubscriptShiftDown, fs);
                let (e, _) = emit_expr(s_expr, cur_x - sub_bbox.as_ref().map(|b| b.advance).unwrap_or(0.0),
                                       baseline_y + shift_down, &lim_ctx);
                elems.extend(e);
                if sup_bbox.is_none() { cur_x += s_bb.advance; }
            }
        }
    }

    // Small gap then operand.
    cur_x += fs * 0.1;
    let (op_elems, op_bbox) = emit_expr(operand, cur_x, baseline_y, ctx);
    elems.extend(op_elems);

    let bbox = MathBBox {
        advance: cur_x - x + op_bbox.advance,
        ascent: (op_size * 0.8)
            .max(sup_bbox.as_ref().map(|b| op_size + b.height()).unwrap_or(0.0))
            .max(op_bbox.ascent),
        descent: (op_size * 0.2)
            .max(sub_bbox.as_ref().map(|b| op_size + b.height()).unwrap_or(0.0))
            .max(op_bbox.descent),
        italic_correction: 0.0,
    };
    (elems, bbox)
}

/// Emit function: name + arg (e.g., sin x, log y).
fn emit_function(
    name: &MathExpr,
    arg: &MathExpr,
    x: f32,
    baseline_y: f32,
    ctx: &MathLayoutContext,
) -> (Vec<LayoutElement>, MathBBox) {
    let fs = ctx.font_size;
    let (name_elems, name_bbox) = emit_expr(name, x, baseline_y, ctx);
    let gap = fs * 0.15;
    let arg_x = x + name_bbox.advance + gap;
    let (arg_elems, arg_bbox) = emit_expr(arg, arg_x, baseline_y, ctx);
    let mut elems = name_elems;
    elems.extend(arg_elems);
    let bbox = MathBBox {
        advance: name_bbox.advance + gap + arg_bbox.advance,
        ascent: name_bbox.ascent.max(arg_bbox.ascent),
        descent: name_bbox.descent.max(arg_bbox.descent),
        italic_correction: arg_bbox.italic_correction,
    };
    (elems, bbox)
}

/// Emit group character: brace/bracket above or below base.
fn emit_group_chr(
    chr: char,
    pos: crate::ir::BarPos,
    base: &MathExpr,
    x: f32,
    baseline_y: f32,
    ctx: &MathLayoutContext,
) -> (Vec<LayoutElement>, MathBBox) {
    use crate::ir::BarPos;
    let table = MathTable::cambria_math();
    let fs = ctx.font_size;
    let base_bbox = layout_expr(base, ctx);
    let (base_elems, _) = emit_expr(base, x, baseline_y, ctx);
    let mut elems = base_elems;

    let gap = table.du_to_pt(table.constants.StretchStackGapAboveMin, fs);
    let chr_size = fs * 0.8;
    let chr_x = x + (base_bbox.advance - chr_size * 0.6) / 2.0;

    let chr_baseline = match pos {
        BarPos::Top => baseline_y - base_bbox.ascent - gap - chr_size * 0.2,
        BarPos::Bot => baseline_y + base_bbox.descent + gap + chr_size * 0.8,
    };
    elems.push(emit_text_at(chr.to_string(), chr_x, chr_baseline, chr_size));

    let mut bbox = base_bbox;
    match pos {
        BarPos::Top => bbox.ascent += gap + chr_size,
        BarPos::Bot => bbox.descent += gap + chr_size,
    }
    (elems, bbox)
}

/// Emit equation array: vertically stacked expressions.
fn emit_eq_array(
    items: &[MathExpr],
    x: f32,
    baseline_y: f32,
    ctx: &MathLayoutContext,
) -> (Vec<LayoutElement>, MathBBox) {
    let table = MathTable::cambria_math();
    let fs = ctx.font_size;

    if items.is_empty() {
        return (vec![], MathBBox::default());
    }
    let bboxes: Vec<MathBBox> = items.iter().map(|e| layout_expr(e, ctx)).collect();
    let gap = table.du_to_pt(table.constants.StackGapMin, fs);
    let total_h: f32 = bboxes.iter().map(|b| b.height()).sum::<f32>()
        + gap * items.len().saturating_sub(1) as f32;
    let common_w = bboxes.iter().map(|b| b.advance).fold(0.0_f32, f32::max);

    let axis = table.du_to_pt(table.constants.AxisHeight, fs);
    let mut cur_y = baseline_y - axis - total_h / 2.0;
    let mut elems = Vec::new();
    for (i, item) in items.iter().enumerate() {
        let bb = &bboxes[i];
        let item_baseline = cur_y + bb.ascent;
        let item_x = x + (common_w - bb.advance) / 2.0;
        let (e, _) = emit_expr(item, item_x, item_baseline, ctx);
        elems.extend(e);
        cur_y += bb.height() + gap;
    }
    let bbox = MathBBox {
        advance: common_w,
        ascent: total_h / 2.0 + axis,
        descent: total_h / 2.0 - axis,
        italic_correction: 0.0,
    };
    (elems, bbox)
}

/// Emit pre-script: sub/sup to the LEFT of base (isotope notation ^14_6 C).
fn emit_prescript(
    base: &MathExpr,
    sub: &MathExpr,
    sup: &MathExpr,
    x: f32,
    baseline_y: f32,
    ctx: &MathLayoutContext,
) -> (Vec<LayoutElement>, MathBBox) {
    let table = MathTable::cambria_math();
    let fs = ctx.font_size;
    let s_ctx = ctx.descend_script();
    let sub_bbox = layout_expr(sub, &s_ctx);
    let sup_bbox = layout_expr(sup, &s_ctx);

    let pre_w = sub_bbox.advance.max(sup_bbox.advance);
    let sup_shift = table.du_to_pt(table.constants.SuperscriptShiftUp, fs);
    let sub_shift = table.du_to_pt(table.constants.SubscriptShiftDown, fs);

    let mut elems = Vec::new();
    // Pre-sup: right-aligned at x + pre_w, raised.
    let sup_x = x + pre_w - sup_bbox.advance;
    let (sup_e, _) = emit_expr(sup, sup_x, baseline_y - sup_shift, &s_ctx);
    elems.extend(sup_e);
    // Pre-sub: right-aligned at x + pre_w, lowered.
    let sub_x = x + pre_w - sub_bbox.advance;
    let (sub_e, _) = emit_expr(sub, sub_x, baseline_y + sub_shift, &s_ctx);
    elems.extend(sub_e);

    // Base after pre-scripts.
    let base_x = x + pre_w;
    let (base_elems, base_bbox) = emit_expr(base, base_x, baseline_y, ctx);
    elems.extend(base_elems);

    let bbox = MathBBox {
        advance: pre_w + base_bbox.advance,
        ascent: base_bbox.ascent.max(sup_shift + sup_bbox.ascent),
        descent: base_bbox.descent.max(sub_shift + sub_bbox.descent),
        italic_correction: base_bbox.italic_correction,
    };
    (elems, bbox)
}

/// Emit bar (overline/underline) as TableBorder above or below base.
fn emit_bar(
    pos: crate::ir::BarPos,
    base: &MathExpr,
    x: f32,
    baseline_y: f32,
    ctx: &MathLayoutContext,
) -> (Vec<LayoutElement>, MathBBox) {
    use crate::ir::BarPos;
    let table = MathTable::cambria_math();
    let fs = ctx.font_size;

    let base_bbox = layout_expr(base, ctx);
    let (base_elems, _) = emit_expr(base, x, baseline_y, ctx);
    let mut elems = base_elems;

    let (gap, thick, extra) = match pos {
        BarPos::Top => (
            table.du_to_pt(table.constants.OverbarVerticalGap, fs),
            table.du_to_pt(table.constants.OverbarRuleThickness, fs),
            table.du_to_pt(table.constants.OverbarExtraAscender, fs),
        ),
        BarPos::Bot => (
            table.du_to_pt(table.constants.UnderbarVerticalGap, fs),
            table.du_to_pt(table.constants.UnderbarRuleThickness, fs),
            table.du_to_pt(table.constants.UnderbarExtraDescender, fs),
        ),
    };

    // Bar y position.
    let bar_y = match pos {
        BarPos::Top => baseline_y - base_bbox.ascent - gap - thick / 2.0,
        BarPos::Bot => baseline_y + base_bbox.descent + gap + thick / 2.0,
    };

    elems.push(LayoutElement::new(
        x,
        bar_y - thick / 2.0,
        base_bbox.advance,
        thick,
        LayoutContent::TableBorder {
            x1: x, y1: bar_y,
            x2: x + base_bbox.advance, y2: bar_y,
            color: None,
            width: thick,
            style: None,
        },
    ));

    let mut bbox = base_bbox;
    match pos {
        BarPos::Top => bbox.ascent += gap + thick + extra,
        BarPos::Bot => bbox.descent += gap + thick + extra,
    }
    (elems, bbox)
}

/// Emit accent: combining accent char positioned above base, centered via
/// TopAccentAttachment (or geometric center as fallback).
fn emit_accent(
    accent: char,
    base: &MathExpr,
    x: f32,
    baseline_y: f32,
    ctx: &MathLayoutContext,
) -> (Vec<LayoutElement>, MathBBox) {
    let table = MathTable::cambria_math();
    let glyphs = MathGlyphTables::cambria_math();
    let fs = ctx.font_size;

    let base_bbox = layout_expr(base, ctx);
    let (base_elems, _) = emit_expr(base, x, baseline_y, ctx);
    let mut elems = base_elems;

    // Horizontal attachment: look up the substituted first char of base.
    let first_char = extract_first_char(base);
    let attach_x = if let Some(c) = first_char {
        let sub = math_substitute(c);
        glyphs.top_accent_attachment(sub)
            .map(|du| table.du_to_pt(du, fs))
            .unwrap_or(base_bbox.advance / 2.0)
    } else {
        base_bbox.advance / 2.0
    };

    // Accent y: above base's ascent with OverbarVerticalGap gap.
    let acc_size = fs * 0.9; // slightly smaller than base
    let ascent = base_bbox.ascent;
    let accent_y_top = baseline_y - ascent
        - table.du_to_pt(table.constants.OverbarVerticalGap, fs);
    // emit_text_at places text at top = baseline - ascent_approx.
    // We want the accent's BOTTOM at accent_y_top. Approximating accent
    // ascent as fs*0.8, we set its baseline at accent_y_top + fs*0.8.
    let accent_baseline = accent_y_top + acc_size * 0.2;
    // Accent char rendered at (x + attach_x, accent_baseline) but shifted
    // left by half accent width for visual centering.
    let accent_w = acc_size * 0.4;
    let accent_x = x + attach_x - accent_w / 2.0;
    elems.push(emit_text_at(accent.to_string(), accent_x, accent_baseline, acc_size));

    let bbox = MathBBox {
        advance: base_bbox.advance,
        ascent: base_bbox.ascent + table.du_to_pt(table.constants.OverbarVerticalGap, fs)
            + acc_size,
        descent: base_bbox.descent,
        italic_correction: 0.0,
    };
    (elems, bbox)
}

/// Extract first leaf char of a MathExpr (for accent attachment lookup).
fn extract_first_char(expr: &MathExpr) -> Option<char> {
    match expr {
        MathExpr::Text(s) | MathExpr::Run { text: s, .. } => s.chars().next(),
        MathExpr::Seq(children) => children.iter().find_map(extract_first_char),
        MathExpr::Superscript { base, .. } | MathExpr::Subscript { base, .. }
        | MathExpr::SubSuperscript { base, .. } | MathExpr::PreScript { base, .. }
        | MathExpr::Accent { base, .. } | MathExpr::Bar { base, .. }
        | MathExpr::Limit { base, .. } | MathExpr::GroupChar { base, .. }
        | MathExpr::BorderBox { base, .. } => extract_first_char(base),
        MathExpr::BoxExpr(inner) | MathExpr::Phantom(inner) => extract_first_char(inner),
        MathExpr::Radical { radicand, .. } => extract_first_char(radicand),
        _ => None,
    }
}

/// Emit limit: base with lim expression above (limUpp) or below (limLow).
fn emit_limit(
    base: &MathExpr,
    lim: &MathExpr,
    pos: crate::ir::LimitPos,
    x: f32,
    baseline_y: f32,
    ctx: &MathLayoutContext,
) -> (Vec<LayoutElement>, MathBBox) {
    use crate::ir::LimitPos;
    let table = MathTable::cambria_math();
    let fs = ctx.font_size;

    let base_bbox = layout_expr(base, ctx);
    let lim_ctx = ctx.descend_script();
    let lim_bbox = layout_expr(lim, &lim_ctx);

    // Center lim horizontally on base.
    let common_w = base_bbox.advance.max(lim_bbox.advance);
    let base_x = x + (common_w - base_bbox.advance) / 2.0;
    let lim_x = x + (common_w - lim_bbox.advance) / 2.0;

    let (base_elems, _) = emit_expr(base, base_x, baseline_y, ctx);
    let mut elems = base_elems;

    let (lim_baseline, _gap) = match pos {
        LimitPos::Lower => {
            let gap = table.du_to_pt(table.constants.LowerLimitGapMin, fs);
            let drop = table.du_to_pt(table.constants.LowerLimitBaselineDropMin, fs);
            let lb = baseline_y + base_bbox.descent + gap + drop + lim_bbox.ascent;
            (lb, gap)
        }
        LimitPos::Upper => {
            let gap = table.du_to_pt(table.constants.UpperLimitGapMin, fs);
            let rise = table.du_to_pt(table.constants.UpperLimitBaselineRiseMin, fs);
            let lb = baseline_y - base_bbox.ascent - gap - rise;
            (lb, gap)
        }
    };
    let (lim_elems, _) = emit_expr(lim, lim_x, lim_baseline, &lim_ctx);
    elems.extend(lim_elems);

    let bbox = match pos {
        LimitPos::Lower => MathBBox {
            advance: common_w,
            ascent: base_bbox.ascent,
            descent: (baseline_y + base_bbox.descent - baseline_y)
                + (lim_baseline - baseline_y) - base_bbox.descent + lim_bbox.descent,
            italic_correction: 0.0,
        },
        LimitPos::Upper => MathBBox {
            advance: common_w,
            ascent: (baseline_y - lim_baseline) + lim_bbox.ascent,
            descent: base_bbox.descent,
            italic_correction: 0.0,
        },
    };
    (elems, bbox)
}

/// Emit delimiter: begChr on left, content, endChr on right.
/// Delimiter glyphs render at the base font size (not stretched yet —
/// Phase 3 later adds MATH vertical_variants for grow).
fn emit_delimiter(
    beg: char,
    end: char,
    content: &MathExpr,
    x: f32,
    baseline_y: f32,
    ctx: &MathLayoutContext,
) -> (Vec<LayoutElement>, MathBBox) {
    let fs = ctx.font_size;
    let mut elems = Vec::new();

    // Left delimiter char.
    let left_w = fs * 0.45;
    elems.push(emit_text_at(beg.to_string(), x, baseline_y, fs));

    // Content bbox to determine overall advance.
    let content_bbox = layout_expr(content, ctx);
    let content_x = x + left_w;
    let (content_elems, _) = emit_expr(content, content_x, baseline_y, ctx);
    elems.extend(content_elems);

    // Right delimiter char.
    let right_x = content_x + content_bbox.advance;
    let right_w = fs * 0.45;
    elems.push(emit_text_at(end.to_string(), right_x, baseline_y, fs));

    let bbox = MathBBox {
        advance: left_w + content_bbox.advance + right_w,
        ascent: content_bbox.ascent.max(fs * 0.8),
        descent: content_bbox.descent.max(fs * 0.2),
        italic_correction: 0.0,
    };
    (elems, bbox)
}

/// Emit combined sub+superscript: base with sub below and sup above at same x.
fn emit_subsuperscript(
    base: &MathExpr,
    sub: &MathExpr,
    sup: &MathExpr,
    x: f32,
    baseline_y: f32,
    ctx: &MathLayoutContext,
) -> (Vec<LayoutElement>, MathBBox) {
    let table = MathTable::cambria_math();
    let fs = ctx.font_size;
    let (mut base_elems, base_bbox) = emit_expr(base, x, baseline_y, ctx);

    let s_ctx = ctx.descend_script();
    let sup_shift = table.du_to_pt(table.constants.SuperscriptShiftUp, fs);
    let sub_shift = table.du_to_pt(table.constants.SubscriptShiftDown, fs);

    let script_x = x + base_bbox.advance + base_bbox.italic_correction;
    let (sup_e, sup_b) = emit_expr(sup, script_x, baseline_y - sup_shift, &s_ctx);
    let (sub_e, sub_b) = emit_expr(sub, script_x, baseline_y + sub_shift, &s_ctx);
    base_elems.extend(sup_e);
    base_elems.extend(sub_e);

    let bbox = MathBBox {
        advance: base_bbox.advance + base_bbox.italic_correction
            + sup_b.advance.max(sub_b.advance),
        ascent: base_bbox.ascent.max(sup_shift + sup_b.ascent),
        descent: base_bbox.descent.max(sub_shift + sub_b.descent),
        italic_correction: 0.0,
    };
    (base_elems, bbox)
}

/// Emit positioned LayoutElements for a full MathBlock.
/// Returns (elements, total bbox). Origin: top-left at (x, cursor_y).
pub fn emit_math_block(
    block: &MathBlock,
    x: f32,
    cursor_y: f32,
    font_size: f32,
) -> (Vec<LayoutElement>, MathBBox) {
    let ctx = MathLayoutContext {
        font_size,
        style: MathStyle::from_block(block),
    };
    let exprs: &[MathExpr] = match block {
        MathBlock::Inline(xs) => xs,
        MathBlock::Display { content, .. } => content,
    };
    // Pre-compute baseline: first pass finds needed ascent.
    let mut total_bbox = MathBBox::default();
    for e in exprs {
        let b = layout_expr(e, &ctx);
        total_bbox = total_bbox.hstack(&b);
    }
    // Baseline sits at cursor_y + ascent (top-relative).
    let baseline_y = cursor_y + total_bbox.ascent.max(font_size * 0.8);

    let mut elems = Vec::new();
    let mut cur_x = x;
    for e in exprs {
        let (ee, b) = emit_expr(e, cur_x, baseline_y, &ctx);
        elems.extend(ee);
        cur_x += b.advance;
    }
    (elems, total_bbox)
}

/// Flatten a whole MathBlock to a single text string (substituted).
pub fn extract_flat_text_block(block: &MathBlock) -> String {
    let exprs: &[MathExpr] = match block {
        MathBlock::Inline(xs) => xs,
        MathBlock::Display { content, .. } => content,
    };
    let mut out = String::new();
    for e in exprs {
        append_flat(&mut out, e);
    }
    out
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
    fn extract_flat_fraction() {
        let frac = MathExpr::Fraction {
            num: Box::new(MathExpr::Text("a".to_string())),
            den: Box::new(MathExpr::Text("b".to_string())),
            bar_type: FracBarType::Bar,
        };
        let text = extract_flat_text(&frac);
        // Both chars should be math-substituted (𝑎, 𝑏) with '/' between.
        assert_eq!(text, "\u{1D44E}/\u{1D44F}");
    }

    #[test]
    fn extract_flat_superscript() {
        let sup = MathExpr::Superscript {
            base: Box::new(MathExpr::Text("x".to_string())),
            sup: Box::new(MathExpr::Text("2".to_string())),
        };
        assert_eq!(extract_flat_text(&sup), "\u{1D465}^2"); // 𝑥^2
    }

    #[test]
    fn extract_flat_nested_delim() {
        // (a + b) → parenthesized substituted chars
        let inner = MathExpr::Seq(vec![
            MathExpr::Text("a".to_string()),
            MathExpr::Text("+".to_string()),
            MathExpr::Text("b".to_string()),
        ]);
        let d = MathExpr::Delimiter {
            beg: '(', end: ')', sep: None,
            content: Box::new(inner),
        };
        assert_eq!(extract_flat_text(&d), "(\u{1D44E}+\u{1D44F})"); // (𝑎+𝑏)
    }

    #[test]
    fn extract_flat_block() {
        let block = MathBlock::Display {
            content: vec![
                MathExpr::Text("E".to_string()),
                MathExpr::Text("=".to_string()),
                MathExpr::Text("mc".to_string()),
            ],
            jc: MathAlignment::Center,
        };
        let t = extract_flat_text_block(&block);
        // E→𝐸, = unchanged, m→𝑚, c→𝑐
        assert_eq!(t, "\u{1D438}=\u{1D45A}\u{1D450}");
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
