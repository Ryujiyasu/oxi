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
        // Primitives not yet implemented — return zero bbox.
        // Phase 3 fills in Nary / Matrix / Delimiter / Accent / etc.
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
            color: None,
            highlight: None,
            field_type: None,
            character_spacing: 0.0,
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
