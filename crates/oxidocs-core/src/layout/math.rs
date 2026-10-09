// This Source Code Form is subject to the terms of the Mozilla Public
// License, v. 2.0. If a copy of the MPL was not distributed with this
// file, You can obtain one at https://mozilla.org/MPL/2.0/.

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
use crate::layout::{LayoutElement, LayoutContent, TextEffects, FontGlyph};
use crate::font::math_stretch::{StretchTable, StretchPlan, Direction};

struct RadicalGeometry {
    plan: StretchPlan,
    advance: f32,
    ink_top: f32,
    ink_bottom: f32,
    baseline_shift: f32,
    rule_thickness: f32,
}

fn radical_geometry(radicand: &MathExpr, ctx: &MathLayoutContext) -> Option<RadicalGeometry> {
    let data=StretchTable::cambria_math();
    let table=MathTable::cambria_math();
    let fs=ctx.effective_font_size();
    if fs<=0.0 { return None; }
    let scale=fs/data.upm as f32;
    let (a,d)=ink_extent_word(radicand,ctx,true);
    let gap_du = table.constants.RadicalRuleThickness;
    let gap=gap_du as f32*scale;
    let thickness=table.constants.RadicalRuleThickness as f32*scale;
    let target=a+d+gap+thickness;
    let construction=data.construction(Direction::Vert,'\u{221a}')?;
    let plan=data.plan(construction,target as f64/scale as f64).ok()?;
    let upper=plan.placements.iter().map(|p|p.advance_offset as f32+p.glyph.bounds[3] as f32)
        .fold(f32::NEG_INFINITY,f32::max);
    let lower=plan.placements.iter().map(|p|p.advance_offset as f32+p.glyph.bounds[1] as f32)
        .fold(f32::INFINITY,f32::min);
    let height=(upper-lower)*scale;
    let extra=(height-target).max(0.0);
    let top=-a-gap-thickness-extra*0.5;
    let advance=plan.placements.iter().map(|p|p.glyph.advance_width as f32*scale).fold(0.0,f32::max);
    Some(RadicalGeometry{plan,advance,ink_top:top,ink_bottom:top+height,
        baseline_shift:top+upper*scale,rule_thickness:thickness})
}

fn radical_degree(degree: Option<&MathExpr>, ctx: &MathLayoutContext, shape: &RadicalGeometry)
    -> (f32, Option<(f32,f32,MathLayoutContext)>) {
    let Some(expr)=degree else{return (0.0,None)};
    let table=MathTable::cambria_math();let fs=ctx.effective_font_size();
    let dctx=ctx.descend_script().descend_script();let bb=layout_expr(expr,&dctx);
    let before=table.du_to_pt(table.constants.RadicalKernBeforeDegree,fs);
    let after=table.du_to_pt(table.constants.RadicalKernAfterDegree,fs);
    let prefix=(before+bb.advance+after).max(0.0);
    let (_,dd)=ink_extent_word(expr,&dctx,false);
    let raise=(shape.ink_bottom-shape.ink_top)*table.constants.RadicalDegreeBottomRaisePercent as f32/100.0;
    (prefix,Some((before,shape.ink_bottom-raise-dd,dctx)))
}

/// Painted font ink, distinct from the text line box or assembly advance.
pub(crate) fn painted_element_ink(e: &LayoutElement) -> (f32,f32) {
    if let LayoutContent::Text{text,font_size,..}=&e.content {
        let baseline=e.y+e.baseline_offset.unwrap_or(e.height*(2.0/3.0));
        if let Some(g)=e.font_glyph {
            return (baseline-g.bounds_em[3]*font_size,baseline-g.bounds_em[1]*font_size);
        }
        let extents:Vec<_>=text.chars().filter_map(glyph_ink_du).collect();
        if !extents.is_empty() {
            return (extents.iter().map(|(a,_)|baseline-a*font_size).fold(f32::INFINITY,f32::min),
                    extents.iter().map(|(_,d)|baseline+d*font_size).fold(f32::NEG_INFINITY,f32::max));
        }
    }
    (e.y,e.y+e.height)
}

/// Painted extents plus structural whitespace reserved by OpenType MATH.
/// The radical's extra ascender is space, so it is absent from glyph bounds.
/// An assembly is a contiguous group at one x origin; reserve above its whole
/// construction, including the uppermost extender and end glyph.
pub(crate) fn reserved_line_extents(elements: &[LayoutElement]) -> (f32, f32) {
    let mut top=f32::INFINITY;
    let mut bottom=f32::NEG_INFINITY;
    let table=MathTable::cambria_math();
    for (index,e) in elements.iter().enumerate() {
        let (a,d)=painted_element_ink(e);
        top=top.min(a);bottom=bottom.max(d);
        if is_fallback_font_element(e) {
            top=top.min(e.y);bottom=bottom.max(e.y+e.height);
        }
        if let LayoutContent::Text{text,font_size,..}=&e.content {
            if e.font_glyph.is_some() && text == "\u{221a}" {
                let mut construction_top=a;
                for component in &elements[index+1..] {
                    let same_size=matches!(&component.content,
                        LayoutContent::Text{font_size:size,..} if size == font_size);
                    if component.font_glyph.is_none() || component.x != e.x || !same_size { break; }
                    construction_top=construction_top.min(painted_element_ink(component).0);
                }
                top=top.min(construction_top-table.du_to_pt(table.constants.RadicalExtraAscender,*font_size));
            }
        }
    }
    (top,bottom)
}

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
        let constants = &MathTable::cambria_math().constants;
        let percent = match self.style {
            MathStyle::Script => constants.ScriptPercentScaleDown,
            MathStyle::ScriptScript => constants.ScriptScriptPercentScaleDown,
            _ => return self.font_size,
        };
        // Script size is expressed in half-point units after scaling.
        // Word's 10-size sweep fits 73/60 percent with this quantization;
        // the PDF's separate 600dpi rounding belongs to the renderer.
        ((self.font_size * percent as f32 / 100.0 * 2.0).floor() / 2.0).max(0.5)
    }

    /// Fractions use compact shifts below the outer display fraction.
    /// Font-size reduction is a separate policy from selecting those shifts.
    pub fn descend_fraction(&self) -> MathLayoutContext {
        let style = match self.style {
            MathStyle::Display => MathStyle::CompactFullSize,
            MathStyle::DisplayReducedFractions => MathStyle::Text,
            MathStyle::CompactFullSize => MathStyle::CompactFullSize,
            _ => self.style.script_style(),
        };
        MathLayoutContext { font_size: self.font_size, style }
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
/// S527 (coverage): per-char Cambria Math advance estimate (em). The flat 0.5em
/// was far too narrow — Word-measured advances: `=`0.75, `+`0.97, `m`/`M`0.85,
/// `E`0.63, `x`0.54, `i`/`l`0.32, digits 0.56. A per-class estimate (no committed
/// Cambria Math advance table exists yet) keeps operators/wide letters wide and
/// narrow letters narrow, so operator expressions/identifiers don't pack too tight.
pub fn glyph_advance_em(c: char) -> f32 {
    // S1258 (2026-08-29, default ON, opt-out OXI_S1258_DISABLE): ask the real
    // face. ★The lookup takes the SUBSTITUTED character — `a` is drawn as the
    // math italic U+1D44E and that is the glyph whose width Word advances by;
    // measuring the ASCII `a` (which Cambria Math also carries, at a different
    // width) is a second, quieter version of the same bug.
    if std::env::var("OXI_S1258_DISABLE").is_err() {
        let t = crate::font::math_glyphs::MathAdvances::cambria_math();
        if let Some(em) = t.advance_em(math_substitute(c)).or_else(|| t.advance_em(c)) {
            return em;
        }
    }
    match c {
        'i' | 'j' | 'l' | 'ı' | '.' | ',' | ';' | ':' | '\'' | '!' | '|' | 'f' | 't' | 'r' => 0.33,
        'm' | 'w' | 'M' | 'W' => 0.86,
        '=' | '≠' | '≈' | '≡' => 0.75,
        '+' => 0.97,
        '-' | '\u{2212}' | '±' | '∓' | '×' | '÷' | '∗' | '⋅' => 0.88,
        '<' | '>' | '≤' | '≥' => 0.78,
        '(' | ')' | '[' | ']' | '{' | '}' | '/' => 0.4,
        '0'..='9' => 0.56,
        'A'..='Z' => 0.68,
        _ => 0.52,
    }
}

/// Advance of the selected codepoint. Do not italicize an upright run again.
fn painted_glyph_advance_em(c: char) -> f32 {
    crate::font::math_glyphs::MathAdvances::cambria_math().advance_em(c)
        .unwrap_or_else(|| glyph_advance_em(c))
}

/// Select the glyph shape at the math style depth, independently of size.
fn script_level(ctx: &MathLayoutContext) -> u8 {
    match ctx.style { MathStyle::Script => 1, MathStyle::ScriptScript => 2, _ => 0 }
}

fn script_glyph(c: char, level: u8) -> Option<crate::font::math_script_glyphs::ScriptGlyph> {
    crate::font::math_script_glyphs::MathScriptGlyphs::cambria_math().alternate(c,level)
}

fn selected_advance_em(c: char, level: u8) -> f32 {
    script_glyph(c,level).map_or_else(||painted_glyph_advance_em(c),|g|g.advance_em)
}

fn selected_italic_correction(c: char, level: u8, fs: f32) -> f32 {
    if let Some(g)=script_glyph(c,level) { return g.italic_correction_em*fs; }
    let table=MathTable::cambria_math();
    MathGlyphTables::cambria_math().italic_correction(c)
        .map(|du|table.du_to_pt(du,fs)).unwrap_or(0.0)
}

fn selected_ink_em(c: char, level: u8) -> (f32,f32) {
    script_glyph(c,level).map(|g|(g.bounds_em[3],-g.bounds_em[1]))
        .or_else(||glyph_ink_du(c)).unwrap_or((0.7,0.2))
}

/// Literal runs retain ordinary text shaping; `ssty` belongs to math runs.
fn run_script_level(style: &crate::ir::MathRunStyle, ctx: &MathLayoutContext) -> u8 {
    if style.literal { 0 } else { script_level(ctx) }
}

fn space_after_script(ctx: &MathLayoutContext) -> f32 {
    let table=MathTable::cambria_math();
    table.du_to_pt(table.constants.SpaceAfterScript,ctx.effective_font_size())
}

fn kern_leaf_gid(expr: &MathExpr,ctx: &MathLayoutContext) -> Option<u16> {
    let c=match expr {
        MathExpr::Text(text) => {
            let mut chars=text.chars();let c=chars.next()?;
            if chars.next().is_some(){return None;}math_substitute(c)
        }
        MathExpr::Run{text,style} if !style.literal => {
            let mut chars=text.chars();let c=chars.next()?;
            if chars.next().is_some(){return None;}
            crate::font::math_substitute::math_run_substitute(c,style)
        }
        MathExpr::Seq(children) if children.len()==1 => return kern_leaf_gid(&children[0],ctx),
        _ => return None,
    };
    script_glyph(c,script_level(ctx)).map(|g|g.index)
        .or_else(||crate::font::math_kern::MathKerns::cambria_math().ordinary_gid(c))
}

/// OpenType MATH: evaluate both contact heights and use the smaller sum.
/// A compound box has zero corner kerning; adjacent glyphs retain their own
/// corner kerning and font size. Missing font tables contribute zero.
fn script_kern(base: &MathExpr,script: &MathExpr,ctx: &MathLayoutContext,
               shift: f32,superscript: bool) -> f32 {
    use crate::font::math_kern::Corner;
    let script_ctx=ctx.descend_script();
    let(ba,bd)=ink_extent_word(base,ctx,false);
    let(sa,sd)=ink_extent_word(script,&script_ctx,false);
    if superscript {
        let at_script_bottom=math_contact_kern(base,ctx,Corner::TopRight,shift-sd)
            +math_contact_kern(script,&script_ctx,Corner::BottomLeft,-sd);
        let at_base_top=math_contact_kern(base,ctx,Corner::TopRight,ba)
            +math_contact_kern(script,&script_ctx,Corner::BottomLeft,ba-shift);
        at_script_bottom.min(at_base_top)
    } else {
        let at_script_top=math_contact_kern(base,ctx,Corner::BottomRight,sa-shift)
            +math_contact_kern(script,&script_ctx,Corner::TopLeft,sa);
        let at_base_bottom=math_contact_kern(base,ctx,Corner::BottomRight,-bd)
            +math_contact_kern(script,&script_ctx,Corner::TopLeft,shift-bd);
        at_script_top.min(at_base_bottom)
    }
}

/// Separate combined scripts using the font's minimum ink gap. Raise the
/// superscript up to its allowed bottom height, then lower the subscript for
/// any remaining deficit. Both scripts retain their actual style and size.
fn combined_script_shifts(
    sub: &MathExpr,
    sup: &MathExpr,
    ctx: &MathLayoutContext,
    cramped: bool,
) -> (f32, f32) {
    let table = MathTable::cambria_math();
    let fs = ctx.effective_font_size();
    let script_ctx = ctx.descend_script();
    let (sub_ascent, _) = ink_extent_word(sub, &script_ctx, true);
    let (_, sup_descent) = ink_extent_word(sup, &script_ctx, cramped);
    let up_constant = if cramped {
        table.constants.SuperscriptShiftUpCramped
    } else {
        table.constants.SuperscriptShiftUp
    };
    let mut up = table.du_to_pt(up_constant, fs);
    let mut down = table.du_to_pt(table.constants.SubscriptShiftDown, fs);
    let gap = up + down - sup_descent - sub_ascent;
    let minimum = table.du_to_pt(table.constants.SubSuperscriptGapMin, fs);
    if gap < minimum {
        let limit = table.du_to_pt(table.constants.SuperscriptBottomMaxWithSubscript, fs);
        let raise = (limit - (up - sup_descent)).max(0.0).min(minimum - gap);
        up += raise;
        down += (minimum - gap - raise).max(0.0);
    }
    (up, down)
}

struct DelimiterGeometry {
    plan: StretchPlan,
    advance: f32,
    baseline_shift: f32,
    ink_top: f32,
    ink_bottom: f32,
}

fn delimiter_geometry(chr: char, content: &MathExpr, ctx: &MathLayoutContext,
                      cramped: bool) -> Option<DelimiterGeometry> {
    let data = StretchTable::cambria_math();
    let fs = ctx.effective_font_size();
    if fs <= 0.0 { return None; }
    let construction = data.construction(Direction::Vert, chr)?;
    let scale = fs / data.upm as f32;
    let (a, d) = ink_extent_word(content, ctx, cramped);
    let plan = data.plan(construction, (a + d) as f64 / scale as f64).ok()?;
    let upper = plan.placements.iter().map(|p| p.advance_offset as f32 + p.glyph.bounds[3] as f32)
        .fold(f32::NEG_INFINITY, f32::max) * scale;
    let lower = plan.placements.iter().map(|p| p.advance_offset as f32 + p.glyph.bounds[1] as f32)
        .fold(f32::INFINITY, f32::min) * scale;
    let table = MathTable::cambria_math();
    let baseline_shift = (upper + lower) * 0.5 - table.du_to_pt(table.constants.AxisHeight, fs);
    let advance = plan.placements.iter().map(|p| p.glyph.advance_width as f32 * scale)
        .fold(0.0, f32::max);
    Some(DelimiterGeometry { plan, advance, baseline_shift,
        ink_top: baseline_shift - upper, ink_bottom: baseline_shift - lower })
}

fn delimiter_width(chr: char, content: &MathExpr, ctx: &MathLayoutContext) -> f32 {
    if chr == '\0' { return 0.0; }
    delimiter_geometry(chr, content, ctx, false).map_or_else(
        || ctx.effective_font_size() * painted_glyph_advance_em(chr), |g| g.advance)
}

fn delimiter_ink(chr: char, content: &MathExpr, ctx: &MathLayoutContext,
                 cramped: bool) -> (f32, f32) {
    if chr == '\0' { return (0.0, 0.0); }
    if let Some(g) = delimiter_geometry(chr, content, ctx, cramped) {
        return ((-g.ink_top).max(0.0), g.ink_bottom.max(0.0));
    }
    let (a, d) = glyph_ink_du(chr).unwrap_or((0.7, 0.2));
    (a * ctx.effective_font_size(), d * ctx.effective_font_size())
}

fn emit_delimiter_glyph(chr: char, content: &MathExpr, x: f32, baseline: f32,
                        ctx: &MathLayoutContext) -> Vec<LayoutElement> {
    if chr == '\0' { return Vec::new(); }
    let fs = ctx.effective_font_size();
    if let Some(g) = delimiter_geometry(chr, content, ctx, false) {
        let data = StretchTable::cambria_math();
        let scale = fs / data.upm as f32;
        return g.plan.placements.iter().enumerate().map(|(i, p)| {
            let mut e = emit_text_at(if i == 0 { chr.to_string() } else { String::new() }, x,
                baseline + g.baseline_shift - p.advance_offset as f32 * scale, fs);
            e.width = p.glyph.advance_width as f32 * scale;
            e.font_glyph = Some(FontGlyph { index: p.glyph.gid,
                bounds_em: p.glyph.bounds.map(|v| v as f32 / data.upm as f32) });
            e
        }).collect();
    }
    vec![emit_text_at(chr.to_string(), x, baseline, fs)]
}

struct AccentGeometry {
    plan: StretchPlan,
    baseline_shift: f32,
    x_shift: f32,
    ink_top: f32,
    ink_bottom: f32,
}

fn accent_geometry(accent: char,base: &MathExpr,ctx: &MathLayoutContext) -> Option<AccentGeometry> {
    let data=StretchTable::cambria_math();let fs=ctx.effective_font_size();
    if fs<=0.0{return None;}
    let construction=data.construction(Direction::Horiz,accent)?;
    let bbox=layout_expr(base,ctx);let scale=fs/data.upm as f32;
    let plan=data.plan(construction,bbox.advance as f64/scale as f64).ok()?;
    let(ba,_)=ink_extent_word(base,ctx,false);
    let table=MathTable::cambria_math();
    let baseline_shift=-(ba-table.du_to_pt(table.constants.AccentBaseHeight,fs)).max(0.0);
    let x_shift=(bbox.advance-plan.advance_measurement as f32*scale)/2.0;
    let ink_top=plan.placements.iter().map(|p|baseline_shift-p.glyph.bounds[3]as f32*scale)
        .fold(f32::INFINITY,f32::min);
    let ink_bottom=plan.placements.iter().map(|p|baseline_shift-p.glyph.bounds[1]as f32*scale)
        .fold(f32::NEG_INFINITY,f32::max);
    Some(AccentGeometry{plan,baseline_shift,x_shift,ink_top,ink_bottom})
}

/// A nested fraction argument includes space outside its own rule. The
/// enclosing fraction uses that width to center both arguments and draw its
/// rule; the inner rule keeps its tight content width. Singleton rows and
/// transparent boxes retain the structural argument.
fn fraction_argument_side_space(expr: &MathExpr, ctx: &MathLayoutContext) -> f32 {
    match expr {
        MathExpr::Fraction { .. } => ctx.effective_font_size() * 0.1,
        MathExpr::Seq(children) if children.len() == 1 => fraction_argument_side_space(&children[0], ctx),
        MathExpr::BoxExpr(inner) | MathExpr::Phantom(inner) => fraction_argument_side_space(inner, ctx),
        _ => 0.0,
    }
}

/// A structural radical reserves the font's extra ascender inside its parent
/// fraction. A row retains this space too; it cannot disappear next to text.
/// Corner kerning and optical measurements still use tight glyph ink.
fn fraction_child_extents(expr: &MathExpr,ctx: &MathLayoutContext,cramped: bool)->(f32,f32) {
    match expr {
        MathExpr::Radical{..}=>{
            let(a,d)=ink_extent_word(expr,ctx,cramped);let table=MathTable::cambria_math();
            (a+table.du_to_pt(table.constants.RadicalExtraAscender,ctx.effective_font_size()),d)
        }
        MathExpr::Seq(children)=>children.iter().map(|e|fraction_child_extents(e,ctx,cramped))
            .fold((0.0f32,0.0f32),|(a,d),(ca,cd)|(a.max(ca),d.max(cd))),
        MathExpr::BoxExpr(inner)|MathExpr::Phantom(inner)=>fraction_child_extents(inner,ctx,cramped),
        _=>ink_extent_word(expr,ctx,cramped),
    }
}

struct NaryGeometry {
    plan: StretchPlan,
    scale: f32,
    operator_baseline: f32,
    operator_x: f32,
    sub_position: Option<(f32, f32)>,
    sup_position: Option<(f32, f32)>,
    operand_x: f32,
    bbox: MathBBox,
    ink_top: f32,
    ink_bottom: f32,
}

/// The same selected glyph and positions are used by measurement, ink
/// accounting and emission. Coordinates are relative to the operand baseline.
fn nary_geometry(
    op: char, sub: Option<&MathExpr>, sup: Option<&MathExpr>, operand: &MathExpr,
    lim_loc: crate::ir::LimLoc, grow: bool, ctx: &MathLayoutContext, cramped: bool,
) -> Option<NaryGeometry> {
    use crate::font::math_stretch::Placement;
    use crate::ir::LimLoc;
    let data = StretchTable::cambria_math();
    let table = MathTable::cambria_math();
    let fs = ctx.effective_font_size();
    if fs <= 0.0 { return None; }
    let construction = data.construction(Direction::Vert, op)?;
    let scale = fs / data.upm as f32;
    let axis = table.du_to_pt(table.constants.AxisHeight, fs);
    let (pa, pd) = ink_extent_word(operand, ctx, cramped);
    let target = 2.0 * (pa - axis).max(pd + axis).max(0.0);
    let target = if ctx.style.is_display() {
        target.max(table.du_to_pt(table.constants.DisplayOperatorMinHeight, fs))
    } else { target };
    let plan = if grow || ctx.style.is_display() {
        data.plan(construction, target as f64 / scale as f64).ok()?
    } else {
        StretchPlan { direction: Direction::Vert,
            advance_measurement: (construction.base.bounds[3] - construction.base.bounds[1]) as f64,
            italic_correction: construction.base.italic_correction.unwrap_or(0),
            placements: vec![Placement { glyph: construction.base.clone(), advance_offset: 0.0 }],
            assembled: false }
    };
    let upper = plan.placements.iter()
        .map(|p| p.advance_offset as f32 + p.glyph.bounds[3] as f32)
        .fold(f32::NEG_INFINITY, f32::max) * scale;
    let lower = plan.placements.iter()
        .map(|p| p.advance_offset as f32 + p.glyph.bounds[1] as f32)
        .fold(f32::INFINITY, f32::min) * scale;
    let operator_baseline = (upper + lower) * 0.5 - axis;
    let advance = plan.placements.iter().map(|p| p.glyph.advance_width as f32 * scale)
        .fold(0.0, f32::max);
    // A variant has its own correction. The assembly correction belongs only
    // to a connected assembly, not to every ready-made operator glyph.
    let italic = if plan.assembled { plan.italic_correction as f32 * scale }
        else { plan.placements[0].glyph.italic_correction.unwrap_or(0) as f32 * scale };
    let stacked = matches!(lim_loc, LimLoc::UndOvr)
        || (ctx.style.is_display() && !('\u{222b}'..='\u{2233}').contains(&op));
    let script_ctx = ctx.descend_script();
    let sub_box = sub.map(|expr| layout_expr(expr, &script_ctx));
    let sup_box = sup.map(|expr| layout_expr(expr, &script_ctx));
    let mut operator_x = 0.0;
    let mut sub_position = None;
    let mut sup_position = None;
    let limits_right;
    let mut ink_top = operator_baseline - upper;
    let mut ink_bottom = operator_baseline - lower;
    if stacked {
        let width = advance.max(sub_box.as_ref().map_or(0.0, |b| b.advance))
            .max(sup_box.as_ref().map_or(0.0, |b| b.advance));
        operator_x = (width - advance) * 0.5;
        if let (Some(expr), Some(bbox)) = (sup, sup_box.as_ref()) {
            let (a,d) = ink_extent_word(expr,&script_ctx,cramped);
            let up = table.du_to_pt(table.constants.UpperLimitBaselineRiseMin, fs)
                .max(-ink_top + table.du_to_pt(table.constants.UpperLimitGapMin, fs) + d);
            sup_position = Some(((width - bbox.advance) * 0.5, -up));
            ink_top = ink_top.min(-up - a);
            ink_bottom = ink_bottom.max(-up + d);
        }
        if let (Some(expr), Some(bbox)) = (sub, sub_box.as_ref()) {
            let (a,d) = ink_extent_word(expr,&script_ctx,true);
            let down = table.du_to_pt(table.constants.LowerLimitBaselineDropMin, fs)
                .max(operator_baseline - lower + table.du_to_pt(table.constants.LowerLimitGapMin, fs) + a);
            sub_position = Some(((width - bbox.advance) * 0.5, down));
            ink_top = ink_top.min(down - a);
            ink_bottom = ink_bottom.max(down + d);
        }
        limits_right = width;
    } else {
        let (sup_ascent,sup_descent)=sup.map_or((0.0,0.0),|expr|script_position_extents(expr,&script_ctx,cramped));
        let (sub_ascent,sub_descent)=sub.map_or((0.0,0.0),|expr|script_position_extents(expr,&script_ctx,true));
        let extended=plan.assembled || plan.placements.iter().any(|p|p.glyph.extended_shape);
        let mut up = table.du_to_pt(if cramped { table.constants.SuperscriptShiftUpCramped }
            else { table.constants.SuperscriptShiftUp }, fs);
        let mut down = table.du_to_pt(table.constants.SubscriptShiftDown, fs);
        if sup.is_some() {
            up=up.max(sup_descent+table.du_to_pt(table.constants.SuperscriptBottomMin,fs));
            if extended {up=up.max(-ink_top-table.du_to_pt(table.constants.SuperscriptBaselineDropMax,fs));}
        }
        if sub.is_some() {
            down=down.max(sub_ascent-table.du_to_pt(table.constants.SubscriptTopMax,fs));
            if extended {down=down.max(ink_bottom+table.du_to_pt(table.constants.SubscriptBaselineDropMin,fs));}
        }
        if sub.is_some() && sup.is_some() {
            let minimum_gap = if !extended {
                match (
                    ordinary_fallback_rule_gap(sub.expect("joint lower limit"), &script_ctx),
                    ordinary_fallback_rule_gap(sup.expect("joint upper limit"), &script_ctx),
                ) {
                    (Some(lower), Some(upper)) => lower.max(upper),
                    _ => table.du_to_pt(table.constants.SubSuperscriptGapMin, fs),
                }
            } else {
                table.du_to_pt(table.constants.SubSuperscriptGapMin, fs)
            };
            let deficit=(minimum_gap
                -(up+down-sub_ascent-sup_descent)).max(0.0);
            let upper_bottom_limit = table.du_to_pt(
                table.constants.SuperscriptBottomMaxWithSubscript, fs) + sup_descent;
            // An extended operator has already placed its upper limit beyond
            // the ordinary-base threshold. Balance additional joint clearance
            // around that placement instead of sending it all below the axis.
            let raised = if extended && up >= upper_bottom_limit {
                deficit * 0.5
            } else {
                deficit.min((upper_bottom_limit - up).max(0.0))
            };
            up+=raised;down+=deficit-raised;
        }
        // Ascent of the upper and descent of the lower affect the reserved
        // box, while the opposite sides constrain collision clearance.
        let _=(sup_ascent,sub_descent);
        let mut right = advance;
        if let (Some(expr), Some(bbox)) = (sup, sup_box.as_ref()) {
            let (a,d) = ink_extent_word(expr,&script_ctx,cramped);
            sup_position = Some((advance,-up));
            right = right.max(advance+bbox.advance);
            ink_top = ink_top.min(-up-a);
            ink_bottom = ink_bottom.max(-up+d);
        }
        if let (Some(expr), Some(bbox)) = (sub, sub_box.as_ref()) {
            let (a,d) = ink_extent_word(expr,&script_ctx,true);
            let x = advance-italic;
            sub_position = Some((x,down));
            right = right.max(x+bbox.advance);
            ink_top = ink_top.min(down-a);
            ink_bottom = ink_bottom.max(down+d);
        }
        limits_right = right;
    }
    // Preserve the existing operand spacing policy while isolating the
    // glyph/anchor correction. Its remaining Word residual is measured separately.
    let operand_x = limits_right + fs * 0.1;
    let operand_box = layout_expr(operand,ctx);
    ink_top = ink_top.min(-pa);
    ink_bottom = ink_bottom.max(pd);
    let mut line_top=ink_top;let mut line_bottom=ink_bottom;
    for (expr,position) in [(sub,sub_position),(sup,sup_position)] {
        if let (Some(expr),Some((dx,dy)))=(expr,position) {
            let (elements,_)=emit_expr(expr,dx,dy,&script_ctx);
            for element in &elements {
                if is_fallback_font_element(element) {
                    line_top=line_top.min(element.y);line_bottom=line_bottom.max(element.y+element.height);
                }
            }
        }
    }
    let bbox = MathBBox { advance: operand_x + operand_box.advance,
        ascent: (-line_top).max(operand_box.ascent),
        descent: line_bottom.max(operand_box.descent), italic_correction: operand_box.italic_correction };
    Some(NaryGeometry { plan, scale, operator_baseline, operator_x,
        sub_position, sup_position, operand_x, bbox, ink_top, ink_bottom })
}

fn emit_nary_geometry(
    op: char, sub: Option<&MathExpr>, sup: Option<&MathExpr>, operand: &MathExpr,
    lim_loc: crate::ir::LimLoc, grow: bool, x: f32, baseline: f32, ctx: &MathLayoutContext,
    operator_color: Option<&str>,
) -> Option<(Vec<LayoutElement>, MathBBox)> {
    let shape=nary_geometry(op,sub,sup,operand,lim_loc,grow,ctx,false)?;
    let fs=ctx.effective_font_size();
    let mut elements=Vec::new();
    for placement in &shape.plan.placements {
        let glyph_baseline = baseline + shape.operator_baseline
            - placement.advance_offset as f32 * shape.scale;
        let mut element = emit_text_at(op.to_string(), x + shape.operator_x, glyph_baseline, fs);
        if let LayoutContent::Text { color, .. } = &mut element.content {
            *color = operator_color.map(str::to_owned);
        }
        element.width = placement.glyph.advance_width as f32 * shape.scale;
        element.font_glyph = Some(FontGlyph { index: placement.glyph.gid,
            bounds_em: placement.glyph.bounds.map(|v| v as f32 / StretchTable::cambria_math().upm as f32) });
        elements.push(element);
    }
    let script_ctx=ctx.descend_script();
    if let (Some(expr),Some((dx,dy)))=(sup,shape.sup_position) {
        elements.extend(emit_expr(expr,x+dx,baseline+dy,&script_ctx).0);
    }
    if let (Some(expr),Some((dx,dy)))=(sub,shape.sub_position) {
        elements.extend(emit_expr(expr,x+dx,baseline+dy,&script_ctx).0);
    }
    elements.extend(emit_expr(operand,x+shape.operand_x,baseline,ctx).0);
    Some((elements,shape.bbox))
}

struct ResolvedRunGlyph {
    character: char,
    metrics: crate::font::CatalogGlyphMetrics,
}

fn math_contact_kern(expr:&MathExpr,ctx:&MathLayoutContext,corner:crate::font::math_kern::Corner,height:f32)->f32 {
    if let MathExpr::Seq(children)=expr {
        if children.len()==1 {return math_contact_kern(&children[0],ctx,corner,height);}
    }
    if let MathExpr::Run {text,style}=expr {
        if let Some(run)=style.run_style.as_ref().filter(|run|run.font_family.is_some()) {
            if let Some(glyphs)=resolved_run_glyphs(text,style,ctx) {
                if glyphs.len()==1 {
                    return crate::font::catalog_glyph_corner_kern(run.font_family.as_deref().unwrap(),run.bold,run.italic,
                        glyphs[0].metrics.index,corner,height,resolved_run_context(style,ctx).effective_font_size());
                }
            }
            return 0.0;
        }
    }
    crate::font::math_kern::MathKerns::cambria_math().value(kern_leaf_gid(expr,ctx),corner,height,ctx.effective_font_size())
}

fn resolved_run_context(style: &crate::ir::MathRunStyle, ctx: &MathLayoutContext) -> MathLayoutContext {
    let size = style.run_style.as_ref().and_then(|run|run.font_size)
        .filter(|size|size.is_finite() && *size > 0.0).unwrap_or(ctx.font_size);
    MathLayoutContext { font_size: size, style: ctx.style }
}

fn resolved_run_glyphs(text: &str, style: &crate::ir::MathRunStyle, ctx: &MathLayoutContext)
    -> Option<Vec<ResolvedRunGlyph>>
{
    let run = style.run_style.as_ref()?;
    let family = run.font_family.as_deref()?;
    let level = run_script_level(style, ctx);
    text.chars().map(|c| {
        let original = crate::font::catalog_glyph_metrics(family, run.bold, run.italic, c, 0)?;
        let character = if original.has_math {
            crate::font::math_substitute::math_run_substitute(c, style)
        } else { c };
        let metrics = crate::font::catalog_glyph_metrics(family, run.bold, run.italic, character, level)?;
        Some(ResolvedRunGlyph { character, metrics })
    }).collect()
}

/// Non-MATH leaves have a font box independent of their painted outline.
/// Limit positioning uses the effective script size and external leading;
/// line reservation retains the source size without that leading.
fn fallback_run_font_box(
    glyphs: &[ResolvedRunGlyph], style: &crate::ir::MathRunStyle,
    ctx: &MathLayoutContext, nominal: bool, leading: bool,
) -> Option<(f32, f32)> {
    if glyphs.is_empty() || glyphs.iter().all(|g| g.metrics.has_math) { return None; }
    let run=style.run_style.as_ref()?;
    let metrics=crate::font::catalog_glyph_face_metrics(run.font_family.as_deref()?,run.bold,run.italic)?;
    let context=resolved_run_context(style,ctx);
    // A nominal-size reservation applies when the argument is actually
    // reduced. Full-size leaves in fractions keep their painted ink box.
    if nominal && context.effective_font_size()>=context.font_size {return None;}
    let size=if nominal {context.font_size} else {context.effective_font_size()};
    Some(metrics.design_font_box_pt(size,leading))
}

fn script_position_extents(expr:&MathExpr,ctx:&MathLayoutContext,cramped:bool)->(f32,f32) {
    match expr {
        MathExpr::Run {text,style}=> {
            if let Some(glyphs)=resolved_run_glyphs(text,style,ctx) {
                if let Some(extents)=fallback_run_font_box(&glyphs,style,ctx,false,true) {return extents;}
            }
        },
        MathExpr::Seq(children)=> {
            return children.iter().map(|child|script_position_extents(child,ctx,cramped))
                .fold((0.0_f32,0.0_f32),|(a,d),(ca,cd)|(a.max(ca),d.max(cd)));
        },
        MathExpr::BoxExpr(child)|MathExpr::Phantom(child)=>return script_position_extents(child,ctx,cramped),
        _=>{},
    }
    ink_extent_word(expr,ctx,cramped)
}

fn is_fallback_font_element(element:&LayoutElement)->bool {
    if element.font_glyph.is_none() {return false;}
    let LayoutContent::Text {text,font_family:Some(family),bold,italic,..}=&element.content else{return false;};
    let non_math=text.chars().next().and_then(|c|crate::font::catalog_glyph_metrics(family,*bold,*italic,c,0))
        .is_some_and(|glyph|!glyph.has_math);
    if !non_math {return false;}
    let (ink_top,ink_bottom)=painted_element_ink(element);
    // A font reservation contributes space outside the glyph outline.
    // Exact ink rectangles on full-size leaves do not select this policy.
    element.y<ink_top-0.0001 || element.y+element.height>ink_bottom+0.0001
}

/// A script's font box can exceed its ink even though the glyph is small.
/// Preserve the established ink policy when all leaves use MATH faces.
pub(crate) fn inline_math_typographic_extent(block:&MathBlock,font_size:f32)->Option<(f32,f32)> {
    let (elements,bbox)=emit_math_block(block,0.0,0.0,font_size);
    if !elements.iter().any(is_fallback_font_element) {return None;}
    let baseline=bbox.ascent.max(font_size*0.8);
    let (top,bottom)=reserved_line_extents(&elements);
    if !top.is_finite() || !bottom.is_finite() {return None;}
    let (ia,id)=inline_math_ink_extent(block,font_size);
    Some(((baseline-top).max(ia).max(0.0),(bottom-baseline).max(id).max(0.0)))
}


/// Placement extents are distinct from the nominal font box used to count
/// grid cells. A reduced fallback leaf retains its source ascent reservation,
/// while its effective font descent includes that face's external leading.
/// MATH-only expressions keep their existing placement policy.
pub(crate) fn inline_math_baseline_extent(block: &MathBlock, font_size: f32)
    -> Option<(f32, f32)>
{
    let (elements, bbox) = emit_math_block(block, 0.0, 0.0, font_size);
    if !elements.iter().any(is_fallback_font_element) { return None; }
    let baseline = bbox.ascent.max(font_size * 0.8);
    let (top, mut bottom) = reserved_line_extents(&elements);
    for element in elements.iter().filter(|e| is_fallback_font_element(e)) {
        if let LayoutContent::Text {font_family: Some(family), font_size: size,
            bold, italic, ..} = &element.content
        {
            if let Some(metrics) = crate::font::catalog_glyph_face_metrics(family, *bold, *italic) {
                let (_, descent) = metrics.design_font_box_pt(*size, true);
                if let Some(offset) = element.baseline_offset {
                    bottom = bottom.max(element.y + offset + descent);
                }
            }
        }
    }
    if !top.is_finite() || !bottom.is_finite() { return None; }
    let (ink_ascent, ink_descent) = inline_math_ink_extent(block, font_size);
    Some(((baseline - top).max(ink_ascent).max(0.0),
        (bottom - baseline).max(ink_descent).max(0.0)))
}

fn resolved_run_bbox(glyphs: &[ResolvedRunGlyph], style: &crate::ir::MathRunStyle, ctx: &MathLayoutContext) -> MathBBox {
    let fs=resolved_run_context(style,ctx).effective_font_size();
    glyphs.iter().map(|g|MathBBox { advance:g.metrics.advance_em*fs,
        ascent:g.metrics.bounds_em[3].max(0.0)*fs,
        descent:(-g.metrics.bounds_em[1]).max(0.0)*fs,
        italic_correction:g.metrics.italic_correction_em*fs })
        .fold(MathBBox::default(),|a,b|a.hstack(&b))
}

fn emit_resolved_run(glyphs: &[ResolvedRunGlyph], style: &crate::ir::MathRunStyle,
    x:f32, baseline:f32, ctx:&MathLayoutContext) -> Vec<LayoutElement>
{
    let run=style.run_style.as_ref().expect("resolved run has a generic style");
    let fs=resolved_run_context(style,ctx).effective_font_size();
    let mut pen=x;
    glyphs.iter().map(|g| {
        let mut element=emit_text_at(g.character.to_string(),pen,baseline,fs);
        element.width=g.metrics.advance_em*fs;
        if let Some((ascent,descent))=fallback_run_font_box(glyphs,style,ctx,true,false) {
            element.y=baseline-ascent; element.height=ascent+descent;
            element.baseline_offset=Some(ascent);
        }else if !g.metrics.has_math {
            // Exact ink boxes apply to full-size fallback leaves. Preserve
            // the established MATH-font element geometry used by fractions.
            let ascent=g.metrics.bounds_em[3]*fs;
            let descent=-g.metrics.bounds_em[1]*fs;
            element.y=baseline-ascent;element.height=(ascent+descent).max(0.0);
            element.baseline_offset=Some(ascent);
        }
        element.font_glyph=Some(FontGlyph { index:g.metrics.index, bounds_em:g.metrics.bounds_em });
        if let LayoutContent::Text { font_family,bold,italic,color,underline,strikethrough,double_strikethrough,.. }=&mut element.content {
            *font_family=run.font_family.clone(); *bold=run.bold; *italic=run.italic;
            *color=run.color.clone(); *underline=run.underline; *strikethrough=run.strikethrough;
            *double_strikethrough=run.double_strikethrough;
        }
        pen+=element.width;
        element
    }).collect()
}

/// Transform every run in every primitive; the author document remains
/// separate from the layout copy that receives resolved font families.
pub(super) fn map_math_runs(block:&mut MathBlock,
    mapper:&mut impl FnMut(&str,&crate::ir::MathRunStyle)->Vec<MathExpr>)
{
    fn visit(expr:&mut MathExpr, mapper:&mut impl FnMut(&str,&crate::ir::MathRunStyle)->Vec<MathExpr>) {
        match expr {
            MathExpr::Run {text,style}=> {
                let mut replacement=mapper(text,style);
                *expr=if replacement.len()==1 { replacement.remove(0) }else{MathExpr::Seq(replacement)};
            }
            MathExpr::Text(_)=>{},
            MathExpr::Seq(children)|MathExpr::EqArray(children)=>for child in children { visit(child,mapper); },
            MathExpr::Fraction {num,den,..}=> {visit(num,mapper);visit(den,mapper);},
            MathExpr::Superscript {base,sup}=> {visit(base,mapper);visit(sup,mapper);},
            MathExpr::Subscript {base,sub}=> {visit(base,mapper);visit(sub,mapper);},
            MathExpr::SubSuperscript {base,sub,sup}|MathExpr::PreScript {base,sub,sup}=> {visit(base,mapper);visit(sub,mapper);visit(sup,mapper);},
            MathExpr::Radical {degree,radicand}=> {if let Some(degree)=degree {visit(degree,mapper);}visit(radicand,mapper);},
            MathExpr::Nary {sub,sup,operand,..}=> {if let Some(sub)=sub {visit(sub,mapper);}if let Some(sup)=sup {visit(sup,mapper);}visit(operand,mapper);},
            MathExpr::Delimiter {content,..}=>visit(content,mapper),
            MathExpr::Function {name,arg}=> {visit(name,mapper);visit(arg,mapper);},
            MathExpr::Matrix {rows,..}=>for row in rows {for cell in row {visit(cell,mapper);}},
            MathExpr::Accent {base,..}|MathExpr::Bar {base,..}|MathExpr::GroupChar {base,..}|MathExpr::BorderBox {base,..}=>visit(base,mapper),
            MathExpr::Limit {base,lim,..}=> {visit(base,mapper);visit(lim,mapper);},
            MathExpr::BoxExpr(base)|MathExpr::Phantom(base)=>visit(base,mapper),
        }
    }
    let content=match block { MathBlock::Inline(content)|MathBlock::Display {content,..}=>content };
    for expr in content {visit(expr,mapper);}
}

fn run_text_bbox(text: &str, style: &crate::ir::MathRunStyle, ctx: &MathLayoutContext) -> MathBBox {
    if let Some(glyphs)=resolved_run_glyphs(text,style,ctx) {return resolved_run_bbox(&glyphs,style,ctx);}
    let eff=ctx.effective_font_size();let level=run_script_level(style,ctx);
    text.chars().map(|c| {
        let cp=crate::font::math_substitute::math_run_substitute(c,style);
        MathBBox { advance:eff*selected_advance_em(cp,level),
            ascent:eff*0.7,descent:eff*0.2,
            italic_correction:selected_italic_correction(cp,level,eff) }
    }).fold(MathBBox::default(),|a,b|a.hstack(&b))
}

fn run_ink(text: &str, style: &crate::ir::MathRunStyle, ctx: &MathLayoutContext) -> (f32,f32) {
    if let Some(glyphs)=resolved_run_glyphs(text,style,ctx) {let bbox=resolved_run_bbox(&glyphs,style,ctx);return (bbox.ascent,bbox.descent);}
    let eff=ctx.effective_font_size();let level=run_script_level(style,ctx);
    text.chars().map(|c| {
        let cp=crate::font::math_substitute::math_run_substitute(c,style);
        let(a,d)=selected_ink_em(cp,level);(a*eff,d*eff)
    }).fold((0.0f32,0.0f32),|(a,d),(ca,cd)|(a.max(ca),d.max(cd)))
}

pub fn leaf_char_bbox(c: char, ctx: &MathLayoutContext) -> MathBBox {
    let eff=ctx.effective_font_size();let cp=math_substitute(c);let level=script_level(ctx);
    MathBBox { advance:eff*selected_advance_em(cp,level),
        ascent:eff*0.7,descent:eff*0.2,
        italic_correction:selected_italic_correction(cp,level,eff) }
}

/// S1596 (2026-09-29): Cambria Math glyph ink extents (design units), for the
/// fraction gap test below. Extracted from the installed face's outlines.
fn glyph_ink_du(c: char) -> Option<(f32, f32)> {
    if crate::font::runtime::resolve_registered("Cambria Math", false, false).is_some() {
        return crate::font::runtime::registered_glyph("Cambria Math", false, false, c, 0)
            .map(|g| (g.bounds_em[3], -g.bounds_em[1]));
    }
    static T: std::sync::OnceLock<(f32, std::collections::HashMap<u32, (f32, f32)>)> = std::sync::OnceLock::new();
    let (upm, map) = T.get_or_init(|| {
        let v: serde_json::Value = serde_json::from_str(include_str!("../font/data/cambria_math_glyph_heights.json"))
            .expect("embedded Cambria Math glyph heights should be valid JSON");
        let upm = v["upm"].as_f64().unwrap_or(2048.0) as f32;
        let mut m = std::collections::HashMap::new();
        if let Some(h) = v["heights"].as_object() {
            for (k, val) in h {
                if let (Ok(cp), Some(a)) = (k.parse::<u32>(), val.as_array()) {
                    let top = a.first().and_then(|x| x.as_f64()).unwrap_or(0.0) as f32;
                    let bot = a.get(1).and_then(|x| x.as_f64()).unwrap_or(0.0) as f32;
                    m.insert(cp, (top, bot));
                }
            }
        }
        (upm, m)
    });
    map.get(&(c as u32)).map(|(t, b)| (t / upm, -b / upm))
}

/// Ink (ascent, descent) of an expression above/below its baseline, in points.
/// Leaves use the glyph outlines; structures follow the same composition as
/// `layout_expr`; anything else falls back to the layout box.
fn ink_extent(expr: &MathExpr, ctx: &MathLayoutContext) -> (f32, f32) {
    let table = MathTable::cambria_math();
    match expr {
        MathExpr::Run { text, style } => run_ink(text, style, ctx),
        MathExpr::Text(t) => {
            let eff = ctx.effective_font_size();
            let mut a = 0.0f32;
            let mut d = 0.0f32;
            for c in t.chars() {
                let (ga, gd) = selected_ink_em(math_substitute(c),script_level(ctx));
                a = a.max(ga * eff);
                d = d.max(gd * eff);
            }
            (a, d)
        }
        MathExpr::Seq(children) => children.iter().map(|c| ink_extent(c, ctx))
            .fold((0.0f32, 0.0f32), |(a, d), (ca, cd)| (a.max(ca), d.max(cd))),
        MathExpr::Fraction { num, den, .. } => {
            let sub_ctx = ctx.descend_fraction();
            let fs = ctx.font_size;
            let (up_du, down_du) = if ctx.style.is_display() {
                (table.constants.FractionNumeratorDisplayStyleShiftUp, table.constants.FractionDenominatorDisplayStyleShiftDown)
            } else {
                (table.constants.FractionNumeratorShiftUp, table.constants.FractionDenominatorShiftDown)
            };
            let (up, down) = fraction_shifts(&table, fs, ctx.style.is_display(),
                table.du_to_pt(up_du, fs), table.du_to_pt(down_du, fs), num, den, &sub_ctx);
            let (na, _) = ink_extent(num, &sub_ctx);
            let (_, dd) = ink_extent(den, &sub_ctx);
            (up + na, down + dd)
        }
        MathExpr::Radical { radicand, .. } => {
            let (ra, rd) = ink_extent(radicand, ctx);
            let fs = ctx.font_size;
            let gap_du = table.constants.RadicalRuleThickness;
            let gap = table.du_to_pt(gap_du, fs);
            let thk = table.du_to_pt(table.constants.RadicalRuleThickness, fs);
            (ra + gap + thk, rd)
        }
        MathExpr::Superscript { base, sup } => {
            let (ba, bd) = ink_extent(base, ctx);
            let (sa, _) = ink_extent(sup, &ctx.descend_script());
            let up = table.du_to_pt(table.constants.SuperscriptShiftUp, ctx.effective_font_size());
            (ba.max(sa + up), bd)
        }
        MathExpr::Subscript { base, sub } => {
            let (ba, bd) = ink_extent(base, ctx);
            let (_, sd) = ink_extent(sub, &ctx.descend_script());
            let dn = table.du_to_pt(table.constants.SubscriptShiftDown, ctx.effective_font_size());
            (ba, bd.max(sd + dn))
        }
        _ => {
            let b = layout_expr(expr, ctx);
            (b.ascent, b.descent)
        }
    }
}

/// S1596 (2026-09-29, default ON, opt-out OXI_S1596_DISABLE): a fraction's
/// numerator/denominator shifts are the MATH constants OR whatever keeps the
/// minimum gap between the bar and the numerator's / denominator's INK,
/// whichever is larger (OpenType MATH FractionNumerator/DenominatorGapMin,
/// measured from the bar edges around the axis). MEASURED (`_pb_cjkmath_gen.py`,
/// educational__20d9968b slice, `lines` grid 18pt, Word PDF glyph boxes):
/// b / sqrt(a^2+b^2) spans two grid cells where a/b spans one; in (a/b)/(c/d) the
/// denominator sits 5.40 below the baseline = gap 0.68 + (inner shift 4.43 +
/// ink of `c` at 6pt 2.94) - (axis 3.0 - half bar 0.34), not the constant 5.28
/// and not a 0.7em box (which would give 3 cells).
fn fraction_shifts(table: &MathTable, fs: f32, display: bool, up: f32, down: f32,
                   num: &MathExpr, den: &MathExpr, sub_ctx: &MathLayoutContext) -> (f32, f32) {
    if std::env::var_os("OXI_S1596_DISABLE").is_some() {
        return (up, down);
    }
    let axis = table.du_to_pt(table.constants.AxisHeight, fs);
    let half = table.du_to_pt(table.constants.FractionRuleThickness, fs) / 2.0;
    let (num_gap, den_gap) = if display {
        (table.constants.FractionNumDisplayStyleGapMin, table.constants.FractionDenomDisplayStyleGapMin)
    } else {
        (table.constants.FractionNumeratorGapMin, table.constants.FractionDenominatorGapMin)
    };
    let num_gap = table.du_to_pt(num_gap, fs);
    let den_gap = table.du_to_pt(den_gap, fs);
    // Fraction gaps are measured against glyph ink, including nested
    // delimiters and scripts. Layout boxes reserve additional space and
    // must not inflate the numerator drop or denominator rise.
    let (_, nd) = fraction_child_extents(num, sub_ctx, false);
    let (da, _) = fraction_child_extents(den, sub_ctx, true);
    (up.max(axis + half + num_gap + nd), down.max(den_gap + da - (axis - half)))
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
    let row = math_row_atoms(exprs);
    let exprs = row.as_ref();
    let gaps = atom_gaps(exprs, font_size); // S527 inter-atom math-class spacing
    let mut acc = MathBBox::default();
    for (i, e) in exprs.iter().enumerate() {
        acc.advance += gaps[i];
        let b = layout_expr(e, &ctx);
        acc = acc.hstack(&b);
    }
    acc
}

/// S527 (coverage): TeX-style math-class inter-atom spacing. A math run gets
/// extra space around relation operators (=<>≤≥→…, "thick" 5/18 em) and binary
/// operators (+−×÷±…, "medium" 4/18 em); large operators get "thin" (3/18 em).
/// Without this, `E=m` / `x+y` render with no operator spacing (too narrow).
#[derive(Clone, Copy, PartialEq, Eq)]
enum AClass { Ord, Bin, Rel, Op, Open, Close, Punct }

fn classify_math_char(c: char) -> AClass {
    match c {
        '=' | '<' | '>' | '≠' | '≤' | '≥' | '≈' | '≡' | '∼' | '≅' | '∝'
        | '∈' | '∉' | '∋' | '⊂' | '⊃' | '⊆' | '⊇' | '≪' | '≫' | '≐' | '≑'
        | '→' | '←' | '↔' | '⇒' | '⇐' | '⇔' | '↦' | '≔' | '≜' => AClass::Rel,
        '+' | '-' | '\u{2212}' | '±' | '∓' | '×' | '÷' | '⋅' | '∗' | '∘' | '∙'
        | '∪' | '∩' | '∨' | '∧' | '⊕' | '⊗' | '⊖' | '⊙' | '⊎' | '⊓' | '⊔' => AClass::Bin,
        '(' | '[' | '{' | '⟨' | '⌊' | '⌈' => AClass::Open,
        ')' | ']' | '}' | '⟩' | '⌋' | '⌉' => AClass::Close,
        ',' | ';' => AClass::Punct,
        _ => AClass::Ord,
    }
}

fn classify_atom(e: &MathExpr) -> AClass {
    match e {
        MathExpr::Text(s) | MathExpr::Run { text: s, .. } => {
            let mut it = s.chars();
            match (it.next(), it.next()) {
                (Some(c), None) => classify_math_char(c),
                _ => AClass::Ord, // multi-char identifier
            }
        }
        MathExpr::Nary { .. } => AClass::Op,
        _ => AClass::Ord,
    }
}

/// Per-atom leading gap for a sequence (gap before atom i; gap[0]=0). Reclassifies
/// a binary op to Ord (no space) when it is unary (first, or after Bin/Rel/Op/Open/Punct).
fn atom_gaps(children: &[MathExpr], fs: f32) -> Vec<f32> {
    let thin = fs * 3.0 / 18.0;
    let med = fs * 4.0 / 18.0;
    let thick = fs * 5.0 / 18.0;
    let mut gaps = vec![0.0_f32; children.len()];
    let mut prev: Option<AClass> = None;
    for (i, c) in children.iter().enumerate() {
        let mut cls = classify_atom(c);
        if cls == AClass::Bin {
            let unary = matches!(prev, None | Some(AClass::Bin) | Some(AClass::Rel)
                | Some(AClass::Op) | Some(AClass::Open) | Some(AClass::Punct));
            if unary { cls = AClass::Ord; }
        }
        if let Some(p) = prev {
            gaps[i] = match (p, cls) {
                (AClass::Rel, _) | (_, AClass::Rel) => thick,
                (AClass::Bin, _) | (_, AClass::Bin) => med,
                (AClass::Op, _) | (_, AClass::Op) => thin,
                (AClass::Punct, _) => thin,
                _ => 0.0,
            };
        }
        prev = Some(cls);
    }
    gaps
}

/// A run boundary is formatting, not a mathematical operator boundary.
/// Keep ordinary identifiers together for shaping, and expose operators and
/// punctuation so row spacing sees both sides even in a run such as "+(".
fn math_run_atoms(expr: &MathExpr) -> Option<Vec<MathExpr>> {
    let (text, style) = match expr {
        MathExpr::Text(text) => (text.as_str(), None),
        MathExpr::Run { text, style } if !style.literal => (text.as_str(), Some(style)),
        _ => return None,
    };
    if text.chars().count() < 2 || !text.chars().any(|c| classify_math_char(c) != AClass::Ord) {
        return None;
    }
    let leaf = |text: String| match style {
        Some(style) => MathExpr::Run { text, style: style.clone() },
        None => MathExpr::Text(text),
    };
    let mut atoms = Vec::new();
    let mut identifier = String::new();
    for c in text.chars() {
        if classify_math_char(c) == AClass::Ord {
            identifier.push(c);
        } else {
            if !identifier.is_empty() {
                atoms.push(leaf(std::mem::take(&mut identifier)));
            }
            atoms.push(leaf(c.to_string()));
        }
    }
    if !identifier.is_empty() { atoms.push(leaf(identifier)); }
    Some(atoms)
}

fn math_row_atoms(children: &[MathExpr]) -> std::borrow::Cow<'_, [MathExpr]> {
    if !children.iter().any(|e| math_run_atoms(e).is_some()) {
        return std::borrow::Cow::Borrowed(children);
    }
    let mut atoms = Vec::new();
    for child in children {
        match math_run_atoms(child) {
            Some(parts) => atoms.extend(parts),
            None => atoms.push(child.clone()),
        }
    }
    std::borrow::Cow::Owned(atoms)
}

/// Dispatch bbox computation by MathExpr variant. Phase 2 implements
/// only leaf cases; Phase 3 adds the full primitive set.
pub fn layout_expr(expr: &MathExpr, ctx: &MathLayoutContext) -> MathBBox {
    if let Some(atoms) = math_run_atoms(expr) {
        return layout_expr(&MathExpr::Seq(atoms), ctx);
    }
    match expr {
        MathExpr::Text(s) => leaf_text_bbox(s, ctx),
        MathExpr::Run { text, style } => run_text_bbox(text, style, ctx),
        MathExpr::Seq(children) => {
            let row = math_row_atoms(children);
            let children = row.as_ref();
            let gaps = atom_gaps(children, ctx.font_size);
            let mut acc = MathBBox::default();
            for (i, c) in children.iter().enumerate() {
                acc.advance += gaps[i]; // S527 inter-atom math-class spacing
                acc = acc.hstack(&layout_expr(c, ctx));
            }
            acc
        }
        // Phase 3: full recursive layout for these primitives.
        MathExpr::Fraction { num, den, .. } => {
            let sub_ctx = ctx.descend_fraction();
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
            let (num_shift_up, den_shift_down) = fraction_shifts(
                &table, fs, ctx.style.is_display(),
                table.du_to_pt(num_shift_du, fs), table.du_to_pt(den_shift_du, fs), num, den, &sub_ctx);
            MathBBox {
                advance: (nb.advance + 2.0 * fraction_argument_side_space(num, ctx))
                    .max(db.advance + 2.0 * fraction_argument_side_space(den, ctx)),
                ascent: num_shift_up + nb.ascent,
                descent: den_shift_down + db.descent,
                italic_correction: 0.0,
            }
        }
        MathExpr::Superscript { base, sup } => {
            let bb = layout_expr(base, ctx);
            let sb = layout_expr(sup, &ctx.descend_script());
            let table = MathTable::cambria_math();
            let shift_up = table.du_to_pt(table.constants.SuperscriptShiftUp, ctx.effective_font_size());
            MathBBox {
                advance: bb.advance + bb.italic_correction + script_kern(base,sup,ctx,shift_up,true) + sb.advance + space_after_script(ctx),
                ascent: bb.ascent.max(sb.height() + shift_up),
                descent: bb.descent,
                italic_correction: sb.italic_correction,
            }
        }
        MathExpr::Subscript { base, sub } => {
            let bb = layout_expr(base, ctx);
            let sb = layout_expr(sub, &ctx.descend_script());
            let table = MathTable::cambria_math();
            let shift_down = table.du_to_pt(table.constants.SubscriptShiftDown, ctx.effective_font_size());
            MathBBox {
                advance: bb.advance + script_kern(base,sub,ctx,shift_down,false) + sb.advance + space_after_script(ctx),
                ascent: bb.ascent,
                descent: bb.descent.max(sb.height() + shift_down),
                italic_correction: sb.italic_correction,
            }
        }
        MathExpr::SubSuperscript { base, sub, sup } => {
            let bb = layout_expr(base, ctx);
            let super_b = layout_expr(sup, &ctx.descend_script());
            let sub_b = layout_expr(sub, &ctx.descend_script());
            let (sup_shift, sub_shift) = combined_script_shifts(sub, sup, ctx, false);
            MathBBox {
                advance: bb.advance + bb.italic_correction
                    + (super_b.advance+script_kern(base,sup,ctx,sup_shift,true))
                        .max(sub_b.advance+script_kern(base,sub,ctx,sub_shift,false)) + space_after_script(ctx),
                ascent: bb.ascent.max(super_b.height() + sup_shift),
                descent: bb.descent.max(sub_b.height() + sub_shift),
                italic_correction: 0.0,
            }
        }
        MathExpr::Radical { degree, radicand } => {
            let rb = layout_expr(radicand, ctx);
            if let Some(shape)=radical_geometry(radicand,ctx) {
                let (prefix,placement)=radical_degree(degree.as_deref(),ctx,&shape);
                let mut top=shape.ink_top;
                let mut bottom=shape.ink_bottom;
                if let (Some(expr),Some((_,baseline,dctx)))=(degree.as_deref(),placement) {
                    let(a,d)=ink_extent_word(expr,&dctx,false);top=top.min(baseline-a);bottom=bottom.max(baseline+d);
                }
                let extra=MathTable::cambria_math().du_to_pt(MathTable::cambria_math().constants.RadicalExtraAscender,ctx.effective_font_size());
                return MathBBox{advance:prefix+shape.advance+rb.advance,
                    ascent:(-top).max(0.0)+extra,descent:bottom.max(0.0),italic_correction:0.0};
            }
            let table = MathTable::cambria_math();
            let fs = ctx.font_size;
            let gap_du = table.constants.RadicalRuleThickness;
            let gap = table.du_to_pt(gap_du, fs);
            let thk = table.du_to_pt(table.constants.RadicalRuleThickness, fs);
            let extra = table.du_to_pt(table.constants.RadicalExtraAscender, fs);
            // S528: √ sign width scales only for a TALL radicand (else base size).
            let rad_h = rb.ascent + rb.descent;
            let sign_fs = if rad_h > fs * 1.2 { (rad_h + gap + thk).max(fs) } else { fs };
            MathBBox {
                advance: rb.advance + sign_fs * 0.55,
                ascent: rb.ascent + gap + thk + extra,
                descent: rb.descent,
                italic_correction: 0.0,
            }
        }
        MathExpr::Delimiter { beg, end, content, .. } => {
            let cb = layout_expr(content, ctx);
            let (la, ld) = delimiter_ink(*beg, content, ctx, false);
            let (ra, rd) = delimiter_ink(*end, content, ctx, false);
            MathBBox {
                advance: cb.advance + delimiter_width(*beg, content, ctx) + delimiter_width(*end, content, ctx),
                ascent: cb.ascent.max(la).max(ra),
                descent: cb.descent.max(ld).max(rd),
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
        MathExpr::Accent { accent, base } => {
            if let Some(shape)=accent_geometry(*accent,base,ctx) {
                let mut bbox=layout_expr(base,ctx);
                bbox.ascent=bbox.ascent.max(-shape.ink_top);
                bbox.descent=bbox.descent.max(shape.ink_bottom);
                return bbox;
            }
            let bb = layout_expr(base, ctx);
            let table = MathTable::cambria_math();
            let fs = ctx.font_size;
            let gap = table.du_to_pt(table.constants.OverbarVerticalGap, fs);
            // S525 (coverage): an accent (hat/bar/tilde) is a SMALL glyph just
            // above the base — reserve ~0.35×fs, not 0.9×fs (Word x̂ = 14px tall
            // vs Oxi's 28px with 0.9). Mirrors emit_accent's acc_size.
            MathBBox {
                advance: bb.advance,
                ascent: bb.ascent + gap + fs * 0.35,
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
        MathExpr::Nary { op, sub, sup, operand, lim_loc, grow, .. } => {
            if let Some(shape)=nary_geometry(*op,sub.as_deref(),sup.as_deref(),operand,*lim_loc,*grow,ctx,false) {
                return shape.bbox;
            }
            let fs = ctx.font_size;
            let op_is_integral = ('\u{222B}'..='\u{2233}').contains(op);
            // S653 (coverage): a DISPLAY integral sign is drawn EXTRA-tall —
            // Word ∫ ink ≈ fs×2.34 (26.64pt @fs=10.5, _glyphsize.py) vs the
            // generic fs×1.6 display operator (∑∏…). S525 fixed the integral's
            // limit placement (subSup) but left op_size at the generic display
            // size, so the ∫ rendered ~8.6pt too short (and S652 then honestly
            // reserved that short drawing). Keep ∑/∏ at fs×1.6 (sum reserve was
            // already −0.96, the stacked limits drive its height).
            // S1252: an INLINE (text-style) n-ary operator is drawn at the RUN
            // size. Word truth (_pb_inlmath `nary` / `nary2` arms, PDF spans):
            // `∑Board Committee` comes back as ONE CambriaMath span at size 9.96
            // — the same size as the surrounding text — and the line advance stays
            // the plain 11.66; with visible limits the operator is still 9.96 and
            // the limits 6.96, growing the line by only 0.24. reference__0042471c
            // (the ONLY corpus doc with an inline n-ary: golden 0 / corp_ja 0 /
            // corp_en 1) agrees — Word's whole `∑Dewan Pengawas Syariah (…)` span
            // is size 9.96, while Oxi drew the sigma at fs×1.2 and spent ~4pt of
            // extra line. Display keeps its enlarged operator.
            let op_size = if op_is_integral && ctx.style.is_display() {
                fs * 2.34
            } else if ctx.style.is_display() {
                fs * 1.6
            } else if std::env::var("OXI_S1252_DISABLE").is_err() {
                fs
            } else {
                fs * 1.2
            };
            // ∫ is tall-and-NARROW (advance ≈ 0.36× its height) unlike the
            // roughly-square ∑/∏ (0.6×). Decouple width from the S653 taller
            // op_size so the operand isn't pushed right (Word ∫ total width
            // 17.28pt vs Oxi 24.0 under the 0.6 factor, _glyphwh.py).
            let op_w = op_size * if op_is_integral { 0.36 } else { 0.6 };
            let lim_ctx = ctx.descend_script();
            let sub_b = sub.as_ref().map(|s| layout_expr(s, &lim_ctx));
            let sup_b = sup.as_ref().map(|s| layout_expr(s, &lim_ctx));
            let op_bbox = layout_expr(operand, ctx);
            // S525: must mirror emit_nary's effective limit location so the
            // reserved bbox matches the draw. Stacking (undOvr) reserves
            // op_size+limit above/below; integrals & inline keep subSup (limits
            // to the right -> height ≈ op_size, width += limits).
            // (op_is_integral computed above for the S653 op_size.)
            let stacked = matches!(lim_loc, crate::ir::LimLoc::UndOvr)
                || (ctx.style.is_display() && matches!(lim_loc, crate::ir::LimLoc::SubSup) && !op_is_integral);
            if stacked {
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
            } else {
                // subSup: operator + limits to the right beside it.
                let lim_w = sub_b.as_ref().map(|b| b.advance).unwrap_or(0.0)
                    .max(sup_b.as_ref().map(|b| b.advance).unwrap_or(0.0));
                MathBBox {
                    advance: op_w + lim_w + fs * 0.1 + op_bbox.advance,
                    ascent: (op_size * 0.8)
                        .max(sup_b.as_ref().map(|b| fs * 0.4 + b.height()).unwrap_or(0.0))
                        .max(op_bbox.ascent),
                    descent: (op_size * 0.2)
                        .max(sub_b.as_ref().map(|b| fs * 0.3 + b.height()).unwrap_or(0.0))
                        .max(op_bbox.descent),
                    italic_correction: 0.0,
                }
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
            // S525: mirror emit_group_chr's reduced brace size (0.5×fs, ~0.45 reserve).
            let chr_h = fs * 0.45;
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
            let pre_w = sb.advance.max(pb.advance)+space_after_script(ctx);
            let table = MathTable::cambria_math();
            let fs = ctx.effective_font_size();
            let sup_shift = table.du_to_pt(table.constants.SuperscriptShiftUp, fs);
            let sub_shift = table.du_to_pt(table.constants.SubscriptShiftDown, fs);
            MathBBox {
                advance: pre_w + bb.advance,
                ascent: bb.ascent.max(sup_shift + pb.ascent),
                descent: bb.descent.max(sub_shift + sb.descent),
                italic_correction: bb.italic_correction,
            }
        }
        // S526 (coverage): boxed/grouped wrappers.
        MathExpr::BorderBox { base, .. } => {
            let pad = ctx.font_size * 0.22;
            let bb = layout_expr(base, ctx);
            MathBBox {
                advance: bb.advance + 2.0 * pad,
                ascent: bb.ascent + pad,
                descent: bb.descent + pad,
                italic_correction: 0.0,
            }
        }
        MathExpr::BoxExpr(inner) | MathExpr::Phantom(inner) => layout_expr(inner, ctx),
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
        MathExpr::Text(s) => {
            for c in s.chars() { out.push(math_substitute(c)); }
        }
        MathExpr::Run { text, style } => {
            for c in text.chars() { out.push(crate::font::math_substitute::math_run_substitute(c, style)); }
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
    // S1258: the element's own width, from the REAL advances. It was
    // `chars * font_size * 0.55` -- a flat guess that put
    // `reference__0042471c`'s 47-character maths run at 258.5pt against
    // Word's 239.9. The table (see `MathAdvances`) sums the SUBSTITUTED
    // glyphs, which is what is drawn, and predicts 238.5.
    let approx_width: f32 = if std::env::var("OXI_S1258_DISABLE").is_err() {
        text.chars().map(|c| font_size * painted_glyph_advance_em(c)).sum()
    } else {
        text.chars().count() as f32 * font_size * 0.55
    };
    let mut element=LayoutElement::new(
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
            is_vertical: false, effects: TextEffects::default(),
        },
    );
    element.baseline_offset=Some(ascent_approx);
    element
}

/// Emit the selected `ssty` glyphs, keeping Unicode text and nominal size.
/// Runs with no alternates preserve normal shaping and do not get split.
fn emit_selected_text(text: String, x: f32, baseline: f32,
                      ctx: &MathLayoutContext, level: u8) -> Vec<LayoutElement> {
    let fs=ctx.effective_font_size();
    if !text.chars().any(|c|script_glyph(c,level).is_some()) {
        return vec![emit_text_at(text,x,baseline,fs)];
    }
    let mut elements=Vec::new();let mut pen=x;
    for c in text.chars() {
        let advance=selected_advance_em(c,level)*fs;
        let mut e=emit_text_at(c.to_string(),pen,baseline,fs);
        e.width=advance;
        if let Some(g)=script_glyph(c,level) {
            e.font_glyph=Some(FontGlyph{index:g.index,bounds_em:g.bounds_em});
        }
        elements.push(e);pen+=advance;
    }
    elements
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
    if let Some(atoms) = math_run_atoms(expr) {
        return emit_expr(&MathExpr::Seq(atoms), x, baseline_y, ctx);
    }
    let eff_size = ctx.effective_font_size();
    match expr {
        MathExpr::Run { text, style } => {
            if text.is_empty() { return (vec![], MathBBox::default()); }
            if let Some(glyphs)=resolved_run_glyphs(text,style,ctx) {
                return (emit_resolved_run(&glyphs,style,x,baseline_y,ctx),resolved_run_bbox(&glyphs,style,ctx));
            }
            let selected: String = text.chars().map(|c| crate::font::math_substitute::math_run_substitute(c, style)).collect();
            let bbox = run_text_bbox(text, style, ctx);
            let elements=emit_selected_text(selected,x,baseline_y,ctx,run_script_level(style,ctx));
            (elements,bbox)
        }
        MathExpr::Text(s) => {
            if s.is_empty() {
                return (vec![], MathBBox::default());
            }
            // Apply italic-math substitution per-char.
            let subbed: String = s.chars().map(math_substitute).collect();
            let bbox = leaf_text_bbox(s, ctx);
            let elements=emit_selected_text(subbed,x,baseline_y,ctx,script_level(ctx));
            (elements,bbox)
        }
        MathExpr::Seq(children) => {
            let row = math_row_atoms(children);
            let children = row.as_ref();
            let gaps = atom_gaps(children, ctx.font_size);
            let mut elems = Vec::new();
            let mut cur_x = x;
            let mut total = MathBBox::default();
            for (i, child) in children.iter().enumerate() {
                cur_x += gaps[i]; // S527 inter-atom math-class spacing
                total.advance += gaps[i];
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
        MathExpr::Nary { op, operator_color, sub, sup, operand, lim_loc, grow } => {
            if let Some(result)=emit_nary_geometry(*op,sub.as_deref(),sup.as_deref(),operand,*lim_loc,*grow,x,baseline_y,ctx,operator_color.as_deref()) {
                return result;
            }
            emit_nary(*op, sub.as_deref(), sup.as_deref(), operand, *lim_loc, x, baseline_y, ctx, operator_color.as_deref())
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
        // S526 (coverage): boxed equation — base + a stroked rectangle.
        MathExpr::BorderBox { base, .. } => {
            let pad = ctx.font_size * 0.22;
            let base_bb = layout_expr(base, ctx);
            let (mut elems, _) = emit_expr(base, x + pad, baseline_y, ctx);
            let rect_top = baseline_y - base_bb.ascent - pad;
            let rect_h = base_bb.ascent + base_bb.descent + 2.0 * pad;
            let rect_w = base_bb.advance + 2.0 * pad;
            let rect = LayoutElement::new(x, rect_top, rect_w, rect_h, LayoutContent::BoxRect {
                fill: None,
                stroke_color: Some("#000000".to_string()),
                stroke_width: 0.5,
                corner_radius: 0.0,
            });
            elems.insert(0, rect); // behind the glyphs
            let bbox = MathBBox {
                advance: base_bb.advance + 2.0 * pad,
                ascent: base_bb.ascent + pad,
                descent: base_bb.descent + pad,
                italic_correction: 0.0,
            };
            (elems, bbox)
        }
        MathExpr::BoxExpr(inner) | MathExpr::Phantom(inner) => {
            // BoxExpr is a transparent grouping; Phantom reserves space. Both
            // emit the inner content (phantom ink-suppression is a refinement).
            emit_expr(inner, x, baseline_y, ctx)
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
    let sub_ctx = ctx.descend_fraction();

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
    let (num_shift_up, den_shift_down) = fraction_shifts(
        &table, fs, ctx.style.is_display(),
        table.du_to_pt(num_shift_du, fs), table.du_to_pt(den_shift_du, fs), num, den, &sub_ctx);
    let rule_thick = table.du_to_pt(rule_thick_du, fs);
    let axis_height = table.du_to_pt(table.constants.AxisHeight, fs);

    // Common width: max of num and den advances.
    let common_w = (num_bbox.advance + 2.0 * fraction_argument_side_space(num, ctx))
        .max(den_bbox.advance + 2.0 * fraction_argument_side_space(den, ctx));
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
    let fs = ctx.effective_font_size();
    let (mut base_elems, base_bbox) = emit_expr(base, x, baseline_y, ctx);

    let sup_ctx = ctx.descend_script();
    let shift_up = table.du_to_pt(table.constants.SuperscriptShiftUp, fs);
    let kern=script_kern(base,sup,ctx,shift_up,true);
    let sup_x = x + base_bbox.advance + base_bbox.italic_correction + kern;
    let sup_baseline = baseline_y - shift_up;
    let (sup_elems, sup_bbox) = emit_expr(sup, sup_x, sup_baseline, &sup_ctx);

    base_elems.extend(sup_elems);
    let bbox = MathBBox {
        advance: base_bbox.advance + base_bbox.italic_correction + kern + sup_bbox.advance + space_after_script(ctx),
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
    let fs = ctx.effective_font_size();
    let (mut base_elems, base_bbox) = emit_expr(base, x, baseline_y, ctx);

    let sub_ctx = ctx.descend_script();
    let shift_down = table.du_to_pt(table.constants.SubscriptShiftDown, fs);
    let kern=script_kern(base,sub,ctx,shift_down,false);
    let sub_x = x + base_bbox.advance + kern;
    let sub_baseline = baseline_y + shift_down;
    let (sub_elems, sub_bbox) = emit_expr(sub, sub_x, sub_baseline, &sub_ctx);

    base_elems.extend(sub_elems);
    let bbox = MathBBox {
        advance: base_bbox.advance + kern + sub_bbox.advance + space_after_script(ctx),
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
    if let Some(shape)=radical_geometry(radicand,ctx) {
        let fs=ctx.effective_font_size();let scale=fs/StretchTable::cambria_math().upm as f32;
        let (prefix,degree_placement)=radical_degree(degree,ctx,&shape);
        let root_x=x+prefix;let radicand_x=root_x+shape.advance;
        let mut elements=Vec::new();
        for (i,p) in shape.plan.placements.iter().enumerate() {
            let mut e=emit_text_at(if i==0{"\u{221a}".to_string()}else{String::new()},root_x,
                baseline_y+shape.baseline_shift-p.advance_offset as f32*scale,fs);
            e.width=p.glyph.advance_width as f32*scale;
            e.font_glyph=Some(FontGlyph{index:p.glyph.gid,
                bounds_em:p.glyph.bounds.map(|v|v as f32/StretchTable::cambria_math().upm as f32)});
            elements.push(e);
        }
        elements.extend(emit_expr(radicand,radicand_x,baseline_y,ctx).0);
        let width=layout_expr(radicand,ctx).advance;
        let bar_top=baseline_y+shape.ink_top;let bar_center=bar_top+shape.rule_thickness/2.0;
        elements.push(LayoutElement::new(radicand_x,bar_top,width,shape.rule_thickness,
            LayoutContent::TableBorder{x1:radicand_x,y1:bar_center,x2:radicand_x+width,y2:bar_center,
                color:None,width:shape.rule_thickness,style:None}));
        if let (Some(expr),Some((dx,dy,dctx)))=(degree,degree_placement) {
            elements.extend(emit_expr(expr,x+dx,baseline_y+dy,&dctx).0);
        }
        let expression=MathExpr::Radical{degree:degree.map(|d|Box::new(d.clone())),radicand:Box::new(radicand.clone())};
        return (elements,layout_expr(&expression,ctx));
    }
    let table = MathTable::cambria_math();
    let fs = ctx.font_size;

    // Radicand bbox (at the same style — not script).
    let rad_bbox = layout_expr(radicand, ctx);

    // MATH constants (select display vs inline gap).
    let v_gap_du = table.constants.RadicalRuleThickness;
    let v_gap = table.du_to_pt(v_gap_du, fs);
    let rule_thick = table.du_to_pt(table.constants.RadicalRuleThickness, fs);
    let extra_asc = table.du_to_pt(table.constants.RadicalExtraAscender, fs);

    // S528 (coverage): the √ sign STRETCHES to the radicand height (Word selects
    // a taller MATH glyph variant for tall radicands — fractions, nested radicals).
    // Approximate by rendering the √ glyph at a font size that spans the radicand
    // height + gap + rule, instead of a fixed `fs` (which left tall radicands
    // sticking out the top, e.g. √√x was 36×28 vs Word 47×49).
    // Only stretch the √ for a TALL radicand (fraction, nested radical, etc.);
    // a normal single-line radicand keeps the base √ size. NB: a true stretchy √
    // is taller-NOT-wider (Word selects a narrow tall MATH glyph variant); the
    // renderer ignores text_scale and we have no variant glyphs, so the font-
    // scaled √ is proportionally wider than Word's for tall cases — but it COVERS
    // the radicand (vs a tiny base √ leaving the fraction sticking out before).
    let rad_height = rad_bbox.ascent + rad_bbox.descent;
    let sign_fs = if rad_height > fs * 1.2 {
        (rad_height + v_gap + rule_thick).max(fs)
    } else {
        fs
    };
    let sign_width = sign_fs * 0.55;

    // Radicand inner left edge (after √ sign).
    let radicand_x = x + sign_width;

    // Overbar y: above radicand top, gap above.
    let radicand_top_y = baseline_y - rad_bbox.ascent;
    let bar_y = radicand_top_y - v_gap - rule_thick / 2.0;
    let bar_width = rad_bbox.advance;

    // Render the √ sign at the (possibly stretched) size, baseline placed so the
    // glyph top ≈ bar_y and bottom ≈ the radicand bottom.
    let mut elems = Vec::new();
    let sign_baseline = bar_y + sign_fs * 0.78;
    elems.push(emit_text_at('\u{221A}'.to_string(), x, sign_baseline, sign_fs));

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

    // S525 (coverage): inter-column gap. MathLeading (~1.5%fs) is far too small
    // for a matrix — Word spaces columns by ~0.8em (the default mcSp). Measured:
    // 2-col single-digit matrix Word 42px wide vs Oxi 23 (gap ≈0). Use 0.8em.
    let col_gap = fs * 0.8;
    // Inter-row gap: a smaller fraction (rows are spaced by leading + a bit).
    let row_gap = fs * 0.35;

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
    operator_color: Option<&str>,
) -> (Vec<LayoutElement>, MathBBox) {
    use crate::ir::LimLoc;
    let table = MathTable::cambria_math();
    let fs = ctx.font_size;

    // S524/S525 (coverage, 2026-06-09): Word places n-ary limits ABOVE/BELOW
    // (undOvr) for DISPLAY equations by default for STACKING operators
    // (∑∏⋃⋂⋀⋁⨁⨀ …), but INTEGRALS (∫∬∭∮∯∰∱∲∳, U+222B..U+2233) keep their
    // limits to the RIGHT (subSup) even in display. The parser hard-defaults to
    // SubSup; flip to UndOvr in display EXCEPT for integrals (S525: the integral
    // repro showed Oxi stacked 32×66 vs Word's subSup 41×55). Without the flip
    // the ∑ repro drew subSup at half the bbox-reserved height (31px vs Word 68).
    let op_is_integral = ('\u{222B}'..='\u{2233}').contains(&op);
    let lim_loc = if ctx.style.is_display() && matches!(lim_loc, LimLoc::SubSup) && !op_is_integral {
        LimLoc::UndOvr
    } else {
        lim_loc
    };

    // Operator glyph: render larger if grow or display. S653 (coverage): a
    // DISPLAY integral is drawn extra-tall (Word ∫ ink ≈ fs×2.34); see the
    // layout_expr Nary counterpart.
    // S1252: inline (text-style) n-ary is drawn at the RUN size — see the
    // layout_expr counterpart for the Word measurement.
    let op_size = if op_is_integral && ctx.style.is_display() {
        fs * 2.34
    } else if ctx.style.is_display() {
        fs * 1.6
    } else if std::env::var("OXI_S1252_DISABLE").is_err() {
        fs
    } else {
        fs * 1.2
    };
    // ∫ tall-and-narrow: decouple advance width from the taller op_size (S653).
    let op_w = op_size * if op_is_integral { 0.36 } else { 0.6 };

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
            let mut operator = emit_text_at(op.to_string(), op_x, baseline_y, op_size);
            if let LayoutContent::Text { color, .. } = &mut operator.content {
                *color = operator_color.map(str::to_owned);
            }
            elems.push(operator);

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
            let mut operator = emit_text_at(op.to_string(), cur_x, baseline_y, op_size);
            if let LayoutContent::Text { color, .. } = &mut operator.content {
                *color = operator_color.map(str::to_owned);
            }
            elems.push(operator);
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
    // S525 (coverage): the over/under brace (⏞⏟) is a WIDE, SHORT stretchy
    // glyph — render at ~0.5×fs and reserve ~0.45×fs, not 0.8×fs (under-brace
    // "xyz" was Oxi 37px tall vs Word 22). Center it across the base width.
    let chr_size = fs * 0.5;
    let chr_x = x + (base_bbox.advance - chr_size * 0.6) / 2.0;

    let chr_baseline = match pos {
        BarPos::Top => baseline_y - base_bbox.ascent - gap - chr_size * 0.1,
        BarPos::Bot => baseline_y + base_bbox.descent + gap + chr_size * 0.6,
    };
    elems.push(emit_text_at(chr.to_string(), chr_x, chr_baseline, chr_size));

    let mut bbox = base_bbox;
    match pos {
        BarPos::Top => bbox.ascent += gap + chr_size * 0.9,
        BarPos::Bot => bbox.descent += gap + chr_size * 0.9,
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
    let fs = ctx.effective_font_size();
    let s_ctx = ctx.descend_script();
    let sub_bbox = layout_expr(sub, &s_ctx);
    let sup_bbox = layout_expr(sup, &s_ctx);

    let pre_w = sub_bbox.advance.max(sup_bbox.advance)+space_after_script(ctx);
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
    if let Some(shape)=accent_geometry(accent,base,ctx) {
        let fs=ctx.effective_font_size();let data=StretchTable::cambria_math();let scale=fs/data.upm as f32;
        let(mut elements,bbox)=emit_expr(base,x,baseline_y,ctx);
        for(i,p)in shape.plan.placements.iter().enumerate() {
            let mut e=emit_text_at(if i==0{accent.to_string()}else{String::new()},
                x+shape.x_shift+p.advance_offset as f32*scale,baseline_y+shape.baseline_shift,fs);
            e.width=p.glyph.advance_width as f32*scale;
            e.font_glyph=Some(FontGlyph{index:p.glyph.gid,bounds_em:p.glyph.bounds.map(|v|v as f32/data.upm as f32)});
            elements.push(e);
        }
        return(elements,MathBBox{ascent:bbox.ascent.max(-shape.ink_top),
            descent:bbox.descent.max(shape.ink_bottom),..bbox});
    }
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

    // Accent y: a SMALL glyph sitting just above the base top with a small gap.
    // S525 (coverage): acc_size 0.9→0.6×fs and reserve ~0.35×fs height (the hat
    // ink), not the full glyph em — Word x̂ is 14px tall vs Oxi's 28px before.
    let acc_size = fs * 0.6;
    let ascent = base_bbox.ascent;
    let gap = table.du_to_pt(table.constants.OverbarVerticalGap, fs);
    // The combining-accent glyph draws its ink ABOVE its baseline; place the
    // baseline so the accent sits just above base-top + gap.
    let accent_baseline = baseline_y - ascent - gap + acc_size * 0.55;
    // Accent char rendered at (x + attach_x), shifted left by half its width.
    let accent_w = acc_size * 0.4;
    let accent_x = x + attach_x - accent_w / 2.0;
    elems.push(emit_text_at(accent.to_string(), accent_x, accent_baseline, acc_size));

    let bbox = MathBBox {
        advance: base_bbox.advance,
        ascent: base_bbox.ascent + gap + fs * 0.35,
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

/// Emit a pair of delimiters selected for the content's ink height.
fn emit_delimiter(beg: char, end: char, content: &MathExpr, x: f32,
                  baseline_y: f32, ctx: &MathLayoutContext) -> (Vec<LayoutElement>, MathBBox) {
    let cb = layout_expr(content, ctx);
    let lw = delimiter_width(beg, content, ctx);
    let rw = delimiter_width(end, content, ctx);
    let mut elements = emit_delimiter_glyph(beg, content, x, baseline_y, ctx);
    elements.extend(emit_expr(content, x + lw, baseline_y, ctx).0);
    elements.extend(emit_delimiter_glyph(end, content, x + lw + cb.advance, baseline_y, ctx));
    let (la, ld) = delimiter_ink(beg, content, ctx, false);
    let (ra, rd) = delimiter_ink(end, content, ctx, false);
    (elements, MathBBox { advance: lw + cb.advance + rw,
        ascent: cb.ascent.max(la).max(ra), descent: cb.descent.max(ld).max(rd), italic_correction: 0.0 })
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
    let (mut base_elems, base_bbox) = emit_expr(base, x, baseline_y, ctx);

    let s_ctx = ctx.descend_script();
    let (sup_shift, sub_shift) = combined_script_shifts(sub, sup, ctx, false);

    let script_x = x + base_bbox.advance + base_bbox.italic_correction;
    let sup_kern=script_kern(base,sup,ctx,sup_shift,true);
    let sub_kern=script_kern(base,sub,ctx,sub_shift,false);
    let (sup_e, sup_b) = emit_expr(sup, script_x+sup_kern, baseline_y - sup_shift, &s_ctx);
    let (sub_e, sub_b) = emit_expr(sub, script_x+sub_kern, baseline_y + sub_shift, &s_ctx);
    base_elems.extend(sup_e);
    base_elems.extend(sub_e);

    let bbox = MathBBox {
        advance: base_bbox.advance + base_bbox.italic_correction
            + (sup_b.advance+sup_kern).max(sub_b.advance+sub_kern)+space_after_script(ctx),
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
    // S527 inter-atom math-class spacing applied to the top-level content too.
    let row = math_row_atoms(exprs);
    let exprs = row.as_ref();
    let gaps = atom_gaps(exprs, font_size);
    // Pre-compute baseline: first pass finds needed ascent.
    let mut total_bbox = MathBBox::default();
    for (i, e) in exprs.iter().enumerate() {
        total_bbox.advance += gaps[i];
        let b = layout_expr(e, &ctx);
        total_bbox = total_bbox.hstack(&b);
    }
    // Baseline sits at cursor_y + ascent (top-relative).
    let baseline_y = cursor_y + total_bbox.ascent.max(font_size * 0.8);

    let mut elems = Vec::new();
    let mut cur_x = x;
    for (i, e) in exprs.iter().enumerate() {
        cur_x += gaps[i];
        let (ee, b) = emit_expr(e, cur_x, baseline_y, &ctx);
        elems.extend(ee);
        cur_x += b.advance;
    }
    if matches!(block,MathBlock::Display{..}) && !elems.is_empty()
        && std::env::var("OXI_S652_DISABLE").is_err() {
        let (top,bottom)=reserved_line_extents(&elems);
        if top.is_finite() && bottom >= top {
            let table=MathTable::cambria_math();
            let leading=table.du_to_pt(table.constants.MathLeading,font_size);
            let shift=cursor_y+leading-top;
            for e in &mut elems {
                e.y += shift;
                if let LayoutContent::TableBorder{y1,y2,..}=&mut e.content {
                    *y1 += shift;*y2 += shift;
                }
            }
            total_bbox.ascent=baseline_y+shift-cursor_y;
            total_bbox.descent=(bottom-baseline_y).max(0.0);
        }
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
            host: None,
            reduce_fraction_size: false,
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
        // Word 10.5pt control: nominal7.5pt, PDF7.56 after device rounding.
        assert!((ctx_s.effective_font_size() - 7.5).abs() < 0.01);
        let ctx_ss = ctx_s.descend_script();
        // Same control's second script level: Word/PDF6.0pt.
        assert!((ctx_ss.effective_font_size() - 6.0).abs() < 0.01);
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
            host: None,
            reduce_fraction_size: false,
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

/// S1252 (2026-08-29): the INLINE box of a maths expression — `(advance,
/// ascent, descent)` in points, both vertical parts measured from the EMITTED
/// glyph geometry rather than from `MathBBox`.
///
/// `MathBBox` is a loose over-estimate — the S652 note at the display-maths arm
/// records why (`emit_nary`'s descent double-counts, leaf glyph boxes are a
/// slack 0.8em/0.4em). For a 10.5pt `2π/3` it returns 62pt against a real ink
/// extent of 16.4pt, which as an inline object height would swallow the page.
/// The convention here is `s1244_math_advance`'s: a text element's baseline is
/// `y + 2/3·h`, its ink runs `[baseline − 0.60·fs, baseline + 0.05·fs]`, and a
/// non-text primitive (fraction bar, radical rule) is already tight.
///
/// WORD TRUTH (`_pb_inlmath`, PDF text origins — exact baselines):
/// against the plain 11.64/11.66 pitch the maths grows the line
/// `f(x)=4cos(3x)` by +0.00, `x²` by +0.00, `√x` by +1.68 (ascent side) and
/// `2π/3` by +5.04 (+2.28 ascent, +2.76 descent) — i.e. the two sides compose
/// independently, which is what S1095 does with this ascent/descent pair.
/// S1611: ink (ascent, descent) as Word composes it for counting grid cells.
/// Differs from `ink_extent` in two measured ways (Word PDF of blind-G JA
/// educational__20d9968b p2): every shift scales with the CURRENT script size,
/// and radicands, denominators and subscripts are CRAMPED, taking
/// SuperscriptShiftUpCramped (615du): R=sqrt(a^2+b^2) raises its `2` 3.1pt at
/// 10.5pt (615 -> 3.15; 750 would be 3.85) and the same radicand inside a
/// denominator 2.2pt at 7.56pt (615 -> 2.27).
fn ink_extent_word(expr: &MathExpr, ctx: &MathLayoutContext, cramped: bool) -> (f32, f32) {
    let table = MathTable::cambria_math();
    let eff = ctx.effective_font_size();
    match expr {
        MathExpr::Run { text, style } => run_ink(text, style, ctx),
        MathExpr::Text(_) => ink_extent(expr, ctx),
        MathExpr::Seq(children) => children.iter().map(|c| ink_extent_word(c, ctx, cramped))
            .fold((0.0f32, 0.0f32), |(a, d), (ca, cd)| (a.max(ca), d.max(cd))),
        MathExpr::Fraction { num, den, .. } => {
            let display = ctx.style.is_display();
            let sub_ctx = ctx.descend_fraction();
            let (up_du, down_du, ng, dg) = if display {
                (table.constants.FractionNumeratorDisplayStyleShiftUp, table.constants.FractionDenominatorDisplayStyleShiftDown,
                 table.constants.FractionNumDisplayStyleGapMin, table.constants.FractionDenomDisplayStyleGapMin)
            } else {
                (table.constants.FractionNumeratorShiftUp, table.constants.FractionDenominatorShiftDown,
                 table.constants.FractionNumeratorGapMin, table.constants.FractionDenominatorGapMin)
            };
            let axis = table.du_to_pt(table.constants.AxisHeight, eff);
            let half = table.du_to_pt(table.constants.FractionRuleThickness, eff) / 2.0;
            let (na, nd) = ink_extent_word(num, &sub_ctx, cramped);
            let (da, dd) = ink_extent_word(den, &sub_ctx, true);
            let up = table.du_to_pt(up_du, eff).max(axis + half + table.du_to_pt(ng, eff) + nd);
            let down = table.du_to_pt(down_du, eff).max(table.du_to_pt(dg, eff) + da - (axis - half));
            (up + na, down + dd)
        }
        MathExpr::Radical { degree, radicand } => {
            if let Some(shape)=radical_geometry(radicand,ctx) {
                let mut top=shape.ink_top;let mut bottom=shape.ink_bottom;
                let (_,placement)=radical_degree(degree.as_deref(),ctx,&shape);
                if let (Some(expr),Some((_,baseline,dctx)))=(degree.as_deref(),placement) {
                    let(a,d)=ink_extent_word(expr,&dctx,false);top=top.min(baseline-a);bottom=bottom.max(baseline+d);
                }
                return ((-top).max(0.0),bottom.max(0.0));
            }
            let (ra, rd) = ink_extent_word(radicand, ctx, true);
            let gap_du = table.constants.RadicalRuleThickness;
            (ra + table.du_to_pt(gap_du, eff) + table.du_to_pt(table.constants.RadicalRuleThickness, eff), rd)
        }
        MathExpr::Superscript { base, sup } => {
            let (ba, bd) = ink_extent_word(base, ctx, cramped);
            let (sa, _) = ink_extent_word(sup, &ctx.descend_script(), cramped);
            let up_du = if cramped { table.constants.SuperscriptShiftUpCramped } else { table.constants.SuperscriptShiftUp };
            (ba.max(sa + table.du_to_pt(up_du, eff)), bd)
        }
        MathExpr::Subscript { base, sub } => {
            let (ba, bd) = ink_extent_word(base, ctx, cramped);
            let (_, sd) = ink_extent_word(sub, &ctx.descend_script(), true);
            (ba, bd.max(sd + table.du_to_pt(table.constants.SubscriptShiftDown, eff)))
        }
        MathExpr::SubSuperscript { base, sub, sup } => {
            let (ba, bd) = ink_extent_word(base, ctx, cramped);
            let (sa, _) = ink_extent_word(sup, &ctx.descend_script(), cramped);
            let (_, sd) = ink_extent_word(sub, &ctx.descend_script(), true);
            let (up, down) = combined_script_shifts(sub, sup, ctx, cramped);
            (ba.max(sa + up), bd.max(sd + down))
        }
        // Compose the accent glyph with its base ink on the common baseline.
        // The base's subscript depth stays its ink depth; a loose layout box
        // must not leak back into the fraction's numerator gap through Accent.
        MathExpr::Accent { accent, base } => {
            if let Some(shape)=accent_geometry(*accent,base,ctx) {
                let(a,d)=ink_extent_word(base,ctx,cramped);
                return(a.max(-shape.ink_top),d.max(shape.ink_bottom));
            }
            let (ba, bd) = ink_extent_word(base, ctx, cramped);
            let paint_fs = ctx.font_size;
            let accent_fs = paint_fs * 0.6;
            let gap = table.du_to_pt(table.constants.OverbarVerticalGap, paint_fs);
            let offset = -layout_expr(base, ctx).ascent - gap + accent_fs * 0.55;
            let (aa, ad) = glyph_ink_du(*accent).unwrap_or((0.7, 0.2));
            (ba.max(aa * accent_fs - offset), bd.max(ad * accent_fs + offset))
        }
        // A function name and its argument share the baseline.
        MathExpr::Function { name, arg } => {
            let (na, nd) = ink_extent_word(name, ctx, cramped);
            let (aa, ad) = ink_extent_word(arg, ctx, cramped);
            (na.max(aa), nd.max(ad))
        }
        // Delimiters grow to cover their content; a short content keeps the
        // glyph's own ink.
        MathExpr::Delimiter { beg, end, content, .. } => {
            let (ca, cd) = ink_extent_word(content, ctx, cramped);
            let mut a = ca;
            let mut d = cd;
            for c in [*beg, *end] {
                let (ga, gd) = delimiter_ink(c, content, ctx, cramped);
                a = a.max(ga);
                d = d.max(gd);
            }
            (a, d)
        }
        // n-ary (∑ ∫ …): the operator glyph's own ink, limits either as
        // scripts (subSup) or stacked with the OpenType limit constants (undOvr),
        // and the operand on the shared baseline. reports__5823d5a8 p4: a
        // fraction of two ∑_{i=1}^{n} terms takes 2 cells in Word; the layout-box
        // fallback read it as 36/30pt and gave 4.
        MathExpr::Nary { op, sub, sup, operand, lim_loc, grow, .. } => {
            if let Some(shape)=nary_geometry(*op,sub.as_deref(),sup.as_deref(),operand,*lim_loc,*grow,ctx,cramped) {
                return (-shape.ink_top,shape.ink_bottom);
            }
            let (ga, gd) = glyph_ink_du(*op).unwrap_or((0.8, 0.3));
            let (oa, od) = (ga * eff, gd * eff);
            let sctx = ctx.descend_script();
            let (mut a, mut d) = (oa, od);
            match lim_loc {
                crate::ir::math::LimLoc::UndOvr => {
                    if let Some(s) = sup {
                        let (la, ld) = ink_extent_word(s, &sctx, cramped);
                        let rise = table.du_to_pt(table.constants.UpperLimitBaselineRiseMin, eff)
                            .max(oa + table.du_to_pt(table.constants.UpperLimitGapMin, eff) + ld);
                        a = a.max(rise + la);
                    }
                    if let Some(s) = sub {
                        let (la, ld) = ink_extent_word(s, &sctx, true);
                        let drop = table.du_to_pt(table.constants.LowerLimitBaselineDropMin, eff)
                            .max(od + table.du_to_pt(table.constants.LowerLimitGapMin, eff) + la);
                        d = d.max(drop + ld);
                    }
                }
                crate::ir::math::LimLoc::SubSup => {
                    if let Some(s) = sup {
                        let (sa, _) = ink_extent_word(s, &sctx, cramped);
                        let up_du = if cramped { table.constants.SuperscriptShiftUpCramped } else { table.constants.SuperscriptShiftUp };
                        a = a.max(sa + table.du_to_pt(up_du, eff));
                    }
                    if let Some(s) = sub {
                        let (_, sd) = ink_extent_word(s, &sctx, true);
                        d = d.max(sd + table.du_to_pt(table.constants.SubscriptShiftDown, eff));
                    }
                }
            }
            let (pa, pd) = ink_extent_word(operand, ctx, cramped);
            (a.max(pa), d.max(pd))
        }
        // limLow / limUpp: OpenType MATH lower/upper limit placement.
        MathExpr::Limit { base, lim, pos } => {
            let (ba, bd) = ink_extent_word(base, ctx, cramped);
            let lctx = ctx.descend_script();
            match pos {
                crate::ir::math::LimitPos::Lower => {
                    let (la, ld) = ink_extent_word(lim, &lctx, true);
                    let drop = table.du_to_pt(table.constants.LowerLimitBaselineDropMin, eff)
                        .max(bd + table.du_to_pt(table.constants.LowerLimitGapMin, eff) + la);
                    (ba, bd.max(drop + ld))
                }
                crate::ir::math::LimitPos::Upper => {
                    let (la, ld) = ink_extent_word(lim, &lctx, cramped);
                    let rise = table.du_to_pt(table.constants.UpperLimitBaselineRiseMin, eff)
                        .max(ba + table.du_to_pt(table.constants.UpperLimitGapMin, eff) + ld);
                    (ba.max(rise + la), bd)
                }
            }
        }
        _ => ink_extent(expr, ctx),
    }
}

/// S1611: the per-glyph ink (top, bottom) of a whole maths block about its
/// baseline, composed exactly as `ink_extent` composes a fraction (shifts from
/// `fraction_shifts`, numerator/denominator ink at script size) and a radical
/// (the radicand's depth, not the radical sign's). Used to count grid cells.
pub fn inline_math_ink_extent(block: &MathBlock, font_size: f32) -> (f32, f32) {
    let ctx = MathLayoutContext {
        font_size,
        style: MathStyle::from_block(block),
    };
    let exprs: &[MathExpr] = match block {
        MathBlock::Inline(xs) => xs,
        MathBlock::Display { content, .. } => content,
    };
    exprs.iter().map(|e| ink_extent_word(e, &ctx, false))
        .fold((0.0f32, 0.0f32), |(a, d), (ea, ed)| (a.max(ea), d.max(ed)))
}

pub fn inline_math_ink(block: &MathBlock, font_size: f32) -> (f32, f32, f32) {
    let (elems, bbox) = emit_math_block(block, 0.0, 0.0, font_size);
    let baseline = bbox.ascent.max(font_size * 0.8);
    let (asc_r, desc_r) = (0.60_f32, 0.05_f32);
    let mut ink_top = f32::INFINITY;
    let mut ink_bot = f32::NEG_INFINITY;
    for e in &elems {
        let (lo,hi)=painted_element_ink(e);
        ink_top = ink_top.min(lo);
        ink_bot = ink_bot.max(hi);
    }
    if ink_bot > ink_top {
        (bbox.advance, (baseline - ink_top).max(0.0), (ink_bot - baseline).max(0.0))
    } else {
        (bbox.advance, bbox.ascent, bbox.descent)
    }
}

#[cfg(test)]
mod combined_script_gap_tests {
    use super::*;
    use crate::ir::{MathRunStyle, MathStyleVariant};

    fn bold(text: &str) -> MathExpr {
        MathExpr::Run {
            text: text.to_owned(),
            style: MathRunStyle { math_style: Some(MathStyleVariant::Bold), ..Default::default() },
        }
    }

    #[test]
    fn combined_scripts_match_word_relative_baselines() {
        let base = bold("S");
        let sub = bold("1");
        let sup = bold("2");
        let expr = MathExpr::SubSuperscript {
            base: Box::new(base), sub: Box::new(sub), sup: Box::new(sup),
        };
        let ctx = MathLayoutContext { font_size: 12.0, style: MathStyle::Display };
        let (elements, _) = emit_expr(&expr, 0.0, 100.0, &ctx);
        let baseline = |e: &LayoutElement| e.y + e.baseline_offset.unwrap();
        // Saved Word PDF controls: base117.50 / sup113.06 / sub120.74.
        assert!((100.0 - baseline(&elements[1]) - 4.44).abs() <= 0.35);
        assert!((baseline(&elements[2]) - 100.0 - 3.24).abs() <= 0.35);
    }

    #[test]
    fn combined_scripts_keep_minimum_ink_gap_across_styles() {
        for size in [8.0, 10.5, 12.0, 18.0, 20.0] {
            for style in [MathStyle::Display, MathStyle::CompactFullSize, MathStyle::Text,
                          MathStyle::Script, MathStyle::ScriptScript] {
                for cramped in [false, true] {
                    for (sub, sup) in [(bold("1"), bold("2")),
                                      (MathExpr::Text("p".into()), MathExpr::Text("g".into())),
                                      (MathExpr::Text("".into()), MathExpr::Text("".into()))] {
                        let ctx = MathLayoutContext { font_size: size, style };
                        let (up, down) = combined_script_shifts(&sub, &sup, &ctx, cramped);
                        let (a, _) = ink_extent_word(&sub, &ctx.descend_script(), true);
                        let (_, d) = ink_extent_word(&sup, &ctx.descend_script(), cramped);
                        let table = MathTable::cambria_math();
                        let required = table.du_to_pt(table.constants.SubSuperscriptGapMin, ctx.effective_font_size());
                        assert!(up + down - a - d + 0.0001 >= required);
                        assert!(up.is_finite() && down.is_finite() && up >= 0.0 && down >= 0.0);
                    }
                }
            }
        }
    }
}

#[cfg(test)]
mod delimiter_geometry_tests {
    use super::*;
    use crate::ir::{FracBarType, MathRunStyle, MathStyleVariant};

    #[test]
    fn structural_parentheses_match_word_axis_and_advance() {
        let run = |text: &str| MathExpr::Run { text: text.to_owned(),
            style: MathRunStyle { math_style: Some(MathStyleVariant::BoldItalic), ..Default::default() } };
        let content = MathExpr::Seq(vec![MathExpr::Subscript {
            base: Box::new(run("n")), sub: Box::new(run("1")),
        }, run("-"), run("1")]);
        let expr = MathExpr::Delimiter { beg: '(', end: ')', sep: None, content: Box::new(content) };
        let ctx = MathLayoutContext { font_size: 12.0, style: MathStyle::Display };
        let (elements, _) = emit_expr(&expr, 0.0, 100.0, &ctx);
        let left = &elements[0];
        assert_eq!(left.font_glyph.unwrap().index, 4666);
        assert!((left.y + left.baseline_offset.unwrap() - 99.52).abs() <= 0.35);
        assert!((elements[1].x - 5.04).abs() <= 0.35);
    }

    #[test]
    fn literal_parentheses_retain_the_text_baseline() {
        let expr = MathExpr::Run { text: "(n1)".into(),
            style: MathRunStyle { literal: true, ..Default::default() } };
        let ctx = MathLayoutContext { font_size: 12.0, style: MathStyle::Display };
        let (elements, _) = emit_expr(&expr, 0.0, 100.0, &ctx);
        assert!(!elements.is_empty());
        assert!(elements.iter().all(|e| (e.y + e.baseline_offset.unwrap() - 100.0).abs() < 0.001));
    }

    #[test]
    fn delimiter_plans_cover_content_and_agree_with_emitted_widths() {
        let frac = MathExpr::Fraction { num: Box::new(MathExpr::Text("a".into())),
            den: Box::new(MathExpr::Text("b".into())), bar_type: FracBarType::Bar };
        let nested = MathExpr::Fraction { num: Box::new(frac.clone()), den: Box::new(frac.clone()),
            bar_type: FracBarType::Bar };
        for size in [8.0, 12.0, 18.0] {
            for style in [MathStyle::Display, MathStyle::Text, MathStyle::Script] {
                for (beg, end) in [('(', ')'), ('[', ']'), ('{', '}')] {
                    for content in [&frac, &nested] {
                        let ctx = MathLayoutContext { font_size: size, style };
                        let (a, d) = ink_extent_word(content, &ctx, false);
                        let shape = delimiter_geometry(beg, content, &ctx, false).unwrap();
                        assert!(shape.ink_bottom - shape.ink_top + 0.01 >= a + d);
                        let axis = MathTable::cambria_math().du_to_pt(
                            MathTable::cambria_math().constants.AxisHeight, ctx.effective_font_size());
                        assert!(((shape.ink_top + shape.ink_bottom) * 0.5 + axis).abs() < 0.001);
                        let expr = MathExpr::Delimiter { beg, end, sep: None, content: Box::new(content.clone()) };
                        let layout = layout_expr(&expr, &ctx);
                        let (_, emitted) = emit_expr(&expr, 0.0, 100.0, &ctx);
                        assert!((layout.advance - emitted.advance).abs() < 0.001);
                    }
                }
            }
        }
    }
}

#[cfg(test)]
mod radical_rule_gap_tests {
    use super::*;
    use crate::ir::FracBarType;

    #[test]
    fn fraction_in_radical_matches_word_relative_baselines() {
        let fraction = MathExpr::Fraction { num: Box::new(MathExpr::Text("1".into())),
            den: Box::new(MathExpr::Subscript { base: Box::new(MathExpr::Text("n".into())),
                sub: Box::new(MathExpr::Text("1".into())) }), bar_type: FracBarType::Bar };
        let expr = MathExpr::Radical { degree: None, radicand: Box::new(fraction) };
        for (size, expected) in [(12.0, [-8.88, 8.28, 10.68]), (18.0, [-13.32, 12.48, 16.08])] {
            let ctx = MathLayoutContext { font_size: size, style: MathStyle::Display };
            let (elements, _) = emit_expr(&expr, 0.0, 100.0, &ctx);
            let baseline = |e: &LayoutElement| e.y + e.baseline_offset.unwrap();
            let root = baseline(&elements[0]);
            for (element, reference) in elements[1..4].iter().zip(expected) {
                assert!((baseline(element) - root - reference).abs() <= 0.35);
            }
        }
    }
}

#[cfg(test)]
mod row_token_regression_tests {
    use super::*;
    use crate::ir::{MathRunStyle, MathStyleVariant};

    fn run(text: &str, style: &MathRunStyle) -> MathExpr {
        MathExpr::Run { text: text.into(), style: style.clone() }
    }

    fn emitted_origins(expr: &MathExpr, ctx: &MathLayoutContext) -> (Vec<(String, f32, f32)>, f32) {
        let (elements, bbox) = emit_expr(expr, 20.0, 100.0, ctx);
        let origins = elements.iter().filter_map(|e| match &e.content {
            LayoutContent::Text { text, .. } => Some((text.clone(), e.x, e.y + e.baseline_offset.unwrap())),
            _ => None,
        }).collect();
        (origins, bbox.advance)
    }

    #[test]
    fn equation_spacing_is_independent_of_operator_run_boundaries() {
        for fs in [8.0, 12.0, 18.0] {
            for math_style in [None, Some(MathStyleVariant::BoldItalic), Some(MathStyleVariant::Plain)] {
                let style = MathRunStyle { math_style, ..Default::default() };
                for context_style in [MathStyle::Display, MathStyle::Text, MathStyle::Script] {
                    let ctx = MathLayoutContext { font_size: fs, style: context_style };
                    let grouped = run("a+(b)", &style);
                    let split = MathExpr::Seq(["a", "+", "(", "b", ")"].iter()
                        .map(|s| run(s, &style)).collect());
                    let a = emitted_origins(&grouped, &ctx);
                    let b = emitted_origins(&split, &ctx);
                    assert_eq!(a.0, b.0);
                    assert!((a.1 - b.1).abs() < 0.001);
                    assert!((layout_expr(&grouped, &ctx).advance - a.1).abs() < 0.001);
                    for xs in [vec![grouped.clone()], vec![split.clone()]] {
                        let block = MathBlock::Inline(xs);
                        let (_, emitted) = emit_math_block(&block, 20.0, 100.0, fs);
                        assert!((layout_math_block(&block, fs).advance - emitted.advance).abs() < 0.001);
                    }
                }
            }
        }
    }

    #[test]
    fn normal_text_and_identifiers_keep_their_shaping_runs() {
        let literal = MathRunStyle { literal: true, ..Default::default() };
        let normal = MathRunStyle::default();
        let ctx = MathLayoutContext { font_size: 12.0, style: MathStyle::Text };
        for expr in [run("a+(b)", &literal), run("alpha", &normal), MathExpr::Text("alpha".into())] {
            assert!(math_run_atoms(&expr).is_none());
            let (e, _) = emit_expr(&expr, 20.0, 100.0, &ctx);
            assert_eq!(e.len(), 1);
        }
    }

    #[test]
    fn grouped_binary_operator_restores_both_word_measured_gaps() {
        // Word's saved display controls place the grouped '+' 2.67pt after
        // the preceding operand and '(' a further 2.67pt after the operator
        // advance at 12pt. The previous row omitted both 4/18-em gaps.
        let style = MathRunStyle::default();
        let ctx = MathLayoutContext { font_size: 12.0, style: MathStyle::Text };
        let row = MathExpr::Seq(vec![run("x", &style), run("+(", &style), run("y", &style)]);
        let (elements, _) = emit_expr(&row, 20.0, 100.0, &ctx);
        let plus = &elements[1];
        let open = &elements[2];
        let x_advance = layout_expr(&run("x", &style), &ctx).advance;
        let plus_advance = layout_expr(&run("+", &style), &ctx).advance;
        assert!((plus.x - (20.0 + x_advance) - 2.67).abs() < 0.01);
        assert!((open.x - (plus.x + plus_advance) - 2.67).abs() < 0.01);
    }
}

#[cfg(test)]
mod nary_glyph_geometry_regression_tests {
    use super::*;
    fn leaf(text:&str)->MathExpr {
        MathExpr::Run{text:text.into(),style:crate::ir::MathRunStyle{
            math_style:Some(crate::ir::MathStyleVariant::Plain),..Default::default()}}
    }
    #[test]
    fn lower_only_limit_keeps_the_measured_right_side_anchor() {
        for (op,expected_x) in [('\u{222b}',99.744),('\u{2211}',103.10)] {
            let expr=MathExpr::Nary{operator_color: None, op,sub:Some(Box::new(leaf("α"))),sup:None,
                operand:Box::new(MathExpr::Text("f(x)".into())),lim_loc:crate::ir::LimLoc::SubSup,grow:false};
            let ctx=MathLayoutContext{font_size:10.5,style:MathStyle::Text};
            let (elements,emitted)=emit_expr(&expr,95.664,140.0,&ctx);
            let lower=elements.iter().find(|e|matches!(&e.content,LayoutContent::Text{text,..}if text=="α")).unwrap();
            assert!((lower.x-expected_x).abs()<0.35,"operator {op}: {} vs {expected_x}",lower.x);
            assert!((layout_expr(&expr,&ctx).advance-emitted.advance).abs()<0.001);
        }
    }
    #[test]
    fn grown_operator_emits_the_font_variant_at_its_nominal_size() {
        let ctx=MathLayoutContext{font_size:10.5,style:MathStyle::Text};
        let make=|grow|MathExpr::Nary{operator_color: None, op:'\u{2211}',sub:Some(Box::new(leaf("α"))),sup:Some(Box::new(leaf("β"))),
            operand:Box::new(MathExpr::Text("f(x)g(x)dx".into())),lim_loc:crate::ir::LimLoc::SubSup,grow};
        for (grow,gid) in [(false,963),(true,3532)] {
            let expr=make(grow);let (elements,emitted)=emit_expr(&expr,0.0,100.0,&ctx);
            let operator=&elements[0];assert_eq!(operator.font_glyph.unwrap().index,gid);
            assert!(matches!(&operator.content,LayoutContent::Text{font_size,..}if (*font_size-10.5).abs()<0.001));
            assert!((layout_expr(&expr,&ctx).advance-emitted.advance).abs()<0.001);
        }
    }
}

#[cfg(test)]
mod resolved_math_font_regression_tests {
    use super::*;
    #[test]
    fn fallback_greek_has_word_face_ink_and_an_em_advance() {
        let style=crate::ir::MathRunStyle {math_style:Some(crate::ir::MathStyleVariant::Plain),
            run_style:Some(crate::ir::RunStyle {font_family:Some("MS Mincho".into()),font_size:Some(10.5),..Default::default()}),..Default::default()};
        let ctx=MathLayoutContext {font_size:10.5,style:MathStyle::Script};
        let glyphs=resolved_run_glyphs("α",&style,&ctx).expect("portable real face geometry");
        let bbox=resolved_run_bbox(&glyphs,&style,&ctx);
        assert!((bbox.advance-7.5).abs()<0.001);
        assert_eq!(glyphs[0].metrics.bounds_em,[0.2421875,-0.0078125,0.76953125,0.42578125]);
        let elements=emit_resolved_run(&glyphs,&style,99.744,132.74,&ctx);
        assert!(matches!(&elements[0].content,LayoutContent::Text {font_family:Some(f),font_size,..}if f=="MS Mincho" && (*font_size-7.5).abs()<0.001));
        assert!((elements[0].width-bbox.advance).abs()<0.001);
        assert!((elements[0].y+elements[0].baseline_offset.unwrap()-132.74).abs()<0.001);
    }
}


#[cfg(test)]
mod fallback_typographic_geometry_regression_tests {
    use super::*;
    fn leaf(text:&str)->MathExpr {
        MathExpr::Run {text:text.into(),style:crate::ir::MathRunStyle {
            math_style:Some(crate::ir::MathStyleVariant::Plain),run_style:Some(crate::ir::RunStyle {
                font_family:Some("MS Mincho".into()),font_size:Some(10.5),..Default::default()}),..Default::default()}}
    }
    #[test]
    fn font_signature_reserves_both_sides_without_a_family_exception() {
        let original=crate::font::catalog_glyph_face_metrics("MS Mincho",false,false).unwrap();
        let mut renamed=(*original).clone();renamed.family="arbitrary face name".into();
        assert_eq!(original.design_font_box_pt(7.5,true),renamed.design_font_box_pt(7.5,true));
        let (a,d)=renamed.design_font_box_pt(7.5,true);
        assert!((a-258.0/256.0*7.5).abs()<0.001);
        assert!((d-74.0/256.0*7.5).abs()<0.001);
    }
    #[test]
    fn script_capacity_preserves_nominal_box_and_the_actual_paint_baseline() {
        let expr=leaf("β");let ctx=MathLayoutContext{font_size:10.5,style:MathStyle::Script};
        let (elements,_)=emit_expr(&expr,100.0,135.26,&ctx);let e=&elements[0];
        assert!((e.y+e.baseline_offset.unwrap()-135.26).abs()<0.001);
        assert!((e.height-10.5).abs()<0.001);
        assert!(matches!(&e.content,LayoutContent::Text{font_size,..}if (*font_size-7.5).abs()<0.001));
        let (top,bottom)=painted_element_ink(e);
        assert!((bottom-top-0.828125*7.5).abs()<0.001);
    }
    #[test]
    fn upper_limits_use_the_selected_operator_shape() {
        let ctx=MathLayoutContext{font_size:10.5,style:MathStyle::Text};
        for (op,grow,expected) in [('∫',false,-6.2399902),('∑',false,-3.8400269)] {
            let expr=MathExpr::Nary {operator_color: None, op,sub:None,sup:Some(Box::new(leaf("β"))),
                operand:Box::new(MathExpr::Text("f(x)g(x)dx".into())),lim_loc:crate::ir::LimLoc::SubSup,grow};
            let (elements,_)=emit_expr(&expr,0.0,0.0,&ctx);
            let upper=elements.iter().find(|e|matches!(&e.content,LayoutContent::Text{text,..}if text=="β")).unwrap();
            assert!((upper.y+upper.baseline_offset.unwrap()-expected).abs()<0.35);
        }
    }
    #[test]
    fn integral_upper_limit_counts_the_word_grid_capacity() {
        let font=crate::font::catalog_glyph_face_metrics("MS Mincho",false,false).unwrap();
        for (size,pitch,expected) in [(10.5,18.0,54.0),(14.0,12.0,60.0)] {
            let mut upper=leaf("β");
            if let MathExpr::Run{style,..}=&mut upper {style.run_style.as_mut().unwrap().font_size=Some(size);}
            let block=MathBlock::Inline(vec![MathExpr::Nary {operator_color: None, op:'∫',sub:None,sup:Some(Box::new(upper)),
                operand:Box::new(MathExpr::Text("f(x)g(x)dx".into())),lim_loc:crate::ir::LimLoc::SubSup,grow:false}]);
            let (a,d)=inline_math_typographic_extent(&block,size).unwrap();let (ha,hd)=font.design_font_box_pt(size,true);
            let marker_span=((ha+hd)/pitch).ceil()*pitch+((a.max(ha)+d.max(hd))/pitch).ceil()*pitch;
            assert!((marker_span-expected).abs()<0.001,"size {size}, pitch {pitch}: {marker_span}");
        }
    }
}

#[cfg(test)]
#[test]
fn full_size_foreign_math_leaf_retains_its_painted_box() {
    let style=crate::ir::MathRunStyle {math_style:Some(crate::ir::MathStyleVariant::Plain),
        run_style:Some(crate::ir::RunStyle {font_family:Some("Times New Roman".into()),font_size:Some(12.0),
            ..Default::default()}),..Default::default()};
    let expr=MathExpr::Run {text:"-".into(),style};
    let ctx=MathLayoutContext {font_size:12.0,style:MathStyle::CompactFullSize};
    let (elements,_)=emit_expr(&expr,10.0,100.0,&ctx);let element=&elements[0];
    let (top,bottom)=painted_element_ink(element);
    assert!((element.y-top).abs()<0.001 && (element.y+element.height-bottom).abs()<0.001);
    assert!((element.y+element.baseline_offset.unwrap()-100.0).abs()<0.001);
    assert!(!is_fallback_font_element(element));
}


#[cfg(test)]
mod shared_math_line_baseline_regression_tests {
    use super::*;

    fn limit(text: &str, size: f32) -> MathExpr {
        MathExpr::Run {text: text.into(), style: crate::ir::MathRunStyle {
            math_style: Some(crate::ir::MathStyleVariant::Plain),
            run_style: Some(crate::ir::RunStyle {font_family: Some("MS Mincho".into()),
                font_size: Some(size), ..Default::default()}), ..Default::default()}}
    }

    #[test]
    fn composed_baseline_matches_saved_word_operand_origins() {
        // Saved Word PDF operand baselines relative to the measured grid-line
        // origin. Both limits, two sizes, grown/base shapes and two pitches.
        for (op, sub, sup, grow, size, height, expected) in [
            ('\u{222b}', true, false, false, 10.5, 24.0, 14.0499878),
            ('\u{222b}', true, false, false, 14.0, 24.0, 14.7699585),
            ('\u{2211}', false, true, false, 10.5, 18.0, 14.0499878),
            ('\u{2211}', true, true, true, 10.5, 24.0, 16.4500122),
        ] {
            let block = MathBlock::Inline(vec![MathExpr::Nary {operator_color: None, op,
                sub: sub.then(|| Box::new(limit("\u{03b1}", size))),
                sup: sup.then(|| Box::new(limit("\u{03b2}", size))),
                operand: Box::new(MathExpr::Text("f(x)g(x)dx".into())),
                lim_loc: crate::ir::LimLoc::SubSup, grow}]);
            let capacity_before = inline_math_typographic_extent(&block, size).unwrap();
            let (a, d) = inline_math_baseline_extent(&block, size).unwrap();
            let host = crate::font::catalog_glyph_face_metrics("MS Mincho", false, false).unwrap();
            let (ha, hd) = host.design_font_box_pt(size, true);
            let actual = (height + a.max(ha) - d.max(hd)) * 0.5;
            assert!((actual - expected).abs() <= 0.35,
                "op {op}, size {size}, sub {sub}, sup {sup}, grow {grow}: {actual} vs {expected}");
            assert_eq!(capacity_before, inline_math_typographic_extent(&block, size).unwrap());
        }
    }

    #[test]
    fn math_face_only_expression_retains_its_baseline_policy() {
        let block = MathBlock::Inline(vec![MathExpr::Nary {operator_color: None, op: '\u{222b}',
            sub: Some(Box::new(MathExpr::Text("a".into()))), sup: None,
            operand: Box::new(MathExpr::Text("f(x)".into())),
            lim_loc: crate::ir::LimLoc::SubSup, grow: false}]);
        assert!(inline_math_baseline_extent(&block, 10.5).is_none());
    }
}


#[cfg(test)]
mod extended_joint_limit_baseline_regression_tests {
    use super::*;

    #[test]
    fn jointly_placed_limits_retain_saved_word_host_baselines() {
        for (op, grow, size, line_height, expected) in [
            ('\u{222b}', false, 10.5, 24.0, 16.4500122),
            ('\u{222b}', false, 14.0, 36.0, 24.0100098),
            ('\u{2211}', true, 10.5, 24.0, 16.4500122),
            ('\u{2211}', true, 14.0, 36.0, 24.0100098),
        ] {
            let limit = |text: &str| MathExpr::Run { text: text.into(),
                style: crate::ir::MathRunStyle {
                    math_style: Some(crate::ir::MathStyleVariant::Plain),
                    run_style: Some(crate::ir::RunStyle {
                        font_family: Some("MS Mincho".into()), font_size: Some(size),
                        ..Default::default()
                    }), ..Default::default()
                }
            };
            let block = MathBlock::Inline(vec![MathExpr::Nary {operator_color: None,  op,
                sub: Some(Box::new(limit("\u{03b1}"))),
                sup: Some(Box::new(limit("\u{03b2}"))),
                operand: Box::new(MathExpr::Text("f(x)g(x)dx".into())),
                lim_loc: crate::ir::LimLoc::SubSup, grow
            }]);
            let (a, d) = inline_math_baseline_extent(&block, size).unwrap();
            let host = crate::font::catalog_glyph_face_metrics("MS Mincho", false, false).unwrap();
            let (ha, hd) = host.design_font_box_pt(size, true);
            let baseline = (line_height + a.max(ha) - d.max(hd)) * 0.5;
            assert!((baseline - expected).abs() <= 0.35,
                "op {op}, size {size}: {baseline} vs {expected}");
        }
    }
}


/// A legacy leaf has no MATH constants. Its measured rule thickness supplies
/// clearance for ordinary joint limits; MATH leaves keep their own policy.
fn ordinary_fallback_rule_gap(expr: &MathExpr, ctx: &MathLayoutContext) -> Option<f32> {
    match expr {
        MathExpr::Run { text, style } => {
            let glyphs = resolved_run_glyphs(text, style, ctx)?;
            if glyphs.is_empty() || glyphs.iter().any(|g| g.metrics.has_math) { return None; }
            let run = style.run_style.as_ref()?;
            let thickness = crate::font::catalog_glyph_rule_thickness(
                run.font_family.as_deref()?, run.bold, run.italic)?;
            let nominal_size = resolved_run_context(style, ctx).font_size;
            (nominal_size.is_finite() && nominal_size > 0.0)
                .then_some(4.0 * thickness * nominal_size)
        }
        MathExpr::Seq(children) => {
            let gaps: Option<Vec<f32>> = children.iter()
                .map(|child| ordinary_fallback_rule_gap(child, ctx)).collect();
            gaps?.into_iter().reduce(f32::max)
        }
        MathExpr::BoxExpr(child) | MathExpr::Phantom(child) => ordinary_fallback_rule_gap(child, ctx),
        _ => None,
    }
}

#[cfg(test)]
mod ordinary_joint_limit_relative_position_tests {
    use super::*;

    #[test]
    fn ordinary_joint_limits_retain_saved_relative_positions_at_two_sizes() {
        for (size, expected_lower, expected_upper) in [
            (10.5, 5.5200195, -6.1199951),
            (14.0, 7.4400024, -8.1600342),
        ] {
            let limit = |text: &str| MathExpr::Run { text: text.into(),
                style: crate::ir::MathRunStyle {
                    math_style: Some(crate::ir::MathStyleVariant::Plain),
                    run_style: Some(crate::ir::RunStyle {
                        font_family: Some("MS Mincho".into()), font_size: Some(size),
                        ..Default::default()
                    }), ..Default::default()
                }
            };
            let block = MathBlock::Inline(vec![MathExpr::Nary {operator_color: None,
                op: '\u{2211}', sub: Some(Box::new(limit("\u{03b1}"))),
                sup: Some(Box::new(limit("\u{03b2}"))),
                operand: Box::new(MathExpr::Text("f(x)g(x)dx".into())),
                lim_loc: crate::ir::LimLoc::SubSup, grow: false,
            }]);
            let (elements, _) = emit_math_block(&block, 0.0, 0.0, size);
            let baseline = |text: &str| {
                let e = elements.iter().find(|e| matches!(&e.content,
                    LayoutContent::Text { text: value, .. } if value == text)).unwrap();
                e.y + e.baseline_offset.unwrap()
            };
            // The legacy Text leaf substitutes mathematical italic letters.
            // It may remain one shaping run or emit selected glyphs separately;
            // both carry the same operand baseline. Match its actual first f.
            let operand_element = elements.iter().find(|e| matches!(&e.content,
                LayoutContent::Text { text, .. }
                    if text.starts_with(math_substitute('f'))))
                .expect("emitted mathematical f operand");
            let operand = operand_element.y + operand_element.baseline_offset
                .expect("emitted math operand baseline");
            assert!((baseline("\u{03b1}") - operand - expected_lower).abs() <= 0.15);
            assert!((baseline("\u{03b2}") - operand - expected_upper).abs() <= 0.15);
        }
    }
}

#[cfg(test)]
mod operator_color_regression_tests {
    use super::*;

    #[test]
    fn nary_operator_color_is_independent_of_operand_and_limits() {
        let leaf = |text: &str, color: &str| MathExpr::Run { text: text.into(),
            style: crate::ir::MathRunStyle { run_style: Some(crate::ir::RunStyle {
                font_family: Some("Cambria Math".into()), color: Some(color.into()),
                ..crate::ir::RunStyle::default()
            }), ..crate::ir::MathRunStyle::default() } };
        // '+' exercises the non-MATH-table fallback as well as the real stretch plans.
        for op in ['∑', '∫', '+'] {
            for style in [MathStyle::Display, MathStyle::Text] {
                let expr = MathExpr::Nary { op, operator_color: Some("#FF0000".into()),
                    sub: Some(Box::new(leaf("n", "#00FF00"))), sup: None,
                    operand: Box::new(leaf("x", "#0000FF")), lim_loc: crate::ir::LimLoc::SubSup, grow: true };
                let ctx = MathLayoutContext { font_size: 14.0, style };
                let (elements, bounds) = emit_expr(&expr, 12.0, 100.0, &ctx);
                let mut operators = 0;
                for e in &elements {
                    if let LayoutContent::Text { text, color, .. } = &e.content {
                        let expected = if text == &op.to_string() { operators += 1; "#FF0000" }
                            else if text == "n" || text == "𝑛" { "#00FF00" } else { "#0000FF" };
                        assert_eq!(color.as_deref(), Some(expected), "{op}: {text}");
                    }
                }
                assert!(operators > 0);
                let mut uncolored = expr.clone();
                if let MathExpr::Nary { operator_color, .. } = &mut uncolored { *operator_color = None; }
                let (before, before_bounds) = emit_expr(&uncolored, 12.0, 100.0, &ctx);
                assert_eq!(elements.len(), before.len());
                assert_eq!((bounds.advance,bounds.ascent,bounds.descent,bounds.italic_correction),
                    (before_bounds.advance,before_bounds.ascent,before_bounds.descent,before_bounds.italic_correction));
                for (after, mut before) in elements.iter().zip(before) {
                    if let LayoutContent::Text { text, color, .. } = &mut before.content {
                        if text == &op.to_string() { *color = Some("#FF0000".into()); }
                    }
                    assert_eq!((after.x, after.y, after.width, after.height, after.text_y_off, after.baseline_offset),
                        (before.x, before.y, before.width, before.height, before.text_y_off, before.baseline_offset));
                    assert_eq!(after.font_glyph.as_ref().map(|g| (g.index, g.bounds_em)),
                        before.font_glyph.as_ref().map(|g| (g.index, g.bounds_em)));
                    assert_eq!(std::mem::discriminant(&after.content), std::mem::discriminant(&before.content));
                    if let (LayoutContent::Text { text: a, color: ac, font_size: afs, font_family: aff, .. },
                        LayoutContent::Text { text: b, color: bc, font_size: bfs, font_family: bff, .. }) = (&after.content, &before.content) {
                        assert_eq!((a, ac, afs, aff), (b, bc, bfs, bff));
                    }
                }
            }
        }
    }

    #[test]
    fn older_nary_ir_without_operator_color_remains_readable() {
        let expr: MathExpr = serde_json::from_str(r#"{"Nary":{"op":"∑","sub":null,"sup":null,"operand":{"Text":"x"},"lim_loc":"SubSup","grow":false}}"#).unwrap();
        assert!(matches!(expr, MathExpr::Nary { operator_color: None, .. }));
    }
}
