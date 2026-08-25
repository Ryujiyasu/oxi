// This Source Code Form is subject to the terms of the Mozilla Public
// License, v. 2.0. If a copy of the MPL was not distributed with this
// file, You can obtain one at https://mozilla.org/MPL/2.0/.

//! Oxi PPTX Renderer — renders .pptx slides for pixel-accurate comparison
//! against PowerPoint (deck.pdf, produced by measure_pptx_word.py).
//!
//! Usage:
//!   oxi-pptx-renderer <input.pptx> <output_prefix> [dpi] [--dump-layout=PATH]
//!                     [--supersample=N]
//!
//! Default dpi=150, supersample=3. Produces `<prefix>_s1.png`, `<prefix>_s2.png`
//! ... (one per slide). With `--dump-layout=PATH` it writes the slide-level
//! layout JSON (points) and exits without rendering — this is the Oxi-side
//! measurement target for the pptx Ra loop (each shape's bbox/type/text is
//! compared against the PowerPoint COM oracle's truth.json).
//!
//! Text rendering implements the measured text-frame layout (Spec #4):
//! word-wrap within the effective width (shape width - 14.4pt insets),
//! line advance = font_size x 1.2 x n (n = lnSpc multiple, 1.0 default),
//! first-line baseline from the measured models (multi: +0.75 x advance /
//! single: +A_font x fs), space_before / space_after added per paragraph.

mod emoji;
mod font_adv;

use oxislides_core::ir::{
    GeomCmd, LineEnd, MasterStyleLevel, Presentation, Shape, ShapeContent, SlideAlignment,
    SlideBackgroundImage, SlideBullet, SlideGradient,
};
use serde_json::{json, Value};

fn main() {
    let args: Vec<String> = std::env::args().collect();
    if args.len() < 3 {
        eprintln!(
            "Usage: {} <input.pptx> <output_prefix> [dpi] [--dump-layout=PATH] [--supersample=N]",
            args[0]
        );
        std::process::exit(1);
    }

    let pptx_path = &args[1];
    let output_prefix = &args[2];
    let dpi: u32 = args.get(3).and_then(|s| s.parse().ok()).unwrap_or(150);

    let mut dump_layout: Option<String> = None;
    // The oversampling factor the page is drawn at before it is filtered down.
    // Found on d32 slide 6, whose map is fine line art: 1 / 2 / 3 score
    // 0.8606 / 0.8769 / 0.8804 against PowerPoint. Over the whole dev corpus
    // 3 is worth **+0.002518, 39 decks improved and none regressed** -- the
    // gain is not confined to pictures (d01, all vector, gains 0.0086) because
    // every edge on the page is sampled better.
    //
    // It costs 2.4x the render time (74.7s -> 178.5s for d13's 20 slides), and
    // that trade was made deliberately in favour of fidelity (2026-08-21).
    // `OXI_SUPERSAMPLE` overrides it, and 2 reproduces the pre-change output
    // byte for byte.
    let mut supersample: u32 = std::env::var("OXI_SUPERSAMPLE")
        .ok()
        .and_then(|v| v.parse().ok())
        .filter(|n| (1..=4).contains(n))
        .unwrap_or(3);
    for arg in &args[3..] {
        if let Some(path) = arg.strip_prefix("--dump-layout=") {
            dump_layout = Some(path.to_string());
        }
        if let Some(n) = arg.strip_prefix("--supersample=") {
            supersample = n.parse().unwrap_or(2);
        }
    }

    let data = std::fs::read(pptx_path).expect("Cannot read pptx file");
    let pres = oxislides_core::parser::parse_pptx(&data).expect("Cannot parse pptx");

    eprintln!(
        "Parsed {} slides, size={}x{}pt, DPI={} supersample={}x",
        pres.slides.len(),
        pres.slide_width,
        pres.slide_height,
        dpi,
        supersample
    );

    #[cfg(windows)]
    {
        let n = install_embedded_fonts(&pres);
        if n > 0 {
            eprintln!("Installed {}/{} embedded fonts", n, pres.embedded_fonts.len());
        }
        let cloud = install_cloud_fonts();
        if cloud > 0 {
            eprintln!("Installed {cloud} cloud-cache fonts");
        }
    }

    if let Some(path) = dump_layout {
        #[cfg(windows)]
        dump_layout_json_gdi(&pres, &path);
        #[cfg(not(windows))]
        dump_layout_json_plain(&pres, &path);
        eprintln!("Layout dumped to {}", path);
        return;
    }

    #[cfg(windows)]
    {
        render_slides_gdi(&pres, output_prefix, dpi, supersample);
    }

    #[cfg(not(windows))]
    {
        eprintln!("GDI rendering requires Windows");
        std::process::exit(1);
    }
}

/// Slide-level layout JSON in points — the Oxi-side measurement target.
/// Plain (no GDI): paragraphs carry only runs (no wrapped-line positions).
#[cfg(not(windows))]
fn dump_layout_json_plain(pres: &Presentation, path: &str) {
    let slides: Vec<Value> = pres
        .slides
        .iter()
        .map(|slide| {
            let shapes: Vec<Value> = slide.shapes.iter().map(shape_json).collect();
            json!({
                "index": slide.index,
                "width": pres.slide_width,
                "height": pres.slide_height,
                "background_color": slide.background_color,
                "shapes": shapes,
            })
        })
        .collect();
    let json = json!({
        "presentation": {
            "width": pres.slide_width,
            "height": pres.slide_height,
        },
        "slides": slides,
    });
    let text = serde_json::to_string_pretty(&json).expect("Cannot serialize layout");
    std::fs::write(path, text).expect("Cannot write layout JSON");
}

/// Slide-level layout JSON in points, computed with GDI font metrics so that
/// each text paragraph also carries its wrapped line baselines (slide-absolute
/// y in points) — the Oxi-side target for the text-frame layout spec.
#[cfg(windows)]
fn dump_layout_json_gdi(pres: &Presentation, path: &str) {
    use windows::Win32::Foundation::*;
    use windows::Win32::Graphics::Gdi::*;

    let slides: Vec<Value> = pres
        .slides
        .iter()
        .map(|slide| {
            let shapes: Vec<Value> = slide
                .shapes
                .iter()
                .map(|sh| {
                    let mut v = shape_json(sh);
                    // Attach wrapped-line baselines for text-bearing shapes.
                    if let Some(p) = v.get_mut("content") {
                        let text_shape = matches!(
                            &sh.content,
                            ShapeContent::TextBox { .. } | ShapeContent::AutoShape { .. }
                        );
                        if text_shape {
                            let scale = 1.0; // points
                            let dc = unsafe { GetDC(HWND(std::ptr::null_mut())) };
                            if let Some(paragraphs) = p.get_mut("paragraphs") {
                                if let Some(arr) = paragraphs.as_array_mut() {
                                    let (gh, gv) = geom_text_inset(sh);
                                    let mut cursor_pt = sh.y + sh.t_ins + gv;
                                    let anchor_off = compute_shape_anchor_off(dc, pres, sh);
                                    let master_ctx: &Vec<MasterStyleLevel> =
                                        match sh.ph_type.as_deref() {
                                            Some("title") | Some("ctrTitle") => {
                                                &pres.master_styles.title
                                            }
                                            Some(_) => &pres.master_styles.body,
                                            None => &pres.master_styles.other,
                                        };
                                    // Spec #11: one AutoNum counter set per text box.
                                    let mut counters = std::collections::HashMap::<
                                        (u32, String),
                                        (Option<u32>, u32),
                                    >::new();
                                    let mut prev_fs: Option<f32> = None;
                                    for (i, para_json) in arr.iter_mut().enumerate() {
                                        if let Some(para) = sh_para(&sh.content, i) {
                                            let def_family = resolve_font(pres, sh);
                                            let (bases, marker) = layout_paragraph_baselines(
                                                dc,
                                                para,
                                                &mut cursor_pt,
                                                sh.width,
                                                scale,
                                                i == 0,
                                                &def_family,
                                                sh.l_ins + gh,
                                                sh.r_ins + gh,
                                                &master_ctx[..],
                                                &sh.ph_levels[..],
                                                anchor_off,
                                                &mut counters,
                                                &mut prev_fs,
                                                sh.wrap_text,
                                            );
                                            // Spec #11: surface the marker text so autonum
                                            // number strings can be verified in the dump.
                                            para_json["marker"] = json!(marker.map(|m| m.text));
                                            para_json["line_baselines"] = json!(
                                                bases
                                                    .iter()
                                                    .map(|(_, b, _)| (b * 100.0).round() / 100.0)
                                                    .collect::<Vec<_>>()
                                            );
                                            para_json["line_x_offsets"] = json!(
                                                bases
                                                    .iter()
                                                    .map(|(_, _, x)| (x * 100.0).round() / 100.0)
                                                    .collect::<Vec<_>>()
                                            );
                                        }
                                    }
                                }
                            }
                            unsafe {
                                let _ = ReleaseDC(HWND(std::ptr::null_mut()), dc);
                            }
                        }
                    }
                    v
                })
                .collect();
            json!({
                "index": slide.index,
                "width": pres.slide_width,
                "height": pres.slide_height,
                "background_color": slide.background_color,
                "shapes": shapes,
            })
        })
        .collect();
    let json = json!({
        "presentation": {
            "width": pres.slide_width,
            "height": pres.slide_height,
        },
        "slides": slides,
    });
    let text = serde_json::to_string_pretty(&json).expect("Cannot serialize layout");
    std::fs::write(path, text).expect("Cannot write layout JSON");
}

/// Convenience: the i-th paragraph of a text shape, or None.
fn sh_para(content: &ShapeContent, i: usize) -> Option<&oxislides_core::ir::SlideParagraph> {
    match content {
        ShapeContent::TextBox { paragraphs } | ShapeContent::AutoShape { paragraphs } => {
            paragraphs.get(i)
        }
        _ => None,
    }
}

/// All paragraphs of a text shape, or None (non-text shapes).
fn sh_paragraphs(content: &ShapeContent) -> Option<&Vec<oxislides_core::ir::SlideParagraph>> {
    match content {
        ShapeContent::TextBox { paragraphs } | ShapeContent::AutoShape { paragraphs } => {
            Some(paragraphs)
        }
        _ => None,
    }
}

/// Spec #6: the vertical-anchor offset for a text shape.
///
/// `a:bodyPr/@anchor` (resolved through the placeholder chain by the parser,
/// stored on `Shape.anchor`) shifts the whole text block within the inner area
/// (shape minus t_ins/b_ins):
///   * "ctr" -> offset = (inner_h - block_h) / 2  (vertically centred)
///   * "b"   -> offset = (inner_h - block_h)      (pushed to the bottom)
///   * "t" / None -> 0.0 (top-aligned; the default)
///
/// `block_h` is the measured height of the text block = the cursor advance
/// across all paragraphs with anchor_off = 0, MINUS the first paragraph's
/// `first_off` (the ascent-based first-line placement is not part of the block
/// height — the centring law is `baseline = inner_top + (inner_h - block_h)/2
/// + first_line_ascent`, anchor_probe / anchor_trigger render-truth).
///
/// A block TALLER than the inner area is NOT clamped: both anchors keep their
/// formula and let the block overflow (upward for "b", equally for "ctr"), as
/// PowerPoint does — `anchorb` probe, 12 arms, and d24 s1 for the centre.
#[cfg(windows)]
fn compute_shape_anchor_off(
    dc: windows::Win32::Graphics::Gdi::HDC,
    pres: &Presentation,
    sh: &Shape,
) -> f32 {
    let anchor = sh.anchor.as_deref();
    if anchor != Some("ctr") && anchor != Some("b") {
        return 0.0;
    }
    let Some(paragraphs) = sh_paragraphs(&sh.content) else {
        return 0.0;
    };
    if paragraphs.is_empty() {
        return 0.0;
    }
    let (geom_h_ins, geom_v_ins) = geom_text_inset(sh);
    let inner_h = (sh.height - sh.t_ins - sh.b_ins - 2.0 * geom_v_ins).max(0.0);
    let master_ctx: &Vec<MasterStyleLevel> = match sh.ph_type.as_deref() {
        Some("title") | Some("ctrTitle") => &pres.master_styles.title,
        Some(_) => &pres.master_styles.body,
        None => &pres.master_styles.other,
    };
    let def_family = resolve_font(pres, sh);
    let mut cursor_pt = 0.0_f32;
    let scale = 1.0_f64;
    // Spec #11: AutoNum counters are per text box; this measurement pass only
    // needs the block height, so a throwaway counter set is fine.
    let mut counters = std::collections::HashMap::<(u32, String), (Option<u32>, u32)>::new();
    let mut prev_fs: Option<f32> = None;
    for (i, para) in paragraphs.iter().enumerate() {
        let _ = layout_paragraph_baselines(
            dc,
            para,
            &mut cursor_pt,
            sh.width,
            scale,
            i == 0,
            &def_family,
            sh.l_ins + geom_h_ins,
            sh.r_ins + geom_h_ins,
            &master_ctx[..],
            &sh.ph_levels[..],
            0.0,
            &mut counters,
            &mut prev_fs,
            sh.wrap_text,
        );
    }
    // block_h = the block's advance minus the first paragraph's first_off.
    let para = &paragraphs[0];
    let fs = paragraph_font_size(
        para,
        sh.ph_levels
            .first()
            .and_then(|l| l.font_size)
            .or_else(|| master_ctx.first().and_then(|m| m.font_size)),
        None,
    );
    let first_off = {
        let n = para
            .line_spacing
            .or_else(|| sh.ph_levels.first().and_then(|l| l.line_spacing))
            .unwrap_or(1.0);
        let family = para
            .runs
            .iter()
            .find_map(|r| r.font_family.clone())
            .unwrap_or_else(|| def_family.clone());
        first_baseline_off(&family, fs, n)
    };
    // With the line-box cursor the run already ends at the block's bottom;
    // the old cursor ended one ascent lower, hence the subtraction.
    let block_h = if mixpitch_on() {
        cursor_pt.max(0.0)
    } else {
        (cursor_pt - first_off).max(0.0)
    };
    if anchor == Some("ctr") {
        // ★A block TALLER than its box still centres on it: d24 slide 1's
        // 60pt title needs 178pt in a 91pt box and PowerPoint puts the block's
        // centre (195.6) on the box's (202.4), overflowing equally above and
        // below. Clamping the offset at zero pinned it to the box top, 63pt
        // low. Measured 2026-08-18 from PowerPoint's own render.
        //
        // Gated with the placeholder-style inheritance because it belongs to
        // the same change: leaving it outside made the opt-out arm differ on 9
        // decks, so the A/B was not a before-vs-after.
        if phlevel_on() {
            (inner_h - block_h) / 2.0
        } else {
            (inner_h - block_h).max(0.0) / 2.0
        }
    } else if anchorb_on() {
        // ★The BOTTOM anchor does not clamp either. Probe `anchorb`, 12 arms
        // (t / ctr / b x 1..4 lines of 32pt in a 57.6pt text area): PowerPoint
        // holds the LAST baseline at 201.53 for two, three and four lines and
        // lets the block run off the TOP of the box, matching the unclamped
        // offset to 0.03pt at every count. Clamping at zero pinned the block
        // to the box top instead -- one whole line-height per overflowing
        // line too low. 246 shapes on the dev corpus's slides anchor bottom,
        // all in the eight SlidesCarnival-family decks.
        inner_h - block_h
    } else {
        (inner_h - block_h).max(0.0)
    }
}

/// Theme-default font resolution (Ra loop, theme_default probes, PPTX COM/PDF
/// render-truth): a run with NO explicit `font.name` is rendered in the theme
/// font for its context —
///   * title placeholders (`p:ph @type="title"`) use the theme MAJOR font
///     (`<a:majorFont><a:latin typeface=.../>`),
///   * everything else (plain textboxes, body placeholders, table cells) uses
///     the theme MINOR font (`<a:minorFont><a:latin typeface=.../>`).
/// An explicit run font (already resolved by the caller via `font_family`) wins
/// over the theme; this function is only the fallback for unset runs.
/// The face a paragraph asks for, through the chain ECMA-376 defines: the run
/// first, then the PLACEHOLDER's own `a:lstStyle` (layout, then master), then
/// the master `p:txStyles` level, and only then the theme's major/minor.
///
/// d15's master `body` placeholder declares Barlow Light while every level Oxi
/// used to read -- master txStyles, theme minor, presentation default -- says
/// Arial, so its body text came out in the wrong face and measured wider,
/// wrapping a word early. 188 layout/master placeholders in the dev corpus name
/// a font this way.
fn paragraph_family(
    pres: &Presentation,
    sh: &Shape,
    para: &oxislides_core::ir::SlideParagraph,
    ph_levels: &[MasterStyleLevel],
    master: &[MasterStyleLevel],
) -> String {
    let chosen = paragraph_family_inner(pres, sh, para, ph_levels, master);
    if let Ok(want) = std::env::var("OXI_FAM_DEBUG") {
        let text: String = para.runs.iter().map(|r| r.text.as_str()).collect();
        if text.contains(&want) {
            let fams = |ls: &[MasterStyleLevel]| -> Vec<Option<String>> {
                ls.iter().map(|l| l.font_family.clone()).collect()
            };
            eprintln!(
                "FAM ph_type={:?} lvl={} ph_levels={:?} master={:?} -> {chosen}",
                sh.ph_type,
                para.lvl,
                fams(ph_levels),
                fams(master),
            );
        }
    }
    chosen
}

fn paragraph_family_inner(
    pres: &Presentation,
    sh: &Shape,
    para: &oxislides_core::ir::SlideParagraph,
    ph_levels: &[MasterStyleLevel],
    master: &[MasterStyleLevel],
) -> String {
    if let Some(f) = para.runs.iter().find_map(|r| r.font_family.clone()) {
        return f;
    }
    if phfont_on() {
        let lvl = para.lvl as usize;
        for levels in [ph_levels, master] {
            if let Some(l) = levels.get(lvl.min(levels.len().saturating_sub(1))) {
                if let Some(f) = l.font_family.clone() {
                    return f;
                }
            }
        }
    }
    resolve_font(pres, sh)
}

/// A placeholder's declared face is honoured unless this is set, which restores
/// resolving straight to the theme.
fn phfont_on() -> bool {
    std::env::var("OXI_PHFONT_DISABLE").is_err()
}

fn resolve_font(pres: &Presentation, sh: &Shape) -> String {
    match sh.ph_type.as_deref() {
        // A centered title placeholder is still a TITLE: it uses the theme
        // MAJOR font like a plain "title" (Word render-truth, anchor_trigger
        // V4: ctrTitle renders in Georgia = the theme major face).
        Some("title") | Some("ctrTitle") => pres.major_font.clone(),
        _ => pres.minor_font.clone(),
    }
}

fn alignment_str(a: Option<SlideAlignment>) -> &'static str {
    match a {
        Some(SlideAlignment::Left) => "left",
        Some(SlideAlignment::Center) => "center",
        Some(SlideAlignment::Right) => "right",
        Some(SlideAlignment::Justify) => "justify",
        // Spec #6: a paragraph with no alignment anywhere in the resolution
        // chain (run -> paragraph -> master txStyles level) is "inherit" in
        // the dump — the renderer resolves it per the chain.
        None => "inherit",
    }
}

fn run_json(text: &str, font_size: Option<f32>, bold: bool, italic: bool, color: &Option<String>, font_family: &Option<String>) -> Value {
    json!({
        "text": text,
        "font_size": font_size,
        "bold": bold,
        "italic": italic,
        "color": color,
        "font_family": font_family,
    })
}

fn paragraphs_json(paragraphs: &[oxislides_core::ir::SlideParagraph]) -> Value {
    let items: Vec<Value> = paragraphs
        .iter()
        .map(|p| {
            json!({
                "alignment": alignment_str(p.alignment),
                "line_spacing": p.line_spacing,
                "space_before": p.space_before,
                "space_after": p.space_after,
                "runs": p
                    .runs
                    .iter()
                    .map(|r| run_json(&r.text, r.font_size, r.bold, r.italic, &r.color, &r.font_family))
                    .collect::<Vec<_>>(),
            })
        })
        .collect();
    let text: String = paragraphs
        .iter()
        .flat_map(|p| p.runs.iter().map(|r| r.text.as_str()))
        .collect();
    json!({
        "text": text,
        "paragraphs": items,
    })
}

fn shape_json(sh: &Shape) -> Value {
    let (shape_type, extra) = match &sh.content {
        ShapeContent::AutoShape { paragraphs } => {
            ("autoshape", paragraphs_json(paragraphs))
        }
        ShapeContent::TextBox { paragraphs } => ("text", paragraphs_json(paragraphs)),
        ShapeContent::Table { table } => (
            "table",
            json!({
                "col_widths": table.col_widths,
                "row_heights": table.row_heights,
                "rows": table
                    .rows
                    .iter()
                    .map(|row| {
                        row.iter()
                            .map(|cell| {
                                let mut v = paragraphs_json(&cell.paragraphs);
                                if let Some(o) = v.as_object_mut() {
                                    o.insert("grid_span".into(), json!(cell.grid_span));
                                    o.insert("h_merge".into(), json!(cell.h_merge));
                                    o.insert(
                                        "end_para_sizes".into(),
                                        json!(cell
                                            .paragraphs
                                            .iter()
                                            .map(|p| p.end_para_size)
                                            .collect::<Vec<_>>()),
                                    );
                                }
                                v
                            })
                            .collect::<Vec<_>>()
                    })
                    .collect::<Vec<_>>(),
            }),
        ),
        ShapeContent::Image { data, content_type } => (
            "image",
            json!({
                "content_type": content_type,
                "image_bytes": data.len(),
                "src_rect": sh.src_rect,
                "fill_rect": sh.fill_rect,
            }),
        ),
        ShapeContent::Unsupported { element_type } => ("unsupported", json!({ "element_type": element_type })),
        ShapeContent::Placeholder => ("placeholder", json!({})),
        ShapeContent::Chart { chart } => (
            "chart",
            json!({
                "chart_type": chart.chart_type,
                "bar_dir": chart.bar_dir,
                "grouping": chart.grouping,
                "hole_size": chart.hole_size,
                "bubble_scale": chart.bubble_scale,
                "size_represents": chart.size_represents,
                "series": chart
                    .series
                    .iter()
                    .map(|s| json!({
                        "name": s.name,
                        "values": s.values,
                        "x_values": s.x_values,
                        "sizes": s.sizes,
                        "line_none": s.line_none,
                        "marker_none": s.marker_none,
                    }))
                    .collect::<Vec<_>>(),
                "categories": chart.categories,
                "has_legend": chart.has_legend,
                "auto_title_deleted": chart.auto_title_deleted,
                "explicit_title": chart.explicit_title,
                "marker": chart.marker,
            }),
        ),
    };
    json!({
        "x": sh.x,
        "y": sh.y,
        "w": sh.width,
        "h": sh.height,
        "rotation": sh.rotation,
        "shape_type": sh.shape_type,
        "adjustments": sh.adjustments,
        "type": shape_type,
        "fill_color": sh.fill_color,
        "fill_alpha": sh.fill_alpha,
        "border_color": sh.border_color,
        "border_width": sh.border_width,
        "border_dash": sh.border_dash,
        "head_end": sh.head_end.as_ref().map(|e| json!([e.kind, e.w, e.len])),
        "tail_end": sh.tail_end.as_ref().map(|e| json!([e.kind, e.w, e.len])),
        "text_warp": sh.text_warp,
        "anchor": sh.anchor,
        // The gradient was ABSENT from this dump, so every question asked of it
        // came back `null` and read as "the parser lost it" -- it had not.
        "gradient": sh.gradient.as_ref().map(|g| json!({
            "angle_deg": g.angle_deg,
            "scaled": g.scaled,
            "focus": g.focus,
            "stops": g.stops.iter().map(|s| json!({
                "pos": s.pos, "color": s.color, "alpha": s.alpha
            })).collect::<Vec<_>>(),
        })),
        "l_ins": sh.l_ins,
        "r_ins": sh.r_ins,
        "t_ins": sh.t_ins,
        "b_ins": sh.b_ins,
        "content": extra,
    })
}

/// Parse a `#RRGGBB` or `RRGGBB` hex color into (r, g, b) bytes. Defaults to
/// (0, 0, 0) for malformed input.
fn parse_hex_rgb(s: &str) -> Option<(u8, u8, u8)> {
    let c = s.strip_prefix('#').unwrap_or(s);
    if c.len() != 6 {
        return None;
    }
    let r = u8::from_str_radix(&c[0..2], 16).ok()?;
    let g = u8::from_str_radix(&c[2..4], 16).ok()?;
    let b = u8::from_str_radix(&c[4..6], 16).ok()?;
    Some((r, g, b))
}

/// `a:prstGeom prst="pie"` is drawn as its wedge unless this is set.
fn pie_on() -> bool {
    std::env::var("OXI_PIE_DISABLE").is_err()
}

fn colorref(r: u8, g: u8, b: u8) -> u32 {
    (r as u32) | ((g as u32) << 8) | ((b as u32) << 16)
}

/// Draw the highest-frequency non-rectangular DrawingML presets as their real
/// geometry. Coordinates are built in the shape's local box, then flipped and
/// rotated about its centre before conversion to device pixels.
#[cfg(windows)]
/// Emit the shape's outline as the DC's current path, between BeginPath and
/// EndPath. Returns false when the preset has no geometry here, and then no
/// path has been started.
///
/// Split out of `draw_preset_shape_gdi` so a picture can be CLIPPED to the same
/// outline: 91 `p:pic` across 16 dev decks state a non-rectangular geometry --
/// 53 ellipse, 19 roundRect, 13 custGeom, 6 round2SameRect -- and d11 slide 33's
/// four portraits are circles in PowerPoint and squares in Oxi.
#[cfg(windows)]
unsafe fn emit_shape_path(dc: windows::Win32::Graphics::Gdi::HDC, sh: &Shape, scale: f64) -> bool {
    use windows::Win32::Foundation::POINT;
    use windows::Win32::Graphics::Gdi::*;

    let prst = match sh.shape_type.as_deref() {
        Some(
            p @ ("ellipse" | "roundRect" | "homePlate" | "chevron" | "teardrop" | "pie"
                | "star10" | "wedgeRectCallout" | "blockArc"),
        ) if (p != "pie" || pie_on())
            && (p != "chevron" || chevron_on())
            && (p != "star10" || star10_on())
            && (p != "wedgeRectCallout" || wedgecallout_on())
            && (p != "blockArc" || blockarc_on()) =>
        {
            p
        }
        _ => return false,
    };
    let w = sh.width.max(0.0);
    let h = sh.height.max(0.0);
    if w == 0.0 || h == 0.0 {
        return false;
    }
    let map = |mut lx: f32, mut ly: f32| -> POINT {
        if sh.flip_h { lx = w - lx; }
        if sh.flip_v { ly = h - ly; }
        let dx = lx - w / 2.0;
        let dy = ly - h / 2.0;
        let (sn, cs) = sh.rotation.to_radians().sin_cos();
        let px = sh.x + w / 2.0 + dx * cs - dy * sn;
        let py = sh.y + h / 2.0 + dx * sn + dy * cs;
        POINT { x: (px as f64 * scale).round() as i32, y: (py as f64 * scale).round() as i32 }
    };
    let line = |p: POINT| { let _ = LineTo(dc, p.x, p.y); };
    let bezier = |c1: POINT, c2: POINT, end: POINT| { let _ = PolyBezierTo(dc, &[c1, c2, end]); };
    let _ = BeginPath(dc);
    const K: f32 = 0.552_284_8;
    match prst {
        "ellipse" => {
            let (rx, ry) = (w / 2.0, h / 2.0);
            let p = map(rx, 0.0); let _ = MoveToEx(dc, p.x, p.y, None);
            bezier(map(rx + K * rx, 0.0), map(w, ry - K * ry), map(w, ry));
            bezier(map(w, ry + K * ry), map(rx + K * rx, h), map(rx, h));
            bezier(map(rx - K * rx, h), map(0.0, ry + K * ry), map(0.0, ry));
            bezier(map(0.0, ry - K * ry), map(rx - K * rx, 0.0), map(rx, 0.0));
        }
        "roundRect" => {
            let adj = sh.adjustments.get("adj").copied().unwrap_or(16_667.0);
            let r = (w.min(h) * (adj / 100_000.0)).clamp(0.0, w.min(h) / 2.0);
            let p = map(r, 0.0); let _ = MoveToEx(dc, p.x, p.y, None);
            line(map(w - r, 0.0));
            bezier(map(w - r + K * r, 0.0), map(w, r - K * r), map(w, r));
            line(map(w, h - r));
            bezier(map(w, h - r + K * r), map(w - r + K * r, h), map(w - r, h));
            line(map(r, h));
            bezier(map(r - K * r, h), map(0.0, h - r + K * r), map(0.0, h - r));
            line(map(0.0, r));
            bezier(map(0.0, r - K * r), map(r - K * r, 0.0), map(r, 0.0));
        }
        "homePlate" => {
            let adj = sh.adjustments.get("adj").copied().unwrap_or(50_000.0);
            let d = (w.min(h) * (adj / 100_000.0)).clamp(0.0, w);
            let p = map(0.0, 0.0); let _ = MoveToEx(dc, p.x, p.y, None);
            line(map(w - d, 0.0)); line(map(w, h / 2.0));
            line(map(w - d, h)); line(map(0.0, h));
        }
        "chevron" => {
            // ECMA-376's `chevron`: `homePlate` with the SAME notch cut out of
            // its left edge, so a row of them interlocks. `adj` (default 50000)
            // is the horizontal run of the point as a fraction of the SHORTER
            // side, exactly as homePlate reads it.
            //
            // d35 s17's three process arrows are two of these and one
            // homePlate, all at 50% alpha -- and a translucent preset used to
            // lose its outline entirely, so PowerPoint's interlocking arrows
            // arrived as three plain boxes.
            let adj = sh.adjustments.get("adj").copied().unwrap_or(50_000.0);
            let d = (w.min(h) * (adj / 100_000.0)).clamp(0.0, w);
            let p = map(0.0, 0.0); let _ = MoveToEx(dc, p.x, p.y, None);
            line(map(w - d, 0.0)); line(map(w, h / 2.0));
            line(map(w - d, h)); line(map(0.0, h));
            line(map(d, h / 2.0));
        }
        "pie" => {
            // ECMA-376's `pie`: a wedge of the box's ellipse from `adj1` to
            // `adj2`, both in 60000ths of a degree measured clockwise from 3
            // o'clock (y grows downward), closed through the centre. 34 shapes
            // over 7 dev decks ask for one; d19 slide 30 is four of them at
            // adj1=10788866 adj2=16200000, rotated 0 / 90 / 180 / 270 to make
            // the SWOT circle Oxi was drawing as a teal square.
            let st = sh.adjustments.get("adj1").copied().unwrap_or(0.0);
            let en = sh.adjustments.get("adj2").copied().unwrap_or(16_200_000.0);
            let mut sw = en - st;
            if sw <= 0.0 {
                sw += 21_600_000.0;
            }
            let (rx, ry) = (w / 2.0, h / 2.0);
            let at = |units: f32| {
                let a = (units / 60_000.0).to_radians();
                (rx + rx * a.cos(), ry + ry * a.sin())
            };
            let (sx, sy) = at(st);
            let p = map(sx, sy);
            let _ = MoveToEx(dc, p.x, p.y, None);
            // One segment per degree keeps the polyline inside half a device
            // pixel of the true arc for any shape that fits on a slide.
            let steps = ((sw / 60_000.0).abs().ceil() as usize).clamp(2, 720);
            for i in 1..=steps {
                let (px, py) = at(st + sw * i as f32 / steps as f32);
                line(map(px, py));
            }
            line(map(rx, ry));
        }
        // S-STAR10 (2026-08-25): 14 of these are the snowflakes on d04 slide 13,
        // and with no path they were painted as their bounding box -- fourteen
        // white squares scattered over a world map, which is exactly the
        // "incorrect ink is worse than none" case this file keeps meeting.
        //
        // Read out of PowerPoint's own PDF vectors (a 20-point polyline per
        // star) and normalised to the box, the geometry is:
        //
        //     outer vertex k   theta = -90 + 36k   (hf*cos, sin)
        //     inner vertex k   phi   = -72 + 36k   (hf*r*cos, r*sin),  r = adj/50000
        //
        // in box coordinates where the box spans -1..1 in each axis. Every
        // measured point matches to about 0.1%: the outer u/cos ratios come out
        // 1.0524 / 1.0514 against the declared `hf` of 1.05146, and the inner
        // radius 0.5934 against `adj` 29731 / 50000 = 0.59462.
        //
        // ★The default `hf` is 105146 = 1/cos(18 degrees), which is precisely
        // what makes the outermost pair of points touch the left and right box
        // edges -- so the preset is self-consistent and the factor is not a
        // fudge. `adj`'s default (42533) is ECMA's; d04 states its own, so only
        // `hf` is confirmed by measurement here.
        "star10" => {
            let (rx, ry) = (w / 2.0, h / 2.0);
            let hf = sh.adjustments.get("hf").copied().unwrap_or(105_146.0) / 100_000.0;
            let r = sh.adjustments.get("adj").copied().unwrap_or(42_533.0) / 50_000.0;
            let vertex = |deg: f32, rad: f32| {
                let t = deg.to_radians();
                (rx + hf * rad * t.cos() * rx, ry + rad * t.sin() * ry)
            };
            let (sx, sy) = vertex(-90.0, 1.0);
            let p = map(sx, sy);
            let _ = MoveToEx(dc, p.x, p.y, None);
            for k in 0..10 {
                let kf = k as f32;
                if k > 0 {
                    let (ox, oy) = vertex(-90.0 + 36.0 * kf, 1.0);
                    line(map(ox, oy));
                }
                let (ix, iy) = vertex(-72.0 + 36.0 * kf, r);
                line(map(ix, iy));
            }
        }
        // S-WEDGECALL (2026-08-25): a rectangle with a triangular tail. Eight of
        // them in the corpus, one each on EIGHT different decks (d04 s13,
        // d06 s14, d11 s14, d15 s14, d16 s15, d19 s14, d24 s14, d35 s14) --
        // the same map-label template, drawn as a plain box with no tail.
        //
        // ECMA's guide chain, confirmed against PowerPoint's own PDF vectors on
        // three shapes covering BOTH exit branches, with the rule stated before
        // the last two were measured:
        //
        //   dxPos = adj1/100000 * w      dyPos = adj2/100000 * h
        //   tip   = (hc + dxPos, vc + dyPos)
        //   dz    = |dyPos| - |dxPos * h / w|
        //     dz > 0  -> the tail leaves through a HORIZONTAL edge (bottom if
        //               dyPos > 0 else top), base spanning w*g/12 for
        //               g = (7,10) if dxPos > 0 else (2,5)
        //     dz <= 0 -> a VERTICAL edge (right if dxPos > 0 else left), base
        //               spanning h*g/12 with g chosen the same way off dyPos
        //
        // d06 s14 (50.36x14.98, adj -21428/84287) gives tip (181.97,104.70) and
        // base 175.97..188.56 against PowerPoint's (181.97,104.70) and
        // 175.98..188.57. d15 s14 is the vertical branch -- predicted LEFT edge,
        // tip (201.28,135.81), base 132.90..136.88, all three drawn exactly.
        "wedgeRectCallout" => {
            let a1 = sh.adjustments.get("adj1").copied().unwrap_or(-20_833.0) / 100_000.0;
            let a2 = sh.adjustments.get("adj2").copied().unwrap_or(62_500.0) / 100_000.0;
            let (dx_pos, dy_pos) = (a1 * w, a2 * h);
            let (tx, ty) = (w / 2.0 + dx_pos, h / 2.0 + dy_pos);
            let horizontal = dy_pos.abs() > (dx_pos * h / w.max(1e-6)).abs();
            let g = |positive: bool| if positive { (7.0, 10.0) } else { (2.0, 5.0) };
            // The perimeter clockwise from the top-left, with the tail spliced
            // into whichever edge it leaves by.
            let mut pts: Vec<(f32, f32)> = Vec::with_capacity(7);
            let (g1, g2) = if horizontal { g(dx_pos > 0.0) } else { g(dy_pos > 0.0) };
            let (b1, b2) = if horizontal {
                (w * g1 / 12.0, w * g2 / 12.0)
            } else {
                (h * g1 / 12.0, h * g2 / 12.0)
            };
            pts.push((0.0, 0.0));
            if horizontal && dy_pos <= 0.0 {
                pts.push((b1, 0.0));
                pts.push((tx, ty));
                pts.push((b2, 0.0));
            }
            pts.push((w, 0.0));
            if !horizontal && dx_pos > 0.0 {
                pts.push((w, b1));
                pts.push((tx, ty));
                pts.push((w, b2));
            }
            pts.push((w, h));
            if horizontal && dy_pos > 0.0 {
                pts.push((b2, h));
                pts.push((tx, ty));
                pts.push((b1, h));
            }
            pts.push((0.0, h));
            if !horizontal && dx_pos <= 0.0 {
                pts.push((0.0, b2));
                pts.push((tx, ty));
                pts.push((0.0, b1));
            }
            let p = map(pts[0].0, pts[0].1);
            let _ = MoveToEx(dc, p.x, p.y, None);
            for &(px, py) in &pts[1..] {
                line(map(px, py));
            }
        }
        // S-BLOCKARC (2026-08-25): the LAST preset in the corpus with no path.
        // d24 slide 17 is three of them at rot 60 / 180 / -60 forming a donut,
        // and with no geometry Oxi filled the whole 218x218 box with the first
        // segment's orange -- a solid square where a ring belongs (0.9397).
        //
        // A ring sector: `adj1` / `adj2` are the start and end angles in
        // 60000ths of a degree (the same units and screen sense as `pie`), and
        // ★`adj3` is the ring's THICKNESS, not its inner radius:
        //
        //     inner radius = wd2 * (1 - adj3/50000)
        //
        // Measured off PowerPoint's own raster rather than guessed -- a scan
        // through the donut's centre row puts the coloured band at
        // **0.5835..0.9965** of the half-width, against 1 - 20773/50000 =
        // 0.58454 for the inner edge and 1.0 for the outer. The three segments
        // span 119.3 degrees each and, once each shape's own `rot` is applied,
        // tile 0..360 with the small gaps the deck draws.
        "blockArc" => {
            let st = sh.adjustments.get("adj1").copied().unwrap_or(10_800_000.0);
            let en = sh.adjustments.get("adj2").copied().unwrap_or(0.0);
            let a3 = sh.adjustments.get("adj3").copied().unwrap_or(25_000.0) / 50_000.0;
            let mut sw = en - st;
            if sw <= 0.0 {
                sw += 21_600_000.0;
            }
            let (rx, ry) = (w / 2.0, h / 2.0);
            let (irx, iry) = (rx * (1.0 - a3).max(0.0), ry * (1.0 - a3).max(0.0));
            let at = |units: f32, ax: f32, ay: f32| {
                let a = (units / 60_000.0).to_radians();
                (rx + ax * a.cos(), ry + ay * a.sin())
            };
            // One segment per degree, the same resolution the `pie` arm uses.
            let steps = ((sw / 60_000.0).abs().ceil() as usize).clamp(2, 720);
            let (sx, sy) = at(st, rx, ry);
            let p0 = map(sx, sy);
            let _ = MoveToEx(dc, p0.x, p0.y, None);
            for i in 1..=steps {
                let (px, py) = at(st + sw * i as f32 / steps as f32, rx, ry);
                line(map(px, py));
            }
            for i in 0..=steps {
                let (px, py) = at(st + sw * (steps - i) as f32 / steps as f32, irx, iry);
                line(map(px, py));
            }
        }
        "teardrop" => {
            // Default adj=100000: an ellipse whose upper-right quadrant is
            // pulled to the box corner. This is every corpus override too.
            let (rx, ry) = (w / 2.0, h / 2.0);
            let p = map(0.0, ry); let _ = MoveToEx(dc, p.x, p.y, None);
            bezier(map(0.0, ry - K * ry), map(rx - K * rx, 0.0), map(rx, 0.0));
            bezier(map(rx + w / 6.0, 0.0), map(rx + w / 3.0, 0.0), map(w, 0.0));
            bezier(map(w, h / 6.0), map(w, h / 3.0), map(w, ry));
            bezier(map(w, ry + K * ry), map(rx + K * rx, h), map(rx, h));
            bezier(map(rx - K * rx, h), map(0.0, ry + K * ry), map(0.0, ry));
        }
        _ => unreachable!(),
    }
    let _ = CloseFigure(dc);
    let _ = EndPath(dc);
    true
}

/// A picture is clipped to its shape's outline unless this is set.
fn picclip_on() -> bool {
    std::env::var("OXI_PICCLIP_DISABLE").is_err()
}

unsafe fn draw_preset_shape_gdi(dc: windows::Win32::Graphics::Gdi::HDC, sh: &Shape, scale: f64) -> bool {
    use windows::Win32::Foundation::COLORREF;
    use windows::Win32::Graphics::Gdi::*;

    // A translucent fill is composited by the caller's AlphaBlend path, which
    // this solid-brush path cannot reproduce along a bezier. Decline the shape
    // so it keeps the (rectangular) alpha-correct rendering: an opaque ellipse
    // where a 0%-alpha one belongs is the "incorrect ink is worse than none"
    // slab bug, and correct geometry does not buy back a wrong opacity.
    if sh.fill_color.is_some()
        && sh
            .fill_alpha
            .filter(|_| fill_alpha_on())
            .is_some_and(|a| a < 1.0)
    {
        return false;
    }
    // S-PRESETGRAD (2026-08-24): a PRESET whose only fill is a gradient has no
    // solid brush either. `draw_custom_geometry_gdi` has declined that case
    // since S-GEOMGRAD so `paint_shape_gradient` can clip the ramp to the
    // outline; the preset painter never learned it, and instead traced the
    // path with NULL_BRUSH, painted nothing, and reported SUCCESS -- which
    // stops the gradient painter from running at all.
    //
    // It was invisible until S-GRADSTOP because the generic `a:srgbClr`
    // handler used to catch a self-closing gradient STOP colour and leave it
    // in `fill_color`, so these shapes were painted flat in their last stop's
    // colour. d24 slide 30's SWOT circle is four `pie` wedges built exactly
    // that way: reading the stops properly turned the accidental flat fill off
    // and the circle vanished. Every non-rect preset with a gradient fill is
    // affected (`rect` never reaches the preset painter).
    if presetgrad_on() && sh.fill_color.is_none() && sh.gradient.is_some() {
        return false;
    }
    let w = sh.width.max(0.0);
    let h = sh.height.max(0.0);
    if w == 0.0 || h == 0.0 { return true; }

    let fill_brush = sh.fill_color.as_deref().and_then(parse_hex_rgb)
        .map(|c| CreateSolidBrush(COLORREF(colorref(c.0, c.1, c.2))));
    let old_brush = if let Some(brush) = fill_brush { SelectObject(dc, brush) }
        else { SelectObject(dc, GetStockObject(NULL_BRUSH)) };
    let border_w = sh.border_width.unwrap_or(0.0);
    let border_pen = if border_w > 0.0 {
        let c = sh.border_color.as_deref().and_then(parse_hex_rgb).unwrap_or((0, 0, 0));
        Some(CreatePen(PS_SOLID, (border_w as f64 * scale).round().max(1.0) as i32,
            COLORREF(colorref(c.0, c.1, c.2))))
    } else { None };
    let old_pen = if let Some(pen) = border_pen { SelectObject(dc, pen) }
        else { SelectObject(dc, GetStockObject(NULL_PEN)) };

    if !emit_shape_path(dc, sh, scale) {
        SelectObject(dc, old_pen);
        SelectObject(dc, old_brush);
        if let Some(pen) = border_pen { let _ = DeleteObject(pen); }
        if let Some(brush) = fill_brush { let _ = DeleteObject(brush); }
        return false;
    }
    let _ = StrokeAndFillPath(dc);

    SelectObject(dc, old_pen); SelectObject(dc, old_brush);
    if let Some(pen) = border_pen { let _ = DeleteObject(pen); }
    if let Some(brush) = fill_brush { let _ = DeleteObject(brush); }
    true
}

/// Draw an `a:custGeom` outline as its real path instead of the bounding box.
///
/// Every deck in the dev corpus uses custGeom (11470 shapes on 628 of 886
/// slides) and each one was previously painted as an axis-aligned slab of its
/// fill colour. The path is built in the path's declared local space, mapped
/// onto the shape box, then flipped and rotated about the box centre exactly
/// like the preset path.
///
/// Returns false when the shape must keep the legacy rectangle: no custGeom, a
/// translucent fill (the solid-brush path cannot composite along a bezier --
/// same rule as the presets), or a path whose local space is degenerate.
#[cfg(windows)]
unsafe fn draw_custom_geometry_gdi(
    dc: windows::Win32::Graphics::Gdi::HDC,
    sh: &Shape,
    scale: f64,
) -> bool {
    use windows::Win32::Foundation::COLORREF;
    use windows::Win32::Graphics::Gdi::*;

    if sh.custom_geometry.is_some() && drawable_geometry(sh).is_none() && custgeom_on() {
        // A geometry that exists but is not drawable (unsupported command,
        // degenerate box, no segments) keeps the rectangular fallback.
        return false;
    }
    let geom = match drawable_geometry(sh) {
        Some(g) => g,
        None => return false,
    };
    // A custGeom shape whose fill is a GRADIENT has no solid brush to paint
    // with: this would trace the path with no brush, paint nothing, and then
    // report success -- which stops `paint_shape_gradient`, the one painter
    // that clips a ramp to this very geometry, from ever running. d24's three
    // full-height layout bands are exactly that shape, and the deck rendered
    // bare background where PowerPoint draws half the slide.
    //
    // S-GRADSTROKE (2026-08-24): the `border_width <= 0` half of this test is
    // obsolete. It was written when handing the shape over lost its outline;
    // since S-GEOMALPHA and S-PRESETSTROKE the border pass below strokes the
    // real geometry (custGeom path, else `emit_shape_path`), so a shape can
    // have BOTH its ramp and its outline. d16 slide 13's iceberg is 28 custGeom
    // triangles filled white -> accent6 and outlined 0.75pt in `lt1`: the
    // outline kept the path here, the path had no brush, and the iceberg was
    // drawn as 28 white lines on a white page.
    if geomgrad_on()
        && sh.fill_color.is_none()
        && sh.gradient.is_some()
        && (gradstroke_on() || sh.border_width.unwrap_or(0.0) <= 0.0)
    {
        return false;
    }
    if sh.fill_color.is_some()
        && sh
            .fill_alpha
            .filter(|_| fill_alpha_on())
            .is_some_and(|a| a < 1.0)
    {
        return false;
    }

    let fill_brush = sh
        .fill_color
        .as_deref()
        .and_then(parse_hex_rgb)
        .map(|c| CreateSolidBrush(COLORREF(colorref(c.0, c.1, c.2))));
    let border_w = sh.border_width.unwrap_or(0.0);
    let border_pen = if border_w > 0.0 {
        let c = sh
            .border_color
            .as_deref()
            .and_then(parse_hex_rgb)
            .unwrap_or((0, 0, 0));
        Some(outline_pen(
            (border_w as f64 * scale).round().max(1.0) as i32,
            colorref(c.0, c.1, c.2),
            sh.border_dash.as_deref(),
            border_w as f64 * scale,
            None,
        ))
    } else {
        None
    };

    let mut drew = false;
    for path in &geom.paths {
        if path.commands.is_empty() {
            continue;
        }
        let brush = match (path.fill_none, fill_brush) {
            (false, Some(b)) => b,
            _ => HBRUSH(GetStockObject(NULL_BRUSH).0),
        };
        let old_brush = SelectObject(dc, brush);
        let old_pen = match border_pen {
            Some(pen) => SelectObject(dc, pen),
            None => SelectObject(dc, GetStockObject(NULL_PEN)),
        };

        // PowerPoint render-truth (custgeom_fillrule probe, 2026-08-17, 5 arms
        // exported by PowerPoint itself): a multi-subpath custGeom fills
        // EVEN-ODD, not by non-zero winding. Two nested squares wound the SAME
        // way leave a hole (C1) and a single self-intersecting pentagram is
        // hollow at the centre (C3) -- both of which non-zero winding would
        // fill -- while three nested squares fill the innermost again (C5).
        // 901 of the corpus's 11470 custGeom shapes are multi-subpath and
        // filled, so this is real ink. ALTERNATE is also GDI's default, but
        // the DC is shared with every other draw path, so set it explicitly.
        let old_fill_mode = SetPolyFillMode(dc, ALTERNATE);
        if openpath_on() {
            // `StrokeAndFillPath` CLOSES every open figure before stroking, so
            // a path that never says `a:close` gets an extra segment drawn from
            // its last point back to its first. d06's winding road (two noFill
            // shapes, an 18pt black stroke and a 1.5pt white dashed one, 58
            // points from the top-right corner to the bottom-left) came out with
            // a straight diagonal ruled across the whole slide, on ten slides of
            // that template. PowerPoint strokes only the segments the path
            // declares. Filling still closes implicitly -- that part is right,
            // and is what `FillPath` does -- so the two passes are split rather
            // than the close suppressed.
            let has_fill = !path.fill_none && fill_brush.is_some();
            if has_fill {
                let _ = BeginPath(dc);
                emit_geom_path_gdi(dc, sh, path, scale);
                let _ = EndPath(dc);
                let _ = FillPath(dc);
            }
            if border_pen.is_some() {
                let _ = BeginPath(dc);
                emit_geom_path_gdi(dc, sh, path, scale);
                let _ = EndPath(dc);
                let _ = StrokePath(dc);
            }
        } else {
            let _ = BeginPath(dc);
            emit_geom_path_gdi(dc, sh, path, scale);
            let _ = EndPath(dc);
            let _ = StrokeAndFillPath(dc);
        }
        drew = true;

        SetPolyFillMode(dc, CREATE_POLYGON_RGN_MODE(old_fill_mode));
        SelectObject(dc, old_pen);
        SelectObject(dc, old_brush);
    }

    if let Some(pen) = border_pen {
        let _ = DeleteObject(pen);
    }
    if let Some(brush) = fill_brush {
        let _ = DeleteObject(brush);
    }
    drew
}

/// Emit ONE `a:path` into the DC's current path, WITHOUT painting it.
///
/// Shared by the outline draw and the fill clip so the two can never disagree
/// about where the shape's boundary is.
#[cfg(windows)]
unsafe fn emit_geom_path_gdi(
    dc: windows::Win32::Graphics::Gdi::HDC,
    sh: &Shape,
    path: &oxislides_core::ir::GeomPath,
    scale: f64,
) {
    use windows::Win32::Foundation::POINT;
    use windows::Win32::Graphics::Gdi::*;

    let (w, h) = (sh.width.max(0.0), sh.height.max(0.0));
    // @w / @h are the path's own units. Both are declared on every corpus
    // path; the schema's 0 means "already in the shape's EMU space".
    let (sx, sy) = (
        if path.w > 0.0 { w / path.w } else { 1.0 / 12700.0 },
        if path.h > 0.0 { h / path.h } else { 1.0 / 12700.0 },
    );
    let map = |px: f32, py: f32| -> POINT {
        let (mut lx, mut ly) = (px * sx, py * sy);
        if sh.flip_h {
            lx = w - lx;
        }
        if sh.flip_v {
            ly = h - ly;
        }
        let (dx, dy) = (lx - w / 2.0, ly - h / 2.0);
        let (sn, cs) = sh.rotation.to_radians().sin_cos();
        POINT {
            x: (((sh.x + w / 2.0 + dx * cs - dy * sn) as f64) * scale).round() as i32,
            y: (((sh.y + h / 2.0 + dx * sn + dy * cs) as f64) * scale).round() as i32,
        }
    };
    let mut open = false;
    for cmd in &path.commands {
        match cmd {
            GeomCmd::MoveTo(x, y) => {
                let p = map(*x, *y);
                let _ = MoveToEx(dc, p.x, p.y, None);
                open = true;
            }
            GeomCmd::LineTo(x, y) => {
                if open {
                    let p = map(*x, *y);
                    let _ = LineTo(dc, p.x, p.y);
                }
            }
            GeomCmd::CubicTo(x1, y1, x2, y2, x3, y3) => {
                if open {
                    let _ = PolyBezierTo(dc, &[map(*x1, *y1), map(*x2, *y2), map(*x3, *y3)]);
                }
            }
            GeomCmd::Close => {
                if open {
                    let _ = CloseFigure(dc);
                    open = false;
                }
            }
        }
    }
}

/// The shape's drawable geometry, or None when it must keep the box fallback.
fn drawable_geometry(sh: &Shape) -> Option<&oxislides_core::ir::CustomGeometry> {
    match sh.custom_geometry.as_ref() {
        Some(g)
            if !g.unsupported
                && custgeom_on()
                && sh.width > 0.0
                && sh.height > 0.0
                && g.paths.iter().any(|p| !p.commands.is_empty()) =>
        {
            Some(g)
        }
        _ => None,
    }
}

/// Clip the DC to the shape's outline for the duration of an image blit.
///
/// PowerPoint render-truth (`custgeom_blipfill` probe, 2026-08-17, exported by
/// PowerPoint itself): a shape's `a:blipFill` is CLIPPED to the shape's own
/// geometry, and `a:fillRect` insets/expands the DESTINATION inside that clip.
///   - D2, a triangle path with fillRect 0: the box corners outside the
///     triangle are BLANK, the interior carries the source's lower quadrants.
///   - D3, a rect path with `r="-100000"`: the source's left half is stretched
///     across the whole box (destination twice as wide) and the half that
///     reaches past the box is GONE, not painted beside it.
///   - D6, d28's literal `r="-145344" b="-574764"`: every interior sample is
///     the source's top-left quadrant and nothing spills onto the page.
/// Oxi already modelled fillRect; the missing clip is what let d28's bunting
/// images (2.45 x 6.75 oversized) cover the page and put that deck at the
/// corpus floor (0.5217).
///
/// 2141 of the corpus's shape-level blipFills are on custGeom shapes -- all of
/// them -- so the geometry needed for the clip is exactly what is available.
#[cfg(windows)]
unsafe fn clip_to_geometry_gdi(
    dc: windows::Win32::Graphics::Gdi::HDC,
    sh: &Shape,
    scale: f64,
) -> bool {
    use windows::Win32::Graphics::Gdi::*;

    let geom = match drawable_geometry(sh) {
        Some(g) if blipclip_on() => g,
        // No custGeom, but a PRESET outline clips the raster just the same.
        // d11 slide 33's four portraits are `p:pic` carrying
        // `<a:prstGeom prst="ellipse"/>`: circles in PowerPoint, squares here.
        // 91 pictures across 16 dev decks state a non-rectangular geometry, and
        // 78 of them are a preset rather than a custGeom.
        _ if picclip_on() && emit_shape_path(dc, sh, scale) => {
            let ok = SelectClipPath(dc, RGN_COPY).as_bool();
            if !ok {
                let _ = SelectClipRgn(dc, None);
            }
            return ok;
        }
        _ => return false,
    };
    let _ = BeginPath(dc);
    for path in &geom.paths {
        emit_geom_path_gdi(dc, sh, path, scale);
    }
    let _ = EndPath(dc);
    // Even-odd, the rule measured for custGeom fills, applies to the region a
    // path becomes as well.
    let old_mode = SetPolyFillMode(dc, ALTERNATE);
    let ok = SelectClipPath(dc, RGN_COPY).as_bool();
    SetPolyFillMode(dc, CREATE_POLYGON_RGN_MODE(old_mode));
    if !ok {
        let _ = SelectClipRgn(dc, None);
    }
    ok
}

/// A picture is flipped and rotated with its shape unless this is set.
fn imgrot_on() -> bool {
    std::env::var("OXI_IMGROT_DISABLE").is_err()
}

/// Feeds one embedded font part to `TTLoadEmbeddedFont`'s pull-style reader.
#[cfg(windows)]
struct FontStream {
    data: Vec<u8>,
    pos: usize,
}

/// `READEMBEDPROC`: copy the next `count` bytes into t2embed's buffer.
#[cfg(windows)]
unsafe extern "system" fn read_embedded_font(
    stream: *mut core::ffi::c_void,
    dest: *mut core::ffi::c_void,
    count: u32,
) -> u32 {
    if stream.is_null() || dest.is_null() {
        return 0;
    }
    let s = &mut *(stream as *mut FontStream);
    let n = (count as usize).min(s.data.len().saturating_sub(s.pos));
    if n > 0 {
        std::ptr::copy_nonoverlapping(s.data.as_ptr().add(s.pos), dest as *mut u8, n);
        s.pos += n;
    }
    n as u32
}

#[cfg(windows)]
thread_local! {
    /// The style-suffixed names `install_embedded_fonts` actually registered.
    static EMBEDDED_FACES: std::cell::RefCell<std::collections::HashSet<String>> =
        std::cell::RefCell::new(std::collections::HashSet::new());
}

/// Embedded fonts are given one GDI family PER STYLE unless this is set.
fn embedstyle_on() -> bool {
    std::env::var("OXI_EMBEDSTYLE_DISABLE").is_err()
}

/// The GDI family name one part of an embedded typeface is registered under.
///
/// All four parts of a `p:embeddedFont` used to be renamed to the same
/// `p:font/@typeface`, and GDI then cannot tell them apart. Measured on d24
/// (2026-08-20), asking for weight 400 upright:
///
/// | family            | GDI served |
/// |-------------------|------------|
/// | Fira Sans         | tmItalic=255 |
/// | Fira Sans Light   | upright |
/// | Fira Sans Medium  | tmItalic=255 |
/// | Fira Sans SemiBold| tmItalic=255 |
/// | Montserrat        | upright |
///
/// Three of five families answered a plain request with an ITALIC face, which
/// is why d24's title came out slanted once it finally found its typeface.
/// Loading only the regular part makes it coherent (tmWeight=500 upright, the
/// SemiBold's real weight) but throws away the three real faces and leaves GDI
/// synthesising them, so instead each part gets its own family name and the
/// draw side asks for the exact one.
///
/// LF_FACESIZE caps a family at 31 characters; a name that would not fit keeps
/// the old shared name, which is no worse than before.
fn embedded_face_name(typeface: &str, bold: bool, italic: bool) -> String {
    // Only the ITALIC parts move out. Giving the BOLD part its own family too
    // measured -0.00126 over the corpus (5 improved / 11 regressed, d01 -0.0288):
    // PowerPoint's bold on those decks is WIDER than the embedded bold part, so
    // d01's "Add description here" stopped wrapping to the two lines
    // PowerPoint needs. Whatever GDI was serving for a bold request already
    // matched it better, and this change has no business moving it. Taking the
    // italic parts out of the shared family is enough for the observed harm --
    // a plain request can then only land on the regular or the bold face.
    if !italic || !embedstyle_on() {
        return typeface.to_string();
    }
    let suffix = if bold { " #BI" } else { " #I" };
    if typeface.chars().count() + suffix.chars().count() > 31 {
        return typeface.to_string();
    }
    format!("{typeface}{suffix}")
}

/// The face to actually create for a (family, bold, italic) request, and the
/// weight and slant to ask GDI for. An embedded part carries its own style, so
/// it is requested at weight 400 upright; anything else keeps the request.
#[cfg(windows)]
fn styled_face(family: &str, bold: bool, italic: bool) -> (String, i32, bool) {
    if embedstyle_on() && italic {
        let name = embedded_face_name(family, bold, italic);
        if name != family && EMBEDDED_FACES.with(|f| f.borrow().contains(&name)) {
            // The part carries its own slant; the weight is still asked for,
            // since a family may embed only one italic part.
            return (name, if bold { 700 } else { 400 }, false);
        }
    }
    (
        family.to_string(),
        if bold { 700 } else { 400 },
        italic,
    )
}

/// Install the deck's embedded fonts so GDI can resolve them by name.
///
/// The `.fntdata` parts are EOT (all 262 in the dev corpus are EOT 2.2 with
/// MicroType Express compression), so they cannot be handed to
/// `AddFontMemResourceEx`; `TTLoadEmbeddedFont` is the API that decompresses
/// them, and it is the route PowerPoint itself takes.
///
/// Measured 2026-08-17 on d28's Calistoga: before the call, `CreateFont`
/// ("Calistoga") silently yields MS PGothic and "Abraham Lincoln" measures
/// 417x60 at 60px; after it, `GetTextFace` reports Calistoga and the same
/// string measures 473x102. A privately loaded font does NOT appear in
/// `EnumFontFamiliesEx`, which is expected and does not affect `CreateFont`.
///
/// ★TRAP: `ulPrivs` is a SINGLE license value, not a bitmask -- passing
/// `LICENSE_PREVIEWPRINT | LICENSE_EDITABLE` returns E_EXCEPTION (0x105) with
/// the read callback never invoked, which reads exactly like "this API does
/// not work here". LICENSE_INSTALLABLE (0) loads every corpus part.
///
/// ★Each face is RENAMED to the `p:font/@typeface` the deck's runs ask for,
/// because a part's own family name often is NOT that name. d04 ships
/// `RobotoSlab-regular.fntdata` whose internal family is "Roboto Slab Light"
/// (weight 300) under `typeface="Roboto Slab"`; loading it unrenamed leaves
/// "Roboto Slab" resolvable only through the BOLD part, so GDI serves weight
/// 400 from the 700 face and the whole deck renders bold. With the rename,
/// weight 400 selects tmWeight=300 and weight 700 the real bold face
/// (219px against the 230px GDI synthesises) -- measured 2026-08-17.
///
/// The handles are deliberately leaked: the fonts must stay loaded for the
/// whole run, and the process exits right after rendering.
#[cfg(windows)]
fn install_embedded_fonts(pres: &Presentation) -> usize {
    use windows::Win32::Foundation::HANDLE;
    use windows::Win32::Graphics::Gdi::{
        TTLoadEmbeddedFont, EMBEDDED_FONT_PRIV_STATUS, FONT_LICENSE_PRIVS,
        TTLOAD_EMBEDDED_FONT_STATUS,
    };

    if std::env::var("OXI_EMBEDFONT_DISABLE").is_ok() {
        return 0;
    }
    const TTLOAD_PRIVATE: u32 = 0x0000_0001;
    const LICENSE_INSTALLABLE: u32 = 0x0000_0000;
    let mut loaded = 0;
    for font in &pres.embedded_fonts {
        let mut stream = Box::new(FontStream {
            data: font.data.clone(),
            pos: 0,
        });
        let mut handle = HANDLE::default();
        let mut priv_status = EMBEDDED_FONT_PRIV_STATUS::default();
        let mut status = TTLOAD_EMBEDDED_FONT_STATUS::default();
        let face = embedded_face_name(&font.typeface, font.bold, font.italic);
        let mut win_name: Vec<u16> = face.encode_utf16().collect();
        win_name.push(0);
        let rc = unsafe {
            TTLoadEmbeddedFont(
                &mut handle,
                TTLOAD_PRIVATE,
                &mut priv_status,
                FONT_LICENSE_PRIVS(LICENSE_INSTALLABLE),
                &mut status,
                Some(read_embedded_font),
                stream.as_mut() as *mut FontStream as *const core::ffi::c_void,
                windows::core::PCWSTR(win_name.as_ptr()),
                None,
                None,
            )
        };
        if rc == 0 {
            loaded += 1;
            if face != font.typeface {
                EMBEDDED_FACES.with(|f| f.borrow_mut().insert(face));
            }
            std::mem::forget(stream); // t2embed keeps no reference, but the
                                      // font must outlive this scope anyway
        } else {
            eprintln!(
                "  embedded font '{}' (bold={} italic={}) failed to load: 0x{:x}",
                font.typeface, font.bold, font.italic, rc
            );
        }
    }
    enum_faces_debug();
    if std::env::var("OXI_FD_DEBUG").is_ok() {
        // Ask each family what it is really serving. A deck that labels a part
        // with one weight's name and another weight's data shows up here as two
        // families reporting identical advances.
        let mut seen: std::collections::BTreeSet<&str> = std::collections::BTreeSet::new();
        for font in &pres.embedded_fonts {
            if seen.insert(font.typeface.as_str()) {
                for bold in [false, true] {
                    let _ = fontdata_advance_em(&font.typeface, bold, false, 'a');
                }
            }
        }
    }
    if std::env::var("OXI_DEBUG_EMBED").is_ok() {
        use std::collections::BTreeSet;
        let names: BTreeSet<&str> = pres
            .embedded_fonts
            .iter()
            .map(|f| f.typeface.as_str())
            .collect();
        for name in names {
            for (w, it) in [(400, false), (400, true), (700, false), (700, true)] {
                debug_face(name, w, it);
            }
            // ...and the renamed parts the italic path actually asks for, which
            // the bare-family probe above can never reach.
            for suffix in [" #I", " #BI"] {
                let renamed = format!("{name}{suffix}");
                let known = EMBEDDED_FACES.with(|f| f.borrow().contains(&renamed));
                eprintln!("EMBED   registered({renamed:?}) = {known}");
                if known {
                    debug_face(&renamed, 400, false);
                }
            }
        }
    }
    loaded
}

/// The Office cloud-font cache -- the third place a font can live, after the
/// system fonts and the deck's own embedded parts.
///
/// `%LOCALAPPDATA%\Microsoft\FontCache\4\CloudFonts\<package>\<id>.ttf`. Office
/// downloads these ON DEMAND, so the set grows over time, and GDI never sees any
/// of them: `CreateFontW("IBM Plex Sans")` silently serves a substitute. On the
/// dev corpus that is 2330 runs on d06, 2452 on d16, 2408 on d24, 2368 on d19
/// and 2439 on d35 -- whole decks drawn in the wrong face while PowerPoint's own
/// export names the real one.
///
/// ★The DIRECTORY IS NOT THE FAMILY: `CloudFonts\IBM Plex Sans\` holds both
/// `IBM Plex Sans` and `IBM Plex Sans Condensed`. It names the download package.
/// The family has to come out of each file's own name table.
#[cfg(windows)]
fn cloud_font_root() -> Option<std::path::PathBuf> {
    let local = std::env::var_os("LOCALAPPDATA")?;
    let root = std::path::Path::new(&local)
        .join("Microsoft")
        .join("FontCache")
        .join("4")
        .join("CloudFonts");
    root.is_dir().then_some(root)
}

/// The typographic family of an sfnt blob: name ID 16, else name ID 1.
///
/// ID 16 is the one that groups the weights (`IBM Plex Sans` for both the
/// regular and the bold file); ID 1 splits them on faces that have no ID 16.
#[cfg(windows)]
fn sfnt_family(data: &[u8]) -> Option<String> {
    const SFNT: [[u8; 4]; 3] = [[0x00, 0x01, 0x00, 0x00], *b"OTTO", *b"true"];
    let magic = data.get(0..4)?;
    if !SFNT.iter().any(|tag| tag.as_slice() == magic) {
        return None;
    }
    let tables = u16::from_be_bytes([data[4], data[5]]) as usize;
    let mut name_off = None;
    for index in 0..tables {
        let rec = 12 + 16 * index;
        if data.get(rec..rec + 4)? == b"name" {
            name_off = Some(u32::from_be_bytes([
                data[rec + 8],
                data[rec + 9],
                data[rec + 10],
                data[rec + 11],
            ]) as usize);
            break;
        }
    }
    let base = name_off?;
    let count = u16::from_be_bytes([*data.get(base + 2)?, *data.get(base + 3)?]) as usize;
    let strings =
        base + u16::from_be_bytes([*data.get(base + 4)?, *data.get(base + 5)?]) as usize;
    let mut fallback: Option<String> = None;
    for index in 0..count {
        let rec = base + 6 + 12 * index;
        let field = |at: usize| -> Option<u16> {
            Some(u16::from_be_bytes([
                *data.get(rec + at)?,
                *data.get(rec + at + 1)?,
            ]))
        };
        let platform = field(0)?;
        let name_id = field(6)?;
        if name_id != 1 && name_id != 16 {
            continue;
        }
        let len = field(8)? as usize;
        let off = field(10)? as usize;
        let raw = data.get(strings + off..strings + off + len)?;
        let text = if platform == 3 {
            let units: Vec<u16> = raw
                .chunks_exact(2)
                .map(|pair| u16::from_be_bytes([pair[0], pair[1]]))
                .collect();
            String::from_utf16(&units).ok()?
        } else {
            raw.iter().map(|byte| *byte as char).collect()
        };
        if name_id == 16 {
            return Some(text);
        }
        fallback.get_or_insert(text);
    }
    fallback
}

/// What GDI hands back when asked for this family by name.
#[cfg(windows)]
fn gdi_face_for(family: &str) -> String {
    use windows::Win32::Graphics::Gdi::*;

    let probe = probe_dc();
    let wide: Vec<u16> = family.encode_utf16().chain(std::iter::once(0)).collect();
    unsafe {
        let font = CreateFontW(
            -64,
            0,
            0,
            0,
            400,
            0,
            0,
            0,
            DEFAULT_CHARSET.0 as u32,
            OUT_DEFAULT_PRECIS.0 as u32,
            CLIP_DEFAULT_PRECIS.0 as u32,
            CLEARTYPE_QUALITY.0 as u32,
            (DEFAULT_PITCH.0 | FF_DONTCARE.0) as u32,
            windows::core::PCWSTR(wide.as_ptr()),
        );
        if font.is_invalid() {
            return String::new();
        }
        let old = SelectObject(probe, font);
        let mut name = [0u16; 64];
        let got = GetTextFaceW(probe, Some(&mut name));
        SelectObject(probe, old);
        let _ = DeleteObject(font);
        if got > 0 {
            String::from_utf16_lossy(&name[..(got as usize).saturating_sub(1)])
        } else {
            String::new()
        }
    }
}

/// Register the cloud-cache files for the families GDI cannot already serve.
///
/// Two passes on purpose: the first decides which families are missing, the
/// second registers every FILE of those families. Deciding per file would
/// register the regular, make the family resolvable, and then skip the bold --
/// the same defect the embedded-font rename exists to avoid (d04's Roboto Slab,
/// 2026-08-17).
///
/// Called AFTER `install_embedded_fonts` so a deck's own part always wins: by
/// then GDI serves that family and the cache file for it is skipped.
///
/// `FR_PRIVATE` keeps the registration inside this process -- nothing is
/// installed on the machine and it dies with the process.
#[cfg(windows)]
fn install_cloud_fonts() -> usize {
    use std::collections::BTreeSet;
    use std::os::windows::ffi::OsStrExt;
    use windows::Win32::Graphics::Gdi::{AddFontResourceExW, FR_PRIVATE};

    if std::env::var("OXI_CLOUDFONT_DISABLE").is_ok() {
        return 0;
    }
    let Some(root) = cloud_font_root() else {
        return 0;
    };
    let mut files: Vec<(std::path::PathBuf, String)> = Vec::new();
    let mut stack = vec![root];
    while let Some(dir) = stack.pop() {
        let Ok(entries) = std::fs::read_dir(&dir) else {
            continue;
        };
        for entry in entries.flatten() {
            let path = entry.path();
            if path.is_dir() {
                stack.push(path);
                continue;
            }
            let extension = path
                .extension()
                .and_then(|ext| ext.to_str())
                .unwrap_or_default()
                .to_ascii_lowercase();
            if extension != "ttf" && extension != "otf" {
                continue;
            }
            let Ok(blob) = std::fs::read(&path) else {
                continue;
            };
            if let Some(family) = sfnt_family(&blob) {
                files.push((path, family));
            }
        }
    }
    files.sort();
    let mut missing: BTreeSet<String> = BTreeSet::new();
    let mut checked: BTreeSet<String> = BTreeSet::new();
    for (_path, family) in &files {
        if !checked.insert(family.clone()) {
            continue;
        }
        if !gdi_face_for(family).eq_ignore_ascii_case(family) {
            missing.insert(family.clone());
        }
    }
    let mut loaded = 0;
    for (path, family) in &files {
        if !missing.contains(family) {
            continue;
        }
        let wide: Vec<u16> = path
            .as_os_str()
            .encode_wide()
            .chain(std::iter::once(0))
            .collect();
        let added =
            unsafe { AddFontResourceExW(windows::core::PCWSTR(wide.as_ptr()), FR_PRIVATE, None) };
        if added > 0 {
            loaded += 1;
        } else {
            eprintln!("  cloud font '{}' failed to register", path.display());
        }
    }
    loaded
}

/// What GDI actually serves for one (family, weight, italic) request.
#[cfg(windows)]
fn debug_face(family: &str, weight: i32, italic: bool) {
    use windows::Win32::Graphics::Gdi::*;

    let dc = probe_dc();
    let wide: Vec<u16> = family.encode_utf16().chain(std::iter::once(0)).collect();
    unsafe {
        let font = CreateFontW(
            -64,
            0,
            0,
            0,
            weight,
            u32::from(italic),
            0,
            0,
            DEFAULT_CHARSET.0 as u32,
            OUT_DEFAULT_PRECIS.0 as u32,
            CLIP_DEFAULT_PRECIS.0 as u32,
            CLEARTYPE_QUALITY.0 as u32,
            (DEFAULT_PITCH.0 | FF_DONTCARE.0) as u32,
            windows::core::PCWSTR(wide.as_ptr()),
        );
        if font.is_invalid() {
            eprintln!("EMBED {family:24} w={weight} i={italic}: CreateFont failed");
            return;
        }
        let old = SelectObject(dc, font);
        let mut name = [0u16; 64];
        let n = GetTextFaceW(dc, Some(&mut name));
        let mut tm = TEXTMETRICW::default();
        let _ = GetTextMetricsW(dc, &mut tm);
        // A face name and a slant flag do not say whether the ADVANCES are the
        // ones the deck's own part carries, and that is the thing a wrong face
        // gets wrong. So measure a real string at a real size too: a synthesised
        // oblique keeps the upright widths, a genuine italic part does not.
        let probe: Vec<u16> = "invoke philosophical thoughts ".encode_utf16().collect();
        let mut sz = windows::Win32::Foundation::SIZE::default();
        let _ = GetTextExtentPoint32W(dc, &probe, &mut sz);
        SelectObject(dc, old);
        let _ = DeleteObject(font);
        eprintln!(
            "EMBED {family:24} w={weight} i={italic} -> face={:?} tmWeight={} tmItalic={}              width64={} ({:.2}pt at 36)",
            String::from_utf16_lossy(&name[..(n as usize).saturating_sub(1)]),
            tm.tmWeight,
            tm.tmItalic,
            sz.cx,
            f64::from(sz.cx) * 36.0 / 64.0
        );
    }
}

/// Resample a picture into the page-aligned pixel box its flipped and rotated
/// destination occupies. Returns the buffer and where to blit it, or None when
/// there is nothing to transform.
///
/// PowerPoint render-truth (`img_rotation` probe, 2026-08-17, exported by
/// PowerPoint itself; the source is a 2x2 colour grid so each sample names the
/// source corner that landed there):
///   - E2 `p:pic` @rot=90: box TL <- source BL, TR <- TL, BL <- BR, BR <- TR,
///     i.e. the raster turns CLOCKWISE with the shape, about the box centre --
///     the same convention already derived for outlines.
///   - E4 a shape blipFill with rotWithShape="1" gives the identical map;
///     E5 with "0" leaves the raster upright.
///   - E6 @rot=90 + flipH: box TL <- BR, TR <- TR, BL <- BL, BR <- TL, which
///     is FLIP FIRST in the shape's local box, THEN rotate -- the composition
///     `emit_geom_path_gdi` already uses for the outline.
/// The corpus has 489 rotated shape blipFills, 9 rotated pictures and 257
/// flipped image shapes; all of it was painted axis-aligned before this.
fn transform_picture(
    rgba: &image::RgbaImage,
    src: (i32, i32, i32, i32),
    dest: (f64, f64, f64, f64),
    centre: (f64, f64),
    angle_deg: f64,
    flip_h: bool,
    flip_v: bool,
) -> Option<(image::RgbaImage, i32, i32)> {
    let (sx0, sy0, sw, sh) = src;
    let (dx, dy, dw, dh) = dest;
    if sw <= 0 || sh <= 0 || dw <= 0.0 || dh <= 0.0 {
        return None;
    }
    let (sn, cs) = angle_deg.to_radians().sin_cos();
    let rotate = |px: f64, py: f64| {
        let (rx, ry) = (px - centre.0, py - centre.1);
        (centre.0 + rx * cs - ry * sn, centre.1 + rx * sn + ry * cs)
    };
    let corners = [
        rotate(dx, dy),
        rotate(dx + dw, dy),
        rotate(dx + dw, dy + dh),
        rotate(dx, dy + dh),
    ];
    let min_x = corners.iter().map(|p| p.0).fold(f64::INFINITY, f64::min).floor() as i32;
    let min_y = corners.iter().map(|p| p.1).fold(f64::INFINITY, f64::min).floor() as i32;
    let max_x = corners.iter().map(|p| p.0).fold(f64::NEG_INFINITY, f64::max).ceil() as i32;
    let max_y = corners.iter().map(|p| p.1).fold(f64::NEG_INFINITY, f64::max).ceil() as i32;
    let (out_w, out_h) = ((max_x - min_x) as i64, (max_y - min_y) as i64);
    // A negative fillRect can blow the destination up to many times the page;
    // refuse rather than allocate unboundedly (the caller keeps its old path).
    if out_w <= 0 || out_h <= 0 || out_w * out_h > 64_000_000 {
        return None;
    }
    let mut out = image::RgbaImage::new(out_w as u32, out_h as u32);
    let (iw, ih) = (rgba.width() as i32, rgba.height() as i32);
    for oy in 0..out_h {
        for ox in 0..out_w {
            let px = min_x as f64 + ox as f64 + 0.5;
            let py = min_y as f64 + oy as f64 + 0.5;
            // Inverse rotation back into the destination rect's own frame.
            let (rx, ry) = (px - centre.0, py - centre.1);
            let ux = centre.0 + rx * cs + ry * sn;
            let uy = centre.1 - rx * sn + ry * cs;
            let mut fx = (ux - dx) / dw;
            let mut fy = (uy - dy) / dh;
            if !(0.0..1.0).contains(&fx) || !(0.0..1.0).contains(&fy) {
                continue; // outside the picture: stays fully transparent
            }
            if flip_h {
                fx = 1.0 - fx;
            }
            if flip_v {
                fy = 1.0 - fy;
            }
            // Bilinear sample of the srcRect-cropped source.
            let sxf = sx0 as f64 + fx * sw as f64 - 0.5;
            let syf = sy0 as f64 + fy * sh as f64 - 0.5;
            let (x0, y0) = (sxf.floor() as i32, syf.floor() as i32);
            let (tx, ty) = (sxf - x0 as f64, syf - y0 as f64);
            let mut acc = [0.0f64; 4];
            for (i, (ox2, oy2)) in [(0, 0), (1, 0), (0, 1), (1, 1)].into_iter().enumerate() {
                let sx = (x0 + ox2).clamp(0, iw - 1);
                let sy = (y0 + oy2).clamp(0, ih - 1);
                let w = match i {
                    0 => (1.0 - tx) * (1.0 - ty),
                    1 => tx * (1.0 - ty),
                    2 => (1.0 - tx) * ty,
                    _ => tx * ty,
                };
                let p = rgba.get_pixel(sx as u32, sy as u32);
                for c in 0..4 {
                    acc[c] += w * p[c] as f64;
                }
            }
            out.put_pixel(
                ox as u32,
                oy as u32,
                image::Rgba([
                    acc[0].round().clamp(0.0, 255.0) as u8,
                    acc[1].round().clamp(0.0, 255.0) as u8,
                    acc[2].round().clamp(0.0, 255.0) as u8,
                    acc[3].round().clamp(0.0, 255.0) as u8,
                ]),
            );
        }
    }
    Some((out, min_x, min_y))
}

/// The pre-S-TBLCELL table rendering, kept so `OXI_TBLCELL_DISABLE` reproduces
/// the shipped output exactly and the A/B is a real before-vs-after.
#[cfg(windows)]
unsafe fn draw_table_legacy(
    mem_dc: windows::Win32::Graphics::Gdi::HDC,
    pres: &Presentation,
    sh: &Shape,
    table: &oxislides_core::ir::Table,
    x: i32,
    y: i32,
    scale: f64,
) {
    use windows::Win32::Foundation::COLORREF;
    use windows::Win32::Graphics::Gdi::*;

    let pen = CreatePen(PS_SOLID, (1.0 * scale).round() as i32, COLORREF(colorref(0, 0, 0)));
    let old_pen = SelectObject(mem_dc, pen);
    let _ = SelectObject(mem_dc, GetStockObject(NULL_BRUSH));
    let mut cy = y;
    let mut cy_pt = 0.0f64;
    for (r, row) in table.rows.iter().enumerate() {
        let h_pt = table.row_heights.get(r).copied().unwrap_or(0.0) as f64;
        let ph = if tbledge_on() {
            cy = y + (cy_pt * scale).round() as i32;
            y + ((cy_pt + h_pt) * scale).round() as i32 - cy
        } else {
            (h_pt * scale).round() as i32
        };
        let mut cx = x;
        let mut cx_pt = 0.0f64;
        for (c, cell) in row.iter().enumerate() {
            let w_pt = table.col_widths.get(c).copied().unwrap_or(0.0) as f64;
            let pw = if tbledge_on() {
                cx = x + (cx_pt * scale).round() as i32;
                x + ((cx_pt + w_pt) * scale).round() as i32 - cx
            } else {
                (w_pt * scale).round() as i32
            };
            let _ = Rectangle(mem_dc, cx, cy, cx + pw, cy + ph);
            let mut cursor_y = cy + (0.06 * scale).round() as i32;
            let mut prev_fs: Option<f32> = None;
            for p in &cell.paragraphs {
                // The legacy cell size is max(runs, 18pt) -- an 18pt floor the
                // rest of the engine does not have. Keep it for paragraphs that
                // carry text; only the EMPTY ones take the paragraph-mark rule.
                let legacy = p.runs.iter().filter_map(|r| r.font_size).fold(18.0, f32::max);
                let fs = if emptypara_on() && p.runs.iter().all(|r| r.text.is_empty()) {
                    p.end_para_size.or(prev_fs).unwrap_or(legacy)
                } else {
                    legacy
                };
                prev_fs = Some(fs);
                let text: String = p.runs.iter().map(|r| r.text.as_str()).collect();
                if text.trim().is_empty() {
                    cursor_y += (fs as f64 * scale * 1.2).round() as i32;
                    continue;
                }
                let family = p
                    .runs
                    .iter()
                    .find_map(|r| r.font_family.clone())
                    .unwrap_or_else(|| resolve_font(pres, sh));
                let color = p.runs.iter().find_map(|r| r.color.clone());
                draw_text_line(
                    mem_dc,
                    cx + (0.06 * scale).round() as i32,
                    cursor_y,
                    &text,
                    fs,
                    &family,
                    color.as_deref(),
                    scale,
                );
                cursor_y += (fs as f64 * scale * 1.2).round() as i32;
            }
            cx += pw;
            cx_pt += w_pt;
        }
        cy += ph;
        cy_pt += h_pt;
    }
    SelectObject(mem_dc, old_pen);
    let _ = DeleteObject(pen);
}

fn srgb_to_linear(c: f64) -> f64 {
    let c = c / 255.0;
    if c <= 0.04045 {
        c / 12.92
    } else {
        ((c + 0.055) / 1.055).powf(2.4)
    }
}

fn linear_to_srgb(c: f64) -> f64 {
    let c = c.clamp(0.0, 1.0);
    let s = if c <= 0.003_130_8 {
        c * 12.92
    } else {
        1.055 * c.powf(1.0 / 2.4) - 0.055
    };
    (s * 255.0).round().clamp(0.0, 255.0)
}

/// Colour of a gradient ramp at `t` (0..1).
///
/// Stops are interpolated in LINEAR RGB, not sRGB. Measured against three
/// PowerPoint ramps -- the probe's RED->BLUE, d04's FFD966->FF9900 and d15's
/// 572D7E->0F0B19 -- linear-RGB lerp beats plain sRGB lerp on every one of
/// them (d04 max error 10.8 vs 24.2, d15 15.8 vs 20.0). A 10-16/255 residual
/// remains: PowerPoint's ramp is not a plain lerp in any single space (d15
/// favours cos-squared, the probe favours 1-t^2), so this is the measured best
/// of the simple models rather than an exact match.
/// Linear interpolation of the ramp's per-stop alpha at `t` (1.0 = opaque).
fn gradient_alpha_at(g: &SlideGradient, t: f64) -> f64 {
    let stops = &g.stops;
    if stops.is_empty() {
        return 1.0;
    }
    let t = t.clamp(0.0, 1.0);
    let last = stops.len() - 1;
    if t <= stops[0].pos as f64 {
        return stops[0].alpha as f64;
    }
    if t >= stops[last].pos as f64 {
        return stops[last].alpha as f64;
    }
    for i in 0..last {
        let (a, b) = (&stops[i], &stops[i + 1]);
        if t >= a.pos as f64 && t <= b.pos as f64 {
            let span = (b.pos - a.pos) as f64;
            let u = if span.abs() < 1e-9 {
                0.0
            } else {
                (t - a.pos as f64) / span
            };
            return a.alpha as f64 + (b.alpha as f64 - a.alpha as f64) * u;
        }
    }
    1.0
}

/// True when any stop is translucent, i.e. the ramp must be composited rather
/// than painted as opaque bands.
fn gradient_has_alpha(g: &SlideGradient) -> bool {
    g.stops.iter().any(|s| s.alpha < 1.0)
}

fn gradient_color_at(g: &SlideGradient, t: f64) -> (u8, u8, u8) {
    let stops = &g.stops;
    if stops.is_empty() {
        return (255, 255, 255);
    }
    let parse = |s: &str| parse_hex_rgb(s).unwrap_or((0, 0, 0));
    let t = t.clamp(0.0, 1.0);
    let last = stops.len() - 1;
    if t <= stops[0].pos as f64 {
        return parse(&stops[0].color);
    }
    if t >= stops[last].pos as f64 {
        return parse(&stops[last].color);
    }
    // PowerPoint interpolates the two stop counts DIFFERENTLY, and its own PDF
    // says so: a two-stop ramp is exported as a Size-256 sampled function whose
    // midpoint is (186,0,187) for red->blue -- neither sRGB-linear (128,0,128)
    // nor a plain linear-RGB blend -- while a three-stop ramp is exported as
    // exact `FunctionType 2, N 1` segments over the raw sRGB values, and its
    // pixels match sRGB-linear to within 1.  Measured over all 256 samples of
    // every probe arm: two stops fit linear-RGB with a smoothstep ease
    // (max 20 / avg 6.1, versus 37/23.9 for a plain linear-RGB blend and
    // 59/43.2 for sRGB), three stops fit sRGB-linear exactly.  Both shapes are
    // in the corpus (58 two-stop and 17 three-stop slides), so both are honoured.
    let two_stop = stops.len() == 2;
    for i in 0..last {
        let (p0, p1) = (stops[i].pos as f64, stops[i + 1].pos as f64);
        if t >= p0 && t <= p1 {
            let f = if (p1 - p0).abs() < 1e-9 {
                0.0
            } else {
                (t - p0) / (p1 - p0)
            };
            let a = parse(&stops[i].color);
            let b = parse(&stops[i + 1].color);
            if two_stop {
                let e = f * f * (3.0 - 2.0 * f);
                let mix = |x: u8, y: u8| {
                    linear_to_srgb(
                        srgb_to_linear(x as f64) * (1.0 - e) + srgb_to_linear(y as f64) * e,
                    ) as u8
                };
                return (mix(a.0, b.0), mix(a.1, b.1), mix(a.2, b.2));
            }
            let mix = |x: u8, y: u8| (x as f64 * (1.0 - f) + y as f64 * f).round() as u8;
            return (mix(a.0, b.0), mix(a.1, b.1), mix(a.2, b.2));
        }
    }
    parse(&stops[last].color)
}

/// Paint a slide background picture. Returns false when the image could not be
/// decoded, so the caller can fall back to the gradient/flat fill -- the
/// compatible bitmap is UNINITIALISED, so something must always paint it.
///
/// PowerPoint render-truth (dev corpus, 2026-08): the exported PDF places the
/// background image at exactly the page rect on every deck that has one -- d04
/// slide10, d06 slide10/11, d16 slide1, d19 slide10 all give
/// `Rect(0, ~0, 720, 405)` for a 1280x720 source -- and none of them carries a
/// soft mask. So the fill is a plain full-page stretch at full opacity, which
/// is what every one of the corpus's 22 background fills asks for anyway:
/// they are all `<a:stretch><a:fillRect/></a:stretch>` with a bare
/// `<a:alphaModFix/>` (no `amt`), i.e. no crop, no insets, no alpha, no tile.
/// Those variants are therefore NOT modelled -- nothing measured to model from.
#[cfg(windows)]
unsafe fn paint_bg_image(
    dc: windows::Win32::Graphics::Gdi::HDC,
    w: i32,
    h: i32,
    img: &SlideBackgroundImage,
) -> bool {
    use windows::Win32::Graphics::Gdi::*;

    if w <= 0 || h <= 0 {
        return false;
    }
    let dyn_img = match image::load_from_memory(&img.data) {
        Ok(d) => d,
        Err(_) => return false,
    };
    let rgba = dyn_img.to_rgba8();
    let (iw, ih) = (rgba.width() as i32, rgba.height() as i32);
    if iw <= 0 || ih <= 0 {
        return false;
    }
    let mut bgra = Vec::with_capacity((iw * ih * 4) as usize);
    for px in rgba.pixels() {
        bgra.push(px[2]);
        bgra.push(px[1]);
        bgra.push(px[0]);
        bgra.push(px[3]);
    }
    let bmi = BITMAPINFO {
        bmiHeader: BITMAPINFOHEADER {
            biSize: std::mem::size_of::<BITMAPINFOHEADER>() as u32,
            biWidth: iw,
            biHeight: -ih, // top-down
            biPlanes: 1,
            biBitCount: 32,
            biCompression: 0, // BI_RGB
            ..Default::default()
        },
        ..Default::default()
    };
    let _ = StretchDIBits(
        dc,
        0,
        0,
        w,
        h,
        0,
        0,
        iw,
        ih,
        Some(bgra.as_ptr() as *const _),
        &bmi,
        DIB_RGB_COLORS,
        SRCCOPY,
    );
    true
}

/// Composite a picture that carries transparency.
///
/// `StretchDIBits` on a BI_RGB 32-bpp DIB throws the alpha byte away, so a PNG
/// with transparency is painted as the raw RGB stored *under* its transparent
/// pixels (usually black or white) instead of letting the page show through.
/// PowerPoint composites it. `AlphaBlend` does that, but it needs a source *DC*
/// rather than a DIB pointer, and PREMULTIPLIED source pixels, so the bitmap is
/// built as a DIB section and selected into a scratch DC.
///
/// Returns false when any GDI step fails, so the caller can fall back to the
/// opaque blit and never leave the picture unpainted.
#[cfg(windows)]
unsafe fn alpha_blit(
    dst: windows::Win32::Graphics::Gdi::HDC,
    dx: i32,
    dy: i32,
    dw: i32,
    dh: i32,
    sx0: i32,
    sy0: i32,
    sw: i32,
    sh: i32,
    iw: i32,
    ih: i32,
    rgba: &image::RgbaImage,
) -> bool {
    use windows::Win32::Graphics::Gdi::*;

    if dw <= 0 || dh <= 0 || sw <= 0 || sh <= 0 || iw <= 0 || ih <= 0 {
        return false;
    }
    let bmi = BITMAPINFO {
        bmiHeader: BITMAPINFOHEADER {
            biSize: std::mem::size_of::<BITMAPINFOHEADER>() as u32,
            biWidth: iw,
            biHeight: -ih, // top-down
            biPlanes: 1,
            biBitCount: 32,
            biCompression: 0, // BI_RGB
            ..Default::default()
        },
        ..Default::default()
    };
    let mut bits: *mut core::ffi::c_void = std::ptr::null_mut();
    let hbm = match CreateDIBSection(dst, &bmi, DIB_RGB_COLORS, &mut bits, None, 0) {
        Ok(b) if !bits.is_null() => b,
        _ => return false,
    };
    {
        // premultiplied BGRA, top-down (matches the negative biHeight)
        let n = (iw as usize) * (ih as usize);
        let out = std::slice::from_raw_parts_mut(bits as *mut u8, n * 4);
        for (i, p) in rgba.pixels().enumerate().take(n) {
            let a = p[3] as u32;
            out[i * 4] = ((p[2] as u32 * a + 127) / 255) as u8;
            out[i * 4 + 1] = ((p[1] as u32 * a + 127) / 255) as u8;
            out[i * 4 + 2] = ((p[0] as u32 * a + 127) / 255) as u8;
            out[i * 4 + 3] = p[3];
        }
    }
    let src_dc = CreateCompatibleDC(dst);
    if src_dc.0.is_null() {
        let _ = DeleteObject(hbm);
        return false;
    }
    let old = SelectObject(src_dc, hbm);
    let old_stretch = begin_smooth_blit(dst, (sw, sh), (dw, dh));
    let bf = BLENDFUNCTION {
        BlendOp: AC_SRC_OVER as u8,
        BlendFlags: 0,
        SourceConstantAlpha: 255,
        AlphaFormat: AC_SRC_ALPHA as u8,
    };
    let ok = AlphaBlend(dst, dx, dy, dw, dh, src_dc, sx0, sy0, sw, sh, bf).as_bool();
    end_smooth_blit(dst, old_stretch);
    SelectObject(src_dc, old);
    let _ = DeleteDC(src_dc);
    let _ = DeleteObject(hbm);
    ok
}

/// `<a:alpha>` on a shape fill is composited unless this is set.
fn fill_alpha_on() -> bool {
    std::env::var("OXI_FILLALPHA_DISABLE").is_err()
}

/// `<a:alpha>` on a RUN's own colour is composited unless this is set. Only
/// the fitted-text path reads it; a run alpha outside WordArt does not occur
/// in the dev corpus.
fn txwarp_alpha_on() -> bool {
    std::env::var("OXI_TXWARPALPHA_DISABLE").is_err()
}

/// `a:custGeom` outlines are drawn as their real path unless this is set, in
/// which case the shape keeps its pre-S-CUSTGEOM bounding-box rendering.
fn custgeom_on() -> bool {
    std::env::var("OXI_CUSTGEOM_DISABLE").is_err()
}

/// A translucent solid fill is clipped to the shape's outline unless this is
/// set, which restores the bounding-box fill.
fn geomalpha_on() -> bool {
    std::env::var("OXI_GEOMALPHA_DISABLE").is_err()
}

/// A picture blended with per-pixel alpha is shrunk with a real filter before
/// the blend unless this is set, which leaves it to AlphaBlend's own sampler.
fn alphasmooth_on() -> bool {
    std::env::var("OXI_ALPHASMOOTH_DISABLE").is_err()
}

/// Images are resampled with GDI's HALFTONE filter unless this is set, which
/// restores the default BLACKONWHITE (drop-sample) mode.
fn imgsmooth_on() -> bool {
    std::env::var("OXI_IMGSMOOTH_DISABLE").is_err()
}

/// Set HALFTONE for the duration of a blit; returns the previous mode.
///
/// GDI's default stretch mode is BLACKONWHITE, which DROPS rows and columns
/// when it shrinks a bitmap. d28's engraved Lincoln portrait is a fine halftone
/// scaled down by more than 2x, and drop-sampling it produces a different moire
/// than PowerPoint's filtered one -- the portrait sits at exactly the right
/// place (best-fit shift 0,0) and still carries mean|d| 17.3 across its box.
/// MSDN requires SetBrushOrgEx after selecting HALFTONE.
#[cfg(windows)]
unsafe fn begin_smooth_blit(
    dc: windows::Win32::Graphics::Gdi::HDC,
    src: (i32, i32),
    dst: (i32, i32),
) -> i32 {
    use windows::Win32::Graphics::Gdi::*;
    // Only when the blit SHRINKS. Averaging is what a downscale needs, and the
    // corpus agrees (d05 +0.0065, d10 +0.0038, d08 +0.0032 with it on); on an
    // upscale it only softens edges PowerPoint keeps, which is where the same
    // arm lost ground (d36 -0.0009, d12 / d29 -0.0005).
    if !imgsmooth_on() || (dst.0 >= src.0 && dst.1 >= src.1) {
        return 0;
    }
    let old = SetStretchBltMode(dc, HALFTONE);
    let _ = SetBrushOrgEx(dc, 0, 0, None);
    old
}

#[cfg(windows)]
unsafe fn end_smooth_blit(dc: windows::Win32::Graphics::Gdi::HDC, old: i32) {
    use windows::Win32::Graphics::Gdi::*;
    if old != 0 {
        SetStretchBltMode(dc, STRETCH_BLT_MODE(old));
    }
}

/// Each line wraps against the width left between its own start and the shared
/// right edge unless this is set, which wraps every line at the full inner
/// width.
fn wrapwidth_on() -> bool {
    std::env::var("OXI_WRAPWIDTH_DISABLE").is_err()
}

/// An underlined run is drawn with its rule unless this is set.
fn underline_on() -> bool {
    std::env::var("OXI_UNDERLINE_DISABLE").is_err()
}

/// Italic text is drawn slanted unless this is set, which restores upright.
fn paraitalic_on() -> bool {
    std::env::var("OXI_PARAITALIC_DISABLE").is_err()
}

/// A uniformly bold paragraph is drawn bold unless this is set, which restores
/// the weight-400 single-style path.
fn parabold_on() -> bool {
    std::env::var("OXI_PARABOLD_DISABLE").is_err()
}

/// A shape's `a:blipFill` is clipped to its outline unless this is set.
fn blipclip_on() -> bool {
    std::env::var("OXI_BLIPCLIP_DISABLE").is_err()
}

/// Table cells honour their own fill / borders / margins / anchor and their
/// runs' real font size unless this is set, which restores the legacy grid.
fn tblcell_on() -> bool {
    std::env::var("OXI_TBLCELL_DISABLE").is_err()
}

/// A face outside the measured table gets its ascent from GDI unless this is
/// set, which restores the 0.9685 average for every one of them.
fn rtbaseline_on() -> bool {
    std::env::var("OXI_RTBASELINE_DISABLE").is_err()
}

/// Each line claims its own ascent and leaves its own descent unless this is
/// set, which restores the flat `prev_size * 1.2` step between paragraphs.
fn mixpitch_on() -> bool {
    std::env::var("OXI_MIXPITCH_DISABLE").is_err()
}

/// A custGeom shape filled with a gradient is left to the gradient painter
/// unless this is set, which restores the geometry path claiming it.
fn geomgrad_on() -> bool {
    std::env::var("OXI_GEOMGRAD_DISABLE").is_err()
}

/// An open custGeom subpath is stroked open unless this is set, which restores
/// the single `StrokeAndFillPath` pass that closes it first.
fn openpath_on() -> bool {
    std::env::var("OXI_OPENPATH_DISABLE").is_err()
}

/// A line is drawn broken when its `a:prstDash` says so unless this is set.
fn prstdash_on() -> bool {
    std::env::var("OXI_PRSTDASH_DISABLE").is_err()
}

/// The on/off run lengths of a `prstDash` preset, in multiples of the LINE
/// WIDTH.
///
/// Read out of PowerPoint's own PDFs (`tools/metrics/read_pptx_dash.py`): every
/// dashed stroke in the dev corpus carries its pattern verbatim in the PDF's
/// `d` operator, and dividing by the stroke width gives the same small integers
/// at every width from 0.75pt to 6pt — dash `[3 2.25]` at 0.75pt and
/// `[24 18]` at 6pt are both 4 on / 3 off. The `sys*` presets are the
/// ECMA-376 §20.1.10.49 values; the corpus does not use them, so they are not
/// measured.
fn dash_pattern(preset: &str) -> Option<&'static [u32]> {
    Some(match preset {
        "dot" => &[1, 3],
        "dash" => &[4, 3],
        "lgDash" => &[8, 3],
        "dashDot" => &[4, 3, 1, 3],
        "lgDashDot" => &[8, 3, 1, 3],
        "lgDashDotDot" => &[8, 3, 1, 3, 1, 3],
        "sysDash" => &[3, 1],
        "sysDot" => &[1, 1],
        "sysDashDot" => &[3, 1, 1, 1],
        "sysDashDotDot" => &[3, 1, 1, 1, 1, 1],
        _ => return None,
    })
}

/// An open path's stroke ends where `a:ln@cap` says unless this is set, which
/// restores GDI's default round cap everywhere.
fn line_cap_on() -> bool {
    std::env::var("OXI_LINECAP_DISABLE").is_err()
}

/// A pen for a shape outline: broken when the shape declares a `prstDash`.
///
/// `CreatePen`'s PS_DASH is cosmetic — it is ignored for any pen wider than one
/// pixel, which is most of them — so a wide dashed line needs a GEOMETRIC pen
/// with an explicit user style in device units.
///
/// `cap` is the open path's `a:ln@cap`, and `None` means the caller is stroking
/// a CLOSED outline where the cap cannot show and the legacy pen is kept as it
/// was. It matters because `CreatePen(PS_SOLID)` gives GDI's ROUND cap, which
/// overhangs the endpoint by half the line width, while PowerPoint's flat one
/// does not reach the endpoint at all: d15's 1.5pt connector runs its stem from
/// 39.786 with the tip of its head at 36.780.
///
/// PowerPoint honours the attribute exactly, and the repro's cap slide reads it
/// straight out of the PDF's stroke state: absent and "flat" both come out
/// `J=0` (butt), "rnd" `J=1`, "sq" `J=2`, at every line width. So neither
/// answer can be assumed for the other — the corpus's connectors say flat 1326
/// times and "rnd" 31 times, the widest of those 2.00pt, which is 2px of
/// overhang at 150 DPI.
#[cfg(windows)]
fn outline_pen(
    width_px: i32,
    color: u32,
    dash: Option<&str>,
    width_dev: f64,
    cap: Option<&str>,
) -> windows::Win32::Graphics::Gdi::HPEN {
    use windows::Win32::Foundation::COLORREF;
    use windows::Win32::Graphics::Gdi::*;

    // "flat" is the schema default, and it is also what a closed outline's
    // pen has always used, so an absent `cap` maps to it either way.
    let end_cap = match cap.filter(|_| line_cap_on()) {
        Some("rnd") => PS_ENDCAP_ROUND,
        Some("sq") => PS_ENDCAP_SQUARE,
        _ => PS_ENDCAP_FLAT,
    };
    let pattern = dash.filter(|_| prstdash_on()).and_then(dash_pattern);
    let Some(pattern) = pattern else {
        // A round cap IS GDI's default for a wide PS_SOLID pen, so leaving the
        // legacy pen alone is both cheaper and byte-identical.
        if cap.is_none() || !line_cap_on() || end_cap == PS_ENDCAP_ROUND {
            return unsafe { CreatePen(PS_SOLID, width_px, COLORREF(color)) };
        }
        let brush = LOGBRUSH {
            lbStyle: BS_SOLID,
            lbColor: COLORREF(color),
            lbHatch: 0,
        };
        return unsafe {
            let pen = ExtCreatePen(
                PS_GEOMETRIC | PS_SOLID | end_cap | PS_JOIN_MITER,
                width_px.max(1) as u32,
                &brush,
                None,
            );
            if pen.is_invalid() {
                CreatePen(PS_SOLID, width_px, COLORREF(color))
            } else {
                pen
            }
        };
    };
    // The run lengths are multiples of the TRUE line width, not of the pen's
    // rounded integer width: a 0.75pt line at this scale is 4.7 device units,
    // and rounding it to 5 first stretches every dash by 6%.
    let unit = if width_dev > 0.0 {
        width_dev
    } else {
        width_px.max(1) as f64
    };
    let style: Vec<u32> = pattern
        .iter()
        .map(|n| ((f64::from(*n) * unit).round() as u32).max(1))
        .collect();
    let brush = LOGBRUSH {
        lbStyle: BS_SOLID,
        lbColor: COLORREF(color),
        lbHatch: 0,
    };
    unsafe {
        let pen = ExtCreatePen(
            PS_GEOMETRIC | PS_USERSTYLE | end_cap | PS_JOIN_MITER,
            width_px.max(1) as u32,
            &brush,
            Some(&style),
        );
        if pen.is_invalid() {
            CreatePen(PS_SOLID, width_px, COLORREF(color))
        } else {
            pen
        }
    }
}

/// Table row and column EDGES are rounded from the exact running sum unless
/// this is set, which restores rounding each cell's size and adding those up.
///
/// d06 s29's Gantt grid is the tell: 14 columns of 31.917pt each: PowerPoint's
/// right edge lands at 719.52 and Oxi's at 718.80, drifting a tenth of a pixel
/// per column until the last third of the table is a whole pixel out. Every
/// rule and every bar in it moves with the drift, which is why LibreOffice --
/// which does not accumulate -- beat Oxi on that slide by 0.0296 while being
/// WORSE on its title.
fn tbledge_on() -> bool {
    std::env::var("OXI_TBLEDGE_DISABLE").is_err()
}

/// A cell spans the columns its `gridSpan` claims unless this is set, which
/// restores giving every cell exactly one column.
fn hmerge_on() -> bool {
    std::env::var("OXI_HMERGE_DISABLE").is_err()
}

/// The width one cell occupies, in points: its own column plus the ones its
/// `gridSpan` swallows.
///
/// d35 / d24 / d11 slide 29 share a Gantt template whose "Week 1" and "Week 2"
/// headers each span seven columns; without this they were drawn in a single
/// narrow column, which is also what made wrapping break them into "Wee / k 1".
fn cell_width_pt(
    table: &oxislides_core::ir::Table,
    col: usize,
    cell: &oxislides_core::ir::TableCell,
) -> f32 {
    let span = if hmerge_on() {
        (cell.grid_span.max(1) as usize).min(table.col_widths.len().saturating_sub(col).max(1))
    } else {
        1
    };
    table
        .col_widths
        .iter()
        .skip(col)
        .take(span)
        .copied()
        .sum::<f32>()
}

/// An empty paragraph counts toward its row's height only when this is set.
///
/// It is the other half of `cellbase_on` — counting it moves every row below
/// DOWN, the baseline model moves the text within a row UP — and the two were
/// each measured alone before that was noticed. Ink-band centres on d35 s35
/// against PowerPoint's own page, mean absolute error over 9 bands:
///
/// | arm                | error |
/// |--------------------|-------|
/// | shipped (both off) | 0.213pt |
/// | cellbase only      | 1.066pt |
/// | emptycell only     | 1.600pt |
/// | **both on**        | 0.533pt |
///
/// So they genuinely are each other's missing half, and still do not beat the
/// state that ships. The residual is a uniform +0.64pt, and d16 s36 wants the
/// empty cell counted (+0.0192) while d35 s35 does not — so what decides it is
/// still unknown. Both stay off.
fn emptycell_on() -> bool {
    std::env::var("OXI_EMPTYCELL_DISABLE").is_err()
}

/// A cell's text sits on the font's own baseline only when this is set.
///
/// DERIVED but HELD OPT-IN (2026-08-21). On d25 s7 PowerPoint's own PDF gives
/// three baselines whose offset below each line's top is 0.9807 / 0.9761 /
/// 0.9717 em — the face's ascent (Arial 0.9727) — so the model is right there,
/// and the slide gains 0.0123. But the corpus reads 4 improved / 2 regressed
/// (d35 s35 −0.0239, d11 s35 −0.0075 against d19 s35 +0.0036), and on d35 s35
/// the same measurement implies A ≈ 1.63 em, which no face has. Working that
/// back: those rows are consistent with the SAME model once row 0's growth is
/// taken into account (18.55pt declared, but 7pt text between 7.199pt insets
/// needs ~23pt), i.e. the residual is in the row-growth amount, not in the
/// baseline. One document is not three; pin the growth first, then re-gate.
fn cellbase_on() -> bool {
    std::env::var("OXI_CELLBASE_DISABLE").is_err()
}

/// A centred or bottom-anchored cell is positioned by the height of its
/// WRAPPED text unless this is set, which restores counting one line per
/// paragraph.
fn cellblock_on() -> bool {
    std::env::var("OXI_CELLBLOCK_DISABLE").is_err()
}

/// A translucent PRESET shape is stroked along its own outline unless this is
/// set, which restores a rectangle around its box.
fn presetstroke_on() -> bool {
    std::env::var("OXI_PRESETSTROKE_DISABLE").is_err()
}

/// `a:prstGeom prst="chevron"` — ECMA-376's homePlate with the same notch cut
/// out of its left edge.
///
/// ★UNPARKED 2026-08-25. It was held opt-in on 2026-08-24 because d35 s17's
/// three process arrows went 0.9669 -> 0.9912 (past LibreOffice's 0.9853) while
/// **d15 s17, the same template at a different aspect, LOST 0.0050** for reasons
/// that were not understood. The note at the time read "PowerPoint is not
/// compositing these three translucent shapes the way stacking them implies".
///
/// That was never a compositing rule -- it was Oxi's own GRADIENT rendering.
/// These arrows are translucent gradient fills, and S-GRADLIN / S-GRADSTOP /
/// S-GRADROT / S-PRESETGRAD / S-GRADSTROKE all landed after that measurement.
/// Re-measured on the same two slides with the same flag:
///
///     d35 s17   0.9708 -> 0.9912   (+0.0204)
///     d15 s17   0.9437 -> 0.9472   (+0.0035)   <- was -0.0050
///
/// Both slides now improve, so the reason to hold it is gone. ★The lesson is
/// about the PARK, not the shape: a change held back "until the failure is
/// explained" has to be RE-MEASURED when anything under it moves, or it stays
/// parked on evidence that has expired.
fn chevron_on() -> bool {
    std::env::var("OXI_CHEVRON_DISABLE").is_err()
}

/// A cell centres on its line's VISIBLE width unless this is set, which restores
/// counting a trailing space as ink.
fn celltrim_on() -> bool {
    std::env::var("OXI_CELLTRIM_DISABLE").is_err()
}

/// A table cell wraps its text unless this is set, which restores drawing each
/// paragraph as one line however wide the column is.
///
/// It needed horizontal merging first: while `gridSpan` was unmodelled a header
/// PowerPoint spans across seven columns measured one column wide, and wrapping
/// it there broke "Week 1" into "Wee / k 1" (d35 s29 −0.0914 and the same on
/// d24 / d11). With the span in place the width is the real one.
fn cellwrap_on() -> bool {
    std::env::var("OXI_CELLWRAP_DISABLE").is_err()
}

/// An empty paragraph is sized by its paragraph mark unless this is set,
/// which restores the pre-S-EMPTYPARA fallback to the inherited level default.
fn emptypara_on() -> bool {
    std::env::var("OXI_EMPTYPARA_DISABLE").is_err()
}

/// The font size that governs a paragraph's line advance.
///
/// A paragraph with text is sized by its runs. A paragraph with NONE is sized
/// by its paragraph mark, and PowerPoint's own export says so exactly (probe
/// `emptypara`, 10 arms, and `emptypara2`, 6 paired questions, 2026-08-18):
///
/// * `a:endParaRPr/@sz` wins -- 7 / 10 / 24 / 40pt arms advance by sz * 1.2 * n
///   on the nose, and it beats an `a:rPr` on a textless run (run 10pt +
///   endParaRPr 40pt renders 40pt).
/// * With no `endParaRPr`, the PRECEDING paragraph's size is used: prev=24
///   gives 28.80pt, prev=32 gives 38.43pt, and prev=10/next=24 gives 12.00pt,
///   so it is the paragraph before and not the one after.
/// * Only then does the inherited level default apply.
///
/// d28 slide 13 is the corpus case: a 10pt `endParaRPr` between two 10pt
/// paragraphs, under a 14pt inherited default, which Oxi drew 15px too tall at
/// 150dpi and pushed the whole lower half of the text block down with it.
fn paragraph_font_size(
    para: &oxislides_core::ir::SlideParagraph,
    inherited: Option<f32>,
    prev_fs: Option<f32>,
) -> f32 {
    let explicit = para
        .runs
        .iter()
        .filter_map(|r| r.font_size)
        .fold(None, |acc: Option<f32>, x| Some(acc.map_or(x, |a| a.max(x))));
    if emptypara_on() && para.runs.iter().all(|r| r.text.is_empty()) {
        if let Some(fs) = para.end_para_size.or(prev_fs) {
            return fs;
        }
    }
    explicit.or(inherited).unwrap_or(18.0)
}

/// `a:alphaModFix/@amt` scales a picture's opacity unless this is set.
fn imgalpha_on() -> bool {
    std::env::var("OXI_IMGALPHA_DISABLE").is_err()
}

/// The srcRect top crop is converted to StretchDIBits' bottom-up ySrc unless
/// this is set, which restores the pre-S-SRCFLIP (vertically swapped) crop.
fn srcrect_flip_on() -> bool {
    std::env::var("OXI_SRCFLIP_DISABLE").is_err()
}

/// Composite a solid fill at a constant opacity.
///
/// Derived from PowerPoint: `<a:alpha val="N"/>` inside a solidFill is a
/// constant `a = N/100000` composited straight source-over on sRGB bytes,
/// `out = a*src + (1-a)*dst`. Its PDF says the same thing at the operator
/// level -- `/BM /Normal` with `/ca` in a `/DeviceRGB` transparency group --
/// and a 10-arm probe over white / red / green backdrops (stacked translucent
/// rects included) matches to within 2/255, the residual being PowerPoint
/// quantising the alpha to 8 bits. `AlphaBlend` with `SourceConstantAlpha` and
/// no per-pixel alpha is exactly that blend, so the source is a 1x1 solid
/// stretched over the rect.
///
/// Returns false when any GDI step fails.
#[cfg(windows)]
unsafe fn alpha_fill(
    dst: windows::Win32::Graphics::Gdi::HDC,
    r: &windows::Win32::Foundation::RECT,
    rgb: (u8, u8, u8),
    alpha: f32,
) -> bool {
    use windows::Win32::Graphics::Gdi::*;

    let (dw, dh) = (r.right - r.left, r.bottom - r.top);
    if dw <= 0 || dh <= 0 {
        return false;
    }
    let bmi = BITMAPINFO {
        bmiHeader: BITMAPINFOHEADER {
            biSize: std::mem::size_of::<BITMAPINFOHEADER>() as u32,
            biWidth: 1,
            biHeight: -1, // top-down
            biPlanes: 1,
            biBitCount: 32,
            biCompression: 0, // BI_RGB
            ..Default::default()
        },
        ..Default::default()
    };
    let mut bits: *mut core::ffi::c_void = std::ptr::null_mut();
    let hbm = match CreateDIBSection(dst, &bmi, DIB_RGB_COLORS, &mut bits, None, 0) {
        Ok(b) if !bits.is_null() => b,
        _ => return false,
    };
    {
        // BGRA. AlphaFormat is 0 below, so the source is taken as opaque and
        // the alpha byte is ignored -- the opacity is SourceConstantAlpha.
        let out = std::slice::from_raw_parts_mut(bits as *mut u8, 4);
        out[0] = rgb.2;
        out[1] = rgb.1;
        out[2] = rgb.0;
        out[3] = 255;
    }
    let src_dc = CreateCompatibleDC(dst);
    if src_dc.0.is_null() {
        let _ = DeleteObject(hbm);
        return false;
    }
    let old = SelectObject(src_dc, hbm);
    let bf = BLENDFUNCTION {
        BlendOp: AC_SRC_OVER as u8,
        BlendFlags: 0,
        // PowerPoint quantises the alpha to 8 bits -- its PDF carries
        // /ca .50196 = round(0.5*255)/255 for val="50000" -- and this rounding
        // reproduces that.
        SourceConstantAlpha: (alpha.clamp(0.0, 1.0) * 255.0).round() as u8,
        AlphaFormat: 0,
    };
    let ok = AlphaBlend(dst, r.left, r.top, dw, dh, src_dc, 0, 0, 1, 1, bf).as_bool();
    SelectObject(src_dc, old);
    let _ = DeleteDC(src_dc);
    let _ = DeleteObject(hbm);
    ok
}

/// Paint a slide-background gradient over the whole page.
///
/// GDI has no primitive that covers both cases (`GradientFill` is a two-point
/// linear ramp only), so the ramp is drawn as 256 bands -- the resolution
/// PowerPoint itself uses, its PDF shading Function being a Size-256 sampled
/// array.
///
/// Linear (`a:lin`): the axis is centred on the page, runs along the angle and
/// spans the page, |w cos| + |h sin| -- probe B1's measured axis is exactly the
/// page width (0,270)->(720,270), and B3's 45-degree axis is 890.9pt =
/// (720+540)/sqrt(2). Each band is therefore a rotated quad. `scaled="1"` makes
/// the angle 45 degrees in NORMALIZED space, i.e. the direction is
/// proportional to (cos/w, sin/h): probe B6 measured a 3:4 direction on a 4:3
/// page, with both off-axis corners landing on t=0.5.
///
/// Radial (`a:path path="circle"`): concentric bands about the focus, drawn
/// outside-in after flooding the page with the end colour so rounding cannot
/// leave the corners unpainted.
#[cfg(windows)]
unsafe fn paint_bg_gradient(
    dc: windows::Win32::Graphics::Gdi::HDC,
    w: i32,
    h: i32,
    g: &SlideGradient,
) {
    use windows::Win32::Foundation::{COLORREF, POINT, RECT};
    use windows::Win32::Graphics::Gdi::*;

    const N: i32 = 256;
    if w <= 0 || h <= 0 {
        return;
    }

    // A ramp whose stops carry `a:alpha` cannot be painted as opaque bands:
    // d06's layout wash is 020F2B at 33.7% over 010C16 at 0%, which PowerPoint
    // renders as a faint darkening and an opaque painter renders as a slab.
    // Only that case takes the compositing route, so every fully opaque
    // gradient keeps its existing byte-for-byte output.
    if gradient_has_alpha(g)
        && std::env::var("OXI_GRADALPHA_DISABLE").is_err()
        && (w as i64) * (h as i64) <= 64_000_000
    {
        let mut img = image::RgbaImage::new(w as u32, h as u32);
        let t_at: Box<dyn Fn(f64, f64) -> f64> = if let Some((fx, fy)) = g.focus {
            let cx = fx as f64 * w as f64;
            let cy = fy as f64 * h as f64;
            let r_max = [
                (0.0, 0.0),
                (w as f64, 0.0),
                (0.0, h as f64),
                (w as f64, h as f64),
            ]
            .iter()
            .map(|(px, py)| ((px - cx).powi(2) + (py - cy).powi(2)).sqrt())
            .fold(0.0f64, f64::max);
            Box::new(move |x, y| {
                if r_max < 1e-9 {
                    0.0
                } else {
                    (((x - cx).powi(2) + (y - cy).powi(2)).sqrt() / r_max).clamp(0.0, 1.0)
                }
            })
        } else {
            let th = (g.angle_deg.unwrap_or(0.0) as f64).to_radians();
            let (mut dx, mut dy) = (th.cos(), th.sin());
            if g.scaled {
                dx /= w as f64;
                dy /= h as f64;
            }
            let len = (dx * dx + dy * dy).sqrt();
            if len < 1e-12 {
                dx = 1.0;
                dy = 0.0;
            } else {
                dx /= len;
                dy /= len;
            }
            let axis = (w as f64 * dx).abs() + (h as f64 * dy).abs();
            let (cx, cy) = (w as f64 / 2.0, h as f64 / 2.0);
            Box::new(move |x, y| {
                if axis < 1e-9 {
                    0.0
                } else {
                    (((x - cx) * dx + (y - cy) * dy) / axis + 0.5).clamp(0.0, 1.0)
                }
            })
        };
        for py in 0..h {
            for px in 0..w {
                let t = t_at(px as f64 + 0.5, py as f64 + 0.5);
                let (r, gg, b) = gradient_color_at(g, t);
                let a = (gradient_alpha_at(g, t) * 255.0).round().clamp(0.0, 255.0) as u8;
                img.put_pixel(px as u32, py as u32, image::Rgba([r, gg, b, a]));
            }
        }
        alpha_blit(dc, 0, 0, w, h, 0, 0, w, h, w, h, &img);
        return;
    }

    if let Some((fx, fy)) = g.focus {
        let cx = fx as f64 * w as f64;
        let cy = fy as f64 * h as f64;
        // t=1 sits on the FARTHEST page corner: measured on d04 (centred focus,
        // r=413.05 = the corner distance of a 720x405 page) and d15
        // (bottom-right focus, r=826.09 = the distance to the opposite corner).
        let r_max = [
            (0.0, 0.0),
            (w as f64, 0.0),
            (0.0, h as f64),
            (w as f64, h as f64),
        ]
        .iter()
        .map(|(x, y)| ((x - cx).powi(2) + (y - cy).powi(2)).sqrt())
        .fold(0.0_f64, f64::max);

        let end = gradient_color_at(g, 1.0);
        let brush = CreateSolidBrush(COLORREF(colorref(end.0, end.1, end.2)));
        let rect = RECT {
            left: 0,
            top: 0,
            right: w,
            bottom: h,
        };
        FillRect(dc, &rect, brush);
        let _ = DeleteObject(brush);

        let old_pen = SelectObject(dc, GetStockObject(NULL_PEN));
        for i in (0..N).rev() {
            let r = r_max * (i + 1) as f64 / N as f64;
            let c = gradient_color_at(g, (i as f64 + 0.5) / N as f64);
            let b = CreateSolidBrush(COLORREF(colorref(c.0, c.1, c.2)));
            let old = SelectObject(dc, b);
            let _ = Ellipse(
                dc,
                (cx - r).round() as i32,
                (cy - r).round() as i32,
                (cx + r).round() as i32,
                (cy + r).round() as i32,
            );
            SelectObject(dc, old);
            let _ = DeleteObject(b);
        }
        SelectObject(dc, old_pen);
        return;
    }

    let ang = g.angle_deg.unwrap_or(0.0) as f64;
    let th = ang.to_radians();
    let (mut dx, mut dy) = (th.cos(), th.sin());
    if g.scaled {
        dx /= w as f64;
        dy /= h as f64;
    }
    let len = (dx * dx + dy * dy).sqrt();
    if len < 1e-12 {
        dx = 1.0;
        dy = 0.0;
    } else {
        dx /= len;
        dy /= len;
    }
    let axis = (w as f64 * dx).abs() + (h as f64 * dy).abs();
    let (cx, cy) = (w as f64 / 2.0, h as f64 / 2.0);
    // Perpendicular half-extent: the page diagonal covers any rotation.
    let half = ((w as f64) * (w as f64) + (h as f64) * (h as f64)).sqrt();
    let (nx, ny) = (-dy, dx);

    let old_pen = SelectObject(dc, GetStockObject(NULL_PEN));
    for i in 0..N {
        let t0 = i as f64 / N as f64;
        let t1 = (i + 1) as f64 / N as f64;
        let c = gradient_color_at(g, (t0 + t1) / 2.0);
        // Overlap adjacent bands by a pixel so integer rounding cannot leave a
        // seam between the quads.
        let a0 = (t0 - 0.5) * axis - 1.0;
        let a1 = (t1 - 0.5) * axis + 1.0;
        let pts = [
            (cx + dx * a0 + nx * half, cy + dy * a0 + ny * half),
            (cx + dx * a1 + nx * half, cy + dy * a1 + ny * half),
            (cx + dx * a1 - nx * half, cy + dy * a1 - ny * half),
            (cx + dx * a0 - nx * half, cy + dy * a0 - ny * half),
        ];
        let quad: [POINT; 4] = [
            POINT {
                x: pts[0].0.round() as i32,
                y: pts[0].1.round() as i32,
            },
            POINT {
                x: pts[1].0.round() as i32,
                y: pts[1].1.round() as i32,
            },
            POINT {
                x: pts[2].0.round() as i32,
                y: pts[2].1.round() as i32,
            },
            POINT {
                x: pts[3].0.round() as i32,
                y: pts[3].1.round() as i32,
            },
        ];
        let b = CreateSolidBrush(COLORREF(colorref(c.0, c.1, c.2)));
        let old = SelectObject(dc, b);
        let _ = Polygon(dc, &quad);
        SelectObject(dc, old);
        let _ = DeleteObject(b);
    }
    SelectObject(dc, old_pen);
}

/// `a:headEnd` / `a:tailEnd` decorations are drawn unless this is set.
fn line_ends_on() -> bool {
    std::env::var("OXI_LINEEND_DISABLE").is_err()
}

/// How many line widths the `@w` / `@len` size tokens are worth.
fn line_end_factor(tok: &str) -> f64 {
    match tok {
        "sm" => 2.0,
        "lg" => 5.0,
        _ => 3.0,
    }
}

/// The width the size tokens actually scale: the line's own, but never under
/// **2.00pt**.
///
/// Read off PowerPoint's own PDF, which carries each head as a real filled
/// path, so these are exact rather than pixel estimates. Above the floor the
/// factor is plain: a `med` triangle is 9.000pt across on a 3.00pt line and
/// 22.500pt on a 7.50pt one. Below it the factor alone is wrong: a `sm`
/// triangle on a 1.50pt line is 4.000pt across, not 3.000, and a `med` oval on
/// a 0.75pt line is 6.000pt, not 2.250 -- both land exactly on their factor
/// times 2.00.
///
/// `gen_pptx_lineend.py` sweeps 5 line widths x 5 (w, len) pairs x 5 types and
/// `read_pptx_lineend_probe.py` reads the result: for oval, triangle, stealth
/// and diamond all 100 measurements are the predicted size to three decimals,
/// and the 0.75pt and 1.50pt rows are IDENTICAL while 3.00 / 4.50 / 6.00 scale
/// linearly -- flat below the floor, proportional above it. `w` and `len` act
/// on the two axes independently (a `sm`/`lg` head is 4.000 across by 10.000
/// along). Four dev decks agree.
fn line_end_unit(lw_pt: f64) -> f64 {
    lw_pt.max(2.0)
}

/// Draw one line-end decoration.
///
/// `(ex, ey)` is the decorated end in device pixels and `(bx, by)` any point
/// behind it along the line; together they fix the direction the decoration
/// points. `lw_pt` is the stroke width in POINTS -- the floor in
/// `line_end_unit` is a point value, so it cannot be applied after scaling.
#[cfg(windows)]
unsafe fn draw_line_end(
    dc: windows::Win32::Graphics::Gdi::HDC,
    end: &LineEnd,
    ex: i32,
    ey: i32,
    bx: i32,
    by: i32,
    lw_pt: f64,
    scale: f64,
    rgb: (u8, u8, u8),
) {
    use windows::Win32::Foundation::{COLORREF, POINT};
    use windows::Win32::Graphics::Gdi::*;

    let (dx, dy) = ((ex - bx) as f64, (ey - by) as f64);
    let len = (dx * dx + dy * dy).sqrt();
    if len < 0.5 || lw_pt <= 0.0 {
        return;
    }
    // u runs along the line towards the decorated end, v across it.
    let (ux, uy) = (dx / len, dy / len);
    let (vx, vy) = (-uy, ux);
    let unit = line_end_unit(lw_pt) * scale;
    let across = line_end_factor(&end.w) * unit;
    let along = line_end_factor(&end.len) * unit;
    let at = |u: f64, v: f64| POINT {
        x: (ex as f64 + u * ux + v * vx).round() as i32,
        y: (ey as f64 + u * uy + v * vy).round() as i32,
    };

    // Each kind's outline in (along, across), origin at the declared endpoint,
    // as PowerPoint's PDF draws it. The oval is CENTRED on the end -- d35's
    // circles have their centres exactly on the connector's endpoints and
    // reach half a diameter past. The pointed kinds put their TIP on the end
    // and their back a whole `along` behind it (d15's triangle: tip 36.780 =
    // the declared endpoint, back 40.780 = 4.000 later, base spanning 4.000).
    let pts: Vec<POINT> = match end.kind.as_str() {
        "oval" => (0..48)
            .map(|i| {
                let t = f64::from(i) * std::f64::consts::TAU / 48.0;
                at(along / 2.0 * t.cos(), across / 2.0 * t.sin())
            })
            .collect(),
        "triangle" => vec![
            at(0.0, 0.0),
            at(-along, across / 2.0),
            at(-along, -across / 2.0),
        ],
        // The notch reaches two thirds of the way back: d32's stealth head is
        // tip 1198.730, back corners 1185.230 (13.500 = 3 x 4.50 behind) and
        // notch vertex 1189.730, i.e. 9.000 back of 13.500.
        "stealth" => vec![
            at(0.0, 0.0),
            at(-along, across / 2.0),
            at(-along * 2.0 / 3.0, 0.0),
            at(-along, -across / 2.0),
        ],
        // Centred like the oval, not tip-on-end like the pointed kinds: in the
        // repro every diamond reaches exactly half its length past the end.
        "diamond" => vec![
            at(along / 2.0, 0.0),
            at(0.0, across / 2.0),
            at(-along / 2.0, 0.0),
            at(0.0, -across / 2.0),
        ],
        // "arrow" is an open V that PowerPoint STROKES rather than fills, so it
        // does not follow the law above and needs its own derivation. No
        // dev-corpus line asks for one; leaving it undrawn beats drawing it
        // wrong.
        _ => return,
    };

    let brush = CreateSolidBrush(COLORREF(colorref(rgb.0, rgb.1, rgb.2)));
    let pen = CreatePen(PS_SOLID, 1, COLORREF(colorref(rgb.0, rgb.1, rgb.2)));
    let old_b = SelectObject(dc, brush);
    let old_p = SelectObject(dc, pen);
    let _ = Polygon(dc, &pts);
    SelectObject(dc, old_p);
    SelectObject(dc, old_b);
    let _ = DeleteObject(pen);
    let _ = DeleteObject(brush);
}

#[cfg(windows)]
fn render_slides_gdi(pres: &Presentation, prefix: &str, dpi: u32, supersample: u32) {
    use windows::Win32::Foundation::*;
    use windows::Win32::Graphics::Gdi::*;

    let render_dpi = dpi * supersample.max(1);
    let scale = render_dpi as f64 / 72.0;

    for (si, slide) in pres.slides.iter().enumerate() {
        let out_w = (pres.slide_width as f64 * dpi as f64 / 72.0).round() as u32;
        let out_h = (pres.slide_height as f64 * dpi as f64 / 72.0).round() as u32;
        let w = (pres.slide_width as f64 * scale).round() as i32;
        let h = (pres.slide_height as f64 * scale).round() as i32;

        unsafe {
            let screen_dc = GetDC(HWND(std::ptr::null_mut()));
            let mem_dc = CreateCompatibleDC(screen_dc);
            let bitmap = CreateCompatibleBitmap(screen_dc, w, h);
            let old_bmp = SelectObject(mem_dc, bitmap);

            // Background, in PowerPoint's own precedence: the picture when the
            // slide (or its layout/master) has one, else the gradient, else the
            // flat colour, else white. The compatible bitmap starts
            // UNINITIALISED, so exactly one of these must always run -- when a
            // picture cannot be decoded paint_bg_image reports false and we
            // fall through to the next fill rather than leave garbage pixels.
            let painted = match slide.background_image.as_ref() {
                Some(bi) => paint_bg_image(mem_dc, w, h, bi),
                None => false,
            };
            if painted {
                // the picture already covers the whole page
            } else if let Some(g) = slide.background_gradient.as_ref() {
                paint_bg_gradient(mem_dc, w, h, g);
            } else {
                let bg = slide
                    .background_color
                    .as_deref()
                    .and_then(parse_hex_rgb)
                    .unwrap_or((255, 255, 255));
                let bg_brush = CreateSolidBrush(COLORREF(colorref(bg.0, bg.1, bg.2)));
                let rect = RECT {
                    left: 0,
                    top: 0,
                    right: w,
                    bottom: h,
                };
                FillRect(mem_dc, &rect, bg_brush);
                let _ = DeleteObject(bg_brush);
            }
            SetBkMode(mem_dc, TRANSPARENT);

            for sh in &slide.shapes {
                let x = (sh.x as f64 * scale).round() as i32;
                let y = (sh.y as f64 * scale).round() as i32;
                let ew = (sh.width as f64 * scale).round() as i32;
                let eh = (sh.height as f64 * scale).round() as i32;

                // A connector (p:cxnSp) is a LINE, not a box. PowerPoint
                // render-truth (dev corpus vs the exported PDF): it runs
                // from the xfrm box's top-left to its bottom-right, a flip
                // selects the other diagonal, and @rot rotates both
                // endpoints about the box centre -- verified to 0.01pt on
                // plain, flipH+rot180 and rot=-90 connectors. No connector
                // in the corpus carries text (0 of 1357), so the box draw
                // and the text pass below are both skipped.
                if sh
                    .shape_type
                    .as_deref()
                    .is_some_and(|p| p.contains("Connector"))
                {
                    let bw = sh.border_width.unwrap_or(0.0);
                    if bw > 0.0 {
                        let col = sh
                            .border_color
                            .as_deref()
                            .and_then(parse_hex_rgb)
                            .unwrap_or((0, 0, 0));
                        let (x0, y0, x1, y1) = if sh.flip_h != sh.flip_v {
                            (sh.x + sh.width, sh.y, sh.x, sh.y + sh.height)
                        } else {
                            (sh.x, sh.y, sh.x + sh.width, sh.y + sh.height)
                        };
                        let cx = (sh.x + sh.width / 2.0) as f64;
                        let cy = (sh.y + sh.height / 2.0) as f64;
                        let (sn, cs) = (sh.rotation as f64).to_radians().sin_cos();
                        let map = |px: f32, py: f32| {
                            let (dx, dy) = (px as f64 - cx, py as f64 - cy);
                            (
                                ((cx + dx * cs - dy * sn) * scale).round() as i32,
                                ((cy + dx * sn + dy * cs) * scale).round() as i32,
                            )
                        };
                        // S-BENTCONN (2026-08-25): `bentConnector3` is an ELBOW,
                        // not a diagonal. Its local path is
                        //     (0,0) -> (adj*w, 0) -> (adj*w, h) -> (w, h)
                        // with `adj = adj1/100000` (default 0.5), flipped and
                        // turned about the box centre by the same `map` a
                        // straight connector uses. Read straight out of
                        // PowerPoint's own PDF vectors on d11 slide 12 and
                        // matched on ALL FOUR points of both flip states to
                        // 0.01pt -- e.g. the unflipped 55.77x109.37 box at
                        // rot=-90 gives (149.09,217.95) (149.09,190.06)
                        // (258.46,190.06) (258.46,162.18), exactly what
                        // PowerPoint drew.
                        //
                        // 18 of them in the corpus (d11 / d19 / d24, six each),
                        // and each sits on an org-chart slide where the
                        // connectors ARE the diagram: Oxi drew six diagonals
                        // across three trees.
                        let bent = bentconn_on()
                            && sh.shape_type.as_deref() == Some("bentConnector3");
                        let elbow: Vec<(i32, i32)> = if bent {
                            let adj = sh
                                .adjustments
                                .get("adj1")
                                .copied()
                                .unwrap_or(50_000.0)
                                / 100_000.0;
                            let (w, h) = (sh.width, sh.height);
                            let xb = w * adj;
                            [(0.0, 0.0), (xb, 0.0), (xb, h), (w, h)]
                                .iter()
                                .map(|&(mut lx, mut ly)| {
                                    if sh.flip_h {
                                        lx = w - lx;
                                    }
                                    if sh.flip_v {
                                        ly = h - ly;
                                    }
                                    map(sh.x + lx, sh.y + ly)
                                })
                                .collect()
                        } else {
                            Vec::new()
                        };
                        let (ax, ay) = map(x0, y0);
                        let (bx, by) = map(x1, y1);
                        let pen = outline_pen(
                            (bw as f64 * scale).round().max(1.0) as i32,
                            colorref(col.0, col.1, col.2),
                            sh.border_dash.as_deref(),
                            bw as f64 * scale,
                            Some(sh.line_cap.as_deref().unwrap_or("flat")),
                        );
                        let old_pen = SelectObject(mem_dc, pen);
                        if elbow.len() == 4 {
                            let _ = MoveToEx(mem_dc, elbow[0].0, elbow[0].1, None);
                            for pt in &elbow[1..] {
                                let _ = LineTo(mem_dc, pt.0, pt.1);
                            }
                        } else {
                            let _ = MoveToEx(mem_dc, ax, ay, None);
                            let _ = LineTo(mem_dc, bx, by);
                        }
                        SelectObject(mem_dc, old_pen);
                        let _ = DeleteObject(pen);
                        if line_ends_on() {
                            let lw = f64::from(bw);
                            // A decoration points along the segment it sits on,
                            // which for an elbow is the first / last one.
                            let (ha, hb, ta, tb) = if elbow.len() == 4 {
                                (elbow[0], elbow[1], elbow[3], elbow[2])
                            } else {
                                ((ax, ay), (bx, by), (bx, by), (ax, ay))
                            };
                            if let Some(h) = sh.head_end.as_ref() {
                                draw_line_end(mem_dc, h, ha.0, ha.1, hb.0, hb.1, lw, scale, col);
                            }
                            if let Some(t) = sh.tail_end.as_ref() {
                                draw_line_end(mem_dc, t, ta.0, ta.1, tb.0, tb.1, lw, scale, col);
                            }
                        }
                    }
                    continue;
                }

                // DrawingML geometry: an explicit custGeom outline first (it is
                // the shape's real boundary), then the named presets.
                // Unsupported ones retain the legacy rectangular fallback below.
                //
                // custGeom is NOT gated on AutoShape the way the presets are: a
                // shape with custom geometry has no prstGeom, so the parser
                // classifies it as TextBox when it carries text and Placeholder
                // when it does not -- gating on AutoShape would skip every one
                // of them. Pictures keep their own draw path untouched.
                let drew_preset = !matches!(&sh.content, ShapeContent::Image { .. })
                    && (draw_custom_geometry_gdi(mem_dc, sh, scale)
                        || (matches!(&sh.content, ShapeContent::AutoShape { .. })
                            && draw_preset_shape_gdi(mem_dc, sh, scale)));

                // A gradient fill paints the shape's own area (clipped to
                // its outline) and stands in for the solid fill below.
                let drew_gradient = match sh.gradient.as_ref() {
                    Some(g) if !drew_preset => paint_shape_gradient(mem_dc, sh, g, scale),
                    _ => false,
                };

                // Fill. A preset path fills itself, so this rectangular
                // fallback only runs for the shapes it declined.
                if !drew_preset && !drew_gradient {
                    if let Some(fill) = &sh.fill_color {
                        if let Some((r, g, b)) = parse_hex_rgb(fill) {
                            // <a:alpha> makes the fill translucent; PowerPoint
                            // composites it straight source-over on sRGB bytes
                            // (S-FILLALPHA), which is what AlphaBlend's
                            // SourceConstantAlpha does. An absent or full alpha
                            // takes the plain opaque FillRect.
                            //
                            // A translucent fill NEVER falls back to the opaque
                            // brush: painting a 0%-alpha shape solid is the very
                            // bug this fixes -- d23 carries 15 of them, custGeom
                            // frames whose only visible ink is a green outline,
                            // and each was drawn as a black slab over a quarter
                            // of the page. "Incorrect ink is worse than none."
                            let r2 = RECT {
                                left: x,
                                top: y,
                                right: x + ew,
                                bottom: y + eh,
                            };
                            match sh.fill_alpha.filter(|_| fill_alpha_on()) {
                                Some(a) if a < 1.0 => {
                                    if a > 0.0 {
                                        // A translucent fill cannot be painted
                                        // with a GDI brush, so
                                        // `draw_custom_geometry_gdi` declines
                                        // the shape and this box is what gets
                                        // painted. Clip it to the outline, or
                                        // d04 slide 13's world map -- one
                                        // custGeom of 154 subpaths at 50%
                                        // alpha -- arrives as a single pink
                                        // slab over the continents.
                                        let clipped = geomalpha_on()
                                            && clip_to_geometry_gdi(mem_dc, sh, scale);
                                        alpha_fill(mem_dc, &r2, (r, g, b), a);
                                        if clipped {
                                            let _ = SelectClipRgn(mem_dc, None);
                                        }
                                    }
                                }
                                _ => {
                                    let brush =
                                        CreateSolidBrush(COLORREF(colorref(r, g, b)));
                                    FillRect(mem_dc, &r2, brush);
                                    let _ = DeleteObject(brush);
                                }
                            }
                        }
                    }
                }

                // Border
                let border_w = sh.border_width.unwrap_or(0.0);
                if !drew_preset && border_w > 0.0 {
                    let col = sh
                        .border_color
                        .as_deref()
                        .and_then(parse_hex_rgb)
                        .unwrap_or((0, 0, 0));
                    let pen = outline_pen(
                        (border_w as f64 * scale).round() as i32,
                        colorref(col.0, col.1, col.2),
                        sh.border_dash.as_deref(),
                        border_w as f64 * scale,
                        None,
                    );
                    let old_pen = SelectObject(mem_dc, pen);
                    let old_brush = SelectObject(mem_dc, GetStockObject(NULL_BRUSH));
                    // A shape that HAS an outline should be stroked along it,
                    // not around its box. This only arises when the geometry
                    // path declined the shape (a translucent fill), and d04
                    // slide 13 is the case: PowerPoint traces every coastline
                    // in 0.75pt white, Oxi drew one rectangle around the map.
                    let stroked = if geomalpha_on() {
                        match drawable_geometry(sh) {
                            Some(geom) => {
                                let _ = BeginPath(mem_dc);
                                for path in &geom.paths {
                                    emit_geom_path_gdi(mem_dc, sh, path, scale);
                                }
                                let _ = EndPath(mem_dc);
                                StrokePath(mem_dc).as_bool()
                            }
                            // ★A PRESET outline is an outline too. The fill
                            // above is already clipped to it, so leaving the
                            // stroke as a rectangle drew a box around a shape
                            // that had just been cut to an arrow -- d35 s17's
                            // three process chevrons at 50% alpha came out as
                            // boxes with a faint wedge inside them. 132 non-rect
                            // presets over 8 dev decks carry a translucent fill
                            // and reach this path.
                            None if presetstroke_on() => {
                                let _ = BeginPath(mem_dc);
                                let ok = emit_shape_path(mem_dc, sh, scale);
                                let _ = EndPath(mem_dc);
                                ok && StrokePath(mem_dc).as_bool()
                            }
                            None => false,
                        }
                    } else {
                        false
                    };
                    if !stroked {
                        let _ = Rectangle(mem_dc, x, y, x + ew, y + eh);
                    }
                    SelectObject(mem_dc, old_brush);
                    SelectObject(mem_dc, old_pen);
                    let _ = DeleteObject(pen);
                }

                // Text (Spec #4 layout: wrap at word boundaries within the
                // effective width, place each line at its baseline). AutoShapes
                // with a text body render their text too.
                match &sh.content {
                    ShapeContent::TextBox { paragraphs }
                    | ShapeContent::AutoShape { paragraphs } => {
                        // WordArt: an AUTOSHAPE carrying `a:prstTxWarp` has its
                        // text's INK BOX mapped onto the shape box exactly, and
                        // stretched independently in each axis. Derived from
                        // PowerPoint's own export (probe `txwarp`, 6 arms over
                        // three faces and four aspect ratios: ink offset 0.000
                        // and ink/box ratio 1.000 every time), and confirmed on
                        // the corpus specimen d35 s4 (box 75.9 x 303.5, ink
                        // 75.8 x 297.5). A plain TEXT BOX with the same element
                        // is left alone -- the same probe's textbox arms all
                        // render at the default 18pt.
                        if txwarp_on()
                            && sh.text_warp.is_some()
                            && matches!(sh.content, ShapeContent::AutoShape { .. })
                        {
                            if draw_warped_text(mem_dc, pres, sh, paragraphs, scale) {
                                continue;
                            }
                        }
                        // S-TEXTROT (2026-08-24): the text of a turned shape is
                        // laid out in the shape's OWN box, exactly as if `rot`
                        // were zero, and the whole result is then turned about
                        // that box's centre. So nothing above this line changes
                        // -- only the paint does, through one world transform
                        // that carries the glyphs, the highlight boxes and the
                        // bullet markers together.
                        //
                        // WordArt takes its own path above and stays upright;
                        // no corpus `prstTxWarp` shape is turned.
                        let turned_text = begin_turned_text(mem_dc, sh, paragraphs, scale);
                        let (geom_h_ins, geom_v_ins) = geom_text_inset(sh);
                        let left_x =
                            x + ((sh.l_ins + geom_h_ins) as f64 * scale).round() as i32;
                        let right_x = x
                            + ((sh.width - sh.r_ins - geom_h_ins) as f64 * scale).round() as i32;
                        let mut cursor_pt = sh.y + sh.t_ins + geom_v_ins;
                        let anchor_off = compute_shape_anchor_off(mem_dc, pres, sh);
                        let master_ctx: &Vec<MasterStyleLevel> = match sh.ph_type.as_deref() {
                            Some("title") | Some("ctrTitle") => &pres.master_styles.title,
                            Some(_) => &pres.master_styles.body,
                            None => &pres.master_styles.other,
                        };
                        // Spec #11: one AutoNum counter set per text box (the
                        // counters continue across the box's paragraphs, keyed
                        // by (level, kind) with a startAt reset, inside
                        // layout_paragraph_baselines).
                        let mut counters =
                            std::collections::HashMap::<(u32, String), (Option<u32>, u32)>::new();
                        let mut prev_fs: Option<f32> = None;
                        for (pi, p) in paragraphs.iter().enumerate() {
                            // Effective font size: a run's explicit sz wins (the
                            // max over runs); else the master txStyles level
                            // default (Spec #5, phfs probe: V2 layout sz is
                            // ignored, V3 run 14pt overrides master 32pt); else
                            // the engine default 18pt.
                            let m = if master_ctx.is_empty() {
                                None
                            } else {
                                Some(&master_ctx[(p.lvl as usize).min(master_ctx.len() - 1)])
                            };
                            // The LAYOUT placeholder's own a:lstStyle sits
                            // between the run and the master txStyles: d24's
                            // master titleStyle has no size or colour while its
                            // layout ctrTitle declares sz=6000 + lt1, and
                            // PowerPoint draws 60pt white.
                            let phl = if sh.ph_levels.is_empty() {
                                None
                            } else {
                                Some(&sh.ph_levels[(p.lvl as usize).min(sh.ph_levels.len() - 1)])
                            };
                            let m_fs = phl
                                .and_then(|l| l.font_size)
                                .or_else(|| m.and_then(|mm| mm.font_size));
                            let fs = paragraph_font_size(p, m_fs, prev_fs);
                            let family = effective_family(
                                mem_dc,
                                &paragraph_family(
                                    pres,
                                    sh,
                                    p,
                                    &sh.ph_levels[..],
                                    m.map(std::slice::from_ref).unwrap_or(&[]),
                                ),
                            );
                            let color = p
                                .runs
                                .iter()
                                .find_map(|r| r.color.clone())
                                .or_else(|| phl.and_then(|l| l.color.clone()));
                            // A level can declare the highlight too, and then
                            // every run inherits it: d35's master title level
                            // carries the white slab behind "BIG CONCEPT".
                            // The default for runs that state NO colour. It must be
                            // the level's, never another run's: d16 slide 5's
                            // quotation is two uncoloured runs around one accent1
                            // run, and taking the first colour found painted the
                            // whole quotation blue where PowerPoint has black.
                            // 30 paragraphs over 9 decks mix the two.
                            let run_default_color = if runcolordef_on() {
                                phl.and_then(|l| l.color.clone())
                            } else {
                                color.clone()
                            };
                            let para_highlight = if highlightlvl_on() {
                                phl.and_then(|l| l.highlight.clone())
                            } else {
                                None
                            };
                            let (lines, marker) = layout_paragraph_baselines(
                                mem_dc,
                                p,
                                &mut cursor_pt,
                                sh.width,
                                scale,
                                pi == 0,
                                &family,
                                sh.l_ins + geom_h_ins,
                                sh.r_ins + geom_h_ins,
                                &master_ctx[..],
                                &sh.ph_levels[..],
                                anchor_off,
                                &mut counters,
                                &mut prev_fs,
                                sh.wrap_text,
                            );
                            if let Some(m) = &marker {
                                let marker_x =
                                    left_x + (m.x_pt as f64 * scale).round() as i32;
                                draw_text_baseline(
                                    mem_dc,
                                    marker_x,
                                    m.baseline,
                                    &m.text,
                                    m.fs,
                                    &m.font,
                                    color.as_deref(),
                                    scale,
                                );
                            }
                            // Spec #6: horizontal alignment resolution — a
                            // paragraph with no explicit alignment inherits
                            // the master txStyles level's algn (then Left).
                            let align = p
                                .alignment
                                .unwrap_or(m.and_then(|mm| mm.algn).unwrap_or(SlideAlignment::Left));
                            let is_justify = matches!(align, SlideAlignment::Justify);
                            let n_lines = lines.len();
                            // OXI_DUMP_LINES=1 prints what the text layout
                            // actually produced. Pixel forensics cannot tell a
                            // short last line from the next paragraph's first
                            // line, and that ambiguity is exactly what makes
                            // line-breaking differences hard to attribute.
                            if std::env::var("OXI_DUMP_LINES").is_ok() {
                                eprintln!(
                                    "LINES slide={} shape=({:.1},{:.1}) para={} fs={:.2} n={} lnSpc={:?}",
                                    slide.index, sh.x, sh.y, pi, fs, n_lines, p.line_spacing
                                );
                                for (li, (t, b, xo)) in lines.iter().enumerate() {
                                    eprintln!(
                                        "  L{:<2} baseline={:.2}pt x_off={:.2} {:?}",
                                        li,
                                        *b as f64 / scale,
                                        xo,
                                        t.chars().take(60).collect::<String>()
                                    );
                                }
                            }
                            // A paragraph whose runs are all bold is UNIFORM,
                            // so it never takes the per-run path -- and the
                            // single-style path called `draw_text_baseline`,
                            // which hardcodes weight 400. d19 slide 2's
                            // headings are one run each with `b="1"` and came
                            // out regular; 1377 paragraphs across all 40 decks
                            // are all-bold like that.
                            // A LEVEL asks for WEIGHT the same way: d11's
                            // master title placeholder declares
                            // `<a:defRPr b="1" sz="3200">` with Kulim Park, and
                            // Oxi took the size and the face from that level
                            // while leaving the weight at 400 -- "Team
                            // Presentation" 264.5pt of ink against
                            // PowerPoint's 271.7 at 32pt.
                            let lvl_bold =
                                lvlbold_on() && phl.and_then(|l| l.bold).unwrap_or(false);
                            let para_weight =
                                if (p.runs.iter().any(|r| r.bold) || lvl_bold) && parabold_on() {
                                    700
                                } else {
                                    400
                                };
                            // A LEVEL can ask for italic too: d16's layout body
                            // level declares `<a:defRPr i="1"/>` and PowerPoint
                            // sets the whole quotation slanted.
                            let lvl_italic = lvlitalic_on() && phl.is_some_and(|l| l.italic);
                            let para_italic = lvl_italic || p.runs.iter().any(|r| r.italic);
                            let para_ul = p.runs.iter().any(|r| r.underline);
                            // A highlight needs the per-run path even when the
                            // paragraph is one run, since only that path knows
                            // where a run starts and ends on the line.
                            let has_highlight = highlight_on()
                                && (para_highlight.is_some()
                                    || p.runs.iter().any(|r| r.highlight.is_some()));
                            let styled = runstyle_on()
                                && (has_highlight
                                    || (p.runs.len() > 1
                                        && (p.runs.iter().any(|r| r.bold != p.runs[0].bold)
                                            || p.runs.iter().any(|r| r.color != p.runs[0].color)
                                            || p.runs
                                                .iter()
                                                .any(|r| r.font_size != p.runs[0].font_size)
                                            || p.runs.iter().any(|r| r.italic != p.runs[0].italic)
                                            || (underline_on()
                                                && p.runs.iter().any(|r| {
                                                    r.underline != p.runs[0].underline
                                                })))));
                            let mut line_off = 0usize;
                            for (i, (line_text, baseline, x_off)) in
                                lines.into_iter().enumerate()
                            {
                                let this_off = line_off;
                                line_off += line_text.chars().count();
                                // The trailing newline of a soft-broken line is
                                // accounting, not ink: it is counted above so
                                // the run mapping stays aligned, and dropped
                                // here so nothing tries to draw it.
                                let line_text = line_text.trim_end_matches('\n').to_string();
                                if line_text.trim().is_empty() {
                                    continue;
                                }
                                if styled && !(is_justify && i + 1 < n_lines) {
                                    let line_x = left_x
                                        + (x_off as f64 * scale).round() as i32;
                                    draw_line_runs(
                                        mem_dc,
                                        line_x,
                                        baseline,
                                        &line_text,
                                        this_off,
                                        &p.runs,
                                        &family,
                                        fs,
                                        run_default_color.as_deref(),
                                        para_highlight.as_deref(),
                                        lvl_italic,
                                        lvl_bold,
                                        scale,
                                    );
                                    continue;
                                }
                                if is_justify && i + 1 < n_lines {
                                    // Non-final justified line: spread the
                                    // stretch over the inter-word gaps.
                                    draw_text_justify(
                                        mem_dc,
                                        left_x,
                                        right_x,
                                        baseline,
                                        &line_text,
                                        fs,
                                        &family,
                                        color.as_deref(),
                                        scale,
                                    );
                                } else {
                                    let line_x = left_x
                                        + (x_off as f64 * scale).round() as i32;
                                    draw_text_baseline_wiu(
                                        mem_dc,
                                        line_x,
                                        baseline,
                                        &line_text,
                                        fs,
                                        &family,
                                        color.as_deref(),
                                        scale,
                                        para_weight,
                                        para_italic,
                                        para_ul,
                                    );
                                }
                            }
                        }
                        end_turned_text(mem_dc, turned_text);
                    }
                    ShapeContent::Table { table } => {
                        // A DrawingML cell states its own fill, its four
                        // borders, its margins and its anchor, so none of it
                        // needs the table style resolved. The pre-S-TBLCELL
                        // path ignored all of it: it stroked a black 1px
                        // rectangle around every cell (PowerPoint draws the
                        // few rules the cells actually declare, and an
                        // invisible edge is written as alpha=0, not omitted)
                        // and it sized text with `fold(18.0, f32::max)`, which
                        // rendered an 8pt planner grid at 18pt.
                        if !tblcell_on() {
                            draw_table_legacy(mem_dc, pres, sh, table, x, y, scale);
                            continue;
                        }
                        // `a:tr/@h` is a MINIMUM. PowerPoint grows a row until
                        // its tallest cell's text fits inside that cell's own
                        // margins, and everything below moves down with it.
                        // d35 slide 35's planner header declares 18.55pt and
                        // holds 7pt text in a cell with marT = marB = 7.2pt --
                        // twice the default -- so PowerPoint draws it 23.5pt and
                        // Oxi's seven data rows all sat 5.3pt high. 4 of the
                        // corpus's 181 table rows need the growth, one each in
                        // d11, d15, d16 and d35: the same template page.
                        let row_h: Vec<f32> = table
                            .rows
                            .iter()
                            .enumerate()
                            .map(|(r, row)| {
                                let declared =
                                    table.row_heights.get(r).copied().unwrap_or(0.0);
                                if !tblgrow_on() {
                                    return declared;
                                }
                                // Only a cell whose runs all STATE a size can be
                                // measured here: the draw path falls back to 18pt
                                // for a run that inherits one, and feeding that
                                // guess into the row height grew rows that should
                                // not move -- four decks lost 0.003 to 0.006 each.
                                let mut need = 0.0f32;
                                for (c, cell) in row.iter().enumerate() {
                                    let mut text = 0.0f32;
                                    let mut sized = true;
                                    for p in &cell.paragraphs {
                                        // An EMPTY paragraph still occupies a
                                        // line, sized by its paragraph mark.
                                        let empty_sz = p
                                            .runs
                                            .iter()
                                            .all(|r| r.text.trim().is_empty())
                                            .then_some(p.end_para_size)
                                            .flatten()
                                            .filter(|_| emptycell_on());
                                        if let Some(sz) = empty_sz {
                                            text += sz * 1.2;
                                            continue;
                                        }
                                        match p.runs.iter().find_map(|r| r.font_size) {
                                            Some(fs) => {
                                                // A wrapped paragraph is as many
                                                // lines as the column forces, and
                                                // the row has to hold all of them.
                                                let mut n = 1usize;
                                                if cellwrap_on() {
                                                    let inner = cell_width_pt(table, c, cell)
                                                        - cell.mar_l
                                                        - cell.mar_r;
                                                    let body: String = p
                                                        .runs
                                                        .iter()
                                                        .map(|r| r.text.as_str())
                                                        .collect();
                                                    if inner > 0.0 && !body.trim().is_empty() {
                                                        let family = effective_family(
                                                            mem_dc,
                                                            &paragraph_family(
                                                                pres,
                                                                sh,
                                                                p,
                                                                &sh.ph_levels[..],
                                                                &[],
                                                            ),
                                                        );
                                                        let bold =
                                                            p.runs.iter().any(|r| r.bold);
                                                        n = gdi_wrap_lines(
                                                            mem_dc, &body, inner, inner,
                                                            scale, fs, &family, bold, false,
                                                            Some((&p.runs[..], 0)),
                                                        )
                                                        .len()
                                                        .max(1);
                                                    }
                                                }
                                                text += fs * 1.2 * n as f32;
                                            }
                                            None => sized = false,
                                        }
                                    }
                                    if sized && !cell.paragraphs.is_empty() {
                                        need = need.max(text + cell.mar_t + cell.mar_b);
                                    }
                                }
                                if std::env::var("OXI_ROW_DEBUG").is_ok() {
                                    eprintln!(
                                        "ROW r={r} declared={declared:.3} need={need:.3} \
                                         -> {:.3}",
                                        declared.max(need)
                                    );
                                }
                                declared.max(need)
                            })
                            .collect();
                        let mut cy = y;
                        let mut cy_pt = 0.0f64;
                        for (r, row) in table.rows.iter().enumerate() {
                            let h_pt = row_h.get(r).copied().unwrap_or(0.0) as f64;
                            let ph = if tbledge_on() {
                                cy = y + (cy_pt * scale).round() as i32;
                                y + ((cy_pt + h_pt) * scale).round() as i32 - cy
                            } else {
                                (h_pt * scale).round() as i32
                            };
                            let mut cx = x;
                            let mut cx_pt = 0.0f64;
                            for (c, cell) in row.iter().enumerate() {
                                let w_pt = cell_width_pt(table, c, cell) as f64;
                                let pw = if tbledge_on() {
                                    cx = x + (cx_pt * scale).round() as i32;
                                    x + ((cx_pt + w_pt) * scale).round() as i32 - cx
                                } else {
                                    (w_pt * scale).round() as i32
                                };
                                // A continuation of the cell to its left owns no
                                // ink of its own: painting it would lay a second
                                // fill and a second set of rules over the spanning
                                // cell. It must not advance the cursor either --
                                // its column is already inside the spanning cell's
                                // width, and advancing again pushed d35 s29's
                                // "Week 2" six columns past its place, off the
                                // table.
                                if hmerge_on() && cell.h_merge {
                                    continue;
                                }
                                let cell_rect = RECT {
                                    left: cx,
                                    top: cy,
                                    right: cx + pw,
                                    bottom: cy + ph,
                                };
                                if let Some((rr, gg, bb)) =
                                    cell.fill_color.as_deref().and_then(parse_hex_rgb)
                                {
                                    // The corpus states cell washes as an alpha
                                    // on the fill colour (every cell of d19
                                    // slide 13 is 21355A at 15.6%), so painting
                                    // it opaque puts a navy slab over the table.
                                    match cell.fill_alpha.filter(|_| fill_alpha_on()) {
                                        Some(a) if a < 1.0 => {
                                            if a > 0.0 {
                                                alpha_fill(mem_dc, &cell_rect, (rr, gg, bb), a);
                                            }
                                        }
                                        _ => {
                                            let brush =
                                                CreateSolidBrush(COLORREF(colorref(rr, gg, bb)));
                                            FillRect(mem_dc, &cell_rect, brush);
                                            let _ = DeleteObject(brush);
                                        }
                                    }
                                }
                                // Borders, in the IR's L/R/T/B order. A side
                                // with alpha 0 is declared invisible.
                                for (side, border) in cell.borders.iter().enumerate() {
                                    let Some(b) = border else { continue };
                                    if b.alpha <= 0.0 || b.width <= 0.0 {
                                        continue;
                                    }
                                    let Some((rr, gg, bb)) = parse_hex_rgb(&b.color) else {
                                        continue;
                                    };
                                    let bpen = CreatePen(
                                        PS_SOLID,
                                        (b.width as f64 * scale).round().max(1.0) as i32,
                                        COLORREF(colorref(rr, gg, bb)),
                                    );
                                    let old_bpen = SelectObject(mem_dc, bpen);
                                    let (x0, y0, x1, y1) = match side {
                                        0 => (cx, cy, cx, cy + ph),
                                        1 => (cx + pw, cy, cx + pw, cy + ph),
                                        2 => (cx, cy, cx + pw, cy),
                                        _ => (cx, cy + ph, cx + pw, cy + ph),
                                    };
                                    let _ = MoveToEx(mem_dc, x0, y0, None);
                                    let _ = LineTo(mem_dc, x1, y1);
                                    SelectObject(mem_dc, old_bpen);
                                    let _ = DeleteObject(bpen);
                                }

                                // Text: the cell's own margins, its anchor and
                                // the paragraph's alignment, with the run's
                                // real size.
                                let left = cx + (cell.mar_l as f64 * scale).round() as i32;
                                let right = cx + pw - (cell.mar_r as f64 * scale).round() as i32;
                                let line_h = |p: &oxislides_core::ir::SlideParagraph| {
                                    let fs = p
                                        .runs
                                        .iter()
                                        .find_map(|r| r.font_size)
                                        .unwrap_or(18.0);
                                    (fs, fs as f64 * scale * 1.2)
                                };
                                // The block a centred or bottom-anchored cell is
                                // positioned by is as tall as the WRAPPED text,
                                // not as the paragraph count: since cells wrap,
                                // counting one line per paragraph makes the
                                // block look short and pushes the text down by
                                // half of every line it forgot.
                                let cell_inner_pt = (right - left).max(1) as f64 / scale;
                                let total: f64 = cell
                                    .paragraphs
                                    .iter()
                                    .map(|p| {
                                        let (fs, adv) = line_h(p);
                                        let mut n = 1usize;
                                        if cellwrap_on() && cellblock_on() {
                                            let body: String =
                                                p.runs.iter().map(|r| r.text.as_str()).collect();
                                            if !body.trim().is_empty() {
                                                let family = effective_family(
                                                    mem_dc,
                                                    &paragraph_family(
                                                        pres, sh, p, &sh.ph_levels[..], &[],
                                                    ),
                                                );
                                                let bold = p.runs.iter().any(|r| r.bold);
                                                n = gdi_wrap_lines(
                                                    mem_dc,
                                                    &body,
                                                    cell_inner_pt as f32,
                                                    cell_inner_pt as f32,
                                                    scale,
                                                    fs,
                                                    &family,
                                                    bold,
                                                    false,
                                                    Some((&p.runs[..], 0)),
                                                )
                                                .len()
                                                .max(1);
                                            }
                                        }
                                        adv * n as f64
                                    })
                                    .sum();
                                let inner_top = cy + (cell.mar_t as f64 * scale).round() as i32;
                                let inner_bot = cy + ph - (cell.mar_b as f64 * scale).round() as i32;
                                let mut cursor_y = match cell.anchor.as_deref() {
                                    Some("ctr") => {
                                        inner_top
                                            + (((inner_bot - inner_top) as f64 - total) / 2.0)
                                                .max(0.0)
                                                .round() as i32
                                    }
                                    Some("b") => {
                                        (inner_bot as f64 - total).round().max(inner_top as f64)
                                            as i32
                                    }
                                    _ => inner_top,
                                };
                                for p in &cell.paragraphs {
                                    let (fs, advance) = line_h(p);
                                    let text: String =
                                        p.runs.iter().map(|r| r.text.as_str()).collect();
                                    if text.trim().is_empty() {
                                        cursor_y += advance.round() as i32;
                                        continue;
                                    }
                                    let family = effective_family(
                                        mem_dc,
                                        &paragraph_family(
                                            pres, sh, p, &sh.ph_levels[..], &[],
                                        ),
                                    );
                                    let color = p.runs.iter().find_map(|r| r.color.clone());
                                    let bold = p.runs.iter().any(|r| r.bold);
                                    if cellwrap_on() {
                                        // A cell wraps its text like any other
                                        // text frame; this path used to draw the
                                        // paragraph as ONE line, so d25 slide 7's
                                        // 54-character body ran 297pt across a
                                        // 213.6pt column and over its neighbour.
                                        let inner_pt =
                                            (right - left).max(1) as f64 / scale;
                                        let lines = gdi_wrap_lines(
                                            mem_dc,
                                            &text,
                                            inner_pt as f32,
                                            inner_pt as f32,
                                            scale,
                                            fs,
                                            &family,
                                            bold,
                                            false,
                                            Some((&p.runs[..], 0)),
                                        );
                                        for line in &lines {
                                            // ★A trailing space is not ink and
                                            // must not be centred on. The text
                                            // FRAME path has excluded it since
                                            // S-ADVEXACT ("trailing spaces
                                            // excluded; final visible char
                                            // included"); the cell path still
                                            // measured the whole line, so a
                                            // wrapped line that ends on a space
                                            // was centred half a space too far
                                            // LEFT. d25 s7: PowerPoint's first
                                            // body line has its ink centred at
                                            // 145.28 and Oxi's at 143.84 --
                                            // 1.44pt, against half of 11pt
                                            // Arial's space (1.53).
                                            let measured = if celltrim_on() {
                                                line.trim_end()
                                            } else {
                                                line.as_str()
                                            };
                                            let lw = measure_text_width(
                                                mem_dc, measured, fs, &family, bold, scale,
                                            );
                                            let lx = match p.alignment {
                                                Some(SlideAlignment::Center) => {
                                                    left
                                                        + (((right - left) as f64 - lw) / 2.0)
                                                            .round()
                                                            as i32
                                                }
                                                Some(SlideAlignment::Right) => {
                                                    right - lw.round() as i32
                                                }
                                                _ => left,
                                            };
                                            if cellbase_on() {
                                                // A cell is a text frame: its
                                                // baseline sits the FONT's own
                                                // ascent below the line box, not
                                                // wherever GDI's TextOut top
                                                // happens to put it. Measured on
                                                // d25 s7 against PowerPoint's own
                                                // PDF (read_pptx_cellbase.py):
                                                // its three baselines imply
                                                // A = 0.9807 / 0.9761 / 0.9717 em
                                                // below each line's top, which is
                                                // the face's ascent (Arial
                                                // 0.9727), while TextOut anchors
                                                // by tmAscent.
                                                let base_pt = cursor_y as f32 / scale as f32
                                                    + font_baseline_offset_em(&family) * fs;
                                                draw_text_baseline_wiu(
                                                    mem_dc,
                                                    lx,
                                                    base_pt,
                                                    line,
                                                    fs,
                                                    &family,
                                                    color.as_deref(),
                                                    scale,
                                                    if bold { 700 } else { 400 },
                                                    false,
                                                    false,
                                                );
                                            } else {
                                                draw_text_line(
                                                    mem_dc,
                                                    lx,
                                                    cursor_y,
                                                    line,
                                                    fs,
                                                    &family,
                                                    color.as_deref(),
                                                    scale,
                                                );
                                            }
                                            cursor_y += advance.round() as i32;
                                        }
                                        continue;
                                    }
                                    let w = measure_text_width(
                                        mem_dc, &text, fs, &family, bold, scale,
                                    );
                                    let tx = match p.alignment {
                                        Some(SlideAlignment::Center) => {
                                            left + (((right - left) as f64 - w) / 2.0).round()
                                                as i32
                                        }
                                        Some(SlideAlignment::Right) => {
                                            right - w.round() as i32
                                        }
                                        _ => left,
                                    };
                                    draw_text_line(
                                        mem_dc,
                                        tx,
                                        cursor_y,
                                        &text,
                                        fs,
                                        &family,
                                        color.as_deref(),
                                        scale,
                                    );
                                    cursor_y += advance.round() as i32;
                                }
                                cx += pw;
                                cx_pt += w_pt;
                            }
                            cy += ph;
                            cy_pt += h_pt;
                        }
                    }
                    ShapeContent::Chart { chart } => {
                        // Step 2-4 + Step 5: clustered column chart. Geometry
                        // from the Word render-truth (fitz get_drawings +
                        // rawdict, chart1/2/2b/3 2026-08-06):
                        //   plot area: left = sh.x+41.4, top = sh.y+51.4
                        //     (auto-title present) or sh.y+16.0 (no title),
                        //     right = sh.x+w-11.0, bottom = sh.y+h-39.9
                        //   bars: width = pitch / (n_ser + 1.5), height =
                        //     val/max_axis x plot_h, bottom on the X axis,
                        //     cluster centred on the category centre, bars
                        //     touching within a cluster, colour = theme
                        //     accent: per-POINT for a single series
                        //     (varyColors), per-SERIES for multiple
                        //   value axis: Calibri 18pt right-aligned to
                        //     plot_left-16.64, baseline = tick_y+5.2
                        //   category names: centred on each category centre,
                        //     baseline = plot_bot+28.67
                        //   auto title: with a SINGLE series Word shows an
                        //     automatic chart title = the series name
                        //     (Calibri-Bold 21.62pt, centred on the frame,
                        //     baseline = sh.y+28.03); >=2 series -> none.
                        //     A legend is drawn only when the chart XML
                        //     declares <c:legend> (none of the 4 probes do).
                        if chart.chart_type == "pie"
                            || chart.chart_type == "doughnut"
                        {
                            // A DOUGHNUT is a pie with a hole: identical
                            // titles, slice sweep, legend and vertical
                            // geometry, so it shares this branch and differs
                            // only where marked `is_doughnut`. Word
                            // render-truth 2026-08-09 (chart_doughnut, 8 arms;
                            // 600dpi pixel scan of the ring + fitz
                            // get_drawings/rawdict for the legend):
                            //   r_in = r_out * holeSize/100 — measured
                            //     0.5010/0.5011/0.5013/0.5014 at holeSize 50
                            //     and 0.2510 at holeSize 25.
                            //   vertical geometry == the pie's within 0.16pt:
                            //     bottom sy+shh-11 (measured -11.16), top
                            //     sy+46.37 auto-title (46.44) / sy+40.7
                            //     explicit (40.80), and data labels shrink
                            //     both sides by 15.78 (measured 15.72/15.84).
                            //   data labels sit on the MIDDLE of the ring
                            //     band, i.e. at (r_in+r_out)/2 along the
                            //     slice's mid-angle: measured 74.66 against
                            //     (49.92+99.42)/2 = 74.67.
                            let is_doughnut = chart.chart_type == "doughnut";
                            // Pie chart (Word render-truth 2026-08-06, fitz
                            // get_drawings on chart_pie / chart_pie3 + the
                            // XML autoTitleDeleted census):
                            //   auto title: drawn when the series count == 1
                            //     AND the XML does NOT declare
                            //     <c:autoTitleDeleted val="1"/>. The series
                            //     name is drawn Calibri-Bold 21.62pt centred
                            //     on the frame, baseline = sh.y+28.03 (the
                            //     same rule as the bar auto title).
                            //   circle geometry (frame-size independent,
                            //     derived from 6 measured probes A-F):
                            //     center_x = sx + sw/2
                            //     bottom   = sy + shh - 11
                            //     top      = sy + 11    (untitled)
                            //              = sy + 46.37 (titled)
                            //     r        = (bottom-top)/2
                            //     center_y = (top+bottom)/2
                            //   slices: start at -90 deg (12 o'clock) and
                            //     sweep CLOCKWISE; angle = value/total*360.
                            //     colour = theme accent(i+1) per CATEGORY
                            //     (varyColors: a single series colours each
                            //     point in order). Fill only, no outline
                            //     (the Word PDF slice paths are closed
                            //     2c+2l fills with stroke=None).
                            //   legend (when <c:legend> is declared):
                            //     per-category swatch + category name,
                            //     right-aligned overlay (same geometry as
                            //     the bar legend EXCEPT legend_y0 is centred
                            //     on the CIRCLE centre, not the frame
                            //     centre - chart_pie2 slide2/3 render-truth
                            //     2026-08-06).
                            let axis_family = "Calibri";
                            let sx = sh.x as f64;
                            let sy = sh.y as f64;
                            let sw = sh.width as f64;
                            let shh = sh.height as f64;
                            let has_explicit_title = chart.explicit_title.is_some();
                            // Auto title (single series, NOT autoTitleDeleted,
                            // AND no explicit <c:title> — an explicit title
                            // suppresses the auto series-name title, same as
                            // the bar/line branches.  NOTE: unlike bar/line,
                            // the pie draws its auto title for ANY series
                            // count -- Word renders a multi-series pie with
                            // ONLY the first series (the second <c:ser> is
                            // ignored, measured chart_pie_multi 2026-08-08:
                            // slice angles == series[0].values/their sum, so
                            // 'Rev' 19.2/21.4/16.7 -> 120.6/134.4/104.9 deg)
                            // and titles it with series[0].name ('Rev').
                            let has_title_draw = !chart.auto_title_deleted
                                && !has_explicit_title;
                            if let Some(first) = chart.series.first() {
                                if has_title_draw {
                                    let tfs = 21.62f32;
                                    let lw = font_adv::line_hmtx_width_pt(
                                        &first.name,
                                        tfs,
                                        axis_family,
                                    )
                                    .unwrap_or_else(|| {
                                        first.name.chars().count() as f32 * tfs * 0.5
                                    }) as f64;
                                    let frame_cx = sx + sw / 2.0;
                                    draw_text_baseline_w(
                                        mem_dc,
                                        ((frame_cx - lw / 2.0) * scale).round() as i32,
                                        (sy + 28.03) as f32,
                                        &first.name,
                                        tfs,
                                        axis_family,
                                        None,
                                        scale,
                                        700,
                                    );
                                }
                            }
                            // EXPLICIT <c:title> text: Word draws it as Arial
                            // 18pt (regular), centred on the frame, baseline
                            // sy+24.43 (chart_title_pie render-truth
                            // 2026-08-07: origin=(194.66,96.43), same as the
                            // bar/line explicit title). It suppresses the
                            // auto series-name title.
                            if let Some(title) = &chart.explicit_title {
                                let tfs = 18.0f32;
                                let lw = font_adv::line_hmtx_width_pt(
                                    title,
                                    tfs,
                                    "Arial",
                                )
                                .unwrap_or_else(|| {
                                    title.chars().count() as f32 * tfs * 0.5
                                }) as f64;
                                let frame_cx = sx + sw / 2.0;
                                draw_text_baseline_w(
                                    mem_dc,
                                    ((frame_cx - lw / 2.0) * scale).round() as i32,
                                    (sy + 24.43) as f32,
                                    title,
                                    tfs,
                                    "Arial",
                                    None,
                                    scale,
                                    400,
                                );
                            }
                            // A NON-OVERLAY legend (<c:overlay val="0"/>)
                            // takes a band on the right and the circle is
                            // centred in what remains; an overlay legend (a
                            // bare <c:legend/>) leaves the frame centre alone.
                            // chart_doughnut 2026-08-09, ring centre vs the
                            // legend swatch x0 over three label widths
                            // (10.42 / 63.81 / 102.28pt): plot_right =
                            // swatch_x0 - 7.32 (measured 7.36/7.31/7.30), and
                            // with no legend at all the centre lands on the
                            // frame centre (269.94 vs 270.00). The swatch x0
                            // itself uses the same formula as the legend
                            // block drawn further down.
                            //
                            // A label wider than the legend's share of the
                            // frame wraps first (see `legend_label_cap`), and
                            // the block then takes the WRAPPED width.
                            let legend_fs = 18.0f32;
                            let legend_cap = legend_label_cap(sw);
                            let legend_lines: Vec<Vec<String>> = chart
                                .categories
                                .iter()
                                .map(|name| {
                                    wrap_legend_label(
                                        name,
                                        legend_fs,
                                        axis_family,
                                        legend_cap,
                                    )
                                })
                                .collect();
                            let legend_label_w = legend_lines
                                .iter()
                                .flat_map(|ls| ls.iter())
                                .map(|s| {
                                    font_adv::line_hmtx_width_pt(
                                        s,
                                        legend_fs,
                                        axis_family,
                                    )
                                    .unwrap_or_else(|| {
                                        s.chars().count() as f32 * legend_fs * 0.5
                                    }) as f64
                                })
                                .fold(0.0f64, f64::max);
                            let legend_max_lines = legend_lines
                                .iter()
                                .map(|ls| ls.len())
                                .max()
                                .unwrap_or(1);
                            let banded_right = if chart.has_legend
                                && !chart.legend_overlay
                            {
                                let swatch_x0 = (sx + sw)
                                    - 10.0
                                    - legend_label_w
                                    - 4.62
                                    - 9.89;
                                Some(swatch_x0 - 7.32)
                            } else {
                                None
                            };
                            let circle_cx = match banded_right {
                                Some(right) => (sx + right) / 2.0,
                                None => sx + sw / 2.0,
                            };
                            // Data labels (c:dLbls) may shrink the pie circle
                            // (see the OUTSIDE_END rule below) — top/bottom
                            // must be mutable.
                            let mut circle_bot = sy + shh - 11.0;
                            // Pie circle top: untitled sy+11 / auto-title
                            // sy+46.37 / EXPLICIT title sy+40.7. The explicit
                            // value = bar explicit plot_top (45.69) − 5.0,
                            // consistent with the other two (16−5 = 11 and
                            // 51.4−5 = 46.37 — the pie circle sits 5pt above
                            // the bar plot area). chart_title_pie render-truth
                            // 2026-08-07: circle top = 112.70 = sy+40.7.
                            let mut circle_top = sy + if has_explicit_title {
                                40.7
                            } else if has_title_draw {
                                46.37
                            } else {
                                11.0
                            };
                            // PIE data labels (c:dLbls): present when a
                            // <c:dLbls> with showVal=1 is declared. CENTER
                            // labels leave the circle unchanged; OUTSIDE_END
                            // (the pie default when no <c:dLblPos> is written,
                            // i.e. datalabel_position == "") shrinks the
                            // circle inward by 15.78pt on BOTH sides, keeping
                            // circle_cy fixed — chart_datalabel_pie P1/P2
                            // render-truth 2026-08-07: r 115.31 -> 99.53,
                            // circle_top 118.37 -> 134.15, circle_bot
                            // 349.0 -> 333.21, circle_cy unchanged at 233.68.
                            let has_pie_labels =
                                chart.has_data_labels && chart.show_val;
                            let pie_labels_outside = has_pie_labels
                                && chart.datalabel_position != "ctr";
                            if pie_labels_outside {
                                circle_top += 15.78;
                                circle_bot -= 15.78;
                            }
                            // The circle is bound by whichever of the plot
                            // area's height and width is smaller (in every
                            // measured arm the height binds).
                            let r = match banded_right {
                                Some(right) => ((circle_bot - circle_top) / 2.0)
                                    .min((right - sx) / 2.0),
                                None => (circle_bot - circle_top) / 2.0,
                            };
                            let circle_cy = (circle_top + circle_bot) / 2.0;
                            // Doughnut hole: r_in = r_out * holeSize/100.
                            let r_in = if is_doughnut {
                                r * (chart.hole_size / 100.0).clamp(0.0, 0.95)
                            } else {
                                0.0
                            };
                            let bx0 = ((circle_cx - r) * scale).round() as i32;
                            let by0 = ((circle_cy - r) * scale).round() as i32;
                            let bx1 = ((circle_cx + r) * scale).round() as i32;
                            let by1 = ((circle_cy + r) * scale).round() as i32;
                            // total = the FIRST series' values only.  Word
                            // renders a multi-series pie with ONLY series[0]
                            // (chart_pie_multi 2026-08-08: slice angles ==
                            // series[0].values/their sum, the 2nd <c:ser> is
                            // ignored), so summing across ALL series would
                            // shrink every slice.
                            //
                            // A multi-series DOUGHNUT is the exception: Word
                            // draws one CONCENTRIC RING PER SERIES, splitting
                            // the [r_in, r] annulus into n equal-width bands
                            // with series 0 INNERMOST.  chart_doughnut_resid
                            // S2 (2026-08-09, 600dpi scan along +x): the band
                            // edges go 66.72 / 99.84 | 99.96 / 133.08 for one
                            // hole and two series -- two equal 33.12 bands --
                            // and a colour sample at 122 deg (where the two
                            // series disagree on which category owns the ray)
                            // puts series1 (West/accent2) INSIDE and series2
                            // (East/accent1) outside.  Each ring's angles come
                            // from its OWN series total.
                            let ring_count = if is_doughnut {
                                chart.series.len().max(1)
                            } else {
                                1
                            };
                            let ring_w = (r - r_in) / ring_count as f64;
                            let _ =
                                SelectObject(mem_dc, GetStockObject(NULL_PEN));
                            for (si, ser) in
                                chart.series.iter().enumerate().take(ring_count)
                            {
                                let ring_in = r_in + ring_w * si as f64;
                                let ring_out = r_in + ring_w * (si + 1) as f64;
                                let total: f64 =
                                    ser.values.iter().copied().sum();
                                let mut start_deg = -90.0f64;
                                for (ci, v) in ser.values.iter().enumerate() {
                                    if total <= 0.0 || *v <= 0.0 {
                                        continue;
                                    }
                                    let sweep = v / total * 360.0;
                                    let end_deg = start_deg + sweep;
                                    let to_rad = |deg: f64| {
                                        deg * std::f64::consts::PI / 180.0
                                    };
                                    let p1 = (
                                        circle_cx
                                            + ring_out
                                                * (to_rad(start_deg)).cos(),
                                        circle_cy
                                            + ring_out
                                                * (to_rad(start_deg)).sin(),
                                    );
                                    let p2 = (
                                        circle_cx
                                            + ring_out * (to_rad(end_deg)).cos(),
                                        circle_cy
                                            + ring_out * (to_rad(end_deg)).sin(),
                                    );
                                    let col_hex = pres
                                        .theme_colors
                                        .get(&format!("accent{}", ci + 1))
                                        .map(|s| s.as_str())
                                        .or_else(|| DEFAULT_ACCENT.get(ci).copied());
                                    if let Some(rgb) =
                                        col_hex.and_then(parse_hex_rgb)
                                    {
                                        let brush = CreateSolidBrush(COLORREF(
                                            colorref(rgb.0, rgb.1, rgb.2),
                                        ));
                                        let old_brush =
                                            SelectObject(mem_dc, brush);
                                        if is_doughnut {
                                            // GDI has no annular-sector
                                            // primitive, so flatten the band
                                            // into a polygon: the outer arc
                                            // start->end, then the inner arc
                                            // back. 1-degree steps are well
                                            // under a pixel at 150dpi and the
                                            // fill stays background-agnostic
                                            // (punching the hole with a
                                            // background-coloured circle
                                            // would not be).
                                            let steps = (sweep.abs().ceil()
                                                as usize)
                                                .clamp(2, 720);
                                            let mut pts: Vec<POINT> =
                                                Vec::with_capacity(
                                                    (steps + 1) * 2,
                                                );
                                            let at = |rad: f64, deg: f64| {
                                                POINT {
                                                    x: ((circle_cx
                                                        + rad
                                                            * to_rad(deg).cos())
                                                        * scale)
                                                        .round()
                                                        as i32,
                                                    y: ((circle_cy
                                                        + rad
                                                            * to_rad(deg).sin())
                                                        * scale)
                                                        .round()
                                                        as i32,
                                                }
                                            };
                                            for i in 0..=steps {
                                                let t = i as f64
                                                    / steps as f64;
                                                pts.push(at(
                                                    ring_out,
                                                    start_deg + sweep * t,
                                                ));
                                            }
                                            for i in (0..=steps).rev() {
                                                let t = i as f64
                                                    / steps as f64;
                                                pts.push(at(
                                                    ring_in,
                                                    start_deg + sweep * t,
                                                ));
                                            }
                                            let old_pen = SelectObject(
                                                mem_dc,
                                                GetStockObject(NULL_PEN),
                                            );
                                            let _ = Polygon(mem_dc, &pts);
                                            SelectObject(mem_dc, old_pen);
                                        } else {
                                            // GDI Pie() sweeps COUNTER-CLOCKWISE from
                                            // (xr1,yr1) to (xr2,yr2). Word's slices
                                            // sweep CLOCKWISE from -90 deg, so pass
                                            // the endpoints in reverse order (p2 =
                                            // the clockwise END of the slice).
                                            let _ = Pie(
                                                mem_dc,
                                                bx0,
                                                by0,
                                                bx1,
                                                by1,
                                                (p2.0 * scale).round() as i32,
                                                (p2.1 * scale).round() as i32,
                                                (p1.0 * scale).round() as i32,
                                                (p1.1 * scale).round() as i32,
                                            );
                                        }
                                        SelectObject(mem_dc, old_brush);
                                        let _ = DeleteObject(brush);
                                    }
                                    start_deg = end_deg;
                                }
                            }
                            // PIE data labels (c:dLbls): Word renders each
                            // point's value in Calibri 18pt black on the
                            // slice's mid-angle ray (chart_datalabel_pie
                            // render-truth 2026-08-07):
                            //   CENTER:      anchor at 0.5*r along the
                            //                mid-angle (no circle shrink).
                            //   OUTSIDE_END: anchor at 0.78*r (shrunk
                            //                circle) along the mid-angle.
                            //   baseline = anchor.y + 6.2 and the text is
                            //   horizontally centred at anchor.x — the same
                            //   vertical rule as the line/bar data labels.
                            //   Format: numFmt "0.0%" -> value*100 one-decimal
                            //   + "%"; otherwise Word prints the value with
                            //   its decimals ('19.2' etc.), so keep the raw
                            //   number.
                            if has_pie_labels {
                                let dlfs = 18.0f32;
                                let num_fmt = chart.number_format.clone();
                                let format_label = |v: f64| -> String {
                                    if num_fmt == "0.0%" {
                                        format!("{:.1}%", v * 100.0)
                                    } else {
                                        format!("{}", v)
                                    }
                                };
                                // One label pass per RING (a pie has exactly
                                // one).  Multi-series doughnut labels are not
                                // measured; each ring reuses the measured
                                // single-ring rule on its own band.
                                for (si, ser) in chart
                                    .series
                                    .iter()
                                    .enumerate()
                                    .take(ring_count)
                                {
                                    let ring_in = r_in + ring_w * si as f64;
                                    let ring_out =
                                        r_in + ring_w * (si + 1) as f64;
                                    let total: f64 =
                                        ser.values.iter().copied().sum();
                                    let mut lab_deg = -90.0f64;
                                    for v in ser.values.iter() {
                                        if total <= 0.0 || *v <= 0.0 {
                                            continue;
                                        }
                                        let sweep = v / total * 360.0;
                                        let end_deg = lab_deg + sweep;
                                        let mid_deg = (lab_deg + end_deg) / 2.0;
                                        let to_rad = |deg: f64| {
                                            deg * std::f64::consts::PI / 180.0
                                        };
                                        let text = format_label(*v);
                                        let lw = font_adv::line_hmtx_width_pt(
                                            &text,
                                            dlfs,
                                            axis_family,
                                        )
                                        .unwrap_or_else(|| {
                                            text.chars().count() as f32
                                                * dlfs
                                                * 0.5
                                        }) as f64;
                                        let label_r = if is_doughnut {
                                            // A doughnut has no room outside
                                            // its band, so Word centres the
                                            // label ON the band: measured
                                            // 74.66 against (r_in+r_out)/2 =
                                            // (49.92+99.42)/2 = 74.67 on all
                                            // three slices of chart_doughnut
                                            // slide 3 (the ring still shrinks
                                            // by 15.78 per side as a pie's
                                            // would).
                                            (ring_in + ring_out) / 2.0
                                        } else if pie_labels_outside {
                                            // OUTSIDE_END: Word places the
                                            // label CENTRE at 0.78 * the
                                            // PRE-shrink radius minus
                                            // 0.37 * label width (width-ramp
                                            // probe 2026-08-08: constant-outer-
                                            // edge model disproven; this linear
                                            // fit matches all widths within
                                            // ~1pt). The 15.78pt shrink moved
                                            // circle_top/bot inward, so the
                                            // pre-shrink radius = r + 15.78.
                                            (r + 15.78) * 0.78 - 0.37 * lw
                                        } else {
                                            // CENTER: anchor at 0.5*r (no
                                            // circle shrink).
                                            r * 0.5
                                        };
                                        let anchor = (
                                            circle_cx
                                                + label_r * to_rad(mid_deg).cos(),
                                            circle_cy
                                                + label_r * to_rad(mid_deg).sin(),
                                        );
                                        let lx = anchor.0 - lw / 2.0;
                                        draw_text_baseline(
                                            mem_dc,
                                            (lx * scale).round() as i32,
                                            (anchor.1 + 6.2) as f32,
                                            &text,
                                            dlfs,
                                            axis_family,
                                            None,
                                            scale,
                                        );
                                        lab_deg = end_deg;
                                    }
                                }
                            }

                            // Legend (when <c:legend> declared): per-category
                            // swatch + category name, right-aligned overlay,
                            // vertically centred on the CIRCLE centre.
                            if chart.has_legend {
                                let lfs = legend_fs;
                                let n_cat = chart.categories.len().max(1);
                                let max_label_w = legend_label_w;
                                let swatch_w = 9.89f64;
                                let gap = 4.62f64;
                                // Every row grows by the extra text lines of
                                // the TALLEST entry: chart_doughnut_resid
                                // measures a uniform 27.75 for all-single-line
                                // legends, 49.51 when one entry wraps to two
                                // lines and 93.02 when one wraps to four, i.e.
                                // 27.75 + (lines-1) * 21.76.  Text lines
                                // inside an entry sit 21.99 apart.
                                let text_line_pitch = 21.99f64;
                                let row_pitch = 27.75
                                    + (legend_max_lines as f64 - 1.0) * 21.76;
                                let legend_right = (sx + sw) - 10.0;
                                let swatch_x1 = legend_right - max_label_w - gap;
                                let swatch_x0 = swatch_x1 - swatch_w;
                                let label_x0 = swatch_x1 + gap;
                                // Block height + placement, from the 8-arm
                                // chart_legendvert sweep (L = 1..5 x n = 2..4):
                                // the block is n * row_pitch tall, is centred
                                // on the CIRCLE centre, and the first swatch
                                // sits 8.97 below the block top.  WHICH entry
                                // wraps is irrelevant -- first / middle / last /
                                // all-wrap render byte-identically (B1/B3/B4/B7
                                // all put the swatches at 168.35/217.86/267.37);
                                // only the tallest entry's line count matters.
                                // n moves the first swatch by exactly
                                // row_pitch/2 per entry (193.10 / 168.35 /
                                // 143.59 for n = 2/3/4).
                                //
                                // When the block would leave the frame Word
                                // re-measures it with the LAST row shrunk to
                                // its own lines (chart_doughnut_resid S8,
                                // L=4 n=3: natural 279.06 overflows, tight
                                // 213.79 fits -> first swatch 135.72), and if
                                // that still does not fit it DROPS trailing
                                // entries (B8, L=5 n=3: renders only two
                                // swatches, at 127.84 = the n=2 natural block).
                                let swatch_off = 8.97f64;
                                let mut n_shown = n_cat;
                                let mut legend_h =
                                    n_shown as f64 * row_pitch;
                                loop {
                                    let top = circle_cy - legend_h / 2.0;
                                    if top >= sy - 0.01
                                        && top + legend_h <= sy + shh + 0.01
                                    {
                                        break;
                                    }
                                    let last_lines = legend_lines
                                        .get(n_shown.saturating_sub(1))
                                        .map(|l| l.len().max(1))
                                        .unwrap_or(1);
                                    let tight = (n_shown as f64 - 1.0)
                                        * row_pitch
                                        + (last_lines as f64 * 21.76 + 5.99);
                                    let ttop = circle_cy - tight / 2.0;
                                    if tight < legend_h
                                        && ttop >= sy - 0.01
                                        && ttop + tight <= sy + shh + 0.01
                                    {
                                        legend_h = tight;
                                        break;
                                    }
                                    if n_shown <= 1 {
                                        break;
                                    }
                                    n_shown -= 1;
                                    legend_h = n_shown as f64 * row_pitch;
                                }
                                let legend_y0 =
                                    circle_cy - legend_h / 2.0 + swatch_off;
                                for (ci, name) in chart
                                    .categories
                                    .iter()
                                    .enumerate()
                                    .take(n_shown)
                                {
                                    let sw_y =
                                        legend_y0 + ci as f64 * row_pitch;
                                    let col_hex = pres
                                        .theme_colors
                                        .get(&format!("accent{}", ci + 1))
                                        .map(|s| s.as_str())
                                        .or_else(|| {
                                            DEFAULT_ACCENT.get(ci).copied()
                                        });
                                    if let Some(rgb) =
                                        col_hex.and_then(parse_hex_rgb)
                                    {
                                        let brush = CreateSolidBrush(COLORREF(
                                            colorref(rgb.0, rgb.1, rgb.2),
                                        ));
                                        let old_brush =
                                            SelectObject(mem_dc, brush);
                                        let r = RECT {
                                            left: (swatch_x0 * scale).round() as i32,
                                            top: (sw_y * scale).round() as i32,
                                            right: (swatch_x1 * scale).round() as i32,
                                            bottom: ((sw_y + swatch_w) * scale).round() as i32,
                                        };
                                        let _ = FillRect(mem_dc, &r, brush);
                                        SelectObject(mem_dc, old_brush);
                                        let _ = DeleteObject(brush);
                                    }
                                    let label_baseline =
                                        sw_y + swatch_w + 0.28;
                                    let lines = legend_lines
                                        .get(ci)
                                        .cloned()
                                        .unwrap_or_else(|| {
                                            vec![name.to_string()]
                                        });
                                    for (li, line) in lines.iter().enumerate() {
                                        draw_text_baseline(
                                            mem_dc,
                                            (label_x0 * scale).round() as i32,
                                            (label_baseline
                                                + li as f64 * text_line_pitch)
                                                as f32,
                                            line,
                                            lfs,
                                            axis_family,
                                            None,
                                            scale,
                                        );
                                    }
                                }
                            }
                        } else if chart.chart_type == "area"
                            && std::env::var("OXI_AREA_DISABLE").is_err()
                        {
                        // ================= AREA chart (<c:areaChart>) =================
                        // Word render-truth measured 2026-08-09 with fitz
                        // get_drawings/rawdict on pipeline_data/pptx_probes/
                        // chart_area (8 arms: standard / stacked /
                        // percentStacked x 1..3 series x auto / no title) and
                        // chart_area_leg (6 arms isolating the legend band,
                        // the label width and an explicit title).
                        //
                        //   plot_left  = sx + 6.50 + widest VALUE label + 16.70
                        //     -- the same axis inset the horizontal bar uses
                        //     for its category labels.  113.45 with "0".."25"
                        //     labels, 135.44 with "0%".."100%": both EXACT,
                        //     and it reproduces the line chart's hardcoded
                        //     sx+41.4 and the percentStacked column's sx+63.44.
                        //   plot_top   = sy + 45.69 explicit title / sy + 51.40
                        //     auto title / sy + 16.00 none   (== the line chart)
                        //   plot_bot   = sy + shh - 39.90
                        //   band_right = legend ? swatch_x0 - 18.15
                        //                      : sx + sw - 11.0
                        //   plot_right = band_right - w(LAST category label)/2
                        //     ★ area categories sit AT the plot edges, so the
                        //     outermost label overhangs by half its width.
                        //     chart_area_leg G1/G2/G3/G6 (legend label widths
                        //     18.23 / 32.63 / 77.47 / 90.99) and G4 (no legend)
                        //     all reproduce plot_right to 0.04pt.
                        //   ★ data points sit at the CATEGORY BOUNDARIES, not
                        //     at band centres like the line/column charts:
                        //     x_i = plot_left + i * plot_w/(n_cat-1)
                        //     (113.45 / 202.92 / 292.44 / 381.98 for n=4).
                        //   value scale: standard -> nice_axis_max(max value)
                        //     in 5 steps; stacked -> nice_axis_max(max category
                        //     SUM) in (max/5).round() steps; percentStacked ->
                        //     0..100% in 10 steps  (== the column chart).
                        //   draw order (PDF): frame, gridlines, area fills,
                        //     axis lines + ticks, legend.  A gridline crossing
                        //     a filled band is covered by the accent colour,
                        //     so the gridlines go down BEFORE the fills.
                        //   fills: standard -> every series is its own polygon
                        //     closed down to the X axis, painted in series
                        //     order (later series cover earlier ones); stacked
                        //     -> the band between the running cumulative
                        //     curves.  The paths are fill-only (no outline).
                        //   category ticks sit at the DATA x (n ticks, not
                        //     n+1); labels are centred there with baseline
                        //     plot_bot + 28.67.
                        //   legend rows are the SERIES, and ★ a STACKED area
                        //     REVERSES them (chart_area S4/S5/S6 put Ser2 on
                        //     top; the standard arms keep Ser1 on top).  The
                        //     block is n_ser * row_pitch centred on
                        //     sy + shh/2 + the title offset (0 / 14.85 with an
                        //     explicit title / 17.68 with the auto title --
                        //     S7 pins the no-title case at 0, which the line
                        //     chart could not separate because its auto title
                        //     only ever fires at n==1), with the first swatch
                        //     8.97 below the block top: the doughnut rule.
                        let axis_family = "Calibri";
                        let axis_fs = 18.0f32;
                        let sx = sh.x as f64;
                        let sy = sh.y as f64;
                        let sw = sh.width as f64;
                        let shh = sh.height as f64;
                        let text_w = |t: &str, fs: f32| -> f64 {
                            font_adv::line_hmtx_width_pt(t, fs, axis_family)
                                .unwrap_or_else(|| {
                                    t.chars().count() as f32 * fs * 0.5
                                }) as f64
                        };
                        let n_cat = chart.categories.len().max(1);
                        let n_ser = chart.series.len().max(1);
                        let is_100pct = chart.grouping == "percentStacked";
                        let is_stacked = chart.grouping == "stacked" || is_100pct;
                        let has_explicit_title = chart.explicit_title.is_some();
                        let has_auto_title = chart.series.len() == 1
                            && !chart.auto_title_deleted
                            && !has_explicit_title;

                        // ---- value axis ----
                        let raw_max = if is_stacked {
                            (0..n_cat)
                                .map(|ci| {
                                    chart
                                        .series
                                        .iter()
                                        .map(|s| {
                                            s.values.get(ci).copied().unwrap_or(0.0)
                                        })
                                        .sum::<f64>()
                                })
                                .fold(0.0f64, f64::max)
                        } else {
                            chart
                                .series
                                .iter()
                                .flat_map(|s| s.values.iter().copied())
                                .fold(0.0f64, f64::max)
                        };
                        // NEGATIVE data (chart_negative N6, 2026-08-10): the
                        // axis spans zero, the area closes on the ZERO line
                        // and the category names hang off it, so the 39.9pt
                        // bottom band collapses to the plain 16.0 margin.
                        let raw_min = if is_stacked {
                            (0..n_cat)
                                .map(|ci| {
                                    chart
                                        .series
                                        .iter()
                                        .map(|s| {
                                            s.values.get(ci).copied().unwrap_or(0.0)
                                        })
                                        .filter(|v| *v < 0.0)
                                        .sum::<f64>()
                                })
                                .fold(0.0f64, f64::min)
                        } else {
                            chart
                                .series
                                .iter()
                                .flat_map(|s| s.values.iter().copied())
                                .fold(0.0f64, f64::min)
                        };
                        let has_neg = raw_min < 0.0;
                        let plot_bot_pre =
                            sy + shh - if has_neg { 16.0 } else { 39.9 };
                        let plot_top_pre = if has_explicit_title {
                            sy + 45.69
                        } else if has_auto_title {
                            sy + 51.40
                        } else {
                            sy + 16.0
                        };
                        let (axis_min, max_axis, axis_steps) = if is_100pct {
                            (0.0, 100.0, 10usize)
                        } else if has_neg {
                            nice_axis_range(
                                raw_min,
                                raw_max,
                                plot_bot_pre - plot_top_pre,
                                VERT_MIN_SPACING,
                            )
                        } else {
                            let m = nice_axis_max(raw_max);
                            let steps = if is_stacked {
                                ((m / 5.0).round() as usize).max(1)
                            } else {
                                5usize
                            };
                            (0.0, m, steps)
                        };
                        let axis_span = (max_axis - axis_min).max(1e-9);
                        let axis_label = |i: usize| -> String {
                            let v = axis_min + axis_span * i as f64 / axis_steps as f64;
                            if is_100pct {
                                format!("{:.0}%", v)
                            } else {
                                format!("{}", v.round() as i64)
                            }
                        };
                        let val_label_w = (0..=axis_steps)
                            .map(|i| text_w(&axis_label(i), axis_fs))
                            .fold(0.0f64, f64::max);

                        let plot_left = sx + 6.50 + val_label_w + 16.70;
                        let plot_top = plot_top_pre;
                        let plot_bot = plot_bot_pre;
                        let plot_h = plot_bot - plot_top;
                        let val_y =
                            |v: f64| plot_bot - ((v - axis_min) / axis_span) * plot_h;
                        let zero_y = val_y(0.0);

                        // ---- legend block geometry (rows = series) ----
                        let legend_fs = 18.0f32;
                        let legend_cap = legend_label_cap(sw);
                        let legend_lines: Vec<Vec<String>> = chart
                            .series
                            .iter()
                            .map(|s| {
                                wrap_legend_label(
                                    &s.name,
                                    legend_fs,
                                    axis_family,
                                    legend_cap,
                                )
                            })
                            .collect();
                        let legend_label_w = legend_lines
                            .iter()
                            .flat_map(|ls| ls.iter())
                            .map(|s| text_w(s, legend_fs))
                            .fold(0.0f64, f64::max);
                        let legend_max_lines = legend_lines
                            .iter()
                            .map(|ls| ls.len())
                            .max()
                            .unwrap_or(1);
                        let swatch_w = 9.89f64;
                        let legend_gap = 4.62f64;
                        let legend_right = (sx + sw) - 10.0;
                        let swatch_x1 = legend_right - legend_label_w - legend_gap;
                        let swatch_x0 = swatch_x1 - swatch_w;
                        let label_x0 = swatch_x1 + legend_gap;
                        // A legend that declares <c:overlay val="0"/> takes a
                        // band off the plot; a bare <c:legend/> overlays it.
                        // chart_area_dlbls D6 (band, plot_right 381.98) vs D7
                        // (bare, plot_right 446.38 with the swatches sitting
                        // INSIDE the plot at x 410.79) -- the same
                        // discriminator pie/doughnut derived.
                        let band_right = if chart.has_legend && !chart.legend_overlay
                        {
                            swatch_x0 - 18.15
                        } else {
                            sx + sw - 11.0
                        };
                        let last_cat_w = chart
                            .categories
                            .last()
                            .map(|c| text_w(c, axis_fs))
                            .unwrap_or(0.0);
                        let plot_right = band_right - last_cat_w / 2.0;
                        let plot_w = (plot_right - plot_left).max(1.0);
                        let step_x = if n_cat > 1 {
                            plot_w / (n_cat - 1) as f64
                        } else {
                            0.0
                        };
                        let cat_x = |ci: usize| plot_left + step_x * ci as f64;

                        // ---- value-axis labels (right-aligned to
                        // plot_left-16.70, baseline tick_y+5.24) ----
                        for i in 0..=axis_steps {
                            let tick_y =
                                plot_bot - plot_h * i as f64 / axis_steps as f64;
                            let label = axis_label(i);
                            let lx = plot_left - 16.70 - text_w(&label, axis_fs);
                            draw_text_baseline(
                                mem_dc,
                                (lx * scale).round() as i32,
                                (tick_y + 5.24) as f32,
                                &label,
                                axis_fs,
                                axis_family,
                                None,
                                scale,
                            );
                        }

                        // ---- major gridlines (behind the fills) ----
                        {
                            let grid_pen =
                                CreatePen(PS_SOLID, 2, COLORREF(colorref(0, 0, 0)));
                            let old_pen = SelectObject(mem_dc, grid_pen);
                            let _ = SelectObject(mem_dc, GetStockObject(NULL_BRUSH));
                            let gl = (plot_left * scale).round() as i32;
                            let gr = (plot_right * scale).round() as i32;
                            // With negatives Word also rules the bottom tick
                            // (the axis line has moved up to zero).
                            let grid_from = if has_neg { 0 } else { 1 };
                            for i in grid_from..=axis_steps {
                                let gy = ((plot_bot
                                    - plot_h * i as f64 / axis_steps as f64)
                                    * scale)
                                    .round() as i32;
                                let _ = MoveToEx(mem_dc, gl, gy, None);
                                let _ = LineTo(mem_dc, gr, gy);
                            }
                            SelectObject(mem_dc, old_pen);
                            let _ = DeleteObject(grid_pen);
                        }

                        // ---- area fills ----
                        // Data labels (<c:dLbls> + <c:showVal val="1"/>) sit at
                        // the vertical CENTRE of the band that series paints,
                        // horizontally centred on the data point, baseline at
                        // centre + 6.22.  Word render-truth chart_area_dlbls
                        // D1-D7 (24 labels, every one within 0.04pt):
                        //   standard  band = [point, plot_bot]
                        //   stacked   band = [own cumulative, previous
                        //                     cumulative]  (D3 '12.3' baseline
                        //                     162.96 = (121.06+192.44)/2+6.21)
                        //   ★ 100% stacked still prints the RAW value, not the
                        //     percentage (D4 shows '16.7'/'8.5', not '56%').
                        let mut dlbl_slots: Vec<(f64, f64, f64)> = Vec::new();
                        let mut cum = vec![0.0f64; n_cat];
                        for (si, ser) in chart.series.iter().enumerate() {
                            let col_hex = pres
                                .theme_colors
                                .get(&format!("accent{}", si + 1))
                                .map(|s| s.as_str())
                                .or_else(|| DEFAULT_ACCENT.get(si % 6).copied());
                            let rgb = match col_hex.and_then(parse_hex_rgb) {
                                Some(v) => v,
                                None => continue,
                            };
                            // Upper boundary (cumulative when stacked) and the
                            // lower boundary the polygon closes back along.
                            let mut upper: Vec<f64> = Vec::with_capacity(n_cat);
                            let mut lower: Vec<f64> = Vec::with_capacity(n_cat);
                            for ci in 0..n_cat {
                                let v = ser.values.get(ci).copied().unwrap_or(0.0);
                                if is_stacked {
                                    let base = cum[ci];
                                    let total: f64 = if is_100pct {
                                        chart
                                            .series
                                            .iter()
                                            .map(|s| {
                                                s.values
                                                    .get(ci)
                                                    .copied()
                                                    .unwrap_or(0.0)
                                            })
                                            .sum()
                                    } else {
                                        1.0
                                    };
                                    let vv = if is_100pct {
                                        if total > 0.0 {
                                            v / total * 100.0
                                        } else {
                                            0.0
                                        }
                                    } else {
                                        v
                                    };
                                    lower.push(base);
                                    cum[ci] = base + vv;
                                    upper.push(cum[ci]);
                                } else {
                                    lower.push(0.0);
                                    upper.push(v);
                                }
                            }
                            if chart.has_data_labels && chart.show_val {
                                for ci in 0..n_cat {
                                    let v =
                                        ser.values.get(ci).copied().unwrap_or(0.0);
                                    let mid = (val_y(upper[ci])
                                        + val_y(lower[ci]))
                                        / 2.0;
                                    dlbl_slots.push((cat_x(ci), mid, v));
                                }
                            }
                            let mut pts: Vec<POINT> =
                                Vec::with_capacity(n_cat * 2);
                            for ci in 0..n_cat {
                                pts.push(POINT {
                                    x: (cat_x(ci) * scale).round() as i32,
                                    y: (val_y(upper[ci]) * scale).round() as i32,
                                });
                            }
                            for ci in (0..n_cat).rev() {
                                pts.push(POINT {
                                    x: (cat_x(ci) * scale).round() as i32,
                                    y: (val_y(lower[ci]) * scale).round() as i32,
                                });
                            }
                            if pts.len() >= 3 {
                                let brush = CreateSolidBrush(COLORREF(colorref(
                                    rgb.0, rgb.1, rgb.2,
                                )));
                                let old_brush = SelectObject(mem_dc, brush);
                                let old_pen =
                                    SelectObject(mem_dc, GetStockObject(NULL_PEN));
                                let _ = Polygon(mem_dc, &pts);
                                SelectObject(mem_dc, old_pen);
                                SelectObject(mem_dc, old_brush);
                                let _ = DeleteObject(brush);
                            }
                        }

                        // ---- axis lines + ticks (on top of the fills) ----
                        {
                            let axis_pen =
                                CreatePen(PS_SOLID, 2, COLORREF(colorref(0, 0, 0)));
                            let old_pen = SelectObject(mem_dc, axis_pen);
                            let _ = SelectObject(mem_dc, GetStockObject(NULL_BRUSH));
                            let pl = (plot_left * scale).round() as i32;
                            let pr = (plot_right * scale).round() as i32;
                            let pb = (plot_bot * scale).round() as i32;
                            let pz = (zero_y * scale).round() as i32;
                            let pt_i = (plot_top * scale).round() as i32;
                            let _ = MoveToEx(mem_dc, pl, pt_i, None);
                            let _ = LineTo(mem_dc, pl, pb);
                            let _ = MoveToEx(mem_dc, pl, pz, None);
                            let _ = LineTo(mem_dc, pr, pz);
                            for i in 0..=axis_steps {
                                let ty = ((plot_bot
                                    - plot_h * i as f64 / axis_steps as f64)
                                    * scale)
                                    .round() as i32;
                                let _ = MoveToEx(
                                    mem_dc,
                                    ((plot_left - 5.71) * scale).round() as i32,
                                    ty,
                                    None,
                                );
                                let _ = LineTo(mem_dc, pl, ty);
                            }
                            for ci in 0..n_cat {
                                let tx = (cat_x(ci) * scale).round() as i32;
                                let _ = MoveToEx(mem_dc, tx, pz, None);
                                let _ = LineTo(
                                    mem_dc,
                                    tx,
                                    ((zero_y + 5.71) * scale).round() as i32,
                                );
                            }
                            SelectObject(mem_dc, old_pen);
                            let _ = DeleteObject(axis_pen);
                        }

                        // ---- data labels (band centre, on top of the fills) ----
                        if !dlbl_slots.is_empty() {
                            let num_fmt = chart.number_format.clone();
                            for (lx_c, mid, v) in &dlbl_slots {
                                let text = if num_fmt == "0.0%" {
                                    format!("{:.1}%", v * 100.0)
                                } else if num_fmt == "0%" {
                                    format!("{}%", (v * 100.0).round() as i64)
                                } else {
                                    // Word prints the raw value with no trailing
                                    // zero: 22.0 -> "22", 21.4 -> "21.4".
                                    format!("{}", v)
                                };
                                let lw = font_adv::line_hmtx_width_pt(
                                    &text, axis_fs, axis_family,
                                )
                                .unwrap_or_else(|| {
                                    text.chars().count() as f32 * axis_fs * 0.5
                                }) as f64;
                                draw_text_baseline(
                                    mem_dc,
                                    ((lx_c - lw / 2.0) * scale).round() as i32,
                                    (mid + 6.22) as f32,
                                    &text,
                                    axis_fs,
                                    axis_family,
                                    None,
                                    scale,
                                );
                            }
                        }

                        // ---- category labels (centred on the data x) ----
                        for (ci, name) in chart.categories.iter().enumerate() {
                            let lw = text_w(name, axis_fs);
                            draw_text_baseline(
                                mem_dc,
                                ((cat_x(ci) - lw / 2.0) * scale).round() as i32,
                                (zero_y + 28.67) as f32,
                                name,
                                axis_fs,
                                axis_family,
                                None,
                                scale,
                            );
                        }

                        // ---- titles ----
                        if let Some(title) = &chart.explicit_title {
                            let tfs = 18.0f32;
                            let lw = font_adv::line_hmtx_width_pt(title, tfs, "Arial")
                                .unwrap_or_else(|| {
                                    title.chars().count() as f32 * tfs * 0.5
                                }) as f64;
                            draw_text_baseline_w(
                                mem_dc,
                                ((sx + sw / 2.0 - lw / 2.0) * scale).round() as i32,
                                (sy + 24.43) as f32,
                                title,
                                tfs,
                                "Arial",
                                None,
                                scale,
                                400,
                            );
                        } else if has_auto_title {
                            let first = &chart.series[0];
                            let tfs = 21.62f32;
                            let lw = text_w(&first.name, tfs);
                            draw_text_baseline_w(
                                mem_dc,
                                ((sx + sw / 2.0 - lw / 2.0) * scale).round() as i32,
                                (sy + 28.03) as f32,
                                &first.name,
                                tfs,
                                axis_family,
                                None,
                                scale,
                                700,
                            );
                        }

                        // ---- legend ----
                        if chart.has_legend {
                            let text_line_pitch = 21.99f64;
                            let row_pitch =
                                27.75 + (legend_max_lines as f64 - 1.0) * 21.76;
                            let title_off = if has_explicit_title {
                                14.85
                            } else if has_auto_title {
                                17.68
                            } else {
                                0.0
                            };
                            let block_h = n_ser as f64 * row_pitch;
                            let legend_y0 =
                                sy + shh / 2.0 + title_off - block_h / 2.0 + 8.97;
                            for row in 0..n_ser {
                                // A stacked area paints series 0 at the BOTTOM,
                                // and its legend follows: Ser2 on top.
                                let si = if is_stacked { n_ser - 1 - row } else { row };
                                let sw_y = legend_y0 + row as f64 * row_pitch;
                                let col_hex = pres
                                    .theme_colors
                                    .get(&format!("accent{}", si + 1))
                                    .map(|s| s.as_str())
                                    .or_else(|| DEFAULT_ACCENT.get(si % 6).copied());
                                if let Some(rgb) = col_hex.and_then(parse_hex_rgb) {
                                    let brush = CreateSolidBrush(COLORREF(colorref(
                                        rgb.0, rgb.1, rgb.2,
                                    )));
                                    let old_brush = SelectObject(mem_dc, brush);
                                    let r = RECT {
                                        left: (swatch_x0 * scale).round() as i32,
                                        top: (sw_y * scale).round() as i32,
                                        right: (swatch_x1 * scale).round() as i32,
                                        bottom: ((sw_y + swatch_w) * scale).round()
                                            as i32,
                                    };
                                    let _ = FillRect(mem_dc, &r, brush);
                                    SelectObject(mem_dc, old_brush);
                                    let _ = DeleteObject(brush);
                                }
                                let label_baseline = sw_y + swatch_w + 0.28;
                                if let Some(lines) = legend_lines.get(si) {
                                    for (li, line) in lines.iter().enumerate() {
                                        draw_text_baseline(
                                            mem_dc,
                                            (label_x0 * scale).round() as i32,
                                            (label_baseline
                                                + li as f64 * text_line_pitch)
                                                as f32,
                                            line,
                                            legend_fs,
                                            axis_family,
                                            None,
                                            scale,
                                        );
                                    }
                                }
                            }
                        }
                        } else if chart.chart_type == "bubble"
                            && std::env::var("OXI_BUBBLE_DISABLE").is_err()
                        {
                        // BUBBLE chart (Word render-truth 2026-08-10: the 8-arm
                        // probe chart_bubble U1..U8 plus a 10-arm frame-HEIGHT
                        // sweep, an 8-arm frame-WIDTH sweep and a 6-arm
                        // bubbleScale/sizeRepresents sweep = 20 measured arms,
                        // residuals <= 0.004pt).
                        //
                        // The plot geometry is the SCATTER model (both axes
                        // numeric, plot_left = sx + 6.50 + w(widest Y label) +
                        // 16.70, plot_right = band_right - w(last X label)/2,
                        // X through horiz_value_axis and Y through
                        // nice_axis_max/nice_axis_range) -- but unlike scatter a
                        // bubble chart DOES draw an automatic title for a single
                        // series, so
                        //   plot_top = sy + 51.35  (auto title;   U1 123.35)
                        //            = sy + 45.69  (explicit tit; U8 117.69)
                        //            = sy + 16.0   (no title;     U3  87.99)
                        //
                        // ---- the bubble SIZE law -------------------------
                        //   avail_w = sw - 10.0
                        //   avail_h = (sy + shh + 6.0) - plot_top
                        //             ^ the FRAME bottom, NOT plot_bot: U7's
                        //               negative data moves plot_bot to
                        //               sy+shh-16 yet leaves r_max at 28.0,
                        //               identical to U1.
                        //   d_max   = min(avail_w, avail_h)
                        //             * 3*scale / (3*scale + 1000)   [percent]
                        //   r_i     = (d_max/2) * (size_i/size_max)        ["w"]
                        //           = (d_max/2) * sqrt(size_i/size_max)  [area]
                        // The bubbleScale response SATURATES; 50/100/200/300
                        // measured r_max 15.82 / 28.00 / 45.49 / 57.47 against
                        // 15.82 / 28.00 / 45.50 / 57.48 predicted (at scale 100
                        // the factor reduces to the clean 3/13).  size_max is
                        // GLOBAL across every series (U8's two series share it).
                        //
                        //   bubbles: accent(series) FILL, no stroke.
                        //   data labels: left edge cx + r + 8.55, baseline
                        //     cy + 6.20 (U6 measured 8.53/8.55/8.57 and
                        //     6.20/6.21/6.20 over the three bubbles).
                        //   legend: CIRCULAR swatch r=4.94 centred at
                        //     label_x0 - 9.59, label baseline swatch_cy + 5.24,
                        //     a block of n*27.75 centred on sy + shh/2 +
                        //     title_off (0 / 14.85 explicit / 17.68 auto) --
                        //     the same block model as line/area.  The band it
                        //     takes is swatch_cx - 23.22 (U8).
                        let axis_family = "Calibri";
                        let axis_fs = 18.0f32;
                        let sx = sh.x as f64;
                        let sy = sh.y as f64;
                        let sw = sh.width as f64;
                        let shh = sh.height as f64;
                        let label_w = |s: &str| {
                            font_adv::line_hmtx_width_pt(s, 18.0, axis_family)
                                .unwrap_or_else(|| {
                                    s.chars().count() as f32 * 18.0 * 0.5
                                }) as f64
                        };
                        let has_explicit_title = chart.explicit_title.is_some();
                        let has_auto_title = chart.series.len() == 1
                            && !chart.auto_title_deleted
                            && !has_explicit_title;

                        let (y_min_data, y_max_data) = chart
                            .series
                            .iter()
                            .flat_map(|s| s.values.iter().copied())
                            .fold((0.0f64, 0.0f64), |(lo, hi), v| {
                                (lo.min(v), hi.max(v))
                            });
                        let y_has_neg = y_min_data < 0.0;
                        // REAL extremes (the folds above are seeded with 0.0, so
                        // they already carry the union with the origin).  The
                        // bubble expansion has to grow from the data itself:
                        // padding min(0, data) would drag a positive-only axis
                        // below zero, which Word never does.
                        let (y_lo_real, y_hi_real) = chart
                            .series
                            .iter()
                            .flat_map(|s| s.values.iter().copied())
                            .fold((f64::INFINITY, f64::NEG_INFINITY), |(lo, hi), v| {
                                (lo.min(v), hi.max(v))
                            });
                        let (y_lo_real, y_hi_real) = if y_lo_real.is_finite() {
                            (y_lo_real, y_hi_real)
                        } else {
                            (0.0, 0.0)
                        };
                        let plot_top = sy
                            + if has_explicit_title {
                                45.69
                            } else if has_auto_title {
                                51.35
                            } else {
                                16.0
                            };
                        let plot_bot = sy + shh - if y_has_neg { 16.0 } else { 39.9 };
                        let plot_h = plot_bot - plot_top;

                        // ---- bubble radii (the size law above) ----
                        // r_max depends only on the frame and plot_top, so it is
                        // known before either axis -- which matters because the
                        // axes are expanded by it (below).
                        let size_max = chart
                            .series
                            .iter()
                            .flat_map(|s| s.sizes.iter().copied())
                            .fold(0.0f64, |a, v| a.max(v.abs()));
                        let scale_pct = chart.bubble_scale.max(1.0);
                        let avail_w = sw - 10.0;
                        let avail_h = (sy + shh + 6.0) - plot_top;
                        let r_max = avail_w.min(avail_h) * 3.0 * scale_pct
                            / (3.0 * scale_pct + 1000.0)
                            / 2.0;
                        let by_width = chart.size_represents == "w";

                        // ---- the bubble AXIS-EXTENT law ------------------
                        // Word expands a bubble value axis by the LARGEST bubble
                        // radius at BOTH ends (converted to data units through the
                        // axis itself) INSTEAD of the ordinary 5% headroom, then
                        // settles on the fixed point.  U7 is the discriminator:
                        // y in [-8, 12], r_max 28, plot_h 220.70.  The plain rule
                        // gives [-10, 15]; Word draws [-15, 20], and
                        //   ppu = 220.70/35 = 6.306,  28/6.306 = 4.44
                        //   eff = [-12.44, 16.44] -> step 5 -> [-15, 20]
                        // reproduces it and is stable.  PER-POINT radii do NOT
                        // (they leave [-10, 15]); r_max at both ends does.  Every
                        // positive arm is unchanged because the expansion lands
                        // where the headroom already did (U1 21.4 + 28/7.87 =
                        // 24.96 -> 25, U3 24.86 -> 25, U8 24.94 -> 25).
                        // Dividing by AXIS_HEADROOM cancels the 5% that
                        // nice_axis_range / nice_axis_max apply internally.
                        // The expansion is r_max less ~1pt.  The window is
                        // pinned by two arms that straddle it: chart_bubble_scale
                        // s3 (plot_h 148.75, r 22.46) keeps 25 -- so the pad must
                        // be <= 21.42 = r - 1.04 -- while chart_bubble_size s4
                        // (plot_h 196.75, r 57.47, axis 30) must exceed 30 -- so
                        // the pad is > 56.40 = r - 1.07.  Hence r_max - c with
                        // c in [1.04, 1.07); the ~1pt is probably the bubble
                        // outline, and no pure multiple of r_max satisfies both.
                        let pad_pt = (r_max - 1.05).max(0.0);
                        // Word sizes the value axis with the ordinary ~5-tick
                        // step (nice_axis_max's rule) and COARSENS it only when
                        // the ticks would crowd below VERT_MIN_SPACING.  The
                        // division count then falls out of the rounded max --
                        // which is why chart_bubble_size s3/s4 show 6 and 7
                        // divisions (0..30 / 0..35 by 5) while the tall frames
                        // stay at 5, and the 160/200pt frames drop to 2 and 3.
                        // Taking the FINEST step that clears the spacing instead
                        // gives 11 divisions on a tall plot, which Word never does.
                        let pick_y = |lo: f64, hi: f64, len: f64| -> (f64, f64, usize) {
                            let span = (hi - lo).max(1e-9);
                            let raw = span / 5.0;
                            let mut mag = 10f64.powf(raw.log10().floor());
                            let resid = raw / mag;
                            let mut m = if resid < 1.5 {
                                1.0
                            } else if resid < 3.0 {
                                2.0
                            } else if resid < 7.0 {
                                5.0
                            } else {
                                mag *= 10.0;
                                1.0
                            };
                            let mut out = (lo, hi.max(1.0), 1usize);
                            for _ in 0..24 {
                                let step = m * mag;
                                let amax = (hi / step - 1e-9).ceil() * step;
                                let amin = (lo / step + 1e-9).floor() * step;
                                let div = ((amax - amin) / step).round().max(1.0);
                                out = (amin, amax, div as usize);
                                if div <= 1.0 || len / div >= VERT_MIN_SPACING {
                                    break;
                                }
                                if m == 1.0 {
                                    m = 2.0;
                                } else if m == 2.0 {
                                    m = 5.0;
                                } else {
                                    m = 1.0;
                                    mag *= 10.0;
                                }
                            }
                            out
                        };
                        let (mut y_min, mut y_axis, mut y_steps) =
                            pick_y(y_min_data, y_max_data, plot_h);
                        for _ in 0..3 {
                            let span = (y_axis - y_min).max(1e-9);
                            let pad = if plot_h > 0.0 {
                                pad_pt * span / plot_h
                            } else {
                                0.0
                            };
                            let elo = (y_lo_real - pad).min(0.0);
                            let ehi = (y_hi_real + pad).max(0.0);
                            let (lo, hi, dv) = pick_y(elo, ehi, plot_h);
                            if (lo - y_min).abs() < 1e-9
                                && (hi - y_axis).abs() < 1e-9
                            {
                                break;
                            }
                            y_min = lo;
                            y_axis = hi;
                            y_steps = dv.max(1);
                        }
                        let (y_min, y_axis, y_steps) = (y_min, y_axis, y_steps);
                        let y_span = (y_axis - y_min).max(1e-9);
                        let y_step = y_span / y_steps as f64;
                        let y_lab_w = (0..=y_steps)
                            .map(|i| {
                                label_w(&fmt_axis_value(
                                    y_min + y_span * i as f64 / y_steps as f64,
                                    y_step,
                                ))
                            })
                            .fold(0.0f64, f64::max);
                        let x_min_data = chart
                            .series
                            .iter()
                            .flat_map(|s| s.x_values.iter().copied())
                            .fold(0.0f64, f64::min);
                        let x_has_neg = x_min_data < 0.0;
                        let (x_lo_real, x_hi_real) = chart
                            .series
                            .iter()
                            .flat_map(|s| s.x_values.iter().copied())
                            .fold((f64::INFINITY, f64::NEG_INFINITY), |(lo, hi), v| {
                                (lo.min(v), hi.max(v))
                            });
                        let (x_lo_real, x_hi_real) = if x_lo_real.is_finite() {
                            (x_lo_real, x_hi_real)
                        } else {
                            (0.0, 0.0)
                        };
                        let val_y =
                            |v: f64| plot_bot - ((v - y_min) / y_span) * plot_h;
                        let zero_y = val_y(0.0);

                        // Legend band (a bare <c:legend/> overlays the plot --
                        // the pie/doughnut/area discriminator).
                        let legend_active = chart.has_legend && !chart.legend_overlay;
                        let cap = legend_label_cap(sw);
                        let mut legend_lines: Vec<Vec<String>> = Vec::new();
                        let mut max_label_w = 0.0f64;
                        if legend_active {
                            for s in chart.series.iter() {
                                let lines =
                                    wrap_legend_label(&s.name, 18.0, axis_family, cap);
                                for l in lines.iter() {
                                    max_label_w = max_label_w.max(label_w(l));
                                }
                                legend_lines.push(lines);
                            }
                        }
                        let legend_label_x0 = sx + sw - 10.0 - max_label_w;
                        let legend_swatch_cx = legend_label_x0 - 9.59;
                        let band_right = if legend_active {
                            legend_swatch_cx - 23.22
                        } else {
                            sx + sw - 11.0
                        };

                        // Numeric X axis: plot_right depends on the width of the
                        // last tick label, which depends on the axis, which
                        // depends on plot_w -> converge in 2 passes.
                        let _x_max_data = chart
                            .series
                            .iter()
                            .flat_map(|s| s.x_values.iter().copied())
                            .fold(0.0f64, f64::max);
                        let plot_left_pos = sx + 6.50 + y_lab_w + 16.70;
                        let mut plot_left = plot_left_pos;
                        let mut plot_right = band_right;
                        let mut x_min = 0.0f64;
                        let mut x_axis = 1.0f64;
                        let mut x_div = 1usize;
                        // Same bubble expansion on X, inside the loop that
                        // already resolved the label-width/axis circularity
                        // (pass 0 seeds it without the expansion because the span
                        // is not known yet).
                        // Finest 1/2/5 step whose tick spacing clears
                        // label_width + BUBBLE_LABEL_GAP (see the constant).
                        let pick_x = |lo: f64, hi: f64, len: f64| -> (f64, f64, usize) {
                            for k in -6i32..=9 {
                                let mag = 10f64.powi(k);
                                for m in [1.0f64, 2.0, 5.0] {
                                    let step = m * mag;
                                    let amax = (hi / step - 1e-9).ceil() * step;
                                    let amin = (lo / step + 1e-9).floor() * step;
                                    let div = ((amax - amin) / step).round();
                                    if div < 1.0 || div > 1000.0 {
                                        continue;
                                    }
                                    let wl = label_w(&fmt_axis_value(amin, step)).max(
                                        label_w(&fmt_axis_value(amax, step)),
                                    );
                                    if len / div >= wl + BUBBLE_LABEL_GAP {
                                        return (amin, amax, div as usize);
                                    }
                                }
                            }
                            (lo.min(0.0), hi.max(1.0), 1)
                        };
                        for pass in 0..4 {
                            let pw = (plot_right - plot_left).max(1.0);
                            let pad = if pass == 0 {
                                0.0
                            } else {
                                pad_pt * (x_axis - x_min).max(1e-9) / pw
                            };
                            let elo = (x_lo_real - pad).min(0.0);
                            let ehi = (x_hi_real + pad).max(0.0);
                            let (lo, hi, dv) = pick_x(elo, ehi, pw);
                            x_min = if x_has_neg { lo } else { 0.0 };
                            x_axis = hi;
                            x_div = dv.max(1);
                            let step = (x_axis - x_min) / x_div as f64;
                            if x_has_neg {
                                plot_left = sx
                                    + 11.0
                                    + label_w(&fmt_axis_value(x_min, step)) / 2.0;
                            }
                            plot_right = band_right
                                - label_w(&fmt_axis_value(x_axis, step)) / 2.0;
                        }
                        let plot_left = plot_left;
                        let plot_w = plot_right - plot_left;
                        let x_span = (x_axis - x_min).max(1e-9);
                        let x_step = x_span / x_div as f64;
                        let val_x =
                            |v: f64| plot_left + ((v - x_min) / x_span) * plot_w;
                        let zero_x = val_x(0.0);
                        let y_axis_x = if x_has_neg { zero_x } else { plot_left };
                        let x_axis_y = if y_has_neg { zero_y } else { plot_bot };

                        let radius = |sz: f64| -> f64 {
                            if size_max <= 0.0 {
                                return 0.0;
                            }
                            let f = sz.abs() / size_max;
                            if by_width {
                                r_max * f
                            } else {
                                r_max * f.sqrt()
                            }
                        };

                        // Y axis labels (right edge y_axis_x-16.70).
                        for i in 0..=y_steps {
                            let val = y_min + y_span * i as f64 / y_steps as f64;
                            let tick_y = plot_bot - plot_h * i as f64 / y_steps as f64;
                            let label = fmt_axis_value(val, y_step);
                            let lx = y_axis_x - 16.70 - label_w(&label);
                            draw_text_baseline(
                                mem_dc,
                                (lx * scale).round() as i32,
                                (tick_y + 5.20) as f32,
                                &label,
                                18.0,
                                axis_family,
                                None,
                                scale,
                            );
                        }

                        // Horizontal major gridlines, BEFORE the bubbles.
                        let grid_pen =
                            CreatePen(PS_SOLID, 2, COLORREF(colorref(0, 0, 0)));
                        let old_grid_pen = SelectObject(mem_dc, grid_pen);
                        let _ = SelectObject(mem_dc, GetStockObject(NULL_BRUSH));
                        let gl = (plot_left * scale).round() as i32;
                        let gr = (plot_right * scale).round() as i32;
                        let grid_from = if y_has_neg { 0 } else { 1 };
                        for i in grid_from..=y_steps {
                            let grid_y = plot_bot - plot_h * i as f64 / y_steps as f64;
                            let gy = (grid_y * scale).round() as i32;
                            let _ = MoveToEx(mem_dc, gl, gy, None);
                            let _ = LineTo(mem_dc, gr, gy);
                        }
                        SelectObject(mem_dc, old_grid_pen);
                        let _ = DeleteObject(grid_pen);

                        // Bubbles: accent fill, NO stroke.
                        for (si, s) in chart.series.iter().enumerate() {
                            let (fill_hex, _) = line_series_colors(si);
                            if let Some(rgb) = parse_hex_rgb(&fill_hex) {
                                let brush = CreateSolidBrush(COLORREF(colorref(
                                    rgb.0, rgb.1, rgb.2,
                                )));
                                let old_brush = SelectObject(mem_dc, brush);
                                let _ = SelectObject(mem_dc, GetStockObject(NULL_PEN));
                                for (i, v) in s.values.iter().enumerate() {
                                    let xv =
                                        s.x_values.get(i).copied().unwrap_or(i as f64);
                                    let r = radius(
                                        s.sizes.get(i).copied().unwrap_or(0.0),
                                    );
                                    if r > 0.0 {
                                        draw_bubble_circle(
                                            mem_dc,
                                            val_x(xv),
                                            val_y(*v),
                                            r,
                                            scale,
                                        );
                                    }
                                }
                                SelectObject(mem_dc, old_brush);
                                let _ = DeleteObject(brush);
                            }
                        }

                        // Data labels: left edge cx + r + 8.55, baseline
                        // cy + 6.20.
                        if chart.has_data_labels && chart.show_val {
                            let num_fmt = chart.number_format.clone();
                            for s in chart.series.iter() {
                                for (i, v) in s.values.iter().enumerate() {
                                    let xv =
                                        s.x_values.get(i).copied().unwrap_or(i as f64);
                                    let r = radius(
                                        s.sizes.get(i).copied().unwrap_or(0.0),
                                    );
                                    let text = if num_fmt == "0.0%" {
                                        format!("{:.1}%", v * 100.0)
                                    } else {
                                        format!("{}", v)
                                    };
                                    draw_text_baseline(
                                        mem_dc,
                                        ((val_x(xv) + r + 8.55) * scale).round() as i32,
                                        (val_y(*v) + 6.20) as f32,
                                        &text,
                                        axis_fs,
                                        axis_family,
                                        None,
                                        scale,
                                    );
                                }
                            }
                        }

                        // X axis tick labels, centred on the tick.
                        for i in 0..=x_div {
                            let val = x_min + x_span * i as f64 / x_div as f64;
                            let tick_x = plot_left + plot_w * i as f64 / x_div as f64;
                            let label = fmt_axis_value(val, x_step);
                            draw_text_baseline(
                                mem_dc,
                                ((tick_x - label_w(&label) / 2.0) * scale).round() as i32,
                                (x_axis_y + 28.67) as f32,
                                &label,
                                18.0,
                                axis_family,
                                None,
                                scale,
                            );
                        }

                        // Title: explicit <c:title> (Arial 18pt, baseline
                        // sy+24.43) else the AUTOMATIC single-series title
                        // (Calibri-Bold 21.62pt, baseline sy+28.03).
                        if let Some(title) = &chart.explicit_title {
                            let tfs = 18.0f32;
                            let lw = font_adv::line_hmtx_width_pt(title, tfs, "Arial")
                                .unwrap_or_else(|| {
                                    title.chars().count() as f32 * tfs * 0.5
                                }) as f64;
                            draw_text_baseline_w(
                                mem_dc,
                                ((sx + sw / 2.0 - lw / 2.0) * scale).round() as i32,
                                (sy + 24.43) as f32,
                                title,
                                tfs,
                                "Arial",
                                None,
                                scale,
                                400,
                            );
                        } else if has_auto_title {
                            let tfs = 21.62f32;
                            let name = chart.series[0].name.clone();
                            let lw = font_adv::line_hmtx_width_pt(&name, tfs, axis_family)
                                .unwrap_or_else(|| {
                                    name.chars().count() as f32 * tfs * 0.5
                                }) as f64;
                            draw_text_baseline_w(
                                mem_dc,
                                ((sx + sw / 2.0 - lw / 2.0) * scale).round() as i32,
                                (sy + 28.03) as f32,
                                &name,
                                tfs,
                                axis_family,
                                None,
                                scale,
                                700,
                            );
                        }

                        // Legend: circular swatch r=4.94 + label.
                        if legend_active {
                            let n = chart.series.len().max(1);
                            let row_pitch = 27.75
                                + (legend_lines
                                    .iter()
                                    .map(|l| l.len())
                                    .max()
                                    .unwrap_or(1)
                                    .saturating_sub(1)) as f64
                                    * 21.76;
                            let title_off = if has_explicit_title {
                                14.85
                            } else if has_auto_title {
                                17.68
                            } else {
                                0.0
                            };
                            let block_top = sy + shh / 2.0 + title_off
                                - n as f64 * row_pitch / 2.0;
                            for (si, lines) in legend_lines.iter().enumerate() {
                                let cy = block_top
                                    + row_pitch / 2.0
                                    + si as f64 * row_pitch;
                                let (fill_hex, _) = line_series_colors(si);
                                if let Some(rgb) = parse_hex_rgb(&fill_hex) {
                                    let brush = CreateSolidBrush(COLORREF(colorref(
                                        rgb.0, rgb.1, rgb.2,
                                    )));
                                    let old_brush = SelectObject(mem_dc, brush);
                                    let _ =
                                        SelectObject(mem_dc, GetStockObject(NULL_PEN));
                                    draw_bubble_circle(
                                        mem_dc,
                                        legend_swatch_cx,
                                        cy,
                                        4.94,
                                        scale,
                                    );
                                    SelectObject(mem_dc, old_brush);
                                    let _ = DeleteObject(brush);
                                }
                                for (li, line) in lines.iter().enumerate() {
                                    draw_text_baseline(
                                        mem_dc,
                                        (legend_label_x0 * scale).round() as i32,
                                        (cy + 5.24 + li as f64 * 21.99) as f32,
                                        line,
                                        18.0,
                                        axis_family,
                                        None,
                                        scale,
                                    );
                                }
                            }
                        }

                        // Axis lines + ticks (each axis rides the OTHER axis'
                        // zero crossing when that axis spans zero).
                        let axis_pen =
                            CreatePen(PS_SOLID, 2, COLORREF(colorref(0, 0, 0)));
                        let old_axis_pen = SelectObject(mem_dc, axis_pen);
                        let _ = SelectObject(mem_dc, GetStockObject(NULL_BRUSH));
                        let pl = (plot_left * scale).round() as i32;
                        let pt = (plot_top * scale).round() as i32;
                        let pr = (plot_right * scale).round() as i32;
                        let pb = (plot_bot * scale).round() as i32;
                        let ax_x = (y_axis_x * scale).round() as i32;
                        let ax_y = (x_axis_y * scale).round() as i32;
                        let _ = MoveToEx(mem_dc, ax_x, pt, None);
                        let _ = LineTo(mem_dc, ax_x, pb);
                        let _ = MoveToEx(mem_dc, pl, ax_y, None);
                        let _ = LineTo(mem_dc, pr, ax_y);
                        let _ = MoveToEx(mem_dc, pl, pt, None);
                        let _ = LineTo(mem_dc, pr, pt);
                        for i in 0..=y_steps {
                            let tick_y = plot_bot - plot_h * i as f64 / y_steps as f64;
                            let ty = (tick_y * scale).round() as i32;
                            let _ = MoveToEx(
                                mem_dc,
                                ((y_axis_x - 5.71) * scale).round() as i32,
                                ty,
                                None,
                            );
                            let _ = LineTo(mem_dc, ax_x, ty);
                        }
                        for i in 0..=x_div {
                            let tick_x = plot_left + plot_w * i as f64 / x_div as f64;
                            let tx = (tick_x * scale).round() as i32;
                            let _ = MoveToEx(mem_dc, tx, ax_y, None);
                            let _ = LineTo(
                                mem_dc,
                                tx,
                                ((x_axis_y + 5.71) * scale).round() as i32,
                            );
                        }
                        SelectObject(mem_dc, old_axis_pen);
                        let _ = DeleteObject(axis_pen);
                        } else if chart.chart_type == "scatter"
                            && std::env::var("OXI_SCATTER_DISABLE").is_err()
                        {
                        // XY SCATTER chart (Word render-truth 2026-08-09,
                        // fitz get_drawings + rawdict over the 8-arm probe
                        // chart_scatter S1..S8 -- markers / markers 2ser /
                        // lines+markers / smooth / lines-only / legend /
                        // data labels / no-title):
                        //   plot_left  = sx + 6.50 + w(widest Y label) + 16.70
                        //     (all 8 arms 113.45 = 72 + 6.50 + 18.25 + 16.70;
                        //      the same rule the area and horizontal-bar
                        //      branches use)
                        //   plot_top   = sy + 16.0  (measured 87.99 on EVERY
                        //     arm -- scatter draws NO automatic title even at
                        //     one series, unlike line/pie/bar)
                        //   plot_bot   = sy + shh - 39.9   (320.10)
                        //   band_right = legend ? swatch_x0 - 18.15
                        //                       : sx + sw - 11.0
                        //   plot_right = band_right - w(last X label)/2
                        //     (S1 452.44 = 457.0 - 9.13/2; S6 388.03)
                        //   X axis is NUMERIC and follows horiz_value_axis
                        //     (5% headroom, >=57pt tick spacing): S1 x-max 4
                        //     -> 0..5 in 5 (spacing 67.8), S6's narrower plot
                        //     -> 0..6 in 3 (spacing 91.5; step 1 would give
                        //     54.9 < 57).  plot_right depends on the last
                        //     label's width and the axis depends on plot_w,
                        //     so the two converge in 2 passes.
                        //   point = (plot_left + xv/x_axis*plot_w,
                        //            plot_bot  - yv/y_axis*plot_h)   [x from 0]
                        //   Y axis = nice_axis_max in 5 steps (22.0 -> 25).
                        //   markers: 9.9pt when the series draws NO line
                        //     (S1/S2/S6/S7/S8 measured 9.84 diamond / 9.96
                        //     square) and 6.96pt when it does (S3/S4), per
                        //     series shape+fill exactly like the line branch.
                        //   polyline: border colour w=2.25 through the points
                        //     when the series does not declare <a:noFill>;
                        //     a SMOOTH series (S4) is drawn with the same
                        //     straight segments as S3 in the PDF.
                        //   data labels: LEFT-aligned at point_x + 11.09,
                        //     baseline point_y + 6.22 (S7: widths 18.25 and
                        //     31.96 share the same dx0 -> left edge, not
                        //     centred), Calibri 18pt, raw value.
                        //   legend: per-series MARKER swatch (no line swatch)
                        //     centred at label_x0 - 9.68, rows pitched 27.75
                        //     with the block frame-vertically centred
                        //     (S6 n=2: centres 202.08/229.74 about 216 =
                        //     sy+shh/2), label baseline = swatch_cy + 5.28.
                        let axis_family = "Calibri";
                        let axis_fs = 18.0f32;
                        let sx = sh.x as f64;
                        let sy = sh.y as f64;
                        let sw = sh.width as f64;
                        let shh = sh.height as f64;
                        let label_w = |s: &str| {
                            font_adv::line_hmtx_width_pt(s, 18.0, axis_family)
                                .unwrap_or_else(|| {
                                    s.chars().count() as f32 * 18.0 * 0.5
                                }) as f64
                        };

                        // NEGATIVE data (chart_negative N4/N5, 2026-08-10):
                        // each axis spans zero and the OTHER axis' line, ticks
                        // and labels ride the zero crossing -- the Y labels
                        // right-align 16.64pt left of zero_x (N5: 222.53 for
                        // zero_x 239.18) and the X labels hang zero_y + 28.68
                        // (N4/N5: 299.54 for zero_y 270.86).  The 39.9pt bottom
                        // band collapses to 16.0 because the X labels are no
                        // longer under the plot.
                        let (y_min_data, y_max_data) = chart
                            .series
                            .iter()
                            .flat_map(|s| s.values.iter().copied())
                            .fold((0.0f64, 0.0f64), |(lo, hi), v| {
                                (lo.min(v), hi.max(v))
                            });
                        let y_has_neg = y_min_data < 0.0;
                        let plot_top = sy + 16.0;
                        let plot_bot = sy + shh - if y_has_neg { 16.0 } else { 39.9 };
                        let plot_h = plot_bot - plot_top;
                        let (y_min, y_axis, y_steps) = if y_has_neg {
                            nice_axis_range(
                                y_min_data,
                                y_max_data,
                                plot_h,
                                VERT_MIN_SPACING,
                            )
                        } else {
                            (0.0, nice_axis_max(y_max_data), 5usize)
                        };
                        let y_span = (y_axis - y_min).max(1e-9);
                        let y_step = y_span / y_steps as f64;
                        let y_lab_w = (0..=y_steps)
                            .map(|i| {
                                label_w(&fmt_axis_value(
                                    y_min + y_span * i as f64 / y_steps as f64,
                                    y_step,
                                ))
                            })
                            .fold(0.0f64, f64::max);
                        let x_min_data = chart
                            .series
                            .iter()
                            .flat_map(|s| s.x_values.iter().copied())
                            .fold(0.0f64, f64::min);
                        let x_has_neg = x_min_data < 0.0;
                        let val_y =
                            |v: f64| plot_bot - ((v - y_min) / y_span) * plot_h;
                        let zero_y = val_y(0.0);

                        // Legend band (only a legend that declares
                        // <c:overlay val="0"/> takes a band; a bare
                        // <c:legend/> overlays the plot -- the pie/doughnut
                        // and area discriminator).
                        let legend_active = chart.has_legend && !chart.legend_overlay;
                        let cap = legend_label_cap(sw);
                        let mut legend_lines: Vec<Vec<String>> = Vec::new();
                        let mut max_label_w = 0.0f64;
                        if legend_active {
                            for s in chart.series.iter() {
                                let lines =
                                    wrap_legend_label(&s.name, 18.0, axis_family, cap);
                                for l in lines.iter() {
                                    max_label_w = max_label_w.max(label_w(l));
                                }
                                legend_lines.push(lines);
                            }
                        }
                        let legend_label_x0 = sx + sw - 10.0 - max_label_w;
                        let legend_swatch_cx = legend_label_x0 - 9.68;
                        let band_right = if legend_active {
                            (legend_swatch_cx - 9.9 / 2.0) - 18.15
                        } else {
                            sx + sw - 11.0
                        };

                        // Numeric X axis: plot_right depends on the width of
                        // the last tick label, which depends on the axis,
                        // which depends on plot_w -> converge in 2 passes.
                        let x_max_data = chart
                            .series
                            .iter()
                            .flat_map(|s| s.x_values.iter().copied())
                            .fold(0.0f64, f64::max);
                        let plot_left_pos = sx + 6.50 + y_lab_w + 16.70;
                        let mut plot_left = plot_left_pos;
                        let mut plot_right = band_right;
                        let mut x_min = 0.0f64;
                        let mut x_axis = 1.0f64;
                        let mut x_div = 1usize;
                        for _ in 0..2 {
                            let pw = (plot_right - plot_left).max(1.0);
                            if x_has_neg {
                                let (lo, hi, dv) = nice_axis_range(
                                    x_min_data,
                                    x_max_data,
                                    pw,
                                    HORIZ_MIN_SPACING,
                                );
                                x_min = lo;
                                x_axis = hi;
                                x_div = dv.max(1);
                            } else {
                                let (ax, dv) = horiz_value_axis(x_max_data, pw);
                                x_min = 0.0;
                                x_axis = ax;
                                x_div = dv;
                            }
                            let step = (x_axis - x_min) / x_div as f64;
                            if x_has_neg {
                                // Both edges leave half the outermost label
                                // (N5: 90.32 = 72 + 11.0 + w("-4")/2).
                                plot_left = sx
                                    + 11.0
                                    + label_w(&fmt_axis_value(x_min, step)) / 2.0;
                            }
                            plot_right = band_right
                                - label_w(&fmt_axis_value(x_axis, step)) / 2.0;
                        }
                        let plot_left = plot_left;
                        let plot_w = plot_right - plot_left;
                        let x_span = (x_axis - x_min).max(1e-9);
                        let x_step = x_span / x_div as f64;
                        let val_x =
                            |v: f64| plot_left + ((v - x_min) / x_span) * plot_w;
                        let zero_x = val_x(0.0);

                        // Each axis line rides the OTHER axis' zero crossing.
                        let y_axis_x = if x_has_neg { zero_x } else { plot_left };
                        let x_axis_y = if y_has_neg { zero_y } else { plot_bot };

                        // Y axis labels (right edge y_axis_x-16.70,
                        // baseline tick_y+5.20).
                        for i in 0..=y_steps {
                            let val = y_min + y_span * i as f64 / y_steps as f64;
                            let tick_y = plot_bot - plot_h * i as f64 / y_steps as f64;
                            let label = fmt_axis_value(val, y_step);
                            let lx = y_axis_x - 16.70 - label_w(&label);
                            draw_text_baseline(
                                mem_dc,
                                (lx * scale).round() as i32,
                                (tick_y + 5.20) as f32,
                                &label,
                                18.0,
                                axis_family,
                                None,
                                scale,
                            );
                        }

                        // Horizontal major gridlines at the value ticks
                        // i=1..=y_steps, BEFORE the series.
                        let grid_pen =
                            CreatePen(PS_SOLID, 2, COLORREF(colorref(0, 0, 0)));
                        let old_grid_pen = SelectObject(mem_dc, grid_pen);
                        let _ = SelectObject(mem_dc, GetStockObject(NULL_BRUSH));
                        let gl = (plot_left * scale).round() as i32;
                        let gr = (plot_right * scale).round() as i32;
                        let grid_from = if y_has_neg { 0 } else { 1 };
                        for i in grid_from..=y_steps {
                            let grid_y = plot_bot - plot_h * i as f64 / y_steps as f64;
                            let gy = (grid_y * scale).round() as i32;
                            let _ = MoveToEx(mem_dc, gl, gy, None);
                            let _ = LineTo(mem_dc, gr, gy);
                        }
                        SelectObject(mem_dc, old_grid_pen);
                        let _ = DeleteObject(grid_pen);

                        // Points per series: x from the series' OWN xVal.
                        let series_pts: Vec<Vec<(f64, f64)>> = chart
                            .series
                            .iter()
                            .map(|s| {
                                s.values
                                    .iter()
                                    .enumerate()
                                    .map(|(i, v)| {
                                        let xv =
                                            s.x_values.get(i).copied().unwrap_or(i as f64);
                                        (val_x(xv), val_y(*v))
                                    })
                                    .collect()
                            })
                            .collect();

                        // Connecting lines (series without <a:ln><a:noFill/>).
                        for (si, pts) in series_pts.iter().enumerate() {
                            if chart.series.get(si).map_or(false, |s| s.line_none) {
                                continue;
                            }
                            let (_, border_hex) = line_series_colors(si);
                            if let Some(rgb) = parse_hex_rgb(&border_hex) {
                                let line_pen = CreatePen(
                                    PS_SOLID,
                                    (2.25 * scale).round().max(1.0) as i32,
                                    COLORREF(colorref(rgb.0, rgb.1, rgb.2)),
                                );
                                let old_line_pen = SelectObject(mem_dc, line_pen);
                                let _ = SelectObject(mem_dc, GetStockObject(NULL_BRUSH));
                                for w in pts.windows(2) {
                                    let _ = MoveToEx(
                                        mem_dc,
                                        (w[0].0 * scale).round() as i32,
                                        (w[0].1 * scale).round() as i32,
                                        None,
                                    );
                                    let _ = LineTo(
                                        mem_dc,
                                        (w[1].0 * scale).round() as i32,
                                        (w[1].1 * scale).round() as i32,
                                    );
                                }
                                SelectObject(mem_dc, old_line_pen);
                                let _ = DeleteObject(line_pen);
                            }
                        }

                        // Markers (series without <c:symbol val="none"/>):
                        // 9.9pt when the series has no connecting line,
                        // 6.96pt when it does.
                        for (si, pts) in series_pts.iter().enumerate() {
                            let ser = match chart.series.get(si) {
                                Some(s) => s,
                                None => continue,
                            };
                            if ser.marker_none {
                                continue;
                            }
                            let (fill_hex, _) = line_series_colors(si);
                            if let Some(rgb) = parse_hex_rgb(&fill_hex) {
                                let m_brush =
                                    CreateSolidBrush(COLORREF(colorref(rgb.0, rgb.1, rgb.2)));
                                let old_m_brush = SelectObject(mem_dc, m_brush);
                                let _ = SelectObject(mem_dc, GetStockObject(NULL_PEN));
                                let mr = if ser.line_none { 9.9 / 2.0 } else { 6.96 / 2.0 };
                                for (px, py) in pts.iter() {
                                    draw_line_marker(
                                        mem_dc,
                                        line_marker_shape(si),
                                        *px,
                                        *py,
                                        mr,
                                        scale,
                                    );
                                }
                                SelectObject(mem_dc, old_m_brush);
                                let _ = DeleteObject(m_brush);
                            }
                        }

                        // Data labels: LEFT-aligned at point_x + 11.09,
                        // baseline point_y + 6.22.
                        if chart.has_data_labels && chart.show_val {
                            let num_fmt = chart.number_format.clone();
                            for (si, pts) in series_pts.iter().enumerate() {
                                let vals = match chart.series.get(si) {
                                    Some(s) => &s.values,
                                    None => continue,
                                };
                                for (pt, v) in pts.iter().zip(vals.iter()) {
                                    let text = if num_fmt == "0.0%" {
                                        format!("{:.1}%", v * 100.0)
                                    } else {
                                        format!("{}", v)
                                    };
                                    draw_text_baseline(
                                        mem_dc,
                                        ((pt.0 + 11.09) * scale).round() as i32,
                                        (pt.1 + 6.22) as f32,
                                        &text,
                                        axis_fs,
                                        axis_family,
                                        None,
                                        scale,
                                    );
                                }
                            }
                        }

                        // X axis tick labels, centred on the tick.
                        for i in 0..=x_div {
                            let val = x_min + x_span * i as f64 / x_div as f64;
                            let tick_x = plot_left + plot_w * i as f64 / x_div as f64;
                            let label = fmt_axis_value(val, x_step);
                            draw_text_baseline(
                                mem_dc,
                                ((tick_x - label_w(&label) / 2.0) * scale).round() as i32,
                                (x_axis_y + 28.67) as f32,
                                &label,
                                18.0,
                                axis_family,
                                None,
                                scale,
                            );
                        }

                        // EXPLICIT <c:title>: Arial 18pt centred, baseline
                        // sy+24.43 (same as every other chart type; scatter
                        // never draws an AUTOMATIC title).
                        if let Some(title) = &chart.explicit_title {
                            let tfs = 18.0f32;
                            let lw = font_adv::line_hmtx_width_pt(title, tfs, "Arial")
                                .unwrap_or_else(|| {
                                    title.chars().count() as f32 * tfs * 0.5
                                }) as f64;
                            let frame_cx = sx + sw / 2.0;
                            draw_text_baseline_w(
                                mem_dc,
                                ((frame_cx - lw / 2.0) * scale).round() as i32,
                                (sy + 24.43) as f32,
                                title,
                                tfs,
                                "Arial",
                                None,
                                scale,
                                400,
                            );
                        }

                        // Legend: marker swatch + label, rows pitched 27.75,
                        // block frame-vertically centred.
                        if legend_active {
                            let n = chart.series.len().max(1);
                            let row_pitch = 27.75
                                + (legend_lines
                                    .iter()
                                    .map(|l| l.len())
                                    .max()
                                    .unwrap_or(1)
                                    .saturating_sub(1)) as f64
                                    * 21.76;
                            let block_top = sy + shh / 2.0 - n as f64 * row_pitch / 2.0;
                            for (si, lines) in legend_lines.iter().enumerate() {
                                let cy = block_top + row_pitch / 2.0 + si as f64 * row_pitch;
                                let (fill_hex, _) = line_series_colors(si);
                                if let Some(rgb) = parse_hex_rgb(&fill_hex) {
                                    let m_brush = CreateSolidBrush(COLORREF(colorref(
                                        rgb.0, rgb.1, rgb.2,
                                    )));
                                    let old_m_brush = SelectObject(mem_dc, m_brush);
                                    let _ = SelectObject(mem_dc, GetStockObject(NULL_PEN));
                                    draw_line_marker(
                                        mem_dc,
                                        line_marker_shape(si),
                                        legend_swatch_cx,
                                        cy,
                                        9.9 / 2.0,
                                        scale,
                                    );
                                    SelectObject(mem_dc, old_m_brush);
                                    let _ = DeleteObject(m_brush);
                                }
                                for (li, line) in lines.iter().enumerate() {
                                    draw_text_baseline(
                                        mem_dc,
                                        (legend_label_x0 * scale).round() as i32,
                                        (cy + 5.28 + li as f64 * 21.99) as f32,
                                        line,
                                        18.0,
                                        axis_family,
                                        None,
                                        scale,
                                    );
                                }
                            }
                        }

                        // Axis lines + ticks: Y axis at plot_left, X axis at
                        // plot_bot, plot top edge, Y ticks 0..=y_steps at
                        // plot_left-5.71, X ticks 0..=x_div at plot_bot+5.71.
                        let axis_pen =
                            CreatePen(PS_SOLID, 2, COLORREF(colorref(0, 0, 0)));
                        let old_axis_pen = SelectObject(mem_dc, axis_pen);
                        let _ = SelectObject(mem_dc, GetStockObject(NULL_BRUSH));
                        let pl = (plot_left * scale).round() as i32;
                        let pt = (plot_top * scale).round() as i32;
                        let pr = (plot_right * scale).round() as i32;
                        let pb = (plot_bot * scale).round() as i32;
                        let ax_x = (y_axis_x * scale).round() as i32;
                        let ax_y = (x_axis_y * scale).round() as i32;
                        let _ = MoveToEx(mem_dc, ax_x, pt, None);
                        let _ = LineTo(mem_dc, ax_x, pb);
                        let _ = MoveToEx(mem_dc, pl, ax_y, None);
                        let _ = LineTo(mem_dc, pr, ax_y);
                        let _ = MoveToEx(mem_dc, pl, pt, None);
                        let _ = LineTo(mem_dc, pr, pt);
                        for i in 0..=y_steps {
                            let tick_y = plot_bot - plot_h * i as f64 / y_steps as f64;
                            let ty = (tick_y * scale).round() as i32;
                            let _ = MoveToEx(
                                mem_dc,
                                ((y_axis_x - 5.71) * scale).round() as i32,
                                ty,
                                None,
                            );
                            let _ = LineTo(mem_dc, ax_x, ty);
                        }
                        for i in 0..=x_div {
                            let tick_x = plot_left + plot_w * i as f64 / x_div as f64;
                            let tx = (tick_x * scale).round() as i32;
                            let _ = MoveToEx(mem_dc, tx, ax_y, None);
                            let _ = LineTo(
                                mem_dc,
                                tx,
                                ((x_axis_y + 5.71) * scale).round() as i32,
                            );
                        }
                        SelectObject(mem_dc, old_axis_pen);
                        let _ = DeleteObject(axis_pen);
                        } else if chart.chart_type == "radar"
                            && std::env::var("OXI_RADAR_DISABLE").is_err()
                        {
                        // Radar / spider chart (Word render-truth 2026-08-10,
                        // 18 arms: chart_radar R1-R8 + chart_radar_geo G1-G10,
                        // fitz get_drawings items + rawdict spans).
                        //
                        //   radar box   top = plot_top + 17.44,
                        //               bottom = sy + shh - 33.44
                        //     plot_top  = sy+16.0 (no title) / sy+51.4 (auto
                        //               title = 1 series) / sy+45.69 (explicit)
                        //   Lmax        = widest WRAPPED category-label LINE
                        //   plot_left   = sx + Lmax + 13.5
                        //   band_right  = legend ? swatch_x0 - 4.64 : sx + sw
                        //   plot_right  = band_right - Lmax - 13.5
                        //   r  = min(box height, box width) / 2   cx,cy = centre
                        //   Verified on the width-limited (G6 197.0/60.67 vs
                        //   197.0/60.73, G7 162.0/38.45 vs 38.47), height-
                        //   limited (G3 270.0/120.69 vs 120.65) and legend-
                        //   limited (G4 235.70/86.39, G5 200.87/57.57) arms.
                        //   n_cat has NO effect on the radius (G1/G8/G9 equal).
                        //
                        //   Category label wrap cap = 0.25*A - 5.6 where
                        //   A = band_right - sx.  Measured windows: A=180 ->
                        //   [38.05,42.10), A=250 -> [50.83,57.80), A=257.73 ->
                        //   [57.80,63.81), A>=327 -> no wrap; together they pin
                        //   the slope to (0.2020,0.2821) and, at 0.25, the
                        //   intercept to (4.70,6.63].
                        //
                        //   Divisions: the ~5-tick nice step on the 5%-headroom
                        //   max, coarsened AT MOST ONCE while r_axis/div < 19.0,
                        //   where r_axis ignores the LEGEND band (it uses the
                        //   label reservation computed against the full frame
                        //   width).  That is what separates R4 (r_axis 92.86 ->
                        //   18.57 -> 3 divisions) from G4 (110.56 -> 22.11 -> 5)
                        //   even though both render r = 86.4; R8 (19.14) and G4
                        //   bracket the threshold from above while R4/R2 (18.57)
                        //   bracket it from below.  Falsified: divisions from
                        //   the final r, from 2r, from the plot height, from
                        //   n_cat, from the series count.
                        //
                        //   Rings are POLYGONS (n_cat vertices) at k*r/div,
                        //   k=1..div, #868686 w=0.75; spokes run centre->vertex
                        //   in the same grey; each series is a CLOSED polygon
                        //   stroked in its border colour at w=3.75 (line charts
                        //   use 2.25).  Word draws NO markers in any arm --
                        //   RADAR_MARKERS (chart5, radarStyle="marker" with no
                        //   c:symbol) has series radii identical to the plain
                        //   arm and no marker geometry at all -- so none are
                        //   drawn here.
                        //
                        //   Value labels: right edge at cx-18.23, baseline
                        //   ring_y+6.24.  Category labels hang off an anchor at
                        //   radius 1.0406*r (measured k = 1.04034..1.04107 over
                        //   arms spanning r 38..121): the label box's LEFT edge
                        //   sits on the anchor on the right half, its RIGHT edge
                        //   on the left half, and it is centred at the exact top
                        //   and bottom vertices; the baseline is anchor+5.25 on
                        //   the sides, anchor-6.70 at the top and anchor+17.15
                        //   at the bottom, and a wrapped label centres its line
                        //   block on that baseline (21.99 line pitch).  Checked
                        //   against every label of all 18 arms: 71 labels, max
                        //   error 0.12pt.
                        //
                        // RESIDUAL (unmeasured, deliberately not implemented):
                        //   RADAR_FILLED's polygon is exported by Word as a
                        //   masked RASTER (Image25, 173x169) carrying a vertical
                        //   GRADIENT (sampled 151,190,252 -> 66,130,206), so
                        //   only a solid accent fill is drawn here.
                        let axis_family = "Calibri";
                        let axis_fs = 18.0f32;
                        let sx = sh.x as f64;
                        let sy = sh.y as f64;
                        let sw = sh.width as f64;
                        let shh = sh.height as f64;
                        let n_cat = chart.categories.len().max(1);
                        let n_ser = chart.series.len().max(1);
                        let has_explicit_title = chart.explicit_title.is_some();
                        let has_auto_title = chart.series.len() == 1
                            && !chart.auto_title_deleted
                            && !has_explicit_title;
                        let filled = chart.grouping == "filled";
                        let wid = |t: &str| {
                            font_adv::line_hmtx_width_pt(t, axis_fs, axis_family)
                                .unwrap_or_else(|| {
                                    t.chars().count() as f32 * axis_fs * 0.5
                                }) as f64
                        };
                        let plot_top = if has_explicit_title {
                            sy + 45.69
                        } else if has_auto_title {
                            sy + 51.4
                        } else {
                            sy + 16.0
                        };
                        let radar_top = plot_top + 17.44;
                        let radar_bot = sy + shh - 33.44;

                        // ---- legend band (line swatch, doughnut block law) ----
                        let legend_fs = 18.0f32;
                        let legend_cap = legend_label_cap(sw);
                        let legend_lines: Vec<Vec<String>> = chart
                            .series
                            .iter()
                            .map(|s| {
                                wrap_legend_label(
                                    &s.name, legend_fs, axis_family, legend_cap,
                                )
                            })
                            .collect();
                        let legend_max_lines =
                            legend_lines.iter().map(|l| l.len()).max().unwrap_or(1);
                        let max_legend_w = legend_lines
                            .iter()
                            .flatten()
                            .map(|l| wid(l))
                            .fold(0.0f64, f64::max);
                        let legend_right = sx + sw - 10.0;
                        let label_x0 = legend_right - max_legend_w;
                        let swatch_x0 = label_x0 - 21.29;
                        let band_right = if chart.has_legend {
                            swatch_x0 - 4.64
                        } else {
                            sx + sw
                        };

                        // ---- category labels + radius ----
                        let wrap_cats = |a: f64| -> (Vec<Vec<String>>, f64) {
                            let cap = 0.25 * a - 5.6;
                            let lines: Vec<Vec<String>> = chart
                                .categories
                                .iter()
                                .map(|c| {
                                    wrap_legend_label(c, axis_fs, axis_family, cap)
                                })
                                .collect();
                            let m = lines
                                .iter()
                                .flatten()
                                .map(|l| wid(l))
                                .fold(0.0f64, f64::max);
                            (lines, m)
                        };
                        let (cat_lines, lmax) = wrap_cats(band_right - sx);
                        let plot_left = sx + lmax + 13.5;
                        let plot_right = band_right - lmax - 13.5;
                        let r = ((radar_bot - radar_top) / 2.0)
                            .min((plot_right - plot_left) / 2.0)
                            .max(1.0);
                        let cx = (plot_left + plot_right) / 2.0;
                        let cy = (radar_top + radar_bot) / 2.0;

                        // ---- value axis (r_axis ignores the legend band) ----
                        let (_, lmax0) = wrap_cats(sw);
                        let r_axis = ((radar_bot - radar_top) / 2.0)
                            .min((sw - 2.0 * (lmax0 + 13.5)) / 2.0)
                            .max(1.0);
                        let max_val = chart
                            .series
                            .iter()
                            .flat_map(|s| s.values.iter().copied())
                            .fold(0.0f64, f64::max);
                        let hi = max_val.max(1e-9) * AXIS_HEADROOM;
                        let raw = hi / 5.0;
                        let mut mag = 10f64.powf(raw.log10().floor());
                        let resid = raw / mag;
                        let mut mult = if resid < 1.5 {
                            1.0
                        } else if resid < 3.0 {
                            2.0
                        } else if resid < 7.0 {
                            5.0
                        } else {
                            mag *= 10.0;
                            1.0
                        };
                        let mut step = mult * mag;
                        let mut axis_max = (hi / step - 1e-9).ceil() * step;
                        let mut div = (axis_max / step).round().max(1.0);
                        if div > 1.0 && r_axis / div < 19.0 {
                            if mult == 1.0 {
                                mult = 2.0;
                            } else if mult == 2.0 {
                                mult = 5.0;
                            } else {
                                mult = 1.0;
                                mag *= 10.0;
                            }
                            step = mult * mag;
                            axis_max = (hi / step - 1e-9).ceil() * step;
                            div = (axis_max / step).round().max(1.0);
                        }
                        let div_n = div as usize;
                        let vert = |i: usize, rad: f64| -> (f64, f64) {
                            let th = i as f64 * std::f64::consts::TAU / n_cat as f64;
                            (cx + rad * th.sin(), cy - rad * th.cos())
                        };
                        let dev = |v: f64| (v * scale).round() as i32;

                        // ---- rings + spokes ----
                        let grid_pen = CreatePen(
                            PS_SOLID,
                            (0.75 * scale).round().max(1.0) as i32,
                            COLORREF(colorref(0x86, 0x86, 0x86)),
                        );
                        let old_grid_pen = SelectObject(mem_dc, grid_pen);
                        let _ = SelectObject(mem_dc, GetStockObject(NULL_BRUSH));
                        for k in 1..=div_n {
                            let rr = r * k as f64 / div;
                            let (x0, y0) = vert(0, rr);
                            let _ = MoveToEx(mem_dc, dev(x0), dev(y0), None);
                            for i in 1..=n_cat {
                                let (x, y) = vert(i % n_cat, rr);
                                let _ = LineTo(mem_dc, dev(x), dev(y));
                            }
                        }
                        for i in 0..n_cat {
                            let (x, y) = vert(i, r);
                            let _ = MoveToEx(mem_dc, dev(cx), dev(cy), None);
                            let _ = LineTo(mem_dc, dev(x), dev(y));
                        }
                        SelectObject(mem_dc, old_grid_pen);
                        let _ = DeleteObject(grid_pen);

                        // ---- series polygons ----
                        for (si, s) in chart.series.iter().enumerate() {
                            let pts: Vec<(f64, f64)> = (0..n_cat)
                                .map(|i| {
                                    let v = s.values.get(i).copied().unwrap_or(0.0);
                                    vert(i, r * (v / axis_max).max(0.0))
                                })
                                .collect();
                            if pts.len() < 2 {
                                continue;
                            }
                            let (fill_hex, border_hex) = line_series_colors(si);
                            if filled {
                                if let Some(rgb) = parse_hex_rgb(&fill_hex) {
                                    use windows::Win32::Foundation::POINT;
                                    use windows::Win32::Graphics::Gdi::Polygon;
                                    let brush = CreateSolidBrush(COLORREF(colorref(
                                        rgb.0, rgb.1, rgb.2,
                                    )));
                                    let old_brush = SelectObject(mem_dc, brush);
                                    let _ =
                                        SelectObject(mem_dc, GetStockObject(NULL_PEN));
                                    let poly: Vec<POINT> = pts
                                        .iter()
                                        .map(|(x, y)| POINT {
                                            x: dev(*x),
                                            y: dev(*y),
                                        })
                                        .collect();
                                    let _ = Polygon(mem_dc, &poly);
                                    SelectObject(mem_dc, old_brush);
                                    let _ = DeleteObject(brush);
                                }
                            }
                            if let Some(rgb) = parse_hex_rgb(&border_hex) {
                                let pen = CreatePen(
                                    PS_SOLID,
                                    (3.75 * scale).round().max(1.0) as i32,
                                    COLORREF(colorref(rgb.0, rgb.1, rgb.2)),
                                );
                                let old_pen = SelectObject(mem_dc, pen);
                                let _ = SelectObject(mem_dc, GetStockObject(NULL_BRUSH));
                                let _ =
                                    MoveToEx(mem_dc, dev(pts[0].0), dev(pts[0].1), None);
                                for k in 1..=pts.len() {
                                    let q = pts[k % pts.len()];
                                    let _ = LineTo(mem_dc, dev(q.0), dev(q.1));
                                }
                                SelectObject(mem_dc, old_pen);
                                let _ = DeleteObject(pen);
                            }
                        }

                        // ---- value-axis labels ----
                        for k in 0..=div_n {
                            let v = axis_max * k as f64 / div;
                            let label = fmt_axis_value(v, step);
                            let lw = wid(&label);
                            let ty = cy - r * k as f64 / div;
                            draw_text_baseline(
                                mem_dc,
                                dev(cx - 18.23 - lw),
                                (ty + 6.24) as f32,
                                &label,
                                axis_fs,
                                axis_family,
                                None,
                                scale,
                            );
                        }

                        // ---- category labels ----
                        for (i, lines) in cat_lines.iter().enumerate() {
                            let th = i as f64 * std::f64::consts::TAU / n_cat as f64;
                            let (ax, ay) = vert(i, r * 1.0406);
                            let bw = lines.iter().map(|l| wid(l)).fold(0.0f64, f64::max);
                            let sn = th.sin();
                            let box_x0 = if sn > 1e-6 {
                                ax
                            } else if sn < -1e-6 {
                                ax - bw
                            } else {
                                ax - bw / 2.0
                            };
                            let base = if sn.abs() < 1e-6 {
                                ay + if th.cos() > 0.0 { -6.70 } else { 17.15 }
                            } else {
                                ay + 5.25
                            };
                            let first = base - (lines.len() as f64 - 1.0) * 21.99 / 2.0;
                            for (li, line) in lines.iter().enumerate() {
                                let lw = wid(line);
                                draw_text_baseline(
                                    mem_dc,
                                    dev(box_x0 + (bw - lw) / 2.0),
                                    (first + li as f64 * 21.99) as f32,
                                    line,
                                    axis_fs,
                                    axis_family,
                                    None,
                                    scale,
                                );
                            }
                        }

                        // ---- legend ----
                        if chart.has_legend {
                            let text_line_pitch = 21.99f64;
                            let row_pitch =
                                27.75 + (legend_max_lines as f64 - 1.0) * 21.76;
                            let block_h = n_ser as f64 * row_pitch;
                            let legend_y0 = cy - block_h / 2.0 + 8.97;
                            for (si, lines) in legend_lines.iter().enumerate() {
                                let sw_y = legend_y0 + si as f64 * row_pitch;
                                let (_, border_hex) = line_series_colors(si);
                                if let Some(rgb) = parse_hex_rgb(&border_hex) {
                                    let lg_pen = CreatePen(
                                        PS_SOLID,
                                        (2.25 * scale).round().max(1.0) as i32,
                                        COLORREF(colorref(rgb.0, rgb.1, rgb.2)),
                                    );
                                    let old_lg_pen = SelectObject(mem_dc, lg_pen);
                                    let _ =
                                        SelectObject(mem_dc, GetStockObject(NULL_BRUSH));
                                    let ly = dev(sw_y + 4.945);
                                    let _ = MoveToEx(mem_dc, dev(swatch_x0), ly, None);
                                    let _ = LineTo(mem_dc, dev(swatch_x0 + 19.20), ly);
                                    SelectObject(mem_dc, old_lg_pen);
                                    let _ = DeleteObject(lg_pen);
                                }
                                for (li, line) in lines.iter().enumerate() {
                                    draw_text_baseline(
                                        mem_dc,
                                        dev(label_x0),
                                        (sw_y + 10.17 + li as f64 * text_line_pitch)
                                            as f32,
                                        line,
                                        legend_fs,
                                        axis_family,
                                        None,
                                        scale,
                                    );
                                }
                            }
                        }

                        // ---- titles ----
                        if let Some(title) = &chart.explicit_title {
                            let tfs = 18.0f32;
                            let lw = font_adv::line_hmtx_width_pt(title, tfs, "Arial")
                                .unwrap_or_else(|| {
                                    title.chars().count() as f32 * tfs * 0.5
                                })
                                as f64;
                            draw_text_baseline_w(
                                mem_dc,
                                dev(sx + sw / 2.0 - lw / 2.0),
                                (sy + 24.43) as f32,
                                title,
                                tfs,
                                "Arial",
                                None,
                                scale,
                                400,
                            );
                        } else if has_auto_title {
                            let first = &chart.series[0];
                            let tfs = 21.62f32;
                            let lw = font_adv::line_hmtx_width_pt(
                                &first.name, tfs, axis_family,
                            )
                            .unwrap_or_else(|| {
                                first.name.chars().count() as f32 * tfs * 0.5
                            }) as f64;
                            draw_text_baseline_w(
                                mem_dc,
                                dev(sx + sw / 2.0 - lw / 2.0),
                                (sy + 28.03) as f32,
                                &first.name,
                                tfs,
                                axis_family,
                                None,
                                scale,
                                700,
                            );
                        }
                        } else if chart.chart_type == "stock"
                            && std::env::var("OXI_STOCK_DISABLE").is_err()
                        {
                        // Stock chart (Word render-truth 2026-08-10, the
                        // 8-arm chart_stock probe K1..K8 read with fitz
                        // get_drawings + rawdict).  <c:stockChart> reuses the
                        // LINE chart's plot geometry; what is new is that the
                        // series carry <a:ln><a:noFill/> (nothing joins the
                        // points) and the data is carried by two decorations:
                        //   <c:hiLowLines/>  one vertical rule per category
                        //                    spanning min..max ACROSS ALL
                        //                    SERIES  (K1 Q1: High 24.0 ->
                        //                    134.40, Low 18.2 -> 179.28)
                        //   <c:upDownBars>   a box between the FIRST and LAST
                        //                    series (open..close), width
                        //                    pitch/(1+gapWidth/100) centred on
                        //                    the band centre (34.32 at pitch
                        //                    85.89 and 27.20 at pitch 68.00,
                        //                    both exact at the default 150)
                        //
                        //   plot_left  = sx + 6.50 + w(widest value label)
                        //                + 16.70  (= 113.45 on every arm; the
                        //                same gutter the area / horizontal-bar
                        //                branches use, and the value the LINE
                        //                branch hardcodes as sx+41.4 for its
                        //                own two-digit labels)
                        //   plot_right = legend ? label_x0 - 32.66
                        //                       : sx + sw - 11.0
                        //                (32.66 = swatch 9.89 + gap 4.62 +
                        //                18.15, the doughnut/area legend band;
                        //                K2 386.01 / K7 385.44 measured)
                        //   plot_top   = sy + 16.0, or sy + 45.69 with an
                        //                explicit <c:title>  (a stock chart
                        //                always has >= 3 series so the
                        //                single-series auto title never fires)
                        //   plot_bot   = sy + shh - 39.9
                        //   data x     = plot_left + pitch*(i+0.5)  (band
                        //                centres, NOT category boundaries --
                        //                so plot_right is NOT pulled in by
                        //                half a label the way area/bar are)
                        let axis_family = "Calibri";
                        let axis_fs = 18.0f32;
                        let sx = sh.x as f64;
                        let sy = sh.y as f64;
                        let sw = sh.width as f64;
                        let shh = sh.height as f64;
                        let has_explicit_title = chart.explicit_title.is_some();
                        let tw = |s: &str, fs: f32| -> f64 {
                            font_adv::line_hmtx_width_pt(s, fs, axis_family)
                                .unwrap_or_else(|| {
                                    s.chars().count() as f32 * fs * 0.5
                                }) as f64
                        };
                        let max_val = chart
                            .series
                            .iter()
                            .flat_map(|s| s.values.iter().copied())
                            .fold(0.0f64, f64::max);
                        let (max_axis, axis_steps) = nice_axis_max_div(max_val);
                        let step = max_axis / axis_steps as f64;
                        let mut widest_val = 0.0f64;
                        for i in 0..=axis_steps {
                            let w = tw(&fmt_axis_value(step * i as f64, step), axis_fs);
                            if w > widest_val {
                                widest_val = w;
                            }
                        }
                        let plot_left = sx + 6.50 + widest_val + 16.70;
                        let plot_top = if has_explicit_title {
                            sy + 45.69
                        } else {
                            sy + 16.0
                        };
                        let plot_bot = sy + shh - 39.9;
                        let plot_h = plot_bot - plot_top;
                        let max_label_w = chart
                            .series
                            .iter()
                            .map(|s| tw(&s.name, axis_fs))
                            .fold(0.0f64, f64::max);
                        let legend_label_x0 = sx + sw - 10.0 - max_label_w;
                        let plot_right = if chart.has_legend {
                            legend_label_x0 - 32.66
                        } else {
                            sx + sw - 11.0
                        };
                        let plot_w = plot_right - plot_left;
                        let n_cat = chart.categories.len().max(1);
                        let pitch = plot_w / n_cat as f64;
                        let val_y =
                            |v: f64| plot_bot - (v / max_axis.max(1e-9)) * plot_h;
                        let pen_w = (0.75 * scale).round().max(1.0) as i32;

                        // Major gridlines i=1..=axis_steps (i=0 is the X axis
                        // line, i=axis_steps is the plot's top edge), drawn
                        // UNDER the hi-low rules and the up/down boxes.
                        let grid_pen =
                            CreatePen(PS_SOLID, pen_w, COLORREF(colorref(0, 0, 0)));
                        let old_grid_pen = SelectObject(mem_dc, grid_pen);
                        let _ = SelectObject(mem_dc, GetStockObject(NULL_BRUSH));
                        let gl = (plot_left * scale).round() as i32;
                        let gr = (plot_right * scale).round() as i32;
                        for i in 1..=axis_steps {
                            let gy = ((plot_bot - plot_h * i as f64
                                / axis_steps as f64)
                                * scale)
                                .round() as i32;
                            let _ = MoveToEx(mem_dc, gl, gy, None);
                            let _ = LineTo(mem_dc, gr, gy);
                        }
                        SelectObject(mem_dc, old_grid_pen);
                        let _ = DeleteObject(grid_pen);

                        // <c:hiLowLines/>: one black w=0.75 vertical rule per
                        // category, from the MAXIMUM to the MINIMUM of every
                        // series at that category.
                        if chart.hi_low_lines && !chart.series.is_empty() {
                            let hl_pen = CreatePen(
                                PS_SOLID,
                                pen_w,
                                COLORREF(colorref(0, 0, 0)),
                            );
                            let old_hl_pen = SelectObject(mem_dc, hl_pen);
                            let _ =
                                SelectObject(mem_dc, GetStockObject(NULL_BRUSH));
                            for ci in 0..n_cat {
                                let mut lo = f64::INFINITY;
                                let mut hi = f64::NEG_INFINITY;
                                for s in chart.series.iter() {
                                    if let Some(v) = s.values.get(ci) {
                                        lo = lo.min(*v);
                                        hi = hi.max(*v);
                                    }
                                }
                                if !lo.is_finite() || !hi.is_finite() {
                                    continue;
                                }
                                let cx = ((plot_left + pitch * (ci as f64 + 0.5))
                                    * scale)
                                    .round() as i32;
                                let _ = MoveToEx(
                                    mem_dc,
                                    cx,
                                    (val_y(hi) * scale).round() as i32,
                                    None,
                                );
                                let _ = LineTo(
                                    mem_dc,
                                    cx,
                                    (val_y(lo) * scale).round() as i32,
                                );
                            }
                            SelectObject(mem_dc, old_hl_pen);
                            let _ = DeleteObject(hl_pen);
                        }

                        // <c:upDownBars>: a box from the FIRST series' value
                        // (open) to the LAST series' value (close).  Word
                        // paints it #F9F9F9 when the close is ABOVE the open
                        // and #3F3F3F when below, both with a black w=0.75
                        // outline, and draws it OVER the hi-low rules.
                        if chart.up_down_bars && chart.series.len() >= 2 {
                            let bar_w = pitch / (1.0 + chart.up_down_gap / 100.0);
                            let first = &chart.series[0];
                            let last = &chart.series[chart.series.len() - 1];
                            let ud_pen = CreatePen(
                                PS_SOLID,
                                pen_w,
                                COLORREF(colorref(0, 0, 0)),
                            );
                            for ci in 0..n_cat {
                                let (o, c) = match (
                                    first.values.get(ci),
                                    last.values.get(ci),
                                ) {
                                    (Some(a), Some(b)) => (*a, *b),
                                    _ => continue,
                                };
                                let cx = plot_left + pitch * (ci as f64 + 0.5);
                                let y0 = val_y(o.max(c));
                                let y1 = val_y(o.min(c));
                                let r = RECT {
                                    left: ((cx - bar_w / 2.0) * scale).round() as i32,
                                    top: (y0 * scale).round() as i32,
                                    right: ((cx + bar_w / 2.0) * scale).round() as i32,
                                    bottom: (y1 * scale).round() as i32,
                                };
                                let fill = if c >= o {
                                    colorref(0xf9, 0xf9, 0xf9)
                                } else {
                                    colorref(0x3f, 0x3f, 0x3f)
                                };
                                let brush = CreateSolidBrush(COLORREF(fill));
                                let _ = FillRect(mem_dc, &r, brush);
                                let _ = DeleteObject(brush);
                                let old_ud_pen = SelectObject(mem_dc, ud_pen);
                                let _ = SelectObject(
                                    mem_dc,
                                    GetStockObject(NULL_BRUSH),
                                );
                                let _ = MoveToEx(mem_dc, r.left, r.top, None);
                                let _ = LineTo(mem_dc, r.right, r.top);
                                let _ = LineTo(mem_dc, r.right, r.bottom);
                                let _ = LineTo(mem_dc, r.left, r.bottom);
                                let _ = LineTo(mem_dc, r.left, r.top);
                                SelectObject(mem_dc, old_ud_pen);
                            }
                            let _ = DeleteObject(ud_pen);
                        }

                        // Connecting polylines / markers: a stock chart's
                        // series declare <a:ln><a:noFill/> and
                        // <c:symbol val="none"/>, so nothing is drawn; the
                        // per-series flags keep that honest rather than
                        // hard-coding "stock never draws lines".
                        let series_pts: Vec<Vec<(f64, f64)>> = chart
                            .series
                            .iter()
                            .map(|s| {
                                s.values
                                    .iter()
                                    .enumerate()
                                    .map(|(ci, v)| {
                                        (
                                            plot_left + pitch * (ci as f64 + 0.5),
                                            val_y(*v),
                                        )
                                    })
                                    .collect()
                            })
                            .collect();
                        for (si, pts) in series_pts.iter().enumerate() {
                            if chart.series[si].line_none {
                                continue;
                            }
                            let (_, border_hex) = line_series_colors(si);
                            if let Some(rgb) = parse_hex_rgb(&border_hex) {
                                let line_pen = CreatePen(
                                    PS_SOLID,
                                    (2.25 * scale).round().max(1.0) as i32,
                                    COLORREF(colorref(rgb.0, rgb.1, rgb.2)),
                                );
                                let old_line_pen = SelectObject(mem_dc, line_pen);
                                let _ =
                                    SelectObject(mem_dc, GetStockObject(NULL_BRUSH));
                                for w in pts.windows(2) {
                                    let _ = MoveToEx(
                                        mem_dc,
                                        (w[0].0 * scale).round() as i32,
                                        (w[0].1 * scale).round() as i32,
                                        None,
                                    );
                                    let _ = LineTo(
                                        mem_dc,
                                        (w[1].0 * scale).round() as i32,
                                        (w[1].1 * scale).round() as i32,
                                    );
                                }
                                SelectObject(mem_dc, old_line_pen);
                                let _ = DeleteObject(line_pen);
                            }
                        }
                        if chart.marker {
                            for (si, pts) in series_pts.iter().enumerate() {
                                if chart.series[si].marker_none {
                                    continue;
                                }
                                let (fill_hex, _) = line_series_colors(si);
                                if let Some(rgb) = parse_hex_rgb(&fill_hex) {
                                    let m_brush = CreateSolidBrush(COLORREF(
                                        colorref(rgb.0, rgb.1, rgb.2),
                                    ));
                                    let old_m_brush = SelectObject(mem_dc, m_brush);
                                    let _ =
                                        SelectObject(mem_dc, GetStockObject(NULL_PEN));
                                    for (px, py) in pts.iter() {
                                        draw_line_marker(
                                            mem_dc,
                                            line_marker_shape(si),
                                            *px,
                                            *py,
                                            6.96 / 2.0,
                                            scale,
                                        );
                                    }
                                    SelectObject(mem_dc, old_m_brush);
                                    let _ = DeleteObject(m_brush);
                                }
                            }
                        }

                        // Value labels: Calibri 18pt, right edge =
                        // plot_left-16.64, baseline = tick_y+5.22 (same rule
                        // as the line/bar branches).
                        for i in 0..=axis_steps {
                            let val = step * i as f64;
                            let label = fmt_axis_value(val, step);
                            let lw = tw(&label, axis_fs);
                            draw_text_baseline(
                                mem_dc,
                                ((plot_left - 16.64 - lw) * scale).round() as i32,
                                (val_y(val) + 5.22) as f32,
                                &label,
                                axis_fs,
                                axis_family,
                                None,
                                scale,
                            );
                        }

                        // Category names centred under each band.
                        for (ci, name) in chart.categories.iter().enumerate() {
                            let cx = plot_left + pitch * (ci as f64 + 0.5);
                            let lw = tw(name, axis_fs);
                            draw_text_baseline(
                                mem_dc,
                                ((cx - lw / 2.0) * scale).round() as i32,
                                (plot_bot + 28.67) as f32,
                                name,
                                axis_fs,
                                axis_family,
                                None,
                                scale,
                            );
                        }

                        // Axis lines + ticks: Y at plot_left, X at plot_bot,
                        // Y ticks i=0..=axis_steps hanging 5.71 to the left,
                        // X ticks i=0..=n_cat hanging 5.71 below.
                        let axis_pen =
                            CreatePen(PS_SOLID, pen_w, COLORREF(colorref(0, 0, 0)));
                        let old_axis_pen = SelectObject(mem_dc, axis_pen);
                        let _ = SelectObject(mem_dc, GetStockObject(NULL_BRUSH));
                        let pl = (plot_left * scale).round() as i32;
                        let pt = (plot_top * scale).round() as i32;
                        let pr = (plot_right * scale).round() as i32;
                        let pb = (plot_bot * scale).round() as i32;
                        let _ = MoveToEx(mem_dc, pl, pt, None);
                        let _ = LineTo(mem_dc, pl, pb);
                        let _ = MoveToEx(mem_dc, pl, pb, None);
                        let _ = LineTo(mem_dc, pr, pb);
                        for i in 0..=axis_steps {
                            let ty = ((plot_bot - plot_h * i as f64
                                / axis_steps as f64)
                                * scale)
                                .round() as i32;
                            let _ = MoveToEx(
                                mem_dc,
                                ((plot_left - 5.71) * scale).round() as i32,
                                ty,
                                None,
                            );
                            let _ = LineTo(mem_dc, pl, ty);
                        }
                        for i in 0..=n_cat {
                            let tx =
                                ((plot_left + pitch * i as f64) * scale).round() as i32;
                            let _ = MoveToEx(mem_dc, tx, pb, None);
                            let _ = LineTo(
                                mem_dc,
                                tx,
                                ((plot_bot + 5.71) * scale).round() as i32,
                            );
                        }
                        SelectObject(mem_dc, old_axis_pen);
                        let _ = DeleteObject(axis_pen);

                        // Legend: the doughnut/area block law -- n rows of
                        // 27.75pt centred on sy + shh/2 (+14.85 with an
                        // explicit title), first swatch top = block top +
                        // 8.97, label baseline = swatch bottom + 0.28, all
                        // labels left-aligned at
                        // label_x0 = sx + sw - 10 - max_label_w.
                        // (K2 predicted 193.52 vs measured 193.46; K7
                        // 194.49 vs 194.45.)  No swatch is painted for a
                        // stock series because its line is noFill and its
                        // marker is none -- Word draws labels only.
                        if chart.has_legend {
                            let n = chart.series.len();
                            let title_off = if has_explicit_title { 14.85 } else { 0.0 };
                            let block_top = sy + shh / 2.0 + title_off
                                - n as f64 * 27.75 / 2.0;
                            for (si, s) in chart.series.iter().enumerate() {
                                let sw_top = block_top + 8.97 + si as f64 * 27.75;
                                if !s.line_none {
                                    let (_, border_hex) = line_series_colors(si);
                                    if let Some(rgb) = parse_hex_rgb(&border_hex) {
                                        let lg_pen = CreatePen(
                                            PS_SOLID,
                                            (2.25 * scale).round().max(1.0) as i32,
                                            COLORREF(colorref(rgb.0, rgb.1, rgb.2)),
                                        );
                                        let old_lg_pen = SelectObject(mem_dc, lg_pen);
                                        let _ = SelectObject(
                                            mem_dc,
                                            GetStockObject(NULL_BRUSH),
                                        );
                                        let ly = ((sw_top + 9.89 / 2.0) * scale)
                                            .round() as i32;
                                        let lx0 = ((legend_label_x0 - 4.62 - 9.89)
                                            * scale)
                                            .round() as i32;
                                        let lx1 =
                                            ((legend_label_x0 - 4.62) * scale).round()
                                                as i32;
                                        let _ = MoveToEx(mem_dc, lx0, ly, None);
                                        let _ = LineTo(mem_dc, lx1, ly);
                                        SelectObject(mem_dc, old_lg_pen);
                                        let _ = DeleteObject(lg_pen);
                                    }
                                }
                                draw_text_baseline(
                                    mem_dc,
                                    (legend_label_x0 * scale).round() as i32,
                                    (sw_top + 9.89 + 0.28) as f32,
                                    &s.name,
                                    axis_fs,
                                    axis_family,
                                    None,
                                    scale,
                                );
                            }
                        }

                        // EXPLICIT <c:title>: Arial 18pt regular, centred on
                        // the frame, baseline sy+24.43 (same as every other
                        // chart type).
                        if let Some(title) = &chart.explicit_title {
                            let tfs = 18.0f32;
                            let lw = font_adv::line_hmtx_width_pt(title, tfs, "Arial")
                                .unwrap_or_else(|| {
                                    title.chars().count() as f32 * tfs * 0.5
                                }) as f64;
                            draw_text_baseline_w(
                                mem_dc,
                                ((sx + sw / 2.0 - lw / 2.0) * scale).round() as i32,
                                (sy + 24.43) as f32,
                                title,
                                tfs,
                                "Arial",
                                None,
                                scale,
                                400,
                            );
                        }
                        } else if chart.chart_type == "line" {
                        // Line chart (Word render-truth 2026-08-06, fitz
                        // get_drawings + rawdict on the 7-variant probe
                        // P0-P6; chart_line = the P1 configuration):
                        //   plot area: left = sx+41.4, top = sy+51.4
                        //     (1 series -> auto title present), bottom =
                        //     sy+shh-39.9 (default band; the crowded-label
                        //     78.62 band is derived-but-unimplemented =
                        //     two-line label wrapping not yet rendered).
                        //     plot_right = WITH a legend: sx+sw-41.4-103.82
                        //     (right legend band), without: sx+sw-41.4-11
                        //     (same as the bars).
                        //   category pitch = plot_w/n_cat; point i x =
                        //     plot_left + pitch/2 + i*pitch
                        //   value scale = nice_axis_max(max) in 5 steps
                        //     (6 labels 0,5,...,25); point i y = plot_bot -
                        //     (val_i/max_axis)*plot_h
                        //   gridlines: 5 black lines at the value ticks
                        //     i=1..=axis_steps, drawn BEFORE the polyline
                        //     (the chart_line valAx declares
                        //     <c:majorGridlines/>; line charts always carry
                        //     gridlines - NOT gated on is_stacked)
                        //   polyline: (n_cat-1) accent1 #4F81BD segments
                        //     w=2.25 joining consecutive points
                        //   markers: 6.96pt filled accent1 circles at every
                        //     point, gated on <c:marker val="1"/>
                        //     (LINE_MARKERS)
                        //   legend (when <c:legend> declared) is a LINE
                        //     swatch: horizontal accent1 line w=2.25 of
                        //     length 19.20 at legend_left=plot_right+15.65,
                        //     a 6.96pt marker circle centred on it, and the
                        //     series name Calibri 18pt beside it;
                        //     legend_y0 = sy + shh/2 + 17.68 (FRAME-relative,
                        //     P0/P1/P4 render-truth)
                        let axis_family = "Calibri";
                        let axis_fs = 18.0f32;
                        let sx = sh.x as f64;
                        let sy = sh.y as f64;
                        let sw = sh.width as f64;
                        let shh = sh.height as f64;
                        let has_auto_title = chart.series.len() == 1;
                        let has_explicit_title = chart.explicit_title.is_some();
                        // NEGATIVE data (chart_negative N3, 2026-08-10): the
                        // value axis spans zero, the category names hang off
                        // the ZERO line instead of the plot bottom, and the
                        // value gutter grows by the minus sign.
                        let (min_val, max_val) = chart
                            .series
                            .iter()
                            .flat_map(|s| s.values.iter().copied())
                            .fold((0.0f64, 0.0f64), |(lo, hi), v| {
                                (lo.min(v), hi.max(v))
                            });
                        let has_neg = min_val < 0.0;
                        let plot_left_0 = sx + 41.4;
                        let plot_top = if has_explicit_title {
                            // An explicit <c:title> shifts the plot down by
                            // the title line: plot_top = sy+45.69
                            // (chart_title_line/chart_title_line2 render-truth
                            // 2026-08-07; Arial 18pt title, same as the bar
                            // explicit title, vs the auto title's 21.62pt
                            // Calibri-Bold at sy+51.4).
                            sy + 45.69
                        } else if has_auto_title {
                            sy + 51.4
                        } else {
                            sy + 16.0
                        };
                        // plot_w (measured P1..P6): WITH a legend the right
                        // band is 103.82pt (frame-width independent),
                        // WITHOUT it the frame right inset is 11.0pt; then
                        // plot_right = plot_left + plot_w (P3 no-legend 396w
                        // -> 457.00 = 113.45 + 343.6; do NOT subtract the
                        // 41.4 left inset a second time).
                        // The right band is 103.82 = 15.65 (swatch gap) + 19.20
                        // (swatch) + 2.09 + w("Series 1") + 10.0 (frame inset);
                        // with a different series name it must be computed
                        // (N3's "Ser1" gives 79.61, plot_right 388.38 = Word
                        // 388.37).  The positive path keeps the measured
                        // constant so every existing line probe is unchanged.
                        let legend_band = if !chart.has_legend {
                            11.0
                        } else if has_neg {
                            let w = chart
                                .series
                                .iter()
                                .map(|s| {
                                    font_adv::line_hmtx_width_pt(
                                        &s.name, axis_fs, axis_family,
                                    )
                                    .unwrap_or_else(|| {
                                        s.name.chars().count() as f32 * axis_fs * 0.5
                                    }) as f64
                                })
                                .fold(0.0f64, f64::max);
                            15.65 + 21.29 + w + 10.0
                        } else {
                            103.82
                        };
                        let plot_right = sx + sw - legend_band;
                        let plot_w = plot_right - plot_left_0;
                        let n_cat = chart.categories.len().max(1);
                        // Crowded category labels -> the 78.62pt bottom band
                        // (measured P0/P4): when the widest category label
                        // EXCEEDS the category pitch (ratio > 1.0) Word keeps
                        // the labels on ONE line by growing the label band;
                        // else the normal 39.9pt band. Threshold re-derived
                        // from the 7 probe variants P0..P6 using the REAL
                        // Calibri hmtx label widths: P0/P4 ratio 1.273 ->
                        // crowded, P3 0.929 / P2 0.900 / P6 0.891 / P1 0.764 /
                        // P5 0.208 -> not crowded. The old window (0.64,0.88]
                        // was a knife-edge misfit (it used a 0.5x-char
                        // fallback width for the unsupported Calibri family).
                        let widest_label = chart
                            .categories
                            .iter()
                            .map(|c| {
                                font_adv::line_hmtx_width_pt(c, 18.0, axis_family)
                                    .unwrap_or_else(|| {
                                        c.chars().count() as f32 * 18.0 * 0.5
                                    }) as f64
                            })
                            .fold(0.0f64, f64::max);
                        let crowded = widest_label / (plot_w / n_cat as f64) > 1.0;
                        // With negatives the category names sit ON the zero
                        // line, so the label band collapses to the plain
                        // 16.0pt margin (N3: plot_bot 344.01 = sy+shh-16).
                        let plot_bot = if has_neg {
                            sy + shh - 16.0
                        } else if crowded {
                            sy + shh - 78.62
                        } else {
                            sy + shh - 39.9
                        };
                        let plot_h = plot_bot - plot_top;
                        // A positive-only line chart always uses 5 divisions
                        // (measured); a signed one needs the general
                        // zero-spanning search (N3: 25..-10 = 7 divisions).
                        let (axis_min, max_axis, axis_steps) = if has_neg {
                            nice_axis_range(min_val, max_val, plot_h, VERT_MIN_SPACING)
                        } else {
                            (0.0, nice_axis_max(max_val), 5usize)
                        };
                        let axis_span = (max_axis - axis_min).max(1e-9);
                        let val_y =
                            |v: f64| plot_bot - ((v - axis_min) / axis_span) * plot_h;
                        let zero_y = val_y(0.0);
                        let plot_left = if has_neg {
                            let mut widest = 0.0f64;
                            for i in 0..=axis_steps {
                                let val =
                                    axis_min + axis_span * i as f64 / axis_steps as f64;
                                let label = format!("{}", val.round() as i64);
                                let lw = font_adv::line_hmtx_width_pt(
                                    &label, axis_fs, axis_family,
                                )
                                .unwrap_or_else(|| {
                                    label.chars().count() as f32 * axis_fs * 0.5
                                }) as f64;
                                widest = widest.max(lw);
                            }
                            sx + 6.5 + widest + 16.7
                        } else {
                            plot_left_0
                        };
                        let plot_w = plot_right - plot_left;
                        let pitch = plot_w / n_cat as f64;

                        // Value axis labels (Calibri 18pt, right edge =
                        // plot_left-16.64, baseline = tick_y+5.22; same
                        // rule as the bars).
                        for i in 0..=axis_steps {
                            let val = axis_min + axis_span * i as f64 / axis_steps as f64;
                            let tick_y = val_y(val);
                            let label = format!("{}", val.round() as i64);
                            let lw = font_adv::line_hmtx_width_pt(&label, 18.0, axis_family)
                                .unwrap_or_else(|| {
                                    label.chars().count() as f32 * 18.0 * 0.5
                                }) as f64;
                            let lx = plot_left - 16.64 - lw;
                            draw_text_baseline(
                                mem_dc,
                                (lx * scale).round() as i32,
                                (tick_y + 5.22) as f32,
                                &label,
                                18.0,
                                axis_family,
                                None,
                                scale,
                            );
                        }

                        // Major gridlines: 5 black lines at the value ticks
                        // (i=0 is the X axis line, drawn in the axis
                        // section below), BEFORE the polyline.
                        let grid_pen = CreatePen(
                            PS_SOLID,
                            2,
                            COLORREF(colorref(0, 0, 0)),
                        );
                        let old_grid_pen = SelectObject(mem_dc, grid_pen);
                        let _ = SelectObject(mem_dc, GetStockObject(NULL_BRUSH));
                        let gl = (plot_left * scale).round() as i32;
                        let gr = (plot_right * scale).round() as i32;
                        for i in 1..=axis_steps {
                            let grid_y = plot_bot - plot_h * i as f64 / axis_steps as f64;
                            let gy = (grid_y * scale).round() as i32;
                            let _ = MoveToEx(mem_dc, gl, gy, None);
                            let _ = LineTo(mem_dc, gr, gy);
                        }
                        SelectObject(mem_dc, old_grid_pen);
                        let _ = DeleteObject(grid_pen);

                        // Data points PER SERIES (point i x =
                        // plot_left + pitch/2 + i*pitch, point i y =
                        // plot_bot - (val_i/max_axis)*plot_h).
                        let series_pts: Vec<Vec<(f64, f64)>> = chart
                            .series
                            .iter()
                            .map(|s| {
                                s.values
                                    .iter()
                                    .enumerate()
                                    .map(|(ci, v)| {
                                        let x = plot_left + pitch * (ci as f64 + 0.5);
                                        (x, val_y(*v))
                                    })
                                    .collect()
                            })
                            .collect();

                        // Polylines: PER SERIES (n_cat-1) border-colour
                        // segments w=2.25 joining consecutive points
                        // (measured border colours: S1 #4A7EBB / S2
                        // #BE4B48 / S3 #98B954).
                        for (si, pts) in series_pts.iter().enumerate() {
                            let (_, border_hex) = line_series_colors(si);
                            if let Some(rgb) = parse_hex_rgb(&border_hex) {
                                let line_pen = CreatePen(
                                    PS_SOLID,
                                    (2.25 * scale).round().max(1.0) as i32,
                                    COLORREF(colorref(rgb.0, rgb.1, rgb.2)),
                                );
                                let old_line_pen = SelectObject(mem_dc, line_pen);
                                let _ = SelectObject(mem_dc, GetStockObject(NULL_BRUSH));
                                for w in pts.windows(2) {
                                    let _ = MoveToEx(
                                        mem_dc,
                                        (w[0].0 * scale).round() as i32,
                                        (w[0].1 * scale).round() as i32,
                                        None,
                                    );
                                    let _ = LineTo(
                                        mem_dc,
                                        (w[1].0 * scale).round() as i32,
                                        (w[1].1 * scale).round() as i32,
                                    );
                                }
                                SelectObject(mem_dc, old_line_pen);
                                let _ = DeleteObject(line_pen);
                            }
                        }

                        // Markers: 6.96pt filled PER-SERIES shapes at every
                        // point (gated on <c:marker val="1"/>). Word renders
                        // the LINE_MARKERS data marker as a per-series shape
                        // (S1 diamond / S2 square / S3 triangle, measured)
                        // filled with the series fill colour; same-colour
                        // stroke (w=0.75) so the outline is invisible.
                        if chart.marker {
                            for (si, pts) in series_pts.iter().enumerate() {
                                let (fill_hex, _) = line_series_colors(si);
                                if let Some(rgb) = parse_hex_rgb(&fill_hex) {
                                    let m_brush = CreateSolidBrush(COLORREF(
                                        colorref(rgb.0, rgb.1, rgb.2),
                                    ));
                                    let old_m_brush = SelectObject(mem_dc, m_brush);
                                    let _ = SelectObject(mem_dc, GetStockObject(NULL_PEN));
                                    let mr = 6.96 / 2.0;
                                    for (px, py) in pts.iter() {
                                        draw_line_marker(
                                            mem_dc,
                                            line_marker_shape(si),
                                            *px,
                                            *py,
                                            mr,
                                            scale,
                                        );
                                    }
                                    SelectObject(mem_dc, old_m_brush);
                                    let _ = DeleteObject(m_brush);
                                }
                            }
                        }

                        // Data labels (c:dLbls): Word renders each data
                        // point's value in Calibri 18pt black, centred on the
                        // point, baseline = point_y + 6.2 (chart_datalabel_line
                        // render-truth 2026-08-07: '19.2' at (164.76,175.20)
                        // for point (155.25,169.0); same for both single- and
                        // multi-series). Format: numFmt "0.0%" -> value*100
                        // one-decimal + "%" (same as the bars).
                        if chart.has_data_labels && chart.show_val {
                            let num_fmt = chart.number_format.clone();
                            let format_label = |v: f64| -> String {
                                if num_fmt == "0.0%" {
                                    format!("{:.1}%", v * 100.0)
                                } else {
                                    // Word prints the RAW value, no rounding:
                                    // chart_datalabel_line p1 reads
                                    // "19.2 21.4 16.7" (not "19 21 17").
                                    format!("{}", v)
                                }
                            };
                            for (si, pts) in series_pts.iter().enumerate() {
                                for (pt, s) in pts.iter().zip(
                                    chart.series.get(si).map(|s| &s.values).into_iter().flatten(),
                                ) {
                                    let text = format_label(*s);
                                    // LEFT-aligned 9.50pt to the RIGHT of the
                                    // point -- NOT centred on it.  Word
                                    // render-truth chart_datalabel_line: all 12
                                    // labels across p1/p2/p3 sit at dx0 +9.46
                                    // ..+9.55 while their widths range 18.25
                                    // ("19") .. 63.07 ("1920.0%"), so the
                                    // offset is to the label's LEFT EDGE.
                                    let lx = pt.0 + 9.50;
                                    draw_text_baseline(
                                        mem_dc,
                                        (lx * scale).round() as i32,
                                        (pt.1 + 6.2) as f32,
                                        &text,
                                        axis_fs,
                                        axis_family,
                                        None,
                                        scale,
                                    );
                                }
                            }
                        }

                        // Category names centred on each category centre.
                        // In the CROWDED 78.62pt band Word wraps a label that
                        // does not fit its pitch into TWO centred lines
                        // (chart_line 'Midwest' -> "Mid"/"west", render-truth
                        // 2026-08-08: line1 ink y=299.57..326.37 -> baseline
                        // plot_bot+44.99, line2 y=322.09..348.93 -> baseline
                        // plot_bot+67.55, baseline gap 22.56; a single-line
                        // label sits at plot_bot+43.82 = East bottom 325.20).
                        // Split point = the index that best balances the two
                        // halves' Calibri hmtx widths.
                        for (ci, name) in chart.categories.iter().enumerate() {
                            let cat_center = plot_left + pitch * (ci as f64 + 0.5);
                            let lw = font_adv::line_hmtx_width_pt(name, 18.0, axis_family)
                                .unwrap_or_else(|| {
                                    name.chars().count() as f32 * 18.0 * 0.5
                                }) as f64;
                            let lx = cat_center - lw / 2.0;
                            let single_bl = if crowded { 43.8 } else { 28.67 };
                            // two-line wrap: crowded AND the whole label is
                            // wider than its pitch
                            if crowded && lw > pitch {
                                let chars: Vec<char> = name.chars().collect();
                                let n = chars.len();
                                if n >= 2 {
                                    let mut best: Option<(usize, f64)> = None;
                                    for i in 1..n {
                                        let a: String = chars[..i].iter().collect();
                                        let b: String = chars[i..].iter().collect();
                                        let wa =
                                            font_adv::line_hmtx_width_pt(&a, 18.0, axis_family)
                                                .unwrap_or_else(|| {
                                                    a.chars().count() as f32 * 18.0 * 0.5
                                                }) as f64;
                                        let wb =
                                            font_adv::line_hmtx_width_pt(&b, 18.0, axis_family)
                                                .unwrap_or_else(|| {
                                                    b.chars().count() as f32 * 18.0 * 0.5
                                                }) as f64;
                                        let d = (wa - wb).abs();
                                        if best.map_or(true, |(_, bd)| d < bd) {
                                            best = Some((i, d));
                                        }
                                    }
                                    if let Some((i, _)) = best {
                                        let a: String = chars[..i].iter().collect();
                                        let b: String = chars[i..].iter().collect();
                                        let wa =
                                            font_adv::line_hmtx_width_pt(&a, 18.0, axis_family)
                                                .unwrap_or_else(|| {
                                                    a.chars().count() as f32 * 18.0 * 0.5
                                                }) as f64;
                                        let wb =
                                            font_adv::line_hmtx_width_pt(&b, 18.0, axis_family)
                                                .unwrap_or_else(|| {
                                                    b.chars().count() as f32 * 18.0 * 0.5
                                                }) as f64;
                                        draw_text_baseline(
                                            mem_dc,
                                            ((cat_center - wa / 2.0) * scale).round() as i32,
                                            (zero_y + 45.0) as f32,
                                            &a,
                                            18.0,
                                            axis_family,
                                            None,
                                            scale,
                                        );
                                        draw_text_baseline(
                                            mem_dc,
                                            ((cat_center - wb / 2.0) * scale).round() as i32,
                                            (zero_y + 67.55) as f32,
                                            &b,
                                            18.0,
                                            axis_family,
                                            None,
                                            scale,
                                        );
                                        continue;
                                    }
                                }
                            }
                            draw_text_baseline(
                                mem_dc,
                                (lx * scale).round() as i32,
                                (zero_y + single_bl) as f32,
                                name,
                                18.0,
                                axis_family,
                                None,
                                scale,
                            );
                        }

                        // EXPLICIT <c:title> text: Word draws it as Arial
                        // 18pt (regular), centred on the frame, baseline
                        // sy+24.43 (chart_title_line/chart_title_line2
                        // render-truth 2026-08-07: origin=(194.66,96.43),
                        // same as the bar explicit title). It suppresses the
                        // automatic series-name title.
                        if let Some(title) = &chart.explicit_title {
                            let tfs = 18.0f32;
                            let lw = font_adv::line_hmtx_width_pt(title, tfs, "Arial")
                                .unwrap_or_else(|| {
                                    title.chars().count() as f32 * tfs * 0.5
                                }) as f64;
                            let frame_cx = sx + sw / 2.0;
                            draw_text_baseline_w(
                                mem_dc,
                                ((frame_cx - lw / 2.0) * scale).round() as i32,
                                (sy + 24.43) as f32,
                                title,
                                tfs,
                                "Arial",
                                None,
                                scale,
                                400,
                            );
                        }

                        // Automatic title (single series only -> the series
                        // name Calibri-Bold 21.62pt centred on the frame,
                        // baseline sy+28.03; same rule as the bars. Word
                        // shows the series name as the auto title only for
                        // a single series; an explicit <c:title> suppresses
                        // it).
                        if chart.series.len() == 1 && !has_explicit_title {
                            let first = &chart.series[0];
                            let tfs = 21.62f32;
                            let lw = font_adv::line_hmtx_width_pt(&first.name, tfs, axis_family)
                                .unwrap_or_else(|| {
                                    first.name.chars().count() as f32 * tfs * 0.5
                                }) as f64;
                            let frame_cx = sx + sw / 2.0;
                            draw_text_baseline_w(
                                mem_dc,
                                ((frame_cx - lw / 2.0) * scale).round() as i32,
                                (sy + 28.03) as f32,
                                &first.name,
                                tfs,
                                axis_family,
                                None,
                                scale,
                                700,
                            );
                        }

                        // Line legend (when <c:legend> declared): a horizontal
                        // accent1 line swatch w=2.25 of length 19.20 at
                        // legend_left = plot_right+15.65, a 6.96pt marker
                        // circle centred on it (centre = legend_left+9.57),
                        // and the series name Calibri 18pt at
                        // x0 = legend_left+21.29, baseline = legend_y0+5.24.
                        // legend_y0 = sy + shh/2 + 17.68 (FRAME-relative;
                        // single-series measured only).
                        if chart.has_legend {
                            // Per-series legend: n>=2 is frame-vertically
                            // centred (legend_y0 = sy + shh/2 -
                            // (n-1)*27.75/2, measured chart_line3), n==1
                            // keeps the single-series offset
                            // sy+shh/2+17.68 (chart_line SSIM). Each row:
                            // border-colour line swatch w=2.25 of length
                            // 19.20 at legend_left = plot_right+15.65, the
                            // series' 6.96pt marker centred on it (centre
                            // = legend_left+9.57), and the series name
                            // Calibri 18pt at x0 = legend_left+21.29,
                            // baseline = row_y+5.24.
                            let n = chart.series.len();
                            // Explicit `<c:title>` probes (chart_title_line/
                            // chart_title_line2) anchor the legend block at
                            // sy + shh/2 + 14.85 (Word-measured, both n=1 and
                            // n=2), while no-title probes keep +17.68 (n<=1)
                            // / frame-vertical centering (n>=2).
                            let legend_y0 = if has_explicit_title {
                                sy + shh / 2.0 + 14.85
                                    - (n as f64 - 1.0) * 27.75 / 2.0
                            } else if n <= 1 {
                                sy + shh / 2.0 + 17.68
                            } else {
                                sy + shh / 2.0 - (n as f64 - 1.0) * 27.75 / 2.0
                            };
                            let legend_left = plot_right + 15.65;
                            for (si, s) in chart.series.iter().enumerate() {
                                let row_y = legend_y0 + si as f64 * 27.75;
                                let (fill_hex, border_hex) = line_series_colors(si);
                                if let Some(rgb) = parse_hex_rgb(&border_hex) {
                                    let lg_pen = CreatePen(
                                        PS_SOLID,
                                        (2.25 * scale).round().max(1.0) as i32,
                                        COLORREF(colorref(rgb.0, rgb.1, rgb.2)),
                                    );
                                    let old_lg_pen = SelectObject(mem_dc, lg_pen);
                                    let _ = SelectObject(mem_dc, GetStockObject(NULL_BRUSH));
                                    let ly = (row_y * scale).round() as i32;
                                    let _ = MoveToEx(
                                        mem_dc,
                                        (legend_left * scale).round() as i32,
                                        ly,
                                        None,
                                    );
                                    let _ = LineTo(
                                        mem_dc,
                                        ((legend_left + 19.20) * scale).round() as i32,
                                        ly,
                                    );
                                    SelectObject(mem_dc, old_lg_pen);
                                    let _ = DeleteObject(lg_pen);
                                }
                                if let Some(rgb) = parse_hex_rgb(&fill_hex) {
                                    let m_brush = CreateSolidBrush(COLORREF(
                                        colorref(rgb.0, rgb.1, rgb.2),
                                    ));
                                    let old_m_brush = SelectObject(mem_dc, m_brush);
                                    let _ = SelectObject(mem_dc, GetStockObject(NULL_PEN));
                                    draw_line_marker(
                                        mem_dc,
                                        line_marker_shape(si),
                                        legend_left + 9.57,
                                        row_y,
                                        6.96 / 2.0,
                                        scale,
                                    );
                                    SelectObject(mem_dc, old_m_brush);
                                    let _ = DeleteObject(m_brush);
                                }
                                draw_text_baseline(
                                    mem_dc,
                                    ((legend_left + 21.29) * scale).round() as i32,
                                    (row_y + 5.24) as f32,
                                    &s.name,
                                    18.0,
                                    axis_family,
                                    None,
                                    scale,
                                );
                            }
                        }

                        // Axis lines + ticks (same as the bars: Y axis
                        // vertical plot_left, X axis horizontal plot_bot,
                        // Y ticks 0..=axis_steps at plot_left-5.7, X ticks
                        // 0..=n_cat at category boundaries, plot frame top
                        // edge painted).
                        let axis_pen = CreatePen(
                            PS_SOLID,
                            2,
                            COLORREF(colorref(0, 0, 0)),
                        );
                        let old_axis_pen = SelectObject(mem_dc, axis_pen);
                        let _ = SelectObject(mem_dc, GetStockObject(NULL_BRUSH));
                        let pl = (plot_left * scale).round() as i32;
                        let pt = (plot_top * scale).round() as i32;
                        let pr = (plot_right * scale).round() as i32;
                        let pb = (plot_bot * scale).round() as i32;
                        let _ = MoveToEx(mem_dc, pl, pt, None);
                        let _ = LineTo(mem_dc, pl, pb);
                        let _ = MoveToEx(mem_dc, pl, pb, None);
                        let _ = LineTo(mem_dc, pr, pb);
                        let _ = MoveToEx(mem_dc, pl, pt, None);
                        let _ = LineTo(mem_dc, pr, pt);
                        for i in 0..=axis_steps {
                            let tick_y = plot_bot - plot_h * i as f64 / axis_steps as f64;
                            let ty = (tick_y * scale).round() as i32;
                            let _ = MoveToEx(mem_dc, ((plot_left - 5.7) * scale).round() as i32, ty, None);
                            let _ = LineTo(mem_dc, pl, ty);
                        }
                        // Category ticks hang off the ZERO line (N3: the four
                        // ticks run 280.97 -> 286.68 while plot_bot is 344.01).
                        let zt = (zero_y * scale).round() as i32;
                        for i in 0..=n_cat {
                            let tick_x = plot_left + pitch * i as f64;
                            let tx = (tick_x * scale).round() as i32;
                            let _ = MoveToEx(mem_dc, tx, zt, None);
                            let _ = LineTo(mem_dc, tx, ((zero_y + 5.7) * scale).round() as i32);
                        }
                        SelectObject(mem_dc, old_axis_pen);
                        let _ = DeleteObject(axis_pen);
                        } else if chart.bar_dir == "bar"
                            && std::env::var("OXI_HBAR_DISABLE").is_err()
                        {
                        // ============ HORIZONTAL bar chart (barDir="bar") ============
                        // Word render-truth, measured 2026-08-09 via fitz
                        // get_drawings/rawdict on pipeline_data/pptx_probes/
                        // chart_bar (5 slides: 1-series, 2-series, stacked,
                        // legend, data labels) + chart_bar_axis (14-arm sweep).
                        //
                        //   plot_left  = sx + 6.50 + widest category label + 16.70
                        //     (159.07 for [East,West,Midwest], 105.62 for [A,B,C]
                        //      -- two probes 53pt apart in label width, both EXACT)
                        //   plot_top   = sy + 46.37 with the auto title / sy + 11.00
                        //     without (pages 1&5 vs 2,3,4)
                        //   plot_bot   = sy + sh - 39.90  (same as the column chart)
                        //   plot_right = band_right - w(axis-max label)/2, where
                        //     band_right = sx + sw - 11.00, or, with a legend,
                        //     legend_swatch_x0 - 18.15  (447.88 / 352.42 measured)
                        //   category 0 is the BOTTOM row; inside a cluster series 0
                        //     is the bottom-most bar; a stacked bar grows rightwards
                        //   bar height = pitch/(n_ser+1.5) clustered, pitch*0.4 stacked
                        //   value labels centred under their tick, baseline
                        //     plot_bot + 28.67; category labels right-aligned at
                        //     plot_left - 16.70, baseline cat_center + 5.22
                        //   ticks: value 5.71pt below the axis (div+1), category
                        //     5.72pt left of the axis (n_cat+1)
                        //   gridlines: VERTICAL, full plot height, i=1..div (the
                        //     i=0 line coincides with the category axis); there is
                        //     NO horizontal frame edge at plot_top
                        let sx = sh.x as f64;
                        let sy = sh.y as f64;
                        let sw = sh.width as f64;
                        let shh = sh.height as f64;
                        let axis_fs = 18.0f32;
                        let axis_family = "Calibri";
                        let has_auto_title = chart.series.len() == 1;
                        let has_explicit_title = chart.explicit_title.is_some();
                        let is_stacked = chart.grouping == "stacked"
                            || chart.grouping == "percentStacked";
                        let n_cat = chart.categories.len().max(1);
                        let n_ser = chart.series.len().max(1);

                        let text_w = |t: &str, fs: f32| -> f64 {
                            font_adv::line_hmtx_width_pt(t, fs, axis_family)
                                .unwrap_or_else(|| {
                                    t.chars().count() as f32 * fs * 0.5
                                }) as f64
                        };

                        let cat_label_w = chart
                            .categories
                            .iter()
                            .map(|c| text_w(c, axis_fs))
                            .fold(0.0f64, f64::max);
                        // NEGATIVE data (chart_negative N7, 2026-08-10): the
                        // category names move INSIDE, right-aligned to the zero
                        // line, so the left gutter now only has to hold half of
                        // the leftmost VALUE label (94.88 measured = sx + 11.0 +
                        // w("-10")/2); plot_bot keeps the 39.9 band because the
                        // value labels are still underneath.
                        let raw_min_h = if is_stacked {
                            (0..n_cat)
                                .map(|ci| {
                                    chart
                                        .series
                                        .iter()
                                        .map(|s| {
                                            s.values.get(ci).copied().unwrap_or(0.0)
                                        })
                                        .filter(|v| *v < 0.0)
                                        .sum::<f64>()
                                })
                                .fold(0.0f64, f64::min)
                        } else {
                            chart
                                .series
                                .iter()
                                .flat_map(|s| s.values.iter().copied())
                                .fold(0.0f64, f64::min)
                        };
                        let has_neg = raw_min_h < 0.0;
                        let plot_left_pos = sx + 6.50 + cat_label_w + 16.70;
                        let plot_top = if has_explicit_title {
                            // MEASURED 2026-08-09 (chart_bar_resid slides 1/2/5,
                            // Word PDF category-tick tops): 112.70 on a frame at
                            // sy=72.  The earlier value was 40.66, adopted by
                            // analogy with the column chart; the probe puts it at
                            // 40.70, so the analogy was right to 0.04pt.
                            sy + 40.70
                        } else if has_auto_title {
                            sy + 46.37
                        } else {
                            sy + 11.0
                        };
                        let plot_bot = sy + shh - 39.9;
                        let plot_h = plot_bot - plot_top;

                        // Legend block: identical geometry to the column chart
                        // (chart_bar page 4 reproduces swatch_x0 379.74 and
                        // label_x0 394.25 exactly) but the ROWS RUN BOTTOM-UP,
                        // matching the bar order (series 0 lowest).
                        let legend_lfs = 18.0f32;
                        let legend_label_w = chart
                            .series
                            .iter()
                            .map(|s| text_w(&s.name, legend_lfs))
                            .fold(0.0f64, f64::max);
                        let legend_swatch_w = 9.89f64;
                        let legend_gap = 4.62f64;
                        let legend_row_pitch = 27.75f64;
                        let legend_right = (sx + sw) - 10.0;
                        let legend_swatch_x1 =
                            legend_right - legend_label_w - legend_gap;
                        let legend_swatch_x0 = legend_swatch_x1 - legend_swatch_w;
                        let band_right = if chart.has_legend {
                            legend_swatch_x0 - 18.15
                        } else {
                            sx + sw - 11.0
                        };

                        let raw_max = if is_stacked {
                            (0..n_cat)
                                .map(|ci| {
                                    chart
                                        .series
                                        .iter()
                                        .map(|s| {
                                            s.values.get(ci).copied().unwrap_or(0.0)
                                        })
                                        .sum::<f64>()
                                })
                                .fold(0.0f64, f64::max)
                        } else {
                            chart
                                .series
                                .iter()
                                .flat_map(|s| s.values.iter().copied())
                                .fold(0.0f64, f64::max)
                        };
                        // plot_right depends on the axis-max label width while the
                        // axis choice depends on plot_w -> settle with one refine
                        // pass (the half-label moves plot_w by under 10pt).
                        let mut axis_min = 0.0f64;
                        let mut axis_max = nice_axis_max(raw_max);
                        let mut axis_steps = 5usize;
                        let mut plot_left = plot_left_pos;
                        let mut plot_right = band_right - 9.13;
                        for _ in 0..2 {
                            if has_neg {
                                let (lo, hi, d) = nice_axis_range(
                                    raw_min_h,
                                    raw_max,
                                    plot_right - plot_left,
                                    HORIZ_MIN_SPACING,
                                );
                                axis_min = lo;
                                axis_max = hi;
                                axis_steps = d.max(1);
                            } else {
                                let (m, d) = horiz_value_axis(
                                    raw_max,
                                    plot_right - plot_left,
                                );
                                axis_min = 0.0;
                                axis_max = m;
                                axis_steps = d.max(1);
                            }
                            let step =
                                (axis_max - axis_min) / axis_steps as f64;
                            if has_neg {
                                // Both edges leave half of the OUTERMOST label.
                                plot_left = sx
                                    + 11.0
                                    + text_w(
                                        &fmt_axis_value(axis_min, step),
                                        axis_fs,
                                    ) / 2.0;
                            }
                            plot_right = band_right
                                - text_w(&fmt_axis_value(axis_max, step), axis_fs)
                                    / 2.0;
                        }
                        let plot_left = plot_left;
                        let plot_w = plot_right - plot_left;
                        let pitch = plot_h / n_cat as f64;
                        let axis_span = (axis_max - axis_min).max(1e-9);
                        let axis_step = axis_span / axis_steps as f64;
                        let val_x =
                            |v: f64| plot_left + ((v - axis_min) / axis_span) * plot_w;
                        let zero_x = val_x(0.0);

                        // ---- value-axis gridlines (vertical, behind the bars) ----
                        {
                            let grid_pen =
                                CreatePen(PS_SOLID, 2, COLORREF(colorref(0, 0, 0)));
                            let old_pen = SelectObject(mem_dc, grid_pen);
                            let _ =
                                SelectObject(mem_dc, GetStockObject(NULL_BRUSH));
                            for i in 1..=axis_steps {
                                let gx = plot_left
                                    + plot_w * i as f64 / axis_steps as f64;
                                let gxi = (gx * scale).round() as i32;
                                let _ = MoveToEx(
                                    mem_dc,
                                    gxi,
                                    (plot_top * scale).round() as i32,
                                    None,
                                );
                                let _ = LineTo(
                                    mem_dc,
                                    gxi,
                                    (plot_bot * scale).round() as i32,
                                );
                            }
                            SelectObject(mem_dc, old_pen);
                            let _ = DeleteObject(grid_pen);
                        }

                        // ---- bars ----
                        let vary_points = chart.series.len() == 1;
                        for ci in 0..n_cat {
                            let cat_center = plot_bot - pitch * (ci as f64 + 0.5);
                            if is_stacked {
                                let bar_h = pitch * 0.4;
                                let by0 = cat_center - bar_h / 2.0;
                                let mut cum_pos = 0.0f64;
                                let mut cum_neg = 0.0f64;
                                for (si, series) in chart.series.iter().enumerate() {
                                    let v =
                                        series.values.get(ci).copied().unwrap_or(0.0);
                                    if v == 0.0 {
                                        continue;
                                    }
                                    let seg_w = (v / axis_span * plot_w).abs();
                                    let (bx0, bx1) = if v > 0.0 {
                                        let a = zero_x + cum_pos;
                                        cum_pos += seg_w;
                                        (a, a + seg_w)
                                    } else {
                                        let b = zero_x - cum_neg;
                                        cum_neg += seg_w;
                                        (b - seg_w, b)
                                    };
                                    let neg = v < 0.0;
                                    let col_hex = pres
                                        .theme_colors
                                        .get(&format!("accent{}", si + 1))
                                        .map(|s| s.as_str())
                                        .or_else(|| DEFAULT_ACCENT.get(si).copied());
                                    if let Some(rgb0) =
                                        col_hex.and_then(parse_hex_rgb)
                                    {
                                        // invertIfNegative defaults to TRUE.
                                        let rgb = if neg {
                                            (255u8, 255u8, 255u8)
                                        } else {
                                            rgb0
                                        };
                                        let brush = CreateSolidBrush(COLORREF(
                                            colorref(rgb.0, rgb.1, rgb.2),
                                        ));
                                        let old_brush = SelectObject(mem_dc, brush);
                                        let r = RECT {
                                            left: (bx0 * scale).round() as i32,
                                            top: (by0 * scale).round() as i32,
                                            right: (bx1 * scale).round() as i32,
                                            bottom: ((by0 + bar_h) * scale).round()
                                                as i32,
                                        };
                                        let _ = FillRect(mem_dc, &r, brush);
                                        SelectObject(mem_dc, old_brush);
                                        let _ = DeleteObject(brush);
                                        if neg {
                                            draw_neg_bar_outline(mem_dc, &r, scale);
                                        }
                                    }
                                }
                            } else {
                                let bar_h = pitch / (n_ser as f64 + 1.5);
                                let cluster_h = bar_h * n_ser as f64;
                                for (si, series) in chart.series.iter().enumerate() {
                                    let v =
                                        series.values.get(ci).copied().unwrap_or(0.0);
                                    let accent_idx = if vary_points { ci } else { si };
                                    let col_hex = pres
                                        .theme_colors
                                        .get(&format!("accent{}", accent_idx + 1))
                                        .map(|s| s.as_str())
                                        .or_else(|| {
                                            DEFAULT_ACCENT.get(accent_idx).copied()
                                        });
                                    if let Some(rgb0) =
                                        col_hex.and_then(parse_hex_rgb)
                                    {
                                        // invertIfNegative defaults to TRUE.
                                        let rgb = if v < 0.0 {
                                            (255u8, 255u8, 255u8)
                                        } else {
                                            rgb0
                                        };
                                        let brush = CreateSolidBrush(COLORREF(
                                            colorref(rgb.0, rgb.1, rgb.2),
                                        ));
                                        let old_brush = SelectObject(mem_dc, brush);
                                        // series 0 is the BOTTOM bar of the cluster
                                        let by0 = cat_center + cluster_h / 2.0
                                            - (si as f64 + 1.0) * bar_h;
                                        let vx = val_x(v);
                                        let (bx0, bx1) = if v >= 0.0 {
                                            (zero_x, vx)
                                        } else {
                                            (vx, zero_x)
                                        };
                                        let r = RECT {
                                            left: (bx0 * scale).round() as i32,
                                            top: (by0 * scale).round() as i32,
                                            right: (bx1 * scale).round() as i32,
                                            bottom: ((by0 + bar_h) * scale).round()
                                                as i32,
                                        };
                                        let _ = FillRect(mem_dc, &r, brush);
                                        SelectObject(mem_dc, old_brush);
                                        let _ = DeleteObject(brush);
                                        if v < 0.0 {
                                            draw_neg_bar_outline(mem_dc, &r, scale);
                                        }
                                    }
                                }
                            }
                        }

                        // ---- data labels ----
                        // chart_bar page 5 render-truth: OUTSIDE_END puts the label
                        // 6.06pt right of the bar end, baseline cat_center + 6.22.
                        // STACKED centres each label in its own segment, baseline
                        // cat_center + 6.22 -- MEASURED 2026-08-09 on
                        // chart_bar_resid slides 3/4: all 6 segment/label centre
                        // pairs agree within 0.1pt and the baseline offset reads
                        // 6.19/6.23/6.26 across the three rows.
                        if chart.has_data_labels && chart.show_val {
                            let num_fmt = chart.number_format.clone();
                            let fmt = |v: f64| -> String {
                                if num_fmt == "0.0%" {
                                    format!("{:.1}%", v * 100.0)
                                } else if num_fmt == "0%" {
                                    format!("{}%", (v * 100.0).round() as i64)
                                } else if (v - v.round()).abs() < 1e-9 {
                                    format!("{}", v.round() as i64)
                                } else {
                                    let s = format!("{:.4}", v);
                                    s.trim_end_matches('0').trim_end_matches('.').to_string()
                                }
                            };
                            for ci in 0..n_cat {
                                let cat_center =
                                    plot_bot - pitch * (ci as f64 + 0.5);
                                let mut cum = 0.0f64;
                                for (si, series) in chart.series.iter().enumerate() {
                                    let v =
                                        series.values.get(ci).copied().unwrap_or(0.0);
                                    let seg_w = if axis_max > 0.0 {
                                        v / axis_max * plot_w
                                    } else {
                                        0.0
                                    };
                                    let label = fmt(v);
                                    let lw = text_w(&label, axis_fs);
                                    let (lx, ly) = if is_stacked {
                                        (
                                            plot_left + cum + seg_w / 2.0 - lw / 2.0,
                                            cat_center + 6.22,
                                        )
                                    } else {
                                        let bar_h =
                                            pitch / (n_ser as f64 + 1.5);
                                        let cluster_h = bar_h * n_ser as f64;
                                        let by0 = cat_center + cluster_h / 2.0
                                            - (si as f64 + 1.0) * bar_h;
                                        (
                                            plot_left + seg_w + 6.06,
                                            by0 + bar_h / 2.0 + 6.22,
                                        )
                                    };
                                    draw_text_baseline(
                                        mem_dc,
                                        (lx * scale).round() as i32,
                                        ly as f32,
                                        &label,
                                        axis_fs,
                                        axis_family,
                                        None,
                                        scale,
                                    );
                                    cum += seg_w;
                                }
                            }
                        }

                        // ---- value-axis labels (below the axis, centred) ----
                        for i in 0..=axis_steps {
                            let v = axis_min + axis_step * i as f64;
                            let label = fmt_axis_value(v, axis_step);
                            let lw = text_w(&label, axis_fs);
                            let tick_x =
                                plot_left + plot_w * i as f64 / axis_steps as f64;
                            draw_text_baseline(
                                mem_dc,
                                ((tick_x - lw / 2.0) * scale).round() as i32,
                                (plot_bot + 28.67) as f32,
                                &label,
                                axis_fs,
                                axis_family,
                                None,
                                scale,
                            );
                        }

                        // ---- category labels (left of the axis, right-aligned) ----
                        // With negatives the category axis IS the zero line, and
                        // the names right-align 16.64pt to its left (N7: 'Q3'
                        // ends at 166.46, zero_x 183.13).
                        let cat_axis_x = if has_neg { zero_x } else { plot_left };
                        for (ci, cat) in chart.categories.iter().enumerate() {
                            let cat_center = plot_bot - pitch * (ci as f64 + 0.5);
                            let lw = text_w(cat, axis_fs);
                            draw_text_baseline(
                                mem_dc,
                                ((cat_axis_x - 16.70 - lw) * scale).round() as i32,
                                (cat_center + 5.22) as f32,
                                cat,
                                axis_fs,
                                axis_family,
                                None,
                                scale,
                            );
                        }

                        // ---- explicit / automatic chart title ----
                        if let Some(title) = chart.explicit_title.as_ref() {
                            let tfs = 18.0f32;
                            let lw = text_w(title, tfs);
                            draw_text_baseline_w(
                                mem_dc,
                                (((sx + sw / 2.0) - lw / 2.0) * scale).round() as i32,
                                (sy + 24.43) as f32,
                                title,
                                tfs,
                                "Arial",
                                None,
                                scale,
                                400,
                            );
                        } else if has_auto_title {
                            if let Some(first) = chart.series.first() {
                                let tfs = 21.62f32;
                                let lw = text_w(&first.name, tfs);
                                draw_text_baseline_w(
                                    mem_dc,
                                    (((sx + sw / 2.0) - lw / 2.0) * scale).round()
                                        as i32,
                                    (sy + 28.03) as f32,
                                    &first.name,
                                    tfs,
                                    axis_family,
                                    None,
                                    scale,
                                    700,
                                );
                            }
                        }

                        // ---- legend ----
                        // MEASURED 2026-08-09: the row order follows the visual
                        // stacking direction.  CLUSTERED bars grow bottom-up
                        // inside a cluster, so the legend mirrors (series 0 at the
                        // BOTTOM: chart_bar p4 puts Cost above Revenue); STACKED
                        // segments grow left-to-right, so the legend stays natural
                        // (series 0 on TOP: chart_bar_resid p4 puts Revenue above
                        // Cost).  Both pages share the same swatch x/y, so only the
                        // row index differs.
                        if chart.has_legend {
                            let legend_total_h =
                                (n_ser as f64 - 1.0) * legend_row_pitch
                                    + legend_swatch_w;
                            let legend_y0 =
                                (sy + shh / 2.0) - legend_total_h / 2.0;
                            for (si, series) in chart.series.iter().enumerate() {
                                let row = if is_stacked { si } else { n_ser - 1 - si };
                                let sw_y =
                                    legend_y0 + row as f64 * legend_row_pitch;
                                let col_hex = pres
                                    .theme_colors
                                    .get(&format!("accent{}", si + 1))
                                    .map(|s| s.as_str())
                                    .or_else(|| DEFAULT_ACCENT.get(si).copied());
                                if let Some(rgb) = col_hex.and_then(parse_hex_rgb) {
                                    let brush = CreateSolidBrush(COLORREF(colorref(
                                        rgb.0, rgb.1, rgb.2,
                                    )));
                                    let old_brush = SelectObject(mem_dc, brush);
                                    let r = RECT {
                                        left: (legend_swatch_x0 * scale).round()
                                            as i32,
                                        top: (sw_y * scale).round() as i32,
                                        right: (legend_swatch_x1 * scale).round()
                                            as i32,
                                        bottom: ((sw_y + legend_swatch_w) * scale)
                                            .round()
                                            as i32,
                                    };
                                    let _ = FillRect(mem_dc, &r, brush);
                                    SelectObject(mem_dc, old_brush);
                                    let _ = DeleteObject(brush);
                                }
                                draw_text_baseline(
                                    mem_dc,
                                    ((legend_swatch_x1 + legend_gap) * scale).round()
                                        as i32,
                                    (sw_y + legend_swatch_w + 0.28) as f32,
                                    &series.name,
                                    legend_lfs,
                                    axis_family,
                                    None,
                                    scale,
                                );
                            }
                        }

                        // ---- axis lines + ticks ----
                        {
                            let axis_pen =
                                CreatePen(PS_SOLID, 2, COLORREF(colorref(0, 0, 0)));
                            let old_pen = SelectObject(mem_dc, axis_pen);
                            let _ =
                                SelectObject(mem_dc, GetStockObject(NULL_BRUSH));
                            let pl = (plot_left * scale).round() as i32;
                            let pt = (plot_top * scale).round() as i32;
                            let pr = (plot_right * scale).round() as i32;
                            let pb = (plot_bot * scale).round() as i32;
                            // category axis (vertical) + value axis (horizontal)
                            // -- the category axis rides the ZERO line when the
                            // value axis has negatives (N7: 183.13).
                            let pz = (cat_axis_x * scale).round() as i32;
                            let _ = MoveToEx(mem_dc, pz, pt, None);
                            let _ = LineTo(mem_dc, pz, pb);
                            let _ = MoveToEx(mem_dc, pl, pb, None);
                            let _ = LineTo(mem_dc, pr, pb);
                            // value ticks below the axis
                            for i in 0..=axis_steps {
                                let tx = ((plot_left
                                    + plot_w * i as f64 / axis_steps as f64)
                                    * scale)
                                    .round() as i32;
                                let _ = MoveToEx(mem_dc, tx, pb, None);
                                let _ = LineTo(
                                    mem_dc,
                                    tx,
                                    ((plot_bot + 5.71) * scale).round() as i32,
                                );
                            }
                            // category ticks left of the axis
                            for i in 0..=n_cat {
                                let ty = ((plot_top + pitch * i as f64) * scale)
                                    .round() as i32;
                                let _ = MoveToEx(
                                    mem_dc,
                                    ((cat_axis_x - 5.72) * scale).round() as i32,
                                    ty,
                                    None,
                                );
                                let _ = LineTo(mem_dc, pz, ty);
                            }
                            SelectObject(mem_dc, old_pen);
                            let _ = DeleteObject(axis_pen);
                        }
                        } else {
                        let sx = sh.x as f64;
                        let sy = sh.y as f64;
                        let sw = sh.width as f64;
                        let shh = sh.height as f64;
                        let has_auto_title = chart.series.len() == 1;
                        let has_explicit_title = chart.explicit_title.is_some();
                        let is_100pct = chart.grouping == "percentStacked";
                        let is_stacked = chart.grouping == "stacked" || is_100pct;
                        let n_cat = chart.categories.len().max(1);
                        let n_ser = chart.series.len().max(1);
                        let axis_fs = 18.0f32;
                        let axis_family = "Calibri";
                        // NEGATIVE data (chart_negative probe, 2026-08-10):
                        // the axis spans zero, so both ends of the data range
                        // matter.  A STACKED chart accumulates each SIGN away
                        // from zero independently, so its range is the largest
                        // per-category POSITIVE sum and the smallest per-category
                        // NEGATIVE sum (N8: +29.7 / -19.7 -> axis -30..40).
                        let (min_val, max_val) = if is_100pct {
                            // percentStacked: fixed 0..100 scale (10-step %-axis).
                            (0.0, 100.0)
                        } else if is_stacked {
                            let mut lo = 0.0f64;
                            let mut hi = 0.0f64;
                            for ci in 0..n_cat {
                                let mut pos = 0.0f64;
                                let mut neg = 0.0f64;
                                for s in &chart.series {
                                    let v = s.values.get(ci).copied().unwrap_or(0.0);
                                    if v >= 0.0 {
                                        pos += v;
                                    } else {
                                        neg += v;
                                    }
                                }
                                hi = hi.max(pos);
                                lo = lo.min(neg);
                            }
                            (lo, hi)
                        } else {
                            let mut lo = 0.0f64;
                            let mut hi = 0.0f64;
                            for s in &chart.series {
                                for v in s.values.iter().copied() {
                                    hi = hi.max(v);
                                    lo = lo.min(v);
                                }
                            }
                            (lo, hi)
                        };
                        // With negatives the category names move INSIDE the
                        // plot (they hang off the zero line), so the 39.9pt
                        // bottom band that normally holds them collapses to
                        // the plain 16.0pt margin: N1/N2/N8 all put plot_bot
                        // at sy+shh-16.0 (344.01) instead of 320.10.
                        let has_neg = min_val < 0.0;
                        let plot_top = if has_explicit_title {
                            // An explicit <c:title> shifts the plot down by
                            // the title line: plot_top = sy+45.69
                            // (chart_title/chart_title2 render-truth
                            // 2026-08-07; Arial 18pt title, vs the auto
                            // title's 21.62pt Calibri-Bold at sy+51.4).
                            sy + 45.69
                        } else if has_auto_title {
                            sy + 51.4
                        } else {
                            sy + 16.0
                        };
                        let plot_bot = sy + shh - if has_neg { 16.0 } else { 39.9 };
                        let plot_h = plot_bot - plot_top;

                        // Value axis labels (0..max_axis in even steps),
                        // right-aligned to a fixed gutter. For a CLUSTERED
                        // chart the scale is 0..max_axis in 5 steps (6
                        // labels). For a STACKED chart Word scales to the
                        // largest per-category series SUM (chart_stacked:
                        // Q2 sum 36.4 -> nice max 40) and draws one label
                        // per 5-step tick, i.e. (max_axis/5)+1 labels
                        // (0,5,...,40 = 9 labels, render-truth 2026-08-06).
                        //
                        // percentStacked's axis is FIXED at 100, so it must not
                        // go through the nice-range search (whose 5% headroom
                        // would round 100 up to 120).
                        let (axis_min, max_axis, axis_steps) = if is_100pct {
                            // percentStacked: 0%,10%,...,100% (11 labels).
                            (0.0, 100.0, 10usize)
                        } else {
                            nice_axis_range(min_val, max_val, plot_h, VERT_MIN_SPACING)
                        };
                        let axis_span = (max_axis - axis_min).max(1e-9);
                        // The value gutter is measured from the WIDEST value
                        // label: plot_left = sx + 6.50 + w + 16.70 (the same
                        // rule the horizontal-bar and area branches use, and
                        // the rule the 41.4 / 63.44 constants below are the
                        // 2-digit / "100%" evaluations of).  A negative axis
                        // grows the labels by the minus sign, so it must be
                        // computed: N1/N2/N8 all put plot_left at 118.96 =
                        // 72 + 6.50 + w("-10") 23.76 + 16.70.
                        let plot_left = if has_neg {
                            let mut widest = 0.0f64;
                            for i in 0..=axis_steps {
                                let val = axis_min
                                    + (max_axis - axis_min) * i as f64
                                        / axis_steps as f64;
                                let label = format!("{}", val.round() as i64);
                                let lw = font_adv::line_hmtx_width_pt(
                                    &label, axis_fs, axis_family,
                                )
                                .unwrap_or_else(|| {
                                    label.chars().count() as f32 * axis_fs * 0.5
                                }) as f64;
                                widest = widest.max(lw);
                            }
                            sx + 6.5 + widest + 16.7
                        } else if is_100pct {
                            sx + 63.44
                        } else {
                            sx + 41.4
                        };
                        let plot_right = sx + sw - 11.0;
                        let plot_w = plot_right - plot_left;
                        // Value -> y.  With axis_min == 0 (the positive-only
                        // case) this is exactly the previous `plot_bot -
                        // v/max_axis*plot_h`, so every existing probe is
                        // unchanged.
                        let val_y =
                            |v: f64| plot_bot - ((v - axis_min) / axis_span) * plot_h;
                        // The category axis sits at the value 0, NOT at the
                        // bottom of the plot (N1: bars grow from y=280.97 and
                        // the category names sit at 280.97+28.68).
                        let zero_y = val_y(0.0);
                        for i in 0..=axis_steps {
                            // percentStacked labels the fixed 0..100 scale
                            // as "0%".."100%" (render-truth: '0%' x=96.77 /
                            // '100%' x=78.50, right-aligned, baseline
                            // tick_y + 5.2).
                            let val =
                                axis_min + axis_span * i as f64 / axis_steps as f64;
                            let tick_y = val_y(val);
                            let label = if is_100pct {
                                format!("{:.0}%", val)
                            } else {
                                format!("{}", val.round() as i64)
                            };
                            let lw = font_adv::line_hmtx_width_pt(&label, axis_fs, axis_family)
                                .unwrap_or_else(|| {
                                    label.chars().count() as f32 * axis_fs * 0.5
                                }) as f64;
                            let lx = plot_left - 16.64 - lw;
                            draw_text_baseline(
                                mem_dc,
                                (lx * scale).round() as i32,
                                (tick_y + 5.2) as f32,
                                &label,
                                axis_fs,
                                axis_family,
                                None,
                                scale,
                            );
                        }

                        // Horizontal major gridlines (Word render-truth
                        // 2026-08-06, fitz get_drawings items level):
                        // Word draws black lines spanning the full plot
                        // width (plot_left->plot_right) at the value-tick
                        // rows i=1..=axis_steps (i=0 is the X axis line,
                        // drawn in the axis section below). They render
                        // BEHIND the bars (a bar covers its crossing), so
                        // they are drawn before the bars. STACKED = 8
                        // lines (chart_stacked), CLUSTERED = 5 lines
                        // (chart1/2/3/2b all identical, y = plot_bot -
                        // i*plot_h/5), LINE = 5 lines (chart_line).
                        if axis_steps > 0 {
                            let grid_pen = CreatePen(
                                PS_SOLID,
                                2,
                                COLORREF(colorref(0, 0, 0)),
                            );
                            let old_grid_pen =
                                SelectObject(mem_dc, grid_pen);
                            let _ = SelectObject(
                                mem_dc,
                                GetStockObject(NULL_BRUSH),
                            );
                            let gl = (plot_left * scale).round() as i32;
                            let gr = (plot_right * scale).round() as i32;
                            for i in 1..=axis_steps {
                                let grid_y = plot_bot
                                    - plot_h * i as f64 / axis_steps as f64;
                                let gy = (grid_y * scale).round() as i32;
                                let _ = MoveToEx(mem_dc, gl, gy, None);
                                let _ = LineTo(mem_dc, gr, gy);
                            }
                            SelectObject(mem_dc, old_grid_pen);
                            let _ = DeleteObject(grid_pen);
                        }

                        // Bars: one cluster per category, series side by side
                        // (touching within a cluster). Colour rule: with a
                        // SINGLE series Word colours each data POINT with the
                        // theme accents in order (varyColors default); with
                        // multiple series each SERIES takes one accent.
                        // Bar width derived 2026-08-06 (Word get_drawings):
                        //   chart1 (n_ser=1) bar_w = 45.81 = pitch/(1+1.5)
                        //   chart2 (n_ser=2) bar_w = 32.71 = pitch/(2+1.5)
                        //   chart3 (n_ser=3) bar_w = 25.44 = pitch/(3+1.5)
                        //   i.e. bar_w = pitch / (n_ser + 1.5) exactly.
                        // STACKED: one bar per category (width = pitch*0.4,
                        // render-truth 45.72 = 114.5*0.4, chart_stacked
                        // 2026-08-06); series 0 is the BOTTOM segment and
                        // each later series stacks on the one below it.
                        let pitch = plot_w / n_cat as f64;
                        let vary_points = chart.series.len() == 1;
                        if is_stacked {
                            let bar_w = pitch * 0.4;
                            for ci in 0..n_cat {
                                let cat_center =
                                    plot_left + pitch * (ci as f64 + 0.5);
                                let bx0 = cat_center - bar_w / 2.0;
                                let mut cum_pos = 0.0f64;
                                let mut cum_neg = 0.0f64;
                                for (si, series) in
                                    chart.series.iter().enumerate()
                                {
                                    let v = series
                                        .values
                                        .get(ci)
                                        .copied()
                                        .unwrap_or(0.0);
                                    // percentStacked: each category is 100% of
                                    // its series SUM (render-truth 2026-08-07:
                                    // Q1 blue 19.2/29.7*plot_h, red
                                    // 10.5/29.7*plot_h; the stack fills the
                                    // plot height).
                                    let seg_h = if is_100pct {
                                        let sum_cat = chart
                                            .series
                                            .iter()
                                            .map(|s| {
                                                s.values
                                                    .get(ci)
                                                    .copied()
                                                    .unwrap_or(0.0)
                                            })
                                            .sum::<f64>();
                                        if sum_cat > 0.0 {
                                            v / sum_cat * plot_h
                                        } else {
                                            0.0
                                        }
                                    } else {
                                        (v / axis_span * plot_h).abs()
                                    };
                                    // Each SIGN stacks away from zero
                                    // independently (N8 render-truth: Q2's two
                                    // negative segments run 234.29->265.38 and
                                    // 265.38->306.34, i.e. downward from the
                                    // zero line, not upward from plot_bot).
                                    let (by0, by1) = if v >= 0.0 {
                                        let b1 = zero_y - cum_pos;
                                        let b0 = b1 - seg_h;
                                        cum_pos += seg_h;
                                        (b0, b1)
                                    } else {
                                        let b0 = zero_y + cum_neg;
                                        let b1 = b0 + seg_h;
                                        cum_neg += seg_h;
                                        (b0, b1)
                                    };
                                    let accent_idx =
                                        if vary_points { ci } else { si };
                                    let col_hex = pres
                                        .theme_colors
                                        .get(&format!(
                                            "accent{}",
                                            accent_idx + 1
                                        ))
                                        .map(|s| s.as_str())
                                        .or_else(|| {
                                            DEFAULT_ACCENT
                                                .get(accent_idx)
                                                .copied()
                                        });
                                    if let Some(rgb) =
                                        col_hex.and_then(parse_hex_rgb)
                                    {
                                        // Negative segments invert to white
                                        // here too (N8: both of Q2's segments
                                        // are #FFFFFF, not accent).
                                        let neg = v < 0.0;
                                        let rgb = if neg {
                                            (255u8, 255u8, 255u8)
                                        } else {
                                            rgb
                                        };
                                        let brush = CreateSolidBrush(
                                            COLORREF(colorref(
                                                rgb.0, rgb.1, rgb.2,
                                            )),
                                        );
                                        let old_brush =
                                            SelectObject(mem_dc, brush);
                                        let r = RECT {
                                            left: (bx0 * scale).round() as i32,
                                            top: (by0 * scale).round() as i32,
                                            right: ((bx0 + bar_w) * scale)
                                                .round() as i32,
                                            bottom: (by1 * scale).round() as i32,
                                        };
                                        let _ = FillRect(mem_dc, &r, brush);
                                        SelectObject(mem_dc, old_brush);
                                        let _ = DeleteObject(brush);
                                        if neg {
                                            draw_neg_bar_outline(
                                                mem_dc, &r, scale,
                                            );
                                        }
                                    }
                                }
                            }
                        } else {
                        let bar_w = pitch / (n_ser as f64 + 1.5);
                        let cluster_w = bar_w * n_ser as f64;
                        let vary_points = chart.series.len() == 1;
                        for (si, series) in chart.series.iter().enumerate() {
                            for (ci, v) in series.values.iter().enumerate() {
                                let accent_idx = if vary_points { ci } else { si };
                                let col_hex = pres
                                    .theme_colors
                                    .get(&format!("accent{}", accent_idx + 1))
                                    .map(|s| s.as_str())
                                    .or_else(|| DEFAULT_ACCENT.get(accent_idx).copied());
                                if let Some(rgb) = col_hex.and_then(parse_hex_rgb) {
                                    // "Invert if negative" is ON by default in
                                    // Word: a negative bar is painted WHITE and
                                    // keeps only its outline (N1 Q2 / all of N2
                                    // render as #FFFFFF with a thin dark edge).
                                    let neg = *v < 0.0;
                                    let rgb = if neg { (255u8, 255u8, 255u8) } else { rgb };
                                    let brush = CreateSolidBrush(COLORREF(colorref(
                                        rgb.0, rgb.1, rgb.2,
                                    )));
                                    let old_brush = SelectObject(mem_dc, brush);
                                    let cat_center =
                                        plot_left + pitch * (ci as f64 + 0.5);
                                    let bx0 = cat_center
                                        - cluster_w / 2.0
                                        + si as f64 * bar_w;
                                    // Bars grow from the ZERO line, downward
                                    // when the value is negative (N1: Q2's -8.5
                                    // runs 280.97..334.56, i.e. below zero).
                                    let vy = val_y(*v);
                                    let (by0, by1) = if *v >= 0.0 {
                                        (vy, zero_y)
                                    } else {
                                        (zero_y, vy)
                                    };
                                    let r = RECT {
                                        left: (bx0 * scale).round() as i32,
                                        top: (by0 * scale).round() as i32,
                                        right: ((bx0 + bar_w) * scale).round() as i32,
                                        bottom: (by1 * scale).round() as i32,
                                    };
                                    let _ = FillRect(mem_dc, &r, brush);
                                    SelectObject(mem_dc, old_brush);
                                    let _ = DeleteObject(brush);
                                    if neg {
                                        draw_neg_bar_outline(mem_dc, &r, scale);
                                    }
                                }
                            }
                        }
                        }

                        // Data labels (c:dLbls): Word renders each bar's
                        // value in Calibri 18pt black, centred on the bar,
                        // positioned by c:dLblPos (outEnd / ctr / inEnd).
                        // Measured 2026-08-06 (chart_dlbls S1-S4, PDF
                        // baseline vs bar-top): OUTSIDE_END baseline
                        // ~ bar_top - 9.28, INSIDE_END ~ bar_top + 21.70,
                        // CENTER ~ bar vertical centre + 6.2. Format:
                        // numFmt "0.0%" -> value*100 one-decimal + "%".
                        if chart.has_data_labels && chart.show_val {
                            let num_fmt = chart.number_format.clone();
                            let format_label = |v: f64| -> String {
                                if num_fmt == "0.0%" {
                                    format!("{:.1}%", v * 100.0)
                                } else if num_fmt == "0%" {
                                    format!(
                                        "{}%",
                                        (v * 100.0).round() as i64
                                    )
                                } else {
                                    // Word prints the RAW value, no rounding:
                                    // chart_bar p5 reads "19.2 21.4 16.7" and
                                    // chart_stacked100_dlbls p1 "…10.5 15 12.3"
                                    // (15.0 loses its trailing zero).
                                    format!("{}", v)
                                }
                            };
                            // Default data-label position: STACKED charts
                            // centre their labels (COM position = -4108,
                            // chart_dlbls S5); CLUSTERED place them above
                            // the bar (OUTSIDE_END, S1).
                            let dlbl_pos = if chart.datalabel_position.is_empty() {
                                if is_stacked {
                                    "ctr"
                                } else {
                                    "outEnd"
                                }
                            } else {
                                chart.datalabel_position.as_str()
                            };
                            if is_stacked {
                                let bar_w = pitch * 0.4;
                                for ci in 0..n_cat {
                                    let cat_center =
                                        plot_left + pitch * (ci as f64 + 0.5);
                                    let bx0 = cat_center - bar_w / 2.0;
                                    let bar_center = bx0 + bar_w / 2.0;
                                    // Mirror the drawing block: each sign
                                    // stacks away from the zero line.  With
                                    // no negatives cum_neg stays 0 and this is
                                    // the previous plot_bot-anchored geometry.
                                    let mut cum_pos = 0.0f64;
                                    let mut cum_neg = 0.0f64;
                                    for s in chart.series.iter() {
                                        let v = s
                                            .values
                                            .get(ci)
                                            .copied()
                                            .unwrap_or(0.0);
                                        if v == 0.0 {
                                            continue;
                                        }
                                        let seg_h = if is_100pct {
                                            let sum_cat = chart
                                                .series
                                                .iter()
                                                .map(|s| {
                                                    s.values
                                                        .get(ci)
                                                        .copied()
                                                        .unwrap_or(0.0)
                                                })
                                                .sum::<f64>();
                                            if sum_cat > 0.0 {
                                                v / sum_cat * plot_h
                                            } else {
                                                0.0
                                            }
                                        } else {
                                            (v / axis_span * plot_h).abs()
                                        };
                                        let (by0, by1) = if v > 0.0 {
                                            let b1 = zero_y - cum_pos;
                                            let b0 = b1 - seg_h;
                                            cum_pos += seg_h;
                                            (b0, b1)
                                        } else {
                                            let b0 = zero_y + cum_neg;
                                            let b1 = b0 + seg_h;
                                            cum_neg += seg_h;
                                            (b0, b1)
                                        };
                                        let _ = by1;
                                        let text = format_label(v);
                                        let lw = font_adv::line_hmtx_width_pt(
                                            &text,
                                            axis_fs,
                                            axis_family,
                                        )
                                        .unwrap_or_else(|| {
                                            text.chars().count() as f32
                                                * axis_fs
                                                * 0.5
                                        }) as f64;
                                        let lx = bar_center - lw / 2.0;
                                        let baseline = match dlbl_pos {
                                            "inEnd" => by0 + 21.70,
                                            "ctr" => {
                                                by0 + seg_h / 2.0 + 6.2
                                            }
                                            _ => by0 - 9.28, // outEnd
                                        };
                                        draw_text_baseline(
                                            mem_dc,
                                            (lx * scale).round() as i32,
                                            baseline as f32,
                                            &text,
                                            axis_fs,
                                            axis_family,
                                            None,
                                            scale,
                                        );
                                    }
                                }
                            } else {
                                let bar_w = pitch / (n_ser as f64 + 1.5);
                                let cluster_w = bar_w * n_ser as f64;
                                for ci in 0..n_cat {
                                    let cat_center =
                                        plot_left + pitch * (ci as f64 + 0.5);
                                    for (si, s) in chart.series.iter().enumerate()
                                    {
                                        let v = s
                                            .values
                                            .get(ci)
                                            .copied()
                                            .unwrap_or(0.0);
                                        if v == 0.0 {
                                            continue;
                                        }
                                        let bx0 = cat_center
                                            - cluster_w / 2.0
                                            + si as f64 * bar_w;
                                        let bar_center = bx0 + bar_w / 2.0;
                                        let vy = val_y(v);
                                        let (by0, by1) = if v > 0.0 {
                                            (vy, zero_y)
                                        } else {
                                            (zero_y, vy)
                                        };
                                        let bar_h = by1 - by0;
                                        let text = format_label(v);
                                        let lw = font_adv::line_hmtx_width_pt(
                                            &text,
                                            axis_fs,
                                            axis_family,
                                        )
                                        .unwrap_or_else(|| {
                                            text.chars().count() as f32
                                                * axis_fs
                                                * 0.5
                                        }) as f64;
                                        let lx = bar_center - lw / 2.0;
                                        // The measured offsets are for a bar
                                        // that grows UP.  A negative bar's
                                        // "outer end" is its BOTTOM edge, so
                                        // the label mirrors: the measured 9.28
                                        // leaves an ink gap of 9.28 - descent
                                        // (0.211*fs) above the edge, and the
                                        // mirror puts the same gap below it
                                        // (baseline = edge + gap + ascent).
                                        // UNMEASURED - no probe arm pairs a
                                        // negative bar with c:dLbls.
                                        let baseline = if v > 0.0 {
                                            match dlbl_pos {
                                                "inEnd" => by0 + 21.70,
                                                "ctr" => by0 + bar_h / 2.0 + 6.2,
                                                _ => by0 - 9.28, // outEnd
                                            }
                                        } else {
                                            let fs = axis_fs as f64;
                                            let gap = 9.28 - 0.211 * fs;
                                            match dlbl_pos {
                                                "inEnd" => by1 - 21.70 + 12.4,
                                                "ctr" => by0 + bar_h / 2.0 + 6.2,
                                                _ => by1 + gap + 0.75 * fs,
                                            }
                                        };
                                        draw_text_baseline(
                                            mem_dc,
                                            (lx * scale).round() as i32,
                                            baseline as f32,
                                            &text,
                                            axis_fs,
                                            axis_family,
                                            None,
                                            scale,
                                        );
                                    }
                                }
                            }
                        }

                        // Category names centred on each category centre.
                        for (ci, name) in chart.categories.iter().enumerate() {
                            let cat_center =
                                plot_left + pitch * (ci as f64 + 0.5);
                            let lw = font_adv::line_hmtx_width_pt(name, axis_fs, axis_family)
                                .unwrap_or_else(|| {
                                    name.chars().count() as f32 * axis_fs * 0.5
                                }) as f64;
                            let lx = cat_center - lw / 2.0;
                            // Category names hang off the ZERO line (N1: bars
                            // start at 280.97 and the names sit at 309.65;
                            // N8: zero 234.29, names 262.94).  zero_y ==
                            // plot_bot whenever the data has no negatives.
                            draw_text_baseline(
                                mem_dc,
                                (lx * scale).round() as i32,
                                (zero_y + 28.67) as f32,
                                name,
                                axis_fs,
                                axis_family,
                                None,
                                scale,
                            );
                        }

                        // EXPLICIT <c:title> text: Word draws it as Arial
                        // 18pt (regular), centred on the frame, baseline
                        // sy+24.43, and it suppresses the automatic
                        // series-name title (chart_title / chart_title2
                        // render-truth 2026-08-07: origin=(194.66,96.43),
                        // frame_cx = 270.06, plot_top = sy+45.69).
                        if let Some(title) = &chart.explicit_title {
                            let tfs = 18.0f32;
                            let lw = font_adv::line_hmtx_width_pt(title, tfs, "Arial")
                                .unwrap_or_else(|| {
                                    title.chars().count() as f32 * tfs * 0.5
                                }) as f64;
                            let frame_cx = sx + sw / 2.0;
                            draw_text_baseline_w(
                                mem_dc,
                                ((frame_cx - lw / 2.0) * scale).round() as i32,
                                (sy + 24.43) as f32,
                                title,
                                tfs,
                                "Arial",
                                None,
                                scale,
                                400,
                            );
                        }

                        // Automatic chart title: with a SINGLE series Word
                        // shows the series name as the chart title
                        // (Calibri-Bold 21.62pt, centred on the frame,
                        // baseline = sh.y+28.03 - chart1/chart2b render-truth
                        // 2026-08-06: 'Series 1' / 'Revenue' at
                        // origin=(235.37/231.17,100.03), frame_cx = sh.x+sh.w/2
                        // = 270.0). A <c:legend> is drawn only when declared
                        // (none of the 4 probes carry one -> not drawn).
                        // Gated on has_auto_title: a MULTI-series chart (e.g.
                        // chart2 / chart3 / chart_stacked / percentStacked)
                        // has NO automatic title (render-truth: only the
                        // single-series chart1 / chart2b render one).
                        if has_auto_title && !has_explicit_title {
                        if let Some(first) = chart.series.first() {
                            let tfs = 21.62f32;
                            let lw = font_adv::line_hmtx_width_pt(
                                &first.name,
                                tfs,
                                axis_family,
                            )
                            .unwrap_or_else(|| {
                                first.name.chars().count() as f32 * tfs * 0.5
                            }) as f64;
                            let frame_cx = sx + sw / 2.0;
                            draw_text_baseline_w(
                                mem_dc,
                                ((frame_cx - lw / 2.0) * scale).round() as i32,
                                (sy + 28.03) as f32,
                                &first.name,
                                tfs,
                                axis_family,
                                None,
                                scale,
                                700,
                            );
                        }
                        }

                        // Legend (when <c:legend> is declared): per-series
                        // swatch (accent colour 9.89x9.89pt) + series name
                        // (Calibri 18pt). Placement DERIVED from Word
                        // get_drawings + rawdict on chart_legend (2 series)
                        // and chart_legend3 (3 series), 2026-08-06:
                        //   right-aligned overlay: legend_right = frame right
                        //     - 10.0; swatch_x1 = legend_right - max_label_w
                        //     - 4.62 (max over series names, Calibri 18pt);
                        //     swatch_x0 = swatch_x1 - 9.89; label_x0 =
                        //     swatch_x1 + 4.62
                        //   vertically centred on the frame:
                        //     legend_total_h = (n_ser-1)*27.75 + 9.89;
                        //     legend_y0 = sy + sh/2 - legend_total_h/2
                        //     (2 ser 197.18 / 3 ser 183.30 = Word EXACT)
                        //   row pitch 27.75; label baseline = swatch bottom
                        //     + 0.28
                        //   plot area is NOT shrunk (overlay; COM
                        //     Legend.IncludeInLayout = False; plot_top/X-axis
                        //     identical to the no-legend chart3)
                        if chart.has_legend {
                            let lfs = 18.0f32;
                            let max_label_w = chart
                                .series
                                .iter()
                                .map(|s| {
                                    font_adv::line_hmtx_width_pt(
                                        &s.name,
                                        lfs,
                                        axis_family,
                                    )
                                    .unwrap_or_else(|| {
                                        s.name.chars().count() as f32 * lfs * 0.5
                                    }) as f64
                                })
                                .fold(0.0f64, f64::max);
                            let swatch_w = 9.89f64;
                            let gap = 4.62f64;
                            let row_pitch = 27.75f64;
                            let legend_right = (sx + sw) - 10.0;
                            let swatch_x1 = legend_right - max_label_w - gap;
                            let swatch_x0 = swatch_x1 - swatch_w;
                            let label_x0 = swatch_x1 + gap;
                            let legend_total_h =
                                (n_ser as f64 - 1.0) * row_pitch + swatch_w;
                            let legend_y0 = (sy + shh / 2.0)
                                - legend_total_h / 2.0;
                            for (si, series) in
                                chart.series.iter().enumerate()
                            {
                                let sw_y =
                                    legend_y0 + si as f64 * row_pitch;
                                let col_hex = pres
                                    .theme_colors
                                    .get(&format!("accent{}", si + 1))
                                    .map(|s| s.as_str())
                                    .or_else(|| DEFAULT_ACCENT.get(si).copied());
                                if let Some(rgb) =
                                    col_hex.and_then(parse_hex_rgb)
                                {
                                    let brush = CreateSolidBrush(COLORREF(
                                        colorref(rgb.0, rgb.1, rgb.2),
                                    ));
                                    let old_brush =
                                        SelectObject(mem_dc, brush);
                                    let r = RECT {
                                        left: (swatch_x0 * scale).round() as i32,
                                        top: (sw_y * scale).round() as i32,
                                        right: (swatch_x1 * scale).round() as i32,
                                        bottom: ((sw_y + swatch_w) * scale).round() as i32,
                                    };
                                    let _ = FillRect(mem_dc, &r, brush);
                                    SelectObject(mem_dc, old_brush);
                                    let _ = DeleteObject(brush);
                                }
                                let label_baseline =
                                    sw_y + swatch_w + 0.28;
                                draw_text_baseline(
                                    mem_dc,
                                    (label_x0 * scale).round() as i32,
                                    label_baseline as f32,
                                    &series.name,
                                    lfs,
                                    axis_family,
                                    None,
                                    scale,
                                );
                            }
                        }

                        // Axis lines + ticks (chart1 render-truth, fitz
                        // get_drawings 2026-08-06, per-item line paths):
                        //   Y axis line: vertical (plot_left, plot_top) ->
                        //     (plot_left, plot_bot)
                        //   X axis line: horizontal (plot_left, plot_bot) ->
                        //     (plot_right, plot_bot)
                        //   Y ticks: 6, x from plot_left-5.7 to plot_left at
                        //     y = plot_top + i*plot_h/5 (i=0..=5)
                        //   X ticks: n_cat+1, y from plot_bot to plot_bot+5.7
                        //     at x = plot_left + i*pitch (i=0..=n_cat; the
                        //     CATEGORY BOUNDARIES, not the centres)
                        let axis_pen = CreatePen(
                            PS_SOLID,
                            2,
                            COLORREF(colorref(0, 0, 0)),
                        );
                        let old_axis_pen = SelectObject(mem_dc, axis_pen);
                        let _ = SelectObject(mem_dc, GetStockObject(NULL_BRUSH));

                        let pl = (plot_left * scale).round() as i32;
                        let pt = (plot_top * scale).round() as i32;
                        let pr = (plot_right * scale).round() as i32;
                        let pb = (plot_bot * scale).round() as i32;

                        let _ = MoveToEx(mem_dc, pl, pt, None);
                        let _ = LineTo(mem_dc, pl, pb);
                        let _ = MoveToEx(mem_dc, pl, pb, None);
                        let _ = LineTo(mem_dc, pr, pb);
                        // Plot frame TOP edge (render-truth: Word paints a line
                        // (113.45,123.35)->(457.00,123.35) at plot_top; the chart
                        // frame / plot frame right edge are declared but NOT
                        // painted - pixel-checked 2026-08-06)
                        let _ = MoveToEx(mem_dc, pl, pt, None);
                        let _ = LineTo(mem_dc, pr, pt);

                        for i in 0..=axis_steps {
                            let tick_y =
                                plot_bot - plot_h * i as f64 / axis_steps as f64;
                            let ty = (tick_y * scale).round() as i32;
                            let _ = MoveToEx(mem_dc, ((plot_left - 5.7) * scale).round() as i32, ty, None);
                            let _ = LineTo(mem_dc, pl, ty);
                        }
                        // Category ticks hang from the CATEGORY AXIS, which sits
                        // at value 0 (crosses=autoZero) - not at the bottom of
                        // the plot.  N8 render-truth: the four ticks run
                        // 234.29 -> 240.00 = zero_y -> zero_y+5.71, while the
                        // plot bottom is 344.01.  With no negatives zero_y ==
                        // plot_bot, so this is the previous geometry.
                        let zt = (zero_y * scale).round() as i32;
                        for i in 0..=n_cat {
                            let tick_x = plot_left + pitch * i as f64;
                            let tx = (tick_x * scale).round() as i32;
                            let _ = MoveToEx(mem_dc, tx, zt, None);
                            let _ = LineTo(mem_dc, tx, ((zero_y + 5.7) * scale).round() as i32);
                        }

                        SelectObject(mem_dc, old_axis_pen);
                        let _ = DeleteObject(axis_pen);
                        }
                    }
                    ShapeContent::Image { data, content_type } => {
                        // Decode the embedded image (PNG/JPEG media part) and
                        // draw it scaled into the shape rect. rotation=0 only
                        // for now (a rotation-aware path is left for a later
                        // step; deck background images are typically unrotated).
                        //
                        // Image fill geometry (a:blipFill) — Word render-truth
                        // (01__Biology deck, 2026-08):
                        //   - a:srcRect (l/t/r/b, 0..1) CROPS THE SOURCE. The
                        //     full-bleed background PNG is 500x500 but the deck
                        //     crops t=21.875% b=21.874% -> visible 500x281.26,
                        //     whose aspect 1.778 matches the 720x405 box exactly
                        //     (keeps the image un-distorted).
                        //   - a:stretch/a:fillRect (l/t/r/b, may be negative)
                        //     INSETS THE DESTINATION. The photo expands it
                        //     t=b=-22.646% -> dest 296.99 x 410.29pt, whose
                        //     aspect 0.724 matches the portrait 977x1350 source.
                        // Both are "keep aspect" mechanisms Word uses for
                        // pictures; without them StretchDIBits distorts.
                        let (sl, st, sr, sb) = sh
                            .src_rect
                            .map(|(l, t, r, b)| (l, t, r, b))
                            .unwrap_or((0.0, 0.0, 0.0, 0.0));
                        // S-SRCDEGEN (2026-08-26): a crop that leaves NOTHING
                        // shows nothing. `a:srcRect` gives each edge as a
                        // fraction to trim, so `l + r >= 1` (or `t + b >= 1`)
                        // selects an empty region -- and d37's layout 4 asks for
                        // exactly that, `l=74246 r=118526` (the right edge
                        // trimmed by 118%, past the left one).
                        //
                        // PowerPoint draws nothing for it. Oxi ignored the
                        // degeneracy and painted the WHOLE picture, so a
                        // mostly-opaque 2048x1636 image covered the right half
                        // of the slide in pale grey where the deck's tan
                        // background belongs. d37 s5 -- the deck's worst slide
                        // at 0.8775 -- is that rectangle.
                        //
                        // Only 2 of the corpus's 412 srcRects are degenerate and
                        // both are this one crop (the layout's, and the copy
                        // slide 5 carries), but the cost of drawing it is half a
                        // slide.
                        if srcdegen_on() && (sl + sr >= 1.0 || st + sb >= 1.0) {
                            continue;
                        }
                        let (dl, dt, dr, db) = sh
                            .fill_rect
                            .map(|(l, t, r, b)| (l, t, r, b))
                            .unwrap_or((0.0, 0.0, 0.0, 0.0));
                        match image::load_from_memory(data) {
                            Ok(dyn_img) => {
                                let mut rgba = dyn_img.to_rgba8();
                                // `a:alphaModFix/@amt` scales the WHOLE
                                // picture's opacity (d32's title map is a city
                                // plan at 7%: a dark texture in PowerPoint, a
                                // stark white overlay when it is ignored).
                                if let Some(a) = sh.image_alpha.filter(|_| imgalpha_on()) {
                                    if a < 1.0 {
                                        for p in rgba.pixels_mut() {
                                            p[3] = (p[3] as f32 * a).round().clamp(0.0, 255.0)
                                                as u8;
                                        }
                                    }
                                }
                                let rgba = rgba;
                                let (iw, ih) = (rgba.width() as i32, rgba.height() as i32);
                                if iw <= 0 || ih <= 0 {
                                    continue;
                                }
                                // Source sub-rect after srcRect crop (px).
                                let sl = sl as f64;
                                let st = st as f64;
                                let sr = sr as f64;
                                let sb = sb as f64;
                                let dl = dl as f64;
                                let dt = dt as f64;
                                let dr = dr as f64;
                                let db = db as f64;
                                let sx0 = (iw as f64 * sl).round() as i32;
                                let sy0 = (ih as f64 * st).round() as i32;
                                let sw = ((iw as f64 * (1.0 - sl - sr)).round() as i32)
                                    .max(1);
                                let shh = ((ih as f64 * (1.0 - st - sb)).round() as i32)
                                    .max(1);
                                // Destination rect after fillRect insets (px;
                                // negative inset expands beyond the box).
                                let (dx, dy, dw, dh) = if imgrect_on() {
                                    // From the EXACT points, covering every
                                    // pixel the rectangle touches.
                                    let bw = sh.width as f64 * scale;
                                    let bh = sh.height as f64 * scale;
                                    let x0 = sh.x as f64 * scale + bw * dl;
                                    let y0 = sh.y as f64 * scale + bh * dt;
                                    let x1 = x0 + bw * (1.0 - dl - dr);
                                    let y1 = y0 + bh * (1.0 - dt - db);
                                    // ★A hair of coverage is not coverage. The
                                    // `embedsplit` master picture is
                                    // 720.0002pt wide -- 1500.0005px, five
                                    // ten-thousandths past the slide's right
                                    // edge -- and a bare `ceil` gave it a
                                    // 1501st column, stretching the whole
                                    // background 0.07% and costing every page
                                    // of that probe 0.0036. An edge within
                                    // EPS of a pixel boundary belongs to the
                                    // pixel it is on.
                                    const EPS: f64 = 0.02;
                                    let left = (x0 + EPS).floor() as i32;
                                    let top = (y0 + EPS).floor() as i32;
                                    (
                                        left,
                                        top,
                                        ((x1 - EPS).ceil() as i32 - left).max(1),
                                        ((y1 - EPS).ceil() as i32 - top).max(1),
                                    )
                                } else {
                                    (
                                        (x as f64 + ew as f64 * dl).round() as i32,
                                        (y as f64 + eh as f64 * dt).round() as i32,
                                        ((ew as f64 * (1.0 - dl - dr)).round() as i32).max(1),
                                        ((eh as f64 * (1.0 - dt - db)).round() as i32).max(1),
                                    )
                                };
                                // A blipFill belongs to the SHAPE, so it is
                                // clipped to the shape's outline -- both the
                                // part of the box the outline does not cover
                                // and the part a negative fillRect pushes
                                // outside the box (S-BLIPCLIP render-truth).
                                let clipped = clip_to_geometry_gdi(mem_dc, sh, scale);
                                // The picture turns and mirrors with its shape
                                // (S-IMGROT render-truth). Resample it into a
                                // page-aligned buffer whose margins are
                                // transparent, then composite that -- an
                                // axis-aligned blit cannot express a rotation.
                                let turns = imgrot_on()
                                    && (sh.flip_h || sh.flip_v || sh.rotation != 0.0)
                                    && (sh.rot_with_shape || sh.flip_h || sh.flip_v);
                                let angle = if sh.rot_with_shape { sh.rotation as f64 } else { 0.0 };
                                let turned = if turns {
                                    transform_picture(
                                        &rgba,
                                        (sx0, sy0, sw, shh),
                                        (dx as f64, dy as f64, dw as f64, dh as f64),
                                        (
                                            (sh.x + sh.width / 2.0) as f64 * scale,
                                            (sh.y + sh.height / 2.0) as f64 * scale,
                                        ),
                                        angle,
                                        sh.flip_h,
                                        sh.flip_v,
                                    )
                                } else {
                                    None
                                };
                                // A picture whose media carries transparency
                                // has to be COMPOSITED over the page -- the
                                // opaque blit below drops the alpha byte and
                                // paints the RGB stored under the transparent
                                // pixels. A fully opaque picture keeps the
                                // exact SRCCOPY path (byte-identical), and a
                                // failed AlphaBlend falls back to it too.
                                // Resample once, here, for both blits.
                                //
                                // GDI's own samplers are what made this
                                // necessary: `AlphaBlend` ignores the DC's
                                // stretch mode entirely, and
                                // `SetStretchBltMode(HALFTONE)` is a measured
                                // no-op on a 32-bpp BI_RGB `StretchDIBits`
                                // (setting it leaves d28 slide 3 byte-identical,
                                // same sha). Both then drop-sample.
                                //
                                // It applies to any scale, not only a shrink:
                                // the page is drawn at `supersample`x, so d28's
                                // 2048px-wide engraved portrait is ENLARGED into
                                // a 3000px box before the final 2x downsample
                                // takes it to 1500 -- which is why every
                                // shrink-only gate missed it.
                                //
                                // Two other models were measured and are worse.
                                // Band-limiting to the final size and scaling up
                                // (Triangle, then Nearest so the closing box
                                // filter is exact) costs d28 a further -0.0041.
                                // Filtering only a TRUE enlargement -- source
                                // narrower than the final box -- gives d28 back
                                // but loses most of the corpus gain (d22 0.9170
                                // -> 0.9123, d38 0.9739 -> 0.9728). Straight to
                                // the supersampled box wins by measurement.
                                let scaled;
                                let (rgba, sx0, sy0, sw, shh, iw, ih) = if alphasmooth_on()
                                    && turned.is_none()
                                    && (dw != sw || dh != shh)
                                    && dw > 0
                                    && dh > 0
                                {
                                    let sub = image::imageops::crop_imm(
                                        &rgba, sx0 as u32, sy0 as u32, sw as u32, shh as u32,
                                    )
                                    .to_image();
                                    scaled = image::imageops::resize(
                                        &sub,
                                        dw as u32,
                                        dh as u32,
                                        image::imageops::FilterType::Triangle,
                                    );
                                    (&scaled, 0, 0, dw, dh, dw, dh)
                                } else {
                                    (&rgba, sx0, sy0, sw, shh, iw, ih)
                                };
                                let composited = match &turned {
                                    // The resampled buffer always has
                                    // transparent corners, so it must go
                                    // through the compositing blit.
                                    Some((buf, bx, by)) => alpha_blit(
                                        mem_dc,
                                        *bx,
                                        *by,
                                        buf.width() as i32,
                                        buf.height() as i32,
                                        0,
                                        0,
                                        buf.width() as i32,
                                        buf.height() as i32,
                                        buf.width() as i32,
                                        buf.height() as i32,
                                        buf,
                                    ),
                                    None => {
                                        std::env::var("OXI_ALPHABLEND_DISABLE").is_err()
                                            && rgba.pixels().any(|p| p[3] != 255)
                                            && alpha_blit(
                                                mem_dc, dx, dy, dw, dh, sx0, sy0, sw, shh, iw,
                                                ih, &rgba,
                                            )
                                    }
                                };
                                if !composited {
                                // RGBA -> BGRA for the 32-bpp DIB
                                let mut bgra = Vec::with_capacity((iw * ih * 4) as usize);
                                for p in rgba.pixels() {
                                    bgra.push(p[2]);
                                    bgra.push(p[1]);
                                    bgra.push(p[0]);
                                    bgra.push(p[3]);
                                }
                                let bmi = BITMAPINFO {
                                    bmiHeader: BITMAPINFOHEADER {
                                        biSize: std::mem::size_of::<BITMAPINFOHEADER>() as u32,
                                        biWidth: iw,
                                        biHeight: -ih, // top-down
                                        biPlanes: 1,
                                        biBitCount: 32,
                                        biCompression: 0, // BI_RGB
                                        ..Default::default()
                                    },
                                    ..Default::default()
                                };
                                // ★StretchDIBits measures ySrc from the
                                // BOTTOM of the DIB even when the bitmap is
                                // top-down (negative biHeight), so a srcRect
                                // top crop must be converted. PowerPoint
                                // render-truth (srcrect probe F2/F3,
                                // 2026-08-18): with `b="50000"` PowerPoint
                                // keeps the source's TOP half (rules land at
                                // 0.20/0.40/0.60/0.80/1.00) while Oxi kept the
                                // BOTTOM half -- t and b came out swapped, and
                                // d30 slide 16's photo was the 1.15x vertical
                                // zoom that produced.
                                let sy_bottom_up = if srcrect_flip_on() {
                                    (ih - (sy0 + shh)).max(0)
                                } else {
                                    sy0
                                };
                                let old_stretch = begin_smooth_blit(mem_dc, (sw, shh), (dw, dh));
                                let _ = StretchDIBits(
                                    mem_dc,
                                    dx,
                                    dy,
                                    dw,
                                    dh,
                                    sx0,
                                    sy_bottom_up,
                                    sw,
                                    shh,
                                    Some(bgra.as_ptr() as *const _),
                                    &bmi,
                                    DIB_RGB_COLORS,
                                    SRCCOPY,
                                );
                                end_smooth_blit(mem_dc, old_stretch);
                                }
                                if clipped {
                                    let _ = SelectClipRgn(mem_dc, None);
                                }
                            }
                            Err(_) => {}
                        }
                    }
                    _ => {}
                }
            }

            // Extract bitmap pixels
            let mut bmi = BITMAPINFO {
                bmiHeader: BITMAPINFOHEADER {
                    biSize: std::mem::size_of::<BITMAPINFOHEADER>() as u32,
                    biWidth: w,
                    biHeight: -h, // top-down
                    biPlanes: 1,
                    biBitCount: 32,
                    biCompression: 0, // BI_RGB
                    ..Default::default()
                },
                ..Default::default()
            };

            let mut pixels = vec![0u8; (w * h * 4) as usize];
            GetDIBits(
                mem_dc,
                bitmap,
                0,
                h as u32,
                Some(pixels.as_mut_ptr() as *mut _),
                &mut bmi,
                DIB_RGB_COLORS,
            );

            // BGRA -> RGB
            let mut rgb_pixels = Vec::with_capacity((w * h * 3) as usize);
            for i in 0..(w * h) as usize {
                rgb_pixels.push(pixels[i * 4 + 2]); // R
                rgb_pixels.push(pixels[i * 4 + 1]); // G
                rgb_pixels.push(pixels[i * 4]); // B
            }

            let img = image::RgbImage::from_raw(w as u32, h as u32, rgb_pixels)
                .expect("Failed to create image");
            let final_img = if supersample > 1 && (w as u32 != out_w || h as u32 != out_h) {
                let dynamic = image::DynamicImage::ImageRgb8(img);
                dynamic
                    .resize_exact(out_w, out_h, image::imageops::FilterType::Lanczos3)
                    .to_rgb8()
            } else {
                img
            };
            let out_path = format!("{}_s{}.png", prefix, si + 1);
            final_img.save(&out_path).expect("Failed to save PNG");
            eprintln!(
                "  Saved {} ({}x{})",
                out_path,
                final_img.width(),
                final_img.height()
            );

            SelectObject(mem_dc, old_bmp);
            let _ = DeleteObject(bitmap);
            let _ = DeleteDC(mem_dc);
            let _ = ReleaseDC(HWND(std::ptr::null_mut()), screen_dc);
        }
    }
}

#[cfg(windows)]
fn draw_text_line(
    dc: windows::Win32::Graphics::Gdi::HDC,
    x: i32,
    y: i32,
    text: &str,
    font_size: f32,
    family: &str,
    color: Option<&str>,
    scale: f64,
) {
    use windows::Win32::Foundation::*;
    use windows::Win32::Graphics::Gdi::*;
    use windows::core::PCWSTR;

    let height = (font_size as f64 * scale).round() as i32;
    // Negative lfHeight = character height; CJK needs the eastAsia charset so
    // Japanese text resolves. wcsdup via UTF-16.
    let wide: Vec<u16> = family.encode_utf16().collect();
    let mut family_buf = vec![0u16; wide.len() + 1];
    family_buf[..wide.len()].copy_from_slice(&wide);

    let font = unsafe {
        CreateFontW(
            -height,
            0,
            0,
            0,
            // FW_NORMAL
            400,
            0,
            0,
            0,
            // DEFAULT_CHARSET
            1,
            0, // OUT_DEFAULT_PRECIS
            0, // CLIP_DEFAULT_PRECIS
            5, // CLEARTYPE_QUALITY — matches Word GDI rendering
            0, // DEFAULT_PITCH
            PCWSTR(family_buf.as_ptr()),
        )
    };
    if font.is_invalid() {
        return;
    }

    let rgb = color.and_then(parse_hex_rgb).unwrap_or((0, 0, 0));
    let old_color = unsafe { SetTextColor(dc, COLORREF(colorref(rgb.0, rgb.1, rgb.2))) };
    let old_font = unsafe { SelectObject(dc, font) };

    let wtext: Vec<u16> = text.encode_utf16().collect();
    unsafe {
        let _ = TextOutW(dc, x, y, &wtext);
    }

    unsafe {
        SelectObject(dc, old_font);
        SetTextColor(dc, old_color);
        let _ = DeleteObject(font);
    }
}

// ---------------------------------------------------------------------------
// Text-frame layout (Spec #4): wrapped lines + baseline positions, computed
// with GDI font metrics so the wrapped line breaks and line advances match
// PowerPoint. Used by BOTH the GDI renderer (draw) and --dump-layout (JSON).
//
// Measured models (Ra loop, spec4d/spec4e):
//   * default line advance (no lnSpc)      = font_size x 1.2
//   * explicit lnSpc n                     = font_size x 1.2 x n  (linear)
//   * first-line baseline, n != 1 (multi)  = text_area_top + 0.75 x advance
//   * first-line baseline, n == 1 (single) = text_area_top + A_font x fs
//       A_font = hhea_asc + hhea_lineGap (font-dependent; table below)
//   * space_before / space_after           = added around each paragraph
//   * inner insets                         = shape l_ins/r_ins/t_ins/b_ins
//       (a:bodyPr; placeholders default to top/bottom 3.6pt, left/right 7.2pt)
// ---------------------------------------------------------------------------

/// Create a GDI font for the given family/size (negative lfHeight = char height).
#[cfg(windows)]
fn create_font_for_w(
    family: &str,
    font_size: f32,
    weight: i32,
    scale: f64,
) -> windows::Win32::Graphics::Gdi::HFONT {
    create_font_for_wi(family, font_size, weight, false, scale)
}

/// As `create_font_for_w`, with the slant. `CreateFontW`'s italic argument was
/// hardcoded to 0, so `a:rPr/@i` reached the MEASUREMENT (`runtime_advance_em`
/// takes it) and never the page -- italic text was drawn upright. Only three
/// paragraphs in the dev corpus are italic (d16, d17, d35, one each), so this
/// is closing the family rather than moving the number.
#[cfg(windows)]
fn create_font_for_wi(
    family: &str,
    font_size: f32,
    weight: i32,
    italic: bool,
    scale: f64,
) -> windows::Win32::Graphics::Gdi::HFONT {
    create_font_for_wiu(family, font_size, weight, italic, false, scale)
}

/// As `create_font_for_wi`, with the underline. GDI draws the rule itself, at
/// the position and thickness the face declares, which is what PowerPoint's
/// export shows for the hyperlinks these decks put in their instructions.
#[cfg(windows)]
fn create_font_for_wiu(
    family: &str,
    font_size: f32,
    weight: i32,
    italic: bool,
    underline: bool,
    scale: f64,
) -> windows::Win32::Graphics::Gdi::HFONT {
    use windows::Win32::Graphics::Gdi::*;
    use windows::core::PCWSTR;
    let height = (font_size as f64 * scale).round() as i32;
    // An embedded part registered for this exact style IS the bold / italic
    // face, so it is asked for plain.
    let asked = family;
    let want_italic = italic;
    let (family, weight, italic) =
        styled_face(family, weight >= 700, italic && paraitalic_on());
    if want_italic && std::env::var("OXI_DEBUG_EMBED").is_ok() {
        eprintln!(
            "FACE asked={asked:?} want_i=true -> {family:?} w={weight} i={italic} sz={font_size}"
        );
    }
    let wide: Vec<u16> = family.encode_utf16().collect();
    let mut family_buf = vec![0u16; wide.len() + 1];
    family_buf[..wide.len()].copy_from_slice(&wide);
    unsafe {
        CreateFontW(
            -height,
            0,
            0,
            0,
            weight,
            u32::from(italic),
            u32::from(underline && underline_on()),
            0,
            1,
            0,
            0,
            5,
            0,
            PCWSTR(family_buf.as_ptr()),
        )
    }
}

/// Regular-weight font (the common case).
#[cfg(windows)]
fn create_font_for(
    family: &str,
    font_size: f32,
    scale: f64,
) -> windows::Win32::Graphics::Gdi::HFONT {
    create_font_for_w(family, font_size, 400, scale)
}

/// Office 2016+ default accent colours. Used only as a fallback when the
/// theme's clrScheme does not declare accentN — real charts resolve their
/// series colours from `Presentation.theme_colors` (Spec #10).
const DEFAULT_ACCENT: [&str; 6] = [
    "4472C4", "ED7D31", "A5A5A5", "FFC000", "5B9BD5", "70AD47",
];

/// Per-series LINE chart (fill, border) colours, measured from a Word PDF
/// export of a 3-series LINE_MARKERS chart (2026-08-06):
///   S1 fill #4F81BD / border #4A7EBB
///   S2 fill #C0504D / border #BE4B48
///   S3 fill #9BBB59 / border #98B954
/// S4+ are UNMEASURED -> fall back to DEFAULT_ACCENT with a same-colour
/// border (the measured S1-S3 differ from DEFAULT_ACCENT; only add more
/// measured entries when a 4+-series specimen is verified).
fn line_series_colors(si: usize) -> (String, String) {
    const MEASURED: [(&str, &str); 3] = [
        ("4F81BD", "4A7EBB"),
        ("C0504D", "BE4B48"),
        ("9BBB59", "98B954"),
    ];
    if let Some((f, b)) = MEASURED.get(si) {
        (f.to_string(), b.to_string())
    } else {
        let accent = DEFAULT_ACCENT[si % DEFAULT_ACCENT.len()];
        (accent.to_string(), accent.to_string())
    }
}

/// Per-series LINE chart data-marker SHAPE (measured, 2026-08-06):
///   index 0 = diamond (4 lines joining the bbox side-midpoints)
///   index 1 = square (6.96x6.96 rect)
///   index 2 = triangle (3 lines: top, right, bottom-left)
///   index 3+ = diamond fallback (UNMEASURED; Word's marker cycle for
///   series beyond the measured three is not yet known).
fn line_marker_shape(si: usize) -> u8 {
    match si {
        0 => 0, // diamond
        1 => 1, // square
        2 => 2, // triangle
        _ => 0, // diamond fallback (unmeasured)
    }
}

/// Draw one 6.96pt data marker (per-series shape) centred at (cx, cy),
/// with the caller's brush selected (fill colour) and NULL_PEN (same-
/// colour stroke invisible, = Word's w=0.75 outline).
fn draw_line_marker(
    dc: windows::Win32::Graphics::Gdi::HDC,
    shape: u8,
    cx: f64,
    cy: f64,
    mr: f64,
    scale: f64,
) {
    use windows::Win32::Foundation::POINT;
    use windows::Win32::Graphics::Gdi::{Polygon, Rectangle};
    let hx = (cx * scale).round() as i32;
    let hy = (cy * scale).round() as i32;
    match shape {
        1 => {
            // square: 6.96x6.96 rect
            unsafe {
                let _ = Rectangle(
                    dc,
                    ((cx - mr) * scale).round() as i32,
                    ((cy - mr) * scale).round() as i32,
                    ((cx + mr) * scale).round() as i32,
                    ((cy + mr) * scale).round() as i32,
                );
            }
        }
        2 => {
            // triangle: top, right, bottom-left (3 lines)
            let pts = [
                POINT {
                    x: hx,
                    y: ((cy - mr) * scale).round() as i32,
                },
                POINT {
                    x: ((cx + mr) * scale).round() as i32,
                    y: ((cy + mr) * scale).round() as i32,
                },
                POINT {
                    x: ((cx - mr) * scale).round() as i32,
                    y: ((cy + mr) * scale).round() as i32,
                },
            ];
            unsafe {
                let _ = Polygon(dc, &pts);
            }
        }
        _ => {
            // diamond: top, right, bottom, left (4 lines)
            let pts = [
                POINT {
                    x: hx,
                    y: ((cy - mr) * scale).round() as i32,
                },
                POINT {
                    x: ((cx + mr) * scale).round() as i32,
                    y: hy,
                },
                POINT {
                    x: hx,
                    y: ((cy + mr) * scale).round() as i32,
                },
                POINT {
                    x: ((cx - mr) * scale).round() as i32,
                    y: hy,
                },
            ];
            unsafe {
                let _ = Polygon(dc, &pts);
            }
        }
    }
}

/// The legend's share of the chart width: a label wider than this wraps.
///
/// DERIVED (chart_legendwrap 2026-08-09): a category label made of one very
/// long word cannot break at a space, so Word force-breaks it at the last
/// character that fits -- the widest resulting line IS the cap.  Sweeping the
/// frame width over 200/240/280/320/360/396/440/500/560/600pt with a 4.12pt
/// glyph quantum leaves exactly one linear law, `cap = a*sw + b` with a in
/// [0.3321, 0.3436]; a = 1/3 lies inside it, and at a = 1/3 the intercept
/// window is [-21.74, -21.03), so:
///
///     cap = frame_width / 3 - 21.4pt
///
/// It is HEIGHT-independent (same cap at frame heights 180 / 288 / 400) and it
/// reproduces the six independent chart_doughnut_resid arms.  After wrapping,
/// the legend block shrinks to the WRAPPED width: the measured label x0 is
/// (sx+sw) - 10 - widest_line on every arm.
fn legend_label_cap(frame_w: f64) -> f64 {
    frame_w / 3.0 - 21.4
}

/// Wrap one legend label to `cap`: greedy at spaces, and a single word wider
/// than the cap is force-broken at the last character that fits (Word
/// 2026-08-09: 'Abcdefghijklmnop' -> 'Abcdefghijklm' + 'nop').
fn wrap_legend_label(text: &str, fs: f32, family: &str, cap: f64) -> Vec<String> {
    let width = |s: &str| {
        font_adv::line_hmtx_width_pt(s, fs, family)
            .unwrap_or_else(|| s.chars().count() as f32 * fs * 0.5) as f64
    };
    if cap <= 0.0 || width(text) <= cap {
        return vec![text.to_string()];
    }
    let mut out: Vec<String> = Vec::new();
    let mut line = String::new();
    for word in text.split(' ') {
        if !line.is_empty() {
            let cand = format!("{} {}", line, word);
            if width(&cand) <= cap {
                line = cand;
                continue;
            }
            out.push(std::mem::take(&mut line));
        }
        if width(word) <= cap {
            line = word.to_string();
            continue;
        }
        let mut cur = String::new();
        for ch in word.chars() {
            let mut probe = cur.clone();
            probe.push(ch);
            if !cur.is_empty() && width(&probe) > cap {
                out.push(std::mem::take(&mut cur));
            }
            cur.push(ch);
        }
        line = cur;
    }
    if !line.is_empty() {
        out.push(line);
    }
    if out.is_empty() {
        vec![text.to_string()]
    } else {
        out
    }
}

/// "Nice" ceiling for the value axis: the smallest multiple of a 1/2/5×10^k
/// step that is >= max. Chart1 render-truth: max 21.4 -> step 5 -> 25.
/// Format a value-axis label the way Word does: the number of decimals comes
/// from the tick step, and trailing zeros are trimmed ("0", "0.5", "1", "2.5"
/// for a 0.5 step -- chart_bar_axis slide 1 render-truth).
fn fmt_axis_value(v: f64, step: f64) -> String {
    let dec = if step >= 1.0 {
        0usize
    } else {
        (-step.log10()).ceil().max(0.0) as usize
    };
    let mut s = format!("{:.*}", dec, v);
    if dec > 0 {
        while s.ends_with('0') {
            s.pop();
        }
        if s.ends_with('.') {
            s.pop();
        }
    }
    s
}

/// HORIZONTAL value-axis auto scale -> (axis_max, division count).
///
/// DERIVED 2026-08-09 from a 14-arm Word sweep (chart_bar_axis: data range
/// 2.2..480 at a fixed frame, plus 196/342/546pt plot widths at range 34) plus
/// the 5 chart_bar pages. Word picks the FINEST 1/2/5x10^k step whose resulting
/// tick spacing along the axis is at least ~57pt; the axis maximum is that step
/// rounded up past the data PLUS 5% headroom. The sweep proves the rule is
/// axis-LENGTH dependent (range 34 gives 2 / 4 / 8 divisions at 196/342/546pt).
///
/// The 5% headroom is what makes the axis maximum come out at Word's value in
/// the arms a plain ceil(max/step)*step misses: data max 78 with a 20 step is
/// 100 in Word (ceil(78*1.05/20)*20), not 80.  With it, the threshold window
/// closes to (56.28, 57.76] over all 19 measured points and 57.0 sits inside;
/// the model then reproduces 19 of 19 (the earlier headroom-free model with a
/// 53pt threshold reproduced 15).
/// Value axis for data that may contain NEGATIVE values (both orientations).
///
/// DERIVED from the 8-arm `chart_negative` probe (Word render-truth, 2026-08-10):
///
///  * the axis always spans zero -- the modelled data range is
///    `[min(data_min,0), max(data_max,0)]`, each end given the same 5%
///    headroom the positive-only axis already uses (`AXIS_HEADROOM`).
///    N2 (all data negative, -19.2..-8.5) has axis_max exactly 0, and N4/N5
///    (scatter X starting at 1) have axis_min exactly 0.
///  * `axis_max = ceil(hi/step)*step`, `axis_min = floor(lo/step)*step`.
///  * `step` is the FINEST 1/2/5 x 10^k whose tick spacing
///    (`plot_len / div`) is at least `min_spacing` -- the same selection the
///    horizontal axis already made, now shared by both orientations.
///
/// The four measured negative arms are reproduced exactly:
///   N1 col   19.2/-8.5   span 29.085 -> step 5,  -10..25, 7 div (31.5pt)
///   N2 col   all-neg     span 20.16  -> step 5,  -25..0,  5 div (44.1pt)
///   N8 stack +29.7/-19.7 span 51.87  -> step 10, -30..40, 7 div (36.6pt)
///   N7 bar   19.2/-8.5   span 29.085 -> step 10, -10..30, 4 div (88.2pt)
///   N5 sc-X  2/-2        span 4.2    -> step 2,  -4..4,   4 div (74.4pt)
/// and every positive-only probe is bit-identical because `data_min >= 0`
/// forces `axis_min = 0`, which reduces this to the previous formula.
fn nice_axis_range(
    data_min: f64,
    data_max: f64,
    plot_len: f64,
    min_spacing: f64,
) -> (f64, f64, usize) {
    let lo = data_min.min(0.0) * AXIS_HEADROOM;
    let hi = data_max.max(0.0) * AXIS_HEADROOM;
    if plot_len <= 0.0 || !(hi - lo).is_finite() || hi - lo <= 0.0 {
        return (0.0, 1.0, 1);
    }
    let mut best: Option<(f64, f64, f64, usize)> = None; // (step, amin, amax, div)
    for k in -6i32..=9 {
        let mag = 10f64.powi(k);
        for m in [1.0f64, 2.0, 5.0] {
            let step = m * mag;
            let amax = (hi / step - 1e-9).ceil() * step;
            let amin = (lo / step + 1e-9).floor() * step;
            if !amax.is_finite() || !amin.is_finite() || amax - amin <= 0.0 {
                continue;
            }
            let div = ((amax - amin) / step).round() as usize;
            if div == 0 || div > 200 {
                continue;
            }
            if plot_len / div as f64 >= min_spacing {
                match best {
                    Some((bs, _, _, _)) if bs <= step => {}
                    _ => best = Some((step, amin, amax, div)),
                }
            }
        }
    }
    match best {
        Some((_, amin, amax, div)) => (amin, amax, div),
        // Plot too small for any nice step: fall back to a single division.
        None => {
            let amax = if data_max <= 0.0 {
                0.0
            } else {
                nice_axis_max(data_max)
            };
            let amin = if data_min >= 0.0 {
                0.0
            } else {
                -nice_axis_max(-data_min)
            };
            (amin, amax, 1)
        }
    }
}

fn horiz_value_axis(max_val: f64, plot_w: f64) -> (f64, usize) {
    if max_val <= 0.0 || plot_w <= 0.0 {
        return (1.0, 1);
    }
    let (_, axis_max, div) =
        nice_axis_range(0.0, max_val, plot_w, HORIZ_MIN_SPACING);
    (axis_max, div)
}

/// VERTICAL value-axis maximum.
///
/// Word gives the value axis 5% headroom above the data, exactly as the
/// horizontal axis does (see `horiz_value_axis`).  The separating specimen is
/// the stacked AREA probe chart_area S4/S5: category sums peak at 34.3, and
/// Word's axis is 40 -- `ceil(34.3/5)*5 = 35` without the headroom, but
/// `34.3*1.05 = 36.015` pushes the step from 5 to 10 and lands on 40, which the
/// render-truth confirms (the red band's top at x=150pt sits at y=140.9, i.e.
/// 30.884/40 of the plot height; the headroom-free 35 put it at 115.3).
/// Every previously measured vertical-axis probe is insensitive to it
/// (chart1 21.4 -> 25, chart_stacked 36.4 -> 40, chart_line 22.0 -> 25 with and
/// without), so this only ever changes charts that were wrong.
const AXIS_HEADROOM: f64 = 1.05;

/// Minimum tick spacing for a VERTICAL value axis, in points.
///
/// DERIVED as a window from the negative-value probe plus the whole recorded
/// battery: the step that Word CHOSE and the next finer one it REJECTED bracket
/// it from both sides.
///   rejected: chart_stacked step 2 (11.6pt), N1 step 2 (13.8), N2 step 2
///             (20.1), N8 step 5 (21.3), chart_scatter Y step 2 (21.3)
///   accepted: chart_stacked step 5 (29.0), N1 step 5 (31.5), chart1 step 5
///             (39.35), N2 step 5 (44.1)
/// so the window is (21.3, 29.0]; 25.0 sits mid-window.  Every positive-only
/// probe picks the same step it did under the old fixed 5-division rule.
/// Outline for an INVERTED (negative) bar.
///
/// `invertIfNegative` defaults to TRUE, so Word paints a negative bar white and
/// keeps a thin dark edge -- without the edge the bar would be invisible
/// (chart_negative N1/N2/N7 render-truth, 2026-08-10).
unsafe fn draw_neg_bar_outline(
    mem_dc: windows::Win32::Graphics::Gdi::HDC,
    r: &windows::Win32::Foundation::RECT,
    scale: f64,
) {
    // The GDI imports in this file are function-local, so a top-level helper
    // needs its own `use` (same as the LINE marker helper).
    use windows::Win32::Foundation::COLORREF;
    use windows::Win32::Graphics::Gdi::{
        CreatePen, DeleteObject, GetStockObject, Rectangle, SelectObject,
        NULL_BRUSH, PS_SOLID,
    };
    let pen = CreatePen(
        PS_SOLID,
        (0.75 * scale).round().max(1.0) as i32,
        COLORREF(colorref(0x3b, 0x3b, 0x3b)),
    );
    let old_pen = SelectObject(mem_dc, pen);
    let nb = GetStockObject(NULL_BRUSH);
    let ob2 = SelectObject(mem_dc, nb);
    let _ = Rectangle(mem_dc, r.left, r.top, r.right, r.bottom);
    SelectObject(mem_dc, ob2);
    SelectObject(mem_dc, old_pen);
    let _ = DeleteObject(pen);
}

/// Filled circle for a BUBBLE chart datum.  The GDI imports in this file are
/// function-local, so a top-level helper carries its own `use` (the same reason
/// the negative-bar outline and LINE marker helpers do).
unsafe fn draw_bubble_circle(
    mem_dc: windows::Win32::Graphics::Gdi::HDC,
    cx: f64,
    cy: f64,
    r: f64,
    scale: f64,
) {
    use windows::Win32::Graphics::Gdi::Ellipse;
    let _ = Ellipse(
        mem_dc,
        ((cx - r) * scale).round() as i32,
        ((cy - r) * scale).round() as i32,
        ((cx + r) * scale).round() as i32,
        ((cy + r) * scale).round() as i32,
    );
}

const VERT_MIN_SPACING: f64 = 25.0;

/// Minimum tick spacing for a HORIZONTAL value axis, in points (2026-08-09
/// 14-arm sweep: the window closed to (56.28, 57.76]).
const HORIZ_MIN_SPACING: f64 = 57.0;

/// Clear space Word keeps between two tick labels on a bubble/scatter numeric
/// X axis, in points: a step is usable when its tick spacing is at least
/// `label_width + BUBBLE_LABEL_GAP`.
///
/// Derived 2026-08-10 from every measured bubble/scatter X axis (24 sweep arms
/// + the 8-arm probe).  Neither a constant spacing nor a multiple of the label
/// width can separate them:
///   accepted 61.50/9.13  60.75/9.13  110.75/9.13  60.35/14.63  90.53/14.63
///            68.02/22.76 (spacing/label width)
///   rejected 30.75/9.13  51.10/14.63 63.29/22.76  48.43/22.76
/// -- 63.29 is rejected while 60.35 is accepted (kills a constant), and 3.49x
/// is rejected while 2.99x is accepted (kills a ratio).  "width + gap" fits all
/// ten with gap in (40.5, 45.3]; 43.0 is the midpoint.
///
/// The horizontal-BAR threshold is NOT this rule (bar accepts 48.34 with an
/// 18.26pt label = gap 30.1, and its own sweep contradicts itself across
/// chart_bar_axis pg13 / chart_bar_resid p4), so HORIZ_MIN_SPACING is left
/// alone and this applies to the bubble branch only.
const BUBBLE_LABEL_GAP: f64 = 43.0;

fn nice_axis_max(max_val: f64) -> f64 {
    if max_val <= 0.0 {
        return 1.0;
    }
    let max_val = max_val * AXIS_HEADROOM;
    let n_ticks = 5.0;
    let raw = max_val / n_ticks;
    let mag = 10f64.powf(raw.log10().floor());
    let resid = raw / mag;
    let step = if resid < 1.5 {
        1.0
    } else if resid < 3.0 {
        2.0
    } else if resid < 7.0 {
        5.0
    } else {
        10.0
    } * mag;
    (max_val / step).ceil() * step
}

/// `nice_axis_max` plus the DIVISION COUNT Word draws.
///
/// `nice_axis_max` picks a step internally and then throws it away; the number
/// of major divisions Word renders is `axis_max / step`, which is NOT always 5.
/// Word render-truth (chart_stock K1/K3/K4/K5, 2026-08-10): max 27.2 -> 28.56
/// after the 5% headroom -> raw 5.712 -> step 5 -> axis 30 -> SIX divisions,
/// and the 400pt-tall K5 arm draws the same six, so the count is a property of
/// the DATA, not of the plot height.
fn nice_axis_max_div(max_val: f64) -> (f64, usize) {
    let axis_max = nice_axis_max(max_val);
    if max_val <= 0.0 {
        return (axis_max, 5);
    }
    let padded = max_val * AXIS_HEADROOM;
    let raw = padded / 5.0;
    let mag = 10f64.powf(raw.log10().floor());
    let resid = raw / mag;
    let step = if resid < 1.5 {
        1.0
    } else if resid < 3.0 {
        2.0
    } else if resid < 7.0 {
        5.0
    } else {
        10.0
    } * mag;
    let div = (axis_max / step).round().max(1.0) as usize;
    (axis_max, div)
}

/// Paint a shape's `a:gradFill`, clipped to its outline.
///
/// The ramp model is the one already derived for slide backgrounds (linear
/// angle / scaled, or a circular focus running to the farthest corner); only
/// the area differs. The corpus has 302 slide-level gradient shapes on 35
/// slides in 4 decks plus 60 on layout shapes, and d24's title slide is built
/// entirely from them -- without this it renders as one flat slab.
#[cfg(windows)]
unsafe fn paint_shape_gradient(
    dc: windows::Win32::Graphics::Gdi::HDC,
    sh: &Shape,
    g: &SlideGradient,
    scale: f64,
) -> bool {
    use windows::Win32::Graphics::Gdi::*;

    if !shapegrad_on() {
        return false;
    }
    let w = (sh.width as f64 * scale).round() as i32;
    let h = (sh.height as f64 * scale).round() as i32;
    if w <= 0 || h <= 0 {
        return false;
    }
    // Clip first (the region is captured in device space), then shift the
    // origin so the background painter's 0..w,0..h maps onto the shape box.
    let clipped = clip_to_geometry_gdi(dc, sh, scale);
    let mut clip_rgn = None;
    if !clipped {
        let x = (sh.x as f64 * scale).round() as i32;
        let y = (sh.y as f64 * scale).round() as i32;
        let rgn = CreateRectRgn(x, y, x + w, y + h);
        let _ = SelectClipRgn(dc, rgn);
        clip_rgn = Some(rgn);
    }
    let mut old_org = windows::Win32::Foundation::POINT::default();
    let _ = SetViewportOrgEx(
        dc,
        (sh.x as f64 * scale).round() as i32,
        (sh.y as f64 * scale).round() as i32,
        Some(&mut old_org),
    );
    let turned = gradient_turned_with_shape(sh, g);
    paint_bg_gradient(dc, w, h, turned.as_ref().unwrap_or(g));
    let _ = SetViewportOrgEx(dc, old_org.x, old_org.y, None);
    let _ = SelectClipRgn(dc, None);
    if let Some(rgn) = clip_rgn {
        let _ = DeleteObject(rgn);
    }
    true
}

/// Turn a shape's linear ramp by the shape's own `a:xfrm`.
///
/// S-GRADROT (2026-08-24). `paint_shape_gradient` ran the ramp at the declared
/// `a:lin@ang` and ignored `rot` / `flipH` / `flipV` entirely, so d06's layout
/// wash -- `ang="5400012"` on a band at `rot="10800000"` -- came out mirrored
/// top-for-bottom: same amplitude, exactly reversed, on 15 of the deck's 39
/// slides. Both reference renderers beat Oxi on three of them.
///
/// The `gradrot` probe (38 arms, PowerPoint's own PDF, a black->white ramp fit
/// by least squares over each shape) measured the whole composition:
///
/// * `ang` maps 1:1 to the screen direction (0 = brightening rightward,
///   growing clockwise) -- 8 arms, every one within 0.02 degrees.
/// * the shape's `rot` is ADDED -- 12 arms.
/// * `rotWithShape="0"` pins the ramp to the page and drops the `rot` term;
///   ABSENT behaves as `"1"` -- 6 arms each. (Every gradFill in the dev corpus
///   omits the attribute, so the default is the load-bearing half.)
/// * a flip MIRRORS the ramp axis: `flipH` gives `180 - ang`, `flipV` gives
///   `-ang`, both give `180 + ang` -- 6 arms.
/// * the two compose in the DrawingML xfrm order, FLIP FIRST then rotate.
///   Every corpus shape is at `rot=180`, where the two orders agree, so this
///   was the one part the corpus could not settle: block E was authored with
///   both predictions written down first and PowerPoint answered
///   flip-then-rotate 4 times out of 4 (225 / 45 / 45 / 315 against the other
///   reading's 45 / 225 / 225 / 135).
///
/// A radial ramp is left alone: its focus would have to move with the shape,
/// and the corpus has no rotated one to measure against.
fn gradient_turned_with_shape(sh: &Shape, g: &SlideGradient) -> Option<SlideGradient> {
    if !gradrot_on() || g.focus.is_some() {
        return None;
    }
    let rot = if g.rot_with_shape { sh.rotation as f64 } else { 0.0 };
    if rot == 0.0 && !sh.flip_h && !sh.flip_v {
        return None;
    }
    let ang = g.angle_deg.unwrap_or(0.0) as f64;
    let flipped = match (sh.flip_h, sh.flip_v) {
        (true, false) => 180.0 - ang,
        (false, true) => -ang,
        (true, true) => 180.0 + ang,
        (false, false) => ang,
    };
    Some(SlideGradient {
        angle_deg: Some((flipped + rot).rem_euclid(360.0) as f32),
        ..g.clone()
    })
}

/// A shape with BOTH a gradient fill and an outline keeps its ramp unless this
/// is set (which restores losing the fill to the stroke).
fn gradstroke_on() -> bool {
    std::env::var("OXI_GRADSTROKE_DISABLE").is_err()
}

/// A preset shape hands a gradient-only fill to the ramp painter unless this
/// is set.
fn presetgrad_on() -> bool {
    std::env::var("OXI_PRESETGRAD_DISABLE").is_err()
}

/// Turn a shape's TEXT with the shape, for the duration of its text pass.
///
/// S-TEXTROT (2026-08-24). `create_font_for_wiu` passes 0 for `CreateFontW`'s
/// escapement and orientation, so a turned shape's text was always drawn across
/// the page. d35 slide 34's competitor matrix runs "LOW VALUE 1" / "HIGH VALUE
/// 1" down each side at -90 degrees, and both reference renderers beat Oxi on
/// that slide. **174 text shapes over 17 of the 40 dev decks carry a rotation**,
/// 133 of them exactly +/-90 (d39 56, d40 50, d30 10, d36 9, d19 8, d20 7 ...).
/// No corpus shape uses `bodyPr@vert` -- this is `a:xfrm@rot` alone.
///
/// The model, from the `textrot` probe (17 arms against PowerPoint's own PDF):
///
/// * the baseline runs at the shape's `rot`, in the same sense -- read off the
///   PDF's own span direction at 0 / 90 / 180 / 270, exact.
/// * the text turns about the UNTURNED box centre. The probe draws a hairline
///   frame on the same `a:xfrm`, and its ink centre is (360.00, 288.00) at every
///   angle; rotating each arm's local baseline origin about that point predicts
///   all 13 measured origins to within 0.13pt.
/// * ★the text is laid out in the shape's own box FIRST and turned afterwards,
///   NOT laid out in the turned bounding box. A line too long for the 288pt box
///   breaks into 2 lines at rot 0 and into 2 lines at rot +/-90 as well -- had
///   it been wrapping against the turned 72pt width it would have taken many
///   more. Anchor and alignment likewise stay in the shape's own frame
///   (`ctr`/`ctr` and `b`/`r` arms both predicted exactly).
///
/// Off-axis angles could not be read from the text layer at all: PowerPoint
/// exports text that is not axis-aligned as vector OUTLINES, so `get_text`
/// returns nothing for 30 / 45 / 135 / -45. The probe answers those with ink
/// instead -- turn the rot=0 arm's black pixels about the box centre and the
/// predicted bounding box must be the one PowerPoint drew. Max error over the
/// eight turned arms: **0.53pt at 150dpi**, i.e. one device pixel.
///
/// So the layout needs no change whatsoever, only the paint. A world transform
/// carries everything the pass draws -- glyphs, highlight boxes, bullet markers,
/// underline rules -- where font escapement would have turned only the baseline
/// of one `TextOut` and left the boxes square.
#[cfg(windows)]
unsafe fn begin_turned_text(
    dc: windows::Win32::Graphics::Gdi::HDC,
    sh: &Shape,
    paragraphs: &[oxislides_core::ir::SlideParagraph],
    scale: f64,
) -> Option<(i32, windows::Win32::Graphics::Gdi::XFORM)> {
    use windows::Win32::Graphics::Gdi::*;

    if !textrot_on() || sh.rotation == 0.0 || !sh.rotation.is_finite() {
        return None;
    }
    // ★Do not touch the device context for a shape with nothing to draw.
    // Installing the transform at all moves the shape's own edges by one
    // SUPERSAMPLED sub-pixel -- d06 slide 31's canvas band is an AutoShape at
    // rot=180 carrying ZERO paragraphs, and merely bracketing its (empty) text
    // pass shifted the band's bottom rule and four cell rules by a third of a
    // final pixel, for -0.0041. Bisected with a scaffold: setting GM_ADVANCED
    // and restoring it is inert, so it is the transform itself, and a 180-degree
    // turn maps the box onto itself, leaving only the rounding. Nothing is drawn
    // for an empty paragraph list, so the guard costs nothing and the question
    // of what GDI re-rounds does not arise.
    if paragraphs.is_empty() {
        return None;
    }
    let (sn, cs) = (sh.rotation as f64).to_radians().sin_cos();
    let cx = (sh.x as f64 + sh.width as f64 / 2.0) * scale;
    let cy = (sh.y as f64 + sh.height as f64 / 2.0) * scale;
    // Screen y grows downward, so a clockwise turn is [[c,-s],[s,c]]; GDI reads
    // the matrix transposed (x' = x*eM11 + y*eM21 + eDx).
    let m = XFORM {
        eM11: cs as f32,
        eM12: sn as f32,
        eM21: -sn as f32,
        eM22: cs as f32,
        eDx: (cx - cs * cx + sn * cy) as f32,
        eDy: (cy - sn * cx - cs * cy) as f32,
    };
    let old_mode = SetGraphicsMode(dc, GM_ADVANCED);
    if old_mode == 0 {
        return None;
    }
    let mut old = XFORM::default();
    if !GetWorldTransform(dc, &mut old).as_bool() {
        SetGraphicsMode(dc, GRAPHICS_MODE(old_mode));
        return None;
    }
    if !SetWorldTransform(dc, &m).as_bool() {
        SetGraphicsMode(dc, GRAPHICS_MODE(old_mode));
        return None;
    }
    Some((old_mode, old))
}

/// Put back whatever `begin_turned_text` replaced.
#[cfg(windows)]
unsafe fn end_turned_text(
    dc: windows::Win32::Graphics::Gdi::HDC,
    saved: Option<(i32, windows::Win32::Graphics::Gdi::XFORM)>,
) {
    use windows::Win32::Graphics::Gdi::*;
    if let Some((mode, old)) = saved {
        let _ = SetWorldTransform(dc, &old);
        SetGraphicsMode(dc, GRAPHICS_MODE(mode));
    }
}

/// Position italic text by the italic face's own advances.
///
/// ★UNPARKED 2026-08-25 (was `OXI_ITALADV=1`, now opt-out). Kept the old note
/// below because it is the reasoning that held it back; what changed is that
/// BOTH of its blockers were removed by later work:
///
///   * d17 s4 (-0.0037) was never about italic at all -- S-HMTXSTYLE fixed it
///     (the plain-weight advance table was answering for bold text). It no
///     longer appears in the diff.
///   * d16 s5 went from +0.0058 to **+0.0767** once S-RUNALIGN made the line
///     centre on the width it is actually drawn at. The two fixes compound: the
///     right advances are only worth having if the alignment uses them too.
///
/// Re-measured on the four decks that carry italic:
///
///     d16 s5    0.8907 -> 0.9674   (+0.0767)
///     d35 s25   0.9529 -> 0.9646   (+0.0116)
///     d16 s17   0.9624 -> 0.9699   (+0.0076)
///     d15 s5    0.9192 -> 0.9079   (-0.0112)   <- see below
///
/// **3 up / 1 down, net +0.0847 = corpus +0.000096.**
///
/// ★The one loss is a REFERENCE artifact, not a regression. d15 s5's stored
/// truth PDF is a COLD export: PowerPoint uses a deck's embedded italic parts
/// only from the SECOND open of a session, and the first falls back to the
/// upright part with a synthetic slant. Measured on that very deck --
/// open #1 `Barlow-Bold` width 360.06, opens #2 AND #3 `Barlow-BoldItalic`
/// width 347.85. Oxi now matches the WARM answer, which is the reproducible
/// one; the cold state is the one-off. See `pptx-truth-pdf-first-open-is-cold`.
/// If the corpus references are ever re-exported warm, this slide flips to a
/// gain.
///
/// The original parking note follows.
///
/// PARKED OPT-IN (`OXI_ITALADV=1`), because the rule was INCOMPLETE, not wrong.
/// It is exactly right on d16 -- the line widths become PowerPoint's to the
/// pixel (1026/1021/910/483 against 1026/1020/910/483, from 1057/1024/940/496)
/// and s5 gains +0.0058, s17 +0.0076. But d15 s5 LOSES 0.0112, and the reason
/// is that PowerPoint itself does not always use the embedded italic part:
///
///   d16  level-italic, regular run -> `SourceSansPro-Italic`      (embedded)
///   d16  level-italic, bold run    -> `SourceSansPro-BlackItalic` (embedded)
///   d15  level-italic, bold        -> `Barlow-Bold`, SKEWED       (upright!)
///
/// Both decks embed all four parts of the family in question, and d15's
/// `Barlow #BI` registers fine here (tmWeight 700, 471.38pt against the upright
/// bold's 487.12pt) -- PowerPoint simply declined it and synthesised the slant,
/// keeping the upright advances. Until the rule that DECIDES between those two
/// is measured, shipping this would be stacking an exception on a spec that is
/// not yet derived. Corpus with it on: **+0.000011, 2 decks up / 2 down**.
fn italadv_on() -> bool {
    std::env::var("OXI_ITALADV_DISABLE").is_err()
}

/// A shape's text is turned with the shape unless this is set.
fn textrot_on() -> bool {
    std::env::var("OXI_TEXTROT_DISABLE").is_err()
}

/// A shape gradient rides its shape's rotation unless this is set.
fn gradrot_on() -> bool {
    std::env::var("OXI_GRADROT_DISABLE").is_err()
}

/// Shape gradients are painted unless this is set.
fn shapegrad_on() -> bool {
    std::env::var("OXI_SHAPEGRAD_DISABLE").is_err()
}

/// Fonts PowerPoint could not resolve are drawn in Calibri unless this is set.
fn fontsub_on() -> bool {
    std::env::var("OXI_FONTSUB_DISABLE").is_err()
}

#[cfg(windows)]
thread_local! {
    /// requested family -> the family to actually draw with.
    static FAMILY_CACHE: std::cell::RefCell<std::collections::HashMap<String, String>> =
        std::cell::RefCell::new(std::collections::HashMap::new());
}

/// The family to draw `requested` with.
///
/// PowerPoint render-truth (`fontfallback` probe, 2026-08-18): a face it
/// cannot resolve is rendered in **Calibri** -- Mali, Jua, "Zzyzx
/// Nonexistent", Fira Sans and Lobster all came back as Calibri in
/// PowerPoint's own PDF, while Nunito (which IS present here) stayed Nunito.
/// It is NOT the theme font: d19 asks for Mali, its theme is Arial, and
/// PowerPoint drew Calibri.
///
/// Oxi previously handed the name to GDI's font mapper, which picks by PANOSE
/// and gave a face 15% taller in the caps (33.1pt against PowerPoint's 37.9pt
/// on d19 slide 1 at 52pt). Asking GDI what it actually selected is the
/// portable way to detect the substitution: a resolved family answers with its
/// own name.
#[cfg(windows)]
fn effective_family(dc: windows::Win32::Graphics::Gdi::HDC, requested: &str) -> String {
    use windows::Win32::Graphics::Gdi::*;

    if !fontsub_on() || requested.is_empty() {
        return requested.to_string();
    }
    FAMILY_CACHE.with(|cache| {
        if let Some(hit) = cache.borrow().get(requested) {
            return hit.clone();
        }
        let probe = probe_dc();
        let wide: Vec<u16> = requested.encode_utf16().chain(std::iter::once(0)).collect();
        let resolved = unsafe {
            let font = CreateFontW(
                -64,
                0,
                0,
                0,
                400,
                0,
                0,
                0,
                DEFAULT_CHARSET.0 as u32,
                OUT_DEFAULT_PRECIS.0 as u32,
                CLIP_DEFAULT_PRECIS.0 as u32,
                CLEARTYPE_QUALITY.0 as u32,
                (DEFAULT_PITCH.0 | FF_DONTCARE.0) as u32,
                windows::core::PCWSTR(wide.as_ptr()),
            );
            if font.is_invalid() {
                requested.to_string()
            } else {
                let old = SelectObject(probe, font);
                let mut name = [0u16; 64];
                let n = GetTextFaceW(probe, Some(&mut name));
                SelectObject(probe, old);
                let _ = DeleteObject(font);
                let got = if n > 0 {
                    String::from_utf16_lossy(&name[..(n as usize).saturating_sub(1)])
                } else {
                    String::new()
                };
                if got.eq_ignore_ascii_case(requested) {
                    requested.to_string()
                } else {
                    "Calibri".to_string()
                }
            }
        };
        let _ = dc;
        cache
            .borrow_mut()
            .insert(requested.to_string(), resolved.clone());
        resolved
    })
}

/// Draw one wrapped line as its RUN segments, each in its own style.
///
/// A paragraph's runs can differ in bold, colour and size -- 75 / 73 / 38
/// paragraphs in the dev corpus do, on 70 slides across 17 decks -- and the
/// pre-S-RUNSTYLE path drew the whole line in the FIRST run's colour at one
/// weight, so "**Lead-in:** rest" came out uniformly styled. (No corpus
/// paragraph mixes font FAMILIES, so the wrap still measures the line with one
/// family; only the drawing is split.)
///
/// `line_start` is the line's offset in the paragraph's concatenated text,
/// which is recoverable because the wrap partitions that text in order without
/// dropping characters.
#[cfg(windows)]
#[allow(clippy::too_many_arguments)]
unsafe fn draw_line_runs(
    dc: windows::Win32::Graphics::Gdi::HDC,
    x: i32,
    baseline: f32,
    line_text: &str,
    line_start: usize,
    runs: &[oxislides_core::ir::SlideRun],
    default_family: &str,
    default_fs: f32,
    default_color: Option<&str>,
    // The level's own `a:highlight`, used by runs that declare none.
    default_highlight: Option<&str>,
    // The level's `a:defRPr/@i`, which every run inherits.
    default_italic: bool,
    // The level's `a:defRPr/@b`, likewise. d11's master title placeholder
    // declares it and every title in the deck is bold.
    default_bold: bool,
    scale: f64,
) {
    // Walk the runs, clipping each to this line's character range.
    let line_chars: Vec<char> = line_text.chars().collect();
    let line_end = line_start + line_chars.len();

    // The highlight box is the LINE's, not the run's: a 48pt neighbour made
    // an 18pt highlighted run's box 53.64pt tall, exactly the 48pt face's
    // ascent plus descent (`highlight` probe, 2026-08-19). So the tallest run
    // on this line sets both edges.
    let mut box_up = 0.0f32;
    let mut box_down = 0.0f32;
    if highlight_on()
        && (default_highlight.is_some() || runs.iter().any(|r| r.highlight.is_some()))
    {
        let mut at = 0usize;
        for run in runs {
            let n = run.text.chars().count();
            let (rs, re) = (at, at + n);
            at = re;
            if rs.max(line_start) >= re.min(line_end) {
                continue;
            }
            let fs = run.font_size.unwrap_or(default_fs);
            let family = effective_family(dc, run.font_family.as_deref().unwrap_or(default_family));
            let (up, down) = highlight_extent_pt(&family, fs);
            if up + down > box_up + box_down {
                box_up = up;
                box_down = down;
            }
        }
    }

    let mut cursor_x = x;
    let mut at = 0usize; // char offset of the current run's start
    for run in runs {
        let n = run.text.chars().count();
        let (rs, re) = (at, at + n);
        at = re;
        let from = rs.max(line_start);
        let to = re.min(line_end);
        if from >= to {
            continue;
        }
        let seg: String = line_chars[from - line_start..to - line_start].iter().collect();
        if seg.is_empty() {
            continue;
        }
        let fs = run.font_size.unwrap_or(default_fs);
        let family = &effective_family(dc, run.font_family.as_deref().unwrap_or(default_family));
        let color = run.color.as_deref().or(default_color);
        let bold = run.bold || default_bold;
        let weight = if bold { 700 } else { 400 };
        let w = runtime_width_px(dc, &seg, fs, family, bold, run.italic, scale)
            .or_else(|| font_adv::text_hmtx_px(&seg, fs, family, scale))
            .unwrap_or_else(|| {
                measure_text_width(dc, &seg, fs, family, bold, scale).round() as i32
            });
        // Behind the glyphs, and exactly as wide as this run's advance -- the
        // probe's `HIGH ` arm put the box's right edge on the trailing space's
        // right edge, not on the last letter's.
        if highlight_on() {
            if let Some(hex) = run
                .highlight
                .as_deref()
                .or(default_highlight)
                .filter(|_| box_up + box_down > 0.0)
            {
                if let Some((r, g, b)) = parse_hex_rgb(hex) {
                    let base_px = (baseline as f64 * scale).round() as i32;
                    let rect = windows::Win32::Foundation::RECT {
                        left: cursor_x,
                        top: base_px - (box_up as f64 * scale).round() as i32,
                        right: cursor_x + w,
                        bottom: base_px + (box_down as f64 * scale).round() as i32,
                    };
                    let brush = windows::Win32::Graphics::Gdi::CreateSolidBrush(
                        windows::Win32::Foundation::COLORREF(colorref(r, g, b)),
                    );
                    if !brush.is_invalid() {
                        unsafe {
                            windows::Win32::Graphics::Gdi::FillRect(dc, &rect, brush);
                            let _ = windows::Win32::Graphics::Gdi::DeleteObject(brush);
                        }
                    }
                }
            }
        }
        draw_text_baseline_wiu(
            dc,
            cursor_x,
            baseline,
            &seg,
            fs,
            family,
            color,
            scale,
            weight,
            run.italic || default_italic,
            run.underline,
        );
        cursor_x += w;
    }
}

/// The layout / master placeholder `a:lstStyle` chain is applied unless this
/// is set.
fn phlevel_on() -> bool {
    std::env::var("OXI_PHLEVEL_DISABLE").is_err()
}

/// A bottom-anchored block taller than its box overflows upward rather than
/// being clamped to the box top, unless this is set.
fn anchorb_on() -> bool {
    std::env::var("OXI_ANCHORB_DISABLE").is_err()
}

/// Run-level styling within a paragraph is applied unless this is set.
fn runstyle_on() -> bool {
    std::env::var("OXI_RUNSTYLE_DISABLE").is_err()
}

/// EM size the runtime advance probe measures at. At 2048 units per em the
/// integer quantisation `GetCharABCWidthsW` still applies is 1/2048 of an em,
/// i.e. below 0.05% -- far under the 1.6% error GDI shows at render size.
#[cfg(windows)]
const ADVANCE_PROBE_EM: i32 = 2048;

/// GDI reports character advances as integers scaled to the probe size, so the
/// probe size IS the measurement resolution.
///
/// 2048 is exact for the common 2048-unit TrueType em -- Lobster's advances land
/// on whole units there -- and only quantises a 1000-unit CFF em, so probing
/// finer is opt-in until a corpus measurement says it is worth the cache churn:
/// on d09 it moved the deck +0.0001 with the per-slide signs mixed.
///
/// ★RE-MEASURED 2026-08-25 (trap #86: a parked flag must be re-checked when the
/// ground under it moves, and a great deal moved that day -- S-RUNALIGN,
/// S-HMTXSTYLE, S-ITALADV and S-SUBTITLE all changed how text is measured or
/// which face it resolves to). The answer did NOT change: over d09 / d13 / d05
/// it is **22 slides up / 27 down, net +0.0024**, still mixed signs for a
/// negligible net. Stays parked -- re-measuring is for finding out, not for
/// unparking.
#[cfg(windows)]
fn advance_probe_em() -> i32 {
    if std::env::var("OXI_ADVPREC_ENABLE").is_ok() {
        16384
    } else {
        ADVANCE_PROBE_EM
    }
}

#[cfg(windows)]
thread_local! {
    /// (family, weight, italic) -> char -> advance in EM units.
    static BASELINE_CACHE: std::cell::RefCell<
        std::collections::HashMap<String, Option<f32>>,
    > = std::cell::RefCell::new(std::collections::HashMap::new());
    /// (family, weight, italic) -> char -> ink reach past the advance, EM units.
    static OVERHANG_CACHE: std::cell::RefCell<
        std::collections::HashMap<(String, i32, bool), std::collections::HashMap<char, Option<f32>>>,
    > = std::cell::RefCell::new(std::collections::HashMap::new());
    /// The 16384-per-em advance probe backing the master-unit break test.
    /// Separate from ADVANCE_CACHE so render positions keep the shipped
    /// 2048-probe values byte-for-byte while the BREAK test gets advances
    /// precise enough to round to the correct 1/8pt bucket (the 2048 probe's
    /// +-1/4096 em error misplaces a display-size glyph's bucket routinely).
    static PRECISE_CACHE: std::cell::RefCell<
        std::collections::HashMap<(String, i32, bool), std::collections::HashMap<char, Option<f32>>>,
    > = std::cell::RefCell::new(std::collections::HashMap::new());
    /// (face, weight, italic) -> that face's own design advances.
    static FACE_ADV_CACHE: std::cell::RefCell<
        std::collections::HashMap<(String, i32, bool), Option<FaceAdvances>>,
    > = std::cell::RefCell::new(std::collections::HashMap::new());
    static ADVANCE_CACHE: std::cell::RefCell<
        std::collections::HashMap<(String, i32, bool), std::collections::HashMap<char, Option<f32>>>,
    > = std::cell::RefCell::new(std::collections::HashMap::new());
    /// A DC used ONLY for advance probing.
    ///
    /// ★Probing on the DC the renderer draws with made the output
    /// NON-DETERMINISTIC: the same binary rendering d19 twice differed on
    /// slide 39, an icon row where every glyph shifted by one position, while
    /// the opt-out arm was byte-stable across runs. Selecting a probe font
    /// into a DC mid-draw perturbs GDI's font-linking state for the glyphs
    /// that follow. A private DC keeps the two apart.
    static PROBE_DC: std::cell::Cell<isize> = const { std::cell::Cell::new(0) };
}

/// The dedicated probe DC, created on first use.
#[cfg(windows)]
fn probe_dc() -> windows::Win32::Graphics::Gdi::HDC {
    use windows::Win32::Foundation::HWND;
    use windows::Win32::Graphics::Gdi::*;
    PROBE_DC.with(|cell| {
        let raw = cell.get();
        if raw != 0 {
            return HDC(raw as *mut core::ffi::c_void);
        }
        unsafe {
            let screen = GetDC(HWND(std::ptr::null_mut()));
            let dc = CreateCompatibleDC(screen);
            let _ = ReleaseDC(HWND(std::ptr::null_mut()), screen);
            cell.set(dc.0 as isize);
            dc
        }
    })
}

/// How far a glyph's ink reaches past its advance, in EM units, never negative.
///
/// A script face joins its letters by drawing outside the advance box: Lobster's
/// `o` overhangs by 0.123 em. Whether the BREAK test counts any of that is the
/// open question `trailing_overhang_px` documents -- the full ABC-C term is
/// falsified, so this feeds only the opt-in probe and the OXI_ADV_DEBUG dump.
#[cfg(windows)]
fn runtime_overhang_em(family: &str, bold: bool, italic: bool, ch: char) -> Option<f32> {
    use windows::Win32::Graphics::Gdi::*;

    let dc = probe_dc();
    let weight = if bold { 700 } else { 400 };
    let key = (family.to_string(), weight, italic);
    let (face, weight, italic) = styled_face(family, bold, italic);
    OVERHANG_CACHE.with(|cache| {
        let mut cache = cache.borrow_mut();
        let per_font = cache.entry(key).or_default();
        if let Some(hit) = per_font.get(&ch) {
            return *hit;
        }
        let wide: Vec<u16> = face.encode_utf16().chain(std::iter::once(0)).collect();
        let value = unsafe {
            let probe_em = advance_probe_em();
            let font = CreateFontW(
                -probe_em,
                0,
                0,
                0,
                weight,
                u32::from(italic),
                0,
                0,
                DEFAULT_CHARSET.0 as u32,
                OUT_DEFAULT_PRECIS.0 as u32,
                CLIP_DEFAULT_PRECIS.0 as u32,
                CLEARTYPE_QUALITY.0 as u32,
                (DEFAULT_PITCH.0 | FF_DONTCARE.0) as u32,
                windows::core::PCWSTR(wide.as_ptr()),
            );
            if font.is_invalid() {
                None
            } else {
                let old = SelectObject(dc, font);
                let mut abc = ABC::default();
                let code = ch as u32;
                let ok = GetCharABCWidthsW(dc, code, code, &mut abc).as_bool();
                SelectObject(dc, old);
                let _ = DeleteObject(font);
                if ok {
                    Some((-abc.abcC).max(0) as f32 / probe_em as f32)
                } else {
                    None
                }
            }
        };
        per_font.insert(ch, value);
        if let Ok(want) = std::env::var("OXI_ADV_DEBUG") {
            if face.contains(&want) {
                eprintln!("OVH {face} '{ch}' {value:?}");
            }
        }
        value
    })
}

/// The design advance of `ch` in EM units, read from the font GDI actually
/// resolved for `family` (including a privately loaded embedded one).
///
/// `font_adv`'s tables already encode the rule -- PowerPoint places glyphs at
/// the TrueType design advance, not at GDI's hinted, pixel-snapped one -- but
/// they cover three hardcoded families. Every embedded font (262 parts in the
/// dev corpus) fell through to the hinted path, where a d28 body line measures
/// **+3.9pt (+1.6%)** wider than its design width; that is what pushes a word
/// onto the next line. Measured 2026-08-18 against PowerPoint's own PDF, whose
/// per-character pen positions sit within 0.06pt of the design advances.
#[cfg(windows)]
fn runtime_advance_em(family: &str, bold: bool, italic: bool, ch: char) -> Option<f32> {
    use windows::Win32::Graphics::Gdi::*;

    let dc = probe_dc();

    let weight = if bold { 700 } else { 400 };
    let key = (family.to_string(), weight, italic);
    // Measure the face that will be DRAWN: an embedded bold or italic part is
    // its own GDI family, and asking the base name for weight 700 would
    // measure a synthesised face instead.
    let (face, weight, italic) = styled_face(family, bold, italic);
    ADVANCE_CACHE.with(|cache| {
        let mut cache = cache.borrow_mut();
        let per_font = cache.entry(key).or_default();
        if let Some(hit) = per_font.get(&ch) {
            return *hit;
        }
        let wide: Vec<u16> = face.encode_utf16().chain(std::iter::once(0)).collect();
        let value = unsafe {
            let probe_em = advance_probe_em();
            let font = CreateFontW(
                -probe_em,
                0,
                0,
                0,
                weight,
                u32::from(italic),
                0,
                0,
                DEFAULT_CHARSET.0 as u32,
                OUT_DEFAULT_PRECIS.0 as u32,
                CLIP_DEFAULT_PRECIS.0 as u32,
                CLEARTYPE_QUALITY.0 as u32,
                (DEFAULT_PITCH.0 | FF_DONTCARE.0) as u32,
                windows::core::PCWSTR(wide.as_ptr()),
            );
            if font.is_invalid() {
                None
            } else {
                let old = SelectObject(dc, font);
                let mut abc = ABC::default();
                let code = ch as u32;
                let ok = GetCharABCWidthsW(dc, code, code, &mut abc).as_bool();
                SelectObject(dc, old);
                let _ = DeleteObject(font);
                if ok {
                    Some((abc.abcA + abc.abcB as i32 + abc.abcC) as f32 / probe_em as f32)
                } else {
                    None
                }
            }
        };
        per_font.insert(ch, value);
        if let Ok(want) = std::env::var("OXI_ADV_DEBUG") {
            if face.contains(&want) {
                eprintln!("ADV {face} '{ch}' {value:?}");
            }
        }
        value
    })
}

/// Enumerate every registered face whose name contains `OXI_ENUM_DEBUG`,
/// with the weight and slant GDI records for it. The deck's embedded parts are
/// private to this process, so nothing outside it can see this list.
#[cfg(windows)]
fn enum_faces_debug() {
    use windows::Win32::Graphics::Gdi::*;

    let Ok(want) = std::env::var("OXI_ENUM_DEBUG") else {
        return;
    };
    unsafe extern "system" fn cb(
        lf: *const LOGFONTW,
        tm: *const TEXTMETRICW,
        _kind: u32,
        want: windows::Win32::Foundation::LPARAM,
    ) -> i32 {
        let name = String::from_utf16_lossy(&(*lf).lfFaceName);
        let name = name.trim_end_matches(' ');
        let want = &*(want.0 as *const String);
        if name.to_lowercase().contains(&want.to_lowercase()) {
            eprintln!(
                "ENUM {name:26} lfWeight={:4} lfItalic={} tmWeight={:4} tmItalic={}",
                (*lf).lfWeight,
                (*lf).lfItalic,
                (*tm).tmWeight,
                (*tm).tmItalic
            );
        }
        1
    }
    let dc = probe_dc();
    let mut lf = LOGFONTW {
        lfCharSet: DEFAULT_CHARSET,
        ..Default::default()
    };
    unsafe {
        EnumFontFamiliesExW(
            dc,
            &mut lf,
            Some(cb),
            windows::Win32::Foundation::LPARAM(&want as *const String as isize),
            0,
        );
    }
}

/// The design advances of one face, keyed by code point, in font units.
#[cfg(windows)]
struct FaceAdvances {
    upem: f32,
    by_cp: std::collections::HashMap<u32, u16>,
}

#[cfg(windows)]
fn be16(b: &[u8], off: usize) -> Option<u16> {
    Some(u16::from_be_bytes([*b.get(off)?, *b.get(off + 1)?]))
}

#[cfg(windows)]
fn be32(b: &[u8], off: usize) -> Option<u32> {
    Some(u32::from_be_bytes([
        *b.get(off)?,
        *b.get(off + 1)?,
        *b.get(off + 2)?,
        *b.get(off + 3)?,
    ]))
}

/// One `sfnt` table of the font currently selected into `dc`.
///
/// `GetFontData` reaches the deck's OWN embedded parts: `TTLoadEmbeddedFont`
/// registers them privately in this process, and GDI hands their tables back
/// like any other face's. There is no file on disk to open for those, which is
/// why the design advance was previously out of reach for them.
#[cfg(windows)]
fn font_table(dc: windows::Win32::Graphics::Gdi::HDC, tag: &[u8; 4]) -> Option<Vec<u8>> {
    use windows::Win32::Graphics::Gdi::*;

    // GetFontData takes the tag with its FIRST byte in the LOW byte.
    let t = u32::from_le_bytes(*tag);
    unsafe {
        let n = GetFontData(dc, t, 0, None, 0);
        // GDI_ERROR, the failure return, is 0xFFFFFFFF.
        if n == u32::MAX || n == 0 {
            return None;
        }
        let mut buf = vec![0u8; n as usize];
        let got = GetFontData(dc, t, 0, Some(buf.as_mut_ptr().cast()), n);
        if got == u32::MAX {
            return None;
        }
        Some(buf)
    }
}

/// code point -> glyph id, from a `cmap` table (formats 4 and 12).
#[cfg(windows)]
fn cmap_by_code_point(cmap: &[u8]) -> Option<std::collections::HashMap<u32, u16>> {
    let n = be16(cmap, 2)? as usize;
    let mut best: Option<usize> = None;
    let mut best_score = -1i32;
    for i in 0..n {
        let rec = 4 + 8 * i;
        let score = match (be16(cmap, rec)?, be16(cmap, rec + 2)?) {
            (3, 10) => 4,
            (3, 1) => 3,
            (0, _) => 2,
            (3, 0) => 1,
            _ => 0,
        };
        if score > best_score {
            best_score = score;
            best = Some(be32(cmap, rec + 4)? as usize);
        }
    }
    let sub = cmap.get(best?..)?;
    let mut map = std::collections::HashMap::new();
    match be16(sub, 0)? {
        4 => {
            let segx2 = be16(sub, 6)? as usize;
            let ends = 14;
            let starts = ends + segx2 + 2;
            let deltas = starts + segx2;
            let ranges = deltas + segx2;
            for s in 0..segx2 / 2 {
                let end = be16(sub, ends + 2 * s)?;
                let start = be16(sub, starts + 2 * s)?;
                if start > end {
                    continue;
                }
                let delta = be16(sub, deltas + 2 * s)?;
                let ro = be16(sub, ranges + 2 * s)?;
                for cp in start..=end {
                    if cp == 0xFFFF {
                        continue;
                    }
                    let g = if ro == 0 {
                        cp.wrapping_add(delta)
                    } else {
                        let at = ranges + 2 * s + ro as usize + 2 * (cp - start) as usize;
                        match be16(sub, at) {
                            Some(0) | None => continue,
                            Some(g) => g.wrapping_add(delta),
                        }
                    };
                    if g != 0 {
                        map.insert(u32::from(cp), g);
                    }
                }
            }
        }
        12 => {
            let groups = be32(sub, 12)? as usize;
            for g in 0..groups.min(4096) {
                let rec = 16 + 12 * g;
                let (start, end, gid) = (be32(sub, rec)?, be32(sub, rec + 4)?, be32(sub, rec + 8)?);
                if end < start || end - start > 0xFFFF {
                    continue;
                }
                for cp in start..=end {
                    map.insert(cp, (gid + (cp - start)) as u16);
                }
            }
        }
        _ => return None,
    }
    Some(map)
}

/// Every code point's design advance for the font selected into `dc`.
#[cfg(windows)]
fn read_face_advances(dc: windows::Win32::Graphics::Gdi::HDC) -> Option<FaceAdvances> {
    let head = font_table(dc, b"head")?;
    let upem = f32::from(be16(&head, 18)?);
    if upem <= 0.0 {
        return None;
    }
    let hhea = font_table(dc, b"hhea")?;
    let num_h = be16(&hhea, 34)? as usize;
    if num_h == 0 {
        return None;
    }
    let hmtx = font_table(dc, b"hmtx")?;
    let cmap = font_table(dc, b"cmap")?;
    let mut by_cp = std::collections::HashMap::new();
    for (cp, gid) in cmap_by_code_point(&cmap)? {
        // Glyphs past the last full metric all carry that metric's advance.
        let g = (gid as usize).min(num_h - 1);
        if let Some(a) = be16(&hmtx, 4 * g) {
            by_cp.insert(cp, a);
        }
    }
    if by_cp.is_empty() {
        return None;
    }
    Some(FaceAdvances { upem, by_cp })
}

/// A glyph's DESIGN advance in EM units, read out of the font GDI actually
/// selected -- the answer to "which face is being served under this name?".
///
/// ★This is a DIAGNOSTIC (`OXI_FD_DEBUG`), not an advance source. It is here
/// because a wrong face is otherwise undetectable: `GetTextFaceW` echoes the
/// name that was ASKED for, `EnumFontFamiliesEx` cannot see a privately loaded
/// part at all, and a face name plus a weight number says nothing about the
/// advances the glyphs actually carry. Reading the face's own `hmtx` back is
/// the only thing that does.
///
/// What it found (2026-08-26, d15). The deck embeds a part under
/// `typeface="Barlow Light"` whose data is Barlow REGULAR: asked for "Barlow
/// Light" GDI returns `a`=0.511 `w`=0.721, byte-identical to what the same
/// deck's "Barlow" part returns, and not the `a`=0.506 `w`=0.705 a real
/// Barlow-Light carries. Registering it shadows the genuine Barlow Light in
/// the Office cloud cache (`CloudFonts/Barlow/25577919585.ttf`, usWeightClass
/// 300), which is the face PowerPoint's own export used -- its PDF subset says
/// Barlow-Light, 300, `a`=0.506. So d15 renders its whole body one weight too
/// heavy, +0.54% per line, and slide 2 loses the word "will" off the end of a
/// 273.47pt box.
///
/// The fix is NOT settled and must not be guessed at. Across the six dev decks
/// where an embedded part disagrees with the local font of the same name,
/// PowerPoint's own output follows the LOCAL font on d10 and d15 and the
/// EMBEDDED part on d35 -- and the corpus PDFs come from at least three export
/// dates, with `pptx-truth-pdf-first-open-is-cold` already on record. Which of
/// those two groups is a PowerPoint rule and which is an artefact of when the
/// reference was exported has to be settled before any precedence changes.
#[cfg(windows)]
fn fontdata_advance_em(family: &str, bold: bool, italic: bool, ch: char) -> Option<f32> {
    use windows::Win32::Graphics::Gdi::*;

    let (face, weight, italic) = styled_face(family, bold, italic);
    FACE_ADV_CACHE.with(|cache| {
        let mut cache = cache.borrow_mut();
        let entry = cache
            .entry((face.clone(), weight, italic))
            .or_insert_with(|| {
                let dc = probe_dc();
                let wide: Vec<u16> = face.encode_utf16().chain(std::iter::once(0)).collect();
                unsafe {
                    let font = CreateFontW(
                        -2048,
                        0,
                        0,
                        0,
                        weight,
                        u32::from(italic),
                        0,
                        0,
                        DEFAULT_CHARSET.0 as u32,
                        OUT_DEFAULT_PRECIS.0 as u32,
                        CLIP_DEFAULT_PRECIS.0 as u32,
                        CLEARTYPE_QUALITY.0 as u32,
                        (DEFAULT_PITCH.0 | FF_DONTCARE.0) as u32,
                        windows::core::PCWSTR(wide.as_ptr()),
                    );
                    if font.is_invalid() {
                        return None;
                    }
                    let old = SelectObject(dc, font);
                    let read = read_face_advances(dc);
                    let mut real = [0u16; 64];
                    let n = GetTextFaceW(dc, Some(&mut real));
                    SelectObject(dc, old);
                    let _ = DeleteObject(font);
                    if std::env::var("OXI_FD_DEBUG").is_ok() {
                        let real = String::from_utf16_lossy(&real[..(n as usize).saturating_sub(1)]);
                        match &read {
                            Some(f) => eprintln!(
                                "FD {face:22} w={weight} i={italic} -> {real:22} upem={} glyphs={} 'a'={:?} 'w'={:?}",
                                f.upem,
                                f.by_cp.len(),
                                f.by_cp.get(&(97u32)).map(|a| f32::from(*a) / f.upem),
                                f.by_cp.get(&(119u32)).map(|a| f32::from(*a) / f.upem),
                            ),
                            None => eprintln!("FD {face:22} w={weight} i={italic} -> {real:22} NO TABLES"),
                        }
                    }
                    read
                }
            });
        entry
            .as_ref()
            .and_then(|f| f.by_cp.get(&(ch as u32)).map(|a| f32::from(*a) / f.upem))
    })
}

/// A glyph advance in EM units from a 16384-per-em GDI probe -- the
/// break-test twin of `runtime_advance_em`, eight times finer. 16384 exceeds
/// every common unitsPerEm (1000 CFF, 2000, 2048), so the integer GDI width
/// recovers the design advance exactly instead of to the render probe's
/// 1/2048.
#[cfg(windows)]
fn precise_advance_em(family: &str, bold: bool, italic: bool, ch: char) -> Option<f32> {
    use windows::Win32::Graphics::Gdi::*;

    const PRECISE_PROBE_EM: i32 = 16384;
    let dc = probe_dc();
    let weight = if bold { 700 } else { 400 };
    let key = (family.to_string(), weight, italic);
    let (face, weight, italic) = styled_face(family, bold, italic);
    PRECISE_CACHE.with(|cache| {
        let mut cache = cache.borrow_mut();
        let per_font = cache.entry(key).or_default();
        if let Some(hit) = per_font.get(&ch) {
            return *hit;
        }
        let wide: Vec<u16> = face.encode_utf16().chain(std::iter::once(0)).collect();
        let value = unsafe {
            let font = CreateFontW(
                -PRECISE_PROBE_EM,
                0,
                0,
                0,
                weight,
                u32::from(italic),
                0,
                0,
                DEFAULT_CHARSET.0 as u32,
                OUT_DEFAULT_PRECIS.0 as u32,
                CLIP_DEFAULT_PRECIS.0 as u32,
                CLEARTYPE_QUALITY.0 as u32,
                (DEFAULT_PITCH.0 | FF_DONTCARE.0) as u32,
                windows::core::PCWSTR(wide.as_ptr()),
            );
            if font.is_invalid() {
                None
            } else {
                let old = SelectObject(dc, font);
                let mut abc = ABC::default();
                let code = ch as u32;
                let ok = GetCharABCWidthsW(dc, code, code, &mut abc).as_bool();
                SelectObject(dc, old);
                let _ = DeleteObject(font);
                if ok {
                    Some(
                        (abc.abcA + abc.abcB as i32 + abc.abcC) as f32
                            / PRECISE_PROBE_EM as f32,
                    )
                } else {
                    None
                }
            }
        };
        per_font.insert(ch, value);
        if let Ok(want) = std::env::var("OXI_ADV_DEBUG") {
            if face.contains(&want) {
                eprintln!("PRC {face} '{ch}' {value:?}");
            }
        }
        value
    })
}

/// PowerPoint's break test measures in master units unless this is set.
fn masterunit_on() -> bool {
    std::env::var("OXI_MASTERUNIT_DISABLE").is_err()
}

/// The room a candidate line claims, in MASTER UNITS (1/8 pt -- the
/// PowerPoint-97 1/576-inch unit, still governing text measurement).
///
/// DERIVED 2026-08-21 (wrapfit COM sweeps, rounds 1-3): each glyph's design
/// advance at the font size is rounded to the nearest master unit and the
/// line fits iff the SUM, exact in master units, is <= the effective box
/// width (inclusive; the width side is EMU-precise, not quantised). The
/// brackets exclude round-of-sum, floor, ceil, every px-grid quantum, and
/// any trailing-ink term (Segoe Script 't', ink +0.22 em past its advance,
/// thresholds exactly like '.'). The rule reproduces d09 s13 "Happy Holi!":
/// master sum 546.5pt over a 546.4128pt box breaks, while the float sum
/// 546.399pt would fit.
///
/// None when any glyph's advance is unknown -- the caller falls back to the
/// legacy pixel test. That includes any non-BMP character
/// (`GetCharABCWidthsW` is a UCS-2 API; asked about U+1F60A it measures some
/// other char's glyph and poisons the sum -- the first mu1 gate run lost
/// -0.08 on each of the d11/d24/d35 emoji charwrap slides this way) and any
/// text the family lacks a glyph for, where GDI would silently measure a
/// font-link fallback with the base font's metrics (the d19 s39 icon-row
/// non-determinism lesson -- see runtime_dx_px).
#[cfg(windows)]
fn master_units(text: &str, fs: f32, family: &str, bold: bool, italic: bool) -> Option<i64> {
    if text.chars().any(|c| c as u32 > 0xFFFF) {
        return None;
    }
    if !font_has_all_glyphs(family, bold, italic, text) {
        return None;
    }
    let mut sum: i64 = 0;
    for ch in text.chars() {
        let em = font_adv::hmtx_advance_em(family, ch)
            .or_else(|| precise_advance_em(family, bold, italic, ch))?;
        sum += f64::from(em * fs * 8.0).round() as i64;
    }
    Some(sum)
}

/// The styles a paragraph's runs impose on a candidate line, so the break test
/// can measure each character at the size and weight it is DRAWN at.
///
/// `draw_line_runs` already walks runs by character offset in the paragraph's
/// concatenated text; this is the same walk on the measuring side. 112 of the
/// corpus's 6366 non-empty paragraphs switch style mid-paragraph (99 change
/// weight, 57 size, 11 slant; none change family), and measuring those as one
/// style breaks them in the wrong place.
#[cfg(windows)]
#[derive(Clone, Copy)]
struct RunStyles<'a> {
    runs: &'a [oxislides_core::ir::SlideRun],
    /// Characters of the paragraph already committed to earlier lines.
    line_start: usize,
}

/// Per-run master units are used when the paragraph has run styles unless this
/// is set, which restores measuring the whole line in one style.
fn runmeasure_on() -> bool {
    std::env::var("OXI_RUNMEASURE_DISABLE").is_err()
}

/// Master units for `text`, each character measured at its own run's size and
/// weight. None when any character's advance is unknown.
#[cfg(windows)]
fn master_units_runs(
    text: &str,
    fs: f32,
    family: &str,
    bold: bool,
    italic: bool,
    styles: RunStyles<'_>,
) -> Option<i64> {
    if text.chars().any(|c| c as u32 > 0xFFFF) {
        return None;
    }
    // Coverage is asked ONCE for the line, as the single-style path does: the
    // per-character form issued a GDI font creation per character per candidate
    // prefix, which is quadratic in the paragraph.
    if !font_has_all_glyphs(family, bold, italic, text) {
        return None;
    }
    let mut sum: i64 = 0;
    for (i, ch) in text.chars().enumerate() {
        // Which run owns this character. Runs are contiguous and in order, so
        // this is a running total, not a search over the whole paragraph.
        let at = styles.line_start + i;
        let mut seen = 0usize;
        let mut run_fs = fs;
        let mut run_bold = bold;
        let mut run_italic = italic;
        for run in styles.runs {
            let n = run.text.chars().count();
            if at < seen + n {
                run_fs = run.font_size.unwrap_or(fs);
                run_bold = run.bold;
                run_italic = run.italic;
                break;
            }
            seen += n;
        }
        let em = font_adv::hmtx_advance_em(family, ch)
            .or_else(|| precise_advance_em(family, run_bold, run_italic, ch))?;
        sum += f64::from(em * run_fs * 8.0).round() as i64;
    }
    Some(sum)
}

/// One fit test for every break site: master units when derivable, the
/// legacy pixel measure otherwise.
#[cfg(windows)]
#[allow(clippy::too_many_arguments)]
fn fits_line(
    dc: windows::Win32::Graphics::Gdi::HDC,
    text: &str,
    fs: f32,
    family: &str,
    bold: bool,
    italic: bool,
    width_pt: f32,
    width_px: i32,
    scale: f64,
    styles: Option<RunStyles<'_>>,
) -> bool {
    if masterunit_on() {
        let mu = match styles.filter(|_| runmeasure_on()) {
            Some(s) => master_units_runs(text, fs, family, bold, italic, s),
            None => master_units(text, fs, family, bold, italic),
        };
        if let Some(mu) = mu {
            let fits = mu as f64 / 8.0 <= f64::from(width_pt) + 1e-6;
            if let Ok(want) = std::env::var("OXI_FIT_DEBUG") {
                if text.contains(&want) {
                    eprintln!(
                        "FIT {:>8.2}pt vs box {width_pt:.2}pt  fits={fits}  {text:?}",
                        mu as f64 / 8.0
                    );
                }
            }
            return fits;
        }
    }
    measure_wrap(dc, text, fs, family, bold, italic, scale) <= width_px
}

/// Exact width of `text` in POINTS with every character measured at its own
/// run's size, weight and slant.
///
/// S-RUNALIGN (2026-08-25). The WRAP has measured per run since S-RUNMEASURE
/// (`master_units_runs` behind `fits_line`), but the width that CENTRES or
/// RIGHT-ALIGNS the finished line did not: it took `bold` as "any run in the
/// paragraph is bold" and measured the whole line in that one style. One bold
/// word therefore made every line of its paragraph measure bold, and a centred
/// paragraph started from a width that is never drawn.
///
/// d16 slide 5 is the specimen: a 4-line centred quotation whose second line
/// carries one bold run. Every line measured in `Source Sans Pro #BI` -- which
/// that deck embeds as **Black** Italic -- so line 1 came out 531.36pt against a
/// 514.22pt text area, wider than the box, and `(area_w - line_w).max(0.0)`
/// clamped its offset to zero: the line was not centred at all.
///
/// **76 multi-run paragraphs over 13 of the 40 dev decks mix weights** (d24 11,
/// d11 10, d15 10, d19 10, d16 8, d06 7, d35 7 ...), and every one was aligned
/// from a width measured in a face half of it is not drawn in.
///
/// Master units are deliberately NOT used here: those are the BREAK model
/// (1/8pt per glyph, `pptx-master-unit-break-law`), while alignment is judged on
/// the exact advance sum, which is what the single-style path already used.
/// Returns None if any segment's advance is unknown, so the caller keeps its
/// existing fallbacks, and a single-run paragraph never reaches this at all.
#[cfg(windows)]
fn line_width_pt_runs(
    dc: windows::Win32::Graphics::Gdi::HDC,
    text: &str,
    fs: f32,
    family: &str,
    bold: bool,
    italic: bool,
    scale: f64,
    styles: RunStyles<'_>,
) -> Option<f32> {
    let style_at = |at: usize| -> (u32, bool, bool) {
        let mut seen = 0usize;
        for run in styles.runs {
            let n = run.text.chars().count();
            if at < seen + n {
                // f32 has no Eq, and the segment walk needs one; the size is a
                // half-point value, so hundredths of a point are lossless.
                let size = run.font_size.unwrap_or(fs);
                return ((size * 100.0).round() as u32, run.bold, run.italic || italic);
            }
            seen += n;
        }
        ((fs * 100.0).round() as u32, bold, italic)
    };
    let measure = |seg: &str, st: (u32, bool, bool)| -> Option<f32> {
        if seg.is_empty() {
            return Some(0.0);
        }
        let (sz, sb, si) = st;
        let sfs = sz as f32 / 100.0;
        hmtx_width_styled(seg, sfs, family, sb, si).or_else(|| {
            runtime_width_px(dc, seg, sfs, family, sb, si, scale)
                .map(|px| px as f32 / scale as f32)
        })
    };
    // Walk maximal same-style segments and measure each with the single-style
    // path, so this cannot disagree with it on a uniform paragraph.
    let mut total = 0.0f32;
    let mut seg = String::new();
    let mut cur: Option<(u32, bool, bool)> = None;
    for (i, ch) in text.chars().enumerate() {
        let st = style_at(styles.line_start + i);
        if let Some(c) = cur {
            if c != st {
                total += measure(&seg, c)?;
                seg.clear();
            }
        }
        cur = Some(st);
        seg.push(ch);
    }
    if let Some(c) = cur {
        total += measure(&seg, c)?;
    }
    Some(total)
}

/// A picture whose `a:srcRect` crops away everything is skipped unless this is
/// set (which restores drawing the whole image).
fn srcdegen_on() -> bool {
    std::env::var("OXI_SRCDEGEN_DISABLE").is_err()
}

/// `blockArc` is drawn as its ring sector unless this is set (which restores
/// painting it as its bounding box).
fn blockarc_on() -> bool {
    std::env::var("OXI_BLOCKARC_DISABLE").is_err()
}

/// `wedgeRectCallout` keeps its tail unless this is set (which restores drawing
/// it as a plain rectangle).
fn wedgecallout_on() -> bool {
    std::env::var("OXI_WEDGECALL_DISABLE").is_err()
}

/// `star10` is drawn as its ten-pointed outline unless this is set (which
/// restores painting it as its bounding box).
fn star10_on() -> bool {
    std::env::var("OXI_STAR10_DISABLE").is_err()
}

/// `bentConnector3` is drawn as its elbow unless this is set (which restores
/// drawing it as a diagonal between the box corners).
fn bentconn_on() -> bool {
    std::env::var("OXI_BENTCONN_DISABLE").is_err()
}

/// A line is aligned on a width measured per RUN unless this is set (which
/// restores measuring the whole line in the paragraph's heaviest style).
fn runalign_on() -> bool {
    std::env::var("OXI_RUNALIGN_DISABLE").is_err()
}

/// The hmtx design-advance table, but only for text it actually describes.
///
/// S-HMTXSTYLE (2026-08-25). `font_adv`'s tables are keyed by FAMILY NAME alone
/// and hold `arial` / `arialbd` / `calibri` -- no italic, and nothing ever maps a
/// bold request onto `arialbd`. So every styled run in a table family was
/// measured with the REGULAR advances while being drawn with the real bold or
/// italic face.
///
/// d17 slide 4's "HAPPY DESIGNING!" is Arial Bold Italic at 24.6pt. PowerPoint
/// draws it 492px wide; Oxi drew 489 and, once the dx array was corrected,
/// drew the right 492 but still CENTRED it on the regular-Arial width, landing
/// the line 2px right. The table has to decline styled text, not answer for it.
#[cfg(windows)]
fn hmtx_width_styled(text: &str, fs: f32, family: &str, bold: bool, italic: bool) -> Option<f32> {
    if hmtxstyle_on() && (bold || italic) {
        return None;
    }
    font_adv::line_hmtx_width_pt(text, fs, family)
}

/// The hmtx table is declined for bold / italic text unless this is set.
fn hmtxstyle_on() -> bool {
    std::env::var("OXI_HMTXSTYLE_DISABLE").is_err()
}

/// WordArt text fitting is applied unless this is set.
fn txwarp_on() -> bool {
    std::env::var("OXI_TXWARP_DISABLE").is_err()
}

/// Draw an autoshape's text stretched onto the shape box (WordArt).
///
/// Returns false when the shape does not qualify -- no single line of text, an
/// unmeasurable face -- so the caller falls back to normal layout.
#[cfg(windows)]
fn draw_warped_text(
    dc: windows::Win32::Graphics::Gdi::HDC,
    pres: &Presentation,
    sh: &Shape,
    paragraphs: &[oxislides_core::ir::SlideParagraph],
    scale: f64,
) -> bool {
    use windows::Win32::Foundation::COLORREF;
    use windows::Win32::Graphics::Gdi::*;
    use windows::core::PCWSTR;

    let texts: Vec<String> = paragraphs
        .iter()
        .map(|p| p.runs.iter().map(|r| r.text.as_str()).collect::<String>())
        .filter(|t| !t.trim().is_empty())
        .collect();
    if texts.len() != 1 || sh.width <= 0.0 || sh.height <= 0.0 {
        return false;
    }
    let text = &texts[0];
    let para = match paragraphs.iter().find(|p| {
        !p.runs
            .iter()
            .map(|r| r.text.as_str())
            .collect::<String>()
            .trim()
            .is_empty()
    }) {
        Some(p) => p,
        None => return false,
    };
    let family = effective_family(
        dc,
        &paragraph_family(pres, sh, para, &sh.ph_levels[..], &pres.master_styles.other),
    );
    let bold = para.runs.iter().any(|r| r.bold);
    let italic = para.runs.iter().any(|r| r.italic);
    let Some((ix0, ix1, iy0, iy1)) = text_ink_box_em(&family, bold, italic, text) else {
        return false;
    };
    let (ink_w, ink_h) = (ix1 - ix0, iy1 - iy0);
    if ink_w <= 0.0 || ink_h <= 0.0 {
        return false;
    }
    // Vertical scale sets the font size; the horizontal one is asked for
    // through the pen width, which is how GDI stretches glyphs.
    let fs_px = (sh.height as f64 * scale) / f64::from(ink_h);
    let target_w_px = (sh.width as f64) * scale;
    let natural_w_px = f64::from(ink_w) * fs_px;
    if fs_px < 1.0 || natural_w_px <= 0.0 {
        return false;
    }
    let (face, weight, italic_flag) = styled_face(&family, bold, italic);
    let wide: Vec<u16> = face.encode_utf16().chain(std::iter::once(0)).collect();
    unsafe {
        // Measure the face's natural average width at this height, then ask for
        // the stretched one.
        let probe = CreateFontW(
            -(fs_px.round() as i32),
            0,
            0,
            0,
            weight,
            u32::from(italic_flag),
            0,
            0,
            DEFAULT_CHARSET.0 as u32,
            OUT_DEFAULT_PRECIS.0 as u32,
            CLIP_DEFAULT_PRECIS.0 as u32,
            CLEARTYPE_QUALITY.0 as u32,
            (DEFAULT_PITCH.0 | FF_DONTCARE.0) as u32,
            PCWSTR(wide.as_ptr()),
        );
        if probe.is_invalid() {
            return false;
        }
        let old_probe = SelectObject(dc, probe);
        let mut tm = TEXTMETRICW::default();
        let ok = GetTextMetricsW(dc, &mut tm).as_bool();
        SelectObject(dc, old_probe);
        let _ = DeleteObject(probe);
        if !ok || tm.tmAveCharWidth <= 0 {
            return false;
        }
        let stretched = (f64::from(tm.tmAveCharWidth) * target_w_px / natural_w_px)
            .round()
            .max(1.0) as i32;
        let font = CreateFontW(
            -(fs_px.round() as i32),
            stretched,
            0,
            0,
            weight,
            u32::from(italic_flag),
            0,
            0,
            DEFAULT_CHARSET.0 as u32,
            OUT_DEFAULT_PRECIS.0 as u32,
            CLIP_DEFAULT_PRECIS.0 as u32,
            CLEARTYPE_QUALITY.0 as u32,
            (DEFAULT_PITCH.0 | FF_DONTCARE.0) as u32,
            PCWSTR(wide.as_ptr()),
        );
        if font.is_invalid() {
            return false;
        }
        // NOT selected into `dc` yet: the blend path needs it in a memory DC,
        // and one GDI object cannot be selected into two DCs at once -- doing
        // that left the memory DC with its default font and drew a translucent
        // RECTANGLE instead of the numeral.
        // Place the ink's top-left on the shape's: the pen sits left of the ink
        // by its left bearing and above the baseline by the ink's top.
        let k = target_w_px / natural_w_px;
        let pen_x = (f64::from(sh.x) * scale) - f64::from(ix0) * fs_px * k;
        let baseline = (f64::from(sh.y) * scale) + f64::from(iy1) * fs_px;
        let color = para
            .runs
            .iter()
            .find_map(|r| r.color.clone())
            .or_else(|| sh.fill_color.clone());
        let rgb = color.as_deref().and_then(parse_hex_rgb).unwrap_or((0, 0, 0));
        let alpha = para.runs.iter().find_map(|r| r.color_alpha).unwrap_or(1.0);
        let wtext: Vec<u16> = text.encode_utf16().collect();
        let px = pen_x.round() as i32;
        let py = baseline.round() as i32;
        // A translucent run (d35's numerals are white at 26.9%) has to be
        // blended, and GDI's text has no alpha: draw it onto a COPY of the
        // destination and blend that back with a constant alpha, so the
        // gradient behind still shows through.
        if txwarp_alpha_on() && alpha < 0.999 {
            // Clamped to the surface: a shape at the very edge would otherwise
            // copy from outside the bitmap and blend that garbage back in.
            let mut clip = windows::Win32::Foundation::RECT::default();
            let _ = GetClipBox(dc, &mut clip);
            let bx = ((f64::from(sh.x) * scale).floor() as i32 - 4).max(clip.left);
            let by = ((f64::from(sh.y) * scale).floor() as i32 - 4).max(clip.top);
            let bw = ((f64::from(sh.x + sh.width) * scale).ceil() as i32 + 4)
                .min(clip.right)
                - bx;
            let bh = ((f64::from(sh.y + sh.height) * scale).ceil() as i32 + 4)
                .min(clip.bottom)
                - by;
            let mem = CreateCompatibleDC(dc);
            let bmp = CreateCompatibleBitmap(dc, bw.max(1), bh.max(1));
            if bw > 0 && bh > 0 && !mem.0.is_null() && !bmp.is_invalid() {
                let old_bmp = SelectObject(mem, bmp);
                let _ = BitBlt(mem, 0, 0, bw, bh, dc, bx, by, SRCCOPY);
                // A fresh DC is OPAQUE over white: without this the text cell
                // is painted white and swallows a white glyph whole.
                SetBkMode(mem, TRANSPARENT);
                let old_f = SelectObject(mem, font);
                SetTextColor(mem, COLORREF(colorref(rgb.0, rgb.1, rgb.2)));
                let old_a = SetTextAlign(mem, TA_LEFT | TA_BASELINE);
                let _ = TextOutW(mem, px - bx, py - by, &wtext);
                SetTextAlign(mem, TEXT_ALIGN_OPTIONS(old_a));
                SelectObject(mem, old_f);
                let bf = BLENDFUNCTION {
                    BlendOp: AC_SRC_OVER as u8,
                    BlendFlags: 0,
                    SourceConstantAlpha: (alpha * 255.0_f32).round().clamp(0.0, 255.0) as u8,
                    AlphaFormat: 0,
                };
                let _ = AlphaBlend(dc, bx, by, bw, bh, mem, 0, 0, bw, bh, bf);
                SelectObject(mem, old_bmp);
                let _ = DeleteObject(bmp);
                let _ = DeleteDC(mem);
                let _ = DeleteObject(font);
                return true;
            }
            if !mem.0.is_null() {
                let _ = DeleteDC(mem);
            }
            if !bmp.is_invalid() {
                let _ = DeleteObject(bmp);
            }
        }
        let old = SelectObject(dc, font);
        let old_color = SetTextColor(dc, COLORREF(colorref(rgb.0, rgb.1, rgb.2)));
        let old_align = SetTextAlign(dc, TA_LEFT | TA_BASELINE);
        let _ = TextOutW(dc, px, py, &wtext);
        SetTextAlign(dc, TEXT_ALIGN_OPTIONS(old_align));
        SetTextColor(dc, old_color);
        SelectObject(dc, old);
        let _ = DeleteObject(font);
    }
    true
}

/// The ink box of `text` in EM units: (left, right, bottom, top) where the
/// vertical pair is measured up from the baseline.
///
/// `GetGlyphOutlineW(GGO_METRICS)` reports each glyph's black box and its
/// origin relative to the pen, which is what a WordArt fit needs -- the
/// advance width is not the ink.
#[cfg(windows)]
fn text_ink_box_em(
    family: &str,
    bold: bool,
    italic: bool,
    text: &str,
) -> Option<(f32, f32, f32, f32)> {
    use windows::Win32::Graphics::Gdi::*;

    const PROBE: i32 = 2048;
    let dc = probe_dc();
    let (face, weight, italic_flag) = styled_face(family, bold, italic);
    let wide: Vec<u16> = face.encode_utf16().chain(std::iter::once(0)).collect();
    unsafe {
        let font = CreateFontW(
            -PROBE,
            0,
            0,
            0,
            weight,
            u32::from(italic_flag),
            0,
            0,
            DEFAULT_CHARSET.0 as u32,
            OUT_DEFAULT_PRECIS.0 as u32,
            CLIP_DEFAULT_PRECIS.0 as u32,
            CLEARTYPE_QUALITY.0 as u32,
            (DEFAULT_PITCH.0 | FF_DONTCARE.0) as u32,
            windows::core::PCWSTR(wide.as_ptr()),
        );
        if font.is_invalid() {
            return None;
        }
        let old = SelectObject(dc, font);
        let mut pen = 0.0f32;
        let (mut x0, mut x1, mut y0, mut y1) = (f32::MAX, f32::MIN, f32::MAX, f32::MIN);
        let mat = MAT2 {
            eM11: FIXED { fract: 0, value: 1 },
            eM12: FIXED { fract: 0, value: 0 },
            eM21: FIXED { fract: 0, value: 0 },
            eM22: FIXED { fract: 0, value: 1 },
        };
        for ch in text.chars() {
            let mut gm = GLYPHMETRICS::default();
            let got = GetGlyphOutlineW(
                dc,
                ch as u32,
                GGO_METRICS,
                &mut gm,
                0,
                None,
                &mat,
            );
            if got == GDI_ERROR as u32 {
                SelectObject(dc, old);
                let _ = DeleteObject(font);
                return None;
            }
            if gm.gmBlackBoxX > 0 && gm.gmBlackBoxY > 0 {
                let gx0 = pen + gm.gmptGlyphOrigin.x as f32;
                let gx1 = gx0 + gm.gmBlackBoxX as f32;
                let gy1 = gm.gmptGlyphOrigin.y as f32;
                let gy0 = gy1 - gm.gmBlackBoxY as f32;
                x0 = x0.min(gx0);
                x1 = x1.max(gx1);
                y0 = y0.min(gy0);
                y1 = y1.max(gy1);
            }
            pen += gm.gmCellIncX as f32;
        }
        SelectObject(dc, old);
        let _ = DeleteObject(font);
        if x1 <= x0 || y1 <= y0 {
            return None;
        }
        let em = PROBE as f32;
        Some((x0 / em, x1 / em, y0 / em, y1 / em))
    }
}

/// True when `family` itself contains a glyph for every char of `text`.
#[cfg(windows)]
fn font_has_all_glyphs(family: &str, bold: bool, italic: bool, text: &str) -> bool {
    use windows::Win32::Graphics::Gdi::*;

    if text.is_empty() {
        return false;
    }
    let dc = probe_dc();
    let (face, weight, italic) = styled_face(family, bold, italic);
    let wide: Vec<u16> = face.encode_utf16().chain(std::iter::once(0)).collect();
    let wtext: Vec<u16> = text.encode_utf16().collect();
    unsafe {
        let font = CreateFontW(
            -ADVANCE_PROBE_EM,
            0,
            0,
            0,
            weight,
            u32::from(italic),
            0,
            0,
            DEFAULT_CHARSET.0 as u32,
            OUT_DEFAULT_PRECIS.0 as u32,
            CLIP_DEFAULT_PRECIS.0 as u32,
            CLEARTYPE_QUALITY.0 as u32,
            (DEFAULT_PITCH.0 | FF_DONTCARE.0) as u32,
            windows::core::PCWSTR(wide.as_ptr()),
        );
        if font.is_invalid() {
            return false;
        }
        let old = SelectObject(dc, font);
        let mut indices = vec![0u16; wtext.len()];
        let wtext_z: Vec<u16> = wtext.iter().copied().chain(std::iter::once(0)).collect();
        let n = GetGlyphIndicesW(
            dc,
            windows::core::PCWSTR(wtext_z.as_ptr()),
            wtext.len() as i32,
            indices.as_mut_ptr(),
            GGI_MARK_NONEXISTING_GLYPHS,
        );
        SelectObject(dc, old);
        let _ = DeleteObject(font);
        n != GDI_ERROR as u32 && indices.iter().all(|g| *g != 0xFFFF)
    }
}

/// Per-character device advances for `text` at the design metrics, or None
/// when any character has no advance in this font (then the caller keeps the
/// GDI path rather than mixing two metric sources within one line).
#[cfg(windows)]
fn runtime_dx_px(
    _dc: windows::Win32::Graphics::Gdi::HDC,
    text: &str,
    fs: f32,
    family: &str,
    bold: bool,
    italic: bool,
    scale: f64,
) -> Option<Vec<i32>> {
    if !advance_exact_on() {
        return None;
    }
    // Round the CUMULATIVE position, not each advance. `ExtTextOutW` needs
    // integer steps, but rounding every advance on its own makes the drawn line
    // the sum of 55 roundings -- several pixels adrift of the design width, which
    // is what PowerPoint both breaks and centres against. Taking differences of
    // rounded running positions keeps every glyph within half a pixel of its
    // exact place AND makes the total equal the rounded exact width, so the same
    // number can be used for wrapping, for alignment, and for drawing.
    let mut dx = Vec::with_capacity(text.len());
    let mut acc = 0.0f64;
    let mut prev = 0i32;
    let plan = if emoji_on() {
        run_plan(family, bold, italic, text)
    } else {
        None
    };
    // ★Without a plan, only a font that has EVERY glyph may be measured.
    // `GetCharABCWidthsW` alone consults GDI's font-link chain and its success
    // is not stable run to run -- the same binary drew d19 slide 39's icon row
    // shifted by one glyph on a second run while the layout dump stayed
    // byte-identical, i.e. the non-determinism was entirely in whether this
    // path was taken. Asking for glyph indices with
    // GGI_MARK_NONEXISTING_GLYPHS is a direct, stable question about the font,
    // and it is also the correct one: a fallback glyph must not be advanced by
    // the base font's metrics.
    //
    // A plan is the stable answer to the same question, character by
    // character: every one of them names the face that owns it, so an emoji
    // the family lacks is measured against the emoji face's own advance
    // instead of disqualifying the whole run.
    if plan.is_none() && !font_has_all_glyphs(family, bold, italic, text) {
        return None;
    }
    for (i, ch) in text.chars().enumerate() {
        let em = match plan.as_ref().map(|p| p[i]) {
            Some(CharPlan::Base(em) | CharPlan::Color(_, em) | CharPlan::Symbol(_, em)) => em,
            Some(CharPlan::Skip) => 0.0,
            None => runtime_advance_em(family, bold, italic, ch)?,
        };
        acc += em as f64 * fs as f64 * scale;
        let pos = if advwidth_on() {
            acc.round() as i32
        } else {
            prev + (em * fs * scale as f32).round() as i32
        };
        dx.push(pos - prev);
        prev = pos;
    }
    Some(dx)
}

/// Design width of `text` in device pixels, or None (see `runtime_dx_px`).
///
/// Summed in EM units and rounded ONCE. `runtime_dx_px` has to round every
/// character on its own because `ExtTextOutW` takes an integer dx array, but a
/// line's width is not the sum of those roundings -- PowerPoint measures the
/// exact design width, and d28 slide 8's centred body lines came out 0.4% to
/// 2.0% wide, enough to shift a 250pt line by up to 5pt.
#[cfg(windows)]
fn runtime_width_px(
    dc: windows::Win32::Graphics::Gdi::HDC,
    text: &str,
    fs: f32,
    family: &str,
    bold: bool,
    italic: bool,
    scale: f64,
) -> Option<i32> {
    runtime_dx_px(dc, text, fs, family, bold, italic, scale).map(|dx| dx.iter().sum())
}

/// Glyph positions are rounded cumulatively unless this is set, which restores
/// rounding each advance on its own.
fn advwidth_on() -> bool {
    std::env::var("OXI_ADVWIDTH_DISABLE").is_err()
}

/// Text is measured and drawn at the font's design advances unless this is
/// set, which restores GDI's hinted metrics.
fn advance_exact_on() -> bool {
    std::env::var("OXI_ADVEXACT_DISABLE").is_err()
}

/// Measure `text` in device pixels with a font this call creates itself.
///
/// `gdi_measure_text_px` needs the caller to have selected the font already;
/// the table path draws one cell at a time and has no font selected, so it
/// needs the self-contained form.
#[cfg(windows)]
fn measure_text_width(
    dc: windows::Win32::Graphics::Gdi::HDC,
    text: &str,
    font_size: f32,
    family: &str,
    bold: bool,
    scale: f64,
) -> f64 {
    use windows::Win32::Graphics::Gdi::*;
    use windows::core::PCWSTR;

    let height = (font_size as f64 * scale).round() as i32;
    let (face, weight, italic) = styled_face(family, bold, false);
    let wide: Vec<u16> = face.encode_utf16().chain(std::iter::once(0)).collect();
    unsafe {
        let font = CreateFontW(
            -height,
            0,
            0,
            0,
            weight,
            u32::from(italic),
            0,
            0,
            DEFAULT_CHARSET.0 as u32,
            OUT_DEFAULT_PRECIS.0 as u32,
            CLIP_DEFAULT_PRECIS.0 as u32,
            CLEARTYPE_QUALITY.0 as u32,
            (DEFAULT_PITCH.0 | FF_DONTCARE.0) as u32,
            PCWSTR(wide.as_ptr()),
        );
        let old = SelectObject(dc, font);
        let w = gdi_measure_text_px(dc, text);
        SelectObject(dc, old);
        let _ = DeleteObject(font);
        w as f64
    }
}

/// Measure the width of `text` in device pixels (font must be selected).
#[cfg(windows)]
fn gdi_measure_text_px(dc: windows::Win32::Graphics::Gdi::HDC, text: &str) -> i32 {
    use windows::Win32::Foundation::*;
    use windows::Win32::Graphics::Gdi::*;
    let wtext: Vec<u16> = text.encode_utf16().collect();
    let mut size = SIZE::default();
    unsafe {
        let _ = GetTextExtentPoint32W(dc, &wtext, &mut size);
    }
    size.cx
}

/// The width the wrap judges against -- the same chain the `fits` test uses.
#[cfg(windows)]
fn measure_wrap(
    dc: windows::Win32::Graphics::Gdi::HDC,
    text: &str,
    fs: f32,
    family: &str,
    bold: bool,
    italic: bool,
    scale: f64,
) -> i32 {
    let width = if advance_exact_on() {
        runtime_width_px(dc, text, fs, family, bold, italic, scale)
            .or_else(|| font_adv::text_hmtx_px(text, fs, family, scale))
            .unwrap_or_else(|| gdi_measure_text_px(dc, text))
    } else {
        gdi_measure_text_px(dc, text)
    };
    width + trailing_overhang_px(text, fs, family, bold, italic, scale)
}

/// The room the last glyph on a candidate line needs beyond its advance.
///
/// FALSIFIED as a default (2026-08-21) and held opt-in: adding the full GDI
/// ABC-C overhang flipped 20 knife-edge lines across six dev decks and the
/// corpus said 1 improved / 4 regressed -- PowerPoint keeps lines this term
/// breaks (d20 "Table Of Contents", d34 Pacifico at 0.002pt slack). Yet the
/// pure design-advance sum is not PowerPoint's test either: it breaks four
/// lines that float-fit with 0.01-0.10pt to spare (d09 "Happy Holi!", d20
/// "About Us" / "Who we are?" / "Our Projects"), and coarse per-glyph
/// quantisation (96dpi int / 1/16px) contradicts the Pacifico keeps outright
/// (`tools/metrics/analyze_pptx_wrapfit.py` scores every hypothesis against
/// all twenty cases). Whatever PowerPoint adds is far smaller than the ABC-C
/// ink reach; pinning it needs a width-sweep repro against PowerPoint COM.
#[cfg(windows)]
fn trailing_overhang_px(
    text: &str,
    fs: f32,
    family: &str,
    bold: bool,
    italic: bool,
    scale: f64,
) -> i32 {
    if std::env::var("OXI_TRAILINK_ENABLE").is_err() {
        return 0;
    }
    let Some(last) = text.chars().rev().find(|c| !c.is_whitespace()) else {
        return 0;
    };
    let Some(em) = runtime_overhang_em(family, bold, italic, last) else {
        return 0;
    };
    (f64::from(em * fs) * scale).round() as i32
}

/// A word wider than its line breaks inside itself unless this is set.
fn charwrap_on() -> bool {
    std::env::var("OXI_CHARWRAP_DISABLE").is_err()
}

/// A hyphen offers a break after it unless this is set.
fn hyphbrk_on() -> bool {
    std::env::var("OXI_HYPHBRK_DISABLE").is_err()
}

/// The pieces a line may end on.
///
/// Spaces are the obvious opportunity; a hyphen is the other one PowerPoint
/// honours. The `charwrap` probe's `alpha-beta-gamma-delta-epsilon-zeta-...`
/// in a 165.6pt box came back as `alpha-beta-gamma-` / `delta-epsilon-zeta-` /
/// `eta-theta-iota`, so the break lands AFTER the hyphen and not at the last
/// character that fits (which would have left `...gamma-de` on line 0).
///
/// Slash and dot are NOT opportunities: the same probe's
/// `https://www.example.com/some/rather/long/path/index.html` broke inside
/// `long` rather than after `rather/`.
///
/// 76 paragraphs across 31 dev decks carry a mid-word hyphen.
fn break_pieces(text: &str) -> Vec<&str> {
    if !hyphbrk_on() {
        return text.split_inclusive(' ').collect();
    }
    let mut out = Vec::new();
    for chunk in text.split_inclusive(' ') {
        let mut start = 0usize;
        for (i, ch) in chunk.char_indices() {
            // A trailing hyphen would otherwise produce an empty tail piece.
            if ch == '-' && i + 1 < chunk.len() {
                out.push(&chunk[start..i + 1]);
                start = i + 1;
            }
        }
        out.push(&chunk[start..]);
    }
    out
}

/// Wrap `text` at word boundaries to fit `effective_width_pt`.
#[cfg(windows)]
fn gdi_wrap_lines(
    dc: windows::Win32::Graphics::Gdi::HDC,
    text: &str,
    first_width_pt: f32,
    rest_width_pt: f32,
    scale: f64,
    fs: f32,
    family: &str,
    bold: bool,
    italic: bool,
    runs: Option<(&[oxislides_core::ir::SlideRun], usize)>,
) -> Vec<String> {
    let first_px = (first_width_pt as f64 * scale).round().max(1.0) as i32;
    let rest_px = (rest_width_pt as f64 * scale).round().max(1.0) as i32;
    let mut width_px = first_px;
    let mut width_pt = first_width_pt;
    let mut lines: Vec<String> = Vec::new();
    let mut current = String::new();
    let mut current_w = 0i32;
    // Characters already committed to finished lines: the candidate's offset in
    // the paragraph, which is what maps a character back to its run.
    let mut emitted = 0usize;
    let styles = |start: usize| {
        runs.map(|(runs, base)| RunStyles {
            runs,
            line_start: base + start,
        })
    };
    let trim_on = std::env::var("OXI_WRAPTRIM_DISABLE").is_err();
    for word in break_pieces(text) {
        // A line's trailing space HANGS past the right edge -- it is not part
        // of the width the break is judged against. Measured on d28 slide 13
        // (2026-08-18): "National Cemetery in Gettysburg, Pennsylvania. In
        // just" is 1034px in its own font against a 1036px box and PowerPoint
        // keeps it whole, but with the trailing space it is 1047px, so the
        // per-word accumulation broke before "just" and the paragraph needed
        // 11 lines where PowerPoint needs 10.
        //
        // Measuring the candidate PREFIX rather than summing per-word widths
        // also drops the per-word integer-pixel rounding, which pushed the
        // same way.
        let fits = if trim_on {
            let mut candidate = current.clone();
            candidate.push_str(word);
            let trimmed = candidate.trim_end();
            fits_line(dc, trimmed, fs, family, bold, italic, width_pt, width_px, scale, styles(emitted))
        } else {
            current_w + gdi_measure_text_px(dc, word) <= width_px
        };
        if !current.is_empty() && !fits {
            emitted += current.chars().count();
            lines.push(std::mem::take(&mut current));
            current_w = 0;
            // Every line after the first is judged against the continuation
            // width, which a hanging indent or a bullet makes narrower.
            width_px = rest_px;
            width_pt = rest_width_pt;
        }
        // A single "word" wider than the line has to break INSIDE itself --
        // splitting on spaces alone leaves it as one overflowing line. d11 and
        // d24 slide 38 are 53 emoji with no space between them in a 490pt box;
        // PowerPoint lays them out in four rows and Oxi drew one that ran off
        // the page. 45 paragraphs across nine decks carry a space-free run of
        // 30 characters or more (long URLs are the other kind).
        if charwrap_on() && current.is_empty() {
            let mut rest = word;
            loop {
                let trimmed = rest.trim_end();
                if trimmed.is_empty()
                    || fits_line(
                        dc, trimmed, fs, family, bold, italic, width_pt, width_px, scale,
                        styles(emitted),
                    )
                {
                    break;
                }
                // Longest prefix that fits, never empty so the loop ends.
                let mut last_ok = 0usize;
                for (i, ch) in rest.char_indices() {
                    let end = i + ch.len_utf8();
                    if fits_line(
                        dc, rest[..end].trim_end(), fs, family, bold, italic,
                        width_pt, width_px, scale, styles(emitted),
                    ) {
                        last_ok = end;
                    } else {
                        break;
                    }
                }
                let take = if last_ok > 0 {
                    last_ok
                } else {
                    rest.char_indices().nth(1).map(|(i, _)| i).unwrap_or(rest.len())
                };
                emitted += rest[..take].chars().count();
                lines.push(rest[..take].to_string());
                rest = &rest[take..];
                width_px = rest_px;
                width_pt = rest_width_pt;
            }
            current.push_str(rest);
            current_w += gdi_measure_text_px(dc, rest);
            continue;
        }
        current.push_str(word);
        current_w += gdi_measure_text_px(dc, word);
    }
    if !current.is_empty() {
        lines.push(current);
    }
    if lines.is_empty() {
        lines.push(String::new());
    }
    if let Ok(want) = std::env::var("OXI_LINE_DEBUG") {
        if text.contains(&want) {
            eprintln!(
                "LINE fam={family:?} fs={fs} bold={bold} first={first_width_pt} rest={rest_width_pt} \
                 scale={scale:.4} mu={:?} lines={lines:?}",
                master_units(text, fs, family, bold, italic),
            );
        }
    }
    lines
}

/// First-line baseline offset factor (em) measured from PowerPoint render-truth
/// (Ra loop, spec6). The first line of a single-spaced (n == 1) paragraph sits at
/// text-area-top + em*fs.
///
/// Derived via two independent probes (spec6_baseline/):
///   * multitop   — fs=192, shape top swept 12pt, 12 points/font  -> em = (X-3.6)/192
///   * fssweep    — fs {8..192} regression, box top = 0            -> tIns = 3.6pt (margin_top
///                  0.05in) and em (slope) confirmed independent of fs
/// em is font-specific and constant across font size. The closed-form model
/// `1.2 * (win_asc + 0.5) / win_total` (model1b) reproduces all six measured
/// fonts to ~±0.0003 em, but the measured table is exact for these fonts, so it
/// is used directly; unknown fonts fall back to the model1b-style average.
#[cfg(windows)]
/// Where the baseline sits inside a line box, as a multiple of the font size.
///
/// The line box is `1.2 * size` (probe `lineheight`, 8 faces; `ascentsplit`,
/// 12 more -- all 1.2010 within the PDF's 0.03pt quantum), and probe
/// `mixedpitch` showed the step between two lines of DIFFERENT size is
/// `descent(first) + ascent(second)`, so this one number decides both the first
/// baseline and every paragraph boundary.
///
/// The split is the FONT's own, read from OS/2 (probe `ascentsplit`, 12 faces
/// spanning 0.88 to 1.06):
///
/// * `fsSelection` bit 7 (USE_TYPO_METRICS) clear -> `usWinAscent` /
///   `usWinDescent`. Arial 0.9724 vs measured 0.9728, Goudy Stout 0.8938 vs
///   0.8948, Haettenschweiler 1.0614 vs 1.0627 -- worst error 0.0021.
/// * bit set -> `sTypoAscender + sTypoLineGap` over that plus `-sTypoDescender`.
///   Noto Serif 0.9419 vs 0.9427, Reem Kufi 0.8800 vs 0.8797, and the two faces
///   d28 EMBEDS: Calistoga 0.9231 vs 0.9243, Jua (the one with a line gap)
///   1.0080 vs 1.0083.
///
/// The line gap belongs on the ascent side and NOT in the win branch: Jua's
/// 250-unit gap is what takes it from 0.96 to 1.008, while adding Arial's hhea
/// gap to its win ascent would give 0.9789 against a measured 0.9728.
fn font_baseline_offset_em(family: &str) -> f32 {
    if rtbaseline_on() {
        if let Some(a) = runtime_baseline_offset_em(family) {
            return a;
        }
    }
    // Measured before the rule was derived; kept as the offline fallback for a
    // face GDI cannot hand back tables for. Each is within 0.0005 of what the
    // rule computes for it.
    match family.to_ascii_lowercase().as_str() {
        "arial" => 0.97274,
        "times new roman" => 0.96587,
        "calibri" => 0.93648,
        "segoe ui" => 0.97399,
        "georgia" => 0.96899,
        "verdana" => 0.99275,
        _ => 0.9685,
    }
}

/// Read the ascent split straight out of the resolved face's OS/2 table.
#[cfg(windows)]
fn runtime_baseline_offset_em(family: &str) -> Option<f32> {
    use windows::Win32::Graphics::Gdi::*;

    BASELINE_CACHE.with(|cache| {
        if let Some(hit) = cache.borrow().get(family) {
            return *hit;
        }
        let dc = probe_dc();
        let wide: Vec<u16> = family.encode_utf16().chain(std::iter::once(0)).collect();
        let value = unsafe {
            let font = CreateFontW(
                -ADVANCE_PROBE_EM,
                0,
                0,
                0,
                400,
                0,
                0,
                0,
                DEFAULT_CHARSET.0 as u32,
                OUT_DEFAULT_PRECIS.0 as u32,
                CLIP_DEFAULT_PRECIS.0 as u32,
                CLEARTYPE_QUALITY.0 as u32,
                (DEFAULT_PITCH.0 | FF_DONTCARE.0) as u32,
                windows::core::PCWSTR(wide.as_ptr()),
            );
            if font.is_invalid() {
                None
            } else {
                let old = SelectObject(dc, font);
                let os2 = read_font_table(dc, b"OS/2");
                SelectObject(dc, old);
                let _ = DeleteObject(font);
                os2.filter(|t| t.len() >= 78).and_then(|t| {
                    let u16at = |o: usize| u16::from_be_bytes([t[o], t[o + 1]]) as f32;
                    let i16at = |o: usize| i16::from_be_bytes([t[o], t[o + 1]]) as f32;
                    let use_typo = (u16at(62) as u32) & 0x80 != 0;
                    let (asc, desc) = if use_typo {
                        (i16at(68) + i16at(72), -i16at(70))
                    } else {
                        (u16at(74), u16at(76))
                    };
                    // The numbers behind every line-box decision. Printing them
                    // is the difference between measuring what the renderer
                    // reads and inferring it from where a baseline landed --
                    // which is how a font whose OS/2 disagrees with PowerPoint
                    // gets mistaken for a broken layout rule.
                    if std::env::var("OXI_DEBUG_BASELINE").is_ok() {
                        eprintln!(
                            "BASELINE {family:24} use_typo={use_typo} asc={asc} desc={desc}                              sTypo=({}, {}, gap {}) usWin=({}, {}) em={:.4}",
                            i16at(68),
                            i16at(70),
                            i16at(72),
                            u16at(74),
                            u16at(76),
                            1.2 * asc / (asc + desc),
                        );
                    }
                    if asc + desc > 0.0 {
                        Some(1.2 * asc / (asc + desc))
                    } else {
                        None
                    }
                })
            }
        };
        cache.borrow_mut().insert(family.to_string(), value);
        value
    })
}

/// `a:bodyPr/@wrap="none"` is honoured unless this is set.
fn wrapnone_on() -> bool {
    std::env::var("OXI_WRAPNONE_DISABLE").is_err()
}

/// The extra inset a non-rectangular preset geometry puts on its own text.
///
/// PowerPoint lays a shape's text out in the rectangle INSCRIBED in its
/// geometry, not in its bounding box. For an ellipse that is `w/sqrt2` by
/// `h/sqrt2`, centred -- and d35 s34's competitor matrix says so to a
/// hundredth of a point on all seven of its bubbles:
///
///   w=78.76 "Our company" (50.87pt) fits 55.69      -> one line, and it is
///   w=68.98 "Competitor"  (42.61pt) fits 48.77      -> one line, twice
///   w=57.52 "Competitor"  (42.61pt) EXCEEDS 40.67   -> breaks, three times
///   w=36.78 "Competitor"           exceeds 26.01    -> breaks after "Comp"
///
/// and the horizontal origin lands on it too: the w=57.52 bubble at x=108.89
/// has an inscribed left edge of 117.315, its 39.39pt first line centres at
/// **117.955**, and PowerPoint drew it at **117.96**. Reading the box at full
/// width fits every one of those lines on one line instead.
///
/// 108 text-bearing ellipses across 8 dev decks. Opt-out
/// `OXI_GEOMINSET_DISABLE`.
fn geom_text_inset(sh: &Shape) -> (f32, f32) {
    if !geominset_on() {
        return (0.0, 0.0);
    }
    match sh.shape_type.as_deref() {
        Some("ellipse") => {
            let k = (1.0 - std::f32::consts::FRAC_1_SQRT_2) / 2.0;
            (sh.width * k, sh.height * k)
        }
        _ => (0.0, 0.0),
    }
}

/// A preset geometry insets its own text unless this is set.
fn geominset_on() -> bool {
    std::env::var("OXI_GEOMINSET_DISABLE").is_err()
}

/// A picture covers every pixel its exact rectangle TOUCHES unless this is
/// set, which restores a rectangle rounded from the shape's rounded box.
///
/// PowerPoint's export is vector: the rasteriser antialiases a fractional edge,
/// so the image reaches into the pixel row that its top only partly covers.
/// d08 s11's lower photo sits at y=202.5pt = 421.875px and PowerPoint's content
/// starts in row **421**, while Oxi rounded to 422 and stretched the source into
/// 422 rows instead of 423. The error is not a shift but a SCALE: 1px at the top
/// of the photo, nothing at the bottom, which is exactly `[round, round+round)`
/// against `[floor, ceil)`.
fn imgrect_on() -> bool {
    std::env::var("OXI_IMGRECT_DISABLE").is_err()
}

/// A soft break that changes SIZE steps by the mixed-pitch rule unless this is
/// set, which restores the paragraph's flat advance for every line in it.
///
/// d11 s33's caption is one paragraph: "Imani Jackson" at 12pt, `<a:br/>`, then
/// "JOB TITLE" at 8.04pt. PowerPoint steps 10.560pt between those baselines --
/// the 12pt line's descent plus the 8.04pt line's ascent -- and the flat rule
/// steps 14.40. The 3.81pt is then carried by every line after it, so the whole
/// block below the photos sits low. The same template is in d16 and d35.
fn brpitch_on() -> bool {
    std::env::var("OXI_BRPITCH_DISABLE").is_err()
}

/// `<a:br/>` ends the line it stands on unless this is set.
fn softbreak_on() -> bool {
    std::env::var("OXI_SOFTBREAK_DISABLE").is_err()
}

/// A table row grows to fit its cells unless this is set.
fn tblgrow_on() -> bool {
    std::env::var("OXI_TBLGROW_DISABLE").is_err()
}

/// A level's `a:defRPr/@i` is honoured unless this is set.
fn lvlitalic_on() -> bool {
    std::env::var("OXI_LVLITALIC_DISABLE").is_err()
}

/// Where the FIRST baseline of a text block sits below the text-area top.
///
/// Derived 2026-08-23 from PowerPoint's own export, probe `firstline`: 31 arms
/// over Arial / Calibri / Verdana / Georgia / Segoe Script / Comic Sans MS x
/// `lnSpc` 70 / 80 / 90 / 95 / 100 / 110 / 120% x 20 / 40 / 60pt, in a box with
/// every inset 0 and `anchor="t"`, so the measured baseline IS the offset.
///
/// The rule is not about the ascent at all -- it is about the DESCENT, and only
/// then does it collapse to one expression. With `P = 1.2 * fs` the natural line
/// box and `D0 = P - face * fs` the face's own share of it below the baseline:
///
///     n <= 1:  D = max( D0 + 0.25 * P * (n - 1),  min(D0, 0.25 * P * n) )
///     n >  1:  D = max( D0,                       0.25 * P * n          )
///     off = P * n - D
///
/// i.e. the baseline sits its own descent above the box's bottom; that descent
/// is **capped at a quarter of the box**, and a face already deeper than that
/// quarter gives up a quarter of whatever the box loses. Above single spacing
/// the quarter becomes a FLOOR instead of a cap, which is why 120% reads
/// 43.25pt at 40pt for both Arial and Calibri although their faces differ by
/// 0.036 em, while Segoe Script -- whose descent is deeper than the quarter --
/// keeps its own and lands at 42.65.
///
/// All 31 arms match within 0.072pt and every residual is POSITIVE: the per-face
/// constant already present at n == 1 (Arial 0.035, Calibri 0.050, Verdana
/// 0.062, Georgia 0.072, Segoe Script 0.054, Comic Sans MS 0.011) plus a flat
/// 0.050 wherever the quarter binds. That is the existing baseline residual, not
/// a gap in this rule.
///
/// ★What the first attempt got wrong, and why the corpus caught it: reading the
/// same data as an ASCENT rule gives `max(face - (1-n) * P, 0.75 * P * n)`,
/// which fits every face whose `face` exceeds 0.75 * P and none that does not.
/// The four faces of the first probe round were all above it. d04 slide 1 --
/// 58pt **Satisfy**, whose face is 0.7877 -- was 6.5pt low under that rule, and
/// the corpus read 114 slides down against 116 up. Segoe Script (0.8249) is the
/// installed face that reproduces it, and it is in the probe for that reason.
#[cfg(windows)]
fn first_baseline_off(family: &str, fs: f32, n: f32) -> f32 {
    if !firstline_on() {
        // ★The parenthesis is load-bearing: the pre-change code computed
        // `0.75 * adv` with `adv = fs * 1.2 * n`, and `0.75 * fs * 1.2 * n`
        // associates the other way. The 1-ULP difference flipped one page of
        // d37 in the opt-out arm -- a byte-identity proof lost to float
        // association, not to a leak.
        return if (n - 1.0).abs() > 1e-4 {
            0.75 * (fs * 1.2 * n)
        } else {
            font_baseline_offset_em(family) * fs
        };
    }
    let pitch = fs * 1.2;
    let natural_descent = pitch - font_baseline_offset_em(family) * fs;
    let quarter = 0.25 * pitch;
    let descent = if n <= 1.0 {
        (natural_descent + quarter * (n - 1.0)).max(natural_descent.min(quarter * n))
    } else {
        natural_descent.max(quarter * n)
    };
    pitch * n - descent
}

/// A level's `a:defRPr/@b` is honoured unless this is set.
fn lvlbold_on() -> bool {
    std::env::var("OXI_LVLBOLD_DISABLE").is_err()
}

/// The joined first-baseline rule is used unless this is set, which restores
/// the two unrelated models either side of `lnSpc` 100%.
fn firstline_on() -> bool {
    std::env::var("OXI_FIRSTLINE_DISABLE").is_err()
}

/// A run with no colour takes the LEVEL's rather than a sibling run's unless
/// this is set.
fn runcolordef_on() -> bool {
    std::env::var("OXI_RUNCOLORDEF_DISABLE").is_err()
}

/// A level's inherited `a:highlight` is honoured unless this is set.
fn highlightlvl_on() -> bool {
    std::env::var("OXI_HIGHLIGHTLVL_DISABLE").is_err()
}

/// `a:highlight` boxes are drawn unless this is set.
fn highlight_on() -> bool {
    std::env::var("OXI_HIGHLIGHT_DISABLE").is_err()
}

#[cfg(windows)]
thread_local! {
    static HHEA_CACHE: std::cell::RefCell<
        std::collections::HashMap<String, Option<(f32, f32)>>,
    > = std::cell::RefCell::new(std::collections::HashMap::new());
}

/// A face's hhea ascent and descent, in em.
///
/// This is the pair the highlight box is built from, and it is NOT the pair
/// `runtime_baseline_offset_em` uses: that one follows fsSelection bit 7 to the
/// typo metrics, and the two faces measured that set the bit -- Bahnschrift
/// (0.7527 / 0.2473) and Cascadia Mono (0.9201 / 0.2420) -- came out of
/// PowerPoint's export on their hhea values, not their typo ones
/// (`highlight` probe, 2026-08-19).
#[cfg(windows)]
fn hhea_extent_em(family: &str) -> Option<(f32, f32)> {
    use windows::Win32::Graphics::Gdi::*;

    HHEA_CACHE.with(|cache| {
        if let Some(hit) = cache.borrow().get(family) {
            return *hit;
        }
        let dc = probe_dc();
        let wide: Vec<u16> = family.encode_utf16().chain(std::iter::once(0)).collect();
        let value = unsafe {
            let font = CreateFontW(
                -ADVANCE_PROBE_EM,
                0,
                0,
                0,
                400,
                0,
                0,
                0,
                DEFAULT_CHARSET.0 as u32,
                OUT_DEFAULT_PRECIS.0 as u32,
                CLIP_DEFAULT_PRECIS.0 as u32,
                CLEARTYPE_QUALITY.0 as u32,
                (DEFAULT_PITCH.0 | FF_DONTCARE.0) as u32,
                windows::core::PCWSTR(wide.as_ptr()),
            );
            if font.is_invalid() {
                None
            } else {
                let old = SelectObject(dc, font);
                let parsed = (|| {
                    let hhea = read_font_table(dc, b"hhea")?;
                    let head = read_font_table(dc, b"head")?;
                    let upem = u16::from_be_bytes([*head.get(18)?, *head.get(19)?]) as f32;
                    let asc = i16::from_be_bytes([*hhea.get(4)?, *hhea.get(5)?]) as f32;
                    let desc = i16::from_be_bytes([*hhea.get(6)?, *hhea.get(7)?]) as f32;
                    if upem <= 0.0 {
                        return None;
                    }
                    Some((asc / upem, -desc / upem))
                })();
                SelectObject(dc, old);
                let _ = DeleteObject(font);
                parsed
            }
        };
        cache.borrow_mut().insert(family.to_string(), value);
        value
    })
}

/// The highlight box for a run of this face and size: how far it reaches above
/// the baseline and below it, in points.
///
/// Measured on seven faces: the box is the font's DESIGN height, ascent plus
/// descent, with its bottom on the line box's bottom -- which is the 1.2 em
/// line split in the same ascent-to-descent proportion. Arial 18pt came back
/// 0.891 / 0.228 em against the rule's 0.8896 / 0.2276, Courier New 0.814 /
/// 0.319 against 0.8147 / 0.3181, Georgia 0.906 / 0.231 against 0.9050 /
/// 0.2312. Line spacing does not move it (150% and 70% arms are identical).
#[cfg(windows)]
fn highlight_extent_pt(family: &str, fs: f32) -> (f32, f32) {
    match hhea_extent_em(family) {
        Some((a, d)) if a + d > 0.0 => {
            let below = 1.2 * d / (a + d);
            ((a + d - below) * fs, below * fs)
        }
        // A face GDI cannot hand tables for keeps the line box itself, which
        // is the same box about 7% too tall at the top.
        _ => {
            let a = font_baseline_offset_em(family);
            (a * fs, (1.2 - a) * fs)
        }
    }
}

/// Colour emoji are drawn from their COLR layers unless this is set.
fn emoji_on() -> bool {
    std::env::var("OXI_EMOJI_DISABLE").is_err()
}

/// The face Windows resolves emoji to. PowerPoint's own PDF export draws them
/// in colour, and this is the only colour font on a stock Windows install.
const EMOJI_FAMILY: &str = "Segoe UI Emoji";

#[cfg(windows)]
thread_local! {
    static EMOJI_FONT: std::cell::RefCell<Option<Option<std::rc::Rc<emoji::ColorFont>>>> =
        const { std::cell::RefCell::new(None) };
    /// (family, weight, italic, char) -> does the family itself have the glyph?
    static GLYPH_CACHE: std::cell::RefCell<
        std::collections::HashMap<(String, i32, bool, char), bool>,
    > = std::cell::RefCell::new(std::collections::HashMap::new());
}

/// The parsed colour font, or None when this machine has no COLR emoji face.
#[cfg(windows)]
fn emoji_font() -> Option<std::rc::Rc<emoji::ColorFont>> {
    use windows::Win32::Graphics::Gdi::*;

    if !emoji_on() {
        return None;
    }
    EMOJI_FONT.with(|cell| {
        if let Some(hit) = cell.borrow().as_ref() {
            return hit.clone();
        }
        let dc = probe_dc();
        let wide: Vec<u16> = EMOJI_FAMILY
            .encode_utf16()
            .chain(std::iter::once(0))
            .collect();
        let value = unsafe {
            let font = CreateFontW(
                -ADVANCE_PROBE_EM,
                0,
                0,
                0,
                400,
                0,
                0,
                0,
                DEFAULT_CHARSET.0 as u32,
                OUT_DEFAULT_PRECIS.0 as u32,
                CLIP_DEFAULT_PRECIS.0 as u32,
                CLEARTYPE_QUALITY.0 as u32,
                (DEFAULT_PITCH.0 | FF_DONTCARE.0) as u32,
                windows::core::PCWSTR(wide.as_ptr()),
            );
            if font.is_invalid() {
                None
            } else {
                let old = SelectObject(dc, font);
                let parsed = (|| {
                    emoji::ColorFont::from_tables(
                        read_font_table(dc, b"COLR")?,
                        read_font_table(dc, b"CPAL")?,
                        read_font_table(dc, b"cmap")?,
                        read_font_table(dc, b"hmtx")?,
                        &read_font_table(dc, b"hhea")?,
                        &read_font_table(dc, b"head")?,
                    )
                })()
                .map(std::rc::Rc::new);
                SelectObject(dc, old);
                let _ = DeleteObject(font);
                parsed
            }
        };
        *cell.borrow_mut() = Some(value.clone());
        value
    })
}

/// True when `family` has its own glyph for `ch`.
///
/// A false here is what sends the character to the font-link chain, and it is
/// the discriminator for colour: PowerPoint paints an emoji in colour exactly
/// when the requested face could not supply it and Windows reached for Segoe
/// UI Emoji instead.
#[cfg(windows)]
fn family_has_glyph(family: &str, bold: bool, italic: bool, ch: char) -> bool {
    use windows::Win32::Graphics::Gdi::*;

    let weight = if bold { 700 } else { 400 };
    let key = (family.to_string(), weight, italic, ch);
    GLYPH_CACHE.with(|cache| {
        if let Some(hit) = cache.borrow().get(&key) {
            return *hit;
        }
        let dc = probe_dc();
        let (face, weight, italic) = styled_face(family, bold, italic);
        let wide: Vec<u16> = face.encode_utf16().chain(std::iter::once(0)).collect();
        let value = unsafe {
            let font = CreateFontW(
                -ADVANCE_PROBE_EM,
                0,
                0,
                0,
                weight,
                u32::from(italic),
                0,
                0,
                DEFAULT_CHARSET.0 as u32,
                OUT_DEFAULT_PRECIS.0 as u32,
                CLIP_DEFAULT_PRECIS.0 as u32,
                CLEARTYPE_QUALITY.0 as u32,
                (DEFAULT_PITCH.0 | FF_DONTCARE.0) as u32,
                windows::core::PCWSTR(wide.as_ptr()),
            );
            if font.is_invalid() {
                false
            } else {
                let old = SelectObject(dc, font);
                let mut buf = [0u16; 3];
                let n = ch.encode_utf16(&mut buf).len();
                buf[n] = 0;
                let mut idx = [0u16; 2];
                let got = GetGlyphIndicesW(
                    dc,
                    windows::core::PCWSTR(buf.as_ptr()),
                    n as i32,
                    idx.as_mut_ptr(),
                    GGI_MARK_NONEXISTING_GLYPHS,
                );
                SelectObject(dc, old);
                let _ = DeleteObject(font);
                got != GDI_ERROR as u32 && idx[..n].iter().all(|g| *g != 0xFFFF)
            }
        };
        cache.borrow_mut().insert(key, value);
        value
    })
}

/// The face that supplies an emoji's TEXT presentation. Measured: PowerPoint
/// advanced U+2764 by 0.9719 em and U+1F321 by 0.9989 em, and this font's own
/// hmtx holds 0.9717 and 1.0000 (`emojiadv` probe, 2026-08-19). The colour
/// arms all advanced 1.3709 em against Segoe UI Emoji's 1.3730.
const SYMBOL_FAMILY: &str = "Segoe UI Symbol";

#[cfg(windows)]
thread_local! {
    static SYMBOL_FONT: std::cell::RefCell<Option<Option<std::rc::Rc<emoji::ColorFont>>>> =
        const { std::cell::RefCell::new(None) };
}

/// The monochrome fallback face, read for its metrics only.
#[cfg(windows)]
fn symbol_font() -> Option<std::rc::Rc<emoji::ColorFont>> {
    use windows::Win32::Graphics::Gdi::*;

    SYMBOL_FONT.with(|cell| {
        if let Some(hit) = cell.borrow().as_ref() {
            return hit.clone();
        }
        let dc = probe_dc();
        let wide: Vec<u16> = SYMBOL_FAMILY
            .encode_utf16()
            .chain(std::iter::once(0))
            .collect();
        let value = unsafe {
            let font = CreateFontW(
                -ADVANCE_PROBE_EM,
                0,
                0,
                0,
                400,
                0,
                0,
                0,
                DEFAULT_CHARSET.0 as u32,
                OUT_DEFAULT_PRECIS.0 as u32,
                CLIP_DEFAULT_PRECIS.0 as u32,
                CLEARTYPE_QUALITY.0 as u32,
                (DEFAULT_PITCH.0 | FF_DONTCARE.0) as u32,
                windows::core::PCWSTR(wide.as_ptr()),
            );
            if font.is_invalid() {
                None
            } else {
                let old = SelectObject(dc, font);
                let parsed = (|| {
                    emoji::ColorFont::metrics_only(
                        read_font_table(dc, b"cmap")?,
                        read_font_table(dc, b"hmtx")?,
                        &read_font_table(dc, b"hhea")?,
                        &read_font_table(dc, b"head")?,
                    )
                })()
                .map(std::rc::Rc::new);
                SelectObject(dc, old);
                let _ = DeleteObject(font);
                parsed
            }
        };
        *cell.borrow_mut() = Some(value.clone());
        value
    })
}

/// What one character of a run needs.
#[cfg(windows)]
#[derive(Clone, Copy)]
enum CharPlan {
    /// The requested face draws it, at this advance in em.
    Base(f32),
    /// Colour emoji: the COLR layers of this glyph of the emoji face.
    Color(u16, f32),
    /// Text presentation: this glyph of the monochrome face, painted in the
    /// run's own colour.
    Symbol(u16, f32),
    /// A variation selector. It chooses the presentation and takes no width.
    Skip,
}

/// How to draw every character of `text`, or None when any of them needs
/// something this does not model -- in which case the caller keeps the plain
/// path, which is what every non-emoji run in the corpus takes.
///
/// The rule is PowerPoint's, measured (`emojipres` + `emojiadv`, 2026-08-19):
/// a face that owns the glyph draws it (Arial keeps its own U+263A even with
/// U+FE0F after it), otherwise Emoji_Presentation -- or the variation
/// selector overriding it -- picks between the colour face and the monochrome
/// one.
#[cfg(windows)]
fn run_plan(family: &str, bold: bool, italic: bool, text: &str) -> Option<Vec<CharPlan>> {
    let chars: Vec<char> = text.chars().collect();
    let mut out = Vec::with_capacity(chars.len());
    for (i, &ch) in chars.iter().enumerate() {
        if ch == emoji::VS16 || ch == emoji::VS15 {
            out.push(CharPlan::Skip);
            continue;
        }
        if family_has_glyph(family, bold, italic, ch) {
            out.push(CharPlan::Base(runtime_advance_em(family, bold, italic, ch)?));
            continue;
        }
        // Not an emoji at all -- a missing CJK glyph, say -- so this function
        // has nothing to offer and the run keeps GDI's font linking.
        let cf = emoji_font()?;
        let egid = cf.gid(ch)?;
        cf.layers(egid)?;
        let colored = match chars.get(i + 1) {
            Some(&n) if n == emoji::VS16 => true,
            Some(&n) if n == emoji::VS15 => false,
            _ => emoji::emoji_presentation(ch),
        };
        if colored {
            out.push(CharPlan::Color(egid, cf.advance_em(egid)?));
        } else {
            let sf = symbol_font()?;
            let sgid = sf.gid(ch)?;
            out.push(CharPlan::Symbol(sgid, sf.advance_em(sgid)?));
        }
    }
    Some(out)
}

/// One SFNT table of the font currently selected into `dc`.
#[cfg(windows)]
fn read_font_table(dc: windows::Win32::Graphics::Gdi::HDC, tag: &[u8; 4]) -> Option<Vec<u8>> {
    use windows::Win32::Graphics::Gdi::GetFontData;

    // GetFontData wants the tag as it appears in the file, read little-endian.
    let t = u32::from_le_bytes(*tag);
    unsafe {
        let n = GetFontData(dc, t, 0, None, 0);
        if n == 0 || n == u32::MAX {
            return None;
        }
        let mut buf = vec![0u8; n as usize];
        let got = GetFontData(
            dc,
            t,
            0,
            Some(buf.as_mut_ptr() as *mut core::ffi::c_void),
            n,
        );
        if got == u32::MAX { None } else { Some(buf) }
    }
}

#[cfg(not(windows))]
fn runtime_baseline_offset_em(_family: &str) -> Option<f32> {
    None
}

/// Bullet / AutoNum marker to draw at the start of a paragraph's first line
/// (Spec #8 / Spec #11). `text` is the marker string ("•" for buChar, or a
/// rendered auto-number like "1." / "(i)" for buAutoNum); `x_pt` is the
/// marker's left edge in pt relative to P0 (the shape's left text inset);
/// `baseline` is the slide-absolute baseline in pt (== line 0's).
struct MarkerInfo {
    text: String,
    font: String,
    x_pt: f32,
    baseline: f32,
    fs: f32,
}

/// Lay out one paragraph: advance `cursor_pt` (text-area top) by space_before,
/// wrap the run text, and return each line's (text, slide-absolute baseline in
/// pt, x-offset from the left inset in pt) plus an optional bullet marker.
/// Advances `cursor_pt` past the paragraph (incl. space_after).
///
/// Alignment model (Ra loop, spec5a/spec5b, PowerPoint PDF render-truth):
///   * Left      : every line starts at the left inset (x-offset 0).
///   * Center    : each line is centred on the text area: offset = (W - w)/2.
///   * Right     : each line ends at the right inset: offset = W - w.
///   * Justify   : non-final lines are stretched so the last word's right edge
///                 reaches the right inset (inter-word gaps are spread evenly,
///                 done at draw time); the FINAL line (and a 1-line paragraph)
///                 is left-aligned like Left.
/// `w` is the LOGICAL line width (GetTextExtentPoint32W advance sum, spaces
/// included); the ink bbox centre/right edge is offset by side bearings, so the
/// rule is anchored to the logical width (wave-1 finding).
///
/// Bullet / indent geometry (Spec #8, bulletph + bullet5 render-truth): the
/// paragraph's own pPr (marL / indent / bullet / space_before) wins over the
/// inherited master txStyles level. Master spcBef applies between paragraphs
/// only (never on the first line). See the inline comment for the measured
/// first-line / marker rule.
#[cfg(windows)]
fn layout_paragraph_baselines(
    dc: windows::Win32::Graphics::Gdi::HDC,
    para: &oxislides_core::ir::SlideParagraph,
    cursor_pt: &mut f32,
    shape_width: f32,
    scale: f64,
    is_first: bool,
    default_family: &str,
    l_ins: f32,
    r_ins: f32,
    master: &[MasterStyleLevel],
    ph_levels: &[MasterStyleLevel],
    anchor_off: f32,
    counters: &mut std::collections::HashMap<(u32, String), (Option<u32>, u32)>,
    prev_fs: &mut Option<f32>,
    // `a:bodyPr/@wrap` -- false lets the paragraph run past the box.
    wrap_text: bool,
) -> (Vec<(String, f32, f32)>, Option<MarkerInfo>) {
    use windows::Win32::Graphics::Gdi::*;
    // Master txStyles level for this paragraph's outline level (Spec #8).
    let mut m = if master.is_empty() {
        MasterStyleLevel::default()
    } else {
        let idx = (para.lvl as usize).min(master.len() - 1);
        master[idx].clone()
    };
    // The LAYOUT placeholder's own a:lstStyle overrides the master level, field
    // by field. Resolving it only in the draw loop and not here wrapped d24's
    // title at the master's 18pt and then drew it at the layout's 60pt, so the
    // line ran off the box instead of breaking into PowerPoint's three.
    if !ph_levels.is_empty() {
        let l = &ph_levels[(para.lvl as usize).min(ph_levels.len() - 1)];
        if l.font_size.is_some() {
            m.font_size = l.font_size;
        }
        if l.color.is_some() {
            m.color = l.color.clone();
        }
        if l.algn.is_some() {
            m.algn = l.algn;
        }
        if l.line_spacing.is_some() {
            m.line_spacing = l.line_spacing;
        }
        if l.bold.is_some() {
            m.bold = l.bold;
        }
    }
    // Effective font size: a run's explicit sz wins (the max over runs);
    // otherwise the master txStyles level default (Spec #5, phfs probe: V3
    // run 14pt overrides master 32pt); else the engine default. An EMPTY
    // paragraph is sized by its paragraph mark instead -- see
    // `paragraph_font_size`.
    let fs = paragraph_font_size(para, m.font_size, *prev_fs);
    *prev_fs = Some(fs);
    // The paragraph's own lnSpc wins; otherwise the placeholder chain's
    // (d24's master title placeholder says 90%, which is what makes its
    // 60pt title step 64.8pt instead of 72pt).
    let n = para.line_spacing.or(m.line_spacing).unwrap_or(1.0);
    let text: String = para.runs.iter().map(|r| r.text.as_str()).collect();
    // The same chain the draw path uses, so the wrap measures the face that
    // will actually be drawn: run, then the placeholder's own lstStyle (layout
    // then master), then whatever the caller resolved.
    let family = effective_family(
        dc,
        &para
            .runs
            .iter()
            .find_map(|r| r.font_family.clone())
            .or_else(|| {
                if !phfont_on() {
                    return None;
                }
                let lvl = para.lvl as usize;
                [ph_levels, master].into_iter().find_map(|levels| {
                    levels
                        .get(lvl.min(levels.len().saturating_sub(1)))
                        .and_then(|l| l.font_family.clone())
                })
            })
            .unwrap_or_else(|| default_family.to_string()),
    );
    let mar_l = para.mar_l.unwrap_or(m.mar_l);
    let indent = para.indent.unwrap_or(m.indent);
    let bullet = if matches!(para.bullet, SlideBullet::Inherit) {
        m.bullet
    } else {
        para.bullet.clone()
    };

    if let Some(sb) = para.space_before {
        *cursor_pt += sb;
    } else if !is_first {
        // Master spcBef (a:spcPct) — a fraction of the line advance, applied
        // between paragraphs only (the first paragraph's first line gets none).
        if let Some(pct) = m.spc_bef_pct {
            *cursor_pt += pct * fs * 1.2 * n;
        }
    }
    // Spec #6: vertical anchoring (a:bodyPr/@anchor resolved through the
    // placeholder chain). `anchor_off` shifts the first baseline of the whole
    // text block: anchor="ctr" centres the block in the inner area (offset
    // (inner_h - block_h)/2), anchor="b" pushes it to the bottom (inner_h -
    // block_h). It applies only to the FIRST paragraph's first line — the
    // cursor advance after the block must NOT re-apply it (it is baked into
    // `text_area_top`).
    let text_area_top = *cursor_pt + if is_first { anchor_off } else { 0.0 };
    let effective_width = (shape_width - l_ins - r_ins).max(0.0);
    // gdi_measure_text_px requires the target font to be selected into the DC;
    // otherwise the wrap measures with the DC default font and packs far too
    // many characters per line.
    let font = create_font_for(&family, fs, scale);
    let old_font = unsafe { SelectObject(dc, font) };
    // The WRAP has to measure at the weight the line is drawn at, level bold
    // included -- measuring d11's titles at 400 and drawing them at 700 would
    // break them a word later than PowerPoint.
    let bold = para.runs.iter().any(|r| r.bold) || (lvlbold_on() && m.bold.unwrap_or(false));
    // S-ITALADV (2026-08-24): and at the SLANT it is drawn at, level italic
    // included -- the same argument as the bold line above, which this one was
    // missing. The draw loop resolves italic as
    // `phl.is_some_and(|l| l.italic) || any run`, so measuring it as "any run"
    // alone measured d16's quotation UPRIGHT (Source Sans Pro, 464.62pt for the
    // line) and drew it in the deck's own italic part (446.06pt). The centred
    // line was then started from a width 4% too large -- every line about 20px
    // left of PowerPoint, even once the drawn advances were exact.
    let lvl_italic = italadv_on()
        && lvlitalic_on()
        && !ph_levels.is_empty()
        && ph_levels[(para.lvl as usize).min(ph_levels.len() - 1)].italic;
    let italic = para.runs.iter().any(|r| r.italic) || lvl_italic;
    let area_w = effective_width;
    let adv = fs * 1.2 * n;
    let first_off = first_baseline_off(&family, fs, n);
    // `cursor_pt` between paragraphs is the BOTTOM of the previous line box,
    // and every paragraph's first baseline sits its own ascent below it. When
    // the sizes match this is the old flat `adv` step exactly; when they differ
    // it is not, and PowerPoint's own export says the difference is real.
    // Probe `mixedpitch` (4 faces x 8 size pairs, 2026-08-18) fits
    //     step = d * prev_size + a * next_size,   a + d = 1.2004
    // with d = 0.2284 (Arial) / 0.2322 (Georgia) / 0.2636 (Calibri) / 0.2088
    // (Verdana) -- each within 0.0015 of that face's own
    // `1.2 * tmDescent / (tmAscent + tmDescent)`, i.e. the 1.2 line height
    // split by the FONT's ascent:descent ratio, which is what
    // `font_baseline_offset_em` already holds for the ascent half. d28's title
    // is 55pt then 66pt: PowerPoint steps 159px at 150dpi, the flat rule gives
    // 137px, and this gives 159.7px.
    //
    // The `lnSpc != 100%` arms fit the same way once the ascent switches to
    // 0.75 of the advance (which `first_off` already does): 10->40 at 150%
    // measures 58.470 / 56.310 / 43.350pt for both / next / prev against a
    // predicted 58.5 / 56.276 / 43.396.
    if mixpitch_on() || is_first {
        *cursor_pt += first_off;
    }

    // Bullet / indent geometry, all offsets relative to P0 (Spec #8, measured):
    //   para_left = P0 + marL;  indent > 0: text_1st = para_left + indent,
    //       marker = para_left
    //   indent <= 0: text_1st = max(para_left, P0 - indent),
    //       marker = text_1st + indent
    //   continuation lines = para_left;  render_x = max(text_1st, marker+bullet_w)
    let para_left_rel = mar_l;
    let (line0_x_off, marker_rel) = if indent > 0.0 {
        (mar_l + indent, mar_l)
    } else {
        let t = mar_l.max(-indent);
        (t, t + indent)
    };
    let mut line0_x_off = line0_x_off;
    let mut marker: Option<MarkerInfo> = None;
    match &bullet {
        SlideBullet::Char { ch, font } => {
            let marker_family = font.clone().unwrap_or_else(|| family.clone());
            let marker_w =
                font_adv::bullet_advance_em(&marker_family, *ch).unwrap_or(0.0) * fs;
            line0_x_off = line0_x_off.max(marker_rel + marker_w);
            marker = Some(MarkerInfo {
                text: ch.to_string(),
                font: marker_family,
                x_pt: marker_rel,
                baseline: 0.0, // line 0's baseline, filled in the line loop
                fs,
            });
        }
        SlideBullet::AutoNum { kind, start_at } => {
            // Spec #11: the auto-number counter is per (level, kind). The
            // sequence CONTINUES while startAt stays the same (absent==absent,
            // or the same value); it starts a NEW list whenever startAt
            // changes — present/absent or a different value — resetting to
            // startAt.unwrap_or(1). Word truth: autonum4 G [None][5][None]
            // renders 1,5,1 (the None list restarts after the [5] list);
            // autonum p1 L0 1,2,3..4 continues across interleaved levels
            // because the (lvl,kind) key never changed its startAt.
            let key = (para.lvl, kind.clone());
            let entry = counters.entry(key).or_insert((None, 0u32));
            let (last_start, c) = *entry;
            let n = if c == 0 || last_start != *start_at {
                start_at.unwrap_or(1) // new list (first use, or startAt changed)
            } else {
                c // same list: continue the sequence
            };
            *entry = (*start_at, n + 1);
            let text = autonum_text(kind, n);
            // The number is ASCII (digits / I V X / a-z A-Z) so its width is
            // the hmtx design-advance sum (unsupported families -> 0.0, the
            // same fallback buChar uses).
            let marker_family = family.clone();
            let marker_w =
                font_adv::line_hmtx_width_pt(&text, fs, &marker_family).unwrap_or(0.0);
            line0_x_off = line0_x_off.max(marker_rel + marker_w);
            marker = Some(MarkerInfo {
                text,
                font: marker_family,
                x_pt: marker_rel,
                baseline: 0.0, // line 0's baseline, filled in the line loop
                fs,
            });
        }
        _ => {}
    }

    // PowerPoint wraps each line against the same RIGHT EDGE, so the width a
    // line may use is the inner width MINUS that line's own left offset -- and
    // with a hanging indent or a bullet, line 0's offset differs from the rest.
    // Probe `wrapwidth` (2026-08-19, 7 arms): with marL18/ind-18 and no bullet,
    // line 0 starts at the inner left and runs a 232.07pt span while the
    // continuations start 18pt in and run at most 221.26 -- both stopping at
    // the same 316.8. Wrapping everything at the full 237.6 and shifting
    // afterwards let a continuation run 18pt past the inset, which is what put
    // an extra word on `bulletph`'s line 0.
    // `wrap="none"` means the text does not break at all: PowerPoint draws it
    // past the box edge. Every arm of the COM-built `embedsplit` probes is
    // such a box, auto-sized to 14.5pt around 20pt of text.
    let first_w = (effective_width - if wrapwidth_on() { line0_x_off } else { 0.0 }).max(1.0);
    let rest_w = (effective_width - if wrapwidth_on() { para_left_rel } else { 0.0 }).max(1.0);
    // `<a:br/>` arrives as a newline in the run stream and ends the line where
    // it stands. Each segment wraps on its own, and the newline is kept on the
    // END of the line it closed so the caller's character accounting -- which
    // maps lines back to runs -- still lines up.
    let lines = if !wrap_text && wrapnone_on() {
        vec![text.clone()]
    } else if softbreak_on() && text.contains('\n') {
        let mut out: Vec<String> = Vec::new();
        let mut seg_base = 0usize;
        for (si, seg) in text.split('\n').enumerate() {
            if si > 0 {
                match out.last_mut() {
                    Some(last) => last.push('\n'),
                    None => out.push("\n".to_string()),
                }
            }
            let w = if out.is_empty() { first_w } else { rest_w };
            let mut part = gdi_wrap_lines(
                dc, seg, w, rest_w, scale, fs, &family, bold, italic,
                Some((&para.runs[..], seg_base)),
            );
            seg_base += seg.chars().count() + 1;
            if part.is_empty() {
                part.push(String::new());
            }
            out.extend(part);
        }
        out
    } else {
        gdi_wrap_lines(
            dc, &text, first_w, rest_w, scale, fs, &family, bold, italic,
            Some((&para.runs[..], 0)),
        )
    };
    let n_lines = lines.len();

    // The size each LINE is set in. A soft break carries the size of the run it
    // stands in, so a paragraph can change size without ending -- and the step
    // across such a break is `descent(prev) + ascent(next)`, the same
    // mixed-pitch rule the paragraph boundary already uses, not the flat
    // paragraph advance.
    let line_sizes: Vec<f32> = {
        let mut sizes = Vec::with_capacity(n_lines);
        let mut at = 0usize;
        for line in &lines {
            let len = line.chars().count().max(1);
            let mut best: Option<f32> = None;
            let mut seen = 0usize;
            for run in &para.runs {
                let run_len = run.text.chars().count();
                if seen < at + len && at < seen + run_len {
                    if let Some(size) = run.font_size {
                        best = Some(best.map_or(size, |b: f32| b.max(size)));
                    }
                }
                seen += run_len;
            }
            sizes.push(best.unwrap_or(fs));
            at += line.chars().count();
        }
        sizes
    };
    let mixed_lines = brpitch_on() && line_sizes.iter().any(|s| (s - fs).abs() > 1e-4);
    // Only the mixed case leaves the flat arithmetic: an all-one-size paragraph
    // must keep `i * adv` exactly, down to the float association.
    let mixed_baselines: Vec<f32> = if mixed_lines {
        let ascent = |size: f32| first_baseline_off(&family, size, n);
        let mut baselines = Vec::with_capacity(n_lines);
        let mut at = text_area_top
            + if mixpitch_on() || is_first { ascent(line_sizes[0]) } else { 0.0 };
        for (i, size) in line_sizes.iter().enumerate() {
            if i > 0 {
                let prev = line_sizes[i - 1];
                at += (1.2 * prev * n - ascent(prev)) + ascent(*size);
            }
            baselines.push(at);
        }
        baselines
    } else {
        Vec::new()
    };

    let mut out = Vec::with_capacity(n_lines);
    let mut align_at = 0usize; // char offset of this line within the paragraph
    for (i, line) in lines.iter().enumerate() {
        let baseline = if mixed_lines {
            mixed_baselines[i]
        } else {
            text_area_top
                + if mixpitch_on() || is_first { first_off } else { 0.0 }
                + i as f32 * adv
        };
        // Logical line width in pt = hmtx design-advance sum of the VISIBLE
        // characters (trailing spaces excluded; final visible char included).
        // GDI's measured width (hinted / pixel-snapped) over-measures a line by
        // ~1.5-3.75pt vs PowerPoint, so we prefer the hmtx table and fall back
        // to the GDI measurement only for unsupported fonts/characters.
        // A line that ends on a soft break carries the newline for the
        // caller's accounting; it is not ink and must not be measured.
        let ink = line.trim_end_matches('\n');
        // S-RUNALIGN: measure the line the way it is DRAWN -- each run at its
        // own size, weight and slant -- exactly as the wrap already does.
        let per_run = if runalign_on() && para.runs.len() > 1 {
            line_width_pt_runs(
                dc,
                ink.trim_end(),
                fs,
                &family,
                bold,
                italic,
                scale,
                RunStyles { runs: &para.runs, line_start: align_at },
            )
        } else {
            None
        };
        align_at += line.chars().count();
        let line_w = per_run
            .or_else(|| hmtx_width_styled(ink, fs, &family, bold, italic))
            .or_else(|| {
                runtime_width_px(dc, ink.trim_end(), fs, &family, bold, italic, scale)
                    .map(|px| px as f32 / scale as f32)
            })
            .unwrap_or_else(|| gdi_measure_text_px(dc, line) as f32 / scale as f32);
        // Spec #6: horizontal alignment resolution — a paragraph with no
        // explicit alignment inherits the master txStyles level's algn (then
        // the default Left). The run level carries no alignment; the chain is
        // paragraph -> master txStyles level.
        let align = para
            .alignment
            .unwrap_or(m.algn.unwrap_or(SlideAlignment::Left));
        let is_justify_last = matches!(align, SlideAlignment::Justify) && i + 1 == n_lines;
        let align_off = match align {
            SlideAlignment::Center => (area_w - line_w).max(0.0) / 2.0,
            SlideAlignment::Right => (area_w - line_w).max(0.0),
            SlideAlignment::Justify if is_justify_last => 0.0,
            _ => 0.0,
        };
        if i == 0 {
            if let Some(mk) = marker.as_mut() {
                mk.baseline = baseline;
            }
        }
        let base_off = if i == 0 { line0_x_off } else { para_left_rel };
        out.push((line.clone(), baseline, base_off + align_off));
    }
    let _ = unsafe { SelectObject(dc, old_font) };
    // Leaving the paragraph, the cursor stops at the last line box's bottom --
    // one descent below the last baseline, which is `adv - first_off`.
    *cursor_pt = if mixed_lines {
        let last = *line_sizes.last().unwrap_or(&fs);
        mixed_baselines[n_lines - 1] + (1.2 * last * n - first_baseline_off(&family, last, n))
    } else {
        text_area_top
            + if !mixpitch_on() && is_first { first_off } else { 0.0 }
            + n_lines as f32 * adv
    };
    if let Some(sa) = para.space_after {
        *cursor_pt += sa;
    }
    (out, marker)
}

/// Render an auto-number marker string for a buAutoNum `kind` at count `n`
/// (Spec #11, autonum4.pdf 16-scheme render-truth, all PowerPoint-valid kinds).
/// The body is decimal, uppercase/lowercase roman, or uppercase/lowercase
/// alpha (Excel-column style: 1=A .. 26=Z .. 27=AA); the wrapper is
/// `<body>.` (Period), `<body>)` (ParenR), `(<body>)` (ParenBoth), or the
/// bare body (Plain). Unknown kinds fall back to arabicPeriod ("N.") — the
/// *Plain roman/alpha kinds crash PowerPoint and never appear in real docs.
fn autonum_text(kind: &str, n: u32) -> String {
    let body = if kind.starts_with("romanUc") {
        to_roman(n)
    } else if kind.starts_with("romanLc") {
        to_roman(n).to_lowercase()
    } else if kind.starts_with("alphaUc") {
        to_alpha(n)
    } else if kind.starts_with("alphaLc") {
        to_alpha(n).to_lowercase()
    } else {
        n.to_string()
    };
    if kind.ends_with("ParenBoth") {
        format!("({})", body)
    } else if kind.ends_with("ParenR") {
        format!("{})", body)
    } else if kind.ends_with("Period") {
        format!("{}.", body)
    } else {
        body // Plain and any unknown kind
    }
}

/// Standard greedy decimal -> uppercase Roman numeral (1=I .. 3999=MMMCMXCIX).
fn to_roman(mut n: u32) -> String {
    const ROMAN: [(u32, &str); 13] = [
        (1000, "M"),
        (900, "CM"),
        (500, "D"),
        (400, "CD"),
        (100, "C"),
        (90, "XC"),
        (50, "L"),
        (40, "XL"),
        (10, "X"),
        (9, "IX"),
        (5, "V"),
        (4, "IV"),
        (1, "I"),
    ];
    let mut s = String::new();
    for &(v, r) in ROMAN.iter() {
        while n >= v {
            s.push_str(r);
            n -= v;
        }
    }
    s
}

/// Excel-column style alphabetic numbering: 1=A, 2=B, .. 26=Z, 27=AA.
fn to_alpha(mut n: u32) -> String {
    let mut s = String::new();
    while n > 0 {
        let d = ((n - 1) % 26) as u8;
        s.insert(0, (b'A' + d) as char);
        n = (n - 1) / 26;
    }
    s
}

/// Draw text at a baseline position (converts baseline -> cell top via tmAscent).
#[cfg(windows)]
fn draw_text_baseline_w(
    dc: windows::Win32::Graphics::Gdi::HDC,
    x: i32,
    baseline_pt: f32,
    text: &str,
    font_size: f32,
    family: &str,
    color: Option<&str>,
    scale: f64,
    weight: i32,
) {
    draw_text_baseline_wi(dc, x, baseline_pt, text, font_size, family, color, scale, weight, false)
}

/// As `draw_text_baseline_w`, with the slant.
#[cfg(windows)]
fn draw_text_baseline_wi(
    dc: windows::Win32::Graphics::Gdi::HDC,
    x: i32,
    baseline_pt: f32,
    text: &str,
    font_size: f32,
    family: &str,
    color: Option<&str>,
    scale: f64,
    weight: i32,
    italic: bool,
) {
    draw_text_baseline_wiu(
        dc, x, baseline_pt, text, font_size, family, color, scale, weight, italic, false,
    )
}

/// Draw a line that contains colour emoji, one character at a time.
///
/// Returns false when the line has no colour character, so the caller keeps
/// its single `ExtTextOutW`. GDI cannot paint a COLR glyph -- it draws the
/// base outline, which is why Oxi's emoji were black line art against
/// PowerPoint's colour ones -- so each emoji is painted here as its own stack
/// of layer glyphs, back to front, at one pen position.
#[cfg(windows)]
#[allow(clippy::too_many_arguments)]
fn draw_color_run(
    dc: windows::Win32::Graphics::Gdi::HDC,
    x: i32,
    baseline_px: i32,
    text: &str,
    font_size: f32,
    family: &str,
    scale: f64,
    weight: i32,
    italic: bool,
    underline: bool,
) -> bool {
    use windows::Win32::Foundation::*;
    use windows::Win32::Graphics::Gdi::*;
    use windows::core::PCWSTR;

    let bold = weight >= 700;
    let cf = match emoji_font() {
        Some(cf) => cf,
        None => return false,
    };
    let plan = match run_plan(family, bold, italic, text) {
        Some(p) => p,
        None => return false,
    };
    // Every character is the requested face's own: nothing to do here.
    if plan.iter().all(|p| matches!(p, CharPlan::Base(_))) {
        return false;
    }
    let dx = match runtime_dx_px(dc, text, font_size, family, bold, italic, scale) {
        Some(dx) if dx.len() == plan.len() => dx,
        // Without agreeing advances the characters would not land where the
        // wrap measured them; the plain path at least keeps them adjacent.
        _ => return false,
    };

    // The emoji face has its own ascent, so its glyphs sit on the line's
    // baseline rather than on the base font's cell top.
    let efont = create_font_for_wiu(EMOJI_FAMILY, font_size, 400, false, false, scale);
    if efont.is_invalid() {
        return false;
    }
    let old_font = unsafe { SelectObject(dc, efont) };
    let mut tm = TEXTMETRICW::default();
    unsafe {
        let _ = GetTextMetricsW(dc, &mut tm);
    }
    let emoji_y = baseline_px - tm.tmAscent;
    unsafe {
        SelectObject(dc, old_font);
    }

    // The monochrome face needs the same treatment as the colour one: its own
    // handle, and its own ascent.
    let sfont = if plan.iter().any(|p| matches!(p, CharPlan::Symbol(..))) {
        create_font_for_wiu(SYMBOL_FAMILY, font_size, weight, italic, underline, scale)
    } else {
        windows::Win32::Graphics::Gdi::HFONT::default()
    };
    let symbol_y = if sfont.is_invalid() {
        emoji_y
    } else {
        let mut stm = TEXTMETRICW::default();
        unsafe {
            let old = SelectObject(dc, sfont);
            let _ = GetTextMetricsW(dc, &mut stm);
            SelectObject(dc, old);
        }
        baseline_px - stm.tmAscent
    };

    // Ordinary characters in the same line keep the base font, and its own
    // ascent -- the two faces do not share one.
    let base = create_font_for_wiu(family, font_size, weight, italic, underline, scale);
    let base_y = if base.is_invalid() {
        emoji_y
    } else {
        let mut btm = TEXTMETRICW::default();
        unsafe {
            let old = SelectObject(dc, base);
            let _ = GetTextMetricsW(dc, &mut btm);
            SelectObject(dc, old);
        }
        baseline_px - btm.tmAscent
    };
    let mut text_color = unsafe { GetTextColor(dc) };
    let mut pen = x;
    let mut i = 0usize;
    for (ci, ch) in text.chars().enumerate() {
        match plan[ci] {
            CharPlan::Color(gid, _) => {
                let layers = cf.layers(gid).unwrap_or_default();
                unsafe {
                    let old = SelectObject(dc, efont);
                    for (lg, pi) in layers {
                        let paint = if pi == 0xFFFF {
                            None
                        } else {
                            match cf.color(pi) {
                                // A fully transparent layer paints nothing;
                                // GDI text has no alpha to honour a partial
                                // one, so it is drawn opaque.
                                Some((_, _, _, 0)) => continue,
                                Some((r, g, b, _)) => Some(colorref(r, g, b)),
                                None => None,
                            }
                        };
                        if let Some(c) = paint {
                            text_color = SetTextColor(dc, COLORREF(c));
                        }
                        let one = [lg];
                        let _ = ExtTextOutW(
                            dc,
                            pen,
                            emoji_y,
                            ETO_GLYPH_INDEX,
                            None,
                            PCWSTR(one.as_ptr()),
                            1,
                            None,
                        );
                        if paint.is_some() {
                            text_color = SetTextColor(dc, COLORREF(text_color.0));
                        }
                    }
                    SelectObject(dc, old);
                }
            }
            // The monochrome presentation, painted in the run's own colour --
            // which is why d11 slide 38's U+2764 is the deck's dark navy in
            // PowerPoint and not Segoe UI Emoji's red.
            CharPlan::Symbol(gid, _) if !sfont.is_invalid() => unsafe {
                let old = SelectObject(dc, sfont);
                let one = [gid];
                let _ = ExtTextOutW(
                    dc,
                    pen,
                    symbol_y,
                    ETO_GLYPH_INDEX,
                    None,
                    PCWSTR(one.as_ptr()),
                    1,
                    None,
                );
                SelectObject(dc, old);
            },
            CharPlan::Base(_) if !base.is_invalid() => {
                let mut buf = [0u16; 2];
                let wch = ch.encode_utf16(&mut buf);
                unsafe {
                    let old = SelectObject(dc, base);
                    let _ = ExtTextOutW(
                        dc,
                        pen,
                        base_y,
                        ETO_OPTIONS(0),
                        None,
                        PCWSTR(wch.as_ptr()),
                        wch.len() as u32,
                        None,
                    );
                    SelectObject(dc, old);
                }
            }
            _ => {}
        }
        pen += dx[i];
        i += 1;
    }
    unsafe {
        let _ = DeleteObject(efont);
        if !base.is_invalid() {
            let _ = DeleteObject(base);
        }
        if !sfont.is_invalid() {
            let _ = DeleteObject(sfont);
        }
    }
    true
}

/// As `draw_text_baseline_wi`, with the underline.
#[cfg(windows)]
#[allow(clippy::too_many_arguments)]
fn draw_text_baseline_wiu(
    dc: windows::Win32::Graphics::Gdi::HDC,
    x: i32,
    baseline_pt: f32,
    text: &str,
    font_size: f32,
    family: &str,
    color: Option<&str>,
    scale: f64,
    weight: i32,
    italic: bool,
    underline: bool,
) {
    use windows::Win32::Foundation::*;
    use windows::Win32::Graphics::Gdi::*;
    use windows::core::PCWSTR;
    let font = create_font_for_wiu(family, font_size, weight, italic, underline, scale);
    if font.is_invalid() {
        return;
    }
    let rgb = color.and_then(parse_hex_rgb).unwrap_or((0, 0, 0));
    let old_color = unsafe { SetTextColor(dc, COLORREF(colorref(rgb.0, rgb.1, rgb.2))) };
    let old_font = unsafe { SelectObject(dc, font) };
    let mut tm = TEXTMETRICW::default();
    unsafe {
        let _ = GetTextMetricsW(dc, &mut tm);
    }
    let ascent_px = tm.tmAscent as i32;
    let y = (baseline_pt as f64 * scale).round() as i32 - ascent_px;
    // A line carrying colour emoji is painted character by character, since
    // one ExtTextOutW can only use one font and cannot paint COLR layers.
    if emoji_on()
        && draw_color_run(
            dc,
            x,
            (baseline_pt as f64 * scale).round() as i32,
            text,
            font_size,
            family,
            scale,
            weight,
            italic,
            underline,
        )
    {
        unsafe {
            SelectObject(dc, old_font);
            SetTextColor(dc, old_color);
            let _ = DeleteObject(font);
        }
        return;
    }
    let wtext: Vec<u16> = text.encode_utf16().collect();
    // When the family has an hmtx table, draw each char at its design
    // advance (Dx) so glyphs land exactly where PowerPoint's PDF export
    // places them. Otherwise fall back to the hinted GDI text.
    // Draw at the design advances so the glyphs land where the wrap measured
    // them (and where PowerPoint puts them). The measured `font_adv` tables win
    // when they cover the family; otherwise the advances come from the font GDI
    // resolved, which is what makes embedded faces work.
    // S-ITALADV (2026-08-24): the glyphs were drawn in the ITALIC face while
    // their positions came from the UPRIGHT one -- `runtime_dx_px` was handed a
    // hardcoded `false`, and `line_hmtx_dx_px` takes no style at all.
    //
    // d16 slide 5's quotation is Source Sans Pro Italic at 36pt, and the deck
    // embeds that exact part. Asking GDI for the string's width:
    //     "Source Sans Pro"    italic=1  ->  464.62pt   (a synthesised oblique,
    //                                        i.e. the UPRIGHT advances)
    //     "Source Sans Pro #I"           ->  446.06pt   (the deck's own part)
    //     PowerPoint's own span          ->  446.12pt
    // The right face was already being selected to DRAW with -- 0.06pt from
    // PowerPoint -- but the dx array came from the upright measurement, so every
    // line came out **+3.3%** wide, and a centred line therefore started left of
    // where PowerPoint starts it. The error is a clean scale, not accumulated
    // rounding: the last word of the line sits at 428.16pt against PowerPoint's
    // 414.72pt, a ratio of 1.032 that matches the ink ratio 1.033.
    //
    // The hmtx table is skipped outright for italic because it holds no italic
    // data (Arial / Arial Bold / Calibri only); `runtime_dx_px` then measures
    // the face that is actually drawn, through the same `styled_face` the font
    // selection uses. For Arial this changes nothing -- Arial Italic carries the
    // same advances as Arial -- so only families with a genuinely narrower
    // italic move.
    let dx = if (italic && italadv_on()) || (weight >= 700 && hmtxstyle_on()) {
        None
    } else {
        font_adv::line_hmtx_dx_px(text, font_size, family, scale)
    }
    .or_else(|| {
        let it = italic && italadv_on();
        runtime_dx_px(dc, text, font_size, family, weight >= 700, it, scale)
    });
    if let Some(dx) = dx {
        unsafe {
            let _ = ExtTextOutW(
                dc,
                x,
                y,
                ETO_OPTIONS(0),
                None,
                PCWSTR(wtext.as_ptr()),
                wtext.len() as u32,
                Some(dx.as_ptr()),
            );
        }
    } else {
        unsafe {
            let _ = TextOutW(dc, x, y, &wtext);
        }
    }
    unsafe {
        SelectObject(dc, old_font);
        SetTextColor(dc, old_color);
        let _ = DeleteObject(font);
    }
}

/// Regular-weight text at a baseline position.
#[cfg(windows)]
fn draw_text_baseline(
    dc: windows::Win32::Graphics::Gdi::HDC,
    x: i32,
    baseline_pt: f32,
    text: &str,
    font_size: f32,
    family: &str,
    color: Option<&str>,
    scale: f64,
) {
    draw_text_baseline_w(dc, x, baseline_pt, text, font_size, family, color, scale, 400)
}

/// Draw a justified (non-final) line: split into words, then spread the
/// stretch evenly over the inter-word gaps so the last word's right edge
/// reaches `right_x` (the right text inset). The measured model: PowerPoint
/// justify stretches the gaps between words; the final line of a paragraph is
/// left-aligned (handled by the caller via the last-line flag).
#[cfg(windows)]
fn draw_text_justify(
    dc: windows::Win32::Graphics::Gdi::HDC,
    left_x: i32,
    right_x: i32,
    baseline_pt: f32,
    text: &str,
    font_size: f32,
    family: &str,
    color: Option<&str>,
    scale: f64,
) {
    use windows::Win32::Foundation::*;
    use windows::Win32::Graphics::Gdi::*;
    use windows::core::PCWSTR;
    let font = create_font_for(family, font_size, scale);
    if font.is_invalid() {
        return;
    }
    let rgb = color.and_then(parse_hex_rgb).unwrap_or((0, 0, 0));
    let old_color = unsafe { SetTextColor(dc, COLORREF(colorref(rgb.0, rgb.1, rgb.2))) };
    let old_font = unsafe { SelectObject(dc, font) };
    let mut tm = TEXTMETRICW::default();
    unsafe {
        let _ = GetTextMetricsW(dc, &mut tm);
    }
    let ascent_px = tm.tmAscent as i32;
    let y = (baseline_pt as f64 * scale).round() as i32 - ascent_px;

    // Split into words (ignore leading/trailing spaces).
    let words: Vec<&str> = text.split(' ').filter(|w| !w.is_empty()).collect();
    if words.len() <= 1 {
        let wtext: Vec<u16> = text.encode_utf16().collect();
        unsafe {
            let _ = TextOutW(dc, left_x, y, &wtext);
        }
        unsafe {
            SelectObject(dc, old_font);
            SetTextColor(dc, old_color);
            let _ = DeleteObject(font);
        }
        return;
    }

    // When the family has an hmtx table, both the natural width (from which
    // the justify stretch is computed) and the word/gap placement use the
    // design advance, matching PowerPoint's PDF export. Otherwise fall back
    // to the hinted GDI metrics.
    let hmtx = font_adv::line_hmtx_dx_px(text, font_size, family, scale).is_some();
    let (space_w, word_ws): (i32, Vec<i32>) = if hmtx {
        let sw = font_adv::space_hmtx_px(font_size, family, scale).unwrap_or(0);
        let ww = words
            .iter()
            .map(|w| font_adv::text_hmtx_px(w, font_size, family, scale).unwrap_or(0))
            .collect();
        (sw, ww)
    } else {
        let sw = gdi_measure_text_px(dc, " ");
        let ww = words.iter().map(|w| gdi_measure_text_px(dc, w)).collect();
        (sw, ww)
    };
    let natural: i32 = word_ws.iter().sum::<i32>() + space_w * (words.len() - 1) as i32;
    let stretch = (right_x - left_x - natural).max(0);
    let n_gaps = (words.len() - 1) as i32;
    let per_gap = if n_gaps > 0 { stretch / n_gaps } else { 0 };
    let rem = if n_gaps > 0 { stretch % n_gaps } else { 0 };

    let mut x = left_x;
    for (i, word) in words.iter().enumerate() {
        let wtext: Vec<u16> = word.encode_utf16().collect();
        if hmtx {
            if let Some(dx) = font_adv::line_hmtx_dx_px(word, font_size, family, scale) {
                unsafe {
                    let _ = ExtTextOutW(
                        dc,
                        x,
                        y,
                        ETO_OPTIONS(0),
                        None,
                        PCWSTR(wtext.as_ptr()),
                        wtext.len() as u32,
                        Some(dx.as_ptr()),
                    );
                }
            } else {
                unsafe {
                    let _ = TextOutW(dc, x, y, &wtext);
                }
            }
        } else {
            unsafe {
                let _ = TextOutW(dc, x, y, &wtext);
            }
        }
        x += word_ws[i];
        if i + 1 < words.len() {
            // Spread the integer remainder over the first `rem` gaps.
            let extra = per_gap + if (i as i32) < rem { 1 } else { 0 };
            x += space_w + extra;
        }
    }

    unsafe {
        SelectObject(dc, old_font);
        SetTextColor(dc, old_color);
        let _ = DeleteObject(font);
    }
}
