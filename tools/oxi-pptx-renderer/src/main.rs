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
//! Default dpi=150, supersample=2. Produces `<prefix>_s1.png`, `<prefix>_s2.png`
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

mod font_adv;

use oxislides_core::ir::{
    MasterStyleLevel, Presentation, Shape, ShapeContent, SlideAlignment, SlideBullet,
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
    let mut supersample: u32 = 2;
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
                                    let mut cursor_pt = sh.y + sh.t_ins;
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
                                                sh.l_ins,
                                                sh.r_ins,
                                                &master_ctx[..],
                                                anchor_off,
                                                &mut counters,
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
/// When the block is taller than the inner area the offset clamps to 0
/// (top-aligned; conservative — overflow behaviour not yet measured).
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
    let inner_h = (sh.height - sh.t_ins - sh.b_ins).max(0.0);
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
    for (i, para) in paragraphs.iter().enumerate() {
        let _ = layout_paragraph_baselines(
            dc,
            para,
            &mut cursor_pt,
            sh.width,
            scale,
            i == 0,
            &def_family,
            sh.l_ins,
            sh.r_ins,
            &master_ctx[..],
            0.0,
            &mut counters,
        );
    }
    // block_h = the block's advance minus the first paragraph's first_off.
    let para = &paragraphs[0];
    let fs = para
        .runs
        .iter()
        .filter_map(|r| r.font_size)
        .fold(None, |acc: Option<f32>, x| Some(acc.map_or(x, |a| a.max(x))))
        .unwrap_or(master_ctx.first().and_then(|m| m.font_size).unwrap_or(18.0));
    let first_off = {
        let n = para.line_spacing.unwrap_or(1.0);
        let family = para
            .runs
            .iter()
            .find_map(|r| r.font_family.clone())
            .unwrap_or_else(|| def_family.clone());
        if (n - 1.0).abs() > 1e-4 {
            0.75 * fs * 1.2 * n
        } else {
            font_baseline_offset_em(&family) * fs
        }
    };
    let block_h = (cursor_pt - first_off).max(0.0);
    let extra = (inner_h - block_h).max(0.0);
    if anchor == Some("ctr") {
        extra / 2.0
    } else {
        extra
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
                            .map(|cell| paragraphs_json(&cell.paragraphs))
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
                "series": chart
                    .series
                    .iter()
                    .map(|s| json!({ "name": s.name, "values": s.values }))
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
        "type": shape_type,
        "fill_color": sh.fill_color,
        "border_color": sh.border_color,
        "border_width": sh.border_width,
        "anchor": sh.anchor,
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

fn colorref(r: u8, g: u8, b: u8) -> u32 {
    (r as u32) | ((g as u32) << 8) | ((b as u32) << 16)
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

            // Background (slide background color if set, else white)
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
            SetBkMode(mem_dc, TRANSPARENT);

            for sh in &slide.shapes {
                let x = (sh.x as f64 * scale).round() as i32;
                let y = (sh.y as f64 * scale).round() as i32;
                let ew = (sh.width as f64 * scale).round() as i32;
                let eh = (sh.height as f64 * scale).round() as i32;

                // Fill
                if let Some(fill) = &sh.fill_color {
                    if let Some((r, g, b)) = parse_hex_rgb(fill) {
                        let brush = CreateSolidBrush(COLORREF(colorref(r, g, b)));
                        let r2 = RECT {
                            left: x,
                            top: y,
                            right: x + ew,
                            bottom: y + eh,
                        };
                        FillRect(mem_dc, &r2, brush);
                        let _ = DeleteObject(brush);
                    }
                }

                // Border
                let border_w = sh.border_width.unwrap_or(0.0);
                if border_w > 0.0 {
                    let col = sh
                        .border_color
                        .as_deref()
                        .and_then(parse_hex_rgb)
                        .unwrap_or((0, 0, 0));
                    let pen = CreatePen(
                        PS_SOLID,
                        (border_w as f64 * scale).round() as i32,
                        COLORREF(colorref(col.0, col.1, col.2)),
                    );
                    let old_pen = SelectObject(mem_dc, pen);
                    let _ = SelectObject(mem_dc, GetStockObject(NULL_BRUSH));
                    let _ = Rectangle(mem_dc, x, y, x + ew, y + eh);
                    SelectObject(mem_dc, old_pen);
                    let _ = DeleteObject(pen);
                }

                // Text (Spec #4 layout: wrap at word boundaries within the
                // effective width, place each line at its baseline). AutoShapes
                // with a text body render their text too.
                match &sh.content {
                    ShapeContent::TextBox { paragraphs }
                    | ShapeContent::AutoShape { paragraphs } => {
                        let left_x = x + (sh.l_ins as f64 * scale).round() as i32;
                        let right_x = x
                            + ((sh.width - sh.r_ins) as f64 * scale).round() as i32;
                        let mut cursor_pt = sh.y + sh.t_ins;
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
                            let m_fs = m.and_then(|mm| mm.font_size);
                            let fs = p
                                .runs
                                .iter()
                                .filter_map(|r| r.font_size)
                                .fold(None, |acc: Option<f32>, x| {
                                    Some(acc.map_or(x, |a| a.max(x)))
                                })
                                .unwrap_or(m_fs.unwrap_or(18.0));
                            let family = p
                                .runs
                                .iter()
                                .find_map(|r| r.font_family.clone())
                                .unwrap_or_else(|| resolve_font(pres, sh));
                            let color = p.runs.iter().find_map(|r| r.color.clone());
                            let (lines, marker) = layout_paragraph_baselines(
                                mem_dc,
                                p,
                                &mut cursor_pt,
                                sh.width,
                                scale,
                                pi == 0,
                                &family,
                                sh.l_ins,
                                sh.r_ins,
                                &master_ctx[..],
                                anchor_off,
                                &mut counters,
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
                            for (i, (line_text, baseline, x_off)) in
                                lines.into_iter().enumerate()
                            {
                                if line_text.trim().is_empty() {
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
                                    draw_text_baseline(
                                        mem_dc,
                                        line_x,
                                        baseline,
                                        &line_text,
                                        fs,
                                        &family,
                                        color.as_deref(),
                                        scale,
                                    );
                                }
                            }
                        }
                    }
                    ShapeContent::Table { table } => {
                        // Grid lines: one rectangle per cell (borrow the
                        // black border pen already set above when border_w>0;
                        // use a dedicated pen here so a borderless table still
                        // shows its grid).
                        let pen = CreatePen(
                            PS_SOLID,
                            (1.0 * scale).round() as i32,
                            COLORREF(colorref(0, 0, 0)),
                        );
                        let old_pen = SelectObject(mem_dc, pen);
                        let _ = SelectObject(mem_dc, GetStockObject(NULL_BRUSH));
                        let mut cy = y;
                        for (r, row) in table.rows.iter().enumerate() {
                            let ph = table
                                .row_heights
                                .get(r)
                                .copied()
                                .unwrap_or(0.0) as f64
                                * scale;
                            let ph = ph.round() as i32;
                            let mut cx = x;
                            for (c, cell) in row.iter().enumerate() {
                                let pw = table
                                    .col_widths
                                    .get(c)
                                    .copied()
                                    .unwrap_or(0.0) as f64
                                    * scale;
                                let pw = pw.round() as i32;
                                let _ = Rectangle(mem_dc, cx, cy, cx + pw, cy + ph);
                                // Cell text (top-left, minimal inset)
                                let mut cursor_y = cy + (0.06 * scale).round() as i32;
                                for p in &cell.paragraphs {
                                    let fs = p
                                        .runs
                                        .iter()
                                        .filter_map(|r| r.font_size)
                                        .fold(18.0, f32::max);
                                    let text: String =
                                        p.runs.iter().map(|r| r.text.as_str()).collect();
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
                            }
                            cy += ph;
                        }
                        SelectObject(mem_dc, old_pen);
                        let _ = DeleteObject(pen);
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
                        let max_axis = if is_100pct {
                            100.0
                        } else {
                            nice_axis_max(raw_max)
                        };
                        let axis_steps = if is_100pct {
                            10usize
                        } else if is_stacked {
                            ((max_axis / 5.0).round() as usize).max(1)
                        } else {
                            5usize
                        };
                        let axis_label = |i: usize| -> String {
                            let v = max_axis * i as f64 / axis_steps as f64;
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
                        let plot_top = if has_explicit_title {
                            sy + 45.69
                        } else if has_auto_title {
                            sy + 51.40
                        } else {
                            sy + 16.0
                        };
                        let plot_bot = sy + shh - 39.9;
                        let plot_h = plot_bot - plot_top;

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
                        let val_y = |v: f64| {
                            if max_axis > 0.0 {
                                plot_bot - (v / max_axis) * plot_h
                            } else {
                                plot_bot
                            }
                        };

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
                            for i in 1..=axis_steps {
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
                            let pt_i = (plot_top * scale).round() as i32;
                            let _ = MoveToEx(mem_dc, pl, pt_i, None);
                            let _ = LineTo(mem_dc, pl, pb);
                            let _ = MoveToEx(mem_dc, pl, pb, None);
                            let _ = LineTo(mem_dc, pr, pb);
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
                                let _ = MoveToEx(mem_dc, tx, pb, None);
                                let _ = LineTo(
                                    mem_dc,
                                    tx,
                                    ((plot_bot + 5.71) * scale).round() as i32,
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
                                (plot_bot + 28.67) as f32,
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
                        let plot_left = sx + 41.4;
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
                        let plot_w = if chart.has_legend {
                            sw - 41.4 - 103.82
                        } else {
                            sw - 41.4 - 11.0
                        };
                        let plot_right = plot_left + plot_w;
                        let plot_w = plot_right - plot_left;
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
                        let plot_bot = if crowded {
                            sy + shh - 78.62
                        } else {
                            sy + shh - 39.9
                        };
                        let plot_h = plot_bot - plot_top;
                        let pitch = plot_w / n_cat as f64;
                        let max_val = chart
                            .series
                            .iter()
                            .flat_map(|s| s.values.iter().copied())
                            .fold(0.0f64, f64::max);
                        let max_axis = nice_axis_max(max_val);
                        let axis_steps = 5usize;

                        // Value axis labels (Calibri 18pt, right edge =
                        // plot_left-16.64, baseline = tick_y+5.22; same
                        // rule as the bars).
                        for i in 0..=axis_steps {
                            let val = max_axis * i as f64 / axis_steps as f64;
                            let tick_y = plot_bot - (val / max_axis) * plot_h;
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
                                        let y = if max_axis > 0.0 {
                                            plot_bot - (v / max_axis) * plot_h
                                        } else {
                                            plot_bot
                                        };
                                        (x, y)
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
                                    format!("{}", v.round() as i64)
                                }
                            };
                            for (si, pts) in series_pts.iter().enumerate() {
                                for (pt, s) in pts.iter().zip(
                                    chart.series.get(si).map(|s| &s.values).into_iter().flatten(),
                                ) {
                                    let text = format_label(*s);
                                    let lw = font_adv::line_hmtx_width_pt(
                                        &text,
                                        axis_fs,
                                        axis_family,
                                    )
                                    .unwrap_or_else(|| {
                                        text.chars().count() as f32 * axis_fs * 0.5
                                    }) as f64;
                                    let lx = pt.0 - lw / 2.0;
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
                                            (plot_bot + 45.0) as f32,
                                            &a,
                                            18.0,
                                            axis_family,
                                            None,
                                            scale,
                                        );
                                        draw_text_baseline(
                                            mem_dc,
                                            ((cat_center - wb / 2.0) * scale).round() as i32,
                                            (plot_bot + 67.55) as f32,
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
                                (plot_bot + single_bl) as f32,
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
                        for i in 0..=n_cat {
                            let tick_x = plot_left + pitch * i as f64;
                            let tx = (tick_x * scale).round() as i32;
                            let _ = MoveToEx(mem_dc, tx, pb, None);
                            let _ = LineTo(mem_dc, tx, ((plot_bot + 5.7) * scale).round() as i32);
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
                        let plot_left = sx + 6.50 + cat_label_w + 16.70;
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
                        let mut axis_max = nice_axis_max(raw_max);
                        let mut axis_steps = 5usize;
                        let mut plot_right = band_right - 9.13;
                        for _ in 0..2 {
                            let (m, d) =
                                horiz_value_axis(raw_max, plot_right - plot_left);
                            axis_max = m;
                            axis_steps = d.max(1);
                            let step = m / axis_steps as f64;
                            plot_right = band_right
                                - text_w(&fmt_axis_value(m, step), axis_fs) / 2.0;
                        }
                        let plot_w = plot_right - plot_left;
                        let pitch = plot_h / n_cat as f64;
                        let axis_step = axis_max / axis_steps as f64;

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
                                let mut cum = 0.0f64;
                                for (si, series) in chart.series.iter().enumerate() {
                                    let v =
                                        series.values.get(ci).copied().unwrap_or(0.0);
                                    if v <= 0.0 {
                                        continue;
                                    }
                                    let seg_w = if axis_max > 0.0 {
                                        v / axis_max * plot_w
                                    } else {
                                        0.0
                                    };
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
                                        let old_brush = SelectObject(mem_dc, brush);
                                        let r = RECT {
                                            left: ((plot_left + cum) * scale).round()
                                                as i32,
                                            top: (by0 * scale).round() as i32,
                                            right: ((plot_left + cum + seg_w) * scale)
                                                .round()
                                                as i32,
                                            bottom: ((by0 + bar_h) * scale).round()
                                                as i32,
                                        };
                                        let _ = FillRect(mem_dc, &r, brush);
                                        SelectObject(mem_dc, old_brush);
                                        let _ = DeleteObject(brush);
                                    }
                                    cum += seg_w;
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
                                    if let Some(rgb) =
                                        col_hex.and_then(parse_hex_rgb)
                                    {
                                        let brush = CreateSolidBrush(COLORREF(
                                            colorref(rgb.0, rgb.1, rgb.2),
                                        ));
                                        let old_brush = SelectObject(mem_dc, brush);
                                        // series 0 is the BOTTOM bar of the cluster
                                        let by0 = cat_center + cluster_h / 2.0
                                            - (si as f64 + 1.0) * bar_h;
                                        let bw = if axis_max > 0.0 {
                                            v / axis_max * plot_w
                                        } else {
                                            0.0
                                        };
                                        let r = RECT {
                                            left: (plot_left * scale).round() as i32,
                                            top: (by0 * scale).round() as i32,
                                            right: ((plot_left + bw) * scale).round()
                                                as i32,
                                            bottom: ((by0 + bar_h) * scale).round()
                                                as i32,
                                        };
                                        let _ = FillRect(mem_dc, &r, brush);
                                        SelectObject(mem_dc, old_brush);
                                        let _ = DeleteObject(brush);
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
                            let v = axis_step * i as f64;
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
                        for (ci, cat) in chart.categories.iter().enumerate() {
                            let cat_center = plot_bot - pitch * (ci as f64 + 0.5);
                            let lw = text_w(cat, axis_fs);
                            draw_text_baseline(
                                mem_dc,
                                ((plot_left - 16.70 - lw) * scale).round() as i32,
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
                            let _ = MoveToEx(mem_dc, pl, pt, None);
                            let _ = LineTo(mem_dc, pl, pb);
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
                                    ((plot_left - 5.72) * scale).round() as i32,
                                    ty,
                                    None,
                                );
                                let _ = LineTo(mem_dc, pl, ty);
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
                        let plot_left = if is_100pct { sx + 63.44 } else { sx + 41.4 };
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
                        let plot_right = sx + sw - 11.0;
                        let plot_bot = sy + shh - 39.9;
                        let plot_w = plot_right - plot_left;
                        let plot_h = plot_bot - plot_top;
                        let n_cat = chart.categories.len().max(1);
                        let n_ser = chart.series.len().max(1);
                        let axis_fs = 18.0f32;
                        let axis_family = "Calibri";

                        // Value axis labels (0..max_axis in even steps),
                        // right-aligned to a fixed gutter. For a CLUSTERED
                        // chart the scale is 0..max_axis in 5 steps (6
                        // labels). For a STACKED chart Word scales to the
                        // largest per-category series SUM (chart_stacked:
                        // Q2 sum 36.4 -> nice max 40) and draws one label
                        // per 5-step tick, i.e. (max_axis/5)+1 labels
                        // (0,5,...,40 = 9 labels, render-truth 2026-08-06).
                        let max_val = if is_100pct {
                            // percentStacked: fixed 0..100 scale (10-step %-axis).
                            100.0
                        } else if is_stacked {
                            // largest per-category sum over all series
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
                        // percentStacked's axis is FIXED at 100, so it must not
                        // go through nice_axis_max (whose 5% headroom would
                        // round 100 up to 120).
                        let max_axis = if is_100pct {
                            100.0
                        } else {
                            nice_axis_max(max_val)
                        };
                        let axis_steps = if is_100pct {
                            // percentStacked: 0%,10%,...,100% (11 labels).
                            10usize
                        } else if is_stacked {
                            (max_axis / 5.0).round().max(1.0) as usize
                        } else {
                            5usize
                        };
                        for i in 0..=axis_steps {
                            // percentStacked labels the fixed 0..100 scale
                            // as "0%".."100%" (render-truth: '0%' x=96.77 /
                            // '100%' x=78.50, right-aligned, baseline
                            // tick_y + 5.2).
                            let val = if is_100pct {
                                100.0 * i as f64 / axis_steps as f64
                            } else {
                                max_axis * i as f64 / axis_steps as f64
                            };
                            let tick_y = plot_bot - (val / max_axis) * plot_h;
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
                                let mut cum_h = 0.0;
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
                                    } else if max_axis > 0.0 {
                                        v / max_axis * plot_h
                                    } else {
                                        0.0
                                    };
                                    let by1 = plot_bot - cum_h;
                                    let by0 = by1 - seg_h;
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
                                    }
                                    cum_h += seg_h;
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
                                    let brush = CreateSolidBrush(COLORREF(colorref(
                                        rgb.0, rgb.1, rgb.2,
                                    )));
                                    let old_brush = SelectObject(mem_dc, brush);
                                    let cat_center =
                                        plot_left + pitch * (ci as f64 + 0.5);
                                    let bx0 = cat_center
                                        - cluster_w / 2.0
                                        + si as f64 * bar_w;
                                    let bar_h = if max_axis > 0.0 {
                                        v / max_axis * plot_h
                                    } else {
                                        0.0
                                    };
                                    let by0 = plot_bot - bar_h;
                                    let r = RECT {
                                        left: (bx0 * scale).round() as i32,
                                        top: (by0 * scale).round() as i32,
                                        right: ((bx0 + bar_w) * scale).round() as i32,
                                        bottom: (plot_bot * scale).round() as i32,
                                    };
                                    let _ = FillRect(mem_dc, &r, brush);
                                    SelectObject(mem_dc, old_brush);
                                    let _ = DeleteObject(brush);
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
                                    format!("{}", v.round() as i64)
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
                                    let mut cum_h = 0.0f64;
                                    for s in chart.series.iter() {
                                        let v = s
                                            .values
                                            .get(ci)
                                            .copied()
                                            .unwrap_or(0.0);
                                        if v <= 0.0 {
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
                                        } else if max_axis > 0.0 {
                                            v / max_axis * plot_h
                                        } else {
                                            0.0
                                        };
                                        let by1 = plot_bot - cum_h;
                                        let by0 = by1 - seg_h;
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
                                        cum_h += seg_h;
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
                                        if v <= 0.0 {
                                            continue;
                                        }
                                        let bx0 = cat_center
                                            - cluster_w / 2.0
                                            + si as f64 * bar_w;
                                        let bar_center = bx0 + bar_w / 2.0;
                                        let bar_h = if max_axis > 0.0 {
                                            v / max_axis * plot_h
                                        } else {
                                            0.0
                                        };
                                        let by0 = plot_bot - bar_h;
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
                                                by0 + bar_h / 2.0 + 6.2
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
                            draw_text_baseline(
                                mem_dc,
                                (lx * scale).round() as i32,
                                (plot_bot + 28.67) as f32,
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
                        for i in 0..=n_cat {
                            let tick_x = plot_left + pitch * i as f64;
                            let tx = (tick_x * scale).round() as i32;
                            let _ = MoveToEx(mem_dc, tx, pb, None);
                            let _ = LineTo(mem_dc, tx, ((plot_bot + 5.7) * scale).round() as i32);
                        }

                        SelectObject(mem_dc, old_axis_pen);
                        let _ = DeleteObject(axis_pen);
                        }
                    }
                    ShapeContent::Image { data, content_type: _ } => {
                        // Decode the embedded image (PNG/JPEG media part) and
                        // draw it scaled into the shape rect. rotation=0 only
                        // for now (a rotation-aware path is left for a later
                        // step; deck background images are typically unrotated).
                        if let Ok(dyn_img) = image::load_from_memory(data) {
                            let rgba = dyn_img.to_rgba8();
                            let (iw, ih) = (rgba.width() as i32, rgba.height() as i32);
                            if iw <= 0 || ih <= 0 {
                                continue;
                            }
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
                            let _ = StretchDIBits(
                                mem_dc,
                                x,
                                y,
                                ew,
                                eh,
                                0,
                                0,
                                iw,
                                ih,
                                Some(bgra.as_ptr() as *const _),
                                &bmi,
                                DIB_RGB_COLORS,
                                SRCCOPY,
                            );
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
    use windows::Win32::Graphics::Gdi::*;
    use windows::core::PCWSTR;
    let height = (font_size as f64 * scale).round() as i32;
    let wide: Vec<u16> = family.encode_utf16().collect();
    let mut family_buf = vec![0u16; wide.len() + 1];
    family_buf[..wide.len()].copy_from_slice(&wide);
    unsafe {
        CreateFontW(
            -height, 0, 0, 0, weight, 0, 0, 0, 1, 0, 0, 5, 0,
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
fn horiz_value_axis(max_val: f64, plot_w: f64) -> (f64, usize) {
    if max_val <= 0.0 || plot_w <= 0.0 {
        return (1.0, 1);
    }
    const MIN_SPACING: f64 = 57.0;
    const HEADROOM: f64 = 1.05;
    let mut best: Option<(f64, f64, usize)> = None; // (step, axis_max, div)
    for k in -6i32..=9 {
        let mag = 10f64.powi(k);
        for m in [1.0f64, 2.0, 5.0] {
            let step = m * mag;
            let axis_max = (max_val * HEADROOM / step - 1e-9).ceil() * step;
            if !axis_max.is_finite() || axis_max <= 0.0 {
                continue;
            }
            let div = (axis_max / step).round() as usize;
            if div == 0 || div > 200 {
                continue;
            }
            if plot_w / div as f64 >= MIN_SPACING {
                match best {
                    Some((bs, _, _)) if bs <= step => {}
                    _ => best = Some((step, axis_max, div)),
                }
            }
        }
    }
    match best {
        Some((_, axis_max, div)) => (axis_max, div),
        // Plot too narrow for any nice step: fall back to a single division.
        None => (nice_axis_max(max_val), 1),
    }
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

/// Wrap `text` at word boundaries to fit `effective_width_pt`.
#[cfg(windows)]
fn gdi_wrap_lines(
    dc: windows::Win32::Graphics::Gdi::HDC,
    text: &str,
    effective_width_pt: f32,
    scale: f64,
) -> Vec<String> {
    let width_px = (effective_width_pt as f64 * scale).round().max(1.0) as i32;
    let mut lines: Vec<String> = Vec::new();
    let mut current = String::new();
    let mut current_w = 0i32;
    for word in text.split_inclusive(' ') {
        let w = gdi_measure_text_px(dc, word);
        if !current.is_empty() && current_w + w > width_px {
            lines.push(std::mem::take(&mut current));
            current_w = 0;
        }
        current.push_str(word);
        current_w += w;
    }
    if !current.is_empty() {
        lines.push(current);
    }
    if lines.is_empty() {
        lines.push(String::new());
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
fn font_baseline_offset_em(family: &str) -> f32 {
    match family.to_ascii_lowercase().as_str() {
        "arial" => 0.97274,
        "times new roman" => 0.96587,
        "calibri" => 0.93648,
        "segoe ui" => 0.97399,
        "georgia" => 0.96899,
        "verdana" => 0.99275,
        _ => 0.9685, // avg of the six measured fonts (model1b: 1.2*(win_asc+0.5)/win_total)
    }
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
    anchor_off: f32,
    counters: &mut std::collections::HashMap<(u32, String), (Option<u32>, u32)>,
) -> (Vec<(String, f32, f32)>, Option<MarkerInfo>) {
    use windows::Win32::Graphics::Gdi::*;
    // Master txStyles level for this paragraph's outline level (Spec #8).
    let m = if master.is_empty() {
        MasterStyleLevel::default()
    } else {
        let idx = (para.lvl as usize).min(master.len() - 1);
        master[idx].clone()
    };
    // Effective font size: a run's explicit sz wins (the max over runs);
    // otherwise the master txStyles level default (Spec #5, phfs probe: V3
    // run 14pt overrides master 32pt); else the engine default.
    let fs = para
        .runs
        .iter()
        .filter_map(|r| r.font_size)
        .fold(None, |acc: Option<f32>, x| Some(acc.map_or(x, |a| a.max(x))))
        .unwrap_or(m.font_size.unwrap_or(18.0));
    let n = para.line_spacing.unwrap_or(1.0);
    let text: String = para.runs.iter().map(|r| r.text.as_str()).collect();
    let family = para
        .runs
        .iter()
        .find_map(|r| r.font_family.clone())
        .unwrap_or_else(|| default_family.to_string());
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
    let lines = gdi_wrap_lines(dc, &text, effective_width, scale);
    let area_w = effective_width;
    let n_lines = lines.len();
    let adv = fs * 1.2 * n;
    let first_off = if (n - 1.0).abs() > 1e-4 {
        0.75 * adv
    } else {
        font_baseline_offset_em(&family) * fs
    };
    // The baseline offset (ascent-based first-line placement) applies ONLY to
    // the text area's FIRST line. Between paragraphs the line grid continues
    // at the plain `adv` pitch (Word render-truth: paragraph gap == one line
    // height when space_after == 0). So for paragraphs after the first,
    // `cursor_pt` already sits at the first-line baseline and we must NOT add
    // `first_off` again (that double-count was a +16.89pt gap per paragraph).
    if is_first {
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

    let mut out = Vec::with_capacity(n_lines);
    for (i, line) in lines.iter().enumerate() {
        let baseline = text_area_top + if is_first { first_off } else { 0.0 } + i as f32 * adv;
        // Logical line width in pt = hmtx design-advance sum of the VISIBLE
        // characters (trailing spaces excluded; final visible char included).
        // GDI's measured width (hinted / pixel-snapped) over-measures a line by
        // ~1.5-3.75pt vs PowerPoint, so we prefer the hmtx table and fall back
        // to the GDI measurement only for unsupported fonts/characters.
        let line_w = font_adv::line_hmtx_width_pt(line, fs, &family)
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
    *cursor_pt = text_area_top + if is_first { first_off } else { 0.0 } + n_lines as f32 * adv;
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
    use windows::Win32::Foundation::*;
    use windows::Win32::Graphics::Gdi::*;
    use windows::core::PCWSTR;
    let font = create_font_for_w(family, font_size, weight, scale);
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
    let wtext: Vec<u16> = text.encode_utf16().collect();
    // When the family has an hmtx table, draw each char at its design
    // advance (Dx) so glyphs land exactly where PowerPoint's PDF export
    // places them. Otherwise fall back to the hinted GDI text.
    if let Some(dx) = font_adv::line_hmtx_dx_px(text, font_size, family, scale) {
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
