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

use oxislides_core::ir::{Presentation, Shape, ShapeContent, SlideAlignment};
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
                                    let mut cursor_pt = sh.y + MARGIN_TOP;
                                    for (i, para_json) in arr.iter_mut().enumerate() {
                                        if let Some(para) = sh_para(&sh.content, i) {
                                            let bases = layout_paragraph_baselines(
                                                dc,
                                                para,
                                                &mut cursor_pt,
                                                sh.width,
                                                scale,
                                                i == 0,
                                            );
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

fn alignment_str(a: SlideAlignment) -> &'static str {
    match a {
        SlideAlignment::Left => "left",
        SlideAlignment::Center => "center",
        SlideAlignment::Right => "right",
        SlideAlignment::Justify => "justify",
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
                        let left_x = x + (MARGIN_LEFT as f64 * scale).round() as i32;
                        let right_x = x
                            + ((sh.width - MARGIN_RIGHT) as f64 * scale).round() as i32;
                        let mut cursor_pt = sh.y + MARGIN_TOP;
                        for (pi, p) in paragraphs.iter().enumerate() {
                            let fs = p
                                .runs
                                .iter()
                                .filter_map(|r| r.font_size)
                                .fold(18.0, f32::max);
                            let family = p
                                .runs
                                .iter()
                                .find_map(|r| r.font_family.clone())
                                .unwrap_or_else(|| "Calibri".to_string());
                            let color = p.runs.iter().find_map(|r| r.color.clone());
                            let lines = layout_paragraph_baselines(
                                mem_dc,
                                p,
                                &mut cursor_pt,
                                sh.width,
                                scale,
                                pi == 0,
                            );
                            let is_justify = matches!(
                                p.alignment,
                                oxislides_core::ir::SlideAlignment::Justify
                            );
                            let n_lines = lines.len();
                            for (i, (line_text, baseline, x_off)) in lines.into_iter().enumerate()
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
                                        .unwrap_or_else(|| "Calibri".to_string());
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
//   * inner insets                         = top/bottom 3.6pt, left/right 7.2pt
// ---------------------------------------------------------------------------

#[cfg(windows)]
const MARGIN_TOP: f32 = 3.6;
#[cfg(windows)]
const MARGIN_LEFT: f32 = 7.2;
#[cfg(windows)]
const MARGIN_RIGHT: f32 = 7.2;

/// Create a GDI font for the given family/size (negative lfHeight = char height).
#[cfg(windows)]
fn create_font_for(
    family: &str,
    font_size: f32,
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
            -height, 0, 0, 0, 400, 0, 0, 0, 1, 0, 0, 5, 0,
            PCWSTR(family_buf.as_ptr()),
        )
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

/// A_font = hhea_asc + hhea_lineGap (fontTools-measured), the first-line
/// baseline offset factor for a single-spaced (n == 1) paragraph.
#[cfg(windows)]
fn font_baseline_offset_em(family: &str) -> f32 {
    match family.to_ascii_lowercase().as_str() {
        "calibri" => 0.9707,
        "arial" => 0.9380,
        "times new roman" => 0.9336,
        _ => 0.9380, // Arial-like default
    }
}

/// Lay out one paragraph: advance `cursor_pt` (text-area top) by space_before,
/// wrap the run text, and return each line's (text, slide-absolute baseline in
/// pt, x-offset from the left inset in pt). Advances `cursor_pt` past the
/// paragraph (incl. space_after).
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
#[cfg(windows)]
fn layout_paragraph_baselines(
    dc: windows::Win32::Graphics::Gdi::HDC,
    para: &oxislides_core::ir::SlideParagraph,
    cursor_pt: &mut f32,
    shape_width: f32,
    scale: f64,
    is_first: bool,
) -> Vec<(String, f32, f32)> {
    use windows::Win32::Graphics::Gdi::*;
    let fs = para
        .runs
        .iter()
        .filter_map(|r| r.font_size)
        .fold(18.0, f32::max);
    let n = para.line_spacing.unwrap_or(1.0);
    let text: String = para.runs.iter().map(|r| r.text.as_str()).collect();
    let family = para
        .runs
        .iter()
        .find_map(|r| r.font_family.clone())
        .unwrap_or_else(|| "Calibri".to_string());

    if let Some(sb) = para.space_before {
        *cursor_pt += sb;
    }
    let text_area_top = *cursor_pt;
    let effective_width = (shape_width - MARGIN_LEFT - MARGIN_RIGHT).max(0.0);
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
        let is_justify_last =
            matches!(para.alignment, oxislides_core::ir::SlideAlignment::Justify)
                && i + 1 == n_lines;
        let x_off = match para.alignment {
            oxislides_core::ir::SlideAlignment::Center => (area_w - line_w).max(0.0) / 2.0,
            oxislides_core::ir::SlideAlignment::Right => (area_w - line_w).max(0.0),
            oxislides_core::ir::SlideAlignment::Justify if is_justify_last => 0.0,
            _ => 0.0,
        };
        out.push((line.clone(), baseline, x_off));
    }
    let _ = unsafe { SelectObject(dc, old_font) };
    *cursor_pt = text_area_top + if is_first { first_off } else { 0.0 } + n_lines as f32 * adv;
    if let Some(sa) = para.space_after {
        *cursor_pt += sa;
    }
    out
}

/// Draw text at a baseline position (converts baseline -> cell top via tmAscent).
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
