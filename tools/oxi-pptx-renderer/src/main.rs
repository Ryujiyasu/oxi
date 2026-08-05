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
//! Text rendering is deliberately minimal (one line per paragraph, no wrapping
//! yet) — paragraph/line layout within a shape is the first spec target of the
//! Ra loop and will replace this scaffold.

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
        let json = dump_layout_json(&pres);
        let text = serde_json::to_string_pretty(&json).expect("Cannot serialize layout");
        std::fs::write(&path, text).expect("Cannot write layout JSON");
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
fn dump_layout_json(pres: &Presentation) -> Value {
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
    json!({
        "presentation": {
            "width": pres.slide_width,
            "height": pres.slide_height,
        },
        "slides": slides,
    })
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

                // Text (one line per paragraph, no wrapping yet — the first
                // Ra-loop spec target will replace this scaffold). AutoShapes
                // with a text body render their text too.
                match &sh.content {
                    ShapeContent::TextBox { paragraphs }
                    | ShapeContent::AutoShape { paragraphs } => {
                        let mut cursor_y = y;
                        for p in paragraphs {
                            let fs = p
                                .runs
                                .iter()
                                .filter_map(|r| r.font_size)
                                .fold(18.0, f32::max);
                            let text: String = p.runs.iter().map(|r| r.text.as_str()).collect();
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
                            draw_text_line(mem_dc, x, cursor_y, &text, fs, &family, color.as_deref(), scale);
                            cursor_y += (fs as f64 * scale * 1.2).round() as i32;
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
