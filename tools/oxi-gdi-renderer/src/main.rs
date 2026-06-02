//! Oxi GDI Renderer: Render .docx layout using Windows GDI for pixel-accurate comparison with Word.
//!
//! Usage: oxi-gdi-renderer input.docx output_prefix
//!   Produces: output_prefix_p1.png, output_prefix_p2.png, ...

use std::path::Path;

fn main() {
    let args: Vec<String> = std::env::args().collect();
    if args.len() < 3 {
        eprintln!("Usage: {} <input.docx> <output_prefix> [dpi] [--exclude=text,border,shading,box,image,clip] [--supersample=N]", args[0]);
        std::process::exit(1);
    }

    let docx_path = &args[1];
    let output_prefix = &args[2];
    let dpi: u32 = args.get(3).and_then(|s| s.parse().ok()).unwrap_or(150);

    // Parse --exclude flag and --supersample=N flag from any argument
    let mut exclude: Vec<String> = Vec::new();
    // 2x supersampling is the default as of 2026-04-16: CJK fonts ≤10.5pt hit the GDI
    // no-AA fallback with CLEARTYPE_QUALITY, but supersampling + Lanczos downscale
    // restores grayscale AA matching Word EMF output.
    let mut supersample: u32 = 2;
    let mut dump_layout: Option<String> = None;
    for arg in &args[3..] {
        if let Some(list) = arg.strip_prefix("--exclude=") {
            exclude = list.split(',').map(|s| s.trim().to_lowercase()).collect();
        }
        if let Some(n) = arg.strip_prefix("--supersample=") {
            supersample = n.parse().unwrap_or(2);
        }
        if let Some(path) = arg.strip_prefix("--dump-layout=") {
            dump_layout = Some(path.to_string());
        }
    }

    // Parse document
    let data = std::fs::read(docx_path).expect("Cannot read docx file");
    let doc = oxidocs_core::parser::parse_docx(&data).expect("Cannot parse docx");

    // Layout. S483: render in Word's "final" view (hide <w:del>, show <w:ins>)
    // to match the final/clean-view Word ground-truth PNGs. Opt-out OXI_S483_DISABLE.
    let engine = oxidocs_core::layout::LayoutEngine::for_document(&doc);
    let engine = if std::env::var("OXI_S483_DISABLE").is_ok() {
        engine
    } else {
        engine.with_show_revisions(oxidocs_core::ir::ShowRevisions::Final)
    };
    let result = engine.layout(&doc);

    eprintln!("Parsed {} pages, DPI={} supersample={}x", result.pages.len(), dpi, supersample);
    if !exclude.is_empty() {
        eprintln!("Excluding: {:?}", exclude);
    }

    if let Some(ref path) = dump_layout {
        dump_layout_json(&result, path);
        eprintln!("Layout dumped to {}", path);
        return;
    }

    // Render each page with GDI
    #[cfg(windows)]
    {
        render_pages_gdi(&result, output_prefix, dpi, supersample, &exclude);
    }

    #[cfg(not(windows))]
    {
        eprintln!("GDI rendering requires Windows");
        std::process::exit(1);
    }
}

#[cfg(windows)]
/// Parse a `#RRGGBB` hex color into (r, g, b) bytes. Defaults to (0, 0, 0)
/// for malformed input — used by R-05g balloon + connector renderers.
fn parse_hex_rgb(s: &str) -> (u8, u8, u8) {
    let c = s.strip_prefix('#').unwrap_or(s);
    if c.len() != 6 {
        return (0, 0, 0);
    }
    let r = u8::from_str_radix(&c[0..2], 16).unwrap_or(0);
    let g = u8::from_str_radix(&c[2..4], 16).unwrap_or(0);
    let b = u8::from_str_radix(&c[4..6], 16).unwrap_or(0);
    (r, g, b)
}

fn render_pages_gdi(result: &oxidocs_core::layout::LayoutResult, prefix: &str, dpi: u32, supersample: u32, exclude: &[String]) {
    use windows::Win32::Graphics::Gdi::*;
    use windows::Win32::Foundation::*;
    use windows::core::*;

    // Supersample: render at render_dpi internally, then downscale to output dpi.
    let render_dpi = dpi * supersample.max(1);
    let scale = render_dpi as f64 / 72.0;

    for (page_idx, page) in result.pages.iter().enumerate() {
        let out_w = (page.width as f64 * dpi as f64 / 72.0).round() as u32;
        let out_h = (page.height as f64 * dpi as f64 / 72.0).round() as u32;
        let w = (page.width as f64 * scale).round() as i32;
        let h = (page.height as f64 * scale).round() as i32;

        unsafe {
            // Create memory DC and bitmap
            let screen_dc = GetDC(HWND(std::ptr::null_mut()));
            let mem_dc = CreateCompatibleDC(screen_dc);
            let bitmap = CreateCompatibleBitmap(screen_dc, w, h);
            let old_bmp = SelectObject(mem_dc, bitmap);

            // Fill white background
            let white_brush = CreateSolidBrush(COLORREF(0x00FFFFFF));
            let rect = RECT { left: 0, top: 0, right: w, bottom: h };
            FillRect(mem_dc, &rect, white_brush);
            let _ = DeleteObject(white_brush);

            // Set text mode
            SetBkMode(mem_dc, TRANSPARENT);

            // Render each element (skip excluded types)
            for elem in &page.elements {
                let type_name = match &elem.content {
                    oxidocs_core::layout::LayoutContent::Text { .. } => "text",
                    oxidocs_core::layout::LayoutContent::TableBorder { .. } => "border",
                    oxidocs_core::layout::LayoutContent::CellShading { .. } => "shading",
                    oxidocs_core::layout::LayoutContent::BoxRect { .. } => "box",
                    oxidocs_core::layout::LayoutContent::Image { .. } => "image",
                    oxidocs_core::layout::LayoutContent::ClipStart => "clip",
                    oxidocs_core::layout::LayoutContent::ClipEnd => "clip",
                    oxidocs_core::layout::LayoutContent::PresetShape { .. } => "shape",
                    // R-05a: balloon variants are layout-only stubs for now;
                    // the GDI renderer will gain real handlers in R-05g.
                    oxidocs_core::layout::LayoutContent::Balloon { .. } => "balloon",
                    oxidocs_core::layout::LayoutContent::BalloonConnector { .. } => "balloon_connector",
                };
                if exclude.iter().any(|e| e == type_name) {
                    continue;
                }

                let x = (elem.x as f64 * scale).round() as i32;
                let y = (elem.y as f64 * scale).round() as i32;
                let ew = (elem.width as f64 * scale).round() as i32;
                let eh = (elem.height as f64 * scale).round() as i32;

                match &elem.content {
                    oxidocs_core::layout::LayoutContent::Text {
                        text, font_size, font_family, bold, italic, color, underline, underline_style, strikethrough, double_strikethrough, highlight, character_spacing, text_scale, is_vertical, ..
                    } => {
                        // Session 75 Phase D (2026-05-17): elem.y is LINE BOX TOP;
                        // glyph_y = LBT + text_y_off is where TextOutW/underline/
                        // strikethrough/highlight should draw to preserve pre-Phase-D
                        // pixel positions. See memory/session71_y_convention_refactor_design.md.
                        let text_y_off_px = (elem.text_y_off as f64 * scale).round() as i32;
                        let glyph_y = y + text_y_off_px;
                        let fs = (*font_size as f64 * scale).round() as i32;
                        // 2026-04-19: Apply horizontal text scale (OOXML w:w).
                        // CreateFontW lfWidth=0 → default aspect; positive value → specified glyph width.
                        let lf_width = if (*text_scale - 100.0).abs() > 0.01 {
                            (*font_size as f64 * (*text_scale as f64 / 100.0) * scale * 0.5).round() as i32
                        } else { 0 };
                        let family = font_family.as_deref().unwrap_or("Calibri");

                        // Parse color
                        let rgb = color.as_deref()
                            .and_then(|c| {
                                let c = c.strip_prefix('#').unwrap_or(c);
                                if c.len() == 6 {
                                    let r = u8::from_str_radix(&c[0..2], 16).ok()?;
                                    let g = u8::from_str_radix(&c[2..4], 16).ok()?;
                                    let b = u8::from_str_radix(&c[4..6], 16).ok()?;
                                    Some(COLORREF((r as u32) | ((g as u32) << 8) | ((b as u32) << 16)))
                                } else { None }
                            })
                            .unwrap_or(COLORREF(0x00000000));

                        // Draw highlight background before text
                        if let Some(ref hl) = highlight {
                            let hl_rgb = {
                                let c = hl.strip_prefix('#').unwrap_or(hl);
                                if c.len() == 6 {
                                    let r = u8::from_str_radix(&c[0..2], 16).unwrap_or(255);
                                    let g = u8::from_str_radix(&c[2..4], 16).unwrap_or(255);
                                    let b = u8::from_str_radix(&c[4..6], 16).unwrap_or(0);
                                    COLORREF((r as u32) | ((g as u32) << 8) | ((b as u32) << 16))
                                } else {
                                    COLORREF(0x0000FFFF) // default yellow
                                }
                            };
                            let hl_brush = CreateSolidBrush(hl_rgb);
                            // Phase D: highlight tracks glyph (not line box) to preserve pre-Phase-D pixels.
                            let r = RECT { left: x, top: glyph_y, right: x + ew, bottom: glyph_y + eh };
                            FillRect(mem_dc, &r, hl_brush);
                            let _ = DeleteObject(hl_brush);
                        }

                        SetTextColor(mem_dc, rgb);

                        // Create font
                        // Session 132 (2026-05-20): vertical writing — set
                        // lfEscapement = lfOrientation = -900 (tenths of CCW
                        // degrees) for 90° CW rotation. tbRlV cells flow top-
                        // to-bottom with each character rotated 90° CW. GDI
                        // takes the TextOutW origin as the rotated baseline
                        // start. For the simple case (vAlign=top, single
                        // paragraph), the layout-emitted x,y is at cell left-
                        // top; we shift x by font_size to anchor at the right
                        // edge of the first (top-most) rotated character
                        // since rotated chars extend leftward from origin.
                        let weight = if *bold { 700i32 } else { 400i32 };
                        let ital = if *italic { 1u32 } else { 0u32 };
                        let family_wide: Vec<u16> = family.encode_utf16().chain(std::iter::once(0)).collect();
                        let (escapement, orientation) = if *is_vertical {
                            (-900i32, -900i32)
                        } else {
                            (0i32, 0i32)
                        };
                        let font = CreateFontW(
                            -fs, lf_width, escapement, orientation, weight,
                            ital, 0, 0,
                            1, // DEFAULT_CHARSET
                            0, 0,
                            5, // CLEARTYPE_QUALITY — matches Word GDI rendering
                            0,
                            PCWSTR(family_wide.as_ptr()),
                        );
                        let old_font = SelectObject(mem_dc, font);

                        // Apply character spacing (justify gap, explicit cs from XML)
                        let cs_px = (*character_spacing as f64 * scale).round() as i32;
                        if cs_px != 0 {
                            SetTextCharacterExtra(mem_dc, cs_px);
                        }

                        // Draw text
                        let text_wide: Vec<u16> = text.encode_utf16().collect();
                        // For vertical text, anchor the rotated glyphs along
                        // the right edge of the cell. The TextOutW origin in
                        // rotated mode is at the BASELINE start of the first
                        // glyph (= top-right of unrotated char). Shifting
                        // x_draw rightward by fs places the rotated chars
                        // starting at the cell's right side. This matches
                        // Word's tbRlV reading direction (text flows down-
                        // right).
                        let (x_draw, y_draw) = if *is_vertical {
                            (x + fs, glyph_y)
                        } else {
                            (x, glyph_y)
                        };
                        TextOutW(mem_dc, x_draw, y_draw, &text_wide);

                        // Reset character extra
                        if cs_px != 0 {
                            SetTextCharacterExtra(mem_dc, 0);
                        }

                        // Underline: use OTM metrics for correct positioning
                        if *underline {
                            let mut tm = TEXTMETRICW::default();
                            GetTextMetricsW(mem_dc, &mut tm);
                            let ascent = tm.tmAscent;
                            // OTM underline metrics (font-specific)
                            let (ul_offset, ul_size) = {
                                let mut otm_buf = vec![0u8; 512];
                                let got = GetOutlineTextMetricsW(mem_dc, 512,
                                    Some(otm_buf.as_mut_ptr() as *mut OUTLINETEXTMETRICW));
                                if got > 0 {
                                    let otm = &*(otm_buf.as_ptr() as *const OUTLINETEXTMETRICW);
                                    (otm.otmsUnderscorePosition.abs(), otm.otmsUnderscoreSize.max(1) as i32)
                                } else {
                                    (3_i32, 1_i32) // fallback
                                }
                            };
                            let pen = CreatePen(PS_SOLID, ul_size, rgb);
                            let old_pen = SelectObject(mem_dc, pen);
                            let is_double = underline_style.as_deref() == Some("double");
                            if is_double {
                                let ul_y1 = glyph_y + ascent + ul_offset;
                                let ul_y2 = ul_y1 + ul_size + 2;
                                MoveToEx(mem_dc, x, ul_y1, None);
                                LineTo(mem_dc, x + ew, ul_y1);
                                MoveToEx(mem_dc, x, ul_y2, None);
                                LineTo(mem_dc, x + ew, ul_y2);
                            } else {
                                let ul_y = glyph_y + ascent + ul_offset;
                                MoveToEx(mem_dc, x, ul_y, None);
                                LineTo(mem_dc, x + ew, ul_y);
                            }
                            SelectObject(mem_dc, old_pen);
                            let _ = DeleteObject(pen);
                        }

                        if *strikethrough || *double_strikethrough {
                            let mut tm = TEXTMETRICW::default();
                            GetTextMetricsW(mem_dc, &mut tm);
                            let st_y = glyph_y + tm.tmAscent / 2;
                            let st_pen = CreatePen(PS_SOLID, 1, rgb);
                            let old_pen = SelectObject(mem_dc, st_pen);
                            if *double_strikethrough {
                                // R66: two parallel lines around the single-strike y.
                                // Gap is 0.18 × fs (supersampled px) — larger than the
                                // DWrite 0.08 because GDI here renders in a supersampled
                                // DC then downsamples, halving the visual separation.
                                let gap = ((fs as f64 * 0.18).round() as i32).max(3);
                                let st_y1 = st_y - gap / 2;
                                let st_y2 = st_y1 + gap;
                                MoveToEx(mem_dc, x, st_y1, None);
                                LineTo(mem_dc, x + ew, st_y1);
                                MoveToEx(mem_dc, x, st_y2, None);
                                LineTo(mem_dc, x + ew, st_y2);
                            } else {
                                MoveToEx(mem_dc, x, st_y, None);
                                LineTo(mem_dc, x + ew, st_y);
                            }
                            SelectObject(mem_dc, old_pen);
                            let _ = DeleteObject(st_pen);
                        }

                        SelectObject(mem_dc, old_font);
                        let _ = DeleteObject(font);
                    }

                    oxidocs_core::layout::LayoutContent::TableBorder { x1, y1, x2, y2, color, width, style } => {
                        let bw = (*width as f64 * scale).max(1.0) as i32;
                        let rgb = color.as_deref()
                            .and_then(|c| {
                                let c = c.strip_prefix('#').unwrap_or(c);
                                if c.len() == 6 {
                                    let r = u8::from_str_radix(&c[0..2], 16).ok()?;
                                    let g = u8::from_str_radix(&c[2..4], 16).ok()?;
                                    let b = u8::from_str_radix(&c[4..6], 16).ok()?;
                                    Some(COLORREF((r as u32) | ((g as u32) << 8) | ((b as u32) << 16)))
                                } else { None }
                            })
                            .unwrap_or(COLORREF(0x00000000));

                        // S480: map OOXML border art style -> GDI cosmetic dash pen.
                        // (Cosmetic PS_DASH/PS_DOT draw at 1px width; GDI is the
                        // pagination/fallback path, not the SSIM gate, so the dash
                        // appearance is approximate. OXI_S480_DISABLE -> solid.)
                        let pen_style = if std::env::var("OXI_S480_DISABLE").is_ok() {
                            PS_SOLID
                        } else {
                            // S480: only the heavy "Stroked" art borders (see DWrite
                            // s480_dash_style); thin line-style borders stay solid.
                            match style.as_deref() {
                                Some("dashDotStroked") => PS_DASHDOT,
                                Some("dashDotDotStroked") => PS_DASHDOTDOT,
                                _ => PS_SOLID,
                            }
                        };
                        // Cosmetic dashed pens require width 1; geometric solid keeps bw.
                        let pen = if pen_style == PS_SOLID {
                            CreatePen(PS_SOLID, bw, rgb)
                        } else {
                            CreatePen(pen_style, 1, rgb)
                        };
                        let old_pen = SelectObject(mem_dc, pen);
                        let bx1 = (*x1 as f64 * scale) as i32;
                        let by1 = (*y1 as f64 * scale) as i32;
                        let bx2 = (*x2 as f64 * scale) as i32;
                        let by2 = (*y2 as f64 * scale) as i32;
                        MoveToEx(mem_dc, bx1, by1, None);
                        LineTo(mem_dc, bx2, by2);
                        SelectObject(mem_dc, old_pen);
                        let _ = DeleteObject(pen);
                    }

                    oxidocs_core::layout::LayoutContent::CellShading { color: shade_color } => {
                        let rgb = {
                            let c = shade_color.strip_prefix('#').unwrap_or(shade_color);
                            if c.len() == 6 {
                                let r = u8::from_str_radix(&c[0..2], 16).unwrap_or(240);
                                let g = u8::from_str_radix(&c[2..4], 16).unwrap_or(240);
                                let b = u8::from_str_radix(&c[4..6], 16).unwrap_or(240);
                                COLORREF((r as u32) | ((g as u32) << 8) | ((b as u32) << 16))
                            } else {
                                COLORREF(0x00F0F0F0)
                            }
                        };
                        let brush = CreateSolidBrush(rgb);
                        let r = RECT { left: x, top: y, right: x + ew, bottom: y + eh };
                        FillRect(mem_dc, &r, brush);
                        let _ = DeleteObject(brush);
                    }

                    oxidocs_core::layout::LayoutContent::BoxRect { fill, stroke_color, stroke_width, corner_radius } => {
                        // Word optimizes out BoxRect with white fill and no border — they
                        // paint white-over-white and are visually indistinguishable from
                        // the page background. Oxi's naive paint obscures text beneath.
                        // COM-confirmed on b35_tokumei_08_01 p1 textbox (白fill + noFill枠).
                        let is_invisible_white = match fill.as_deref() {
                            Some(f) => {
                                let c = f.strip_prefix('#').unwrap_or(f);
                                c.eq_ignore_ascii_case("ffffff") || c.eq_ignore_ascii_case("fff")
                            }
                            None => false,
                        } && stroke_color.is_none();
                        if is_invisible_white {
                            continue;
                        }
                        let cr = (*corner_radius as f64 * scale) as i32;
                        // Create fill brush
                        let fill_brush = if let Some(ref fill_hex) = fill {
                            let c = fill_hex.strip_prefix('#').unwrap_or(fill_hex);
                            if c.len() == 6 {
                                let r = u8::from_str_radix(&c[0..2], 16).unwrap_or(255);
                                let g = u8::from_str_radix(&c[2..4], 16).unwrap_or(255);
                                let b = u8::from_str_radix(&c[4..6], 16).unwrap_or(255);
                                Some(CreateSolidBrush(COLORREF((r as u32) | ((g as u32) << 8) | ((b as u32) << 16))))
                            } else { None }
                        } else { None };
                        // Create stroke pen
                        let stroke_pen = if let Some(ref sc) = stroke_color {
                            let c = sc.strip_prefix('#').unwrap_or(sc);
                            if c.len() == 6 {
                                let r = u8::from_str_radix(&c[0..2], 16).unwrap_or(0);
                                let g = u8::from_str_radix(&c[2..4], 16).unwrap_or(0);
                                let b = u8::from_str_radix(&c[4..6], 16).unwrap_or(0);
                                let sw = (*stroke_width as f64 * scale).max(1.0) as i32;
                                Some(CreatePen(PS_SOLID, sw, COLORREF((r as u32) | ((g as u32) << 8) | ((b as u32) << 16))))
                            } else { None }
                        } else { None };
                        // Select brush and pen
                        let old_brush = fill_brush.map(|b| SelectObject(mem_dc, b));
                        let no_brush = if fill_brush.is_none() { Some(SelectObject(mem_dc, GetStockObject(NULL_BRUSH))) } else { None };
                        let old_pen = stroke_pen.map(|p| SelectObject(mem_dc, p));
                        let no_pen = if stroke_pen.is_none() { Some(SelectObject(mem_dc, GetStockObject(NULL_PEN))) } else { None };
                        // Draw rounded or regular rectangle
                        if cr > 0 {
                            RoundRect(mem_dc, x, y, x + ew, y + eh, cr * 2, cr * 2);
                        } else {
                            Rectangle(mem_dc, x, y, x + ew, y + eh);
                        }
                        // Cleanup
                        if let Some(ob) = old_brush { SelectObject(mem_dc, ob); }
                        if let Some(nb) = no_brush { SelectObject(mem_dc, nb); }
                        if let Some(op) = old_pen { SelectObject(mem_dc, op); }
                        if let Some(np) = no_pen { SelectObject(mem_dc, np); }
                        if let Some(ref _f) = fill_brush { /* DeleteObject handled by HGDIOBJ drop */ }
                        if let Some(ref _s) = stroke_pen { /* DeleteObject handled by HGDIOBJ drop */ }
                    }

                    oxidocs_core::layout::LayoutContent::ClipStart => {
                        SaveDC(mem_dc);
                        let rgn = CreateRectRgn(x, y, x + ew, y + eh);
                        SelectClipRgn(mem_dc, rgn);
                        let _ = DeleteObject(rgn);
                    }
                    oxidocs_core::layout::LayoutContent::ClipEnd => {
                        RestoreDC(mem_dc, -1);
                    }
                    oxidocs_core::layout::LayoutContent::Image { ref data, .. } => {
                        if !data.is_empty() {
                            // Decode image and draw via GDI StretchDIBits
                            if let Ok(img) = image::load_from_memory(data) {
                                let rgba = img.to_rgba8();
                                let (pw, ph) = rgba.dimensions();
                                // Create DIB and draw
                                let mut bmi = BITMAPINFO {
                                    bmiHeader: BITMAPINFOHEADER {
                                        biSize: std::mem::size_of::<BITMAPINFOHEADER>() as u32,
                                        biWidth: pw as i32,
                                        biHeight: -(ph as i32), // top-down
                                        biPlanes: 1,
                                        biBitCount: 32,
                                        biCompression: 0,
                                        ..Default::default()
                                    },
                                    ..Default::default()
                                };
                                // BGRA pixel data
                                let mut pixels: Vec<u8> = Vec::with_capacity((pw * ph * 4) as usize);
                                for pixel in rgba.pixels() {
                                    pixels.push(pixel[2]); // B
                                    pixels.push(pixel[1]); // G
                                    pixels.push(pixel[0]); // R
                                    pixels.push(pixel[3]); // A
                                }
                                StretchDIBits(mem_dc, x, y, ew, eh,
                                    0, 0, pw as i32, ph as i32,
                                    Some(pixels.as_ptr() as *const _),
                                    &bmi, DIB_RGB_COLORS, SRCCOPY);
                            }
                        }
                    }
                    oxidocs_core::layout::LayoutContent::PresetShape { shape_type, stroke_color, stroke_width } => {
                        // Parse stroke color once — used by all shape branches.
                        let (sr, sg, sb) = if let Some(ref sc) = stroke_color {
                            let c = sc.strip_prefix('#').unwrap_or(sc);
                            if c.len() == 6 {
                                (
                                    u8::from_str_radix(&c[0..2], 16).unwrap_or(0),
                                    u8::from_str_radix(&c[2..4], 16).unwrap_or(0),
                                    u8::from_str_radix(&c[4..6], 16).unwrap_or(0),
                                )
                            } else { (0, 0, 0) }
                        } else { (0, 0, 0) };
                        let color_ref = COLORREF((sr as u32) | ((sg as u32) << 8) | ((sb as u32) << 16));
                        let sw_px = (*stroke_width as f64 * scale).round().max(1.0) as i32;

                        match shape_type.as_str() {
                            "bracketPair" => {
                                let pen = CreatePen(PS_SOLID, sw_px, color_ref);
                                let old_pen = SelectObject(mem_dc, pen);
                                let old_brush = SelectObject(mem_dc, GetStockObject(NULL_BRUSH));
                                // OOXML bracketPair: corner radius = 8.387% of min(w,h)
                                let r = ((ew.min(eh) as f64) * 0.08387).round().max(2.0) as i32;
                                let k = 0.5522847;
                                let kr = (r as f64 * k).round() as i32;
                                // Left bracket
                                {
                                    let pts = [
                                        POINT { x: x + r, y },
                                        POINT { x: x + r - kr, y },
                                        POINT { x: x, y: y + r - kr },
                                        POINT { x: x, y: y + r },
                                        POINT { x: x, y: y + eh - r },
                                        POINT { x: x, y: y + eh - r },
                                        POINT { x: x, y: y + eh - r },
                                        POINT { x: x, y: y + eh - r + kr },
                                        POINT { x: x + r - kr, y: y + eh },
                                        POINT { x: x + r, y: y + eh },
                                    ];
                                    PolyBezier(mem_dc, &pts);
                                }
                                // Right bracket
                                {
                                    let pts = [
                                        POINT { x: x + ew - r, y },
                                        POINT { x: x + ew - r + kr, y },
                                        POINT { x: x + ew, y: y + r - kr },
                                        POINT { x: x + ew, y: y + r },
                                        POINT { x: x + ew, y: y + eh - r },
                                        POINT { x: x + ew, y: y + eh - r },
                                        POINT { x: x + ew, y: y + eh - r },
                                        POINT { x: x + ew, y: y + eh - r + kr },
                                        POINT { x: x + ew - r + kr, y: y + eh },
                                        POINT { x: x + ew - r, y: y + eh },
                                    ];
                                    PolyBezier(mem_dc, &pts);
                                }
                                SelectObject(mem_dc, old_brush);
                                SelectObject(mem_dc, old_pen);
                                let _ = DeleteObject(pen);
                            }
                            // OOXML prstGeom prst="line": diagonal line from top-left to
                            // bottom-right of the bounding box. With width=0 → vertical;
                            // with height=0 → horizontal. Used as divider/separator.
                            // 3a4f9fbe1a83 uses 27 of these (24 inside textboxes).
                            "line" => {
                                let pen = CreatePen(PS_SOLID, sw_px, color_ref);
                                let old_pen = SelectObject(mem_dc, pen);
                                let pts = [
                                    POINT { x, y },
                                    POINT { x: x + ew, y: y + eh },
                                ];
                                Polyline(mem_dc, &pts);
                                SelectObject(mem_dc, old_pen);
                                let _ = DeleteObject(pen);
                            }
                            // OOXML prstGeom prst="leftBracket" | "rightBracket":
                            // single "[" or "]" shape, outline-only (no fill),
                            // drawn as rounded-corner path along one side of the bbox.
                            // Shape adj default = 8.387% corner radius (same as bracketPair).
                            "leftBracket" | "rightBracket" => {
                                let pen = CreatePen(PS_SOLID, sw_px, color_ref);
                                let old_pen = SelectObject(mem_dc, pen);
                                let old_brush = SelectObject(mem_dc, GetStockObject(NULL_BRUSH));
                                let r = ((ew.min(eh) as f64) * 0.08387).round().max(2.0) as i32;
                                let k = 0.5522847;
                                let kr = (r as f64 * k).round() as i32;
                                let pts = if shape_type == "leftBracket" {
                                    // "[" — curves at top-left + bottom-left of bbox
                                    [
                                        POINT { x: x + r, y },                       // top start (a bit right of corner)
                                        POINT { x: x + r - kr, y },                   // cubic handle 1
                                        POINT { x, y: y + r - kr },                    // cubic handle 2
                                        POINT { x, y: y + r },                          // top-left curve end
                                        POINT { x, y: y + eh - r },                    // straight down
                                        POINT { x, y: y + eh - r },                    // (duplicate for cubic)
                                        POINT { x, y: y + eh - r },                    //
                                        POINT { x, y: y + eh - r + kr },               // cubic handle
                                        POINT { x: x + r - kr, y: y + eh },            // cubic handle
                                        POINT { x: x + r, y: y + eh },                // bottom end
                                    ]
                                } else {
                                    // "]" — curves at top-right + bottom-right of bbox
                                    [
                                        POINT { x: x + ew - r, y },
                                        POINT { x: x + ew - r + kr, y },
                                        POINT { x: x + ew, y: y + r - kr },
                                        POINT { x: x + ew, y: y + r },
                                        POINT { x: x + ew, y: y + eh - r },
                                        POINT { x: x + ew, y: y + eh - r },
                                        POINT { x: x + ew, y: y + eh - r },
                                        POINT { x: x + ew, y: y + eh - r + kr },
                                        POINT { x: x + ew - r + kr, y: y + eh },
                                        POINT { x: x + ew - r, y: y + eh },
                                    ]
                                };
                                PolyBezier(mem_dc, &pts);
                                SelectObject(mem_dc, old_brush);
                                SelectObject(mem_dc, old_pen);
                                let _ = DeleteObject(pen);
                            }
                            _ => {}
                        }
                    }
                    // R-05g: render comment balloon as a rounded-rect filled
                    // with the author's tint, then draw author header line +
                    // body + indented reply blocks inside.
                    oxidocs_core::layout::LayoutContent::Balloon {
                        author, author_color_index, resolved, body, replies, ..
                    } => {
                        let tint_hex = oxidocs_core::layout::comment_balloon_fill(*author_color_index, *resolved);
                        let (br, bg, bb) = parse_hex_rgb(tint_hex);
                        let fill = CreateSolidBrush(COLORREF((br as u32) | ((bg as u32) << 8) | ((bb as u32) << 16)));
                        // Subtle border in a slightly darker shade of the tint.
                        let border_pen = CreatePen(PS_SOLID, 1, COLORREF(((br.saturating_sub(40)) as u32) | (((bg.saturating_sub(40)) as u32) << 8) | (((bb.saturating_sub(40)) as u32) << 16)));
                        let old_brush = SelectObject(mem_dc, fill);
                        let old_pen = SelectObject(mem_dc, border_pen);
                        let radius = (4.0 * scale).round().max(1.0) as i32;
                        RoundRect(mem_dc, x, y, x + ew, y + eh, radius * 2, radius * 2);
                        SelectObject(mem_dc, old_brush);
                        SelectObject(mem_dc, old_pen);
                        let _ = DeleteObject(fill);
                        let _ = DeleteObject(border_pen);

                        // Header + body text. Use 9pt-equivalent for compactness.
                        let header_fs = (9.0 * scale).round() as i32;
                        let body_fs = (10.0 * scale).round() as i32;
                        let family_wide: Vec<u16> = "Calibri".encode_utf16().chain(std::iter::once(0)).collect();
                        let header_font = CreateFontW(-header_fs, 0, 0, 0, 700, 0, 0, 0, 1, 0, 0, 5, 0, PCWSTR(family_wide.as_ptr()));
                        let body_font = CreateFontW(-body_fs, 0, 0, 0, 400, 0, 0, 0, 1, 0, 0, 5, 0, PCWSTR(family_wide.as_ptr()));

                        let pad = (4.0 * scale).round() as i32;
                        let mut text_y = y + pad;
                        let text_x = x + pad;
                        let text_right = x + ew - pad;

                        // Author header (bold, slightly darker than body color).
                        SetBkMode(mem_dc, TRANSPARENT);
                        SetTextColor(mem_dc, COLORREF(0x00404040));
                        let old_f = SelectObject(mem_dc, header_font);
                        let header_str = format!("{}", author);
                        let header_wide: Vec<u16> = header_str.encode_utf16().collect();
                        let mut header_rect = RECT { left: text_x, top: text_y, right: text_right, bottom: text_y + header_fs + pad };
                        DrawTextW(mem_dc, &mut header_wide.clone(), &mut header_rect, DT_LEFT | DT_TOP | DT_NOCLIP);
                        text_y = header_rect.bottom + 1;
                        SelectObject(mem_dc, old_f);
                        let _ = DeleteObject(header_font);

                        // Body text — wrap inside balloon width.
                        SetTextColor(mem_dc, COLORREF(0x00000000));
                        let old_f = SelectObject(mem_dc, body_font);
                        let body_wide: Vec<u16> = body.encode_utf16().collect();
                        let mut body_rect = RECT { left: text_x, top: text_y, right: text_right, bottom: y + eh - pad };
                        let body_h = DrawTextW(mem_dc, &mut body_wide.clone(), &mut body_rect, DT_LEFT | DT_TOP | DT_WORDBREAK);
                        text_y += body_h + pad;
                        SelectObject(mem_dc, old_f);

                        // Replies (R-08) — indented ~10pt, each with its own author chip.
                        let indent_px = (10.0 * scale).round() as i32;
                        for reply in replies.iter() {
                            // Reply author chip
                            SetTextColor(mem_dc, COLORREF(0x00606060));
                            let header_font2 = CreateFontW(-header_fs, 0, 0, 0, 700, 0, 0, 0, 1, 0, 0, 5, 0, PCWSTR(family_wide.as_ptr()));
                            let oldf = SelectObject(mem_dc, header_font2);
                            let h_str = format!("{}", reply.author);
                            let h_wide: Vec<u16> = h_str.encode_utf16().collect();
                            let mut hr = RECT { left: text_x + indent_px, top: text_y, right: text_right, bottom: text_y + header_fs + pad };
                            DrawTextW(mem_dc, &mut h_wide.clone(), &mut hr, DT_LEFT | DT_TOP | DT_NOCLIP);
                            text_y = hr.bottom + 1;
                            SelectObject(mem_dc, oldf);
                            let _ = DeleteObject(header_font2);
                            // Reply body
                            let body_font2 = CreateFontW(-body_fs, 0, 0, 0, 400, 0, 0, 0, 1, 0, 0, 5, 0, PCWSTR(family_wide.as_ptr()));
                            let oldf = SelectObject(mem_dc, body_font2);
                            SetTextColor(mem_dc, COLORREF(0x00000000));
                            let r_wide: Vec<u16> = reply.body.encode_utf16().collect();
                            let mut rr = RECT { left: text_x + indent_px, top: text_y, right: text_right, bottom: y + eh - pad };
                            let rh = DrawTextW(mem_dc, &mut r_wide.clone(), &mut rr, DT_LEFT | DT_TOP | DT_WORDBREAK);
                            text_y += rh + pad;
                            SelectObject(mem_dc, oldf);
                            let _ = DeleteObject(body_font2);
                        }

                        let _ = DeleteObject(body_font);
                    }
                    // R-05g: dotted connector line from the inline anchor to
                    // the balloon's left edge, in the author's tint hue.
                    // Word's connector is roughly 0.75pt thick — dim grey
                    // dotted, visible against white background. We use a
                    // medium grey (#808080) at 1pt for v1 since PS_DOT with
                    // a colored pen is hard to see at small thickness.
                    oxidocs_core::layout::LayoutContent::BalloonConnector {
                        from_x, from_y, to_x, to_y, color_hex: _,
                    } => {
                        // Slightly thicker grey dotted line so it's visible
                        // at typical screen DPI without being intrusive.
                        let pen = CreatePen(
                            PS_DOT,
                            (1.0 * scale).round().max(1.0) as i32,
                            COLORREF(0x00808080),
                        );
                        let old_pen = SelectObject(mem_dc, pen);
                        let fx = (*from_x as f64 * scale).round() as i32;
                        let fy = (*from_y as f64 * scale).round() as i32;
                        let tx = (*to_x as f64 * scale).round() as i32;
                        let ty = (*to_y as f64 * scale).round() as i32;
                        let _ = MoveToEx(mem_dc, fx, fy, None);
                        let _ = LineTo(mem_dc, tx, ty);
                        SelectObject(mem_dc, old_pen);
                        let _ = DeleteObject(pen);
                    }
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
            GetDIBits(mem_dc, bitmap, 0, h as u32, Some(pixels.as_mut_ptr() as *mut _), &mut bmi, DIB_RGB_COLORS);

            // Convert BGRA to RGB and save as PNG
            let mut rgb_pixels = Vec::with_capacity((w * h * 3) as usize);
            for i in 0..(w * h) as usize {
                rgb_pixels.push(pixels[i * 4 + 2]); // R
                rgb_pixels.push(pixels[i * 4 + 1]); // G
                rgb_pixels.push(pixels[i * 4]);      // B
            }

            let img = image::RgbImage::from_raw(w as u32, h as u32, rgb_pixels)
                .expect("Failed to create image");
            let final_img = if supersample > 1 && (w as u32 != out_w || h as u32 != out_h) {
                let dynamic = image::DynamicImage::ImageRgb8(img);
                dynamic.resize_exact(out_w, out_h, image::imageops::FilterType::Lanczos3).to_rgb8()
            } else {
                img
            };
            let out_path = format!("{}_p{}.png", prefix, page_idx + 1);
            final_img.save(&out_path).expect("Failed to save PNG");
            eprintln!("  Saved {} ({}x{})", out_path, final_img.width(), final_img.height());

            // Cleanup
            SelectObject(mem_dc, old_bmp);
            let _ = DeleteObject(bitmap);
            DeleteDC(mem_dc);
            ReleaseDC(HWND(std::ptr::null_mut()), screen_dc);
        }
    }
}

/// Dump layout elements to JSON for per-element x/y diffing against Word DML.
/// Minimal schema — enough to locate each text fragment in document space.
fn dump_layout_json(result: &oxidocs_core::layout::LayoutResult, path: &str) {
    use oxidocs_core::layout::LayoutContent;
    use std::fmt::Write;
    let mut out = String::from("{\n  \"pages\": [\n");
    for (pi, page) in result.pages.iter().enumerate() {
        if pi > 0 { out.push_str(",\n"); }
        write!(&mut out, "    {{\"page\": {}, \"width\": {:.3}, \"height\": {:.3}, \"elements\": [\n",
               pi + 1, page.width, page.height).unwrap();
        let mut first = true;
        for el in &page.elements {
            let (kind, text_json, font_size) = match &el.content {
                LayoutContent::Text { text, font_size, .. } => {
                    let mut esc = String::with_capacity(text.len());
                    for c in text.chars() {
                        match c {
                            '\\' => esc.push_str("\\\\"),
                            '"'  => esc.push_str("\\\""),
                            '\n' => esc.push_str("\\n"),
                            '\r' => esc.push_str("\\r"),
                            '\t' => esc.push_str("\\t"),
                            c if (c as u32) < 0x20 => {
                                use std::fmt::Write as _;
                                write!(&mut esc, "\\u{:04x}", c as u32).unwrap();
                            }
                            c => esc.push(c),
                        }
                    }
                    ("text", format!("\"{}\"", esc), *font_size)
                }
                LayoutContent::Image { .. } => ("image", "null".to_string(), 0.0),
                LayoutContent::TableBorder { .. } => ("border", "null".to_string(), 0.0),
                LayoutContent::CellShading { .. } => ("shading", "null".to_string(), 0.0),
                _ => ("other", "null".to_string(), 0.0),
            };
            if !first { out.push_str(",\n"); }
            first = false;
            let pi_json = el.paragraph_index.map(|v| v.to_string()).unwrap_or_else(|| "null".to_string());
            let ri_json = el.run_index.map(|v| v.to_string()).unwrap_or_else(|| "null".to_string());
            let co_json = el.char_offset.map(|v| v.to_string()).unwrap_or_else(|| "null".to_string());
            // R7.32: emit cell_para_idx so aggregate_dump can split cell paragraphs
            let cpi_json = el.cell_paragraph_index.map(|v| v.to_string()).unwrap_or_else(|| "null".to_string());
            // R7.44: emit cell_row_idx / cell_col_idx so aggregate_dump can
            // distinguish cells sharing (block_idx, cpi=0) — fixes the
            // "千円千円千円千円" collapse.
            let cri_json = el.cell_row_index.map(|v| v.to_string()).unwrap_or_else(|| "null".to_string());
            let cci_json = el.cell_col_index.map(|v| v.to_string()).unwrap_or_else(|| "null".to_string());
            // Session 73 Phase B (2026-05-17): emit text_y_off so element_iou_diff.py
            // and measure_pagination_oxi.py can subtract it to compare against Word's
            // Information(6) (LINE BOX TOP convention). See
            // memory/session71_y_convention_refactor_design.md.
            write!(&mut out,
                "      {{\"type\": \"{}\", \"x\": {:.3}, \"y\": {:.3}, \"w\": {:.3}, \"h\": {:.3}, \"text\": {}, \"font_size\": {:.2}, \"para_idx\": {}, \"run_idx\": {}, \"char_offset\": {}, \"cell_para_idx\": {}, \"cell_row_idx\": {}, \"cell_col_idx\": {}, \"text_y_off\": {:.3}}}",
                kind, el.x, el.y, el.width, el.height, text_json, font_size, pi_json, ri_json, co_json, cpi_json, cri_json, cci_json, el.text_y_off).unwrap();
        }
        out.push_str("\n    ]}");
    }
    out.push_str("\n  ]\n}\n");
    std::fs::write(path, out).expect("Failed to write dump-layout JSON");
}
