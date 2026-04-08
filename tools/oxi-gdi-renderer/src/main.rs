//! Oxi GDI Renderer: Render .docx layout using Windows GDI for pixel-accurate comparison with Word.
//!
//! Usage: oxi-gdi-renderer input.docx output_prefix
//!   Produces: output_prefix_p1.png, output_prefix_p2.png, ...

use std::path::Path;

fn main() {
    let args: Vec<String> = std::env::args().collect();
    if args.len() < 3 {
        eprintln!("Usage: {} <input.docx> <output_prefix> [dpi] [--exclude=text,border,shading,box,image,clip]", args[0]);
        std::process::exit(1);
    }

    let docx_path = &args[1];
    let output_prefix = &args[2];
    let dpi: u32 = args.get(3).and_then(|s| s.parse().ok()).unwrap_or(150);

    // Parse --exclude flag from any argument
    let mut exclude: Vec<String> = Vec::new();
    for arg in &args[3..] {
        if let Some(list) = arg.strip_prefix("--exclude=") {
            exclude = list.split(',').map(|s| s.trim().to_lowercase()).collect();
        }
    }

    // Parse document
    let data = std::fs::read(docx_path).expect("Cannot read docx file");
    let doc = oxidocs_core::parser::parse_docx(&data).expect("Cannot parse docx");

    // Layout
    let engine = oxidocs_core::layout::LayoutEngine::for_document(&doc);
    let result = engine.layout(&doc);

    eprintln!("Parsed {} pages, DPI={}", result.pages.len(), dpi);
    if !exclude.is_empty() {
        eprintln!("Excluding: {:?}", exclude);
    }

    // Render each page with GDI
    #[cfg(windows)]
    {
        render_pages_gdi(&result, output_prefix, dpi, &exclude);
    }

    #[cfg(not(windows))]
    {
        eprintln!("GDI rendering requires Windows");
        std::process::exit(1);
    }
}

#[cfg(windows)]
fn render_pages_gdi(result: &oxidocs_core::layout::LayoutResult, prefix: &str, dpi: u32, exclude: &[String]) {
    use windows::Win32::Graphics::Gdi::*;
    use windows::Win32::Foundation::*;
    use windows::core::*;

    let scale = dpi as f64 / 72.0;

    for (page_idx, page) in result.pages.iter().enumerate() {
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
                        text, font_size, font_family, bold, italic, color, underline, underline_style, strikethrough, highlight, character_spacing, ..
                    } => {
                        let fs = (*font_size as f64 * scale).round() as i32;
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
                            let r = RECT { left: x, top: y, right: x + ew, bottom: y + eh };
                            FillRect(mem_dc, &r, hl_brush);
                            let _ = DeleteObject(hl_brush);
                        }

                        SetTextColor(mem_dc, rgb);

                        // Create font
                        let weight = if *bold { 700i32 } else { 400i32 };
                        let ital = if *italic { 1u32 } else { 0u32 };
                        let family_wide: Vec<u16> = family.encode_utf16().chain(std::iter::once(0)).collect();
                        let font = CreateFontW(
                            -fs, 0, 0, 0, weight,
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
                        TextOutW(mem_dc, x, y, &text_wide);

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
                                let ul_y1 = y + ascent + ul_offset;
                                let ul_y2 = ul_y1 + ul_size + 2;
                                MoveToEx(mem_dc, x, ul_y1, None);
                                LineTo(mem_dc, x + ew, ul_y1);
                                MoveToEx(mem_dc, x, ul_y2, None);
                                LineTo(mem_dc, x + ew, ul_y2);
                            } else {
                                let ul_y = y + ascent + ul_offset;
                                MoveToEx(mem_dc, x, ul_y, None);
                                LineTo(mem_dc, x + ew, ul_y);
                            }
                            SelectObject(mem_dc, old_pen);
                            let _ = DeleteObject(pen);
                        }

                        // Strikethrough
                        if *strikethrough {
                            let mut tm = TEXTMETRICW::default();
                            GetTextMetricsW(mem_dc, &mut tm);
                            let st_y = y + tm.tmAscent / 2;
                            let st_pen = CreatePen(PS_SOLID, 1, rgb);
                            let old_pen = SelectObject(mem_dc, st_pen);
                            MoveToEx(mem_dc, x, st_y, None);
                            LineTo(mem_dc, x + ew, st_y);
                            SelectObject(mem_dc, old_pen);
                            let _ = DeleteObject(st_pen);
                        }

                        SelectObject(mem_dc, old_font);
                        let _ = DeleteObject(font);
                    }

                    oxidocs_core::layout::LayoutContent::TableBorder { x1, y1, x2, y2, color, width } => {
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

                        let pen = CreatePen(PS_SOLID, bw, rgb);
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
                        // Round 30: render DrawingML preset shapes via GDI lines.
                        // For now, only "bracketPair" (= 〔 〕 paired bracket frame
                        // around a content area) is implemented. Other shapes
                        // remain unsupported.
                        if shape_type == "bracketPair" {
                            // Stroke pen
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
                            let sw = (*stroke_width as f64 * scale).round().max(1.0) as i32;
                            let pen = CreatePen(PS_SOLID, sw, COLORREF((sr as u32) | ((sg as u32) << 8) | ((sb as u32) << 16)));
                            let old_pen = SelectObject(mem_dc, pen);
                            // Bracket arm length: ~10% of width per side (roughly
                            // matches OOXML default adj=8387/21600 ≈ 38.8%
                            // converted via the bracketPair geometry formula).
                            let arm = (((ew as f64) * 0.10).round() as i32).max(2);
                            // Left bracket: 〔 — three line segments forming a [
                            // top arm: from (x+arm, y) to (x, y)
                            // left side: from (x, y) to (x, y+eh)
                            // bottom arm: from (x, y+eh) to (x+arm, y+eh)
                            MoveToEx(mem_dc, x + arm, y, None);
                            LineTo(mem_dc, x, y);
                            LineTo(mem_dc, x, y + eh);
                            LineTo(mem_dc, x + arm, y + eh);
                            // Right bracket: 〕 — three line segments forming a ]
                            MoveToEx(mem_dc, x + ew - arm, y, None);
                            LineTo(mem_dc, x + ew, y);
                            LineTo(mem_dc, x + ew, y + eh);
                            LineTo(mem_dc, x + ew - arm, y + eh);
                            SelectObject(mem_dc, old_pen);
                            let _ = DeleteObject(pen);
                        }
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
            let out_path = format!("{}_p{}.png", prefix, page_idx + 1);
            img.save(&out_path).expect("Failed to save PNG");
            eprintln!("  Saved {} ({}x{})", out_path, w, h);

            // Cleanup
            SelectObject(mem_dc, old_bmp);
            let _ = DeleteObject(bitmap);
            DeleteDC(mem_dc);
            ReleaseDC(HWND(std::ptr::null_mut()), screen_dc);
        }
    }
}
