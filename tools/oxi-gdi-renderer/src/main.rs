//! Oxi GDI Renderer: Render .docx layout using Windows GDI for pixel-accurate comparison with Word.
//!
//! Usage: oxi-gdi-renderer input.docx output_prefix
//!   Produces: output_prefix_p1.png, output_prefix_p2.png, ...

use std::path::Path;

fn main() {
    let args: Vec<String> = std::env::args().collect();
    if args.len() < 3 {
        eprintln!("Usage: {} <input.docx> <output_prefix>", args[0]);
        std::process::exit(1);
    }

    let docx_path = &args[1];
    let output_prefix = &args[2];
    let dpi: u32 = args.get(3).and_then(|s| s.parse().ok()).unwrap_or(150);

    // Parse document
    let data = std::fs::read(docx_path).expect("Cannot read docx file");
    let doc = oxidocs_core::parser::parse_docx(&data).expect("Cannot parse docx");

    // Layout
    let engine = oxidocs_core::layout::LayoutEngine::for_document(&doc);
    let result = engine.layout(&doc);

    eprintln!("Parsed {} pages, DPI={}", result.pages.len(), dpi);

    // Render each page with GDI
    #[cfg(windows)]
    {
        render_pages_gdi(&result, output_prefix, dpi);
    }

    #[cfg(not(windows))]
    {
        eprintln!("GDI rendering requires Windows");
        std::process::exit(1);
    }
}

#[cfg(windows)]
fn render_pages_gdi(result: &oxidocs_core::layout::LayoutResult, prefix: &str, dpi: u32) {
    use windows::Win32::Graphics::Gdi::*;
    use windows::Win32::Foundation::*;
    use windows::core::*;

    let scale = dpi as f64 / 72.0;

    for (page_idx, page) in result.pages.iter().enumerate() {
        let w = (page.width as f64 * scale) as i32;
        let h = (page.height as f64 * scale) as i32;

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

            // Render each element
            for elem in &page.elements {
                let x = (elem.x as f64 * scale) as i32;
                let y = (elem.y as f64 * scale) as i32;
                let ew = (elem.width as f64 * scale) as i32;
                let eh = (elem.height as f64 * scale) as i32;

                match &elem.content {
                    oxidocs_core::layout::LayoutContent::Text {
                        text, font_size, font_family, bold, color, underline, underline_style, ..
                    } => {
                        let fs = (*font_size as f64 * scale) as i32;
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

                        SetTextColor(mem_dc, rgb);

                        // Create font
                        let weight = if *bold { 700i32 } else { 400i32 };
                        let family_wide: Vec<u16> = family.encode_utf16().chain(std::iter::once(0)).collect();
                        let font = CreateFontW(
                            -fs, 0, 0, 0, weight,
                            0, 0, 0,
                            1, // DEFAULT_CHARSET
                            0, 0, 0, 0,
                            PCWSTR(family_wide.as_ptr()),
                        );
                        let old_font = SelectObject(mem_dc, font);

                        // Draw text
                        let text_wide: Vec<u16> = text.encode_utf16().collect();
                        TextOutW(mem_dc, x, y, &text_wide);

                        // Underline
                        if *underline {
                            let ul_y = y + fs + (fs as f64 * 0.15) as i32;
                            let ul_w = (fs as f64 * 0.05).max(1.0) as i32;
                            let pen = CreatePen(PS_SOLID, ul_w, rgb);
                            let old_pen = SelectObject(mem_dc, pen);
                            let is_double = underline_style.as_deref() == Some("double");
                            if is_double {
                                let ul_y1 = y + fs + (fs as f64 * 0.10) as i32;
                                let ul_y2 = y + fs + (fs as f64 * 0.22) as i32;
                                MoveToEx(mem_dc, x, ul_y1, None);
                                LineTo(mem_dc, x + ew, ul_y1);
                                MoveToEx(mem_dc, x, ul_y2, None);
                                LineTo(mem_dc, x + ew, ul_y2);
                            } else {
                                MoveToEx(mem_dc, x, ul_y, None);
                                LineTo(mem_dc, x + ew, ul_y);
                            }
                            SelectObject(mem_dc, old_pen);
                            let _ = DeleteObject(pen);
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

                    _ => {} // Skip ClipStart/End, Image, PresetShape for now
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
