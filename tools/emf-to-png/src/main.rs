//! EMF to PNG converter using Windows GDI PlayEnhMetaFile.
//! Usage: emf-to-png input.emf output.png [dpi]

fn main() {
    let args: Vec<String> = std::env::args().collect();
    if args.len() < 3 {
        eprintln!("Usage: {} <input.emf> <output.png> [dpi]", args[0]);
        std::process::exit(1);
    }
    let emf_path = &args[1];
    let out_path = &args[2];
    let dpi: u32 = args.get(3).and_then(|s| s.parse().ok()).unwrap_or(150);

    #[cfg(windows)]
    {
        convert_emf_to_png(emf_path, out_path, dpi);
    }
    #[cfg(not(windows))]
    {
        eprintln!("Windows only");
        std::process::exit(1);
    }
}

#[cfg(windows)]
fn convert_emf_to_png(emf_path: &str, out_path: &str, dpi: u32) {
    use windows::Win32::Graphics::Gdi::*;
    use windows::Win32::Foundation::*;
    use windows::core::*;

    let emf_data = std::fs::read(emf_path).expect("Cannot read EMF file");

    unsafe {
        // Create EMF from bytes
        let hemf = SetEnhMetaFileBits(&emf_data);
        if hemf.is_invalid() {
            eprintln!("Failed to create EMF from bytes");
            std::process::exit(1);
        }

        // Get EMF header for dimensions
        let mut header = ENHMETAHEADER::default();
        let hdr_size = std::mem::size_of::<ENHMETAHEADER>() as u32;
        GetEnhMetaFileHeader(hemf, hdr_size, Some(&mut header));

        let frame_w = header.rclFrame.right - header.rclFrame.left;
        let frame_h = header.rclFrame.bottom - header.rclFrame.top;

        // Derive page size from EMF frame + device metrics.
        // EMF frame is in 0.01mm. We need full-page pixel size.
        // For Word CopyAsPicture, the EMF content covers the page margins area.
        // We play the EMF into the full page rectangle for simplest mapping.
        let frame_w_mm = frame_w as f64 / 100.0;
        let frame_h_mm = frame_h as f64 / 100.0;
        // Full page including margins (approximate from frame * page/content ratio)
        // Use MulDiv-based sizes for Letter/A4 detection
        let page_w_pt: f64;
        let page_h_pt: f64;
        if frame_h_mm > 270.0 { // A4-ish
            page_w_pt = 595.3;
            page_h_pt = 841.9;
        } else { // Letter
            page_w_pt = 612.0;
            page_h_pt = 792.0;
        }
        let w = (page_w_pt * dpi as f64 / 72.0).round() as i32;
        let h = (page_h_pt * dpi as f64 / 72.0).round() as i32;
        eprintln!("EMF frame: {:.1}x{:.1}mm, page: {:.0}x{:.0}pt, rendering to {}x{}px at {}DPI",
            frame_w_mm, frame_h_mm, page_w_pt, page_h_pt, w, h, dpi);

        // Create memory DC and bitmap
        let screen_dc = GetDC(HWND(std::ptr::null_mut()));
        let mem_dc = CreateCompatibleDC(screen_dc);
        let bitmap = CreateCompatibleBitmap(screen_dc, w, h);
        let old_bmp = SelectObject(mem_dc, bitmap);

        // White background
        let white_brush = CreateSolidBrush(COLORREF(0x00FFFFFF));
        let rect = RECT { left: 0, top: 0, right: w, bottom: h };
        FillRect(mem_dc, &rect, white_brush);
        let _ = DeleteObject(white_brush);

        // Word CopyAsPicture EMF frame = content bounding box (text area only).
        // Frame is in 0.01mm. Convert to pt for page mapping.
        let frame_w_pt = frame_w_mm / 25.4 * 72.0;
        let frame_h_pt = frame_h_mm / 25.4 * 72.0;
        // gen_memo margins: left=90pt, top=72pt (Letter page)
        // TODO: auto-detect or accept as CLI args
        let margin_l_pt = 90.0_f64;
        let margin_t_pt = 72.0_f64;
        let scale_factor = dpi as f64 / 72.0;
        let origin_x = (margin_l_pt * scale_factor).round() as i32;
        let origin_y = (margin_t_pt * scale_factor).round() as i32;
        let content_w = (frame_w_pt * scale_factor).round() as i32;
        let content_h = (frame_h_pt * scale_factor).round() as i32;
        let play_rect = RECT {
            left: origin_x,
            top: origin_y,
            right: origin_x + content_w,
            bottom: origin_y + content_h,
        };
        eprintln!("Frame: {:.1}x{:.1}pt, play: ({},{}) -> ({},{})",
            frame_w_pt, frame_h_pt, play_rect.left, play_rect.top, play_rect.right, play_rect.bottom);
        let ok = PlayEnhMetaFile(mem_dc, hemf, &play_rect);
        if !ok.as_bool() {
            eprintln!("PlayEnhMetaFile failed");
        }

        // Extract pixels
        let mut bmi = BITMAPINFO {
            bmiHeader: BITMAPINFOHEADER {
                biSize: std::mem::size_of::<BITMAPINFOHEADER>() as u32,
                biWidth: w,
                biHeight: -h, // top-down
                biPlanes: 1,
                biBitCount: 32,
                biCompression: 0,
                ..Default::default()
            },
            ..Default::default()
        };
        let mut pixels = vec![0u8; (w * h * 4) as usize];
        GetDIBits(mem_dc, bitmap, 0, h as u32,
            Some(pixels.as_mut_ptr() as *mut _), &mut bmi, DIB_RGB_COLORS);

        // Convert BGRA to RGB
        let mut rgb = Vec::with_capacity((w * h * 3) as usize);
        for i in 0..(w * h) as usize {
            rgb.push(pixels[i * 4 + 2]); // R
            rgb.push(pixels[i * 4 + 1]); // G
            rgb.push(pixels[i * 4]);      // B
        }

        let img = image::RgbImage::from_raw(w as u32, h as u32, rgb)
            .expect("Failed to create image");
        img.save(out_path).expect("Failed to save PNG");
        eprintln!("Saved {} ({}x{})", out_path, w, h);

        // Cleanup
        SelectObject(mem_dc, old_bmp);
        let _ = DeleteObject(bitmap);
        DeleteDC(mem_dc);
        ReleaseDC(HWND(std::ptr::null_mut()), screen_dc);
        let _ = DeleteEnhMetaFile(hemf);
    }
}
