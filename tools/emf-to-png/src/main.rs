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

        // Use A4 page size (595.3 x 841.9 pt) to match Oxi full-page output
        let page_w_pt = 595.3_f64;
        let page_h_pt = 841.9_f64;
        let w = (page_w_pt * dpi as f64 / 72.0).ceil() as i32;
        let h = (page_h_pt * dpi as f64 / 72.0).ceil() as i32;
        eprintln!("EMF frame: {}x{} (0.01mm), rendering to {}x{} px at {} DPI (A4 page)", frame_w, frame_h, w, h, dpi);

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

        // Play EMF into content area (margin-based mapping).
        // CopyAsPicture EMF has ~5.6% coordinate overshoot from TextBox overhang,
        // which is a known limitation. The margin-based mapping gives SSIM=1.0
        // at the top where coordinates are most accurate.
        let margin_l = (42.55 * dpi as f64 / 72.0).round() as i32;
        let margin_t = (56.7 * dpi as f64 / 72.0).round() as i32;
        let margin_r = (42.55 * dpi as f64 / 72.0).round() as i32;
        let margin_b = (56.7 * dpi as f64 / 72.0).round() as i32;
        let content_rect = RECT {
            left: margin_l,
            top: margin_t,
            right: w - margin_r,
            bottom: h - margin_b,
        };
        let ok = PlayEnhMetaFile(mem_dc, hemf, &content_rect);
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
