//! Oxi DirectWrite/Direct2D Renderer.
//!
//! Replaces the GDI renderer (`tools/oxi-gdi-renderer`) with DirectWrite for
//! text and Direct2D for shapes/images. Goal: match Word's rendering pipeline
//! (Word uses DirectWrite + Direct2D internally) for pixel-accurate SSIM.
//!
//! Usage: oxi-dwrite-renderer input.docx output_prefix [dpi] [--exclude=...] [--supersample=N] [--dump-layout=PATH]
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

    let mut exclude: Vec<String> = Vec::new();
    // Direct2D rendering already produces grayscale ClearType-equivalent AA at
    // the target DPI, so default supersample is 1x (vs GDI's 2x). Override
    // with --supersample=N if needed for explicit comparison.
    let mut supersample: u32 = 1;
    let mut dump_layout: Option<String> = None;
    for arg in &args[3..] {
        if let Some(list) = arg.strip_prefix("--exclude=") {
            exclude = list.split(',').map(|s| s.trim().to_lowercase()).collect();
        }
        if let Some(n) = arg.strip_prefix("--supersample=") {
            supersample = n.parse().unwrap_or(1);
        }
        if let Some(path) = arg.strip_prefix("--dump-layout=") {
            dump_layout = Some(path.to_string());
        }
    }

    let data = std::fs::read(docx_path).expect("Cannot read docx file");
    let doc = oxidocs_core::parser::parse_docx(&data).expect("Cannot parse docx");

    let engine = oxidocs_core::layout::LayoutEngine::for_document(&doc);
    let result = engine.layout(&doc);

    eprintln!("Parsed {} pages, DPI={} supersample={}x (DirectWrite renderer)",
              result.pages.len(), dpi, supersample);
    if !exclude.is_empty() {
        eprintln!("Excluding: {:?}", exclude);
    }

    if let Some(ref path) = dump_layout {
        dump_layout_json(&result, path);
        eprintln!("Layout dumped to {}", path);
        return;
    }

    #[cfg(windows)]
    {
        render_pages_dwrite(&result, output_prefix, dpi, supersample, &exclude)
            .expect("DirectWrite rendering failed");
    }

    #[cfg(not(windows))]
    {
        eprintln!("DirectWrite rendering requires Windows");
        std::process::exit(1);
    }
}

#[cfg(windows)]
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

#[cfg(windows)]
fn rgb_to_d2d_color(r: u8, g: u8, b: u8) -> windows::Win32::Graphics::Direct2D::Common::D2D1_COLOR_F {
    windows::Win32::Graphics::Direct2D::Common::D2D1_COLOR_F {
        r: r as f32 / 255.0,
        g: g as f32 / 255.0,
        b: b as f32 / 255.0,
        a: 1.0,
    }
}

#[cfg(windows)]
fn render_pages_dwrite(
    result: &oxidocs_core::layout::LayoutResult,
    prefix: &str,
    dpi: u32,
    supersample: u32,
    exclude: &[String],
) -> windows::core::Result<()> {
    use windows::core::*;
    use windows::Win32::System::Com::*;
    use windows::Win32::Graphics::Direct2D::*;
    use windows::Win32::Graphics::Direct2D::Common::*;
    use windows::Win32::Graphics::DirectWrite::*;
    use windows::Win32::Graphics::Imaging::*;

    unsafe {
        // 1. Initialize COM (for WIC factory)
        CoInitializeEx(None, COINIT_MULTITHREADED).ok()?;

        // 2. Create WIC factory for PNG output
        let wic_factory: IWICImagingFactory = CoCreateInstance(
            &CLSID_WICImagingFactory,
            None,
            CLSCTX_INPROC_SERVER,
        )?;

        // 3. Create Direct2D factory
        let d2d_factory: ID2D1Factory = D2D1CreateFactory::<ID2D1Factory>(
            D2D1_FACTORY_TYPE_SINGLE_THREADED,
            None,
        )?;

        // 4. Create DirectWrite factory
        let dwrite_factory: IDWriteFactory = DWriteCreateFactory(
            DWRITE_FACTORY_TYPE_SHARED,
        )?;

        let render_dpi = dpi * supersample.max(1);
        let scale_dpi = render_dpi as f32;

        for (page_idx, page) in result.pages.iter().enumerate() {
            let out_w = (page.width as f64 * dpi as f64 / 72.0).round() as u32;
            let out_h = (page.height as f64 * dpi as f64 / 72.0).round() as u32;
            let render_w = (page.width as f64 * render_dpi as f64 / 72.0).round() as u32;
            let render_h = (page.height as f64 * render_dpi as f64 / 72.0).round() as u32;

            // 5. Create WIC bitmap (BGRA premul)
            let wic_bitmap: IWICBitmap = wic_factory.CreateBitmap(
                render_w,
                render_h,
                &GUID_WICPixelFormat32bppPBGRA,
                WICBitmapCacheOnLoad,
            )?;

            // 6. Create WIC bitmap render target
            let rt_props = D2D1_RENDER_TARGET_PROPERTIES {
                r#type: D2D1_RENDER_TARGET_TYPE_DEFAULT,
                pixelFormat: D2D1_PIXEL_FORMAT {
                    format: windows::Win32::Graphics::Dxgi::Common::DXGI_FORMAT_B8G8R8A8_UNORM,
                    alphaMode: D2D1_ALPHA_MODE_PREMULTIPLIED,
                },
                dpiX: scale_dpi,
                dpiY: scale_dpi,
                usage: D2D1_RENDER_TARGET_USAGE_NONE,
                minLevel: D2D1_FEATURE_LEVEL_DEFAULT,
            };
            let rt: ID2D1RenderTarget = d2d_factory.CreateWicBitmapRenderTarget(
                &wic_bitmap,
                &rt_props,
            )?;

            // 7. Begin draw
            rt.BeginDraw();
            rt.Clear(Some(&D2D1_COLOR_F { r: 1.0, g: 1.0, b: 1.0, a: 1.0 }));

            // TODO: render page elements (text, shapes, borders, images, clipping)
            // Step 3+ in the porting plan. For now scaffold only renders white.
            render_page_elements(&rt, &dwrite_factory, &d2d_factory, page, exclude)?;

            // 8. End draw
            rt.EndDraw(None, None)?;

            // 9. Encode WIC bitmap to PNG
            let out_path = format!("{}_p{}.png", prefix, page_idx + 1);
            save_wic_bitmap_as_png(&wic_factory, &wic_bitmap, &out_path, out_w, out_h)?;
            eprintln!("  Saved {} ({}x{})", out_path, out_w, out_h);
        }

        Ok(())
    }
}

#[cfg(windows)]
const PT_TO_DIP: f32 = 96.0 / 72.0;

#[cfg(windows)]
unsafe fn render_page_elements(
    rt: &windows::Win32::Graphics::Direct2D::ID2D1RenderTarget,
    dwrite_factory: &windows::Win32::Graphics::DirectWrite::IDWriteFactory,
    _d2d_factory: &windows::Win32::Graphics::Direct2D::ID2D1Factory,
    page: &oxidocs_core::layout::LayoutPage,
    exclude: &[String],
) -> windows::core::Result<()> {
    use oxidocs_core::layout::LayoutContent;
    use windows::Win32::Graphics::Direct2D::*;
    use windows::Win32::Graphics::Direct2D::Common::*;

    for el in &page.elements {
        match &el.content {
            LayoutContent::Text {
                text, font_size, font_family, bold, italic, color,
                underline: _, underline_style: _, strikethrough,
                double_strikethrough: _, highlight, character_spacing,
                text_scale: _, ..
            } => {
                if exclude.iter().any(|e| e == "text") { continue; }
                render_text(
                    rt, dwrite_factory,
                    el.x, el.y, el.width, el.height,
                    text, font_family.as_deref().unwrap_or("Calibri"),
                    *font_size, *bold, *italic, color.as_deref(),
                    *strikethrough, highlight.as_deref(), *character_spacing,
                )?;
            }
            LayoutContent::BoxRect { fill, stroke_color, stroke_width, corner_radius } => {
                if exclude.iter().any(|e| e == "box") { continue; }
                render_box_rect(rt, el.x, el.y, el.width, el.height,
                    fill.as_deref(), stroke_color.as_deref(), *stroke_width, *corner_radius)?;
            }
            LayoutContent::CellShading { color } => {
                if exclude.iter().any(|e| e == "shading") { continue; }
                render_filled_rect(rt, el.x, el.y, el.width, el.height, color)?;
            }
            LayoutContent::TableBorder { x1, y1, x2, y2, color, width } => {
                if exclude.iter().any(|e| e == "border") { continue; }
                render_line(rt, *x1, *y1, *x2, *y2, color.as_deref(), *width)?;
            }
            LayoutContent::ClipStart => {
                if exclude.iter().any(|e| e == "clip") { continue; }
                let r = D2D_RECT_F {
                    left:   el.x * PT_TO_DIP,
                    top:    el.y * PT_TO_DIP,
                    right:  (el.x + el.width)  * PT_TO_DIP,
                    bottom: (el.y + el.height) * PT_TO_DIP,
                };
                rt.PushAxisAlignedClip(&r, D2D1_ANTIALIAS_MODE_PER_PRIMITIVE);
            }
            LayoutContent::ClipEnd => {
                if exclude.iter().any(|e| e == "clip") { continue; }
                rt.PopAxisAlignedClip();
            }
            LayoutContent::Image { ref data, .. } => {
                if exclude.iter().any(|e| e == "image") { continue; }
                if !data.is_empty() {
                    render_image(rt, &page_factory_handle(rt)?, el.x, el.y, el.width, el.height, data)?;
                }
            }
            LayoutContent::PresetShape { shape_type, stroke_color, stroke_width } => {
                render_preset_shape(rt, el.x, el.y, el.width, el.height,
                    shape_type, stroke_color.as_deref(), *stroke_width)?;
            }
            // TODO Step 7: Balloon, BalloonConnector (low priority for SSIM)
            _ => {}
        }
    }
    Ok(())
}

#[cfg(windows)]
fn page_factory_handle(
    rt: &windows::Win32::Graphics::Direct2D::ID2D1RenderTarget,
) -> windows::core::Result<windows::Win32::Graphics::Direct2D::ID2D1Factory> {
    unsafe { rt.GetFactory() }
}

#[cfg(windows)]
unsafe fn render_image(
    rt: &windows::Win32::Graphics::Direct2D::ID2D1RenderTarget,
    _d2d_factory: &windows::Win32::Graphics::Direct2D::ID2D1Factory,
    x_pt: f32, y_pt: f32, w_pt: f32, h_pt: f32,
    data: &[u8],
) -> windows::core::Result<()> {
    use windows::Win32::Graphics::Direct2D::*;
    use windows::Win32::Graphics::Direct2D::Common::*;

    // Decode via image crate (jpg/png/etc.)
    let img = match image::load_from_memory(data) {
        Ok(i) => i,
        Err(_) => return Ok(()),
    };
    let rgba = img.to_rgba8();
    let (pw, ph) = rgba.dimensions();

    // BGRA premul for D2D
    let mut bgra: Vec<u8> = Vec::with_capacity((pw * ph * 4) as usize);
    for px in rgba.pixels() {
        let a = px[3];
        let pre = |c: u8| ((c as u16 * a as u16 + 127) / 255) as u8;
        bgra.push(pre(px[2])); // B
        bgra.push(pre(px[1])); // G
        bgra.push(pre(px[0])); // R
        bgra.push(a);
    }

    let bmp_props = D2D1_BITMAP_PROPERTIES {
        pixelFormat: D2D1_PIXEL_FORMAT {
            format: windows::Win32::Graphics::Dxgi::Common::DXGI_FORMAT_B8G8R8A8_UNORM,
            alphaMode: D2D1_ALPHA_MODE_PREMULTIPLIED,
        },
        dpiX: 96.0,
        dpiY: 96.0,
    };
    let size = D2D_SIZE_U { width: pw, height: ph };
    let bitmap = rt.CreateBitmap(size, Some(bgra.as_ptr() as *const _), pw * 4, &bmp_props)?;

    let dest = D2D_RECT_F {
        left:   x_pt * PT_TO_DIP,
        top:    y_pt * PT_TO_DIP,
        right:  (x_pt + w_pt) * PT_TO_DIP,
        bottom: (y_pt + h_pt) * PT_TO_DIP,
    };
    rt.DrawBitmap(&bitmap, Some(&dest), 1.0,
        D2D1_BITMAP_INTERPOLATION_MODE_LINEAR, None);
    Ok(())
}

#[cfg(windows)]
unsafe fn render_preset_shape(
    rt: &windows::Win32::Graphics::Direct2D::ID2D1RenderTarget,
    x_pt: f32, y_pt: f32, w_pt: f32, h_pt: f32,
    shape_type: &str,
    stroke_color: Option<&str>,
    stroke_width: f32,
) -> windows::core::Result<()> {
    use windows::Win32::Graphics::Direct2D::Common::*;

    let (r, g, b) = stroke_color.map(parse_hex_rgb).unwrap_or((0, 0, 0));
    let brush = rt.CreateSolidColorBrush(&rgb_to_d2d_color(r, g, b), None)?;
    let sw_dip = stroke_width.max(0.5) * PT_TO_DIP;

    let x = x_pt * PT_TO_DIP;
    let y = y_pt * PT_TO_DIP;
    let w = w_pt * PT_TO_DIP;
    let h = h_pt * PT_TO_DIP;

    match shape_type {
        "line" => {
            // Diagonal line top-left to bottom-right
            rt.DrawLine(
                D2D_POINT_2F { x, y },
                D2D_POINT_2F { x: x + w, y: y + h },
                &brush, sw_dip, None,
            );
        }
        // bracketPair, leftBracket, rightBracket: bezier-curved brackets.
        // For first cut, approximate with simple lines (3% radius).
        // Pixel-perfect bezier port can be Step 5b.
        "bracketPair" | "leftBracket" | "rightBracket" => {
            let radius = (w.min(h) * 0.08387).max(2.0);
            if shape_type == "bracketPair" || shape_type == "leftBracket" {
                rt.DrawLine(D2D_POINT_2F { x: x + radius, y },
                    D2D_POINT_2F { x, y: y + radius }, &brush, sw_dip, None);
                rt.DrawLine(D2D_POINT_2F { x, y: y + radius },
                    D2D_POINT_2F { x, y: y + h - radius }, &brush, sw_dip, None);
                rt.DrawLine(D2D_POINT_2F { x, y: y + h - radius },
                    D2D_POINT_2F { x: x + radius, y: y + h }, &brush, sw_dip, None);
            }
            if shape_type == "bracketPair" || shape_type == "rightBracket" {
                rt.DrawLine(D2D_POINT_2F { x: x + w - radius, y },
                    D2D_POINT_2F { x: x + w, y: y + radius }, &brush, sw_dip, None);
                rt.DrawLine(D2D_POINT_2F { x: x + w, y: y + radius },
                    D2D_POINT_2F { x: x + w, y: y + h - radius }, &brush, sw_dip, None);
                rt.DrawLine(D2D_POINT_2F { x: x + w, y: y + h - radius },
                    D2D_POINT_2F { x: x + w - radius, y: y + h }, &brush, sw_dip, None);
            }
        }
        _ => {}
    }
    Ok(())
}

#[cfg(windows)]
unsafe fn render_filled_rect(
    rt: &windows::Win32::Graphics::Direct2D::ID2D1RenderTarget,
    x_pt: f32, y_pt: f32, w_pt: f32, h_pt: f32,
    color: &str,
) -> windows::core::Result<()> {
    use windows::Win32::Graphics::Direct2D::Common::*;
    let (r, g, b) = parse_hex_rgb(color);
    let brush = rt.CreateSolidColorBrush(&rgb_to_d2d_color(r, g, b), None)?;
    let rect = D2D_RECT_F {
        left:   x_pt * PT_TO_DIP,
        top:    y_pt * PT_TO_DIP,
        right:  (x_pt + w_pt) * PT_TO_DIP,
        bottom: (y_pt + h_pt) * PT_TO_DIP,
    };
    rt.FillRectangle(&rect, &brush);
    Ok(())
}

#[cfg(windows)]
unsafe fn render_box_rect(
    rt: &windows::Win32::Graphics::Direct2D::ID2D1RenderTarget,
    x_pt: f32, y_pt: f32, w_pt: f32, h_pt: f32,
    fill: Option<&str>,
    stroke_color: Option<&str>,
    stroke_width: f32,
    corner_radius: f32,
) -> windows::core::Result<()> {
    use windows::Win32::Graphics::Direct2D::*;
    use windows::Win32::Graphics::Direct2D::Common::*;

    // White-with-no-border: skip (matches GDI optimization).
    let is_invisible_white = match fill {
        Some(f) => {
            let c = f.strip_prefix('#').unwrap_or(f);
            c.eq_ignore_ascii_case("ffffff") || c.eq_ignore_ascii_case("fff")
        }
        None => false,
    } && stroke_color.is_none();
    if is_invisible_white {
        return Ok(());
    }

    let rect = D2D_RECT_F {
        left:   x_pt * PT_TO_DIP,
        top:    y_pt * PT_TO_DIP,
        right:  (x_pt + w_pt) * PT_TO_DIP,
        bottom: (y_pt + h_pt) * PT_TO_DIP,
    };
    let cr_dip = corner_radius * PT_TO_DIP;

    // Fill
    if let Some(fill_hex) = fill {
        let (r, g, b) = parse_hex_rgb(fill_hex);
        let brush = rt.CreateSolidColorBrush(&rgb_to_d2d_color(r, g, b), None)?;
        if cr_dip > 0.0 {
            let rounded = D2D1_ROUNDED_RECT { rect, radiusX: cr_dip, radiusY: cr_dip };
            rt.FillRoundedRectangle(&rounded, &brush);
        } else {
            rt.FillRectangle(&rect, &brush);
        }
    }
    // Stroke
    if let Some(stroke) = stroke_color {
        let (r, g, b) = parse_hex_rgb(stroke);
        let brush = rt.CreateSolidColorBrush(&rgb_to_d2d_color(r, g, b), None)?;
        let sw_dip = stroke_width * PT_TO_DIP;
        if cr_dip > 0.0 {
            let rounded = D2D1_ROUNDED_RECT { rect, radiusX: cr_dip, radiusY: cr_dip };
            rt.DrawRoundedRectangle(&rounded, &brush, sw_dip, None);
        } else {
            rt.DrawRectangle(&rect, &brush, sw_dip, None);
        }
    }
    Ok(())
}

#[cfg(windows)]
unsafe fn render_line(
    rt: &windows::Win32::Graphics::Direct2D::ID2D1RenderTarget,
    x1_pt: f32, y1_pt: f32, x2_pt: f32, y2_pt: f32,
    color: Option<&str>,
    width_pt: f32,
) -> windows::core::Result<()> {
    use windows::Win32::Graphics::Direct2D::Common::*;
    let (r, g, b) = color.map(parse_hex_rgb).unwrap_or((0, 0, 0));
    let brush = rt.CreateSolidColorBrush(&rgb_to_d2d_color(r, g, b), None)?;
    rt.DrawLine(
        D2D_POINT_2F { x: x1_pt * PT_TO_DIP, y: y1_pt * PT_TO_DIP },
        D2D_POINT_2F { x: x2_pt * PT_TO_DIP, y: y2_pt * PT_TO_DIP },
        &brush,
        width_pt * PT_TO_DIP,
        None,
    );
    Ok(())
}

#[cfg(windows)]
unsafe fn render_text(
    rt: &windows::Win32::Graphics::Direct2D::ID2D1RenderTarget,
    dwrite_factory: &windows::Win32::Graphics::DirectWrite::IDWriteFactory,
    x_pt: f32, y_pt: f32, w_pt: f32, h_pt: f32,
    text: &str,
    font_family: &str,
    font_size_pt: f32,
    bold: bool,
    italic: bool,
    color: Option<&str>,
    strikethrough: bool,
    highlight: Option<&str>,
    character_spacing_pt: f32,
) -> windows::core::Result<()> {
    use windows::core::*;
    use windows::Win32::Graphics::Direct2D::*;
    use windows::Win32::Graphics::Direct2D::Common::*;
    use windows::Win32::Graphics::DirectWrite::*;

    if text.is_empty() {
        return Ok(());
    }

    // Highlight background covers the element rect before glyphs are drawn.
    if let Some(hl) = highlight {
        let (hr, hg, hb) = parse_hex_rgb(hl);
        let hl_brush = rt.CreateSolidColorBrush(&rgb_to_d2d_color(hr, hg, hb), None)?;
        let rect = D2D_RECT_F {
            left:   x_pt * PT_TO_DIP,
            top:    y_pt * PT_TO_DIP,
            right:  (x_pt + w_pt) * PT_TO_DIP,
            bottom: (y_pt + h_pt) * PT_TO_DIP,
        };
        rt.FillRectangle(&rect, &hl_brush);
    }

    // Create IDWriteTextFormat. font_size in DIPs at 96 DPI standard.
    let weight = if bold { DWRITE_FONT_WEIGHT_BOLD } else { DWRITE_FONT_WEIGHT_NORMAL };
    let style = if italic { DWRITE_FONT_STYLE_ITALIC } else { DWRITE_FONT_STYLE_NORMAL };
    let family_wide: Vec<u16> = font_family.encode_utf16()
        .chain(std::iter::once(0)).collect();
    let locale_wide: Vec<u16> = "ja-jp".encode_utf16()
        .chain(std::iter::once(0)).collect();

    let format: IDWriteTextFormat = dwrite_factory.CreateTextFormat(
        PCWSTR(family_wide.as_ptr()),
        None,
        weight,
        style,
        DWRITE_FONT_STRETCH_NORMAL,
        font_size_pt * PT_TO_DIP,
        PCWSTR(locale_wide.as_ptr()),
    )?;

    // Disable wrapping for per-char elements (each char is its own element)
    format.SetWordWrapping(DWRITE_WORD_WRAPPING_NO_WRAP)?;

    // Create solid color brush
    let (r, g, b) = color.map(parse_hex_rgb).unwrap_or((0, 0, 0));
    let brush_color = rgb_to_d2d_color(r, g, b);
    let brush: ID2D1SolidColorBrush = rt.CreateSolidColorBrush(&brush_color, None)?;

    // Convert UTF-16 for DirectWrite
    let text_wide: Vec<u16> = text.encode_utf16().collect();

    // Create text layout. Width/height generous enough so text doesn't clip.
    // Use 10000 DIPs ≈ 7500pt to avoid wrapping or clipping at element bounds.
    let _ = (w_pt, h_pt); // currently unused; element bounds may be tight
    let layout_w_dip = 10000.0_f32;
    let layout_h_dip = (font_size_pt * 2.0) * PT_TO_DIP;

    let layout: IDWriteTextLayout = dwrite_factory.CreateTextLayout(
        &text_wide,
        &format,
        layout_w_dip,
        layout_h_dip,
    )?;

    // OOXML w:spacing val="N" gives extra inter-glyph advance in 20ths of a
    // point. Oxi's IR converts that to points before reaching the renderer.
    // GDI applies it via SetTextCharacterExtra (one int trailing pixels per
    // char). DirectWrite's equivalent is IDWriteTextLayout1::SetCharacterSpacing
    // (leading, trailing, minAdvance, range) — putting the whole offset on
    // trailing matches GDI.
    if character_spacing_pt.abs() > 0.001 {
        if let Ok(layout1) = layout.cast::<IDWriteTextLayout1>() {
            let trailing = character_spacing_pt * PT_TO_DIP;
            let range = DWRITE_TEXT_RANGE { startPosition: 0, length: text_wide.len() as u32 };
            let _ = layout1.SetCharacterSpacing(0.0, trailing, 0.0, range);
        }
    }

    let origin = D2D_POINT_2F {
        x: x_pt * PT_TO_DIP,
        y: y_pt * PT_TO_DIP,
    };

    rt.DrawTextLayout(
        origin,
        &layout,
        &brush,
        D2D1_DRAW_TEXT_OPTIONS_NONE,
    );

    // Strikethrough drawn manually as a horizontal line spanning the element
    // width — matches GDI's MoveToEx/LineTo path. DirectWrite's SetStrikethrough
    // skips whitespace, which breaks Oxi's per-char element model. (Underline
    // intentionally not yet ported — both SetUnderline and DrawLine variants
    // regress -0.18 to -0.20 net p.1 SSIM on the 19 underline-containing
    // baseline docs; see pipeline_data/dwrite_underline_investigation_2026-05-02.md.)
    if strikethrough {
        let mut count: u32 = 0;
        let _ = layout.GetLineMetrics(None, &mut count);
        let baseline_dip = if count > 0 {
            let mut metrics = vec![DWRITE_LINE_METRICS::default(); count as usize];
            if layout.GetLineMetrics(Some(&mut metrics), &mut count).is_ok() {
                metrics[0].baseline
            } else { font_size_pt * PT_TO_DIP * 0.8 }
        } else { font_size_pt * PT_TO_DIP * 0.8 };
        let thickness = (font_size_pt * PT_TO_DIP * 0.06).max(1.0);
        let y_dip = y_pt * PT_TO_DIP + baseline_dip - baseline_dip * 0.30;
        rt.DrawLine(
            D2D_POINT_2F { x: x_pt * PT_TO_DIP, y: y_dip },
            D2D_POINT_2F { x: (x_pt + w_pt) * PT_TO_DIP, y: y_dip },
            &brush, thickness, None,
        );
    }

    Ok(())
}

#[cfg(windows)]
unsafe fn save_wic_bitmap_as_png(
    wic_factory: &windows::Win32::Graphics::Imaging::IWICImagingFactory,
    src_bitmap: &windows::Win32::Graphics::Imaging::IWICBitmap,
    path: &str,
    out_w: u32,
    out_h: u32,
) -> windows::core::Result<()> {
    use windows::core::*;
    use windows::Win32::Graphics::Imaging::*;

    // If supersampled, downscale via WIC scaler to output dimensions
    let mut src_w: u32 = 0;
    let mut src_h: u32 = 0;
    src_bitmap.GetSize(&mut src_w, &mut src_h)?;

    let source: IWICBitmapSource = if src_w != out_w || src_h != out_h {
        let scaler = wic_factory.CreateBitmapScaler()?;
        scaler.Initialize(
            src_bitmap,
            out_w,
            out_h,
            WICBitmapInterpolationModeFant,
        )?;
        scaler.cast()?
    } else {
        src_bitmap.cast()?
    };

    // Create stream — use Windows-style path separators
    let win_path: String = path.replace('/', "\\");
    let stream: IWICStream = wic_factory.CreateStream()?;
    let path_wide: Vec<u16> = win_path.encode_utf16().chain(std::iter::once(0)).collect();
    // GENERIC_WRITE = 0x40000000 (no need for full FILE_GENERIC_WRITE mask)
    stream.InitializeFromFilename(
        PCWSTR(path_wide.as_ptr()),
        0x40000000,
    )?;

    // Create PNG encoder
    let encoder: IWICBitmapEncoder = wic_factory.CreateEncoder(
        &GUID_ContainerFormatPng,
        std::ptr::null(),
    )?;
    encoder.Initialize(&stream, WICBitmapEncoderNoCache)?;

    let mut frame: Option<IWICBitmapFrameEncode> = None;
    let mut bag: Option<windows::Win32::System::Com::StructuredStorage::IPropertyBag2> = None;
    encoder.CreateNewFrame(&mut frame, &mut bag)?;
    let frame = frame.unwrap();
    frame.Initialize(bag.as_ref())?;
    frame.SetSize(out_w, out_h)?;

    // Use 32bpp BGRA pixel format (PNG supports it)
    let mut pixel_format = GUID_WICPixelFormat32bppBGRA;
    frame.SetPixelFormat(&mut pixel_format)?;

    frame.WriteSource(&source, std::ptr::null())?;
    frame.Commit()?;
    encoder.Commit()?;

    Ok(())
}

fn dump_layout_json(result: &oxidocs_core::layout::LayoutResult, path: &str) {
    // Reuse format from oxi-gdi-renderer for compatibility
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
            write!(&mut out,
                "      {{\"type\": \"{}\", \"x\": {:.3}, \"y\": {:.3}, \"w\": {:.3}, \"h\": {:.3}, \"text\": {}, \"font_size\": {:.2}, \"para_idx\": {}, \"run_idx\": {}, \"char_offset\": {}}}",
                kind, el.x, el.y, el.width, el.height, text_json, font_size, pi_json, ri_json, co_json).unwrap();
        }
        out.push_str("\n    ]}");
    }
    out.push_str("\n  ]\n}\n");
    std::fs::write(path, out).expect("Failed to write dump-layout JSON");
}
