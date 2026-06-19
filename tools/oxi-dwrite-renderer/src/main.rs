//! Oxi DirectWrite/Direct2D Renderer.
//!
//! Replaces the GDI renderer (`tools/oxi-gdi-renderer`) with DirectWrite for
//! text and Direct2D for shapes/images. Goal: match Word's rendering pipeline
//! (Word uses DirectWrite + Direct2D internally) for pixel-accurate SSIM.
//!
//! Usage: oxi-dwrite-renderer input.docx output_prefix [dpi] [--exclude=...] [--supersample=N] [--dump-layout=PATH]
//!   Produces: output_prefix_p1.png, output_prefix_p2.png, ...

use std::path::Path;

// S494: per-glyph dump buffer (thread-local to avoid threading through render_text's
// many params). Each page: (width_pt, height_pt, Vec<(char, x_pt, top_pt, baseline_pt, fs_pt, family)>).
// baseline_pt (S494 gate fix) = the EXACT y where DWrite draws the glyph origin
// (top + the font's DWrite ascent via GetLineMetrics). Emitting the baseline directly
// lets the per-glyph gate place glyphs without guessing a per-font ascent K (the fixed
// K=0.859 systematically mis-placed non-Mincho/Latin baselines, e.g. Calibri winAscent
// 0.952 — the cause of the spurious Latin-doc "drops" in the first per-glyph baseline).
thread_local! {
    static GLYPH_DUMP: std::cell::RefCell<Option<Vec<(f32, f32, Vec<(char, f32, f32, f32, f32, String)>)>>>
        = const { std::cell::RefCell::new(None) };
}

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
    // S493d (2026-06-04) — default supersample 1->2. Even with the S460 OUTLINE
    // grayscale AA, native-res (1x) leaves the dense-CJK glyph cores HARDER than
    // Word's ClearType-screenshot grayscale (render-truth confirmed 0e7af positions
    // are perfect, so the residual is weight/AA texture). 2x supersample softens the
    // coverage AA toward Word: full-corpus ss1->ss2 canary (235 docs, RGB SSIM gate)
    // bottom-10 sum +0.0375 (strictly up = Phase-3 primary), mean +0.0009, 10 docs
    // improve >0.005, only d77a regresses (-0.008, non-bottom-10). 0e7af doc-mean
    // 0.8758->0.8907 (+0.0149; p2 +0.028, p7 +0.033). Render-only (positions/pagination
    // unchanged -> Phase-1 safe). 4x render cost (was deliberately avoided pre-S493 but
    // the bottom-N gain is real). Override with --supersample=1 for the legacy fast path.
    // S501 SHIP (2026-06-07, default supersample 2->3, user-approved the 3x verify cost) —
    // ss3 softens the coverage AA further toward Word's smooth grayscale. PASSES Phase-3
    // bottom-N: bottom-20 SUM ss2 16.3152 -> ss3 16.3330 (+0.0178, 15 up / 5 down). Weight-
    // capped gains (0e7af +0.0113, 683f +0.0056, d77a +0.0036, 1ec1 +0.0033, 15076 +0.0023)
    // outweigh position-capped regressions (b35 -0.0054 over-softens its mispositioned
    // glyphs; 4 others). ss4 = no further gain (16x cost). Render-only (Phase-1 safe) and
    // verify-pipeline-only (browser product renders at native res, unaffected). COST: 9x
    // render (vs ss2 4x) -> every verify/gate run ~3x (full baseline refresh ~90 min);
    // accepted. --supersample=1/2 overrides for a fast verify.
    // S622 (2026-06-19): env override OXI_SUPERSAMPLE for A/B testing (the
    // --supersample flag still wins if given). Default 3 (S501).
    let mut supersample: u32 = std::env::var("OXI_SUPERSAMPLE")
        .ok().and_then(|v| v.parse().ok()).unwrap_or(3);
    let mut dump_layout: Option<String> = None;
    // S494: --dump-glyphs=PATH emits each glyph's EXACT per-char position from
    // DirectWrite (IDWriteTextLayout::HitTestTextPosition), which includes DWrite's
    // CJK<->half-width autoSpace AND charGrid stretch — the gate-render positions
    // (dwrite ≈ Word). Faithful for ALL docs incl. tables (unlike GDI dump-glyphs).
    let mut dump_glyphs: Option<String> = None;
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
        if let Some(path) = arg.strip_prefix("--dump-glyphs=") {
            dump_glyphs = Some(path.to_string());
        }
    }
    if dump_glyphs.is_some() {
        GLYPH_DUMP.with(|g| *g.borrow_mut() = Some(Vec::new()));
    }

    let data = std::fs::read(docx_path).expect("Cannot read docx file");
    let doc = oxidocs_core::parser::parse_docx(&data).expect("Cannot parse docx");

    // S483: render in Word's "final" view (accept revisions: hide <w:del>
    // content, show <w:ins> as normal text) to match the Word ground-truth
    // PNGs, which are captured in final/clean view. Oxi defaulted to
    // ShowRevisions::All (deletion markup shown). Opt-out OXI_S483_DISABLE.
    let engine = oxidocs_core::layout::LayoutEngine::for_document(&doc);
    let engine = if std::env::var("OXI_S483_DISABLE").is_ok() {
        engine
    } else {
        engine.with_show_revisions(oxidocs_core::ir::ShowRevisions::Final)
    };
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

    if let Some(path) = dump_glyphs {
        GLYPH_DUMP.with(|g| {
            if let Some(pages) = g.borrow().as_ref() {
                dump_glyphs_json(pages, &path);
                eprintln!("Glyphs dumped to {}", path);
            }
        });
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
            // S494: open a new page entry in the glyph-dump buffer (if dumping).
            GLYPH_DUMP.with(|g| {
                if let Some(v) = g.borrow_mut().as_mut() {
                    v.push((page.width, page.height, Vec::new()));
                }
            });
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

            // S460 (2026-05-31) ★ SHIP — the default WIC bitmap text rendering
            // produced a coarse GDI-compatible 4-bit (17-level) grayscale alpha
            // with SHARP edges, while Word's EMF output is smooth 255-level AA
            // (measured: Oxi 17 unique gray values vs Word 255 on 0e7af p2). On
            // dense small CJK body pages this edge mismatch capped SSIM ~0.7.
            // The prior CLAUDE.md premise "Direct2D produces grayscale ClearType-
            // equivalent AA so 1x supersample matches Word" was never pixel-
            // verified and is FALSE. Fix: set GRAYSCALE antialias + an OUTLINE
            // custom rendering params (unhinted pure-coverage AA) → smooth
            // 256-level AA at NATIVE resolution (no supersample perf cost).
            // outline beat natural (9 lvl, hinted, no gain) and natural_sym
            // (17 lvl, +0.044 on 0e7af) on 0e7af; gamma has no effect on outline.
            // GATE (full 410-page recompute): mean 0.9067→0.9079 (+0.0011),
            // bottom-3 sum +0.0019 / bottom-10 +0.0239, <0.70 8→7, ≥0.95
            // 174→176; 0e7af doc-mean 0.821→0.876 (p6 +0.119, p2 +0.087), only
            // 4 pages regress >0.005 (worst LOD p11 −0.0099, MS Gothic code).
            // Rendering-only (no layout change) → element.y / pagination / Phase-1
            // unaffected. Opt out / override via OXI_S460_RMODE ("off"/"legacy"
            // disables; "natural"/"natural_sym"/"aliased" alternatives);
            // OXI_S460_GAMMA overrides gamma (no-op for outline).
            let s460_mode = std::env::var("OXI_S460_RMODE")
                .unwrap_or_else(|_| "outline".to_string());
            if s460_mode != "off" && s460_mode != "legacy" {
                rt.SetTextAntialiasMode(D2D1_TEXT_ANTIALIAS_MODE_GRAYSCALE);
                let gamma: f32 = std::env::var("OXI_S460_GAMMA").ok()
                    .and_then(|v| v.parse().ok()).unwrap_or(1.8);
                let rmode = match s460_mode.as_str() {
                    "natural" => DWRITE_RENDERING_MODE_NATURAL,
                    "natural_sym" => DWRITE_RENDERING_MODE_NATURAL_SYMMETRIC,
                    "aliased" => DWRITE_RENDERING_MODE_ALIASED,
                    _ => DWRITE_RENDERING_MODE_OUTLINE,
                };
                if let Ok(params) = dwrite_factory.CreateCustomRenderingParams(
                    gamma, 0.0, 0.0,
                    DWRITE_PIXEL_GEOMETRY_FLAT,
                    rmode,
                ) {
                    rt.SetTextRenderingParams(&params);
                }
            }

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

    // S518 (2026-06-09): per-line BODY reference (font/size) for aligning a Symbol-font
    // list marker to its line's body baseline. The Symbol bullet () renders in the
    // Symbol font whose DWrite ascent is ~2pt smaller than the body font at fs11, so
    // (y+off-1+symbol_ascent) lands ~2pt ABOVE the body — runs on one line don't share a
    // baseline (DrawTextLayout uses each font's own ascent). Word draws the bullet on the
    // line baseline. SCOPED to Symbol-font markers (paragraph_index None + family "Symbol")
    // so b837 number markers (S517, body font) and / bullets (body font) are untouched.
    // Key by line-box top (el.y) rounded to 0.25pt; first body text element wins.
    let mut line_body_ref: std::collections::HashMap<i32, (f32, String, f32, bool, bool)> =
        std::collections::HashMap::new();
    for el in &page.elements {
        if let LayoutContent::Text { font_size, font_family, bold, italic, .. } = &el.content {
            if el.paragraph_index.is_some() {
                let key = (el.y * 4.0).round() as i32;
                line_body_ref.entry(key).or_insert_with(|| (
                    el.y + el.text_y_off,
                    font_family.as_deref().unwrap_or("Calibri").to_string(),
                    *font_size, *bold, *italic,
                ));
            }
        }
    }

    for el in &page.elements {
        match &el.content {
            LayoutContent::Text {
                text, font_size, font_family, bold, italic, color,
                underline: _, underline_style, strikethrough, double_strikethrough,
                highlight, character_spacing,
                text_scale, is_vertical, ..
            } => {
                if exclude.iter().any(|e| e == "text") { continue; }
                // Session 75 Phase D (2026-05-17): el.y is LINE BOX TOP; pass
                // el.y + el.text_y_off as the glyph-top y to preserve pre-Phase-D
                // pixel positions. See memory/session71_y_convention_refactor_design.md.
                let is_double_underline = underline_style.as_deref() == Some("double");
                // S518: align a Symbol-font marker to its line's body baseline.
                let mut glyph_top_y = el.y + el.text_y_off;
                let fam = font_family.as_deref().unwrap_or("Calibri");
                if el.paragraph_index.is_none() && fam == "Symbol"
                    && std::env::var("OXI_S518_BULLET_DISABLE").is_err()
                {
                    if let Some((body_top, body_fam, body_fs, body_b, body_i)) =
                        line_body_ref.get(&((el.y * 4.0).round() as i32))
                    {
                        let body_asc = font_ascent_pt(dwrite_factory, body_fam, *body_fs, *body_b, *body_i);
                        let mark_asc = font_ascent_pt(dwrite_factory, fam, *font_size, *bold, *italic);
                        // marker baseline := body baseline (both share el.y+text_y_off after S517)
                        glyph_top_y = *body_top + (body_asc - mark_asc);
                    }
                }
                // S546 (2026-06-12): FALSIFIED full-corpus — naive element-x
                // 96dpi snap (round(x/0.75)*0.75) regressed 216/410 pages net
                // −1.1369 (the horizontal analog of S468 VSNAP). Word snaps the
                // TRUE cumulative per-char x; Oxi element x already carries
                // compression/justify adjustments, so snapping injects ±0.375
                // noise instead of aligning. Kept opt-IN for experiments.
                let snap_x = if std::env::var("OXI_S546_XSNAP").is_ok() {
                    (el.x / 0.75).round() * 0.75
                } else { el.x };
                render_text(
                    rt, dwrite_factory,
                    snap_x, glyph_top_y, el.width, el.height,
                    text, fam,
                    *font_size, *bold, *italic, color.as_deref(),
                    *strikethrough, *double_strikethrough, is_double_underline,
                    highlight.as_deref(), *character_spacing,
                    *is_vertical, *text_scale,
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
            LayoutContent::TableBorder { x1, y1, x2, y2, color, width, style } => {
                if exclude.iter().any(|e| e == "border") { continue; }
                render_line(rt, *x1, *y1, *x2, *y2, color.as_deref(), *width, style.as_deref())?;
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
            LayoutContent::PresetShape { shape_type, stroke_color, stroke_width, flip_h, flip_v, arrow_head, arrow_tail } => {
                render_preset_shape(rt, el.x, el.y, el.width, el.height,
                    shape_type, stroke_color.as_deref(), *stroke_width, *flip_h, *flip_v, *arrow_head, *arrow_tail)?;
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
    flip_h: bool,
    flip_v: bool,
    arrow_head: bool,
    arrow_tail: bool,
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
        // S493h: connector shapes — flipH/flipV-aware endpoints (the S490 attempt drew a
        // fixed TL→BR diagonal, wrong when flipped → regressed 3a4f). The connector runs
        // between two diagonal corners of its bbox; flipH/flipV pick which corners.
        //   start = (flipH? x+w : x, flipV? y+h : y);  end = opposite corner.
        // straightConnector1 = the diagonal; bentConnector3 = an H–V–H Z through the bend.
        // (3a4f is Phase-1-broken — gate on 2ea81a/1636d28, not 3a4f.)
        "line" | "straightConnector1" | "bentConnector3" | "bentConnector2" => {
            let sx = if flip_h { x + w } else { x };
            let sy = if flip_v { y + h } else { y };
            let ex = if flip_h { x } else { x + w };
            let ey = if flip_v { y } else { y + h };
            // arrowhead anchors: (tip, point the segment comes FROM)
            let (tail_from_x, tail_from_y, head_from_x, head_from_y);
            if shape_type == "bentConnector3" || shape_type == "bentConnector2" {
                // Z/L bend: horizontal to mid-x, vertical, horizontal to end.
                let mx = (sx + ex) * 0.5;
                rt.DrawLine(D2D_POINT_2F { x: sx, y: sy }, D2D_POINT_2F { x: mx, y: sy }, &brush, sw_dip, None);
                rt.DrawLine(D2D_POINT_2F { x: mx, y: sy }, D2D_POINT_2F { x: mx, y: ey }, &brush, sw_dip, None);
                rt.DrawLine(D2D_POINT_2F { x: mx, y: ey }, D2D_POINT_2F { x: ex, y: ey }, &brush, sw_dip, None);
                tail_from_x = mx; tail_from_y = ey; head_from_x = mx; head_from_y = sy;
            } else {
                rt.DrawLine(D2D_POINT_2F { x: sx, y: sy }, D2D_POINT_2F { x: ex, y: ey }, &brush, sw_dip, None);
                tail_from_x = sx; tail_from_y = sy; head_from_x = ex; head_from_y = ey;
            }
            // S493i: filled-triangle arrowheads (a:tailEnd at end, a:headEnd at start).
            if arrow_tail {
                draw_arrowhead(rt, ex, ey, ex - tail_from_x, ey - tail_from_y, sw_dip, &brush)?;
            }
            if arrow_head {
                draw_arrowhead(rt, sx, sy, sx - head_from_x, sy - head_from_y, sw_dip, &brush)?;
            }
        }
        // S490: ellipse/oval outline (e.g. ○ option markers circling a choice).
        // Was unhandled → Word drew the ring, Oxi drew nothing. Stroke-only
        // (render_preset_shape has no fill; these markers are noFill rings).
        "ellipse" | "oval" => {
            use windows::Win32::Graphics::Direct2D::D2D1_ELLIPSE;
            let ellipse = D2D1_ELLIPSE {
                point: D2D_POINT_2F { x: x + w * 0.5, y: y + h * 0.5 },
                radiusX: w * 0.5,
                radiusY: h * 0.5,
            };
            rt.DrawEllipse(&ellipse, &brush, sw_dip, None);
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

/// S493i: draw a filled-triangle connector arrowhead. Tip at (tx,ty); (dx,dy) is the line
/// direction INTO the tip. Size scales with the stroke width (Word's "med" triangle ≈ a few
/// line-widths). DIP coordinates.
#[cfg(windows)]
unsafe fn draw_arrowhead(
    rt: &windows::Win32::Graphics::Direct2D::ID2D1RenderTarget,
    tx: f32, ty: f32, dx: f32, dy: f32, sw_dip: f32,
    brush: &windows::Win32::Graphics::Direct2D::ID2D1SolidColorBrush,
) -> windows::core::Result<()> {
    use windows::Win32::Graphics::Direct2D::*;
    use windows::Win32::Graphics::Direct2D::Common::*;
    let len = (dx * dx + dy * dy).sqrt();
    if len < 0.01 { return Ok(()); }
    let (ux, uy) = (dx / len, dy / len);   // unit direction into the tip
    let (px, py) = (-uy, ux);              // perpendicular
    let alen = (sw_dip * 3.0).max(5.0);    // arrowhead length
    let ahw = (sw_dip * 1.8).max(3.0);     // arrowhead half-width
    let (bx, by) = (tx - ux * alen, ty - uy * alen); // base centre
    let tip = D2D_POINT_2F { x: tx, y: ty };
    let p1 = D2D_POINT_2F { x: bx + px * ahw, y: by + py * ahw };
    let p2 = D2D_POINT_2F { x: bx - px * ahw, y: by - py * ahw };
    let factory = rt.GetFactory()?;
    let geom: ID2D1PathGeometry = factory.CreatePathGeometry()?;
    let sink = geom.Open()?;
    sink.BeginFigure(tip, D2D1_FIGURE_BEGIN_FILLED);
    sink.AddLine(p1);
    sink.AddLine(p2);
    sink.EndFigure(D2D1_FIGURE_END_CLOSED);
    sink.Close()?;
    rt.FillGeometry(&geom, brush, None);
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

// S480: map an OOXML border art style (w:val) to a D2D1 dash style.
// None = solid stroke. OXI_S480_DISABLE forces solid (pre-S480 behavior).
#[cfg(windows)]
fn s480_dash_style(style: Option<&str>)
    -> Option<windows::Win32::Graphics::Direct2D::D2D1_DASH_STYLE>
{
    use windows::Win32::Graphics::Direct2D::*;
    if std::env::var("OXI_S480_DISABLE").is_ok() { return None; }
    // S480 ships only the HEAVY ornamental "Stroked" art borders, where any
    // breaking of the thick solid bar clearly improves the match (gate-confirmed
    // +0.0044 on the tokumei dashDotStroked box family). Thin line-style borders
    // (dashed/dashSmallGap/dotted/dotDash) are left SOLID: Word renders their
    // small gaps nearly-solid, and a coarse width-scaled dash matches WORSE than
    // solid (459f05 −0.0078, 3a4f −0.0035). Those need a measured fine cadence
    // (follow-up); see S480 memo.
    match style? {
        "dashDotStroked" => Some(D2D1_DASH_STYLE_DASH_DOT),
        "dashDotDotStroked" => Some(D2D1_DASH_STYLE_DASH_DOT_DOT),
        _ => None,
    }
}

#[cfg(windows)]
unsafe fn render_line(
    rt: &windows::Win32::Graphics::Direct2D::ID2D1RenderTarget,
    x1_pt: f32, y1_pt: f32, x2_pt: f32, y2_pt: f32,
    color: Option<&str>,
    width_pt: f32,
    style: Option<&str>,
) -> windows::core::Result<()> {
    use windows::Win32::Graphics::Direct2D::*;
    use windows::Win32::Graphics::Direct2D::Common::*;
    let (r, g, b) = color.map(parse_hex_rgb).unwrap_or((0, 0, 0));
    let brush = rt.CreateSolidColorBrush(&rgb_to_d2d_color(r, g, b), None)?;
    // Build a dashed stroke style when the border declares a dash art style.
    let stroke_style: Option<ID2D1StrokeStyle> = match s480_dash_style(style) {
        Some(dash) => {
            let factory: ID2D1Factory = rt.GetFactory()?;
            let props = D2D1_STROKE_STYLE_PROPERTIES {
                startCap: D2D1_CAP_STYLE_FLAT,
                endCap: D2D1_CAP_STYLE_FLAT,
                dashCap: if dash == D2D1_DASH_STYLE_DOT { D2D1_CAP_STYLE_ROUND } else { D2D1_CAP_STYLE_FLAT },
                lineJoin: D2D1_LINE_JOIN_MITER,
                miterLimit: 10.0,
                dashStyle: dash,
                dashOffset: 0.0,
            };
            factory.CreateStrokeStyle(&props, None).ok()
        }
        None => None,
    };
    rt.DrawLine(
        D2D_POINT_2F { x: x1_pt * PT_TO_DIP, y: y1_pt * PT_TO_DIP },
        D2D_POINT_2F { x: x2_pt * PT_TO_DIP, y: y2_pt * PT_TO_DIP },
        &brush,
        width_pt * PT_TO_DIP,
        stroke_style.as_ref(),
    );
    Ok(())
}

#[cfg(windows)]
// S518 (2026-06-09): the DWrite ascent (GetLineMetrics[0].baseline) for a font/size,
// in points — the SAME value DrawTextLayout uses to place the baseline below the
// layout-rect top. Used to align a Symbol-font list marker (whose ascent is ~2pt
// smaller than the body at fs11) to its line's BODY baseline. Returns fs*0.8 on any
// DWrite error (harmless fallback; the caller only uses the DIFFERENCE of two ascents).
unsafe fn font_ascent_pt(
    dwrite_factory: &windows::Win32::Graphics::DirectWrite::IDWriteFactory,
    font_family: &str, font_size_pt: f32, bold: bool, italic: bool,
) -> f32 {
    use windows::core::*;
    use windows::Win32::Graphics::DirectWrite::*;
    let weight = if bold { DWRITE_FONT_WEIGHT_BOLD } else { DWRITE_FONT_WEIGHT_NORMAL };
    let style = if italic { DWRITE_FONT_STYLE_ITALIC } else { DWRITE_FONT_STYLE_NORMAL };
    let family_wide: Vec<u16> = font_family.encode_utf16().chain(std::iter::once(0)).collect();
    let locale_wide: Vec<u16> = "ja-jp".encode_utf16().chain(std::iter::once(0)).collect();
    let fmt = match dwrite_factory.CreateTextFormat(
        PCWSTR(family_wide.as_ptr()), None, weight, style,
        DWRITE_FONT_STRETCH_NORMAL, font_size_pt * PT_TO_DIP, PCWSTR(locale_wide.as_ptr())) {
        Ok(f) => f, Err(_) => return font_size_pt * 0.8,
    };
    let text_wide: Vec<u16> = "X".encode_utf16().collect();
    let layout = match dwrite_factory.CreateTextLayout(
        &text_wide, &fmt, 1000.0, font_size_pt * 2.0 * PT_TO_DIP) {
        Ok(l) => l, Err(_) => return font_size_pt * 0.8,
    };
    let mut lcount: u32 = 0;
    let _ = layout.GetLineMetrics(None, &mut lcount);
    if lcount > 0 {
        let mut lm = vec![DWRITE_LINE_METRICS::default(); lcount as usize];
        if layout.GetLineMetrics(Some(&mut lm), &mut lcount).is_ok() {
            return lm[0].baseline / PT_TO_DIP;
        }
    }
    font_size_pt * 0.8
}

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
    double_strikethrough: bool,
    double_underline: bool,
    highlight: Option<&str>,
    character_spacing_pt: f32,
    is_vertical: bool,
    text_scale_pct: f32,
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

    // S493f (env OXI_S493F_FAUXBOLD): faux-bold for fonts WITHOUT a real bold face.
    // Word (GDI/DWrite) synthesizes bold for MS Mincho/Gothic (+~40% ink, measured
    // bold/reg=1.40); DWrite's CreateTextFormat with BOLD weight silently falls back to
    // the REGULAR face for these (no simulation) → Oxi rendered bold CJK at regular
    // weight (bold/reg=1.00). Detect "no bold face" via the system collection; if so,
    // OVERPRINT the layout at +dx (GDI's faux-bold mechanism). Latin fonts with a real
    // bold face keep it (no overprint). OXI_FAUXBOLD_DX (pt) tunes the smear.
    // DWrite's system collection reports a nominal bold FACE for MS Gothic/Mincho
    // (GetFirstMatchingFont(BOLD) returns weight 700) yet DrawTextLayout renders it at
    // REGULAR weight (clean per-line repro: bold/reg ink = 1.000 vs Word's 1.396). So the
    // collection's bold-face report is NOT a reliable "real bold renders" signal. The
    // legacy MS bitmap-era families (ＭＳ 明朝/ゴシック/Ｐ明朝/Ｐゴシック/UI Gothic) lack a
    // real bold face and need faux-bold; Word (GDI) synthesizes it (+~40% ink). SCOPED to
    // the "ＭＳ"/"MS " families only — modern CJK fonts (游ゴシック/メイリオ = Yu/Meiryo)
    // DO have a real bold face DWrite renders, so faux-bolding them would DOUBLE-bold.
    let is_ms_legacy = font_family.starts_with("ＭＳ")
        || font_family.to_ascii_lowercase().starts_with("ms ");
    // Default ON (opt-out OXI_S493F_DISABLE): faux-bold MS-legacy CJK bold to match Word.
    let faux_bold: bool = bold && is_ms_legacy && std::env::var("OXI_S493F_DISABLE").is_err();
    // dx=0.4pt overprint => bold-line ink ratio 1.41 ≈ Word's measured 1.396 (calibrated
    // on the MS Gothic/Mincho repro: REG vs BOLD lines, Word bold/reg=1.396).
    let faux_dx: f32 = std::env::var("OXI_FAUXBOLD_DX").ok()
        .and_then(|v| v.parse().ok()).unwrap_or(0.4_f32) * PT_TO_DIP;

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
            // S627 (2026-06-19) FALSIFIED+REVERTED: a renderer-side discrete
            // error-diffused justify here was INERT (justified CJK lines are
            // ~1-char-per-fragment with character_spacing=0; the justify expansion
            // lives in the layout's frag_spacing_after, not here). The layout-side
            // version (mod.rs ~5716) was then tried and also regressed — Word's exact
            // justify bump PLACEMENT differs from error-diffusion. See mod.rs.
            let trailing = character_spacing_pt * PT_TO_DIP;
            let range = DWRITE_TEXT_RANGE { startPosition: 0, length: text_wide.len() as u32 };
            let _ = layout1.SetCharacterSpacing(0.0, trailing, 0.0, range);
        }
    }

    // S494: per-char EXACT positions for --dump-glyphs. HitTestTextPosition gives
    // each char's x (DIP, relative to the layout origin), which includes DWrite's
    // autoSpace + charGrid spacing — the gate-render positions. Horizontal text only.
    if !is_vertical {
        let dumping = GLYPH_DUMP.with(|g| g.borrow().is_some());
        if dumping {
            // S494 gate fix: the EXACT baseline = render origin (y_pt - 1.0) + the font's
            // DWrite ascent. GetLineMetrics[0].baseline is that ascent in DIP (what
            // DrawTextLayout uses to place the baseline below the layout-rect top). This
            // is the per-font ground truth (Calibri ~0.95em, Mincho ~0.86em), so the gate
            // no longer needs a fixed K to reconstruct it.
            let mut lcount: u32 = 0;
            let _ = layout.GetLineMetrics(None, &mut lcount);
            let asc_pt = if lcount > 0 {
                let mut lm = vec![DWRITE_LINE_METRICS::default(); lcount as usize];
                if layout.GetLineMetrics(Some(&mut lm), &mut lcount).is_ok() {
                    lm[0].baseline / PT_TO_DIP
                } else { font_size_pt * 0.8 }
            } else { font_size_pt * 0.8 };
            // Use the RENDER origin (y_pt - 1.0) — the SAME -1.0 glyph-top shift the actual
            // DrawTextLayout applies below — so the gate replicates where Oxi's product
            // ACTUALLY draws the baseline, then compares that to word_png. A/B on 4 docs
            // (0e7af/gen2_045 docGrid=none, b35/1ec1 table): including the -1.0 is decisively
            // better for the body docs (0e7af 0.784->0.905 = matches/exceeds the dwrite gate
            // = positions clean; gen2_045 0.931->0.950) and within ±0.02 (table vertical
            // noise) for the grid docs. Consistent with the `top` field (also y_pt - 1.0).
            let baseline_pt = (y_pt - 1.0) + asc_pt;
            let mut u16i: u32 = 0;
            let mut collected: Vec<(char, f32, f32, f32, f32, String)> = Vec::new();
            for ch in text.chars() {
                let mut px: f32 = 0.0;
                let mut py: f32 = 0.0;
                let mut hm = DWRITE_HIT_TEST_METRICS::default();
                let _ = layout.HitTestTextPosition(u16i, false, &mut px, &mut py, &mut hm);
                if !ch.is_whitespace() {
                    collected.push((
                        ch,
                        x_pt + px / PT_TO_DIP,  // absolute x in pt (HitTest = DWrite's
                                                // actual autoSpace/charGrid positions)
                        y_pt - 1.0,             // glyph top = the dwrite RENDER origin.y
                                                // (the -1.0 glyph-top alignment fix), so the
                                                // dump matches where dwrite ACTUALLY draws =
                                                // word_png. (page-1-only bench hid this; the
                                                // DENSE continuation pages need it — 0e7af
                                                // p2 was 2px low without it.)
                        baseline_pt,            // S494: exact glyph-origin baseline y (pt)
                        font_size_pt,
                        font_family.to_string(),
                    ));
                }
                u16i += ch.len_utf16() as u32;
            }
            GLYPH_DUMP.with(|g| {
                if let Some(v) = g.borrow_mut().as_mut() {
                    if let Some(last) = v.last_mut() {
                        last.2.extend(collected);
                    }
                }
            });
        }
    }

    // DWrite glyph-top alignment fix (2026-05-03): shift origin UP 1.0pt to
    // align rendered glyph top with Oxi's IR element.y. DirectWrite's
    // DrawTextLayout places baseline at origin.y + font_ascent, with leading
    // gap above caps; the empirical 1.0pt offset is the residual between
    // Oxi IR's intended glyph-top y and DirectWrite's actual glyph-top
    // placement at default origin. Full-corpus canary swept -0.25 .. -1.25pt
    // (5 values) and -1.0pt produced peak NET +2.2198 (110 wins / 41 regs)
    // on 177-doc p.1 baseline. See pipeline_data/dwrite_origin_shift_2026-05-03.md.
    // Session 489 (2026-06-02): Japanese vertical writing (tategaki) draws CJK
    // glyphs UPRIGHT, stacked top-to-bottom. S133 (below) rotated the WHOLE run
    // 90° CW, which rotates the GLYPHS sideways too — correct only for Latin,
    // WRONG for CJK. Word keeps kanji/kana upright (pixel-confirmed on
    // 7ead52 申請者/連絡担当窓口, 2ea81a 予納する理由, 459f05, ed025c — all
    // pure-CJK vertical cell labels). Draw each char upright at its stacked y;
    // per-char vertical advance = one em (font_size), matching
    // layout::vert_para_height (Σ n_chars·font_size). Whitespace advances but
    // is not drawn. Legacy rotation kept behind OXI_S489_ROTATE_LEGACY.
    if is_vertical && std::env::var("OXI_S489_ROTATE_LEGACY").is_err() {
        let adv_dip = font_size_pt * PT_TO_DIP;
        // Centre each full-width glyph (advance ≈ one em) horizontally within
        // the cell content column [x_pt, x_pt+w_pt] — Word centres tategaki
        // labels in their column. w_pt is the cell content width.
        let cx_dip = (x_pt + ((w_pt - font_size_pt) * 0.5).max(0.0)) * PT_TO_DIP;
        let mut yy = (y_pt - 1.0) * PT_TO_DIP;
        for ch in text.chars() {
            if !ch.is_whitespace() {
                let cw: Vec<u16> = ch.to_string().encode_utf16().collect();
                if let Ok(clayout) = dwrite_factory.CreateTextLayout(
                    &cw, &format, layout_w_dip, layout_h_dip,
                ) {
                    let corigin = D2D_POINT_2F { x: cx_dip, y: yy };
                    rt.DrawTextLayout(
                        corigin, &clayout, &brush, D2D1_DRAW_TEXT_OPTIONS_NONE);
                    if faux_bold {
                        let c2 = D2D_POINT_2F { x: cx_dip + faux_dx, y: yy };
                        rt.DrawTextLayout(c2, &clayout, &brush, D2D1_DRAW_TEXT_OPTIONS_NONE);
                    }
                }
            }
            yy += adv_dip;
        }
        return Ok(());
    }
    // Session 133 (2026-05-20): vertical writing rotation (LEGACY — see S489).
    // For tbRlV cells: apply a 90° CW rotation around the pivot point
    // (x_pt + font_size_pt, y_pt). Pre-rotation, text would extend
    // rightward from this pivot; post-rotation, it extends downward
    // (and the rotated glyphs occupy x ∈ [x_pt, x_pt + font_size_pt],
    // matching the cell's right-edge anchor convention from GDI S132).
    let prior_transform = if is_vertical {
        let mut saved = windows::Foundation::Numerics::Matrix3x2::default();
        rt.GetTransform(&mut saved);
        let pivot = D2D_POINT_2F {
            x: (x_pt + font_size_pt) * PT_TO_DIP,
            y: y_pt * PT_TO_DIP,
        };
        let mut rot = windows::Foundation::Numerics::Matrix3x2::default();
        windows::Win32::Graphics::Direct2D::D2D1MakeRotateMatrix(90.0, pivot, &mut rot);
        rt.SetTransform(&rot);
        Some(saved)
    } else {
        None
    };

    let origin = if is_vertical {
        // Draw at pivot point; rotation matrix maps it to itself.
        // Text extends rightward in local frame, which after rotation
        // becomes downward in screen frame.
        D2D_POINT_2F {
            x: (x_pt + font_size_pt) * PT_TO_DIP,
            y: (y_pt - 1.0) * PT_TO_DIP,
        }
    } else {
        // Original origin: DWrite glyph-top alignment fix (2026-05-03)
        // shifts y up by 1pt; see comment in pre-S133 code path.
        D2D_POINT_2F {
            x: x_pt * PT_TO_DIP,
            y: (y_pt - 1.0) * PT_TO_DIP,
        }
    };

    // S529 (coverage): w:w character-width scaling (text_scale %). The layout
    // already narrows the per-char ADVANCE to text_scale%, but the renderer was
    // drawing each glyph at FULL width within the narrowed slot → CJK glyphs
    // overlapped/crammed for w:w<100 (and were spread for >100). Word compresses
    // each GLYPH horizontally. Apply a horizontal scale = text_scale/100 around
    // the glyph's left edge (origin.x) so the glyph fills exactly its (already
    // scaled) advance. Skips vertical text. Affects 5 corpus docs using <w:w>.
    let ww_scale = text_scale_pct / 100.0;
    let ww_applied = !is_vertical && (ww_scale - 1.0).abs() > 0.005;
    let ww_saved = if ww_applied {
        let mut saved = windows::Foundation::Numerics::Matrix3x2::default();
        rt.GetTransform(&mut saved);
        let ax = origin.x; // left edge anchor (DIP)
        let m = windows::Foundation::Numerics::Matrix3x2 {
            M11: ww_scale, M12: 0.0, M21: 0.0, M22: 1.0,
            M31: ax * (1.0 - ww_scale), M32: 0.0,
        };
        rt.SetTransform(&(m * saved));
        Some(saved)
    } else { None };

    rt.DrawTextLayout(
        origin,
        &layout,
        &brush,
        D2D1_DRAW_TEXT_OPTIONS_NONE,
    );
    // S493f: faux-bold overprint (no real bold face) — second pass +faux_dx right.
    if faux_bold {
        let o2 = D2D_POINT_2F { x: origin.x + faux_dx, y: origin.y };
        rt.DrawTextLayout(o2, &layout, &brush, D2D1_DRAW_TEXT_OPTIONS_NONE);
    }
    if let Some(saved) = ww_saved {
        rt.SetTransform(&saved);
    }

    // Restore prior transform if we changed it for rotation.
    if let Some(saved) = prior_transform {
        rt.SetTransform(&saved);
    }

    // Strikethrough drawn manually as a horizontal line spanning the element
    // width — matches GDI's MoveToEx/LineTo path. DirectWrite's SetStrikethrough
    // skips whitespace, which breaks Oxi's per-char element model. (Single
    // underline intentionally not ported — both SetUnderline and DrawLine
    // variants regressed -0.18 net p.1 SSIM on 19 underline-containing baseline
    // docs in the 2026-05-02 canary; see
    // pipeline_data/dwrite_underline_investigation_2026-05-02.md. Double
    // underline (w:val="double") is scoped to one baseline doc — 1ec1091177b1
    // — and rendered explicitly below per fix-path B from the same investigation.)
    if strikethrough || double_strikethrough {
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
        let x_start = x_pt * PT_TO_DIP;
        let x_end = (x_pt + w_pt) * PT_TO_DIP;
        if double_strikethrough {
            // R66 v1: two parallel lines centered on the single-strike y.
            // Gap mirrors the double_underline heuristic below (font_size * 0.08).
            let gap_dip = (font_size_pt * PT_TO_DIP * 0.08).max(1.0);
            let y1 = y_dip - gap_dip * 0.5;
            let y2 = y_dip + gap_dip * 0.5;
            rt.DrawLine(
                D2D_POINT_2F { x: x_start, y: y1 },
                D2D_POINT_2F { x: x_end,   y: y1 },
                &brush, thickness, None,
            );
            rt.DrawLine(
                D2D_POINT_2F { x: x_start, y: y2 },
                D2D_POINT_2F { x: x_end,   y: y2 },
                &brush, thickness, None,
            );
        } else {
            rt.DrawLine(
                D2D_POINT_2F { x: x_start, y: y_dip },
                D2D_POINT_2F { x: x_end,   y: y_dip },
                &brush, thickness, None,
            );
        }
    }

    if double_underline {
        let mut count: u32 = 0;
        let _ = layout.GetLineMetrics(None, &mut count);
        let baseline_dip = if count > 0 {
            let mut metrics = vec![DWRITE_LINE_METRICS::default(); count as usize];
            if layout.GetLineMetrics(Some(&mut metrics), &mut count).is_ok() {
                metrics[0].baseline
            } else { font_size_pt * PT_TO_DIP * 0.8 }
        } else { font_size_pt * PT_TO_DIP * 0.8 };
        // S493g (2026-06-04): the two underlines were touching ("looks like one line").
        // Word's double underline (1ec1 20pt title, pixel-measured @150dpi): two ~2px lines,
        // 5px center-to-center with a clear ~3px gap. Old gap 0.08·fs (center-to-center) ≈ the
        // 0.06·fs thickness → edge-to-edge ~0 → merged. Match Word: gap 0.12·fs (center 5px),
        // thinner lines 0.05·fs (~2px). Opt-out OXI_S493G_DISABLE restores the old 0.08/0.06.
        let legacy_du = std::env::var("OXI_S493G_DISABLE").is_ok();
        let thickness = (font_size_pt * PT_TO_DIP * if legacy_du { 0.06 } else { 0.05 }).max(1.0);
        let offset_dip = (font_size_pt * PT_TO_DIP * 0.15).max(1.0);
        let gap_dip = (font_size_pt * PT_TO_DIP * if legacy_du { 0.08 } else { 0.12 }).max(1.0);
        let y1_dip = y_pt * PT_TO_DIP + baseline_dip + offset_dip;
        let y2_dip = y1_dip + gap_dip;
        let x_start = x_pt * PT_TO_DIP;
        let x_end = (x_pt + w_pt) * PT_TO_DIP;
        rt.DrawLine(
            D2D_POINT_2F { x: x_start, y: y1_dip },
            D2D_POINT_2F { x: x_end,   y: y1_dip },
            &brush, thickness, None,
        );
        rt.DrawLine(
            D2D_POINT_2F { x: x_start, y: y2_dip },
            D2D_POINT_2F { x: x_end,   y: y2_dip },
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

    // S623 (2026-06-19): downscale filter knob. Fant (default) is area-averaging =
    // SOFT (spreads glyph ink ~8% wider than Word's MuPDF analytic rasterization).
    // OXI_DOWNSCALE={fant|cubic|hqcubic|linear} to test a sharper downscale toward
    // MuPDF's crisper coverage. Gated on full-corpus ssim_ab.
    let interp = match std::env::var("OXI_DOWNSCALE").ok().as_deref() {
        Some("cubic") => WICBitmapInterpolationModeCubic,
        Some("hqcubic") => WICBitmapInterpolationModeHighQualityCubic,
        Some("linear") => WICBitmapInterpolationModeLinear,
        _ => WICBitmapInterpolationModeFant,
    };
    let source: IWICBitmapSource = if src_w != out_w || src_h != out_h {
        let scaler = wic_factory.CreateBitmapScaler()?;
        scaler.Initialize(
            src_bitmap,
            out_w,
            out_h,
            interp,
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

/// S494: per-char EXACT glyph positions (pt) from DirectWrite. Same schema as the
/// GDI --dump-glyphs so oxi_via_mupdf.py consumes either. dwrite positions include
/// DWrite's autoSpace + charGrid (the gate render), so faithful for tables too.
fn dump_glyphs_json(pages: &[(f32, f32, Vec<(char, f32, f32, f32, f32, String)>)], path: &str) {
    use std::fmt::Write;
    let mut out = String::from("{\n  \"pages\": [\n");
    for (pi, (pw, ph, glyphs)) in pages.iter().enumerate() {
        if pi > 0 { out.push_str(",\n"); }
        write!(&mut out, "    {{\"page\": {}, \"width\": {:.3}, \"height\": {:.3}, \"glyphs\": [\n",
               pi + 1, pw, ph).unwrap();
        let mut first = true;
        for (ch, x, top, baseline, fs, fam) in glyphs {
            let esc_ch = match *ch { '\\' => "\\\\".to_string(), '"' => "\\\"".to_string(), c => c.to_string() };
            let esc_fam = fam.replace('\\', "\\\\").replace('"', "\\\"");
            if !first { out.push_str(",\n"); }
            first = false;
            write!(&mut out,
                "      {{\"char\": \"{}\", \"x\": {:.3}, \"top\": {:.3}, \"baseline\": {:.3}, \"font_size\": {:.2}, \"font_family\": \"{}\"}}",
                esc_ch, x, top, baseline, fs, esc_fam).unwrap();
        }
        out.push_str("\n    ]}");
    }
    out.push_str("\n  ]\n}\n");
    std::fs::write(path, out).expect("Failed to write dump-glyphs JSON");
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
            // Session 73 Phase B: emit text_y_off for Y-convention refactor.
            // See memory/session71_y_convention_refactor_design.md.
            write!(&mut out,
                "      {{\"type\": \"{}\", \"x\": {:.3}, \"y\": {:.3}, \"w\": {:.3}, \"h\": {:.3}, \"text\": {}, \"font_size\": {:.2}, \"para_idx\": {}, \"run_idx\": {}, \"char_offset\": {}, \"text_y_off\": {:.3}}}",
                kind, el.x, el.y, el.width, el.height, text_json, font_size, pi_json, ri_json, co_json, el.text_y_off).unwrap();
        }
        out.push_str("\n    ]}");
    }
    out.push_str("\n  ]\n}\n");
    std::fs::write(path, out).expect("Failed to write dump-layout JSON");
}
