/* This Source Code Form is subject to the terms of the Mozilla Public
 * License, v. 2.0. If a copy of the MPL was not distributed with this
 * file, You can obtain one at https://mozilla.org/MPL/2.0/. */

//! Does Direct2D on THIS machine draw Segoe UI Emoji's COLR **v1** paint tree?
//!
//! Oxi paints the v0 layer list itself, which is flat: blind 09 slide 38's
//! emoji come out as solid silhouettes against PowerPoint's shaded ones, and
//! the same slide is in five decks. The v1 graph is what carries the shading --
//! a census of the font's BaseGlyphList counts 76921 radial and 33647 linear
//! gradients -- and DirectWrite is the engine PowerPoint renders through, so
//! the question is only whether this API path exposes it.
//!
//!     cargo run --release --example d2d_emoji -- <out.png> [emoji] [size]
//!
//! Writes the glyph on a transparent ground. Flat fills mean v0; a gradient
//! means v1 and the integration is worth building.

#[cfg(windows)]
fn main() -> windows::core::Result<()> {
    use windows::core::{Interface, PCWSTR};
    use windows::Win32::Graphics::Direct2D::Common::*;
    use windows::Win32::Graphics::Direct2D::*;
    use windows::Win32::Graphics::DirectWrite::*;
    use windows::Win32::Graphics::Imaging::*;
    use windows::Win32::System::Com::*;

    let args: Vec<String> = std::env::args().collect();
    let out = args.get(1).cloned().unwrap_or_else(|| "d2d_emoji.png".into());
    let text: Vec<u16> = args
        .get(2)
        .map(|s| s.as_str())
        .unwrap_or("\u{1F44B}")
        .encode_utf16()
        .collect();
    let size: f32 = args.get(3).and_then(|s| s.parse().ok()).unwrap_or(128.0);
    let px = (size * 1.6) as u32;

    unsafe {
        CoInitializeEx(None, COINIT_APARTMENTTHREADED).ok()?;
        let wic: IWICImagingFactory =
            CoCreateInstance(&CLSID_WICImagingFactory, None, CLSCTX_INPROC_SERVER)?;
        let bmp = wic.CreateBitmap(px, px, &GUID_WICPixelFormat32bppPBGRA, WICBitmapCacheOnLoad)?;
        let d2d: ID2D1Factory = D2D1CreateFactory(D2D1_FACTORY_TYPE_SINGLE_THREADED, None)?;
        let props = D2D1_RENDER_TARGET_PROPERTIES {
            r#type: D2D1_RENDER_TARGET_TYPE_DEFAULT,
            pixelFormat: D2D1_PIXEL_FORMAT {
                format: windows::Win32::Graphics::Dxgi::Common::DXGI_FORMAT_B8G8R8A8_UNORM,
                alphaMode: D2D1_ALPHA_MODE_PREMULTIPLIED,
            },
            dpiX: 96.0,
            dpiY: 96.0,
            usage: D2D1_RENDER_TARGET_USAGE_NONE,
            minLevel: D2D1_FEATURE_LEVEL_DEFAULT,
        };
        let rt = d2d.CreateWicBitmapRenderTarget(&bmp, &props)?;
        let dw: IDWriteFactory = DWriteCreateFactory(DWRITE_FACTORY_TYPE_SHARED)?;
        let family: Vec<u16> = "Segoe UI Emoji\0".encode_utf16().collect();
        let locale: Vec<u16> = "en-us\0".encode_utf16().collect();
        let fmt = dw.CreateTextFormat(
            PCWSTR(family.as_ptr()),
            None,
            DWRITE_FONT_WEIGHT_NORMAL,
            DWRITE_FONT_STYLE_NORMAL,
            DWRITE_FONT_STRETCH_NORMAL,
            size,
            PCWSTR(locale.as_ptr()),
        )?;
        let brush = rt.CreateSolidColorBrush(
            &D2D1_COLOR_F { r: 0.0, g: 0.0, b: 0.0, a: 1.0 },
            None,
        )?;
        rt.BeginDraw();
        rt.Clear(Some(&D2D1_COLOR_F { r: 0.0, g: 0.0, b: 0.0, a: 0.0 }));
        let rect = D2D_RECT_F { left: 0.0, top: 0.0, right: px as f32, bottom: px as f32 };
        rt.DrawText(
            &text,
            &fmt,
            &rect,
            &brush,
            D2D1_DRAW_TEXT_OPTIONS_ENABLE_COLOR_FONT,
            DWRITE_MEASURING_MODE_NATURAL,
        );
        rt.EndDraw(None, None)?;

        let lock = bmp.Lock(
            &WICRect { X: 0, Y: 0, Width: px as i32, Height: px as i32 },
            WICBitmapLockRead.0 as u32,
        )?;
        let stride = lock.GetStride()?;
        let mut size_out = 0u32;
        let mut data = std::ptr::null_mut();
        lock.GetDataPointer(&mut size_out, &mut data)?;
        let raw = std::slice::from_raw_parts(data, size_out as usize);
        let mut img = image::RgbaImage::new(px, px);
        let mut colours = std::collections::HashSet::new();
        for y in 0..px {
            for x in 0..px {
                let o = (y * stride + x * 4) as usize;
                let (b, g, r, a) = (raw[o], raw[o + 1], raw[o + 2], raw[o + 3]);
                img.put_pixel(x, y, image::Rgba([r, g, b, a]));
                if a > 200 {
                    colours.insert((r, g, b));
                }
            }
        }
        img.save(&out).unwrap();
        println!("wrote {out} ({px}x{px}), {} distinct opaque colours", colours.len());
        println!("  (a flat v0 render of U+1F44B has ~3; a v1 gradient has hundreds)");
    }
    Ok(())
}

#[cfg(not(windows))]
fn main() {}
