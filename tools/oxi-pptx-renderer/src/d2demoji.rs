/* This Source Code Form is subject to the terms of the Mozilla Public
 * License, v. 2.0. If a copy of the MPL was not distributed with this
 * file, You can obtain one at https://mozilla.org/MPL/2.0/. */

//! Colour emoji rasterised by DirectWrite, which is the engine PowerPoint
//! draws through.
//!
//! `emoji.rs` paints the COLR **v0** layer list -- a flat list of (glyph,
//! palette colour) -- because GDI has no colour-font support at all. Windows
//! 11's Segoe UI Emoji is a COLR **v1** font, and v1 is where its shading
//! lives: a census of its `BaseGlyphList` counts 76921 radial and 33647 linear
//! gradient paints. The v0 list is the fallback the same file keeps for old
//! rasterisers, so painting it gives correct SHAPES in flat colour -- which is
//! exactly what blind 09 slide 38 showed against PowerPoint's shaded ones, on
//! a slide that five decks of the corpus share.
//!
//! Direct2D on this machine does draw the v1 tree (`examples/d2d_emoji.rs`:
//! 3102 distinct opaque colours in one 204px U+1F44B against the ~3 a v0
//! render gives), so the glyph is rendered into a premultiplied BGRA bitmap
//! here and blended onto the GDI surface by the caller.

#![cfg(windows)]

use std::cell::RefCell;
use std::collections::HashMap;

use windows::core::PCWSTR;
use windows::Win32::Graphics::Direct2D::Common::*;
use windows::Win32::Graphics::Direct2D::*;
use windows::Win32::Graphics::DirectWrite::*;
use windows::Win32::Graphics::Imaging::*;
use windows::Win32::System::Com::*;

/// One rasterised glyph: premultiplied BGRA, plus where its pen origin sits.
pub struct Raster {
    pub w: i32,
    pub h: i32,
    /// Pixels from the bitmap's top edge down to the text baseline.
    pub baseline: i32,
    /// Pixels from the bitmap's left edge right to the pen position.
    pub pen_x: i32,
    pub bgra: Vec<u8>,
}

struct Factories {
    wic: IWICImagingFactory,
    d2d: ID2D1Factory,
    dw: IDWriteFactory,
}

thread_local! {
    static FACTORIES: RefCell<Option<Option<Factories>>> = const { RefCell::new(None) };
    /// (char, size in whole pixels) -> the raster, or None when it cannot be
    /// drawn. Rendering one glyph costs about a millisecond, and a slide can
    /// ask for the same emoji dozens of times.
    static CACHE: RefCell<HashMap<(char, u32), Option<std::rc::Rc<Raster>>>> =
        RefCell::new(HashMap::new());
}

/// The glyph is drawn this far inside the bitmap, so a negative side bearing
/// or an overhanging shadow is not clipped.
const PAD: f32 = 8.0;

fn factories<R>(f: impl FnOnce(&Factories) -> R) -> Option<R> {
    FACTORIES.with(|cell| {
        let mut slot = cell.borrow_mut();
        if slot.is_none() {
            *slot = Some(build_factories());
        }
        slot.as_ref().unwrap().as_ref().map(f)
    })
}

fn build_factories() -> Option<Factories> {
    unsafe {
        // The renderer never calls CoInitialize itself; a second call on the
        // same thread returns S_FALSE, which is not an error.
        let _ = CoInitializeEx(None, COINIT_APARTMENTTHREADED);
        let wic: IWICImagingFactory =
            CoCreateInstance(&CLSID_WICImagingFactory, None, CLSCTX_INPROC_SERVER).ok()?;
        let d2d: ID2D1Factory =
            D2D1CreateFactory(D2D1_FACTORY_TYPE_SINGLE_THREADED, None).ok()?;
        let dw: IDWriteFactory = DWriteCreateFactory(DWRITE_FACTORY_TYPE_SHARED).ok()?;
        Some(Factories { wic, d2d, dw })
    }
}

/// Rasterise `ch` at `size_px` (the em size in device pixels).
pub fn raster(ch: char, size_px: f32) -> Option<std::rc::Rc<Raster>> {
    if !(1.0..=4096.0).contains(&size_px) {
        return None;
    }
    let key = (ch, size_px.round() as u32);
    if let Some(hit) = CACHE.with(|c| c.borrow().get(&key).cloned()) {
        return hit;
    }
    let made = render(ch, key.1 as f32).map(std::rc::Rc::new);
    CACHE.with(|c| c.borrow_mut().insert(key, made.clone()));
    made
}

fn render(ch: char, size_px: f32) -> Option<Raster> {
    factories(|f| unsafe { render_with(f, ch, size_px) })?
}

unsafe fn render_with(f: &Factories, ch: char, size_px: f32) -> Option<Raster> {
    let mut text: Vec<u16> = ch.to_string().encode_utf16().collect();
    // U+FE0F asks for the colour presentation of a character that also has a
    // text one -- the same choice `run_plan` already made to reach this path.
    text.push(0xFE0F);
    let box_px = (size_px * 2.0 + PAD * 2.0).ceil() as u32;
    let bmp = f
        .wic
        .CreateBitmap(box_px, box_px, &GUID_WICPixelFormat32bppPBGRA, WICBitmapCacheOnLoad)
        .ok()?;
    let props = D2D1_RENDER_TARGET_PROPERTIES {
        r#type: D2D1_RENDER_TARGET_TYPE_DEFAULT,
        pixelFormat: D2D1_PIXEL_FORMAT {
            format: windows::Win32::Graphics::Dxgi::Common::DXGI_FORMAT_B8G8R8A8_UNORM,
            alphaMode: D2D1_ALPHA_MODE_PREMULTIPLIED,
        },
        // 96 dpi makes one DIP one pixel, so `size_px` is both the em size the
        // caller asked for and the size DirectWrite lays out at.
        dpiX: 96.0,
        dpiY: 96.0,
        usage: D2D1_RENDER_TARGET_USAGE_NONE,
        minLevel: D2D1_FEATURE_LEVEL_DEFAULT,
    };
    let rt = f.d2d.CreateWicBitmapRenderTarget(&bmp, &props).ok()?;
    let family: Vec<u16> = "Segoe UI Emoji\0".encode_utf16().collect();
    let locale: Vec<u16> = "en-us\0".encode_utf16().collect();
    let fmt = f
        .dw
        .CreateTextFormat(
            PCWSTR(family.as_ptr()),
            None,
            DWRITE_FONT_WEIGHT_NORMAL,
            DWRITE_FONT_STYLE_NORMAL,
            DWRITE_FONT_STRETCH_NORMAL,
            size_px,
            PCWSTR(locale.as_ptr()),
        )
        .ok()?;
    let layout = f
        .dw
        .CreateTextLayout(&text, &fmt, box_px as f32, box_px as f32)
        .ok()?;
    // The baseline is what ties the bitmap to the line: everything else on the
    // line is placed from it, so the emoji has to be blitted by it too.
    let mut lines = [DWRITE_LINE_METRICS::default(); 4];
    let mut n_lines = 0u32;
    layout
        .GetLineMetrics(Some(&mut lines), &mut n_lines)
        .ok()?;
    if n_lines == 0 {
        return None;
    }
    let baseline = lines[0].baseline;
    let brush = rt
        .CreateSolidColorBrush(&D2D1_COLOR_F { r: 0.0, g: 0.0, b: 0.0, a: 1.0 }, None)
        .ok()?;
    rt.BeginDraw();
    rt.Clear(Some(&D2D1_COLOR_F { r: 0.0, g: 0.0, b: 0.0, a: 0.0 }));
    rt.DrawTextLayout(
        D2D_POINT_2F { x: PAD, y: PAD },
        &layout,
        &brush,
        D2D1_DRAW_TEXT_OPTIONS_ENABLE_COLOR_FONT,
    );
    rt.EndDraw(None, None).ok()?;

    let lock = bmp
        .Lock(
            &WICRect { X: 0, Y: 0, Width: box_px as i32, Height: box_px as i32 },
            WICBitmapLockRead.0 as u32,
        )
        .ok()?;
    let stride = lock.GetStride().ok()? as usize;
    let mut len = 0u32;
    let mut data = std::ptr::null_mut();
    lock.GetDataPointer(&mut len, &mut data).ok()?;
    let raw = std::slice::from_raw_parts(data, len as usize);
    let w = box_px as usize;
    let mut bgra = vec![0u8; w * w * 4];
    let mut any = false;
    for y in 0..w {
        let src = &raw[y * stride..y * stride + w * 4];
        bgra[y * w * 4..(y + 1) * w * 4].copy_from_slice(src);
        if !any && src.chunks_exact(4).any(|p| p[3] != 0) {
            any = true;
        }
    }
    if !any {
        return None;
    }
    Some(Raster {
        w: box_px as i32,
        h: box_px as i32,
        baseline: (PAD + baseline).round() as i32,
        pen_x: PAD.round() as i32,
        bgra,
    })
}
