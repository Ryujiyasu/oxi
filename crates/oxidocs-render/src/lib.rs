// This Source Code Form is subject to the terms of the Mozilla Public
// License, v. 2.0. If a copy of the MPL was not distributed with this
// file, You can obtain one at https://mozilla.org/MPL/2.0/.

//! Oxidocs pages drawn to pixels.
//!
//! The pages are the `PdfDocument` the PDF writer writes, so what is drawn
//! here is what the PDF shows. A whole page ([`raster::render_page`]) serves
//! thumbnails; a [`Tile`] of a page at any zoom serves a viewer that scrolls
//! and pinches, drawing only the part of the page on screen, so the same
//! pages can be shown natively on a tablet or a phone.

pub mod font_program;
pub mod fonts;
pub mod pdf;
#[cfg(feature = "raster")]
pub mod raster;

use oxipdf_core::ir::PdfDocument;

#[cfg(feature = "raster")]
/// A rectangle of a page drawn at some zoom: `width` × `height` pixels of
/// RGBA, rows top to bottom, not premultiplied (what Android's `Bitmap` and
/// iOS's `CGImage` take as `ARGB_8888` / `RGBA8888` after a channel swap, or
/// directly as non-premultiplied RGBA).
#[derive(Clone, Debug, PartialEq, Eq)]
pub struct Tile {
    pub width: u32,
    pub height: u32,
    pub rgba: Vec<u8>,
}

#[cfg(feature = "raster")]
/// Draw the part of page `index` (0-based) of `doc` that starts at the point
/// (`left`, `top`) of the page, `width` × `height` pixels, at `scale` pixels
/// per point (96 dpi is 4/3). Parts of the tile beyond the page are paper
/// white. `stand_in` is a TrueType face for text in fonts the document did
/// not embed. None when the page does not exist, the size is zero or the
/// scale is not a positive number.
///
/// A viewer asks for the tiles that cover its window, at the zoom it shows,
/// and keeps the ones it already has as it scrolls; changing the zoom asks
/// for new tiles of the same pages, with no layout work.
pub fn render_tile(
    doc: &PdfDocument,
    index: usize,
    scale: f32,
    left: f32,
    top: f32,
    width: u32,
    height: u32,
    stand_in: Option<&[u8]>,
) -> Option<Tile> {
    if !(scale.is_finite() && scale > 0.0) || width == 0 || height == 0 {
        return None;
    }
    let page = doc.pages.get(index)?;
    let pixmap = raster::render_region(doc, page, scale, (left, top), width, height, stand_in)?;
    Some(Tile { width, height, rgba: unpremultiplied(&pixmap) })
}

/// The size of page `index` in points, (width, height).
pub fn page_size(doc: &PdfDocument, index: usize) -> Option<(f64, f64)> {
    doc.pages.get(index).map(|p| (p.width, p.height))
}

#[cfg(feature = "raster")]
fn unpremultiplied(pixmap: &tiny_skia::Pixmap) -> Vec<u8> {
    let mut out = Vec::with_capacity(pixmap.data().len());
    for px in pixmap.pixels() {
        let c = px.demultiply();
        out.extend_from_slice(&[c.red(), c.green(), c.blue(), c.alpha()]);
    }
    out
}

#[cfg(feature = "raster")]
#[cfg(test)]
mod tests {
    use super::*;
    use oxipdf_core::ir::{Color, ContentElement, DocumentInfo, Page, PathData, PathOp, PdfVersion, Rectangle};
    use std::collections::HashMap;

    /// A US-letter page with a filled rectangle at (100, 100)–(300, 200) pt.
    fn page_with_box() -> PdfDocument {
        PdfDocument {
            version: PdfVersion::new(1, 7),
            info: DocumentInfo::default(),
            pages: vec![Page {
                width: 612.0,
                height: 792.0,
                media_box: Rectangle { llx: 0.0, lly: 0.0, urx: 612.0, ury: 792.0 },
                crop_box: None,
                contents: vec![ContentElement::Path(PathData {
                    operations: vec![
                        PathOp::MoveTo(100.0, 100.0),
                        PathOp::LineTo(300.0, 100.0),
                        PathOp::LineTo(300.0, 200.0),
                        PathOp::LineTo(100.0, 200.0),
                        PathOp::ClosePath,
                    ],
                    stroke: None,
                    fill: Some(Color::Rgb(1.0, 0.0, 0.0)),
                })],
                rotation: 0,
            }],
            outline: vec![],
            embedded_fonts: HashMap::new(),
        }
    }

    fn pixel(tile: &Tile, x: u32, y: u32) -> [u8; 4] {
        let i = ((y * tile.width + x) * 4) as usize;
        tile.rgba[i..i + 4].try_into().unwrap()
    }

    #[test]
    fn a_tile_is_the_same_pixels_as_that_part_of_the_whole_page() {
        let doc = page_with_box();
        let scale = 2.0;
        let whole = render_tile(&doc, 0, scale, 0.0, 0.0, 1224, 1584, None).unwrap();
        // The tile whose top-left is the point (150, 120): pixel (300, 240) of the page.
        let tile = render_tile(&doc, 0, scale, 150.0, 120.0, 256, 256, None).unwrap();
        for (x, y) in [(0, 0), (10, 10), (200, 150), (255, 255)] {
            assert_eq!(pixel(&tile, x, y), pixel(&whole, 300 + x, 240 + y), "tile pixel ({x}, {y})");
        }
        // Inside the box: red. Below it (y = 120 + 200/2 = 220 pt > 200 pt): white.
        assert_eq!(pixel(&tile, 10, 10), [255, 0, 0, 255]);
        assert_eq!(pixel(&tile, 10, 200), [255, 255, 255, 255]);
    }

    #[test]
    fn zoom_scales_the_page_without_layout() {
        let doc = page_with_box();
        // At 4 px/pt the box's left edge (100 pt) is at pixel 400.
        let tile = render_tile(&doc, 0, 4.0, 0.0, 100.0, 512, 8, None).unwrap();
        assert_eq!(pixel(&tile, 398, 4), [255, 255, 255, 255]);
        assert_eq!(pixel(&tile, 402, 4), [255, 0, 0, 255]);
    }

    #[test]
    fn beyond_the_page_is_white_and_bad_requests_are_none() {
        let doc = page_with_box();
        let tile = render_tile(&doc, 0, 1.0, 600.0, 780.0, 64, 64, None).unwrap();
        assert_eq!(pixel(&tile, 63, 63), [255, 255, 255, 255]);
        assert!(render_tile(&doc, 1, 1.0, 0.0, 0.0, 64, 64, None).is_none());
        assert!(render_tile(&doc, 0, 0.0, 0.0, 0.0, 64, 64, None).is_none());
        assert!(render_tile(&doc, 0, 1.0, 0.0, 0.0, 0, 64, None).is_none());
        assert_eq!(page_size(&doc, 0), Some((612.0, 792.0)));
    }
}
