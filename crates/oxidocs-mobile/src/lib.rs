// SPDX-License-Identifier: MIT OR Apache-2.0

//! Oxidocs for native apps. A Swift or Kotlin viewer opens a document, asks
//! how many pages it has and how large they are, and draws the tiles of them
//! that are on screen, at the zoom it shows. Everything is laid out and drawn
//! in Rust (`oxidocs-render`), the same pages the browser build and the PDF
//! export draw, so a document looks the same on a tablet as in Word's tab.
//!
//! Generated bindings: `uniffi-bindgen generate --library <the built
//! liboxidocs_mobile> --language swift|kotlin`.

use std::sync::Arc;

use oxipdf_core::ir::PdfDocument;

uniffi::setup_scaffolding!();

#[derive(Debug, thiserror::Error, uniffi::Error)]
pub enum OxiError {
    /// The bytes are not a document Oxidocs can open, or opening it failed.
    #[error("cannot open: {reason}")]
    Open { reason: String },
    /// The file is a format the native viewer does not draw yet.
    #[error("not supported yet: {format}")]
    Unsupported { format: String },
}

/// A page's size in points (1/72 inch).
#[derive(Debug, Clone, Copy, PartialEq, uniffi::Record)]
pub struct PageSize {
    pub width: f64,
    pub height: f64,
}

/// `width` × `height` pixels of RGBA, rows top to bottom, not premultiplied.
#[derive(Debug, Clone, PartialEq, uniffi::Record)]
pub struct Tile {
    pub width: u32,
    pub height: u32,
    pub rgba: Vec<u8>,
}

/// An opened document: its pages, laid out once.
#[derive(uniffi::Object)]
pub struct Document {
    pages: PdfDocument,
}

#[uniffi::export]
impl Document {
    /// Open a document from its bytes. `name` is its file name; its
    /// extension decides the format when the bytes alone do not.
    #[uniffi::constructor]
    pub fn open(bytes: Vec<u8>, name: String) -> Result<Arc<Self>, OxiError> {
        match format_of(&bytes, &name) {
            Format::Word => {
                let pages = oxidocs_render::pdf::docx_pages(&bytes, None)
                    .map_err(|reason| OxiError::Open { reason })?;
                Ok(Arc::new(Self { pages }))
            }
            Format::Other(format) => Err(OxiError::Unsupported { format }),
        }
    }

    pub fn page_count(&self) -> u32 {
        self.pages.pages.len() as u32
    }

    /// The size of page `index` (0-based), or None past the last page.
    pub fn page_size(&self, index: u32) -> Option<PageSize> {
        oxidocs_render::page_size(&self.pages, index as usize).map(|(width, height)| PageSize { width, height })
    }

    /// Draw the part of page `index` that starts at the point (`left`, `top`)
    /// of the page, `width` × `height` pixels at `scale` pixels per point.
    /// A viewer asks for the tiles covering its window; zooming asks for new
    /// tiles of the same pages and lays nothing out again. None for a page
    /// that does not exist, a zero size or a scale that is not positive.
    pub fn render_tile(&self, index: u32, scale: f32, left: f32, top: f32, width: u32, height: u32) -> Option<Tile> {
        oxidocs_render::render_tile(&self.pages, index as usize, scale, left, top, width, height, stand_in())
            .map(|t| Tile { width: t.width, height: t.height, rgba: t.rgba })
    }
}

/// Register a font program the app found on the device or bundles, for
/// `family` in that style, before opening documents that use it. False when
/// no member of `bytes` names itself `family` in that style.
#[uniffi::export]
pub fn register_font(family: String, bold: bool, italic: bool, bytes: Vec<u8>) -> bool {
    oxidocs_render::fonts::try_register_family(&family, bold, italic, &bytes)
}

/// The face drawn for text whose font the pages name but do not embed.
fn stand_in() -> Option<&'static [u8]> {
    oxidocs_fonts::face("Carlito")
}

enum Format {
    Word,
    Other(String),
}

fn format_of(bytes: &[u8], name: &str) -> Format {
    let lower = name.to_lowercase();
    let ext = lower.rsplit_once('.').map(|(_, e)| e).unwrap_or("");
    match ext {
        "docx" | "docm" | "dotx" | "dotm" => Format::Word,
        "" => {
            let has = |part: &[u8]| bytes.windows(part.len()).any(|w| w == part);
            if has(b"word/document.xml") { Format::Word } else { Format::Other("unknown".into()) }
        }
        other => Format::Other(other.to_string()),
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    fn fixture(name: &str) -> Vec<u8> {
        std::fs::read(format!("{}/../../tests/fixtures/{name}", env!("CARGO_MANIFEST_DIR"))).unwrap()
    }

    #[test]
    fn a_word_document_opens_and_draws_tiles() {
        let doc = Document::open(fixture("basic_test.docx"), "basic_test.docx".into()).unwrap();
        assert!(doc.page_count() >= 1);
        let size = doc.page_size(0).unwrap();
        assert!(size.width > 500.0 && size.height > 700.0, "{size:?}");
        let tile = doc.render_tile(0, 2.0, 0.0, 0.0, 256, 256).unwrap();
        assert_eq!((tile.width, tile.height, tile.rgba.len()), (256, 256, 256 * 256 * 4));
        // The top-left corner of a page is margin: paper white.
        assert_eq!(&tile.rgba[..4], &[255, 255, 255, 255]);
        // Somewhere on the first page there is ink.
        let whole_w = (size.width * 1.0) as u32;
        let whole = doc.render_tile(0, 1.0, 0.0, 0.0, whole_w, (size.height) as u32).unwrap();
        assert!(whole.rgba.chunks_exact(4).any(|p| p[0] < 128), "no ink on page 1");
        assert!(doc.page_size(doc.page_count()).is_none());
        assert!(doc.render_tile(0, 0.0, 0.0, 0.0, 16, 16).is_none());
    }

    #[test]
    fn formats_not_drawn_yet_are_named_and_junk_is_refused() {
        match Document::open(vec![1, 2, 3], "deck.pptx".into()) {
            Err(OxiError::Unsupported { format }) => assert_eq!(format, "pptx"),
            other => panic!("{:?}", other.map(|_| ())),
        }
        assert!(matches!(Document::open(vec![1, 2, 3], "a.docx".into()), Err(OxiError::Open { .. })));
    }
}
