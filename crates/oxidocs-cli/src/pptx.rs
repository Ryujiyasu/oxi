// This Source Code Form is subject to the terms of the Mozilla Public
// License, v. 2.0. If a copy of the MPL was not distributed with this
// file, You can obtain one at https://mozilla.org/MPL/2.0/.

//! `.pptx` to PDF pages, drawn by `oxidocs_render::pptx` with the faces this
//! machine has: the font index, the metric-compatible substitutes, the East
//! Asian and symbol fallbacks `docx-to-pdf` uses, embedded from their files.

use std::collections::BTreeSet;

use oxidocs_core::ir::UnsupportedContent;
pub use oxidocs_render::pptx::TextPages;
use oxidocs_render::pptx::{FaceId, SlideFonts};
use oxipdf_core::ir::{EmbeddedFont, PdfDocument};
use oxislides_core::ir::Presentation;

/// The pages of a presentation, one per slide, and what on them is not drawn
/// as PowerPoint draws it.
pub fn pages(pres: &Presentation) -> (PdfDocument, Vec<UnsupportedContent>) {
    oxidocs_render::pptx::pages(pres, &MachineFonts)
}

/// The canvas the spreadsheet PDF path draws on, with this machine's faces.
pub fn text_pages() -> TextPages {
    TextPages::new(&MachineFonts)
}

struct MachineFonts;

fn lookup(name: &str, bold: bool, italic: bool) -> Option<FaceId> {
    let index = super::font_index();
    index
        .find(name, bold, italic)
        .or_else(|| super::substitute_for(name).and_then(|s| index.find(s, bold, italic)))
        .map(|face| FaceId { name: face.path.to_string_lossy().to_string(), index: face.index })
}

impl SlideFonts for MachineFonts {
    fn resolve(&self, family: &str, bold: bool, italic: bool) -> Option<FaceId> {
        lookup(family, bold, italic).or_else(|| {
            super::split_style_suffix(family).and_then(|(base, b, i)| lookup(&base, bold || b, italic || i))
        })
    }

    fn stand_in(&self, bold: bool, italic: bool) -> Option<FaceId> {
        ["Liberation Sans", "Carlito", "Arial", "DejaVu Sans"]
            .iter()
            .find_map(|stand_in| lookup(stand_in, bold, italic))
    }

    fn fallbacks(&self) -> Vec<FaceId> {
        let east_asian = super::CJK_FONT_PATHS
            .iter()
            .map(|p| p.to_string())
            .find(|p| std::path::Path::new(p).exists())
            .or_else(|| super::bundled_face_file("OxiGothic.ttf").map(|p| p.to_string_lossy().to_string()));
        east_asian
            .into_iter()
            .chain(super::FALLBACK_FACES.iter().map(|p| p.to_string()))
            .filter(|p| std::path::Path::new(p).exists())
            .map(|name| FaceId { name, index: 0 })
            .collect()
    }

    fn has_face(&self, family: &str, bold: bool, italic: bool) -> bool {
        super::font_index().find(family, bold, italic).is_some()
    }

    fn load(&self, id: &FaceId) -> Option<Vec<u8>> {
        std::fs::read(&id.name).ok()
    }

    fn embed(&self, id: &FaceId, _data: &[u8], chars: &BTreeSet<char>, key: &str) -> Option<EmbeddedFont> {
        if id.index == 0 {
            super::subset_font_file(&id.name, chars, key)
        } else {
            super::embed_collection_member(&id.name, id.index, key)
        }
    }
}
