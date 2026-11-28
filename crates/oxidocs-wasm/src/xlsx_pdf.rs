// SPDX-License-Identifier: MIT OR Apache-2.0

//! Spreadsheet print pages as PDF in the browser. The layout is the command
//! line's (`oxicells_core::print`); this canvas measures and draws with the
//! faces the module carries: fonts the page registered, the bundled East
//! Asian stand-ins, and the bundled Latin stand-ins.

use oxipdf_core::ir::*;
use std::cell::RefCell;
use std::collections::{BTreeSet, HashMap};
use wasm_bindgen::prelude::*;

#[derive(Default)]
struct Canvas {
    pages: Vec<Page>,
    /// Faces loaded so far, by resource name. Measuring loads them too.
    fonts: RefCell<HashMap<String, EmbeddedFont>>,
}

impl Canvas {
    /// The resource that draws `ch` for this family and style: a registered
    /// program when it has the character, otherwise the bundled stand-in the
    /// .docx export would pick.
    fn face(&self, family: &str, bold: bool, italic: bool, ch: char) -> String {
        let mut fonts = self.fonts.borrow_mut();
        if crate::runtime_fonts::has_font_program(family, bold, italic) {
            let key = oxidocs_render::fonts::resource_name(family, bold, italic);
            if !fonts.contains_key(&key) {
                if let Some(font) = oxidocs_render::fonts::embedded_font(family, bold, italic) {
                    fonts.insert(key.clone(), font);
                }
            }
            if fonts.get(&key).is_some_and(|f| f.unicode_to_gid.contains_key(&(ch as u32))) {
                return key;
            }
        }
        let key = crate::resolve_wasm_font(family, bold, italic, &ch.to_string());
        if !fonts.contains_key(&key) {
            let font = match key.as_str() {
                "OxiCJK-Gothic" => Some(crate::get_cjk_fonts().0.clone()),
                "OxiCJK-Mincho" => Some(crate::get_cjk_fonts().1.clone()),
                other => crate::latin_face(other).cloned(),
            };
            if let Some(font) = font {
                fonts.insert(key.clone(), font);
            }
        }
        key
    }

    /// Advance of `ch` in em, from the face's own width table.
    fn advance(&self, key: &str, ch: char) -> f64 {
        let fallback = if (ch as u32) > 0x2E7F { 1.0 } else { 0.5 };
        self.fonts
            .borrow()
            .get(key)
            .and_then(|font| {
                let gid = font.unicode_to_gid.get(&(ch as u32))?;
                font.cid_widths.get(gid).map(|w| f64::from(*w) / 1000.0)
            })
            .unwrap_or(fallback)
    }
}

impl oxicells_core::print::PrintCanvas for Canvas {
    fn page(&mut self, width: f64, height: f64) {
        self.pages.push(Page {
            width,
            height,
            media_box: Rectangle { llx: 0.0, lly: 0.0, urx: width, ury: height },
            crop_box: None,
            contents: Vec::new(),
            rotation: 0,
        });
    }

    fn element(&mut self, element: ContentElement) {
        if let Some(page) = self.pages.last_mut() {
            page.contents.push(element);
        }
    }

    fn width(&self, text: &str, family: &str, size: f64, bold: bool, italic: bool) -> f64 {
        text.chars()
            .map(|ch| self.advance(&self.face(family, bold, italic, ch), ch))
            .sum::<f64>()
            * size
    }

    fn text(&mut self, text: &str, family: &str, size: f64, bold: bool, italic: bool, x: f64, y: f64, color: Color) {
        let mut pen = x;
        let mut open: Option<TextSpan> = None;
        let mut spans = Vec::new();
        for ch in text.chars() {
            let key = self.face(family, bold, italic, ch);
            let advance = self.advance(&key, ch) * size;
            match open.as_mut().filter(|span| span.font_name == key) {
                Some(span) => span.text.push(ch),
                None => {
                    spans.extend(open.take());
                    // The bundled East Asian faces have no bold member: bold is
                    // drawn filled and stroked, as in the .docx export.
                    let fake_bold = bold && key.starts_with("OxiCJK-");
                    open = Some(TextSpan {
                        x: pen,
                        y,
                        text: ch.to_string(),
                        font_name: key,
                        font_size: size,
                        fill_color: color,
                        character_spacing: 0.0,
                        stroke_width: if fake_bold { size / 35.0 } else { 0.0 },
                        skew: 0.0,
                        horizontal_scale: 100.0,
                        stroke_color: None,
                    });
                }
            }
            pen += advance;
        }
        spans.extend(open);
        for span in spans {
            self.element(ContentElement::Text(span));
        }
    }
}

/// Convert a workbook (.xlsx, .xlsm, .xltx, .xltm) to PDF bytes: the print
/// areas of each visible sheet, or its used range, laid out with the sheet's
/// page setup from the values saved in the file. Formulas are not
/// recalculated. `title` becomes the PDF's title when it is not empty.
#[wasm_bindgen]
pub fn xlsx_to_pdf(data: &[u8], title: &str) -> Result<Vec<u8>, JsError> {
    let doc = workbook_pdf(data, title).map_err(|e| JsError::new(&e))?;
    oxipdf_core::write_pdf_checked(&doc).map_err(|e| JsError::new(&format!("PDF glyph encoding: {e:?}")))
}

fn workbook_pdf(data: &[u8], title: &str) -> Result<PdfDocument, String> {
    let mut canvas = Canvas::default();
    oxicells_core::print::pages(data, &mut canvas)?;
    let Canvas { pages, fonts } = canvas;
    let mut fonts = fonts.into_inner();

    // Embed only the faces drawn, each cut down to the characters drawn.
    let mut drawn: HashMap<&str, BTreeSet<char>> = HashMap::new();
    for element in pages.iter().flat_map(|p| &p.contents) {
        if let ContentElement::Text(span) = element {
            drawn.entry(span.font_name.as_str()).or_default().extend(span.text.chars());
        }
    }
    fonts.retain(|key, _| drawn.contains_key(key.as_str()));
    for (key, font) in fonts.iter_mut() {
        if let Some(subset) = drawn.get(key.as_str()).and_then(|chars| crate::subset_embedded_font(font, chars)) {
            *font = subset;
        }
    }

    let title = title.trim();
    Ok(PdfDocument {
        version: PdfVersion::new(1, 7),
        info: DocumentInfo {
            title: (!title.is_empty()).then(|| title.to_string()),
            producer: Some("Oxi (xlsx→PDF)".to_string()),
            ..Default::default()
        },
        pages,
        outline: Vec::new(),
        embedded_fonts: fonts,
    })
}

#[cfg(test)]
mod tests {
    use super::*;

    fn fixture(name: &str) -> Vec<u8> {
        std::fs::read(format!("{}/../../tests/fixtures/{name}", env!("CARGO_MANIFEST_DIR"))).unwrap()
    }

    #[test]
    fn workbook_pages_embed_only_the_faces_drawn() {
        let doc = workbook_pdf(&fixture("basic_test.xlsx"), "請求書").unwrap();
        assert!(!doc.pages.is_empty());
        let drawn: BTreeSet<&str> = doc.pages.iter().flat_map(|p| &p.contents)
            .filter_map(|e| match e { ContentElement::Text(s) => Some(s.font_name.as_str()), _ => None })
            .collect();
        assert!(!drawn.is_empty());
        let embedded: BTreeSet<&str> = doc.embedded_fonts.keys().map(String::as_str).collect();
        assert_eq!(drawn, embedded);
        assert_eq!(doc.info.title.as_deref(), Some("請求書"));
        assert!(oxipdf_core::write_pdf_checked(&doc).unwrap().starts_with(b"%PDF-"));
    }

    #[test]
    fn east_asian_text_is_measured_with_the_bundled_face() {
        use oxicells_core::print::PrintCanvas;
        let canvas = Canvas::default();
        let w = canvas.width("請求書", "ＭＳ Ｐゴシック", 10.0, false, false);
        assert!((w - 30.0).abs() < 3.0, "{w}");
        let latin = canvas.width("abc", "Calibri", 10.0, false, false);
        assert!(latin > 10.0 && latin < 20.0, "{latin}");
    }
}
