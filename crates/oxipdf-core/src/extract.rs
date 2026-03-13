//! Text extraction from PDF documents.

use crate::ir::*;

/// Extract all text from a PDF document, page by page.
pub fn extract_text(doc: &PdfDocument) -> Vec<PageText> {
    doc.pages
        .iter()
        .enumerate()
        .map(|(i, page)| {
            let spans = extract_page_text(page);
            PageText {
                page_number: i + 1,
                spans,
            }
        })
        .collect()
}

/// Extract all text from a PDF as a single string.
pub fn extract_text_string(doc: &PdfDocument) -> String {
    let pages = extract_text(doc);
    pages
        .iter()
        .map(|p| p.to_string())
        .collect::<Vec<_>>()
        .join("\n\n")
}

/// Text content of a single page.
#[derive(Debug, Clone)]
pub struct PageText {
    pub page_number: usize,
    pub spans: Vec<TextPosition>,
}

/// A positioned text fragment.
#[derive(Debug, Clone)]
pub struct TextPosition {
    pub x: f64,
    pub y: f64,
    pub text: String,
    pub font_size: f64,
}

impl PageText {
    /// Get the text content as a single string, with spans sorted by position.
    pub fn to_string(&self) -> String {
        let mut sorted = self.spans.clone();
        // Sort by y (top to bottom), then x (left to right).
        sorted.sort_by(|a, b| {
            a.y.partial_cmp(&b.y)
                .unwrap_or(std::cmp::Ordering::Equal)
                .then(a.x.partial_cmp(&b.x).unwrap_or(std::cmp::Ordering::Equal))
        });

        let mut result = String::new();
        let mut last_y: Option<f64> = None;

        for span in &sorted {
            if let Some(ly) = last_y {
                // If y changes significantly, add a newline.
                if (span.y - ly).abs() > span.font_size * 0.5 {
                    result.push('\n');
                } else {
                    result.push(' ');
                }
            }
            result.push_str(&span.text);
            last_y = Some(span.y);
        }

        result
    }
}

fn extract_page_text(page: &Page) -> Vec<TextPosition> {
    page.contents
        .iter()
        .filter_map(|el| match el {
            ContentElement::Text(span) => Some(TextPosition {
                x: span.x,
                y: span.y,
                text: span.text.clone(),
                font_size: span.font_size,
            }),
            _ => None,
        })
        .collect()
}

#[cfg(test)]
mod tests {
    use super::*;

    fn make_doc_with_text(texts: Vec<(&str, f64, f64)>) -> PdfDocument {
        let contents: Vec<ContentElement> = texts
            .into_iter()
            .map(|(text, x, y)| {
                ContentElement::Text(TextSpan {
                    x,
                    y,
                    text: text.to_string(),
                    font_name: "F1".into(),
                    font_size: 12.0,
                    fill_color: Color::Gray(0.0),
                })
            })
            .collect();

        PdfDocument {
            version: PdfVersion::new(1, 7),
            info: DocumentInfo::default(),
            pages: vec![Page {
                width: 612.0,
                height: 792.0,
                media_box: Rectangle { llx: 0.0, lly: 0.0, urx: 612.0, ury: 792.0 },
                crop_box: None,
                contents,
                rotation: 0,
            }],
            outline: Vec::new(),
        }
    }

    #[test]
    fn test_extract_single_line() {
        let doc = make_doc_with_text(vec![("Hello World", 72.0, 72.0)]);
        let text = extract_text_string(&doc);
        assert_eq!(text, "Hello World");
    }

    #[test]
    fn test_extract_multi_line() {
        let doc = make_doc_with_text(vec![
            ("Line 1", 72.0, 72.0),
            ("Line 2", 72.0, 90.0),
        ]);
        let text = extract_text_string(&doc);
        assert!(text.contains("Line 1"));
        assert!(text.contains("Line 2"));
        assert!(text.contains('\n'));
    }

    #[test]
    fn test_extract_same_line() {
        let doc = make_doc_with_text(vec![
            ("Hello", 72.0, 72.0),
            ("World", 120.0, 72.0),
        ]);
        let text = extract_text_string(&doc);
        assert_eq!(text, "Hello World");
    }

    #[test]
    fn test_extract_empty() {
        let doc = PdfDocument {
            version: PdfVersion::new(1, 7),
            info: DocumentInfo::default(),
            pages: vec![Page {
                width: 612.0,
                height: 792.0,
                media_box: Rectangle { llx: 0.0, lly: 0.0, urx: 612.0, ury: 792.0 },
                crop_box: None,
                contents: vec![],
                rotation: 0,
            }],
            outline: Vec::new(),
        };
        let text = extract_text_string(&doc);
        assert!(text.is_empty());
    }

    #[test]
    fn test_extract_multi_page() {
        let doc = PdfDocument {
            version: PdfVersion::new(1, 7),
            info: DocumentInfo::default(),
            pages: vec![
                Page {
                    width: 612.0,
                    height: 792.0,
                    media_box: Rectangle { llx: 0.0, lly: 0.0, urx: 612.0, ury: 792.0 },
                    crop_box: None,
                    contents: vec![ContentElement::Text(TextSpan {
                        x: 72.0, y: 72.0,
                        text: "Page 1".into(),
                        font_name: "F1".into(),
                        font_size: 12.0,
                        fill_color: Color::Gray(0.0),
                    })],
                    rotation: 0,
                },
                Page {
                    width: 612.0,
                    height: 792.0,
                    media_box: Rectangle { llx: 0.0, lly: 0.0, urx: 612.0, ury: 792.0 },
                    crop_box: None,
                    contents: vec![ContentElement::Text(TextSpan {
                        x: 72.0, y: 72.0,
                        text: "Page 2".into(),
                        font_name: "F1".into(),
                        font_size: 12.0,
                        fill_color: Color::Gray(0.0),
                    })],
                    rotation: 0,
                },
            ],
            outline: Vec::new(),
        };
        let pages = extract_text(&doc);
        assert_eq!(pages.len(), 2);
        assert_eq!(pages[0].page_number, 1);
        assert_eq!(pages[1].page_number, 2);
        assert!(pages[0].to_string().contains("Page 1"));
        assert!(pages[1].to_string().contains("Page 2"));
    }

    #[test]
    fn test_roundtrip_text_extraction() {
        // Write a PDF, parse it, extract text.
        let doc = make_doc_with_text(vec![("Roundtrip Test", 72.0, 72.0)]);
        let pdf_bytes = crate::write_pdf(&doc);
        let parsed = crate::parse_pdf(&pdf_bytes).unwrap();
        let text = extract_text_string(&parsed);
        assert!(text.contains("Roundtrip Test"));
    }
}
