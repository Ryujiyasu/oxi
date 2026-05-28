//! Public-API integration tests for oxipdf-core.
//!
//! Pins the externally-visible behavior of the public PDF pipeline:
//!   write_pdf → parse_pdf → extract_text
//!
//! Validates document version, metadata, multi-page structure, and the
//! error path on malformed input. These are not unit tests of internal
//! helpers — they verify the supported API surface a consumer would call.
//!
//! Companion to the in-src `extract::tests` (which already covers one
//! round-trip case); this file expands coverage to the broader public
//! pipeline and codifies the boundary so refactors notice when the
//! shape of the public IR changes.

use oxipdf_core::ir::*;
use std::collections::HashMap;

fn simple_doc(text: &str) -> PdfDocument {
    PdfDocument {
        version: PdfVersion::new(1, 7),
        info: DocumentInfo::default(),
        pages: vec![Page {
            width: 612.0,
            height: 792.0,
            media_box: Rectangle { llx: 0.0, lly: 0.0, urx: 612.0, ury: 792.0 },
            crop_box: None,
            contents: vec![ContentElement::Text(TextSpan {
                x: 72.0,
                y: 72.0,
                text: text.to_string(),
                font_name: "F1".into(),
                font_size: 12.0,
                fill_color: Color::Gray(0.0),
                character_spacing: 0.0,
            })],
            rotation: 0,
        }],
        outline: Vec::new(),
        embedded_fonts: HashMap::new(),
    }
}

#[test]
fn write_pdf_produces_valid_pdf_header() {
    let doc = simple_doc("hello");
    let bytes = oxipdf_core::write_pdf(&doc);
    assert!(bytes.starts_with(b"%PDF-1.7\n"), "expected PDF-1.7 header");
    // Trailer ends with %%EOF (optionally followed by whitespace).
    let tail = std::str::from_utf8(&bytes[bytes.len().saturating_sub(64)..]).unwrap_or("");
    assert!(tail.contains("%%EOF"), "expected %%EOF in trailer, got tail: {tail:?}");
}

#[test]
fn write_then_parse_preserves_page_count() {
    let mut doc = simple_doc("page 1");
    doc.pages.push(Page {
        width: 612.0,
        height: 792.0,
        media_box: Rectangle { llx: 0.0, lly: 0.0, urx: 612.0, ury: 792.0 },
        crop_box: None,
        contents: vec![ContentElement::Text(TextSpan {
            x: 72.0, y: 100.0, text: "page 2".into(),
            font_name: "F1".into(), font_size: 12.0,
            fill_color: Color::Gray(0.0), character_spacing: 0.0,
        })],
        rotation: 0,
    });
    let bytes = oxipdf_core::write_pdf(&doc);
    let parsed = oxipdf_core::parse_pdf(&bytes).expect("parse round-trip");
    assert_eq!(parsed.pages.len(), 2);
}

#[test]
fn write_then_parse_preserves_text() {
    let doc = simple_doc("Roundtrip ASCII text");
    let bytes = oxipdf_core::write_pdf(&doc);
    let parsed = oxipdf_core::parse_pdf(&bytes).expect("parse round-trip");
    let extracted = oxipdf_core::extract_text_string(&parsed);
    assert!(
        extracted.contains("Roundtrip ASCII text"),
        "expected text in extraction, got {extracted:?}"
    );
}

#[test]
fn write_then_parse_preserves_pdf_version() {
    let doc = simple_doc("v");
    let bytes = oxipdf_core::write_pdf(&doc);
    let parsed = oxipdf_core::parse_pdf(&bytes).expect("parse round-trip");
    assert_eq!(parsed.version.major, 1);
    assert_eq!(parsed.version.minor, 7);
}

#[test]
fn write_then_parse_preserves_page_dimensions() {
    let doc = simple_doc("dim");
    let bytes = oxipdf_core::write_pdf(&doc);
    let parsed = oxipdf_core::parse_pdf(&bytes).expect("parse round-trip");
    assert!((parsed.pages[0].width - 612.0).abs() < 0.5);
    assert!((parsed.pages[0].height - 792.0).abs() < 0.5);
}

#[test]
fn extract_text_returns_one_entry_per_page() {
    let mut doc = simple_doc("p1");
    doc.pages.push(Page {
        width: 612.0, height: 792.0,
        media_box: Rectangle { llx: 0.0, lly: 0.0, urx: 612.0, ury: 792.0 },
        crop_box: None,
        contents: vec![ContentElement::Text(TextSpan {
            x: 72.0, y: 72.0, text: "p2".into(),
            font_name: "F1".into(), font_size: 12.0,
            fill_color: Color::Gray(0.0), character_spacing: 0.0,
        })],
        rotation: 0,
    });
    let pages = oxipdf_core::extract_text(&doc);
    assert_eq!(pages.len(), 2);
    assert_eq!(pages[0].page_number, 1);
    assert_eq!(pages[1].page_number, 2);
}

#[test]
fn extract_text_skips_non_text_elements() {
    // Path + Text on same page → extract_text returns only the text.
    let doc = PdfDocument {
        version: PdfVersion::new(1, 7),
        info: DocumentInfo::default(),
        pages: vec![Page {
            width: 612.0, height: 792.0,
            media_box: Rectangle { llx: 0.0, lly: 0.0, urx: 612.0, ury: 792.0 },
            crop_box: None,
            contents: vec![
                ContentElement::Path(PathData {
                    operations: vec![PathOp::MoveTo(0.0, 0.0), PathOp::LineTo(100.0, 100.0)],
                    stroke: Some(StrokeStyle {
                        color: Color::Gray(0.0), width: 1.0,
                        line_cap: LineCap::Butt, line_join: LineJoin::Miter,
                    }),
                    fill: None,
                }),
                ContentElement::Text(TextSpan {
                    x: 72.0, y: 72.0, text: "TEXT".into(),
                    font_name: "F1".into(), font_size: 12.0,
                    fill_color: Color::Gray(0.0), character_spacing: 0.0,
                }),
            ],
            rotation: 0,
        }],
        outline: Vec::new(),
        embedded_fonts: HashMap::new(),
    };
    let pages = oxipdf_core::extract_text(&doc);
    assert_eq!(pages[0].spans.len(), 1, "only the Text element should be extracted");
    assert_eq!(pages[0].spans[0].text, "TEXT");
}

#[test]
fn parse_rejects_garbage_input() {
    // Random bytes are not a valid PDF; parser must return Err, not panic.
    let result = oxipdf_core::parse_pdf(b"this is not a pdf file at all");
    assert!(result.is_err(), "expected parse error on garbage input");
}

#[test]
fn parse_rejects_empty_input() {
    let result = oxipdf_core::parse_pdf(b"");
    assert!(result.is_err(), "expected parse error on empty input");
}

#[test]
fn write_pdf_is_deterministic_for_identical_input() {
    // Two write_pdf calls on the same IR must produce byte-identical output
    // (no clock-based / RNG-based fields). Required for content-addressed
    // caches and reproducible signing workflows.
    let doc = simple_doc("determinism");
    let a = oxipdf_core::write_pdf(&doc);
    let b = oxipdf_core::write_pdf(&doc);
    assert_eq!(a, b, "write_pdf must be deterministic");
}

#[test]
fn rectangle_width_height_compute_from_corners() {
    let r = Rectangle { llx: 10.0, lly: 20.0, urx: 110.0, ury: 220.0 };
    assert_eq!(r.width(), 100.0);
    assert_eq!(r.height(), 200.0);
}

#[test]
fn pdf_version_display_formats_as_dotted() {
    let v = PdfVersion::new(1, 7);
    assert_eq!(format!("{v}"), "1.7");
    let v2 = PdfVersion::new(2, 0);
    assert_eq!(format!("{v2}"), "2.0");
}

#[test]
fn empty_page_round_trip_produces_no_text() {
    let doc = PdfDocument {
        version: PdfVersion::new(1, 7),
        info: DocumentInfo::default(),
        pages: vec![Page {
            width: 612.0, height: 792.0,
            media_box: Rectangle { llx: 0.0, lly: 0.0, urx: 612.0, ury: 792.0 },
            crop_box: None,
            contents: vec![],
            rotation: 0,
        }],
        outline: Vec::new(),
        embedded_fonts: HashMap::new(),
    };
    let bytes = oxipdf_core::write_pdf(&doc);
    let parsed = oxipdf_core::parse_pdf(&bytes).expect("empty page round-trip");
    assert_eq!(parsed.pages.len(), 1);
    let extracted = oxipdf_core::extract_text_string(&parsed);
    assert!(extracted.is_empty(), "empty page must extract to empty string, got {extracted:?}");
}
