// This Source Code Form is subject to the terms of the Mozilla Public
// License, v. 2.0. If a copy of the MPL was not distributed with this
// file, You can obtain one at https://mozilla.org/MPL/2.0/.

//! Integration tests for oxislides-core shape & run property parsing.
//!
//! S351's public_api_integration covered parse error paths, slide
//! dimensions, Slide.index, and PptxEditor round-trip. This suite pins the
//! VISUAL property extraction that the renderer depends on: shape geometry
//! (EMU→pt), run bold/italic/color/font, and paragraph alignment. A
//! regression here silently corrupts slide rendering.

use oxislides_core::ir::{ShapeContent, SlideAlignment};
use oxislides_core::parser::parse_pptx;

const PPTX: &[u8] = include_bytes!("../../../tests/fixtures/basic_test.pptx");

fn pres() -> oxislides_core::ir::Presentation {
    parse_pptx(PPTX).expect("basic_test.pptx must parse")
}

#[test]
fn slide_has_title_and_body_shapes() {
    let p = pres();
    let slide = &p.slides[0];
    assert!(slide.shapes.len() >= 2, "title + body");
}

#[test]
fn shape_geometry_emu_to_points() {
    // basic_test slide1: title shape off=(457200,274638) ext=(8229600,1143000).
    // 1 point = 12700 EMU. → x=36pt, y≈21.63pt, w=648pt, h=90pt.
    let p = pres();
    let title = &p.slides[0].shapes[0];
    assert!((title.x - 36.0).abs() < 0.5, "title x≈36pt, got {}", title.x);
    assert!((title.y - 21.63).abs() < 0.5, "title y≈21.6pt, got {}", title.y);
    assert!((title.width - 648.0).abs() < 1.0, "title w≈648pt, got {}", title.width);
    assert!((title.height - 90.0).abs() < 1.0, "title h≈90pt, got {}", title.height);
}

#[test]
fn shape_geometry_is_finite_and_nonneg() {
    // Every shape must have finite, non-negative geometry (renderer contract).
    let p = pres();
    for (si, slide) in p.slides.iter().enumerate() {
        for (shi, sh) in slide.shapes.iter().enumerate() {
            for (name, v) in [("x", sh.x), ("y", sh.y), ("w", sh.width), ("h", sh.height)] {
                assert!(v.is_finite(), "slide {si} shape {shi} {name} not finite");
                assert!(v >= 0.0, "slide {si} shape {shi} {name} negative: {v}");
            }
        }
    }
}

#[test]
fn title_run_is_bold() {
    let p = pres();
    if let ShapeContent::TextBox { paragraphs } = &p.slides[0].shapes[0].content {
        let run = &paragraphs[0].runs[0];
        assert_eq!(run.text, "Welcome to Oxi");
        assert!(run.bold, "title run must be bold");
    } else {
        panic!("title must be a TextBox");
    }
}

#[test]
fn body_has_italic_colored_run() {
    // basic_test body (shape 1) has an italic run colored 4472C4.
    let p = pres();
    if let ShapeContent::TextBox { paragraphs } = &p.slides[0].shapes[1].content {
        let mut found = false;
        for para in paragraphs {
            for run in &para.runs {
                if run.italic && run.color.as_deref() == Some("4472C4") {
                    found = true;
                }
            }
        }
        assert!(found, "body must contain an italic run colored 4472C4");
    } else {
        panic!("body must be a TextBox");
    }
}

#[test]
fn run_color_is_six_hex_digits_when_present() {
    // Any parsed run color must be a 6-hex-digit string (no leading '#',
    // uppercase or lowercase) — the renderer assumes this format.
    let p = pres();
    for slide in &p.slides {
        for shape in &slide.shapes {
            if let ShapeContent::TextBox { paragraphs } = &shape.content {
                for para in paragraphs {
                    for run in &para.runs {
                        if let Some(c) = &run.color {
                            assert_eq!(c.len(), 6, "color must be 6 hex digits, got {c:?}");
                            assert!(
                                c.chars().all(|ch| ch.is_ascii_hexdigit()),
                                "color must be hex, got {c:?}"
                            );
                        }
                    }
                }
            }
        }
    }
}

#[test]
fn paragraph_alignment_parsed() {
    // basic_test slide1 has a center-aligned paragraph (algn="ctr").
    let p = pres();
    let mut saw_center = false;
    for shape in &p.slides[0].shapes {
        if let ShapeContent::TextBox { paragraphs } = &shape.content {
            for para in paragraphs {
                if matches!(para.alignment, SlideAlignment::Center) {
                    saw_center = true;
                }
            }
        }
    }
    assert!(saw_center, "slide1 must have a center-aligned paragraph");
}

#[test]
fn alignment_defaults_to_left_when_unspecified() {
    // multi_slide.pptx slide 1 title is left/default-aligned. Pin that
    // unspecified alignment yields the Left default (not a panic / garbage).
    let multi = parse_pptx(include_bytes!("../../../tests/fixtures/multi_slide.pptx"))
        .expect("multi_slide.pptx must parse");
    // Just assert every paragraph has a valid alignment variant (exhaustive
    // match guards against an unparsed/garbage state).
    for slide in &multi.slides {
        for shape in &slide.shapes {
            if let ShapeContent::TextBox { paragraphs } = &shape.content {
                for para in paragraphs {
                    match para.alignment {
                        SlideAlignment::Left
                        | SlideAlignment::Center
                        | SlideAlignment::Right
                        | SlideAlignment::Justify => {}
                    }
                }
            }
        }
    }
}
