// This Source Code Form is subject to the terms of the Mozilla Public
// License, v. 2.0. If a copy of the MPL was not distributed with this
// file, You can obtain one at https://mozilla.org/MPL/2.0/.

//! Integration tests: parse `mc:AlternateContent` and `w:object` (OLE)
//! end-to-end and verify the resulting Block::Image / Paragraph.shapes
//! routing after `parse_docx`.
//!
//! Parser code paths tested:
//!   - [parser/ooxml.rs:4017](crates/oxidocs-core/src/parser/ooxml.rs#L4017)
//!     `parse_alternate_content`: prefers <mc:Choice> (DrawingML)
//!     over <mc:Fallback> (VML legacy).
//!   - [parser/ooxml.rs:3931](crates/oxidocs-core/src/parser/ooxml.rs#L3931)
//!     `parse_ole_object`: <w:object> with embedded
//!     <v:shape><v:imagedata r:id="..."/></v:shape> → Image with
//!     alt_text="OLE Object" HARDCODED.
//!
//! Non-obvious behaviors pinned:
//!   - mc:Choice WINS over mc:Fallback when both are present
//!     (parser/ooxml.rs:4036-4042 + line 4043 `result.is_none()`
//!     gate prevents Fallback from overwriting Choice). Modern docs
//!     ship both for forward compat — the parser must follow the
//!     OOXML rule of preferring Choice's richer DrawingML over
//!     Fallback's VML.
//!   - When ONLY Fallback is present, parser falls back to pict
//!     path (line 4043). The result is a VML shape, routed to
//!     Paragraph.shapes (per S317 routing discovery).
//!   - parse_ole_object HARDCODES alt_text="OLE Object" (line 4002).
//!     A regression that left alt_text=None or stored the rel_id
//!     would silently affect accessibility tools that read
//!     alt_text for object-embedded content (charts, equations,
//!     spreadsheet snippets).
//!   - parse_ole_object returns DrawingResult.image=None when no
//!     <v:imagedata> is present (line 4009-4011). The paragraph
//!     produces NO inline Block::Image — a regression that emitted
//!     an empty Image would surface as a phantom 0×0 image at
//!     the paragraph position.
//!
//! Fixtures live in `tools/fixtures/alternate_content_ole_samples/`
//! and are authored by
//! `tools/metrics/build_alternate_content_ole_repro_fixtures.py`.

use std::fs;

use oxidocs_core::ir::{Block, Document};
use oxidocs_core::parse_docx;

fn fixture_path(name: &str) -> std::path::PathBuf {
    let crate_root = std::env::current_dir().unwrap();
    let workspace_root = crate_root.parent().unwrap().parent().unwrap();
    workspace_root
        .join("tools")
        .join("fixtures")
        .join("alternate_content_ole_samples")
        .join(name)
}

fn load(name: &str) -> Option<Document> {
    let path = fixture_path(name);
    if !path.exists() {
        eprintln!("skipping: {} not found", path.display());
        return None;
    }
    let data = fs::read(&path).expect("read fixture");
    Some(parse_docx(&data).expect("parse fixture"))
}

/// The OLE preview image, wherever the parser routes it.
///
/// S974 (2026-07-21) moved a BODY `<w:object>` preview off the sibling
/// `Block::Image` path and onto its host run (`inline_object_image`), because
/// Word flows it inline in the host line — measured on
/// administrative__00018048, whose two 228.5x48.2 banners both sit on their
/// host line. The preview and its metadata must survive either way, so this
/// helper looks in both places; the tests below pin alt_text / dimensions /
/// position, which are the contract, not the storage slot.
fn first_inline_image(doc: &Document) -> Option<&oxidocs_core::ir::Image> {
    for page in &doc.pages {
        for block in &page.blocks {
            match block {
                Block::Image(img) => return Some(img),
                Block::Paragraph(p) => {
                    for r in &p.runs {
                        if let Some(img) = r.style.inline_object_image.as_deref() {
                            return Some(img);
                        }
                    }
                }
                _ => {}
            }
        }
    }
    None
}

fn first_paragraph_shapes(doc: &Document) -> Vec<&oxidocs_core::ir::Shape> {
    for page in &doc.pages {
        for block in &page.blocks {
            if let Block::Paragraph(p) = block {
                if !p.shapes.is_empty() {
                    return p.shapes.iter().collect();
                }
            }
        }
    }
    Vec::new()
}

#[test]
fn v1_ac_choice_only_drawing_routes_to_block_image() {
    let Some(doc) = load("v1_ac_choice_only_drawing.docx") else { return };

    // mc:Choice contains a DrawingML inline image. Parser at line
    // 4036-4042 routes parse_drawing output. The image's alt_text
    // ("from-choice") confirms Choice was the source.
    let img = first_inline_image(&doc).expect(
        "mc:Choice-only with <w:drawing> must produce Block::Image",
    );
    assert_eq!(
        img.alt_text.as_deref(),
        Some("from-choice"),
        "Choice's drawing was the source"
    );
}

#[test]
fn v1_ac_choice_wins_over_fallback() {
    let Some(doc) = load("v1_ac_choice_wins_over_fallback.docx") else { return };

    // BOTH Choice (DrawingML) and Fallback (VML pict) are present.
    // Parser MUST prefer Choice (line 4036 fires first; line 4043's
    // `result.is_none()` gate prevents Fallback from overwriting).
    let img = first_inline_image(&doc).expect(
        "Choice's drawing must produce Block::Image (NOT VML shape)",
    );
    assert_eq!(
        img.alt_text.as_deref(),
        Some("from-choice"),
        "Choice wins — alt_text confirms (Fallback's red rect would NOT have this)"
    );

    // Equally important: Fallback's VML shape must NOT also surface
    // in Paragraph.shapes. The parser is single-result: once Choice
    // populated result, the pict branch's `if dr.has_content()`
    // check still runs but the gate `result.is_none()` is now false.
    let shapes = first_paragraph_shapes(&doc);
    assert!(
        shapes.is_empty(),
        "Fallback's VML shape must NOT also surface — Choice is single-source-of-truth"
    );
}

#[test]
fn v1_ac_fallback_only_pict_routes_to_paragraph_shapes() {
    let Some(doc) = load("v1_ac_fallback_only_pict.docx") else { return };

    // No Choice — parser falls back to pict (line 4043). VML shape
    // goes to Paragraph.shapes per S317 routing.
    let shapes = first_paragraph_shapes(&doc);
    assert_eq!(
        shapes.len(),
        1,
        "Fallback-only with <w:pict> → 1 shape in Paragraph.shapes"
    );
    assert_eq!(
        shapes[0].shape_type, "rect",
        "<v:rect> from Fallback → \"rect\""
    );
    assert_eq!(
        shapes[0].fill.as_deref(),
        Some("00FF00"),
        "Fallback's VML fill survives end-to-end"
    );

    // No Block::Image (Choice was absent).
    assert!(
        first_inline_image(&doc).is_none(),
        "Fallback-only → no Block::Image (DrawingML was absent)"
    );
}

#[test]
fn v1_ole_with_imagedata_pins_hardcoded_alt_text_and_dimensions() {
    let Some(doc) = load("v1_ole_with_imagedata_preview.docx") else { return };

    let img = first_inline_image(&doc).expect(
        "<w:object> with <v:imagedata> must produce Block::Image preview",
    );

    // alt_text is HARDCODED to "OLE Object" (parser/ooxml.rs:4002).
    // A regression that left alt_text=None or stored the rel_id
    // would silently affect accessibility readers.
    assert_eq!(
        img.alt_text.as_deref(),
        Some("OLE Object"),
        "parse_ole_object HARDCODES alt_text=\"OLE Object\""
    );

    // Width/height from <v:shape style="..."> via parse_css_length.
    assert!(
        (img.width - 120.0).abs() < 0.001,
        "v:shape style=\"width:120pt\" → 120.0pt, got {}",
        img.width
    );
    assert!(
        (img.height - 60.0).abs() < 0.001,
        "v:shape style=\"height:60pt\" → 60.0pt, got {}",
        img.height
    );

    // OLE preview is INLINE (position=None) per parse_ole_object
    // line 4004. NOT a floating image.
    assert!(
        img.position.is_none(),
        "OLE preview is inline (parser hardcodes position=None)"
    );
}

#[test]
fn v1_ole_no_imagedata_produces_no_image() {
    let Some(doc) = load("v1_ole_no_imagedata.docx") else { return };

    // No <v:imagedata> inside <w:object> → rel_id stays None →
    // DrawingResult.image = None (parser/ooxml.rs:4009-4011). The
    // paragraph produces NO Block::Image. A regression that emitted
    // an empty/zero-sized Image would surface as a phantom 0×0 image.
    assert!(
        first_inline_image(&doc).is_none(),
        "<w:object> without <v:imagedata> → DrawingResult.image=None → \
         NO Block::Image (NOT a phantom empty image)"
    );
}

#[test]
fn all_five_fixtures_produce_either_image_or_shape_or_nothing() {
    let cases: &[(&str, bool, bool)] = &[
        // (filename, expects_image, expects_shape)
        ("v1_ac_choice_only_drawing.docx", true, false),
        ("v1_ac_choice_wins_over_fallback.docx", true, false),
        ("v1_ac_fallback_only_pict.docx", false, true),
        ("v1_ole_with_imagedata_preview.docx", true, false),
        ("v1_ole_no_imagedata.docx", false, false),
    ];
    for (name, exp_img, exp_shape) in cases {
        let path = fixture_path(name);
        if !path.exists() {
            eprintln!("skipping {}", name);
            continue;
        }
        let data = fs::read(&path).unwrap();
        let doc = parse_docx(&data)
            .unwrap_or_else(|e| panic!("failed to parse {}: {:?}", name, e));
        let has_img = first_inline_image(&doc).is_some();
        let has_shape = !first_paragraph_shapes(&doc).is_empty();
        assert_eq!(has_img, *exp_img, "{} expects image={}", name, exp_img);
        assert_eq!(has_shape, *exp_shape, "{} expects shape={}", name, exp_shape);
    }
}
