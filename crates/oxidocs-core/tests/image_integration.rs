// This Source Code Form is subject to the terms of the Mozilla Public
// License, v. 2.0. If a copy of the MPL was not distributed with this
// file, You can obtain one at https://mozilla.org/MPL/2.0/.

//! Integration tests: parse inline `<w:drawing>` images end-to-end and
//! verify `Block::Image` populates with alt_text, dimensions, data blob,
//! and content_type.
//!
//! Parser code paths tested:
//! - [parser/ooxml.rs:359](crates/oxidocs-core/src/parser/ooxml.rs#L359):
//!   loads `word/media/imageN.{png,jpg,...}` binary via the image rels.
//! - [parser/ooxml.rs:2993](crates/oxidocs-core/src/parser/ooxml.rs#L2993):
//!   extracts `wp:docPr/@descr` as `Image.alt_text`.
//! - [parser/ooxml.rs:3407](crates/oxidocs-core/src/parser/ooxml.rs#L3407):
//!   reads `wp:extent/@cx,@cy` (EMU; divides by 12700 for pt) as
//!   width/height.
//!
//! Fixtures live in `tools/fixtures/image_samples/` and are authored by
//! `tools/metrics/build_image_repro_fixtures.py` (S292). Each fixture
//! embeds a minimal 1×1 RGBA PNG (~68 bytes) at `word/media/image1.png`.

use std::fs;

use oxidocs_core::ir::{Block, Document, Image};
use oxidocs_core::parser::parse_docx;

fn fixture_path(name: &str) -> std::path::PathBuf {
    let crate_root = std::env::current_dir().unwrap();
    let workspace_root = crate_root.parent().unwrap().parent().unwrap();
    workspace_root.join("tools").join("fixtures").join("image_samples").join(name)
}

fn collect_images(doc: &Document) -> Vec<&Image> {
    let mut out = Vec::new();
    for page in &doc.pages {
        walk_blocks(&page.blocks, &mut out);
    }
    out
}

fn walk_blocks<'a>(blocks: &'a [Block], out: &mut Vec<&'a Image>) {
    for b in blocks {
        match b {
            Block::Image(img) => out.push(img),
            Block::Table(t) => {
                for row in &t.rows {
                    for cell in &row.cells {
                        walk_blocks(&cell.blocks, out);
                    }
                }
            }
            _ => {}
        }
    }
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

#[test]
fn v1_simple_image_has_alt_text_and_one_inch_dimensions() {
    let Some(doc) = load("v1_simple.docx") else { return };
    let images = collect_images(&doc);
    assert_eq!(images.len(), 1, "exactly 1 image expected");
    let img = images[0];
    assert_eq!(img.alt_text.as_deref(), Some("Test alt text"));
    // 914400 EMU = 1 inch = 72pt
    assert!((img.width - 72.0).abs() < 0.1, "width {}", img.width);
    assert!((img.height - 72.0).abs() < 0.1, "height {}", img.height);
    assert_eq!(img.content_type.as_deref(), Some("image/png"));
    // Minimal 1×1 RGBA PNG is ~68 bytes
    assert!(img.data.len() > 0, "image data must be non-empty");
    assert!(img.data.starts_with(b"\x89PNG"),
        "data must start with PNG signature");
    // Inline (not floating)
    assert!(img.position.is_none(), "inline image has no FloatingPosition");
}

#[test]
fn v1_no_alt_image_has_none_alt_text() {
    // Image without `descr` attribute on wp:docPr → alt_text None.
    let Some(doc) = load("v1_no_alt.docx") else { return };
    let images = collect_images(&doc);
    assert_eq!(images.len(), 1);
    assert!(images[0].alt_text.is_none(),
        "image without descr should have alt_text = None");
    // Other fields still populated
    assert_eq!(images[0].content_type.as_deref(), Some("image/png"));
    assert!(images[0].data.len() > 0);
}

#[test]
fn v1_custom_size_carries_2in_by_1in_dimensions() {
    // cx=1828800 cy=914400 EMU = 144×72pt = 2in × 1in
    let Some(doc) = load("v1_custom_size.docx") else { return };
    let images = collect_images(&doc);
    assert_eq!(images.len(), 1);
    let img = images[0];
    assert!((img.width - 144.0).abs() < 0.1, "width {}", img.width);
    assert!((img.height - 72.0).abs() < 0.1, "height {}", img.height);
    assert_eq!(img.alt_text.as_deref(), Some("Custom-sized image"));
}

#[test]
fn all_three_fixtures_parse_with_valid_png_data() {
    // Smoke test: each fixture loads, image data starts with PNG signature.
    for name in ["v1_simple.docx", "v1_no_alt.docx", "v1_custom_size.docx"] {
        let path = fixture_path(name);
        if !path.exists() {
            eprintln!("skipping {}", name);
            continue;
        }
        let data = fs::read(&path).unwrap();
        let doc = parse_docx(&data)
            .unwrap_or_else(|e| panic!("failed to parse {}: {:?}", name, e));
        let images = collect_images(&doc);
        assert_eq!(images.len(), 1, "{} should have exactly 1 image", name);
        let img = images[0];
        assert!(img.data.starts_with(b"\x89PNG"),
            "{} image data must be a PNG", name);
        assert_eq!(img.content_type.as_deref(), Some("image/png"),
            "{} content_type", name);
    }
}
