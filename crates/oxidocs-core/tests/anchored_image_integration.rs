//! Integration tests: parse `<w:drawing>` with `<wp:anchor>` (floating)
//! end-to-end and verify the FloatingPosition / WrapType / ImageCrop
//! routing after `parse_docx`.
//!
//! `image_integration.rs` (S292) already covers INLINE images
//! (`<wp:inline>`): extent EMU→pt, alt_text, embedded PNG blob,
//! `Image.position.is_none()` confirms inline routing. This file
//! exercises the OTHER half of `parse_drawing` (parser/ooxml.rs:2958):
//! anchored images that route to `Page.floating_images` (NOT
//! `Block::Image` inside paragraph blocks).
//!
//! Non-obvious behaviors pinned:
//!   - Routing: image with position=Some → Page.floating_images
//!     (parser/ooxml.rs:832-836 dispatches on position.is_some()).
//!     `Page.blocks` for these fixtures has NO Block::Image entry.
//!     The paragraph that contained the `<w:drawing>` still exists
//!     but is otherwise empty.
//!   - posOffset Text content → EMU/12700 → pt (NOT a divisor of
//!     914400, which is the inch divisor). 914400 EMU = 72pt.
//!     A regression that used /914400 (inch) instead of /12700
//!     (point) would silently produce widths/positions 12.7× too
//!     small.
//!   - positionH/V `relativeFrom` attribute captured as
//!     h_relative/v_relative (string, NOT enum at the IR level —
//!     downstream layout dispatches on the string).
//!   - `<wp:align>...</wp:align>` Text content captured as
//!     h_align/v_align (alternative to posOffset). When align is
//!     used, pos_x/pos_y STAY at default 0.0 (no offset implied).
//!   - WrapType enum has FOUR variants (None, Square, Tight,
//!     TopAndBottom). Each maps from a distinct `<wp:wrap*>`
//!     Empty element. A regression that conflated Square and
//!     TopAndBottom would silently break layout of wrapped-vs-
//!     non-wrapped text around floating images.
//!   - srcRect l/t/r/b in 1/1000th percent units (val/1000 →
//!     percent). NOT raw percent, NOT 1/100,000 (EMU-style).
//!     A regression that used the wrong divisor would silently
//!     produce wildly wrong crop rectangles.
//!
//! Fixtures live in `tools/fixtures/anchored_image_samples/` and
//! are authored by
//! `tools/metrics/build_anchored_image_repro_fixtures.py`.

use std::fs;

use oxidocs_core::ir::{Block, Document, WrapType};
use oxidocs_core::parse_docx;

fn fixture_path(name: &str) -> std::path::PathBuf {
    let crate_root = std::env::current_dir().unwrap();
    let workspace_root = crate_root.parent().unwrap().parent().unwrap();
    workspace_root
        .join("tools")
        .join("fixtures")
        .join("anchored_image_samples")
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

#[test]
fn v1_anchor_pos_offset_emu_to_points_and_relative_strings() {
    let Some(doc) = load("v1_anchor_pos_offset.docx") else { return };
    let page = &doc.pages[0];

    // ROUTING: anchored image lives in Page.floating_images, NOT
    // Block::Image inside Page.blocks. This is the headline
    // distinction from inline images (S292).
    assert_eq!(
        page.floating_images.len(),
        1,
        "anchored image routes to Page.floating_images"
    );
    let inline_count = page
        .blocks
        .iter()
        .filter_map(|b| if let Block::Image(_) = b { Some(()) } else { None })
        .count();
    assert_eq!(
        inline_count, 0,
        "anchored image is NOT in Page.blocks as Block::Image"
    );

    let img = &page.floating_images[0];
    let pos = img
        .position
        .as_ref()
        .expect("anchored image must have FloatingPosition");

    // 914400 EMU / 12700 = 72.0pt. The /12700 divisor (EMU→pt) is
    // DIFFERENT from the inch divisor /914400. A regression mixing
    // them would produce x=0.0788pt (way wrong).
    assert!(
        (pos.x - 72.0).abs() < 0.001,
        "posOffset=914400 EMU → 72pt (val/12700), got {}",
        pos.x
    );
    assert!(
        (pos.y - 36.0).abs() < 0.001,
        "posOffset=457200 EMU → 36pt, got {}",
        pos.y
    );

    // relativeFrom captured verbatim — downstream layout dispatches
    // on these strings ("column" vs "margin" vs "page" matters).
    assert_eq!(
        pos.h_relative.as_deref(),
        Some("column"),
        "relativeFrom verbatim"
    );
    assert_eq!(pos.v_relative.as_deref(), Some("paragraph"));

    // When posOffset is used, align stays None.
    assert!(
        pos.h_align.is_none(),
        "posOffset path → h_align stays None"
    );
    assert!(pos.v_align.is_none());

    // wrapNone → WrapType::None (the variant, not the absence).
    assert_eq!(img.wrap_type, Some(WrapType::None));
    assert_eq!(img.alt_text.as_deref(), Some("positioned"));
}

#[test]
fn v1_anchor_pos_align_uses_string_branch_with_zero_offsets() {
    let Some(doc) = load("v1_anchor_pos_align.docx") else { return };
    let page = &doc.pages[0];
    assert_eq!(page.floating_images.len(), 1);
    let img = &page.floating_images[0];
    let pos = img.position.as_ref().expect("FloatingPosition required");

    // align branch is an ALTERNATIVE to posOffset. When align is
    // used, pos_x / pos_y STAY at 0.0 (no offset implied by align).
    assert_eq!(
        pos.x, 0.0,
        "align branch → pos.x stays 0.0 (no offset implied)"
    );
    assert_eq!(pos.y, 0.0);

    // align Text content captured as string. "center" and "top"
    // are downstream-dispatch keys.
    assert_eq!(
        pos.h_align.as_deref(),
        Some("center"),
        "<wp:align>center</wp:align> → h_align=Some(\"center\")"
    );
    assert_eq!(pos.v_align.as_deref(), Some("top"));

    assert_eq!(pos.h_relative.as_deref(), Some("margin"));
    assert_eq!(pos.v_relative.as_deref(), Some("page"));
}

#[test]
fn v1_anchor_wrap_square_pins_square_enum_variant() {
    let Some(doc) = load("v1_anchor_wrap_square.docx") else { return };
    let img = &doc.pages[0].floating_images[0];

    // wrapSquare → WrapType::Square. Distinct from None / Tight /
    // TopAndBottom — downstream layout dispatches on this enum.
    assert_eq!(
        img.wrap_type,
        Some(WrapType::Square),
        "<wp:wrapSquare/> → WrapType::Square"
    );
}

#[test]
fn v1_anchor_wrap_topandbottom_pins_distinct_enum_variant() {
    let Some(doc) = load("v1_anchor_wrap_topandbottom.docx") else { return };
    let img = &doc.pages[0].floating_images[0];

    // wrapTopAndBottom → WrapType::TopAndBottom. A regression that
    // conflated this with Square would silently change wrap behavior.
    assert_eq!(
        img.wrap_type,
        Some(WrapType::TopAndBottom),
        "<wp:wrapTopAndBottom/> → WrapType::TopAndBottom (distinct from Square)"
    );
}

#[test]
fn v1_anchor_crop_srcrect_thousandths_of_percent_divisor() {
    let Some(doc) = load("v1_anchor_crop_srcrect.docx") else { return };
    let img = &doc.pages[0].floating_images[0];

    let crop = img
        .crop
        .as_ref()
        .expect("srcRect with non-zero edges must populate crop");

    // val / 1000 → percent. 10000 → 10.0%. Distinct from raw
    // percent (val=10) and from EMU-style (val/100000).
    assert!(
        (crop.left - 10.0).abs() < 0.001,
        "l=10000 → 10.0% (val/1000), got {}",
        crop.left
    );
    assert!(
        (crop.top - 20.0).abs() < 0.001,
        "t=20000 → 20.0%, got {}",
        crop.top
    );
    assert!(
        (crop.right - 30.0).abs() < 0.001,
        "r=30000 → 30.0%, got {}",
        crop.right
    );
    assert!(
        (crop.bottom - 40.0).abs() < 0.001,
        "b=40000 → 40.0%, got {}",
        crop.bottom
    );
}

#[test]
fn all_five_fixtures_route_to_floating_images_not_block_image() {
    // Smoke + ROUTING: confirm every fixture produces exactly one
    // floating image and zero inline Block::Image. The routing
    // dispatch at parser/ooxml.rs:832-836 is the structural
    // distinction from inline images.
    let cases: &[&str] = &[
        "v1_anchor_pos_offset.docx",
        "v1_anchor_pos_align.docx",
        "v1_anchor_wrap_square.docx",
        "v1_anchor_wrap_topandbottom.docx",
        "v1_anchor_crop_srcrect.docx",
    ];
    for name in cases {
        let path = fixture_path(name);
        if !path.exists() {
            eprintln!("skipping {}", name);
            continue;
        }
        let data = fs::read(&path).unwrap();
        let doc = parse_docx(&data)
            .unwrap_or_else(|e| panic!("failed to parse {}: {:?}", name, e));
        let page = &doc.pages[0];
        assert_eq!(
            page.floating_images.len(),
            1,
            "{} should produce exactly 1 floating image",
            name
        );
        let inline = page
            .blocks
            .iter()
            .filter_map(|b| if let Block::Image(_) = b { Some(()) } else { None })
            .count();
        assert_eq!(
            inline, 0,
            "{} should have 0 inline Block::Image (anchored routes elsewhere)",
            name
        );
    }
}
