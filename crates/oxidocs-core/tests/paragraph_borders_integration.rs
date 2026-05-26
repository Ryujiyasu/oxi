//! Integration tests: parse `<w:pBdr>` and verify
//! `Paragraph.style.borders` populates with the expected per-side
//! `BorderDef` values.
//!
//! Parser code path tested:
//!   - [parser/ooxml.rs:1915](crates/oxidocs-core/src/parser/ooxml.rs#L1915)
//!     routes `<w:pBdr>` inside `<w:pPr>` to
//!     [parse_paragraph_borders](crates/oxidocs-core/src/parser/ooxml.rs#L2362),
//!   - [parser/ooxml.rs:2394](crates/oxidocs-core/src/parser/ooxml.rs#L2394)
//!     `parse_border_attrs` maps each `<w:top>/<w:bottom>/<w:left>/<w:right>/<w:between>`
//!     element's attributes:
//!       - `w:val="none"|"nil"` → None (border suppressed)
//!       - `w:val=<style>`     → `BorderDef.style` verbatim
//!       - `w:sz=<eighths-pt>` → `BorderDef.width = sz / 8.0` in pt
//!       - `w:color="auto"`    → `BorderDef.color = Some("000000")`
//!       - `w:color=<hex>`     → `BorderDef.color = Some(<hex>)`
//!   - `<w:start>` / `<w:end>` aliases map to `borders.left` / `borders.right`.
//!
//! Fixtures live in `tools/fixtures/paragraph_borders_samples/` and are
//! authored by `tools/metrics/build_paragraph_borders_repro_fixtures.py` (S302).

use std::fs;

use oxidocs_core::ir::{Block, Document, Paragraph};
use oxidocs_core::parser::parse_docx;

fn fixture_path(name: &str) -> std::path::PathBuf {
    let crate_root = std::env::current_dir().unwrap();
    let workspace_root = crate_root.parent().unwrap().parent().unwrap();
    workspace_root
        .join("tools")
        .join("fixtures")
        .join("paragraph_borders_samples")
        .join(name)
}

fn first_paragraph(doc: &Document) -> &Paragraph {
    doc.pages
        .iter()
        .flat_map(|p| p.blocks.iter())
        .find_map(|b| if let Block::Paragraph(p) = b { Some(p) } else { None })
        .expect("first paragraph")
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
fn v1_all_sides_populates_top_bottom_left_right() {
    let Some(doc) = load("v1_all_sides.docx") else { return };
    let p = first_paragraph(&doc);
    let b = p.style.borders.as_ref().expect("paragraph has borders");
    assert!(b.top.is_some(), "top set");
    assert!(b.bottom.is_some(), "bottom set");
    assert!(b.left.is_some(), "left set");
    assert!(b.right.is_some(), "right set");
    // `between` was not declared, so it stays None.
    assert!(b.between.is_none(), "between absent");

    // Each side: style="single", width=0.5pt (sz=4), color="000000".
    for (label, opt) in [
        ("top", &b.top),
        ("bottom", &b.bottom),
        ("left", &b.left),
        ("right", &b.right),
    ] {
        let bd = opt.as_ref().unwrap_or_else(|| panic!("{} present", label));
        assert_eq!(bd.style, "single", "{} style", label);
        assert!((bd.width - 0.5).abs() < 1e-4, "{} width {}", label, bd.width);
        assert_eq!(bd.color.as_deref(), Some("000000"), "{} color", label);
    }
}

#[test]
fn v1_between_surfaces_in_borders() {
    let Some(doc) = load("v1_between.docx") else { return };
    let p = first_paragraph(&doc);
    let b = p.style.borders.as_ref().expect("paragraph has borders");
    let between = b.between.as_ref().expect("between border set");
    assert_eq!(between.style, "single");
    assert!((between.width - 0.5).abs() < 1e-4);
    // The other sides were not declared.
    assert!(b.top.is_none() && b.bottom.is_none() && b.left.is_none() && b.right.is_none());
}

#[test]
fn v1_start_end_aliases_map_to_left_and_right() {
    // OOXML newer/bidi-friendly naming: `<w:start>`/`<w:end>` aliases.
    // Parser maps them onto the same `left`/`right` slots as the
    // legacy `<w:left>`/`<w:right>` elements.
    let Some(doc) = load("v1_start_end_aliases.docx") else { return };
    let p = first_paragraph(&doc);
    let b = p.style.borders.as_ref().expect("paragraph has borders");
    assert!(b.left.is_some(), "`<w:start>` lands on borders.left");
    assert!(b.right.is_some(), "`<w:end>` lands on borders.right");
    // Top/bottom/between were not declared.
    assert!(b.top.is_none());
    assert!(b.bottom.is_none());
    assert!(b.between.is_none());
}

#[test]
fn v1_color_and_width_preserves_hex_and_maps_auto() {
    // top: sz=12 (1.5pt), color="FF0000" (verbatim hex)
    // bottom: sz=4 (0.5pt), color="auto" → "000000"
    let Some(doc) = load("v1_color_and_width.docx") else { return };
    let p = first_paragraph(&doc);
    let b = p.style.borders.as_ref().expect("paragraph has borders");
    let top = b.top.as_ref().expect("top border set");
    assert!((top.width - 1.5).abs() < 1e-4, "top width {}", top.width);
    assert_eq!(
        top.color.as_deref(),
        Some("FF0000"),
        "hex color preserved verbatim (no normalization)"
    );

    let bot = b.bottom.as_ref().expect("bottom border set");
    assert!((bot.width - 0.5).abs() < 1e-4);
    assert_eq!(
        bot.color.as_deref(),
        Some("000000"),
        "`auto` color materialized as black hex"
    );
}

#[test]
fn v1_none_and_nil_vals_suppress_the_side() {
    // val="none" and val="nil" → BorderDef returned as None even
    // though the element is present. This is the parser's way of
    // letting a child paragraph opt OUT of an inherited border.
    let Some(doc) = load("v1_none_suppresses.docx") else { return };
    let p = first_paragraph(&doc);
    // pBdr was emitted with two children, so Paragraph.borders is
    // Some(ParagraphBorders { ... }) — but both sides are None.
    let b = p.style.borders.as_ref().expect("paragraph has borders");
    assert!(b.top.is_none(), "val=none → no top border");
    assert!(b.bottom.is_none(), "val=nil → no bottom border");
    assert!(b.left.is_none() && b.right.is_none() && b.between.is_none());
}

#[test]
fn all_five_fixtures_parse_with_expected_border_presence() {
    let cases: &[(&str, [bool; 5])] = &[
        // (name, [top, bottom, left, right, between])
        ("v1_all_sides.docx", [true, true, true, true, false]),
        ("v1_between.docx", [false, false, false, false, true]),
        ("v1_start_end_aliases.docx", [false, false, true, true, false]),
        ("v1_color_and_width.docx", [true, true, false, false, false]),
        ("v1_none_suppresses.docx", [false, false, false, false, false]),
    ];
    for (name, expected) in cases {
        let path = fixture_path(name);
        if !path.exists() {
            eprintln!("skipping {}", name);
            continue;
        }
        let data = fs::read(&path).unwrap();
        let doc = parse_docx(&data)
            .unwrap_or_else(|e| panic!("failed to parse {}: {:?}", name, e));
        let p = first_paragraph(&doc);
        let b = p.style.borders.as_ref().expect("borders present");
        let got = [
            b.top.is_some(),
            b.bottom.is_some(),
            b.left.is_some(),
            b.right.is_some(),
            b.between.is_some(),
        ];
        assert_eq!(&got, expected, "{} border presence", name);
    }
}
