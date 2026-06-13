// This Source Code Form is subject to the terms of the Mozilla Public
// License, v. 2.0. If a copy of the MPL was not distributed with this
// file, You can obtain one at https://mozilla.org/MPL/2.0/.

//! Integration tests: parse `<w:sectPr>` (page size, margins, columns)
//! end-to-end and verify `Page.size` / `Page.margin` / `Page.columns`
//! populate correctly.
//!
//! Parser code paths tested:
//! - [parser/ooxml.rs:5571](crates/oxidocs-core/src/parser/ooxml.rs#L5571)
//!   `<w:pgSz>`: parses width/height in twentieths-of-a-pt; swaps width/height
//!   when `orient="landscape"` AND width<height.
//! - [parser/ooxml.rs:5596](crates/oxidocs-core/src/parser/ooxml.rs#L5596)
//!   `<w:pgMar>`: parses top/bottom/left/right margins.
//! - [parser/ooxml.rs:5766](crates/oxidocs-core/src/parser/ooxml.rs#L5766)
//!   `<w:cols>`: parses column count and inter-column space.
//!
//! Fixtures live in `tools/fixtures/section_samples/` and are authored by
//! `tools/metrics/build_section_repro_fixtures.py` (S290).

use std::fs;

use oxidocs_core::parser::parse_docx;

fn fixture_path(name: &str) -> std::path::PathBuf {
    let crate_root = std::env::current_dir().unwrap();
    let workspace_root = crate_root.parent().unwrap().parent().unwrap();
    workspace_root.join("tools").join("fixtures").join("section_samples").join(name)
}

fn load(name: &str) -> Option<oxidocs_core::ir::Document> {
    let path = fixture_path(name);
    if !path.exists() {
        eprintln!("skipping: {} not found", path.display());
        return None;
    }
    let data = fs::read(&path).expect("read fixture");
    Some(parse_docx(&data).expect("parse fixture"))
}

#[test]
fn v1_a4_portrait_has_595_by_841_with_72pt_margins() {
    let Some(doc) = load("v1_a4_portrait.docx") else { return };
    assert!(!doc.pages.is_empty());
    let p = &doc.pages[0];
    // A4 = 11906 × 16838 twentieths-of-pt → 595.3 × 841.9 pt
    assert!((p.size.width - 595.3).abs() < 0.1, "width {}", p.size.width);
    assert!((p.size.height - 841.9).abs() < 0.1, "height {}", p.size.height);
    // 1440 twips = 72pt (1 inch)
    assert!((p.margin.top - 72.0).abs() < 0.1);
    assert!((p.margin.bottom - 72.0).abs() < 0.1);
    assert!((p.margin.left - 72.0).abs() < 0.1);
    assert!((p.margin.right - 72.0).abs() < 0.1);
    assert!(p.columns.is_none(), "default = no explicit column layout");
}

#[test]
fn v1_a4_landscape_swaps_width_and_height() {
    // Word stores landscape as orient="landscape" with the raw width
    // attribute still being the SMALLER dimension. The parser must swap
    // so width > height for landscape display.
    let Some(doc) = load("v1_a4_landscape.docx") else { return };
    let p = &doc.pages[0];
    assert!((p.size.width - 841.9).abs() < 0.1,
        "landscape width should be the larger dimension, got {}", p.size.width);
    assert!((p.size.height - 595.3).abs() < 0.1,
        "landscape height should be the smaller dimension, got {}", p.size.height);
    assert!(p.size.width > p.size.height,
        "landscape ⇒ width > height invariant");
}

#[test]
fn v1_custom_margins_each_side_distinct() {
    // top=2000tw → 100pt, bottom=1000tw → 50pt,
    // left=1600tw → 80pt, right=800tw → 40pt
    let Some(doc) = load("v1_custom_margins.docx") else { return };
    let p = &doc.pages[0];
    assert!((p.margin.top - 100.0).abs() < 0.5,    "top {}", p.margin.top);
    assert!((p.margin.bottom - 50.0).abs() < 0.5,  "bottom {}", p.margin.bottom);
    assert!((p.margin.left - 80.0).abs() < 0.5,    "left {}", p.margin.left);
    assert!((p.margin.right - 40.0).abs() < 0.5,   "right {}", p.margin.right);
}

#[test]
fn v1_two_columns_has_column_layout_with_18pt_space() {
    // <w:cols w:num="2" w:space="360"/> → ColumnLayout { num: 2,
    // space: Some(18.0pt) }
    let Some(doc) = load("v1_two_columns.docx") else { return };
    let p = &doc.pages[0];
    let cols = p.columns.as_ref().expect("two-column fixture must have columns");
    assert_eq!(cols.num, 2, "column count");
    let space = cols.space.expect("space should be parsed from w:space");
    assert!((space - 18.0).abs() < 0.1, "expected ~18pt space, got {}", space);
}

#[test]
fn all_four_fixtures_parse_with_expected_section_shape() {
    let cases: &[(&str, f32, f32, u32)] = &[
        // (filename, width_pt, height_pt, expected_columns_num)
        ("v1_a4_portrait.docx",   595.3, 841.9, 1),
        ("v1_a4_landscape.docx",  841.9, 595.3, 1),
        ("v1_custom_margins.docx", 595.3, 841.9, 1),
        ("v1_two_columns.docx",   595.3, 841.9, 2),
    ];
    for (name, exp_w, exp_h, exp_cols) in cases {
        let path = fixture_path(name);
        if !path.exists() {
            eprintln!("skipping {}", name);
            continue;
        }
        let data = fs::read(&path).unwrap();
        let doc = parse_docx(&data)
            .unwrap_or_else(|e| panic!("failed to parse {}: {:?}", name, e));
        let p = &doc.pages[0];
        assert!((p.size.width - exp_w).abs() < 0.2, "{} width", name);
        assert!((p.size.height - exp_h).abs() < 0.2, "{} height", name);
        let actual_cols = p.columns.as_ref().map(|c| c.num).unwrap_or(1);
        assert_eq!(actual_cols, *exp_cols, "{} column count", name);
    }
}
