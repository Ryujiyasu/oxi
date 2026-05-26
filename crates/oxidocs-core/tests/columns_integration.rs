//! Integration tests: parse `<w:cols>` end-to-end and verify the
//! `Page.columns: Option<ColumnLayout>` field after `parse_docx`.
//!
//! `section_integration.rs` already covers the common num=2 + space
//! case (and the default-single = None case). This file deepens
//! coverage to the parser's non-obvious branches that S290 didn't
//! pin:
//!
//!   - [parser/ooxml.rs:5574](crates/oxidocs-core/src/parser/ooxml.rs#L5574)
//!     `columns = Some(ColumnLayout)` is gated by `if num > 1`. An
//!     explicit `<w:cols w:num="1"/>` in XML therefore still results
//!     in `Page.columns = None` — the parser collapses the trivial
//!     1-column case rather than carrying a degenerate ColumnLayout.
//!   - [parser/ooxml.rs:5540](crates/oxidocs-core/src/parser/ooxml.rs#L5540)
//!     `equalWidth="0"` (or `"false"`) → `equal_width = false`. Any
//!     other value (including absent) → `equal_width = true`.
//!   - [parser/ooxml.rs:5550](crates/oxidocs-core/src/parser/ooxml.rs#L5550)
//!     Child `<w:col w:w="X" w:space="Y"/>` entries populate the
//!     `ColumnDef { width, space }` vector. Width and space are
//!     converted from twips to points (val / 20).
//!   - When `<w:cols>` is missing from sectPr, `Page.columns = None`.
//!
//! Fixtures live in `tools/fixtures/columns_samples/` and are
//! authored by `tools/metrics/build_columns_repro_fixtures.py` (S307).

use std::fs;

use oxidocs_core::ir::Document;
use oxidocs_core::parse_docx;

fn fixture_path(name: &str) -> std::path::PathBuf {
    let crate_root = std::env::current_dir().unwrap();
    let workspace_root = crate_root.parent().unwrap().parent().unwrap();
    workspace_root
        .join("tools")
        .join("fixtures")
        .join("columns_samples")
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
fn v1_explicit_single_collapses_to_none() {
    // `<w:cols w:num="1"/>` is short-circuited by `if num > 1` in the
    // parser — the IR carries `Page.columns = None` even though the
    // XML element was present. This matches "no columns layout" as a
    // single canonical IR state regardless of how the docx encoded it.
    let Some(doc) = load("v1_explicit_single.docx") else { return };
    let p = &doc.pages[0];
    assert!(
        p.columns.is_none(),
        "explicit <w:cols w:num=\"1\"/> still collapses to Page.columns=None"
    );
}

#[test]
fn v1_three_equal_columns_with_36pt_space() {
    let Some(doc) = load("v1_three_equal.docx") else { return };
    let cols = doc.pages[0]
        .columns
        .as_ref()
        .expect("three-column fixture must populate columns");
    assert_eq!(cols.num, 3);
    let space = cols.space.expect("space attribute must be parsed");
    assert!((space - 36.0).abs() < 0.05, "720 twips → 36pt, got {}", space);
    // No explicit equalWidth attribute → defaults to true.
    assert!(cols.equal_width, "absent equalWidth defaults to true");
    // No per-column children declared (Empty-element form) → empty vec.
    assert!(
        cols.columns.is_empty(),
        "no <w:col> children → ColumnLayout.columns is empty"
    );
}

#[test]
fn v1_equal_width_false_flips_flag() {
    let Some(doc) = load("v1_equal_width_false.docx") else { return };
    let cols = doc.pages[0].columns.as_ref().expect("columns must be Some");
    assert_eq!(cols.num, 2);
    assert!(
        !cols.equal_width,
        "equalWidth=\"0\" → equal_width flips to false"
    );
}

#[test]
fn v1_per_column_defs_pins_widths_and_space() {
    // Start-element `<w:cols>` with two child `<w:col>` entries:
    //   col[0]: w=4000tw (200pt), space=240tw (12pt)
    //   col[1]: w=5000tw (250pt), no space attribute → None
    let Some(doc) = load("v1_per_column_defs.docx") else { return };
    let cols = doc.pages[0]
        .columns
        .as_ref()
        .expect("per-column fixture must populate columns");
    assert_eq!(cols.num, 2);
    assert!(!cols.equal_width);
    assert_eq!(cols.columns.len(), 2, "two <w:col> children → 2 ColumnDef");

    let c0 = &cols.columns[0];
    assert!((c0.width - 200.0).abs() < 0.05, "col[0] width = 200pt");
    assert_eq!(
        c0.space.map(|s| (s - 12.0).abs() < 0.05),
        Some(true),
        "col[0] space = 12pt"
    );

    let c1 = &cols.columns[1];
    assert!((c1.width - 250.0).abs() < 0.05, "col[1] width = 250pt");
    assert!(
        c1.space.is_none(),
        "col[1] has no w:space attribute → ColumnDef.space stays None"
    );
}

#[test]
fn v1_no_cols_leaves_page_columns_none() {
    // sectPr without `<w:cols>` → Page.columns = None. Same canonical
    // None state as v1_explicit_single (the 1-column collapse).
    let Some(doc) = load("v1_no_cols.docx") else { return };
    assert!(
        doc.pages[0].columns.is_none(),
        "sectPr without <w:cols> → Page.columns stays None"
    );
}

#[test]
fn all_five_fixtures_parse_with_expected_column_presence() {
    let cases: &[(&str, Option<u32>)] = &[
        ("v1_explicit_single.docx", None),
        ("v1_three_equal.docx", Some(3)),
        ("v1_equal_width_false.docx", Some(2)),
        ("v1_per_column_defs.docx", Some(2)),
        ("v1_no_cols.docx", None),
    ];
    for (name, expected_num) in cases {
        let path = fixture_path(name);
        if !path.exists() {
            eprintln!("skipping {}", name);
            continue;
        }
        let data = fs::read(&path).unwrap();
        let doc = parse_docx(&data)
            .unwrap_or_else(|e| panic!("failed to parse {}: {:?}", name, e));
        let actual = doc.pages[0].columns.as_ref().map(|c| c.num);
        assert_eq!(&actual, expected_num, "{} column num mismatch", name);
    }
}
