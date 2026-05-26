//! Integration tests: parse `<w:tbl><w:tblPr>...</w:tblPr></w:tbl>`
//! end-to-end and verify `Table.style: TableStyle` after `parse_docx`.
//!
//! `table_integration.rs` (S294) covers row/cell structure (gridSpan,
//! vMerge, shd). `cell_properties_integration.rs` (S310) covers tcPr.
//! This file fills the `<w:tblPr>` surface that no integration test
//! pinned: tblW (dxa vs pct conversion divisors), tblBorders
//! (has_inside_h flag, color="auto" suppression — OPPOSITE of cell
//! borders), tblLayout, tblInd, tblCellSpacing, jc, tblLook (with
//! noHBand inversion), and tblpPr (floating-table position).
//!
//! Parser code path tested:
//!   - [parser/ooxml.rs:4795](crates/oxidocs-core/src/parser/ooxml.rs#L4795)
//!     `parse_table_properties` end-to-end.

use std::fs;

use oxidocs_core::ir::{Block, Document, Table};
use oxidocs_core::parse_docx;

fn fixture_path(name: &str) -> std::path::PathBuf {
    let crate_root = std::env::current_dir().unwrap();
    let workspace_root = crate_root.parent().unwrap().parent().unwrap();
    workspace_root
        .join("tools")
        .join("fixtures")
        .join("table_properties_samples")
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

fn tables(doc: &Document) -> Vec<&Table> {
    doc.pages
        .iter()
        .flat_map(|p| p.blocks.iter())
        .filter_map(|b| if let Block::Table(t) = b { Some(t) } else { None })
        .collect()
}

#[test]
fn v1_tblw_dxa_pct_use_different_divisors() {
    let Some(doc) = load("v1_tblw_dxa_and_pct.docx") else { return };
    let ts = tables(&doc);
    assert_eq!(ts.len(), 2, "doc contains exactly two tables");

    // table[0]: tblW w=5000 type="dxa" → width = 250pt (val / 20).
    let dxa = &ts[0].style;
    assert_eq!(dxa.width_type.as_deref(), Some("dxa"));
    let wd = dxa.width.expect("dxa table must have width populated");
    assert!(
        (wd - 250.0).abs() < 0.5,
        "tblW w=5000 type=dxa → 250pt (twips/20), got {}",
        wd
    );

    // table[1]: tblW w=2500 type="pct" → width = 50.0 (val / 50).
    // The /50 divisor (50ths of a percent) is DIFFERENT from the
    // dxa /20 divisor. A regression that used /20 for pct would
    // produce width=125 instead of 50 — silently wrong.
    let pct = &ts[1].style;
    assert_eq!(pct.width_type.as_deref(), Some("pct"));
    let wp = pct.width.expect("pct table must have width populated");
    assert!(
        (wp - 50.0).abs() < 0.5,
        "tblW w=2500 type=pct → 50.0 (50ths of a percent, val/50), got {}",
        wp
    );
}

#[test]
fn v1_tbl_borders_has_inside_h_flag_and_color_auto_suppression() {
    let Some(doc) = load("v1_tbl_borders_inside_h.docx") else { return };
    let t = tables(&doc)[0];
    let s = &t.style;

    // tblBorders declared → explicit_borders flag set (gates the
    // b35-style "border-at-margin-minus-padding" offset).
    assert!(
        s.explicit_borders,
        "<w:tblBorders> directly under <w:tblPr> → explicit_borders=true"
    );

    // At least one non-suppressed border → border=true.
    assert!(s.border, "non-suppressed sides → border=true");

    // has_inside_h is the SIDE-SPECIFIC flag, set only when insideH is
    // present and non-suppressed. A regression that set it for top or
    // bottom (or unset it for insideH val=single) would silently lose
    // inter-row separator rendering.
    assert!(
        s.has_inside_h,
        "<w:insideH w:val=\"single\"/> → has_inside_h=true"
    );

    // tbl border color="auto" → border_color stays None (SUPPRESSION).
    // This is the OPPOSITE of cell borders (S310) where color="auto"
    // materializes to "000000". Pinning the divergence between the
    // two parser branches is the headline of this test.
    //
    // Top declared color="000000" before bottom declared "auto", so
    // border_color is set ONCE (the first-write-wins path at line
    // 4895 — `if border_color.is_none() { ... }`). The "000000"
    // hex must survive.
    assert_eq!(
        s.border_color.as_deref(),
        Some("000000"),
        "first explicit color wins; bottom's color=\"auto\" does not overwrite"
    );

    // sz=8 → 1.0pt width (1/8-pt units, same as cell borders).
    let bw = s.border_width.expect("sz=8 must populate border_width");
    assert!(
        (bw - 1.0).abs() < 0.001,
        "sz=8 → 1.0pt (val / 8), got {}",
        bw
    );

    assert_eq!(
        s.border_style.as_deref(),
        Some("single"),
        "any non-suppressed border materializes border_style=\"single\""
    );
}

#[test]
fn v1_tbl_layout_jc_indent_and_cell_spacing() {
    let Some(doc) = load("v1_tbl_layout_jc_indent.docx") else { return };
    let s = &tables(&doc)[0].style;

    assert_eq!(s.layout.as_deref(), Some("fixed"), "tblLayout type=fixed");
    assert_eq!(s.alignment.as_deref(), Some("center"), "jc val=center");

    let ind = s.indent.expect("tblInd populates indent");
    assert!(
        (ind - 36.0).abs() < 0.001,
        "tblInd w=720 → 36pt (twips/20), got {}",
        ind
    );

    let sp = s
        .cell_spacing
        .expect("tblCellSpacing populates cell_spacing");
    assert!(
        (sp - 5.0).abs() < 0.001,
        "tblCellSpacing w=100 → 5pt (twips/20), got {}",
        sp
    );
}

#[test]
fn v1_tbl_look_attr_form_inverts_no_band_flags() {
    let Some(doc) = load("v1_tbl_look_attr_form.docx") else { return };
    let s = &tables(&doc)[0].style;

    let look = s.tbl_look.as_ref().expect("tblLook present → Some");
    assert!(look.first_row, "firstRow=\"1\" → first_row=true");
    assert!(!look.last_row, "lastRow=\"0\" → last_row=false");
    assert!(
        look.first_column,
        "firstColumn=\"1\" → first_column=true"
    );
    assert!(
        !look.last_column,
        "lastColumn=\"0\" → last_column=false"
    );

    // noHBand=1 → banded_rows=false (INVERTED at parser/ooxml.rs:4990).
    // noVBand=0 → banded_columns=true.
    // The inversion is non-obvious: a regression that copied the value
    // directly would silently invert visual banding on every doc.
    assert!(
        !look.banded_rows,
        "noHBand=\"1\" → banded_rows=false (INVERTED, NOT copied verbatim)"
    );
    assert!(
        look.banded_columns,
        "noVBand=\"0\" → banded_columns=true (INVERTED)"
    );
}

#[test]
fn v1_tbl_pos_floating_pins_tblp_pr_axes_and_anchors() {
    let Some(doc) = load("v1_tbl_pos_floating.docx") else { return };
    let s = &tables(&doc)[0].style;

    let p = s.position.as_ref().expect("tblpPr present → position Some");

    // tblpX/Y in twips → pt (val / 20).
    assert!(
        (p.x - 72.0).abs() < 0.001,
        "tblpX=1440 → 72pt, got {}",
        p.x
    );
    assert!(
        (p.y - 36.0).abs() < 0.001,
        "tblpY=720 → 36pt, got {}",
        p.y
    );

    // Anchors stored verbatim — these are enum-like strings ("margin",
    // "page", "text") that downstream layout dispatches on.
    assert_eq!(p.h_anchor.as_deref(), Some("margin"));
    assert_eq!(p.v_anchor.as_deref(), Some("page"));

    // All four from-text distances are twips → pt.
    assert!(
        (p.left_from_text - 9.0).abs() < 0.001,
        "leftFromText=180 → 9pt"
    );
    assert!(
        (p.right_from_text - 9.0).abs() < 0.001,
        "rightFromText=180 → 9pt"
    );
    assert!(
        (p.top_from_text - 10.0).abs() < 0.001,
        "topFromText=200 → 10pt"
    );
    assert!(
        (p.bottom_from_text - 10.0).abs() < 0.001,
        "bottomFromText=200 → 10pt"
    );
}

#[test]
fn all_five_fixtures_parse_with_expected_table_counts() {
    let cases: &[(&str, usize)] = &[
        ("v1_tblw_dxa_and_pct.docx", 2),
        ("v1_tbl_borders_inside_h.docx", 1),
        ("v1_tbl_layout_jc_indent.docx", 1),
        ("v1_tbl_look_attr_form.docx", 1),
        ("v1_tbl_pos_floating.docx", 1),
    ];
    for (name, expected_count) in cases {
        let path = fixture_path(name);
        if !path.exists() {
            eprintln!("skipping {}", name);
            continue;
        }
        let data = fs::read(&path).unwrap();
        let doc = parse_docx(&data)
            .unwrap_or_else(|e| panic!("failed to parse {}: {:?}", name, e));
        assert_eq!(
            tables(&doc).len(),
            *expected_count,
            "{} table count",
            name
        );
    }
}
