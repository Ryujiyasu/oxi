// This Source Code Form is subject to the terms of the Mozilla Public
// License, v. 2.0. If a copy of the MPL was not distributed with this
// file, You can obtain one at https://mozilla.org/MPL/2.0/.

//! Integration tests: parse `<w:tc><w:tcPr>...</w:tcPr></w:tc>` end-to-end
//! and verify `TableCell.{borders, margins, width, v_align, shading}`
//! after `parse_docx`.
//!
//! `table_integration.rs` (S294) covers gridSpan, vMerge, and shd hex
//! storage. `vertical_integration.rs` covers `textDirection`. This
//! file fills the remaining cell-properties surface in
//! [parser/ooxml.rs:5246](crates/oxidocs-core/src/parser/ooxml.rs#L5246)
//! `parse_cell_properties`:
//!
//!   - tcBorders: top/bottom/left/right + `<w:start>`/`<w:end>` ALIAS
//!     routing to left/right (parser/ooxml.rs:5365-5366). Width = sz/8
//!     (1/8 pt units), color="auto" materializes to "000000" (NOT the
//!     literal "auto"), val="none"/"nil" SUPPRESSES the side — but on
//!     the CELL path it is stored as the S482 `{style:"none"}` SENTINEL,
//!     not as `None` (see `v1_tc_borders_pins_sz_color_and_none_suppression`).
//!   - tcMar: top/bottom/left/right in twips → pt (val/20). All four
//!     distinct values catch a mis-route between sides.
//!   - tcW: cell width in twips → pt (val/20).
//!   - vAlign: stored verbatim as the val string (top/center/bottom).
//!   - shd: `<w:shd w:fill="auto"/>` → shading=None (SUPPRESSION,
//!     parser/ooxml.rs:5312). Explicit hex stored verbatim. Symmetric
//!     with the rPr color="auto"→None branch pinned in S309 but on the
//!     cell-level shd field.
//!
//! Fixtures live in `tools/fixtures/cell_properties_samples/` and are
//! authored by `tools/metrics/build_cell_properties_repro_fixtures.py`.

use std::fs;

use oxidocs_core::ir::{Block, Document, Table, TableCell};
use oxidocs_core::parse_docx;

fn fixture_path(name: &str) -> std::path::PathBuf {
    let crate_root = std::env::current_dir().unwrap();
    let workspace_root = crate_root.parent().unwrap().parent().unwrap();
    workspace_root
        .join("tools")
        .join("fixtures")
        .join("cell_properties_samples")
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

fn first_table(doc: &Document) -> &Table {
    doc.pages
        .iter()
        .flat_map(|p| p.blocks.iter())
        .find_map(|b| if let Block::Table(t) = b { Some(t) } else { None })
        .expect("first table")
}

fn cell_text(cell: &TableCell) -> String {
    cell.blocks
        .iter()
        .filter_map(|b| if let Block::Paragraph(p) = b { Some(p) } else { None })
        .flat_map(|p| p.runs.iter())
        .map(|r| r.text.as_str())
        .collect::<String>()
}

#[test]
fn v1_tc_borders_pins_sz_color_and_none_suppression() {
    let Some(doc) = load("v1_tc_borders_lrtb.docx") else { return };
    let t = first_table(&doc);
    let cell0 = &t.rows[0].cells[0];
    assert_eq!(cell_text(cell0), "border-cell");

    let b = cell0.borders.as_ref().expect("tcBorders must populate borders");

    // top: val=single, sz=8 → width=1.0pt, color=000000 verbatim.
    let top = b.top.as_ref().expect("top border present");
    assert_eq!(top.style, "single", "val stored verbatim");
    assert!(
        (top.width - 1.0).abs() < 0.001,
        "sz=8 → 1.0pt (sz/8), got {}",
        top.width
    );
    assert_eq!(top.color.as_deref(), Some("000000"));

    // bottom: sz=16 → 2.0pt; color="auto" materializes to "000000"
    // (NOT the literal "auto" — parser/ooxml.rs:2415). A consumer
    // that compares hex strings would silently break if "auto" leaked.
    let bottom = b.bottom.as_ref().expect("bottom border present");
    assert!(
        (bottom.width - 2.0).abs() < 0.001,
        "sz=16 → 2.0pt, got {}",
        bottom.width
    );
    assert_eq!(
        bottom.color.as_deref(),
        Some("000000"),
        "color=auto → materialized as \"000000\" hex, NOT stored as literal \"auto\""
    );

    // left: dashed style, FF0000 explicit color (no auto-materialization)
    let left = b.left.as_ref().expect("left border present");
    assert_eq!(left.style, "dashed");
    assert_eq!(left.color.as_deref(), Some("FF0000"));

    // right: val="none" → SUPPRESSED, but stored as the S482 SENTINEL
    // `Some(BorderDef { style: "none", width: 0.0 })`, NOT as `None`.
    //
    // ★S482 (and S911, which depends on it) REQUIRE the sentinel — do not
    // "fix" this back to None. `parse_border_attrs` itself does return None
    // for val="none"/"nil" (parser/ooxml.rs:3297), but `parse_cell_borders`
    // (parser/ooxml.rs:7340) re-wraps an EXPLICIT nil as the sentinel,
    // because on the CELL path the two cases are semantically DIFFERENT
    // (ECMA-376: a cell border overrides the table border):
    //   - edge ABSENT       → inherit the table-level border (tblBorders)
    //   - edge EXPLICIT nil → suppress; draw NOTHING, do not inherit
    // Collapsing both to `None` loses that distinction and makes an
    // explicit nil fall through to the table border. Consumers that read
    // the sentinel:
    //   - S482  layout/mod.rs:22519 `resolve_border` — `style == "none"`
    //           returns (None, 0.0, None) instead of the table fallback
    //           (31420af: nil top/left over an all-single tblBorders drew
    //           spurious rules).
    //   - S911  layout/mod.rs:16346/16379/16455 `rowbox2_border_pad` — an
    //           explicitly-dead row boundary contributes NO border-box pad
    //           (legal 0001482d: +1.0/row → the wp217-220 tail).
    // The sentinel is opt-out via OXI_S482_DISABLE (which restores the
    // pre-S482 `None`); this test pins the DEFAULT behavior.
    let right = b.right.as_ref().expect("explicit nil is kept as the S482 sentinel");
    assert_eq!(
        right.style, "none",
        "val=\"none\" → S482 sentinel style=\"none\" (suppress, do NOT inherit the table border)"
    );
    assert_eq!(right.width, 0.0, "a suppressed edge has zero width");

    // The second cell has no tcBorders → borders stays None.
    // (This is the ABSENT case the sentinel must stay distinguishable from.)
    let cell1 = &t.rows[0].cells[1];
    assert!(
        cell1.borders.is_none(),
        "cell with no tcBorders → borders=None"
    );
}

#[test]
fn v1_tc_borders_start_end_route_to_left_and_right() {
    // OOXML's newer/bidi-friendly `<w:start>`/`<w:end>` are ALIASES
    // for `<w:left>`/`<w:right>` (parser/ooxml.rs:5365-5366 match arm).
    // A regression that introduced a separate `borders.start`/`.end`
    // field would silently lose the styling on docs that emit aliases.
    let Some(doc) = load("v1_tc_borders_start_end_aliases.docx") else { return };
    let t = first_table(&doc);
    let cell = &t.rows[0].cells[0];

    let b = cell.borders.as_ref().expect("tcBorders must populate");

    let left = b.left.as_ref().expect("`<w:start>` lands on borders.left");
    assert_eq!(left.style, "single");
    assert_eq!(left.color.as_deref(), Some("0000FF"));
    assert!((left.width - 1.0).abs() < 0.001, "sz=8 → 1.0pt");

    let right = b.right.as_ref().expect("`<w:end>` lands on borders.right");
    assert_eq!(right.style, "double");
    assert_eq!(right.color.as_deref(), Some("00FF00"));
    assert!(
        (right.width - 1.5).abs() < 0.001,
        "sz=12 → 1.5pt, got {}",
        right.width
    );

    // top/bottom were not declared → stay None.
    assert!(b.top.is_none());
    assert!(b.bottom.is_none());
}

#[test]
fn v1_tc_margins_twips_to_pt_per_side() {
    let Some(doc) = load("v1_tc_margins.docx") else { return };
    let t = first_table(&doc);
    let cell = &t.rows[0].cells[0];

    let m = cell.margins.as_ref().expect("tcMar must populate margins");

    // All four sides distinct so a top/bottom or left/right swap is caught.
    assert!(
        (m.top.unwrap() - 5.0).abs() < 0.001,
        "top w=100 → 5.0pt (twips/20)"
    );
    assert!(
        (m.bottom.unwrap() - 10.0).abs() < 0.001,
        "bottom w=200 → 10.0pt"
    );
    assert!(
        (m.left.unwrap() - 15.0).abs() < 0.001,
        "left w=300 → 15.0pt"
    );
    assert!(
        (m.right.unwrap() - 20.0).abs() < 0.001,
        "right w=400 → 20.0pt"
    );
}

#[test]
fn v1_tc_width_valign_distinguish_three_cells() {
    let Some(doc) = load("v1_tc_width_valign.docx") else { return };
    let t = first_table(&doc);
    let row = &t.rows[0];
    assert_eq!(row.cells.len(), 3);

    let c0 = &row.cells[0];
    let c1 = &row.cells[1];
    let c2 = &row.cells[2];

    // cell[0]: explicit tcW=3000tw → width=150pt; vAlign=center.
    let w0 = c0.width.expect("cell[0] must have width");
    assert!(
        (w0 - 150.0).abs() < 0.5,
        "tcW=3000tw → 150pt (twips/20), got {}",
        w0
    );
    assert_eq!(c0.v_align.as_deref(), Some("center"));

    // cell[1]: tcW=2000tw → 100pt; vAlign=bottom.
    let w1 = c1.width.expect("cell[1] must have width");
    assert!(
        (w1 - 100.0).abs() < 0.5,
        "tcW=2000tw → 100pt, got {}",
        w1
    );
    assert_eq!(c1.v_align.as_deref(), Some("bottom"));

    // cell[2]: no tcW override. width MAY be populated by the parser
    // from the grid (it's the rendered width); v_align stays None
    // since no vAlign was declared and the field has no default.
    assert!(
        c2.v_align.is_none(),
        "cell with no vAlign → v_align=None (no default)"
    );
}

#[test]
fn v1_tc_shd_auto_suppression_vs_explicit_hex() {
    // Pins the parser/ooxml.rs:5312 branch: `<w:shd w:fill="auto"/>`
    // is the "no shading" sentinel and SUPPRESSES storage. An explicit
    // hex value is stored verbatim. Symmetric with rPr color="auto"
    // (pinned in S309) but on the cell-level shd field.
    let Some(doc) = load("v1_tc_shd_auto_suppression.docx") else { return };
    let t = first_table(&doc);
    let row = &t.rows[0];

    assert_eq!(cell_text(&row.cells[0]), "shd-auto");
    assert!(
        row.cells[0].shading.is_none(),
        "fill=\"auto\" SUPPRESSES shading (NOT stored as literal \"auto\")"
    );

    assert_eq!(cell_text(&row.cells[1]), "shd-red");
    assert_eq!(
        row.cells[1].shading.as_deref(),
        Some("FF0000"),
        "explicit hex fill stored verbatim"
    );
}

#[test]
fn all_five_fixtures_parse_with_expected_cell_shape() {
    // Smoke + structure: each fixture parses and the first row's
    // cell count matches the structure we authored. Catches a
    // future regression that drops or merges cells while tcPr is
    // populated.
    let cases: &[(&str, usize)] = &[
        ("v1_tc_borders_lrtb.docx", 2),
        ("v1_tc_borders_start_end_aliases.docx", 1),
        ("v1_tc_margins.docx", 1),
        ("v1_tc_width_valign.docx", 3),
        ("v1_tc_shd_auto_suppression.docx", 2),
    ];
    for (name, expected_cells) in cases {
        let path = fixture_path(name);
        if !path.exists() {
            eprintln!("skipping {}", name);
            continue;
        }
        let data = fs::read(&path).unwrap();
        let doc = parse_docx(&data)
            .unwrap_or_else(|e| panic!("failed to parse {}: {:?}", name, e));
        let t = first_table(&doc);
        assert_eq!(
            t.rows[0].cells.len(),
            *expected_cells,
            "{} row[0] cell count",
            name
        );
    }
}
