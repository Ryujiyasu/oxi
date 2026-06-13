// This Source Code Form is subject to the terms of the Mozilla Public
// License, v. 2.0. If a copy of the MPL was not distributed with this
// file, You can obtain one at https://mozilla.org/MPL/2.0/.

//! Integration tests: parse `<w:tr><w:trPr>...</w:trPr></w:tr>` end-to-end
//! and verify `TableRow.{height, height_rule, header, cant_split,
//! grid_before, cell_margins_override}` after `parse_docx`.
//!
//! Completes the properties-hierarchy coverage (run S309 / cell S310 /
//! tbl S311 / row S312). Parser code path tested:
//!   - [parser/ooxml.rs:5059](crates/oxidocs-core/src/parser/ooxml.rs#L5059)
//!     `parse_table_row` end-to-end (trHeight, tblHeader, cantSplit,
//!     gridBefore, tblPrEx > tblCellMar override).
//!
//! Non-obvious behaviors pinned:
//!   - trHeight val/20 (twips → pt) AND hRule stored verbatim. The
//!     two height rules ("exact" vs "atLeast") drive RADICALLY
//!     different layout policy — clip-to-height vs grow-past-height.
//!     A regression that dropped hRule storage would conflate them.
//!   - tblHeader / cantSplit / gridBefore are INDEPENDENT flags
//!     (all can co-occur on the same row).
//!   - gridBefore=2 lets a row START at grid column 2 — physical
//!     cell count (1) differs from logical grid position. A
//!     regression that consumed gridBefore as cell-count would
//!     misalign the entire table.
//!   - tblPrEx > tblCellMar populates ROW.cell_margins_override
//!     (parser/ooxml.rs:5116), a SEPARATE field from the table-level
//!     default_cell_margins pinned in S311. Both shapes exist in
//!     the IR because Word allows per-row escape from the table
//!     default.
//!   - tblPrEx > tblCellMar `<w:start>` / `<w:end>` ALIASES route
//!     to left/right at parser/ooxml.rs:5104-5105 — same alias
//!     branch as cell-level tcBorders pinned in S310, but at the
//!     row-margin code path.
//!
//! Fixtures live in `tools/fixtures/row_properties_samples/` and are
//! authored by `tools/metrics/build_row_properties_repro_fixtures.py`.

use std::fs;

use oxidocs_core::ir::{Block, Document, Table};
use oxidocs_core::parse_docx;

fn fixture_path(name: &str) -> std::path::PathBuf {
    let crate_root = std::env::current_dir().unwrap();
    let workspace_root = crate_root.parent().unwrap().parent().unwrap();
    workspace_root
        .join("tools")
        .join("fixtures")
        .join("row_properties_samples")
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

#[test]
fn v1_tr_height_exact_pins_val_and_rule() {
    let Some(doc) = load("v1_tr_height_exact.docx") else { return };
    let row = &first_table(&doc).rows[0];

    let h = row.height.expect("trHeight must populate height");
    assert!(
        (h - 20.0).abs() < 0.001,
        "trHeight val=400 → 20pt (twips/20), got {}",
        h
    );

    // hRule stored VERBATIM as "exact" — the layout dispatches on
    // this exact string. A regression that normalized to (say) an
    // enum at parse-time without keeping the string would still pass
    // a "row has height_rule" check but break the layout's match
    // arms.
    assert_eq!(
        row.height_rule.as_deref(),
        Some("exact"),
        "hRule=\"exact\" stored verbatim — drives clip-to-height policy"
    );
}

#[test]
fn v1_tr_height_atleast_distinct_from_exact() {
    let Some(doc) = load("v1_tr_height_atleast.docx") else { return };
    let row = &first_table(&doc).rows[0];

    let h = row.height.expect("trHeight must populate");
    assert!(
        (h - 30.0).abs() < 0.001,
        "trHeight val=600 → 30pt, got {}",
        h
    );

    // "atLeast" is the OPPOSITE policy from "exact" — content
    // overflow GROWS the row past the declared height. A
    // regression that lost the distinction would silently over-
    // clip OR over-grow swathes of rows.
    assert_eq!(
        row.height_rule.as_deref(),
        Some("atLeast"),
        "hRule=\"atLeast\" stored verbatim — drives grow-past-height policy"
    );
}

#[test]
fn v1_tr_header_cant_split_are_independent_flags() {
    let Some(doc) = load("v1_tr_header_cant_split.docx") else { return };
    let t = first_table(&doc);

    // Row 0 has BOTH tblHeader and cantSplit — independent flags.
    let r0 = &t.rows[0];
    assert!(r0.header, "<w:tblHeader/> → header=true");
    assert!(r0.cant_split, "<w:cantSplit/> → cant_split=true");

    // Row 1 has NEITHER → both flags default false (no leak across rows).
    let r1 = &t.rows[1];
    assert!(
        !r1.header,
        "subsequent row without tblHeader → header=false (no leak from row 0)"
    );
    assert!(
        !r1.cant_split,
        "subsequent row without cantSplit → cant_split=false"
    );
}

#[test]
fn v1_tr_grid_before_skips_leading_columns() {
    let Some(doc) = load("v1_tr_grid_before.docx") else { return };
    let t = first_table(&doc);

    // Row 0: gridBefore=2 — the row STARTS at grid column 2.
    // Physical cell count (1) ≠ logical grid position (column 2).
    // A regression that consumed gridBefore as a cell-count would
    // misalign the whole table.
    let r0 = &t.rows[0];
    assert_eq!(
        r0.grid_before, 2,
        "<w:gridBefore w:val=\"2\"/> → grid_before=2"
    );
    assert_eq!(
        r0.cells.len(),
        1,
        "physical cell count = 1 (only the cell at column 2 is materialized)"
    );

    // Row 1: no gridBefore → default 0; full 3-cell row.
    let r1 = &t.rows[1];
    assert_eq!(r1.grid_before, 0, "default grid_before = 0 (no leak)");
    assert_eq!(r1.cells.len(), 3);

    // The table's grid still spans 3 columns regardless of row 0's
    // gridBefore — the grid is a TABLE-level property, the
    // gridBefore is a ROW-level offset.
    assert_eq!(
        t.grid_columns.len(),
        3,
        "table grid unaffected by row gridBefore"
    );
}

#[test]
fn v1_tr_tblpr_ex_cellmar_override_with_start_end_aliases() {
    let Some(doc) = load("v1_tr_tblpr_ex_cellmar_override.docx") else { return };
    let row = &first_table(&doc).rows[0];

    let m = row
        .cell_margins_override
        .as_ref()
        .expect("tblPrEx > tblCellMar must populate cell_margins_override");

    // All four sides distinct values so a top/bottom or left/right
    // swap is caught structurally. start/end aliases at
    // parser/ooxml.rs:5104-5105 route to .left/.right respectively.
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
        "<w:start w:w=\"300\"/> ALIAS → left=15.0pt (NOT a separate .start field)"
    );
    assert!(
        (m.right.unwrap() - 20.0).abs() < 0.001,
        "<w:end w:w=\"400\"/> ALIAS → right=20.0pt"
    );
}

#[test]
fn all_five_fixtures_parse_with_expected_row_count() {
    let cases: &[(&str, usize)] = &[
        ("v1_tr_height_exact.docx", 1),
        ("v1_tr_height_atleast.docx", 1),
        ("v1_tr_header_cant_split.docx", 2),
        ("v1_tr_grid_before.docx", 2),
        ("v1_tr_tblpr_ex_cellmar_override.docx", 1),
    ];
    for (name, expected_rows) in cases {
        let path = fixture_path(name);
        if !path.exists() {
            eprintln!("skipping {}", name);
            continue;
        }
        let data = fs::read(&path).unwrap();
        let doc = parse_docx(&data)
            .unwrap_or_else(|e| panic!("failed to parse {}: {:?}", name, e));
        let t = first_table(&doc);
        assert_eq!(t.rows.len(), *expected_rows, "{} row count", name);
    }
}
