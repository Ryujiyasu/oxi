//! Integration tests: parse `<w:tbl>` end-to-end and verify
//! `Block::Table` produces correct row/cell structure including
//! horizontal merge (gridSpan), vertical merge (vMerge), and cell
//! shading.
//!
//! Parser code paths tested:
//! - [parser/ooxml.rs:759](crates/oxidocs-core/src/parser/ooxml.rs#L759):
//!   `<w:tbl>` body element entry.
//! - [parser/ooxml.rs:4572](crates/oxidocs-core/src/parser/ooxml.rs#L4572):
//!   `<w:tr>` row.
//! - [parser/ooxml.rs:5065](crates/oxidocs-core/src/parser/ooxml.rs#L5065):
//!   `<w:tc>` cell.
//! - [parser/ooxml.rs:5278](crates/oxidocs-core/src/parser/ooxml.rs#L5278):
//!   `<w:gridSpan>` → `cell.grid_span`.
//! - [parser/ooxml.rs:5286](crates/oxidocs-core/src/parser/ooxml.rs#L5286):
//!   `<w:vMerge>` → `cell.v_merge` (Some("restart") | Some("continue")).
//! - [parser/ooxml.rs:5295](crates/oxidocs-core/src/parser/ooxml.rs#L5295):
//!   `<w:shd>` → `cell.shading` (hex color string).
//!
//! Fixtures live in `tools/fixtures/table_samples/` and are authored by
//! `tools/metrics/build_table_repro_fixtures.py` (S294).

use std::fs;

use oxidocs_core::ir::{Block, Document, Table, TableCell};
use oxidocs_core::parser::parse_docx;

fn fixture_path(name: &str) -> std::path::PathBuf {
    let crate_root = std::env::current_dir().unwrap();
    let workspace_root = crate_root.parent().unwrap().parent().unwrap();
    workspace_root.join("tools").join("fixtures").join("table_samples").join(name)
}

fn first_table(doc: &Document) -> &Table {
    doc.pages.iter()
        .flat_map(|p| p.blocks.iter())
        .find_map(|b| if let Block::Table(t) = b { Some(t) } else { None })
        .expect("first table")
}

fn cell_text(cell: &TableCell) -> String {
    cell.blocks.iter()
        .filter_map(|b| if let Block::Paragraph(p) = b { Some(p) } else { None })
        .flat_map(|p| p.runs.iter())
        .map(|r| r.text.as_str())
        .collect::<String>()
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
fn v1_simple_2x2_has_4_cells_in_doc_order() {
    let Some(doc) = load("v1_simple_2x2.docx") else { return };
    let t = first_table(&doc);
    assert_eq!(t.rows.len(), 2, "2 rows expected");
    for row in &t.rows {
        assert_eq!(row.cells.len(), 2, "2 cells per row");
    }
    let expected = [["A1", "A2"], ["B1", "B2"]];
    for (ri, row) in t.rows.iter().enumerate() {
        for (ci, cell) in row.cells.iter().enumerate() {
            assert_eq!(cell_text(cell), expected[ri][ci],
                "row[{}] cell[{}]", ri, ci);
            assert_eq!(cell.grid_span, 1, "default grid_span = 1");
            assert!(cell.v_merge.is_none(), "no vertical merge");
            assert!(cell.shading.is_none(), "no shading");
        }
    }
    // grid_columns from <w:tblGrid><w:gridCol w:w="2500"/>×2 → 125pt each
    assert_eq!(t.grid_columns.len(), 2);
    assert!((t.grid_columns[0] - 125.0).abs() < 0.5);
    assert!((t.grid_columns[1] - 125.0).abs() < 0.5);
}

#[test]
fn v1_horizontal_merge_first_row_has_gridspan_2() {
    // Row 0: a single cell with `<w:gridSpan w:val="2"/>` covering both
    // columns. Row 1: two regular cells.
    let Some(doc) = load("v1_horizontal_merge.docx") else { return };
    let t = first_table(&doc);
    assert_eq!(t.rows.len(), 2);

    // Row 0: exactly 1 cell with grid_span=2
    assert_eq!(t.rows[0].cells.len(), 1, "merged row has 1 cell");
    let merged = &t.rows[0].cells[0];
    assert_eq!(merged.grid_span, 2, "h-merge spans 2 columns");
    assert_eq!(cell_text(merged), "Wide header");

    // Row 1: 2 cells, neither merged
    assert_eq!(t.rows[1].cells.len(), 2);
    assert_eq!(t.rows[1].cells[0].grid_span, 1);
    assert_eq!(t.rows[1].cells[1].grid_span, 1);
    assert_eq!(cell_text(&t.rows[1].cells[0]), "Left");
    assert_eq!(cell_text(&t.rows[1].cells[1]), "Right");
}

#[test]
fn v1_vertical_merge_uses_restart_and_continue() {
    // Row 0 col 0: vMerge=restart (starts the merge, carries text)
    // Row 1 col 0: vMerge=continue (continuation, empty text)
    let Some(doc) = load("v1_vertical_merge.docx") else { return };
    let t = first_table(&doc);
    assert_eq!(t.rows.len(), 2);

    // Both rows still have 2 physical cells (vMerge doesn't collapse rows).
    assert_eq!(t.rows[0].cells.len(), 2);
    assert_eq!(t.rows[1].cells.len(), 2);

    // Row 0 col 0: vMerge=restart, carries text "Tall left"
    assert_eq!(t.rows[0].cells[0].v_merge.as_deref(), Some("restart"));
    assert_eq!(cell_text(&t.rows[0].cells[0]), "Tall left");

    // Row 1 col 0: vMerge=continue (no val attribute → defaults to "continue")
    assert_eq!(t.rows[1].cells[0].v_merge.as_deref(), Some("continue"));
    assert_eq!(cell_text(&t.rows[1].cells[0]), "",
        "continuation cell has no rendered text");

    // The right column cells are independent (no vMerge)
    assert!(t.rows[0].cells[1].v_merge.is_none());
    assert!(t.rows[1].cells[1].v_merge.is_none());
    assert_eq!(cell_text(&t.rows[0].cells[1]), "Top right");
    assert_eq!(cell_text(&t.rows[1].cells[1]), "Bottom right");
}

#[test]
fn v1_cell_shading_captures_hex_fill_color() {
    let Some(doc) = load("v1_cell_shading.docx") else { return };
    let t = first_table(&doc);
    assert_eq!(t.rows.len(), 1);
    let row = &t.rows[0];
    assert_eq!(row.cells.len(), 2);

    // First cell: no shading
    assert!(row.cells[0].shading.is_none(), "first cell unshaded");
    assert_eq!(cell_text(&row.cells[0]), "Plain");

    // Second cell: shading="FFFF00" (yellow)
    assert_eq!(row.cells[1].shading.as_deref(), Some("FFFF00"),
        "shading fill captured as hex color");
    assert_eq!(cell_text(&row.cells[1]), "Yellow");
}

#[test]
fn all_four_fixtures_parse_with_expected_table_shape() {
    let cases: &[(&str, usize, usize)] = &[
        // (filename, expected_row_count, expected_first_row_cell_count)
        ("v1_simple_2x2.docx",        2, 2),
        ("v1_horizontal_merge.docx",  2, 1),  // merged row has 1 cell
        ("v1_vertical_merge.docx",    2, 2),
        ("v1_cell_shading.docx",      1, 2),
    ];
    for (name, exp_rows, exp_first_row_cells) in cases {
        let path = fixture_path(name);
        if !path.exists() {
            eprintln!("skipping {}", name);
            continue;
        }
        let data = fs::read(&path).unwrap();
        let doc = parse_docx(&data)
            .unwrap_or_else(|e| panic!("failed to parse {}: {:?}", name, e));
        let t = first_table(&doc);
        assert_eq!(t.rows.len(), *exp_rows, "{} row count", name);
        assert_eq!(t.rows[0].cells.len(), *exp_first_row_cells,
            "{} first row cell count", name);
    }
}
