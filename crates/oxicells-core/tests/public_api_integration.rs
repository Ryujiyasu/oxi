// This Source Code Form is subject to the terms of the Mozilla Public
// License, v. 2.0. If a copy of the MPL was not distributed with this
// file, You can obtain one at https://mozilla.org/MPL/2.0/.

//! Integration tests for oxicells-core public API.
//!
//! Pins `parse_cell_ref`, `col_to_letter`, `parse_xlsx`, and
//! `evaluate_sheet_formulas` behavior so silent regressions in the Excel
//! engine surface in CI. This is the first integration suite for the
//! oxicells-core crate (zero-test before S350).
//!
//! Why these specific cases:
//!
//! - `parse_cell_ref` and `col_to_letter` are inverses across the full
//!   Excel column range (A..XFD = 0..16383). The round-trip property
//!   tests below pin that invariant — a regression that breaks
//!   round-trip at any single column would corrupt every cell address
//!   touched downstream.
//!
//! - `parse_cell_ref` accepts lowercase input (silently uppercases via
//!   `to_ascii_uppercase`) and tolerates missing row digits (returns
//!   row=0). These are NON-OBVIOUS behaviors callers may depend on;
//!   they are pinned here.
//!
//! - `parse_xlsx` against an empty (no sheets) workbook surfaces parser
//!   error handling — not covered by basic_test.xlsx (which has 1 sheet)
//!   nor multi_sheet.xlsx.
//!
//! - `evaluate_sheet_formulas` is called only via internal paths in lib
//!   tests; the public-API integration covers it end-to-end.

use oxicells_core::editor::col_to_letter;
use oxicells_core::parser::{parse_cell_ref, parse_xlsx};
use oxicells_core::{evaluate_sheet_formulas, XlsxEditor};
use oxicells_core::ir::{Cell, CellStyle, CellValue, Row, Sheet};

// ────────────────────────────────────────────────────────────────────
// parse_cell_ref + col_to_letter (public API)
// ────────────────────────────────────────────────────────────────────

#[test]
fn parse_cell_ref_basic() {
    assert_eq!(parse_cell_ref("A1"), (0, 0));
    assert_eq!(parse_cell_ref("Z1"), (25, 0));
    assert_eq!(parse_cell_ref("AA1"), (26, 0));
    assert_eq!(parse_cell_ref("AZ1"), (51, 0));
    assert_eq!(parse_cell_ref("BA1"), (52, 0));
    assert_eq!(parse_cell_ref("ZZ1"), (701, 0));
    assert_eq!(parse_cell_ref("AAA1"), (702, 0));
}

#[test]
fn parse_cell_ref_excel_max_column_xfd() {
    // Excel's hard-coded maximum column = 16384 columns = XFD (col index 16383).
    assert_eq!(parse_cell_ref("XFD1"), (16383, 0));
    assert_eq!(parse_cell_ref("XFD1048576"), (16383, 1048575));
}

#[test]
fn parse_cell_ref_lowercase_accepted() {
    // `to_ascii_uppercase` inside parse_cell_ref means lowercase letters
    // are treated as their uppercase equivalents. Pin this — callers
    // building references from formula strings rely on it.
    assert_eq!(parse_cell_ref("a1"), (0, 0));
    assert_eq!(parse_cell_ref("aa1"), (26, 0));
    assert_eq!(parse_cell_ref("Ab2"), (27, 1));
}

#[test]
fn parse_cell_ref_no_row_digit_defaults_zero() {
    // When the input has no digits, row_str.parse::<u32>() fails →
    // unwrap_or(1).saturating_sub(1) = 0. This is the documented
    // graceful-degrade for the rare case of a column-only reference.
    assert_eq!(parse_cell_ref("A"), (0, 0));
    assert_eq!(parse_cell_ref("AB"), (27, 0));
    assert_eq!(parse_cell_ref("Z"), (25, 0));
}

#[test]
fn parse_cell_ref_row_only_collapses_to_a() {
    // Without any letter prefix, col stays 0 → unwrap_or returns col=0
    // (NOT a panic). Test covers the digit-only edge case.
    assert_eq!(parse_cell_ref("1"), (0, 0));
    assert_eq!(parse_cell_ref("42"), (0, 41));
}

#[test]
fn col_to_letter_basic() {
    assert_eq!(col_to_letter(0), "A");
    assert_eq!(col_to_letter(1), "B");
    assert_eq!(col_to_letter(25), "Z");
    assert_eq!(col_to_letter(26), "AA");
    assert_eq!(col_to_letter(27), "AB");
    assert_eq!(col_to_letter(51), "AZ");
    assert_eq!(col_to_letter(52), "BA");
    assert_eq!(col_to_letter(701), "ZZ");
    assert_eq!(col_to_letter(702), "AAA");
}

#[test]
fn col_to_letter_excel_max() {
    // Excel's max column count = 16384, indexed 0..16383.
    assert_eq!(col_to_letter(16383), "XFD");
}

#[test]
fn parse_cell_ref_col_to_letter_roundtrip_full_range() {
    // Critical invariant: col_to_letter ∘ parse_cell_ref = id for the
    // full Excel column range. A regression in either function that
    // breaks the round-trip would corrupt every address read or written.
    for col in 0u32..16384 {
        let letter = col_to_letter(col);
        let parsed = parse_cell_ref(&format!("{}1", letter));
        assert_eq!(
            parsed.0, col,
            "round-trip failed at col {}: letter={}, parsed_col={}",
            col, letter, parsed.0
        );
    }
}

// ────────────────────────────────────────────────────────────────────
// parse_xlsx public entry point (smoke + error path)
// ────────────────────────────────────────────────────────────────────

#[test]
fn parse_xlsx_basic_fixture() {
    let data = include_bytes!("../../../tests/fixtures/basic_test.xlsx");
    let wb = parse_xlsx(data).expect("parse_xlsx must succeed on basic_test.xlsx");
    assert_eq!(wb.sheets.len(), 1, "basic_test has 1 sheet");
    assert!(!wb.sheets[0].name.is_empty(), "sheet name not empty");
    assert!(!wb.sheets[0].rows.is_empty(), "rows not empty");
}

#[test]
fn parse_xlsx_invalid_data_returns_err() {
    // Garbage input must produce Err (not panic). This is the parser's
    // contract — corrupt files in user uploads must not crash the WASM
    // host or CLI.
    let result = parse_xlsx(b"not a zip file");
    assert!(result.is_err(), "garbage input must error, not panic");
}

#[test]
fn parse_xlsx_empty_bytes_returns_err() {
    let result = parse_xlsx(b"");
    assert!(result.is_err());
}

// ────────────────────────────────────────────────────────────────────
// evaluate_sheet_formulas (public API)
// ────────────────────────────────────────────────────────────────────

fn make_cell(col: u32, value: CellValue, formula: Option<&str>) -> Cell {
    Cell {
        col,
        value,
        style: CellStyle::default(),
        formula: formula.map(|s| s.to_string()),
    }
}

fn make_sheet_with_formula(formula: &str, input_values: &[(u32, u32, CellValue)]) -> Sheet {
    // Build a sheet with input cells (row, col, value) and ONE formula cell
    // at the next free position (placed at the END so its formula evaluates
    // against the input snapshot).
    use std::collections::BTreeMap;
    let mut rows_map: BTreeMap<u32, Vec<Cell>> = BTreeMap::new();
    for (r, c, v) in input_values {
        rows_map.entry(*r).or_default().push(make_cell(*c, v.clone(), None));
    }
    // formula cell at row 100 col 0
    rows_map
        .entry(100)
        .or_default()
        .push(make_cell(0, CellValue::Empty, Some(formula)));
    let rows: Vec<Row> = rows_map
        .into_iter()
        .map(|(index, cells)| Row {
            index,
            cells,
            height: None,
        })
        .collect();
    Sheet {
        name: "TestSheet".to_string(),
        rows,
        col_count: 10,
        col_widths: vec![],
        default_col_width: 8.43,
        default_row_height: 15.0,
        merge_cells: vec![],
        unsupported_elements: vec![],
    }
}

fn formula_result(sheet: &Sheet) -> &CellValue {
    // Find the formula cell (row 100 col 0)
    let row = sheet.rows.iter().find(|r| r.index == 100).expect("formula row");
    let cell = row.cells.iter().find(|c| c.col == 0).expect("formula cell");
    &cell.value
}

#[test]
fn evaluate_sum_function() {
    let mut sheet = make_sheet_with_formula(
        "=SUM(A1:A3)",
        &[
            (1, 0, CellValue::Number(10.0)),
            (2, 0, CellValue::Number(20.0)),
            (3, 0, CellValue::Number(30.0)),
        ],
    );
    evaluate_sheet_formulas(&mut sheet);
    match formula_result(&sheet) {
        CellValue::Number(n) => assert!((*n - 60.0).abs() < f64::EPSILON, "SUM = 60, got {}", n),
        other => panic!("expected Number, got {:?}", other),
    }
}

#[test]
fn evaluate_arithmetic_expression() {
    // =2 + 3 * 4 = 14 (operator precedence)
    let mut sheet = make_sheet_with_formula("=2 + 3 * 4", &[]);
    evaluate_sheet_formulas(&mut sheet);
    match formula_result(&sheet) {
        CellValue::Number(n) => assert!((*n - 14.0).abs() < f64::EPSILON, "= 14, got {}", n),
        other => panic!("expected Number, got {:?}", other),
    }
}

#[test]
fn evaluate_cell_ref() {
    let mut sheet = make_sheet_with_formula(
        "=A1",
        &[(1, 0, CellValue::Number(42.0))],
    );
    evaluate_sheet_formulas(&mut sheet);
    match formula_result(&sheet) {
        CellValue::Number(n) => assert!((*n - 42.0).abs() < f64::EPSILON),
        other => panic!("expected Number(42), got {:?}", other),
    }
}

#[test]
fn evaluate_average_function() {
    let mut sheet = make_sheet_with_formula(
        "=AVERAGE(A1:A4)",
        &[
            (1, 0, CellValue::Number(10.0)),
            (2, 0, CellValue::Number(20.0)),
            (3, 0, CellValue::Number(30.0)),
            (4, 0, CellValue::Number(40.0)),
        ],
    );
    evaluate_sheet_formulas(&mut sheet);
    match formula_result(&sheet) {
        CellValue::Number(n) => assert!((*n - 25.0).abs() < f64::EPSILON, "AVG = 25, got {}", n),
        other => panic!("expected Number, got {:?}", other),
    }
}

// ────────────────────────────────────────────────────────────────────
// XlsxEditor public API (smoke)
// ────────────────────────────────────────────────────────────────────

#[test]
fn xlsx_editor_set_cell_and_workbook_accessor() {
    let data = include_bytes!("../../../tests/fixtures/basic_test.xlsx");
    let mut editor = XlsxEditor::new(data).expect("editor must construct");
    // Read-only workbook accessor returns the parsed IR
    let wb = editor.workbook();
    assert_eq!(wb.sheets.len(), 1);

    // set_cell records an edit in memory (does not mutate workbook directly)
    editor.set_cell(0, 1, 0, "newvalue".to_string());
    // Workbook accessor still returns original (set_cell defers actual mutation)
    let wb_after = editor.workbook();
    // Sheet count unchanged
    assert_eq!(wb_after.sheets.len(), 1);
}
