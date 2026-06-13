// This Source Code Form is subject to the terms of the Mozilla Public
// License, v. 2.0. If a copy of the MPL was not distributed with this
// file, You can obtain one at https://mozilla.org/MPL/2.0/.

//! Integration tests for oxicells-core formula evaluation.
//!
//! S350's public_api_integration covered SUM/AVERAGE/arithmetic/cell-refs.
//! This suite pins the rest of the evaluator's public surface (via
//! `evaluate_sheet_formulas`): MIN/MAX/COUNT/COUNTA/IF, ABS/ROUND,
//! CONCATENATE/LEN, nested functions, operator precedence, division, and
//! error propagation (#NAME? for unknown functions). The evaluator is a
//! self-contained module; pinning its behavior guards the Excel engine's
//! computed-value path.

use oxicells_core::evaluate_sheet_formulas;
use oxicells_core::ir::{Cell, CellStyle, CellValue, Row, Sheet};
use std::collections::BTreeMap;

fn cell(col: u32, value: CellValue, formula: Option<&str>) -> Cell {
    Cell {
        col,
        value,
        style: CellStyle::default(),
        formula: formula.map(|s| s.to_string()),
    }
}

/// Build a sheet from input cells (row_1based, col_0based, value) plus a
/// single formula placed at row 100 col 0 (evaluated against the inputs).
fn eval(formula: &str, inputs: &[(u32, u32, CellValue)]) -> CellValue {
    let mut rows_map: BTreeMap<u32, Vec<Cell>> = BTreeMap::new();
    for (r, c, v) in inputs {
        rows_map.entry(*r).or_default().push(cell(*c, v.clone(), None));
    }
    rows_map
        .entry(100)
        .or_default()
        .push(cell(0, CellValue::Empty, Some(formula)));
    let rows: Vec<Row> = rows_map
        .into_iter()
        .map(|(index, cells)| Row { index, cells, height: None })
        .collect();
    let mut sheet = Sheet {
        name: "S".into(),
        rows,
        col_count: 10,
        col_widths: vec![],
        default_col_width: 8.43,
        default_row_height: 15.0,
        merge_cells: vec![],
        unsupported_elements: vec![],
    };
    evaluate_sheet_formulas(&mut sheet);
    let row = sheet.rows.iter().find(|r| r.index == 100).unwrap();
    row.cells.iter().find(|c| c.col == 0).unwrap().value.clone()
}

fn num(v: CellValue) -> f64 {
    match v {
        CellValue::Number(n) => n,
        other => panic!("expected Number, got {other:?}"),
    }
}

const COL_A: u32 = 0;

fn col_a(vals: &[f64]) -> Vec<(u32, u32, CellValue)> {
    vals.iter()
        .enumerate()
        .map(|(i, v)| ((i + 1) as u32, COL_A, CellValue::Number(*v)))
        .collect()
}

// ── range aggregate functions ──────────────────────────────────────

#[test]
fn min_over_range() {
    assert_eq!(num(eval("=MIN(A1:A4)", &col_a(&[5.0, 2.0, 9.0, 3.0]))), 2.0);
}

#[test]
fn max_over_range() {
    assert_eq!(num(eval("=MAX(A1:A4)", &col_a(&[5.0, 2.0, 9.0, 3.0]))), 9.0);
}

#[test]
fn count_counts_numeric_cells() {
    // COUNT counts numeric cells in the range.
    assert_eq!(num(eval("=COUNT(A1:A4)", &col_a(&[5.0, 2.0, 9.0, 3.0]))), 4.0);
}

#[test]
fn sum_and_average_consistency() {
    let inputs = col_a(&[10.0, 20.0, 30.0]);
    assert_eq!(num(eval("=SUM(A1:A3)", &inputs)), 60.0);
    assert_eq!(num(eval("=AVERAGE(A1:A3)", &inputs)), 20.0);
}

// ── IF ─────────────────────────────────────────────────────────────

#[test]
fn if_true_branch() {
    // IF(cond, then, else): numeric cond != 0 is true.
    let v = eval("=IF(1, 100, 200)", &[]);
    assert_eq!(num(v), 100.0);
}

#[test]
fn if_false_branch() {
    let v = eval("=IF(0, 100, 200)", &[]);
    assert_eq!(num(v), 200.0);
}

// ── scalar functions ───────────────────────────────────────────────

#[test]
fn abs_negative() {
    assert_eq!(num(eval("=ABS(0-7)", &[])), 7.0);
}

#[test]
fn round_to_places() {
    // ROUND(x, n) — use a non-PI-like value to avoid clippy::approx_constant.
    let v = eval("=ROUND(5.678, 1)", &[]);
    assert!((num(v) - 5.7).abs() < 1e-9, "ROUND(5.678,1)=5.7");
}

#[test]
fn len_of_string_literal() {
    let v = eval("=LEN(\"hello\")", &[]);
    assert_eq!(num(v), 5.0);
}

#[test]
fn concatenate_strings() {
    let v = eval("=CONCATENATE(\"ab\", \"cd\")", &[]);
    match v {
        CellValue::String(s) => assert_eq!(s, "abcd"),
        other => panic!("expected String, got {other:?}"),
    }
}

// ── arithmetic / precedence / division ─────────────────────────────

#[test]
fn precedence_mul_before_add() {
    assert_eq!(num(eval("=2+3*4", &[])), 14.0);
}

#[test]
fn parentheses_override_precedence() {
    assert_eq!(num(eval("=(2+3)*4", &[])), 20.0);
}

#[test]
fn division() {
    assert_eq!(num(eval("=10/4", &[])), 2.5);
}

// ── nested functions ───────────────────────────────────────────────

#[test]
fn nested_sum_in_arithmetic() {
    let v = eval("=SUM(A1:A3)*2", &col_a(&[1.0, 2.0, 3.0]));
    assert_eq!(num(v), 12.0);
}

#[test]
fn nested_function_in_function() {
    // MAX of a SUM and a literal.
    let v = eval("=MAX(SUM(A1:A2), 100)", &col_a(&[10.0, 20.0]));
    assert_eq!(num(v), 100.0);
}

// ── error propagation ──────────────────────────────────────────────

#[test]
fn unknown_function_yields_name_error() {
    let v = eval("=BOGUSFN(1,2)", &[]);
    match v {
        CellValue::Error(e) => assert!(e.contains("#NAME?"), "expected #NAME?, got {e:?}"),
        other => panic!("expected Error, got {other:?}"),
    }
}

#[test]
fn cell_reference_resolves() {
    // =A1 + A2 with A1=7, A2=8 → 15
    let v = eval("=A1+A2", &col_a(&[7.0, 8.0]));
    assert_eq!(num(v), 15.0);
}
