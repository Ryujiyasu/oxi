// This Source Code Form is subject to the terms of the Mozilla Public
// License, v. 2.0. If a copy of the MPL was not distributed with this
// file, You can obtain one at https://mozilla.org/MPL/2.0/.

//! Formula evaluation, delegated to `oxicells-calc`.
//!
//! This module used to carry its own evaluator. It handled arithmetic and eleven
//! functions, had no comparison operators, and evaluated every formula against a
//! snapshot taken before evaluation began — so a formula that referenced another
//! formula read a stale value. Measured against the cached values in the 285
//! workbooks under `tools/golden-test/documents/xlsx`, it agreed with Excel on
//! about 15% of the 29,672 formula cells; `IF(A1>10,…)` alone produced `#VALUE!`
//! in over 23,000 of them.
//!
//! `oxicells-calc` builds a dependency graph and recalculates in topological
//! order. On the same corpus it agrees with Excel on 99.5%.
//!
//! # When to recalculate at all
//!
//! Usually: don't. An `.xlsx` already contains the values Excel computed, and
//! for display those values *are* the right answer — recalculating a freshly
//! loaded file can only introduce divergence. [`crate::parse_xlsx`] therefore
//! keeps them and calls [`fill_missing_formula_values`], which only computes
//! cells the file left without a cached value.
//!
//! Recalculate when the workbook has been *changed*: that is what
//! [`evaluate_workbook_formulas`] is for.

use crate::ir::{CellValue, Sheet, Workbook};
use oxicells_calc::reference::col_to_letters;
use oxicells_calc::{ExcelError, Value};

pub use oxicells_calc::{
    shift_formula_references, translate_formula_references, ReferenceShift, ShiftAxis,
};

/// Recalculate every formula in a single sheet, overwriting cached values.
///
/// Cross-sheet references cannot resolve here and become `#REF!`; use
/// [`evaluate_workbook_formulas`] when the workbook has more than one sheet.
pub fn evaluate_sheet_formulas(sheet: &mut Sheet) {
    recalculate(std::slice::from_mut(sheet), Overwrite::All);
}

/// Recalculate every formula in the workbook, overwriting cached values.
///
/// Use after editing. Cross-sheet references resolve correctly.
pub fn evaluate_workbook_formulas(workbook: &mut Workbook) {
    recalculate(&mut workbook.sheets, Overwrite::All);
}

/// Compute only those formula cells the file left without a cached value,
/// leaving everything Excel already calculated untouched.
pub fn fill_missing_formula_values(workbook: &mut Workbook) {
    recalculate(&mut workbook.sheets, Overwrite::OnlyMissing);
}

#[derive(Clone, Copy, PartialEq, Eq)]
enum Overwrite {
    All,
    OnlyMissing,
}

fn recalculate(sheets: &mut [Sheet], mode: Overwrite) {
    let mut book = oxicells_calc::Workbook::new();

    for sheet in sheets.iter() {
        book.add_sheet(&sheet.name);
        for row in &sheet.rows {
            for cell in &row.cells {
                let addr = a1(cell.col, row.index);
                match formula_text(cell.formula.as_deref()) {
                    // A formula we cannot parse falls back to its cached value,
                    // so one bad cell does not poison everything downstream.
                    Some(text) if book.set_formula(&sheet.name, &addr, text).is_ok() => {}
                    _ => {
                        let _ = book.set_value(&sheet.name, &addr, to_calc(&cell.value));
                    }
                }
            }
        }
    }

    book.recalculate();

    for sheet in sheets.iter_mut() {
        let name = sheet.name.clone();
        for row in &mut sheet.rows {
            for cell in &mut row.cells {
                if formula_text(cell.formula.as_deref()).is_none() {
                    continue;
                }
                if mode == Overwrite::OnlyMissing && !matches!(cell.value, CellValue::Empty) {
                    continue;
                }
                cell.value = from_calc(&book.value(&name, &a1(cell.col, row.index)));
            }
        }
    }
}

/// A shared-formula follower carries an empty `<f>` element; treat that as
/// having no formula rather than as an empty expression.
fn formula_text(formula: Option<&str>) -> Option<&str> {
    formula.filter(|f| !f.trim().is_empty())
}

fn a1(col: u32, row_1based: u32) -> String {
    let mut out = col_to_letters(col);
    out.push_str(&row_1based.to_string());
    out
}

fn to_calc(value: &CellValue) -> Value {
    match value {
        CellValue::Empty => Value::Blank,
        CellValue::String(s) => Value::Text(s.clone()),
        CellValue::Number(n) => Value::Number(*n),
        CellValue::Boolean(b) => Value::Logical(*b),
        CellValue::Error(s) => Value::Error(parse_error_text(s)),
    }
}

fn from_calc(value: &Value) -> CellValue {
    match value {
        Value::Blank => CellValue::Empty,
        Value::Number(n) => CellValue::Number(*n),
        Value::Text(s) => CellValue::String(s.clone()),
        Value::Logical(b) => CellValue::Boolean(*b),
        Value::Error(e) => CellValue::Error(e.as_str().to_string()),
    }
}

fn parse_error_text(s: &str) -> ExcelError {
    match s.to_ascii_uppercase().as_str() {
        "#NULL!" => ExcelError::Null,
        "#DIV/0!" => ExcelError::DivZero,
        "#REF!" => ExcelError::Ref,
        "#NAME?" => ExcelError::Name,
        "#NUM!" => ExcelError::Num,
        "#N/A" => ExcelError::NA,
        _ => ExcelError::Value,
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::ir::{Cell, CellStyle, Row, Sheet};

    fn make_sheet(data: Vec<(u32, u32, CellValue, Option<String>)>) -> Sheet {
        let mut rows_map: std::collections::BTreeMap<u32, Vec<Cell>> =
            std::collections::BTreeMap::new();
        for (r, c, val, formula) in data {
            rows_map.entry(r).or_default().push(Cell {
                col: c,
                value: val,
                style: CellStyle::default(),
                formula,
            });
        }
        let rows: Vec<Row> = rows_map
            .into_iter()
            .map(|(idx, cells)| Row {
                index: idx + 1, // 1-based
                cells,
                height: None,
                hidden: false,
            })
            .collect();
        Sheet {
            name: "Sheet1".to_string(),
            rows,
            col_count: 5,
            col_widths: vec![],
            default_col_width: 8.43,
            default_row_height: 15.0,
            merge_cells: vec![],
            hidden_cols: vec![],
            auto_filter: None,
            unsupported_elements: vec![],
        }
    }

    #[test]
    fn test_simple_arithmetic() {
        let mut sheet = make_sheet(vec![
            (0, 0, CellValue::Number(10.0), None),
            (0, 1, CellValue::Number(20.0), None),
            (0, 2, CellValue::Empty, Some("A1+B1".to_string())),
        ]);
        evaluate_sheet_formulas(&mut sheet);
        assert!(
            matches!(&sheet.rows[0].cells[2].value, CellValue::Number(n) if (*n - 30.0).abs() < f64::EPSILON)
        );
    }

    #[test]
    fn test_sum_function() {
        let mut sheet = make_sheet(vec![
            (0, 0, CellValue::Number(1.0), None),
            (1, 0, CellValue::Number(2.0), None),
            (2, 0, CellValue::Number(3.0), None),
            (3, 0, CellValue::Empty, Some("SUM(A1:A3)".to_string())),
        ]);
        evaluate_sheet_formulas(&mut sheet);
        assert!(
            matches!(&sheet.rows[3].cells[0].value, CellValue::Number(n) if (*n - 6.0).abs() < f64::EPSILON)
        );
    }

    #[test]
    fn test_average_function() {
        let mut sheet = make_sheet(vec![
            (0, 0, CellValue::Number(10.0), None),
            (1, 0, CellValue::Number(20.0), None),
            (2, 0, CellValue::Number(30.0), None),
            (3, 0, CellValue::Empty, Some("AVERAGE(A1:A3)".to_string())),
        ]);
        evaluate_sheet_formulas(&mut sheet);
        assert!(
            matches!(&sheet.rows[3].cells[0].value, CellValue::Number(n) if (*n - 20.0).abs() < f64::EPSILON)
        );
    }

    #[test]
    fn test_min_max() {
        let mut sheet = make_sheet(vec![
            (0, 0, CellValue::Number(5.0), None),
            (1, 0, CellValue::Number(3.0), None),
            (2, 0, CellValue::Number(8.0), None),
            (3, 0, CellValue::Empty, Some("MIN(A1:A3)".to_string())),
            (3, 1, CellValue::Empty, Some("MAX(A1:A3)".to_string())),
        ]);
        evaluate_sheet_formulas(&mut sheet);
        assert!(
            matches!(&sheet.rows[3].cells[0].value, CellValue::Number(n) if (*n - 3.0).abs() < f64::EPSILON)
        );
        assert!(
            matches!(&sheet.rows[3].cells[1].value, CellValue::Number(n) if (*n - 8.0).abs() < f64::EPSILON)
        );
    }

    /// The old evaluator could not parse `A1>5` at all and left `#VALUE!` here.
    /// This is the single largest source of the ~23,000 wrong cells it produced
    /// across the golden-test corpus.
    #[test]
    fn test_if_with_comparison() {
        let mut sheet = make_sheet(vec![
            (0, 0, CellValue::Number(10.0), None),
            (0, 1, CellValue::Empty, Some("IF(A1>5,\"yes\",\"no\")".to_string())),
        ]);
        evaluate_sheet_formulas(&mut sheet);
        assert!(matches!(&sheet.rows[0].cells[1].value, CellValue::String(s) if s == "yes"));
    }

    #[test]
    fn test_division_by_zero() {
        let mut sheet = make_sheet(vec![
            (0, 0, CellValue::Number(10.0), None),
            (0, 1, CellValue::Number(0.0), None),
            (0, 2, CellValue::Empty, Some("A1/B1".to_string())),
        ]);
        evaluate_sheet_formulas(&mut sheet);
        assert!(matches!(&sheet.rows[0].cells[2].value, CellValue::Error(s) if s == "#DIV/0!"));
    }

    #[test]
    fn test_nested_arithmetic() {
        let mut sheet = make_sheet(vec![
            (0, 0, CellValue::Number(2.0), None),
            (0, 1, CellValue::Number(3.0), None),
            (0, 2, CellValue::Number(4.0), None),
            (0, 3, CellValue::Empty, Some("A1*B1+C1".to_string())),
        ]);
        evaluate_sheet_formulas(&mut sheet);
        assert!(
            matches!(&sheet.rows[0].cells[3].value, CellValue::Number(n) if (*n - 10.0).abs() < f64::EPSILON)
        );
    }

    #[test]
    fn test_count_function() {
        let mut sheet = make_sheet(vec![
            (0, 0, CellValue::Number(1.0), None),
            (1, 0, CellValue::String("hello".to_string()), None),
            (2, 0, CellValue::Number(3.0), None),
            (3, 0, CellValue::Empty, Some("COUNT(A1:A3)".to_string())),
        ]);
        evaluate_sheet_formulas(&mut sheet);
        assert!(
            matches!(&sheet.rows[3].cells[0].value, CellValue::Number(n) if (*n - 2.0).abs() < f64::EPSILON)
        );
    }

    /// A formula that reads another formula's result. The old snapshot-based
    /// evaluator returned the stale pre-evaluation value here.
    #[test]
    fn test_formula_chain_resolves_in_order() {
        let mut sheet = make_sheet(vec![
            (0, 0, CellValue::Number(1.0), None),
            (0, 1, CellValue::Empty, Some("A1+1".to_string())),
            (0, 2, CellValue::Empty, Some("B1+1".to_string())),
        ]);
        evaluate_sheet_formulas(&mut sheet);
        assert!(
            matches!(&sheet.rows[0].cells[2].value, CellValue::Number(n) if (*n - 3.0).abs() < f64::EPSILON)
        );
    }

    #[test]
    fn fill_missing_leaves_cached_values_alone() {
        let sheet = make_sheet(vec![
            (0, 0, CellValue::Number(1.0), None),
            // Excel cached 99 here; it disagrees with the formula on purpose.
            (0, 1, CellValue::Number(99.0), Some("A1+1".to_string())),
            (0, 2, CellValue::Empty, Some("A1+5".to_string())),
        ]);
        let mut workbook = Workbook {
            sheets: vec![sheet],
        };
        fill_missing_formula_values(&mut workbook);
        let cells = &workbook.sheets[0].rows[0].cells;
        // Cached value kept.
        assert!(matches!(&cells[1].value, CellValue::Number(n) if (*n - 99.0).abs() < f64::EPSILON));
        // Missing value filled in.
        assert!(matches!(&cells[2].value, CellValue::Number(n) if (*n - 6.0).abs() < f64::EPSILON));
    }

    #[test]
    fn cross_sheet_references_resolve_at_workbook_level() {
        let mut data = make_sheet(vec![(0, 0, CellValue::Number(42.0), None)]);
        data.name = "Data".to_string();
        let main = make_sheet(vec![(0, 0, CellValue::Empty, Some("Data!A1*2".to_string()))]);
        let mut workbook = Workbook {
            sheets: vec![data, main],
        };
        evaluate_workbook_formulas(&mut workbook);
        assert!(
            matches!(&workbook.sheets[1].rows[0].cells[0].value, CellValue::Number(n) if (*n - 84.0).abs() < f64::EPSILON)
        );
    }
}
