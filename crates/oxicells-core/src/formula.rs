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
    format_number, formula_from_r1c1, formula_to_r1c1, move_formula_references,
    shift_formula_references, translate_formula_references, transpose_formula_references, CellMove,
    ReferenceShift, ShiftAxis,
};

/// Recalculate every formula in a single sheet, overwriting cached values.
///
/// Cross-sheet references cannot resolve here and become `#REF!`; use
/// [`evaluate_workbook_formulas`] when the workbook has more than one sheet.
/// How many of a workbook's formulas this engine cannot even read.
///
/// A formula that will not parse keeps the value the file was saved with, so
/// it costs nothing visible and agrees with any oracle by default. Counting
/// them is the only way a measurement can tell "we got this right" from "we
/// never looked".
pub fn unparsed_formulas(workbook: &Workbook) -> usize {
    workbook
        .sheets
        .iter()
        .flat_map(|sheet| sheet.rows.iter())
        .flat_map(|row| row.cells.iter())
        .filter_map(|cell| cell.formula.as_deref())
        .filter(|formula| {
            let text = formula.strip_prefix('=').unwrap_or(formula);
            oxicells_calc::parse(text).is_err()
        })
        .count()
}

pub fn evaluate_sheet_formulas(sheet: &mut Sheet) {
    // One sheet on its own carries no workbook, so it has no names either.
    recalculate(std::slice::from_mut(sheet), &[], Overwrite::All, None);
}

/// Recalculate every formula in the workbook, overwriting cached values.
///
/// Use after editing. Cross-sheet references resolve correctly.
pub fn evaluate_workbook_formulas(workbook: &mut Workbook) {
    let names = workbook.defined_names.clone();
    recalculate(&mut workbook.sheets, &names, Overwrite::All, None);
}

/// The same, told what the moment is.
///
/// `now` is a serial: whole days since the last day of 1899, with the time of
/// day after the point. It is what `TODAY()` and `NOW()` answer.
///
/// Somewhere without a clock has to be told — a browser build has no
/// `SystemTime` — and somewhere with one may still want to pin the moment, so
/// that a sheet worked out twice comes out the same both times.
pub fn evaluate_workbook_formulas_at(workbook: &mut Workbook, now: f64) {
    let names = workbook.defined_names.clone();
    recalculate(&mut workbook.sheets, &names, Overwrite::All, Some(now));
}

/// Work one formula out against the workbook as it stands, without storing it
/// anywhere. `sheet` is where a bare reference like `A1` is read from.
///
/// The whole book has to be handed to the engine before it can look a
/// reference up, so this costs a pass over every cell — the same pass a
/// recalculation makes. It is for the odd expression a macro evaluates, not
/// for a loop over one.
///
/// `None` means the text is not a formula at all. A formula that fails comes
/// back as the error value a cell would show, which is what
/// `Application.Evaluate` hands a macro.
pub fn evaluate_expression(
    workbook: &Workbook,
    sheet: usize,
    formula: &str,
    now: Option<f64>,
) -> Option<oxicells_calc::Value> {
    let name = workbook.sheets.get(sheet)?.name.clone();
    let book = assemble(&workbook.sheets, &workbook.defined_names, now);
    book.evaluate(&name, formula).ok()
}

/// Compute only those formula cells the file left without a cached value,
/// leaving everything Excel already calculated untouched.
pub fn fill_missing_formula_values(workbook: &mut Workbook) {
    let names = workbook.defined_names.clone();
    recalculate(&mut workbook.sheets, &names, Overwrite::OnlyMissing, None);
}

/// Put the workbook to the engine: its names, its sheets, its tables and
/// every cell, either as the formula it holds or as the value it holds.
fn assemble(
    sheets: &[Sheet],
    names: &[(String, String)],
    now: Option<f64>,
) -> oxicells_calc::Workbook {
    let mut book = oxicells_calc::Workbook::new();
    if let Some(moment) = now {
        book.set_now(moment);
    }

    // Named before anything is asked of them: a formula saying `SUM(sales)`
    // means one of these, and a name that will not parse is simply left
    // undefined, which is what it already was.
    for (name, refers_to) in names {
        let _ = book.define_name(name, refers_to);
    }

    for sheet in sheets.iter() {
        book.add_sheet(&sheet.name);
        // A table's own name is how a formula reaches its columns:
        // `tblNomina[[#This Row],[DATE]]`. The IR counts a table's rows from
        // ONE and its columns from zero; the engine counts both from zero, so
        // the rows have to be shifted down on the way across. Handing them
        // over as they stand puts every table one row lower than it is, which
        // reads the headings as data and loses the last row.
        for table in &sheet.tables {
            if table.name.is_empty() {
                continue;
            }
            book.add_table(
                &sheet.name,
                &table.name,
                (
                    table.start_row.saturating_sub(1),
                    table.end_row.saturating_sub(1),
                ),
                (table.start_col, table.end_col),
                table.header_rows,
                table.columns.clone(),
            );
        }
        for row in &sheet.rows {
            for cell in &row.cells {
                let addr = a1(cell.col, row.index);
                match formula_text(cell.formula.as_deref()) {
                    // A formula we cannot parse falls back to its cached value,
                    // so one bad cell does not poison everything downstream.
                    //
                    // Worth knowing when measuring: a formula that will not
                    // parse therefore keeps the answer the file was saved with
                    // and looks, to anything comparing the two, like a perfect
                    // agreement. Improving the parser can LOWER a measured
                    // score by letting formulas through to a gap behind them.
                    // `unparsed_formulas` counts them so that cannot hide.
                    Some(text) if book.set_formula(&sheet.name, &addr, text).is_ok() => {}
                    _ => {
                        let _ = book.set_value(&sheet.name, &addr, to_calc(&cell.value));
                    }
                }
            }
        }
    }

    book
}

#[derive(Clone, Copy, PartialEq, Eq)]
enum Overwrite {
    All,
    OnlyMissing,
}

fn recalculate(
    sheets: &mut [Sheet],
    names: &[(String, String)],
    mode: Overwrite,
    now: Option<f64>,
) {
    // Which formula cells the file left without an answer. A workbook Excel
    // wrote has none, and then there is nothing to do at all — which is worth
    // asking BEFORE building anything, since building it is the cost.
    let missing: Vec<(String, (u32, u32))> = if mode == Overwrite::OnlyMissing {
        let mut found = Vec::new();
        for sheet in sheets.iter() {
            for row in &sheet.rows {
                for cell in &row.cells {
                    if formula_text(cell.formula.as_deref()).is_some()
                        && matches!(cell.value, CellValue::Empty)
                    {
                        found.push((sheet.name.clone(), (cell.col, row.index.saturating_sub(1))));
                    }
                }
            }
        }
        if found.is_empty() {
            return;
        }
        found
    } else {
        Vec::new()
    };

    let mut book = assemble(sheets, names, now);

    match mode {
        Overwrite::All => {
            book.recalculate();
        }
        // Only the cells that need an answer. Whatever they read is either
        // already cached or is another of these, and a missing input comes
        // earlier in the same order.
        Overwrite::OnlyMissing => {
            book.recalculate_these(&missing);
        }
    }

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
                runs: Vec::new(),
            });
        }
        let rows: Vec<Row> = rows_map
            .into_iter()
            .map(|(idx, cells)| Row {
                index: idx + 1, // 1-based
                cells,
                height: None,
                custom_height: false,
                style_font: None,
                thick_top: false,
                thick_bottom: false,
                hidden: false,
            })
            .collect();
        Sheet {
            visibility: Default::default(),
            name: "Sheet1".to_string(),
            rows,
            col_count: 5,
            col_widths: vec![],
            default_col_width: 8.43,
            default_row_height: 15.0,
            default_row_custom: false,
            col_fonts: vec![],
            normal_font: None,
            first_font: None,
            frozen_rows: 0,
            frozen_cols: 0,
            merge_cells: vec![],
            hidden_cols: vec![],
            auto_filter: None,
            declared_range: None,
            tables: Vec::new(),
            drawings: Vec::new(),
            comments: Vec::new(),
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
            ..Default::default()
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
            ..Default::default()
        };
        evaluate_workbook_formulas(&mut workbook);
        assert!(
            matches!(&workbook.sheets[1].rows[0].cells[0].value, CellValue::Number(n) if (*n - 84.0).abs() < f64::EPSILON)
        );
    }

    /// What the formula parser cannot yet read, counted over a corpus and
    /// ranked by how many WORKBOOKS each construct blocks.
    ///
    /// This is a build-order instrument, not a test: it decides what to
    /// implement next by measuring, rather than by guessing which Excel
    /// features matter. Run it with
    ///
    /// ```text
    /// cargo test -p oxicells-core formula_coverage_census -- --ignored --nocapture
    /// ```
    ///
    /// against `pipeline_data/xlsx_corpus` (see the fetch script in
    /// tools/metrics). Rank by books rather than by cells: one workbook full
    /// of the same unsupported formula is one problem, not nine hundred.
    #[test]
    #[ignore = "minutes over the whole corpus; a build-order instrument, asked for by name"]
    fn formula_coverage_census() {
        use std::collections::{BTreeMap, BTreeSet};

        let roots = [
            std::path::Path::new("../../pipeline_data/xlsx_corpus/init"),
            std::path::Path::new("../../pipeline_data/xlsx_corpus/golden"),
        ];
        let entries: Vec<_> = roots
            .iter()
            .filter_map(|root| std::fs::read_dir(root).ok())
            .flat_map(|entries| entries.flatten())
            .collect();
        if entries.is_empty() {
            println!("no corpus found; fetch it with tools/metrics/fetch_xlsx_corpus.py");
            return;
        }

        let mut books = 0usize;
        let mut blocked_books = 0usize;
        let mut formulas = 0usize;
        let mut refused = 0usize;
        let mut reasons: BTreeMap<String, (usize, usize, String)> = BTreeMap::new();

        for entry in entries {
            let path = entry.path();
            if path.extension().and_then(|e| e.to_str()) != Some("xlsx") {
                continue;
            }
            let Ok(bytes) = std::fs::read(&path) else {
                continue;
            };
            let Ok(workbook) = crate::parser::parse_xlsx_preserving_values(&bytes) else {
                continue;
            };
            books += 1;
            let mut here: BTreeSet<String> = BTreeSet::new();
            for sheet in &workbook.sheets {
                for row in &sheet.rows {
                    for cell in &row.cells {
                        let Some(formula) = cell.formula.as_deref() else {
                            continue;
                        };
                        if formula.trim().is_empty() {
                            continue;
                        }
                        formulas += 1;
                        if let Err(error) = oxicells_calc::parse(formula) {
                            refused += 1;
                            let reason = census_reason(&error.to_string(), formula);
                            here.insert(reason.clone());
                            let slot = reasons.entry(reason).or_insert((0, 0, String::new()));
                            slot.0 += 1;
                            if slot.2.is_empty() {
                                slot.2 = formula.chars().take(60).collect();
                            }
                        }
                    }
                }
            }
            if !here.is_empty() {
                blocked_books += 1;
                for reason in here {
                    reasons.entry(reason).or_default().1 += 1;
                }
            }
        }

        println!(
            "\n{books} books, {formulas} formulas, {refused} refused, {blocked_books} books blocked"
        );
        let mut ranked: Vec<_> = reasons.into_iter().collect();
        ranked.sort_by_key(|(_, (cells, books, _))| (usize::MAX - *books, usize::MAX - *cells));
        println!(
            "\n{:<34} {:>7} {:>7}  example",
            "what it could not read", "cells", "books"
        );
        for (reason, (cells, books, example)) in ranked {
            println!("{reason:<34} {cells:>7} {books:>7}  {example}");
        }
    }

    /// Group a refusal by the CONSTRUCT it stumbled on rather than by the
    /// parser's wording: the wording says where it stopped, the construct says
    /// what to build.
    fn census_reason(error: &str, formula: &str) -> String {
        // A bracketed NUMBER is another workbook -- `[1]Assistente!R6:S23` --
        // and it can sit anywhere, not only at the front. Ask that before the
        // structured-reference shape, which is a NAME followed by brackets.
        if error.contains("another workbook")
            || formula
                .match_indices('[')
                .any(|(at, _)| formula[at + 1..].starts_with(|c: char| c.is_ascii_digit()))
        {
            return "another workbook [1]Sheet!A1".to_string();
        }
        if formula.contains('{') {
            return "array constant {..}".to_string();
        }
        if formula.contains('[') && formula.contains(']') {
            return "structured reference Table[..]".to_string();
        }
        if formula.contains('@') {
            return "implicit intersection @".to_string();
        }
        if error.contains("requires a cell reference on both sides") {
            return "a `:` whose side is not a plain reference".to_string();
        }
        if error.contains("trailing input") {
            return "trailing input (often the intersection space)".to_string();
        }
        if error.contains("unexpected character") {
            let shown = error.split('\'').nth(1).unwrap_or("?");
            return format!("unexpected character {shown:?}");
        }
        if error.contains("unexpected end") {
            return "unexpected end".to_string();
        }
        format!("other: {}", error.chars().take(40).collect::<String>())
    }
}
