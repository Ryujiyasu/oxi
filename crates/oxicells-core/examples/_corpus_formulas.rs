// This Source Code Form is subject to the terms of the Mozilla Public
// License, v. 2.0. If a copy of the MPL was not distributed with this
// file, You can obtain one at https://mozilla.org/MPL/2.0/.

//! Does the evaluator get the same answers the file was saved with?
//!
//! Every formula cell in a workbook arrives with the answer Excel worked out
//! when it last saved. Recomputing them and comparing is therefore a test with
//! a real mark scheme, and it needs no Excel to run.
//!
//! The claim on the corpus everything was derived on is 99.5% of 29,672 cells.
//! That corpus cannot say whether the figure generalises, since it is where the
//! evaluator was taught. This asks a bundle nothing was fitted to.
//!
//! Prints one line per workbook: how many formulas, how many agreed, and the
//! first few that did not, with the function they lead with — a name that
//! recurs across the disagreements is a gap rather than an accident.

use oxicells_core::ir::CellValue;

fn main() {
    let mut agreed = 0usize;
    let mut asked = 0usize;
    let mut blank = 0usize;
    for path in std::env::args().skip(1) {
        let data = match std::fs::read(&path) {
            Ok(data) => data,
            Err(_) => continue,
        };
        let Ok(book) = oxicells_core::parse_xlsx(&data) else {
            continue;
        };
        // What the file was saved holding, keyed by where the cell IS. Pairing
        // the two lists by position would be a wrong answer waiting to happen:
        // any cell added or dropped in between shifts every one after it, and
        // the disagreements would all be someone else's.
        let mut held: std::collections::BTreeMap<(usize, u32, u32), (String, CellValue)> =
            std::collections::BTreeMap::new();
        for (which, sheet) in book.sheets.iter().enumerate() {
            for row in &sheet.rows {
                for cell in &row.cells {
                    if let Some(formula) = &cell.formula {
                        held.insert(
                            (which, row.index, cell.col),
                            (formula.clone(), cell.value.clone()),
                        );
                    }
                }
            }
        }
        if held.is_empty() {
            continue;
        }
        // And what it comes to when worked out afresh.
        let mut again = book.clone();
        oxicells_core::formula::evaluate_workbook_formulas(&mut again);
        let mut after: std::collections::BTreeMap<(usize, u32, u32), CellValue> =
            std::collections::BTreeMap::new();
        for (which, sheet) in again.sheets.iter().enumerate() {
            for row in &sheet.rows {
                for cell in &row.cells {
                    if cell.formula.is_some() {
                        after.insert((which, row.index, cell.col), cell.value.clone());
                    }
                }
            }
        }
        for (at, (formula, was)) in &held {
            let Some(now) = after.get(at).cloned() else { continue };
            let was = Some(was.clone());
            let now = Some(now);
            asked += 1;
            if same(was.as_ref(), now.as_ref()) {
                agreed += 1;
            } else {
                if matches!(now, Some(CellValue::Empty) | None) {
                    blank += 1;
                }
                println!(
                    "{}\t{}\t{:?}\t{:?}\t{}",
                    leads(formula),
                    path,
                    brief(was.as_ref()),
                    brief(now.as_ref()),
                    formula.replace('\t', " "),
                );
            }
        }
    }
    eprintln!(
        "  {asked} formulas, {agreed} agreed ({:.2}%), {} came back empty",
        if asked == 0 { 0.0 } else { 100.0 * agreed as f64 / asked as f64 },
        blank,
    );
}

/// The function a formula leads with, which is what groups its failures.
fn leads(formula: &str) -> String {
    let body = formula.trim_start_matches(['=', '+', '-', ' ']);
    let name: String = body
        .chars()
        .take_while(|one| one.is_ascii_alphanumeric() || *one == '.' || *one == '_')
        .collect();
    if name.is_empty() {
        "(no function)".to_string()
    } else {
        name.to_ascii_uppercase()
    }
}

fn brief(value: Option<&CellValue>) -> String {
    match value {
        Some(CellValue::Number(n)) => format!("{n}"),
        Some(CellValue::String(s)) => s.chars().take(24).collect(),
        Some(CellValue::Boolean(b)) => format!("{b}"),
        Some(CellValue::Error(e)) => e.clone(),
        Some(CellValue::Empty) | None => "(empty)".to_string(),
    }
}

/// Two answers agree when they are the same kind and the same value; numbers
/// to a part in a billion, since a cached answer was rounded on its way into
/// the file and reading it back exactly is not the question being asked.
fn same(was: Option<&CellValue>, now: Option<&CellValue>) -> bool {
    match (was, now) {
        (Some(CellValue::Number(a)), Some(CellValue::Number(b))) => {
            let scale = a.abs().max(b.abs()).max(1.0);
            (a - b).abs() / scale < 1e-9
        }
        (Some(CellValue::String(a)), Some(CellValue::String(b))) => a == b,
        (Some(CellValue::Boolean(a)), Some(CellValue::Boolean(b))) => a == b,
        (Some(CellValue::Error(a)), Some(CellValue::Error(b))) => a == b,
        (Some(CellValue::Empty) | None, Some(CellValue::Empty) | None) => true,
        _ => false,
    }
}
