// This Source Code Form is subject to the terms of the Mozilla Public
// License, v. 2.0. If a copy of the MPL was not distributed with this
// file, You can obtain one at https://mozilla.org/MPL/2.0/.

//! Which cells a sheet's conditional rules catch, and what they put on.
//!
//! Excel keeps these rules on the sheet rather than on the cells, so nothing in
//! a cell says how it will look: the rule has to be run against the value to
//! find out. Over half the rules in a four-hundred-workbook sweep are
//! `expression` rules, which means running a formula once per cell — the book
//! is therefore assembled a single time here and every cell evaluated against
//! it, rather than through `formula::evaluate_expression`, which assembles the
//! whole book on each call and says so.
//!
//! Rows come out counted from zero, the way the rules state their ranges, and
//! not the way `Row::index` counts them.

use std::collections::HashMap;

use crate::ir::{Cell, CellValue, ConditionalRule, DiffStyle, Workbook};

/// A cell a rule caught, and the look the rule puts over whatever the cell
/// already wears.
#[derive(Debug, Clone, PartialEq, serde::Serialize, serde::Deserialize)]
pub struct Caught {
    /// Counted from zero, like the ranges the rules state.
    pub row: u32,
    pub col: u32,
    pub style: DiffStyle,
}

/// The most cells one rule will be run over.
///
/// A rule stated over a whole column covers a million of them, and a sheet with
/// a hundred such rules would otherwise stall the page it is drawn on. The
/// ranges are clipped to what the sheet actually holds first, so this only
/// bites on a sheet that really is that big.
const MOST: usize = 200_000;

/// What each of a sheet's rules catches, lowest priority first, so a later
/// entry for the same cell is the one that wins.
pub fn caught(workbook: &Workbook, sheet: usize) -> Vec<Caught> {
    let Some(held) = workbook.sheets.get(sheet) else {
        return Vec::new();
    };
    if held.conditional_rules.is_empty() {
        return Vec::new();
    }

    // What the sheet actually holds, so a rule stated over a whole column is
    // run over the cells that exist rather than over the column.
    let mut cells: HashMap<(u32, u32), &Cell> = HashMap::new();
    let mut last_row = 0u32;
    let mut last_col = 0u32;
    for line in &held.rows {
        // `Row::index` counts from one; everything here counts from zero.
        let row = line.index.saturating_sub(1);
        for cell in &line.cells {
            last_row = last_row.max(row);
            last_col = last_col.max(cell.col);
            cells.insert((row, cell.col), cell);
        }
    }

    // The engine is only worth building when a rule needs one, and then only
    // once for the whole sheet.
    let book = held
        .conditional_rules
        .iter()
        .any(needs_a_formula)
        .then(|| {
            crate::formula::assemble_for_evaluation(
                &workbook.sheets,
                &workbook.defined_names,
                &workbook.external_books,
            )
        });
    let name = held.name.clone();

    let mut out = Vec::new();
    for rule in &held.conditional_rules {
        let Some(style) = rule.style.clone() else { continue };
        let places = places_of(rule, last_row, last_col);
        if places.is_empty() {
            continue;
        }
        // Two rules count how often a value appears rather than what it is, so
        // the whole range has to be counted before any cell can be judged.
        let tally = matches!(rule.kind.as_str(), "duplicateValues" | "uniqueValues")
            .then(|| {
                let mut seen: HashMap<String, usize> = HashMap::new();
                for &(row, col) in &places {
                    let said = shown(cells.get(&(row, col)).copied());
                    if !said.is_empty() {
                        *seen.entry(said).or_default() += 1;
                    }
                }
                seen
            });

        for (row, col) in places {
            let cell = cells.get(&(row, col)).copied();
            let hit = match rule.kind.as_str() {
                "expression" => book.as_ref().is_some_and(|book| {
                    // The formula is written for the top-left of the rule's
                    // first range and is relative to it, so each cell gets the
                    // same formula with its references moved to suit.
                    let Some((top, left, ..)) = rule.ranges.first() else {
                        return false;
                    };
                    let Some(formula) = rule.formulas.first() else {
                        return false;
                    };
                    let moved = crate::formula::translate_formula_references(
                        formula,
                        row as i64 - *top as i64,
                        col as i64 - *left as i64,
                    );
                    match moved {
                        Ok(moved) => book
                            .evaluate(&name, &moved)
                            .map(|value| truthy(&value))
                            .unwrap_or(false),
                        Err(_) => false,
                    }
                }),
                "cellIs" => cell_is(rule, cell, book.as_ref(), &name),
                "containsText" => contains(rule, cell, true),
                "notContainsText" => contains(rule, cell, false),
                "beginsWith" => {
                    let said = shown(cell).to_lowercase();
                    let wanted = rule.text.clone().unwrap_or_default().to_lowercase();
                    !wanted.is_empty() && said.starts_with(&wanted)
                }
                "endsWith" => {
                    let said = shown(cell).to_lowercase();
                    let wanted = rule.text.clone().unwrap_or_default().to_lowercase();
                    !wanted.is_empty() && said.ends_with(&wanted)
                }
                "containsBlanks" => shown(cell).trim().is_empty(),
                "notContainsBlanks" => !shown(cell).trim().is_empty(),
                "duplicateValues" | "uniqueValues" => {
                    let said = shown(cell);
                    let times = tally
                        .as_ref()
                        .and_then(|seen| seen.get(&said))
                        .copied()
                        .unwrap_or(0);
                    if rule.kind == "duplicateValues" {
                        times > 1
                    } else {
                        times == 1
                    }
                }
                // A rule this does not know is not guessed at. Saying nothing
                // leaves the cell as the file dressed it, which is closer to
                // right than a colour picked on a hunch.
                _ => false,
            };
            if hit {
                out.push(Caught {
                    row,
                    col,
                    style: style.clone(),
                });
            }
        }
    }
    out
}

/// Whether a rule has to be worked out by the engine rather than read off the
/// cell.
fn needs_a_formula(rule: &ConditionalRule) -> bool {
    if rule.kind == "expression" {
        return true;
    }
    // A `cellIs` bound can itself be a formula — `$B$1` or `AVERAGE(A:A)` —
    // and only the engine can say what those come to.
    rule.kind == "cellIs"
        && rule
            .formulas
            .iter()
            .any(|held| held.parse::<f64>().is_err() && !is_quoted(held))
}

/// The cells a rule covers, clipped to what the sheet holds and capped.
fn places_of(rule: &ConditionalRule, last_row: u32, last_col: u32) -> Vec<(u32, u32)> {
    let mut out = Vec::new();
    for &(top, left, bottom, right) in &rule.ranges {
        let bottom = bottom.min(last_row);
        let right = right.min(last_col);
        for row in top..=bottom {
            for col in left..=right {
                if out.len() >= MOST {
                    return out;
                }
                out.push((row, col));
            }
        }
    }
    out
}

fn is_quoted(held: &str) -> bool {
    let held = held.trim();
    held.len() >= 2 && held.starts_with('"') && held.ends_with('"')
}

fn shown(cell: Option<&Cell>) -> String {
    cell.map(|cell| cell.value.display()).unwrap_or_default()
}

fn truthy(value: &oxicells_calc::Value) -> bool {
    match value {
        oxicells_calc::Value::Logical(held) => *held,
        oxicells_calc::Value::Number(held) => *held != 0.0,
        _ => false,
    }
}

fn contains(rule: &ConditionalRule, cell: Option<&Cell>, wanted: bool) -> bool {
    let said = shown(cell).to_lowercase();
    let looking = rule.text.clone().unwrap_or_default().to_lowercase();
    if looking.is_empty() {
        return false;
    }
    said.contains(&looking) == wanted
}

/// A `cellIs` rule: the cell's own value against one or two bounds.
fn cell_is(
    rule: &ConditionalRule,
    cell: Option<&Cell>,
    book: Option<&oxicells_calc::Workbook>,
    sheet: &str,
) -> bool {
    let bound = |at: usize| -> Option<oxicells_calc::Value> {
        let held = rule.formulas.get(at)?.trim();
        if let Ok(number) = held.parse::<f64>() {
            return Some(oxicells_calc::Value::Number(number));
        }
        if is_quoted(held) {
            return Some(oxicells_calc::Value::Text(
                held[1..held.len() - 1].to_string(),
            ));
        }
        book?.evaluate(sheet, held).ok()
    };
    let Some(first) = bound(0) else { return false };
    let mine = match cell.map(|cell| &cell.value) {
        Some(CellValue::Number(held)) => oxicells_calc::Value::Number(*held),
        Some(CellValue::Boolean(held)) => oxicells_calc::Value::Logical(*held),
        Some(CellValue::String(held)) => oxicells_calc::Value::Text(held.clone()),
        // A rule tests what is there; an empty cell is not tested at all,
        // which is what Excel does with every operator but `notEqual`.
        _ => {
            return rule.operator.as_deref() == Some("notEqual");
        }
    };
    let order = compare(&mine, &first);
    match rule.operator.as_deref().unwrap_or("equal") {
        "equal" => order == Some(std::cmp::Ordering::Equal),
        "notEqual" => order != Some(std::cmp::Ordering::Equal),
        "greaterThan" => order == Some(std::cmp::Ordering::Greater),
        "lessThan" => order == Some(std::cmp::Ordering::Less),
        "greaterThanOrEqual" => matches!(
            order,
            Some(std::cmp::Ordering::Greater) | Some(std::cmp::Ordering::Equal)
        ),
        "lessThanOrEqual" => matches!(
            order,
            Some(std::cmp::Ordering::Less) | Some(std::cmp::Ordering::Equal)
        ),
        "between" | "notBetween" => {
            let Some(second) = bound(1) else { return false };
            let low = compare(&mine, &first);
            let high = compare(&mine, &second);
            let inside = matches!(
                (low, high),
                (
                    Some(std::cmp::Ordering::Greater) | Some(std::cmp::Ordering::Equal),
                    Some(std::cmp::Ordering::Less) | Some(std::cmp::Ordering::Equal)
                )
            );
            if rule.operator.as_deref() == Some("between") {
                inside
            } else {
                !inside
            }
        }
        _ => false,
    }
}

/// Two values put in order, where they are the same sort of thing. Text is
/// compared without regard to case, the way Excel compares it.
fn compare(mine: &oxicells_calc::Value, theirs: &oxicells_calc::Value) -> Option<std::cmp::Ordering> {
    use oxicells_calc::Value;
    match (mine, theirs) {
        (Value::Number(a), Value::Number(b)) => a.partial_cmp(b),
        (Value::Logical(a), Value::Logical(b)) => Some(a.cmp(b)),
        (Value::Text(a), Value::Text(b)) => Some(a.to_lowercase().cmp(&b.to_lowercase())),
        // A number written into a text rule, or the other way round: Excel
        // reads the text as a number when it can.
        (Value::Text(a), Value::Number(b)) => a.trim().parse::<f64>().ok()?.partial_cmp(b),
        (Value::Number(a), Value::Text(b)) => a.partial_cmp(&b.trim().parse::<f64>().ok()?),
        _ => None,
    }
}
