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

/// The most cells one rule will be run over when the cell alone answers it.
///
/// A rule stated over a whole column covers a million of them, and a sheet with
/// a hundred such rules would otherwise stall the page it is drawn on. This
/// kind of rule is a string comparison, so the cap can be generous.
const MOST: usize = 200_000;

/// The same, for a rule the formula engine has to answer.
///
/// Each of those cells is a parse and an evaluation rather than a comparison,
/// so the same number of them costs a great deal more.
const DEARER: usize = 20_000;

/// What a sheet's rules put on each cell they catch, one entry per cell.
///
/// Excel does not pick a winning rule. It lays the matching rules over one
/// another in precedence order — the lowest `priority` number first — and each
/// rule contributes only the parts the rules above it left unset, so a rule
/// that sets weight alone and a rule that sets fill alone both show. A rule
/// marked `stopIfTrue` ends the stack for the cell it catches.
pub fn caught(workbook: &Workbook, sheet: usize) -> Vec<Caught> {
    let Some(held) = workbook.sheets.get(sheet) else {
        return Vec::new();
    };
    if held.conditional_rules.is_empty() {
        return Vec::new();
    }

    // What the sheet holds, to look a cell up by where it is. A place a rule
    // covers and the file does not record is a blank cell, which several kinds
    // of rule have an answer for.
    let mut cells: HashMap<(u32, u32), &Cell> = HashMap::new();
    for line in &held.rows {
        // `Row::index` counts from one; everything here counts from zero.
        let row = line.index.saturating_sub(1);
        for cell in &line.cells {
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

    // What each caught cell wears so far, and the cells a `stopIfTrue` rule has
    // already finished with.
    let mut worn: HashMap<(u32, u32), DiffStyle> = HashMap::new();
    let mut finished: std::collections::HashSet<(u32, u32)> = std::collections::HashSet::new();
    let mut order: Vec<&ConditionalRule> = held.conditional_rules.iter().collect();
    order.sort_by_key(|rule| rule.priority);

    for rule in order {
        let Some(style) = rule.style.clone() else { continue };
        let places = places_of(rule, needs_a_formula(rule));
        if places.is_empty() {
            continue;
        }
        // A rank rule reads the whole range before it can judge one cell: the
        // cut is the Nth number down from the top (or up from the bottom), and
        // every cell at or past it is caught, ties and all.
        let cut = (rule.kind == "top10")
            .then(|| cut_of(rule, &places, &cells))
            .flatten();

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
            // A rule above this one caught the cell and said to stop.
            if finished.contains(&(row, col)) {
                continue;
            }
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
                        // Read AT the cell, not merely about it: `ROW()` and
                        // `COLUMN()` with no argument answer with wherever the
                        // formula is standing, and `MOD(ROW(),2)=0` — banded
                        // rows — is the commonest conditional rule written.
                        Ok(moved) => book
                            .evaluate_at(&name, &moved, (col, row))
                            .map(|value| truthy(&value))
                            .unwrap_or(false),
                        Err(_) => false,
                    }
                }),
                "cellIs" => cell_is(rule, cell, book.as_ref(), &name, (col, row)),
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
                "top10" => cut.is_some_and(|edge| {
                    let Some(mine) = number_of(cell) else { return false };
                    if rule.bottom {
                        mine <= edge
                    } else {
                        mine >= edge
                    }
                }),
                // A rule this does not know is not guessed at. Saying nothing
                // leaves the cell as the file dressed it, which is closer to
                // right than a colour picked on a hunch.
                _ => false,
            };
            if hit {
                layer(worn.entry((row, col)).or_default(), &style);
                if rule.stop_if_true {
                    finished.insert((row, col));
                }
            }
        }
    }

    let mut out: Vec<Caught> = worn
        .into_iter()
        .map(|((row, col), style)| Caught { row, col, style })
        .collect();
    out.sort_unstable_by_key(|hit| (hit.row, hit.col));
    out
}

/// Add what a rule sets to what the rules above it already set, and nothing
/// more: the one that got there first keeps each part.
fn layer(worn: &mut DiffStyle, adding: &DiffStyle) {
    if worn.font_color.is_none() {
        worn.font_color.clone_from(&adding.font_color);
    }
    if worn.bg_color.is_none() {
        worn.bg_color.clone_from(&adding.bg_color);
    }
    if worn.number_format.is_none() {
        worn.number_format.clone_from(&adding.number_format);
    }
    worn.bold = worn.bold.or(adding.bold);
    worn.italic = worn.italic.or(adding.italic);
    worn.underline = worn.underline.or(adding.underline);
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

/// The cells a rule covers, capped by what answering it costs.
///
/// Not clipped to the cells the file records. Excel runs a rule over every cell
/// of its stated range whether anything was ever typed there or not, and the
/// commonest proof of it is a `containsBlanks` rule, which colours precisely the
/// cells that hold nothing. Clipping to the written extent lost the last cell of
/// `tests/fixtures/conditional/blanks.xlsx`, which Excel colours.
fn places_of(rule: &ConditionalRule, dearer: bool) -> Vec<(u32, u32)> {
    let most = if dearer { DEARER } else { MOST };
    let mut out = Vec::new();
    for &(top, left, bottom, right) in &rule.ranges {
        for row in top..=bottom {
            for col in left..=right {
                if out.len() >= most {
                    return out;
                }
                out.push((row, col));
            }
        }
    }
    out
}

/// The number a cell holds, if it holds one. Text that looks like a number is
/// not one — Excel's rank rules ignore it, as they ignore blanks.
fn number_of(cell: Option<&Cell>) -> Option<f64> {
    match cell.map(|cell| &cell.value) {
        Some(CellValue::Number(held)) => Some(*held),
        _ => None,
    }
}

/// The value at the edge of a `top10` rule: the Nth from the top of the range,
/// or from the bottom when the rule says so. `None` when nothing in the range
/// is a number, which is a rule that catches nothing.
fn cut_of(
    rule: &ConditionalRule,
    places: &[(u32, u32)],
    cells: &HashMap<(u32, u32), &Cell>,
) -> Option<f64> {
    let mut numbers: Vec<f64> = places
        .iter()
        .filter_map(|at| number_of(cells.get(at).copied()))
        .collect();
    if numbers.is_empty() {
        return None;
    }
    // Excel writes no `rank` for the "Top 10 Items" preset it names itself.
    let asked = rule.rank.unwrap_or(10).max(1) as usize;
    let take = if rule.percent {
        // A share of the cells holding a number, and never none of them.
        (numbers.len() * asked.min(100) / 100).max(1)
    } else {
        asked
    }
    .min(numbers.len());
    if rule.bottom {
        numbers.sort_by(|a, b| a.partial_cmp(b).unwrap_or(std::cmp::Ordering::Equal));
    } else {
        numbers.sort_by(|a, b| b.partial_cmp(a).unwrap_or(std::cmp::Ordering::Equal));
    }
    numbers.get(take - 1).copied()
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
///
/// `place` is the cell being judged, `(column, row)` from zero. A bound that is
/// a formula is written for the top-left of the rule's first range, like an
/// `expression` rule's, so it is moved to this cell before being read — a rule
/// stated as "greater than `$B2`" means a different row in every row.
fn cell_is(
    rule: &ConditionalRule,
    cell: Option<&Cell>,
    book: Option<&oxicells_calc::Workbook>,
    sheet: &str,
    place: (u32, u32),
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
        let book = book?;
        let (top, left, ..) = rule.ranges.first()?;
        let moved = crate::formula::translate_formula_references(
            held,
            place.1 as i64 - *top as i64,
            place.0 as i64 - *left as i64,
        )
        .unwrap_or_else(|_| held.to_string());
        book.evaluate_at(sheet, &moved, place).ok()
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

/// Two values put in order the way Excel puts them.
///
/// Within a kind this is what anyone would expect, text without regard to case.
/// Across kinds Excel does NOT read a number out of the text: every number
/// comes before every piece of text, and every piece of text before either
/// logical value. So a cell holding the TEXT "50" is greater than 50, and a
/// `greaterThan 50` rule colours it — measured against Excel's `DisplayFormat`
/// on `tests/fixtures/conditional/cell_is.xlsx`, where reading "50" as fifty
/// left that cell uncoloured and Excel coloured it.
fn compare(mine: &oxicells_calc::Value, theirs: &oxicells_calc::Value) -> Option<std::cmp::Ordering> {
    use oxicells_calc::Value;
    /// Numbers, then text, then logicals.
    fn rank(value: &Value) -> Option<u8> {
        match value {
            Value::Number(_) => Some(0),
            Value::Text(_) => Some(1),
            Value::Logical(_) => Some(2),
            _ => None,
        }
    }
    match (mine, theirs) {
        (Value::Number(a), Value::Number(b)) => a.partial_cmp(b),
        (Value::Logical(a), Value::Logical(b)) => Some(a.cmp(b)),
        (Value::Text(a), Value::Text(b)) => Some(a.to_lowercase().cmp(&b.to_lowercase())),
        _ => Some(rank(mine)?.cmp(&rank(theirs)?)),
    }
}
