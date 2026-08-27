// This Source Code Form is subject to the terms of the Mozilla Public
// License, v. 2.0. If a copy of the MPL was not distributed with this
// file, You can obtain one at https://mozilla.org/MPL/2.0/.

//! Which formulas this engine cannot read at all, and what it says about them.
//!
//! A formula that will not parse keeps the value the file was saved with, so
//! it costs nothing visible and agrees with any oracle by default. That makes
//! it invisible to the agreement measurement and, worse, makes the parser look
//! better the less of the language it understands.
//!
//! Prints one line per unreadable formula: the complaint, the workbook, and
//! the formula itself, so a complaint that recurs can be told from one that
//! does not.

fn main() {
    let mut unread = 0usize;
    let mut total = 0usize;
    for path in std::env::args().skip(1) {
        let Ok(data) = std::fs::read(&path) else {
            continue;
        };
        let Ok(book) = oxicells_core::parse_xlsx(&data) else {
            continue;
        };
        for sheet in &book.sheets {
            for row in &sheet.rows {
                for cell in &row.cells {
                    let Some(formula) = cell.formula.as_deref() else {
                        continue;
                    };
                    total += 1;
                    let text = formula.strip_prefix('=').unwrap_or(formula);
                    if let Err(why) = oxicells_calc::parse(text) {
                        unread += 1;
                        println!(
                            "{}\t{}\t{}",
                            complaint(&why),
                            path,
                            formula.replace('\t', " "),
                        );
                    }
                }
            }
        }
    }
    eprintln!("  {total} formulas, {unread} of them unreadable");
}

/// The shape of a complaint, with the particulars taken out, so that the same
/// gap reported about twenty different cells counts as one gap.
fn complaint(why: &oxicells_calc::ParseError) -> String {
    let said = format!("{why:?}");
    // `UnexpectedToken("...")` and friends carry the offending text, which is
    // exactly what has to go if these are to be counted.
    match said.find('(') {
        Some(at) => {
            let kind = &said[..at];
            let inside = said[at + 1..].trim_end_matches(')').trim_matches('"');
            // A message is worth keeping; a lone token is not.
            if inside.len() > 12 {
                format!("{kind}: {inside}")
            } else {
                kind.to_string()
            }
        }
        None => said,
    }
}
