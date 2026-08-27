// This Source Code Form is subject to the terms of the Mozilla Public
// License, v. 2.0. If a copy of the MPL was not distributed with this
// file, You can obtain one at https://mozilla.org/MPL/2.0/.

//! Opens one workbook and says what was in it, as a line of JSON.
//!
//! Used to sweep a corpus nothing has been fitted to. What matters is not any
//! single number but the shape of the whole: how many opened at all, and what
//! the ones that did turn out to contain — a corpus full of features the
//! engine walks past is a different problem from one it cannot read.

use std::collections::BTreeSet;

fn main() {
    let path = match std::env::args().nth(1) {
        Some(path) => path,
        None => {
            eprintln!("give it a workbook");
            std::process::exit(2);
        }
    };
    let data = match std::fs::read(&path) {
        Ok(data) => data,
        Err(error) => {
            eprintln!("{error}");
            std::process::exit(1);
        }
    };
    let book = match oxicells_core::parse_xlsx(&data) {
        Ok(book) => book,
        Err(error) => {
            eprintln!("{error}");
            std::process::exit(1);
        }
    };

    let mut cells = 0usize;
    let mut formulas = 0usize;
    let mut merges = 0usize;
    let mut frozen = 0usize;
    let mut styled = 0usize;
    let mut wrapped = 0usize;
    let mut dated = 0usize;
    let mut rows = 0usize;
    let mut unread: BTreeSet<String> = BTreeSet::new();
    for sheet in &book.sheets {
        rows += sheet.rows.len();
        merges += sheet.merge_cells.len();
        if sheet.frozen_rows > 0 || sheet.frozen_cols > 0 {
            frozen += 1;
        }
        for name in &sheet.unsupported_elements {
            unread.insert(name.clone());
        }
        for row in &sheet.rows {
            for cell in &row.cells {
                cells += 1;
                if cell.formula.is_some() {
                    formulas += 1;
                }
                let style = &cell.style;
                if style.bold || style.italic || style.underline
                    || style.font_color.is_some() || style.bg_color.is_some()
                    || style.horizontal_align.is_some() || style.indent > 0
                {
                    styled += 1;
                }
                if style.wrap_text {
                    wrapped += 1;
                }
                // A date is a number wearing a format that mentions a day or a
                // year, which is the only thing that makes it one.
                if let Some(format) = &style.number_format {
                    let bare: String = format.replace('\\', "");
                    if bare.contains('y') || bare.contains('d')
                        || bare.contains('Y') || bare.contains('D')
                    {
                        dated += 1;
                    }
                }
            }
        }
    }

    let unread: Vec<String> = unread.into_iter().collect();
    println!(
        "{{\"sheets\":{},\"rows\":{},\"cells\":{},\"formulas\":{},\"merges\":{},\
         \"frozen\":{},\"styled\":{},\"wrapped\":{},\"dated\":{},\"unread\":{}}}",
        book.sheets.len(),
        rows,
        cells,
        formulas,
        merges,
        frozen,
        styled,
        wrapped,
        dated,
        // Written by hand rather than with a serialiser, so the example needs
        // nothing the crate does not already depend on.
        format!(
            "[{}]",
            unread
                .iter()
                .map(|one| format!("{:?}", one))
                .collect::<Vec<_>>()
                .join(",")
        ),
    );
}
