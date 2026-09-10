// This Source Code Form is subject to the terms of the Mozilla Public
// License, v. 2.0. If a copy of the MPL was not distributed with this
// file, You can obtain one at https://mozilla.org/MPL/2.0/.

//! What this engine thinks a sheet's conditional rules catch, as JSON lines.
//!
//! The other half of `tools/metrics/xlsx_conditional_agreement.py`, which asks
//! Excel the same question through `Range.DisplayFormat` — the property that
//! answers with what a person actually sees, conditional formatting included.
//! Comparing the two is the only way to tell "we applied a rule" from "we
//! applied the rule Excel applies".
//!
//!     cargo run --release -p oxicells-core --example _conditional_dump -- book.xlsx
//!
//! Rows and columns come out counted from zero.

/// `Option<bool>` in JSON: `true`, `false`, or `null` — the Rust debug spelling
/// is not JSON, and a reader that skips unparseable lines would drop exactly
/// the hits that carry a weight.
fn yes(held: Option<bool>) -> &'static str {
    match held {
        Some(true) => "true",
        Some(false) => "false",
        None => "null",
    }
}

fn main() {
    let Some(path) = std::env::args().nth(1) else {
        eprintln!("give it a workbook");
        std::process::exit(2);
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

    for (at, sheet) in book.sheets.iter().enumerate() {
        // The ranges the rules cover, so the comparison can also ask about the
        // cells that were NOT caught: a rule that fires nowhere and a rule that
        // fires everywhere are both wrong, and only one of them shows up in a
        // list of hits.
        for rule in &sheet.conditional_rules {
            for (top, left, bottom, right) in &rule.ranges {
                println!(
                    "{{\"kind\":\"range\",\"sheet\":{at},\"name\":{:?},\
                     \"top\":{top},\"left\":{left},\"bottom\":{bottom},\"right\":{right},\
                     \"rule\":{:?},\"styled\":{},\"op\":{:?},\"formulas\":{:?}}}",
                    sheet.name,
                    rule.kind,
                    rule.style.is_some(),
                    rule.operator.clone().unwrap_or_default(),
                    rule.formulas,
                );
            }
        }
        for hit in oxicells_core::conditional::caught(&book, at) {
            let style = &hit.style;
            println!(
                "{{\"kind\":\"hit\",\"sheet\":{at},\"row\":{},\"col\":{},\
                 \"bg\":{:?},\"fg\":{:?},\"bold\":{},\"italic\":{}}}",
                hit.row,
                hit.col,
                style.bg_color.clone().unwrap_or_default(),
                style.font_color.clone().unwrap_or_default(),
                yes(style.bold),
                yes(style.italic),
            );
        }
    }
}
