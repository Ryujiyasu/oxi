// This Source Code Form is subject to the terms of the Mozilla Public
// License, v. 2.0. If a copy of the MPL was not distributed with this
// file, You can obtain one at https://mozilla.org/MPL/2.0/.

//! Where does a recalculation's time actually go?
//!
//! A workbook of 936,000 cells and 22,864 formulas took 37 seconds, and after
//! the dependency builder stopped walking the sheet it took 22. Knowing which
//! phase the remaining 22 belongs to decides what to do next: building the
//! graph is a cost the editor pays on every keystroke, while evaluating is a
//! cost it should only pay for the cells a change can reach.
use std::time::Instant;

const SEPARATORS: [char; 2] = ['/', '\\'];

fn main() {
    for path in std::env::args().skip(1) {
        let data = std::fs::read(&path).expect("read");
        let started = Instant::now();
        let Ok(book) = oxicells_core::parse_xlsx(&data) else {
            continue;
        };
        let parsed = started.elapsed();

        let started = Instant::now();
        let mut engine = oxicells_calc::Workbook::new();
        for sheet in &book.sheets {
            engine.add_sheet(&sheet.name);
            for row in &sheet.rows {
                for cell in &row.cells {
                    let a1 = format!(
                        "{}{}",
                        oxicells_calc::reference::col_to_letters(cell.col),
                        row.index
                    );
                    match &cell.formula {
                        Some(text) => {
                            let _ = engine.set_formula(&sheet.name, &a1, text);
                        }
                        None => {
                            let _ = engine.set_value(&sheet.name, &a1, as_value(&cell.value));
                        }
                    }
                }
            }
        }
        let loaded = started.elapsed();

        let started = Instant::now();
        let report = engine.recalculate();
        let worked = started.elapsed();

        // And what one edit costs. The graph still has to be built — a cell
        // that has just been given a formula may read anything — but only the
        // formulas the change reaches are worked out.
        let first = book
            .sheets
            .first()
            .map(|sheet| sheet.name.clone())
            .unwrap_or_default();
        let started = Instant::now();
        let touched = engine.recalculate_after(&[(first, (0, 0))]);
        let edit = started.elapsed();

        println!(
            "  {:<38} parse {:>6}ms  load {:>5}ms  full {:>6}ms ({:>5} cells)  edit {:>5}ms ({} cells)",
            path.rsplit(SEPARATORS).next().unwrap_or(&path),
            parsed.as_millis(),
            loaded.as_millis(),
            worked.as_millis(),
            report.evaluated,
            edit.as_millis(),
            touched.evaluated,
        );
    }
}

fn as_value(value: &oxicells_core::ir::CellValue) -> oxicells_calc::Value {
    use oxicells_core::ir::CellValue;
    match value {
        CellValue::Number(n) => oxicells_calc::Value::Number(*n),
        CellValue::String(s) => oxicells_calc::Value::text(s.clone()),
        CellValue::Boolean(b) => oxicells_calc::Value::Logical(*b),
        _ => oxicells_calc::Value::Blank,
    }
}
