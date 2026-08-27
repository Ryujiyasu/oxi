// This Source Code Form is subject to the terms of the Mozilla Public
// License, v. 2.0. If a copy of the MPL was not distributed with this
// file, You can obtain one at https://mozilla.org/MPL/2.0/.

// How long does recalculating a whole workbook take, with no bridge in the way?
//
// The browser needs to know what a typed formula comes to. Sending the whole
// IR across the wasm bridge to find out costs about 17us a cell, which on a
// half-million-cell sheet is several seconds. This asks what the work itself
// costs, so it is clear whether the bridge is the problem or the evaluator is.
use std::time::Instant;

fn main() {
    for path in std::env::args().skip(1) {
        let data = std::fs::read(&path).expect("read");
        let book = match oxicells_core::parse_xlsx(&data) {
            Ok(book) => book,
            Err(error) => {
                println!("  {path}: {error}");
                continue;
            }
        };
        let cells: usize = book
            .sheets
            .iter()
            .flat_map(|sheet| sheet.rows.iter())
            .map(|row| row.cells.len())
            .sum();
        let formulas: usize = book
            .sheets
            .iter()
            .flat_map(|sheet| sheet.rows.iter())
            .flat_map(|row| row.cells.iter())
            .filter(|cell| cell.formula.is_some())
            .count();
        let mut runs = Vec::new();
        for _ in 0..5 {
            let mut copy = book.clone();
            let started = Instant::now();
            oxicells_core::formula::evaluate_workbook_formulas(&mut copy);
            runs.push(started.elapsed().as_secs_f64() * 1000.0);
        }
        runs.sort_by(|a, b| a.partial_cmp(b).unwrap());
        let name = std::path::Path::new(&path)
            .file_name()
            .map(|one| one.to_string_lossy().into_owned())
            .unwrap_or_else(|| path.clone());
        println!("  {name:<44}{cells:>8} cells{formulas:>6} formulas   median {:.1}ms", runs[2]);
    }
}
