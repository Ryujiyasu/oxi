// This Source Code Form is subject to the terms of the Mozilla Public
// License, v. 2.0. If a copy of the MPL was not distributed with this
// file, You can obtain one at https://mozilla.org/MPL/2.0/.

//! The notes a workbook keeps pinned open on its first sheet.
//!
//! ```text
//! cargo run --release -p oxicells-core --example comment_dump -- path\to\book.xlsx
//! ```

use std::path::Path;

fn main() {
    let target = std::env::args()
        .nth(1)
        .unwrap_or_else(|| r"tools\golden-test\documents\xlsx".to_string());
    let path = Path::new(&target);
    let books: Vec<_> = if path.is_dir() {
        let mut held: Vec<_> = std::fs::read_dir(path)
            .expect("the directory can be read")
            .flatten()
            .map(|entry| entry.path())
            .filter(|path| path.extension().is_some_and(|kind| kind == "xlsx"))
            .collect();
        held.sort();
        held
    } else {
        vec![path.to_path_buf()]
    };

    for book in books {
        let Ok(bytes) = std::fs::read(&book) else { continue };
        let Ok(workbook) = oxicells_core::parser::parse_xlsx_preserving_values(&bytes) else {
            continue;
        };
        let Some(sheet) = workbook.sheets.first() else { continue };
        if sheet.comments.is_empty() {
            continue;
        }
        println!(
            "{:<46} {} pinned open",
            book.file_stem().unwrap_or_default().to_string_lossy(),
            sheet.comments.len()
        );
        for note in sheet.comments.iter().take(4) {
            println!(
                "    from cell {},{} + {},{}emu  {:.0}x{:.0}pt  {:?}",
                note.from.col,
                note.from.row,
                note.from.col_off,
                note.from.row_off,
                note.size.0,
                note.size.1,
                note.text
                    .paragraphs
                    .iter()
                    .map(|held| held.text.chars().take(14).collect::<String>())
                    .collect::<Vec<_>>()
            );
        }
    }
}
