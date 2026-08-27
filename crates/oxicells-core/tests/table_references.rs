// This Source Code Form is subject to the terms of the Mozilla Public
// License, v. 2.0. If a copy of the MPL was not distributed with this
// file, You can obtain one at https://mozilla.org/MPL/2.0/.

//! Where a table sits, as a formula that names it sees it.
//!
//! The IR counts a table's rows from ONE and the evaluator counts every row
//! from zero. Handing the numbers across as they stood put every table one row
//! lower than it is, and it showed up twice in one workbook: a row asking how
//! far below the heading it was answered 7 where Excel said 8, and a lookup
//! over a table's columns read the headings as data, lost the last row, and
//! came back `#N/A`.
//!
//! Nothing in the calculator's own tests could catch it — they add a table by
//! calling `add_table` directly, in the calculator's own counting — so the
//! test has to start from a file.

use oxicells_core::{editor::CellEditValue, parse_xlsx, XlsxEditor};

const FIXTURE: &[u8] = include_bytes!("fixtures/frozen.xlsx");

/// The table part: a heading row at row 1 and three rows of data under it,
/// spanning A1:C4.
const TABLE: &str = concat!(
    r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>"#,
    r#"<table xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main""#,
    r#" id="1" name="Staff" displayName="Staff" ref="A1:C4" totalsRowShown="0">"#,
    r#"<autoFilter ref="A1:C4"/>"#,
    r#"<tableColumns count="3">"#,
    r#"<tableColumn id="1" name="ID"/>"#,
    r#"<tableColumn id="2" name="NAME"/>"#,
    r#"<tableColumn id="3" name="PAY"/>"#,
    r#"</tableColumns>"#,
    r#"</table>"#,
);

const SHEET_RELS: &str = concat!(
    r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>"#,
    r#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">"#,
    r#"<Relationship Id="rId9""#,
    r#" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/table""#,
    r#" Target="../tables/table1.xml"/>"#,
    r#"</Relationships>"#,
);

/// The fixture with a table part bolted on, and the cells the table describes
/// actually filled in.
fn a_workbook_with_a_table(formulas: &[(u32, u32, &str)]) -> Vec<u8> {
    let mut editor = XlsxEditor::new(FIXTURE).expect("opens");
    for (at, heading) in ["ID", "NAME", "PAY"].iter().enumerate() {
        editor.set_cell_value(
            0,
            1,
            at as u32,
            CellEditValue::String(heading.to_string()),
        );
    }
    // Three rows of data, rows 2 to 4.
    for row in 2..=4u32 {
        editor.set_cell_value(0, row, 0, CellEditValue::Number(row as f64 - 1.0));
        editor.set_cell_value(0, row, 1, CellEditValue::String(format!("w{row}")));
        editor.set_cell_value(0, row, 2, CellEditValue::Number(row as f64 * 100.0));
    }
    for (row, col, formula) in formulas {
        editor.set_cell_value(0, *row, *col, CellEditValue::Formula(formula.to_string()));
    }
    with_table(&editor.save().expect("saves"))
}

fn with_table(source: &[u8]) -> Vec<u8> {
    let mut out = Vec::new();
    {
        let reader = std::io::Cursor::new(source);
        let mut archive = zip::ZipArchive::new(reader).expect("a zip");
        let mut writer = zip::ZipWriter::new(std::io::Cursor::new(&mut out));
        for at in 0..archive.len() {
            let mut part = archive.by_index(at).expect("a part");
            let name = part.name().to_string();
            if name == "xl/worksheets/_rels/sheet1.xml.rels" {
                continue;
            }
            let mut bytes = Vec::new();
            std::io::copy(&mut part, &mut bytes).expect("reads");
            writer
                .start_file(name, zip::write::SimpleFileOptions::default())
                .expect("starts");
            use std::io::Write;
            writer.write_all(&bytes).expect("writes");
        }
        use std::io::Write;
        for (name, body) in [
            ("xl/tables/table1.xml", TABLE),
            ("xl/worksheets/_rels/sheet1.xml.rels", SHEET_RELS),
        ] {
            writer
                .start_file(name, zip::write::SimpleFileOptions::default())
                .expect("starts");
            writer.write_all(body.as_bytes()).expect("writes");
        }
        writer.finish().expect("finishes");
    }
    out
}

/// What a cell shows, as text. `CellValue` states no equality of its own, and
/// what is being asked here is what the sheet reads as anyway.
fn shown(book: &oxicells_core::ir::Workbook, row: u32, col: u32) -> String {
    book.sheets[0]
        .rows
        .iter()
        .find(|one| one.index == row)
        .and_then(|one| one.cells.iter().find(|cell| cell.col == col))
        .map(|cell| cell.value.display())
        .unwrap_or_default()
}

#[test]
fn the_table_is_read_with_its_name_and_its_columns() {
    let book = parse_xlsx(&a_workbook_with_a_table(&[])).expect("parses");
    let table = &book.sheets[0].tables[0];
    assert_eq!(table.name, "Staff");
    assert_eq!(table.columns, vec!["ID", "NAME", "PAY"]);
    // The IR counts a table's rows from one: the heading is row 1.
    assert_eq!((table.start_row, table.end_row), (1, 4));
    assert_eq!((table.start_col, table.end_col), (0, 2));
}

#[test]
fn a_table_sits_where_the_file_says_it_sits() {
    // `ROW() - ROW(Staff[[#Headers],[ID]])` is how far below the heading this
    // row is. Written in row 4 it is 3, and it comes out 2 if the heading is
    // reported one row too far down.
    let mut book = parse_xlsx(&a_workbook_with_a_table(&[
        (4, 4, "=ROW()-ROW(Staff[[#Headers],[ID]])"),
        (5, 4, "=SUM(Staff[PAY])"),
        (6, 4, "=Staff[[#Headers],[PAY]]"),
        (2, 5, "=Staff[[#This Row],[PAY]]"),
    ]))
    .expect("parses");
    oxicells_core::formula::evaluate_workbook_formulas(&mut book);
    assert_eq!(shown(&book, 4, 4), "3", "three rows below the heading");
    // 200 + 300 + 400, the data and not the heading.
    assert_eq!(shown(&book, 5, 4), "900");
    assert_eq!(shown(&book, 6, 4), "PAY", "the heading itself");
    // Row 2 is the first row of data.
    assert_eq!(shown(&book, 2, 5), "200");
}
