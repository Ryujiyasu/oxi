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

#[test]
fn tables_added_renamed_resized_and_removed_survive_save() {
    let source = with_table(FIXTURE);
    let template = parse_xlsx(&source).unwrap().sheets[0].tables[0].clone();
    let mut editor = XlsxEditor::new(FIXTURE).unwrap();
    let mut book = editor.workbook().clone();
    let mut table = template;
    table.name = "AddedTable".into();
    table.style = Some("TableStyleMedium2".into());
    book.sheets[0].tables.push(table);
    editor.apply_workbook(&book).unwrap();
    let saved = editor.save().unwrap();
    let reopened = parse_xlsx(&saved).unwrap();
    assert_eq!(reopened.sheets[0].tables.len(), 1);
    assert_eq!(reopened.sheets[0].tables[0].name, "AddedTable");

    let mut editor = XlsxEditor::new(&saved).unwrap();
    let mut book = editor.workbook().clone();
    let table = &mut book.sheets[0].tables[0];
    table.name = "RenamedTable".into();
    table.columns[1] = "DisplayName".into();
    table.end_row += 1;
    table.banded_rows = true;
    editor.apply_workbook(&book).unwrap();
    let renamed = editor.save().unwrap();
    let reopened = parse_xlsx(&renamed).unwrap();
    let table = &reopened.sheets[0].tables[0];
    assert_eq!(table.name, "RenamedTable");
    assert_eq!(table.columns[1], "DisplayName");
    assert_eq!(table.end_row, 5);
    assert!(table.banded_rows);

    let mut editor = XlsxEditor::new(&renamed).unwrap();
    let mut book = editor.workbook().clone();
    book.sheets[0].tables.clear();
    editor.apply_workbook(&book).unwrap();
    let deleted = editor.save().unwrap();
    assert!(parse_xlsx(&deleted).unwrap().sheets[0].tables.is_empty());
    let zip = zip::ZipArchive::new(std::io::Cursor::new(&deleted)).unwrap();
    assert!(!zip.file_names().any(|name| name.starts_with("xl/tables/")));
}

#[test]
fn simultaneous_table_renames_and_resizes_retain_source_identity() {
    let source = with_table(FIXTURE);
    let template = parse_xlsx(&source).unwrap().sheets[0].tables[0].clone();
    let mut editor = XlsxEditor::new(FIXTURE).unwrap();
    let mut book = editor.workbook().clone();
    for (name, offset) in [("First", 0), ("Second", 3)] {
        let mut table = template.clone();
        table.origin_part = None;
        table.name = name.into();
        table.start_col += offset;
        table.end_col += offset;
        book.sheets[0].tables.push(table);
    }
    editor.apply_workbook(&book).unwrap();
    let saved = editor.save().unwrap();
    let mut editor = XlsxEditor::new(&saved).unwrap();
    let mut book = editor.workbook().clone();
    let originals = book.sheets[0].tables.iter().map(|table| (format!("{}Updated", table.name), table.origin_part.clone())).collect::<std::collections::HashMap<_, _>>();
    book.sheets[0].tables.reverse();
    for table in &mut book.sheets[0].tables {
        table.name.push_str("Updated");
        table.end_row += 1;
    }
    editor.apply_workbook(&book).unwrap();
    let reopened = parse_xlsx(&editor.save().unwrap()).unwrap();
    for table in &reopened.sheets[0].tables {
        assert_eq!(Some(&table.origin_part), originals.get(&table.name));
        assert_eq!(table.end_row, 5);
    }
}

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
    with_table_xml(source, TABLE)
}

fn with_table_xml(source: &[u8], table: &str) -> Vec<u8> {
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
            ("xl/tables/table1.xml", table),
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

#[test]
fn table_style_and_filter_buttons_survive_edits_and_save() {
    let table_xml = TABLE.replace("<autoFilter ref=\"A1:C4\"/>",
        "<autoFilter ref=\"A1:C4\"><filterColumn colId=\"0\" hiddenButton=\"1\"><filters><filter val=\"x\"/></filters></filterColumn><filterColumn colId=\"2\" showButton=\"0\"/></autoFilter>")
        .replace("</table>", "<tableStyleInfo name=\"TableStyleMedium2\" showFirstColumn=\"1\" showLastColumn=\"1\" showRowStripes=\"0\" showColumnStripes=\"1\"/></table>");
    let source = with_table_xml(FIXTURE, &table_xml);
    let mut editor = XlsxEditor::new(&source).unwrap();
    let mut book = editor.workbook().clone();
    let table = &mut book.sheets[0].tables[0];
    assert!(table.emphasize_first_column && table.emphasize_last_column && table.banded_columns);
    assert!(!table.banded_rows);
    assert!(table.show_auto_filter);
    assert_eq!(table.hidden_filter_buttons, vec![0, 2]);
    table.emphasize_first_column = false;
    table.emphasize_last_column = false;
    table.banded_columns = false;
    table.banded_rows = true;
    table.hidden_filter_buttons = vec![1, 2];
    editor.apply_workbook(&book).unwrap();
    let bytes = editor.save().unwrap();
    let reopened = parse_xlsx(&bytes).unwrap();
    let table = &reopened.sheets[0].tables[0];
    assert!(!table.emphasize_first_column && !table.emphasize_last_column && !table.banded_columns);
    assert!(table.banded_rows && table.show_auto_filter);
    assert_eq!(table.hidden_filter_buttons, vec![1, 2]);
    use std::io::Read;
    let mut zip = zip::ZipArchive::new(std::io::Cursor::new(&bytes)).unwrap();
    let mut xml = String::new();
    zip.by_name("xl/tables/table1.xml").unwrap().read_to_string(&mut xml).unwrap();
    assert!(xml.contains("<filters><filter val=\"x\"/></filters>"), "existing criteria must survive button edits");
    let mut editor = XlsxEditor::new(&bytes).unwrap();
    let mut book = editor.workbook().clone();
    book.sheets[0].tables[0].show_auto_filter = false;
    book.sheets[0].tables[0].hidden_filter_buttons.clear();
    editor.apply_workbook(&book).unwrap();
    let saved = editor.save().unwrap();
    let table = &parse_xlsx(&saved).unwrap().sheets[0].tables[0];
    assert!(!table.show_auto_filter);
    assert!(table.hidden_filter_buttons.is_empty());
}

#[test]
fn empty_hidden_tables_and_explicit_filter_resets_survive_save() {
    let source_xml = TABLE.replace("ref=\"A1:C4\" totalsRowShown=\"0\"", "ref=\"A2:C2\" headerRowCount=\"0\" insertRow=\"1\" totalsRowShown=\"0\"")
        .replace("<autoFilter ref=\"A1:C4\"/>", "");
    let source = with_table_xml(FIXTURE, &source_xml);
    let mut editor = XlsxEditor::new(&source).unwrap();
    let book = editor.workbook().clone();
    let table = &book.sheets[0].tables[0];
    assert!(table.insert_row);
    assert_eq!(table.header_rows, 0);
    assert!(!table.show_auto_filter);
    editor.apply_workbook(&book).unwrap();
    let reopened = parse_xlsx(&editor.save().unwrap()).unwrap();
    let table = &reopened.sheets[0].tables[0];
    assert!(table.insert_row);
    assert_eq!((table.start_row, table.end_row, table.header_rows), (2, 2, 0));

    // Criteria may contain extensions whose prefixes were declared only on
    // the source table. They must remain valid after moving to a new part.
    let source_xml = TABLE.replace(" id=\"1\"", " xmlns:custom=\"urn:oxi:test-filter\" id=\"1\"")
        .replace("<autoFilter ref=\"A1:C4\"/>", "<autoFilter ref=\"A1:C4\"><filterColumn colId=\"0\"><filters><filter val=\"x\"/></filters><extLst><ext uri=\"test\"><custom:state/></ext></extLst></filterColumn></autoFilter>");
    let source = with_table_xml(FIXTURE, &source_xml);
    let mut editor = XlsxEditor::new(&source).unwrap();
    let mut book = editor.workbook().clone();
    let table = &mut book.sheets[0].tables[0];
    assert!(table.auto_filter_xml.as_deref().unwrap().contains("xmlns:custom=\"urn:oxi:test-filter\""));
    table.origin_part = None;
    editor.apply_workbook(&book).unwrap();
    let saved = editor.save().unwrap();
    let filter = parse_xlsx(&saved).unwrap().sheets[0].tables[0].auto_filter_xml.clone().unwrap();
    assert!(filter.contains("<filter val=\"x\"/>") && filter.contains("<custom:state/>") && filter.contains("xmlns:custom="));

    // Clear and re-enable before saving: the original XML must not restore
    // the removed criterion or its extension state.
    let mut editor = XlsxEditor::new(&source).unwrap();
    let mut book = editor.workbook().clone();
    let table = &mut book.sheets[0].tables[0];
    table.auto_filter_xml = None;
    table.show_auto_filter = true;
    editor.apply_workbook(&book).unwrap();
    let reopened = parse_xlsx(&editor.save().unwrap()).unwrap();
    let table = &reopened.sheets[0].tables[0];
    assert!(table.show_auto_filter);
    let filter = table.auto_filter_xml.as_deref().unwrap();
    assert!(!filter.contains("filterColumn") && !filter.contains("custom:state"));
}

#[test]
fn empty_table_formula_references_keep_excels_insertion_row() {
    for (headers, first, totals) in [(1, 1, 0), (0, 2, 0), (1, 1, 1), (0, 2, 1), (0, 1, 0)] {
        let table = TABLE.replace("ref=\"A1:C4\" totalsRowShown=\"0\"", &format!("ref=\"A{first}:C{}\" headerRowCount=\"{headers}\" totalsRowCount=\"{totals}\" insertRow=\"1\"", first + headers + totals))
            .replace("<autoFilter ref=\"A1:C4\"/>", "");
        let mut editor = XlsxEditor::new(FIXTURE).unwrap();
        for row in first..=first + headers + totals {
            for column in 0..3 { editor.set_cell_value(0, row, column, CellEditValue::Empty); }
        }
        for (row, formula) in [(1, "SUM(Staff[ID])"), (2, "COUNTA(Staff[ID])"), (3, "ROWS(Staff[ID])"), (4, "INDEX(Staff[ID],1)"), (5, "ROWS(Staff[#All])"), (6, "ROWS(Staff[#Headers])")] {
            editor.set_cell_value(0, row, 5, CellEditValue::Formula(formula.into()));
        }
        let source = with_table_xml(&editor.save().unwrap(), &table);
        let mut book = parse_xlsx(&source).unwrap();
        assert!(book.sheets[0].tables[0].insert_row);
        oxicells_core::formula::evaluate_workbook_formulas(&mut book);
        for (row, expected) in [(1, "0"), (2, "0"), (3, "1"), (4, "0")] { assert_eq!(shown(&book, row, 5), expected); }
        assert_eq!(shown(&book, 5, 5), (headers + totals + 1).to_string());
        assert_eq!(shown(&book, 6, 5), if headers == 0 { "#REF!" } else { "1" });
    }
}

#[test]
fn imported_table_criteria_control_subtotal_visibility() {
    for filtered in [false, true] {
        let source = a_workbook_with_a_table(&[(1, 5, "SUBTOTAL(9,Staff[ID])"), (2, 5, "SUBTOTAL(109,Staff[ID])")]);
        let mut book = parse_xlsx(&source).unwrap();
        for row in &mut book.sheets[0].rows { if row.index == 3 || row.index == 4 { row.hidden = true; } }
        book.sheets[0].tables[0].auto_filter_xml = Some(if filtered {
            "<autoFilter><filterColumn colId=\"0\"><filters><filter val=\"1\"/></filters></filterColumn></autoFilter>"
        } else { "<autoFilter><filterColumn colId=\"0\" hiddenButton=\"1\"/></autoFilter>" }.into());
        assert_eq!(book.sheets[0].tables[0].has_filter_criteria(), filtered);
        let mut editor = XlsxEditor::new(&source).unwrap();
        editor.apply_workbook(&book).unwrap();
        let mut book = parse_xlsx(&editor.save().unwrap()).unwrap();
        assert!(book.sheets[0].auto_filter.is_none(), "criteria are in the table part");
        oxicells_core::formula::evaluate_workbook_formulas(&mut book);
        assert_eq!(shown(&book, 1, 5), if filtered { "1" } else { "6" });
        assert_eq!(shown(&book, 2, 5), "1");
    }
}

#[test]
fn totals_functions_labels_and_custom_utf16_formulas_roundtrip_and_clear() {
    use oxicells_core::ir::{CellValue, TableTotals};
    let source = a_workbook_with_a_table(&[]);
    let mut book = parse_xlsx(&source).unwrap();
    let label = vec![116, 111, 116, 97, 108, 0xD83D, 38];
    let mut formula: Vec<u16> = "SUM(Staff[ID])&\"a".encode_utf16().collect();
    formula.extend([0xDC00, b'b' as u16, b'"' as u16]);
    book.sheets[0].tables[0].totals_rows = 1;
    book.sheets[0].tables[0].totals = vec![
        TableTotals { label: Some(label.clone()), ..Default::default() },
        TableTotals { function: Some("custom".into()), formula: Some(formula.clone()), label: None },
        TableTotals { function: Some("sum".into()), ..Default::default() },
    ];
    let row = book.sheets[0].rows.iter_mut().find(|row| row.index == 4).unwrap();
    for cell in &mut row.cells {
        match cell.col {
            0 => { cell.value = CellValue::from_utf16(label.clone()); cell.formula = None; cell.formula_utf16 = None; }
            1 => { cell.value = CellValue::Empty; cell.formula = Some(String::from_utf16_lossy(&formula)); cell.formula_utf16 = Some(formula.clone()); }
            2 => { cell.value = CellValue::Number(500.0); cell.formula = Some("SUBTOTAL(109,Staff[PAY])".into()); cell.formula_utf16 = None; }
            _ => {}
        }
    }
    let mut editor = XlsxEditor::new(&source).unwrap();
    editor.apply_workbook(&book).unwrap();
    let saved = editor.save().unwrap();
    let reopened = parse_xlsx(&saved).unwrap();
    assert_eq!(reopened.sheets[0].tables[0].totals, book.sheets[0].tables[0].totals);
    let mut editor = XlsxEditor::new(&saved).unwrap();
    let mut book = reopened;
    book.sheets[0].tables[0].totals.fill(TableTotals::default());
    let row = book.sheets[0].rows.iter_mut().find(|row| row.index == 4).unwrap();
    for cell in row.cells.iter_mut().filter(|cell| cell.col < 3) {
        cell.value = CellValue::Empty; cell.formula = None; cell.formula_utf16 = None;
    }
    editor.apply_workbook(&book).unwrap();
    let saved = editor.save().unwrap();
    assert!(parse_xlsx(&saved).unwrap().sheets[0].tables[0].totals.iter().all(TableTotals::is_empty));
    use std::io::Read;
    let mut zip = zip::ZipArchive::new(std::io::Cursor::new(saved)).unwrap();
    let mut xml = String::new();
    zip.by_name("xl/tables/table1.xml").unwrap().read_to_string(&mut xml).unwrap();
    assert!(!xml.contains("totalsRowFunction") && !xml.contains("totalsRowLabel") && !xml.contains("totalsRowFormula"));
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
