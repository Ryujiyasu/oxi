// This Source Code Form is subject to the terms of the Mozilla Public
// License, v. 2.0. If a copy of the MPL was not distributed with this
// file, You can obtain one at https://mozilla.org/MPL/2.0/.

//! What survives a round trip when a workbook gains, loses or renames a sheet.

use oxicells_core::editor::XlsxEditor;
use oxicells_core::ir::{Cell, CellStyle, CellValue, Row, Sheet};
use oxicells_core::parser::parse_xlsx;

const FIXTURE: &[u8] = include_bytes!("fixtures/hidden_rows_cols.xlsx");

fn names(workbook: &oxicells_core::ir::Workbook) -> Vec<String> {
    workbook.sheets.iter().map(|s| s.name.clone()).collect()
}

fn blank_sheet(name: &str) -> Sheet {
    Sheet {
        name: name.to_string(),
        rows: Vec::new(),
        col_count: 0,
        col_widths: Vec::new(),
        default_col_width: 8.43,
        default_row_height: 15.0,
        merge_cells: Vec::new(),
        hidden_cols: Vec::new(),
        auto_filter: None,
        declared_range: None,
        tables: Vec::new(),
        unsupported_elements: Vec::new(),
    }
}

fn saved_after(change: impl FnOnce(&mut oxicells_core::ir::Workbook)) -> Vec<u8> {
    let mut workbook = parse_xlsx(FIXTURE).expect("the fixture parses");
    change(&mut workbook);
    let mut editor = XlsxEditor::new(FIXTURE).expect("the fixture opens");
    editor
        .apply_workbook(&workbook)
        .expect("the change is one the editor can write");
    editor.save().expect("the workbook saves")
}

#[test]
fn a_sheet_a_run_added_is_saved() {
    let saved = saved_after(|workbook| {
        let mut added = blank_sheet("Added");
        added.rows.push(Row {
            index: 1,
            height: None,
            hidden: false,
            cells: vec![Cell {
                col: 0,
                value: CellValue::Number(42.0),
                style: CellStyle::default(),
                formula: None,
            }],
        });
        workbook.sheets.push(added);
    });

    let reread = parse_xlsx(&saved).expect("the saved workbook parses");
    assert_eq!(names(&reread), vec!["Sheet1".to_string(), "Added".to_string()]);
    assert_eq!(reread.sheets[1].rows[0].cells[0].value.display(), "42");
    // The sheet that was already there is untouched.
    let hidden: Vec<u32> = reread.sheets[0]
        .rows
        .iter()
        .filter(|row| row.hidden)
        .map(|row| row.index)
        .collect();
    assert_eq!(hidden, vec![2, 4]);
}

#[test]
fn a_sheet_a_run_added_in_front_keeps_its_place() {
    let saved = saved_after(|workbook| {
        workbook.sheets.insert(0, blank_sheet("Front"));
    });
    let reread = parse_xlsx(&saved).expect("parses");
    assert_eq!(names(&reread), vec!["Front".to_string(), "Sheet1".to_string()]);
}

#[test]
fn a_sheet_a_run_removed_is_gone() {
    let saved = saved_after(|workbook| {
        workbook.sheets.push(blank_sheet("Spare"));
    });
    // Take it away again, starting from the workbook that has it.
    let mut workbook = parse_xlsx(&saved).expect("parses");
    workbook.sheets.retain(|sheet| sheet.name != "Spare");
    let mut editor = XlsxEditor::new(&saved).expect("opens");
    editor.apply_workbook(&workbook).expect("writes");
    let saved = editor.save().expect("saves");

    let reread = parse_xlsx(&saved).expect("parses");
    assert_eq!(names(&reread), vec!["Sheet1".to_string()]);
}

#[test]
fn a_renamed_sheet_keeps_what_it_holds() {
    let saved = saved_after(|workbook| {
        workbook.sheets[0].name = "Renamed".to_string();
    });
    let reread = parse_xlsx(&saved).expect("parses");
    assert_eq!(names(&reread), vec!["Renamed".to_string()]);
    // Renaming does not disturb the rows it hides.
    let hidden: Vec<u32> = reread.sheets[0]
        .rows
        .iter()
        .filter(|row| row.hidden)
        .map(|row| row.index)
        .collect();
    assert_eq!(hidden, vec![2, 4]);
}

#[test]
fn leaving_the_sheets_alone_writes_nothing_back() {
    let workbook = parse_xlsx(FIXTURE).expect("parses");
    let mut editor = XlsxEditor::new(FIXTURE).expect("opens");
    editor.apply_workbook(&workbook).expect("nothing to write");
    assert!(!editor.has_edits());
}

