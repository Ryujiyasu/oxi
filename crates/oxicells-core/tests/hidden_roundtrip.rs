// This Source Code Form is subject to the terms of the Mozilla Public
// License, v. 2.0. If a copy of the MPL was not distributed with this
// file, You can obtain one at https://mozilla.org/MPL/2.0/.

//! What survives a round trip through the editor, for rows and columns a
//! workbook hides.
//!
//! The fixture was written by Excel 16: five rows of data, with rows 2 and 4
//! hidden and column B hidden.

use oxicells_core::editor::{CellEditValue, XlsxEditor};
use oxicells_core::parser::parse_xlsx;

const FIXTURE: &[u8] = include_bytes!("fixtures/hidden_rows_cols.xlsx");

fn hidden_rows(workbook: &oxicells_core::ir::Workbook) -> Vec<u32> {
    workbook.sheets[0]
        .rows
        .iter()
        .filter(|row| row.hidden)
        .map(|row| row.index)
        .collect()
}

#[test]
fn reading_a_workbook_finds_the_rows_and_columns_it_hides() {
    let workbook = parse_xlsx(FIXTURE).expect("the fixture parses");
    assert_eq!(hidden_rows(&workbook), vec![2, 4]);
    assert_eq!(workbook.sheets[0].hidden_cols, vec![1]);
}

/// The editor patches the worksheet XML in place rather than writing it afresh,
/// so what it was not asked to change should come back untouched.
#[test]
fn editing_a_cell_leaves_hidden_rows_and_columns_alone() {
    let mut editor = XlsxEditor::new(FIXTURE).expect("the fixture opens");
    editor.set_cell_value(0, 1, 0, CellEditValue::Number(99.0));
    let saved = editor.save().expect("the workbook saves");

    let workbook = parse_xlsx(&saved).expect("the saved workbook parses");
    assert_eq!(hidden_rows(&workbook), vec![2, 4]);
    assert_eq!(workbook.sheets[0].hidden_cols, vec![1]);
}

/// Editing a cell *on* a hidden row must not reveal it.
#[test]
fn editing_a_hidden_rows_cell_keeps_it_hidden() {
    let mut editor = XlsxEditor::new(FIXTURE).expect("the fixture opens");
    editor.set_cell_value(0, 2, 0, CellEditValue::Number(77.0));
    let saved = editor.save().expect("the workbook saves");

    let workbook = parse_xlsx(&saved).expect("the saved workbook parses");
    assert_eq!(hidden_rows(&workbook), vec![2, 4]);
}

/// Inserting a cell into a row the fixture never wrote must not disturb the
/// rows around it.
#[test]
fn adding_a_row_leaves_the_hidden_ones_hidden() {
    let mut editor = XlsxEditor::new(FIXTURE).expect("the fixture opens");
    editor.set_cell_value(0, 9, 0, CellEditValue::Number(1.0));
    let saved = editor.save().expect("the workbook saves");

    let workbook = parse_xlsx(&saved).expect("the saved workbook parses");
    assert_eq!(hidden_rows(&workbook), vec![2, 4]);
    assert_eq!(workbook.sheets[0].hidden_cols, vec![1]);
}

/// The editor can hide a row that was visible, and reveal one that was not.
#[test]
fn the_editor_can_change_which_rows_are_hidden() {
    let mut editor = XlsxEditor::new(FIXTURE).expect("the fixture opens");
    editor.set_row_hidden(0, 2, false); // was hidden
    editor.set_row_hidden(0, 3, true); // was visible
    let saved = editor.save().expect("the workbook saves");

    let workbook = parse_xlsx(&saved).expect("the saved workbook parses");
    assert_eq!(hidden_rows(&workbook), vec![3, 4]);
    // The cells on those rows are untouched by the change.
    let row3 = workbook.sheets[0]
        .rows
        .iter()
        .find(|row| row.index == 3)
        .expect("row 3 is still there");
    assert_eq!(row3.cells[0].value.display(), "30");
}

#[test]
fn the_editor_can_change_which_columns_are_hidden() {
    let mut editor = XlsxEditor::new(FIXTURE).expect("the fixture opens");
    editor.set_col_hidden(0, 1, false); // column B was hidden
    editor.set_col_hidden(0, 2, true); // column C was visible
    let saved = editor.save().expect("the workbook saves");

    let workbook = parse_xlsx(&saved).expect("the saved workbook parses");
    assert_eq!(workbook.sheets[0].hidden_cols, vec![2]);
}

/// Hiding a row the sheet has no record of still has to be written down.
#[test]
fn hiding_an_empty_row_writes_it_out() {
    let mut editor = XlsxEditor::new(FIXTURE).expect("the fixture opens");
    editor.set_row_hidden(0, 12, true);
    let saved = editor.save().expect("the workbook saves");

    let workbook = parse_xlsx(&saved).expect("the saved workbook parses");
    assert_eq!(hidden_rows(&workbook), vec![2, 4, 12]);
}

/// Both kinds of change at once, alongside a cell edit.
#[test]
fn cells_rows_and_columns_all_change_together() {
    let mut editor = XlsxEditor::new(FIXTURE).expect("the fixture opens");
    editor.set_cell_value(0, 1, 0, CellEditValue::Number(99.0));
    editor.set_row_hidden(0, 5, true);
    editor.set_col_hidden(0, 0, true);
    let saved = editor.save().expect("the workbook saves");

    let workbook = parse_xlsx(&saved).expect("the saved workbook parses");
    assert_eq!(hidden_rows(&workbook), vec![2, 4, 5]);
    assert_eq!(workbook.sheets[0].hidden_cols, vec![0, 1]);
    let first = workbook.sheets[0]
        .rows
        .iter()
        .find(|row| row.index == 1)
        .expect("row 1 is still there");
    assert_eq!(first.cells[0].value.display(), "99");
}
