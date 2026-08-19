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
