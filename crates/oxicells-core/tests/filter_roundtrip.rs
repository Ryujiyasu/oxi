// This Source Code Form is subject to the terms of the Mozilla Public
// License, v. 2.0. If a copy of the MPL was not distributed with this
// file, You can obtain one at https://mozilla.org/MPL/2.0/.

//! What survives a round trip through the editor, for the filter a sheet is
//! under.
//!
//! The fixture was written by Excel 16: a heading row and three rows of data,
//! filtered on the first column for "apple", which leaves row 3 hidden.

use oxicells_core::editor::XlsxEditor;
use oxicells_core::ir::{AutoFilter, AutoFilterColumn};
use oxicells_core::parser::parse_xlsx;

const FIXTURE: &[u8] = include_bytes!("fixtures/filtered.xlsx");

#[test]
fn reading_a_workbook_finds_the_filter_it_is_under() {
    let workbook = parse_xlsx(FIXTURE).expect("the fixture parses");
    let filter = workbook.sheets[0]
        .auto_filter
        .as_ref()
        .expect("the sheet is filtered");

    // A1:B4, with rows one-based and columns zero-based.
    assert_eq!(
        (
            filter.start_row,
            filter.start_col,
            filter.end_row,
            filter.end_col
        ),
        (1, 0, 4, 1)
    );
    assert_eq!(filter.columns.len(), 1);
    assert_eq!(filter.columns[0].field, 1);
    assert_eq!(filter.columns[0].criteria, vec!["apple".to_string()]);
}

/// The rows the filter rejected are hidden, and stay that way.
#[test]
fn the_rows_the_filter_rejected_are_hidden() {
    let workbook = parse_xlsx(FIXTURE).expect("the fixture parses");
    let hidden: Vec<u32> = workbook.sheets[0]
        .rows
        .iter()
        .filter(|row| row.hidden)
        .map(|row| row.index)
        .collect();
    assert_eq!(hidden, vec![3]);
}

fn filter_of(workbook: &oxicells_core::ir::Workbook) -> Option<AutoFilter> {
    workbook.sheets[0].auto_filter.clone()
}

/// A filter the editor sets survives the trip out and back.
#[test]
fn a_filter_the_editor_sets_is_saved() {
    let mut editor = XlsxEditor::new(FIXTURE).expect("the fixture opens");
    editor.set_auto_filter(
        0,
        Some(AutoFilter {
            start_row: 1,
            start_col: 0,
            end_row: 4,
            end_col: 1,
            columns: vec![AutoFilterColumn {
                field: 2,
                criteria: vec![">15".to_string()],
                either: false,
            }],
        }),
    );
    let saved = editor.save().expect("the workbook saves");

    let filter = filter_of(&parse_xlsx(&saved).expect("parses")).expect("still filtered");
    assert_eq!(filter.columns.len(), 1);
    assert_eq!(filter.columns[0].field, 2);
    assert_eq!(filter.columns[0].criteria, vec![">15".to_string()]);
}

#[test]
fn a_filter_can_be_taken_away() {
    let mut editor = XlsxEditor::new(FIXTURE).expect("the fixture opens");
    editor.set_auto_filter(0, None);
    let saved = editor.save().expect("the workbook saves");
    assert!(filter_of(&parse_xlsx(&saved).expect("parses")).is_none());
}

/// Two criteria on one column come back as they went in.
#[test]
fn two_criteria_survive_the_trip() {
    let mut editor = XlsxEditor::new(FIXTURE).expect("the fixture opens");
    editor.set_auto_filter(
        0,
        Some(AutoFilter {
            start_row: 1,
            start_col: 0,
            end_row: 4,
            end_col: 1,
            columns: vec![AutoFilterColumn {
                field: 2,
                criteria: vec![">=10".to_string(), "<=20".to_string()],
                either: false,
            }],
        }),
    );
    let saved = editor.save().expect("the workbook saves");

    let filter = filter_of(&parse_xlsx(&saved).expect("parses")).expect("still filtered");
    assert_eq!(
        filter.columns[0].criteria,
        vec![">=10".to_string(), "<=20".to_string()]
    );
    assert!(!filter.columns[0].either);
}

#[test]
fn an_unchanged_workbook_writes_no_filter() {
    let workbook = parse_xlsx(FIXTURE).expect("parses");
    let mut editor = XlsxEditor::new(FIXTURE).expect("opens");
    editor.apply_workbook(&workbook).expect("nothing to write");
    assert!(!editor.has_edits());
}

