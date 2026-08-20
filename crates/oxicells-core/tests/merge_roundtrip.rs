// This Source Code Form is subject to the terms of the Mozilla Public
// License, v. 2.0. If a copy of the MPL was not distributed with this
// file, You can obtain one at https://mozilla.org/MPL/2.0/.

//! What survives a round trip through the editor, for the cells a sheet merges.
//!
//! The fixture was written by Excel 16: five rows of data, with B1:C1 and
//! A4:A5 merged.

use oxicells_core::editor::XlsxEditor;
use oxicells_core::ir::MergeCell;
use oxicells_core::parser::parse_xlsx;

const FIXTURE: &[u8] = include_bytes!("fixtures/merged_cells.xlsx");

fn merges(workbook: &oxicells_core::ir::Workbook) -> Vec<(u32, u32, u32, u32)> {
    let mut listed: Vec<_> = workbook.sheets[0]
        .merge_cells
        .iter()
        .map(|merge| (merge.start_row, merge.start_col, merge.end_row, merge.end_col))
        .collect();
    listed.sort_unstable();
    listed
}

#[test]
fn reading_a_workbook_finds_the_cells_it_merges() {
    let workbook = parse_xlsx(FIXTURE).expect("the fixture parses");
    // B1:C1 and A4:A5, as rows are one-based and columns zero-based.
    assert_eq!(merges(&workbook), vec![(1, 1, 1, 2), (4, 0, 5, 0)]);
}

#[test]
fn the_editor_can_replace_what_a_sheet_merges() {
    let mut editor = XlsxEditor::new(FIXTURE).expect("the fixture opens");
    editor.set_merges(
        0,
        vec![MergeCell {
            start_row: 2,
            start_col: 1,
            end_row: 2,
            end_col: 2,
        }],
    );
    let saved = editor.save().expect("the workbook saves");

    let workbook = parse_xlsx(&saved).expect("the saved workbook parses");
    assert_eq!(merges(&workbook), vec![(2, 1, 2, 2)]);
}

#[test]
fn the_editor_can_take_every_merge_away() {
    let mut editor = XlsxEditor::new(FIXTURE).expect("the fixture opens");
    editor.set_merges(0, Vec::new());
    let saved = editor.save().expect("the workbook saves");

    let workbook = parse_xlsx(&saved).expect("the saved workbook parses");
    assert!(merges(&workbook).is_empty());
}

/// The whole way round, as a VBA run would do it.
#[test]
fn a_changed_workbook_saves_the_merges_it_changed() {
    let mut workbook = parse_xlsx(FIXTURE).expect("the fixture parses");
    workbook.sheets[0].merge_cells.push(MergeCell {
        start_row: 3,
        start_col: 1,
        end_row: 3,
        end_col: 2,
    });

    let mut editor = XlsxEditor::new(FIXTURE).expect("the fixture opens");
    editor
        .apply_workbook(&workbook)
        .expect("the change is one the editor can write");
    let saved = editor.save().expect("the workbook saves");

    let reread = parse_xlsx(&saved).expect("the saved workbook parses");
    assert_eq!(
        merges(&reread),
        vec![(1, 1, 1, 2), (3, 1, 3, 2), (4, 0, 5, 0)]
    );
}

/// Listing the same merges in another order is not a change.
#[test]
fn reordering_the_merges_writes_nothing_back() {
    let mut workbook = parse_xlsx(FIXTURE).expect("the fixture parses");
    workbook.sheets[0].merge_cells.reverse();

    let mut editor = XlsxEditor::new(FIXTURE).expect("the fixture opens");
    editor.apply_workbook(&workbook).expect("nothing to write");
    assert!(!editor.has_edits());
}

