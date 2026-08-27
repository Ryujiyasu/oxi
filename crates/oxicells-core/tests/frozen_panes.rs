// This Source Code Form is subject to the terms of the Mozilla Public
// License, v. 2.0. If a copy of the MPL was not distributed with this
// file, You can obtain one at https://mozilla.org/MPL/2.0/.

//! Reading a sheet's frozen panes.
//!
//! A workbook says `<pane ySplit="1" topLeftCell="A2" state="frozen"/>` to mean
//! "hold the top row in view while the rest scrolls". It is the commonest thing
//! a real sheet does, and the parser used to walk past it and list it under what
//! it could not show.
//!
//! The fixture was written by Excel 16 and holds four sheets, frozen four ways:
//! at A2 (the top row), at B1 (the first column), at C4 (three rows and two
//! columns), and not at all. What Excel wrote for each:
//!
//! ```text
//! <pane ySplit="1" topLeftCell="A2" activePane="bottomLeft" state="frozen"/>
//! <pane xSplit="1" topLeftCell="B1" activePane="topRight" state="frozen"/>
//! <pane xSplit="2" ySplit="3" topLeftCell="C4" activePane="bottomRight" state="frozen"/>
//! (no pane element at all)
//! ```

use oxicells_core::parse_xlsx;

const FIXTURE: &[u8] = include_bytes!("fixtures/frozen.xlsx");

#[test]
fn the_splits_are_counted_in_cells() {
    let workbook = parse_xlsx(FIXTURE).expect("the fixture parses");
    let frozen: Vec<(u32, u32)> = workbook
        .sheets
        .iter()
        .map(|sheet| (sheet.frozen_rows, sheet.frozen_cols))
        .collect();
    assert_eq!(
        frozen,
        vec![(1, 0), (0, 1), (3, 2), (0, 0)],
        "the top row, the first column, three rows and two columns, and none",
    );
}

#[test]
fn a_frozen_pane_is_no_longer_listed_as_something_it_cannot_show() {
    let workbook = parse_xlsx(FIXTURE).expect("the fixture parses");
    for sheet in &workbook.sheets {
        assert!(
            !sheet
                .unsupported_elements
                .iter()
                .any(|one| one.contains("pane") || one.contains("Frozen")),
            "{} still reports {:?}",
            sheet.name,
            sheet.unsupported_elements,
        );
    }
}

#[test]
fn a_sheet_that_freezes_nothing_says_nothing() {
    let workbook = parse_xlsx(FIXTURE).expect("the fixture parses");
    let last = workbook.sheets.last().expect("four sheets");
    assert_eq!((last.frozen_rows, last.frozen_cols), (0, 0));
}

#[test]
fn a_freeze_survives_being_written_back() {
    // The pane lives in a part the editor does not touch, so it rides along in
    // the original XML. That is worth pinning: a workbook whose frozen header
    // came loose after an edit would be a change nobody asked for.
    let mut editor = oxicells_core::XlsxEditor::new(FIXTURE).expect("opens");
    editor.set_cell_value(
        0,
        1,
        0,
        oxicells_core::editor::CellEditValue::String("changed".to_string()),
    );
    let written = editor.save().expect("saves");
    let after = parse_xlsx(&written).expect("the written file parses");
    let frozen: Vec<(u32, u32)> = after
        .sheets
        .iter()
        .map(|sheet| (sheet.frozen_rows, sheet.frozen_cols))
        .collect();
    assert_eq!(frozen, vec![(1, 0), (0, 1), (3, 2), (0, 0)]);
}
