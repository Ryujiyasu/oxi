// This Source Code Form is subject to the terms of the Mozilla Public
// License, v. 2.0. If a copy of the MPL was not distributed with this
// file, You can obtain one at https://mozilla.org/MPL/2.0/.

//! Freezing and unfreezing a sheet's panes.
//!
//! The pane lives inside `<sheetView>`, which nearly every sheet already has —
//! an unfrozen one still says `<sheetViews><sheetView workbookViewId="0"/>`.
//! So the work is replacing that element's `<pane>` while leaving the view's
//! own attributes alone: a sheet someone had turned the gridlines off on must
//! not come back with them on.
//!
//! The fixture's four sheets are frozen at the top row, at the first column,
//! at C4, and not at all.

use oxicells_core::{parse_xlsx, XlsxEditor};

const FIXTURE: &[u8] = include_bytes!("fixtures/frozen.xlsx");

fn frozen(bytes: &[u8]) -> Vec<(u32, u32)> {
    parse_xlsx(bytes)
        .expect("parses")
        .sheets
        .iter()
        .map(|sheet| (sheet.frozen_rows, sheet.frozen_cols))
        .collect()
}

#[test]
fn a_sheet_that_was_not_frozen_can_be() {
    let mut editor = XlsxEditor::new(FIXTURE).expect("opens");
    editor.set_frozen_panes(3, 2, 0);
    let written = editor.save().expect("saves");
    assert_eq!(frozen(&written)[3], (2, 0));
}

#[test]
fn a_freeze_can_be_moved() {
    let mut editor = XlsxEditor::new(FIXTURE).expect("opens");
    editor.set_frozen_panes(2, 1, 1);
    let written = editor.save().expect("saves");
    assert_eq!(frozen(&written)[2], (1, 1), "was three rows and two columns");
}

#[test]
fn a_freeze_can_be_taken_away() {
    let mut editor = XlsxEditor::new(FIXTURE).expect("opens");
    editor.set_frozen_panes(0, 0, 0);
    let written = editor.save().expect("saves");
    assert_eq!(frozen(&written)[0], (0, 0), "was the top row");
}

#[test]
fn the_other_sheets_are_left_alone() {
    let mut editor = XlsxEditor::new(FIXTURE).expect("opens");
    editor.set_frozen_panes(1, 4, 0);
    let written = editor.save().expect("saves");
    let after = frozen(&written);
    assert_eq!(after[0], (1, 0));
    assert_eq!(after[1], (4, 0), "the one that was changed");
    assert_eq!(after[2], (3, 2));
    assert_eq!(after[3], (0, 0));
}

#[test]
fn a_column_only_freeze_writes_only_the_split_it_has() {
    let mut editor = XlsxEditor::new(FIXTURE).expect("opens");
    editor.set_frozen_panes(0, 0, 2);
    let written = editor.save().expect("saves");
    assert_eq!(frozen(&written)[0], (0, 2));
}

#[test]
fn the_views_own_attributes_come_through() {
    // A sheet with the gridlines turned off has that on the sheetView, beside
    // the pane. Replacing the pane must not replace the view.
    let mut editor = XlsxEditor::new(FIXTURE).expect("opens");
    editor.set_frozen_panes(2, 2, 2);
    let written = editor.save().expect("saves");
    let mut archive = zip::ZipArchive::new(std::io::Cursor::new(&written)).expect("a zip");
    let mut xml = String::new();
    for at in 0..archive.len() {
        let mut part = archive.by_index(at).expect("a part");
        if part.name().ends_with("sheet3.xml") {
            use std::io::Read;
            part.read_to_string(&mut xml).expect("reads");
        }
    }
    assert!(xml.contains(r#"tabSelected="1""#), "the view kept its own say: {xml:.400}");
    assert!(xml.contains(r#"workbookViewId="0""#), "{xml:.400}");
    assert!(xml.contains(r#"state="frozen""#), "{xml:.400}");
    assert_eq!(xml.matches("<pane ").count(), 1, "exactly one pane: {xml:.400}");
}

#[test]
fn a_freeze_changed_in_the_ir_reaches_the_file() {
    // The browser edits an IR rather than calling the editor, so the change
    // has to be noticed by the diff as well as writable by hand.
    let mut workbook = parse_xlsx(FIXTURE).expect("parses");
    workbook.sheets[3].frozen_rows = 2;
    workbook.sheets[3].frozen_cols = 1;
    workbook.sheets[0].frozen_rows = 0;
    let mut editor = XlsxEditor::new(FIXTURE).expect("opens");
    editor.apply_workbook(&workbook).expect("applies");
    let written = editor.save().expect("saves");
    let after = frozen(&written);
    assert_eq!(after[3], (2, 1), "a sheet given a freeze");
    assert_eq!(after[0], (0, 0), "and one whose freeze was taken away");
    assert_eq!(after[2], (3, 2), "with the others left alone");
}
