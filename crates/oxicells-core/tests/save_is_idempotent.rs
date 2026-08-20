// This Source Code Form is subject to the terms of the Mozilla Public
// License, v. 2.0. If a copy of the MPL was not distributed with this
// file, You can obtain one at https://mozilla.org/MPL/2.0/.

//! Saving a workbook nobody edited must give back the workbook that was read.
//!
//! This is the path VBA takes: a macro opens a workbook, changes some of it,
//! and the editor writes the result. If the write drifts when nothing changed,
//! every macro that touches one cell quietly rewrites the rest of the sheet.

use oxicells_core::editor::XlsxEditor;
use oxicells_core::parser::parse_xlsx;

fn survives_a_save(fixture: &[u8], name: &str) {
    let opened = parse_xlsx(fixture).expect("the fixture parses");

    let mut editor = XlsxEditor::new(fixture).expect("the fixture opens");
    editor
        .apply_workbook(&opened)
        .expect("a workbook applies to the editor it came from");
    let saved = editor.save().expect("the workbook saves");
    let reopened = parse_xlsx(&saved).expect("the saved workbook parses");

    assert_eq!(
        opened.sheets.len(),
        reopened.sheets.len(),
        "{name}: the sheet count changed"
    );
    for (before, after) in opened.sheets.iter().zip(&reopened.sheets) {
        assert_eq!(before.name, after.name, "{name}: a sheet was renamed");
        assert_eq!(
            format!("{:?}", before.rows),
            format!("{:?}", after.rows),
            "{name}: the cells of {} changed",
            before.name
        );
        assert_eq!(
            format!("{:?}", before.merge_cells),
            format!("{:?}", after.merge_cells),
            "{name}: the merges of {} changed",
            before.name
        );
        assert_eq!(
            format!("{:?}", before.auto_filter),
            format!("{:?}", after.auto_filter),
            "{name}: the filter of {} changed",
            before.name
        );
        assert_eq!(
            before.hidden_cols, after.hidden_cols,
            "{name}: the hidden columns of {} changed",
            before.name
        );
    }
}

#[test]
fn hidden_rows_and_columns_survive_a_save() {
    survives_a_save(include_bytes!("fixtures/hidden_rows_cols.xlsx"), "hidden");
}

#[test]
fn merges_survive_a_save() {
    survives_a_save(include_bytes!("fixtures/merged_cells.xlsx"), "merged");
}

#[test]
fn a_filter_survives_a_save() {
    survives_a_save(include_bytes!("fixtures/filtered.xlsx"), "filtered");
}

#[test]
fn what_the_editor_cannot_keep_still_leaves_the_rest_alone() {
    survives_a_save(include_bytes!("fixtures/unsupported_bits.xlsx"), "unsupported");
}
