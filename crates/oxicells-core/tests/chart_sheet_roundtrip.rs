// This Source Code Form is subject to the terms of the Mozilla Public
// License, v. 2.0. If a copy of the MPL was not distributed with this
// file, You can obtain one at https://mozilla.org/MPL/2.0/.

use oxicells_core::{ir::SheetKind, parse_xlsx, XlsxEditor};

#[test]
fn native_chart_sheet_identity_and_relationships_survive_reordering() {
    let original = include_bytes!("../../../tests/fixtures/vba/chart_document_save_refresh.xlsm");
    let excel_saved = include_bytes!("../../../tests/fixtures/vba/chart_document_save_refresh_refreshed.xlsm");
    let mut edited = parse_xlsx(original).unwrap();
    assert_eq!(edited.sheets[0].kind, SheetKind::Chart);
    assert_eq!(edited.sheets[1].kind, SheetKind::Worksheet);
    assert_eq!(edited.sheets[0].drawings.len(), 1, "empty native charts remain objects");
    edited.sheets.swap(0, 1);
    let mut editor = XlsxEditor::new(original).unwrap();
    editor.apply_workbook(&edited).unwrap();
    let saved = editor.save().unwrap();
    let actual = parse_xlsx(&saved).unwrap();
    let expected = parse_xlsx(excel_saved).unwrap();
    let identities = |book: &oxicells_core::ir::Workbook| {
        book.sheets.iter().map(|sheet| (sheet.name.clone(), sheet.kind)).collect::<Vec<_>>()
    };
    assert_eq!(identities(&actual), identities(&expected));
    assert_eq!(actual.sheets[1].drawings.len(), 1);
}
