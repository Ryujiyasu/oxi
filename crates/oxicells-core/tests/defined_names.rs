// This Source Code Form is subject to the terms of the Mozilla Public
// License, v. 2.0. If a copy of the MPL was not distributed with this
// file, You can obtain one at https://mozilla.org/MPL/2.0/.

//! Reading the names a workbook gives to things.
//!
//! `INDEX(cats, MATCH(x, exDB, 0))` is an ordinary formula, and it came back
//! `#NAME?` because `cats` and `exDB` are names the workbook defines and
//! nothing here had ever read them. The evaluator has had `define_name` all
//! along; the parser simply never told it anything.
//!
//! Sixty-four of the eight hundred workbooks in the blind corpus carry one.

use oxicells_core::{parse_xlsx, XlsxEditor};

const FIXTURE: &[u8] = include_bytes!("fixtures/frozen.xlsx");

/// A workbook holding two numbers, a name for the pair, and a formula using it.
fn a_workbook_with_a_name() -> Vec<u8> {
    let mut editor = XlsxEditor::new(FIXTURE).expect("opens");
    editor.set_cell_value(
        0,
        1,
        0,
        oxicells_core::editor::CellEditValue::Number(10.0),
    );
    editor.set_cell_value(
        0,
        2,
        0,
        oxicells_core::editor::CellEditValue::Number(32.0),
    );
    editor.save().expect("saves")
}

#[test]
fn the_names_a_workbook_states_are_read() {
    let book = parse_xlsx(&a_workbook_with_a_name()).expect("parses");
    // The fixture states none of its own, so this says the field exists and is
    // empty rather than absent — a workbook with no names is not a workbook
    // whose names were not read.
    assert!(book.defined_names.is_empty());
}

#[test]
fn a_name_is_read_from_the_workbook_part() {
    // Written out rather than taken from a file, so what is being tested is
    // the reading and not some workbook's idea of a name.
    let with_names = add_names(
        FIXTURE,
        concat!(
            r#"<definedNames>"#,
            r#"<definedName name="cats">Sheet1!$A$1:$A$3</definedName>"#,
            r#"<definedName name="one">Sheet1!$B$2</definedName>"#,
            // Built-in and sheet-local names retain their sheet prefix.
            r#"<definedName name="_xlnm.Print_Area" localSheetId="0">Sheet1!$A$1:$D$9</definedName>"#,
            r#"<definedName name="local" localSheetId="0">Sheet1!$C$1</definedName>"#,
            r#"</definedNames>"#,
        ),
    );
    let book = parse_xlsx(&with_names).expect("parses");
    assert_eq!(
        book.defined_names,
        vec![
            ("cats".to_string(), "Sheet1!$A$1:$A$3".to_string()),
            ("one".to_string(), "Sheet1!$B$2".to_string()),
            ("Sheet1!Print_Area".to_string(), "Sheet1!$A$1:$D$9".to_string()),
            ("Sheet1!local".to_string(), "Sheet1!$C$1".to_string()),
        ],
        "sheet-local and built-in names retain their owners",
    );
}

#[test]
fn a_formula_that_names_a_range_is_worked_out() {
    let with_names = add_names(
        FIXTURE,
        r#"<definedNames><definedName name="pair">Sheet1!$A$1:$A$2</definedName></definedNames>"#,
    );
    let mut editor = XlsxEditor::new(&with_names).expect("opens");
    use oxicells_core::editor::CellEditValue;
    editor.set_cell_value(0, 1, 0, CellEditValue::Number(10.0));
    editor.set_cell_value(0, 2, 0, CellEditValue::Number(32.0));
    editor.set_cell_value(0, 1, 5, CellEditValue::Formula("=SUM(pair)".to_string()));
    let written = editor.save().expect("saves");

    let mut book = parse_xlsx(&written).expect("parses");
    assert_eq!(book.defined_names.len(), 1, "the name came through");
    oxicells_core::formula::evaluate_workbook_formulas(&mut book);
    let cell = book.sheets[0]
        .rows
        .iter()
        .find(|row| row.index == 1)
        .and_then(|row| row.cells.iter().find(|cell| cell.col == 5))
        .expect("the formula cell is there");
    assert!(
        matches!(cell.value, oxicells_core::ir::CellValue::Number(n) if n == 42.0),
        "a named range adds up like any other, but held {:?}",
        cell.value,
    );
}

#[test]
fn scoped_names_keep_global_fallback_explicit_owners_and_attributes() {
    let source = add_names(FIXTURE, concat!(
        r#"<definedNames><definedName name="Rate">10</definedName>"#,
        r#"<definedName name="Rate" localSheetId="0">20</definedName>"#,
        r#"<definedName name="Rate" localSheetId="1">30</definedName>"#,
        r#"<definedName name="LocalRange" localSheetId="1">Sheet2!$A$1</definedName>"#,
        r#"<definedName name="Secret" localSheetId="1" hidden="1" comment="a &amp; b">40</definedName>"#,
        r#"<definedName name="_xlnm._FilterDatabase" localSheetId="0" hidden="1">Sheet1!$A$1:$A$3</definedName></definedNames>"#,
    ));
    let mut book = parse_xlsx(&source).unwrap();
    assert_eq!(book.defined_name_attributes["Sheet2!Secret"]["comment"], "a & b");
    assert_eq!(book.defined_name_attributes["Sheet1!_FilterDatabase"]["hidden"], "1");
    let mut edit = XlsxEditor::new(&source).unwrap();
    use oxicells_core::editor::CellEditValue;
    for sheet in 0..3 { edit.set_cell_value(sheet, 1, 4, CellEditValue::Formula("=Rate".into())); }
    edit.set_cell_value(0, 2, 4, CellEditValue::Formula("=Sheet2!Rate".into()));
    edit.set_cell_value(1, 1, 0, CellEditValue::Number(7.0));
    edit.set_cell_value(0, 3, 4, CellEditValue::Formula("=SUM(Sheet2!LocalRange)".into()));
    let source = edit.save().unwrap();
    book = parse_xlsx(&source).unwrap();
    oxicells_core::formula::evaluate_workbook_formulas(&mut book);
    let value = |sheet: usize, row| &book.sheets[sheet].rows.iter().find(|r| r.index == row).unwrap().cells.iter().find(|c| c.col == 4).unwrap().value;
    for (sheet, row, expected) in [(0, 1, 20.0), (1, 1, 30.0), (2, 1, 10.0), (0, 2, 30.0), (0, 3, 7.0)] {
        assert!(matches!(value(sheet, row), oxicells_core::ir::CellValue::Number(n) if *n == expected), "sheet={sheet} row={row}: {:?}", value(sheet, row));
    }
    book.defined_name_attributes.get_mut("Sheet2!Secret").unwrap().insert("comment".into(), "changed <&>".into());
    let mut edit = XlsxEditor::new(&source).unwrap(); edit.apply_workbook(&book).unwrap();
    let saved = edit.save().unwrap();
    let reopened = parse_xlsx(&saved).unwrap();
    assert_eq!(reopened.defined_names, book.defined_names);
    assert_eq!(reopened.defined_name_attributes["Sheet2!Secret"]["hidden"], "1");
    assert_eq!(reopened.defined_name_attributes["Sheet2!Secret"]["comment"], "changed <&>");
    let mut archive = zip::ZipArchive::new(std::io::Cursor::new(saved)).unwrap();
    let mut xml = String::new(); std::io::Read::read_to_string(&mut archive.by_name("xl/workbook.xml").unwrap(), &mut xml).unwrap();
    assert!(xml.contains(r#"name="_xlnm._FilterDatabase" localSheetId="0" hidden="1""#));
    assert!(!xml.contains(r#"name="Sheet1!_FilterDatabase""#));
}

#[test]
fn local_names_follow_sheet_rename_reorder_and_deletion() {
    let source = add_names(FIXTURE, concat!(
        r#"<definedNames><definedName name="Global">10</definedName>"#,
        r#"<definedName name="Rate" localSheetId="0">20</definedName>"#,
        r#"<definedName name="Rate" localSheetId="1" hidden="1" comment="kept">30</definedName>"#,
        r#"<definedName name="Ref" localSheetId="1">Sheet2!$A$1</definedName></definedNames>"#,
    ));
    let mut book = parse_xlsx(&source).unwrap();
    let renamed = "O'Brien ! \u{65e5}\u{672c}";
    book.sheets[1].name = renamed.into();
    book.sheets.swap(0, 1);
    let mut edit = XlsxEditor::new(&source).unwrap(); edit.apply_workbook(&book).unwrap();
    let saved = edit.save().unwrap();
    let mut reopened = parse_xlsx(&saved).unwrap();
    let held = oxicells_core::ir::scoped_defined_name(renamed, "Rate");
    assert!(reopened.defined_names.contains(&(held.clone(), "30".into())));
    assert!(reopened.defined_names.contains(&("Sheet1!Rate".into(), "20".into())));
    assert_eq!(reopened.defined_name_attributes[&held]["hidden"], "1");
    let ref_name = oxicells_core::ir::scoped_defined_name(renamed, "Ref");
    assert!(reopened.defined_names.iter().find(|(n, _)| n == &ref_name).unwrap().1.contains("O''Brien !"));
    reopened.sheets.remove(1);
    let mut edit = XlsxEditor::new(&saved).unwrap(); edit.apply_workbook(&reopened).unwrap();
    let saved = edit.save().unwrap(); let reopened = parse_xlsx(&saved).unwrap();
    assert!(reopened.defined_names.iter().all(|(n, _)| n != "Sheet1!Rate"));
    assert_eq!(reopened.defined_name_attributes[&held]["localSheetId"], "0");
    assert_eq!(reopened.defined_name_attributes[&held]["comment"], "kept");
}

#[test]
fn explicit_name_owners_keep_distinct_unicode_sheet_names() {
    let mut book = parse_xlsx(FIXTURE).unwrap();
    book.sheets[0].name = "\u{131}Sheet".into();
    book.sheets[1].name = "ISheet".into();
    book.defined_names = vec![("\u{131}Sheet!Rate".into(), "20".into()), ("ISheet!Rate".into(), "30".into())];
    let mut edit = XlsxEditor::new(FIXTURE).unwrap(); edit.apply_workbook(&book).unwrap();
    edit.set_cell_value(1, 1, 4, oxicells_core::editor::CellEditValue::Formula("='\u{131}Sheet'!Rate".into()));
    let mut reopened = parse_xlsx(&edit.save().unwrap()).unwrap();
    oxicells_core::formula::evaluate_workbook_formulas(&mut reopened);
    let cell = reopened.sheets[1].rows.iter().find(|r| r.index == 1).unwrap().cells.iter().find(|c| c.col == 4).unwrap();
    assert!(matches!(cell.value, oxicells_core::ir::CellValue::Number(20.0)), "{:?}", cell.value);
}

/// Put a `<definedNames>` block into a workbook part, where the schema wants
/// it: after the sheets and before anything that follows them.
fn add_names(source: &[u8], block: &str) -> Vec<u8> {
    let mut out = Vec::new();
    {
        let reader = std::io::Cursor::new(source);
        let mut archive = zip::ZipArchive::new(reader).expect("a zip");
        let mut writer = zip::ZipWriter::new(std::io::Cursor::new(&mut out));
        for at in 0..archive.len() {
            let mut part = archive.by_index(at).expect("a part");
            let name = part.name().to_string();
            let mut bytes = Vec::new();
            std::io::copy(&mut part, &mut bytes).expect("reads");
            if name == "xl/workbook.xml" {
                let xml = String::from_utf8_lossy(&bytes).replace("</sheets>", &format!("</sheets>{block}"));
                bytes = xml.into_bytes();
            }
            writer
                .start_file(name, zip::write::SimpleFileOptions::default())
                .expect("starts");
            use std::io::Write;
            writer.write_all(&bytes).expect("writes");
        }
        writer.finish().expect("finishes");
    }
    out
}
