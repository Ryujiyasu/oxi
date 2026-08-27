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
            // Excel's own, which is not what a formula means by a name.
            r#"<definedName name="_xlnm.Print_Area" localSheetId="0">Sheet1!$A$1:$D$9</definedName>"#,
            // Scoped to a single sheet: two sheets may each mean something
            // different by the same word, so it cannot go in a list held for
            // the whole workbook.
            r#"<definedName name="local" localSheetId="1">Sheet2!$C$1</definedName>"#,
            r#"</definedNames>"#,
        ),
    );
    let book = parse_xlsx(&with_names).expect("parses");
    assert_eq!(
        book.defined_names,
        vec![
            ("cats".to_string(), "Sheet1!$A$1:$A$3".to_string()),
            ("one".to_string(), "Sheet1!$B$2".to_string()),
        ],
        "the built-in and the sheet-scoped ones are left out",
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
