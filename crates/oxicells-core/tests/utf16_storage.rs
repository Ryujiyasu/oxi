// This Source Code Form is subject to the terms of the Mozilla Public
// License, v. 2.0. If a copy of the MPL was not distributed with this
// file, You can obtain one at https://mozilla.org/MPL/2.0/.
use oxicells_core::{
    editor::CellEditValue, ir::CellValue, parse_xlsx_preserving_values, XlsxEditor,
};
use std::io::{Cursor, Read, Write};
const BASE: &[u8] = include_bytes!("../../../tests/fixtures/basic_test.xlsx");

#[test]
fn calculation_language_context_does_not_change_saved_package() {
    let mut workbook = parse_xlsx_preserving_values(BASE).unwrap();
    workbook.calculation.dbcs_language = Some(false);
    let mut editor = XlsxEditor::new(BASE).unwrap();
    editor.apply_workbook(&workbook).unwrap();
    let saved = editor.save().unwrap();
    assert_eq!(saved, BASE);
    let reopened = parse_xlsx_preserving_values(&saved).unwrap();
    assert_eq!(reopened.calculation.dbcs_language, None);
}

#[test]
fn formulas_preserve_literal_ooxml_escape_sequences_on_save() {
    let mut editor = XlsxEditor::new(BASE).unwrap();
    editor.set_cell_value(0, 1, 0, CellEditValue::Formula("=\"literal _x0041_\"".into()));
    let saved = editor.save().unwrap();
    let reopened = parse_xlsx_preserving_values(&saved).unwrap();
    assert_eq!(reopened.sheets[0].rows[0].cells[0].formula.as_deref(), Some("\"literal _x0041_\""));
    let mut zip = zip::ZipArchive::new(Cursor::new(saved)).unwrap();
    let mut sheet = String::new();
    zip.by_name("xl/worksheets/sheet1.xml").unwrap().read_to_string(&mut sheet).unwrap();
    assert!(sheet.contains("literal _x005F_x0041_"));
}

#[test]
fn formula_utf16_survives_read_calculation_edit_and_save() {
    let sheet = r#"<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData><row r="1"><c r="A1" t="str"><f>"a_xD83D_b"</f><v>a_xD83D_b</v></c><c r="B1" t="str"><f>FORMULATEXT(A1)</f><v>="a_xD83D_b"</v></c></row></sheetData></worksheet>"#;
    let original = rewrite(&[("xl/worksheets/sheet1.xml", sheet)]);
    let mut book = parse_xlsx_preserving_values(&original).unwrap();
    let formula = vec![34, 97, 0xD83D, 98, 34];
    assert_eq!(book.sheets[0].rows[0].cells[0].formula_utf16.as_ref(), Some(&formula));
    oxicells_core::formula::evaluate_workbook_formulas(&mut book);
    assert_eq!(book.sheets[0].rows[0].cells[0].value.string_units().unwrap().as_ref(), [97, 0xD83D, 98]);
    assert_eq!(book.sheets[0].rows[0].cells[1].value.string_units().unwrap().as_ref(), [61, 34, 97, 0xD83D, 98, 34]);
    let changed = vec![34, 97, 0xDE00, 98, 34];
    let cell = &mut book.sheets[0].rows[0].cells[0];
    cell.formula = Some(String::from_utf16_lossy(&changed));
    cell.formula_utf16 = Some(changed.clone());
    oxicells_core::formula::evaluate_workbook_formulas(&mut book);
    let mut editor = XlsxEditor::new(&original).unwrap();
    editor.apply_workbook(&book).unwrap();
    let saved = editor.save().unwrap();
    let reopened = parse_xlsx_preserving_values(&saved).unwrap();
    let cells = &reopened.sheets[0].rows[0].cells;
    assert_eq!(cells[0].formula_utf16.as_ref(), Some(&changed));
    assert_eq!(cells[0].value.string_units().unwrap().as_ref(), [97, 0xDE00, 98]);
    assert_eq!(cells[1].value.string_units().unwrap().as_ref(), [61, 34, 97, 0xDE00, 98, 34]);
}

#[test]
fn replacing_formula_does_not_reuse_stale_utf16_source() {
    let sheet = r#"<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData><row r="1"><c r="A1" t="str"><f>"a_xD83D_b"</f><v>a_xD83D_b</v></c></row></sheetData></worksheet>"#;
    let original = rewrite(&[("xl/worksheets/sheet1.xml", sheet)]);
    let mut book = parse_xlsx_preserving_values(&original).unwrap();
    book.sheets[0].rows[0].cells[0].formula = Some("1+2".into());
    oxicells_core::formula::evaluate_workbook_formulas(&mut book);
    assert!(matches!(book.sheets[0].rows[0].cells[0].value, CellValue::Number(3.0)));
    let mut editor = XlsxEditor::new(&original).unwrap();
    editor.apply_workbook(&book).unwrap();
    let reopened = parse_xlsx_preserving_values(&editor.save().unwrap()).unwrap();
    assert_eq!(reopened.sheets[0].rows[0].cells[0].formula.as_deref(), Some("1+2"));
    assert_eq!(reopened.sheets[0].rows[0].cells[0].formula_utf16, None);
}

#[test]
fn formula_utf16_row_and_column_insertions_preserve_literal_units() {
    let sheet = r#"<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData><row r="2"><c r="B2" t="str"><f>A1&amp;"a_xD83D_b"</f><v>a_xD83D_b</v></c></row></sheetData></worksheet>"#;
    let original = rewrite(&[("xl/worksheets/sheet1.xml", sheet)]);
    let mut book = parse_xlsx_preserving_values(&original).unwrap();
    let name = book.sheets[0].name.clone();
    oxicells_core::bands::insert(&mut book, &name, oxicells_core::bands::Band::Rows, 1, 1);
    oxicells_core::bands::insert(&mut book, &name, oxicells_core::bands::Band::Columns, 0, 1);
    let cell = &mut book.sheets[0].rows.iter_mut().find(|row| row.index == 3).unwrap().cells[0];
    let expected: Vec<u16> = "B2&\"a".encode_utf16().chain([0xD83D, 98, 34]).collect();
    assert_eq!(cell.formula_units(), Some(expected.as_slice()));
    assert!(cell.rewrite_formula(|_| Err("cannot rewrite".into())).is_err());
    assert_eq!(cell.formula_units(), Some(expected.as_slice()));
    let mut editor = XlsxEditor::new(&original).unwrap();
    editor.apply_workbook(&book).unwrap();
    let saved = editor.save().unwrap();
    let reopened = parse_xlsx_preserving_values(&saved).unwrap();
    let cell = &reopened.sheets[0].rows.iter().find(|row| row.index == 3).unwrap().cells[0];
    assert_eq!(cell.col, 2);
    assert_eq!(cell.formula_units(), Some(expected.as_slice()));
}

fn rewrite(parts: &[(&str, &str)]) -> Vec<u8> {
    let mut source = zip::ZipArchive::new(Cursor::new(BASE)).unwrap();
    let mut out = zip::ZipWriter::new(Cursor::new(Vec::new()));
    for at in 0..source.len() {
        let mut part = source.by_index(at).unwrap();
        let name = part.name().to_owned();
        let mut bytes = Vec::new();
        part.read_to_end(&mut bytes).unwrap();
        if let Some((_, text)) = parts.iter().find(|(path, _)| *path == name) {
            bytes = text.as_bytes().to_vec();
        }
        out.start_file(name, zip::write::SimpleFileOptions::default())
            .unwrap();
        out.write_all(&bytes).unwrap();
    }
    out.finish().unwrap().into_inner()
}

#[test]
fn shared_and_inline_strings_decode_surrogates_without_replacing_or_unescaping_twice() {
    let sst = r#"<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><si><t>_xD83D_</t></si><si><t>_xDE00_</t></si><si><r><t>_xD83D_</t></r><r><t>_xDE00_</t></r></si><si><t>_x005F_xD83D_</t></si></sst>"#;
    let sheet = r#"<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData><row r="1"><c r="A1" t="s"><v>0</v></c><c r="B1" t="s"><v>1</v></c><c r="C1" t="s"><v>2</v></c><c r="D1" t="s"><v>3</v></c><c r="E1" t="inlineStr"><is><t>_xD83D_</t></is></c><c r="F1" t="str"><v>_xDE00_</v></c></row></sheetData></worksheet>"#;
    let bytes = rewrite(&[
        ("xl/sharedStrings.xml", sst),
        ("xl/worksheets/sheet1.xml", sheet),
    ]);
    let book = parse_xlsx_preserving_values(&bytes).unwrap();
    let expected = [
        vec![0xD83D],
        vec![0xDE00],
        vec![0xD83D, 0xDE00],
        "_xD83D_".encode_utf16().collect(),
        vec![0xD83D],
        vec![0xDE00],
    ];
    for (cell, units) in book.sheets[0].rows[0].cells.iter().zip(expected) {
        assert_eq!(cell.value.string_units().unwrap().as_ref(), units);
    }
}

#[test]
fn edited_and_cached_strings_survive_xml_save_and_reopen_with_original_units() {
    let samples = [
        vec![0xD83D],
        vec![0xDE00],
        vec![0xD83D, 0xDE00],
        vec![65, 0, 0xD83D, 13, 10, 0xFFFF],
        "_xD83D_ _x005F_ <&>".encode_utf16().collect(),
    ];
    let mut editor = XlsxEditor::new(BASE).unwrap();
    for (at, units) in samples.iter().enumerate() {
        editor.set_cell_value(0, 1, at as u32, CellEditValue::Utf16(units.clone()));
        editor.set_cell_value(
            0,
            2,
            at as u32,
            CellEditValue::CachedFormula {
                formula: format!("{}1", (b'A' + at as u8) as char),
                reach: None,
                value: CellValue::from_utf16(units.clone()),
            },
        );
    }
    let saved = editor.save().unwrap();
    let book = parse_xlsx_preserving_values(&saved).unwrap();
    for row in book.sheets[0].rows.iter().take(2) {
        for (at, units) in samples.iter().enumerate() {
            let cell = row.cells.iter().find(|cell| cell.col == at as u32).unwrap();
            assert_eq!(cell.value.string_units().unwrap().as_ref(), units);
        }
    }
    let mut next = book.clone();
    next.sheets[0].rows[0].cells[0].value = CellValue::Utf16(vec![0xDE00, 0xD83D]);
    let mut editor = XlsxEditor::new(&saved).unwrap();
    editor.apply_workbook(&next).unwrap();
    let reopened = parse_xlsx_preserving_values(&editor.save().unwrap()).unwrap();
    assert_eq!(
        reopened.sheets[0].rows[0].cells[0]
            .value
            .string_units()
            .unwrap()
            .as_ref(),
        [0xDE00, 0xD83D]
    );
}

#[test]
fn direct_formula_references_retain_utf16_values() {
    let mut editor = XlsxEditor::new(BASE).unwrap();
    editor.set_cell_value(0, 1, 0, CellEditValue::Utf16(vec![0xD83D]));
    editor.set_cell_value(0, 1, 1, CellEditValue::Formula("A1".into()));
    let mut book = oxicells_core::parse_xlsx(&editor.save().unwrap()).unwrap();
    oxicells_core::formula::evaluate_workbook_formulas(&mut book);
    let cell = &book.sheets[0].rows[0].cells[1];
    assert_eq!(cell.value.string_units().unwrap().as_ref(), [0xD83D]);
}

#[test]
fn workbook_compatibility_version_controls_recalculation_after_save_and_reopen() {
    let mut editor = XlsxEditor::new(BASE).unwrap();
    editor.set_cell_value(0, 1, 0, CellEditValue::Utf16(vec![0xD83D, 0xDE00]));
    editor.set_cell_value(0, 1, 1, CellEditValue::Formula("LEN(A1)".into()));
    let original = editor.save().unwrap();
    for version in [2, 1] {
        let mut book = parse_xlsx_preserving_values(&original).unwrap();
        book.calculation.compatibility_version = version;
        let mut editor = XlsxEditor::new(&original).unwrap();editor.apply_workbook(&book).unwrap();
        let saved = editor.save().unwrap();let mut reopened = parse_xlsx_preserving_values(&saved).unwrap();
        assert_eq!(reopened.calculation.compatibility_version, version);
        oxicells_core::formula::evaluate_workbook_formulas(&mut reopened);
        assert!(matches!(reopened.sheets[0].rows[0].cells[1].value, CellValue::Number(n) if n == if version==2 {1.0}else{2.0}));
    }
}

#[test]
fn rich_text_save_keeps_value_units_run_fonts_and_format_only_edits() {
    use oxicells_core::ir::TextRun;
    let mut book=parse_xlsx_preserving_values(BASE).unwrap();
    let cell=&mut book.sheets[0].rows[0].cells[0];
    cell.value=CellValue::Utf16(vec![0xD83D,65,13,10,66]);
    cell.runs=vec![TextRun{text:"\u{FFFD}A\n".into(),dressed:true,bold:true,color:Some("#FF0000".into()),..Default::default()},TextRun{text:"B".into(),dressed:true,italic:true,..Default::default()}];
    let mut editor=XlsxEditor::new(BASE).unwrap();editor.apply_workbook(&book).unwrap();
    let saved=editor.save().unwrap();let mut reopened=parse_xlsx_preserving_values(&saved).unwrap();
    let cell=&mut reopened.sheets[0].rows[0].cells[0];
    assert_eq!(cell.value.string_units().unwrap().as_ref(),[0xD83D,65,13,10,66]);
    assert_eq!(cell.runs.len(),2);assert!(cell.runs[0].bold);assert_eq!(cell.runs[0].color.as_deref(),Some("FF0000"));assert!(cell.runs[1].italic);
    cell.runs[0].bold=false;
    let mut editor=XlsxEditor::new(&saved).unwrap();editor.apply_workbook(&reopened).unwrap();
    let updated=parse_xlsx_preserving_values(&editor.save().unwrap()).unwrap();
    assert!(!updated.sheets[0].rows[0].cells[0].runs[0].bold);
    assert_eq!(updated.sheets[0].rows[0].cells[0].value.string_units().unwrap().as_ref(),[0xD83D,65,13,10,66]);
}

#[test]
fn literal_escape_split_across_runs_survives_format_edit_and_save() {
    let sheet = r#"<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData><row r="1"><c r="A1" t="inlineStr"><is><r><rPr><b/></rPr><t>_x0</t></r><r><rPr><i/></rPr><t>041_</t></r></is></c></row></sheetData></worksheet>"#;
    let original = rewrite(&[("xl/worksheets/sheet1.xml", sheet)]);
    let mut book = parse_xlsx_preserving_values(&original).unwrap();
    let cell = &mut book.sheets[0].rows[0].cells[0];
    assert_eq!(cell.value.string_units().unwrap().as_ref(), "_x0041_".encode_utf16().collect::<Vec<_>>());
    assert!(cell.runs[0].bold);
    assert!(cell.runs[1].italic);
    cell.runs[1].bold = true;
    let mut editor = XlsxEditor::new(&original).unwrap();
    editor.apply_workbook(&book).unwrap();
    let saved = editor.save().unwrap();
    let reopened = parse_xlsx_preserving_values(&saved).unwrap();
    let cell = &reopened.sheets[0].rows[0].cells[0];
    assert_eq!(cell.value.string_units().unwrap().as_ref(), "_x0041_".encode_utf16().collect::<Vec<_>>());
    assert!(cell.runs[0].bold);
    assert!(cell.runs[1].bold && cell.runs[1].italic);
}

#[test]
fn cells_and_row_formats_emit_each_inserted_row_once_in_numeric_order() {
    for data in [
        "<sheetData/>",
        "<sheetData><row r=\"1\"><c r=\"A1\"><v>1</v></c></row><row r=\"10\"/></sheetData>",
    ] {
        let xml = format!("<worksheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\">{data}</worksheet>");
        let source = rewrite(&[("xl/worksheets/sheet1.xml", &xml)]);
        let mut editor = XlsxEditor::new(&source).unwrap();
        editor.set_cell_value(0, 6, 0, CellEditValue::Utf16(vec![0xD83D]));
        editor.set_row_height(0, 6, Some(25.0), true);
        editor.set_row_hidden(0, 3, true);
        editor.set_row_height(0, 10, Some(45.0), true);
        editor.set_cell_value(0, 12, 0, CellEditValue::String("last".into()));
        editor.set_row_height(0, 12, Some(30.0), true);
        let saved = editor.save().unwrap();
        let book = parse_xlsx_preserving_values(&saved).unwrap();
        let rows = &book.sheets[0].rows;
        assert!(
            rows.windows(2).all(|pair| pair[0].index < pair[1].index),
            "duplicate or unsorted rows: {:?}",
            rows.iter().map(|row| row.index).collect::<Vec<_>>()
        );
        assert_eq!(rows.iter().filter(|row| row.index == 6).count(), 1);
        assert_eq!(
            rows.iter().find(|row| row.index == 6).unwrap().height,
            Some(25.0)
        );
        assert_eq!(
            rows.iter().find(|row| row.index == 12).unwrap().height,
            Some(30.0)
        );
        assert_eq!(
            rows.iter().find(|row| row.index == 10).unwrap().height,
            Some(45.0)
        );
    }
}
