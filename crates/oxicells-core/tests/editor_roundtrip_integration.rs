//! Integration tests for XlsxEditor round-trip preservation.
//!
//! Parallel to oxidocs-core's editor_roundtrip_integration (S378). Pins the
//! CLAUDE.md "round-trip preservation (open->save->reopen)" metric for the
//! Excel editor, and documents the behavioral DIFFERENCE from DocxEditor:
//! XlsxEditor returns the ORIGINAL bytes on a no-edit save (byte-identical),
//! whereas DocxEditor re-serializes.

use oxicells_core::ir::CellValue;
use oxicells_core::parser::parse_xlsx;
use oxicells_core::XlsxEditor;
use oxicells_core::editor::CellEdit;

const XLSX: &[u8] = include_bytes!("../../../tests/fixtures/basic_test.xlsx");
const MULTI: &[u8] = include_bytes!("../../../tests/fixtures/multi_sheet.xlsx");

#[test]
fn no_edit_save_is_byte_identical() {
    // XlsxEditor returns original_data.clone() when there are no edits —
    // byte-identical (DISTINCT from DocxEditor which re-serializes, S378).
    let editor = XlsxEditor::new(XLSX).expect("open");
    assert!(!editor.has_edits());
    let saved = editor.save().expect("save");
    assert_eq!(saved, XLSX, "no-edit XlsxEditor save must equal input bytes");
}

#[test]
fn has_edits_flag_transitions() {
    let mut editor = XlsxEditor::new(XLSX).expect("open");
    assert!(!editor.has_edits());
    editor.set_cell(0, 2, 0, "edited".to_string());
    assert!(editor.has_edits());
}

#[test]
fn set_cell_defers_until_save() {
    // set_cell records the edit but does NOT mutate the parsed workbook IR
    // (S350 contract). workbook() still reflects the original until save.
    let mut editor = XlsxEditor::new(XLSX).expect("open");
    let before = editor.workbook().sheets[0].rows.len();
    editor.set_cell(0, 2, 0, "deferred".to_string());
    let after = editor.workbook().sheets[0].rows.len();
    assert_eq!(before, after, "set_cell must not restructure the parsed IR");
}

#[test]
fn edit_survives_save_and_reparse() {
    // Set a cell, save, re-parse, and confirm the new value is present.
    let mut editor = XlsxEditor::new(XLSX).expect("open");
    // Row 2 (1-based), col 0 — a data cell in basic_test.
    editor.set_cell(0, 2, 0, "RoundTripValue".to_string());
    let saved = editor.save().expect("save");
    let wb = parse_xlsx(&saved).expect("reparse");
    let found = wb.sheets[0].rows.iter().any(|r| {
        r.cells.iter().any(|c| {
            matches!(&c.value, CellValue::String(s) if s == "RoundTripValue")
        })
    });
    assert!(found, "edited cell value must survive save + re-parse");
}

#[test]
fn save_is_deterministic_with_edit() {
    let mut e1 = XlsxEditor::new(XLSX).expect("open");
    e1.set_cell(0, 2, 0, "DET".to_string());
    let a = e1.save().expect("save1");
    let mut e2 = XlsxEditor::new(XLSX).expect("open");
    e2.set_cell(0, 2, 0, "DET".to_string());
    let b = e2.save().expect("save2");
    assert_eq!(a, b, "same edit -> identical bytes");
}

#[test]
fn apply_edits_batch_equals_individual() {
    let mut e_batch = XlsxEditor::new(XLSX).expect("open");
    e_batch.apply_edits(&[CellEdit {
        sheet_index: 0,
        row: 2,
        col: 0,
        new_value: "BATCH".to_string(),
    }]);
    let batch_saved = e_batch.save().expect("save batch");

    let mut e_indiv = XlsxEditor::new(XLSX).expect("open");
    e_indiv.set_cell(0, 2, 0, "BATCH".to_string());
    let indiv_saved = e_indiv.save().expect("save indiv");

    assert_eq!(batch_saved, indiv_saved, "apply_edits == individual set_cell");
}

#[test]
fn multi_sheet_edit_isolation() {
    // Editing sheet 0 must not corrupt sheet 1's content.
    let orig = parse_xlsx(MULTI).expect("parse multi");
    assert!(orig.sheets.len() >= 2, "multi_sheet has >= 2 sheets");
    let s1_texts: Vec<String> = orig.sheets[1]
        .rows
        .iter()
        .flat_map(|r| r.cells.iter().map(|c| c.value.display()))
        .collect();

    let mut editor = XlsxEditor::new(MULTI).expect("open");
    editor.set_cell(0, 1, 0, "Sheet0Edit".to_string());
    let saved = editor.save().expect("save");
    let wb = parse_xlsx(&saved).expect("reparse");

    assert_eq!(wb.sheets.len(), orig.sheets.len(), "sheet count preserved");
    let s1_after: Vec<String> = wb.sheets[1]
        .rows
        .iter()
        .flat_map(|r| r.cells.iter().map(|c| c.value.display()))
        .collect();
    assert_eq!(s1_texts, s1_after, "sheet 1 must be untouched by a sheet-0 edit");
}

#[test]
fn out_of_range_sheet_edit_is_safe() {
    // Editing a non-existent sheet index must not panic on save.
    let mut editor = XlsxEditor::new(XLSX).expect("open");
    editor.set_cell(99, 1, 0, "ghost".to_string());
    let saved = editor.save().expect("save must not panic on out-of-range sheet");
    let wb = parse_xlsx(&saved).expect("reparse");
    assert_eq!(wb.sheets.len(), 1, "no sheet added by ghost edit");
}
