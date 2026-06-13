// This Source Code Form is subject to the terms of the Mozilla Public
// License, v. 2.0. If a copy of the MPL was not distributed with this
// file, You can obtain one at https://mozilla.org/MPL/2.0/.

//! Integration tests for DocxEditor round-trip PRESERVATION.
//!
//! CLAUDE.md lists "Round-trip preservation (open → save → reopen, IR
//! equality)" as an independent quality metric. The in-src editor tests
//! check specific edited fields; this suite pins the IR-EQUALITY and
//! determinism invariants the round-trip contract depends on (used by the
//! WASM editor + CLI save path):
//!   - save() with no edits preserves ALL document text (full IR text equality)
//!   - save() is deterministic (same bytes on repeat)
//!   - a single edit changes ONLY the target run, leaving all others intact
//!   - edits survive save → re-parse
//!   - apply_edits batch ≡ individual set_run_text

use oxidocs_core::editor::DocxEditor;
use oxidocs_core::ir::{Block, Document};
use oxidocs_core::parse_docx;

const DOCX: &[u8] = include_bytes!("../../../tests/fixtures/basic_test.docx");

/// Flatten every run's text across all pages/blocks (recursing into tables),
/// in document order — a stable fingerprint of the document's text content.
fn flatten_text(doc: &Document) -> Vec<String> {
    fn walk(blocks: &[Block], out: &mut Vec<String>) {
        for b in blocks {
            match b {
                Block::Paragraph(p) => {
                    for r in &p.runs {
                        out.push(r.text.clone());
                    }
                }
                Block::Table(t) => {
                    for row in &t.rows {
                        for cell in &row.cells {
                            walk(&cell.blocks, out);
                        }
                    }
                }
                _ => {}
            }
        }
    }
    let mut out = Vec::new();
    for page in &doc.pages {
        walk(&page.blocks, &mut out);
    }
    out
}

#[test]
fn save_no_edits_preserves_all_text() {
    let orig = parse_docx(DOCX).expect("parse orig");
    let editor = DocxEditor::new(DOCX).expect("open");
    assert!(!editor.has_edits());
    let saved = editor.save().expect("save");
    let reparsed = parse_docx(&saved).expect("reparse");
    assert_eq!(
        flatten_text(&orig),
        flatten_text(&reparsed),
        "save with no edits must preserve all document text (IR text equality)"
    );
}

#[test]
fn save_is_deterministic() {
    let editor = DocxEditor::new(DOCX).expect("open");
    let a = editor.save().expect("save 1");
    let b = editor.save().expect("save 2");
    assert_eq!(a, b, "save() must be deterministic (identical bytes on repeat)");
}

#[test]
fn no_edit_save_reserializes_but_preserves_ir() {
    // NOTE (pinned behavior): unlike PptxEditor/XlsxEditor (which return the
    // original bytes when there are no edits), DocxEditor ALWAYS re-serializes
    // the document XML on save(). So no-edit save is NOT byte-identical to the
    // input (ZIP metadata + minor XML serialization differ). The contract it
    // DOES guarantee is IR equality (text content preserved) + determinism,
    // both verified in the other tests. This test documents the re-serialize
    // behavior so a future "preserve original bytes on no-edit" optimization
    // is a conscious change, not a silent regression.
    let editor = DocxEditor::new(DOCX).expect("open");
    let saved = editor.save().expect("save");
    assert_ne!(saved, DOCX, "DocxEditor re-serializes (does not return input bytes)");
    // But the re-serialized output must still parse and preserve text.
    let reparsed = parse_docx(&saved).expect("reparse");
    let orig = parse_docx(DOCX).expect("parse orig");
    assert_eq!(flatten_text(&orig), flatten_text(&reparsed));
}

#[test]
fn single_edit_changes_only_target_run() {
    let orig_texts = flatten_text(&parse_docx(DOCX).expect("parse"));

    let mut editor = DocxEditor::new(DOCX).expect("open");
    editor.set_run_text(0, 0, "EDITED_HEADING".to_string());
    let saved = editor.save().expect("save");
    let reparsed = parse_docx(&saved).expect("reparse");
    let new_texts = flatten_text(&reparsed);

    // The edited run (page0 block0 run0) text differs; everything else equal.
    if let Block::Paragraph(p) = &reparsed.pages[0].blocks[0] {
        assert_eq!(p.runs[0].text, "EDITED_HEADING");
    } else {
        panic!("expected paragraph at page0 block0");
    }
    // Same number of runs (no structural change).
    assert_eq!(
        orig_texts.len(),
        new_texts.len(),
        "edit must not add/remove runs"
    );
    // Exactly ONE run text changed.
    let diffs = orig_texts
        .iter()
        .zip(new_texts.iter())
        .filter(|(a, b)| a != b)
        .count();
    assert_eq!(diffs, 1, "exactly one run text should change, got {diffs}");
}

#[test]
fn edit_survives_roundtrip() {
    let mut editor = DocxEditor::new(DOCX).expect("open");
    editor.set_run_text(0, 0, "ROUNDTRIP_TEXT".to_string());
    let saved = editor.save().expect("save");
    // Re-open the saved bytes in a NEW editor and save again — edit persists.
    let editor2 = DocxEditor::new(&saved).expect("reopen");
    let saved2 = editor2.save().expect("save2");
    let reparsed = parse_docx(&saved2).expect("reparse");
    if let Block::Paragraph(p) = &reparsed.pages[0].blocks[0] {
        assert_eq!(p.runs[0].text, "ROUNDTRIP_TEXT");
    } else {
        panic!("expected paragraph");
    }
}

#[test]
fn apply_edits_batch_equals_individual() {
    use oxidocs_core::editor::TextEdit;
    // Batch path
    let mut e_batch = DocxEditor::new(DOCX).expect("open");
    e_batch.apply_edits(&[TextEdit {
        paragraph_index: 0,
        run_index: 0,
        new_text: "BATCH".to_string(),
    }]);
    let batch_saved = e_batch.save().expect("save batch");

    // Individual path
    let mut e_indiv = DocxEditor::new(DOCX).expect("open");
    e_indiv.set_run_text(0, 0, "BATCH".to_string());
    let indiv_saved = e_indiv.save().expect("save indiv");

    assert_eq!(
        batch_saved, indiv_saved,
        "apply_edits batch must equal individual set_run_text"
    );
}

#[test]
fn has_edits_flag_transitions() {
    let mut editor = DocxEditor::new(DOCX).expect("open");
    assert!(!editor.has_edits(), "fresh editor: no edits");
    editor.set_run_text(0, 0, "X".to_string());
    assert!(editor.has_edits(), "after edit: has edits");
}
