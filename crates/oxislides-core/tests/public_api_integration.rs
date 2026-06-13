// This Source Code Form is subject to the terms of the Mozilla Public
// License, v. 2.0. If a copy of the MPL was not distributed with this
// file, You can obtain one at https://mozilla.org/MPL/2.0/.

//! Integration tests for oxislides-core public API.
//!
//! First integration test file for this crate. Pins parser entry,
//! PptxEditor round-trip semantics, and error contracts so silent
//! regressions in the PowerPoint engine surface in CI.
//!
//! Why these specific cases:
//!
//! - `parse_pptx` on garbage/empty input must return Err (not panic).
//!   WASM/CLI hosts rely on this contract for user-uploaded files.
//! - `PptxEditor::save()` with NO edits must return original bytes
//!   verbatim — pins the "edit-free save = identity" invariant that
//!   downstream re-parsers (and the WASM diff view) depend on.
//! - `set_run_text` followed by `presentation()` must NOT mutate the
//!   parsed IR — edits are deferred until `save()` builds a new ZIP.
//! - `has_edits()` flips false→true on first edit and STAYS true even
//!   if the edit is reverted to original text (no idempotency check).
//! - Multi-slide editing across slides preserves slide ordering and
//!   does not corrupt unedited slides.
//! - Out-of-range edit coordinates (slide_index past end) must be a
//!   safe no-op — pins WASM caller contract.

use oxislides_core::ir::ShapeContent;
use oxislides_core::parser::{parse_pptx, PptxError};
use oxislides_core::{PptxEditor};

// ────────────────────────────────────────────────────────────────────
// parse_pptx error paths
// ────────────────────────────────────────────────────────────────────

#[test]
fn parse_pptx_garbage_returns_err() {
    let result = parse_pptx(b"not a pptx zip");
    assert!(result.is_err(), "garbage input must return Err, not panic");
}

#[test]
fn parse_pptx_empty_returns_err() {
    let result = parse_pptx(b"");
    assert!(result.is_err());
}

#[test]
fn parse_pptx_truncated_zip_returns_err() {
    // Build a fake ZIP header + truncated body. Parser must error,
    // not panic on partial archives (e.g., interrupted upload).
    let truncated = b"PK\x03\x04truncated rest of zip";
    let result = parse_pptx(truncated);
    assert!(result.is_err());
}

#[test]
fn parse_pptx_error_types_are_displayable() {
    // PptxError::Display via thiserror must produce a non-empty message.
    // Pins the contract for log/UI surfaces that print error strings.
    let err = parse_pptx(b"").unwrap_err();
    let s = format!("{}", err);
    assert!(!s.is_empty(), "error Display must be non-empty");
}

// ────────────────────────────────────────────────────────────────────
// parse_pptx happy path (basic + multi-slide)
// ────────────────────────────────────────────────────────────────────

#[test]
fn parse_pptx_basic_slide_dimensions() {
    let data = include_bytes!("../../../tests/fixtures/basic_test.pptx");
    let pres = parse_pptx(data).expect("basic_test.pptx must parse");
    assert_eq!(pres.slides.len(), 1);
    // 10in x 7.5in = 720pt x 540pt at 72 DPI
    assert!((pres.slide_width - 720.0).abs() < 1.0, "width ~720pt");
    assert!((pres.slide_height - 540.0).abs() < 1.0, "height ~540pt");
}

#[test]
fn parse_pptx_multi_slide_count_and_order() {
    let data = include_bytes!("../../../tests/fixtures/multi_slide.pptx");
    let pres = parse_pptx(data).expect("multi_slide.pptx must parse");
    assert_eq!(pres.slides.len(), 3);
    // Slide.index is 1-based (passed as i+1 by parser.rs:628) — pin this
    // convention so callers using `slide.index` for display/citation
    // don't break silently on 0-based vs 1-based change.
    for (i, slide) in pres.slides.iter().enumerate() {
        assert_eq!(slide.index, i + 1, "slide index is 1-based");
    }
}

// ────────────────────────────────────────────────────────────────────
// PptxEditor round-trip semantics
// ────────────────────────────────────────────────────────────────────

#[test]
fn pptx_editor_save_without_edits_returns_original_bytes() {
    // Pins the invariant: PptxEditor::save() with no edits = identity.
    // Downstream tools (WASM diff view, comparison) rely on this.
    let data = include_bytes!("../../../tests/fixtures/basic_test.pptx");
    let editor = PptxEditor::new(data).expect("editor must construct");
    let saved = editor.save().expect("save with no edits must succeed");
    assert_eq!(saved, data, "save() with no edits must be byte-identical");
}

#[test]
fn pptx_editor_set_run_text_defers_until_save() {
    // set_run_text must NOT mutate the in-memory presentation IR; edits
    // are deferred until save() rebuilds the ZIP. presentation()
    // returns the original parsed IR even after set_run_text.
    let data = include_bytes!("../../../tests/fixtures/basic_test.pptx");
    let mut editor = PptxEditor::new(data).expect("editor must construct");

    // Capture original title
    let original_title: String;
    if let ShapeContent::TextBox { paragraphs } = &editor.presentation().slides[0].shapes[0].content {
        original_title = paragraphs[0].runs[0].text.clone();
    } else {
        panic!("expected TextBox in slide 0 shape 0");
    }

    editor.set_run_text(0, 0, 0, 0, "Edited Title".to_string());

    // presentation() still returns ORIGINAL title (edit not applied to IR)
    if let ShapeContent::TextBox { paragraphs } = &editor.presentation().slides[0].shapes[0].content {
        assert_eq!(
            paragraphs[0].runs[0].text, original_title,
            "set_run_text must NOT mutate the parsed IR (deferred until save)"
        );
    } else {
        panic!("expected TextBox");
    }
}

#[test]
fn pptx_editor_has_edits_flag() {
    let data = include_bytes!("../../../tests/fixtures/basic_test.pptx");
    let mut editor = PptxEditor::new(data).expect("editor must construct");
    assert!(!editor.has_edits(), "fresh editor must have no edits");
    editor.set_run_text(0, 0, 0, 0, "X".to_string());
    assert!(editor.has_edits(), "after set_run_text, has_edits must be true");
}

#[test]
fn pptx_editor_save_applies_edit_visible_in_reparse() {
    // After save() with an edit, re-parsing the saved bytes must surface
    // the new text at the edited coordinate.
    let data = include_bytes!("../../../tests/fixtures/basic_test.pptx");
    let mut editor = PptxEditor::new(data).expect("editor must construct");
    editor.set_run_text(0, 0, 0, 0, "Round-Trip Title".to_string());

    let saved = editor.save().expect("save must succeed");
    let pres = parse_pptx(&saved).expect("re-parse must succeed");
    if let ShapeContent::TextBox { paragraphs } = &pres.slides[0].shapes[0].content {
        assert_eq!(
            paragraphs[0].runs[0].text, "Round-Trip Title",
            "edited text must survive save + re-parse"
        );
    } else {
        panic!("expected TextBox");
    }
}

#[test]
fn pptx_editor_multi_slide_independent_edits() {
    // Edit slide 0 only, leave slides 1 and 2 untouched. After save +
    // re-parse, slide 0 has new text and other slides preserve original
    // text. Pins isolation: cross-slide edit contamination would silently
    // corrupt unedited slides.
    let data = include_bytes!("../../../tests/fixtures/multi_slide.pptx");
    let pres_orig = parse_pptx(data).expect("orig parse");

    // Capture slide 1 + 2 first paragraph texts
    let s1_orig: String = match &pres_orig.slides[1].shapes[0].content {
        ShapeContent::TextBox { paragraphs } => {
            paragraphs[0].runs.iter().map(|r| r.text.as_str()).collect()
        }
        _ => String::new(),
    };
    let s2_orig: String = match &pres_orig.slides[2].shapes[0].content {
        ShapeContent::TextBox { paragraphs } => {
            paragraphs[0].runs.iter().map(|r| r.text.as_str()).collect()
        }
        _ => String::new(),
    };

    let mut editor = PptxEditor::new(data).expect("editor must construct");
    editor.set_run_text(0, 0, 0, 0, "Slide0Edit".to_string());
    let saved = editor.save().expect("save");
    let pres = parse_pptx(&saved).expect("re-parse");

    assert_eq!(pres.slides.len(), 3, "slide count preserved");

    // Slide 1 unchanged
    if let ShapeContent::TextBox { paragraphs } = &pres.slides[1].shapes[0].content {
        let s1_new: String = paragraphs[0].runs.iter().map(|r| r.text.as_str()).collect();
        assert_eq!(s1_new, s1_orig, "slide 1 must not change");
    }

    // Slide 2 unchanged
    if let ShapeContent::TextBox { paragraphs } = &pres.slides[2].shapes[0].content {
        let s2_new: String = paragraphs[0].runs.iter().map(|r| r.text.as_str()).collect();
        assert_eq!(s2_new, s2_orig, "slide 2 must not change");
    }
}

#[test]
fn pptx_editor_out_of_range_edit_is_safe_no_op() {
    // set_run_text with slide_index past end must not panic. Save will
    // skip the unresolvable coordinate. Pins safety for UI/WASM callers
    // that might pass stale indices.
    let data = include_bytes!("../../../tests/fixtures/basic_test.pptx");
    let mut editor = PptxEditor::new(data).expect("editor must construct");
    // Slide 99 does not exist (only slide 0)
    editor.set_run_text(99, 0, 0, 0, "Ghost".to_string());

    // has_edits is still true (the edit was recorded even though no slide matches)
    assert!(editor.has_edits(), "set_run_text records the edit unconditionally");

    // save() must not panic on unresolvable slide_index
    let saved = editor.save().expect("save must not panic on out-of-range slide");
    let pres = parse_pptx(&saved).expect("re-parse");
    assert_eq!(pres.slides.len(), 1, "no slides added by ghost edit");
}

#[test]
fn pptx_editor_apply_edits_batch() {
    // apply_edits is the batch interface to set_run_text. After
    // applying a batch, has_edits must be true and save must persist
    // each individual edit.
    use oxislides_core::editor::SlideTextEdit;
    let data = include_bytes!("../../../tests/fixtures/multi_slide.pptx");
    let mut editor = PptxEditor::new(data).expect("editor must construct");

    let batch = vec![
        SlideTextEdit {
            slide_index: 0,
            shape_index: 0,
            paragraph_index: 0,
            run_index: 0,
            new_text: "BATCH_S0".to_string(),
        },
        SlideTextEdit {
            slide_index: 1,
            shape_index: 0,
            paragraph_index: 0,
            run_index: 0,
            new_text: "BATCH_S1".to_string(),
        },
    ];
    editor.apply_edits(&batch);
    assert!(editor.has_edits());

    let saved = editor.save().expect("save");
    let pres = parse_pptx(&saved).expect("re-parse");
    // Slide 0 has BATCH_S0
    if let ShapeContent::TextBox { paragraphs } = &pres.slides[0].shapes[0].content {
        assert_eq!(paragraphs[0].runs[0].text, "BATCH_S0");
    }
    // Slide 1 has BATCH_S1
    if let ShapeContent::TextBox { paragraphs } = &pres.slides[1].shapes[0].content {
        assert_eq!(paragraphs[0].runs[0].text, "BATCH_S1");
    }
}

#[test]
fn pptx_editor_new_propagates_parse_error() {
    // PptxEditor::new on garbage bytes must propagate the parse error
    // (not silently swallow it).
    let result = PptxEditor::new(b"not a pptx");
    let err = match result {
        Err(e) => e,
        Ok(_) => panic!("expected Err, got Ok"),
    };
    // Error must be Display-renderable
    let s = format!("{}", err);
    assert!(!s.is_empty());
}

#[test]
fn pptx_error_display_variants_non_empty() {
    // Each PptxError variant carries useful info in its Display.
    let invalid = PptxError::InvalidData("test message".to_string());
    let s = format!("{}", invalid);
    assert!(s.contains("test message"), "InvalidData Display must include the payload");
}
