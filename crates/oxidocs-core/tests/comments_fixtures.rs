//! End-to-end parse of the comments/tracked-changes fixture set.
//!
//! These tests run the full `parse_docx` pipeline against the 10 fixtures in
//! `tools/fixtures/comments_samples/`. Coverage corresponds to attack-matrix
//! rows P-01 (comments.xml body → Comment IR) and P-03/P-04 (tracked changes).
//!
//! COM-validated ground truth: see
//! `docs/spec/comments_tracked_changes/com_measurements/comments_tracked_changes_com.json`.

use std::path::{Path, PathBuf};

fn fixture(name: &str) -> PathBuf {
    // Tests run from the crate root (crates/oxidocs-core); go up two levels.
    let manifest_dir = Path::new(env!("CARGO_MANIFEST_DIR"));
    manifest_dir.join("../../tools/fixtures/comments_samples").join(name)
}

fn read_fixture(name: &str) -> Option<Vec<u8>> {
    std::fs::read(fixture(name)).ok()
}

/// P-01: Comment body is parsed with author + initials + date.
///
/// Word COM (2026-04-18): Comments.Count=1, Author="Alice Reviewer",
/// Initial="AR", Scope.Text="brown fox".
#[test]
fn fixture_01_comment_fields_roundtrip() {
    let Some(bytes) = read_fixture("fixture_01_single_comment.docx") else {
        eprintln!("skipping: fixture_01 missing");
        return;
    };
    let doc = oxidocs_core::parse_docx(&bytes).expect("parse fixture_01");

    // P-01: the comment body surfaces on Document.comments with all metadata.
    assert_eq!(doc.comments.len(), 1, "expected exactly one comment");
    let c = &doc.comments[0];
    assert_eq!(c.id, "0");
    assert_eq!(c.author.as_deref(), Some("Alice Reviewer"));
    assert_eq!(c.initials.as_deref(), Some("AR"));
    assert_eq!(c.date.as_deref(), Some("2026-04-18T10:00:00Z"));

    // P-02: commentRangeStart/End AND commentReference are preserved on runs so
    // the renderer can locate highlight boundaries + balloon anchor after layout.
    let mut found_range_start = false;
    let mut found_range_end = false;
    let mut found_reference = false;
    for page in &doc.pages {
        for block in &page.blocks {
            if let oxidocs_core::ir::Block::Paragraph(p) = block {
                for run in &p.runs {
                    if run.comment_range_start.iter().any(|id| id == "0") {
                        found_range_start = true;
                    }
                    if run.comment_range_end.iter().any(|id| id == "0") {
                        found_range_end = true;
                    }
                    if run.comment_references.iter().any(|id| id == "0") {
                        found_reference = true;
                    }
                }
            }
        }
    }
    assert!(found_range_start, "commentRangeStart id=0 must survive to a run");
    assert!(found_range_end, "commentRangeEnd id=0 must survive to a run");
    assert!(found_reference, "commentReference id=0 must survive to a run");
}

/// P-10: comments_extended.xml merges onto Comment (reply + resolved fields).
///
/// Even though fixture_02 fails `Documents.Open` in Word (validator rejects it
/// for a still-unidentified schema defect), the XML is syntactically valid and
/// the Oxi parser must still extract the reply pointer and resolved flag, so
/// that when the fixture is repaired the renderer work needs no adjustments.
#[test]
fn fixture_02_comments_extended_reply_threaded() {
    let Some(bytes) = read_fixture("fixture_02_comment_with_reply.docx") else {
        eprintln!("skipping: fixture_02 missing");
        return;
    };
    let doc = oxidocs_core::parse_docx(&bytes).expect("parse fixture_02");
    assert_eq!(doc.comments.len(), 2, "expected 2 comments");

    let by_id: std::collections::HashMap<_, _> =
        doc.comments.iter().map(|c| (c.id.as_str(), c)).collect();
    let parent = by_id.get("0").expect("parent comment id=0");
    let reply = by_id.get("1").expect("reply comment id=1");

    // Parent paragraph id is captured from the body's first w:p@w14:paraId.
    assert_eq!(parent.para_id.as_deref(), Some("00000010"));
    assert!(parent.parent_para_id.is_none(), "parent has no grandparent");
    assert!(!parent.resolved);

    // Reply points back at parent via parent_para_id.
    assert_eq!(reply.para_id.as_deref(), Some("00000011"));
    assert_eq!(reply.parent_para_id.as_deref(), Some("00000010"));
    assert!(!reply.resolved);
}

#[test]
fn fixture_03_comments_extended_resolved_flag() {
    let Some(bytes) = read_fixture("fixture_03_resolved_comment.docx") else {
        eprintln!("skipping: fixture_03 missing");
        return;
    };
    let doc = oxidocs_core::parse_docx(&bytes).expect("parse fixture_03");
    assert_eq!(doc.comments.len(), 1);
    let c = &doc.comments[0];
    assert_eq!(c.id, "0");
    assert_eq!(c.para_id.as_deref(), Some("00000010"));
    assert!(c.resolved, "w15:done='1' must land on Comment.resolved");
}
