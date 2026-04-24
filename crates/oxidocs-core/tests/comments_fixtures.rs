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

    // P-02 (partial): commentRangeStart/End are preserved on runs so the
    // renderer can locate highlight boundaries after layout.
    let mut found_range_start = false;
    let mut found_range_end = false;
    for page in &doc.pages {
        for block in &page.blocks {
            if let oxidocs_core::ir::Block::Paragraph(p) = block {
                for run in &p.runs {
                    if !run.comment_range_start.is_empty() {
                        found_range_start = true;
                    }
                    if !run.comment_range_end.is_empty() {
                        found_range_end = true;
                    }
                }
            }
        }
    }
    assert!(
        found_range_start,
        "commentRangeStart marker must survive to a run"
    );
    assert!(
        found_range_end,
        "commentRangeEnd marker must survive to a run"
    );
}
