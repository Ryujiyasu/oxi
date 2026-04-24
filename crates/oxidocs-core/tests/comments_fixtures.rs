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

fn collect_runs(doc: &oxidocs_core::Document) -> Vec<&oxidocs_core::ir::Run> {
    let mut runs = Vec::new();
    for page in &doc.pages {
        for block in &page.blocks {
            if let oxidocs_core::ir::Block::Paragraph(p) = block {
                for run in &p.runs {
                    runs.push(run);
                }
            }
        }
    }
    runs
}

/// P-03: `<w:ins>` runs carry `tracked_change.change_type == "insert"` with
/// author + date preserved. Deleted text from `<w:delText>` is preserved (not
/// dropped) per attack-matrix P-04 notes.
#[test]
fn fixture_05_single_ins() {
    let Some(bytes) = read_fixture("fixture_05_single_ins.docx") else {
        eprintln!("skipping: fixture_05 missing");
        return;
    };
    let doc = oxidocs_core::parse_docx(&bytes).expect("parse fixture_05");
    let ins_runs: Vec<_> = collect_runs(&doc)
        .into_iter()
        .filter(|r| r.tracked_change.as_ref().map(|t| t.change_type.as_str()) == Some("insert"))
        .collect();
    assert_eq!(ins_runs.len(), 1, "one <w:ins> run expected");
    let run = ins_runs[0];
    assert_eq!(run.text, "INSERTED TEXT ");
    let tc = run.tracked_change.as_ref().unwrap();
    assert!(tc.author.is_some(), "w:author must survive");
    assert!(tc.date.is_some(), "w:date must survive");
}

/// P-04: `<w:del>` runs carry `tracked_change.change_type == "delete"` and the
/// deleted text (from `<w:delText>`) is preserved verbatim.
#[test]
fn fixture_06_single_del() {
    let Some(bytes) = read_fixture("fixture_06_single_del.docx") else {
        eprintln!("skipping: fixture_06 missing");
        return;
    };
    let doc = oxidocs_core::parse_docx(&bytes).expect("parse fixture_06");
    let del_runs: Vec<_> = collect_runs(&doc)
        .into_iter()
        .filter(|r| r.tracked_change.as_ref().map(|t| t.change_type.as_str()) == Some("delete"))
        .collect();
    assert_eq!(del_runs.len(), 1);
    assert_eq!(del_runs[0].text, "DELETED TEXT ");
}

/// P-03+P-04: mixed ins/del in one paragraph preserves XML order.
#[test]
fn fixture_07_mixed_ins_del() {
    let Some(bytes) = read_fixture("fixture_07_mixed_ins_del.docx") else {
        eprintln!("skipping: fixture_07 missing");
        return;
    };
    let doc = oxidocs_core::parse_docx(&bytes).expect("parse fixture_07");
    let revisions: Vec<_> = collect_runs(&doc)
        .into_iter()
        .filter_map(|r| {
            r.tracked_change
                .as_ref()
                .map(|t| (t.change_type.clone(), r.text.clone()))
        })
        .collect();
    assert_eq!(
        revisions,
        vec![
            ("insert".to_string(), "ins1 ".to_string()),
            ("delete".to_string(), "del1 ".to_string()),
            ("insert".to_string(), "ins2".to_string()),
        ],
        "three revisions must preserve authoring (XML) order"
    );
}

/// P-05: `<w:moveFrom>` and `<w:moveTo>` wrap runs the same way ins/del do.
/// Both sides share `w:id`, which becomes `tracked_change.pair_id` so the
/// renderer can draw move arrows between the two halves.
#[test]
fn fixture_08_move_from_to_pair() {
    let Some(bytes) = read_fixture("fixture_08_move_from_to.docx") else {
        eprintln!("skipping: fixture_08 missing");
        return;
    };
    let doc = oxidocs_core::parse_docx(&bytes).expect("parse fixture_08");
    let moves: Vec<_> = collect_runs(&doc)
        .into_iter()
        .filter_map(|r| {
            r.tracked_change.as_ref().and_then(|t| match t.change_type.as_str() {
                "moveFrom" | "moveTo" => Some((t.change_type.clone(), t.pair_id.clone(), r.text.clone())),
                _ => None,
            })
        })
        .collect();
    assert_eq!(moves.len(), 2, "one moveFrom + one moveTo expected");
    // Both sides carry the same text "moved clause".
    for (_, _, text) in &moves {
        assert_eq!(text, "moved clause");
    }
    let kinds: Vec<_> = moves.iter().map(|(k, _, _)| k.as_str()).collect();
    assert!(kinds.contains(&"moveFrom"));
    assert!(kinds.contains(&"moveTo"));
    // Note: `<w:moveFrom>` / `<w:moveTo>` wrappers each carry their *own*
    // `w:id`; the actual from↔to pairing lives on the surrounding
    // `moveFromRangeStart` / `moveToRangeStart` pair via `w:name`
    // (revisions_notes.md §1.2). For Phase 2 parser we preserve the wrapper
    // id per-run; R-11 walks the range markers to draw the arrow.
    let from_id = moves.iter().find(|(k, _, _)| k == "moveFrom").and_then(|(_, id, _)| id.clone());
    let to_id = moves.iter().find(|(k, _, _)| k == "moveTo").and_then(|(_, id, _)| id.clone());
    assert!(from_id.is_some(), "moveFrom wrapper w:id must be captured");
    assert!(to_id.is_some(), "moveTo wrapper w:id must be captured");
}

/// P-12: people.xml populates Document.people with two reviewers.
#[test]
fn fixture_10_people_two_reviewers() {
    let Some(bytes) = read_fixture("fixture_10_multiple_reviewers.docx") else {
        eprintln!("skipping: fixture_10 missing");
        return;
    };
    let doc = oxidocs_core::parse_docx(&bytes).expect("parse fixture_10");
    // Word COM: fixture fails Documents.Open for a validator defect, but the
    // underlying people.xml is syntactically valid and our parser must surface
    // both reviewers so R-02 (author palette) has the list it needs.
    assert_eq!(doc.people.len(), 2, "expected exactly two reviewers");
    let authors: Vec<_> = doc.people.iter().map(|p| p.author.as_str()).collect();
    assert_eq!(authors, vec!["Alice Reviewer", "Bob Reviewer"]);
    // presenceInfo merges onto each Person.
    assert_eq!(doc.people[0].user_id.as_deref(), Some("Alice Reviewer"));
    assert_eq!(doc.people[1].user_id.as_deref(), Some("Bob Reviewer"));
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
