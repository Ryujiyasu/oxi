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

/// P-06: `<w:rPrChange>` carries the prior rPr so the renderer can annotate
/// "formatting changed". Fixture 09 toggles a run from plain to bold.
#[test]
fn fixture_09_rpr_change_bold() {
    let Some(bytes) = read_fixture("fixture_09_rPrChange_bold.docx") else {
        eprintln!("skipping: fixture_09 missing");
        return;
    };
    let doc = oxidocs_core::parse_docx(&bytes).expect("parse fixture_09");
    let run = collect_runs(&doc)
        .into_iter()
        .find(|r| r.text == "Now bold (was plain).")
        .expect("bold run present");
    assert!(run.style.bold, "current state is bold");
    let pc = run.rpr_change.as_ref().expect("rpr_change must be populated");
    assert_eq!(pc.id.as_deref(), Some("300"));
    assert!(pc.author.is_some());
    assert!(pc.date.is_some());
    let prior = pc.prior_run_style.as_ref().expect("prior_run_style must be populated");
    assert!(!prior.bold, "prior state was plain (not bold)");
}

/// P-12: people.xml populates Document.people with two reviewers.
/// I-03: the same reviewers populate Document.authors with stable color_indices.
#[test]
fn fixture_10_people_two_reviewers() {
    let Some(bytes) = read_fixture("fixture_10_multiple_reviewers.docx") else {
        eprintln!("skipping: fixture_10 missing");
        return;
    };
    let doc = oxidocs_core::parse_docx(&bytes).expect("parse fixture_10");
    // P-12: people.xml content is preserved verbatim.
    assert_eq!(doc.people.len(), 2, "expected exactly two reviewers");
    let people_authors: Vec<_> = doc.people.iter().map(|p| p.author.as_str()).collect();
    assert_eq!(people_authors, vec!["Alice Reviewer", "Bob Reviewer"]);
    assert_eq!(doc.people[0].user_id.as_deref(), Some("Alice Reviewer"));
    assert_eq!(doc.people[1].user_id.as_deref(), Some("Bob Reviewer"));

    // I-03: authors palette is derived in first-seen order. people.xml seeds
    // it, so Alice gets color_index=0 and Bob gets color_index=1 even if Bob's
    // <w:del> appears in the document body before Alice's <w:ins>.
    assert_eq!(doc.authors.len(), 2, "expected exactly two palette entries");
    assert_eq!(doc.authors[0].display, "Alice Reviewer");
    assert_eq!(doc.authors[0].color_index, 0);
    assert_eq!(doc.authors[1].display, "Bob Reviewer");
    assert_eq!(doc.authors[1].color_index, 1);
}

/// I-03: when people.xml is absent, the palette falls back to first-seen
/// order across tracked changes. Fixture 05 has only Alice's <w:ins>.
#[test]
fn fixture_05_authors_palette_from_tracked_changes_only() {
    let Some(bytes) = read_fixture("fixture_05_single_ins.docx") else {
        eprintln!("skipping: fixture_05 missing");
        return;
    };
    let doc = oxidocs_core::parse_docx(&bytes).expect("parse fixture_05");
    assert!(doc.people.is_empty(), "fixture 05 has no people.xml");
    assert_eq!(doc.authors.len(), 1, "single insert author");
    assert_eq!(doc.authors[0].color_index, 0);
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

// ---------------------------------------------------------------------------
// R-01 / R-03 / R-11 — Layout-level integration tests.
//
// These run the full parse → layout pipeline and inspect emitted
// `LayoutContent::Text` properties. Ground truth (author RGB, hard-coded green
// for moves) is captured in
// `docs/spec/comments_tracked_changes/com_measurements/PIXEL_PASS_README.md`.
// ---------------------------------------------------------------------------

fn layout_doc(doc: &oxidocs_core::Document) -> oxidocs_core::layout::LayoutResult {
    let engine = oxidocs_core::layout::LayoutEngine::for_document(doc);
    engine.layout(doc)
}

fn collect_text_elements_with(
    res: &oxidocs_core::layout::LayoutResult,
    needle: &str,
) -> Vec<(bool, bool, Option<String>)> {
    let mut out = Vec::new();
    for page in &res.pages {
        for el in &page.elements {
            if let oxidocs_core::layout::LayoutContent::Text {
                text,
                underline,
                strikethrough,
                color,
                ..
            } = &el.content
            {
                if text.contains(needle) {
                    out.push((*underline, *strikethrough, color.clone()));
                }
            }
        }
    }
    out
}

/// R-01: a `<w:ins>` run lays out as underlined text in the author's palette
/// color. For Alice (palette slot 0), Word renders #D03337 (COM-confirmed
/// 2026-04-25 in fixture_05).
#[test]
fn fixture_05_layout_ins_underline_in_author_color() {
    let Some(bytes) = read_fixture("fixture_05_single_ins.docx") else {
        eprintln!("skipping: fixture_05 missing");
        return;
    };
    let doc = oxidocs_core::parse_docx(&bytes).expect("parse fixture_05");
    let result = layout_doc(&doc);

    // Layout splits text at word boundaries — match a single word from the
    // ins range to find at least one underlined element.
    let ins_hits = collect_text_elements_with(&result, "INSERTED");
    assert!(!ins_hits.is_empty(), "INSERTED must be laid out");
    for (underline, strike, color) in &ins_hits {
        assert!(*underline, "ins run must render underlined");
        assert!(!strike, "ins run must NOT render strikethrough");
        assert_eq!(
            color.as_deref(),
            Some("#D03337"),
            "Alice's ins must use palette slot 0 (#D03337)"
        );
    }

    // Adjacent normal runs MUST NOT be touched by the revision pre-pass.
    // "Before" comes from the leading non-revision run.
    let normal_hits = collect_text_elements_with(&result, "Before");
    assert!(!normal_hits.is_empty());
    for (underline, strike, _color) in &normal_hits {
        assert!(!*underline, "normal text must not be underlined");
        assert!(!*strike, "normal text must not be struck");
    }
}

/// R-03: a `<w:del>` run lays out with strikethrough in the author's palette
/// color. The text itself is preserved (Word's default "All Markup" view).
#[test]
fn fixture_06_layout_del_strikethrough_in_author_color() {
    let Some(bytes) = read_fixture("fixture_06_single_del.docx") else {
        eprintln!("skipping: fixture_06 missing");
        return;
    };
    let doc = oxidocs_core::parse_docx(&bytes).expect("parse fixture_06");
    let result = layout_doc(&doc);

    let del_hits = collect_text_elements_with(&result, "DELETED");
    assert!(!del_hits.is_empty(), "DELETED TEXT must be laid out");
    for (underline, strike, color) in &del_hits {
        assert!(!*underline, "del run must NOT be underlined");
        assert!(*strike, "del run must render strikethrough");
        assert_eq!(color.as_deref(), Some("#D03337"));
    }
}

/// R-02: two distinct authors get the first two palette slots. Alice (slot 0)
/// → #D03337, Bob (slot 1) → #5B2C90. Both are COM-confirmed against Word
/// 16.0 in fixture_10.
#[test]
fn fixture_10_layout_two_authors_get_distinct_colors() {
    let Some(bytes) = read_fixture("fixture_10_multiple_reviewers.docx") else {
        eprintln!("skipping: fixture_10 missing");
        return;
    };
    let doc = oxidocs_core::parse_docx(&bytes).expect("parse fixture_10");
    let result = layout_doc(&doc);

    let alice_ins = collect_text_elements_with(&result, "ALICE");
    assert!(!alice_ins.is_empty(), "Alice's ins missing from layout");
    for (underline, _, color) in &alice_ins {
        assert!(*underline);
        assert_eq!(color.as_deref(), Some("#D03337"), "Alice = palette slot 0");
    }

    let bob_del = collect_text_elements_with(&result, "REMOVE");
    assert!(!bob_del.is_empty(), "Bob's del missing from layout");
    for (_, strike, color) in &bob_del {
        assert!(*strike);
        assert_eq!(color.as_deref(), Some("#5B2C90"), "Bob = palette slot 1");
    }
}

/// S-03: `revisions::accept_all` permanently bakes accepted state into the
/// IR. After the call, `<w:ins>` runs are normal text and `<w:del>` runs are
/// gone. Idempotent: layout output identical to running with
/// `ShowRevisions::All` on the post-accept doc (which now has 0 tracked
/// changes).
#[test]
fn s03_accept_all_drops_deletions_and_clears_tracked_changes() {
    let Some(bytes) = read_fixture("fixture_07_mixed_ins_del.docx") else {
        eprintln!("skipping: fixture_07 missing");
        return;
    };
    let mut doc = oxidocs_core::parse_docx(&bytes).expect("parse fixture_07");

    // Sanity: pre-accept doc has 3 revisions.
    let pre_count = collect_runs(&doc)
        .iter()
        .filter(|r| r.tracked_change.is_some())
        .count();
    assert_eq!(pre_count, 3, "fixture_07 starts with 3 revisions");

    oxidocs_core::revisions::accept_all(&mut doc);

    let post_count = collect_runs(&doc)
        .iter()
        .filter(|r| r.tracked_change.is_some())
        .count();
    assert_eq!(post_count, 0, "after accept_all, no run carries tracked_change");

    let texts: Vec<String> = collect_runs(&doc).iter().map(|r| r.text.clone()).collect();
    let joined = texts.join("|");
    assert!(!joined.contains("del1"), "accepted del must be removed; got {joined}");
    assert!(joined.contains("ins1"), "accepted ins survives as normal text");
    assert!(joined.contains("ins2"), "accepted ins2 survives");
}

/// S-03: `reject_all` is the mirror of `accept_all` — keeps deletions,
/// drops insertions.
#[test]
fn s03_reject_all_drops_insertions_keeps_deletions() {
    let Some(bytes) = read_fixture("fixture_07_mixed_ins_del.docx") else {
        eprintln!("skipping: fixture_07 missing");
        return;
    };
    let mut doc = oxidocs_core::parse_docx(&bytes).expect("parse fixture_07");
    oxidocs_core::revisions::reject_all(&mut doc);

    let texts: Vec<String> = collect_runs(&doc).iter().map(|r| r.text.clone()).collect();
    let joined = texts.join("|");
    assert!(joined.contains("del1"), "rejected del survives");
    assert!(!joined.contains("ins1"), "rejected ins removed; got {joined}");
    assert!(!joined.contains("ins2"), "rejected ins2 removed");
}

/// S-03: per-id targeting. `accept_revision(id)` only touches the
/// revision whose `pair_id` matches.
#[test]
fn s03_accept_revision_by_id_leaves_others_untouched() {
    let Some(bytes) = read_fixture("fixture_07_mixed_ins_del.docx") else {
        eprintln!("skipping: fixture_07 missing");
        return;
    };
    let mut doc = oxidocs_core::parse_docx(&bytes).expect("parse fixture_07");

    // fixture_07 builder uses w:id="100" for ins1, "101" for del1, "102" for ins2.
    oxidocs_core::revisions::accept_revision(&mut doc, "101");

    let runs = collect_runs(&doc);
    let texts: Vec<String> = runs.iter().map(|r| r.text.clone()).collect();
    let joined = texts.join("|");
    // del1 (id=101) is gone (accepted = removed)
    assert!(!joined.contains("del1"), "id=101 (del) must be accepted away");
    // ins1 / ins2 still carry tracked_change (untouched)
    let with_tc: Vec<&str> = runs.iter()
        .filter(|r| r.tracked_change.is_some())
        .map(|r| r.text.as_str())
        .collect();
    assert!(
        with_tc.iter().any(|t| t.contains("ins1")),
        "ins1 should still be tracked; with_tc={with_tc:?}"
    );
    assert!(
        with_tc.iter().any(|t| t.contains("ins2")),
        "ins2 should still be tracked"
    );
}

/// S-02: `ShowRevisions::Final` drops `<w:del>` runs and renders surviving
/// `<w:ins>` runs without underline or color (post-edit / accepted view).
#[test]
fn s02_show_revisions_final_drops_del_and_strips_ins_styling() {
    let Some(bytes) = read_fixture("fixture_07_mixed_ins_del.docx") else {
        eprintln!("skipping: fixture_07 missing");
        return;
    };
    let doc = oxidocs_core::parse_docx(&bytes).expect("parse fixture_07");
    let engine = oxidocs_core::layout::LayoutEngine::for_document(&doc)
        .with_show_revisions(oxidocs_core::ir::ShowRevisions::Final);
    let result = engine.layout(&doc);

    let mut texts: Vec<String> = Vec::new();
    let mut underlined: Vec<bool> = Vec::new();
    let mut struck: Vec<bool> = Vec::new();
    let mut colors: Vec<Option<String>> = Vec::new();
    for page in &result.pages {
        for el in &page.elements {
            if let oxidocs_core::layout::LayoutContent::Text {
                text, underline, strikethrough, color, ..
            } = &el.content
            {
                if !text.trim().is_empty() {
                    texts.push(text.clone());
                    underlined.push(*underline);
                    struck.push(*strikethrough);
                    colors.push(color.clone());
                }
            }
        }
    }

    let joined = texts.join("|");
    // del1 must be gone
    assert!(
        !joined.contains("del1"),
        "Final view should drop <w:del>; got texts={joined}"
    );
    // ins1 / ins2 must still be present (as normal text).
    assert!(joined.contains("ins1"), "Final view keeps ins; got {joined}");
    assert!(joined.contains("ins2"), "Final view keeps ins2; got {joined}");
    // No element should have underline=true under Final (revision styling stripped).
    assert!(
        underlined.iter().all(|&u| !u),
        "Final view: ins runs should not be underlined"
    );
    // No element should be struck under Final (no del runs survive).
    assert!(struck.iter().all(|&s| !s));
    // No element should carry the Alice author color.
    assert!(
        colors.iter().all(|c| c.as_deref() != Some("#D03337")),
        "Final view should not author-tint surviving ins runs"
    );
}

/// S-02: `ShowRevisions::Original` drops `<w:ins>` runs and renders
/// surviving `<w:del>` runs as normal (pre-edit / rejected view).
#[test]
fn s02_show_revisions_original_drops_ins_and_strips_del_styling() {
    let Some(bytes) = read_fixture("fixture_07_mixed_ins_del.docx") else {
        eprintln!("skipping: fixture_07 missing");
        return;
    };
    let doc = oxidocs_core::parse_docx(&bytes).expect("parse fixture_07");
    let engine = oxidocs_core::layout::LayoutEngine::for_document(&doc)
        .with_show_revisions(oxidocs_core::ir::ShowRevisions::Original);
    let result = engine.layout(&doc);

    let mut texts: Vec<String> = Vec::new();
    for page in &result.pages {
        for el in &page.elements {
            if let oxidocs_core::layout::LayoutContent::Text { text, .. } = &el.content {
                if !text.trim().is_empty() {
                    texts.push(text.clone());
                }
            }
        }
    }
    let joined = texts.join("|");
    assert!(!joined.contains("ins1"), "Original: ins1 must be dropped");
    assert!(!joined.contains("ins2"), "Original: ins2 must be dropped");
    assert!(joined.contains("del1"), "Original: del must survive");
}

/// S-02: `ShowRevisions::Simple` keeps all runs visible as normal text but
/// the per-line margin change bar (R-10) still fires for revision-bearing
/// lines.
#[test]
fn s02_show_revisions_simple_skips_color_keeps_margin_bar() {
    let Some(bytes) = read_fixture("fixture_05_single_ins.docx") else {
        eprintln!("skipping: fixture_05 missing");
        return;
    };
    let doc = oxidocs_core::parse_docx(&bytes).expect("parse fixture_05");
    let engine = oxidocs_core::layout::LayoutEngine::for_document(&doc)
        .with_show_revisions(oxidocs_core::ir::ShowRevisions::Simple);
    let result = engine.layout(&doc);

    let mut underlined_with_color = 0;
    let mut margin_bars = 0;
    for page in &result.pages {
        for el in &page.elements {
            match &el.content {
                oxidocs_core::layout::LayoutContent::Text { underline, color, .. } => {
                    if *underline && color.as_deref() == Some("#D03337") {
                        underlined_with_color += 1;
                    }
                }
                oxidocs_core::layout::LayoutContent::BoxRect { fill, .. }
                    if el.width <= 2.0 && fill.as_deref() == Some("#424242") =>
                {
                    margin_bars += 1;
                }
                _ => {}
            }
        }
    }
    assert_eq!(
        underlined_with_color, 0,
        "Simple view: ins runs must NOT be author-tinted-underlined"
    );
    assert!(
        margin_bars >= 1,
        "Simple view: at least one margin change bar should still fire"
    );
}

/// S-01: `show_comments=false` suppresses every comment-related visual:
/// balloons, connectors, in-line range highlight, and body-width compression.
/// Tracked-change rendering (R-01/R-03/etc.) is independent and stays on —
/// that's what S-02 controls.
#[test]
fn s01_show_comments_false_suppresses_all_comment_visuals() {
    let Some(bytes) = read_fixture("fixture_01_single_comment.docx") else {
        eprintln!("skipping: fixture_01 missing");
        return;
    };
    let doc = oxidocs_core::parse_docx(&bytes).expect("parse fixture_01");

    // With show_comments=false: no balloons, no connectors, no range highlight,
    // and the body uses full width (1 line instead of 2).
    let engine_off = oxidocs_core::layout::LayoutEngine::for_document(&doc).with_show_comments(false);
    let result_off = engine_off.layout(&doc);

    let mut balloons = 0;
    let mut connectors = 0;
    let mut highlighted = 0;
    let mut text_y_set = std::collections::BTreeSet::new();
    for page in &result_off.pages {
        for el in &page.elements {
            match &el.content {
                oxidocs_core::layout::LayoutContent::Balloon { .. } => balloons += 1,
                oxidocs_core::layout::LayoutContent::BalloonConnector { .. } => connectors += 1,
                oxidocs_core::layout::LayoutContent::Text { text, highlight, .. } => {
                    if !text.is_empty() {
                        text_y_set.insert(el.y.round() as i32);
                    }
                    if highlight.is_some() {
                        highlighted += 1;
                    }
                }
                _ => {}
            }
        }
    }
    assert_eq!(balloons, 0, "show_comments=false should suppress balloons");
    assert_eq!(connectors, 0, "show_comments=false should suppress connectors");
    assert_eq!(
        highlighted, 0,
        "show_comments=false should suppress inline range highlight"
    );
    assert_eq!(
        text_y_set.len(),
        1,
        "with show_comments=false body uses full width — fixture_01 fits on 1 line"
    );

    // Sanity — with the default (show_comments=true), all three are present.
    let engine_on = oxidocs_core::layout::LayoutEngine::for_document(&doc);
    let result_on = engine_on.layout(&doc);
    let mut on_balloons = 0;
    for page in &result_on.pages {
        for el in &page.elements {
            if matches!(&el.content, oxidocs_core::layout::LayoutContent::Balloon { .. }) {
                on_balloons += 1;
            }
        }
    }
    assert_eq!(on_balloons, 1, "default behavior emits 1 balloon");
}

/// R-05f: when a comment has a reply (Comment.parent_para_id set), the
/// reply's body folds into the parent's `Balloon.replies` Vec rather than
/// emitting as a standalone balloon. Word renders the reply indented inside
/// the same balloon (single physical balloon for the parent + child thread).
#[test]
fn fixture_02_reply_folds_into_parent_balloon_replies_vec() {
    let Some(bytes) = read_fixture("fixture_02_comment_with_reply.docx") else {
        eprintln!("skipping: fixture_02 missing");
        return;
    };
    let doc = oxidocs_core::parse_docx(&bytes).expect("parse fixture_02");
    let result = layout_doc(&doc);

    let mut balloons = 0_usize;
    let mut found_reply_body: Option<String> = None;
    let mut found_reply_author: Option<String> = None;
    for page in &result.pages {
        for el in &page.elements {
            if let oxidocs_core::layout::LayoutContent::Balloon { replies, body, .. } = &el.content {
                balloons += 1;
                assert!(body.contains("Why?"), "parent body should contain 'Why?'; got {body:?}");
                if !replies.is_empty() {
                    let r = &replies[0];
                    found_reply_body = Some(r.body.clone());
                    found_reply_author = Some(r.author.clone());
                }
            }
        }
    }
    assert_eq!(
        balloons, 1,
        "expected 1 standalone balloon (reply must fold into parent); got {balloons}"
    );
    assert_eq!(
        found_reply_body.as_deref(),
        Some("Following up."),
        "reply body must be exposed via Balloon.replies"
    );
    assert_eq!(found_reply_author.as_deref(), Some("Alice Reviewer"));
}

/// R-05e: alongside every Balloon LayoutElement, one `BalloonConnector`
/// element is emitted. Connector starts at the inline anchor coordinates
/// (commentReference X/Y on the body) and ends at the balloon's left edge,
/// ~5pt below the balloon's top.
#[test]
fn fixture_01_emits_balloon_connector_paired_with_balloon() {
    let Some(bytes) = read_fixture("fixture_01_single_comment.docx") else {
        eprintln!("skipping: fixture_01 missing");
        return;
    };
    let doc = oxidocs_core::parse_docx(&bytes).expect("parse fixture_01");
    let result = layout_doc(&doc);

    let mut connectors: Vec<(f32, f32, f32, f32, String)> = Vec::new();
    let mut balloon_left: Option<f32> = None;
    let mut balloon_top: Option<f32> = None;
    for page in &result.pages {
        for el in &page.elements {
            match &el.content {
                oxidocs_core::layout::LayoutContent::BalloonConnector {
                    from_x, from_y, to_x, to_y, color_hex,
                } => {
                    connectors.push((*from_x, *from_y, *to_x, *to_y, color_hex.clone()));
                }
                oxidocs_core::layout::LayoutContent::Balloon { .. } => {
                    balloon_left = Some(el.x);
                    balloon_top = Some(el.y);
                }
                _ => {}
            }
        }
    }

    assert_eq!(connectors.len(), 1, "expected exactly 1 BalloonConnector");
    let (from_x, from_y, to_x, to_y, color) = &connectors[0];
    let bl = balloon_left.expect("balloon must also be emitted");
    let bt = balloon_top.expect("balloon must also be emitted");

    // Connector ends at the balloon's left edge.
    assert!(
        (*to_x - bl).abs() < 0.01,
        "connector to_x={to_x} should equal balloon_left={bl}"
    );
    // Ends 5pt below the balloon's top.
    assert!(
        (*to_y - (bt + 5.0)).abs() < 0.01,
        "connector to_y={to_y} should equal balloon_top+5={}",
        bt + 5.0
    );
    // Starts at the body — the anchor X must be smaller than the balloon's left edge.
    assert!(
        *from_x < bl,
        "anchor x={from_x} must be left of balloon (left={bl})"
    );
    // Color = Alice slot 0 unresolved tint.
    assert_eq!(color, "#FAE6E7", "connector color should be Alice's tint");
    // Sanity on from_y: should be on the body (positive, but well above balloon's top).
    assert!(*from_y > 0.0 && *from_y <= bt + 5.0);
}

/// R-05c: each visible comment emits one `LayoutContent::Balloon` element on
/// the page where its scope begins. Width is 293.8pt (unresolved) or 190.1pt
/// (resolved); right edge sits ~4pt from the page edge; anchor Y matches the
/// rendered Y of the `commentRangeStart` line.
#[test]
fn fixture_01_emits_one_balloon_for_single_comment() {
    let Some(bytes) = read_fixture("fixture_01_single_comment.docx") else {
        eprintln!("skipping: fixture_01 missing");
        return;
    };
    let doc = oxidocs_core::parse_docx(&bytes).expect("parse fixture_01");
    let result = layout_doc(&doc);

    let mut balloons: Vec<(String, String, bool, f32, f32, f32, f32)> = Vec::new();
    for page in &result.pages {
        for el in &page.elements {
            if let oxidocs_core::layout::LayoutContent::Balloon {
                comment_id, author, resolved, ..
            } = &el.content
            {
                balloons.push((
                    comment_id.clone(),
                    author.clone(),
                    *resolved,
                    el.x,
                    el.y,
                    el.width,
                    el.height,
                ));
            }
        }
    }

    assert_eq!(balloons.len(), 1, "fixture_01 should emit exactly 1 balloon");
    let (cid, author, resolved, x, _y, w, h) = &balloons[0];
    assert_eq!(cid, "0");
    assert_eq!(author, "Alice Reviewer");
    assert!(!resolved);
    // unresolved width = 293.8pt.
    assert!(
        (*w - 293.8).abs() < 0.01,
        "unresolved balloon width should be 293.8pt; got {w}"
    );
    // Right edge ≈ page_width − 4pt → for A4 (595.3pt): right = 591.4pt; left = 591.4 − 293.8 = 297.6pt.
    assert!(
        (*x - 297.5).abs() < 0.5,
        "balloon x should be ~297.6pt for default A4; got {x}"
    );
    assert!(*h > 8.0, "balloon should have nonzero height; got {h}");
}

/// R-05c: a resolved comment emits a narrower balloon (190.1pt vs 293.8pt).
#[test]
fn fixture_03_emits_resolved_balloon_with_narrower_width() {
    let Some(bytes) = read_fixture("fixture_03_resolved_comment.docx") else {
        eprintln!("skipping: fixture_03 missing");
        return;
    };
    let doc = oxidocs_core::parse_docx(&bytes).expect("parse fixture_03");
    let result = layout_doc(&doc);

    let mut found = false;
    for page in &result.pages {
        for el in &page.elements {
            if let oxidocs_core::layout::LayoutContent::Balloon { resolved, .. } = &el.content {
                assert!(*resolved, "fixture_03's balloon must carry resolved=true");
                assert!(
                    (el.width - 190.1).abs() < 0.01,
                    "resolved balloon width should be 190.1pt; got {}",
                    el.width
                );
                found = true;
            }
        }
    }
    assert!(found, "fixture_03 should emit one resolved balloon");
}

/// R-05b: when the document has any comments, the body content width is
/// reduced by the balloon column width (293.8pt + buffer). The text "The
/// quick brown fox jumps over the lazy dog." (~245pt at 11pt Calibri) wraps
/// to 2 lines instead of fitting on one line.
#[test]
fn fixture_01_body_width_compresses_when_comments_present() {
    fn distinct_text_line_count(result: &oxidocs_core::layout::LayoutResult) -> usize {
        let mut ys: std::collections::BTreeSet<i32> = std::collections::BTreeSet::new();
        for page in &result.pages {
            for el in &page.elements {
                if let oxidocs_core::layout::LayoutContent::Text { text, .. } = &el.content {
                    if !text.is_empty() {
                        // Quantize y to nearest pt to merge same-line fragments.
                        ys.insert(el.y.round() as i32);
                    }
                }
            }
        }
        ys.len()
    }

    let bytes_with = match read_fixture("fixture_01_single_comment.docx") {
        Some(b) => b,
        None => {
            eprintln!("skipping: fixture_01 missing");
            return;
        }
    };
    let with_doc = oxidocs_core::parse_docx(&bytes_with).expect("parse fixture_01");
    let with_result = layout_doc(&with_doc);

    // fixture_01's body sentence is ~245pt at default 11pt Calibri. Without
    // compression body width = 451pt → fits on 1 line. With R-05b's 317.8pt
    // balloon column reservation, body width drops to ~133pt → text wraps to
    // ≥2 lines.
    let lines = distinct_text_line_count(&with_result);
    assert!(
        lines >= 2,
        "comment-bearing fixture_01 should wrap to ≥2 lines under R-05b body compression; got {lines} lines"
    );

    // Sanity: same fixture parsed without comments would still produce 1 line.
    // Build by stripping the comments from the doc and re-laying out, to
    // isolate the compression effect from text content.
    let mut without = with_doc.clone();
    without.comments.clear();
    let without_result = layout_doc(&without);
    let without_lines = distinct_text_line_count(&without_result);
    assert_eq!(
        without_lines, 1,
        "with comments cleared, fixture_01 should fit on 1 line; got {without_lines}"
    );
}

/// R-10: a paragraph-mark insert/delete (`<w:pPr>/<w:rPr>/<w:ins>` or
/// `<w:pPr>/<w:rPr>/<w:del>` — P-09's `paragraph_mark_revision` field)
/// fires the change bar even when no run carries `tracked_change` /
/// `rpr_change` and there is no `pPrChange` either. Verifies the second
/// half of the paragraph-level R-10 path.
#[test]
fn r10_fires_for_paragraph_mark_revision() {
    let Some(bytes) = read_fixture("fixture_05_single_ins.docx") else {
        eprintln!("skipping: fixture_05 missing");
        return;
    };
    let mut doc = oxidocs_core::parse_docx(&bytes).expect("parse fixture_05");

    for page in &mut doc.pages {
        for block in &mut page.blocks {
            if let oxidocs_core::ir::Block::Paragraph(p) = block {
                for run in &mut p.runs {
                    run.tracked_change = None;
                    run.rpr_change = None;
                }
                p.ppr_change = None;
                p.paragraph_mark_revision = Some(oxidocs_core::ir::TrackedChange {
                    change_type: "insert".into(),
                    author: Some("Alice Reviewer".into()),
                    date: None,
                    pair_id: Some("777".into()),
                });
            }
        }
    }

    let result = layout_doc(&doc);
    let bars = result
        .pages
        .iter()
        .flat_map(|p| p.elements.iter())
        .filter(|el| {
            matches!(&el.content, oxidocs_core::layout::LayoutContent::BoxRect { fill, .. }
                if el.width <= 2.0 && fill.as_deref() == Some("#424242"))
        })
        .count();
    assert!(
        bars >= 1,
        "paragraph_mark_revision alone must fire ≥1 margin change bar; got {bars}"
    );
}

/// R-10: paragraph-level revision (`pPrChange`) fires the change bar even
/// when no run on the line carries `tracked_change` / `rpr_change`. Verified
/// by mutating fixture_05's parsed IR to strip run-level revisions and
/// install a synthetic `ppr_change` instead.
#[test]
fn r10_fires_for_paragraph_level_ppr_change() {
    let Some(bytes) = read_fixture("fixture_05_single_ins.docx") else {
        eprintln!("skipping: fixture_05 missing");
        return;
    };
    let mut doc = oxidocs_core::parse_docx(&bytes).expect("parse fixture_05");

    // Strip every run-level revision so only the synthetic ppr_change can
    // trigger R-10. This isolates the paragraph-level code path.
    for page in &mut doc.pages {
        for block in &mut page.blocks {
            if let oxidocs_core::ir::Block::Paragraph(p) = block {
                for run in &mut p.runs {
                    run.tracked_change = None;
                    run.rpr_change = None;
                }
                p.ppr_change = Some(oxidocs_core::ir::PropertyChange {
                    id: Some("999".into()),
                    author: Some("Alice Reviewer".into()),
                    date: None,
                    prior_run_style: None,
                    prior_paragraph_style: None,
                    prior_alignment: None,
                });
            }
        }
    }

    let result = layout_doc(&doc);
    let mut bars = 0;
    for page in &result.pages {
        for el in &page.elements {
            if let oxidocs_core::layout::LayoutContent::BoxRect { fill, .. } = &el.content {
                if el.width <= 2.0 && fill.as_deref() == Some("#424242") {
                    bars += 1;
                }
            }
        }
    }
    assert!(
        bars >= 1,
        "ppr_change alone must fire ≥1 margin change bar; got {bars}"
    );
}

/// R-10: every line containing a revision-bearing run gets a single dark-grey
/// margin change bar (1.5pt thick) emitted as `LayoutContent::BoxRect`. Lines
/// without revisions get no bar.
#[test]
fn fixture_05_layout_emits_revision_change_bar() {
    let Some(bytes) = read_fixture("fixture_05_single_ins.docx") else {
        eprintln!("skipping: fixture_05 missing");
        return;
    };
    let doc = oxidocs_core::parse_docx(&bytes).expect("parse fixture_05");
    let result = layout_doc(&doc);

    let mut bars: Vec<(f32, f32, f32, f32, Option<String>)> = Vec::new();
    for page in &result.pages {
        for el in &page.elements {
            if let oxidocs_core::layout::LayoutContent::BoxRect { fill, .. } = &el.content {
                // Filter to thin (≤2pt wide) bars to exclude any unrelated
                // BoxRects (e.g., box outlines).
                if el.width <= 2.0 {
                    bars.push((el.x, el.y, el.width, el.height, fill.clone()));
                }
            }
        }
    }

    assert_eq!(
        bars.len(),
        1,
        "fixture_05 has one revision-bearing line, expected exactly 1 change bar; got {} ({:?})",
        bars.len(),
        bars
    );
    let (x, _y, w, h, fill) = &bars[0];
    assert_eq!(fill.as_deref(), Some("#424242"), "change bar fill");
    assert!(*w <= 2.0, "change bar should be thin");
    assert!(*h > 8.0, "change bar should span the line height");
    assert!(
        *x >= 0.0 && *x < 72.0,
        "change bar should sit in the left margin (0..72pt for 1in margin), got x={x}"
    );
}

/// R-04: in-line comment-range highlight. Runs strictly BETWEEN
/// `commentRangeStart` and `commentRangeEnd` carry a background highlight
/// matching the author's tint. Runs outside the range must NOT be highlighted.
#[test]
fn fixture_01_layout_comment_range_highlight_inline() {
    let Some(bytes) = read_fixture("fixture_01_single_comment.docx") else {
        eprintln!("skipping: fixture_01 missing");
        return;
    };
    let doc = oxidocs_core::parse_docx(&bytes).expect("parse fixture_01");
    let result = layout_doc(&doc);

    fn highlights_for(
        result: &oxidocs_core::layout::LayoutResult,
        needle: &str,
    ) -> Vec<Option<String>> {
        let mut out = Vec::new();
        for page in &result.pages {
            for el in &page.elements {
                if let oxidocs_core::layout::LayoutContent::Text { text, highlight, .. } =
                    &el.content
                {
                    if text.contains(needle) {
                        out.push(highlight.clone());
                    }
                }
            }
        }
        out
    }

    // "brown" is inside the range "brown fox" — must be tinted.
    let brown = highlights_for(&result, "brown");
    assert!(!brown.is_empty(), "'brown' element missing from layout");
    for h in &brown {
        assert_eq!(
            h.as_deref(),
            Some("#FAE6E7"),
            "Alice's unresolved range highlight must use slot-0 tint"
        );
    }

    // "The" (before the range) must NOT be tinted.
    let leading = highlights_for(&result, "The");
    assert!(!leading.is_empty());
    for h in &leading {
        assert!(
            h.is_none(),
            "pre-range text must not be highlighted, got {h:?}"
        );
    }

    // "jumps" (after the range) must NOT be tinted.
    let trailing = highlights_for(&result, "jumps");
    assert!(!trailing.is_empty());
    for h in &trailing {
        assert!(
            h.is_none(),
            "post-range text must not be highlighted, got {h:?}"
        );
    }
}

/// R-09 (in-line half): a resolved comment (`<w15:done="1"/>`) uses the
/// desaturated tint palette instead of the unresolved palette. Slot 0 drops
/// from #FAE6E7 to #F1EDEC — chroma falls to near-grey while lightness is
/// preserved (COM-confirmed 2026-04-25 from fixture_03 balloon background).
#[test]
fn fixture_03_layout_resolved_comment_uses_desaturated_tint() {
    let Some(bytes) = read_fixture("fixture_03_resolved_comment.docx") else {
        eprintln!("skipping: fixture_03 missing");
        return;
    };
    let doc = oxidocs_core::parse_docx(&bytes).expect("parse fixture_03");
    let result = layout_doc(&doc);

    fn highlight_for(
        result: &oxidocs_core::layout::LayoutResult,
        needle: &str,
    ) -> Option<String> {
        for page in &result.pages {
            for el in &page.elements {
                if let oxidocs_core::layout::LayoutContent::Text { text, highlight, .. } =
                    &el.content
                {
                    if text.contains(needle) {
                        return highlight.clone();
                    }
                }
            }
        }
        None
    }

    // "reviewed" is inside the resolved range "has been reviewed" — must be
    // tinted with the *resolved* palette, not the unresolved one.
    let h = highlight_for(&result, "reviewed");
    assert_eq!(
        h.as_deref(),
        Some("#F1EDEC"),
        "resolved range must use the desaturated tint, not #FAE6E7"
    );
}

/// R-04: multi-paragraph range. The comment covers 3 paragraphs; all three
/// must be highlighted. Verifies the state machine carries `open` across
/// paragraph boundaries, and that the parser's new anchor-run fallback (for
/// `<w:commentRangeStart>` as the first child of a paragraph) still emits the
/// id onto the IR.
#[test]
fn fixture_04_layout_multi_paragraph_range_highlight() {
    let Some(bytes) = read_fixture("fixture_04_multi_para_range.docx") else {
        eprintln!("skipping: fixture_04 missing");
        return;
    };
    let doc = oxidocs_core::parse_docx(&bytes).expect("parse fixture_04");
    let result = layout_doc(&doc);

    fn first_highlight_for(
        result: &oxidocs_core::layout::LayoutResult,
        needle: &str,
    ) -> Option<Option<String>> {
        for page in &result.pages {
            for el in &page.elements {
                if let oxidocs_core::layout::LayoutContent::Text { text, highlight, .. } =
                    &el.content
                {
                    if text.contains(needle) {
                        return Some(highlight.clone());
                    }
                }
            }
        }
        None
    }

    for para_marker in &["First", "Second", "Third"] {
        let h = first_highlight_for(&result, para_marker).unwrap_or_else(|| {
            panic!("'{para_marker}' not found in layout output")
        });
        assert_eq!(
            h.as_deref(),
            Some("#FAE6E7"),
            "{para_marker} must be inside the range and get Alice's tint"
        );
    }
}

/// R-11: `<w:moveFrom>` and `<w:moveTo>` always render in green (#2B6033)
/// regardless of the author's palette slot — Word's hard-coded behavior
/// (COM-confirmed 2026-04-25 in fixture_08).
///
/// R-11 v2 (R66, 2026-04-29): moveFrom uses **double** strikethrough and
/// moveTo uses **double** underline. Confirmed by pixel-sampling fixture_08
/// rendered output: two full-width green lines at y=164/167 (strike) and
/// y=220/222 (underline), 1pt apart, on the "moved clause" runs.
#[test]
fn fixture_08_layout_moves_render_in_green() {
    let Some(bytes) = read_fixture("fixture_08_move_from_to.docx") else {
        eprintln!("skipping: fixture_08 missing");
        return;
    };
    let doc = oxidocs_core::parse_docx(&bytes).expect("parse fixture_08");
    let result = layout_doc(&doc);

    // Walk LayoutContent::Text fragments matching "moved" and capture the
    // full styling tuple — R-11 needs underline_style and double_strikethrough,
    // not just the bool pair the shared helper exposes.
    type MoveHit = (bool, Option<String>, bool, bool, Option<String>);
    let move_hits: Vec<MoveHit> = {
        let mut out = Vec::new();
        for page in &result.pages {
            for el in &page.elements {
                if let oxidocs_core::layout::LayoutContent::Text {
                    text,
                    underline,
                    underline_style,
                    strikethrough,
                    double_strikethrough,
                    color,
                    ..
                } = &el.content
                {
                    if text.contains("moved") {
                        out.push((
                            *underline,
                            underline_style.clone(),
                            *strikethrough,
                            *double_strikethrough,
                            color.clone(),
                        ));
                    }
                }
            }
        }
        out
    };
    assert!(
        move_hits.len() >= 2,
        "expected ≥2 'moved' fragments (one moveFrom + one moveTo); got {}",
        move_hits.len()
    );
    for (_, _, _, _, color) in &move_hits {
        assert_eq!(
            color.as_deref(),
            Some("#2B6033"),
            "moveFrom/moveTo render in fixed green regardless of author"
        );
    }

    // moveFrom: strikethrough=true AND double_strikethrough=true (no underline).
    let any_double_struck = move_hits
        .iter()
        .any(|(u, _, s, ds, _)| *s && *ds && !*u);
    assert!(
        any_double_struck,
        "moveFrom must render with double strikethrough (R-11 v2, R66 COM-confirmed)"
    );

    // moveTo: underline=true AND underline_style=Some(\"double\") (no strikethrough).
    let any_double_underlined = move_hits
        .iter()
        .any(|(u, us, s, _, _)| *u && us.as_deref() == Some("double") && !*s);
    assert!(
        any_double_underlined,
        "moveTo must render with double underline (R-11 v2, R66 COM-confirmed)"
    );
}

/// R-12 (R67, 2026-04-29): a `<w:rPrChange>` run anchors a "Formatted: …"
/// balloon in the right margin. Pixel-confirmed against fixture_09's
/// rendered PNG: balloon column starts at x≈401pt (resolved-balloon left
/// edge), balloon top is at y≈158pt, body line "Formatted: Bold" sits at
/// y≈166pt. The Oxi side emits the balloon with author = revision author,
/// resolved=true (narrow grey geometry), and a body string built from the
/// style diff between `rpr_change.prior_run_style` and the run's current
/// style.
#[test]
fn fixture_09_layout_emits_rprchange_margin_balloon() {
    let Some(bytes) = read_fixture("fixture_09_rPrChange_bold.docx") else {
        eprintln!("skipping: fixture_09 missing");
        return;
    };
    let doc = oxidocs_core::parse_docx(&bytes).expect("parse fixture_09");
    let result = layout_doc(&doc);

    let mut balloon_count = 0_usize;
    let mut found_body: Option<String> = None;
    let mut found_author: Option<String> = None;
    let mut found_resolved: Option<bool> = None;
    for page in &result.pages {
        for el in &page.elements {
            if let oxidocs_core::layout::LayoutContent::Balloon {
                comment_id,
                body,
                author,
                resolved,
                ..
            } = &el.content
            {
                if comment_id.starts_with("rprchange:") {
                    balloon_count += 1;
                    found_body = Some(body.clone());
                    found_author = Some(author.clone());
                    found_resolved = Some(*resolved);
                }
            }
        }
    }
    assert_eq!(
        balloon_count, 1,
        "fixture_09 has 1 rPrChange (bold toggle) → must emit exactly 1 R-12 balloon"
    );
    let body = found_body.unwrap();
    assert!(
        body.starts_with("Formatted:"),
        "R-12 balloon body must start with 'Formatted:'; got {body:?}"
    );
    assert!(
        body.contains("Bold"),
        "fixture_09 toggles bold on; body should mention 'Bold'; got {body:?}"
    );
    assert_eq!(
        found_author.as_deref(),
        Some("Alice Reviewer"),
        "R-12 balloon author must be the rPrChange author"
    );
    assert_eq!(
        found_resolved,
        Some(true),
        "R-12 balloons use the narrow (resolved-width) geometry to mirror Word's 'Formatted' balloon"
    );
}

/// R-12 v2 (R69, 2026-04-29): a `<w:pPrChange>` paragraph anchors a
/// "Formatted: …" balloon in the right margin, mirroring v1's run-level
/// rPrChange behaviour. fixture_13 toggles paragraph indent (0 → 720dxa
/// = 36pt left); the balloon body must mention "Indent Left" and the
/// synthetic comment_id must use the `pprchange:` prefix to keep
/// run-level and paragraph-level entries distinguishable.
#[test]
fn fixture_13_layout_emits_pprchange_margin_balloon() {
    let Some(bytes) = read_fixture("fixture_13_pPrChange_indent.docx") else {
        eprintln!("skipping: fixture_13 missing");
        return;
    };
    let doc = oxidocs_core::parse_docx(&bytes).expect("parse fixture_13");
    let result = layout_doc(&doc);

    let mut pprchange_count = 0_usize;
    let mut rprchange_count = 0_usize;
    let mut found_body: Option<String> = None;
    let mut found_author: Option<String> = None;
    let mut found_resolved: Option<bool> = None;
    for page in &result.pages {
        for el in &page.elements {
            if let oxidocs_core::layout::LayoutContent::Balloon {
                comment_id,
                body,
                author,
                resolved,
                ..
            } = &el.content
            {
                if comment_id.starts_with("pprchange:") {
                    pprchange_count += 1;
                    found_body = Some(body.clone());
                    found_author = Some(author.clone());
                    found_resolved = Some(*resolved);
                } else if comment_id.starts_with("rprchange:") {
                    rprchange_count += 1;
                }
            }
        }
    }
    assert_eq!(
        pprchange_count, 1,
        "fixture_13 has 1 pPrChange (indent toggle) → must emit exactly 1 pprchange-prefixed balloon"
    );
    assert_eq!(
        rprchange_count, 0,
        "fixture_13 has no rPrChange — no rprchange-prefixed balloon expected"
    );
    let body = found_body.unwrap();
    assert!(
        body.starts_with("Formatted:"),
        "R-12 v2 balloon body must start with 'Formatted:'; got {body:?}"
    );
    assert!(
        body.contains("Indent Left"),
        "fixture_13 toggles indent → body should mention 'Indent Left'; got {body:?}"
    );
    assert_eq!(
        found_author.as_deref(),
        Some("Alice Reviewer"),
        "R-12 v2 balloon author must be the pPrChange author"
    );
    assert_eq!(
        found_resolved,
        Some(true),
        "R-12 balloons use narrow (resolved-width) geometry"
    );
}

/// R-12 v1 + R71 (2026-04-29): a multi-property `<w:rPrChange>` (font
/// family + font size in a single change) lays out as ONE balloon whose
/// body lists both diffs comma-separated. Confirms `describe_rpr_diff`'s
/// font_family branch (added R71) and the helper's comma-join behaviour
/// when more than one property toggles in a single revision.
#[test]
fn fixture_14_layout_rprchange_multi_property_describe_diff() {
    let Some(bytes) = read_fixture("fixture_14_rPrChange_font.docx") else {
        eprintln!("skipping: fixture_14 missing");
        return;
    };
    let doc = oxidocs_core::parse_docx(&bytes).expect("parse fixture_14");
    let result = layout_doc(&doc);

    let mut found_body: Option<String> = None;
    let mut balloon_count = 0_usize;
    for page in &result.pages {
        for el in &page.elements {
            if let oxidocs_core::layout::LayoutContent::Balloon {
                comment_id, body, ..
            } = &el.content
            {
                if comment_id.starts_with("rprchange:") {
                    balloon_count += 1;
                    found_body = Some(body.clone());
                }
            }
        }
    }
    assert_eq!(
        balloon_count, 1,
        "fixture_14 has 1 rPrChange (font + size) → 1 balloon"
    );
    let body = found_body.unwrap();
    assert!(
        body.starts_with("Formatted:"),
        "body must start with 'Formatted:'; got {body:?}"
    );
    assert!(
        body.contains("Font:"),
        "fixture_14 changes font_family → body must mention 'Font:'; got {body:?}"
    );
    assert!(
        body.contains("Times New Roman"),
        "the new font name must appear in the body; got {body:?}"
    );
    assert!(
        body.contains("Font Size:"),
        "fixture_14 changes font_size too → body must mention 'Font Size:'; got {body:?}"
    );
    assert!(
        body.contains("14pt"),
        "the new font size must appear in the body; got {body:?}"
    );
    // Comma-join behaviour: there must be a comma separating the two
    // property diffs (the only literal comma in this body).
    assert!(
        body.contains(", "),
        "multi-property diff must be comma-separated; got {body:?}"
    );
}

/// R-12 v3.5 (R72, 2026-04-29): a `<w:pPrChange>` that toggles paragraph
/// alignment surfaces "Alignment: …" in the balloon body. R69 v2 left
/// this gap because Paragraph.alignment lives outside ParagraphStyle;
/// R72 adds `prior_alignment` to PropertyChange so the parser can
/// capture the prior `<w:jc>` and the helper can render the diff.
#[test]
fn fixture_15_layout_pprchange_alignment_toggle() {
    let Some(bytes) = read_fixture("fixture_15_pPrChange_alignment.docx") else {
        eprintln!("skipping: fixture_15 missing");
        return;
    };
    let doc = oxidocs_core::parse_docx(&bytes).expect("parse fixture_15");

    // Parser side: prior_alignment must be captured from the inner
    // pPr's <w:jc w:val="left"/> child.
    let mut found_prior_alignment: Option<oxidocs_core::ir::Alignment> = None;
    for page in &doc.pages {
        for block in &page.blocks {
            if let oxidocs_core::ir::Block::Paragraph(p) = block {
                if let Some(pc) = p.ppr_change.as_ref() {
                    found_prior_alignment = pc.prior_alignment;
                }
            }
        }
    }
    assert_eq!(
        found_prior_alignment,
        Some(oxidocs_core::ir::Alignment::Left),
        "parser must capture prior_alignment = Left from inner pPr/jc"
    );

    // Layout side: balloon body must mention the alignment toggle.
    let result = layout_doc(&doc);
    let mut found_body: Option<String> = None;
    let mut balloon_count = 0_usize;
    for page in &result.pages {
        for el in &page.elements {
            if let oxidocs_core::layout::LayoutContent::Balloon {
                comment_id, body, ..
            } = &el.content
            {
                if comment_id.starts_with("pprchange:") {
                    balloon_count += 1;
                    found_body = Some(body.clone());
                }
            }
        }
    }
    assert_eq!(balloon_count, 1, "fixture_15 has 1 pPrChange → 1 balloon");
    let body = found_body.unwrap();
    assert!(
        body.starts_with("Formatted:"),
        "body must start with 'Formatted:'; got {body:?}"
    );
    assert!(
        body.contains("Alignment:"),
        "alignment toggle must surface as 'Alignment:'; got {body:?}"
    );
    assert!(
        body.contains("Centered"),
        "current alignment is Center → body should label it 'Centered'; got {body:?}"
    );
}

/// R86 (2026-04-29): describe_rpr_diff covers 3 more axes — small_caps,
/// all_caps, character_spacing. fixture_16 toggles all_caps + character
/// _spacing in a single rPrChange to exercise (a) the new branches and
/// (b) the comma-join across multiple new-axis diffs.
#[test]
fn fixture_16_layout_rprchange_caps_and_spacing() {
    let Some(bytes) = read_fixture("fixture_16_rPrChange_caps_spacing.docx") else {
        eprintln!("skipping: fixture_16 missing");
        return;
    };
    let doc = oxidocs_core::parse_docx(&bytes).expect("parse fixture_16");
    let result = layout_doc(&doc);

    let mut found_body: Option<String> = None;
    let mut balloon_count = 0_usize;
    for page in &result.pages {
        for el in &page.elements {
            if let oxidocs_core::layout::LayoutContent::Balloon {
                comment_id, body, ..
            } = &el.content
            {
                if comment_id.starts_with("rprchange:") {
                    balloon_count += 1;
                    found_body = Some(body.clone());
                }
            }
        }
    }
    assert_eq!(
        balloon_count, 1,
        "fixture_16 has 1 rPrChange → 1 balloon"
    );
    let body = found_body.unwrap();
    assert!(
        body.starts_with("Formatted:"),
        "body must start with 'Formatted:'; got {body:?}"
    );
    assert!(
        body.contains("All Caps"),
        "all_caps toggle must surface as 'All Caps'; got {body:?}"
    );
    assert!(
        body.contains("Character Spacing"),
        "character_spacing toggle must surface as 'Character Spacing:'; got {body:?}"
    );
    // Multi-axis diff: comma-separated.
    assert!(
        body.contains(", "),
        "multi-axis diff must be comma-joined; got {body:?}"
    );
}

/// R87 (2026-04-29): describe_rpr_diff covers vertical_align + shading
/// (and font_family_east_asia, untested by this fixture). fixture_17
/// toggles vertical_align=superscript + shading=#FFFF00 in a single
/// rPrChange to exercise both new branches plus comma-join.
#[test]
fn fixture_17_layout_rprchange_valign_and_shading() {
    let Some(bytes) = read_fixture("fixture_17_rPrChange_vAlign_shading.docx") else {
        eprintln!("skipping: fixture_17 missing");
        return;
    };
    let doc = oxidocs_core::parse_docx(&bytes).expect("parse fixture_17");
    let result = layout_doc(&doc);

    let mut found_body: Option<String> = None;
    let mut balloon_count = 0_usize;
    for page in &result.pages {
        for el in &page.elements {
            if let oxidocs_core::layout::LayoutContent::Balloon {
                comment_id, body, ..
            } = &el.content
            {
                if comment_id.starts_with("rprchange:") {
                    balloon_count += 1;
                    found_body = Some(body.clone());
                }
            }
        }
    }
    assert_eq!(
        balloon_count, 1,
        "fixture_17 has 1 rPrChange → 1 balloon"
    );
    let body = found_body.unwrap();
    assert!(
        body.starts_with("Formatted:"),
        "body must start with 'Formatted:'; got {body:?}"
    );
    assert!(
        body.contains("Superscript"),
        "vertical_align=Superscript must surface; got {body:?}"
    );
    assert!(
        body.contains("Shading:"),
        "shading toggle must surface as 'Shading:'; got {body:?}"
    );
    // Multi-axis diff: comma-separated.
    assert!(
        body.contains(", "),
        "multi-axis diff must be comma-joined; got {body:?}"
    );
}

/// R88 (2026-04-29): describe_ppr_diff covers paragraph shading.
/// fixture_18 toggles `<w:pPr>/<w:shd w:fill="FFFF00"/>` (yellow
/// paragraph bg); prior pPr empty. Body must surface
/// "Paragraph Shading: FFFF00" (the "Paragraph " prefix
/// disambiguates from the run-level "Shading:" added in R87).
#[test]
fn fixture_18_layout_pprchange_paragraph_shading() {
    let Some(bytes) = read_fixture("fixture_18_pPrChange_shading.docx") else {
        eprintln!("skipping: fixture_18 missing");
        return;
    };
    let doc = oxidocs_core::parse_docx(&bytes).expect("parse fixture_18");
    let result = layout_doc(&doc);

    let mut found_body: Option<String> = None;
    let mut balloon_count = 0_usize;
    for page in &result.pages {
        for el in &page.elements {
            if let oxidocs_core::layout::LayoutContent::Balloon {
                comment_id, body, ..
            } = &el.content
            {
                if comment_id.starts_with("pprchange:") {
                    balloon_count += 1;
                    found_body = Some(body.clone());
                }
            }
        }
    }
    assert_eq!(
        balloon_count, 1,
        "fixture_18 has 1 pPrChange → 1 balloon"
    );
    let body = found_body.unwrap();
    assert!(
        body.starts_with("Formatted:"),
        "body must start with 'Formatted:'; got {body:?}"
    );
    assert!(
        body.contains("Paragraph Shading:"),
        "paragraph shading toggle must surface; got {body:?}"
    );
    assert!(
        body.contains("FFFF00"),
        "the new shading hex must appear in the body; got {body:?}"
    );
}

/// R89 (2026-04-29): describe_ppr_diff covers keep_next + keep_lines.
/// fixture_19 toggles keep_next via pPrChange. Body must mention
/// "Keep With Next".
///
/// Originally R89 attempted num_id/num_ilvl axes but numPr is parsed
/// separately into NumPrRef returned as a 4th tuple element of
/// parse_paragraph_properties — it never populates ParagraphStyle's
/// num_id, so prior_paragraph_style.num_id is always None. Wiring
/// numPr through PropertyChange's prior_num_pr is a future R72-style
/// 3-layer extension; R89 ships the simpler keep_* bool axes.
#[test]
fn fixture_19_layout_pprchange_keep_next() {
    let Some(bytes) = read_fixture("fixture_19_pPrChange_keep_next.docx") else {
        eprintln!("skipping: fixture_19 missing");
        return;
    };
    let doc = oxidocs_core::parse_docx(&bytes).expect("parse fixture_19");
    let result = layout_doc(&doc);

    let mut found_body: Option<String> = None;
    let mut balloon_count = 0_usize;
    for page in &result.pages {
        for el in &page.elements {
            if let oxidocs_core::layout::LayoutContent::Balloon {
                comment_id, body, ..
            } = &el.content
            {
                if comment_id.starts_with("pprchange:") {
                    balloon_count += 1;
                    found_body = Some(body.clone());
                }
            }
        }
    }
    assert_eq!(
        balloon_count, 1,
        "fixture_19 has 1 pPrChange → 1 balloon"
    );
    let body = found_body.unwrap();
    assert!(
        body.starts_with("Formatted:"),
        "body must start with 'Formatted:'; got {body:?}"
    );
    assert!(
        body.contains("Keep With Next"),
        "keep_next toggle must surface as 'Keep With Next'; got {body:?}"
    );
}

/// R93 (2026-04-30): describe_ppr_diff covers paragraph borders via
/// side-presence summary (no PartialEq derive needed). fixture_20
/// adds a bottom border via pPrChange whose prior pPr was empty
/// (no border). Body must mention "Borders Added".
#[test]
fn fixture_20_layout_pprchange_borders_added() {
    let Some(bytes) = read_fixture("fixture_20_pPrChange_borders.docx") else {
        eprintln!("skipping: fixture_20 missing");
        return;
    };
    let doc = oxidocs_core::parse_docx(&bytes).expect("parse fixture_20");
    let result = layout_doc(&doc);

    let mut found_body: Option<String> = None;
    let mut balloon_count = 0_usize;
    for page in &result.pages {
        for el in &page.elements {
            if let oxidocs_core::layout::LayoutContent::Balloon {
                comment_id, body, ..
            } = &el.content
            {
                if comment_id.starts_with("pprchange:") {
                    balloon_count += 1;
                    found_body = Some(body.clone());
                }
            }
        }
    }
    assert_eq!(
        balloon_count, 1,
        "fixture_20 has 1 pPrChange → 1 balloon"
    );
    let body = found_body.unwrap();
    assert!(
        body.starts_with("Formatted:"),
        "body must start with 'Formatted:'; got {body:?}"
    );
    assert!(
        body.contains("Borders Added"),
        "borders toggle must surface as 'Borders Added'; got {body:?}"
    );
}

/// R94 (2026-04-30): describe_ppr_diff covers tab_stops via position-
/// only summary (mirror of R93 borders side-summary). fixture_21 adds
/// 3 tab stops via pPrChange whose prior pPr was empty.
#[test]
fn fixture_21_layout_pprchange_tabs_added() {
    let Some(bytes) = read_fixture("fixture_21_pPrChange_tabs.docx") else {
        eprintln!("skipping: fixture_21 missing");
        return;
    };
    let doc = oxidocs_core::parse_docx(&bytes).expect("parse fixture_21");
    let result = layout_doc(&doc);

    let mut found_body: Option<String> = None;
    let mut balloon_count = 0_usize;
    for page in &result.pages {
        for el in &page.elements {
            if let oxidocs_core::layout::LayoutContent::Balloon {
                comment_id, body, ..
            } = &el.content
            {
                if comment_id.starts_with("pprchange:") {
                    balloon_count += 1;
                    found_body = Some(body.clone());
                }
            }
        }
    }
    assert_eq!(
        balloon_count, 1,
        "fixture_21 has 1 pPrChange → 1 balloon"
    );
    let body = found_body.unwrap();
    assert!(
        body.starts_with("Formatted:"),
        "body must start with 'Formatted:'; got {body:?}"
    );
    assert!(
        body.contains("Tab Stops Added"),
        "tabs toggle must surface as 'Tab Stops Added'; got {body:?}"
    );
}

/// R95 (2026-04-30): describe_ppr_diff covers num_id (and num_ilvl,
/// not exercised here since prior_ilvl == current_ilvl == 0). R89
/// originally attempted these but parser asymmetry blocked them;
/// R95 patches the parser to mirror inline numPr onto
/// style.num_id/num_ilvl. fixture_22 attaches a paragraph to an
/// inline list (numId=1 ilvl=0) via pPrChange whose prior pPr was
/// empty. Body must contain "Numbering: list 1".
#[test]
#[allow(non_snake_case)]
fn fixture_22_layout_pprchange_inline_numPr_attach() {
    let Some(bytes) = read_fixture("fixture_22_pPrChange_numPr.docx") else {
        eprintln!("skipping: fixture_22 missing");
        return;
    };
    let doc = oxidocs_core::parse_docx(&bytes).expect("parse fixture_22");
    let result = layout_doc(&doc);

    let mut found_body: Option<String> = None;
    let mut balloon_count = 0_usize;
    for page in &result.pages {
        for el in &page.elements {
            if let oxidocs_core::layout::LayoutContent::Balloon {
                comment_id, body, ..
            } = &el.content
            {
                if comment_id.starts_with("pprchange:") {
                    balloon_count += 1;
                    found_body = Some(body.clone());
                }
            }
        }
    }
    assert_eq!(
        balloon_count, 1,
        "fixture_22 has 1 pPrChange → 1 balloon"
    );
    let body = found_body.unwrap();
    assert!(
        body.starts_with("Formatted:"),
        "body must start with 'Formatted:'; got {body:?}"
    );
    assert!(
        body.contains("Numbering: list 1"),
        "num_id attach must surface as 'Numbering: list 1'; got {body:?}"
    );
}

/// R98 (2026-04-30): describe_rpr_diff covers outline + emboss + imprint
/// (3 NEW non-R72 rPr axes). fixture_23 toggles outline + emboss in a
/// single rPrChange to exercise both new branches plus comma-join.
#[test]
fn fixture_23_layout_rprchange_outline_emboss() {
    let Some(bytes) = read_fixture("fixture_23_rPrChange_outline_emboss.docx") else {
        eprintln!("skipping: fixture_23 missing");
        return;
    };
    let doc = oxidocs_core::parse_docx(&bytes).expect("parse fixture_23");
    let result = layout_doc(&doc);

    let mut found_body: Option<String> = None;
    let mut balloon_count = 0_usize;
    for page in &result.pages {
        for el in &page.elements {
            if let oxidocs_core::layout::LayoutContent::Balloon {
                comment_id, body, ..
            } = &el.content
            {
                if comment_id.starts_with("rprchange:") {
                    balloon_count += 1;
                    found_body = Some(body.clone());
                }
            }
        }
    }
    assert_eq!(
        balloon_count, 1,
        "fixture_23 has 1 rPrChange → 1 balloon"
    );
    let body = found_body.unwrap();
    assert!(
        body.starts_with("Formatted:"),
        "body must start with 'Formatted:'; got {body:?}"
    );
    assert!(
        body.contains("Outline"),
        "outline toggle must surface; got {body:?}"
    );
    assert!(
        body.contains("Emboss"),
        "emboss toggle must surface; got {body:?}"
    );
    assert!(
        body.contains(", "),
        "multi-axis diff must be comma-joined; got {body:?}"
    );
}

/// R108 (2026-04-30): describe_ppr_diff extension — bidi + page_break_after
/// + text_alignment (3 more NEW non-R72 ppr axes). fixture_27 toggles bidi
/// ON and text_alignment="top" in one pPrChange.
#[test]
#[allow(non_snake_case)]
fn fixture_27_layout_pprchange_bidi_textAlign() {
    let Some(bytes) = read_fixture("fixture_27_pPrChange_bidi_textAlign.docx") else {
        eprintln!("skipping: fixture_27 missing");
        return;
    };
    let doc = oxidocs_core::parse_docx(&bytes).expect("parse fixture_27");
    let result = layout_doc(&doc);

    let mut found_body: Option<String> = None;
    let mut balloon_count = 0_usize;
    for page in &result.pages {
        for el in &page.elements {
            if let oxidocs_core::layout::LayoutContent::Balloon {
                comment_id, body, ..
            } = &el.content
            {
                if comment_id.starts_with("pprchange:") {
                    balloon_count += 1;
                    found_body = Some(body.clone());
                }
            }
        }
    }
    assert_eq!(
        balloon_count, 1,
        "fixture_27 has 1 pPrChange → 1 balloon"
    );
    let body = found_body.unwrap();
    assert!(
        body.starts_with("Formatted:"),
        "body must start with 'Formatted:'; got {body:?}"
    );
    assert!(
        body.contains("Right-to-Left"),
        "bidi toggle must surface; got {body:?}"
    );
    assert!(
        body.contains("Text Alignment: top"),
        "text_alignment toggle must surface; got {body:?}"
    );
    assert!(
        body.contains(", "),
        "multi-axis diff must be comma-joined; got {body:?}"
    );
}

/// R107 (2026-04-30): describe_ppr_diff NEW non-R72 axes — page_break_before
/// + widow_control + contextual_spacing. fixture_26 toggles
/// page_break_before ON and widow_control OFF in one pPrChange.
#[test]
#[allow(non_snake_case)]
fn fixture_26_layout_pprchange_pageBreak_widow() {
    let Some(bytes) = read_fixture("fixture_26_pPrChange_pageBreak_widow.docx") else {
        eprintln!("skipping: fixture_26 missing");
        return;
    };
    let doc = oxidocs_core::parse_docx(&bytes).expect("parse fixture_26");
    let result = layout_doc(&doc);

    let mut found_body: Option<String> = None;
    let mut balloon_count = 0_usize;
    for page in &result.pages {
        for el in &page.elements {
            if let oxidocs_core::layout::LayoutContent::Balloon {
                comment_id, body, ..
            } = &el.content
            {
                if comment_id.starts_with("pprchange:") {
                    balloon_count += 1;
                    found_body = Some(body.clone());
                }
            }
        }
    }
    assert_eq!(
        balloon_count, 1,
        "fixture_26 has 1 pPrChange → 1 balloon"
    );
    let body = found_body.unwrap();
    assert!(
        body.starts_with("Formatted:"),
        "body must start with 'Formatted:'; got {body:?}"
    );
    assert!(
        body.contains("Page Break Before"),
        "page_break_before toggle must surface; got {body:?}"
    );
    assert!(
        body.contains("Not Widow/Orphan Control"),
        "widow_control off must surface; got {body:?}"
    );
    assert!(
        body.contains(", "),
        "multi-axis diff must be comma-joined; got {body:?}"
    );
}

/// R100 (2026-04-30): describe_rpr_diff extension — highlight + position
/// + emphasis_mark (3 user-visible Word props, Option-typed). fixture_25
/// toggles all three in one rPrChange.
#[test]
fn fixture_25_layout_rprchange_highlight_position_em() {
    let Some(bytes) = read_fixture("fixture_25_rPrChange_highlight_position.docx") else {
        eprintln!("skipping: fixture_25 missing");
        return;
    };
    let doc = oxidocs_core::parse_docx(&bytes).expect("parse fixture_25");
    let result = layout_doc(&doc);

    let mut found_body: Option<String> = None;
    let mut balloon_count = 0_usize;
    for page in &result.pages {
        for el in &page.elements {
            if let oxidocs_core::layout::LayoutContent::Balloon {
                comment_id, body, ..
            } = &el.content
            {
                if comment_id.starts_with("rprchange:") {
                    balloon_count += 1;
                    found_body = Some(body.clone());
                }
            }
        }
    }
    assert_eq!(
        balloon_count, 1,
        "fixture_25 has 1 rPrChange → 1 balloon"
    );
    let body = found_body.unwrap();
    assert!(
        body.starts_with("Formatted:"),
        "body must start with 'Formatted:'; got {body:?}"
    );
    assert!(
        body.contains("Highlight: yellow"),
        "highlight toggle must surface; got {body:?}"
    );
    assert!(
        body.contains("Position:"),
        "position toggle must surface; got {body:?}"
    );
    assert!(
        body.contains("Emphasis Mark: dot"),
        "emphasis_mark toggle must surface; got {body:?}"
    );
    assert!(
        body.contains(", "),
        "multi-axis diff must be comma-joined; got {body:?}"
    );
}

/// R99 (2026-04-30): describe_rpr_diff extension — shadow + vanish +
/// double_strikethrough (3 more NEW non-R72 rPr axes peer to R98).
/// fixture_24 toggles all three in one rPrChange, exercises the new
/// branches plus comma-join across them.
#[test]
fn fixture_24_layout_rprchange_shadow_vanish_dstrike() {
    let Some(bytes) = read_fixture("fixture_24_rPrChange_shadow_vanish_dstrike.docx") else {
        eprintln!("skipping: fixture_24 missing");
        return;
    };
    let doc = oxidocs_core::parse_docx(&bytes).expect("parse fixture_24");
    let result = layout_doc(&doc);

    let mut found_body: Option<String> = None;
    let mut balloon_count = 0_usize;
    for page in &result.pages {
        for el in &page.elements {
            if let oxidocs_core::layout::LayoutContent::Balloon {
                comment_id, body, ..
            } = &el.content
            {
                if comment_id.starts_with("rprchange:") {
                    balloon_count += 1;
                    found_body = Some(body.clone());
                }
            }
        }
    }
    assert_eq!(
        balloon_count, 1,
        "fixture_24 has 1 rPrChange → 1 balloon"
    );
    let body = found_body.unwrap();
    assert!(
        body.starts_with("Formatted:"),
        "body must start with 'Formatted:'; got {body:?}"
    );
    assert!(
        body.contains("Shadow"),
        "shadow toggle must surface; got {body:?}"
    );
    assert!(
        body.contains("Hidden"),
        "vanish toggle must surface as 'Hidden'; got {body:?}"
    );
    assert!(
        body.contains("Double Strikethrough"),
        "dstrike toggle must surface; got {body:?}"
    );
    assert!(
        body.contains(", "),
        "multi-axis diff must be comma-joined; got {body:?}"
    );
}

// ---------------------------------------------------------------------------
// fixture_11 — CJK body with one ins + one del.
//
// PHASE_2_CLOSEOUT.md known-limitation #5 noted that the existing fixtures are
// Latin-only and the strikethrough Y on CJK glyphs has not been verified.
// fixture_11 is the smallest case that exercises R-01 / R-03 styling on
// MS Mincho 24pt content; it covers the IR / layout side. The actual
// pixel-level Y position is a renderer-side concern (TextOutW + GDI font
// metrics) so verifying it via cargo tests is out of scope — the fixture
// instead pins the prerequisite: the IR carries the right tracked_change
// and the layout pre-pass applies underline/strikethrough + author color
// regardless of script.
// ---------------------------------------------------------------------------

const F11_INS_TEXT: &str = "挿入された文字";
const F11_DEL_TEXT: &str = "削除された文字";

/// fixture_11 parser side: ins + del runs preserve CJK text and tracked-change
/// metadata identically to the Latin fixtures.
#[test]
fn fixture_11_cjk_ins_del_parse_roundtrip() {
    let Some(bytes) = read_fixture("fixture_11_cjk_revisions.docx") else {
        eprintln!("skipping: fixture_11 missing");
        return;
    };
    let doc = oxidocs_core::parse_docx(&bytes).expect("parse fixture_11");

    let revisions: Vec<(String, String)> = collect_runs(&doc)
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
            ("insert".to_string(), F11_INS_TEXT.to_string()),
            ("delete".to_string(), F11_DEL_TEXT.to_string()),
        ],
        "fixture_11 must surface one CJK ins + one CJK del in document order"
    );

    // Author + date metadata still intact.
    let with_tc: Vec<&oxidocs_core::ir::Run> = collect_runs(&doc)
        .into_iter()
        .filter(|r| r.tracked_change.is_some())
        .collect();
    for run in &with_tc {
        let tc = run.tracked_change.as_ref().unwrap();
        assert_eq!(tc.author.as_deref(), Some("Alice Reviewer"));
        assert!(tc.date.is_some(), "w:date must survive on CJK runs too");
    }
}

/// fixture_11 layout side: R-01 styles the CJK ins as underlined Alice-red,
/// R-03 styles the CJK del as struck Alice-red. Adjacent normal CJK runs are
/// left untouched.
///
/// Note: the body layout splits CJK content into per-glyph `LayoutContent::Text`
/// fragments (one element per character; matches the per-glyph TextOutW
/// emission the GDI renderer needs for CJK kerning / spacing). Tests below
/// match on individual CJK characters rather than multi-char substrings.
#[test]
fn fixture_11_cjk_layout_ins_underline_and_del_strikethrough() {
    let Some(bytes) = read_fixture("fixture_11_cjk_revisions.docx") else {
        eprintln!("skipping: fixture_11 missing");
        return;
    };
    let doc = oxidocs_core::parse_docx(&bytes).expect("parse fixture_11");
    let result = layout_doc(&doc);

    // R-01: every CJK glyph that came from the <w:ins> run must be underlined
    // + Alice-red. Pick the first two CJK characters of "挿入された文字".
    for needle in &["挿", "入", "た", "文", "字"] {
        let hits = collect_text_elements_with(&result, needle);
        assert!(
            !hits.is_empty(),
            "ins CJK glyph '{needle}' must reach layout"
        );
        // Glyphs from the ins run must all be underlined Alice-red. The same
        // glyph may appear elsewhere in the doc — filter to the underlined
        // hits and require at least one.
        let ins_styled: Vec<_> = hits
            .iter()
            .filter(|(u, _, c)| *u && c.as_deref() == Some("#D03337"))
            .collect();
        assert!(
            !ins_styled.is_empty(),
            "ins CJK glyph '{needle}' must surface underlined+red at least once; got {hits:?}"
        );
        for (_underline, strike, _color) in &ins_styled {
            assert!(!*strike, "ins CJK glyph '{needle}' must NOT be strikethrough");
        }
    }

    // R-03: every CJK glyph from the <w:del> run must be strikethrough +
    // Alice-red. "削除された文字" — pick characters unique to del so we don't
    // collide with the ins run.
    for needle in &["削", "除"] {
        let hits = collect_text_elements_with(&result, needle);
        assert!(
            !hits.is_empty(),
            "del CJK glyph '{needle}' must reach layout"
        );
        let del_styled: Vec<_> = hits
            .iter()
            .filter(|(_, s, c)| *s && c.as_deref() == Some("#D03337"))
            .collect();
        assert!(
            !del_styled.is_empty(),
            "del CJK glyph '{needle}' must surface struck+red at least once; got {hits:?}"
        );
        for (underline, _strike, _color) in &del_styled {
            assert!(
                !*underline,
                "del CJK glyph '{needle}' must NOT be underlined"
            );
        }
    }

    // Adjacent normal CJK runs must not be touched by the revision pre-pass.
    // "前段落。" precedes the ins run; pick a glyph unique to that prefix.
    for needle in &["前", "段", "落"] {
        let hits = collect_text_elements_with(&result, needle);
        assert!(
            !hits.is_empty(),
            "normal CJK glyph '{needle}' must reach layout"
        );
        for (underline, strike, color) in &hits {
            assert!(!*underline, "normal CJK glyph '{needle}' must not be underlined");
            assert!(!*strike, "normal CJK glyph '{needle}' must not be struck");
            assert_ne!(
                color.as_deref(),
                Some("#D03337"),
                "normal CJK glyph '{needle}' must not be author-tinted"
            );
        }
    }
}

// ---------------------------------------------------------------------------
// fixture_12 — three reviewers, exercising palette slot 2.
//
// Slots 0 (#D03337) and 1 (#5B2C90) are COM-confirmed via fixture_05/06/10.
// Slot 2 in REVISION_AUTHOR_PALETTE is "#2B6033" (Word's documented green —
// also used for moves regardless of author). PHASE_2_CLOSEOUT.md known-
// limitation #9 noted that slots 2-7 lack ground-truth confirmation; the
// fixture below is the smallest input that surfaces slot 2 on the Oxi side
// so a future Word-side pixel pass can sample it.
//
// Body: "Start. ALICE INS middle1 BOB DEL middle2 CAROL INS. End."
// people.xml seeds the palette in author-order so the assignment is stable.
// ---------------------------------------------------------------------------

/// fixture_12 parser side: people.xml seeds three authors and the palette
/// hands them indices 0/1/2 in that order.
#[test]
fn fixture_12_three_reviewers_palette_assigns_slots_0_1_2() {
    let Some(bytes) = read_fixture("fixture_12_three_reviewers.docx") else {
        eprintln!("skipping: fixture_12 missing");
        return;
    };
    let doc = oxidocs_core::parse_docx(&bytes).expect("parse fixture_12");

    assert_eq!(doc.people.len(), 3, "expected 3 reviewers in people.xml");
    let people_authors: Vec<_> = doc.people.iter().map(|p| p.author.as_str()).collect();
    assert_eq!(
        people_authors,
        vec!["Alice Reviewer", "Bob Reviewer", "Carol Reviewer"]
    );

    assert_eq!(doc.authors.len(), 3, "expected 3 palette entries");
    assert_eq!(doc.authors[0].display, "Alice Reviewer");
    assert_eq!(doc.authors[0].color_index, 0);
    assert_eq!(doc.authors[1].display, "Bob Reviewer");
    assert_eq!(doc.authors[1].color_index, 1);
    assert_eq!(doc.authors[2].display, "Carol Reviewer");
    assert_eq!(
        doc.authors[2].color_index, 2,
        "Carol must land on palette slot 2 (the new slot under test)"
    );
}

/// fixture_12 layout side: each author's revision run renders in the
/// palette color for its slot. The third author lands on slot 2 (`#2B6033`),
/// which is currently sourced from Word's documented Office reviewing
/// palette and not yet COM-confirmed against an actual Word render.
/// Asserting against the constant pins the Oxi side; the comment above
/// the test records the open ground-truth question.
#[test]
fn fixture_12_layout_third_author_uses_palette_slot_2() {
    let Some(bytes) = read_fixture("fixture_12_three_reviewers.docx") else {
        eprintln!("skipping: fixture_12 missing");
        return;
    };
    let doc = oxidocs_core::parse_docx(&bytes).expect("parse fixture_12");
    let result = layout_doc(&doc);

    let alice_ins = collect_text_elements_with(&result, "ALICE");
    assert!(!alice_ins.is_empty(), "Alice's ins missing from layout");
    for (underline, _, color) in &alice_ins {
        assert!(*underline);
        assert_eq!(color.as_deref(), Some("#D03337"), "Alice = slot 0");
    }

    let bob_del = collect_text_elements_with(&result, "BOB");
    assert!(!bob_del.is_empty(), "Bob's del missing from layout");
    for (_, strike, color) in &bob_del {
        assert!(*strike);
        assert_eq!(color.as_deref(), Some("#5B2C90"), "Bob = slot 1");
    }

    let carol_ins = collect_text_elements_with(&result, "CAROL");
    assert!(!carol_ins.is_empty(), "Carol's ins missing from layout");
    for (underline, strike, color) in &carol_ins {
        assert!(*underline, "Carol's ins must be underlined");
        assert!(!*strike, "Carol's ins must not be strikethrough");
        assert_eq!(
            color.as_deref(),
            Some("#478103"),
            "Carol = slot 2 (#478103, COM-confirmed R65 2026-04-29 via fixture_12 pixel pass)"
        );
    }
}
