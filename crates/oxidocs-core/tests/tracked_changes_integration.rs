//! Tracked-change IR deepening tests (ins / del / moveFrom / moveTo).
//!
//! comments_fixtures.rs already exercises the parser for these four wrappers
//! but its assertions stop at "author is Some / date is Some / change_type
//! matches". This file pins the **exact** author strings, ISO-8601 dates, and
//! per-wrapper `w:id` values that the existing fixtures encode — catching
//! silent regressions where attributes are dropped, swapped, or normalized.
//!
//! Parser code paths exercised:
//! - [parser/ooxml.rs:1101](crates/oxidocs-core/src/parser/ooxml.rs#L1101):
//!   `"ins" | "del" | "moveFrom" | "moveTo"` wrapper dispatch.
//! - [parser/ooxml.rs:1112-1121](crates/oxidocs-core/src/parser/ooxml.rs#L1112):
//!   `w:author` / `w:date` / `w:id` attribute extraction into TrackedChange.
//!
//! Fixtures live in `tools/fixtures/comments_samples/` and are authored by
//! `tools/fixtures/comments_samples/build_comments_samples.py`.

use std::path::{Path, PathBuf};

use oxidocs_core::ir::{Block, Document, Run};

fn fixture(name: &str) -> PathBuf {
    let manifest_dir = Path::new(env!("CARGO_MANIFEST_DIR"));
    manifest_dir.join("../../tools/fixtures/comments_samples").join(name)
}

fn read_fixture(name: &str) -> Option<Vec<u8>> {
    std::fs::read(fixture(name)).ok()
}

fn collect_runs(doc: &Document) -> Vec<&Run> {
    let mut runs = Vec::new();
    for page in &doc.pages {
        for block in &page.blocks {
            if let Block::Paragraph(p) = block {
                for run in &p.runs {
                    runs.push(run);
                }
            }
        }
    }
    runs
}

/// fixture_05 (`<w:ins w:id="100" w:author="Alice Reviewer"
/// w:date="2026-04-18T10:00:00Z">`) — pin the exact author string, ISO-8601
/// date, and wrapper `w:id` that map onto `TrackedChange`. The companion
/// test in `comments_fixtures.rs` only asserts `is_some`, so a parser
/// regression that dropped `w:date` while keeping `w:author` (or vice
/// versa) would slip through.
#[test]
fn fixture_05_single_ins_pins_exact_author_date_and_id() {
    let Some(bytes) = read_fixture("fixture_05_single_ins.docx") else {
        eprintln!("skipping: fixture_05 missing");
        return;
    };
    let doc = oxidocs_core::parse_docx(&bytes).expect("parse fixture_05");

    let ins_runs: Vec<_> = collect_runs(&doc)
        .into_iter()
        .filter(|r| {
            r.tracked_change.as_ref().map(|t| t.change_type.as_str()) == Some("insert")
        })
        .collect();
    assert_eq!(ins_runs.len(), 1, "exactly one <w:ins> run");
    let tc = ins_runs[0].tracked_change.as_ref().unwrap();

    // Exact pinning: any change to author/date/id parsing flips one of these.
    assert_eq!(tc.author.as_deref(), Some("Alice Reviewer"));
    assert_eq!(tc.date.as_deref(), Some("2026-04-18T10:00:00Z"));
    assert_eq!(tc.pair_id.as_deref(), Some("100"));
    assert_eq!(tc.change_type, "insert",
        "change_type literal pinned (parser maps \"ins\" -> \"insert\")");
}

/// fixture_07 has three revisions with distinct `w:id` (100, 101, 102) and
/// distinct `w:date` (10:00:00Z for ins/del, 10:05:00Z for the second ins).
/// Pin the exact (change_type, pair_id, date, text) tuple per revision **in
/// XML order** — catches reorder / id-swap / date-swap regressions that the
/// `change_type + text` companion check would miss.
#[test]
fn fixture_07_mixed_pins_per_revision_id_and_date_tuples() {
    let Some(bytes) = read_fixture("fixture_07_mixed_ins_del.docx") else {
        eprintln!("skipping: fixture_07 missing");
        return;
    };
    let doc = oxidocs_core::parse_docx(&bytes).expect("parse fixture_07");

    let revs: Vec<_> = collect_runs(&doc)
        .into_iter()
        .filter_map(|r| {
            r.tracked_change.as_ref().map(|t| {
                (
                    t.change_type.clone(),
                    t.pair_id.clone().unwrap_or_default(),
                    t.date.clone().unwrap_or_default(),
                    r.text.clone(),
                )
            })
        })
        .collect();

    assert_eq!(
        revs,
        vec![
            ("insert".into(), "100".into(), "2026-04-18T10:00:00Z".into(), "ins1 ".into()),
            ("delete".into(), "101".into(), "2026-04-18T10:00:00Z".into(), "del1 ".into()),
            ("insert".into(), "102".into(), "2026-04-18T10:05:00Z".into(), "ins2".into()),
        ],
        "revisions must surface in XML order with exact id/date/text triples"
    );
}

/// fixture_08 wraps "moved clause" in BOTH `<w:moveFrom w:id="201">` and
/// `<w:moveTo w:id="202">`. The companion test in `comments_fixtures.rs`
/// only asserts `pair_id.is_some()` for each side; this test pins the
/// **wrapper-local** ids (201 ≠ 202) and the **shared** range-pair id (200
/// on `moveFromRangeStart` / `moveToRangeStart`).
///
/// Note (revisions_notes.md §1.2): per-wrapper `w:id` on moveFrom/moveTo
/// is NOT the from↔to pairing key — that lives on the surrounding range
/// markers via `w:name="move1"` + range-marker `w:id=200`. This test pins
/// the wrapper-side ids only (200 is captured via range-marker handling, a
/// separate code path not asserted here).
#[test]
fn fixture_08_move_wrappers_have_distinct_per_side_ids() {
    let Some(bytes) = read_fixture("fixture_08_move_from_to.docx") else {
        eprintln!("skipping: fixture_08 missing");
        return;
    };
    let doc = oxidocs_core::parse_docx(&bytes).expect("parse fixture_08");

    let mut from_ids: Vec<String> = Vec::new();
    let mut to_ids: Vec<String> = Vec::new();
    for r in collect_runs(&doc) {
        if let Some(t) = &r.tracked_change {
            match t.change_type.as_str() {
                "moveFrom" => {
                    if let Some(id) = &t.pair_id { from_ids.push(id.clone()); }
                }
                "moveTo" => {
                    if let Some(id) = &t.pair_id { to_ids.push(id.clone()); }
                }
                _ => {}
            }
        }
    }
    assert_eq!(from_ids, vec!["201".to_string()],
        "moveFrom wrapper id pinned to 201 (per-side, not the shared range id)");
    assert_eq!(to_ids, vec!["202".to_string()],
        "moveTo wrapper id pinned to 202 (per-side, not the shared range id)");
    assert_ne!(from_ids, to_ids,
        "moveFrom and moveTo wrappers each carry their own id; they MUST differ");
}

/// fixture_10 has Alice's `<w:ins>` and Bob's `<w:del>` in the same
/// paragraph. The companion `fixture_10_people_two_reviewers` test exercises
/// the people.xml palette but does not pin **per-revision author
/// attribution** — i.e., that the parser doesn't accidentally cross-wire
/// author strings between adjacent ins/del wrappers.
#[test]
fn fixture_10_multi_reviewer_attribution_does_not_cross_wires() {
    let Some(bytes) = read_fixture("fixture_10_multiple_reviewers.docx") else {
        eprintln!("skipping: fixture_10 missing");
        return;
    };
    let doc = oxidocs_core::parse_docx(&bytes).expect("parse fixture_10");

    let by_kind: std::collections::HashMap<&str, (String, String)> = collect_runs(&doc)
        .into_iter()
        .filter_map(|r| {
            r.tracked_change.as_ref().map(|t| (
                match t.change_type.as_str() {
                    "insert" => "insert",
                    "delete" => "delete",
                    _ => "other",
                },
                (
                    t.author.clone().unwrap_or_default(),
                    t.pair_id.clone().unwrap_or_default(),
                ),
            ))
        })
        .collect();

    let (alice_author, alice_id) = by_kind.get("insert").expect("ins revision present");
    assert_eq!(alice_author, "Alice Reviewer", "ins -> Alice");
    assert_eq!(alice_id, "400", "ins wrapper id pinned");

    let (bob_author, bob_id) = by_kind.get("delete").expect("del revision present");
    assert_eq!(bob_author, "Bob Reviewer", "del -> Bob (NOT cross-wired to Alice)");
    assert_eq!(bob_id, "401", "del wrapper id pinned");
}

/// fixture_11 contains CJK ins ("挿入された文字") + CJK del ("削除された文字")
/// with `xml:space="preserve"`. Pin that the parser preserves CJK glyphs
/// **byte-for-byte** in both `Run.text` (for ins via `<w:t>`) and `Run.text`
/// (for del via `<w:delText>`). Any future normalization (NFC/NFD swap,
/// halfwidth-fullwidth conversion, whitespace collapse) would flip this.
#[test]
fn fixture_11_cjk_ins_del_text_preserved_byte_exact() {
    let Some(bytes) = read_fixture("fixture_11_cjk_revisions.docx") else {
        eprintln!("skipping: fixture_11 missing");
        return;
    };
    let doc = oxidocs_core::parse_docx(&bytes).expect("parse fixture_11");

    let ins_text: Vec<_> = collect_runs(&doc)
        .into_iter()
        .filter(|r| {
            r.tracked_change.as_ref().map(|t| t.change_type.as_str()) == Some("insert")
        })
        .map(|r| r.text.clone())
        .collect();
    let del_text: Vec<_> = collect_runs(&doc)
        .into_iter()
        .filter(|r| {
            r.tracked_change.as_ref().map(|t| t.change_type.as_str()) == Some("delete")
        })
        .map(|r| r.text.clone())
        .collect();

    assert_eq!(ins_text, vec!["挿入された文字".to_string()],
        "CJK ins text preserved verbatim (no normalization)");
    assert_eq!(del_text, vec!["削除された文字".to_string()],
        "CJK del text from <w:delText> preserved verbatim");

    // Byte-length check: 7 CJK chars × 3 bytes UTF-8 = 21 bytes each.
    assert_eq!(ins_text[0].len(), 21, "7 CJK chars × 3 bytes UTF-8");
    assert_eq!(del_text[0].len(), 21, "7 CJK chars × 3 bytes UTF-8");
}
