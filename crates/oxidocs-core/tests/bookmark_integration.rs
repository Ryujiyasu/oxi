//! Integration tests: parse `<w:bookmarkStart>` / `<w:bookmarkEnd>` and verify
//! the `Run.bookmark_name` anchor-run shape.
//!
//! Parser code path tested: `parser/ooxml.rs:1341` (the `"bookmarkStart"`
//! arm) — creates an empty anchor `Run` with `bookmark_name=Some(name)`,
//! filters Word's auto-inserted `_GoBack` cursor marker, and treats
//! `<w:bookmarkEnd>` as a no-op (anchor is placed at start). Companion to
//! `hyperlink_integration.rs` which already exercises the *consumer* side
//! of bookmarks via `<w:hyperlink w:anchor="bookmark1">`.
//!
//! Fixtures live in `tools/fixtures/bookmark_samples/` and are authored by
//! `tools/metrics/build_bookmark_repro_fixtures.py` (S274).

use std::fs;

use oxidocs_core::ir::{Block, Document, Run};
use oxidocs_core::parser::parse_docx;

fn fixture_path(name: &str) -> std::path::PathBuf {
    let crate_root = std::env::current_dir().unwrap();
    let workspace_root = crate_root.parent().unwrap().parent().unwrap();
    workspace_root.join("tools").join("fixtures").join("bookmark_samples").join(name)
}

fn collect_runs(doc: &Document) -> Vec<&Run> {
    doc.pages.iter()
        .flat_map(|p| p.blocks.iter())
        .filter_map(|b| if let Block::Paragraph(p) = b { Some(p) } else { None })
        .flat_map(|p| p.runs.iter())
        .collect()
}

fn load(name: &str) -> Option<Document> {
    let path = fixture_path(name);
    if !path.exists() {
        eprintln!("skipping: {} not found", path.display());
        return None;
    }
    let data = fs::read(&path).expect("read fixture");
    Some(parse_docx(&data).expect("parse fixture"))
}

#[test]
fn v1_basic_bookmark_produces_empty_anchor_run_before_text() {
    // <w:bookmarkStart w:name="section1"/> + text. Parser places an empty
    // anchor run (text=""), then a content run carrying the visible text.
    let Some(doc) = load("v1_basic.docx") else { return };
    let runs = collect_runs(&doc);
    assert_eq!(runs.len(), 2, "expected 1 anchor run + 1 content run");

    assert_eq!(runs[0].text, "");
    assert_eq!(runs[0].bookmark_name.as_deref(), Some("section1"));

    assert_eq!(runs[1].text, "Hello, world.");
    assert!(runs[1].bookmark_name.is_none());
}

#[test]
fn v1_goback_bookmark_is_filtered_out_of_ir() {
    // Word inserts `_GoBack` to remember the cursor on save. It MUST NOT
    // produce an anchor run — only the deliberately-authored "real_anchor"
    // should survive. Pins the explicit filter at parser/ooxml.rs:1347.
    let Some(doc) = load("v1_goback_skipped.docx") else { return };
    let runs = collect_runs(&doc);

    // No run should ever carry `_GoBack`.
    for r in &runs {
        assert_ne!(r.bookmark_name.as_deref(), Some("_GoBack"),
            "_GoBack should be filtered, never appear in IR");
    }

    // Exactly one bookmark anchor (the real one).
    let anchors: Vec<&Run> = runs.iter().filter(|r| r.bookmark_name.is_some()).copied().collect();
    assert_eq!(anchors.len(), 1, "only the non-_GoBack anchor survives");
    assert_eq!(anchors[0].bookmark_name.as_deref(), Some("real_anchor"));
    assert_eq!(anchors[0].text, "", "anchor run carries no visible text");
}

#[test]
fn v1_multiple_bookmarks_preserve_document_order() {
    // Three bookmarks intro / body / end interleaved with content runs.
    // Document order matters — consumers iterate runs to map anchor positions
    // to surrounding text.
    let Some(doc) = load("v1_multiple.docx") else { return };
    let runs = collect_runs(&doc);
    assert_eq!(runs.len(), 6);

    let bookmark_seq: Vec<&str> = runs.iter()
        .filter_map(|r| r.bookmark_name.as_deref())
        .collect();
    assert_eq!(bookmark_seq, vec!["intro", "body", "end"]);

    // Sanity: every anchor run is empty; content runs are non-empty.
    for r in &runs {
        if r.bookmark_name.is_some() {
            assert_eq!(r.text, "", "anchor run text must be empty");
        } else {
            assert!(!r.text.is_empty(), "content run must have text");
        }
    }
}

#[test]
fn v1_bookmark_around_text_anchors_at_start_not_end() {
    // bookmarkStart precedes the run, bookmarkEnd follows it. The parser
    // collapses this to an empty anchor at the START position; the text run
    // inside the wrap is preserved verbatim with bookmark_name=None.
    // There is no "end anchor" run — bookmarkEnd is a no-op
    // (parser/ooxml.rs:1373).
    let Some(doc) = load("v1_around_text.docx") else { return };
    let runs = collect_runs(&doc);
    assert_eq!(runs.len(), 4, "before + anchor + inside + after");

    assert_eq!(runs[0].text, "Before. ");
    assert!(runs[0].bookmark_name.is_none());

    assert_eq!(runs[1].text, "");
    assert_eq!(runs[1].bookmark_name.as_deref(), Some("wrap"));

    assert_eq!(runs[2].text, "Inside wrapped span.");
    assert!(runs[2].bookmark_name.is_none(),
        "wrapped text must NOT inherit bookmark_name");

    assert_eq!(runs[3].text, " After.");
    assert!(runs[3].bookmark_name.is_none());
}

#[test]
fn all_four_fixtures_parse_and_produce_expected_anchor_counts() {
    // Smoke test: 1 / 1 / 3 / 1 anchor runs respectively.
    let cases: &[(&str, usize)] = &[
        ("v1_basic.docx",          1),
        ("v1_goback_skipped.docx", 1),
        ("v1_multiple.docx",       3),
        ("v1_around_text.docx",    1),
    ];
    for (name, exp_anchors) in cases {
        let path = fixture_path(name);
        if !path.exists() {
            eprintln!("skipping {}", name);
            continue;
        }
        let data = fs::read(&path).unwrap();
        let doc = parse_docx(&data)
            .unwrap_or_else(|e| panic!("failed to parse {}: {:?}", name, e));
        let n = collect_runs(&doc).iter()
            .filter(|r| r.bookmark_name.is_some())
            .count();
        assert_eq!(n, *exp_anchors, "{} anchor count", name);
    }
}
