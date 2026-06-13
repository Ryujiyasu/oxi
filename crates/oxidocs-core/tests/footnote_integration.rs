// This Source Code Form is subject to the terms of the Mozilla Public
// License, v. 2.0. If a copy of the MPL was not distributed with this
// file, You can obtain one at https://mozilla.org/MPL/2.0/.

//! Integration tests: parse `<w:footnoteReference>` / `<w:endnoteReference>`
//! elements end-to-end and verify the `Run.footnote_ref` / `Run.endnote_ref`
//! fields are populated, AND that the per-section renumber pass
//! (`renumber_note_refs` at `parser/ooxml.rs:5907`) rewrites `run.text` to a
//! sequential "1","2","3",... regardless of the raw XML `w:id` (which Word
//! permits to have gaps and to start at 2 since id=1 is the separator).
//!
//! Fixtures live in `tools/fixtures/footnote_samples/` and are authored by
//! `tools/metrics/build_footnote_endnote_repro_fixtures.py` (S273, written
//! direct-to-fixtures per the S272 hyperlink no-COM-needed variant).
//!
//! Companion to `hyperlink_integration.rs`, `vertical_integration.rs`,
//! `ruby_integration.rs`, `omml_integration.rs`, `comments_fixtures.rs`.

use std::fs;

use oxidocs_core::ir::{Block, Document, Run};
use oxidocs_core::parser::parse_docx;

fn fixture_path(name: &str) -> std::path::PathBuf {
    let crate_root = std::env::current_dir().unwrap();
    let workspace_root = crate_root.parent().unwrap().parent().unwrap();
    workspace_root.join("tools").join("fixtures").join("footnote_samples").join(name)
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
fn v1_single_footnote_populates_ref_and_renders_sequence_one() {
    // <w:footnoteReference w:id="1"/> inside a body run. Parser should produce
    // a 3-run paragraph (plain + footnote-marker + plain) and the marker run
    // carries footnote_ref=Some(1), text="1".
    let Some(doc) = load("v1_footnote.docx") else { return };
    let runs = collect_runs(&doc);
    assert_eq!(runs.len(), 3);
    assert_eq!(runs[0].text, "Body text with footnote");
    assert!(runs[0].footnote_ref.is_none());
    assert_eq!(runs[1].text, "1");
    assert_eq!(runs[1].footnote_ref, Some(1));
    assert!(runs[1].endnote_ref.is_none());
    assert_eq!(runs[2].text, " and more body text.");
    assert!(runs[2].footnote_ref.is_none());
}

#[test]
fn v1_single_endnote_populates_ref_and_renders_sequence_one() {
    // Symmetric to the footnote case. Even though the parser initially writes
    // run.text = "[1]" at parse time, renumber_note_refs overwrites that to
    // the plain sequence number "1" once the section is fully parsed.
    let Some(doc) = load("v1_endnote.docx") else { return };
    let runs = collect_runs(&doc);
    assert_eq!(runs.len(), 3);
    assert_eq!(runs[1].text, "1");
    assert_eq!(runs[1].endnote_ref, Some(1));
    assert!(runs[1].footnote_ref.is_none());
}

#[test]
fn v1_mixed_paragraph_carries_both_note_kinds_independently() {
    // One paragraph: plain + footnote(1) + plain + endnote(1) + plain.
    // Footnote and endnote numbering are INDEPENDENT — both can be "1" in
    // the same paragraph; the field that's Some distinguishes which kind.
    let Some(doc) = load("v1_mixed.docx") else { return };
    let runs = collect_runs(&doc);
    assert_eq!(runs.len(), 5);

    let texts: Vec<&str> = runs.iter().map(|r| r.text.as_str()).collect();
    assert_eq!(texts, vec!["See footnote", "1", " and endnote", "1", "."]);

    assert_eq!(runs[1].footnote_ref, Some(1));
    assert!(runs[1].endnote_ref.is_none());
    assert!(runs[3].footnote_ref.is_none());
    assert_eq!(runs[3].endnote_ref, Some(1));
}

#[test]
fn v1_renumber_preserves_raw_id_but_renders_per_section_sequence() {
    // Two footnoteReferences with raw ids 2 and 5 (gaps are legal in OOXML).
    // Word displays them as "1" and "2" — verifies renumber_note_refs at
    // parser/ooxml.rs:5907 rewrites run.text without touching footnote_ref
    // (which keeps the raw XML id, useful for round-tripping or correlating
    // back to the footnotes.xml entries).
    let Some(doc) = load("v1_renumber.docx") else { return };
    let runs = collect_runs(&doc);
    assert_eq!(runs.len(), 5);

    assert_eq!(runs[1].text, "1", "first footnote marker should render as seq=1");
    assert_eq!(runs[1].footnote_ref, Some(2), "raw id=2 preserved");

    assert_eq!(runs[3].text, "2", "second footnote marker should render as seq=2");
    assert_eq!(runs[3].footnote_ref, Some(5), "raw id=5 preserved");
}

#[test]
fn all_four_fixtures_parse_with_expected_note_refs() {
    // Smoke test: each fixture parses, and the per-fixture expected
    // footnote/endnote counts match. Catches regressions where the parser
    // would crash on missing footnotes.xml or drop refs entirely.
    let cases: &[(&str, usize, usize)] = &[
        ("v1_footnote.docx", 1, 0),
        ("v1_endnote.docx",  0, 1),
        ("v1_mixed.docx",    1, 1),
        ("v1_renumber.docx", 2, 0),
    ];
    for (name, exp_fn, exp_en) in cases {
        let path = fixture_path(name);
        if !path.exists() {
            eprintln!("skipping {}", name);
            continue;
        }
        let data = fs::read(&path).unwrap();
        let doc = parse_docx(&data)
            .unwrap_or_else(|e| panic!("failed to parse {}: {:?}", name, e));
        let runs = collect_runs(&doc);
        let n_fn = runs.iter().filter(|r| r.footnote_ref.is_some()).count();
        let n_en = runs.iter().filter(|r| r.endnote_ref.is_some()).count();
        assert_eq!(n_fn, *exp_fn, "{} footnote count", name);
        assert_eq!(n_en, *exp_en, "{} endnote count", name);
    }
}
