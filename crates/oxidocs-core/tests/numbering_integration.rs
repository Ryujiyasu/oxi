// This Source Code Form is subject to the terms of the Mozilla Public
// License, v. 2.0. If a copy of the MPL was not distributed with this
// file, You can obtain one at https://mozilla.org/MPL/2.0/.

//! Integration tests: parse `<w:numPr>` + `word/numbering.xml` end-to-end
//! and verify `ParagraphStyle.{num_id, num_ilvl, list_marker, list_suff}`
//! populate, including across-paragraph counter state.
//!
//! Parser code path tested: [parser/ooxml.rs:1599](crates/oxidocs-core/src/parser/ooxml.rs#L1599)
//! resolves `numId + ilvl` against the numbering definitions, advances
//! per-(numId, ilvl) counters via `resolve_marker_full`, and writes the
//! resolved marker text + suffix back onto ParagraphStyle.
//!
//! Fixtures live in `tools/fixtures/numbering_samples/` and are authored
//! by `tools/metrics/build_numbering_repro_fixtures.py` (S279).

use std::fs;

use oxidocs_core::ir::{Block, Document, Paragraph};
use oxidocs_core::parser::parse_docx;

fn fixture_path(name: &str) -> std::path::PathBuf {
    let crate_root = std::env::current_dir().unwrap();
    let workspace_root = crate_root.parent().unwrap().parent().unwrap();
    workspace_root.join("tools").join("fixtures").join("numbering_samples").join(name)
}

fn collect_paras(doc: &Document) -> Vec<&Paragraph> {
    doc.pages.iter()
        .flat_map(|p| p.blocks.iter())
        .filter_map(|b| if let Block::Paragraph(p) = b { Some(p) } else { None })
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
fn v1_decimal_advances_counter_across_three_paragraphs() {
    // numId=1, ilvl=0, numFmt=decimal, lvlText="%1."
    // 3 paragraphs → markers "1.", "2.", "3.".
    let Some(doc) = load("v1_decimal.docx") else { return };
    let paras = collect_paras(&doc);
    assert_eq!(paras.len(), 3);

    let expected = [("1.", "Alpha"), ("2.", "Bravo"), ("3.", "Charlie")];
    for (i, (exp_marker, exp_text)) in expected.iter().enumerate() {
        let p = paras[i];
        assert_eq!(p.style.num_id.as_deref(), Some("1"));
        assert_eq!(p.style.num_ilvl, 0);
        assert_eq!(p.style.list_marker.as_deref(), Some(*exp_marker),
            "paragraph {} marker", i);
        // Default suffix is "tab".
        assert_eq!(p.style.list_suff.as_deref(), Some("tab"));
        let text: String = p.runs.iter().map(|r| r.text.as_str()).collect();
        assert_eq!(text, *exp_text);
    }
}

#[test]
fn v1_bullet_uses_lvltext_literal_for_every_paragraph() {
    // numFmt=bullet, lvlText="•" — each paragraph gets the SAME marker
    // text (no counter advancement for bullets).
    let Some(doc) = load("v1_bullet.docx") else { return };
    let paras = collect_paras(&doc);
    assert_eq!(paras.len(), 3);

    for p in &paras {
        assert_eq!(p.style.num_id.as_deref(), Some("1"));
        assert_eq!(p.style.num_ilvl, 0);
        assert_eq!(p.style.list_marker.as_deref(), Some("•"));
    }
}

#[test]
fn v1_two_levels_uses_distinct_lvltext_per_ilvl() {
    // Same numId, two levels:
    //   ilvl=0 decimal lvlText="%1." → "1.", "2."
    //   ilvl=1 lowerLetter lvlText="%2)" → "a)", "b)"
    let Some(doc) = load("v1_two_levels.docx") else { return };
    let paras = collect_paras(&doc);
    assert_eq!(paras.len(), 4);

    let cases: &[(u8, &str, &str)] = &[
        (0, "1.", "Top 1"),
        (1, "a)", "Nested 1a"),
        (1, "b)", "Nested 1b"),
        (0, "2.", "Top 2"),
    ];
    for (i, (exp_ilvl, exp_marker, exp_text)) in cases.iter().enumerate() {
        let p = paras[i];
        assert_eq!(p.style.num_ilvl, *exp_ilvl, "paragraph {} ilvl", i);
        assert_eq!(p.style.list_marker.as_deref(), Some(*exp_marker),
            "paragraph {} marker", i);
        let text: String = p.runs.iter().map(|r| r.text.as_str()).collect();
        assert_eq!(text, *exp_text);
    }
}

#[test]
fn v1_two_numids_maintain_independent_counters() {
    // numId=1 and numId=2 each have their own ilvl=0 counter starting at 1.
    // Interleaving paragraphs across numIds must NOT cross-pollute counters.
    let Some(doc) = load("v1_two_numids.docx") else { return };
    let paras = collect_paras(&doc);
    assert_eq!(paras.len(), 5);

    let cases: &[(&str, &str, &str)] = &[
        ("1", "1.", "List A first"),
        ("1", "2.", "List A second"),
        ("2", "1.", "List B first"),
        ("2", "2.", "List B second"),
        ("1", "3.", "List A third"),
    ];
    for (i, (exp_num_id, exp_marker, exp_text)) in cases.iter().enumerate() {
        let p = paras[i];
        assert_eq!(p.style.num_id.as_deref(), Some(*exp_num_id),
            "paragraph {} num_id", i);
        assert_eq!(p.style.list_marker.as_deref(), Some(*exp_marker),
            "paragraph {} marker (independent-counter advancement)", i);
        let text: String = p.runs.iter().map(|r| r.text.as_str()).collect();
        assert_eq!(text, *exp_text);
    }
}

#[test]
fn all_four_fixtures_parse_with_expected_paragraph_counts() {
    let cases: &[(&str, usize)] = &[
        ("v1_decimal.docx",     3),
        ("v1_bullet.docx",      3),
        ("v1_two_levels.docx",  4),
        ("v1_two_numids.docx",  5),
    ];
    for (name, expected) in cases {
        let path = fixture_path(name);
        if !path.exists() {
            eprintln!("skipping {}", name);
            continue;
        }
        let data = fs::read(&path).unwrap();
        let doc = parse_docx(&data)
            .unwrap_or_else(|e| panic!("failed to parse {}: {:?}", name, e));
        let n = collect_paras(&doc).len();
        assert_eq!(n, *expected, "{} paragraph count", name);
        // Every paragraph should carry a num_id.
        for p in collect_paras(&doc) {
            assert!(p.style.num_id.is_some(),
                "{} every paragraph should carry num_id", name);
            assert!(p.style.list_marker.is_some(),
                "{} every paragraph should resolve a list_marker", name);
        }
    }
}
