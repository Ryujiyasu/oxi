//! Integration tests: parse `<w:tabs>` / `<w:tab>` elements end-to-end
//! and verify `Paragraph.style.tab_stops` populates with the expected
//! position, alignment, and leader.
//!
//! Parser code path tested: [parser/ooxml.rs:2422](crates/oxidocs-core/src/parser/ooxml.rs#L2422)
//! reads each `<w:tab>` inside `<w:tabs>` and maps:
//!   - `w:pos` (twips, divided by 20 → pt)
//!   - `w:val` → TabStopAlignment (`left`/`center`/`right`/`end`/`decimal`)
//!   - `w:leader` → Option<String> ("none" → None, anything else → Some)
//!
//! Fixtures live in `tools/fixtures/tabstops_samples/` and are authored
//! by `tools/metrics/build_tabstops_repro_fixtures.py` (S291).

use std::fs;

use oxidocs_core::ir::{Block, Document, Paragraph, TabStopAlignment};
use oxidocs_core::parser::parse_docx;

fn fixture_path(name: &str) -> std::path::PathBuf {
    let crate_root = std::env::current_dir().unwrap();
    let workspace_root = crate_root.parent().unwrap().parent().unwrap();
    workspace_root.join("tools").join("fixtures").join("tabstops_samples").join(name)
}

fn first_paragraph(doc: &Document) -> &Paragraph {
    doc.pages.iter()
        .flat_map(|p| p.blocks.iter())
        .find_map(|b| if let Block::Paragraph(p) = b { Some(p) } else { None })
        .expect("first paragraph")
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
fn v1_simple_has_single_left_tab_at_200pt() {
    let Some(doc) = load("v1_simple.docx") else { return };
    let p = first_paragraph(&doc);
    assert_eq!(p.style.tab_stops.len(), 1, "single tab stop");
    let t = &p.style.tab_stops[0];
    assert!((t.position - 200.0).abs() < 0.05, "pos {}", t.position);
    assert_eq!(t.alignment, TabStopAlignment::Left);
    assert!(t.leader.is_none(), "no leader on simple tab");
}

#[test]
fn v1_multi_carries_left_center_right_decimal_in_order() {
    let Some(doc) = load("v1_multi.docx") else { return };
    let p = first_paragraph(&doc);
    assert_eq!(p.style.tab_stops.len(), 4);

    let expected: &[(f32, TabStopAlignment)] = &[
        (100.0, TabStopAlignment::Left),
        (200.0, TabStopAlignment::Center),
        (300.0, TabStopAlignment::Right),
        (400.0, TabStopAlignment::Decimal),
    ];
    for (i, (exp_pos, exp_align)) in expected.iter().enumerate() {
        let t = &p.style.tab_stops[i];
        assert!((t.position - exp_pos).abs() < 0.05,
            "tab[{}] pos {} expected {}", i, t.position, exp_pos);
        assert_eq!(t.alignment, *exp_align, "tab[{}] alignment", i);
    }
}

#[test]
fn v1_with_leader_captures_dot_leader() {
    // TOC-style: `<w:tab w:val="right" w:pos="8000" w:leader="dot"/>`.
    // The parser stores the leader verbatim ("dot"); the renderer is
    // responsible for mapping "dot"/"hyphen"/"underscore"/"middleDot"/"heavy"
    // to the visual fill character.
    let Some(doc) = load("v1_with_leader.docx") else { return };
    let p = first_paragraph(&doc);
    assert_eq!(p.style.tab_stops.len(), 1);
    let t = &p.style.tab_stops[0];
    assert_eq!(t.alignment, TabStopAlignment::Right);
    assert_eq!(t.leader.as_deref(), Some("dot"),
        "leader string preserved verbatim");
}

#[test]
fn v1_no_tabs_leaves_tab_stops_empty() {
    // Paragraph without `<w:tabs>` → tab_stops vec is empty.
    let Some(doc) = load("v1_no_tabs.docx") else { return };
    let p = first_paragraph(&doc);
    assert!(p.style.tab_stops.is_empty(),
        "no-tabs paragraph yields empty tab_stops");
}

#[test]
fn all_four_fixtures_parse_with_expected_tab_count() {
    let cases: &[(&str, usize)] = &[
        ("v1_simple.docx",      1),
        ("v1_multi.docx",       4),
        ("v1_with_leader.docx", 1),
        ("v1_no_tabs.docx",     0),
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
        let p = first_paragraph(&doc);
        assert_eq!(p.style.tab_stops.len(), *expected, "{} tab count", name);
    }
}
