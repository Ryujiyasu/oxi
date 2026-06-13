// This Source Code Form is subject to the terms of the Mozilla Public
// License, v. 2.0. If a copy of the MPL was not distributed with this
// file, You can obtain one at https://mozilla.org/MPL/2.0/.

//! Integration tests: parse `<w:fldChar>` + `<w:instrText>` field
//! groups end-to-end and verify `Run.field_type` populates correctly.
//!
//! Parser code path tested: [parser/ooxml.rs:2656](crates/oxidocs-core/src/parser/ooxml.rs#L2656)
//! inspects the `<w:instrText>` content of a complete field (delimited by
//! `<w:fldChar fldCharType="begin"/separate/end"/>`) and:
//!   - PAGE      → `Run.field_type = Some(FieldType::Page)`,     text rewritten to `"#"`
//!   - NUMPAGES  → `Run.field_type = Some(FieldType::NumPages)`, text rewritten to `"#"`
//!   - DATE/TIME → text rewritten to the field-code string, no field_type
//!   - TOC/HYPERLINK/REF/AUTHOR-class → text-only handling, no field_type
//!
//! Fixtures live in `tools/fixtures/field_samples/` and are authored by
//! `tools/metrics/build_field_repro_fixtures.py` (S278).

use std::fs;

use oxidocs_core::ir::{Block, Document, FieldType, Run};
use oxidocs_core::parser::parse_docx;

fn fixture_path(name: &str) -> std::path::PathBuf {
    let crate_root = std::env::current_dir().unwrap();
    let workspace_root = crate_root.parent().unwrap().parent().unwrap();
    workspace_root.join("tools").join("fixtures").join("field_samples").join(name)
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
fn v1_page_field_carries_page_variant_and_hash_placeholder() {
    // Body: "Page " + <PAGE field> + " of total."
    // Parser collapses the begin/separate/end markers; the instrText run
    // becomes the field-carrying run with text rewritten to "#" and
    // field_type = Some(Page).
    let Some(doc) = load("v1_page.docx") else { return };
    let runs = collect_runs(&doc);
    let field_runs: Vec<&Run> = runs.iter()
        .filter(|r| r.field_type.is_some())
        .copied()
        .collect();
    assert_eq!(field_runs.len(), 1, "expected exactly 1 field-bearing run");
    assert_eq!(field_runs[0].field_type, Some(FieldType::Page));
    assert_eq!(field_runs[0].text, "#",
        "PAGE field rewrites text to placeholder '#'");
    // Surrounding visible text preserved.
    let visible: String = runs.iter()
        .filter(|r| !r.text.is_empty() && r.field_type.is_none())
        .map(|r| r.text.as_str())
        .collect();
    assert_eq!(visible, "Page  of total.");
}

#[test]
fn v1_numpages_field_carries_numpages_variant() {
    // Body: "Total: " + <NUMPAGES field> + " pages."
    let Some(doc) = load("v1_numpages.docx") else { return };
    let runs = collect_runs(&doc);
    let field_runs: Vec<&Run> = runs.iter()
        .filter(|r| r.field_type.is_some())
        .copied()
        .collect();
    assert_eq!(field_runs.len(), 1);
    assert_eq!(field_runs[0].field_type, Some(FieldType::NumPages));
    assert_eq!(field_runs[0].text, "#",
        "NUMPAGES field rewrites text to placeholder '#'");
}

#[test]
fn v1_page_of_numpages_carries_both_variants_in_order() {
    // Common header pattern: "Page # of #." with PAGE then NUMPAGES.
    // Pins that both fields coexist independently in one paragraph and
    // the order is preserved.
    let Some(doc) = load("v1_page_of_numpages.docx") else { return };
    let runs = collect_runs(&doc);
    let field_runs: Vec<&Run> = runs.iter()
        .filter(|r| r.field_type.is_some())
        .copied()
        .collect();
    assert_eq!(field_runs.len(), 2,
        "should have 1 PAGE and 1 NUMPAGES run");
    assert_eq!(field_runs[0].field_type, Some(FieldType::Page));
    assert_eq!(field_runs[1].field_type, Some(FieldType::NumPages));
    // Visible text between fields preserved.
    let visible: String = runs.iter()
        .filter(|r| !r.text.is_empty() && r.field_type.is_none())
        .map(|r| r.text.as_str())
        .collect();
    assert_eq!(visible, "Page  of .");
}

#[test]
fn v1_date_field_does_not_carry_field_type_variant() {
    // DATE is handled (text rewritten to the field-code string) but the
    // FieldType enum currently only has Page/NumPages variants. So DATE
    // produces no field_type, just text rewriting. Pins the current
    // behavior so a future FieldType::Date variant can't silently regress.
    let Some(doc) = load("v1_date.docx") else { return };
    let runs = collect_runs(&doc);
    let with_ft: Vec<&Run> = runs.iter()
        .filter(|r| r.field_type.is_some())
        .copied()
        .collect();
    assert!(with_ft.is_empty(),
        "DATE field should NOT produce a field_type variant (yet)");
    // The instrText run keeps the field-code string as its rendered text.
    let has_date_marker = runs.iter().any(|r| r.text.contains("DATE"));
    assert!(has_date_marker,
        "DATE field-code text should appear somewhere in the run stream");
}

#[test]
fn all_four_fixtures_parse_with_expected_field_counts() {
    // Smoke test: count field-bearing runs per fixture.
    let cases: &[(&str, usize, &[FieldType])] = &[
        ("v1_page.docx",              1, &[FieldType::Page]),
        ("v1_numpages.docx",          1, &[FieldType::NumPages]),
        ("v1_page_of_numpages.docx",  2, &[FieldType::Page, FieldType::NumPages]),
        ("v1_date.docx",              0, &[]),
    ];
    for (name, expected_count, expected_types) in cases {
        let path = fixture_path(name);
        if !path.exists() {
            eprintln!("skipping {}", name);
            continue;
        }
        let data = fs::read(&path).unwrap();
        let doc = parse_docx(&data)
            .unwrap_or_else(|e| panic!("failed to parse {}: {:?}", name, e));
        let runs = collect_runs(&doc);
        let actual_types: Vec<FieldType> = runs.iter()
            .filter_map(|r| r.field_type)
            .collect();
        assert_eq!(actual_types.len(), *expected_count, "{} field count", name);
        assert_eq!(actual_types, *expected_types, "{} field type sequence", name);
    }
}
