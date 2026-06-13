// This Source Code Form is subject to the terms of the Mozilla Public
// License, v. 2.0. If a copy of the MPL was not distributed with this
// file, You can obtain one at https://mozilla.org/MPL/2.0/.

//! Integration tests: parse `<w:headerReference>` / `<w:footerReference>`
//! end-to-end and verify `Page.header` / `Page.footer` populate from
//! word/headerN.xml / word/footerN.xml parts.
//!
//! Parser code path tested: [parser/ooxml.rs:5709](crates/oxidocs-core/src/parser/ooxml.rs#L5709)
//! collects header/footer relationships from `<w:sectPr>`, and the
//! layout-engine pass at parser/ooxml.rs:307 (via `parse_header_footer_xml`)
//! reads the referenced parts and attaches the parsed blocks to each
//! page's `Page.header` / `Page.footer` Vec<Block>. `<w:titlePg/>` causes
//! the first page to use the "first"-type reference instead of "default".
//!
//! Fixtures live in `tools/fixtures/header_footer_samples/` and are authored
//! by `tools/metrics/build_header_footer_repro_fixtures.py` (S286).

use std::fs;

use oxidocs_core::ir::{Block, Document};
use oxidocs_core::parser::parse_docx;

fn fixture_path(name: &str) -> std::path::PathBuf {
    let crate_root = std::env::current_dir().unwrap();
    let workspace_root = crate_root.parent().unwrap().parent().unwrap();
    workspace_root.join("tools").join("fixtures").join("header_footer_samples").join(name)
}

fn block_text(b: &Block) -> String {
    match b {
        Block::Paragraph(p) => p.runs.iter().map(|r| r.text.as_str()).collect(),
        _ => String::new(),
    }
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
fn v1_simple_doc_has_default_header_and_footer_on_each_page() {
    let Some(doc) = load("v1_simple.docx") else { return };
    assert!(!doc.pages.is_empty());
    for page in &doc.pages {
        assert_eq!(page.header.len(), 1, "every page has 1 header block");
        assert_eq!(page.footer.len(), 1, "every page has 1 footer block");
        let hdr_text: String = page.header.iter().map(block_text).collect();
        let ftr_text: String = page.footer.iter().map(block_text).collect();
        assert_eq!(hdr_text, "Header content - default");
        assert_eq!(ftr_text, "Footer content - default");
    }
}

#[test]
fn v1_header_only_doc_has_no_footer_blocks() {
    let Some(doc) = load("v1_header_only.docx") else { return };
    for page in &doc.pages {
        assert_eq!(page.header.len(), 1,
            "header is present");
        assert!(page.footer.is_empty(),
            "footer is empty when not referenced in sectPr");
    }
}

#[test]
fn v1_footer_only_doc_has_no_header_blocks() {
    let Some(doc) = load("v1_footer_only.docx") else { return };
    for page in &doc.pages {
        assert!(page.header.is_empty(),
            "header is empty when not referenced in sectPr");
        assert_eq!(page.footer.len(), 1,
            "footer is present");
    }
}

#[test]
fn v1_title_page_uses_first_type_header_on_page_1() {
    // sectPr has TWO header references: type=default (rIdHdr1) and
    // type=first (rIdHdr2). titlePg is set. On page 1 the layout must
    // pick the "first"-type header. (For a one-page doc, all pages are
    // "page 1" so the first-type wins everywhere.)
    let Some(doc) = load("v1_title_page.docx") else { return };
    assert!(!doc.pages.is_empty());
    let page1 = &doc.pages[0];
    assert_eq!(page1.header.len(), 1,
        "first page has exactly 1 header (the first-type one)");
    let hdr_text: String = page1.header.iter().map(block_text).collect();
    assert_eq!(hdr_text, "First-page-only header",
        "titlePg + type=first wins on page 1");
    // Footer is the default one (only one footer reference).
    let ftr_text: String = page1.footer.iter().map(block_text).collect();
    assert_eq!(ftr_text, "Default footer");
}

#[test]
fn all_four_fixtures_parse_with_expected_header_footer_presence() {
    let cases: &[(&str, bool, bool)] = &[
        // (filename, has_header_on_first_page, has_footer_on_first_page)
        ("v1_simple.docx",       true,  true),
        ("v1_header_only.docx",  true,  false),
        ("v1_footer_only.docx",  false, true),
        ("v1_title_page.docx",   true,  true),
    ];
    for (name, exp_hdr, exp_ftr) in cases {
        let path = fixture_path(name);
        if !path.exists() {
            eprintln!("skipping {}", name);
            continue;
        }
        let data = fs::read(&path).unwrap();
        let doc = parse_docx(&data)
            .unwrap_or_else(|e| panic!("failed to parse {}: {:?}", name, e));
        assert!(!doc.pages.is_empty(), "{} must have ≥1 page", name);
        let p1 = &doc.pages[0];
        assert_eq!(!p1.header.is_empty(), *exp_hdr,
            "{} header presence (page 1)", name);
        assert_eq!(!p1.footer.is_empty(), *exp_ftr,
            "{} footer presence (page 1)", name);
    }
}
