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
//! page's `Page.header` / `Page.footer` Vec<Block>.
//!
//! S755: `Page.header`/`Page.footer` always carry the DEFAULT-type part;
//! `<w:titlePg/>` and `<w:evenAndOddHeaders/>` variants land in
//! `Page.header_first` / `Page.header_even` (+ footer_*) and the LAYOUT
//! selects per rendered page number. See
//! `v1_title_page_uses_first_type_header_on_page_1`.
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
    // type=first (rIdHdr2). titlePg is set.
    //
    // ★S755 (2026-07-06) moved the per-page variant SELECTION out of the
    // parser and into the LAYOUT — do not "fix" this back to baking the
    // first-type header into `Page.header`. The parser now ALWAYS bakes the
    // DEFAULT type into `Page.header`/`Page.footer` and parses the variants
    // into `Page.header_first`/`header_even` (+ footer_*); the layout picks
    // per RENDERED page number (layout/mod.rs:6604 for the render,
    // layout/mod.rs:3078 for the body-top geometry):
    //   page 1 of a titlePg section  → header_first
    //   even physical page (evenAndOddHeaders) → header_even
    //   everything else              → header
    // WHY: the old parser bake applied the first-type header to EVERY page of
    // the section, so a TALL first-page header pushed the body down on all of
    // them (probextitlepg +1×6). `Page.header` is therefore the default-type
    // header even on a titlePg doc — that is not a bug.
    // OXI_S755_DISABLE restores the legacy bake; this test pins the DEFAULT.
    let Some(doc) = load("v1_title_page.docx") else { return };
    assert!(!doc.pages.is_empty());
    let page1 = &doc.pages[0];

    assert!(page1.title_pg, "sectPr <w:titlePg/> → Page.title_pg=true");

    // Page.header carries the DEFAULT-type header (the layout swaps it out
    // for header_first on page 1).
    assert_eq!(page1.header.len(), 1, "Page.header = the 1 default-type header");
    let hdr_text: String = page1.header.iter().map(block_text).collect();
    assert_eq!(hdr_text, "Default header (non-title pages)",
        "S755: Page.header is ALWAYS the default type, even under titlePg");

    // Page.header_first carries the type=first header the layout renders on
    // page 1. It is only populated when titlePg is set (parser/ooxml.rs:157).
    let hdr_first: String = page1.header_first.iter().map(block_text).collect();
    assert_eq!(hdr_first, "First-page-only header",
        "S755: titlePg + type=first → Page.header_first (layout picks it on page 1)");

    // The fixture declares only a DEFAULT footer reference, so footer_first
    // stays empty. NOTE: per ECMA-376 §17.10.2 (and the S755 layout), titlePg
    // with NO first-type reference renders a BLANK footer on page 1 — the
    // layout does NOT fall back to `Page.footer` there. `Page.footer` below is
    // the default-type footer that pages 2+ render.
    let ftr_text: String = page1.footer.iter().map(block_text).collect();
    assert_eq!(ftr_text, "Default footer", "Page.footer = the default-type footer");
    assert!(page1.footer_first.is_empty(),
        "no type=first footerReference → footer_first empty (page 1 footer is blank)");
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
