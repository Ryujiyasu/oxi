// This Source Code Form is subject to the terms of the Mozilla Public
// License, v. 2.0. If a copy of the MPL was not distributed with this
// file, You can obtain one at https://mozilla.org/MPL/2.0/.

//! Integration tests: parse `<w:hdr>` / `<w:ftr>` INNER content
//! end-to-end and verify `Page.header` / `Page.footer` block lists
//! after `parse_docx`. Deepening pass over header_footer_integration
//! (S286) which covered routing (default/first ref, title_pg) but
//! NOT the inner content branches.
//!
//! Parser code path tested:
//!   - [parser/ooxml.rs:5835](crates/oxidocs-core/src/parser/ooxml.rs#L5835)
//!     `parse_header_footer_xml` end-to-end on its four content
//!     branches: <w:p> (paragraph), <w:tbl> (table),
//!     <w:sdt><w:sdtContent><w:p> (sdt-wrapped paragraph),
//!     <w:sdt><w:sdtContent><w:tbl> (sdt-wrapped table).
//!
//! Non-obvious behaviors pinned:
//!   - Multiple <w:p> children → multiple Block::Paragraph entries
//!     in Page.header, source-order preserved.
//!   - <w:tbl> inside <w:hdr> → Block::Table at the SAME LEVEL as
//!     Block::Paragraph (line 5854-5857). No nesting under a
//!     synthetic wrapper.
//!   - <w:sdt> WRAPPER is NOT a Block — its <w:sdtContent> children
//!     (paragraphs OR tables) are harvested DIRECTLY into
//!     Page.header (parser/ooxml.rs:5858-5891). A regression that
//!     stored sdt as its own block, or that dropped sdt content
//!     entirely, would silently affect Word docs using content
//!     controls in headers (common in templates).
//!   - <w:sdtPr> (sibling of <w:sdtContent> inside sdt) is SKIPPED
//!     by the parser. The id/lock/placeholder metadata is discarded
//!     intentionally; only sdtContent matters for layout.
//!   - parse_header_footer_xml DELEGATES to parse_paragraph for <w:p>
//!     (line 5851), so paragraph properties (pPr/jc) and run
//!     properties (rPr/b) propagate end-to-end through the
//!     header/footer path — same as body paragraphs. Pinning this
//!     catches an accidental override where the header parser
//!     forgot to plumb ctx/styles.

use std::fs;

use oxidocs_core::ir::{Alignment, Block, Document};
use oxidocs_core::parse_docx;

fn fixture_path(name: &str) -> std::path::PathBuf {
    let crate_root = std::env::current_dir().unwrap();
    let workspace_root = crate_root.parent().unwrap().parent().unwrap();
    workspace_root
        .join("tools")
        .join("fixtures")
        .join("header_footer_content_samples")
        .join(name)
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

fn block_text(b: &Block) -> String {
    match b {
        Block::Paragraph(p) => p.runs.iter().map(|r| r.text.as_str()).collect(),
        Block::Table(t) => t
            .rows
            .iter()
            .flat_map(|row| row.cells.iter())
            .flat_map(|cell| cell.blocks.iter())
            .map(block_text)
            .collect::<Vec<_>>()
            .join("|"),
        _ => String::new(),
    }
}

#[test]
fn v1_header_two_paragraphs_preserve_source_order() {
    let Some(doc) = load("v1_header_two_paragraphs_order.docx") else { return };
    let page = &doc.pages[0];

    assert_eq!(
        page.header.len(),
        2,
        "two <w:p> children → 2 Block entries in Page.header"
    );
    assert_eq!(block_text(&page.header[0]), "first");
    assert_eq!(block_text(&page.header[1]), "second");

    // Both must be Block::Paragraph variants — NOT Block::Table or
    // anything else. A regression that wrapped them in a synthetic
    // container would surface here.
    assert!(matches!(page.header[0], Block::Paragraph(_)));
    assert!(matches!(page.header[1], Block::Paragraph(_)));
}

#[test]
fn v1_header_table_routes_to_block_table() {
    let Some(doc) = load("v1_header_table.docx") else { return };
    let page = &doc.pages[0];

    assert_eq!(page.header.len(), 1, "single <w:tbl> → 1 Block entry");
    let block = &page.header[0];

    // The block MUST be Block::Table — same routing as body-level
    // tables. parser/ooxml.rs:5854-5857.
    let table = match block {
        Block::Table(t) => t,
        _ => panic!(
            "expected Block::Table inside Page.header, got: {:?}",
            std::mem::discriminant(block),
        ),
    };
    assert_eq!(table.rows.len(), 1);
    assert_eq!(table.rows[0].cells.len(), 2);

    let cell_a_text: String = table.rows[0].cells[0]
        .blocks
        .iter()
        .map(block_text)
        .collect();
    let cell_b_text: String = table.rows[0].cells[1]
        .blocks
        .iter()
        .map(block_text)
        .collect();
    assert_eq!(cell_a_text, "hdr-cell-A");
    assert_eq!(cell_b_text, "hdr-cell-B");
}

#[test]
fn v1_header_sdt_paragraph_harvests_inner_directly() {
    let Some(doc) = load("v1_header_sdt_paragraph.docx") else { return };
    let page = &doc.pages[0];

    // The sdt WRAPPER is NOT a Block. Its inner paragraph is
    // harvested DIRECTLY into Page.header. So there's exactly ONE
    // entry, and it is Block::Paragraph (NOT some synthetic sdt
    // variant).
    assert_eq!(
        page.header.len(),
        1,
        "<w:sdt><w:sdtContent><w:p>...</w:p></w:sdtContent></w:sdt> \
         → 1 Block (the inner paragraph, NOT the sdt wrapper)"
    );
    assert!(
        matches!(page.header[0], Block::Paragraph(_)),
        "inner is Block::Paragraph (sdt is transparent)"
    );
    assert_eq!(block_text(&page.header[0]), "inside-sdt");
}

#[test]
fn v1_header_sdt_table_harvests_table_directly() {
    let Some(doc) = load("v1_header_sdt_table.docx") else { return };
    let page = &doc.pages[0];

    // The OTHER sdtContent branch (line 5871-5874): table inside
    // sdt → Block::Table directly, same as table NOT wrapped.
    assert_eq!(
        page.header.len(),
        1,
        "<w:sdt><w:sdtContent><w:tbl>...</w:tbl></w:sdtContent></w:sdt> \
         → 1 Block (the inner table)"
    );
    let table = match &page.header[0] {
        Block::Table(t) => t,
        _ => panic!("expected Block::Table inside sdt → Page.header"),
    };
    assert_eq!(table.rows.len(), 1);
    assert_eq!(table.rows[0].cells.len(), 2);
}

#[test]
fn v1_footer_paragraph_properties_propagate_through_parse_header_footer_xml() {
    let Some(doc) = load("v1_footer_para_with_properties.docx") else { return };
    let page = &doc.pages[0];

    assert_eq!(page.footer.len(), 1);
    let para = match &page.footer[0] {
        Block::Paragraph(p) => p,
        _ => panic!("expected Block::Paragraph in footer"),
    };

    // pPr/jc=center → Paragraph.alignment = Center. The header/footer
    // parser delegates to parse_paragraph, so paragraph properties
    // propagate end-to-end. A regression that forgot to plumb
    // ctx/styles into parse_header_footer_xml's parse_paragraph call
    // would surface as Left (default) here.
    assert_eq!(
        para.alignment,
        Alignment::Center,
        "footer paragraph jc=center propagates through parse_header_footer_xml"
    );

    // rPr/b on the run → run.style.bold = true. Same propagation
    // expectation for run-level properties.
    let run = para.runs.first().expect("first run");
    assert!(
        run.style.bold,
        "footer paragraph run rPr/b propagates → run.style.bold=true"
    );
    assert_eq!(run.text, "centered-bold");
}

#[test]
fn all_five_fixtures_produce_expected_header_or_footer_block_count() {
    // Smoke: every fixture must produce exactly 1 page with a
    // non-empty header or footer matching its content.
    let cases: &[(&str, usize, usize)] = &[
        // (filename, expected_header_count, expected_footer_count)
        ("v1_header_two_paragraphs_order.docx", 2, 0),
        ("v1_header_table.docx", 1, 0),
        ("v1_header_sdt_paragraph.docx", 1, 0),
        ("v1_header_sdt_table.docx", 1, 0),
        ("v1_footer_para_with_properties.docx", 0, 1),
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
        let page = &doc.pages[0];
        assert_eq!(
            page.header.len(),
            *exp_hdr,
            "{} header block count",
            name
        );
        assert_eq!(
            page.footer.len(),
            *exp_ftr,
            "{} footer block count",
            name
        );
    }
}
