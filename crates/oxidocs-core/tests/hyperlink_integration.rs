//! Integration tests: parse w:hyperlink elements end-to-end and verify
//! `Run.url` is populated correctly for both external URL relationships
//! (r:id → DOC_RELS) and internal bookmark anchors (w:anchor → "#name").
//!
//! Fixtures live in `tools/fixtures/hyperlink_samples/` and are authored by
//! `tools/metrics/build_hyperlink_repro_fixtures.py` (S272). Parser code
//! path tested: `parser/ooxml.rs:1080` (the `"hyperlink"` arm), which calls
//! `parse_hyperlink_runs` and propagates `link_url` to every contained run.
//!
//! Companion to `vertical_integration.rs` (tbRlV cells),
//! `ruby_integration.rs` (furigana), `omml_integration.rs` (math), and
//! `comments_fixtures.rs` (tracked-changes balloons).

use std::fs;

use oxidocs_core::ir::{Block, Document, Run};
use oxidocs_core::parser::parse_docx;

fn fixture_path(name: &str) -> std::path::PathBuf {
    // Tests run with CWD at the crate root; fixtures are two levels up.
    let crate_root = std::env::current_dir().unwrap();
    let workspace_root = crate_root.parent().unwrap().parent().unwrap();
    workspace_root.join("tools").join("fixtures").join("hyperlink_samples").join(name)
}

/// Walk every run in document order. Returned in document order so the
/// Nth run in the file can be asserted by index.
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
fn v1_external_url_resolves_via_doc_rels() {
    // <w:hyperlink r:id="rId3"> with rId3 → "https://www.example.com/" in
    // word/_rels/document.xml.rels. Verifies the rels-lookup branch of the
    // parser ("r:id" attribute → ctx.hyperlinks.get).
    let Some(doc) = load("v1_external.docx") else { return };
    let runs = collect_runs(&doc);
    assert_eq!(runs.len(), 3, "expected: plain + hyperlink + plain");
    assert_eq!(runs[0].text, "Click here: ");
    assert!(runs[0].url.is_none());
    assert_eq!(runs[1].text, "Example");
    assert_eq!(runs[1].url.as_deref(), Some("https://www.example.com/"));
    assert_eq!(runs[2].text, ".");
    assert!(runs[2].url.is_none());
}

#[test]
fn v1_internal_anchor_is_prefixed_with_hash() {
    // <w:hyperlink w:anchor="section1"> with no rels entry. Verifies the
    // anchor branch (parser prefixes the bookmark name with "#").
    let Some(doc) = load("v1_anchor.docx") else { return };
    let runs = collect_runs(&doc);
    assert_eq!(runs.len(), 3);
    assert_eq!(runs[1].text, "Section 1");
    assert_eq!(runs[1].url.as_deref(), Some("#section1"),
        "anchor link should carry '#'-prefixed bookmark name");
    assert!(runs[0].url.is_none() && runs[2].url.is_none());
}

#[test]
fn v1_mixed_paragraph_carries_both_url_kinds() {
    // One paragraph: plain + external + plain + anchor + plain.
    // Exercises both code paths back-to-back inside a single paragraph;
    // also pins ordering after the parser flattens hyperlink contents.
    let Some(doc) = load("v1_mixed.docx") else { return };
    let runs = collect_runs(&doc);
    assert_eq!(runs.len(), 5);

    let url_at: Vec<Option<&str>> = runs.iter().map(|r| r.url.as_deref()).collect();
    let text_at: Vec<&str> = runs.iter().map(|r| r.text.as_str()).collect();
    assert_eq!(text_at, vec!["Start ", "Ext", " middle ", "Intro", " end."]);
    assert_eq!(
        url_at,
        vec![None, Some("https://docs.rs/"), None, Some("#intro"), None]
    );
}

#[test]
fn v1_multirun_link_propagates_url_to_every_contained_run() {
    // Single <w:hyperlink r:id="rId3"> wrapping two <w:r>s (one bold, one
    // plain). Both runs must carry the SAME url; bold-ness must be preserved
    // independently — verifies link_url is propagated per-run and rPr
    // attributes are NOT overwritten by hyperlink wrapping.
    let Some(doc) = load("v1_multirun.docx") else { return };
    let runs = collect_runs(&doc);
    assert_eq!(runs.len(), 4);
    assert_eq!(runs[0].text, "Mixed-style link: ");
    assert!(runs[0].url.is_none());

    // Inside the hyperlink: two runs, both carrying https://crates.io/.
    assert_eq!(runs[1].text, "Bold");
    assert_eq!(runs[1].url.as_deref(), Some("https://crates.io/"));
    assert!(runs[1].style.bold, "first link-run keeps its w:b rPr");

    assert_eq!(runs[2].text, " Plain");
    assert_eq!(runs[2].url.as_deref(), Some("https://crates.io/"));
    assert!(!runs[2].style.bold, "second link-run is not bold");

    assert_eq!(runs[3].text, ".");
    assert!(runs[3].url.is_none());
}

#[test]
fn all_four_fixtures_parse_with_at_least_one_url_run() {
    // Smoke test: every fixture parses and yields at least one Run.url=Some(_).
    for name in ["v1_external.docx", "v1_anchor.docx", "v1_mixed.docx", "v1_multirun.docx"] {
        let path = fixture_path(name);
        if !path.exists() {
            eprintln!("skipping {}", name);
            continue;
        }
        let data = fs::read(&path).unwrap();
        let doc = parse_docx(&data)
            .unwrap_or_else(|e| panic!("failed to parse {}: {:?}", name, e));
        let urls: Vec<_> = collect_runs(&doc).into_iter()
            .filter_map(|r| r.url.as_deref())
            .collect();
        assert!(!urls.is_empty(), "{} should produce at least 1 url-bearing run", name);
    }
}
