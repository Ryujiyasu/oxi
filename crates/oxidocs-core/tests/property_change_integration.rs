// This Source Code Form is subject to the terms of the Mozilla Public
// License, v. 2.0. If a copy of the MPL was not distributed with this
// file, You can obtain one at https://mozilla.org/MPL/2.0/.

//! Integration tests: IR-level assertions on `<w:pPrChange>` and
//! `<w:rPrChange>` after `parse_docx`. Deepening pass over the existing
//! comments_samples fixtures (S306, modeled on S293 OMML deepening and
//! S297 tracked-changes deepening).
//!
//! `comments_fixtures.rs` covers layout-side balloon emission for these
//! fixtures; this file pins the raw IR fields surfaced under
//! `Paragraph.ppr_change` / `Run.rpr_change` so a future refactor that
//! drops or renames a `PropertyChange` field is caught at the parser
//! seam rather than via downstream layout assertions.
//!
//! Parser code path tested:
//!   - [parser/ooxml.rs:1926](crates/oxidocs-core/src/parser/ooxml.rs#L1926)
//!     pPrChange handler — reads w:id/w:author/w:date attrs, drains the
//!     inner `<w:pPr>` via `parse_paragraph_properties`, lands result
//!     under `Paragraph.ppr_change`.
//!   - The rPrChange path inside `parse_run_properties` mirrors the
//!     pPrChange shape and lands under `Run.rpr_change`.
//!
//! Non-obvious behaviors pinned:
//!   - `<w:pPr/>` Empty element (no children) still materializes
//!     `prior_paragraph_style = Some(ParagraphStyle::default())` —
//!     i.e., presence of an empty pPr means "prior state was the empty
//!     paragraph style", not "no prior state captured".
//!   - `prior_alignment` is only populated when the inner pPr declared
//!     an explicit `<w:jc>` child. Missing jc → None, NOT Left default
//!     — alignment lives outside ParagraphStyle so this field exists
//!     specifically to track jc as a tracked-change field.
//!   - id/author/date attributes preserved verbatim, including the
//!     ISO-8601 date string.

use std::fs;

use oxidocs_core::ir::{Alignment, Block, Document, Paragraph, PropertyChange, Run};
use oxidocs_core::parse_docx;

fn fixture_path(name: &str) -> std::path::PathBuf {
    let crate_root = std::env::current_dir().unwrap();
    let workspace_root = crate_root.parent().unwrap().parent().unwrap();
    workspace_root
        .join("tools")
        .join("fixtures")
        .join("comments_samples")
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

fn paragraphs(doc: &Document) -> Vec<&Paragraph> {
    doc.pages
        .iter()
        .flat_map(|p| p.blocks.iter())
        .filter_map(|b| if let Block::Paragraph(p) = b { Some(p) } else { None })
        .collect()
}

fn collect_runs(doc: &Document) -> Vec<&Run> {
    paragraphs(doc).iter().flat_map(|p| p.runs.iter()).collect()
}

fn find_ppr_change(doc: &Document) -> Option<(&Paragraph, &PropertyChange)> {
    paragraphs(doc)
        .into_iter()
        .find_map(|p| p.ppr_change.as_ref().map(|pc| (p, pc)))
}

fn find_rpr_change(doc: &Document) -> Option<(&Run, &PropertyChange)> {
    collect_runs(doc)
        .into_iter()
        .find_map(|r| r.rpr_change.as_ref().map(|pc| (r, pc)))
}

#[test]
fn fixture_13_ppr_change_indent_pins_id_author_date() {
    let Some(doc) = load("fixture_13_pPrChange_indent.docx") else { return };
    let (_p, pc) = find_ppr_change(&doc).expect("fixture 13 has pPrChange");
    assert_eq!(pc.id.as_deref(), Some("600"), "id verbatim");
    assert_eq!(pc.author.as_deref(), Some("Alice Reviewer"));
    assert_eq!(
        pc.date.as_deref(),
        Some("2026-04-18T10:00:00Z"),
        "ISO-8601 date verbatim"
    );
    // The inner `<w:pPr/>` is empty Empty (self-closing), so
    // prior_paragraph_style materializes as Some(default).
    assert!(
        pc.prior_paragraph_style.is_some(),
        "empty <w:pPr/> still triggers Some(default) prior_paragraph_style"
    );
    // No inner <w:jc> → prior_alignment must stay None (not Left fallback).
    assert!(
        pc.prior_alignment.is_none(),
        "missing <w:jc> in prior pPr → prior_alignment None, NOT Left default"
    );
}

#[test]
fn fixture_14_rpr_change_font_pins_id_author_date() {
    let Some(doc) = load("fixture_14_rPrChange_font.docx") else { return };
    let (_r, pc) = find_rpr_change(&doc).expect("fixture 14 has rPrChange");
    assert_eq!(pc.id.as_deref(), Some("700"));
    assert_eq!(pc.author.as_deref(), Some("Alice Reviewer"));
    assert_eq!(pc.date.as_deref(), Some("2026-04-18T10:00:00Z"));
    assert!(
        pc.prior_run_style.is_some(),
        "empty <w:rPr/> in rPrChange materializes Some(default) prior_run_style"
    );
}

#[test]
fn fixture_15_ppr_change_alignment_captures_prior_jc_left() {
    // The ONLY fixture whose inner pPr declares a child element
    // (<w:jc w:val="left"/>). Pins:
    //   - prior_alignment = Some(Left) — the jc value is captured
    //     in a SEPARATE field from prior_paragraph_style (since
    //     Paragraph.alignment lives outside ParagraphStyle).
    //   - prior_paragraph_style is still Some(...) but its alignment
    //     field is unrelated — ParagraphStyle has no alignment.
    let Some(doc) = load("fixture_15_pPrChange_alignment.docx") else { return };
    let (_p, pc) = find_ppr_change(&doc).expect("fixture 15 has pPrChange");
    assert_eq!(pc.id.as_deref(), Some("800"));
    assert_eq!(pc.author.as_deref(), Some("Alice Reviewer"));
    assert_eq!(pc.date.as_deref(), Some("2026-04-18T10:00:00Z"));
    assert!(pc.prior_paragraph_style.is_some());
    assert_eq!(
        pc.prior_alignment,
        Some(Alignment::Left),
        "inner <w:jc w:val=\"left\"/> materializes prior_alignment = Left"
    );
}

#[test]
fn fixture_18_ppr_change_shading_pins_prior_paragraph_style_some() {
    // fixture_18 toggles paragraph shading via pPrChange. The prior
    // pPr is empty (no children), so prior_paragraph_style is Some(default)
    // — same shape as fixture_13 but distinct fixture confirms the parser
    // doesn't conflate the two fixture indices.
    let Some(doc) = load("fixture_18_pPrChange_shading.docx") else { return };
    let (_p, pc) = find_ppr_change(&doc).expect("fixture 18 has pPrChange");
    assert_eq!(pc.id.as_deref(), Some("1100"));
    assert_eq!(pc.author.as_deref(), Some("Alice Reviewer"));
    assert!(pc.prior_paragraph_style.is_some());
    assert!(
        pc.prior_alignment.is_none(),
        "fixture_18's inner pPr has no jc"
    );
}

#[test]
fn rpr_change_id_remains_distinct_across_fixtures() {
    // Pin the per-fixture id values so a future serializer refactor
    // that accidentally normalizes ids (e.g., re-numbering on save)
    // is caught immediately. Walks the rPrChange-bearing fixtures
    // and asserts each holds its expected literal id.
    let cases: &[(&str, &str)] = &[
        ("fixture_09_rPrChange_bold.docx", "300"),
        ("fixture_14_rPrChange_font.docx", "700"),
        ("fixture_16_rPrChange_caps_spacing.docx", "900"),
        ("fixture_17_rPrChange_vAlign_shading.docx", "1000"),
    ];
    for (name, expected_id) in cases {
        let Some(doc) = load(name) else { continue };
        let (_r, pc) = find_rpr_change(&doc)
            .unwrap_or_else(|| panic!("{} has rPrChange", name));
        assert_eq!(
            pc.id.as_deref(),
            Some(*expected_id),
            "{} rPrChange id should be {expected_id}",
            name
        );
        assert!(
            pc.prior_run_style.is_some(),
            "{} prior_run_style must be Some after parse",
            name
        );
        // None of these fixtures have an inner pPr inside rPrChange,
        // so prior_paragraph_style stays None on rpr_change.
        assert!(
            pc.prior_paragraph_style.is_none(),
            "{} rPrChange must NOT populate prior_paragraph_style",
            name
        );
    }
}

#[test]
fn ppr_change_and_rpr_change_do_not_cross_wire() {
    // fixture_13 has a pPrChange but NO rPrChange on any run.
    // fixture_14 has an rPrChange but NO pPrChange on any paragraph.
    // Crossed wiring (e.g., parser writing rPrChange into ppr_change)
    // is a subtle regression that wouldn't fail single-fixture tests.
    let Some(doc13) = load("fixture_13_pPrChange_indent.docx") else { return };
    assert!(find_ppr_change(&doc13).is_some(), "fixture 13 has pPrChange");
    assert!(
        collect_runs(&doc13).iter().all(|r| r.rpr_change.is_none()),
        "fixture 13: no run should carry rpr_change"
    );

    let Some(doc14) = load("fixture_14_rPrChange_font.docx") else { return };
    assert!(find_rpr_change(&doc14).is_some(), "fixture 14 has rPrChange");
    assert!(
        paragraphs(&doc14).iter().all(|p| p.ppr_change.is_none()),
        "fixture 14: no paragraph should carry ppr_change"
    );
}
