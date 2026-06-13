// This Source Code Form is subject to the terms of the Mozilla Public
// License, v. 2.0. If a copy of the MPL was not distributed with this
// file, You can obtain one at https://mozilla.org/MPL/2.0/.

//! Integration tests: parse `<w:docDefaults>` + `<w:style>` blocks
//! end-to-end and verify `Document.styles.{doc_default_run_style,
//! doc_default_para_style, doc_default_alignment, default_paragraph_style_id,
//! styles}` populate correctly. Also pins the `basedOn` inheritance
//! flattening done by `resolve_style_inheritance`.
//!
//! Parser code paths tested:
//! - [parser/styles.rs:36](crates/oxidocs-core/src/parser/styles.rs#L36)
//!   `parse_styles`: dispatches `<w:docDefaults>` / `<w:style>` blocks
//!   into `StyleSheet`.
//! - [parser/styles.rs:48-65](crates/oxidocs-core/src/parser/styles.rs#L48)
//!   `<w:rPrDefault>` / `<w:pPrDefault>` capture into
//!   `doc_default_run_style` / `doc_default_para_style` /
//!   `doc_default_alignment`.
//! - [parser/styles.rs:82-84](crates/oxidocs-core/src/parser/styles.rs#L82)
//!   `<w:style w:type="paragraph" w:default="1">` captures
//!   `default_paragraph_style_id`.
//! - [parser/styles.rs:132](crates/oxidocs-core/src/parser/styles.rs#L132)
//!   `resolve_style_inheritance`: merges parent `ParagraphStyle` /
//!   `RunStyle` into each child along the `<w:basedOn>` chain.
//! - [parser/styles.rs:1165](crates/oxidocs-core/src/parser/styles.rs#L1165)
//!   `<w:basedOn>` reference capture into `StyleDefinition.based_on`.
//! - [parser/styles.rs:1339](crates/oxidocs-core/src/parser/styles.rs#L1339)
//!   `<w:jc>` inside `<w:style>` populates
//!   `StyleDefinition.alignment` (separate from `ParagraphStyle`,
//!   which has no alignment field).
//!
//! Fixtures live in `tools/fixtures/styles_samples/` and are authored by
//! `tools/metrics/build_styles_repro_fixtures.py` (S298).

use std::fs;

use oxidocs_core::ir::Alignment;
use oxidocs_core::parser::parse_docx;

fn fixture_path(name: &str) -> std::path::PathBuf {
    let crate_root = std::env::current_dir().unwrap();
    let workspace_root = crate_root.parent().unwrap().parent().unwrap();
    workspace_root.join("tools").join("fixtures").join("styles_samples").join(name)
}

fn load(name: &str) -> Option<oxidocs_core::ir::Document> {
    let path = fixture_path(name);
    if !path.exists() {
        eprintln!("skipping: {} not found", path.display());
        return None;
    }
    let data = fs::read(&path).expect("read fixture");
    Some(parse_docx(&data).expect("parse fixture"))
}

#[test]
fn v1_doc_defaults_captures_run_font_size_and_para_alignment() {
    // rPrDefault: ascii="Calibri", sz=24 (=12pt), color=2E74B5
    // pPrDefault: jc=center, spacing line=276 (auto → 1.15)
    let Some(doc) = load("v1_doc_defaults.docx") else { return };

    let rs = doc.styles.doc_default_run_style.as_ref()
        .expect("rPrDefault must produce a doc_default_run_style");
    assert_eq!(rs.font_family.as_deref(), Some("Calibri"),
        "rPrDefault ascii→font_family");
    let sz = rs.font_size.expect("sz=24 must populate font_size");
    // sz val is in half-points, parser divides by 2 → 12pt.
    assert!((sz - 12.0).abs() < 0.01, "sz=24 → 12pt, got {}", sz);
    assert_eq!(rs.color.as_deref(), Some("2E74B5"),
        "rPrDefault color=2E74B5");

    let ps = doc.styles.doc_default_para_style.as_ref()
        .expect("pPrDefault must produce a doc_default_para_style");
    // line=276 with lineRule="auto" → 276/240 = 1.15
    let ls = ps.line_spacing.expect("spacing line must populate line_spacing");
    assert!((ls - 1.15).abs() < 0.01, "line=276/auto → 1.15, got {}", ls);

    // jc inside pPrDefault is stored separately on doc_default_alignment,
    // NOT on doc_default_para_style (ParagraphStyle has no alignment field).
    assert_eq!(doc.styles.doc_default_alignment, Some(Alignment::Center),
        "pPrDefault jc=center → Alignment::Center");
}

#[test]
fn v1_default_paragraph_style_id_captured_from_default_attr() {
    // Only one style, id="Normal", with w:default="1". No docDefaults.
    // Even without rPrDefault/pPrDefault, the parser should still record
    // the default paragraph style id, since that lookup drives
    // implicit-style inheritance for unstyled paragraphs.
    let Some(doc) = load("v1_default_para_style.docx") else { return };

    assert_eq!(doc.styles.default_paragraph_style_id.as_deref(), Some("Normal"),
        "type=paragraph + default=1 → default_paragraph_style_id=Normal");

    // The style itself is stored in the styles map.
    let def = doc.styles.styles.get("Normal")
        .expect("style id=Normal must be in styles map");
    assert_eq!(def.style_id, "Normal");
    assert!(def.based_on.is_none(),
        "Normal has no basedOn in this fixture");
    // No rPr → no default_run_style attached.
    assert!(def.paragraph.default_run_style.is_none(),
        "no rPr in style → default_run_style stays None");

    // No docDefaults at all → doc_default_* fields must stay None.
    assert!(doc.styles.doc_default_run_style.is_none(),
        "fixture has no rPrDefault");
    assert!(doc.styles.doc_default_para_style.is_none(),
        "fixture has no pPrDefault");
}

#[test]
fn v1_basedon_chain_merges_parent_font_size_into_child_bold() {
    // Normal:   <w:rPr><w:sz w:val="22"/></w:rPr>  (11pt only, no bold)
    // Heading1: basedOn="Normal", <w:rPr><w:b/></w:rPr>  (bold only, no sz)
    //
    // After resolve_style_inheritance the merged Heading1 must inherit
    // font_size from Normal while keeping its own bold. This is the
    // load-bearing property of the inheritance walk: a child with one
    // rPr field should not lose unrelated parent rPr fields.
    let Some(doc) = load("v1_basedon_chain.docx") else { return };

    let normal = doc.styles.styles.get("Normal")
        .expect("Normal style must exist");
    assert!(normal.based_on.is_none());
    let normal_rs = normal.paragraph.default_run_style.as_ref()
        .expect("Normal rPr.sz must populate default_run_style");
    let normal_sz = normal_rs.font_size.expect("Normal sz=22");
    assert!((normal_sz - 11.0).abs() < 0.01,
        "Normal sz=22 → 11pt, got {}", normal_sz);
    assert!(!normal_rs.bold, "Normal must not have bold set");

    let h1 = doc.styles.styles.get("Heading1")
        .expect("Heading1 style must exist");
    assert_eq!(h1.based_on.as_deref(), Some("Normal"),
        "Heading1 must record basedOn=Normal");

    let h1_rs = h1.paragraph.default_run_style.as_ref()
        .expect("Heading1 must end up with a default_run_style (from rPr.b)");
    assert!(h1_rs.bold,
        "Heading1's own <w:b/> must survive inheritance merge");
    let h1_sz = h1_rs.font_size.expect(
        "Heading1 font_size must be inherited from Normal sz=22",
    );
    assert!((h1_sz - 11.0).abs() < 0.01,
        "Heading1 inherits 11pt from Normal, got {}", h1_sz);
}

#[test]
fn v1_paragraph_style_captures_jc_and_left_indent() {
    // <w:style id="CenteredIndent" type="paragraph">
    //   <w:pPr><w:jc w:val="center"/><w:ind w:left="720"/></w:pPr>
    // </w:style>
    //
    // jc goes into StyleDefinition.alignment (not ParagraphStyle.*).
    // ind w:left=720 (twips÷20=36pt) goes into ParagraphStyle.indent_left.
    let Some(doc) = load("v1_para_style_alignment.docx") else { return };

    let def = doc.styles.styles.get("CenteredIndent")
        .expect("CenteredIndent style must exist");
    assert_eq!(def.alignment, Some(Alignment::Center),
        "pPr jc=center → StyleDefinition.alignment = Center");

    let left = def.paragraph.indent_left.expect(
        "ind w:left=720 must populate indent_left",
    );
    // 720 twentieths-of-pt / 20 = 36pt
    assert!((left - 36.0).abs() < 0.01,
        "ind left=720 → 36pt, got {}", left);

    // Independent indents must remain None to prove only `left` was set.
    assert!(def.paragraph.indent_right.is_none(),
        "ind has no w:right attr → indent_right stays None");
    assert!(def.paragraph.indent_first_line.is_none(),
        "ind has no w:firstLine attr → indent_first_line stays None");
}

#[test]
fn all_four_fixtures_parse_with_expected_style_shape() {
    // (filename, expected styles-map size, expected has_doc_default_rs, has_default_para_id)
    let cases: &[(&str, usize, bool, bool)] = &[
        ("v1_doc_defaults.docx",         0, true,  false),
        ("v1_default_para_style.docx",   1, false, true),
        ("v1_basedon_chain.docx",        2, false, true),
        ("v1_para_style_alignment.docx", 1, false, false),
    ];
    for (name, exp_n_styles, exp_has_rs, exp_has_default_id) in cases {
        let path = fixture_path(name);
        if !path.exists() {
            eprintln!("skipping {}", name);
            continue;
        }
        let data = fs::read(&path).unwrap();
        let doc = parse_docx(&data)
            .unwrap_or_else(|e| panic!("failed to parse {}: {:?}", name, e));
        assert_eq!(doc.styles.styles.len(), *exp_n_styles,
            "{} styles map size", name);
        assert_eq!(doc.styles.doc_default_run_style.is_some(), *exp_has_rs,
            "{} doc_default_run_style presence", name);
        assert_eq!(doc.styles.default_paragraph_style_id.is_some(), *exp_has_default_id,
            "{} default_paragraph_style_id presence", name);
    }
}
