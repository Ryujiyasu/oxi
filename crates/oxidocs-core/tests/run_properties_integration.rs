// This Source Code Form is subject to the terms of the Mozilla Public
// License, v. 2.0. If a copy of the MPL was not distributed with this
// file, You can obtain one at https://mozilla.org/MPL/2.0/.

//! Integration tests: parse inline `<w:r><w:rPr>...</w:rPr>` end-to-end
//! and verify the per-field `Run.style: RunStyle` shape after `parse_docx`.
//!
//! `styles_integration.rs` (S298) covers `<w:rPrDefault>` and style-sheet
//! `<w:rPr>` merge. `comments_fixtures.rs::fixture_09` exercises
//! `Run.style.bold` via an rPrChange toggle. But no integration test
//! pins the breadth of fields that `parse_run_properties` materializes
//! when an rPr is declared inline on a body run (parser/ooxml.rs:4189).
//!
//! Non-obvious behaviors pinned:
//!   - [parser/ooxml.rs:4291](crates/oxidocs-core/src/parser/ooxml.rs#L4291)
//!     `<w:u w:val="none"/>` SUPPRESSES underline (sets `underline=false`)
//!     even though the element is present — the val attribute polarity-
//!     flips the boolean. `<w:u w:val="single"/>` populates both
//!     `underline=true` AND `underline_style=Some("single")`.
//!   - [parser/ooxml.rs:4300-4303](crates/oxidocs-core/src/parser/ooxml.rs#L4300)
//!     `<w:dstrike/>` sets BOTH `strikethrough=true` AND
//!     `double_strikethrough=true` (not just the dstrike flag — single-
//!     strikethrough renderers still draw a strike on dstrike runs).
//!   - [parser/ooxml.rs:4329](crates/oxidocs-core/src/parser/ooxml.rs#L4329)
//!     `<w:sz w:val="22"/>` → 11.0pt (half-points / 2). Half values
//!     are preserved: val="23" → 11.5pt (not rounded to 12pt).
//!   - [parser/ooxml.rs:4396-4404](crates/oxidocs-core/src/parser/ooxml.rs#L4396)
//!     `<w:color w:val="auto"/>` → `color = None` (the "auto" sentinel
//!     suppresses storage, NOT stored as the string "auto"). Explicit
//!     hex stored verbatim.
//!   - [parser/ooxml.rs:4413](crates/oxidocs-core/src/parser/ooxml.rs#L4413)
//!     `<w:spacing w:val="40"/>` → 2.0pt (twips / 20). Note: distinct
//!     from `<w:sz>` half-point divisor — character_spacing is twips.
//!   - [parser/ooxml.rs:4478](crates/oxidocs-core/src/parser/ooxml.rs#L4478)
//!     `<w:kern w:val="22"/>` → 11.0pt (half-points / 2, same divisor
//!     as sz, NOT twips like spacing).
//!   - [parser/ooxml.rs:4522](crates/oxidocs-core/src/parser/ooxml.rs#L4522)
//!     `<w:position w:val="6"/>` → +3.0pt (half-points / 2, signed —
//!     val="-6" → -3.0pt for lowered text).
//!   - [parser/ooxml.rs:4421](crates/oxidocs-core/src/parser/ooxml.rs#L4421)
//!     `<w:w w:val="80"/>` → `text_scale = Some(80.0)` (raw percentage,
//!     NO conversion — w:w is "percentage" semantics in OOXML).
//!   - [parser/ooxml.rs:4254-4260](crates/oxidocs-core/src/parser/ooxml.rs#L4254)
//!     `<w:rFonts w:ascii="Arial" w:eastAsia="ＭＳ Ｐ明朝"/>` populates
//!     `font_family` AND `font_family_east_asia` as SEPARATE fields,
//!     and sets `has_explicit_east_asia=true` (the latter is the
//!     gate for §4.6.3 Latin-space-adjacent-CJK widening per S136).
//!   - [parser/ooxml.rs:4445](crates/oxidocs-core/src/parser/ooxml.rs#L4445)
//!     `<w:webHidden/>` is NOT an alias for `<w:vanish/>`: per ISO/IEC
//!     29500-1 §17.3.2.44 it hides text only in WEB PAGE VIEW and "should
//!     not affect a normal paginated view", so Oxi (a paginated renderer)
//!     treats it as a NO-OP. Only `<w:vanish/>` (§17.3.2.45) sets `vanish`.
//!   - [parser/ooxml.rs:4425-4430](crates/oxidocs-core/src/parser/ooxml.rs#L4425)
//!     `<w:smallCaps/>` and `<w:caps/>` are INDEPENDENT flags — both
//!     can be true simultaneously (a regression here would silently
//!     swap one for the other).
//!   - [parser/ooxml.rs:4431-4441](crates/oxidocs-core/src/parser/ooxml.rs#L4431)
//!     `<w:shd w:fill="FFFF00"/>` populates run-level `shading`
//!     (distinct from paragraph-level shading; pinned here at the
//!     parser seam, not via layout).
//!
//! Fixtures live in `tools/fixtures/run_properties_samples/` and are
//! authored by `tools/metrics/build_run_properties_repro_fixtures.py`
//! (S308). The fixture's styles.xml is intentionally empty (no
//! docDefaults) so the inline rPr is the SOLE source of run
//! formatting — any field populated on `Run.style` must have come
//! from the inline rPr.

use std::fs;

use oxidocs_core::ir::{Block, Document, Paragraph, Run, VerticalAlign};
use oxidocs_core::parse_docx;

fn fixture_path(name: &str) -> std::path::PathBuf {
    let crate_root = std::env::current_dir().unwrap();
    let workspace_root = crate_root.parent().unwrap().parent().unwrap();
    workspace_root
        .join("tools")
        .join("fixtures")
        .join("run_properties_samples")
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

fn find_run<'a>(doc: &'a Document, text: &str) -> &'a Run {
    collect_runs(doc)
        .into_iter()
        .find(|r| r.text == text)
        .unwrap_or_else(|| panic!("run with text {:?} not found", text))
}

#[test]
fn v1_bold_italic_underline_with_none_polarity() {
    let Some(doc) = load("v1_bold_italic_underline.docx") else { return };

    let plain = find_run(&doc, "plain");
    assert!(!plain.style.bold, "plain run is not bold");
    assert!(!plain.style.italic, "plain run is not italic");
    assert!(!plain.style.underline, "plain run is not underlined");
    assert!(plain.style.underline_style.is_none(), "plain has no u style");

    let bi = find_run(&doc, "bold-italic");
    assert!(bi.style.bold, "<w:b/> → bold=true");
    assert!(bi.style.italic, "<w:i/> → italic=true");
    assert!(!bi.style.underline, "no <w:u> → underline=false");

    let us = find_run(&doc, "underline-single");
    assert!(us.style.underline, "<w:u w:val=\"single\"/> → underline=true");
    assert_eq!(
        us.style.underline_style.as_deref(),
        Some("single"),
        "underline_style preserves the val attr verbatim"
    );

    // The non-obvious branch: `val="none"` flips underline back to
    // false even though the <w:u> element IS present. Parser hits
    // the Empty-element handler at parser/ooxml.rs:4286, defaults
    // underline=true, then the val=none check (line 4291) sets it
    // back to false. underline_style stays None in this case.
    let un = find_run(&doc, "underline-none");
    assert!(
        !un.style.underline,
        "<w:u w:val=\"none\"/> SUPPRESSES underline despite element presence"
    );
    assert!(
        un.style.underline_style.is_none(),
        "val=none does NOT populate underline_style"
    );
}

#[test]
fn v1_dstrike_sets_both_strike_and_double_strike_flags() {
    let Some(doc) = load("v1_strike_dstrike_vertalign.docx") else { return };

    let plain = find_run(&doc, "plain");
    assert!(!plain.style.strikethrough);
    assert!(!plain.style.double_strikethrough);
    assert!(plain.style.vertical_align.is_none());

    let strike = find_run(&doc, "strike");
    assert!(strike.style.strikethrough, "<w:strike/> → strikethrough=true");
    assert!(
        !strike.style.double_strikethrough,
        "<w:strike/> does NOT set double_strikethrough"
    );

    // The non-obvious branch: dstrike sets BOTH flags so a renderer
    // that only checks `strikethrough` still draws *a* strike on
    // double-strike runs.
    let dstrike = find_run(&doc, "dstrike");
    assert!(
        dstrike.style.strikethrough,
        "<w:dstrike/> sets strikethrough=true (single-strike fallback)"
    );
    assert!(
        dstrike.style.double_strikethrough,
        "<w:dstrike/> ALSO sets double_strikethrough=true"
    );

    let sup = find_run(&doc, "super");
    assert_eq!(
        sup.style.vertical_align,
        Some(VerticalAlign::Superscript),
        "vertAlign val=superscript → Superscript enum"
    );

    let sub = find_run(&doc, "sub");
    assert_eq!(
        sub.style.vertical_align,
        Some(VerticalAlign::Subscript),
        "vertAlign val=subscript → Subscript enum"
    );
}

#[test]
fn v1_sz_half_points_color_auto_suppressed_highlight_verbatim() {
    let Some(doc) = load("v1_sz_color_highlight.docx") else { return };

    // Half-point arithmetic: val=22 → 11.0pt exactly.
    let sz22 = find_run(&doc, "sz22");
    let fs22 = sz22.style.font_size.expect("sz=22 must populate font_size");
    assert!(
        (fs22 - 11.0).abs() < 0.001,
        "sz val=22 → 11.0pt, got {}",
        fs22
    );

    // Half values are preserved (NOT rounded): val=23 → 11.5pt.
    let sz23 = find_run(&doc, "sz23-half");
    let fs23 = sz23.style.font_size.expect("sz=23 must populate font_size");
    assert!(
        (fs23 - 11.5).abs() < 0.001,
        "sz val=23 → 11.5pt (half preserved, not rounded), got {}",
        fs23
    );

    // color val="auto" → None (the "auto" sentinel suppresses storage,
    // NOT stored verbatim as the string "auto"). A regression that
    // stores "auto" would break downstream color resolution which
    // expects either None or a hex string.
    let cauto = find_run(&doc, "color-auto");
    assert!(
        cauto.style.color.is_none(),
        "color val=auto → None (not stored as the string \"auto\")"
    );

    // Explicit hex preserved verbatim (no case-folding, no shortening).
    let cred = find_run(&doc, "color-red");
    assert_eq!(
        cred.style.color.as_deref(),
        Some("FF0000"),
        "hex color stored verbatim"
    );

    // highlight stored as the enum-string verbatim (e.g. "yellow",
    // "green") — NOT translated to a hex value at parse time.
    let hl = find_run(&doc, "highlight-yellow");
    assert_eq!(
        hl.style.highlight.as_deref(),
        Some("yellow"),
        "highlight val stored as the enum-string verbatim"
    );
}

#[test]
fn v1_kern_position_spacing_unit_distinctions() {
    // Pins the three DIFFERENT unit conversions on the run:
    //   - kern: half-points / 2
    //   - position: half-points / 2 (signed)
    //   - spacing: twips / 20
    //   - w (text_scale): raw percentage, no conversion
    let Some(doc) = load("v1_kern_position_spacing.docx") else { return };

    let kern = find_run(&doc, "kern22");
    let kv = kern.style.kern.expect("<w:kern val=22/> must populate kern");
    assert!(
        (kv - 11.0).abs() < 0.001,
        "kern val=22 → 11.0pt (half-points / 2), got {}",
        kv
    );

    let raised = find_run(&doc, "pos-raised");
    let pv = raised.style.position.expect("position must populate");
    assert!(
        (pv - 3.0).abs() < 0.001,
        "position val=6 → +3.0pt (half-points / 2), got {}",
        pv
    );

    // Negative position survives the half-point divide (signed division).
    let lowered = find_run(&doc, "pos-lowered");
    let lv = lowered.style.position.expect("negative position must populate");
    assert!(
        (lv - (-3.0)).abs() < 0.001,
        "position val=-6 → -3.0pt (signed half-points / 2), got {}",
        lv
    );

    let sp = find_run(&doc, "spacing-40tw");
    let cs = sp.style.character_spacing.expect("character_spacing must populate");
    assert!(
        (cs - 2.0).abs() < 0.001,
        "spacing val=40 → 2.0pt (twips / 20), got {}",
        cs
    );

    let scale = find_run(&doc, "scale80");
    let ts = scale.style.text_scale.expect("text_scale must populate");
    assert!(
        (ts - 80.0).abs() < 0.001,
        "w val=80 → text_scale=80.0 (raw percentage, no conversion), got {}",
        ts
    );
}

#[test]
// Renamed from `..._vanish_alias` (2026-07-17): webHidden is NOT a vanish
// alias — see the webHidden block below (ISO/IEC 29500-1 §17.3.2.44).
fn v1_rfonts_cjk_separation_caps_independence_webhidden_not_vanish() {
    let Some(doc) = load("v1_rfonts_caps_vanish.docx") else { return };

    // rFonts ascii AND eastAsia → two SEPARATE fields, NOT a single
    // font_family that collapses both. has_explicit_east_asia=true
    // gates §4.6.3 widening (S136) and must reflect the EXPLICIT
    // eastAsia attribute (not a theme fallback).
    let arial = find_run(&doc, "arial-cjk");
    assert_eq!(
        arial.style.font_family.as_deref(),
        Some("Arial"),
        "rFonts ascii=Arial → font_family=Arial"
    );
    assert_eq!(
        arial.style.font_family_east_asia.as_deref(),
        Some("ＭＳ Ｐ明朝"),
        "rFonts eastAsia preserved verbatim in font_family_east_asia (UTF-8 CJK)"
    );
    assert!(
        arial.style.has_explicit_east_asia,
        "explicit eastAsia attr → has_explicit_east_asia=true (gates S136 widening)"
    );

    // smallCaps and caps are INDEPENDENT flags. A regression that
    // collapsed them (e.g., "caps OR smallCaps") would not catch
    // a downstream consumer that needs to distinguish "ALL caps"
    // from "small caps" rendering.
    let sc = find_run(&doc, "smallcaps-caps");
    assert!(sc.style.small_caps, "<w:smallCaps/> → small_caps=true");
    assert!(sc.style.all_caps, "<w:caps/> → all_caps=true");

    // webHidden is NOT an alias for vanish — it is a NO-OP in a paginated
    // renderer. ★Do not "fix" this back to `vanish=true`: the parser
    // deliberately splits them at all three sites (ooxml.rs:6201 run rPr,
    // ooxml.rs:2596 the ¶-mark pPr/rPr, styles.rs:637 style rPr) —
    // commit 8166ea9e, "ToC page-number rendering (uk_local_spending)".
    //
    // ISO/IEC 29500-1 §17.3.2.44 webHidden (Web Hidden Text), verbatim:
    //   "This element specifies whether the contents of this run shall be
    //    hidden from display at display time in a document WHEN THE DOCUMENT
    //    IS BEING DISPLAYED IN A WEB PAGE VIEW. ... As well, this setting
    //    SHOULD NOT AFFECT A NORMAL PAGINATED VIEW of the document."
    // Oxi renders the paginated view (it clones Word's print/PDF layout), so
    // webHidden must NOT hide. §17.3.2.45 vanish is the all-views flag.
    //
    // This is load-bearing, not pedantry: Word marks ToC tab-leaders and
    // PAGEREF page-number runs webHidden (they are meaningless in Web Layout,
    // which has no pages). Treating it as vanish dropped EVERY ToC page number
    // in uk_local_spending. The ¶-mark site matters too: webHidden must not
    // trigger the S673v hidden-mark empty-paragraph collapse.
    let wh = find_run(&doc, "hidden-webhidden");
    assert!(
        !wh.style.vanish,
        "<w:webHidden/> is WEB-LAYOUT-ONLY (§17.3.2.44) — it must NOT set vanish, \
         which would hide it in the paginated render (ToC page numbers)"
    );

    // vanish (§17.3.2.45) IS the all-views hidden flag.
    let v = find_run(&doc, "hidden-vanish");
    assert!(v.style.vanish, "<w:vanish/> → vanish=true");

    // Run-level shading (separate from paragraph shading). fill stored
    // verbatim as hex. "auto" suppression is already exercised on
    // color; this case pins the explicit-hex path on shd.
    let shd = find_run(&doc, "shaded");
    assert_eq!(
        shd.style.shading.as_deref(),
        Some("FFFF00"),
        "shd fill=FFFF00 → run-level shading hex verbatim"
    );
}

#[test]
fn all_five_fixtures_parse_and_preserve_run_count() {
    // Smoke + structural: each fixture parses without error and the
    // run text is preserved in document order. Catches a future
    // regression that drops or reorders runs while inline rPr is
    // present (each run carries its own rPr in these fixtures).
    let cases: &[(&str, &[&str])] = &[
        (
            "v1_bold_italic_underline.docx",
            &["plain", "bold-italic", "underline-single", "underline-none"],
        ),
        (
            "v1_strike_dstrike_vertalign.docx",
            &["plain", "strike", "dstrike", "super", "sub"],
        ),
        (
            "v1_sz_color_highlight.docx",
            &[
                "sz22",
                "sz23-half",
                "color-auto",
                "color-red",
                "highlight-yellow",
            ],
        ),
        (
            "v1_kern_position_spacing.docx",
            &[
                "kern22",
                "pos-raised",
                "pos-lowered",
                "spacing-40tw",
                "scale80",
            ],
        ),
        (
            "v1_rfonts_caps_vanish.docx",
            &[
                "arial-cjk",
                "smallcaps-caps",
                "hidden-webhidden",
                "hidden-vanish",
                "shaded",
            ],
        ),
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
        let texts: Vec<&str> = collect_runs(&doc).iter().map(|r| r.text.as_str()).collect();
        assert_eq!(&texts, expected, "{} run text order", name);
    }
}
