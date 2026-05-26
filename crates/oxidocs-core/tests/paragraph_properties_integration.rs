//! Integration tests: parse `<w:p><w:pPr>...</w:pPr></w:p>` end-to-end
//! and verify `Paragraph.{alignment, style}` shape after `parse_docx`.
//!
//! `parse_paragraph_properties` (parser/ooxml.rs:1746) is the largest
//! single parser function in oxidocs. Several sub-features have
//! dedicated integration suites:
//!   - <w:pBdr>      → paragraph_borders_integration   (S302)
//!   - <w:tabs>      → tabstops_integration            (S286)
//!   - <w:numPr>     → numbering_integration           (S272)
//!   - <w:sectPr>    → section_integration             (S279)
//!   - <w:pPrChange> → property_change_integration     (S306)
//!
//! This file fills the remaining INLINE pPr surface — jc aliases,
//! ind twip-priority, spacing modes (incl. negative-line Word quirk),
//! widowControl has_explicit tracking, and the four boolean toggles
//! (wordWrap / autoSpaceDE / autoSpaceDN / snapToGrid).
//!
//! Non-obvious behaviors pinned:
//!   - jc aliases (parser/ooxml.rs:1998-2004): "left"/"start" → Left,
//!     "right"/"end" → Right, "both" → Justify (counter-intuitive
//!     name), "distribute" → Distribute. All five enum-reachable
//!     variants exercised from one fixture.
//!   - ind twip-priority (CLAUDE.md 2026-04-10, parser/ooxml.rs:
//!     2175-2180): when BOTH `left` (twip) AND `leftChars` are
//!     declared, twip wins; `indent_left_chars` STAYS None even
//!     though the attribute is present. Reason: Word pre-computes
//!     the twip from the char count + font metrics, so twip is
//!     authoritative.
//!   - ind hanging → NEGATIVE indent_first_line (parser/ooxml.rs:
//!     2158). "hanging" is the OOXML name for a negative first-line
//!     indent. hangingChars mirrors this sign-flip
//!     (parser/ooxml.rs:2186).
//!   - line spacing rules diverge:
//!     * lineRule="auto" → line_spacing = line/240 (multiplier)
//!     * lineRule="exact"/"atLeast" → line_spacing = line/20 (pt)
//!     * negative `line` + missing-or-auto lineRule → Word quirk
//!       (COM-confirmed S114 2026-05-15 on d4d126):
//!       treated as wdLineSpaceExactly with |val|/20 pt. The parser
//!       FABRICATES `line_spacing_rule = "exact"` to match Word.
//!   - widowControl has THREE distinguishable states:
//!     missing (default true, has_explicit=false) ≠ `<w:widowControl/>`
//!     (true, has_explicit=true) ≠ `<w:widowControl w:val="0"/>`
//!     (false, has_explicit=true). Inheritance and OXI_FORCE_WIDOW
//!     dispatch on the explicit flag.
//!   - Four boolean toggles (word_wrap / auto_space_de / auto_space_dn
//!     / snap_to_grid) all default TRUE in `ParagraphStyle::default()`
//!     and flip to FALSE when w:val="0". A regression on the polarity
//!     would silently affect EVERY paragraph in EVERY doc.
//!
//! Fixtures live in `tools/fixtures/paragraph_properties_samples/`
//! and are authored by
//! `tools/metrics/build_paragraph_properties_repro_fixtures.py`.

use std::fs;

use oxidocs_core::ir::{Alignment, Block, Document, Paragraph};
use oxidocs_core::parse_docx;

fn fixture_path(name: &str) -> std::path::PathBuf {
    let crate_root = std::env::current_dir().unwrap();
    let workspace_root = crate_root.parent().unwrap().parent().unwrap();
    workspace_root
        .join("tools")
        .join("fixtures")
        .join("paragraph_properties_samples")
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

fn find_para<'a>(doc: &'a Document, text: &str) -> &'a Paragraph {
    paragraphs(doc)
        .into_iter()
        .find(|p| p.runs.iter().any(|r| r.text == text))
        .unwrap_or_else(|| panic!("paragraph with run text {:?} not found", text))
}

#[test]
fn v1_pp_jc_aliases_cover_all_alignment_variants() {
    let Some(doc) = load("v1_pp_jc_aliases.docx") else { return };

    assert_eq!(find_para(&doc, "jc-left").alignment, Alignment::Left);
    assert_eq!(
        find_para(&doc, "jc-start-alias").alignment,
        Alignment::Left,
        "jc val=\"start\" is an alias for Left (NOT a separate Start variant)"
    );
    assert_eq!(find_para(&doc, "jc-right").alignment, Alignment::Right);
    assert_eq!(
        find_para(&doc, "jc-end-alias").alignment,
        Alignment::Right,
        "jc val=\"end\" is an alias for Right"
    );
    assert_eq!(find_para(&doc, "jc-center").alignment, Alignment::Center);
    assert_eq!(
        find_para(&doc, "jc-both-equals-justify").alignment,
        Alignment::Justify,
        "jc val=\"both\" is OOXML's name for Justify (NOT a separate variant)"
    );
    assert_eq!(
        find_para(&doc, "jc-distribute").alignment,
        Alignment::Distribute
    );
}

#[test]
fn v1_pp_ind_twip_priority_and_hanging_sign_flip() {
    let Some(doc) = load("v1_pp_ind_twip_priority.docx") else { return };

    // Paragraph 1: BOTH `left="720"` AND `leftChars="200"`. Twip
    // wins — leftChars is NOT stored (CLAUDE.md 2026-04-10 rule).
    let twip_wins = find_para(&doc, "twip-wins-over-chars");
    let il = twip_wins
        .style
        .indent_left
        .expect("left=720 must populate indent_left");
    assert!(
        (il - 36.0).abs() < 0.001,
        "left=720 → 36pt (twips/20), got {}",
        il
    );
    assert!(
        twip_wins.style.indent_left_chars.is_none(),
        "twip-priority: leftChars SUPPRESSED when twip left also present \
         (CLAUDE.md 2026-04-10 rule, parser/ooxml.rs:2175-2180)"
    );

    // Paragraph 2: ONLY leftChars (no twip). Now stored — there's
    // no authoritative twip to defer to.
    let chars_alone = find_para(&doc, "chars-alone-stored");
    assert!(
        chars_alone.style.indent_left.is_none(),
        "no twip left → indent_left=None"
    );
    let lc = chars_alone
        .style
        .indent_left_chars
        .expect("leftChars alone IS stored (no twip to defer to)");
    assert!(
        (lc - 200.0).abs() < 0.001,
        "leftChars stored raw (hundredths of a char unit), got {}",
        lc
    );

    // Paragraph 3: hanging → NEGATIVE indent_first_line.
    let hanging = find_para(&doc, "hanging-becomes-negative");
    let fl = hanging
        .style
        .indent_first_line
        .expect("hanging populates indent_first_line");
    assert!(
        (fl - (-6.0)).abs() < 0.001,
        "hanging=120 → indent_first_line=-6.0pt (negative, parser/ooxml.rs:2158), got {}",
        fl
    );

    // Paragraph 4: hangingChars also negative (mirrors hanging).
    let hanging_chars = find_para(&doc, "hanging-chars-also-negative");
    let flc = hanging_chars
        .style
        .indent_first_line_chars
        .expect("hangingChars populates indent_first_line_chars");
    assert!(
        (flc - (-100.0)).abs() < 0.001,
        "hangingChars=100 → indent_first_line_chars=-100 (parser/ooxml.rs:2186), got {}",
        flc
    );
}

#[test]
fn v1_pp_spacing_three_modes_plus_negative_line_quirk() {
    let Some(doc) = load("v1_pp_spacing_line_modes.docx") else { return };

    // Paragraph 1: lineRule=auto → line/240 (multiplier).
    let auto = find_para(&doc, "line-auto-2x");
    let ls_auto = auto.style.line_spacing.expect("line populates");
    assert!(
        (ls_auto - 2.0).abs() < 0.001,
        "line=480 lineRule=auto → 2.0 multiplier (val/240), got {}",
        ls_auto
    );
    assert!(
        auto.style.line_spacing_rule.is_none(),
        "lineRule=auto does NOT populate line_spacing_rule (rule stays None for the auto path)"
    );

    // Paragraph 2: lineRule=exact → line/20 (pt), rule="exact".
    let exact = find_para(&doc, "line-exact-12pt");
    let ls_exact = exact.style.line_spacing.expect("line populates");
    assert!(
        (ls_exact - 12.0).abs() < 0.001,
        "line=240 lineRule=exact → 12pt (val/20), got {}",
        ls_exact
    );
    assert_eq!(
        exact.style.line_spacing_rule.as_deref(),
        Some("exact"),
        "exact rule materialized on style"
    );

    // Paragraph 3: line=-240 with NO lineRule — Word quirk.
    // COM-confirmed S114 2026-05-15 on d4d126: treated as
    // wdLineSpaceExactly |val|/20. Parser FABRICATES rule="exact"
    // to match Word.
    let neg = find_para(&doc, "line-negative-equals-exact");
    let ls_neg = neg.style.line_spacing.expect("negative line populates");
    assert!(
        (ls_neg - 12.0).abs() < 0.001,
        "line=-240 (no lineRule) → 12pt EXACT (Word quirk, |val|/20), got {}",
        ls_neg
    );
    assert_eq!(
        neg.style.line_spacing_rule.as_deref(),
        Some("exact"),
        "Word quirk FABRICATES rule=\"exact\" for negative line (parser/ooxml.rs:2107-2113)"
    );

    // Paragraph 4: before/after standalone (twips → pt).
    let ba = find_para(&doc, "before-after-twips");
    let sb = ba.style.space_before.expect("before populates");
    let sa = ba.style.space_after.expect("after populates");
    assert!((sb - 6.0).abs() < 0.001, "before=120 → 6pt, got {}", sb);
    assert!((sa - 12.0).abs() < 0.001, "after=240 → 12pt, got {}", sa);
}

#[test]
fn v1_pp_widow_control_three_distinguishable_states() {
    let Some(doc) = load("v1_pp_widow_control_explicit.docx") else { return };

    // Missing widowControl → default true, has_explicit=false.
    let none = find_para(&doc, "no-widow-attr");
    assert!(none.style.widow_control, "default widow_control = true");
    assert!(
        !none.style.has_explicit_widow_control,
        "missing attr → has_explicit_widow_control=false (downstream inheritance can still apply)"
    );

    // `<w:widowControl/>` (no val) → true, has_explicit=true.
    let on = find_para(&doc, "explicit-on");
    assert!(on.style.widow_control, "presence with no val → true");
    assert!(
        on.style.has_explicit_widow_control,
        "presence → has_explicit_widow_control=true (DISTINCT from missing)"
    );

    // `<w:widowControl w:val="0"/>` → false, has_explicit=true.
    let off = find_para(&doc, "explicit-off");
    assert!(!off.style.widow_control, "val=0 → widow_control=false");
    assert!(
        off.style.has_explicit_widow_control,
        "explicit off also marked as explicit (drives OXI_FORCE_WIDOW dispatch)"
    );
}

#[test]
fn v1_pp_four_boolean_toggles_default_true_flip_to_false_on_val_zero() {
    // The four toggles — wordWrap / autoSpaceDE / autoSpaceDN /
    // snapToGrid — all default TRUE in ParagraphStyle::default()
    // and flip to FALSE on val="0". A regression that inverted the
    // polarity would silently affect EVERY paragraph in EVERY doc.
    let Some(doc) = load("v1_pp_word_wrap_autospace.docx") else { return };

    let off = find_para(&doc, "four-off");
    assert!(
        !off.style.word_wrap,
        "<w:wordWrap w:val=\"0\"/> → word_wrap=false (S301 discriminator)"
    );
    assert!(
        !off.style.auto_space_de,
        "<w:autoSpaceDE w:val=\"0\"/> → auto_space_de=false"
    );
    assert!(
        !off.style.auto_space_dn,
        "<w:autoSpaceDN w:val=\"0\"/> → auto_space_dn=false"
    );
    assert!(
        !off.style.snap_to_grid,
        "<w:snapToGrid w:val=\"0\"/> → snap_to_grid=false"
    );

    // Bare paragraph (no pPr toggles) → all four stay at the Rust-
    // side default (true). A regression that defaulted any to false
    // would surface here.
    let on = find_para(&doc, "four-default-on");
    assert!(on.style.word_wrap, "default word_wrap=true");
    assert!(on.style.auto_space_de, "default auto_space_de=true");
    assert!(on.style.auto_space_dn, "default auto_space_dn=true");
    assert!(on.style.snap_to_grid, "default snap_to_grid=true");
}

#[test]
fn all_five_fixtures_parse_with_expected_paragraph_counts() {
    let cases: &[(&str, usize)] = &[
        ("v1_pp_jc_aliases.docx", 7),
        ("v1_pp_ind_twip_priority.docx", 4),
        ("v1_pp_spacing_line_modes.docx", 4),
        ("v1_pp_widow_control_explicit.docx", 3),
        ("v1_pp_word_wrap_autospace.docx", 2),
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
        assert_eq!(
            paragraphs(&doc).len(),
            *expected,
            "{} paragraph count",
            name
        );
    }
}
