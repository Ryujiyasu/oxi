//! Integration tests: parse `word/theme/theme1.xml` end-to-end and
//! verify Run.style.color / Run.style.font_family_east_asia after
//! `parse_docx` resolves themeColor / eastAsiaTheme references.
//!
//! ThemeColors itself is private to the parser module, so the test
//! observes theme parsing INDIRECTLY through Run-level resolution
//! at parser/ooxml.rs:4362-4400 (rPr color) and 4269-4276 (rFonts).
//!
//! Parser code paths tested:
//!   - [parser/theme.rs:108](crates/oxidocs-core/src/parser/theme.rs#L108)
//!     `parse_theme` reads clrScheme (dk1/lt1/dk2/lt2/accent1-6/
//!     hlink/folHlink) and fontScheme (majorFont / minorFont).
//!   - [parser/theme.rs:22](crates/oxidocs-core/src/parser/theme.rs#L22)
//!     `ThemeColors::resolve` aliases:
//!     "dark1"|"text1" → "dk1";
//!     "light1"|"background1" → "lt1";
//!     "dark2"|"text2" → "dk2";
//!     "light2"|"background2" → "lt2";
//!     "hyperlink" → "hlink";
//!     "followedHyperlink" → "folHlink";
//!     other → passthrough.
//!   - [parser/theme.rs:153](crates/oxidocs-core/src/parser/theme.rs#L153)
//!     `sysClr` uses `lastClr` attribute (NOT `val`).
//!   - [parser/styles.rs:15](crates/oxidocs-core/src/parser/styles.rs#L15)
//!     `resolve_theme_font_pub` maps "minorEastAsia"/"majorEastAsia"
//!     → theme.minor_font_ea / major_font_ea.
//!
//! Non-obvious behaviors pinned:
//!   - When a run has BOTH `w:val` AND `w:themeColor` AND theme
//!     resolves, the THEME WINS. The val attribute is the fallback
//!     ONLY if theme.resolve() returns None
//!     (parser/ooxml.rs:4379-4400).
//!   - sysClr uses `lastClr` attribute, NOT `val` — distinct from
//!     srgbClr which uses `val`. A regression mixing the two would
//!     silently produce empty hex for sysClr-encoded scheme entries.
//!   - resolve() collapses TWO aliases per dk/lt scheme entry
//!     ("dark1"/"text1" both → "dk1"). A regression that dropped
//!     one alias would silently affect docs that use the other
//!     spelling.
//!   - Explicit ea typeface in theme PREVENTS the end-of-parse
//!     Meiryo fallback (theme.rs:251-256). When theme has explicit
//!     <a:ea typeface="ＭＳ Ｐ明朝"/>, the resolved font is
//!     verbatim — NOT replaced by Meiryo.

use std::fs;

use oxidocs_core::ir::{Block, Document, Paragraph, Run};
use oxidocs_core::parse_docx;

fn fixture_path(name: &str) -> std::path::PathBuf {
    let crate_root = std::env::current_dir().unwrap();
    let workspace_root = crate_root.parent().unwrap().parent().unwrap();
    workspace_root
        .join("tools")
        .join("fixtures")
        .join("theme_samples")
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

fn first_run(doc: &Document) -> &Run {
    let para = doc
        .pages
        .iter()
        .flat_map(|p| p.blocks.iter())
        .find_map(|b| if let Block::Paragraph(p) = b { Some(p) } else { None })
        .expect("first paragraph") as &Paragraph;
    para.runs.first().expect("first run")
}

#[test]
fn v1_theme_resolve_accent1_theme_wins_over_val() {
    let Some(doc) = load("v1_theme_resolve_accent1.docx") else { return };
    let r = first_run(&doc);

    // theme accent1="00FF00", w:val="FF0000". The theme MUST win
    // because parser/ooxml.rs:4379-4400 checks themeColor first
    // and assigns style.color from the resolved hex; the val
    // branch only fires when theme.resolve() returns None.
    assert_eq!(
        r.style.color.as_deref(),
        Some("00FF00"),
        "themeColor=accent1 with theme accent1=00FF00 → \"00FF00\" \
         (theme WINS over w:val=\"FF0000\")"
    );
}

#[test]
fn v1_theme_resolve_alias_text1_maps_to_dk1() {
    let Some(doc) = load("v1_theme_resolve_alias_text1.docx") else { return };
    let r = first_run(&doc);

    // resolve("text1") → "dk1". Theme dk1=123456. Pins the
    // "text1" alias arm at theme.rs:25.
    assert_eq!(
        r.style.color.as_deref(),
        Some("123456"),
        "themeColor=\"text1\" alias resolves to dk1 → \"123456\""
    );
}

#[test]
fn v1_theme_resolve_hyperlink_alias_maps_to_hlink() {
    let Some(doc) = load("v1_theme_resolve_hyperlink_alias.docx") else { return };
    let r = first_run(&doc);

    // resolve("hyperlink") → "hlink". Theme hlink=0563C1. Pins
    // the "hyperlink" alias arm at theme.rs:35.
    assert_eq!(
        r.style.color.as_deref(),
        Some("0563C1"),
        "themeColor=\"hyperlink\" alias resolves to hlink → \"0563C1\""
    );
}

#[test]
fn v1_theme_sysclr_uses_lastclr_attribute_not_val() {
    let Some(doc) = load("v1_theme_sysclr_lastclr.docx") else { return };
    let r = first_run(&doc);

    // Theme dk1 was encoded as <a:sysClr val="windowText"
    // lastClr="333333"/>. Parser at theme.rs:153-163 reads
    // `lastClr` (NOT `val`). val="windowText" is the named system
    // color and would be wrong as hex. A regression mixing val/lastClr
    // would silently surface "windowText" or empty.
    assert_eq!(
        r.style.color.as_deref(),
        Some("333333"),
        "<a:sysClr lastClr=\"333333\"/> uses `lastClr` (NOT `val`) → \"333333\""
    );
}

#[test]
fn v1_theme_font_minor_ea_resolve_avoids_meiryo_fallback() {
    let Some(doc) = load("v1_theme_font_minor_ea_resolve.docx") else { return };
    let r = first_run(&doc);

    // Theme has explicit <a:ea typeface="ＭＳ Ｐ明朝"/> on minorFont.
    // Run has <w:rFonts w:eastAsiaTheme="minorEastAsia"/>. Resolver
    // returns theme.minor_font_ea = "ＭＳ Ｐ明朝". The Meiryo
    // fallback at theme.rs:251-256 should NOT fire because the
    // theme set ea explicitly during parsing.
    assert_eq!(
        r.style.font_family_east_asia.as_deref(),
        Some("ＭＳ Ｐ明朝"),
        "explicit theme.minorFont.ea PREVENTS Meiryo fallback; \
         eastAsiaTheme=\"minorEastAsia\" resolves to verbatim \"ＭＳ Ｐ明朝\""
    );
}

#[test]
fn all_five_fixtures_parse_with_resolved_run_attribute() {
    let cases: &[&str] = &[
        "v1_theme_resolve_accent1.docx",
        "v1_theme_resolve_alias_text1.docx",
        "v1_theme_resolve_hyperlink_alias.docx",
        "v1_theme_sysclr_lastclr.docx",
        "v1_theme_font_minor_ea_resolve.docx",
    ];
    for name in cases {
        let path = fixture_path(name);
        if !path.exists() {
            eprintln!("skipping {}", name);
            continue;
        }
        let data = fs::read(&path).unwrap();
        let doc = parse_docx(&data)
            .unwrap_or_else(|e| panic!("failed to parse {}: {:?}", name, e));
        let r = first_run(&doc);
        // First four pin color; last one pins font_family_east_asia.
        let has_resolved = r.style.color.is_some()
            || r.style.font_family_east_asia.is_some();
        assert!(
            has_resolved,
            "{}: at least one of color/font_family_east_asia must be resolved \
             by theme parser",
            name,
        );
    }
}
