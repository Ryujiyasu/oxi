//! Integration tests: parse `word/fontTable.xml` and verify
//! `Document.styles.font_table` populates with the expected
//! `FontInfo` records keyed by `<w:font w:name="...">`.
//!
//! Parser code path tested:
//!   - [parser/ooxml.rs:475](crates/oxidocs-core/src/parser/ooxml.rs#L475)
//!     `parse_font_table` opens each `<w:font>` Start element, reads
//!     the `w:name` attribute as HashMap key, then collects child
//!     `<w:panose1>` / `<w:charset>` / `<w:family>` / `<w:pitch>` Empty
//!     elements' `w:val` attributes verbatim. Entry is committed to
//!     `StyleSheet.font_table` on the corresponding `</w:font>`.
//!   - Each FontInfo field is `Option<String>` — None when the child
//!     element is absent (no normalization, no defaulting).
//!   - `word/fontTable.xml` absent → `font_table` is an empty HashMap.
//!
//! Fixtures live in `tools/fixtures/font_info_samples/` and are
//! authored by `tools/metrics/build_font_info_repro_fixtures.py` (S303).

use std::fs;

use oxidocs_core::ir::Document;
use oxidocs_core::parser::parse_docx;

fn fixture_path(name: &str) -> std::path::PathBuf {
    let crate_root = std::env::current_dir().unwrap();
    let workspace_root = crate_root.parent().unwrap().parent().unwrap();
    workspace_root
        .join("tools")
        .join("fixtures")
        .join("font_info_samples")
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

#[test]
fn v1_basic_carries_all_four_fields_for_calibri() {
    let Some(doc) = load("v1_basic.docx") else { return };
    let info = doc
        .styles
        .font_table
        .get("Calibri")
        .expect("Calibri entry in font_table");
    assert_eq!(
        info.panose1.as_deref(),
        Some("020F0502020204030204"),
        "PANOSE-1 hex preserved verbatim"
    );
    assert_eq!(info.charset.as_deref(), Some("00"), "ANSI charset");
    assert_eq!(info.family.as_deref(), Some("swiss"));
    assert_eq!(info.pitch.as_deref(), Some("variable"));
}

#[test]
fn v1_partial_leaves_omitted_fields_as_none() {
    // Only `<w:panose1>` was emitted; other child elements absent.
    // FontInfo fields are Option<String>, so missing children stay None
    // (no defaulting, no empty-string fallback).
    let Some(doc) = load("v1_partial.docx") else { return };
    let info = doc
        .styles
        .font_table
        .get("ＭＳ 明朝")
        .expect("CJK-keyed font entry");
    assert_eq!(info.panose1.as_deref(), Some("02020609040205080304"));
    assert!(info.charset.is_none(), "no charset child → None");
    assert!(info.family.is_none(), "no family child → None");
    assert!(info.pitch.is_none(), "no pitch child → None");
}

#[test]
fn v1_multiple_fonts_keys_each_entry_independently() {
    // 3 <w:font> blocks in the same fontTable.xml become 3 HashMap
    // entries; per-field values stay attached to the correct key
    // (no cross-wiring between consecutive entries).
    let Some(doc) = load("v1_multiple_fonts.docx") else { return };
    let ft = &doc.styles.font_table;
    assert_eq!(ft.len(), 3, "three independent entries");

    let times = ft.get("Times New Roman").expect("Times present");
    assert_eq!(times.family.as_deref(), Some("roman"));
    assert_eq!(times.pitch.as_deref(), Some("variable"));

    let courier = ft.get("Courier New").expect("Courier present");
    assert_eq!(courier.family.as_deref(), Some("modern"));
    assert_eq!(courier.pitch.as_deref(), Some("fixed"));

    let wingdings = ft.get("Wingdings").expect("Wingdings present");
    assert_eq!(
        wingdings.charset.as_deref(),
        Some("02"),
        "Symbol charset preserved as hex"
    );
    assert_eq!(wingdings.family.as_deref(), Some("auto"));
    assert_eq!(wingdings.pitch.as_deref(), Some("default"));

    // PANOSE strings stay attached to their own font (no cross-wire).
    assert_eq!(
        times.panose1.as_deref(),
        Some("02020603050405020304")
    );
    assert_eq!(
        courier.panose1.as_deref(),
        Some("02070309020205020404")
    );
}

#[test]
fn v1_panose_verbatim_preserves_mixed_case_hex() {
    // The parser stores `w:val` verbatim. Even mixed-case hex
    // (lowercase "abcdef" + uppercase "ABCD") is kept as-is — no
    // .to_uppercase(), no zero-padding, no stripping.
    let Some(doc) = load("v1_panose_verbatim.docx") else { return };
    let info = doc
        .styles
        .font_table
        .get("PanoseTest")
        .expect("PanoseTest entry");
    assert_eq!(
        info.panose1.as_deref(),
        Some("abcdef0123456789ABCD"),
        "PANOSE hex stored verbatim with original casing"
    );
    let p = info.panose1.as_deref().unwrap();
    assert_eq!(p.len(), 20, "PANOSE-1 always 20 hex chars (10 bytes)");
}

#[test]
fn v1_no_fonttable_yields_empty_font_table() {
    // word/fontTable.xml absent → read_part returns Err and
    // parse_font_table returns early without inserting anything.
    let Some(doc) = load("v1_no_fonttable.docx") else { return };
    assert!(
        doc.styles.font_table.is_empty(),
        "missing fontTable.xml → empty HashMap (not Option::None)"
    );
}

#[test]
fn all_five_fixtures_parse_with_expected_entry_count() {
    let cases: &[(&str, usize)] = &[
        ("v1_basic.docx", 1),
        ("v1_partial.docx", 1),
        ("v1_multiple_fonts.docx", 3),
        ("v1_panose_verbatim.docx", 1),
        ("v1_no_fonttable.docx", 0),
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
            doc.styles.font_table.len(),
            *expected,
            "{} font_table size",
            name
        );
    }
}
