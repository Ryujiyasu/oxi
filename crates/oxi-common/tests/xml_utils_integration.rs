//! Integration tests for `oxi_common::xml_utils`.
//!
//! Pins the 6 shared utility functions (local_name, get_attr, get_raw_attr,
//! twips_to_pt, half_pt_to_pt, emu_to_pt) used across oxicells-core and
//! oxislides-core parsers/editors. These functions sit on the hot path of
//! every XML element read in those crates — a silent regression here would
//! silently corrupt sheet / slide parsing without a single failing assertion
//! elsewhere.
//!
//! Coverage anchors (S336, 2026-05-27):
//! - local_name: rfind(':') semantics — last colon wins, no-colon = identity,
//!   empty-input contract, multi-colon edge case.
//! - get_attr: namespaced lookup by LOCAL name (`w:val` matched as `val`).
//! - get_raw_attr: lookup by raw key OR `ends_with(":<key>")` — must NOT
//!   confuse `customval` with `val`.
//! - twips_to_pt / half_pt_to_pt / emu_to_pt: unit conversion constants
//!   (1440 tw = 72 pt = 914400 EMU; 21 hp = 10.5 pt — Word's size=22 ⇒ 11pt).
//!
//! Numerical constants are NOT speculative — they are derived from the OOXML
//! spec (ECMA-376) and are stable. Any regression here flags a constant drift.

use oxi_common::xml_utils::{emu_to_pt, get_attr, get_raw_attr, half_pt_to_pt, local_name, twips_to_pt};
use quick_xml::events::Event;
use quick_xml::reader::Reader;

// ---------------------------------------------------------------------------
// local_name: extract local name from a potentially-namespaced XML tag.
// Implementation uses rfind(':'); semantics pinned below.
// ---------------------------------------------------------------------------

#[test]
fn local_name_strips_namespace_prefix() {
    assert_eq!(local_name(b"w:body"), "body");
    assert_eq!(local_name(b"r:embed"), "embed");
    assert_eq!(local_name(b"xmlns:w"), "w");
}

#[test]
fn local_name_returns_identity_when_no_colon() {
    assert_eq!(local_name(b"body"), "body");
    assert_eq!(local_name(b"Relationship"), "Relationship");
}

#[test]
fn local_name_handles_empty_input() {
    // Pinned contract: empty bytes → empty String (not panic, not None).
    // Downstream code in parsers relies on this for short-circuit comparisons
    // (e.g., `if local_name(name) == "Relationship" { ... }`).
    assert_eq!(local_name(b""), "");
}

#[test]
fn local_name_multi_colon_uses_last_colon() {
    // rfind(':') means the LAST colon is the separator. This matters for
    // pathological inputs like `a:b:c` — local part is `c`, not `b:c`.
    assert_eq!(local_name(b"a:b:c"), "c");
}

#[test]
fn local_name_handles_invalid_utf8_as_empty() {
    // local_name uses `std::str::from_utf8(...).unwrap_or("")`. Pin this
    // graceful fallback so a malformed byte sequence never panics inside the
    // parser hot loop.
    let bad: &[u8] = &[0xff, 0xfe, 0xfd];
    assert_eq!(local_name(bad), "");
}

// ---------------------------------------------------------------------------
// get_attr: lookup by LOCAL attribute name (namespace-insensitive).
// ---------------------------------------------------------------------------

fn first_start_event(xml: &str) -> quick_xml::events::BytesStart<'static> {
    // Drive the reader until the first Start/Empty event and return an owned
    // copy. Tests stay isolated even though quick_xml's events borrow.
    let mut reader = Reader::from_str(xml);
    loop {
        match reader.read_event().expect("xml parse") {
            Event::Start(e) => return e.into_owned(),
            Event::Empty(e) => return e.into_owned(),
            Event::Eof => panic!("no Start/Empty event in: {xml}"),
            _ => continue,
        }
    }
}

#[test]
fn get_attr_finds_namespaced_attribute_by_local_name() {
    // `w:val="11"` must be findable as "val".
    let e = first_start_event(r#"<w:sz w:val="22"/>"#);
    assert_eq!(get_attr(&e, "val").as_deref(), Some("22"));
}

#[test]
fn get_attr_finds_unprefixed_attribute() {
    let e = first_start_event(r#"<Relationship Id="rId1" Target="word/document.xml"/>"#);
    assert_eq!(get_attr(&e, "Id").as_deref(), Some("rId1"));
    assert_eq!(get_attr(&e, "Target").as_deref(), Some("word/document.xml"));
}

#[test]
fn get_attr_returns_none_for_missing_attribute() {
    let e = first_start_event(r#"<w:sz w:val="22"/>"#);
    assert!(get_attr(&e, "nonexistent").is_none());
}

#[test]
fn get_attr_returns_first_match_when_duplicate_local_names() {
    // If two attributes share a local name (e.g., `w:val` AND `r:val`), the
    // current implementation iterates `e.attributes().flatten()` and returns
    // the first match. Pin this so future "namespace-aware" changes are
    // intentional (renderers depend on encounter-order semantics).
    let e = first_start_event(r#"<x w:val="first" r:val="second"/>"#);
    assert_eq!(get_attr(&e, "val").as_deref(), Some("first"));
}

// ---------------------------------------------------------------------------
// get_raw_attr: lookup by RAW key, OR by suffix `:<key>` (handles `xmlns:`).
// ---------------------------------------------------------------------------

#[test]
fn get_raw_attr_matches_exact_key() {
    let e = first_start_event(r#"<Relationship Id="rId1"/>"#);
    assert_eq!(get_raw_attr(&e, "Id").as_deref(), Some("rId1"));
}

#[test]
fn get_raw_attr_matches_namespaced_via_suffix() {
    // Raw key `w:val` ends_with `:val` → matches when searching for "val".
    let e = first_start_event(r#"<w:sz w:val="22"/>"#);
    assert_eq!(get_raw_attr(&e, "val").as_deref(), Some("22"));
}

#[test]
fn get_raw_attr_does_not_match_inner_substring() {
    // `customval` ends with `val` (substring) but NOT with `:val` (suffix).
    // Pin that get_raw_attr requires the colon to avoid false positives —
    // critical for OOXML where attributes like `customval` exist.
    let e = first_start_event(r#"<x customval="bad"/>"#);
    assert!(get_raw_attr(&e, "val").is_none());
}

// ---------------------------------------------------------------------------
// twips_to_pt: 1 pt = 20 twips (Word's whole-document unit).
// ---------------------------------------------------------------------------

#[test]
fn twips_to_pt_1440_equals_1_inch_72pt() {
    // 1440 twips = 1 inch = 72 pt (the unit Word uses for margins / page
    // dimensions in pageMar). This is an OOXML spec constant.
    assert_eq!(twips_to_pt(1440.0), 72.0);
}

#[test]
fn twips_to_pt_basic_ratio() {
    assert_eq!(twips_to_pt(20.0), 1.0);
    assert_eq!(twips_to_pt(0.0), 0.0);
    assert_eq!(twips_to_pt(-40.0), -2.0); // sign preservation
}

// ---------------------------------------------------------------------------
// half_pt_to_pt: Word stores font size as half-points (sz=22 ⇒ 11pt).
// ---------------------------------------------------------------------------

#[test]
fn half_pt_to_pt_word_font_size_semantics() {
    // Word's `<w:sz w:val="22"/>` means font size = 11pt.
    assert_eq!(half_pt_to_pt(22.0), 11.0);
    // 21 half-points = 10.5pt (MS Mincho default in many JP docs).
    assert_eq!(half_pt_to_pt(21.0), 10.5);
    assert_eq!(half_pt_to_pt(0.0), 0.0);
}

// ---------------------------------------------------------------------------
// emu_to_pt: 914400 EMU = 1 inch = 72 pt. Used for DrawingML image sizes.
// ---------------------------------------------------------------------------

#[test]
fn emu_to_pt_914400_equals_1_inch_72pt() {
    assert_eq!(emu_to_pt(914400.0), 72.0);
}

#[test]
fn emu_to_pt_basic_ratio() {
    assert_eq!(emu_to_pt(12700.0), 1.0);
    assert_eq!(emu_to_pt(0.0), 0.0);
}
