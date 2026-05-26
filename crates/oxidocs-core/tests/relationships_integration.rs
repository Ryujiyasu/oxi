//! Integration tests: parse `.rels` XML directly via `parse_relationships`.
//!
//! `parse_relationships` is the public entry for `_rels/.rels`, `word/_rels/
//! document.xml.rels`, header/footer .rels, etc. It returns a
//! `HashMap<Id, Relationship>` consumed internally by parser/ooxml.rs to
//! resolve image, hyperlink, header, footer, comments, footnote, endnote,
//! settings, theme, fontTable, numbering, styles, customXml, etc. references.
//!
//! Parser code paths tested:
//!   - [parser/relationships.rs:18](crates/oxidocs-core/src/parser/relationships.rs#L18)
//!     `parse_relationships` reads `<Relationship Id Type Target/>` elements
//!     and returns `HashMap<Id, Relationship>`.
//!
//! Non-obvious behaviors pinned:
//!   - Both `<Relationship/>` self-closing AND `<Relationship>...</Relationship>`
//!     paired start/end are accepted (Event::Empty | Event::Start match).
//!     Load-bearing: real .rels files mix both forms depending on writer.
//!   - **Empty Id is SKIPPED** (not inserted). Prevents `HashMap` collision
//!     on missing `Id` attribute (the HashMap key would otherwise be `""`
//!     and last-such-Relationship would silently overwrite the previous).
//!   - **Namespace prefixes stripped via `local_name`**: `<r:Relationship>`
//!     and `<Relationships:Relationship>` are both treated as Relationship.
//!     A regression that compared `e.name()` byte-for-byte would silently
//!     drop ALL relationships from namespaced .rels files.
//!   - **Duplicate Id: LAST WINS** (HashMap.insert overwrites). Validates
//!     that the iteration order — root → child order — determines the
//!     surviving entry. A change to "first wins" would silently flip
//!     resolution for any .rels containing duplicate Ids (rare in valid
//!     OOXML but observed in legacy/buggy writers).
//!   - **Unknown attributes ignored** (TargetMode, Schema, etc.) — forward
//!     compatibility with OPC extensions.
//!   - **Target stored verbatim** (no URL decoding, no `xml:base` resolution,
//!     no leading-slash normalization). Image references like
//!     `media/image1.png` and absolute `/word/media/image2.png` survive.

use oxidocs_core::parser::relationships::parse_relationships;

/// Fixture: minimal single self-closing Relationship.
const FIXTURE_SELF_CLOSING: &str = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>"#;

/// Fixture: paired Start/End form (some writers prefer this).
const FIXTURE_PAIRED: &str = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/image1.png"></Relationship>
</Relationships>"#;

/// Fixture: empty `Id` — should be SKIPPED (not inserted).
const FIXTURE_EMPTY_ID: &str = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/skipped.png"/>
  <Relationship Id="rIdOK" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/kept.png"/>
</Relationships>"#;

/// Fixture: duplicate Id — LAST WINS.
const FIXTURE_DUPLICATE_ID: &str = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rIdDup" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/first.png"/>
  <Relationship Id="rIdDup" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/last.png"/>
</Relationships>"#;

/// Fixture: namespace-prefixed Relationship element.
const FIXTURE_NAMESPACED: &str = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<r:Relationships xmlns:r="http://schemas.openxmlformats.org/package/2006/relationships">
  <r:Relationship Id="rIdNs" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" Target="https://example.com"/>
</r:Relationships>"#;

/// Fixture: unknown attributes (TargetMode, custom Schema) — must be ignored.
const FIXTURE_UNKNOWN_ATTRS: &str = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rIdExt" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" Target="https://example.com" TargetMode="External" Schema="ignored"/>
</Relationships>"#;

#[test]
fn self_closing_relationship_inserted() {
    let rels = parse_relationships(FIXTURE_SELF_CLOSING).expect("parse ok");
    assert_eq!(rels.len(), 1);
    let r = rels.get("rId1").expect("rId1 present");
    assert_eq!(r.target, "word/document.xml");
    assert_eq!(
        r.rel_type,
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument"
    );
}

#[test]
fn paired_start_end_relationship_inserted() {
    let rels = parse_relationships(FIXTURE_PAIRED).expect("parse ok");
    assert_eq!(rels.len(), 1, "paired Start/End form must be accepted");
    let r = rels.get("rId2").expect("rId2 present");
    assert_eq!(r.target, "media/image1.png");
}

#[test]
fn empty_id_is_skipped() {
    let rels = parse_relationships(FIXTURE_EMPTY_ID).expect("parse ok");
    assert_eq!(rels.len(), 1, "empty-Id entry must NOT be inserted");
    assert!(rels.get("").is_none(), "empty key must not exist");
    assert!(rels.get("rIdOK").is_some(), "non-empty entry survives");
    assert_eq!(rels["rIdOK"].target, "media/kept.png");
}

#[test]
fn duplicate_id_last_wins() {
    let rels = parse_relationships(FIXTURE_DUPLICATE_ID).expect("parse ok");
    assert_eq!(rels.len(), 1, "duplicate Id collapses to one entry");
    let r = rels.get("rIdDup").expect("rIdDup present");
    assert_eq!(r.target, "media/last.png", "last <Relationship> wins via HashMap.insert");
}

#[test]
fn namespace_prefix_stripped() {
    let rels = parse_relationships(FIXTURE_NAMESPACED).expect("parse ok");
    assert_eq!(rels.len(), 1, "r:Relationship must be recognized (local_name strips prefix)");
    let r = rels.get("rIdNs").expect("rIdNs present");
    assert_eq!(r.target, "https://example.com");
}

#[test]
fn unknown_attributes_ignored() {
    let rels = parse_relationships(FIXTURE_UNKNOWN_ATTRS).expect("parse ok");
    let r = rels.get("rIdExt").expect("rIdExt present");
    // TargetMode and Schema not captured — only Id/Type/Target survive
    assert_eq!(r.target, "https://example.com");
    assert!(
        r.rel_type.contains("hyperlink"),
        "rel_type unaffected by extra attributes"
    );
}
