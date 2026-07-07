// This Source Code Form is subject to the terms of the Mozilla Public
// License, v. 2.0. If a copy of the MPL was not distributed with this
// file, You can obtain one at https://mozilla.org/MPL/2.0/.

//! Integration tests for oxi-common::relationships::parse_relationships.
//!
//! parse_relationships turns a `.rels` XML part into an Id→Relationship map.
//! It is used by every OOXML engine to resolve part targets (images, slides,
//! headers, hyperlinks). These tests pin attribute extraction, both element
//! forms (self-closing vs open/close), namespace tolerance, and the
//! empty/malformed contracts using inline XML (no fixture needed).

use oxidocs_common::relationships::parse_relationships;

const RELS_NS: &str =
    r#"http://schemas.openxmlformats.org/package/2006/relationships"#;

fn wrap(inner: &str) -> String {
    format!(r#"<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="{RELS_NS}">{inner}</Relationships>"#)
}

#[test]
fn parses_multiple_relationships() {
    let xml = wrap(
        r#"<Relationship Id="rId1" Type="http://x/officeDocument" Target="word/document.xml"/>
           <Relationship Id="rId2" Type="http://x/styles" Target="word/styles.xml"/>
           <Relationship Id="rId3" Type="http://x/image" Target="media/image1.png"/>"#,
    );
    let rels = parse_relationships(&xml).expect("parse");
    assert_eq!(rels.len(), 3);
    assert_eq!(rels["rId1"].target, "word/document.xml");
    assert_eq!(rels["rId2"].target, "word/styles.xml");
    assert_eq!(rels["rId3"].target, "media/image1.png");
}

#[test]
fn extracts_all_three_attributes() {
    let xml = wrap(r#"<Relationship Id="rId7" Type="http://example/hyperlink" Target="https://a.example/x"/>"#);
    let rels = parse_relationships(&xml).expect("parse");
    let r = &rels["rId7"];
    assert_eq!(r.id, "rId7");
    assert_eq!(r.rel_type, "http://example/hyperlink");
    assert_eq!(r.target, "https://a.example/x");
}

#[test]
fn open_close_element_form_also_parsed() {
    // Both <Relationship .../> (Empty) and <Relationship>...</Relationship>
    // (Start) forms must be handled. The latter is unusual but valid XML.
    let xml = wrap(r#"<Relationship Id="rId1" Type="t" Target="a.xml"></Relationship>"#);
    let rels = parse_relationships(&xml).expect("parse");
    assert_eq!(rels.len(), 1);
    assert_eq!(rels["rId1"].target, "a.xml");
}

#[test]
fn relationship_without_id_is_skipped() {
    // An entry with empty/absent Id is not inserted (would be unaddressable).
    let xml = wrap(
        r#"<Relationship Type="t" Target="orphan.xml"/>
           <Relationship Id="rId1" Type="t" Target="keep.xml"/>"#,
    );
    let rels = parse_relationships(&xml).expect("parse");
    assert_eq!(rels.len(), 1);
    assert!(rels.contains_key("rId1"));
}

#[test]
fn missing_type_or_target_yields_empty_strings() {
    // Absent Type/Target default to empty string (not an error).
    let xml = wrap(r#"<Relationship Id="rId1"/>"#);
    let rels = parse_relationships(&xml).expect("parse");
    assert_eq!(rels["rId1"].rel_type, "");
    assert_eq!(rels["rId1"].target, "");
}

#[test]
fn duplicate_id_last_wins() {
    let xml = wrap(
        r#"<Relationship Id="rId1" Type="t" Target="first.xml"/>
           <Relationship Id="rId1" Type="t" Target="second.xml"/>"#,
    );
    let rels = parse_relationships(&xml).expect("parse");
    assert_eq!(rels.len(), 1);
    assert_eq!(rels["rId1"].target, "second.xml");
}

#[test]
fn empty_relationships_element_yields_empty_map() {
    let xml = wrap("");
    let rels = parse_relationships(&xml).expect("parse");
    assert!(rels.is_empty());
}

#[test]
fn namespace_prefixed_element_parsed() {
    // local_name() strips the namespace prefix, so r:Relationship works.
    let xml = format!(
        r#"<?xml version="1.0"?><r:Relationships xmlns:r="{RELS_NS}"><r:Relationship Id="rId1" Type="t" Target="a.xml"/></r:Relationships>"#
    );
    let rels = parse_relationships(&xml).expect("parse");
    assert_eq!(rels.len(), 1);
    assert_eq!(rels["rId1"].target, "a.xml");
}

#[test]
fn empty_input_is_ok_empty() {
    // Completely empty input is not an error — yields an empty map.
    let rels = parse_relationships("").expect("parse empty");
    assert!(rels.is_empty());
}

#[test]
fn targetmode_external_ignored_but_target_kept() {
    // Extra attributes (TargetMode) are ignored; Target is still captured.
    let xml = wrap(
        r#"<Relationship Id="rId1" Type="t" Target="https://ext/x" TargetMode="External"/>"#,
    );
    let rels = parse_relationships(&xml).expect("parse");
    assert_eq!(rels["rId1"].target, "https://ext/x");
}
