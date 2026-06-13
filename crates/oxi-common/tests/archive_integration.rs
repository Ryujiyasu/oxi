// This Source Code Form is subject to the terms of the Mozilla Public
// License, v. 2.0. If a copy of the MPL was not distributed with this
// file, You can obtain one at https://mozilla.org/MPL/2.0/.

//! Integration tests for oxi-common::archive (OoxmlArchive).
//!
//! OoxmlArchive is the ZIP-reading layer beneath ALL three OOXML engines.
//! It wraps the `zip` crate with the security limits (entry count, size,
//! compression ratio). These tests pin the public read API + error contract
//! using a real OOXML fixture (basic_test.xlsx is a valid ZIP package).

use oxi_common::archive::OoxmlArchive;
use oxi_common::OxiError;

// A real OOXML ZIP package (xlsx) committed as a test fixture.
const XLSX: &[u8] = include_bytes!("../../../tests/fixtures/basic_test.xlsx");

#[test]
fn new_opens_valid_ooxml_zip() {
    let arch = OoxmlArchive::new(XLSX);
    assert!(arch.is_ok(), "valid xlsx ZIP must open");
}

#[test]
fn new_rejects_garbage() {
    // Non-ZIP bytes must error (not panic).
    assert!(OoxmlArchive::new(b"not a zip file at all").is_err());
}

#[test]
fn new_rejects_empty() {
    assert!(OoxmlArchive::new(b"").is_err());
}

#[test]
fn file_names_lists_entries() {
    let arch = OoxmlArchive::new(XLSX).expect("open");
    let names = arch.file_names();
    assert!(!names.is_empty(), "archive must list entries");
    // Every OOXML package contains the content-types part.
    assert!(
        names.iter().any(|n| n == "[Content_Types].xml"),
        "OOXML package must contain [Content_Types].xml, got: {names:?}"
    );
}

#[test]
fn read_part_reads_content_types() {
    let mut arch = OoxmlArchive::new(XLSX).expect("open");
    let ct = arch.read_part("[Content_Types].xml").expect("read content types");
    assert!(ct.contains("<Types"), "content-types XML must contain <Types");
}

#[test]
fn read_part_missing_returns_missing_part_error() {
    let mut arch = OoxmlArchive::new(XLSX).expect("open");
    let err = match arch.read_part("does/not/exist.xml") {
        Err(e) => e,
        Ok(_) => panic!("expected MissingPart error"),
    };
    assert!(matches!(err, OxiError::MissingPart(ref n) if n == "does/not/exist.xml"));
}

#[test]
fn try_read_part_missing_returns_none() {
    let mut arch = OoxmlArchive::new(XLSX).expect("open");
    // Present part → Some
    assert!(arch.try_read_part("[Content_Types].xml").expect("ok").is_some());
    // Absent part → None (NOT an error)
    assert_eq!(arch.try_read_part("ppt/slides/slide999.xml").expect("ok"), None);
}

#[test]
fn read_binary_part_reads_bytes() {
    let mut arch = OoxmlArchive::new(XLSX).expect("open");
    let bytes = arch
        .read_binary_part("[Content_Types].xml")
        .expect("read binary");
    assert!(!bytes.is_empty());
    // Binary read of the same part should match the text read's byte length
    // (content-types is ASCII/UTF-8).
    let text = OoxmlArchive::new(XLSX)
        .unwrap()
        .read_part("[Content_Types].xml")
        .unwrap();
    assert_eq!(bytes.len(), text.len());
}

#[test]
fn read_binary_part_missing_errors() {
    let mut arch = OoxmlArchive::new(XLSX).expect("open");
    assert!(arch.read_binary_part("nope.bin").is_err());
}

#[test]
fn read_part_repeatable() {
    // Reading the same part twice must yield identical content (no consumed
    // stream surprises across calls).
    let mut arch = OoxmlArchive::new(XLSX).expect("open");
    let a = arch.read_part("[Content_Types].xml").expect("first");
    let b = arch.read_part("[Content_Types].xml").expect("second");
    assert_eq!(a, b);
}
