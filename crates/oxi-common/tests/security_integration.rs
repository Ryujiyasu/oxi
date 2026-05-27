//! Integration tests for oxi-common::security public API.
//!
//! The security module is the ZIP-slip / path-traversal / decompression-bomb
//! defense for ALL three OOXML engines (Word/Excel/PowerPoint open untrusted
//! user files through it). A regression here is a security hole, so the public
//! contract is pinned exhaustively — especially the edge cases the in-src unit
//! tests don't cover (drive letters, null bytes, mixed separators, base_dir
//! traversal, boundary sizes).

use oxi_common::security::{
    sanitize_rel_target, validate_input_size, validate_path, validate_zip_entry,
    MAX_COMPRESSION_RATIO, MAX_ENTRY_SIZE, MAX_INPUT_SIZE,
};

// ────────────────────────────────────────────────────────────────────
// validate_path — the core zip-slip defense
// ────────────────────────────────────────────────────────────────────

#[test]
fn validate_path_accepts_normal_ooxml_paths() {
    for p in [
        "word/document.xml",
        "xl/worksheets/sheet1.xml",
        "ppt/slides/slide1.xml",
        "[Content_Types].xml",
        "word/_rels/document.xml.rels",
        "word/media/image1.png",
    ] {
        assert!(validate_path(p).is_ok(), "should accept {p:?}");
    }
}

#[test]
fn validate_path_rejects_empty() {
    assert!(validate_path("").is_err());
}

#[test]
fn validate_path_rejects_absolute_unix_and_windows() {
    assert!(validate_path("/etc/passwd").is_err());
    assert!(validate_path("\\windows\\system32").is_err());
}

#[test]
fn validate_path_rejects_drive_letters() {
    // Windows drive-letter absolute paths must be rejected (zip slip vector).
    assert!(validate_path("C:\\Windows\\system32").is_err());
    assert!(validate_path("D:/data/x.xml").is_err());
    assert!(validate_path("z:relative.xml").is_err());
    // Lowercase drive letter too.
    assert!(validate_path("c:\\x").is_err());
}

#[test]
fn validate_path_non_drive_colon_is_allowed() {
    // A colon NOT in drive-letter position (byte[1] != ':') is not treated as
    // a drive letter. "ab:cd" has ':' at index 2, so it's not rejected by the
    // drive-letter rule (pins the exact byte[1]=='::' check).
    assert!(validate_path("ab:cd.xml").is_ok());
}

#[test]
fn validate_path_rejects_dotdot_traversal() {
    assert!(validate_path("../../../etc/passwd").is_err());
    assert!(validate_path("word/../../secret.xml").is_err());
    assert!(validate_path("..\\..\\windows\\system32").is_err());
    // ".." as the very first component.
    assert!(validate_path("../x").is_err());
    // ".." in the middle with forward slashes.
    assert!(validate_path("a/b/../c").is_err());
}

#[test]
fn validate_path_allows_dotdot_substring_not_component() {
    // "..foo" or "foo.." are NOT traversal (only a bare ".." component is).
    // Pins that the check is component-wise, not substring.
    assert!(validate_path("word/..foo/x.xml").is_ok());
    assert!(validate_path("word/foo../x.xml").is_ok());
    assert!(validate_path("word/a..b/x.xml").is_ok());
}

#[test]
fn validate_path_rejects_null_byte() {
    assert!(validate_path("word/doc\0.xml").is_err());
}

#[test]
fn validate_path_single_dot_component_allowed() {
    // A single "." component is not traversal (stays in place).
    assert!(validate_path("word/./document.xml").is_ok());
}

// ────────────────────────────────────────────────────────────────────
// sanitize_rel_target — relationship target resolution
// ────────────────────────────────────────────────────────────────────

#[test]
fn sanitize_rel_target_relative_prepends_base() {
    assert_eq!(
        sanitize_rel_target("word", "media/image1.png").unwrap(),
        "word/media/image1.png"
    );
}

#[test]
fn sanitize_rel_target_empty_base_makes_path_absolute_and_rejects() {
    // Pin ACTUAL behavior: with empty base_dir, the internal validation path
    // becomes format!("{}/{}", "", target) = "/target", which validate_path
    // rejects as absolute. Callers always pass a non-empty base ("word",
    // "ppt", "xl"), so this edge case is a reject, not a passthrough.
    assert!(sanitize_rel_target("", "document.xml").is_err());
}

#[test]
fn sanitize_rel_target_absolute_strips_leading_slash() {
    // Absolute targets (e.g. "/word/styles.xml") strip the leading slash.
    assert_eq!(
        sanitize_rel_target("ppt", "/word/styles.xml").unwrap(),
        "word/styles.xml"
    );
}

#[test]
fn sanitize_rel_target_rejects_traversal_escaping_base() {
    // A relative target that escapes the base via ".." must be rejected.
    assert!(sanitize_rel_target("word", "../../../etc/passwd").is_err());
    assert!(sanitize_rel_target("word", "../secret.xml").is_err());
}

#[test]
fn sanitize_rel_target_rejects_absolute_traversal() {
    // Absolute target that still contains traversal after stripping slash.
    assert!(sanitize_rel_target("word", "/../../etc/passwd").is_err());
}

// ────────────────────────────────────────────────────────────────────
// validate_zip_entry — decompression-bomb defense
// ────────────────────────────────────────────────────────────────────

#[test]
fn validate_zip_entry_normal_ok() {
    assert!(validate_zip_entry("word/document.xml", 1000, 5000).is_ok());
}

#[test]
fn validate_zip_entry_at_max_size_boundary() {
    // Exactly MAX_ENTRY_SIZE is allowed; +1 is rejected. Use a compressed
    // size large enough that the ratio check (uncompressed/compressed <= 100)
    // does NOT also fire — otherwise the ratio guard masks the size boundary.
    let comp = MAX_ENTRY_SIZE / MAX_COMPRESSION_RATIO; // ratio exactly == limit
    assert!(validate_zip_entry("x", comp, MAX_ENTRY_SIZE).is_ok());
    assert!(validate_zip_entry("x", comp, MAX_ENTRY_SIZE + 1).is_err());
}

#[test]
fn validate_zip_entry_compression_ratio_boundary() {
    // ratio == MAX is allowed (uses strict > comparison); ratio > MAX rejected.
    let comp = 1000u64;
    assert!(validate_zip_entry("x", comp, comp * MAX_COMPRESSION_RATIO).is_ok());
    assert!(validate_zip_entry("x", comp, comp * MAX_COMPRESSION_RATIO + comp + 1).is_err());
}

#[test]
fn validate_zip_entry_zero_compressed_no_divide_by_zero() {
    // compressed_size == 0 must not panic (division guard).
    assert!(validate_zip_entry("empty.xml", 0, 0).is_ok());
    assert!(validate_zip_entry("stored.xml", 0, 100).is_ok());
}

// ────────────────────────────────────────────────────────────────────
// validate_input_size
// ────────────────────────────────────────────────────────────────────

#[test]
fn validate_input_size_small_ok() {
    assert!(validate_input_size(&[0u8; 1024]).is_ok());
    assert!(validate_input_size(&[]).is_ok());
}

#[test]
fn security_constants_are_sane() {
    // Pin the documented limits so an accidental edit (e.g. dropping a zero)
    // surfaces in CI.
    assert_eq!(MAX_INPUT_SIZE, 256 * 1024 * 1024);
    assert_eq!(MAX_ENTRY_SIZE, 128 * 1024 * 1024);
    assert_eq!(MAX_COMPRESSION_RATIO, 100);
}
