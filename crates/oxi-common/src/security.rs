//! Security limits and validation for OOXML document processing.
//!
//! Defends against:
//! - ZIP bombs (decompression bombs)
//! - Path traversal (zip slip)
//! - Excessive resource consumption

use crate::OxiError;

/// Maximum input file size: 256 MB
pub const MAX_INPUT_SIZE: usize = 256 * 1024 * 1024;

/// Maximum decompressed size for a single ZIP entry: 128 MB
pub const MAX_ENTRY_SIZE: u64 = 128 * 1024 * 1024;

/// Maximum total decompressed size across all entries: 512 MB
pub const MAX_TOTAL_DECOMPRESSED: u64 = 512 * 1024 * 1024;

/// Maximum number of entries in a ZIP archive
pub const MAX_ZIP_ENTRIES: usize = 10_000;

/// Maximum decompression ratio (compressed vs decompressed) to detect zip bombs
pub const MAX_COMPRESSION_RATIO: u64 = 100;

/// Validate input data size before parsing.
pub fn validate_input_size(data: &[u8]) -> Result<(), OxiError> {
    if data.len() > MAX_INPUT_SIZE {
        return Err(OxiError::Security(format!(
            "Input file too large: {} bytes (max {} bytes)",
            data.len(),
            MAX_INPUT_SIZE
        )));
    }
    Ok(())
}

/// Validate a ZIP entry before reading it.
/// Checks for excessive decompressed size and suspicious compression ratios.
pub fn validate_zip_entry(
    name: &str,
    compressed_size: u64,
    uncompressed_size: u64,
) -> Result<(), OxiError> {
    // Check absolute size limit
    if uncompressed_size > MAX_ENTRY_SIZE {
        return Err(OxiError::Security(format!(
            "ZIP entry '{}' too large: {} bytes (max {} bytes)",
            name, uncompressed_size, MAX_ENTRY_SIZE
        )));
    }

    // Check compression ratio (zip bomb detection)
    if compressed_size > 0 && uncompressed_size / compressed_size > MAX_COMPRESSION_RATIO {
        return Err(OxiError::Security(format!(
            "ZIP entry '{}' has suspicious compression ratio: {}x (max {}x)",
            name,
            uncompressed_size / compressed_size,
            MAX_COMPRESSION_RATIO
        )));
    }

    Ok(())
}

/// Validate that a path from a ZIP entry or relationship target is safe.
/// Rejects path traversal attempts (zip slip).
pub fn validate_path(path: &str) -> Result<(), OxiError> {
    // Reject empty paths
    if path.is_empty() {
        return Err(OxiError::Security("Empty path".to_string()));
    }

    // Reject absolute paths (Unix and Windows)
    if path.starts_with('/') || path.starts_with('\\') {
        return Err(OxiError::Security(format!(
            "Absolute path rejected: '{}'",
            path
        )));
    }

    // Check for Windows drive letters (e.g., "C:\...")
    if path.len() >= 2 && path.as_bytes()[1] == b':' && path.as_bytes()[0].is_ascii_alphabetic() {
        return Err(OxiError::Security(format!(
            "Absolute path with drive letter rejected: '{}'",
            path
        )));
    }

    // Reject path traversal sequences
    for component in path.split(&['/', '\\']) {
        if component == ".." {
            return Err(OxiError::Security(format!(
                "Path traversal rejected: '{}'",
                path
            )));
        }
    }

    // Reject paths with null bytes
    if path.contains('\0') {
        return Err(OxiError::Security(format!(
            "Null byte in path rejected: '{}'",
            path
        )));
    }

    Ok(())
}

/// Sanitize a relationship target path.
/// Returns a safe resolved path relative to the base directory, or an error.
pub fn sanitize_rel_target(base_dir: &str, target: &str) -> Result<String, OxiError> {
    // Absolute targets: strip leading slash and use as-is
    if target.starts_with('/') {
        let stripped = target.trim_start_matches('/');
        validate_path(stripped)?;
        return Ok(stripped.to_string());
    }

    // Validate the target itself
    validate_path(&format!("{}/{}", base_dir, target))?;

    // Simple resolution: prepend base_dir
    if base_dir.is_empty() {
        Ok(target.to_string())
    } else {
        Ok(format!("{}/{}", base_dir, target))
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_validate_input_size() {
        assert!(validate_input_size(&[0u8; 100]).is_ok());
        // We won't allocate MAX_INPUT_SIZE+1 bytes in tests, just check the logic
    }

    #[test]
    fn test_validate_zip_entry() {
        // Normal entry
        assert!(validate_zip_entry("doc.xml", 1000, 5000).is_ok());

        // Too large
        assert!(validate_zip_entry("huge.xml", 100, MAX_ENTRY_SIZE + 1).is_err());

        // Suspicious ratio
        assert!(validate_zip_entry("bomb.xml", 100, 100 * MAX_COMPRESSION_RATIO + 100).is_err());

        // Zero compressed size (avoid division by zero)
        assert!(validate_zip_entry("empty.xml", 0, 0).is_ok());
    }

    #[test]
    fn test_validate_path_safe() {
        assert!(validate_path("word/document.xml").is_ok());
        assert!(validate_path("xl/worksheets/sheet1.xml").is_ok());
        assert!(validate_path("ppt/slides/slide1.xml").is_ok());
    }

    #[test]
    fn test_validate_path_traversal() {
        assert!(validate_path("../../../etc/passwd").is_err());
        assert!(validate_path("word/../../secret.xml").is_err());
        assert!(validate_path("..\\..\\windows\\system32").is_err());
    }

    #[test]
    fn test_validate_path_absolute() {
        assert!(validate_path("/etc/passwd").is_err());
        assert!(validate_path("\\windows\\system32").is_err());
        assert!(validate_path("C:\\windows").is_err());
    }

    #[test]
    fn test_validate_path_null_byte() {
        assert!(validate_path("word/doc\0.xml").is_err());
    }

    #[test]
    fn test_validate_path_empty() {
        assert!(validate_path("").is_err());
    }

    #[test]
    fn test_sanitize_rel_target() {
        assert_eq!(
            sanitize_rel_target("word", "media/image1.png").unwrap(),
            "word/media/image1.png"
        );
        assert_eq!(
            sanitize_rel_target("ppt", "/ppt/slides/slide1.xml").unwrap(),
            "ppt/slides/slide1.xml"
        );
        assert!(sanitize_rel_target("word", "../../etc/passwd").is_err());
    }
}
