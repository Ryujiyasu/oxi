//! PDF signature verification.
//!
//! Extracts signature information from signed PDFs and provides
//! the data needed for verification by external crypto libraries.

use crate::error::PdfError;
use super::types::{SignatureInfo, VerificationResult};

/// Maximum PDF size for signature scanning (512 MB).
const MAX_VERIFY_SIZE: usize = 512 * 1024 * 1024;
/// Maximum number of signatures to extract from a single PDF.
const MAX_SIGNATURES: usize = 256;

/// Extract all digital signatures from a PDF.
pub fn verify_pdf_signatures(pdf_data: &[u8]) -> Result<Vec<SignatureInfo>, PdfError> {
    if pdf_data.len() > MAX_VERIFY_SIZE {
        return Err(PdfError::Parse(format!(
            "PDF too large for signature verification ({} bytes, max {} MB)",
            pdf_data.len(),
            MAX_VERIFY_SIZE / (1024 * 1024)
        )));
    }

    let mut signatures = Vec::new();

    // Scan for signature dictionaries.
    let text = String::from_utf8_lossy(pdf_data);
    let mut search_from = 0;

    while let Some(pos) = text[search_from..].find("/Type /Sig") {
        let abs_pos = search_from + pos;
        // Find the enclosing dictionary.
        if let Some(sig_info) = extract_signature_at(pdf_data, &text, abs_pos) {
            signatures.push(sig_info);
            if signatures.len() >= MAX_SIGNATURES {
                break;
            }
        }
        search_from = abs_pos + 10;
    }

    Ok(signatures)
}

/// Verify a single signature's data integrity.
///
/// This checks whether the signed byte ranges match the PKCS#7 data.
/// Actual cryptographic verification requires an external provider
/// (e.g., OpenSSL, ring, or a platform-specific crypto API).
pub fn check_integrity(
    pdf_data: &[u8],
    sig: &SignatureInfo,
) -> VerificationResult {
    // Extract the signed data from byte ranges.
    let [off1, len1, off2, len2] = sig.byte_range;

    // Guard against integer overflow and out-of-bounds.
    let end1 = off1.checked_add(len1);
    let end2 = off2.checked_add(len2);
    let (end1, end2) = match (end1, end2) {
        (Some(e1), Some(e2)) => (e1, e2),
        _ => {
            return VerificationResult::Invalid {
                reason: "byte range arithmetic overflow".into(),
            };
        }
    };

    if end1 > pdf_data.len() as u64 || end2 > pdf_data.len() as u64 {
        return VerificationResult::Invalid {
            reason: "byte range exceeds file size".into(),
        };
    }

    if off2 < end1 {
        return VerificationResult::Invalid {
            reason: "byte ranges overlap".into(),
        };
    }

    // Check if the signature covers the whole document.
    let total_covered = len1 + len2;
    let gap = off2 - end1; // the signature hex string

    if off1 != 0 || end2 as usize != pdf_data.len() {
        // Signature doesn't cover the whole file — potentially partial.
        return VerificationResult::Unknown {
            reason: "signature does not cover entire document".into(),
        };
    }

    if sig.pkcs7_data.is_empty() {
        return VerificationResult::Invalid {
            reason: "empty PKCS#7 data".into(),
        };
    }

    // At this point, we have the signed data and PKCS#7 blob.
    // Actual verification requires a crypto library to:
    // 1. Parse the PKCS#7 SignedData structure
    // 2. Verify the signature against the certificate
    // 3. Check the certificate chain
    // For now, return Unknown since we don't have crypto deps yet.
    VerificationResult::Unknown {
        reason: format!(
            "crypto verification not yet implemented (signed {} bytes, gap {} bytes, pkcs7 {} bytes)",
            total_covered, gap, sig.pkcs7_data.len()
        ),
    }
}

fn extract_signature_at(
    raw_data: &[u8],
    text: &str,
    sig_pos: usize,
) -> Option<SignatureInfo> {
    // Find the start of the dictionary containing this /Type /Sig.
    let dict_start = text[..sig_pos].rfind("<<")?;
    let dict_end = find_matching_dict_end(text, dict_start)?;
    let dict_text = &text[dict_start..dict_end + 2];

    let name = extract_paren_value(dict_text, "/Name");
    let reason = extract_paren_value(dict_text, "/Reason");
    let location = extract_paren_value(dict_text, "/Location");
    let date = extract_paren_value(dict_text, "/M");

    // Extract ByteRange.
    let byte_range = extract_byte_range(dict_text)?;

    // Extract Contents (hex-encoded PKCS#7).
    let pkcs7_data = extract_hex_contents(dict_text)?;

    let covers_whole = byte_range[0] == 0
        && (byte_range[2] + byte_range[3]) as usize == raw_data.len();

    Some(SignatureInfo {
        name,
        reason,
        location,
        date,
        pkcs7_data,
        byte_range,
        covers_whole_document: covers_whole,
    })
}

fn find_matching_dict_end(text: &str, start: usize) -> Option<usize> {
    let bytes = text.as_bytes();
    let mut depth = 0;
    let mut i = start;
    while i + 1 < bytes.len() {
        if bytes[i] == b'<' && bytes[i + 1] == b'<' {
            depth += 1;
            i += 2;
        } else if bytes[i] == b'>' && bytes[i + 1] == b'>' {
            depth -= 1;
            if depth == 0 {
                return Some(i);
            }
            i += 2;
        } else {
            i += 1;
        }
    }
    None
}

fn extract_paren_value(dict: &str, key: &str) -> Option<String> {
    let pos = dict.find(key)?;
    let after = &dict[pos + key.len()..];
    let paren_start = after.find('(')?;
    let paren_end = after[paren_start + 1..].find(')')?;
    Some(after[paren_start + 1..paren_start + 1 + paren_end].to_string())
}

fn extract_byte_range(dict: &str) -> Option<[u64; 4]> {
    let pos = dict.find("/ByteRange")?;
    let after = &dict[pos + 10..];
    let bracket_start = after.find('[')?;
    let bracket_end = after.find(']')?;
    let inner = after[bracket_start + 1..bracket_end].trim();

    let nums: Vec<u64> = inner
        .split_whitespace()
        .filter_map(|s| s.parse().ok())
        .collect();

    if nums.len() >= 4 {
        Some([nums[0], nums[1], nums[2], nums[3]])
    } else {
        None
    }
}

fn extract_hex_contents(dict: &str) -> Option<Vec<u8>> {
    let pos = dict.find("/Contents")?;
    let after = &dict[pos + 9..];
    let hex_start = after.find('<')?;
    let hex_end = after[hex_start + 1..].find('>')?;
    let hex_str = &after[hex_start + 1..hex_start + 1 + hex_end];

    // Decode hex, skipping zero padding.
    let trimmed = hex_str.trim_end_matches('0');
    if trimmed.is_empty() {
        return Some(Vec::new());
    }

    let bytes: Vec<u8> = (0..trimmed.len())
        .step_by(2)
        .filter_map(|i| {
            let end = (i + 2).min(trimmed.len());
            u8::from_str_radix(&trimmed[i..end], 16).ok()
        })
        .collect();

    Some(bytes)
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::signature::signer::sign_pdf;
    use crate::signature::types::SignatureConfig;

    struct DummyProvider;

    impl crate::signature::types::SignatureProvider for DummyProvider {
        fn sign(&self, _data: &[u8]) -> Result<Vec<u8>, PdfError> {
            Ok(vec![0x30, 0x82, 0x00, 0x04, 0xDE, 0xAD, 0xBE, 0xEF])
        }
        fn certificate(&self) -> Result<Vec<u8>, PdfError> {
            Ok(vec![])
        }
        fn signer_name(&self) -> String {
            "Verifier Test".into()
        }
    }

    #[test]
    fn test_extract_signatures() {
        let doc = crate::ir::PdfDocument {
            version: crate::ir::PdfVersion::new(1, 7),
            info: crate::ir::DocumentInfo::default(),
            pages: vec![crate::ir::Page {
                width: 612.0,
                height: 792.0,
                media_box: crate::ir::Rectangle {
                    llx: 0.0, lly: 0.0, urx: 612.0, ury: 792.0,
                },
                crop_box: None,
                contents: vec![],
                rotation: 0,
            }],
            outline: Vec::new(),
            embedded_fonts: std::collections::HashMap::new(),
        };
        let pdf = crate::write_pdf(&doc);
        let signed = sign_pdf(&pdf, &DummyProvider, &SignatureConfig::default()).unwrap();

        let sigs = verify_pdf_signatures(&signed).unwrap();
        assert_eq!(sigs.len(), 1);
        assert_eq!(sigs[0].name, Some("Verifier Test".into()));
        assert!(!sigs[0].pkcs7_data.is_empty());
    }

    #[test]
    fn test_check_integrity_unknown() {
        let doc = crate::ir::PdfDocument {
            version: crate::ir::PdfVersion::new(1, 7),
            info: crate::ir::DocumentInfo::default(),
            pages: vec![crate::ir::Page {
                width: 612.0,
                height: 792.0,
                media_box: crate::ir::Rectangle {
                    llx: 0.0, lly: 0.0, urx: 612.0, ury: 792.0,
                },
                crop_box: None,
                contents: vec![],
                rotation: 0,
            }],
            outline: Vec::new(),
            embedded_fonts: std::collections::HashMap::new(),
        };
        let pdf = crate::write_pdf(&doc);
        let signed = sign_pdf(&pdf, &DummyProvider, &SignatureConfig::default()).unwrap();

        let sigs = verify_pdf_signatures(&signed).unwrap();
        let result = check_integrity(&signed, &sigs[0]);

        // Should return Unknown since we don't have crypto verification yet.
        match result {
            VerificationResult::Unknown { .. } => {} // expected
            other => panic!("expected Unknown, got {:?}", other),
        }
    }

    #[test]
    fn test_no_signatures() {
        let doc = crate::ir::PdfDocument {
            version: crate::ir::PdfVersion::new(1, 7),
            info: crate::ir::DocumentInfo::default(),
            pages: vec![crate::ir::Page {
                width: 612.0,
                height: 792.0,
                media_box: crate::ir::Rectangle {
                    llx: 0.0, lly: 0.0, urx: 612.0, ury: 792.0,
                },
                crop_box: None,
                contents: vec![],
                rotation: 0,
            }],
            outline: Vec::new(),
            embedded_fonts: std::collections::HashMap::new(),
        };
        let pdf = crate::write_pdf(&doc);
        let sigs = verify_pdf_signatures(&pdf).unwrap();
        assert!(sigs.is_empty());
    }
}
