//! PDF signing: embed a PAdES-compatible digital signature into a PDF.
//!
//! The signing process follows the PDF incremental update model:
//! 1. Append a new signature dictionary with a /ByteRange placeholder
//! 2. Calculate the byte ranges (everything except the signature hex string)
//! 3. Hash the byte ranges
//! 4. Call the SignatureProvider to produce a PKCS#7 signature
//! 5. Inject the signature into the reserved space

use crate::error::PdfError;
use super::types::{SignatureConfig, SignatureProvider};

/// The size reserved for the PKCS#7 signature (in bytes).
/// 8192 bytes = 16384 hex characters. Enough for most certificates.
const SIGNATURE_RESERVED_BYTES: usize = 8192;

/// Sign a PDF document, returning the signed PDF bytes.
///
/// This performs an incremental update — the original PDF is preserved
/// and the signature is appended.
pub fn sign_pdf(
    pdf_data: &[u8],
    provider: &dyn SignatureProvider,
    config: &SignatureConfig,
) -> Result<Vec<u8>, PdfError> {
    let signer_name = provider.signer_name();

    // Build the signature dictionary object.
    let sig_dict = build_signature_dict(&signer_name, config);

    // Calculate where to insert the ByteRange and Contents.
    // For now, we append an incremental update to the existing PDF.
    let mut output = pdf_data.to_vec();

    // Ensure the original PDF ends properly.
    if !output.ends_with(b"\n") {
        output.push(b'\n');
    }

    // Append the signature object.
    let sig_obj_offset = output.len();
    let sig_obj_num = find_next_obj_num(pdf_data)?;

    write_bytes(
        &mut output,
        format!("{sig_obj_num} 0 obj\n{sig_dict}\nendobj\n").as_bytes(),
    );

    // Find the /Contents placeholder position within the output.
    let contents_marker = b"/Contents <";
    let contents_pos = find_last_occurrence(&output, contents_marker)
        .ok_or_else(|| PdfError::Parse("failed to find /Contents placeholder".into()))?;
    let hex_start = contents_pos + contents_marker.len();
    let hex_end = hex_start + SIGNATURE_RESERVED_BYTES * 2;

    // Calculate ByteRange: [0, hex_start, hex_end+1, total_len - hex_end - 1]
    // We need to know the final file length first, so we build the xref + trailer now.
    let xref_offset = output.len();

    // Minimal xref for the incremental update.
    write_bytes(
        &mut output,
        format!(
            "xref\n{sig_obj_num} 1\n"
        )
        .as_bytes(),
    );
    let entry = format!("{:010} 00000 n\r\n", sig_obj_offset);
    write_bytes(&mut output, entry.as_bytes());

    // Find the original trailer's /Size to compute new size.
    let new_size = sig_obj_num + 1;

    // Find the original startxref.
    let prev_xref = crate::parser::xref::find_startxref(pdf_data)?;

    write_bytes(
        &mut output,
        format!(
            "trailer\n<< /Size {new_size} /Root 1 0 R /Prev {prev_xref} >>\n\
             startxref\n{xref_offset}\n%%EOF\n"
        )
        .as_bytes(),
    );

    // Now compute the byte range.
    let total_len = output.len() as u64;
    let byte_range = [
        0u64,
        hex_start as u64,
        (hex_end + 1) as u64,
        total_len - (hex_end + 1) as u64,
    ];

    // Update the ByteRange in the signature dictionary.
    update_byte_range(&mut output, &byte_range)?;

    // Extract the bytes to sign (everything except the hex signature).
    let mut data_to_sign = Vec::new();
    data_to_sign.extend_from_slice(&output[..hex_start]);
    data_to_sign.extend_from_slice(&output[hex_end + 1..]);

    // Call the provider to sign.
    let signature = provider.sign(&data_to_sign)?;

    // Encode the signature as hex and inject it.
    let hex_sig = encode_hex(&signature);
    if hex_sig.len() > SIGNATURE_RESERVED_BYTES * 2 {
        return Err(PdfError::Parse(
            "signature too large for reserved space".into(),
        ));
    }

    // Pad with zeros.
    let mut padded = hex_sig;
    while padded.len() < SIGNATURE_RESERVED_BYTES * 2 {
        padded.push('0');
    }

    // Write the hex signature into the reserved space.
    output[hex_start..hex_end].copy_from_slice(padded.as_bytes());

    Ok(output)
}

fn build_signature_dict(signer_name: &str, config: &SignatureConfig) -> String {
    let hex_placeholder = "0".repeat(SIGNATURE_RESERVED_BYTES * 2);
    // ByteRange placeholder — will be updated after we know the file layout.
    let byte_range_placeholder = format!("{:<40}", "[0 0 0 0]");

    let mut dict = format!(
        "<< /Type /Sig /Filter /Adobe.PPKLite /SubFilter /adbe.pkcs7.detached \
         /ByteRange {byte_range_placeholder} \
         /Contents <{hex_placeholder}> \
         /Name ({signer_name})"
    );

    if let Some(ref reason) = config.reason {
        dict.push_str(&format!(" /Reason ({reason})"));
    }
    if let Some(ref location) = config.location {
        dict.push_str(&format!(" /Location ({location})"));
    }
    if let Some(ref contact) = config.contact_info {
        dict.push_str(&format!(" /ContactInfo ({contact})"));
    }

    dict.push_str(" >>");
    dict
}

fn find_next_obj_num(pdf_data: &[u8]) -> Result<u32, PdfError> {
    // Find the /Size in the trailer to determine the next available object number.
    find_trailer_size(pdf_data)
        .ok_or_else(|| PdfError::Parse("cannot determine next object number".into()))
}

fn find_trailer_size(data: &[u8]) -> Option<u32> {
    let text = String::from_utf8_lossy(data);
    let size_pos = text.rfind("/Size ")?;
    let after = &text[size_pos + 6..];
    let num_str: String = after
        .chars()
        .take_while(|c| c.is_ascii_digit())
        .collect();
    num_str.parse().ok()
}

fn update_byte_range(output: &mut [u8], byte_range: &[u64; 4]) -> Result<(), PdfError> {
    let marker = b"/ByteRange ";
    let pos = find_last_occurrence(output, marker)
        .ok_or_else(|| PdfError::Parse("ByteRange marker not found".into()))?;

    let range_str = format!(
        "[{} {} {} {}]",
        byte_range[0], byte_range[1], byte_range[2], byte_range[3]
    );
    // Pad to fill the placeholder space (40 chars).
    let padded = format!("{:<40}", range_str);

    let start = pos + marker.len();
    if start + padded.len() > output.len() {
        return Err(PdfError::Parse("ByteRange overflow".into()));
    }
    output[start..start + padded.len()].copy_from_slice(padded.as_bytes());
    Ok(())
}

fn find_last_occurrence(data: &[u8], needle: &[u8]) -> Option<usize> {
    data.windows(needle.len())
        .rposition(|w| w == needle)
}

fn write_bytes(buf: &mut Vec<u8>, data: &[u8]) {
    buf.extend_from_slice(data);
}

fn encode_hex(data: &[u8]) -> String {
    data.iter().map(|b| format!("{:02X}", b)).collect()
}

#[cfg(test)]
mod tests {
    use super::*;

    /// A dummy signature provider for testing.
    struct DummyProvider;

    impl SignatureProvider for DummyProvider {
        fn sign(&self, _data: &[u8]) -> Result<Vec<u8>, PdfError> {
            // Return a dummy signature.
            Ok(vec![0x30, 0x82, 0x00, 0x04, 0xDE, 0xAD, 0xBE, 0xEF])
        }

        fn certificate(&self) -> Result<Vec<u8>, PdfError> {
            Ok(vec![0x30, 0x82, 0x00, 0x02, 0xCA, 0xFE])
        }

        fn signer_name(&self) -> String {
            "Test Signer".into()
        }
    }

    #[test]
    fn test_sign_pdf() {
        // Create a simple PDF to sign.
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
        };
        let pdf_bytes = crate::write_pdf(&doc);

        let provider = DummyProvider;
        let config = SignatureConfig {
            reason: Some("テスト署名".into()),
            location: Some("Tokyo".into()),
            ..Default::default()
        };

        let signed = sign_pdf(&pdf_bytes, &provider, &config).unwrap();

        // The signed PDF should be larger than the original.
        assert!(signed.len() > pdf_bytes.len());

        // Should contain PAdES signature markers.
        let text = String::from_utf8_lossy(&signed);
        assert!(text.contains("/Type /Sig"));
        assert!(text.contains("/Filter /Adobe.PPKLite"));
        assert!(text.contains("/SubFilter /adbe.pkcs7.detached"));
        assert!(text.contains("/Name (Test Signer)"));
        assert!(text.contains("/Reason (テスト署名)"));
        assert!(text.contains("/Location (Tokyo)"));

        // ByteRange should have been updated from the placeholder.
        assert!(!text.contains("[0 0 0 0]"));

        // The signature hex should contain our dummy bytes.
        assert!(text.contains("DEADBEEF"));
    }

    #[test]
    fn test_encode_hex() {
        assert_eq!(encode_hex(&[0xDE, 0xAD]), "DEAD");
        assert_eq!(encode_hex(&[0x00, 0xFF]), "00FF");
    }

    #[test]
    fn test_build_signature_dict() {
        let config = SignatureConfig {
            reason: Some("Approved".into()),
            location: Some("Tokyo".into()),
            ..Default::default()
        };
        let dict = build_signature_dict("John", &config);
        assert!(dict.contains("/Type /Sig"));
        assert!(dict.contains("/Name (John)"));
        assert!(dict.contains("/Reason (Approved)"));
        assert!(dict.contains("/Location (Tokyo)"));
    }
}
