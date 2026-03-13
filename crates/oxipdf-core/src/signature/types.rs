//! Signature types for PDF digital signatures (PAdES / PKCS#7).

use crate::error::PdfError;
use serde::Serialize;

/// Trait for pluggable signature providers.
///
/// Implement this trait to support different signing backends:
/// - Software certificates (PKCS#12 / PFX files)
/// - Hardware tokens (PKCS#11 / smart cards)
/// - Cloud signing services
/// - oxihanko (digital stamp + signature)
pub trait SignatureProvider {
    /// Return the DER-encoded PKCS#7 signature for the given data.
    fn sign(&self, data: &[u8]) -> Result<Vec<u8>, PdfError>;

    /// Return the DER-encoded signing certificate.
    fn certificate(&self) -> Result<Vec<u8>, PdfError>;

    /// Return the signer's common name (for display in /Name field).
    fn signer_name(&self) -> String;
}

/// Configuration for signing a PDF.
#[derive(Debug, Clone)]
pub struct SignatureConfig {
    /// Reason for signing (e.g. "承認", "Approved").
    pub reason: Option<String>,
    /// Location of signing (e.g. "Tokyo").
    pub location: Option<String>,
    /// Contact info for the signer.
    pub contact_info: Option<String>,
    /// Visible signature appearance rectangle (page coords in points).
    /// If None, the signature is invisible.
    pub appearance: Option<SignatureAppearance>,
}

/// Visible signature appearance on the page.
#[derive(Debug, Clone)]
pub struct SignatureAppearance {
    /// Page index (0-based) where the signature appears.
    pub page: usize,
    /// Rectangle for the signature widget (x, y, width, height in points).
    pub x: f64,
    pub y: f64,
    pub width: f64,
    pub height: f64,
    /// Optional stamp/seal image (PNG bytes). Used by oxihanko.
    pub stamp_image: Option<Vec<u8>>,
}

impl Default for SignatureConfig {
    fn default() -> Self {
        Self {
            reason: None,
            location: None,
            contact_info: None,
            appearance: None,
        }
    }
}

/// Information about a signature found in a PDF.
#[derive(Debug, Clone, Serialize)]
pub struct SignatureInfo {
    /// Signer's name from the /Name field.
    pub name: Option<String>,
    /// Reason for signing.
    pub reason: Option<String>,
    /// Location of signing.
    pub location: Option<String>,
    /// Signing date (as PDF date string).
    pub date: Option<String>,
    /// The raw PKCS#7 signature bytes.
    pub pkcs7_data: Vec<u8>,
    /// Byte range that was signed [offset1, len1, offset2, len2].
    pub byte_range: [u64; 4],
    /// Whether the signature covers the entire document.
    pub covers_whole_document: bool,
}

/// Result of signature verification.
#[derive(Debug, Clone, Serialize)]
pub enum VerificationResult {
    /// Signature is valid — data integrity confirmed.
    Valid {
        signer_name: String,
    },
    /// Signature is invalid — data has been modified.
    Invalid {
        reason: String,
    },
    /// Cannot verify — missing certificate or unsupported algorithm.
    Unknown {
        reason: String,
    },
}
