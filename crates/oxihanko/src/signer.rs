//! Hanko-based PDF signer.
//!
//! Combines a stamp image with a digital signature to produce
//! a signed PDF with a visible Japanese hanko stamp.

use oxipdf_core::error::PdfError;
use oxipdf_core::signature::types::{
    SignatureAppearance, SignatureConfig, SignatureProvider,
};
use oxipdf_core::signature::signer::sign_pdf;

use crate::stamp::{generate_stamp_svg, StampConfig};

/// Configuration for hanko-stamped PDF signing.
#[derive(Debug, Clone)]
pub struct HankoSignConfig {
    /// Stamp appearance configuration.
    pub stamp: StampConfig,
    /// Page to place the stamp on (0-based).
    pub page: usize,
    /// X position of the stamp (in points from left).
    pub x: f64,
    /// Y position of the stamp (in points from top).
    pub y: f64,
    /// Size of the stamp on the page (in points).
    pub stamp_size: f64,
    /// Signature reason.
    pub reason: Option<String>,
    /// Signing location.
    pub location: Option<String>,
}

impl Default for HankoSignConfig {
    fn default() -> Self {
        Self {
            stamp: StampConfig::default(),
            page: 0,
            x: 450.0,   // Right side of A4
            y: 100.0,   // Near top
            stamp_size: 60.0,
            reason: None,
            location: None,
        }
    }
}

/// Sign a PDF with a hanko stamp.
///
/// This function:
/// 1. Generates a hanko stamp SVG
/// 2. Creates a visible signature appearance with the stamp
/// 3. Signs the PDF using the provided signature provider
pub fn sign_pdf_with_hanko(
    pdf_data: &[u8],
    provider: &dyn SignatureProvider,
    config: &HankoSignConfig,
) -> Result<Vec<u8>, PdfError> {
    let stamp_svg = generate_stamp_svg(&config.stamp);
    let stamp_bytes = stamp_svg.into_bytes();

    let sig_config = SignatureConfig {
        reason: config.reason.clone(),
        location: config.location.clone(),
        contact_info: None,
        appearance: Some(SignatureAppearance {
            page: config.page,
            x: config.x,
            y: config.y,
            width: config.stamp_size,
            height: config.stamp_size,
            stamp_image: Some(stamp_bytes),
        }),
    };

    sign_pdf(pdf_data, provider, &sig_config)
}

/// Generate a hanko stamp SVG for preview (without signing).
pub fn preview_stamp(config: &StampConfig) -> String {
    generate_stamp_svg(config)
}

#[cfg(test)]
mod tests {
    use super::*;

    struct DummyProvider;

    impl SignatureProvider for DummyProvider {
        fn sign(&self, _data: &[u8]) -> Result<Vec<u8>, PdfError> {
            Ok(vec![0x30, 0x82, 0xDE, 0xAD])
        }
        fn certificate(&self) -> Result<Vec<u8>, PdfError> {
            Ok(vec![])
        }
        fn signer_name(&self) -> String {
            "山田太郎".into()
        }
    }

    #[test]
    fn test_sign_pdf_with_hanko() {
        let doc = oxipdf_core::ir::PdfDocument {
            version: oxipdf_core::ir::PdfVersion::new(1, 7),
            info: oxipdf_core::ir::DocumentInfo::default(),
            pages: vec![oxipdf_core::ir::Page {
                width: 595.0,
                height: 842.0,
                media_box: oxipdf_core::ir::Rectangle {
                    llx: 0.0, lly: 0.0, urx: 595.0, ury: 842.0,
                },
                crop_box: None,
                contents: vec![],
                rotation: 0,
            }],
            outline: Vec::new(),
        };
        let pdf = oxipdf_core::write_pdf(&doc);

        let config = HankoSignConfig {
            stamp: StampConfig {
                name: "山田".into(),
                ..StampConfig::default()
            },
            reason: Some("承認".into()),
            location: Some("東京".into()),
            ..Default::default()
        };

        let signed = sign_pdf_with_hanko(&pdf, &DummyProvider, &config).unwrap();
        assert!(signed.len() > pdf.len());

        let text = String::from_utf8_lossy(&signed);
        assert!(text.contains("/Type /Sig"));
        assert!(text.contains("山田太郎")); // signer name
    }

    #[test]
    fn test_preview_stamp() {
        let config = StampConfig {
            name: "鈴木".into(),
            ..StampConfig::default()
        };
        let svg = preview_stamp(&config);
        assert!(svg.contains("鈴木"));
        assert!(svg.contains("<svg"));
    }
}
