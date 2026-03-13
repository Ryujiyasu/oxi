pub mod error;
pub mod extract;
pub mod ir;
pub mod parser;
pub mod signature;
pub mod writer;

pub use error::PdfError;
pub use extract::{extract_text, extract_text_string};
pub use ir::PdfDocument;
pub use parser::parse_pdf;
pub use signature::{sign_pdf, verify_pdf_signatures, SignatureConfig, SignatureProvider};
pub use writer::write_pdf;
