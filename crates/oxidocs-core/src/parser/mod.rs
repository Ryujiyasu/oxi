mod numbering;
mod ooxml;
mod relationships;
mod styles;
pub mod theme;

use crate::ir::Document;
use thiserror::Error;

pub use ooxml::OoxmlParser;

#[derive(Error, Debug)]
pub enum ParseError {
    #[error("failed to read zip archive: {0}")]
    ZipError(#[from] zip::result::ZipError),
    #[error("failed to parse XML: {0}")]
    XmlError(#[from] quick_xml::Error),
    #[error("missing required part: {0}")]
    MissingPart(String),
    #[error("invalid attribute: {0}")]
    InvalidAttribute(String),
    #[error("I/O error: {0}")]
    IoError(#[from] std::io::Error),
}

pub fn parse_docx(data: &[u8]) -> Result<Document, ParseError> {
    let parser = OoxmlParser::new(data)?;
    parser.parse()
}
