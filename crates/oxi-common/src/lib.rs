pub mod archive;
pub mod relationships;
pub mod security;
pub mod xml_utils;

use thiserror::Error;

#[derive(Error, Debug)]
pub enum OxiError {
    #[error("ZIP error: {0}")]
    Zip(#[from] zip::result::ZipError),

    #[error("XML error: {0}")]
    Xml(#[from] quick_xml::Error),

    #[error("Missing part: {0}")]
    MissingPart(String),

    #[error("Invalid attribute: {0}")]
    InvalidAttribute(String),

    #[error("IO error: {0}")]
    Io(#[from] std::io::Error),

    #[error("Parse error: {0}")]
    Parse(String),

    #[error("Security error: {0}")]
    Security(String),
}
