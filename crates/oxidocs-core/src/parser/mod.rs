use crate::ir::{Document, DocumentMetadata, Page, PageSize, Margin, StyleSheet};
use thiserror::Error;

#[derive(Error, Debug)]
pub enum ParseError {
    #[error("failed to read docx file: {0}")]
    ReadError(String),
    #[error("invalid document structure: {0}")]
    InvalidStructure(String),
}

pub fn parse_docx(data: &[u8]) -> Result<Document, ParseError> {
    let _docx = docx_rs::read_docx(data)
        .map_err(|e| ParseError::ReadError(e.to_string()))?;

    // TODO: Convert docx-rs document to our IR
    Ok(Document {
        pages: vec![Page {
            blocks: vec![],
            size: PageSize::default(),
            margin: Margin::default(),
        }],
        styles: StyleSheet::default(),
        metadata: DocumentMetadata::default(),
    })
}
