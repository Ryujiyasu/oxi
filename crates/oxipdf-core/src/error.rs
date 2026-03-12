use thiserror::Error;

#[derive(Error, Debug)]
pub enum PdfError {
    #[error("parse error: {0}")]
    Parse(String),

    #[error("unsupported feature: {0}")]
    Unsupported(String),

    #[error("I/O error: {0}")]
    Io(#[from] std::io::Error),
}
