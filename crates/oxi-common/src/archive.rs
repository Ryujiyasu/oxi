use std::io::{Cursor, Read};
use zip::ZipArchive;

use crate::OxiError;

/// Generic OOXML archive wrapper for reading parts from ZIP files
pub struct OoxmlArchive {
    archive: ZipArchive<Cursor<Vec<u8>>>,
}

impl OoxmlArchive {
    pub fn new(data: &[u8]) -> Result<Self, OxiError> {
        let cursor = Cursor::new(data.to_vec());
        let archive = ZipArchive::new(cursor)?;
        Ok(Self { archive })
    }

    /// Read a text part from the archive
    pub fn read_part(&mut self, name: &str) -> Result<String, OxiError> {
        let mut file = self
            .archive
            .by_name(name)
            .map_err(|_| OxiError::MissingPart(name.to_string()))?;
        let mut contents = String::new();
        file.read_to_string(&mut contents)?;
        Ok(contents)
    }

    /// Read a binary part from the archive
    pub fn read_binary_part(&mut self, name: &str) -> Result<Vec<u8>, OxiError> {
        let mut file = self
            .archive
            .by_name(name)
            .map_err(|_| OxiError::MissingPart(name.to_string()))?;
        let mut contents = Vec::new();
        file.read_to_end(&mut contents)?;
        Ok(contents)
    }

    /// Try to read a text part, returning None if not found
    pub fn try_read_part(&mut self, name: &str) -> Result<Option<String>, OxiError> {
        match self.read_part(name) {
            Ok(s) => Ok(Some(s)),
            Err(OxiError::MissingPart(_)) => Ok(None),
            Err(e) => Err(e),
        }
    }

    /// List all file names in the archive
    pub fn file_names(&self) -> Vec<String> {
        (0..self.archive.len())
            .filter_map(|i| self.archive.name_for_index(i).map(|s| s.to_string()))
            .collect()
    }
}
