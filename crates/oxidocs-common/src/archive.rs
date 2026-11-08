// This Source Code Form is subject to the terms of the Mozilla Public
// License, v. 2.0. If a copy of the MPL was not distributed with this
// file, You can obtain one at https://mozilla.org/MPL/2.0/.

use std::io::{Cursor, Read};
use zip::ZipArchive;

use crate::security;
use crate::OxiError;

/// Opens a ZIP package, first repairing one damage Word and LibreOffice
/// tolerate: an end-of-central-directory record whose comment length runs
/// past the end of the file (seen on packages whose tail went through a
/// CRLF-rewriting transfer). The length is clamped to the bytes actually
/// present; everything else must still be a valid archive.
pub fn open_zip(mut data: Vec<u8>) -> zip::result::ZipResult<ZipArchive<Cursor<Vec<u8>>>> {
    clamp_eocd_comment(&mut data);
    ZipArchive::new(Cursor::new(data))
}

/// Clamps an overlong EOCD comment length; true when it changed the bytes.
/// Only the last EOCD signature is considered, and only when no record in
/// the tail is already consistent.
fn clamp_eocd_comment(data: &mut [u8]) -> bool {
    const SIG: &[u8; 4] = b"PK\x05\x06";
    const FIXED: usize = 22;
    let start = data.len().saturating_sub(FIXED + u16::MAX as usize);
    let mut found = None;
    for at in (start..=data.len().saturating_sub(FIXED)).rev() {
        if &data[at..at + 4] != SIG {
            continue;
        }
        let declared = u16::from_le_bytes([data[at + 20], data[at + 21]]) as usize;
        let present = data.len() - at - FIXED;
        if declared <= present {
            return false;
        }
        found.get_or_insert((at, present));
    }
    let Some((at, present)) = found else { return false };
    data[at + 20..at + 22].copy_from_slice(&(present as u16).to_le_bytes());
    true
}

/// Generic OOXML archive wrapper for reading parts from ZIP files
pub struct OoxmlArchive {
    archive: ZipArchive<Cursor<Vec<u8>>>,
}

impl OoxmlArchive {
    pub fn new(data: &[u8]) -> Result<Self, OxiError> {
        security::validate_input_size(data)?;

        let archive = open_zip(data.to_vec())?;

        if archive.len() > security::MAX_ZIP_ENTRIES {
            return Err(OxiError::Security(format!(
                "Too many ZIP entries: {} (max {})",
                archive.len(),
                security::MAX_ZIP_ENTRIES
            )));
        }

        Ok(Self { archive })
    }

    /// Read a text part from the archive with size validation
    pub fn read_part(&mut self, name: &str) -> Result<String, OxiError> {
        let mut file = self
            .archive
            .by_name(name)
            .map_err(|_| OxiError::MissingPart(name.to_string()))?;

        security::validate_zip_entry(
            name,
            file.compressed_size(),
            file.size(),
        )?;

        let mut contents = String::with_capacity(file.size().min(security::MAX_ENTRY_SIZE) as usize);
        file.read_to_string(&mut contents)?;
        Ok(contents)
    }

    /// Read a binary part from the archive with size validation
    pub fn read_binary_part(&mut self, name: &str) -> Result<Vec<u8>, OxiError> {
        let mut file = self
            .archive
            .by_name(name)
            .map_err(|_| OxiError::MissingPart(name.to_string()))?;

        security::validate_zip_entry(
            name,
            file.compressed_size(),
            file.size(),
        )?;

        let mut contents = Vec::with_capacity(file.size().min(security::MAX_ENTRY_SIZE) as usize);
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

    /// Try to read a binary part, returning None if not found
    pub fn try_read_bytes(&mut self, name: &str) -> Result<Option<Vec<u8>>, OxiError> {
        match self.read_binary_part(name) {
            Ok(bytes) => Ok(Some(bytes)),
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

#[cfg(test)]
mod open_zip_tests {
    use super::*;
    use std::io::Write;

    fn package() -> Vec<u8> {
        let mut out = Cursor::new(Vec::new());
        let mut zip = zip::ZipWriter::new(&mut out);
        zip.start_file("a.xml", zip::write::SimpleFileOptions::default()).unwrap();
        zip.write_all(b"<a/>").unwrap();
        zip.finish().unwrap();
        out.into_inner()
    }

    #[test]
    fn overlong_comment_length_is_clamped() {
        let mut data = package();
        let n = data.len();
        data[n - 2..].copy_from_slice(&0x0d00u16.to_le_bytes());
        data.extend_from_slice(b"\n\r\n");
        assert!(ZipArchive::new(Cursor::new(data.clone())).is_err());
        let mut archive = open_zip(data).unwrap();
        let mut text = String::new();
        archive.by_name("a.xml").unwrap().read_to_string(&mut text).unwrap();
        assert_eq!(text, "<a/>");
    }

    #[test]
    fn valid_package_is_untouched() {
        let mut data = package();
        assert!(!clamp_eocd_comment(&mut data));
        assert_eq!(data, package());
    }
}
