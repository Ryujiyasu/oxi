// This Source Code Form is subject to the terms of the Mozilla Public
// License, v. 2.0. If a copy of the MPL was not distributed with this
// file, You can obtain one at https://mozilla.org/MPL/2.0/.

//! Structural B5/Win64 `_VBA_PROJECT` metadata, without content searches.

use super::{
    decode_identifier_table, CacheContext, DecodeError, DecodeErrorKind, IdentifierTable, Reader,
};

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct CachedProjectReference<'a> {
    /// Optional UTF-16LE identity; no filesystem access or library loading.
    pub identity: Option<&'a [u8]>,
    pub metadata: [u8; 12],
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct CachedProjectModule<'a> {
    pub kind: u16,
    /// Encoded UTF-16LE metadata; these are not assumed to be CFB stream paths.
    pub name: Option<&'a [u8]>,
    pub cookie: Option<&'a [u8]>,
    pub alias: Option<&'a [u8]>,
    pub name_metadata: [u8; 4],
    pub header: [u8; 10],
    pub bindings: Vec<[u32; 2]>,
    pub trailer: [u8; 11],
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct ProjectCache<'a> {
    pub context: CacheContext,
    pub code_page: u16,
    pub reference_flags: u16,
    pub references: Vec<CachedProjectReference<'a>>,
    pub namespace_bindings: Vec<[u16; 2]>,
    pub project_name_word: u16,
    pub project_flags: u16,
    pub description: Option<&'a [u8]>,
    pub help_file: Option<&'a [u8]>,
    /// Fixed-format fields retained for future metadata interpretation.
    pub metadata: [u8; 100],
    pub modules: Vec<CachedProjectModule<'a>>,
    pub index_header: [u8; 10],
    pub index_bytes: &'a [u8],
    pub index_trailer: [u8; 6],
    pub name_table_offset: usize,
    pub identifiers: IdentifierTable<'a>,
    /// Through the last decoded name; following project structures remain.
    pub consumed: usize,
}

fn invalid(reader: &Reader<'_>) -> DecodeError {
    reader.error(DecodeErrorKind::InvalidProjectCache)
}
fn string<'a>(reader: &mut Reader<'a>) -> Result<Option<&'a [u8]>, DecodeError> {
    let length = reader.word()?;
    if length == u16::MAX {
        return Ok(None);
    }
    if length & 1 != 0 {
        return Err(invalid(reader));
    }
    reader.slice(usize::from(length)).map(Some)
}
fn count(reader: &mut Reader<'_>, minimum_size: usize) -> Result<usize, DecodeError> {
    let count = usize::from(reader.word()?);
    let remaining = reader.bytes.len() - reader.at;
    if count > remaining / minimum_size {
        return Err(reader.error(DecodeErrorKind::Truncated {
            needed: count * minimum_size,
            remaining,
        }));
    }
    Ok(count)
}
fn dword(reader: &mut Reader<'_>) -> Result<u32, DecodeError> {
    Ok(u32::from_le_bytes(reader.take()?))
}

/// Locate names by traversing version-checked references, module descriptors
/// and the length-prefixed index. Does not activate or execute cached code.
pub fn decode_project_cache(
    bytes: &[u8],
    context: CacheContext,
) -> Result<ProjectCache<'_>, DecodeError> {
    context.validate()?;
    let mut reader = Reader::new(bytes, 0);
    let header = reader.take::<30>()?;
    if header[..2] != [0xcc, 0x61] {
        return Err(invalid(&reader));
    }
    let actual = CacheContext {
        version: u16::from_le_bytes([header[2], header[3]]),
        sys_kind: u32::from_le_bytes(header[18..22].try_into().unwrap()),
    };
    actual.validate()?;
    if actual != context {
        return Err(invalid(&reader));
    }
    let reference_count = count(&mut reader, 14)?;
    let reference_flags = reader.word()?;
    let mut references = Vec::with_capacity(reference_count);
    for _ in 0..reference_count {
        references.push(CachedProjectReference {
            identity: string(&mut reader)?,
            metadata: reader.take()?,
        });
    }
    let module_count = count(&mut reader, 2)?;
    let kinds = (0..module_count)
        .map(|_| reader.word())
        .collect::<Result<Vec<_>, _>>()?;
    let binding_count = count(&mut reader, 4)?;
    let mut namespace_bindings = Vec::with_capacity(binding_count);
    for _ in 0..binding_count {
        namespace_bindings.push([reader.word()?, reader.word()?]);
    }
    let project_name_word = reader.word()?;
    let project_flags = reader.word()?;
    let description = string(&mut reader)?;
    let help_file = string(&mut reader)?;
    let metadata = reader.take()?;
    if usize::from(reader.word()?) != module_count {
        return Err(invalid(&reader));
    }
    if module_count > (bytes.len() - reader.at) / 31 {
        return Err(invalid(&reader));
    }
    let mut modules = Vec::with_capacity(module_count);
    for kind in kinds {
        let name = string(&mut reader)?;
        let cookie = string(&mut reader)?;
        let name_metadata = reader.take()?;
        let alias = string(&mut reader)?;
        let header = reader.take::<10>()?;
        let binding_count = usize::from(u16::from_le_bytes([header[8], header[9]]));
        if binding_count > (bytes.len() - reader.at) / 8 {
            return Err(invalid(&reader));
        }
        let mut bindings = Vec::with_capacity(binding_count);
        for _ in 0..binding_count {
            bindings.push([dword(&mut reader)?, dword(&mut reader)?]);
        }
        modules.push(CachedProjectModule {
            kind,
            name,
            cookie,
            alias,
            name_metadata,
            header,
            bindings,
            trailer: reader.take()?,
        });
    }
    let index_header = reader.take::<10>()?;
    if index_header[4..6] != [1, 1] {
        return Err(invalid(&reader));
    }
    let index_size = u32::from_le_bytes(index_header[6..10].try_into().unwrap()) as usize;
    let index_bytes = reader.slice(index_size)?;
    let index_trailer = reader.take()?;
    let name_table_offset = reader.at;
    let identifiers =
        decode_identifier_table(&bytes[name_table_offset..], context).map_err(|mut error| {
            error.offset += name_table_offset;
            error
        })?;
    reader.slice(identifiers.consumed())?;
    Ok(ProjectCache {
        context,
        code_page: u16::from_le_bytes([header[16], header[17]]),
        reference_flags,
        references,
        namespace_bindings,
        project_name_word,
        project_flags,
        description,
        help_file,
        metadata,
        modules,
        index_header,
        index_bytes,
        index_trailer,
        name_table_offset,
        identifiers,
        consumed: reader.at,
    })
}
