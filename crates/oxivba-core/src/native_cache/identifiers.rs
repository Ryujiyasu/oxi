// This Source Code Form is subject to the terms of the Mozilla Public
// License, v. 2.0. If a copy of the MPL was not distributed with this
// file, You can obtain one at https://mozilla.org/MPL/2.0/.

//! Identifier records from a caller-located native cache table span.

use super::{CacheContext, DecodeError, DecodeErrorKind, Reader};

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct CachedIdentifier<'a> {
    pub symbol: u16,
    /// Encoded bytes, not UTF-8. The host chooses the effective code page.
    /// Internal cached names need not be valid VBA language identifiers.
    pub name: &'a [u8],
    pub flags: u8,
    /// Preserved for subsequent metadata readers; not interpreted as pointers.
    pub prefix: [u8; 4],
    pub extension: Option<[u8; 6]>,
}

/// Additional encoded names do not allocate a source-symbol ordinal. Their
/// reference words are preserved without assuming they are user symbols.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct CachedUnindexedIdentifier<'a> {
    pub before_symbol: u16,
    pub reference: u16,
    pub name: &'a [u8],
    pub flags: u8,
    pub extension: Option<[u8; 6]>,
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct IdentifierTable<'a> {
    total: u16,
    reserved: u16,
    record_count: u16,
    identifiers: Vec<CachedIdentifier<'a>>,
    unindexed: Vec<CachedUnindexedIdentifier<'a>>,
    consumed: usize,
}

impl<'a> IdentifierTable<'a> {
    pub fn total(&self) -> u16 {
        self.total
    }
    pub fn reserved(&self) -> u16 {
        self.reserved
    }
    pub fn identifiers(&self) -> &[CachedIdentifier<'a>] {
        &self.identifiers
    }
    pub fn record_count(&self) -> u16 {
        self.record_count
    }
    pub fn unindexed_identifiers(&self) -> &[CachedUnindexedIdentifier<'a>] {
        &self.unindexed
    }
    /// The table can be followed by other project-cache structures.
    pub fn consumed(&self) -> usize {
        self.consumed
    }

    /// Resolve an encoded source symbol. Reserved, odd, or absent references
    /// return None; this reader does not invent names for built-in symbols.
    pub fn get(&self, symbol: u16) -> Option<&CachedIdentifier<'a>> {
        if symbol & 1 != 0 {
            return None;
        }
        let ordinal = u32::from(symbol / 2).checked_sub(u32::from(self.reserved) + 1)?;
        self.identifiers.get(ordinal as usize)
    }
}

/// Decode a B5/Win64 identifier table starting at its six-byte count header.
/// Locating this table in a complete project and decoding its byte strings
/// remain host responsibilities. No content-search heuristic is used here.
pub fn decode_identifier_table(
    bytes: &[u8],
    context: CacheContext,
) -> Result<IdentifierTable<'_>, DecodeError> {
    context.validate()?;
    let mut reader = Reader::new(bytes, 0);
    let total = reader.word()?;
    let named = reader.word()?;
    let reserved = reader.word()?;
    let indexed = total.checked_sub(reserved);
    if indexed.is_none_or(|count| count > named) {
        return Err(reader.error(DecodeErrorKind::InvalidIdentifierCounts {
            total,
            named,
            reserved,
        }));
    }
    let indexed = indexed.unwrap();
    let extra_count = usize::from(named - indexed);
    if total > u16::MAX / 2 {
        return Err(reader.error(DecodeErrorKind::IdentifierSymbolOverflow));
    }
    let count = usize::from(named);
    // Each nonempty record needs at least four prefix bytes, a length and
    // flags, and one name byte. Check before allocating the record vector.
    let remaining = bytes.len() - reader.at;
    if count > remaining / 7 {
        return Err(reader.error(DecodeErrorKind::Truncated {
            needed: count * 7,
            remaining,
        }));
    }
    let mut identifiers = Vec::with_capacity(usize::from(indexed));
    let mut unindexed = Vec::with_capacity(extra_count);
    for ordinal in 0..indexed {
        let prefix = reader.take()?;
        let symbol = ((u32::from(reserved) + u32::from(ordinal) + 1) * 2) as u16;
        while reader.peek_word() == Some(0) {
            if unindexed.len() == extra_count {
                return Err(reader.error(DecodeErrorKind::InvalidIdentifierCounts {
                    total,
                    named,
                    reserved,
                }));
            }
            reader.word()?;
            let reference = reader.word()?;
            let length = usize::from(reader.take::<1>()?[0]);
            let flags = reader.take::<1>()?[0];
            if length == 0 {
                return Err(reader.error(DecodeErrorKind::EmptyIdentifier));
            }
            let extension = if flags & 0x80 != 0 {
                Some(reader.take()?)
            } else {
                None
            };
            let name = reader.slice(length)?;
            unindexed.push(CachedUnindexedIdentifier {
                before_symbol: symbol,
                reference,
                name,
                flags,
                extension,
            });
        }
        let length = usize::from(reader.take::<1>()?[0]);
        let flags = reader.take::<1>()?[0];
        if length == 0 {
            return Err(reader.error(DecodeErrorKind::EmptyIdentifier));
        }
        let extension = if flags & 0x80 != 0 {
            Some(reader.take()?)
        } else {
            None
        };
        let name = reader.slice(length)?;
        identifiers.push(CachedIdentifier {
            symbol,
            name,
            flags,
            prefix,
            extension,
        });
    }
    if unindexed.len() != extra_count {
        return Err(reader.error(DecodeErrorKind::InvalidIdentifierCounts {
            total,
            named,
            reserved,
        }));
    }
    Ok(IdentifierTable {
        total,
        reserved,
        record_count: named,
        identifiers,
        unindexed,
        consumed: reader.at,
    })
}
