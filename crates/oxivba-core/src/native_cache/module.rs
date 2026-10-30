// This Source Code Form is subject to the terms of the Mozilla Public
// License, v. 2.0. If a copy of the MPL was not distributed with this
// file, You can obtain one at https://mozilla.org/MPL/2.0/.

//! Bounded module line frames and declaration records for B5/Win64 caches.

use super::{CacheContext, DecodeError, DecodeErrorKind};
use std::collections::BTreeSet;

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct CachedLine<'a> {
    /// Zero-based logical source line; no source decoding is performed.
    pub index: u16,
    pub flags: u16,
    pub extra: u16,
    pub reserved: u16,
    pub code: &'a [u8],
}

/// Inline type codes retain their qualifiers. Referenced types (including
/// arrays and declared object types) require a separate metadata resolver.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum CachedType {
    Builtin { kind: u8, qualifiers: u8 },
    Reference(u32),
}

impl CachedType {
    fn from_word(word: u32) -> Self {
        if word >> 16 == 0xffff {
            Self::Builtin {
                kind: word as u8,
                qualifiers: (word >> 8) as u8,
            }
        } else {
            Self::Reference(word)
        }
    }
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct CachedParameter {
    pub declaration_offset: u32,
    pub flags: u16,
    /// Raw identifier reference; array shape belongs to type metadata.
    pub name_word: u16,
    pub type_descriptor: CachedType,
    pub storage_word: u32,
    pub mode_flags: u32,
}

impl CachedParameter {
    /// Remaining modifier/default flags are preserved in mode_flags.
    pub fn passing_mode(&self) -> Option<super::super::ast::ParamMode> {
        use super::super::ast::ParamMode;
        match self.mode_flags & 7 {
            1 => Some(ParamMode::ParamArray),
            2 => Some(ParamMode::ByRef),
            4 => Some(ParamMode::ByVal),
            _ => None,
        }
    }
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct CachedProcedure<'a> {
    pub source_line: u16,
    pub definition_opcode: u16,
    /// Includes default-value instructions preceding the declaration. They
    /// are not discarded, evaluated, or assumed to be part of the body.
    pub definition_prefix: &'a [u8],
    pub declaration_offset: u32,
    pub flags: u16,
    pub name_word: u16,
    pub member_id: u32,
    pub return_type: Option<CachedType>,
    pub parameters: Vec<CachedParameter>,
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct ModuleCache<'a> {
    bytes: &'a [u8],
    lines: Vec<CachedLine<'a>>,
}

fn fail(offset: usize, kind: DecodeErrorKind) -> DecodeError {
    DecodeError { offset, kind }
}
fn span(bytes: &[u8], at: usize, count: usize) -> Result<&[u8], DecodeError> {
    bytes
        .get(at..)
        .and_then(|tail| tail.get(..count))
        .ok_or_else(|| {
            fail(
                at,
                DecodeErrorKind::Truncated {
                    needed: count,
                    remaining: bytes.len().saturating_sub(at),
                },
            )
        })
}
fn word(bytes: &[u8], at: usize) -> Result<u16, DecodeError> {
    let value = span(bytes, at, 2)?;
    Ok(u16::from_le_bytes([value[0], value[1]]))
}
fn dword(bytes: &[u8], at: usize) -> Result<u32, DecodeError> {
    let value = span(bytes, at, 4)?;
    Ok(u32::from_le_bytes([value[0], value[1], value[2], value[3]]))
}
fn add(at: usize, count: usize) -> Result<usize, DecodeError> {
    at.checked_add(count)
        .ok_or_else(|| fail(at, DecodeErrorKind::InvalidModuleCache))
}

/// Decode a module's performance-cache prefix, excluding compressed source.
/// Header offsets describe the measured input format, not procedure positions.
pub fn decode_module_cache(
    bytes: &[u8],
    context: CacheContext,
) -> Result<ModuleCache<'_>, DecodeError> {
    context.validate()?;
    let marker = add(dword(bytes, 25)? as usize, 60)?;
    if span(bytes, marker, 4)? != [0xfe, 0xca, 1, 0] {
        return Err(fail(marker, DecodeErrorKind::InvalidModuleCache));
    }
    let count = usize::from(word(bytes, add(marker, 4)?)?);
    let table = add(marker, 6)?;
    let end = add(table, count * 12)?;
    span(bytes, table, count * 12)?;
    if span(bytes, end, 6)? != [0xff, 0xff, 0xff, 0xff, 1, 1] {
        return Err(fail(end, DecodeErrorKind::InvalidModuleCache));
    }
    let code_size = dword(bytes, add(end, 6)?)? as usize;
    let code = add(end, 10)?;
    let payload = span(bytes, code, code_size)?;
    let mut lines = Vec::with_capacity(count);
    for index in 0..count {
        let at = table + index * 12;
        let length = usize::from(word(bytes, at + 4)?);
        let offset = dword(bytes, at + 8)?;
        let data = if length == 0 {
            if offset != u32::MAX {
                return Err(fail(at + 8, DecodeErrorKind::InvalidModuleCache));
            }
            &[][..]
        } else {
            span(payload, offset as usize, length)
                .map_err(|_| fail(at + 8, DecodeErrorKind::InvalidModuleCache))?
        };
        lines.push(CachedLine {
            index: index as u16,
            flags: word(bytes, at)?,
            extra: word(bytes, at + 2)?,
            reserved: word(bytes, at + 6)?,
            code: data,
        });
    }
    Ok(ModuleCache { bytes, lines })
}

impl<'a> ModuleCache<'a> {
    pub fn lines(&self) -> &[CachedLine<'a>] {
        &self.lines
    }

    /// Read procedure references from framed source lines and follow their
    /// bounded declaration/parameter offsets. Raw opcodes, modifiers, default
    /// instructions and unresolved type references remain available to callers.
    pub fn procedures(&self) -> Result<Vec<CachedProcedure<'a>>, DecodeError> {
        let mut references = Vec::new();
        for line in &self.lines {
            if line.code.len() < 6 {
                continue;
            }
            let tail = line.code.len() - 6;
            let opcode = word(line.code, tail)?;
            if opcode & 0x3ff == 0x96 {
                references.push((line, tail, opcode, dword(line.code, tail + 2)?));
            }
        }
        if references.is_empty() {
            return Ok(Vec::new());
        }
        let header = dword(self.bytes, 17)? as usize;
        let origin = add(header, 16)?;
        let size = dword(self.bytes, add(header, 12)?)? as usize;
        let declarations = span(self.bytes, origin, size)?;
        let mut result = Vec::with_capacity(references.len());
        for (line, tail, opcode, offset) in references {
            let at = offset as usize;
            let record = span(declarations, at, 88)?;
            let flags = word(record, 0)?;
            if flags & 15 != 11 {
                return Err(fail(
                    add(origin, at)?,
                    DecodeErrorKind::InvalidDeclarationRecord,
                ));
            }
            let mut next = dword(record, 56)?;
            let mut visited = BTreeSet::new();
            let mut parameters = Vec::new();
            while next != u32::MAX {
                if !visited.insert(next) || visited.len() > declarations.len() / 32 {
                    return Err(fail(
                        add(origin, next as usize)?,
                        DecodeErrorKind::CyclicParameterChain,
                    ));
                }
                let raw = span(declarations, next as usize, 32)?;
                let parameter_flags = word(raw, 0)?;
                if parameter_flags & 15 != 9 {
                    return Err(fail(
                        add(origin, next as usize)?,
                        DecodeErrorKind::InvalidDeclarationRecord,
                    ));
                }
                parameters.push(CachedParameter {
                    declaration_offset: next,
                    flags: parameter_flags,
                    name_word: word(raw, 2)?,
                    type_descriptor: CachedType::from_word(dword(raw, 16)?),
                    storage_word: dword(raw, 8)?,
                    mode_flags: dword(raw, 28)?,
                });
                next = dword(raw, 24)?;
            }
            let return_word = dword(record, 60)?;
            result.push(CachedProcedure {
                source_line: line.index,
                definition_opcode: opcode,
                definition_prefix: &line.code[..tail],
                declaration_offset: offset,
                flags,
                name_word: word(record, 2)?,
                member_id: dword(record, 8)?,
                return_type: if return_word == u32::MAX {
                    None
                } else {
                    Some(CachedType::from_word(return_word))
                },
                parameters,
            });
        }
        Ok(result)
    }
}
