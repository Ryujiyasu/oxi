// This Source Code Form is subject to the terms of the Mozilla Public
// License, v. 2.0. If a copy of the MPL was not distributed with this
// file, You can obtain one at https://mozilla.org/MPL/2.0/.

//! Global runtime objects and canonical encoded names, independent of source
//! identifiers. ANSI decoding, case comparison, companion invocation signatures
//! and execution-target selection remain separate operations.

use super::{
    decode_runtime_stream_directory, CacheContext, DecodeError, DecodeErrorKind, Reader,
    RuntimeCacheLink, RuntimeHeapSpan, RuntimeStreamDirectory, RuntimeStreamKind,
};

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum RuntimeNameEncoding {
    Ansi,
    Utf16Le,
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct RuntimeEncodedName<'a> {
    pub encoding: RuntimeNameEncoding,
    /// ANSI bytes require the native name code page; no Latin-1/UTF-8 guess is made.
    pub bytes: &'a [u8],
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct RuntimeReadyModule {
    pub root: RuntimeCacheLink,
    pub type_link: RuntimeCacheLink,
    pub extra: Option<RuntimeCacheLink>,
    pub standard_storage: Option<([u8; 12], Option<RuntimeCacheLink>)>,
    /// Canonical name-table entries; these do not replace companion target links.
    pub procedure_names: Vec<RuntimeCacheLink>,
    pub trailer: u16,
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub enum RuntimeProjectModuleState {
    Ready(RuntimeReadyModule),
    Unavailable { raw: [u8; 24], word: u32 },
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct RuntimeProjectModule {
    pub flags: u32,
    pub name: RuntimeCacheLink,
    pub state: RuntimeProjectModuleState,
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct RuntimeProjectReferenceEntry {
    pub target: RuntimeCacheLink,
    pub raw: [u8; 16],
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub enum RuntimeProjectObjectKind<'a> {
    Name(RuntimeEncodedName<'a>),
    Guid([u8; 16]),
    LinkArray {
        /// Observed values include 0 and 2. This is not the number of links.
        word: u16,
        entries: [RuntimeCacheLink; 2],
        extra: Option<RuntimeCacheLink>,
    },
    Field {
        target: RuntimeCacheLink,
        raw: [u8; 12],
    },
    Member {
        target: RuntimeCacheLink,
        extra: Option<RuntimeCacheLink>,
        flags: [u8; 4],
        signature: RuntimeCacheLink,
        raw: [u8; 12],
    },
    ReferenceLibrary {
        guid: RuntimeCacheLink,
        extra: Option<RuntimeCacheLink>,
        settings: [u8; 8],
        path: Option<RuntimeCacheLink>,
        name: RuntimeCacheLink,
        tail: [u8; 18],
    },
    Bstr(&'a [u8]),
    ValueRecord {
        raw: &'a [u8],
        following_references: Vec<RuntimeCacheLink>,
    },
    Signature {
        header: [u8; 4],
        return_type: u8,
        parameter_types: &'a [u8],
        owner: Option<RuntimeCacheLink>,
    },
    Zeroed,
    Vacant,
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct RuntimeProjectObject<'a> {
    pub address: u64,
    pub allocation_units: u64,
    pub serialized_offset: usize,
    pub kind: RuntimeProjectObjectKind<'a>,
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct RuntimeProjectHeap<'a> {
    pub directory: RuntimeStreamDirectory,
    pub header_words: [u64; 4],
    pub zero_arena: Vec<RuntimeHeapSpan>,
    pub owner: RuntimeCacheLink,
    pub fixed: [u8; 12],
    pub name: RuntimeCacheLink,
    pub guid: [u8; 16],
    pub settings: [u8; 18],
    pub active_word: u16,
    pub extra: Option<RuntimeCacheLink>,
    pub modules: Vec<RuntimeProjectModule>,
    pub references: Vec<RuntimeProjectReferenceEntry>,
    pub objects: Vec<RuntimeProjectObject<'a>>,
    pub consumed: usize,
}

impl<'a> RuntimeProjectHeap<'a> {
    pub fn object_for_link(&self, link: RuntimeCacheLink) -> Option<&RuntimeProjectObject<'a>> {
        if link.context != 0 || link.key & 15 != 1 {
            return None;
        }
        let mut matches = self.objects.iter().filter(|o| o.address == link.key - 1);
        let object = matches.next()?;
        matches.next().is_none().then_some(object)
    }
    pub fn name_for_link(&self, link: RuntimeCacheLink) -> Option<RuntimeEncodedName<'a>> {
        match self.object_for_link(link)?.kind {
            RuntimeProjectObjectKind::Name(name) => Some(name),
            _ => None,
        }
    }
    pub fn module_name(&self, index: usize) -> Option<RuntimeEncodedName<'a>> {
        self.name_for_link(self.modules.get(index)?.name)
    }
    /// Reads a canonical entry name. Entry order does not select executable code
    /// when companion target links have been changed independently.
    pub fn procedure_name(&self, module: usize, entry: usize) -> Option<RuntimeEncodedName<'a>> {
        let RuntimeProjectModuleState::Ready(ready) = &self.modules.get(module)?.state else {
            return None;
        };
        self.name_for_link(*ready.procedure_names.get(entry)?)
    }
}

fn invalid(r: &Reader<'_>) -> DecodeError {
    r.error(DecodeErrorKind::InvalidRuntimeHeap)
}
fn dword(r: &mut Reader<'_>) -> Result<u32, DecodeError> {
    Ok(u32::from_le_bytes(r.take()?))
}
fn qword(r: &mut Reader<'_>) -> Result<u64, DecodeError> {
    Ok(u64::from_le_bytes(r.take()?))
}
fn link(r: &mut Reader<'_>) -> Result<RuntimeCacheLink, DecodeError> {
    Ok(RuntimeCacheLink {
        key: qword(r)?,
        context: dword(r)?,
    })
}
fn optional(r: &mut Reader<'_>) -> Result<Option<RuntimeCacheLink>, DecodeError> {
    let key = qword(r)?;
    Ok(if key == u64::MAX {
        None
    } else {
        Some(RuntimeCacheLink {
            key,
            context: dword(r)?,
        })
    })
}
fn limit(r: &Reader<'_>, units: u64) -> Result<u64, DecodeError> {
    units
        .checked_mul(16)
        .and_then(|n| n.checked_add(16))
        .ok_or_else(|| invalid(r))
}
fn advance(r: &Reader<'_>, address: u64, units: u64, end: u64) -> Result<u64, DecodeError> {
    if units == 0 || units & 1 != 0 {
        return Err(invalid(r));
    }
    units
        .checked_mul(8)
        .and_then(|n| n.checked_add(16))
        .and_then(|n| n.checked_add(address))
        .filter(|n| *n <= end)
        .ok_or_else(|| invalid(r))
}
fn count_links(r: &mut Reader<'_>) -> Result<Vec<RuntimeCacheLink>, DecodeError> {
    let n = usize::from(r.word()?);
    if n > r.bytes.len().saturating_sub(r.at) / 12 {
        return Err(invalid(r));
    }
    (0..n).map(|_| link(r)).collect()
}
fn sized_bytes<'a>(r: &mut Reader<'a>, n: u64, wide: bool) -> Result<&'a [u8], DecodeError> {
    let n = usize::try_from(n).map_err(|_| invalid(r))?;
    if wide && n & 1 != 0 {
        return Err(invalid(r));
    }
    r.slice(n)
}

/// Decode the global `__SRP_0` stream, including its stream directory. Names
/// preserve their stored encoding; project settings do not override native ANSI
/// name encoding. Unsupported object encodings are rejected, never scanned past.
pub fn decode_runtime_project_heap(
    bytes: &[u8],
    context: CacheContext,
) -> Result<RuntimeProjectHeap<'_>, DecodeError> {
    let directory = decode_runtime_stream_directory(bytes, context)?;
    let mut r = Reader::new(bytes, 0);
    r.at = directory.root_heap_offset;
    if r.word()? != 0x5572 {
        return Err(invalid(&r));
    }
    let mut header_words = [0; 4];
    for word in &mut header_words {
        *word = qword(&mut r)?;
    }
    let zero_end = limit(&r, header_words[1])?;
    let object_end = limit(&r, header_words[0])?;
    let mut address = 16;
    let mut zero_arena = Vec::new();
    loop {
        let tag = qword(&mut r)?;
        if tag == 0 {
            break;
        }
        let kind = (tag >> 56) as u8;
        if !matches!(kind, 0x7e | 0x7f) {
            return Err(invalid(&r));
        }
        let units = tag & 0x00ff_ffff_ffff_ffff;
        let next = advance(&r, address, units, zero_end)?;
        zero_arena.push(RuntimeHeapSpan {
            address,
            allocation_units: units,
            zeroed: kind == 0x7e,
        });
        address = next;
    }
    if address != zero_end || qword(&mut r)? != 34 {
        return Err(invalid(&r));
    }
    let owner = link(&mut r)?;
    if owner
        != (RuntimeCacheLink {
            key: 17,
            context: 1 << 16,
        })
    {
        return Err(invalid(&r));
    }
    let fixed = r.take()?;
    let name = link(&mut r)?;
    let guid = r.take()?;
    let settings = r.take()?;
    let active_word = r.word()?;
    let extra = optional(&mut r)?;
    let count = usize::from(r.word()?);
    if count > r.bytes.len().saturating_sub(r.at) / 44 {
        return Err(invalid(&r));
    }
    let mut modules = Vec::with_capacity(count);
    for index in 0..count {
        let flags = dword(&mut r)?;
        let state = if flags & 0x8000 != 0 {
            let root = link(&mut r)?;
            let stream = directory
                .unique_stream(Some(index as u16), RuntimeStreamKind::Primary)
                .ok_or_else(|| invalid(&r))?;
            if root
                != (RuntimeCacheLink {
                    key: 17,
                    context: u32::from(stream) << 16,
                })
            {
                return Err(invalid(&r));
            }
            let type_link = link(&mut r)?;
            let extra = optional(&mut r)?;
            let standard_storage = if flags & 15 == 1 {
                Some((r.take()?, optional(&mut r)?))
            } else {
                None
            };
            let procedure_names = count_links(&mut r)?;
            if r.word()? != 0xffff {
                return Err(invalid(&r));
            }
            let trailer = r.word()?;
            RuntimeProjectModuleState::Ready(RuntimeReadyModule {
                root,
                type_link,
                extra,
                standard_storage,
                procedure_names,
                trailer,
            })
        } else {
            let raw = r.take()?;
            if raw != [255; 24] {
                return Err(invalid(&r));
            }
            RuntimeProjectModuleState::Unavailable {
                raw,
                word: dword(&mut r)?,
            }
        };
        modules.push(RuntimeProjectModule {
            flags,
            name: link(&mut r)?,
            state,
        });
    }
    let count = usize::from(r.word()?);
    if count > r.bytes.len().saturating_sub(r.at) / 28 {
        return Err(invalid(&r));
    }
    let mut references = Vec::with_capacity(count);
    for _ in 0..count {
        references.push(RuntimeProjectReferenceEntry {
            target: link(&mut r)?,
            raw: r.take()?,
        });
    }
    address = advance(&r, 16, 34, object_end)?;
    let mut objects = Vec::new();
    loop {
        let serialized_offset = r.at;
        let tag = qword(&mut r)?;
        if tag == 0 {
            break;
        }
        let allocation_units = tag & 0x00ff_ffff_ffff_ffff;
        let next = advance(&r, address, allocation_units, object_end)?;
        let kind = match (tag >> 56) as u8 {
            k @ (1 | 2) => {
                let n = qword(&mut r)?;
                RuntimeProjectObjectKind::Name(RuntimeEncodedName {
                    encoding: if k == 1 {
                        RuntimeNameEncoding::Utf16Le
                    } else {
                        RuntimeNameEncoding::Ansi
                    },
                    bytes: sized_bytes(&mut r, n, k == 1)?,
                })
            }
            3 => RuntimeProjectObjectKind::Guid(r.take()?),
            5 => RuntimeProjectObjectKind::LinkArray {
                word: r.word()?,
                entries: [link(&mut r)?, link(&mut r)?],
                extra: optional(&mut r)?,
            },
            6 => RuntimeProjectObjectKind::Field {
                target: link(&mut r)?,
                raw: r.take()?,
            },
            7 => RuntimeProjectObjectKind::Member {
                target: link(&mut r)?,
                extra: optional(&mut r)?,
                flags: r.take()?,
                signature: link(&mut r)?,
                raw: r.take()?,
            },
            0x0a => RuntimeProjectObjectKind::ReferenceLibrary {
                guid: link(&mut r)?,
                extra: optional(&mut r)?,
                settings: r.take()?,
                path: optional(&mut r)?,
                name: link(&mut r)?,
                tail: r.take()?,
            },
            0x0b => {
                let n = dword(&mut r)?;
                RuntimeProjectObjectKind::Bstr(sized_bytes(&mut r, u64::from(n), true)?)
            }
            0x0d => {
                let n = usize::from(r.word()?);
                let raw = r.slice(n)?;
                let mut following_references = Vec::new();
                while r
                    .bytes
                    .get(r.at..)
                    .and_then(|b| b.get(..8))
                    .is_some_and(|b| u64::from_le_bytes(b.try_into().unwrap()) & 15 == 1)
                {
                    following_references.push(link(&mut r)?);
                }
                RuntimeProjectObjectKind::ValueRecord {
                    raw,
                    following_references,
                }
            }
            0x0e => {
                let header: [u8; 4] = r.take()?;
                if !matches!(header[0], 0 | 2 | 3) {
                    return Err(invalid(&r));
                }
                let return_type = r.take::<1>()?[0];
                let parameter_types = r.slice(usize::from(header[1]))?;
                let owner = if header[0] == 3 {
                    Some(link(&mut r)?)
                } else {
                    None
                };
                RuntimeProjectObjectKind::Signature {
                    header,
                    return_type,
                    parameter_types,
                    owner,
                }
            }
            0x7e => RuntimeProjectObjectKind::Zeroed,
            0x7f => RuntimeProjectObjectKind::Vacant,
            k => return Err(r.error(DecodeErrorKind::UnsupportedRuntimeHeapObject(k))),
        };
        objects.push(RuntimeProjectObject {
            address,
            allocation_units,
            serialized_offset,
            kind,
        });
        address = next;
    }
    if address != object_end || r.at != bytes.len() {
        return Err(invalid(&r));
    }
    let result = RuntimeProjectHeap {
        directory,
        header_words,
        zero_arena,
        owner,
        fixed,
        name,
        guid,
        settings,
        active_word,
        extra,
        modules,
        references,
        objects,
        consumed: r.at,
    };
    if result.name_for_link(result.name).is_none() {
        return Err(invalid(&r));
    }
    for (i, module) in result.modules.iter().enumerate() {
        if result.module_name(i).is_none() {
            return Err(invalid(&r));
        }
        if let RuntimeProjectModuleState::Ready(ready) = &module.state {
            if ready
                .procedure_names
                .iter()
                .any(|l| result.name_for_link(*l).is_none())
            {
                return Err(invalid(&r));
            }
        }
    }
    Ok(result)
}
