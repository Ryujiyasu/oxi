// This Source Code Form is subject to the terms of the Mozilla Public
// License, v. 2.0. If a copy of the MPL was not distributed with this
// file, You can obtain one at https://mozilla.org/MPL/2.0/.

//! Runtime module prefixes, including class and host-event binding tables.

use super::{
    decode_runtime_record, CacheContext, DecodeError, DecodeErrorKind, Reader, RuntimeRecord,
};

/// Opaque packed cache key and context word. Neither field is a native pointer.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct RuntimeCacheLink {
    pub key: u64,
    pub context: u32,
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct RuntimeHostEvent {
    pub dispatch_id: u32,
    pub member_id: u32,
    pub target: RuntimeCacheLink,
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub enum RuntimeBindingSlots<'a> {
    Class([Option<RuntimeCacheLink>; 2]),
    Host {
        capacity: u16,
        events: Vec<RuntimeHostEvent>,
        unused: &'a [u8],
    },
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct RuntimeClassBinding<'a> {
    pub version: [u8; 8],
    pub type_links: Vec<Option<RuntimeCacheLink>>,
    pub storage_header: [u8; 8],
    pub shared_links: [Option<RuntimeCacheLink>; 6],
    pub slots: RuntimeBindingSlots<'a>,
    pub directory_header: Option<[u16; 2]>,
    pub directory_word: u16,
    /// Includes cached field accessors; not a one-to-one procedure-slot map.
    pub invocation_links: Vec<RuntimeCacheLink>,
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct RuntimeRelocation {
    pub handle: u64,
    pub offset: u32,
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct RuntimeMethodImage<'a> {
    pub code: &'a [u8],
    pub frame_info: &'a [u8],
    /// Tagged records immediately following the image; ownership is unresolved.
    pub following_references: Vec<RuntimeRelocation>,
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct RuntimeModuleCache<'a> {
    pub header_words: [u64; 4],
    pub template_marker: u64,
    pub template_data: Option<[u8; 16]>,
    pub root_metadata: [u8; 40],
    /// Recorded odd companion-stream index; no stream-name guessing is used.
    pub companion_stream: u16,
    pub table_words: [u16; 3],
    pub auxiliary_links: Vec<RuntimeCacheLink>,
    pub class_binding: Option<RuntimeClassBinding<'a>>,
    pub methods: Vec<RuntimeMethodImage<'a>>,
    /// Through method images/relocations; further module structures may follow.
    pub consumed: usize,
}

fn invalid(r: &Reader<'_>) -> DecodeError {
    r.error(DecodeErrorKind::InvalidRuntimeModule)
}
fn dword(r: &mut Reader<'_>) -> Result<u32, DecodeError> {
    Ok(u32::from_le_bytes(r.take()?))
}
fn qword(r: &mut Reader<'_>) -> Result<u64, DecodeError> {
    Ok(u64::from_le_bytes(r.take()?))
}
fn peek_dword(r: &Reader<'_>) -> Option<u32> {
    Some(u32::from_le_bytes(
        r.bytes.get(r.at..)?.get(..4)?.try_into().ok()?,
    ))
}
fn peek_qword(r: &Reader<'_>) -> Option<u64> {
    Some(u64::from_le_bytes(
        r.bytes.get(r.at..)?.get(..8)?.try_into().ok()?,
    ))
}
fn bound_count(r: &Reader<'_>, count: usize, width: usize) -> Result<(), DecodeError> {
    let remaining = r.bytes.len() - r.at;
    if count > remaining / width {
        return Err(r.error(DecodeErrorKind::Truncated {
            needed: count.saturating_mul(width),
            remaining,
        }));
    }
    Ok(())
}
fn link(r: &mut Reader<'_>) -> Result<RuntimeCacheLink, DecodeError> {
    Ok(RuntimeCacheLink {
        key: qword(r)?,
        context: dword(r)?,
    })
}
fn nullable32(r: &mut Reader<'_>) -> Result<Option<RuntimeCacheLink>, DecodeError> {
    if peek_dword(r) == Some(u32::MAX) {
        dword(r)?;
        Ok(None)
    } else {
        link(r).map(Some)
    }
}
fn nullable64(r: &mut Reader<'_>) -> Result<Option<RuntimeCacheLink>, DecodeError> {
    let key = qword(r)?;
    if key == u64::MAX {
        Ok(None)
    } else {
        Ok(Some(RuntimeCacheLink {
            key,
            context: dword(r)?,
        }))
    }
}
fn links(r: &mut Reader<'_>, count: usize) -> Result<Vec<RuntimeCacheLink>, DecodeError> {
    bound_count(r, count, 12)?;
    (0..count).map(|_| link(r)).collect()
}
fn class_binding<'a>(r: &mut Reader<'a>) -> Result<RuntimeClassBinding<'a>, DecodeError> {
    let version = r.take()?;
    if version != [1, 0, 1, 0, 0, 0, 1, 0] {
        return Err(invalid(r));
    }
    let mut type_links = Vec::new();
    while let Some(value) = peek_dword(r) {
        if value == u32::MAX || value & 1 != 0 {
            type_links.push(nullable32(r)?);
        } else {
            break;
        }
    }
    let storage_header = r.take::<8>()?;
    let capacity = u16::from_le_bytes([storage_header[2], storage_header[3]]);
    let mut shared_links = [None; 6];
    for value in &mut shared_links {
        *value = nullable32(r)?;
    }
    let (slots, directory_header) = if capacity == 2 {
        let slots = RuntimeBindingSlots::Class([nullable64(r)?, nullable64(r)?]);
        let header = [r.word()?, r.word()?];
        if header[0] != 24 {
            return Err(invalid(r));
        }
        (slots, Some(header))
    } else {
        let count = usize::from(r.word()?);
        if count > usize::from(capacity) {
            return Err(invalid(r));
        }
        bound_count(r, count, 20)?;
        let mut events = Vec::with_capacity(count);
        for _ in 0..count {
            events.push(RuntimeHostEvent {
                dispatch_id: dword(r)?,
                member_id: dword(r)?,
                target: RuntimeCacheLink { key: 0, context: 0 },
            });
        }
        for event in &mut events {
            event.target = link(r)?;
        }
        let unused = r.slice((usize::from(capacity) - count) * 8)?;
        if unused.iter().any(|value| *value != 255) || dword(r)? != u32::MAX {
            return Err(invalid(r));
        }
        (
            RuntimeBindingSlots::Host {
                capacity,
                events,
                unused,
            },
            None,
        )
    };
    let count = usize::from(r.word()?);
    let directory_word = r.word()?;
    let invocation_links = links(r, count)?;
    Ok(RuntimeClassBinding {
        version,
        type_links,
        storage_header,
        shared_links,
        slots,
        directory_header,
        directory_word,
        invocation_links,
    })
}
fn record<'a>(r: &mut Reader<'a>, context: CacheContext) -> Result<RuntimeRecord<'a>, DecodeError> {
    let start = r.at;
    let decoded = decode_runtime_record(&r.bytes[start..], context).map_err(|mut error| {
        error.offset += start;
        error
    })?;
    r.slice(decoded.consumed)?;
    Ok(decoded.record)
}

/// Read a caller-selected B5/Win64 runtime module stream through its method
/// images. Module identity, heap-link resolution, activation and execution are
/// separate; the reader never searches for names or instruction patterns.
/// Eight lookahead bytes from the following section are required but remain
/// unconsumed, so a truncated final reference is not accepted as end-of-table.
pub fn decode_runtime_module_cache(
    bytes: &[u8],
    context: CacheContext,
) -> Result<RuntimeModuleCache<'_>, DecodeError> {
    context.validate()?;
    let mut r = Reader::new(bytes, 0);
    if r.word()? != 0x5572 {
        return Err(invalid(&r));
    }
    let header_words = [
        qword(&mut r)?,
        qword(&mut r)?,
        qword(&mut r)?,
        qword(&mut r)?,
    ];
    let template_marker = qword(&mut r)?;
    let template_data = match template_marker {
        0 => None,
        0x7e00000000000004 => Some(r.take()?),
        _ => return Err(invalid(&r)),
    };
    let root_metadata = r.take::<40>()?;
    let companion_stream = u16::from_le_bytes([root_metadata[30], root_metadata[31]]);
    if companion_stream & 1 == 0 {
        return Err(invalid(&r));
    }
    let first = r.word()?;
    let second = r.word()?;
    let count = usize::from(r.word()?);
    let third = r.word()?;
    let auxiliary_count = dword(&mut r)? as usize;
    let auxiliary_links = links(&mut r, auxiliary_count)?;
    let class_binding = if template_data.is_none() {
        Some(class_binding(&mut r)?)
    } else {
        None
    };
    bound_count(&r, count, 16)?;
    let mut methods = Vec::with_capacity(count);
    for _ in 0..count {
        let RuntimeRecord::Image { code, frame_info } = record(&mut r, context)? else {
            return Err(invalid(&r));
        };
        let mut relocations = Vec::new();
        while peek_qword(&r).is_some_and(|tag| tag & 1 != 0) {
            let RuntimeRecord::Reference { handle, offset } = record(&mut r, context)? else {
                return Err(invalid(&r));
            };
            relocations.push(RuntimeRelocation { handle, offset });
        }
        methods.push(RuntimeMethodImage {
            code,
            frame_info,
            following_references: relocations,
        });
    }
    bound_count(&r, 1, 8)?;
    Ok(RuntimeModuleCache {
        header_words,
        template_marker,
        template_data,
        root_metadata,
        companion_stream,
        table_words: [first, second, third],
        auxiliary_links,
        class_binding,
        methods,
        consumed: r.at,
    })
}
