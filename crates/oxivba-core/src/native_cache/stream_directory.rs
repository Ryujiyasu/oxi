// This Source Code Form is subject to the terms of the Mozilla Public
// License, v. 2.0. If a copy of the MPL was not distributed with this
// file, You can obtain one at https://mozilla.org/MPL/2.0/.

//! Runtime stream ownership from the beginning of a B5/Win64 `__SRP_0`.

use super::{CacheContext, DecodeError, DecodeErrorKind, Reader};

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum RuntimeStreamKind {
    Primary,
    Companion,
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub enum RuntimeStreamSlot {
    Vacant,
    Occupied {
        kind: RuntimeStreamKind,
        /// None denotes project-wide data, not a module at index zero.
        module_index: Option<u16>,
        /// Remaining wire fields are retained without assigning load-state semantics.
        metadata: [u8; 6],
    },
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct RuntimeStreamDirectory {
    pub context: CacheContext,
    pub format_marker: u8,
    /// Slot index is the runtime stream index; vacant slots keep their positions.
    pub slots: Vec<RuntimeStreamSlot>,
    pub heap_flags: u16,
    pub root_heap_offset: usize,
    /// Root heap signature is checked but remains unconsumed.
    pub consumed: usize,
}

impl RuntimeStreamDirectory {
    /// Return a unique match, leaving absent or ambiguous ownership unresolved.
    /// Primary and companion streams are looked up independently, not inferred
    /// from slot adjacency, module order or the presence of a physical stream.
    pub fn unique_stream(&self, module_index: Option<u16>, kind: RuntimeStreamKind) -> Option<u16> {
        let mut found = None;
        for (index, slot) in self.slots.iter().enumerate() {
            if let RuntimeStreamSlot::Occupied {
                kind: actual,
                module_index: owner,
                ..
            } = slot
            {
                if *actual == kind && *owner == module_index {
                    if found.is_some() {
                        return None;
                    }
                    found = Some(u16::try_from(index).ok()?);
                }
            }
        }
        found
    }
}

fn invalid(r: &Reader<'_>) -> DecodeError {
    r.error(DecodeErrorKind::InvalidRuntimeDirectory)
}

/// Decode the stream table by its declared slot count. This does not open a
/// compound file, parse the root heap, recover names or activate cached code.
/// Module numbers refer to cached project metadata, not displayed source text.
pub fn decode_runtime_stream_directory(
    bytes: &[u8],
    context: CacheContext,
) -> Result<RuntimeStreamDirectory, DecodeError> {
    context.validate()?;
    let mut r = Reader::new(bytes, 0);
    if r.word()? != 0x4b93 {
        return Err(invalid(&r));
    }
    let format_marker = r.take::<1>()?[0];
    let version = r.take::<1>()?[0];
    let platform = r.word()?;
    let actual = CacheContext {
        version: u16::from(version),
        sys_kind: u32::from(platform),
    };
    actual.validate()?;
    if actual != context || format_marker != 0x2a {
        return Err(invalid(&r));
    }
    let count = usize::from(r.word()?);
    if count < 2 {
        return Err(invalid(&r));
    }
    let remaining = bytes.len() - r.at;
    if count > remaining / 2 {
        return Err(r.error(DecodeErrorKind::Truncated {
            needed: count * 2,
            remaining,
        }));
    }
    let mut slots = Vec::with_capacity(count);
    for _ in 0..count {
        let tag = r.word()?;
        let kind = match tag {
            0 => RuntimeStreamKind::Primary,
            2 => RuntimeStreamKind::Companion,
            5 => {
                slots.push(RuntimeStreamSlot::Vacant);
                continue;
            }
            _ => return Err(r.error(DecodeErrorKind::UnsupportedRuntimeStreamKind(tag))),
        };
        let owner = r.word()?;
        slots.push(RuntimeStreamSlot::Occupied {
            kind,
            module_index: (owner != u16::MAX).then_some(owner),
            metadata: r.take()?,
        });
    }
    if !matches!(
        slots[0],
        RuntimeStreamSlot::Occupied {
            kind: RuntimeStreamKind::Primary,
            module_index: None,
            ..
        }
    ) || !matches!(
        slots[1],
        RuntimeStreamSlot::Occupied {
            kind: RuntimeStreamKind::Companion,
            module_index: None,
            ..
        }
    ) {
        return Err(invalid(&r));
    }
    let heap_flags = r.word()?;
    let root_heap_offset = r.at;
    // Require complete lookahead without consuming the following root heap.
    let marker = r.slice(2)?;
    if marker != [0x72, 0x55] {
        return Err(invalid(&r));
    }
    Ok(RuntimeStreamDirectory {
        context: actual,
        format_marker,
        slots,
        heap_flags,
        root_heap_offset,
        consumed: root_heap_offset,
    })
}
