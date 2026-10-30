// This Source Code Form is subject to the terms of the Mozilla Public
// License, v. 2.0. If a copy of the MPL was not distributed with this
// file, You can obtain one at https://mozilla.org/MPL/2.0/.

//! B5 Win64 module-object arenas and callable links. This is not a project
//! loader: companion names/signatures and activation remain separate inputs.

use super::{
    decode_runtime_module_cache, CacheContext, DecodeError, DecodeErrorKind, Reader,
    RuntimeCacheLink, RuntimeMethodImage, RuntimeModuleCache,
};

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct RuntimeHeapSpan {
    /// Logical cache address, independent of the serialized byte position.
    pub address: u64,
    pub allocation_units: u64,
    pub zeroed: bool,
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub enum RuntimeHeapObjectKind<'a> {
    StandardMethod {
        code_slot: u16,
        signature: RuntimeCacheLink,
    },
    MemberMethod {
        /// Includes class/host flags; its interpretation is left to the caller.
        identity: u32,
        signature: RuntimeCacheLink,
        secondary: Option<RuntimeCacheLink>,
    },
    Field {
        storage_offset: u16,
        type_word: u32,
    },
    Signature {
        header: [u8; 4],
        return_type: u8,
        parameter_types: &'a [u8],
    },
    Zeroed,
    Vacant,
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct RuntimeHeapObject<'a> {
    pub address: u64,
    pub allocation_units: u64,
    pub serialized_offset: usize,
    pub kind: RuntimeHeapObjectKind<'a>,
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct RuntimeModuleHeap<'a> {
    pub module: RuntimeModuleCache<'a>,
    pub zero_arena: Vec<RuntimeHeapSpan>,
    pub root_allocation_units: u64,
    /// Null image markers after the prefix; no code-slot ownership is inferred.
    pub null_image_markers: usize,
    pub objects: Vec<RuntimeHeapObject<'a>>,
    pub consumed: usize,
}

impl<'a> RuntimeModuleHeap<'a> {
    /// Resolves only the observed undisplaced, tagged links in this module's
    /// stream. The companion's exposed name and argument binding are separate.
    pub fn object_for_link(
        &self,
        module_stream: u16,
        link: RuntimeCacheLink,
    ) -> Option<&RuntimeHeapObject<'a>> {
        if link.context != u32::from(module_stream) << 16 || link.key & 15 != 1 {
            return None;
        }
        let address = link.key - 1;
        let mut matches = self.objects.iter().filter(|o| o.address == address);
        let object = matches.next()?;
        matches.next().is_none().then_some(object)
    }

    /// Selects an available standard-module image through its cached callable
    /// descriptor, not through displayed source names or declaration order.
    pub fn standard_method_for_link(
        &self,
        module_stream: u16,
        link: RuntimeCacheLink,
    ) -> Option<&RuntimeMethodImage<'a>> {
        let RuntimeHeapObjectKind::StandardMethod { code_slot, .. } =
            self.object_for_link(module_stream, link)?.kind
        else {
            return None;
        };
        self.module.methods.get(usize::from(code_slot))
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
fn limit(r: &Reader<'_>, units: u64) -> Result<u64, DecodeError> {
    units
        .checked_mul(16)
        .and_then(|v| v.checked_add(16))
        .ok_or_else(|| invalid(r))
}
fn advance(r: &Reader<'_>, address: u64, units: u64, end: u64) -> Result<u64, DecodeError> {
    if units == 0 || units & 1 != 0 {
        return Err(invalid(r));
    }
    let next = units
        .checked_mul(8)
        .and_then(|v| v.checked_add(16))
        .and_then(|v| v.checked_add(address))
        .filter(|v| *v <= end)
        .ok_or_else(|| invalid(r))?;
    Ok(next)
}

/// Decode the complete two-arena module stream in the covered cache context.
/// Allocation sizes determine logical keys; serialized payloads have distinct
/// lengths. Unknown object kinds are rejected rather than searched past.
pub fn decode_runtime_module_heap(
    bytes: &[u8],
    context: CacheContext,
) -> Result<RuntimeModuleHeap<'_>, DecodeError> {
    let module = decode_runtime_module_cache(bytes, context)?;
    let mut r = Reader::new(bytes, 0);
    r.at = 34; // signature and four words, already checked by the module reader
    let zero_end = limit(&r, module.header_words[1])?;
    let object_end = limit(&r, module.header_words[0])?;
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
    if address != zero_end {
        return Err(invalid(&r));
    }
    let root_units = qword(&mut r)?;
    if root_units >> 56 != 0 || r.at > module.consumed {
        return Err(invalid(&r));
    }
    address = advance(&r, 16, root_units, object_end)?;
    r.at = module.consumed;
    let mut null_image_markers = 0;
    while bytes.get(r.at..).and_then(|b| b.get(..8)) == Some(&[255; 8]) {
        qword(&mut r)?;
        null_image_markers += 1;
    }
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
            0x12 => RuntimeHeapObjectKind::StandardMethod {
                code_slot: r.word()?,
                signature: link(&mut r)?,
            },
            0x13 => {
                let identity = dword(&mut r)?;
                let signature = link(&mut r)?;
                let key = qword(&mut r)?;
                let secondary = if key == u64::MAX {
                    None
                } else {
                    Some(RuntimeCacheLink {
                        key,
                        context: dword(&mut r)?,
                    })
                };
                RuntimeHeapObjectKind::MemberMethod {
                    identity,
                    signature,
                    secondary,
                }
            }
            0x14 => RuntimeHeapObjectKind::Field {
                storage_offset: r.word()?,
                type_word: dword(&mut r)?,
            },
            0x0e => {
                let header: [u8; 4] = r.take()?;
                if header[0] != 0 {
                    return Err(invalid(&r));
                }
                let return_type = r.take::<1>()?[0];
                let parameter_types = r.slice(usize::from(header[1]))?;
                RuntimeHeapObjectKind::Signature {
                    header,
                    return_type,
                    parameter_types,
                }
            }
            0x7e => RuntimeHeapObjectKind::Zeroed,
            0x7f => RuntimeHeapObjectKind::Vacant,
            kind => return Err(r.error(DecodeErrorKind::UnsupportedRuntimeHeapObject(kind))),
        };
        objects.push(RuntimeHeapObject {
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
    Ok(RuntimeModuleHeap {
        module,
        zero_arena,
        root_allocation_units: root_units,
        null_image_markers,
        objects,
        consumed: r.at,
    })
}
