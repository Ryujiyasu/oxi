// This Source Code Form is subject to the terms of the Mozilla Public
// License, v. 2.0. If a copy of the MPL was not distributed with this
// file, You can obtain one at https://mozilla.org/MPL/2.0/.

//! Paired module type information. Serialized method order supplies invocation
//! entries; member IDs and primary target code slots are independent fields.

use super::{
    CacheContext, DecodeError, DecodeErrorKind, Reader, RuntimeCacheLink, RuntimeHeapObjectKind,
    RuntimeHeapSpan, RuntimeModuleHeap, RuntimeProjectHeap, RuntimeProjectModuleState,
    RuntimeStreamKind,
};

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct RuntimeCompanionType {
    pub word: u16,
    pub array_word: Option<u16>,
    pub reference: Option<RuntimeCacheLink>,
}
impl RuntimeCompanionType {
    pub fn base(&self) -> u16 {
        self.word & 0x3f
    }
    pub fn is_by_ref(&self) -> bool {
        self.word & 0x40 != 0
    }
    pub fn is_array(&self) -> bool {
        self.word & 0x80 != 0
    }
    pub fn is_optional(&self) -> bool {
        self.word & 0x100 != 0
    }
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct RuntimeCompanionMethod<'a> {
    pub serialized_offset: usize,
    pub native_size: u16,
    pub target: RuntimeCacheLink,
    pub member_id: u32,
    pub flags: u16,
    pub location_word: u32,
    pub extra_links: [Option<RuntimeCacheLink>; 2],
    pub help_word: u32,
    pub parameter_links: Vec<Option<RuntimeCacheLink>>,
    /// Encoded variant/default data, including any native padding. Not source literals.
    pub default_data: &'a [u8],
    pub return_type: RuntimeCompanionType,
    /// Includes hidden COM parameters for class/host methods.
    pub parameter_types: Vec<RuntimeCompanionType>,
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct RuntimeCompanionField {
    pub native_size: u16,
    pub target: RuntimeCacheLink,
    pub extra: Option<RuntimeCacheLink>,
    pub help_word: u32,
    pub member_id: u32,
    pub flags: u16,
    pub location_word: u16,
    pub value_type: RuntimeCompanionType,
    pub storage_offset: u32,
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct RuntimePositionalArity {
    pub required: usize,
    pub maximum: Option<usize>,
}
impl RuntimePositionalArity {
    pub fn accepts(self, count: usize) -> bool {
        count >= self.required && self.maximum.is_none_or(|max| count <= max)
    }
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct RuntimeCompanionCache<'a> {
    pub header_words: [u64; 4],
    pub zero_arena: Vec<RuntimeHeapSpan>,
    pub owner: RuntimeCacheLink,
    pub root_raw: [u8; 22],
    pub primary_stream: u16,
    pub companion_stream: u16,
    pub standard_module: bool,
    pub methods: Vec<RuntimeCompanionMethod<'a>>,
    pub fields: Vec<Option<RuntimeCompanionField>>,
    pub tail_word: u16,
    pub type_links: Vec<RuntimeCacheLink>,
    pub object_arena: Vec<RuntimeHeapSpan>,
    pub consumed: usize,
}

impl RuntimeCompanionCache<'_> {
    /// Positional argument counts for covered standard-module invocations.
    /// Class/host signatures contain hidden parameters and require an adapter.
    /// This does not perform argument conversion, default decoding or execution.
    /// Host API argument limits remain separate from this signature-level arity.
    pub fn positional_arity(&self, entry: usize) -> Option<RuntimePositionalArity> {
        if !self.standard_module {
            return None;
        }
        let method = self.methods.get(entry)?;
        let variadic = method.flags & 0xfc00 == 0xfc00;
        if variadic
            && !method
                .parameter_types
                .last()
                .is_some_and(|t| t.is_array() && t.is_by_ref() && t.base() == 15)
        {
            return None;
        }
        let fixed = method
            .parameter_types
            .len()
            .checked_sub(usize::from(variadic))?;
        let required = method.parameter_types[..fixed]
            .iter()
            .rposition(|t| !t.is_optional())
            .map_or(0, |i| i + 1);
        Some(RuntimePositionalArity {
            required,
            maximum: if variadic { None } else { Some(fixed) },
        })
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
fn value_type(r: &mut Reader<'_>) -> Result<RuntimeCompanionType, DecodeError> {
    let word = r.word()?;
    let array_word = if word & 0x80 != 0 {
        Some(r.word()?)
    } else {
        None
    };
    let reference = if matches!(word & 0x3f, 20 | 30) {
        Some(link(r)?)
    } else {
        None
    };
    Ok(RuntimeCompanionType {
        word,
        array_word,
        reference,
    })
}
fn limit(r: &Reader<'_>, units: u64) -> Result<u64, DecodeError> {
    units
        .checked_mul(16)
        .and_then(|v| v.checked_add(16))
        .ok_or_else(|| invalid(r))
}
fn arena(
    r: &mut Reader<'_>,
    mut address: u64,
    end: u64,
) -> Result<Vec<RuntimeHeapSpan>, DecodeError> {
    let mut spans = Vec::new();
    loop {
        let tag = qword(r)?;
        if tag == 0 {
            break;
        }
        let kind = (tag >> 56) as u8;
        let units = tag & 0x00ff_ffff_ffff_ffff;
        if !matches!(kind, 0x7e | 0x7f) || units == 0 || units & 1 != 0 {
            return Err(invalid(r));
        }
        let next = units
            .checked_mul(8)
            .and_then(|v| v.checked_add(16))
            .and_then(|v| v.checked_add(address))
            .filter(|v| *v <= end)
            .ok_or_else(|| invalid(r))?;
        spans.push(RuntimeHeapSpan {
            address,
            allocation_units: units,
            zeroed: kind == 0x7e,
        });
        address = next;
    }
    if address != end {
        return Err(invalid(r));
    }
    Ok(spans)
}

/// Decode a module companion using its primary heap's callable descriptor count.
/// Method records have no standalone count/zero terminator; a field count follows
/// them. Source declaration counts and member-ID ordering are not used to frame it.
pub fn decode_runtime_companion_cache<'a>(
    bytes: &'a [u8],
    context: CacheContext,
    primary: &RuntimeModuleHeap<'_>,
    primary_stream: u16,
) -> Result<RuntimeCompanionCache<'a>, DecodeError> {
    context.validate()?;
    let mut r = Reader::new(bytes, 0);
    if r.word()? != 0x5572 {
        return Err(invalid(&r));
    }
    let mut header_words = [0; 4];
    for value in &mut header_words {
        *value = qword(&mut r)?;
    }
    let zero_end = limit(&r, header_words[1])?;
    let object_end = limit(&r, header_words[0])?;
    let zero_arena = arena(&mut r, 16, zero_end)?;
    if qword(&mut r)? != 26 {
        return Err(invalid(&r));
    }
    let owner = link(&mut r)?;
    if owner
        != (RuntimeCacheLink {
            key: 17,
            context: u32::from(primary_stream) << 16,
        })
    {
        return Err(invalid(&r));
    }
    let root_raw = r.take()?;
    let count = primary
        .objects
        .iter()
        .filter(|o| {
            matches!(
                o.kind,
                RuntimeHeapObjectKind::StandardMethod { .. }
                    | RuntimeHeapObjectKind::MemberMethod { .. }
            )
        })
        .count();
    if count > r.bytes.len().saturating_sub(r.at) / 50 {
        return Err(invalid(&r));
    }
    let mut methods = Vec::with_capacity(count);
    for _ in 0..count {
        let serialized_offset = r.at;
        let native_size = r.word()?;
        if native_size < 56 || native_size & 7 != 0 {
            return Err(invalid(&r));
        }
        let target = link(&mut r)?;
        let object = primary
            .object_for_link(primary_stream, target)
            .ok_or_else(|| invalid(&r))?;
        if !matches!(
            object.kind,
            RuntimeHeapObjectKind::StandardMethod { .. }
                | RuntimeHeapObjectKind::MemberMethod { .. }
        ) {
            return Err(invalid(&r));
        }
        let member_id = dword(&mut r)?;
        let flags = r.word()?;
        let location_word = dword(&mut r)?;
        let extra_links = [optional(&mut r)?, optional(&mut r)?];
        let help_word = dword(&mut r)?;
        let params = usize::from((flags & 0xfc) >> 2);
        if params > r.bytes.len().saturating_sub(r.at) / 8 {
            return Err(invalid(&r));
        }
        let mut parameter_links = Vec::with_capacity(params);
        for _ in 0..params {
            parameter_links.push(optional(&mut r)?);
        }
        let default_size = usize::try_from(dword(&mut r)?).map_err(|_| invalid(&r))?;
        let default_data = r.slice(default_size)?;
        let return_type = value_type(&mut r)?;
        let mut parameter_types = Vec::with_capacity(params);
        for _ in 0..params {
            parameter_types.push(value_type(&mut r)?);
        }
        methods.push(RuntimeCompanionMethod {
            serialized_offset,
            native_size,
            target,
            member_id,
            flags,
            location_word,
            extra_links,
            help_word,
            parameter_links,
            default_data,
            return_type,
            parameter_types,
        });
    }
    let count = usize::from(r.word()?);
    if count > r.bytes.len().saturating_sub(r.at) / 2 {
        return Err(invalid(&r));
    }
    let mut fields = Vec::with_capacity(count);
    for _ in 0..count {
        let native_size = r.word()?;
        if native_size == 0 {
            fields.push(None);
            continue;
        }
        if native_size < 42 || native_size & 1 != 0 {
            return Err(invalid(&r));
        }
        fields.push(Some(RuntimeCompanionField {
            native_size,
            target: link(&mut r)?,
            extra: optional(&mut r)?,
            help_word: dword(&mut r)?,
            member_id: dword(&mut r)?,
            flags: r.word()?,
            location_word: r.word()?,
            value_type: value_type(&mut r)?,
            storage_offset: dword(&mut r)?,
        }));
    }
    let tail_word = r.word()?;
    if tail_word != 0 {
        return Err(invalid(&r));
    }
    let count = usize::from(r.word()?);
    if count > r.bytes.len().saturating_sub(r.at) / 12 {
        return Err(invalid(&r));
    }
    let mut type_links = Vec::with_capacity(count);
    for _ in 0..count {
        type_links.push(link(&mut r)?);
    }
    let object_arena = arena(&mut r, 16 + 16 + 26 * 8, object_end)?;
    if r.at != bytes.len() {
        return Err(invalid(&r));
    }
    Ok(RuntimeCompanionCache {
        header_words,
        zero_arena,
        owner,
        root_raw,
        primary_stream,
        companion_stream: primary.module.companion_stream,
        standard_module: primary.module.class_binding.is_none(),
        methods,
        fields,
        tail_word,
        type_links,
        object_arena,
        consumed: r.at,
    })
}

/// Align canonical runtime name entries with companion signatures/targets using
/// the project stream directory and paired primary heap. Member IDs do not select
/// entries, and exposed signatures need not match the independently selected code.
pub fn decode_runtime_invocation_table<'a>(
    project: &RuntimeProjectHeap<'_>,
    module_index: usize,
    primary: &RuntimeModuleHeap<'_>,
    companion_bytes: &'a [u8],
    context: CacheContext,
) -> Result<RuntimeCompanionCache<'a>, DecodeError> {
    let r = Reader::new(companion_bytes, 0);
    let module = project
        .modules
        .get(module_index)
        .ok_or_else(|| invalid(&r))?;
    let RuntimeProjectModuleState::Ready(ready) = &module.state else {
        return Err(invalid(&r));
    };
    let index = u16::try_from(module_index).map_err(|_| invalid(&r))?;
    let stream = project
        .directory
        .unique_stream(Some(index), RuntimeStreamKind::Primary)
        .ok_or_else(|| invalid(&r))?;
    let companion = project
        .directory
        .unique_stream(Some(index), RuntimeStreamKind::Companion)
        .ok_or_else(|| invalid(&r))?;
    if ready.root
        != (RuntimeCacheLink {
            key: 17,
            context: u32::from(stream) << 16,
        })
        || companion != primary.module.companion_stream
    {
        return Err(invalid(&r));
    }
    let result = decode_runtime_companion_cache(companion_bytes, context, primary, stream)?;
    if result.methods.len() != ready.procedure_names.len() {
        return Err(invalid(&r));
    }
    Ok(result)
}
