// This Source Code Form is subject to the terms of the Mozilla Public
// License, v. 2.0. If a copy of the MPL was not distributed with this
// file, You can obtain one at https://mozilla.org/MPL/2.0/.

//! B5 Win64 execution-frame sizes and cleanup locations. Opaque native fields
//! are preserved, never dereferenced or treated as serialized host pointers.

use super::{CacheContext, DecodeError, DecodeErrorKind, Reader, RuntimeMethodImage};

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct RuntimeCleanupSlot {
    pub offset: i32,
    pub kind: u16,
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct RuntimeCleanupTable<'a> {
    pub offset: usize,
    pub header: &'a [u8],
    pub slots: Vec<RuntimeCleanupSlot>,
    pub consumed: usize,
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct RuntimeFrameLayout<'a> {
    pub cached_word: u64,
    /// Includes the measured eight-byte ABI prefix and hidden result arguments.
    pub argument_bytes: u16,
    pub local_bytes: u16,
    pub code_bytes: u32,
    pub raw_header: &'a [u8],
    pub first_cleanup: RuntimeCleanupTable<'a>,
    pub alignment: &'a [u8],
    pub second_cleanup: RuntimeCleanupTable<'a>,
    pub consumed: usize,
}

fn invalid(r: &Reader<'_>) -> DecodeError {
    r.error(DecodeErrorKind::InvalidRuntimeFrame)
}

fn cleanup<'a>(r: &mut Reader<'a>) -> Result<RuntimeCleanupTable<'a>, DecodeError> {
    let offset = r.at;
    let size = u32::from_le_bytes(r.take()?) as usize;
    if size < 20 || (size - 20) % 6 != 0 {
        return Err(invalid(r));
    }
    let payload = r.slice(size - 4)?;
    let count = usize::from(u16::from_le_bytes([payload[6], payload[7]]));
    if count != (size - 20) / 6 {
        return Err(invalid(r));
    }
    let slots = payload[16..]
        .chunks_exact(6)
        .map(|s| RuntimeCleanupSlot {
            offset: i32::from_le_bytes(s[..4].try_into().expect("four-byte offset")),
            kind: u16::from_le_bytes([s[4], s[5]]),
        })
        .collect();
    Ok(RuntimeCleanupTable {
        offset,
        header: &payload[..16],
        slots,
        consumed: size,
    })
}

/// Decode the complete frame-info portion of a method image. Cleanup-table
/// count/size/extent are checked; unknown cleanup kinds remain explicit data.
/// No procedure identity, parameter type or lifetime policy is inferred here.
pub fn decode_runtime_frame_layout<'a>(
    image: &RuntimeMethodImage<'a>,
    context: CacheContext,
) -> Result<RuntimeFrameLayout<'a>, DecodeError> {
    context.validate()?;
    let mut r = Reader::new(image.frame_info, 0);
    let cached_word = u64::from_le_bytes(r.take()?);
    let argument_bytes = r.word()?;
    let local_bytes = r.word()?;
    let code_bytes = u32::from_le_bytes(r.take()?);
    let second_offset = u32::from_le_bytes(r.take()?) as usize;
    let raw_header = r.slice(16)?;
    if code_bytes as usize != image.code.len() {
        return Err(invalid(&r));
    }
    let first_cleanup = cleanup(&mut r)?;
    let aligned = r.at.checked_add(3).ok_or_else(|| invalid(&r))? & !3;
    if second_offset != aligned {
        return Err(invalid(&r));
    }
    let alignment = r.slice(aligned - r.at)?;
    let second_cleanup = cleanup(&mut r)?;
    if r.at != image.frame_info.len() {
        return Err(invalid(&r));
    }
    Ok(RuntimeFrameLayout {
        cached_word,
        argument_bytes,
        local_bytes,
        code_bytes,
        raw_header,
        first_cleanup,
        alignment,
        second_cleanup,
        consumed: r.at,
    })
}
