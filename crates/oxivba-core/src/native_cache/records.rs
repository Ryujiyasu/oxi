// This Source Code Form is subject to the terms of the Mozilla Public
// License, v. 2.0. If a copy of the MPL was not distributed with this
// file, You can obtain one at https://mozilla.org/MPL/2.0/.

//! Framed runtime records from a caller-located B5/Win64 record boundary.

use super::{CacheContext, DecodeError, DecodeErrorKind, Reader};

#[derive(Debug, Clone, PartialEq, Eq)]
pub enum RuntimeRecord<'a> {
    /// Instruction bytes and opaque frame data. Neither is executed here.
    Image {
        code: &'a [u8],
        frame_info: &'a [u8],
    },
    /// Tagged cache-object handle and relocation offset, not a native address.
    /// Resolution and validation against the owning object remain separate.
    Reference { handle: u64, offset: u32 },
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct DecodedRuntimeRecord<'a> {
    pub record: RuntimeRecord<'a>,
    pub consumed: usize,
}

/// Read one tagged record, preserving references between runtime images.
/// Even leading words give an image's combined code/frame-data byte length;
/// object handles use the low-nibble tag 1. No payload search, stream-number
/// assumption, procedure-slot binding or frame-pointer resolution is performed.
pub fn decode_runtime_record(
    bytes: &[u8],
    context: CacheContext,
) -> Result<DecodedRuntimeRecord<'_>, DecodeError> {
    context.validate()?;
    let mut reader = Reader::new(bytes, 0);
    let tag = u64::from_le_bytes(reader.take()?);
    let record = if tag & 1 != 0 {
        if tag & 15 != 1 {
            return Err(reader.error(DecodeErrorKind::InvalidRuntimeRecordTag(tag)));
        }
        RuntimeRecord::Reference {
            handle: tag,
            offset: u32::from_le_bytes(reader.take()?),
        }
    } else {
        let code_size = u64::from_le_bytes(reader.take()?);
        if code_size & 1 != 0 || code_size > tag {
            return Err(reader.error(DecodeErrorKind::InvalidRuntimeImageSizes {
                image_size: tag,
                code_size,
            }));
        }
        let image_size = usize::try_from(tag).map_err(|_| {
            reader.error(DecodeErrorKind::InvalidRuntimeImageSizes {
                image_size: tag,
                code_size,
            })
        })?;
        // Bound the complete image before converting or slicing its code size.
        let image = reader.slice(image_size)?;
        let (code, frame_info) = image.split_at(code_size as usize);
        RuntimeRecord::Image { code, frame_info }
    };
    Ok(DecodedRuntimeRecord {
        record,
        consumed: reader.at,
    })
}
