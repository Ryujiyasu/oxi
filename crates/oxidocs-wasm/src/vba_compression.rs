// SPDX-License-Identifier: MIT OR Apache-2.0

//! Bounds-checked MS-OVBA compressed stream decoding.

use std::path::Path;

pub(crate) fn read_stream(
    project: &ovba::Project,
    path: impl AsRef<Path>,
    offset: usize,
) -> Result<Vec<u8>, String> {
    let bytes = project
        .read_stream(path)
        .map_err(|error| error.to_string())?;
    let compressed = bytes
        .get(offset..)
        .ok_or("VBA source offset is outside the module stream")?;
    decompress(compressed)
}

pub(crate) fn decompress(input: &[u8]) -> Result<Vec<u8>, String> {
    if input.first() != Some(&1) {
        return Err("invalid VBA compressed stream signature".into());
    }
    let mut at = 1usize;
    let mut output = Vec::new();
    while at < input.len() {
        let header_bytes = input
            .get(at..at + 2)
            .ok_or("truncated VBA compressed chunk header")?;
        let header = u16::from_le_bytes(header_bytes.try_into().unwrap());
        if (header >> 12) & 7 != 3 {
            return Err("invalid VBA compressed chunk signature".into());
        }
        let size = usize::from(header & 0x0fff) + 3;
        let end = at
            .checked_add(size)
            .filter(|end| *end <= input.len())
            .ok_or("truncated VBA compressed chunk")?;
        at += 2;
        let start = output.len();
        let limit = oxidocs_common::security::MAX_ENTRY_SIZE as usize;
        if header & 0x8000 == 0 {
            if size != 4098 {
                return Err("invalid VBA raw chunk size".into());
            }
            if start > limit.saturating_sub(4096) {
                return Err("decompressed VBA stream exceeds the size limit".into());
            }
            output.extend_from_slice(&input[at..end]);
            at = end;
            continue;
        }
        while at < end {
            let flags = input[at];
            at += 1;
            for bit in 0..8 {
                if at == end {
                    break;
                }
                if flags & (1 << bit) == 0 {
                    if output.len() - start == 4096 {
                        return Err("VBA compressed chunk expands beyond 4096 bytes".into());
                    }
                    if output.len() == limit {
                        return Err("decompressed VBA stream exceeds the size limit".into());
                    }
                    output.push(input[at]);
                    at += 1;
                } else {
                    let token_bytes = input
                        .get(at..at + 2)
                        .filter(|_| at + 2 <= end)
                        .ok_or("truncated VBA copy token")?;
                    let token = u16::from_le_bytes(token_bytes.try_into().unwrap());
                    at += 2;
                    let available = output.len() - start;
                    if available == 0 {
                        return Err("VBA copy token has no source bytes".into());
                    }
                    let offset_bits = if available <= 16 {
                        4
                    } else {
                        usize::BITS - (available - 1).leading_zeros()
                    };
                    let length_bits = 16 - offset_bits;
                    let offset = usize::from(token >> length_bits) + 1;
                    let length = usize::from(token & ((1u16 << length_bits) - 1)) + 3;
                    if offset > available {
                        return Err("VBA copy token points before the chunk".into());
                    }
                    if available + length > 4096 {
                        return Err("VBA compressed chunk expands beyond 4096 bytes".into());
                    }
                    if length > limit - output.len() {
                        return Err("decompressed VBA stream exceeds the size limit".into());
                    }
                    for _ in 0..length {
                        output.push(output[output.len() - offset]);
                    }
                }
            }
        }
    }
    Ok(output)
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn incomplete_chunks_and_invalid_back_references_return_errors() {
        for input in [
            vec![],
            vec![1, 0],
            vec![1, 0xff, 0xbf, 0],
            vec![1, 2, 0xb0, 1, 0, 0],
            vec![1, 1, 0xb0, 1, 0],
        ] {
            assert!(decompress(&input).is_err(), "{input:?}");
        }
        // Literal a, then an overlapping copy with offset1 and length3.
        assert_eq!(decompress(&[1, 3, 0xb0, 2, b'a', 0, 0]).unwrap(), b"aaaa");
        // The same legal overlap must not expand a chunk beyond4K.
        assert!(decompress(&[1, 3, 0xb0, 2, b'a', 0xff, 0x0f]).is_err());
    }

    #[test]
    fn native_fixture_streams_keep_their_original_decoded_bytes() {
        let package = include_bytes!("../../../tests/fixtures/vba/save_calculation.xlsm");
        let mut archive = oxidocs_common::archive::OoxmlArchive::new(package).unwrap();
        let project =
            ovba::open_project(archive.read_binary_part("xl/vbaProject.bin").unwrap()).unwrap();
        for module in &project.modules {
            let path = format!("/VBA/{}", module.stream_name);
            assert_eq!(
                read_stream(&project, &path, module.text_offset).unwrap(),
                project
                    .decompress_stream_from(&path, module.text_offset)
                    .unwrap()
            );
        }
    }

    #[test]
    fn copy_offsets_change_width_at_chunk_positions_and_reset_between_chunks() {
        for literal_count in [1usize, 16, 17, 32, 33, 256, 257, 2048, 2049] {
            let mut data = Vec::new();
            let literals: Vec<u8> = (0..literal_count).map(|index| index as u8).collect();
            for group in literals.chunks(8) {
                data.push(0);
                data.extend_from_slice(group);
            }
            let offset_bits = (usize::BITS - (literal_count - 1).leading_zeros()).max(4);
            let token = ((literal_count - 1) as u16) << (16 - offset_bits);
            if literal_count % 8 == 0 {
                data.push(1);
            } else {
                let last_flag = data.len() - literal_count % 8 - 1;
                data[last_flag] = 1 << (literal_count % 8);
            }
            data.extend_from_slice(&token.to_le_bytes());
            let header = 0xb000 | (data.len() as u16 - 1);
            let mut input = vec![1, 0xff, 0x3f];
            input.extend_from_slice(&[b'X'; 4096]);
            input.extend_from_slice(&header.to_le_bytes());
            input.extend_from_slice(&data);
            let mut expected = vec![b'X'; 4096];
            expected.extend_from_slice(&literals);
            for index in 0..3 {
                expected.push(literals[index % literal_count]);
            }
            assert_eq!(decompress(&input).unwrap(), expected, "{literal_count}");
        }
        // A previous chunk cannot supply the first copy token in a new chunk.
        let mut input = vec![1, 0xff, 0x3f];
        input.extend_from_slice(&[b'X'; 4096]);
        input.extend_from_slice(&[2, 0xb0, 1, 0, 0]);
        assert!(decompress(&input).unwrap_err().contains("no source bytes"));
    }
}
