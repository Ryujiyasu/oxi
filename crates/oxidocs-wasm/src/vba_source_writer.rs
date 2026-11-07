// SPDX-License-Identifier: MIT OR Apache-2.0

//! Update source in existing VBA modules while retaining project identities.

use crate::vba_module_directory::{identities, module_source, module_table_offset};
use std::collections::{BTreeMap, BTreeSet, HashMap};
use std::io::{Cursor, Write};

/// MS-OVBA 2.4.1: chunks decompress to 4096 bytes except the final chunk.
/// Raw fallback, including final zero padding, follows section 2.4.1.3.10.
pub(crate) fn compress(bytes: &[u8]) -> Vec<u8> {
    let mut output = vec![1];
    for chunk in bytes.chunks(4096) {
        let mut encoded = Vec::new();
        let mut positions: HashMap<[u8; 3], Vec<usize>> = HashMap::new();
        let mut at = 0usize;
        while at < chunk.len() && encoded.len() <= 4096 {
            let flag_at = encoded.len();
            encoded.push(0);
            for bit in 0..8 {
                if at == chunk.len() {
                    break;
                }
                let offset_bits = if at <= 16 {
                    4
                } else {
                    usize::BITS - (at - 1).leading_zeros()
                };
                let length_bits = 16 - offset_bits;
                let max_length = ((1usize << length_bits) + 2).min(chunk.len() - at);
                let mut best = (0usize, 0usize);
                if at + 3 <= chunk.len() {
                    let key: [u8; 3] = chunk[at..at + 3].try_into().unwrap();
                    if let Some(candidates) = positions.get(&key) {
                        // Bounded work for long comments and repetitive source.
                        for &previous in candidates.iter().rev().take(64) {
                            let mut length = 3;
                            while length < max_length
                                && chunk[previous + length] == chunk[at + length]
                            {
                                length += 1;
                            }
                            if length > best.1 {
                                best = (at - previous, length);
                            }
                            if length == max_length {
                                break;
                            }
                        }
                    }
                }
                let length = if best.1 >= 3 {
                    encoded[flag_at] |= 1 << bit;
                    let token = (((best.0 - 1) << length_bits) | (best.1 - 3)) as u16;
                    encoded.extend_from_slice(&token.to_le_bytes());
                    best.1
                } else {
                    encoded.push(chunk[at]);
                    1
                };
                for position in at..at + length {
                    if position + 3 <= chunk.len() {
                        positions
                            .entry(chunk[position..position + 3].try_into().unwrap())
                            .or_default()
                            .push(position);
                    }
                }
                at += length;
            }
        }
        if at == chunk.len() && encoded.len() <= 4096 {
            let header = 0xb000 | (encoded.len() as u16 - 1);
            output.extend_from_slice(&header.to_le_bytes());
            output.extend_from_slice(&encoded);
        } else {
            output.extend_from_slice(&0x3fffu16.to_le_bytes());
            output.extend_from_slice(chunk);
            output.resize(output.len() + 4096 - chunk.len(), 0);
        }
    }
    output
}

fn compress_source(bytes: &[u8]) -> Vec<u8> {
    // Keep literal blocks below the native stored-size limit. Source streams
    // must avoid raw blocks and zero padding, including incompressible text.
    let mut output = vec![1];
    for chunk in bytes.chunks(3600) {
        let mut body = Vec::new();
        for literal in chunk.chunks(8) {
            body.push(0);
            body.extend_from_slice(literal);
        }
        let size = body.len();
        output.extend_from_slice(&(0xb000 | (size as u16-1)).to_le_bytes());
        output.extend_from_slice(&body);
    }
    output
}

fn source_body(text: &str, class: bool) -> Result<String, String> {
    let text = text.replace("\r\n", "\n");
    let mut lines = text.lines();
    let mut first = lines.next().unwrap_or("");
    if class && first.trim().to_ascii_uppercase().starts_with("VERSION ") {
        first = lines.next().unwrap_or("");
        if first.trim().eq_ignore_ascii_case("BEGIN") {
            loop {
                let line = lines.next().ok_or("unterminated exported class header")?;
                if line.trim().eq_ignore_ascii_case("END") {
                    break;
                }
            }
            first = lines.next().unwrap_or("");
        }
    }
    Ok(std::iter::once(first).chain(lines).collect::<Vec<_>>().join("\n"))
}

fn comparable(text: &str) -> String {
    text.replace("\r\n", "\n").trim_end_matches(['\n', '\0']).to_owned()
}

fn module_record(output: &mut Vec<u8>, tag: u16, bytes: &[u8]) -> Result<(), String> {
    let length = u32::try_from(bytes.len()).map_err(|_| "VBA module record is too large")?;
    output.extend_from_slice(&tag.to_le_bytes());
    output.extend_from_slice(&length.to_le_bytes());
    output.extend_from_slice(bytes);
    Ok(())
}

fn unicode_name(name: &str) -> Vec<u8> {
    name.encode_utf16().flat_map(u16::to_le_bytes).collect()
}

const MODULE_NAME_BEST_FIT_1252: &[(u16, u8)] = &[
    (0x0081, 0x81), (0x008d, 0x8d), (0x008f, 0x8f), (0x0090, 0x90), (0x009d, 0x9d), (0x0100, 0x41),
    (0x0101, 0x61), (0x0102, 0x41), (0x0103, 0x61), (0x0104, 0x41), (0x0105, 0x61), (0x0106, 0x43),
    (0x0107, 0x63), (0x0108, 0x43), (0x0109, 0x63), (0x010a, 0x43), (0x010b, 0x63), (0x010c, 0x43),
    (0x010d, 0x63), (0x010e, 0x44), (0x010f, 0x64), (0x0110, 0xd0), (0x0111, 0x64), (0x0112, 0x45),
    (0x0113, 0x65), (0x0114, 0x45), (0x0115, 0x65), (0x0116, 0x45), (0x0117, 0x65), (0x0118, 0x45),
    (0x0119, 0x65), (0x011a, 0x45), (0x011b, 0x65), (0x011c, 0x47), (0x011d, 0x67), (0x011e, 0x47),
    (0x011f, 0x67), (0x0120, 0x47), (0x0121, 0x67), (0x0122, 0x47), (0x0123, 0x67), (0x0124, 0x48),
    (0x0125, 0x68), (0x0126, 0x48), (0x0127, 0x68), (0x0128, 0x49), (0x0129, 0x69), (0x012a, 0x49),
    (0x012b, 0x69), (0x012c, 0x49), (0x012d, 0x69), (0x012e, 0x49), (0x012f, 0x69), (0x0130, 0x49),
    (0x0131, 0x69), (0x0134, 0x4a), (0x0135, 0x6a), (0x0136, 0x4b), (0x0137, 0x6b), (0x0139, 0x4c),
    (0x013a, 0x6c), (0x013b, 0x4c), (0x013c, 0x6c), (0x013d, 0x4c), (0x013e, 0x6c), (0x0141, 0x4c),
    (0x0142, 0x6c), (0x0143, 0x4e), (0x0144, 0x6e), (0x0145, 0x4e), (0x0146, 0x6e), (0x0147, 0x4e),
    (0x0148, 0x6e), (0x014c, 0x4f), (0x014d, 0x6f), (0x014e, 0x4f), (0x014f, 0x6f), (0x0150, 0x4f),
    (0x0151, 0x6f), (0x0154, 0x52), (0x0155, 0x72), (0x0156, 0x52), (0x0157, 0x72), (0x0158, 0x52),
    (0x0159, 0x72), (0x015a, 0x53), (0x015b, 0x73), (0x015c, 0x53), (0x015d, 0x73), (0x015e, 0x53),
    (0x015f, 0x73), (0x0162, 0x54), (0x0163, 0x74), (0x0164, 0x54), (0x0165, 0x74), (0x0166, 0x54),
    (0x0167, 0x74), (0x0168, 0x55), (0x0169, 0x75), (0x016a, 0x55), (0x016b, 0x75), (0x016c, 0x55),
    (0x016d, 0x75), (0x016e, 0x55), (0x016f, 0x75), (0x0170, 0x55), (0x0171, 0x75), (0x0172, 0x55),
    (0x0173, 0x75), (0x0174, 0x57), (0x0175, 0x77), (0x0176, 0x59), (0x0177, 0x79), (0x0179, 0x5a),
    (0x017a, 0x7a), (0x017b, 0x5a), (0x017c, 0x7a), (0x0180, 0x62), (0x0189, 0xd0), (0x0191, 0x83),
    (0x0197, 0x49), (0x019a, 0x6c), (0x019f, 0x4f), (0x01a0, 0x4f), (0x01a1, 0x6f), (0x01ab, 0x74),
    (0x01ae, 0x54), (0x01af, 0x55), (0x01b0, 0x75), (0x01b6, 0x7a), (0x01c0, 0x7c), (0x01c3, 0x21),
    (0x01cd, 0x41), (0x01ce, 0x61), (0x01cf, 0x49), (0x01d0, 0x69), (0x01d1, 0x4f), (0x01d2, 0x6f),
    (0x01d3, 0x55), (0x01d4, 0x75), (0x01d5, 0x55), (0x01d6, 0x75), (0x01d7, 0x55), (0x01d8, 0x75),
    (0x01d9, 0x55), (0x01da, 0x75), (0x01db, 0x55), (0x01dc, 0x75), (0x01de, 0x41), (0x01df, 0x61),
    (0x01e4, 0x47), (0x01e5, 0x67), (0x01e6, 0x47), (0x01e7, 0x67), (0x01e8, 0x4b), (0x01e9, 0x6b),
    (0x01ea, 0x4f), (0x01eb, 0x6f), (0x01ec, 0x4f), (0x01ed, 0x6f), (0x01f0, 0x6a), (0x0261, 0x67),
    (0x02b9, 0x27), (0x02ba, 0x22), (0x02bc, 0x27), (0x02c4, 0x5e), (0x02c8, 0x27), (0x02c9, 0xaf),
    (0x02ca, 0xb4), (0x02cb, 0x60), (0x02cd, 0x5f), (0x02da, 0xb0), (0x0300, 0x60), (0x0301, 0xb4),
    (0x0302, 0x5e), (0x0303, 0x7e), (0x0304, 0xaf), (0x0305, 0xaf), (0x0308, 0xa8), (0x030a, 0xb0),
    (0x030e, 0x22), (0x0327, 0xb8), (0x0331, 0x5f), (0x0332, 0x5f), (0x037e, 0x3b), (0x0393, 0x47),
    (0x0398, 0x54), (0x03a3, 0x53), (0x03a6, 0x46), (0x03a9, 0x4f), (0x03b1, 0x61), (0x03b2, 0xdf),
    (0x03b4, 0x64), (0x03b5, 0x65), (0x03bc, 0xb5), (0x03c0, 0x70), (0x03c3, 0x73), (0x03c4, 0x74),
    (0x03c6, 0x66), (0x04bb, 0x68), (0x0589, 0x3a), (0x066a, 0x25), (0x2000, 0x20), (0x2001, 0x20),
    (0x2002, 0x20), (0x2003, 0x20), (0x2004, 0x20), (0x2005, 0x20), (0x2006, 0x20), (0x2007, 0xa0),
    (0x2008, 0x20), (0x2009, 0x20), (0x200a, 0x20), (0x2010, 0x2d), (0x2011, 0x2d), (0x2012, 0x2d),
    (0x2015, 0x2d), (0x2017, 0x3d), (0x201b, 0x27), (0x201f, 0x22), (0x2024, 0xb7), (0x202f, 0xa0),
    (0x2032, 0x27), (0x2033, 0x22), (0x2035, 0x60), (0x2036, 0x22), (0x2044, 0x2f), (0x2070, 0xb0),
    (0x2074, 0x34), (0x2075, 0x35), (0x2076, 0x36), (0x2077, 0x37), (0x2078, 0x38), (0x207f, 0x6e),
    (0x2080, 0x30), (0x2081, 0x31), (0x2082, 0x32), (0x2083, 0x33), (0x2084, 0x34), (0x2085, 0x35),
    (0x2086, 0x36), (0x2087, 0x37), (0x2088, 0x38), (0x2089, 0x39), (0x20a1, 0xa2), (0x20a4, 0xa3),
    (0x20a7, 0x50), (0x2102, 0x43), (0x2107, 0x45), (0x210a, 0x67), (0x210b, 0x48), (0x210c, 0x48),
    (0x210d, 0x48), (0x210e, 0x68), (0x2110, 0x49), (0x2111, 0x49), (0x2112, 0x4c), (0x2113, 0x6c),
    (0x2115, 0x4e), (0x2118, 0x50), (0x2119, 0x50), (0x211a, 0x51), (0x211b, 0x52), (0x211c, 0x52),
    (0x211d, 0x52), (0x2124, 0x5a), (0x2128, 0x5a), (0x212a, 0x4b), (0x212b, 0xc5), (0x212c, 0x42),
    (0x212d, 0x43), (0x212e, 0x65), (0x212f, 0x65), (0x2130, 0x45), (0x2131, 0x46), (0x2133, 0x4d),
    (0x2134, 0x6f), (0x2205, 0xd8), (0x2212, 0x2d), (0x2213, 0xb1), (0x2215, 0x2f), (0x2216, 0x5c),
    (0x2217, 0x2a), (0x2218, 0xb0), (0x2219, 0xb7), (0x221a, 0x76), (0x221e, 0x38), (0x2223, 0x7c),
    (0x2229, 0x6e), (0x2236, 0x3a), (0x223c, 0x7e), (0x2248, 0x98), (0x2261, 0x3d), (0x2264, 0x3d),
    (0x2265, 0x3d), (0x226a, 0xab), (0x226b, 0xbb), (0x22c5, 0xb7), (0x2302, 0xa6), (0x2303, 0x5e),
    (0x2310, 0xac), (0x2320, 0x28), (0x2321, 0x29), (0x2329, 0x3c), (0x232a, 0x3e), (0x2500, 0x2d),
    (0x2502, 0xa6), (0x250c, 0x2b), (0x2510, 0x2b), (0x2514, 0x2b), (0x2518, 0x2b), (0x251c, 0x2b),
    (0x2524, 0xa6), (0x252c, 0x2d), (0x2534, 0x2d), (0x253c, 0x2b), (0x2550, 0x2d), (0x2551, 0xa6),
    (0x2552, 0x2b), (0x2553, 0x2b), (0x2554, 0x2b), (0x2555, 0x2b), (0x2556, 0x2b), (0x2557, 0x2b),
    (0x2558, 0x2b), (0x2559, 0x2b), (0x255a, 0x2b), (0x255b, 0x2b), (0x255c, 0x2b), (0x255d, 0x2b),
    (0x255e, 0xa6), (0x255f, 0xa6), (0x2560, 0xa6), (0x2561, 0xa6), (0x2562, 0xa6), (0x2563, 0xa6),
    (0x2564, 0x2d), (0x2565, 0x2d), (0x2566, 0x2d), (0x2567, 0x2d), (0x2568, 0x2d), (0x2569, 0x2d),
    (0x256a, 0x2b), (0x256b, 0x2b), (0x256c, 0x2b), (0x2580, 0xaf), (0x2584, 0x5f), (0x2588, 0xa6),
    (0x258c, 0xa6), (0x2590, 0xa6), (0x2591, 0xa6), (0x2592, 0xa6), (0x2593, 0xa6), (0x25a0, 0xa6),
    (0x263c, 0xa4), (0x2758, 0x7c), (0x3000, 0x20), (0x3008, 0x3c), (0x3009, 0x3e), (0x300a, 0xab),
    (0x300b, 0xbb), (0x301a, 0x5b), (0x301b, 0x5d), (0x30fb, 0xb7), (0xff01, 0x21), (0xff02, 0x22),
    (0xff03, 0x23), (0xff04, 0x24), (0xff05, 0x25), (0xff06, 0x26), (0xff07, 0x27), (0xff08, 0x28),
    (0xff09, 0x29), (0xff0a, 0x2a), (0xff0b, 0x2b), (0xff0c, 0x2c), (0xff0d, 0x2d), (0xff0e, 0x2e),
    (0xff0f, 0x2f), (0xff10, 0x30), (0xff11, 0x31), (0xff12, 0x32), (0xff13, 0x33), (0xff14, 0x34),
    (0xff15, 0x35), (0xff16, 0x36), (0xff17, 0x37), (0xff18, 0x38), (0xff19, 0x39), (0xff1a, 0x3a),
    (0xff1b, 0x3b), (0xff1c, 0x3c), (0xff1d, 0x3d), (0xff1e, 0x3e), (0xff20, 0x40), (0xff21, 0x41),
    (0xff22, 0x42), (0xff23, 0x43), (0xff24, 0x44), (0xff25, 0x45), (0xff26, 0x46), (0xff27, 0x47),
    (0xff28, 0x48), (0xff29, 0x49), (0xff2a, 0x4a), (0xff2b, 0x4b), (0xff2c, 0x4c), (0xff2d, 0x4d),
    (0xff2e, 0x4e), (0xff2f, 0x4f), (0xff30, 0x50), (0xff31, 0x51), (0xff32, 0x52), (0xff33, 0x53),
    (0xff34, 0x54), (0xff35, 0x55), (0xff36, 0x56), (0xff37, 0x57), (0xff38, 0x58), (0xff39, 0x59),
    (0xff3a, 0x5a), (0xff3b, 0x5b), (0xff3c, 0x5c), (0xff3d, 0x5d), (0xff3e, 0x5e), (0xff3f, 0x5f),
    (0xff40, 0x60), (0xff41, 0x61), (0xff42, 0x62), (0xff43, 0x63), (0xff44, 0x64), (0xff45, 0x65),
    (0xff46, 0x66), (0xff47, 0x67), (0xff48, 0x68), (0xff49, 0x69), (0xff4a, 0x6a), (0xff4b, 0x6b),
    (0xff4c, 0x6c), (0xff4d, 0x6d), (0xff4e, 0x6e), (0xff4f, 0x6f), (0xff50, 0x70), (0xff51, 0x71),
    (0xff52, 0x72), (0xff53, 0x73), (0xff54, 0x74), (0xff55, 0x75), (0xff56, 0x76), (0xff57, 0x77),
    (0xff58, 0x78), (0xff59, 0x79), (0xff5a, 0x7a), (0xff5b, 0x7b), (0xff5c, 0x7c), (0xff5d, 0x7d),
    (0xff5e, 0x7e),
];

fn encoded_module_name(name: &str, code_page: u16) -> Result<Vec<u8>, String> {
    let encoding = codepage::to_encoding(code_page)
        .ok_or_else(|| format!("unsupported VBA code page {code_page}"))?;
    let (encoded, _, replaced) = encoding.encode(name);
    if !replaced { return Ok(encoded.into_owned()); }
    // Keep the real identifier in the Unicode records and source. The legacy
    // name uses default characters, rather than HTML numeric references.
    let mut output = Vec::new();
    for character in name.chars() {
        let mut utf8 = [0; 4];
        let (bytes, _, replaced) = encoding.encode(character.encode_utf8(&mut utf8));
        if !replaced {
            output.extend_from_slice(&bytes);
        } else if code_page == 1252 && (character as u32) <= 0xffff {
            let mapped = MODULE_NAME_BEST_FIT_1252.binary_search_by_key(&(character as u16), |entry| entry.0)
                .ok().map(|index| MODULE_NAME_BEST_FIT_1252[index].1).unwrap_or(b'?');
            output.push(mapped);
        } else {
            output.extend(std::iter::repeat_n(b'?', character.len_utf16()));
        }
    }
    Ok(output)
}

fn valid_module_name(name: &str) -> bool {
    let mut characters = name.chars();
    name.encode_utf16().count() <= 31
        && characters.next().is_some_and(char::is_alphabetic)
        && characters.all(|c| c.is_alphanumeric() || c == '_')
}

fn filtered_project_name_map(bytes: &[u8], removed: &BTreeSet<String>) -> Result<Vec<u8>, String> {
    if !bytes.ends_with(&[0, 0]) { return Err("invalid VBA module name map terminator".into()); }
    let mut output = Vec::with_capacity(bytes.len());
    let mut at = 0;
    let end = bytes.len() - 2;
    while at < end {
        let ansi_end = bytes[at..end].iter().position(|byte| *byte == 0)
            .map(|offset| at + offset).ok_or("invalid VBA ANSI name map entry")?;
        let wide_start = ansi_end + 1;
        let mut wide_end = wide_start;
        while wide_end + 1 < end && bytes[wide_end..wide_end + 2] != [0, 0] {
            wide_end += 2;
        }
        if wide_end + 1 >= end { return Err("invalid VBA Unicode name map entry".into()); }
        let units = bytes[wide_start..wide_end].chunks_exact(2)
            .map(|pair| u16::from_le_bytes([pair[0], pair[1]])).collect::<Vec<_>>();
        let name = String::from_utf16(&units).map_err(|_| "invalid VBA Unicode module name map")?;
        if !removed.contains(&name.to_ascii_lowercase()) {
            output.extend_from_slice(&bytes[at..wide_end + 2]);
        }
        at = wide_end + 2;
    }
    if at != end { return Err("invalid VBA module name map length".into()); }
    output.extend_from_slice(&[0, 0]);
    Ok(output)
}

/// Update an existing project's source. Document modules can be omitted as
/// they retain their native identity; omitted ordinary modules are removed.
pub(crate) fn update_project(project_bytes: &[u8], source: &str, source_code_page: Option<u16>) -> Result<Vec<u8>, String> {
    update_project_with_workbook_document(project_bytes, source, source_code_page, None)
}

/// Only the workbook loader may authorize creation of its missing native
/// component. Ordinary source imports cannot manufacture document modules.
pub(crate) fn update_project_with_workbook_document(project_bytes: &[u8], source: &str, source_code_page: Option<u16>, workbook_document: Option<&str>) -> Result<Vec<u8>, String> {
    update_project_inner(project_bytes, source, source_code_page, workbook_document, false)
}

/// Insert a loader-created component without interpreting or replacing any
/// existing module's source, including bytes from a different host code page.
pub(crate) fn insert_workbook_document(project_bytes: &[u8], source: &str, source_code_page: Option<u16>, name: &str) -> Result<Vec<u8>, String> {
    update_project_inner(project_bytes, source, source_code_page, Some(name), true)
}

fn update_project_inner(project_bytes: &[u8], source: &str, source_code_page: Option<u16>, workbook_document: Option<&str>, retain_omitted: bool) -> Result<Vec<u8>, String> {
    struct AddedModule {
        name: String,
        class: bool,
        document: bool,
        encoded_name: Vec<u8>,
        source: Vec<u8>,
    }
    if source.trim().is_empty() {
        return Err("VBA project source is empty".into());
    }
    if source.contains('\0') {
        return Err("VBA source contains a NUL character".into());
    }
    let project = ovba::open_project(project_bytes.to_vec())
        .map_err(|error| format!("cannot read VBA project: {error}"))?;
    let names = identities(&project, source_code_page)?;
    let declared_encoding = codepage::to_encoding(project.information.code_page)
        .ok_or_else(|| format!("unsupported VBA code page {}", project.information.code_page))?;
    let source_code_page_value = source_code_page.unwrap_or(project.information.code_page);
    let encoding = codepage::to_encoding(source_code_page_value)
        .ok_or_else(|| format!("unsupported VBA source code page {source_code_page_value}"))?;
    let old_sources: Vec<String> = project.modules.iter().zip(&names)
        .map(|(module, name)| module_source(&project, module, name, source_code_page))
        .collect::<Result<_, _>>()?;
    let project_text = project.read_stream("/PROJECT").map_err(|e|e.to_string())?;
    let project_text = declared_encoding.decode_without_bom_handling_and_without_replacement(&project_text)
        .ok_or("invalid VBA PROJECT text")?;
    let documents = crate::vba_module_directory::document_module_names(&project, &names, &project_text);
    let mut edits = BTreeMap::new();
    let mut seen = BTreeSet::new();
    let mut new_names = BTreeSet::new();
    let mut additions: Vec<AddedModule> = Vec::new();
    for module in oxivba_core::parser::source_modules(source) {
        let index = if let Some(name) = &module.named {
            let matches: Vec<usize> = names.iter().enumerate()
                .filter(|(_, identity)| identity.name.eq_ignore_ascii_case(name))
                .map(|(index, _)| index).collect();
            if matches.is_empty() {
                if !valid_module_name(name) {
                    return Err(format!("invalid VBA module name: {name}"));
                }
                if !new_names.insert(name.to_ascii_lowercase()) {
                    return Err(format!("duplicate VBA module {name}"));
                }
                let mut text = source_body(&module.text, module.class)?;
                let document = workbook_document.is_some_and(|document| document.eq_ignore_ascii_case(name));
                if document && (!module.class || !text.lines().any(|line| line.trim().eq_ignore_ascii_case(
                    "Attribute VB_Base = \"0{00020819-0000-0000-C000-000000000046}\""))) {
                    return Err(format!("invalid generated workbook document: {name}"));
                }
                let class_base = "Attribute VB_Base = \"0{FCFB3D2A-A0FA-1068-A738-08002B3371B5}\"";
                if text.lines().any(|line| {
                    line.trim_start().to_ascii_lowercase().starts_with("attribute vb_base")
                        && (!module.class || (!document && !line.trim().eq_ignore_ascii_case(class_base)))
                }) {
                    return Err(format!("VBA document module creation is not implemented: {name}"));
                }
                if module.class && !document {
                    let mut missing = Vec::new();
                    if !text.lines().any(|line| line.trim_start().to_ascii_lowercase().starts_with("attribute vb_base")) {
                        missing.push(class_base.to_string());
                    }
                    for attribute in ["VB_GlobalNameSpace", "VB_Creatable", "VB_PredeclaredId", "VB_Exposed", "VB_TemplateDerived", "VB_Customizable"] {
                        let prefix = format!("attribute {attribute}").to_ascii_lowercase();
                        if !text.lines().any(|line| line.trim_start().to_ascii_lowercase().starts_with(&prefix)) {
                            missing.push(format!("Attribute {attribute} = False"));
                        }
                    }
                    if !missing.is_empty() {
                        let mut lines = Vec::new();
                        for line in text.lines() {
                            lines.push(line.to_owned());
                            if line.trim_start().to_ascii_lowercase().starts_with("attribute vb_name") {
                                lines.extend(missing.iter().cloned());
                            }
                        }
                        text = lines.join("\n");
                    }
                }
                let text = text.trim_end_matches('\n').replace('\n', "\r\n") + "\r\n";
                let (encoded, _, replaced) = encoding.encode(&text);
                if replaced {
                    return Err(format!("VBA source in {name} cannot be represented in code page {source_code_page_value}"));
                }
                let name_bytes = if module.class && !document {
                    encoded_module_name(name, 1252)?
                } else {
                    let name_encoding = if document { declared_encoding } else { encoding };
                    let (bytes, _, replaced) = name_encoding.encode(name);
                    if replaced { return Err(format!("VBA module name {name} cannot be represented in its source code page")); }
                    bytes.into_owned()
                };
                additions.push(AddedModule { name: name.clone(), class: module.class, document,
                    encoded_name: name_bytes, source: encoded.into_owned() });
                continue;
            }
            if matches.len() != 1 {
                return Err(format!("ambiguous VBA module: {name}"));
            }
            matches[0]
        } else {
            let matches: Vec<usize> = project.modules.iter().enumerate()
                .filter(|(_, existing)| matches!(existing.module_type, ovba::ModuleType::Procedural))
                .map(|(index, _)| index).collect();
            if matches.len() != 1 || module.class {
                return Err("VBA source must identify its module with Attribute VB_Name".into());
            }
            matches[0]
        };
        if !seen.insert(index) {
            return Err(format!("duplicate VBA module {}", names[index].name));
        }
        let class = matches!(project.modules[index].module_type, ovba::ModuleType::DocClsDesigner);
        if module.class && !class {
            return Err(format!("VBA module kind changed: {}", names[index].name));
        }
        let mut text = source_body(&module.text, class)?;
        if module.named.is_none() {
            let attributes = old_sources[index].lines()
                .take_while(|line| line.trim().is_empty() || line.trim_start().starts_with("Attribute "))
                .filter(|line| line.trim_start().starts_with("Attribute "))
                .collect::<Vec<_>>().join("\n");
            text = format!("{attributes}\n{text}");
        } else {
            // CodePane/exported class text can omit hidden module metadata.
            // In particular VB_Base is required to rebuild a source-only class.
            let attribute_key = |line: &str| -> Option<String> {
                let rest = line.trim().strip_prefix("Attribute ")?;
                let key = rest.split_once('=')?.0.trim();
                (!key.contains('.')).then(|| key.to_ascii_lowercase())
            };
            let present: BTreeSet<_> = text.lines().filter_map(attribute_key).collect();
            let missing: Vec<_> = old_sources[index].lines()
                .take_while(|line| line.trim().is_empty() || line.trim_start().starts_with("Attribute "))
                .filter(|line| attribute_key(line).is_some_and(|key| !present.contains(&key)))
                .collect();
            if !missing.is_empty() {
                let mut lines = Vec::new();
                for line in text.lines() {
                    lines.push(line.to_owned());
                    if attribute_key(line).is_some_and(|key| key=="vb_name") {
                        lines.extend(missing.iter().map(|line| (*line).to_owned()));
                    }
                }
                text = lines.join("\n");
            }
        }
        if comparable(&text) != comparable(&old_sources[index]) {
            let text = text.trim_end_matches('\n').replace('\n', "\r\n") + "\r\n";
            let (encoded, _, replaced) = encoding.encode(&text);
            if replaced {
                return Err(format!("VBA source in {} cannot be represented in code page {}", names[index].name, source_code_page_value));
            }
            edits.insert(index, encoded.into_owned());
        }
    }
    let removed: BTreeSet<usize> = project.modules.iter().enumerate()
        .filter(|(index, _)| !retain_omitted && !seen.contains(index) && !documents.contains(&names[*index].name.to_ascii_lowercase()))
        .map(|(index, _)| index).collect();
    let mut project_names: BTreeSet<String> = project.modules.iter().enumerate()
        .filter(|(index, _)| !removed.contains(index))
        .map(|(_, module)| module.name.to_ascii_lowercase()).collect();
    for addition in &additions {
        let name = declared_encoding.decode_without_bom_handling_and_without_replacement(&addition.encoded_name)
            .ok_or("invalid VBA module name encoding")?;
        if !project_names.insert(name.to_ascii_lowercase()) {
            return Err(format!("VBA module name {} has an ambiguous project code page identity", addition.name));
        }
    }
    if edits.is_empty() && additions.is_empty() && removed.is_empty() {
        return Ok(project_bytes.to_vec());
    }
    let mut directory = crate::vba_compression::read_stream(&project, "/VBA/dir", 0)
        .map_err(|error| format!("cannot read VBA directory: {error}"))?;
    let table_offset = if additions.is_empty() && removed.is_empty() { None }
        else { Some(module_table_offset(&directory, &project, source_code_page)?) };
    let mut sources = Vec::with_capacity(names.len() - removed.len() + additions.len() + 4);
    let mut removed_streams = Vec::new();
    for (index, (module, name)) in project.modules.iter().zip(&names).enumerate() {
        let path = format!("/VBA/{}", name.stream_name);
        if removed.contains(&index) {
            removed_streams.push(path);
            continue;
        }
        let bytes = if let Some(edited) = edits.remove(&index) {
            edited
        } else {
            crate::vba_compression::read_stream(&project, &path, module.text_offset)
                .map_err(|error| format!("cannot read VBA source {}: {error}", name.name))?
        };
        sources.push((path, compress_source(&bytes)));
        directory.get_mut(name.source_offset_position..name.source_offset_position + 4)
            .ok_or("invalid VBA source offset field")?.fill(0);
    }
    if let Some(table) = table_offset {
        let total = u16::try_from(names.len() - removed.len() + additions.len()).map_err(|_| "too many VBA modules")?;
        directory.get_mut(table + 6..table + 8).ok_or("invalid VBA module count")?
            .copy_from_slice(&total.to_le_bytes());
        if !directory.ends_with(&[0x10, 0, 0, 0, 0, 0]) {
            return Err("invalid VBA module table terminator".into());
        }
        let mut rebuilt = directory.get(..table + 16).ok_or("invalid VBA module table header")?.to_vec();
        for (index, name) in names.iter().enumerate() {
            if !removed.contains(&index) {
                rebuilt.extend_from_slice(directory.get(name.record_start..name.record_end).ok_or("invalid VBA module record bounds")?);
            }
        }
        directory = rebuilt;
        let old_map = project.read_stream("/PROJECTwm")
            .map_err(|error| format!("cannot read VBA module name map: {error}"))?;
        let removed_names: BTreeSet<String> = removed.iter().map(|index| names[*index].name.to_ascii_lowercase()).collect();
        let removed_project_names: BTreeSet<String> = removed.iter()
            .map(|index| project.modules[*index].name.to_ascii_lowercase()).collect();
        let mut map = filtered_project_name_map(&old_map, &removed_names)?;
        map.truncate(map.len() - 2);
        let mut properties = project_text.to_string();
        if !removed_names.is_empty() {
            properties = properties.split_inclusive('\n').filter(|line| {
                let Some((kind, name)) = line.trim_end_matches(['\r', '\n']).split_once('=') else { return true; };
                !matches!(kind, "Module" | "Class") || !removed_project_names.contains(&name.to_ascii_lowercase())
            }).collect();
        }
        let mut insertion = None;
        let mut position = 0;
        for line in properties.split_inclusive('\n') {
            position += line.len();
            if ["Document=", "Module=", "Class=", "BaseClass="].iter()
                .any(|prefix| line.starts_with(prefix))
            {
                insertion = Some(position);
            }
        }
        let insertion = insertion.ok_or("VBA PROJECT has no module declarations")?;
        let mut added_properties = String::new();
        let mut streams: BTreeSet<String> = names.iter().map(|name| name.stream_name.to_ascii_lowercase()).collect();
        let mut serial = names.len();
        for AddedModule { name, class, document, encoded_name, source: bytes } in additions {
            let stream_name = loop {
                serial += 1;
                let candidate = format!("OxiModule{serial}");
                if streams.insert(candidate.to_ascii_lowercase())
                    && project.read_stream(format!("/VBA/{candidate}")).is_err()
                {
                    break candidate;
                }
            };
            let wide_name = unicode_name(&name);
            module_record(&mut directory, 0x0019, &encoded_name)?;
            if class { module_record(&mut directory, 0x0047, &wide_name)?; }
            module_record(&mut directory, 0x001a, stream_name.as_bytes())?;
            module_record(&mut directory, 0x0032, &unicode_name(&stream_name))?;
            module_record(&mut directory, 0x001c, &[])?;
            module_record(&mut directory, 0x0048, &[])?;
            module_record(&mut directory, 0x0031, &0u32.to_le_bytes())?;
            module_record(&mut directory, 0x001e, &0u32.to_le_bytes())?;
            module_record(&mut directory, 0x002c, &0xffffu16.to_le_bytes())?;
            module_record(&mut directory, if class { 0x0022 } else { 0x0021 }, &[])?;
            module_record(&mut directory, 0x002b, &[])?;
            map.extend_from_slice(&encoded_name);
            map.push(0);
            map.extend_from_slice(&wide_name);
            map.extend_from_slice(&[0, 0]);
            added_properties.push_str(if document { "Document=" } else if class { "Class=" } else { "Module=" });
            let project_name = declared_encoding.decode_without_bom_handling_and_without_replacement(&encoded_name)
                .ok_or("invalid VBA module name encoding")?;
            added_properties.push_str(&project_name);
            if document { added_properties.push_str("/&H00000000"); }
            added_properties.push_str("\r\n");
            sources.push((format!("/VBA/{stream_name}"), compress_source(&bytes)));
        }
        module_record(&mut directory, 0x0010, &[])?;
        map.extend_from_slice(&[0, 0]);
        properties.insert_str(insertion, &added_properties);
        let (properties, _, replaced) = declared_encoding.encode(&properties);
        if replaced { return Err("VBA PROJECT declarations cannot be represented in its code page".into()); }
        sources.push(("/PROJECT".into(), properties.into_owned()));
        sources.push(("/PROJECTwm".into(), map));
    }
    let mut header = project.read_stream("/VBA/_VBA_PROJECT")
        .map_err(|error| format!("cannot read VBA project version: {error}"))?;
    if header.len() < 7 || header[..2] != [0xcc, 0x61] {
        return Err("invalid VBA project version header".into());
    }
    header.truncate(7);
    header[2..4].copy_from_slice(&0xffffu16.to_le_bytes());
    header[4] = 0;
    drop(project);
    let mut compound = cfb::CompoundFile::open(Cursor::new(project_bytes.to_vec()))
        .map_err(|error| format!("cannot open VBA compound file: {error}"))?;
    let caches: Vec<_> = compound.walk().filter(|entry| entry.is_stream())
        .filter(|entry| entry.path().parent().and_then(|p|p.to_str()).is_some_and(|p|p.eq_ignore_ascii_case("/VBA")))
        .filter(|entry| entry.name().to_ascii_uppercase().starts_with("__SRP_"))
        .map(|entry| entry.path().to_path_buf()).collect();
    for path in caches {
        compound.remove_stream(path).map_err(|error| error.to_string())?;
    }
    for path in removed_streams {
        compound.remove_stream(path).map_err(|error| error.to_string())?;
    }
    sources.push(("/VBA/dir".into(), compress(&directory)));
    sources.push(("/VBA/_VBA_PROJECT".into(), header));
    for (path, bytes) in sources {
        compound.create_stream(path).map_err(|error| error.to_string())?
            .write_all(&bytes).map_err(|error| error.to_string())?;
    }
    Ok(compound.into_inner().into_inner())
}

#[cfg(test)]
mod unicode_module_tests {
    use super::*;

    fn original_project() -> Vec<u8> {
        let template = include_bytes!("../../../tests/fixtures/vba/named_project_ascii.xlsm");
        let mut archive = oxidocs_common::archive::OoxmlArchive::new(template).unwrap();
        archive.read_binary_part("xl/vbaProject.bin").unwrap()
    }

    const BASE: &str = "Attribute VB_Name = \"ProjectSourceProbe\"\nPublic Function Control() As Long\n Control = 1\nEnd Function\n";

    fn source_with_class(name: &str) -> String {
        format!("{BASE}Attribute VB_Name = \"UnicodeWriterProbe\"\nPublic Function Probe() As String\n Dim obj As {name}\n Set obj = New {name}\n Probe = obj.Label & \"|\" & TypeName(obj)\nEnd Function\nVERSION 1.0 CLASS\nAttribute VB_Name = \"{name}\"\nPublic Property Get Label() As String\n Label = \"FreshClassOK\"\nEnd Property\n")
    }

    #[test]
    fn unicode_module_addition_rejects_colliding_project_names() {
        let source = source_with_class("\u{9867}\u{5ba2}");
        let extra = "VERSION 1.0 CLASS\nAttribute VB_Name = \"\u{5e33}\u{7968}\"\nPublic marker As Long\n";
        let combined = source.clone() + extra;
        let original = original_project();
        assert!(update_project(&original, &combined, Some(932)).unwrap_err().contains("ambiguous project code page identity"));
        let changed = update_project(&original, &source, Some(932)).unwrap();
        assert!(update_project(&changed, &combined, Some(932)).unwrap_err().contains("ambiguous project code page identity"));
    }

    #[test]
    fn unicode_module_addition_supports_standard_names_and_class_best_fit() {
        let name = "\u{65b0}\u{898f}";
        let source = format!("{BASE}Attribute VB_Name = \"{name}\"\nPublic Function Probe() As String\n Probe = \"UnicodeModuleOK\"\nEnd Function\n");
        let changed = update_project(&original_project(), &source, Some(932)).unwrap();
        let project = ovba::open_project(changed.clone()).unwrap();
        assert!(identities(&project, Some(932)).unwrap().iter().any(|m| m.name == name));
        assert_eq!(update_project(&changed, &source, Some(932)).unwrap(), changed);
        let removed = ovba::open_project(update_project(&changed, BASE, Some(932)).unwrap()).unwrap();
        assert!(!identities(&removed, Some(932)).unwrap().iter().any(|m| m.name == name));
        let properties = String::from_utf8(removed.read_stream("/PROJECT").unwrap()).unwrap();
        assert_eq!(properties.lines().filter(|line| line.starts_with("Module=")).collect::<Vec<_>>(),
            vec!["Module=ProjectSourceProbe"]);

        let source = source_with_class("\u{3a9}");
        let changed = update_project(&original_project(), &source, Some(932)).unwrap();
        let project = ovba::open_project(changed).unwrap();
        assert!(identities(&project, Some(932)).unwrap().iter().any(|m| m.name == "\u{3a9}"));
        let properties = String::from_utf8(project.read_stream("/PROJECT").unwrap()).unwrap();
        assert!(properties.lines().any(|line| line == "Class=O"));
    }

    #[test]
    fn unicode_module_addition_preserves_names_and_removes_ansi_declarations() {
        let name = "\u{9867}\u{5ba2}";
        let source = source_with_class(name);
        let changed = update_project(&original_project(), &source, Some(932)).unwrap();
        let project = ovba::open_project(changed.clone()).unwrap();
        assert_eq!(project.information.code_page, 1252);
        let names = identities(&project, Some(932)).unwrap();
        let index = names.iter().position(|m| m.name == name).unwrap();
        let body = module_source(&project, &project.modules[index], &names[index], Some(932)).unwrap();
        assert!(body.contains(&format!("Attribute VB_Name = \"{name}\"")));
        let properties = String::from_utf8(project.read_stream("/PROJECT").unwrap()).unwrap();
        assert!(properties.lines().any(|line| line == "Class=??"));
        assert_eq!(update_project(&changed, &source, Some(932)).unwrap(), changed);
        let removed = update_project(&changed, BASE, Some(932)).unwrap();
        let project = ovba::open_project(removed).unwrap();
        assert!(!identities(&project, Some(932)).unwrap().iter().any(|m| m.name == name));
        let properties = String::from_utf8(project.read_stream("/PROJECT").unwrap()).unwrap();
        assert!(!properties.lines().any(|line| line == "Class=??"));
    }
}
