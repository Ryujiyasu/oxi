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
    update_project_inner(project_bytes, source, source_code_page, workbook_document, false, false, false)
}

/// Insert a loader-created component without interpreting or replacing any
/// existing module's source, including bytes from a different host code page.
pub(crate) fn insert_workbook_document(project_bytes: &[u8], source: &str, source_code_page: Option<u16>, name: &str) -> Result<Vec<u8>, String> {
    update_project_inner(project_bytes, source, source_code_page, Some(name), true, false, false)
}

/// Build a copied worksheet's project from its document component only.
/// Ordinary source imports keep their independent omitted-document contract.
pub(crate) fn copy_document_project(project_bytes: &[u8], name: &str, source_code_page: Option<u16>) -> Result<Vec<u8>, String> {
    let project = ovba::open_project(project_bytes.to_vec()).map_err(|error| error.to_string())?;
    let names = identities(&project, source_code_page)?;
    let encoding = codepage::to_encoding(project.information.code_page).ok_or("unsupported VBA document code page")?;
    let properties = project.read_stream("/PROJECT").map_err(|error| error.to_string())?;
    let text = encoding.decode_without_bom_handling_and_without_replacement(&properties).ok_or("invalid VBA document declarations")?;
    let documents = crate::vba_module_directory::document_module_names(&project, &names, &text);
    if !documents.contains(&name.to_ascii_lowercase()) { return Err(format!("copied document module {name} is missing")); }
    let matches = names.iter().enumerate().filter(|(_, identity)| identity.name.eq_ignore_ascii_case(name)).collect::<Vec<_>>();
    if matches.len() != 1 { return Err(format!("copied document module {name} is ambiguous")); }
    let (index, identity) = matches[0];
    let source = module_source(&project, &project.modules[index], identity, source_code_page)?;
    let source = format!("VERSION 1.0 CLASS\n{source}");
    update_project_inner(project_bytes, &source, source_code_page, None, false, true, false)
}

/// Only a real sheet copy may authorize a new native document component.
pub(crate) fn import_sheet_document(project_bytes: &[u8], source: &str, name: &str, source_code_page: Option<u16>) -> Result<Vec<u8>,String> {
    update_project_inner(project_bytes, source, source_code_page, Some(name), true, false, true)
}

fn update_project_inner(project_bytes: &[u8], source: &str, source_code_page: Option<u16>, workbook_document: Option<&str>, retain_omitted: bool, remove_omitted_documents: bool, allow_sheet_document: bool) -> Result<Vec<u8>, String> {
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
                let expected_bases: &[&str] = if allow_sheet_document {
                    &[r#"Attribute VB_Base = "0{00020820-0000-0000-C000-000000000046}""#, r#"Attribute VB_Base = "0{00020821-0000-0000-C000-000000000046}""#]
                } else { &[r#"Attribute VB_Base = "0{00020819-0000-0000-C000-000000000046}""#] };
                if document && (!module.class || !text.lines().any(|line| expected_bases.iter().any(|base| line.trim().eq_ignore_ascii_case(base)))) {
                    return Err(format!("invalid generated native document: {name}"));
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
        .filter(|(index, _)| !retain_omitted && !seen.contains(index) && (remove_omitted_documents || !documents.contains(&names[*index].name.to_ascii_lowercase())))
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
                let declaration = matches!(kind, "Module" | "Class")
                    || remove_omitted_documents && matches!(kind, "Document" | "BaseClass");
                let name = if kind == "Document" { name.split('/').next().unwrap_or(name) } else { name };
                !declaration || !removed_project_names.contains(&name.to_ascii_lowercase())
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
    fn copied_document_project_keeps_only_the_selected_complete_native_sheet_source() {
        use std::io::Read;
        // Self-authored complete original Excel208 project corpus: CC BY-SA 4.0.
        const PACKED: &str = concat!(
            "504b0304140000000800d4926d5d512ddd41278c000000da01000b00000070726f6a6563742e62696eecbd097c13c7f5",
            "38fe7657b2e5fbc0808180176140026324198c3128b66cd998d3c636579a04cbd61a2b08cb4832b12909324720274712",
            "a02790a4699aa60d49d3343736e4a0495aa049137a7c0339daa627a4c737b44de3ff5c2bed4a2b6353d26ffbfb57feac",
            "349e9de3cdcc7b6fde7b33f3e6d4c9ac73871f1bfd2e447dae06013eed4f8204451c8f1e4efe2793fe8f9f4ffbfbfbe5",
            "e89bd1d3ffdfcf7fcce71fe8494463f643f4e8d0a3470f1e73037a52d193849e64f4a4a0e715f49cfaeff8fe3ff1f907",
            "fbad071ffa0b820855d08e7efdd01dcd0a06fc8c40182397997989b41732cbbff083d5ffc30928dc2bd2b8e550018e21",
            "d5a8fe18104792ebd7c549b37b5219a957ae5ff9ae11dac0030158817ac10f6ba199f4c7da41d73f1a780ed30aa60f0c",
            "c360f37d8ebb749ac17c7850d3e350f373ac0c3c2632fd637e8e79026e9792fe97a0270d3de9e8c900320540167ab2d1",
            "330c3d392c2f7e46a2702e7a4601ee238031e8b90a3d63d1330e3d79e811d1331e3d46f44c404f3e7a26a267127a26a3",
            "c7841e337aa6a0672a7a0ad0330d3d85e8998e1e0b7aace8b1a1a7083d33d033133dc5e899859e12f4cc464f297ae6a0",
            "672e7aec80e7368032f494a307e35e057a2ad1e3444f157aaad1330f3d35e8998f9e05e859889e45e859ccda58cb7eff",
            "133e1b19ac9bd0ef4d40e7e8cde809e1f15ff118388241bfa7b9330892b8bc62f512d73a14b08bc6c63644182b7cfeb5",
            "cd3e4411c6b4649c0e70429c0e5657b802381d182d9fb7582c360b9458674f43210b90ef69951640df161243be6714df",
            "04ac180fb06256c33cafafd9e55de28275524387ab4512c12e56bbbc01290d921964409356fa4172055dcdde01d2d4a1",
            "346ea9c5ebf24be09eef46f035faa1538a01beaaabc31790dc401274c696d328c1ba0eaf2b283925f07b3648b824a055",
            "263b8075980404acce40107ceb3c1b31648a02073536dc5fe2f67f439b0452d04a7a0c588522a014a4e3510fd08e079b",
            "a5c486fb59eeed695069a141daeda8e3a1f826daf1c08a416c7735ed7840f5918e2795e21682143346957edaf103a5a9",
            "f3d38e07f44b3a1e483f44034f3a1e243a329d1ae5344ab4e3c129d18e8f8f0f9d1008d28e07af142910ff4d06b1a1ce",
            "b972da220fb448ed0169da7c3748ed414fab47f243a9585939ad62d534687070231f0f7184f30960797801bf3b017279",
            "be2609c4e49c841a01cd0ff081cea0433cf0aae5158eba90df7783d412d4c368c4dccaf9646e036229dfe29313793f07",
            "8684ec243e3bd490f4772f07a9bc31792e9f3c829f1d08ba7d06af74353f2a80265870a3f905752498d2d0d493753d08",
            "53ae85799482669c2942a367d0e92af5498821237a99602bb44c00cb84cad26b5778dac1edbb31706d4377a03c28ad2b",
            "b25d2b4836280c7a9b27d42eaa02d1d1893ac315ccf5f8dab72edd5a55dbdaeaf4b46cabaa85d610781aa1654bd55743",
            "867bb655d9c0595de2b4cca89c06332baa1dd3ac166b79c5b40a67d5cc5095e30647a87246486fdb59153ab006fcae75",
            "62b5076175e0da4adfba75be764e97b0d8d3e2f7057cd01a141bda08995d5b5b5d3dbfb20aacc5d72e6ea82d74e62e5a",
            "346147e6d6c5a2b5582cb488b5cd3d7de2a2724fb3dfe5efdeda9491b5e54d01866d1176fe7514b7a5197119cc640073",
            "99b5f320776a8f21f44d4fe88915bdabfcf0109a907b395f2f37fa87290fdb7e9a7255c5d89a0acefaf2c6dbc655f305",
            "c7c6dff2536385903cb5821b55b10a918bf450d03aefd89886f2915239dc10eccd1c5d93685b92f8fe98fdb73df9da98",
            "2447966b83042eb1599a07d9fc70173cbc0186d7858a7ca12209c6dd795b92ed9ea416f1fa15e36b443eb38223b323fd",
            "ee1322336696211276185e77a13997bc95297a8f22ad01cd434de877bd22cf63e11007870ddcf79f8f4bffae0d20d5f9",
            "7dcd9211a1fe648acfb0c843f119e6bb293e83e4a7f80caba63538a6cd28b4405a726d4710a1018888dabc9e160f04d3",
            "92eb3a9bbd8818c4eacef696a0077cede20a8f3bd806a4065418a0d282dd26b3088e80b8c8d7be06552aa20f4a05c136",
            "9a0a489dc16e4c93368badd88ab98fd596965c8570335232ad0c3cac3240a088f5aef6b550eff3054d15ddcbc1e515a5",
            "ae0e843c81007ee9082072459dd0bec68cf03840c28020200080d3b34e74b5076e442dc52f972316e371b5a32689805e",
            "e3378847a31e737474e04a5db8e98550b5c1e5edc47cc4446a0a043018665c24ca85a1a947925e106543580c1d92a3dd",
            "bd1c71c34ec9c40a34e35621ee435b853a9375610bb086fa80b6cae187359deb50c7982aa01b958121821b19ace022b0",
            "d25e254d04d6ac7a20795156cc98081ce062708097c1819a4ce15054aa3594188e6af0f9d7757a5d3218a8d980db8d3b",
            "58517b8158815f83e895d6b85a1a101342d30c624c56b89c4f3a921eb19e28cba183c983f5cbbffe9bc8bfff7fafbf9b",
            "a04685cf8798bcabdd4c880f082532f2c2c487f16883c7d71988e60d9882e90b34dd63f2c3d487e9acb0d2e56d01848c",
            "98725862f41ebc8c3c812440ef01ff877277815719b3d8d5de899804630008bbd74826e362b0962e9d69341756a2e9de",
            "e5aff4b50731e904184dcf6fc5280dae966eb1b14dc2f42a920f2577540018175b51ee6aa4e9613a21b2ab5d4dc14610",
            "27a938d32494c66c2445216e4ac43fd22d320513b8ac408a45a582d765c3bcdb0ef58cba81962b4e0251c18750b94633",
            "d07245a479bb11f4ac6425298b76102b11dd9aea5dc0aa3282b910916f279afacd1146120183f42260014ac1ea944da6",
            "fffbd7e1d1c1606c02d71a97a7bdd428c2245699aa405219d8cc431a0b8a0b40597125e2672d1417008f35ea850d80b1",
            "4663de40221fca2591388a1b0d2d28bbcb5f2f75202981cc1e88b70166ee12993796530e0b66367180874d1c4070d723",
            "3a80e16c3219388426f32150e5f7fbfc2654172ac68cf00501cf1a88bf509d5e607506b1486fac52f40fcbc5c60b23c5",
            "7c68c570347677806422a09971ae0dcd0c9e1884c415b8fc402bc0a36c6c4015206423152c4289593166652e346ca83d",
            "68eeb58a8d3e71118299a69231218c990d40ca47c503291fd5072c026142100f7b294175521d22ea9615a6c5689ecfa7",
            "e51580e82910ad6675ed4ba4aea0e86131555e20b410b7c768c9b45310e9b1d6e03e342a5aba814cb96699be281d0099",
            "78110600997891380f0431f064ae987b1d7ebfabbb1e488d082d00e305ae06242a25a09e27336f44b8608c08a105887e",
            "df8d381950cc20adc0e8d282847391323979a2760289f7a289ba3d40dfa96513b7671d929e30bb4058853a3c3a0dc6b8",
            "0082165cddf2781184600849d2d4b68b042541ac970288178980bb9b01463adc0ffec24a4c6bac15a4c7085848dc401d",
            "beacc2877acc1d1e419bdcabf8834044d8424104540d4a4f0a5cd2b9ae1929448a51a680e056ccf32124038b02b55043",
            "dc1eccd04841e0a705cd06d61a31d21ccceef1f02831dc612d256c90a278052070dd320298314a02410cf41afc26d21a",
            "b72615e00f6668640811435b54414b9269d907b42fdc32d629fb391ca0d81326103405d0ff3193216c91d0bb8a4e5915",
            "26542f98cd0ae42701421ef89502543a61a86bc7154337eb17acc53b6c0adeb208586b880c488810444633aceb96a9ba",
            "8e2451a204ab03586b0831aac602954349b0121a820de002a466401da0318766141ada2707492043953f46a0e7076943",
            "9554b43fffe9f2d76a66015f8d46a01e6a6101544125b63a0cf233fa32fa7f387a1259983b7df45fa57f620e8e2d3712",
            "61d2443440c20a78953c353cff23baf342b3ab65ed2257b3c4d260c6e0a1a980cc0622921603412491b87deb60795822",
            "70fa907e84cd458a4a238aa87aa662f204a56c32213079229e22ea2155b0598331c5f90160f204a15ac2dda15dc13889",
            "d00426e3354516243251506d98291206a5e4100d6ce206ca361153ac625c1149e526225601d370ab8069b848b0adaaaf",
            "afad2f6c84557555265409122dcde14915cb27ad62f4544c6a6f663a61cc74140b4683028c4558f0a0ec56c55ca97c82",
            "5832954f22524c347f52cb74443e51890f6a86873a9cca27f90cf802248650f944593b954f580c6ef580021729199b12",
            "48a7b0813313795fd15295781596d349208c51d8dc5807c4348022815a0b44a48407a51a5f20e8003f552e2ac1159008",
            "be2111c3d721f9a9aaa450b90a44a6ba07bc3ed42db21812d6d85db424082ef42000643c2c106b81901b968b293abb3a",
            "a083a95ad0eef07a5c0122051335cf4564706ac435abc82a2c04057cd0e96f51918352747269586170b66660829707e8",
            "0b59066f21b5d39220a890aadcf1aa60e2bd76ae5697073110bfa4258b49880e01a9236e29a0915d22320b7e0b01362d",
            "42b7c966b1580a44403fb30a449b05ac33f18fad08c8cf6cf2aaa898bc9a1196aa1049b964f50e0f079e662d3184df20",
            "05c5009aae716fdac3ca9303294f8e1946b312575979c4184486978af9517a4380141599d24d46ab315a11a15802ac18",
            "84ea364ade0a1a2460f9501f12b064e5d289954b19d529d770c99af1420fc679bb38436c8436a62b315a9610c152b830",
            "ccb2aa8da43d4789d1ac54176813819587353fbb383346f7c24501818b89cc95cbb1408a47695658d2a9c2d8ab3504c5",
            "831e8259c6680ec9c82b08a4e3aec67ccc223a50d02593f042c06fe6a29181e2189ec9fa007538191b532560c04d1291",
            "a609469aa2aa9886cab198cdb25192c2114de657e311a4704465477080ad3866702363417ace0412437a401044e1d834",
            "54b4f91278883aa208d78fc64b8d5a73511545b11d2e4f768ed9d87a006cae430a1feb0e6024180093baa5d34428b284",
            "7b83e216041538af2c390e9fd7229002d16803a32cb03325336c0cf00161c5445c41131d9a21d04cd42255186330935a",
            "8c5ccd2dc652ccebeca2ad14b1364ca7a522213463975131944adcc0598d2e632988cd24d40cb8901632576a1981a850",
            "51db8c179b4c0c7bcd31a48791164d157802b550258cf618ae8df170456762d462a9c1ca5203220e0c11ab42a1c3614e",
            "4253cb0c04490528358139c052471a389fa5c62841b916213cd22d2cb58272d53a65c04b0df116da40dacf0356c11823",
            "059df231ed6eb1b16e21c3c47043812db448228114d1e1c61047f506411868a5d96a8129ea889350457d0930451d7312",
            "3fd1d1119d46547f32d3cb664e2354234cc7b8a531942e8501954cdada8a359549d87a47581aacc6aca9c2e442b35573",
            "81d86256e422cc5c511709d01916332b5a2b81b090acc405da2429189664062e5989b4f204a4906db01a0192cbdfd286",
            "da0cac354a8474296ac76d46f372345f15599b8956a25ee369a04523be5a2012d82ea57cc72907ef72c0ad8660d8bc54",
            "48cb06d2705436b428357eb23234bf556629ac7faf58bb3013c23c0858e52d4814712b1963b4f98d04d8904659e271db",
            "800d29b3a81562c19e54607201ad0035d03d4003953c138b5f88b7489a36a4da7660f4310fb0fdc8c270a3024bc380c5",
            "617972a964bb188820de0a58a6f34bd440b889180895161a13ed288d35394d9bb1cafe3164150fdb95a2163a8850dae2",
            "f3c6331962235cb462484c83485ea5e632aa216a2b8f545e88a73c128b1c31da618e0398e5608ea398f1a3d90ed2e9f1",
            "0f360e02060c1b07d536239b4a73723320a19d0269a745a2810534b20ace814061ed98e70336ba4aac6705818f158486",
            "77768c0c8b3f6a739b9d18078159ff90acb74865bf8bb11a2fab40931b6d0d460a73346fc37a28b6c4a1821729ada268",
            "c011d0d838082c77ace95cd376c950916ac3d8524930346a65829467420375a3d91ccdb3896e8a5f29408d9e0949bf00",
            "437a62a627c641a5557491d22a8a48d64a0da7caae53993e9192ac61a9bc94e9730073a42dd61c29a37f8c99158f0250",
            "1beda21893259a007de2b298e644831a2e07d35e6c39365a0eb58243406b28b57a556d1e650c07988d405e51a04dc6c6",
            "de0242fa666d3b321959f43a1ac3974840edc1ca5c2a3e1a2dba2ad73388a4b99c62996a2d27c208d505b21513caff30",
            "57843a669100c62b116fc3fba9fc26caf6301362b5cbfa9f43b9b848055632ca6dfe15a64935e02c29721203098b7042",
            "95c5420df7cd46f5eaa0036c6ce51088ec8f7200c9828a885eb6741429561969da1580d3e2d2a3219c110561a36f6d37",
            "f808f6b32a8a67cd42384f23483933aa1c9531e5cc549763b515cd98199da6589d86edeaa349da8124413a24d0855faf",
            "8bc063b74eb7c8cbbae10696907559bc7403789d142f9346d55511d5f3d3ac91c560a2f154d8d4ba94455d4545541fc6",
            "649fa1ce6e5367afb4aab3a3818f6a69a50db794d92dc1ce7002e820317cc695014e5ba4b6711a6f2e3346c1e3b4aae1",
            "b1aae171dad4f0d8a2b317a9b317a9b35759a3d0cf1a0561b535aa3936c69619fac23c2b2d00580d45d1b8310b8f17dd",
            "b0006ca95dc4d2ba622f5b21d9d21a2874b891c082375996da8db5e0f7acf1b453fe83f4d27aa91569e781461f7a07f6",
            "cf1b897209c69b8c432cb052f282979627012bcf0e93315d38b09172838477c705a110672336cec9e3f31df908eb9859",
            "0d1bd4253f5198ac54c60af30d6a478fd97184053aca57b120270b5988e975d03d47d0ae588f277646608bb0742c707e",
            "3a13a1fc2085ad716ca94d2dc35611868f4ddd8a4d4d66300f1e54267be269bd929961f1263485721681b0206c46555b",
            "bf6473638188973780185c59392e852ae64082bc8b2c58a0d90a1b5cdb3115902dac1adbc0884593a974dad653d950a7",
            "d84ec6fa10bf45bd41260c6ceca4e0c748e2684ea19a2a9e92228a698cb9d112bb4d81d80981d9cec20ccd0acc7c39b0",
            "ddd13a80694a69269c1121c6384602924f8124264228685e3636bbdc08efc1d8d2ec8ad884d4ba0b158e5dc084e3580d",
            "952ab8b1f61edc47c0d0100f3735a4476bb12440552fb0ab37be5463b4a940c0625069dbafa4864a4c09a80f51f13292",
            "a874446cb1a3dd0011e3226d3aa10bac9c63a58128e757dc2021ebfd889d965100b1595729ab6b09c04c4b4365297b9e",
            "28b101604a2c1572b0820eb482322314c8d61da53cafb2dc445b41fee90662518d2ad1cc5885170a906c5b201abb22e6",
            "46d650295a3ea4e2252953b156431b8af090369472bfb049d2c4461909be585ac61519556b6be10d348c0c808d3f7ae5",
            "0f6f3820523bdebd413577b2d395648fec41994755775a0e629f4057b2c8b445f4767ba4c0a1ebeb317b836bd1248668",
            "b0a1cd85d727d94cd2056c266957aae91a3b14716ef48ae4c680e1053db2886d52116315412d32872826a958056a7e80",
            "985d914237a8ece6d84d6cb889c0960399c505db65bd120936a2b6af95970103128ec5c721144b6a66ad952ba4d3d3e5",
            "5e5c2d923cbc01dc5689c6015d4b47f4244e25c21f65dea46c5c25aa911840e5da06372c549a90da5ba446c4745be8da",
            "a557bd8f5163f1721d5e69082f5ec65962f433653a462e20735a2bb506695a613ae85e5408a8769046d9656822e80c84",
            "6534caeb2fb94955227449372700ed65d132a819556b9e415d4150d212331346c851b58d4dcd0ce81676a0628fbc2535",
            "660d0c5742949818cb913f6c5f8992add0a8639110b39901642b2dae79f99b78f1972c80937db8c01427aac9d9a3b7c8",
            "2a2a95b5109c4d51a9d200ead756ee4db2444f6aa41bfc6dca4952bdb196286995c094344d3eca501388f513b15020d6",
            "4f59414114e30746315460a14c45ce658e6c2df32bcc0f2d641f4c785396df24b14121740d080d6598e36ccf53cf96e1",
            "3dc06458e290cce0f896d3877736489883c8323ec23acc46f0c92c447564b30ddb49c01233502586bd8c379115f79876",
            "117ec5a8db876b972e93592d73b72246e45a2305e4bda01897809d7f5048c771f60987b7b2e33d496cb7727bb4a5c4e9",
            "8175a2ab3528f95720861694c26ca700c52284f4577bfc81602457014d8c44a1169f62c706c316322ce0634c4a315eaa",
            "2519153c6cb4555c229ac575799531880851877ba3145c420ed4644c4d121064d63360530a46093acd60cd402b3b225f",
            "20f44b8c55987e5593362d8e76cb0adc59c0c65fb5f74b6d375a0c56b6859db06579058e742290bec57545431865ab88",
            "c325e848a089dd473632486c3ec7a0d252685d188f8280f108132f21469c13681b624cd1e81db0618fb18cd2b637e0b1",
            "0737a37c9c46695d1eec667b664102365c41d9d8f999117e2536b123f01139c25a13256ff031f20651d4d8c667d7983d",
            "1179334c5015a8c50a62cf39113a0706175bdaa13cbc8150fa25ac0e05b276de1990e4131036d546a838c209e61b8c17",
            "284f2e686c0b6f965a111548f5922b86ba8994c3acb9912de8581124b0c75a02c2a76c70d7d15aa996167fce65300f78",
            "da02b303450c9e94113bd024197a9203d8ec47e5b1588e6d0f6b065885227d2b9f7c19d4b91b6a14a4fc423170daf286",
            "1a4295a9358ec0a0d281d0f8a0b9928e0fd863a7245aa93f62745729316a71855a2864e90b8d22302641760a2a30c14c",
            "b658df3cd06c1ed9eeeea77569f26799f3cb46025c07b0d66809ffac35e15d9baa750f936c3d5dac30beca26138c9340",
            "5a23b78b9c0c8a6845a4791212535d9493c1cd9155a9985e35470e983011546da08d735289d601b4790837a2f41d06ea",
            "e5487a64e89878403697d0101ea64d916152742fe5f33747f87c0c67536c15651dceea52e17ca517d8e9286c180bb051",
            "8e3fbf8b766da9407b85879e89518846225182f1562dd262f924285d2e8b553155f64222dbc84269ccbab83fbc161a51",
            "afe53555f5ac17d9b714b545981ebc62efe45dc8d166277f7859d21fbddcaa1e6e923f1a8da30d2d7273ec6155282c7e",
            "abd42eb6f42e9bca080761e5449dca8d0036b869941c32269be8350c1b789d3ed2d3749b2f78c51bc336a0aaf6a0bf9b",
            "a4c1331642a5760d4bfa80566032aa54fb5028e5cca6a06551b0c41883340e24515d17a3361b3a34b63786f79920a8c1",
            "df3d80ca4bb114da1996fe33fa68045fa28f4f0fe63478d884ad3414ca7b51b4b741e023ec683c510c55ee10ff93bb57",
            "21e229d1ad85b0fcd28178896cc90aeff4c0ca1d198bc1a83e8dae358879bb1b895d8952bcd7d58cb434857483918b68",
            "4558ad02a2bd6144a24725cc4a4d6e00dc18c8da840f6c34bb801ed8f0921369240ad8190eacff4dc27049c02c028d48",
            "ca5e43c026f26018b0c11156056206ed2d6d54cf964f85b2b125729346dba9a54456cee228823246292d5b7144b3a8fd",
            "e74cda0e0013d1e40586f8eada00a2b4a55404daa748f3a67d4ae4266384e264c34fecb6f3b8869fc12daa85098d5402",
            "d88685ad4b7159b56a07898952339015c7b866535958539fa81e589b618a2d8b598ce477a6d8b2ec6a5b0e9515819dc7",
            "8e3ed8ada8f472a43f0d614d5399566df8a8f0236196a22c592b1aa20588ce845e820ad4601d7d488a551ad7c6892bd5",
            "b48c0eee484b256229de86165f8704a40da678cb99da860df91c7641fcd92a8ee215bd18a6a4afc18854047122314c05",
            "52ef09586c2d5d521477a3468cb540631386a61b804a1f557a89d11f3b128892612f2d41d383d42a4b00dbd92013918c",
            "874b34ac32b47ef037d2fa63540419f9b130bb2972ee298623c52c93688adbd1a7786559788935b29d429eec54ea4898",
            "af4aff86102ed184b012511ca586462014fdef002ac1e2b896ec018df7da73dc9024107a0892c9b9a8e4662682d06954",
            "7359eb5f266630ab39391483d46293cd629d8917e3a95ac5581c224f2d2f3f0e3465ac6927e78cf08047e40d1775b985",
            "b79b50191d3bc3688fe77260e095ab0231e0820d927b7e7b476750b517a4404326f9674dc8aabaec9a1bb35815aa35ef",
            "68d314e2acca18ca59a3364b219c84254571a7481567931d9310ee09762a9e52e6e9b486d9975cf2128d899e1929a9d9",
            "d2690bb35cb299830d570b3b1532143bbaa3133bf9c39907611893ad564eab7a9fd92cedc5f8182a883133abcd604c16",
            "89d6bcd5ddb27c284e6262309ca9cccc59c826c54c146d7d8f0b86da574dd4c662199e2597f031431147d3c78ccc0f9d",
            "0a7e88f5087031cc9699cc676623af97a055220be491a527a4d30fb0f8c48ef2fe97735c09ce81491a772326d016aa03",
            "0dc914ce1c2602266952e0955dfa8a883097b602ab67f8c5ff36e292a64d53ee43793bf40036bb2b29de3a1580899897",
            "fe4b6bd7b6c8470479198c28bb798c2c138f116aaeffa90c6331cba2112dffcab2ea01b9a63c630c605afda7e6b40157",
            "b8a5ea4e6f78ed8aec9d26cc37bc5efb19233f998de2ac93c97aa453612d20875fe96c04885d5e61eb7fccf98ee53e9c",
            "1cb14aa6749a80edca20fb2ab46a47398064b9cc35dde53ea055ca8bbb7196e735fc7951e8154b17f5d23a896c47a476",
            "cccfd65a89a68ed8762df6b9f159672d5b1731cf304b0b10539c48ccba340a88750eb39d6af9200cf329a0d82b63475a",
            "09765603d85b0d6532314742350c7e2c0a98c24388a8413e4f5145dd22cbd96313db63152694df69d4dc294de16b7703",
            "854f8ada854345333a4a7e20db74ec041306a3422d9102a817808d2d59952745031b5b26b6531d918ef340a697d8999a",
            "8dede03ba156ee040c1a339f1382a52386a4bfa9f8f08407d8e109e51a82b191e894b43946a3038cc602a49e980789ea",
            "4a29ae2188268ec0c00a65d8808d376f53f20c684a95daf6ef7f8554895df2446d96ba7c639e82678a51ee49e449eaca",
            "6d84a22a0630a68ec44ba66200531a911a41f53ebaf37f483b20b184094c698c6fe98fb1440e7a73e450f4c841dbb6b1",
            "704bdc3859c9e14aecec096b84d4474640b5c1be8192b048f6842f06eca14981a205221037920a5b3bf181241aebb127",
            "8c81252b52052917b13b281555ab0a3122288187a646c4581a47f550ad0dd3d4886397c643482a6dc400ef8c01dee941",
            "f24f37e39024253e12251a57a0b235b15764a6005632d135d976edc1ad5e5fc1e15e823417e2378bf561346b8a8897ff",
            "61dab9d38360ee9e1ff0b15ca6789b1eb4378b47dc6b79904e8f99a055e56f16c5e348db67ae707b0823b42a0e603995",
            "ba422985831c940c0ba1b6c838337a8acf748962ac8a21cab86c3f5528286cbc35f7b4c44a5bf6b07b8081b6813139ae",
            "919af8c01c759278c925cd82b6b059309ed23058667929e35d69e455d4f9d192ff5365318e3abb444b9dd546b6ff1c08",
            "11ab0d02256ad9923468508784cf71ac0e8c18157cdda6e6eb8c183f63deb6cc0dad0ea5f3ac4b39400efbc4a42275ec",
            "219758156a2a30919a7930ae6dbe01c8b62e7ab42b667fcf3277ab23da1d9bac6fc9a70f891f64ec968b6dff8a3e4c19",
            "d32e2cc715cf502855974a1ce5e6325ad119cc6a96b3bbdd05eb3c2d449f06a4ddcaa7b5a47672c58f5b5336a76b4e31",
            "879e59c76b31e1b0bece0a767fb61a0fd93920cf7bddd0eea24da406807f427167cd00a9524299e41d48ad48260a0469",
            "4f818ff514f36ec936e12b4e48478e5c2bc46dc7e02f26d0de975310666831dbbc2ffb80d815b9c56000b9891ab5d952",
            "1373883694b9db11b688b302b58ee1584a45ec109e9a2f80d91a44301a6340d5987cbb09e644b0d71431b5b646ce4114",
            "50b5d86c8c92019668c900b8486004e741dd628aec8341c88286926e3a2e4005569032071c0bf9c487a639bf54cb9663",
            "bc52228442008fb2e50c2c8631909703eb54b2d5e6322d9a7176050d0efd06777143145d0c864b60cf136e2a923ba50e",
            "c4b1dd527b4bf7e56f21223682067a1e7a798cfb84cf6ee1ab81ba22a00632bbb69cf0afd993d4d8061e763f1f10ef20",
            "eba40010ef20647b76a9ddc884252c91600f04e4746200e9f4d83bc864ba0b5aa95513ff22188f268f877c67be55b632",
            "5eaed14271a0cf0418160c0a740607c713147ac162855ea0cd9f6397440627043aa33cdfccfa8fd99c34b085830ca59b",
            "f439107243d4062dddffa782fde04c6ec4e50dd05b27d716123f3701344ec0d0078907d8e0829705c27c9e15ec54568a",
            "4d0dd87f3da55789eac2caae433a7d9c9569ed93a08333733762398e028bbb1dda19978be3695093a1864fbd1744269c",
            "088f2c1083f49e499cc81308b2c3168c1ce314681f80af0f46531968e7ad8ca435d6d21a8de5b1d2a11a8c6ad4062323",
            "1d3a6314bfa951583888af99d2c8ab28ff534c1723de6bc39758aa2dfbb823713fa2198cf857324197b7c1df42d5fd82",
            "48c9b88908fd0bc4022c76ad92c2c7b2c998506c253013caa3641d90b99da620a6c56fe4db9f8869817242606cb40196",
            "2d3651b2260d0c7c8e6eb9b8ce3cf4cd518bade1cd51034b258a0dd2039f5f905b5113e5ad6cd6ffe98af21076495d02",
            "426dead6748632b88e1a78bb6690dc6b4a985d782c3e338b4285046b3ced14b51a8278c3be8668f6cf7389cb5b445750",
            "b7d3a64dddffa26d8403534a5cb1a234ae6fbb9241491c31a31326c67f0ae707e33daddad3ee09401ba91befb4209811",
            "b3a07809558c14d246c57f027fe48a30cdcd37f1c4a525577c7fc620506bc0f31414e9e51847648be9e0040672072b3d",
            "a7db40cfe00eec65e772d730624fa2475fe118ada0c4dfbd37f46e590c4cbd89e17e33e372bf813cf15cc6499f4b6b06",
            "118157753654b51971c87be3e8edb9a440c5f9637fe4262ec5f9e3813608c63beb5cba4463492fc2134a97c0cc7fb9eb",
            "08c2253a03d467cc15700c8308035879038bb284b3686e9662b7572d02767b15bbc9b8d52301bbbd6a15b0dbab34ad17",
            "844cebfc40af4a467d48af4a8e72f9a075a1321e6e7a55b2cd0296e2c159a0714140ef5cc6173ac8bb43883ba9c8fe0c",
            "798f946723550798a0a8612b55add3af03544c78b91745584a51805d2488b820c2976501c02e36dadd2e3fb849abc8c9",
            "6f7f5885a2cedb11efac64f722d27c408b59e101920581061be5298064b195466d6ac4b500aac64d9b4cf26c0c1fb171",
            "01ca53846ba13ea6b19287d199561a068c9548fc03e212a19d9588484f5e3a68a72becd0222f63b321c52779b166103d",
            "9fd2d53fd61ec069632e87526d7d5e8640a55d06b4316d31b3b0032324d53980b65c997808f3451d60ccc0279107eb01",
            "450a1f248fef470f2ba4b14ef20b142eafd4fb7b86ee0e8d4041949a5619d58d31675737b8b49c2d129300395d1ea19d",
            "42bce06032de4846077bbffdf1f34af92bfa7a1f5a3b3673d303f6124841638ccb920df2ceb7c11c63d55e978ae3da51",
            "db9b9be64ebcb0ff3be664ac139893315c45340a84252996cbac7469aef607a1beb9f6328f9b2b989323faeaf30d83f1",
            "3441d90329001cf28d4622730b000ee616407177dc60eccb0a0e4d9c0345f6d45de25038be061d9bc6142b55ff8f0959",
            "b5972764e1f55a72477c77cc32adb6fcc576aa0113bf02c0c42f790158b9b36b20758d1ea7081f491f8c3bd3cbd90a55",
            "5a7b69975bd1f665c50cdf1076007419e70ee2b90043300d78b87380a1fccc0e772ae81dd3578784743899de3bfce4fa",
            "ac6037bd355463e6d910b33613df8549a0d31b541d122d887bd7df90dd932048f126460229995e8c0159f488f5fcacb6",
            "0e3249c80d4ce06037eda4a9dd5fe30a24bc0f901a019168dfe67163b3bb31669b3713b6f0ee3834da35484e41e950b9",
            "74e2504c5b69cacd872c9311a88446ce9c514107185c94bdcb9da0f462d3ea02e6d384ceccd8f32f48f22d550a51a84e",
            "6e85a8e9dd44c3d64fcdee95f4ae196dc147bd077219119494325a8cba3f9040274feb0c038079c2b4680850b1472414",
            "2d5dece992dc262d2f2064b70b9a73917021935ea5b5343ef0d122846282c695907b0e8d4b3ab18b7e6d8141392112b0",
            "62a4515c3fe9e7d8718ff161136bd4d1f2a3a79a89b0e2d41940c32d7504e4cd2e742b3ae05d088a8d1a57d0e54381c2",
            "0d7b40ed7edf13e5a9e8ffd435de403a34f17b1fe9287a2bdd50d67d1de10dff31bbd3629d49d0616a4489a92f094dc3",
            "46fcda897a6c56dfffed1f706f3076f5c1f67d13b1bdded1e074440eb968dcf71b712f1f5e1bc11fe6fe5aa4773d90ab",
            "7c15f72e060a44b2515c637b76bd31629fa31046db5c15faa8d14137730f648914c5f0f6ec069c5aa593c56cf860a981",
            "6efb8e1a1d6a8e8d8579085bca89db15a69d2b9a73c9937d71c542cc1cc8206babaa5abefff0c4096c03b8a8da4b44f0",
            "0d02781f7820b2caf1af37320df552f5ff7e06fd31e892c1060f9950b0078773e12916ce811a38c1c2d9500ee758b80e",
            "b2b98be1f436a83147e257b2f00894fe21161e074dd06b8e94798e8553a1042ea2b0aec7003c18c03005c587920854f8",
            "26fb0a149f80fe32493cce0b903b45ceeb04cb944839e52c6c8110d4b1f068545788850b60371c62e14278104eb1702e",
            "5c053095860d30027259b804e51559381f85574e8db4373435d257bba746da754891e6a9a9913e39a188cf2d90ebbd1a",
            "4c0591be2a2f90d3e4c0ca8248bbda0a22fdd355a0dd3f2145fa9d8a320f1444da7ea420d296738ab0615a0436cbb448",
            "faf2699171ef60611bbc003ba745e0e99da60dcf094599a71479cf4d8bc0965b18a9cb521881bfa930527e476174f9a9",
            "a8f46ce82aa4f8908e7fc2799db05b51cea1c2485d470a23e37b8285af829570aa50c6873be04261a4bde2741a9e0bdf",
            "8692e99172baa647cadf3d3d02e7a1e9dafdf0104b9f0e463812ce5b0ebd8abca762f2a692bc67a6d3368e463fe7a647",
            "da0896080c9996086c6238bc004a58d888c6b78e85a7a0b67459647832e1218b8cb7f9f0140d1b72a0004eb1b088fae7",
            "4316ce853a3058e5f00210ad729aeba14411bf928579047f070b97c029e8b2caf830028e58e5b1c8865e6b04fe13d648",
            "bbce5923fd73c1aaee1ffce1388458f03e8ac1f62d62d5978dfab24d5f36e9338b7e264a7b2ab3bfbf2b11e0baf2bfe8",
            "ea72e96f88fde2d2f7196a482dd7c15f74bde85784abf97cb899e7e0cbf8f6bf0efcffd320c0354516d4677351bc095a",
            "7980a7119c55a54f430654d5d7d7d61736aeaaab32a14448078219fc44b80da55c2966c16478908770b921142bc2165e",
            "0f1270f028d8cfe0721a4af1db31e1f4f633f9c0732234f21cf7087040dfee45f55fcf8be8fdd328ae14b2483c4e85d3",
            "e48b7f4350e6c37245bd16f491cb790d9af82c04b51bf5675b82129e95a2b23c39ef5a94ae1ca5f3a0df12f4fb377806",
            "f6e59e412de1e0114825fdf51447fbed14fbfd90fda22691df5cf66b61bf35ecb709fd3e02a7121f81b7d173163dbf42",
            "cfefd1f367f4fc2dd189b0241141f11ce9ad6fa3f601e8d178e08ebc80daa007727b141b8f42789c07968e03dcb37884",
            "709bac804be250498ff3f43d4fdee35c17c8083aad91326888a6d39174b9aafa4a14f549cd349d9ea49b3067ed24dc96",
            "55a6e8fa12c87b60e5f0e098152983e21f4d978c709e863251cb75288f1c9f8a6ac5fda02c7b237f037b9f13ce375a95",
            "2f47910fe798cee2c785d3e7abd28f53a42770ce8ee0fabacf7739af6967f0cf8ec0bf91e035c6a75c1da607dcdff8db",
            "467a9d279096cec0787451c0ed7c188d89217cc3311e4f4c0798b65ccd2d28f5137c3560bcc039738186717b9f24f1b8",
            "e42ef4df533cc5575cef2181c6bbc2b923a5a94b00c070741038f048bf40e8bb84c3ff3fc2f0eb4b300f8d127e5f884a",
            "bb01ec8baa96981cf43d17f53e176e809a4693a3a880bee7a3de3f89f257ad4409acf4bd10f5fe29f4fe69841d76076e",
            "c7193e1a0e9c6a3281425d3f8dc7b5abeba5f14f86d30baa785c1b693faae77534566da4be5ebea2fc180a396c058e22",
            "331d1f3dbda69c509bbdb17efe62930847f82f133c7d02e1742efa7d92a71cee455e407d7b357f0b7f372c5d85db91c9",
            "c7a61b0f2f2bd2ded486e1f89093f12151bec99bd0c9467e85397e7d2758194f0932de1de2e2d527a77d9ab4bb8ba3e9",
            "3b487aadf29f6275bccaeb483e8a5fe531e53fa5a8434e4b2902d793c9513c3370b47fc7c3eba8f79fe1f184faba01f3",
            "001cff0cc3e04d8ce3e6c30f517ecc751f42bf98df9e20f3d2195e9e972c02e39fecb789fd7609321f7f93e033a6503c",
            "be3b4139beb8e48d3c4e87f1650fa188537c04be1fa1dfbd3c86af8ee4fb114a974430009787290adfd18cf3df86dee3",
            "f941395fec617193512b30bc78bed940e69bfe7e79f63ec1d1dfae1191f7727972197b817e44544ea47736a0b6e1d1c8",
            "473312ed235a0b1bcf444af10e1b850fd3c96dc0c5c088e3f700a78253ae279286d7c87b8a97f328e196eb51962dc38f",
            "d37585d3f12cdd293e5efb706adc463eaa8d7239af29cadda0f89fd0730284e7733c56cb37516c509683d3d5b0791ccf",
            "e7fb7229fd50be8925102c55e0f48f02bc2387f1db66f28b78be5529f74472907736e53bfc899442de1745bfc73c065f",
            "deeb4329d7af89d4fdc89248dd28df8ce87c8fc06c0b793353fde6191d892d8ead4707f8525ee5ac6b82729efc5f42ff",
            "af287f8ba5a633c508527785aabd8f00eddb8aa896e2f424be283a9ea7f1512da0fd8ddf5446f5a7b2bf2a63fa13c7de",
            "5c46de15a9e1e254124c74fdce1878051a1f032f4768a8ca6a0375f994f6ab63e205123f2f2a5ed9cf15e567c2fdfa34",
            "a440ad7cf9adab1bde84040ecb01f6cf93ab3c6f7a13de41343009f1a09f227ef5739e0fe74b66f9f0623fcdc5837db2",
            "08e750da042e8bcc9ff436db2c882de511188e28601f8ff11ef3cd7db9994cdedac9472474ca7b27c37b249d28c8726e",
            "9b4e861ff359939efeae64bf1d7a3a8f5e08cb9d543ec59fc1c8a7321d28538d47bd8825264ccf07f88824837f9b5d6e",
            "f2dbd2ec22b2954064ab081f50cfe7329f899ecfe5783aff517c1489a44a67739ecc66449ee3a3dfcbb3384d23d77b21",
            "ce3c1e5d4f99a29e138a7a8e70d1efe5d95b5dcf4e367f8754f3f7e384bf52fcc5b2a17a06c7e9ebb878e9d433b85c8f",
            "c8e6ef5cc5fcbd52c4b3279dbf81c46fe4b567ef0f4839188fb096b42f77a75e0bdf3ee439c5fca27ef7425877fb25c1",
            "c7dda41c84e6a49c43047f52d0db4364bc7f87f0e537245d1b49d79520e36b6622c35bf65b9e08acd5c528cf0546f740",
            "4ac2786c022dfc71aaf05a9b6eb05ce154e1354d772e261d91a732f0fb3f92724c0816992fda492ecaf91787f9d99f79",
            "99ffc9716abe1289c7b413a929f6bd92cfe3f17d2a5d2d9f51b8d572d946024f226007edd8ed7a1619013a3a7f60f8d0",
            "41cab9a06a0f1e8f95e9783c5626c61fb7ff25e36621e94e25cae3566360f21dfa45d21df9de6d90c7433d7eb1fda8dd",
            "5e799c2fd5cf18cec9f0373e5e7fd3f79f5ce2fda708a6bff14a8bc227bc5aa2fa34fcff26c88a338ab88f2f5e621cb5",
            "fafd5c2aeecf870cdafd8ee10301a77b8aa48324b9df0f24d1df23ecf74412eef9734957a2df45d009321f1ca8ff317d",
            "88a903a52b0ff3a98b29727b12844bd18188d2d0f1b85981c391f1c0783e190cc2c0e33a987a701932f5e03a692d03d7",
            "1a9fbe3fcb7af9987a07c6ab9dc9581fbb18c6978e6459ef5273f3ab79250fdfa8b0fb517b03fd5f3d7fa8b19ef2c9dc",
            "e481e6998d91f93789a69b0cc904af31c6a279275986534ca1bf25296a7eff549216bf4f17e8bcaae6dfc942847f8792",
            "62e701f93d86a729e99fe3aba902e5abb949b4df33497bea52b4e9f93e0447b680f3dd17e6ab99a41fce1870be0f53e4",
            "7e08a5627a3e90aa45cf72bf50eb11b0d23ea3f9cff0cfcd7f439de74e25ca78a4d5df580a976f15c922fd89fb324788",
            "4f07a144dcaf4752e57e35a4e17e15d306c727f5f40030a8618dcc4fe960972fcd20775668f50630e8659a8bee9d08e5",
            "6997b744a33c259da9f98556e9ca774b346a8ecfcfae7c3d4b3ef37a4608f1c64d0b3f9af4183f4ad20643afd8ee8f6d",
            "d3b90cdf88fe45f29f4893f1ab3c1de3d74af28d652d2d3c53eac054871a270c05fff00a8e7c4e901d138c438751e96c",
            "66d0c22811ae12a8e5b5245cff0d0a3c7f04d1a296de7e69bcbe144668f183cf0a2f460f8017e304edf6c5e28b85e8e5",
            "3bd3e5f11633f0489790efba8ccf66bc2f353e83919307c77faecc38c8237ba5cb8b8f87ff1ef0e139f4ca9433e30a96",
            "63142ebf1c5118182f074f3717c761ba69cb88c03549019796de934fe4a25e92af37235e3a6c2749032acd44ec45f791",
            "765603e5dc58ce12e1013e22efd0f70dec3d9673568e8bc43b593c594f1e27cb690ff07465ce44e03211b84a32b5e1a2",
            "e5d42aea7f1a851be53b04d79a8c0eec8419225217ce3b4531af9c184be4964c99cfb465610e1322df07c8f791ac48ff",
            "fd27f19b68ad00a2f60d140a63d8ca040eabf70afc1e28dc1cd4c34268846a4d2aa6632bbf8f50258e7728e26728e2ab",
            "a3e28d6c9dd619a71e0b7bbf22fc3e1e770268c7f83f6668f4a7ec25b97fe47514aa3f155c91f9ac690c915fb2643c13",
            "b3c97c46beebc8775bf6e0f16c8620f7946ccba7f1c543c2bf089e5d6939279aaf1e0adb4a21ce08d29135a8da14591f",
            "f9d7c8cdffec7c7025ca9f0cb601f06d460cbe514a2a16b4fa4d0b0fcfe4623c0c656bf353b53d82f2e11221b28e2cb1",
            "f872128ff9f9ce5c395f88a7a9d9fa632eadaf8dd47791d5375790e7244eb3fe68bd1dcf4fa5641e104939bb87c9f453",
            "938369a629478b5eae14ff8dd6fbf1dc22bb3a943d1d9a50ee726659d243019a6fcca0cdade3e7af50e4af08e7d7d247",
            "d5f813df3ea1457f115b14d5e68786ef72aeab8581fa3316dfda4690f5871c79dc4e9171fb907cc3f0c1f2bbea21ea6d",
            "c961d7729dc101d6079df1d60765be3852cdefc2650e8e3f2eb65e39fef8ff0e1f74c6e56fea71c36332118d097e5fc3",
            "c6bf7ad0f3ed856118ef7287cb78d7311c63dc4ef27d280ede69e11775efa584b33a2ee6e1587a7214435913239fc8eb",
            "f235b6e8798faecbd744adbfe31d61f54234146fa0d84681e2729d30117eceeba0101609743d5d7632d6290550ba4582",
            "09e17b049ff355f8d920fb0d93dd8631af61d1b87da5f1ba26669efa779877170c30ef2e12641c8c876f5863d8970b23",
            "86c2a7b4f66928f121560efa6ce5b42b2597c58ed1c023b45ca18f614d6f5feeca115af00db5dc95829cefcaf4bf524e",
            "90e1dd9941f4c711329f3947388c61e460f90bf17714572e1f9afe474bb32a768da875c0cb9b3f062af51a61a076c45f",
            "ff1647caf0abe53fdc5fd70b93e15a414e6f22e9bb46cafd7b6124e9df5cb2ba40be4b7223edb8bcfe8e27073e0d6960",
            "97cf57e2658238729d0cf7606c7f148e0420a726ab2723cdfbb3d1cbe3e9dff1f4f54be9df5a7a3bd6b777a646e38fdc",
            "1797b69d69ebdb4dc250f0c9904af78f82fc0bdc3bdf054730e8f734770641129757aca67e08eca2b1017b7f2547aab1",
            "b770e6ff09e4a322209f1501f9b008c8a745202db9b6839ed6adeac267173df868669ddfb3c1d50347a85307ae8439e6",
            "77040e9103ce89e52d683e909d6ea709790d704fd2787ca15a08dfa8e65bb71c16f212574139f5b6b4e4c46a7cbe91df",
            "c237409dcb8f1d411cf2633f4f6013b92d1c5fccbcef8c3073157a4904ec2df62ad1c3250b8bb924eccc667e805eb6cd",
            "f179ec2686f4e49eb41e0e6af099ccf26b8a2c4673e19685b63cf82a3f3c31fbd62aee5b55a5c64c119ffab59b1c1d45",
            "3df7f07714567163609fc9183ebb62a8ab32850acc46b3d9254cc287d44335a1d58d86ee0e697b8d5ddcd03c5268495d",
            "72e79c861d731649edf37479ba695c51b5cfdf5b66ef15c919e0798644db3c2e2975f86dc224d1b8b338106a5961c2c7",
            "7d85d12176c877bbada7271b9f39153d3b8acf278e4e1db565c142bbb9bcce30a6e25aae6a3e715b90963caeaa977b69",
            "ebb19d70bba745e46ed959d11d946a7c81f220f11853fe44a5c11590b6dd86cf1b4fe625c7d6e3cb0bc4ca3d01efb5be",
            "a043dfb7f3259dcb99b970922774d5b6270b44e1b8cbbb5d97eedab1d1e1f5b88ef5862ae86976d89e5aedeadd50ff40",
            "c0d7e9ff43cbf7bfb5753ef74d1767ecad727ce7c7f3f4cdf5fa967abdbb5e7f9fb9c726504708f765a5ec703bb22497",
            "bbd2e79612df1b7d43768d6067fed94c368bc552b0403c04b30e8375e661b02d28ba0f661f86a2e2c330e32e5b7e69d3",
            "f50b0aeca265f175355250d44db48bdbfc8e7378ff9f717b45e59ad70db5e275af76571b8ee718ad3f4bb47de368a6d1",
            "b9c93063b3c1387ed4174b3e9735f37543c1a8cae5a103d3fa0a67796f4d35161fcd323a6e9bb52aebea733b2ca2035c",
            "cf2f4899dbcbdf99e2cff97a9529d970619a68b53cb8437738590719cfd88665d846641c37ae48b795a7af15d6ce482f",
            "2a4a2f9a965e553ffbefdf7ff1f98ca20733f847975e7b2ee957c33226702e7dc1d4732b0b44a3adffe5e25011f7c078",
            "7df9da1d5f69da9c19282787378c77882eae672ddec0592a368f3f842f236a69026be907a2bb098c5de55b4eeb96667f",
            "bcc25024e408c6e342bd3081ff6815d766209789482d9b82a6e3295fd9104abe6fc577cbc44f1b2e646625bccfaf48dc",
            "c1bf61dfc2ffdd9ab28b1773f481b95bf88f9276f1ffd8ce17d5676de1d767dcca7f33e39de47f1c4c16cb6fe8d4df93",
            "bc3a793edf987c360ffde424bf9364caab852f9ebee614760d12f289aeae9c51a7b04710f0f8df87f2d47f3c33befadc",
            "babbc697f35f5d37676fcff88c61d2b9d517ef7a4f67721588179a0bc416f735b5f9678537f6c20ae674ea8bc1bf6c2a",
            "ecd73d268862f72461e9e87569749fe35f77a5b58abbd27edafc68d2fd493afdf38fa6890b8dfdc2a3693f7aa4304f7f",
            "f431c10a3f48bb7bcc1e6ed7a8c05c7bd1527ed7a8f6c299e5fad402f1abeef3a3fa841950a66b2f3c3feae06afd357a",
            "c394d969fa1d7611bcbdc44187f4a6b9b5e0b979be730ff7be72da50f695c0e9d73fe61f7840b7f7fe4d594f55dcdc94",
            "34fe58bf2db52efdb1c4e11377bd74e407f385cd2fbdb2f9a5baa48ff6f87d3772f72ebc3bb7c5e7ade5cab9ce75ed6f",
            "34dff333deeda97b3870e1c7bb334fee7e66d29beb337ac7cd0fdc97e47efeef8f8c2b2beb3af5933d08e0c3c3bbca84",
            "f1dc32835d5c56e1ebfca4dd5df5dcc42f58f9d4a4e9a31dfc59dfe8a59c05c49474ddecbc29c3ee7fb461f4e1d623bb",
            "df7a63d14349e6abfa0e7a52e6bf723ab1e787dc28bb385ad7ef3a7983f0ca8ca5a8d83379bd869fc053882b544cfcc6",
            "6b95c34c8712cac61aee3bd9c18d4c78d4f16c82d1f1a9edc39f958eaafbf1cf4ecd4fad78fbf654ee76bdc3fff477c1",
            "cffc1e86c829f85223d451df188b2af0a5a36ee226aa5c2a106d66734a570b7393624ed3258b06e2a1c02ffa6e443342",
            "ca26abf910e651cb921fb0a66e38cf557365b02abdccc697e90229feefeb6f72f8134fd8c5a434788bdfc4bdd5d0e2f2",
            "bafc0b85df0a6f9be09b05224c35bfbbbb0a33ae9eb494595b920518b323993a59716d4dc4cc72476aca7ce3b8e5f754",
            "6cb36f9fd023687aace2f449a213cd25929fb9465ce4d3b5afd93a4b3ea0ef2007f427f6fcc406f7195dfcafdafcb0c2",
            "34a9c65952e4ace2deabd42556592c5b7eddbcd1b830d3363f539fd2977e7b52d1c83d497d39b727cdb835c9d8e8dbb9",
            "b6db7722bf78d6acd02ddb2689db46cda872543a2bf52f15cdfc7e86d5563463e64baf8d2cfe7e62a3bf73dbeaed19b7",
            "cc0a25c2bdd8db745f651dde287fc79c927243c443771df630f4eac88a571ba6fdaf3525f1b60596d7125e9df77a5a72",
            "d546dbab6995691d0b13b3d72cb9aef274729ef1e6b2d7e63a5fcd7d6dd272e794d4b5c39caf26174d1956f5a43539f1",
            "f6e4eadb856f26cfbb5dd8fe5ab2832b211e0b4af8a07b0e71f7ee03e6ee7d096077efd803372f8ee74aedc65abfc7b0",
            "c6d35ee1c3eedd42f4fe021f128eec9f37927de3a9c69b1e18df6fc0bbc31bbdded70d93df5de1e80d6ce8a5ae7b2a32",
            "7beeefa57bc3b7a5dc1f4c6b5a622b1f7fdd463c9570c611d85b039a7139e862feaaee0a8442dc4de2e39ef635cd3797",
            "55725ba7272786d2b618f2bfece858d071ee1657dfcec2aabea75c2541c9f4f564c757fbb38352269a9382eb86a16968",
            "929f4f9937ac403c3ec2e53f81a6a8deb6859edefbf2efd95d20d676f465e209c8316f189e7f027d868a1c9fef50477b",
            "28b7dae5dd1690968cc51e90b69ce67af46ce217b68ae38500f15d3fb65ef06ffd59539679726dbbb8134dfbbd3f468c",
            "0d46548bef3e731c0b03b7e5a2e9d782a580a6c7aa7eb01b7b33d4a5dbc5897776a189a57763a5a1aff7a4c1daffbaa1",
            "d1f0e2f726274de07a322625f036de746c1562eecd2e775f53a3b1a5d965f47ebf27c1de7b5521ffe5f70de7b25f9965",
            "173f3056c3cc0ae323497d29f6d167e764eed1bf9fdbbbe1eb7ce17bbadd49fd90f446724ed25e43c7d85752f374df99",
            "f28fc3f640cfcce1be32fd6cfdcf6ffc7dda1f1b765fd7d296915e3e2e237dfeb879e793de474cd2642c9b9e2e5e2fa6",
            "84ca74d3d3fd3717aed64be76bd3efdb372661d7e8c0235795198a0c3d7515648231b58c3e2d7e15490dbdf95d33f37b",
            "84c0f9d4b2bcf6c2bc3127f527f5e2c9d3f36eb9bbed2add1cbbe87fe8dda9bd8467fa7559fbe67c32cf67f95ae38d3d",
            "755f5cd0e10cea2e22beb165bfad2894b4fbdb6f1cccb8a0cba87b76df7137f63757e65bd5573dacea85056d4287a1b1",
            "3bd821f5794c79abae5a35f5c9f9213a4b05370baf7ce5da83a9d7dfdf20790fe94eb9ebb08690df1ab071d23bee3da9",
            "39cbee4bfde932292bcf38e10b4d5ddec085d4b1dc54f1cd9b7f559a10e2ffa24f4bde3ce2f3239a37f6e2abaa1bfd16",
            "735ec7fb2f7f4e6a5de773ffe87aeec81f466670ae17f3ae5dbd87c302c6dc4eff5b8230bf87bbb06183e7a2af33f08b",
            "5b1eb8717c42259fc16f7ef78655631f1226086787e9a6f6afe9eaffc0db7fcdb1745f5b7efaab3f3c6a5cbdf9ec76fb",
            "91ca7bf45ff11cd4fbeb4ca65d37f56f6b9ff8e6d4ca9359f7e5f78eefe2bd27b2a98b9b1b270c7b37c9d9b9b8e7e10f",
            "9efde3ebf6fec7f6cfba2343e0844efd07f77da4cb48eb87b4f90913dee898719ab7c2f7b9f37d7f7a66a11ec92adf5b",
            "7077d2da69e5199bb317ff993ff6f6def96f26f52c107fe3fd3430efe2ef26fc3929af2e33e5e0887b4724faa61eb963",
            "e4ec6f3e11ff6e461f07216ab1085f0ae50589ddf45ca5eb30279ab09b1424cb53817b53a5ab9cbaa74b334e9d2b5a85",
            "c6c4e76081fe9ab4e45cc48313b664ddbbccdd5a4e3cff0574f748bdd4e71f7704f6f552af42fc3e21394bb7163b1a9b",
            "b4c1e37bb13300075a3abdfcfddcf3994482cf77b50643927f85df1394b6a7e90ac42d49d5f83aa377ef486a905a7ca1",
            "e53027d1d8b7c766271e66beb565f6570b77966d9d93931ebaadcb6b18514ebb5618a5f611a3beb92545c8d875ccb26b",
            "0cf6eba30f6db5858ac8d5ae378bf3ea8594131577addf35ac6a8ab4d990212c9e91e15cd198714b4f687ae5bda18da6",
            "7953cdbd362392135727265418b3122a4755246cea49a59ea3f1f4988bd487f9ab7a8f764ab6abcda7ecbb568f0cd6ed",
            "58fd5ca5ef6f9d70a39f7f6ff6cf4f64c1e2d7774e3e9b6914aff72129677b67a05a4a58046e240c9fcb6f17f202c303",
            "cd52abcf2f9da9975cf7edba7d83ffd8cc9ec0472d37178801e3b6c9fcf1e2cd81b213eb2a133607b66cdc1cd81c98df",
            "fa9e78eb24c4770e3dee7ad805efcf5f3aee4634113ed0d198d85bf6a7869b7b1e9ddf5a99672a0c767c50f85aef970e",
            "b68a12eabefbfc5f4ff947b3efe6973b7f577f9fb0a5fbfeb496619bbf3f3e75d7d9116747349d69bdc939dffed12f46",
            "229ae5b71fcdfdf0dcd1dcfea3b9b7fca6523c9afb8ddc9d3786cadca33f5e3daa6759c628f18d65c30c19a3323c73fa",
            "7deeabedabbb1fdb7975c8ef49dde012bb1fbc1593f705c9d1ee6e4edb16dab3a3fef321c65b8f56e795871a8765cebf",
            "bfe2b99cc6ee8e5e09bb6c331dd3ad18d3776c1cbde86ef78bdcac279a9ae6dcc189df9c939f6817cf6f836523bfd0fa",
            "5e42e18f5267b62e9cdb5fab6b302dd13526dd3af7d5ec5379dfd10919abff71103be4acf31f405aa76ec4c17ded6b7e",
            "713d3067a3526fb00a3b20ed5991b73aed3b239cab732feec704bfa9d3fff2fe09c2552631f5833b6a81b9b5c3176d53",
            "97aa1791c0314168e2b6ee41b2ebe8ef5c955c36da1f0ce96f9fe8081d39b1416a589f5658651d5d6efa9b694da6f9a5",
            "117f5bfee0e8e774a643cf1fd5b53fa983af0d5ff952a67d697a880a94fe3fa56e7e709498f2b8f9013e7bea7bf31b0d",
            "936e3ab5a065cf531feffec7dc61fa0773e675605f6a17865f5df2f3be4691ecad38d3e7755d6c96bc7bf2ee9e79eff8",
            "dae35b6697bd71f43b89fd63c772d5c71f4b6d769520a5b72fe1dd9752b949e2e66643403f39b9293bef656e59decb48",
            "87aef0bbda5bda1a73fdae9683d3fbd310d73bfbcb29d3e79d2c148b36ef6c5fafaf7db1accad5f783da179daf97e56c",
            "7ddb52ea17dfc9361adfe02f58fc29f64f44cbd905871ebd90fd85aa55bb8f4ebdf17b533b2cfbf5d683fa83fabce2e0",
            "dfa4bc62f1b5ba5f1ba66667bca17ba37fd7ee9327dd1f0ccbe6847d739fcefcc68e8f933f372dc1d721ae5ab572d2a9",
            "11e74bfbcb5fdf6b7fc432ecea8ac24a7eceb8e3297919d9f50f77243c3ab9fcc12acb13c0dc46b3aba811df1289d096",
            "b05b24f77135b4607f588da825aacb113289035dd0bc525be0eee116510f68bdcc7b9237cbdd4479c826bebc30d1b048",
            "f0e45c451c83f5180a1813b10d70674a266624c27869836b91b7932f81f1de004cb254a68e0d11b6125a6b7384d28d78",
            "cf2b88fe1035661985114bb7adb7a366dc69bac36ceea5827488df927efb2df3a160395c2dd9ccb795fc7c89f5e1129d",
            "1dac9b12ac85e5dfab37f1c6a5a6c50e53d5b5d2a2f2fefe09e0fced047ecaeb136eff16b4f74fdcf5878907bef2f6e3",
            "992f2361f0aa5fb39d3855fe726c96a8e8eef58a4d645771e2ff6cf9fdf25e26d5785fdf3941cc9fc64f15ad95fa1657",
            "885d41e74afebcc39dca4fba3f6dabbe26ab29b972cbf2503652a6ad3343d772cfe73a0201cf9af67589a1007f57fe28",
            "e26fad4a0cb61c1e8538575357bb7074f7b776d33bc32fb677741af7e48f3f78f7afbef0e5bb7ffd618a5d746e72beeb",
            "16da5a2694951d30ee293b603ef0eb5b93ef75fa7e51feb4f4de3d4eebc47b1fd057b4563da9b77d4d8f84a957a6373e",
            "d526b5cf7ca2dc7edb11c7f6e2b342deaec5df7cc434514c9ab535217fc347071f46bac7ade284c2ef7cfbbd719b75ba",
            "fb53cd89de9dbf5b2a0e3bb833d07ff4f9835fe27f9b306cc3dbf9badb27272e2d7b7f4cd99dcb32f85059e96f43f512",
            "200916cb010d2299a966944a5d5f6a5a786a8e0f8cbfe82910efaef947ff1c57ff9cd5dbf65ac697356cbb38f33cecec",
            "9f537a6b59637bef988399b6845fb5645fd4f7cf916c4773560e3ba2efcfdb89fffce3ca8a77251fcfbce3aefeb2e283",
            "77647df196fc9bb67fb465ff9809e9fdb5b3ceec393aed8d7b43fab299aefd21ee6867f6e62dcd8737dffe787a9de07d",
            "74c6d917bf7c3e6107773ea1efebf6650f9d4ff8c9f90444c75fdcfc4563d9c2de11677f396c1bf427e0bfc2b3396773",
            "5e3b3cf67c8271f5215ffbd99cbc86c453773a0b0ff512ff699f48c1e2a74e8e393966eef5478d478d7f1a9f7ab2d015",
            "3c9f80ff6a9fdc153afabcf17cc2b984ea4eef2f4a3ffeede2ddf65d57f9335e681a97f1c2db9bfb535ef01d7eeb44de",
            "759986afcfd78daefd913dfb6dfbe633e3ff27efba31d72df7fdf5b1bd5e29f88361ef7c676b7fc24b7cffdddcdcdc23",
            "c963b9e207ae4d7894affd36f7950feba575126270c2dd6f7f470a909bb43f91dceb9b27a4ec3dfc9394b2e4b2e41dc1",
            "6fef5ef93cb7d8e7165d673fb18aa77735642c9d3afdb1398fc3f863d54fbc8b15b8ddfa0793b3eda71e0f04778f73fb",
            "d69df55ffda8feec931f7eef90bee1217d59d22cdf338167438273bb70e1c5f9addcfd82e9bb804d7cd8c207be75cb65",
            "cfd1b47e517213be80069d80e66ff4e85ad6428755db897e26e134f9e212e9e5006ce4aec7b6c087120e9533f7f9d748",
            "0f9a13738457a50db0c9c52d6da27716b4a5654d1d21a2e4bb5b404494d444cc8fbc393573d205d1586b149625de0a75",
            "fab732c5a9a2039e45e3544edd596f9d8b954d4ab2dc96cf198d0ea3b1c09a6b369af36e8495887339b65c47dcbee936",
            "baea30bd85925bb8cec86d9d87859dd8f2c6e706b8f4722a606d4f93a5aed0b776a4ede0f70bc3042f71ec3aac89f86e",
            "dc59d693bb88f7f54ca4a49a89177e7b92e46b6af18d7d671083e493f4a59eeda384117d59c3099b2cefaa285852b873",
            "ccd6d49da3ed62f9de358e503df1b87eec6edb6421a931332475796b7216bb2c98756e4ddd96bdd86a424cb37c5825d4",
            "f92b434cf4aa34181075cf5b856da5cf0e9f72e728c74badc7b3320e5a8484ca5bc4de3cc23bbb6d3dbdd9442413136f",
            "b85f48797139e299bdcb4dafe7649b21d558da93e6f0fb3f4402d73197e9fb232b17f67e6958df533046f4948fb68a8d",
            "bdc407a269de1d4e83de91d4c0ff863a0424ee0e6b920a440bb67f9e48e6938cf54f184bb7b72dee69db7ebd5499fb6c",
            "4fb2a06f00fd37dc23f40e4e7ffba9cac2e4eae17a4429f957a7bd92e10ce96f9d80d76e9fe95e97b36287eec3c58726",
            "cd9aac3feabaf778fe1a3e61a68e835b368d2f12fa8b3f2da676d5fecab4aa4962f77b5b6098907e5c78794cff5b9fcb",
            "d7ad1df3f575c70d23d7d6751ffcc2697edafc80effd2f98f27609faef6ddfd9bbc47597f55dee96bb3c0efe53db023e",
            "7ffab9bef777b78c3ab8ab6dca972a139a53ace553774d2c157519bffffa5edbb9aea37bdb7b938eee7dabf8fcd15bbf",
            "756bcb7bfc4967d30f278cd18dfee192bee00dbbbea1b799f3ffe774f95167877bda9ff7de96e69df1648573c7b06f1d",
            "c92a2915363b45a9d5be1d7e9557fcb8fbd455dec0fea57f9ce5cee8afd7e72f7db02189f767e9f83f19a667ce7505a7",
            "653e91d96fe8374c12ffde6fe0f6da3252b27cf8ffff357cfcb96271f8e624ff38f7b523fb2737ae3f9f93f3ec55015b",
            "c658e9236fc0b63aa7ec4cb125f9ecec53078f1e7af6d0327739bd29b8fcf30bcf78dadd4fec19bf5f32389aabcb02d3",
            "1a0c6533c672bf3c541e9a1f3021d14ffca6c9af378b658743015d965460d89dbaed812aef7301e94782b1fba4705270",
            "f70a9f14cfb85f7f5aa84d08a677f0e76a3bf08dc36efd0c61f777e627b4ae5e39f609b2f3a7c6d3b2bc79a5576a1c26",
            "f5d20b7ed70ffbe4819bbd878face7de68fc43ce374c557a71f307babb933246668cecff5efff7e0e9499f8aef2111ea",
            "cba91d599bf427337f90394a9f522985fa1ec86cf5f86f09343f724726be11d532e2571dd21ddcd9a75cedcdd97f5db3",
            "fabbfdcd8e9ee10ddf7597fad6def554d977f28ffdaa25eba8e3ee04bbb86bcdf923e78ff8f6fef448c1f9239b9ff57d",
            "baeefc11ef5d9d7f34f0d758c6584ac5ebdffed478cf9aa4cf7dfe86450593f61d1b53dff86981ffad99de7b747b1b0f",
            "ea3eb4ecd3ddfffeb8afe92a8cc6f31ded09afddfd44d34dd3ece2d934bbf8e3b4fe5da36b257fb3576f4bfae30bb70f",
            "3b9fd75fb5fc83a49f161e1a79d075d095e03f9a902efeea858c16bee5d127629db1666211895f891d8832b929c0191a",
            "99035b03628fdc739837e662e7b59b306bd4bed4d0a47628dd4c65b1ace7b8498b80485eca8b0045c2a5b951b27b7ac4",
            "76ce11b98ccb4e6ac8fc7606bc98f38e704bba648f5cf06d35ce30176e996b13868fbc1b0b69bb654dcfa19b48969a43",
            "93a2a4b3ad998d6d1e607773fa746b0b7bbe4c8c5a595b0e97da8d5beea3d7606293968454226cd20aa5d8278be49e17",
            "e8db203511eb153f16928c21ba75c928147e19b1fe2f6ff945a8d61b086d4fb708a9b72c8405986b6d2fdf642b4fc4f7",
            "c441a09cead4af8c337f69db3067da12ae7a59da626b55ca9d7b164a5bf39c3ec4ceca7dcf981699cdb7996a5c609af5",
            "6ac2b8c59cb1eafa3b47bc988b84c23f6cea5b503f9a5b545a9d985475fca64fcdd98663653ac3f109074ab38e83edaf",
            "9b274b3669f2cb334d877b57735d5c652dbd8b31d89f5ff062524ad10fda3e3686ee389797777b5fe6ede4b6b463f579",
            "b7e7fd619f6bbb70bb7bfa2ff7148841814be175f4bebd8b689a1472f36e7ba9b677f44de30e3fbf0def7df9f6b6d297",
            "f63cdbb533df56d3e77abf5e2c301ab99f4cbdb1c676eea6493a6b6bcab1f4fb7fc9bff8950629a89b609f7267c2a2d6",
            "fc8a7b4df8bebe712d1599e5f7547e3d7f548158d08b04ea5552c0b9f158c14f26dc6b175bbeb4253df040ced97b1f3b",
            "736fc3b2c595fb2a933f770c3eb9cebc79ffe4af6ede7feead07e66ddeff7393f1b569eefd67efb37fed679bf7f76fde",
            "2f1ea838e8766ede7fa8e9deabfb8b3f69e9de3d3127f36cd2fed639d73e3a75f32f7d13c6ffbaaa3fab59df9178b2a4",
            "7debc99d3fd85921adf1b4db2acf1167bb9b77f627fc2561737fe9ea9b56df74ecce3747cca9ffb8d8d9bfb1e4cfc533",
            "848a75673b6774bfd5f926fff11faf3acbbfc9db8e768dfceaf96bdfbcebf78746cc7ad0f1cde492a225df48fee4c6fe",
            "4f461cd1fd7af2b1b4bdb3cfaeadd8ff77b370f0c12f3d48afd4ea3f3b76c58367c7968b95a93f101e4dfec5c1fe5d1b",
            "3e3afccde4dd8e83d7bcf3d56dc3368f2f0b346c1edfa49f9e9a31e5d41a47671b9acc77732d6556bb95fa8fc6fe38f8",
            "6e8b750db97a6a8f65942875651607024774fd873b02dff53deb7b767defb140f3b36f3afefef3178ece69bf30e2d1b4",
            "fbd2fe34ea24e73bc63509bb66ef9a3df350f3f927c5bf4f6fcd7bacf06f33c62fdf9f757fd28c745759139652de361a",
            "47d4ff65efaba64945a59b273cc95953fefc76ff9bb3bf35fa54c68b89e757bc9775ae64baa17fba617782bb6458da87",
            "49078719179fcf3df88bdca36337af5af2dee9959d812d2fe7d9266d7de5b8a17634b6ad59fbefe7d7e876674d4f3626",
            "5bf97339279d19af39b72fba30bd418f85963bbf6b75fd78664f21d20e423fff43bb3fa5c576dce03c6a3d6a5df5db72",
            "ea75ba2239f07e8ffeb8ae40f4f441dfc3017d7fc6889345cbff3e417834b33ea13cfb8745b0eb31915ea600aa1ba098",
            "a75b70314fb73aa48c6d239776472e2e6f6497761b30b71b87799d48d4417e8468ff10e1a5619c437f43ea8fe0f53a72",
            "f3c14c1bd88ce1dd07e54c639bea17e96d4f632b1b7291ca91e3e3664966f11c5d03d1f353127ec8b17b3d6ce6c3c252",
            "228f7d2dc0057ae97e035d3154eb444f28992ccc7660c964c770dd38ee9e0690982fe65080fa62de99dd4b1767770ee7",
            "b39b8cd8e7f2adf3f4774a376eade1a608390ec8d9b33e3da781e372d8d53d67107b4a78213dc7093949f6438423a7cd",
            "a9760c73edf6630fc7ce1d7da6aa8b888b0b6589992fafd1f1bd4d099b5e6e3aba96381b2e41128a70d5ce7bb6ded680",
            "84f48e9ef5429e90543de7adc7ee99c8c1cd6c26c80e1846a725f7d89194cc957f8edde2f4dfcf15ff84fafbfbdfe400",
            "d6f02bd17ffbf98be8ff8ff513a18e5b233c84c207f933e8fb4941283fc89f43a15feaebb00f4e18c38fd4630f746b84",
            "43fdd8c3caf748cedfea758073e1b8e7049c720d8fb74bd152bea6c7b5b5a2da5ee6b1cf3e24bc02fe6e45df1febcd40",
            "6bf82ef70b6e7dc23abe994180213944a04a23501d20504552b69394bb555018c250e0d4bb091c1d8af6a529da2797e4",
            "67751e62751e50a4c425ece54fa06f9ca2465557724c5d35a43e1b69b7b245b417d4b0efd097f4530f866bf829aa1c7e",
            "92e354b8e61d7a13fabe8da46cbee4589dd218357ec8a3768a8c1aae379bd4ab1c4b0cf787bc3c8a2ff33f578ce81432",
            "a253e010a95daef321459dcf93ffe4f4181337f0ae30568a70017e87ca736a60ac180eb12d56089f2e129c3c47da7875",
            "f87d6238f410797f887c1f20dfbbc9770df97e0a7d0b407be816e6672593fdca103e8cbe3fd5e371c8830b7afb19791c",
            "fb75532113f691702d77047d5fd0c33bf8ed01c5db03e42dc6fe4ff56de87b1ad472b8be31688473d1ef479c48fefb94",
            "f46a211aa7b38201aed125c19fb9728e83af92543fd37bf5002f0838fc1b8477b8ce230cef70d974e4ae4590c6424847",
            "42aead8dd586db337d10b53d279c53944ddb3798b22d4328fb40b887ada88ef56b645ad0eae14796e0b7da3d1ce9d312",
            "060986c836a83e35c4f4e91a1efb05be48c665b62552b6b2954583287b0dd96dfa07069356293306554a0de9255cca4c",
            "d02e6754dc72f45ca49c7242c9b1b98be3e43ea35f46cee32be1a0bd12af3db306dd1e5a0ec4694fc910cbe1e294337b",
            "88e5f071ca291de238e5c52967cea0ca192ac5cd1d02c54560b4c781f1ea2bd4f723af50df970db11c214e39e557a85d",
            "8e2bd4ae8a2bd4aeca4195138ff6e3718e33fa919c92f6178539b513e439ab2a0cd159a11a95300f9590cb17b012ded3",
            "a7c377c33336a728eb4018076b4086633ec825c9b084504e5ad2693dc6dd05406702cadb17a2ff0e303e74f1d3580870",
            "5c2c04bf43520f2e11bffd8d3e41c5f9d7f0d984eef0e1a1c789bc426594f5e17df4063070584628d190507234e2b4a4",
            "925c3292b87e81c92e47c8f701f21d22df542af939cbd105b25482e1dbcf6321e545fe0281ef7a32aa5896b9a04fd2e1",
            "b7387ebd708af5cc87aa1e2d213dfa61ff7f3ff13fd8fff2f7d1f32a7af0bef5d310b985e053f6fb1b14f7a122fe4a7f",
            "7e8dcafe2da611459d57f2e35f06ccc7693968ffb2d73733dcdb2c63ef68f61b7699a00a0b90190e639fc6005bd0f391",
            "22b14b4126a74744c28edc48f8dd5191f09e3191b0756c24bc7e5824fcd838052c8af21d8af2d72bc22e45991f6528ea",
            "5594f3ae5151e644455d9314f1a648989b12092fd5cb211ec68f57d495a848a388ef51841f5384df5584b314f05815e1",
            "f58af06145b84f117e777224ec9aa080c11c09ef5184fb1461c754455e45f8a3708887ac698af48af07a45f8b0229c35",
            "2312b64e57c0a308f728c28f29c2ef2ac30a78f628dad5a708bbf2146914e13e4538cba2805f115eaf081f56844f2bc2",
            "9c3512b62ac22e45788f22dca7087fa4088fb745c24b15e1c38a704724a8a0391d2c9da948af089f5684b9e248d8a108",
            "af57843f52f4c961457c9662badb335b01a722fc91a2cf4f972aea9da3a857117629c2d630bdf3b04711ff9822fc6e51",
            "243c7e5624dca3083fa60867cd55c42bc21fd923e1f58af8ac72451a45bd7b1c8a7214e1d38af0478af0d20a45f98af0",
            "1e45f8314538ab32127628c2eb15e13e45787c59247c5811ce52184d1c8af07a45f8b0227c5a11e6aa2261ab22bc5e11",
            "dea308f729c259d58a7845f8234578fc3c45998a702f93ab9cec770d8f678f97f458bea1ba0a9e14a806f4910e4b5faf",
            "f01f907435e1f82785d4f2833cb6f17cc4617b0f96267bd16fa286548b53c932f121527e3b608b55322aaba35f96a068",
            "fe1422d9e254d1d2138efb3367e004d847729de4d608581e8dd8b8f07fb486df206913a7c1cfcb7c31b19b8961d8ffa8",
            "d3e1db8e487bcb70db8845b680c094ca6c1f0bb9f875508bdb05ace033fb9c32cf3e6e877e25fa7d87948a5baab68362",
            "2bd8453d4ec9c15e456e970e97fdb14e4fac7775a4ef13109c2e94a7bc5f6d6f2530f553db277ef79cd0d11f6dfdc530",
            "5c43acbfc57cc4b2b628aa179a590b65e8eac2f09c8cb1bd26a8ea555b1365dc12a197e0d5e5d81223ec2e211c6a23ef",
            "a9159cda0fcbfb65eba240c6988782349ab62955c6e75879fd3d326240be13481f602da88f8c612e840c17f5d826782a",
            "91a6f912795bcd5fd41bd0efdb8974ac23b117b0de83627955ec1914fb2b142ba8627b51ecef51ac4e15fb108afd338a",
            "d5ab6277a3d8bf255208e5580ae941be0b0ba4e43e311ad3caef23714ff2e5a4855842dc4b2897b60163f622828325a4",
            "bd1d9f625a35305ac5ad1d1e43ab9280531d2518f4327f428f732bcbe558b9af9272af57f42968f429ad650428fbb461",
            "10ed89573bcf6a37f203b52ad6fa71a956e958b91b072c3777c8e5ea59b9ff4bcabdf612bd7541ff36f61cafc3758de3",
            "a98deef2fb2a81d5bd4418a84db13608ad364d51948b9705fde1fff0bce1021baa072b2e6bf82d9768a3b29453e0fdff",
            "dafb1af0b88a23c19e37a35fff3032c608d92163d91099d8ce48966d1930d1bf65f08fb0842dc0c11a494fd6c068469e",
            "1919398091219778b95cce4e08f16dfeec24ece9086105c9b2be4040103eaf2ed904c1e5b2c6c926cacfe594dbecae7c",
            "9b5bc46d4e6fbbaabb5ff7fb936624b30949c6d69bf7fa755757575757555757f7f8991ffef3c8c1bfc88116f9c8b77c",
            "1f08c0f8b2d3a0cd85066d260d9ef74369378c975a302e31318e23c637d972ce86d56cb55d6da96db559db2ecc3960cb",
            "395b6d90ce7a6d649af51ac89812171d0bb9848eed9b661a10600114effe774272eb7f26e7d34adffa321eedd3284197",
            "135582662a0344cdc544f857decc594135cb5f6937e6423af417abf55de49bda0896803426cb35f2ac36644bf3d1b413",
            "661a2b7b352d3b80ad3d9a2b5b2bdfbfdb025bd6e70d5b62cd74f1c5c0656843dd91274769d022a7bf65b6e731be6a09",
            "35b6e758730be93be5c83de4925b48cb6bf3ecb94fb8e4f6f3dcad8edc0398fb9c06fe37d136b7367c344f685e565250",
            "cdad05671d79871c7905fe3f77e43de1c82bb06738aa7907941e01f3e0bb9a81293032a70250fa651c176c04bd183841",
            "67786c0459f92e847c172f10f04fa23d497b0e6b6cc5be86f632bb78c8bc3b61def54d3b6d5a487b1dc7fc4ae2a7fcd3",
            "6e8005db87b2a5a340d2fbe3d82280dfed3b84b267391d2943fcb98d4b0a5957f334f49c578dbff60df996d07bc87521",
            "a78a5ebf530d3ea4a77c71ff791c7d8007786c217733f7d842ad5053d5b4a0e839eddc02fbf82c453a3d5668a7d3ff29",
            "b834745ae5a0d32385ff7674faeebce8c47054a9b26481b0a366a3ca4086f4594d15b1953e234add439ef4813631cdf0",
            "18d62ada08cf0336fa654bb557abc15bd21e38e40f65413571dd84d429a6e9cef10af55e0e1a95de3fa3b599ad16637d",
            "1ac7facd2829e02dcc7682d3328e239fde9f467caea4da1fe65de0fa55673c901b729169b12a21e73a65742203730ed5",
            "3f243e379b77ecf774e55c67d210b3188db4e3f53c5e47f03a84d713781dc02b21fb4d5832282b8876423e5e095efb70",
            "65a319af404d8db4e1b50aaf21764f21a210e61f8dc4b9c3f957fc44a7997d01d7a21e7bd23f9334853591cfe15548d3",
            "b3a6bc56e36de2da331869237a1856639c30440fc3db679073053cb587992e80b79339055c175cf0c968a7b2b739da29",
            "d3e823e1c1388d6d633cc7e0bee09f40bb92b5cdca6741d2846b51d9cda9199f39ded30fd4e4c7381c6ab6f3a9f6fe1c",
            "d1ffceb9b3dad36e9aff3091730c5847bbe83b6e301b13a8e436cb845c2f627f9dd3e23ebb4521ac841f11e973ca04ee",
            "e730d71bb84a0996f0b39c5f06957a984c7984485b5643f83012555bf6345a79c29685b760cb42fd3b4c5b56a4825f62",
            "2711de0096ca207e5e035c84850b2960e142dab3dc3a62f8a8e36710797f50193f8386bb35f2054dea13784bf509a680",
            "ec612d60d11052b30024bbcc86b4d791662b09d8cce7159a3da6788e8e5aa04a2dc2d6b4d9fcc5bb06a11520d7058c71",
            "039b43435d1a3484c50f5a0172432e98c100ec3243ead16f05dced0d2691544aacf05b29b12b2b4aac7250428da03c6a",
            "813a7f4a7c77ce94502d0ad6eeb280d3a29842be856888a99ca0c1e22ba673c6618e8d54999916cca25069f18a52f763",
            "44b528180592b9ff85b6858dd543781fe4e356d6992d8da40d313c9d198dc495c9a521c57e90234c681778db4125fc59",
            "b39556fd721372dc59435a0ea7d172386f48cbe1d4b435c613729c72b11a4066f7f840deb6b9c873b9ec94e32acfcfe2",
            "55b518d82f08b34fbe79771a6d8253781d44795f8dd71e2c5386d7e1696183586d820ff2a729e2ad137e48e616ffc0f4",
            "1f2bf9666037c5f8a07fd8901ab8cf10734e2bdc131cee900bdc21052e9464f36d23c7eabb076b0e6a821a7a38f70046",
            "439c7b0023365bb4f757909c425acc45ffba45a20ce31bb602c0fcdb4c230f19522f33db6f8897103d2cac33b88e21ae",
            "b0dc5c6afe6e2ad34caf905a3345cc8621f737b10758fbcefa8b895cda2b2361ac23d3164e20766378758bbf49f2fa27",
            "4cabd2c94150d7e7d102643dcde29fed7d3b867d5bad35131697fd0c9744acf5cb88d0e36a8bbd2cd3311cf770159a95",
            "4171b3694a7c528e9e7270e330e7466f8c999577ca9056de8489f7f368679ed37ead79593ec7b1f65d58fb244af1fc6a",
            "7b6d83666dafe700bc3d84ed1398b93629132773def8c74c68dfa4dd6a520a24d69b392d9c0ecd54e29f522871914704",
            "4fe7007d5a5dec33c8f5839cb5b984a8b88918b9cc61c8183986d7de2ccbbf8114031bd19b6273c14a44ed8e295a6893",
            "c9e74e3e14fa8771f6b8c2d931e40a48811948bba279601efb29c456acd7bd99739bb9b22865418763ddb010653d9457",
            "d70ead9aae0aeb3d9fd1a85439036a9432b3988ce0d8dfe1221b64208fdf942753a6141472d33d86af1ddfb519ee6b82",
            "63283f593cdf095e424834d11b99c9cbcfa2a4fa86128b98d264f049906a17809b6934a25bda346f7cd7ac127244e98b",
            "5197be185524e408f211e3e451b33fad23dd1b06ebcf5185f74791f74f70de1ff2e0fd510bef8fce20999df26380cb8f",
            "cc6b12f2631447ea80599bec553692e6d2062155465da19e7f5ba0b2f103f11572bcb719d615793622e56e2836b2458e",
            "2acc31c67384798e324566c0fa443638bf913360938f07318258ca9225c42e47000ba81bea85fa8a4d2e14bd345f9e62",
            "b26922a3f1a0f232c038ac2595f13b81e36e2ef6dc5ef3fd02f3ce4b1e55e1358cd732bcb275a262bc8ea2b46276dea3",
            "1c52933952329556c5961c4ff8ac324ba4e7d37f007b3e12eb6aee29fad00c9ea2df3f89c5e6d6a0b7ff4a6bf5bd7d32",
            "ec9cf68ce29fba3475b01d48ea0864d0dda5e5de2ce9e794117fc96d73b546ab1dc3e4da5e53ae59ad152fd9f629840b",
            "b267cc60768c9becb9f492f9eda4cbc770ec8c1956bb4e52aa99d3a249a158b58362553c5758a158990bc516cf28adff",
            "483937ca4dbd4d94bb943a4db4a8989c47b92c3d41f2b3d6bce3043035da7943eab531bc3a56ad4836da4dd56843dc87",
            "14f609496a5f576a2662cd44f881c694be005862d54978734678cf9c37983707f2433ee64957d7a958bcc55f13318b66",
            "bd2af865cc847d3217fa6a9f62e58c5aea0898753c8feb45e7b4a022a56f72852c7661cf0c4ff5dd3218af289a5b8da7",
            "6c431a663793623d3b6e30ffa0830308e1fd3d8ad711bc0e61eebb79cde366bf39b57c765e17c0e57378155e176647ab",
            "9af505e2e6ef386f38fd1d90f63a8e30e1ef685366efe7959e2a557a8af971860cbb1f67c4605ad55a53955993f0e30c",
            "618fbad7649de7c3aa80b3e5629e0f6fc53c9f51c1aa1f4f29bc74c250e7f9c70d39cf1f7448b301c33acf172b93500a",
            "7233c9a9cef31997854c7e2826c7b1cfe73e63dfeb3a633f65c8fdffc70de18df6a194607ceae79cf70a2f2156d3dd38",
            "6f767b1872bc079ff30931eda76d668a98df430fb9dbcac5747e003854bb50a2c4bc930e73e1af70b3994739c1ea679d",
            "e58f291ac09b1b99cd3ce6e022f7110918bd99733b69555aefe6eb9c26728c9c768c91b3198e91d386ddd7c9c6c8691c",
            "8da73c7d9d31575fa7f7d83f6db8fb3a9db579f93abd5bc2342ca3a9d3063fad5042ae45037ddcb43de41276bec44d58",
            "4599c378035bece6ab54a156670955da21e71519d685dc39a1c89a1e87ac69e71289d9f27798b6bc1c75f69d0aaad751",
            "583ce319f1bbda1f005d5d952926c338aee6e363f4de71e0f4343259759ac92a9e7d75ae186b8c7ef5662f88b68dbbb4",
            "6d5c19cbe7712c4f627f8e9bf4b1f2ab370c461fd5db85f68f6547ce9d2e1c306ee180f11964083bb500e4f77e5cb775",
            "c277e3b07185f719f40f13557277f0953ed5c7950dec3770770b8c097597ca84c1ecc23e435aecacfee7915b3369cd07",
            "326acd129f1c27d798eb96d6f659e7218bcc79c8dcdbe93667c9478865087180433ce1d18e3284185420120be58216ca",
            "a9731da805f2c29a29e0059060cd14b00bcec87d5e2d9333c83f52343b8a7a71e8ef3b45a7de368a662a2199e602599d",
            "9d64b66a2eaa217da037b2f33e33bd252341179a777dc6ecf3f37cbc1243cc0afd3863f7236d3412c4eb1887673ff502",
            "acb557b847277bfb7bdcb0dbdf6cf7cb8f71af13cc3f9fb54525884f31d5ced9443d4e602bc6b18d6eb6f8afb9df7a87",
            "e9bf567576f63e6ba6b3d95c9ec5283abd406e27f1583d5d22b24a7a832ea5efdb8db3ddbc504c2f82cd731799bb275b",
            "8573c0038edb9877c261be01e017b13b9751e6cdac29e36dd54c219e791e787af59cd863c5a04f12bb5563f75dcfcdff",
            "f9aa631d8f4124083184109b38c4760faa861062be02716a5af5092f24766f2640873c2041019f1097a0035cbacedf4f",
            "cc7492eabbcddc333b774a120e716a5a429c9c9e0f2527a6ad3ac98b9250cbc49c289acdc8152751425ceb5c39d9a9db",
            "321b5f808b1aa95eec03f99add9c8cf922cf1bf6dd0e52c779799c99469bc228c749bc4e4c4b4f74c890fee87cbc7e9f",
            "bbb59b0a8424f0f2c79c552831e2428911451f9c35e41ae688d22b23863d66ddad2f4690bbfa4c7d30a6c071ea036f5c",
            "58af8c285ce5acdf8dab462c5c35a2c85fe69b6df790bf996192691c489fc17c2b4e9cddc6f888a231e05e9c7fe09c65",
            "362b52a4c9103b2b33ab41f6caab4a2cd5a8c16699aa0d6ac5c28bfa5ef54829c96836843c39aa48c9ccd74ae7debeb0",
            "a57d8596f609d9f687d94e662d80ad9c47dc79d46b5c0b19fbce6aef1f4abfae254e393668cab1b19c8e77609b2a50de",
            "ce26a18748fb3bae656ebda5b669aa20f23bd2a6a98cdb345f3dcdec0cb01898f53757c9e48c42c84cbb032e569ffc04",
            "da56d9c5218c1ac20ad466d83531b38703f011e7017d9f97f810ff1633fdebcdb171d177de60f6c6a44b3b27b19de3f9",
            "9db49dcc26bc547178e0032943bc328dc4f35a71f493769eb0d45ca9cf3466f05a5b0e40ef87ca8a5c1a738588eac979",
            "33a78ba87e9c737c77269c62c2a2a5d7e13bb0196be084075718ba2b8c5cd7bcddc4ea3792b16bacbe9d667d2cb60476",
            "9a59a9bec82419446c039de613ffb898431b36e99829bded2d3b686b59bbd9a63846144ce70077f610277f0e9afcf93a",
            "7a50611d741467384ff1b6338e853926bc1f70e1c120398e6dc894168e791b111c78c2903b95fe59ec233657c4803384",
            "f710f27f0ac7c82bdcdbc76659438ac419761989c3869c65415eb1ea3dac48bf61c3ee57769379c328addb0d31cb1a55",
            "e0386759deb830e937ac486f67fd6ed27bd822bd5979d5cb15f398656586c96cb3ac0983cdb2dab9d473e2eca6e1860d",
            "39cb1a46add58eb57d94cceef5cabc06d92b320a4b68d21183cdb5ac1a54c565b539e6d89c9dcd2877f83e8267f20ce2",
            "f9456b35e1edf3f1d93adf3d8c67edfd259e742de0b1bdc5c5d89e28ee2d5e8a3b7fbdfadaab55d22686a773dab301b9",
            "739941d711fa5d087d2d71f6f5a0d9d7c212f5827437427a0a21a9f69f37ef30fbcf0b5e17c29bf2c44c85c4ac2e2f48",
            "dd08696f606e14fc51decc14bc07a1b374364703eb80cdd132970ac212622389c9f4a7b09611235b3fe5d4f4fc797f72",
            "5ae5fd420bef0b9fa2750c2c33b9ff71e4fa47cc7301581b4464e17ce5955ce560d6e65c69ec8c00c94cca012ed67dd4",
            "411fe81be931941f476487696d161b32e6711475d78431f37e5da7b5a97a1c01373f3fc9e5515e68c2352e247b6d3769",
            "5c1a6d773e03386ebff660873362fcf6b5e64c712a30feeff218ff99683815ce010f386e2b044e38d9ca61fb0c429c46",
            "70e9249b8493ef01c78b03ac70e6a709c51ad05945b69639646b88cbd66245b6068db9cad657790c8258411310cfa3b4",
            "9ec25311c31c6293470f43ae37f0fc0a01718ccbe9b3065b03825600de8033e00a35426d50cb185f031a36c42f600450",
            "628c4fcf8d57fe48c3f9d3d079a687170de5991e02bf4ba54be5badd7c46b57aaecd5c47b573fd6f2e3a399f847da0fb",
            "b2f3ff9c35ec679ab8ad01b215bf90217578d0b0c7b69c478d3ca6e865f67b2f4c3b8fe37927c25f3262ea0efbce94cc",
            "bc3dd6bd26edc4e98738a70d3b72b313d44e1061a1fd7d4e35b1e797676caa27754b9f923c773ba5fddaf47184f0eced",
            "cce3d6c79126d53eb739ff329eed79b3c56a2cd02462d148c4886373d83b0363c66c67bf309f1adc3fcfcf0db81411fd",
            "99f9de8a4918db93297d983d0867e43adf4d206f1de7752e0c085cdff951c0ea59520b3d7c12996192f99e84014c8b71",
            "3e829d5ea70df557417b09f82e06397739db367304329ec564b0d3fa4434f0206228eb0dbad61b9fa5de996385bdeb55",
            "7595372d99aeca9696d9c3ae25766fce6945bf03f6d6c8fe09ae0b45dbe652a7bacf809fa7e5d0bda7785fb42bd8b439",
            "b0619a39dfc40672b75d92917069a23d436408e5c35ca23ddd76279cc637a70cfb1e857194488328d31bf92cf494eb2c",
            "74c4f8438e7464bf8007b6448288b2db1db50f9ab5bf86f8c32fe001b59945d5479f9a0c79d69ffd17f020cdfd17f000",
            "22bcc5136f39b736cd30c29b4c6c0f5d82b84c80939c575ca63ad2bdfb8b8d7456c239ab1d33319a428ce61781e9c50f",
            "738f4804af8e5512b168bf83fef169fbbc27f37aad1189631ce2a8027104215623c4208758e6d133d5867d6e76d612db",
            "192422f2107016b327a80d6a39ab4424326e0c98dc38973efe7da3e1d4bc69385f29d58ee375884b29d10f5629036917",
            "90224cc2ec0da4c8cbd82ad6b6f339fd487d154b7186c2ef56d4683e69f281ae9a7bd4a8617498f91699776a44e8f8b4",
            "7d7e388ad711bc9e55a246ab0d1935cae68d4d78fd16d7a62b5cb5ea9891edfe5da655cf7b68436f18ea9ec86ab31f2f",
            "fa4e1b72d7673ac35d9f5ee5ddf8d4ad3cd32b106fd24fe4ce5815d261cf9db18c8f241c764aadcf03cebd59c2d13ce0",
            "0c6404e7cf289cb7383ec4039293c63a964e62acd1724dfccafc5b398c0f199c1ec3696740da6bf83b5147c8c773e12c",
            "da0f92001dcff7d1f13c8e5cd263309b012043be21f479ad40ea032ef753d862ee1127224ec2cb9e80763de0d1aa4c76",
            "32ab70921e70dc24ae138eaa6bbc399ee91a56c2694f4cd87a3e6f5e1c943d46ce7d76a71ddaef14d75527145d751c75",
            "55e63df0464eb3a2ab0639c4010562dfb43871fc22fe12dd727e62b95b5ff420c411c517db336dddb928ce2f069c0157",
            "a8116a835a7ab8f603bc7ab8f613bc39171a325f6c2634743b5f79be520ce6d1620ed8cf478a53b3652691ed9aed946f",
            "45d69a8dcde5d81cd047dc349b7abaf309bc1e67273de395fd86541f5e7ba6e5def51ebc36a3741c415fe8105e7fc111",
            "e8314fa9f03a3961d28506930a0d2615ad3289fd31c2fbe3bc477f4c5afa635219cb43a636704272b34e2695b16c87a3",
            "79c071f38dbbc36156d15d1e70dcc6eca422edec700e78c0711ba94e38eaf8f2ee1136beec94c8bb4414cd9f1745bd38",
            "c48b8e72d6c6cabf906579ebcca1cfb0ef09613273b6f9708f430a83ada89e2f0e900122a4837c9c54e42360d0ec2223",
            "404a8ce0d8ab7691120eefb32925dc4fc8ed33ec51b293cad867562d93001fe618e49bbf02cb467df6dc957d6f0c2934",
            "3cedd11b9969b0997ae3f40cbd3164f6c6a59156d2769ddfd892b6eb5cc796d35f9f99d4665cc96693822b8137b2e34a",
            "a6cbed5c797a46ae1c52b8f265ae83aade86139340a3be5d33aea319da1a99ee873b6ea8670d656ba13f8fe5dd22357b",
            "94d1d66eb85ba00f7a5aa0c70d7bb45a1b1f73e386fc8560e183b7e2e24539ef138a84cc6794633179e386b411a1ae99",
            "ce2652cf96995f2b9b2c92a5d071aa5153c6adcd944f9c3b32b2b13fcf9ba33848cee258cade021511711386db88ee31",
            "dc4f4962639cd99b4c96bdcc4bc873dd328d9b67349cc0fef396613d854b38a59eb0c4c03bcf402e461ce6bff322c83d",
            "40f7bb9e2431aaf4d9980bd6638a5c1a35a42768cc432e79c350cfdf67b41a437e3bc5f96dd883dfc62cfcc6caabb1e8",
            "5eebbe9961c2ac5c58891a24f29721549c32f9650836de193567aa99f53ee492bf93a0c6a1b765547b1b8efb6ac3b997",
            "6b82cb35800c1021fd57fc1ccf361cefd5882dac20bf9973cc3c29458d4b87d1c02497884b9fc098e57330de72983efc",
            "b892cb1a97cee03be3d24734394632a7b1f59733ce69bff1bb4793071dd1e4de7dc06c412f782ccefd0e259adc0a49fe",
            "8e898873f782c4a2c9cffae7d66e114dee059d45bdffc6134fb5c52cea9df1291bfb63c4eec160da353bfead329c31e5",
            "c179f171be858f0b2d7c0c90f33df87999c9b92ca67c9312530eb0ac31e573973b4e3f4b665206601cd6724c85142453",
            "288fb38b3c2b36e46a3bd379729d5efe9616d37255868cff0e1a32f26ccc102729f9f9af7af298354aac88a93a34a16d",
            "17135f92e4e58a3c3e52149079ae1679480bd1cc028104c9cbd3b4025146232fe5cc06d74f8afc328f480d90a27c99ca",
            "2b7e306442b8835624f2e690930532ef6a91c7df03c898b8e492a2853257c98c10f3c8c945326f99c8a31d214bf34dc4",
            "f3495150665ac601969a992350bdc85c40ce14c9cc6aad8180c853488a96ca3c2b66c47001397985ccabf686096e2129",
            "2a9659ac555ab13469b4889cb94a1671efafc5a468c56ccd36015e46cebc4b66563305f24c4483a4282433899ef12064",
            "1139b3d2bde1266d9690a2d51e5954a04bf3cd12979333d7b8934a366529292a937944ef78b4fb0a72668d3b0a66ab97",
            "11df95328b48bd9214ad95a99cbf2c3c20f9a5989c5c27f35af8be4072ca55a46895b39e12521476a43ea8f678c0ecf1",
            "e5e4a57299d7d2890566352b4851a5cc24d8c28a5481499f7791a737bad3c704783529aa925944ea4a5274a34c5dccab",
            "7105f06e726ccb2c7584c8c5eb3db2706bdc63f89592935bdd0b9abdb38af8de2fb358c0c823eb7f673fc9b21db5897e",
            "12ef2a3b1c89f5eb646da862cd9ad0b521527a7d29fdaa6b21e964596722d6df4be2a9358b0a43a11009f10f09352692",
            "a164e25e12da1a023071c2c1ac25a1f235a1d644e83602c9904a2099a62e228566710b1c5a891b9c0a84d3490085786a",
            "51a1524abda9492623e4c81ebd2f914c87c8567c3eb287b06768cefdd09c964e12894592908d24d3bcc965b4056b69f5",
            "24b6c6d140f8ecd207481ad0a32f2dafe80b5a172dad946a887791d0f66e9ed2104be94aa91a8258b1da695b4bf7de4f",
            "185a9118616825d384a125086502c49bc6fe78673a9a20f145858b0a9b9351723892d6219540723cd4d243227d7a4dbc",
            "6b2f12bcacf608dc84106288d4a4427b23c96884c4d36b42f4013a371a3fc831ac8ff686929c623529126a612f191a49",
            "bd8f2410e9d623a44fdf15e9d5cb0e1344d4ce2ddb53bb3bee267a27a32fd14daa6eef0ec13b72377f47a0706b8f1e27",
            "8b0a4daa26191168457847a0f7a0f3c8b5963ee51cb57e2f5651b1c69de0a2395b693771562021c60a96bee09cc0a084",
            "1ae870607d08705a7a2845099214f248c4b04b428d8475499c128a949bd615bb0992cba8fdf71373e28e92ac3d9ffc7f",
            "437cd655ae0f53a66eeeef88453b25b010d93d106d4d46e229d2ad27774692f7e874a8da3bcd250f7055aa47277a7a5d",
            "2f4b595711ae2aa55520b2143c41640b296e0b286e2f5d2d710bda70931f785fa4c9b61dd2eced84f30caa48b9a2b2a9",
            "19d44e3ca1b152d5e44ca92c41cdd3f68067898bb63a2fc3f2e58a3ebf7cc6f2d2fa1438fb69fb9fbe56965f9a71fb9d",
            "b0001a6dbf620c5c316bfb1740fbaf932596cd887f8dadfd0bb07cd17b65f92bb32edf46562a9605b522db176644bf1a",
            "e5bec8c10b7ea447157969bd847d5546dc50fe3e59a264c6d6d847da021c69b72af6cbf2197bf390d2023b672dc016b4",
            "919f5448682b66a4cd4905dad32a6d5c78ce8f9c52a39845efca803655e4e94db2c4d519952852cca2776754e2986225",
            "8532eab1959b658995338f3f17995145562ae655694623e6e44db2c4aaac47bc3c318488388607c93c3ebe1b9e19f4e1",
            "bcd54fc24fdcac9dc8254d9ad65440464a97e636f9298e5d3f0fe407a86e58b1b7b6a679309900ed97434ae860abd60a",
            "7d8709d9fa156d599e96f4d129e792026dc9604bc1bfc47c64a176f3921bb5423a7bd8924a77051331fd26edaa144913",
            "d245a8a942f466b288f4906bee22feebc8fe6df785c3e18af070e586f03a921f08d49142bfb6d8b7241caedcf4c02a52",
            "b13ebc2abcaaee7ab27f5f34de95b83745f6b71c49a5f5de0d89f8ce48bc3f12fb650995e08384187e1fa1d6586f7f2c",
            "52f189ff1d24854f96f9a0e3ea135443e9add1ce7bbeffa74152c0536febea6e49470eeaa9bbd6d0ac8148775a4fee4b",
            "46d3faed6de6736334994abf7898aa1bf6dca27726e25d9ffd3e4de050eaa8c5475f00f025e33219f52d03bfe1cb50bc",
            "3fa50bd476bd07e077e8dd89a4be478f74b5ad0a92dc400a33ffe66741b228806a5b68edeb7e2271de9348a49b93890e",
            "fd8baf07c9e2c0bd89e43d5461eae986783a79e49ae764935b23070fea5d80d555a54192f3a17c244e2cd2a1c7ae8a48",
            "346ba906eeeca17ab8534f522a2c168dd263b196ce449f8e6ff65c210b30b80dc96422f978885a083cb926958a1e8cf7",
            "eaf134165b3841510e44fad389de483ada195f0d0d4e450eeb5ddbe37dfd69fd57b2e41e9dea7f3ddec989f5ec7768fb",
            "7957d645629d945e69bdb13f16bbff9f6599bd099a1a8de99cf6cbb7c9ee17affeee0b41b2d0aca257efed607d74e5fd",
            "9246bb74ca4548a3fd3fa6035ab49dd74a0d8d96b4de97ba690f204f1b95c4b4d4d1bd949ebcb6fa68327de4c77b25e1",
            "30617b2ac1ca2f9d868e8d4293cbd3f79af715fdf7cab6502eac491eec07d2dd42d9fbeb4fc857f547e291de68a76852",
            "4f4190e4f18ed4e311cad15d3fa0b917d98942fb2ef578134532d00ddcfbe80791bb9071538b8364a9683db582bb10df",
            "7abd4f8f77d13e38f2999d940eaca35a12fdc94e1d19f0aa1f5108a2c57a4c4feb6707249c56c0c40667e9e392e7d2f0",
            "fed8454a440e62473495a6463695242fdc4619cb919ada3c2c330fc45a929d7b22f1837a4997a4fb40ec763df58be5c7",
            "68f50fd1bf691719fac7cfefff470e955afd60348e4c48c5485a6fbf5b8ee7c6683c9aea91efbe5625dfed8b76a57b70",
            "66c6a4cf27964ae90099fb53201d1eff2a2d619349281da0a6ad41b2cc4d724061941f131417cd307cbe769a25a44174",
            "7d29feaedb6a0db6febf4703edcbdaf35ecd67de2fd660b76c5083add8efa3a57248261429406f9c8f98e696aaeddd52",
            "829fb9be8cb4d593dafa569f2fe7988fd5e2f33563c5ce4fd823fdb7f1a1c64494245d5a9ae9279f683e0bc00c3ee039",
            "5baecd9a2da30f80c9b6fe4bf9996ffdcd640fd94d6e260da48eb4927b496f96e597122debfa8f10794251a062bf5faf",
            "589f8e759055bb7734846afad383687124e224fc30f13ddcb0bbbb3bdaf9a186e3bb49f7206596ce871a4a3e3758f4e8",
            "a0ef5843457d23a9aa0f57d6addb584b1a6bd69587cb6bd79da8ad6fd838d850533358575239985371bc61f0d4c16484",
            "f4861aa96a4dedafae4bf4f626e281dc9d24da994ca412dd69024e96a4deb59fec6e6cdc5ed750be89ecdfd9b27b7dfd",
            "8e1db9ab3e127c7867a87cd3fa707e6877c7b197423ba21dc3c948f208b985f82e7be87bc7fde4f287fcc7dfbaeaa121",
            "d2da134deda3f6dc60472271cf36523c78f943ad835f8e9214d93772fbe924b987748cf81223beeb4abebba0e2c28215",
            "234b08b9d054eb2bffe607ab3f76759d569df385b52faffcf085d2dafcf7d6fab4ab6aef6f01cbb0fcba6d2f2f6fa95e",
            "a557b7a5494f392969caabd895f7b3e59f2a7cf6dbcb0b46ded542957d0da176e136b2e4989f1c8f90c3832b9a4972f0",
            "a6d73baab512f24841c5a305bb43c117be46c8ce1befdab7720d4c87eb6e0cbeacf5266fabe63de1f5cd07ca51f17d98",
            "7d9bb39625fcdb5c7622eee2c59a2787dfe3abdbb505cad20c51cf55144b720b60956e2b2927ef2361330dbe6be8244e",
            "7daea57938aa0be0fbfda64b5cc3b5987c91b9100b9793ebe9b5d202a29ea6bae5aab255bcd9f6bcc5527185d8a3bc00",
            "e4f91ea2933e3a6d8a904e7a576b96843c113a13ecb4941db03c45c4fa43de4b0a91055d2b8078ab657ab59378be42fe",
            "dc4807509ccadc5a1be61b6dcf9b2cf56f35171e34df1da4207f266cbe42b139798d4c1f54b05921a0901622bae1af23",
            "7f81ca1356bae1efe7015900e6b51a09f9e8fc926c23f7d18e877f15f4af92caab46b28ea7842d77758e34f55f256ddb",
            "0374de5e49d65348ab689fada225e8c4944ac2249dd81ea4572a2b28d300ade834970edefd344782a6f5d26bdcf66627",
            "88242c99a27fdd747e1ca2adeba13092344717cdb19752bbc6fcde4ceb2de74f0df8544f76d07fab680a088a7e5a3246",
            "61d4d2ef14c286fa12145a8896ee431e82d408ad298af8a488fa2971a15615ad71c3bca8554ef1044a8533a2961b4da8",
            "f4a6ffa2c8fdfbf15d82a6efb7a497d3daf653aab4d13a1a284dd6e37d03add30d62031d25500aa80525d753eca09e0e",
            "72374dedc43c3b68b90ec431895a477e9e276e5cb5c1469d6ce9548158a874dac7475c82ea56a04d0bc503bc283aa5da",
            "069a7f3f3e71970a964fd33baa0c694b76d03642aff7d334e03fd9e7eaa7ccecf10aca4b5421d22b60544771dd48f9a8",
            "914258472914a67fb5f4ae96be6fa06f584b6ae83fd6824a9e5e61b6a42ae31e9fcff8d84df33792edd8e78c0376d21c",
            "bb9591311b3fcda5ffc133a65109930bc1236429b5d497d1bf627a5f42bf57d03f66e14b8b86b9c5b2367009c4bfc872",
            "4e0b9e90e36ff94931b5be7a5002ecc3d14e4d051c23f7500b2b4ce5c426daa26e9a16a6e37023e51c98757897318c0f",
            "5fe0c0519007008d8548789dfe4b53822d45c1e004fb5ecd9acf303ef52c83542c412da159a89541b330aee8a0774b11",
            "8813e07acd2db761dcb58fc978c2a723302d2160a27042bd33e62fbf43b864f2297e87e17bf4b5dbbffaf933876e79f6",
            "fae0b5a7761e380669308406cf76feafa16b73770efdf481bfcbfbccb75788f4df047fd9b2fda98fd63e7e2ee7de8f97",
            "1efc8648876f884319444e1ef63d49727d930d70bf80f83e544c0c3fbae39f26017cf091b43e90ce59d830d0a9c706df",
            "1b24fe857b6b6bdefc2978e4f645e3e59b5e3aca6f3754e43dc86f37550ed05b7f6067a4f3abcf044920408b6cfaca2a",
            "7eb7f9297a971f68660ef8f2c22b832477612add9588e99f680f9242a880bf7ce1457887f31efd8afe2059b8504c2c60",
            "5e71ffcf82a46090b9320e34c0c27124ad17bf0145d86c61a22448f2023b135dfd31bdbca388e65e083304f439d75e01",
            "c8eed1537afaf29bc0a9d89fd26febea7e615c7af16afafa62d14e9c813d7e9d4c563c233ddf08926583c28da7bca811",
            "dee21f5c0892cb6d05f55af496032a5fbe20fd80e818bcb010d0c245f09fdf42b11fb42c041cba1d5cd175897877f460",
            "7f528f14529272cf646fa24bffd261da10fe9c8803fc1b7e01a4a614d077c76347b407c0fbbb1bdfc0f38b2d5059ebbd",
            "94ba9fdc013488f48213fab3512824a87c4f313cdd168fa65b2927bc4501fa02d1f620ad99d4a43af7bdba0d60b6e8e9",
            "fe3e74d6fef20712076c46ea9b2b0034a546323a702a0ee5ea7a92fbbef0a77077b8a32e79f96a76b7a3bbf51a490d04",
            "56f21fa4b316bdb7155fa0c002224ba2efc8b91765c7d4eba974348ef41f7b8e320b4f6e8ea4d27a4b9fde198dc45eb8",
            "45d680e99fcd912efa8118266145a9bf29557abc3f0d75011987f2a10760a10311bce521cae7bcb5c9c4bd7ff22ab07d",
            "6722d6f420b499453aacd660c984859420ec272fd28e25db53b8f070db8fe1616f24d97aa44ffff33e58bd3fdcc1023e",
            "da75da8800effaba484a4f953c01a0f481bea49e4a81ff7edddf02f565c2c04fe4ea8118115f6f85d5953a3aaf8f74a6",
            "f52402fafb17a557da7c934a1ea4d4618b40e1c392ce8d8978fab6a07cae4dc4ba1e7e8b3650b8f693d1c37af28e5f02",
            "65f689459c97a6690de440ed81543a796067b46bba0dc21ce8f3e148f20030c06222513d50af7747fa63e9bb5f0e92cb",
            "023a439c36878e9858eacb5f03fc3be96d47a4f39e1db0ea937aeecf20638c76171d355d895ea4ebfeff0cf482101e11",
            "c1f37c8374edd71e49eb4d89545a2c50001da6ee029ccde590c92320a252b144baa210404594b58ca91fc26a42448a84",
            "9a5834927ae2155c8fc0e5856dcfc2d088b405e1abf30efceaba330822a83b128dd1f1baad017b0bbabd8ef2524a7b05",
            "ba1ba374a8b0fbf77f0ecba2c0148f6037c4f448f267df91fcda48b1a8fda78e2029e20926ad792c52fc9fae9512a385",
            "96eeeca93d77585299d22416e9d46b1f5a4a9126bbfa6119e98beb80a36a44044e32fdeeb85c98e964a15e7d5ba003ba",
            "a29414c063c8b63710c6c358f2339f54d6b9006d2aa1d2946ea9fb7e2825b429fdfee593b24db06e93faee8bd0f09aae",
            "ae37932ab2dd94235b13c79f5404326de5611d65fb35bbe46211f42ce2015dfae06394408e2ee72f7705e4e2dfee64f4",
            "60348e8b93931f857e800108f84cfd5f09ba458fe9485a709d7ff6cb341b274d0a5fe85db77f51ae9799ab7fb8d0f8df",
            "283b5c1ea003f37034d19f5274c3df9e942a41511afeb3ddaf7ce9f3db97079f1b204f90aa1b9f13919c7e5f37599a7f",
            "4ceec8b2b864a423438492fab55ea2e5e7cb5064f38cb805e0af68a13627cc023aa9696b7a3f7cdc93617b2bce471751",
            "d776df8dad4a1547d56d72ff8c6e9bb5e0b69171f864c4c5e7c59c319536f78bea3a62ae296b4a859202df7b67c1e4eb",
            "149395eb657a3020ef43a266f49d696af33a6c6896db9e2b6ccf1b2c1dd88af39423f46a75de55d89e37d89ead9eba3a",
            "5bad75b6f247c9fb6defadf0ea6df9eb95f7404c98915658521a1d29db30856f22580001c9bb292745e94c1958b486de",
            "c3fcf388c51bb895ced64bb9dfaf94ceb7d50eceb37828776207efa278d9b8eea4c22fa2f2a5b6caebd03d12b334712b",
            "798fc9fe62e7c32e9cd2eb268ab040f21eb292aca6e8af56480cef3b68de2e4b4a27a6cdc461bb2987c96026429a150e",
            "e3aec9052b10377024f453ccc14192a44c0213fc7b4819596369c32e8a9328773596abc7721de87690a5c091bb4612ce",
            "9fa684d3dc448439d41be834151c70fde86cd1ed6535b5ac8a1174932a6eee47171dd03e45bb4f1d36474dea8b5d227b",
            "798de0fe113915dc9061ee47f1d381fd6987d898212d2a282d447b85d3be0eebeea4f963f616239b591aac8a662b1979",
            "e72f583f2302b0d297a623613bede03ef46aad11f1f633a2be13bb71bdd9158fd2da732d98a9e468b13cd55b9e769ba4",
            "82b0cf56dae483f41ff8a1646da594e94bc95aac532dbbc7f2546379da67810bed3842db08830aa0ee45ef1ef3de41cb",
            "ca2c92702daf718d23b51653d5c1ed4643f15e0eda7a75d0faff2b3fce62843d37aab534a17469b2c966c9902a8b37d9",
            "646d934d760214c15e7c5b06fadfc0bfa97ae552e6d9aee20cd71a4a7f53a4f80fd55bf15525ab6cbfc43125f62a2f68",
            "22de3cd4426ea37c54e601e14ee5e90394e6d6be103278a36b6ab9d90325587f0b850e7ddd4fe1aa024c15ae7b689b5b",
            "282e3572f0ca2840b105c912f4c46bd6c44a8e0c7ce2703521c2ecc14fa2a4b075acf14de4b7fcf9c42dd7fd851f231e",
            "c4a98d9a7927cf8ad5945363899946cc349f9926ee726cff0849de36c84931c84b56dbbe39258f721a9b2b9e5edf221f",
            "efffa3b9b37c6bb3bcf782c33bc9e41011b22e35ad8f983b04e9e769c5c93cf1bd8e8f9c7df260f5c3bf1cfb8f7db70e",
            "3cea2372a58de5f09b6ada4f1e7e4457a006c86b72cba7998bf554b95287b554ee0ca5224a29dff188524a23875c0c87",
            "133e671a956d853235a2dc9f51ee2f2ae6d16b4aba6f81bcaf51ee0f29f76794fbd7947bb9af939655ee0f29f76794fb",
            "d7947b9f3caa8c942bf711e5fea472ff92727f51b95fb958dedfaadc1f53ee9f56ee2f2af72b2f53ca2af7c794fba795",
            "fb2285a96e55ee8f29f72f29f73ea507cb2dbd49cb283d49f96a86e89b9702deef2ee678bf8be479bf1332507488ea4f",
            "e5a868421b3127aa2821b6d0991e54bb20966e5b319efe31605560dcd6d26a7989baebf7d3120793111191c20252f8c3",
            "4e11942262526805f0b7797d39fd6ad80c41290265bfc4817743e14f5c460caca2433b8a14aa3faddc0f13fb4734e572",
            "77dbc17fc123fddfbdfed6a79f1afa61fdc7befdd6079e7ae5997a3b6d99cf8ce3aa89ed010e82481a3027f8fe642291",
            "decfeecb37ed6f68ab6bd8b1bea1ad4180175c81ee7a418988220104257c84f56791c2636794fb314f4a8403ae5689f6",
            "6ed9061eda2e22db2bf633ef3e0637d909c15e09546f55d85aa02afae798f2ee35e57ed213d51d6353ebfedf9dcb83df",
            "f89fe40952ffa33d82dc62976896fc2742a2444494bd29ac5b44538ec983f995a6306a3dadbcab51466bd0b128aadad6",
            "d292bdca83eb7ebaf14ffed3dfd47eb2f1d967be7943efcdefff6791bea4de35bf26d49390001cad454bd1921cf020eb",
            "798fda57bf9e937347cd3f341eff87cecdb7ee28ec12e9218ffc65c6969265d58f577ffac08503535ffade16915ee481",
            "ad9040a6dfd32ecc54d7af3006c5fe7e37f7af282f349be20615af1629c54d37a2805de480cd33d8012bbe463b608b87",
            "d16ea15a3d8aa25a2170e54e103b85ccdd207644940d1ef657ca260f078eea0e0e7b5de62e0ebb3e90bb36ec55291b35",
            "ec2db66ed6b01754366c8857424bdb366dd85fdb7666d85fdb3661d819ebdf62fe6147c9b697c2fedab69fc2de65960d",
            "14f6b1e0b65b421534d7137b09b77d11768c6c61eb766ad843d7edefede1eb763acb10765152ccaddd43d5b96c5b5482",
            "e75b84f0418ad77cc2a2ea18242107855931930bef7d941b0f2d91e9e38accf6764fdf811155a63cb64944de22ff1bf9",
            "aeba354723672e97b5bca6dc0bacf876c1c5fec073e013cacd5733a9ce8906dabbaa97ac814ec121f0790f594f27e9b7",
            "93669a52c6d15d6377b5ba00f5ff0f1bce1526269f46ef949b7bcc95b01b286123f2643512520c3395b05abeda9e16da",
            "9e99a0de0bab0b2b647abb02b554406507d7cc18ce7996c279fa5d32fdb40776be3c8e5daeda65e34a6e8f5a304d9907",
            "fac5202c50e701879419a59c70f85cf3bea69c7a1156f2725cf2850d755811ebf9a28cb9282a720bc0b03c2a320bf9a8",
            "2e94724cf20576f6255351568c6bb9782dde88268a35363b0610ea20f20acda0ec4814af8498b5ee2bb4371e7711da1b",
            "6f6ef6b3232b77ce0938c27ae39bd8ece071d79a48346d5edcbfc69fe6b5e955f9a83339cf703831a9b304b629733a97",
            "f8b44c3fdbebb796ded7b0b96a73e59686ca750d750d15eb2ac3f535eb6a1baa2ad685c3151b363654d46f6eacac78a0",
            "7451617da2131966ab8af5fbae6d0af38f9283216c79c7227ab69a582f2a048db6b554ce4069154d7aac0fd7a107d280",
            "5a9826edd593c067d4caefa39d48bb6b43c5d6d20d5b3654545450b83443ddce6d5b4bcb6bcbb7346e2caf2bafac08ab",
            "7f8076732d2db06943657dd5862d0d55958de18a8d61f647df6eabdb5abab17ce386dac6ca705d18fe6a1b4be178903b",
            "c1380c350ca441732643dbe3dd890f2c2a345b54bef5be0d551bacfb4fb2ff0489e6035504ead83032db7f721ffdfbef",
            "6e819873fca8016b970eea1f46fd15f59b2ac3ebea1ab784d79597d735aeab6aa8dc40c74d4db86e4b797938bcb1e681",
            "1ba87972833a0e2867e15ca88f9aca94a1d4b1b435bc3664feaf5b54c8469123558ca0ad159bd686e0af7c5379e5dad0",
            "e6f086b5a14585b323fdc7cf25fbfc2b504b01021400140000000800d4926d5d512ddd41278c000000da01000b000000",
            "000000000000000080010000000070726f6a6563742e62696e504b0506000000000100010039000000508c00000000",
        );
        let bytes = PACKED.as_bytes().chunks_exact(2).map(|pair| u8::from_str_radix(std::str::from_utf8(pair).unwrap(), 16).unwrap()).collect::<Vec<_>>();
        let mut archive = zip::ZipArchive::new(std::io::Cursor::new(bytes)).unwrap();
        let mut original = Vec::new();archive.by_name("project.bin").unwrap().read_to_end(&mut original).unwrap();
        let original_guard = original.clone();
        let before = ovba::open_project(original.clone()).unwrap();
        let before_names = identities(&before, Some(932)).unwrap();
        let before_index = before_names.iter().position(|name| name.name == "Sheet1").unwrap();
        let before_source = module_source(&before, &before.modules[before_index], &before_names[before_index], Some(932)).unwrap();
        assert!(before_source.contains("OxiTransferMarker"));
        assert!(before_names.len() > 1);
        // Ordinary source imports retain every omitted document and its body.
        let encoding = codepage::to_encoding(before.information.code_page).unwrap();
        let properties = before.read_stream("/PROJECT").unwrap();
        let properties = encoding.decode_without_bom_handling_and_without_replacement(&properties).unwrap();
        let documents = crate::vba_module_directory::document_module_names(&before, &before_names, &properties);
        let imported = ovba::open_project(update_project(&original, BASE, Some(932)).unwrap()).unwrap();
        let imported_names = identities(&imported, Some(932)).unwrap();
        let properties = imported.read_stream("/PROJECT").unwrap();
        let properties = encoding.decode_without_bom_handling_and_without_replacement(&properties).unwrap();
        assert_eq!(crate::vba_module_directory::document_module_names(&imported, &imported_names, &properties), documents);
        for name in &documents {
            let old = before_names.iter().position(|identity| identity.name.eq_ignore_ascii_case(name)).unwrap();
            let new = imported_names.iter().position(|identity| identity.name.eq_ignore_ascii_case(name)).unwrap();
            assert_eq!(module_source(&before, &before.modules[old], &before_names[old], Some(932)).unwrap(),
                module_source(&imported, &imported.modules[new], &imported_names[new], Some(932)).unwrap());
        }

        let copied = copy_document_project(&original, "Sheet1", Some(932)).unwrap();
        let after = ovba::open_project(copied.clone()).unwrap();
        let after_names = identities(&after, Some(932)).unwrap();
        assert_eq!(after_names.len(), 1);
        assert_eq!(after_names[0].name, "Sheet1");
        assert_eq!(module_source(&after, &after.modules[0], &after_names[0], Some(932)).unwrap(), before_source);
        let properties = after.read_stream("/PROJECT").unwrap();
        let encoding = codepage::to_encoding(after.information.code_page).unwrap();
        let properties = encoding.decode_without_bom_handling_and_without_replacement(&properties).unwrap();
        let declarations = properties.lines().filter(|line| ["Document=", "Module=", "Class=", "BaseClass="].iter().any(|prefix| line.starts_with(prefix))).collect::<Vec<_>>();
        assert_eq!(declarations.len(), 1);
        assert!(declarations[0].starts_with("Document=Sheet1/"));
        assert_eq!(crate::vba_module_directory::document_module_names(&after, &after_names, &properties), std::collections::BTreeSet::from(["sheet1".to_owned()]));
        assert!(copy_document_project(&original, "MissingSheet", Some(932)).is_err());
        assert_eq!(original, original_guard);
        assert_eq!(copy_document_project(&copied, "Sheet1", Some(932)).unwrap(), copied);
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
