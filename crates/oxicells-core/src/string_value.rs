// This Source Code Form is subject to the terms of the Mozilla Public
// License, v. 2.0. If a copy of the MPL was not distributed with this
// file, You can obtain one at https://mozilla.org/MPL/2.0/.

//! Lossless OOXML escaped strings, including unpaired UTF-16 code units.

pub(crate) fn decode(text: &str) -> Vec<u16> {
    let mut units = Vec::with_capacity(text.len());
    let mut remaining = text;
    while !remaining.is_empty() {
        let bytes = remaining.as_bytes();
        if bytes.len() >= 7
            && bytes[0] == b'_'
            && bytes[1] == b'x'
            && bytes[6] == b'_'
            && bytes[2..6].iter().all(u8::is_ascii_hexdigit)
        {
            units.push(u16::from_str_radix(&remaining[2..6], 16).unwrap());
            remaining = &remaining[7..];
        } else {
            let ch = remaining.chars().next().unwrap();
            let mut buffer = [0; 2];
            units.extend_from_slice(ch.encode_utf16(&mut buffer));
            remaining = &remaining[ch.len_utf8()..];
        }
    }
    units
}

pub(crate) fn encode(units: &[u16]) -> String {
    use std::fmt::Write;
    let mut out = String::new();
    let mut at = 0;
    while at < units.len() {
        let unit = units[at];
        let literal_escape = unit == 95
            && units.get(at + 1) == Some(&120)
            && units.get(at + 6) == Some(&95)
            && units.get(at + 2..at + 6).is_some_and(|hex| {
                hex.iter()
                    .all(|u| matches!(u, 48..=57 | 65..=70 | 97..=102))
            });
        if literal_escape {
            out.push_str("_x005F_");
        } else if (0xD800..=0xDBFF).contains(&unit)
            && units
                .get(at + 1)
                .is_some_and(|u| (0xDC00..=0xDFFF).contains(u))
        {
            let low = units[at + 1];
            out.push(
                char::from_u32(0x10000 + ((unit as u32 - 0xD800) << 10) + low as u32 - 0xDC00)
                    .unwrap(),
            );
            at += 1;
        } else if matches!(unit, 0..=8 | 11..=13 | 14..=31 | 0xD800..=0xDFFF | 0xFFFE..=0xFFFF) {
            write!(&mut out, "_x{unit:04X}_").unwrap();
        } else {
            out.push(char::from_u32(unit as u32).unwrap());
        }
        at += 1;
    }
    out
}
