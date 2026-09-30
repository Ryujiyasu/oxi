// This Source Code Form is subject to the terms of the Mozilla Public
// License, v. 2.0. If a copy of the MPL was not distributed with this
// file, You can obtain one at https://mozilla.org/MPL/2.0/.

//! The case table `UCase`, `LCase` and `StrConv` work by. It is not
//! Unicode's: every letter maps to exactly one letter and back, so `ß`,
//! the ligatures, the final sigma, the dotless i, the long s, the titlecase
//! digraphs and the letters Unicode cased late (Georgian Mtavruli, Cherokee
//! small letters, most of Latin Extended-D) stay as they are. Measured in
//! Excel's VBA over every UTF-16 unit: 973 pairs, the same both ways.

/// Runs of (lower, upper, count, stride): `lower + k * stride` and
/// `upper + k * stride` are a pair for every `k < count`.
const RUNS: &[(u16, u16, u16, u16)] = &[
    (0x0061, 0x0041, 26, 1), (0x00E0, 0x00C0, 23, 1), (0x00F8, 0x00D8, 7, 1),
    (0x00FF, 0x0178, 1, 1), (0x0101, 0x0100, 24, 2), (0x0133, 0x0132, 3, 2),
    (0x013A, 0x0139, 8, 2), (0x014B, 0x014A, 23, 2), (0x017A, 0x0179, 3, 2),
    (0x0180, 0x0243, 1, 1), (0x0183, 0x0182, 2, 2), (0x0188, 0x0187, 1, 1),
    (0x018C, 0x018B, 1, 1), (0x0192, 0x0191, 1, 1), (0x0195, 0x01F6, 1, 1),
    (0x0199, 0x0198, 1, 1), (0x019A, 0x023D, 1, 1), (0x019E, 0x0220, 1, 1),
    (0x01A1, 0x01A0, 3, 2), (0x01A8, 0x01A7, 1, 1), (0x01AD, 0x01AC, 1, 1),
    (0x01B0, 0x01AF, 1, 1), (0x01B4, 0x01B3, 2, 2), (0x01B9, 0x01B8, 1, 1),
    (0x01BD, 0x01BC, 1, 1), (0x01BF, 0x01F7, 1, 1), (0x01C6, 0x01C4, 1, 1),
    (0x01C9, 0x01C7, 1, 1), (0x01CC, 0x01CA, 1, 1), (0x01CE, 0x01CD, 8, 2),
    (0x01DD, 0x018E, 1, 1), (0x01DF, 0x01DE, 9, 2), (0x01F3, 0x01F1, 1, 1),
    (0x01F5, 0x01F4, 1, 1), (0x01F9, 0x01F8, 20, 2), (0x0223, 0x0222, 9, 2),
    (0x023C, 0x023B, 1, 1), (0x0242, 0x0241, 1, 1), (0x0247, 0x0246, 5, 2),
    (0x0250, 0x2C6F, 1, 1), (0x0251, 0x2C6D, 1, 1), (0x0253, 0x0181, 1, 1),
    (0x0254, 0x0186, 1, 1), (0x0256, 0x0189, 2, 1), (0x0259, 0x018F, 1, 1),
    (0x025B, 0x0190, 1, 1), (0x0260, 0x0193, 1, 1), (0x0263, 0x0194, 1, 1),
    (0x0268, 0x0197, 1, 1), (0x0269, 0x0196, 1, 1), (0x026B, 0x2C62, 1, 1),
    (0x026F, 0x019C, 1, 1), (0x0271, 0x2C6E, 1, 1), (0x0272, 0x019D, 1, 1),
    (0x0275, 0x019F, 1, 1), (0x027D, 0x2C64, 1, 1), (0x0280, 0x01A6, 1, 1),
    (0x0283, 0x01A9, 1, 1), (0x0288, 0x01AE, 1, 1), (0x0289, 0x0244, 1, 1),
    (0x028A, 0x01B1, 2, 1), (0x028C, 0x0245, 1, 1), (0x0292, 0x01B7, 1, 1),
    (0x0371, 0x0370, 2, 2), (0x0377, 0x0376, 1, 1), (0x037B, 0x03FD, 3, 1),
    (0x03AC, 0x0386, 1, 1), (0x03AD, 0x0388, 3, 1), (0x03B1, 0x0391, 17, 1),
    (0x03C3, 0x03A3, 9, 1), (0x03CC, 0x038C, 1, 1), (0x03CD, 0x038E, 2, 1),
    (0x03D7, 0x03CF, 1, 1), (0x03D9, 0x03D8, 12, 2), (0x03F2, 0x03F9, 1, 1),
    (0x03F8, 0x03F7, 1, 1), (0x03FB, 0x03FA, 1, 1), (0x0430, 0x0410, 32, 1),
    (0x0450, 0x0400, 16, 1), (0x0461, 0x0460, 17, 2), (0x048B, 0x048A, 27, 2),
    (0x04C2, 0x04C1, 7, 2), (0x04CF, 0x04C0, 1, 1), (0x04D1, 0x04D0, 42, 2),
    (0x0561, 0x0531, 38, 1), (0x1D79, 0xA77D, 1, 1), (0x1D7D, 0x2C63, 1, 1),
    (0x1E01, 0x1E00, 75, 2), (0x1EA1, 0x1EA0, 48, 2), (0x1F00, 0x1F08, 8, 1),
    (0x1F10, 0x1F18, 6, 1), (0x1F20, 0x1F28, 8, 1), (0x1F30, 0x1F38, 8, 1),
    (0x1F40, 0x1F48, 6, 1), (0x1F51, 0x1F59, 4, 2), (0x1F60, 0x1F68, 8, 1),
    (0x1F70, 0x1FBA, 2, 1), (0x1F72, 0x1FC8, 4, 1), (0x1F76, 0x1FDA, 2, 1),
    (0x1F78, 0x1FF8, 2, 1), (0x1F7A, 0x1FEA, 2, 1), (0x1F7C, 0x1FFA, 2, 1),
    (0x1F80, 0x1F88, 8, 1), (0x1F90, 0x1F98, 8, 1), (0x1FA0, 0x1FA8, 8, 1),
    (0x1FB0, 0x1FB8, 2, 1), (0x1FB3, 0x1FBC, 1, 1), (0x1FC3, 0x1FCC, 1, 1),
    (0x1FD0, 0x1FD8, 2, 1), (0x1FE0, 0x1FE8, 2, 1), (0x1FE5, 0x1FEC, 1, 1),
    (0x1FF3, 0x1FFC, 1, 1), (0x214E, 0x2132, 1, 1), (0x2170, 0x2160, 16, 1),
    (0x2184, 0x2183, 1, 1), (0x24D0, 0x24B6, 26, 1), (0x2C30, 0x2C00, 47, 1),
    (0x2C61, 0x2C60, 1, 1), (0x2C65, 0x023A, 1, 1), (0x2C66, 0x023E, 1, 1),
    (0x2C68, 0x2C67, 3, 2), (0x2C73, 0x2C72, 1, 1), (0x2C76, 0x2C75, 1, 1),
    (0x2C81, 0x2C80, 50, 2), (0x2D00, 0x10A0, 38, 1), (0xA641, 0xA640, 16, 2),
    (0xA663, 0xA662, 6, 2), (0xA681, 0xA680, 12, 2), (0xA723, 0xA722, 7, 2),
    (0xA733, 0xA732, 31, 2), (0xA77A, 0xA779, 2, 2), (0xA77F, 0xA77E, 5, 2),
    (0xA78C, 0xA78B, 1, 1), (0xFF41, 0xFF21, 26, 1),
];

/// The upper-case form of one UTF-16 unit, or the unit itself.
pub fn upper_unit(unit: u16) -> u16 {
    for &(lower, upper, count, stride) in RUNS {
        if unit >= lower && unit < lower + count * stride && (unit - lower) % stride == 0 {
            return upper + (unit - lower);
        }
    }
    unit
}

/// The lower-case form of one UTF-16 unit, or the unit itself.
pub fn lower_unit(unit: u16) -> u16 {
    for &(lower, upper, count, stride) in RUNS {
        if unit >= upper && unit < upper + count * stride && (unit - upper) % stride == 0 {
            return lower + (unit - upper);
        }
    }
    unit
}

/// Text in upper case, one unit for one unit.
pub fn upper(text: &str) -> String {
    map(text, upper_unit)
}

/// Text in lower case, one unit for one unit.
pub fn lower(text: &str) -> String {
    map(text, lower_unit)
}

fn map(text: &str, unit: fn(u16) -> u16) -> String {
    if text.is_ascii() {
        return text.chars().map(|ch| char::from(unit(ch as u16) as u8)).collect();
    }
    let units: Vec<u16> = text.encode_utf16().map(unit).collect();
    String::from_utf16_lossy(&units)
}

/// Letters a text comparison (StrComp, Option Compare Text) folds together
/// beyond the case table: measured over every letter Unicode and the table
/// case differently, these 62 compare equal to the letter they fold to --
/// the titlecase digraphs, the final sigma, the Kelvin, Angstrom and Ohm
/// signs, the capital sharp s -- while the rest (the dotless i, the long s,
/// Georgian Mtavruli, Cherokee small letters, ...) stay apart.
const COMPARE_EXTRA: &[(u16, u16)] = &[
    (0x01C5, 0x01C6), (0x01C8, 0x01C9), (0x01CB, 0x01CC), (0x01F2, 0x01F3), (0x03C2, 0x03C3),
    (0x03F5, 0x03B5), (0x0524, 0x0525), (0x0526, 0x0527), (0x0528, 0x0529), (0x052A, 0x052B),
    (0x052C, 0x052D), (0x052E, 0x052F), (0x10C7, 0x2D27), (0x10CD, 0x2D2D), (0x13F5, 0x13FD),
    (0x1CBD, 0x10FD), (0x1CBE, 0x10FE), (0x1CBF, 0x10FF), (0x1E9E, 0x00DF), (0x1FBE, 0x03B9),
    (0x2126, 0x03C9), (0x212A, 0x006B), (0x212B, 0x00E5), (0x2C2F, 0x2C5F), (0x2C70, 0x0252),
    (0x2C7E, 0x023F), (0x2C7F, 0x0240), (0x2CEB, 0x2CEC), (0x2CED, 0x2CEE), (0x2CF2, 0x2CF3),
    (0xA660, 0xA661), (0xA698, 0xA699), (0xA69A, 0xA69B), (0xA78D, 0x0265), (0xA790, 0xA791),
    (0xA792, 0xA793), (0xA796, 0xA797), (0xA798, 0xA799), (0xA79A, 0xA79B), (0xA79C, 0xA79D),
    (0xA79E, 0xA79F), (0xA7A0, 0xA7A1), (0xA7A2, 0xA7A3), (0xA7A4, 0xA7A5), (0xA7A6, 0xA7A7),
    (0xA7A8, 0xA7A9), (0xA7B3, 0xAB53), (0xA7B4, 0xA7B5), (0xA7B6, 0xA7B7), (0xA7B8, 0xA7B9),
    (0xA7BA, 0xA7BB), (0xA7BC, 0xA7BD), (0xA7BE, 0xA7BF), (0xA7C0, 0xA7C1), (0xA7C2, 0xA7C3),
    (0xA7C4, 0xA794), (0xA7C7, 0xA7C8), (0xA7C9, 0xA7CA), (0xA7D0, 0xA7D1), (0xA7D6, 0xA7D7),
    (0xA7D8, 0xA7D9), (0xA7F5, 0xA7F6),
];

/// One UTF-16 unit as a text comparison reads it.
pub fn compare_unit(unit: u16) -> u16 {
    match COMPARE_EXTRA.binary_search_by_key(&unit, |&(from, _)| from) {
        Ok(at) => COMPARE_EXTRA[at].1,
        Err(_) => lower_unit(unit),
    }
}

/// Text as a text comparison reads it, one unit for one unit.
pub fn compare_fold(text: &str) -> String {
    map(text, compare_unit)
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn letters_map_one_for_one() {
        assert_eq!(upper("stra\u{df}e \u{e9}\u{ff}"), "STRA\u{df}E \u{c9}\u{178}");
        assert_eq!(lower("\u{130}\u{1e9e}ABC"), "\u{130}\u{1e9e}abc");
        assert_eq!(upper("\u{3c2}\u{3c3}\u{10d0}\u{24d0}\u{ff41}"), "\u{3c2}\u{3a3}\u{10d0}\u{24b6}\u{ff21}");
    }

    #[test]
    fn comparison_folds_a_few_more() {
        assert_eq!(compare_fold("\u{3c2}\u{3a3}\u{1c5}\u{212a}"), "\u{3c3}\u{3c3}\u{1c6}k");
        assert_eq!(compare_fold("\u{131}\u{17f}\u{1c90}"), "\u{131}\u{17f}\u{1c90}");
    }
}
