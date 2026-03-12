//! Standard PDF font encodings.
//!
//! When a font has no /ToUnicode CMap, we fall back to its /Encoding
//! (or the default encoding) to map character codes to Unicode.

/// Decode a byte using WinAnsiEncoding to a char.
/// WinAnsiEncoding is based on Windows code page 1252.
pub fn win_ansi_decode(byte: u8) -> char {
    match byte {
        // Standard ASCII range
        0x00..=0x7F => byte as char,
        // Windows-1252 specific mappings (0x80-0x9F differ from Latin-1)
        0x80 => '\u{20AC}', // Euro sign
        0x81 => '\u{FFFD}', // undefined
        0x82 => '\u{201A}', // Single low-9 quotation mark
        0x83 => '\u{0192}', // Latin small letter f with hook
        0x84 => '\u{201E}', // Double low-9 quotation mark
        0x85 => '\u{2026}', // Horizontal ellipsis
        0x86 => '\u{2020}', // Dagger
        0x87 => '\u{2021}', // Double dagger
        0x88 => '\u{02C6}', // Modifier letter circumflex accent
        0x89 => '\u{2030}', // Per mille sign
        0x8A => '\u{0160}', // Latin capital letter S with caron
        0x8B => '\u{2039}', // Single left-pointing angle quotation mark
        0x8C => '\u{0152}', // Latin capital ligature OE
        0x8D => '\u{FFFD}', // undefined
        0x8E => '\u{017D}', // Latin capital letter Z with caron
        0x8F => '\u{FFFD}', // undefined
        0x90 => '\u{FFFD}', // undefined
        0x91 => '\u{2018}', // Left single quotation mark
        0x92 => '\u{2019}', // Right single quotation mark
        0x93 => '\u{201C}', // Left double quotation mark
        0x94 => '\u{201D}', // Right double quotation mark
        0x95 => '\u{2022}', // Bullet
        0x96 => '\u{2013}', // En dash
        0x97 => '\u{2014}', // Em dash
        0x98 => '\u{02DC}', // Small tilde
        0x99 => '\u{2122}', // Trade mark sign
        0x9A => '\u{0161}', // Latin small letter s with caron
        0x9B => '\u{203A}', // Single right-pointing angle quotation mark
        0x9C => '\u{0153}', // Latin small ligature oe
        0x9D => '\u{FFFD}', // undefined
        0x9E => '\u{017E}', // Latin small letter z with caron
        0x9F => '\u{0178}', // Latin capital letter Y with diaeresis
        // 0xA0-0xFF are same as Latin-1
        _ => byte as char,
    }
}

/// Decode a byte using MacRomanEncoding to a char.
pub fn mac_roman_decode(byte: u8) -> char {
    match byte {
        0x00..=0x7F => byte as char,
        0x80 => '\u{00C4}', // Ä
        0x81 => '\u{00C5}', // Å
        0x82 => '\u{00C7}', // Ç
        0x83 => '\u{00C9}', // É
        0x84 => '\u{00D1}', // Ñ
        0x85 => '\u{00D6}', // Ö
        0x86 => '\u{00DC}', // Ü
        0x87 => '\u{00E1}', // á
        0x88 => '\u{00E0}', // à
        0x89 => '\u{00E2}', // â
        0x8A => '\u{00E4}', // ä
        0x8B => '\u{00E3}', // ã
        0x8C => '\u{00E5}', // å
        0x8D => '\u{00E7}', // ç
        0x8E => '\u{00E9}', // é
        0x8F => '\u{00E8}', // è
        0x90 => '\u{00EA}', // ê
        0x91 => '\u{00EB}', // ë
        0x92 => '\u{00ED}', // í
        0x93 => '\u{00EC}', // ì
        0x94 => '\u{00EE}', // î
        0x95 => '\u{00EF}', // ï
        0x96 => '\u{00F1}', // ñ
        0x97 => '\u{00F3}', // ó
        0x98 => '\u{00F2}', // ò
        0x99 => '\u{00F4}', // ô
        0x9A => '\u{00F6}', // ö
        0x9B => '\u{00F5}', // õ
        0x9C => '\u{00FA}', // ú
        0x9D => '\u{00F9}', // ù
        0x9E => '\u{00FB}', // û
        0x9F => '\u{00FC}', // ü
        0xA0 => '\u{2020}', // †
        0xA1 => '\u{00B0}', // °
        0xA5 => '\u{2022}', // •
        0xC1 => '\u{00C0}', // grave double
        0xC7 => '\u{00AB}', // «
        0xC8 => '\u{00BB}', // »
        0xC9 => '\u{2026}', // …
        0xCA => '\u{00A0}', // non-breaking space
        0xD0 => '\u{2013}', // –
        0xD1 => '\u{2014}', // —
        0xD2 => '\u{201C}', // "
        0xD3 => '\u{201D}', // "
        0xD4 => '\u{2018}', // '
        0xD5 => '\u{2019}', // '
        _ => byte as char,
    }
}

/// Known font encoding types.
#[derive(Debug, Clone)]
pub enum FontEncoding {
    WinAnsi,
    MacRoman,
    /// Custom /Differences array overrides on top of a base encoding.
    Differences(Vec<(u8, String)>),
}

/// Decode bytes using a FontEncoding.
pub fn decode_with_encoding(bytes: &[u8], encoding: &FontEncoding) -> String {
    match encoding {
        FontEncoding::WinAnsi => bytes.iter().map(|&b| win_ansi_decode(b)).collect(),
        FontEncoding::MacRoman => bytes.iter().map(|&b| mac_roman_decode(b)).collect(),
        FontEncoding::Differences(_diffs) => {
            // For now, fall back to WinAnsi as base
            bytes.iter().map(|&b| win_ansi_decode(b)).collect()
        }
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_win_ansi_ascii() {
        assert_eq!(win_ansi_decode(b'A'), 'A');
        assert_eq!(win_ansi_decode(b' '), ' ');
    }

    #[test]
    fn test_win_ansi_special() {
        assert_eq!(win_ansi_decode(0x80), '\u{20AC}'); // Euro
        assert_eq!(win_ansi_decode(0x93), '\u{201C}'); // Left double quote
        assert_eq!(win_ansi_decode(0x94), '\u{201D}'); // Right double quote
        assert_eq!(win_ansi_decode(0x96), '\u{2013}'); // En dash
        assert_eq!(win_ansi_decode(0x97), '\u{2014}'); // Em dash
    }

    #[test]
    fn test_win_ansi_latin1_range() {
        assert_eq!(win_ansi_decode(0xE9), 'é');
        assert_eq!(win_ansi_decode(0xFC), 'ü');
    }
}
