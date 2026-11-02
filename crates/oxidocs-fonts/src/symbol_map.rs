// This Source Code Form is subject to the terms of the Mozilla Public
// License, v. 2.0. If a copy of the MPL was not distributed with this
// file, You can obtain one at https://mozilla.org/MPL/2.0/.

//! Unicode for the characters of a symbol font.
//!
//! A .docx that uses Symbol or Wingdings stores the font's own codes, usually
//! in the U+F020..U+F0FF private-use block (a `w:sym` or a run in that font).
//! Word draws them with the font. Where the font is not on the machine -- it
//! is Microsoft's and cannot be shipped -- the only faithful thing left is the
//! Unicode character that names the same symbol, drawn from a face that has it.

/// The Unicode character the symbol font `family` draws for `ch`, when `ch` is
/// one of the font's codes (U+F020..U+F0FF, or the same code below U+0100) and
/// the symbol has a Unicode equivalent. `None` for any other family or code.
pub fn symbol_font_to_unicode(family: &str, ch: char) -> Option<char> {
    let code = match ch as u32 {
        c @ 0xF020..=0xF0FF => c - 0xF000,
        c @ 0x21..=0xFF => c,
        _ => return None,
    } as u8;
    let table: &[(u8, char)] = match family.to_ascii_lowercase().replace(' ', "").as_str() {
        "symbol" => ADOBE_SYMBOL,
        "wingdings" => WINGDINGS,
        _ => return None,
    };
    table.binary_search_by_key(&code, |&(c, _)| c).ok().map(|i| table[i].1)
}

/// Adobe Symbol Encoding to Unicode, from the Unicode Consortium's
/// VENDORS/ADOBE/symbol.txt (table version 1.0, 2011-07-12), which Unicode,
/// Inc. grants the right to use freely. Codes that file maps into Adobe's
/// private-use bracket pieces are left out; the first of two mappings is kept.
const ADOBE_SYMBOL: &[(u8, char)] = &[
    (0x21, '\u{0021}'), (0x22, '\u{2200}'), (0x23, '\u{0023}'), (0x24, '\u{2203}'), (0x25, '\u{0025}'), (0x26, '\u{0026}'),
    (0x27, '\u{220B}'), (0x28, '\u{0028}'), (0x29, '\u{0029}'), (0x2A, '\u{2217}'), (0x2B, '\u{002B}'), (0x2C, '\u{002C}'),
    (0x2D, '\u{2212}'), (0x2E, '\u{002E}'), (0x2F, '\u{002F}'), (0x30, '\u{0030}'), (0x31, '\u{0031}'), (0x32, '\u{0032}'),
    (0x33, '\u{0033}'), (0x34, '\u{0034}'), (0x35, '\u{0035}'), (0x36, '\u{0036}'), (0x37, '\u{0037}'), (0x38, '\u{0038}'),
    (0x39, '\u{0039}'), (0x3A, '\u{003A}'), (0x3B, '\u{003B}'), (0x3C, '\u{003C}'), (0x3D, '\u{003D}'), (0x3E, '\u{003E}'),
    (0x3F, '\u{003F}'), (0x40, '\u{2245}'), (0x41, '\u{0391}'), (0x42, '\u{0392}'), (0x43, '\u{03A7}'), (0x44, '\u{0394}'),
    (0x45, '\u{0395}'), (0x46, '\u{03A6}'), (0x47, '\u{0393}'), (0x48, '\u{0397}'), (0x49, '\u{0399}'), (0x4A, '\u{03D1}'),
    (0x4B, '\u{039A}'), (0x4C, '\u{039B}'), (0x4D, '\u{039C}'), (0x4E, '\u{039D}'), (0x4F, '\u{039F}'), (0x50, '\u{03A0}'),
    (0x51, '\u{0398}'), (0x52, '\u{03A1}'), (0x53, '\u{03A3}'), (0x54, '\u{03A4}'), (0x55, '\u{03A5}'), (0x56, '\u{03C2}'),
    (0x57, '\u{03A9}'), (0x58, '\u{039E}'), (0x59, '\u{03A8}'), (0x5A, '\u{0396}'), (0x5B, '\u{005B}'), (0x5C, '\u{2234}'),
    (0x5D, '\u{005D}'), (0x5E, '\u{22A5}'), (0x5F, '\u{005F}'), (0x61, '\u{03B1}'), (0x62, '\u{03B2}'), (0x63, '\u{03C7}'),
    (0x64, '\u{03B4}'), (0x65, '\u{03B5}'), (0x66, '\u{03C6}'), (0x67, '\u{03B3}'), (0x68, '\u{03B7}'), (0x69, '\u{03B9}'),
    (0x6A, '\u{03D5}'), (0x6B, '\u{03BA}'), (0x6C, '\u{03BB}'), (0x6D, '\u{00B5}'), (0x6E, '\u{03BD}'), (0x6F, '\u{03BF}'),
    (0x70, '\u{03C0}'), (0x71, '\u{03B8}'), (0x72, '\u{03C1}'), (0x73, '\u{03C3}'), (0x74, '\u{03C4}'), (0x75, '\u{03C5}'),
    (0x76, '\u{03D6}'), (0x77, '\u{03C9}'), (0x78, '\u{03BE}'), (0x79, '\u{03C8}'), (0x7A, '\u{03B6}'), (0x7B, '\u{007B}'),
    (0x7C, '\u{007C}'), (0x7D, '\u{007D}'), (0x7E, '\u{223C}'), (0xA0, '\u{20AC}'), (0xA1, '\u{03D2}'), (0xA2, '\u{2032}'),
    (0xA3, '\u{2264}'), (0xA4, '\u{2044}'), (0xA5, '\u{221E}'), (0xA6, '\u{0192}'), (0xA7, '\u{2663}'), (0xA8, '\u{2666}'),
    (0xA9, '\u{2665}'), (0xAA, '\u{2660}'), (0xAB, '\u{2194}'), (0xAC, '\u{2190}'), (0xAD, '\u{2191}'), (0xAE, '\u{2192}'),
    (0xAF, '\u{2193}'), (0xB0, '\u{00B0}'), (0xB1, '\u{00B1}'), (0xB2, '\u{2033}'), (0xB3, '\u{2265}'), (0xB4, '\u{00D7}'),
    (0xB5, '\u{221D}'), (0xB6, '\u{2202}'), (0xB7, '\u{2022}'), (0xB8, '\u{00F7}'), (0xB9, '\u{2260}'), (0xBA, '\u{2261}'),
    (0xBB, '\u{2248}'), (0xBC, '\u{2026}'), (0xBF, '\u{21B5}'), (0xC0, '\u{2135}'), (0xC1, '\u{2111}'), (0xC2, '\u{211C}'),
    (0xC3, '\u{2118}'), (0xC4, '\u{2297}'), (0xC5, '\u{2295}'), (0xC6, '\u{2205}'), (0xC7, '\u{2229}'), (0xC8, '\u{222A}'),
    (0xC9, '\u{2283}'), (0xCA, '\u{2287}'), (0xCB, '\u{2284}'), (0xCC, '\u{2282}'), (0xCD, '\u{2286}'), (0xCE, '\u{2208}'),
    (0xCF, '\u{2209}'), (0xD0, '\u{2220}'), (0xD1, '\u{2207}'), (0xD5, '\u{220F}'), (0xD6, '\u{221A}'), (0xD7, '\u{22C5}'),
    (0xD8, '\u{00AC}'), (0xD9, '\u{2227}'), (0xDA, '\u{2228}'), (0xDB, '\u{21D4}'), (0xDC, '\u{21D0}'), (0xDD, '\u{21D1}'),
    (0xDE, '\u{21D2}'), (0xDF, '\u{21D3}'), (0xE0, '\u{25CA}'), (0xE1, '\u{2329}'), (0xE5, '\u{2211}'), (0xF1, '\u{232A}'),
    (0xF2, '\u{222B}'), (0xF3, '\u{2320}'), (0xF5, '\u{2321}'),
];

/// Wingdings codes with an unambiguous Unicode equivalent: the bullets, boxes,
/// check marks, arrows and numbered circles documents use. A code whose
/// equivalent is in doubt is left out, so it stays a missing glyph rather than
/// becoming the wrong symbol.
const WINGDINGS: &[(u8, char)] = &[
    (0x22, '✂'), (0x23, '✁'), (0x29, '✆'), (0x36, '⌛'), (0x3E, '✇'), (0x3F, '✍'),
    (0x41, '✌'), (0x4A, '☺'), (0x4C, '☹'), (0x4E, '☠'), (0x51, '✈'), (0x52, '☼'),
    (0x54, '❄'), (0x56, '✞'), (0x58, '✠'), (0x59, '✡'), (0x5A, '☪'), (0x5B, '☯'),
    (0x5C, 'ॐ'), (0x5D, '☸'), (0x5E, '♈'), (0x5F, '♉'), (0x60, '♊'), (0x61, '♋'),
    (0x62, '♌'), (0x63, '♍'), (0x64, '♎'), (0x65, '♏'), (0x66, '♐'), (0x67, '♑'),
    (0x68, '♒'), (0x69, '♓'), (0x6C, '●'), (0x6E, '■'), (0x6F, '□'), (0x71, '❑'),
    (0x72, '❒'), (0x75, '◆'), (0x76, '❖'), (0x78, '⌧'), (0x7A, '⌘'), (0x80, '⓪'),
    (0x81, '①'), (0x82, '②'), (0x83, '③'), (0x84, '④'), (0x85, '⑤'), (0x86, '⑥'),
    (0x87, '⑦'), (0x88, '⑧'), (0x89, '⑨'), (0x8A, '⑩'), (0x8B, '⓿'), (0x8C, '❶'),
    (0x8D, '❷'), (0x8E, '❸'), (0x8F, '❹'), (0x90, '❺'), (0x91, '❻'), (0x92, '❼'),
    (0x93, '❽'), (0x94, '❾'), (0x95, '❿'), (0x9E, '·'), (0x9F, '•'), (0xA4, '◉'),
    (0xA5, '◎'), (0xA7, '▪'), (0xA8, '◻'), (0xAA, '✦'), (0xAB, '★'), (0xAC, '✶'),
    (0xAD, '✴'), (0xAE, '✹'), (0xAF, '✵'), (0xB1, '⌖'), (0xB5, '✪'), (0xB6, '✰'),
    (0xD5, '⌫'), (0xD6, '⌦'), (0xD8, '➢'), (0xEF, '⇦'), (0xF0, '⇨'), (0xF1, '⇧'),
    (0xF2, '⇩'), (0xF3, '⬄'), (0xF4, '⇳'), (0xFB, '✘'), (0xFC, '✔'), (0xFD, '☒'),
    (0xFE, '☑'),
];

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn tables_are_sorted_for_binary_search() {
        for table in [ADOBE_SYMBOL, WINGDINGS] {
            assert!(table.windows(2).all(|w| w[0].0 < w[1].0));
        }
    }

    #[test]
    fn word_bullets_and_check_boxes() {
        // Word's default bullet, and the ones in its bullet gallery
        assert_eq!(symbol_font_to_unicode("Symbol", '\u{F0B7}'), Some('•'));
        assert_eq!(symbol_font_to_unicode("Wingdings", '\u{F0A7}'), Some('▪'));
        assert_eq!(symbol_font_to_unicode("Wingdings", '\u{F076}'), Some('❖'));
        assert_eq!(symbol_font_to_unicode("Wingdings", '\u{F0D8}'), Some('➢'));
        assert_eq!(symbol_font_to_unicode("Wingdings", '\u{F0FC}'), Some('✔'));
        assert_eq!(symbol_font_to_unicode("Wingdings", '\u{F06C}'), Some('●'));
        assert_eq!(symbol_font_to_unicode("Wingdings", '\u{F0FE}'), Some('☑'));
        // a run typed in the font, not a w:sym
        assert_eq!(symbol_font_to_unicode("Wingdings", 'l'), Some('●'));
        assert_eq!(symbol_font_to_unicode("Symbol", 'a'), Some('α'));
    }

    #[test]
    fn other_families_and_unknown_codes_are_left_alone() {
        assert_eq!(symbol_font_to_unicode("Arial", '\u{F0B7}'), None);
        assert_eq!(symbol_font_to_unicode("Wingdings", '\u{F07F}'), None);
        assert_eq!(symbol_font_to_unicode("Wingdings", 'あ'), None);
    }
}
