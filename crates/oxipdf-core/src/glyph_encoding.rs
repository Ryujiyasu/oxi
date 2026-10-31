// This Source Code Form is subject to the terms of the Mozilla Public
// License, v. 2.0. If a copy of the MPL was not distributed with this
// file, You can obtain one at https://mozilla.org/MPL/2.0/.

use crate::ir::{EmbeddedFont, FontFormat, PdfDocument, TextSpan};
use crate::writer::escape_name;
use std::collections::{BTreeMap, HashMap};

#[derive(Debug, thiserror::Error, PartialEq)]
pub enum GlyphEncodingError {
    #[error("Selected glyph count does not match Unicode scalar count")]
    InvalidGlyphCount,
    #[error("Selected glyphs require an actual embedded font: {0}")]
    MissingFontProgram(String),
    #[error("Selected glyph is outside the supplied font")]
    InvalidGlyphIndex,
    #[error("A font resource needs more than 65535 character codes")]
    TooManyCodes,
}

pub(crate) struct FontEncoding {
    pub(crate) cid_to_gid: BTreeMap<u16, u16>,
    pub(crate) cid_to_unicode: HashMap<u16, u32>,
    pairs: HashMap<(u16, u32), u16>,
    cmap: HashMap<u32, u16>,
    direct_gids: bool,
}

pub(crate) type FontEncodings = HashMap<String, FontEncoding>;

impl FontEncoding {
    fn new(font: &EmbeddedFont) -> Self {
        Self {
            cid_to_gid: BTreeMap::new(),
            cid_to_unicode: HashMap::new(),
            pairs: HashMap::new(),
            cmap: font.unicode_to_gid.clone(),
            direct_gids: font.format == FontFormat::OpenTypeCff,
        }
    }

    fn glyphs(&self, span: &TextSpan, selected: Option<&[u16]>) -> Vec<u16> {
        selected.map_or_else(
            || {
                span.text
                    .chars()
                    .map(|ch| {
                        if self.cmap.is_empty() {
                            ch as u16
                        } else {
                            self.cmap.get(&(ch as u32)).copied().unwrap_or(0)
                        }
                    })
                    .collect()
            },
            |indices| indices.to_vec(),
        )
    }

    fn add(&mut self, gid: u16, unicode: u32) -> Result<(), GlyphEncodingError> {
        if self.pairs.contains_key(&(gid, unicode)) {
            return Ok(());
        }
        let cid = if self.direct_gids {
            if self
                .cid_to_unicode
                .get(&gid)
                .is_some_and(|old| *old != unicode)
            {
                // A CFF CID is its glyph, so two characters drawn with one
                // glyph -- a no-break space and a space, a Kangxi radical and
                // its unified ideograph, or any two characters the font lacks
                // (.notdef) -- share one code. The page draws the same either
                // way; ToUnicode can name only one, so keep the canonical one.
                // Refusing here threw away whole documents over copy-paste.
                if alias_rank(unicode) < alias_rank(self.cid_to_unicode[&gid]) {
                    self.cid_to_unicode.insert(gid, unicode);
                }
                self.pairs.insert((gid, unicode), gid);
                return Ok(());
            }
            gid
        } else {
            u16::try_from(self.pairs.len() + 1).map_err(|_| GlyphEncodingError::TooManyCodes)?
        };
        self.pairs.insert((gid, unicode), cid);
        self.cid_to_gid.insert(cid, gid);
        self.cid_to_unicode.insert(cid, unicode);
        Ok(())
    }

    pub(crate) fn encode(&self, span: &TextSpan, selected: Option<&[u16]>) -> String {
        use std::fmt::Write;
        let mut hex = String::new();
        for (ch, gid) in span.text.chars().zip(self.glyphs(span, selected)) {
            write!(hex, "{:04X}", self.pairs[&(gid, ch as u32)]).unwrap();
        }
        hex
    }

    pub(crate) fn gid_stream(&self) -> Vec<u8> {
        let max = self.cid_to_gid.keys().copied().max().unwrap_or(0);
        let mut bytes = vec![0; (usize::from(max) + 1) * 2];
        for (&cid, &gid) in &self.cid_to_gid {
            let at = usize::from(cid) * 2;
            bytes[at..at + 2].copy_from_slice(&gid.to_be_bytes());
        }
        bytes
    }
}

/// How poor a choice `unicode` is for the text a shared glyph extracts to:
/// 0 for an ordinary character, 1 for a compatibility or private-use form that
/// names the same thing as an ordinary one.
fn alias_rank(unicode: u32) -> u8 {
    match unicode {
        0x00A0
        | 0x2E80..=0x2EFF // CJK Radicals Supplement
        | 0x2F00..=0x2FDF // Kangxi Radicals
        | 0xE000..=0xF8FF // Private Use Area (symbol fonts)
        | 0xF900..=0xFAFF // CJK Compatibility Ideographs
        | 0x2F800..=0x2FA1F => 1,
        _ => 0,
    }
}

fn ttf_glyph_count(bytes: &[u8]) -> Option<u16> {
    let count = u16::from_be_bytes(bytes.get(4..6)?.try_into().ok()?);
    for i in 0..usize::from(count) {
        let at = 12 + i * 16;
        if bytes.get(at..at + 4)? == b"maxp" {
            let offset = u32::from_be_bytes(bytes.get(at + 8..at + 12)?.try_into().ok()?) as usize;
            return Some(u16::from_be_bytes(
                bytes.get(offset + 4..offset + 6)?.try_into().ok()?,
            ));
        }
    }
    None
}

pub(crate) fn font_encodings(doc: &PdfDocument) -> Result<FontEncodings, GlyphEncodingError> {
    let mut encodings = FontEncodings::new();
    for page in &doc.pages {
        for element in &page.contents {
            let Some(span) = element.text_span() else {
                continue;
            };
            let indices = element.glyph_indices();
            if indices.is_some_and(|ids| ids.len() != span.text.chars().count()) {
                return Err(GlyphEncodingError::InvalidGlyphCount);
            }
            let font = doc.embedded_fonts.get(&span.font_name);
            if indices.is_some() && font.map_or(true, |font| font.data.is_empty()) {
                return Err(GlyphEncodingError::MissingFontProgram(
                    span.font_name.clone(),
                ));
            }
            let Some(font) = font else { continue };
            if let Some(ids) = indices {
                if font.format == FontFormat::TrueType {
                    let count =
                        ttf_glyph_count(&font.data).ok_or(GlyphEncodingError::InvalidGlyphIndex)?;
                    if ids.iter().any(|gid| *gid >= count) {
                        return Err(GlyphEncodingError::InvalidGlyphIndex);
                    }
                }
            }
            let encoding = encodings
                .entry(escape_name(&span.font_name))
                .or_insert_with(|| FontEncoding::new(font));
            for (ch, gid) in span.text.chars().zip(encoding.glyphs(span, indices)) {
                encoding.add(gid, ch as u32)?;
            }
        }
    }
    Ok(encodings)
}

#[cfg(test)]
mod tests {
    use super::*;

    fn cff() -> FontEncoding {
        FontEncoding {
            cid_to_gid: BTreeMap::new(),
            cid_to_unicode: HashMap::new(),
            pairs: HashMap::new(),
            cmap: HashMap::new(),
            direct_gids: true,
        }
    }

    #[test]
    fn cff_glyph_shared_by_two_characters_extracts_the_ordinary_one() {
        // no-break space first, space second: the space wins
        let mut e = cff();
        e.add(1, 0x00A0).unwrap();
        e.add(1, 0x0020).unwrap();
        assert_eq!(e.cid_to_unicode[&1], 0x0020);
        // unified ideograph first, Kangxi radical second: the ideograph stays
        e.add(18580, 0x6238).unwrap();
        e.add(18580, 0x2F3E).unwrap();
        assert_eq!(e.cid_to_unicode[&18580], 0x6238);
        // both characters still encode, to the one shared code
        assert_eq!(e.pairs[&(1, 0x00A0)], 1);
        assert_eq!(e.pairs[&(18580, 0x2F3E)], 18580);
        assert_eq!(e.cid_to_gid.len(), 2);
    }

    #[test]
    fn cff_notdef_shared_by_missing_characters_keeps_the_first() {
        let mut e = cff();
        e.add(0, 0xF0FE).unwrap();
        e.add(0, 0xF06C).unwrap();
        assert_eq!(e.cid_to_unicode[&0], 0xF0FE);
        assert_eq!(e.pairs[&(0, 0xF06C)], 0);
    }

    #[test]
    fn truetype_glyph_shared_by_two_characters_gets_two_codes() {
        let mut e = FontEncoding { direct_gids: false, ..cff() };
        e.add(1, 0x0020).unwrap();
        e.add(1, 0x00A0).unwrap();
        assert_ne!(e.pairs[&(1, 0x0020)], e.pairs[&(1, 0x00A0)]);
    }
}
