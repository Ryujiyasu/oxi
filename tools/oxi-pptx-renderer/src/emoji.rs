/* This Source Code Form is subject to the terms of the Mozilla Public
 * License, v. 2.0. If a copy of the MPL was not distributed with this
 * file, You can obtain one at https://mozilla.org/MPL/2.0/. */

//! Colour emoji: the COLR/CPAL tables of a layered colour font.
//!
//! PowerPoint draws emoji in colour and Oxi drew the black outline glyph,
//! because GDI has no COLR support -- `ExtTextOutW` renders a colour font's
//! base outline and stops. The layers are in the font file, so the renderer
//! reads them itself: COLR maps a base glyph to a list of (glyph, palette
//! index) layers, CPAL holds the palette, and drawing the layers back to front
//! at one pen position reproduces the glyph PowerPoint shows.
//!
//! Only the v0 records are read. Segoe UI Emoji declares COLR version 1, but
//! v1 is a superset: its baseGlyphRecords / layerRecords arrays are still
//! populated with the flat layer lists that predate the v1 gradient
//! machinery. Reading them gives 3372 base glyphs on the Windows 11 font,
//! U+1F44B among them at 9 layers starting #FFC83D.
//!
//! This module is deliberately free of Windows types so the parsing can be
//! exercised on hand-built tables; the caller hands in the table bytes.

/// A parsed colour font: enough of COLR, CPAL, cmap and hmtx to place and
/// paint one layered glyph.
pub struct ColorFont {
    colr: Vec<u8>,
    cpal: Vec<u8>,
    cmap: Vec<u8>,
    hmtx: Vec<u8>,
    n_base: usize,
    base_off: usize,
    layer_off: usize,
    n_layer: usize,
    /// Byte offset of the first colour record, and palette 0's first entry.
    pal_first: usize,
    pal0: usize,
    n_records: usize,
    /// Offset of the cmap subtable to look glyphs up in, and its format.
    cmap_sub: usize,
    cmap_fmt: u16,
    num_h_metrics: usize,
    upem: f32,
}

fn u16at(b: &[u8], o: usize) -> Option<u16> {
    Some(u16::from_be_bytes([*b.get(o)?, *b.get(o + 1)?]))
}

fn u32at(b: &[u8], o: usize) -> Option<u32> {
    Some(u32::from_be_bytes([
        *b.get(o)?,
        *b.get(o + 1)?,
        *b.get(o + 2)?,
        *b.get(o + 3)?,
    ]))
}

impl ColorFont {
    /// Parse the tables. Returns None unless every one of them is present and
    /// well-formed enough to answer a lookup -- a font without COLR is not a
    /// colour font and the caller must keep its ordinary drawing.
    pub fn from_tables(
        colr: Vec<u8>,
        cpal: Vec<u8>,
        cmap: Vec<u8>,
        hmtx: Vec<u8>,
        hhea: &[u8],
        head: &[u8],
    ) -> Option<Self> {
        let n_base = u16at(&colr, 2)? as usize;
        let base_off = u32at(&colr, 4)? as usize;
        let layer_off = u32at(&colr, 8)? as usize;
        let n_layer = u16at(&colr, 12)? as usize;
        if n_base == 0 || base_off + n_base * 6 > colr.len() {
            return None;
        }

        let n_pal = u16at(&cpal, 4)? as usize;
        let n_records = u16at(&cpal, 6)? as usize;
        let pal_first = u32at(&cpal, 8)? as usize;
        if n_pal == 0 || pal_first + n_records * 4 > cpal.len() {
            return None;
        }
        // Palettes are indexed indirectly: palette p's entries start at colour
        // record colorRecordIndices[p]. Only palette 0 is used -- CPAL's other
        // palettes are alternate themes a document has no way to select.
        let pal0 = u16at(&cpal, 12)? as usize;

        // Prefer the format-12 subtable: every emoji above U+FFFF is
        // unreachable through format 4, which is a 16-bit map.
        let n_sub = u16at(&cmap, 2)? as usize;
        let (mut sub, mut fmt) = (0usize, 0u16);
        for i in 0..n_sub {
            let off = u32at(&cmap, 4 + 8 * i + 4)? as usize;
            let f = u16at(&cmap, off)?;
            if f == 12 {
                sub = off;
                fmt = 12;
                break;
            }
            if f == 4 && fmt == 0 {
                sub = off;
                fmt = 4;
            }
        }
        if fmt == 0 {
            return None;
        }

        let num_h_metrics = u16at(hhea, 34)? as usize;
        let upem = u16at(head, 18)? as f32;
        if num_h_metrics == 0 || upem <= 0.0 {
            return None;
        }

        Some(ColorFont {
            colr,
            cpal,
            cmap,
            hmtx,
            n_base,
            base_off,
            layer_off,
            n_layer,
            pal_first,
            pal0,
            n_records,
            cmap_sub: sub,
            cmap_fmt: fmt,
            num_h_metrics,
            upem,
        })
    }

    /// Glyph id for a character, or None when the font does not cover it.
    pub fn gid(&self, ch: char) -> Option<u16> {
        let cp = ch as u32;
        match self.cmap_fmt {
            12 => self.gid_fmt12(cp),
            _ if cp <= 0xFFFF => self.gid_fmt4(cp as u16),
            _ => None,
        }
    }

    fn gid_fmt12(&self, cp: u32) -> Option<u16> {
        let b = &self.cmap;
        let base = self.cmap_sub;
        let n = u32at(b, base + 12)? as usize;
        let (mut lo, mut hi) = (0usize, n);
        while lo < hi {
            let mid = (lo + hi) / 2;
            let g = base + 16 + 12 * mid;
            let start = u32at(b, g)?;
            let end = u32at(b, g + 4)?;
            if cp < start {
                hi = mid;
            } else if cp > end {
                lo = mid + 1;
            } else {
                let first = u32at(b, g + 8)?;
                return u16::try_from(first + (cp - start)).ok();
            }
        }
        None
    }

    fn gid_fmt4(&self, cp: u16) -> Option<u16> {
        let b = &self.cmap;
        let base = self.cmap_sub;
        let seg2 = u16at(b, base + 6)? as usize;
        let ends = base + 14;
        let starts = ends + seg2 + 2;
        let deltas = starts + seg2;
        let ranges = deltas + seg2;
        for i in (0..seg2).step_by(2) {
            if u16at(b, ends + i)? < cp {
                continue;
            }
            if u16at(b, starts + i)? > cp {
                return None;
            }
            let delta = u16at(b, deltas + i)?;
            let ro = u16at(b, ranges + i)?;
            let g = if ro == 0 {
                cp.wrapping_add(delta)
            } else {
                let at = ranges + i + ro as usize + 2 * (cp - u16at(b, starts + i)?) as usize;
                let g = u16at(b, at)?;
                if g == 0 {
                    return None;
                }
                g.wrapping_add(delta)
            };
            return if g == 0 { None } else { Some(g) };
        }
        None
    }

    /// The layers of a base glyph, back to front, as (glyph, palette index).
    /// A palette index of 0xFFFF means "paint in the current text colour".
    pub fn layers(&self, gid: u16) -> Option<Vec<(u16, u16)>> {
        let (mut lo, mut hi) = (0usize, self.n_base);
        while lo < hi {
            let mid = (lo + hi) / 2;
            let rec = self.base_off + 6 * mid;
            let g = u16at(&self.colr, rec)?;
            if gid < g {
                hi = mid;
            } else if gid > g {
                lo = mid + 1;
            } else {
                let first = u16at(&self.colr, rec + 2)? as usize;
                let count = u16at(&self.colr, rec + 4)? as usize;
                if first + count > self.n_layer {
                    return None;
                }
                let mut out = Vec::with_capacity(count);
                for i in 0..count {
                    let l = self.layer_off + 4 * (first + i);
                    out.push((u16at(&self.colr, l)?, u16at(&self.colr, l + 2)?));
                }
                return Some(out);
            }
        }
        None
    }

    /// Palette-0 colour as (r, g, b, a). CPAL stores its records BGRA.
    pub fn color(&self, index: u16) -> Option<(u8, u8, u8, u8)> {
        let rec = self.pal0 + index as usize;
        if rec >= self.n_records {
            return None;
        }
        let o = self.pal_first + 4 * rec;
        Some((
            *self.cpal.get(o + 2)?,
            *self.cpal.get(o + 1)?,
            *self.cpal.get(o)?,
            *self.cpal.get(o + 3)?,
        ))
    }

    /// Advance width of a glyph in em units.
    pub fn advance_em(&self, gid: u16) -> Option<f32> {
        // Past the last long metric every glyph repeats the last advance.
        let i = (gid as usize).min(self.num_h_metrics - 1);
        Some(u16at(&self.hmtx, 4 * i)? as f32 / self.upem)
    }

    /// A font read only for its metrics: no COLR, no CPAL. Segoe UI Symbol
    /// supplies the text-presentation glyphs and is not a colour font, but the
    /// advance still has to come from its own hmtx -- GDI's char-width calls
    /// take a 16-bit code point and cannot be asked about U+1F321.
    pub fn metrics_only(cmap: Vec<u8>, hmtx: Vec<u8>, hhea: &[u8], head: &[u8]) -> Option<Self> {
        // A COLR/CPAL pair with no records: `layers` and `color` then search
        // empty ranges and answer None, which is what a plain font means.
        let mut f =
            Self::from_tables(EMPTY_COLR.to_vec(), EMPTY_CPAL.to_vec(), cmap, hmtx, hhea, head)?;
        f.n_base = 0;
        f.n_records = 0;
        Some(f)
    }
}

/// One base-glyph record so `from_tables` accepts it; `metrics_only` then
/// zeroes the count, leaving a font that answers no colour question.
const EMPTY_COLR: [u8; 20] = [
    0, 0, // version
    0, 1, // numBaseGlyphRecords
    0, 0, 0, 14, // baseGlyphRecordsOffset
    0, 0, 0, 20, // layerRecordsOffset
    0, 0, // numLayerRecords
    0, 0, 0, 0, 0, 0, // the one base glyph record
];
const EMPTY_CPAL: [u8; 14] = [
    0, 0, // version
    0, 1, // numPaletteEntries
    0, 1, // numPalettes
    0, 0, // numColorRecords
    0, 0, 0, 14, // offsetFirstColorRecord
    0, 0, // colorRecordIndices[0]
];

/// Code points whose Unicode Emoji_Presentation is Yes -- the ones a renderer
/// paints in colour with no variation selector asked for.
///
/// PowerPoint honours the property (`emojipres` probe, 2026-08-19): U+2764,
/// U+1F321 and U+1F441 came out of its PDF export monochrome in the run
/// colour, the same three with U+FE0F after them came out in colour, and
/// U+270B / U+231A / U+1F600 were colour with no selector at all. Ranges from
/// Unicode emoji-data.txt.
const EMOJI_PRESENTATION: &[(u32, u32)] = &[
    (0x231A, 0x231B),
    (0x23E9, 0x23EC),
    (0x23F0, 0x23F0),
    (0x23F3, 0x23F3),
    (0x25FD, 0x25FE),
    (0x2614, 0x2615),
    (0x2648, 0x2653),
    (0x267F, 0x267F),
    (0x2693, 0x2693),
    (0x26A1, 0x26A1),
    (0x26AA, 0x26AB),
    (0x26BD, 0x26BE),
    (0x26C4, 0x26C5),
    (0x26CE, 0x26CE),
    (0x26D4, 0x26D4),
    (0x26EA, 0x26EA),
    (0x26F2, 0x26F3),
    (0x26F5, 0x26F5),
    (0x26FA, 0x26FA),
    (0x26FD, 0x26FD),
    (0x2705, 0x2705),
    (0x270A, 0x270B),
    (0x2728, 0x2728),
    (0x274C, 0x274C),
    (0x274E, 0x274E),
    (0x2753, 0x2755),
    (0x2757, 0x2757),
    (0x2795, 0x2797),
    (0x27B0, 0x27B0),
    (0x27BF, 0x27BF),
    (0x2B1B, 0x2B1C),
    (0x2B50, 0x2B50),
    (0x2B55, 0x2B55),
    (0x1F004, 0x1F004),
    (0x1F0CF, 0x1F0CF),
    (0x1F18E, 0x1F18E),
    (0x1F191, 0x1F19A),
    (0x1F1E6, 0x1F1FF),
    (0x1F201, 0x1F201),
    (0x1F21A, 0x1F21A),
    (0x1F22F, 0x1F22F),
    (0x1F232, 0x1F236),
    (0x1F238, 0x1F23A),
    (0x1F250, 0x1F251),
    (0x1F300, 0x1F320),
    (0x1F32D, 0x1F335),
    (0x1F337, 0x1F37C),
    (0x1F37E, 0x1F393),
    (0x1F3A0, 0x1F3CA),
    (0x1F3CF, 0x1F3D3),
    (0x1F3E0, 0x1F3F0),
    (0x1F3F4, 0x1F3F4),
    (0x1F3F8, 0x1F43E),
    (0x1F440, 0x1F440),
    (0x1F442, 0x1F4FC),
    (0x1F4FF, 0x1F53D),
    (0x1F54B, 0x1F54E),
    (0x1F550, 0x1F567),
    (0x1F57A, 0x1F57A),
    (0x1F595, 0x1F596),
    (0x1F5A4, 0x1F5A4),
    (0x1F5FB, 0x1F64F),
    (0x1F680, 0x1F6C5),
    (0x1F6CC, 0x1F6CC),
    (0x1F6D0, 0x1F6D2),
    (0x1F6D5, 0x1F6D7),
    (0x1F6DC, 0x1F6DF),
    (0x1F6EB, 0x1F6EC),
    (0x1F6F4, 0x1F6FC),
    (0x1F7E0, 0x1F7EB),
    (0x1F7F0, 0x1F7F0),
    (0x1F90C, 0x1F93A),
    (0x1F93C, 0x1F945),
    (0x1F947, 0x1F9FF),
    (0x1FA70, 0x1FA7C),
    (0x1FA80, 0x1FA88),
    (0x1FA90, 0x1FABD),
    (0x1FABF, 0x1FAC5),
    (0x1FACE, 0x1FADB),
    (0x1FAE0, 0x1FAE8),
    (0x1FAF0, 0x1FAF8),
];

/// U+FE0F asks for the colour glyph, U+FE0E for the monochrome one.
pub const VS16: char = '\u{FE0F}';
pub const VS15: char = '\u{FE0E}';

/// True when this character is drawn in colour with no selector after it.
pub fn emoji_presentation(ch: char) -> bool {
    let cp = ch as u32;
    EMOJI_PRESENTATION
        .binary_search_by(|(lo, hi)| {
            if cp < *lo {
                std::cmp::Ordering::Greater
            } else if cp > *hi {
                std::cmp::Ordering::Less
            } else {
                std::cmp::Ordering::Equal
            }
        })
        .is_ok()
}

#[cfg(test)]
mod tests {
    use super::*;

    /// A colour font built by hand: gid 7 has two layers, painted from
    /// palette 0 as #112233 and "use the text colour".
    fn tables() -> ColorFont {
        let mut colr = vec![0u8; 14];
        colr[2..4].copy_from_slice(&1u16.to_be_bytes()); // one base glyph
        colr[4..8].copy_from_slice(&14u32.to_be_bytes()); // baseGlyphRecords
        colr[8..12].copy_from_slice(&20u32.to_be_bytes()); // layerRecords
        colr[12..14].copy_from_slice(&2u16.to_be_bytes()); // two layers
        colr.extend_from_slice(&7u16.to_be_bytes()); // gid
        colr.extend_from_slice(&0u16.to_be_bytes()); // firstLayerIndex
        colr.extend_from_slice(&2u16.to_be_bytes()); // numLayers
        colr.extend_from_slice(&11u16.to_be_bytes()); // layer 0 glyph
        colr.extend_from_slice(&0u16.to_be_bytes()); // layer 0 palette entry
        colr.extend_from_slice(&12u16.to_be_bytes()); // layer 1 glyph
        colr.extend_from_slice(&0xFFFFu16.to_be_bytes()); // "text colour"

        let mut cpal = vec![0u8; 14];
        cpal[2..4].copy_from_slice(&1u16.to_be_bytes()); // entries per palette
        cpal[4..6].copy_from_slice(&1u16.to_be_bytes()); // palettes
        cpal[6..8].copy_from_slice(&1u16.to_be_bytes()); // colour records
        cpal[8..12].copy_from_slice(&14u32.to_be_bytes()); // first record
        cpal.extend_from_slice(&[0x33, 0x22, 0x11, 0xFF]); // BGRA

        // cmap with one format-12 group mapping U+1F600..U+1F601 to gid 7..8.
        let mut cmap = vec![0u8; 12];
        cmap[2..4].copy_from_slice(&1u16.to_be_bytes());
        cmap[4..6].copy_from_slice(&3u16.to_be_bytes());
        cmap[6..8].copy_from_slice(&10u16.to_be_bytes());
        cmap[8..12].copy_from_slice(&12u32.to_be_bytes());
        cmap.extend_from_slice(&12u16.to_be_bytes()); // format
        cmap.extend_from_slice(&0u16.to_be_bytes());
        cmap.extend_from_slice(&0u32.to_be_bytes());
        cmap.extend_from_slice(&0u32.to_be_bytes());
        cmap.extend_from_slice(&1u32.to_be_bytes()); // one group
        cmap.extend_from_slice(&0x1F600u32.to_be_bytes());
        cmap.extend_from_slice(&0x1F601u32.to_be_bytes());
        cmap.extend_from_slice(&7u32.to_be_bytes());

        // hmtx: 8 long metrics, glyph 7 advancing 1024 of a 2048 em.
        let mut hmtx = Vec::new();
        for i in 0..8u16 {
            hmtx.extend_from_slice(&(if i == 7 { 1024u16 } else { 512 }).to_be_bytes());
            hmtx.extend_from_slice(&0u16.to_be_bytes());
        }
        let mut hhea = vec![0u8; 36];
        hhea[34..36].copy_from_slice(&8u16.to_be_bytes());
        let mut head = vec![0u8; 54];
        head[18..20].copy_from_slice(&2048u16.to_be_bytes());

        ColorFont::from_tables(colr, cpal, cmap, hmtx, &hhea, &head).unwrap()
    }

    #[test]
    fn maps_a_non_bmp_character_through_format_12() {
        let f = tables();
        assert_eq!(f.gid('\u{1F600}'), Some(7));
        assert_eq!(f.gid('\u{1F601}'), Some(8));
        assert_eq!(f.gid('a'), None);
    }

    #[test]
    fn reads_the_layer_list_back_to_front() {
        let f = tables();
        assert_eq!(f.layers(7), Some(vec![(11, 0), (12, 0xFFFF)]));
        // A glyph with no COLR record is drawn the ordinary way.
        assert_eq!(f.layers(8), None);
        assert!(f.gid('\u{1F600}').and_then(|g| f.layers(g)).is_some());
    }

    #[test]
    fn unswaps_the_bgra_palette_record() {
        let f = tables();
        assert_eq!(f.color(0), Some((0x11, 0x22, 0x33, 0xFF)));
        assert_eq!(f.color(1), None);
    }

    #[test]
    fn advance_comes_back_in_em_units() {
        let f = tables();
        assert_eq!(f.advance_em(7), Some(0.5));
        // Glyph 7 is the last long metric, so every glyph past it repeats it.
        assert_eq!(f.advance_em(900), Some(0.5));
        assert_eq!(f.advance_em(3), Some(0.25));
    }

    #[test]
    fn presentation_follows_the_unicode_property() {
        // The three the probe measured monochrome, and the three it measured
        // in colour with no selector.
        assert!(!emoji_presentation('\u{2764}'));
        assert!(!emoji_presentation('\u{1F321}'));
        assert!(!emoji_presentation('\u{1F441}'));
        assert!(emoji_presentation('\u{270B}'));
        assert!(emoji_presentation('\u{231A}'));
        assert!(emoji_presentation('\u{1F600}'));
        // Ordinary text is not emoji at all.
        assert!(!emoji_presentation('A'));
        assert!(!emoji_presentation('\u{00A9}'));
    }

    #[test]
    fn presentation_ranges_are_sorted_for_the_binary_search() {
        for pair in EMOJI_PRESENTATION.windows(2) {
            assert!(pair[0].1 < pair[1].0, "{:?} then {:?}", pair[0], pair[1]);
        }
        for (lo, hi) in EMOJI_PRESENTATION {
            assert!(lo <= hi);
        }
    }

    #[test]
    fn a_metrics_only_font_answers_no_colour_question() {
        let f = tables();
        let mut hhea = vec![0u8; 36];
        hhea[34..36].copy_from_slice(&8u16.to_be_bytes());
        let mut head = vec![0u8; 54];
        head[18..20].copy_from_slice(&2048u16.to_be_bytes());
        let m = ColorFont::metrics_only(f.cmap.clone(), f.hmtx.clone(), &hhea, &head).unwrap();
        assert_eq!(m.gid('\u{1F600}'), Some(7));
        assert_eq!(m.advance_em(7), Some(0.5));
        assert_eq!(m.layers(7), None);
        assert_eq!(m.color(0), None);
    }
}
