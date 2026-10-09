// This Source Code Form is subject to the terms of the Mozilla Public
// License, v. 2.0. If a copy of the MPL was not distributed with this
// file, You can obtain one at https://mozilla.org/MPL/2.0/.

//! Static TrueType layout bounds from the same programme used for painting.

use skrifa::raw::{FileRef, TableProvider};
use skrifa::raw::types::{GlyphId, Tag};

/// Read the declared design-unit bounds of a static TrueType glyph.
/// Composite transforms may contain fractional coordinates: reconstructing
/// and truncating them can disagree with the source's integer glyph header.
/// Variable and non-TrueType outlines must retain their derived bounds.
pub fn static_glyph_bounds(bytes: &[u8], face_index: u32, glyph: u16) -> Option<[i32; 4]> {
    let font = match FileRef::new(bytes).ok()? {
        FileRef::Font(font) if face_index == 0 => font,
        FileRef::Collection(collection) => collection.get(face_index).ok()?,
        _ => return None,
    };
    if font.table_data(Tag::new(b"fvar")).is_some() || glyph >= font.maxp().ok()?.num_glyphs() {
        return None;
    }
    let loca = font.loca(None).ok()?;
    let glyf = font.glyf().ok()?;
    let glyph = loca.get_glyf(GlyphId::new(u32::from(glyph)), &glyf).ok()?;
    let Some(glyph) = glyph else { return Some([0; 4]); };
    if glyph.number_of_contours() == 0 { return Some([0; 4]); }
    Some([i32::from(glyph.x_min()), i32::from(glyph.y_min()),
          i32::from(glyph.x_max()), i32::from(glyph.y_max())])
}

#[cfg(test)]
mod tests {
    use super::static_glyph_bounds;

    // Authored parser fixtures contain only table headers; no font names or
    // third-party programme/contour data is embedded in these tests.
    fn sfnt(base: u32, long: bool, variable: bool, empty: bool, bounds: [i16; 4]) -> Vec<u8> {
        let mut head = vec![0; 54];
        head[50..52].copy_from_slice(&i16::from(long).to_be_bytes());
        let mut maxp = 0x0000_5000u32.to_be_bytes().to_vec();
        maxp.extend_from_slice(&1u16.to_be_bytes());
        let mut glyf = (-1i16).to_be_bytes().to_vec();
        for value in bounds { glyf.extend_from_slice(&value.to_be_bytes()); }
        let end = if empty { 0 } else { glyf.len() as u32 };
        let loca = if long { [0, end].into_iter().flat_map(u32::to_be_bytes).collect() }
                   else { [0, (end/2) as u16].into_iter().flat_map(u16::to_be_bytes).collect() };
        let mut tables = vec![(*b"glyf", glyf), (*b"head", head), (*b"loca", loca), (*b"maxp", maxp)];
        if variable { tables.push((*b"fvar", vec![0; 16])); }
        tables.sort_by_key(|(tag, _)| *tag);
        let mut data = 0x0001_0000u32.to_be_bytes().to_vec();
        for value in [tables.len() as u16, 0, 0, 0] { data.extend_from_slice(&value.to_be_bytes()); }
        let mut offset = 12+16*tables.len() as u32;
        for (tag, bytes) in &tables {
            data.extend_from_slice(tag);
            for value in [0, base+offset, bytes.len() as u32] { data.extend_from_slice(&value.to_be_bytes()); }
            offset += (bytes.len() as u32+3)&!3;
        }
        for (_, bytes) in tables {
            data.extend_from_slice(&bytes);
            while data.len()%4 != 0 { data.push(0); }
        }
        data
    }

    #[test]
    fn preserves_declared_composite_bounds_for_both_loca_formats() {
        for long in [false, true] {
            let bytes=sfnt(0, long, false, false, [173, 2, 1653, 1480]);
            assert_eq!(static_glyph_bounds(&bytes, 0, 0), Some([173, 2, 1653, 1480]));
        }
    }

    #[test]
    fn empty_glyphs_have_no_ink_and_invalid_ids_do_not_alias_them() {
        let bytes=sfnt(0, false, false, true, [10, 20, 30, 40]);
        assert_eq!(static_glyph_bounds(&bytes, 0, 0), Some([0; 4]));
        assert_eq!(static_glyph_bounds(&bytes, 0, 1), None);
        assert_eq!(static_glyph_bounds(&bytes, 1, 0), None);
    }

    #[test]
    fn variable_programmes_do_not_reuse_default_header_bounds() {
        let bytes=sfnt(0, false, true, false, [-50, -20, 100, 200]);
        assert_eq!(static_glyph_bounds(&bytes, 0, 0), None);
    }

    #[test]
    fn reads_selected_collection_member_and_rejects_missing_members() {
        let first=sfnt(20, false, false, false, [10, 20, 30, 40]);
        let second_offset=20+first.len() as u32;
        let second=sfnt(second_offset, true, false, false, [-100, -200, 300, 400]);
        let mut bytes=b"ttcf".to_vec();
        for value in [0x0001_0000u32, 2, 20, second_offset] { bytes.extend_from_slice(&value.to_be_bytes()); }
        bytes.extend(first); bytes.extend(second);
        assert_eq!(static_glyph_bounds(&bytes, 0, 0), Some([10, 20, 30, 40]));
        assert_eq!(static_glyph_bounds(&bytes, 1, 0), Some([-100, -200, 300, 400]));
        assert_eq!(static_glyph_bounds(&bytes, 2, 0), None);
    }

    #[test]
    fn malformed_offsets_and_truncated_tables_are_rejected_without_panicking() {
        assert_eq!(static_glyph_bounds(b"", 0, 0), None);
        let original=sfnt(0, false, false, false, [1, 2, 3, 4]);
        for length in 0..original.len() {
            let _=static_glyph_bounds(&original[..length], 0, 0);
        }
        let mut bytes=original;
        let loca_record=12+2*16;
        let offset=u32::from_be_bytes(bytes[loca_record+8..loca_record+12].try_into().unwrap()) as usize;
        bytes[offset..offset+2].copy_from_slice(&20u16.to_be_bytes());
        assert_eq!(static_glyph_bounds(&bytes, 0, 0), None);
    }
}
