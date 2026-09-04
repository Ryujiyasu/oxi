// This Source Code Form is subject to the terms of the Mozilla Public
// License, v. 2.0. If a copy of the MPL was not distributed with this
// file, You can obtain one at https://mozilla.org/MPL/2.0/.

use std::collections::HashMap;

use crate::ir::{EmbeddedFont, FontFormat};

/// Find a table in a TTF/OTF font by its 4-byte tag.
/// Returns `(offset, length)` within `data`.
pub fn find_table(data: &[u8], tag: &[u8; 4]) -> Option<(usize, usize)> {
    if data.len() < 12 {
        return None;
    }
    let num_tables = u16::from_be_bytes([data[4], data[5]]) as usize;
    for i in 0..num_tables {
        let rec = 12 + i * 16;
        if rec + 16 > data.len() {
            return None;
        }
        if &data[rec..rec + 4] == tag {
            let offset = u32::from_be_bytes([
                data[rec + 8],
                data[rec + 9],
                data[rec + 10],
                data[rec + 11],
            ]) as usize;
            let length = u32::from_be_bytes([
                data[rec + 12],
                data[rec + 13],
                data[rec + 14],
                data[rec + 15],
            ]) as usize;
            return Some((offset, length));
        }
    }
    None
}

/// Parse the cmap table from a TTF/OTF font, returning Unicode codepoint → GID mapping.
pub fn parse_cmap_table(font_data: &[u8]) -> HashMap<u32, u16> {
    let mut result = HashMap::new();
    if font_data.len() < 12 {
        return result;
    }

    let num_tables = u16::from_be_bytes([font_data[4], font_data[5]]) as usize;

    // Find the cmap table
    let mut cmap_offset = 0usize;
    let mut cmap_length = 0usize;
    for i in 0..num_tables {
        let rec = 12 + i * 16;
        if rec + 16 > font_data.len() {
            break;
        }
        if &font_data[rec..rec + 4] == b"cmap" {
            cmap_offset = u32::from_be_bytes([
                font_data[rec + 8],
                font_data[rec + 9],
                font_data[rec + 10],
                font_data[rec + 11],
            ]) as usize;
            cmap_length = u32::from_be_bytes([
                font_data[rec + 12],
                font_data[rec + 13],
                font_data[rec + 14],
                font_data[rec + 15],
            ]) as usize;
            break;
        }
    }
    if cmap_offset == 0 || cmap_offset + 4 > font_data.len() {
        return result;
    }

    let cmap = &font_data[cmap_offset..font_data.len().min(cmap_offset + cmap_length)];
    if cmap.len() < 4 {
        return result;
    }

    let num_subtables = u16::from_be_bytes([cmap[2], cmap[3]]) as usize;

    // Prefer: Platform 3 Encoding 10 (Windows UCS-4, Format 12) > Platform 3 Encoding 1 (Windows BMP, Format 4)
    // > Platform 0 (Unicode)
    let mut best_offset = 0usize;
    let mut best_priority = 0u8;

    for i in 0..num_subtables {
        let rec = 4 + i * 8;
        if rec + 8 > cmap.len() {
            break;
        }
        let platform = u16::from_be_bytes([cmap[rec], cmap[rec + 1]]);
        let encoding = u16::from_be_bytes([cmap[rec + 2], cmap[rec + 3]]);
        let offset = u32::from_be_bytes([
            cmap[rec + 4],
            cmap[rec + 5],
            cmap[rec + 6],
            cmap[rec + 7],
        ]) as usize;

        let priority = match (platform, encoding) {
            (3, 10) => 5, // Windows UCS-4 (best, supports all Unicode)
            (0, 4) => 4,  // Unicode full
            (3, 1) => 3,  // Windows BMP
            (0, 3) => 3,  // Unicode BMP
            (0, _) => 2,  // Any Unicode platform
            // Windows symbol encoding. Wingdings, Symbol and their relatives
            // carry only this one, mapping the F000..F0FF private-use block —
            // which is exactly what a .docx asks for when it uses them. Ranked
            // last so a font that has both is still read as Unicode, but
            // reachable, because otherwise these families have no cmap at all
            // and their glyphs never get embedded.
            (3, 0) => 1,
            _ => 0,
        };

        if priority > best_priority {
            best_priority = priority;
            best_offset = offset;
        }
    }

    if best_offset == 0 || best_offset + 2 > cmap.len() {
        return result;
    }

    let subtable = &cmap[best_offset..];
    if subtable.len() < 2 {
        return result;
    }
    let format = u16::from_be_bytes([subtable[0], subtable[1]]);

    match format {
        4 => parse_cmap_format4(subtable, &mut result),
        12 => parse_cmap_format12(subtable, &mut result),
        _ => {}
    }

    result
}

/// Parse cmap subtable format 4 (BMP).
pub fn parse_cmap_format4(data: &[u8], result: &mut HashMap<u32, u16>) {
    if data.len() < 14 {
        return;
    }
    let seg_count = u16::from_be_bytes([data[6], data[7]]) as usize / 2;
    let header_size = 14;

    if data.len() < header_size + seg_count * 8 {
        return;
    }

    let end_codes = &data[header_size..];
    let start_codes = &data[header_size + seg_count * 2 + 2..]; // +2 for reservedPad
    let id_deltas = &data[header_size + seg_count * 4 + 2..];
    let id_range_offsets_start = header_size + seg_count * 6 + 2;
    let id_range_offsets = &data[id_range_offsets_start..];

    for seg in 0..seg_count {
        let off = seg * 2;
        if off + 2 > end_codes.len() || off + 2 > start_codes.len() {
            break;
        }
        let end_code = u16::from_be_bytes([end_codes[off], end_codes[off + 1]]);
        let start_code = u16::from_be_bytes([start_codes[off], start_codes[off + 1]]);
        if off + 2 > id_deltas.len() || off + 2 > id_range_offsets.len() {
            break;
        }
        let id_delta = i16::from_be_bytes([id_deltas[off], id_deltas[off + 1]]);
        let id_range_offset =
            u16::from_be_bytes([id_range_offsets[off], id_range_offsets[off + 1]]);

        if start_code == 0xFFFF {
            break;
        }

        for code in start_code..=end_code {
            let gid = if id_range_offset == 0 {
                (code as i32 + id_delta as i32) as u16
            } else {
                // idRangeOffset points into the glyphIdArray
                let glyph_idx_offset = id_range_offsets_start
                    + off
                    + id_range_offset as usize
                    + (code - start_code) as usize * 2;
                if glyph_idx_offset + 2 <= data.len() {
                    let glyph_id = u16::from_be_bytes([
                        data[glyph_idx_offset],
                        data[glyph_idx_offset + 1],
                    ]);
                    if glyph_id == 0 {
                        0
                    } else {
                        (glyph_id as i32 + id_delta as i32) as u16
                    }
                } else {
                    0
                }
            };
            if gid != 0 {
                result.insert(code as u32, gid);
            }
        }
    }
}

/// Parse cmap subtable format 12 (full Unicode).
pub fn parse_cmap_format12(data: &[u8], result: &mut HashMap<u32, u16>) {
    if data.len() < 16 {
        return;
    }
    let num_groups =
        u32::from_be_bytes([data[12], data[13], data[14], data[15]]) as usize;

    for i in 0..num_groups {
        let off = 16 + i * 12;
        if off + 12 > data.len() {
            break;
        }
        let start_code =
            u32::from_be_bytes([data[off], data[off + 1], data[off + 2], data[off + 3]]);
        let end_code = u32::from_be_bytes([
            data[off + 4],
            data[off + 5],
            data[off + 6],
            data[off + 7],
        ]);
        let start_gid = u32::from_be_bytes([
            data[off + 8],
            data[off + 9],
            data[off + 10],
            data[off + 11],
        ]);

        for code in start_code..=end_code {
            let gid = start_gid + (code - start_code);
            if gid != 0 && gid <= 0xFFFF {
                result.insert(code, gid as u16);
            }
        }
    }
}

/// Check whether a TTF/OTF font contains a CFF table.
pub fn has_cff_table(data: &[u8]) -> bool {
    if data.len() < 12 {
        return false;
    }
    let num_tables = u16::from_be_bytes([data[4], data[5]]) as usize;
    for i in 0..num_tables {
        let off = 12 + i * 16;
        if off + 4 > data.len() {
            return false;
        }
        if &data[off..off + 4] == b"CFF " {
            return true;
        }
    }
    false
}

/// Extract raw CFF data from an OTF font file.
pub fn extract_cff_from_otf(data: &[u8]) -> Option<Vec<u8>> {
    if data.len() < 12 {
        return None;
    }
    let num_tables = u16::from_be_bytes([data[4], data[5]]) as usize;
    for i in 0..num_tables {
        let rec_off = 12 + i * 16;
        if rec_off + 16 > data.len() {
            return None;
        }
        if &data[rec_off..rec_off + 4] == b"CFF " {
            let offset = u32::from_be_bytes([
                data[rec_off + 8],
                data[rec_off + 9],
                data[rec_off + 10],
                data[rec_off + 11],
            ]) as usize;
            let length = u32::from_be_bytes([
                data[rec_off + 12],
                data[rec_off + 13],
                data[rec_off + 14],
                data[rec_off + 15],
            ]) as usize;
            if offset + length <= data.len() {
                return Some(data[offset..offset + length].to_vec());
            }
        }
    }
    None
}

/// Parse hhea + hmtx tables to get per-GID advance widths in 1/1000 em units.
pub fn parse_ttf_widths(font_data: &[u8]) -> HashMap<u16, u16> {
    let mut result = HashMap::new();

    // 1. Parse 'head' table -> unitsPerEm (u16 at offset 18 within the table)
    let (head_off, head_len) = match find_table(font_data, b"head") {
        Some(v) => v,
        None => return result,
    };
    if head_off + head_len > font_data.len() || head_len < 20 {
        return result;
    }
    let head = &font_data[head_off..head_off + head_len];
    let units_per_em = u16::from_be_bytes([head[18], head[19]]) as u32;
    if units_per_em == 0 {
        return result;
    }

    // 2. Parse 'hhea' table -> numOfLongHorMetrics (u16 at offset 34 within the table)
    let (hhea_off, hhea_len) = match find_table(font_data, b"hhea") {
        Some(v) => v,
        None => return result,
    };
    if hhea_off + hhea_len > font_data.len() || hhea_len < 36 {
        return result;
    }
    let hhea = &font_data[hhea_off..hhea_off + hhea_len];
    let num_long_hor_metrics = u16::from_be_bytes([hhea[34], hhea[35]]) as usize;

    // 3. Parse 'hmtx' table -> read advanceWidth for each long horizontal metric
    let (hmtx_off, hmtx_len) = match find_table(font_data, b"hmtx") {
        Some(v) => v,
        None => return result,
    };
    if hmtx_off + hmtx_len > font_data.len() {
        return result;
    }
    let hmtx = &font_data[hmtx_off..hmtx_off + hmtx_len];

    // Each longHorMetric is 4 bytes: advanceWidth(u16) + lsb(i16)
    for gid in 0..num_long_hor_metrics {
        let off = gid * 4;
        if off + 2 > hmtx.len() {
            break;
        }
        let advance = u16::from_be_bytes([hmtx[off], hmtx[off + 1]]) as u32;
        // Round rather than truncate: at 2048 units/em the floor costs up to
        // one 1/1000-em unit on every glyph, all in the same direction, so the
        // error accumulates along a line instead of cancelling.
        let width_1000 = ((advance * 1000 + units_per_em / 2) / units_per_em) as u16;
        result.insert(gid as u16, width_1000);
    }

    // Remaining GIDs (if any) share the last advanceWidth
    if num_long_hor_metrics > 0 {
        let last_off = (num_long_hor_metrics - 1) * 4;
        if last_off + 2 <= hmtx.len() {
            let last_advance = u16::from_be_bytes([hmtx[last_off], hmtx[last_off + 1]]) as u32;
            let last_width_1000 =
                ((last_advance * 1000 + units_per_em / 2) / units_per_em) as u16;

            // leftSideBearing entries follow: 2 bytes each
            let remaining_start = num_long_hor_metrics * 4;
            let remaining_count = (hmtx.len().saturating_sub(remaining_start)) / 2;
            for i in 0..remaining_count {
                let gid = (num_long_hor_metrics + i) as u16;
                result.insert(gid, last_width_1000);
            }
        }
    }

    result
}

/// Parse the name table for the PostScript name (nameID = 6).
pub fn parse_ps_name(font_data: &[u8]) -> Option<String> {
    let (name_off, name_len) = find_table(font_data, b"name")?;
    if name_off + name_len > font_data.len() || name_len < 6 {
        return None;
    }
    let name_table = &font_data[name_off..name_off + name_len];
    let count = u16::from_be_bytes([name_table[2], name_table[3]]) as usize;
    let string_offset = u16::from_be_bytes([name_table[4], name_table[5]]) as usize;

    // First pass: look for platformID=3 (Windows), nameID=6
    // Second pass: look for platformID=1 (Mac), nameID=6
    for target_platform in &[3u16, 1u16] {
        for i in 0..count {
            let rec = 6 + i * 12;
            if rec + 12 > name_table.len() {
                break;
            }
            let platform_id = u16::from_be_bytes([name_table[rec], name_table[rec + 1]]);
            let _encoding_id = u16::from_be_bytes([name_table[rec + 2], name_table[rec + 3]]);
            let _language_id = u16::from_be_bytes([name_table[rec + 4], name_table[rec + 5]]);
            let name_id = u16::from_be_bytes([name_table[rec + 6], name_table[rec + 7]]);
            let length = u16::from_be_bytes([name_table[rec + 8], name_table[rec + 9]]) as usize;
            let offset = u16::from_be_bytes([name_table[rec + 10], name_table[rec + 11]]) as usize;

            if platform_id != *target_platform || name_id != 6 {
                continue;
            }

            let str_start = string_offset + offset;
            if str_start + length > name_table.len() {
                continue;
            }
            let str_data = &name_table[str_start..str_start + length];

            if platform_id == 3 {
                // Windows: UTF-16BE
                let chars: Vec<u16> = str_data
                    .chunks_exact(2)
                    .map(|c| u16::from_be_bytes([c[0], c[1]]))
                    .collect();
                return Some(String::from_utf16_lossy(&chars));
            } else {
                // Mac: ASCII/Latin-1
                return Some(str_data.iter().map(|&b| b as char).collect());
            }
        }
    }
    None
}

/// Extract the single-font data from a TTC (TrueType Collection) file.
/// If the data is not a TTC, returns it as-is.
/// Lift one member of a TrueType Collection out into a standalone font.
///
/// A TTC stores each member's table directory at its own offset, but the
/// offsets *inside* those records are absolute from the start of the file.
/// Slicing the collection at the directory and treating the result as a font
/// therefore reads every table from the wrong place — which is why anything
/// that went through this path (Cambria, MS Gothic, MS Mincho — all shipped as
/// .ttc on Windows) could only be handled by an external subsetter that knew
/// about collections. Rebuilding the member as its own sfnt keeps every later
/// reader honest, because the offsets it sees are its own.
pub fn extract_ttc_face(data: &[u8], face: u32) -> Option<Vec<u8>> {
    if data.len() < 16 || &data[0..4] != b"ttcf" {
        return None;
    }
    let read_u32 = |at: usize| -> Option<u32> {
        Some(u32::from_be_bytes([
            *data.get(at)?,
            *data.get(at + 1)?,
            *data.get(at + 2)?,
            *data.get(at + 3)?,
        ]))
    };
    let num_fonts = read_u32(8)?;
    if face >= num_fonts {
        return None;
    }
    let dir = read_u32(12 + 4 * face as usize)? as usize;
    let sfnt_version = data.get(dir..dir + 4)?.to_vec();
    let num_tables = u16::from_be_bytes([*data.get(dir + 4)?, *data.get(dir + 5)?]) as usize;

    // Collect (tag, bytes) for every table this member names.
    let mut tables: Vec<([u8; 4], &[u8])> = Vec::with_capacity(num_tables);
    for i in 0..num_tables {
        let rec = dir + 12 + i * 16;
        let mut tag = [0u8; 4];
        tag.copy_from_slice(data.get(rec..rec + 4)?);
        let offset = read_u32(rec + 8)? as usize;
        let length = read_u32(rec + 12)? as usize;
        let end = offset.checked_add(length)?;
        tables.push((tag, data.get(offset..end.min(data.len()))?));
    }
    tables.sort_by_key(|(tag, _)| *tag);

    // sfnt layout: header, then one 16-byte record per table, then the table
    // data itself, each padded to a 4-byte boundary.
    let header_len = 12 + tables.len() * 16;
    let mut records = Vec::with_capacity(tables.len() * 16);
    let mut body: Vec<u8> = Vec::new();
    for (tag, bytes) in &tables {
        let offset = header_len + body.len();
        records.extend_from_slice(tag);
        records.extend_from_slice(&0u32.to_be_bytes()); // checksum: readers we
                                                        // feed do not verify it
        records.extend_from_slice(&(offset as u32).to_be_bytes());
        records.extend_from_slice(&(bytes.len() as u32).to_be_bytes());
        body.extend_from_slice(bytes);
        while body.len() % 4 != 0 {
            body.push(0);
        }
    }

    let count = tables.len() as u16;
    let entry_selector = (15u16.saturating_sub(count.leading_zeros() as u16)).min(15);
    let search_range = (1u16 << entry_selector).saturating_mul(16);
    let mut out = Vec::with_capacity(header_len + body.len());
    out.extend_from_slice(&sfnt_version);
    out.extend_from_slice(&count.to_be_bytes());
    out.extend_from_slice(&search_range.to_be_bytes());
    out.extend_from_slice(&entry_selector.to_be_bytes());
    out.extend_from_slice(&(count.saturating_mul(16).saturating_sub(search_range)).to_be_bytes());
    out.extend_from_slice(&records);
    out.extend_from_slice(&body);
    Some(out)
}

/// Build an `EmbeddedFont` from raw TTF/TTC/OTF bytes (no subsetting).
pub fn embedded_font_from_ttf(font_data: &[u8]) -> EmbeddedFont {
    embedded_font_from_face(font_data, 0)
}

/// Build an `EmbeddedFont` from one face of a font file. `face` selects the
/// member of a TrueType Collection and is ignored for a plain TTF/OTF.
pub fn embedded_font_from_face(font_data: &[u8], face: u32) -> EmbeddedFont {
    // A collection member is lifted out as its own sfnt first: the tables must
    // be read, and embedded, with offsets that belong to the font we name.
    let lifted = extract_ttc_face(font_data, face);
    let otf_data: &[u8] = lifted.as_deref().unwrap_or(font_data);

    let unicode_to_gid = parse_cmap_table(otf_data);
    let cid_widths = parse_ttf_widths(otf_data);
    let ps_name = parse_ps_name(otf_data);

    let is_cff = otf_data.starts_with(b"OTTO") || has_cff_table(otf_data);

    let (data, format) = if is_cff {
        if let Some(cff) = extract_cff_from_otf(otf_data) {
            (cff, FontFormat::OpenTypeCff)
        } else {
            (otf_data.to_vec(), FontFormat::OpenTypeCff)
        }
    } else {
        (otf_data.to_vec(), FontFormat::TrueType)
    };

    EmbeddedFont {
        data,
        format,
        unicode_to_gid,
        cid_widths,
        ps_name,
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    /// Build a two-member collection whose tables carry recognisable bytes, so
    /// a lifted face can be checked against what it was supposed to contain.
    fn synthetic_ttc() -> Vec<u8> {
        fn dir(tables: &[([u8; 4], u32, u32)]) -> Vec<u8> {
            let mut out = Vec::new();
            out.extend_from_slice(&0x0001_0000u32.to_be_bytes()); // sfntVersion
            out.extend_from_slice(&(tables.len() as u16).to_be_bytes());
            out.extend_from_slice(&0u16.to_be_bytes()); // searchRange
            out.extend_from_slice(&0u16.to_be_bytes()); // entrySelector
            out.extend_from_slice(&0u16.to_be_bytes()); // rangeShift
            for (tag, offset, length) in tables {
                out.extend_from_slice(tag);
                out.extend_from_slice(&0u32.to_be_bytes());
                out.extend_from_slice(&offset.to_be_bytes());
                out.extend_from_slice(&length.to_be_bytes());
            }
            out
        }

        // Header: tag(4) version(4) numFonts(4) offsets(4*2) = 20 bytes.
        let header_len = 20u32;
        let dir_len = 12 + 16 * 2; // two tables each
        let dir0_at = header_len;
        let dir1_at = dir0_at + dir_len as u32;
        let data_at = dir1_at + dir_len as u32;

        let face0_a = b"FACE-ZERO-TABLE-A".to_vec();
        let face0_b = b"FACE-ZERO-TABLE-B".to_vec();
        let face1_a = b"FACE-ONE-TABLE-A".to_vec();

        let a0 = data_at;
        let b0 = a0 + face0_a.len() as u32;
        let a1 = b0 + face0_b.len() as u32;

        let mut out = Vec::new();
        out.extend_from_slice(b"ttcf");
        out.extend_from_slice(&0x0001_0000u32.to_be_bytes());
        out.extend_from_slice(&2u32.to_be_bytes());
        out.extend_from_slice(&dir0_at.to_be_bytes());
        out.extend_from_slice(&dir1_at.to_be_bytes());
        out.extend_from_slice(&dir(&[
            (*b"AAAA", a0, face0_a.len() as u32),
            (*b"BBBB", b0, face0_b.len() as u32),
        ]));
        out.extend_from_slice(&dir(&[
            (*b"AAAA", a1, face1_a.len() as u32),
            (*b"BBBB", b0, face0_b.len() as u32),
        ]));
        out.extend_from_slice(&face0_a);
        out.extend_from_slice(&face0_b);
        out.extend_from_slice(&face1_a);
        out
    }

    #[test]
    fn extract_ttc_face_lifts_the_named_member() {
        let ttc = synthetic_ttc();

        let face0 = extract_ttc_face(&ttc, 0).expect("face 0");
        let (off, len) = find_table(&face0, b"AAAA").expect("AAAA in face 0");
        assert_eq!(&face0[off..off + len], b"FACE-ZERO-TABLE-A");

        let face1 = extract_ttc_face(&ttc, 1).expect("face 1");
        let (off, len) = find_table(&face1, b"AAAA").expect("AAAA in face 1");
        assert_eq!(
            &face1[off..off + len],
            b"FACE-ONE-TABLE-A",
            "face 1 must not come back holding face 0's tables"
        );

        // Shared tables are copied into each lifted face, not referenced.
        let (off, len) = find_table(&face1, b"BBBB").expect("BBBB in face 1");
        assert_eq!(&face1[off..off + len], b"FACE-ZERO-TABLE-B");

        assert!(extract_ttc_face(&ttc, 2).is_none(), "no third member exists");
        assert!(
            extract_ttc_face(b"not a collection at all", 0).is_none(),
            "a plain font is not a collection"
        );
    }

    /// The offsets a lifted face carries must be its own. Reading the
    /// collection sliced at the directory — what the code used to do — makes
    /// every table land in the wrong place.
    #[test]
    fn lifted_face_offsets_are_self_relative() {
        let ttc = synthetic_ttc();
        let face0 = extract_ttc_face(&ttc, 0).expect("face 0");
        let (off, _) = find_table(&face0, b"AAAA").expect("AAAA");
        assert!(off >= 12 + 2 * 16, "table data must follow the directory");
        assert!(off < face0.len(), "offset must address the lifted font");
    }
}
