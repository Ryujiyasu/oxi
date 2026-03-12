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
            (3, 10) => 4, // Windows UCS-4 (best, supports all Unicode)
            (0, 4) => 3,  // Unicode full
            (3, 1) => 2,  // Windows BMP
            (0, 3) => 2,  // Unicode BMP
            (0, _) => 1,  // Any Unicode platform
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
        let width_1000 = (advance * 1000 / units_per_em) as u16;
        result.insert(gid as u16, width_1000);
    }

    // Remaining GIDs (if any) share the last advanceWidth
    if num_long_hor_metrics > 0 {
        let last_off = (num_long_hor_metrics - 1) * 4;
        if last_off + 2 <= hmtx.len() {
            let last_advance = u16::from_be_bytes([hmtx[last_off], hmtx[last_off + 1]]) as u32;
            let last_width_1000 = (last_advance * 1000 / units_per_em) as u16;

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
fn resolve_ttc(font_data: &[u8]) -> &[u8] {
    if font_data.len() >= 16 && &font_data[0..4] == b"ttcf" {
        // TTC header: tag(4) + version(4) + numFonts(4) + offsets(4*numFonts)
        let _num_fonts =
            u32::from_be_bytes([font_data[8], font_data[9], font_data[10], font_data[11]])
                as usize;
        let first_offset =
            u32::from_be_bytes([font_data[12], font_data[13], font_data[14], font_data[15]])
                as usize;
        if first_offset < font_data.len() {
            return &font_data[first_offset..];
        }
    }
    font_data
}

/// Build an `EmbeddedFont` from raw TTF/TTC/OTF bytes (no subsetting).
pub fn embedded_font_from_ttf(font_data: &[u8]) -> EmbeddedFont {
    let otf_data = resolve_ttc(font_data);

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
        // For TTC, we need the whole file (offset-based tables reference the original data).
        // For plain TTF, otf_data == font_data so this is fine either way.
        (font_data.to_vec(), FontFormat::TrueType)
    };

    EmbeddedFont {
        data,
        format,
        unicode_to_gid,
        cid_widths,
        ps_name,
    }
}
