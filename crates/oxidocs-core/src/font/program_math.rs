// This Source Code Form is subject to the terms of the Mozilla Public
// License, v. 2.0. If a copy of the MPL was not distributed with this
// file, You can obtain one at https://mozilla.org/MPL/2.0/.

//! Read math geometry from the same caller-owned programme used for painting.
//! No font names, fixed glyph indices, or generated metric files live here.

use std::collections::HashMap;
use std::sync::Arc;
use rustybuzz::ttf_parser::{Face, GlyphId};
use super::math_constants::{MathConstants, MathTable};
use super::math_script_glyphs::ScriptGlyph;
use super::math_stretch::{Construction, Direction, Glyph, Part, StretchTable, Variant};

pub(crate) struct ProgramMath {
    pub constants: Arc<MathTable>,
    pub stretch: Arc<StretchTable>,
}

fn stretch_glyph(face: &Face<'_>, bytes: &[u8], face_index: u32, gid: GlyphId) -> Option<Glyph> {
    if gid.0 >= face.number_of_glyphs() { return None; }
    let bounds = super::static_glyph_bounds(bytes, face_index, gid.0).or_else(||
        face.glyph_bounding_box(gid).map(|r| [r.x_min as i32, r.y_min as i32, r.x_max as i32, r.y_max as i32]))
        .unwrap_or([0; 4]);
    let info = face.tables().math.and_then(|m| m.glyph_info);
    Some(Glyph {
        gid: gid.0, advance_width: face.glyph_hor_advance(gid)? as u32, bounds,
        italic_correction: info.and_then(|i| i.italic_corrections)
            .and_then(|c| c.get(gid)).map(|v| v.value as i32),
        extended_shape: info.and_then(|i| i.extended_shapes)
            .is_some_and(|coverage| coverage.get(gid).is_some()),
    })
}

impl ProgramMath {
    pub fn from_face(face: &Face<'_>, bytes: &[u8], face_index: u32) -> Option<Self> {
        let math = face.tables().math?;
        let constants = math.constants?;
        let variants = math.variants?;
        let mut codepoints = HashMap::<u16, u32>::new();
        for subtable in face.tables().cmap?.subtables {
            if !subtable.is_unicode() { continue; }
            subtable.codepoints(|cp| {
                if let Some(gid) = char::from_u32(cp).and_then(|c| face.glyph_index(c)) {
                    codepoints.entry(gid.0).and_modify(|old| *old = (*old).min(cp)).or_insert(cp);
                }
            });
        }
        let mut constructions = Vec::new();
        for (direction, available) in [(Direction::Vert, variants.vertical_constructions),
                                        (Direction::Horiz, variants.horizontal_constructions)] {
            for index in 0..face.number_of_glyphs() {
                let gid = GlyphId(index);
                let Some(construction) = available.get(gid) else { continue; };
                let ready = construction.variants.into_iter().map(|v| Some(Variant {
                    glyph: stretch_glyph(face, bytes, face_index, v.variant_glyph)?,
                    advance_measurement: v.advance_measurement as u32,
                })).collect::<Option<Vec<_>>>()?;
                let assembly = match construction.assembly {
                    Some(a) => a.parts.into_iter().map(|p| Some(Part {
                        glyph: stretch_glyph(face, bytes, face_index, p.glyph_id)?,
                        start: p.start_connector_length as u32, end: p.end_connector_length as u32,
                        full_advance: p.full_advance as u32, flags: p.part_flags.0,
                    })).collect::<Option<Vec<_>>>()?,
                    None => Vec::new(),
                };
                constructions.push(Construction {
                    direction, base: stretch_glyph(face, bytes, face_index, gid)?, codepoint: codepoints.get(&index).copied(),
                    variants: ready, assembly,
                    italic_correction: construction.assembly
                        .map(|a| a.italics_correction.value as i32).unwrap_or(0),
                });
            }
        }
        Some(Self {
            constants: Arc::new(MathTable { upm: face.units_per_em() as u32,
                constants: MathConstants {
                    SuperscriptShiftUp: constants.superscript_shift_up().value as i32,
                    SuperscriptShiftUpCramped: constants.superscript_shift_up_cramped().value as i32,
                    SuperscriptBottomMin: constants.superscript_bottom_min().value as i32,
                    SuperscriptBaselineDropMax: constants.superscript_baseline_drop_max().value as i32,
                    SuperscriptBottomMaxWithSubscript: constants.superscript_bottom_max_with_subscript().value as i32,
                    SubscriptShiftDown: constants.subscript_shift_down().value as i32,
                    SubscriptTopMax: constants.subscript_top_max().value as i32,
                    SubscriptBaselineDropMin: constants.subscript_baseline_drop_min().value as i32,
                    SubSuperscriptGapMin: constants.sub_superscript_gap_min().value as i32,
                    SpaceAfterScript: constants.space_after_script().value as i32,
                    FractionNumeratorShiftUp: constants.fraction_numerator_shift_up().value as i32,
                    FractionDenominatorShiftDown: constants.fraction_denominator_shift_down().value as i32,
                    FractionNumeratorGapMin: constants.fraction_numerator_gap_min().value as i32,
                    FractionDenominatorGapMin: constants.fraction_denominator_gap_min().value as i32,
                    FractionRuleThickness: constants.fraction_rule_thickness().value as i32,
                    FractionNumeratorDisplayStyleShiftUp: constants.fraction_numerator_display_style_shift_up().value as i32,
                    FractionDenominatorDisplayStyleShiftDown: constants.fraction_denominator_display_style_shift_down().value as i32,
                    FractionNumDisplayStyleGapMin: constants.fraction_num_display_style_gap_min().value as i32,
                    FractionDenomDisplayStyleGapMin: constants.fraction_denom_display_style_gap_min().value as i32,
                    SkewedFractionHorizontalGap: constants.skewed_fraction_horizontal_gap().value as i32,
                    SkewedFractionVerticalGap: constants.skewed_fraction_vertical_gap().value as i32,
                    RadicalVerticalGap: constants.radical_vertical_gap().value as i32,
                    RadicalDisplayStyleVerticalGap: constants.radical_display_style_vertical_gap().value as i32,
                    RadicalRuleThickness: constants.radical_rule_thickness().value as i32,
                    RadicalExtraAscender: constants.radical_extra_ascender().value as i32,
                    RadicalKernBeforeDegree: constants.radical_kern_before_degree().value as i32,
                    RadicalKernAfterDegree: constants.radical_kern_after_degree().value as i32,
                    RadicalDegreeBottomRaisePercent: constants.radical_degree_bottom_raise_percent() as i32,
                    UpperLimitGapMin: constants.upper_limit_gap_min().value as i32,
                    UpperLimitBaselineRiseMin: constants.upper_limit_baseline_rise_min().value as i32,
                    LowerLimitGapMin: constants.lower_limit_gap_min().value as i32,
                    LowerLimitBaselineDropMin: constants.lower_limit_baseline_drop_min().value as i32,
                    StackTopShiftUp: constants.stack_top_shift_up().value as i32,
                    StackTopDisplayStyleShiftUp: constants.stack_top_display_style_shift_up().value as i32,
                    StackBottomShiftDown: constants.stack_bottom_shift_down().value as i32,
                    StackBottomDisplayStyleShiftDown: constants.stack_bottom_display_style_shift_down().value as i32,
                    StackGapMin: constants.stack_gap_min().value as i32,
                    StackDisplayStyleGapMin: constants.stack_display_style_gap_min().value as i32,
                    StretchStackTopShiftUp: constants.stretch_stack_top_shift_up().value as i32,
                    StretchStackBottomShiftDown: constants.stretch_stack_bottom_shift_down().value as i32,
                    StretchStackGapAboveMin: constants.stretch_stack_gap_above_min().value as i32,
                    StretchStackGapBelowMin: constants.stretch_stack_gap_below_min().value as i32,
                    OverbarVerticalGap: constants.overbar_vertical_gap().value as i32,
                    OverbarRuleThickness: constants.overbar_rule_thickness().value as i32,
                    OverbarExtraAscender: constants.overbar_extra_ascender().value as i32,
                    UnderbarVerticalGap: constants.underbar_vertical_gap().value as i32,
                    UnderbarRuleThickness: constants.underbar_rule_thickness().value as i32,
                    UnderbarExtraDescender: constants.underbar_extra_descender().value as i32,
                    AxisHeight: constants.axis_height().value as i32,
                    AccentBaseHeight: constants.accent_base_height().value as i32,
                    FlattenedAccentBaseHeight: constants.flattened_accent_base_height().value as i32,
                    MathLeading: constants.math_leading().value as i32,
                    ScriptPercentScaleDown: constants.script_percent_scale_down() as i32,
                    ScriptScriptPercentScaleDown: constants.script_script_percent_scale_down() as i32,
                    DelimitedSubFormulaMinHeight: constants.delimited_sub_formula_min_height() as i32,
                    DisplayOperatorMinHeight: constants.display_operator_min_height() as i32,
                } }),
            stretch: Arc::new(StretchTable { upm: face.units_per_em() as u32,
                min_connector_overlap: variants.min_connector_overlap as u32, constructions }),
        })
    }
}

/// Use the same cmap selection as the registered programme's advance metrics.
/// This includes Windows Symbol mappings and their conventional low-byte aliases.
fn nominal_glyph(bytes: &[u8], face_index: u32, c: char) -> Option<GlyphId> {
    use skrifa::MetadataProvider;
    let font = match skrifa::raw::FileRef::new(bytes).ok()? {
        skrifa::raw::FileRef::Font(font) if face_index == 0 => font,
        skrifa::raw::FileRef::Collection(collection) => collection.get(face_index).ok()?,
        _ => return None,
    };
    let gid = font.charmap().map(c)?.to_u32();
    Some(GlyphId(u16::try_from(gid).ok()?))
}

/// Shape script alternates using the programme's GSUB instead of a saved ID.
/// A level-zero character is selected directly through this face's cmap.
pub(crate) fn glyph(bytes: &[u8], face_index: u32, c: char, level: u8) -> Option<ScriptGlyph> {
    let face = rustybuzz::Face::from_slice(bytes, face_index)?;
    let gid = if level == 0 {
        nominal_glyph(bytes, face_index, c)?
    } else {
        let mut buffer = rustybuzz::UnicodeBuffer::new();
        buffer.push_str(&c.to_string());
        buffer.guess_segment_properties();
        if let Some(script) = rustybuzz::Script::from_iso15924_tag(rustybuzz::ttf_parser::Tag::from_bytes(b"Zmth")) {
            buffer.set_script(script);
        }
        let feature = rustybuzz::Feature::new(rustybuzz::ttf_parser::Tag::from_bytes(b"ssty"), u32::from(level.min(2)), ..);
        let shaped = rustybuzz::shape(&face, &[feature], buffer);
        let info = shaped.glyph_infos();
        if info.len() != 1 || info[0].glyph_id == 0 || info[0].glyph_id > u16::MAX as u32 { return None; }
        GlyphId(info[0].glyph_id as u16)
    };
    let em = face.units_per_em() as f32;
    let bounds_em = super::static_glyph_bounds(bytes, face_index, gid.0).or_else(||
        face.glyph_bounding_box(gid).map(|r| [i32::from(r.x_min), i32::from(r.y_min),
                                           i32::from(r.x_max), i32::from(r.y_max)]))
        .map(|bounds| bounds.map(|value| value as f32/em)).unwrap_or([0.0; 4]);
    let math = face.tables().math.and_then(|m| m.glyph_info);
    Some(ScriptGlyph { index: gid.0, advance_em: f32::from(face.glyph_hor_advance(gid)?)/em,
        bounds_em, italic_correction_em: math.and_then(|i| i.italic_corrections)
            .and_then(|v| v.get(gid)).map_or(0.0, |v| f32::from(v.value)/em),
        top_accent_attachment_em: math.and_then(|i| i.top_accent_attachments)
            .and_then(|v| v.get(gid)).map(|v| f32::from(v.value)/em),
    })
}

pub(crate) fn corner_kern(bytes: &[u8], face_index: u32, gid: u16,
    corner: super::math_kern::Corner, height: f32, size: f32) -> f32 {
    if !height.is_finite() || !size.is_finite() || size <= 0.0 { return 0.0; }
    let Some(face) = Face::parse(bytes, face_index).ok() else { return 0.0; };
    let Some(info) = face.tables().math.and_then(|m|m.glyph_info)
        .and_then(|i|i.kern_infos).and_then(|i|i.get(GlyphId(gid))) else { return 0.0; };
    let selected = match corner {
        super::math_kern::Corner::TopRight => info.top_right,
        super::math_kern::Corner::TopLeft => info.top_left,
        super::math_kern::Corner::BottomRight => info.bottom_right,
        super::math_kern::Corner::BottomLeft => info.bottom_left,
    };
    let Some(kern) = selected else { return 0.0; };
    let em = f32::from(face.units_per_em());
    let target = height * em / size;
    let index = (0..kern.count()).find(|&i| kern.height(i)
        .is_some_and(|h| f32::from(h.value) > target)).unwrap_or(kern.count());
    kern.kern(index).map_or(0.0, |v| f32::from(v.value) * size / em)
}


#[cfg(test)]
mod nominal_mapping_tests {
    use super::nominal_glyph;

    // Self-authored cmap-only parser fixtures. They contain no font programme
    // outlines, names, or third-party font bytes and never leave test memory.
    fn format4(cp: u16, gid: u16) -> Vec<u8> {
        let values = [4, 32, 0, 4, 4, 1, 0, cp, 0xffff, 0, cp, 0xffff,
                      gid.wrapping_sub(cp), 1, 0, 0];
        values.into_iter().flat_map(u16::to_be_bytes).collect()
    }

    fn format12(groups: &[(u32, u32)]) -> Vec<u8> {
        let mut data = Vec::new();
        data.extend_from_slice(&12u16.to_be_bytes());
        data.extend_from_slice(&0u16.to_be_bytes());
        data.extend_from_slice(&(16u32 + 12 * groups.len() as u32).to_be_bytes());
        data.extend_from_slice(&0u32.to_be_bytes());
        data.extend_from_slice(&(groups.len() as u32).to_be_bytes());
        for &(cp, gid) in groups {
            for value in [cp, cp, gid] { data.extend_from_slice(&value.to_be_bytes()); }
        }
        data
    }

    fn cmap_sfnt(tables: &[(u16, u16, Vec<u8>)]) -> Vec<u8> {
        let mut cmap = Vec::new();
        cmap.extend_from_slice(&0u16.to_be_bytes());
        cmap.extend_from_slice(&(tables.len() as u16).to_be_bytes());
        let mut offset = 4u32 + 8 * tables.len() as u32;
        for (platform, encoding, data) in tables {
            cmap.extend_from_slice(&platform.to_be_bytes());
            cmap.extend_from_slice(&encoding.to_be_bytes());
            cmap.extend_from_slice(&offset.to_be_bytes());
            offset += data.len() as u32;
        }
        for (_, _, data) in tables { cmap.extend_from_slice(data); }
        let mut sfnt = Vec::new();
        sfnt.extend_from_slice(&0x0001_0000u32.to_be_bytes());
        for value in [1u16, 16, 0, 0] { sfnt.extend_from_slice(&value.to_be_bytes()); }
        sfnt.extend_from_slice(b"cmap");
        for value in [0u32, 28, cmap.len() as u32] { sfnt.extend_from_slice(&value.to_be_bytes()); }
        sfnt.extend_from_slice(&cmap);
        sfnt
    }

    #[test]
    fn symbol_mapping_resolves_actual_pua_and_conventional_low_byte() {
        let bytes = cmap_sfnt(&[(3, 0, format4(0xf021, 3))]);
        assert_eq!(nominal_glyph(&bytes, 0, '\u{f021}').unwrap().0, 3);
        assert_eq!(nominal_glyph(&bytes, 0, '!').unwrap().0, 3);
        assert!(nominal_glyph(&bytes, 0, '\u{f022}').is_none());
        assert!(nominal_glyph(&bytes, 0, '\u{1021}').is_none());
    }

    #[test]
    fn symbol_subtable_wins_over_conflicting_unicode_mapping() {
        let bytes = cmap_sfnt(&[(3, 0, format4(0xf021, 3)),
                               (3, 10, format12(&[(0xf021, 8)]))]);
        assert_eq!(nominal_glyph(&bytes, 0, '\u{f021}').unwrap().0, 3);
    }

    #[test]
    fn unicode_full_repertoire_covers_bmp_and_non_bmp() {
        let bytes = cmap_sfnt(&[(3, 1, format4(0x0041, 2)),
                               (3, 10, format12(&[(0x0041, 4), (0x1f600, 7)]))]);
        assert_eq!(nominal_glyph(&bytes, 0, 'A').unwrap().0, 4);
        assert_eq!(nominal_glyph(&bytes, 0, '\u{1f600}').unwrap().0, 7);
        assert!(nominal_glyph(&bytes, 0, 'B').is_none());
        assert!(nominal_glyph(&bytes, 1, 'A').is_none());
        assert!(nominal_glyph(&[], 0, 'A').is_none());
    }

    #[test]
    fn collection_face_index_selects_its_own_character_map() {
        let mut first = cmap_sfnt(&[(3, 1, format4(0x0041, 3))]);
        let mut second = cmap_sfnt(&[(3, 1, format4(0x0041, 9))]);
        let first_offset = 20u32;
        let second_offset = first_offset + first.len() as u32;
        first[20..24].copy_from_slice(&(first_offset + 28).to_be_bytes());
        second[20..24].copy_from_slice(&(second_offset + 28).to_be_bytes());
        let mut ttc = Vec::from(*b"ttcf");
        for value in [0x0001_0000u32, 2, first_offset, second_offset] {
            ttc.extend_from_slice(&value.to_be_bytes());
        }
        ttc.extend_from_slice(&first);
        ttc.extend_from_slice(&second);
        assert_eq!(nominal_glyph(&ttc, 0, 'A').unwrap().0, 3);
        assert_eq!(nominal_glyph(&ttc, 1, 'A').unwrap().0, 9);
        assert!(nominal_glyph(&ttc, 2, 'A').is_none());
    }
}
