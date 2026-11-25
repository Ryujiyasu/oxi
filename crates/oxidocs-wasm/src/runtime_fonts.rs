// SPDX-License-Identifier: MIT OR Apache-2.0

use std::cell::{Cell, RefCell};
use wasm_bindgen::prelude::*;
use crate::font_program::FontPrograms;

thread_local! {
    static PROGRAMS: RefCell<FontPrograms> = RefCell::new(FontPrograms::default());
    static FONT_REVISION: Cell<u64> = const { Cell::new(0) };
}

/// Register caller-supplied font bytes in memory. The face's own name table
/// and style metadata must agree; no OS paths or family exception lists.
#[wasm_bindgen]
pub fn register_font_program(
    family: &str, bold: bool, italic: bool, bytes: &[u8], face_index: u32,
) -> Result<(), JsError> {
    register_in_memory(family, bold, italic, bytes, face_index).map_err(|e| JsError::new(&e))
}

#[wasm_bindgen]
pub fn clear_font_programs() {
    PROGRAMS.with(|programs| *programs.borrow_mut() = FontPrograms::default());
    oxidocs_core::font::runtime::clear_memory_fonts();
    FONT_REVISION.with(|revision| revision.set(revision.get().wrapping_add(1)));
}

#[wasm_bindgen]
pub fn has_font_program(family: &str, bold: bool, italic: bool) -> bool {
    PROGRAMS.with(|programs| programs.borrow().program(family, bold, italic).is_some())
}

/// Painting commands are returned to the renderer in memory, independently
/// from layout/source dumps. They are never stored as font data tables.
#[wasm_bindgen]
pub fn get_font_glyph_outline(
    family: &str, bold: bool, italic: bool, index: u16, expected_bounds_em: &[f32],
) -> Result<JsValue, JsError> {
    let bounds: [f32; 4] = expected_bounds_em.try_into()
        .map_err(|_| JsError::new("Expected four glyph bounds"))?;
    let painting = PROGRAMS.with(|programs| programs.borrow().selected_glyph(
        family, bold, italic, index, bounds,
    )).map_err(|e| JsError::new(&format!("Selected glyph: {e:?}")))?;
    serde_wasm_bindgen::to_value(&painting).map_err(|e| JsError::new(&e.to_string()))
}

pub(crate) fn resource_name(family: &str, bold: bool, italic: bool) -> String {
    format!("OxiRuntime-{family}-{}-{}", u8::from(bold), u8::from(italic))
}

pub(crate) fn embedded_font(
    family: &str, bold: bool, italic: bool,
) -> Option<oxipdf_core::ir::EmbeddedFont> {
    PROGRAMS.with(|programs| {
        let programs = programs.borrow();
        let (bytes, index) = programs.program(family, bold, italic)?;
        Some(oxipdf_core::font_util::embedded_font_from_face(bytes, index))
    })
}

pub(crate) fn validate_layout_glyphs(
    layout: &oxidocs_core::layout::LayoutResult,
) -> Result<(), String> {
    PROGRAMS.with(|programs| {
        let programs = programs.borrow();
        for element in layout.pages.iter().flat_map(|page| &page.elements) {
            if let Some(glyph) = &element.font_glyph {
                if let oxidocs_core::layout::LayoutContent::Text {
                    font_family, bold, italic, text, ..
                } = &element.content {
                    if text.is_empty() { continue; }
                    let family = font_family.as_deref().ok_or("Selected glyph has no family")?;
                    programs.selected_glyph(family, *bold, *italic, glyph.index, glyph.bounds_em)
                        .map_err(|e| format!("Selected glyph font program: {e:?}"))?;
                }
            }
        }
        Ok(())
    })
}

/// Resolve the requested name/style from the actual program's name tables,
/// including every collection member. No filename or family exception map.
#[wasm_bindgen]
pub fn try_register_font_program_family(family: &str, bold: bool, italic: bool, bytes: &[u8]) -> bool {
    let count = ttf_parser::fonts_in_collection(bytes).unwrap_or(1);
    for index in 0..count {
        if register_in_memory(family, bold, italic, bytes, index).is_ok() { return true; }
    }
    false
}

#[wasm_bindgen]
pub fn register_font_program_family(family: &str, bold: bool, italic: bool, bytes: &[u8]) -> Result<(), JsError> {
    if try_register_font_program_family(family, bold, italic, bytes) {
        Ok(())
    } else {
        Err(JsError::new("Font program has no matching family/style"))
    }
}

/// Complete SFNT bytes for the CSS face; unlike PDF embedding, preserve the
/// OTF wrapper around CFF tables. Returned bytes stay in the client memory.
#[wasm_bindgen]
pub fn get_registered_font_sfnt(family: &str, bold: bool, italic: bool) -> Result<Vec<u8>, JsError> {
    PROGRAMS.with(|programs| {
        let programs = programs.borrow();
        let (bytes, index) = programs.program(family, bold, italic)
            .ok_or_else(|| JsError::new("Font program is not registered"))?;
        Ok(oxipdf_core::font_util::extract_ttc_face(bytes, index).unwrap_or_else(|| bytes.to_vec()))
    })
}

fn register_in_memory(
    family: &str, bold: bool, italic: bool, bytes: &[u8], face_index: u32,
) -> Result<(), String> {
    let unchanged = PROGRAMS.with(|programs| programs.borrow().program(family, bold, italic)
        .is_some_and(|(existing, index)| index == face_index && existing == bytes));
    if unchanged { return Ok(()); }
    let prepared = oxidocs_core::font::runtime::prepare_memory_font(family, bold, italic, bytes, face_index)
        .map_err(str::to_owned)?;
    PROGRAMS.with(|programs| programs.borrow_mut().register(
        family, bold, italic, bytes.to_vec(), face_index,
    )).map_err(|error| format!("Font registration: {error:?}"))?;
    prepared.publish();
    FONT_REVISION.with(|revision| revision.set(revision.get().wrapping_add(1)));
    Ok(())
}

/// A successful registration changes both layout and painting together.
/// Callers use this revision to relayout after asynchronous font loading.
#[wasm_bindgen]
pub fn font_program_revision() -> u64 {
    FONT_REVISION.with(|revision| revision.get())
}

/// The face Oxi ships to draw `family` when the machine does not have it:
/// Carlito for Calibri, Liberation for Arial / Times New Roman / Courier New,
/// Caladea for Cambria, OxiMincho / OxiGothic for a Japanese family. Null
/// for a family with no stand-in.
///
/// It is for painting only. Registering it as `family`'s program would lay
/// the text out with the stand-in's widths instead of the measured ones, so a
/// caller hands it to CSS (`FontFace`) and the layout keeps its own metrics.
/// The Japanese pair has no bold or italic member; the browser slants or
/// emboldens the regular one.
#[wasm_bindgen]
pub fn stand_in_font_program(family: &str, bold: bool, italic: bool) -> Option<Vec<u8>> {
    if let Some(base) = crate::latin_stand_in(family) {
        let key = match (bold, italic) {
            (false, false) => base.to_string(),
            (true, false) => format!("{base}-B"),
            (false, true) => format!("{base}-I"),
            (true, true) => format!("{base}-BI"),
        };
        return oxidocs_fonts::face(&key).map(<[u8]>::to_vec);
    }
    if !east_asian_family(family) { return None; }
    oxidocs_fonts::face(oxidocs_fonts::cjk_stand_in(family).0).map(<[u8]>::to_vec)
}

/// The face Oxi ships to draw what an East Asian `family` it stands in for
/// sets proportionally -- its Latin when that is not half-width, its narrow
/// kana and punctuation when its kana are proportional: OxiGothicP /
/// OxiMinchoP. Null for a family that sets neither (ＭＳ ゴシック, ＭＳ 明朝,
/// BIZ UDゴシック), which the stand-in itself draws right, and for one with no
/// Japanese stand-in.
///
/// `stand_in_font_program` sets ASCII at half an em and kana at a full em, as
/// ＭＳ ゴシック does; given to CSS alone it drew 游ゴシック or メイリオ in a
/// typewriter face with a gap after each word, and ＭＳ Ｐゴシック's half-em
/// 「（」 a full em wide. A caller adds this face under the same family with
/// `unicodeRange: stand_in_proportional_range(family)`, after the main one, so
/// the browser takes those characters from it.
#[wasm_bindgen]
pub fn stand_in_proportional_font_program(family: &str) -> Option<Vec<u8>> {
    stand_in_proportional_range(family)?;
    oxidocs_fonts::face(oxidocs_fonts::cjk_proportional_stand_in(family).0).map(<[u8]>::to_vec)
}

/// The CSS `unicode-range` of the characters `family` takes from
/// `stand_in_proportional_font_program`: those the measured widths say it
/// sets proportionally (`FontMetrics::draws_proportional`). Null when there
/// are none.
#[wasm_bindgen]
pub fn stand_in_proportional_range(family: &str) -> Option<String> {
    if crate::latin_stand_in(family).is_some() || !east_asian_family(family) { return None; }
    let metrics = crate::metrics_registry().get(family);
    let mut ranges: Vec<(u32, u32)> = Vec::new();
    for c in [' '..='~', '\u{3000}'..='\u{30FF}', '\u{FF01}'..='\u{FF60}'].into_iter().flatten() {
        if !oxidocs_fonts::in_cjk_proportional_stand_in(c) || !metrics.draws_proportional(c) { continue; }
        match ranges.last_mut() {
            Some((_, end)) if *end + 1 == c as u32 => *end = c as u32,
            _ => ranges.push((c as u32, c as u32)),
        }
    }
    if ranges.is_empty() { return None; }
    Some(ranges.iter()
        .map(|&(a, b)| if a == b { format!("U+{a:04X}") } else { format!("U+{a:04X}-{b:04X}") })
        .collect::<Vec<_>>()
        .join(", "))
}

/// Where to cut an element of `text` drawn in the CSS stand-in of `family`
/// (main face plus `stand_in_proportional_font_program` over
/// `stand_in_proportional_range`), so each piece sits where the layout's
/// measured widths put it: a JSON array of `[text, x, scale]`, x from the
/// element's left edge and `scale` the horizontal factor to draw it at. The
/// cuts are the PDF export's (`proportional_pieces`): a Latin word scaled to its
/// measured width, each kana or mark the family sets proportionally in its own
/// cell, narrowed and if still too wide centred on it. Null when the family
/// sets nothing proportionally or `text` has nothing it does.
#[wasm_bindgen]
pub fn stand_in_pieces(family: &str, text: &str, width: f64, font_size: f64, spacing: f64) -> Option<String> {
    if crate::latin_stand_in(family).is_some() || !east_asian_family(family) { return None; }
    let face_key = match oxidocs_fonts::cjk_stand_in(family).0 {
        "OxiMincho" => "OxiCJK-Mincho",
        _ => "OxiCJK-Gothic",
    };
    let pieces = crate::proportional_pieces(text, family, face_key, 0.0, width, font_size, spacing)?;
    let out: Vec<(String, f64, f64)> = pieces.into_iter()
        .map(|(piece, _, x, hscale)| (piece, x, hscale / 100.0))
        .collect();
    serde_json::to_string(&out).ok()
}

/// A family named in Japanese, or one of the Latin-spelled Japanese families
/// Word documents carry (MS Gothic, Yu Mincho, Meiryo, BIZ UDPGothic,
/// HGSoeiKakugothicUB, UD Digi Kyokasho…). The Latin "Gothic" faces are not.
fn east_asian_family(family: &str) -> bool {
    if family.chars().any(|c| c as u32 > 0x2E7F) { return true; }
    let lower = family.to_lowercase();
    const LATIN_GOTHICS: [&str; 6] =
        ["century gothic", "franklin gothic", "trade gothic", "news gothic", "letter gothic", "highway gothic"];
    if LATIN_GOTHICS.iter().any(|latin| lower.starts_with(latin)) { return false; }
    ["gothic", "mincho", "meiryo", "kyokasho", "hiragino", "osaka", "kozuka"].iter().any(|word| lower.contains(word))
}

#[cfg(test)]
mod stand_in_tests {
    use super::*;

    #[test]
    fn latin_and_japanese_families_get_their_shipped_face() {
        assert_eq!(stand_in_font_program("Calibri", true, false).as_deref(), oxidocs_fonts::face("Carlito-B"));
        assert_eq!(stand_in_font_program("Times New Roman", false, true).as_deref(), oxidocs_fonts::face("LiberationSerif-I"));
        assert_eq!(stand_in_font_program("ＭＳ 明朝", false, false).as_deref(), Some(oxidocs_fonts::OXI_MINCHO));
        assert_eq!(stand_in_font_program("Yu Gothic", true, false).as_deref(), Some(oxidocs_fonts::OXI_GOTHIC));
        assert_eq!(stand_in_font_program("Meiryo UI", false, false).as_deref(), Some(oxidocs_fonts::OXI_GOTHIC));
        assert_eq!(stand_in_font_program("HGSoeiKakugothicUB", false, false).as_deref(), Some(oxidocs_fonts::OXI_GOTHIC));
        assert_eq!(stand_in_font_program("UD Digi Kyokasho N-R", false, false).as_deref(), Some(oxidocs_fonts::OXI_GOTHIC));
    }

    #[test]
    fn latin_families_without_a_stand_in_get_none() {
        assert!(stand_in_font_program("Century Gothic", false, false).is_none());
        assert_eq!(stand_in_proportional_font_program("游ゴシック").as_deref(), Some(oxidocs_fonts::OXI_GOTHIC_P));
        assert_eq!(stand_in_proportional_font_program("ＭＳ Ｐ明朝").as_deref(), Some(oxidocs_fonts::OXI_MINCHO_P));
        assert_eq!(stand_in_proportional_font_program("メイリオ").as_deref(), Some(oxidocs_fonts::OXI_GOTHIC_P));
        assert!(stand_in_proportional_font_program("ＭＳ ゴシック").is_none());
        assert!(stand_in_proportional_font_program("ＭＳ 明朝").is_none());
        assert!(stand_in_proportional_font_program("Calibri").is_none());
        // 游ゴシック's kana are full-width: its range is the Latin alone.
        assert_eq!(stand_in_proportional_range("游ゴシック").as_deref(), Some("U+0020-007E"));
        let pieces: Vec<(String, f64, f64)> = serde_json::from_str(
            &stand_in_pieces("ＭＳ Ｐゴシック", "（何）", 20.0, 10.0, 0.0).unwrap()).unwrap();
        assert_eq!(pieces.iter().map(|p| p.0.as_str()).collect::<Vec<_>>(), ["（", "何", "）"]);
        assert_eq!(pieces[1].1, 5.0);
        assert!(stand_in_pieces("ＭＳ ゴシック", "（何）", 20.0, 10.0, 0.0).is_none());
        let pgothic = stand_in_proportional_range("ＭＳ Ｐゴシック").unwrap();
        assert!(pgothic.starts_with("U+0020-007E, U+30"), "{pgothic}");
        assert!(pgothic.contains("FF08"), "（ is in: {pgothic}");
        assert!(stand_in_font_program("Segoe Script", false, false).is_none());
        assert!(stand_in_font_program("Brush Script MT", false, false).is_none());
        assert!(stand_in_font_program("MS Sans Serif", false, false).is_none());
    }
}
