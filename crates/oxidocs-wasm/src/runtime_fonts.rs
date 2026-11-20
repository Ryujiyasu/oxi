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
        assert!(stand_in_font_program("Segoe Script", false, false).is_none());
        assert!(stand_in_font_program("Brush Script MT", false, false).is_none());
        assert!(stand_in_font_program("MS Sans Serif", false, false).is_none());
    }
}
