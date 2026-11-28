// This Source Code Form is subject to the terms of the Mozilla Public
// License, v. 2.0. If a copy of the MPL was not distributed with this
// file, You can obtain one at https://mozilla.org/MPL/2.0/.

//! Font programs registered at run time: the faces a browser reads from the
//! machine, or a phone app from its bundle or the system. They are kept in
//! memory, laid out with (`oxidocs_core::font::runtime`) and drawn and
//! embedded from (`crate::pdf`). Process-wide, like the layout's own store.

use std::sync::atomic::{AtomicU64, Ordering};
use std::sync::{Mutex, MutexGuard, OnceLock};

use crate::font_program::{FontPrograms, GlyphPainting};

fn programs() -> MutexGuard<'static, FontPrograms> {
    static PROGRAMS: OnceLock<Mutex<FontPrograms>> = OnceLock::new();
    PROGRAMS.get_or_init(|| Mutex::new(FontPrograms::default()))
        .lock()
        .unwrap_or_else(|poisoned| poisoned.into_inner())
}

static REVISION: AtomicU64 = AtomicU64::new(0);

fn bump() {
    REVISION.fetch_add(1, Ordering::Relaxed);
}

/// Register `bytes` (member `face_index` of a collection) as `family`'s
/// program for that style. The face's own name table and style must agree.
pub fn register(family: &str, bold: bool, italic: bool, bytes: &[u8], face_index: u32) -> Result<(), String> {
    let unchanged = programs().program(family, bold, italic)
        .is_some_and(|(existing, index)| index == face_index && existing == bytes);
    if unchanged { return Ok(()); }
    let prepared = oxidocs_core::font::runtime::prepare_memory_font(family, bold, italic, bytes, face_index)
        .map_err(str::to_owned)?;
    programs().register(family, bold, italic, bytes.to_vec(), face_index)
        .map_err(|error| format!("Font registration: {error:?}"))?;
    prepared.publish();
    bump();
    Ok(())
}

/// Register whichever member of `bytes` names itself `family` in that style.
pub fn try_register_family(family: &str, bold: bool, italic: bool, bytes: &[u8]) -> bool {
    let count = ttf_parser::fonts_in_collection(bytes).unwrap_or(1);
    (0..count).any(|index| register(family, bold, italic, bytes, index).is_ok())
}

/// Forget every registered program.
pub fn clear() {
    *programs() = FontPrograms::default();
    oxidocs_core::font::runtime::clear_memory_fonts();
    bump();
}

pub fn has_font_program(family: &str, bold: bool, italic: bool) -> bool {
    programs().program(family, bold, italic).is_some()
}

/// Changes whenever a program is registered or cleared, so a caller knows
/// to lay out again.
pub fn revision() -> u64 {
    REVISION.load(Ordering::Relaxed)
}

/// The PDF resource name of a registered program.
pub fn resource_name(family: &str, bold: bool, italic: bool) -> String {
    format!("OxiRuntime-{family}-{}-{}", u8::from(bold), u8::from(italic))
}

pub fn embedded_font(family: &str, bold: bool, italic: bool) -> Option<oxipdf_core::ir::EmbeddedFont> {
    let programs = programs();
    let (bytes, index) = programs.program(family, bold, italic)?;
    Some(oxipdf_core::font_util::embedded_font_from_face(bytes, index))
}

/// The complete SFNT of a registered program (a collection's member lifted
/// out of it), for a CSS face or a platform font.
pub fn registered_sfnt(family: &str, bold: bool, italic: bool) -> Option<Vec<u8>> {
    let programs = programs();
    let (bytes, index) = programs.program(family, bold, italic)?;
    Some(oxipdf_core::font_util::extract_ttc_face(bytes, index).unwrap_or_else(|| bytes.to_vec()))
}

/// The outline of glyph `index` of a registered program, checked against
/// the bounds the layout measured for it.
pub fn selected_glyph(
    family: &str, bold: bool, italic: bool, index: u16, bounds_em: [f32; 4],
) -> Result<GlyphPainting, String> {
    programs().selected_glyph(family, bold, italic, index, bounds_em).map_err(|e| format!("Selected glyph: {e:?}"))
}

/// Every glyph a page selects must be in its registered program.
pub fn validate_layout_glyphs(layout: &oxidocs_core::layout::LayoutResult) -> Result<(), String> {
    let programs = programs();
    for element in layout.pages.iter().flat_map(|page| &page.elements) {
        if let Some(glyph) = &element.font_glyph {
            if let oxidocs_core::layout::LayoutContent::Text { font_family, bold, italic, text, .. } = &element.content {
                if text.is_empty() { continue; }
                let family = font_family.as_deref().ok_or("Selected glyph has no family")?;
                programs.selected_glyph(family, *bold, *italic, glyph.index, glyph.bounds_em)
                    .map_err(|e| format!("Selected glyph font program: {e:?}"))?;
            }
        }
    }
    Ok(())
}
