// This Source Code Form is subject to the terms of the Mozilla Public
// License, v. 2.0. If a copy of the MPL was not distributed with this
// file, You can obtain one at https://mozilla.org/MPL/2.0/.

//! Numeric metrics for OpenType `ssty` alternate glyphs. Script font scaling
//! and script glyph selection are separate operations; these metrics are in em.
use serde::Deserialize;
use std::collections::HashMap;
use std::sync::OnceLock;

#[derive(Debug, Clone, Copy, Deserialize)]
pub struct ScriptGlyph {
    pub index: u16,
    pub advance_em: f32,
    pub bounds_em: [f32; 4],
    pub italic_correction_em: f32,
    pub top_accent_attachment_em: Option<f32>,
}

#[derive(Deserialize)]
pub struct MathScriptGlyphs {
    forms: HashMap<u32, Vec<ScriptGlyph>>,
}

impl MathScriptGlyphs {
    pub fn cambria_math() -> &'static Self {
        static TABLE: OnceLock<MathScriptGlyphs> = OnceLock::new();
        TABLE.get_or_init(|| serde_json::from_str(
            include_str!("data/cambria_math_script_glyphs.json"))
            .expect("embedded script glyph metrics must be valid"))
    }

    /// Level zero uses the ordinary character glyph. Fonts may supply one or
    /// two alternate shapes; deeper scripts use the final available alternate.
    pub fn alternate(&self, c: char, level: u8) -> Option<ScriptGlyph> {
        if super::runtime::resolve_registered("Cambria Math", false, false).is_some() {
            return super::runtime::registered_glyph("Cambria Math", false, false, c, level);
        }
        if level == 0 { return None; }
        let forms = self.forms.get(&(c as u32))?;
        if forms.len() < 2 { return None; }
        forms.get((level as usize).min(forms.len() - 1)).copied()
    }
}
