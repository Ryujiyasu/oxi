// This Source Code Form is subject to the terms of the Mozilla Public
// License, v. 2.0. If a copy of the MPL was not distributed with this
// file, You can obtain one at https://mozilla.org/MPL/2.0/.

//! Cambria Math per-glyph MATH table data.
//!
//! Loaded from `data/cambria_math_glyph_tables.json` (180KB) via
//! `include_str!`. Exposes three lookup functions indexed by Unicode
//! codepoint (after `math_substitute` has been applied):
//!
//! - `italic_correction(cp)`: extra horizontal space after italic
//!   glyphs before a superscript (e.g., ∫ needs 415 DU at 2048 UPM).
//! - `top_accent_attachment(cp)`: horizontal offset from glyph origin
//!   to the optical center where accents should align (e.g., Α has
//!   attachment 668 DU, arranging ^ directly above).
//! - `vertical_variants(cp)`: ordered list of grow variants for
//!   stretchy glyphs (∑, ∫, √, parentheses, etc.) with their advance
//!   measurement.
//!
//! All values are in font design units (UPM = 2048). Divide by UPM and
//! multiply by font size to get points.

use serde::Deserialize;
use std::collections::HashMap;
use std::sync::OnceLock;

/// Italic correction entry for one glyph.
#[derive(Debug, Clone, Deserialize)]
struct ItalicCorrEntry {
    #[serde(default)]
    codepoint: Option<u32>,
    italic_correction: i32,
}

/// Top-accent attachment entry for one glyph.
#[derive(Debug, Clone, Deserialize)]
struct TopAccentEntry {
    #[serde(default)]
    codepoint: Option<u32>,
    top_accent_attachment: i32,
}

/// One grow-variant of a stretchy glyph.
#[derive(Debug, Clone, Deserialize)]
pub struct GlyphVariant {
    /// Variant glyph name (font-internal).
    pub variant_glyph: String,
    /// Height (for vertical) or width (for horizontal) in design units.
    pub advance: i32,
}

/// Vertical-variant entry for one base glyph.
#[derive(Debug, Clone, Deserialize)]
struct VertVariantEntry {
    #[serde(default)]
    codepoint: Option<u32>,
    variants: Vec<GlyphVariant>,
}

/// Raw file schema.
#[derive(Debug, Deserialize)]
struct GlyphTablesFile {
    #[serde(rename = "italic_correction")]
    italic_correction: Vec<ItalicCorrEntry>,
    #[serde(rename = "top_accent_attachment")]
    top_accent_attachment: Vec<TopAccentEntry>,
    #[serde(rename = "vertical_variants")]
    vertical_variants: Vec<VertVariantEntry>,
}

/// Parsed, indexed lookup tables. Built once from bundled JSON.
pub struct MathGlyphTables {
    italic_corr: HashMap<u32, i32>,
    top_accent: HashMap<u32, i32>,
    vert_variants: HashMap<u32, Vec<GlyphVariant>>,
}

impl MathGlyphTables {
    /// Load bundled Cambria Math glyph tables (lazy, cached).
    pub fn cambria_math() -> &'static Self {
        static TABLES: OnceLock<MathGlyphTables> = OnceLock::new();
        TABLES.get_or_init(|| {
            let json = include_str!("data/cambria_math_glyph_tables.json");
            let file: GlyphTablesFile = serde_json::from_str(json)
                .expect("cambria_math_glyph_tables.json parse");

            let mut italic_corr: HashMap<u32, i32> = HashMap::new();
            for e in file.italic_correction {
                if let Some(cp) = e.codepoint {
                    italic_corr.insert(cp, e.italic_correction);
                }
            }

            let mut top_accent: HashMap<u32, i32> = HashMap::new();
            for e in file.top_accent_attachment {
                if let Some(cp) = e.codepoint {
                    top_accent.insert(cp, e.top_accent_attachment);
                }
            }

            let mut vert_variants: HashMap<u32, Vec<GlyphVariant>> = HashMap::new();
            for e in file.vertical_variants {
                if let Some(cp) = e.codepoint {
                    if !e.variants.is_empty() {
                        vert_variants.insert(cp, e.variants);
                    }
                }
            }

            MathGlyphTables { italic_corr, top_accent, vert_variants }
        })
    }

    /// Italic correction in design units for the given character.
    /// Returns None when the character has no entry (geometric bbox can
    /// be used as a fallback in that case).
    pub fn italic_correction(&self, c: char) -> Option<i32> {
        self.italic_corr.get(&(c as u32)).copied()
    }

    /// Top-accent attachment horizontal offset in design units.
    /// When None, use half the glyph advance as geometric center.
    pub fn top_accent_attachment(&self, c: char) -> Option<i32> {
        self.top_accent.get(&(c as u32)).copied()
    }

    /// Select the smallest vertical grow-variant whose advance ≥ target_du.
    /// Returns `None` if the glyph has no variants (use the base glyph).
    /// If all variants are smaller than target, returns the largest variant
    /// (assembly from parts would be the next step, not implemented here).
    pub fn select_vertical_variant(
        &self,
        c: char,
        target_du: i32,
    ) -> Option<&GlyphVariant> {
        let variants = self.vert_variants.get(&(c as u32))?;
        // Variants are ordered smallest → largest; pick first that fits.
        variants.iter()
            .find(|v| v.advance >= target_du)
            .or_else(|| variants.last())
    }

    /// Raw list of vertical grow variants for a character.
    pub fn vertical_variants(&self, c: char) -> &[GlyphVariant] {
        self.vert_variants
            .get(&(c as u32))
            .map(|v| v.as_slice())
            .unwrap_or(&[])
    }

    /// Number of glyphs with italic correction data.
    pub fn n_italic_correction(&self) -> usize { self.italic_corr.len() }

    /// Number of glyphs with top-accent attachment data.
    pub fn n_top_accent(&self) -> usize { self.top_accent.len() }

    /// Number of stretchy base glyphs with vertical variants.
    pub fn n_vertical_variant_bases(&self) -> usize { self.vert_variants.len() }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn tables_load() {
        let t = MathGlyphTables::cambria_math();
        // From Phase 1 measurement:
        assert!(t.n_italic_correction() > 100, "italic_corr = {}", t.n_italic_correction());
        assert!(t.n_top_accent() > 300, "top_accent = {}", t.n_top_accent());
        assert!(t.n_vertical_variant_bases() > 40, "variants = {}", t.n_vertical_variant_bases());
    }

    #[test]
    fn italic_correction_known_values() {
        let t = MathGlyphTables::cambria_math();
        // Integral ∫ (U+222B) has italic correction 415 (Phase 1 verified).
        assert_eq!(t.italic_correction('∫'), Some(415));
    }

    #[test]
    fn top_accent_attachment_known() {
        let t = MathGlyphTables::cambria_math();
        // Α (U+0391 Greek Alpha) has attachment 668 per Phase 1 measurement.
        assert_eq!(t.top_accent_attachment('Α'), Some(668));
        // Combining grave (U+0300) has negative attachment -364.
        assert_eq!(t.top_accent_attachment('\u{0300}'), Some(-364));
    }

    #[test]
    fn vertical_variant_selection_radical() {
        let t = MathGlyphTables::cambria_math();
        // √ (U+221A) has 6 variants with advances 1972, 2544, ..., 4569.
        // Requesting 2000 DU should pick variant ≥ 2000, which is 2544.
        let v = t.select_vertical_variant('√', 2000);
        assert!(v.is_some(), "no variant returned");
        let v = v.unwrap();
        assert!(v.advance >= 2000, "got advance {}", v.advance);
        // Default √ has 6+ variants
        assert!(t.vertical_variants('√').len() >= 6);
    }

    #[test]
    fn vertical_variant_sum() {
        let t = MathGlyphTables::cambria_math();
        let variants = t.vertical_variants('∑');
        assert!(!variants.is_empty(), "∑ has no variants");
        // Smallest variant should be smaller than largest.
        assert!(variants[0].advance < variants[variants.len()-1].advance);
    }

    #[test]
    fn unknown_glyph_returns_none() {
        let t = MathGlyphTables::cambria_math();
        // Japanese 日 should not have any MATH-table data.
        assert_eq!(t.italic_correction('日'), None);
        assert_eq!(t.top_accent_attachment('日'), None);
        assert!(t.vertical_variants('日').is_empty());
    }

    #[test]
    fn huge_target_returns_largest() {
        let t = MathGlyphTables::cambria_math();
        // Asking for a much larger radical than available should return
        // the largest variant (assembly-from-parts would handle larger).
        let v = t.select_vertical_variant('√', 999_999);
        let vs = t.vertical_variants('√');
        assert_eq!(v.unwrap().advance, vs.last().unwrap().advance);
    }
}
