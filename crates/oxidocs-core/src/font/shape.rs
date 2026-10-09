// This Source Code Form is subject to the terms of the Mozilla Public
// License, v. 2.0. If a copy of the MPL was not distributed with this
// file, You can obtain one at https://mozilla.org/MPL/2.0/.

//! Complex-script shaping (Devanagari and its kin) via rustybuzz.
//!
//! The per-codepoint advance sum the layout uses everywhere else is wrong for a
//! complex script: a virama conjunct (क् + ष = क्ष) shapes to ONE narrow glyph,
//! not the two full consonants summed, and non-spacing matras / marks that the
//! sum still counts collapse onto their base. MEASURED against Word (Nirmala UI
//! 20pt): क्ष is 14.25pt, the sum is 29.5; राष्ट्रीय is 44.25, the sum 66.5 —
//! while a conjunct-free word (सरकार, भारत) already matched within ~1pt.
//! rustybuzz — the same OpenType shaping model DirectWrite and Word use —
//! reproduces every one within 0.3pt, so the shaped cluster advance is the
//! right width source.
//!
//! This module is reached ONLY for a run that carries a complex-script
//! character (see `is_complex_script`), which no document in the frozen
//! Latin/CJK corpus does — so that corpus is byte-identical by construction.

use std::cell::RefCell;
use std::collections::HashMap;

struct ShapeFace<'a> {
    face: rustybuzz::Face<'a>,
    upm: f32,
}

#[derive(Default)]
struct ShapeCache {
    generation: usize,
    faces: HashMap<(String, bool, bool), Option<ShapeFace<'static>>>,
}

thread_local! {
    // (family, bold, italic) -> the shapeable face, or None when this machine
    // has no such file. rustybuzz::Face is not Send, so the cache is per-thread;
    // layout runs one document on one thread.
    static FACES: RefCell<ShapeCache> = RefCell::new(ShapeCache::default());
}

thread_local! {
    // Faces embedded in the document being laid out (default ON, opt-out OXI_EMBEDDED_FONTS_DISABLE).
    static DOCUMENT_FONTS: RefCell<HashMap<(String, bool, bool), Vec<u8>>> = RefCell::new(HashMap::new());
}

/// Restores the previous document's faces when the layout that set them ends.
pub(crate) struct DocumentFontScope(HashMap<(String, bool, bool), Vec<u8>>);
impl Drop for DocumentFontScope {
    fn drop(&mut self) {
        DOCUMENT_FONTS.with(|fonts| *fonts.borrow_mut() = std::mem::take(&mut self.0));
    }
}

pub(crate) fn enter_document_fonts(doc: &crate::ir::Document) -> DocumentFontScope {
    let styles = &doc.styles;
    let mut current = HashMap::new();
    if std::env::var_os("OXI_EMBEDDED_FONTS_DISABLE").is_none() {
        let usable = super::usable_embedded_families(doc);
        for (name, info) in &styles.font_table {
            if !usable.contains(&name.to_lowercase()) {
                continue;
            }
            if info.embedded_faces.is_empty()
                || super::runtime::resolve(name, false, false).is_some()
                || super::FontMetricsRegistry::load().supports_family(name)
            {
                continue;
            }
            for face in &info.embedded_faces {
                current.insert((name.to_lowercase(), face.bold, face.italic), face.data.clone());
            }
        }
    }
    DocumentFontScope(DOCUMENT_FONTS.with(|fonts| std::mem::replace(&mut *fonts.borrow_mut(), current)))
}

fn with_face<R>(
    family: &str,
    bold: bool,
    italic: bool,
    f: impl FnOnce(Option<&ShapeFace<'_>>) -> R,
) -> R {
    let mut callback = Some(f);
    let embedded = DOCUMENT_FONTS.with(|fonts| {
        let fonts = fonts.borrow();
        let data = fonts.get(&(family.to_lowercase(), bold, italic))
            .or_else(|| fonts.get(&(family.to_lowercase(), false, false)))?;
        let face = rustybuzz::Face::from_slice(data, 0)?;
        let upm = face.units_per_em() as f32;
        Some(callback.take().unwrap()(Some(&ShapeFace { face, upm })))
    });
    if let Some(result) = embedded {
        return result;
    }
    let f = callback.unwrap();
    // Borrow a registered programme only within this shaping call. The Arc
    // keeps the font alive through a concurrent clear or replacement, and no
    // self-referential/static face retains superseded caller bytes.
    let registered = super::runtime::registered_font_file_for(family, bold, italic)
        .or_else(|| if bold || italic {
            super::runtime::registered_font_file_for(family, false, false)
        } else { None });
    if let Some((bytes, index)) = registered {
        let shaped = rustybuzz::Face::from_slice(&bytes, index).and_then(|face| {
            let upm = face.units_per_em() as f32;
            (upm > 0.0).then_some(ShapeFace { face, upm })
        });
        return f(shaped.as_ref());
    }
    FACES.with(|c| {
        let mut cache = c.borrow_mut();
        let generation = super::runtime::memory_font_generation();
        if cache.generation != generation {
            cache.faces.clear();
            cache.generation = generation;
        }
        let m = &mut cache.faces;
        // Try the requested style; if the machine lacks that exact face, fall
        // back to the regular face of the same family rather than giving up
        // (a bold Devanagari run on a box with only the regular file still
        // shapes correctly, just without the bold widths).
        let key = (family.to_string(), bold, italic);
        if !m.contains_key(&key) {
            let built = build_face(family, bold, italic)
                .or_else(|| if bold || italic { build_face(family, false, false) } else { None });
            m.insert(key.clone(), built);
        }
        f(m.get(&key).and_then(|o| o.as_ref()))
    })
}

fn build_face(family: &str, bold: bool, italic: bool) -> Option<ShapeFace<'static>> {
    let (bytes, idx) = super::runtime::font_file_for(family, bold, italic)?;
    let face = rustybuzz::Face::from_slice(bytes, idx)?;
    let upm = face.units_per_em() as f32;
    if upm <= 0.0 {
        return None;
    }
    Some(ShapeFace { face, upm })
}

/// Is `family` installed on this machine and does it carry a glyph for `c`?
/// Used to decide whether a run's own font can draw the script or Word's
/// fallback (Nirmala UI, the Windows Devanagari UI font) takes over.
pub fn family_covers(family: &str, c: char) -> bool {
    with_face(family, false, false, |sf| {
        sf.map_or(false, |sf| sf.face.glyph_index(c).is_some())
    })
}

/// Per-CHARACTER advances (points) for `text` shaped with `family`, with each
/// shaped cluster's total advance placed on the cluster's FIRST character and
/// `0.0` on the rest, plus a parallel flag marking those cluster-first chars.
///
/// The layout only needs a prefix sum that is correct at cluster boundaries
/// (a complex script never breaks inside a cluster), which this gives: summing
/// the advances of any run of whole clusters reproduces the shaped width. The
/// flag lets the caller add letter-spacing once per cluster, as Word does.
/// `None` when the family is not installed / not shapeable.
pub fn cluster_advances(
    family: &str,
    bold: bool,
    italic: bool,
    text: &str,
    font_size: f32,
) -> Option<(Vec<f32>, Vec<bool>)> {
    with_face(family, bold, italic, |sf| {
        let sf = sf?;
        let n = text.chars().count();
        if n == 0 {
            return Some((Vec::new(), Vec::new()));
        }
        let mut buf = rustybuzz::UnicodeBuffer::new();
        buf.push_str(text);
        let glyphs = rustybuzz::shape(&sf.face, &[], buf);
        let infos = glyphs.glyph_infos();
        let pos = glyphs.glyph_positions();

        // Cluster values are byte offsets into `text`, always on a char
        // boundary. Sum every glyph's advance into its cluster.
        let mut cluster_adv: std::collections::BTreeMap<u32, i32> = Default::default();
        for (info, p) in infos.iter().zip(pos.iter()) {
            *cluster_adv.entry(info.cluster).or_insert(0) += p.x_advance;
        }
        // byte offset (of each char start) -> char index
        let mut byte_to_char: HashMap<usize, usize> = HashMap::with_capacity(n);
        for (ci, (bi, _)) in text.char_indices().enumerate() {
            byte_to_char.insert(bi, ci);
        }
        let mut adv = vec![0.0f32; n];
        let mut start = vec![false; n];
        for (cl, a) in cluster_adv {
            if let Some(&ci) = byte_to_char.get(&(cl as usize)) {
                adv[ci] += a as f32 / sf.upm * font_size;
                start[ci] = true;
            }
        }
        Some((adv, start))
    })
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn absent_font_is_none_not_a_panic() {
        // The shaper must never panic on a font this machine lacks; the caller
        // keeps its own path. (Kept assertion-light: WHICH fonts are installed
        // is a property of the machine, like runtime::resolve's tests.)
        assert!(cluster_advances("Zzquartz No Such Face", false, false, "\u{0915}", 12.0).is_none());
        assert!(!family_covers("Zzquartz No Such Face", '\u{0915}'));
    }

    #[test]
    fn empty_text_shapes_to_empty() {
        // Any installed or absent font: no chars -> no advances, no panic.
        if let Some((adv, start)) = cluster_advances("Nirmala UI", false, false, "", 12.0) {
            assert!(adv.is_empty() && start.is_empty());
        }
    }

    #[test]
    fn conjunct_is_narrower_than_the_codepoint_sum() {
        // The whole point of shaping: क् + ष = क्ष is ONE narrow glyph, so the
        // cluster advance is far below the two consonants summed. Only asserted
        // where a Devanagari font is installed (Nirmala UI on Windows); skipped
        // otherwise so CI without the face still passes.
        let ks = "\u{0915}\u{094D}\u{0937}"; // क्ष
        if let Some((adv, start)) = cluster_advances("Nirmala UI", false, false, ks, 20.0) {
            let total: f32 = adv.iter().sum();
            // MEASURED against Word: 14.25pt; the per-codepoint sum is 29.5.
            assert!(total > 8.0 && total < 20.0, "क्ष cluster width {total} off");
            // one cluster: the first char carries it, the rest are zero.
            assert!(start[0]);
            assert_eq!(start.iter().filter(|&&s| s).count(), 1);
        }
    }
}
