// This Source Code Form is subject to the terms of the Mozilla Public
// License, v. 2.0. If a copy of the MPL was not distributed with this
// file, You can obtain one at https://mozilla.org/MPL/2.0/.

//! Portable metrics for installed faces, generated without font programs.
//!
//! Calibrated registry entries remain authoritative. This catalog supplies
//! missing faces before disk resolution, so Linux and WASM do not depend on
//! the fonts installed on the Windows measurement machine. Each face is a
//! separate gzip member and is decoded only when requested.

use std::collections::HashMap;
use std::sync::{OnceLock, RwLock};

use serde::Deserialize;

use super::{FontMetrics, RawFontMetrics};

const DATA: &[u8] = include_bytes!("data/font_catalog_metrics.gz");

#[derive(Deserialize)]
struct Face {
    key: String,
    offset: usize,
    length: usize,
    families: Vec<String>,
    full_names: Vec<String>,
    bold: bool,
    italic: bool,
    weight: u16,
    width_class: u16,
    priority: u16,
    #[serde(default)]
    codepage_range1: Option<u32>,
    #[serde(default)]
    average_width_em: Option<f32>,
}

#[derive(Deserialize)]
struct Index {
    schema: u32,
    faces: Vec<Face>,
}

struct Catalog {
    faces: Vec<Face>,
    names: HashMap<String, Vec<usize>>,
    metrics: Vec<OnceLock<FontMetrics>>,
    glyph_geometry: Vec<OnceLock<Option<FaceGlyphGeometry>>>,
}

impl Catalog {
    fn load() -> Self {
        let index: Index = serde_json::from_str(include_str!("data/font_catalog_index.json"))
            .expect("generated font catalog index must be valid");
        assert_eq!(index.schema, 1, "unsupported font catalog schema");
        let mut names: HashMap<String, Vec<usize>> = HashMap::new();
        for (i, face) in index.faces.iter().enumerate() {
            assert!(
                face.offset
                    .checked_add(face.length)
                    .is_some_and(|end| end <= DATA.len()),
                "font catalog member is outside the metric data: {}",
                face.key
            );
            for name in face.families.iter().chain(&face.full_names) {
                let entries = names.entry(name.clone()).or_default();
                if entries.last() != Some(&i) {
                    entries.push(i);
                }
            }
        }
        let metrics = (0..index.faces.len()).map(|_| OnceLock::new()).collect();
        let glyph_geometry = (0..index.faces.len()).map(|_| OnceLock::new()).collect();
        Self {
            faces: index.faces,
            names,
            metrics,
            glyph_geometry,
        }
    }

    fn select(&self, name: &str, bold: bool, italic: bool) -> Option<usize> {
        self.names
            .get(name)?
            .iter()
            .copied()
            .filter(|&i| {
                let face = &self.faces[i];
                if face.families.iter().any(|n| n == name) {
                    // A family name must select the requested style, not whichever
                    // face happened to be enumerated first.
                    face.bold == bold && face.italic == italic
                } else {
                    // A full name can explicitly name a styled face. PostScript
                    // names alone are metadata, not Word family aliases.
                    (!bold || face.bold) && (!italic || face.italic)
                }
            })
            .min_by_key(|&i| {
                let face = &self.faces[i];
                (
                    face.priority,
                    face.weight.abs_diff(if bold { 700 } else { 400 }),
                    face.width_class.abs_diff(5),
                    i,
                )
            })
    }

    fn metrics(&self, i: usize) -> &FontMetrics {
        self.metrics[i].get_or_init(|| {
            let face = &self.faces[i];
            let decoder =
                flate2::read::GzDecoder::new(&DATA[face.offset..face.offset + face.length]);
            let raw: RawFontMetrics = serde_json::from_reader(decoder)
                .expect("generated font catalog member must contain metrics");
            let em = f32::from(raw.units_per_em);
            assert!(em > 0.0, "font catalog contains zero units per em");
            FontMetrics {
                synthetic_bold_advance: 0.0,
                average_width_em: face.average_width_em.filter(|v| v.is_finite() && *v > 0.0)
                    .or_else(|| raw.average_width.filter(|v| *v > 0).map(|v| v as f32 / em)),
                family: raw.family,
                units_per_em: raw.units_per_em,
                ascent: f32::from(raw.ascender) / em,
                descent: -f32::from(raw.descender) / em,
                line_gap: f32::from(raw.line_gap) / em,
                win_ascent: f32::from(raw.win_ascent) / em,
                win_descent: f32::from(raw.win_descent) / em,
                typo_ascent: f32::from(raw.typo_ascender) / em,
                typo_descent: -f32::from(raw.typo_descender) / em,
                typo_line_gap: f32::from(raw.typo_line_gap) / em,
                use_typo_metrics: raw.use_typo_metrics,
                codepage_range1: face.codepage_range1,
                sym_coverage: Vec::new(),
                char_widths: raw
                    .widths
                    .into_iter()
                    .filter_map(|(cp, advance)| {
                        char::from_u32(cp).map(|c| (c, f32::from(advance) / em))
                    })
                    .collect(),
            }
        })
    }
}

fn catalog() -> &'static Catalog {
    static CATALOG: OnceLock<Catalog> = OnceLock::new();
    CATALOG.get_or_init(Catalog::load)
}

/// Resolve a real styled face without sharing the regular face's calibrated
/// width-table key. A regular-family key would silently replace these advances.
pub(super) fn resolve_styled(family: &str, bold: bool, italic: bool) -> Option<&'static FontMetrics> {
    type Cache = HashMap<(String, bool, bool), &'static FontMetrics>;
    static CACHE: OnceLock<RwLock<Cache>> = OnceLock::new();
    let key = (family.trim().to_lowercase(), bold, italic);
    let cache = CACHE.get_or_init(|| RwLock::new(HashMap::new()));
    if let Some(metrics) = cache.read().expect("font catalog cache poisoned").get(&key) {
        return Some(*metrics);
    }
    let catalog = catalog();
    let index = catalog.select(&key.0, bold, italic)?;
    let face = &catalog.faces[index];
    let mut cache = cache.write().expect("font catalog cache poisoned");
    Some(*cache.entry(key.clone()).or_insert_with(|| {
        let mut metrics = catalog.metrics(index).clone();
        metrics.family = if face.families.iter().any(|name| name == &key.0) {
            let suffix = match (face.bold, face.italic) {
                (true, true) => " Bold Italic",
                (true, false) => " Bold",
                (false, true) => " Italic",
                (false, false) => "",
            };
            format!("{family}{suffix}")
        } else {
            family.to_owned()
        };
        Box::leak(Box::new(metrics))
    }))
}

/// A regular East Asian face without a real bold member is emboldened
/// rather than substituted. Resolve family aliases before checking members,
/// so an explicit regular full name does not hide a family's real bold face.
pub(super) fn synthetic_bold_eligible(family: &str) -> bool {
    let catalog = catalog();
    let name = family.trim().to_lowercase();
    let Some(index) = catalog.select(&name, false, false) else { return false; };
    let face = &catalog.faces[index];
    !face.bold && face.codepage_range1.is_some_and(|bits| bits & (0x1f << 17) != 0)
        && !face.families.iter().any(|alias| catalog.select(alias, true, false).is_some())
}

/// The catalog supplies this same fallback for faces outside the compact
/// calibrated registry. Metrics stay immutable and are cached per family.
pub(super) fn resolve_synthetic_bold(family: &str) -> Option<&'static FontMetrics> {
    if !synthetic_bold_eligible(family) { return None; }
    type Cache = HashMap<String, &'static FontMetrics>;
    static CACHE: OnceLock<RwLock<Cache>> = OnceLock::new();
    let key = family.trim().to_lowercase();
    let cache = CACHE.get_or_init(|| RwLock::new(HashMap::new()));
    if let Some(metrics) = cache.read().expect("font catalog cache poisoned").get(&key) {
        return Some(*metrics);
    }
    let regular = resolve(family, false, false)?;
    let mut cache = cache.write().expect("font catalog cache poisoned");
    Some(*cache.entry(key).or_insert_with(|| {
        let mut synthetic = regular.clone();
        synthetic.synthetic_bold_advance = 1.0 / f32::from(regular.units_per_em);
        Box::leak(Box::new(synthetic))
    }))
}

pub(super) fn average_width_em(family: &str) -> Option<f32> {
    let name = family.trim().to_lowercase();
    let catalog = catalog();
    let index = catalog.select(&name, false, false).or_else(|| {
        catalog.names.get(&name).and_then(|faces| faces.first().copied())
    })?;
    catalog.faces[index].average_width_em.filter(|v| v.is_finite() && *v > 0.0)
}

pub(super) fn codepage_range1(family: &str) -> Option<u32> {
    let name = family.trim().to_lowercase();
    let catalog = catalog();
    let index = catalog.select(&name, false, false).or_else(|| {
        catalog
            .names
            .get(&name)
            .and_then(|faces| faces.first().copied())
    })?;
    catalog.faces[index].codepage_range1
}

/// Preserve the requested name, as disk resolution does, because existing
/// calibrated spacing rules distinguish localized family names.
pub(super) fn resolve(family: &str, bold: bool, italic: bool) -> Option<&'static FontMetrics> {
    type Cache = HashMap<(String, bool, bool), &'static FontMetrics>;
    static CACHE: OnceLock<RwLock<Cache>> = OnceLock::new();
    let key = (family.trim().to_lowercase(), bold, italic);
    let cache = CACHE.get_or_init(|| RwLock::new(HashMap::new()));
    if let Some(metrics) = cache.read().expect("font catalog cache poisoned").get(&key) {
        return Some(*metrics);
    }
    let catalog = catalog();
    let index = catalog.select(&key.0, bold, italic)?;
    let mut cache = cache.write().expect("font catalog cache poisoned");
    Some(*cache.entry(key).or_insert_with(|| {
        let mut metrics = catalog.metrics(index).clone();
        metrics.family = family.to_owned();
        Box::leak(Box::new(metrics))
    }))
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn every_catalog_member_has_valid_metrics() {
        let catalog = catalog();
        assert!(!catalog.faces.is_empty());
        for i in 0..catalog.faces.len() {
            let face = &catalog.faces[i];
            let metrics = catalog.metrics(i);
            assert!(metrics.units_per_em > 0, "{}", face.key);
            assert!(
                !metrics.char_widths.is_empty(),
                "{} has no character map",
                face.key
            );
            assert!(metrics
                .char_widths
                .values()
                .all(|w| w.is_finite() && *w >= 0.0));
        }
    }

    #[test]
    fn localized_names_share_widths_without_an_installed_font() {
        let en = resolve("HGPSoeiPresenceEB", false, false).unwrap();
        let ja = resolve("HGP創英ﾌﾟﾚｾﾞﾝｽEB", false, false).unwrap();
        assert_eq!(en.units_per_em, 256);
        assert_eq!(en.char_widths, ja.char_widths);
        assert_eq!(en.char_widths[&'A'], 170.0 / 256.0);
        assert_eq!(en.char_widths.len(), 7484);
    }

    #[test]
    fn styles_select_distinct_faces_and_missing_names_do_not_substitute() {
        let regular = resolve("Gill Sans Nova", false, false).unwrap();
        let italic = resolve("Gill Sans Nova", false, true).unwrap();
        assert_ne!(regular.char_widths, italic.char_widths);
        assert!(resolve("Oxi nonexistent catalog fixture", false, false).is_none());
        assert!(resolve("MS-Mincho", false, false).is_none());
    }
}

#[derive(Deserialize)]
struct FaceGlyphGeometry {
    face_key: String,
    units_per_em: u16,
    has_math: bool,
    glyphs: HashMap<u32, [f32; 7]>,
    script_glyphs: HashMap<u8, HashMap<u32, [f32; 7]>>,
    #[serde(default)]
    corner_kerns: HashMap<u16, HashMap<String, CatalogCornerKern>>,
}

#[derive(Deserialize)]
struct CatalogCornerKern { heights: Vec<i32>, values: Vec<i32> }

#[derive(Deserialize)]
struct GlyphGeometryRecord {
    #[serde(default)]
    glyph_geometry: Option<FaceGlyphGeometry>,
}

#[derive(Debug, Clone, Copy)]
pub(crate) struct CatalogGlyphMetrics {
    pub index: u16,
    pub advance_em: f32,
    pub bounds_em: [f32; 4],
    pub italic_correction_em: f32,
    pub has_math: bool,
}

/// Select through the same family/full-name/style index as ordinary metrics.
/// Missing geometry is explicit; it is never synthesized from a line box.
pub(crate) fn glyph_metrics(family: &str, bold: bool, italic: bool, c: char, script_level: u8)
    -> Option<CatalogGlyphMetrics>
{
    if super::runtime::resolve_registered(family, bold, italic).is_some() {
        let g = super::runtime::registered_glyph(family, bold, italic, c, script_level)?;
        return Some(CatalogGlyphMetrics { index: g.index, advance_em: g.advance_em,
            bounds_em: g.bounds_em, italic_correction_em: g.italic_correction_em,
            has_math: super::runtime::registered_has_math(family, bold, italic) });
    }
    let catalog = catalog();
    let name = family.trim().to_lowercase();
    let i = catalog.select(&name, bold, italic)
        .or_else(|| catalog.select(&name, false, false))?;
    let face = &catalog.faces[i];
    let data = catalog.glyph_geometry[i].get_or_init(|| {
        let decoder = flate2::read::GzDecoder::new(&DATA[face.offset..face.offset+face.length]);
        let record: GlyphGeometryRecord = serde_json::from_reader(decoder)
            .expect("generated catalog glyph geometry must be valid");
        if let Some(ref geometry) = record.glyph_geometry {
            assert_eq!(geometry.face_key, face.key, "glyph geometry face identity mismatch");
            assert!(geometry.units_per_em > 0, "glyph geometry has zero units per em");
        }
        record.glyph_geometry
    }).as_ref()?;
    let cp = c as u32;
    let values = data.script_glyphs.get(&script_level).and_then(|map|map.get(&cp))
        .or_else(||data.glyphs.get(&cp))?;
    let em = f32::from(data.units_per_em);
    Some(CatalogGlyphMetrics { index: values[0] as u16, advance_em: values[1]/em,
        bounds_em: [values[2]/em, values[3]/em, values[4]/em, values[5]/em],
        italic_correction_em: values[6]/em, has_math: data.has_math })
}

pub(crate) fn glyph_corner_kern(family:&str,bold:bool,italic:bool,gid:u16,
    corner:super::math_kern::Corner,height:f32,size:f32)->f32
{
    if let Some(value) = super::runtime::registered_corner_kern(family, bold, italic, gid, corner, height, size) {
        return value;
    }
    if !size.is_finite() || size<=0.0 {return 0.0;}
    // Resolve through the identical selector and lazily initialize geometry.
    let _=glyph_metrics(family,bold,italic,' ',0);
    let catalog=catalog();let name=family.trim().to_lowercase();
    let Some(i)=catalog.select(&name,bold,italic).or_else(||catalog.select(&name,false,false))else{return 0.0;};
    let Some(data)=catalog.glyph_geometry[i].get().and_then(|data|data.as_ref())else{return 0.0;};
    let key=match corner {
        super::math_kern::Corner::TopRight=>"top_right",super::math_kern::Corner::TopLeft=>"top_left",
        super::math_kern::Corner::BottomRight=>"bottom_right",super::math_kern::Corner::BottomLeft=>"bottom_left",
    };
    let Some(kern)=data.corner_kerns.get(&gid).and_then(|map|map.get(key))else{return 0.0;};
    let em=f32::from(data.units_per_em);let design_height=height*em/size;
    let index=kern.heights.partition_point(|h|*h as f32<=design_height);
    kern.values.get(index).map_or(0.0,|v|*v as f32*size/em)
}

/// Font-box metrics from the same selected face as numeric glyph geometry.
pub(crate) fn glyph_face_metrics(family:&str,bold:bool,italic:bool)->Option<super::FontMetricsRef<'static>> {
    if let Some(metrics) = super::runtime::resolve_registered(family, bold, italic) { return Some(metrics); }
    let catalog=catalog();let name=family.trim().to_lowercase();
    let index=catalog.select(&name,bold,italic).or_else(||catalog.select(&name,false,false))?;
    Some(catalog.metrics(index).into())
}


#[derive(Deserialize)]
struct RuleFaceMetrics {
    units_per_em: u16,
    underline_thickness: i16,
}

#[derive(Deserialize)]
struct RuleMetricIndex {
    schema: u32,
    faces: HashMap<String, RuleFaceMetrics>,
}

/// Numeric rule thickness from the same selected face as glyph geometry.
/// An unmeasured face remains absent; no other face's value is substituted.
pub(crate) fn glyph_rule_thickness(family: &str, bold: bool, italic: bool) -> Option<f32> {
    if super::runtime::resolve_registered(family, bold, italic).is_some() {
        return super::runtime::registered_rule_thickness(family, bold, italic);
    }
    static RULES: OnceLock<RuleMetricIndex> = OnceLock::new();
    let rules = RULES.get_or_init(|| {
        let index: RuleMetricIndex = serde_json::from_str(include_str!("data/font_rule_metrics.json"))
            .expect("generated rule metrics must be valid");
        assert_eq!(index.schema, 1, "unsupported rule metric schema");
        index
    });
    let catalog = catalog();
    let name = family.trim().to_lowercase();
    let i = catalog.select(&name, bold, italic)
        .or_else(|| catalog.select(&name, false, false))?;
    let raw = rules.faces.get(&catalog.faces[i].key)?;
    if raw.units_per_em == 0 || raw.underline_thickness <= 0 { return None; }
    Some(f32::from(raw.underline_thickness) / f32::from(raw.units_per_em))
}
