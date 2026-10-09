// This Source Code Form is subject to the terms of the Mozilla Public
// License, v. 2.0. If a copy of the MPL was not distributed with this
// file, You can obtain one at https://mozilla.org/MPL/2.0/.

//! Resolve a font Word can resolve but the metrics tables do not carry.
//!
//! S1171 (2026-08-19, default ON, opt-out `OXI_S1171_DISABLE`).
//!
//! The shipped tables cover the faces measured on CI. A document may name a
//! face that is installed HERE and that Word therefore lays out correctly,
//! while Oxi, finding no table, decides the font is unresolvable and takes the
//! S1146 script fallback (Latin-named → Cambria). That is the wrong answer to
//! the wrong question: S1146 is about what Word does when IT cannot resolve a
//! name, so applying it to a name Word CAN resolve invents a different font.
//!
//! `educational__00252fa88ac64d0d` is the specimen. Its Normal style is
//! `Gill Sans Nova` 16pt justified; Word embeds that face in its PDF, so Word
//! resolved it. The machine has it only as an OFFICE CLOUD FONT, under
//! `%LOCALAPPDATA%\Microsoft\FontCache\4\CloudFonts\<Family>\<id>.ttf`, where
//! the files are named by numeric id — so a face is identifiable only by
//! reading its own `name` table. Oxi laid the document out in Cambria, wrapped
//! one line early, and lost a line at the page bottom.
//!
//! Three sources are searched, which is the set Word itself uses (the
//! `font_audit_three_sources` note): the Office cloud cache, the per-user font
//! directory, and the system font directory.
//!
//! Metrics come from the file, not from a formula: head/hhea/OS-2 for the
//! vertical box and cmap+hmtx for every advance, all normalised to 1em exactly
//! as the generated tables are. A face that cannot be found or cannot be
//! parsed returns `None` and the caller keeps its existing fallback, so this
//! can only ever REPLACE an invented font with the real one.
//!
//! Resolved faces are leaked and cached: a process sees a handful of families,
//! and leaking lets the registry keep handing out `&FontMetrics` without
//! changing every signature to thread a lifetime or a lock guard.

use std::collections::HashMap;
use std::path::{Path, PathBuf};
use std::sync::{OnceLock, RwLock};

use super::{FontMetrics, FontMetricsRef};

/// family/bold/italic → the resolved face, or `None` when it is not installed.
/// `None` is cached too: a miss costs one directory walk per process.
fn cache() -> &'static RwLock<HashMap<(String, bool, bool), Option<&'static FontMetrics>>> {
    static CACHE: OnceLock<RwLock<HashMap<(String, bool, bool), Option<&'static FontMetrics>>>> =
        OnceLock::new();
    CACHE.get_or_init(|| RwLock::new(HashMap::new()))
}

/// Directories Word draws faces from, in the order it prefers them.
fn search_roots() -> Vec<PathBuf> {
    let mut out = Vec::new();
    if let Ok(local) = std::env::var("LOCALAPPDATA") {
        // The Office cloud cache keeps one directory per family, holding files
        // named by numeric id -- the family is the DIRECTORY, not the filename.
        out.push(Path::new(&local).join(r"Microsoft\FontCache\4\CloudFonts"));
        out.push(Path::new(&local).join(r"Microsoft\Windows\Fonts"));
    }
    if let Ok(win) = std::env::var("SystemRoot") {
        out.push(Path::new(&win).join("Fonts"));
    } else {
        out.push(PathBuf::from(r"C:\Windows\Fonts"));
    }
    out
}

fn is_font_file(p: &Path) -> bool {
    matches!(
        p.extension().and_then(|e| e.to_str()).map(str::to_ascii_lowercase).as_deref(),
        Some("ttf") | Some("otf") | Some("ttc")
    )
}

/// Every font file under `root`, one level deep (the cloud cache nests by family).
fn candidate_files(root: &Path) -> Vec<PathBuf> {
    let mut out = Vec::new();
    let Ok(entries) = std::fs::read_dir(root) else {
        return out;
    };
    for e in entries.flatten() {
        let p = e.path();
        if p.is_dir() {
            if let Ok(inner) = std::fs::read_dir(&p) {
                out.extend(inner.flatten().map(|i| i.path()).filter(|i| is_font_file(i)));
            }
        } else if is_font_file(&p) {
            out.push(p);
        }
    }
    out
}

/// Does this face's own `name` table say it is the family/style we want?
fn face_matches(font: &skrifa::FontRef, family: &str, bold: bool, italic: bool) -> bool {
    use skrifa::MetadataProvider;
    let want = family.trim().to_ascii_lowercase();
    let mut fam_ok = false;
    let mut style = String::new();
    for rec in font.localized_strings(skrifa::string::StringId::FAMILY_NAME) {
        if rec.to_string().trim().to_ascii_lowercase() == want {
            fam_ok = true;
        }
    }
    // TYPOGRAPHIC_FAMILY_NAME carries the real family when the legacy field was
    // split into 4-style groups ("Gill Sans Nova Cond Lt" and friends).
    if !fam_ok {
        for rec in font.localized_strings(skrifa::string::StringId::TYPOGRAPHIC_FAMILY_NAME) {
            if rec.to_string().trim().to_ascii_lowercase() == want {
                fam_ok = true;
            }
        }
    }
    if !fam_ok {
        return false;
    }
    if let Some(rec) = font
        .localized_strings(skrifa::string::StringId::SUBFAMILY_NAME)
        .next()
    {
        style = rec.to_string().to_ascii_lowercase();
    }
    let has_bold = style.contains("bold");
    let has_italic = style.contains("italic") || style.contains("oblique");
    has_bold == bold && has_italic == italic
}

/// Build the metrics the layout engine needs, straight out of the file.
pub(super) fn metrics_from(font: &skrifa::FontRef, family: &str) -> Option<FontMetrics> {
    use skrifa::raw::TableProvider;
    use skrifa::MetadataProvider;

    let upm = font.head().ok()?.units_per_em();
    let em = upm as f32;
    let hhea = font.hhea().ok()?;
    let (asc, desc, gap) = (
        hhea.ascender().to_i16() as f32 / em,
        (-(hhea.descender().to_i16() as f32)) / em,
        hhea.line_gap().to_i16() as f32 / em,
    );

    // OS/2 is optional in principle; fall back to the hhea box rather than
    // inventing zeros, which would collapse the line height.
    let (win_a, win_d, typo_a, typo_d, typo_gap, use_typo) = match font.os2() {
        Ok(os2) => (
            os2.us_win_ascent() as f32 / em,
            os2.us_win_descent() as f32 / em,
            os2.s_typo_ascender() as f32 / em,
            (-(os2.s_typo_descender() as f32)) / em,
            os2.s_typo_line_gap() as f32 / em,
            os2.fs_selection().bits() & 0x80 != 0,
        ),
        Err(_) => (asc, desc, asc, desc, gap, false),
    };

    // Every advance the document could ask for, normalised to 1em. Walking the
    // charmap (rather than a fixed ASCII range) keeps punctuation and the
    // curly quotes this corpus is full of measured rather than guessed.
    let charmap = font.charmap();
    let glyph_metrics = font.glyph_metrics(skrifa::instance::Size::unscaled(), skrifa::instance::LocationRef::default());
    let mut char_widths = HashMap::new();
    for (cp, gid) in charmap.mappings() {
        if let Some(ch) = char::from_u32(cp) {
            if let Some(adv) = glyph_metrics.advance_width(gid) {
                char_widths.insert(ch, adv / em);
            }
        }
    }
    if char_widths.is_empty() {
        return None;
    }

    Some(FontMetrics {
        synthetic_bold_advance: 0.0,
        average_width_em: font.os2().ok().map(|t| t.x_avg_char_width() as f32 / em)
            .filter(|v| v.is_finite() && *v > 0.0),
        family: family.to_string(),
        units_per_em: upm,
        ascent: asc,
        descent: desc,
        line_gap: gap,
        win_ascent: win_a,
        win_descent: win_d,
        typo_ascent: typo_a,
        typo_descent: typo_d,
        typo_line_gap: typo_gap,
        use_typo_metrics: use_typo,
        codepage_range1: font.os2().ok().and_then(|os2| os2.ul_code_page_range_1()),
        sym_coverage: Vec::new(),
        char_widths,
    })
}

fn load_from_disk(family: &str, bold: bool, italic: bool) -> Option<&'static FontMetrics> {
    for root in search_roots() {
        for path in candidate_files(&root) {
            let Ok(data) = std::fs::read(&path) else {
                continue;
            };
            // S1272 (2026-09-02): a .ttc holds SEVERAL faces and `FontRef::new`
            // reads none of them -- it only accepts a single-font file. Windows
            // ships most of the Japanese families that way, and the one the
            // document names is rarely the first face in the file:
            //
            //   BIZ-UDGothicR.ttc  -> BIZ UDゴシック (monospaced) + BIZ UDPゴシック
            //   meiryo.ttc         -> メイリオ + Meiryo UI
            //   YuGothM.ttc / HG*  -> likewise
            //
            // So every one of those resolved to None, the layout kept the em as
            // each character's advance, and a PROPORTIONAL face wrapped early on
            // every line. technical__898a80c889101e85 (BIZ UDPゴシック 18pt):
            // Word fits 25 chars on the line at advances 13.68..18.00, Oxi fitted
            // 23 at a flat 18.00 and ran one page long.
            //
            // Walk the collection instead of giving up on it.
            let faces: Vec<skrifa::FontRef> = match skrifa::raw::FileRef::new(&data) {
                Ok(skrifa::raw::FileRef::Font(f)) => vec![f],
                Ok(skrifa::raw::FileRef::Collection(c)) => {
                    (0..c.len()).filter_map(|i| c.get(i).ok()).collect()
                }
                Err(_) => continue,
            };
            for font in faces {
                if !face_matches(&font, family, bold, italic) {
                    continue;
                }
                if let Some(m) = metrics_from(&font, family) {
                    if std::env::var("OXI_DBG_FONTRT").is_ok() {
                        eprintln!(
                            "[FONTRT] resolved {:?} bold={} italic={} from {}",
                            family,
                            bold,
                            italic,
                            path.display()
                        );
                    }
                    return Some(Box::leak(Box::new(m)));
                }
            }
        }
    }
    None
}

/// The face named by `family`, if this machine has it and the tables do not.
///
/// Returns `None` for a face that is genuinely absent, which is the case S1146
/// exists for -- the caller must keep that fallback.
pub fn resolve(family: &str, bold: bool, italic: bool) -> Option<FontMetricsRef<'static>> {
    if let Some(font) = memory_font(family, bold, italic) { return Some(font.metrics.into()); }
    if std::env::var("OXI_S1171_DISABLE").is_ok() {
        return None;
    }
    let key = (family.to_ascii_lowercase(), bold, italic);
    if let Some(hit) = cache().read().ok().and_then(|c| c.get(&key).copied()) {
        return hit.map(FontMetricsRef::from);
    }
    let found = load_from_disk(family, bold, italic);
    if let Ok(mut c) = cache().write() {
        c.insert(key, found);
    }
    found.map(FontMetricsRef::from)
}

/// The raw font bytes and the face index within them for `family`, if this
/// machine has it. The shaper (rustybuzz) needs the file itself, not the parsed
/// metrics `resolve` returns. Only the matching file is leaked; misses are
/// cached so a repeated lookup is one directory walk per process.
pub(crate) fn font_file_for(
    family: &str,
    bold: bool,
    italic: bool,
) -> Option<(&'static [u8], u32)> {
    static FCACHE: OnceLock<RwLock<HashMap<(String, bool, bool), Option<(&'static [u8], u32)>>>> =
        OnceLock::new();
    let cache = FCACHE.get_or_init(|| RwLock::new(HashMap::new()));
    let key = (family.to_ascii_lowercase(), bold, italic);
    if let Some(hit) = cache.read().ok().and_then(|c| c.get(&key).copied()) {
        return hit;
    }
    let mut found: Option<(&'static [u8], u32)> = None;
    'outer: for root in search_roots() {
        for path in candidate_files(&root) {
            let Ok(data) = std::fs::read(&path) else {
                continue;
            };
            // Match without leaking: find the face index this file offers, if any.
            let matched: Option<u32> = match skrifa::raw::FileRef::new(&data) {
                Ok(skrifa::raw::FileRef::Font(f)) => {
                    face_matches(&f, family, bold, italic).then_some(0)
                }
                Ok(skrifa::raw::FileRef::Collection(c)) => (0..c.len()).find(|&i| {
                    c.get(i)
                        .map(|f| face_matches(&f, family, bold, italic))
                        .unwrap_or(false)
                }),
                Err(_) => None,
            };
            if let Some(idx) = matched {
                let leaked: &'static [u8] = Box::leak(data.into_boxed_slice());
                found = Some((leaked, idx));
                break 'outer;
            }
        }
    }
    if let Ok(mut c) = cache.write() {
        c.insert(key, found);
    }
    found
}

#[cfg(test)]
mod tests {
    /// The resolver must never panic or invent a face: on a machine without the
    /// font it returns None and the caller keeps its own fallback. Kept
    /// assertion-light on purpose -- WHICH faces are installed is a property of
    /// the machine, not of the code, so asserting on one would fail on CI.
    #[test]
    fn absent_face_resolves_to_none_not_a_substitute() {
        assert!(super::resolve("Zzquartz Nonexistent Face", false, false).is_none());
    }

    /// Whatever a machine does have, a resolved face must carry real metrics.
    #[test]
    fn resolved_faces_are_self_consistent() {
        for root in super::search_roots() {
            for path in super::candidate_files(&root).into_iter().take(3) {
                let Ok(data) = std::fs::read(&path) else { continue };
                let Ok(font) = skrifa::FontRef::new(&data) else { continue };
                if let Some(m) = super::metrics_from(&font, "probe") {
                    assert!(m.units_per_em > 0);
                    assert!(m.ascent > 0.0, "{} has no ascent", path.display());
                    assert!(!m.char_widths.is_empty());
                }
            }
        }
    }
}

// This Source Code Form is subject to the terms of the Mozilla Public
// License, v. 2.0. If a copy of the MPL was not distributed with this
// file, You can obtain one at https://mozilla.org/MPL/2.0/.

// A successful registration or clear invalidates shaping faces on every
// thread. This counter contains no programme data and is independent of any
// caller's UI revision counter. Wraparound is handled by equality checking.
static MEMORY_FONT_GENERATION: std::sync::atomic::AtomicUsize =
    std::sync::atomic::AtomicUsize::new(0);

pub(crate) fn memory_font_generation() -> usize {
    MEMORY_FONT_GENERATION.load(std::sync::atomic::Ordering::Acquire)
}

#[derive(Clone)]
struct MemoryFont {
    metrics: std::sync::Arc<FontMetrics>,
    bytes: std::sync::Arc<[u8]>,
    face_index: u32,
    math: Option<std::sync::Arc<super::program_math::ProgramMath>>,
}

fn memory_fonts() -> &'static RwLock<HashMap<(String, bool, bool), MemoryFont>> {
    static FONTS: OnceLock<RwLock<HashMap<(String, bool, bool), MemoryFont>>> = OnceLock::new();
    FONTS.get_or_init(|| RwLock::new(HashMap::new()))
}

fn memory_key(family: &str, bold: bool, italic: bool) -> (String, bool, bool) {
    (family.split_whitespace().collect::<Vec<_>>().join(" ").to_lowercase(), bold, italic)
}

/// Validated font data ready to publish together with its painting program.
/// Preparing has no effect on the process's font resolver.
pub struct PreparedMemoryFont {
    aliases: Vec<String>,
    bold: bool,
    italic: bool,
    metrics: FontMetrics,
    bytes: Vec<u8>,
    face_index: u32,
    math: Option<std::sync::Arc<super::program_math::ProgramMath>>,
}

/// Parse a caller's real font without any filesystem access. Family names and
/// styles come from that face's metadata; an unrelated collection member is
/// never substituted. The metrics parser is shared with the native resolver.
pub fn prepare_memory_font(
    family: &str, bold: bool, italic: bool, bytes: &[u8], face_index: u32,
) -> Result<PreparedMemoryFont, &'static str> {
    let face = rustybuzz::ttf_parser::Face::parse(bytes, face_index)
        .map_err(|_| "Invalid font program or collection member")?;
    if face.is_bold() != bold || face.is_italic() != italic {
        return Err("Font style does not match its metadata");
    }
    let aliases: Vec<String> = face.names().into_iter()
        .filter(|name| name.name_id == 1 || name.name_id == 16)
        .filter_map(|name| name.to_string())
        .map(|name| memory_key(&name, bold, italic).0).collect();
    if !aliases.contains(&memory_key(family, bold, italic).0) {
        return Err("Font family does not match its metadata");
    }
    let font = match skrifa::raw::FileRef::new(bytes).map_err(|_| "Invalid font container")? {
        skrifa::raw::FileRef::Font(font) if face_index == 0 => font,
        skrifa::raw::FileRef::Collection(collection) => collection.get(face_index)
            .map_err(|_| "Invalid font collection member")?,
        _ => return Err("Invalid font face index"),
    };
    let mut metrics = metrics_from(&font, family).ok_or("Font has no usable layout metrics")?;
    // The real cmap determines symbol coverage for this supplied programme.
    // An empty legacy bitmap must not hide a present or absent glyph.
    let coverage_len: usize = super::SYM_RANGES.iter()
        .map(|(first, last)| (last - first + 1) as usize).sum();
    metrics.sym_coverage = vec![0; coverage_len.div_ceil(8)];
    for &character in metrics.char_widths.keys() {
        if let Some(index) = super::sym_range_index(character) {
            metrics.sym_coverage[index >> 3] |= 1 << (index & 7);
        }
    }
    let math = super::program_math::ProgramMath::from_face(&face, bytes, face_index).map(std::sync::Arc::new);
    Ok(PreparedMemoryFont { aliases, bold, italic, metrics, bytes: bytes.to_vec(), face_index, math })
}

impl PreparedMemoryFont {
    /// Publish owned metrics and programme bytes together. Existing lookups
    /// retain their own Arc, so replacing or clearing the registry releases
    /// data after its last caller without invalidating an in-flight lookup.
    pub fn publish(self) {
        let metrics = std::sync::Arc::new(self.metrics);
        let bytes: std::sync::Arc<[u8]> = self.bytes.into();
        let font = MemoryFont { metrics, bytes, face_index: self.face_index, math: self.math };
        let mut fonts = memory_fonts().write().expect("Memory font registry poisoned");
        let replaced: Vec<_> = self.aliases.iter().filter_map(|alias|
            fonts.get(&(alias.clone(), self.bold, self.italic)).map(|old| old.metrics.clone())).collect();
        fonts.retain(|_, old| !replaced.iter().any(|metrics| std::sync::Arc::ptr_eq(metrics, &old.metrics)));
        for alias in self.aliases { fonts.insert((alias, self.bold, self.italic), font.clone()); }
        MEMORY_FONT_GENERATION.fetch_add(1, std::sync::atomic::Ordering::Release);
    }
}

/// Remove registrations from subsequent lookups. A live metrics or programme
/// handle keeps its own data alive until that caller releases it.
pub fn clear_memory_fonts() {
    let mut fonts = memory_fonts().write().expect("Memory font registry poisoned");
    fonts.clear();
    MEMORY_FONT_GENERATION.fetch_add(1, std::sync::atomic::Ordering::Release);
}

fn memory_font(family: &str, bold: bool, italic: bool) -> Option<MemoryFont> {
    memory_fonts().read().ok()?.get(&memory_key(family, bold, italic)).cloned()
}

/// Explicitly registered face only. It outranks a shipped metric table for an
/// older programme; an ordinary native lookup still keeps its calibrated data.
pub fn resolve_registered(family: &str, bold: bool, italic: bool) -> Option<FontMetricsRef<'static>> {
    memory_font(family, bold, italic).map(|font| font.metrics.into())
}

pub(crate) fn registered_font_file_for(family: &str, bold: bool, italic: bool)
    -> Option<(std::sync::Arc<[u8]>, u32)> {
    memory_font(family, bold, italic).map(|font| (font.bytes, font.face_index))
}

pub fn has_registered_family(family: &str) -> bool {
    [(false, false), (true, false), (false, true), (true, true)]
        .into_iter().any(|(bold, italic)| memory_font(family, bold, italic).is_some())
}


pub(crate) fn registered_math(family: &str) -> Option<std::sync::Arc<super::program_math::ProgramMath>> {
    memory_font(family, false, false)?.math
}

pub(crate) fn registered_glyph(family: &str, bold: bool, italic: bool, c: char, level: u8)
    -> Option<super::math_script_glyphs::ScriptGlyph> {
    let font = memory_font(family, bold, italic)?;
    super::program_math::glyph(&font.bytes, font.face_index, c, level)
}

pub(crate) fn registered_corner_kern(family: &str, bold: bool, italic: bool, gid: u16,
    corner: super::math_kern::Corner, height: f32, size: f32) -> Option<f32> {
    let font = memory_font(family, bold, italic)?;
    Some(super::program_math::corner_kern(&font.bytes, font.face_index, gid, corner, height, size))
}

pub(crate) fn registered_has_math(family: &str, bold: bool, italic: bool) -> bool {
    memory_font(family, bold, italic).and_then(|font|
        rustybuzz::ttf_parser::Face::parse(&font.bytes, font.face_index).ok()
            .map(|face| face.tables().math.is_some())).unwrap_or(false)
}

pub(crate) fn registered_rule_thickness(family: &str, bold: bool, italic: bool) -> Option<f32> {
    let font = memory_font(family, bold, italic)?;
    let face = rustybuzz::ttf_parser::Face::parse(&font.bytes, font.face_index).ok()?;
    let thickness = face.underline_metrics()?.thickness;
    (thickness > 0).then(|| f32::from(thickness)/f32::from(face.units_per_em()))
}
