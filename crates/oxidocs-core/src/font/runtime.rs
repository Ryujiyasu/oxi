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

use super::FontMetrics;

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
fn metrics_from(font: &skrifa::FontRef, family: &str) -> Option<FontMetrics> {
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
pub fn resolve(family: &str, bold: bool, italic: bool) -> Option<&'static FontMetrics> {
    if std::env::var("OXI_S1171_DISABLE").is_ok() {
        return None;
    }
    let key = (family.to_ascii_lowercase(), bold, italic);
    if let Some(hit) = cache().read().ok().and_then(|c| c.get(&key).copied()) {
        return hit;
    }
    let found = load_from_disk(family, bold, italic);
    if let Ok(mut c) = cache().write() {
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
