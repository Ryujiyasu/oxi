// This Source Code Form is subject to the terms of the Mozilla Public
// License, v. 2.0. If a copy of the MPL was not distributed with this
// file, You can obtain one at https://mozilla.org/MPL/2.0/.

//! The faces Oxi ships, as bytes.
//!
//! Two things need them and neither can hold them. The browser build embeds
//! them, because a page cannot read the machine's font files and a PDF written
//! there can only carry what travelled with the code. The command-line tool
//! reads them off disk, because a rendered page has to be drawn with the same
//! outlines the layout was measured against. Before this crate existed the
//! browser build reached across the tree into the tool's own folder with
//! `include_bytes!("../../oxidocs-cli/fonts/...")` — which works in a checkout
//! and cannot work anywhere else, because a published crate carries only its
//! own directory.
//!
//! ## Why three crates
//!
//! The files are in `oxidocs-fonts-latin` and `oxidocs-fonts-jp`, and this
//! crate holds no font at all — only the code that puts them back together.
//! The reason is arithmetic: packed, the Latin faces are 3.5 MiB and the
//! Japanese pair is 7.5, and crates.io takes 10 MiB per crate. Together they
//! are 11.0 and cannot be published; apart they both fit with room to spare.
//! Depend on this one and the split does not concern you.
//!
//! Anyone who wants only one half can name that half directly and carry less.

pub use oxidocs_fonts_jp::{OXI_GOTHIC, OXI_GOTHIC_P, OXI_MINCHO, OXI_MINCHO_P};

mod symbol_map;
pub use symbol_map::symbol_font_to_unicode;

/// Every face, under the key the renderer asks for it by and the file name it
/// is stored under: the Latin ones first, then the Japanese pair.
pub fn faces() -> impl Iterator<Item = &'static (&'static str, &'static str, &'static [u8])> {
    oxidocs_fonts_latin::FACES
        .iter()
        .chain(oxidocs_fonts_jp::FACES.iter())
}

/// One face by key, or nothing for a key this does not ship.
pub fn face(key: &str) -> Option<&'static [u8]> {
    faces().find(|(name, ..)| *name == key).map(|(.., data)| *data)
}

/// Just the Latin faces, in the shape the PDF writer wants: a key and its
/// bytes.
pub fn latin_faces() -> Vec<(&'static str, &'static [u8])> {
    oxidocs_fonts_latin::FACES
        .iter()
        .map(|(key, _, data)| (*key, *data))
        .collect()
}

/// The shipped Japanese face that stands in for an East Asian `family` when the
/// machine does not have it: `OxiMincho` for a Mincho (serif) family,
/// `OxiGothic` for everything else. Answered with the key and the file name, so
/// a caller can look the face up in bytes or on disk.
///
/// The command-line tool and the browser build both ask here. Each used to
/// decide on its own, and the tool's answer was "whatever sans face the system
/// has" — so a document set in ＭＳ 明朝 came out of `docx-to-pdf` in Gothic
/// while the browser drew it in Mincho.
pub fn cjk_stand_in(family: &str) -> (&'static str, &'static str) {
    let lower = family.to_lowercase();
    if family.contains("明朝") || lower.contains("mincho") || lower.contains("serif") {
        ("OxiMincho", "OxiMincho.ttf")
    } else {
        ("OxiGothic", "OxiGothic.ttf")
    }
}

/// The companion of `cjk_stand_in` that draws what a proportional family sets
/// proportionally: `OxiMinchoP` beside `OxiMincho`, `OxiGothicP` beside
/// `OxiGothic`. The main faces set printable ASCII at half an em and kana and
/// CJK punctuation at a full em, as ＭＳ ゴシック / ＭＳ 明朝 do, so 游ゴシック,
/// メイリオ or ＭＳ Ｐゴシック drawn with them alone came out in a typewriter
/// face with a gap after every word, and ＭＳ Ｐゴシック's 「（」 and 「・」 a full
/// em wide in a cell measured half that. The companion carries Noto's own
/// proportional Latin and its 'palt' kana and punctuation (see
/// tools/fonts/make_proportional_companion.py). Which characters a family
/// takes from it is the caller's to say (the layout's measured widths know).
pub fn cjk_proportional_stand_in(family: &str) -> (&'static str, &'static str) {
    match cjk_stand_in(family).0 {
        "OxiMincho" => ("OxiMinchoP", "OxiMinchoP.ttf"),
        _ => ("OxiGothicP", "OxiGothicP.ttf"),
    }
}

/// Whether `c` is one the face `cjk_proportional_stand_in` names can draw:
/// printable ASCII and the space, U+3000-30FF (CJK punctuation, kana) and
/// U+FF01-FF60 (the full-width forms).
pub fn in_cjk_proportional_stand_in(c: char) -> bool {
    matches!(c, ' '..='~' | '\u{3000}'..='\u{30FF}' | '\u{FF01}'..='\u{FF60}')
}

/// The folders holding these faces on disk, for a binary still running inside
/// the source tree it was built in. Empty once it has left that tree, which is
/// the point: an absolute path baked in at compile time means nothing on
/// somebody else's machine.
pub fn source_dirs(inside: &std::path::Path) -> Vec<std::path::PathBuf> {
    [
        oxidocs_fonts_latin::source_dir(inside),
        oxidocs_fonts_jp::source_dir(inside),
    ]
    .into_iter()
    .flatten()
    .collect()
}

#[cfg(test)]
mod tests {
    use super::*;

    /// Every face is a TrueType file with something in it. A build that picked
    /// up an empty or truncated font would otherwise fail much later, in a
    /// renderer, as a page of blank paper.
    #[test]
    fn every_face_is_a_font() {
        let mut seen = 0;
        for (name, file, data) in faces() {
            assert!(data.len() > 10_000, "{name} is too small to be a font");
            assert!(file.ends_with(".ttf"), "{name} is stored as {file}");
            // `0x00010000` for TrueType outlines, `OTTO` for CFF.
            let tag = &data[..4];
            assert!(
                tag == [0x00, 0x01, 0x00, 0x00] || tag == b"OTTO" || tag == b"true",
                "{name} does not start like a font: {tag:?}",
            );
            seen += 1;
        }
        assert_eq!(seen, 24, "a face went missing from one of the two halves");
    }

    /// A Mincho family is stood in for by the serif face, in either script and
    /// either width; every other East Asian family by the sans one. Both
    /// answers name a face this crate actually ships.
    #[test]
    fn mincho_families_get_the_serif_face() {
        for family in ["ＭＳ 明朝", "ＭＳ Ｐ明朝", "MS Mincho", "Yu Mincho", "游明朝", "Noto Serif CJK JP"] {
            assert_eq!(cjk_stand_in(family).0, "OxiMincho", "{family}");
        }
        for family in ["ＭＳ ゴシック", "MS PGothic", "游ゴシック", "メイリオ", "Meiryo"] {
            assert_eq!(cjk_stand_in(family).0, "OxiGothic", "{family}");
        }
        for family in ["ＭＳ 明朝", "メイリオ"] {
            let (key, file) = cjk_stand_in(family);
            assert!(face(key).is_some(), "{key} is not shipped");
            assert!(faces().any(|(k, f, _)| *k == key && *f == file), "{key} is not stored as {file}");
        }
    }

    /// The proportional companion follows the main face's kind and is
    /// shipped, with Noto's own advances rather than the main face's half em.
    #[test]
    fn proportional_companion_follows_the_main_face() {
        assert_eq!(cjk_proportional_stand_in("游明朝").0, "OxiMinchoP");
        assert_eq!(cjk_proportional_stand_in("メイリオ").0, "OxiGothicP");
        for family in ["游明朝", "メイリオ"] {
            let (key, file) = cjk_proportional_stand_in(family);
            assert!(faces().any(|(k, f, _)| *k == key && *f == file), "{key} is not stored as {file}");
        }
        assert!(in_cjk_proportional_stand_in('W') && in_cjk_proportional_stand_in(' ') && in_cjk_proportional_stand_in('~'));
        assert!(in_cjk_proportional_stand_in('あ') && in_cjk_proportional_stand_in('（') && in_cjk_proportional_stand_in('Ａ'));
        assert!(!in_cjk_proportional_stand_in('漢') && !in_cjk_proportional_stand_in('\u{A0}') && !in_cjk_proportional_stand_in('\t'));
        assert!(!in_cjk_proportional_stand_in('ｱ'), "the half-width forms are half an em in the main face already");
    }

    /// No key is claimed by both halves, or `face` would answer with whichever
    /// happened to come first.
    #[test]
    fn the_halves_do_not_overlap() {
        let mut keys: Vec<&str> = faces().map(|(key, ..)| *key).collect();
        let before = keys.len();
        keys.sort_unstable();
        keys.dedup();
        assert_eq!(keys.len(), before, "two faces answer to the same key");
    }

    /// Every shipped face is named in LICENSE-FONTS, the notice that travels
    /// with the binary and the npm package: the faces are OFL-1.1, not the
    /// code's licence, and a face added without its notice would ship
    /// without its terms.
    #[test]
    fn every_shipped_face_has_its_notice() {
        let notice = std::fs::read_to_string(
            std::path::Path::new(env!("CARGO_MANIFEST_DIR")).join("../../LICENSE-FONTS"),
        )
        .expect("LICENSE-FONTS at the repository root");
        assert!(notice.contains("SIL OPEN FONT LICENSE Version 1.1"));
        for (_, file, _) in faces() {
            // "LiberationSans-Bold.ttf" is listed as "Liberation*-*.ttf",
            // "Carlito-Regular.ttf" as "Carlito-*.ttf".
            let family = file.split(['-', '.']).next().unwrap();
            let listed = notice.contains(&format!("({file})"))
                || notice.contains(&format!("({family}-*.ttf)"))
                || (family.starts_with("Liberation") && notice.contains("(Liberation*-*.ttf)"));
            assert!(listed, "{file} is shipped but not named in LICENSE-FONTS");
        }
        // The Python package ships its own copy beside the binary it bundles.
        let copy = std::path::Path::new(env!("CARGO_MANIFEST_DIR"))
            .join("../oxidocs-python/python/oxidocs/LICENSE-FONTS");
        assert_eq!(
            std::fs::read_to_string(copy).expect("the Python package's LICENSE-FONTS"),
            notice,
            "crates/oxidocs-python/python/oxidocs/LICENSE-FONTS is not a copy of LICENSE-FONTS",
        );
    }
}
