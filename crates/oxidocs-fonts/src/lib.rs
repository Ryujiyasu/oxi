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

pub use oxidocs_fonts_jp::{OXI_GOTHIC, OXI_MINCHO};

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
        assert_eq!(seen, 22, "a face went missing from one of the two halves");
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
}
