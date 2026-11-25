// This Source Code Form is subject to the terms of the Mozilla Public
// License, v. 2.0. If a copy of the MPL was not distributed with this
// file, You can obtain one at https://mozilla.org/MPL/2.0/.

//! The Japanese faces Oxi ships, as bytes.
//!
//! Held apart from the Latin ones for one reason: together they are 11.0 MiB
//! packed and crates.io takes 10. Apart they are 7.5 and 3.5, and both fit.
//! `oxidocs-fonts` puts them back together for anyone who wants both.

/// The Japanese sans face.
pub static OXI_GOTHIC: &[u8] = include_bytes!("../fonts/OxiGothic.ttf");
/// The Japanese serif face.
pub static OXI_MINCHO: &[u8] = include_bytes!("../fonts/OxiMincho.ttf");
/// The sans face's proportional companion: printable ASCII at Noto's own
/// advances, and kana, CJK punctuation and full-width forms at Noto's 'palt'
/// widths. The main face sets ASCII at half an em and kana at a full em, as
/// ＭＳ ゴシック does; a family that sets them proportionally (游ゴシック's and
/// メイリオ's Latin, ＭＳ Ｐゴシック's Latin, kana and 「（・」) takes them from here.
pub static OXI_GOTHIC_P: &[u8] = include_bytes!("../fonts/OxiGothicP.ttf");
/// The serif face's proportional companion, for 游明朝, ＭＳ Ｐ明朝 and the like.
pub static OXI_MINCHO_P: &[u8] = include_bytes!("../fonts/OxiMinchoP.ttf");

/// Each face under the key the renderer asks for it by, with the file name it
/// is stored under for a caller that wants it on disk.
pub static FACES: &[(&str, &str, &[u8])] = &[
    ("OxiGothic", "OxiGothic.ttf", OXI_GOTHIC),
    ("OxiMincho", "OxiMincho.ttf", OXI_MINCHO),
    ("OxiGothicP", "OxiGothicP.ttf", OXI_GOTHIC_P),
    ("OxiMinchoP", "OxiMinchoP.ttf", OXI_MINCHO_P),
];

/// Where this crate's own `fonts/` sits, for a binary still running inside the
/// source tree it was built in.
///
/// `CARGO_MANIFEST_DIR` is an absolute path baked in at compile time, so a
/// binary handed to somebody else would otherwise read the build machine's
/// working directory — a path that by then may hold nothing, or somebody
/// else's fonts. `inside` is the running executable; the answer is `None`
/// unless it is still under the same tree.
pub fn source_dir(inside: &std::path::Path) -> Option<std::path::PathBuf> {
    let manifest = std::path::PathBuf::from(env!("CARGO_MANIFEST_DIR"));
    let root = manifest.parent()?.parent()?;
    if inside.starts_with(root) {
        Some(manifest.join("fonts"))
    } else {
        None
    }
}
