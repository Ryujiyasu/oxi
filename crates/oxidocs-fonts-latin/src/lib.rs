// This Source Code Form is subject to the terms of the Mozilla Public
// License, v. 2.0. If a copy of the MPL was not distributed with this
// file, You can obtain one at https://mozilla.org/MPL/2.0/.

//! The Latin faces Oxi ships, as bytes.
//!
//! These are the free faces whose advance widths match their Microsoft
//! counterparts, measured rather than assumed (see `FONT_CHECKS.md` beside
//! `oxidocs-fonts`): Carlito for Calibri, and the three Liberation faces for
//! Arial, Times New Roman and Courier New — identical on all 95 printable
//! ASCII characters, so substituting one moves no glyph. Caladea stands in for
//! Cambria on shape alone; its widths differ by up to 0.19 em, so a Cambria
//! document reflows within the line.
//!
//! No Microsoft font is here, and none ever will be. What ships is what may
//! ship; the licences travel in `fonts/`.
//!
//! Held apart from the Japanese pair for one reason: together they are 11.0
//! MiB packed and crates.io takes 10.

/// Every face under the key the renderer asks for it by — the family name,
/// then `-B`, `-I` or `-BI` for the weight and slope — with the file name it
/// is stored under for a caller that wants it on disk.
pub static FACES: &[(&str, &str, &[u8])] = &[
    ("Carlito", "Carlito-Regular.ttf", include_bytes!("../fonts/Carlito-Regular.ttf")),
    ("Carlito-B", "Carlito-Bold.ttf", include_bytes!("../fonts/Carlito-Bold.ttf")),
    ("Carlito-I", "Carlito-Italic.ttf", include_bytes!("../fonts/Carlito-Italic.ttf")),
    ("Carlito-BI", "Carlito-BoldItalic.ttf", include_bytes!("../fonts/Carlito-BoldItalic.ttf")),
    ("LiberationSans", "LiberationSans-Regular.ttf",
     include_bytes!("../fonts/LiberationSans-Regular.ttf")),
    ("LiberationSans-B", "LiberationSans-Bold.ttf",
     include_bytes!("../fonts/LiberationSans-Bold.ttf")),
    ("LiberationSans-I", "LiberationSans-Italic.ttf",
     include_bytes!("../fonts/LiberationSans-Italic.ttf")),
    ("LiberationSans-BI", "LiberationSans-BoldItalic.ttf",
     include_bytes!("../fonts/LiberationSans-BoldItalic.ttf")),
    ("LiberationSerif", "LiberationSerif-Regular.ttf",
     include_bytes!("../fonts/LiberationSerif-Regular.ttf")),
    ("LiberationSerif-B", "LiberationSerif-Bold.ttf",
     include_bytes!("../fonts/LiberationSerif-Bold.ttf")),
    ("LiberationSerif-I", "LiberationSerif-Italic.ttf",
     include_bytes!("../fonts/LiberationSerif-Italic.ttf")),
    ("LiberationSerif-BI", "LiberationSerif-BoldItalic.ttf",
     include_bytes!("../fonts/LiberationSerif-BoldItalic.ttf")),
    ("LiberationMono", "LiberationMono-Regular.ttf",
     include_bytes!("../fonts/LiberationMono-Regular.ttf")),
    ("LiberationMono-B", "LiberationMono-Bold.ttf",
     include_bytes!("../fonts/LiberationMono-Bold.ttf")),
    ("LiberationMono-I", "LiberationMono-Italic.ttf",
     include_bytes!("../fonts/LiberationMono-Italic.ttf")),
    ("LiberationMono-BI", "LiberationMono-BoldItalic.ttf",
     include_bytes!("../fonts/LiberationMono-BoldItalic.ttf")),
    ("Caladea", "Caladea-Regular.ttf", include_bytes!("../fonts/Caladea-Regular.ttf")),
    ("Caladea-B", "Caladea-Bold.ttf", include_bytes!("../fonts/Caladea-Bold.ttf")),
    ("Caladea-I", "Caladea-Italic.ttf", include_bytes!("../fonts/Caladea-Italic.ttf")),
    ("Caladea-BI", "Caladea-BoldItalic.ttf", include_bytes!("../fonts/Caladea-BoldItalic.ttf")),
];

/// Where this crate's own `fonts/` sits, for a binary still running inside the
/// source tree it was built in. `None` once it has left that tree — see the
/// note on the same function in `oxidocs-fonts-jp`.
pub fn source_dir(inside: &std::path::Path) -> Option<std::path::PathBuf> {
    let manifest = std::path::PathBuf::from(env!("CARGO_MANIFEST_DIR"));
    let root = manifest.parent()?.parent()?;
    if inside.starts_with(root) {
        Some(manifest.join("fonts"))
    } else {
        None
    }
}
