// This Source Code Form is subject to the terms of the Mozilla Public
// License, v. 2.0. If a copy of the MPL was not distributed with this
// file, You can obtain one at https://mozilla.org/MPL/2.0/.

//! Resolve caller-requested family/style from font programs already installed
//! on the host. The index holds paths and face indices, never outline tables.
//! Font bytes are read into memory only when returned to the caller.

use std::collections::HashMap;
use std::path::{Path, PathBuf};
use ttf_parser::Face;

#[derive(Clone, Debug, Eq, PartialEq)]
struct FaceLocation {
    path: PathBuf,
    index: u32,
}

pub struct ResolvedFontProgram {
    bytes: Vec<u8>,
    pub face_index: u32,
}

impl ResolvedFontProgram {
    pub fn bytes(&self) -> &[u8] { &self.bytes }
    pub fn into_bytes(self) -> Vec<u8> { self.bytes }
}

#[derive(Debug, PartialEq)]
pub enum ResolveError {
    MissingFace,
    UnreadableProgram,
    ChangedFace,
}

#[derive(Default)]
pub struct InstalledFontPrograms {
    faces: HashMap<(String, bool, bool), FaceLocation>,
    indexed_faces: usize,
}

fn key(name: &str) -> String {
    name.split_whitespace().collect::<Vec<_>>().join(" ").to_lowercase()
}

fn aliases(face: &Face<'_>) -> Vec<String> {
    face.names().into_iter()
        .filter(|name| name.name_id == 1 || name.name_id == 16)
        .filter_map(|name| name.to_string()).map(|name| key(&name)).collect()
}

fn font_file(path: &Path) -> bool {
    matches!(path.extension().and_then(|e| e.to_str()).map(str::to_ascii_lowercase).as_deref(),
        Some("ttf") | Some("otf") | Some("ttc"))
}

impl InstalledFontPrograms {
    /// Earlier supplied files have priority. The caller owns platform policy;
    /// this resolver makes no claim that this order is Word's universal rule.
    pub fn from_files(files: impl IntoIterator<Item = PathBuf>) -> Self {
        let mut registry = Self::default();
        for path in files {
            let Ok(bytes) = std::fs::read(&path) else { continue };
            let count = ttf_parser::fonts_in_collection(&bytes).unwrap_or(1);
            for index in 0..count {
                let Ok(face) = Face::parse(&bytes, index) else { continue };
                let names = aliases(&face);
                if names.is_empty() { continue; }
                let bold = face.is_bold();
                let italic = face.is_italic();
                for name in names {
                    registry.faces.entry((name, bold, italic))
                        .or_insert_with(|| FaceLocation { path: path.clone(), index });
                }
                registry.indexed_faces += 1;
            }
        }
        registry
    }

    /// Root order is preserved; filenames within a root are deterministic.
    /// One directory level covers family folders in the Office cloud cache.
    pub fn from_roots(roots: impl IntoIterator<Item = PathBuf>) -> Self {
        let mut files = Vec::new();
        for root in roots {
            let Ok(entries) = std::fs::read_dir(&root) else { continue };
            let mut group = Vec::new();
            for entry in entries.flatten() {
                let path = entry.path();
                if path.is_dir() {
                    if let Ok(nested) = std::fs::read_dir(path) {
                        group.extend(nested.flatten().map(|entry| entry.path()).filter(|p| font_file(p)));
                    }
                } else if font_file(&path) {
                    group.push(path);
                }
            }
            group.sort();
            files.extend(group);
        }
        Self::from_files(files)
    }

    pub fn indexed_faces(&self) -> usize { self.indexed_faces }

    /// Revalidate the selected face at each read. A changed file must never
    /// silently return an unrelated family/style at the old collection index.
    pub fn resolve(&self, family: &str, bold: bool, italic: bool) -> Result<ResolvedFontProgram, ResolveError> {
        let location = self.faces.get(&(key(family), bold, italic)).ok_or(ResolveError::MissingFace)?;
        let bytes = std::fs::read(&location.path).map_err(|_| ResolveError::UnreadableProgram)?;
        let face = Face::parse(&bytes, location.index).map_err(|_| ResolveError::ChangedFace)?;
        if face.is_bold() != bold || face.is_italic() != italic || !aliases(&face).contains(&key(family)) {
            return Err(ResolveError::ChangedFace);
        }
        Ok(ResolvedFontProgram { bytes, face_index: location.index })
    }
}

/// Match the existing native layout resolver's three host font sources.
/// No family names or proprietary font data are part of this source code.
pub fn windows_font_roots() -> Vec<PathBuf> {
    let mut roots = Vec::new();
    if let Some(local) = std::env::var_os("LOCALAPPDATA") {
        let local = PathBuf::from(local);
        roots.push(local.join("Microsoft").join("FontCache").join("4").join("CloudFonts"));
        roots.push(local.join("Microsoft").join("Windows").join("Fonts"));
    }
    if let Some(windows) = std::env::var_os("SystemRoot") {
        roots.push(PathBuf::from(windows).join("Fonts"));
    }
    roots
}

#[cfg(test)]
mod tests {
    use super::*;

    fn font_path() -> PathBuf {
        PathBuf::from(std::env::var_os("PHASE1_PROTOTYPE_FONT_FILE").expect("Test-only installed TTC input required"))
    }

    #[test]
    fn actual_collection_names_and_face_index_resolve_without_filename_rules() {
        let path = font_path();
        let registry = InstalledFontPrograms::from_files([path.clone()]);
        let math = registry.resolve("Cambria Math", false, false).unwrap();
        let ordinary = registry.resolve("Cambria", false, false).unwrap();
        assert_eq!(math.face_index, 1);
        assert_eq!(ordinary.face_index, 0);
        assert_eq!(math.bytes(), std::fs::read(path).unwrap());
        let face = Face::parse(math.bytes(), math.face_index).unwrap();
        assert_eq!(face.glyph_index('\u{222b}').unwrap().0, 1516);
        assert!(registry.indexed_faces() >= 2);
    }

    #[test]
    fn family_normalization_retains_exact_face_and_rejects_unavailable_style() {
        let registry = InstalledFontPrograms::from_files([font_path()]);
        assert_eq!(registry.resolve("  cAmBrIa   mAtH  ", false, false).unwrap().face_index, 1);
        assert!(matches!(registry.resolve("Cambria Math", true, false), Err(ResolveError::MissingFace)));
        assert!(matches!(registry.resolve("Unknown test family", false, false), Err(ResolveError::MissingFace)));
    }

    #[test]
    fn unreadable_input_is_skipped_without_inventing_a_substitute() {
        let registry = InstalledFontPrograms::from_files([PathBuf::from("__phase1_nonexistent_font_file__"), font_path()]);
        assert_eq!(registry.resolve("Cambria Math", false, false).unwrap().face_index, 1);
        let empty = InstalledFontPrograms::from_roots([PathBuf::from("__phase1_nonexistent_font_root__")]);
        assert_eq!(empty.indexed_faces(), 0);
        assert!(matches!(empty.resolve("Cambria Math", false, false), Err(ResolveError::MissingFace)));
    }
}
