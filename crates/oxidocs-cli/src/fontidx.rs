// This Source Code Form is subject to the terms of the Mozilla Public
// License, v. 2.0. If a copy of the MPL was not distributed with this
// file, You can obtain one at https://mozilla.org/MPL/2.0/.

//! Family-name to font-file index, built by scanning the platform's font
//! directories and reading each face's `name` table.
//!
//! The PDF writer used to reach for one hardcoded path (Calibri) and map
//! every sans family onto it, so an Arial document was drawn with Calibri
//! outlines, and a Verdana or Tahoma document fell through to an unembedded
//! base-14 font that the viewer substituted at will. Those families are
//! installed on the machine; nothing was looking for them. This index
//! answers "which file holds <family> at <weight/slope>" for whatever is
//! actually present.
//!
//! Windows keeps user-installed faces and the Office cloud-font cache
//! outside the system font directory, so all three roots are scanned.

use std::collections::HashMap;
use std::path::{Path, PathBuf};

/// One face inside a font file. `index` is the TTC member index, 0 for a
/// plain TTF/OTF.
#[derive(Debug, Clone)]
pub struct Face {
    pub path: PathBuf,
    pub index: u32,
    pub family: String,
    /// Every name this face answers to: the typographic family and the legacy
    /// one, which differ on Apple's system fonts.
    pub aliases: Vec<String>,
    pub bold: bool,
    pub italic: bool,
}

#[derive(Debug, Default)]
pub struct FontIndex {
    /// (lowercased family, bold, italic) -> face
    faces: HashMap<(String, bool, bool), Face>,
}

impl FontIndex {
    /// Exact match on family and style, then the usual degradations: an
    /// italic request falls back to the upright face, a bold request to the
    /// regular one. `None` means the family is absent entirely, which is the
    /// caller's cue to substitute or to warn.
    pub fn find(&self, family: &str, bold: bool, italic: bool) -> Option<&Face> {
        let key = family.trim().to_lowercase();
        for (b, i) in [(bold, italic), (bold, false), (false, italic), (false, false)] {
            if let Some(face) = self.faces.get(&(key.clone(), b, i)) {
                return Some(face);
            }
        }
        None
    }

    pub fn len(&self) -> usize {
        self.faces.len()
    }

    pub fn is_empty(&self) -> bool {
        self.faces.is_empty()
    }

    /// Scan `extra` (the bundled fonts directory beside the executable)
    /// followed by every platform font root. The first insert for a key
    /// wins, so bundled substitutes are only reached when the real family is
    /// missing -- callers pass the substitute under its own family name, not
    /// under the name it stands in for.
    pub fn build(extra: &[PathBuf]) -> FontIndex {
        let mut index = FontIndex::default();
        for dir in extra.iter().cloned().chain(system_font_dirs()) {
            index.scan_dir(&dir, 0);
        }
        index
    }

    fn scan_dir(&mut self, dir: &Path, depth: usize) {
        if depth > 6 {
            return;
        }
        let entries = match std::fs::read_dir(dir) {
            Ok(entries) => entries,
            Err(_) => return,
        };
        for entry in entries.flatten() {
            let path = entry.path();
            if path.is_dir() {
                self.scan_dir(&path, depth + 1);
                continue;
            }
            let ext = path
                .extension()
                .and_then(|e| e.to_str())
                .map(|e| e.to_lowercase())
                .unwrap_or_default();
            if !matches!(ext.as_str(), "ttf" | "otf" | "ttc" | "otc") {
                continue;
            }
            let data = match std::fs::read(&path) {
                Ok(data) => data,
                Err(_) => continue,
            };
            for face in read_faces(&data, &path) {
                for alias in face.aliases.clone() {
                    self.faces
                        .entry((alias.to_lowercase(), face.bold, face.italic))
                        .or_insert_with(|| face.clone());
                }
            }
        }
    }
}

fn system_font_dirs() -> Vec<PathBuf> {
    let mut dirs = Vec::new();
    if cfg!(windows) {
        if let Ok(windir) = std::env::var("WINDIR") {
            dirs.push(PathBuf::from(windir).join("Fonts"));
        }
        if let Ok(local) = std::env::var("LOCALAPPDATA") {
            let local = PathBuf::from(local);
            // User-installed faces, and the Office cloud-font cache whose
            // files are named by id rather than by family.
            dirs.push(local.join("Microsoft").join("Windows").join("Fonts"));
            dirs.push(local.join("Microsoft").join("FontCache"));
        }
    } else if cfg!(target_os = "macos") {
        dirs.push(PathBuf::from("/System/Library/Fonts"));
        dirs.push(PathBuf::from("/Library/Fonts"));
        // Most of the East Asian families a Mac reports are not files under
        // /System/Library/Fonts at all; they are on-demand assets, and a
        // Japanese document is exactly what needs them.
        dirs.push(PathBuf::from(
            "/System/Library/AssetsV2/com_apple_MobileAsset_Font7",
        ));
        // Homebrew: /usr/local on Intel, /opt/homebrew on Apple silicon.
        dirs.push(PathBuf::from("/usr/local/share/fonts"));
        dirs.push(PathBuf::from("/opt/homebrew/share/fonts"));
        if let Ok(home) = std::env::var("HOME") {
            dirs.push(PathBuf::from(home).join("Library").join("Fonts"));
        }
    } else {
        dirs.push(PathBuf::from("/usr/share/fonts"));
        dirs.push(PathBuf::from("/usr/local/share/fonts"));
        if let Ok(home) = std::env::var("HOME") {
            let home = PathBuf::from(home);
            dirs.push(home.join(".fonts"));
            dirs.push(home.join(".local").join("share").join("fonts"));
        }
    }
    dirs
}

fn be16(d: &[u8], o: usize) -> Option<u16> {
    Some(u16::from_be_bytes([*d.get(o)?, *d.get(o + 1)?]))
}

fn be32(d: &[u8], o: usize) -> Option<u32> {
    Some(u32::from_be_bytes([
        *d.get(o)?,
        *d.get(o + 1)?,
        *d.get(o + 2)?,
        *d.get(o + 3)?,
    ]))
}

/// Every face in the file: one for a TTF/OTF, N for a TrueType Collection.
fn read_faces(data: &[u8], path: &Path) -> Vec<Face> {
    let mut out = Vec::new();
    if data.len() < 12 {
        return out;
    }
    if &data[0..4] == b"ttcf" {
        let count = be32(data, 8).unwrap_or(0).min(64);
        for i in 0..count {
            if let Some(offset) = be32(data, 12 + 4 * i as usize) {
                if let Some(face) = read_face_at(data, offset as usize, path, i) {
                    out.push(face);
                }
            }
        }
    } else if let Some(face) = read_face_at(data, 0, path, 0) {
        out.push(face);
    }
    out
}

fn read_face_at(data: &[u8], table_dir: usize, path: &Path, index: u32) -> Option<Face> {
    let num_tables = be16(data, table_dir + 4)? as usize;
    let mut name = None;
    let mut head = None;
    for i in 0..num_tables.min(512) {
        let rec = table_dir + 12 + i * 16;
        let tag = data.get(rec..rec + 4)?;
        let offset = be32(data, rec + 8)? as usize;
        let length = be32(data, rec + 12)? as usize;
        match tag {
            b"name" => name = Some((offset, length)),
            b"head" => head = Some(offset),
            _ => {}
        }
    }
    let (name_off, name_len) = name?;
    let end = name_off.checked_add(name_len)?;
    let (families, subfamily) = read_name_table(data.get(name_off..end.min(data.len()))?)?;
    let family = families.first()?.clone();

    // head.macStyle bit 0 is bold, bit 1 is italic. The subfamily string is
    // the cross-check: some faces leave macStyle clear.
    let mac_style = head.and_then(|o| be16(data, o + 44)).unwrap_or(0);
    let sub = subfamily.to_lowercase();
    let bold = mac_style & 0x1 != 0 || sub.contains("bold");
    let italic = mac_style & 0x2 != 0 || sub.contains("italic") || sub.contains("oblique");

    Some(Face {
        path: path.to_path_buf(),
        index,
        family,
        aliases: families,
        bold,
        italic,
    })
}

/// How much a name record is worth as "the name a document would write".
/// English wins: Apple writes nameID 1 twice on platform 3 — langID 1033 and
/// then 1041 — so taking the last platform-3 record indexed Hiragino Sans
/// under `ヒラギノ角ゴシック W3` and left the English name unreachable.
fn name_rank(platform: u16, language: u16) -> u8 {
    match (platform, language) {
        (3, 1033) => 5, // Windows, English (United States)
        (0, _) => 4,    // Unicode: no language, so never a localized surprise
        (1, 0) => 3,    // Macintosh, English
        (3, _) => 2,    // Windows, some other language
        (1, _) => 1,    // Macintosh, some other language
        _ => 0,
    }
}

/// The names a face can be asked for, plus its subfamily.
///
/// Two nameIDs carry a family: 1 is the legacy family, which on Apple's system
/// fonts has the weight welded on (`Hiragino Sans W3`), and 16 is the
/// typographic family (`Hiragino Sans`), which is what a document usually
/// names. Both are returned so either spelling resolves.
fn read_name_table(table: &[u8]) -> Option<(Vec<String>, String)> {
    let count = be16(table, 2)? as usize;
    let storage = be16(table, 4)? as usize;
    // (nameID) -> (rank, text)
    let mut best: std::collections::HashMap<u16, (u8, String)> =
        std::collections::HashMap::new();
    for i in 0..count.min(4096) {
        let rec = 6 + i * 12;
        let platform = be16(table, rec)?;
        let language = be16(table, rec + 4)?;
        let name_id = be16(table, rec + 6)?;
        let length = be16(table, rec + 8)? as usize;
        let offset = be16(table, rec + 10)? as usize;
        if !matches!(name_id, 1 | 2 | 16 | 17) {
            continue;
        }
        let start = storage.checked_add(offset)?;
        let bytes = match table.get(start..start.checked_add(length)?) {
            Some(bytes) => bytes,
            None => continue,
        };
        let text = if platform == 3 || platform == 0 {
            let units: Vec<u16> = bytes
                .chunks_exact(2)
                .map(|c| u16::from_be_bytes([c[0], c[1]]))
                .collect();
            String::from_utf16_lossy(&units)
        } else {
            bytes.iter().map(|b| *b as char).collect()
        };
        let text = text.trim().to_string();
        if text.is_empty() {
            continue;
        }
        let rank = name_rank(platform, language);
        match best.get(&name_id) {
            Some((seen, _)) if *seen >= rank => {}
            _ => {
                best.insert(name_id, (rank, text));
            }
        }
    }

    let mut families = Vec::new();
    for id in [16u16, 1] {
        if let Some((_, text)) = best.get(&id) {
            if !families.contains(text) {
                families.push(text.clone());
            }
        }
    }
    if families.is_empty() {
        return None;
    }
    // Subfamily: typographic (17) if the face declares one, else legacy (2).
    let subfamily = best
        .get(&17)
        .or_else(|| best.get(&2))
        .map(|(_, text)| text.clone())
        .unwrap_or_else(|| "Regular".to_string());
    Some((families, subfamily))
}

#[cfg(test)]
mod tests {
    use super::*;

    /// Build a `name` table from (platform, language, nameID, text) records.
    fn name_table(records: &[(u16, u16, u16, &str)]) -> Vec<u8> {
        let mut storage: Vec<u8> = Vec::new();
        let mut entries: Vec<u8> = Vec::new();
        for (platform, language, name_id, text) in records {
            let bytes: Vec<u8> = if *platform == 3 || *platform == 0 {
                text.encode_utf16().flat_map(|u| u.to_be_bytes()).collect()
            } else {
                text.bytes().collect()
            };
            let offset = storage.len() as u16;
            entries.extend_from_slice(&platform.to_be_bytes());
            entries.extend_from_slice(&1u16.to_be_bytes()); // encodingID
            entries.extend_from_slice(&language.to_be_bytes());
            entries.extend_from_slice(&name_id.to_be_bytes());
            entries.extend_from_slice(&(bytes.len() as u16).to_be_bytes());
            entries.extend_from_slice(&offset.to_be_bytes());
            storage.extend_from_slice(&bytes);
        }
        let count = records.len() as u16;
        let storage_at = 6 + count * 12;
        let mut out = Vec::new();
        out.extend_from_slice(&0u16.to_be_bytes()); // format
        out.extend_from_slice(&count.to_be_bytes());
        out.extend_from_slice(&storage_at.to_be_bytes());
        out.extend_from_slice(&entries);
        out.extend_from_slice(&storage);
        out
    }

    /// Apple writes nameID 1 twice on platform 3 — English, then Japanese —
    /// and puts the family a document actually names in nameID 16. Taking the
    /// last platform-3 record indexed Hiragino Sans under its Japanese name
    /// and left `Hiragino Sans` unreachable.
    #[test]
    fn localized_name_does_not_displace_the_english_one() {
        let table = name_table(&[
            (1, 0, 1, "Hiragino Sans"),
            (1, 11, 1, "ヒラギノ角ゴシック"),
            (3, 1033, 1, "Hiragino Sans W3"),
            (3, 1041, 1, "ヒラギノ角ゴシック W3"),
            (3, 1033, 16, "Hiragino Sans"),
            (3, 1033, 2, "W3"),
        ]);
        let (families, subfamily) = read_name_table(&table).expect("names");

        assert_eq!(
            families[0], "Hiragino Sans",
            "the typographic family is what a document names"
        );
        assert!(
            families.contains(&"Hiragino Sans W3".to_string()),
            "the legacy family must stay reachable too, got {families:?}"
        );
        assert!(
            !families.iter().any(|f| f.contains('ヒ')),
            "a localized name must not become the family, got {families:?}"
        );
        assert_eq!(subfamily, "W3");
    }

    /// A plain Windows font carries one family on platform 3 and nothing else;
    /// it must come through unchanged.
    #[test]
    fn a_single_family_is_read_as_is() {
        let table = name_table(&[(3, 1033, 1, "Arial"), (3, 1033, 2, "Bold")]);
        let (families, subfamily) = read_name_table(&table).expect("names");
        assert_eq!(families, vec!["Arial".to_string()]);
        assert_eq!(subfamily, "Bold");
    }
}
