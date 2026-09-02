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
        if depth > 3 {
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
                self.faces
                    .entry((face.family.to_lowercase(), face.bold, face.italic))
                    .or_insert(face);
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
    let (family, subfamily) = read_name_table(data.get(name_off..end.min(data.len()))?)?;

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
        bold,
        italic,
    })
}

/// Family (nameID 1) and subfamily (nameID 2). Windows and Unicode records
/// are UTF-16BE; Mac Roman records are read as Latin-1, which is correct for
/// the ASCII family names this index is asked about.
fn read_name_table(table: &[u8]) -> Option<(String, String)> {
    let count = be16(table, 2)? as usize;
    let storage = be16(table, 4)? as usize;
    let mut family: Option<String> = None;
    let mut subfamily: Option<String> = None;
    for i in 0..count.min(4096) {
        let rec = 6 + i * 12;
        let platform = be16(table, rec)?;
        let name_id = be16(table, rec + 6)?;
        let length = be16(table, rec + 8)? as usize;
        let offset = be16(table, rec + 10)? as usize;
        if name_id != 1 && name_id != 2 {
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
        // Prefer the Windows record when a face carries both.
        if name_id == 1 && (family.is_none() || platform == 3) {
            family = Some(text);
        } else if name_id == 2 && (subfamily.is_none() || platform == 3) {
            subfamily = Some(text);
        }
    }
    Some((family?, subfamily.unwrap_or_else(|| "Regular".to_string())))
}
