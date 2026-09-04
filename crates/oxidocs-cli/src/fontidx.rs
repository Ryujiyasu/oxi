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
    /// `OS/2.usWeightClass`, 400 when the face carries no `OS/2` table. Two
    /// faces of one family routinely land in the same style slot, and this is
    /// what decides between them.
    pub weight: u16,
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
        // The usual degradations first, then any style at all. A family can
        // exist in one style only -- Arial Black is bold by weight and has no
        // upright sibling -- and answering ABSENT for it would send a family
        // the machine actually holds to a viewer substitute.
        let preferred = [(bold, italic), (bold, false), (false, italic), (false, false)];
        let last_resort = [(true, false), (false, true), (true, true)];
        for (b, i) in preferred.into_iter().chain(last_resort) {
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
    /// followed by every platform font root. Bundled substitutes are only
    /// reached when the real family is missing -- callers pass the substitute
    /// under its own family name, not under the name it stands in for -- so
    /// the two never contend for one key. Where faces of one family do
    /// contend, `displaces` decides, not the order the directory came back in.
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
                    let key = (alias.to_lowercase(), face.bold, face.italic);
                    let keep = match self.faces.get(&key) {
                        Some(held) => displaces(face.bold, held.weight, face.weight),
                        None => true,
                    };
                    if keep {
                        self.faces.insert(key, face.clone());
                    }
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
    let mut os2 = None;
    for i in 0..num_tables.min(512) {
        let rec = table_dir + 12 + i * 16;
        let tag = data.get(rec..rec + 4)?;
        let offset = be32(data, rec + 8)? as usize;
        let length = be32(data, rec + 12)? as usize;
        match tag {
            b"name" => name = Some((offset, length)),
            b"head" => head = Some(offset),
            b"OS/2" => os2 = Some(offset),
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
    // OS/2.usWeightClass sits four bytes into the table, after version and
    // xAvgCharWidth. Apple's Japanese families leave macStyle clear on every
    // weight and spell the weight only in the subfamily -- W0 through W9 --
    // so the number is the only thing that separates Light from Heavy.
    let weight = os2.and_then(|o| be16(data, o + 4)).unwrap_or(400);
    let sub = subfamily.to_lowercase();
    let bold = mac_style & 0x1 != 0 || sub.contains("bold") || weight >= 600;
    let italic = mac_style & 0x2 != 0 || sub.contains("italic") || sub.contains("oblique");

    Some(Face {
        path: path.to_path_buf(),
        index,
        family,
        aliases: families,
        weight,
        bold,
        italic,
    })
}

/// The weight a style slot is nominally asking for.
fn target_weight(bold: bool) -> u16 {
    if bold {
        700
    } else {
        400
    }
}

/// Whether a candidate face should take a slot from the one already in it.
///
/// Apple ships Hiragino Sans as ten separate collections, W0 through W9, and
/// every one of them reports the same typographic family with macStyle clear.
/// They all land in the same slot, so keeping whichever arrived first made the
/// choice depend on the order `read_dir` happened to return -- Heavy on this
/// machine, Light on the next, from the same document. Nearest to the slot's
/// nominal weight is the same answer everywhere.
fn displaces(bold: bool, held: u16, candidate: u16) -> bool {
    let target = i32::from(target_weight(bold));
    (i32::from(candidate) - target).abs() < (i32::from(held) - target).abs()
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

    /// An sfnt carrying just the three tables the index reads: `name`, `head`
    /// for macStyle, and `OS/2` for usWeightClass.
    fn synth_face(records: &[(u16, u16, u16, &str)], mac_style: u16, weight: u16) -> Vec<u8> {
        let name = name_table(records);
        let mut head = vec![0u8; 54];
        head[44..46].copy_from_slice(&mac_style.to_be_bytes());
        let mut os2 = vec![0u8; 78];
        os2[4..6].copy_from_slice(&weight.to_be_bytes());

        let tables: [(&[u8; 4], &Vec<u8>); 3] = [(b"OS/2", &os2), (b"head", &head), (b"name", &name)];
        let mut out = Vec::new();
        out.extend_from_slice(&0x0001_0000u32.to_be_bytes()); // sfntVersion
        out.extend_from_slice(&(tables.len() as u16).to_be_bytes());
        out.extend_from_slice(&[0u8; 6]); // searchRange, entrySelector, rangeShift
        let mut offset = 12 + tables.len() * 16;
        let mut body = Vec::new();
        for (tag, data) in tables {
            out.extend_from_slice(tag);
            out.extend_from_slice(&0u32.to_be_bytes()); // checksum
            out.extend_from_slice(&(offset as u32).to_be_bytes());
            out.extend_from_slice(&(data.len() as u32).to_be_bytes());
            offset += data.len();
            body.extend_from_slice(data);
        }
        out.extend_from_slice(&body);
        out
    }

    fn hiragino(weight: u16) -> Vec<u8> {
        let sub = format!("W{}", weight / 100);
        synth_face(
            &[
                (3, 1033, 1, "Hiragino Sans"),
                (3, 1033, 16, "Hiragino Sans"),
                (3, 1033, 2, &sub),
            ],
            0,
            weight,
        )
    }

    /// Apple leaves macStyle clear on all ten weights of Hiragino Sans, so the
    /// number in `OS/2` is the only thing separating W0 from W9.
    #[test]
    fn weight_class_is_read_when_mac_style_says_nothing() {
        let faces = read_faces(&hiragino(800), Path::new("W8.ttc"));
        assert_eq!(faces.len(), 1);
        assert_eq!(faces[0].weight, 800);
        assert!(faces[0].bold, "800 is bold however clear macStyle is");

        let light = read_faces(&hiragino(300), Path::new("W3.ttc"));
        assert_eq!(light[0].weight, 300);
        assert!(!light[0].bold, "300 must not land in the bold slot");
    }

    /// The regression this guards: ten files calling themselves `Hiragino
    /// Sans` collided on one key, and whichever the directory listed first
    /// won -- so the same document drew Heavy on one machine and Light on
    /// another. Insert them in the worst order and the answer must not move.
    #[test]
    fn the_slot_goes_to_the_nearest_weight_not_the_first_scanned() {
        for order in [
            vec![800u16, 900, 100, 400, 300, 600, 700],
            vec![400u16, 700, 800, 900, 100, 300, 600],
            vec![900u16, 800, 700, 600, 400, 300, 100],
        ] {
            let mut index = FontIndex::default();
            for weight in &order {
                for face in read_faces(&hiragino(*weight), Path::new("x.ttc")) {
                    for alias in face.aliases.clone() {
                        let key = (alias.to_lowercase(), face.bold, face.italic);
                        let keep = match index.faces.get(&key) {
                            Some(held) => displaces(face.bold, held.weight, face.weight),
                            None => true,
                        };
                        if keep {
                            index.faces.insert(key, face.clone());
                        }
                    }
                }
            }
            let regular = index.find("Hiragino Sans", false, false).expect("regular");
            let bold = index.find("Hiragino Sans", true, false).expect("bold");
            assert_eq!(regular.weight, 400, "regular slot, scanned as {order:?}");
            assert_eq!(bold.weight, 700, "bold slot, scanned as {order:?}");
        }
    }

    #[test]
    fn displaces_prefers_the_nearer_weight_and_keeps_ties() {
        assert!(displaces(false, 800, 400), "400 is nearer regular than 800");
        assert!(!displaces(false, 400, 800));
        assert!(displaces(true, 400, 700), "700 is nearer bold than 400");
        assert!(!displaces(false, 400, 400), "a tie leaves the slot alone");
    }

    /// A family that exists in one style only must still resolve. Arial Black
    /// is 900 with no upright sibling, and answering ABSENT would send a font
    /// the machine holds to a viewer substitute.
    #[test]
    fn a_family_present_in_one_style_only_still_resolves() {
        let mut index = FontIndex::default();
        let data = synth_face(
            &[(3, 1033, 1, "Arial Black"), (3, 1033, 2, "Regular")],
            0,
            900,
        );
        for face in read_faces(&data, Path::new("ArialBlack.ttf")) {
            index
                .faces
                .insert(("arial black".into(), face.bold, face.italic), face);
        }
        assert!(
            index.find("Arial Black", false, false).is_some(),
            "a regular request must reach the only face there is"
        );
    }
}
