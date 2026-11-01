// This Source Code Form is subject to the terms of the Mozilla Public
// License, v. 2.0. If a copy of the MPL was not distributed with this
// file, You can obtain one at https://mozilla.org/MPL/2.0/.

//! GDI-measured character widths: family → ppem → codepoint → width in pixels.
//!
//! These correct for GDI hinting that round(advance * ppem / upm) cannot
//! reproduce. The table is 2.5M entries; `build.rs` packs each (family, ppem)
//! map as its own deflated chunk, and a map is inflated the first time it is
//! asked for and kept for the life of the process. A document touches a few
//! sizes of a few faces, so start-up no longer pays for the rest.

use std::collections::HashMap;
use std::io::Read;
use std::sync::OnceLock;

include!(concat!(env!("OUT_DIR"), "/gdi_widths_index.rs"));

static BLOB: &[u8] = include_bytes!(concat!(env!("OUT_DIR"), "/gdi_widths.bin"));

/// family → ppem → chunk number.
fn index() -> &'static HashMap<&'static str, HashMap<u32, usize>> {
    static INDEX: OnceLock<HashMap<&'static str, HashMap<u32, usize>>> = OnceLock::new();
    INDEX.get_or_init(|| {
        let mut out: HashMap<&'static str, HashMap<u32, usize>> = HashMap::new();
        for (i, &(family, ppem, ..)) in GDI_CHUNKS.iter().enumerate() {
            out.entry(GDI_FAMILIES[usize::from(family)]).or_default().insert(ppem, i);
        }
        out
    })
}

/// One slot per chunk, filled the first time the chunk is asked for.
type Slots = Box<[OnceLock<HashMap<u32, u32>>]>;

fn decoded() -> &'static [OnceLock<HashMap<u32, u32>>] {
    static DECODED: OnceLock<Slots> = OnceLock::new();
    DECODED.get_or_init(|| (0..GDI_CHUNKS.len()).map(|_| OnceLock::new()).collect())
}

fn decode(chunk: usize) -> HashMap<u32, u32> {
    let (_, _, offset, len, count) = GDI_CHUNKS[chunk];
    let packed = &BLOB[offset as usize..(offset + len) as usize];
    let mut raw = Vec::new();
    flate2::read::DeflateDecoder::new(packed)
        .read_to_end(&mut raw)
        .expect("gdi width chunk: inflate failed");
    let mut out = HashMap::with_capacity(count as usize);
    let mut bytes = raw.into_iter();
    let mut cp = 0u32;
    while let Some(delta) = varint(&mut bytes) {
        cp += delta;
        let width = varint(&mut bytes).expect("gdi width chunk: truncated");
        out.insert(cp, width);
    }
    out
}

fn varint(bytes: &mut impl Iterator<Item = u8>) -> Option<u32> {
    let mut v = 0u32;
    let mut shift = 0;
    loop {
        let b = bytes.next()?;
        v |= u32::from(b & 0x7F) << shift;
        if b & 0x80 == 0 {
            return Some(v);
        }
        shift += 7;
    }
}

/// Whether the table measures `family` at any size.
pub(crate) fn has_family(family: &str) -> bool {
    index().contains_key(family)
}

/// The widths of `family` at `ppem`, codepoint → pixels.
pub(crate) fn widths(family: &str, ppem: u32) -> Option<&'static HashMap<u32, u32>> {
    let &chunk = index().get(family)?.get(&ppem)?;
    Some(decoded()[chunk].get_or_init(|| decode(chunk)))
}

#[cfg(test)]
mod tests {
    use super::*;

    /// The packed table holds exactly what the JSON did: every family, every
    /// size, every entry, every width.
    #[test]
    fn packed_table_matches_the_source_json() {
        let gz: &[u8] = include_bytes!("data/gdi_width_overrides.json.gz");
        let mut json = String::new();
        flate2::read::GzDecoder::new(gz).read_to_string(&mut json).unwrap();
        let raw: HashMap<String, HashMap<String, HashMap<String, u32>>> =
            serde_json::from_str(&json).unwrap();
        let mut entries = 0;
        for (family, ppems) in &raw {
            assert!(has_family(family), "{family}");
            for (ppem, chars) in ppems {
                let Ok(ppem) = ppem.parse::<u32>() else { continue };
                let map = widths(family, ppem).unwrap_or_else(|| panic!("{family} {ppem}"));
                let expected: HashMap<u32, u32> =
                    chars.iter().filter_map(|(c, w)| Some((c.parse().ok()?, *w))).collect();
                assert_eq!(map, &expected, "{family} {ppem}");
                entries += expected.len();
            }
        }
        assert_eq!(entries, GDI_CHUNKS.iter().map(|c| c.4 as usize).sum::<usize>());
        assert!(widths("No Such Face", 16).is_none());
        assert!(!has_family("No Such Face"));
    }
}
