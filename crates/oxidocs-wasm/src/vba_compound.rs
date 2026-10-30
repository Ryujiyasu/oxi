// SPDX-License-Identifier: MIT OR Apache-2.0

//! Repack live compound-file objects without changing their contents or metadata.

use std::collections::{BTreeMap, BTreeSet};
use std::io::{self, Cursor};
use std::path::PathBuf;

const END: u32 = 0xffff_fffe;
const NONE: u32 = 0xffff_ffff;

fn word(bytes: &[u8], offset: usize) -> Result<u32, String> {
    let value = bytes.get(offset..offset.checked_add(4).ok_or("compound offset overflow")?)
        .ok_or("truncated compound field")?;
    Ok(u32::from_le_bytes(value.try_into().unwrap()))
}

fn sector_start(bytes: &[u8], size: usize, id: u32) -> Result<usize, String> {
    let start = (id as usize).checked_add(1).and_then(|id| id.checked_mul(size))
        .ok_or("compound sector offset overflow")?;
    if start.checked_add(size).is_none_or(|end| end > bytes.len()) {
        return Err("compound sector outside file".into());
    }
    Ok(start)
}

/// Physical directory locations indexed by logical paths. Allocation links are
/// rebuilt, while CLSID/state/timestamp bytes can be retained verbatim. Using
/// raw timestamps also retains values a platform's SystemTime cannot represent.
fn directory_entries(bytes: &[u8]) -> Result<(BTreeMap<PathBuf, usize>, Vec<usize>), String> {
    if bytes.get(..8) != Some(&[0xd0, 0xcf, 0x11, 0xe0, 0xa1, 0xb1, 0x1a, 0xe1]) {
        return Err("invalid compound signature".into());
    }
    let shift = bytes.get(30..32).ok_or("truncated compound header")?;
    let size = match u16::from_le_bytes(shift.try_into().unwrap()) {
        9 => 512,
        12 => 4096,
        _ => return Err("invalid compound sector size".into()),
    };
    if bytes.len() < size || bytes.len() % size != 0 {
        return Err("invalid compound file length".into());
    }
    let sectors = bytes.len() / size - 1;
    let fat_count = word(bytes, 44)? as usize;
    let difat_count = word(bytes, 72)? as usize;
    if fat_count > sectors || difat_count > sectors {
        return Err("invalid compound allocation counts".into());
    }
    let mut fat_sectors = Vec::new();
    for at in (76..512).step_by(4) {
        let id = word(bytes, at)?;
        if id != NONE { fat_sectors.push(id); }
    }
    let mut id = word(bytes, 68)?;
    let mut seen = BTreeSet::new();
    for _ in 0..difat_count {
        if !seen.insert(id) { return Err("cyclic compound DIFAT".into()); }
        let start = sector_start(bytes, size, id)?;
        for at in (start..start + size - 4).step_by(4) {
            let id = word(bytes, at)?;
            if id != NONE { fat_sectors.push(id); }
        }
        id = word(bytes, start + size - 4)?;
    }
    if fat_sectors.len() != fat_count || difat_count > 0 && id != END {
        return Err("inconsistent compound DIFAT".into());
    }
    let mut fat = Vec::new();
    for id in fat_sectors {
        let start = sector_start(bytes, size, id)?;
        for at in (start..start + size).step_by(4) { fat.push(word(bytes, at)?); }
    }
    let mut entry_offsets = Vec::new();
    id = word(bytes, 48)?;
    seen.clear();
    while id != END {
        if !seen.insert(id) { return Err("cyclic compound directory".into()); }
        let start = sector_start(bytes, size, id)?;
        entry_offsets.extend((start..start + size).step_by(128));
        id = *fat.get(id as usize).ok_or("invalid compound directory chain")?;
    }
    let mut pending = vec![(0_u32, PathBuf::from("/"))];
    let mut locations = BTreeMap::new();
    seen.clear();
    while let Some((id, parent)) = pending.pop() {
        if id == NONE { continue; }
        if !seen.insert(id) { return Err("cyclic compound directory tree".into()); }
        let at = *entry_offsets.get(id as usize).ok_or("invalid compound directory entry")?;
        let kind = bytes[at + 66];
        let path = if id == 0 {
            if kind != 5 { return Err("missing compound root storage".into()); }
            PathBuf::from("/")
        } else {
            if !matches!(kind, 1 | 2) { return Err("invalid live compound object".into()); }
            let count = u16::from_le_bytes(bytes[at + 64..at + 66].try_into().unwrap()) as usize;
            if !(2..=64).contains(&count) || count % 2 != 0 { return Err("invalid compound name length".into()); }
            let units: Vec<_> = bytes[at..at + count - 2].chunks_exact(2)
                .map(|word| u16::from_le_bytes(word.try_into().unwrap())).collect();
            parent.join(String::from_utf16(&units).map_err(|_| "invalid compound name")?)
        };
        pending.push((word(bytes, at + 68)?, parent.clone()));
        pending.push((word(bytes, at + 72)?, parent));
        if matches!(kind, 1 | 5) { pending.push((word(bytes, at + 76)?, path.clone())); }
        if locations.insert(path, at).is_some() { return Err("duplicate compound path".into()); }
    }
    Ok((locations, entry_offsets))
}

fn directory_locations(bytes: &[u8]) -> Result<BTreeMap<PathBuf, usize>, String> {
    directory_entries(bytes).map(|(locations, _)| locations)
}

fn put(bytes: &mut [u8], at: usize, value: u32) {
    bytes[at..at + 4].copy_from_slice(&value.to_le_bytes());
}

/// Start with the original live directory tree and empty streams. Creating a
/// new compound file or storage through cfb asks SystemTime for the current
/// time, which is unavailable on wasm32-unknown-unknown. Opening this valid
/// empty allocation and writing existing streams needs no platform clock.
fn empty_allocation(bytes: &[u8], locations: &BTreeMap<PathBuf, usize>, offsets: &[usize]) -> Result<Vec<u8>, String> {
    let size = 1_usize << u16::from_le_bytes(bytes[30..32].try_into().unwrap());
    let mut paths: Vec<_> = locations.keys().collect();
    paths.sort_by(|a, b| a.components().count().cmp(&b.components().count()).then_with(|| a.cmp(b)));
    let directory_count = (paths.len() * 128).div_ceil(size);
    let words = size / 4;
    let mut fat_count = 1;
    let mut difat_count = 0;
    loop {
        let count = (directory_count + fat_count + difat_count).div_ceil(words);
        let extended = count.saturating_sub(109).div_ceil(words - 1);
        if count == fat_count && extended == difat_count { break; }
        fat_count = count;
        difat_count = extended;
    }
    let total = directory_count + fat_count + difat_count;
    let count = u32::try_from(total).map_err(|_| "compound allocation too large")?;
    let mut output = vec![0; (total + 1).checked_mul(size).ok_or("compound allocation overflow")?];
    output[..size].copy_from_slice(&bytes[..size]);
    put(&mut output, 40, if size == 4096 { directory_count as u32 } else { 0 });
    put(&mut output, 44, fat_count as u32);
    put(&mut output, 48, 0);
    put(&mut output, 60, END);
    put(&mut output, 64, 0);
    put(&mut output, 68, if difat_count == 0 { END } else { (directory_count + fat_count) as u32 });
    put(&mut output, 72, difat_count as u32);
    for at in (76..512).step_by(4) { put(&mut output, at, NONE); }
    let fat_ids: Vec<_> = (directory_count..directory_count + fat_count).map(|id| id as u32).collect();
    for (index, id) in fat_ids.iter().take(109).enumerate() { put(&mut output, 76 + index * 4, *id); }
    let mut fat = vec![NONE; fat_count * words];
    for (id, link) in fat.iter_mut().take(directory_count).enumerate() {
        *link = if id + 1 == directory_count { END } else { id as u32 + 1 };
    }
    for id in &fat_ids { fat[*id as usize] = 0xffff_fffd; }
    for id in directory_count + fat_count..total { fat[id] = 0xffff_fffc; }
    for (index, id) in fat_ids.iter().enumerate() {
        let start = sector_start(&output, size, *id)?;
        for (at, value) in fat[index * words..(index + 1) * words].iter().enumerate() {
            put(&mut output, start + at * 4, *value);
        }
    }
    for index in 0..difat_count {
        let id = (directory_count + fat_count + index) as u32;
        let start = sector_start(&output, size, id)?;
        for at in 0..words - 1 {
            let value = fat_ids.get(109 + index * (words - 1) + at).copied().unwrap_or(NONE);
            put(&mut output, start + at * 4, value);
        }
        put(&mut output, start + size - 4, if id + 1 == count { END } else { id + 1 });
    }
    let ids: BTreeMap<_, _> = paths.iter().enumerate().map(|(id, path)| (locations[*path], id as u32)).collect();
    for index in 0..directory_count * size / 128 {
        let at = size + index * 128;
        output[at + 67] = 1;
        for field in [68, 72, 76] { put(&mut output, at + field, NONE); }
        put(&mut output, at + 116, END);
    }
    for (index, path) in paths.into_iter().enumerate() {
        let old = locations[path];
        let new = size + index * 128;
        output[new..new + 128].copy_from_slice(&bytes[old..old + 128]);
        for field in [68, 72, 76] {
            let link = word(bytes, old + field)?;
            let link = if link == NONE { NONE } else {
                let offset = offsets.get(link as usize).ok_or("invalid compound object link")?;
                *ids.get(offset).ok_or("unreachable compound object link")?
            };
            put(&mut output, new + field, link);
        }
        put(&mut output, new + 116, if bytes[old + 66] == 1 { 0 } else { END });
        output[new + 120..new + 128].fill(0);
    }
    Ok(output)
}

/// Remove unused sectors while retaining every reachable storage/stream,
/// including compiler caches and unknown application-owned objects.
pub(crate) fn compact(bytes: &[u8]) -> Result<Vec<u8>, String> {
    let mut original = cfb::CompoundFile::open(Cursor::new(bytes)).map_err(|error| error.to_string())?;
    let (old_locations, offsets) = directory_entries(bytes)?;
    let mut entries: Vec<_> = original.walk().collect();
    entries.sort_by(|a, b| a.path().components().count().cmp(&b.path().components().count())
        .then_with(|| a.path().cmp(b.path())));
    let skeleton = empty_allocation(bytes, &old_locations, &offsets)?;
    let mut packed = cfb::CompoundFile::open(Cursor::new(skeleton))
        .map_err(|error| error.to_string())?;
    for entry in entries {
        if entry.is_stream() {
            let mut input = original.open_stream(entry.path()).map_err(|error| error.to_string())?;
            let mut output = packed.open_stream(entry.path()).map_err(|error| error.to_string())?;
            io::copy(&mut input, &mut output).map_err(|error| error.to_string())?;
        }
    }
    packed.flush().map_err(|error| error.to_string())?;
    let mut output = packed.into_inner().into_inner();
    let new_locations = directory_locations(&output)?;
    if old_locations.keys().ne(new_locations.keys()) { return Err("compound objects changed while packing".into()); }
    for (path, old) in old_locations {
        let new = new_locations[&path];
        // Keep allocation, lengths and red-black tree links from the rebuilt
        // directory; copy only class identity, state and exact timestamp bytes.
        output[new + 80..new + 116].copy_from_slice(&bytes[old + 80..old + 116]);
    }
    for range in [8..26, 34..40, 52..56] { output[range.clone()].copy_from_slice(&bytes[range]); }
    Ok(output)
}

#[cfg(test)]
mod tests {
    use super::*;
    use std::io::{Read, Write};

    fn contents(bytes: &[u8]) -> BTreeMap<PathBuf, Vec<u8>> {
        let mut compound = cfb::CompoundFile::open(Cursor::new(bytes)).unwrap();
        let paths: Vec<_> = compound.walk().filter(|entry| entry.is_stream()).map(|entry| entry.path().to_path_buf()).collect();
        paths.into_iter().map(|path| {
            let mut data = Vec::new();
            compound.open_stream(&path).unwrap().read_to_end(&mut data).unwrap();
            (path, data)
        }).collect()
    }

    #[test]
    fn compaction_keeps_live_objects_metadata_and_caches_in_both_versions() {
        for version in [cfb::Version::V3, cfb::Version::V4] {
            let mut compound = cfb::CompoundFile::create_with_version(version, Cursor::new(Vec::new())).unwrap();
            for path in ["/VBA", "/Other", "/Other/Empty"] { compound.create_storage(path).unwrap(); }
            for (path, data) in [("/PROJECT", &b"project\0\xff"[..]), ("/VBA/dir", &b"directory"[..]),
                ("/Other/dir", &b"unknown application data"[..]), ("/VBA/__SRP_0", &[0, 1, 255][..]), ("/Empty", &[][..])] {
                compound.create_stream(path).unwrap().write_all(data).unwrap();
            }
            compound.set_storage_clsid("/", "00112233-4455-6677-8899-aabbccddeeff".parse().unwrap()).unwrap();
            compound.set_storage_clsid("/Other", "fedcba98-7654-3210-fedc-ba9876543210".parse().unwrap()).unwrap();
            compound.set_state_bits("/", 0x1234_5678).unwrap();
            compound.set_state_bits("/VBA/dir", 0xdead_beef).unwrap();
            compound.create_stream("/Discarded").unwrap().write_all(&vec![0x53; 1024 * 1024]).unwrap();
            compound.remove_stream("/Discarded").unwrap();
            compound.flush().unwrap();
            let mut original = compound.into_inner().into_inner();
            let locations = directory_locations(&original).unwrap();
            for (index, at) in locations.values().enumerate() {
                // Retain exact raw timestamp fields, even values a permissive
                // reader maps to zero or an out-of-range SystemTime fallback.
                original[at + 100..at + 108].copy_from_slice(&(index as u64 * 17).to_le_bytes());
                original[at + 108..at + 116].copy_from_slice(&(u64::MAX - index as u64).to_le_bytes());
            }
            let packed = compact(&original).unwrap();
            assert!(packed.len() < original.len() / 8);
            assert_eq!(cfb::CompoundFile::open(Cursor::new(&packed)).unwrap().version(), version);
            assert_eq!(contents(&packed), contents(&original));
            let new_locations = directory_locations(&packed).unwrap();
            assert_eq!(locations.keys().collect::<Vec<_>>(), new_locations.keys().collect::<Vec<_>>());
            for (path, old) in locations {
                let new = new_locations[&path];
                assert_eq!(&original[old + 80..old + 116], &packed[new + 80..new + 116], "{}", path.display());
            }
            assert_eq!(compact(&packed).unwrap(), packed);
        }
    }

    #[test]
    fn compaction_retains_large_streams_through_extended_difat() {
        let data = vec![0x91; 8 * 1024 * 1024];
        let mut compound = cfb::CompoundFile::create_with_version(cfb::Version::V3, Cursor::new(Vec::new())).unwrap();
        compound.create_stream("/Large").unwrap().write_all(&data).unwrap();
        compound.flush().unwrap();
        let original = compound.into_inner().into_inner();
        assert!(word(&original, 72).unwrap() > 0);
        let packed = compact(&original).unwrap();
        assert_eq!(contents(&packed)[&PathBuf::from("/Large")], data);
        assert_eq!(compact(&packed).unwrap(), packed);
    }

    #[test]
    fn empty_directory_allocation_supports_extended_difat() {
        let seed = cfb::CompoundFile::create_with_version(cfb::Version::V3, Cursor::new(Vec::new())).unwrap()
            .into_inner().into_inner();
        let root = directory_locations(&seed).unwrap()[&PathBuf::from("/")];
        // A perfect, all-black sibling tree of 65535 empty streams needs
        // more than the 109 FAT sectors addressable directly in the header.
        let count = 65_536;
        let mut bytes = vec![0; 512 + count * 128];
        bytes[..512].copy_from_slice(&seed[..512]);
        bytes[512..640].copy_from_slice(&seed[root..root + 128]);
        let offsets: Vec<_> = (0..count).map(|id| 512 + id * 128).collect();
        let mut locations = BTreeMap::from([(PathBuf::from("/"), 512)]);
        for (id, at) in offsets.iter().copied().enumerate().skip(1) {
            let name = format!("S{id:05}");
            let units: Vec<_> = name.encode_utf16().collect();
            for (index, unit) in units.iter().enumerate() {
                bytes[at + index * 2..at + index * 2 + 2].copy_from_slice(&unit.to_le_bytes());
            }
            bytes[at + 64..at + 66].copy_from_slice(&((units.len() * 2 + 2) as u16).to_le_bytes());
            bytes[at + 66] = 2;
            bytes[at + 67] = 1;
            for field in [68, 72, 76] { put(&mut bytes, at + field, NONE); }
            put(&mut bytes, at + 116, END);
            locations.insert(PathBuf::from("/").join(name), at);
        }
        fn balance(bytes: &mut [u8], low: usize, high: usize) -> u32 {
            if low == high { return NONE; }
            let middle = low + (high - low) / 2;
            let left = balance(bytes, low, middle);
            let right = balance(bytes, middle + 1, high);
            put(bytes, 512 + middle * 128 + 68, left);
            put(bytes, 512 + middle * 128 + 72, right);
            middle as u32
        }
        let child = balance(&mut bytes, 1, count);
        put(&mut bytes, 512 + 76, child);
        let packed = empty_allocation(&bytes, &locations, &offsets).unwrap();
        assert!(word(&packed, 72).unwrap() > 0);
        let compound = cfb::CompoundFile::open(Cursor::new(&packed)).unwrap();
        assert_eq!(compound.walk().count(), count);
        assert!(compound.walk().filter(|entry| entry.is_stream()).all(|entry| entry.len() == 0));
        assert_eq!(compact(&packed).unwrap(), packed);
    }

    #[test]
    fn truncated_or_cyclic_allocation_is_rejected() {
        for length in [0, 7, 31, 127, 511] { assert!(compact(&vec![0; length]).is_err()); }
        let mut compound = cfb::CompoundFile::create_with_version(cfb::Version::V3, Cursor::new(Vec::new())).unwrap();
        compound.create_storage("/VBA").unwrap();
        let mut bytes = compound.into_inner().into_inner();
        bytes[48..52].copy_from_slice(&0x7fff_ffff_u32.to_le_bytes());
        assert!(directory_locations(&bytes).is_err());
        assert!(compact(&bytes).is_err());
    }
}
