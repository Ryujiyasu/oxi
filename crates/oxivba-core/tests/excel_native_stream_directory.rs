// This Source Code Form is subject to the terms of the Mozilla Public
// License, v. 2.0. If a copy of the MPL was not distributed with this
// file, You can obtain one at https://mozilla.org/MPL/2.0/.

use oxivba_core::native_cache::{
    decode_runtime_module_cache, decode_runtime_stream_directory, CacheContext, DecodeErrorKind,
    RuntimeStreamKind as Kind, RuntimeStreamSlot as Slot,
};

fn hex(text: &str) -> Vec<u8> {
    text.as_bytes()
        .chunks_exact(2)
        .map(|p| u8::from_str_radix(std::str::from_utf8(p).unwrap(), 16).unwrap())
        .collect()
}
fn cases() -> impl Iterator<Item = Vec<&'static str>> {
    include_str!("../../../tests/fixtures/vba/native_cache_stream_directory.tsv")
        .lines()
        .filter(|l| !l.starts_with('#') && !l.starts_with("id\t"))
        .map(|l| l.split('\t').collect())
}
fn kind(tag: &str) -> Kind {
    match tag {
        "0" => Kind::Primary,
        "2" => Kind::Companion,
        _ => panic!("kind"),
    }
}

#[test]
fn native_stream_tables_preserve_vacancies_and_cached_module_ownership() {
    let (mut projects, mut modules) = (0, 0);
    for row in cases() {
        let bytes = hex(row[1]);
        let directory = decode_runtime_stream_directory(&bytes, CacheContext::B5_WIN64).unwrap();
        assert_eq!(directory.slots.len(), row[2].parse::<usize>().unwrap());
        assert_eq!(directory.root_heap_offset, row[3].parse::<usize>().unwrap());
        assert_eq!(directory.consumed + 2, bytes.len());
        assert_eq!(directory.unique_stream(None, Kind::Primary), Some(0));
        assert_eq!(directory.unique_stream(None, Kind::Companion), Some(1));
        let mut occupied = 0;
        for expected in row[4].split(';') {
            let f: Vec<_> = expected.split(':').collect();
            let index = f[0].parse::<usize>().unwrap();
            let Slot::Occupied {
                kind: actual,
                module_index,
                metadata,
            } = &directory.slots[index]
            else {
                panic!("occupied")
            };
            assert_eq!(*actual, kind(f[1]));
            assert_eq!(
                *module_index,
                if f[2] == "-" {
                    None
                } else {
                    Some(f[2].parse().unwrap())
                }
            );
            assert_eq!(metadata.to_vec(), hex(f[3]));
            occupied += 1;
        }
        assert_eq!(
            directory
                .slots
                .iter()
                .filter(|s| matches!(s, Slot::Occupied { .. }))
                .count(),
            occupied
        );
        let expected: Vec<_> = if row[6] == "-" {
            vec![]
        } else {
            row[6]
                .split(';')
                .map(|s| {
                    s.split(':')
                        .map(|v| v.parse::<u16>().unwrap())
                        .collect::<Vec<_>>()
                })
                .collect()
        };
        for values in &expected {
            assert_eq!(
                directory.unique_stream(Some(values[0]), Kind::Primary),
                Some(values[1])
            );
            assert_eq!(
                directory.unique_stream(Some(values[0]), Kind::Companion),
                Some(values[2])
            );
            modules += 1;
        }
        for index in 0..row[5].split(',').count() {
            if !expected.iter().any(|v| usize::from(v[0]) == index) {
                assert_eq!(
                    directory.unique_stream(Some(index as u16), Kind::Primary),
                    None
                );
            }
        }
        projects += 1;
    }
    assert_eq!((projects, modules), (66, 265));
}

#[test]
fn directory_links_cached_module_labels_to_the_actual_runtime_module_prefix() {
    for (project, label, module_fixture, expected_stream) in [
        ("classes", "CacheClassPlain", "native/__SRP_2", 2),
        ("classes", "CacheClassFields", "native/__SRP_4", 4),
        ("documents", "ThisWorkbook", "documents/__SRP_2", 2),
        ("documents", "Sheet1", "documents/__SRP_4", 4),
        ("records", "SaveProbe", "records/__SRP_6", 6),
    ] {
        let row = cases().find(|r| r[0] == project).unwrap();
        let bytes = hex(row[1]);
        let directory = decode_runtime_stream_directory(&bytes, CacheContext::B5_WIN64).unwrap();
        let encoded = label
            .encode_utf16()
            .flat_map(u16::to_le_bytes)
            .collect::<Vec<_>>();
        let index = row[5]
            .split(',')
            .position(|name| hex(name) == encoded)
            .unwrap() as u16;
        let stream = directory.unique_stream(Some(index), Kind::Primary).unwrap();
        assert_eq!(stream, expected_stream);
        let module = include_str!("../../../tests/fixtures/vba/native_cache_runtime_modules.tsv")
            .lines()
            .find(|l| l.split('\t').next() == Some(module_fixture))
            .unwrap();
        let module_bytes = hex(module.split('\t').nth(1).unwrap());
        let parsed = decode_runtime_module_cache(&module_bytes, CacheContext::B5_WIN64).unwrap();
        assert_eq!(
            directory.unique_stream(Some(index), Kind::Companion),
            Some(parsed.companion_stream)
        );
    }
}

fn synthetic(slots: &[Slot]) -> Vec<u8> {
    let mut bytes = hex("934b2ab50300");
    bytes.extend_from_slice(&(slots.len() as u16).to_le_bytes());
    for slot in slots {
        match slot {
            Slot::Vacant => bytes.extend_from_slice(&5u16.to_le_bytes()),
            Slot::Occupied {
                kind,
                module_index,
                metadata,
            } => {
                bytes.extend_from_slice(
                    &(if *kind == Kind::Primary { 0u16 } else { 2u16 }).to_le_bytes(),
                );
                bytes.extend_from_slice(&module_index.unwrap_or(u16::MAX).to_le_bytes());
                bytes.extend_from_slice(metadata);
            }
        }
    }
    bytes.extend_from_slice(&[0, 0, 0x72, 0x55]);
    bytes
}
fn occupied(kind: Kind, owner: Option<u16>) -> Slot {
    Slot::Occupied {
        kind,
        module_index: owner,
        metadata: [0, 0, 0, 0, 1, 0],
    }
}

#[test]
fn sparse_nonadjacent_and_ambiguous_slots_do_not_create_guessed_bindings() {
    let mut slots = vec![
        occupied(Kind::Primary, None),
        occupied(Kind::Companion, None),
        occupied(Kind::Primary, Some(42)),
        Slot::Vacant,
        Slot::Vacant,
        occupied(Kind::Companion, Some(42)),
        Slot::Vacant,
        Slot::Vacant,
    ];
    // Root-signature bytes in an opaque metadata field are not a heap boundary.
    let Slot::Occupied { metadata, .. } = &mut slots[0] else {
        unreachable!()
    };
    metadata[..2].copy_from_slice(&[0x72, 0x55]);
    let bytes = synthetic(&slots);
    let parsed = decode_runtime_stream_directory(&bytes, CacheContext::B5_WIN64).unwrap();
    assert_eq!(parsed.unique_stream(Some(42), Kind::Primary), Some(2));
    assert_eq!(parsed.unique_stream(Some(42), Kind::Companion), Some(5));
    assert_eq!(parsed.unique_stream(Some(0), Kind::Primary), None);
    assert_eq!(parsed.root_heap_offset + 2, bytes.len());
    slots[6] = occupied(Kind::Primary, Some(42));
    let parsed =
        decode_runtime_stream_directory(&synthetic(&slots), CacheContext::B5_WIN64).unwrap();
    assert_eq!(parsed.unique_stream(Some(42), Kind::Primary), None);
    assert_eq!(parsed.unique_stream(Some(42), Kind::Companion), Some(5));
}

#[test]
fn stream_directory_bounds_contexts_tags_and_root_ownership_are_checked() {
    for row in cases() {
        let mut bytes = hex(row[1]);
        for length in 0..bytes.len() {
            assert!(
                decode_runtime_stream_directory(&bytes[..length], CacheContext::B5_WIN64).is_err(),
                "{} {length}",
                row[0]
            );
        }
        let consumed = row[3].parse::<usize>().unwrap();
        bytes.extend_from_slice(&[0xff; 32]);
        assert_eq!(
            decode_runtime_stream_directory(&bytes, CacheContext::B5_WIN64)
                .unwrap()
                .consumed,
            consumed
        );
    }
    let mut bytes = hex(cases().next().unwrap()[1]);
    for (at, data) in [
        (0, vec![0, 0]),
        (2, vec![0]),
        (3, vec![180]),
        (4, vec![1, 0]),
        (6, u16::MAX.to_le_bytes().to_vec()),
        (10, vec![0, 0]),
    ] {
        let mut bad = bytes.clone();
        bad[at..at + data.len()].copy_from_slice(&data);
        assert!(decode_runtime_stream_directory(&bad, CacheContext::B5_WIN64).is_err());
    }
    bytes[8..10].copy_from_slice(&7u16.to_le_bytes());
    assert_eq!(
        decode_runtime_stream_directory(&bytes, CacheContext::B5_WIN64)
            .unwrap_err()
            .kind,
        DecodeErrorKind::UnsupportedRuntimeStreamKind(7)
    );
    let bytes = hex(cases().next().unwrap()[1]);
    for context in [
        CacheContext {
            version: 182,
            sys_kind: 3,
        },
        CacheContext {
            version: 181,
            sys_kind: 1,
        },
    ] {
        assert!(decode_runtime_stream_directory(&bytes, context).is_err());
    }
}
