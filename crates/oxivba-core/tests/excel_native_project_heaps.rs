// This Source Code Form is subject to the terms of the Mozilla Public
// License, v. 2.0. If a copy of the MPL was not distributed with this
// file, You can obtain one at https://mozilla.org/MPL/2.0/.

use oxivba_core::native_cache::{
    decode_runtime_module_heap, decode_runtime_numeric_assignment, decode_runtime_project_heap,
    CacheContext, RuntimeCacheLink, RuntimeEncodedName, RuntimeNameEncoding,
    RuntimeProjectModuleState, RuntimeProjectObjectKind,
};

fn hex(text: &str) -> Vec<u8> {
    text.as_bytes()
        .chunks_exact(2)
        .map(|p| u8::from_str_radix(std::str::from_utf8(p).unwrap(), 16).unwrap())
        .collect()
}
fn cases() -> impl Iterator<Item = Vec<&'static str>> {
    include_str!("../../../tests/fixtures/vba/native_cache_project_heaps.tsv")
        .lines()
        .filter(|l| !l.starts_with('#') && !l.starts_with("id\t"))
        .map(|l| l.split('\t').collect())
}
fn name_eq(actual: RuntimeEncodedName<'_>, expected: &str) {
    let (kind, raw) = expected.split_once(':').unwrap();
    assert_eq!(
        actual.encoding,
        if kind == "1" {
            RuntimeNameEncoding::Utf16Le
        } else {
            RuntimeNameEncoding::Ansi
        }
    );
    assert_eq!(actual.bytes, hex(raw));
}
fn kind(object: &RuntimeProjectObjectKind<'_>) -> u8 {
    match object {
        RuntimeProjectObjectKind::Name(n) => {
            if n.encoding == RuntimeNameEncoding::Utf16Le {
                1
            } else {
                2
            }
        }
        RuntimeProjectObjectKind::Guid(_) => 3,
        RuntimeProjectObjectKind::LinkArray { .. } => 5,
        RuntimeProjectObjectKind::Field { .. } => 6,
        RuntimeProjectObjectKind::Member { .. } => 7,
        RuntimeProjectObjectKind::ReferenceLibrary { .. } => 10,
        RuntimeProjectObjectKind::Bstr(_) => 11,
        RuntimeProjectObjectKind::ValueRecord { .. } => 13,
        RuntimeProjectObjectKind::Signature { .. } => 14,
        RuntimeProjectObjectKind::Zeroed => 126,
        RuntimeProjectObjectKind::Vacant => 127,
    }
}

#[test]
fn native_global_heaps_recover_encoded_names_and_module_ownership_without_source() {
    let (mut count, mut modules, mut procedures) = (0, 0, 0);
    for row in cases() {
        let bytes = hex(row[1]);
        let heap = decode_runtime_project_heap(&bytes, CacheContext::B5_WIN64)
            .unwrap_or_else(|e| panic!("{}: {e}", row[0]));
        assert_eq!(heap.consumed, bytes.len());
        name_eq(heap.name_for_link(heap.name).unwrap(), row[2]);
        let names = row[3].split(';').collect::<Vec<_>>();
        let entries = row[4].split(';').collect::<Vec<_>>();
        assert_eq!(heap.modules.len(), names.len());
        for (index, module) in heap.modules.iter().enumerate() {
            name_eq(heap.module_name(index).unwrap(), names[index]);
            if let RuntimeProjectModuleState::Ready(ready) = &module.state {
                let expected = if entries[index] == "-" {
                    vec![]
                } else {
                    entries[index].split(',').collect()
                };
                assert_eq!(ready.procedure_names.len(), expected.len());
                for (entry, name) in expected.iter().enumerate() {
                    name_eq(heap.procedure_name(index, entry).unwrap(), name);
                    procedures += 1;
                }
                assert_eq!(ready.root.key, 17);
                assert_eq!(ready.root.context & 0xffff, 0);
            } else {
                assert!(heap.procedure_name(index, 0).is_none());
            }
            modules += 1;
        }
        assert_eq!(
            heap.objects
                .iter()
                .map(|o| kind(&o.kind))
                .collect::<Vec<_>>(),
            row[5]
                .split(',')
                .map(|v| v.parse::<u8>().unwrap())
                .collect::<Vec<_>>()
        );
        assert!(heap.objects.iter().all(|o| o.address % 16 == 0));
        assert!(heap.objects.iter().all(|o| match &o.kind {
            RuntimeProjectObjectKind::ValueRecord { raw, .. } => raw.iter().all(|b| *b == 0),
            _ => true,
        }));
        count += 1;
    }
    assert_eq!((count, modules, procedures), (87, 363, 1694));
}

#[test]
fn native_named_calls_use_canonical_entries_and_independent_companion_code_targets() {
    let (mut queries, mut absent, mut evaluated) = (0, 0, 0);
    for row in cases().filter(|r| r[6] != "-") {
        let bytes = hex(row[1]);
        let project = decode_runtime_project_heap(&bytes, CacheContext::B5_WIN64).unwrap();
        let primary = hex(row[7]);
        let heap = decode_runtime_module_heap(&primary, CacheContext::B5_WIN64).unwrap();
        let links = row[8]
            .split(';')
            .map(|s| {
                let (key, context) = s.split_once(':').unwrap();
                RuntimeCacheLink {
                    key: key.parse().unwrap(),
                    context: context.parse().unwrap(),
                }
            })
            .collect::<Vec<_>>();
        for query in row[6].split(';') {
            let fields = query.split(':').collect::<Vec<_>>();
            let module = fields[0].parse::<usize>().unwrap();
            let name = hex(fields[1]);
            let RuntimeProjectModuleState::Ready(ready) = &project.modules[module].state else {
                panic!("ready module required")
            };
            let matches = (0..ready.procedure_names.len())
                .filter(|i| {
                    let n = project.procedure_name(module, *i).unwrap();
                    n.encoding == RuntimeNameEncoding::Ansi && n.bytes == name
                })
                .collect::<Vec<_>>();
            if fields[2] == "-" {
                assert!(matches.is_empty());
                absent += 1;
            } else {
                let entry = fields[2].parse::<usize>().unwrap();
                assert_eq!(matches, vec![entry]);
                let image = heap
                    .standard_method_for_link((ready.root.context >> 16) as u16, links[entry])
                    .unwrap();
                let slot = fields[3].parse::<usize>().unwrap();
                assert!(std::ptr::eq(image, &heap.module.methods[slot]));
                // These authored arithmetic controls have two six-byte prologue
                // instructions. This is an assignment-span oracle, not a VM loader.
                if fields[4] != "0" && image.code.len() > 12 {
                    if let Ok(assignment) =
                        decode_runtime_numeric_assignment(&image.code[12..], CacheContext::B5_WIN64)
                    {
                        assert_eq!(
                            assignment.evaluate_double(7.0).unwrap(),
                            fields[5].parse::<f64>().unwrap(),
                            "{}: {}",
                            row[0],
                            fields[1]
                        );
                        evaluated += 1;
                    }
                }
            }
            queries += 1;
        }
    }
    assert_eq!((queries, absent, evaluated), (62, 8, 35));
}

#[test]
fn japanese_native_names_preserve_ansi_bytes_and_wide_strings_remain_distinct() {
    let row = cases().find(|r| r[0] == "japanese-reopen-exact").unwrap();
    let bytes = hex(row[1]);
    let heap = decode_runtime_project_heap(&bytes, CacheContext::B5_WIN64).unwrap();
    let names = [
        ("Driver", "447269766572"),
        ("Alpha", "416c706861"),
        ("日本語", "93fa967b8cea"),
        ("計算", "8c768e5a"),
    ];
    for (entry, (_, raw)) in names.iter().enumerate() {
        let name = heap.procedure_name(2, entry).unwrap();
        assert_eq!(name.encoding, RuntimeNameEncoding::Ansi);
        assert_eq!(name.bytes, hex(raw));
    }
    let mut wide = 0;
    for row in cases() {
        let bytes = hex(row[1]);
        let heap = decode_runtime_project_heap(&bytes, CacheContext::B5_WIN64).unwrap();
        for object in &heap.objects {
            if let RuntimeProjectObjectKind::Name(name) = object.kind {
                if name.encoding == RuntimeNameEncoding::Utf16Le {
                    assert_eq!(name.bytes.len() % 2, 0);
                    let units = name
                        .bytes
                        .chunks_exact(2)
                        .map(|b| u16::from_le_bytes([b[0], b[1]]))
                        .collect::<Vec<_>>();
                    assert!(String::from_utf16(&units).is_ok());
                    wide += 1;
                }
            }
        }
    }
    assert!(wide >= 81);
}

#[test]
fn global_heap_rejects_every_truncation_and_trailing_data() {
    for name in ["japanese-reopen-exact", "classes", "records"] {
        let row = cases().find(|r| r[0] == name).unwrap();
        let mut bytes = hex(row[1]);
        for end in 0..bytes.len() {
            assert!(
                decode_runtime_project_heap(&bytes[..end], CacheContext::B5_WIN64).is_err(),
                "{name}: {end}"
            );
        }
        bytes.push(0);
        assert!(decode_runtime_project_heap(&bytes, CacheContext::B5_WIN64).is_err());
    }
}

#[test]
fn global_names_reject_wrong_context_non_names_and_duplicate_addresses() {
    let row = cases().find(|r| r[0] == "japanese-reopen-exact").unwrap();
    let bytes = hex(row[1]);
    let mut heap = decode_runtime_project_heap(&bytes, CacheContext::B5_WIN64).unwrap();
    let link = heap.modules[2].name;
    assert!(heap.name_for_link(link).is_some());
    for bad in [
        RuntimeCacheLink {
            context: 1 << 16,
            ..link
        },
        RuntimeCacheLink {
            key: link.key + 1,
            ..link
        },
        RuntimeCacheLink {
            key: u64::MAX,
            ..link
        },
    ] {
        assert!(heap.name_for_link(bad).is_none());
    }
    let guid = heap
        .objects
        .iter()
        .find(|o| matches!(o.kind, RuntimeProjectObjectKind::Guid(_)))
        .unwrap();
    assert!(heap
        .name_for_link(RuntimeCacheLink {
            key: guid.address + 1,
            context: 0
        })
        .is_none());
    heap.objects
        .push(heap.object_for_link(link).unwrap().clone());
    assert!(heap.name_for_link(link).is_none());
    assert!(heap.module_name(usize::MAX).is_none());
    assert!(heap.procedure_name(2, usize::MAX).is_none());
}

#[test]
fn global_heap_rejects_corrupt_owner_sizes_object_tags_and_name_lengths() {
    let row = cases().find(|r| r[0] == "japanese-reopen-exact").unwrap();
    let bytes = hex(row[1]);
    let heap = decode_runtime_project_heap(&bytes, CacheContext::B5_WIN64).unwrap();
    let start = heap.directory.root_heap_offset;
    let object = heap.objects[0].serialized_offset;
    let mut root = start + 34;
    while u64::from_le_bytes(bytes[root..root + 8].try_into().unwrap()) != 0 {
        root += 8;
    }
    root += 8;
    for (at, data) in [
        (start + 2, u64::MAX.to_le_bytes().to_vec()),
        (root, 36u64.to_le_bytes().to_vec()),
        (root + 16, vec![0; 4]),
        (object + 7, vec![0x7d]),
        (object + 8, u64::MAX.to_le_bytes().to_vec()),
    ] {
        let mut bad = bytes.clone();
        bad[at..at + data.len()].copy_from_slice(&data);
        assert!(
            decode_runtime_project_heap(&bad, CacheContext::B5_WIN64).is_err(),
            "{at}"
        );
    }
    assert!(decode_runtime_project_heap(
        &bytes,
        CacheContext {
            version: 0xb4,
            sys_kind: 3
        }
    )
    .is_err());
    assert!(decode_runtime_project_heap(
        &bytes,
        CacheContext {
            version: 0xb5,
            sys_kind: 1
        }
    )
    .is_err());
    let mut checked = false;
    for row in cases() {
        let mut bytes = hex(row[1]);
        let heap = decode_runtime_project_heap(&bytes, CacheContext::B5_WIN64).unwrap();
        if let Some(object) = heap.objects.iter().find(|o| {
            matches!(
                o.kind,
                RuntimeProjectObjectKind::Name(RuntimeEncodedName {
                    encoding: RuntimeNameEncoding::Utf16Le,
                    ..
                })
            )
        }) {
            let at = object.serialized_offset + 8;
            bytes[at] |= 1;
            assert!(decode_runtime_project_heap(&bytes, CacheContext::B5_WIN64).is_err());
            checked = true;
            break;
        }
    }
    assert!(checked);
}
