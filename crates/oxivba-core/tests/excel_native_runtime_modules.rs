// This Source Code Form is subject to the terms of the Mozilla Public
// License, v. 2.0. If a copy of the MPL was not distributed with this
// file, You can obtain one at https://mozilla.org/MPL/2.0/.

use oxivba_core::native_cache::{
    decode_runtime_module_cache, decode_runtime_numeric_assignment, CacheContext,
    RuntimeBindingSlots,
};

fn hex(text: &str) -> Vec<u8> {
    text.as_bytes()
        .chunks_exact(2)
        .map(|p| u8::from_str_radix(std::str::from_utf8(p).unwrap(), 16).unwrap())
        .collect()
}
fn cases() -> impl Iterator<Item = Vec<&'static str>> {
    include_str!("../../../tests/fixtures/vba/native_cache_runtime_modules.tsv")
        .lines()
        .filter(|l| !l.starts_with('#') && !l.starts_with("id\t"))
        .map(|l| l.split('\t').collect())
}
fn numbers(text: &str) -> Vec<usize> {
    if text == "-" {
        vec![]
    } else {
        text.split(',').map(|v| v.parse().unwrap()).collect()
    }
}
fn location(bytes: &[u8], slice: &[u8]) -> usize {
    slice.as_ptr() as usize - bytes.as_ptr() as usize
}

#[test]
fn native_runtime_modules_follow_headers_bindings_and_method_records() {
    let (mut count, mut images) = (0, 0);
    for row in cases() {
        let bytes = hex(row[1]);
        let module = decode_runtime_module_cache(&bytes, CacheContext::B5_WIN64)
            .unwrap_or_else(|e| panic!("{}: {e}", row[0]));
        assert_eq!(module.consumed, row[2].parse::<usize>().unwrap());
        assert_eq!(module.consumed + 8, bytes.len());
        assert_eq!(module.companion_stream, row[3].parse::<u16>().unwrap());
        assert_eq!(
            module
                .methods
                .iter()
                .map(|m| m.code.len())
                .collect::<Vec<_>>(),
            numbers(row[5])
        );
        assert_eq!(
            module
                .methods
                .iter()
                .map(|m| m.following_references.len())
                .collect::<Vec<_>>(),
            numbers(row[6])
        );
        assert!(module
            .methods
            .iter()
            .all(|m| m.frame_info.iter().all(|b| *b == 0)));
        match module.class_binding.as_ref() {
            None => assert_eq!(row[4], "standard"),
            Some(binding) => {
                assert_eq!(
                    binding.invocation_links.len(),
                    row[9].parse::<usize>().unwrap()
                );
                match &binding.slots {
                    RuntimeBindingSlots::Class(slots) => {
                        assert_eq!(row[4], "class");
                        assert_eq!(
                            slots
                                .iter()
                                .map(|s| usize::from(s.is_some()))
                                .collect::<Vec<_>>(),
                            numbers(row[8])
                        );
                    }
                    RuntimeBindingSlots::Host {
                        capacity, events, ..
                    } => {
                        assert_eq!(row[4], format!("host{capacity}"));
                        let actual = events
                            .iter()
                            .map(|e| {
                                format!(
                                    "{}:{}:{}:{}",
                                    e.dispatch_id, e.member_id, e.target.key, e.target.context
                                )
                            })
                            .collect::<Vec<_>>()
                            .join(";");
                        assert_eq!(if actual.is_empty() { "-" } else { &actual }, row[7]);
                    }
                }
            }
        }
        count += 1;
        images += module.methods.len();
    }
    assert_eq!((count, images), (15, 133));
}

#[test]
fn multiple_host_events_keep_separate_identity_and_target_arrays() {
    let row = cases()
        .find(|r| r[0] == "multiple-documents/__SRP_2")
        .unwrap();
    let bytes = hex(row[1]);
    let module = decode_runtime_module_cache(&bytes, CacheContext::B5_WIN64).unwrap();
    let binding = module.class_binding.as_ref().unwrap();
    let RuntimeBindingSlots::Host {
        capacity,
        events,
        unused,
    } = &binding.slots
    else {
        panic!("host")
    };
    assert_eq!(*capacity, 55);
    assert_eq!(events.len(), 2);
    assert_eq!(unused.len(), 53 * 8);
    assert_eq!(
        (events[0].dispatch_id, events[0].member_id),
        (1564, 0x60030000)
    );
    assert_eq!(
        (events[1].dispatch_id, events[1].member_id),
        (1558, 0x60030001)
    );
    assert_eq!(
        events.iter().map(|e| e.target.context).collect::<Vec<_>>(),
        [2 << 16, 2 << 16]
    );
    assert_ne!(events[0].target.key, events[1].target.key);
    assert_eq!(module.methods.len(), 3);
    // Host auxiliary keys are opaque; even values are not runtime image records.
    assert_eq!(module.auxiliary_links[0].key, 592);
}

#[test]
fn class_fields_initializers_and_numeric_method_images_remain_distinct() {
    for (name, index, expected) in [
        ("long/__SRP_2", 4, 21.0),
        ("native/__SRP_2", 0, 14.0),
        ("native/__SRP_2", 1, 21.0),
        ("native/__SRP_4", 0, 28.0),
    ] {
        let row = cases().find(|r| r[0] == name).unwrap();
        let bytes = hex(row[1]);
        let module = decode_runtime_module_cache(&bytes, CacheContext::B5_WIN64).unwrap();
        let numeric = decode_runtime_numeric_assignment(
            &module.methods[index].code[12..],
            CacheContext::B5_WIN64,
        )
        .unwrap();
        assert_eq!(numeric.evaluate_double(7.0).unwrap(), expected);
    }
    let row = cases().find(|r| r[0] == "native/__SRP_c").unwrap();
    let bytes = hex(row[1]);
    let empty = decode_runtime_module_cache(&bytes, CacheContext::B5_WIN64).unwrap();
    assert!(empty.methods.is_empty());
    assert_eq!(empty.class_binding.unwrap().invocation_links.len(), 4);
    let row = cases().find(|r| r[0] == "native/__SRP_e").unwrap();
    let bytes = hex(row[1]);
    let initialized = decode_runtime_module_cache(&bytes, CacheContext::B5_WIN64).unwrap();
    let RuntimeBindingSlots::Class(slots) = initialized.class_binding.unwrap().slots else {
        panic!("class")
    };
    assert_eq!(slots.iter().filter(|s| s.is_some()).count(), 1);
    assert_eq!(initialized.methods.len(), 2);
}

#[test]
fn runtime_module_prefixes_require_complete_records_and_boundary_lookahead() {
    for row in cases() {
        let mut bytes = hex(row[1]);
        for length in 0..bytes.len() {
            assert!(
                decode_runtime_module_cache(&bytes[..length], CacheContext::B5_WIN64).is_err(),
                "{} {length}",
                row[0]
            );
        }
        let consumed = row[2].parse::<usize>().unwrap();
        bytes.extend_from_slice(&[0xff; 32]);
        assert_eq!(
            decode_runtime_module_cache(&bytes, CacheContext::B5_WIN64)
                .unwrap()
                .consumed,
            consumed
        );
    }
}

#[test]
fn opaque_frame_bytes_do_not_change_runtime_module_boundaries() {
    let row = cases().find(|r| r[0] == "long/__SRP_2").unwrap();
    let mut bytes = hex(row[1]);
    let module = decode_runtime_module_cache(&bytes, CacheContext::B5_WIN64).unwrap();
    let offset = location(&bytes, module.methods[0].frame_info);
    let consumed = module.consumed;
    let sizes = module
        .methods
        .iter()
        .map(|m| m.code.len())
        .collect::<Vec<_>>();
    bytes[offset..offset + 8].copy_from_slice(&0x5572u64.to_le_bytes());
    bytes[offset + 8..offset + 16].copy_from_slice(&0x1231u64.to_le_bytes());
    let changed = decode_runtime_module_cache(&bytes, CacheContext::B5_WIN64).unwrap();
    assert_eq!(changed.consumed, consumed);
    assert_eq!(
        changed
            .methods
            .iter()
            .map(|m| m.code.len())
            .collect::<Vec<_>>(),
        sizes
    );
}

#[test]
fn malformed_header_counts_and_host_binding_tables_fail_safely() {
    let row = cases().find(|r| r[0] == "long/__SRP_2").unwrap();
    let bytes = hex(row[1]);
    for (at, data) in [
        (0, vec![0, 0]),
        (34, 1u64.to_le_bytes().to_vec()),
        (88, 2u16.to_le_bytes().to_vec()),
        (102, u16::MAX.to_le_bytes().to_vec()),
        (106, u32::MAX.to_le_bytes().to_vec()),
    ] {
        let mut bad = bytes.clone();
        bad[at..at + data.len()].copy_from_slice(&data);
        assert!(
            decode_runtime_module_cache(&bad, CacheContext::B5_WIN64).is_err(),
            "offset{at}"
        );
    }
    for context in [
        CacheContext {
            version: 180,
            sys_kind: 3,
        },
        CacheContext {
            version: 181,
            sys_kind: 1,
        },
    ] {
        assert!(decode_runtime_module_cache(&bytes, context).is_err());
    }
    let row = cases()
        .find(|r| r[0] == "multiple-documents/__SRP_2")
        .unwrap();
    let bytes = hex(row[1]);
    let module = decode_runtime_module_cache(&bytes, CacheContext::B5_WIN64).unwrap();
    let RuntimeBindingSlots::Host {
        capacity,
        events,
        unused,
    } = &module.class_binding.as_ref().unwrap().slots
    else {
        panic!("host")
    };
    let unused_at = location(&bytes, unused);
    let count_at = unused_at - events.len() * 20 - 2;
    let end = unused_at + unused.len();
    for (at, data) in [
        (count_at, (capacity + 1).to_le_bytes().to_vec()),
        (unused_at, vec![0]),
        (end, 0u32.to_le_bytes().to_vec()),
        (end + 4, u16::MAX.to_le_bytes().to_vec()),
    ] {
        let mut bad = bytes.clone();
        bad[at..at + data.len()].copy_from_slice(&data);
        assert!(
            decode_runtime_module_cache(&bad, CacheContext::B5_WIN64).is_err(),
            "host offset{at}"
        );
    }
}
