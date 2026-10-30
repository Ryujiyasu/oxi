// This Source Code Form is subject to the terms of the Mozilla Public
// License, v. 2.0. If a copy of the MPL was not distributed with this
// file, You can obtain one at https://mozilla.org/MPL/2.0/.

use oxivba_core::native_cache::{
    decode_runtime_module_heap, CacheContext, RuntimeCacheLink, RuntimeHeapObjectKind,
};

fn hex(text: &str) -> Vec<u8> {
    text.as_bytes()
        .chunks_exact(2)
        .map(|p| u8::from_str_radix(std::str::from_utf8(p).unwrap(), 16).unwrap())
        .collect()
}
fn cases() -> impl Iterator<Item = Vec<&'static str>> {
    include_str!("../../../tests/fixtures/vba/native_cache_module_heaps.tsv")
        .lines()
        .filter(|l| !l.starts_with('#') && !l.starts_with("id\t"))
        .map(|l| l.split('\t').collect())
}
fn joined(values: Vec<String>) -> String {
    if values.is_empty() {
        "-".into()
    } else {
        values.join(";")
    }
}

#[test]
fn native_module_heaps_validate_both_arenas_and_all_typed_objects() {
    let (mut count, mut standard_count, mut member_count, mut signature_count) = (0, 0, 0, 0);
    for row in cases() {
        let bytes = hex(row[1]);
        let heap = decode_runtime_module_heap(&bytes, CacheContext::B5_WIN64)
            .unwrap_or_else(|e| panic!("{}: {e}", row[0]));
        assert_eq!(heap.consumed, bytes.len());
        assert_eq!(heap.root_allocation_units, row[2].parse::<u64>().unwrap());
        assert_eq!(heap.null_image_markers, row[3].parse::<usize>().unwrap());
        assert_eq!(heap.objects.len(), row[4].parse::<usize>().unwrap());
        let (mut standard, mut members, mut fields, mut signatures) =
            (vec![], vec![], vec![], vec![]);
        for object in &heap.objects {
            assert_eq!(object.address % 16, 0);
            assert!(object.allocation_units > 0 && object.allocation_units % 2 == 0);
            match &object.kind {
                RuntimeHeapObjectKind::StandardMethod {
                    code_slot,
                    signature,
                } => {
                    standard.push(format!(
                        "{}:{code_slot}:{}:{}",
                        object.address, signature.key, signature.context
                    ));
                }
                RuntimeHeapObjectKind::MemberMethod {
                    identity,
                    signature,
                    secondary,
                } => {
                    members.push(format!(
                        "{}:{identity}:{}:{}:{}:{}",
                        object.address,
                        signature.key,
                        signature.context,
                        secondary.map_or(u64::MAX, |s| s.key),
                        secondary.map_or_else(|| "-".into(), |s| s.context.to_string())
                    ));
                }
                RuntimeHeapObjectKind::Field {
                    storage_offset,
                    type_word,
                } => {
                    fields.push(format!("{}:{storage_offset}:{type_word}", object.address));
                }
                RuntimeHeapObjectKind::Signature {
                    header,
                    return_type,
                    parameter_types,
                } => {
                    assert_eq!(usize::from(header[1]), parameter_types.len());
                    let raw = header
                        .iter()
                        .copied()
                        .chain([*return_type])
                        .chain(parameter_types.iter().copied())
                        .map(|v| format!("{v:02x}"))
                        .collect::<String>();
                    signatures.push(format!("{}:{raw}", object.address));
                }
                RuntimeHeapObjectKind::Zeroed | RuntimeHeapObjectKind::Vacant => {}
            }
        }
        standard_count += standard.len();
        member_count += members.len();
        signature_count += signatures.len();
        assert_eq!(joined(standard), row[5], "{}", row[0]);
        assert_eq!(joined(members), row[6], "{}", row[0]);
        assert_eq!(joined(fields), row[7], "{}", row[0]);
        assert_eq!(joined(signatures), row[8], "{}", row[0]);
        assert!(heap
            .module
            .methods
            .iter()
            .all(|m| m.frame_info.iter().all(|b| *b == 0)));
        count += 1;
    }
    assert_eq!(
        (count, standard_count, member_count, signature_count),
        (271, 1461, 165, 1644)
    );
}

#[test]
fn native_companion_links_select_code_independently_of_exposed_argument_binding() {
    let (mut calls, mut separate_signatures) = (0, 0);
    for row in cases().filter(|r| r[9] != "-") {
        let bytes = hex(row[1]);
        let heap = decode_runtime_module_heap(&bytes, CacheContext::B5_WIN64).unwrap();
        for binding in row[9].split(';') {
            let fields = binding.split(':').collect::<Vec<_>>();
            let link = RuntimeCacheLink {
                key: fields[1].parse().unwrap(),
                context: fields[2].parse().unwrap(),
            };
            let expected_slot = fields[3].parse::<usize>().unwrap();
            let expected_target_parameters = fields[4].parse::<usize>().unwrap();
            // Measured by Application.Run, before and after explicit source access.
            let accepted_external_parameters = fields[5].parse::<usize>().unwrap();
            let descriptor = heap.object_for_link(2, link).unwrap();
            let RuntimeHeapObjectKind::StandardMethod {
                code_slot,
                signature,
            } = descriptor.kind
            else {
                panic!("standard callable required")
            };
            assert_eq!(usize::from(code_slot), expected_slot);
            let code = heap.standard_method_for_link(2, link).unwrap();
            assert!(std::ptr::eq(code, &heap.module.methods[expected_slot]));
            let RuntimeHeapObjectKind::Signature {
                parameter_types, ..
            } = heap.object_for_link(2, signature).unwrap().kind
            else {
                panic!("signature required")
            };
            assert_eq!(parameter_types.len(), expected_target_parameters);
            if row[0].contains("companion-key-swap") {
                assert_ne!(parameter_types.len(), accepted_external_parameters);
                separate_signatures += 1;
            } else {
                assert_eq!(parameter_types.len(), accepted_external_parameters);
            }
            calls += 1;
        }
    }
    assert_eq!((calls, separate_signatures), (12, 4));
}

#[test]
fn runtime_callable_lookup_rejects_other_streams_displacements_and_ambiguous_keys() {
    let row = cases().find(|r| r[0] == "arity-ordered-control").unwrap();
    let bytes = hex(row[1]);
    let mut heap = decode_runtime_module_heap(&bytes, CacheContext::B5_WIN64).unwrap();
    let link = RuntimeCacheLink {
        key: 0x261,
        context: 2 << 16,
    };
    assert!(heap.standard_method_for_link(2, link).is_some());
    assert!(heap.standard_method_for_link(4, link).is_none());
    assert!(heap
        .standard_method_for_link(
            2,
            RuntimeCacheLink {
                context: link.context | 1,
                ..link
            }
        )
        .is_none());
    for key in [0, 1, 0x260, 0x263, u64::MAX] {
        assert!(heap
            .standard_method_for_link(2, RuntimeCacheLink { key, ..link })
            .is_none());
    }
    let descriptor = heap.object_for_link(2, link).unwrap().clone();
    heap.objects.push(descriptor);
    assert!(heap.object_for_link(2, link).is_none());
    heap.objects.pop();
    let object = heap
        .objects
        .iter_mut()
        .find(|o| o.address == link.key - 1)
        .unwrap();
    if let RuntimeHeapObjectKind::StandardMethod { code_slot, .. } = &mut object.kind {
        *code_slot = u16::MAX;
    }
    assert!(heap.standard_method_for_link(2, link).is_none());
}

#[test]
fn runtime_heap_rejects_every_truncation_and_trailing_bytes() {
    for name in [
        "arity-ordered-control",
        "classes/CacheClassFields",
        "multiple-documents/ThisWorkbook",
    ] {
        let row = cases().find(|r| r[0] == name).unwrap();
        let mut bytes = hex(row[1]);
        for end in 0..bytes.len() {
            assert!(
                decode_runtime_module_heap(&bytes[..end], CacheContext::B5_WIN64).is_err(),
                "{name}: {end}"
            );
        }
        bytes.push(0);
        assert!(decode_runtime_module_heap(&bytes, CacheContext::B5_WIN64).is_err());
    }
}

#[test]
fn runtime_heap_rejects_corrupt_sizes_kinds_signatures_and_contexts() {
    let row = cases().find(|r| r[0] == "arity-ordered-control").unwrap();
    let bytes = hex(row[1]);
    let heap = decode_runtime_module_heap(&bytes, CacheContext::B5_WIN64).unwrap();
    let descriptor = heap
        .objects
        .iter()
        .find(|o| matches!(o.kind, RuntimeHeapObjectKind::StandardMethod { .. }))
        .unwrap()
        .serialized_offset;
    let signature = heap
        .objects
        .iter()
        .find(|o| matches!(o.kind, RuntimeHeapObjectKind::Signature { .. }))
        .unwrap()
        .serialized_offset;
    for (at, data) in [
        (2, u64::MAX.to_le_bytes().to_vec()),
        (10, u64::MAX.to_le_bytes().to_vec()),
        (descriptor + 7, vec![0x7d]),
        (descriptor, vec![0; 7]),
        (signature + 8, vec![1]),
        (signature + 9, vec![255]),
    ] {
        let mut bad = bytes.clone();
        bad[at..at + data.len()].copy_from_slice(&data);
        assert!(
            decode_runtime_module_heap(&bad, CacheContext::B5_WIN64).is_err(),
            "offset {at}"
        );
    }
    assert!(decode_runtime_module_heap(
        &bytes,
        CacheContext {
            version: 0xb4,
            sys_kind: 3
        }
    )
    .is_err());
    assert!(decode_runtime_module_heap(
        &bytes,
        CacheContext {
            version: 0xb5,
            sys_kind: 1
        }
    )
    .is_err());
}
