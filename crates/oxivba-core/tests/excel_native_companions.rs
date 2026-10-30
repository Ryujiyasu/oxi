// This Source Code Form is subject to the terms of the Mozilla Public
// License, v. 2.0. If a copy of the MPL was not distributed with this
// file, You can obtain one at https://mozilla.org/MPL/2.0/.

use oxivba_core::native_cache::{
    decode_runtime_companion_cache, decode_runtime_invocation_table, decode_runtime_module_heap,
    decode_runtime_numeric_assignment, decode_runtime_project_heap, CacheContext,
    RuntimeCompanionType, RuntimeHeapObjectKind, RuntimeNameEncoding, RuntimeProjectModuleState,
};

fn hex(text: &str) -> Vec<u8> {
    text.as_bytes()
        .chunks_exact(2)
        .map(|b| u8::from_str_radix(std::str::from_utf8(b).unwrap(), 16).unwrap())
        .collect()
}
fn cases() -> impl Iterator<Item = Vec<&'static str>> {
    include_str!("../../../tests/fixtures/vba/native_cache_companions.tsv")
        .lines()
        .filter(|l| !l.starts_with('#') && !l.starts_with("id\t"))
        .map(|l| l.split('\t').collect())
}
fn typed(t: &RuntimeCompanionType) -> String {
    format!(
        "{},{},{}",
        t.word,
        t.array_word.map_or_else(|| "-".into(), |v| v.to_string()),
        t.reference
            .map_or_else(|| "-".into(), |v| format!("{},{}", v.key, v.context))
    )
}
fn list(values: Vec<String>) -> String {
    if values.is_empty() {
        "-".into()
    } else {
        values.join(";")
    }
}

#[test]
fn native_companions_frame_from_primary_descriptors_and_preserve_types_fields_and_links() {
    let (mut count, mut methods, mut fields, mut materialized) = (0, 0, 0, 0);
    for row in cases() {
        let g = hex(row[2]);
        let p = hex(row[3]);
        let c = hex(row[4]);
        let project = decode_runtime_project_heap(&g, CacheContext::B5_WIN64).unwrap();
        let primary = decode_runtime_module_heap(&p, CacheContext::B5_WIN64).unwrap();
        let table = decode_runtime_invocation_table(
            &project,
            row[1].parse().unwrap(),
            &primary,
            &c,
            CacheContext::B5_WIN64,
        )
        .unwrap_or_else(|e| panic!("{}: {e}", row[0]));
        assert_eq!(table.consumed, c.len());
        assert_eq!(
            table.standard_module,
            primary.module.class_binding.is_none()
        );
        assert_eq!(
            table.methods.len(),
            primary
                .objects
                .iter()
                .filter(|o| matches!(
                    o.kind,
                    RuntimeHeapObjectKind::StandardMethod { .. }
                        | RuntimeHeapObjectKind::MemberMethod { .. }
                ))
                .count()
        );
        let actual = table
            .methods
            .iter()
            .map(|m| {
                assert_eq!(m.parameter_links.len(), m.parameter_types.len());
                format!(
                    "{}:{}:{}:{}:{}:{}:{}:{}",
                    m.member_id,
                    m.flags,
                    m.target.key,
                    m.target.context,
                    m.native_size,
                    typed(&m.return_type),
                    if m.parameter_types.is_empty() {
                        "-".into()
                    } else {
                        m.parameter_types
                            .iter()
                            .map(typed)
                            .collect::<Vec<_>>()
                            .join("/")
                    },
                    if m.default_data.is_empty() {
                        "-".into()
                    } else {
                        m.default_data
                            .iter()
                            .map(|v| format!("{v:02x}"))
                            .collect::<String>()
                    }
                )
            })
            .collect();
        assert_eq!(list(actual), row[5], "{}", row[0]);
        let actual = table
            .fields
            .iter()
            .map(|f| {
                f.as_ref().map_or_else(
                    || "null".into(),
                    |f| {
                        format!(
                            "{}:{}:{}",
                            f.member_id,
                            f.storage_offset,
                            typed(&f.value_type)
                        )
                    },
                )
            })
            .collect();
        assert_eq!(list(actual), row[6], "{}", row[0]);
        assert_eq!(
            list(
                table
                    .type_links
                    .iter()
                    .map(|l| format!("{}:{}", l.key, l.context))
                    .collect()
            ),
            row[7]
        );
        if !table.standard_module {
            assert!((0..table.methods.len()).all(|i| table.positional_arity(i).is_none()));
        }
        methods += table.methods.len();
        fields += table.fields.len();
        materialized += table.fields.iter().filter(|f| f.is_some()).count();
        count += 1;
    }
    assert_eq!((count, methods, fields, materialized), (292, 1740, 79, 27));
}

#[test]
fn native_named_invocation_shapes_follow_companion_signatures_not_target_code_or_member_ids() {
    let (mut queries, mut absent, mut count_errors, mut evaluated) = (0, 0, 0, 0);
    for row in cases().filter(|r| r[8] != "-") {
        let g = hex(row[2]);
        let p = hex(row[3]);
        let c = hex(row[4]);
        let index = row[1].parse::<usize>().unwrap();
        let project = decode_runtime_project_heap(&g, CacheContext::B5_WIN64).unwrap();
        let primary = decode_runtime_module_heap(&p, CacheContext::B5_WIN64).unwrap();
        let table =
            decode_runtime_invocation_table(&project, index, &primary, &c, CacheContext::B5_WIN64)
                .unwrap();
        for query in row[8].split(';') {
            let f = query.split(':').collect::<Vec<_>>();
            let name = hex(f[0]);
            let argc = f[1].parse::<usize>().unwrap();
            let ok = f[2] == "ok";
            let matches = (0..table.methods.len())
                .filter(|i| {
                    let n = project.procedure_name(index, *i).unwrap();
                    n.encoding == RuntimeNameEncoding::Ansi && n.bytes == name
                })
                .collect::<Vec<_>>();
            if matches.is_empty() {
                assert!(!ok);
                absent += 1;
            } else {
                assert_eq!(matches.len(), 1);
                let entry = matches[0];
                let shape = table.positional_arity(entry).unwrap();
                assert_eq!(shape.accepts(argc), ok, "{}: {} argc{argc}", row[0], f[0]);
                if !ok {
                    count_errors += 1;
                } else {
                    let image = primary
                        .standard_method_for_link(table.primary_stream, table.methods[entry].target)
                        .unwrap();
                    if argc > 0 && image.code.len() > 12 {
                        if let Ok(assignment) = decode_runtime_numeric_assignment(
                            &image.code[12..],
                            CacheContext::B5_WIN64,
                        ) {
                            let value = String::from_utf8(hex(f[3]))
                                .unwrap()
                                .parse::<f64>()
                                .unwrap();
                            assert_eq!(
                                assignment.evaluate_double(7.0).unwrap(),
                                value,
                                "{}: {}",
                                row[0],
                                f[0]
                            );
                            evaluated += 1;
                        }
                    }
                }
            }
            queries += 1;
        }
    }
    assert_eq!((queries, absent, count_errors), (109, 8, 15));
    assert!(evaluated >= 40);
}

#[test]
fn native_matrix_keeps_return_types_reference_modes_optional_data_and_paramarray_distinct() {
    let row = cases().find(|r| r[0] == "matrix-reopen").unwrap();
    let g = hex(row[2]);
    let p = hex(row[3]);
    let c = hex(row[4]);
    let project = decode_runtime_project_heap(&g, CacheContext::B5_WIN64).unwrap();
    let primary = decode_runtime_module_heap(&p, CacheContext::B5_WIN64).unwrap();
    let table =
        decode_runtime_invocation_table(&project, 2, &primary, &c, CacheContext::B5_WIN64).unwrap();
    assert_eq!(table.methods.len(), 15);
    assert_eq!(
        table.methods[..6]
            .iter()
            .map(|m| m.return_type.base())
            .collect::<Vec<_>>(),
        [0, 11, 8, 3, 16, 15]
    );
    assert!(!table.methods[8].parameter_types[0].is_by_ref());
    assert!(table.methods[9].parameter_types[0].is_by_ref());
    let optional = &table.methods[10];
    assert!(optional.parameter_types[0].is_optional());
    assert_eq!(optional.default_data, [3, 0, 5, 0, 0, 0]);
    let text = &table.methods[11];
    assert!(text.parameter_types[0].is_optional());
    assert_eq!(text.default_data, b"\x08\x00\x05\x00cache\x00");
    let shape = table.positional_arity(10).unwrap();
    assert_eq!((shape.required, shape.maximum), (0, Some(1)));
    let variadic = &table.methods[12];
    assert_eq!(variadic.flags & 0xfc00, 0xfc00);
    assert!(variadic.parameter_types[0].is_array());
    assert_eq!(variadic.parameter_types[0].base(), 15);
    let shape = table.positional_arity(12).unwrap();
    assert_eq!((shape.required, shape.maximum), (0, None));
    assert!((0..=3).all(|n| shape.accepts(n)));
    let array = &table.methods[13].parameter_types[0];
    assert!(array.is_array() && array.is_by_ref());
    assert_eq!((array.base(), array.array_word), (8, Some(0)));
}

#[test]
fn cached_entry_order_survives_member_id_changes_and_keeps_target_signature_independent() {
    for name in [
        "probe-ordered-reopen-companion-member-ids",
        "arity-ordered-companion-key-swap",
    ] {
        let row = cases().find(|r| r[0] == name).unwrap();
        let g = hex(row[2]);
        let p = hex(row[3]);
        let c = hex(row[4]);
        let project = decode_runtime_project_heap(&g, CacheContext::B5_WIN64).unwrap();
        let primary = decode_runtime_module_heap(&p, CacheContext::B5_WIN64).unwrap();
        let table =
            decode_runtime_invocation_table(&project, 2, &primary, &c, CacheContext::B5_WIN64)
                .unwrap();
        assert_eq!(project.procedure_name(2, 1).unwrap().bytes, b"Alpha");
        if name.contains("member-ids") {
            assert_eq!(table.methods[1].member_id & 0xffff, 3);
            let RuntimeHeapObjectKind::StandardMethod { code_slot, .. } = primary
                .object_for_link(2, table.methods[1].target)
                .unwrap()
                .kind
            else {
                panic!()
            };
            assert_eq!(code_slot, 1);
        } else {
            assert_eq!(table.positional_arity(1).unwrap().maximum, Some(1));
            let RuntimeHeapObjectKind::StandardMethod {
                code_slot,
                signature,
            } = primary
                .object_for_link(2, table.methods[1].target)
                .unwrap()
                .kind
            else {
                panic!()
            };
            assert_eq!(code_slot, 2);
            let RuntimeHeapObjectKind::Signature {
                parameter_types, ..
            } = primary.object_for_link(2, signature).unwrap().kind
            else {
                panic!()
            };
            assert_eq!(parameter_types.len(), 2);
        }
    }
}

#[test]
fn companion_metadata_rejects_every_truncation_and_trailing_bytes() {
    for name in [
        "matrix/SaveProbe",
        "classes/CacheClassFields",
        "records/SaveProbe",
    ] {
        let row = cases().find(|r| r[0] == name).unwrap();
        let p = hex(row[3]);
        let mut c = hex(row[4]);
        let primary = decode_runtime_module_heap(&p, CacheContext::B5_WIN64).unwrap();
        let stream = (u32::from_le_bytes(c[58..62].try_into().unwrap()) >> 16) as u16;
        for end in 0..c.len() {
            assert!(
                decode_runtime_companion_cache(&c[..end], CacheContext::B5_WIN64, &primary, stream)
                    .is_err(),
                "{name}: {end}"
            );
        }
        c.push(0);
        assert!(
            decode_runtime_companion_cache(&c, CacheContext::B5_WIN64, &primary, stream).is_err()
        );
    }
}

#[test]
fn companion_metadata_rejects_corrupt_counts_sizes_owners_targets_and_contexts() {
    let row = cases().find(|r| r[0] == "matrix/SaveProbe").unwrap();
    let p = hex(row[3]);
    let c = hex(row[4]);
    let mut primary = decode_runtime_module_heap(&p, CacheContext::B5_WIN64).unwrap();
    let table = decode_runtime_companion_cache(&c, CacheContext::B5_WIN64, &primary, 2).unwrap();
    let method = table.methods[10].serialized_offset;
    let defaults = table.methods[10].default_data.as_ptr() as usize - c.as_ptr() as usize;
    for (at, bytes) in [
        (2, u64::MAX.to_le_bytes().to_vec()),
        (58, vec![0; 4]),
        (84, vec![1, 0]),
        (method + 2, u64::MAX.to_le_bytes().to_vec()),
        (defaults - 4, u32::MAX.to_le_bytes().to_vec()),
    ] {
        let mut bad = c.clone();
        bad[at..at + bytes.len()].copy_from_slice(&bytes);
        assert!(
            decode_runtime_companion_cache(&bad, CacheContext::B5_WIN64, &primary, 2).is_err(),
            "{at}"
        );
    }
    assert!(decode_runtime_companion_cache(&c, CacheContext::B5_WIN64, &primary, 4).is_err());
    assert!(decode_runtime_companion_cache(
        &c,
        CacheContext {
            version: 0xb4,
            sys_kind: 3
        },
        &primary,
        2
    )
    .is_err());
    let at = primary
        .objects
        .iter()
        .position(|o| matches!(o.kind, RuntimeHeapObjectKind::StandardMethod { .. }))
        .unwrap();
    primary.objects.remove(at);
    assert!(decode_runtime_companion_cache(&c, CacheContext::B5_WIN64, &primary, 2).is_err());
}

#[test]
fn invocation_table_rejects_mismatched_project_pairings_and_exposes_only_standard_arity() {
    let row = cases().find(|r| r[0] == "matrix-reopen").unwrap();
    let g = hex(row[2]);
    let p = hex(row[3]);
    let c = hex(row[4]);
    let mut project = decode_runtime_project_heap(&g, CacheContext::B5_WIN64).unwrap();
    let mut primary = decode_runtime_module_heap(&p, CacheContext::B5_WIN64).unwrap();
    let mut table =
        decode_runtime_invocation_table(&project, 2, &primary, &c, CacheContext::B5_WIN64).unwrap();
    assert!(table.positional_arity(usize::MAX).is_none());
    table.methods[8].parameter_types[0].word |= 0x100;
    table.methods[8].parameter_types.push(RuntimeCompanionType {
        word: 11,
        array_word: None,
        reference: None,
    });
    assert_eq!(table.positional_arity(8).unwrap().required, 2);
    assert!(decode_runtime_invocation_table(
        &project,
        usize::MAX,
        &primary,
        &c,
        CacheContext::B5_WIN64
    )
    .is_err());
    assert!(
        decode_runtime_invocation_table(&project, 0, &primary, &c, CacheContext::B5_WIN64).is_err()
    );
    primary.module.companion_stream = 99;
    assert!(
        decode_runtime_invocation_table(&project, 2, &primary, &c, CacheContext::B5_WIN64).is_err()
    );
    primary.module.companion_stream = 3;
    let RuntimeProjectModuleState::Ready(ready) = &mut project.modules[2].state else {
        panic!()
    };
    ready.procedure_names.pop();
    assert!(
        decode_runtime_invocation_table(&project, 2, &primary, &c, CacheContext::B5_WIN64).is_err()
    );
}
