// This Source Code Form is subject to the terms of the Mozilla Public
// License, v. 2.0. If a copy of the MPL was not distributed with this
// file, You can obtain one at https://mozilla.org/MPL/2.0/.

use oxivba_core::{native_cache::*, RuntimeErrorKind, Value};
const CTX: CacheContext = CacheContext::B5_WIN64;
fn hex(s: &str) -> Vec<u8> {
    s.as_bytes()
        .chunks_exact(2)
        .map(|b| u8::from_str_radix(std::str::from_utf8(b).unwrap(), 16).unwrap())
        .collect()
}
fn cases() -> impl Iterator<Item = Vec<&'static str>> {
    include_str!("../../../tests/fixtures/vba/native_cache_frames.tsv")
        .lines()
        .filter(|l| !l.starts_with('#') && !l.starts_with("id\t"))
        .map(|l| l.split('\t').collect())
}
fn arguments(s: &str) -> Vec<RuntimeScalarArgument> {
    if s == "-" {
        return vec![];
    }
    s.split(',')
        .map(|v| {
            RuntimeScalarArgument::Value(if let Some(s) = v.strip_prefix("b:") {
                Value::Boolean(s == "true")
            } else {
                let s = v.strip_prefix("n:").unwrap();
                if s.contains('.') {
                    Value::Double(s.parse().unwrap())
                } else {
                    let n = s.parse::<i64>().unwrap();
                    if i32::try_from(n).is_ok() {
                        Value::Integer(n)
                    } else {
                        Value::LongLong(n)
                    }
                }
            })
        })
        .collect()
}
fn expected(kind: &str, text: &str) -> Value {
    match kind {
        "byte" => Value::Byte(text.parse().unwrap()),
        "i16" => Value::Int16(text.parse().unwrap()),
        "long" => Value::Integer(text.parse().unwrap()),
        "i64" => Value::LongLong(text.parse().unwrap()),
        "f32" => Value::Single(text.parse().unwrap()),
        "f64" => Value::Double(text.parse().unwrap()),
        "cy" => Value::Currency(text.parse().unwrap()),
        "date" => Value::Date(text.parse().unwrap()),
        "bool" => Value::Boolean(text == "true"),
        "empty" => Value::Empty,
        _ => panic!("kind {kind}"),
    }
}
fn cleanup(t: &RuntimeCleanupTable<'_>) -> String {
    if t.slots.is_empty() {
        "-".into()
    } else {
        t.slots
            .iter()
            .map(|s| format!("{}:{}", s.offset, s.kind))
            .collect::<Vec<_>>()
            .join(",")
    }
}

#[test]
fn native_frames_preserve_argument_locals_and_both_cleanup_tables() {
    let mut count = 0;
    for row in cases() {
        let bytes = hex(row[2]);
        let primary = decode_runtime_module_heap(&bytes, CTX).unwrap();
        for text in row[5].split(';') {
            let f = text.split('~').collect::<Vec<_>>();
            let image = &primary.module.methods[f[0].parse::<usize>().unwrap()];
            let frame = decode_runtime_frame_layout(image, CTX).unwrap();
            assert_eq!(frame.argument_bytes, f[1].parse::<u16>().unwrap());
            assert_eq!(frame.local_bytes, f[2].parse::<u16>().unwrap());
            assert_eq!(frame.code_bytes as usize, image.code.len());
            assert_eq!(frame.consumed, image.frame_info.len());
            assert_eq!(cleanup(&frame.first_cleanup), f[3]);
            assert_eq!(cleanup(&frame.second_cleanup), f[4]);
            count += 1;
        }
    }
    assert_eq!(count, 104);
}

#[test]
fn native_scalar_functions_execute_complete_cached_images_without_source() {
    let mut count = 0;
    for row in cases() {
        let p = hex(row[2]);
        let c = hex(row[3]);
        let primary = decode_runtime_module_heap(&p, CTX).unwrap();
        let table = decode_runtime_companion_cache(&c, CTX, &primary, 2).unwrap();
        for line in include_str!("../../../tests/fixtures/vba/native_scalar_calls.tsv")
            .lines()
            .filter(|l| !l.starts_with('#') && !l.starts_with("id\t"))
        {
            let f = line.split('\t').collect::<Vec<_>>();
            if f[0] != row[0] {
                continue;
            }
            let result = execute_runtime_scalar_method(
                &primary,
                &table,
                f[1].parse().unwrap(),
                &arguments(f[3]),
                CTX,
                1000,
            );
            assert_eq!(
                result
                    .as_ref()
                    .unwrap_or_else(|e| panic!("{} {}: {e}", row[0], f[2])),
                &expected(f[4], f[5]),
                "{} {}",
                row[0],
                f[2]
            );
            count += 1;
        }
    }
    assert_eq!(count, 106);
}

#[test]
fn native_project_link_pair_zero_word_does_not_drop_its_two_links() {
    for row in cases() {
        let g = hex(row[1]);
        let p = hex(row[2]);
        let c = hex(row[3]);
        let project = decode_runtime_project_heap(&g, CTX).unwrap();
        let primary = decode_runtime_module_heap(&p, CTX).unwrap();
        assert!(project.objects.iter().any(|o|matches!(&o.kind,RuntimeProjectObjectKind::LinkArray{word:0,entries,..} if entries.len()==2)));
        let index = project
            .modules
            .iter()
            .position(
                |m| matches!(&m.state,RuntimeProjectModuleState::Ready(r) if r.root.context==2<<16),
            )
            .unwrap();
        let table = decode_runtime_invocation_table(&project, index, &primary, &c, CTX).unwrap();
        assert_eq!(table.methods.len(), row[4].split(',').count());
        for (i, name) in row[4].split(',').enumerate() {
            assert_eq!(
                project.procedure_name(index, i).unwrap().bytes,
                name.as_bytes()
            );
        }
    }
}

#[test]
fn native_byref_updates_aliases_and_numeric_errors_keep_their_kind() {
    let row = cases().nth(1).unwrap();
    let p = hex(row[2]);
    let c = hex(row[3]);
    let primary = decode_runtime_module_heap(&p, CTX).unwrap();
    let table = decode_runtime_companion_cache(&c, CTX, &primary, 2).unwrap();
    let entry = |s| row[4].split(',').position(|n| n == s).unwrap();
    let reference =
        RuntimeScalarReference::new(RuntimeScalarType::Double, Value::Double(7.0)).unwrap();
    let value = execute_runtime_scalar_method(
        &primary,
        &table,
        entry("Mutate"),
        &[RuntimeScalarArgument::Reference(reference.clone())],
        CTX,
        1000,
    )
    .unwrap();
    assert_eq!(value, Value::Double(27.0));
    assert_eq!(reference.value(), Value::Double(9.0));
    let failure = execute_runtime_scalar_method(
        &primary,
        &table,
        entry("Divide"),
        &arguments("n:7"),
        CTX,
        1000,
    )
    .unwrap_err();
    assert!(
        matches!(failure,RuntimeScalarError::Runtime(e) if e.kind==RuntimeErrorKind::DivisionByZero)
    );
    let wrong = RuntimeScalarReference::new(RuntimeScalarType::Long, Value::Integer(7)).unwrap();
    assert!(matches!(
        execute_runtime_scalar_method(
            &primary,
            &table,
            entry("Mutate"),
            &[RuntimeScalarArgument::Reference(wrong)],
            CTX,
            1000
        ),
        Err(RuntimeScalarError::ByRefTypeMismatch)
    ));
    assert!(matches!(
        execute_runtime_scalar_method(&primary, &table, entry("Two"), &[], CTX, 1000),
        Err(RuntimeScalarError::Arity)
    ));
    assert!(matches!(
        execute_runtime_scalar_method(&primary, &table, entry("ByValString"), &[], CTX, 1000),
        Err(RuntimeScalarError::Arity)
    ));
    assert!(matches!(
        execute_runtime_scalar_method(
            &primary,
            &table,
            entry("LocalRefs"),
            &[RuntimeScalarArgument::Value(Value::String("cache".into()))],
            CTX,
            1000
        ),
        Err(RuntimeScalarError::Unsupported(_))
    ));
}

#[test]
fn frame_reader_checks_every_prefix_counts_sizes_and_image_pairing() {
    let row = cases().next().unwrap();
    let p = hex(row[2]);
    let primary = decode_runtime_module_heap(&p, CTX).unwrap();
    for image in &primary.module.methods {
        for end in 0..image.frame_info.len() {
            let mut short = image.clone();
            short.frame_info = &image.frame_info[..end];
            assert!(decode_runtime_frame_layout(&short, CTX).is_err());
        }
        for at in [12, 16, 36, 46] {
            let mut bytes = image.frame_info.to_vec();
            bytes[at] ^= 1;
            let mut bad = image.clone();
            bad.frame_info = &bytes;
            assert!(
                decode_runtime_frame_layout(&bad, CTX).is_err(),
                "field {at}"
            );
        }
        assert!(decode_runtime_frame_layout(
            image,
            CacheContext {
                version: 180,
                sys_kind: 3
            }
        )
        .is_err());
    }
}

#[test]
fn cached_scalar_binding_uses_defaults_code_targets_and_typed_references() {
    let row = cases().nth(1).unwrap();
    let p = hex(row[2]);
    let c = hex(row[3]);
    let primary = decode_runtime_module_heap(&p, CTX).unwrap();
    let mut table = decode_runtime_companion_cache(&c, CTX, &primary, 2).unwrap();
    let entry = |s| row[4].split(',').position(|n| n == s).unwrap();
    let optional = entry("OptionalNumber");
    assert_eq!(
        execute_runtime_scalar_method(
            &primary,
            &table,
            optional,
            &[RuntimeScalarArgument::Value(Value::Missing)],
            CTX,
            1000
        )
        .unwrap(),
        Value::Double(7.0)
    );
    assert_eq!(
        execute_runtime_scalar_method(
            &primary,
            &table,
            optional,
            &[RuntimeScalarArgument::Value(Value::Empty)],
            CTX,
            1000
        )
        .unwrap(),
        Value::Double(0.0)
    );
    let reference =
        RuntimeScalarReference::new(RuntimeScalarType::Double, Value::Double(7.0)).unwrap();
    assert_eq!(
        execute_runtime_scalar_method(
            &primary,
            &table,
            entry("ByValDouble"),
            &[RuntimeScalarArgument::Reference(reference.clone())],
            CTX,
            1000
        )
        .unwrap(),
        Value::Double(7.0)
    );
    assert_eq!(reference.value(), Value::Double(7.0));
    // A canonical entry's target is independent of its position and member ID.
    table.methods[entry("ByValDouble")].target = table.methods[entry("Branch")].target;
    table.methods[entry("ByValDouble")].member_id = u32::MAX;
    assert_eq!(
        execute_runtime_scalar_method(
            &primary,
            &table,
            entry("ByValDouble"),
            &arguments("n:7"),
            CTX,
            1000
        )
        .unwrap(),
        Value::Double(8.0)
    );
    // Unequal target arities exist in Excel. Until their physical frame is
    // modeled, fail explicitly instead of inventing a value for absent storage.
    table.methods[entry("ByValDouble")].target = table.methods[entry("Two")].target;
    assert!(matches!(
        execute_runtime_scalar_method(
            &primary,
            &table,
            entry("ByValDouble"),
            &arguments("n:7"),
            CTX,
            1000
        ),
        Err(RuntimeScalarError::Unsupported("scalar frame ABI"))
    ));
}

#[test]
fn invalid_cached_code_is_rejected_before_byref_writes_and_cycles_are_bounded() {
    let row = cases().nth(1).unwrap();
    let p = hex(row[2]);
    let c = hex(row[3]);
    let mut primary = decode_runtime_module_heap(&p, CTX).unwrap();
    let table = decode_runtime_companion_cache(&c, CTX, &primary, 2).unwrap();
    let entry = |s| row[4].split(',').position(|n| n == s).unwrap();
    let slot = |p: &RuntimeModuleHeap<'_>, i: usize| match p
        .object_for_link(2, table.methods[i].target)
        .unwrap()
        .kind
    {
        RuntimeHeapObjectKind::StandardMethod { code_slot, .. } => usize::from(code_slot),
        _ => panic!(),
    };
    let mutate_slot = slot(&primary, entry("Mutate"));
    let mut bad = primary.module.methods[mutate_slot].code.to_vec();
    let return_at = bad.windows(2).rposition(|b| b == [0x73, 0x02]).unwrap();
    bad[return_at..return_at + 2].copy_from_slice(&[0xff, 0xff]);
    primary.module.methods[mutate_slot].code = &bad;
    let reference =
        RuntimeScalarReference::new(RuntimeScalarType::Double, Value::Double(7.0)).unwrap();
    assert!(execute_runtime_scalar_method(
        &primary,
        &table,
        entry("Mutate"),
        &[RuntimeScalarArgument::Reference(reference.clone())],
        CTX,
        1000
    )
    .is_err());
    assert_eq!(reference.value(), Value::Double(7.0));
    let branch_slot = slot(&primary, entry("Branch"));
    let mut cycle = primary.module.methods[branch_slot].code.to_vec();
    let jump = cycle.windows(2).position(|b| b == [0xc6, 0x02]).unwrap();
    cycle[jump + 2..jump + 6].copy_from_slice(&0_u32.to_le_bytes());
    primary.module.methods[branch_slot].code = &cycle;
    assert!(matches!(
        execute_runtime_scalar_method(
            &primary,
            &table,
            entry("Branch"),
            &arguments("n:2"),
            CTX,
            100
        ),
        Err(RuntimeScalarError::InstructionLimit)
    ));
    let mut malformed = cycle.clone();
    malformed[jump + 2..jump + 6].copy_from_slice(&8_u32.to_le_bytes());
    primary.module.methods[branch_slot].code = &malformed;
    assert!(matches!(
        execute_runtime_scalar_method(
            &primary,
            &table,
            entry("Branch"),
            &arguments("n:2"),
            CTX,
            100
        ),
        Err(RuntimeScalarError::InvalidCode(_))
    ));
}

#[test]
fn native_shared_byref_arguments_keep_one_storage_location() {
    let row = cases().nth(2).unwrap();
    let p = hex(row[2]);
    let c = hex(row[3]);
    let primary = decode_runtime_module_heap(&p, CTX).unwrap();
    let table = decode_runtime_companion_cache(&c, CTX, &primary, 2).unwrap();
    let entry = row[4].split(',').position(|n| n == "Twin").unwrap();
    let first = RuntimeScalarReference::new(RuntimeScalarType::Double, Value::Double(7.0)).unwrap();
    let second =
        RuntimeScalarReference::new(RuntimeScalarType::Double, Value::Double(4.0)).unwrap();
    let call = |a: RuntimeScalarReference, b: RuntimeScalarReference| {
        execute_runtime_scalar_method(
            &primary,
            &table,
            entry,
            &[
                RuntimeScalarArgument::Reference(a),
                RuntimeScalarArgument::Reference(b),
            ],
            CTX,
            1000,
        )
        .unwrap()
    };
    assert_eq!(call(first.clone(), second.clone()), Value::Double(21.0));
    assert_eq!(first.value(), Value::Double(9.0));
    assert_eq!(second.value(), Value::Double(12.0));
    let shared =
        RuntimeScalarReference::new(RuntimeScalarType::Double, Value::Double(7.0)).unwrap();
    assert_eq!(call(shared.clone(), shared.clone()), Value::Double(54.0));
    assert_eq!(shared.value(), Value::Double(27.0));
    // Excel's AliasState returns 81: result 54 plus the shared caller value 27.
}
