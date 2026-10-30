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
    include_str!("../../../tests/fixtures/vba/native_cache_calls.tsv")
        .lines()
        .filter(|l| !l.starts_with('#') && !l.starts_with("id\t"))
        .map(|l| l.split('\t').collect())
}
fn entry(row: &[&str], name: &str) -> usize {
    row[4].split(',').position(|n| n == name).unwrap()
}
fn args(s: &str) -> Vec<RuntimeScalarArgument> {
    if s == "-" {
        return vec![];
    }
    s.split(',')
        .map(|s| {
            RuntimeScalarArgument::Value(if let Some(b) = s.strip_prefix("b:") {
                Value::Boolean(b == "true")
            } else {
                let n = s.strip_prefix("n:").unwrap();
                if n.contains('.') {
                    Value::Double(n.parse().unwrap())
                } else {
                    let v = n.parse::<i64>().unwrap();
                    if i32::try_from(v).is_ok() {
                        Value::Integer(v)
                    } else {
                        Value::LongLong(v)
                    }
                }
            })
        })
        .collect()
}
fn expected(k: &str, s: &str) -> Value {
    match k {
        "byte" => Value::Byte(s.parse().unwrap()),
        "i16" => Value::Int16(s.parse().unwrap()),
        "long" => Value::Integer(s.parse().unwrap()),
        "i64" => Value::LongLong(s.parse().unwrap()),
        "f32" => Value::Single(s.parse().unwrap()),
        "f64" => Value::Double(s.parse().unwrap()),
        "cy" => Value::Currency(s.parse().unwrap()),
        "date" => Value::Date(s.parse().unwrap()),
        "bool" => Value::Boolean(s == "true"),
        _ => panic!("{k}"),
    }
}

#[test]
fn native_local_calls_match_typed_nested_recursive_and_reopened_outcomes() {
    let mut count = 0;
    for row in cases() {
        let p = hex(row[2]);
        let c = hex(row[3]);
        let primary = decode_runtime_module_heap(&p, CTX).unwrap();
        let table = decode_runtime_companion_cache(&c, CTX, &primary, 2).unwrap();
        for line in include_str!("../../../tests/fixtures/vba/native_cached_call_profiles.tsv")
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
                &args(f[3]),
                CTX,
                10000,
            );
            assert_eq!(
                result.unwrap_or_else(|e| panic!("{} {}: {e}", f[0], f[2])),
                expected(f[4], f[5]),
                "{} {}",
                f[0],
                f[2]
            );
            count += 1;
        }
    }
    assert_eq!(count, 132);
}

#[test]
fn canonical_entry_names_and_primary_call_targets_are_independent() {
    for row in cases() {
        let g = hex(row[1]);
        let p = hex(row[2]);
        let c = hex(row[3]);
        let project = decode_runtime_project_heap(&g, CTX).unwrap();
        let primary = decode_runtime_module_heap(&p, CTX).unwrap();
        let index = project
            .modules
            .iter()
            .position(
                |m| matches!(&m.state,RuntimeProjectModuleState::Ready(r)if r.root.context==2<<16),
            )
            .unwrap();
        let table = decode_runtime_invocation_table(&project, index, &primary, &c, CTX).unwrap();
        for (i, name) in row[4].split(',').enumerate() {
            assert_eq!(
                project.procedure_name(index, i).unwrap().bytes,
                name.as_bytes()
            );
        }
        if row[0] == "intervention" {
            assert_eq!(
                execute_runtime_scalar_method(
                    &primary,
                    &table,
                    entry(&row, "MutationState"),
                    &[],
                    CTX,
                    1000
                )
                .unwrap(),
                Value::Double(14.0)
            );
            assert_eq!(
                execute_runtime_scalar_method(
                    &primary,
                    &table,
                    entry(&row, "Mutate"),
                    &args("n:7"),
                    CTX,
                    1000
                )
                .unwrap(),
                Value::Double(27.0)
            );
        }
    }
}

#[test]
fn compiled_callsite_defaults_do_not_use_changed_external_defaults() {
    let row = cases().find(|r| r[0] == "default_intervention").unwrap();
    let p = hex(row[2]);
    let c = hex(row[3]);
    let primary = decode_runtime_module_heap(&p, CTX).unwrap();
    let mut table = decode_runtime_companion_cache(&c, CTX, &primary, 2).unwrap();
    let call = |table: &RuntimeCompanionCache<'_>, name: &str| {
        execute_runtime_scalar_method(&primary, table, entry(&row, name), &[], CTX, 1000).unwrap()
    };
    assert_eq!(call(&table, "OptionalNumber"), Value::Double(12.5));
    assert_eq!(call(&table, "OptionalCall"), Value::Double(7.0));
    assert_eq!(call(&table, "LiteralCall"), Value::Double(3.5));
    // External metadata can disappear/change without renaming the primary target.
    table.methods[entry(&row, "OptionalNumber")].default_data = &[];
    assert_eq!(call(&table, "OptionalNumber"), Value::Double(0.0));
    assert_eq!(call(&table, "OptionalCall"), Value::Double(7.0));
}

#[test]
fn forwarded_and_temporary_references_keep_aliases_and_error_side_effects() {
    let row = cases().find(|r| r[0] == "order").unwrap();
    let p = hex(row[2]);
    let c = hex(row[3]);
    let primary = decode_runtime_module_heap(&p, CTX).unwrap();
    let table = decode_runtime_companion_cache(&c, CTX, &primary, 2).unwrap();
    let reference =
        RuntimeScalarReference::new(RuntimeScalarType::Double, Value::Double(7.0)).unwrap();
    let value = execute_runtime_scalar_method(
        &primary,
        &table,
        entry(&row, "Forward"),
        &[RuntimeScalarArgument::Reference(reference.clone())],
        CTX,
        1000,
    )
    .unwrap();
    assert_eq!(value, Value::Double(27.0));
    assert_eq!(reference.value(), Value::Double(9.0));
    let reference =
        RuntimeScalarReference::new(RuntimeScalarType::Double, Value::Double(7.0)).unwrap();
    assert!(
        matches!(execute_runtime_scalar_method(&primary,&table,entry(&row,"NestedFail"),&[RuntimeScalarArgument::Reference(reference.clone())],CTX,1000),Err(RuntimeScalarError::Runtime(e))if e.kind==RuntimeErrorKind::DivisionByZero)
    );
    assert_eq!(reference.value(), Value::Double(9.0)); // Native CaughtState returns9.
    for (name, want) in [
        ("TemporaryRef", 27.0),
        ("UninitializedRef", 6.0),
        ("AliasState", 81.0),
    ] {
        assert_eq!(
            execute_runtime_scalar_method(&primary, &table, entry(&row, name), &[], CTX, 1000)
                .unwrap(),
            Value::Double(want)
        );
    }
}

#[test]
fn recursive_calls_share_the_instruction_budget_without_rust_recursion() {
    let row = cases().find(|r| r[0] == "order").unwrap();
    let p = hex(row[2]);
    let c = hex(row[3]);
    let primary = decode_runtime_module_heap(&p, CTX).unwrap();
    let table = decode_runtime_companion_cache(&c, CTX, &primary, 2).unwrap();
    let recursive = entry(&row, "Recursive");
    assert_eq!(
        execute_runtime_scalar_method(&primary, &table, recursive, &args("n:3"), CTX, 1000)
            .unwrap(),
        Value::Double(6.0)
    );
    assert!(matches!(
        execute_runtime_scalar_method(&primary, &table, recursive, &args("n:100000"), CTX, 100),
        Err(RuntimeScalarError::InstructionLimit)
    ));
}

#[test]
fn invalid_call_targets_and_abis_are_rejected_before_caller_side_effects() {
    let row = cases().find(|r| r[0] == "order").unwrap();
    let p = hex(row[2]);
    let c = hex(row[3]);
    let original = decode_runtime_module_heap(&p, CTX).unwrap();
    let table = decode_runtime_companion_cache(&c, CTX, &original, 2).unwrap();
    let target = table.methods[entry(&row, "Mutate")].target;
    let index = original
        .module
        .auxiliary_links
        .iter()
        .position(|l| *l == target)
        .unwrap();
    let mut primary = original.clone();
    primary.module.auxiliary_links[index] = RuntimeCacheLink {
        key: target.key,
        context: 0,
    };
    assert!(execute_runtime_scalar_method(
        &primary,
        &table,
        entry(&row, "ForwardState"),
        &[],
        CTX,
        1000
    )
    .is_err());
    let mut primary = original.clone();
    primary.module.auxiliary_links[index] = table.methods[entry(&row, "Two")].target;
    assert!(matches!(
        execute_runtime_scalar_method(
            &primary,
            &table,
            entry(&row, "ForwardState"),
            &[],
            CTX,
            1000
        ),
        Err(RuntimeScalarError::Unsupported("cached call ABI"))
    ));
    let mut primary = original.clone();
    let reference =
        RuntimeScalarReference::new(RuntimeScalarType::Double, Value::Double(7.0)).unwrap();
    assert!(execute_runtime_scalar_method(
        &primary,
        &table,
        entry(&row, "Forward"),
        &[RuntimeScalarArgument::Reference(reference.clone())],
        CTX,
        0
    )
    .is_err());
    assert_eq!(reference.value(), Value::Double(7.0));
    let target = table.methods[entry(&row, "Failing")].target;
    let RuntimeHeapObjectKind::StandardMethod { code_slot, .. } =
        primary.object_for_link(2, target).unwrap().kind
    else {
        panic!()
    };
    let image = &mut primary.module.methods[usize::from(code_slot)];
    let mut code = image.code.to_vec();
    let at = code.windows(2).rposition(|b| b == [0x73, 0x02]).unwrap();
    code[at..at + 2].copy_from_slice(&[0xff, 0xff]);
    image.code = &code;
    assert!(execute_runtime_scalar_method(
        &primary,
        &table,
        entry(&row, "NestedFail"),
        &[RuntimeScalarArgument::Reference(reference.clone())],
        CTX,
        1000
    )
    .is_err());
    assert_eq!(reference.value(), Value::Double(7.0));
}
