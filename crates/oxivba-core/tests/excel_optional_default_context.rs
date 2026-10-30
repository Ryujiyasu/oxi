// This Source Code Form is subject to the terms of the Mozilla Public
// License, v. 2.0. If a copy of the MPL was not distributed with this
// file, You can obtain one at https://mozilla.org/MPL/2.0/.

use oxivba_core::{parse_project, Host, ObjectRef, Runtime, Value};

struct ApplicationHost;
impl Host for ApplicationHost {
    fn call(
        &mut self,
        receiver: Option<&ObjectRef>,
        name: &str,
        _args: &[Value],
    ) -> Result<Option<Value>, String> {
        Ok(
            if receiver.is_none() && name.eq_ignore_ascii_case("Application") {
                Some(Value::Object(ObjectRef::new(1, "Application")))
            } else {
                None
            },
        )
    }
    fn get(&mut self, _receiver: &ObjectRef, _name: &str) -> Result<Option<Value>, String> {
        Ok(None)
    }
    fn set(&mut self, _receiver: &ObjectRef, _name: &str, _value: Value) -> Result<bool, String> {
        Ok(false)
    }
}

fn verify_context(source: &str, expected: &str) {
    let (module, classes) = parse_project(source).unwrap();
    let mut mismatches = Vec::new();
    for line in expected
        .lines()
        .filter(|s| !s.starts_with('#') && !s.starts_with("procedure\t"))
    {
        let row = line.split('\t').collect::<Vec<_>>();
        let expected = if row[1] == "string" {
            Value::String(row[2].into())
        } else {
            Value::Integer(row[2].parse().unwrap())
        };
        let mut host = ApplicationHost;
        let mut runtime = Runtime::new(&module)
            .with_classes(&classes)
            .with_host(&mut host);
        let actual = runtime.call_macro(row[0], vec![]);
        if actual.as_ref().ok() != Some(&expected) {
            mismatches.push(format!("{}: {actual:?}, expected {expected:?}", row[0]));
        }
    }
    assert!(mismatches.is_empty(), "{}", mismatches.join("\n"));
}

#[test]
fn optional_defaults_keep_native_internal_and_automation_contexts_separate() {
    verify_context(
        include_str!("../../../tests/fixtures/vba/optional_default_context.bas"),
        include_str!("../../../tests/fixtures/vba/optional_default_context.tsv"),
    );
}

#[test]
fn forwarded_missing_uses_byref_defaults_without_changing_the_caller_variant() {
    verify_context(
        include_str!("../../../tests/fixtures/vba/optional_default_reference_context.bas"),
        include_str!("../../../tests/fixtures/vba/optional_default_reference_context.tsv"),
    );
}

#[test]
fn native_optional_primitive_and_multiple_defaults_match_excel() {
    let (module, classes) = parse_project(include_str!(
        "../../../tests/fixtures/vba/optional_default_types.bas"
    ))
    .unwrap();
    let mut mismatches = Vec::new();
    for line in include_str!("../../../tests/fixtures/vba/optional_default_types.tsv")
        .lines()
        .filter(|l| !l.starts_with('#') && !l.starts_with("procedure\t"))
    {
        let f = line.split('\t').collect::<Vec<_>>();
        let expected = match f[1] {
            "string" => Value::String(f[2].into()),
            "bool" => Value::Boolean(f[2] == "true"),
            "double" => Value::Double(f[2].parse().unwrap()),
            _ => Value::Integer(f[2].parse().unwrap()),
        };
        let mut host = ApplicationHost;
        let mut runtime = Runtime::new(&module)
            .with_classes(&classes)
            .with_host(&mut host);
        let actual = runtime.call_macro(f[0], vec![]);
        if actual.as_ref().ok() != Some(&expected) {
            mismatches.push(format!("{}: {actual:?}, expected {expected:?}", f[0]));
        }
    }
    assert!(mismatches.is_empty(), "{}", mismatches.join("\n"));
}
