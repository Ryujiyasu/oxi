// This Source Code Form is subject to the terms of the Mozilla Public
// License, v. 2.0. If a copy of the MPL was not distributed with this
// file, You can obtain one at https://mozilla.org/MPL/2.0/.

use std::collections::{BTreeMap, BTreeSet};

use oxivba_core::ast::Visibility;
use oxivba_core::{
    analyse, host_error, host_error_described, host_error_from, parse_module, ErrorEnvironment,
    ErrorLanguage, Host, ModuleItem, ObjectRef, Runtime, RuntimeErrorKind, Value,
};

const HELP_FILE: &str = r"C:\Oxi\VBA\1041\VbLR6.chm";

fn environment() -> ErrorEnvironment {
    ErrorEnvironment {
        language: ErrorLanguage::Japanese,
        help_file: HELP_FILE.to_string(),
    }
}

#[test]
fn japanese_error_catalog_matches_all_excel_numbers() {
    let expected: BTreeMap<i64, &str> =
        include_str!("../../../tests/fixtures/vba/error_messages_ja.tsv")
            .lines()
            .map(|line| {
                let (number, description) = line.split_once('\t').unwrap();
                (number.parse().unwrap(), description)
            })
            .collect();
    assert_eq!(expected.len(), 89);
    let generic = expected[&95];
    let module = parse_module(
        "Public Function Describe(ByVal code As Long) As String\nDescribe = Error$(code)\nEnd Function",
    ).unwrap();
    assert_eq!(analyse(&module).metrics.unparsed, 0);
    let mut runtime = Runtime::new(&module).with_error_environment(environment());
    for code in 0..=65_535 {
        let description = expected.get(&code).copied().unwrap_or(generic);
        assert_eq!(
            runtime.call("Describe", vec![Value::Integer(code)]),
            Ok(Value::String(description.to_owned())),
            "Error$({code})"
        );
    }
}

#[test]
fn japanese_error_snapshots_match_excel() {
    let module = parse_module(include_str!(
        "../../../tests/fixtures/vba/localized_errors_ja.bas"
    ))
    .unwrap();
    assert_eq!(analyse(&module).metrics.unparsed, 0);
    let names: BTreeSet<_> = module
        .items
        .iter()
        .filter_map(|item| match item {
            ModuleItem::Procedure(procedure)
                if procedure.visibility == Visibility::Public
                    && !procedure.name.starts_with("Host") =>
            {
                Some(procedure.name.as_str())
            }
            _ => None,
        })
        .collect();
    let mut compared = BTreeSet::new();
    let mut failures = Vec::new();
    for case in include_str!("../../../tests/fixtures/vba/localized_errors_ja.tsv").lines() {
        let (name, expected) = case.split_once('\t').unwrap();
        if name.starts_with("Host") {
            continue;
        }
        assert!(compared.insert(name));
        let expected = Value::String(expected.replace("{VBA_HELP}", HELP_FILE));
        let actual = Runtime::new(&module)
            .with_error_environment(environment())
            .call(name, vec![]);
        if actual != Ok(expected.clone()) {
            failures.push(format!("{name}: expected {expected:?}, got {actual:?}"));
        }
    }
    assert_eq!(compared, names);
    assert!(failures.is_empty(), "{}", failures.join("\n"));
}

#[test]
fn numbered_host_errors_localize_and_explicit_descriptions_are_preserved() {
    struct Refusal(bool);
    impl Host for Refusal {
        fn call(
            &mut self,
            _: Option<&ObjectRef>,
            _: &str,
            _: &[Value],
        ) -> Result<Option<Value>, String> {
            Err(if self.0 {
                host_error_described(6, "Overflow")
            } else {
                host_error(6, "engine diagnostic")
            })
        }
        fn get(&mut self, _: &ObjectRef, _: &str) -> Result<Option<Value>, String> {
            Ok(None)
        }
        fn set(&mut self, _: &ObjectRef, _: &str, _: Value) -> Result<bool, String> {
            Ok(false)
        }
    }
    let module = parse_module("Function Main() As String\nOn Error Resume Next\nHostRefusal\nMain = Err.Description\nEnd Function").unwrap();
    for (explicit, expected) in [(false, "オーバーフローしました。"), (true, "Overflow")]
    {
        let mut host = Refusal(explicit);
        assert_eq!(
            Runtime::new(&module)
                .with_host(&mut host)
                .with_error_environment(environment())
                .call("Main", vec![]),
            Ok(Value::String(expected.to_owned()))
        );
    }
}

#[test]
fn explicit_host_sources_are_never_internal_markers() {
    struct Refusal(&'static str);
    impl Host for Refusal {
        fn call(
            &mut self,
            _: Option<&ObjectRef>,
            _: &str,
            _: &[Value],
        ) -> Result<Option<Value>, String> {
            Err(host_error_from(513, self.0, "custom"))
        }
        fn get(&mut self, _: &ObjectRef, _: &str) -> Result<Option<Value>, String> {
            Ok(None)
        }
        fn set(&mut self, _: &ObjectRef, _: &str, _: Value) -> Result<bool, String> {
            Ok(false)
        }
    }
    let module = parse_module("Function Main() As String\nOn Error Resume Next\nHostRefusal\nMain = Err.Source & \"|\" & Err.Description\nEnd Function").unwrap();
    for source in ["VBA", "Collection", "Microsoft Excel"] {
        let mut host = Refusal(source);
        assert_eq!(
            Runtime::new(&module)
                .with_host(&mut host)
                .with_error_environment(environment())
                .call("Main", vec![]),
            Ok(Value::String(format!("{source}|custom")))
        );
    }
}

#[test]
fn unhandled_custom_error_text_keeps_control_characters() {
    let module = parse_module(
        "Sub Main()\nErr.Raise 513, \"VBA\", \"before\" & Chr(1) & \"after\"\nEnd Sub",
    )
    .unwrap();
    let failure = Runtime::new(&module).call("Main", vec![]).unwrap_err();
    assert_eq!(failure.vba_source.as_deref(), Some("VBA"));
    assert_eq!(failure.message, "before\u{1}after");
    assert!(failure.to_string().contains("before\u{1}after"));
}

#[test]
fn renamed_project_sources_match_excel() {
    let module = parse_module(include_str!(
        "../../../tests/fixtures/vba/localized_errors_ja.bas"
    ))
    .unwrap();
    let mut compared = 0;
    for line in include_str!("../../../tests/fixtures/vba/project_sources_ja.tsv").lines() {
        let mut fields = line.splitn(3, '\t');
        let project = fields.next().unwrap();
        let procedure = fields.next().unwrap();
        let expected = fields.next().unwrap().replace("{VBA_HELP}", HELP_FILE);
        if procedure.starts_with("Host") {
            continue;
        }
        let actual = Runtime::new(&module)
            .with_project_name(project)
            .with_error_environment(environment())
            .call(procedure, vec![]);
        assert_eq!(actual, Ok(Value::String(expected)), "{project}.{procedure}");
        compared += 1;
    }
    assert_eq!(compared, 33);
}

#[test]
fn special_error_metadata_and_control_flow_match_excel() {
    let module = parse_module(include_str!("../../../tests/fixtures/vba/special_errors.bas")).unwrap();
    assert_eq!(analyse(&module).metrics.unparsed, 0);
    let mut compared = 0;
    for line in include_str!("../../../tests/fixtures/vba/special_errors_ja.tsv").lines() {
        let fields: Vec<_> = line.splitn(6, '\t').collect();
        assert_eq!(fields.len(), 6);
        let [procedure, number, mode, outcome, progress, expected] = fields.as_slice() else { unreachable!() };
        let mut args = vec![Value::Integer(number.parse().unwrap())];
        if *mode != "omitted" {
            args.push(Value::String(if *mode == "empty" { "" } else { mode }.to_string()));
        }
        let mut runtime = Runtime::new(&module).with_error_environment(environment());
        let actual = runtime.call(procedure, args);
        if *outcome == "terminate" {
            let failure = actual.unwrap_err();
            assert_eq!(failure.kind, RuntimeErrorKind::ExecutionTerminated, "{line}");
            let hresult: i32 = expected.parse().unwrap();
            assert_eq!(failure.vba_number, Some((hresult as u32 & 0xffff) as i64), "{line}");
        } else {
            assert_eq!(actual, Ok(Value::String(expected.replace("{VBA_HELP}", HELP_FILE))), "{line}");
        }
        assert_eq!(runtime.call("ReadProgress", vec![]), Ok(Value::String(progress.to_string())), "{line}");
        compared += 1;
    }
    assert_eq!(compared, 37);
}

#[test]
fn error_statement_and_raise_sweeps_match_all_excel_numbers() {
    let module = parse_module(include_str!("../../../tests/fixtures/vba/special_errors.bas")).unwrap();
    let mut captures: BTreeMap<&str, BTreeMap<&str, &str>> = BTreeMap::new();
    for line in include_str!("../../../tests/fixtures/vba/special_error_sweeps_ja.tsv").lines() {
        let mut fields = line.splitn(3, '\t');
        let kind = fields.next().unwrap();
        let number = fields.next().unwrap();
        let expected = fields.next().unwrap();
        assert!(captures.entry(kind).or_default().insert(number, expected).is_none());
    }
    assert_eq!(captures.len(), 2);
    for (kind, use_raise) in [("statement", false), ("raise", true)] {
        let Value::Array(values) = Runtime::new(&module)
            .with_error_environment(environment())
            .call("Sweep", vec![Value::Boolean(use_raise)]).unwrap() else { panic!("{kind} array result") };
        assert_eq!(values.lower_bound(1), Some(0));
        assert_eq!(values.values.len(), 65_536);
        let expected = &captures[kind];
        for (number, actual) in values.values.iter().enumerate() {
            let number = number.to_string();
            let snapshot = expected.get(number.as_str()).unwrap_or(&expected["*"])
                .replace("{NUMBER}", &number);
            assert_eq!(actual, &Value::String(snapshot), "{kind} {number}");
        }
    }
}

#[test]
fn err_automation_dispatch_matches_excel() {
    let module = parse_module(include_str!("../../../tests/fixtures/vba/err_dispatch.bas")).unwrap();
    assert_eq!(analyse(&module).metrics.unparsed, 0);
    let mut compared = 0;
    let mut failures = Vec::new();
    for line in include_str!("../../../tests/fixtures/vba/err_dispatch_ja.tsv").lines() {
        let fields: Vec<_> = line.splitn(4, '\t').collect();
        let args = vec![Value::String(fields[0].to_string()), Value::Integer(fields[1].parse().unwrap()), Value::Integer(fields[2].parse().unwrap())];
        let expected = Value::String(fields[3].replace("{VBA_HELP}", HELP_FILE));
        let actual = Runtime::new(&module).with_error_environment(environment()).call("DispatchMatrix", args);
        if actual != Ok(expected.clone()) {
            failures.push(format!("{}: expected {expected:?}, got {actual:?}", &line[..line.find('\t').unwrap()]));
        }
        compared += 1;
    }
    assert_eq!(compared, 420);
    assert!(failures.is_empty(), "{} mismatches: {}", failures.len(), failures.join("\n"));
}

#[test]
fn collection_automation_dispatch_matches_excel() {
    let module = parse_module(include_str!("../../../tests/fixtures/vba/collection_dispatch.bas")).unwrap();
    assert_eq!(analyse(&module).metrics.unparsed, 0);
    let mut compared = 0;
    let mut failures = Vec::new();
    for line in include_str!("../../../tests/fixtures/vba/collection_dispatch_ja.tsv").lines() {
        let fields: Vec<_> = line.splitn(4, '\t').collect();
        let args = vec![Value::String(fields[0].to_string()), Value::Integer(fields[1].parse().unwrap()), Value::Integer(fields[2].parse().unwrap())];
        let expected = Value::String(fields[3].replace("{VBA_HELP}", HELP_FILE));
        let actual = Runtime::new(&module).with_error_environment(environment()).call("DispatchMatrix", args);
        if actual != Ok(expected.clone()) { failures.push(format!("{} {} {}: expected {expected:?}, got {actual:?}", fields[0], fields[1], fields[2])); }
        compared += 1;
    }
    assert_eq!(compared, 180);
    assert!(failures.is_empty(), "{} mismatches: {}", failures.len(), failures.join("\n"));
}
