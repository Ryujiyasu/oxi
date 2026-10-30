// This Source Code Form is subject to the terms of the Mozilla Public
// License, v. 2.0. If a copy of the MPL was not distributed with this
// file, You can obtain one at https://mozilla.org/MPL/2.0/.

use std::path::Path;

use oxivba_core::{
    analyse, parse_project, ErrorEnvironment, ErrorLanguage, Runtime, RuntimeErrorKind, Value,
};

#[test]
fn bare_print_output_fields_are_not_parenthesized_argument_errors() {
    let module = oxivba_core::parse_module("Sub Probe()\nDebug.Print 1,,\nEnd Sub\n").unwrap();
    assert_eq!(analyse(&module).metrics.unparsed, 0);
    let procedure = module
        .items
        .iter()
        .find_map(|item| match item {
            oxivba_core::ModuleItem::Procedure(procedure) => Some(procedure),
            _ => None,
        })
        .unwrap();
    assert!(oxivba_core::check_procedure(&module, procedure).is_ok());
}

#[test]
fn procedure_compilation_and_execution_order_match_excel() {
    let fixtures = Path::new(env!("CARGO_MANIFEST_DIR")).join("../../tests/fixtures/vba/compile");
    let cases = std::fs::read_to_string(fixtures.join("cases.tsv")).unwrap();
    let mut failures = Vec::new();
    let mut compared = 0;
    for line in cases.lines() {
        let fields: Vec<_> = line.splitn(6, '\t').collect();
        assert_eq!(fields.len(), 6);
        let [name, main, other, outcome, expected, trace] = fields.as_slice() else {
            unreachable!()
        };
        let mut source = format!(
            "Attribute VB_Name = \"EntryProbe\"\n{}",
            std::fs::read_to_string(fixtures.join(main)).unwrap()
        );
        if !other.is_empty() {
            source += &format!(
                "\nAttribute VB_Name = \"OtherProbe\"\n{}",
                std::fs::read_to_string(fixtures.join(other)).unwrap()
            );
        }
        let (module, classes) = parse_project(&source).unwrap();
        assert_eq!(analyse(&module).metrics.unparsed, 0, "{name}");
        let mut runtime = Runtime::new(&module)
            .with_classes(&classes)
            .with_error_environment(ErrorEnvironment {
                language: ErrorLanguage::Japanese,
                help_file: String::new(),
            });
        let actual = runtime.call("Probe", vec![]);
        let matching = if *outcome == "compile" {
            match &actual {
                Err(failure) => {
                    failure.kind == RuntimeErrorKind::Compile
                        && failure.vba_number.is_none()
                        && failure.message == format!("コンパイル エラー: {expected}")
                }
                _ => false,
            }
        } else {
            actual == Ok(Value::String(expected.to_string()))
        };
        if !matching {
            failures.push(format!(
                "{name}: expected {outcome} {expected:?}, got {actual:?}"
            ));
        }
        let actual_trace = runtime.call("ReadTrace", vec![]);
        if actual_trace != Ok(Value::String(trace.to_string())) {
            failures.push(format!(
                "{name}: expected trace {trace:?}, got {actual_trace:?}"
            ));
        }
        compared += 1;
    }
    assert_eq!(compared, 78);
    assert!(
        failures.is_empty(),
        "{} mismatches:\n{}",
        failures.len(),
        failures.join("\n")
    );
}
