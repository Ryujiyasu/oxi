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


fn pointer_compile_source(declaration: &str, expression: &str, helper: &str) -> String {
    format!(r#"
Attribute VB_Name = "PointerCompile"
Option Explicit
Public trace As Long
{helper}
Public Function Probe() As Boolean
 {declaration}
 trace=1
 Probe=({expression})
End Function
Public Function ReadTrace() As Long
 ReadTrace=trace
End Function
"#)
}

#[test]
fn varptr_whole_array_arguments_fail_during_compilation() {
    for (declaration, expression) in [
        ("Dim values() As Object", "VarPtr(values)<>0"),
        ("Dim values() As Long", "VarPtr(values)<>0"),
        ("Dim values() As Variant", "VarPtr(values)<>0"),
        ("Dim values() As Object", "VarPtr((values))<>0"),
        ("Dim values() As Object", "VarPtr(values())<>0"),
        ("Dim values() As Object", "VBA.VarPtr(values)<>0"),
    ] {
        let source = pointer_compile_source(declaration,expression,"");
        let failure = oxivba_core::compile_project(&source).expect_err(expression);
        assert_eq!(failure.error.kind, oxivba_core::CompileErrorKind::TypeMismatch, "{expression}");
        let (module, classes) = parse_project(&source).unwrap();
        let mut runtime = Runtime::new(&module).with_classes(&classes);
        let failure = runtime.call("Probe",vec![]).expect_err(expression);
        assert_eq!(failure.kind, RuntimeErrorKind::Compile, "{expression}");
        assert_eq!(failure.vba_number, None, "{expression}");
        assert_eq!(runtime.call("ReadTrace",vec![]).unwrap(), Value::Integer(0), "{expression}");
    }
}

#[test]
fn varptr_variant_slots_elements_and_source_shadows_remain_valid() {
    for (declaration, expression, helper) in [
        ("Dim values As Variant:values=Array(1,2)", "VarPtr(values)<>0", ""),
        ("Dim values(0 To 0) As Object", "VarPtr(values(0))<>0", ""),
        ("Dim values() As Object", "VarPtr(values)=77", "Private Function VarPtr(ByRef value() As Object) As Long\n VarPtr=77\nEnd Function\n"),
    ] {
        let source = pointer_compile_source(declaration,expression,helper);
        oxivba_core::compile_project(&source).unwrap_or_else(|failure|panic!("{expression}: {failure:?}"));
        let (module, classes) = parse_project(&source).unwrap();
        let mut runtime = Runtime::new(&module).with_classes(&classes);
        assert_eq!(runtime.call("Probe",vec![]).unwrap(), Value::Boolean(true), "{expression}");
        assert_eq!(runtime.call("ReadTrace",vec![]).unwrap(), Value::Integer(1), "{expression}");
    }
}
