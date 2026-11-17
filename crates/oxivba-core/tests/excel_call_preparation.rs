// This Source Code Form is subject to the terms of the Mozilla Public
// License, v. 2.0. If a copy of the MPL was not distributed with this
// file, You can obtain one at https://mozilla.org/MPL/2.0/.

use oxivba_core::{
    parse_project, ErrorEnvironment, ErrorLanguage, Runtime, RuntimeErrorKind, Value,
};

#[test]
fn source_argument_preparation_matches_excel() {
    let fixtures =
        std::path::Path::new(env!("CARGO_MANIFEST_DIR")).join("../../tests/fixtures/vba");
    let rows: Vec<Vec<_>> = include_str!("../../../tests/fixtures/vba/call_preparation.tsv")
        .lines()
        .map(|line| line.split('\t').collect())
        .collect();
    let mut at = 0;
    while at < rows.len() {
        let name = rows[at][0];
        let source = std::fs::read_to_string(fixtures.join(rows[at][1])).unwrap();
        let (module, classes) = parse_project(&source).unwrap();
        let environment = || ErrorEnvironment {
            language: ErrorLanguage::Japanese,
            help_file: String::new(),
        };
        let mut continuous = Runtime::new(&module)
            .with_classes(&classes)
            .with_error_environment(environment());
        let mut saved = None;
        while at < rows.len() && rows[at][0] == name {
            let fields = &rows[at];
            let verify = |actual: Result<Value, oxivba_core::RuntimeError>| {
                if fields[3] == "compile" {
                    let error = actual.unwrap_err();
                    assert_eq!(error.kind, RuntimeErrorKind::Compile, "{name}");
                    assert_eq!(error.vba_number, None);
                    assert_eq!(
                        error.message,
                        format!("コンパイル エラー: {}", fields[4]),
                        "{name}"
                    );
                } else {
                    assert_eq!(
                        actual.unwrap_or_else(|error| panic!("{name}: {error:?}")),
                        Value::String(fields[4].into()),
                        "{name}"
                    );
                }
            };
            verify(continuous.call(fields[2], vec![]));
            let (module, classes) = parse_project(&source).unwrap();
            let mut restored = Runtime::new(&module)
                .with_classes(&classes)
                .with_error_environment(environment());
            if let Some(previous) = saved.take() {
                restored = restored.with_project_state(previous);
            }
            verify(restored.call(fields[2], vec![]));
            saved = Some(restored.into_project_state());
            at += 1;
        }
    }
}

fn argument_error(binding: &str, expression: &str) -> Value {
    let helpers = r#"
Private Function TakeValue(ByVal first As Variant, ByVal second As Variant) As Long
 TakeValue = 1
End Function
Private Function TakeNumber(ByVal first As Long, ByVal second As Variant) As Long
 TakeNumber = 1
End Function
Private Function TakeByteLong(ByVal first As Byte, ByVal second As Long) As Long
 TakeByteLong = 1
End Function
Private Function TakeLongByte(ByVal first As Long, ByVal second As Byte) As Long
 TakeLongByte = 1
End Function
"#;
    let (declaration, target) = match binding {
        "global" => (String::new(), expression.to_string()),
        "typed" => (
            " Dim worker As ErrorWorker\n Set worker = New ErrorWorker\n".into(),
            format!("worker.{expression}"),
        ),
        "callbyname" => {
            let (method, arguments) = expression.split_once('(').unwrap();
            (
                " Dim worker As Object\n Set worker = New ErrorWorker\n".into(),
                format!("CallByName(worker, \"{method}\", 1, {arguments}"),
            )
        }
        "dispatch" => (
            " Dim worker As Object\n Set worker = New ErrorWorker\n".into(),
            format!("worker.{expression}"),
        ),
        _ => panic!("unknown binding"),
    };
    let source = format!(
        "Attribute VB_Name = \"ArgumentErrors\"\n{helpers}\n\
         Public Function Probe() As Long\n\
          Dim large As Double, zero As Double, huge As Double\n\
          Dim byteOverflow As Variant, bad As Variant, sink As Variant\n\
         {declaration}\
          large = 1e308: zero = 0: huge = 1e20: byteOverflow = 300: bad = \"bad\"\n\
          On Error Resume Next\n\
          Err.Clear\n\
          sink = {target}\n\
          Probe = Err.Number\n\
         End Function\n\
         VERSION 1.0 CLASS\nAttribute VB_Name = \"ErrorWorker\"\n{}",
        helpers.replace("Private Function", "Public Function")
    );
    let (module, classes) = parse_project(&source).unwrap();
    Runtime::new(&module)
        .with_classes(&classes)
        .call("Probe", vec![])
        .unwrap()
}

#[test]
fn source_call_argument_operators_use_reverse_binding_order() {
    for (expression, expected) in [
        ("TakeValue(1/zero, large*large)", 6),
        ("TakeValue(large*large, 1/zero)", 11),
        ("TakeValue(CInt(huge), 1/zero)", 11),
        ("TakeValue(1/zero, CInt(huge))", 6),
        ("TakeNumber(bad, 1/zero)", 11),
        ("TakeNumber(1/zero, CInt(huge))", 6),
    ] {
        assert_eq!(argument_error("global", expression), Value::Integer(expected), "{expression}");
    }
}

const BINDING_ERRORS: &[(&str, i64, i64)] = &[
    ("TakeByteLong(1/zero, huge)", 6, 11),
    ("TakeLongByte(1/zero, byteOverflow)", 6, 11),
    ("TakeByteLong(large*large, bad)", 13, 6),
    ("TakeValue(large*large, bad)", 6, 6),
    ("TakeByteLong(byteOverflow, 1/zero)", 11, 11),
    ("TakeByteLong(CInt(huge), bad)", 13, 6),
    ("TakeByteLong(bad, CInt(huge))", 6, 6),
    ("TakeByteLong(byteOverflow, bad)", 13, 6),
    ("TakeLongByte(bad, byteOverflow)", 6, 13),
    ("TakeByteLong(bad, huge)", 6, 13),
    ("TakeLongByte(huge, bad)", 13, 6),
];

#[test]
fn source_call_argument_conversions_precede_earlier_operand_errors() {
    for &(expression, expected, _) in BINDING_ERRORS {
        assert_eq!(argument_error("global", expression), Value::Integer(expected), "{expression}");
    }
}

#[test]
fn source_call_argument_typed_methods_keep_conversion_order() {
    for &(expression, expected, _) in BINDING_ERRORS {
        assert_eq!(argument_error("typed", expression), Value::Integer(expected), "{expression}");
    }
}

#[test]
fn source_call_argument_dispatch_methods_evaluate_operands_before_binding() {
    for &(expression, _, expected) in BINDING_ERRORS {
        assert_eq!(argument_error("dispatch", expression), Value::Integer(expected), "{expression}");
    }
}

#[test]
fn source_call_argument_err_temporaries_follow_signature_and_receiver_timing() {
    for (receiver, mode, expected) in [
        ("ErrorSink", "ByVal", "5:7"),
        ("Object", "ByVal", "7:7"),
        ("Object", "ByRef", "7:7"),
        ("ErrorSink", "ByRef", "5:7"),
    ] {
        let argument = if mode == "ByRef" { "(held)" } else { "held" };
        let source = format!(
            "Attribute VB_Name = \"ErrorCaller\"\n\
             Public worker As ErrorSink\nPublic Log As String\n\
             Public Property Get Selected() As {receiver}\n\
              Err.Number = 7\n Set Selected = worker\nEnd Property\n\
             Public Function Probe() As String\n\
              Dim held As Variant\n Set worker = New ErrorSink\n Set held = Err\n\
              Err.Number = 5\n Selected.Change {argument}\n Probe = Log\nEnd Function\n\
             VERSION 1.0 CLASS\nAttribute VB_Name = \"ErrorSink\"\n\
             Public Sub Change({mode} value As Long)\n\
              ErrorCaller.Log = CStr(value) & \":\" & CStr(Err.Number)\nEnd Sub\n"
        );
        let (module, classes) = parse_project(&source).unwrap();
        let result = Runtime::new(&module)
            .with_classes(&classes)
            .call("Probe", vec![])
            .unwrap();
        assert_eq!(result, Value::String(expected.into()), "{receiver} {mode}");
    }
}


#[test]
fn source_call_argument_indexed_setters_keep_dispatch_conversion_order() {
    for (binding, expected) in [("global", 6), ("typed", 6), ("dispatch", 13)] {
        let (declaration, target) = match binding {
            "global" => ("", "Target"),
            "typed" => (" Dim worker As SetterWorker\n Set worker = New SetterWorker\n", "worker.Target"),
            "dispatch" => (" Dim worker As Object\n Set worker = New SetterWorker\n", "worker.Target"),
            _ => unreachable!(),
        };
        let setter = "Public Property Let Target(ByVal first As Long, ByVal second As Long, ByVal value As Byte)\nEnd Property\n";
        let source = format!(
            "Attribute VB_Name = \"SetterErrors\"\n{setter}\n\
             Public Function Probe() As Long\n\
              Dim bad As Variant, huge As Double, byteOverflow As Variant\n\
             {declaration}\
              bad = \"bad\": huge = 1e20: byteOverflow = 300\n\
              On Error Resume Next\n Err.Clear\n\
              {target}(bad, huge) = byteOverflow\n\
              Probe = Err.Number\nEnd Function\n\
             VERSION 1.0 CLASS\nAttribute VB_Name = \"SetterWorker\"\n{setter}"
        );
        let (module, classes) = parse_project(&source).unwrap();
        let result = Runtime::new(&module)
            .with_classes(&classes)
            .call("Probe", vec![])
            .unwrap();
        assert_eq!(result, Value::Integer(expected), "{binding}");
    }
}

#[test]
fn source_call_argument_callbyname_keeps_dispatch_conversion_order() {
    for &(expression, _, expected) in &BINDING_ERRORS[BINDING_ERRORS.len() - 4..] {
        assert_eq!(argument_error("callbyname", expression), Value::Integer(expected), "{expression}");
    }
}


#[test]
fn source_call_argument_indexed_setters_defer_dispatch_rhs_default() {
    for (receiver, expected) in [("DefaultSetter", "13|rhs;"), ("Object", "13|")] {
        let source = format!(
            "Attribute VB_Name = \"DefaultCaller\"\nPublic Log As String\n\
             Public Function Probe() As String\n\
              Dim worker As {receiver}, box As DefaultValue, rhs As Variant, bad As Variant\n\
              Set worker = New DefaultSetter\n Set box = New DefaultValue\n Set rhs = box\n\
              bad = \"bad\": Log = \"\"\n On Error Resume Next\n Err.Clear\n\
              worker.Target(bad, 0) = rhs\n\
              Probe = CStr(Err.Number) & \"|\" & Log\nEnd Function\n\
             VERSION 1.0 CLASS\nAttribute VB_Name = \"DefaultSetter\"\n\
             Public Property Let Target(ByVal first As Long, ByVal second As Long, ByVal value As Byte)\nEnd Property\n\
             VERSION 1.0 CLASS\nAttribute VB_Name = \"DefaultValue\"\n\
             Public Property Get Value() As Variant\nAttribute Value.VB_UserMemId = 0\n\
              DefaultCaller.Log = DefaultCaller.Log & \"rhs;\"\n Value = 1\nEnd Property\n"
        );
        let (module, classes) = parse_project(&source).unwrap();
        let result = Runtime::new(&module)
            .with_classes(&classes)
            .call("Probe", vec![])
            .unwrap();
        assert_eq!(result, Value::String(expected.into()), "{receiver}");
    }
}
