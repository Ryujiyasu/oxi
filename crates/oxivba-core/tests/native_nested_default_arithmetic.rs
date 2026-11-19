// This Source Code Form is subject to the terms of the Mozilla Public
// License, v. 2.0. If a copy of the MPL was not distributed with this
// file, You can obtain one at https://mozilla.org/MPL/2.0/.

use oxivba_core::{parse_project, Host, ObjectRef, Runtime, Value};
use std::{cell::RefCell, rc::Rc};

struct TraceHost(Rc<RefCell<String>>);

impl Host for TraceHost {
    fn call(&mut self, receiver: Option<&ObjectRef>, name: &str, args: &[Value]) -> Result<Option<Value>, String> {
        if receiver.is_none() && name.eq_ignore_ascii_case("Application") {
            return Ok(Some(Value::Object(ObjectRef::new(42, "Application"))));
        }
        if receiver.is_some_and(|r| r.kind == "Application") && name.eq_ignore_ascii_case("Note") {
            let Some(Value::String(text)) = args.first() else { return Err("Note requires text".into()); };
            self.0.borrow_mut().push_str(text);
            return Ok(Some(Value::Empty));
        }
        Ok(None)
    }
    fn get(&mut self, _: &ObjectRef, _: &str) -> Result<Option<Value>, String> { Ok(None) }
    fn set(&mut self, _: &ObjectRef, _: &str, _: Value) -> Result<bool, String> { Ok(false) }
}

const OWNER: &str = r#"Attribute VB_Name = "Factory"
Public mode As Long
Public Function Make(ByVal kind As Long) As Object
 mode = kind
 Set Make = New Rhs
End Function
VERSION 1.0 CLASS
Attribute VB_Name = "Rhs"
Public Property Get Value() As Variant
Attribute Value.VB_UserMemId = 0
 Application.Note "O"
 If Factory.mode = 2 Then
  Value = 7
 Else
  Set Value = New Nested
 End If
End Property
VERSION 1.0 CLASS
Attribute VB_Name = "Nested"
Public Property Get Value() As Long
Attribute Value.VB_UserMemId = 0
 Application.Note "N"
 If Factory.mode = 0 Then
  Value = 3
 Else
  Value = 7
 End If
End Property
"#;

const CALLER: &str = r#"Attribute VB_Name = "Driver"
Public Function Operand() As Long
 Application.Note "R"
 Operand = 1
End Function
Public Function Literal(ByVal target As Object) As Variant
 Dim result As Variant
 result = target + 1
 Literal = result
End Function
Public Function WithOperand(ByVal target As Object) As Variant
 Dim result As Variant
 result = target + Operand()
 WithOperand = result
End Function
Public Function Boundaries(ByVal target As Object, ByVal selector As Long) As Variant
 Dim result As Variant
 Dim data(0 To 1) As Long
 Dim nilValue As Object
 On Error GoTo Failed
 Select Case selector
 Case 0: result = (target) + Operand()
 Case 1: result = target + Null
 Case 2: result = target / 0
 Case 3: result = target + (1 / 0)
 Case 4: result = target + nilValue
 Case 5: result = target + "bad"
 Case 6: result = target + UBound(data)
 Case 7: result = target + IIf(True, 1, 2)
 Case 8: result = target + Val("1")
 End Select
 Boundaries = result
 Exit Function
Failed:
 Boundaries = Err.Number
End Function

Public Function NullOperand() As Variant
 Application.Note "R"
 NullOperand = Null
End Function
Public Function NullOperators(ByVal target As Object, ByVal selector As Long) As Variant
 Dim result As Variant, nullable As Variant
 nullable = Null
 Select Case selector
 Case 0
  result = target & Null
 Case 1
  result = target And Null
 Case 2
  result = target Or Null
 Case 3
  result = target Xor Null
 Case 4
  result = target Eqv Null
 Case 5
  result = target Imp Null
 Case 6
  result = target Like Null
 Case 7
  result = target = Null
 Case 8
  result = target <> Null
 Case 9
  result = target < Null
 Case 10
  result = target > Null
 Case 11
  result = target <= Null
 Case 12
  result = target >= Null
 Case 13
  result = target ^ Null
 Case 14
  result = target \ Null
 Case 15
  result = target Mod Null
 Case 16
  result = Null & target
 Case 17
  result = Null And target
 Case 18
  result = Null Or target
 Case 19
  result = Null Xor target
 End Select
 NullOperators = result
End Function
Public Function GroupedNullAddition(ByVal target As Object) As Variant
 GroupedNullAddition = Null + (target)
End Function
Public Function NullPositions(ByVal target As Object, ByVal selector As Long) As Variant
 Dim result As Variant, nullable As Variant
 nullable = Null
 Select Case selector
 Case 0
  result = target - Null
 Case 1
  result = target * Null
 Case 2
  result = target / Null
 Case 3
  result = Null + target
 Case 4
  result = Null - target
 Case 5
  result = Null * target
 Case 6
  result = Null / target
 Case 7
  result = Null ^ target
 Case 8
  result = Null \ target
 Case 9
  result = Null Mod target
 Case 10
  result = Null Eqv target
 Case 11
  result = Null Imp target
 Case 12
  result = Null Like target
 Case 13
  result = target + NullOperand()
 Case 14
  result = target & NullOperand()
 Case 15
  result = target And NullOperand()
 Case 16
  result = target = NullOperand()
 Case 17
  result = target + nullable
 Case 18
  result = target & nullable
 Case 19
  result = target = nullable
 End Select
 NullPositions = result
End Function
"#;

fn probe(foreign: bool, mode: i64, procedure: &str) -> (Value, String) {
    probe_args(foreign, mode, procedure, vec![])
}

fn probe_args(foreign: bool, mode: i64, procedure: &str, extra: Vec<Value>) -> (Value, String) {
    let observed = Rc::new(RefCell::new(String::new()));
    let mut a_host = TraceHost(observed.clone());
    let mut b_host = TraceHost(observed.clone());
    let (a, ac) = parse_project(OWNER).unwrap();
    let (b, bc) = parse_project(CALLER).unwrap();
    let combined = format!("{OWNER}\n{CALLER}");
    let (module, classes) = parse_project(&combined).unwrap();
    let result = if foreign {
        let mut root = Runtime::new(&b).with_classes(&bc);
        root.set_application_project_identity(3).unwrap();
        let mut owner = Runtime::new(&a).with_classes(&ac).with_host(&mut a_host);
        owner.set_application_project_identity(1).unwrap();
        assert!(root.register_application_project(owner).is_ok());
        let mut caller = Runtime::new(&b).with_classes(&bc).with_host(&mut b_host);
        caller.set_application_project_identity(2).unwrap();
        assert!(root.register_application_project(caller).is_ok());
        let object = root.in_application_project(1, |r| r.call("Factory.Make", vec![Value::Integer(mode)])).unwrap();
        let mut args = vec![object];
        args.extend(extra);
        root.in_application_project(2, |r| r.call(procedure, args)).unwrap()
    } else {
        let mut runtime = Runtime::new(&module).with_classes(&classes).with_host(&mut a_host);
        runtime.set_application_project_identity(1).unwrap();
        let object = runtime.call("Factory.Make", vec![Value::Integer(mode)]).unwrap();
        let mut args = vec![object];
        args.extend(extra);
        runtime.call(procedure, args).unwrap()
    };
    let trace = observed.borrow().clone();
    (result, trace)
}

#[test]
fn nested_literal_addition_preserves_the_native_variant_result() {
    // Excel witness6523: nested Long3/Long7 produce Long0/Date0 here;
    // the scalar Variant7 control still produces Integer8.
    for foreign in [false, true] {
        for (mode, expected) in [(0, Value::Integer(0)), (1, Value::Date(0.0)), (2, Value::Int16(8))] {
            let (result, trace) = probe(foreign, mode, "Driver.Literal");
            assert_eq!(result, expected, "mode {mode}, foreign {foreign}");
            assert_eq!(trace, if mode == 2 { "O" } else { "ON" });
        }
    }
}

#[test]
fn nested_function_operand_runs_between_the_two_default_getters() {
    // Excel witness6536: O (outer default), R (right operand), N (nested
    // default); the scalar control has no N. Function arithmetic is normal.
    for foreign in [false, true] {
        for mode in 0..3 {
            let (result, trace) = probe(foreign, mode, "Driver.WithOperand");
            assert_eq!(result, Value::Integer(if mode == 0 { 4 } else { 8 }));
            assert_eq!(trace, if mode == 2 { "OR" } else { "ORN" }, "mode {mode}, foreign {foreign}");
        }
    }
}

#[test]
fn nested_default_error_precedence_and_operand_storage_match_excel() {
    // Native witness6556: errors and Null can stop before N; parentheses
    // must not read the outer default twice. Inline bounds differ from
    // runtime-call temporaries such as IIf and Val.
    for foreign in [false, true] {
        for mode in 0..2 {
            let normal = if mode == 0 { 4 } else { 8 };
            let cases = [
                (Value::Integer(normal), "ORN"),
                (Value::Null, "O"),
                (Value::Integer(11), "ON"),
                (Value::Integer(11), "O"),
                (Value::Integer(91), "O"),
                (Value::Integer(13), "ON"),
                (if mode == 0 { Value::Integer(0) } else { Value::Date(0.0) }, "ON"),
                (Value::Integer(normal), "ON"),
                (Value::Double(normal as f64), "ON"),
            ];
            for (selector, (expected, trace)) in cases.into_iter().enumerate() {
                let actual = probe_args(foreign, mode, "Driver.Boundaries", vec![Value::Integer(selector as i64)]);
                assert_eq!(actual, (expected, trace.to_owned()), "selector {selector}, mode {mode}, foreign {foreign}");
            }
        }
    }
}

#[test]
fn nested_null_operators_keep_native_values_and_getter_order() {
    // Native witnesses 6760: two stable captures, 240 call points.
    for foreign in [false, true] {
        for mode in 0..3 {
            let cases = match (mode, foreign) {
                (0, false) => [
                    (Value::Integer(0), "ON"),
                    (Value::Integer(0), "ON"),
                    (Value::Integer(0), "ON"),
                    (Value::Integer(0), "ON"),
                    (Value::Integer(0), "ON"),
                    (Value::Integer(0), "ON"),
                    (Value::Integer(0), "ON"),
                    (Value::Null, "ON"),
                    (Value::Null, "ON"),
                    (Value::Null, "ON"),
                    (Value::Null, "ON"),
                    (Value::Null, "ON"),
                    (Value::Null, "ON"),
                    (Value::Integer(0), "ON"),
                    (Value::Integer(0), "ON"),
                    (Value::Integer(0), "ON"),
                    (Value::String("3".into()), "ON"),
                    (Value::Null, "ON"),
                    (Value::Integer(3), "ON"),
                    (Value::Null, "ON"),
                ],
                (0, true) => [
                    (Value::Integer(0), "ON"),
                    (Value::Integer(0), "ON"),
                    (Value::Integer(0), "ON"),
                    (Value::Integer(0), "ON"),
                    (Value::Integer(0), "ON"),
                    (Value::Integer(0), "ON"),
                    (Value::Integer(0), "ON"),
                    (Value::Null, "ON"),
                    (Value::Null, "ON"),
                    (Value::Null, "ON"),
                    (Value::Null, "ON"),
                    (Value::Null, "ON"),
                    (Value::Null, "ON"),
                    (Value::Integer(0), "ON"),
                    (Value::Integer(0), "ON"),
                    (Value::Integer(0), "ON"),
                    (Value::String("3".into()), "ON"),
                    (Value::Null, "ON"),
                    (Value::Integer(3), "ON"),
                    (Value::Null, "ON"),
                ],
                (1, false) => [
                    (Value::Date(0.0), "ON"),
                    (Value::Date(0.0), "ON"),
                    (Value::Date(0.0), "ON"),
                    (Value::Date(0.0), "ON"),
                    (Value::Date(0.0), "ON"),
                    (Value::Date(0.0), "ON"),
                    (Value::Date(0.0), "ON"),
                    (Value::Null, "ON"),
                    (Value::Null, "ON"),
                    (Value::Null, "ON"),
                    (Value::Null, "ON"),
                    (Value::Null, "ON"),
                    (Value::Null, "ON"),
                    (Value::Date(0.0), "ON"),
                    (Value::Date(0.0), "ON"),
                    (Value::Date(0.0), "ON"),
                    (Value::String("7".into()), "ON"),
                    (Value::Null, "ON"),
                    (Value::Integer(7), "ON"),
                    (Value::Null, "ON"),
                ],
                (1, true) => [
                    (Value::Date(0.0), "ON"),
                    (Value::Date(0.0), "ON"),
                    (Value::Date(0.0), "ON"),
                    (Value::Date(0.0), "ON"),
                    (Value::Date(0.0), "ON"),
                    (Value::Date(0.0), "ON"),
                    (Value::Date(0.0), "ON"),
                    (Value::Null, "ON"),
                    (Value::Null, "ON"),
                    (Value::Null, "ON"),
                    (Value::Null, "ON"),
                    (Value::Null, "ON"),
                    (Value::Null, "ON"),
                    (Value::Date(0.0), "ON"),
                    (Value::Date(0.0), "ON"),
                    (Value::Date(0.0), "ON"),
                    (Value::String("7".into()), "ON"),
                    (Value::Null, "ON"),
                    (Value::Integer(7), "ON"),
                    (Value::Null, "ON"),
                ],
                (2, false) => [
                    (Value::String("7".into()), "O"),
                    (Value::Null, "O"),
                    (Value::Int16(7), "O"),
                    (Value::Null, "O"),
                    (Value::Null, "O"),
                    (Value::Int16(-8), "O"),
                    (Value::Null, "O"),
                    (Value::Null, "O"),
                    (Value::Null, "O"),
                    (Value::Null, "O"),
                    (Value::Null, "O"),
                    (Value::Null, "O"),
                    (Value::Null, "O"),
                    (Value::Null, "O"),
                    (Value::Null, "O"),
                    (Value::Null, "O"),
                    (Value::String("7".into()), "O"),
                    (Value::Null, "O"),
                    (Value::Int16(7), "O"),
                    (Value::Null, "O"),
                ],
                (2, true) => [
                    (Value::String("7".into()), "O"),
                    (Value::Null, "O"),
                    (Value::Int16(7), "O"),
                    (Value::Null, "O"),
                    (Value::Null, "O"),
                    (Value::Int16(-8), "O"),
                    (Value::Null, "O"),
                    (Value::Null, "O"),
                    (Value::Null, "O"),
                    (Value::Null, "O"),
                    (Value::Null, "O"),
                    (Value::Null, "O"),
                    (Value::Null, "O"),
                    (Value::Null, "O"),
                    (Value::Null, "O"),
                    (Value::Null, "O"),
                    (Value::String("7".into()), "O"),
                    (Value::Null, "O"),
                    (Value::Int16(7), "O"),
                    (Value::Null, "O"),
                ],
                _ => unreachable!(),
            };
            for (selector, (expected, trace)) in cases.into_iter().enumerate() {
                let actual = probe_args(foreign, mode, "Driver.NullOperators", vec![Value::Integer(selector as i64)]);
                assert_eq!(actual, (expected, trace.to_owned()), "selector {selector}, mode {mode}, foreign {foreign}");
            }
        }
    }
}

#[test]
fn null_positions_and_temporaries_keep_native_evaluation() {
    // Native witnesses 6764: two stable captures, 240 call points.
    for foreign in [false, true] {
        for mode in 0..3 {
            let cases = match (mode, foreign) {
                (0, false) => [
                    (Value::Null, "O"),
                    (Value::Null, "O"),
                    (Value::Null, "O"),
                    (Value::Null, "O"),
                    (Value::Null, "O"),
                    (Value::Null, "O"),
                    (Value::Null, "O"),
                    (Value::Null, "ON"),
                    (Value::Null, "ON"),
                    (Value::Null, "ON"),
                    (Value::Null, "ON"),
                    (Value::Integer(3), "ON"),
                    (Value::Null, "ON"),
                    (Value::Null, "OR"),
                    (Value::String("3".into()), "ORN"),
                    (Value::Null, "ORN"),
                    (Value::Null, "ORN"),
                    (Value::Null, "O"),
                    (Value::Integer(0), "ON"),
                    (Value::Null, "ON"),
                ],
                (0, true) => [
                    (Value::Null, "O"),
                    (Value::Null, "O"),
                    (Value::Null, "O"),
                    (Value::Null, "O"),
                    (Value::Null, "O"),
                    (Value::Null, "O"),
                    (Value::Null, "O"),
                    (Value::Null, "ON"),
                    (Value::Null, "ON"),
                    (Value::Null, "ON"),
                    (Value::Null, "ON"),
                    (Value::Integer(3), "ON"),
                    (Value::Null, "ON"),
                    (Value::Null, "OR"),
                    (Value::String("3".into()), "ORN"),
                    (Value::Null, "ORN"),
                    (Value::Null, "ORN"),
                    (Value::Null, "O"),
                    (Value::Integer(0), "ON"),
                    (Value::Null, "ON"),
                ],
                (1, false) => [
                    (Value::Null, "O"),
                    (Value::Null, "O"),
                    (Value::Null, "O"),
                    (Value::Null, "O"),
                    (Value::Null, "O"),
                    (Value::Null, "O"),
                    (Value::Null, "O"),
                    (Value::Null, "ON"),
                    (Value::Null, "ON"),
                    (Value::Null, "ON"),
                    (Value::Null, "ON"),
                    (Value::Integer(7), "ON"),
                    (Value::Null, "ON"),
                    (Value::Null, "OR"),
                    (Value::String("7".into()), "ORN"),
                    (Value::Null, "ORN"),
                    (Value::Null, "ORN"),
                    (Value::Null, "O"),
                    (Value::Date(0.0), "ON"),
                    (Value::Null, "ON"),
                ],
                (1, true) => [
                    (Value::Null, "O"),
                    (Value::Null, "O"),
                    (Value::Null, "O"),
                    (Value::Null, "O"),
                    (Value::Null, "O"),
                    (Value::Null, "O"),
                    (Value::Null, "O"),
                    (Value::Null, "ON"),
                    (Value::Null, "ON"),
                    (Value::Null, "ON"),
                    (Value::Null, "ON"),
                    (Value::Integer(7), "ON"),
                    (Value::Null, "ON"),
                    (Value::Null, "OR"),
                    (Value::String("7".into()), "ORN"),
                    (Value::Null, "ORN"),
                    (Value::Null, "ORN"),
                    (Value::Null, "O"),
                    (Value::Date(0.0), "ON"),
                    (Value::Null, "ON"),
                ],
                (2, false) => [
                    (Value::Null, "O"),
                    (Value::Null, "O"),
                    (Value::Null, "O"),
                    (Value::Null, "O"),
                    (Value::Null, "O"),
                    (Value::Null, "O"),
                    (Value::Null, "O"),
                    (Value::Null, "O"),
                    (Value::Null, "O"),
                    (Value::Null, "O"),
                    (Value::Null, "O"),
                    (Value::Int16(7), "O"),
                    (Value::Null, "O"),
                    (Value::Null, "OR"),
                    (Value::String("7".into()), "OR"),
                    (Value::Null, "OR"),
                    (Value::Null, "OR"),
                    (Value::Null, "O"),
                    (Value::String("7".into()), "O"),
                    (Value::Null, "O"),
                ],
                (2, true) => [
                    (Value::Null, "O"),
                    (Value::Null, "O"),
                    (Value::Null, "O"),
                    (Value::Null, "O"),
                    (Value::Null, "O"),
                    (Value::Null, "O"),
                    (Value::Null, "O"),
                    (Value::Null, "O"),
                    (Value::Null, "O"),
                    (Value::Null, "O"),
                    (Value::Null, "O"),
                    (Value::Int16(7), "O"),
                    (Value::Null, "O"),
                    (Value::Null, "OR"),
                    (Value::String("7".into()), "OR"),
                    (Value::Null, "OR"),
                    (Value::Null, "OR"),
                    (Value::Null, "O"),
                    (Value::String("7".into()), "O"),
                    (Value::Null, "O"),
                ],
                _ => unreachable!(),
            };
            for (selector, (expected, trace)) in cases.into_iter().enumerate() {
                let actual = probe_args(foreign, mode, "Driver.NullPositions", vec![Value::Integer(selector as i64)]);
                assert_eq!(actual, (expected, trace.to_owned()), "selector {selector}, mode {mode}, foreign {foreign}");
            }
        }
    }
}

#[test]
fn grouped_right_default_stops_before_nested_getter_with_null_addition() {
    // Native witness6853: two stable captures; selectors16/36/56/76.
    for foreign in [false, true] {
        for mode in 0..3 {
            assert_eq!(probe(foreign, mode, "Driver.GroupedNullAddition"),
                (Value::Null, "O".to_owned()), "mode {mode}, foreign {foreign}");
        }
    }
}
