// This Source Code Form is subject to the terms of the Mozilla Public
// License, v. 2.0. If a copy of the MPL was not distributed with this
// file, You can obtain one at https://mozilla.org/MPL/2.0/.

use oxivba_core::{parse_project, Runtime, Value};

const RECEIVER: &str = r#"Attribute VB_Name = "Factory"
Public Function Make() As Object
 Set Make = New Receiver
End Function
VERSION 1.0 CLASS
Attribute VB_Name = "Receiver"
Private count As Long
Public Property Get CounterValue() As Long
 CounterValue = count
End Property
Public Property Let CounterValue(ByVal value As Long)
 count = value
End Property
"#;

const CALLER: &str = r#"Attribute VB_Name = "Driver"
Public evaluations As Long
Public Function Probe(ByVal target As Object) As String
 Dim rhs As Object, number As Long
 Set rhs = New Rhs
 On Error Resume Next
 Err.Clear
 target.CounterValue = rhs
 number = Err.Number
 Probe = CStr(number) & "|" & CStr(evaluations) & "|" & CStr(target.CounterValue)
End Function
VERSION 1.0 CLASS
Attribute VB_Name = "Rhs"
Public Property Get Value() As Long
Attribute Value.VB_UserMemId = 0
 Driver.evaluations = Driver.evaluations + 1
 Value = 7
End Property
"#;

#[test]
fn foreign_property_value_reads_the_callers_default_member_once() {
    // Native Excel witness6414: a live receiver accepts the other workbook's
    // default scalar value, evaluates it once, and does not raise an error.
    let combined = format!("{RECEIVER}\n{CALLER}");
    let (module, classes) = parse_project(&combined).unwrap();
    let mut same = Runtime::new(&module).with_classes(&classes);
    let object = same.call("Factory.Make", vec![]).unwrap();
    assert_eq!(same.call("Driver.Probe", vec![object]).unwrap(), Value::String("0|1|7".into()), "same-project control");

    let (a, a_classes) = parse_project(RECEIVER).unwrap();
    let (b, b_classes) = parse_project(CALLER).unwrap();
    let mut root = Runtime::new(&module).with_classes(&classes);
    root.set_application_project_identity(3).unwrap();
    let mut owner = Runtime::new(&a).with_classes(&a_classes);
    owner.set_application_project_identity(1).unwrap();
    assert!(root.register_application_project(owner).is_ok());
    let mut caller = Runtime::new(&b).with_classes(&b_classes);
    caller.set_application_project_identity(2).unwrap();
    assert!(root.register_application_project(caller).is_ok());
    let object = root.in_application_project(1, |r| r.call("Factory.Make", vec![])).unwrap();
    let result = root.in_application_project(2, |r| r.call("Driver.Probe", vec![object])).unwrap();
    assert_eq!(result, Value::String("0|1|7".into()), "foreign receiver and caller-owned RHS");
}
