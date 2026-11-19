// This Source Code Form is subject to the terms of the Mozilla Public
// License, v. 2.0. If a copy of the MPL was not distributed with this
// file, You can obtain one at https://mozilla.org/MPL/2.0/.

use oxivba_core::{parse_project, Runtime, Value};

#[test]
fn disconnected_property_assignment_does_not_end_the_calling_project() {
    // Excel witness6285: all six operations preserve fresh termination after
    // the stale receiver's project has closed, including direct Property Let.
    let source = r#"Attribute VB_Name = "Driver"
Public trace As String
Public Function Make() As Object
 Set Make = New Widget
End Function
Public Function Probe(ByVal stale As Object, ByVal mode As Long) As String
 Dim fresh As Object, result As Variant, obj As Object, number As Long
 On Error Resume Next
 trace = ""
 Err.Clear
 Select Case mode
 Case 0
  result = stale.CounterValue
 Case 1
  stale.CounterValue = 7
 Case 2
  Set obj = stale.ChildValue
 Case 3
  Set stale.ChildValue = Nothing
 Case 4
  result = CallByName(stale, "CounterValue", vbLet, 7)
 Case 5
  result = CallByName(stale, "ChildValue", vbSet, Nothing)
 End Select
 number = Err.Number
 Set fresh = New Widget
 Set fresh = Nothing
 Probe = CStr(number) & "|" & trace
End Function
VERSION 1.0 CLASS
Attribute VB_Name = "Widget"
Private count As Long
Private child As Object
Private Sub Class_Initialize()
 Driver.trace = Driver.trace & "I"
End Sub
Private Sub Class_Terminate()
 Driver.trace = Driver.trace & "T"
End Sub
Public Property Get CounterValue() As Long
 CounterValue = count
End Property
Public Property Let CounterValue(ByVal value As Long)
 count = value
End Property
Public Property Get ChildValue() As Object
 Set ChildValue = child
End Property
Public Property Set ChildValue(ByVal value As Object)
 Set child = value
End Property
"#;
    let (module, classes) = parse_project(source).unwrap();
    for mode in 0..6 {
        let mut root = Runtime::new(&module).with_classes(&classes);
        root.set_application_project_identity(3).unwrap();
        for id in [1, 2] {
            let mut context = Runtime::new(&module).with_classes(&classes);
            context.set_application_project_identity(id).unwrap();
            assert!(root.register_application_project(context).is_ok());
        }
        let old = root.in_application_project(1, |r| r.call("Driver.Make", vec![])).unwrap();
        root.reset_application_project(1).unwrap();
        let result = root.in_application_project(2, |r| r.call("Driver.Probe", vec![old, Value::Integer(mode)])).unwrap();
        let number = if mode >= 4 { 440 } else { -2147418105_i64 };
        assert_eq!(result, Value::String(format!("{number}|IT")), "mode {mode}");
    }
}
