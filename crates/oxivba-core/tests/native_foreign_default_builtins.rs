// This Source Code Form is subject to the terms of the Mozilla Public
// License, v. 2.0. If a copy of the MPL was not distributed with this
// file, You can obtain one at https://mozilla.org/MPL/2.0/.

use oxivba_core::{parse_project, Runtime, Value};

#[test]
fn foreign_default_getter_metadata_preserves_predicate_errors_and_nested_len() {
    // Excel witness6435: predicates retain a failing default getter's Err;
    // nested Len follows the second default and returns the measured Date tag.
    let owner = r#"Attribute VB_Name = "Factory"
Public selector As Long, evaluations As Long
Public Function Make(ByVal mode As Long) As Object
 selector = mode
 evaluations = 0
 Set Make = New Rhs
End Function
VERSION 1.0 CLASS
Attribute VB_Name = "Rhs"
Public Property Get Value() As Variant
Attribute Value.VB_UserMemId = 0
 Factory.evaluations = Factory.evaluations + 1
 If Factory.selector = 1 Then Err.Raise 5, "ForeignDefault6431", "foreign default getter"
 If Factory.selector = 2 Then
  Set Value = New Nested
 Else
  Value = 7
 End If
End Property
Public Function ReadCount() As Long
 ReadCount = Factory.evaluations
End Function
VERSION 1.0 CLASS
Attribute VB_Name = "Nested"
Public Property Get Value() As Long
Attribute Value.VB_UserMemId = 0
 Factory.evaluations = Factory.evaluations + 1
 Value = 7
End Property
"#;
    let caller = r#"Attribute VB_Name = "Driver"
Public Function Probe(ByVal rhs As Object, ByVal mode As Long) As String
 Dim result As Variant, number As Long, source As String
 On Error Resume Next
 Err.Clear
 If mode = 2 Then
  result = Len(rhs)
  result = VarType(result)
 Else
  result = VarType(rhs)
 End If
 number = Err.Number
 source = Err.Source
 Probe = CStr(result) & "|" & CStr(number) & "|" & CStr(rhs.ReadCount()) & "|" & source
End Function
"#;
    let (a, ac) = parse_project(owner).unwrap();
    let (b, bc) = parse_project(caller).unwrap();
    for (mode, expected) in [(0, "2|0|1|"), (1, "9|5|1|ForeignDefault6431"), (2, "7|0|2|")] {
        let mut root = Runtime::new(&b).with_classes(&bc);
        root.set_application_project_identity(3).unwrap();
        let mut owner = Runtime::new(&a).with_classes(&ac);
        owner.set_application_project_identity(1).unwrap();
        assert!(root.register_application_project(owner).is_ok());
        let mut caller = Runtime::new(&b).with_classes(&bc);
        caller.set_application_project_identity(2).unwrap();
        assert!(root.register_application_project(caller).is_ok());
        let rhs = root.in_application_project(1, |r| r.call("Factory.Make", vec![Value::Integer(mode)])).unwrap();
        let result = root.in_application_project(2, |r| r.call("Driver.Probe", vec![rhs, Value::Integer(mode)])).unwrap();
        assert_eq!(result, Value::String(expected.into()), "selector {mode}");
    }
}
