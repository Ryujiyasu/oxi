// This Source Code Form is subject to the terms of the Mozilla Public
// License, v. 2.0. If a copy of the MPL was not distributed with this
// file, You can obtain one at https://mozilla.org/MPL/2.0/.
use oxivba_core::{parser::parse_project, runtime::{Runtime, Value}};

#[test]
fn releasing_a_disconnected_owner_does_not_end_a_fresh_owner() {
    // Excel witness6034: after closing and reopening the origin, releasing
    // an old reference must not suppress the new instance's termination.
    let source = r#"Attribute VB_Name = "Driver"
Public trace As String
Public Function Make() As Object
 Set Make = New Widget
End Function
Public Function Probe(ByVal stale As Object) As String
 Dim fresh As Object
 trace = ""
 Set fresh = New Widget
 Set stale = Nothing
 Set fresh = Nothing
 Probe = trace
End Function
VERSION 1.0 CLASS
Attribute VB_Name = "Widget"
Private Sub Class_Initialize()
 Driver.trace = Driver.trace & "I"
End Sub
Private Sub Class_Terminate()
 Driver.trace = Driver.trace & "T"
End Sub
"#;
    let (module, classes) = parse_project(source).unwrap();
    let mut root = Runtime::new(&module).with_classes(&classes);
    root.set_application_project_identity(3).unwrap();
    for id in [1, 2] {
        let mut context = Runtime::new(&module).with_classes(&classes);
        context.set_application_project_identity(id).unwrap();
        assert!(root.register_application_project(context).is_ok());
    }
    let old = root.in_application_project(1, |r| r.call("Driver.Make", vec![])).unwrap();
    root.reset_application_project(1).unwrap();
    let result = root.in_application_project(2, |r| r.call("Driver.Probe", vec![old])).unwrap();
    assert_eq!(result, Value::String("IT".into()));
}
