// This Source Code Form is subject to the terms of the Mozilla Public
// License, v. 2.0. If a copy of the MPL was not distributed with this
// file, You can obtain one at https://mozilla.org/MPL/2.0/.

use oxivba_core::{parse_project, Host, ObjectRef, Runtime, Value};

#[derive(Default)]
struct ClosingHost(bool);

impl Host for ClosingHost {
    fn call(&mut self, receiver: Option<&ObjectRef>, name: &str, _: &[Value]) -> Result<Option<Value>, String> {
        if receiver.is_none() && name.eq_ignore_ascii_case("Application") {
            return Ok(Some(Value::Object(ObjectRef::new(42, "Application"))));
        }
        if receiver.is_some_and(|r| r.kind == "Application") && name.eq_ignore_ascii_case("CloseOrigin") {
            self.0 = true;
            return Ok(Some(Value::Empty));
        }
        Ok(None)
    }
    fn get(&mut self, _: &ObjectRef, _: &str) -> Result<Option<Value>, String> { Ok(None) }
    fn set(&mut self, _: &ObjectRef, _: &str, _: Value) -> Result<bool, String> { Ok(false) }
    fn take_closed_application_projects(&mut self) -> Vec<u64> {
        if std::mem::take(&mut self.0) { vec![1] } else { vec![] }
    }
}

#[test]
fn closing_a_property_receiver_during_value_conversion_keeps_the_caller_running() {
    // Native witness6414 selector2: the default getter closes the receiving
    // workbook successfully, then the assignment reports the disconnected HRESULT.
    let a = r#"Attribute VB_Name = "Factory"
Public Function Make() As Object
 Set Make = New Receiver
End Function
VERSION 1.0 CLASS
Attribute VB_Name = "Receiver"
Public Property Let CounterValue(ByVal value As Long)
End Property
"#;
    let b = r#"Attribute VB_Name = "Driver"
Public evaluations As Long
Public Function Probe(ByVal target As Object) As String
 Dim rhs As Object, number As Long
 Set rhs = New Rhs
 On Error Resume Next
 Err.Clear
 target.CounterValue = rhs
 number = Err.Number
 Probe = CStr(number) & "|" & CStr(evaluations)
End Function
VERSION 1.0 CLASS
Attribute VB_Name = "Rhs"
Public Property Get Value() As Long
Attribute Value.VB_UserMemId = 0
 Driver.evaluations = Driver.evaluations + 1
 Application.CloseOrigin
 Value = 7
End Property
"#;
    let (am, ac) = parse_project(a).unwrap();
    let (bm, bc) = parse_project(b).unwrap();
    let mut host = ClosingHost::default();
    let mut root = Runtime::new(&bm).with_classes(&bc);
    root.set_application_project_identity(3).unwrap();
    let mut owner = Runtime::new(&am).with_classes(&ac);
    owner.set_application_project_identity(1).unwrap();
    assert!(root.register_application_project(owner).is_ok());
    let mut caller = Runtime::new(&bm).with_classes(&bc).with_host(&mut host);
    caller.set_application_project_identity(2).unwrap();
    assert!(root.register_application_project(caller).is_ok());
    let target = root.in_application_project(1, |r| r.call("Factory.Make", vec![])).unwrap();
    let result = root.in_application_project(2, |r| r.call("Driver.Probe", vec![target])).unwrap();
    assert_eq!(result, Value::String("-2147418105|1".into()));
}
