// This Source Code Form is subject to the terms of the Mozilla Public
// License, v. 2.0. If a copy of the MPL was not distributed with this
// file, You can obtain one at https://mozilla.org/MPL/2.0/.
use oxivba_core::{parser::parse_project, runtime::{Runtime, Value}};
fn observed(value: &Value) -> String {
    match value {
        Value::String(v) => format!("{v:?}"),
        Value::NullString => "\"\"".into(),
        Value::Integer(v) => v.to_string(),
        Value::Boolean(v) => v.to_string(),
        Value::Empty | Value::Null => "null".into(),
        Value::Array(a) => format!("[{}]", a.values.iter().map(observed).collect::<Vec<_>>().join(",")),
        other => panic!("Unmeasured result: {other:?}"),
    }
}

fn compare_japanese(source: &str, expected: [&str; 2]) {
    let (module, classes) = parse_project(source).unwrap();
    let mut runtime = Runtime::new(&module).with_classes(&classes).with_error_environment(oxivba_core::ErrorEnvironment { language: oxivba_core::ErrorLanguage::Japanese, help_file: r"C:\Program Files\Common Files\Microsoft Shared\VBA\VBA7.1\1041\VbLR6.chm".into() });
    for row in expected {
        let value = runtime.call("Predeclared5288.Probe", vec![]).unwrap();
        assert_eq!(observed(&value), row, "continuous runtime");
    }
    let mut state = None;
    for row in expected {
        let mut restored = Runtime::new(&module).with_classes(&classes).with_error_environment(oxivba_core::ErrorEnvironment { language: oxivba_core::ErrorLanguage::Japanese, help_file: r"C:\Program Files\Common Files\Microsoft Shared\VBA\VBA7.1\1041\VbLR6.chm".into() });
        if let Some(saved) = state.take() { restored = restored.with_project_state(saved); }
        let value = restored.call("Predeclared5288.Probe", vec![]).unwrap();
        assert_eq!(observed(&value), row, "restored project storage");
        state = Some(restored.into_project_state());
    }
}

#[test]
fn native_predeclared_rebound_held_error() {
    compare_japanese(r########"Attribute VB_Name = "Predeclared5288"
Option Explicit
Public trace As String,initCount As Long,termCount As Long
Public held As Object
Public Sub MarkInit()
 trace=trace & "I":initCount=initCount+1
End Sub
Public Sub NoteSelf(ByVal same As Boolean)
 If same Then trace=trace & "S" Else trace=trace & "D"
End Sub
Public Sub MarkTerm()
 trace=trace & "T":termCount=termCount+1
End Sub
Public Function Probe() As Variant
 Dim first As Long,second As Long,before As Variant
 
 On Error Resume Next
 Err.Clear
 first=Automatic5288.Bump():second=held.Bump()
 before=Array(trace,initCount,termCount)
 Probe=Array(before,first,second,Err.Number,Err.Description,Err.Source,Err.HelpFile,Err.HelpContext)
End Function
Attribute VB_Name = "Other5299"
Option Explicit
Public Function Touch() As Long
 Touch=Automatic5288.Bump()
End Function
VERSION 1.0 CLASS
Attribute VB_Name = "Automatic5288"
Attribute VB_PredeclaredId = True
Option Explicit
Private counter As Long,tag As Long
Private Sub Class_Initialize()
 Predeclared5288.MarkInit
 tag=Predeclared5288.initCount
 If Predeclared5288.initCount=1 Then Set Automatic5288=New Automatic5288: Set Predeclared5288.held=Automatic5288: Err.Raise 513, "Ctor", "failed with held replacement"
End Sub
Private Sub Class_Terminate()
 Predeclared5288.MarkTerm
End Sub
Public Function Bump() As Long
 counter=counter+1:Bump=tag*100+counter
End Function
Public Function SameInstance() As Long
 If Me Is Automatic5288 Then SameInstance=1 Else SameInstance=0
End Function
Public Function Outer() As Long
 Outer=Automatic5288.Bump()
End Function
"########, [
        r########"[["II",2,0],0,201,513,"failed with held replacement","Ctor","C:\\Program Files\\Common Files\\Microsoft Shared\\VBA\\VBA7.1\\1041\\VbLR6.chm",1000513]"########,
        r########"[["III",3,0],301,202,0,"","","",0]"########,
    ]);
}

#[test]
fn native_predeclared_rebound_held_success() {
    compare_japanese(r########"Attribute VB_Name = "Predeclared5288"
Option Explicit
Public trace As String,initCount As Long,termCount As Long
Public held As Object
Public Sub MarkInit()
 trace=trace & "I":initCount=initCount+1
End Sub
Public Sub NoteSelf(ByVal same As Boolean)
 If same Then trace=trace & "S" Else trace=trace & "D"
End Sub
Public Sub MarkTerm()
 trace=trace & "T":termCount=termCount+1
End Sub
Public Function Probe() As Variant
 Dim first As Long,second As Long,before As Variant
 
 On Error Resume Next
 Err.Clear
 first=Automatic5288.Bump():second=held.Bump()
 before=Array(trace,initCount,termCount)
 Probe=Array(before,first,second,Err.Number,Err.Description,Err.Source,Err.HelpFile,Err.HelpContext)
End Function
Attribute VB_Name = "Other5299"
Option Explicit
Public Function Touch() As Long
 Touch=Automatic5288.Bump()
End Function
VERSION 1.0 CLASS
Attribute VB_Name = "Automatic5288"
Attribute VB_PredeclaredId = True
Option Explicit
Private counter As Long,tag As Long
Private Sub Class_Initialize()
 Predeclared5288.MarkInit
 tag=Predeclared5288.initCount
 If Predeclared5288.initCount=1 Then Set Automatic5288=New Automatic5288: Set Predeclared5288.held=Automatic5288
End Sub
Private Sub Class_Terminate()
 Predeclared5288.MarkTerm
End Sub
Public Function Bump() As Long
 counter=counter+1:Bump=tag*100+counter
End Function
Public Function SameInstance() As Long
 If Me Is Automatic5288 Then SameInstance=1 Else SameInstance=0
End Function
Public Function Outer() As Long
 Outer=Automatic5288.Bump()
End Function
"########, [
        r########"[["IIT",2,1],201,202,0,"","","",0]"########,
        r########"[["IIT",2,1],203,204,0,"","","",0]"########,
    ]);
}
