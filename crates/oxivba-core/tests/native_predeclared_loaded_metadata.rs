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
fn native_predeclared_loaded_terminate_clear_again() {
    compare_japanese(r########"VERSION 1.0 CLASS
Attribute VB_Name = "ThisWorkbook"
Attribute VB_Base = "0{00020819-0000-0000-C000-000000000046}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = True

VERSION 1.0 CLASS
Attribute VB_Name = "Sheet1"
Attribute VB_Base = "0{00020820-0000-0000-C000-000000000046}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = True

Attribute VB_Name = "Predeclared5288"
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
 first=Automatic5288.Bump():Set Automatic5288=Nothing
 second=Automatic5288.Bump():before=Array(trace,initCount,termCount)
 Probe=Array(before,first,second,Err.Number,Err.Description,Err.Source,Err.HelpFile,Err.HelpContext)
End Function

Attribute VB_Name = "Other5299"
Option Explicit
Public Function Touch() As Long
 Touch=Automatic5288.Bump()
End Function

VERSION 1.0 CLASS
Attribute VB_Name = "Automatic5288"
Attribute VB_Base = "0{FCFB3D2A-A0FA-1068-A738-08002B3371B5}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False
Attribute VB_PredeclaredId = True
Option Explicit
Private counter As Long,tag As Long
Private child As Child5751
Private Sub Class_Initialize()
 Predeclared5288.MarkInit
 tag=Predeclared5288.initCount
 
End Sub
Private Sub Class_Terminate()
 Predeclared5288.MarkTerm
 If Predeclared5288.termCount=1 Then Set Automatic5288=Nothing
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

VERSION 1.0 CLASS
Attribute VB_Name = "Child5751"
Attribute VB_Base = "0{FCFB3D2A-A0FA-1068-A738-08002B3371B5}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False
Attribute VB_PredeclaredId = False
Option Explicit
Private Sub Class_Initialize()
 Predeclared5288.trace=Predeclared5288.trace & "C"
End Sub
Private Sub Class_Terminate()
 Predeclared5288.trace=Predeclared5288.trace & "X"
End Sub

"########, ["[[\"ITI\",2,1],101,201,0,\"\",\"\",\"\",0]","[[\"ITITI\",3,2],202,301,0,\"\",\"\",\"\",0]"]);
}

#[test]
fn native_predeclared_loaded_terminate_rebind_self() {
    compare_japanese(r########"VERSION 1.0 CLASS
Attribute VB_Name = "ThisWorkbook"
Attribute VB_Base = "0{00020819-0000-0000-C000-000000000046}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = True

VERSION 1.0 CLASS
Attribute VB_Name = "Sheet1"
Attribute VB_Base = "0{00020820-0000-0000-C000-000000000046}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = True

Attribute VB_Name = "Predeclared5288"
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
 first=Automatic5288.Bump():Set Automatic5288=Nothing
 second=Automatic5288.Bump():before=Array(trace,initCount,termCount)
 Probe=Array(before,first,second,Err.Number,Err.Description,Err.Source,Err.HelpFile,Err.HelpContext)
End Function

Attribute VB_Name = "Other5299"
Option Explicit
Public Function Touch() As Long
 Touch=Automatic5288.Bump()
End Function

VERSION 1.0 CLASS
Attribute VB_Name = "Automatic5288"
Attribute VB_Base = "0{FCFB3D2A-A0FA-1068-A738-08002B3371B5}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False
Attribute VB_PredeclaredId = True
Option Explicit
Private counter As Long,tag As Long
Private child As Child5751
Private Sub Class_Initialize()
 Predeclared5288.MarkInit
 tag=Predeclared5288.initCount
 
End Sub
Private Sub Class_Terminate()
 Predeclared5288.MarkTerm
 If Predeclared5288.termCount=1 Then Set Automatic5288=Me
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

VERSION 1.0 CLASS
Attribute VB_Name = "Child5751"
Attribute VB_Base = "0{FCFB3D2A-A0FA-1068-A738-08002B3371B5}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False
Attribute VB_PredeclaredId = False
Option Explicit
Private Sub Class_Initialize()
 Predeclared5288.trace=Predeclared5288.trace & "C"
End Sub
Private Sub Class_Terminate()
 Predeclared5288.trace=Predeclared5288.trace & "X"
End Sub

"########, ["[[\"IT\",1,1],101,102,0,\"\",\"\",\"\",0]","[[\"ITI\",2,1],103,201,0,\"\",\"\",\"\",0]"]);
}

#[test]
fn native_predeclared_loaded_terminate_handled_error() {
    compare_japanese(r########"VERSION 1.0 CLASS
Attribute VB_Name = "ThisWorkbook"
Attribute VB_Base = "0{00020819-0000-0000-C000-000000000046}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = True

VERSION 1.0 CLASS
Attribute VB_Name = "Sheet1"
Attribute VB_Base = "0{00020820-0000-0000-C000-000000000046}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = True

Attribute VB_Name = "Predeclared5288"
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
 first=Automatic5288.Bump():Set Automatic5288=Nothing
 second=Automatic5288.Bump():before=Array(trace,initCount,termCount)
 Probe=Array(before,first,second,Err.Number,Err.Description,Err.Source,Err.HelpFile,Err.HelpContext)
End Function

Attribute VB_Name = "Other5299"
Option Explicit
Public Function Touch() As Long
 Touch=Automatic5288.Bump()
End Function

VERSION 1.0 CLASS
Attribute VB_Name = "Automatic5288"
Attribute VB_Base = "0{FCFB3D2A-A0FA-1068-A738-08002B3371B5}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False
Attribute VB_PredeclaredId = True
Option Explicit
Private counter As Long,tag As Long
Private child As Child5751
Private Sub Class_Initialize()
 Predeclared5288.MarkInit
 tag=Predeclared5288.initCount
 
End Sub
Private Sub Class_Terminate()
 Predeclared5288.MarkTerm
 On Error Resume Next
 Err.Raise 5
 Predeclared5288.trace=Predeclared5288.trace & CStr(Err.Number)
 Err.Clear
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

VERSION 1.0 CLASS
Attribute VB_Name = "Child5751"
Attribute VB_Base = "0{FCFB3D2A-A0FA-1068-A738-08002B3371B5}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False
Attribute VB_PredeclaredId = False
Option Explicit
Private Sub Class_Initialize()
 Predeclared5288.trace=Predeclared5288.trace & "C"
End Sub
Private Sub Class_Terminate()
 Predeclared5288.trace=Predeclared5288.trace & "X"
End Sub

"########, ["[[\"IT5I\",2,1],101,201,0,\"\",\"\",\"\",0]","[[\"IT5IT5I\",3,2],202,301,0,\"\",\"\",\"\",0]"]);
}
