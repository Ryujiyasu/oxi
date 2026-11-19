// This Source Code Form is subject to the terms of the Mozilla Public
// License, v. 2.0. If a copy of the MPL was not distributed with this
// file, You can obtain one at https://mozilla.org/MPL/2.0/.

use oxivba_core::{parser::parse_project, runtime::{Runtime, Value}};

// COM text observations do not certify null-BSTR pointer identity.
fn observed(value: &Value) -> String {
    match value {
        Value::String(v) => format!("{v:?}"),
        Value::NullString => "\"\"".into(),
        Value::Integer(v) => v.to_string(),
        Value::Boolean(v) => v.to_string(),
        Value::Empty | Value::Null => "null".into(),
        Value::Array(a) => format!("[{}]", a.values.iter().map(observed).collect::<Vec<_>>().join(",")),
        other => format!("unexpected:{other:?}"),
    }
}

fn compare(source: &str, expected: [&str; 2]) {
    let (module, classes) = parse_project(source).unwrap();
    let mut runtime = Runtime::new(&module).with_classes(&classes);
    for row in expected {
        let value = runtime.call("Predeclared5288.Probe", vec![]).unwrap();
        assert_eq!(observed(&value), row, "continuous runtime");
    }
    let mut state = None;
    for row in expected {
        let mut restored = Runtime::new(&module).with_classes(&classes);
        if let Some(saved) = state.take() { restored = restored.with_project_state(saved); }
        let value = restored.call("Predeclared5288.Probe", vec![]).unwrap();
        assert_eq!(observed(&value), row, "restored project storage");
        state = Some(restored.into_project_state());
    }
}

#[test]
fn native_predeclared_ordinary_new() {
    compare(r########"Attribute VB_Name = "Predeclared5288"
Option Explicit
Public Function Probe() As Variant
 Dim item As Object, first As Long, second As Long, before As Variant
 On Error Resume Next
 Err.Clear
 Set item=New Automatic5288
 before=Array(Err.Number,TypeName(item),VarType(item))
 first=item.Bump()
 second=item.Bump()
 Probe=Array(before,first,second,Err.Number,Err.Description,Err.Source,Err.HelpFile,Err.HelpContext)
End Function
VERSION 1.0 CLASS
Attribute VB_Name = "Automatic5288"
Attribute VB_PredeclaredId = False
Option Explicit
Private counter As Long
Public Function Bump() As Long
 counter=counter+1
 Bump=counter
End Function
"########, [
        r########"[[0,"Automatic5288",9],1,2,0,"","","",0]"########,
        r########"[[0,"Automatic5288",9],1,2,0,"","","",0]"########,
    ]);
}

#[test]
fn native_predeclared_predeclared_new() {
    compare(r########"Attribute VB_Name = "Predeclared5288"
Option Explicit
Public Function Probe() As Variant
 Dim item As Object, first As Long, second As Long, before As Variant
 On Error Resume Next
 Err.Clear
 Set item=New Automatic5288
 before=Array(Err.Number,TypeName(item),VarType(item))
 first=item.Bump()
 second=item.Bump()
 Probe=Array(before,first,second,Err.Number,Err.Description,Err.Source,Err.HelpFile,Err.HelpContext)
End Function
VERSION 1.0 CLASS
Attribute VB_Name = "Automatic5288"
Attribute VB_PredeclaredId = True
Option Explicit
Private counter As Long
Public Function Bump() As Long
 counter=counter+1
 Bump=counter
End Function
"########, [
        r########"[[0,"Automatic5288",9],1,2,0,"","","",0]"########,
        r########"[[0,"Automatic5288",9],1,2,0,"","","",0]"########,
    ]);
}

#[test]
fn native_predeclared_predeclared_direct() {
    compare(r########"Attribute VB_Name = "Predeclared5288"
Option Explicit
Public Function Probe() As Variant
 Dim item As Object, first As Long, second As Long, before As Variant
 On Error Resume Next
 Err.Clear
 Set item=Automatic5288
 before=Array(Err.Number,TypeName(item),VarType(item))
 first=item.Bump()
 second=item.Bump()
 Probe=Array(before,first,second,Err.Number,Err.Description,Err.Source,Err.HelpFile,Err.HelpContext)
End Function
VERSION 1.0 CLASS
Attribute VB_Name = "Automatic5288"
Attribute VB_PredeclaredId = True
Option Explicit
Private counter As Long
Public Function Bump() As Long
 counter=counter+1
 Bump=counter
End Function
"########, [
        r########"[[0,"Automatic5288",9],1,2,0,"","","",0]"########,
        r########"[[0,"Automatic5288",9],3,4,0,"","","",0]"########,
    ]);
}

#[test]
fn native_predeclared_lazy() {
    compare(r########"Attribute VB_Name = "Predeclared5288"
Option Explicit
Public trace As String,initCount As Long,termCount As Long
Public held As Object
Public Sub MarkInit()
 trace=trace & "I":initCount=initCount+1
End Sub
Public Sub MarkTerm()
 trace=trace & "T":termCount=termCount+1
End Sub
Public Function Probe() As Variant
 Dim first As Long,second As Long,before As Variant
 
 On Error Resume Next
 Err.Clear
 before=Array(trace,initCount,termCount)
 first=Automatic5288.Bump()
 second=Automatic5288.Bump()
 before=Array(before,Array(trace,initCount,termCount),VarType(Automatic5288))
 Probe=Array(before,first,second,Err.Number,Err.Description,Err.Source,Err.HelpFile,Err.HelpContext)
End Function
Attribute VB_Name = "Other5296"
Option Explicit
Public Function Touch() As Long
 Touch=Automatic5288.Bump()
End Function
VERSION 1.0 CLASS
Attribute VB_Name = "Automatic5288"
Attribute VB_PredeclaredId = True
Option Explicit
Private counter As Long
Private Sub Class_Initialize()
 Predeclared5288.MarkInit
End Sub
Private Sub Class_Terminate()
 Predeclared5288.MarkTerm
End Sub
Public Function Bump() As Long
 counter=counter+1:Bump=counter
End Function
Public Function SameInstance() As Long
 If Me Is Automatic5288 Then SameInstance=1 Else SameInstance=0
End Function
Public Function Outer() As Long
 Outer=Automatic5288.Bump()
End Function
"########, [
        r########"[[["",0,0],["I",1,0],9],1,2,0,"","","",0]"########,
        r########"[[["I",1,0],["I",1,0],9],3,4,0,"","","",0]"########,
    ]);
}

#[test]
fn native_predeclared_identity() {
    compare(r########"Attribute VB_Name = "Predeclared5288"
Option Explicit
Public trace As String,initCount As Long,termCount As Long
Public held As Object
Public Sub MarkInit()
 trace=trace & "I":initCount=initCount+1
End Sub
Public Sub MarkTerm()
 trace=trace & "T":termCount=termCount+1
End Sub
Public Function Probe() As Variant
 Dim first As Long,second As Long,before As Variant
 Dim a As Object,b As Object,other As Object
 On Error Resume Next
 Err.Clear
 Set a=Automatic5288:Set b=Automatic5288:Set other=New Automatic5288
 first=a.Bump():second=other.Bump()
 before=Array(Array(a Is b,a Is other),Array(trace,initCount,termCount),VarType(a))
 Probe=Array(before,first,second,Err.Number,Err.Description,Err.Source,Err.HelpFile,Err.HelpContext)
End Function
Attribute VB_Name = "Other5296"
Option Explicit
Public Function Touch() As Long
 Touch=Automatic5288.Bump()
End Function
VERSION 1.0 CLASS
Attribute VB_Name = "Automatic5288"
Attribute VB_PredeclaredId = True
Option Explicit
Private counter As Long
Private Sub Class_Initialize()
 Predeclared5288.MarkInit
End Sub
Private Sub Class_Terminate()
 Predeclared5288.MarkTerm
End Sub
Public Function Bump() As Long
 counter=counter+1:Bump=counter
End Function
Public Function SameInstance() As Long
 If Me Is Automatic5288 Then SameInstance=1 Else SameInstance=0
End Function
Public Function Outer() As Long
 Outer=Automatic5288.Bump()
End Function
"########, [
        r########"[[[true,false],["II",2,0],9],1,1,0,"","","",0]"########,
        r########"[[[true,false],["IITI",3,1],9],2,1,0,"","","",0]"########,
    ]);
}

#[test]
fn native_predeclared_self_reference() {
    compare(r########"Attribute VB_Name = "Predeclared5288"
Option Explicit
Public trace As String,initCount As Long,termCount As Long
Public held As Object
Public Sub MarkInit()
 trace=trace & "I":initCount=initCount+1
End Sub
Public Sub MarkTerm()
 trace=trace & "T":termCount=termCount+1
End Sub
Public Function Probe() As Variant
 Dim first As Long,second As Long,before As Variant
 
 On Error Resume Next
 Err.Clear
 first=Automatic5288.SameInstance()
 second=Automatic5288.Outer()
 before=Array(trace,initCount,termCount)
 Probe=Array(before,first,second,Err.Number,Err.Description,Err.Source,Err.HelpFile,Err.HelpContext)
End Function
Attribute VB_Name = "Other5296"
Option Explicit
Public Function Touch() As Long
 Touch=Automatic5288.Bump()
End Function
VERSION 1.0 CLASS
Attribute VB_Name = "Automatic5288"
Attribute VB_PredeclaredId = True
Option Explicit
Private counter As Long
Private Sub Class_Initialize()
 Predeclared5288.MarkInit
End Sub
Private Sub Class_Terminate()
 Predeclared5288.MarkTerm
End Sub
Public Function Bump() As Long
 counter=counter+1:Bump=counter
End Function
Public Function SameInstance() As Long
 If Me Is Automatic5288 Then SameInstance=1 Else SameInstance=0
End Function
Public Function Outer() As Long
 Outer=Automatic5288.Bump()
End Function
"########, [
        r########"[["I",1,0],1,1,0,"","","",0]"########,
        r########"[["I",1,0],1,2,0,"","","",0]"########,
    ]);
}

#[test]
fn native_predeclared_shadow() {
    compare(r########"Attribute VB_Name = "Predeclared5288"
Option Explicit
Public trace As String,initCount As Long,termCount As Long
Public held As Object
Public Sub MarkInit()
 trace=trace & "I":initCount=initCount+1
End Sub
Public Sub MarkTerm()
 trace=trace & "T":termCount=termCount+1
End Sub
Public Function Probe() As Variant
 Dim first As Long,second As Long,before As Variant
 Dim Automatic5288 As Object
 On Error Resume Next
 Err.Clear
 Set Automatic5288=New Automatic5288
 first=Automatic5288.Bump():second=Automatic5288.Bump()
 before=Array(trace,initCount,termCount)
 Probe=Array(before,first,second,Err.Number,Err.Description,Err.Source,Err.HelpFile,Err.HelpContext)
End Function
Attribute VB_Name = "Other5296"
Option Explicit
Public Function Touch() As Long
 Touch=Automatic5288.Bump()
End Function
VERSION 1.0 CLASS
Attribute VB_Name = "Automatic5288"
Attribute VB_PredeclaredId = True
Option Explicit
Private counter As Long
Private Sub Class_Initialize()
 Predeclared5288.MarkInit
End Sub
Private Sub Class_Terminate()
 Predeclared5288.MarkTerm
End Sub
Public Function Bump() As Long
 counter=counter+1:Bump=counter
End Function
Public Function SameInstance() As Long
 If Me Is Automatic5288 Then SameInstance=1 Else SameInstance=0
End Function
Public Function Outer() As Long
 Outer=Automatic5288.Bump()
End Function
"########, [
        r########"[["I",1,0],1,2,0,"","","",0]"########,
        r########"[["ITI",2,1],1,2,0,"","","",0]"########,
    ]);
}

#[test]
fn native_predeclared_cross_module() {
    compare(r########"Attribute VB_Name = "Predeclared5288"
Option Explicit
Public trace As String,initCount As Long,termCount As Long
Public held As Object
Public Sub MarkInit()
 trace=trace & "I":initCount=initCount+1
End Sub
Public Sub MarkTerm()
 trace=trace & "T":termCount=termCount+1
End Sub
Public Function Probe() As Variant
 Dim first As Long,second As Long,before As Variant
 
 On Error Resume Next
 Err.Clear
 first=Automatic5288.Bump():second=Other5296.Touch()
 before=Array(trace,initCount,termCount)
 Probe=Array(before,first,second,Err.Number,Err.Description,Err.Source,Err.HelpFile,Err.HelpContext)
End Function
Attribute VB_Name = "Other5296"
Option Explicit
Public Function Touch() As Long
 Touch=Automatic5288.Bump()
End Function
VERSION 1.0 CLASS
Attribute VB_Name = "Automatic5288"
Attribute VB_PredeclaredId = True
Option Explicit
Private counter As Long
Private Sub Class_Initialize()
 Predeclared5288.MarkInit
End Sub
Private Sub Class_Terminate()
 Predeclared5288.MarkTerm
End Sub
Public Function Bump() As Long
 counter=counter+1:Bump=counter
End Function
Public Function SameInstance() As Long
 If Me Is Automatic5288 Then SameInstance=1 Else SameInstance=0
End Function
Public Function Outer() As Long
 Outer=Automatic5288.Bump()
End Function
"########, [
        r########"[["I",1,0],1,2,0,"","","",0]"########,
        r########"[["I",1,0],3,4,0,"","","",0]"########,
    ]);
}

#[test]
fn native_predeclared_held() {
    compare(r########"Attribute VB_Name = "Predeclared5288"
Option Explicit
Public trace As String,initCount As Long,termCount As Long
Public held As Object
Public Sub MarkInit()
 trace=trace & "I":initCount=initCount+1
End Sub
Public Sub MarkTerm()
 trace=trace & "T":termCount=termCount+1
End Sub
Public Function Probe() As Variant
 Dim first As Long,second As Long,before As Variant
 
 On Error Resume Next
 Err.Clear
 If held Is Nothing Then Set held=Automatic5288
 first=held.Bump():second=Automatic5288.Bump()
 before=Array(trace,initCount,termCount)
 Probe=Array(before,first,second,Err.Number,Err.Description,Err.Source,Err.HelpFile,Err.HelpContext)
End Function
Attribute VB_Name = "Other5296"
Option Explicit
Public Function Touch() As Long
 Touch=Automatic5288.Bump()
End Function
VERSION 1.0 CLASS
Attribute VB_Name = "Automatic5288"
Attribute VB_PredeclaredId = True
Option Explicit
Private counter As Long
Private Sub Class_Initialize()
 Predeclared5288.MarkInit
End Sub
Private Sub Class_Terminate()
 Predeclared5288.MarkTerm
End Sub
Public Function Bump() As Long
 counter=counter+1:Bump=counter
End Function
Public Function SameInstance() As Long
 If Me Is Automatic5288 Then SameInstance=1 Else SameInstance=0
End Function
Public Function Outer() As Long
 Outer=Automatic5288.Bump()
End Function
"########, [
        r########"[["I",1,0],1,2,0,"","","",0]"########,
        r########"[["I",1,0],3,4,0,"","","",0]"########,
    ]);
}

#[test]
fn native_predeclared_init_self() {
    compare(r########"Attribute VB_Name = "Predeclared5288"
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
 before=Array(trace,initCount,termCount)
 first=Automatic5288.Bump():second=Automatic5288.Bump()
 before=Array(before,Array(trace,initCount,termCount),VarType(Automatic5288))
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
Private counter As Long
Private Sub Class_Initialize()
 Predeclared5288.MarkInit
 If Predeclared5288.initCount<3 Then Predeclared5288.NoteSelf (Me Is Automatic5288)
End Sub
Private Sub Class_Terminate()
 Predeclared5288.MarkTerm
End Sub
Public Function Bump() As Long
 counter=counter+1:Bump=counter
End Function
Public Function SameInstance() As Long
 If Me Is Automatic5288 Then SameInstance=1 Else SameInstance=0
End Function
Public Function Outer() As Long
 Outer=Automatic5288.Bump()
End Function
"########, [
        r########"[[["",0,0],["IS",1,0],9],1,2,0,"","","",0]"########,
        r########"[[["IS",1,0],["IS",1,0],9],3,4,0,"","","",0]"########,
    ]);
}

#[test]
fn native_predeclared_init_error() {
    compare(r########"Attribute VB_Name = "Predeclared5288"
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
 Dim savedErr As Variant
 On Error Resume Next
 Err.Clear
 first=Automatic5288.Bump()
 savedErr=Array(Err.Number,Err.Description,Err.Source,Err.HelpFile,Err.HelpContext)
 Err.Clear
 second=Automatic5288.Bump()
 before=Array(savedErr,Array(trace,initCount,termCount),VarType(Automatic5288))
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
Private counter As Long
Private Sub Class_Initialize()
 Predeclared5288.MarkInit
 If Predeclared5288.initCount=1 Then Err.Raise 513,"init-source","init-failure","init-help",73
End Sub
Private Sub Class_Terminate()
 Predeclared5288.MarkTerm
End Sub
Public Function Bump() As Long
 counter=counter+1:Bump=counter
End Function
Public Function SameInstance() As Long
 If Me Is Automatic5288 Then SameInstance=1 Else SameInstance=0
End Function
Public Function Outer() As Long
 Outer=Automatic5288.Bump()
End Function
"########, [
        r########"[[[513,"init-failure","init-source","init-help",73],["II",2,0],9],0,1,0,"","","",0]"########,
        r########"[[[0,"","","",0],["II",2,0],9],2,3,0,"","","",0]"########,
    ]);
}


#[test]
fn native_predeclared_holdout_clear() {
    compare(r########"Attribute VB_Name = "Predeclared5288"
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
Public Function TouchRef(ByRef item As Automatic5288) As Long
 TouchRef=item.Bump()
End Function
Public Function TouchVal(ByVal item As Automatic5288) As Long
 TouchVal=item.Bump()
End Function
Public Function Probe() As Variant
 Dim first As Long,second As Long,before As Variant
 
 On Error Resume Next
 Err.Clear
 Set held=Automatic5288
 first=held.Bump()
 Set Automatic5288=Nothing
 second=Automatic5288.Bump()
 before=Array(held Is Automatic5288,held.Bump(),Array(trace,initCount,termCount))
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
Private counter As Long
Private Sub Class_Initialize()
 Predeclared5288.MarkInit
 
End Sub
Private Sub Class_Terminate()
 Predeclared5288.MarkTerm
End Sub
Public Function Bump() As Long
 counter=counter+1:Bump=counter
End Function
Public Function SameInstance() As Long
 If Me Is Automatic5288 Then SameInstance=1 Else SameInstance=0
End Function
Public Function Outer() As Long
 Outer=Automatic5288.Bump()
End Function
"########, [
        r########"[[false,2,["II",2,0]],1,1,0,"","","",0]"########,
        r########"[[false,3,["IITI",3,1]],2,1,0,"","","",0]"########
    ]);
}

#[test]
fn native_predeclared_holdout_rebind() {
    compare(r########"Attribute VB_Name = "Predeclared5288"
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
Public Function TouchRef(ByRef item As Automatic5288) As Long
 TouchRef=item.Bump()
End Function
Public Function TouchVal(ByVal item As Automatic5288) As Long
 TouchVal=item.Bump()
End Function
Public Function Probe() As Variant
 Dim first As Long,second As Long,before As Variant
 Dim fresh As Automatic5288
 On Error Resume Next
 Err.Clear
 Set held=Automatic5288
 first=held.Bump()
 Set fresh=New Automatic5288
 Set Automatic5288=fresh
 second=Automatic5288.Bump()
 before=Array(held Is Automatic5288,fresh Is Automatic5288,Array(trace,initCount,termCount))
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
Private counter As Long
Private Sub Class_Initialize()
 Predeclared5288.MarkInit
 
End Sub
Private Sub Class_Terminate()
 Predeclared5288.MarkTerm
End Sub
Public Function Bump() As Long
 counter=counter+1:Bump=counter
End Function
Public Function SameInstance() As Long
 If Me Is Automatic5288 Then SameInstance=1 Else SameInstance=0
End Function
Public Function Outer() As Long
 Outer=Automatic5288.Bump()
End Function
"########, [
        r########"[[false,true,["II",2,0]],1,1,0,"","","",0]"########,
        r########"[[false,true,["IITI",3,1]],2,1,0,"","","",0]"########
    ]);
}

#[test]
fn native_predeclared_holdout_typed_byref() {
    compare(r########"Attribute VB_Name = "Predeclared5288"
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
Public Function TouchRef(ByRef item As Automatic5288) As Long
 TouchRef=item.Bump()
End Function
Public Function TouchVal(ByVal item As Automatic5288) As Long
 TouchVal=item.Bump()
End Function
Public Function Probe() As Variant
 Dim first As Long,second As Long,before As Variant
 
 On Error Resume Next
 Err.Clear
 first=TouchRef(Automatic5288)
 second=Automatic5288.Bump()
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
Private counter As Long
Private Sub Class_Initialize()
 Predeclared5288.MarkInit
 
End Sub
Private Sub Class_Terminate()
 Predeclared5288.MarkTerm
End Sub
Public Function Bump() As Long
 counter=counter+1:Bump=counter
End Function
Public Function SameInstance() As Long
 If Me Is Automatic5288 Then SameInstance=1 Else SameInstance=0
End Function
Public Function Outer() As Long
 Outer=Automatic5288.Bump()
End Function
"########, [
        r########"[["I",1,0],1,2,0,"","","",0]"########,
        r########"[["I",1,0],3,4,0,"","","",0]"########
    ]);
}

#[test]
fn native_predeclared_holdout_typed_byval() {
    compare(r########"Attribute VB_Name = "Predeclared5288"
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
Public Function TouchRef(ByRef item As Automatic5288) As Long
 TouchRef=item.Bump()
End Function
Public Function TouchVal(ByVal item As Automatic5288) As Long
 TouchVal=item.Bump()
End Function
Public Function Probe() As Variant
 Dim first As Long,second As Long,before As Variant
 
 On Error Resume Next
 Err.Clear
 first=TouchVal(Automatic5288)
 second=Automatic5288.Bump()
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
Private counter As Long
Private Sub Class_Initialize()
 Predeclared5288.MarkInit
 
End Sub
Private Sub Class_Terminate()
 Predeclared5288.MarkTerm
End Sub
Public Function Bump() As Long
 counter=counter+1:Bump=counter
End Function
Public Function SameInstance() As Long
 If Me Is Automatic5288 Then SameInstance=1 Else SameInstance=0
End Function
Public Function Outer() As Long
 Outer=Automatic5288.Bump()
End Function
"########, [
        r########"[["I",1,0],1,2,0,"","","",0]"########,
        r########"[["I",1,0],3,4,0,"","","",0]"########
    ]);
}

// Match the measured Excel error language and help-file environment explicitly.
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
fn native_predeclared_storage_wrong_object() {
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
Public Function ClearRef(ByRef item As Automatic5288) As Long
 ClearRef=item.Bump()
 Set item=Nothing
End Function
Public Function TouchRef(ByRef item As Automatic5288) As Long
 TouchRef=item.Bump()
End Function
Public Function TouchVal(ByVal item As Automatic5288) As Long
 TouchVal=item.Bump()
End Function
Public Function Probe() As Variant
 Dim first As Long,second As Long,before As Variant
 Dim wrong As Object, savedErr As Variant
 On Error Resume Next
 Err.Clear
 Set wrong=New Collection
 Set Automatic5288=wrong
 savedErr=Array(Err.Number,Err.Description,Err.Source,Err.HelpFile,Err.HelpContext)
 Err.Clear
 first=Automatic5288.Bump():second=Automatic5288.Bump()
 before=Array(savedErr,TypeName(Automatic5288),Array(trace,initCount,termCount))
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
Private counter As Long
Private Sub Class_Initialize()
 Predeclared5288.MarkInit
 
End Sub
Private Sub Class_Terminate()
 Predeclared5288.MarkTerm
End Sub
Public Function Bump() As Long
 counter=counter+1:Bump=counter
End Function
Public Function SameInstance() As Long
 If Me Is Automatic5288 Then SameInstance=1 Else SameInstance=0
End Function
Public Function Outer() As Long
 Outer=Automatic5288.Bump()
End Function
"########, [
        r########"[[[13,"型が一致しません。","VBAProject","C:\\Program Files\\Common Files\\Microsoft Shared\\VBA\\VBA7.1\\1041\\VbLR6.chm",1000013],"Automatic5288",["I",1,0]],1,2,0,"","","",0]"########,
        r########"[[[13,"型が一致しません。","VBAProject","C:\\Program Files\\Common Files\\Microsoft Shared\\VBA\\VBA7.1\\1041\\VbLR6.chm",1000013],"Automatic5288",["I",1,0]],3,4,0,"","","",0]"########
    ]);
}

#[test]
fn native_predeclared_storage_byref_clear() {
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
Public Function ClearRef(ByRef item As Automatic5288) As Long
 ClearRef=item.Bump()
 Set item=Nothing
End Function
Public Function TouchRef(ByRef item As Automatic5288) As Long
 TouchRef=item.Bump()
End Function
Public Function TouchVal(ByVal item As Automatic5288) As Long
 TouchVal=item.Bump()
End Function
Public Function Probe() As Variant
 Dim first As Long,second As Long,before As Variant
 
 On Error Resume Next
 Err.Clear
 Set held=Automatic5288
 first=ClearRef(Automatic5288)
 second=Automatic5288.Bump()
 before=Array(held Is Automatic5288,held.Bump(),Array(trace,initCount,termCount))
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
Private counter As Long
Private Sub Class_Initialize()
 Predeclared5288.MarkInit
 
End Sub
Private Sub Class_Terminate()
 Predeclared5288.MarkTerm
End Sub
Public Function Bump() As Long
 counter=counter+1:Bump=counter
End Function
Public Function SameInstance() As Long
 If Me Is Automatic5288 Then SameInstance=1 Else SameInstance=0
End Function
Public Function Outer() As Long
 Outer=Automatic5288.Bump()
End Function
"########, [
        r########"[[false,2,["II",2,0]],1,1,0,"","","",0]"########,
        r########"[[false,3,["IITI",3,1]],2,1,0,"","","",0]"########
    ]);
}


#[test]
fn native_predeclared_storage_init_clear() {
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
 first=Automatic5288.Bump():second=Automatic5288.Bump()
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
 If Predeclared5288.initCount=1 Then Set Automatic5288=Nothing
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
        r########"[["ITI",2,1],0,201,91,"オブジェクト変数または With ブロック変数が設定されていません。","VBAProject","C:\\Program Files\\Common Files\\Microsoft Shared\\VBA\\VBA7.1\\1041\\VbLR6.chm",1000091]"########,
        r########"[["ITI",2,1],202,203,0,"","","",0]"########
    ]);
}

#[test]
fn native_predeclared_storage_init_rebind() {
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
 first=Automatic5288.Bump():second=Automatic5288.Bump()
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
 If Predeclared5288.initCount=1 Then Set Automatic5288=New Automatic5288
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
        r########"[["IIT",2,1],203,204,0,"","","",0]"########
    ]);
}

#[test]
fn native_predeclared_failed_init_clear() {
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
 first=Automatic5288.Bump():second=Automatic5288.Bump()
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
 If Predeclared5288.initCount=1 Then Set Automatic5288=Nothing: Err.Raise 513, "Ctor", "failed after clear"
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
        r########"[["II",2,0],0,201,513,"failed after clear","Ctor","C:\\Program Files\\Common Files\\Microsoft Shared\\VBA\\VBA7.1\\1041\\VbLR6.chm",1000513]"########,
        r########"[["II",2,0],202,203,0,"","","",0]"########,
    ]);
}

#[test]
fn native_predeclared_failed_init_rebind() {
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
 first=Automatic5288.Bump():second=Automatic5288.Bump()
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
 If Predeclared5288.initCount=1 Then Set Automatic5288=New Automatic5288: Err.Raise 513, "Ctor", "failed after rebind"
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
        r########"[["IITI",3,1],0,301,513,"failed after rebind","Ctor","C:\\Program Files\\Common Files\\Microsoft Shared\\VBA\\VBA7.1\\1041\\VbLR6.chm",1000513]"########,
        r########"[["IITI",3,1],302,303,0,"","","",0]"########,
    ]);
}
