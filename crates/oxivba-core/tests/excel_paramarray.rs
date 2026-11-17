// This Source Code Form is subject to the terms of the Mozilla Public
// License, v. 2.0. If a copy of the MPL was not distributed with this
// file, You can obtain one at https://mozilla.org/MPL/2.0/.

use oxivba_core::{parse_project, Runtime, Value};

#[test]
fn paramarray_len_and_nested_arrays_match_excel() {
    let fixtures = std::path::Path::new(env!("CARGO_MANIFEST_DIR")).join("../../tests/fixtures/vba");
    let rows: Vec<Vec<_>> = include_str!("../../../tests/fixtures/vba/paramarray.tsv").lines().map(|row| row.split('\t').collect()).collect();
    let mut at = 0;
    while at < rows.len() {
        let name = rows[at][0];
        let source = std::fs::read_to_string(fixtures.join(rows[at][1])).unwrap();
        let (module, classes) = parse_project(&source).unwrap();
        let mut continuous = Runtime::new(&module).with_classes(&classes);
        let mut previous = None;
        while at < rows.len() && rows[at][0] == name {
            let fields = &rows[at];
            let expected = Value::String(fields[3].into());
            assert_eq!(continuous.call(fields[2], vec![]).unwrap_or_else(|failure| panic!("continuous {name}: {failure:?}")), expected, "continuous {name}");
            let (module, classes) = parse_project(&source).unwrap();
            let mut runtime = Runtime::new(&module).with_classes(&classes);
            if let Some(saved) = previous.take() { runtime = runtime.with_project_state(saved); }
            assert_eq!(runtime.call(fields[2], vec![]).unwrap_or_else(|failure| panic!("resumed {name}: {failure:?}")), expected, "resumed {name}");
            previous = Some(runtime.into_project_state());
            at += 1;
        }
    }
}


fn array_report(source: &str) -> Vec<Value> {
    let (module, classes) = parse_project(source).unwrap();
    match Runtime::new(&module).with_classes(&classes).call("Probe", vec![]).unwrap_or_else(|error| panic!("{error:?}; operation={:?}", source.lines().find(|line| line.contains("sink=")))) {
        Value::Array(array) => array.values,
        value => panic!("expected report array, received {value:?}"),
    }
}

fn boxing_report(element: &str, operation: &str) -> Vec<Value> {
    let (declaration, value) = if element == "record" {
        ("Dim holder As ArrayRecord".into(), "holder.data")
    } else {
        (format!("Dim data() As {element}"), "data")
    };
    array_report(&format!(r#"
Attribute VB_Name = "Boxing"
Private Type ArrayRecord
 data() As Object
End Type
Private Function TakeValue(ByVal value As Variant) As Long
 TakeValue=1
End Function
Private Function TakeReference(ByRef value As Variant) As Long
 TakeReference=1
End Function
Private Function TakeTyped(ByRef value() As Object) As Long
 TakeTyped=1
End Function
Private Function TakePack(ParamArray values() As Variant) As Long
 TakePack=1
End Function
Public Function Probe() As Variant
 {declaration}
 Dim copy() As Object,sink As Variant,lo As Long,hi As Long
 Dim beforeErr As Long,actionErr As Long,loErr As Long,hiErr As Long,sinkType As Long
 On Error Resume Next
 Err.Clear
 lo=LBound({value}):beforeErr=Err.Number
 Err.Clear
 {operation}
 actionErr=Err.Number:sinkType=VarType(sink)
 Err.Clear
 lo=LBound({value}):loErr=Err.Number
 Err.Clear
 hi=UBound({value}):hiErr=Err.Number
 Probe=Array(beforeErr,actionErr,sinkType,lo,loErr,hi,hiErr)
End Function
VERSION 1.0 CLASS
Attribute VB_Name = "ArrayItem"
Public Value As Long
"#,operation=operation.replace('$',value)))
}

#[test]
fn variant_array_boxing_preserves_element_types_and_original_places() {
    for (element, direct_type, materialized) in [
        ("Object",8201,true),("ArrayItem",8201,true),("record",8201,true),
        ("Variant",8204,false),("Long",8195,false),("String",8200,false),
    ] {
        for (operation, value_type) in [
            ("sink=$",direct_type),("sink=TakeValue($)",3),("sink=TakeReference($)",3),
            ("sink=Array($)",8204),("sink=TakePack($)",3),("sink=TypeName($)",8),
        ] {
            let expected = if materialized { vec![9,0,value_type,0,0,-1,0] } else { vec![9,0,value_type,0,9,0,9] };
            assert_eq!(boxing_report(element,operation),expected.into_iter().map(Value::Integer).collect::<Vec<_>>(),"{element}: {operation}");
        }
    }
    for (operation, value_type) in [
        ("sink=VarType($)",3),("sink=IsArray($)",11),("sink=IsObject($)",11),
        ("sink=IsEmpty($)",11),("sink=IsNull($)",11),("sink=CVar($)",8201),
        ("sink=($)",8201),("sink=TakeValue(($))",3),("sink=TakeReference(($))",3),
    ] {
        assert_eq!(boxing_report("Object",operation),[9,0,value_type,0,0,-1,0].into_iter().map(Value::Integer).collect::<Vec<_>>(),"{operation}");
    }
}

#[test]
fn variant_array_boxing_keeps_typed_array_paths_uninitialized() {
    for (operation, value_type) in [("copy=$:sink=1",2),("sink=TakeTyped($)",3)] {
        assert_eq!(boxing_report("Object",operation),[9,0,value_type,0,9,0,9].into_iter().map(Value::Integer).collect::<Vec<_>>(),"{operation}");
    }
}

fn timing_report(operation: &str) -> Vec<Value> {
    array_report(&format!(r#"
Attribute VB_Name = "ArrayTiming"
Public Trace As String
Private Type BoxRecord
 value As Variant
End Type
Private Function Grow(ByRef data() As Object) As Long
 ReDim data(1 To 2)
 Trace=Trace & "G;"
 Grow=0
End Function
Private Function TakeValue(ByVal value As Variant,ByVal other As Variant) As Variant
 TakeValue=value
End Function
Private Function TakeReference(ByRef value As Variant,ByVal other As Variant) As Variant
 TakeReference=value
End Function
Private Function PackFirst(ParamArray values() As Variant) As Variant
 PackFirst=values(0)
End Function
Public Property Let Target(ByVal first As Long,ByVal second As Long,ByVal value As Variant)
 Trace=Trace & "B;"
End Property
Public Function Probe() As Variant
 Dim data() As Object,sink As Variant,zero As Double,large As Double,huge As Double,bad As Variant
 Dim actionErr As Long,sinkType As Long,sl As Long,su As Long,sle As Long,sue As Long,dl As Long,du As Long,dle As Long,due As Long
 Trace="":large=1e308:huge=1e20:bad="bad"
 On Error Resume Next
 Err.Clear
 {operation}
 actionErr=Err.Number:sinkType=VarType(sink)
 If IsArray(sink) Then
  Err.Clear
  sl=LBound(sink):sle=Err.Number
  Err.Clear
  su=UBound(sink):sue=Err.Number
  sink="array"
 End If
 Err.Clear
 dl=LBound(data):dle=Err.Number
 Err.Clear
 du=UBound(data):due=Err.Number
 Probe=Array(actionErr,sinkType,CStr(sink),sl,sle,su,sue,dl,dle,du,due,Trace)
End Function
VERSION 1.0 CLASS
Attribute VB_Name = "ArrayWorker"
Public Function TakeValue(ByVal value As Variant,ByVal other As Variant) As Variant
 TakeValue=value
End Function
Public Function TakeTyped(ByRef value() As Object,ByVal other As Variant) As Long
 TakeTyped=1
End Function
Public Function TakeNumber(ByVal value As Long,ByVal other As Variant) As Long
 TakeNumber=1
End Function
Public Property Let Target(ByVal first As Long,ByVal second As Long,ByVal value As Variant)
 ArrayTiming.Trace=ArrayTiming.Trace & "B;"
End Property
Public Property Let TargetByte(ByVal first As Long,ByVal second As Long,ByVal value As Byte)
 ArrayTiming.Trace=ArrayTiming.Trace & "B;"
End Property
"#))
}

fn timing_expected(error: i64, value_type: i64, text: &str, sink: i8, original: i8, trace: &str) -> Vec<Value> {
    let bounds = |state| match state {
        -1 => [0,0,0,0], 0 => [0,9,0,9], 1 => [0,0,-1,0], 2 => [1,0,2,0], _ => unreachable!(),
    };
    let mut values=vec![Value::Integer(error),Value::Integer(value_type),Value::String(text.into())];
    values.extend(bounds(sink).into_iter().map(Value::Integer));
    values.extend(bounds(original).into_iter().map(Value::Integer));
    values.push(Value::String(trace.into()));
    values
}

#[test]
fn variant_array_boxing_keeps_fixed_call_operand_order() {
    for (operation,error) in [
        ("sink=IIf(True,1/zero,large*large)",6),("sink=IIf(True,large*large,1/zero)",11),
        ("sink=IIf(1/zero,large*large,0)",6),("sink=IIf(CInt(huge),1/zero,0)",11),
    ] { assert_eq!(timing_report(operation),timing_expected(error,0,"",-1,0,""),"{operation}"); }
    for operation in ["sink=IIf(True,data,Grow(data))","sink=IIf(False,Grow(data),data)","sink=TakeValue(data,Grow(data))","sink=TakeReference(data,Grow(data))"] {
        assert_eq!(timing_report(operation),timing_expected(0,8201,"array",2,2,"G;"),"{operation}");
    }
    for operation in ["sink=TakeValue(1/zero,Grow(data))","sink=IIf(True,1/zero,Grow(data))"] {
        assert_eq!(timing_report(operation),timing_expected(11,0,"",-1,2,"G;"),"{operation}");
    }
    assert_eq!(timing_report("sink=IIf(False,data,1)"),timing_expected(0,2,"1",-1,1,""));
}

#[test]
fn variant_array_boxing_keeps_variadic_operand_order() {
    for operation in ["sink=Array(1/zero,Grow(data))","sink=Choose(1,1/zero,Grow(data))","sink=PackFirst(1/zero,Grow(data))","sink=Switch(True,1/zero,False,Grow(data))"] {
        assert_eq!(timing_report(operation),timing_expected(11,0,"",-1,0,""),"{operation}");
    }
    assert_eq!(timing_report("sink=PackFirst(large*large,Grow(data))"),timing_expected(6,0,"",-1,0,""));
    assert_eq!(timing_report("sink=PackFirst(data,1/zero,Grow(data))"),timing_expected(11,0,"",-1,1,""));
    for operation in ["sink=Choose(1,data,Grow(data))","sink=Array(data,Grow(data)):sink=sink(0)","sink=PackFirst(data,Grow(data))"] {
        assert_eq!(timing_report(operation),timing_expected(0,8201,"array",2,2,"G;"),"{operation}");
    }
    for operation in ["sink=Choose(2,data,1)","sink=Switch(False,data,True,1)"] {
        assert_eq!(timing_report(operation),timing_expected(0,2,"1",-1,1,""),"{operation}");
    }
}

#[test]
fn variant_array_boxing_preserves_grouped_snapshots() {
    for operation in ["sink=Array((data),Grow(data)):sink=sink(0)","sink=Choose(1,(data),Grow(data))","sink=PackFirst((data),Grow(data))"] {
        assert_eq!(timing_report(operation),timing_expected(0,8201,"array",1,2,"G;"),"{operation}");
    }
    assert_eq!(timing_report("sink=IIf(True,(data),Grow(data))"),timing_expected(0,8201,"array",2,2,"G;"));
}

#[test]
fn variant_array_boxing_keeps_late_call_boundaries() {
    let declare="Dim worker As Object:Set worker=New ArrayWorker:";
    for operation in ["sink=worker.TakeValue(data,LBound(data))","sink=CallByName(worker,\"TakeValue\",1,data,LBound(data))"] {
        assert_eq!(timing_report(&(declare.to_owned()+operation)),timing_expected(0,8201,"array",1,1,""),"{operation}");
    }
    assert_eq!(timing_report(&(declare.to_owned()+"sink=worker.TakeTyped(data,LBound(data))")),timing_expected(0,3,"1",-1,1,""));
    assert_eq!(timing_report(&(declare.to_owned()+"sink=worker.TakeNumber(data,LBound(data))")),timing_expected(13,0,"",-1,1,""));
    assert_eq!(timing_report(&(declare.to_owned()+"sink=worker.TakeValue(data,1/zero)")),timing_expected(11,0,"",-1,1,""));
    assert_eq!(timing_report(&(declare.to_owned()+"sink=worker.TakeValue(1/zero,TypeName(data))")),timing_expected(11,0,"",-1,0,""));
    assert_eq!(timing_report("Dim worker As ArrayWorker:Set worker=New ArrayWorker:sink=worker.TakeValue(data,LBound(data))"),timing_expected(9,0,"",-1,0,""));
}

#[test]
fn variant_array_boxing_preserves_assignment_argument_effects() {
    let declare="Dim worker As Object:Set worker=New ArrayWorker:";
    for operation in ["worker.Target(bad,0)=data","worker.TargetByte(bad,0)=data","sink=CallByName(worker,\"Target\",4,bad,0,data)"] {
        assert_eq!(timing_report(&(declare.to_owned()+operation)),timing_expected(13,0,"",-1,1,""),"{operation}");
    }
    assert_eq!(timing_report(&(declare.to_owned()+"worker.Missing=data")),timing_expected(438,0,"",-1,1,""));
    assert_eq!(timing_report(&(declare.to_owned()+"worker.Target(1/zero,0)=data")),timing_expected(11,0,"",-1,0,""));
    for operation in ["Target(bad,0)=data","Dim worker As ArrayWorker:Set worker=New ArrayWorker:worker.Target(bad,0)=data"] {
        assert_eq!(timing_report(operation),timing_expected(13,0,"",-1,1,""),"{operation}");
    }
    for operation in ["Dim holder(0 To 0) As Variant:holder(0)=data:sink=1","Dim record As BoxRecord:record.value=data:sink=1"] {
        assert_eq!(timing_report(operation),timing_expected(0,2,"1",-1,1,""),"{operation}");
    }
}


#[test]
fn variant_array_boxing_preserves_erased_variant_state() {
    let source = r#"
Public Function Probe() As Variant
 Dim data() As Object,value As Variant,ignored As Variant,lo As Long,number As Long
 ReDim data(0 To 1)
 value=data
 Erase value
 ignored=TypeName(value)
 ignored=VarType(value)
 ignored=IsArray(value)
 On Error Resume Next
 Err.Clear
 lo=LBound(value):number=Err.Number
 Probe=Array(VarType(value),number)
End Function
"#;
    assert_eq!(array_report(source),vec![Value::Integer(8201),Value::Integer(9)]);
}


fn grouped_variant_report(operation: &str) -> Vec<Value> {
    let source = format!(r#"
Attribute VB_Name = "GroupedProbe"
Option Explicit
Public Trace As String
Public Function Describe(ByVal value As Variant) As Variant
 Describe=Array(VarType(value),TypeName(value),IsObject(value))
End Function
Private Function TakeValue(ByVal value As Variant,Optional ByVal other As Variant) As Variant
 TakeValue=Describe(value)
End Function
Private Function TakeRef(ByRef value As Variant,Optional ByVal other As Variant) As Variant
 TakeRef=Describe(value)
End Function
Private Function TakePack(ParamArray values() As Variant) As Variant
 TakePack=Describe(values(0))
End Function
Public Function Later() As Long
 Trace=Trace & "L;":Later=0
End Function
Public Function Change(ByRef value As Variant) As Long
 Trace=Trace & "L;":Set value=Nothing:Change=0
End Function
Public Function Probe() As Variant
 Dim typed As GroupedPayload,obj As Object,boxed As Variant,receiver As GroupedReceiver,late As Object
 Dim answer As Variant,held As Variant,number As Long,errors As Variant,zero As Double
 Set typed=New GroupedPayload:Set obj=typed:Set boxed=obj
 Set receiver=New GroupedReceiver:Set late=receiver:Trace=""
 On Error Resume Next
 Err.Clear
 {operation}
 number=Err.Number
 errors=Array(Err.Number,Err.Description,Err.Source,Err.HelpFile,Err.HelpContext)
 Probe=Array(Array(number,answer,Trace),errors)
End Function
VERSION 1.0 CLASS
Attribute VB_Name = "GroupedPayload"
Public Property Get Value() As Long
Attribute Value.VB_UserMemId = 0
 GroupedProbe.Trace=GroupedProbe.Trace & "G;"
 Value=7
End Property
VERSION 1.0 CLASS
Attribute VB_Name = "GroupedReceiver"
Public Function TakeValue(ByVal value As Variant,Optional ByVal other As Variant) As Variant
 TakeValue=GroupedProbe.Describe(value)
End Function
"#);
    let report = array_report(&source);
    match &report[0] {
        Value::Array(array) => array.values.clone(),
        other => panic!("expected grouped argument report, received {other:?}"),
    }
}

fn check_grouped_variant_report(operation: &str, number: i64, kind: i64, name: &str, object: bool, trace: &str) {
    let actual = grouped_variant_report(operation);
    assert_eq!(actual[0], Value::Integer(number), "{operation}: {actual:?}");
    if number == 0 {
        let Value::Array(description) = &actual[1] else { panic!("{operation}: {actual:?}"); };
        assert_eq!(description.values, vec![Value::Integer(kind), Value::String(name.into()), Value::Boolean(object)], "{operation}");
    } else {
        assert_eq!(actual[1], Value::Empty, "{operation}");
    }
    assert_eq!(actual[2], Value::String(trace.into()), "{operation}");
}

#[test]
fn grouped_variant_arguments_preserve_the_formal_parameter_context() {
    check_grouped_variant_report(r#"answer=TakeValue(boxed)"#, 0, 3, "GroupedPayload", true, "G;");
    check_grouped_variant_report(r#"answer=TakeValue((boxed))"#, 0, 3, "GroupedPayload", true, "G;");
    check_grouped_variant_report(r#"answer=TakeRef((boxed))"#, 0, 3, "Long", false, "G;");
    check_grouped_variant_report(r#"answer=TakePack((boxed))"#, 0, 3, "Long", false, "G;");
    check_grouped_variant_report(r#"answer=late.TakeValue((boxed))"#, 0, 3, "GroupedPayload", true, "G;");
    check_grouped_variant_report(r#"answer=receiver.TakeValue((boxed))"#, 0, 3, "GroupedPayload", true, "G;");
    check_grouped_variant_report(r#"answer=CallByName(late,"TakeValue",1,(boxed))"#, 0, 3, "Long", false, "G;");
    check_grouped_variant_report(r#"answer=CallByName(receiver,"TakeValue",1,(boxed))"#, 0, 3, "Long", false, "G;");
    check_grouped_variant_report(r#"answer=TakePack((obj))"#, 0, 3, "Long", false, "G;");
    check_grouped_variant_report(r#"answer=CallByName(late,"TakeValue",1,(obj))"#, 0, 3, "Long", false, "G;");
    check_grouped_variant_report(r#"answer=TakeValue((obj))"#, 0, 3, "Long", false, "G;");
    check_grouped_variant_report(r#"held=Array((boxed)):answer=Describe(held(0))"#, 0, 3, "Long", false, "G;");
    check_grouped_variant_report(r#"held=IIf(True,(boxed),0):answer=Describe(held)"#, 0, 3, "Long", false, "G;");
    check_grouped_variant_report(r#"held=Choose(1,(boxed)):answer=Describe(held)"#, 0, 3, "Long", false, "G;");
    check_grouped_variant_report(r#"held=Switch(True,(boxed)):answer=Describe(held)"#, 0, 3, "Long", false, "G;");
    check_grouped_variant_report(r#"held=(boxed):answer=Describe(held)"#, 0, 3, "Long", false, "G;");
    check_grouped_variant_report(r#"answer=CallByName(late,"TakeValue",1,boxed)"#, 0, 3, "GroupedPayload", true, "G;");
    check_grouped_variant_report(r#"answer=TakePack(boxed)"#, 0, 3, "GroupedPayload", true, "G;");
    check_grouped_variant_report(r#"answer=TakeValue((typed))"#, 0, 3, "Long", false, "G;");
    check_grouped_variant_report(r#"answer=TakeValue(((boxed)))"#, 0, 3, "GroupedPayload", true, "G;");
    check_grouped_variant_report(r#"answer=TakePack(((boxed)))"#, 0, 3, "Long", false, "G;");
    check_grouped_variant_report(r#"answer=CallByName(late,"TakeValue",1,((boxed)))"#, 0, 3, "Long", false, "G;");
}

#[test]
fn grouped_variant_intrinsic_operands_preserve_object_values() {
    check_grouped_variant_report(r#"answer=TakeValue(IIf(True,(boxed),0))"#, 0, 3, "Long", false, "G;");
    check_grouped_variant_report(r#"answer=TakeValue(Choose(1,(boxed)))"#, 0, 3, "Long", false, "G;");
    check_grouped_variant_report(r#"answer=TakeValue(Switch(True,(boxed)))"#, 0, 3, "Long", false, "G;");
    check_grouped_variant_report(r#"answer=TakeValue(IIf(True,boxed,0))"#, 0, 3, "GroupedPayload", true, "G;");
    check_grouped_variant_report(r#"answer=TakeValue(Choose(1,boxed))"#, 0, 3, "GroupedPayload", true, "G;");
    check_grouped_variant_report(r#"answer=TakeValue(Switch(True,boxed))"#, 0, 3, "GroupedPayload", true, "G;");
}

#[test]
fn grouped_variant_arguments_preserve_callback_order() {
    check_grouped_variant_report(r#"answer=TakeRef((boxed),Later())"#, 0, 3, "Long", false, "L;G;");
    check_grouped_variant_report(r#"answer=TakeValue((boxed),Later())"#, 0, 3, "GroupedPayload", true, "L;G;");
    check_grouped_variant_report(r#"answer=TakePack((boxed),Later())"#, 0, 3, "Long", false, "G;L;");
    check_grouped_variant_report(r#"answer=TakeValue(IIf(True,(boxed),Later()))"#, 0, 3, "Long", false, "L;G;");
    check_grouped_variant_report(r#"answer=TakeValue(Choose(1,(boxed),Later()))"#, 0, 3, "Long", false, "G;L;");
    check_grouped_variant_report(r#"answer=TakeValue(Switch(True,(boxed),False,Later()))"#, 0, 3, "Long", false, "G;L;");
    check_grouped_variant_report(r#"answer=CallByName(late,"TakeValue",1,(boxed),Later())"#, 0, 3, "Long", false, "L;G;");
    check_grouped_variant_report(r#"answer=TakeRef((boxed),Change(boxed))"#, 91, 0, "", false, "L;");
    check_grouped_variant_report(r#"answer=TakeValue((boxed),Change(boxed))"#, 0, 9, "Nothing", true, "L;");
    check_grouped_variant_report(r#"answer=TakePack((boxed),Change(boxed))"#, 0, 3, "Long", false, "G;L;");
    check_grouped_variant_report(r#"answer=TakeValue(IIf(True,(boxed),Change(boxed)))"#, 91, 0, "", false, "L;");
    check_grouped_variant_report(r#"answer=TakeValue(Choose(1,(boxed),Change(boxed)))"#, 0, 3, "Long", false, "G;L;");
    check_grouped_variant_report(r#"answer=TakeValue(Switch(True,(boxed),False,Change(boxed)))"#, 0, 3, "Long", false, "G;L;");
    check_grouped_variant_report(r#"answer=CallByName(late,"TakeValue",1,(boxed),Change(boxed))"#, 91, 0, "", false, "L;");
}

#[test]
fn grouped_variant_arguments_preserve_errors_and_switch_selection() {
    check_grouped_variant_report(r#"answer=TakeValue(Switch(boxed,11,True,22))"#, 0, 2, "Integer", false, "G;");
    check_grouped_variant_report(r#"answer=TakeValue(Switch(True,boxed,boxed,0))"#, 0, 3, "GroupedPayload", true, "G;");
    check_grouped_variant_report(r#"answer=TakeValue(Switch(False,0,boxed,11))"#, 0, 2, "Integer", false, "G;");
    check_grouped_variant_report(r#"answer=TakeValue(Switch(True,boxed,Nothing,0))"#, 0, 3, "GroupedPayload", true, "G;");
    check_grouped_variant_report(r#"answer=TakeValue(Switch(False,0,Nothing,1))"#, 91, 0, "", false, "");
    check_grouped_variant_report(r#"answer=TakeValue(Switch(Null,0,True,boxed))"#, 0, 3, "GroupedPayload", true, "G;");
    check_grouped_variant_report(r#"answer=TakeValue(Switch(CVErr(5),0,True,boxed))"#, 13, 0, "", false, "");
    check_grouped_variant_report(r#"answer=TakeValue(Switch(True,boxed,CVErr(5),0))"#, 0, 3, "GroupedPayload", true, "G;");
    check_grouped_variant_report(r#"answer=TakeValue(Switch(0,0,7,boxed))"#, 0, 3, "GroupedPayload", true, "G;");
    check_grouped_variant_report(r#"answer=CallByName(late,"TakeValue",1,1/zero,Later())"#, 11, 0, "", false, "L;");
    check_grouped_variant_report(r#"answer=CallByName(late,"TakeValue",1,(boxed),1/zero)"#, 11, 0, "", false, "G;");
    check_grouped_variant_report(r#"answer=CallByName(late,"TakeValue",1,1/zero,(boxed))"#, 11, 0, "", false, "");
    check_grouped_variant_report(r#"answer=CallByName(late,"TakeValue",1,(obj),Later())"#, 0, 3, "Long", false, "G;L;");
    check_grouped_variant_report(r#"answer=CallByName(late,"TakeValue",1,boxed,Change(boxed))"#, 0, 9, "Nothing", true, "L;");
    check_grouped_variant_report(r#"answer=CallByName(late,"TakeValue",1,((boxed)),Later())"#, 0, 3, "Long", false, "L;G;");
    check_grouped_variant_report(r#"answer=TakeValue(IIf(True,(boxed),1/zero))"#, 11, 0, "", false, "");
}

fn typed_array_writeback_report(element: &str, operation: &str, invocation: &str) -> Vec<Value> {
    let pack_operation = operation.replace("value", "values(0)");
    array_report(&format!(r#"
Private Function FailValue() As Variant
 Err.Raise 5
End Function
Private Function MutateRef(ByRef value As Variant) As Long
 On Error Resume Next
 Err.Clear
 {operation}
 Dim number As Long:number=Err.Number
 Err.Clear
 MutateRef=number
End Function
Private Function MutateValue(ByVal value As Variant) As Long
 On Error Resume Next
 Err.Clear
 {operation}
 Dim number As Long:number=Err.Number
 Err.Clear
 MutateValue=number
End Function
Private Function MutatePack(ParamArray values() As Variant) As Long
 On Error Resume Next
 Err.Clear
 {pack_operation}
 Dim number As Long:number=Err.Number
 Err.Clear
 MutatePack=number
End Function
Public Function Probe() As Variant
 Dim data() As {element},inside As Long,outside As Long,lo As Long,hi As Long,le As Long,he As Long
 ReDim data(1 To 1)
 On Error Resume Next
 Err.Clear
 inside={invocation}
 outside=Err.Number
 Err.Clear
 lo=LBound(data):le=Err.Number
 Err.Clear
 hi=UBound(data):he=Err.Number
 Probe=Array(inside,outside,lo,le,hi,he,VarType(data))
End Function
"#))
}

#[test]
fn typed_array_variant_references_reject_whole_value_replacement() {
    for (element, kind) in [("Object", 8201), ("Long", 8195), ("Variant", 8204), ("Byte", 8209), ("String", 8200)] {
        for invocation in ["MutateRef(data)", "MutatePack(data)"] {
            for (operation, number) in [("value=Empty",458),("value=Null",458),("value=CVErr(2000)",458),("Set value=Nothing",13),("value=FailValue()",5)] {
                assert_eq!(typed_array_writeback_report(element,operation,invocation),
                    [number,0,1,0,1,0,kind].into_iter().map(Value::Integer).collect::<Vec<_>>(),
                    "{element}: {invocation}: {operation}");
            }
        }
    }
    for operation in ["value=Array(7,8)", "value=42"] {
        assert_eq!(typed_array_writeback_report("Object",operation,"MutateRef(data)"),
            [458,0,1,0,1,0,8201].into_iter().map(Value::Integer).collect::<Vec<_>>());
    }
}

#[test]
fn typed_array_variant_copies_remain_independently_replaceable() {
    for invocation in ["MutateRef((data))", "MutateValue(data)", "MutateValue((data))", "MutatePack((data))"] {
        for operation in ["value=Array(7,8)","value=42"] {
            assert_eq!(typed_array_writeback_report("Object",operation,invocation),
                [0,0,1,0,1,0,8201].into_iter().map(Value::Integer).collect::<Vec<_>>(), "{invocation}: {operation}");
        }
    }
    assert_eq!(typed_array_writeback_report("Object","ReDim value(2 To 3)","MutateRef(data)"),
        [0,0,2,0,3,0,8201].into_iter().map(Value::Integer).collect::<Vec<_>>());
    assert_eq!(typed_array_writeback_report("Object","Erase value","MutateRef(data)"),
        [0,0,0,9,0,9,8201].into_iter().map(Value::Integer).collect::<Vec<_>>());
}
