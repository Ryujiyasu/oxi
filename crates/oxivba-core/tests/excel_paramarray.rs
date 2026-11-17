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
