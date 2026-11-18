// This Source Code Form is subject to the terms of the Mozilla Public
// License, v. 2.0. If a copy of the MPL was not distributed with this
// file, You can obtain one at https://mozilla.org/MPL/2.0/.

use oxivba_core::{ErrorEnvironment, ErrorLanguage};
use oxivba_core::parser::parse_project;
use oxivba_core::runtime::{Runtime, Value};

// Compare every captured returned field, including explicit VarType witnesses
// and the complete five-field Err record. Numeric observations retain all bits.
fn snapshot(value: &Value) -> String {
    let number = match value {
        Value::Byte(n) => Some(f64::from(*n)),
        Value::Int16(n) => Some(f64::from(*n)),
        Value::Integer(n) | Value::LongLong(n) => Some(*n as f64),
        Value::Single(n) => Some(f64::from(*n)),
        Value::Double(n) => Some(*n),
        _ => None,
    };
    if let Some(n) = number { return format!("N{:016x}", n.to_bits()); }
    match value {
        Value::Empty | Value::Null => "E".into(),
        Value::Boolean(true) => "T".into(),
        Value::Boolean(false) => "F".into(),
        Value::String(s) => format!("S{}:{s}", s.len()),
        Value::NullString => "S0:".into(),
        Value::Array(array) => format!("A{}[{}]", array.values.len(),
            array.values.iter().map(snapshot).collect::<Vec<_>>().join(";")),
        other => panic!("unexpected recorded value kind: {other:?}"),
    }
}

fn compare(source: &str, rows: &str, help_file: &str) {
    let (module, classes) = parse_project(source).unwrap();
    let mut differences = Vec::new();
    for row in rows.lines() {
        let mut fields = row.splitn(3, '\t');
        let kind = fields.next().unwrap().parse::<i64>().unwrap();
        let pathname = fields.next().unwrap();
        let expected = fields.next().unwrap();
        let actual = Runtime::new(&module)
            .with_classes(&classes)
            .with_project_name("VBAProject")
            .with_error_environment(ErrorEnvironment {
                language: ErrorLanguage::Japanese,
                help_file: help_file.into(),
            })
            .call("Probe", vec![Value::Integer(kind), Value::String(pathname.into())]);
        match actual {
            Ok(value) if snapshot(&value) == expected => {},
            other => differences.push(format!("case={kind}, actual={other:?}, expected={expected}")),
        }
    }
    assert!(differences.is_empty(), "{} Native differences: {:?}", differences.len(),
        &differences[..differences.len().min(4)]);
}

// Native Excel source SHA-256 aa0195c9dfd3f8922013922e616f76d44f4bb74bd04b6ef6b1828aab94b693fb.
#[test]
fn native_len_typed_assignment() {
 compare(r####"Attribute VB_Name = "LenCases4953"
Option Explicit
Private target As String, trace As String
Public Sub Mark(ByVal token As String)
 trace=trace & token & "|"
End Sub
Public Function Probe(ByVal kind As Long,ByVal pathname As String) As Variant
 Dim item As Object, answer As Long, savedErr As Variant
 target=pathname:trace=""
 savedErr=Array(0,"","","",0)
 On Error GoTo Caught
 Err.Clear
 Mark "before"
 Select Case kind
 Case 0: Set item=New Leaf4953:answer=Len(item)
 Case 1: Set item=New Parent4953:answer=CStr(item)
 Case 2: Set item=New Parent4953:answer=CLng(item)
 Case 3: Set item=New Parent4953:answer=Len(item)
 Case 4: Set item=New VariantParent4953:answer=Len(item)
 Case 5: Set item=New TypedParent4953:answer=Len(item)
 Case 6: Set item=New PlainParent4953:answer=Len(item)
 Case 7: Set item=New Parent4953:answer=LenB(item)
 End Select
 Mark "after"
 GoTo Done
Caught:
 savedErr=Array(Err.Number,Err.Description,Err.Source,Err.HelpFile,Err.HelpContext)
 Mark "caught:" & CStr(savedErr(0))
Done:
 Probe=Array(Array(kind,trace,answer,TypeName(answer),VarType(answer)),savedErr)
End Function
VERSION 1.0 CLASS
Attribute VB_Name = "Leaf4953"
Option Explicit
Public Property Get Scalar() As Long
Attribute Scalar.VB_UserMemId = 0
 LenCases4953.Mark "leaf"
 Scalar=83
End Property
VERSION 1.0 CLASS
Attribute VB_Name = "Parent4953"
Option Explicit
Public Property Get Scalar() As Object
Attribute Scalar.VB_UserMemId = 0
 LenCases4953.Mark "parent"
 Set Scalar=New Leaf4953
End Property
VERSION 1.0 CLASS
Attribute VB_Name = "VariantParent4953"
Option Explicit
Public Property Get Scalar() As Variant
Attribute Scalar.VB_UserMemId = 0
 LenCases4953.Mark "variant_parent"
 Set Scalar=New Leaf4953
End Property
VERSION 1.0 CLASS
Attribute VB_Name = "TypedParent4953"
Option Explicit
Public Property Get Scalar() As Leaf4953
Attribute Scalar.VB_UserMemId = 0
 LenCases4953.Mark "typed_parent"
 Set Scalar=New Leaf4953
End Property
VERSION 1.0 CLASS
Attribute VB_Name = "Plain4953"
Option Explicit
Public Value As Long
VERSION 1.0 CLASS
Attribute VB_Name = "PlainParent4953"
Option Explicit
Public Property Get Scalar() As Object
Attribute Scalar.VB_UserMemId = 0
 LenCases4953.Mark "plain_parent"
 Set Scalar=New Plain4953
 Scalar.Value=17
End Property
"####, r####"0	E:\oxi-vba-public-artifacts-20261115\types4959\owned-data-4960-0-0.bin	A2[A5[N0000000000000000;S18:before|leaf|after|;N4000000000000000;S4:Long;N4008000000000000];A5[N0000000000000000;S0:;S0:;S0:;N0000000000000000]]
1	E:\oxi-vba-public-artifacts-20261115\types4959\owned-data-4960-0-1.bin	A2[A5[N3ff0000000000000;S25:before|parent|leaf|after|;N4054c00000000000;S4:Long;N4008000000000000];A5[N0000000000000000;S0:;S0:;S0:;N0000000000000000]]
2	E:\oxi-vba-public-artifacts-20261115\types4959\owned-data-4960-0-2.bin	A2[A5[N4000000000000000;S25:before|parent|leaf|after|;N4054c00000000000;S4:Long;N4008000000000000];A5[N0000000000000000;S0:;S0:;S0:;N0000000000000000]]
3	E:\oxi-vba-public-artifacts-20261115\types4959\owned-data-4960-0-3.bin	A2[A5[N4008000000000000;S30:before|parent|leaf|caught:458|;N0000000000000000;S4:Long;N4008000000000000];A5[N407ca00000000000;S109:Visual Basic でサポートされていないオートメーションが変数で使用されています。;S10:VBAProject;S72:C:\Program Files\Common Files\Microsoft Shared\VBA\VBA7.1\1041\VbLR6.chm;N412e881400000000]]
4	E:\oxi-vba-public-artifacts-20261115\types4959\owned-data-4960-0-4.bin	A2[A5[N4010000000000000;S38:before|variant_parent|leaf|caught:458|;N0000000000000000;S4:Long;N4008000000000000];A5[N407ca00000000000;S109:Visual Basic でサポートされていないオートメーションが変数で使用されています。;S10:VBAProject;S72:C:\Program Files\Common Files\Microsoft Shared\VBA\VBA7.1\1041\VbLR6.chm;N412e881400000000]]
5	E:\oxi-vba-public-artifacts-20261115\types4959\owned-data-4960-0-5.bin	A2[A5[N4014000000000000;S36:before|typed_parent|leaf|caught:458|;N0000000000000000;S4:Long;N4008000000000000];A5[N407ca00000000000;S109:Visual Basic でサポートされていないオートメーションが変数で使用されています。;S10:VBAProject;S72:C:\Program Files\Common Files\Microsoft Shared\VBA\VBA7.1\1041\VbLR6.chm;N412e881400000000]]
6	E:\oxi-vba-public-artifacts-20261115\types4959\owned-data-4960-0-6.bin	A2[A5[N4018000000000000;S31:before|plain_parent|caught:438|;N0000000000000000;S4:Long;N4008000000000000];A5[N407b600000000000;S102:オブジェクトは、このプロパティまたはメソッドをサポートしていません。;S10:VBAProject;S72:C:\Program Files\Common Files\Microsoft Shared\VBA\VBA7.1\1041\VbLR6.chm;N412e87ec00000000]]
7	E:\oxi-vba-public-artifacts-20261115\types4959\owned-data-4960-0-7.bin	A2[A5[N401c000000000000;S30:before|parent|leaf|caught:458|;N0000000000000000;S4:Long;N4008000000000000];A5[N407ca00000000000;S109:Visual Basic でサポートされていないオートメーションが変数で使用されています。;S10:VBAProject;S72:C:\Program Files\Common Files\Microsoft Shared\VBA\VBA7.1\1041\VbLR6.chm;N412e881400000000]]"####, r####"C:\Program Files\Common Files\Microsoft Shared\VBA\VBA7.1\1041\VbLR6.chm"####);
}
