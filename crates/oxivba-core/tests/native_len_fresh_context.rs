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

// Native Excel source SHA-256 2e45872845f4f2be0b8e2060f34596ef5ee7bb1c405b431d30acea00f05ba4a9.
#[test]
fn native_len_fresh_context() {
 compare(r####"Attribute VB_Name = "Consume5052"
Option Explicit
Private sink As Variant, slots(0 To 0) As Variant
Private target As String, trace As String
Private leafResult As Long
Public Function GetLeafResult() As Long
 GetLeafResult=leafResult
End Function
Public Sub Mark(ByVal token As String)
 trace=trace & token & "|"
End Sub
Public Function ReadByRef(ByRef arg As Variant) As Long
 Mark "byref_enter"
 ReadByRef=VarType(arg)
End Function
Public Function ReadByVal(ByVal arg As Variant) As Long
 Mark "byval_enter"
 ReadByVal=VarType(arg)
End Function
Public Sub ReadByRefSub(ByRef arg As Variant)
 Mark "byref_sub_enter"
End Sub
Public Function ReturnHeld(ByRef arg As Variant) As Variant
 Mark "return_enter"
 ReturnHeld=arg
End Function
Public Function FreshLen(ByVal item As Object,ByVal useBytes As Boolean) As Variant
 Mark "fresh_enter"
 If useBytes Then
  FreshLen=LenB(item)
 Else
  FreshLen=Len(item)
 End If
 Mark "fresh_after"
End Function
Public Function Probe(ByVal kind As Long,ByVal pathname As String) As Variant
 Dim item As Object, value As Variant, copied As Variant, observed As Variant
 Dim number As Long, seed As Long, operation As Long, seedErr As Variant, savedErr As Variant, tags As Variant
 target=pathname:trace="":tags=Array(16,83):leafResult=tags(kind\28):seed=(kind Mod 28)\14:operation=kind Mod 14
 On Error Resume Next
 Err.Clear
 Set item=New Parent5052
 Mark "seed_before"
 If seed=0 Then
  value=Len(item)
 Else
  value=LenB(item)
 End If
 seedErr=Array(Err.Number,Err.Description,Err.Source,Err.HelpFile,Err.HelpContext)
 Mark "seed_after"
 Err.Clear
 Select Case operation
 Case 0: observed=VarType((value))
 Case 1: observed=IsObject((value))
 Case 2: observed=IsEmpty((value))
 Case 3: observed=IsNumeric((value))
 Case 4: observed=IsNull((value))
 Case 5: observed=IsDate((value))
 Case 6: observed=IsError((value))
 Case 7: observed=IsArray((value))
 Case 8: observed=VarType((FreshLen(item,seed=1)))
 Case 9: observed=ReadByRef(FreshLen(item,seed=1))
 Case 10: observed=ReadByRef((FreshLen(item,seed=1)))
 Case 11: observed=ReadByVal(FreshLen(item,seed=1))
 Case 12: observed=ReadByVal((FreshLen(item,seed=1)))
 Case 13: copied=(FreshLen(item,seed=1)):observed=VarType(copied)
 End Select
 savedErr=Array(Err.Number,Err.Description,Err.Source,Err.HelpFile,Err.HelpContext)
 Mark "operation_after"
 Probe=Array(Array(kind,seed,operation,trace,observed,TypeName(observed),VarType(observed),seedErr,leafResult),savedErr)
End Function
VERSION 1.0 CLASS
Attribute VB_Name = "Leaf5052"
Option Explicit
Public Property Get Scalar() As Long
Attribute Scalar.VB_UserMemId = 0
 Consume5052.Mark "leaf"
 Scalar=Consume5052.GetLeafResult()
End Property
VERSION 1.0 CLASS
Attribute VB_Name = "Parent5052"
Option Explicit
Public Property Get Scalar() As Object
Attribute Scalar.VB_UserMemId = 0
 Consume5052.Mark "parent"
 Set Scalar=New Leaf5052
End Property
"####, r####"0	E:\oxi-vba-public-artifacts-20261115\types5052\owned-data-5053-0-0.bin	A2[A9[N0000000000000000;N0000000000000000;N0000000000000000;S51:seed_before|parent|leaf|seed_after|operation_after|;N4030000000000000;S4:Long;N4008000000000000;A5[N0000000000000000;S0:;S0:;S0:;N0000000000000000];N4030000000000000];A5[N0000000000000000;S0:;S0:;S0:;N0000000000000000]]
1	E:\oxi-vba-public-artifacts-20261115\types5052\owned-data-5053-0-1.bin	A2[A9[N3ff0000000000000;N0000000000000000;N3ff0000000000000;S51:seed_before|parent|leaf|seed_after|operation_after|;F;S7:Boolean;N4026000000000000;A5[N0000000000000000;S0:;S0:;S0:;N0000000000000000];N4030000000000000];A5[N0000000000000000;S0:;S0:;S0:;N0000000000000000]]
2	E:\oxi-vba-public-artifacts-20261115\types5052\owned-data-5053-0-2.bin	A2[A9[N4000000000000000;N0000000000000000;N4000000000000000;S51:seed_before|parent|leaf|seed_after|operation_after|;F;S7:Boolean;N4026000000000000;A5[N0000000000000000;S0:;S0:;S0:;N0000000000000000];N4030000000000000];A5[N0000000000000000;S0:;S0:;S0:;N0000000000000000]]
3	E:\oxi-vba-public-artifacts-20261115\types5052\owned-data-5053-0-3.bin	A2[A9[N4008000000000000;N0000000000000000;N4008000000000000;S51:seed_before|parent|leaf|seed_after|operation_after|;F;S7:Boolean;N4026000000000000;A5[N0000000000000000;S0:;S0:;S0:;N0000000000000000];N4030000000000000];A5[N0000000000000000;S0:;S0:;S0:;N0000000000000000]]
4	E:\oxi-vba-public-artifacts-20261115\types5052\owned-data-5053-0-4.bin	A2[A9[N4010000000000000;N0000000000000000;N4010000000000000;S51:seed_before|parent|leaf|seed_after|operation_after|;F;S7:Boolean;N4026000000000000;A5[N0000000000000000;S0:;S0:;S0:;N0000000000000000];N4030000000000000];A5[N0000000000000000;S0:;S0:;S0:;N0000000000000000]]
5	E:\oxi-vba-public-artifacts-20261115\types5052\owned-data-5053-0-5.bin	A2[A9[N4014000000000000;N0000000000000000;N4014000000000000;S51:seed_before|parent|leaf|seed_after|operation_after|;F;S7:Boolean;N4026000000000000;A5[N0000000000000000;S0:;S0:;S0:;N0000000000000000];N4030000000000000];A5[N0000000000000000;S0:;S0:;S0:;N0000000000000000]]
6	E:\oxi-vba-public-artifacts-20261115\types5052\owned-data-5053-0-6.bin	A2[A9[N4018000000000000;N0000000000000000;N4018000000000000;S51:seed_before|parent|leaf|seed_after|operation_after|;F;S7:Boolean;N4026000000000000;A5[N0000000000000000;S0:;S0:;S0:;N0000000000000000];N4030000000000000];A5[N0000000000000000;S0:;S0:;S0:;N0000000000000000]]
7	E:\oxi-vba-public-artifacts-20261115\types5052\owned-data-5053-0-7.bin	A2[A9[N401c000000000000;N0000000000000000;N401c000000000000;S51:seed_before|parent|leaf|seed_after|operation_after|;F;S7:Boolean;N4026000000000000;A5[N0000000000000000;S0:;S0:;S0:;N0000000000000000];N4030000000000000];A5[N0000000000000000;S0:;S0:;S0:;N0000000000000000]]
8	E:\oxi-vba-public-artifacts-20261115\types5052\owned-data-5053-0-8.bin	A2[A9[N4020000000000000;N0000000000000000;N4020000000000000;S87:seed_before|parent|leaf|seed_after|fresh_enter|parent|leaf|fresh_after|operation_after|;N4030000000000000;S4:Long;N4008000000000000;A5[N0000000000000000;S0:;S0:;S0:;N0000000000000000];N4030000000000000];A5[N0000000000000000;S0:;S0:;S0:;N0000000000000000]]
9	E:\oxi-vba-public-artifacts-20261115\types5052\owned-data-5053-0-9.bin	A2[A9[N4022000000000000;N0000000000000000;N4022000000000000;S99:seed_before|parent|leaf|seed_after|fresh_enter|parent|leaf|fresh_after|byref_enter|operation_after|;N4030000000000000;S4:Long;N4008000000000000;A5[N0000000000000000;S0:;S0:;S0:;N0000000000000000];N4030000000000000];A5[N0000000000000000;S0:;S0:;S0:;N0000000000000000]]
10	E:\oxi-vba-public-artifacts-20261115\types5052\owned-data-5053-0-10.bin	A2[A9[N4024000000000000;N0000000000000000;N4024000000000000;S99:seed_before|parent|leaf|seed_after|fresh_enter|parent|leaf|fresh_after|byref_enter|operation_after|;N4030000000000000;S4:Long;N4008000000000000;A5[N0000000000000000;S0:;S0:;S0:;N0000000000000000];N4030000000000000];A5[N0000000000000000;S0:;S0:;S0:;N0000000000000000]]
11	E:\oxi-vba-public-artifacts-20261115\types5052\owned-data-5053-0-11.bin	A2[A9[N4026000000000000;N0000000000000000;N4026000000000000;S99:seed_before|parent|leaf|seed_after|fresh_enter|parent|leaf|fresh_after|byval_enter|operation_after|;N4030000000000000;S4:Long;N4008000000000000;A5[N0000000000000000;S0:;S0:;S0:;N0000000000000000];N4030000000000000];A5[N0000000000000000;S0:;S0:;S0:;N0000000000000000]]
12	E:\oxi-vba-public-artifacts-20261115\types5052\owned-data-5053-0-12.bin	A2[A9[N4028000000000000;N0000000000000000;N4028000000000000;S99:seed_before|parent|leaf|seed_after|fresh_enter|parent|leaf|fresh_after|byval_enter|operation_after|;N4030000000000000;S4:Long;N4008000000000000;A5[N0000000000000000;S0:;S0:;S0:;N0000000000000000];N4030000000000000];A5[N0000000000000000;S0:;S0:;S0:;N0000000000000000]]
13	E:\oxi-vba-public-artifacts-20261115\types5052\owned-data-5053-0-13.bin	A2[A9[N402a000000000000;N0000000000000000;N402a000000000000;S87:seed_before|parent|leaf|seed_after|fresh_enter|parent|leaf|fresh_after|operation_after|;N4030000000000000;S4:Long;N4008000000000000;A5[N0000000000000000;S0:;S0:;S0:;N0000000000000000];N4030000000000000];A5[N0000000000000000;S0:;S0:;S0:;N0000000000000000]]
14	E:\oxi-vba-public-artifacts-20261115\types5052\owned-data-5053-0-14.bin	A2[A9[N402c000000000000;N3ff0000000000000;N0000000000000000;S51:seed_before|parent|leaf|seed_after|operation_after|;N4030000000000000;S4:Long;N4008000000000000;A5[N0000000000000000;S0:;S0:;S0:;N0000000000000000];N4030000000000000];A5[N0000000000000000;S0:;S0:;S0:;N0000000000000000]]
15	E:\oxi-vba-public-artifacts-20261115\types5052\owned-data-5053-0-15.bin	A2[A9[N402e000000000000;N3ff0000000000000;N3ff0000000000000;S51:seed_before|parent|leaf|seed_after|operation_after|;F;S7:Boolean;N4026000000000000;A5[N0000000000000000;S0:;S0:;S0:;N0000000000000000];N4030000000000000];A5[N0000000000000000;S0:;S0:;S0:;N0000000000000000]]
16	E:\oxi-vba-public-artifacts-20261115\types5052\owned-data-5053-0-16.bin	A2[A9[N4030000000000000;N3ff0000000000000;N4000000000000000;S51:seed_before|parent|leaf|seed_after|operation_after|;F;S7:Boolean;N4026000000000000;A5[N0000000000000000;S0:;S0:;S0:;N0000000000000000];N4030000000000000];A5[N0000000000000000;S0:;S0:;S0:;N0000000000000000]]
17	E:\oxi-vba-public-artifacts-20261115\types5052\owned-data-5053-0-17.bin	A2[A9[N4031000000000000;N3ff0000000000000;N4008000000000000;S51:seed_before|parent|leaf|seed_after|operation_after|;F;S7:Boolean;N4026000000000000;A5[N0000000000000000;S0:;S0:;S0:;N0000000000000000];N4030000000000000];A5[N0000000000000000;S0:;S0:;S0:;N0000000000000000]]
18	E:\oxi-vba-public-artifacts-20261115\types5052\owned-data-5053-0-18.bin	A2[A9[N4032000000000000;N3ff0000000000000;N4010000000000000;S51:seed_before|parent|leaf|seed_after|operation_after|;F;S7:Boolean;N4026000000000000;A5[N0000000000000000;S0:;S0:;S0:;N0000000000000000];N4030000000000000];A5[N0000000000000000;S0:;S0:;S0:;N0000000000000000]]
19	E:\oxi-vba-public-artifacts-20261115\types5052\owned-data-5053-0-19.bin	A2[A9[N4033000000000000;N3ff0000000000000;N4014000000000000;S51:seed_before|parent|leaf|seed_after|operation_after|;F;S7:Boolean;N4026000000000000;A5[N0000000000000000;S0:;S0:;S0:;N0000000000000000];N4030000000000000];A5[N0000000000000000;S0:;S0:;S0:;N0000000000000000]]
20	E:\oxi-vba-public-artifacts-20261115\types5052\owned-data-5053-0-20.bin	A2[A9[N4034000000000000;N3ff0000000000000;N4018000000000000;S51:seed_before|parent|leaf|seed_after|operation_after|;F;S7:Boolean;N4026000000000000;A5[N0000000000000000;S0:;S0:;S0:;N0000000000000000];N4030000000000000];A5[N0000000000000000;S0:;S0:;S0:;N0000000000000000]]
21	E:\oxi-vba-public-artifacts-20261115\types5052\owned-data-5053-0-21.bin	A2[A9[N4035000000000000;N3ff0000000000000;N401c000000000000;S51:seed_before|parent|leaf|seed_after|operation_after|;F;S7:Boolean;N4026000000000000;A5[N0000000000000000;S0:;S0:;S0:;N0000000000000000];N4030000000000000];A5[N0000000000000000;S0:;S0:;S0:;N0000000000000000]]
22	E:\oxi-vba-public-artifacts-20261115\types5052\owned-data-5053-0-22.bin	A2[A9[N4036000000000000;N3ff0000000000000;N4020000000000000;S87:seed_before|parent|leaf|seed_after|fresh_enter|parent|leaf|fresh_after|operation_after|;N4030000000000000;S4:Long;N4008000000000000;A5[N0000000000000000;S0:;S0:;S0:;N0000000000000000];N4030000000000000];A5[N0000000000000000;S0:;S0:;S0:;N0000000000000000]]
23	E:\oxi-vba-public-artifacts-20261115\types5052\owned-data-5053-0-23.bin	A2[A9[N4037000000000000;N3ff0000000000000;N4022000000000000;S99:seed_before|parent|leaf|seed_after|fresh_enter|parent|leaf|fresh_after|byref_enter|operation_after|;N4030000000000000;S4:Long;N4008000000000000;A5[N0000000000000000;S0:;S0:;S0:;N0000000000000000];N4030000000000000];A5[N0000000000000000;S0:;S0:;S0:;N0000000000000000]]
24	E:\oxi-vba-public-artifacts-20261115\types5052\owned-data-5053-0-24.bin	A2[A9[N4038000000000000;N3ff0000000000000;N4024000000000000;S99:seed_before|parent|leaf|seed_after|fresh_enter|parent|leaf|fresh_after|byref_enter|operation_after|;N4030000000000000;S4:Long;N4008000000000000;A5[N0000000000000000;S0:;S0:;S0:;N0000000000000000];N4030000000000000];A5[N0000000000000000;S0:;S0:;S0:;N0000000000000000]]
25	E:\oxi-vba-public-artifacts-20261115\types5052\owned-data-5053-0-25.bin	A2[A9[N4039000000000000;N3ff0000000000000;N4026000000000000;S99:seed_before|parent|leaf|seed_after|fresh_enter|parent|leaf|fresh_after|byval_enter|operation_after|;N4030000000000000;S4:Long;N4008000000000000;A5[N0000000000000000;S0:;S0:;S0:;N0000000000000000];N4030000000000000];A5[N0000000000000000;S0:;S0:;S0:;N0000000000000000]]
26	E:\oxi-vba-public-artifacts-20261115\types5052\owned-data-5053-0-26.bin	A2[A9[N403a000000000000;N3ff0000000000000;N4028000000000000;S99:seed_before|parent|leaf|seed_after|fresh_enter|parent|leaf|fresh_after|byval_enter|operation_after|;N4030000000000000;S4:Long;N4008000000000000;A5[N0000000000000000;S0:;S0:;S0:;N0000000000000000];N4030000000000000];A5[N0000000000000000;S0:;S0:;S0:;N0000000000000000]]
27	E:\oxi-vba-public-artifacts-20261115\types5052\owned-data-5053-0-27.bin	A2[A9[N403b000000000000;N3ff0000000000000;N402a000000000000;S87:seed_before|parent|leaf|seed_after|fresh_enter|parent|leaf|fresh_after|operation_after|;N4030000000000000;S4:Long;N4008000000000000;A5[N0000000000000000;S0:;S0:;S0:;N0000000000000000];N4030000000000000];A5[N0000000000000000;S0:;S0:;S0:;N0000000000000000]]
28	E:\oxi-vba-public-artifacts-20261115\types5052\owned-data-5053-0-28.bin	A2[A9[N403c000000000000;N0000000000000000;N0000000000000000;S51:seed_before|parent|leaf|seed_after|operation_after|;E;S5:Empty;N0000000000000000;A5[N0000000000000000;S0:;S0:;S0:;N0000000000000000];N4054c00000000000];A5[N407ca00000000000;S109:Visual Basic でサポートされていないオートメーションが変数で使用されています。;S10:VBAProject;S72:C:\Program Files\Common Files\Microsoft Shared\VBA\VBA7.1\1041\VbLR6.chm;N412e881400000000]]
29	E:\oxi-vba-public-artifacts-20261115\types5052\owned-data-5053-0-29.bin	A2[A9[N403d000000000000;N0000000000000000;N3ff0000000000000;S51:seed_before|parent|leaf|seed_after|operation_after|;E;S5:Empty;N0000000000000000;A5[N0000000000000000;S0:;S0:;S0:;N0000000000000000];N4054c00000000000];A5[N407ca00000000000;S109:Visual Basic でサポートされていないオートメーションが変数で使用されています。;S10:VBAProject;S72:C:\Program Files\Common Files\Microsoft Shared\VBA\VBA7.1\1041\VbLR6.chm;N412e881400000000]]
30	E:\oxi-vba-public-artifacts-20261115\types5052\owned-data-5053-0-30.bin	A2[A9[N403e000000000000;N0000000000000000;N4000000000000000;S51:seed_before|parent|leaf|seed_after|operation_after|;E;S5:Empty;N0000000000000000;A5[N0000000000000000;S0:;S0:;S0:;N0000000000000000];N4054c00000000000];A5[N407ca00000000000;S109:Visual Basic でサポートされていないオートメーションが変数で使用されています。;S10:VBAProject;S72:C:\Program Files\Common Files\Microsoft Shared\VBA\VBA7.1\1041\VbLR6.chm;N412e881400000000]]
31	E:\oxi-vba-public-artifacts-20261115\types5052\owned-data-5053-0-31.bin	A2[A9[N403f000000000000;N0000000000000000;N4008000000000000;S51:seed_before|parent|leaf|seed_after|operation_after|;E;S5:Empty;N0000000000000000;A5[N0000000000000000;S0:;S0:;S0:;N0000000000000000];N4054c00000000000];A5[N407ca00000000000;S109:Visual Basic でサポートされていないオートメーションが変数で使用されています。;S10:VBAProject;S72:C:\Program Files\Common Files\Microsoft Shared\VBA\VBA7.1\1041\VbLR6.chm;N412e881400000000]]
32	E:\oxi-vba-public-artifacts-20261115\types5052\owned-data-5053-0-32.bin	A2[A9[N4040000000000000;N0000000000000000;N4010000000000000;S51:seed_before|parent|leaf|seed_after|operation_after|;E;S5:Empty;N0000000000000000;A5[N0000000000000000;S0:;S0:;S0:;N0000000000000000];N4054c00000000000];A5[N407ca00000000000;S109:Visual Basic でサポートされていないオートメーションが変数で使用されています。;S10:VBAProject;S72:C:\Program Files\Common Files\Microsoft Shared\VBA\VBA7.1\1041\VbLR6.chm;N412e881400000000]]
33	E:\oxi-vba-public-artifacts-20261115\types5052\owned-data-5053-0-33.bin	A2[A9[N4040800000000000;N0000000000000000;N4014000000000000;S51:seed_before|parent|leaf|seed_after|operation_after|;E;S5:Empty;N0000000000000000;A5[N0000000000000000;S0:;S0:;S0:;N0000000000000000];N4054c00000000000];A5[N407ca00000000000;S109:Visual Basic でサポートされていないオートメーションが変数で使用されています。;S10:VBAProject;S72:C:\Program Files\Common Files\Microsoft Shared\VBA\VBA7.1\1041\VbLR6.chm;N412e881400000000]]
34	E:\oxi-vba-public-artifacts-20261115\types5052\owned-data-5053-0-34.bin	A2[A9[N4041000000000000;N0000000000000000;N4018000000000000;S51:seed_before|parent|leaf|seed_after|operation_after|;E;S5:Empty;N0000000000000000;A5[N0000000000000000;S0:;S0:;S0:;N0000000000000000];N4054c00000000000];A5[N407ca00000000000;S109:Visual Basic でサポートされていないオートメーションが変数で使用されています。;S10:VBAProject;S72:C:\Program Files\Common Files\Microsoft Shared\VBA\VBA7.1\1041\VbLR6.chm;N412e881400000000]]
35	E:\oxi-vba-public-artifacts-20261115\types5052\owned-data-5053-0-35.bin	A2[A9[N4041800000000000;N0000000000000000;N401c000000000000;S51:seed_before|parent|leaf|seed_after|operation_after|;E;S5:Empty;N0000000000000000;A5[N0000000000000000;S0:;S0:;S0:;N0000000000000000];N4054c00000000000];A5[N407ca00000000000;S109:Visual Basic でサポートされていないオートメーションが変数で使用されています。;S10:VBAProject;S72:C:\Program Files\Common Files\Microsoft Shared\VBA\VBA7.1\1041\VbLR6.chm;N412e881400000000]]
36	E:\oxi-vba-public-artifacts-20261115\types5052\owned-data-5053-0-36.bin	A2[A9[N4042000000000000;N0000000000000000;N4020000000000000;S87:seed_before|parent|leaf|seed_after|fresh_enter|parent|leaf|fresh_after|operation_after|;N4054c00000000000;S4:Long;N4008000000000000;A5[N0000000000000000;S0:;S0:;S0:;N0000000000000000];N4054c00000000000];A5[N0000000000000000;S0:;S0:;S0:;N0000000000000000]]
37	E:\oxi-vba-public-artifacts-20261115\types5052\owned-data-5053-0-37.bin	A2[A9[N4042800000000000;N0000000000000000;N4022000000000000;S99:seed_before|parent|leaf|seed_after|fresh_enter|parent|leaf|fresh_after|byref_enter|operation_after|;N4054c00000000000;S4:Long;N4008000000000000;A5[N0000000000000000;S0:;S0:;S0:;N0000000000000000];N4054c00000000000];A5[N0000000000000000;S0:;S0:;S0:;N0000000000000000]]
38	E:\oxi-vba-public-artifacts-20261115\types5052\owned-data-5053-0-38.bin	A2[A9[N4043000000000000;N0000000000000000;N4024000000000000;S99:seed_before|parent|leaf|seed_after|fresh_enter|parent|leaf|fresh_after|byref_enter|operation_after|;N4054c00000000000;S4:Long;N4008000000000000;A5[N0000000000000000;S0:;S0:;S0:;N0000000000000000];N4054c00000000000];A5[N0000000000000000;S0:;S0:;S0:;N0000000000000000]]
39	E:\oxi-vba-public-artifacts-20261115\types5052\owned-data-5053-0-39.bin	A2[A9[N4043800000000000;N0000000000000000;N4026000000000000;S87:seed_before|parent|leaf|seed_after|fresh_enter|parent|leaf|fresh_after|operation_after|;E;S5:Empty;N0000000000000000;A5[N0000000000000000;S0:;S0:;S0:;N0000000000000000];N4054c00000000000];A5[N407ca00000000000;S109:Visual Basic でサポートされていないオートメーションが変数で使用されています。;S10:VBAProject;S72:C:\Program Files\Common Files\Microsoft Shared\VBA\VBA7.1\1041\VbLR6.chm;N412e881400000000]]
40	E:\oxi-vba-public-artifacts-20261115\types5052\owned-data-5053-0-40.bin	A2[A9[N4044000000000000;N0000000000000000;N4028000000000000;S87:seed_before|parent|leaf|seed_after|fresh_enter|parent|leaf|fresh_after|operation_after|;E;S5:Empty;N0000000000000000;A5[N0000000000000000;S0:;S0:;S0:;N0000000000000000];N4054c00000000000];A5[N407ca00000000000;S109:Visual Basic でサポートされていないオートメーションが変数で使用されています。;S10:VBAProject;S72:C:\Program Files\Common Files\Microsoft Shared\VBA\VBA7.1\1041\VbLR6.chm;N412e881400000000]]
41	E:\oxi-vba-public-artifacts-20261115\types5052\owned-data-5053-0-41.bin	A2[A9[N4044800000000000;N0000000000000000;N402a000000000000;S87:seed_before|parent|leaf|seed_after|fresh_enter|parent|leaf|fresh_after|operation_after|;N4054c00000000000;S4:Long;N4008000000000000;A5[N0000000000000000;S0:;S0:;S0:;N0000000000000000];N4054c00000000000];A5[N0000000000000000;S0:;S0:;S0:;N0000000000000000]]
42	E:\oxi-vba-public-artifacts-20261115\types5052\owned-data-5053-0-42.bin	A2[A9[N4045000000000000;N3ff0000000000000;N0000000000000000;S51:seed_before|parent|leaf|seed_after|operation_after|;E;S5:Empty;N0000000000000000;A5[N0000000000000000;S0:;S0:;S0:;N0000000000000000];N4054c00000000000];A5[N407ca00000000000;S109:Visual Basic でサポートされていないオートメーションが変数で使用されています。;S10:VBAProject;S72:C:\Program Files\Common Files\Microsoft Shared\VBA\VBA7.1\1041\VbLR6.chm;N412e881400000000]]
43	E:\oxi-vba-public-artifacts-20261115\types5052\owned-data-5053-0-43.bin	A2[A9[N4045800000000000;N3ff0000000000000;N3ff0000000000000;S51:seed_before|parent|leaf|seed_after|operation_after|;E;S5:Empty;N0000000000000000;A5[N0000000000000000;S0:;S0:;S0:;N0000000000000000];N4054c00000000000];A5[N407ca00000000000;S109:Visual Basic でサポートされていないオートメーションが変数で使用されています。;S10:VBAProject;S72:C:\Program Files\Common Files\Microsoft Shared\VBA\VBA7.1\1041\VbLR6.chm;N412e881400000000]]
44	E:\oxi-vba-public-artifacts-20261115\types5052\owned-data-5053-0-44.bin	A2[A9[N4046000000000000;N3ff0000000000000;N4000000000000000;S51:seed_before|parent|leaf|seed_after|operation_after|;E;S5:Empty;N0000000000000000;A5[N0000000000000000;S0:;S0:;S0:;N0000000000000000];N4054c00000000000];A5[N407ca00000000000;S109:Visual Basic でサポートされていないオートメーションが変数で使用されています。;S10:VBAProject;S72:C:\Program Files\Common Files\Microsoft Shared\VBA\VBA7.1\1041\VbLR6.chm;N412e881400000000]]
45	E:\oxi-vba-public-artifacts-20261115\types5052\owned-data-5053-0-45.bin	A2[A9[N4046800000000000;N3ff0000000000000;N4008000000000000;S51:seed_before|parent|leaf|seed_after|operation_after|;E;S5:Empty;N0000000000000000;A5[N0000000000000000;S0:;S0:;S0:;N0000000000000000];N4054c00000000000];A5[N407ca00000000000;S109:Visual Basic でサポートされていないオートメーションが変数で使用されています。;S10:VBAProject;S72:C:\Program Files\Common Files\Microsoft Shared\VBA\VBA7.1\1041\VbLR6.chm;N412e881400000000]]
46	E:\oxi-vba-public-artifacts-20261115\types5052\owned-data-5053-0-46.bin	A2[A9[N4047000000000000;N3ff0000000000000;N4010000000000000;S51:seed_before|parent|leaf|seed_after|operation_after|;E;S5:Empty;N0000000000000000;A5[N0000000000000000;S0:;S0:;S0:;N0000000000000000];N4054c00000000000];A5[N407ca00000000000;S109:Visual Basic でサポートされていないオートメーションが変数で使用されています。;S10:VBAProject;S72:C:\Program Files\Common Files\Microsoft Shared\VBA\VBA7.1\1041\VbLR6.chm;N412e881400000000]]
47	E:\oxi-vba-public-artifacts-20261115\types5052\owned-data-5053-0-47.bin	A2[A9[N4047800000000000;N3ff0000000000000;N4014000000000000;S51:seed_before|parent|leaf|seed_after|operation_after|;E;S5:Empty;N0000000000000000;A5[N0000000000000000;S0:;S0:;S0:;N0000000000000000];N4054c00000000000];A5[N407ca00000000000;S109:Visual Basic でサポートされていないオートメーションが変数で使用されています。;S10:VBAProject;S72:C:\Program Files\Common Files\Microsoft Shared\VBA\VBA7.1\1041\VbLR6.chm;N412e881400000000]]
48	E:\oxi-vba-public-artifacts-20261115\types5052\owned-data-5053-0-48.bin	A2[A9[N4048000000000000;N3ff0000000000000;N4018000000000000;S51:seed_before|parent|leaf|seed_after|operation_after|;E;S5:Empty;N0000000000000000;A5[N0000000000000000;S0:;S0:;S0:;N0000000000000000];N4054c00000000000];A5[N407ca00000000000;S109:Visual Basic でサポートされていないオートメーションが変数で使用されています。;S10:VBAProject;S72:C:\Program Files\Common Files\Microsoft Shared\VBA\VBA7.1\1041\VbLR6.chm;N412e881400000000]]
49	E:\oxi-vba-public-artifacts-20261115\types5052\owned-data-5053-0-49.bin	A2[A9[N4048800000000000;N3ff0000000000000;N401c000000000000;S51:seed_before|parent|leaf|seed_after|operation_after|;E;S5:Empty;N0000000000000000;A5[N0000000000000000;S0:;S0:;S0:;N0000000000000000];N4054c00000000000];A5[N407ca00000000000;S109:Visual Basic でサポートされていないオートメーションが変数で使用されています。;S10:VBAProject;S72:C:\Program Files\Common Files\Microsoft Shared\VBA\VBA7.1\1041\VbLR6.chm;N412e881400000000]]
50	E:\oxi-vba-public-artifacts-20261115\types5052\owned-data-5053-0-50.bin	A2[A9[N4049000000000000;N3ff0000000000000;N4020000000000000;S87:seed_before|parent|leaf|seed_after|fresh_enter|parent|leaf|fresh_after|operation_after|;N4054c00000000000;S4:Long;N4008000000000000;A5[N0000000000000000;S0:;S0:;S0:;N0000000000000000];N4054c00000000000];A5[N0000000000000000;S0:;S0:;S0:;N0000000000000000]]
51	E:\oxi-vba-public-artifacts-20261115\types5052\owned-data-5053-0-51.bin	A2[A9[N4049800000000000;N3ff0000000000000;N4022000000000000;S99:seed_before|parent|leaf|seed_after|fresh_enter|parent|leaf|fresh_after|byref_enter|operation_after|;N4054c00000000000;S4:Long;N4008000000000000;A5[N0000000000000000;S0:;S0:;S0:;N0000000000000000];N4054c00000000000];A5[N0000000000000000;S0:;S0:;S0:;N0000000000000000]]
52	E:\oxi-vba-public-artifacts-20261115\types5052\owned-data-5053-0-52.bin	A2[A9[N404a000000000000;N3ff0000000000000;N4024000000000000;S99:seed_before|parent|leaf|seed_after|fresh_enter|parent|leaf|fresh_after|byref_enter|operation_after|;N4054c00000000000;S4:Long;N4008000000000000;A5[N0000000000000000;S0:;S0:;S0:;N0000000000000000];N4054c00000000000];A5[N0000000000000000;S0:;S0:;S0:;N0000000000000000]]
53	E:\oxi-vba-public-artifacts-20261115\types5052\owned-data-5053-0-53.bin	A2[A9[N404a800000000000;N3ff0000000000000;N4026000000000000;S87:seed_before|parent|leaf|seed_after|fresh_enter|parent|leaf|fresh_after|operation_after|;E;S5:Empty;N0000000000000000;A5[N0000000000000000;S0:;S0:;S0:;N0000000000000000];N4054c00000000000];A5[N407ca00000000000;S109:Visual Basic でサポートされていないオートメーションが変数で使用されています。;S10:VBAProject;S72:C:\Program Files\Common Files\Microsoft Shared\VBA\VBA7.1\1041\VbLR6.chm;N412e881400000000]]
54	E:\oxi-vba-public-artifacts-20261115\types5052\owned-data-5053-0-54.bin	A2[A9[N404b000000000000;N3ff0000000000000;N4028000000000000;S87:seed_before|parent|leaf|seed_after|fresh_enter|parent|leaf|fresh_after|operation_after|;E;S5:Empty;N0000000000000000;A5[N0000000000000000;S0:;S0:;S0:;N0000000000000000];N4054c00000000000];A5[N407ca00000000000;S109:Visual Basic でサポートされていないオートメーションが変数で使用されています。;S10:VBAProject;S72:C:\Program Files\Common Files\Microsoft Shared\VBA\VBA7.1\1041\VbLR6.chm;N412e881400000000]]
55	E:\oxi-vba-public-artifacts-20261115\types5052\owned-data-5053-0-55.bin	A2[A9[N404b800000000000;N3ff0000000000000;N402a000000000000;S87:seed_before|parent|leaf|seed_after|fresh_enter|parent|leaf|fresh_after|operation_after|;N4054c00000000000;S4:Long;N4008000000000000;A5[N0000000000000000;S0:;S0:;S0:;N0000000000000000];N4054c00000000000];A5[N0000000000000000;S0:;S0:;S0:;N0000000000000000]]"####, r####"C:\Program Files\Common Files\Microsoft Shared\VBA\VBA7.1\1041\VbLR6.chm"####);
}
