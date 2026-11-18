// SPDX-License-Identifier: MIT OR Apache-2.0

use super::*;
use serde_json::json;

#[test]
fn automation_subtype_output_is_an_opaque_descriptor_not_a_reusable_payload() {
    for type_code in [16, 82, 83, 84, 95, 99] {
        let descriptor = serde_json::to_value(OutputValue::from(
            Value::AutomationVariant(type_code),
        )).unwrap();
        assert_eq!(descriptor, json!({
            "type": "automation_variant", "value": { "type_code": type_code }
        }));
        assert_eq!(decode_typed_scalar(&descriptor).unwrap_err(),
            "Automation subtype descriptors do not contain a reusable value");
    }
}

#[test]
fn result_scalar_descriptors_preserve_vba_subtypes_when_reused_as_arguments() {
    let cases = [
        (Value::Empty, "empty"),
        (Value::Missing, "missing"),
        (Value::Nothing, "nothing"),
        (Value::Null, "null"),
        (Value::NullString, "null_string"),
        (Value::String(String::new()), "string"),
        (Value::Boolean(false), "boolean"),
        (Value::Boolean(true), "boolean"),
        (Value::Byte(0), "byte"),
        (Value::Byte(u8::MAX), "byte"),
        (Value::Int16(i16::MIN), "integer"),
        (Value::Int16(i16::MAX), "integer"),
        (Value::Integer(i64::from(i32::MIN)), "long"),
        (Value::Integer(i64::from(i32::MAX)), "long"),
        (Value::LongLong(i64::MIN), "long_long"),
        (Value::LongLong(i64::MAX), "long_long"),
        (Value::Single(1.234_567_9), "single"),
        (Value::Single(f32::from_bits(1)), "single"),
        (Value::Single(f32::MAX), "single"),
        (Value::Double(f64::MAX), "double"),
        (Value::Date(0.0), "date"),
        (Value::Date(-0.5), "date"),
        (Value::Date(43_890.502_835_648_15), "date"),
        (Value::Currency(i64::MIN), "currency"),
        (Value::Currency(i64::MAX), "currency"),
        (Value::Error(2007), "error"),
    ];
    for (value, tag) in cases {
        let descriptor = serde_json::to_value(OutputValue::from(value.clone())).unwrap();
        assert_eq!(descriptor["type"], tag, "{value:?}");
        if tag == "long_long" {
            assert!(descriptor["value"].is_string());
        }
        let restored = decode_typed_scalar(&descriptor).unwrap();
        assert_eq!(
            std::mem::discriminant(&restored),
            std::mem::discriminant(&value)
        );
        assert_eq!(restored, value);
    }
}

#[test]
fn result_float_descriptors_preserve_signed_zero_and_nonfinite_payload_bits() {
    for bits in [
        0u32,
        0x8000_0000,
        1,
        0x7f80_0000,
        0xff80_0000,
        0x7fc0_0042,
        0xffc0_0317,
    ] {
        let descriptor =
            serde_json::to_value(OutputValue::from(Value::Single(f32::from_bits(bits)))).unwrap();
        if bits == 0x8000_0000 {
            assert_eq!(descriptor["value"], json!({"bits":"80000000"}));
        }
        let Value::Single(restored) = decode_typed_scalar(&descriptor).unwrap() else {
            panic!("lost Single")
        };
        assert_eq!(restored.to_bits(), bits);
    }
    for bits in [
        0u64,
        0x8000_0000_0000_0000,
        1,
        0x7ff0_0000_0000_0000,
        0xfff0_0000_0000_0000,
        0x7ff8_0000_0000_0042,
        0xfff8_0000_0000_0317,
    ] {
        for value in [
            Value::Double(f64::from_bits(bits)),
            Value::Date(f64::from_bits(bits)),
        ] {
            let descriptor = serde_json::to_value(OutputValue::from(value.clone())).unwrap();
            if bits == 0x8000_0000_0000_0000 {
                assert_eq!(descriptor["value"], json!({"bits":"8000000000000000"}));
            }
            let restored = decode_typed_scalar(&descriptor).unwrap();
            assert_eq!(
                std::mem::discriminant(&restored),
                std::mem::discriminant(&value)
            );
            let (Value::Double(restored) | Value::Date(restored)) = restored else {
                panic!("lost floating type")
            };
            assert_eq!(restored.to_bits(), bits);
        }
    }
    for descriptor in [
        json!({"type":"single","value":{"bits":"7ff8000000000042"}}),
        json!({"type":"double","value":{"bits":"7fc00042"}}),
        json!({"type":"date","value":{"bits":"0x00000000000000"}}),
        json!({"type":"double","value":{"bits":"0000000000000000","extra":0}}),
        json!({"type":"double","value":{"bits":0}}),
    ] {
        assert!(decode_typed_scalar(&descriptor).is_err(), "{descriptor}");
    }
}

#[test]
fn result_arrays_retain_element_type_bounds_and_resize_contract() {
    for (default, values, tag) in [
        (
            Value::Empty,
            vec![Value::Byte(7), Value::Date(43_890.0), Value::NullString],
            "variant",
        ),
        (
            Value::Int16(0),
            vec![Value::Int16(-7), Value::Int16(0), Value::Int16(13)],
            "integer",
        ),
        (
            Value::Integer(0),
            vec![
                Value::Integer(i32::MIN.into()),
                Value::Integer(0),
                Value::Integer(i32::MAX.into()),
            ],
            "long",
        ),
        (
            Value::LongLong(0),
            vec![
                Value::LongLong(i64::MIN),
                Value::LongLong(0),
                Value::LongLong(i64::MAX),
            ],
            "long_long",
        ),
        (
            Value::Single(0.0),
            vec![Value::Single(1.25), Value::Single(0.0), Value::Single(-2.5)],
            "single",
        ),
        (
            Value::Date(0.0),
            vec![Value::Date(-0.5), Value::Date(0.0), Value::Date(43_890.5)],
            "date",
        ),
        (
            Value::NullString,
            vec![
                Value::NullString,
                Value::String("one".into()),
                Value::String(String::new()),
            ],
            "string",
        ),
        (
            Value::Nothing,
            vec![Value::Nothing, Value::Nothing, Value::Nothing],
            "object",
        ),
    ] {
        for resizable in [false, true] {
            let array = ArrayValue {
                dimensions: vec![ArrayDimension {
                    lower_bound: -2,
                    length: 3,
                }],
                values: values.clone(),
                element_default: Box::new(default.clone()),
                resizable,
            };
            let descriptor =
                serde_json::to_value(OutputValue::from(Value::Array(array.clone()))).unwrap();
            let payload = &descriptor["value"];
            assert_eq!(payload["element_type"], tag);
            assert_eq!(payload["lower_bound"], -2);
            assert_eq!(
                payload["dimensions"],
                json!([{"lower_bound":-2,"length":3}])
            );
            assert_eq!(payload["resizable"], resizable);
            let metadata = json!({"element_type":payload["element_type"],"dimensions":payload["dimensions"],"resizable":payload["resizable"]});
            let schema = decode_vba_input_array_schema(&metadata, 3).unwrap();
            let values = payload["values"]
                .as_array()
                .unwrap()
                .iter()
                .map(decode_typed_scalar)
                .collect::<Result<Vec<_>, _>>()
                .unwrap();
            assert_eq!(
                finish_vba_input_array(schema, values).unwrap(),
                Value::Array(array)
            );
        }
    }
    for dimensions in [
        vec![],
        vec![ArrayDimension {
            lower_bound: 0,
            length: 0,
        }],
    ] {
        let descriptor = serde_json::to_value(OutputValue::from(Value::Array(ArrayValue {
            dimensions: dimensions.clone(),
            values: vec![],
            element_default: Box::new(Value::Int16(0)),
            resizable: true,
        })))
        .unwrap();
        assert_eq!(descriptor["value"]["element_type"], "integer");
        assert_eq!(
            descriptor["value"]["dimensions"].as_array().unwrap().len(),
            dimensions.len()
        );
    }
}

// Native Excel source SHA-256 31b715600bec8fb03b43670167757a9a250f9f06ff8c39eae38d89f1d0ae2359; only physical Mark IO is omitted.
    fn len_automation_workbook() -> Workbook {
        Workbook {
            sheets: vec![Sheet {
                kind: Default::default(),
                name: "Sheet1".to_string(),
                visibility: Default::default(),
                rows: Vec::new(),
                col_count: 0,
                col_widths: Vec::new(),
                col_custom_widths: Default::default(),
                best_fit_cols: Default::default(),
                cached_spills: None,
                pending_used_bounds: None,
                column_width_font: None,
                default_col_width: 0.0,
                default_col_width_explicit: false,
                default_row_height: 15.0,
                default_row_custom: false,
                col_fonts: vec![],
                band_styles: Default::default(),
                normal_font: None,
                first_font: None,
                frozen_rows: 0,
                frozen_cols: 0,
                drawings: Vec::new(),
                comments: Vec::new(),
                threaded_comments: Vec::new(),
                merge_cells: Vec::new(),
                hidden_cols: Vec::new(),
                auto_filter: None,
                declared_range: None,
                tables: Vec::new(),
                unsupported_elements: Vec::new(),
                hyperlinks: Vec::new(),
                protected: false,
                protection: None,
                form_controls: Vec::new(),
                validations: Vec::new(),
                conditional_rules: Vec::new(),
                tab_color: None,
                drawing_id_blocks: Vec::new(),
        legacy_drawing_order: Vec::new(),
        drawing_groups: Vec::new(),
                header_footer: None,
                origin_id: None,
            }],
            ..Default::default()
        }
    }
fn len_cell_snapshot(value: &Value) -> String {
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
            array.values.iter().map(len_cell_snapshot).collect::<Vec<_>>().join(";")),
        other => panic!("unexpected recorded value kind: {other:?}"),
    }
}


#[test]
fn native_len_cell_boundary_matches_excel() {
    use oxivba_core::{ErrorEnvironment, ErrorLanguage};
    let (module, classes) = oxivba_core::parser::parse_project(r####"Attribute VB_Name = "Consume5031"
Option Explicit
Private target As String, trace As String
Private leafResult As Long
Public Function GetLeafResult() As Long
 GetLeafResult=leafResult
End Function
Public Sub Mark(ByVal token As String)
 trace=trace & token & "|"
End Sub
Public Function Probe(ByVal kind As Long,ByVal pathname As String) As Variant
 Dim item As Object, value As Variant, copied As Variant, observed As Variant
 Dim number As Long, seed As Long, operation As Long, seedErr As Variant, savedErr As Variant, tags As Variant
 target=pathname:trace="":tags=Array(16,83):leafResult=tags(kind\28):seed=(kind Mod 28)\14:operation=kind Mod 14
 On Error Resume Next
 Err.Clear
 ThisWorkbook.Worksheets(1).Cells(1,1).Value=123
 Set item=New Parent5031
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
 Case 0: ThisWorkbook.Worksheets(1).Cells(1,1).Value=value
 Case 1: ThisWorkbook.Worksheets(1).Cells(1,1).Value2=value
 Case 2
  If seed=0 Then
   ThisWorkbook.Worksheets(1).Cells(1,1).Value=Len(item)
  Else
   ThisWorkbook.Worksheets(1).Cells(1,1).Value=LenB(item)
  End If
 Case 3
  If seed=0 Then
   ThisWorkbook.Worksheets(1).Cells(1,1).Value2=Len(item)
  Else
   ThisWorkbook.Worksheets(1).Cells(1,1).Value2=LenB(item)
  End If
 End Select
 savedErr=Array(Err.Number,Err.Description,Err.Source,Err.HelpFile,Err.HelpContext)
 Mark "operation_after"
 observed=Array(ThisWorkbook.Worksheets(1).Cells(1,1).Value2,TypeName(ThisWorkbook.Worksheets(1).Cells(1,1).Value2),VarType(ThisWorkbook.Worksheets(1).Cells(1,1).Value2))
 Probe=Array(Array(kind,seed,operation,trace,observed,TypeName(observed),VarType(observed),seedErr,leafResult),savedErr)
End Function
VERSION 1.0 CLASS
Attribute VB_Name = "Leaf5031"
Option Explicit
Public Property Get Scalar() As Long
Attribute Scalar.VB_UserMemId = 0
 Consume5031.Mark "leaf"
 Scalar=Consume5031.GetLeafResult()
End Property
VERSION 1.0 CLASS
Attribute VB_Name = "Parent5031"
Option Explicit
Public Property Get Scalar() As Object
Attribute Scalar.VB_UserMemId = 0
 Consume5031.Mark "parent"
 Set Scalar=New Leaf5031
End Property
"####).unwrap();
    for row in r####"0	E:\oxi-vba-public-artifacts-20261115\types5031\owned-data-5032-0-0.bin	A2[A9[N0000000000000000;N0000000000000000;N0000000000000000;S51:seed_before|parent|leaf|seed_after|operation_after|;A3[N0000000000000000;S6:Double;N4014000000000000];S9:Variant();N40c0060000000000;A5[N0000000000000000;S0:;S0:;S0:;N0000000000000000];N4030000000000000];A5[N0000000000000000;S0:;S0:;S0:;N0000000000000000]]
1	E:\oxi-vba-public-artifacts-20261115\types5031\owned-data-5032-0-1.bin	A2[A9[N3ff0000000000000;N0000000000000000;N3ff0000000000000;S51:seed_before|parent|leaf|seed_after|operation_after|;A3[N0000000000000000;S6:Double;N4014000000000000];S9:Variant();N40c0060000000000;A5[N0000000000000000;S0:;S0:;S0:;N0000000000000000];N4030000000000000];A5[N0000000000000000;S0:;S0:;S0:;N0000000000000000]]
2	E:\oxi-vba-public-artifacts-20261115\types5031\owned-data-5032-0-2.bin	A2[A9[N4000000000000000;N0000000000000000;N4000000000000000;S63:seed_before|parent|leaf|seed_after|parent|leaf|operation_after|;A3[N0000000000000000;S6:Double;N4014000000000000];S9:Variant();N40c0060000000000;A5[N0000000000000000;S0:;S0:;S0:;N0000000000000000];N4030000000000000];A5[N0000000000000000;S0:;S0:;S0:;N0000000000000000]]
3	E:\oxi-vba-public-artifacts-20261115\types5031\owned-data-5032-0-3.bin	A2[A9[N4008000000000000;N0000000000000000;N4008000000000000;S63:seed_before|parent|leaf|seed_after|parent|leaf|operation_after|;A3[N0000000000000000;S6:Double;N4014000000000000];S9:Variant();N40c0060000000000;A5[N0000000000000000;S0:;S0:;S0:;N0000000000000000];N4030000000000000];A5[N0000000000000000;S0:;S0:;S0:;N0000000000000000]]
14	E:\oxi-vba-public-artifacts-20261115\types5031\owned-data-5032-0-14.bin	A2[A9[N402c000000000000;N3ff0000000000000;N0000000000000000;S51:seed_before|parent|leaf|seed_after|operation_after|;A3[N0000000000000000;S6:Double;N4014000000000000];S9:Variant();N40c0060000000000;A5[N0000000000000000;S0:;S0:;S0:;N0000000000000000];N4030000000000000];A5[N0000000000000000;S0:;S0:;S0:;N0000000000000000]]
15	E:\oxi-vba-public-artifacts-20261115\types5031\owned-data-5032-0-15.bin	A2[A9[N402e000000000000;N3ff0000000000000;N3ff0000000000000;S51:seed_before|parent|leaf|seed_after|operation_after|;A3[N0000000000000000;S6:Double;N4014000000000000];S9:Variant();N40c0060000000000;A5[N0000000000000000;S0:;S0:;S0:;N0000000000000000];N4030000000000000];A5[N0000000000000000;S0:;S0:;S0:;N0000000000000000]]
16	E:\oxi-vba-public-artifacts-20261115\types5031\owned-data-5032-0-16.bin	A2[A9[N4030000000000000;N3ff0000000000000;N4000000000000000;S63:seed_before|parent|leaf|seed_after|parent|leaf|operation_after|;A3[N0000000000000000;S6:Double;N4014000000000000];S9:Variant();N40c0060000000000;A5[N0000000000000000;S0:;S0:;S0:;N0000000000000000];N4030000000000000];A5[N0000000000000000;S0:;S0:;S0:;N0000000000000000]]
17	E:\oxi-vba-public-artifacts-20261115\types5031\owned-data-5032-0-17.bin	A2[A9[N4031000000000000;N3ff0000000000000;N4008000000000000;S63:seed_before|parent|leaf|seed_after|parent|leaf|operation_after|;A3[N0000000000000000;S6:Double;N4014000000000000];S9:Variant();N40c0060000000000;A5[N0000000000000000;S0:;S0:;S0:;N0000000000000000];N4030000000000000];A5[N0000000000000000;S0:;S0:;S0:;N0000000000000000]]
28	E:\oxi-vba-public-artifacts-20261115\types5031\owned-data-5032-0-28.bin	A2[A9[N403c000000000000;N0000000000000000;N0000000000000000;S51:seed_before|parent|leaf|seed_after|operation_after|;A3[N405ec00000000000;S6:Double;N4014000000000000];S9:Variant();N40c0060000000000;A5[N0000000000000000;S0:;S0:;S0:;N0000000000000000];N4054c00000000000];A5[N408f600000000000;S84:アプリケーション定義またはオブジェクト定義のエラーです。;S10:VBAProject;S72:C:\Program Files\Common Files\Microsoft Shared\VBA\VBA7.1\1041\VbLR6.chm;N412e853e00000000]]
29	E:\oxi-vba-public-artifacts-20261115\types5031\owned-data-5032-0-29.bin	A2[A9[N403d000000000000;N0000000000000000;N3ff0000000000000;S51:seed_before|parent|leaf|seed_after|operation_after|;A3[N405ec00000000000;S6:Double;N4014000000000000];S9:Variant();N40c0060000000000;A5[N0000000000000000;S0:;S0:;S0:;N0000000000000000];N4054c00000000000];A5[N408f600000000000;S84:アプリケーション定義またはオブジェクト定義のエラーです。;S10:VBAProject;S72:C:\Program Files\Common Files\Microsoft Shared\VBA\VBA7.1\1041\VbLR6.chm;N412e853e00000000]]
30	E:\oxi-vba-public-artifacts-20261115\types5031\owned-data-5032-0-30.bin	A2[A9[N403e000000000000;N0000000000000000;N4000000000000000;S63:seed_before|parent|leaf|seed_after|parent|leaf|operation_after|;A3[N405ec00000000000;S6:Double;N4014000000000000];S9:Variant();N40c0060000000000;A5[N0000000000000000;S0:;S0:;S0:;N0000000000000000];N4054c00000000000];A5[N408f600000000000;S84:アプリケーション定義またはオブジェクト定義のエラーです。;S10:VBAProject;S72:C:\Program Files\Common Files\Microsoft Shared\VBA\VBA7.1\1041\VbLR6.chm;N412e853e00000000]]
31	E:\oxi-vba-public-artifacts-20261115\types5031\owned-data-5032-0-31.bin	A2[A9[N403f000000000000;N0000000000000000;N4008000000000000;S63:seed_before|parent|leaf|seed_after|parent|leaf|operation_after|;A3[N405ec00000000000;S6:Double;N4014000000000000];S9:Variant();N40c0060000000000;A5[N0000000000000000;S0:;S0:;S0:;N0000000000000000];N4054c00000000000];A5[N408f600000000000;S84:アプリケーション定義またはオブジェクト定義のエラーです。;S10:VBAProject;S72:C:\Program Files\Common Files\Microsoft Shared\VBA\VBA7.1\1041\VbLR6.chm;N412e853e00000000]]
42	E:\oxi-vba-public-artifacts-20261115\types5031\owned-data-5032-0-42.bin	A2[A9[N4045000000000000;N3ff0000000000000;N0000000000000000;S51:seed_before|parent|leaf|seed_after|operation_after|;A3[N405ec00000000000;S6:Double;N4014000000000000];S9:Variant();N40c0060000000000;A5[N0000000000000000;S0:;S0:;S0:;N0000000000000000];N4054c00000000000];A5[N408f600000000000;S84:アプリケーション定義またはオブジェクト定義のエラーです。;S10:VBAProject;S72:C:\Program Files\Common Files\Microsoft Shared\VBA\VBA7.1\1041\VbLR6.chm;N412e853e00000000]]
43	E:\oxi-vba-public-artifacts-20261115\types5031\owned-data-5032-0-43.bin	A2[A9[N4045800000000000;N3ff0000000000000;N3ff0000000000000;S51:seed_before|parent|leaf|seed_after|operation_after|;A3[N405ec00000000000;S6:Double;N4014000000000000];S9:Variant();N40c0060000000000;A5[N0000000000000000;S0:;S0:;S0:;N0000000000000000];N4054c00000000000];A5[N408f600000000000;S84:アプリケーション定義またはオブジェクト定義のエラーです。;S10:VBAProject;S72:C:\Program Files\Common Files\Microsoft Shared\VBA\VBA7.1\1041\VbLR6.chm;N412e853e00000000]]
44	E:\oxi-vba-public-artifacts-20261115\types5031\owned-data-5032-0-44.bin	A2[A9[N4046000000000000;N3ff0000000000000;N4000000000000000;S63:seed_before|parent|leaf|seed_after|parent|leaf|operation_after|;A3[N405ec00000000000;S6:Double;N4014000000000000];S9:Variant();N40c0060000000000;A5[N0000000000000000;S0:;S0:;S0:;N0000000000000000];N4054c00000000000];A5[N408f600000000000;S84:アプリケーション定義またはオブジェクト定義のエラーです。;S10:VBAProject;S72:C:\Program Files\Common Files\Microsoft Shared\VBA\VBA7.1\1041\VbLR6.chm;N412e853e00000000]]
45	E:\oxi-vba-public-artifacts-20261115\types5031\owned-data-5032-0-45.bin	A2[A9[N4046800000000000;N3ff0000000000000;N4008000000000000;S63:seed_before|parent|leaf|seed_after|parent|leaf|operation_after|;A3[N405ec00000000000;S6:Double;N4014000000000000];S9:Variant();N40c0060000000000;A5[N0000000000000000;S0:;S0:;S0:;N0000000000000000];N4054c00000000000];A5[N408f600000000000;S84:アプリケーション定義またはオブジェクト定義のエラーです。;S10:VBAProject;S72:C:\Program Files\Common Files\Microsoft Shared\VBA\VBA7.1\1041\VbLR6.chm;N412e853e00000000]]"####.lines() {
        let mut fields = row.splitn(3, '\t');
        let kind = fields.next().unwrap().parse::<i64>().unwrap();
        let pathname = fields.next().unwrap();
        let expected = fields.next().unwrap();
        let mut book = len_automation_workbook();
        let mut host = WorkbookHost::new(&mut book, 0).unwrap();
        host.error_language = ErrorLanguage::Japanese;
        let value = oxivba_core::runtime::Runtime::new(&module).with_classes(&classes)
            .with_project_name("VBAProject").with_host(&mut host)
            .with_error_environment(ErrorEnvironment { language: ErrorLanguage::Japanese, help_file: r####"C:\Program Files\Common Files\Microsoft Shared\VBA\VBA7.1\1041\VbLR6.chm"####.into() })
            .call("Probe", vec![Value::Integer(kind), Value::String(pathname.into())]).unwrap();
        assert_eq!(len_cell_snapshot(&value), expected, "Native cell boundary case={kind}");
    }
}
// Native source SHA-256 9e722038a11dd9564584f59b279f46721eb1e40eaadf898afdcf98eecd771162; physical Mark IO omitted, Native B2=0 seed preserved.

#[test]
fn native_len_find_and_print_matches_excel() {
    use oxivba_core::{ErrorEnvironment, ErrorLanguage};
    let (module, classes) = oxivba_core::parser::parse_project(r####"Attribute VB_Name = "Consume5040"
Option Explicit
Private target As String, trace As String
Private leafResult As Long
Public Function GetLeafResult() As Long
 GetLeafResult=leafResult
End Function
Public Sub Mark(ByVal token As String)
 trace=trace & token & "|"
End Sub
Public Function Probe(ByVal kind As Long,ByVal pathname As String) As Variant
 Dim item As Object, found As Object, value As Variant, copied As Variant, observed As Variant
 Dim number As Long, seed As Long, operation As Long, seedErr As Variant, savedErr As Variant, tags As Variant
 target=pathname:trace="":tags=Array(16,83):leafResult=tags(kind\28):seed=(kind Mod 28)\14:operation=kind Mod 14
 On Error Resume Next
 Err.Clear
 ThisWorkbook.Worksheets(1).Cells(1,1).Value=123
 ThisWorkbook.Worksheets(1).Cells(2,2).Value=0
 Set item=New Parent5040
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
 Case 0: Set found=ThisWorkbook.Worksheets(1).Cells.Find(What:=value):observed=(found Is Nothing)
 Case 1
  If seed=0 Then
   Set found=ThisWorkbook.Worksheets(1).Cells.Find(What:=Len(item))
  Else
   Set found=ThisWorkbook.Worksheets(1).Cells.Find(What:=LenB(item))
  End If
  observed=(found Is Nothing)
 Case 2: Debug.Print value:observed="print_after"
 Case 3
  If seed=0 Then
   Debug.Print Len(item)
  Else
   Debug.Print LenB(item)
  End If
  observed="print_after"
 End Select
 savedErr=Array(Err.Number,Err.Description,Err.Source,Err.HelpFile,Err.HelpContext)
 Mark "operation_after"
 Probe=Array(Array(kind,seed,operation,trace,observed,TypeName(observed),VarType(observed),seedErr,leafResult),savedErr)
End Function
VERSION 1.0 CLASS
Attribute VB_Name = "Leaf5040"
Option Explicit
Public Property Get Scalar() As Long
Attribute Scalar.VB_UserMemId = 0
 Consume5040.Mark "leaf"
 Scalar=Consume5040.GetLeafResult()
End Property
VERSION 1.0 CLASS
Attribute VB_Name = "Parent5040"
Option Explicit
Public Property Get Scalar() As Object
Attribute Scalar.VB_UserMemId = 0
 Consume5040.Mark "parent"
 Set Scalar=New Leaf5040
End Property
"####).unwrap();
    for row in r####"0	E:\oxi-vba-public-artifacts-20261115\types5040\owned-data-5041-0-0.bin	A2[A9[N0000000000000000;N0000000000000000;N0000000000000000;S51:seed_before|parent|leaf|seed_after|operation_after|;F;S7:Boolean;N4026000000000000;A5[N0000000000000000;S0:;S0:;S0:;N0000000000000000];N4030000000000000];A5[N0000000000000000;S0:;S0:;S0:;N0000000000000000]]
1	E:\oxi-vba-public-artifacts-20261115\types5040\owned-data-5041-0-1.bin	A2[A9[N3ff0000000000000;N0000000000000000;N3ff0000000000000;S63:seed_before|parent|leaf|seed_after|parent|leaf|operation_after|;F;S7:Boolean;N4026000000000000;A5[N0000000000000000;S0:;S0:;S0:;N0000000000000000];N4030000000000000];A5[N0000000000000000;S0:;S0:;S0:;N0000000000000000]]
2	E:\oxi-vba-public-artifacts-20261115\types5040\owned-data-5041-0-2.bin	A2[A9[N4000000000000000;N0000000000000000;N4000000000000000;S51:seed_before|parent|leaf|seed_after|operation_after|;S11:print_after;S6:String;N4020000000000000;A5[N0000000000000000;S0:;S0:;S0:;N0000000000000000];N4030000000000000];A5[N0000000000000000;S0:;S0:;S0:;N0000000000000000]]
3	E:\oxi-vba-public-artifacts-20261115\types5040\owned-data-5041-0-3.bin	A2[A9[N4008000000000000;N0000000000000000;N4008000000000000;S63:seed_before|parent|leaf|seed_after|parent|leaf|operation_after|;S11:print_after;S6:String;N4020000000000000;A5[N0000000000000000;S0:;S0:;S0:;N0000000000000000];N4030000000000000];A5[N0000000000000000;S0:;S0:;S0:;N0000000000000000]]
14	E:\oxi-vba-public-artifacts-20261115\types5040\owned-data-5041-0-14.bin	A2[A9[N402c000000000000;N3ff0000000000000;N0000000000000000;S51:seed_before|parent|leaf|seed_after|operation_after|;F;S7:Boolean;N4026000000000000;A5[N0000000000000000;S0:;S0:;S0:;N0000000000000000];N4030000000000000];A5[N0000000000000000;S0:;S0:;S0:;N0000000000000000]]
15	E:\oxi-vba-public-artifacts-20261115\types5040\owned-data-5041-0-15.bin	A2[A9[N402e000000000000;N3ff0000000000000;N3ff0000000000000;S63:seed_before|parent|leaf|seed_after|parent|leaf|operation_after|;F;S7:Boolean;N4026000000000000;A5[N0000000000000000;S0:;S0:;S0:;N0000000000000000];N4030000000000000];A5[N0000000000000000;S0:;S0:;S0:;N0000000000000000]]
16	E:\oxi-vba-public-artifacts-20261115\types5040\owned-data-5041-0-16.bin	A2[A9[N4030000000000000;N3ff0000000000000;N4000000000000000;S51:seed_before|parent|leaf|seed_after|operation_after|;S11:print_after;S6:String;N4020000000000000;A5[N0000000000000000;S0:;S0:;S0:;N0000000000000000];N4030000000000000];A5[N0000000000000000;S0:;S0:;S0:;N0000000000000000]]
17	E:\oxi-vba-public-artifacts-20261115\types5040\owned-data-5041-0-17.bin	A2[A9[N4031000000000000;N3ff0000000000000;N4008000000000000;S63:seed_before|parent|leaf|seed_after|parent|leaf|operation_after|;S11:print_after;S6:String;N4020000000000000;A5[N0000000000000000;S0:;S0:;S0:;N0000000000000000];N4030000000000000];A5[N0000000000000000;S0:;S0:;S0:;N0000000000000000]]
28	E:\oxi-vba-public-artifacts-20261115\types5040\owned-data-5041-0-28.bin	A2[A9[N403c000000000000;N0000000000000000;N0000000000000000;S51:seed_before|parent|leaf|seed_after|operation_after|;T;S7:Boolean;N4026000000000000;A5[N0000000000000000;S0:;S0:;S0:;N0000000000000000];N4054c00000000000];A5[N407ca00000000000;S109:Visual Basic でサポートされていないオートメーションが変数で使用されています。;S10:VBAProject;S72:C:\Program Files\Common Files\Microsoft Shared\VBA\VBA7.1\1041\VbLR6.chm;N412e881400000000]]
29	E:\oxi-vba-public-artifacts-20261115\types5040\owned-data-5041-0-29.bin	A2[A9[N403d000000000000;N0000000000000000;N3ff0000000000000;S63:seed_before|parent|leaf|seed_after|parent|leaf|operation_after|;T;S7:Boolean;N4026000000000000;A5[N0000000000000000;S0:;S0:;S0:;N0000000000000000];N4054c00000000000];A5[N407ca00000000000;S109:Visual Basic でサポートされていないオートメーションが変数で使用されています。;S10:VBAProject;S72:C:\Program Files\Common Files\Microsoft Shared\VBA\VBA7.1\1041\VbLR6.chm;N412e881400000000]]
30	E:\oxi-vba-public-artifacts-20261115\types5040\owned-data-5041-0-30.bin	A2[A9[N403e000000000000;N0000000000000000;N4000000000000000;S51:seed_before|parent|leaf|seed_after|operation_after|;S11:print_after;S6:String;N4020000000000000;A5[N0000000000000000;S0:;S0:;S0:;N0000000000000000];N4054c00000000000];A5[N407ca00000000000;S109:Visual Basic でサポートされていないオートメーションが変数で使用されています。;S10:VBAProject;S72:C:\Program Files\Common Files\Microsoft Shared\VBA\VBA7.1\1041\VbLR6.chm;N412e881400000000]]
31	E:\oxi-vba-public-artifacts-20261115\types5040\owned-data-5041-0-31.bin	A2[A9[N403f000000000000;N0000000000000000;N4008000000000000;S63:seed_before|parent|leaf|seed_after|parent|leaf|operation_after|;S11:print_after;S6:String;N4020000000000000;A5[N0000000000000000;S0:;S0:;S0:;N0000000000000000];N4054c00000000000];A5[N407ca00000000000;S109:Visual Basic でサポートされていないオートメーションが変数で使用されています。;S10:VBAProject;S72:C:\Program Files\Common Files\Microsoft Shared\VBA\VBA7.1\1041\VbLR6.chm;N412e881400000000]]
42	E:\oxi-vba-public-artifacts-20261115\types5040\owned-data-5041-0-42.bin	A2[A9[N4045000000000000;N3ff0000000000000;N0000000000000000;S51:seed_before|parent|leaf|seed_after|operation_after|;T;S7:Boolean;N4026000000000000;A5[N0000000000000000;S0:;S0:;S0:;N0000000000000000];N4054c00000000000];A5[N407ca00000000000;S109:Visual Basic でサポートされていないオートメーションが変数で使用されています。;S10:VBAProject;S72:C:\Program Files\Common Files\Microsoft Shared\VBA\VBA7.1\1041\VbLR6.chm;N412e881400000000]]
43	E:\oxi-vba-public-artifacts-20261115\types5040\owned-data-5041-0-43.bin	A2[A9[N4045800000000000;N3ff0000000000000;N3ff0000000000000;S63:seed_before|parent|leaf|seed_after|parent|leaf|operation_after|;T;S7:Boolean;N4026000000000000;A5[N0000000000000000;S0:;S0:;S0:;N0000000000000000];N4054c00000000000];A5[N407ca00000000000;S109:Visual Basic でサポートされていないオートメーションが変数で使用されています。;S10:VBAProject;S72:C:\Program Files\Common Files\Microsoft Shared\VBA\VBA7.1\1041\VbLR6.chm;N412e881400000000]]
44	E:\oxi-vba-public-artifacts-20261115\types5040\owned-data-5041-0-44.bin	A2[A9[N4046000000000000;N3ff0000000000000;N4000000000000000;S51:seed_before|parent|leaf|seed_after|operation_after|;S11:print_after;S6:String;N4020000000000000;A5[N0000000000000000;S0:;S0:;S0:;N0000000000000000];N4054c00000000000];A5[N407ca00000000000;S109:Visual Basic でサポートされていないオートメーションが変数で使用されています。;S10:VBAProject;S72:C:\Program Files\Common Files\Microsoft Shared\VBA\VBA7.1\1041\VbLR6.chm;N412e881400000000]]
45	E:\oxi-vba-public-artifacts-20261115\types5040\owned-data-5041-0-45.bin	A2[A9[N4046800000000000;N3ff0000000000000;N4008000000000000;S63:seed_before|parent|leaf|seed_after|parent|leaf|operation_after|;S11:print_after;S6:String;N4020000000000000;A5[N0000000000000000;S0:;S0:;S0:;N0000000000000000];N4054c00000000000];A5[N407ca00000000000;S109:Visual Basic でサポートされていないオートメーションが変数で使用されています。;S10:VBAProject;S72:C:\Program Files\Common Files\Microsoft Shared\VBA\VBA7.1\1041\VbLR6.chm;N412e881400000000]]"####.lines() {
        let mut fields = row.splitn(3, '\t');
        let kind = fields.next().unwrap().parse::<i64>().unwrap();
        let pathname = fields.next().unwrap();
        let expected = fields.next().unwrap();
        let mut book = len_automation_workbook();
        let mut host = WorkbookHost::new(&mut book, 0).unwrap();
        host.error_language = ErrorLanguage::Japanese;
        let value = oxivba_core::runtime::Runtime::new(&module).with_classes(&classes)
            .with_project_name("VBAProject").with_host(&mut host)
            .with_error_environment(ErrorEnvironment { language: ErrorLanguage::Japanese, help_file: r####"C:\Program Files\Common Files\Microsoft Shared\VBA\VBA7.1\1041\VbLR6.chm"####.into() })
            .call("Probe", vec![Value::Integer(kind), Value::String(pathname.into())]).unwrap();
        assert_eq!(len_cell_snapshot(&value), expected, "Native host consumption case={kind}");
    }
}

// Native source SHA-256 93640b7b25d2b100f7008631f9d0c50bfbaf2f11a2c546ff3f26721f66ab8971; physical Mark IO omitted, Native B2=0 seed preserved.

#[test]
fn native_len_find_address_matches_excel() {
    use oxivba_core::{ErrorEnvironment, ErrorLanguage};
    let (module, classes) = oxivba_core::parser::parse_project(r####"Attribute VB_Name = "Consume5043"
Option Explicit
Private target As String, trace As String
Private leafResult As Long
Public Function GetLeafResult() As Long
 GetLeafResult=leafResult
End Function
Public Sub Mark(ByVal token As String)
 trace=trace & token & "|"
End Sub
Public Function Probe(ByVal kind As Long,ByVal pathname As String) As Variant
 Dim item As Object, found As Object, value As Variant, copied As Variant, observed As Variant
 Dim number As Long, seed As Long, operation As Long, seedErr As Variant, savedErr As Variant, tags As Variant
 target=pathname:trace="":tags=Array(16,83):leafResult=tags(kind\28):seed=(kind Mod 28)\14:operation=kind Mod 14
 On Error Resume Next
 Err.Clear
 ThisWorkbook.Worksheets(1).Cells(1,1).Value=123
 ThisWorkbook.Worksheets(1).Cells(2,2).Value=0
 ThisWorkbook.Worksheets(1).Cells(1,2).Value="alpha"
 ThisWorkbook.Worksheets(1).Cells(1,3).Value=0
 ThisWorkbook.Worksheets(1).Cells(1,4).Value="0"
 Set item=New Parent5043
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
 Case 0: Set found=ThisWorkbook.Worksheets(1).Cells.Find(What:=value)
 Case 1
  If seed=0 Then
   Set found=ThisWorkbook.Worksheets(1).Cells.Find(What:=Len(item))
  Else
   Set found=ThisWorkbook.Worksheets(1).Cells.Find(What:=LenB(item))
  End If
 Case 2: Set found=ThisWorkbook.Worksheets(1).Cells.Find(What:=CLng(value))
 Case 3: Set found=ThisWorkbook.Worksheets(1).Cells.Find(What:=0)
 End Select
 savedErr=Array(Err.Number,Err.Description,Err.Source,Err.HelpFile,Err.HelpContext)
 If found Is Nothing Then
  observed=Array(True,"","")
 Else
  observed=Array(False,found.Address,CStr(found.Value2))
 End If
 Mark "operation_after"
 Probe=Array(Array(kind,seed,operation,trace,observed,TypeName(observed),VarType(observed),seedErr,leafResult),savedErr)
End Function
VERSION 1.0 CLASS
Attribute VB_Name = "Leaf5043"
Option Explicit
Public Property Get Scalar() As Long
Attribute Scalar.VB_UserMemId = 0
 Consume5043.Mark "leaf"
 Scalar=Consume5043.GetLeafResult()
End Property
VERSION 1.0 CLASS
Attribute VB_Name = "Parent5043"
Option Explicit
Public Property Get Scalar() As Object
Attribute Scalar.VB_UserMemId = 0
 Consume5043.Mark "parent"
 Set Scalar=New Leaf5043
End Property
"####).unwrap();
    for row in r####"0	E:\oxi-vba-public-artifacts-20261115\types5043\owned-data-5044-0-0.bin	A2[A9[N0000000000000000;N0000000000000000;N0000000000000000;S51:seed_before|parent|leaf|seed_after|operation_after|;A3[F;S4:$C$1;S1:0];S9:Variant();N40c0060000000000;A5[N0000000000000000;S0:;S0:;S0:;N0000000000000000];N4030000000000000];A5[N0000000000000000;S0:;S0:;S0:;N0000000000000000]]
1	E:\oxi-vba-public-artifacts-20261115\types5043\owned-data-5044-0-1.bin	A2[A9[N3ff0000000000000;N0000000000000000;N3ff0000000000000;S63:seed_before|parent|leaf|seed_after|parent|leaf|operation_after|;A3[F;S4:$C$1;S1:0];S9:Variant();N40c0060000000000;A5[N0000000000000000;S0:;S0:;S0:;N0000000000000000];N4030000000000000];A5[N0000000000000000;S0:;S0:;S0:;N0000000000000000]]
2	E:\oxi-vba-public-artifacts-20261115\types5043\owned-data-5044-0-2.bin	A2[A9[N4000000000000000;N0000000000000000;N4000000000000000;S51:seed_before|parent|leaf|seed_after|operation_after|;A3[F;S4:$C$1;S1:0];S9:Variant();N40c0060000000000;A5[N0000000000000000;S0:;S0:;S0:;N0000000000000000];N4030000000000000];A5[N0000000000000000;S0:;S0:;S0:;N0000000000000000]]
3	E:\oxi-vba-public-artifacts-20261115\types5043\owned-data-5044-0-3.bin	A2[A9[N4008000000000000;N0000000000000000;N4008000000000000;S51:seed_before|parent|leaf|seed_after|operation_after|;A3[F;S4:$C$1;S1:0];S9:Variant();N40c0060000000000;A5[N0000000000000000;S0:;S0:;S0:;N0000000000000000];N4030000000000000];A5[N0000000000000000;S0:;S0:;S0:;N0000000000000000]]
14	E:\oxi-vba-public-artifacts-20261115\types5043\owned-data-5044-0-14.bin	A2[A9[N402c000000000000;N3ff0000000000000;N0000000000000000;S51:seed_before|parent|leaf|seed_after|operation_after|;A3[F;S4:$C$1;S1:0];S9:Variant();N40c0060000000000;A5[N0000000000000000;S0:;S0:;S0:;N0000000000000000];N4030000000000000];A5[N0000000000000000;S0:;S0:;S0:;N0000000000000000]]
15	E:\oxi-vba-public-artifacts-20261115\types5043\owned-data-5044-0-15.bin	A2[A9[N402e000000000000;N3ff0000000000000;N3ff0000000000000;S63:seed_before|parent|leaf|seed_after|parent|leaf|operation_after|;A3[F;S4:$C$1;S1:0];S9:Variant();N40c0060000000000;A5[N0000000000000000;S0:;S0:;S0:;N0000000000000000];N4030000000000000];A5[N0000000000000000;S0:;S0:;S0:;N0000000000000000]]
16	E:\oxi-vba-public-artifacts-20261115\types5043\owned-data-5044-0-16.bin	A2[A9[N4030000000000000;N3ff0000000000000;N4000000000000000;S51:seed_before|parent|leaf|seed_after|operation_after|;A3[F;S4:$C$1;S1:0];S9:Variant();N40c0060000000000;A5[N0000000000000000;S0:;S0:;S0:;N0000000000000000];N4030000000000000];A5[N0000000000000000;S0:;S0:;S0:;N0000000000000000]]
17	E:\oxi-vba-public-artifacts-20261115\types5043\owned-data-5044-0-17.bin	A2[A9[N4031000000000000;N3ff0000000000000;N4008000000000000;S51:seed_before|parent|leaf|seed_after|operation_after|;A3[F;S4:$C$1;S1:0];S9:Variant();N40c0060000000000;A5[N0000000000000000;S0:;S0:;S0:;N0000000000000000];N4030000000000000];A5[N0000000000000000;S0:;S0:;S0:;N0000000000000000]]
28	E:\oxi-vba-public-artifacts-20261115\types5043\owned-data-5044-0-28.bin	A2[A9[N403c000000000000;N0000000000000000;N0000000000000000;S51:seed_before|parent|leaf|seed_after|operation_after|;A3[T;S0:;S0:];S9:Variant();N40c0060000000000;A5[N0000000000000000;S0:;S0:;S0:;N0000000000000000];N4054c00000000000];A5[N407ca00000000000;S109:Visual Basic でサポートされていないオートメーションが変数で使用されています。;S10:VBAProject;S72:C:\Program Files\Common Files\Microsoft Shared\VBA\VBA7.1\1041\VbLR6.chm;N412e881400000000]]
29	E:\oxi-vba-public-artifacts-20261115\types5043\owned-data-5044-0-29.bin	A2[A9[N403d000000000000;N0000000000000000;N3ff0000000000000;S63:seed_before|parent|leaf|seed_after|parent|leaf|operation_after|;A3[T;S0:;S0:];S9:Variant();N40c0060000000000;A5[N0000000000000000;S0:;S0:;S0:;N0000000000000000];N4054c00000000000];A5[N407ca00000000000;S109:Visual Basic でサポートされていないオートメーションが変数で使用されています。;S10:VBAProject;S72:C:\Program Files\Common Files\Microsoft Shared\VBA\VBA7.1\1041\VbLR6.chm;N412e881400000000]]
30	E:\oxi-vba-public-artifacts-20261115\types5043\owned-data-5044-0-30.bin	A2[A9[N403e000000000000;N0000000000000000;N4000000000000000;S51:seed_before|parent|leaf|seed_after|operation_after|;A3[T;S0:;S0:];S9:Variant();N40c0060000000000;A5[N0000000000000000;S0:;S0:;S0:;N0000000000000000];N4054c00000000000];A5[N407ca00000000000;S109:Visual Basic でサポートされていないオートメーションが変数で使用されています。;S10:VBAProject;S72:C:\Program Files\Common Files\Microsoft Shared\VBA\VBA7.1\1041\VbLR6.chm;N412e881400000000]]
31	E:\oxi-vba-public-artifacts-20261115\types5043\owned-data-5044-0-31.bin	A2[A9[N403f000000000000;N0000000000000000;N4008000000000000;S51:seed_before|parent|leaf|seed_after|operation_after|;A3[F;S4:$C$1;S1:0];S9:Variant();N40c0060000000000;A5[N0000000000000000;S0:;S0:;S0:;N0000000000000000];N4054c00000000000];A5[N0000000000000000;S0:;S0:;S0:;N0000000000000000]]
42	E:\oxi-vba-public-artifacts-20261115\types5043\owned-data-5044-0-42.bin	A2[A9[N4045000000000000;N3ff0000000000000;N0000000000000000;S51:seed_before|parent|leaf|seed_after|operation_after|;A3[T;S0:;S0:];S9:Variant();N40c0060000000000;A5[N0000000000000000;S0:;S0:;S0:;N0000000000000000];N4054c00000000000];A5[N407ca00000000000;S109:Visual Basic でサポートされていないオートメーションが変数で使用されています。;S10:VBAProject;S72:C:\Program Files\Common Files\Microsoft Shared\VBA\VBA7.1\1041\VbLR6.chm;N412e881400000000]]
43	E:\oxi-vba-public-artifacts-20261115\types5043\owned-data-5044-0-43.bin	A2[A9[N4045800000000000;N3ff0000000000000;N3ff0000000000000;S63:seed_before|parent|leaf|seed_after|parent|leaf|operation_after|;A3[T;S0:;S0:];S9:Variant();N40c0060000000000;A5[N0000000000000000;S0:;S0:;S0:;N0000000000000000];N4054c00000000000];A5[N407ca00000000000;S109:Visual Basic でサポートされていないオートメーションが変数で使用されています。;S10:VBAProject;S72:C:\Program Files\Common Files\Microsoft Shared\VBA\VBA7.1\1041\VbLR6.chm;N412e881400000000]]
44	E:\oxi-vba-public-artifacts-20261115\types5043\owned-data-5044-0-44.bin	A2[A9[N4046000000000000;N3ff0000000000000;N4000000000000000;S51:seed_before|parent|leaf|seed_after|operation_after|;A3[T;S0:;S0:];S9:Variant();N40c0060000000000;A5[N0000000000000000;S0:;S0:;S0:;N0000000000000000];N4054c00000000000];A5[N407ca00000000000;S109:Visual Basic でサポートされていないオートメーションが変数で使用されています。;S10:VBAProject;S72:C:\Program Files\Common Files\Microsoft Shared\VBA\VBA7.1\1041\VbLR6.chm;N412e881400000000]]
45	E:\oxi-vba-public-artifacts-20261115\types5043\owned-data-5044-0-45.bin	A2[A9[N4046800000000000;N3ff0000000000000;N4008000000000000;S51:seed_before|parent|leaf|seed_after|operation_after|;A3[F;S4:$C$1;S1:0];S9:Variant();N40c0060000000000;A5[N0000000000000000;S0:;S0:;S0:;N0000000000000000];N4054c00000000000];A5[N0000000000000000;S0:;S0:;S0:;N0000000000000000]]"####.lines() {
        let mut fields = row.splitn(3, '\t');
        let kind = fields.next().unwrap().parse::<i64>().unwrap();
        let pathname = fields.next().unwrap();
        let expected = fields.next().unwrap();
        let mut book = len_automation_workbook();
        let mut host = WorkbookHost::new(&mut book, 0).unwrap();
        host.error_language = ErrorLanguage::Japanese;
        let value = oxivba_core::runtime::Runtime::new(&module).with_classes(&classes)
            .with_project_name("VBAProject").with_host(&mut host)
            .with_error_environment(ErrorEnvironment { language: ErrorLanguage::Japanese, help_file: r####"C:\Program Files\Common Files\Microsoft Shared\VBA\VBA7.1\1041\VbLR6.chm"####.into() })
            .call("Probe", vec![Value::Integer(kind), Value::String(pathname.into())]).unwrap();
        assert_eq!(len_cell_snapshot(&value), expected, "Native host consumption case={kind}");
    }
}
