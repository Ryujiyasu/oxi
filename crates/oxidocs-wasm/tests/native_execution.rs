// SPDX-License-Identifier: MIT OR Apache-2.0

#![cfg(feature = "suite")]

use oxicells_core::ir::{CellValue, Workbook};
use oxidocs_wasm::{
    read_spreadsheet_vba_project_native, NativeVbaAction, NativeVbaCall, SpreadsheetVbaSession,
};
use oxivba_core::Value;

const SOURCE: &str = concat!(
    "Attribute VB_Name = \"NativeRun\"\n",
    "Private total As Long\n",
    "Public Function Step(ByVal delta As Long) As Long\n",
    " total = total + delta\n Range(\"A1\").Value = total\n Step = total\nEnd Function\n",
    "Public Function FailAfterWrite() As Variant\n",
    " Range(\"A1\").Value = 9\n Err.Raise 5\nEnd Function\n",
    "Public Function TypedValues() As Variant\n",
    " TypedValues = Array(CByte(255), CSng(1.25), CDate(1.5), ",
    "CLngLng(\"9007199254740993\"), ChrW(-10240))\nEnd Function\n"
);

const OBJECT_SOURCE: &str = concat!(
    "Attribute VB_Name = \"NativeObjects\"\n",
    "Public Function Cell() As Object\n Set Cell = Range(\"A1\")\nEnd Function\n"
);

fn workbook() -> Workbook {
    oxicells_core::parse_xlsx(include_bytes!("../../../tests/fixtures/basic_test.xlsx")).unwrap()
}

fn call(procedure: &str, args: Vec<Value>) -> NativeVbaCall<'_> {
    NativeVbaCall::new(SOURCE, procedure, args, 45_000.0, 1)
}

#[test]
fn native_execution_keeps_project_storage_and_workbook_edits() {
    let mut session = SpreadsheetVbaSession::default();
    let first = session
        .run_native(workbook(), call("NativeRun.Step", vec![Value::Int16(2)]))
        .unwrap();
    assert!(
        matches!(first.result(), Value::Integer(2)),
        "result {:?}; error {}",
        first.result(),
        serde_json::from_str::<serde_json::Value>(&first.to_json().unwrap()).unwrap()["error"]
    );
    assert!(matches!(
        first.workbook().sheets[0].rows[0].cells[0].value,
        CellValue::Number(2.0)
    ));
    let second = session
        .run_native(
            first.into_workbook(),
            call("NativeRun.Step", vec![Value::Integer(3)]),
        )
        .unwrap();
    assert!(matches!(second.result(), Value::Integer(5)));
    assert!(matches!(
        second.workbook().sheets[0].rows[0].cells[0].value,
        CellValue::Number(5.0)
    ));
}

#[test]
fn native_execution_preserves_vba_result_subtypes_and_utf16() {
    let mut session = SpreadsheetVbaSession::default();
    let outcome = session
        .run_native(workbook(), call("NativeRun.TypedValues", vec![]))
        .unwrap();
    let Value::Array(values) = outcome.result() else {
        panic!("expected the VBA array result");
    };
    assert_eq!(values.values.len(), 5);
    assert!(matches!(values.values[0], Value::Byte(255)));
    assert!(matches!(values.values[1], Value::Single(v) if v.to_bits() == 1.25f32.to_bits()));
    assert!(matches!(values.values[2], Value::Date(v) if v.to_bits() == 1.5f64.to_bits()));
    assert!(matches!(
        values.values[3],
        Value::LongLong(9_007_199_254_740_993)
    ));
    assert!(matches!(&values.values[4], Value::Utf16(v) if v == &[0xd800]));
}

#[test]
fn native_execution_invalid_inputs_keep_state_and_runtime_errors_keep_edits() {
    let mut session = SpreadsheetVbaSession::default();
    let first = session
        .run_native(workbook(), call("NativeRun.Step", vec![Value::Integer(2)]))
        .unwrap();
    let book = first.into_workbook();
    let mut invalid_clock = call("NativeRun.Step", vec![Value::Integer(100)]);
    invalid_clock.current_time = f64::NAN;
    assert!(session.run_native(book.clone(), invalid_clock).is_err());
    let mut invalid_options = call("NativeRun.Step", vec![Value::Integer(100)]);
    invalid_options.options = "[]";
    assert!(session.run_native(book.clone(), invalid_options).is_err());
    let next = session
        .run_native(book, call("NativeRun.Step", vec![Value::Integer(3)]))
        .unwrap();
    assert!(
        matches!(next.result(), Value::Integer(5)),
        "result {:?}; error {}",
        next.result(),
        serde_json::from_str::<serde_json::Value>(&next.to_json().unwrap()).unwrap()["error"]
    );
    let failed = session
        .run_native(
            next.into_workbook(),
            call("NativeRun.FailAfterWrite", vec![]),
        )
        .unwrap();
    assert!(matches!(failed.result(), Value::Empty));
    assert!(matches!(
        failed.workbook().sheets[0].rows[0].cells[0].value,
        CellValue::Number(9.0)
    ));
    let details: serde_json::Value = serde_json::from_str(&failed.to_json().unwrap()).unwrap();
    assert_eq!(details["error"]["kind"], "runtime");
    assert_eq!(details["workbook_changed"], true);
}

#[test]
fn native_execution_compilation_and_prepare_save_do_not_run_a_macro_or_reset_storage() {
    let mut session = SpreadsheetVbaSession::default();
    let first = session
        .run_native(workbook(), call("NativeRun.Step", vec![Value::Integer(2)]))
        .unwrap();
    let mut compile = call("NativeRun.Step", vec![]);
    compile.action = NativeVbaAction::CompileProject;
    let compiled = session.run_native(first.into_workbook(), compile).unwrap();
    let details: serde_json::Value = serde_json::from_str(&compiled.to_json().unwrap()).unwrap();
    assert_eq!(details["compilation"]["compiled"], true);
    assert!(details["error"].is_null());
    assert!(matches!(
        compiled.workbook().sheets[0].rows[0].cells[0].value,
        CellValue::Number(2.0)
    ));
    let mut prepare = call("NativeRun.Step", vec![]);
    prepare.action = NativeVbaAction::PrepareSave;
    let prepared = session
        .run_native(compiled.into_workbook(), prepare)
        .unwrap();
    let next = session
        .run_native(
            prepared.into_workbook(),
            call("NativeRun.Step", vec![Value::Integer(3)]),
        )
        .unwrap();
    assert!(matches!(next.result(), Value::Integer(5)));

    let invalid = concat!(
        "Public Sub Bump(ByRef value As Long)\n value = value + 1\nEnd Sub\n",
        "Public Sub Bad()\n Dim value As Variant\n",
        " Range(\"A1\").Value = 99\n Bump value\nEnd Sub\n"
    );
    let before = serde_json::to_value(next.workbook()).unwrap();
    let refused = SpreadsheetVbaSession::default()
        .run_native(
            next.into_workbook(),
            NativeVbaCall::new(invalid, "Bad", vec![], 45_000.0, 1),
        )
        .unwrap();
    let details: serde_json::Value = serde_json::from_str(&refused.to_json().unwrap()).unwrap();
    assert_eq!(details["error"]["kind"], "compile");
    assert_eq!(details["workbook_changed"], false);
    assert_eq!(serde_json::to_value(refused.workbook()).unwrap(), before);
}

#[test]
fn native_execution_retained_objects_are_session_owned_and_explicitly_released() {
    let mut session = SpreadsheetVbaSession::default();
    let object = session
        .run_native(
            workbook(),
            NativeVbaCall::new(OBJECT_SOURCE, "NativeObjects.Cell", vec![], 45_000.0, 1),
        )
        .unwrap();
    assert!(matches!(object.result(), Value::Object(_)));
    let details: serde_json::Value = serde_json::from_str(&object.to_json().unwrap()).unwrap();
    let reference = details["result"]["value"]["reference"]
        .as_str()
        .unwrap()
        .to_owned();
    let mut book = object.into_workbook();
    for (kind, args) in [(4, vec![Value::Integer(17)]), (2, vec![])] {
        let mut invoke = NativeVbaCall::new(OBJECT_SOURCE, "", args, 45_000.0, 1);
        invoke.action = NativeVbaAction::Invoke {
            reference: reference.clone(),
            member: "Value".into(),
            kind,
        };
        let result = session.run_native(book, invoke).unwrap();
        let details: serde_json::Value = serde_json::from_str(&result.to_json().unwrap()).unwrap();
        assert!(details["error"].is_null(), "{details}");
        if kind == 2 {
            assert!(matches!(result.result(), Value::Double(17.0)));
        }
        book = result.into_workbook();
    }
    let mut foreign = NativeVbaCall::new(OBJECT_SOURCE, "", vec![], 45_000.0, 1);
    foreign.action = NativeVbaAction::Invoke {
        reference: reference.clone(),
        member: "Value".into(),
        kind: 2,
    };
    assert!(SpreadsheetVbaSession::default()
        .run_native(book.clone(), foreign)
        .is_err());

    // Releasing uses the loaded project even when an editor has pending invalid source.
    let mut release = NativeVbaCall::new("not valid VBA", "", vec![], 45_000.0, 1);
    release.action = NativeVbaAction::ReleaseAll;
    let released = session.run_native(book, release).unwrap();
    let details: serde_json::Value = serde_json::from_str(&released.to_json().unwrap()).unwrap();
    assert!(details["error"].is_null(), "{details}");
    let mut stale = NativeVbaCall::new(OBJECT_SOURCE, "", vec![], 45_000.0, 1);
    stale.action = NativeVbaAction::Invoke {
        reference,
        member: "Value".into(),
        kind: 2,
    };
    assert!(session.run_native(released.into_workbook(), stale).is_err());
}

#[test]
fn native_execution_save_copy_reopens_with_the_macro_edits() {
    let bytes = include_bytes!("../../../tests/fixtures/vba/row_height_state.xlsm");
    let options = serde_json::json!({
        "applicationState": {"displayAlerts": false, "enableEvents": false},
        "fileContext": {"packageBytes": bytes.as_slice(),
            "fullName": "C:\\work\\input.xlsm", "folders": ["C:\\work"]}
    })
    .to_string();
    let mut session = SpreadsheetVbaSession::default();
    let mut write = call("NativeRun.Step", vec![Value::Integer(23)]);
    write.file_name = Some("input.xlsm".into());
    write.options = &options;
    let written = session
        .run_native(oxicells_core::parse_xlsx(bytes).unwrap(), write)
        .unwrap();
    assert!(matches!(written.result(), Value::Integer(23)));
    let mut save = call("", vec![]);
    save.file_name = Some("input.xlsm".into());
    save.action = NativeVbaAction::SaveWorkbook {
        copy_name: Some("C:\\work\\copy.xlsm".into()),
    };
    let saved = session.run_native(written.into_workbook(), save).unwrap();
    let details: serde_json::Value = serde_json::from_str(&saved.to_json().unwrap()).unwrap();
    assert!(details["error"].is_null(), "{details}");
    assert_eq!(details["save"]["completed"], true);
    assert_eq!(details["save"]["path"], "C:\\work\\copy.xlsm");
    let package: Vec<u8> =
        serde_json::from_value(details["save"]["package_bytes"].clone()).unwrap();
    let reopened = oxicells_core::parse_xlsx(&package).unwrap();
    assert!(matches!(
        reopened.sheets[0].rows[0].cells[0].value,
        CellValue::Number(23.0)
    ));
    assert!(matches!(
        saved.workbook().sheets[0].rows[0].cells[0].value,
        CellValue::Number(23.0)
    ));
}

#[test]
fn native_execution_typed_nonfinite_results_are_never_silently_json_null() {
    let source =
        "Public Function Echo(ByVal value As Variant) As Variant\n Echo = value\nEnd Function\n";
    let mut session = SpreadsheetVbaSession::default();
    let result = session
        .run_native(
            workbook(),
            NativeVbaCall::new(
                source,
                "Echo",
                vec![Value::Double(f64::INFINITY)],
                45_000.0,
                1,
            ),
        )
        .unwrap();
    assert!(matches!(result.result(), Value::Double(v) if *v == f64::INFINITY));
    assert!(result
        .to_json()
        .unwrap_err()
        .contains("non-finite VBA result"));
}

#[test]
fn native_project_reads_packaged_source_and_binds_code_names_after_a_sheet_rename() {
    let bytes = include_bytes!("../../../tests/fixtures/vba/row_height_state.xlsm");
    let mut info: serde_json::Value =
        serde_json::from_str(&read_spreadsheet_vba_project_native(bytes, Some(932)).unwrap())
            .unwrap();
    assert_eq!(info["projectName"], "VBAProject");
    let documents = info["documentModules"].as_array().unwrap();
    assert!(documents.iter().any(|doc| doc["name"] == "ThisWorkbook"
        && doc["baseType"] == "Excel.Workbook"
        && doc["codeName"] == "ThisWorkbook"));
    assert!(documents.iter().any(|doc| doc["name"] == "Sheet1"
        && doc["baseType"] == "Excel.Worksheet"
        && doc["codeName"] == "Sheet1"
        && doc["sheetIndex"] == 0));
    let mut source = info
        .as_object_mut()
        .unwrap()
        .remove("source")
        .unwrap()
        .as_str()
        .unwrap()
        .to_owned();
    assert!(source.contains("Attribute VB_Name = \"SaveProbe\""));
    let options = info.to_string();
    let mut session = SpreadsheetVbaSession::default();
    let mut call = NativeVbaCall::new(
        &source,
        "SaveProbe.HeightProbeIdentity",
        vec![],
        45_000.0,
        1,
    );
    call.options = &options;
    let identity = session
        .run_native(oxicells_core::parse_xlsx(bytes).unwrap(), call)
        .unwrap();
    assert!(matches!(identity.result(), Value::Integer(2_026_102_025)));

    let mut book = identity.into_workbook();
    book.sheets[0].name = "Renamed".into();
    source.push_str("\r\nPublic Function WriteBoundSheet() As Long\r\n Sheet1.Range(\"A1\").Value = 23\r\n WriteBoundSheet = Sheet1.Index\r\nEnd Function\r\n");
    let mut call = NativeVbaCall::new(&source, "SaveProbe.WriteBoundSheet", vec![], 45_000.0, 1);
    call.options = &options;
    let written = session.run_native(book, call).unwrap();
    assert!(
        matches!(written.result(), Value::Integer(1)),
        "{}",
        written.to_json().unwrap()
    );
    assert_eq!(written.workbook().sheets[0].name, "Renamed");
    assert!(matches!(
        written.workbook().sheets[0].rows[0].cells[0].value,
        CellValue::Number(23.0)
    ));
}

#[test]
fn native_project_without_vba_reports_absence_without_javascript() {
    let bytes = include_bytes!("../../../tests/fixtures/basic_test.xlsx");
    let info: serde_json::Value =
        serde_json::from_str(&read_spreadsheet_vba_project_native(bytes, None).unwrap()).unwrap();
    assert!(info["source"].is_null());
    assert!(info["projectName"].is_null());
    assert!(info["documentModules"].is_array());
}

#[test]
fn native_project_rejects_invalid_packages_and_unsupported_source_encodings() {
    let error = read_spreadsheet_vba_project_native(b"not an OOXML archive", None).unwrap_err();
    assert!(!error.is_empty());
    let bytes = include_bytes!("../../../tests/fixtures/vba/row_height_state.xlsm");
    let error = read_spreadsheet_vba_project_native(bytes, Some(65_535)).unwrap_err();
    assert!(!error.is_empty());
}

const CODEC_SOURCE: &str = concat!(
    "Attribute VB_Name = \"NativeCodec\"\n",
    "Public Function Echo(ByVal value As Variant) As Variant\n Echo = value\nEnd Function\n",
    "Public Function TypedArray() As Variant\n Dim data(-2 To -1) As Integer\n",
    " data(-2) = 7\n data(-1) = 9\n TypedArray = data\nEnd Function\n",
    "Public Function InspectArray(ByVal data As Variant) As Variant\n",
    " InspectArray = Array(VarType(data), LBound(data), UBound(data), data(-2), data(-1))\nEnd Function\n",
    "Public Function Pointer(ByVal text As String) As LongLong\n Pointer = StrPtr(text)\nEnd Function\n"
);

#[test]
fn native_values_json_round_trips_subtypes_and_array_shape() {
    let source = format!("{SOURCE}{CODEC_SOURCE}");
    let mut session = SpreadsheetVbaSession::default();
    let result = session
        .run_native(
            workbook(),
            NativeVbaCall::new(&source, "NativeRun.TypedValues", vec![], 45_000.0, 1),
        )
        .unwrap();
    let json: serde_json::Value = serde_json::from_str(&result.to_typed_json().unwrap()).unwrap();
    let values = &json["result"]["value"]["values"];
    assert_eq!(values[0], serde_json::json!({"type":"byte","value":255}));
    assert_eq!(values[1], serde_json::json!({"type":"single","value":1.25}));
    assert_eq!(values[2], serde_json::json!({"type":"date","value":1.5}));
    assert_eq!(
        values[3],
        serde_json::json!({"type":"long_long","value":"9007199254740993"})
    );
    assert_eq!(
        values[4],
        serde_json::json!({"type":"utf16","value":[0xd800]})
    );
    let args = session
        .decode_native_arguments(&serde_json::json!([json["result"]]).to_string())
        .unwrap();
    let echoed = session
        .run_native(
            result.into_workbook(),
            NativeVbaCall::new(&source, "NativeCodec.Echo", args, 45_000.0, 1),
        )
        .unwrap();
    let again: serde_json::Value = serde_json::from_str(&echoed.to_typed_json().unwrap()).unwrap();
    assert_eq!(again["result"], json["result"]);

    let typed = session
        .run_native(
            echoed.into_workbook(),
            NativeVbaCall::new(&source, "NativeCodec.TypedArray", vec![], 45_000.0, 1),
        )
        .unwrap();
    let json: serde_json::Value = serde_json::from_str(&typed.to_typed_json().unwrap()).unwrap();
    assert_eq!(
        json["result"]["value"]["element_default"],
        serde_json::json!({"type":"integer","value":0})
    );
    assert_eq!(
        json["result"]["value"]["dimensions"],
        serde_json::json!([{"lower_bound":-2,"length":2}])
    );
    let args = session
        .decode_native_arguments(&serde_json::json!([json["result"]]).to_string())
        .unwrap();
    let inspected = session
        .run_native(
            typed.into_workbook(),
            NativeVbaCall::new(&source, "NativeCodec.InspectArray", args, 45_000.0, 1),
        )
        .unwrap();
    let json: serde_json::Value = serde_json::from_str(&inspected.to_json().unwrap()).unwrap();
    let fields: Vec<_> = json["result"]["value"]["values"]
        .as_array()
        .unwrap()
        .iter()
        .map(|v| v["value"].clone())
        .collect();
    assert_eq!(
        fields,
        serde_json::json!([8194, -2, -1, 7, 9])
            .as_array()
            .unwrap()
            .clone()
    );
}

#[test]
fn native_values_json_preserves_exact_numbers_and_bstr_contents() {
    let descriptors = vec![
        serde_json::json!({"type":"currency","value":"-922337203685477.5808"}),
        serde_json::json!({"type":"decimal","value":{"negative":false,"magnitude":"79228162514264337593543950335","scale":28}}),
        serde_json::json!({"type":"double","value":{"bits":"7ff8000000000001"}}),
        serde_json::json!({"type":"single","value":{"bits":"7f800000"}}),
        serde_json::json!({"type":"utf16","value":[0xd800]}),
        serde_json::json!({"type":"bstr","value":[65,0,66]}),
    ];
    let mut session = SpreadsheetVbaSession::default();
    let mut book = workbook();
    for descriptor in descriptors {
        let args = session
            .decode_native_arguments(&serde_json::json!([descriptor]).to_string())
            .unwrap();
        let result = session
            .run_native(
                book,
                NativeVbaCall::new(CODEC_SOURCE, "NativeCodec.Echo", args, 45_000.0, 1),
            )
            .unwrap();
        let json: serde_json::Value =
            serde_json::from_str(&result.to_typed_json().unwrap()).unwrap();
        assert!(json["error"].is_null(), "{json}");
        assert_eq!(json["result"], descriptor);
        book = result.into_workbook();
    }
    let args = session
        .decode_native_arguments(r#"[{"type":"null_string"}]"#)
        .unwrap();
    let null = session
        .run_native(
            book,
            NativeVbaCall::new(CODEC_SOURCE, "NativeCodec.Pointer", args, 45_000.0, 1),
        )
        .unwrap();
    assert!(matches!(null.result(), Value::LongLong(0)));
}

#[test]
fn native_values_json_objects_release_user_classes_in_the_owning_session() {
    const SOURCE: &str = concat!(
        "Attribute VB_Name = \"NativeLife\"\n",
        "Public Function MakeBox() As Object\n Set MakeBox = New Box\nEnd Function\n",
        "Public Function CallBox(ByVal item As Object) As Long\n CallBox = item.Touch()\nEnd Function\n",
        "VERSION 1.0 CLASS\nAttribute VB_Name = \"Box\"\n",
        "Public Function Touch() As Long\n Range(\"A1\").Value = 17\n Touch = 17\nEnd Function\n",
        "Private Sub Class_Terminate()\n ThisWorkbook.Worksheets(1).Range(\"A1\").Value = 41\nEnd Sub\n"
    );
    let mut session = SpreadsheetVbaSession::default();
    let object = session
        .run_native(
            workbook(),
            NativeVbaCall::new(SOURCE, "NativeLife.MakeBox", vec![], 45_000.0, 1),
        )
        .unwrap();
    let json: serde_json::Value = serde_json::from_str(&object.to_typed_json().unwrap()).unwrap();
    let reference = json["result"]["value"]["reference"]
        .as_str()
        .unwrap()
        .to_owned();
    let descriptor = serde_json::json!([json["result"]]).to_string();
    assert!(SpreadsheetVbaSession::default()
        .decode_native_arguments(&descriptor)
        .is_err());
    let book = object.into_workbook();
    let args = session.decode_native_arguments(&descriptor).unwrap();
    let touched = session
        .run_native(
            book,
            NativeVbaCall::new(SOURCE, "NativeLife.CallBox", args, 45_000.0, 1),
        )
        .unwrap();
    assert!(matches!(touched.result(), Value::Integer(17)));
    let mut release = NativeVbaCall::new(SOURCE, "", vec![], 45_000.0, 1);
    release.action = NativeVbaAction::Release {
        references: vec![reference],
    };
    let released = session
        .run_native(touched.into_workbook(), release)
        .unwrap();
    let json: serde_json::Value = serde_json::from_str(&released.to_json().unwrap()).unwrap();
    assert!(json["error"].is_null(), "{json}");
    assert!(matches!(
        released.workbook().sheets[0].rows[0].cells[0].value,
        CellValue::Number(41.0)
    ));
    assert!(session.decode_native_arguments(&descriptor).is_err());
}

#[test]
fn native_values_json_rejects_invalid_arguments_without_resetting_storage() {
    let mut session = SpreadsheetVbaSession::default();
    let first = session
        .run_native(workbook(), call("NativeRun.Step", vec![Value::Integer(2)]))
        .unwrap();
    for invalid in [
        r#"{}"#,
        r#"[{"type":"long","value":1,"extra":0}]"#,
        r#"[{"type":"empty","value":null}]"#,
        r#"[{"type":"long","value":2147483648}]"#,
        r#"[{"type":"long_long","value":1}]"#,
        r#"[{"type":"double","value":{"bits":"not hex"}}]"#,
        r#"[{"type":"decimal","value":{"negative":false,"magnitude":"1","scale":29}}]"#,
        r#"[{"type":"array","value":{"dimensions":[{"lower_bound":0,"length":2}],"element_default":{"type":"empty"},"resizable":true,"values":[]}}]"#,
    ] {
        assert!(
            session.decode_native_arguments(invalid).is_err(),
            "{invalid}"
        );
    }
    let next = session
        .run_native(
            first.into_workbook(),
            call("NativeRun.Step", vec![Value::Integer(3)]),
        )
        .unwrap();
    assert!(matches!(next.result(), Value::Integer(5)));
}

#[test]
fn native_save_as_converts_the_package_and_keeps_the_live_project() {
    let bytes = include_bytes!("../../../tests/fixtures/vba/row_height_state.xlsm");
    let options = serde_json::json!({
        "applicationState":{"displayAlerts":false,"enableEvents":false},
        "fileContext":{"packageBytes":bytes.as_slice(),"fullName":"C:\\work\\input.xlsm","folders":["C:\\work"]}
    }).to_string();
    let mut session = SpreadsheetVbaSession::default();
    let mut write = call("NativeRun.Step", vec![Value::Integer(23)]);
    write.options = &options;
    write.file_name = Some("input.xlsm".into());
    let written = session
        .run_native(oxicells_core::parse_xlsx(bytes).unwrap(), write)
        .unwrap();
    assert!(matches!(written.result(), Value::Integer(23)));
    let mut save = call("", vec![]);
    save.action = NativeVbaAction::SaveAs {
        file_name: "C:\\work\\result.xlsx".into(),
        format: 51,
    };
    let saved = session.run_native(written.into_workbook(), save).unwrap();
    let details: serde_json::Value = serde_json::from_str(&saved.to_json().unwrap()).unwrap();
    assert!(details["error"].is_null(), "{details}");
    assert_eq!(details["save"]["completed"], true);
    let package = saved.saved_package_bytes().unwrap();
    let project: serde_json::Value =
        serde_json::from_str(&read_spreadsheet_vba_project_native(package, None).unwrap()).unwrap();
    assert!(project["source"].is_null());
    let reopened = oxicells_core::parse_xlsx(package).unwrap();
    assert!(matches!(
        reopened.sheets[0].rows[0].cells[0].value,
        CellValue::Number(23.0)
    ));
    let next = session
        .run_native(
            saved.into_workbook(),
            call("NativeRun.Step", vec![Value::Integer(1)]),
        )
        .unwrap();
    assert!(matches!(next.result(), Value::Integer(24)));
}

#[test]
fn native_window_collection_calls_preserve_complete_excel_observations() {
    let fixture = native_window_fixture();
    let group = fixture["groups"]
        .as_array()
        .unwrap()
        .iter()
        .find(|group| group["id"] == "binding_and_coercion")
        .unwrap();
    assert_eq!(group["fields"].as_array().unwrap().len(), 20);
    assert_eq!(group["cases"].as_array().unwrap().len(), 112);
    replay_native_window_group(group);
}

#[test]
fn native_window_lifecycle_and_state_assignments_preserve_complete_excel_observations() {
    let fixture = native_window_fixture();
    let groups = fixture["groups"].as_array().unwrap();
    assert_eq!(groups.len(), 19);
    assert_eq!(
        groups
            .iter()
            .map(|group| group["cases"].as_array().unwrap().len())
            .sum::<usize>(),
        2698
    );
    for group in groups
        .iter()
        .take(17)
        .filter(|group| group["id"] != "binding_and_coercion")
    {
        replay_native_window_group(group);
    }
}

#[test]
fn native_window_frame_state_transitions_match_excel() {
    let fixture = native_window_fixture();
    let groups = fixture["groups"].as_array().unwrap();
    assert_eq!(groups.len(), 19);
    assert_eq!(groups[17]["id"], "window_frame_display_transitions");
    assert_eq!(groups[18]["id"], "window_cross_book_display_transitions");
    assert_eq!(groups[17]["cases"].as_array().unwrap().len(), 6);
    assert_eq!(groups[18]["cases"].as_array().unwrap().len(), 4);
    for group in &groups[17..] {
        replay_native_window_group(group);
    }
}

fn native_window_fixture() -> serde_json::Value {
    serde_json::from_str(include_str!(
        "../../../tests/fixtures/vba/window_collections_native.json"
    ))
    .unwrap()
}

#[test]
fn native_json_preserves_double_bits() {
    let source = concat!(
        "Attribute VB_Name = \"NativeFloatBoundary\"\n",
        "Public Function Probe() As Double\n",
        " Probe = 0.9999997615814208984375\nEnd Function\n"
    );
    let mut session = SpreadsheetVbaSession::default();
    let outcome = session
        .run_native(
            workbook(),
            NativeVbaCall::new(source, "NativeFloatBoundary.Probe", vec![], 46_000.0, 1),
        )
        .unwrap();
    let expected = 0x3fef_ffff_8000_0000;
    let Value::Double(raw) = outcome.result() else {
        panic!("expected a Double result");
    };
    assert_eq!(raw.to_bits(), expected, "VBA literal before JSON");
    for wire in [outcome.to_json().unwrap(), outcome.to_typed_json().unwrap()] {
        let decoded: serde_json::Value = serde_json::from_str(&wire).unwrap();
        assert_eq!(
            decoded["result"]["value"].as_f64().unwrap().to_bits(),
            expected,
            "outgoing JSON must retain the raw VBA Double"
        );
    }
    let arguments = session
        .decode_native_arguments(r#"[{"type":"double","value":0.9999997615814209}]"#)
        .unwrap();
    let Value::Double(decoded) = arguments[0] else {
        panic!("expected a Double argument");
    };
    assert_eq!(decoded.to_bits(), expected, "incoming JSON Double");
}

fn replay_native_window_group(group: &serde_json::Value) {
    // Exercise the public options decoder as well as execution. The native
    // observations include Japanese Err descriptions and a host help-file path;
    // replaying them with the API's English defaults would change the input.
    let hex = group["carrier_bytes_hex"].as_str().unwrap();
    let carrier: Vec<u8> = (0..hex.len())
        .step_by(2)
        .map(|at| u8::from_str_radix(&hex[at..at + 2], 16).unwrap())
        .collect();
    let file_name = group["file_name"].as_str().unwrap();
    let options = serde_json::json!({
        "errorLanguage": "ja-JP",
        "vbaHelpFile": group["help_file"],
        "projectName": "VBAProject",
        "applicationState": {
            "visible": group["initial_application_visible"].as_bool().unwrap_or(false),
            "enableEvents": false, "displayAlerts": false,
            "standardFontStartup": {"profileUtf16": group["startup_profile_utf16"]}
        },
        "fileContext": {
            "packageBytes": carrier, "fullName": format!(r"C:\work\{file_name}"),
            "readOnly": true, "folders": [r"C:\work"]
        }
    })
    .to_string();
    let workbook = oxicells_core::parse_xlsx_preserving_values(&carrier).unwrap();
    let fields = group["fields"].as_array().unwrap();
    let cases = group["cases"].as_array().unwrap();
    for reverse in [false, true] {
        for position in 0..cases.len() {
            let case = &cases[if reverse {
                cases.len() - 1 - position
            } else {
                position
            }];
            let mode = case["mode"].as_i64().unwrap();
            let binding = case["binding"].as_i64().unwrap();
            let context = format!(
                "{} reverse={reverse}, mode={mode}, binding={binding}",
                group["id"]
            );
            let mut args = vec![Value::Integer(mode), Value::Integer(binding)];
            if let Some(bits) = case["input_operand_bits"].as_str() {
                args.push(Value::Double(f64::from_bits(
                    u64::from_str_radix(bits, 16).unwrap(),
                )));
            }
            let mut call = NativeVbaCall::new(
                group["program"].as_str().unwrap(),
                group["procedure"].as_str().unwrap_or("Probe"),
                args,
                46_000.0,
                1,
            );
            call.file_name = Some(file_name.into());
            call.options = &options;
            let outcome = SpreadsheetVbaSession::default()
                .run_native(workbook.clone(), call)
                .unwrap_or_else(|error| panic!("{context}: {error}"));
            let details: serde_json::Value =
                serde_json::from_str(&outcome.to_json().unwrap()).unwrap();
            assert!(
                details["error"].is_null(),
                "{context}: {}",
                details["error"]
            );
            if let Some(expected) = group.get("expected_application_visible") {
                assert_eq!(
                    details["application_state"]["visible"], *expected,
                    "{context}"
                );
            }
            if !group["mutates_caller_inputs"].as_bool().unwrap_or(false) {
                assert_eq!(details["workbook_changed"], false, "{context}");
            }
            // Empty file lists are omitted by the public outcome serializer.
            assert!(details.get("files").is_none(), "{context}");
            assert!(details.get("deleted_files").is_none(), "{context}");
            let actual = details["result"]["value"]["values"].as_array().unwrap();
            let expected = case["expected"].as_array().unwrap();
            assert_eq!(actual.len(), fields.len(), "{context}");
            assert_eq!(expected.len(), fields.len(), "{context}");
            for (at, ((field, actual), expected)) in
                fields.iter().zip(actual).zip(expected).enumerate()
            {
                if at == 8 {
                    if let Some(bits) = case["expected_operand_bits"].as_str() {
                        assert_eq!(
                            actual["value"].as_f64().unwrap().to_bits(),
                            u64::from_str_radix(bits, 16).unwrap(),
                            "{context}, field={field}: exact operand bits"
                        );
                        continue;
                    }
                }
                let mut value = actual["value"].clone();
                if group["id"] == "property_types"
                    && at < 39
                    && at % 3 == 0
                    && case["expected"][at + 1] == 10
                    && actual["type"] == "error"
                {
                    // Native COM captures VT_ERROR as a signed SCODE; the
                    // public API carries its VBA error number. Compare the
                    // exact mapping while retaining the native fixture value.
                    let code = value.as_u64().unwrap();
                    assert!(code <= u64::from(u16::MAX));
                    value =
                        serde_json::Value::from(i64::from((0x800a_0000_u32 | code as u32) as i32));
                }
                assert_eq!(&value, expected, "{context}, field={field}");
            }
        }
    }
}
