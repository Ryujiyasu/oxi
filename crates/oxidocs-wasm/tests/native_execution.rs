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
