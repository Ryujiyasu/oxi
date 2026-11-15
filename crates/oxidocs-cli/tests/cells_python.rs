// This Source Code Form is subject to the terms of the Mozilla Public
// License, v. 2.0. If a copy of the MPL was not distributed with this
// file, You can obtain one at https://mozilla.org/MPL/2.0/.

#![cfg(feature = "vba-execution")]

use std::{fs, path::Path, process::Command, time::SystemTime};

/// Exercises Python, JSONL, the real native engine and persisted OOXML together.
#[test]
fn python_runs_the_native_workbook_contract() {
    let root = Path::new(env!("CARGO_MANIFEST_DIR")).join("../..");
    let bytes = fs::read(root.join("tests/fixtures/vba/row_height_state.xlsm")).unwrap();
    let metadata: serde_json::Value = serde_json::from_str(
        &oxidocs_wasm::read_spreadsheet_vba_project_native(&bytes, None).unwrap(),
    )
    .unwrap();
    let source = format!(
        "{}\n{}",
        metadata["source"].as_str().unwrap(),
        concat!(
            "Attribute VB_Name = \"PythonBridge\"\n",
            "Private total As Long\n",
            "Public Function Step(ByVal delta As Long) As Long\n",
            " total = total + delta\n Range(\"A1\").Value = total\n Step = total\nEnd Function\n",
            "Public Function Echo(ByVal item As Variant) As Variant\n Echo = item\nEnd Function\n",
            "Public Function ReadCell() As Variant\n ReadCell = Range(\"A1\").Value\nEnd Function\n",
            "Public Function FailAfterWrite() As Variant\n Range(\"A1\").Value = 9\n Err.Raise 5\nEnd Function\n",
            "Public Function Bounds(ByVal item As Variant) As Variant\n",
            " Bounds = Array(VarType(item), LBound(item, 1), UBound(item, 1), LBound(item, 2), UBound(item, 2), item(1, 1))\nEnd Function\n",
            "Public Function MakeBox() As Object\n Set MakeBox = New BridgeBox\nEnd Function\n",
            "Public Function CallBox(ByVal item As Object) As Long\n CallBox = item.Touch()\nEnd Function\n",
            "Public Sub Auto_Open()\n Range(\"A1\").Value = 999\nEnd Sub\n",
            "VERSION 1.0 CLASS\nAttribute VB_Name = \"BridgeBox\"\n",
            "Public Function Touch() As Long\n Range(\"A1\").Value = 17\n Touch = 17\nEnd Function\n",
            "Private Sub Class_Terminate()\n ThisWorkbook.Worksheets(1).Range(\"A1\").Value = 41\nEnd Sub\n"
        )
    );
    let package = oxidocs_wasm::save_spreadsheet_vba_source(&bytes, &source, None).unwrap();
    let nonce = SystemTime::now()
        .duration_since(SystemTime::UNIX_EPOCH)
        .unwrap()
        .as_nanos();
    let temp = std::env::temp_dir().join(format!("oxicells-python-{}-{nonce}", std::process::id()));
    fs::create_dir(&temp).unwrap();
    let fixture = temp.join("bridge.xlsm");
    fs::write(&fixture, package).unwrap();
    let mut python = Command::new(std::env::var("PYTHON").unwrap_or_else(|_| "python".into()));
    python.arg(root.join("crates/oxicells-python/tests/test_execution.py"));
    if std::env::var("OXICELLS_TEST_INSTALLED").as_deref() == Ok("1") {
        // Exercise discovery of the engine and facade in the installed wheel.
        python
            .env_remove("OXICELLS_BINARY")
            .env_remove("PYTHONPATH");
    } else {
        python
            .env("OXICELLS_BINARY", env!("CARGO_BIN_EXE_oxicells"))
            .env("PYTHONPATH", root.join("crates/oxicells-python/python"));
    }
    let result = python
        .env("OXICELLS_TEST_FIXTURE", &fixture)
        .env("PYTHONDONTWRITEBYTECODE", "1")
        .status();
    // This directory is created exclusively by this test; never remove caller paths.
    let resolved = temp.canonicalize().unwrap();
    assert_eq!(
        resolved.parent(),
        Some(std::env::temp_dir().canonicalize().unwrap().as_path())
    );
    fs::remove_dir_all(&resolved).unwrap();
    assert!(result
        .expect("Python is required for the binding contract test")
        .success());
}
