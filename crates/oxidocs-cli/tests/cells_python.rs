// This Source Code Form is subject to the terms of the Mozilla Public
// License, v. 2.0. If a copy of the MPL was not distributed with this
// file, You can obtain one at https://mozilla.org/MPL/2.0/.

#![cfg(feature = "vba-execution")]

use std::{fs, path::Path, process::Command, time::SystemTime};

#[test]
fn source_code_page_also_preserves_native_document_identity() {
    let root = Path::new(env!("CARGO_MANIFEST_DIR")).join("../../tests/fixtures/vba");
    let cases: serde_json::Value =
        serde_json::from_slice(&fs::read(root.join("unicode_modules.json")).unwrap()).unwrap();
    for case in cases.as_array().unwrap() {
        let bytes = fs::read(root.join(case["package_fixture"].as_str().unwrap())).unwrap();
        let metadata: serde_json::Value = serde_json::from_str(
            &oxidocs_wasm::read_spreadsheet_vba_project_native(&bytes, Some(932)).unwrap(),
        )
        .unwrap();
        if case["name"] == "unicode_workbook" {
            // These JSON expectations describe source execution with explicit
            // document bindings. This saved package still names ThisWorkbook;
            // opening the actual file in Excel leaves the renamed class unbound
            // and Driver.Probe raises 429. Preserve that package distinction.
            let documents = metadata["documentModules"].as_array().unwrap();
            let unbound = documents
                .iter()
                .find(|d| d["name"] == "日本語台帳")
                .unwrap();
            assert_eq!(unbound["unbound"], true);
            assert!(unbound["codeName"].is_null());
            assert!(documents.iter().any(|d| d["name"] == "ThisWorkbook"
                && d["codeName"] == "ThisWorkbook"
                && d["unbound"] != true));
            continue;
        }
        for expected in case["document_modules"].as_array().unwrap() {
            let actual = metadata["documentModules"]
                .as_array()
                .unwrap()
                .iter()
                .find(|document| document["name"] == expected["name"])
                .unwrap_or_else(|| {
                    panic!(
                        "{}: document {} missing from {}",
                        case["name"], expected["name"], metadata["documentModules"]
                    )
                });
            for (key, value) in expected.as_object().unwrap() {
                assert_eq!(&actual[key], value, "{}: {key}: {actual}", case["name"]);
            }
            assert_ne!(actual["unbound"], true, "{}: {actual}", case["name"]);
        }
    }
}

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
            "Private Type FloatStorage\n payload As Double\nEnd Type\n",
            "Private Type IntegerStorage\n payload As LongLong\nEnd Type\n",
            "Private total As Long\n",
            "Public Function DoubleBits(ByVal item As Double) As LongLong\n",
            " Dim floating As FloatStorage, integral As IntegerStorage\n",
            " floating.payload = item\n LSet integral = floating\n DoubleBits = integral.payload\nEnd Function\n",
            "Public Function FromBits(ByVal encoded As LongLong) As Double\n",
            " Dim floating As FloatStorage, integral As IntegerStorage\n",
            " integral.payload = encoded\n LSet floating = integral\n FromBits = floating.payload\nEnd Function\n",
            "Public Function Step(ByVal delta As Long) As Long\n",
            " total = total + delta\n Range(\"A1\").Value = total\n Step = total\nEnd Function\n",
            "Public Function Echo(ByVal item As Variant) As Variant\n Echo = item\nEnd Function\n",
            "Public Function Environment() As Variant\n",
            " Dim n As Long, d As String, s As String, h As String, c As Long\n",
            " Dim encoded As String, bytes As String, i As Long\n",
            " On Error Resume Next\n Err.Raise 5\n",
            " n = Err.Number: d = Err.Description: s = Err.Source: h = Err.HelpFile: c = Err.HelpContext\n",
            " Err.Clear\n encoded = StrConv(ChrW(12354), vbFromUnicode)\n",
            " For i = 1 To LenB(encoded)\n bytes = bytes & Right$(\"0\" & Hex$(AscB(MidB$(encoded, i, 1))), 2)\n Next\n",
            " bytes = bytes & \"|\"\n encoded = StrConv(ChrW(12354), vbFromUnicode, 2048)\n",
            " For i = 1 To LenB(encoded)\n bytes = bytes & Right$(\"0\" & Hex$(AscB(MidB$(encoded, i, 1))), 2)\n Next\n",
            " Environment = Array(n, d, s, h, c, bytes, Err.Number)\nEnd Function\n",
            "Public Function InputHeader(ByVal mode As Long, ByVal count As Long) As Variant\n",
            " Dim value As Variant\n Open ThisWorkbook.FullName For Binary Access Read Shared As #1\n",
            " Select Case mode\n Case 0: value = InputB(count, 1)\n Case 1: value = InputB$(count, 1)\n",
            " Case 2: value = Input(count, 1)\n Case 3: value = Input$(count, 1)\n End Select\n",
            " InputHeader = Array(TypeName(value), VarType(value), LenB(value), Len(value), value, Seek(1), EOF(1), Loc(1))\n",
            " Close #1\nEnd Function\n",
            "Public Function ReadCell() As Variant\n ReadCell = Range(\"A1\").Value\nEnd Function\n",
            "Public Function FailAfterWrite() As Variant\n Range(\"A1\").Value = 9\n Err.Raise 5\nEnd Function\n",
            "Public Function Bounds(ByVal item As Variant) As Variant\n",
            " Bounds = Array(VarType(item), LBound(item, 1), UBound(item, 1), LBound(item, 2), UBound(item, 2), item(1, 1))\nEnd Function\n",
            "Public Function MakeBox() As Object\n Set MakeBox = New BridgeBox\nEnd Function\n",
            "Public Function CallBox(ByVal item As Object) As Long\n CallBox = item.Touch()\nEnd Function\n",
            "Public Sub Auto_Open()\n Range(\"A1\").Value = 999\nEnd Sub\n",
            "Public Sub ExternalCalls(ByVal target As String)\n",
            " Dim remote As Object\n Call Shell(\"never-run.exe\", 1)\n",
            " Set remote = CreateObject(\"MSXML2.XMLHTTP\")\n",
            " Set remote = CreateObject(target)\nEnd Sub\n",
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
        .env(
            "OXICELLS_TEST_VBA_FIXTURES",
            root.join("tests/fixtures/vba"),
        )
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
