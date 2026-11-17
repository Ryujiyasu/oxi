// SPDX-License-Identifier: MIT OR Apache-2.0

use super::*;

const SOURCE: &str = r#"
Attribute VB_Name = "BufferedFiles"
Public Function BeginWrite(ByVal appendMode As Boolean, ByVal amount As Long) As Long
 If appendMode Then
  Open "C:\work\data.bin" For Append As #1
 Else
  Open "C:\work\data.bin" For Output As #1
 End If
 Print #1, String$(amount,"A");
 BeginWrite=FileLen("C:\work\data.bin")
End Function
Public Function Size() As Long
 Size=FileLen("C:\work\data.bin")
End Function
Public Sub Finish()
 Close #1
End Sub
Public Sub BinaryWrite()
 Dim octet As Byte
 octet=90
 Open "C:\work\binary.bin" For Binary As #2
 Put #2,,octet
End Sub
"#;

fn workbook() -> Workbook {
    oxicells_core::parse_xlsx(include_bytes!("../../../tests/fixtures/basic_test.xlsx")).unwrap()
}

fn invoke(
    session: &mut SpreadsheetVbaSession,
    procedure: &str,
    args: Vec<Value>,
    options: &serde_json::Value,
) -> Value {
    let options = options.to_string();
    let mut call = NativeVbaCall::new(SOURCE, procedure, args, 46000.0, 1);
    call.options = &options;
    let out = session.run_native(workbook(), call).unwrap();
    assert!(out.outcome.error.is_none());
    out.result().clone()
}

fn initial() -> serde_json::Value {
    serde_json::json!({"fileContext": {
        "fullName": "C:\\work\\book.xlsm",
        "folders": ["C:\\work"],
        "files": [{"path": "C:\\work\\data.bin", "bytes": [79,76,68]}],
        "modifiedTimes": [{"path": "C:\\work\\data.bin", "value": 43890.50283564815}]
    }})
}

fn resume() -> serde_json::Value {
    serde_json::json!({"fileContext": {"fullName": "C:\\work\\book.xlsm"}})
}

fn close_snapshot(session: &SpreadsheetVbaSession) -> Option<serde_json::Value> {
    session
        .state
        .as_ref()
        .and_then(|state| state.host.file_io.snapshot_after_close())
        .map(|state| serde_json::to_value(state).unwrap()["fileSystemState"].clone())
}

fn bytes(state: &serde_json::Value) -> Vec<u8> {
    state["files"]
        .as_array()
        .unwrap()
        .iter()
        .find(|file| {
            file["path"]
                .as_str()
                .unwrap()
                .eq_ignore_ascii_case("C:\\work\\data.bin")
        })
        .unwrap()["bytes"]
        .as_array()
        .unwrap()
        .iter()
        .map(|value| value.as_u64().unwrap() as u8)
        .collect()
}

#[test]
fn buffered_close_snapshot_preserves_output_without_flushing_the_live_project() {
    for append in [false, true] {
        for amount in [3, 513] {
            let mut session = SpreadsheetVbaSession::new();
            assert!(close_snapshot(&session).is_none());
            let original_size = if append { 3 } else { 0 };
            assert_eq!(
                invoke(
                    &mut session,
                    "BeginWrite",
                    vec![Value::Boolean(append), Value::Integer(amount)],
                    &initial()
                ),
                Value::Integer(original_size)
            );
            let closed = close_snapshot(&session).unwrap();
            assert_eq!(close_snapshot(&session).unwrap(), closed);
            assert_eq!(
                invoke(&mut session, "Size", vec![], &resume()),
                Value::Integer(original_size)
            );
            let physical =
                serde_json::to_value(session.state.as_ref().unwrap().host.file_io.snapshot())
                    .unwrap();
            assert_ne!(bytes(&physical), bytes(&closed));
            let mut expected = if append { b"OLD".to_vec() } else { Vec::new() };
            expected.extend(std::iter::repeat_n(b'A', amount as usize));
            assert_eq!(bytes(&closed), expected);

            let mut restored = SpreadsheetVbaSession::new();
            let mut options = resume();
            options["fileSystemState"] = closed;
            assert_eq!(
                invoke(&mut restored, "Size", vec![], &options),
                Value::Integer(expected.len() as i64)
            );
            assert!(close_snapshot(&restored).is_none());
            // The original session still owns its channel and pending output.
            invoke(&mut session, "Finish", vec![], &resume());
            assert_eq!(
                invoke(&mut session, "Size", vec![], &resume()),
                Value::Integer(expected.len() as i64)
            );
        }
    }
}

#[test]
fn buffered_close_snapshot_disappears_after_close_and_empty_output() {
    let mut session = SpreadsheetVbaSession::new();
    invoke(
        &mut session,
        "BeginWrite",
        vec![Value::Boolean(false), Value::Integer(3)],
        &initial(),
    );
    assert_eq!(bytes(&close_snapshot(&session).unwrap()), b"AAA");
    invoke(&mut session, "Finish", vec![], &resume());
    assert!(close_snapshot(&session).is_none());
    invoke(
        &mut session,
        "BeginWrite",
        vec![Value::Boolean(false), Value::Integer(0)],
        &resume(),
    );
    assert!(close_snapshot(&session).is_none());
    let physical =
        serde_json::to_value(session.state.as_ref().unwrap().host.file_io.snapshot()).unwrap();
    assert!(bytes(&physical).is_empty());
}

#[test]
fn buffered_close_snapshot_dates_only_files_with_unflushed_output() {
    let mut session = SpreadsheetVbaSession::new();
    invoke(
        &mut session,
        "BeginWrite",
        vec![Value::Boolean(false), Value::Integer(3)],
        &initial(),
    );
    invoke(&mut session, "BinaryWrite", vec![], &resume());
    let descriptor = serde_json::to_value(
        session
            .state
            .as_ref()
            .unwrap()
            .host
            .file_io
            .snapshot_after_close()
            .unwrap(),
    )
    .unwrap();
    assert_eq!(
        descriptor["modifiedFiles"],
        serde_json::json!(["C:\\work\\data.bin"])
    );
    let binary = descriptor["fileSystemState"]["files"]
        .as_array()
        .unwrap()
        .iter()
        .find(|file| {
            file["path"]
                .as_str()
                .unwrap()
                .eq_ignore_ascii_case("C:\\work\\binary.bin")
        })
        .unwrap();
    assert_eq!(binary["bytes"], serde_json::json!([90]));
}
