// SPDX-License-Identifier: MIT OR Apache-2.0

#![cfg(feature = "suite")]

#[test]
fn native_compilation_json_preserves_byref_diagnostics_without_execution() {
    let source = concat!(
        "Attribute VB_Name = \"NativeCheck\"\n",
        "Sub Main()\nvalue = CLng(12)\nAcceptLong value\nEnd Sub\n",
        "Private Sub AcceptLong(ByRef value As Long)\nErr.Raise 5\nEnd Sub\n"
    );
    let read = |text: &str, options: &str| -> serde_json::Value {
        serde_json::from_str(&oxidocs_wasm::compile_spreadsheet_vba_native(text, options).unwrap())
            .unwrap()
    };
    let rejected = read(source, r#"{"errorLanguage":"en","currentYear":2026}"#);
    assert_eq!(rejected["compiled"], false);
    assert_eq!(rejected["error"]["kind"], "compile");
    assert_eq!(rejected["error"]["module"], "NativeCheck");
    assert_eq!(rejected["error"]["procedure"], "Main");
    assert_eq!(rejected["error"]["line"], 4);
    assert!(rejected["error"]["message"]
        .as_str()
        .unwrap()
        .contains("ByRef argument type mismatch"));
    let japanese = read(source, r#"{"errorLanguage":"ja","currentYear":2026}"#);
    assert_eq!(japanese["error"]["line"], rejected["error"]["line"]);
    assert_ne!(japanese["error"]["message"], rejected["error"]["message"]);
    // Parentheses allow conversion. Err.Raise would fail if compilation ran it.
    let accepted = source.replace("AcceptLong value", "AcceptLong (value)");
    assert_eq!(
        read(&accepted, "null"),
        serde_json::json!({"compiled":true})
    );
    assert_eq!(read(&accepted, "{}"), serde_json::json!({"compiled":true}));
}

#[test]
fn native_compilation_json_rejects_invalid_options() {
    let source = "Sub Main()\nEnd Sub\n";
    for options in [
        "",
        "[]",
        "true",
        r#"{"currentYear":2026.5}"#,
        r#"{"currentYear":99}"#,
        r#"{"currentYear":10000}"#,
        r#"{"errorLanguage":"unknown"}"#,
        r#"{"unknown":true}"#,
        r#"{"referenceTypes":[{"name":"Test","unknown":true}]}"#,
        r#"{"currentYear":2026,"currentYear":2027}"#,
    ] {
        assert!(
            oxidocs_wasm::compile_spreadsheet_vba_native(source, options).is_err(),
            "{options}"
        );
    }
}
