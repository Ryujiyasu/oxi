// This Source Code Form is subject to the terms of the Mozilla Public
// License, v. 2.0. If a copy of the MPL was not distributed with this
// file, You can obtain one at https://mozilla.org/MPL/2.0/.

use oxivba_core::{parse_module, Runtime, Value};

fn values(value: Value) -> Vec<Value> {
    match value {
        Value::Array(array) => array.values,
        other => panic!("expected an array, got {other:?}"),
    }
}

#[test]
fn sequential_writes_remain_buffered_across_sub_and_macro_returns() {
    let module = parse_module(
        r#"Private Sub StartWrite()
 Open "test.bin" For Output As #1
 Print #1, "AAA";
End Sub
Function Probe() As Variant
 StartWrite
 Probe=FileLen("test.bin")
End Function
Function Finish() As Variant
 Dim before As Long
 before=FileLen("test.bin")
 Close #1
 Finish=Array(before,FileLen("test.bin"))
End Function"#,
    )
    .unwrap();
    let mut runtime = Runtime::new(&module);
    runtime
        .file_system_mut()
        .replace_bytes("test.bin", b"OLD".to_vec())
        .unwrap();
    assert_eq!(runtime.call("Probe", vec![]).unwrap(), Value::Integer(0));
    assert_eq!(runtime.file_system_mut().bytes("test.bin").unwrap(), b"");
    let snapshot = runtime.file_system_mut().closed_snapshot();
    assert_eq!(snapshot.files[0].1, b"");
    assert_eq!(
        values(runtime.call("Finish", vec![]).unwrap()),
        vec![Value::Integer(0), Value::Integer(3)]
    );
    assert_eq!(runtime.file_system_mut().bytes("test.bin").unwrap(), b"AAA");
}

#[test]
fn lof_flushes_output_without_refreshing_the_cached_file_length() {
    for (mode, initial, final_length) in [("Output", 0, 3), ("Append", 3, 6)] {
        let source = format!(
            "Function Probe() As Variant\nDim a As Long,b As Long,c As Long\n\
             Open \"test.bin\" For {mode} As #1\nPrint #1,\"AAA\";\n\
             a=FileLen(\"test.bin\"):b=LOF(1):c=FileLen(\"test.bin\")\n\
             Probe=Array(a,b,c)\nEnd Function\n"
        );
        let module = parse_module(&source).unwrap();
        let mut runtime = Runtime::new(&module);
        runtime
            .file_system_mut()
            .replace_bytes("test.bin", b"OLD".to_vec())
            .unwrap();
        assert_eq!(
            values(runtime.call("Probe", vec![]).unwrap()),
            vec![
                Value::Integer(initial),
                Value::Integer(final_length),
                Value::Integer(initial)
            ]
        );
        assert_eq!(
            runtime.file_system_mut().bytes("test.bin").unwrap().len(),
            final_length as usize
        );
    }
}

#[test]
fn file_size_and_attributes_do_not_flush_numbered_output() {
    let module = parse_module(
        r#"Function Probe() As Variant
 Dim fs As Object,a As Variant,b As Variant
 Set fs=CreateObject("Scripting.FileSystemObject")
 Open "test.bin" For Output As #1
 Print #1,"AAA";
 a=fs.GetFile("test.bin").Size
 b=GetAttr("test.bin")
 Probe=Array(a,b,FileLen("test.bin"))
End Function"#,
    )
    .unwrap();
    let mut runtime = Runtime::new(&module);
    runtime
        .file_system_mut()
        .replace_bytes("test.bin", b"OLD".to_vec())
        .unwrap();
    assert_eq!(
        values(runtime.call("Probe", vec![]).unwrap()),
        vec![Value::Integer(0), Value::Integer(32), Value::Integer(0)]
    );
    assert_eq!(runtime.file_system_mut().bytes("test.bin").unwrap(), b"");
}

#[test]
fn binary_and_random_file_sizes_refresh_the_directory_entry() {
    for (mode, record, payload) in [("Binary", "", "AAAAA"), ("Random", " Len=16", "AAA")] {
        let source = format!(
            "Function Probe() As Variant\nDim fs As Object,a As Variant,b As Variant,c As Variant\n\
             Set fs=CreateObject(\"Scripting.FileSystemObject\")\n\
             Open \"test.bin\" For {mode} As #1{record}\nPut #1,,\"{payload}\"\n\
             a=FileLen(\"test.bin\"):b=fs.GetFile(\"test.bin\").Size:c=FileLen(\"test.bin\")\n\
             Close #1\nProbe=Array(a,b,c,FileLen(\"test.bin\"))\nEnd Function\n"
        );
        let module = parse_module(&source).unwrap();
        let mut runtime = Runtime::new(&module);
        runtime
            .file_system_mut()
            .replace_bytes("test.bin", b"OLD".to_vec())
            .unwrap();
        assert_eq!(
            values(runtime.call("Probe", vec![]).unwrap()),
            vec![
                Value::Integer(3),
                Value::Integer(5),
                Value::Integer(5),
                Value::Integer(5)
            ]
        );
    }
}

#[test]
fn fso_writes_are_physical_before_filelen_refreshes() {
    let module = parse_module(
        r#"Function Probe() As Variant
 Dim fs As Object,s As Object,a As Variant
 Set fs=CreateObject("Scripting.FileSystemObject")
 Set s=fs.OpenTextFile("test.bin",8,True,0)
 s.Write "AAA"
 a=FileLen("test.bin")
 Probe=Array(a,fs.GetFile("test.bin").Size,FileLen("test.bin"))
End Function"#,
    )
    .unwrap();
    let mut runtime = Runtime::new(&module);
    runtime
        .file_system_mut()
        .replace_bytes("test.bin", b"OLD".to_vec())
        .unwrap();
    assert_eq!(
        values(runtime.call("Probe", vec![]).unwrap()),
        vec![Value::Integer(3), Value::Integer(6), Value::Integer(6)]
    );
    assert_eq!(
        runtime.file_system_mut().bytes("test.bin").unwrap(),
        b"OLDAAA"
    );
}

#[test]
fn unicode_bom_does_not_change_the_cached_creation_length() {
    let module = parse_module(
        r#"Function Probe() As Variant
 Dim fs As Object,s As Object,a As Variant,b As Variant,c As Variant
 Set fs=CreateObject("Scripting.FileSystemObject")
 Set s=fs.CreateTextFile("test.bin",True,True)
 a=FileLen("test.bin")
 s.Write "AAA"
 b=FileLen("test.bin")
 c=GetAttr("test.bin")
 Probe=Array(a,b,FileLen("test.bin"))
End Function"#,
    )
    .unwrap();
    let mut runtime = Runtime::new(&module);
    runtime
        .file_system_mut()
        .replace_bytes("test.bin", b"OLD".to_vec())
        .unwrap();
    assert_eq!(
        values(runtime.call("Probe", vec![]).unwrap()),
        vec![Value::Integer(0), Value::Integer(0), Value::Integer(8)]
    );
    assert_eq!(
        runtime.file_system_mut().bytes("test.bin").unwrap(),
        &[0xff, 0xfe, 65, 0, 65, 0, 65, 0]
    );
}

#[test]
fn sequential_buffer_boundaries_expose_only_written_bytes() {
    for (mode, initial) in [("Output", 0), ("Append", 3)] {
        for (count, written) in [(511, 0), (512, 0), (513, 512), (1024, 1024)] {
            let source = format!(
                "Function Probe() As Variant\nOpen \"test.bin\" For {mode} As #1\n\
                 Print #1,String$({count},\"A\");\nProbe=FileLen(\"test.bin\")\nEnd Function\n"
            );
            let module = parse_module(&source).unwrap();
            let mut runtime = Runtime::new(&module);
            runtime
                .file_system_mut()
                .replace_bytes("test.bin", b"OLD".to_vec())
                .unwrap();
            assert_eq!(
                runtime.call("Probe", vec![]).unwrap(),
                Value::Integer(initial)
            );
            assert_eq!(
                runtime.file_system_mut().bytes("test.bin").unwrap().len(),
                initial as usize + written
            );
            assert_eq!(
                runtime.file_system_mut().fresh_length("test.bin").unwrap(),
                initial as usize + written
            );
            assert_eq!(
                runtime.file_system_mut().query_length("test.bin").unwrap(),
                initial as usize + written
            );
            runtime.file_system_mut().close(&[1]);
            assert_eq!(
                runtime.file_system_mut().query_length("test.bin").unwrap(),
                initial as usize + count
            );
        }
    }
}

#[test]
fn closing_a_reader_publishes_growth_but_not_an_unextended_write_time() {
    for reader in [
        "Open \"test.bin\" For Binary Access Read As #2",
        "Set reader=fs.OpenTextFile(\"test.bin\",1,False,0)",
    ] {
        for count in [0, 3, 5] {
            let close = if reader.starts_with("Open ") {
                "Close #2"
            } else {
                "reader.Close"
            };
            let source = format!(
                "Function Probe() As Variant\n\
                 Dim fs As Object,reader As Object,a As Variant,b As Variant,c As Variant,d As Variant\n\
                 Set fs=CreateObject(\"Scripting.FileSystemObject\")\n\
                 Open \"test.bin\" For Binary As #1\n{reader}\n\
                 Put #1,,String$({count},\"A\")\n\
                 a=FileLen(\"test.bin\"):b=(CDbl(FileDateTime(\"test.bin\"))=43890.50283564815)\n\
                 {close}\n\
                 c=FileLen(\"test.bin\"):d=(CDbl(FileDateTime(\"test.bin\"))=43890.50283564815)\n\
                 Close #1\n\
                 Probe=Array(a,b,c,d,FileLen(\"test.bin\"),(CDbl(FileDateTime(\"test.bin\"))=43890.50283564815))\n\
                 End Function\n"
            );
            let module = parse_module(&source).unwrap();
            let mut runtime = Runtime::new(&module).with_current_time(46343.5);
            runtime
                .file_system_mut()
                .replace_bytes("test.bin", b"OLD".to_vec())
                .unwrap();
            runtime
                .file_system_mut()
                .set_modified_time(r"C:\test.bin", 43890.50283564815)
                .unwrap();
            let length = count.max(3);
            assert_eq!(
                values(runtime.call("Probe", vec![]).unwrap()),
                vec![
                    Value::Integer(3),
                    Value::Boolean(true),
                    Value::Integer(length),
                    Value::Boolean(count <= 3),
                    Value::Integer(length),
                    Value::Boolean(count == 0),
                ],
                "reader={reader}, count={count}"
            );
        }
    }
}

#[test]
fn fresh_file_metadata_keeps_the_entry_time_until_the_length_changes() {
    for count in [3, 5] {
        let source = format!(
            "Function Probe() As Variant\n\
             Dim fs As Object,a As Variant,b As Variant,c As Variant\n\
             Set fs=CreateObject(\"Scripting.FileSystemObject\")\n\
             Open \"test.bin\" For Binary As #1\nPut #1,,String$({count},\"A\")\n\
             a=fs.GetFile(\"test.bin\").Size\n\
             b=(CDbl(fs.GetFile(\"test.bin\").DateLastModified)=43890.50283564815)\n\
             c=(CDbl(FileDateTime(\"test.bin\"))=43890.50283564815)\n\
             Close #1\n\
             Probe=Array(a,b,c,(CDbl(FileDateTime(\"test.bin\"))=43890.50283564815))\n\
             End Function\n"
        );
        let module = parse_module(&source).unwrap();
        let mut runtime = Runtime::new(&module).with_current_time(46343.5);
        runtime
            .file_system_mut()
            .replace_bytes("test.bin", b"OLD".to_vec())
            .unwrap();
        runtime
            .file_system_mut()
            .set_modified_time(r"C:\test.bin", 43890.50283564815)
            .unwrap();
        assert_eq!(
            values(runtime.call("Probe", vec![]).unwrap()),
            vec![
                Value::Integer(count),
                Value::Boolean(count == 3),
                Value::Boolean(count == 3),
                Value::Boolean(false)
            ]
        );
    }
}
