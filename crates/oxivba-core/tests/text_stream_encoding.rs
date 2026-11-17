// This Source Code Form is subject to the terms of the Mozilla Public
// License, v. 2.0. If a copy of the MPL was not distributed with this
// file, You can obtain one at https://mozilla.org/MPL/2.0/.

use oxivba_core::{parse_module, AnsiCodec, Host, ObjectRef, Runtime, Value};

struct JapaneseHost;

impl Host for JapaneseHost {
    fn ansi_codec(&self) -> Option<AnsiCodec> {
        AnsiCodec::from_code_page(932)
    }
    fn get(&mut self, _: &ObjectRef, _: &str) -> Result<Option<Value>, String> {
        Ok(None)
    }
    fn set(&mut self, _: &ObjectRef, _: &str, _: Value) -> Result<bool, String> {
        Ok(false)
    }
    fn call(&mut self, _: Option<&ObjectRef>, _: &str, _: &[Value]) -> Result<Option<Value>, String> {
        Ok(None)
    }
}

fn run(body: &str, seed: &[u8]) -> (Value, Vec<u8>) {
    let source = format!(
        "Option Explicit\nFunction Probe() As Variant\nDim fs As Object, s As Object, a As Variant, b As Variant, c As Variant\n\
         Set fs = CreateObject(\"Scripting.FileSystemObject\")\n{body}\nEnd Function\n"
    );
    let module = parse_module(&source).unwrap();
    let mut host = JapaneseHost;
    let mut runtime = Runtime::new(&module).with_host(&mut host);
    runtime.file_system_mut().replace_bytes("test.bin", seed.to_vec()).unwrap();
    let result = runtime.call("Probe", vec![]).unwrap();
    (result, runtime.file_system_mut().bytes("test.bin").unwrap().to_vec())
}

fn values(value: Value) -> Vec<Value> {
    match value {
        Value::Array(array) => array.values,
        other => panic!("expected an array, got {other:?}"),
    }
}

#[test]
fn ansi_output_uses_the_host_code_page_and_rejects_lossy_writes_atomically() {
    let (result, bytes) = run(
        "Set s = fs.CreateTextFile(\"test.bin\", True, False)\n\
         s.Write \"A\" & ChrW(&H65E5) & ChrW(&H672C) & \"B\"\n\
         On Error Resume Next\n\
         s.Write ChrW(&HD83D) & ChrW(&HDE00)\n\
         a = Err.Number\n\
         Probe = Array(a, s.Line, s.Column)\n\
         s.Close",
        b"old",
    );
    assert_eq!(values(result), vec![Value::Integer(5), Value::Integer(1), Value::Integer(5)]);
    assert_eq!(bytes, [0x41, 0x93, 0xfa, 0x96, 0x7b, 0x42]);
}

#[test]
fn unicode_output_preserves_utf16_units_and_counts_surrogate_halves() {
    let (result, bytes) = run(
        "Set s = fs.CreateTextFile(\"test.bin\", True, True)\n\
         s.Write \"A\" & ChrW(&HD83D) & ChrW(&HDE00) & ChrW(&HD83D)\n\
         Probe = Array(s.Line, s.Column)\n\
         s.Close",
        b"old",
    );
    assert_eq!(values(result), vec![Value::Integer(1), Value::Integer(5)]);
    assert_eq!(bytes, [0xff, 0xfe, 0x41, 0, 0x3d, 0xd8, 0, 0xde, 0x3d, 0xd8]);
}

#[test]
fn ansi_reads_count_characters_and_the_last_line_keeps_its_line_number() {
    let (result, _) = run(
        "Set s = fs.OpenTextFile(\"test.bin\", 1, False, 0)\n\
         a = s.Read(3)\n\
         b = s.ReadLine\n\
         c = s.ReadLine\n\
         Probe = Array(a, b, c, s.Line, s.Column, s.AtEndOfStream)\n\
         s.Close",
        &[0x41, 0x93, 0xfa, 0x96, 0x7b, 0x42, 13, 10, 0x43, 0x44],
    );
    assert_eq!(values(result), vec![
        Value::from_utf16(vec![0x41, 0x65e5, 0x672c]), Value::String("B".into()),
        Value::String("CD".into()), Value::Integer(2), Value::Integer(3), Value::Boolean(true),
    ]);
}

#[test]
fn unicode_reads_split_surrogates_and_ignore_an_incomplete_final_unit() {
    let (result, _) = run(
        "Set s = fs.OpenTextFile(\"test.bin\", 1, False, -1)\n\
         a = s.Read(1)\n\
         b = s.Read(1)\n\
         c = s.AtEndOfStream\n\
         On Error Resume Next\n\
         s.Read(1)\n\
         Probe = Array(a, b, c, Err.Number, s.Line, s.Column)\n\
         s.Close",
        &[0xff, 0xfe, 0x3d, 0xd8, 0, 0xde, 0x41],
    );
    assert_eq!(values(result), vec![
        Value::from_utf16(vec![0xd83d]), Value::from_utf16(vec![0xde00]),
        Value::Boolean(true), Value::Integer(62), Value::Integer(1), Value::Integer(3),
    ]);
}

#[test]
fn carriage_return_alone_is_data_and_readline_only_consumes_lf_or_crlf() {
    let (result, _) = run(
        "Set s = fs.OpenTextFile(\"test.bin\", 1)\n\
         a = s.Read(1)\n\
         b = s.AtEndOfLine\n\
         c = s.ReadLine\n\
         Probe = Array(a, b, c, s.Line, s.Column, s.AtEndOfStream)\n\
         s.Close",
        b"A\rB",
    );
    assert_eq!(values(result), vec![
        Value::String("A".into()), Value::Boolean(false), Value::String("\rB".into()),
        Value::Integer(1), Value::Integer(4), Value::Boolean(true),
    ]);
}

#[test]
fn unicode_append_keeps_existing_bytes_and_emits_one_bom_for_an_empty_file() {
    for seed in [b"".as_slice(), &[0xff, 0xfe, 0x41, 0]] {
        let (result, bytes) = run(
            "Set s = fs.OpenTextFile(\"test.bin\", 8, True, -1)\n\
             a = s.Column\n\
             s.Write \"B\"\n\
             Probe = Array(a, s.Column)\n\
             s.Close",
            seed,
        );
        assert_eq!(values(result), vec![Value::Integer(1), Value::Integer(2)]);
        let mut expected = if seed.is_empty() { vec![0xff, 0xfe] } else { seed.to_vec() };
        expected.extend_from_slice(&[0x42, 0]);
        assert_eq!(bytes, expected);
    }
}

#[test]
fn ansi_input_drops_an_incomplete_final_dbcs_character() {
    let (result, _) = run(
        "Set s = fs.OpenTextFile(\"test.bin\", 1, False, 0)\n\
         a = s.ReadAll\n\
         Probe = Array(a, s.Column, s.AtEndOfStream)\n\
         s.Close",
        &[0x41, 0x9e],
    );
    assert_eq!(values(result), vec![Value::String("A".into()), Value::Integer(2), Value::Boolean(true)]);
}

#[test]
fn explicit_unicode_input_honors_a_big_endian_bom() {
    let (result, _) = run(
        "Set s = fs.OpenTextFile(\"test.bin\", 1, False, -1)\n\
         a = s.ReadAll\n\
         Probe = Array(a, s.Column, s.AtEndOfStream)\n\
         s.Close",
        &[0xfe, 0xff, 0, 0x41, 0, 0x42],
    );
    assert_eq!(values(result), vec![Value::String("AB".into()), Value::Integer(3), Value::Boolean(true)]);
}

#[test]
fn default_input_detects_byte_order_without_a_bom_and_rejects_invalid_unicode() {
    let cases = [
        (vec![0xe5, 0x65, 0x2c, 0x67], vec![0x65e5, 0x672c]),
        (vec![0x65, 0xe5, 0x67, 0x2c], vec![0x65e5, 0x672c]),
        (vec![0x41, 0, 0x42, 0], vec![0x41, 0x42]),
        (vec![0, 0x41, 0, 0x42], vec![0x41, 0x42]),
        (vec![0x41, 0x42], vec![0x41, 0x42]),
        (vec![0x41, 0, 0, 0, 0x42, 0], vec![0x41, 0, 0, 0, 0x42, 0]),
        (vec![0xfe, 0xff, 0, 0x41, 0, 0x42], vec![0xf8f2, 0xf8f3, 0, 0x41, 0, 0x42]),
    ];
    for (seed, expected) in cases {
        let (result, bytes) = run(
            "Set s = fs.OpenTextFile(\"test.bin\", 1, False, -2)\n\
             a = s.ReadAll\n\
             Probe = Array(a, s.Column, s.AtEndOfStream)\n\
             s.Close",
            &seed,
        );
        assert_eq!(values(result), vec![
            Value::from_utf16(expected.clone()), Value::Integer(expected.len() as i64 + 1), Value::Boolean(true),
        ]);
        assert_eq!(bytes, seed);
    }
}

#[test]
fn explicit_unicode_and_default_append_preserve_big_endian_order() {
    for format in [-1, -2] {
        let body = format!(
            "Set s = fs.OpenTextFile(\"test.bin\", 8, True, {format})\n\
             a = s.Column\n\
             s.Write \"B\"\n\
             b = s.Column\n\
             Probe = Array(a, b)\n\
             s.Close"
        );
        let (result, bytes) = run(&body, &[0x65, 0xe5, 0x67, 0x2c]);
        assert_eq!(values(result), vec![
            Value::Integer(1), Value::Integer(2),
        ]);
        assert_eq!(bytes, [0x65, 0xe5, 0x67, 0x2c, 0, 0x42]);
    }
}

#[test]
fn default_detection_checks_the_entire_file_length_after_its_initial_sample() {
    let mut seed = b"A\0".repeat(64);
    seed.extend_from_slice(b"\0\0A");
    let (result, bytes) = run(
        "Set s = fs.OpenTextFile(\"test.bin\", 1, False, -2)\n\
         a = s.ReadAll\n\
         Probe = Array(a, s.Column, s.AtEndOfStream)\n\
         s.Close",
        &seed,
    );
    assert_eq!(values(result), vec![
        Value::from_utf16(seed.iter().map(|&byte| u16::from(byte)).collect()),
        Value::Integer(seed.len() as i64 + 1), Value::Boolean(true),
    ]);
    assert_eq!(bytes, seed);
}

#[test]
fn ansi_append_column_counts_a_dbcs_group_separately_from_decoded_units() {
    for (seed, column) in [(vec![0x81, 0], 2), (vec![0x9e], 2), (vec![0x41, 0x9e], 3)] {
        let (result, bytes) = run(
            "Set s = fs.OpenTextFile(\"test.bin\", 8, False, 0)\n\
             a = s.Column\n\
             s.Write \"B\"\n\
             Probe = Array(a, s.Column)\n\
             s.Close",
            &seed,
        );
        assert_eq!(values(result), vec![Value::Integer(column), Value::Integer(column + 1)]);
        let mut expected = seed;
        expected.push(0x42);
        assert_eq!(bytes, expected);
    }
}

#[test]
fn opening_unicode_append_completes_an_odd_unit_before_any_text_write() {
    for seed in [vec![0x9e], vec![0x41], vec![0x41, 0x93, 0xfa], vec![0xff, 0xfe, 0x41, 0, 0x42], vec![0xfe, 0xff, 0, 0x41, 0]] {
        let (result, bytes) = run(
            "Set s = fs.OpenTextFile(\"test.bin\", 8, False, -1)\n\
             Open \"test.bin\" For Binary Access Read As #9\n\
             a = LOF(9)\n\
             Close #9\n\
             Probe = Array(a, s.Line, s.Column)\n\
             s.Close",
            &seed,
        );
        assert_eq!(values(result), vec![
            Value::Integer(seed.len() as i64 + 1), Value::Integer(1), Value::Integer(1),
        ]);
        let mut expected = seed;
        expected.push(0);
        assert_eq!(bytes, expected);
    }
}

#[test]
fn constructor_type_errors_and_invalid_modes_preserve_existing_file_bytes() {
    for (expression, number) in [
        ("fs.OpenTextFile(\"test.bin\", 2, True, Null)", 13),
        ("fs.OpenTextFile(\"test.bin\", 2, True, CVErr(13))", 13),
        ("fs.OpenTextFile(\"test.bin\", 2, True, 2147483648#)", 6),
        ("fs.OpenTextFile(\"test.bin\", 2147483648#, True, Null)", 6),
        ("fs.OpenTextFile(\"test.bin\", 0, True, Null)", 13),
        ("fs.OpenTextFile(\"test.bin\", Null, True, 2147483648#)", 13),
        ("fs.OpenTextFile(\"test.bin\", 1, Null, 2147483648#)", 13),
        ("fs.OpenTextFile(\"test.bin\", Empty, True, 0)", 5),
        ("fs.OpenTextFile(\"test.bin\", 0, True, 0)", 5),
        ("fs.CreateTextFile(\"test.bin\", True, \"bad\")", 13),
        ("fs.CreateTextFile(\"test.bin\", False, \"bad\")", 13),
    ] {
        let (result, bytes) = run(
            &format!("On Error Resume Next\nSet s = {expression}\nProbe = Err.Number"),
            b"original",
        );
        assert_eq!(result, Value::Integer(number), "{expression}");
        assert_eq!(bytes, b"original", "{expression}");
    }
}

#[test]
fn constructor_integer_arguments_use_ties_to_even_rounding() {
    for mode in ["1.5", "2.5", "\"1.5\""] {
        let (result, bytes) = run(
            &format!("Set s = fs.OpenTextFile(\"test.bin\", {mode}, True, 0)\nProbe = Array(s.Line, s.Column)\ns.Close"),
            b"original",
        );
        assert_eq!(values(result), vec![Value::Integer(1), Value::Integer(1)]);
        assert!(bytes.is_empty());
    }
    for (format, expected) in [("-0.6", vec![0xff, 0xfe]), ("-0.5", vec![]), ("-1.5", vec![])] {
        let (_, bytes) = run(
            &format!("Set s = fs.OpenTextFile(\"test.bin\", 2, True, {format})\ns.Close\nProbe = 0"),
            b"original",
        );
        assert_eq!(bytes, expected);
    }
}

#[test]
fn constructor_boolean_arguments_distinguish_empty_from_omission() {
    for (arguments, error, expected) in [
        ("Empty, False", 58, b"original".to_vec()),
        (", False", 0, vec![]),
        ("\"True\", \"False\"", 0, vec![]),
        ("True, \"True\"", 0, vec![0xff, 0xfe]),
    ] {
        let (result, bytes) = run(
            &format!("On Error Resume Next\nSet s = fs.CreateTextFile(\"test.bin\", {arguments})\na = Err.Number\nProbe = a\ns.Close"),
            b"original",
        );
        assert_eq!(result, Value::Integer(error));
        assert_eq!(bytes, expected);
    }
}
