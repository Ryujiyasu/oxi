// This Source Code Form is subject to the terms of the Mozilla Public
// License, v. 2.0. If a copy of the MPL was not distributed with this
// file, You can obtain one at https://mozilla.org/MPL/2.0/.
use oxicells_calc::{ExcelError, Value, Workbook};

fn book(version: u8) -> Workbook {
    let mut book = Workbook::new();
    book.add_sheet("S");
    book.set_dbcs_language(false);
    book.set_compatibility_version(version);
    let values = [
        Value::from_utf16(vec![97, 0xD83D, 0xDE00, 98]),
        Value::from_utf16(vec![0xD83D]),
        Value::from_utf16(vec![0xDE00]),
        Value::text("Tokyo東京"),
        Value::Number(12345.0),
        Value::Logical(true),
        Value::Error(ExcelError::DivZero),
        Value::Blank,
    ];
    for (i, value) in values.into_iter().enumerate() {
        book.set_value("S", &format!("A{}", i + 1), value).unwrap();
    }
    for (col, values) in [("B", vec![-1, 0, 1, 2]), ("D", vec![1, 2, 3])] {
        for (i, n) in values.into_iter().enumerate() {
            book.set_value("S", &format!("{col}{}", i + 1), Value::Number(n as f64)).unwrap();
        }
    }
    for (i, value) in [Value::text("a"), Value::from_utf16(vec![0xDE00]), Value::text("~?")].into_iter().enumerate() {
        book.set_value("S", &format!("C{}", i + 1), value).unwrap();
    }
    for (address, n) in [("E12", 1), ("F12", 2), ("G12", 3)] {
        book.set_value("S", address, Value::Number(n as f64)).unwrap();
    }
    book
}

fn scalar(value: &Value) -> String {
    match value {
        Value::Error(error) => format!("E:{}", match error {
            ExcelError::Null => 1, ExcelError::DivZero => 2, ExcelError::Value => 3,
            ExcelError::Ref => 4, ExcelError::Name => 5, ExcelError::Num => 6,
            ExcelError::NA => 7, _ => 8,
        }),
        Value::Number(n) => format!("5:{n}"),
        Value::Logical(b) => format!("11:{}", if *b {"True"} else {"False"}),
        Value::Blank => "0:".into(),
        Value::Text(_) | Value::Utf16(_) => {
            let units = value.to_utf16().unwrap();
            format!("S:{}{}", units.len(), units.iter().map(|u| format!(":{}", *u as i16)).collect::<String>())
        }
    }
}

#[test]
fn native_text_arrays_preserve_shapes_units_coercion_and_element_errors() {
    let mut count = 0;
    let mut failures = Vec::new();
    for line in include_str!("../../../tests/fixtures/vba/worksheet_byte_arrays.tsv").lines().filter(|l| !l.starts_with('#') && !l.is_empty()) {
        let fields: Vec<_> = line.split('\t').collect();
        let book = book(fields[0].parse().unwrap());
        let (width, height, values, array) = book.evaluate_block("S", fields[1]).unwrap();
        let prefix = if !array { "V".to_string() }
            else if height == 1 {format!("A1:1:{width}")}
            else {format!("A2:1:{height}:1:{width}")};
        let actual = format!("{prefix}|{}", values.iter().map(scalar).collect::<Vec<_>>().join("|"));
        if actual != fields[2] {
            failures.push(format!("{line}\tactual:{actual}"));
        }
        count += 1;
    }
    assert_eq!(count, 272);
    assert!(failures.is_empty(), "{} differences:\n{}", failures.len(), failures.iter().take(20).cloned().collect::<Vec<_>>().join("\n"));
}
