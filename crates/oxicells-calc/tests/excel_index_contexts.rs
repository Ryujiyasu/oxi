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
    book.define_name("OriginArray", "={\"abc\"}").unwrap();
    book.define_name("OriginCell", "=S!$A$1").unwrap();
    book.define_name("NativeRows", "=S!$1:$3").unwrap();
    book.define_name("NativeArea", "=S!$A$1:$O$3").unwrap();
    book.set_value("S", "E1", Value::from_utf16(vec![0xD83D])).unwrap();
    book.set_value("S", "E2", Value::from_utf16(vec![0xDE00])).unwrap();
    book
}

fn scalar(value: &Value) -> String {
    match value {
        Value::Error(error) => format!("E:{}", match error {
            ExcelError::Null => 1, ExcelError::DivZero => 2, ExcelError::Value => 3,
            ExcelError::Ref => 4, ExcelError::Name => 5, ExcelError::Num => 6,
            ExcelError::NA => 7,
            // These observations use VBA Evaluate, whose host projection
            // maps the calculation-only error to the legacy VALUE code.
            ExcelError::Calc => 3,
            _ => 8,
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

fn grid_report(book: &Workbook, cols: u32) -> String {
    let values: Vec<_> = (0..3).flat_map(|row| (0..cols).map(move |col| (row, col))).map(|(row, col)| {
        let address = oxicells_calc::reference::CellRef::new(12 + col, row).to_a1();
        scalar(&book.value("S", &address))
    }).collect();
    format!("A2:1:3:1:{cols}|{}", values.join("|"))
}


#[test]
fn native_index_distinguishes_legacy_dynamic_and_cse_cells() {
    let mut failures = Vec::new(); let mut count = 0;
    for line in include_str!("../../../tests/fixtures/vba/worksheet_index_contexts.tsv").lines().filter(|line| !line.is_empty() && !line.starts_with('#')) {
        let f: Vec<_> = line.split('\t').collect(); let mut book = book(f[0].parse().unwrap());
        if let Some(initial) = f[3].strip_prefix("cycle") {
            let initial: u32 = initial.parse().unwrap();
            book.set_value("S", "D1", Value::Number(initial as f64)).unwrap();
            book.set_formula("S", "M1", f[2]).unwrap();
            book.set_dynamic("S", "M1").unwrap();
            let mut stages = Vec::new();
            for stage in ["entry", "repeat", "changed", "normal", "full"] {
                if stage == "changed" { book.set_value("S", "D1", Value::Number((1 - initial) as f64)).unwrap(); }
                book.recalculate();
                stages.push(format!("{stage}:{}", grid_report(&book, 15)));
            }
            let actual = stages.join("|");
            if actual != f[4] { failures.push(format!("{line}\tactual:{actual}")); }
            count += 1;
            continue;
        }
        if f[3] == "cse" {
            for row in 0..3 { book.set_array_member("S", &format!("M{}", row + 1), f[2], (0, row)).unwrap(); }
        } else {
            book.set_formula("S", "M1", f[2]).unwrap();
            if f[3] == "dynamic" { book.set_dynamic("S", "M1").unwrap(); }
        }
        book.recalculate();
        let cols = if f[1] == "CellIndexGridReport" { 3 } else { 1 };
        let actual = grid_report(&book, cols);
        if actual != f[4] { failures.push(format!("{line}\tactual:{actual}")); }
        count += 1;
    }
    assert_eq!(count, 228); assert!(failures.is_empty(), "{}differences:\n{}", failures.len(), failures.join("\n"));
    let mut circular = book(2);
    circular.set_formula("S", "M1", "INDEX(1:3,1,13)").unwrap();
    assert_eq!(circular.recalculate().circular, vec!["S!M1"]);
}
