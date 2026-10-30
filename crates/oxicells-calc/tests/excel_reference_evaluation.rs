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


fn cases(method: &str) -> Vec<Vec<&'static str>> {
    include_str!("../../../tests/fixtures/vba/worksheet_reference_evaluation.tsv").lines()
        .filter(|line| !line.is_empty() && !line.starts_with('#'))
        .map(|line| line.split('\t').collect::<Vec<_>>()).filter(|fields| fields[1] == method).collect()
}

fn report(answer: oxicells_calc::functions::Arg) -> String {
    use oxicells_calc::functions::Arg;
    match answer {
        Arg::Value(value) => format!("V|{}", scalar(&value)),
        Arg::Range(block) => {
            let reference = block.cell_reference();
            let prefix = match reference {
                Some((_, range)) if !range.is_single() => format!("A2:1:{}:1:{}", block.height, block.width),
                Some(_) => "V".to_string(),
                None if block.height == 1 => format!("A1:1:{}", block.width),
                None => format!("A2:1:{}:1:{}", block.height, block.width),
            };
            format!("{prefix}|{}", block.cells.iter().map(scalar).collect::<Vec<_>>().join("|"))
        }
    }
}

#[test]
fn native_type_tests_distinguish_singleton_arrays_and_references() {
    let rows = cases("ByteArrayCase"); assert_eq!(rows.len(), 88);
    for f in rows { assert_eq!(report(book(f[0].parse().unwrap()).evaluate_argument("S", f[2]).unwrap()), f[4], "{f:?}"); }
}

#[test]
fn native_selection_values_include_zero_negative_and_array_indexes() {
    let rows = cases("SelectionValue"); assert_eq!(rows.len(), 56);
    for f in rows { assert_eq!(report(book(f[0].parse().unwrap()).evaluate_argument("S", f[2]).unwrap()), f[4], "{f:?}"); }
}

#[test]
fn native_selection_identity_survives_array_lookups_and_index_arguments() {
    use oxicells_calc::functions::Arg;
    let rows = cases("OriginShape"); assert_eq!(rows.len(), 56);
    for f in rows {
        let answer = book(f[0].parse().unwrap()).evaluate_argument("S", f[2]).unwrap();
        let actual = match answer {
            Arg::Range(block) if block.cell_reference().is_some() => "Range:True",
            Arg::Range(_) => "Variant():False",
            Arg::Value(Value::Error(_)) => "Error:False",
            _ => panic!("unexpected {f:?}"),
        };
        assert_eq!(actual, f[4], "{f:?}");
    }
}

#[test]
fn native_reference_selectors_evaluate_once_and_do_not_retry_failures() {
    use oxicells_calc::{set_user_function_hook, UserArg};
    use oxicells_calc::functions::Arg;
    use std::{cell::Cell, rc::Rc};
    struct Reset;
    impl Drop for Reset { fn drop(&mut self) { set_user_function_hook(None); } }
    let _reset = Reset;
    let calls = Rc::new(Cell::new(0)); let counted = calls.clone();
    set_user_function_hook(Some(Box::new(move |name, args, _| {
        counted.set(counted.get() + 1);
        let value = if name.eq_ignore_ascii_case("STEPPINGTICK") { Value::Number(if counted.get() == 1 { 1.0 } else { 5.0 }) }
            else { match args.first() { Some(UserArg::Value(value)) => value.clone(), _ => Value::Error(ExcelError::Value) } };
        Some(Arg::Value(value))
    })));
    let rows = cases("ReferenceTrace"); assert_eq!(rows.len(), 12);
    for f in rows {
        calls.set(0);
        let actual = report(book(f[0].parse().unwrap()).evaluate_argument("S", f[2]).unwrap());
        assert_eq!(format!("0|{actual}|calls:{}", calls.get()), f[4], "{f:?}");
    }
    // Native UdfStages establishes one callback per Calculate, after the
    // distinct callback caused by Formula2 assignment. These are core
    // evaluation checks, not claims that VBA Evaluate's replay is fixed.
    for (expr, expected) in [("SelectionTick(1)+SelectionTick(1)", 2), ("INDEX(D1:D3,SelectionTick(5))", 1), ("ROW(INDEX(D1:D3,SelectionTick(1)))", 1), ("TAKE(D1:D3,SelectionTick(1))", 1), ("XLOOKUP(SelectionTick(1),D1:D3,A1:A3)", 1)] {
        calls.set(0); let _ = book(1).evaluate_argument("S", expr).unwrap(); assert_eq!(calls.get(), expected, "{expr}");
    }
    for expr in ["INDEX(D1:D3,SelectionTick(1))", "INDEX(D1:D3,SelectionTick(5))"] {
        calls.set(0); let _ = book(1).evaluate_reference("S", expr); assert_eq!(calls.get(), 1, "{expr}");
    }
}
