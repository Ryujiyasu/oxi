// This Source Code Form is subject to the terms of the Mozilla Public
// License, v. 2.0. If a copy of the MPL was not distributed with this
// file, You can obtain one at https://mozilla.org/MPL/2.0/.
use oxicells_calc::{ExcelError, Value, Workbook};
fn value(text: &str) -> Value {
    let mut fields = text.split(':');
    let length: usize = fields.next().unwrap().parse().unwrap();
    let units: Vec<_> = fields.map(|s| s.parse::<i16>().unwrap() as u16).collect();
    assert_eq!(units.len(), length);
    if units.is_empty() {
        Value::Blank
    } else {
        Value::from_utf16(units)
    }
}
fn report(value: Value) -> String {
    match value {
        Value::Error(error) => format!(
            "E:{}",
            match error {
                ExcelError::Null => 1,
                ExcelError::DivZero => 2,
                ExcelError::Value => 3,
                ExcelError::Ref => 4,
                ExcelError::Name => 5,
                ExcelError::Num => 6,
                ExcelError::NA => 7,
                _ => 8,
            }
        ),
        Value::Number(n) => format!("5:{n}"),
        Value::Logical(b) => format!("11:{}", if b { "True" } else { "False" }),
        Value::Text(_) | Value::Utf16(_) => {
            let units = value.to_utf16().unwrap();
            format!(
                "S:{}{}",
                units.len(),
                units
                    .iter()
                    .map(|u| format!(":{}", *u as i16))
                    .collect::<String>()
            )
        }
        Value::Blank => "0:".into(),
    }
}
fn verify(fixture: &str, expected: usize) {
    let mut count = 0;
    let mut failures = Vec::new();
    for line in fixture.lines().filter(|l| !l.starts_with('#')) {
        let fields: Vec<_> = line.split('\t').collect();
        let mut book = Workbook::new();
        book.add_sheet("S");
        book.set_dbcs_language(false);
        book.set_compatibility_version(fields[0].parse().unwrap());
        for (address, units) in [("J20", fields[1]), ("K20", "1:-10179"), ("L20", "1:-8704")] {
            book.set_value("S", address, value(units)).unwrap();
        }
        let actual = report(book.evaluate("S", fields[2]).unwrap());
        if actual != fields[3] {
            failures.push(format!("{line}\tactual:{actual}"));
        }
        count += 1;
    }
    assert_eq!(count, expected);
    assert!(
        failures.is_empty(),
        "{} differences; first examples:\n{}",
        failures.len(),
        failures
            .iter()
            .take(30)
            .cloned()
            .collect::<Vec<_>>()
            .join("\n")
    );
}
#[test]
fn native_byte_boundaries_and_search_replace_preserve_exact_units_and_error_types() {
    verify(
        include_str!("../../../tests/fixtures/vba/worksheet_byte_boundaries.tsv"),
        3294,
    );
}
#[test]
fn native_character_scanning_and_search_boundaries_match_both_versions() {
    verify(
        include_str!("../../../tests/fixtures/vba/worksheet_byte_controls.tsv"),
        6240,
    );
}
#[test]
fn native_numeric_argument_limits_are_specific_to_each_function() {
    verify(
        include_str!("../../../tests/fixtures/vba/worksheet_byte_limits.tsv"),
        648,
    );
}
#[test]
fn native_long_result_limits_count_original_utf16_units() {
    verify(
        include_str!("../../../tests/fixtures/vba/worksheet_byte_long_limits.tsv"),
        28,
    );
}
#[test]
fn native_multiple_pairs_and_trailing_high_surrogates_follow_directional_scanning() {
    verify(
        include_str!("../../../tests/fixtures/vba/worksheet_byte_ends.tsv"),
        448,
    );
}
#[test]
fn external_nonfinite_positions_and_counts_are_rejected_without_panicking() {
    use oxicells_calc::functions::{call_with_text_context, Arg};
    for fn_name in ["LEFTB", "RIGHTB", "MIDB", "REPLACEB", "FINDB", "SEARCHB"] {
        for n in [f64::NAN, f64::INFINITY, f64::NEG_INFINITY] {
            let text = || Arg::Value(Value::Text("abc".into()));
            let number = |n| Arg::Value(Value::Number(n));
            let args = match fn_name {
                "LEFTB" | "RIGHTB" => vec![text(), number(n)],
                "MIDB" => vec![text(), number(n), number(1.0)],
                "REPLACEB" => vec![text(), number(n), number(1.0), text()],
                _ => vec![text(), text(), number(n)],
            };
            assert_eq!(
                call_with_text_context(fn_name, &args, 2, false),
                Value::Error(ExcelError::Value),
                "{fn_name} {n}"
            );
            if matches!(fn_name, "MIDB" | "REPLACEB") {
                let mut args = args;
                args[1] = number(1.0);
                args[2] = number(n);
                assert_eq!(
                    call_with_text_context(fn_name, &args, 2, false),
                    Value::Error(ExcelError::Value),
                    "{fn_name} count {n}"
                );
            }
        }
    }
}
