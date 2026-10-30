// This Source Code Form is subject to the terms of the Mozilla Public
// License, v. 2.0. If a copy of the MPL was not distributed with this
// file, You can obtain one at https://mozilla.org/MPL/2.0/.
use oxicells_calc::{Value, Workbook};
fn units(report: &str) -> Value {
    let mut fields = report.split(':');
    let size: usize = fields.next().unwrap().parse().unwrap();
    let units: Vec<u16> = fields.map(|n| n.parse::<i16>().unwrap() as u16).collect();
    assert_eq!(size, units.len());
    if units.is_empty() {
        Value::Blank
    } else {
        Value::from_utf16(units)
    }
}
fn report(value: Value) -> String {
    match value {
        Value::Error(_) => "E".into(),
        Value::Number(n) => format!("5:{n}"),
        Value::Logical(b) => format!("11:{}", if b { "True" } else { "False" }),
        Value::Text(_) | Value::Utf16(_) => {
            let u = value.to_utf16().unwrap();
            format!(
                "S:{}{}",
                u.len(),
                u.iter()
                    .map(|u| format!(":{}", *u as i16))
                    .collect::<String>()
            )
        }
        Value::Blank => "0:".into(),
    }
}
fn verify(fixture: &str, expected: usize) {
    let mut count = 0;
    for line in fixture.lines().filter(|line| !line.starts_with('#')) {
        let fields: Vec<_> = line.split('\t').collect();
        let mut book = Workbook::new();
        book.add_sheet("S");
        book.set_compatibility_version(fields[0].parse().unwrap());
        book.set_dbcs_language(false);
        book.set_value("S", "J20", units(fields[1])).unwrap();
        book.set_value("S", "K20", units("1:-10179")).unwrap();
        assert_eq!(
            report(book.evaluate("S", fields[2]).unwrap()),
            fields[3],
            "{line}"
        );
        count += 1;
    }
    assert_eq!(count, expected);
}
#[test]
fn native_text_extensions_preserve_units_in_both_compatibility_versions() {
    verify(
        include_str!("../../../tests/fixtures/vba/worksheet_text_extensions.tsv"),
        440,
    );
}
#[test]
fn native_byte_functions_use_the_calculation_language_context() {
    verify(
        include_str!("../../../tests/fixtures/vba/worksheet_byte_extensions.tsv"),
        112,
    );
}
