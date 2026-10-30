// This Source Code Form is subject to the terms of the Mozilla Public
// License, v. 2.0. If a copy of the MPL was not distributed with this
// file, You can obtain one at https://mozilla.org/MPL/2.0/.
use oxicells_calc::{Value, Workbook};
fn value(report: &str) -> Value {
    let mut fields = report.split(':');
    let length: usize = fields.next().unwrap().parse().unwrap();
    let units: Vec<u16> = fields
        .map(|text| text.parse::<i16>().unwrap() as u16)
        .collect();
    assert_eq!(units.len(), length);
    Value::from_utf16(units)
}
fn report(value: &Value) -> String {
    match value {
        Value::Error(_) => "E".into(),
        Value::Number(number) => format!("5:{number}"),
        Value::Logical(state) => format!("11:{}", if *state { "True" } else { "False" }),
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
#[test]
fn native_compatibility_version_text_matrix() {
    let mut count = 0;
    for line in include_str!("../../../tests/fixtures/vba/worksheet_utf16_matrix.tsv")
        .lines()
        .filter(|line| !line.starts_with('#'))
    {
        let fields: Vec<_> = line.split('\t').collect();
        let mut book = Workbook::new();
        book.add_sheet("S");
        book.set_compatibility_version(fields[0].parse().unwrap());
        book.set_value("S", "J20", value(fields[1])).unwrap();
        book.set_value("S", "K20", value("1:-10179")).unwrap();
        book.set_value("S", "L20", value("1:-8704")).unwrap();
        let actual = book.evaluate("S", fields[2]).unwrap();
        assert_eq!(report(&actual), fields[3], "{line}");
        count += 1;
    }
    assert_eq!(count, 308);
}
#[test]
fn native_managed_text_expressions_and_comparisons() {
    let mut count = 0;
    for line in include_str!("../../../tests/fixtures/vba/worksheet_utf16.tsv")
        .lines()
        .filter(|line| !line.starts_with('#'))
    {
        let fields: Vec<_> = line.split('\t').collect();
        let mut book = Workbook::new();
        book.add_sheet("S");
        book.set_compatibility_version(fields[0].parse().unwrap());
        for (address, units) in [
            ("A1", "1:-10179"),
            ("A2", "1:-8704"),
            ("A3", "2:-10179:-8704"),
        ] {
            book.set_value("S", address, value(units)).unwrap();
        }
        let actual = book.evaluate("S", fields[1]).unwrap();
        assert_eq!(report(&actual), fields[2], "{line}");
        count += 1;
    }
    assert_eq!(count, 44);
}
