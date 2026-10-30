// This Source Code Form is subject to the terms of the Mozilla Public
// License, v. 2.0. If a copy of the MPL was not distributed with this
// file, You can obtain one at https://mozilla.org/MPL/2.0/.
use oxivba_core::{parse_project, Runtime, Value};
const SOURCE: &str = include_str!("../../../tests/fixtures/vba/string_boundaries.bas");
#[test]
fn native_utf16_string_operations_preserve_unpaired_units() {
    verify("UnicodeCases");
}
#[test]
fn native_bstr_allocation_copy_and_reference_relations() {
    verify("PointerCases");
}
fn verify(name: &str) {
    let (module, classes) = parse_project(SOURCE).unwrap();
    let expected = include_str!("../../../tests/fixtures/vba/string_boundaries.tsv")
        .lines()
        .find_map(|l| l.strip_prefix(&format!("{name}\t")))
        .unwrap();
    let actual = Runtime::new(&module)
        .with_classes(&classes)
        .call(name, vec![]);
    assert_eq!(actual.unwrap(), Value::String(expected.into()));
}
#[test]
fn utf16_values_never_replace_unpaired_surrogates() {
    let units = vec![0xd83d, 0, 0xde00, 0xffff];
    let value = Value::from_utf16(units.clone());
    assert_eq!(value.string_units().unwrap().as_ref(), units);
    assert!(value.unicode_text().unwrap().is_err());
    assert_eq!(Value::NullString, Value::String(String::new()));
}
#[test]
fn native_text_comparison_and_filter_do_not_merge_different_surrogates() {
    let source = include_str!("../../../tests/fixtures/vba/string_text_boundaries.bas");
    let expected = include_str!("../../../tests/fixtures/vba/string_text_boundaries.tsv")
        .lines()
        .find_map(|l| l.strip_prefix("TextCases\t"))
        .unwrap();
    let (module, classes) = parse_project(source).unwrap();
    assert_eq!(
        Runtime::new(&module)
            .with_classes(&classes)
            .call("TextCases", vec![])
            .unwrap(),
        Value::String(expected.into())
    );
}
#[test]
fn native_strptr_coercions_distinguish_empty_null_error_and_nothing() {
    let source = include_str!("../../../tests/fixtures/vba/string_error_boundaries.bas");
    let expected = include_str!("../../../tests/fixtures/vba/string_error_boundaries.tsv")
        .lines()
        .find_map(|l| l.strip_prefix("ErrorCases\t"))
        .unwrap();
    let (module, classes) = parse_project(source).unwrap();
    assert_eq!(
        Runtime::new(&module)
            .with_classes(&classes)
            .call("ErrorCases", vec![])
            .unwrap(),
        Value::String(expected.into())
    );
}
#[test]
fn native_literal_pool_and_full_assignment_allocations_remain_distinct() {
    let source = include_str!("../../../tests/fixtures/vba/string_literal_boundaries.bas");
    let expected = include_str!("../../../tests/fixtures/vba/string_literal_boundaries.tsv")
        .lines()
        .find_map(|l| l.strip_prefix("LiteralCases\t"))
        .unwrap();
    let (module, classes) = parse_project(source).unwrap();
    let mut runtime = Runtime::new(&module).with_classes(&classes);
    assert_eq!(
        runtime.call("LiteralCases", vec![]).unwrap(),
        Value::String(expected.into())
    );
}
