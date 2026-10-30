// This Source Code Form is subject to the terms of the Mozilla Public
// License, v. 2.0. If a copy of the MPL was not distributed with this
// file, You can obtain one at https://mozilla.org/MPL/2.0/.

use oxivba_core::ast::{ModuleItem, Visibility};
use oxivba_core::{analyse, execute, parse_module, Value};
use std::collections::BTreeSet;

#[test]
fn excel_err_properties() {
    let source = include_str!("../../../tests/fixtures/vba/err_properties.bas");
    let expected = include_str!("../../../tests/fixtures/vba/err_properties.tsv");
    let module = parse_module(source).expect("valid VBA fixture");
    assert_eq!(analyse(&module).metrics.unparsed, 0);
    let public_functions: BTreeSet<_> = module
        .items
        .iter()
        .filter_map(|item| match item {
            ModuleItem::Procedure(procedure) if procedure.visibility == Visibility::Public => {
                Some(procedure.name.as_str())
            }
            _ => None,
        })
        .collect();
    let mut case_names = BTreeSet::new();
    let mut failures = Vec::new();
    let mut count = 0;
    for case in expected.lines() {
        let (name, expected) = case.split_once('\t').expect("case and Excel result");
        assert!(case_names.insert(name), "duplicate case: {name}");
        let actual = execute(&module, name, vec![]);
        if actual != Ok(Value::String(expected.to_owned())) {
            failures.push(format!("{name}: expected {expected:?}, got {actual:?}"));
        }
        count += 1;
    }
    assert_eq!(
        case_names, public_functions,
        "every public fixture procedure must be compared"
    );
    assert!(
        failures.is_empty(),
        "{} / {count} mismatches:\n{}",
        failures.len(),
        failures.join("\n")
    );
}
