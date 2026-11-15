// This Source Code Form is subject to the terms of the Mozilla Public
// License, v. 2.0. If a copy of the MPL was not distributed with this
// file, You can obtain one at https://mozilla.org/MPL/2.0/.

use oxivba_core::{parse_project, Runtime, Value};

#[test]
fn varptr_identity_and_layout_match_excel() {
    let (module, classes) = parse_project(include_str!("../../../tests/fixtures/vba/varptr_identity.bas")).unwrap();
    let yes = Value::Boolean(true);
    let no = Value::Boolean(false);
    let n = Value::Integer;
    // Complete eight-field captures from 64-bit Excel. Cases 0..17 have two
    // clean Native captures; cases 18..20 have one clean capture apiece.
    // Oxi runs every case forward and backward, retaining project storage.
    // Capture transport normalizes numeric widths; pointer VarType/TypeName are
    // explicit fields below, so pointer-width errors cannot hide in that transport.
    let expected = [
        vec![n(0),yes.clone(),yes.clone(),yes.clone(),yes.clone(),yes.clone(),yes.clone(),n(20)],
        vec![n(0),yes.clone(),yes.clone(),yes.clone(),no,yes.clone(),n(20),Value::String("LongLong".into())],
        vec![n(0),yes.clone(),yes.clone(),yes.clone(),yes.clone(),yes.clone(),yes.clone(),yes.clone()],
        vec![n(0),yes.clone(),yes.clone(),n(2),n(4),yes.clone(),n(20),Value::String("LongLong".into())],
        vec![n(0),yes.clone(),yes.clone(),n(4),n(8),yes.clone(),n(20),Value::String("LongLong".into())],
        vec![n(0),yes.clone(),yes.clone(),n(24),n(48),yes.clone(),n(20),Value::String("LongLong".into())],
        vec![n(0),yes.clone(),n(4),n(8),n(16),yes.clone(),n(20),Value::String("LongLong".into())],
        vec![n(0),yes.clone(),yes.clone(),yes.clone(),n(20),Value::String("LongLong".into()),Value::Empty,Value::Empty],
        vec![n(0),yes.clone(),yes.clone(),n(20),Value::String("LongLong".into()),Value::Empty,Value::Empty,Value::Empty],
        vec![n(0),yes.clone(),Value::Boolean(false),n(20),Value::String("LongLong".into()),Value::Empty,Value::Empty,Value::Empty],
        vec![n(0),n(2),n(6),n(10),yes.clone(),n(20),Value::String("LongLong".into()),Value::Empty],
        vec![n(0),yes.clone(),yes.clone(),yes.clone(),yes.clone(),n(20),Value::String("LongLong".into()),Value::Empty],
        vec![n(0),yes.clone(),Value::Boolean(false),Value::Boolean(false),Value::Boolean(false),n(20),Value::String("LongLong".into()),Value::Empty],
        vec![n(0),yes.clone(),yes.clone(),Value::Boolean(false),Value::Boolean(false),Value::Boolean(false),n(20),Value::String("LongLong".into())],
        vec![n(0),yes.clone(),yes.clone(),n(43),n(43),n(3),n(3),yes.clone()],
        vec![n(0),yes.clone(),Value::Boolean(false),Value::Boolean(false),Value::Boolean(false),Value::Boolean(false),Value::Boolean(false),Value::Boolean(false)],
        vec![n(13),n(13),n(42),n(42),n(3),n(3),yes.clone(),yes.clone()],
        vec![n(0),Value::Boolean(false),Value::Boolean(false),Value::Boolean(false),Value::Boolean(false),Value::Boolean(false),Value::Boolean(false),Value::Empty],
        vec![n(0),yes.clone(),yes.clone(),Value::Boolean(false),Value::Boolean(false),Value::Boolean(false),n(20),Value::Empty],
        vec![n(0),yes.clone(),yes.clone(),Value::Boolean(false),Value::Boolean(false),Value::Boolean(false),n(20),Value::Empty],
        vec![n(0),yes.clone(),yes,Value::Boolean(false),Value::Boolean(false),Value::Boolean(false),n(20),Value::Empty],
    ];
    let mut runtime = Runtime::new(&module).with_classes(&classes);
    for case in (0..expected.len()).chain((0..expected.len()).rev()) {
        let result = runtime.call("Probe",vec![n(case as i64),n(0)]).unwrap();
        let Value::Array(values) = result else { panic!("case {case}: expected captured array"); };
        let actual: Vec<_> = values.values.iter().map(|value| match value {
            Value::Byte(value) => n(i64::from(*value)),
            Value::Int16(value) => n(i64::from(*value)),
            Value::LongLong(value) => n(*value),
            value => value.clone(),
        }).collect();
        assert_eq!(actual,expected[case],"case {case}");
        let global = runtime.call("PG", vec![]).unwrap();
        let persistent = runtime.call("PST", vec![]).unwrap();
        runtime = Runtime::new(&module).with_classes(&classes).with_project_state(runtime.into_project_state());
        assert_eq!(runtime.call("PG", vec![]).unwrap(), global);
        assert_eq!(runtime.call("PST", vec![]).unwrap(), persistent);
    }
}

#[test]
fn reference_places_and_implicit_types_match_excel() {
    let fixtures = std::path::Path::new(env!("CARGO_MANIFEST_DIR")).join("../../tests/fixtures/vba");
    let rows: Vec<Vec<_>> = include_str!("../../../tests/fixtures/vba/reference_places.tsv").lines().map(|row| row.split('\t').collect()).collect();
    let mut at = 0;
    while at < rows.len() {
        let name = rows[at][0];
        let source = std::fs::read_to_string(fixtures.join(rows[at][1])).unwrap();
        let (module, classes) = parse_project(&source).unwrap();
        let mut continuous = Runtime::new(&module).with_classes(&classes);
        let mut previous = None;
        while at < rows.len() && rows[at][0] == name {
            let fields = &rows[at];
            let expected = Value::String(fields[3].into());
            assert_eq!(continuous.call(fields[2], vec![]).unwrap_or_else(|failure| panic!("continuous {name}: {failure:?}")), expected, "continuous {name}");
            let (module, classes) = parse_project(&source).unwrap();
            let mut runtime = Runtime::new(&module).with_classes(&classes);
            if let Some(saved) = previous.take() { runtime = runtime.with_project_state(saved); }
            assert_eq!(runtime.call(fields[2], vec![]).unwrap_or_else(|failure| panic!("resumed {name}: {failure:?}")), expected, "resumed {name}");
            previous = Some(runtime.into_project_state());
            at += 1;
        }
    }
}
