// This Source Code Form is subject to the terms of the Mozilla Public
// License, v. 2.0. If a copy of the MPL was not distributed with this
// file, You can obtain one at https://mozilla.org/MPL/2.0/.

use oxivba_core::{parse_project, Runtime, Value};

#[test]
fn paramarray_len_and_nested_arrays_match_excel() {
    let fixtures = std::path::Path::new(env!("CARGO_MANIFEST_DIR")).join("../../tests/fixtures/vba");
    let rows: Vec<Vec<_>> = include_str!("../../../tests/fixtures/vba/paramarray.tsv").lines().map(|row| row.split('\t').collect()).collect();
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
