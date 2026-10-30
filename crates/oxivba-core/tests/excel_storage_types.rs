// This Source Code Form is subject to the terms of the Mozilla Public
// License, v. 2.0. If a copy of the MPL was not distributed with this
// file, You can obtain one at https://mozilla.org/MPL/2.0/.

use oxivba_core::{
    parse_project, ErrorEnvironment, ErrorLanguage, Runtime, RuntimeErrorKind, Value,
};

#[test]
fn storage_types_and_member_visibility_match_excel() {
    let fixtures =
        std::path::Path::new(env!("CARGO_MANIFEST_DIR")).join("../../tests/fixtures/vba");
    let rows: Vec<Vec<_>> = include_str!("../../../tests/fixtures/vba/storage_types.tsv")
        .lines()
        .map(|line| line.split('\t').collect())
        .collect();
    let mut at = 0;
    while at < rows.len() {
        let name = rows[at][0];
        let source = std::fs::read_to_string(fixtures.join(rows[at][1])).unwrap();
        let (module, classes) = parse_project(&source).unwrap();
        let environment = || ErrorEnvironment {
            language: ErrorLanguage::Japanese,
            help_file: String::new(),
        };
        let mut continuous = Runtime::new(&module)
            .with_classes(&classes)
            .with_error_environment(environment());
        let mut saved = None;
        while at < rows.len() && rows[at][0] == name {
            let fields = &rows[at];
            let verify = |actual: Result<Value, oxivba_core::RuntimeError>| {
                if fields[3] == "compile" {
                    let error = actual.unwrap_err();
                    assert_eq!(error.kind, RuntimeErrorKind::Compile, "{name}");
                    assert_eq!(error.vba_number, None);
                    assert_eq!(
                        error.message,
                        format!("コンパイル エラー: {}", fields[4]),
                        "{name}"
                    );
                } else {
                    assert_eq!(
                        actual.unwrap_or_else(|error| panic!("{name}: {error:?}")),
                        Value::String(fields[4].into()),
                        "{name}"
                    );
                }
            };
            verify(continuous.call(fields[2], vec![]));
            let (module, classes) = parse_project(&source).unwrap();
            let mut restored = Runtime::new(&module)
                .with_classes(&classes)
                .with_error_environment(environment());
            if let Some(previous) = saved.take() {
                restored = restored.with_project_state(previous);
            }
            verify(restored.call(fields[2], vec![]));
            saved = Some(restored.into_project_state());
            at += 1;
        }
    }
}
