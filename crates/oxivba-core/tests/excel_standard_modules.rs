// This Source Code Form is subject to the terms of the Mozilla Public
// License, v. 2.0. If a copy of the MPL was not distributed with this
// file, You can obtain one at https://mozilla.org/MPL/2.0/.

use oxivba_core::{parse_project, ErrorEnvironment, ErrorLanguage, Runtime, RuntimeErrorKind};

#[test]
fn standard_module_visibility_and_ambiguity_match_excel_compile_errors() {
    let fixtures = std::path::Path::new(env!("CARGO_MANIFEST_DIR")).join("../../tests/fixtures/vba");
    for row in include_str!("../../../tests/fixtures/vba/standard_compile.tsv").lines() {
        let fields: Vec<_> = row.split('\t').collect();
        let source = std::fs::read_to_string(fixtures.join(fields[1])).unwrap();
        let (module, classes) = parse_project(&source).unwrap();
        let failure = Runtime::new(&module).with_classes(&classes)
            .with_error_environment(ErrorEnvironment { language: ErrorLanguage::Japanese, help_file: String::new() })
            .call(fields[2], vec![]).unwrap_err();
        assert_eq!(failure.kind, RuntimeErrorKind::Compile, "{}", fields[0]);
        assert_eq!(failure.vba_number, None);
        assert_eq!(failure.message, format!("コンパイル エラー: {}", fields[3]), "{}", fields[0]);
    }
}
