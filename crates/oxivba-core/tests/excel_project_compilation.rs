// This Source Code Form is subject to the terms of the Mozilla Public
// License, v. 2.0. If a copy of the MPL was not distributed with this
// file, You can obtain one at https://mozilla.org/MPL/2.0/.

use oxivba_core::{
    compile_project, compile_project_with_references, parse_project, CompileReferences,
    ErrorLanguage, Runtime, Value,
};

fn excel_references() -> CompileReferences {
    let mut references = CompileReferences::default();
    references.library = std::sync::Arc::new(
        oxivba_core::CompileReferenceLibrary::from_tsv(include_str!(
            "../../oxidocs-wasm/src/vba_reference_signatures.tsv"
        ))
        .unwrap(),
    );
    for row in include_str!("../../oxidocs-wasm/src/vba_builtin_references.tsv")
        .lines()
        .filter(|row| !row.is_empty() && !row.starts_with('#'))
    {
        let fields: Vec<_> = row.split('\t').collect();
        match fields[0] {
            "N" => references.namespaces.push(fields[1].into()),
            "T" => references.types.push((fields[1].into(), fields[2] == "1")),
            "S" => references.symbols.push((
                fields[1].into(),
                if fields[2] == "-" {
                    None
                } else {
                    Some(fields[2].parse().unwrap())
                },
            )),
            _ => panic!("reference row"),
        }
    }
    references.symbols.push(("Sheet1".into(), None));
    references.current_year = Some(2026);
    references
}

#[test]
fn document_module_compilation_matches_excel() {
    let fixtures =
        std::path::Path::new(env!("CARGO_MANIFEST_DIR")).join("../../tests/fixtures/vba");
    let base = excel_references();
    let mut mismatches = Vec::new();
    for line in include_str!("../../../tests/fixtures/vba/document_module_compilation.tsv").lines()
    {
        let case: Vec<_> = line.split('\t').collect();
        let mut references = base.clone();
        references
            .document_modules
            .push((case[4].into(), case[5].into()));
        let source = std::fs::read_to_string(fixtures.join(case[1])).unwrap();
        let result = compile_project_with_references(&source, &references);
        let matches = if case[2] == "compile" {
            result
                .as_ref()
                .err()
                .is_some_and(|failure| failure.error.message(ErrorLanguage::Japanese) == case[3])
        } else {
            result.is_ok()
        };
        if !matches {
            mismatches.push(format!("{}: {:?}", case[0], result));
        }
    }
    assert!(mismatches.is_empty(), "{}", mismatches.join("\n"));
}

#[test]
fn project_compilation_reference_context_matches_excel() {
    let fixtures =
        std::path::Path::new(env!("CARGO_MANIFEST_DIR")).join("../../tests/fixtures/vba");
    let references = excel_references();
    for row in include_str!("../../../tests/fixtures/vba/project_compilation_context.tsv").lines() {
        let fields: Vec<_> = row.split('\t').collect();
        let source = std::fs::read_to_string(fixtures.join(fields[1])).unwrap();
        let result = compile_project_with_references(&source, &references);
        if fields[3] == "compile" {
            let failure = result
                .err()
                .unwrap_or_else(|| panic!("{} compiled unexpectedly", fields[0]));
            assert_eq!(
                failure.error.message(ErrorLanguage::Japanese),
                fields[4],
                "{}",
                fields[0]
            );
        } else {
            result.unwrap_or_else(|error| panic!("{}: {error:?}", fields[0]));
        }
    }
}

#[test]
fn reference_library_compilation_matches_excel() {
    let fixtures =
        std::path::Path::new(env!("CARGO_MANIFEST_DIR")).join("../../tests/fixtures/vba");
    let references = excel_references();
    for row in include_str!("../../../tests/fixtures/vba/reference_compilation.tsv").lines() {
        let fields: Vec<_> = row.split('\t').collect();
        let source = std::fs::read_to_string(fixtures.join(fields[1])).unwrap();
        let result = compile_project_with_references(&source, &references);
        if fields[3] == "compile" {
            let error = result
                .err()
                .unwrap_or_else(|| panic!("{} compiled unexpectedly", fields[0]));
            assert_eq!(
                error.error.message(ErrorLanguage::Japanese),
                fields[4],
                "{}",
                fields[0]
            );
        } else {
            result.unwrap_or_else(|error| panic!("{}: {error:?}", fields[0]));
        }
    }
}

#[test]
fn argument_binding_compilation_matches_excel() {
    let fixtures =
        std::path::Path::new(env!("CARGO_MANIFEST_DIR")).join("../../tests/fixtures/vba");
    let references = excel_references();
    for row in include_str!("../../../tests/fixtures/vba/argument_binding.tsv").lines() {
        let fields: Vec<_> = row.split('\t').collect();
        let source = std::fs::read_to_string(fixtures.join(fields[1])).unwrap();
        let result = compile_project_with_references(&source, &references);
        if fields[3] == "compile" {
            let error = result
                .err()
                .unwrap_or_else(|| panic!("{} compiled unexpectedly", fields[0]));
            assert_eq!(
                error.error.message(ErrorLanguage::Japanese),
                fields[4],
                "{}",
                fields[0]
            );
        } else {
            result.unwrap_or_else(|error| panic!("{}: {error:?}", fields[0]));
        }
    }
}

#[test]
fn name_resolution_and_object_declarations_match_excel() {
    let fixtures =
        std::path::Path::new(env!("CARGO_MANIFEST_DIR")).join("../../tests/fixtures/vba");
    let references = excel_references();
    for row in include_str!("../../../tests/fixtures/vba/name_resolution.tsv").lines() {
        let fields: Vec<_> = row.split('\t').collect();
        let source = std::fs::read_to_string(fixtures.join(fields[1])).unwrap();
        let result = compile_project_with_references(&source, &references);
        if fields[3] == "compile" {
            let error = result
                .err()
                .unwrap_or_else(|| panic!("{} compiled unexpectedly", fields[0]));
            assert_eq!(
                error.error.message(ErrorLanguage::Japanese),
                fields[4],
                "{}",
                fields[0]
            );
        } else {
            result.unwrap_or_else(|error| panic!("{}: {error:?}", fields[0]));
            let (module, classes) = parse_project(&source).unwrap();
            let mut runtime = Runtime::new(&module).with_classes(&classes);
            for value in &fields[4..] {
                compile_project_with_references(&source, &references).unwrap();
                assert_eq!(
                    runtime
                        .call(fields[2], vec![])
                        .unwrap_or_else(|error| panic!("{}: {error:?}", fields[0])),
                    Value::String((*value).into()),
                    "{}",
                    fields[0]
                );
            }
        }
    }
}

#[test]
fn expression_compilation_and_result_indices_match_excel() {
    let fixtures =
        std::path::Path::new(env!("CARGO_MANIFEST_DIR")).join("../../tests/fixtures/vba");
    let references = excel_references();
    for row in include_str!("../../../tests/fixtures/vba/expression_compilation.tsv").lines() {
        let fields: Vec<_> = row.split('\t').collect();
        let source = std::fs::read_to_string(fixtures.join(fields[1])).unwrap();
        let result = compile_project_with_references(&source, &references);
        if fields[3] == "compile" {
            let error = result
                .err()
                .unwrap_or_else(|| panic!("{} compiled unexpectedly", fields[0]));
            assert_eq!(
                error.error.message(ErrorLanguage::Japanese),
                fields[4],
                "{}",
                fields[0]
            );
        } else {
            result.unwrap_or_else(|error| panic!("{}: {error:?}", fields[0]));
            let (module, classes) = parse_project(&source).unwrap();
            let mut runtime = Runtime::new(&module).with_classes(&classes);
            for value in &fields[4..] {
                compile_project_with_references(&source, &references).unwrap();
                assert_eq!(
                    runtime
                        .call(fields[2], vec![])
                        .unwrap_or_else(|error| panic!("{}: {error:?}", fields[0])),
                    Value::String((*value).into()),
                    "{}",
                    fields[0]
                );
            }
        }
    }
}

#[test]
fn logical_lines_and_continued_branches_match_excel() {
    let fixtures =
        std::path::Path::new(env!("CARGO_MANIFEST_DIR")).join("../../tests/fixtures/vba");
    let references = excel_references();
    let mut mismatches = Vec::new();
    for row in include_str!("../../../tests/fixtures/vba/logical_line_compilation.tsv").lines() {
        let fields: Vec<_> = row.split('\t').collect();
        let source = std::fs::read_to_string(fixtures.join(fields[1])).unwrap();
        let result = compile_project_with_references(&source, &references);
        if fields[3] == "compile" {
            match result {
                Err(error) if error.error.message(ErrorLanguage::Japanese) == fields[4] => {}
                actual => mismatches.push(format!(
                    "{} expected {}: {:?}",
                    fields[0], fields[4], actual
                )),
            }
        } else {
            if let Err(error) = result {
                mismatches.push(format!("{} rejected: {error:?}", fields[0]));
                continue;
            }
            let (module, classes) = parse_project(&source).unwrap();
            let mut runtime = Runtime::new(&module).with_classes(&classes);
            for value in &fields[4..] {
                let actual = runtime.call(fields[2], vec![]);
                if actual.as_ref() != Ok(&Value::String((*value).into())) {
                    mismatches.push(format!("{} expected value {value}: {actual:?}", fields[0]));
                    break;
                }
            }
        }
    }
    assert!(mismatches.is_empty(), "{}", mismatches.join("\n"));
}

#[test]
fn explicit_project_compilation_matches_excel() {
    let fixtures =
        std::path::Path::new(env!("CARGO_MANIFEST_DIR")).join("../../tests/fixtures/vba");
    for row in include_str!("../../../tests/fixtures/vba/project_compilation.tsv").lines() {
        let fields: Vec<_> = row.split('\t').collect();
        let source = std::fs::read_to_string(fixtures.join(fields[1])).unwrap();
        let check = compile_project(&source);
        if fields[3] == "compile" {
            let error = check
                .err()
                .unwrap_or_else(|| panic!("{} compiled unexpectedly", fields[0]));
            assert_eq!(
                error.error.message(ErrorLanguage::Japanese),
                fields[4],
                "{} ({})",
                fields[0],
                error.module
            );
            continue;
        }
        check.unwrap_or_else(|error| panic!("{}: {error:?}", fields[0]));
        let (module, classes) = parse_project(&source).unwrap();
        let mut runtime = Runtime::new(&module).with_classes(&classes);
        for expected in &fields[5..] {
            compile_project(&source).unwrap();
            assert_eq!(
                runtime
                    .call(fields[2], vec![])
                    .unwrap_or_else(|error| panic!("{}: {error:?}", fields[0])),
                Value::String((*expected).into()),
                "{}",
                fields[0]
            );
        }
    }
}
