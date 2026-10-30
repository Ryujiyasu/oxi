// This Source Code Form is subject to the terms of the Mozilla Public
// License, v. 2.0. If a copy of the MPL was not distributed with this
// file, You can obtain one at https://mozilla.org/MPL/2.0/.

use oxivba_core::native_cache::{
    decode_identifier_table, decode_module_cache, decode_project_cache, CacheContext,
};

fn hex(text: &str) -> Vec<u8> {
    text.as_bytes()
        .chunks_exact(2)
        .map(|pair| u8::from_str_radix(std::str::from_utf8(pair).unwrap(), 16).unwrap())
        .collect()
}
fn cases() -> impl Iterator<Item = Vec<&'static str>> {
    include_str!("../../../tests/fixtures/vba/native_cache_projects.tsv")
        .lines()
        .filter(|line| !line.starts_with('#') && !line.starts_with("id\t"))
        .map(|line| line.split('\t').collect())
}
fn optional_hex(text: &str) -> Option<Vec<u8>> {
    (text != "-").then(|| hex(text))
}
fn utf16(bytes: &[u8]) -> String {
    String::from_utf16(
        &bytes
            .chunks_exact(2)
            .map(|p| u16::from_le_bytes([p[0], p[1]]))
            .collect::<Vec<_>>(),
    )
    .unwrap()
}
fn word(bytes: &[u8], at: usize) -> usize {
    u16::from_le_bytes(bytes[at..at + 2].try_into().unwrap()) as usize
}
fn put_word(bytes: &mut [u8], at: usize, value: u16) {
    bytes[at..at + 2].copy_from_slice(&value.to_le_bytes());
}

#[test]
fn native_project_structures_locate_names_across_all_cache_fixtures() {
    let (mut projects, mut modules, mut indexed, mut extra) = (0, 0, 0, 0);
    for row in cases() {
        let bytes = hex(row[1]);
        let project = decode_project_cache(&bytes, CacheContext::B5_WIN64)
            .unwrap_or_else(|e| panic!("{}: {e}", row[0]));
        assert_eq!(project.consumed, bytes.len(), "{}", row[0]);
        assert_eq!(project.name_table_offset, row[2].parse::<usize>().unwrap());
        assert_eq!(project.references.len(), row[3].parse::<usize>().unwrap());
        let names: Vec<_> = row[4].split(',').map(optional_hex).collect();
        assert_eq!(project.modules.len(), names.len());
        for (module, name) in project.modules.iter().zip(names) {
            assert_eq!(module.name, name.as_deref());
        }
        assert_eq!(
            project.identifiers.identifiers().len(),
            row[5].parse::<usize>().unwrap()
        );
        let expected: Vec<_> = if row[6] == "-" {
            vec![]
        } else {
            row[6].split(';').collect()
        };
        assert_eq!(
            project.identifiers.unindexed_identifiers().len(),
            expected.len()
        );
        for (actual, expected) in project
            .identifiers
            .unindexed_identifiers()
            .iter()
            .zip(expected)
        {
            let fields: Vec<_> = expected.split(',').collect();
            assert_eq!(actual.before_symbol, fields[0].parse::<u16>().unwrap());
            assert_eq!(actual.reference, fields[1].parse::<u16>().unwrap());
            assert_eq!(actual.flags, fields[2].parse::<u8>().unwrap());
            assert_eq!(
                actual.extension.map(|v| v.to_vec()),
                optional_hex(fields[3])
            );
            assert_eq!(actual.name, hex(fields[4]));
        }
        if row[7] != "-" {
            assert_eq!(
                project
                    .identifiers
                    .get(project.project_name_word)
                    .unwrap()
                    .name,
                hex(row[7])
            );
        }
        assert_eq!(project.description, optional_hex(row[8]).as_deref());
        assert_eq!(project.help_file, optional_hex(row[9]).as_deref());
        projects += 1;
        modules += project.modules.len();
        indexed += project.identifiers.identifiers().len();
        extra += project.identifiers.unindexed_identifiers().len();
    }
    assert_eq!((projects, modules, indexed, extra), (63, 279, 3610, 5));
}

#[test]
fn native_metadata_survives_variable_descriptions_help_and_fifth_reference() {
    let row = cases().find(|row| row[0] == "native_metadata").unwrap();
    let bytes = hex(row[1]);
    let p = decode_project_cache(&bytes, CacheContext::B5_WIN64).unwrap();
    assert_eq!(p.references.len(), 5);
    assert!(utf16(p.references[4].identity.unwrap()).contains("scrrun.dll"));
    assert_eq!(
        p.identifiers.get(p.project_name_word).unwrap().name,
        b"LayoutMetadata"
    );
    assert_eq!(
        utf16(p.description.unwrap()),
        "Native layout probe: description with nonempty metadata"
    );
    assert_eq!(utf16(p.help_file.unwrap()), "layout-probe.chm");
    assert_eq!(
        p.modules
            .iter()
            .map(|m| utf16(m.name.unwrap()))
            .collect::<Vec<_>>(),
        ["ThisWorkbook", "Sheet1", "SaveProbe"]
    );
    assert_eq!(p.name_table_offset, 2525);
}

#[test]
fn additional_names_preserve_references_without_shifting_user_symbol_ids() {
    for row in cases().filter(|r| r[6] != "-") {
        let bytes = hex(row[1]);
        let p = decode_project_cache(&bytes, CacheContext::B5_WIN64).unwrap();
        let table = &p.identifiers;
        assert_eq!(table.get(526).unwrap().name, b"Excel");
        assert_eq!(
            usize::from(table.record_count()),
            table.identifiers().len() + table.unindexed_identifiers().len()
        );
        for name in table.unindexed_identifiers() {
            assert_eq!(name.before_symbol, 526);
            assert!(table.get(name.reference).is_none());
        }
        if row[0] == "run_places_arrays" {
            let names = table.unindexed_identifiers();
            assert_eq!((names[0].reference, names[0].name), (11, &b"b"[..]));
            assert_eq!((names[1].reference, names[1].name), (177, &b"Text"[..]));
        } else {
            let name = &table.unindexed_identifiers()[0];
            assert_eq!((name.reference, name.name), (177, &b"text"[..]));
            assert_eq!(name.flags, 0x84);
            assert!(name.extension.is_some());
        }
        for name in table.identifiers() {
            assert_eq!(table.get(name.symbol), Some(name));
        }
    }
}

#[test]
fn project_names_bind_module_procedures_without_fixture_symbol_locations() {
    let row = cases().find(|r| r[0] == "save_calculation").unwrap();
    let bytes = hex(row[1]);
    let p = decode_project_cache(&bytes, CacheContext::B5_WIN64).unwrap();
    let row = include_str!("../../../tests/fixtures/vba/native_cache_modules.tsv")
        .lines()
        .find(|r| r.starts_with("normal-original\t"))
        .unwrap();
    let module_bytes = hex(row.split('\t').nth(1).unwrap());
    let module = decode_module_cache(&module_bytes, CacheContext::B5_WIN64).unwrap();
    let names = module
        .procedures()
        .unwrap()
        .iter()
        .map(|proc| {
            String::from_utf8(p.identifiers.get(proc.name_word).unwrap().name.to_vec()).unwrap()
        })
        .collect::<Vec<_>>();
    assert_eq!(
        names,
        ["Reset", "Configure", "ModeOnly", "OnSaveOnly", "Twice"]
    );
}

#[test]
fn name_location_follows_lengths_past_decoy_tables_inside_reference_text() {
    let row = cases().find(|r| r[0] == "save_calculation").unwrap();
    let mut bytes = hex(row[1]);
    let expected_offset = row[2].parse::<usize>().unwrap();
    let decoy = hex("0701010006010000000001004100");
    assert_eq!(
        decode_identifier_table(&decoy, CacheContext::B5_WIN64)
            .unwrap()
            .identifiers()[0]
            .name,
        b"A"
    );
    let length = word(&bytes, 34);
    let at = 36 + length;
    put_word(&mut bytes, 34, (length + decoy.len()) as u16);
    bytes.splice(at..at, decoy.iter().copied());
    let p = decode_project_cache(&bytes, CacheContext::B5_WIN64).unwrap();
    assert_eq!(p.name_table_offset, expected_offset + decoy.len());
    assert_eq!(p.identifiers.get(526).unwrap().name, b"Excel");
    assert_eq!(
        p.identifiers.identifiers().len(),
        row[5].parse::<usize>().unwrap()
    );
}

#[test]
fn malformed_project_counts_strings_context_and_auxiliary_records_are_rejected() {
    let row = cases().find(|r| r[0] == "project_run_arguments").unwrap();
    let bytes = hex(row[1]);
    let name_at = row[2].parse::<usize>().unwrap();
    for (at, value) in [
        (0, 0),
        (2, 180),
        (18, 1),
        (30, u16::MAX),
        (34, 1),
        (name_at + 2, 52),
        (name_at + 2, 54),
    ] {
        let mut bad = bytes.clone();
        put_word(&mut bad, at, value);
        assert!(
            decode_project_cache(&bad, CacheContext::B5_WIN64).is_err(),
            "offset{at} value{value}"
        );
    }
    let mut bad = bytes.clone();
    put_word(&mut bad, name_at + 14, 0);
    assert!(decode_project_cache(&bad, CacheContext::B5_WIN64).is_err());
    assert!(decode_project_cache(
        &bytes,
        CacheContext {
            version: 181,
            sys_kind: 1
        }
    )
    .is_err());
}

#[test]
fn every_truncated_native_project_prefix_is_bounded_and_trailing_data_is_retained() {
    for row in cases() {
        let mut bytes = hex(row[1]);
        let end = bytes.len();
        for length in 0..end {
            assert!(
                decode_project_cache(&bytes[..length], CacheContext::B5_WIN64).is_err(),
                "{} {length}",
                row[0]
            );
        }
        bytes.extend_from_slice(&[0xff; 32]);
        assert_eq!(
            decode_project_cache(&bytes, CacheContext::B5_WIN64)
                .unwrap()
                .consumed,
            end
        );
    }
}
