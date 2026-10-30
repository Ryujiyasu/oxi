// This Source Code Form is subject to the terms of the Mozilla Public
// License, v. 2.0. If a copy of the MPL was not distributed with this
// file, You can obtain one at https://mozilla.org/MPL/2.0/.

use oxivba_core::ast::ParamMode;
use oxivba_core::native_cache::{decode_module_cache, CacheContext, CachedType, DecodeErrorKind};

fn hex(text: &str) -> Vec<u8> {
    text.as_bytes()
        .chunks_exact(2)
        .map(|pair| u8::from_str_radix(std::str::from_utf8(pair).unwrap(), 16).unwrap())
        .collect()
}
fn cases() -> impl Iterator<Item = Vec<&'static str>> {
    include_str!("../../../tests/fixtures/vba/native_cache_modules.tsv")
        .lines()
        .filter(|line| !line.starts_with('#') && !line.starts_with("id\t"))
        .map(|line| line.split('\t').collect())
}
fn type_word(value: CachedType) -> u32 {
    match value {
        CachedType::Builtin { kind, qualifiers } => {
            0xffff0000 | u32::from(qualifiers) << 8 | u32::from(kind)
        }
        CachedType::Reference(value) => value,
    }
}
fn numbers(text: &str) -> Vec<u32> {
    text.split(',').map(|item| item.parse().unwrap()).collect()
}
fn word(bytes: &[u8], at: usize) -> usize {
    u16::from_le_bytes(bytes[at..at + 2].try_into().unwrap()) as usize
}
fn dword(bytes: &[u8], at: usize) -> usize {
    u32::from_le_bytes(bytes[at..at + 4].try_into().unwrap()) as usize
}
fn put(bytes: &mut [u8], at: usize, value: u32) {
    bytes[at..at + 4].copy_from_slice(&value.to_le_bytes());
}

#[test]
fn native_module_frames_and_declarations_preserve_observed_records() {
    let (mut lines, mut procedures, mut parameters) = (0, 0, 0);
    for row in cases() {
        let bytes = hex(row[1]);
        let module = decode_module_cache(&bytes, CacheContext::B5_WIN64).unwrap();
        assert_eq!(module.lines().len(), row[2].parse::<usize>().unwrap());
        let actual = module.procedures().unwrap();
        let expected: Vec<_> = row[3].split('|').collect();
        assert_eq!(actual.len(), expected.len(), "{}", row[0]);
        for (proc, expected) in actual.iter().zip(expected) {
            let fields: Vec<_> = expected.split('/').collect();
            assert_eq!(
                vec![
                    u32::from(proc.source_line),
                    u32::from(proc.definition_opcode),
                    proc.declaration_offset,
                    u32::from(proc.flags),
                    u32::from(proc.name_word),
                    proc.member_id,
                    proc.return_type.map(type_word).unwrap_or(u32::MAX)
                ],
                numbers(fields[0]),
                "{}",
                row[0]
            );
            assert_eq!(
                proc.definition_prefix,
                if fields[1] == "-" {
                    vec![]
                } else {
                    hex(fields[1])
                }
            );
            let params: Vec<_> = if fields[2] == "-" {
                vec![]
            } else {
                fields[2].split(';').collect()
            };
            assert_eq!(proc.parameters.len(), params.len());
            for (parameter, expected) in proc.parameters.iter().zip(params) {
                assert_eq!(
                    vec![
                        parameter.declaration_offset,
                        u32::from(parameter.flags),
                        u32::from(parameter.name_word),
                        parameter.storage_word,
                        type_word(parameter.type_descriptor),
                        parameter.mode_flags
                    ],
                    numbers(expected)
                );
                parameters += 1;
            }
            procedures += 1;
        }
        lines += module.lines().len();
    }
    assert_eq!((lines, procedures, parameters), (156, 27, 27));
}

#[test]
fn typed_signatures_keep_defaults_modes_arrays_and_unresolved_types() {
    let row = cases().find(|row| row[0] == "signatures").unwrap();
    let bytes = hex(row[1]);
    let module = decode_module_cache(&bytes, CacheContext::B5_WIN64).unwrap();
    let procedures = module.procedures().unwrap();
    // Native source order: Driver, Long, Boolean, Double, Integer, Single,
    // Currency, Date, String, Object, Variant, Byte, LongLong, Worksheet,
    // Array, ParamArray, private Sub.
    let kinds = [3, 3, 11, 5, 2, 4, 6, 7, 8, 9, 12, 17, 20];
    for (proc, kind) in procedures.iter().zip(kinds) {
        assert_eq!(
            proc.return_type,
            Some(CachedType::Builtin {
                kind,
                qualifiers: 0
            })
        );
    }
    assert_eq!(
        procedures[1].parameters[0].passing_mode(),
        Some(ParamMode::ByVal)
    );
    assert_eq!(
        procedures[3].parameters[0].passing_mode(),
        Some(ParamMode::ByRef)
    );
    assert_eq!(
        procedures[3].parameters[0].type_descriptor,
        CachedType::Builtin {
            kind: 5,
            qualifiers: 1
        }
    );
    assert_eq!(procedures[8].definition_prefix, hex("fa00b900040073656564"));
    assert_eq!(procedures[8].parameters[0].mode_flags, 0x684);
    assert!(matches!(
        procedures[13].return_type,
        Some(CachedType::Reference(_))
    ));
    for proc in &procedures[14..16] {
        let param = &proc.parameters[0];
        assert_eq!(param.name_word, 628);
        assert!(matches!(param.type_descriptor, CachedType::Reference(_)));
    }
    assert_eq!(
        procedures[15].parameters[0].passing_mode(),
        Some(ParamMode::ParamArray)
    );
    assert_eq!(procedures[16].definition_opcode, 0x0496);
    assert_eq!(procedures[16].return_type, None);
}

#[test]
fn malformed_cache_offsets_records_and_parameter_cycles_fail() {
    let row = cases().find(|row| row[0] == "signatures").unwrap();
    let original = hex(row[1]);
    let marker = dword(&original, 25) + 60;
    let count = word(&original, marker + 4);
    let table = marker + 6;
    let end = table + count * 12;
    let origin = dword(&original, 17) + 16;
    let module = decode_module_cache(&original, CacheContext::B5_WIN64).unwrap();
    let procedures = module.procedures().unwrap();
    let proc_at = origin + procedures[1].declaration_offset as usize;
    let param_offset = procedures[1].parameters[0].declaration_offset;
    for (at, value) in [(25, u32::MAX), (end + 6, u32::MAX), (table + 8, u32::MAX)] {
        let mut bytes = original.clone();
        put(&mut bytes, at, value);
        assert!(decode_module_cache(&bytes, CacheContext::B5_WIN64).is_err());
    }
    let mut bytes = original.clone();
    put(&mut bytes, proc_at + 56, u32::MAX - 1);
    assert!(decode_module_cache(&bytes, CacheContext::B5_WIN64)
        .unwrap()
        .procedures()
        .is_err());
    let mut bytes = original.clone();
    put(
        &mut bytes,
        origin + param_offset as usize + 24,
        param_offset,
    );
    assert_eq!(
        decode_module_cache(&bytes, CacheContext::B5_WIN64)
            .unwrap()
            .procedures()
            .unwrap_err()
            .kind,
        DecodeErrorKind::CyclicParameterChain
    );
    for at in [proc_at, origin + param_offset as usize] {
        let mut bytes = original.clone();
        bytes[at] &= 0xf0;
        assert_eq!(
            decode_module_cache(&bytes, CacheContext::B5_WIN64)
                .unwrap()
                .procedures()
                .unwrap_err()
                .kind,
            DecodeErrorKind::InvalidDeclarationRecord
        );
    }
}

#[test]
fn module_prefix_truncation_and_unsupported_contexts_fail_without_panics() {
    for row in cases() {
        let bytes = hex(row[1]);
        for end in 0..bytes.len() {
            assert!(
                decode_module_cache(&bytes[..end], CacheContext::B5_WIN64).is_err(),
                "{} {end}",
                row[0]
            );
        }
        for context in [
            CacheContext {
                version: 180,
                sys_kind: 3,
            },
            CacheContext {
                version: 181,
                sys_kind: 1,
            },
        ] {
            assert!(decode_module_cache(&bytes, context).is_err());
        }
    }
}
