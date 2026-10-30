// This Source Code Form is subject to the terms of the Mozilla Public
// License, v. 2.0. If a copy of the MPL was not distributed with this
// file, You can obtain one at https://mozilla.org/MPL/2.0/.

use oxivba_core::native_cache::{
    decode_identifier_table, decode_module_numeric_assignment, CacheContext, DecodeErrorKind,
    Location,
};

fn hex(text: &str) -> Vec<u8> {
    text.as_bytes()
        .chunks_exact(2)
        .map(|pair| u8::from_str_radix(std::str::from_utf8(pair).unwrap(), 16).unwrap())
        .collect()
}
fn cases() -> impl Iterator<Item = Vec<&'static str>> {
    include_str!("../../../tests/fixtures/vba/native_cache_identifiers.tsv")
        .lines()
        .filter(|line| !line.starts_with('#') && !line.starts_with("id\t"))
        .map(|line| line.split('\t').collect())
}

#[test]
fn cached_symbols_resolve_native_english_japanese_and_renamed_bindings() {
    let mut tables = 0;
    let mut bindings = 0;
    let mut names = 0;
    for row in cases() {
        let bytes = hex(row[1]);
        let table = decode_identifier_table(&bytes, CacheContext::B5_WIN64).unwrap();
        assert_eq!(table.consumed(), bytes.len(), "{}", row[0]);
        assert_eq!(table.identifiers().len(), row[2].parse::<usize>().unwrap());
        for expected in row[3].split(',') {
            assert!(
                table
                    .identifiers()
                    .iter()
                    .any(|name| name.name == hex(expected)),
                "{} required name",
                row[0]
            );
        }
        if row[4] != "-" {
            let code =
                decode_module_numeric_assignment(&hex(row[4]), CacheContext::B5_WIN64).unwrap();
            let Location::Symbol(input) = code.input else {
                panic!("source symbol")
            };
            let Location::Symbol(output) = code.output else {
                panic!("source symbol")
            };
            assert_eq!(
                table.get(input).unwrap().name,
                hex(row[5]),
                "{} parameter",
                row[0]
            );
            assert_eq!(
                table.get(output).unwrap().name,
                hex(row[6]),
                "{} procedure",
                row[0]
            );
            bindings += 1;
        }
        names += table.identifiers().len();
        tables += 1;
    }
    assert_eq!((tables, names, bindings), (13, 372, 5));
}

#[test]
fn identifier_tables_are_bounded_and_do_not_consume_following_metadata() {
    for row in cases() {
        let mut bytes = hex(row[1]);
        for end in 0..bytes.len() {
            assert!(
                decode_identifier_table(&bytes[..end], CacheContext::B5_WIN64).is_err(),
                "{} truncated at{end}",
                row[0]
            );
        }
        let end = bytes.len();
        bytes.extend_from_slice(&[0xff; 32]);
        assert_eq!(
            decode_identifier_table(&bytes, CacheContext::B5_WIN64)
                .unwrap()
                .consumed(),
            end
        );
    }
    let invalid = hex("020001000000");
    assert!(matches!(
        decode_identifier_table(&invalid, CacheContext::B5_WIN64)
            .unwrap_err()
            .kind,
        DecodeErrorKind::InvalidIdentifierCounts { .. }
    ));
    assert!(decode_identifier_table(&hex("ffff0000ffff"), CacheContext::B5_WIN64).is_err());
    assert!(decode_identifier_table(
        &hex("000000000000"),
        CacheContext {
            version: 0xb6,
            sys_kind: 3
        }
    )
    .is_err());
}

#[test]
fn reserved_and_odd_references_are_not_invented_and_name_bytes_are_preserved() {
    let mut longest = 0;
    for row in cases() {
        let bytes = hex(row[1]);
        let table = decode_identifier_table(&bytes, CacheContext::B5_WIN64).unwrap();
        assert!(table.get(0).is_none());
        assert!(table.get(table.reserved() * 2).is_none());
        assert!(table.get(u16::MAX).is_none());
        for item in table.identifiers() {
            assert_eq!(table.get(item.symbol).unwrap(), item);
            assert!(table.get(item.symbol | 1).is_none());
            assert_eq!(item.extension.is_some(), item.flags & 0x80 != 0);
            longest = longest.max(item.name.len());
        }
    }
    assert_eq!(longest, 254);
}
