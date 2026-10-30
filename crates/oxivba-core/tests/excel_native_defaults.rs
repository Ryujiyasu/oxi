// This Source Code Form is subject to the terms of the Mozilla Public
// License, v. 2.0. If a copy of the MPL was not distributed with this
// file, You can obtain one at https://mozilla.org/MPL/2.0/.

use oxivba_core::{
    native_cache::{
        decode_companion_default_values, decode_runtime_companion_cache,
        decode_runtime_module_heap, CacheContext, CompanionDefaultValue as D,
    },
    Value,
};
fn hex(s: &str) -> Vec<u8> {
    s.as_bytes()
        .chunks_exact(2)
        .map(|b| u8::from_str_radix(std::str::from_utf8(b).unwrap(), 16).unwrap())
        .collect()
}
fn cases() -> impl Iterator<Item = Vec<&'static str>> {
    include_str!("../../../tests/fixtures/vba/native_cache_defaults.tsv")
        .lines()
        .filter(|l| !l.starts_with('#') && !l.starts_with("id\t"))
        .map(|l| l.split('\t').collect())
}
fn expected(s: &str) -> D<'_> {
    match s {
        "empty" => D::Empty,
        "null" => D::Null,
        "nothing" => D::Nothing,
        _ => {
            let (k, v) = s.split_once(':').unwrap();
            match k {
                "byte" => D::Byte(v.parse().unwrap()),
                "i16" => D::Integer(v.parse().unwrap()),
                "long" => D::Long(v.parse().unwrap()),
                "i64" => D::LongLong(v.parse().unwrap()),
                "f32" => D::Single(v.parse().unwrap()),
                "f64" => D::Double(v.parse().unwrap()),
                "cy" => D::Currency(v.parse().unwrap()),
                "date" => D::Date(v.parse().unwrap()),
                "bool" => D::Boolean(v.parse().unwrap()),
                _ => panic!("string compared separately"),
            }
        }
    }
}

#[test]
fn native_default_buffers_preserve_scalar_types_sequences_and_parameter_positions() {
    let (mut projects, mut values) = (0, 0);
    for r in cases() {
        let p = hex(r[1]);
        let c = hex(r[2]);
        let primary = decode_runtime_module_heap(&p, CacheContext::B5_WIN64).unwrap();
        let table =
            decode_runtime_companion_cache(&c, CacheContext::B5_WIN64, &primary, 2).unwrap();
        for entry in r[3].split(';') {
            let f = entry.split('~').collect::<Vec<_>>();
            let method = &table.methods[f[0].parse::<usize>().unwrap()];
            let actual = method.optional_defaults(CacheContext::B5_WIN64).unwrap();
            let exp = f[2].split('|').collect::<Vec<_>>();
            assert_eq!(actual.len(), exp.len(), "{} {}", r[0], f[1]);
            let indices = method
                .parameter_types
                .iter()
                .enumerate()
                .filter_map(|(i, t)| t.is_optional().then_some(i))
                .collect::<Vec<_>>();
            assert_eq!(
                actual.iter().map(|v| v.parameter_index).collect::<Vec<_>>(),
                indices
            );
            for (a, e) in actual.iter().zip(exp) {
                if e == "implicit" {
                    assert!(a.encoded.is_none());
                } else {
                    let v = &a.encoded.as_ref().unwrap().value;
                    if let Some(raw) = e.strip_prefix("ansi:") {
                        assert_eq!(v, &D::AnsiString(&hex(raw)));
                    } else {
                        assert_eq!(v, &expected(e), "{} {}", r[0], f[1]);
                    }
                }
                values += 1;
            }
        }
        projects += 1;
    }
    assert_eq!((projects, values), (4, 58));
}

#[test]
fn conversion_keeps_integer_precision_currency_date_missing_and_host_string_decoding() {
    let number = D::LongLong(9007199254740993)
        .to_value::<()>(|_| Err(()))
        .unwrap();
    assert_eq!(number, Value::LongLong(9007199254740993));
    assert_eq!(
        D::Currency(-123456).to_value::<()>(|_| Err(())).unwrap(),
        Value::Currency(-123456)
    );
    assert_eq!(
        D::Date(38020.170208333336)
            .to_value::<()>(|_| Err(()))
            .unwrap(),
        Value::Date(38020.170208333336)
    );
    assert_eq!(
        D::Error(-2147352572).to_value::<()>(|_| Err(())).unwrap(),
        Value::Missing
    );
    let raw = hex("93fa967b8cea");
    let v = D::AnsiString(&raw)
        .to_value(|bytes| {
            assert_eq!(bytes, raw);
            Ok::<_, ()>("日本語".into())
        })
        .unwrap();
    assert_eq!(v, Value::String("日本語".into()));
    assert_eq!(
        D::AnsiString(&raw).to_value(|_| Err::<String, _>("codec unavailable")),
        Err("codec unavailable")
    );
}

#[test]
fn native_alignment_padding_is_not_string_content_or_a_zero_requirement() {
    let bytes = hex("08000300610062ff1100ff7e030007000000");
    let values = decode_companion_default_values(&bytes, CacheContext::B5_WIN64).unwrap();
    assert_eq!(values.len(), 3);
    assert_eq!(values[0].value, D::AnsiString(b"a\0b"));
    assert_eq!(values[0].padding, [255]);
    assert_eq!(values[1].value, D::Byte(255));
    assert_eq!(values[1].padding, [126]);
    assert_eq!(values[2].value, D::Long(7));
    assert_eq!(
        values.iter().map(|v| v.consumed).sum::<usize>(),
        bytes.len()
    );
}

#[test]
fn truncated_values_are_rejected_at_every_incomplete_boundary() {
    for raw in [
        "1100ff00",
        "030005000000",
        "14000100000000002000",
        "080005006361636865ff",
        "090000000000",
        "070026bf58728590e240",
    ] {
        let b = hex(raw);
        for end in 1..b.len() {
            assert!(
                decode_companion_default_values(&b[..end], CacheContext::B5_WIN64).is_err(),
                "{raw} {end}"
            );
        }
    }
    assert!(decode_companion_default_values(&hex("0800ffff"), CacheContext::B5_WIN64).is_err());
    assert!(decode_companion_default_values(&hex("090001000000"), CacheContext::B5_WIN64).is_err());
    assert!(decode_companion_default_values(&hex("ffff"), CacheContext::B5_WIN64).is_err());
}

#[test]
fn encoded_empty_is_not_inferred_from_an_absent_buffer_and_invalid_counts_fail() {
    let r = cases().find(|r| r[0] == "mixed-defaults").unwrap();
    let p = hex(r[1]);
    let c = hex(r[2]);
    let primary = decode_runtime_module_heap(&p, CacheContext::B5_WIN64).unwrap();
    let mut table =
        decode_runtime_companion_cache(&c, CacheContext::B5_WIN64, &primary, 2).unwrap();
    let mixed = table.methods[1]
        .optional_defaults(CacheContext::B5_WIN64)
        .unwrap();
    assert_eq!(mixed[0].encoded.as_ref().unwrap().value, D::Empty);
    assert_eq!(mixed[2].encoded.as_ref().unwrap().value, D::Empty);
    assert!(table.methods[3]
        .optional_defaults(CacheContext::B5_WIN64)
        .unwrap()
        .iter()
        .all(|v| v.encoded.is_none()));
    table.methods[1].parameter_types.pop();
    assert!(table.methods[1]
        .optional_defaults(CacheContext::B5_WIN64)
        .is_err());
    assert!(decode_companion_default_values(
        &[],
        CacheContext {
            version: 0xb4,
            sys_kind: 3
        }
    )
    .is_err());
    assert!(decode_companion_default_values(
        &[],
        CacheContext {
            version: 0xb5,
            sys_kind: 1
        }
    )
    .is_err());
}
