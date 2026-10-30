// This Source Code Form is subject to the terms of the Mozilla Public
// License, v. 2.0. If a copy of the MPL was not distributed with this
// file, You can obtain one at https://mozilla.org/MPL/2.0/.

use oxivba_core::native_cache::{
    decode_module_numeric_assignment, decode_runtime_numeric_assignment, CacheContext,
    DecodeErrorKind, Location, Representation, Scalar,
};
use oxivba_core::{RuntimeErrorKind, Value};

fn hex(text: &str) -> Vec<u8> {
    assert_eq!(text.len() % 2, 0);
    text.as_bytes()
        .chunks_exact(2)
        .map(|pair| u8::from_str_radix(std::str::from_utf8(pair).unwrap(), 16).unwrap())
        .collect()
}

fn cases() -> impl Iterator<Item = Vec<&'static str>> {
    include_str!("../../../tests/fixtures/vba/native_cache_numeric.tsv")
        .lines()
        .filter(|line| !line.starts_with('#') && !line.starts_with("id\t"))
        .map(|line| line.split('\t').collect())
}

#[test]
fn cached_numeric_instructions_match_native_excel_observations() {
    let mut count = 0;
    for fields in cases() {
        let source = hex(fields[4]);
        let runtime = hex(fields[5]);
        let expected: f64 = fields[3].parse().unwrap();
        let source_ir = decode_module_numeric_assignment(&source, CacheContext::B5_WIN64).unwrap();
        let runtime_ir =
            decode_runtime_numeric_assignment(&runtime, CacheContext::B5_WIN64).unwrap();
        assert_eq!(
            source_ir.evaluate_double(7.0).unwrap(),
            expected,
            "{} module",
            fields[0]
        );
        assert_eq!(
            runtime_ir.evaluate_double(7.0).unwrap(),
            expected,
            "{} runtime",
            fields[0]
        );
        assert_eq!(source_ir.representation, Representation::ModulePcode);
        assert_eq!(runtime_ir.representation, Representation::Runtime);
        assert!(matches!(source_ir.input, Location::Symbol(_)));
        assert_eq!(runtime_ir.input, Location::FrameOffset(8));
        assert_eq!(runtime_ir.output, Location::FrameOffset(-8));
        assert_eq!(source_ir.consumed, source.len());
        assert_eq!(runtime_ir.consumed, runtime.len());
        count += 1;
    }
    assert_eq!(count, 48);
}

#[test]
fn partial_cache_instructions_fail_without_panicking_or_consuming_the_next_span() {
    for fields in cases() {
        let source = hex(fields[4]);
        let mut runtime = hex(fields[5]);
        for end in 0..source.len() {
            assert!(
                decode_module_numeric_assignment(&source[..end], CacheContext::B5_WIN64).is_err(),
                "{} source prefix{end}",
                fields[0]
            );
        }
        for end in 0..runtime.len() {
            assert!(
                decode_runtime_numeric_assignment(&runtime[..end], CacheContext::B5_WIN64).is_err(),
                "{} runtime prefix{end}",
                fields[0]
            );
        }
        let end = runtime.len();
        runtime.extend_from_slice(&[0xff; 16]);
        assert_eq!(
            decode_runtime_numeric_assignment(&runtime, CacheContext::B5_WIN64)
                .unwrap()
                .consumed,
            end
        );
    }
    let source = hex(cases().next().unwrap()[4]);
    for context in [
        CacheContext {
            version: 0xb6,
            sys_kind: 3,
        },
        CacheContext {
            version: 0xb5,
            sys_kind: 1,
        },
    ] {
        assert!(matches!(
            decode_module_numeric_assignment(&source, context)
                .unwrap_err()
                .kind,
            DecodeErrorKind::UnsupportedContext(_)
        ));
    }
    let mut unknown = source;
    unknown[4..6].copy_from_slice(&0xffffu16.to_le_bytes());
    assert!(matches!(
        decode_module_numeric_assignment(&unknown, CacheContext::B5_WIN64)
            .unwrap_err()
            .kind,
        DecodeErrorKind::UnsupportedOpcode(0xffff)
    ));
}

#[test]
fn currency_payloads_keep_all_bits_and_use_the_existing_checked_negation() {
    let make = |coefficient: i64, negate: bool| {
        let mut bytes = hex("20000100a900");
        bytes.extend_from_slice(&coefficient.to_le_bytes());
        if negate {
            bytes.extend_from_slice(&0x16u16.to_le_bytes());
        }
        bytes.extend_from_slice(&hex("0f0027000200"));
        bytes
    };
    let exact =
        decode_module_numeric_assignment(&make(i64::MAX, false), CacheContext::B5_WIN64).unwrap();
    assert_eq!(exact.constant.scalar, Scalar::Currency(i64::MAX));
    assert_eq!(exact.constant.scalar.value(), Value::Currency(i64::MAX));
    let overflow =
        decode_module_numeric_assignment(&make(i64::MIN, true), CacheContext::B5_WIN64).unwrap();
    let error = overflow.evaluate_double(7.0).unwrap_err();
    assert_eq!(error.kind, RuntimeErrorKind::Overflow);
    assert_eq!(error.line, None);
}

#[test]
fn decoded_division_uses_vba_errors_instead_of_host_floating_point_results() {
    let row = cases().find(|fields| fields[1] == "0").unwrap();
    let mut source = hex(row[4]);
    let at = source.len() - 6;
    source[at..at + 2].copy_from_slice(&0x10u16.to_le_bytes());
    let assignment = decode_module_numeric_assignment(&source, CacheContext::B5_WIN64).unwrap();
    let error = assignment.evaluate_double(7.0).unwrap_err();
    assert_eq!(error.kind, RuntimeErrorKind::DivisionByZero);
    assert_eq!(error.line, None);
}
