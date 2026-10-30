// This Source Code Form is subject to the terms of the Mozilla Public
// License, v. 2.0. If a copy of the MPL was not distributed with this
// file, You can obtain one at https://mozilla.org/MPL/2.0/.

use oxivba_core::native_cache::{
    decode_module_numeric_assignment, decode_runtime_numeric_assignment, CacheContext,
    DecodeErrorKind, Location,
};

fn hex(text: &str) -> Vec<u8> {
    text.as_bytes()
        .chunks_exact(2)
        .map(|pair| u8::from_str_radix(std::str::from_utf8(pair).unwrap(), 16).unwrap())
        .collect()
}
fn cases() -> impl Iterator<Item = Vec<&'static str>> {
    include_str!("../../../tests/fixtures/vba/native_cache_byref.tsv")
        .lines()
        .filter(|line| !line.starts_with('#') && !line.starts_with("id\t"))
        .map(|line| line.split('\t').collect())
}

#[test]
fn byref_numeric_spans_preserve_native_source_execution_differences() {
    let (mut count, mut differences) = (0, 0);
    for row in cases() {
        let source =
            decode_module_numeric_assignment(&hex(row[1]), CacheContext::B5_WIN64).unwrap();
        let runtime =
            decode_runtime_numeric_assignment(&hex(row[2]), CacheContext::B5_WIN64).unwrap();
        let input = row[3].parse::<f64>().unwrap();
        let shown = source.evaluate_double(input).unwrap();
        let executed = runtime.evaluate_double(input).unwrap();
        assert_eq!(shown, row[4].parse::<f64>().unwrap(), "{} source", row[0]);
        assert_eq!(
            executed,
            row[5].parse::<f64>().unwrap(),
            "{} runtime",
            row[0]
        );
        assert_eq!(runtime.input, Location::IndirectFrameOffset(8));
        assert_eq!(runtime.output, Location::FrameOffset(-8));
        count += 1;
        differences += usize::from(shown != executed);
    }
    assert_eq!((count, differences), (9, 2));
}

#[test]
fn indirect_loads_retain_reference_binding_and_signed_frame_locations() {
    let row = cases().next().unwrap();
    let mut bytes = hex(row[2]);
    let indirect = decode_runtime_numeric_assignment(&bytes, CacheContext::B5_WIN64).unwrap();
    bytes[..2].copy_from_slice(&0x29du16.to_le_bytes());
    let direct = decode_runtime_numeric_assignment(&bytes, CacheContext::B5_WIN64).unwrap();
    assert_eq!(indirect.input, Location::IndirectFrameOffset(8));
    assert_eq!(direct.input, Location::FrameOffset(8));
    // Evaluation receives the value after the caller resolves its binding.
    assert_eq!(indirect.evaluate_double(7.0), direct.evaluate_double(7.0));
    bytes[..2].copy_from_slice(&0x2edu16.to_le_bytes());
    bytes[2..6].copy_from_slice(&(-24i32).to_le_bytes());
    assert_eq!(
        decode_runtime_numeric_assignment(&bytes, CacheContext::B5_WIN64)
            .unwrap()
            .input,
        Location::IndirectFrameOffset(-24)
    );
}

#[test]
fn indirect_load_spans_reject_truncation_unknown_loads_and_wrong_contexts() {
    for row in cases() {
        let mut bytes = hex(row[2]);
        let end = bytes.len();
        for length in 0..end {
            assert!(
                decode_runtime_numeric_assignment(&bytes[..length], CacheContext::B5_WIN64)
                    .is_err()
            );
        }
        bytes.extend_from_slice(&[0xff; 16]);
        assert_eq!(
            decode_runtime_numeric_assignment(&bytes, CacheContext::B5_WIN64)
                .unwrap()
                .consumed,
            end
        );
        bytes[..2].copy_from_slice(&0x29eu16.to_le_bytes());
        assert_eq!(
            decode_runtime_numeric_assignment(&bytes, CacheContext::B5_WIN64)
                .unwrap_err()
                .kind,
            DecodeErrorKind::UnsupportedOpcode(0x29e)
        );
        bytes[..2].copy_from_slice(&0x2edu16.to_le_bytes());
        assert!(decode_runtime_numeric_assignment(
            &bytes,
            CacheContext {
                version: 180,
                sys_kind: 3
            }
        )
        .is_err());
        assert!(decode_runtime_numeric_assignment(
            &bytes,
            CacheContext {
                version: 181,
                sys_kind: 1
            }
        )
        .is_err());
    }
}
