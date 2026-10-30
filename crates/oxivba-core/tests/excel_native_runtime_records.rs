// This Source Code Form is subject to the terms of the Mozilla Public
// License, v. 2.0. If a copy of the MPL was not distributed with this
// file, You can obtain one at https://mozilla.org/MPL/2.0/.

use oxivba_core::native_cache::{
    decode_runtime_numeric_assignment, decode_runtime_record, CacheContext, DecodeErrorKind,
    Location, RuntimeRecord,
};

fn hex(text: &str) -> Vec<u8> {
    text.as_bytes()
        .chunks_exact(2)
        .map(|pair| u8::from_str_radix(std::str::from_utf8(pair).unwrap(), 16).unwrap())
        .collect()
}
fn cases() -> impl Iterator<Item = Vec<&'static str>> {
    include_str!("../../../tests/fixtures/vba/native_cache_runtime_records.tsv")
        .lines()
        .filter(|line| !line.starts_with('#') && !line.starts_with("id\t"))
        .map(|line| line.split('\t').collect())
}

#[test]
fn native_runtime_records_preserve_interleaved_references_and_numeric_code() {
    let (mut images, mut references, mut numeric) = (0, 0, 0);
    for row in cases() {
        let bytes = hex(row[1]);
        let mut at = 0;
        for shape in row[2].split(';') {
            let fields: Vec<_> = shape.split(',').collect();
            let decoded = decode_runtime_record(&bytes[at..], CacheContext::B5_WIN64).unwrap();
            match decoded.record {
                RuntimeRecord::Image { code, frame_info } => {
                    assert_eq!(fields[0], "I");
                    assert_eq!(
                        code.len() + frame_info.len(),
                        fields[1].parse::<usize>().unwrap()
                    );
                    assert_eq!(code.len(), fields[2].parse::<usize>().unwrap());
                    assert!(frame_info.iter().all(|b| *b == 0));
                    assert_eq!(decoded.consumed, 16 + code.len() + frame_info.len());
                    if row[3] != "-" {
                        let assignment =
                            decode_runtime_numeric_assignment(&code[12..], CacheContext::B5_WIN64)
                                .unwrap();
                        assert_eq!(
                            assignment.evaluate_double(7.0).unwrap(),
                            row[3].parse::<f64>().unwrap()
                        );
                        assert_eq!(
                            assignment.input,
                            if row[0] == "numeric-byref" {
                                Location::IndirectFrameOffset(8)
                            } else {
                                Location::FrameOffset(8)
                            }
                        );
                        numeric += 1;
                    }
                    images += 1;
                }
                RuntimeRecord::Reference { handle, offset } => {
                    assert_eq!(fields[0], "R");
                    assert_eq!(handle, fields[1].parse::<u64>().unwrap());
                    assert_eq!(offset, fields[2].parse::<u32>().unwrap());
                    assert_eq!(decoded.consumed, 12);
                    references += 1;
                }
            }
            at += decoded.consumed;
        }
        assert_eq!(at, bytes.len(), "{}", row[0]);
    }
    assert_eq!((images, references, numeric), (11, 25, 2));
}

#[test]
fn seventeen_record_type_references_do_not_shift_the_following_function() {
    let row = cases().find(|r| r[0] == "record-layout").unwrap();
    let bytes = hex(row[1]);
    let mut at = 0;
    let mut reference_counts = Vec::new();
    while at < bytes.len() {
        let decoded = decode_runtime_record(&bytes[at..], CacheContext::B5_WIN64).unwrap();
        match decoded.record {
            RuntimeRecord::Image { .. } => reference_counts.push(0),
            RuntimeRecord::Reference { .. } => *reference_counts.last_mut().unwrap() += 1,
        }
        at += decoded.consumed;
    }
    // Native source order: Driver, OneType, SameType, TwoTypes, ManyTypes,
    // NestedType, RecordArray. The same type in two locals still has two refs.
    assert_eq!(reference_counts, [0, 1, 2, 2, 17, 1, 1]);
}

#[test]
fn runtime_records_are_bounded_and_opaque_payload_is_not_scanned_for_tags() {
    for row in cases() {
        let bytes = hex(row[1]);
        let mut at = 0;
        while at < bytes.len() {
            let record = decode_runtime_record(&bytes[at..], CacheContext::B5_WIN64).unwrap();
            for length in 0..record.consumed {
                assert!(
                    decode_runtime_record(&bytes[at..at + length], CacheContext::B5_WIN64).is_err(),
                    "{} {at} {length}",
                    row[0]
                );
            }
            at += record.consumed;
        }
    }
    let row = cases().find(|r| r[0] == "record-layout").unwrap();
    let mut bytes = hex(row[1]);
    let record = decode_runtime_record(&bytes, CacheContext::B5_WIN64).unwrap();
    let RuntimeRecord::Image { code, .. } = record.record else {
        panic!("image")
    };
    let payload = 16 + code.len();
    let consumed = record.consumed;
    // A pointer-looking word inside frame data is opaque, not a record boundary.
    bytes[payload..payload + 8].copy_from_slice(&0x1231u64.to_le_bytes());
    assert_eq!(
        decode_runtime_record(&bytes, CacheContext::B5_WIN64)
            .unwrap()
            .consumed,
        consumed
    );
    assert!(matches!(
        decode_runtime_record(&bytes[consumed..], CacheContext::B5_WIN64)
            .unwrap()
            .record,
        RuntimeRecord::Image { .. }
    ));
}

#[test]
fn invalid_runtime_tags_sizes_and_contexts_fail_without_resolving_handles() {
    for tag in [3, 15, u64::MAX] {
        let bytes = [tag.to_le_bytes().as_slice(), &[0; 16]].concat();
        assert_eq!(
            decode_runtime_record(&bytes, CacheContext::B5_WIN64)
                .unwrap_err()
                .kind,
            DecodeErrorKind::InvalidRuntimeRecordTag(tag)
        );
    }
    for (size, code) in [
        (0u64, 2u64),
        (2, 1),
        (u64::MAX - 1, 0),
        (u64::MAX - 1, u64::MAX - 1),
    ] {
        let bytes = [size.to_le_bytes(), code.to_le_bytes()].concat();
        assert!(decode_runtime_record(&bytes, CacheContext::B5_WIN64).is_err());
    }
    let handle = u64::MAX - 14;
    let bytes = [
        handle.to_le_bytes().as_slice(),
        u32::MAX.to_le_bytes().as_slice(),
    ]
    .concat();
    assert_eq!(
        decode_runtime_record(&bytes, CacheContext::B5_WIN64)
            .unwrap()
            .record,
        RuntimeRecord::Reference {
            handle,
            offset: u32::MAX
        }
    );
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
        assert!(decode_runtime_record(&bytes, context).is_err());
    }
    let empty = decode_runtime_record(&[0; 16], CacheContext::B5_WIN64).unwrap();
    assert_eq!(
        empty.record,
        RuntimeRecord::Image {
            code: &[],
            frame_info: &[]
        }
    );
    assert_eq!(empty.consumed, 16);
}
