// SPDX-License-Identifier: MIT OR Apache-2.0

use super::*;
use serde_json::json;

#[test]
fn result_scalar_descriptors_preserve_vba_subtypes_when_reused_as_arguments() {
    let cases = [
        (Value::Empty, "empty"),
        (Value::Missing, "missing"),
        (Value::Nothing, "nothing"),
        (Value::Null, "null"),
        (Value::NullString, "null_string"),
        (Value::String(String::new()), "string"),
        (Value::Boolean(false), "boolean"),
        (Value::Boolean(true), "boolean"),
        (Value::Byte(0), "byte"),
        (Value::Byte(u8::MAX), "byte"),
        (Value::Int16(i16::MIN), "integer"),
        (Value::Int16(i16::MAX), "integer"),
        (Value::Integer(i64::from(i32::MIN)), "long"),
        (Value::Integer(i64::from(i32::MAX)), "long"),
        (Value::LongLong(i64::MIN), "long_long"),
        (Value::LongLong(i64::MAX), "long_long"),
        (Value::Single(1.234_567_9), "single"),
        (Value::Single(f32::from_bits(1)), "single"),
        (Value::Single(f32::MAX), "single"),
        (Value::Double(f64::MAX), "double"),
        (Value::Date(0.0), "date"),
        (Value::Date(-0.5), "date"),
        (Value::Date(43_890.502_835_648_15), "date"),
        (Value::Currency(i64::MIN), "currency"),
        (Value::Currency(i64::MAX), "currency"),
        (Value::Error(2007), "error"),
    ];
    for (value, tag) in cases {
        let descriptor = serde_json::to_value(OutputValue::from(value.clone())).unwrap();
        assert_eq!(descriptor["type"], tag, "{value:?}");
        if tag == "long_long" {
            assert!(descriptor["value"].is_string());
        }
        let restored = decode_typed_scalar(&descriptor).unwrap();
        assert_eq!(
            std::mem::discriminant(&restored),
            std::mem::discriminant(&value)
        );
        assert_eq!(restored, value);
    }
}

#[test]
fn result_float_descriptors_preserve_signed_zero_and_nonfinite_payload_bits() {
    for bits in [
        0u32,
        0x8000_0000,
        1,
        0x7f80_0000,
        0xff80_0000,
        0x7fc0_0042,
        0xffc0_0317,
    ] {
        let descriptor =
            serde_json::to_value(OutputValue::from(Value::Single(f32::from_bits(bits)))).unwrap();
        if bits == 0x8000_0000 {
            assert_eq!(descriptor["value"], json!({"bits":"80000000"}));
        }
        let Value::Single(restored) = decode_typed_scalar(&descriptor).unwrap() else {
            panic!("lost Single")
        };
        assert_eq!(restored.to_bits(), bits);
    }
    for bits in [
        0u64,
        0x8000_0000_0000_0000,
        1,
        0x7ff0_0000_0000_0000,
        0xfff0_0000_0000_0000,
        0x7ff8_0000_0000_0042,
        0xfff8_0000_0000_0317,
    ] {
        for value in [
            Value::Double(f64::from_bits(bits)),
            Value::Date(f64::from_bits(bits)),
        ] {
            let descriptor = serde_json::to_value(OutputValue::from(value.clone())).unwrap();
            if bits == 0x8000_0000_0000_0000 {
                assert_eq!(descriptor["value"], json!({"bits":"8000000000000000"}));
            }
            let restored = decode_typed_scalar(&descriptor).unwrap();
            assert_eq!(
                std::mem::discriminant(&restored),
                std::mem::discriminant(&value)
            );
            let (Value::Double(restored) | Value::Date(restored)) = restored else {
                panic!("lost floating type")
            };
            assert_eq!(restored.to_bits(), bits);
        }
    }
    for descriptor in [
        json!({"type":"single","value":{"bits":"7ff8000000000042"}}),
        json!({"type":"double","value":{"bits":"7fc00042"}}),
        json!({"type":"date","value":{"bits":"0x00000000000000"}}),
        json!({"type":"double","value":{"bits":"0000000000000000","extra":0}}),
        json!({"type":"double","value":{"bits":0}}),
    ] {
        assert!(decode_typed_scalar(&descriptor).is_err(), "{descriptor}");
    }
}

#[test]
fn result_arrays_retain_element_type_bounds_and_resize_contract() {
    for (default, values, tag) in [
        (
            Value::Empty,
            vec![Value::Byte(7), Value::Date(43_890.0), Value::NullString],
            "variant",
        ),
        (
            Value::Int16(0),
            vec![Value::Int16(-7), Value::Int16(0), Value::Int16(13)],
            "integer",
        ),
        (
            Value::Integer(0),
            vec![
                Value::Integer(i32::MIN.into()),
                Value::Integer(0),
                Value::Integer(i32::MAX.into()),
            ],
            "long",
        ),
        (
            Value::LongLong(0),
            vec![
                Value::LongLong(i64::MIN),
                Value::LongLong(0),
                Value::LongLong(i64::MAX),
            ],
            "long_long",
        ),
        (
            Value::Single(0.0),
            vec![Value::Single(1.25), Value::Single(0.0), Value::Single(-2.5)],
            "single",
        ),
        (
            Value::Date(0.0),
            vec![Value::Date(-0.5), Value::Date(0.0), Value::Date(43_890.5)],
            "date",
        ),
        (
            Value::NullString,
            vec![
                Value::NullString,
                Value::String("one".into()),
                Value::String(String::new()),
            ],
            "string",
        ),
        (
            Value::Nothing,
            vec![Value::Nothing, Value::Nothing, Value::Nothing],
            "object",
        ),
    ] {
        for resizable in [false, true] {
            let array = ArrayValue {
                dimensions: vec![ArrayDimension {
                    lower_bound: -2,
                    length: 3,
                }],
                values: values.clone(),
                element_default: Box::new(default.clone()),
                resizable,
            };
            let descriptor =
                serde_json::to_value(OutputValue::from(Value::Array(array.clone()))).unwrap();
            let payload = &descriptor["value"];
            assert_eq!(payload["element_type"], tag);
            assert_eq!(payload["lower_bound"], -2);
            assert_eq!(
                payload["dimensions"],
                json!([{"lower_bound":-2,"length":3}])
            );
            assert_eq!(payload["resizable"], resizable);
            let metadata = json!({"element_type":payload["element_type"],"dimensions":payload["dimensions"],"resizable":payload["resizable"]});
            let schema = decode_vba_input_array_schema(&metadata, 3).unwrap();
            let values = payload["values"]
                .as_array()
                .unwrap()
                .iter()
                .map(decode_typed_scalar)
                .collect::<Result<Vec<_>, _>>()
                .unwrap();
            assert_eq!(
                finish_vba_input_array(schema, values).unwrap(),
                Value::Array(array)
            );
        }
    }
    for dimensions in [
        vec![],
        vec![ArrayDimension {
            lower_bound: 0,
            length: 0,
        }],
    ] {
        let descriptor = serde_json::to_value(OutputValue::from(Value::Array(ArrayValue {
            dimensions: dimensions.clone(),
            values: vec![],
            element_default: Box::new(Value::Int16(0)),
            resizable: true,
        })))
        .unwrap();
        assert_eq!(descriptor["value"]["element_type"], "integer");
        assert_eq!(
            descriptor["value"]["dimensions"].as_array().unwrap().len(),
            dimensions.len()
        );
    }
}
