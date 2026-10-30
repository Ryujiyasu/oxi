// This Source Code Form is subject to the terms of the Mozilla Public
// License, v. 2.0. If a copy of the MPL was not distributed with this
// file, You can obtain one at https://mozilla.org/MPL/2.0/.

//! Encoded Automation defaults. These describe companion invocation metadata;
//! compiled internal VBA calls can supply different defaults for the same entry.

use super::{CacheContext, DecodeError, DecodeErrorKind, Reader, RuntimeCompanionMethod};
use crate::runtime::Value;

#[derive(Debug, Clone, PartialEq)]
pub enum CompanionDefaultValue<'a> {
    Empty,
    Null,
    Byte(u8),
    Integer(i16),
    Long(i32),
    LongLong(i64),
    Single(f32),
    Double(f64),
    Currency(i64),
    Date(f64),
    Boolean(i16),
    AnsiString(&'a [u8]),
    Error(i32),
    Nothing,
}

impl CompanionDefaultValue<'_> {
    /// ANSI decoding belongs to the target host. Currency, date, integer width
    /// and Missing remain typed; no decimal-text or f64 conversion is used.
    pub fn to_value<E>(
        &self,
        decode_ansi: impl FnOnce(&[u8]) -> Result<String, E>,
    ) -> Result<Value, E> {
        Ok(match self {
            Self::Empty => Value::Empty,
            Self::Null => Value::Null,
            Self::Byte(v) => Value::Byte(*v),
            Self::Integer(v) => Value::Int16(*v),
            Self::Long(v) => Value::Integer(i64::from(*v)),
            Self::LongLong(v) => Value::LongLong(*v),
            Self::Single(v) => Value::Single(*v),
            Self::Double(v) => Value::Double(*v),
            Self::Currency(v) => Value::Currency(*v),
            Self::Date(v) => Value::Date(*v),
            Self::Boolean(v) => Value::Boolean(*v != 0),
            Self::AnsiString(bytes) => Value::String(decode_ansi(bytes)?),
            Self::Error(-2147352572) => Value::Missing, // DISP_E_PARAMNOTFOUND
            Self::Error(code) => Value::Error(i64::from(*code)),
            Self::Nothing => Value::Nothing,
        })
    }
}

#[derive(Debug, Clone, PartialEq)]
pub struct DecodedCompanionDefault<'a> {
    pub offset: usize,
    pub value: CompanionDefaultValue<'a>,
    /// Native alignment bytes may be uninitialized; they are not string data.
    pub padding: &'a [u8],
    pub consumed: usize,
}

#[derive(Debug, Clone, PartialEq)]
pub struct RuntimeOptionalDefault<'a> {
    pub parameter_index: usize,
    /// None means the metadata has no encoded default buffer. Its implicit
    /// value depends on the parameter type and invocation context.
    pub encoded: Option<DecodedCompanionDefault<'a>>,
}

impl<'a> RuntimeCompanionMethod<'a> {
    pub fn optional_defaults(
        &self,
        context: CacheContext,
    ) -> Result<Vec<RuntimeOptionalDefault<'a>>, DecodeError> {
        context.validate()?;
        let indices = self
            .parameter_types
            .iter()
            .enumerate()
            .filter_map(|(i, t)| t.is_optional().then_some(i))
            .collect::<Vec<_>>();
        if self.default_data.is_empty() {
            return Ok(indices
                .into_iter()
                .map(|parameter_index| RuntimeOptionalDefault {
                    parameter_index,
                    encoded: None,
                })
                .collect());
        }
        let values = decode_companion_default_values(self.default_data, context)?;
        if values.len() != indices.len() {
            return Err(DecodeError {
                offset: self.default_data.len(),
                kind: DecodeErrorKind::InvalidRuntimeHeap,
            });
        }
        Ok(indices
            .into_iter()
            .zip(values)
            .map(|(parameter_index, value)| RuntimeOptionalDefault {
                parameter_index,
                encoded: Some(value),
            })
            .collect())
    }
}

/// Decode a complete companion default-value buffer. Each value begins with an
/// Automation VARTYPE; byte and ANSI string payloads align to two bytes. This is
/// distinct from the companion's internal type words and from global BSTRs.
pub fn decode_companion_default_values(
    bytes: &[u8],
    context: CacheContext,
) -> Result<Vec<DecodedCompanionDefault<'_>>, DecodeError> {
    context.validate()?;
    let mut r = Reader::new(bytes, 0);
    let mut values = Vec::new();
    while r.at < bytes.len() {
        let offset = r.at;
        let tag = r.word()?;
        let value = match tag {
            0 => CompanionDefaultValue::Empty,
            1 => CompanionDefaultValue::Null,
            2 => CompanionDefaultValue::Integer(i16::from_le_bytes(r.take()?)),
            3 => CompanionDefaultValue::Long(i32::from_le_bytes(r.take()?)),
            4 => CompanionDefaultValue::Single(f32::from_le_bytes(r.take()?)),
            5 => CompanionDefaultValue::Double(f64::from_le_bytes(r.take()?)),
            6 => CompanionDefaultValue::Currency(i64::from_le_bytes(r.take()?)),
            7 => CompanionDefaultValue::Date(f64::from_le_bytes(r.take()?)),
            8 => {
                let n = usize::from(r.word()?);
                CompanionDefaultValue::AnsiString(r.slice(n)?)
            }
            9 => {
                if u32::from_le_bytes(r.take()?) != 0 {
                    return Err(r.error(DecodeErrorKind::InvalidRuntimeHeap));
                }
                CompanionDefaultValue::Nothing
            }
            10 => CompanionDefaultValue::Error(i32::from_le_bytes(r.take()?)),
            11 => CompanionDefaultValue::Boolean(i16::from_le_bytes(r.take()?)),
            17 => CompanionDefaultValue::Byte(r.take::<1>()?[0]),
            20 => CompanionDefaultValue::LongLong(i64::from_le_bytes(r.take()?)),
            tag => return Err(r.error(DecodeErrorKind::UnsupportedCompanionDefault(tag))),
        };
        let padding = r.slice(r.at & 1)?;
        values.push(DecodedCompanionDefault {
            offset,
            value,
            padding,
            consumed: r.at - offset,
        });
    }
    Ok(values)
}
