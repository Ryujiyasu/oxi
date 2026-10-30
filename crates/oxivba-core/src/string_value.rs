// This Source Code Form is subject to the terms of the Mozilla Public
// License, v. 2.0. If a copy of the MPL was not distributed with this
// file, You can obtain one at https://mozilla.org/MPL/2.0/.

//! Lossless VBA string values. UTF-16 code units, including lone surrogates,
//! are data; a Rust Unicode scalar string is used only when it is lossless.
use crate::runtime::Value;
use std::borrow::Cow;

impl Value {
    pub fn from_utf16(units: Vec<u16>) -> Self {
        match String::from_utf16(&units) {
            Ok(text) => Self::String(text),
            Err(_) => Self::Utf16(units),
        }
    }
    pub fn is_string(&self) -> bool {
        matches!(self, Self::String(_) | Self::Utf16(_) | Self::NullString)
    }
    pub fn string_units(&self) -> Option<Cow<'_, [u16]>> {
        match self {
            Self::String(s) => Some(Cow::Owned(s.encode_utf16().collect())),
            Self::Utf16(s) => Some(Cow::Borrowed(s)),
            Self::NullString => Some(Cow::Borrowed(&[])),
            _ => None,
        }
    }
    /// Conversion for interfaces requiring Unicode scalar text. Invalid units
    /// are reported, never replaced. VBA string operations use string_units.
    pub fn unicode_text(&self) -> Option<Result<Cow<'_, str>, std::string::FromUtf16Error>> {
        match self {
            Self::String(s) => Some(Ok(Cow::Borrowed(s))),
            Self::NullString => Some(Ok(Cow::Borrowed(""))),
            Self::Utf16(s) => Some(String::from_utf16(s).map(Cow::Owned)),
            _ => None,
        }
    }
}
impl PartialEq for Value {
    fn eq(&self, other: &Self) -> bool {
        if let (Self::String(a), Self::String(b)) = (self, other) {
            return a == b;
        }
        if let (Some(a), Some(b)) = (self.string_units(), other.string_units()) {
            return a == b;
        }
        use Value::*;
        match (self, other) {
            (Empty, Empty) | (Missing, Missing) | (Nothing, Nothing) | (Null, Null) => true,
            (Boolean(a), Boolean(b)) => a == b,
            (Int16(a), Int16(b)) => a == b,
            (Integer(a), Integer(b)) => a == b,
            (LongLong(a), LongLong(b)) => a == b,
            (Byte(a), Byte(b)) => a == b,
            (Single(a), Single(b)) => a == b,
            (Double(a), Double(b)) => a == b,
            (Currency(a), Currency(b)) => a == b,
            (Decimal(a), Decimal(b)) => a == b,
            (Date(a), Date(b)) => a == b,
            (Error(a), Error(b)) => a == b,
            (Array(a), Array(b)) => a == b,
            (Record(a), Record(b)) => a == b,
            (Object(a), Object(b)) => a == b,
            _ => false,
        }
    }
}
