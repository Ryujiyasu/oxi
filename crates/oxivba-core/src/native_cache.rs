// This Source Code Form is subject to the terms of the Mozilla Public
// License, v. 2.0. If a copy of the MPL was not distributed with this
// file, You can obtain one at https://mozilla.org/MPL/2.0/.

//! Version-checked readers for numeric VBA instructions and cache metadata.
//!
//! These readers accept instruction, module-cache or project-cache bytes. They do not
//! open compound files or validate a complete executable cache. Module frames expose
//! procedure declarations; runtime entry mapping and type resolution remain
//! separate. Module p-code and runtime cache instructions are distinct inputs.
//! Unsupported instructions return errors rather than falling back to source.

use crate::ast::{BinaryOp, UnaryOp};
use crate::runtime::{self, RuntimeError, RuntimeErrorKind, Value};

mod identifiers;
pub use identifiers::{
    decode_identifier_table, CachedIdentifier, CachedUnindexedIdentifier, IdentifierTable,
};
mod module;
pub use module::{
    decode_module_cache, CachedLine, CachedParameter, CachedProcedure, CachedType, ModuleCache,
};
mod project;
pub use project::{
    decode_project_cache, CachedProjectModule, CachedProjectReference, ProjectCache,
};
mod records;
pub use records::{decode_runtime_record, DecodedRuntimeRecord, RuntimeRecord};
mod runtime_module;
pub use runtime_module::{
    decode_runtime_module_cache, RuntimeBindingSlots, RuntimeCacheLink, RuntimeClassBinding,
    RuntimeHostEvent, RuntimeMethodImage, RuntimeModuleCache, RuntimeRelocation,
};
mod module_heap;
pub use module_heap::{
    decode_runtime_module_heap, RuntimeHeapObject, RuntimeHeapObjectKind, RuntimeHeapSpan,
    RuntimeModuleHeap,
};
mod project_heap;
pub use project_heap::{
    decode_runtime_project_heap, RuntimeEncodedName, RuntimeNameEncoding, RuntimeProjectHeap,
    RuntimeProjectModule, RuntimeProjectModuleState, RuntimeProjectObject,
    RuntimeProjectObjectKind, RuntimeProjectReferenceEntry, RuntimeReadyModule,
};
mod companion;
pub use companion::{
    decode_runtime_companion_cache, decode_runtime_invocation_table, RuntimeCompanionCache,
    RuntimeCompanionField, RuntimeCompanionMethod, RuntimeCompanionType, RuntimePositionalArity,
};
mod defaults;
pub use defaults::{
    decode_companion_default_values, CompanionDefaultValue, DecodedCompanionDefault,
    RuntimeOptionalDefault,
};
mod frame;
pub use frame::{
    decode_runtime_frame_layout, RuntimeCleanupSlot, RuntimeCleanupTable, RuntimeFrameLayout,
};
mod scalar_vm;
pub use scalar_vm::{
    execute_runtime_scalar_method, RuntimeScalarArgument, RuntimeScalarError,
    RuntimeScalarReference, RuntimeScalarType,
};
mod stream_directory;
pub use stream_directory::{
    decode_runtime_stream_directory, RuntimeStreamDirectory, RuntimeStreamKind, RuntimeStreamSlot,
};

/// Input-cache version and platform, independent of the decoder's host ABI.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct CacheContext {
    pub version: u16,
    pub sys_kind: u32,
}

impl CacheContext {
    /// The cache context covered by the native numeric conformance corpus.
    pub const B5_WIN64: Self = Self {
        version: 0xb5,
        sys_kind: 3,
    };

    fn validate(self) -> Result<(), DecodeError> {
        if self == Self::B5_WIN64 {
            Ok(())
        } else {
            Err(DecodeError {
                offset: 0,
                kind: DecodeErrorKind::UnsupportedContext(self),
            })
        }
    }
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct DecodeError {
    pub offset: usize,
    pub kind: DecodeErrorKind,
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub enum DecodeErrorKind {
    UnsupportedContext(CacheContext),
    Truncated {
        needed: usize,
        remaining: usize,
    },
    UnsupportedOpcode(u16),
    UnexpectedOpcode {
        expected: u16,
        actual: u16,
    },
    TrailingInstructions,
    InvalidIdentifierCounts {
        total: u16,
        named: u16,
        reserved: u16,
    },
    EmptyIdentifier,
    IdentifierSymbolOverflow,
    InvalidModuleCache,
    InvalidDeclarationRecord,
    CyclicParameterChain,
    InvalidProjectCache,
    InvalidRuntimeRecordTag(u64),
    InvalidRuntimeModule,
    InvalidRuntimeDirectory,
    InvalidRuntimeHeap,
    InvalidRuntimeFrame,
    UnsupportedRuntimeHeapObject(u8),
    UnsupportedCompanionDefault(u16),
    UnsupportedRuntimeStreamKind(u16),
    InvalidRuntimeImageSizes {
        image_size: u64,
        code_size: u64,
    },
}

impl std::fmt::Display for DecodeError {
    fn fmt(&self, formatter: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        write!(
            formatter,
            "native VBA cache decode error at byte {}: {:?}",
            self.offset, self.kind
        )
    }
}

impl std::error::Error for DecodeError {}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum Representation {
    ModulePcode,
    Runtime,
}

/// Symbol indices and signed frame locations need different binding tables.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum Location {
    Symbol(u16),
    FrameOffset(i32),
    /// The frame slot holds a reference to the value (for example ByRef Double).
    /// The caller must resolve this reference before supplying the input value.
    IndirectFrameOffset(i32),
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum LiteralNotation {
    Decimal,
    Hexadecimal,
    Keyword,
    Optimized,
}

/// Retains exact integer/Currency storage instead of passing it through f64.
#[derive(Debug, Clone, Copy, PartialEq)]
pub enum Scalar {
    Integer(i16),
    Long(i32),
    Single(f32),
    Double(f64),
    Currency(i64),
    Boolean(bool),
}

impl Scalar {
    pub fn value(self) -> Value {
        match self {
            Self::Integer(value) => Value::Int16(value),
            Self::Long(value) => Value::Integer(i64::from(value)),
            Self::Single(value) => Value::Single(value),
            Self::Double(value) => Value::Double(value),
            Self::Currency(value) => Value::Currency(value),
            Self::Boolean(value) => Value::Boolean(value),
        }
    }
}

#[derive(Debug, Clone, Copy, PartialEq)]
pub struct Constant {
    pub scalar: Scalar,
    pub notation: LiteralNotation,
    /// A typed unary operation, not a sign silently folded into the payload.
    pub negated: bool,
}

#[derive(Debug, Clone, PartialEq)]
pub struct NumericAssignment {
    pub representation: Representation,
    pub input: Location,
    pub constant: Constant,
    pub operation: BinaryOp,
    pub output: Location,
    pub convert_constant_to_double: bool,
    /// Runtime spans can be followed by more instructions; only this prefix
    /// was consumed. A decoded assignment is not a decoded whole function.
    pub consumed: usize,
}

impl NumericAssignment {
    /// Evaluate the span for an already resolved Double input and Double output, using the same
    /// typed unary/arithmetic/coercion rules as source execution. Location
    /// binding, procedure flow, and error handlers remain the caller's job.
    pub fn evaluate_double(&self, input: f64) -> Result<f64, RuntimeError> {
        let mut value = self.constant.scalar.value();
        if self.constant.negated {
            value = runtime::unary(UnaryOp::Neg, value).map_err(|message| {
                let kind = if message.to_ascii_lowercase().contains("overflow") {
                    RuntimeErrorKind::Overflow
                } else {
                    RuntimeErrorKind::TypeMismatch
                };
                runtime::error(kind, message, None)
            })?;
        }
        if self.convert_constant_to_double {
            value = as_double(value)?;
        }
        let value = runtime::binary(self.operation, Value::Double(input), value, false)
            .map_err(|(kind, message)| runtime::error(kind, message, None))?;
        match as_double(value)? {
            Value::Double(value) => Ok(value),
            _ => unreachable!("Double coercion returns a Double"),
        }
    }
}

fn as_double(value: Value) -> Result<Value, RuntimeError> {
    // This path accepts numeric/Boolean values only, so calendar context is
    // irrelevant. The span has no source line until the caller binds it.
    runtime::coerce_declared(value, "Double", 0, 2000).map_err(|mut error| {
        error.line = None;
        error
    })
}

struct Reader<'a> {
    bytes: &'a [u8],
    at: usize,
    origin: usize,
}

impl<'a> Reader<'a> {
    fn new(bytes: &'a [u8], origin: usize) -> Self {
        Self {
            bytes,
            at: 0,
            origin,
        }
    }
    fn error(&self, kind: DecodeErrorKind) -> DecodeError {
        DecodeError {
            offset: self.origin + self.at,
            kind,
        }
    }
    fn take<const N: usize>(&mut self) -> Result<[u8; N], DecodeError> {
        let bytes = self.slice(N)?;
        let mut result = [0; N];
        result.copy_from_slice(bytes);
        Ok(result)
    }
    fn slice(&mut self, count: usize) -> Result<&'a [u8], DecodeError> {
        let bytes = self
            .bytes
            .get(self.at..)
            .and_then(|tail| tail.get(..count))
            .ok_or_else(|| {
                self.error(DecodeErrorKind::Truncated {
                    needed: count,
                    remaining: self.bytes.len().saturating_sub(self.at),
                })
            })?;
        self.at += count;
        Ok(bytes)
    }
    fn word(&mut self) -> Result<u16, DecodeError> {
        Ok(u16::from_le_bytes(self.take()?))
    }
    fn peek_word(&self) -> Option<u16> {
        let bytes = self.bytes.get(self.at..)?.get(..2)?;
        Some(u16::from_le_bytes([bytes[0], bytes[1]]))
    }
    fn expect(&mut self, expected: u16) -> Result<(), DecodeError> {
        let offset = self.origin + self.at;
        let actual = self.word()?;
        if actual == expected {
            Ok(())
        } else {
            Err(DecodeError {
                offset,
                kind: DecodeErrorKind::UnexpectedOpcode { expected, actual },
            })
        }
    }
    fn unsupported(&self, opcode: u16) -> DecodeError {
        DecodeError {
            offset: self.origin + self.at - 2,
            kind: DecodeErrorKind::UnsupportedOpcode(opcode),
        }
    }
}

fn source_constant(bytes: &[u8]) -> Result<Constant, DecodeError> {
    let mut reader = Reader::new(bytes, 4);
    let opcode = reader.word()?;
    let notation = match opcode {
        0xaf | 0xb0 => LiteralNotation::Hexadecimal,
        0xba | 0x4ba => LiteralNotation::Keyword,
        _ => LiteralNotation::Decimal,
    };
    let scalar = match opcode {
        0xac | 0xaf => Scalar::Integer(i16::from_le_bytes(reader.take()?)),
        0xad | 0xb0 => Scalar::Long(i32::from_le_bytes(reader.take()?)),
        0xb6 => Scalar::Single(f32::from_le_bytes(reader.take()?)),
        0xb7 => Scalar::Double(f64::from_le_bytes(reader.take()?)),
        0xa9 => Scalar::Currency(i64::from_le_bytes(reader.take()?)),
        0xba | 0x4ba => Scalar::Boolean(opcode == 0x4ba),
        _ => return Err(reader.unsupported(opcode)),
    };
    let negated = reader.at < bytes.len();
    if negated {
        reader.expect(0x16)?;
    }
    if reader.at != bytes.len() {
        return Err(reader.error(DecodeErrorKind::TrailingInstructions));
    }
    Ok(Constant {
        scalar,
        notation,
        negated,
    })
}

/// Decode an exact source/module numeric assignment span. Raw symbol IDs are
/// retained; this function never guesses variable names from compressed text.
pub fn decode_module_numeric_assignment(
    bytes: &[u8],
    context: CacheContext,
) -> Result<NumericAssignment, DecodeError> {
    context.validate()?;
    if bytes.len() < 12 {
        return Err(DecodeError {
            offset: 0,
            kind: DecodeErrorKind::Truncated {
                needed: 12,
                remaining: bytes.len(),
            },
        });
    }
    let mut reader = Reader::new(bytes, 0);
    reader.expect(0x20)?;
    let input = Location::Symbol(reader.word()?);
    let end = bytes.len() - 6;
    let constant = source_constant(&bytes[4..end])?;
    let mut tail = Reader::new(&bytes[end..], end);
    let opcode = tail.word()?;
    let operation = match opcode {
        0xb => BinaryOp::Add,
        0xc => BinaryOp::Sub,
        0xf => BinaryOp::Mul,
        0x10 => BinaryOp::Div,
        _ => return Err(tail.unsupported(opcode)),
    };
    tail.expect(0x27)?;
    let output = Location::Symbol(tail.word()?);
    Ok(NumericAssignment {
        representation: Representation::ModulePcode,
        input,
        constant,
        operation,
        output,
        convert_constant_to_double: false,
        consumed: bytes.len(),
    })
}

/// Decode a runtime Double arithmetic assignment prefix. The caller supplies
/// its start; no fixed procedure offsets or stack locations are embedded here.
pub fn decode_runtime_numeric_assignment(
    bytes: &[u8],
    context: CacheContext,
) -> Result<NumericAssignment, DecodeError> {
    context.validate()?;
    let mut reader = Reader::new(bytes, 0);
    let load = reader.word()?;
    if !matches!(load, 0x29d | 0x2ed) {
        return Err(reader.unsupported(load));
    }
    let offset = i32::from_le_bytes(reader.take()?);
    let input = if load == 0x2ed {
        Location::IndirectFrameOffset(offset)
    } else {
        Location::FrameOffset(offset)
    };
    let opcode = reader.word()?;
    let scalar = match opcode {
        0x670..=0x67a => Scalar::Integer((opcode - 0x670) as i16),
        0x5ed => Scalar::Integer(i16::from_le_bytes(reader.take()?)),
        0x5f0 => Scalar::Long(i32::from_le_bytes(reader.take()?)),
        0x5f2 => Scalar::Single(f32::from_le_bytes(reader.take()?)),
        0x5f3 => Scalar::Currency(i64::from_le_bytes(reader.take()?)),
        0x5f5 => Scalar::Double(f64::from_le_bytes(reader.take()?)),
        _ => return Err(reader.unsupported(opcode)),
    };
    let negated = reader.peek_word() == Some(0x11b);
    if negated {
        if !matches!(scalar, Scalar::Currency(_)) {
            return Err(reader.error(DecodeErrorKind::UnsupportedOpcode(0x11b)));
        }
        reader.word()?;
    }
    let conversion = match scalar {
        Scalar::Integer(_) => Some(0x1a9),
        Scalar::Long(_) => Some(0x1aa),
        Scalar::Currency(_) => Some(0x1ad),
        _ => None,
    };
    if let Some(opcode) = conversion {
        reader.expect(opcode)?;
    }
    let opcode = reader.word()?;
    let operation = match opcode {
        0xc6 => BinaryOp::Add,
        0xd2 => BinaryOp::Sub,
        0xf6 => BinaryOp::Mul,
        0x102 => BinaryOp::Div,
        _ => return Err(reader.unsupported(opcode)),
    };
    reader.expect(0x2bd)?;
    let output = Location::FrameOffset(i32::from_le_bytes(reader.take()?));
    Ok(NumericAssignment {
        representation: Representation::Runtime,
        input,
        constant: Constant {
            scalar,
            notation: LiteralNotation::Optimized,
            negated,
        },
        operation,
        output,
        convert_constant_to_double: conversion.is_some(),
        consumed: reader.at,
    })
}
