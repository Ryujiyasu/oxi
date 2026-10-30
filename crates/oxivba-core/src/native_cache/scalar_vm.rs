// This Source Code Form is subject to the terms of the Mozilla Public
// License, v. 2.0. If a copy of the MPL was not distributed with this
// file, You can obtain one at https://mozilla.org/MPL/2.0/.

//! Source-independent execution of decoded scalar runtime instructions. This
//! is one execution component, not cache activation or a complete VBA VM.

use super::{
    decode_runtime_frame_layout, CacheContext, DecodeError, DecodeErrorKind, Reader,
    RuntimeCacheLink, RuntimeCompanionCache, RuntimeCompanionType, RuntimeHeapObjectKind,
    RuntimeModuleHeap,
};
use crate::{
    ast::BinaryOp,
    runtime::{self, RuntimeError, Value},
};
use std::{cell::RefCell, collections::BTreeMap, rc::Rc};

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum RuntimeScalarType {
    Void,
    Byte,
    Integer,
    Long,
    LongLong,
    Single,
    Double,
    Currency,
    Date,
    Boolean,
}
impl RuntimeScalarType {
    fn from_type(t: &RuntimeCompanionType) -> Result<Self, RuntimeScalarError> {
        if t.is_array() || t.reference.is_some() || t.word & !0x17f != 0 {
            return Err(RuntimeScalarError::Unsupported("non-scalar type"));
        }
        Ok(match t.base() {
            0 => Self::Void,
            3 => Self::Boolean,
            5 => Self::Byte,
            6 => Self::Integer,
            8 => Self::Long,
            10 => Self::Single,
            11 => Self::Double,
            12 => Self::Date,
            13 => Self::Currency,
            17 => Self::LongLong,
            _ => return Err(RuntimeScalarError::Unsupported("non-scalar type")),
        })
    }
    fn name(self) -> &'static str {
        match self {
            Self::Void => "Empty",
            Self::Byte => "Byte",
            Self::Integer => "Integer",
            Self::Long => "Long",
            Self::LongLong => "LongLong",
            Self::Single => "Single",
            Self::Double => "Double",
            Self::Currency => "Currency",
            Self::Date => "Date",
            Self::Boolean => "Boolean",
        }
    }
    fn physical(self) -> Self {
        match self {
            Self::Date => Self::Double,
            Self::Boolean => Self::Integer,
            other => other,
        }
    }
    fn convert(self, value: Value) -> Result<Value, RuntimeScalarError> {
        if matches!(
            value,
            Value::String(_)
                | Value::Object(_)
                | Value::Nothing
                | Value::Array(_)
                | Value::Record(_)
        ) {
            return Err(RuntimeScalarError::Unsupported(
                "host/string/aggregate conversion",
            ));
        }
        if self == Self::Void {
            return Ok(Value::Empty);
        }
        // This component accepts no calendar-dependent string conversion.
        runtime::coerce_declared(value, self.name(), 0, 2000).map_err(|mut error| {
            error.line = None;
            RuntimeScalarError::Runtime(error)
        })
    }
    fn zero(self) -> Result<Value, RuntimeScalarError> {
        self.convert(Value::Empty)
    }
    fn primary(code: u8) -> Result<Self, RuntimeScalarError> {
        Ok(match code {
            0 => Self::Byte,
            1 => Self::Integer,
            2 => Self::Long,
            3 => Self::Single,
            4 => Self::Double,
            5 => Self::Currency,
            9 => Self::Boolean,
            10 => Self::Date,
            11 => Self::LongLong,
            20 => Self::Void,
            _ => return Err(RuntimeScalarError::Unsupported("managed primary type")),
        })
    }
    fn value_type(value: &Value) -> Option<Self> {
        Some(match value {
            Value::Byte(_) => Self::Byte,
            Value::Int16(_) => Self::Integer,
            Value::Integer(_) => Self::Long,
            Value::LongLong(_) => Self::LongLong,
            Value::Single(_) => Self::Single,
            Value::Double(_) => Self::Double,
            Value::Currency(_) => Self::Currency,
            Value::Date(_) => Self::Date,
            Value::Boolean(_) => Self::Boolean,
            _ => return None,
        })
    }
}

#[derive(Debug)]
pub enum RuntimeScalarError {
    Decode(DecodeError),
    Runtime(RuntimeError),
    Unsupported(&'static str),
    InvalidFrame,
    InvalidCode(usize),
    Arity,
    ByRefTypeMismatch,
    InstructionLimit,
}
impl std::fmt::Display for RuntimeScalarError {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        write!(f, "cached scalar execution: {self:?}")
    }
}
impl std::error::Error for RuntimeScalarError {}
impl From<DecodeError> for RuntimeScalarError {
    fn from(e: DecodeError) -> Self {
        Self::Decode(e)
    }
}

#[derive(Debug, Clone)]
pub struct RuntimeScalarReference {
    declared: RuntimeScalarType,
    value: Rc<RefCell<Value>>,
}
impl RuntimeScalarReference {
    pub fn new(declared: RuntimeScalarType, value: Value) -> Result<Self, RuntimeScalarError> {
        if declared == RuntimeScalarType::Void {
            return Err(RuntimeScalarError::InvalidFrame);
        }
        Ok(Self {
            declared,
            value: Rc::new(RefCell::new(declared.convert(value)?)),
        })
    }
    pub fn value(&self) -> Value {
        self.value.borrow().clone()
    }
    pub fn declared_type(&self) -> RuntimeScalarType {
        self.declared
    }
    fn read_physical(&self, kind: RuntimeScalarType) -> Result<Value, RuntimeScalarError> {
        if self.declared != RuntimeScalarType::Void && self.declared.physical() != kind {
            return Err(RuntimeScalarError::InvalidFrame);
        }
        let value = self.value();
        if self.declared == RuntimeScalarType::Void && matches!(value, Value::Empty) {
            let zero = kind.zero()?;
            *self.value.borrow_mut() = zero.clone();
            return Ok(zero);
        }
        if RuntimeScalarType::value_type(&value).map(|t| t.physical()) != Some(kind) {
            return Err(RuntimeScalarError::InvalidFrame);
        }
        kind.convert(value)
    }
    fn write_physical(
        &self,
        kind: RuntimeScalarType,
        value: Value,
    ) -> Result<(), RuntimeScalarError> {
        if self.declared != RuntimeScalarType::Void && self.declared.physical() != kind {
            return Err(RuntimeScalarError::InvalidFrame);
        }
        let converted = kind.convert(value)?;
        if self.declared == RuntimeScalarType::Void {
            let previous = self.value();
            if !matches!(previous, Value::Empty)
                && RuntimeScalarType::value_type(&previous).map(|t| t.physical()) != Some(kind)
            {
                return Err(RuntimeScalarError::InvalidFrame);
            }
            *self.value.borrow_mut() = converted;
        } else {
            *self.value.borrow_mut() = self.declared.convert(converted)?;
        }
        Ok(())
    }
}
#[derive(Debug, Clone)]
pub enum RuntimeScalarArgument {
    Value(Value),
    Reference(RuntimeScalarReference),
}

#[derive(Debug, Clone)]
enum Instruction {
    Statement(Option<usize>),
    Constant(Value),
    Load(RuntimeScalarType, i32, bool),
    Store(RuntimeScalarType, i32, bool),
    ConvertDouble(RuntimeScalarType),
    TransferDouble,
    Address(i32, bool),
    TemporaryDouble(i32),
    Call {
        result: RuntimeScalarType,
        index: u16,
        argument_bytes: u16,
    },
    Binary(BinaryOp),
    Jump(usize),
    JumpFalse(usize),
    Return(RuntimeScalarType),
}

fn instructions(code: &[u8]) -> Result<BTreeMap<usize, (usize, Instruction)>, RuntimeScalarError> {
    use RuntimeScalarType as T;
    let mut r = Reader::new(code, 0);
    let mut out = BTreeMap::new();
    let mut last_return = false;
    while r.at < code.len() {
        // Native code images include one arbitrary alignment word after some
        // final return instructions. It is not an instruction or a jump target.
        if last_return && code.len() - r.at == 2 {
            break;
        }
        let at = r.at;
        let op = r.word()?;
        let instruction = match op {
            0x267 => {
                let relative = u32::from_le_bytes(r.take()?) as usize;
                let next = if relative == 0 {
                    None
                } else {
                    let next = at
                        .checked_add(relative)
                        .ok_or(RuntimeScalarError::InvalidCode(at))?;
                    if relative < 6 || next >= code.len() || next % 2 != 0 {
                        return Err(RuntimeScalarError::InvalidCode(at));
                    }
                    Some(next)
                };
                Instruction::Statement(next)
            }
            0x670..=0x67a => Instruction::Constant(Value::Int16((op - 0x670) as i16)),
            0x5ed => Instruction::Constant(Value::Int16(i16::from_le_bytes(r.take()?))),
            0x5f0 => {
                Instruction::Constant(Value::Integer(i64::from(i32::from_le_bytes(r.take()?))))
            }
            0x5f2 => Instruction::Constant(Value::Single(f32::from_le_bytes(r.take()?))),
            0x5f3 => Instruction::Constant(Value::Currency(i64::from_le_bytes(r.take()?))),
            0x5f4 | 0x5f5 => Instruction::Constant(Value::Double(f64::from_le_bytes(r.take()?))),
            0x293 => Instruction::Load(T::Single, i32::from_le_bytes(r.take()?), false),
            0x294 => Instruction::Load(T::Double, i32::from_le_bytes(r.take()?), false),
            0x29f => Instruction::Address(i32::from_le_bytes(r.take()?), false),
            0x2ef => Instruction::Address(i32::from_le_bytes(r.take()?), true),
            0x40c => Instruction::TemporaryDouble(i32::from_le_bytes(r.take()?)),
            0x219 | 0x1ac => Instruction::TransferDouble,
            0x510 | 0x511 | 0x512 | 0x515 | 0x51b | 0x51c | 0x51d | 0x51f => {
                let result = match op {
                    0x510 => T::Byte,
                    0x511 => T::Integer,
                    0x512 => T::Long,
                    0x515 => T::Currency,
                    0x51b => T::LongLong,
                    0x51c => T::Single,
                    0x51d => T::Double,
                    _ => T::Void,
                };
                Instruction::Call {
                    result,
                    index: r.word()?,
                    argument_bytes: r.word()?,
                }
            }
            0x290 | 0x291 | 0x292 | 0x295 | 0x29b | 0x29c | 0x29d | 0x2e0 | 0x2e1 | 0x2e2
            | 0x2e5 | 0x2eb | 0x2ec | 0x2ed => {
                let indirect = op >= 0x2e0;
                let base = if indirect { op - 0x50 } else { op };
                let kind = match base {
                    0x290 => T::Byte,
                    0x291 => T::Integer,
                    0x292 => T::Long,
                    0x295 => T::Currency,
                    0x29b => T::LongLong,
                    0x29c => T::Single,
                    _ => T::Double,
                };
                Instruction::Load(kind, i32::from_le_bytes(r.take()?), indirect)
            }
            0x2b0 | 0x2b1 | 0x2b2 | 0x2b5 | 0x2bb | 0x2bc | 0x2bd => {
                let kind = match op {
                    0x2b0 => T::Byte,
                    0x2b1 => T::Integer,
                    0x2b2 => T::Long,
                    0x2b5 => T::Currency,
                    0x2bb => T::LongLong,
                    0x2bc => T::Single,
                    _ => T::Double,
                };
                Instruction::Store(kind, i32::from_le_bytes(r.take()?), false)
            }
            0x30d => Instruction::Store(T::Double, i32::from_le_bytes(r.take()?), true),
            0x1a9 => Instruction::ConvertDouble(T::Integer),
            0x1aa => Instruction::ConvertDouble(T::Long),
            0x1ad => Instruction::ConvertDouble(T::Currency),
            0xc6 => Instruction::Binary(BinaryOp::Add),
            0xd2 => Instruction::Binary(BinaryOp::Sub),
            0xf6 => Instruction::Binary(BinaryOp::Mul),
            0x102 => Instruction::Binary(BinaryOp::Div),
            0x96 => Instruction::Binary(BinaryOp::Gt),
            0x85 => Instruction::Binary(BinaryOp::Lt),
            0x2c6 => Instruction::Jump(u32::from_le_bytes(r.take()?) as usize),
            0x2c7 => Instruction::JumpFalse(u32::from_le_bytes(r.take()?) as usize),
            0x26f => Instruction::Return(T::Byte),
            0x270 => Instruction::Return(T::Integer),
            0x271 => Instruction::Return(T::Long),
            0x272 => Instruction::Return(T::Single),
            0x273 => Instruction::Return(T::Double),
            0x274 => Instruction::Return(T::Currency),
            0x27a => Instruction::Return(T::LongLong),
            0x27b => Instruction::Return(T::Void),
            _ => {
                return Err(DecodeError {
                    offset: at,
                    kind: DecodeErrorKind::UnsupportedOpcode(op),
                }
                .into())
            }
        };
        last_return = matches!(instruction, Instruction::Return(_));
        out.insert(at, (r.at, instruction));
    }
    if !last_return {
        return Err(RuntimeScalarError::InvalidCode(r.at));
    }
    for (at, (_, i)) in &out {
        let target = match i {
            Instruction::Jump(n) | Instruction::JumpFalse(n) => Some(*n),
            Instruction::Statement(n) => *n,
            _ => None,
        };
        if target.is_some_and(|n| !out.contains_key(&n)) {
            return Err(RuntimeScalarError::InvalidCode(*at));
        }
    }
    Ok(out)
}

mod calls;

enum Cell {
    Value(RuntimeScalarReference),
    Reference(RuntimeScalarReference),
}
struct Frame {
    cells: BTreeMap<i32, Cell>,
    local_bytes: u16,
}
impl Frame {
    fn local(&self, offset: i32) -> bool {
        offset <= -8 && offset >= -i32::from(self.local_bytes) && offset % 8 == 0
    }
    fn read(
        &self,
        kind: RuntimeScalarType,
        offset: i32,
        indirect: bool,
    ) -> Result<Value, RuntimeScalarError> {
        match self.cells.get(&offset) {
            Some(Cell::Value(r)) if !indirect => r.read_physical(kind),
            Some(Cell::Reference(r)) if indirect => r.read_physical(kind),
            None if !indirect && self.local(offset) => kind.zero(),
            _ => return Err(RuntimeScalarError::InvalidFrame),
        }
    }
    fn write(
        &mut self,
        kind: RuntimeScalarType,
        offset: i32,
        value: Value,
        indirect: bool,
    ) -> Result<(), RuntimeScalarError> {
        match self.cells.get_mut(&offset) {
            Some(Cell::Value(r)) if !indirect => r.write_physical(kind, value)?,
            Some(Cell::Reference(r)) if indirect => r.write_physical(kind, value)?,
            None if !indirect
                && offset <= -8
                && offset >= -i32::from(self.local_bytes)
                && offset % 8 == 0 =>
            {
                self.cells.insert(
                    offset,
                    Cell::Value(RuntimeScalarReference::new(kind, value)?),
                );
            }
            _ => return Err(RuntimeScalarError::InvalidFrame),
        }
        Ok(())
    }
    fn address(
        &mut self,
        offset: i32,
        forwarded: bool,
    ) -> Result<RuntimeScalarReference, RuntimeScalarError> {
        if !forwarded && !self.cells.contains_key(&offset) && self.local(offset) {
            self.cells.insert(
                offset,
                Cell::Value(RuntimeScalarReference {
                    declared: RuntimeScalarType::Void,
                    value: Rc::new(RefCell::new(Value::Empty)),
                }),
            );
        }
        match self.cells.get(&offset) {
            Some(Cell::Value(r)) if !forwarded => Ok(r.clone()),
            Some(Cell::Reference(r)) if forwarded => Ok(r.clone()),
            _ => Err(RuntimeScalarError::InvalidFrame),
        }
    }
}

/// Bind a standard-module companion entry and execute its independently linked
/// primary code image. No source name/order/body is consulted. Supported scalar
/// instructions include locals, direct/indirect reads, arithmetic, branches,
/// local cached calls and typed returns. Calls use primary auxiliary bindings
/// and physical signatures; companion defaults apply only at this boundary.
/// Unsupported types/instructions/ABIs fail explicitly. Callers
/// retain responsibility for project activation, host API policy and code-page
/// conversions. An instruction budget bounds cycles in untrusted cache code.
pub fn execute_runtime_scalar_method(
    primary: &RuntimeModuleHeap<'_>,
    companion: &RuntimeCompanionCache<'_>,
    entry: usize,
    arguments: &[RuntimeScalarArgument],
    context: CacheContext,
    instruction_budget: usize,
) -> Result<Value, RuntimeScalarError> {
    use RuntimeScalarType as T;
    context.validate()?;
    let arity = companion
        .positional_arity(entry)
        .ok_or(RuntimeScalarError::Unsupported("class/variadic signature"))?;
    if !arity.accepts(arguments.len()) {
        return Err(RuntimeScalarError::Arity);
    }
    if arity.maximum.is_none() {
        return Err(RuntimeScalarError::Unsupported("ParamArray"));
    }
    let method = &companion.methods[entry];
    let result_type = T::from_type(&method.return_type)?;
    let image = primary
        .standard_method_for_link(companion.primary_stream, method.target)
        .ok_or(RuntimeScalarError::Unsupported("primary method target"))?;
    let layout = decode_runtime_frame_layout(image, context)?;
    if !layout.first_cleanup.slots.is_empty() || !layout.second_cleanup.slots.is_empty() {
        return Err(RuntimeScalarError::Unsupported("managed frame cleanup"));
    }
    // The scalar ABI has an eight-byte prefix and one eight-byte slot per
    // argument. Variant results/values have a different ABI, never guessed here.
    if usize::from(layout.argument_bytes) != 8 + method.parameter_types.len() * 8
        || layout.local_bytes % 8 != 0
    {
        return Err(RuntimeScalarError::Unsupported("scalar frame ABI"));
    }
    let defaults = method.optional_defaults(context)?;
    let mut frame = Frame {
        cells: BTreeMap::new(),
        local_bytes: layout.local_bytes,
    };
    for (index, t) in method.parameter_types.iter().enumerate() {
        let kind = T::from_type(t)?;
        if kind == T::Void {
            return Err(RuntimeScalarError::Unsupported("Void parameter"));
        }
        let argument = arguments.get(index);
        let cell = if let Some(RuntimeScalarArgument::Reference(reference)) = argument {
            if t.is_by_ref() {
                if reference.declared != kind {
                    return Err(RuntimeScalarError::ByRefTypeMismatch);
                }
                Cell::Reference(reference.clone())
            } else {
                Cell::Value(RuntimeScalarReference::new(kind, reference.value())?)
            }
        } else {
            let value = match argument {
                Some(RuntimeScalarArgument::Value(v)) if !matches!(v, Value::Missing) => {
                    kind.convert(v.clone())?
                }
                _ if t.is_optional() => {
                    let encoded = defaults
                        .iter()
                        .find(|d| d.parameter_index == index)
                        .and_then(|d| d.encoded.as_ref());
                    let value = if let Some(d) = encoded {
                        d.value.to_value::<RuntimeScalarError>(|_| {
                            Err(RuntimeScalarError::Unsupported("ANSI default"))
                        })?
                    } else {
                        Value::Empty
                    };
                    kind.convert(value)?
                }
                _ => return Err(RuntimeScalarError::Arity),
            };
            if t.is_by_ref() {
                Cell::Reference(RuntimeScalarReference::new(kind, value)?)
            } else {
                Cell::Value(RuntimeScalarReference::new(kind, value)?)
            }
        };
        frame.cells.insert(8 + (index as i32) * 8, cell);
    }
    calls::Machine {
        primary,
        stream: companion.primary_stream,
        context,
    }
    .run(method.target, frame, result_type, instruction_budget)
}
