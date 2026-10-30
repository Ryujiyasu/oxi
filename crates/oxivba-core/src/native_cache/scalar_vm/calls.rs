// This Source Code Form is subject to the terms of the Mozilla Public
// License, v. 2.0. If a copy of the MPL was not distributed with this
// file, You can obtain one at https://mozilla.org/MPL/2.0/.

//! Local cached calls use primary auxiliary links and physical signatures.
//! Companion defaults/names are only for the external invocation boundary.

use super::*;

enum Parameter {
    Scalar(RuntimeScalarType),
    Pointer,
}
struct Method {
    code: BTreeMap<usize, (usize, Instruction)>,
    parameters: Vec<Parameter>,
    result: RuntimeScalarType,
    argument_bytes: u16,
    local_bytes: u16,
}
enum StackValue {
    Scalar(Value),
    Address(RuntimeScalarReference),
}
struct Activation {
    method: u64,
    frame: Frame,
    pc: usize,
    stack: Vec<StackValue>,
}
impl Activation {
    fn pop(&mut self, pc: usize) -> Result<Value, RuntimeScalarError> {
        match self.stack.pop() {
            Some(StackValue::Scalar(v)) => Ok(v),
            _ => Err(RuntimeScalarError::InvalidCode(pc)),
        }
    }
}

pub(super) struct Machine<'a, 'b> {
    pub(super) primary: &'a RuntimeModuleHeap<'b>,
    pub(super) stream: u16,
    pub(super) context: CacheContext,
}
impl Machine<'_, '_> {
    fn target(&self, index: u16) -> Result<RuntimeCacheLink, RuntimeScalarError> {
        self.primary
            .module
            .auxiliary_links
            .get(usize::from(index))
            .copied()
            .ok_or(RuntimeScalarError::Unsupported(
                "cached call auxiliary index",
            ))
    }
    fn compile(&self, target: RuntimeCacheLink) -> Result<Method, RuntimeScalarError> {
        let descriptor = self
            .primary
            .object_for_link(self.stream, target)
            .ok_or(RuntimeScalarError::Unsupported("local cached call target"))?;
        let RuntimeHeapObjectKind::StandardMethod {
            code_slot,
            signature,
        } = descriptor.kind
        else {
            return Err(RuntimeScalarError::Unsupported("non-standard cached call"));
        };
        let signature = self.primary.object_for_link(self.stream, signature).ok_or(
            RuntimeScalarError::Unsupported("primary physical signature"),
        )?;
        let RuntimeHeapObjectKind::Signature {
            return_type,
            parameter_types,
            ..
        } = &signature.kind
        else {
            return Err(RuntimeScalarError::Unsupported(
                "primary physical signature",
            ));
        };
        let result = RuntimeScalarType::primary(*return_type)?;
        let mut parameters = Vec::new();
        for code in *parameter_types {
            parameters.push(if *code == 8 {
                Parameter::Pointer
            } else {
                let kind = RuntimeScalarType::primary(*code)?;
                if kind == RuntimeScalarType::Void {
                    return Err(RuntimeScalarError::Unsupported("Void physical argument"));
                }
                Parameter::Scalar(kind)
            });
        }
        let image = self
            .primary
            .module
            .methods
            .get(usize::from(code_slot))
            .ok_or(RuntimeScalarError::InvalidFrame)?;
        let layout = decode_runtime_frame_layout(image, self.context)?;
        if !layout.first_cleanup.slots.is_empty() || !layout.second_cleanup.slots.is_empty() {
            return Err(RuntimeScalarError::Unsupported("managed frame cleanup"));
        }
        if usize::from(layout.argument_bytes) != 8 + 8 * parameters.len()
            || layout.local_bytes % 8 != 0
        {
            return Err(RuntimeScalarError::Unsupported("primary scalar frame ABI"));
        }
        let code = instructions(image.code)?;
        if code.values().any(
            |(_, i)| matches!(i,Instruction::Return(kind) if kind.physical()!=result.physical()),
        ) {
            return Err(RuntimeScalarError::Unsupported("primary return ABI"));
        }
        Ok(Method {
            code,
            parameters,
            result,
            argument_bytes: layout.argument_bytes,
            local_bytes: layout.local_bytes,
        })
    }
    fn graph(&self, root: RuntimeCacheLink) -> Result<BTreeMap<u64, Method>, RuntimeScalarError> {
        let mut methods = BTreeMap::new();
        let mut pending = vec![root];
        while let Some(target) = pending.pop() {
            if target.context != u32::from(self.stream) << 16 {
                return Err(RuntimeScalarError::Unsupported("cross-module cached call"));
            }
            if methods.contains_key(&target.key) {
                continue;
            }
            let method = self.compile(target)?;
            for (_, instruction) in method.code.values() {
                if let Instruction::Call { index, .. } = instruction {
                    pending.push(self.target(*index)?);
                }
            }
            methods.insert(target.key, method);
        }
        for method in methods.values() {
            for (_, instruction) in method.code.values() {
                if let Instruction::Call {
                    index,
                    result,
                    argument_bytes,
                } = instruction
                {
                    let target = self.target(*index)?;
                    let child = methods
                        .get(&target.key)
                        .ok_or(RuntimeScalarError::InvalidFrame)?;
                    if child.argument_bytes != argument_bytes.saturating_add(8)
                        || child.result.physical() != result.physical()
                    {
                        return Err(RuntimeScalarError::Unsupported("cached call ABI"));
                    }
                }
            }
        }
        Ok(methods)
    }
    pub(super) fn run(
        &self,
        root: RuntimeCacheLink,
        frame: Frame,
        result: RuntimeScalarType,
        budget: usize,
    ) -> Result<Value, RuntimeScalarError> {
        use RuntimeScalarType as T;
        // Compile the reachable local graph before allowing any ByRef effects.
        // Cycles are legal; a visited map bounds preparation without recursion.
        let methods = self.graph(root)?;
        if methods[&root.key].result.physical() != result.physical() {
            return Err(RuntimeScalarError::Unsupported("return ABI"));
        }
        let mut activations = vec![Activation {
            method: root.key,
            frame,
            pc: 0,
            stack: Vec::new(),
        }];
        for _ in 0..budget {
            let active = activations
                .last_mut()
                .ok_or(RuntimeScalarError::InvalidFrame)?;
            let pc = active.pc;
            let (next, instruction) = methods[&active.method]
                .code
                .get(&pc)
                .ok_or(RuntimeScalarError::InvalidCode(pc))?;
            active.pc = *next;
            match instruction {
                Instruction::Statement(_) => {}
                Instruction::Constant(v) => active.stack.push(StackValue::Scalar(v.clone())),
                Instruction::Load(kind, offset, indirect) => {
                    active.stack.push(StackValue::Scalar(
                        active.frame.read(*kind, *offset, *indirect)?,
                    ));
                }
                Instruction::Store(kind, offset, indirect) => {
                    let v = active.pop(pc)?;
                    active.frame.write(*kind, *offset, v, *indirect)?;
                }
                Instruction::Address(offset, forwarded) => {
                    active.stack.push(StackValue::Address(
                        active.frame.address(*offset, *forwarded)?,
                    ));
                }
                Instruction::TemporaryDouble(offset) => {
                    let value = active.pop(pc)?;
                    if !matches!(value, Value::Double(_)) {
                        return Err(RuntimeScalarError::InvalidCode(pc));
                    }
                    active.frame.write(T::Double, *offset, value, false)?;
                    active
                        .stack
                        .push(StackValue::Address(active.frame.address(*offset, false)?));
                }
                Instruction::ConvertDouble(source) => {
                    let v = active.pop(pc)?;
                    if T::value_type(&v) != Some(*source) {
                        return Err(RuntimeScalarError::InvalidCode(pc));
                    }
                    active.stack.push(StackValue::Scalar(T::Double.convert(v)?));
                }
                Instruction::TransferDouble => {
                    if !matches!(
                        active.stack.last(),
                        Some(StackValue::Scalar(Value::Double(_)))
                    ) {
                        return Err(RuntimeScalarError::InvalidCode(pc));
                    }
                }
                Instruction::Binary(op) => {
                    let right = active.pop(pc)?;
                    let left = active.pop(pc)?;
                    if !matches!((&left, &right), (Value::Double(_), Value::Double(_))) {
                        return Err(RuntimeScalarError::InvalidCode(pc));
                    }
                    active.stack.push(StackValue::Scalar(
                        runtime::binary(*op, left, right, false).map_err(|(kind, message)| {
                            RuntimeScalarError::Runtime(runtime::error(kind, message, None))
                        })?,
                    ));
                }
                Instruction::Jump(to) => active.pc = *to,
                Instruction::JumpFalse(to) => {
                    let v = active.pop(pc)?;
                    if !matches!(v, Value::Boolean(_) | Value::Int16(_)) {
                        return Err(RuntimeScalarError::InvalidCode(pc));
                    }
                    if matches!(T::Boolean.convert(v)?, Value::Boolean(false)) {
                        active.pc = *to;
                    }
                }
                Instruction::Call { index, .. } => {
                    let target = self.target(*index)?;
                    let callee = &methods[&target.key];
                    let mut frame = Frame {
                        cells: BTreeMap::new(),
                        local_bytes: callee.local_bytes,
                    };
                    // Native callsites stage arguments right to left. Parameter
                    // zero consumes the newest staged value/address.
                    for (i, parameter) in callee.parameters.iter().enumerate() {
                        let cell = match (parameter, active.stack.pop()) {
                            (Parameter::Pointer, Some(StackValue::Address(r))) => {
                                Cell::Reference(r)
                            }
                            (Parameter::Scalar(kind), Some(StackValue::Scalar(v)))
                                if T::value_type(&v).map(|t| t.physical())
                                    == Some(kind.physical()) =>
                            {
                                Cell::Value(RuntimeScalarReference::new(*kind, v)?)
                            }
                            _ => return Err(RuntimeScalarError::InvalidCode(pc)),
                        };
                        frame.cells.insert(8 + 8 * (i as i32), cell);
                    }
                    activations.push(Activation {
                        method: target.key,
                        frame,
                        pc: 0,
                        stack: Vec::new(),
                    });
                }
                Instruction::Return(kind) => {
                    if !active.stack.is_empty() {
                        return Err(RuntimeScalarError::InvalidCode(pc));
                    }
                    let value = if *kind == T::Void {
                        Value::Empty
                    } else {
                        active.frame.read(*kind, -8, false)?
                    };
                    activations.pop();
                    if let Some(parent) = activations.last_mut() {
                        if *kind != T::Void {
                            parent.stack.push(StackValue::Scalar(value));
                        }
                    } else {
                        return result.convert(value);
                    }
                }
            }
        }
        Err(RuntimeScalarError::InstructionLimit)
    }
}
