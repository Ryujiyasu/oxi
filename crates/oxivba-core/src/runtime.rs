// This Source Code Form is subject to the terms of the Mozilla Public
// License, v. 2.0. If a copy of the MPL was not distributed with this
// file, You can obtain one at https://mozilla.org/MPL/2.0/.

//! Host-independent execution of pure VBA procedures.
//!
//! This first browser-runtime slice supports scalar values, local variables,
//! arithmetic, comparisons, branches, loops, and calls between VBA procedures.
//! Office objects require a host adapter. Multidimensional arrays, file I/O,
//! and events fail explicitly rather than being approximated.

use std::{cell::RefCell, collections::BTreeMap, rc::Rc};

use crate::ast::{
    Argument, ArrayBound, BinaryOp, CaseLabel, DoStmt, ExitKind, Expr, ForEachStmt, ForStmt,
    Literal, LoopTest, Module, ModuleItem, ModuleOption, ParamMode, ProcKind, Procedure,
    SelectCaseStmt, Statement, TypeName, UnaryOp, VarDecl, VarItem,
};

#[derive(Debug, Clone, PartialEq)]
pub enum Value {
    Empty,
    /// An omitted `Optional ... As Variant` argument. Unlike `Empty`, this is
    /// observable through VBA's `IsMissing` function.
    Missing,
    Null,
    Boolean(bool),
    Integer(i64),
    Double(f64),
    String(String),
    Array(ArrayValue),
    Object(ObjectRef),
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct ObjectRef {
    pub handle: u64,
    pub kind: String,
}

pub trait Host {
    fn call(
        &mut self,
        receiver: Option<&ObjectRef>,
        name: &str,
        args: &[Value],
    ) -> Result<Option<Value>, String>;

    fn get(&mut self, receiver: &ObjectRef, name: &str) -> Result<Option<Value>, String>;

    fn set(&mut self, receiver: &ObjectRef, name: &str, value: Value) -> Result<bool, String>;

    fn enumerate(&mut self, _receiver: &ObjectRef) -> Result<Option<Vec<Value>>, String> {
        Ok(None)
    }
}

#[derive(Debug, Clone, PartialEq)]
pub struct ArrayValue {
    pub lower_bound: i64,
    pub values: Vec<Value>,
    pub element_default: Box<Value>,
}

impl ArrayValue {
    pub fn upper_bound(&self) -> i64 {
        match self.values.len().checked_sub(1) {
            Some(offset) => self.lower_bound.saturating_add(offset as i64),
            None => self.lower_bound.saturating_sub(1),
        }
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum RuntimeErrorKind {
    ProcedureNotFound,
    ArgumentCount,
    UndefinedVariable,
    TypeMismatch,
    Overflow,
    SubscriptOutOfRange,
    Host,
    DivisionByZero,
    Unsupported,
    StepLimit,
    CallDepth,
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct RuntimeError {
    pub kind: RuntimeErrorKind,
    pub message: String,
    pub line: Option<u32>,
}

impl std::fmt::Display for RuntimeError {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        match self.line {
            Some(line) => write!(f, "line {line}: {}", self.message),
            None => f.write_str(&self.message),
        }
    }
}

impl std::error::Error for RuntimeError {}

pub fn execute(module: &Module, procedure: &str, args: Vec<Value>) -> Result<Value, RuntimeError> {
    Runtime::new(module).call(procedure, args)
}

pub fn execute_with_host(
    module: &Module,
    procedure: &str,
    args: Vec<Value>,
    host: &mut dyn Host,
) -> Result<Value, RuntimeError> {
    Runtime::new(module).with_host(host).call(procedure, args)
}

pub struct Runtime<'a> {
    module: &'a Module,
    host: Option<&'a mut dyn Host>,
    steps: usize,
    max_steps: usize,
    depth: usize,
    max_depth: usize,
}

struct Frame {
    procedure_name: String,
    values: BTreeMap<String, ValueSlot>,
    with_objects: Vec<ObjectRef>,
}

type ValueSlot = Rc<RefCell<Value>>;

enum BoundArgument {
    Value(Value),
    Reference(ValueSlot),
}

enum Flow {
    Continue,
    Exit(ExitKind),
    End,
}

impl<'a> Runtime<'a> {
    pub fn new(module: &'a Module) -> Self {
        Self {
            module,
            host: None,
            steps: 0,
            max_steps: 100_000,
            depth: 0,
            max_depth: 128,
        }
    }

    pub fn with_host(mut self, host: &'a mut dyn Host) -> Self {
        self.host = Some(host);
        self
    }

    pub fn with_limits(mut self, max_steps: usize, max_depth: usize) -> Self {
        self.max_steps = max_steps;
        self.max_depth = max_depth;
        self
    }

    pub fn call(&mut self, name: &str, args: Vec<Value>) -> Result<Value, RuntimeError> {
        self.steps = 0;
        self.depth = 0;
        self.call_procedure(name, args, None)
    }

    fn call_procedure(
        &mut self,
        name: &str,
        args: Vec<Value>,
        line: Option<u32>,
    ) -> Result<Value, RuntimeError> {
        let procedure = self.find_procedure(name, line)?;
        let fixed_count = procedure
            .params
            .iter()
            .position(|param| param.mode == ParamMode::ParamArray)
            .unwrap_or(procedure.params.len());
        if args.len() > fixed_count
            && !procedure
                .params
                .last()
                .is_some_and(|param| param.mode == ParamMode::ParamArray)
        {
            return Err(argument_count_error(&procedure, args.len(), line));
        }

        let received = args.len();
        let mut values = args.into_iter();
        let mut bound = Vec::with_capacity(procedure.params.len());
        for param in &procedure.params[..fixed_count] {
            match values.next() {
                Some(value) => bound.push(BoundArgument::Value(value)),
                None => bound.push(BoundArgument::Value(
                    self.omitted_parameter_value(param, line)?,
                )),
            }
        }
        if let Some(param_array) = procedure.params.get(fixed_count) {
            bound.push(BoundArgument::Value(Value::Array(ArrayValue {
                lower_bound: 0,
                values: values.collect(),
                element_default: Box::new(default_value(&param_array.type_name)),
            })));
        } else if received < fixed_count
            && procedure.params[received..]
                .iter()
                .any(|param| !param.optional)
        {
            return Err(argument_count_error(&procedure, received, line));
        }
        self.invoke_procedure(&procedure, bound, line)
    }

    fn invoke_procedure(
        &mut self,
        procedure: &Procedure,
        args: Vec<BoundArgument>,
        line: Option<u32>,
    ) -> Result<Value, RuntimeError> {
        if self.depth >= self.max_depth {
            return Err(error(
                RuntimeErrorKind::CallDepth,
                "VBA call-depth limit exceeded",
                line,
            ));
        }
        if args.len() != procedure.params.len() {
            return Err(argument_count_error(procedure, args.len(), line));
        }

        let mut frame = Frame {
            procedure_name: key(&procedure.name),
            values: BTreeMap::new(),
            with_objects: Vec::new(),
        };
        for (param, argument) in procedure.params.iter().zip(args) {
            let value = match argument {
                BoundArgument::Value(value) => Rc::new(RefCell::new(value)),
                BoundArgument::Reference(value) => value,
            };
            frame.values.insert(key(&param.name), value);
        }
        if !matches!(procedure.kind, ProcKind::Sub) {
            frame.values.insert(
                frame.procedure_name.clone(),
                Rc::new(RefCell::new(default_return_value(procedure))),
            );
        }

        self.depth += 1;
        let flow = self.exec_body(&procedure.body, &mut frame);
        self.depth -= 1;
        let ended = match flow? {
            Flow::Continue
            | Flow::Exit(ExitKind::Sub | ExitKind::Function | ExitKind::Property) => false,
            Flow::End => true,
            Flow::Exit(kind) => {
                return Err(error(
                    RuntimeErrorKind::Unsupported,
                    format!("unmatched Exit {kind:?}"),
                    Some(procedure.span.line),
                ))
            }
        };
        let value = if ended || matches!(procedure.kind, ProcKind::Sub) {
            Value::Empty
        } else {
            let value = frame
                .values
                .remove(&frame.procedure_name)
                .map(|value| value.borrow().clone())
                .unwrap_or(Value::Empty);
            value
        };
        Ok(value)
    }

    fn find_procedure(&self, name: &str, line: Option<u32>) -> Result<Procedure, RuntimeError> {
        self.module
            .items
            .iter()
            .find_map(|item| match item {
                ModuleItem::Procedure(procedure) if procedure.name.eq_ignore_ascii_case(name) => {
                    Some(procedure.clone())
                }
                _ => None,
            })
            .ok_or_else(|| {
                error(
                    RuntimeErrorKind::ProcedureNotFound,
                    format!("VBA procedure not found: {name}"),
                    line,
                )
            })
    }

    fn omitted_parameter_value(
        &mut self,
        parameter: &crate::ast::Param,
        line: Option<u32>,
    ) -> Result<Value, RuntimeError> {
        if !parameter.optional {
            return Err(error(
                RuntimeErrorKind::ArgumentCount,
                format!("required argument is missing: {}", parameter.name),
                line,
            ));
        }
        if let Some(default) = &parameter.default {
            let mut frame = Frame {
                procedure_name: String::new(),
                values: BTreeMap::new(),
                with_objects: Vec::new(),
            };
            return self.eval_expr(default, &mut frame);
        }
        if parameter.type_name.name.eq_ignore_ascii_case("variant") {
            Ok(Value::Missing)
        } else {
            Ok(default_value(&parameter.type_name))
        }
    }

    fn exec_body(&mut self, body: &[Statement], frame: &mut Frame) -> Result<Flow, RuntimeError> {
        for statement in body {
            self.tick(line_of(statement))?;
            let flow = self.exec_statement(statement, frame)?;
            if !matches!(flow, Flow::Continue) {
                return Ok(flow);
            }
        }
        Ok(Flow::Continue)
    }

    fn exec_statement(
        &mut self,
        statement: &Statement,
        frame: &mut Frame,
    ) -> Result<Flow, RuntimeError> {
        match statement {
            Statement::Assign {
                target,
                value,
                span,
            } => {
                let value = self.eval_expr(value, frame)?;
                self.assign(target, value, frame, span.line)?;
                Ok(Flow::Continue)
            }
            Statement::SetAssign {
                target,
                value,
                span,
            } => {
                let value = self.eval_expr(value, frame)?;
                if !matches!(value, Value::Object(_)) {
                    return Err(error(
                        RuntimeErrorKind::TypeMismatch,
                        "Set requires an object value",
                        Some(span.line),
                    ));
                }
                self.assign(target, value, frame, span.line)?;
                Ok(Flow::Continue)
            }
            Statement::Dim(decl) => {
                self.declare_locals(decl, frame)?;
                Ok(Flow::Continue)
            }
            Statement::ReDim {
                preserve,
                items,
                span,
            } => {
                for item in items {
                    self.redim(item, *preserve, frame, span.line)?;
                }
                Ok(Flow::Continue)
            }
            Statement::If(branch) => {
                if truthy(&self.eval_expr(&branch.condition, frame)?).map_err(|message| {
                    error(
                        RuntimeErrorKind::TypeMismatch,
                        message,
                        Some(branch.span.line),
                    )
                })? {
                    self.exec_body(&branch.then_body, frame)
                } else {
                    for (condition, body) in &branch.else_ifs {
                        if truthy(&self.eval_expr(condition, frame)?).map_err(|message| {
                            error(
                                RuntimeErrorKind::TypeMismatch,
                                message,
                                Some(branch.span.line),
                            )
                        })? {
                            return self.exec_body(body, frame);
                        }
                    }
                    match &branch.else_body {
                        Some(body) => self.exec_body(body, frame),
                        None => Ok(Flow::Continue),
                    }
                }
            }
            Statement::For(loop_) => self.exec_for(loop_, frame),
            Statement::ForEach(loop_) => self.exec_for_each(loop_, frame),
            Statement::Do(loop_) => self.exec_do(loop_, frame),
            Statement::SelectCase(select) => self.exec_select_case(select, frame),
            Statement::While {
                condition,
                body,
                span,
            } => self.exec_while(condition, body, frame, span.line),
            Statement::With { subject, body, .. } => {
                let receiver = self.eval_object(subject, frame, subject.span().line)?;
                frame.with_objects.push(receiver);
                let result = self.exec_body(body, frame);
                frame.with_objects.pop();
                result
            }
            Statement::Call { target, .. } => {
                self.eval_call(target, frame)?;
                Ok(Flow::Continue)
            }
            Statement::Exit { what, .. } => Ok(Flow::Exit(*what)),
            Statement::End { .. } => Ok(Flow::End),
            Statement::Comment { .. } | Statement::Label { .. } | Statement::LineNumber { .. } => {
                Ok(Flow::Continue)
            }
            Statement::Unknown { text, span } => Err(error(
                RuntimeErrorKind::Unsupported,
                format!("cannot execute unparsed VBA: {text}"),
                Some(span.line),
            )),
            other => Err(error(
                RuntimeErrorKind::Unsupported,
                format!(
                    "VBA statement is not executable yet: {}",
                    statement_name(other)
                ),
                line_of(other),
            )),
        }
    }

    fn exec_select_case(
        &mut self,
        select: &SelectCaseStmt,
        frame: &mut Frame,
    ) -> Result<Flow, RuntimeError> {
        let subject = self.eval_expr(&select.subject, frame)?;
        for case in &select.cases {
            for label in &case.labels {
                if self.case_matches(&subject, label, frame, select.span.line)? {
                    return self.exec_body(&case.body, frame);
                }
            }
        }
        match &select.case_else {
            Some(body) => self.exec_body(body, frame),
            None => Ok(Flow::Continue),
        }
    }

    fn case_matches(
        &mut self,
        subject: &Value,
        label: &CaseLabel,
        frame: &mut Frame,
        line: u32,
    ) -> Result<bool, RuntimeError> {
        let compare = |op, lhs, rhs| {
            binary(op, lhs, rhs)
                .map_err(|(kind, message)| error(kind, message, Some(line)))
                .and_then(|value| {
                    truthy(&value).map_err(|message| {
                        error(RuntimeErrorKind::TypeMismatch, message, Some(line))
                    })
                })
        };
        match label {
            CaseLabel::Value(value) => {
                let value = self.eval_expr(value, frame)?;
                compare(BinaryOp::Eq, subject.clone(), value)
            }
            CaseLabel::Range(lower, upper) => {
                let lower = self.eval_expr(lower, frame)?;
                let upper = self.eval_expr(upper, frame)?;
                Ok(compare(BinaryOp::Ge, subject.clone(), lower)?
                    && compare(BinaryOp::Le, subject.clone(), upper)?)
            }
            CaseLabel::Compare(op, value) => {
                let value = self.eval_expr(value, frame)?;
                compare(*op, subject.clone(), value)
            }
        }
    }

    fn exec_for(&mut self, loop_: &ForStmt, frame: &mut Frame) -> Result<Flow, RuntimeError> {
        let from = self.eval_expr(&loop_.from, frame)?;
        let to = self.eval_expr(&loop_.to, frame)?;
        let step = match &loop_.step {
            Some(step) => self.eval_expr(step, frame)?,
            None => Value::Integer(1),
        };
        let limit = number(&to).map_err(|message| {
            error(
                RuntimeErrorKind::TypeMismatch,
                message,
                Some(loop_.span.line),
            )
        })?;
        let increment = number(&step).map_err(|message| {
            error(
                RuntimeErrorKind::TypeMismatch,
                message,
                Some(loop_.span.line),
            )
        })?;
        self.assign(&loop_.counter, from, frame, loop_.span.line)?;

        loop {
            self.tick(Some(loop_.span.line))?;
            let current = self.eval_expr(&loop_.counter, frame)?;
            let current_number = number(&current).map_err(|message| {
                error(
                    RuntimeErrorKind::TypeMismatch,
                    message,
                    Some(loop_.span.line),
                )
            })?;
            if (increment >= 0.0 && current_number > limit)
                || (increment < 0.0 && current_number < limit)
            {
                return Ok(Flow::Continue);
            }
            match self.exec_body(&loop_.body, frame)? {
                Flow::Continue => {}
                Flow::Exit(ExitKind::For) => return Ok(Flow::Continue),
                flow => return Ok(flow),
            }
            let next = numeric_result(current_number + increment, &current, &step);
            self.assign(&loop_.counter, next, frame, loop_.span.line)?;
        }
    }

    fn exec_for_each(
        &mut self,
        loop_: &ForEachStmt,
        frame: &mut Frame,
    ) -> Result<Flow, RuntimeError> {
        let collection = self.eval_expr(&loop_.collection, frame)?;
        let values = match collection {
            Value::Array(array) => array.values,
            Value::Object(object) => {
                self.host_enumerate(&object, loop_.span.line)?
                    .ok_or_else(|| {
                        error(
                            RuntimeErrorKind::TypeMismatch,
                            format!("VBA object is not enumerable: {}", object.kind),
                            Some(loop_.span.line),
                        )
                    })?
            }
            _ => {
                return Err(error(
                    RuntimeErrorKind::TypeMismatch,
                    "For Each requires a VBA array or enumerable object",
                    Some(loop_.span.line),
                ))
            }
        };
        for value in values {
            self.tick(Some(loop_.span.line))?;
            self.assign(&loop_.item, value, frame, loop_.span.line)?;
            match self.exec_body(&loop_.body, frame)? {
                Flow::Continue => {}
                Flow::Exit(ExitKind::For) => return Ok(Flow::Continue),
                flow => return Ok(flow),
            }
        }
        Ok(Flow::Continue)
    }

    fn exec_do(&mut self, loop_: &DoStmt, frame: &mut Frame) -> Result<Flow, RuntimeError> {
        loop {
            self.tick(Some(loop_.span.line))?;
            if let Some(test) = &loop_.pre {
                if !self.loop_continues(test, frame, loop_.span.line)? {
                    return Ok(Flow::Continue);
                }
            }
            match self.exec_body(&loop_.body, frame)? {
                Flow::Continue => {}
                Flow::Exit(ExitKind::Do) => return Ok(Flow::Continue),
                flow => return Ok(flow),
            }
            if let Some(test) = &loop_.post {
                if !self.loop_continues(test, frame, loop_.span.line)? {
                    return Ok(Flow::Continue);
                }
            }
        }
    }

    fn exec_while(
        &mut self,
        condition: &Expr,
        body: &[Statement],
        frame: &mut Frame,
        line: u32,
    ) -> Result<Flow, RuntimeError> {
        loop {
            self.tick(Some(line))?;
            if !self.condition(condition, frame, line)? {
                return Ok(Flow::Continue);
            }
            match self.exec_body(body, frame)? {
                Flow::Continue => {}
                flow => return Ok(flow),
            }
        }
    }

    fn loop_continues(
        &mut self,
        test: &LoopTest,
        frame: &mut Frame,
        line: u32,
    ) -> Result<bool, RuntimeError> {
        let matched = self.condition(&test.condition, frame, line)?;
        Ok(if test.until { !matched } else { matched })
    }

    fn condition(
        &mut self,
        condition: &Expr,
        frame: &mut Frame,
        line: u32,
    ) -> Result<bool, RuntimeError> {
        truthy(&self.eval_expr(condition, frame)?)
            .map_err(|message| error(RuntimeErrorKind::TypeMismatch, message, Some(line)))
    }

    fn declare_locals(&mut self, decl: &VarDecl, frame: &mut Frame) -> Result<(), RuntimeError> {
        for variable in &decl.items {
            let value = match &variable.array_bounds {
                Some(bounds) => {
                    self.make_array(bounds, &variable.type_name, frame, decl.span.line)?
                }
                None => match &variable.value {
                    Some(expr) => self.eval_expr(expr, frame)?,
                    None => default_value(&variable.type_name),
                },
            };
            frame
                .values
                .insert(key(&variable.name), Rc::new(RefCell::new(value)));
        }
        Ok(())
    }

    fn redim(
        &mut self,
        item: &VarItem,
        preserve: bool,
        frame: &mut Frame,
        line: u32,
    ) -> Result<(), RuntimeError> {
        let bounds = item.array_bounds.as_ref().ok_or_else(|| {
            error(
                RuntimeErrorKind::TypeMismatch,
                "ReDim target must have array bounds",
                Some(line),
            )
        })?;
        let mut replacement = match self.make_array(bounds, &item.type_name, frame, line)? {
            Value::Array(array) => array,
            _ => unreachable!(),
        };
        if item.type_name.name.eq_ignore_ascii_case("variant") {
            if let Some(existing) = frame.values.get(&key(&item.name)) {
                if let Value::Array(existing) = &*existing.borrow() {
                    replacement.element_default = existing.element_default.clone();
                    replacement.values.fill(*existing.element_default.clone());
                }
            }
        }
        if preserve {
            let existing = frame.values.get(&key(&item.name)).cloned().ok_or_else(|| {
                error(
                    RuntimeErrorKind::UndefinedVariable,
                    format!("undefined VBA array: {}", item.name),
                    Some(line),
                )
            })?;
            let existing = existing.borrow();
            let Value::Array(existing) = &*existing else {
                return Err(error(
                    RuntimeErrorKind::TypeMismatch,
                    format!("ReDim Preserve target is not an array: {}", item.name),
                    Some(line),
                ));
            };
            if !existing.values.is_empty() && existing.lower_bound != replacement.lower_bound {
                return Err(error(
                    RuntimeErrorKind::SubscriptOutOfRange,
                    "ReDim Preserve cannot change an array's lower bound",
                    Some(line),
                ));
            }
            let first = existing.lower_bound.max(replacement.lower_bound);
            let last = existing.upper_bound().min(replacement.upper_bound());
            for index in first..=last {
                replacement.values[(index - replacement.lower_bound) as usize] =
                    existing.values[(index - existing.lower_bound) as usize].clone();
            }
        }
        let replacement = Value::Array(replacement);
        match frame.values.get(&key(&item.name)) {
            Some(value) => *value.borrow_mut() = replacement,
            None => {
                frame
                    .values
                    .insert(key(&item.name), Rc::new(RefCell::new(replacement)));
            }
        }
        Ok(())
    }

    fn make_array(
        &mut self,
        bounds: &[ArrayBound],
        element_type: &TypeName,
        frame: &mut Frame,
        line: u32,
    ) -> Result<Value, RuntimeError> {
        if bounds.is_empty() {
            return Ok(Value::Array(ArrayValue {
                lower_bound: self.option_base(),
                values: Vec::new(),
                element_default: Box::new(default_value(element_type)),
            }));
        }
        if bounds.len() != 1 {
            return Err(error(
                RuntimeErrorKind::Unsupported,
                "only one-dimensional VBA arrays are executable yet",
                Some(line),
            ));
        }
        let bound = &bounds[0];
        let lower = match &bound.lower {
            Some(lower) => self.array_index(lower, frame, line)?,
            None => self.option_base(),
        };
        let upper = self.array_index(&bound.upper, frame, line)?;
        if upper < lower {
            return Err(error(
                RuntimeErrorKind::SubscriptOutOfRange,
                format!("invalid VBA array bounds: {lower} To {upper}"),
                Some(line),
            ));
        }
        let len = upper
            .checked_sub(lower)
            .and_then(|value| value.checked_add(1))
            .and_then(|value| usize::try_from(value).ok())
            .filter(|len| *len <= 1_000_000)
            .ok_or_else(|| {
                error(
                    RuntimeErrorKind::Overflow,
                    "VBA array is too large for the browser runtime",
                    Some(line),
                )
            })?;
        let default = default_value(element_type);
        Ok(Value::Array(ArrayValue {
            lower_bound: lower,
            values: vec![default.clone(); len],
            element_default: Box::new(default),
        }))
    }

    fn option_base(&self) -> i64 {
        self.module
            .items
            .iter()
            .find_map(|item| match item {
                ModuleItem::Option(ModuleOption::Base(value), _) => Some(*value as i64),
                _ => None,
            })
            .unwrap_or(0)
    }

    fn assign(
        &mut self,
        target: &Expr,
        value: Value,
        frame: &mut Frame,
        line: u32,
    ) -> Result<(), RuntimeError> {
        match target {
            Expr::Ident(name, _) | Expr::TypedIdent { name, .. } => {
                match frame.values.get(&key(name)) {
                    Some(existing) => *existing.borrow_mut() = value,
                    None => {
                        frame.values.insert(key(name), Rc::new(RefCell::new(value)));
                    }
                }
                Ok(())
            }
            Expr::Index { target, args, .. } => {
                let name = expr_name(target).ok_or_else(|| {
                    error(
                        RuntimeErrorKind::Unsupported,
                        "array assignment target must be a local variable",
                        Some(line),
                    )
                })?;
                let index = self.single_array_argument(args, frame, line)?;
                let array = frame.values.get(&key(name)).cloned().ok_or_else(|| {
                    error(
                        RuntimeErrorKind::UndefinedVariable,
                        format!("undefined VBA array: {name}"),
                        Some(line),
                    )
                })?;
                let mut array = array.borrow_mut();
                let Value::Array(array) = &mut *array else {
                    return Err(error(
                        RuntimeErrorKind::TypeMismatch,
                        format!("VBA value is not an array: {name}"),
                        Some(line),
                    ));
                };
                let offset = array_offset(array, index, line)?;
                array.values[offset] = value;
                Ok(())
            }
            Expr::Member {
                object, name, span, ..
            } => {
                let receiver = self.eval_object(object, frame, span.line)?;
                match self.host_set(&receiver, name, value, span.line)? {
                    true => Ok(()),
                    false => Err(error(
                        RuntimeErrorKind::Unsupported,
                        format!("host property is not writable: {}.{name}", receiver.kind),
                        Some(span.line),
                    )),
                }
            }
            Expr::WithMember(name, span) | Expr::WithBangMember(name, span) => {
                let receiver = current_with_object(frame, span.line)?;
                match self.host_set(&receiver, name, value, span.line)? {
                    true => Ok(()),
                    false => Err(error(
                        RuntimeErrorKind::Unsupported,
                        format!("host property is not writable: {}.{name}", receiver.kind),
                        Some(span.line),
                    )),
                }
            }
            _ => Err(error(
                RuntimeErrorKind::Unsupported,
                "assignment target is not executable yet",
                Some(line),
            )),
        }
    }

    fn eval_expr(&mut self, expr: &Expr, frame: &mut Frame) -> Result<Value, RuntimeError> {
        match expr {
            Expr::Literal(literal, _) => Ok(literal_value(literal)),
            Expr::Ident(name, span) | Expr::TypedIdent { name, span, .. } => {
                if let Some(value) = frame.values.get(&key(name)) {
                    return Ok(value.borrow().clone());
                }
                if let Some(value) = self.host_call(None, name, &[], span.line)? {
                    return Ok(value);
                }
                Err(error(
                    RuntimeErrorKind::UndefinedVariable,
                    format!("undefined VBA variable: {name}"),
                    Some(span.line),
                ))
            }
            Expr::Unary { op, operand, span } => {
                let value = self.eval_expr(operand, frame)?;
                unary(*op, value).map_err(|message| {
                    error(RuntimeErrorKind::TypeMismatch, message, Some(span.line))
                })
            }
            Expr::Binary { op, lhs, rhs, span } => {
                let lhs = self.eval_expr(lhs, frame)?;
                let rhs = self.eval_expr(rhs, frame)?;
                binary(*op, lhs, rhs)
                    .map_err(|(kind, message)| error(kind, message, Some(span.line)))
            }
            Expr::Index {
                target, args, span, ..
            } if expr_name(target).is_some_and(|name| {
                frame
                    .values
                    .get(&key(name))
                    .is_some_and(|value| matches!(&*value.borrow(), Value::Array(_)))
            }) =>
            {
                let name = expr_name(target).unwrap();
                let index = self.single_array_argument(args, frame, span.line)?;
                let value = frame.values.get(&key(name)).unwrap().borrow();
                let Value::Array(array) = &*value else {
                    unreachable!()
                };
                Ok(array.values[array_offset(array, index, span.line)?].clone())
            }
            Expr::Index { .. } => self.eval_call(expr, frame),
            Expr::Member {
                object, name, span, ..
            } => {
                let receiver = self.eval_object(object, frame, span.line)?;
                self.host_get(&receiver, name, span.line)?.ok_or_else(|| {
                    error(
                        RuntimeErrorKind::Unsupported,
                        format!("host property is not available: {}.{name}", receiver.kind),
                        Some(span.line),
                    )
                })
            }
            Expr::WithMember(name, span) | Expr::WithBangMember(name, span) => {
                let receiver = current_with_object(frame, span.line)?;
                self.host_get(&receiver, name, span.line)?.ok_or_else(|| {
                    error(
                        RuntimeErrorKind::Unsupported,
                        format!("host property is not available: {}.{name}", receiver.kind),
                        Some(span.line),
                    )
                })
            }
            _ => Err(error(
                RuntimeErrorKind::Unsupported,
                "VBA expression is not executable yet",
                Some(expr.span().line),
            )),
        }
    }

    fn eval_call(&mut self, expr: &Expr, frame: &mut Frame) -> Result<Value, RuntimeError> {
        match expr {
            Expr::Index {
                target,
                args,
                force_by_value,
                span,
            } => {
                if let Expr::Ident(name, _) | Expr::TypedIdent { name, .. } = target.as_ref() {
                    if self.module.items.iter().any(|item| {
                        matches!(item, ModuleItem::Procedure(p) if p.name.eq_ignore_ascii_case(name))
                    }) {
                        return self.call_user_procedure(
                            name,
                            args,
                            *force_by_value,
                            frame,
                            Some(span.line),
                        );
                    }
                }
                let mut values = Vec::with_capacity(args.len());
                for argument in args {
                    let value = argument.value.as_ref().ok_or_else(|| {
                        error(
                            RuntimeErrorKind::Unsupported,
                            "omitted arguments are not executable yet",
                            Some(span.line),
                        )
                    })?;
                    values.push(self.eval_expr(value, frame)?);
                }
                match target.as_ref() {
                    Expr::Ident(name, _) | Expr::TypedIdent { name, .. } => {
                        self.call_named(name, values, Some(span.line))
                    }
                    Expr::Member { object, name, .. } => {
                        let receiver = self.eval_object(object, frame, span.line)?;
                        self.host_call(Some(&receiver), name, &values, span.line)?
                            .ok_or_else(|| {
                                error(
                                    RuntimeErrorKind::Unsupported,
                                    format!(
                                        "host method is not available: {}.{name}",
                                        receiver.kind
                                    ),
                                    Some(span.line),
                                )
                            })
                    }
                    Expr::WithMember(name, _) | Expr::WithBangMember(name, _) => {
                        let receiver = current_with_object(frame, span.line)?;
                        self.host_call(Some(&receiver), name, &values, span.line)?
                            .ok_or_else(|| {
                                error(
                                    RuntimeErrorKind::Unsupported,
                                    format!(
                                        "host method is not available: {}.{name}",
                                        receiver.kind
                                    ),
                                    Some(span.line),
                                )
                            })
                    }
                    _ => Err(error(
                        RuntimeErrorKind::Unsupported,
                        "call target is not executable yet",
                        Some(span.line),
                    )),
                }
            }
            Expr::Ident(name, span) | Expr::TypedIdent { name, span, .. } => {
                self.call_named(name, Vec::new(), Some(span.line))
            }
            Expr::Member {
                object, name, span, ..
            } => {
                let receiver = self.eval_object(object, frame, span.line)?;
                self.host_call(Some(&receiver), name, &[], span.line)?
                    .ok_or_else(|| {
                        error(
                            RuntimeErrorKind::Unsupported,
                            format!("host method is not available: {}.{name}", receiver.kind),
                            Some(span.line),
                        )
                    })
            }
            Expr::WithMember(name, span) | Expr::WithBangMember(name, span) => {
                let receiver = current_with_object(frame, span.line)?;
                if let Some(value) = self.host_call(Some(&receiver), name, &[], span.line)? {
                    return Ok(value);
                }
                self.host_get(&receiver, name, span.line)?.ok_or_else(|| {
                    error(
                        RuntimeErrorKind::Unsupported,
                        format!("host member is not available: {}.{name}", receiver.kind),
                        Some(span.line),
                    )
                })
            }
            _ => Err(error(
                RuntimeErrorKind::Unsupported,
                "call target is not executable yet",
                Some(expr.span().line),
            )),
        }
    }

    fn call_user_procedure(
        &mut self,
        name: &str,
        args: &[Argument],
        force_by_value: bool,
        frame: &mut Frame,
        line: Option<u32>,
    ) -> Result<Value, RuntimeError> {
        let procedure = self.find_procedure(name, line)?;
        let param_array_index = procedure
            .params
            .iter()
            .position(|param| param.mode == ParamMode::ParamArray);
        let mut assigned = vec![None; procedure.params.len()];
        let mut param_array_args = Vec::new();
        let mut next_positional = 0;
        let mut saw_named = false;
        for (argument_index, argument) in args.iter().enumerate() {
            if let Some(argument_name) = &argument.name {
                saw_named = true;
                let parameter_index = procedure
                    .params
                    .iter()
                    .position(|param| param.name.eq_ignore_ascii_case(argument_name))
                    .ok_or_else(|| {
                        error(
                            RuntimeErrorKind::ArgumentCount,
                            format!("named argument not found: {argument_name}"),
                            line,
                        )
                    })?;
                if Some(parameter_index) == param_array_index {
                    return Err(error(
                        RuntimeErrorKind::ArgumentCount,
                        "ParamArray cannot be supplied as a named argument",
                        line,
                    ));
                }
                if assigned[parameter_index].replace(argument_index).is_some() {
                    return Err(error(
                        RuntimeErrorKind::ArgumentCount,
                        format!("argument supplied more than once: {argument_name}"),
                        line,
                    ));
                }
                continue;
            }
            if saw_named {
                return Err(error(
                    RuntimeErrorKind::ArgumentCount,
                    "positional argument cannot follow a named argument",
                    line,
                ));
            }
            while next_positional < procedure.params.len() && assigned[next_positional].is_some() {
                next_positional += 1;
            }
            if Some(next_positional) == param_array_index {
                param_array_args.push(argument_index);
            } else if next_positional < procedure.params.len() {
                assigned[next_positional] = Some(argument_index);
                next_positional += 1;
            } else {
                return Err(argument_count_error(&procedure, args.len(), line));
            }
        }

        let mut bound = Vec::with_capacity(procedure.params.len());
        let mut copybacks = Vec::<(ValueSlot, i64, ValueSlot)>::new();
        for (parameter_index, parameter) in procedure.params.iter().enumerate() {
            if parameter.mode == ParamMode::ParamArray {
                let mut values = Vec::with_capacity(param_array_args.len());
                for argument_index in &param_array_args {
                    let argument = &args[*argument_index];
                    let value = match &argument.value {
                        Some(expression) => self.eval_expr(expression, frame)?,
                        None => Value::Missing,
                    };
                    values.push(value);
                }
                bound.push(BoundArgument::Value(Value::Array(ArrayValue {
                    lower_bound: 0,
                    values,
                    element_default: Box::new(default_value(&parameter.type_name)),
                })));
                continue;
            }
            let Some(argument_index) = assigned[parameter_index] else {
                bound.push(BoundArgument::Value(
                    self.omitted_parameter_value(parameter, line)?,
                ));
                continue;
            };
            let argument = &args[argument_index];
            let Some(expression) = argument.value.as_ref() else {
                bound.push(BoundArgument::Value(
                    self.omitted_parameter_value(parameter, line)?,
                ));
                continue;
            };
            if parameter.mode == ParamMode::ByRef && !force_by_value && !argument.force_by_value {
                if let Some(name) = expr_name(expression) {
                    if let Some(value) = frame.values.get(&key(name)) {
                        bound.push(BoundArgument::Reference(value.clone()));
                        continue;
                    }
                }
                if let Expr::Index { target, args, .. } = expression {
                    let array = expr_name(target)
                        .and_then(|name| frame.values.get(&key(name)))
                        .filter(|value| matches!(&*value.borrow(), Value::Array(_)))
                        .cloned();
                    if let Some(array) = array {
                        let index = self.single_array_argument(
                            args,
                            frame,
                            line.unwrap_or(expression.span().line),
                        )?;
                        if let Some((_, _, value)) =
                            copybacks.iter().find(|(existing, existing_index, _)| {
                                Rc::ptr_eq(existing, &array) && *existing_index == index
                            })
                        {
                            bound.push(BoundArgument::Reference(value.clone()));
                            continue;
                        }
                        let value = Rc::new(RefCell::new(self.eval_expr(expression, frame)?));
                        bound.push(BoundArgument::Reference(value.clone()));
                        copybacks.push((array, index, value));
                        continue;
                    }
                }
            }
            bound.push(BoundArgument::Value(self.eval_expr(expression, frame)?));
        }

        let result = self.invoke_procedure(&procedure, bound, line)?;
        for (array, index, value) in copybacks {
            let value = value.borrow().clone();
            let mut array = array.borrow_mut();
            let Value::Array(array) = &mut *array else {
                return Err(error(
                    RuntimeErrorKind::TypeMismatch,
                    "ByRef array target is no longer an array",
                    line,
                ));
            };
            let offset = array_offset(array, index, line.unwrap_or(0))?;
            array.values[offset] = value;
        }
        Ok(result)
    }

    fn call_named(
        &mut self,
        name: &str,
        args: Vec<Value>,
        line: Option<u32>,
    ) -> Result<Value, RuntimeError> {
        if self.module.items.iter().any(
            |item| matches!(item, ModuleItem::Procedure(p) if p.name.eq_ignore_ascii_case(name)),
        ) {
            return self.call_procedure(name, args, line);
        }
        if let Some(result) = call_builtin(name, &args, line) {
            return result;
        }
        if let Some(value) = self.host_call(None, name, &args, line.unwrap_or(0))? {
            return Ok(value);
        }
        self.call_procedure(name, args, line)
    }

    fn eval_object(
        &mut self,
        expr: &Expr,
        frame: &mut Frame,
        line: u32,
    ) -> Result<ObjectRef, RuntimeError> {
        let value = self.eval_expr(expr, frame).map_err(|failure| {
            if failure.kind == RuntimeErrorKind::ProcedureNotFound {
                error(
                    RuntimeErrorKind::Unsupported,
                    "object expression requires a configured VBA host",
                    Some(line),
                )
            } else {
                failure
            }
        })?;
        match value {
            Value::Object(object) => Ok(object),
            _ => Err(error(
                RuntimeErrorKind::TypeMismatch,
                "VBA member access requires an object",
                Some(line),
            )),
        }
    }

    fn host_call(
        &mut self,
        receiver: Option<&ObjectRef>,
        name: &str,
        args: &[Value],
        line: u32,
    ) -> Result<Option<Value>, RuntimeError> {
        let Some(host) = self.host.as_deref_mut() else {
            return Ok(None);
        };
        host.call(receiver, name, args)
            .map_err(|message| error(RuntimeErrorKind::Host, message, Some(line)))
    }

    fn host_get(
        &mut self,
        receiver: &ObjectRef,
        name: &str,
        line: u32,
    ) -> Result<Option<Value>, RuntimeError> {
        let Some(host) = self.host.as_deref_mut() else {
            return Ok(None);
        };
        host.get(receiver, name)
            .map_err(|message| error(RuntimeErrorKind::Host, message, Some(line)))
    }

    fn host_set(
        &mut self,
        receiver: &ObjectRef,
        name: &str,
        value: Value,
        line: u32,
    ) -> Result<bool, RuntimeError> {
        let Some(host) = self.host.as_deref_mut() else {
            return Ok(false);
        };
        host.set(receiver, name, value)
            .map_err(|message| error(RuntimeErrorKind::Host, message, Some(line)))
    }

    fn host_enumerate(
        &mut self,
        receiver: &ObjectRef,
        line: u32,
    ) -> Result<Option<Vec<Value>>, RuntimeError> {
        let Some(host) = self.host.as_deref_mut() else {
            return Ok(None);
        };
        host.enumerate(receiver)
            .map_err(|message| error(RuntimeErrorKind::Host, message, Some(line)))
    }

    fn single_array_argument(
        &mut self,
        args: &[Argument],
        frame: &mut Frame,
        line: u32,
    ) -> Result<i64, RuntimeError> {
        if args.len() != 1 || args[0].name.is_some() {
            return Err(error(
                RuntimeErrorKind::SubscriptOutOfRange,
                "one-dimensional VBA array requires exactly one index",
                Some(line),
            ));
        }
        let index = args[0].value.as_ref().ok_or_else(|| {
            error(
                RuntimeErrorKind::SubscriptOutOfRange,
                "VBA array index cannot be omitted",
                Some(line),
            )
        })?;
        self.array_index(index, frame, line)
    }

    fn array_index(
        &mut self,
        expr: &Expr,
        frame: &mut Frame,
        line: u32,
    ) -> Result<i64, RuntimeError> {
        let value = self.eval_expr(expr, frame)?;
        let number = number(&value)
            .map_err(|message| error(RuntimeErrorKind::TypeMismatch, message, Some(line)))?;
        let rounded = number.round_ties_even();
        if !rounded.is_finite() || rounded < i64::MIN as f64 || rounded > i64::MAX as f64 {
            Err(error(
                RuntimeErrorKind::Overflow,
                "VBA array index is outside the supported integer range",
                Some(line),
            ))
        } else {
            Ok(rounded as i64)
        }
    }

    fn tick(&mut self, line: Option<u32>) -> Result<(), RuntimeError> {
        self.steps += 1;
        if self.steps > self.max_steps {
            Err(error(
                RuntimeErrorKind::StepLimit,
                "VBA execution step limit exceeded",
                line,
            ))
        } else {
            Ok(())
        }
    }
}

fn error(kind: RuntimeErrorKind, message: impl Into<String>, line: Option<u32>) -> RuntimeError {
    RuntimeError {
        kind,
        message: message.into(),
        line,
    }
}

fn expr_name(expr: &Expr) -> Option<&str> {
    match expr {
        Expr::Ident(name, _) | Expr::TypedIdent { name, .. } => Some(name),
        _ => None,
    }
}

fn current_with_object(frame: &Frame, line: u32) -> Result<ObjectRef, RuntimeError> {
    frame.with_objects.last().cloned().ok_or_else(|| {
        error(
            RuntimeErrorKind::Unsupported,
            "With member used outside a With block",
            Some(line),
        )
    })
}

fn array_offset(array: &ArrayValue, index: i64, line: u32) -> Result<usize, RuntimeError> {
    if array.values.is_empty() || index < array.lower_bound || index > array.upper_bound() {
        return Err(error(
            RuntimeErrorKind::SubscriptOutOfRange,
            format!(
                "VBA array index {index} is outside {} To {}",
                array.lower_bound,
                array.upper_bound()
            ),
            Some(line),
        ));
    }
    Ok((index - array.lower_bound) as usize)
}

fn call_builtin(
    name: &str,
    args: &[Value],
    line: Option<u32>,
) -> Option<Result<Value, RuntimeError>> {
    let name = name.to_ascii_lowercase();
    let known = matches!(
        name.as_str(),
        "abs"
            | "array"
            | "cbool"
            | "cdbl"
            | "clng"
            | "cstr"
            | "ismissing"
            | "lbound"
            | "lcase"
            | "len"
            | "trim"
            | "ubound"
            | "ucase"
    );
    if !known {
        return None;
    }
    Some((|| {
        if name == "array" {
            return Ok(Value::Array(ArrayValue {
                lower_bound: 0,
                values: args.to_vec(),
                element_default: Box::new(Value::Empty),
            }));
        }
        if name == "ismissing" {
            if args.len() != 1 {
                return Err(error(
                    RuntimeErrorKind::ArgumentCount,
                    format!("ismissing expects 1 argument, received {}", args.len()),
                    line,
                ));
            }
            return Ok(Value::Boolean(matches!(args[0], Value::Missing)));
        }
        if matches!(name.as_str(), "lbound" | "ubound") {
            if !(1..=2).contains(&args.len()) {
                return Err(error(
                    RuntimeErrorKind::ArgumentCount,
                    format!("{name} expects 1 or 2 arguments, received {}", args.len()),
                    line,
                ));
            }
            if args.len() == 2 && number(&args[1]).ok().map(f64::round_ties_even) != Some(1.0) {
                return Err(error(
                    RuntimeErrorKind::SubscriptOutOfRange,
                    "only array dimension 1 is supported",
                    line,
                ));
            }
            let Value::Array(array) = &args[0] else {
                return Err(error(
                    RuntimeErrorKind::TypeMismatch,
                    format!("{name} requires an array"),
                    line,
                ));
            };
            return Ok(Value::Integer(if name == "lbound" {
                array.lower_bound
            } else {
                array.upper_bound()
            }));
        }
        if args.len() != 1 {
            return Err(error(
                RuntimeErrorKind::ArgumentCount,
                format!("{name} expects 1 argument, received {}", args.len()),
                line,
            ));
        }
        let value = &args[0];
        let mismatch = |message| error(RuntimeErrorKind::TypeMismatch, message, line);
        match name.as_str() {
            "abs" => match value {
                Value::Null => Ok(Value::Null),
                Value::Integer(value) => value
                    .checked_abs()
                    .map(Value::Integer)
                    .ok_or_else(|| error(RuntimeErrorKind::Overflow, "overflow in Abs", line)),
                _ => Ok(numeric_literal(number(value).map_err(mismatch)?.abs())),
            },
            "cbool" => match value {
                Value::Null => Err(mismatch("invalid use of Null".to_string())),
                _ => Ok(Value::Boolean(truthy(value).map_err(mismatch)?)),
            },
            "cdbl" => Ok(Value::Double(number(value).map_err(mismatch)?)),
            "clng" => {
                let value = number(value).map_err(mismatch)?.round_ties_even();
                if !(-2_147_483_648.0..=2_147_483_647.0).contains(&value) {
                    Err(error(
                        RuntimeErrorKind::Overflow,
                        "overflow converting value to Long",
                        line,
                    ))
                } else {
                    Ok(Value::Integer(value as i64))
                }
            }
            "cstr" => match value {
                Value::Null => Err(mismatch("invalid use of Null".to_string())),
                _ => Ok(Value::String(text(value).map_err(mismatch)?)),
            },
            "lcase" => match value {
                Value::Null => Ok(Value::Null),
                _ => Ok(Value::String(text(value).map_err(mismatch)?.to_lowercase())),
            },
            "len" => match value {
                Value::Null => Err(mismatch("invalid use of Null".to_string())),
                _ => Ok(Value::Integer(
                    text(value).map_err(mismatch)?.encode_utf16().count() as i64,
                )),
            },
            "trim" => match value {
                Value::Null => Ok(Value::Null),
                _ => Ok(Value::String(
                    text(value).map_err(mismatch)?.trim_matches(' ').to_string(),
                )),
            },
            "ucase" => match value {
                Value::Null => Ok(Value::Null),
                _ => Ok(Value::String(text(value).map_err(mismatch)?.to_uppercase())),
            },
            _ => unreachable!(),
        }
    })())
}

fn key(name: &str) -> String {
    name.to_ascii_lowercase()
}

fn literal_value(literal: &Literal) -> Value {
    match literal {
        Literal::Number(value) | Literal::TypedNumber { value, .. } => numeric_literal(*value),
        Literal::LargeInteger { digits, .. } => digits
            .parse::<i64>()
            .map(Value::Integer)
            .unwrap_or_else(|_| Value::Double(digits.parse().unwrap_or(f64::INFINITY))),
        Literal::Str(value) | Literal::Date(value) => Value::String(value.clone()),
        Literal::Bool(value) => Value::Boolean(*value),
        Literal::Empty | Literal::Nothing => Value::Empty,
        Literal::Null => Value::Null,
    }
}

fn numeric_literal(value: f64) -> Value {
    if value.fract() == 0.0 && value >= i64::MIN as f64 && value <= i64::MAX as f64 {
        Value::Integer(value as i64)
    } else {
        Value::Double(value)
    }
}

fn default_value(type_name: &TypeName) -> Value {
    match type_name.name.to_ascii_lowercase().as_str() {
        "boolean" => Value::Boolean(false),
        "byte" | "integer" | "long" | "longlong" | "longptr" | "currency" => Value::Integer(0),
        "single" | "double" | "decimal" => Value::Double(0.0),
        "string" => Value::String(String::new()),
        _ => Value::Empty,
    }
}

fn default_return_value(procedure: &Procedure) -> Value {
    procedure
        .return_type
        .as_ref()
        .map(default_value)
        .unwrap_or(Value::Empty)
}

fn argument_count_error(procedure: &Procedure, received: usize, line: Option<u32>) -> RuntimeError {
    error(
        RuntimeErrorKind::ArgumentCount,
        format!(
            "{} cannot bind {} supplied argument(s) to {} parameter(s)",
            procedure.name,
            received,
            procedure.params.len()
        ),
        line,
    )
}

fn number(value: &Value) -> Result<f64, String> {
    match value {
        Value::Empty | Value::Boolean(false) => Ok(0.0),
        Value::Boolean(true) => Ok(-1.0),
        Value::Integer(value) => Ok(*value as f64),
        Value::Double(value) => Ok(*value),
        Value::String(value) => value
            .parse()
            .map_err(|_| "type mismatch converting String to number".to_string()),
        Value::Null => Err("invalid use of Null".to_string()),
        Value::Missing => Err("invalid use of Missing".to_string()),
        Value::Array(_) => Err("type mismatch converting array to number".to_string()),
        Value::Object(_) => Err("type mismatch converting object to number".to_string()),
    }
}

fn truthy(value: &Value) -> Result<bool, String> {
    match value {
        Value::Empty | Value::Null | Value::Boolean(false) => Ok(false),
        Value::Boolean(true) => Ok(true),
        Value::Integer(value) => Ok(*value != 0),
        Value::Double(value) => Ok(*value != 0.0),
        Value::String(value) if value.is_empty() => Ok(false),
        Value::String(value) if value.eq_ignore_ascii_case("true") => Ok(true),
        Value::String(value) if value.eq_ignore_ascii_case("false") => Ok(false),
        Value::String(value) => value
            .parse::<f64>()
            .map(|number| number != 0.0)
            .map_err(|_| "type mismatch converting String to Boolean".to_string()),
        Value::Array(_) => Err("type mismatch converting array to Boolean".to_string()),
        Value::Object(_) => Err("type mismatch converting object to Boolean".to_string()),
        Value::Missing => Err("invalid use of Missing".to_string()),
    }
}

fn unary(op: UnaryOp, value: Value) -> Result<Value, String> {
    match op {
        UnaryOp::Plus => Ok(numeric_literal(number(&value)?)),
        UnaryOp::Neg => Ok(numeric_literal(-number(&value)?)),
        UnaryOp::Not => match value {
            Value::Boolean(value) => Ok(Value::Boolean(!value)),
            Value::Integer(value) => Ok(Value::Integer(!value)),
            other => Ok(Value::Boolean(!truthy(&other)?)),
        },
    }
}

fn binary(op: BinaryOp, lhs: Value, rhs: Value) -> Result<Value, (RuntimeErrorKind, String)> {
    use BinaryOp::*;
    if matches!(lhs, Value::Array(_) | Value::Object(_) | Value::Missing)
        || matches!(rhs, Value::Array(_) | Value::Object(_) | Value::Missing)
    {
        return Err((
            RuntimeErrorKind::TypeMismatch,
            "VBA arrays and objects cannot be used as scalar operands".to_string(),
        ));
    }
    if matches!(lhs, Value::Null) || matches!(rhs, Value::Null) {
        return Ok(Value::Null);
    }
    let mismatch = |message| (RuntimeErrorKind::TypeMismatch, message);
    let numbers = || {
        Ok::<_, (RuntimeErrorKind, String)>((
            number(&lhs).map_err(mismatch)?,
            number(&rhs).map_err(mismatch)?,
        ))
    };
    match op {
        Concat => Ok(Value::String(format!(
            "{}{}",
            text(&lhs).map_err(mismatch)?,
            text(&rhs).map_err(mismatch)?
        ))),
        Add | Sub | Mul | Div | IntDiv | Mod | Pow => {
            let (a, b) = numbers()?;
            if matches!(op, Div | IntDiv | Mod) && b == 0.0 {
                return Err((
                    RuntimeErrorKind::DivisionByZero,
                    "division by zero".to_string(),
                ));
            }
            Ok(match op {
                Add => numeric_result(a + b, &lhs, &rhs),
                Sub => numeric_result(a - b, &lhs, &rhs),
                Mul => numeric_result(a * b, &lhs, &rhs),
                Div => Value::Double(a / b),
                IntDiv => Value::Integer((a / b).trunc() as i64),
                Mod => Value::Integer((a as i64) % (b as i64)),
                Pow => Value::Double(a.powf(b)),
                _ => unreachable!(),
            })
        }
        Eq | Ne | Lt | Le | Gt | Ge => {
            let ordering = match (&lhs, &rhs) {
                (Value::String(a), Value::String(b)) => a.partial_cmp(b),
                _ => {
                    let (a, b) = numbers()?;
                    a.partial_cmp(&b)
                }
            };
            let equal = lhs == rhs || ordering == Some(std::cmp::Ordering::Equal);
            Ok(Value::Boolean(match op {
                Eq => equal,
                Ne => !equal,
                Lt => ordering == Some(std::cmp::Ordering::Less),
                Le => ordering != Some(std::cmp::Ordering::Greater),
                Gt => ordering == Some(std::cmp::Ordering::Greater),
                Ge => ordering != Some(std::cmp::Ordering::Less),
                _ => unreachable!(),
            }))
        }
        And | Or | Xor | Eqv | Imp => {
            let a = truthy(&lhs).map_err(mismatch)?;
            let b = truthy(&rhs).map_err(mismatch)?;
            Ok(Value::Boolean(match op {
                And => a && b,
                Or => a || b,
                Xor => a ^ b,
                Eqv => a == b,
                Imp => !a || b,
                _ => unreachable!(),
            }))
        }
        Is | Like => Err((
            RuntimeErrorKind::Unsupported,
            format!("operator {op:?} is not executable yet"),
        )),
    }
}

fn numeric_result(value: f64, lhs: &Value, rhs: &Value) -> Value {
    if matches!(lhs, Value::Integer(_)) && matches!(rhs, Value::Integer(_)) && value.fract() == 0.0
    {
        Value::Integer(value as i64)
    } else {
        Value::Double(value)
    }
}

fn text(value: &Value) -> Result<String, String> {
    Ok(match value {
        Value::Empty => String::new(),
        Value::Missing => return Err("invalid use of Missing".to_string()),
        Value::Null => "Null".to_string(),
        Value::Boolean(true) => "True".to_string(),
        Value::Boolean(false) => "False".to_string(),
        Value::Integer(value) => value.to_string(),
        Value::Double(value) => value.to_string(),
        Value::String(value) => value.clone(),
        Value::Array(_) => return Err("type mismatch converting array to String".to_string()),
        Value::Object(_) => return Err("type mismatch converting object to String".to_string()),
    })
}

fn line_of(statement: &Statement) -> Option<u32> {
    match statement {
        Statement::Assign { span, .. }
        | Statement::SetAssign { span, .. }
        | Statement::ReDim { span, .. }
        | Statement::Call { span, .. }
        | Statement::Exit { span, .. }
        | Statement::End { span }
        | Statement::Comment { span, .. }
        | Statement::Label { span, .. }
        | Statement::LineNumber { span, .. }
        | Statement::Unknown { span, .. } => Some(span.line),
        Statement::Dim(decl) => Some(decl.span.line),
        Statement::If(branch) => Some(branch.span.line),
        Statement::SelectCase(select) => Some(select.span.line),
        Statement::For(loop_) => Some(loop_.span.line),
        Statement::ForEach(loop_) => Some(loop_.span.line),
        Statement::Do(loop_) => Some(loop_.span.line),
        Statement::While { span, .. } => Some(span.line),
        _ => None,
    }
}

fn statement_name(statement: &Statement) -> &'static str {
    match statement {
        Statement::SetAssign { .. } => "Set",
        Statement::With { .. } => "With",
        Statement::OnError(_) => "On Error",
        _ => "host-dependent statement",
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::parse_module;

    #[derive(Default)]
    struct SheetHost {
        cells: BTreeMap<(u32, u32), Value>,
    }

    impl SheetHost {
        fn cell_object(row: u32, column: u32) -> Value {
            Value::Object(ObjectRef {
                handle: ((row as u64) << 32) | column as u64,
                kind: "Cell".to_string(),
            })
        }

        fn coordinates(object: &ObjectRef) -> Option<(u32, u32)> {
            (object.kind == "Cell").then_some(((object.handle >> 32) as u32, object.handle as u32))
        }
    }

    impl Host for SheetHost {
        fn call(
            &mut self,
            receiver: Option<&ObjectRef>,
            name: &str,
            args: &[Value],
        ) -> Result<Option<Value>, String> {
            if let Some(receiver) = receiver {
                if name.eq_ignore_ascii_case("offset") {
                    let Some((row, column)) = Self::coordinates(receiver) else {
                        return Ok(None);
                    };
                    let [row_offset, column_offset] = args else {
                        return Err("Offset expects row and column offsets".to_string());
                    };
                    let row = row as i64 + number(row_offset)?.round_ties_even() as i64;
                    let column = column as i64 + number(column_offset)?.round_ties_even() as i64;
                    if row < 1 || column < 1 {
                        return Err("Offset moved outside the sheet".to_string());
                    }
                    return Ok(Some(Self::cell_object(row as u32, column as u32)));
                }
                return Ok(None);
            }
            if name.eq_ignore_ascii_case("range") {
                let [Value::String(address)] = args else {
                    return Err("Range expects one A1 address".to_string());
                };
                let address = address.to_ascii_uppercase();
                let column = address
                    .bytes()
                    .next()
                    .filter(u8::is_ascii_alphabetic)
                    .map(|value| (value - b'A' + 1) as u32)
                    .ok_or_else(|| "invalid Range address".to_string())?;
                let row = address[1..]
                    .parse::<u32>()
                    .map_err(|_| "invalid Range address".to_string())?;
                return Ok(Some(Self::cell_object(row, column)));
            }
            if name.eq_ignore_ascii_case("cells") {
                let [row, column] = args else {
                    return Err("Cells expects row and column".to_string());
                };
                let row = number(row)?.round_ties_even() as u32;
                let column = number(column)?.round_ties_even() as u32;
                return Ok(Some(Self::cell_object(row, column)));
            }
            Ok(None)
        }

        fn get(&mut self, receiver: &ObjectRef, name: &str) -> Result<Option<Value>, String> {
            let Some(coordinates) = Self::coordinates(receiver) else {
                return Ok(None);
            };
            if name.eq_ignore_ascii_case("value") {
                return Ok(Some(
                    self.cells
                        .get(&coordinates)
                        .cloned()
                        .unwrap_or(Value::Empty),
                ));
            }
            Ok(None)
        }

        fn set(&mut self, receiver: &ObjectRef, name: &str, value: Value) -> Result<bool, String> {
            let Some(coordinates) = Self::coordinates(receiver) else {
                return Ok(false);
            };
            if name.eq_ignore_ascii_case("value") {
                self.cells.insert(coordinates, value);
                return Ok(true);
            }
            Ok(false)
        }
    }

    fn run(source: &str, name: &str, args: Vec<Value>) -> Result<Value, RuntimeError> {
        let module = parse_module(source).unwrap();
        execute(&module, name, args)
    }

    #[test]
    fn executes_a_pure_function_with_locals_arithmetic_and_if() {
        let value = run(
            "Option Explicit\n\
             Public Function NetPrice(total As Double, rate As Double, preferred As Boolean) As Double\n\
             Dim result As Double\n\
             If preferred Then\n\
               result = total * (1 - rate)\n\
             Else\n\
               result = total\n\
             End If\n\
             NetPrice = result\n\
             End Function\n",
            "netprice",
            vec![Value::Double(200.0), Value::Double(0.1), Value::Boolean(true)],
        )
        .unwrap();
        assert_eq!(value, Value::Double(180.0));
    }

    #[test]
    fn reads_and_writes_range_and_cells_through_a_host() {
        let module = parse_module(
            "Public Sub WriteSheet()\n\
               Range(\"A1\").Value = 40\n\
               Cells(2, 1).Value = 2\n\
             End Sub\n\
             Public Function SheetTotal() As Long\n\
               SheetTotal = Range(\"A1\").Value + Cells(2, 1).Value\n\
             End Function\n",
        )
        .unwrap();
        let mut host = SheetHost::default();
        let value = {
            let mut runtime = Runtime::new(&module).with_host(&mut host);
            assert_eq!(runtime.call("WriteSheet", vec![]).unwrap(), Value::Empty);
            runtime.call("SheetTotal", vec![]).unwrap()
        };
        assert_eq!(value, Value::Integer(42));
        assert_eq!(host.cells.get(&(1, 1)), Some(&Value::Integer(40)));
        assert_eq!(host.cells.get(&(2, 1)), Some(&Value::Integer(2)));
    }

    #[test]
    fn set_binds_a_host_object_to_a_local_variable() {
        let module = parse_module(
            "Public Function WriteThroughObject() As Long\n\
               Dim cell As Object\n\
               Set cell = Range(\"A1\")\n\
               cell.Value = 42\n\
               WriteThroughObject = cell.Value\n\
             End Function\n",
        )
        .unwrap();
        let mut host = SheetHost::default();
        let value = execute_with_host(&module, "WriteThroughObject", vec![], &mut host).unwrap();
        assert_eq!(value, Value::Integer(42));
        assert_eq!(host.cells.get(&(1, 1)), Some(&Value::Integer(42)));
    }

    #[test]
    fn with_blocks_read_write_call_and_restore_nested_host_objects() {
        let module = parse_module(
            "Public Function WithProbe() As Long\n\
               With Range(\"A1\")\n\
                 .Value = 10\n\
                 With .Offset(1, 0)\n\
                   .Value = 32\n\
                 End With\n\
                 WithProbe = .Value + .Offset(1, 0).Value\n\
               End With\n\
             End Function\n",
        )
        .unwrap();
        let mut host = SheetHost::default();
        let value = execute_with_host(&module, "WithProbe", vec![], &mut host).unwrap();

        assert_eq!(value, Value::Integer(42));
        assert_eq!(host.cells.get(&(1, 1)), Some(&Value::Integer(10)));
        assert_eq!(host.cells.get(&(2, 1)), Some(&Value::Integer(32)));
    }

    #[test]
    fn a_with_member_outside_a_with_block_fails_at_its_source_line() {
        let failure = run(
            "Public Function InvalidWith() As Variant\n\
               InvalidWith = .Value\n\
             End Function\n",
            "InvalidWith",
            vec![],
        )
        .unwrap_err();

        assert_eq!(failure.kind, RuntimeErrorKind::Unsupported);
        assert_eq!(failure.line, Some(2));
        assert!(failure.message.contains("outside a With block"));
    }

    #[test]
    fn set_rejects_scalar_values() {
        let failure = run(
            "Public Sub InvalidSet()\n\
               Dim value As Object\n\
               Set value = 42\n\
             End Sub\n",
            "InvalidSet",
            vec![],
        )
        .unwrap_err();
        assert_eq!(failure.kind, RuntimeErrorKind::TypeMismatch);
        assert_eq!(failure.line, Some(3));
    }

    #[test]
    fn reports_an_unavailable_host_property_at_its_source_line() {
        let module = parse_module(
            "Public Function Missing() As Variant\n\
               Missing = Range(\"A1\").Formula\n\
             End Function\n",
        )
        .unwrap();
        let mut host = SheetHost::default();
        let failure = execute_with_host(&module, "Missing", vec![], &mut host).unwrap_err();
        assert_eq!(failure.kind, RuntimeErrorKind::Unsupported);
        assert_eq!(failure.line, Some(2));
    }

    #[test]
    fn reports_host_failures_at_the_vba_call_site() {
        let module = parse_module(
            "Public Function Broken() As Variant\n\
               Broken = Range(\"bad\").Value\n\
             End Function\n",
        )
        .unwrap();
        let mut host = SheetHost::default();
        let failure = execute_with_host(&module, "Broken", vec![], &mut host).unwrap_err();
        assert_eq!(failure.kind, RuntimeErrorKind::Host);
        assert_eq!(failure.line, Some(2));
        assert!(failure.message.contains("invalid Range address"));
    }

    #[test]
    fn calls_another_vba_function_case_insensitively() {
        let value = run(
            "Private Function Twice(value As Long) As Long\n\
               Twice = value * 2\n\
             End Function\n\
             Public Function Invoice(value As Long) As Long\n\
               Invoice = TWICE(value) + 1\n\
             End Function\n",
            "invoice",
            vec![Value::Integer(20)],
        )
        .unwrap();
        assert_eq!(value, Value::Integer(41));
    }

    #[test]
    fn byref_parameters_share_the_callers_storage_while_byval_does_not() {
        let value = run(
            "Private Sub Combine(ByRef firstValue As Long, ByRef secondValue As Long)\n\
               firstValue = firstValue + 1\n\
               secondValue = firstValue + secondValue\n\
             End Sub\n\
             Private Sub Replace(ByRef value As Long)\n\
               value = 99\n\
             End Sub\n\
             Private Sub ReplaceCopy(ByVal value As Long)\n\
               value = 88\n\
             End Sub\n\
             Public Function ProbeByRef() As String\n\
               Dim shared As Long\n\
               Dim copied As Long\n\
               Dim bareCopied As Long\n\
               shared = 1\n\
               copied = 2\n\
               bareCopied = 3\n\
               Combine shared, shared\n\
               ReplaceCopy copied\n\
               Call Replace((copied))\n\
               Replace (bareCopied)\n\
               ProbeByRef = shared & \"|\" & copied & \"|\" & bareCopied\n\
             End Function\n",
            "ProbeByRef",
            vec![],
        )
        .unwrap();

        assert_eq!(value, Value::String("4|2|3".to_string()));
    }

    #[test]
    fn byref_array_parameters_and_elements_update_the_caller() {
        let value = run(
            "Private Sub Grow(ByRef values() As Long)\n\
               values(1) = 10\n\
               ReDim Preserve values(1 To 3)\n\
               values(3) = 30\n\
             End Sub\n\
             Private Sub Increment(ByRef value As Long)\n\
               value = value + 1\n\
             End Sub\n\
             Private Sub CombineElements(ByRef firstValue As Long, ByRef secondValue As Long)\n\
               firstValue = firstValue + 1\n\
               secondValue = firstValue + secondValue\n\
             End Sub\n\
             Public Function ArrayByRef() As Long\n\
               Dim values() As Long\n\
               ReDim values(1 To 2)\n\
               values(2) = 20\n\
               Grow values\n\
               Increment values(2)\n\
               CombineElements values(2), values(2)\n\
               ArrayByRef = values(1) + values(2) + values(3) + UBound(values) * 100\n\
             End Function\n",
            "ArrayByRef",
            vec![],
        )
        .unwrap();

        assert_eq!(value, Value::Integer(384));
    }

    #[test]
    fn byref_aliases_remain_shared_through_recursive_calls() {
        let value = run(
            "Private Sub CountDown(ByRef value As Long)\n\
               If value > 0 Then\n\
                 value = value - 1\n\
                 CountDown value\n\
               End If\n\
             End Sub\n\
             Public Function RecursiveByRef() As Long\n\
               Dim value As Long\n\
               value = 4\n\
               CountDown value\n\
               RecursiveByRef = value\n\
             End Function\n",
            "RecursiveByRef",
            vec![],
        )
        .unwrap();

        assert_eq!(value, Value::Integer(0));
    }

    #[test]
    fn binds_optional_omitted_and_named_arguments() {
        let value = run(
            "Private Function Describe(ByVal number As Long, Optional label As String = \"item\", Optional count As Long = 3) As String\n\
               Describe = label & \"=\" & (number * count)\n\
             End Function\n\
             Public Function OptionalProbe() As String\n\
               OptionalProbe = Describe(2) & \"|\" & Describe(4, , 5) & \"|\" & Describe(count:=2, number:=6, label:=\"named\")\n\
             End Function\n",
            "OptionalProbe",
            vec![],
        )
        .unwrap();

        assert_eq!(value, Value::String("item=6|item=20|named=12".to_string()));
    }

    #[test]
    fn named_byref_arguments_alias_the_matching_caller_variables() {
        let value = run(
            "Private Sub ReplaceBoth(ByRef first As Long, ByRef second As Long)\n\
               first = 10\n\
               second = 20\n\
             End Sub\n\
             Public Function NamedByRefProbe() As Long\n\
               Dim leftValue As Long\n\
               Dim rightValue As Long\n\
               ReplaceBoth second:=rightValue, first:=leftValue\n\
               NamedByRefProbe = leftValue * 100 + rightValue\n\
             End Function\n",
            "NamedByRefProbe",
            vec![],
        )
        .unwrap();

        assert_eq!(value, Value::Integer(1020));
    }

    #[test]
    fn collects_paramarray_values_and_supports_an_empty_paramarray() {
        let value = run(
            "Private Function SumMany(ByVal base As Long, ParamArray values() As Variant) As Long\n\
               Dim item As Variant\n\
               SumMany = base\n\
               For Each item In values\n\
                 SumMany = SumMany + item\n\
               Next\n\
             End Function\n\
             Public Function ParamArrayProbe() As Long\n\
               ParamArrayProbe = SumMany(10) * 100 + SumMany(10, 1, 2, 3)\n\
             End Function\n",
            "ParamArrayProbe",
            vec![],
        )
        .unwrap();

        assert_eq!(value, Value::Integer(1016));
    }

    #[test]
    fn ismissing_distinguishes_an_omitted_optional_variant_from_empty() {
        let value = run(
            "Private Function MissingFlag(Optional value As Variant) As Long\n\
               If IsMissing(value) Then\n\
                 MissingFlag = 1\n\
               Else\n\
                 MissingFlag = 0\n\
               End If\n\
             End Function\n\
             Public Function MissingProbe() As Long\n\
               Dim emptyValue As Variant\n\
               MissingProbe = MissingFlag() * 10 + MissingFlag(emptyValue)\n\
             End Function\n",
            "MissingProbe",
            vec![],
        )
        .unwrap();

        assert_eq!(value, Value::Integer(10));
    }

    #[test]
    fn direct_runtime_calls_apply_optional_and_paramarray_binding() {
        let optional = run(
            "Public Function AddDefault(ByVal value As Long, Optional amount As Long = 5) As Long\n\
               AddDefault = value + amount\n\
             End Function\n",
            "AddDefault",
            vec![Value::Integer(7)],
        )
        .unwrap();
        assert_eq!(optional, Value::Integer(12));

        let param_array = run(
            "Public Function CountValues(ParamArray values() As Variant) As Long\n\
               CountValues = UBound(values) - LBound(values) + 1\n\
             End Function\n",
            "CountValues",
            vec![Value::Integer(1), Value::Integer(2), Value::Integer(3)],
        )
        .unwrap();
        assert_eq!(param_array, Value::Integer(3));
    }

    #[test]
    fn rejects_unknown_and_duplicate_named_arguments() {
        let unknown = run(
            "Private Sub Target(Optional value As Long = 1)\n\
             End Sub\n\
             Public Sub Probe()\n\
               Target missing:=2\n\
             End Sub\n",
            "Probe",
            vec![],
        )
        .unwrap_err();
        assert_eq!(unknown.kind, RuntimeErrorKind::ArgumentCount);
        assert!(unknown.message.contains("named argument not found"));

        let duplicate = run(
            "Private Sub Target(Optional value As Long = 1)\n\
             End Sub\n\
             Public Sub Probe()\n\
               Target value:=2, value:=3\n\
             End Sub\n",
            "Probe",
            vec![],
        )
        .unwrap_err();
        assert_eq!(duplicate.kind, RuntimeErrorKind::ArgumentCount);
        assert!(duplicate.message.contains("more than once"));
    }

    #[test]
    fn concatenates_strings_and_uses_vba_true_as_minus_one_in_arithmetic() {
        let value = run(
            "Public Function Label(ok As Boolean) As String\n\
               Label = \"result=\" & (10 + ok)\n\
             End Function\n",
            "Label",
            vec![Value::Boolean(true)],
        )
        .unwrap();
        assert_eq!(value, Value::String("result=9".to_string()));
    }

    #[test]
    fn executes_for_loops_with_positive_and_negative_steps() {
        let value = run(
            "Public Function SumSteps() As Long\n\
               Dim total As Long\n\
               Dim i As Long\n\
               For i = 1 To 5\n\
                 total = total + i\n\
               Next i\n\
               For i = 5 To 1 Step -2\n\
                 total = total + i\n\
               Next i\n\
               SumSteps = total\n\
             End Function\n",
            "SumSteps",
            vec![],
        )
        .unwrap();
        assert_eq!(value, Value::Integer(24));
    }

    #[test]
    fn exit_for_only_leaves_the_nearest_for_loop() {
        let value = run(
            "Public Function FirstLarge() As Long\n\
               Dim i As Long\n\
               For i = 1 To 10\n\
                 If i = 4 Then Exit For\n\
               Next i\n\
               FirstLarge = i\n\
             End Function\n",
            "FirstLarge",
            vec![],
        )
        .unwrap();
        assert_eq!(value, Value::Integer(4));
    }

    #[test]
    fn executes_while_and_do_loops_and_consumes_exit_do() {
        let value = run(
            "Public Function CountUp() As Long\n\
               Dim n As Long\n\
               While n < 2\n\
                 n = n + 1\n\
               Wend\n\
               Do Until n = 4\n\
                 n = n + 1\n\
               Loop\n\
               Do\n\
                 n = n + 1\n\
                 If n = 6 Then Exit Do\n\
               Loop\n\
               CountUp = n\n\
             End Function\n",
            "CountUp",
            vec![],
        )
        .unwrap();
        assert_eq!(value, Value::Integer(6));
    }

    #[test]
    fn selects_value_range_comparison_and_else_cases() {
        let source = "Public Function Band(value As Long) As String\n\
               Select Case value\n\
               Case 1, 2\n\
                 Band = \"small\"\n\
               Case 3 To 5\n\
                 Band = \"medium\"\n\
               Case Is >= 6\n\
                 Band = \"large\"\n\
               Case Else\n\
                 Band = \"other\"\n\
               End Select\n\
             End Function\n";
        assert_eq!(
            run(source, "Band", vec![Value::Integer(2)]).unwrap(),
            Value::String("small".to_string())
        );
        assert_eq!(
            run(source, "Band", vec![Value::Integer(4)]).unwrap(),
            Value::String("medium".to_string())
        );
        assert_eq!(
            run(source, "Band", vec![Value::Integer(9)]).unwrap(),
            Value::String("large".to_string())
        );
        assert_eq!(
            run(source, "Band", vec![Value::Integer(0)]).unwrap(),
            Value::String("other".to_string())
        );
    }

    #[test]
    fn executes_common_conversion_string_and_math_builtins() {
        let value = run(
            "Public Function Summary() As String\n\
               Summary = UCase(Trim(\" oxi \")) & \"|\" & CStr(Len(\"A😀\")) & \"|\" & CStr(CLng(2.5)) & \"|\" & CStr(Abs(-3) + CDbl(\"2.5\")) & \"|\" & CStr(CBool(\"True\"))\n\
             End Function\n",
            "Summary",
            vec![],
        )
        .unwrap();
        assert_eq!(value, Value::String("OXI|3|2|5.5|True".to_string()));
    }

    #[test]
    fn a_user_procedure_shadows_a_builtin_name() {
        let value = run(
            "Private Function Len(value As String) As Long\n\
               Len = 99\n\
             End Function\n\
             Public Function Probe() As Long\n\
               Probe = Len(\"x\")\n\
             End Function\n",
            "Probe",
            vec![],
        )
        .unwrap();
        assert_eq!(value, Value::Integer(99));
    }

    #[test]
    fn reports_builtin_argument_errors_at_the_call_site() {
        let failure = run(
            "Public Function Broken() As Long\n\
               Broken = Len()\n\
             End Function\n",
            "Broken",
            vec![],
        )
        .unwrap_err();
        assert_eq!(failure.kind, RuntimeErrorKind::ArgumentCount);
        assert_eq!(failure.line, Some(2));
    }

    #[test]
    fn indexes_fixed_arrays_and_iterates_them_with_for_each() {
        let value = run(
            "Option Base 1\n\
             Public Function ArrayTotal() As Long\n\
               Dim values(3) As Long\n\
               Dim item As Variant\n\
               Dim total As Long\n\
               values(1) = 2\n\
               values(2) = 4\n\
               values(3) = 6\n\
               For Each item In values\n\
                 total = total + item\n\
               Next item\n\
               ArrayTotal = total + LBound(values) * 100 + UBound(values) * 10\n\
             End Function\n",
            "ArrayTotal",
            vec![],
        )
        .unwrap();
        assert_eq!(value, Value::Integer(142));
    }

    #[test]
    fn creates_array_values_and_preserves_elements_across_redim() {
        let value = run(
            "Public Function Resize() As Long\n\
               Dim values() As Long\n\
               ReDim values(1 To 2)\n\
               values(1) = 7\n\
               values(2) = 8\n\
               ReDim Preserve values(1 To 3)\n\
               values(3) = UBound(Array(10, 20, 30))\n\
               Resize = values(1) + values(2) + values(3)\n\
             End Function\n",
            "Resize",
            vec![],
        )
        .unwrap();
        assert_eq!(value, Value::Integer(17));
    }

    #[test]
    fn reports_out_of_range_array_access() {
        let failure = run(
            "Public Function Broken() As Long\n\
               Dim values(1 To 2) As Long\n\
               Broken = values(3)\n\
             End Function\n",
            "Broken",
            vec![],
        )
        .unwrap_err();
        assert_eq!(failure.kind, RuntimeErrorKind::SubscriptOutOfRange);
        assert_eq!(failure.line, Some(3));
    }

    #[test]
    fn redim_preserve_rejects_a_changed_lower_bound() {
        let failure = run(
            "Public Sub Broken()\n\
               Dim values(1 To 2) As Long\n\
               ReDim Preserve values(0 To 3)\n\
             End Sub\n",
            "Broken",
            vec![],
        )
        .unwrap_err();
        assert_eq!(failure.kind, RuntimeErrorKind::SubscriptOutOfRange);
        assert_eq!(failure.line, Some(3));
    }

    #[test]
    fn stops_an_infinite_loop_at_the_step_limit() {
        let module = parse_module(
            "Public Sub Forever()\n\
               Do\n\
               Loop\n\
             End Sub\n",
        )
        .unwrap();
        let failure = Runtime::new(&module)
            .with_limits(5, 4)
            .call("Forever", vec![])
            .unwrap_err();
        assert_eq!(failure.kind, RuntimeErrorKind::StepLimit);
        assert_eq!(failure.line, Some(2));
    }

    #[test]
    fn unsupported_office_members_fail_explicitly() {
        let failure = run(
            "Public Function ReadCell() As Variant\n\
               ReadCell = Range(\"A1\").Value\n\
             End Function\n",
            "ReadCell",
            vec![],
        )
        .unwrap_err();
        assert_eq!(failure.kind, RuntimeErrorKind::Unsupported);
        assert_eq!(failure.line, Some(2));
    }

    #[test]
    fn reports_division_by_zero_with_a_source_line() {
        let failure = run(
            "Public Function Broken() As Double\n\
               Broken = 1 / 0\n\
             End Function\n",
            "Broken",
            vec![],
        )
        .unwrap_err();
        assert_eq!(failure.kind, RuntimeErrorKind::DivisionByZero);
        assert_eq!(failure.line, Some(2));
    }

    #[test]
    fn enforces_the_call_depth_limit() {
        let module = parse_module(
            "Public Function Recurse() As Long\n\
               Recurse = Recurse()\n\
             End Function\n",
        )
        .unwrap();
        let failure = Runtime::new(&module)
            .with_limits(100, 4)
            .call("Recurse", vec![])
            .unwrap_err();
        assert_eq!(failure.kind, RuntimeErrorKind::CallDepth);
        assert_eq!(failure.line, Some(2));
    }

    #[test]
    fn rejects_the_wrong_argument_count() {
        let failure = run(
            "Public Function AddOne(value As Long) As Long\n\
               AddOne = value + 1\n\
             End Function\n",
            "AddOne",
            vec![],
        )
        .unwrap_err();
        assert_eq!(failure.kind, RuntimeErrorKind::ArgumentCount);
        assert_eq!(failure.line, None);
    }
}
