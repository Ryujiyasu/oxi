// This Source Code Form is subject to the terms of the Mozilla Public
// License, v. 2.0. If a copy of the MPL was not distributed with this
// file, You can obtain one at https://mozilla.org/MPL/2.0/.

//! Host-independent execution of pure VBA procedures.
//!
//! The browser runtime supports scalar and multidimensional array values,
//! procedure/module/static storage, structured and label-based control flow,
//! error handlers, and calls between VBA procedures. Office objects require a
//! host adapter. File I/O and events fail explicitly rather than being
//! approximated.

use std::{
    cell::RefCell,
    collections::{BTreeMap, BTreeSet},
    rc::Rc,
};

use crate::ast::{
    AlignedAssignStmt, AlignmentKind, Argument, ArrayBound, BinaryOp, CaseLabel, DoStmt, ExitKind,
    Expr, ForEachStmt, ForStmt, Literal, LoopTest, MidAssignStmt, Module, ModuleItem, ModuleOption,
    OnBranchKind, OnError, ParamMode, ProcKind, Procedure, ResumeTarget, SelectCaseStmt, Statement,
    TypeName, UnaryOp, VarDecl, VarItem,
};

#[derive(Debug, Clone, PartialEq)]
pub enum Value {
    Empty,
    /// An omitted `Optional ... As Variant` argument. Unlike `Empty`, this is
    /// observable through VBA's `IsMissing` function.
    Missing,
    /// The null object reference used by `Set value = Nothing`.
    Nothing,
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
    pub dimensions: Vec<ArrayDimension>,
    pub values: Vec<Value>,
    pub element_default: Box<Value>,
    /// `true` for dynamic arrays and arrays stored in a Variant. Fixed-size
    /// declaration arrays cannot be deallocated or resized.
    pub resizable: bool,
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct ArrayDimension {
    pub lower_bound: i64,
    pub length: usize,
}

impl ArrayValue {
    pub fn lower_bound(&self, dimension: usize) -> Option<i64> {
        self.dimensions
            .get(dimension.checked_sub(1)?)
            .map(|dimension| dimension.lower_bound)
    }

    pub fn upper_bound(&self, dimension: usize) -> Option<i64> {
        self.dimensions
            .get(dimension.checked_sub(1)?)
            .map(|dimension| match dimension.length.checked_sub(1) {
                Some(offset) => dimension.lower_bound.saturating_add(offset as i64),
                None => dimension.lower_bound.saturating_sub(1),
            })
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum RuntimeErrorKind {
    ProcedureNotFound,
    ArgumentCount,
    UndefinedVariable,
    TypeMismatch,
    ObjectVariableNotSet,
    Overflow,
    SubscriptOutOfRange,
    Host,
    UserDefined,
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
    /// Exact number supplied to `Err.Raise`, when this originated as a VBA
    /// user error rather than one of the runtime's built-in failures.
    pub vba_number: Option<i64>,
    pub vba_source: Option<String>,
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
    module_values: BTreeMap<String, ValueSlot>,
    module_constants: BTreeSet<String>,
    module_auto_new: BTreeMap<String, String>,
    module_fixed_strings: BTreeMap<String, usize>,
    module_variants: BTreeSet<String>,
    static_values: BTreeMap<(String, String), ValueSlot>,
    module_initialized: bool,
    internal_objects: BTreeMap<u64, InternalObject>,
    next_internal_handle: u64,
}

struct Frame {
    procedure_name: String,
    source_name: String,
    values: BTreeMap<String, ValueSlot>,
    constants: BTreeSet<String>,
    auto_new: BTreeMap<String, String>,
    fixed_strings: BTreeMap<String, usize>,
    variants: BTreeSet<String>,
    static_procedure: bool,
    with_objects: Vec<ObjectRef>,
    error_mode: ErrorMode,
    error_state: ErrorState,
    error_handler_active: bool,
    error_statement: Option<usize>,
    current_statement: usize,
    gosub_returns: Vec<usize>,
}

#[derive(Clone)]
enum ErrorMode {
    Disabled,
    ResumeNext,
    Goto(String),
}

#[derive(Clone, Default)]
struct ErrorState {
    number: i64,
    description: String,
    source: String,
    line: Option<u32>,
}

enum InternalObject {
    Collection(Vec<CollectionEntry>),
}

struct CollectionEntry {
    value: Value,
    key: Option<String>,
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
    Jump(String),
    GoSub(String),
    Return,
    Resume(ResumeTarget),
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
            module_values: BTreeMap::new(),
            module_constants: BTreeSet::new(),
            module_auto_new: BTreeMap::new(),
            module_fixed_strings: BTreeMap::new(),
            module_variants: BTreeSet::new(),
            static_values: BTreeMap::new(),
            module_initialized: false,
            internal_objects: BTreeMap::new(),
            next_internal_handle: 1_u64 << 63,
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
        self.initialize_module()?;
        self.call_procedure(name, args, None)
    }

    fn initialize_module(&mut self) -> Result<(), RuntimeError> {
        if self.module_initialized {
            return Ok(());
        }
        self.module_initialized = true;
        let result = self.initialize_module_items();
        if result.is_err() {
            self.module_values.clear();
            self.module_constants.clear();
            self.module_auto_new.clear();
            self.module_fixed_strings.clear();
            self.module_variants.clear();
            self.internal_objects.clear();
            self.next_internal_handle = 1_u64 << 63;
            self.module_initialized = false;
        }
        result
    }

    fn initialize_module_items(&mut self) -> Result<(), RuntimeError> {
        let items = self.module.items.clone();
        let mut frame = empty_frame();
        for item in items {
            match item {
                ModuleItem::Variables(declaration) => {
                    for variable in &declaration.items {
                        let name = key(&variable.name);
                        if variable.type_name.is_new {
                            self.module_auto_new
                                .insert(name.clone(), variable.type_name.name.clone());
                        }
                        let value =
                            self.declared_value(variable, &mut frame, declaration.span.line)?;
                        if variable.array_bounds.is_none()
                            && variable.type_name.fixed_length.is_some()
                        {
                            self.module_fixed_strings
                                .insert(name.clone(), string_width(&value));
                        }
                        if variable.array_bounds.is_none()
                            && variable.type_name.name.eq_ignore_ascii_case("variant")
                        {
                            self.module_variants.insert(name.clone());
                        }
                        self.module_values
                            .insert(name.clone(), Rc::new(RefCell::new(value)));
                        if declaration.is_const {
                            self.module_constants.insert(name);
                        }
                    }
                }
                ModuleItem::Enum(definition) => {
                    let mut next = 0_i64;
                    for (name, expression) in definition.members {
                        let value = match expression {
                            Some(expression) => {
                                let value = self.eval_expr(&expression, &mut frame)?;
                                let value = number(&value).map_err(|message| {
                                    error(
                                        RuntimeErrorKind::TypeMismatch,
                                        message,
                                        Some(definition.span.line),
                                    )
                                })?;
                                if !value.is_finite()
                                    || value < i64::MIN as f64
                                    || value > i64::MAX as f64
                                {
                                    return Err(error(
                                        RuntimeErrorKind::Overflow,
                                        "Enum value is outside the supported integer range",
                                        Some(definition.span.line),
                                    ));
                                }
                                value.round_ties_even() as i64
                            }
                            None => next,
                        };
                        let name = key(&name);
                        self.module_values
                            .insert(name.clone(), Rc::new(RefCell::new(Value::Integer(value))));
                        self.module_constants.insert(name);
                        next = value.checked_add(1).ok_or_else(|| {
                            error(
                                RuntimeErrorKind::Overflow,
                                "Enum value overflow",
                                Some(definition.span.line),
                            )
                        })?;
                    }
                }
                _ => {}
            }
        }
        Ok(())
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
                dimensions: vec![ArrayDimension {
                    lower_bound: 0,
                    length: received.saturating_sub(fixed_count),
                }],
                values: values.collect(),
                element_default: Box::new(default_value(&param_array.type_name)),
                resizable: false,
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
            source_name: procedure.name.clone(),
            values: BTreeMap::new(),
            constants: BTreeSet::new(),
            auto_new: BTreeMap::new(),
            fixed_strings: BTreeMap::new(),
            variants: BTreeSet::new(),
            static_procedure: procedure.is_static,
            with_objects: Vec::new(),
            error_mode: ErrorMode::Disabled,
            error_state: ErrorState::default(),
            error_handler_active: false,
            error_statement: None,
            current_statement: 0,
            gosub_returns: Vec::new(),
        };
        for (param, argument) in procedure.params.iter().zip(args) {
            let value = match argument {
                BoundArgument::Value(value) => Rc::new(RefCell::new(value)),
                BoundArgument::Reference(value) => value,
            };
            frame.values.insert(key(&param.name), value);
            if !param.is_array
                && param.mode != ParamMode::ParamArray
                && param.type_name.name.eq_ignore_ascii_case("variant")
            {
                frame.variants.insert(key(&param.name));
            }
        }
        if !matches!(procedure.kind, ProcKind::Sub) {
            frame.values.insert(
                frame.procedure_name.clone(),
                Rc::new(RefCell::new(default_return_value(procedure))),
            );
            if procedure
                .return_type
                .as_ref()
                .is_none_or(|type_name| type_name.name.eq_ignore_ascii_case("variant"))
            {
                frame.variants.insert(frame.procedure_name.clone());
            }
        }
        self.declare_procedure_locals(&procedure.body, &mut frame)?;

        self.depth += 1;
        let flow = self.exec_procedure_body(&procedure.body, &mut frame);
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
            Flow::Jump(label) | Flow::GoSub(label) => {
                return Err(error(
                    RuntimeErrorKind::UndefinedVariable,
                    format!("VBA label not found: {label}"),
                    Some(procedure.span.line),
                ))
            }
            Flow::Return => {
                return Err(error(
                    RuntimeErrorKind::Unsupported,
                    "Return without GoSub",
                    Some(procedure.span.line),
                ))
            }
            Flow::Resume(_) => {
                return Err(error(
                    RuntimeErrorKind::Unsupported,
                    "Resume without an active error handler",
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

    fn exec_procedure_body(
        &mut self,
        body: &[Statement],
        frame: &mut Frame,
    ) -> Result<Flow, RuntimeError> {
        let mut labels = BTreeMap::new();
        for (index, statement) in body.iter().enumerate() {
            match statement {
                Statement::Label { name, .. } => {
                    labels.entry(key(name)).or_insert(index);
                }
                Statement::LineNumber { value, .. } => {
                    labels.entry(value.to_string()).or_insert(index);
                }
                _ => {}
            }
        }

        let mut pc = 0;
        while pc < body.len() {
            frame.current_statement = pc;
            match self.exec_body(&body[pc..=pc], frame)? {
                Flow::Continue => pc += 1,
                Flow::Jump(label) => {
                    pc = label_destination(&labels, &label, line_of(&body[pc]))?;
                }
                Flow::GoSub(label) => {
                    frame.gosub_returns.push(pc + 1);
                    pc = label_destination(&labels, &label, line_of(&body[pc]))?;
                }
                Flow::Return => {
                    pc = frame.gosub_returns.pop().ok_or_else(|| {
                        error(
                            RuntimeErrorKind::Unsupported,
                            "Return without GoSub",
                            line_of(&body[pc]),
                        )
                    })?;
                }
                Flow::Resume(target) => {
                    if !frame.error_handler_active {
                        return Err(error(
                            RuntimeErrorKind::Unsupported,
                            "Resume without an active error handler",
                            line_of(&body[pc]),
                        ));
                    }
                    let failed = frame.error_statement.ok_or_else(|| {
                        error(
                            RuntimeErrorKind::Unsupported,
                            "Resume has no failed statement",
                            line_of(&body[pc]),
                        )
                    })?;
                    frame.error_handler_active = false;
                    frame.error_statement = None;
                    frame.error_state = ErrorState::default();
                    pc = match target {
                        ResumeTarget::Same => failed,
                        ResumeTarget::Next => failed.saturating_add(1),
                        ResumeTarget::Label(label) => {
                            label_destination(&labels, &label, line_of(&body[pc]))?
                        }
                    };
                }
                flow => return Ok(flow),
            }
        }
        Ok(Flow::Continue)
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
                source_name: String::new(),
                values: BTreeMap::new(),
                constants: BTreeSet::new(),
                auto_new: BTreeMap::new(),
                fixed_strings: BTreeMap::new(),
                variants: BTreeSet::new(),
                static_procedure: false,
                with_objects: Vec::new(),
                error_mode: ErrorMode::Disabled,
                error_state: ErrorState::default(),
                error_handler_active: false,
                error_statement: None,
                current_statement: 0,
                gosub_returns: Vec::new(),
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
            let flow = match self.exec_statement(statement, frame) {
                Ok(flow) => flow,
                Err(failure) => self.handle_runtime_error(failure, frame)?,
            };
            if !matches!(flow, Flow::Continue) {
                return Ok(flow);
            }
        }
        Ok(Flow::Continue)
    }

    fn handle_runtime_error(
        &mut self,
        failure: RuntimeError,
        frame: &mut Frame,
    ) -> Result<Flow, RuntimeError> {
        if matches!(
            failure.kind,
            RuntimeErrorKind::StepLimit | RuntimeErrorKind::CallDepth
        ) || frame.error_handler_active
        {
            return Err(failure);
        }
        frame.error_state = ErrorState {
            number: runtime_error_number(&failure),
            description: failure.message.clone(),
            source: failure
                .vba_source
                .clone()
                .unwrap_or_else(|| frame.source_name.clone()),
            line: failure.line,
        };
        match frame.error_mode.clone() {
            ErrorMode::Disabled => Err(failure),
            ErrorMode::ResumeNext => Ok(Flow::Continue),
            ErrorMode::Goto(label) => {
                frame.error_handler_active = true;
                frame.error_statement = Some(frame.current_statement);
                Ok(Flow::Jump(label))
            }
        }
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
                let value = self.let_value(value, span.line)?;
                self.assign(target, value, frame, span.line)?;
                Ok(Flow::Continue)
            }
            Statement::SetAssign {
                target,
                value,
                span,
            } => {
                let value = match value {
                    Expr::EvaluateShortcut { text, .. } => {
                        self.evaluate_shortcut(text, span.line, true)?
                    }
                    _ => self.eval_expr(value, frame)?,
                };
                if !matches!(value, Value::Object(_) | Value::Nothing) {
                    return Err(error(
                        RuntimeErrorKind::TypeMismatch,
                        "Set requires an object value or Nothing",
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
            Statement::Erase { targets, span } => {
                for target in targets {
                    self.erase_array(target, frame, span.line)?;
                }
                Ok(Flow::Continue)
            }
            Statement::MidAssign(statement) => {
                self.exec_mid_assign(statement, frame)?;
                Ok(Flow::Continue)
            }
            Statement::AlignedAssign(statement) => {
                self.exec_aligned_assign(statement, frame)?;
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
            Statement::OnError(mode) => {
                frame.error_mode = match mode {
                    OnError::Goto { label, .. } => ErrorMode::Goto(label.clone()),
                    OnError::Disable { .. } => ErrorMode::Disabled,
                    OnError::ResumeNext { .. } => ErrorMode::ResumeNext,
                };
                if matches!(mode, OnError::Disable { .. }) {
                    frame.error_handler_active = false;
                    frame.error_statement = None;
                }
                Ok(Flow::Continue)
            }
            Statement::OnBranch(branch) => {
                let selector = self.array_index(&branch.selector, frame, branch.span.line)?;
                if selector < 1 || selector as usize > branch.labels.len() {
                    return Ok(Flow::Continue);
                }
                let label = branch.labels[selector as usize - 1].clone();
                Ok(match branch.kind {
                    OnBranchKind::GoTo => Flow::Jump(label),
                    OnBranchKind::GoSub => Flow::GoSub(label),
                })
            }
            Statement::Resume { target, .. } => Ok(Flow::Resume(target.clone())),
            Statement::GoTo { label, .. } => Ok(Flow::Jump(label.clone())),
            Statement::GoSub { label, .. } => Ok(Flow::GoSub(label.clone())),
            Statement::Return { .. } => Ok(Flow::Return),
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

    fn erase_array(
        &mut self,
        target: &Expr,
        frame: &mut Frame,
        line: u32,
    ) -> Result<(), RuntimeError> {
        let value = self.eval_expr(target, frame)?;
        let Value::Array(mut array) = value else {
            return Err(error(
                RuntimeErrorKind::TypeMismatch,
                "Erase target must contain an array",
                Some(line),
            ));
        };
        if array.resizable {
            array.dimensions.clear();
            array.values.clear();
        } else {
            array.values.fill(*array.element_default.clone());
        }
        self.assign(target, Value::Array(array), frame, line)
    }

    fn exec_mid_assign(
        &mut self,
        statement: &MidAssignStmt,
        frame: &mut Frame,
    ) -> Result<(), RuntimeError> {
        let current = match self.eval_expr(&statement.target, frame)? {
            Value::String(value) => value,
            _ => {
                return Err(error(
                    RuntimeErrorKind::TypeMismatch,
                    "Mid statement target must be a String variable",
                    Some(statement.span.line),
                ))
            }
        };
        let start_value = self.eval_expr(&statement.start, frame)?;
        let start = positive_position(&start_value, Some(statement.span.line))?;
        let requested = match &statement.length {
            Some(length) => {
                let value = self.eval_expr(length, frame)?;
                Some(nonnegative_length(&value, Some(statement.span.line))?)
            }
            None => None,
        };
        let replacement = self.eval_expr(&statement.value, frame)?;
        let replacement = match replacement {
            Value::Null => {
                return Err(error(
                    RuntimeErrorKind::TypeMismatch,
                    "invalid use of Null",
                    Some(statement.span.line),
                ))
            }
            value => text(&value).map_err(|message| {
                error(
                    RuntimeErrorKind::TypeMismatch,
                    message,
                    Some(statement.span.line),
                )
            })?,
        };
        let mut current = current.encode_utf16().collect::<Vec<_>>();
        if start >= current.len() {
            return Err(invalid_procedure_call(
                "Mid statement start exceeds the target String length".to_string(),
                Some(statement.span.line),
            ));
        }
        let replacement = replacement.encode_utf16().collect::<Vec<_>>();
        let requested = requested.unwrap_or(replacement.len());
        let count = requested.min(replacement.len()).min(current.len() - start);
        current[start..start + count].copy_from_slice(&replacement[..count]);
        self.assign(
            &statement.target,
            Value::String(String::from_utf16_lossy(&current)),
            frame,
            statement.span.line,
        )
    }

    fn exec_aligned_assign(
        &mut self,
        statement: &AlignedAssignStmt,
        frame: &mut Frame,
    ) -> Result<(), RuntimeError> {
        let current = match self.eval_expr(&statement.target, frame)? {
            Value::String(value) => value,
            _ => {
                return Err(error(
                    RuntimeErrorKind::TypeMismatch,
                    "LSet and RSet targets must be String variables",
                    Some(statement.span.line),
                ))
            }
        };
        let width = expr_name(&statement.target)
            .and_then(|name| self.fixed_string_width(frame, name))
            .unwrap_or_else(|| current.encode_utf16().count());
        let value = self.eval_expr(&statement.value, frame)?;
        let value = match value {
            Value::Null => {
                return Err(error(
                    RuntimeErrorKind::TypeMismatch,
                    "invalid use of Null",
                    Some(statement.span.line),
                ))
            }
            value => text(&value).map_err(|message| {
                error(
                    RuntimeErrorKind::TypeMismatch,
                    message,
                    Some(statement.span.line),
                )
            })?,
        };
        let mut value = value.encode_utf16().take(width).collect::<Vec<_>>();
        let padding = width.saturating_sub(value.len());
        match statement.kind {
            AlignmentKind::Left => value.extend(std::iter::repeat_n(u16::from(b' '), padding)),
            AlignmentKind::Right => {
                let mut aligned = vec![u16::from(b' '); padding];
                aligned.extend(value);
                value = aligned;
            }
        }
        self.assign(
            &statement.target,
            Value::String(String::from_utf16_lossy(&value)),
            frame,
            statement.span.line,
        )
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
        let option_compare_text = self.option_compare_text();
        let compare = |op, lhs, rhs| {
            binary(op, lhs, rhs, option_compare_text)
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
            let name = key(&variable.name);
            if frame.values.contains_key(&name) {
                continue;
            }
            let is_static = decl.is_static || frame.static_procedure;
            let slot = if is_static {
                let static_key = (frame.procedure_name.clone(), name.clone());
                if let Some(value) = self.static_values.get(&static_key) {
                    value.clone()
                } else {
                    let value = self.declared_value(variable, frame, decl.span.line)?;
                    let value = Rc::new(RefCell::new(value));
                    self.static_values.insert(static_key, value.clone());
                    value
                }
            } else {
                Rc::new(RefCell::new(self.declared_value(
                    variable,
                    frame,
                    decl.span.line,
                )?))
            };
            if variable.array_bounds.is_none() && variable.type_name.fixed_length.is_some() {
                frame
                    .fixed_strings
                    .insert(name.clone(), string_width(&slot.borrow()));
            }
            if variable.array_bounds.is_none()
                && variable.type_name.name.eq_ignore_ascii_case("variant")
            {
                frame.variants.insert(name.clone());
            }
            frame.values.insert(name.clone(), slot);
            if variable.type_name.is_new {
                frame
                    .auto_new
                    .insert(name.clone(), variable.type_name.name.clone());
            }
            if decl.is_const {
                frame.constants.insert(name);
            }
        }
        Ok(())
    }

    fn declare_procedure_locals(
        &mut self,
        body: &[Statement],
        frame: &mut Frame,
    ) -> Result<(), RuntimeError> {
        for statement in body {
            match statement {
                Statement::Dim(declaration) => self.declare_locals(declaration, frame)?,
                Statement::If(branch) => {
                    self.declare_procedure_locals(&branch.then_body, frame)?;
                    for (_, body) in &branch.else_ifs {
                        self.declare_procedure_locals(body, frame)?;
                    }
                    if let Some(body) = &branch.else_body {
                        self.declare_procedure_locals(body, frame)?;
                    }
                }
                Statement::SelectCase(select) => {
                    for case in &select.cases {
                        self.declare_procedure_locals(&case.body, frame)?;
                    }
                    if let Some(body) = &select.case_else {
                        self.declare_procedure_locals(body, frame)?;
                    }
                }
                Statement::For(loop_) => {
                    self.declare_procedure_locals(&loop_.body, frame)?;
                }
                Statement::ForEach(loop_) => {
                    self.declare_procedure_locals(&loop_.body, frame)?;
                }
                Statement::Do(loop_) => {
                    self.declare_procedure_locals(&loop_.body, frame)?;
                }
                Statement::While { body, .. } | Statement::With { body, .. } => {
                    self.declare_procedure_locals(body, frame)?;
                }
                _ => {}
            }
        }
        Ok(())
    }

    fn declared_value(
        &mut self,
        variable: &VarItem,
        frame: &mut Frame,
        line: u32,
    ) -> Result<Value, RuntimeError> {
        if variable.type_name.is_new {
            return self.new_object(&variable.type_name.name, line);
        }
        match &variable.array_bounds {
            Some(bounds) => {
                self.make_array(bounds, &variable.type_name, frame, line, bounds.is_empty())
            }
            None => match &variable.value {
                Some(expr) => {
                    let value = self.eval_expr(expr, frame)?;
                    if variable.type_name.fixed_length.is_some() {
                        let width = string_width(&self.declared_default_value(
                            &variable.type_name,
                            frame,
                            line,
                        )?);
                        coerce_string_width(value, width, Some(line))
                    } else {
                        Ok(value)
                    }
                }
                None => self.declared_default_value(&variable.type_name, frame, line),
            },
        }
    }

    fn declared_default_value(
        &mut self,
        type_name: &TypeName,
        frame: &mut Frame,
        line: u32,
    ) -> Result<Value, RuntimeError> {
        let Some(length) = &type_name.fixed_length else {
            return Ok(default_value(type_name));
        };
        let length = self.array_index(length, frame, line)?;
        let length = usize::try_from(length)
            .ok()
            .filter(|length| *length <= 1_000_000)
            .ok_or_else(|| {
                error(
                    RuntimeErrorKind::Overflow,
                    "fixed-length VBA String is outside the browser runtime limit",
                    Some(line),
                )
            })?;
        Ok(Value::String(" ".repeat(length)))
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
        if self.is_constant(frame, &item.name) {
            return Err(constant_assignment_error(&item.name, line));
        }
        let existing_slot = self.lookup_slot(frame, &item.name);
        if let Some(existing) = &existing_slot {
            if matches!(&*existing.borrow(), Value::Array(array) if !array.resizable) {
                return Err(fixed_array_error(&item.name, Some(line)));
            }
        }
        let mut replacement = match self.make_array(bounds, &item.type_name, frame, line, true)? {
            Value::Array(array) => array,
            _ => unreachable!(),
        };
        if item.type_name.name.eq_ignore_ascii_case("variant") {
            if let Some(existing) = &existing_slot {
                if let Value::Array(existing) = &*existing.borrow() {
                    replacement.element_default = existing.element_default.clone();
                    replacement.values.fill(*existing.element_default.clone());
                }
            }
        }
        if preserve {
            let existing = existing_slot.clone().ok_or_else(|| {
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
            if !existing.dimensions.is_empty()
                && !preservable_dimensions(&existing.dimensions, &replacement.dimensions)
            {
                return Err(error(
                    RuntimeErrorKind::SubscriptOutOfRange,
                    "ReDim Preserve can only resize an array's last dimension",
                    Some(line),
                ));
            }
            preserve_array_values(existing, &mut replacement);
        }
        let replacement = Value::Array(replacement);
        match existing_slot {
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
        resizable: bool,
    ) -> Result<Value, RuntimeError> {
        if bounds.is_empty() {
            return Ok(Value::Array(ArrayValue {
                dimensions: Vec::new(),
                values: Vec::new(),
                element_default: Box::new(self.declared_default_value(
                    element_type,
                    frame,
                    line,
                )?),
                resizable,
            }));
        }
        let mut dimensions = Vec::with_capacity(bounds.len());
        let mut total_len = 1usize;
        for bound in bounds {
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
            let length = upper
                .checked_sub(lower)
                .and_then(|value| value.checked_add(1))
                .and_then(|value| usize::try_from(value).ok())
                .ok_or_else(|| {
                    error(
                        RuntimeErrorKind::Overflow,
                        "VBA array dimension is too large for the browser runtime",
                        Some(line),
                    )
                })?;
            total_len = total_len
                .checked_mul(length)
                .filter(|len| *len <= 1_000_000)
                .ok_or_else(|| {
                    error(
                        RuntimeErrorKind::Overflow,
                        "VBA array is too large for the browser runtime",
                        Some(line),
                    )
                })?;
            dimensions.push(ArrayDimension {
                lower_bound: lower,
                length,
            });
        }
        let default = self.declared_default_value(element_type, frame, line)?;
        Ok(Value::Array(ArrayValue {
            dimensions,
            values: vec![default.clone(); total_len],
            element_default: Box::new(default),
            resizable,
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

    fn option_compare_text(&self) -> bool {
        self.module.items.iter().any(|item| {
            matches!(
                item,
                ModuleItem::Option(ModuleOption::Compare(mode), _)
                    if mode.eq_ignore_ascii_case("text")
            )
        })
    }

    fn lookup_slot(&self, frame: &Frame, name: &str) -> Option<ValueSlot> {
        let name = key(name);
        frame
            .values
            .get(&name)
            .or_else(|| self.module_values.get(&name))
            .cloned()
    }

    fn read_variable(
        &mut self,
        frame: &Frame,
        name: &str,
        line: u32,
    ) -> Result<Option<Value>, RuntimeError> {
        let name_key = key(name);
        let Some(slot) = self.lookup_slot(frame, name) else {
            return Ok(None);
        };
        if matches!(&*slot.borrow(), Value::Nothing) {
            let type_name = if frame.values.contains_key(&name_key) {
                frame.auto_new.get(&name_key)
            } else {
                self.module_auto_new.get(&name_key)
            }
            .cloned();
            if let Some(type_name) = type_name {
                let value = self.new_object(&type_name, line)?;
                *slot.borrow_mut() = value;
            }
        }
        let value = slot.borrow().clone();
        Ok(Some(value))
    }

    fn new_object(&mut self, type_name: &str, line: u32) -> Result<Value, RuntimeError> {
        if !type_name.eq_ignore_ascii_case("collection") {
            return Err(error(
                RuntimeErrorKind::Unsupported,
                format!("New is not available for VBA type: {type_name}"),
                Some(line),
            ));
        }
        let handle = self.next_internal_handle;
        self.next_internal_handle = self.next_internal_handle.checked_add(1).ok_or_else(|| {
            error(
                RuntimeErrorKind::Overflow,
                "VBA internal object handle overflow",
                Some(line),
            )
        })?;
        self.internal_objects
            .insert(handle, InternalObject::Collection(Vec::new()));
        Ok(Value::Object(ObjectRef {
            handle,
            kind: "Collection".to_string(),
        }))
    }

    fn is_constant(&self, frame: &Frame, name: &str) -> bool {
        let name = key(name);
        if frame.values.contains_key(&name) {
            frame.constants.contains(&name)
        } else {
            self.module_constants.contains(&name)
        }
    }

    fn fixed_string_width(&self, frame: &Frame, name: &str) -> Option<usize> {
        let name = key(name);
        if frame.values.contains_key(&name) {
            frame.fixed_strings.get(&name).copied()
        } else {
            self.module_fixed_strings.get(&name).copied()
        }
    }

    fn is_variant_variable(&self, frame: &Frame, name: &str) -> bool {
        let name = key(name);
        if frame.values.contains_key(&name) {
            frame.variants.contains(&name)
        } else {
            self.module_variants.contains(&name)
        }
    }

    fn assign(
        &mut self,
        target: &Expr,
        value: Value,
        frame: &mut Frame,
        line: u32,
    ) -> Result<(), RuntimeError> {
        match target {
            Expr::EvaluateShortcut { text, .. } => {
                let receiver = match self.evaluate_shortcut(text, line, true)? {
                    Value::Object(receiver) => receiver,
                    _ => {
                        return Err(error(
                            RuntimeErrorKind::TypeMismatch,
                            "Evaluate assignment target is not an object",
                            Some(line),
                        ))
                    }
                };
                match self.host_set(&receiver, "Value", value, line)? {
                    true => Ok(()),
                    false => Err(error(
                        RuntimeErrorKind::Unsupported,
                        "evaluated object has no writable default Value property",
                        Some(line),
                    )),
                }
            }
            Expr::Ident(name, _) | Expr::TypedIdent { name, .. } => {
                let implicit_variant = self.lookup_slot(frame, name).is_none();
                let mut value = match self.fixed_string_width(frame, name) {
                    Some(width) => coerce_string_width(value, width, Some(line))?,
                    None => value,
                };
                if implicit_variant || self.is_variant_variable(frame, name) {
                    if let Value::Array(array) = &mut value {
                        array.resizable = true;
                    }
                }
                let name = key(name);
                if let Some(existing) = frame.values.get(&name) {
                    if frame.constants.contains(&name) {
                        return Err(constant_assignment_error(&name, line));
                    }
                    *existing.borrow_mut() = value;
                } else if let Some(existing) = self.module_values.get(&name) {
                    if self.module_constants.contains(&name) {
                        return Err(constant_assignment_error(&name, line));
                    }
                    *existing.borrow_mut() = value;
                } else {
                    if implicit_variant {
                        frame.variants.insert(name.clone());
                    }
                    frame.values.insert(name, Rc::new(RefCell::new(value)));
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
                let indices = self.array_arguments(args, frame, line)?;
                let name_key = key(name);
                if (frame.values.contains_key(&name_key) && frame.constants.contains(&name_key))
                    || (!frame.values.contains_key(&name_key)
                        && self.module_constants.contains(&name_key))
                {
                    return Err(constant_assignment_error(name, line));
                }
                let array = self.lookup_slot(frame, name).ok_or_else(|| {
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
                let value = match array.element_default.as_ref() {
                    Value::String(default) if !default.is_empty() => {
                        coerce_string_width(value, default.encode_utf16().count(), Some(line))?
                    }
                    _ => value,
                };
                let offset = array_offset(array, &indices, line)?;
                array.values[offset] = value;
                Ok(())
            }
            Expr::Member {
                object, name, span, ..
            } => {
                let receiver = self.eval_object(object, frame, span.line)?;
                if is_err_object(&receiver) {
                    return err_set(frame, name, value, span.line);
                }
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
            Expr::EvaluateShortcut { text, span } => self.evaluate_shortcut(text, span.line, false),
            Expr::Ident(name, span) | Expr::TypedIdent { name, span, .. } => {
                if let Some(value) = self.read_variable(frame, name, span.line)? {
                    return Ok(value);
                }
                if let Some(value) = builtin_constant(name) {
                    return Ok(value);
                }
                if name.eq_ignore_ascii_case("err") {
                    return Ok(Value::Object(ObjectRef {
                        handle: u64::MAX,
                        kind: "Err".to_string(),
                    }));
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
            Expr::New { type_name, span } => self.new_object(type_name, span.line),
            Expr::Unary { op, operand, span } => {
                let value = self.eval_expr(operand, frame)?;
                unary(*op, value).map_err(|message| {
                    error(RuntimeErrorKind::TypeMismatch, message, Some(span.line))
                })
            }
            Expr::Binary { op, lhs, rhs, span } => {
                let lhs = self.eval_expr(lhs, frame)?;
                let rhs = self.eval_expr(rhs, frame)?;
                binary(*op, lhs, rhs, self.option_compare_text())
                    .map_err(|(kind, message)| error(kind, message, Some(span.line)))
            }
            Expr::TypeOf {
                operand,
                type_name,
                span,
            } => match self.eval_expr(operand, frame)? {
                Value::Object(object) => Ok(Value::Boolean(
                    object.kind.eq_ignore_ascii_case(type_name)
                        || type_name.eq_ignore_ascii_case("object"),
                )),
                Value::Nothing => Ok(Value::Boolean(false)),
                _ => Err(error(
                    RuntimeErrorKind::TypeMismatch,
                    "TypeOf requires an object expression",
                    Some(span.line),
                )),
            },
            Expr::Index {
                target, args, span, ..
            } if expr_name(target).is_some_and(|name| {
                self.lookup_slot(frame, name)
                    .as_ref()
                    .is_some_and(|value| matches!(&*value.borrow(), Value::Array(_)))
            }) =>
            {
                let name = expr_name(target).unwrap();
                let indices = self.array_arguments(args, frame, span.line)?;
                let value = self.lookup_slot(frame, name).unwrap();
                let value = value.borrow();
                let Value::Array(array) = &*value else {
                    unreachable!()
                };
                Ok(array.values[array_offset(array, &indices, span.line)?].clone())
            }
            Expr::Index { .. } => self.eval_call(expr, frame),
            Expr::Member {
                object, name, span, ..
            } => {
                let receiver = self.eval_object(object, frame, span.line)?;
                if is_err_object(&receiver) {
                    return err_property(frame, name, span.line);
                }
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
                    if self.lookup_slot(frame, name).is_some() {
                        match self.read_variable(frame, name, span.line)? {
                            Some(Value::Object(receiver)) => {
                                let mut values = Vec::with_capacity(args.len());
                                for argument in args {
                                    values.push(match argument.value.as_ref() {
                                        Some(value) => self.eval_expr(value, frame)?,
                                        None => Value::Missing,
                                    });
                                }
                                return self
                                    .host_call(Some(&receiver), "Item", &values, span.line)?
                                    .ok_or_else(|| {
                                        error(
                                            RuntimeErrorKind::Unsupported,
                                            format!(
                                                "default member is not available: {}.Item",
                                                receiver.kind
                                            ),
                                            Some(span.line),
                                        )
                                    });
                            }
                            Some(Value::Nothing) => {
                                return Err(error(
                                    RuntimeErrorKind::ObjectVariableNotSet,
                                    "object variable or With block variable not set",
                                    Some(span.line),
                                ));
                            }
                            _ => {}
                        }
                    }
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
                    values.push(match argument.value.as_ref() {
                        Some(value) => self.eval_expr(value, frame)?,
                        None => Value::Missing,
                    });
                }
                match target.as_ref() {
                    Expr::Ident(name, _) | Expr::TypedIdent { name, .. } => {
                        self.call_named(name, values, Some(span.line))
                    }
                    Expr::Member { object, name, .. } => {
                        let receiver = self.eval_object(object, frame, span.line)?;
                        if is_err_object(&receiver) {
                            return err_call(frame, name, &values, span.line);
                        }
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
                if is_err_object(&receiver) {
                    return err_call(frame, name, &[], span.line);
                }
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
        let mut copybacks = Vec::<(ValueSlot, Vec<i64>, ValueSlot)>::new();
        let mut fixed_string_copybacks = Vec::<(ValueSlot, usize, ValueSlot)>::new();
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
                    dimensions: vec![ArrayDimension {
                        lower_bound: 0,
                        length: values.len(),
                    }],
                    values,
                    element_default: Box::new(default_value(&parameter.type_name)),
                    resizable: false,
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
                    if !self.is_constant(frame, name) {
                        self.read_variable(frame, name, expression.span().line)?;
                        if let Some(value) = self.lookup_slot(frame, name) {
                            if let Some(width) = self.fixed_string_width(frame, name) {
                                if let Some((_, _, shared)) = fixed_string_copybacks
                                    .iter()
                                    .find(|(existing, _, _)| Rc::ptr_eq(existing, &value))
                                {
                                    bound.push(BoundArgument::Reference(shared.clone()));
                                    continue;
                                }
                                let temporary = Rc::new(RefCell::new(value.borrow().clone()));
                                bound.push(BoundArgument::Reference(temporary.clone()));
                                fixed_string_copybacks.push((value, width, temporary));
                                continue;
                            }
                            bound.push(BoundArgument::Reference(value));
                            continue;
                        }
                    }
                }
                if let Expr::Index { target, args, .. } = expression {
                    let array = expr_name(target)
                        .and_then(|name| self.lookup_slot(frame, name))
                        .filter(|value| matches!(&*value.borrow(), Value::Array(_)));
                    if let Some(array) = array {
                        let indices = self.array_arguments(
                            args,
                            frame,
                            line.unwrap_or(expression.span().line),
                        )?;
                        if let Some((_, _, value)) =
                            copybacks.iter().find(|(existing, existing_indices, _)| {
                                Rc::ptr_eq(existing, &array) && *existing_indices == indices
                            })
                        {
                            bound.push(BoundArgument::Reference(value.clone()));
                            continue;
                        }
                        let value = Rc::new(RefCell::new(self.eval_expr(expression, frame)?));
                        bound.push(BoundArgument::Reference(value.clone()));
                        copybacks.push((array, indices, value));
                        continue;
                    }
                }
            }
            bound.push(BoundArgument::Value(self.eval_expr(expression, frame)?));
        }

        let result = self.invoke_procedure(&procedure, bound, line)?;
        for (target, width, value) in fixed_string_copybacks {
            *target.borrow_mut() = coerce_string_width(value.borrow().clone(), width, line)?;
        }
        for (array, indices, value) in copybacks {
            let mut value = value.borrow().clone();
            let mut array = array.borrow_mut();
            let Value::Array(array) = &mut *array else {
                return Err(error(
                    RuntimeErrorKind::TypeMismatch,
                    "ByRef array target is no longer an array",
                    line,
                ));
            };
            if let Value::String(default) = array.element_default.as_ref() {
                if !default.is_empty() {
                    value = coerce_string_width(value, default.encode_utf16().count(), line)?;
                }
            }
            let offset = array_offset(array, &indices, line.unwrap_or(0))?;
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
        if let Some(result) = call_builtin(name, &args, line, self.option_compare_text()) {
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
        let value = match expr {
            Expr::EvaluateShortcut { text, .. } => self.evaluate_shortcut(text, line, true),
            _ => self.eval_expr(expr, frame),
        }
        .map_err(|failure| {
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
            Value::Nothing => Err(error(
                RuntimeErrorKind::ObjectVariableNotSet,
                "object variable or With block variable not set",
                Some(line),
            )),
            _ => Err(error(
                RuntimeErrorKind::TypeMismatch,
                "VBA member access requires an object",
                Some(line),
            )),
        }
    }

    fn evaluate_shortcut(
        &mut self,
        text: &str,
        line: u32,
        object_context: bool,
    ) -> Result<Value, RuntimeError> {
        let value = self
            .host_call(None, "Evaluate", &[Value::String(text.to_string())], line)?
            .ok_or_else(|| {
                error(
                    RuntimeErrorKind::Unsupported,
                    "Evaluate requires a configured VBA host",
                    Some(line),
                )
            })?;
        if object_context {
            return match value {
                Value::Object(_) => Ok(value),
                _ => Err(error(
                    RuntimeErrorKind::TypeMismatch,
                    "evaluated expression does not return an object",
                    Some(line),
                )),
            };
        }
        self.let_value(value, line)
    }

    fn let_value(&mut self, value: Value, line: u32) -> Result<Value, RuntimeError> {
        match value {
            Value::Object(receiver) => self.host_get(&receiver, "Value", line)?.ok_or_else(|| {
                error(
                    RuntimeErrorKind::TypeMismatch,
                    format!("{} has no default scalar Value property", receiver.kind),
                    Some(line),
                )
            }),
            Value::Nothing => Err(error(
                RuntimeErrorKind::ObjectVariableNotSet,
                "object variable or With block variable not set",
                Some(line),
            )),
            value => Ok(value),
        }
    }

    fn host_call(
        &mut self,
        receiver: Option<&ObjectRef>,
        name: &str,
        args: &[Value],
        line: u32,
    ) -> Result<Option<Value>, RuntimeError> {
        if let Some(receiver) = receiver {
            if self.internal_objects.contains_key(&receiver.handle) {
                return self.internal_call(receiver, name, args, line).map(Some);
            }
        }
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
        if let Some(object) = self.internal_objects.get(&receiver.handle) {
            return match (object, name.to_ascii_lowercase().as_str()) {
                (InternalObject::Collection(entries), "count") => {
                    Ok(Some(Value::Integer(entries.len() as i64)))
                }
                _ => Ok(None),
            };
        }
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
        if self.internal_objects.contains_key(&receiver.handle) {
            return Ok(false);
        }
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
        if let Some(InternalObject::Collection(entries)) =
            self.internal_objects.get(&receiver.handle)
        {
            return Ok(Some(
                entries.iter().map(|entry| entry.value.clone()).collect(),
            ));
        }
        let Some(host) = self.host.as_deref_mut() else {
            return Ok(None);
        };
        host.enumerate(receiver)
            .map_err(|message| error(RuntimeErrorKind::Host, message, Some(line)))
    }

    fn internal_call(
        &mut self,
        receiver: &ObjectRef,
        name: &str,
        args: &[Value],
        line: u32,
    ) -> Result<Value, RuntimeError> {
        let object = self
            .internal_objects
            .get_mut(&receiver.handle)
            .ok_or_else(|| {
                error(
                    RuntimeErrorKind::ObjectVariableNotSet,
                    "internal VBA object is no longer available",
                    Some(line),
                )
            })?;
        match object {
            InternalObject::Collection(entries) if name.eq_ignore_ascii_case("add") => {
                if !(1..=4).contains(&args.len()) {
                    return Err(error(
                        RuntimeErrorKind::ArgumentCount,
                        format!(
                            "Collection.Add expects 1 to 4 arguments, received {}",
                            args.len()
                        ),
                        Some(line),
                    ));
                }
                let key_value = args
                    .get(1)
                    .filter(|value| !matches!(value, Value::Missing | Value::Empty));
                let key = key_value
                    .map(|value| {
                        text(value).map_err(|message| {
                            error(RuntimeErrorKind::TypeMismatch, message, Some(line))
                        })
                    })
                    .transpose()?;
                if key.as_ref().is_some_and(|key| {
                    entries.iter().any(|entry| {
                        entry
                            .key
                            .as_ref()
                            .is_some_and(|existing| existing.eq_ignore_ascii_case(key))
                    })
                }) {
                    return Err(raised_error(
                        457,
                        "Collection".to_string(),
                        "this key is already associated with an element of this collection"
                            .to_string(),
                        line,
                    ));
                }
                let before = args
                    .get(2)
                    .filter(|value| !matches!(value, Value::Missing | Value::Empty));
                let after = args
                    .get(3)
                    .filter(|value| !matches!(value, Value::Missing | Value::Empty));
                if before.is_some() && after.is_some() {
                    return Err(error(
                        RuntimeErrorKind::ArgumentCount,
                        "Collection.Add cannot specify both Before and After",
                        Some(line),
                    ));
                }
                let position = if let Some(before) = before {
                    collection_index(entries, before, line)?
                } else if let Some(after) = after {
                    collection_index(entries, after, line)? + 1
                } else {
                    entries.len()
                };
                entries.insert(
                    position,
                    CollectionEntry {
                        value: args[0].clone(),
                        key,
                    },
                );
                Ok(Value::Empty)
            }
            InternalObject::Collection(entries) if name.eq_ignore_ascii_case("item") => {
                if args.len() != 1 {
                    return Err(error(
                        RuntimeErrorKind::ArgumentCount,
                        format!(
                            "Collection.Item expects 1 argument, received {}",
                            args.len()
                        ),
                        Some(line),
                    ));
                }
                Ok(entries[collection_index(entries, &args[0], line)?]
                    .value
                    .clone())
            }
            InternalObject::Collection(entries) if name.eq_ignore_ascii_case("remove") => {
                if args.len() != 1 {
                    return Err(error(
                        RuntimeErrorKind::ArgumentCount,
                        format!(
                            "Collection.Remove expects 1 argument, received {}",
                            args.len()
                        ),
                        Some(line),
                    ));
                }
                let index = collection_index(entries, &args[0], line)?;
                entries.remove(index);
                Ok(Value::Empty)
            }
            InternalObject::Collection(_) => Err(error(
                RuntimeErrorKind::Unsupported,
                format!("Collection method is not available: {name}"),
                Some(line),
            )),
        }
    }

    fn array_arguments(
        &mut self,
        args: &[Argument],
        frame: &mut Frame,
        line: u32,
    ) -> Result<Vec<i64>, RuntimeError> {
        let mut indices = Vec::with_capacity(args.len());
        for argument in args {
            if argument.name.is_some() {
                return Err(error(
                    RuntimeErrorKind::SubscriptOutOfRange,
                    "VBA array indices cannot be named",
                    Some(line),
                ));
            }
            let index = argument.value.as_ref().ok_or_else(|| {
                error(
                    RuntimeErrorKind::SubscriptOutOfRange,
                    "VBA array index cannot be omitted",
                    Some(line),
                )
            })?;
            indices.push(self.array_index(index, frame, line)?);
        }
        Ok(indices)
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
        vba_number: None,
        vba_source: None,
    }
}

fn empty_frame() -> Frame {
    Frame {
        procedure_name: String::new(),
        source_name: String::new(),
        values: BTreeMap::new(),
        constants: BTreeSet::new(),
        auto_new: BTreeMap::new(),
        fixed_strings: BTreeMap::new(),
        variants: BTreeSet::new(),
        static_procedure: false,
        with_objects: Vec::new(),
        error_mode: ErrorMode::Disabled,
        error_state: ErrorState::default(),
        error_handler_active: false,
        error_statement: None,
        current_statement: 0,
        gosub_returns: Vec::new(),
    }
}

fn constant_assignment_error(name: &str, line: u32) -> RuntimeError {
    error(
        RuntimeErrorKind::Unsupported,
        format!("VBA constant cannot be assigned: {name}"),
        Some(line),
    )
}

fn raised_error(number: i64, source: String, description: String, line: u32) -> RuntimeError {
    RuntimeError {
        kind: RuntimeErrorKind::UserDefined,
        message: description,
        line: Some(line),
        vba_number: Some(number),
        vba_source: Some(source),
    }
}

fn runtime_error_number(failure: &RuntimeError) -> i64 {
    failure.vba_number.unwrap_or(match failure.kind {
        RuntimeErrorKind::ProcedureNotFound | RuntimeErrorKind::UndefinedVariable => 35,
        RuntimeErrorKind::ArgumentCount => 450,
        RuntimeErrorKind::TypeMismatch => 13,
        RuntimeErrorKind::ObjectVariableNotSet => 91,
        RuntimeErrorKind::Overflow => 6,
        RuntimeErrorKind::SubscriptOutOfRange => 9,
        RuntimeErrorKind::Host => 1004,
        RuntimeErrorKind::UserDefined => 513,
        RuntimeErrorKind::DivisionByZero => 11,
        RuntimeErrorKind::Unsupported => 445,
        RuntimeErrorKind::StepLimit => 6,
        RuntimeErrorKind::CallDepth => 28,
    })
}

fn label_destination(
    labels: &BTreeMap<String, usize>,
    label: &str,
    line: Option<u32>,
) -> Result<usize, RuntimeError> {
    labels
        .get(&key(label))
        .map(|index| index + 1)
        .ok_or_else(|| {
            error(
                RuntimeErrorKind::UndefinedVariable,
                format!("VBA label not found: {label}"),
                line,
            )
        })
}

fn is_err_object(object: &ObjectRef) -> bool {
    object.handle == u64::MAX && object.kind == "Err"
}

fn err_property(frame: &Frame, name: &str, line: u32) -> Result<Value, RuntimeError> {
    if name.eq_ignore_ascii_case("number") {
        Ok(Value::Integer(frame.error_state.number))
    } else if name.eq_ignore_ascii_case("description") {
        Ok(Value::String(frame.error_state.description.clone()))
    } else if name.eq_ignore_ascii_case("source") {
        Ok(Value::String(frame.error_state.source.clone()))
    } else if name.eq_ignore_ascii_case("erl") {
        Ok(Value::Integer(frame.error_state.line.unwrap_or(0) as i64))
    } else if name.eq_ignore_ascii_case("helpfile") {
        Ok(Value::String(String::new()))
    } else if name.eq_ignore_ascii_case("helpcontext") || name.eq_ignore_ascii_case("lastdllerror")
    {
        Ok(Value::Integer(0))
    } else {
        Err(error(
            RuntimeErrorKind::Unsupported,
            format!("Err property is not available: {name}"),
            Some(line),
        ))
    }
}

fn err_set(frame: &mut Frame, name: &str, value: Value, line: u32) -> Result<(), RuntimeError> {
    let mismatch = |message| error(RuntimeErrorKind::TypeMismatch, message, Some(line));
    if name.eq_ignore_ascii_case("number") {
        frame.error_state.number = number(&value).map_err(mismatch)?.round_ties_even() as i64;
    } else if name.eq_ignore_ascii_case("description") {
        frame.error_state.description = text(&value).map_err(mismatch)?;
    } else if name.eq_ignore_ascii_case("source") {
        frame.error_state.source = text(&value).map_err(mismatch)?;
    } else {
        return Err(error(
            RuntimeErrorKind::Unsupported,
            format!("Err property is not writable: {name}"),
            Some(line),
        ));
    }
    Ok(())
}

fn err_call(
    frame: &mut Frame,
    name: &str,
    args: &[Value],
    line: u32,
) -> Result<Value, RuntimeError> {
    if name.eq_ignore_ascii_case("clear") {
        if !args.is_empty() {
            return Err(error(
                RuntimeErrorKind::ArgumentCount,
                format!("Err.Clear expects 0 arguments, received {}", args.len()),
                Some(line),
            ));
        }
        frame.error_state = ErrorState::default();
        return Ok(Value::Empty);
    }
    if name.eq_ignore_ascii_case("raise") {
        if !(1..=5).contains(&args.len()) {
            return Err(error(
                RuntimeErrorKind::ArgumentCount,
                format!(
                    "Err.Raise expects 1 to 5 arguments, received {}",
                    args.len()
                ),
                Some(line),
            ));
        }
        let mismatch = |message| error(RuntimeErrorKind::TypeMismatch, message, Some(line));
        let number = number(&args[0]).map_err(mismatch)?.round_ties_even() as i64;
        let source = args
            .get(1)
            .filter(|value| !matches!(value, Value::Missing | Value::Empty))
            .map(|value| text(value).map_err(mismatch))
            .transpose()?
            .unwrap_or_else(|| frame.source_name.clone());
        let description = args
            .get(2)
            .filter(|value| !matches!(value, Value::Missing | Value::Empty))
            .map(|value| text(value).map_err(mismatch))
            .transpose()?
            .unwrap_or_else(|| format!("VBA error {number}"));
        return Err(raised_error(number, source, description, line));
    }
    Err(error(
        RuntimeErrorKind::Unsupported,
        format!("Err method is not available: {name}"),
        Some(line),
    ))
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

fn collection_index(
    entries: &[CollectionEntry],
    selector: &Value,
    line: u32,
) -> Result<usize, RuntimeError> {
    let index = match selector {
        Value::Integer(value) => usize::try_from(*value)
            .ok()
            .and_then(|value| value.checked_sub(1)),
        Value::Double(value) if value.is_finite() && value.fract() == 0.0 => {
            usize::try_from(*value as i64)
                .ok()
                .and_then(|value| value.checked_sub(1))
        }
        Value::String(key) => entries.iter().position(|entry| {
            entry
                .key
                .as_ref()
                .is_some_and(|existing| existing.eq_ignore_ascii_case(key))
        }),
        _ => {
            return Err(error(
                RuntimeErrorKind::TypeMismatch,
                "Collection index must be a one-based number or String key",
                Some(line),
            ))
        }
    };
    index.filter(|index| *index < entries.len()).ok_or_else(|| {
        raised_error(
            5,
            "Collection".to_string(),
            "invalid procedure call or argument".to_string(),
            line,
        )
    })
}

fn array_offset(array: &ArrayValue, indices: &[i64], line: u32) -> Result<usize, RuntimeError> {
    if indices.len() != array.dimensions.len() || indices.is_empty() {
        return Err(error(
            RuntimeErrorKind::SubscriptOutOfRange,
            format!(
                "VBA array has {} dimensions, but {} indices were supplied",
                array.dimensions.len(),
                indices.len()
            ),
            Some(line),
        ));
    }
    let mut offset = 0usize;
    for (index, dimension) in indices.iter().zip(&array.dimensions) {
        let upper = dimension
            .length
            .checked_sub(1)
            .map(|offset| dimension.lower_bound.saturating_add(offset as i64))
            .unwrap_or_else(|| dimension.lower_bound.saturating_sub(1));
        if dimension.length == 0 || *index < dimension.lower_bound || *index > upper {
            return Err(error(
                RuntimeErrorKind::SubscriptOutOfRange,
                format!(
                    "VBA array index {index} is outside {} To {upper}",
                    dimension.lower_bound
                ),
                Some(line),
            ));
        }
        offset = offset
            .checked_mul(dimension.length)
            .and_then(|offset| offset.checked_add((*index - dimension.lower_bound) as usize))
            .ok_or_else(|| {
                error(
                    RuntimeErrorKind::Overflow,
                    "VBA array offset overflow",
                    Some(line),
                )
            })?;
    }
    Ok(offset)
}

fn preservable_dimensions(existing: &[ArrayDimension], replacement: &[ArrayDimension]) -> bool {
    existing.len() == replacement.len()
        && !existing.is_empty()
        && existing
            .iter()
            .zip(replacement)
            .enumerate()
            .all(|(index, (old, new))| {
                old.lower_bound == new.lower_bound
                    && (index + 1 == existing.len() || old.length == new.length)
            })
}

fn preserve_array_values(existing: &ArrayValue, replacement: &mut ArrayValue) {
    let (Some(old_last), Some(new_last)) =
        (existing.dimensions.last(), replacement.dimensions.last())
    else {
        return;
    };
    let shared = old_last.length.min(new_last.length);
    if shared == 0 {
        return;
    }
    let prefixes = existing.dimensions[..existing.dimensions.len() - 1]
        .iter()
        .map(|dimension| dimension.length)
        .product::<usize>();
    for prefix in 0..prefixes {
        let old_start = prefix * old_last.length;
        let new_start = prefix * new_last.length;
        replacement.values[new_start..new_start + shared]
            .clone_from_slice(&existing.values[old_start..old_start + shared]);
    }
}

fn call_builtin(
    name: &str,
    args: &[Value],
    line: Option<u32>,
    option_compare_text: bool,
) -> Option<Result<Value, RuntimeError>> {
    let name = name.to_ascii_lowercase();
    let known = matches!(
        name.as_str(),
        "abs"
            | "array"
            | "cbool"
            | "cdbl"
            | "chr"
            | "chrw"
            | "clng"
            | "cstr"
            | "asc"
            | "ascw"
            | "instr"
            | "instrrev"
            | "isarray"
            | "isempty"
            | "ismissing"
            | "isnull"
            | "isnumeric"
            | "isobject"
            | "lbound"
            | "lcase"
            | "left"
            | "len"
            | "ltrim"
            | "mid"
            | "replace"
            | "right"
            | "rtrim"
            | "space"
            | "split"
            | "join"
            | "string"
            | "strreverse"
            | "trim"
            | "typename"
            | "ubound"
            | "ucase"
            | "vartype"
    );
    if !known {
        return None;
    }
    Some((|| {
        if matches!(
            name.as_str(),
            "asc"
                | "ascw"
                | "chr"
                | "chrw"
                | "instr"
                | "instrrev"
                | "join"
                | "left"
                | "ltrim"
                | "mid"
                | "replace"
                | "right"
                | "rtrim"
                | "space"
                | "split"
                | "string"
                | "strreverse"
        ) {
            return call_string_builtin(&name, args, line, option_compare_text);
        }
        if name == "array" {
            return Ok(Value::Array(ArrayValue {
                dimensions: vec![ArrayDimension {
                    lower_bound: 0,
                    length: args.len(),
                }],
                values: args.to_vec(),
                element_default: Box::new(Value::Empty),
                resizable: true,
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
        if matches!(
            name.as_str(),
            "isarray" | "isempty" | "isnull" | "isnumeric" | "isobject" | "typename" | "vartype"
        ) {
            if args.len() != 1 {
                return Err(error(
                    RuntimeErrorKind::ArgumentCount,
                    format!("{name} expects 1 argument, received {}", args.len()),
                    line,
                ));
            }
            let value = &args[0];
            return Ok(match name.as_str() {
                "isarray" => Value::Boolean(matches!(value, Value::Array(_))),
                "isempty" => Value::Boolean(matches!(value, Value::Empty)),
                "isnull" => Value::Boolean(matches!(value, Value::Null)),
                "isnumeric" => Value::Boolean(match value {
                    Value::Integer(_) | Value::Double(_) => true,
                    Value::String(value) => value.parse::<f64>().is_ok(),
                    _ => false,
                }),
                "isobject" => Value::Boolean(matches!(value, Value::Object(_) | Value::Nothing)),
                "typename" => Value::String(value_type_name(value)),
                "vartype" => Value::Integer(value_var_type(value)),
                _ => unreachable!(),
            });
        }
        if matches!(name.as_str(), "lbound" | "ubound") {
            if !(1..=2).contains(&args.len()) {
                return Err(error(
                    RuntimeErrorKind::ArgumentCount,
                    format!("{name} expects 1 or 2 arguments, received {}", args.len()),
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
            let dimension = if args.len() == 1 {
                1
            } else {
                let value = number(&args[1])
                    .map_err(|message| error(RuntimeErrorKind::TypeMismatch, message, line))?;
                let rounded = value.round_ties_even();
                if !rounded.is_finite() || rounded < 1.0 || rounded > usize::MAX as f64 {
                    return Err(error(
                        RuntimeErrorKind::SubscriptOutOfRange,
                        "array dimension is out of range",
                        line,
                    ));
                }
                rounded as usize
            };
            let bound = if name == "lbound" {
                array.lower_bound(dimension)
            } else {
                array.upper_bound(dimension)
            }
            .ok_or_else(|| {
                error(
                    RuntimeErrorKind::SubscriptOutOfRange,
                    format!("array dimension {dimension} is out of range"),
                    line,
                )
            })?;
            return Ok(Value::Integer(bound));
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

fn call_string_builtin(
    name: &str,
    args: &[Value],
    line: Option<u32>,
    option_compare_text: bool,
) -> Result<Value, RuntimeError> {
    let mismatch = |message| error(RuntimeErrorKind::TypeMismatch, message, line);
    let wrong_count = |expected: &str| {
        error(
            RuntimeErrorKind::ArgumentCount,
            format!("{name} expects {expected}, received {}", args.len()),
            line,
        )
    };
    let nullable_text = |value: &Value| -> Result<Option<String>, RuntimeError> {
        match value {
            Value::Null => Ok(None),
            _ => text(value).map(Some).map_err(mismatch),
        }
    };

    match name {
        "ltrim" | "rtrim" | "strreverse" => {
            if args.len() != 1 {
                return Err(wrong_count("1 argument"));
            }
            let Some(value) = nullable_text(&args[0])? else {
                return Ok(Value::Null);
            };
            Ok(Value::String(match name {
                "ltrim" => value.trim_start_matches(' ').to_string(),
                "rtrim" => value.trim_end_matches(' ').to_string(),
                "strreverse" => value.chars().rev().collect(),
                _ => unreachable!(),
            }))
        }
        "asc" | "ascw" => {
            if args.len() != 1 {
                return Err(wrong_count("1 argument"));
            }
            let value = nullable_text(&args[0])?
                .ok_or_else(|| mismatch("invalid use of Null".to_string()))?;
            let unit = value.encode_utf16().next().ok_or_else(|| {
                invalid_procedure_call(format!("{name} requires a non-empty String"), line)
            })?;
            let value = if name == "ascw" {
                i64::from(unit as i16)
            } else {
                i64::from(unit)
            };
            Ok(Value::Integer(value))
        }
        "chr" | "chrw" => {
            if args.len() != 1 {
                return Err(wrong_count("1 argument"));
            }
            let value = integer_argument(&args[0], line)?;
            let valid = (name == "chrw" && (-32_768..=65_535).contains(&value))
                || (name == "chr" && (0..=255).contains(&value));
            let unit = if valid {
                value as u16
            } else {
                return Err(invalid_procedure_call(
                    format!("character code is out of range for {name}"),
                    line,
                ));
            };
            Ok(Value::String(String::from_utf16_lossy(&[unit])))
        }
        "left" | "right" => {
            if args.len() != 2 {
                return Err(wrong_count("2 arguments"));
            }
            let Some(value) = nullable_text(&args[0])? else {
                return Ok(Value::Null);
            };
            let length = nonnegative_length(&args[1], line)?;
            let units = value.encode_utf16().collect::<Vec<_>>();
            let length = length.min(units.len());
            let selected = if name == "left" {
                &units[..length]
            } else {
                &units[units.len() - length..]
            };
            Ok(Value::String(String::from_utf16_lossy(selected)))
        }
        "mid" => {
            if !(2..=3).contains(&args.len()) {
                return Err(wrong_count("2 or 3 arguments"));
            }
            let Some(value) = nullable_text(&args[0])? else {
                return Ok(Value::Null);
            };
            let start = positive_position(&args[1], line)?;
            let units = value.encode_utf16().collect::<Vec<_>>();
            if start > units.len() {
                return Ok(Value::String(String::new()));
            }
            let available = units.len() - start;
            let length = if args.len() == 3 {
                nonnegative_length(&args[2], line)?.min(available)
            } else {
                available
            };
            Ok(Value::String(String::from_utf16_lossy(
                &units[start..start + length],
            )))
        }
        "instr" => {
            if !(2..=4).contains(&args.len()) {
                return Err(wrong_count("2 to 4 arguments"));
            }
            let (start, source_index, needle_index, compare_index) = if args.len() == 2 {
                (0, 0, 1, None)
            } else {
                (
                    positive_position(&args[0], line)?,
                    1,
                    2,
                    (args.len() == 4).then_some(3),
                )
            };
            let Some(source) = nullable_text(&args[source_index])? else {
                return Ok(Value::Null);
            };
            let Some(needle) = nullable_text(&args[needle_index])? else {
                return Ok(Value::Null);
            };
            let compare = compare_mode(
                compare_index.map(|index| &args[index]),
                option_compare_text,
                line,
            )?;
            let source = source.encode_utf16().collect::<Vec<_>>();
            let needle = needle.encode_utf16().collect::<Vec<_>>();
            Ok(Value::Integer(
                utf16_find(&source, &needle, start, compare)
                    .map(|offset| offset as i64 + 1)
                    .unwrap_or(0),
            ))
        }
        "instrrev" => {
            if !(2..=4).contains(&args.len()) {
                return Err(wrong_count("2 to 4 arguments"));
            }
            let Some(source) = nullable_text(&args[0])? else {
                return Ok(Value::Null);
            };
            let Some(needle) = nullable_text(&args[1])? else {
                return Ok(Value::Null);
            };
            let source = source.encode_utf16().collect::<Vec<_>>();
            let needle = needle.encode_utf16().collect::<Vec<_>>();
            let start = match args.get(2) {
                None | Some(Value::Missing) => source.len(),
                Some(value) => {
                    let value = integer_argument(value, line)?;
                    if value == -1 {
                        source.len()
                    } else if value < 1 {
                        return Err(invalid_procedure_call(
                            "InStrRev start must be positive or -1".to_string(),
                            line,
                        ));
                    } else {
                        usize::try_from(value)
                            .unwrap_or(usize::MAX)
                            .min(source.len())
                    }
                }
            };
            let compare = compare_mode(args.get(3), option_compare_text, line)?;
            Ok(Value::Integer(
                utf16_rfind(&source, &needle, start, compare)
                    .map(|offset| offset as i64 + 1)
                    .unwrap_or(0),
            ))
        }
        "replace" => {
            if !(3..=6).contains(&args.len()) {
                return Err(wrong_count("3 to 6 arguments"));
            }
            let source = nullable_text(&args[0])?
                .ok_or_else(|| mismatch("invalid use of Null".to_string()))?;
            let needle = nullable_text(&args[1])?
                .ok_or_else(|| mismatch("invalid use of Null".to_string()))?;
            let replacement = nullable_text(&args[2])?
                .ok_or_else(|| mismatch("invalid use of Null".to_string()))?;
            let start = match args.get(3) {
                None | Some(Value::Missing) => 0,
                Some(value) => positive_position(value, line)?,
            };
            let count = match args.get(4) {
                None | Some(Value::Missing) => usize::MAX,
                Some(value) => {
                    let value = integer_argument(value, line)?;
                    if value == -1 {
                        usize::MAX
                    } else if value < 0 {
                        return Err(invalid_procedure_call(
                            "Replace count must be nonnegative or -1".to_string(),
                            line,
                        ));
                    } else {
                        usize::try_from(value).unwrap_or(usize::MAX)
                    }
                }
            };
            let compare = compare_mode(args.get(5), option_compare_text, line)?;
            Ok(Value::String(utf16_replace(
                &source,
                &needle,
                &replacement,
                start,
                count,
                compare,
            )))
        }
        "split" => {
            if !(1..=4).contains(&args.len()) {
                return Err(wrong_count("1 to 4 arguments"));
            }
            let source = nullable_text(&args[0])?
                .ok_or_else(|| mismatch("invalid use of Null".to_string()))?;
            let delimiter = match args.get(1) {
                None | Some(Value::Missing) => " ".to_string(),
                Some(value) => nullable_text(value)?
                    .ok_or_else(|| mismatch("invalid use of Null".to_string()))?,
            };
            let limit = match args.get(2) {
                None | Some(Value::Missing) => usize::MAX,
                Some(value) => {
                    let value = integer_argument(value, line)?;
                    if value == -1 {
                        usize::MAX
                    } else if value < 0 {
                        return Err(invalid_procedure_call(
                            "Split limit must be nonnegative or -1".to_string(),
                            line,
                        ));
                    } else {
                        usize::try_from(value).unwrap_or(usize::MAX)
                    }
                }
            };
            let compare = compare_mode(args.get(3), option_compare_text, line)?;
            let values = utf16_split(&source, &delimiter, limit, compare)
                .into_iter()
                .map(Value::String)
                .collect::<Vec<_>>();
            Ok(Value::Array(ArrayValue {
                dimensions: vec![ArrayDimension {
                    lower_bound: 0,
                    length: values.len(),
                }],
                values,
                element_default: Box::new(Value::String(String::new())),
                resizable: true,
            }))
        }
        "join" => {
            if !(1..=2).contains(&args.len()) {
                return Err(wrong_count("1 or 2 arguments"));
            }
            let Value::Array(array) = &args[0] else {
                return Err(mismatch("Join requires an array".to_string()));
            };
            if array.dimensions.len() != 1 {
                return Err(mismatch(
                    "Join requires a one-dimensional array".to_string(),
                ));
            }
            let delimiter = match args.get(1) {
                None | Some(Value::Missing) => " ".to_string(),
                Some(value) => nullable_text(value)?
                    .ok_or_else(|| mismatch("invalid use of Null".to_string()))?,
            };
            let values = array
                .values
                .iter()
                .map(|value| text(value).map_err(mismatch))
                .collect::<Result<Vec<_>, _>>()?;
            Ok(Value::String(values.join(&delimiter)))
        }
        "space" => {
            if args.len() != 1 {
                return Err(wrong_count("1 argument"));
            }
            let length = nonnegative_length(&args[0], line)?;
            if length > 1_000_000 {
                return Err(error(
                    RuntimeErrorKind::Overflow,
                    "Space result exceeds the browser runtime limit",
                    line,
                ));
            }
            Ok(Value::String(" ".repeat(length)))
        }
        "string" => {
            if args.len() != 2 {
                return Err(wrong_count("2 arguments"));
            }
            let length = nonnegative_length(&args[0], line)?;
            if length > 1_000_000 {
                return Err(error(
                    RuntimeErrorKind::Overflow,
                    "String result exceeds the browser runtime limit",
                    line,
                ));
            }
            let character = match &args[1] {
                Value::String(value) => value.chars().next().ok_or_else(|| {
                    invalid_procedure_call("String requires a character".to_string(), line)
                })?,
                value => {
                    let code = integer_argument(value, line)?;
                    char::from_u32((code & 0xff) as u32).unwrap_or('\u{fffd}')
                }
            };
            Ok(Value::String(character.to_string().repeat(length)))
        }
        _ => unreachable!(),
    }
}

fn builtin_constant(name: &str) -> Option<Value> {
    Some(match name.to_ascii_lowercase().as_str() {
        "vbbinarycompare" => Value::Integer(0),
        "vbtextcompare" => Value::Integer(1),
        "vbusecompareoption" => Value::Integer(-1),
        "vbcrlf" | "vbnewline" => Value::String("\r\n".to_string()),
        "vbcr" => Value::String("\r".to_string()),
        "vblf" => Value::String("\n".to_string()),
        "vbtab" => Value::String("\t".to_string()),
        "vbback" => Value::String("\u{8}".to_string()),
        "vbformfeed" => Value::String("\u{c}".to_string()),
        "vbverticaltab" => Value::String("\u{b}".to_string()),
        "vbnullstring" => Value::String(String::new()),
        _ => return None,
    })
}

fn integer_argument(value: &Value, line: Option<u32>) -> Result<i64, RuntimeError> {
    let value = number(value)
        .map_err(|message| error(RuntimeErrorKind::TypeMismatch, message, line))?
        .round_ties_even();
    if !value.is_finite() || value < i64::MIN as f64 || value > i64::MAX as f64 {
        Err(error(
            RuntimeErrorKind::Overflow,
            "numeric argument is outside the supported integer range",
            line,
        ))
    } else {
        Ok(value as i64)
    }
}

fn nonnegative_length(value: &Value, line: Option<u32>) -> Result<usize, RuntimeError> {
    let value = integer_argument(value, line)?;
    usize::try_from(value)
        .map_err(|_| invalid_procedure_call("String length cannot be negative".to_string(), line))
}

fn positive_position(value: &Value, line: Option<u32>) -> Result<usize, RuntimeError> {
    let value = integer_argument(value, line)?;
    if value < 1 {
        Err(invalid_procedure_call(
            "String position must be positive".to_string(),
            line,
        ))
    } else {
        Ok(usize::try_from(value - 1).unwrap_or(usize::MAX))
    }
}

fn compare_mode(
    value: Option<&Value>,
    option_compare_text: bool,
    line: Option<u32>,
) -> Result<bool, RuntimeError> {
    let mode = match value {
        None | Some(Value::Missing) => return Ok(option_compare_text),
        Some(value) => integer_argument(value, line)?,
    };
    match mode {
        -1 => Ok(option_compare_text),
        0 => Ok(false),
        1 => Ok(true),
        _ => Err(invalid_procedure_call(
            format!("invalid VBA comparison mode: {mode}"),
            line,
        )),
    }
}

fn invalid_procedure_call(message: String, line: Option<u32>) -> RuntimeError {
    RuntimeError {
        kind: RuntimeErrorKind::UserDefined,
        message,
        line,
        vba_number: Some(5),
        vba_source: Some("VBA".to_string()),
    }
}

fn fixed_array_error(name: &str, line: Option<u32>) -> RuntimeError {
    RuntimeError {
        kind: RuntimeErrorKind::UserDefined,
        message: format!("this array is fixed or temporarily locked: {name}"),
        line,
        vba_number: Some(10),
        vba_source: Some("VBA".to_string()),
    }
}

fn utf16_equal(left: &[u16], right: &[u16], text_compare: bool) -> bool {
    if left == right {
        return true;
    }
    text_compare
        && String::from_utf16_lossy(left).to_lowercase()
            == String::from_utf16_lossy(right).to_lowercase()
}

fn utf16_find(source: &[u16], needle: &[u16], start: usize, text_compare: bool) -> Option<usize> {
    if start > source.len() {
        return None;
    }
    if needle.is_empty() {
        return Some(start);
    }
    source[start..]
        .windows(needle.len())
        .position(|window| utf16_equal(window, needle, text_compare))
        .map(|offset| start + offset)
}

fn utf16_rfind(source: &[u16], needle: &[u16], start: usize, text_compare: bool) -> Option<usize> {
    if needle.is_empty() {
        return Some(start.min(source.len()));
    }
    if needle.len() > source.len() || start == 0 {
        return None;
    }
    let last_start = start.saturating_sub(1).min(source.len() - needle.len());
    (0..=last_start).rev().find(|offset| {
        utf16_equal(
            &source[*offset..*offset + needle.len()],
            needle,
            text_compare,
        )
    })
}

fn utf16_replace(
    source: &str,
    needle: &str,
    replacement: &str,
    start: usize,
    count: usize,
    text_compare: bool,
) -> String {
    let source = source.encode_utf16().collect::<Vec<_>>();
    if start >= source.len() {
        return String::new();
    }
    let needle = needle.encode_utf16().collect::<Vec<_>>();
    let replacement = replacement.encode_utf16().collect::<Vec<_>>();
    if needle.is_empty() || count == 0 {
        return String::from_utf16_lossy(&source[start..]);
    }
    let mut result = Vec::new();
    let mut cursor = start;
    let mut replaced = 0;
    while replaced < count {
        let Some(found) = utf16_find(&source, &needle, cursor, text_compare) else {
            break;
        };
        result.extend_from_slice(&source[cursor..found]);
        result.extend_from_slice(&replacement);
        cursor = found + needle.len();
        replaced += 1;
    }
    result.extend_from_slice(&source[cursor..]);
    String::from_utf16_lossy(&result)
}

fn utf16_split(source: &str, delimiter: &str, limit: usize, text_compare: bool) -> Vec<String> {
    if source.is_empty() || limit == 0 {
        return Vec::new();
    }
    if delimiter.is_empty() || limit == 1 {
        return vec![source.to_string()];
    }
    let source = source.encode_utf16().collect::<Vec<_>>();
    let delimiter = delimiter.encode_utf16().collect::<Vec<_>>();
    let mut values = Vec::new();
    let mut cursor = 0;
    while values.len().saturating_add(1) < limit {
        let Some(found) = utf16_find(&source, &delimiter, cursor, text_compare) else {
            break;
        };
        values.push(String::from_utf16_lossy(&source[cursor..found]));
        cursor = found + delimiter.len();
    }
    values.push(String::from_utf16_lossy(&source[cursor..]));
    values
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
        Literal::Empty => Value::Empty,
        Literal::Nothing => Value::Nothing,
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
        "date" => Value::Double(0.0),
        "string" => Value::String(String::new()),
        "object" | "application" | "workbook" | "worksheet" | "range" | "chart" | "shape"
        | "collection" | "dictionary" => Value::Nothing,
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

fn value_type_name(value: &Value) -> String {
    match value {
        Value::Empty => "Empty".to_string(),
        Value::Missing => "Missing".to_string(),
        Value::Nothing => "Nothing".to_string(),
        Value::Null => "Null".to_string(),
        Value::Boolean(_) => "Boolean".to_string(),
        Value::Integer(_) => "Long".to_string(),
        Value::Double(_) => "Double".to_string(),
        Value::String(_) => "String".to_string(),
        Value::Array(_) => "Variant()".to_string(),
        Value::Object(object) => object.kind.clone(),
    }
}

fn value_var_type(value: &Value) -> i64 {
    match value {
        Value::Empty => 0,
        Value::Null => 1,
        Value::Integer(_) => 3,
        Value::Double(_) => 5,
        Value::String(_) => 8,
        Value::Object(_) | Value::Nothing => 9,
        Value::Boolean(_) => 11,
        Value::Array(_) => 8_192 + 12,
        Value::Missing => 12,
    }
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
        Value::Nothing => Err("object variable or With block variable not set".to_string()),
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
        Value::Nothing => Err("object variable or With block variable not set".to_string()),
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

fn binary(
    op: BinaryOp,
    lhs: Value,
    rhs: Value,
    option_compare_text: bool,
) -> Result<Value, (RuntimeErrorKind, String)> {
    use BinaryOp::*;
    if op == Is {
        return match (lhs, rhs) {
            (Value::Nothing, Value::Nothing) => Ok(Value::Boolean(true)),
            (Value::Object(_), Value::Nothing) | (Value::Nothing, Value::Object(_)) => {
                Ok(Value::Boolean(false))
            }
            (Value::Object(left), Value::Object(right)) => Ok(Value::Boolean(left == right)),
            _ => Err((
                RuntimeErrorKind::TypeMismatch,
                "Is requires two object expressions".to_string(),
            )),
        };
    }
    if matches!(
        lhs,
        Value::Array(_) | Value::Object(_) | Value::Missing | Value::Nothing
    ) || matches!(
        rhs,
        Value::Array(_) | Value::Object(_) | Value::Missing | Value::Nothing
    ) {
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
        Is => unreachable!(),
        Like => Ok(Value::Boolean(
            like_pattern(
                &text(&lhs).map_err(mismatch)?,
                &text(&rhs).map_err(mismatch)?,
                option_compare_text,
            )
            .map_err(|message| (RuntimeErrorKind::TypeMismatch, message))?,
        )),
    }
}

#[derive(Debug)]
enum LikeToken {
    Literal(char),
    AnyOne,
    AnyMany,
    Digit,
    Class {
        negated: bool,
        ranges: Vec<(char, char)>,
    },
}

fn like_pattern(value: &str, pattern: &str, text_compare: bool) -> Result<bool, String> {
    let tokens = parse_like_pattern(pattern)?;
    let value = value.chars().collect::<Vec<_>>();
    let mut memo = BTreeMap::new();
    Ok(like_matches(&value, &tokens, 0, 0, text_compare, &mut memo))
}

fn parse_like_pattern(pattern: &str) -> Result<Vec<LikeToken>, String> {
    let characters = pattern.chars().collect::<Vec<_>>();
    let mut tokens = Vec::new();
    let mut index = 0;
    while index < characters.len() {
        match characters[index] {
            '?' => tokens.push(LikeToken::AnyOne),
            '*' => tokens.push(LikeToken::AnyMany),
            '#' => tokens.push(LikeToken::Digit),
            '[' => {
                let close = characters[index + 1..]
                    .iter()
                    .position(|character| *character == ']')
                    .map(|offset| index + 1 + offset)
                    .ok_or_else(|| "invalid Like pattern: missing ]".to_string())?;
                let mut cursor = index + 1;
                let negated = cursor < close && characters[cursor] == '!';
                if negated {
                    cursor += 1;
                }
                if cursor == close {
                    return Err("invalid Like pattern: empty character list".to_string());
                }
                let mut ranges = Vec::new();
                while cursor < close {
                    let start = characters[cursor];
                    if cursor + 2 < close && characters[cursor + 1] == '-' {
                        let end = characters[cursor + 2];
                        ranges.push((start, end));
                        cursor += 3;
                    } else {
                        ranges.push((start, start));
                        cursor += 1;
                    }
                }
                tokens.push(LikeToken::Class { negated, ranges });
                index = close;
            }
            character => tokens.push(LikeToken::Literal(character)),
        }
        index += 1;
    }
    Ok(tokens)
}

fn like_matches(
    value: &[char],
    pattern: &[LikeToken],
    value_index: usize,
    pattern_index: usize,
    text_compare: bool,
    memo: &mut BTreeMap<(usize, usize), bool>,
) -> bool {
    if let Some(result) = memo.get(&(value_index, pattern_index)) {
        return *result;
    }
    let result = match pattern.get(pattern_index) {
        None => value_index == value.len(),
        Some(LikeToken::AnyMany) => {
            like_matches(
                value,
                pattern,
                value_index,
                pattern_index + 1,
                text_compare,
                memo,
            ) || (value_index < value.len()
                && like_matches(
                    value,
                    pattern,
                    value_index + 1,
                    pattern_index,
                    text_compare,
                    memo,
                ))
        }
        Some(token) if value_index < value.len() => {
            let character = value[value_index];
            let matches = match token {
                LikeToken::Literal(expected) => {
                    like_character_equal(character, *expected, text_compare)
                }
                LikeToken::AnyOne => true,
                LikeToken::Digit => character.is_ascii_digit(),
                LikeToken::Class { negated, ranges } => {
                    let found = ranges.iter().any(|(start, end)| {
                        let character = like_fold(character, text_compare);
                        let start = like_fold(*start, text_compare);
                        let end = like_fold(*end, text_compare);
                        start <= character && character <= end
                    });
                    found != *negated
                }
                LikeToken::AnyMany => unreachable!(),
            };
            matches
                && like_matches(
                    value,
                    pattern,
                    value_index + 1,
                    pattern_index + 1,
                    text_compare,
                    memo,
                )
        }
        Some(_) => false,
    };
    memo.insert((value_index, pattern_index), result);
    result
}

fn like_character_equal(left: char, right: char, text_compare: bool) -> bool {
    left == right || (text_compare && left.to_lowercase().eq(right.to_lowercase()))
}

fn like_fold(value: char, text_compare: bool) -> char {
    if text_compare {
        value.to_lowercase().next().unwrap_or(value)
    } else {
        value
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
        Value::Nothing => return Err("object variable or With block variable not set".to_string()),
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

fn string_width(value: &Value) -> usize {
    match value {
        Value::String(value) => value.encode_utf16().count(),
        _ => 0,
    }
}

fn coerce_string_width(
    value: Value,
    width: usize,
    line: Option<u32>,
) -> Result<Value, RuntimeError> {
    let value = match value {
        Value::Null => {
            return Err(error(
                RuntimeErrorKind::TypeMismatch,
                "invalid use of Null",
                line,
            ))
        }
        value => {
            text(&value).map_err(|message| error(RuntimeErrorKind::TypeMismatch, message, line))?
        }
    };
    let mut value = value.encode_utf16().take(width).collect::<Vec<_>>();
    value.resize(width, u16::from(b' '));
    Ok(Value::String(String::from_utf16_lossy(&value)))
}

fn line_of(statement: &Statement) -> Option<u32> {
    match statement {
        Statement::Assign { span, .. }
        | Statement::SetAssign { span, .. }
        | Statement::ReDim { span, .. }
        | Statement::Erase { span, .. }
        | Statement::MidAssign(MidAssignStmt { span, .. })
        | Statement::AlignedAssign(AlignedAssignStmt { span, .. })
        | Statement::Call { span, .. }
        | Statement::Resume { span, .. }
        | Statement::GoTo { span, .. }
        | Statement::GoSub { span, .. }
        | Statement::Return { span }
        | Statement::Exit { span, .. }
        | Statement::End { span }
        | Statement::Stop { span }
        | Statement::Comment { span, .. }
        | Statement::Directive { span, .. }
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
        Statement::With { span, .. } => Some(span.line),
        Statement::OnError(mode) => Some(match mode {
            OnError::Goto { span, .. }
            | OnError::Disable { span }
            | OnError::ResumeNext { span } => span.line,
        }),
        Statement::OnBranch(branch) => Some(branch.span.line),
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
            if name.eq_ignore_ascii_case("range") || name.eq_ignore_ascii_case("evaluate") {
                let [Value::String(address)] = args else {
                    return Err(format!("{name} expects one A1 address"));
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
    fn module_variables_and_static_locals_persist_across_runtime_calls() {
        let module = parse_module(
            "Private moduleCount As Long\n\
             Public Function NextModuleValue() As Long\n\
               moduleCount = moduleCount + 1\n\
               NextModuleValue = moduleCount\n\
             End Function\n\
             Public Function NextStaticValue() As Long\n\
               Static localCount As Long\n\
               localCount = localCount + 1\n\
               NextStaticValue = localCount\n\
             End Function\n\
             Public Static Function NextProcedureStatic() As Long\n\
               Dim procedureCount As Long\n\
               procedureCount = procedureCount + 1\n\
               NextProcedureStatic = procedureCount\n\
             End Function\n",
        )
        .unwrap();
        let mut runtime = Runtime::new(&module);

        assert_eq!(
            runtime.call("NextModuleValue", vec![]).unwrap(),
            Value::Integer(1)
        );
        assert_eq!(
            runtime.call("NextModuleValue", vec![]).unwrap(),
            Value::Integer(2)
        );
        assert_eq!(
            runtime.call("NextStaticValue", vec![]).unwrap(),
            Value::Integer(1)
        );
        assert_eq!(
            runtime.call("NextStaticValue", vec![]).unwrap(),
            Value::Integer(2)
        );
        assert_eq!(
            runtime.call("NextProcedureStatic", vec![]).unwrap(),
            Value::Integer(1)
        );
        assert_eq!(
            runtime.call("NextProcedureStatic", vec![]).unwrap(),
            Value::Integer(2)
        );

        let mut fresh_runtime = Runtime::new(&module);
        assert_eq!(
            fresh_runtime.call("NextModuleValue", vec![]).unwrap(),
            Value::Integer(1)
        );
        assert_eq!(
            fresh_runtime.call("NextStaticValue", vec![]).unwrap(),
            Value::Integer(1)
        );
    }

    #[test]
    fn module_constants_and_enum_members_initialize_in_declaration_order() {
        let value = run(
            "Private Const InitialValue As Long = 10\n\
             Private Const StepValue As Long = InitialValue + 2\n\
             Private Enum WorkState\n\
               Ready = StepValue\n\
               Running\n\
             End Enum\n\
             Public Function DeclarationValues() As Long\n\
               DeclarationValues = InitialValue + StepValue + Ready + Running\n\
             End Function\n",
            "DeclarationValues",
            vec![],
        )
        .unwrap();

        assert_eq!(value, Value::Integer(47));
    }

    #[test]
    fn module_arrays_and_scalars_can_be_mutated_byref_across_procedures() {
        let value = run(
            "Private values() As Long\n\
             Private total As Long\n\
             Private Sub Increment(ByRef value As Long)\n\
               value = value + 1\n\
             End Sub\n\
             Private Sub Prepare()\n\
               ReDim values(1 To 2)\n\
               values(1) = 10\n\
               values(2) = 20\n\
               total = 5\n\
             End Sub\n\
             Private Sub UpdateGlobals()\n\
               Increment values(2)\n\
               Increment total\n\
             End Sub\n\
             Public Function GlobalByRefProbe() As Long\n\
               Prepare\n\
               UpdateGlobals\n\
               GlobalByRefProbe = values(1) + values(2) + total * 100\n\
             End Function\n",
            "GlobalByRefProbe",
            vec![],
        )
        .unwrap();

        assert_eq!(value, Value::Integer(631));
    }

    #[test]
    fn local_variables_shadow_module_variables_without_overwriting_them() {
        let value = run(
            "Private value As Long\n\
             Private Sub SetModuleValue()\n\
               value = 40\n\
             End Sub\n\
             Private Function LocalValue() As Long\n\
               Dim value As Long\n\
               value = 2\n\
               LocalValue = value\n\
             End Function\n\
             Public Function ShadowProbe() As Long\n\
               SetModuleValue\n\
               ShadowProbe = value + LocalValue()\n\
             End Function\n",
            "ShadowProbe",
            vec![],
        )
        .unwrap();

        assert_eq!(value, Value::Integer(42));
    }

    #[test]
    fn procedure_locals_are_scoped_before_their_declaration_executes() {
        let value = run(
            "Private value As Long\n\
             Private Sub SetModuleValue()\n\
               value = 40\n\
             End Sub\n\
             Private Function LateDeclaration() As Long\n\
               LateDeclaration = value\n\
               Dim value As Long\n\
             End Function\n\
             Private Function NestedDeclaration() As Long\n\
               If False Then\n\
                 Dim hidden As Long\n\
               End If\n\
               NestedDeclaration = hidden\n\
             End Function\n\
             Public Function ProcedureScopeProbe() As Long\n\
               SetModuleValue\n\
               ProcedureScopeProbe = LateDeclaration() + NestedDeclaration()\n\
             End Function\n",
            "ProcedureScopeProbe",
            vec![],
        )
        .unwrap();

        assert_eq!(value, Value::Integer(0));
    }

    #[test]
    fn module_and_local_constants_reject_assignment() {
        let module_failure = run(
            "Private Const FixedValue As Long = 7\n\
             Public Sub ChangeModuleConstant()\n\
               FixedValue = 8\n\
             End Sub\n",
            "ChangeModuleConstant",
            vec![],
        )
        .unwrap_err();
        assert_eq!(module_failure.kind, RuntimeErrorKind::Unsupported);
        assert!(module_failure
            .message
            .contains("constant cannot be assigned"));

        let local_failure = run(
            "Public Sub ChangeLocalConstant()\n\
               Const FixedValue As Long = 7\n\
               FixedValue = 8\n\
             End Sub\n",
            "ChangeLocalConstant",
            vec![],
        )
        .unwrap_err();
        assert_eq!(local_failure.kind, RuntimeErrorKind::Unsupported);
        assert!(local_failure
            .message
            .contains("constant cannot be assigned"));
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
    fn evaluates_bracket_shortcuts_in_value_object_and_assignment_contexts() {
        let module = parse_module(
            "Public Function ShortcutProbe() As String\n\
               Dim cell As Object\n\
               Dim copied As Long\n\
               [A1] = 40\n\
               Set cell = [A1]\n\
               [A2] = [A1] + 2\n\
               copied = Range(\"A2\")\n\
               ShortcutProbe = cell.Value & \"|\" & [A2] & \"|\" & copied\n\
             End Function\n",
        )
        .unwrap();
        let mut host = SheetHost::default();
        let result = execute_with_host(&module, "ShortcutProbe", vec![], &mut host).unwrap();

        assert_eq!(result, Value::String("40|42|42".to_string()));
        assert_eq!(host.cells.get(&(1, 1)), Some(&Value::Integer(40)));
        assert_eq!(host.cells.get(&(2, 1)), Some(&Value::Integer(42)));
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
    fn nothing_object_identity_and_typeof_follow_vba_object_semantics() {
        let module = parse_module(
            "Public Function ObjectSemantics() As String\n\
               Dim first As Object\n\
               Dim same As Object\n\
               Dim other As Object\n\
               ObjectSemantics = (first Is Nothing) & \"|\"\n\
               Set first = Range(\"A1\")\n\
               Set same = Range(\"A1\")\n\
               Set other = Range(\"A2\")\n\
               ObjectSemantics = ObjectSemantics & (first Is same) & \"|\" & (first Is other) & \"|\" & (TypeOf first Is Cell) & \"|\" & (TypeOf first Is Object) & \"|\" & TypeName(first) & \"|\" & VarType(first) & \"|\"\n\
               Set same = Nothing\n\
               ObjectSemantics = ObjectSemantics & (same Is Nothing) & \"|\" & IsObject(same) & \"|\" & TypeName(same)\n\
             End Function\n",
        )
        .unwrap();
        let mut host = SheetHost::default();
        let value = execute_with_host(&module, "ObjectSemantics", vec![], &mut host).unwrap();

        assert_eq!(
            value,
            Value::String("True|True|False|True|True|Cell|9|True|True|Nothing".to_string())
        );
    }

    #[test]
    fn dereferencing_nothing_raises_vba_error_91() {
        let module = parse_module(
            "Public Function MissingObject() As Long\n\
               Dim cell As Range\n\
               On Error Resume Next\n\
               MissingObject = cell.Value\n\
               MissingObject = Err.Number\n\
             End Function\n",
        )
        .unwrap();
        let mut host = SheetHost::default();
        let value = execute_with_host(&module, "MissingObject", vec![], &mut host).unwrap();

        assert_eq!(value, Value::Integer(91));
    }

    #[test]
    fn intrinsic_type_predicates_distinguish_vba_value_categories() {
        let value = run(
            "Public Function TypePredicates() As String\n\
               Dim values As Variant\n\
               values = Array(1, 2)\n\
               TypePredicates = IsEmpty(Empty) & \"|\" & IsNull(Null) & \"|\" & IsNumeric(\"12.5\") & \"|\" & IsNumeric(\"no\") & \"|\" & IsArray(values) & \"|\" & TypeName(values) & \"|\" & VarType(values)\n\
             End Function\n",
            "TypePredicates",
            vec![],
        )
        .unwrap();

        assert_eq!(
            value,
            Value::String("True|True|True|False|True|Variant()|8204".to_string())
        );
    }

    #[test]
    fn collection_supports_keys_order_default_item_remove_and_enumeration() {
        let value = run(
            "Public Function CollectionProbe() As String\n\
               Dim items As New Collection\n\
               Dim item As Variant\n\
               Dim joined As String\n\
               items.Add \"second\", \"b\"\n\
               items.Add \"first\", \"a\", 1\n\
               items.Add \"third\", \"c\", , \"b\"\n\
               CollectionProbe = items.Count & \"|\" & items(\"a\") & \"|\" & items.Item(2) & \"|\" & items(3) & \"|\"\n\
               items.Remove \"b\"\n\
               For Each item In items\n\
                 joined = joined & item\n\
               Next\n\
               CollectionProbe = CollectionProbe & items.Count & \"|\" & joined & \"|\"\n\
               Set items = Nothing\n\
               items.Add \"reset\"\n\
               CollectionProbe = CollectionProbe & items.Count & \"|\" & items(1)\n\
             End Function\n",
            "CollectionProbe",
            vec![],
        )
        .unwrap();

        assert_eq!(
            value,
            Value::String("3|first|second|third|2|firstthird|1|reset".to_string())
        );
    }

    #[test]
    fn new_collection_and_module_as_new_reactivate_after_nothing() {
        let value = run(
            "Private shared As New Collection\n\
             Private Sub FillShared()\n\
               shared.Add 20\n\
               shared.Add 22\n\
             End Sub\n\
             Private Sub ClearShared()\n\
               Set shared = Nothing\n\
             End Sub\n\
             Public Function CollectionLifetime() As Long\n\
               Dim local As Collection\n\
               Set local = New Collection\n\
               local.Add 1\n\
               FillShared\n\
               CollectionLifetime = shared(1) + shared(2) + local.Count\n\
               ClearShared\n\
               shared.Add 5\n\
               CollectionLifetime = CollectionLifetime * 10 + shared(1)\n\
             End Function\n",
            "CollectionLifetime",
            vec![],
        )
        .unwrap();

        assert_eq!(value, Value::Integer(435));
    }

    #[test]
    fn collection_reports_duplicate_and_missing_keys_with_vba_numbers() {
        let value = run(
            "Public Function CollectionErrors() As String\n\
               Dim items As New Collection\n\
               Dim duplicate As Long\n\
               Dim missing As Long\n\
               items.Add 1, \"key\"\n\
               On Error Resume Next\n\
               items.Add 2, \"KEY\"\n\
               duplicate = Err.Number\n\
               Err.Clear\n\
               missing = items(\"absent\")\n\
               missing = Err.Number\n\
               CollectionErrors = duplicate & \"|\" & missing\n\
             End Function\n",
            "CollectionErrors",
            vec![],
        )
        .unwrap();

        assert_eq!(value, Value::String("457|5".to_string()));
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
    fn on_error_resume_next_records_and_clears_err_properties() {
        let value = run(
            "Public Function InlineErrors() As String\n\
               Dim first As Long\n\
               Dim second As Long\n\
               Dim value As Long\n\
               On Error Resume Next\n\
               value = 1 / 0\n\
               first = Err.Number\n\
               Err.Clear\n\
               value = \"not a number\" + 1\n\
               second = Err.Number\n\
               InlineErrors = first & \"|\" & second & \"|\" & Err.Description\n\
             End Function\n",
            "InlineErrors",
            vec![],
        )
        .unwrap();

        assert_eq!(
            value,
            Value::String("11|13|type mismatch converting String to number".to_string())
        );
    }

    #[test]
    fn on_error_goto_handles_an_error_raised_by_a_called_procedure() {
        let value = run(
            "Private Sub Explode()\n\
               Err.Raise 1001, \"Worker\", \"boom\"\n\
             End Sub\n\
             Public Function CatchCall() As String\n\
               On Error GoTo Failed\n\
               Explode\n\
               CatchCall = \"not reached\"\n\
               Exit Function\n\
             Failed:\n\
               CatchCall = Err.Number & \"|\" & Err.Source & \"|\" & Err.Description\n\
             End Function\n",
            "CatchCall",
            vec![],
        )
        .unwrap();

        assert_eq!(value, Value::String("1001|Worker|boom".to_string()));
    }

    #[test]
    fn resume_next_continues_after_the_failed_statement() {
        let value = run(
            "Public Function ResumeNextProbe() As Long\n\
               Dim value As Long\n\
               value = 1\n\
               On Error GoTo Failed\n\
               value = 1 / 0\n\
               value = value + 4\n\
               ResumeNextProbe = value\n\
               Exit Function\n\
             Failed:\n\
               Resume Next\n\
             End Function\n",
            "ResumeNextProbe",
            vec![],
        )
        .unwrap();

        assert_eq!(value, Value::Integer(5));
    }

    #[test]
    fn resume_retries_and_resume_label_redirects_execution() {
        let retry = run(
            "Public Function RetryProbe() As Long\n\
               Dim attempts As Long\n\
               On Error GoTo Failed\n\
               If attempts = 0 Then Err.Raise 77\n\
               RetryProbe = attempts\n\
               Exit Function\n\
             Failed:\n\
               attempts = attempts + 1\n\
               Resume\n\
             End Function\n",
            "RetryProbe",
            vec![],
        )
        .unwrap();
        assert_eq!(retry, Value::Integer(1));

        let redirected = run(
            "Public Function RedirectProbe() As Long\n\
               On Error GoTo Failed\n\
               Err.Raise 88\n\
               RedirectProbe = 1\n\
               Exit Function\n\
             Continued:\n\
               RedirectProbe = 42\n\
               Exit Function\n\
             Failed:\n\
               Resume Continued\n\
             End Function\n",
            "RedirectProbe",
            vec![],
        )
        .unwrap();
        assert_eq!(redirected, Value::Integer(42));
    }

    #[test]
    fn an_error_in_an_active_handler_unwinds_to_the_callers_handler() {
        let value = run(
            "Private Sub Inner()\n\
               On Error GoTo InnerFailed\n\
               Err.Raise 100\n\
               Exit Sub\n\
             InnerFailed:\n\
               Err.Raise 200, \"InnerHandler\", \"handler failed\"\n\
             End Sub\n\
             Public Function Outer() As String\n\
               On Error GoTo OuterFailed\n\
               Inner\n\
               Outer = \"not reached\"\n\
               Exit Function\n\
             OuterFailed:\n\
               Outer = Err.Number & \"|\" & Err.Source\n\
             End Function\n",
            "Outer",
            vec![],
        )
        .unwrap();

        assert_eq!(value, Value::String("200|InnerHandler".to_string()));
    }

    #[test]
    fn goto_computed_goto_and_gosub_return_share_the_label_dispatcher() {
        let value = run(
            "Public Function JumpProbe() As Long\n\
               Dim value As Long\n\
               value = 1\n\
               GoSub AddTen\n\
               On 2 GoTo Wrong, Finished\n\
             Wrong:\n\
               value = 999\n\
             Finished:\n\
               JumpProbe = value\n\
               Exit Function\n\
             AddTen:\n\
               value = value + 10\n\
               Return\n\
             End Function\n",
            "JumpProbe",
            vec![],
        )
        .unwrap();

        assert_eq!(value, Value::Integer(11));
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
    fn executes_vba_string_slicing_search_and_replacement_functions() {
        let value = run(
            "Option Compare Text\n\
             Public Function TextTools() As String\n\
               Dim source As String\n\
               source = \"A😀B\"\n\
               TextTools = Left(source, 3) & \"|\" & Mid(source, 2, 2) & \"|\" & Right(source, 1) & \"|\" & Len(source) & \"|\"\n\
               TextTools = TextTools & InStr(1, \"日本ABC\", \"abc\", vbUseCompareOption) & \"|\" & InStrRev(\"abAB\", \"ab\", -1, vbTextCompare) & \"|\"\n\
               TextTools = TextTools & Replace(\"xx-A-a\", \"a\", \"!\", 4, 1, vbTextCompare)\n\
             End Function\n",
            "TextTools",
            vec![],
        )
        .unwrap();

        assert_eq!(value, Value::String("A😀|😀|B|4|3|3|!-a".to_string()));
    }

    #[test]
    fn splits_joins_and_constructs_vba_strings() {
        let value = run(
            "Public Function BuildText() As String\n\
               Dim parts As Variant\n\
               Dim emptyParts As Variant\n\
               parts = Split(\"alpha|beta|gamma\", \"|\", 2)\n\
               emptyParts = Split(\"\")\n\
               BuildText = Join(parts, \"+\") & \"|\" & parts(1) & \"|\" & LBound(emptyParts) & \"|\" & UBound(emptyParts) & \"|\"\n\
               BuildText = BuildText & LTrim(\"  left\") & \"|\" & RTrim(\"right  \") & \"|\" & StrReverse(\"abc\") & \"|\"\n\
               BuildText = BuildText & String(3, \"xy\") & \"|\" & Len(Space(4)) & \"|\" & ChrW(12354) & \"|\" & AscW(\"あ\") & vbTab & \"done\"\n\
             End Function\n",
            "BuildText",
            vec![],
        )
        .unwrap();

        assert_eq!(
            value,
            Value::String(
                "alpha+beta|gamma|beta|gamma|0|-1|left|right|cba|xxx|4|あ|12354\tdone".to_string()
            )
        );
    }

    #[test]
    fn string_functions_propagate_null_and_raise_vba_error_five() {
        let value = run(
            "Public Function StringErrors() As String\n\
               Dim failure As Long\n\
               On Error Resume Next\n\
               StringErrors = Left(\"abc\", -1)\n\
               failure = Err.Number\n\
               Err.Clear\n\
               StringErrors = failure & \"|\" & IsNull(Left(Null, 1)) & \"|\" & IsNull(InStr(Null, \"x\"))\n\
             End Function\n",
            "StringErrors",
            vec![],
        )
        .unwrap();

        assert_eq!(value, Value::String("5|True|True".to_string()));
    }

    #[test]
    fn matches_vba_like_patterns_with_option_compare() {
        let value = run(
            "Option Compare Text\n\
             Public Function PatternSummary() As String\n\
               PatternSummary = (\"Invoice-42\" Like \"invoice-##\") & \"|\" & (\"ABCxyz\" Like \"[A-C]*[!0-9]\") & \"|\" & (\"fileX\" Like \"file#\") & \"|\" & (\"anything\" Like \"*\")\n\
             End Function\n",
            "PatternSummary",
            vec![],
        )
        .unwrap();

        assert_eq!(value, Value::String("True|True|False|True".to_string()));
    }

    #[test]
    fn executes_mid_lset_and_rset_string_assignments() {
        let value = run(
            "Public Function Rewrite() As String\n\
               Dim value As String\n\
               value = \"The dog jumps\"\n\
               Mid$(value, 5, 3) = \"duck\"\n\
               Rewrite = value & \"|\"\n\
               Mid(value, 5) = \"cow jumped over\"\n\
               Rewrite = Rewrite & value & \"|\"\n\
               value = \"0123456789\"\n\
               LSet value = \"<-Left\"\n\
               Rewrite = Rewrite & \"[\" & value & \"]|\"\n\
               RSet value = \"Right->\"\n\
               Rewrite = Rewrite & \"[\" & value & \"]\"\n\
             End Function\n",
            "Rewrite",
            vec![],
        )
        .unwrap();

        assert_eq!(
            value,
            Value::String("The duc jumps|The cow jumpe|[<-Left    ]|[   Right->]".to_string())
        );
    }

    #[test]
    fn preserves_fixed_length_strings_through_assignment_and_byref() {
        let value = run(
            "Private moduleText As String * 5\n\
             Private Sub Expand(ByRef value As String)\n\
               value = \"abcdefgh\"\n\
             End Sub\n\
             Public Function FixedStrings() As String\n\
               Dim localText As String * 6\n\
               Dim entries(1 To 2) As String * 4\n\
               localText = \"ab\"\n\
               Mid(localText, 3, 2) = \"XYZ\"\n\
               RSet localText = \"Q\"\n\
               Expand localText\n\
               entries(1) = \"x\"\n\
               Expand entries(2)\n\
               moduleText = \"module-wide\"\n\
               FixedStrings = \"[\" & localText & \"]|[\" & entries(1) & \"]|[\" & entries(2) & \"]|[\" & moduleText & \"]\"\n\
             End Function\n",
            "FixedStrings",
            vec![],
        )
        .unwrap();

        assert_eq!(
            value,
            Value::String("[abcdef]|[x   ]|[abcd]|[modul]".to_string())
        );
    }

    #[test]
    fn mid_statement_reports_vba_error_five_for_an_invalid_start() {
        let value = run(
            "Public Function MidFailure() As Long\n\
               Dim value As String\n\
               value = \"abc\"\n\
               On Error Resume Next\n\
               Mid(value, 4) = \"x\"\n\
               MidFailure = Err.Number\n\
             End Function\n",
            "MidFailure",
            vec![],
        )
        .unwrap();

        assert_eq!(value, Value::Integer(5));
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
    fn indexes_multidimensional_arrays_and_reports_each_dimension_bound() {
        let value = run(
            "Private Sub Increment(ByRef value As Long)\n\
               value = value + 1\n\
             End Sub\n\
             Public Function Matrix() As String\n\
               Dim values(0 To 1, 2 To 4) As Long\n\
               values(0, 2) = 7\n\
               values(1, 4) = 9\n\
               Increment values(1, 4)\n\
               Matrix = values(0, 2) & \"|\" & values(1, 4) & \"|\" & LBound(values, 1) & \"|\" & UBound(values, 1) & \"|\" & LBound(values, 2) & \"|\" & UBound(values, 2)\n\
             End Function\n",
            "Matrix",
            vec![],
        )
        .unwrap();
        assert_eq!(value, Value::String("7|10|0|1|2|4".to_string()));
    }

    #[test]
    fn redim_preserve_resizes_only_the_last_dimension() {
        let value = run(
            "Public Function ResizeMatrix() As Long\n\
               Dim values() As Long\n\
               ReDim values(1 To 2, 3 To 4)\n\
               values(1, 3) = 10\n\
               values(2, 4) = 20\n\
               ReDim Preserve values(1 To 2, 3 To 5)\n\
               values(2, 5) = 30\n\
               ResizeMatrix = values(1, 3) + values(2, 4) + values(2, 5)\n\
             End Function\n",
            "ResizeMatrix",
            vec![],
        )
        .unwrap();
        assert_eq!(value, Value::Integer(60));

        let failure = run(
            "Public Sub Broken()\n\
               Dim values() As Long\n\
               ReDim values(1 To 2, 1 To 2)\n\
               ReDim Preserve values(1 To 3, 1 To 2)\n\
             End Sub\n",
            "Broken",
            vec![],
        )
        .unwrap_err();
        assert_eq!(failure.kind, RuntimeErrorKind::SubscriptOutOfRange);
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
               Dim values() As Long\n\
               ReDim values(1 To 2)\n\
               ReDim Preserve values(0 To 3)\n\
             End Sub\n",
            "Broken",
            vec![],
        )
        .unwrap_err();
        assert_eq!(failure.kind, RuntimeErrorKind::SubscriptOutOfRange);
        assert_eq!(failure.line, Some(4));
    }

    #[test]
    fn erase_resets_fixed_arrays_and_rejects_redim_with_error_ten() {
        let value = run(
            "Public Function FixedErase() As String\n\
               Dim numbers(1 To 2) As Long\n\
               Dim labels(0 To 1) As String * 3\n\
               Dim objects(1 To 1) As Object\n\
               Dim failure As Long\n\
               numbers(1) = 10\n\
               numbers(2) = 20\n\
               labels(0) = \"abc\"\n\
               Erase numbers, labels, objects\n\
               On Error Resume Next\n\
               ReDim numbers(1 To 3)\n\
               failure = Err.Number\n\
               On Error GoTo 0\n\
               FixedErase = numbers(1) & \"|\" & numbers(2) & \"|\" & LBound(numbers) & \"|\" & UBound(numbers) & \"|[\" & labels(0) & \"]|\" & (objects(1) Is Nothing) & \"|\" & failure\n\
             End Function\n",
            "FixedErase",
            vec![],
        )
        .unwrap();

        assert_eq!(value, Value::String("0|0|1|2|[   ]|True|10".to_string()));
    }

    #[test]
    fn erase_deallocates_dynamic_and_variant_arrays_until_redim() {
        let value = run(
            "Public Function DynamicErase() As String\n\
               Dim dynamicValues() As Long\n\
               Dim fixedValues(1 To 2) As Long\n\
               Dim variantValues As Variant\n\
               Dim firstFailure As Long\n\
               Dim secondFailure As Long\n\
               Dim ignored As Long\n\
               ReDim dynamicValues(2 To 3)\n\
               dynamicValues(2) = 20\n\
               fixedValues(1) = 7\n\
               variantValues = fixedValues\n\
               Erase dynamicValues, variantValues\n\
               On Error Resume Next\n\
               ignored = LBound(dynamicValues)\n\
               firstFailure = Err.Number\n\
               Err.Clear\n\
               ignored = UBound(variantValues)\n\
               secondFailure = Err.Number\n\
               On Error GoTo 0\n\
               ReDim dynamicValues(5 To 5)\n\
               ReDim variantValues(1 To 1)\n\
               dynamicValues(5) = 30\n\
               variantValues(1) = 40\n\
               DynamicErase = firstFailure & \"|\" & secondFailure & \"|\" & dynamicValues(5) & \"|\" & variantValues(1)\n\
             End Function\n",
            "DynamicErase",
            vec![],
        )
        .unwrap();

        assert_eq!(value, Value::String("9|9|30|40".to_string()));
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
