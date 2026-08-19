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
    /// A Variant of subtype Error, normally created by `CVErr`.
    Error(i64),
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

    fn call_named(
        &mut self,
        receiver: Option<&ObjectRef>,
        name: &str,
        args: &[Value],
        _argument_names: &[Option<String>],
    ) -> Result<Option<Value>, String> {
        self.call(receiver, name, args)
    }

    fn get(&mut self, receiver: &ObjectRef, name: &str) -> Result<Option<Value>, String>;

    fn set(&mut self, receiver: &ObjectRef, name: &str, value: Value) -> Result<bool, String>;

    fn set_indexed(
        &mut self,
        _receiver: &ObjectRef,
        _name: &str,
        _args: &[Value],
        _value: Value,
    ) -> Result<bool, String> {
        Ok(false)
    }

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
    random_state: u32,
    random_entropy: u64,
    current_time: f64,
}

struct Frame {
    procedure_name: String,
    source_name: String,
    values: BTreeMap<String, ValueSlot>,
    constants: BTreeSet<String>,
    auto_new: BTreeMap<String, String>,
    fixed_strings: BTreeMap<String, usize>,
    /// The type each variable was declared with, so an assignment can
    /// narrow to it the way VBA does.
    declared: BTreeMap<String, String>,
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
    Dictionary(DictionaryObject),
}

struct CollectionEntry {
    value: Value,
    key: Option<String>,
}

struct DictionaryObject {
    entries: Vec<DictionaryEntry>,
    compare_mode: i64,
    compare_text: bool,
}

struct DictionaryEntry {
    key: Value,
    value: Value,
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
            random_state: 327_680,
            random_entropy: 327_680,
            current_time: default_current_time(),
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

    /// Supplies entropy for `Randomize` calls that omit their numeric seed.
    pub fn with_random_seed(mut self, seed: u64) -> Self {
        self.random_entropy = seed;
        self
    }

    /// Supplies the current local date and time as an OLE Automation serial.
    pub fn with_current_time(mut self, serial: f64) -> Self {
        if serial.is_finite() {
            self.current_time = serial;
        }
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
            declared: BTreeMap::new(),
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
            let narrowed = !param.is_array
                && param.mode == ParamMode::ByVal
                && !param.type_name.is_new;
            if narrowed {
                let taken = value.borrow().clone();
                *value.borrow_mut() =
                    coerce_declared(taken, &param.type_name.name, procedure.span.line)?;
            }
            frame.values.insert(key(&param.name), value);
            if narrowed {
                frame
                    .declared
                    .insert(key(&param.name), param.type_name.name.clone());
            }
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
            match procedure.return_type.as_ref() {
                Some(return_type) => {
                    coerce_declared(value, &return_type.name, procedure.span.line)?
                }
                None => value,
            }
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
                declared: BTreeMap::new(),
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
                frame.error_state = ErrorState::default();
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
            Statement::Call { target, span, .. }
                if matches!(target, Expr::Index { target, .. }
                    if expr_name(target).is_some_and(|name| name.eq_ignore_ascii_case("error"))) =>
            {
                self.exec_error_statement(target, frame, span.line)?;
                Ok(Flow::Continue)
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
            if variable.array_bounds.is_none() && !variable.type_name.is_new {
                frame
                    .declared
                    .insert(name.clone(), variable.type_name.name.clone());
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
        let (kind, object) = if type_name.eq_ignore_ascii_case("collection") {
            ("Collection", InternalObject::Collection(Vec::new()))
        } else if type_name.eq_ignore_ascii_case("dictionary")
            || type_name.eq_ignore_ascii_case("scripting.dictionary")
        {
            (
                "Dictionary",
                InternalObject::Dictionary(DictionaryObject {
                    entries: Vec::new(),
                    compare_mode: 0,
                    compare_text: false,
                }),
            )
        } else {
            return Err(error(
                RuntimeErrorKind::Unsupported,
                format!("New is not available for VBA type: {type_name}"),
                Some(line),
            ));
        };
        let handle = self.next_internal_handle;
        self.next_internal_handle = self.next_internal_handle.checked_add(1).ok_or_else(|| {
            error(
                RuntimeErrorKind::Overflow,
                "VBA internal object handle overflow",
                Some(line),
            )
        })?;
        self.internal_objects.insert(handle, object);
        Ok(Value::Object(ObjectRef {
            handle,
            kind: kind.to_string(),
        }))
    }

    fn create_object(&mut self, args: &[Value], line: Option<u32>) -> Result<Value, RuntimeError> {
        if args.len() != 1 {
            return Err(error(
                RuntimeErrorKind::ArgumentCount,
                format!(
                    "CreateObject expects one String class name, received {} argument(s)",
                    args.len()
                ),
                line,
            ));
        }
        let Value::String(type_name) = &args[0] else {
            return Err(error(
                RuntimeErrorKind::TypeMismatch,
                "CreateObject class name must be a String",
                line,
            ));
        };
        self.new_object(type_name, line.unwrap_or(0))
            .map_err(|failure| {
                if failure.kind == RuntimeErrorKind::Unsupported {
                    RuntimeError {
                        kind: RuntimeErrorKind::UserDefined,
                        message: format!("ActiveX component cannot create object: {type_name}"),
                        line,
                        vba_number: Some(429),
                        vba_source: Some("VBA".to_string()),
                    }
                } else {
                    failure
                }
            })
    }

    fn is_constant(&self, frame: &Frame, name: &str) -> bool {
        let name = key(name);
        if frame.values.contains_key(&name) {
            frame.constants.contains(&name)
        } else {
            self.module_constants.contains(&name)
        }
    }

    fn declared_type(&self, frame: &Frame, name: &str) -> Option<String> {
        frame.declared.get(&key(name)).cloned()
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
                let mut value = match self.declared_type(frame, name) {
                    Some(declared) => coerce_declared(value, &declared, line)?,
                    None => value,
                };
                value = match self.fixed_string_width(frame, name) {
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
                if let Some((receiver, member)) = self.indexed_object_target(target, frame, line)? {
                    let mut values = Vec::with_capacity(args.len());
                    for argument in args {
                        if argument.name.is_some() {
                            return Err(error(
                                RuntimeErrorKind::ArgumentCount,
                                "indexed property arguments cannot be named",
                                Some(line),
                            ));
                        }
                        values.push(match &argument.value {
                            Some(value) => self.eval_expr(value, frame)?,
                            None => Value::Missing,
                        });
                    }
                    return match self.host_set_indexed(&receiver, &member, &values, value, line)? {
                        true => Ok(()),
                        false => Err(error(
                            RuntimeErrorKind::Unsupported,
                            format!(
                                "indexed property is not writable: {}.{member}",
                                receiver.kind
                            ),
                            Some(line),
                        )),
                    };
                }
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

    fn indexed_object_target(
        &mut self,
        target: &Expr,
        frame: &mut Frame,
        line: u32,
    ) -> Result<Option<(ObjectRef, String)>, RuntimeError> {
        match target {
            Expr::Ident(name, _) | Expr::TypedIdent { name, .. } => {
                match self.read_variable(frame, name, line)? {
                    Some(Value::Object(receiver)) => Ok(Some((receiver, "Item".to_string()))),
                    Some(Value::Nothing) => Err(error(
                        RuntimeErrorKind::ObjectVariableNotSet,
                        "object variable or With block variable not set",
                        Some(line),
                    )),
                    _ => Ok(None),
                }
            }
            Expr::Member { object, name, .. } => {
                Ok(Some((self.eval_object(object, frame, line)?, name.clone())))
            }
            Expr::WithMember(name, _) | Expr::WithBangMember(name, _) => {
                Ok(Some((current_with_object(frame, line)?, name.clone())))
            }
            _ => Ok(None),
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
                    return Ok(Value::Integer(frame.error_state.number));
                }
                if name.eq_ignore_ascii_case("erl") {
                    return Ok(Value::Integer(frame.error_state.line.unwrap_or(0) as i64));
                }
                if ["date", "doevents", "now", "rnd", "time", "timer"]
                    .iter()
                    .any(|builtin| name.eq_ignore_ascii_case(builtin))
                {
                    return self.call_named(name, Vec::new(), &[], Some(span.line), frame);
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
                let value = self.scalar_operand(value, span.line)?;
                unary(*op, value).map_err(|message| {
                    error(RuntimeErrorKind::TypeMismatch, message, Some(span.line))
                })
            }
            Expr::Binary { op, lhs, rhs, span } => {
                let lhs = self.eval_expr(lhs, frame)?;
                let rhs = self.eval_expr(rhs, frame)?;
                // `Is` asks about identity, so it keeps the objects themselves.
                let (lhs, rhs) = if *op == BinaryOp::Is {
                    (lhs, rhs)
                } else {
                    (
                        self.scalar_operand(lhs, span.line)?,
                        self.scalar_operand(rhs, span.line)?,
                    )
                };
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
                        || type_name
                            .rsplit('.')
                            .next()
                            .is_some_and(|name| object.kind.eq_ignore_ascii_case(name))
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
                let mut argument_names = Vec::with_capacity(args.len());
                for argument in args {
                    argument_names.push(argument.name.clone());
                    values.push(match argument.value.as_ref() {
                        Some(value) => self.eval_expr(value, frame)?,
                        None => Value::Missing,
                    });
                }
                match target.as_ref() {
                    Expr::Ident(name, _) | Expr::TypedIdent { name, .. } => {
                        self.call_named(name, values, &argument_names, Some(span.line), frame)
                    }
                    Expr::Member { object, name, .. } => {
                        let receiver = self.eval_object(object, frame, span.line)?;
                        if is_err_object(&receiver) {
                            return err_call(frame, name, &values, span.line);
                        }
                        self.host_call_named(
                            Some(&receiver),
                            name,
                            &values,
                            &argument_names,
                            span.line,
                        )?
                        .ok_or_else(|| {
                            error(
                                RuntimeErrorKind::Unsupported,
                                format!("host method is not available: {}.{name}", receiver.kind),
                                Some(span.line),
                            )
                        })
                    }
                    Expr::WithMember(name, _) | Expr::WithBangMember(name, _) => {
                        let receiver = current_with_object(frame, span.line)?;
                        self.host_call_named(
                            Some(&receiver),
                            name,
                            &values,
                            &argument_names,
                            span.line,
                        )?
                        .ok_or_else(|| {
                            error(
                                RuntimeErrorKind::Unsupported,
                                format!("host method is not available: {}.{name}", receiver.kind),
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
                self.call_named(name, Vec::new(), &[], Some(span.line), frame)
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
        argument_names: &[Option<String>],
        line: Option<u32>,
        frame: &Frame,
    ) -> Result<Value, RuntimeError> {
        if self.module.items.iter().any(
            |item| matches!(item, ModuleItem::Procedure(p) if p.name.eq_ignore_ascii_case(name)),
        ) {
            return self.call_procedure(name, args, line);
        }
        if name.eq_ignore_ascii_case("createobject") {
            return self.create_object(&args, line);
        }
        if name.eq_ignore_ascii_case("rnd") {
            return self.call_rnd(&args, line);
        }
        if name.eq_ignore_ascii_case("randomize") {
            return self.call_randomize(&args, line);
        }
        if ["date", "now", "time", "timer"]
            .iter()
            .any(|builtin| name.eq_ignore_ascii_case(builtin))
        {
            return self.call_current_time(name, &args, line);
        }
        if name.eq_ignore_ascii_case("error") {
            return call_error_builtin(&args, frame, line);
        }
        if name.eq_ignore_ascii_case("err") || name.eq_ignore_ascii_case("erl") {
            if !args.is_empty() {
                return Err(error(
                    RuntimeErrorKind::ArgumentCount,
                    format!("{name} expects no arguments, received {}", args.len()),
                    line,
                ));
            }
            return Ok(Value::Integer(if name.eq_ignore_ascii_case("err") {
                frame.error_state.number
            } else {
                frame.error_state.line.unwrap_or(0) as i64
            }));
        }
        let read_args = if builtin_reads_values(name)
            && args.iter().any(|value| matches!(value, Value::Object(_)))
        {
            let mut read = Vec::with_capacity(args.len());
            for value in &args {
                read.push(self.read_argument(value, line.unwrap_or(0))?);
            }
            Some(read)
        } else {
            None
        };
        if let Some(result) = call_builtin(
            name,
            read_args.as_deref().unwrap_or(&args),
            line,
            self.option_compare_text(),
        ) {
            return result;
        }
        let host_value = if argument_names.is_empty() {
            self.host_call(None, name, &args, line.unwrap_or(0))?
        } else {
            self.host_call_named(None, name, &args, argument_names, line.unwrap_or(0))?
        };
        if let Some(value) = host_value {
            return Ok(value);
        }
        self.call_procedure(name, args, line)
    }

    fn exec_error_statement(
        &mut self,
        target: &Expr,
        frame: &mut Frame,
        line: u32,
    ) -> Result<(), RuntimeError> {
        let Expr::Index { args, .. } = target else {
            return Err(error(
                RuntimeErrorKind::ArgumentCount,
                "Error statement expects 1 argument",
                Some(line),
            ));
        };
        if args.len() != 1 || args[0].value.is_none() {
            return Err(error(
                RuntimeErrorKind::ArgumentCount,
                format!(
                    "Error statement expects 1 argument, received {}",
                    args.len()
                ),
                Some(line),
            ));
        }
        let value = self.eval_expr(args[0].value.as_ref().unwrap(), frame)?;
        let number = integer_argument(&value, Some(line))?;
        if !(1..=65_535).contains(&number) {
            return Err(invalid_procedure_call(
                format!("invalid Error statement number: {number}"),
                Some(line),
            ));
        }
        let description = vba_error_description(number);
        Err(raised_error(
            number,
            frame.source_name.clone(),
            if description == "Application-defined or object-defined error" {
                String::new()
            } else {
                description.to_string()
            },
            line,
        ))
    }

    fn call_rnd(&mut self, args: &[Value], line: Option<u32>) -> Result<Value, RuntimeError> {
        if args.len() > 1 {
            return Err(error(
                RuntimeErrorKind::ArgumentCount,
                format!("Rnd expects 0 or 1 arguments, received {}", args.len()),
                line,
            ));
        }
        let argument = match args.first() {
            None | Some(Value::Missing) => 1.0,
            Some(value) => number(value)
                .map_err(|message| error(RuntimeErrorKind::TypeMismatch, message, line))?,
        };
        if argument < 0.0 {
            self.random_state = (argument as f32).to_bits() & 0x00ff_ffff;
        }
        if argument != 0.0 {
            self.random_state = self
                .random_state
                .wrapping_mul(1_140_671_485)
                .wrapping_add(12_820_163)
                & 0x00ff_ffff;
        }
        Ok(Value::Double(f64::from(self.random_state) / 16_777_216.0))
    }

    fn call_randomize(&mut self, args: &[Value], line: Option<u32>) -> Result<Value, RuntimeError> {
        if args.len() > 1 {
            return Err(error(
                RuntimeErrorKind::ArgumentCount,
                format!(
                    "Randomize expects 0 or 1 arguments, received {}",
                    args.len()
                ),
                line,
            ));
        }
        let bits = match args.first() {
            None | Some(Value::Missing) => {
                self.random_entropy = self
                    .random_entropy
                    .wrapping_mul(6_364_136_223_846_793_005)
                    .wrapping_add(1_442_695_040_888_963_407);
                self.random_entropy
            }
            Some(value) => number(value)
                .map_err(|message| error(RuntimeErrorKind::TypeMismatch, message, line))?
                .to_bits(),
        };
        let folded = (bits as u32) ^ (bits >> 32) as u32;
        self.random_state = ((folded & 0xffff) ^ (folded >> 16)) << 8 | (self.random_state & 0xff);
        self.random_state &= 0x00ff_ffff;
        Ok(Value::Empty)
    }

    fn call_current_time(
        &self,
        name: &str,
        args: &[Value],
        line: Option<u32>,
    ) -> Result<Value, RuntimeError> {
        if !args.is_empty() {
            return Err(error(
                RuntimeErrorKind::ArgumentCount,
                format!("{name} expects no arguments, received {}", args.len()),
                line,
            ));
        }
        let value = match name.to_ascii_lowercase().as_str() {
            "now" => self.current_time,
            "date" => self.current_time.floor(),
            "time" => self.current_time.rem_euclid(1.0),
            "timer" => self.current_time.rem_euclid(1.0) * 86_400.0,
            _ => unreachable!(),
        };
        Ok(Value::Double(value))
    }

    fn eval_object(
        &mut self,
        expr: &Expr,
        frame: &mut Frame,
        line: u32,
    ) -> Result<ObjectRef, RuntimeError> {
        if expr_name(expr).is_some_and(|name| name.eq_ignore_ascii_case("err")) {
            return Ok(ObjectRef {
                handle: u64::MAX,
                kind: "Err".to_string(),
            });
        }
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

    /// An object used where VBA wants a value stands for its default member, so
    /// `Range("A1") + 1` reads the cell rather than failing. Object contexts —
    /// `Is`, `TypeName`, `IsObject`, passing an argument — keep the object.
    fn scalar_operand(&mut self, value: Value, line: u32) -> Result<Value, RuntimeError> {
        match value {
            Value::Object(_) => self.let_value(value, line),
            value => Ok(value),
        }
    }

    /// Reads a builtin's argument as a value. An object with a default member
    /// stands for it, but one without keeps its place rather than failing:
    /// `VarType` answers 9 for a worksheet instead of raising.
    fn read_argument(&mut self, value: &Value, line: u32) -> Result<Value, RuntimeError> {
        let Value::Object(object) = value else {
            return Ok(value.clone());
        };
        Ok(self
            .host_get(object, "Value", line)?
            .unwrap_or_else(|| value.clone()))
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

    fn host_call_named(
        &mut self,
        receiver: Option<&ObjectRef>,
        name: &str,
        args: &[Value],
        argument_names: &[Option<String>],
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
        host.call_named(receiver, name, args, argument_names)
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
                (InternalObject::Dictionary(dictionary), "count") => {
                    Ok(Some(Value::Integer(dictionary.entries.len() as i64)))
                }
                (InternalObject::Dictionary(dictionary), "comparemode") => {
                    Ok(Some(Value::Integer(dictionary.compare_mode)))
                }
                (InternalObject::Dictionary(dictionary), "keys") => Ok(Some(dictionary_array(
                    dictionary
                        .entries
                        .iter()
                        .map(|entry| entry.key.clone())
                        .collect(),
                ))),
                (InternalObject::Dictionary(dictionary), "items") => Ok(Some(dictionary_array(
                    dictionary
                        .entries
                        .iter()
                        .map(|entry| entry.value.clone())
                        .collect(),
                ))),
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
        let option_compare_text = self.option_compare_text();
        if let Some(object) = self.internal_objects.get_mut(&receiver.handle) {
            return match object {
                InternalObject::Dictionary(dictionary)
                    if name.eq_ignore_ascii_case("comparemode") =>
                {
                    if !dictionary.entries.is_empty() {
                        return Err(invalid_procedure_call(
                            "CompareMode cannot be changed after adding Dictionary entries"
                                .to_string(),
                            Some(line),
                        ));
                    }
                    let mode = integer_argument(&value, Some(line))?;
                    if !matches!(mode, -1..=1) {
                        return Err(invalid_procedure_call(
                            format!("unsupported Dictionary CompareMode: {mode}"),
                            Some(line),
                        ));
                    }
                    dictionary.compare_mode = mode;
                    dictionary.compare_text = mode == 1 || (mode == -1 && option_compare_text);
                    Ok(true)
                }
                _ => Ok(false),
            };
        }
        let Some(host) = self.host.as_deref_mut() else {
            return Ok(false);
        };
        host.set(receiver, name, value)
            .map_err(|message| error(RuntimeErrorKind::Host, message, Some(line)))
    }

    fn host_set_indexed(
        &mut self,
        receiver: &ObjectRef,
        name: &str,
        args: &[Value],
        value: Value,
        line: u32,
    ) -> Result<bool, RuntimeError> {
        if self.internal_objects.contains_key(&receiver.handle) {
            return self.internal_set_indexed(receiver, name, args, value, line);
        }
        let Some(host) = self.host.as_deref_mut() else {
            return Ok(false);
        };
        host.set_indexed(receiver, name, args, value)
            .map_err(|message| error(RuntimeErrorKind::Host, message, Some(line)))
    }

    fn host_enumerate(
        &mut self,
        receiver: &ObjectRef,
        line: u32,
    ) -> Result<Option<Vec<Value>>, RuntimeError> {
        if let Some(object) = self.internal_objects.get(&receiver.handle) {
            return Ok(Some(match object {
                InternalObject::Collection(entries) => {
                    entries.iter().map(|entry| entry.value.clone()).collect()
                }
                InternalObject::Dictionary(dictionary) => dictionary
                    .entries
                    .iter()
                    .map(|entry| entry.key.clone())
                    .collect(),
            }));
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
            InternalObject::Dictionary(dictionary) if name.eq_ignore_ascii_case("add") => {
                if args.len() != 2 {
                    return Err(error(
                        RuntimeErrorKind::ArgumentCount,
                        format!(
                            "Dictionary.Add expects 2 arguments, received {}",
                            args.len()
                        ),
                        Some(line),
                    ));
                }
                validate_dictionary_key(&args[0], line)?;
                if dictionary_position(dictionary, &args[0]).is_some() {
                    return Err(dictionary_duplicate_key_error(line));
                }
                dictionary.entries.push(DictionaryEntry {
                    key: args[0].clone(),
                    value: args[1].clone(),
                });
                Ok(Value::Empty)
            }
            InternalObject::Dictionary(dictionary) if name.eq_ignore_ascii_case("exists") => {
                if args.len() != 1 {
                    return Err(error(
                        RuntimeErrorKind::ArgumentCount,
                        format!(
                            "Dictionary.Exists expects 1 argument, received {}",
                            args.len()
                        ),
                        Some(line),
                    ));
                }
                validate_dictionary_key(&args[0], line)?;
                Ok(Value::Boolean(
                    dictionary_position(dictionary, &args[0]).is_some(),
                ))
            }
            InternalObject::Dictionary(dictionary) if name.eq_ignore_ascii_case("item") => {
                if args.len() != 1 {
                    return Err(error(
                        RuntimeErrorKind::ArgumentCount,
                        format!(
                            "Dictionary.Item expects 1 argument, received {}",
                            args.len()
                        ),
                        Some(line),
                    ));
                }
                validate_dictionary_key(&args[0], line)?;
                if let Some(index) = dictionary_position(dictionary, &args[0]) {
                    Ok(dictionary.entries[index].value.clone())
                } else {
                    dictionary.entries.push(DictionaryEntry {
                        key: args[0].clone(),
                        value: Value::Empty,
                    });
                    Ok(Value::Empty)
                }
            }
            InternalObject::Dictionary(dictionary) if name.eq_ignore_ascii_case("remove") => {
                if args.len() != 1 {
                    return Err(error(
                        RuntimeErrorKind::ArgumentCount,
                        format!(
                            "Dictionary.Remove expects 1 argument, received {}",
                            args.len()
                        ),
                        Some(line),
                    ));
                }
                validate_dictionary_key(&args[0], line)?;
                let index = dictionary_position(dictionary, &args[0])
                    .ok_or_else(|| dictionary_missing_key_error(line))?;
                dictionary.entries.remove(index);
                Ok(Value::Empty)
            }
            InternalObject::Dictionary(dictionary) if name.eq_ignore_ascii_case("removeall") => {
                if !args.is_empty() {
                    return Err(error(
                        RuntimeErrorKind::ArgumentCount,
                        format!(
                            "Dictionary.RemoveAll expects no arguments, received {}",
                            args.len()
                        ),
                        Some(line),
                    ));
                }
                dictionary.entries.clear();
                Ok(Value::Empty)
            }
            InternalObject::Dictionary(dictionary)
                if name.eq_ignore_ascii_case("keys") || name.eq_ignore_ascii_case("items") =>
            {
                if !args.is_empty() {
                    return Err(error(
                        RuntimeErrorKind::ArgumentCount,
                        format!("Dictionary.{name} expects no arguments"),
                        Some(line),
                    ));
                }
                let values = if name.eq_ignore_ascii_case("keys") {
                    dictionary
                        .entries
                        .iter()
                        .map(|entry| entry.key.clone())
                        .collect()
                } else {
                    dictionary
                        .entries
                        .iter()
                        .map(|entry| entry.value.clone())
                        .collect()
                };
                Ok(dictionary_array(values))
            }
            InternalObject::Dictionary(_) => Err(error(
                RuntimeErrorKind::Unsupported,
                format!("Dictionary method is not available: {name}"),
                Some(line),
            )),
        }
    }

    fn internal_set_indexed(
        &mut self,
        receiver: &ObjectRef,
        name: &str,
        args: &[Value],
        value: Value,
        line: u32,
    ) -> Result<bool, RuntimeError> {
        let Some(InternalObject::Dictionary(dictionary)) =
            self.internal_objects.get_mut(&receiver.handle)
        else {
            return Ok(false);
        };
        if !name.eq_ignore_ascii_case("item") && !name.eq_ignore_ascii_case("key") {
            return Ok(false);
        }
        if args.len() != 1 {
            return Err(error(
                RuntimeErrorKind::ArgumentCount,
                format!(
                    "Dictionary.{name} expects 1 argument, received {}",
                    args.len()
                ),
                Some(line),
            ));
        }
        validate_dictionary_key(&args[0], line)?;
        if name.eq_ignore_ascii_case("key") {
            validate_dictionary_key(&value, line)?;
            let index = dictionary_position(dictionary, &args[0])
                .ok_or_else(|| dictionary_missing_key_error(line))?;
            if dictionary_position(dictionary, &value).is_some_and(|found| found != index) {
                return Err(dictionary_duplicate_key_error(line));
            }
            dictionary.entries[index].key = value;
            return Ok(true);
        }
        if let Some(index) = dictionary_position(dictionary, &args[0]) {
            dictionary.entries[index].value = value;
        } else {
            dictionary.entries.push(DictionaryEntry {
                key: args[0].clone(),
                value,
            });
        }
        Ok(true)
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
        declared: BTreeMap::new(),
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

fn call_error_builtin(
    args: &[Value],
    frame: &Frame,
    line: Option<u32>,
) -> Result<Value, RuntimeError> {
    if args.len() > 1 {
        return Err(error(
            RuntimeErrorKind::ArgumentCount,
            format!("Error expects 0 or 1 arguments, received {}", args.len()),
            line,
        ));
    }
    let Some(value) = args.first() else {
        return Ok(Value::String(frame.error_state.description.clone()));
    };
    let number = integer_argument(value, line)?;
    if !(0..=65_535).contains(&number) {
        return Err(invalid_procedure_call(
            format!("invalid Error number: {number}"),
            line,
        ));
    }
    Ok(Value::String(vba_error_description(number).to_string()))
}

fn vba_error_description(number: i64) -> &'static str {
    match number {
        0 => "",
        5 => "Invalid procedure call or argument",
        6 => "Overflow",
        7 => "Out of memory",
        9 => "Subscript out of range",
        10 => "This array is fixed or temporarily locked",
        11 => "Division by zero",
        13 => "Type mismatch",
        14 => "Out of string space",
        28 => "Out of stack space",
        35 => "Sub or Function not defined",
        48 => "Error in loading DLL",
        52 => "Bad file name or number",
        53 => "File not found",
        54 => "Bad file mode",
        55 => "File already open",
        58 => "File already exists",
        61 => "Disk full",
        62 => "Input past end of file",
        67 => "Too many files",
        68 => "Device unavailable",
        70 => "Permission denied",
        71 => "Disk not ready",
        75 => "Path/File access error",
        76 => "Path not found",
        91 => "Object variable or With block variable not set",
        92 => "For loop not initialized",
        93 => "Invalid pattern string",
        94 => "Invalid use of Null",
        424 => "Object required",
        429 => "ActiveX component can't create object",
        430 => "Class does not support Automation or does not support expected interface",
        432 => "File name or class name not found during Automation operation",
        438 => "Object doesn't support this property or method",
        440 => "Automation error",
        445 => "Object doesn't support this action",
        446 => "Object doesn't support named arguments",
        447 => "Object doesn't support current locale setting",
        448 => "Named argument not found",
        449 => "Argument not optional",
        450 => "Wrong number of arguments or invalid property assignment",
        451 => {
            "Property let procedure not defined and property get procedure did not return an object"
        }
        453 => "Specified DLL function not found",
        457 => "This key is already associated with an element of this collection",
        458 => "Variable uses an Automation type not supported in Visual Basic",
        _ => "Application-defined or object-defined error",
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
            .unwrap_or_else(|| vba_error_description(number).to_string());
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

fn dictionary_array(values: Vec<Value>) -> Value {
    Value::Array(ArrayValue {
        dimensions: vec![ArrayDimension {
            lower_bound: 0,
            length: values.len(),
        }],
        values,
        element_default: Box::new(Value::Empty),
        resizable: true,
    })
}

fn validate_dictionary_key(key: &Value, line: u32) -> Result<(), RuntimeError> {
    if matches!(
        key,
        Value::Array(_) | Value::Null | Value::Missing | Value::Nothing | Value::Object(_)
    ) {
        Err(error(
            RuntimeErrorKind::TypeMismatch,
            "Dictionary key must be a non-Null scalar value",
            Some(line),
        ))
    } else {
        Ok(())
    }
}

fn dictionary_position(dictionary: &DictionaryObject, key: &Value) -> Option<usize> {
    dictionary
        .entries
        .iter()
        .position(|entry| dictionary_keys_equal(&entry.key, key, dictionary.compare_text))
}

fn dictionary_keys_equal(left: &Value, right: &Value, text_compare: bool) -> bool {
    match (left, right) {
        (Value::String(left), Value::String(right)) if text_compare => {
            left.to_lowercase() == right.to_lowercase()
        }
        (Value::String(left), Value::String(right)) => left == right,
        (Value::Integer(left), Value::Integer(right)) => left == right,
        (Value::Double(left), Value::Double(right)) => left == right,
        (Value::Integer(left), Value::Double(right)) => *left as f64 == *right,
        (Value::Double(left), Value::Integer(right)) => *left == *right as f64,
        (Value::Boolean(left), Value::Boolean(right)) => left == right,
        (Value::Empty, Value::Empty) => true,
        _ => false,
    }
}

fn dictionary_duplicate_key_error(line: u32) -> RuntimeError {
    raised_error(
        457,
        "Scripting.Dictionary".to_string(),
        "this key is already associated with an element of this collection".to_string(),
        line,
    )
}

fn dictionary_missing_key_error(line: u32) -> RuntimeError {
    raised_error(
        32_811,
        "Scripting.Dictionary".to_string(),
        "element not found".to_string(),
        line,
    )
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

/// Whether a builtin reads its arguments as values, in which case an object
/// argument stands for its default member — `Len(Range("A1"))` measures the
/// cell's text. Measured across the conversion, string, maths, date and
/// type-predicate builtins, which all read values; `VarType` reads a value too,
/// answering `5` for a numeric cell rather than `9` for the object.
///
/// The exceptions keep the object. `TypeName`, `IsObject` and `IsMissing` ask
/// about the argument itself. `IIf`, `Array` and `Choose` hand Variants
/// straight back, leaving resolution to whatever consumes the result. The
/// array builtins want an array rather than a scalar, and probing what Excel
/// does when handed a range instead wedges it, so they are left alone.
fn builtin_reads_values(name: &str) -> bool {
    !matches!(
        name.to_ascii_lowercase().as_str(),
        "typename"
            | "isobject"
            | "ismissing"
            | "iif"
            | "array"
            | "choose"
            | "lbound"
            | "ubound"
            | "join"
            | "filter"
            | "split"
    )
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
            | "asc"
            | "ascw"
            | "atn"
            | "cbool"
            | "cbyte"
            | "cdate"
            | "ccur"
            | "cdec"
            | "cdbl"
            | "choose"
            | "chr"
            | "chrw"
            | "cint"
            | "clng"
            | "clnglng"
            | "clngptr"
            | "cos"
            | "csng"
            | "cstr"
            | "cvar"
            | "cverr"
            | "dateadd"
            | "datediff"
            | "datepart"
            | "dateserial"
            | "datevalue"
            | "day"
            | "ddb"
            | "doevents"
            | "exp"
            | "filter"
            | "fix"
            | "fv"
            | "format"
            | "formatcurrency"
            | "formatdatetime"
            | "formatnumber"
            | "formatpercent"
            | "hex"
            | "hour"
            | "iif"
            | "instr"
            | "instrrev"
            | "int"
            | "ipmt"
            | "irr"
            | "isarray"
            | "isdate"
            | "isempty"
            | "iserror"
            | "ismissing"
            | "isnull"
            | "isnumeric"
            | "isobject"
            | "join"
            | "lbound"
            | "lcase"
            | "left"
            | "len"
            | "log"
            | "ltrim"
            | "mid"
            | "minute"
            | "mirr"
            | "month"
            | "monthname"
            | "nper"
            | "npv"
            | "oct"
            | "partition"
            | "pmt"
            | "ppmt"
            | "pv"
            | "qbcolor"
            | "rate"
            | "replace"
            | "rgb"
            | "right"
            | "round"
            | "rtrim"
            | "second"
            | "sgn"
            | "sin"
            | "sln"
            | "space"
            | "split"
            | "sqr"
            | "strcomp"
            | "strconv"
            | "str"
            | "string"
            | "strreverse"
            | "switch"
            | "syd"
            | "tan"
            | "timeserial"
            | "timevalue"
            | "trim"
            | "typename"
            | "ubound"
            | "ucase"
            | "val"
            | "vartype"
            | "weekday"
            | "weekdayname"
            | "year"
    );
    if !known {
        return None;
    }
    Some((|| {
        if matches!(
            name.as_str(),
            "format" | "formatcurrency" | "formatdatetime" | "formatnumber" | "formatpercent"
        ) {
            return call_format_builtin(&name, args, line);
        }
        if matches!(
            name.as_str(),
            "monthname" | "str" | "strconv" | "weekdayname"
        ) {
            return call_text_conversion_builtin(&name, args, line);
        }
        if matches!(
            name.as_str(),
            "ddb"
                | "fv"
                | "ipmt"
                | "irr"
                | "mirr"
                | "nper"
                | "npv"
                | "pmt"
                | "ppmt"
                | "pv"
                | "rate"
                | "sln"
                | "syd"
        ) {
            return call_financial_builtin(&name, args, line);
        }
        if name == "partition" {
            return call_partition_builtin(args, line);
        }
        if name == "doevents" {
            if !args.is_empty() {
                return Err(error(
                    RuntimeErrorKind::ArgumentCount,
                    format!("doevents expects no arguments, received {}", args.len()),
                    line,
                ));
            }
            return Ok(Value::Integer(0));
        }
        if name == "qbcolor" {
            if args.len() != 1 {
                return Err(error(
                    RuntimeErrorKind::ArgumentCount,
                    format!("qbcolor expects 1 argument, received {}", args.len()),
                    line,
                ));
            }
            let color = integer_argument(&args[0], line)?;
            const COLORS: [i64; 16] = [
                0x000000, 0x800000, 0x008000, 0x808000, 0x000080, 0x800080, 0x008080, 0xc0c0c0,
                0x808080, 0xff0000, 0x00ff00, 0xffff00, 0x0000ff, 0xff00ff, 0x00ffff, 0xffffff,
            ];
            let result = COLORS.get(color as usize).ok_or_else(|| {
                invalid_procedure_call(format!("invalid QBColor index: {color}"), line)
            })?;
            return Ok(Value::Integer(*result));
        }
        if matches!(
            name.as_str(),
            "cdate"
                | "dateadd"
                | "datediff"
                | "datepart"
                | "dateserial"
                | "datevalue"
                | "day"
                | "hour"
                | "isdate"
                | "minute"
                | "month"
                | "second"
                | "timeserial"
                | "timevalue"
                | "weekday"
                | "year"
        ) {
            return call_date_builtin(&name, args, line);
        }
        if name == "strcomp" {
            if !(2..=3).contains(&args.len()) {
                return Err(error(
                    RuntimeErrorKind::ArgumentCount,
                    format!("strcomp expects 2 or 3 arguments, received {}", args.len()),
                    line,
                ));
            }
            if matches!(args[0], Value::Null) || matches!(args[1], Value::Null) {
                return Ok(Value::Null);
            }
            let left = text(&args[0])
                .map_err(|message| error(RuntimeErrorKind::TypeMismatch, message, line))?;
            let right = text(&args[1])
                .map_err(|message| error(RuntimeErrorKind::TypeMismatch, message, line))?;
            let text_compare = compare_mode(args.get(2), option_compare_text, line)?;
            let (left, right) = if text_compare {
                (left.to_lowercase(), right.to_lowercase())
            } else {
                (left, right)
            };
            let ordering = left.encode_utf16().cmp(right.encode_utf16());
            return Ok(Value::Integer(match ordering {
                std::cmp::Ordering::Less => -1,
                std::cmp::Ordering::Equal => 0,
                std::cmp::Ordering::Greater => 1,
            }));
        }
        if name == "filter" {
            if !(2..=4).contains(&args.len()) {
                return Err(error(
                    RuntimeErrorKind::ArgumentCount,
                    format!("filter expects 2 to 4 arguments, received {}", args.len()),
                    line,
                ));
            }
            let Value::Array(array) = &args[0] else {
                return Err(error(
                    RuntimeErrorKind::TypeMismatch,
                    "Filter requires a one-dimensional String array",
                    line,
                ));
            };
            if array.dimensions.len() != 1 {
                return Err(error(
                    RuntimeErrorKind::TypeMismatch,
                    "Filter requires a one-dimensional String array",
                    line,
                ));
            }
            let needle = match &args[1] {
                Value::String(value) => value.encode_utf16().collect::<Vec<_>>(),
                Value::Null => {
                    return Err(error(
                        RuntimeErrorKind::TypeMismatch,
                        "invalid use of Null",
                        line,
                    ));
                }
                value => text(value)
                    .map_err(|message| error(RuntimeErrorKind::TypeMismatch, message, line))?
                    .encode_utf16()
                    .collect(),
            };
            let include = match args.get(2) {
                None | Some(Value::Missing) => true,
                Some(Value::Null) => {
                    return Err(error(
                        RuntimeErrorKind::TypeMismatch,
                        "invalid use of Null",
                        line,
                    ));
                }
                Some(value) => truthy(value)
                    .map_err(|message| error(RuntimeErrorKind::TypeMismatch, message, line))?,
            };
            let text_compare = compare_mode(args.get(3), option_compare_text, line)?;
            let mut values = Vec::new();
            for value in &array.values {
                let Value::String(value) = value else {
                    return Err(error(
                        RuntimeErrorKind::TypeMismatch,
                        "Filter source array must contain only Strings",
                        line,
                    ));
                };
                let source = value.encode_utf16().collect::<Vec<_>>();
                if utf16_find(&source, &needle, 0, text_compare).is_some() == include {
                    values.push(Value::String(value.clone()));
                }
            }
            return Ok(Value::Array(ArrayValue {
                dimensions: vec![ArrayDimension {
                    lower_bound: 0,
                    length: values.len(),
                }],
                values,
                element_default: Box::new(Value::String(String::new())),
                resizable: true,
            }));
        }
        if name == "rgb" {
            if args.len() != 3 {
                return Err(error(
                    RuntimeErrorKind::ArgumentCount,
                    format!("rgb expects 3 arguments, received {}", args.len()),
                    line,
                ));
            }
            let mut components = [0_i64; 3];
            for (component, value) in components.iter_mut().zip(args) {
                let value = integer_argument(value, line)?;
                if value < 0 {
                    return Err(invalid_procedure_call(
                        "RGB components cannot be negative".to_string(),
                        line,
                    ));
                }
                *component = value.min(255);
            }
            return Ok(Value::Integer(
                components[0] + components[1] * 256 + components[2] * 65_536,
            ));
        }
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
        if name == "iif" {
            if args.len() != 3 {
                return Err(error(
                    RuntimeErrorKind::ArgumentCount,
                    format!("iif expects 3 arguments, received {}", args.len()),
                    line,
                ));
            }
            return match &args[0] {
                Value::Null => Ok(Value::Null),
                condition => Ok(
                    if truthy(condition)
                        .map_err(|message| error(RuntimeErrorKind::TypeMismatch, message, line))?
                    {
                        args[1].clone()
                    } else {
                        args[2].clone()
                    },
                ),
            };
        }
        if name == "choose" {
            if args.len() < 2 {
                return Err(error(
                    RuntimeErrorKind::ArgumentCount,
                    format!(
                        "choose expects at least 2 arguments, received {}",
                        args.len()
                    ),
                    line,
                ));
            }
            if matches!(args[0], Value::Null) {
                return Ok(Value::Null);
            }
            let index = integer_argument(&args[0], line)?;
            return Ok(if index < 1 || index as usize >= args.len() {
                Value::Null
            } else {
                args[index as usize].clone()
            });
        }
        if name == "switch" {
            if args.is_empty() || !args.len().is_multiple_of(2) {
                return Err(error(
                    RuntimeErrorKind::ArgumentCount,
                    format!(
                        "switch expects one or more expression/value pairs, received {} argument(s)",
                        args.len()
                    ),
                    line,
                ));
            }
            for pair in args.chunks_exact(2) {
                if !matches!(pair[0], Value::Null)
                    && truthy(&pair[0])
                        .map_err(|message| error(RuntimeErrorKind::TypeMismatch, message, line))?
                {
                    return Ok(pair[1].clone());
                }
            }
            return Ok(Value::Null);
        }
        if name == "round" {
            if !(1..=2).contains(&args.len()) {
                return Err(error(
                    RuntimeErrorKind::ArgumentCount,
                    format!("round expects 1 or 2 arguments, received {}", args.len()),
                    line,
                ));
            }
            if matches!(args[0], Value::Null) {
                return Ok(Value::Null);
            }
            let value = number(&args[0])
                .map_err(|message| error(RuntimeErrorKind::TypeMismatch, message, line))?;
            let places = match args.get(1) {
                None | Some(Value::Missing) => 0,
                Some(value) => integer_argument(value, line)?,
            };
            if !(0..=15).contains(&places) {
                return Err(invalid_procedure_call(
                    "Round decimal places must be between 0 and 15".to_string(),
                    line,
                ));
            }
            let scale = 10_f64.powi(places as i32);
            return Ok(numeric_literal((value * scale).round_ties_even() / scale));
        }
        if matches!(name.as_str(), "hex" | "oct") {
            if args.len() != 1 {
                return Err(error(
                    RuntimeErrorKind::ArgumentCount,
                    format!("{name} expects 1 argument, received {}", args.len()),
                    line,
                ));
            }
            if matches!(args[0], Value::Null) {
                return Ok(Value::Null);
            }
            let value = number(&args[0])
                .map_err(|message| error(RuntimeErrorKind::TypeMismatch, message, line))?
                .round_ties_even();
            if !value.is_finite() || value < i32::MIN as f64 || value > i32::MAX as f64 {
                return Err(error(
                    RuntimeErrorKind::Overflow,
                    format!("overflow converting value with {name}"),
                    line,
                ));
            }
            let value = value as i32;
            return Ok(Value::String(if name == "hex" {
                format!("{:X}", value as u32)
            } else {
                format!("{:o}", value as u32)
            }));
        }
        if name == "val" {
            if args.len() != 1 {
                return Err(error(
                    RuntimeErrorKind::ArgumentCount,
                    format!("val expects 1 argument, received {}", args.len()),
                    line,
                ));
            }
            if matches!(args[0], Value::Null) {
                return Err(error(
                    RuntimeErrorKind::TypeMismatch,
                    "invalid use of Null",
                    line,
                ));
            }
            let source = text(&args[0])
                .map_err(|message| error(RuntimeErrorKind::TypeMismatch, message, line))?;
            return parse_val(&source)
                .map(numeric_literal)
                .map_err(|message| error(RuntimeErrorKind::TypeMismatch, message, line));
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
        if name == "cverr" {
            if args.len() != 1 {
                return Err(error(
                    RuntimeErrorKind::ArgumentCount,
                    format!("cverr expects 1 argument, received {}", args.len()),
                    line,
                ));
            }
            let number = integer_argument(&args[0], line)?;
            if !(0..=65_535).contains(&number) {
                return Err(invalid_procedure_call(
                    format!("invalid CVErr number: {number}"),
                    line,
                ));
            }
            return Ok(Value::Error(number));
        }
        if matches!(
            name.as_str(),
            "isarray"
                | "isempty"
                | "iserror"
                | "isnull"
                | "isnumeric"
                | "isobject"
                | "typename"
                | "vartype"
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
                "iserror" => Value::Boolean(matches!(value, Value::Error(_))),
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
            "atn" | "cos" | "sin" | "tan" => match value {
                Value::Null => Ok(Value::Null),
                _ => {
                    let value = number(value).map_err(mismatch)?;
                    Ok(Value::Double(match name.as_str() {
                        "atn" => value.atan(),
                        "cos" => value.cos(),
                        "sin" => value.sin(),
                        "tan" => value.tan(),
                        _ => unreachable!(),
                    }))
                }
            },
            "cbool" => match value {
                Value::Null => Err(mismatch("invalid use of Null".to_string())),
                _ => Ok(Value::Boolean(truthy(value).map_err(mismatch)?)),
            },
            "cbyte" => Ok(Value::Integer(convert_integer(value, 0, 255, line)?)),
            "ccur" => {
                let value = number(value).map_err(mismatch)?;
                const LIMIT: f64 = 922_337_203_685_477.6;
                if !value.is_finite() || value.abs() > LIMIT {
                    Err(error(
                        RuntimeErrorKind::Overflow,
                        "overflow converting value to Currency",
                        line,
                    ))
                } else {
                    Ok(numeric_literal(
                        (value * 10_000.0).round_ties_even() / 10_000.0,
                    ))
                }
            }
            "cdec" => {
                let value = number(value).map_err(mismatch)?;
                const LIMIT: f64 = 79_228_162_514_264_337_593_543_950_335.0;
                if !value.is_finite() || value.abs() > LIMIT {
                    Err(error(
                        RuntimeErrorKind::Overflow,
                        "overflow converting value to Decimal",
                        line,
                    ))
                } else {
                    Ok(Value::Double(value))
                }
            }
            "cdbl" => Ok(Value::Double(number(value).map_err(mismatch)?)),
            "cint" => Ok(Value::Integer(convert_integer(
                value, -32_768, 32_767, line,
            )?)),
            "clng" => Ok(Value::Integer(convert_integer(
                value,
                -2_147_483_648,
                2_147_483_647,
                line,
            )?)),
            "clnglng" | "clngptr" => {
                let value = number(value).map_err(mismatch)?.round_ties_even();
                if !value.is_finite()
                    || value < i64::MIN as f64
                    || value >= 9_223_372_036_854_775_808.0
                {
                    Err(error(
                        RuntimeErrorKind::Overflow,
                        format!("overflow converting value with {name}"),
                        line,
                    ))
                } else {
                    Ok(Value::Integer(value as i64))
                }
            }
            "csng" => {
                let value = number(value).map_err(mismatch)?;
                if !value.is_finite() || value.abs() > f64::from(f32::MAX) {
                    Err(error(
                        RuntimeErrorKind::Overflow,
                        "overflow converting value to Single",
                        line,
                    ))
                } else {
                    Ok(Value::Double(f64::from(value as f32)))
                }
            }
            "cstr" => match value {
                Value::Null => Err(mismatch("invalid use of Null".to_string())),
                _ => Ok(Value::String(text(value).map_err(mismatch)?)),
            },
            "cvar" => Ok(value.clone()),
            "fix" | "int" => match value {
                Value::Null => Ok(Value::Null),
                _ => {
                    let value = number(value).map_err(mismatch)?;
                    Ok(numeric_literal(if name == "int" {
                        value.floor()
                    } else {
                        value.trunc()
                    }))
                }
            },
            "exp" => match value {
                Value::Null => Ok(Value::Null),
                _ => {
                    let result = number(value).map_err(mismatch)?.exp();
                    if result.is_finite() {
                        Ok(Value::Double(result))
                    } else {
                        Err(error(RuntimeErrorKind::Overflow, "overflow in Exp", line))
                    }
                }
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
            "log" => match value {
                Value::Null => Ok(Value::Null),
                _ => {
                    let value = number(value).map_err(mismatch)?;
                    if value <= 0.0 {
                        Err(invalid_procedure_call(
                            "Log requires a positive number".to_string(),
                            line,
                        ))
                    } else {
                        Ok(Value::Double(value.ln()))
                    }
                }
            },
            "sgn" => match value {
                Value::Null => Ok(Value::Null),
                _ => {
                    let value = number(value).map_err(mismatch)?;
                    Ok(Value::Integer(if value > 0.0 {
                        1
                    } else if value < 0.0 {
                        -1
                    } else {
                        0
                    }))
                }
            },
            "sqr" => match value {
                Value::Null => Ok(Value::Null),
                _ => {
                    let value = number(value).map_err(mismatch)?;
                    if value < 0.0 {
                        Err(invalid_procedure_call(
                            "Sqr requires a nonnegative number".to_string(),
                            line,
                        ))
                    } else {
                        Ok(numeric_literal(value.sqrt()))
                    }
                }
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

fn call_format_builtin(
    name: &str,
    args: &[Value],
    line: Option<u32>,
) -> Result<Value, RuntimeError> {
    let count_error = || {
        error(
            RuntimeErrorKind::ArgumentCount,
            format!("invalid argument count for {name}: {}", args.len()),
            line,
        )
    };
    let mismatch = |message| error(RuntimeErrorKind::TypeMismatch, message, line);
    match name {
        "format" => {
            if !(1..=4).contains(&args.len()) {
                return Err(count_error());
            }
            let pattern = match args.get(1) {
                None | Some(Value::Missing) => "".to_string(),
                Some(value) => text(value).map_err(mismatch)?,
            };
            let first_day = first_day_of_week(args.get(2), line)?;
            let first_week = first_week_of_year(args.get(3), line)?;
            format_value(&args[0], &pattern, first_day, first_week)
                .map(Value::String)
                .map_err(mismatch)
        }
        "formatdatetime" => {
            if !(1..=2).contains(&args.len()) {
                return Err(count_error());
            }
            let mode = match args.get(1) {
                None | Some(Value::Missing) => 0,
                Some(value) => integer_argument(value, line)?,
            };
            let pattern = match mode {
                0 => "General Date",
                1 => "Long Date",
                2 => "Short Date",
                3 => "Long Time",
                4 => "Short Time",
                _ => {
                    return Err(invalid_procedure_call(
                        format!("invalid date format: {mode}"),
                        line,
                    ))
                }
            };
            format_date(
                value_date_serial(&args[0]).map_err(mismatch)?,
                pattern,
                1,
                1,
            )
            .map(Value::String)
            .map_err(mismatch)
        }
        "formatnumber" | "formatpercent" | "formatcurrency" => {
            if !(1..=5).contains(&args.len()) {
                return Err(count_error());
            }
            if matches!(args[0], Value::Null) {
                return Ok(Value::String(String::new()));
            }
            let value = number(&args[0]).map_err(mismatch)?;
            let digits = match args.get(1) {
                None | Some(Value::Missing) => 2,
                Some(value) => integer_argument(value, line)?,
            };
            if !(-1..=15).contains(&digits) {
                return Err(invalid_procedure_call(
                    "invalid decimal places".to_string(),
                    line,
                ));
            }
            let leading = tristate(args.get(2), true, line)?;
            let parens = tristate(args.get(3), false, line)?;
            let grouping = tristate(args.get(4), true, line)?;
            let percent = name == "formatpercent";
            let mut result = fixed_number(
                value * if percent { 100.0 } else { 1.0 },
                if digits == -1 { 2 } else { digits as usize },
                grouping,
                leading,
                parens,
            );
            if percent {
                result.push('%');
            } else if name == "formatcurrency" {
                result = currency_symbol(result);
            }
            Ok(Value::String(result))
        }
        _ => unreachable!(),
    }
}

fn call_financial_builtin(
    name: &str,
    args: &[Value],
    line: Option<u32>,
) -> Result<Value, RuntimeError> {
    let expected = match name {
        "irr" => 1..=2,
        "npv" => 2..=2,
        "sln" | "mirr" => 3..=3,
        "ddb" => 4..=5,
        "syd" => 4..=4,
        "rate" => 3..=6,
        "ipmt" | "ppmt" => 4..=6,
        _ => 3..=5,
    };
    if !expected.contains(&args.len()) {
        return Err(error(
            RuntimeErrorKind::ArgumentCount,
            format!("invalid argument count for {name}: {}", args.len()),
            line,
        ));
    }
    let mismatch = |message| error(RuntimeErrorKind::TypeMismatch, message, line);
    let numeric = |index: usize, default: f64| -> Result<f64, RuntimeError> {
        match args.get(index) {
            None | Some(Value::Missing) => Ok(default),
            Some(value) => number(value).map_err(mismatch),
        }
    };
    let payment_type = |index: usize| -> Result<f64, RuntimeError> {
        let value = numeric(index, 0.0)?;
        if value == 0.0 || value == 1.0 {
            Ok(value)
        } else {
            Err(invalid_procedure_call(
                format!("invalid payment type: {value}"),
                line,
            ))
        }
    };
    let result = match name {
        "ddb" => financial_ddb(
            numeric(0, 0.0)?,
            numeric(1, 0.0)?,
            numeric(2, 0.0)?,
            numeric(3, 0.0)?,
            numeric(4, 2.0)?,
        ),
        "fv" => financial_fv(
            numeric(0, 0.0)?,
            numeric(1, 0.0)?,
            numeric(2, 0.0)?,
            numeric(3, 0.0)?,
            payment_type(4)?,
        ),
        "pv" => financial_pv(
            numeric(0, 0.0)?,
            numeric(1, 0.0)?,
            numeric(2, 0.0)?,
            numeric(3, 0.0)?,
            payment_type(4)?,
        ),
        "pmt" => financial_pmt(
            numeric(0, 0.0)?,
            numeric(1, 0.0)?,
            numeric(2, 0.0)?,
            numeric(3, 0.0)?,
            payment_type(4)?,
        ),
        "nper" => financial_nper(
            numeric(0, 0.0)?,
            numeric(1, 0.0)?,
            numeric(2, 0.0)?,
            numeric(3, 0.0)?,
            payment_type(4)?,
        ),
        "ipmt" | "ppmt" => {
            let rate = numeric(0, 0.0)?;
            let period = numeric(1, 0.0)?;
            let periods = numeric(2, 0.0)?;
            if period < 1.0 || period > periods {
                Err("payment period is outside the annuity".to_string())
            } else {
                let present = numeric(3, 0.0)?;
                let future = numeric(4, 0.0)?;
                let kind = payment_type(5)?;
                (|| {
                    let payment = financial_pmt(rate, periods, present, future, kind)?;
                    let interest = if kind == 1.0 && period == 1.0 {
                        0.0
                    } else {
                        let balance = financial_fv(rate, period - 1.0, payment, present, kind)?;
                        let interest = balance * rate;
                        if kind == 1.0 {
                            interest / (1.0 + rate)
                        } else {
                            interest
                        }
                    };
                    if name == "ipmt" {
                        Ok(interest)
                    } else {
                        Ok(payment - interest)
                    }
                })()
            }
        }
        "irr" => financial_irr(
            &cash_flow_values(&args[0]).map_err(mismatch)?,
            numeric(1, 0.1)?,
        ),
        "mirr" => financial_mirr(
            &cash_flow_values(&args[0]).map_err(mismatch)?,
            numeric(1, 0.0)?,
            numeric(2, 0.0)?,
        ),
        "npv" => financial_npv(
            numeric(0, 0.0)?,
            &cash_flow_values(&args[1]).map_err(mismatch)?,
        ),
        "rate" => financial_rate(
            numeric(0, 0.0)?,
            numeric(1, 0.0)?,
            numeric(2, 0.0)?,
            numeric(3, 0.0)?,
            payment_type(4)?,
            numeric(5, 0.1)?,
        ),
        "sln" => financial_sln(numeric(0, 0.0)?, numeric(1, 0.0)?, numeric(2, 0.0)?),
        "syd" => financial_syd(
            numeric(0, 0.0)?,
            numeric(1, 0.0)?,
            numeric(2, 0.0)?,
            numeric(3, 0.0)?,
        ),
        _ => unreachable!(),
    }
    .map_err(|message| invalid_procedure_call(message, line))?;
    if result.is_finite() {
        Ok(Value::Double(if result == 0.0 { 0.0 } else { result }))
    } else {
        Err(error(
            RuntimeErrorKind::Overflow,
            format!("{name} result overflow"),
            line,
        ))
    }
}

fn call_partition_builtin(args: &[Value], line: Option<u32>) -> Result<Value, RuntimeError> {
    if args.len() != 4 {
        return Err(error(
            RuntimeErrorKind::ArgumentCount,
            format!("partition expects 4 arguments, received {}", args.len()),
            line,
        ));
    }
    if args.iter().any(|value| matches!(value, Value::Null)) {
        return Ok(Value::Null);
    }
    let number = integer_argument(&args[0], line)?;
    let start = integer_argument(&args[1], line)?;
    let stop = integer_argument(&args[2], line)?;
    let interval = integer_argument(&args[3], line)?;
    if start < 0 || stop <= start || interval < 1 {
        return Err(invalid_procedure_call(
            "invalid Partition range".to_string(),
            line,
        ));
    }
    let before = start
        .checked_sub(1)
        .ok_or_else(|| error(RuntimeErrorKind::Overflow, "Partition range overflow", line))?;
    let after = stop
        .checked_add(1)
        .ok_or_else(|| error(RuntimeErrorKind::Overflow, "Partition range overflow", line))?;
    let width = after.to_string().len().max(before.to_string().len());
    let (lower, upper) = if number < start {
        (String::new(), before.to_string())
    } else if number > stop {
        (after.to_string(), String::new())
    } else {
        let offset = number - start;
        let lower = start + offset / interval * interval;
        let upper = lower.saturating_add(interval - 1).min(stop);
        (lower.to_string(), upper.to_string())
    };
    Ok(Value::String(format!("{lower:>width$}:{upper:>width$}")))
}

fn cash_flow_values(value: &Value) -> Result<Vec<f64>, String> {
    let Value::Array(array) = value else {
        return Err("cash flows must be a one-dimensional numeric array".to_string());
    };
    if array.dimensions.len() != 1 {
        return Err("cash flows must be a one-dimensional numeric array".to_string());
    }
    array.values.iter().map(number).collect()
}

fn financial_ddb(
    cost: f64,
    salvage: f64,
    life: f64,
    period: f64,
    factor: f64,
) -> Result<f64, String> {
    if cost < 0.0 || salvage < 0.0 || life <= 0.0 || period <= 0.0 || factor <= 0.0 {
        return Err("invalid depreciation arguments".to_string());
    }
    if period > life {
        return Err("depreciation period exceeds asset life".to_string());
    }
    if cost <= salvage {
        return Ok(0.0);
    }
    let rate = (factor / life).min(1.0);
    let book_value = cost * (1.0 - rate).powf(period - 1.0);
    Ok((book_value * rate).min((book_value - salvage).max(0.0)))
}

fn financial_sln(cost: f64, salvage: f64, life: f64) -> Result<f64, String> {
    if cost < 0.0 || salvage < 0.0 || life <= 0.0 {
        return Err("invalid depreciation arguments".to_string());
    }
    Ok((cost - salvage) / life)
}

fn financial_syd(cost: f64, salvage: f64, life: f64, period: f64) -> Result<f64, String> {
    if cost < 0.0 || salvage < 0.0 || life <= 0.0 || period <= 0.0 || period > life {
        return Err("invalid depreciation arguments".to_string());
    }
    Ok((cost - salvage) * (life - period + 1.0) * 2.0 / (life * (life + 1.0)))
}

fn financial_irr(values: &[f64], guess: f64) -> Result<f64, String> {
    if values.len() < 2
        || !values.iter().any(|value| *value < 0.0)
        || !values.iter().any(|value| *value > 0.0)
        || !guess.is_finite()
        || guess <= -1.0
    {
        return Err("invalid IRR cash flows or guess".to_string());
    }
    let mut rate = guess;
    for _ in 0..20 {
        let base = 1.0 + rate;
        let mut value = 0.0;
        let mut derivative = 0.0;
        for (period, cash_flow) in values.iter().enumerate() {
            let period = period as f64;
            value += cash_flow / base.powf(period);
            if period != 0.0 {
                derivative -= period * cash_flow / base.powf(period + 1.0);
            }
        }
        if !value.is_finite() || !derivative.is_finite() || derivative == 0.0 {
            break;
        }
        let mut next = rate - value / derivative;
        if next <= -1.0 {
            next = (rate - 1.0) / 2.0;
        }
        if !next.is_finite() {
            break;
        }
        if (next - rate).abs() <= 1e-7 {
            return Ok(next);
        }
        rate = next;
    }
    Err("IRR failed to converge after 20 iterations".to_string())
}

fn financial_mirr(values: &[f64], finance_rate: f64, reinvest_rate: f64) -> Result<f64, String> {
    if values.len() < 2
        || !values.iter().any(|value| *value < 0.0)
        || !values.iter().any(|value| *value > 0.0)
        || finance_rate <= -1.0
        || reinvest_rate <= -1.0
    {
        return Err("invalid MIRR cash flows or rates".to_string());
    }
    let last_period = values.len() - 1;
    let mut positive_future_value = 0.0;
    let mut negative_present_value = 0.0;
    for (period, cash_flow) in values.iter().enumerate() {
        if *cash_flow > 0.0 {
            positive_future_value +=
                cash_flow * (1.0 + reinvest_rate).powf((last_period - period) as f64);
        } else if *cash_flow < 0.0 {
            negative_present_value += cash_flow / (1.0 + finance_rate).powf(period as f64);
        }
    }
    let ratio = -positive_future_value / negative_present_value;
    if !ratio.is_finite() || ratio <= 0.0 {
        return Err("MIRR has no real solution".to_string());
    }
    Ok(ratio.powf(1.0 / last_period as f64) - 1.0)
}

fn financial_npv(rate: f64, values: &[f64]) -> Result<f64, String> {
    if values.is_empty()
        || !values.iter().any(|value| *value < 0.0)
        || !values.iter().any(|value| *value > 0.0)
        || rate == -1.0
    {
        return Err("invalid NPV cash flows or rate".to_string());
    }
    let base = 1.0 + rate;
    Ok(values
        .iter()
        .enumerate()
        .map(|(period, cash_flow)| cash_flow / base.powf((period + 1) as f64))
        .sum())
}

fn annuity_factor(rate: f64, periods: f64) -> Result<f64, String> {
    if rate <= -1.0 {
        return Err("interest rate must be greater than -1".to_string());
    }
    let factor = (1.0 + rate).powf(periods);
    if factor.is_finite() {
        Ok(factor)
    } else {
        Err("annuity factor overflow".to_string())
    }
}

fn financial_fv(
    rate: f64,
    periods: f64,
    payment: f64,
    present: f64,
    kind: f64,
) -> Result<f64, String> {
    if rate == 0.0 {
        return Ok(-(present + payment * periods));
    }
    let factor = annuity_factor(rate, periods)?;
    Ok(-(present * factor + payment * (1.0 + rate * kind) * (factor - 1.0) / rate))
}

fn financial_pv(
    rate: f64,
    periods: f64,
    payment: f64,
    future: f64,
    kind: f64,
) -> Result<f64, String> {
    if rate == 0.0 {
        return Ok(-(future + payment * periods));
    }
    let factor = annuity_factor(rate, periods)?;
    Ok(-(future + payment * (1.0 + rate * kind) * (factor - 1.0) / rate) / factor)
}

fn financial_pmt(
    rate: f64,
    periods: f64,
    present: f64,
    future: f64,
    kind: f64,
) -> Result<f64, String> {
    if periods == 0.0 {
        return Err("payment periods cannot be zero".to_string());
    }
    if rate == 0.0 {
        return Ok(-(future + present) / periods);
    }
    let factor = annuity_factor(rate, periods)?;
    let denominator = (1.0 + rate * kind) * (factor - 1.0);
    if denominator == 0.0 {
        return Err("payment denominator is zero".to_string());
    }
    Ok(-(future + present * factor) * rate / denominator)
}

fn financial_nper(
    rate: f64,
    payment: f64,
    present: f64,
    future: f64,
    kind: f64,
) -> Result<f64, String> {
    if rate == 0.0 {
        if payment == 0.0 {
            return Err("payment cannot be zero".to_string());
        }
        return Ok(-(present + future) / payment);
    }
    if rate <= -1.0 {
        return Err("interest rate must be greater than -1".to_string());
    }
    let adjusted = payment * (1.0 + rate * kind);
    let denominator = present * rate + adjusted;
    let ratio = (adjusted - future * rate) / denominator;
    if denominator == 0.0 || ratio <= 0.0 {
        return Err("annuity has no real payment-period solution".to_string());
    }
    Ok(ratio.ln() / (1.0 + rate).ln())
}

fn financial_equation(
    rate: f64,
    periods: f64,
    payment: f64,
    present: f64,
    future: f64,
    kind: f64,
) -> Result<f64, String> {
    if rate.abs() < 1e-12 {
        return Ok(present + payment * periods + future);
    }
    let factor = annuity_factor(rate, periods)?;
    Ok(present * factor + payment * (1.0 + rate * kind) * (factor - 1.0) / rate + future)
}

fn financial_rate(
    periods: f64,
    payment: f64,
    present: f64,
    future: f64,
    kind: f64,
    guess: f64,
) -> Result<f64, String> {
    if periods <= 0.0 || !guess.is_finite() || guess <= -1.0 {
        return Err("invalid Rate arguments".to_string());
    }
    let mut rate = guess;
    for _ in 0..20 {
        let value = financial_equation(rate, periods, payment, present, future, kind)?;
        let step = (rate.abs() * 1e-6).max(1e-7);
        let lower = (rate - step).max(-0.999_999_999);
        let upper = rate + step;
        let derivative = (financial_equation(upper, periods, payment, present, future, kind)?
            - financial_equation(lower, periods, payment, present, future, kind)?)
            / (upper - lower);
        if derivative == 0.0 || !derivative.is_finite() {
            break;
        }
        let mut next = rate - value / derivative;
        if next <= -1.0 {
            next = (rate - 1.0) / 2.0;
        }
        if !next.is_finite() {
            break;
        }
        if (next - rate).abs() <= 1e-7 {
            return Ok(next);
        }
        rate = next;
    }
    Err("Rate failed to converge after 20 iterations".to_string())
}

fn call_text_conversion_builtin(
    name: &str,
    args: &[Value],
    line: Option<u32>,
) -> Result<Value, RuntimeError> {
    let count_error = || {
        error(
            RuntimeErrorKind::ArgumentCount,
            format!("invalid argument count for {name}: {}", args.len()),
            line,
        )
    };
    let mismatch = |message| error(RuntimeErrorKind::TypeMismatch, message, line);
    match name {
        "str" => {
            if args.len() != 1 {
                return Err(count_error());
            }
            let value = number(&args[0]).map_err(mismatch)?;
            if !value.is_finite() {
                return Err(error(
                    RuntimeErrorKind::Overflow,
                    "overflow converting number with Str",
                    line,
                ));
            }
            let rendered = text(&numeric_literal(value)).map_err(mismatch)?;
            Ok(Value::String(if value.is_sign_negative() {
                rendered
            } else {
                format!(" {rendered}")
            }))
        }
        "monthname" => {
            if !(1..=2).contains(&args.len()) {
                return Err(count_error());
            }
            let month = integer_argument(&args[0], line)?;
            if !(1..=12).contains(&month) {
                return Err(invalid_procedure_call(
                    format!("invalid month number: {month}"),
                    line,
                ));
            }
            let abbreviate = optional_boolean(args.get(1), false, line)?;
            let name = month_name(month as u32);
            Ok(Value::String(if abbreviate {
                name[..3].to_string()
            } else {
                name.to_string()
            }))
        }
        "weekdayname" => {
            if !(1..=3).contains(&args.len()) {
                return Err(count_error());
            }
            let weekday = integer_argument(&args[0], line)?;
            if !(1..=7).contains(&weekday) {
                return Err(invalid_procedure_call(
                    format!("invalid weekday number: {weekday}"),
                    line,
                ));
            }
            let abbreviate = optional_boolean(args.get(1), false, line)?;
            let first_day = first_day_of_week(args.get(2), line)?;
            let absolute = (first_day + weekday - 2).rem_euclid(7) + 1;
            let name = weekday_name_by_number(absolute);
            Ok(Value::String(if abbreviate {
                name[..3].to_string()
            } else {
                name.to_string()
            }))
        }
        "strconv" => {
            if !(2..=3).contains(&args.len()) {
                return Err(count_error());
            }
            if matches!(args[0], Value::Null) {
                return Err(mismatch("invalid use of Null".to_string()));
            }
            let mut value = text(&args[0]).map_err(mismatch)?;
            let conversion = integer_argument(&args[1], line)?;
            if let Some(locale) = args.get(2) {
                integer_argument(locale, line)?;
            }
            if !(0..=255).contains(&conversion)
                || conversion & 4 != 0 && conversion & 8 != 0
                || conversion & 16 != 0 && conversion & 32 != 0
            {
                return Err(invalid_procedure_call(
                    format!("invalid StrConv conversion: {conversion}"),
                    line,
                ));
            }
            if conversion & 4 != 0 {
                value = convert_width_unicode(&value, true);
            }
            if conversion & 16 != 0 {
                value = convert_kana_script(&value, true);
            } else if conversion & 32 != 0 {
                value = convert_kana_script(&value, false);
            }
            if conversion & 8 != 0 {
                value = convert_width_unicode(&value, false);
            }
            value = match conversion & 3 {
                0 => value,
                1 => value.to_uppercase(),
                2 => value.to_lowercase(),
                3 => proper_case(&value),
                _ => unreachable!(),
            };
            Ok(Value::String(value))
        }
        _ => unreachable!(),
    }
}

fn optional_boolean(
    value: Option<&Value>,
    default: bool,
    line: Option<u32>,
) -> Result<bool, RuntimeError> {
    match value {
        None | Some(Value::Missing) => Ok(default),
        Some(Value::Null) => Err(error(
            RuntimeErrorKind::TypeMismatch,
            "invalid use of Null",
            line,
        )),
        Some(value) => {
            truthy(value).map_err(|message| error(RuntimeErrorKind::TypeMismatch, message, line))
        }
    }
}

fn proper_case(value: &str) -> String {
    let mut at_word_start = true;
    let mut result = String::new();
    for character in value.chars() {
        if character.is_alphanumeric() {
            if at_word_start {
                result.extend(character.to_uppercase());
            } else {
                result.extend(character.to_lowercase());
            }
            at_word_start = false;
        } else {
            result.push(character);
            at_word_start = true;
        }
    }
    result
}

fn convert_kana_script(value: &str, katakana: bool) -> String {
    value
        .chars()
        .map(|character| {
            let code = character as u32;
            let converted = if katakana && (0x3041..=0x3096).contains(&code) {
                code + 0x60
            } else if !katakana && (0x30a1..=0x30f6).contains(&code) {
                code - 0x60
            } else if katakana && (0x309d..=0x309e).contains(&code) {
                code + 0x60
            } else if !katakana && (0x30fd..=0x30fe).contains(&code) {
                code - 0x60
            } else {
                code
            };
            char::from_u32(converted).unwrap_or(character)
        })
        .collect()
}

fn convert_width_unicode(value: &str, wide: bool) -> String {
    let mut output = String::new();
    if wide {
        let mut input = value.chars().peekable();
        while let Some(character) = input.next() {
            let code = character as u32;
            if character == ' ' {
                output.push(char::from_u32(0x3000).unwrap());
            } else if (0x21..=0x7e).contains(&code) {
                output.push(char::from_u32(code + 0xfee0).unwrap());
            } else if let Some(mut converted) = halfwidth_kana_code(code) {
                if let Some(mark @ (0xff9e | 0xff9f)) = input.peek().map(|value| *value as u32) {
                    if let Some(composed) = compose_kana_code(converted, mark) {
                        converted = composed;
                        input.next();
                    }
                }
                output.push(char::from_u32(converted).unwrap());
            } else {
                output.push(character);
            }
        }
    } else {
        for character in value.chars() {
            let code = character as u32;
            if code == 0x3000 {
                output.push(' ');
            } else if (0xff01..=0xff5e).contains(&code) {
                output.push(char::from_u32(code - 0xfee0).unwrap());
            } else if let Some((base, mark)) = narrow_voiced_kana_codes(code) {
                output.push(char::from_u32(base).unwrap());
                output.push(char::from_u32(mark).unwrap());
            } else if let Some(halfwidth) = fullwidth_kana_code(code) {
                output.push(char::from_u32(halfwidth).unwrap());
            } else {
                output.push(character);
            }
        }
    }
    output
}

fn halfwidth_kana_code(code: u32) -> Option<u32> {
    const FULLWIDTH: [u32; 63] = [
        0x3002, 0x300c, 0x300d, 0x3001, 0x30fb, 0x30f2, 0x30a1, 0x30a3, 0x30a5, 0x30a7, 0x30a9,
        0x30e3, 0x30e5, 0x30e7, 0x30c3, 0x30fc, 0x30a2, 0x30a4, 0x30a6, 0x30a8, 0x30aa, 0x30ab,
        0x30ad, 0x30af, 0x30b1, 0x30b3, 0x30b5, 0x30b7, 0x30b9, 0x30bb, 0x30bd, 0x30bf, 0x30c1,
        0x30c4, 0x30c6, 0x30c8, 0x30ca, 0x30cb, 0x30cc, 0x30cd, 0x30ce, 0x30cf, 0x30d2, 0x30d5,
        0x30d8, 0x30db, 0x30de, 0x30df, 0x30e0, 0x30e1, 0x30e2, 0x30e4, 0x30e6, 0x30e8, 0x30e9,
        0x30ea, 0x30eb, 0x30ec, 0x30ed, 0x30ef, 0x30f3, 0x3099, 0x309a,
    ];
    let index = usize::try_from(code.checked_sub(0xff61)?).ok()?;
    FULLWIDTH.get(index).copied()
}

fn fullwidth_kana_code(code: u32) -> Option<u32> {
    (0xff61..=0xff9f).find(|halfwidth| halfwidth_kana_code(*halfwidth) == Some(code))
}

fn compose_kana_code(base: u32, mark: u32) -> Option<u32> {
    Some(match (base, mark) {
        (0x30a6, 0xff9e) => 0x30f4,
        (0x30ab | 0x30ad | 0x30af | 0x30b1 | 0x30b3, 0xff9e)
        | (0x30b5 | 0x30b7 | 0x30b9 | 0x30bb | 0x30bd, 0xff9e)
        | (0x30bf | 0x30c1 | 0x30c6 | 0x30c8, 0xff9e)
        | (0x30cf | 0x30d2 | 0x30d5 | 0x30d8 | 0x30db, 0xff9e) => base + 1,
        (0x30c4, 0xff9e) => 0x30c5,
        (0x30cf | 0x30d2 | 0x30d5 | 0x30d8 | 0x30db, 0xff9f) => base + 2,
        (0x30ef, 0xff9e) => 0x30f7,
        (0x30f2, 0xff9e) => 0x30fa,
        _ => return None,
    })
}

fn narrow_voiced_kana_codes(code: u32) -> Option<(u32, u32)> {
    let (base, mark) = match code {
        0x30f4 => (0x30a6, 0xff9e),
        0x30ac | 0x30ae | 0x30b0 | 0x30b2 | 0x30b4 | 0x30b6 | 0x30b8 | 0x30ba | 0x30bc | 0x30be
        | 0x30c0 | 0x30c2 | 0x30c7 | 0x30c9 | 0x30d0 | 0x30d3 | 0x30d6 | 0x30d9 | 0x30dc => {
            (code - 1, 0xff9e)
        }
        0x30c5 => (0x30c4, 0xff9e),
        0x30d1 | 0x30d4 | 0x30d7 | 0x30da | 0x30dd => (code - 2, 0xff9f),
        0x30f7 => (0x30ef, 0xff9e),
        0x30fa => (0x30f2, 0xff9e),
        _ => return None,
    };
    Some((fullwidth_kana_code(base)?, mark))
}

fn tristate(value: Option<&Value>, default: bool, line: Option<u32>) -> Result<bool, RuntimeError> {
    match value {
        None | Some(Value::Missing) => Ok(default),
        Some(value) => match integer_argument(value, line)? {
            -2 => Ok(default),
            -1 => Ok(true),
            0 => Ok(false),
            value => Err(invalid_procedure_call(
                format!("invalid tristate: {value}"),
                line,
            )),
        },
    }
}

fn format_value(
    value: &Value,
    pattern: &str,
    first_day: i64,
    first_week: i64,
) -> Result<String, String> {
    if matches!(value, Value::Null) {
        return Ok(String::new());
    }
    if pattern.is_empty() {
        return text(value);
    }
    let lower = pattern.to_ascii_lowercase();
    let named_date = matches!(
        lower.as_str(),
        "general date"
            | "long date"
            | "medium date"
            | "short date"
            | "long time"
            | "medium time"
            | "short time"
    );
    let custom_date = lower.contains("yyyy")
        || lower.contains("ddd")
        || lower.contains("am/pm")
        || lower.contains("hh")
        || lower.contains("nn")
        || (lower.contains('/') && lower.contains('d'))
        || (lower.contains(':') && lower.contains('h'));
    if named_date || custom_date {
        return format_date(value_date_serial(value)?, pattern, first_day, first_week);
    }
    if let Value::String(value) = value {
        return Ok(match pattern {
            "<" => value.to_lowercase(),
            ">" => value.to_uppercase(),
            _ if pattern.contains('@') => pattern.replace('@', value),
            _ => value.clone(),
        });
    }
    let value = number(value)?;
    Ok(match lower.as_str() {
        "general number" => text(&numeric_literal(value))?,
        "currency" => currency_symbol(fixed_number(value, 2, true, true, true)),
        "fixed" => fixed_number(value, 2, false, true, false),
        "standard" => fixed_number(value, 2, true, true, false),
        "percent" => format!("{}%", fixed_number(value * 100.0, 2, false, true, false)),
        "scientific" => format!("{value:.2E}"),
        "yes/no" => if value == 0.0 { "No" } else { "Yes" }.to_string(),
        "true/false" => if value == 0.0 { "False" } else { "True" }.to_string(),
        "on/off" => if value == 0.0 { "Off" } else { "On" }.to_string(),
        _ => custom_number(value, pattern),
    })
}

fn format_date(
    serial: f64,
    pattern: &str,
    first_day: i64,
    first_week: i64,
) -> Result<String, String> {
    let parts = serial_date_parts(serial)?;
    let lower = pattern.to_ascii_lowercase();
    match lower.as_str() {
        "general date" => {
            let has_date = serial.floor() != 0.0;
            let has_time = serial.rem_euclid(1.0) != 0.0;
            return Ok(match (has_date, has_time) {
                (true, true) => format!(
                    "{}/{}/{} {}",
                    parts.month,
                    parts.day,
                    parts.year,
                    clock(parts, true)
                ),
                (false, true) => clock(parts, true),
                _ => format!("{}/{}/{}", parts.month, parts.day, parts.year),
            });
        }
        "long date" => {
            return Ok(format!(
                "{}, {} {}, {}",
                weekday_name(serial),
                month_name(parts.month),
                parts.day,
                parts.year
            ))
        }
        "medium date" => {
            return Ok(format!(
                "{}-{}-{:02}",
                parts.day,
                &month_name(parts.month)[..3],
                parts.year.rem_euclid(100)
            ))
        }
        "short date" => return Ok(format!("{}/{}/{}", parts.month, parts.day, parts.year)),
        "long time" => return Ok(clock(parts, true)),
        "medium time" => return Ok(clock(parts, false)),
        "short time" => return Ok(format!("{:02}:{:02}", parts.hour, parts.minute)),
        _ => {}
    }
    let has_ampm = lower.contains("am/pm") || lower.contains("a/p");
    let mut output = String::new();
    let mut cursor = 0;
    let mut after_hour = false;
    let bytes = pattern.as_bytes();
    while cursor < pattern.len() {
        if bytes[cursor] == b'"' {
            cursor += 1;
            while cursor < pattern.len() && bytes[cursor] != b'"' {
                let character = pattern[cursor..].chars().next().unwrap();
                output.push(character);
                cursor += character.len_utf8();
            }
            cursor += usize::from(cursor < pattern.len());
            continue;
        }
        if bytes[cursor] == b'\\' && cursor + 1 < pattern.len() {
            cursor += 1;
            let character = pattern[cursor..].chars().next().unwrap();
            output.push(character);
            cursor += character.len_utf8();
            continue;
        }
        let remaining = &lower[cursor..];
        let (length, replacement) = if remaining.starts_with("am/pm") {
            (5, if parts.hour < 12 { "AM" } else { "PM" }.to_string())
        } else if remaining.starts_with("a/p") {
            (3, if parts.hour < 12 { "A" } else { "P" }.to_string())
        } else if remaining.starts_with("yyyy") {
            (4, format!("{:04}", parts.year))
        } else if remaining.starts_with("mmmm") {
            (4, month_name(parts.month).to_string())
        } else if remaining.starts_with("dddd") {
            (4, weekday_name(serial).to_string())
        } else if remaining.starts_with("mmm") {
            (3, month_name(parts.month)[..3].to_string())
        } else if remaining.starts_with("ddd") {
            (3, weekday_name(serial)[..3].to_string())
        } else if remaining.starts_with("yy") {
            (2, format!("{:02}", parts.year.rem_euclid(100)))
        } else if remaining.starts_with("dd") {
            (2, format!("{:02}", parts.day))
        } else if remaining.starts_with("hh") {
            after_hour = true;
            (2, format!("{:02}", hour_value(parts.hour, has_ampm)))
        } else if remaining.starts_with("nn") {
            (2, format!("{:02}", parts.minute))
        } else if remaining.starts_with("ss") {
            (2, format!("{:02}", parts.second))
        } else if remaining.starts_with("ww") {
            (
                2,
                week_of_year(serial.floor() as i64, parts.year, first_day, first_week)?.to_string(),
            )
        } else if remaining.starts_with("mm") {
            let value = if after_hour {
                parts.minute
            } else {
                parts.month
            };
            after_hour = false;
            (2, format!("{value:02}"))
        } else {
            let character = remaining.chars().next().unwrap();
            let replacement = match character {
                'q' => ((parts.month - 1) / 3 + 1).to_string(),
                'm' => {
                    let value = if after_hour {
                        parts.minute
                    } else {
                        parts.month
                    };
                    after_hour = false;
                    value.to_string()
                }
                'd' => parts.day.to_string(),
                'y' => (serial.floor() as i64 - date_serial(parts.year, 1, 1)?.floor() as i64 + 1)
                    .to_string(),
                'h' => {
                    after_hour = true;
                    hour_value(parts.hour, has_ampm).to_string()
                }
                'n' => parts.minute.to_string(),
                's' => parts.second.to_string(),
                'w' => weekday_number(serial.floor() as i64, first_day).to_string(),
                _ => {
                    output.push(pattern[cursor..].chars().next().unwrap());
                    cursor += pattern[cursor..].chars().next().unwrap().len_utf8();
                    continue;
                }
            };
            (character.len_utf8(), replacement)
        };
        output.push_str(&replacement);
        cursor += length;
    }
    Ok(output)
}

fn month_name(month: u32) -> &'static str {
    const NAMES: [&str; 12] = [
        "January",
        "February",
        "March",
        "April",
        "May",
        "June",
        "July",
        "August",
        "September",
        "October",
        "November",
        "December",
    ];
    NAMES[month as usize - 1]
}

fn weekday_name(serial: f64) -> &'static str {
    weekday_name_by_number(weekday_number(serial.floor() as i64, 1))
}

fn weekday_name_by_number(weekday: i64) -> &'static str {
    const NAMES: [&str; 7] = [
        "Sunday",
        "Monday",
        "Tuesday",
        "Wednesday",
        "Thursday",
        "Friday",
        "Saturday",
    ];
    NAMES[weekday as usize - 1]
}

fn hour_value(hour: u32, meridiem: bool) -> u32 {
    if !meridiem {
        return hour;
    }
    let hour = hour % 12;
    if hour == 0 {
        12
    } else {
        hour
    }
}

fn clock(parts: DateParts, seconds: bool) -> String {
    let suffix = if parts.hour < 12 { "AM" } else { "PM" };
    if seconds {
        format!(
            "{}:{:02}:{:02} {suffix}",
            hour_value(parts.hour, true),
            parts.minute,
            parts.second
        )
    } else {
        format!(
            "{}:{:02} {suffix}",
            hour_value(parts.hour, true),
            parts.minute
        )
    }
}

fn fixed_number(value: f64, digits: usize, grouping: bool, leading: bool, parens: bool) -> String {
    let negative = value < 0.0;
    let raw = format!("{:.*}", digits, value.abs());
    let (whole, fraction) = raw.split_once('.').unwrap_or((&raw, ""));
    let mut whole = if grouping {
        group_number(whole)
    } else {
        whole.to_string()
    };
    if !leading && whole == "0" && digits > 0 {
        whole.clear();
    }
    let mut result = if digits == 0 {
        whole
    } else {
        format!("{whole}.{fraction}")
    };
    if negative {
        result = if parens {
            format!("({result})")
        } else {
            format!("-{result}")
        };
    }
    result
}

fn currency_symbol(mut value: String) -> String {
    if value.starts_with('(') {
        value.insert(1, '$');
    } else if let Some(rest) = value.strip_prefix('-') {
        value = format!("-${rest}");
    } else {
        value.insert(0, '$');
    }
    value
}

fn group_number(value: &str) -> String {
    let mut result = String::new();
    for (index, character) in value.chars().enumerate() {
        if index > 0 && (value.len() - index).is_multiple_of(3) {
            result.push(',');
        }
        result.push(character);
    }
    result
}

fn custom_number(value: f64, pattern: &str) -> String {
    let percent = pattern.contains('%');
    let decimal = pattern.find('.');
    let digits = decimal
        .map(|index| {
            pattern[index + 1..]
                .chars()
                .filter(|c| matches!(c, '0' | '#'))
                .count()
        })
        .unwrap_or(0);
    let required = decimal
        .map(|index| pattern[index + 1..].chars().filter(|c| *c == '0').count())
        .unwrap_or(0);
    let whole_pattern = &pattern[..decimal.unwrap_or(pattern.len())];
    let mut result = fixed_number(
        value * if percent { 100.0 } else { 1.0 },
        digits,
        whole_pattern.contains(','),
        whole_pattern.contains('0'),
        false,
    );
    let mut displayed_digits = digits;
    while displayed_digits > required && result.ends_with('0') && result.contains('.') {
        result.pop();
        displayed_digits -= 1;
    }
    if result.ends_with('.') {
        result.pop();
    }
    if percent {
        result.push('%');
    }
    if pattern.starts_with('$') {
        result.insert(0, '$');
    }
    result
}

fn call_date_builtin(name: &str, args: &[Value], line: Option<u32>) -> Result<Value, RuntimeError> {
    let wrong_count = |expected: &str| {
        error(
            RuntimeErrorKind::ArgumentCount,
            format!("{name} expects {expected}, received {}", args.len()),
            line,
        )
    };
    let mismatch = |message| error(RuntimeErrorKind::TypeMismatch, message, line);
    match name {
        "dateserial" | "timeserial" => {
            if args.len() != 3 {
                return Err(wrong_count("3 arguments"));
            }
            let mut parts = [0_i64; 3];
            for (target, value) in parts.iter_mut().zip(args) {
                *target = integer_argument(value, line)?;
                if !(-32_768..=32_767).contains(target) {
                    return Err(error(
                        RuntimeErrorKind::Overflow,
                        format!("{name} argument is outside the Integer range"),
                        line,
                    ));
                }
            }
            let serial = if name == "dateserial" {
                date_serial(parts[0], parts[1], parts[2])
            } else {
                Ok((parts[0] * 3_600 + parts[1] * 60 + parts[2]) as f64 / 86_400.0)
            }
            .map_err(|message| invalid_procedure_call(message, line))?;
            Ok(Value::Double(serial))
        }
        "cdate" | "datevalue" | "timevalue" | "isdate" => {
            if args.len() != 1 {
                return Err(wrong_count("1 argument"));
            }
            let parsed = value_date_serial(&args[0]);
            if name == "isdate" {
                return Ok(Value::Boolean(parsed.is_ok()));
            }
            let serial = parsed.map_err(mismatch)?;
            Ok(Value::Double(match name {
                "datevalue" => serial.floor(),
                "timevalue" => serial.rem_euclid(1.0),
                _ => serial,
            }))
        }
        "year" | "month" | "day" | "hour" | "minute" | "second" => {
            if args.len() != 1 {
                return Err(wrong_count("1 argument"));
            }
            let serial = value_date_serial(&args[0]).map_err(mismatch)?;
            let parts = serial_date_parts(serial).map_err(mismatch)?;
            Ok(Value::Integer(match name {
                "year" => parts.year,
                "month" => i64::from(parts.month),
                "day" => i64::from(parts.day),
                "hour" => i64::from(parts.hour),
                "minute" => i64::from(parts.minute),
                "second" => i64::from(parts.second),
                _ => unreachable!(),
            }))
        }
        "dateadd" => {
            if args.len() != 3 {
                return Err(wrong_count("3 arguments"));
            }
            let interval = text(&args[0]).map_err(mismatch)?.to_ascii_lowercase();
            let amount = integer_argument(&args[1], line)?;
            let serial = value_date_serial(&args[2]).map_err(mismatch)?;
            date_add(&interval, amount, serial)
                .map(Value::Double)
                .map_err(|message| invalid_procedure_call(message, line))
        }
        "datediff" => {
            if !(3..=5).contains(&args.len()) {
                return Err(wrong_count("3 to 5 arguments"));
            }
            let interval = text(&args[0]).map_err(mismatch)?.to_ascii_lowercase();
            let first = value_date_serial(&args[1]).map_err(mismatch)?;
            let second = value_date_serial(&args[2]).map_err(mismatch)?;
            let first_day = first_day_of_week(args.get(3), line)?;
            first_week_of_year(args.get(4), line)?;
            date_diff(&interval, first, second, first_day)
                .map(Value::Integer)
                .map_err(|message| invalid_procedure_call(message, line))
        }
        "datepart" => {
            if !(2..=4).contains(&args.len()) {
                return Err(wrong_count("2 to 4 arguments"));
            }
            let interval = text(&args[0]).map_err(mismatch)?.to_ascii_lowercase();
            let serial = value_date_serial(&args[1]).map_err(mismatch)?;
            let first_day = first_day_of_week(args.get(2), line)?;
            let first_week = first_week_of_year(args.get(3), line)?;
            date_part(&interval, serial, first_day, first_week)
                .map(Value::Integer)
                .map_err(|message| invalid_procedure_call(message, line))
        }
        "weekday" => {
            if !(1..=2).contains(&args.len()) {
                return Err(wrong_count("1 or 2 arguments"));
            }
            let serial = value_date_serial(&args[0]).map_err(mismatch)?;
            let first_day = first_day_of_week(args.get(1), line)?;
            Ok(Value::Integer(weekday_number(
                serial.floor() as i64,
                first_day,
            )))
        }
        _ => unreachable!(),
    }
}

#[derive(Clone, Copy)]
struct DateParts {
    year: i64,
    month: u32,
    day: u32,
    hour: u32,
    minute: u32,
    second: u32,
}

fn date_serial(year: i64, month: i64, day: i64) -> Result<f64, String> {
    let year = match year {
        0..=29 => year + 2_000,
        30..=99 => year + 1_900,
        _ => year,
    };
    let total_months = year
        .checked_mul(12)
        .and_then(|value| value.checked_add(month - 1))
        .ok_or_else(|| "DateSerial result is outside the supported range".to_string())?;
    let normalized_year = total_months.div_euclid(12);
    let normalized_month = total_months.rem_euclid(12) as u32 + 1;
    let days = days_from_civil(normalized_year, normalized_month, 1)
        .checked_add(day - 1)
        .ok_or_else(|| "DateSerial result is outside the supported range".to_string())?;
    let (result_year, _, _) = civil_from_days(days);
    if !(100..=9_999).contains(&result_year) {
        return Err("DateSerial result year must be between 100 and 9999".to_string());
    }
    Ok((days - ole_epoch_days()) as f64)
}

fn value_date_serial(value: &Value) -> Result<f64, String> {
    let serial = match value {
        Value::Integer(value) => *value as f64,
        Value::Double(value) => *value,
        Value::String(value) => parse_date_text(value)?,
        Value::Empty => 0.0,
        Value::Null => return Err("invalid use of Null".to_string()),
        _ => return Err("value cannot be converted to a Date".to_string()),
    };
    if !serial.is_finite() {
        return Err("Date value must be finite".to_string());
    }
    serial_date_parts(serial)?;
    Ok(serial)
}

fn parse_date_text(source: &str) -> Result<f64, String> {
    let source = source.trim().trim_matches('#').trim();
    if source.is_empty() {
        return Err("Date string is empty".to_string());
    }
    if let Ok(value) = source.parse::<f64>() {
        return Ok(value);
    }
    if let Some(value) = parse_named_date_text(source) {
        return value;
    }
    let mut pieces = source.split_whitespace();
    let first = pieces.next().unwrap_or_default();
    if first.contains(':') {
        return parse_time_text(source);
    }
    let date = parse_date_part(first)?;
    let time_text = pieces.collect::<Vec<_>>().join(" ");
    if time_text.is_empty() {
        Ok(date)
    } else {
        Ok(date + parse_time_text(&time_text)?)
    }
}

fn parse_date_part(source: &str) -> Result<f64, String> {
    let delimiter = if source.contains('/') {
        '/'
    } else if source.contains('-') {
        '-'
    } else {
        return Err(format!("unsupported Date string: {source}"));
    };
    let values = source
        .split(delimiter)
        .map(|part| {
            part.parse::<i64>()
                .map_err(|_| format!("invalid Date component: {part}"))
        })
        .collect::<Result<Vec<_>, _>>()?;
    if values.len() != 3 {
        return Err(format!("Date requires three components: {source}"));
    }
    let (year, month, day) =
        if delimiter == '-' && source.split('-').next().is_some_and(|v| v.len() == 4) {
            (values[0], values[1], values[2])
        } else {
            (values[2], values[0], values[1])
        };
    strict_date_serial(year, month, day, source)
}

fn strict_date_serial(year: i64, month: i64, day: i64, source: &str) -> Result<f64, String> {
    let serial = date_serial(year, month, day)?;
    let parts = serial_date_parts(serial)?;
    let expected_year = match year {
        0..=29 => year + 2_000,
        30..=99 => year + 1_900,
        _ => year,
    };
    if parts.year != expected_year || i64::from(parts.month) != month || i64::from(parts.day) != day
    {
        return Err(format!("invalid calendar Date: {source}"));
    }
    Ok(serial)
}

fn parse_named_date_text(source: &str) -> Option<Result<f64, String>> {
    let normalized = source.replace(',', " ");
    let pieces = normalized.split_whitespace().collect::<Vec<_>>();
    if pieces.len() < 3 {
        return None;
    }
    let (year_text, month, day_text) = if let Some(month) = month_number(pieces[0]) {
        (pieces[2], month, pieces[1])
    } else if let Some(month) = month_number(pieces[1]) {
        (pieces[2], month, pieces[0])
    } else {
        return None;
    };
    Some((|| {
        let year = year_text
            .parse::<i64>()
            .map_err(|_| format!("invalid Date year: {year_text}"))?;
        let day = day_text
            .parse::<i64>()
            .map_err(|_| format!("invalid Date day: {day_text}"))?;
        let mut serial = strict_date_serial(year, i64::from(month), day, source)?;
        if pieces.len() > 3 {
            serial += parse_time_text(&pieces[3..].join(" "))?;
        }
        Ok(serial)
    })())
}

fn month_number(name: &str) -> Option<u32> {
    Some(match name.to_ascii_lowercase().as_str() {
        "jan" | "january" => 1,
        "feb" | "february" => 2,
        "mar" | "march" => 3,
        "apr" | "april" => 4,
        "may" => 5,
        "jun" | "june" => 6,
        "jul" | "july" => 7,
        "aug" | "august" => 8,
        "sep" | "sept" | "september" => 9,
        "oct" | "october" => 10,
        "nov" | "november" => 11,
        "dec" | "december" => 12,
        _ => return None,
    })
}

fn parse_time_text(source: &str) -> Result<f64, String> {
    let upper = source.trim().to_ascii_uppercase();
    let (clock, meridiem) = if let Some(clock) = upper.strip_suffix(" AM") {
        (clock, Some(false))
    } else if let Some(clock) = upper.strip_suffix(" PM") {
        (clock, Some(true))
    } else {
        (upper.as_str(), None)
    };
    let values = clock
        .split(':')
        .map(|part| {
            part.parse::<u32>()
                .map_err(|_| format!("invalid Time component: {part}"))
        })
        .collect::<Result<Vec<_>, _>>()?;
    if !(2..=3).contains(&values.len()) || values[1] > 59 || values.get(2).is_some_and(|v| *v > 59)
    {
        return Err(format!("invalid Time string: {source}"));
    }
    let mut hour = values[0];
    if let Some(pm) = meridiem {
        if !(1..=12).contains(&hour) {
            return Err(format!("invalid 12-hour Time: {source}"));
        }
        hour = hour % 12 + if pm { 12 } else { 0 };
    } else if hour > 23 {
        return Err(format!("invalid 24-hour Time: {source}"));
    }
    Ok((hour * 3_600 + values[1] * 60 + values.get(2).copied().unwrap_or(0)) as f64 / 86_400.0)
}

fn serial_date_parts(serial: f64) -> Result<DateParts, String> {
    let whole_days = serial.floor();
    if whole_days < i64::MIN as f64 || whole_days > i64::MAX as f64 {
        return Err("Date value is outside the supported range".to_string());
    }
    let days = ole_epoch_days()
        .checked_add(whole_days as i64)
        .ok_or_else(|| "Date value is outside the supported range".to_string())?;
    let (year, month, day) = civil_from_days(days);
    if !(100..=9_999).contains(&year) {
        return Err("Date year must be between 100 and 9999".to_string());
    }
    let seconds = ((serial - whole_days) * 86_400.0).round() as i64;
    if seconds == 86_400 {
        let (year, month, day) = civil_from_days(days + 1);
        return Ok(DateParts {
            year,
            month,
            day,
            hour: 0,
            minute: 0,
            second: 0,
        });
    }
    Ok(DateParts {
        year,
        month,
        day,
        hour: (seconds / 3_600) as u32,
        minute: ((seconds % 3_600) / 60) as u32,
        second: (seconds % 60) as u32,
    })
}

fn date_add(interval: &str, amount: i64, serial: f64) -> Result<f64, String> {
    let result = match interval {
        "yyyy" | "q" | "m" => {
            let parts = serial_date_parts(serial)?;
            let months = match interval {
                "yyyy" => amount.checked_mul(12),
                "q" => amount.checked_mul(3),
                _ => Some(amount),
            }
            .ok_or_else(|| "DateAdd month interval overflow".to_string())?;
            let total = parts.year * 12 + i64::from(parts.month) - 1 + months;
            let year = total.div_euclid(12);
            let month = total.rem_euclid(12) as u32 + 1;
            let day = parts.day.min(days_in_month(year, month));
            let date = date_serial(year, i64::from(month), i64::from(day))?;
            date + serial.rem_euclid(1.0)
        }
        "y" | "d" | "w" => serial + amount as f64,
        "ww" => serial + amount as f64 * 7.0,
        "h" => serial + amount as f64 / 24.0,
        "n" => serial + amount as f64 / 1_440.0,
        "s" => serial + amount as f64 / 86_400.0,
        _ => return Err(format!("unsupported Date interval: {interval}")),
    };
    serial_date_parts(result)?;
    Ok(result)
}

fn date_diff(interval: &str, first: f64, second: f64, first_day: i64) -> Result<i64, String> {
    let a = serial_date_parts(first)?;
    let b = serial_date_parts(second)?;
    Ok(match interval {
        "yyyy" => b.year - a.year,
        "q" => {
            (b.year * 4 + i64::from((b.month - 1) / 3))
                - (a.year * 4 + i64::from((a.month - 1) / 3))
        }
        "m" => (b.year * 12 + i64::from(b.month)) - (a.year * 12 + i64::from(a.month)),
        "y" | "d" => second.floor() as i64 - first.floor() as i64,
        "w" => (second.floor() as i64 - first.floor() as i64) / 7,
        "ww" => {
            week_boundary_index(second.floor() as i64, first_day)
                - week_boundary_index(first.floor() as i64, first_day)
        }
        "h" => (second * 24.0).floor() as i64 - (first * 24.0).floor() as i64,
        "n" => (second * 1_440.0).floor() as i64 - (first * 1_440.0).floor() as i64,
        "s" => (second * 86_400.0).round() as i64 - (first * 86_400.0).round() as i64,
        _ => return Err(format!("unsupported Date interval: {interval}")),
    })
}

fn date_part(interval: &str, serial: f64, first_day: i64, first_week: i64) -> Result<i64, String> {
    let parts = serial_date_parts(serial)?;
    let day_number = serial.floor() as i64;
    Ok(match interval {
        "yyyy" => parts.year,
        "q" => i64::from((parts.month - 1) / 3 + 1),
        "m" => i64::from(parts.month),
        "y" => {
            days_from_civil(parts.year, parts.month, parts.day) - days_from_civil(parts.year, 1, 1)
                + 1
        }
        "d" => i64::from(parts.day),
        "w" => weekday_number(day_number, first_day),
        "ww" => week_of_year(day_number, parts.year, first_day, first_week)?,
        "h" => i64::from(parts.hour),
        "n" => i64::from(parts.minute),
        "s" => i64::from(parts.second),
        _ => return Err(format!("unsupported Date interval: {interval}")),
    })
}

fn first_day_of_week(value: Option<&Value>, line: Option<u32>) -> Result<i64, RuntimeError> {
    let value = match value {
        None | Some(Value::Missing) => 1,
        Some(value) => integer_argument(value, line)?,
    };
    if !(0..=7).contains(&value) {
        return Err(invalid_procedure_call(
            "firstdayofweek must be between 0 and 7".to_string(),
            line,
        ));
    }
    Ok(if value == 0 { 1 } else { value })
}

fn first_week_of_year(value: Option<&Value>, line: Option<u32>) -> Result<i64, RuntimeError> {
    let value = match value {
        None | Some(Value::Missing) => 1,
        Some(value) => integer_argument(value, line)?,
    };
    if !(0..=3).contains(&value) {
        return Err(invalid_procedure_call(
            "firstweekofyear must be between 0 and 3".to_string(),
            line,
        ));
    }
    Ok(if value == 0 { 1 } else { value })
}

fn weekday_number(day_number: i64, first_day: i64) -> i64 {
    let sunday_based = (day_number + 6).rem_euclid(7) + 1;
    (sunday_based - first_day).rem_euclid(7) + 1
}

fn week_boundary_index(day_number: i64, first_day: i64) -> i64 {
    let anchor = (first_day - 1).rem_euclid(7) + 1;
    (day_number - anchor).div_euclid(7)
}

fn week_of_year(
    day_number: i64,
    year: i64,
    first_day: i64,
    first_week: i64,
) -> Result<i64, String> {
    let start = first_week_start(year, first_day, first_week);
    if day_number < start {
        return week_of_year(day_number, year - 1, first_day, first_week);
    }
    let next_start = first_week_start(year + 1, first_day, first_week);
    if day_number >= next_start {
        return Ok((day_number - next_start).div_euclid(7) + 1);
    }
    Ok((day_number - start).div_euclid(7) + 1)
}

fn first_week_start(year: i64, first_day: i64, first_week: i64) -> i64 {
    let january_first = days_from_civil(year, 1, 1) - ole_epoch_days();
    let offset = weekday_number(january_first, first_day) - 1;
    let containing_start = january_first - offset;
    match first_week {
        1 => containing_start,
        2 if offset <= 3 => containing_start,
        2 => containing_start + 7,
        3 if offset == 0 => containing_start,
        3 => containing_start + 7,
        _ => containing_start,
    }
}

fn days_in_month(year: i64, month: u32) -> u32 {
    match month {
        1 | 3 | 5 | 7 | 8 | 10 | 12 => 31,
        4 | 6 | 9 | 11 => 30,
        2 if year % 4 == 0 && (year % 100 != 0 || year % 400 == 0) => 29,
        2 => 28,
        _ => 0,
    }
}

fn ole_epoch_days() -> i64 {
    days_from_civil(1899, 12, 30)
}

fn days_from_civil(year: i64, month: u32, day: u32) -> i64 {
    let year = year - i64::from(month <= 2);
    let era = year.div_euclid(400);
    let year_of_era = year - era * 400;
    let adjusted_month = i64::from(month) + if month > 2 { -3 } else { 9 };
    let day_of_year = (153 * adjusted_month + 2) / 5 + i64::from(day) - 1;
    let day_of_era = year_of_era * 365 + year_of_era / 4 - year_of_era / 100 + day_of_year;
    era * 146_097 + day_of_era - 719_468
}

fn civil_from_days(days: i64) -> (i64, u32, u32) {
    let days = days + 719_468;
    let era = days.div_euclid(146_097);
    let day_of_era = days - era * 146_097;
    let year_of_era =
        (day_of_era - day_of_era / 1_460 + day_of_era / 36_524 - day_of_era / 146_096) / 365;
    let mut year = year_of_era + era * 400;
    let day_of_year = day_of_era - (365 * year_of_era + year_of_era / 4 - year_of_era / 100);
    let month_prime = (5 * day_of_year + 2) / 153;
    let day = day_of_year - (153 * month_prime + 2) / 5 + 1;
    let month = month_prime + if month_prime < 10 { 3 } else { -9 };
    year += i64::from(month <= 2);
    (year, month as u32, day as u32)
}

fn parse_val(source: &str) -> Result<f64, String> {
    let compact = source
        .chars()
        .filter(|character| !matches!(character, ' ' | '\t' | '\r' | '\n'))
        .collect::<String>();
    let bytes = compact.as_bytes();
    let mut cursor = 0;
    let negative = match bytes.first() {
        Some(b'+') => {
            cursor = 1;
            false
        }
        Some(b'-') => {
            cursor = 1;
            true
        }
        _ => false,
    };
    if bytes.get(cursor) == Some(&b'&') {
        let radix = match bytes.get(cursor + 1).map(u8::to_ascii_uppercase) {
            Some(b'H') => 16,
            Some(b'O') => 8,
            _ => return Ok(0.0),
        };
        let start = cursor + 2;
        let mut end = start;
        while bytes.get(end).is_some_and(|byte| match radix {
            16 => byte.is_ascii_hexdigit(),
            8 => matches!(byte, b'0'..=b'7'),
            _ => false,
        }) {
            end += 1;
        }
        if end == start {
            return Ok(0.0);
        }
        let digits = &compact[start..end];
        let raw = u32::from_str_radix(digits, radix)
            .map_err(|_| "Val radix literal is outside the supported Long range".to_string())?;
        let signed = if negative {
            -(i64::from(raw))
        } else if (radix == 16 && digits.len() <= 4) || (radix == 8 && digits.len() <= 6) {
            i64::from(raw as u16 as i16)
        } else {
            i64::from(raw as i32)
        };
        return Ok(signed as f64);
    }

    while bytes.get(cursor).is_some_and(u8::is_ascii_digit) {
        cursor += 1;
    }
    if bytes.get(cursor) == Some(&b'.') {
        cursor += 1;
        while bytes.get(cursor).is_some_and(u8::is_ascii_digit) {
            cursor += 1;
        }
    }
    if cursor == usize::from(matches!(bytes.first(), Some(b'+') | Some(b'-')))
        || (cursor == 1 && bytes.first() == Some(&b'.'))
        || (cursor == 2
            && matches!(bytes.first(), Some(b'+') | Some(b'-'))
            && bytes.get(1) == Some(&b'.'))
    {
        return Ok(0.0);
    }
    if matches!(
        bytes.get(cursor).map(u8::to_ascii_uppercase),
        Some(b'E' | b'D')
    ) {
        let exponent_start = cursor;
        cursor += 1;
        if matches!(bytes.get(cursor), Some(b'+') | Some(b'-')) {
            cursor += 1;
        }
        let digits_start = cursor;
        while bytes.get(cursor).is_some_and(u8::is_ascii_digit) {
            cursor += 1;
        }
        if cursor == digits_start {
            cursor = exponent_start;
        }
    }
    compact[..cursor]
        .replace(['d', 'D'], "E")
        .parse::<f64>()
        .map_err(|_| "Val could not convert the numeric prefix".to_string())
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
        "vbusesystem" => Value::Integer(0),
        "vbsunday" => Value::Integer(1),
        "vbmonday" => Value::Integer(2),
        "vbtuesday" => Value::Integer(3),
        "vbwednesday" => Value::Integer(4),
        "vbthursday" => Value::Integer(5),
        "vbfriday" => Value::Integer(6),
        "vbsaturday" => Value::Integer(7),
        "vbfirstjan1" => Value::Integer(1),
        "vbfirstfourdays" => Value::Integer(2),
        "vbfirstfullweek" => Value::Integer(3),
        "vbempty" => Value::Integer(0),
        "vbnull" => Value::Integer(1),
        "vbinteger" => Value::Integer(2),
        "vblong" => Value::Integer(3),
        "vbsingle" => Value::Integer(4),
        "vbdouble" => Value::Integer(5),
        "vbcurrency" => Value::Integer(6),
        "vbdate" => Value::Integer(7),
        "vbstring" => Value::Integer(8),
        "vbobject" => Value::Integer(9),
        "vberror" => Value::Integer(10),
        "vbboolean" => Value::Integer(11),
        "vbvariant" => Value::Integer(12),
        "vbdataobject" => Value::Integer(13),
        "vbdecimal" => Value::Integer(14),
        "vbbyte" => Value::Integer(17),
        "vbuserdefinedtype" => Value::Integer(36),
        "vbarray" => Value::Integer(8_192),
        "vbuppercase" => Value::Integer(1),
        "vblowercase" => Value::Integer(2),
        "vbpropercase" => Value::Integer(3),
        "vbwide" => Value::Integer(4),
        "vbnarrow" => Value::Integer(8),
        "vbkatakana" => Value::Integer(16),
        "vbhiragana" => Value::Integer(32),
        "vbunicode" => Value::Integer(64),
        "vbfromunicode" => Value::Integer(128),
        "vbendofperiod" => Value::Integer(0),
        "vbbeginningofperiod" => Value::Integer(1),
        "vbgeneraldate" => Value::Integer(0),
        "vblongdate" => Value::Integer(1),
        "vbshortdate" => Value::Integer(2),
        "vblongtime" => Value::Integer(3),
        "vbshorttime" => Value::Integer(4),
        "vbusedefault" => Value::Integer(-2),
        "vbtrue" => Value::Integer(-1),
        "vbfalse" => Value::Integer(0),
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

fn convert_integer(
    value: &Value,
    minimum: i64,
    maximum: i64,
    line: Option<u32>,
) -> Result<i64, RuntimeError> {
    let value = number(value)
        .map_err(|message| error(RuntimeErrorKind::TypeMismatch, message, line))?
        .round_ties_even();
    if !value.is_finite() || value < minimum as f64 || value > maximum as f64 {
        Err(error(
            RuntimeErrorKind::Overflow,
            "numeric conversion overflow",
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

#[cfg(not(target_arch = "wasm32"))]
fn default_current_time() -> f64 {
    use std::time::{SystemTime, UNIX_EPOCH};

    SystemTime::now()
        .duration_since(UNIX_EPOCH)
        .map(|duration| duration.as_secs_f64() / 86_400.0 + 25_569.0)
        .unwrap_or(0.0)
}

#[cfg(target_arch = "wasm32")]
fn default_current_time() -> f64 {
    0.0
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

/// Narrows a value to a declared VBA type, the way `Function Total() As Long`
/// hands back a Long whatever it was given. Whole types round half to even, so
/// both 42.5 and 43.5 land on an even number, and a value too large for the
/// type is the overflow VBA reports rather than a wrapped number.
///
/// Types this runtime holds no distinct value for — Currency, Date, Decimal,
/// Variant and the object types — pass through untouched.
fn coerce_declared(value: Value, declared: &str, line: u32) -> Result<Value, RuntimeError> {
    let (low, high) = match declared.to_ascii_lowercase().as_str() {
        "byte" => (0.0, 255.0),
        "integer" => (-32_768.0, 32_767.0),
        "long" => (-2_147_483_648.0, 2_147_483_647.0),
        "longlong" | "longptr" => (i64::MIN as f64, i64::MAX as f64),
        "single" => {
            let number = coerce_number(&value, declared, line)?;
            let narrowed = number as f32;
            if !narrowed.is_finite() && number.is_finite() {
                return Err(overflow(declared, line));
            }
            return Ok(Value::Double(narrowed as f64));
        }
        "double" => return Ok(Value::Double(coerce_number(&value, declared, line)?)),
        "boolean" => {
            return Ok(Value::Boolean(match &value {
                Value::Boolean(value) => *value,
                Value::String(_) | Value::Integer(_) | Value::Double(_) => {
                    coerce_number(&value, declared, line)? != 0.0
                }
                _ => return Ok(value),
            }))
        }
        "string" => {
            return match &value {
                Value::Empty => Ok(Value::String(String::new())),
                Value::Object(_) | Value::Nothing | Value::Null => Ok(value),
                value => text(value)
                    .map(Value::String)
                    .map_err(|message| error(RuntimeErrorKind::TypeMismatch, message, Some(line))),
            }
        }
        _ => return Ok(value),
    };
    if matches!(value, Value::Object(_) | Value::Nothing | Value::Null) {
        return Ok(value);
    }
    let number = coerce_number(&value, declared, line)?.round_ties_even();
    if !(low..=high).contains(&number) {
        return Err(overflow(declared, line));
    }
    Ok(Value::Integer(number as i64))
}

fn coerce_number(value: &Value, declared: &str, line: u32) -> Result<f64, RuntimeError> {
    number(value).map_err(|_| {
        error(
            RuntimeErrorKind::TypeMismatch,
            format!(
                "type mismatch storing {} in {declared}",
                value_type_name(value)
            ),
            Some(line),
        )
    })
}

fn overflow(declared: &str, line: u32) -> RuntimeError {
    error(
        RuntimeErrorKind::Overflow,
        format!("value is outside the range of {declared}"),
        Some(line),
    )
}

fn default_value(type_name: &TypeName) -> Value {
    match type_name.name.to_ascii_lowercase().as_str() {
        "boolean" => Value::Boolean(false),
        "byte" | "integer" | "long" | "longlong" | "longptr" | "currency" => Value::Integer(0),
        "single" | "double" | "decimal" => Value::Double(0.0),
        "date" => Value::Double(0.0),
        "string" => Value::String(String::new()),
        "object"
        | "application"
        | "workbook"
        | "worksheet"
        | "range"
        | "chart"
        | "shape"
        | "collection"
        | "dictionary"
        | "scripting.dictionary" => Value::Nothing,
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
        Value::Error(_) => "Error".to_string(),
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
        Value::Error(_) => 10,
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
        Value::Error(value) => Ok(*value as f64),
        Value::String(value) => value
            .trim()
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
        Value::Error(_) => Err("type mismatch converting Error to Boolean".to_string()),
        Value::Object(_) => Err("type mismatch converting object to Boolean".to_string()),
        Value::Missing => Err("invalid use of Missing".to_string()),
        Value::Nothing => Err("object variable or With block variable not set".to_string()),
    }
}

fn unary(op: UnaryOp, value: Value) -> Result<Value, String> {
    if matches!(value, Value::Error(_)) {
        return Err("type mismatch using Error value as an operand".to_string());
    }
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
        Value::Array(_) | Value::Object(_) | Value::Error(_) | Value::Missing | Value::Nothing
    ) || matches!(
        rhs,
        Value::Array(_) | Value::Object(_) | Value::Error(_) | Value::Missing | Value::Nothing
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
        Value::Error(value) => format!("Error {value}"),
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

    /// An object standing in a value context reads as its default member, which
    /// is how `Range("A1") + 1` and `"x" & Range("A1")` work in VBA. Identity
    /// and type questions still see the object itself.
    #[test]
    fn objects_stand_for_their_default_member_in_value_contexts() {
        let module = parse_module(
            "Public Function DefaultMember() As String\n\
               Dim cell As Object\n\
               Range(\"A1\").Value = 42\n\
               Range(\"A2\").Value = \"text\"\n\
               Set cell = Range(\"A1\")\n\
               DefaultMember = (\"x\" & cell) & \"|\" & (cell + 1) & \"|\" & (cell = 42)\n\
               DefaultMember = DefaultMember & \"|\" & (cell > 40) & \"|\" & (-cell) & \"|\" & (\"y\" & Range(\"A2\"))\n\
               DefaultMember = DefaultMember & \"|\" & TypeName(cell) & \"|\" & IsObject(cell) & \"|\" & (cell Is Range(\"A1\"))\n\
             End Function\n",
        )
        .unwrap();
        let mut host = SheetHost::default();
        let value = execute_with_host(&module, "DefaultMember", vec![], &mut host).unwrap();

        assert_eq!(
            value,
            Value::String("x42|43|True|True|-42|ytext|Cell|True|True".to_string())
        );
    }

    /// Builtins that read values see through an object to its default member,
    /// while the ones asking about the argument itself do not. Measured in VBA
    /// against a cell holding 42, one holding text, and an empty one: VarType
    /// answered 5, 8 and 0 rather than 9 for the object. It reads 3 rather than
    /// 5 here because this host keeps the Integer it was handed, where every
    /// number in an Excel cell is a Double.
    #[test]
    fn builtins_read_values_through_an_objects_default_member() {
        let module = parse_module(
            "Public Function ReadThrough() As String\n\
               Dim number As Object\n\
               Dim word As Object\n\
               Dim blank As Object\n\
               Range(\"A1\").Value = 42\n\
               Range(\"A3\").Value = \"text\"\n\
               Set number = Range(\"A1\")\n\
               Set word = Range(\"A3\")\n\
               Set blank = Range(\"A2\")\n\
               ReadThrough = VarType(number) & \"|\" & VarType(word) & \"|\" & VarType(blank)\n\
               ReadThrough = ReadThrough & \"|\" & Len(word) & \"|\" & UCase(word) & \"|\" & CStr(number)\n\
               ReadThrough = ReadThrough & \"|\" & Abs(number) & \"|\" & IsNumeric(number) & \"|\" & IsNumeric(word)\n\
               ReadThrough = ReadThrough & \"|\" & IsEmpty(blank) & \"|\" & TypeName(number) & \"|\" & IsObject(number)\n\
             End Function\n",
        )
        .unwrap();
        let mut host = SheetHost::default();
        let value = execute_with_host(&module, "ReadThrough", vec![], &mut host).unwrap();

        assert_eq!(
            value,
            Value::String(
                "3|8|0|4|TEXT|42|42|True|False|True|Cell|True".to_string()
            )
        );
    }

    /// A typed function narrows whatever it was given. Measured in VBA: whole
    /// types round half to even, so 42.5 falls to 42 while 43.5 climbs to 44.
    #[test]
    fn a_functions_result_takes_its_declared_type() {
        for (declared, assigned, expected) in [
            ("Long", "42.4", Value::Integer(42)),
            ("Long", "42.5", Value::Integer(42)),
            ("Long", "43.5", Value::Integer(44)),
            ("Long", "42.6", Value::Integer(43)),
            ("Long", "-42.5", Value::Integer(-42)),
            ("Integer", "42.5", Value::Integer(42)),
            ("Long", "\"17\"", Value::Integer(17)),
            ("Double", "42", Value::Double(42.0)),
            ("Single", "0.5", Value::Double(0.5)),
            ("String", "42", Value::String("42".to_string())),
            ("Boolean", "3", Value::Boolean(true)),
            ("Boolean", "0", Value::Boolean(false)),
            ("Variant", "42", Value::Integer(42)),
        ] {
            let source =
                format!("Public Function Narrow() As {declared}\n  Narrow = {assigned}\nEnd Function\n");
            let module = parse_module(&source).unwrap();
            assert_eq!(
                execute(&module, "Narrow", vec![]).unwrap(),
                expected,
                "{declared} holding {assigned}"
            );
        }
    }

    /// A declared variable narrows what it is given, on every assignment rather
    /// than only where it was declared. Measured in VBA.
    #[test]
    fn a_declared_variable_narrows_what_it_is_given() {
        let module = parse_module(
            "Public Function Narrow() As String\n\
               Dim whole As Long\n\
               Dim word As String\n\
               Dim flag As Boolean\n\
               Dim loose As Variant\n\
               whole = 42.6\n\
               Narrow = whole & \"|\"\n\
               whole = 42.5\n\
               Narrow = Narrow & whole & \"|\"\n\
               whole = 43.5\n\
               Narrow = Narrow & whole & \"|\"\n\
               whole = \"17\"\n\
               Narrow = Narrow & whole & \"|\"\n\
               word = 42\n\
               flag = 3\n\
               loose = 42.6\n\
               Narrow = Narrow & TypeName(whole) & \"|\" & TypeName(word) & \"|\" & word\n\
               Narrow = Narrow & \"|\" & flag & \"|\" & loose & \"|\" & TypeName(loose)\n\
             End Function\n",
        )
        .unwrap();
        assert_eq!(
            execute(&module, "Narrow", vec![]).unwrap(),
            Value::String("43|42|44|17|Long|String|42|True|42.6|Double".to_string())
        );
    }

    /// A ByVal parameter narrows its argument, including the value standing in
    /// for an omitted optional one.
    #[test]
    fn a_byval_parameter_narrows_its_argument() {
        let module = parse_module(
            "Public Function Narrow() As String\n\
               Narrow = TakesLong(42.6) & \"|\" & TakesLong(42.5) & \"|\" & TakesLong(\"17\")\n\
               Narrow = Narrow & \"|\" & TakesOptional() & \"|\" & TakesOptional(9.7)\n\
               Narrow = Narrow & \"|\" & TakesString(42) & \"|\" & TakesLoose(42.6)\n\
             End Function\n\
             Private Function TakesLong(ByVal n As Long) As String\n\
               TakesLong = TypeName(n) & \":\" & n\n\
             End Function\n\
             Private Function TakesOptional(Optional ByVal n As Long = 5) As String\n\
               TakesOptional = TypeName(n) & \":\" & n\n\
             End Function\n\
             Private Function TakesString(ByVal s As String) As String\n\
               TakesString = TypeName(s) & \":\" & s\n\
             End Function\n\
             Private Function TakesLoose(ByVal v As Variant) As String\n\
               TakesLoose = TypeName(v) & \":\" & v\n\
             End Function\n",
        )
        .unwrap();
        assert_eq!(
            execute(&module, "Narrow", vec![]).unwrap(),
            Value::String(
                "Long:43|Long:42|Long:17|Long:5|Long:10|String:42|Double:42.6".to_string()
            )
        );
    }

    #[test]
    fn a_declared_variable_refuses_what_will_not_fit() {
        for (declared, assigned, kind) in [
            ("Integer", "40000", RuntimeErrorKind::Overflow),
            ("Byte", "300", RuntimeErrorKind::Overflow),
            ("Long", "\"abc\"", RuntimeErrorKind::TypeMismatch),
        ] {
            let source = format!(
                "Public Sub Store()\n  Dim slot As {declared}\n  slot = {assigned}\nEnd Sub\n"
            );
            let module = parse_module(&source).unwrap();
            let error = execute(&module, "Store", vec![])
                .expect_err(&format!("{declared} cannot hold {assigned}"));
            assert_eq!(error.kind, kind, "{declared} holding {assigned}");
        }
    }

    #[test]
    fn a_declared_type_refuses_what_will_not_fit() {
        for (declared, assigned, kind) in [
            ("Integer", "40000", RuntimeErrorKind::Overflow),
            ("Byte", "300", RuntimeErrorKind::Overflow),
            ("Long", "3000000000", RuntimeErrorKind::Overflow),
            ("Long", "\"abc\"", RuntimeErrorKind::TypeMismatch),
        ] {
            let source =
                format!("Public Function Narrow() As {declared}\n  Narrow = {assigned}\nEnd Function\n");
            let module = parse_module(&source).unwrap();
            let error = execute(&module, "Narrow", vec![])
                .expect_err(&format!("{declared} cannot hold {assigned}"));
            assert_eq!(error.kind, kind, "{declared} holding {assigned}");
        }
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
            // VarType reports the type behind an object's default member, so an
            // unset cell reads as vbEmpty. Nothing, having no member to read,
            // stays vbObject.
            Value::String("True|True|False|True|True|Cell|0|True|True|Nothing".to_string())
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
    fn dictionary_supports_createobject_default_item_arrays_and_enumeration() {
        let value = run(
            "Public Function DictionaryProbe() As String\n\
               Dim values As Object\n\
               Dim entryKey As Variant\n\
               Dim keys As Variant\n\
               Dim items As Variant\n\
               Dim walked As String\n\
               Set values = CreateObject(\"Scripting.Dictionary\")\n\
               values.CompareMode = vbTextCompare\n\
               values.Add \"Alpha\", 10\n\
               values(\"beta\") = 20\n\
               values.Item(\"ALPHA\") = 11\n\
               For Each entryKey In values\n\
                 walked = walked & entryKey & \",\"\n\
               Next\n\
               keys = values.Keys\n\
               items = values.Items\n\
               DictionaryProbe = values.Count & \"|\" & values.Exists(\"BETA\") & \"|\" & values(\"alpha\") & \"|\" & keys(0) & \"|\" & keys(1) & \"|\" & items(0) & \"|\" & items(1) & \"|\" & walked & \"|\" & TypeName(values) & \"|\" & (TypeOf values Is Scripting.Dictionary)\n\
               values.Remove \"BETA\"\n\
               DictionaryProbe = DictionaryProbe & \"|\" & values.Count\n\
             End Function\n",
            "DictionaryProbe",
            vec![],
        )
        .unwrap();

        assert_eq!(
            value,
            Value::String("2|True|11|Alpha|beta|11|20|Alpha,beta,|Dictionary|True|1".to_string())
        );
    }

    #[test]
    fn dictionary_reports_vba_errors_and_removeall_resets_comparemode() {
        let value = run(
            "Public Function DictionaryErrors() As String\n\
               Dim values As New Scripting.Dictionary\n\
               Dim duplicate As Long\n\
               Dim modeChange As Long\n\
               Dim missing As Long\n\
               values.CompareMode = vbTextCompare\n\
               values.Add \"key\", 1\n\
               On Error Resume Next\n\
               values.Add \"KEY\", 2\n\
               duplicate = Err.Number\n\
               Err.Clear\n\
               values.CompareMode = vbBinaryCompare\n\
               modeChange = Err.Number\n\
               Err.Clear\n\
               values.Remove \"absent\"\n\
               missing = Err.Number\n\
               Err.Clear\n\
               values.RemoveAll\n\
               values.CompareMode = vbBinaryCompare\n\
               values(\"x\") = 3\n\
               DictionaryErrors = duplicate & \"|\" & modeChange & \"|\" & missing & \"|\" & values.Count & \"|\" & values(\"x\")\n\
             End Function\n",
            "DictionaryErrors",
            vec![],
        )
        .unwrap();

        assert_eq!(value, Value::String("457|5|32811|1|3".to_string()));
    }

    #[test]
    fn dictionary_key_rename_and_use_compare_option_follow_module_settings() {
        let value = run(
            "Option Compare Text\n\
             Public Function DictionaryRename() As String\n\
               Dim values As New Scripting.Dictionary\n\
               values.CompareMode = vbUseCompareOption\n\
               values.Add \"Alpha\", 42\n\
               values.Key(\"ALPHA\") = \"Renamed\"\n\
               DictionaryRename = values.CompareMode & \"|\" & values.Exists(\"RENAMED\") & \"|\" & values.Exists(\"alpha\") & \"|\" & values(\"renamed\")\n\
             End Function\n",
            "DictionaryRename",
            vec![],
        )
        .unwrap();

        assert_eq!(value, Value::String("-1|True|False|42".to_string()));
    }

    #[test]
    fn dictionary_as_new_reactivates_and_unknown_createobject_reports_429() {
        let value = run(
            "Public Function DictionaryLifetime() As String\n\
               Dim values As New Scripting.Dictionary\n\
               Dim createFailure As Long\n\
               values(\"first\") = 1\n\
               Set values = Nothing\n\
               values(\"second\") = 2\n\
               On Error Resume Next\n\
               Set values = CreateObject(\"Oxi.Unknown\")\n\
               createFailure = Err.Number\n\
               DictionaryLifetime = values.Count & \"|\" & values(\"second\") & \"|\" & createFailure\n\
             End Function\n",
            "DictionaryLifetime",
            vec![],
        )
        .unwrap();

        assert_eq!(value, Value::String("1|2|429".to_string()));
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
    fn executes_vba_numeric_type_conversion_functions() {
        let value = run(
            "Public Function ConversionBuiltins() As String\n\
               ConversionBuiltins = CByte(125.5678) & \"|\" & CInt(2344.5) & \"|\" & CInt(2345.5) & \"|\"\n\
               ConversionBuiltins = ConversionBuiltins & CCur(543.214588 * 2) & \"|\" & Round(CSng(75.3421115), 5) & \"|\"\n\
               ConversionBuiltins = ConversionBuiltins & CLngLng(2147483648#) & \"|\" & CLngPtr(42) & \"|\" & CDec(12.5) & \"|\" & CVar(\"variant\")\n\
             End Function\n",
            "ConversionBuiltins",
            vec![],
        )
        .unwrap();

        assert_eq!(
            value,
            Value::String(
                "126|2344|2346|1086.4292|75.34211|2147483648|42|12.5|variant".to_string()
            )
        );
    }

    #[test]
    fn exposes_vba_variant_type_constants() {
        let value = run(
            "Public Function VariantConstants() As String\n\
               VariantConstants = vbEmpty & \"|\" & vbNull & \"|\" & vbInteger & \"|\" & vbLong & \"|\" & vbSingle & \"|\" & vbDouble & \"|\" & vbCurrency & \"|\" & vbDate & \"|\" & vbString & \"|\" & vbObject & \"|\" & vbError & \"|\" & vbBoolean & \"|\" & vbVariant & \"|\" & vbDecimal & \"|\" & vbByte & \"|\" & vbArray\n\
             End Function\n",
            "VariantConstants",
            vec![],
        )
        .unwrap();

        assert_eq!(
            value,
            Value::String("0|1|2|3|4|5|6|7|8|9|10|11|12|14|17|8192".to_string())
        );
    }

    #[test]
    fn conversion_functions_raise_overflow_errors_for_out_of_range_values() {
        let value = run(
            "Public Function ConversionErrors() As String\n\
               Dim byteError As Long\n\
               Dim integerError As Long\n\
               Dim currencyError As Long\n\
               Dim singleError As Long\n\
               On Error Resume Next\n\
               byteError = CByte(-1)\n\
               byteError = Err.Number\n\
               Err.Clear\n\
               integerError = CInt(32768)\n\
               integerError = Err.Number\n\
               Err.Clear\n\
               currencyError = CCur(1E20)\n\
               currencyError = Err.Number\n\
               Err.Clear\n\
               singleError = CSng(1E40)\n\
               singleError = Err.Number\n\
               ConversionErrors = byteError & \"|\" & integerError & \"|\" & currencyError & \"|\" & singleError\n\
             End Function\n",
            "ConversionErrors",
            vec![],
        )
        .unwrap();

        assert_eq!(value, Value::String("6|6|6|6".to_string()));
    }

    #[test]
    fn executes_str_and_unicode_strconv_modes() {
        let value = run(
            "Public Function TextConversions() As String\n\
               TextConversions = \"[\" & Str(459) & \"]|[\" & Str(-459.65) & \"]|\"\n\
               TextConversions = TextConversions & StrConv(\"Oxi VBA\", vbUpperCase) & \"|\" & StrConv(\"Oxi VBA\", vbLowerCase) & \"|\" & StrConv(\"oXI-vBA runtime\", vbProperCase) & \"|\"\n\
               TextConversions = TextConversions & StrConv(\"ABC 123 ｶﾞ\", vbWide) & \"|\" & StrConv(\"ＡＢＣ　１２３　ガ\", vbNarrow) & \"|\"\n\
               TextConversions = TextConversions & StrConv(\"おーぷん\", vbKatakana) & \"|\" & StrConv(\"オープン\", vbHiragana)\n\
             End Function\n",
            "TextConversions",
            vec![],
        )
        .unwrap();

        assert_eq!(
            value,
            Value::String(
                "[ 459]|[-459.65]|OXI VBA|oxi vba|Oxi-Vba Runtime|ＡＢＣ　１２３　ガ|ABC 123 ｶﾞ|オープン|おーぷん"
                    .to_string()
            )
        );
    }

    #[test]
    fn executes_monthname_and_weekdayname() {
        let value = run(
            "Public Function CalendarNames() As String\n\
               CalendarNames = MonthName(2) & \"|\" & MonthName(2, True) & \"|\"\n\
               CalendarNames = CalendarNames & WeekdayName(1) & \"|\" & WeekdayName(1, True, vbMonday) & \"|\" & WeekdayName(7, False, vbMonday)\n\
             End Function\n",
            "CalendarNames",
            vec![],
        )
        .unwrap();

        assert_eq!(
            value,
            Value::String("February|Feb|Sunday|Mon|Sunday".to_string())
        );
    }

    #[test]
    fn text_conversion_functions_report_invalid_arguments_as_vba_errors() {
        let value = run(
            "Public Function TextConversionErrors() As String\n\
               Dim monthError As Long\n\
               Dim weekdayError As Long\n\
               Dim modeError As Long\n\
               Dim nullError As Long\n\
               On Error Resume Next\n\
               monthError = MonthName(0)\n\
               monthError = Err.Number\n\
               Err.Clear\n\
               weekdayError = WeekdayName(8)\n\
               weekdayError = Err.Number\n\
               Err.Clear\n\
               modeError = StrConv(\"x\", vbWide + vbNarrow)\n\
               modeError = Err.Number\n\
               Err.Clear\n\
               nullError = StrConv(Null, vbUpperCase)\n\
               nullError = Err.Number\n\
               TextConversionErrors = monthError & \"|\" & weekdayError & \"|\" & modeError & \"|\" & nullError\n\
             End Function\n",
            "TextConversionErrors",
            vec![],
        )
        .unwrap();

        assert_eq!(value, Value::String("5|5|5|13".to_string()));
    }

    #[test]
    fn executes_vba_numeric_rounding_radix_and_val_functions() {
        let value = run(
            "Public Function NumericBuiltins() As String\n\
               NumericBuiltins = Round(2.5) & \"|\" & Round(3.5) & \"|\" & Round(0.12345, 4) & \"|\"\n\
               NumericBuiltins = NumericBuiltins & Int(-8.4) & \"|\" & Fix(-8.4) & \"|\" & Sgn(-10) & \"|\" & Sgn(0) & \"|\" & Sgn(10) & \"|\" & Sqr(81) & \"|\"\n\
               NumericBuiltins = NumericBuiltins & Hex(459) & \"|\" & Hex(-1) & \"|\" & Oct(459) & \"|\" & Oct(-1) & \"|\"\n\
               NumericBuiltins = NumericBuiltins & Val(\"&HFFFF\") & \"|\" & Val(\" 16 15 198th Street\") & \"|\" & Val(\"-12.5E2x\")\n\
             End Function\n",
            "NumericBuiltins",
            vec![],
        )
        .unwrap();

        assert_eq!(
            value,
            Value::String(
                "2|4|0.1234|-9|-8|-1|0|1|9|1CB|FFFFFFFF|713|37777777777|-1|1615198|-1250"
                    .to_string()
            )
        );
    }

    #[test]
    fn executes_eager_iif_choose_and_switch_selection_functions() {
        let value = run(
            "Private hits As Long\n\
             Private Function Mark(value As Long) As Long\n\
               hits = hits + value\n\
               Mark = value\n\
             End Function\n\
             Public Function SelectionBuiltins() As String\n\
               Dim selected As Long\n\
               selected = IIf(True, Mark(1), Mark(2))\n\
               SelectionBuiltins = selected & \"|\" & hits & \"|\" & Choose(2, \"one\", \"two\", \"three\") & \"|\" & IsNull(Choose(0, 1, 2)) & \"|\"\n\
               SelectionBuiltins = SelectionBuiltins & Switch(False, \"no\", 2 > 1, \"yes\", True, \"late\") & \"|\" & IsNull(Switch(False, 1, False, 2))\n\
             End Function\n",
            "SelectionBuiltins",
            vec![],
        )
        .unwrap();

        assert_eq!(value, Value::String("1|3|two|True|yes|True".to_string()));
    }

    #[test]
    fn numeric_builtins_raise_vba_errors_for_invalid_domains() {
        let value = run(
            "Public Function NumericErrors() As String\n\
               Dim squareRoot As Long\n\
               Dim decimalPlaces As Long\n\
               Dim nullVal As Long\n\
               On Error Resume Next\n\
               squareRoot = Sqr(-1)\n\
               squareRoot = Err.Number\n\
               Err.Clear\n\
               decimalPlaces = Round(1.2, -1)\n\
               decimalPlaces = Err.Number\n\
               Err.Clear\n\
               nullVal = Val(Null)\n\
               nullVal = Err.Number\n\
               NumericErrors = squareRoot & \"|\" & decimalPlaces & \"|\" & nullVal\n\
             End Function\n",
            "NumericErrors",
            vec![],
        )
        .unwrap();

        assert_eq!(value, Value::String("5|5|13".to_string()));
    }

    #[test]
    fn executes_vba_annuity_financial_functions() {
        let value = run(
            "Public Function FinancialBuiltins() As String\n\
               Dim payment As Double\n\
               payment = Pmt(0.1 / 12, 48, 10000, 0, vbEndOfPeriod)\n\
               FinancialBuiltins = Round(payment, 2) & \"|\" & Round(PV(0.1 / 12, 48, payment), 2) & \"|\" & Round(FV(0.1 / 12, 48, payment, 10000), 2) & \"|\"\n\
               FinancialBuiltins = FinancialBuiltins & Round(NPer(0.1 / 12, payment, 10000), 6) & \"|\" & Round(Rate(48, payment, 10000) * 12, 8) & \"|\"\n\
               FinancialBuiltins = FinancialBuiltins & Round(IPmt(0.1 / 12, 1, 48, 10000), 2) & \"|\" & Round(PPmt(0.1 / 12, 1, 48, 10000), 2)\n\
             End Function\n",
            "FinancialBuiltins",
            vec![],
        )
        .unwrap();

        assert_eq!(
            value,
            Value::String("-253.63|10000|0|48|0.1|-83.33|-170.29".to_string())
        );
    }

    #[test]
    fn financial_functions_handle_zero_rate_and_beginning_payments() {
        let value = run(
            "Public Function FinancialEdges() As String\n\
               FinancialEdges = Pmt(0, 10, 1000) & \"|\" & PV(0, 10, -100) & \"|\" & FV(0, 10, -100, 1000) & \"|\" & NPer(0, -100, 1000) & \"|\"\n\
               FinancialEdges = FinancialEdges & IPmt(0.01, 1, 12, 1000, 0, vbBeginningOfPeriod) & \"|\" & Round(Pmt(0.01, 12, 1000, 0, vbBeginningOfPeriod), 6)\n\
             End Function\n",
            "FinancialEdges",
            vec![],
        )
        .unwrap();

        assert_eq!(
            value,
            Value::String("-100|1000|0|10|0|-87.969098".to_string())
        );
    }

    #[test]
    fn financial_functions_report_invalid_domains() {
        let value = run(
            "Public Function FinancialErrors() As String\n\
               Dim paymentType As Long\n\
               Dim zeroPeriods As Long\n\
               Dim badPeriod As Long\n\
               Dim noRate As Long\n\
               On Error Resume Next\n\
               paymentType = PV(0.1, 10, -100, 0, 2)\n\
               paymentType = Err.Number\n\
               Err.Clear\n\
               zeroPeriods = Pmt(0.1, 0, 1000)\n\
               zeroPeriods = Err.Number\n\
               Err.Clear\n\
               badPeriod = IPmt(0.1, 0, 10, 1000)\n\
               badPeriod = Err.Number\n\
               Err.Clear\n\
               noRate = Rate(10, 0, 1000, 1000)\n\
               noRate = Err.Number\n\
               FinancialErrors = paymentType & \"|\" & zeroPeriods & \"|\" & badPeriod & \"|\" & noRate\n\
             End Function\n",
            "FinancialErrors",
            vec![],
        )
        .unwrap();

        assert_eq!(value, Value::String("5|5|5|5".to_string()));
    }

    #[test]
    fn executes_vba_depreciation_functions() {
        let value = run(
            "Public Function DepreciationBuiltins() As String\n\
               DepreciationBuiltins = DDB(1000, 100, 5, 1) & \"|\" & DDB(1000, 100, 5, 2) & \"|\" & Round(DDB(1000, 100, 5, 5), 2) & \"|\"\n\
               DepreciationBuiltins = DepreciationBuiltins & SLN(1000, 100, 5) & \"|\" & SYD(1000, 100, 5, 2)\n\
             End Function\n",
            "DepreciationBuiltins",
            vec![],
        )
        .unwrap();

        assert_eq!(value, Value::String("400|240|29.6|180|240".to_string()));
    }

    #[test]
    fn executes_vba_cash_flow_financial_functions() {
        let value = run(
            "Public Function CashFlowBuiltins() As String\n\
               Dim values(4) As Double\n\
               values(0) = -70000\n\
               values(1) = 22000\n\
               values(2) = 25000\n\
               values(3) = 28000\n\
               values(4) = 31000\n\
               CashFlowBuiltins = Round(IRR(values), 8) & \"|\" & Round(MIRR(values, 0.1, 0.12), 8)\n\
             End Function\n",
            "CashFlowBuiltins",
            vec![],
        )
        .unwrap();

        assert_eq!(value, Value::String("0.17743588|0.15512706".to_string()));
    }

    #[test]
    fn remaining_financial_functions_report_invalid_inputs() {
        let value = run(
            "Public Function FinancialInputErrors() As String\n\
               Dim flows(1) As Double\n\
               Dim badPeriod As Long\n\
               Dim badFlows As Long\n\
               Dim wrongType As Long\n\
               flows(0) = 100\n\
               flows(1) = 200\n\
               On Error Resume Next\n\
               badPeriod = SYD(1000, 100, 5, 6)\n\
               badPeriod = Err.Number\n\
               Err.Clear\n\
               badFlows = IRR(flows)\n\
               badFlows = Err.Number\n\
               Err.Clear\n\
               wrongType = MIRR(42, 0.1, 0.12)\n\
               wrongType = Err.Number\n\
               FinancialInputErrors = badPeriod & \"|\" & badFlows & \"|\" & wrongType\n\
             End Function\n",
            "FinancialInputErrors",
            vec![],
        )
        .unwrap();

        assert_eq!(value, Value::String("5|5|13".to_string()));
    }

    #[test]
    fn executes_vba_npv_partition_and_qbcolor() {
        let value = run(
            "Public Function RemainingPureBuiltins() As String\n\
               Dim values(4) As Double\n\
               values(0) = -70000\n\
               values(1) = 22000\n\
               values(2) = 25000\n\
               values(3) = 28000\n\
               values(4) = 31000\n\
               RemainingPureBuiltins = Round(NPV(0.0625, values), 8) & \"|[\" & Partition(42, 0, 99, 5) & \"]|[\"\n\
               RemainingPureBuiltins = RemainingPureBuiltins & Partition(-2, 0, 99, 5) & \"]|[\" & Partition(101, 0, 99, 5) & \"]|\"\n\
               RemainingPureBuiltins = RemainingPureBuiltins & QBColor(1) & \"|\" & QBColor(7) & \"|\" & QBColor(15)\n\
             End Function\n",
            "RemainingPureBuiltins",
            vec![],
        )
        .unwrap();

        assert_eq!(
            value,
            Value::String(
                "19312.57020954|[ 40: 44]|[   : -1]|[100:   ]|8388608|12632256|16777215"
                    .to_string()
            )
        );
    }

    #[test]
    fn partition_rounds_arguments_and_propagates_null() {
        let value = run(
            "Public Function PartitionEdges() As String\n\
               PartitionEdges = \"[\" & Partition(2.5, 0, 99, 5) & \"]|[\" & Partition(3.5, 0, 99, 5) & \"]|\" & IsNull(Partition(1, Null, 9, 1))\n\
             End Function\n",
            "PartitionEdges",
            vec![],
        )
        .unwrap();

        assert_eq!(value, Value::String("[  0:  4]|[  0:  4]|True".to_string()));
    }

    #[test]
    fn npv_partition_and_qbcolor_report_invalid_inputs() {
        let value = run(
            "Public Function PureBuiltinErrors() As String\n\
               Dim values(1) As Double\n\
               Dim badFlows As Long\n\
               Dim badRange As Long\n\
               Dim badColor As Long\n\
               values(0) = 10\n\
               values(1) = 20\n\
               On Error Resume Next\n\
               badFlows = NPV(0.1, values)\n\
               badFlows = Err.Number\n\
               Err.Clear\n\
               badRange = Partition(5, 10, 5, 1)\n\
               badRange = Err.Number\n\
               Err.Clear\n\
               badColor = QBColor(16)\n\
               badColor = Err.Number\n\
               PureBuiltinErrors = badFlows & \"|\" & badRange & \"|\" & badColor\n\
             End Function\n",
            "PureBuiltinErrors",
            vec![],
        )
        .unwrap();

        assert_eq!(value, Value::String("5|5|5".to_string()));
    }

    #[test]
    fn creates_and_inspects_vba_error_values() {
        let value = run(
            "Public Function ErrorValues() As String\n\
               Dim failure As Variant\n\
               Dim implicitUse As Long\n\
               failure = CVErr(2001)\n\
               On Error Resume Next\n\
               implicitUse = failure & \"x\"\n\
               implicitUse = Err.Number\n\
               ErrorValues = IsError(failure) & \"|\" & IsError(2001) & \"|\" & VarType(failure) & \"|\" & TypeName(failure) & \"|\" & CInt(failure) & \"|\" & implicitUse\n\
             End Function\n",
            "ErrorValues",
            vec![],
        )
        .unwrap();

        assert_eq!(
            value,
            Value::String("True|False|10|Error|2001|13".to_string())
        );
    }

    #[test]
    fn returns_vba_error_descriptions() {
        let value = run(
            "Public Function ErrorDescriptions() As String\n\
               Dim ignored As Double\n\
               Dim latest As String\n\
               On Error Resume Next\n\
               ignored = 1 / 0\n\
               latest = Error()\n\
               ErrorDescriptions = Error(11) & \"|\" & Error(600) & \"|[\" & Error(0) & \"]|\" & latest\n\
             End Function\n",
            "ErrorDescriptions",
            vec![],
        )
        .unwrap();

        assert_eq!(
            value,
            Value::String(
                "Division by zero|Application-defined or object-defined error|[]|division by zero"
                    .to_string()
            )
        );
    }

    #[test]
    fn cverr_and_error_reject_numbers_outside_vba_range() {
        let value = run(
            "Public Function ErrorValueErrors() As String\n\
               Dim badValue As Long\n\
               Dim badMessage As Long\n\
               On Error Resume Next\n\
               badValue = CVErr(65536)\n\
               badValue = Err.Number\n\
               Err.Clear\n\
               badMessage = Error(-1)\n\
               badMessage = Err.Number\n\
               ErrorValueErrors = badValue & \"|\" & badMessage\n\
             End Function\n",
            "ErrorValueErrors",
            vec![],
        )
        .unwrap();

        assert_eq!(value, Value::String("5|5".to_string()));
    }

    #[test]
    fn executes_legacy_error_statement_and_err_default_property() {
        let value = run(
            "Public Function LegacyError() As String\n\
               On Error Resume Next\n\
               Error 11\n\
               LegacyError = Err & \"|\" & Err.Number & \"|\" & Error() & \"|\" & (Erl = Err.Erl)\n\
             End Function\n",
            "LegacyError",
            vec![],
        )
        .unwrap();

        assert_eq!(
            value,
            Value::String("11|11|Division by zero|True".to_string())
        );
    }

    #[test]
    fn on_error_statements_clear_the_err_object() {
        let value = run(
            "Public Function ErrorReset() As String\n\
               Dim ignored As Double\n\
               Dim before As Long\n\
               On Error Resume Next\n\
               ignored = 1 / 0\n\
               before = Err\n\
               On Error Resume Next\n\
               ErrorReset = before & \"|\" & Err & \"|[\" & Err.Description & \"]\"\n\
             End Function\n",
            "ErrorReset",
            vec![],
        )
        .unwrap();

        assert_eq!(value, Value::String("11|0|[]".to_string()));
    }

    #[test]
    fn err_raise_uses_the_standard_description_when_omitted() {
        let value = run(
            "Public Function DefaultRaiseDescription() As String\n\
               On Error Resume Next\n\
               Err.Raise 6\n\
               DefaultRaiseDescription = Err.Number & \"|\" & Err.Description\n\
             End Function\n",
            "DefaultRaiseDescription",
            vec![],
        )
        .unwrap();

        assert_eq!(value, Value::String("6|Overflow".to_string()));
    }

    #[test]
    fn doevents_returns_zero_for_the_browser_host() {
        let value = run(
            "Public Function PumpEvents() As String\n\
               PumpEvents = DoEvents & \"|\" & DoEvents()\n\
             End Function\n",
            "PumpEvents",
            vec![],
        )
        .unwrap();

        assert_eq!(value, Value::String("0|0".to_string()));
    }

    #[test]
    fn doevents_rejects_arguments() {
        let failure = run(
            "Public Function BadPump() As Long\n\
               BadPump = DoEvents(1)\n\
             End Function\n",
            "BadPump",
            vec![],
        )
        .unwrap_err();

        assert_eq!(failure.kind, RuntimeErrorKind::ArgumentCount);
    }

    #[test]
    fn executes_transcendental_math_and_rgb_functions() {
        let value = run(
            "Public Function MathAndColor() As String\n\
               MathAndColor = Round(4 * Atn(1), 6) & \"|\" & Sin(0) & \"|\" & Cos(0) & \"|\" & Tan(0) & \"|\" & Round(Log(Exp(2)), 6) & \"|\"\n\
               MathAndColor = MathAndColor & RGB(255, 128, 1) & \"|\" & RGB(300, 0, 0)\n\
             End Function\n",
            "MathAndColor",
            vec![],
        )
        .unwrap();

        assert_eq!(
            value,
            Value::String("3.141593|0|1|0|2|98559|255".to_string())
        );
    }

    #[test]
    fn executes_stateful_vba_rnd_and_randomize() {
        let value = run(
            "Public Function RandomSequence() As String\n\
               Dim first As Single\n\
               Dim second As Single\n\
               Dim repeated As Single\n\
               Dim seededA As Single\n\
               Dim seededB As Single\n\
               first = Rnd\n\
               second = Rnd()\n\
               repeated = Rnd(0)\n\
               seededA = Rnd(-1)\n\
               seededB = Rnd(-1)\n\
               Rnd -1\n\
               Randomize 42\n\
               Dim replayA As Single\n\
               replayA = Rnd\n\
               Rnd -1\n\
               Randomize 42\n\
               Dim replayB As Single\n\
               replayB = Rnd\n\
               RandomSequence = Round(first, 7) & \"|\" & Round(second, 7) & \"|\" & (second = repeated) & \"|\" & (seededA = seededB) & \"|\" & (replayA = replayB)\n\
             End Function\n",
            "RandomSequence",
            vec![],
        )
        .unwrap();

        assert_eq!(
            value,
            Value::String("0.7055475|0.533424|True|True|True".to_string())
        );
    }

    #[test]
    fn randomize_without_argument_uses_supplied_runtime_entropy() {
        let module = parse_module(
            "Public Function Sample() As Double\n\
               Randomize\n\
               Sample = Rnd\n\
             End Function\n",
        )
        .unwrap();
        let first = Runtime::new(&module)
            .with_random_seed(1)
            .call("Sample", vec![])
            .unwrap();
        let second = Runtime::new(&module)
            .with_random_seed(2)
            .call("Sample", vec![])
            .unwrap();

        assert_ne!(first, second);
    }

    #[test]
    fn random_functions_validate_arguments() {
        let value = run(
            "Public Function RandomErrors() As String\n\
               Dim rndCount As Long\n\
               Dim randomizeCount As Long\n\
               On Error Resume Next\n\
               rndCount = Rnd(1, 2)\n\
               rndCount = Err.Number\n\
               Err.Clear\n\
               Randomize 1, 2\n\
               randomizeCount = Err.Number\n\
               RandomErrors = rndCount & \"|\" & randomizeCount\n\
             End Function\n",
            "RandomErrors",
            vec![],
        )
        .unwrap();

        assert_eq!(value, Value::String("450|450".to_string()));
    }

    #[test]
    fn executes_current_date_time_and_timer_functions() {
        let module = parse_module(
            "Public Function ClockSnapshot() As String\n\
               ClockSnapshot = Format(Now, \"yyyy-mm-dd hh:nn:ss\") & \"|\" & Format(Date, \"yyyy-mm-dd\") & \"|\" & Format(Time, \"hh:nn:ss\") & \"|\" & Round(Timer, 3) & \"|\"\n\
               ClockSnapshot = ClockSnapshot & (Now = Now()) & \"|\" & (Date = Date()) & \"|\" & (Time = Time()) & \"|\" & (Timer = Timer())\n\
             End Function\n",
        )
        .unwrap();
        let clock =
            date_serial(2024, 2, 29).unwrap() + (16.0 * 3_600.0 + 35.0 * 60.0 + 17.25) / 86_400.0;
        let value = Runtime::new(&module)
            .with_current_time(clock)
            .call("ClockSnapshot", vec![])
            .unwrap();

        assert_eq!(
            value,
            Value::String(
                "2024-02-29 16:35:17|2024-02-29|16:35:17|59717.25|True|True|True|True".to_string()
            )
        );
    }

    #[test]
    fn current_time_functions_reject_arguments() {
        let value = run(
            "Public Function ClockErrors() As String\n\
               Dim nowError As Long\n\
               Dim timerError As Long\n\
               On Error Resume Next\n\
               nowError = Now(1)\n\
               nowError = Err.Number\n\
               Err.Clear\n\
               timerError = Timer(1)\n\
               timerError = Err.Number\n\
               ClockErrors = nowError & \"|\" & timerError\n\
             End Function\n",
            "ClockErrors",
            vec![],
        )
        .unwrap();

        assert_eq!(value, Value::String("450|450".to_string()));
    }

    #[test]
    fn filters_string_arrays_and_compares_strings_with_vba_modes() {
        let value = run(
            "Option Compare Text\n\
             Public Function FilterAndCompare() As String\n\
               Dim source As Variant\n\
               Dim included As Variant\n\
               Dim excluded As Variant\n\
               source = Array(\"Alpha\", \"beta\", \"ALPINE\", \"gamma\")\n\
               included = Filter(source, \"alp\", True, vbTextCompare)\n\
               excluded = Filter(source, \"alp\", False, vbTextCompare)\n\
               FilterAndCompare = Join(included, \"+\") & \"|\" & Join(excluded, \"+\") & \"|\"\n\
               FilterAndCompare = FilterAndCompare & StrComp(\"A\", \"a\") & \"|\" & StrComp(\"A\", \"a\", vbBinaryCompare) & \"|\" & IsNull(StrComp(Null, \"a\"))\n\
             End Function\n",
            "FilterAndCompare",
            vec![],
        )
        .unwrap();

        assert_eq!(
            value,
            Value::String("Alpha+ALPINE|beta+gamma|0|-1|True".to_string())
        );
    }

    #[test]
    fn math_filter_and_rgb_domain_errors_use_vba_numbers() {
        let value = run(
            "Public Function FunctionErrors() As String\n\
               Dim logarithm As Long\n\
               Dim exponential As Long\n\
               Dim color As Long\n\
               Dim filtering As Long\n\
               Dim source As Variant\n\
               source = Array(\"text\", 2)\n\
               On Error Resume Next\n\
               logarithm = Log(0)\n\
               logarithm = Err.Number\n\
               Err.Clear\n\
               exponential = Exp(1000)\n\
               exponential = Err.Number\n\
               Err.Clear\n\
               color = RGB(-1, 0, 0)\n\
               color = Err.Number\n\
               Err.Clear\n\
               filtering = UBound(Filter(source, \"x\"))\n\
               filtering = Err.Number\n\
               FunctionErrors = logarithm & \"|\" & exponential & \"|\" & color & \"|\" & filtering\n\
             End Function\n",
            "FunctionErrors",
            vec![],
        )
        .unwrap();

        assert_eq!(value, Value::String("5|6|5|13".to_string()));
    }

    #[test]
    fn executes_vba_date_serial_parsing_and_component_functions() {
        let value = run(
            "Public Function DateComponents() As String\n\
               Dim stamp As Date\n\
               stamp = DateSerial(2024, 2, 29) + TimeSerial(16, 35, 17)\n\
               DateComponents = DateSerial(1899, 12, 30) & \"|\" & DateSerial(2024, 2, 29) & \"|\"\n\
               DateComponents = DateComponents & Year(stamp) & \"-\" & Month(stamp) & \"-\" & Day(stamp) & \" \" & Hour(stamp) & \":\" & Minute(stamp) & \":\" & Second(stamp) & \"|\"\n\
               DateComponents = DateComponents & Year(#2/29/2024 4:35:17 PM#) & \"|\" & Hour(TimeValue(\"4:35:17 PM\")) & \"|\" & Day(DateValue(\"2024-02-29\")) & \"|\" & Day(DateValue(\"February 12, 1969\")) & \"|\" & IsDate(\"not a date\")\n\
             End Function\n",
            "DateComponents",
            vec![],
        )
        .unwrap();

        assert_eq!(
            value,
            Value::String("0|45351|2024-2-29 16:35:17|2024|16|29|12|False".to_string())
        );
    }

    #[test]
    fn executes_vba_dateadd_datediff_and_datepart_intervals() {
        let value = run(
            "Public Function DateIntervals() As String\n\
               Dim monthEnd As Date\n\
               monthEnd = DateAdd(\"m\", 1, DateSerial(2024, 1, 31))\n\
               DateIntervals = Year(monthEnd) & \"-\" & Month(monthEnd) & \"-\" & Day(monthEnd) & \"|\"\n\
               DateIntervals = DateIntervals & DateDiff(\"d\", DateSerial(2024, 1, 31), monthEnd) & \"|\" & DateDiff(\"yyyy\", DateSerial(2023, 12, 31), DateSerial(2024, 1, 1)) & \"|\"\n\
               DateIntervals = DateIntervals & DatePart(\"q\", monthEnd) & \"|\" & DatePart(\"y\", monthEnd) & \"|\" & DatePart(\"w\", DateSerial(2024, 1, 1), vbMonday) & \"|\" & Weekday(DateSerial(2024, 1, 1), vbMonday) & \"|\" & DatePart(\"ww\", DateSerial(2024, 1, 1), vbMonday, vbFirstFourDays) & \"|\"\n\
               DateIntervals = DateIntervals & Hour(TimeSerial(12 - 6, -15, 0)) & \":\" & Minute(TimeSerial(12 - 6, -15, 0))\n\
             End Function\n",
            "DateIntervals",
            vec![],
        )
        .unwrap();

        assert_eq!(
            value,
            Value::String("2024-2-29|29|1|1|60|1|1|1|5:45".to_string())
        );
    }

    #[test]
    fn executes_vba_named_and_custom_format_patterns() {
        let value = run(
            "Public Function FormatPatterns() As String\n\
               Dim stamp As Date\n\
               stamp = DateSerial(1993, 1, 27) + TimeSerial(17, 4, 23)\n\
               FormatPatterns = Format(stamp, \"yyyy-mm-dd hh:nn:ss\") & \"|\" & Format(stamp, \"dddd, mmm d yyyy\") & \"|\"\n\
               FormatPatterns = FormatPatterns & Format(TimeSerial(17, 4, 23), \"hh:mm:ss AM/PM\") & \"|\" & Format(TimeSerial(17, 4, 23), \"h:m:s\") & \"|\"\n\
               FormatPatterns = FormatPatterns & Format(5459.4, \"##,##0.00\") & \"|\" & Format(334.9, \"###0.00##\") & \"|\" & Format(5, \"0.00%\") & \"|\"\n\
               FormatPatterns = FormatPatterns & Format(1234.5, \"Standard\") & \"|\" & Format(\"HELLO\", \"<\") & \"|\" & Format(\"This is it\", \">\")\n\
             End Function\n",
            "FormatPatterns",
            vec![],
        )
        .unwrap();

        assert_eq!(
            value,
            Value::String(
                "1993-01-27 17:04:23|Wednesday, Jan 27 1993|05:04:23 PM|17:4:23|5,459.40|334.90|500.00%|1,234.50|hello|THIS IS IT"
                    .to_string()
            )
        );
    }

    #[test]
    fn executes_vba_format_helper_functions_and_constants() {
        let value = run(
            "Public Function FormatHelpers() As String\n\
               FormatHelpers = FormatNumber(1234.5, 2) & \"|\" & FormatPercent(0.125, 1) & \"|\"\n\
               FormatHelpers = FormatHelpers & FormatCurrency(-1234.5, 2, vbUseDefault, vbTrue, vbTrue) & \"|\"\n\
               FormatHelpers = FormatHelpers & FormatDateTime(DateSerial(1993, 1, 27), vbShortDate) & \"|\" & FormatDateTime(TimeSerial(17, 4, 23), vbLongTime)\n\
             End Function\n",
            "FormatHelpers",
            vec![],
        )
        .unwrap();

        assert_eq!(
            value,
            Value::String("1,234.50|12.5%|($1,234.50)|1/27/1993|5:04:23 PM".to_string())
        );
    }

    #[test]
    fn format_helpers_raise_vba_error_five_for_invalid_modes() {
        let value = run(
            "Public Function FormatErrors() As String\n\
               Dim dateMode As Long\n\
               Dim tristateMode As Long\n\
               On Error Resume Next\n\
               dateMode = FormatDateTime(DateSerial(2024, 1, 1), 5)\n\
               dateMode = Err.Number\n\
               Err.Clear\n\
               tristateMode = FormatNumber(1, 2, 1)\n\
               tristateMode = Err.Number\n\
               FormatErrors = dateMode & \"|\" & tristateMode\n\
             End Function\n",
            "FormatErrors",
            vec![],
        )
        .unwrap();

        assert_eq!(value, Value::String("5|5".to_string()));
    }

    #[test]
    fn date_builtins_report_invalid_dates_and_intervals_as_vba_errors() {
        let value = run(
            "Public Function DateErrors() As String\n\
               Dim badText As Long\n\
               Dim badInterval As Long\n\
               Dim badYear As Long\n\
               On Error Resume Next\n\
               badText = DateValue(\"2024-02-30\")\n\
               badText = Err.Number\n\
               Err.Clear\n\
               badInterval = DateAdd(\"bad\", 1, DateSerial(2024, 1, 1))\n\
               badInterval = Err.Number\n\
               Err.Clear\n\
               badYear = DateSerial(99, -32768, -32768)\n\
               badYear = Err.Number\n\
               DateErrors = badText & \"|\" & badInterval & \"|\" & badYear\n\
             End Function\n",
            "DateErrors",
            vec![],
        )
        .unwrap();

        assert_eq!(value, Value::String("13|5|5".to_string()));
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
