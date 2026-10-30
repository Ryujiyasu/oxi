// This Source Code Form is subject to the terms of the Mozilla Public
// License, v. 2.0. If a copy of the MPL was not distributed with this
// file, You can obtain one at https://mozilla.org/MPL/2.0/.

//! Procedure-level static checks, separate from parse recovery and run-time Err.
//!
//! In Excel's measured compile-on-demand mode a procedure is compiled when
//! called: an invalid uncalled procedure
//! does not prevent a valid entry point from running. Every branch of the called
//! procedure is checked, including ordinary `If False` branches. Conditional
//! compilation has already removed inactive source before this pass.
//!
//! Initial contracts cover CallByName's required object argument, required
//! parameters, named arguments, and the declarations that determine their types.

use std::collections::{BTreeMap, BTreeSet};

use crate::ast::*;
use crate::references::{CompileReferenceLibrary, ReferenceSignature};
use crate::runtime::{RuntimeErrorKind, Value};
use crate::{ErrorLanguage, Span};

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum CompileErrorKind {
    SyntaxError,
    ExpectedExpression,
    ExpectedIdentifier,
    EndIfWithoutBlock,
    TypeMismatch,
    ArgumentNotOptional,
    NamedArgumentInParamArray,
    VariableNotDefined,
    AmbiguousName,
    MemberNotFound,
    ProcedureNotDefined,
    VariableRequired,
    InvalidParamArrayUse,
    InvalidPublicClassMember,
    ArrayOrUserDefinedTypeExpected,
    CannotAssignArray,
    InvalidObjectUse,
    InvalidConstantType,
    InvalidPropertyUse,
    ObjectVariableNotSet,
    ConstantExpressionRequired,
    ValueOutOfRange,
    InvalidFixedStringLength,
    Overflow,
    DivisionByZero,
    InvalidNull,
    UserDefinedTypeNotDefined,
    CircularModuleReference,
    WrongNumberOfArguments,
    NamedArgumentNotFound,
    NamedArgumentAlreadySpecified,
    InheritedMemberIdentifier,
    InvalidNewUse,
    ReadOnlyProperty,
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct CompileError {
    pub kind: CompileErrorKind,
    pub span: Span,
    pub procedure: String,
    pub symbol: Option<String>,
}

/// A diagnostic from explicit project compilation, retaining its component.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct ProjectCompileError {
    pub module: String,
    pub error: CompileError,
}

impl std::fmt::Display for ProjectCompileError {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        write!(f, "{}: {}", self.module, self.error)
    }
}
impl std::error::Error for ProjectCompileError {}

#[derive(Debug, Clone, Default)]
pub struct CompileReferences {
    pub types: Vec<(String, bool)>,
    /// Constants have a value; host globals and functions have no constant value.
    pub symbols: Vec<(String, Option<i64>)>,
    pub namespaces: Vec<String>,
    /// Current source project, for fully qualified class names.
    pub project_name: Option<String>,
    /// Host-owned document code names and their native base types. Source
    /// attributes alone do not distinguish a document from an ordinary class.
    pub document_modules: Vec<(String, String)>,
    pub current_year: Option<i64>,
    pub library: std::sync::Arc<CompileReferenceLibrary>,
}

#[derive(Clone, Copy)]
struct ReferenceContext<'a> {
    strict: bool,
    project_name: &'a str,
    types: &'a BTreeMap<String, bool>,
    symbols: &'a BTreeMap<String, Option<i64>>,
    namespaces: &'a BTreeSet<String>,
    documents: &'a BTreeMap<String, String>,
    current_year: i64,
    library: &'a CompileReferenceLibrary,
}

/// Compile all declarations and procedures without initializing project storage.
/// Host reference types can be supplied separately; `true` marks a numeric type.
pub fn compile_project_with_references(
    source: &str,
    references: &CompileReferences,
) -> Result<(), ProjectCompileError> {
    let components = crate::parser::project_components(source);
    for (_, name, text) in &components {
        if let Err(error) = crate::parse_module(text) {
            let line = match error {
                crate::lexer::LexError::UnterminatedString { line }
                | crate::lexer::LexError::UnterminatedDate { line }
                | crate::lexer::LexError::UnterminatedBracketExpr { line }
                | crate::lexer::LexError::UnexpectedChar { line, .. } => line,
            };
            return Err(ProjectCompileError {
                module: name.clone(),
                error: CompileError {
                    kind: CompileErrorKind::SyntaxError,
                    span: Span {
                        line,
                        start: 0,
                        end: 0,
                    },
                    procedure: String::new(),
                    symbol: None,
                },
            });
        }
    }
    let (globals, class_modules) =
        crate::parse_project(source).expect("components were lexed above");
    let classes: BTreeMap<_, _> = class_modules
        .iter()
        .map(|(name, module)| (key(name), module))
        .collect();
    let types: BTreeMap<_, _> = references
        .types
        .iter()
        .map(|(name, numeric)| (key(name), *numeric))
        .collect();
    let symbols: BTreeMap<_, _> = references
        .symbols
        .iter()
        .map(|(name, value)| (key(name), *value))
        .collect();
    let namespaces: BTreeSet<_> = references.namespaces.iter().map(|name| key(name)).collect();
    let documents: BTreeMap<_, _> = references
        .document_modules
        .iter()
        .map(|(name, base)| {
            (
                key(name),
                references
                    .library
                    .resolve_type(base)
                    .map(|(owner, _)| owner)
                    .unwrap_or_else(|| key(base)),
            )
        })
        .collect();
    let context = ReferenceContext {
        strict: true,
        project_name: references.project_name.as_deref().filter(|name| !name.is_empty()).unwrap_or("VBAProject"),
        types: &types,
        symbols: &symbols,
        namespaces: &namespaces,
        documents: &documents,
        current_year: references.current_year.unwrap_or(2000),
        library: &references.library,
    };
    let mut ordered = Vec::new();
    for (class, name, _) in &components {
        let scope = if *class {
            classes.get(&key(name)).copied()
        } else if globals.standard_modules.is_empty() {
            Some(&globals)
        } else {
            globals
                .standard_modules
                .iter()
                .find(|(found, _)| found.eq_ignore_ascii_case(name))
                .map(|(_, scope)| scope)
        };
        if let Some(scope) = scope {
            ordered.push((name.clone(), scope));
        }
    }
    for (name, scope) in &ordered {
        let mut checker =
            Checker::for_context_with_references(scope, &globals, &classes, None, Some(context));
        checker
            .module_declarations()
            .map_err(|error| ProjectCompileError {
                module: name.clone(),
                error,
            })?;
    }
    for (name, scope) in &ordered {
        for item in &scope.items {
            if let ModuleItem::Procedure(procedure) = item {
                check_with_environment(scope, &globals, &classes, procedure, Some(context), context.project_name)
                    .map_err(|error| ProjectCompileError {
                        module: name.clone(),
                        error,
                    })?;
            }
        }
    }
    Ok(())
}

pub fn compile_project_with_types(
    source: &str,
    reference_types: &[(String, bool)],
) -> Result<(), ProjectCompileError> {
    compile_project_with_references(
        source,
        &CompileReferences {
            types: reference_types.to_vec(),
            ..CompileReferences::default()
        },
    )
}

/// Compile a project against the standard VBA types, with no additional host references.
pub fn compile_project(source: &str) -> Result<(), ProjectCompileError> {
    compile_project_with_types(source, &[])
}

impl CompileError {
    pub fn description(&self, language: ErrorLanguage) -> &'static str {
        match (self.kind, language) {
            (CompileErrorKind::InvalidNewUse,ErrorLanguage::English)=>"Invalid use of New keyword",
            (CompileErrorKind::InvalidNewUse,ErrorLanguage::Japanese)=>"New キーワードの使用法が不正です。",
            (CompileErrorKind::ReadOnlyProperty,ErrorLanguage::English)=>"Can't assign to read-only property",
            (CompileErrorKind::ReadOnlyProperty,ErrorLanguage::Japanese)=>"値の取得のみ可能なプロパティに値を設定することはできません。",
            (CompileErrorKind::InheritedMemberIdentifier, ErrorLanguage::English) => "Member identifier already exists in object module from which this object module derives",
            (CompileErrorKind::InheritedMemberIdentifier, ErrorLanguage::Japanese) => "このメンバー識別子は既にオブジェクト モジュールの中に存在しています。",
            (CompileErrorKind::NamedArgumentAlreadySpecified, ErrorLanguage::English) => "Named argument already specified",
            (CompileErrorKind::NamedArgumentAlreadySpecified, ErrorLanguage::Japanese) => "その名前付き引数は、既に指定されています。",
            (CompileErrorKind::WrongNumberOfArguments, ErrorLanguage::English) => "Wrong number of arguments or invalid property assignment",
            (CompileErrorKind::WrongNumberOfArguments, ErrorLanguage::Japanese) => "引数の数が一致していません。または不正なプロパティを指定しています。",
            (CompileErrorKind::NamedArgumentNotFound, ErrorLanguage::English) => "Named argument not found",
            (CompileErrorKind::NamedArgumentNotFound, ErrorLanguage::Japanese) => "名前付き引数が見つかりません。",
            (CompileErrorKind::UserDefinedTypeNotDefined, ErrorLanguage::English) => "User-defined type not defined",
            (CompileErrorKind::UserDefinedTypeNotDefined, ErrorLanguage::Japanese) => "ユーザ定義型は定義されていません。",
            (CompileErrorKind::CircularModuleReference, ErrorLanguage::English) => "Circular dependencies between modules",
            (CompileErrorKind::CircularModuleReference, ErrorLanguage::Japanese) => "モジュール間で循環参照が定義されています。",
            (CompileErrorKind::InvalidNull, ErrorLanguage::English) => "Invalid use of Null",
            (CompileErrorKind::InvalidNull, ErrorLanguage::Japanese) => "Null の使い方が不正です。",
            (CompileErrorKind::DivisionByZero, ErrorLanguage::English) => "Division by zero",
            (CompileErrorKind::DivisionByZero, ErrorLanguage::Japanese) => "0 で除算しました。",
            (CompileErrorKind::ConstantExpressionRequired, ErrorLanguage::English) => "Constant expression required",
            (CompileErrorKind::ConstantExpressionRequired, ErrorLanguage::Japanese) => "定数式が必要です。",
            (CompileErrorKind::ValueOutOfRange, ErrorLanguage::English) => "Value out of range",
            (CompileErrorKind::ValueOutOfRange, ErrorLanguage::Japanese) => "値が有効範囲にありません。",
            (CompileErrorKind::InvalidFixedStringLength, ErrorLanguage::English) => "Invalid length for fixed-length string",
            (CompileErrorKind::InvalidFixedStringLength, ErrorLanguage::Japanese) => "固定長文字列の長さが不正です。",
            (CompileErrorKind::Overflow, ErrorLanguage::English) => "Overflow",
            (CompileErrorKind::Overflow, ErrorLanguage::Japanese) => "オーバーフローしました。",
            (CompileErrorKind::InvalidConstantType,ErrorLanguage::English)=>"Invalid constant type",
            (CompileErrorKind::InvalidConstantType,ErrorLanguage::Japanese)=>"定数のデータ型が不正です。",
            (CompileErrorKind::InvalidPropertyUse,ErrorLanguage::English)=>"Invalid use of property",
            (CompileErrorKind::InvalidPropertyUse,ErrorLanguage::Japanese)=>"プロパティの使い方が不正です。",
            (CompileErrorKind::ObjectVariableNotSet,ErrorLanguage::English)=>"Object variable or With block variable not set",
            (CompileErrorKind::ObjectVariableNotSet,ErrorLanguage::Japanese)=>"オブジェクト変数または With ブロック変数が設定されていません。",
            (CompileErrorKind::InvalidObjectUse,ErrorLanguage::English)=>"Invalid use of object",
            (CompileErrorKind::InvalidObjectUse,ErrorLanguage::Japanese)=>"オブジェクトの使い方が不正です。",
            (CompileErrorKind::CannotAssignArray, ErrorLanguage::English) => "Can't assign to an array",
            (CompileErrorKind::CannotAssignArray, ErrorLanguage::Japanese) => "配列には割り当てられません。",
            (CompileErrorKind::ArrayOrUserDefinedTypeExpected,ErrorLanguage::English)=>"Type mismatch: array or user-defined type expected",
            (CompileErrorKind::ArrayOrUserDefinedTypeExpected,ErrorLanguage::Japanese)=>"型が一致しません: 配列またはユーザー定義型を指定してください。",
            (CompileErrorKind::ExpectedExpression, ErrorLanguage::English) => "Expected: expression",
            (CompileErrorKind::ExpectedIdentifier, ErrorLanguage::English) => "Expected: identifier",
            (CompileErrorKind::ExpectedIdentifier, ErrorLanguage::Japanese) => "修正候補: 識別子",
            (CompileErrorKind::EndIfWithoutBlock, ErrorLanguage::English) => "End If without block If",
            (CompileErrorKind::EndIfWithoutBlock, ErrorLanguage::Japanese) => "End If に対応する If ブロックがありません。",
            (CompileErrorKind::ExpectedExpression, ErrorLanguage::Japanese) => "修正候補: 式",
            (CompileErrorKind::SyntaxError, ErrorLanguage::English) => "Syntax error",
            (CompileErrorKind::SyntaxError, ErrorLanguage::Japanese) => "構文エラー",
            (CompileErrorKind::TypeMismatch, ErrorLanguage::English) => "Type mismatch",
            (CompileErrorKind::TypeMismatch, ErrorLanguage::Japanese) => "型が一致しません。",
            (CompileErrorKind::ArgumentNotOptional, ErrorLanguage::English) => {
                "Argument not optional"
            }
            (CompileErrorKind::ArgumentNotOptional, ErrorLanguage::Japanese) => {
                "引数は省略できません。"
            }
            (CompileErrorKind::NamedArgumentInParamArray, ErrorLanguage::English) => {
                "Named argument not a ParamArray element"
            }
            (CompileErrorKind::NamedArgumentInParamArray, ErrorLanguage::Japanese) => {
                "ParamArray 内の引数は、名前付き引数ではありません。"
            }
            (CompileErrorKind::VariableNotDefined, ErrorLanguage::English) => {
                "Variable not defined"
            }
            (CompileErrorKind::VariableNotDefined, ErrorLanguage::Japanese) => {
                "変数が定義されていません。"
            }
            (CompileErrorKind::AmbiguousName, ErrorLanguage::English) => "Ambiguous name detected",
            (CompileErrorKind::AmbiguousName, ErrorLanguage::Japanese) => {
                "名前が適切ではありません"
            }
            (CompileErrorKind::MemberNotFound, ErrorLanguage::English) => {
                "Method or data member not found"
            }
            (CompileErrorKind::MemberNotFound, ErrorLanguage::Japanese) => {
                "メソッドまたはデータ メンバーが見つかりません。"
            }
            (CompileErrorKind::ProcedureNotDefined, ErrorLanguage::English) => {
                "Sub or Function not defined"
            }
            (CompileErrorKind::ProcedureNotDefined, ErrorLanguage::Japanese) => {
                "Sub または Function が定義されていません。"
            }
            (CompileErrorKind::VariableRequired, ErrorLanguage::English) => "Variable required",
            (CompileErrorKind::VariableRequired, ErrorLanguage::Japanese) => "変数が必要です。",
            (CompileErrorKind::InvalidParamArrayUse, ErrorLanguage::English) => {
                "Invalid use of ParamArray"
            }
            (CompileErrorKind::InvalidParamArrayUse, ErrorLanguage::Japanese) => {
                "ParamArray の使い方が適切ではありません。"
            }
            (CompileErrorKind::InvalidPublicClassMember, ErrorLanguage::English) => {
                "Constants, fixed-length strings, arrays, user-defined types and Declare statements not allowed as Public members of object modules"
            }
            (CompileErrorKind::InvalidPublicClassMember, ErrorLanguage::Japanese) => {
                "定数、固定長文字列、配列、ユーザー定義型および Declare ステートメントは、オブジェクト モジュールのパブリック メンバーとしては使用できません。"
            }
        }
    }

    pub fn message(&self, language: ErrorLanguage) -> String {
        match &self.symbol {
            Some(symbol) if self.kind == CompileErrorKind::AmbiguousName => {
                format!("{}: {symbol}", self.description(language))
            }
            _ => self.description(language).to_string(),
        }
    }
}

impl std::fmt::Display for CompileError {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        write!(
            f,
            "line {}: Compile error: {}",
            self.span.line,
            self.message(ErrorLanguage::English)
        )
    }
}

impl std::error::Error for CompileError {}

/// Checks a procedure using declarations in its module. This is a semantic pass:
/// parsing and migration analysis can still inspect source that fails this check.
pub fn check_procedure(module: &Module, procedure: &Procedure) -> Result<(), CompileError> {
    check_with_globals(module, module, &BTreeMap::new(), procedure, "VBAProject")
}

pub(crate) fn check_with_globals(
    module: &Module,
    globals: &Module,
    classes: &BTreeMap<String, &Module>,
    procedure: &Procedure,
    project_name: &str,
) -> Result<(), CompileError> {
    check_with_environment(module, globals, classes, procedure, None, project_name)
}

/// Compile on demand with host document surfaces, retaining the ordinary Run
/// mode's existing treatment of unrelated declarations and call signatures.
pub(crate) fn check_with_document_globals(
    module: &Module,
    globals: &Module,
    classes: &BTreeMap<String, &Module>,
    procedure: &Procedure,
    references: &CompileReferences,
    current_year: i64,
    project_name: &str,
) -> Result<(), CompileError> {
    let documents = references
        .document_modules
        .iter()
        .map(|(name, base)| {
            (
                key(name),
                references
                    .library
                    .resolve_type(base)
                    .map(|(name, _)| name)
                    .unwrap_or_else(|| key(base)),
            )
        })
        .collect();
    check_with_environment(
        module,
        globals,
        classes,
        procedure,
        Some(ReferenceContext {
            strict: false,
            project_name,
            types: &BTreeMap::new(),
            symbols: &BTreeMap::new(),
            namespaces: &BTreeSet::new(),
            documents: &documents,
            current_year: references.current_year.unwrap_or(current_year),
            library: &references.library,
        }),
        project_name,
    )
}

fn check_with_environment<'a>(
    module: &'a Module,
    globals: &'a Module,
    classes: &'a BTreeMap<String, &'a Module>,
    procedure: &'a Procedure,
    references: Option<ReferenceContext<'a>>,
    project_name: &'a str,
) -> Result<(), CompileError> {
    let mut checker =
        Checker::for_context_with_references(module, globals, classes, Some(procedure), references);
    checker.project_name = project_name;
    if checker.known_types.is_some() && intrinsic_operator(&key(&procedure.name)) {
        return Err(checker.failure(CompileErrorKind::SyntaxError, procedure.span));
    }
    if !classes.values().any(|class| std::ptr::eq(*class, module)) {
        if let Some(found) = module.items.iter().find_map(|item| match item {
            ModuleItem::Procedure(found)
                if matches!(
                    key(&found.name).as_str(),
                    "clng" | "cint" | "cstr" | "cbool" | "len" | "abs"
                ) =>
            {
                Some(found)
            }
            _ => None,
        }) {
            return Err(checker.failure(CompileErrorKind::SyntaxError, found.span));
        }
    }
    if let Some(span) = procedure.syntax_errors.iter().find(|span| {
        references.is_none_or(|references| !references.strict)
            || checker.logical_line(**span) == checker.logical_line(procedure.span)
    }) {
        return Err(checker.failure(CompileErrorKind::SyntaxError, *span));
    }
    if let Some((name, _)) = classes
        .iter()
        .find(|(_, class)| std::ptr::eq(**class, module))
    {
        checker.check_class(name, procedure.span)?;
    }
    for parameter in &procedure.params {
        if checker.reserved_name(&parameter.name) {
            return Err(checker.failure(CompileErrorKind::SyntaxError, procedure.span));
        }
        checker.check_class(&parameter.type_name.name, procedure.span)?;
        checker.validate_type(&parameter.type_name, procedure.span)?;
    }
    if let Some(kind) = &procedure.return_type {
        checker.check_class(&kind.name, procedure.span)?;
        checker.validate_type(kind, procedure.span)?;
    }
    checker.body(&procedure.body)
}

fn reserved_declaration_name(name: &str) -> bool {
    matches!(
        key(name).as_str(),
        "boolean"
            | "byte"
            | "integer"
            | "long"
            | "longlong"
            | "longptr"
            | "single"
            | "double"
            | "currency"
            | "date"
            | "string"
            | "variant"
            | "decimal"
            | "me"
            | "true"
            | "nothing"
            | "if"
            | "new"
            | "as"
            | "type"
            | "next"
            | "debug"
            | "clng"
            | "len"
            | "abs"
    )
}

fn intrinsic_operator(name: &str) -> bool {
    matches!(
        name,
        "abs"
            | "len"
            | "lenb"
            | "int"
            | "fix"
            | "sgn"
            | "cbool"
            | "cbyte"
            | "ccur"
            | "cdate"
            | "cdbl"
            | "cdec"
            | "cint"
            | "clng"
            | "clnglng"
            | "clngptr"
            | "csng"
            | "cstr"
            | "cvar"
    )
}

pub(crate) fn procedure_header_valid(procedure: &Procedure) -> bool {
    !procedure
        .params
        .iter()
        .any(|parameter| reserved_declaration_name(&parameter.name))
        && !procedure
            .syntax_errors
            .iter()
            .any(|span| span.start == procedure.span.start)
}

#[derive(Debug, Clone, PartialEq, Eq)]
enum Type {
    Scalar,
    Text,
    Variant,
    Object,
    Class(String),
    Reference(String, bool),
    Array(Box<Type>),
    Record(String),
    Unknown,
}

#[derive(Clone)]
struct ConstantBinding<'a> {
    expression: Expr,
    declared: TypeName,
    owner: Option<&'a Module>,
    /// Enum expressions admit intrinsic operators; an omitted member wraps a Long.
    enumeration: Option<bool>,
}

#[derive(Clone, Copy, PartialEq, Eq)]
enum ConstantContext {
    Initializer,
    Bound,
    Enumeration,
}

struct Checker<'a> {
    project_name: &'a str,
    module: &'a Module,
    globals: &'a Module,
    classes: &'a BTreeMap<String, &'a Module>,
    ambiguous: BTreeSet<String>,
    hidden_variables: BTreeSet<String>,
    hidden_functions: BTreeSet<String>,
    own_symbols: BTreeSet<String>,
    procedure: Option<&'a Procedure>,
    known_types: Option<&'a BTreeMap<String, bool>>,
    known_symbols: Option<&'a BTreeMap<String, Option<i64>>>,
    known_namespaces: Option<&'a BTreeSet<String>>,
    reference_library: Option<&'a CompileReferenceLibrary>,
    document_modules: Option<&'a BTreeMap<String, String>>,
    compilation_year: Option<i64>,
    variables: BTreeMap<String, Type>,
    fixed_arrays: BTreeSet<String>,
    functions: BTreeMap<String, Type>,
    constants: BTreeSet<String>,
    constant_expressions: BTreeMap<String, ConstantBinding<'a>>,
    records: BTreeMap<String, &'a TypeDef>,
    enums: BTreeSet<String>,
    later_locals: BTreeSet<String>,
    with_types: Vec<Type>,
    explicit: bool,
    syntax_lines: BTreeMap<usize, Vec<(&'a Expr, bool)>>,
    checked_syntax_lines: std::cell::RefCell<BTreeSet<usize>>,
    line_syntax_errors: BTreeMap<usize, Span>,
}

impl<'a> Checker<'a> {
    fn reserved_name(&self, name: &str) -> bool {
        reserved_declaration_name(name)
            || (self.known_types.is_some() && intrinsic_operator(&key(name)))
    }
    fn for_context(
        module: &'a Module,
        globals: &'a Module,
        classes: &'a BTreeMap<String, &'a Module>,
        procedure: Option<&'a Procedure>,
    ) -> Self {
        Self::for_context_with_references(module, globals, classes, procedure, None)
    }

    fn for_context_with_references(
        module: &'a Module,
        globals: &'a Module,
        classes: &'a BTreeMap<String, &'a Module>,
        procedure: Option<&'a Procedure>,
        references: Option<ReferenceContext<'a>>,
    ) -> Self {
        let mut checker = Self {
            project_name: references.map_or("VBAProject", |references| references.project_name),
            module,
            globals,
            classes,
            ambiguous: BTreeSet::new(),
            hidden_variables: BTreeSet::new(),
            hidden_functions: BTreeSet::new(),
            own_symbols: BTreeSet::new(),
            procedure,
            known_types: references.filter(|r| r.strict).map(|r| r.types),
            known_symbols: references.filter(|r| r.strict).map(|r| r.symbols),
            known_namespaces: references.filter(|r| r.strict).map(|r| r.namespaces),
            reference_library: references.map(|r| r.library),
            document_modules: references.map(|r| r.documents),
            compilation_year: references.map(|r| r.current_year),
            variables: BTreeMap::new(),
            fixed_arrays: BTreeSet::new(),
            functions: BTreeMap::new(),
            constants: BTreeSet::new(),
            constant_expressions: BTreeMap::new(),
            records: BTreeMap::new(),
            enums: BTreeSet::new(),
            later_locals: BTreeSet::new(),
            with_types: Vec::new(),
            syntax_lines: BTreeMap::new(),
            checked_syntax_lines: std::cell::RefCell::new(BTreeSet::new()),
            line_syntax_errors: BTreeMap::new(),
            explicit: module
                .items
                .iter()
                .any(|item| matches!(item, ModuleItem::Option(ModuleOption::Explicit, _))),
        };
        let mut public_names = BTreeSet::new();
        for item in &module.items {
            match item {
                ModuleItem::Variables(d) => checker
                    .own_symbols
                    .extend(d.items.iter().map(|v| key(&v.name))),
                ModuleItem::Procedure(p) => {
                    checker.own_symbols.insert(key(&p.name));
                }
                _ => {}
            }
        }
        for (_, scope) in &globals.standard_modules {
            let own = std::ptr::eq(scope, module);
            let mut scope_names = BTreeSet::new();
            for item in &scope.items {
                match item {
                    ModuleItem::Variables(d) => {
                        for v in &d.items {
                            let name = key(&v.name);
                            if own {
                                checker.own_symbols.insert(name);
                            } else if matches!(
                                d.visibility,
                                Visibility::Public | Visibility::Global
                            ) {
                                scope_names.insert(name);
                            } else {
                                checker.hidden_variables.insert(name);
                            }
                        }
                    }
                    ModuleItem::Procedure(p) => {
                        let name = key(&p.name);
                        if own {
                            checker.own_symbols.insert(name);
                        } else if p.visibility != Visibility::Private {
                            scope_names.insert(name);
                        } else {
                            checker.hidden_functions.insert(name);
                        }
                    }
                    _ => {}
                }
            }
            for name in scope_names {
                if !public_names.insert(name.clone()) {
                    checker.ambiguous.insert(name);
                }
            }
        }
        for source in [globals, module] {
            for item in &source.items {
                match item {
                    ModuleItem::Type(record) => {
                        checker.records.insert(key(&record.name), record);
                    }
                    ModuleItem::Enum(enumeration) => {
                        checker.enums.insert(key(&enumeration.name));
                    }
                    _ => {}
                }
            }
        }
        for source in [globals, module] {
            for item in &source.items {
                match item {
                    ModuleItem::Variables(declaration)
                        if std::ptr::eq(source, module)
                            || matches!(
                                declaration.visibility,
                                Visibility::Public | Visibility::Global
                            ) =>
                    {
                        let owner = checker.definition_scope(source, declaration.span);
                        checker.declaration_in(declaration, Some(owner))
                    }
                    ModuleItem::Enum(enumeration) => {
                        let owner = checker.definition_scope(source, enumeration.span);
                        let mut previous = Expr::Literal(Literal::Number(-1.0), enumeration.span);
                        for (name, _) in &enumeration.members {
                            checker.variables.insert(key(name), Type::Scalar);
                            checker.constants.insert(key(name));
                        }
                        for (name, value) in &enumeration.members {
                            let implicit = value.is_none();
                            let value = value.clone().unwrap_or_else(|| previous.clone());
                            let declared = TypeName {
                                name: "Long".into(),
                                ..TypeName::implicit()
                            };
                            let qualified = format!("{}.{}", enumeration.name, name);
                            checker.constant_expressions.insert(
                                key(name),
                                ConstantBinding {
                                    expression: value.clone(),
                                    declared: declared.clone(),
                                    owner: Some(owner),
                                    enumeration: Some(implicit),
                                },
                            );
                            checker.constant_expressions.insert(
                                key(&qualified),
                                ConstantBinding {
                                    expression: value,
                                    declared,
                                    owner: Some(owner),
                                    enumeration: Some(implicit),
                                },
                            );
                            previous = Expr::Member {
                                object: Box::new(Expr::Ident(
                                    enumeration.name.clone(),
                                    enumeration.span,
                                )),
                                name: name.clone(),
                                suffix: None,
                                span: enumeration.span,
                            };
                        }
                    }
                    ModuleItem::Procedure(found) => {
                        if std::ptr::eq(source, module)
                            || !matches!(found.visibility, Visibility::Private)
                        {
                            if !procedure_header_valid(found) {
                                checker.own_symbols.remove(&key(&found.name));
                                checker.hidden_functions.insert(key(&found.name));
                                continue;
                            }
                            let result = found
                                .return_type
                                .as_ref()
                                .map(|name| checker.declared(name))
                                .unwrap_or(Type::Variant);
                            checker.functions.insert(key(&found.name), result);
                        }
                    }
                    ModuleItem::ExternalProc(found) => {
                        let result = found
                            .return_type
                            .as_ref()
                            .map(|name| checker.declared(name))
                            .unwrap_or(Type::Variant);
                        checker.functions.insert(key(&found.name), result);
                    }
                    _ => {}
                }
            }
        }
        if let Some(procedure) =
            procedure.filter(|p| matches!(p.kind, ProcKind::Function | ProcKind::PropertyGet))
        {
            let result = procedure
                .return_type
                .as_ref()
                .map(|name| checker.declared(name))
                .unwrap_or(Type::Variant);
            checker.variables.insert(key(&procedure.name), result);
        }
        if let Some((name, _)) = classes
            .iter()
            .find(|(_, class)| std::ptr::eq(**class, module))
        {
            checker
                .variables
                .insert("me".into(), Type::Class(name.clone()));
        }
        for param in procedure.into_iter().flat_map(|p| &p.params) {
            checker.constant_expressions.remove(&key(&param.name));
            checker.own_symbols.insert(key(&param.name));
            let mut kind = checker.declared(&param.type_name);
            if param.is_array || param.mode == ParamMode::ParamArray {
                kind = Type::Array(Box::new(kind));
            }
            checker.fixed_arrays.remove(&key(&param.name));
            checker.variables.insert(key(&param.name), kind);
        }
        if let Some(procedure) = procedure {
            collect_locals(&procedure.body, &mut checker.later_locals);
            if references.is_some_and(|r| r.strict) {
                for span in &procedure.syntax_errors {
                    checker
                        .line_syntax_errors
                        .entry(checker.logical_line(*span))
                        .or_insert(*span);
                }
                let mut roots = Vec::new();
                collect_syntax_roots(&procedure.body, &mut roots);
                for root in roots {
                    checker
                        .syntax_lines
                        .entry(checker.logical_line(root.0.span()))
                        .or_default()
                        .push(root);
                }
            }
        }
        if let Some(references) = references {
            checker.install_references(references);
        }
        checker
    }

    fn install_references(&mut self, references: ReferenceContext<'a>) {
        self.project_name = references.project_name;
        self.known_types = references.strict.then_some(references.types);
        self.known_symbols = references.strict.then_some(references.symbols);
        self.known_namespaces = references.strict.then_some(references.namespaces);
        self.reference_library = Some(references.library);
        self.document_modules = Some(references.documents);
        self.compilation_year = Some(references.current_year);
        for (name, value) in references.symbols {
            if self.variables.contains_key(name)
                || self.functions.contains_key(name)
                || self.later_locals.contains(name)
            {
                continue;
            }
            if let Some(value) = value {
                self.variables.insert(name.clone(), Type::Scalar);
                self.constants.insert(name.clone());
                self.constant_expressions.insert(
                    name.clone(),
                    ConstantBinding {
                        expression: Expr::Literal(
                            Literal::Number(*value as f64),
                            Span {
                                line: 1,
                                start: 0,
                                end: 0,
                            },
                        ),
                        declared: TypeName {
                            name: "Long".into(),
                            ..TypeName::implicit()
                        },
                        owner: None,
                        enumeration: None,
                    },
                );
            }
        }
    }

    fn source_class_key(&self, name: &str) -> Option<String> {
        let lowered = key(name);
        let name = match lowered.split_once('.') {
            Some((project, class)) if project.eq_ignore_ascii_case(self.project_name) => class,
            _ => lowered.as_str(),
        };
        self.classes.contains_key(name).then(|| name.to_string())
    }

    fn source_type_key(&self, name: &str) -> String {
        self.source_class_key(name).unwrap_or_else(|| key(name))
    }

    fn validate_type(&self, kind: &TypeName, span: Span) -> Result<(), CompileError> {
        let Some(known) = self.known_types else {
            return Ok(());
        };
        let name = self.source_type_key(&kind.name);
        if name == "decimal" {
            return Err(self.failure(CompileErrorKind::SyntaxError, span));
        }
        if self.records.contains_key(&name)
            || self.enums.contains(&name)
            || self.classes.contains_key(&name)
            || self
                .document_modules
                .is_some_and(|documents| documents.contains_key(&name))
            || known.contains_key(&name)
            || matches!(
                name.as_str(),
                "byte"
                    | "integer"
                    | "long"
                    | "longlong"
                    | "longptr"
                    | "single"
                    | "double"
                    | "currency"
                    | "date"
                    | "boolean"
                    | "string"
                    | "variant"
                    | "object"
                    | "collection"
                    | "vba.collection"
                    | "any"
            )
        {
            Ok(())
        } else {
            Err(self.failure(CompileErrorKind::UserDefinedTypeNotDefined, span))
        }
    }

    fn module_declarations(&mut self) -> Result<(), CompileError> {
        if let Some(span) = self.module.syntax_errors.first() {
            return Err(self.failure(CompileErrorKind::SyntaxError, *span));
        }
        self.declaration_syntax()?;
        if let Some((name, _)) = self
            .classes
            .iter()
            .find(|(_, module)| std::ptr::eq(**module, self.module))
        {
            let span = self
                .module
                .items
                .iter()
                .find_map(|item| match item {
                    ModuleItem::Variables(decl) => Some(decl.span),
                    _ => None,
                })
                .unwrap_or(Span {
                    line: 1,
                    start: 0,
                    end: 0,
                });
            self.check_class(name, span)?;
        }
        for item in &self.module.items {
            match item {
                ModuleItem::Variables(declaration) => {
                    self.statement(&Statement::Dim(declaration.clone()))?
                }
                ModuleItem::Enum(enumeration) => {
                    let mut previous = -1_i64;
                    for (_, expression) in &enumeration.members {
                        previous = match expression {
                            Some(expression) => {
                                self.expression(expression)?;
                                let value = self
                                    .constant_value_in(
                                        expression,
                                        &mut BTreeSet::new(),
                                        Some(self.module),
                                        ConstantContext::Enumeration,
                                    )
                                    .map_err(|kind| self.failure(kind, expression.span()))?
                                    .ok_or_else(|| {
                                        self.failure(
                                            CompileErrorKind::ConstantExpressionRequired,
                                            expression.span(),
                                        )
                                    })?;
                                let value = crate::runtime::coerce_declared(
                                    value,
                                    "Long",
                                    enumeration.span.line,
                                    2000,
                                )
                                .map_err(|error| {
                                    self.failure(
                                        Self::constant_failure(error.kind),
                                        expression.span(),
                                    )
                                })?;
                                crate::runtime::integer_argument(
                                    &value,
                                    Some(enumeration.span.line),
                                )
                                .map_err(|error| {
                                    self.failure(
                                        Self::constant_failure(error.kind),
                                        expression.span(),
                                    )
                                })?
                            }
                            None => i64::from((previous as i32).wrapping_add(1)),
                        };
                    }
                }
                ModuleItem::Type(record) => {
                    for field in &record.fields {
                        let mut field_checker =
                            Self::for_context(self.module, self.globals, self.classes, None);
                        field_checker.project_name = self.project_name;
                        if let (Some(types), Some(symbols), Some(namespaces)) =
                            (self.known_types, self.known_symbols, self.known_namespaces)
                        {
                            field_checker.install_references(ReferenceContext {
                                strict: true,
                                project_name: self.project_name,
                                types,
                                symbols,
                                namespaces,
                                documents: self
                                    .document_modules
                                    .expect("installed document context"),
                                current_year: self.compilation_year.unwrap_or(2000),
                                library: self
                                    .reference_library
                                    .expect("installed reference library"),
                            });
                        }
                        field_checker.statement(&Statement::Dim(VarDecl {
                            visibility: Visibility::Private,
                            is_const: false,
                            is_static: false,
                            items: vec![field.clone()],
                            span: record.span,
                        }))?;
                    }
                }
                ModuleItem::ExternalProc(procedure) => {
                    for param in &procedure.params {
                        self.validate_type(&param.type_name, procedure.span)?;
                    }
                    if let Some(kind) = &procedure.return_type {
                        self.validate_type(kind, procedure.span)?;
                    }
                }
                _ => {}
            }
        }
        Ok(())
    }

    fn constant_syntax(&self, expression: &Expr) -> Result<(), CompileError> {
        self.check_expression_syntax(expression, false)
            .map_err(|mut failure| {
                expression.visit(&mut |part| {
                    if part.span() == failure.span
                        && matches!(part, Expr::Index { target, args, .. }
                    if args.is_empty() && self.intrinsic_unary(target))
                    {
                        failure.kind = CompileErrorKind::ExpectedExpression;
                    }
                });
                failure
            })
    }

    fn variable_syntax(&self, variable: &VarItem, span: Span) -> Result<(), CompileError> {
        if intrinsic_operator(&key(&variable.name)) {
            return Err(self.failure(CompileErrorKind::ExpectedIdentifier, span));
        }
        if self.reserved_name(&variable.name) {
            return Err(self.failure(CompileErrorKind::SyntaxError, span));
        }
        if let Some(length) = &variable.type_name.fixed_length {
            self.constant_syntax(length)?;
        }
        if let Some(bounds) = &variable.array_bounds {
            for bound in bounds {
                if let Some(lower) = &bound.lower {
                    self.constant_syntax(lower)?;
                }
                self.constant_syntax(&bound.upper)?;
            }
        }
        if let Some(value) = &variable.value {
            self.constant_syntax(value)?;
        }
        Ok(())
    }

    /// Declaration grammar is read before declaration values/types are bound.
    /// A procedure's inline body shares its header's logical declaration line.
    fn declaration_syntax(&self) -> Result<(), CompileError> {
        if self.known_types.is_none() {
            return Ok(());
        }
        for item in &self.module.items {
            match item {
                ModuleItem::Variables(declaration) => {
                    for variable in &declaration.items {
                        self.variable_syntax(variable, declaration.span)?;
                    }
                }
                ModuleItem::Enum(enumeration) => {
                    for (_, value) in &enumeration.members {
                        if let Some(value) = value {
                            self.constant_syntax(value)?;
                        }
                    }
                }
                ModuleItem::Type(record) => {
                    for field in &record.fields {
                        self.variable_syntax(field, record.span)?;
                    }
                }
                ModuleItem::Procedure(procedure) => {
                    if let Some(span) = procedure
                        .syntax_errors
                        .iter()
                        .find(|span| self.logical_line(**span) == self.logical_line(procedure.span))
                    {
                        return Err(self.failure(CompileErrorKind::SyntaxError, *span));
                    }
                    if intrinsic_operator(&key(&procedure.name)) {
                        return Err(self.failure(CompileErrorKind::SyntaxError, procedure.span));
                    }
                    for parameter in &procedure.params {
                        if self.reserved_name(&parameter.name) {
                            return Err(self.failure(CompileErrorKind::SyntaxError, procedure.span));
                        }
                        if let Some(value) = &parameter.default {
                            self.check_expression_syntax(value, false)?;
                        }
                    }
                    let mut roots = Vec::new();
                    collect_syntax_roots(&procedure.body, &mut roots);
                    let header = self.logical_line(procedure.span);
                    for (expression, bare_call) in roots {
                        if self.logical_line(expression.span()) == header {
                            self.check_expression_syntax(expression, bare_call)?;
                        }
                    }
                }
                _ => {}
            }
        }
        Ok(())
    }

    fn declared(&self, name: &TypeName) -> Type {
        let name_key = self.source_type_key(&name.name);
        let kind = if self.records.contains_key(&name_key) {
            Type::Record(name_key)
        } else if self.classes.contains_key(&name_key)
            || self
                .document_modules
                .is_some_and(|documents| documents.contains_key(&name_key))
        {
            Type::Class(name_key)
        } else if name_key == "string" {
            Type::Text
        } else if self.enums.contains(&name_key)
            || self
                .known_types
                .is_some_and(|types| types.get(&name_key) == Some(&true))
            || matches!(
                name_key.as_str(),
                "byte"
                    | "integer"
                    | "long"
                    | "longlong"
                    | "longptr"
                    | "single"
                    | "double"
                    | "currency"
                    | "decimal"
                    | "date"
                    | "boolean"
            )
        {
            Type::Scalar
        } else if name_key == "variant" {
            Type::Variant
        } else if let Some((name, closed)) = self
            .reference_library
            .and_then(|library| library.resolve_type(&name_key))
        {
            Type::Reference(name, closed)
        } else {
            Type::Object
        };
        if name.is_array {
            Type::Array(Box::new(kind))
        } else {
            kind
        }
    }

    fn declaration(&mut self, declaration: &VarDecl) {
        self.declaration_in(declaration, None);
    }

    fn declaration_in(&mut self, declaration: &VarDecl, owner: Option<&'a Module>) {
        for variable in &declaration.items {
            let mut kind = self.declared(&variable.type_name);
            if declaration.is_const && kind == Type::Variant {
                if let Some(value) = &variable.value {
                    kind = self.infer(value);
                }
            }
            if variable.array_bounds.is_some() {
                kind = Type::Array(Box::new(kind));
            }
            self.fixed_arrays.remove(&key(&variable.name));
            if variable
                .array_bounds
                .as_ref()
                .is_some_and(|bounds| !bounds.is_empty())
            {
                self.fixed_arrays.insert(key(&variable.name));
            }
            self.variables.insert(key(&variable.name), kind);
            if declaration.is_const {
                self.constants.insert(key(&variable.name));
                if let Some(value) = &variable.value {
                    self.constant_expressions.insert(
                        key(&variable.name),
                        ConstantBinding {
                            expression: value.clone(),
                            declared: variable.type_name.clone(),
                            owner,
                            enumeration: None,
                        },
                    );
                }
            } else {
                self.constant_expressions.remove(&key(&variable.name));
                self.constants.remove(&key(&variable.name));
            }
        }
    }

    fn definition_scope(&self, source: &'a Module, span: Span) -> &'a Module {
        if std::ptr::eq(source, self.globals) {
            for (_, scope) in &self.globals.standard_modules {
                if scope.items.iter().any(|item| match item {
                    ModuleItem::Variables(found) => found.span.start == span.start,
                    ModuleItem::Enum(found) => found.span.start == span.start,
                    _ => false,
                }) {
                    return scope;
                }
            }
        }
        source
    }

    fn scoped_constant(
        scope: &'a Module,
        name: &str,
        private: bool,
    ) -> Option<ConstantBinding<'a>> {
        for item in &scope.items {
            match item {
                ModuleItem::Variables(declaration)
                    if declaration.is_const
                        && (private
                            || matches!(
                                declaration.visibility,
                                Visibility::Public | Visibility::Global
                            )) =>
                {
                    if let Some(variable) = declaration
                        .items
                        .iter()
                        .find(|v| v.name.eq_ignore_ascii_case(name))
                    {
                        return variable.value.as_ref().map(|expression| ConstantBinding {
                            expression: expression.clone(),
                            declared: variable.type_name.clone(),
                            owner: Some(scope),
                            enumeration: None,
                        });
                    }
                }
                ModuleItem::Enum(enumeration)
                    if private || enumeration.visibility != Visibility::Private =>
                {
                    let mut previous = Expr::Literal(Literal::Number(-1.0), enumeration.span);
                    for (member, value) in &enumeration.members {
                        let expression = value.clone().unwrap_or_else(|| previous.clone());
                        let qualified = format!("{}.{}", enumeration.name, member);
                        if member.eq_ignore_ascii_case(name) || qualified.eq_ignore_ascii_case(name)
                        {
                            return Some(ConstantBinding {
                                expression,
                                declared: TypeName {
                                    name: "Long".into(),
                                    ..TypeName::implicit()
                                },
                                owner: Some(scope),
                                enumeration: Some(value.is_none()),
                            });
                        }
                        previous = Expr::Member {
                            object: Box::new(Expr::Ident(
                                enumeration.name.clone(),
                                enumeration.span,
                            )),
                            name: member.clone(),
                            suffix: None,
                            span: enumeration.span,
                        };
                    }
                }
                _ => {}
            }
        }
        None
    }

    fn constant_value(
        &self,
        expression: &Expr,
        visiting: &mut BTreeSet<String>,
    ) -> Result<Option<Value>, CompileErrorKind> {
        self.constant_value_in(expression, visiting, None, ConstantContext::Bound)
    }

    fn constant_failure(kind: RuntimeErrorKind) -> CompileErrorKind {
        match kind {
            RuntimeErrorKind::Overflow => CompileErrorKind::Overflow,
            RuntimeErrorKind::DivisionByZero => CompileErrorKind::DivisionByZero,
            RuntimeErrorKind::InvalidNull => CompileErrorKind::InvalidNull,
            _ => CompileErrorKind::TypeMismatch,
        }
    }

    fn coerce_constant(
        &self,
        value: Value,
        declared: &TypeName,
        span: Span,
    ) -> Result<Option<Value>, CompileErrorKind> {
        if declared.name.eq_ignore_ascii_case("object") {
            return Err(CompileErrorKind::InvalidConstantType);
        }
        if matches!(value, Value::Nothing) {
            return Err(CompileErrorKind::InvalidObjectUse);
        }
        if declared.name.eq_ignore_ascii_case("date") {
            // Text dates need the execution year's parsing context. Numeric
            // dates and Null have context-independent conversion diagnostics.
            if matches!(value, Value::String(_)) && self.compilation_year.is_none() {
                return Ok(None);
            }
            let value = crate::runtime::coerce_declared(
                value,
                &declared.name,
                span.line,
                self.compilation_year.unwrap_or(2000),
            )
            .map_err(|error| Self::constant_failure(error.kind))?;
            return crate::runtime::call_builtin("CDate", &[value], Some(span.line), false, 2000)
                .expect("CDate is an intrinsic")
                .map(Some)
                // The value is already a Date; a failed calendar-range check
                // is overflow, rather than a source-type conversion failure.
                .map_err(|_| CompileErrorKind::Overflow);
        }
        crate::runtime::coerce_declared(value, &declared.name, span.line, 2000)
            .map(Some)
            .map_err(|error| Self::constant_failure(error.kind))
    }

    fn coerce_constant_in(
        &self,
        value: Value,
        declared: &TypeName,
        span: Span,
        owner: Option<&'a Module>,
        visiting: &mut BTreeSet<String>,
    ) -> Result<Option<Value>, CompileErrorKind> {
        let Some(value) = self.coerce_constant(value, declared, span)? else {
            return Ok(None);
        };
        let Some(length) = &declared.fixed_length else {
            return Ok(Some(value));
        };
        let Some(width) =
            self.constant_value_in(length, visiting, owner, ConstantContext::Bound)?
        else {
            return Ok(None);
        };
        if matches!(width, Value::LongLong(_)) {
            return Err(CompileErrorKind::InvalidFixedStringLength);
        }
        let width = crate::runtime::integer_argument(&width, Some(span.line))
            .map_err(|error| Self::constant_failure(error.kind))?;
        if !(1..=65_526).contains(&width) {
            return Err(CompileErrorKind::InvalidFixedStringLength);
        }
        crate::runtime::coerce_string_width(value, width as usize, Some(span.line))
            .map(Some)
            .map_err(|error| Self::constant_failure(error.kind))
    }

    fn constant_value_in(
        &self,
        expression: &Expr,
        visiting: &mut BTreeSet<String>,
        owner: Option<&'a Module>,
        context: ConstantContext,
    ) -> Result<Option<Value>, CompileErrorKind> {
        use crate::runtime;
        let failure = Self::constant_failure;
        let is_variable = |name: &str| match owner {
            Some(scope) => scope.items.iter().any(|item| {
                matches!(item,
                ModuleItem::Variables(declaration) if !declaration.is_const
                    && declaration.items.iter().any(|v| v.name.eq_ignore_ascii_case(name)))
            }),
            None => self.variables.contains_key(&key(name)),
        };
        let mut binding = expression.dotted_name().and_then(|name| match owner {
            Some(scope) => Self::scoped_constant(scope, &name, true),
            None => self.constant_expressions.get(&key(&name)).cloned(),
        });
        if binding.is_none() && owner.is_some() {
            if let Some(name) = expression.dotted_name().filter(|name| !is_variable(name)) {
                if let Some(Some(value)) = self
                    .known_symbols
                    .and_then(|symbols| symbols.get(&key(&name)))
                {
                    binding = Some(ConstantBinding {
                        expression: Expr::Literal(
                            Literal::Number(*value as f64),
                            expression.span(),
                        ),
                        declared: TypeName {
                            name: "Long".into(),
                            ..TypeName::implicit()
                        },
                        owner: None,
                        enumeration: None,
                    });
                }
            }
        }
        if let Expr::Member { object, name, .. } = expression {
            if let Some(qualifier) = bare_name(object) {
                if is_variable(qualifier) {
                    return Err(CompileErrorKind::ConstantExpressionRequired);
                }
                if let Some((_, scope)) = self
                    .globals
                    .standard_modules
                    .iter()
                    .find(|(found, _)| found.eq_ignore_ascii_case(qualifier))
                {
                    binding = Self::scoped_constant(
                        scope,
                        name,
                        std::ptr::eq(scope, owner.unwrap_or(self.module)),
                    );
                    if binding.is_none() {
                        return Err(CompileErrorKind::ConstantExpressionRequired);
                    }
                }
            }
        }
        if let Some(binding) = binding {
            let name = expression.dotted_name().unwrap_or_default();
            let identity = format!(
                "{}:{:p}:{}",
                if binding.owner.is_some() {
                    "module"
                } else {
                    "local"
                },
                binding.owner.unwrap_or(self.module),
                key(&name)
            );
            if !visiting.insert(identity.clone()) {
                return Err(CompileErrorKind::ConstantExpressionRequired);
            }
            let mut value = self.constant_value_in(
                &binding.expression,
                visiting,
                binding.owner,
                if binding.enumeration.is_some() {
                    ConstantContext::Enumeration
                } else {
                    ConstantContext::Initializer
                },
            )?;
            if binding.enumeration == Some(true) {
                value = value
                    .map(|value| {
                        runtime::integer_argument(&value, Some(expression.span().line))
                            .map(|previous| {
                                Value::Integer(i64::from((previous as i32).wrapping_add(1)))
                            })
                            .map_err(|error| Self::constant_failure(error.kind))
                    })
                    .transpose()?;
            }
            let result = match value {
                Some(value) => self.coerce_constant_in(
                    value,
                    &binding.declared,
                    expression.span(),
                    binding.owner,
                    visiting,
                ),
                None => Ok(None),
            };
            visiting.remove(&identity);
            return result;
        }
        match expression {
            Expr::Literal(value, _) => Ok(Some(runtime::literal_value(value))),
            Expr::Group { expression, .. } => {
                self.constant_value_in(expression, visiting, owner, context)
            }
            Expr::Ident(name, _) | Expr::TypedIdent { name, .. } => {
                if is_variable(name) {
                    return Err(CompileErrorKind::ConstantExpressionRequired);
                }
                if let Some(value) = runtime::builtin_constant(name) {
                    return Ok(Some(value));
                }
                let symbol = key(name);
                if intrinsic_operator(&symbol) {
                    return Err(CompileErrorKind::SyntaxError);
                }
                if symbol == "second" {
                    return Err(CompileErrorKind::ArgumentNotOptional);
                }
                if matches!(symbol.as_str(), "date" | "now" | "time" | "timer") {
                    return Err(CompileErrorKind::ConstantExpressionRequired);
                }
                if owner.is_none() && self.later_locals.contains(&symbol) {
                    return Err(CompileErrorKind::VariableNotDefined);
                }
                if self.functions.contains_key(&symbol) {
                    return Err(CompileErrorKind::ConstantExpressionRequired);
                }
                let explicit = owner
                    .map(|scope| {
                        scope.items.iter().any(|item| {
                            matches!(item, ModuleItem::Option(ModuleOption::Explicit, _))
                        })
                    })
                    .unwrap_or(self.explicit);
                if explicit {
                    return Err(CompileErrorKind::VariableNotDefined);
                }
                Err(CompileErrorKind::ConstantExpressionRequired)
            }
            Expr::Unary { op, operand, .. } => {
                let Some(value) = self.constant_value_in(operand, visiting, owner, context)? else {
                    return Ok(None);
                };
                runtime::unary(*op, value).map(Some).map_err(|message| {
                    if message.to_ascii_lowercase().contains("overflow") {
                        CompileErrorKind::Overflow
                    } else {
                        CompileErrorKind::TypeMismatch
                    }
                })
            }
            Expr::Binary { op, lhs, rhs, .. } => {
                let lhs = self.constant_value_in(lhs, visiting, owner, context)?;
                let rhs = self.constant_value_in(rhs, visiting, owner, context)?;
                let (Some(lhs), Some(rhs)) = (lhs, rhs) else {
                    return Ok(None);
                };
                let scope = owner.unwrap_or(self.module);
                let compare_text = scope.items.iter().any(|item|
                    matches!(item, ModuleItem::Option(ModuleOption::Compare(mode), _) if mode.eq_ignore_ascii_case("text")));
                runtime::binary(*op, lhs, rhs, compare_text)
                    .map(Some)
                    .map_err(|(kind, _)| failure(kind))
            }
            Expr::Index { target, args, .. } => {
                let Some(name) = bare_name(target) else {
                    return Err(CompileErrorKind::ConstantExpressionRequired);
                };
                let symbol = key(name);
                if intrinsic_operator(&symbol) && args.is_empty() {
                    return Err(CompileErrorKind::SyntaxError);
                }
                if symbol == "second" && args.is_empty() {
                    return Err(CompileErrorKind::ArgumentNotOptional);
                }
                if context == ConstantContext::Enumeration && self.functions.contains_key(&symbol) {
                    return Err(CompileErrorKind::CircularModuleReference);
                }
                if context == ConstantContext::Initializer
                    || is_variable(name)
                    || self.functions.contains_key(&symbol)
                {
                    return Err(CompileErrorKind::ConstantExpressionRequired);
                }
                if matches!(
                    symbol.as_str(),
                    "date" | "now" | "time" | "timer" | "sqr" | "left" | "chr" | "hex" | "round"
                ) {
                    return Err(CompileErrorKind::ConstantExpressionRequired);
                }
                if !matches!(
                    symbol.as_str(),
                    "abs"
                        | "len"
                        | "lenb"
                        | "cbyte"
                        | "cint"
                        | "clng"
                        | "clnglng"
                        | "clngptr"
                        | "csng"
                        | "cdbl"
                        | "ccur"
                        | "cdec"
                        | "cbool"
                        | "int"
                        | "fix"
                        | "sgn"
                        | "cstr"
                        | "cvar"
                        | "cdate"
                ) {
                    return if self.known_types.is_some() {
                        Err(CompileErrorKind::ConstantExpressionRequired)
                    } else {
                        Ok(None)
                    };
                }
                let mut values = Vec::with_capacity(args.len());
                for argument in args {
                    let Some(value) = &argument.value else {
                        return Ok(None);
                    };
                    let Some(value) = self.constant_value_in(value, visiting, owner, context)?
                    else {
                        return Ok(None);
                    };
                    values.push(value);
                }
                if symbol == "cdate"
                    && self.compilation_year.is_none()
                    && values.iter().any(|value| matches!(value, Value::String(_)))
                {
                    return Ok(None);
                }
                // Text dates use the explicit compilation year when available.
                runtime::call_builtin(
                    name,
                    &values,
                    Some(expression.span().line),
                    false,
                    self.compilation_year.unwrap_or(2000),
                )
                .map(|answer| answer.map(Some).map_err(|error| failure(error.kind)))
                .unwrap_or(Ok(None))
            }
            _ if context == ConstantContext::Initializer => {
                Err(CompileErrorKind::ConstantExpressionRequired)
            }
            _ => Ok(None),
        }
    }

    fn constant_integer(&self, expression: &Expr) -> Result<Option<i64>, CompileError> {
        let value = self
            .constant_value(expression, &mut BTreeSet::new())
            .map_err(|kind| self.failure(kind, expression.span()))?;
        value
            .map(|value| {
                crate::runtime::integer_argument(&value, Some(expression.span().line)).map_err(
                    |error| self.failure(Self::constant_failure(error.kind), expression.span()),
                )
            })
            .transpose()
    }

    fn field(&self, receiver: Type, name: &str) -> Type {
        match receiver {
            Type::Reference(owner, _) => self
                .reference_library
                .and_then(|library| library.member(&owner, name))
                .map(|s| self.reference_result(s))
                .unwrap_or(Type::Unknown),
            Type::Class(class) => self
                .classes
                .get(&class)
                .and_then(|module| {
                    module.items.iter().find_map(|item| match item {
                        ModuleItem::Variables(declaration)
                            if matches!(
                                declaration.visibility,
                                Visibility::Public | Visibility::Global
                            ) =>
                        {
                            declaration
                                .items
                                .iter()
                                .find(|field| field.name.eq_ignore_ascii_case(name))
                                .map(|field| {
                                    let kind = self.declared(&field.type_name);
                                    if field.array_bounds.is_some() {
                                        Type::Array(Box::new(kind))
                                    } else {
                                        kind
                                    }
                                })
                        }
                        ModuleItem::Procedure(procedure)
                            if procedure.name.eq_ignore_ascii_case(name)
                                && procedure.visibility != Visibility::Private
                                && matches!(
                                    procedure.kind,
                                    ProcKind::Function | ProcKind::PropertyGet
                                ) =>
                        {
                            Some(
                                procedure
                                    .return_type
                                    .as_ref()
                                    .map(|kind| self.declared(kind))
                                    .unwrap_or(Type::Variant),
                            )
                        }
                        _ => None,
                    })
                })
                .or_else(|| {
                    self.document_member(&class, name)
                        .map(|s| self.reference_result(s))
                })
                .unwrap_or(Type::Unknown),
            Type::Record(record) => self
                .records
                .get(&record)
                .and_then(|record| {
                    record
                        .fields
                        .iter()
                        .find(|field| field.name.eq_ignore_ascii_case(name))
                })
                .map(|field| {
                    let kind = self.declared(&field.type_name);
                    if field.array_bounds.is_some() {
                        Type::Array(Box::new(kind))
                    } else {
                        kind
                    }
                })
                .unwrap_or(Type::Unknown),
            _ => Type::Unknown,
        }
    }

    fn class_member_visible(&self, class: &str, name: &str) -> bool {
        self.document_member(class, name).is_some()
            || self.classes.get(class).is_some_and(|module| {
                module.items.iter().any(|item| match item {
                    ModuleItem::Variables(declaration) => {
                        matches!(
                            declaration.visibility,
                            Visibility::Public | Visibility::Global
                        ) && declaration
                            .items
                            .iter()
                            .any(|field| field.name.eq_ignore_ascii_case(name))
                    }
                    ModuleItem::Procedure(procedure) => {
                        procedure.name.eq_ignore_ascii_case(name)
                            && procedure.visibility != Visibility::Private
                    }
                    _ => false,
                })
            })
    }

    fn document_member(&self, class: &str, name: &str) -> Option<&'a ReferenceSignature> {
        let base = self.document_modules?.get(class)?;
        self.reference_library?.member(base, name)
    }

    fn implicit_document_member(&self, name: &str) -> Option<&'a ReferenceSignature> {
        if self.document_modules?.is_empty() {
            return None;
        }
        let symbol = key(name);
        if self.variables.contains_key(&symbol)
            || self.functions.contains_key(&symbol)
            || self.later_locals.contains(&symbol)
        {
            return None;
        }
        let (class, _) = self
            .classes
            .iter()
            .find(|(_, module)| std::ptr::eq(**module, self.module))?;
        self.document_member(class, name)
    }

    fn reference_read_only(&self, target: &Expr) -> bool {
        let Some(library) = self.reference_library else {
            return false;
        };
        let (kind, name) = match target {
            Expr::Member { object, name, .. } => (self.infer(object), name.as_str()),
            Expr::WithMember(name, _) => (
                self.with_types.last().cloned().unwrap_or(Type::Unknown),
                name.as_str(),
            ),
            Expr::Ident(name, _) if self.implicit_document_member(name).is_some() => {
                let Some((class, _)) = self
                    .classes
                    .iter()
                    .find(|(_, module)| std::ptr::eq(**module, self.module))
                else {
                    return false;
                };
                (Type::Class(class.clone()), name.as_str())
            }
            Expr::Index { target, .. } => return self.reference_read_only(target),
            _ => return false,
        };
        let owner = match kind {
            Type::Class(class) => self
                .document_modules
                .and_then(|documents| documents.get(&class))
                .cloned(),
            Type::Reference(owner, _) => Some(owner),
            _ => None,
        };
        let Some(owner) = owner else {
            return false;
        };
        library.member(&owner, name).is_some_and(|signature| {
            signature.invocation == 2
                && matches!(self.reference_result(signature), Type::Scalar | Type::Text)
                && !library.member_writable(&owner, name)
        })
    }

    fn default_procedure(&self, class: &str, kind: ProcKind) -> Option<&'a Procedure> {
        let module = *self.classes.get(class)?;
        let name = module.items.iter().find_map(|item| {
            let ModuleItem::Procedure(p) = item else {
                return None;
            };
            p.body.iter().find_map(|statement| {
                let Statement::Comment { text, .. } = statement else {
                    return None;
                };
                let (member, value) = text.strip_prefix("Attribute ")?.split_once('=')?;
                let (owner, attribute) = member.trim().split_once('.')?;
                (attribute.eq_ignore_ascii_case("VB_UserMemId") && value.trim() == "0")
                    .then(|| owner.trim())
            })
        })?;
        module.items.iter().find_map(|item| match item {
            ModuleItem::Procedure(p) if p.kind == kind && p.name.eq_ignore_ascii_case(name) => {
                Some(p)
            }
            _ => None,
        })
    }

    fn intrinsic_index_result(&self, target: &Expr, args: &[Argument]) -> Option<Type> {
        let name = key(bare_name(target)?);
        if self.variables.contains_key(&name) || self.functions.contains_key(&name)
            || self.later_locals.contains(&name)
        {
            return None;
        }
        // Typelib dispatch signatures are less precise than VBA's intrinsic
        // operators. Abs(Long), for example, has a scalar result even though
        // the library's generic function signature returns Variant.
        Some(match name.as_str() {
            "cbyte" | "cint" | "clng" | "clnglng" | "clngptr" | "csng" | "cdbl"
                | "ccur" | "cdate" | "cbool" | "vartype" => Type::Scalar,
            "cstr" | "typename" => Type::Text,
            "abs" | "int" | "fix" => match args.first().and_then(|arg| arg.value.as_ref()).map(|value| self.infer(value)) {
                Some(Type::Scalar | Type::Text) => Type::Scalar,
                _ => Type::Variant,
            },
            "len" | "lenb" => match args.first().and_then(|arg| arg.value.as_ref()).map(|value| self.infer(value)) {
                Some(Type::Scalar | Type::Text | Type::Record(_)) => Type::Scalar,
                _ => Type::Variant,
            },
            "cvar" | "cdec" | "cverr" | "array" => Type::Variant,
            "createobject" | "getobject" => Type::Object,
            _ => return None,
        })
    }

    fn infer(&self, expr: &Expr) -> Type {
        match expr {
            Expr::Group { expression, .. } => match self.infer(expression) {
                Type::Class(class) => self
                    .default_procedure(&class, ProcKind::PropertyGet)
                    .and_then(|p| p.return_type.as_ref())
                    .map(|kind| self.declared(kind))
                    .unwrap_or(Type::Variant),
                Type::Object => Type::Variant,
                other => other,
            },
            Expr::Literal(Literal::Empty | Literal::Null, _) => Type::Variant,
            Expr::Literal(Literal::Nothing, _) => Type::Object,
            Expr::New { type_name, .. } => {
                if let Some(class) = self.source_class_key(type_name) {
                    Type::Class(class)
                } else if let Some((name, closed)) = self
                    .reference_library
                    .and_then(|library| library.resolve_type(type_name))
                {
                    Type::Reference(name, closed)
                } else {
                    Type::Object
                }
            }
            Expr::Literal(Literal::Str(_), _) => Type::Text,
            Expr::Literal(..) | Expr::AddressOf { .. } => Type::Scalar,
            Expr::Ident(name, _) => self
                .variables
                .get(&key(name))
                .or_else(|| self.functions.get(&key(name)))
                .cloned()
                .unwrap_or_else(|| {
                    if let Some(signature) = self.implicit_document_member(name) {
                        return self.reference_result(signature);
                    }
                    if self
                        .document_modules
                        .is_some_and(|documents| documents.contains_key(&key(name)))
                    {
                        return Type::Class(key(name));
                    }
                    if let Some((name, closed)) = self
                        .reference_library
                        .and_then(|library| library.global(name))
                    {
                        return Type::Reference(name, closed);
                    }
                    if let Some(signature) = self.reference_function(expr) {
                        return self.reference_result(signature);
                    }
                    if name.eq_ignore_ascii_case("timer") {
                        return Type::Scalar;
                    }
                    match crate::runtime::builtin_constant(name) {
                        Some(crate::Value::String(_)) => Type::Text,
                        Some(crate::Value::Empty | crate::Value::Null | crate::Value::Missing) => {
                            Type::Variant
                        }
                        Some(_) => Type::Scalar,
                        None => Type::Unknown,
                    }
                }),
            Expr::TypedIdent { suffix: '$', .. } => Type::Text,
            Expr::TypedIdent { .. } => Type::Scalar,
            Expr::Member { object, name, .. } | Expr::Bang { object, name, .. } => {
                if let Some(signature) = self.reference_function(expr) {
                    return self.reference_result(signature);
                }
                self.field(self.infer(object), name)
            }
            Expr::WithMember(name, _) | Expr::WithBangMember(name, _) => self.field(
                self.with_types.last().cloned().unwrap_or(Type::Unknown),
                name,
            ),
            Expr::Index { target, args, .. } => {
                if let Some(kind) = self.intrinsic_index_result(target, args) {
                    return kind;
                }
                if let Some(signature) = self.reference_index_signature(target, args) {
                    return self.reference_result(signature);
                }
                if let Some(result) = self.source_indexed_result(target, args) {
                    return result;
                }
                if self.known_types.is_some() {
                    if let Ok(Some(signature)) = self.source_signature(target, false) {
                        return self.reference_result(&signature);
                    }
                }
                if let Some(name) = bare_name(target) {
                    if let Some(kind) = self.variables.get(&key(name)) {
                        if args.is_empty() && matches!(kind, Type::Array(_)) {
                            return kind.clone();
                        }
                        return match kind {
                            Type::Array(element) => *element.clone(),
                            _ => Type::Unknown,
                        };
                    }
                    if let Some(kind) = self.functions.get(&key(name)) {
                        return kind.clone();
                    }
                    return Type::Unknown;
                }
                match target.as_ref() {
                    Expr::Member { object, .. } if matches!(self.infer(object), Type::Class(_)) => {
                        match self.infer(target) {
                            Type::Array(element) => *element,
                            kind => kind,
                        }
                    }
                    Expr::WithMember(..) | Expr::WithBangMember(..)
                        if matches!(self.with_types.last(), Some(Type::Class(_))) =>
                    {
                        match self.infer(target) {
                            Type::Array(element) => *element,
                            kind => kind,
                        }
                    }
                    _ => match self.infer(target) {
                        Type::Array(element) => *element,
                        _ => Type::Unknown,
                    },
                }
            }
            Expr::Unary { operand, .. } => match self.infer(operand) {
                Type::Text => Type::Scalar,
                kind => kind,
            },
            Expr::Binary { op, lhs, rhs, .. } => {
                let left = self.infer(lhs);
                let right = self.infer(rhs);
                if matches!(left, Type::Scalar | Type::Text)
                    && matches!(right, Type::Scalar | Type::Text)
                {
                    if *op == BinaryOp::Concat
                        || (*op == BinaryOp::Add && left == Type::Text && right == Type::Text)
                    {
                        Type::Text
                    } else {
                        Type::Scalar
                    }
                } else {
                    Type::Variant
                }
            }
            Expr::TypeOf { .. } => Type::Scalar,
            Expr::EvaluateShortcut { .. } => Type::Unknown,
        }
    }

    fn reference_result(&self, signature: &ReferenceSignature) -> Type {
        self.declared(&TypeName {
            name: signature.result.clone(),
            ..TypeName::implicit()
        })
    }

    fn reference_function(&self, target: &Expr) -> Option<&'a ReferenceSignature> {
        let mut name = target.dotted_name()?;
        let root = key(name.split('.').next()?);
        if self.variables.contains_key(&root)
            || self.functions.contains_key(&root)
            || self.later_locals.contains(&root)
            || self
                .globals
                .standard_modules
                .iter()
                .any(|(name, _)| key(name) == root)
        {
            return None;
        }
        if let Expr::TypedIdent { suffix: '$', .. } = target {
            name.push('$');
        }
        self.reference_library?.function(&name)
    }

    fn reference_signature(&self, target: &Expr) -> Option<&'a ReferenceSignature> {
        let library = self.reference_library?;
        if let Expr::Ident(name, _) = target {
            if let Some(signature) = self.implicit_document_member(name) {
                return Some(signature);
            }
        }
        if let Some(signature) = self.reference_function(target) {
            return Some(signature);
        }
        match target {
            Expr::Member { object, name, .. } => match self.infer(object) {
                Type::Reference(owner, _) => library.member(&owner, name),
                Type::Class(owner) => self.document_member(&owner, name),
                _ => None,
            },
            Expr::WithMember(name, _) => match self.with_types.last() {
                Some(Type::Reference(owner, _)) => library.member(owner, name),
                Some(Type::Class(owner)) => self.document_member(owner, name),
                _ => None,
            },
            _ => match self.infer(target) {
                Type::Reference(owner, _) => library.default_member(&owner),
                _ => None,
            },
        }
    }

    fn reference_index_signature(
        &self,
        target: &Expr,
        args: &[Argument],
    ) -> Option<&'a ReferenceSignature> {
        // A source procedure binds its own arguments before its object result
        // can supply a default member. In particular, Factory() is a call,
        // rather than an argumentless invocation of Collection.Item.
        if self.source_signature(target, false).ok().flatten().is_some() {
            return None;
        }
        let signature = self.reference_signature(target)?;
        if signature.invocation == 2 && signature.parameters.is_empty() && !args.is_empty() {
            if let Type::Reference(owner, _) = self.reference_result(signature) {
                if let Some(default) = self.reference_library?.default_member(&owner) {
                    return Some(default);
                }
            }
        }
        Some(signature)
    }

    fn reference_argument_error(
        signature: &ReferenceSignature,
        args: &[Argument],
    ) -> Option<CompileErrorKind> {
        let mut named = false;
        for argument in args {
            if argument.name.is_some() {
                named = true;
            } else if named {
                return Some(CompileErrorKind::SyntaxError);
            }
        }
        if signature.variadic && named {
            return Some(CompileErrorKind::NamedArgumentInParamArray);
        }
        let dispatch = signature.function_kind == 4;
        if !dispatch {
            let mut seen_names = BTreeSet::new();
            for argument in args {
                if let Some(name) = argument.name.as_ref().filter(|name| {
                    signature
                        .parameters
                        .iter()
                        .any(|(p, _)| p.eq_ignore_ascii_case(name))
                }) {
                    if !seen_names.insert(key(name)) {
                        return Some(CompileErrorKind::NamedArgumentAlreadySpecified);
                    }
                }
            }
            if !signature.variadic && args.len() > signature.parameters.len() {
                return Some(CompileErrorKind::WrongNumberOfArguments);
            }
        }
        let mut present = BTreeSet::new();
        for (position, argument) in args.iter().enumerate() {
            let index = if let Some(name) = &argument.name {
                match signature
                    .parameters
                    .iter()
                    .position(|(p, _)| p.eq_ignore_ascii_case(name))
                {
                    Some(index) => index,
                    None => return Some(CompileErrorKind::NamedArgumentNotFound),
                }
            } else {
                position
            };
            if index >= signature.parameters.len() && !signature.variadic {
                return Some(CompileErrorKind::WrongNumberOfArguments);
            }
            if !dispatch
                && argument.value.is_none()
                && signature
                    .parameters
                    .get(index)
                    .is_some_and(|(_, optional)| !*optional)
            {
                return Some(CompileErrorKind::ArgumentNotOptional);
            }
            if argument.value.is_some() {
                let slot = if dispatch { position } else { index };
                if !dispatch && !present.insert(slot) && index < signature.parameters.len() {
                    return Some(CompileErrorKind::NamedArgumentAlreadySpecified);
                }
                present.insert(slot);
            }
        }
        signature
            .parameters
            .iter()
            .enumerate()
            .any(|(index, (_, optional))| {
                !optional
                    && !(signature.variadic && index + 1 == signature.parameters.len())
                    && !present.contains(&index)
            })
            .then_some(CompileErrorKind::ArgumentNotOptional)
    }

    fn intrinsic_unary(&self, target: &Expr) -> bool {
        bare_name(target).is_some_and(|name| {
            !self.variables.contains_key(&key(name))
                && !self.functions.contains_key(&key(name))
                && matches!(
                    key(name).as_str(),
                    "int"
                        | "fix"
                        | "abs"
                        | "sgn"
                        | "len"
                        | "lenb"
                        | "cbool"
                        | "cbyte"
                        | "ccur"
                        | "cdate"
                        | "cdbl"
                        | "cint"
                        | "clng"
                        | "clnglng"
                        | "clngptr"
                        | "csng"
                        | "cstr"
                        | "cvar"
                )
        })
    }

    fn is_call_by_name(&self, expr: &Expr) -> bool {
        if let Some(name) = bare_name(expr) {
            return name.eq_ignore_ascii_case("callbyname")
                && !self.variables.contains_key(&key(name))
                && !self.functions.contains_key(&key(name));
        }
        let Some(name) = expr.dotted_name() else {
            return false;
        };
        let name = key(&name);
        let root = name.split('.').next().unwrap_or_default();
        !self.variables.contains_key(root)
            && !self.functions.contains_key(root)
            && matches!(
                name.as_str(),
                "vba.callbyname" | "vba.interaction.callbyname" | "interaction.callbyname"
            )
    }

    fn source_signature(
        &self,
        target: &Expr,
        write: bool,
    ) -> Result<Option<ReferenceSignature>, CompileErrorKind> {
        let find = |scope: &Module, name: &str| {
            let procedures: Vec<_> = scope
                .items
                .iter()
                .filter_map(|item| match item {
                    ModuleItem::Procedure(p)
                        if p.name.eq_ignore_ascii_case(name)
                            && (std::ptr::eq(scope, self.module)
                                || p.visibility != Visibility::Private) =>
                    {
                        Some(p)
                    }
                    _ => None,
                })
                .collect();
            let setter = procedures
                .iter()
                .copied()
                .find(|p| matches!(p.kind, ProcKind::PropertyLet | ProcKind::PropertySet));
            let reader = procedures.iter().copied().find(|p| {
                matches!(
                    p.kind,
                    ProcKind::Function | ProcKind::Sub | ProcKind::PropertyGet
                )
            });
            let selected = if write { setter.or(reader) } else { reader };
            let Some(selected) = selected else {
                return if setter.is_some() {
                    Err(CompileErrorKind::InvalidPropertyUse)
                } else {
                    Ok(None)
                };
            };
            Ok(Some(Self::procedure_signature(selected, write)))
        };
        if let Some(name) = bare_name(target) {
            if self.variables.contains_key(&key(name))
                && self
                    .procedure
                    .is_none_or(|p| !p.name.eq_ignore_ascii_case(name))
            {
                return Ok(None);
            }
            if let Some(signature) = find(self.module, name)? {
                return Ok(Some(signature));
            }
            for (_, scope) in &self.globals.standard_modules {
                if let Some(signature) = find(scope, name)? {
                    return Ok(Some(signature));
                }
            }
            return Ok(None);
        }
        match target {
            Expr::Member { object, name, .. } => {
                if let Type::Class(class) = self.infer(object) {
                    return self
                        .classes
                        .get(&class)
                        .map_or(Ok(None), |scope| find(scope, name));
                }
                if let Expr::Ident(qualifier, _) = object.as_ref() {
                    if !self.variables.contains_key(&key(qualifier)) {
                        if let Some((_, scope)) = self
                            .globals
                            .standard_modules
                            .iter()
                            .find(|(scope, _)| scope.eq_ignore_ascii_case(qualifier))
                        {
                            return find(scope, name);
                        }
                    }
                }
                Ok(None)
            }
            Expr::WithMember(name, _) => match self.with_types.last() {
                Some(Type::Class(class)) => self
                    .classes
                    .get(class)
                    .map_or(Ok(None), |scope| find(scope, name)),
                _ => Ok(None),
            },
            _ => Ok(None),
        }
    }

    fn procedure_signature(procedure: &Procedure, write: bool) -> ReferenceSignature {
        let setter = matches!(
            procedure.kind,
            ProcKind::PropertyLet | ProcKind::PropertySet
        );
        let count = procedure
            .params
            .len()
            .saturating_sub(usize::from(write && setter));
        let parameters = &procedure.params[..count];
        ReferenceSignature {
            result: if setter || procedure.kind == ProcKind::Sub {
                "void".into()
            } else {
                procedure
                    .return_type
                    .as_ref()
                    .map(|t| key(&t.name))
                    .unwrap_or_else(|| "variant".into())
            },
            parameters: parameters
                .iter()
                .map(|p| (key(&p.name), p.optional))
                .collect(),
            variadic: parameters.iter().any(|p| p.mode == ParamMode::ParamArray),
            invocation: match procedure.kind {
                ProcKind::PropertyGet => 2,
                ProcKind::PropertyLet => 4,
                ProcKind::PropertySet => 8,
                _ => 1,
            },
            function_kind: 3,
        }
    }

    /// A parameterless value-returning procedure can supply the object/array
    /// being indexed. The parentheses then belong to that result, not its call.
    fn source_indexed_owner(&self, target: &Expr, args: &[Argument]) -> Option<Type> {
        if let Ok(Some(signature)) = self.source_signature(target, false) {
            if !args.is_empty() && signature.parameters.is_empty() && signature.result != "void" {
                let kind = self.reference_result(&signature);
                if matches!(
                    kind,
                    Type::Class(_)
                        | Type::Reference(..)
                        | Type::Object
                        | Type::Variant
                        | Type::Array(_)
                ) {
                    return Some(kind);
                }
            }
            return None;
        }
        match self.infer(target) {
            kind @ Type::Class(_) => Some(kind),
            _ => None,
        }
    }

    fn source_default_signature(&self, owner: &Type, write: bool) -> Option<ReferenceSignature> {
        match owner {
            Type::Class(class) => {
                let procedure = if write {
                    self.default_procedure(class, ProcKind::PropertyLet)
                        .or_else(|| self.default_procedure(class, ProcKind::PropertySet))
                        .or_else(|| self.default_procedure(class, ProcKind::PropertyGet))
                } else {
                    self.default_procedure(class, ProcKind::PropertyGet)
                }?;
                Some(Self::procedure_signature(procedure, write))
            }
            Type::Reference(name, _) => self.reference_library?.default_member(name).cloned(),
            _ => None,
        }
    }

    fn source_indexed_result(&self, target: &Expr, args: &[Argument]) -> Option<Type> {
        let owner = self.source_indexed_owner(target, args)?;
        Some(match owner {
            Type::Array(element) => *element,
            _ => self
                .source_default_signature(&owner, false)
                .map(|signature| self.reference_result(&signature))
                .unwrap_or(Type::Variant),
        })
    }

    fn source_index_signature(
        &self,
        target: &Expr,
        args: &[Argument],
        write: bool,
    ) -> Result<Option<ReferenceSignature>, CompileErrorKind> {
        if let Some(owner) = self.source_indexed_owner(target, args) {
            // Named arguments still address the parameterless procedure;
            // positional indices can address its returned default member.
            if self.source_signature(target, false)?.is_some()
                && args.iter().any(|arg| arg.name.is_some())
            {
                return Err(CompileErrorKind::NamedArgumentNotFound);
            }
            return Ok(self.source_default_signature(&owner, write));
        }
        self.source_signature(target, write)
    }

    fn call_parameters(&self, target: &Expr) -> Option<Vec<Param>> {
        let find = |module: &Module, name: &str| {
            module.items.iter().find_map(|item| match item {
                ModuleItem::Procedure(procedure) if procedure.name.eq_ignore_ascii_case(name) => {
                    Some(procedure.params.clone())
                }
                _ => None,
            })
        };
        if let Some(name) = bare_name(target) {
            if self.variables.contains_key(&key(name))
                && self
                    .procedure
                    .is_none_or(|procedure| !name.eq_ignore_ascii_case(&procedure.name))
            {
                return None;
            }
            return find(self.module, name).or_else(|| {
                self.globals
                    .standard_modules
                    .iter()
                    .find_map(|(_, scope)| find(scope, name))
            });
        }
        match target {
            Expr::Member { object, name, .. } => {
                if let Type::Class(class) = self.infer(object) {
                    return self.classes.get(&class).and_then(|scope| find(scope, name));
                }
                if let Expr::Ident(qualifier, _) = object.as_ref() {
                    return self
                        .globals
                        .standard_modules
                        .iter()
                        .find(|(scope, _)| scope.eq_ignore_ascii_case(qualifier))
                        .and_then(|(_, scope)| find(scope, name));
                }
                None
            }
            Expr::WithMember(name, _) => match self.with_types.last() {
                Some(Type::Class(class)) => {
                    self.classes.get(class).and_then(|scope| find(scope, name))
                }
                _ => None,
            },
            _ => None,
        }
    }

    fn is_variable(&self, expr: &Expr) -> bool {
        match expr {
            Expr::Ident(name, _) | Expr::TypedIdent { name, .. } => {
                !self.constants.contains(&key(name))
                    && (self.variables.contains_key(&key(name))
                        || (!self.explicit
                            && !self.functions.contains_key(&key(name))
                            && !name.eq_ignore_ascii_case("timer")
                            && crate::runtime::builtin_constant(name).is_none()))
            }
            Expr::Index { target, .. } => {
                matches!(self.infer(target), Type::Array(_))
                    && !matches!(target.as_ref(),Expr::Member{object,..} if matches!(self.infer(object),Type::Class(_)))
            }
            Expr::Member { object, .. } | Expr::Bang { object, .. } => {
                matches!(self.infer(object), Type::Record(_))
            }
            Expr::WithMember(..) | Expr::WithBangMember(..) => {
                matches!(self.with_types.last(), Some(Type::Record(_)))
            }
            _ => false,
        }
    }

    fn is_paramarray(&self, expr: &Expr) -> bool {
        bare_name(expr).is_some_and(|name| {
            self.procedure
                .into_iter()
                .flat_map(|procedure| &procedure.params)
                .any(|parameter| {
                    parameter.mode == ParamMode::ParamArray
                        && parameter.name.eq_ignore_ascii_case(name)
                })
        })
    }

    fn failure(&self, kind: CompileErrorKind, span: Span) -> CompileError {
        CompileError {
            kind,
            span,
            procedure: self.procedure.map(|p| p.name.clone()).unwrap_or_default(),
            symbol: None,
        }
    }

    fn check_class(&self, name: &str, span: Span) -> Result<(), CompileError> {
        if let Some(module) = self.classes.get(&self.source_type_key(name)) {
            let inherited = |member: &str| {
                inherited_object_member(member)
                    || self.document_member(&key(name), member).is_some()
            };
            for item in &module.items {
                if self.known_types.is_some()
                    || self
                        .document_modules
                        .is_some_and(|documents| documents.contains_key(&key(name)))
                {
                    let invalid = match item {
                        ModuleItem::Procedure(p) if inherited(&p.name) => Some(p.span),
                        ModuleItem::Variables(d)
                            if d.items.iter().any(|field| inherited(&field.name)) =>
                        {
                            Some(d.span)
                        }
                        _ => None,
                    };
                    if let Some(span) = invalid {
                        return Err(self.failure(CompileErrorKind::InheritedMemberIdentifier, span));
                    }
                }
                if let ModuleItem::Variables(declaration) = item {
                    if matches!(
                        declaration.visibility,
                        Visibility::Public | Visibility::Global
                    ) && (declaration.is_const
                        || declaration.items.iter().any(|field| {
                            field.array_bounds.is_some()
                                || field.type_name.is_array
                                || field.type_name.fixed_length.is_some()
                        }))
                    {
                        return Err(self.failure(CompileErrorKind::InvalidPublicClassMember, span));
                    }
                }
            }
        }
        Ok(())
    }

    fn expression(&self, expr: &Expr) -> Result<(), CompileError> {
        self.check_expression(expr, false)
    }

    fn check_expression(&self, expr: &Expr, bare_call: bool) -> Result<(), CompileError> {
        self.check_expression_access(expr, bare_call, false)
    }

    fn check_call_syntax(&self, expr: &Expr, bare_top: bool) -> Result<(), CompileError> {
        let Expr::Index { target, args, .. } = expr else {
            return Ok(());
        };
        let mut named = false;
        if args.iter().any(|arg| {
            if arg.name.is_some() {
                named = true;
                false
            } else {
                named
            }
        }) || (args.last().is_some_and(|arg| arg.value.is_none()) && !bare_top)
            || (self.intrinsic_unary(target)
                && (args.len() != 1 || args[0].name.is_some() || args[0].value.is_none()))
        {
            return Err(self.failure(CompileErrorKind::SyntaxError, expr.span()));
        }
        Ok(())
    }

    fn check_expression_syntax(&self, expr: &Expr, bare_call: bool) -> Result<(), CompileError> {
        if bare_call
            && matches!(expr, Expr::Index { args, force_by_value, .. }
            if args.is_empty() || (*force_by_value && args.len() > 1))
        {
            return Err(self.failure(CompileErrorKind::SyntaxError, expr.span()));
        }
        let mut call_targets = BTreeSet::new();
        expr.visit(&mut |part| {
            if let Expr::Index { target, .. } = part {
                call_targets.insert(target.as_ref() as *const Expr as usize);
            }
        });
        let mut failure = None;
        expr.visit(&mut |part| {
            if failure.is_some() {
                return;
            }
            if let Err(error) = self.check_call_syntax(part, bare_call && std::ptr::eq(part, expr))
            {
                failure = Some(error);
            } else if !call_targets.contains(&(part as *const Expr as usize))
                && self.intrinsic_unary(part)
            {
                failure = Some(self.failure(CompileErrorKind::SyntaxError, part.span()));
            }
        });
        failure.map_or(Ok(()), Err)
    }

    fn logical_line(&self, span: Span) -> usize {
        if self.module.logical_lines.is_empty() {
            span.line as usize
        } else {
            self.module
                .logical_lines
                .partition_point(|line| line.end <= span.start)
        }
    }

    fn check_logical_line_syntax(&self, span: Span) -> Result<(), CompileError> {
        if self.known_types.is_none() {
            return Ok(());
        }
        let line = self.logical_line(span);
        if self.checked_syntax_lines.borrow_mut().insert(line) {
            if let Some(span) = self.line_syntax_errors.get(&line) {
                return Err(self.failure(CompileErrorKind::SyntaxError, *span));
            }
            if let Some(roots) = self.syntax_lines.get(&line) {
                for (expression, bare_call) in roots {
                    self.check_expression_syntax(expression, *bare_call)?;
                }
            }
        }
        Ok(())
    }

    fn resolve_expression_name(
        &self,
        part: &Expr,
        is_call_target: bool,
    ) -> Result<(), CompileError> {
        if let Expr::Member {
            object, name, span, ..
        } = part
        {
            if let Type::Reference(owner, true) = self.infer(object) {
                if self
                    .reference_library
                    .is_some_and(|library| !library.has_member(&owner, name))
                {
                    return Err(self.failure(CompileErrorKind::MemberNotFound, *span));
                }
            }
            if let Type::Class(class) = self.infer(object) {
                if !self.class_member_visible(&class, name) {
                    return Err(self.failure(CompileErrorKind::MemberNotFound, *span));
                }
            }
            if let Expr::Ident(qualifier, _) = object.as_ref() {
                if !self.variables.contains_key(&key(qualifier))
                    && !self.later_locals.contains(&key(qualifier))
                {
                    if let Some((_, scope)) = self
                        .globals
                        .standard_modules
                        .iter()
                        .find(|(found, _)| found.eq_ignore_ascii_case(qualifier))
                    {
                        let own = std::ptr::eq(scope, self.module);
                        let visible = scope.items.iter().any(|item| match item {
                            ModuleItem::Procedure(p) => {
                                p.name.eq_ignore_ascii_case(name)
                                    && (own || p.visibility != Visibility::Private)
                            }
                            ModuleItem::Variables(d) => {
                                (own || matches!(
                                    d.visibility,
                                    Visibility::Public | Visibility::Global
                                )) && d.items.iter().any(|v| v.name.eq_ignore_ascii_case(name))
                            }
                            ModuleItem::Enum(e) => {
                                (own || e.visibility != Visibility::Private)
                                    && (e.name.eq_ignore_ascii_case(name)
                                        || e.members
                                            .iter()
                                            .any(|(member, _)| member.eq_ignore_ascii_case(name)))
                            }
                            _ => false,
                        });
                        if !visible {
                            return Err(self.failure(CompileErrorKind::MemberNotFound, *span));
                        }
                    }
                }
            }
        }
        if let Expr::WithMember(name, span) = part {
            if let Some(Type::Reference(owner, true)) = self.with_types.last() {
                if self
                    .reference_library
                    .is_some_and(|library| !library.has_member(owner, name))
                {
                    return Err(self.failure(CompileErrorKind::MemberNotFound, *span));
                }
            }
            if let Some(Type::Class(class)) = self.with_types.last() {
                if !self.class_member_visible(class, name) {
                    return Err(self.failure(CompileErrorKind::MemberNotFound, *span));
                }
            }
        }
        if let Some(name) = bare_name(part) {
            let symbol = key(name);
            if !self.own_symbols.contains(&symbol) && !self.later_locals.contains(&symbol) {
                if self.ambiguous.contains(&symbol) {
                    let mut found = self.failure(CompileErrorKind::AmbiguousName, part.span());
                    found.symbol = Some(name.to_string());
                    return Err(found);
                }
                if !self.variables.contains_key(&symbol) && !self.functions.contains_key(&symbol) {
                    let kind = if self.hidden_functions.contains(&symbol) {
                        Some(CompileErrorKind::ProcedureNotDefined)
                    } else if self.explicit && self.hidden_variables.contains(&symbol) {
                        Some(CompileErrorKind::VariableNotDefined)
                    } else if self.known_types.is_some()
                        && (self.explicit || is_call_target)
                        && crate::runtime::builtin_constant(name).is_none()
                        && !crate::runtime::is_builtin_function(name)
                        && !self.is_call_by_name(part)
                        && !matches!(
                            symbol.as_str(),
                            "err"
                                | "erl"
                                | "debug"
                                | "vba"
                                | "date"
                                | "now"
                                | "time"
                                | "timer"
                                | "rnd"
                                | "freefile"
                                | "dir"
                                | "curdir"
                        )
                        && !self.enums.contains(&symbol)
                        && !self.classes.contains_key(&symbol)
                        && !self
                            .known_types
                            .is_some_and(|types| types.contains_key(&symbol))
                        && !self
                            .known_symbols
                            .is_some_and(|symbols| symbols.contains_key(&symbol))
                        && !self
                            .known_namespaces
                            .is_some_and(|names| names.contains(&symbol))
                        && !self
                            .document_modules
                            .is_some_and(|documents| documents.contains_key(&symbol))
                        && self.implicit_document_member(name).is_none()
                        && !self
                            .globals
                            .standard_modules
                            .iter()
                            .any(|(module, _)| module.eq_ignore_ascii_case(name))
                    {
                        Some(if is_call_target {
                            CompileErrorKind::ProcedureNotDefined
                        } else {
                            CompileErrorKind::VariableNotDefined
                        })
                    } else {
                        None
                    };
                    if let Some(kind) = kind {
                        return Err(self.failure(kind, part.span()));
                    }
                }
            }
        }
        Ok(())
    }

    fn check_expression_access(
        &self,
        expr: &Expr,
        bare_call: bool,
        write_target: bool,
    ) -> Result<(), CompileError> {
        self.check_logical_line_syntax(expr.span())?;
        if bare_call && matches!(expr,Expr::Index{args,..} if args.is_empty()) {
            return Err(self.failure(CompileErrorKind::SyntaxError, expr.span()));
        }
        if bare_call {
            if let Expr::Index {
                args,
                force_by_value: true,
                span,
                ..
            } = expr
            {
                // Without Call, parentheses group a single argument. They
                // cannot enclose a comma-separated argument list.
                if args.len() > 1 {
                    return Err(self.failure(CompileErrorKind::SyntaxError, *span));
                }
            }
        }
        let mut call_targets = BTreeSet::new();
        expr.visit(&mut |part| {
            if let Expr::Index { target, .. } = part {
                call_targets.insert(target.as_ref() as *const Expr as usize);
            }
        });
        let mut ordered = Vec::new();
        if self.known_types.is_some() {
            self.check_expression_syntax(expr, bare_call)?;
            compilation_expression_order(expr, &mut ordered);
        } else {
            expr.visit(&mut |part| ordered.push(part));
        }
        let mut failure = None;
        ordered.into_iter().for_each(|part| {
            if failure.is_some() {
                return;
            }
            if self.known_types.is_some() {
                if let Expr::Index { target, args, .. } = part {
                    if let Err(error) = self.resolve_expression_name(target, true) {
                        failure = Some(error);
                        return;
                    }
                    if let Some(signature) = self.reference_index_signature(target, args) {
                        if let Some(kind) = Self::reference_argument_error(signature, args)
                            .filter(|_| !self.is_call_by_name(target))
                        {
                            failure = Some(self.failure(kind, part.span()));
                            return;
                        }
                    } else {
                        match self.source_index_signature(
                            target,
                            args,
                            write_target && std::ptr::eq(part, expr),
                        ) {
                            Err(kind) => {
                                failure = Some(self.failure(kind, part.span()));
                                return;
                            }
                            Ok(Some(signature)) => {
                                if let Some(kind) = Self::reference_argument_error(&signature, args)
                                {
                                    failure = Some(self.failure(kind, part.span()));
                                    return;
                                }
                            }
                            Ok(None) => {}
                        }
                    }
                } else if !call_targets.contains(&(part as *const Expr as usize)) {
                    if let Err(error) =
                        self.resolve_expression_name(part, bare_call && std::ptr::eq(part, expr))
                    {
                        failure = Some(error);
                        return;
                    }
                    if self.intrinsic_unary(part) {
                        failure = Some(self.failure(CompileErrorKind::SyntaxError, part.span()));
                        return;
                    }
                    let signature = match part {
                        Expr::Member { .. } | Expr::WithMember(..) => {
                            self.reference_signature(part)
                        }
                        Expr::Ident(..) | Expr::TypedIdent { .. } => self.reference_function(part),
                        _ => None,
                    };
                    if let Some(signature) = signature {
                        if let Some(kind) = Self::reference_argument_error(signature, &[]) {
                            failure = Some(self.failure(kind, part.span()));
                            return;
                        }
                    } else if matches!(part, Expr::Member { .. } | Expr::WithMember(..))
                        || bare_name(part).is_some_and(|name| {
                            !self.variables.contains_key(&key(name))
                                || (bare_call && std::ptr::eq(part, expr))
                        })
                    {
                        match self.source_signature(part, write_target && std::ptr::eq(part, expr))
                        {
                            Err(kind) => {
                                failure = Some(self.failure(kind, part.span()));
                                return;
                            }
                            Ok(Some(signature)) => {
                                if let Some(kind) = Self::reference_argument_error(&signature, &[])
                                {
                                    failure = Some(self.failure(kind, part.span()));
                                    return;
                                }
                            }
                            Ok(None) => {}
                        }
                    }
                }
            }
            let direct_nothing = |value: &Expr| matches!(value, Expr::Literal(Literal::Nothing, _));
            let invalid_nothing = match part {
                Expr::Group { expression, .. } => direct_nothing(expression),
                Expr::Unary { operand, .. } => direct_nothing(operand),
                Expr::Binary { op, lhs, rhs, .. } if *op != BinaryOp::Is => {
                    direct_nothing(lhs) || direct_nothing(rhs)
                }
                Expr::Index { target, args, .. }
                    if bare_name(target).is_some_and(|name| {
                        name.eq_ignore_ascii_case("choose")
                            && !self.functions.contains_key(&key(name))
                    }) =>
                {
                    args.first()
                        .and_then(|arg| arg.value.as_ref())
                        .is_some_and(direct_nothing)
                }
                Expr::Index { target, args, .. }
                    if bare_name(target).is_some_and(|name| {
                        matches!(
                            key(name).as_str(),
                            "cstr"
                                | "clng"
                                | "cbool"
                                | "cint"
                                | "cdbl"
                                | "cbyte"
                                | "ccur"
                                | "cdate"
                                | "csng"
                                | "clnglng"
                                | "clngptr"
                                | "len"
                                | "lenb"
                                | "abs"
                                | "int"
                                | "fix"
                                | "sgn"
                                | "val"
                        ) && !self.functions.contains_key(&key(name))
                    }) =>
                {
                    args.iter()
                        .any(|arg| arg.value.as_ref().is_some_and(direct_nothing))
                }
                _ => false,
            };
            if invalid_nothing {
                failure = Some(self.failure(CompileErrorKind::InvalidObjectUse, part.span()));
                return;
            }
            if let Expr::Group { expression, .. } = part {
                if let Type::Class(class) = self.infer(expression) {
                    if self
                        .default_procedure(&class, ProcKind::PropertyGet)
                        .is_none()
                        && (self
                            .default_procedure(&class, ProcKind::PropertyLet)
                            .is_some()
                            || self
                                .default_procedure(&class, ProcKind::PropertySet)
                                .is_some())
                    {
                        failure =
                            Some(self.failure(CompileErrorKind::InvalidPropertyUse, part.span()));
                        return;
                    }
                }
            }
            if let Expr::New { type_name, span } = part {
                if self
                    .document_modules
                    .is_some_and(|documents| documents.contains_key(&key(type_name)))
                {
                    failure = Some(self.failure(CompileErrorKind::InvalidNewUse, *span));
                    return;
                }
                if let Err(error) = self.check_class(type_name, *span) {
                    failure = Some(error);
                    return;
                }
            }
            if let Err(error) = self.resolve_expression_name(
                part,
                call_targets.contains(&(part as *const Expr as usize))
                    || (bare_call && std::ptr::eq(part, expr)),
            ) {
                failure = Some(error);
                return;
            }
            let Expr::Index {
                target, args, span, ..
            } = part
            else {
                return;
            };
            if let Some(parameters) = self.call_parameters(target) {
                let forced = matches!(
                    part,
                    Expr::Index {
                        force_by_value: true,
                        ..
                    }
                );
                for (index, argument) in args.iter().enumerate() {
                    let parameter = argument
                        .name
                        .as_ref()
                        .and_then(|name| {
                            parameters
                                .iter()
                                .find(|parameter| parameter.name.eq_ignore_ascii_case(name))
                        })
                        .or_else(|| parameters.get(index));
                    if let (Some(parameter), Some(value)) = (parameter, argument.value.as_ref()) {
                        if direct_nothing(value)
                            && matches!(
                                self.declared(&parameter.type_name),
                                Type::Scalar | Type::Text
                            )
                        {
                            failure = Some(
                                self.failure(CompileErrorKind::InvalidObjectUse, value.span()),
                            );
                            return;
                        }
                        if matches!(value, Expr::Group { .. })
                            && matches!(
                                self.declared(&parameter.type_name),
                                Type::Object | Type::Class(_) | Type::Reference(..)
                            )
                            && matches!(self.infer(value), Type::Scalar | Type::Text)
                        {
                            failure =
                                Some(self.failure(CompileErrorKind::TypeMismatch, value.span()));
                            return;
                        }
                    }
                    if (forced || argument.force_by_value)
                        && parameter.is_some_and(|parameter| parameter.is_array)
                    {
                        if let Some(value) = &argument.value {
                            if matches!(self.infer(value), Type::Array(_)) {
                                failure = Some(self.failure(
                                    CompileErrorKind::ArrayOrUserDefinedTypeExpected,
                                    value.span(),
                                ));
                                return;
                            }
                        }
                    }
                }
            }
            // Bare calls also represent Debug.Print's empty output fields.
            // Parenthesized calls require a final argument expression.
            if args.last().is_some_and(|arg| arg.value.is_none())
                && !(bare_call && std::ptr::eq(part, expr))
            {
                failure = Some(self.failure(CompileErrorKind::SyntaxError, *span));
                return;
            }
            if bare_name(target).is_some_and(|name| {
                matches!(key(name).as_str(), "len" | "lenb")
                    && !self.functions.contains_key(&key(name))
                    && !self.variables.contains_key(&key(name))
            }) {
                if let [Argument {
                    value: Some(value), ..
                }] = args.as_slice()
                {
                    if matches!(self.infer(value), Type::Scalar | Type::Record(_))
                        && !self.is_variable(value.without_groups())
                    {
                        failure =
                            Some(self.failure(CompileErrorKind::VariableRequired, value.span()));
                        return;
                    }
                }
            }
            if !self.is_call_by_name(target) {
                return;
            }
            let kind = if args.iter().any(|arg| arg.name.is_some()) {
                Some(CompileErrorKind::NamedArgumentInParamArray)
            } else if args.len() < 3 || args[..3].iter().any(|arg| arg.value.is_none()) {
                Some(CompileErrorKind::ArgumentNotOptional)
            } else {
                let object = args[0].value.as_ref().unwrap();
                let not_yet_declared = bare_name(object).is_some_and(|name| {
                    self.explicit
                        && !self.variables.contains_key(&key(name))
                        && self.later_locals.contains(&key(name))
                });
                if not_yet_declared {
                    Some(CompileErrorKind::VariableNotDefined)
                } else if matches!(
                    self.infer(object),
                    Type::Scalar | Type::Text | Type::Array(_) | Type::Record(_)
                ) {
                    Some(CompileErrorKind::TypeMismatch)
                } else {
                    None
                }
            };
            if let Some(kind) = kind {
                let at = if matches!(
                    kind,
                    CompileErrorKind::TypeMismatch | CompileErrorKind::VariableNotDefined
                ) {
                    args[0].value.as_ref().map(Expr::span).unwrap_or(*span)
                } else {
                    *span
                };
                failure = Some(self.failure(kind, at));
            }
        });
        failure.map_or(Ok(()), Err)
    }

    fn body(&mut self, body: &[Statement]) -> Result<(), CompileError> {
        for statement in body {
            self.statement(statement)?;
        }
        Ok(())
    }

    fn statement(&mut self, statement: &Statement) -> Result<(), CompileError> {
        self.check_logical_line_syntax(statement.span())?;
        if let Statement::Unknown { text, span } = statement {
            let words: Vec<_> = text
                .split_whitespace()
                .filter(|word| *word != "_")
                .collect();
            if self.known_types.is_some()
                && words.len() == 2
                && words[0].eq_ignore_ascii_case("end")
                && words[1].eq_ignore_ascii_case("if")
            {
                return Err(self.failure(CompileErrorKind::EndIfWithoutBlock, *span));
            }
        }
        match statement {
            Statement::Assign { target, value, .. }
            | Statement::SetAssign { target, value, .. } => {
                if self.known_types.is_some() {
                    // Syntax is checked across this assignment before resolving
                    // either operand. Its value is compiled before its place.
                    self.check_expression_syntax(target, false)?;
                    self.check_expression_syntax(value, false)?;
                    self.expression(value)?;
                    self.check_expression_access(target, false, true)?;
                } else {
                    self.check_expression_access(target, false, true)?;
                    self.expression(value)?;
                }
                if matches!(statement, Statement::SetAssign { .. })
                    && matches!(value, Expr::Group { .. })
                    && matches!(self.infer(value), Type::Scalar | Type::Text)
                {
                    return Err(self.failure(CompileErrorKind::TypeMismatch, value.span()));
                }
                if self.reference_read_only(target) {
                    return Err(self.failure(CompileErrorKind::ReadOnlyProperty, target.span()));
                }
                if matches!(statement, Statement::Assign { .. }) {
                    if matches!(value, Expr::Literal(Literal::Nothing, _))
                        && !matches!(
                            self.infer(target),
                            Type::Object | Type::Reference(..) | Type::Variant | Type::Unknown
                        )
                    {
                        return Err(self.failure(CompileErrorKind::InvalidObjectUse, value.span()));
                    }
                    let whole = match target {
                        Expr::Index { target, args, .. } if args.is_empty() => target.as_ref(),
                        _ => target,
                    };
                    if let Expr::Ident(name, _) = whole {
                        if self.fixed_arrays.contains(&key(name)) {
                            return Err(
                                self.failure(CompileErrorKind::CannotAssignArray, target.span())
                            );
                        }
                    }
                    if let Type::Class(class) = self.infer(target) {
                        if let Some(setter) = self.default_procedure(&class, ProcKind::PropertyLet)
                        {
                            let count = setter.params.len().saturating_sub(1);
                            if setter.params[..count]
                                .iter()
                                .any(|p| !p.optional && p.mode != ParamMode::ParamArray)
                            {
                                return Err(self.failure(
                                    CompileErrorKind::ArgumentNotOptional,
                                    target.span(),
                                ));
                            }
                        }
                    }
                }
            }
            Statement::Call {
                target,
                explicit_call,
                ..
            } => {
                let mut receiver = target;
                loop {
                    receiver = match receiver {
                        Expr::Index { target, .. } => target,
                        Expr::Group { expression, .. } => expression,
                        Expr::Member { object, .. } => object,
                        _ => break,
                    };
                }
                // `With New C` is legal; `(New C).Method ...` is not a
                // valid standalone VBA call statement.
                if matches!(receiver, Expr::New { .. }) {
                    return Err(self.failure(CompileErrorKind::SyntaxError, target.span()));
                }
                self.check_expression(target, !explicit_call)?;
            }
            Statement::Dim(declaration) => {
                for variable in &declaration.items {
                    if variable.type_name.is_new
                        && self.document_modules.is_some_and(|documents| {
                            documents.contains_key(&key(&variable.type_name.name))
                        })
                    {
                        return Err(self.failure(CompileErrorKind::InvalidNewUse, declaration.span));
                    }
                    if self.reserved_name(&variable.name)
                        || (!declaration.is_const && variable.value.is_some())
                        || (declaration.is_const && variable.value.is_none())
                        || (variable.type_name.is_new
                            && (matches!(
                                self.declared(&variable.type_name),
                                Type::Scalar | Type::Text | Type::Variant
                            ) || variable.type_name.name.eq_ignore_ascii_case("object")))
                    {
                        return Err(self.failure(CompileErrorKind::SyntaxError, declaration.span));
                    }
                    if declaration.is_const
                        && matches!(&variable.value, Some(Expr::Literal(Literal::Nothing, _)))
                    {
                        let kind = match self.declared(&variable.type_name) {
                            Type::Variant if variable.explicit_type => {
                                CompileErrorKind::ObjectVariableNotSet
                            }
                            Type::Variant | Type::Object | Type::Class(_) | Type::Reference(..) => {
                                CompileErrorKind::InvalidConstantType
                            }
                            _ => CompileErrorKind::InvalidObjectUse,
                        };
                        return Err(self.failure(kind, declaration.span));
                    }
                    self.check_class(&variable.type_name.name, declaration.span)?;
                    self.validate_type(&variable.type_name, declaration.span)?;
                    if let Some(length) = &variable.type_name.fixed_length {
                        if matches!(
                            self.constant_value(length, &mut BTreeSet::new())
                                .map_err(|kind| self.failure(kind, length.span()))?,
                            Some(Value::LongLong(_))
                        ) {
                            return Err(self.failure(
                                CompileErrorKind::InvalidFixedStringLength,
                                length.span(),
                            ));
                        }
                        if self
                            .constant_integer(length)?
                            .is_some_and(|length| !(1..=65_526).contains(&length))
                        {
                            return Err(self.failure(
                                CompileErrorKind::InvalidFixedStringLength,
                                length.span(),
                            ));
                        }
                    }
                    if let Some(bounds) = &variable.array_bounds {
                        for bound in bounds {
                            if let Some(lower) = &bound.lower {
                                self.expression(lower)?;
                            }
                            self.expression(&bound.upper)?;
                            let lower = match &bound.lower {
                                Some(lower) => self.constant_integer(lower)?,
                                None => Some(
                                    self.module
                                        .items
                                        .iter()
                                        .find_map(|item| match item {
                                            ModuleItem::Option(ModuleOption::Base(base), _) => {
                                                Some(i64::from(*base))
                                            }
                                            _ => None,
                                        })
                                        .unwrap_or(0),
                                ),
                            };
                            let upper = self.constant_integer(&bound.upper)?;
                            if matches!((lower, upper), (Some(lower), Some(upper)) if upper < lower)
                            {
                                return Err(self.failure(
                                    CompileErrorKind::ValueOutOfRange,
                                    bound.upper.span(),
                                ));
                            }
                        }
                    }
                    if let Some(value) = &variable.value {
                        self.expression(value)?;
                        if declaration.is_const {
                            let folded = self
                                .constant_value_in(
                                    value,
                                    &mut BTreeSet::new(),
                                    None,
                                    ConstantContext::Initializer,
                                )
                                .map_err(|kind| self.failure(kind, value.span()))?;
                            if let Some(folded) = folded {
                                self.coerce_constant_in(
                                    folded,
                                    &variable.type_name,
                                    value.span(),
                                    None,
                                    &mut BTreeSet::new(),
                                )
                                .map_err(|kind| self.failure(kind, value.span()))?;
                            }
                        }
                    }
                }
                self.declaration(declaration);
            }
            Statement::ReDim { items, .. } => {
                for item in items {
                    if self.is_paramarray(&item.target) {
                        return Err(self
                            .failure(CompileErrorKind::InvalidParamArrayUse, item.target.span()));
                    }
                    self.expression(&item.target)?;
                    for bound in &item.bounds {
                        if let Some(lower) = &bound.lower {
                            self.expression(lower)?;
                        }
                        self.expression(&bound.upper)?;
                    }
                }
            }
            Statement::Erase { targets, .. } => {
                for target in targets {
                    if self.is_paramarray(target) {
                        return Err(
                            self.failure(CompileErrorKind::InvalidParamArrayUse, target.span())
                        );
                    }
                    self.expression(target)?;
                }
            }
            Statement::Open(open) => {
                self.expression(&open.path)?;
                self.expression(&open.file_number)?;
                if let Some(length) = &open.record_len {
                    self.expression(length)?;
                }
            }
            Statement::Close { files, .. } => {
                for file in files {
                    self.expression(file)?;
                }
            }
            Statement::FileOutput(output) => {
                self.expression(&output.file_number)?;
                for item in &output.items {
                    if let Some(value) = &item.value {
                        self.expression(value)?;
                    }
                }
            }
            Statement::FileInput(input) => {
                self.expression(&input.file_number)?;
                for target in &input.targets {
                    self.expression(target)?;
                }
            }
            Statement::FileTransfer(transfer) => {
                self.expression(&transfer.file_number)?;
                if let Some(record) = &transfer.record_number {
                    self.expression(record)?;
                }
                self.expression(&transfer.value)?;
            }
            Statement::FileSeek(seek) => {
                self.expression(&seek.file_number)?;
                self.expression(&seek.position)?;
            }
            Statement::FileSystem(
                FileSystemStmt::Rename {
                    source,
                    destination,
                    ..
                }
                | FileSystemStmt::Copy {
                    source,
                    destination,
                    ..
                },
            ) => {
                self.expression(source)?;
                self.expression(destination)?;
            }
            Statement::FileSystem(FileSystemStmt::Unary { path, .. }) => self.expression(path)?,
            Statement::FileSystem(FileSystemStmt::SetAttr {
                path, attributes, ..
            }) => {
                self.expression(path)?;
                self.expression(attributes)?;
            }
            Statement::FileRecordLock(lock) => {
                self.expression(&lock.file_number)?;
                if let Some(start) = &lock.start {
                    self.expression(start)?;
                }
                if let Some(end) = &lock.end {
                    self.expression(end)?;
                }
            }
            Statement::FileWidth {
                file_number, width, ..
            } => {
                self.expression(file_number)?;
                self.expression(width)?;
            }
            Statement::MidAssign(mid) => {
                self.expression(&mid.target)?;
                self.expression(&mid.start)?;
                if let Some(length) = &mid.length {
                    self.expression(length)?;
                }
                self.expression(&mid.value)?;
            }
            Statement::AlignedAssign(assign) => {
                self.expression(&assign.target)?;
                self.expression(&assign.value)?;
            }
            Statement::If(branch) => {
                self.expression(&branch.condition)?;
                self.body(&branch.then_body)?;
                for (condition, body) in &branch.else_ifs {
                    self.expression(condition)?;
                    self.body(body)?;
                }
                if let Some(body) = &branch.else_body {
                    self.body(body)?;
                }
            }
            Statement::SelectCase(select) => {
                self.expression(&select.subject)?;
                for case in &select.cases {
                    for label in &case.labels {
                        match label {
                            CaseLabel::Value(value) | CaseLabel::Compare(_, value) => {
                                self.expression(value)?
                            }
                            CaseLabel::Range(start, end) => {
                                self.expression(start)?;
                                self.expression(end)?;
                            }
                        }
                    }
                    self.body(&case.body)?;
                }
                if let Some(body) = &select.case_else {
                    self.body(body)?;
                }
            }
            Statement::For(loop_) => {
                self.expression(&loop_.counter)?;
                self.expression(&loop_.from)?;
                self.expression(&loop_.to)?;
                if let Some(step) = &loop_.step {
                    self.expression(step)?;
                }
                self.body(&loop_.body)?;
            }
            Statement::ForEach(loop_) => {
                self.expression(&loop_.item)?;
                self.expression(&loop_.collection)?;
                self.body(&loop_.body)?;
            }
            Statement::Do(loop_) => {
                if let Some(test) = &loop_.pre {
                    self.expression(&test.condition)?;
                }
                self.body(&loop_.body)?;
                if let Some(test) = &loop_.post {
                    self.expression(&test.condition)?;
                }
            }
            Statement::While {
                condition, body, ..
            } => {
                self.expression(condition)?;
                self.body(body)?;
            }
            Statement::With { subject, body, .. } => {
                self.expression(subject)?;
                self.with_types.push(self.infer(subject));
                let answer = self.body(body);
                self.with_types.pop();
                answer?;
            }
            Statement::OnBranch(branch) => self.expression(&branch.selector)?,
            Statement::RaiseEvent(event) => {
                for arg in &event.args {
                    if let Some(value) = &arg.value {
                        self.expression(value)?;
                    }
                }
            }
            Statement::OnError(_)
            | Statement::Resume { .. }
            | Statement::GoTo { .. }
            | Statement::GoSub { .. }
            | Statement::Return { .. }
            | Statement::Exit { .. }
            | Statement::Label { .. }
            | Statement::LineNumber { .. }
            | Statement::End { .. }
            | Statement::Stop { .. }
            | Statement::Comment { .. }
            | Statement::Directive { .. }
            | Statement::Unknown { .. }
            | Statement::FileReset { .. } => {}
        }
        Ok(())
    }
}

/// Collect each statement's expression roots without merging the syntax and
/// semantic phases. Their lexer ranges decide which roots share a logical line.
fn collect_syntax_roots<'e>(body: &'e [Statement], roots: &mut Vec<(&'e Expr, bool)>) {
    macro_rules! expr {
        ($value:expr) => {
            roots.push(($value, false))
        };
    }
    for statement in body {
        match statement {
            Statement::Assign { target, value, .. }
            | Statement::SetAssign { target, value, .. } => {
                expr!(target);
                expr!(value);
            }
            Statement::Call {
                target,
                explicit_call,
                ..
            } => roots.push((target, !explicit_call)),
            Statement::Dim(declaration) => {
                for variable in &declaration.items {
                    if let Some(length) = &variable.type_name.fixed_length {
                        expr!(length);
                    }
                    if let Some(bounds) = &variable.array_bounds {
                        for bound in bounds {
                            if let Some(lower) = &bound.lower {
                                expr!(lower);
                            }
                            expr!(&bound.upper);
                        }
                    }
                    if let Some(value) = &variable.value {
                        expr!(value);
                    }
                }
            }
            Statement::ReDim { items, .. } => {
                for item in items {
                    expr!(&item.target);
                    for bound in &item.bounds {
                        if let Some(lower) = &bound.lower {
                            expr!(lower);
                        }
                        expr!(&bound.upper);
                    }
                }
            }
            Statement::Erase { targets, .. } => {
                for target in targets {
                    expr!(target);
                }
            }
            Statement::Open(s) => {
                expr!(&s.path);
                expr!(&s.file_number);
                if let Some(length) = &s.record_len {
                    expr!(length);
                }
            }
            Statement::Close { files, .. } => {
                for file in files {
                    expr!(file);
                }
            }
            Statement::FileOutput(s) => {
                expr!(&s.file_number);
                for item in &s.items {
                    if let Some(value) = &item.value {
                        expr!(value);
                    }
                }
            }
            Statement::FileInput(s) => {
                expr!(&s.file_number);
                for target in &s.targets {
                    expr!(target);
                }
            }
            Statement::FileTransfer(s) => {
                expr!(&s.file_number);
                if let Some(record) = &s.record_number {
                    expr!(record);
                }
                expr!(&s.value);
            }
            Statement::FileSeek(s) => {
                expr!(&s.file_number);
                expr!(&s.position);
            }
            Statement::FileSystem(
                FileSystemStmt::Rename {
                    source,
                    destination,
                    ..
                }
                | FileSystemStmt::Copy {
                    source,
                    destination,
                    ..
                },
            ) => {
                expr!(source);
                expr!(destination);
            }
            Statement::FileSystem(FileSystemStmt::Unary { path, .. }) => expr!(path),
            Statement::FileSystem(FileSystemStmt::SetAttr {
                path, attributes, ..
            }) => {
                expr!(path);
                expr!(attributes);
            }
            Statement::FileRecordLock(s) => {
                expr!(&s.file_number);
                if let Some(start) = &s.start {
                    expr!(start);
                }
                if let Some(end) = &s.end {
                    expr!(end);
                }
            }
            Statement::FileWidth {
                file_number, width, ..
            } => {
                expr!(file_number);
                expr!(width);
            }
            Statement::MidAssign(s) => {
                expr!(&s.target);
                expr!(&s.start);
                if let Some(length) = &s.length {
                    expr!(length);
                }
                expr!(&s.value);
            }
            Statement::AlignedAssign(s) => {
                expr!(&s.target);
                expr!(&s.value);
            }
            Statement::If(s) => {
                expr!(&s.condition);
                collect_syntax_roots(&s.then_body, roots);
                for (condition, body) in &s.else_ifs {
                    expr!(condition);
                    collect_syntax_roots(body, roots);
                }
                if let Some(body) = &s.else_body {
                    collect_syntax_roots(body, roots);
                }
            }
            Statement::SelectCase(s) => {
                expr!(&s.subject);
                for case in &s.cases {
                    for label in &case.labels {
                        match label {
                            CaseLabel::Value(value) | CaseLabel::Compare(_, value) => expr!(value),
                            CaseLabel::Range(start, end) => {
                                expr!(start);
                                expr!(end);
                            }
                        }
                    }
                    collect_syntax_roots(&case.body, roots);
                }
                if let Some(body) = &s.case_else {
                    collect_syntax_roots(body, roots);
                }
            }
            Statement::For(s) => {
                expr!(&s.counter);
                expr!(&s.from);
                expr!(&s.to);
                if let Some(step) = &s.step {
                    expr!(step);
                }
                collect_syntax_roots(&s.body, roots);
                if let Some(counter) = &s.next_counter {
                    expr!(counter);
                }
            }
            Statement::ForEach(s) => {
                expr!(&s.item);
                expr!(&s.collection);
                collect_syntax_roots(&s.body, roots);
                if let Some(counter) = &s.next_counter {
                    expr!(counter);
                }
            }
            Statement::Do(s) => {
                if let Some(test) = &s.pre {
                    expr!(&test.condition);
                }
                collect_syntax_roots(&s.body, roots);
                if let Some(test) = &s.post {
                    expr!(&test.condition);
                }
            }
            Statement::While {
                condition, body, ..
            } => {
                expr!(condition);
                collect_syntax_roots(body, roots);
            }
            Statement::With { subject, body, .. } => {
                expr!(subject);
                collect_syntax_roots(body, roots);
            }
            Statement::OnBranch(s) => expr!(&s.selector),
            Statement::RaiseEvent(s) => {
                for arg in &s.args {
                    if let Some(value) = &arg.value {
                        expr!(value);
                    }
                }
            }
            Statement::OnError(_)
            | Statement::Resume { .. }
            | Statement::GoTo { .. }
            | Statement::GoSub { .. }
            | Statement::Return { .. }
            | Statement::Exit { .. }
            | Statement::Label { .. }
            | Statement::LineNumber { .. }
            | Statement::End { .. }
            | Statement::Stop { .. }
            | Statement::Comment { .. }
            | Statement::Directive { .. }
            | Statement::Unknown { .. }
            | Statement::FileReset { .. } => {}
        }
    }
}

/// After the statement's syntax pass, compile a call's arguments and receiver
/// before resolving its member and binding its arguments. Other expression
/// operands retain their source order. Compile-on-demand uses its own traversal.
fn compilation_expression_order<'e>(expression: &'e Expr, ordered: &mut Vec<&'e Expr>) {
    if let Expr::Index { target, args, .. } = expression {
        for arg in args {
            if let Some(value) = &arg.value {
                compilation_expression_order(value, ordered);
            }
        }
        compilation_expression_order(target, ordered);
        ordered.push(expression);
        return;
    }
    if let Expr::Member { object, .. } | Expr::Bang { object, .. } = expression {
        compilation_expression_order(object, ordered);
        ordered.push(expression);
        return;
    }
    ordered.push(expression);
    match expression {
        Expr::Group { expression, .. } => compilation_expression_order(expression, ordered),
        Expr::Unary { operand, .. } | Expr::TypeOf { operand, .. } => {
            compilation_expression_order(operand, ordered)
        }
        Expr::Binary { lhs, rhs, .. } => {
            compilation_expression_order(lhs, ordered);
            compilation_expression_order(rhs, ordered);
        }
        Expr::Literal(..)
        | Expr::EvaluateShortcut { .. }
        | Expr::Ident(..)
        | Expr::TypedIdent { .. }
        | Expr::WithMember(..)
        | Expr::WithBangMember(..)
        | Expr::New { .. }
        | Expr::AddressOf { .. } => {}
        Expr::Index { .. } | Expr::Member { .. } | Expr::Bang { .. } => {
            unreachable!("invocation and receiver handled above")
        }
    }
}

fn inherited_object_member(name: &str) -> bool {
    matches!(
        key(name).as_str(),
        "queryinterface"
            | "addref"
            | "release"
            | "gettypeinfocount"
            | "gettypeinfo"
            | "getidsofnames"
            | "invoke"
    )
}

fn key(name: &str) -> String {
    name.to_ascii_lowercase()
}

fn bare_name(expr: &Expr) -> Option<&str> {
    match expr {
        Expr::Ident(name, _) | Expr::TypedIdent { name, .. } => Some(name),
        _ => None,
    }
}

fn collect_locals(body: &[Statement], names: &mut BTreeSet<String>) {
    for statement in body {
        match statement {
            Statement::Dim(declaration) => {
                names.extend(declaration.items.iter().map(|item| key(&item.name)));
            }
            Statement::If(branch) => {
                collect_locals(&branch.then_body, names);
                for (_, body) in &branch.else_ifs {
                    collect_locals(body, names);
                }
                if let Some(body) = &branch.else_body {
                    collect_locals(body, names);
                }
            }
            Statement::SelectCase(select) => {
                for case in &select.cases {
                    collect_locals(&case.body, names);
                }
                if let Some(body) = &select.case_else {
                    collect_locals(body, names);
                }
            }
            Statement::For(loop_) => collect_locals(&loop_.body, names),
            Statement::ForEach(loop_) => collect_locals(&loop_.body, names),
            Statement::Do(loop_) => collect_locals(&loop_.body, names),
            Statement::While { body, .. } | Statement::With { body, .. } => {
                collect_locals(body, names)
            }
            _ => {}
        }
    }
}
