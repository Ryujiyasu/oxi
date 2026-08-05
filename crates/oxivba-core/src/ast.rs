// This Source Code Form is subject to the terms of the Mozilla Public
// License, v. 2.0. If a copy of the MPL was not distributed with this
// file, You can obtain one at https://mozilla.org/MPL/2.0/.

//! VBA syntax tree.
//!
//! Shaped for analysis rather than execution: every node keeps its source span,
//! and constructs that are equivalent at runtime but different to a reader —
//! `Call Foo(x)` versus `Foo x`, a block `If` versus a one-line `If` — stay
//! distinguishable, because a diagnostic that rewrites them has to put them back
//! the way it found them.

use crate::lexer::Span;

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum Visibility {
    /// No keyword. Module-level default is private, procedure default is public.
    Default,
    Private,
    Public,
    Friend,
    /// `Global`, the pre-VB6 spelling of `Public`.
    Global,
}

/// A declared type, kept as written. `Variant` when omitted.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct TypeName {
    pub name: String,
    /// A type suffix such as the `$` in `Dim s$`.
    pub suffix: Option<char>,
}

impl TypeName {
    pub fn implicit() -> TypeName {
        TypeName {
            name: "Variant".to_string(),
            suffix: None,
        }
    }
}

#[derive(Debug, Clone, PartialEq)]
pub struct Module {
    pub items: Vec<ModuleItem>,
}

/// Kept as a flat ordered list: the order of declarations is part of what a
/// reader sees, and a diff that reorders them is not a faithful diff.
#[derive(Debug, Clone, PartialEq)]
pub enum ModuleItem {
    /// `Attribute VB_Name = "Module1"`, emitted by the VBE on export.
    Attribute {
        name: String,
        value: String,
        span: Span,
    },
    Option(ModuleOption, Span),
    /// `DefInt A-C, I` and the other module-level implicit-type directives.
    DefType(DefTypeDecl),
    /// `Dim` / `Private` / `Public` / `Const` at module level.
    Variables(VarDecl),
    Type(TypeDef),
    Enum(EnumDef),
    /// `Declare [PtrSafe] Sub|Function ... Lib "..."`.
    ExternalProc(ExternalProc),
    Implements {
        interface: String,
        span: Span,
    },
    Event {
        name: String,
        params: Vec<Param>,
        span: Span,
    },
    Procedure(Procedure),
    /// A line the parser did not understand, kept verbatim so that nothing is
    /// silently dropped and the count of unparsed input can be reported.
    Unknown {
        text: String,
        span: Span,
    },
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub enum ModuleOption {
    Explicit,
    Base(u32),
    /// `Binary` / `Text` / `Database`; affects string comparison and `Like`.
    Compare(String),
    PrivateModule,
}

#[derive(Debug, Clone, PartialEq)]
pub struct VarDecl {
    pub visibility: Visibility,
    pub is_const: bool,
    pub is_static: bool,
    /// `Dim a, b` declares two variables in one statement.
    pub items: Vec<VarItem>,
    pub span: Span,
}

#[derive(Debug, Clone, PartialEq)]
pub struct VarItem {
    pub name: String,
    /// `Dim a(1 To 10)`. Empty bounds mean a dynamic array: `Dim a()`.
    pub array_bounds: Option<Vec<ArrayBound>>,
    pub type_name: TypeName,
    /// `WithEvents obj As Worksheet`.
    pub with_events: bool,
    /// Only for `Const`.
    pub value: Option<Expr>,
}

#[derive(Debug, Clone, PartialEq)]
pub struct ArrayBound {
    pub lower: Option<Expr>,
    pub upper: Expr,
}

#[derive(Debug, Clone, PartialEq)]
pub struct TypeDef {
    pub visibility: Visibility,
    pub name: String,
    pub fields: Vec<VarItem>,
    pub span: Span,
}

#[derive(Debug, Clone, PartialEq)]
pub struct EnumDef {
    pub visibility: Visibility,
    pub name: String,
    pub members: Vec<(String, Option<Expr>)>,
    pub span: Span,
}

#[derive(Debug, Clone, PartialEq)]
pub struct ExternalProc {
    pub visibility: Visibility,
    pub is_function: bool,
    pub ptr_safe: bool,
    pub name: String,
    pub lib: String,
    pub alias: Option<String>,
    pub params: Vec<Param>,
    pub return_type: Option<TypeName>,
    pub span: Span,
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum ProcKind {
    Sub,
    Function,
    PropertyGet,
    PropertyLet,
    PropertySet,
}

#[derive(Debug, Clone, PartialEq)]
pub struct Procedure {
    pub kind: ProcKind,
    pub visibility: Visibility,
    pub is_static: bool,
    pub name: String,
    pub params: Vec<Param>,
    pub return_type: Option<TypeName>,
    pub body: Vec<Statement>,
    pub span: Span,
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum ParamMode {
    /// The default when nothing is written, which is the usual source of
    /// surprise: a callee can reassign the caller's variable.
    ByRef,
    ByVal,
    ParamArray,
}

#[derive(Debug, Clone, PartialEq)]
pub struct Param {
    pub mode: ParamMode,
    pub optional: bool,
    pub name: String,
    pub is_array: bool,
    pub type_name: TypeName,
    pub default: Option<Expr>,
}

#[derive(Debug, Clone, PartialEq)]
pub enum Statement {
    /// `x = expr`, or `Let x = expr`.
    Assign {
        target: Expr,
        value: Expr,
        span: Span,
    },
    /// `Set x = expr`. Distinct from `Assign`: it binds a reference.
    SetAssign {
        target: Expr,
        value: Expr,
        span: Span,
    },
    /// `Foo a, b` or `Call Foo(a, b)`. `explicit_call` records which was used.
    Call {
        target: Expr,
        explicit_call: bool,
        span: Span,
    },
    Dim(VarDecl),
    ReDim {
        preserve: bool,
        items: Vec<VarItem>,
        span: Span,
    },
    Erase {
        targets: Vec<Expr>,
        span: Span,
    },
    /// `Open path For Input As #1` and its other mode/access/lock variants.
    Open(FileOpenStmt),
    /// `Close`, optionally followed by one or more file numbers.
    Close {
        files: Vec<Expr>,
        span: Span,
    },
    /// `Print #n, ...` or `Write #n, ...`.
    FileOutput(FileOutputStmt),
    /// `Input #n, ...` or `Line Input #n, ...`.
    FileInput(FileInputStmt),
    /// `Get #n, record, value` or `Put #n, record, value`.
    FileTransfer(FileTransferStmt),
    /// `Seek #n, position` (distinct from the `Seek(n)` function).
    FileSeek(FileSeekStmt),
    /// Native filesystem mutation and directory-navigation statements.
    FileSystem(FileSystemStmt),
    /// `Lock` and `Unlock` over an optional record/byte range.
    FileRecordLock(FileRecordLockStmt),
    /// `Width #n, columns` for sequential output files.
    FileWidth {
        file_number: Expr,
        width: Expr,
        span: Span,
    },
    /// `Mid$(target, start[, length]) = replacement`.
    MidAssign(MidAssignStmt),
    /// `LSet target = value` or `RSet target = value`.
    AlignedAssign(AlignedAssignStmt),
    /// `Reset`, which closes every open disk file.
    FileReset {
        span: Span,
    },
    If(IfStmt),
    SelectCase(SelectCaseStmt),
    For(ForStmt),
    ForEach(ForEachStmt),
    Do(DoStmt),
    /// `While ... Wend`, the older spelling.
    While {
        condition: Expr,
        body: Vec<Statement>,
        span: Span,
    },
    With {
        subject: Expr,
        body: Vec<Statement>,
        span: Span,
    },
    OnError(OnError),
    /// Computed `On expr GoTo/GoSub label, ...` dispatch.
    OnBranch(OnBranchStmt),
    /// `RaiseEvent Name(args)` in a class module.
    RaiseEvent(RaiseEventStmt),
    Resume {
        target: ResumeTarget,
        span: Span,
    },
    GoTo {
        label: String,
        span: Span,
    },
    GoSub {
        label: String,
        span: Span,
    },
    Return {
        span: Span,
    },
    Exit {
        what: ExitKind,
        span: Span,
    },
    Label {
        name: String,
        span: Span,
    },
    /// A bare line number acting as a label.
    LineNumber {
        value: u32,
        span: Span,
    },
    /// `End` on its own: terminates execution.
    End {
        span: Span,
    },
    Stop {
        span: Span,
    },
    Comment {
        text: String,
        span: Span,
    },
    /// `#If` and friends, kept verbatim rather than evaluated.
    Directive {
        text: String,
        span: Span,
    },
    /// Anything the parser could not classify, preserved verbatim.
    Unknown {
        text: String,
        span: Span,
    },
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum FileMode {
    Append,
    Binary,
    Input,
    Output,
    Random,
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum FileAccess {
    Read,
    Write,
    ReadWrite,
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum FileLock {
    Shared,
    Read,
    Write,
    ReadWrite,
}

#[derive(Debug, Clone, PartialEq)]
pub struct FileOpenStmt {
    pub path: Expr,
    pub mode: FileMode,
    pub access: Option<FileAccess>,
    pub lock: Option<FileLock>,
    pub file_number: Expr,
    /// `Len = n`, used by random-access files.
    pub record_len: Option<Expr>,
    pub span: Span,
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum FileOutputKind {
    Print,
    Write,
}

#[derive(Debug, Clone, PartialEq)]
pub struct FileOutputStmt {
    pub kind: FileOutputKind,
    pub file_number: Expr,
    /// The separator is semantically visible for `Print #`: comma advances to
    /// a print zone while semicolon continues at the current position.
    pub items: Vec<FileOutputItem>,
    pub span: Span,
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum FileOutputSeparator {
    Comma,
    Semicolon,
}

#[derive(Debug, Clone, PartialEq)]
pub struct FileOutputItem {
    pub separator: FileOutputSeparator,
    /// `None` preserves a trailing or otherwise omitted output field.
    pub value: Option<Expr>,
}

#[derive(Debug, Clone, PartialEq)]
pub struct FileInputStmt {
    /// `true` for `Line Input`, whose destination receives the whole line.
    pub line: bool,
    pub file_number: Expr,
    pub targets: Vec<Expr>,
    pub span: Span,
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct DefTypeDecl {
    /// Canonical VBA type name (`Integer`, `String`, `Object`, ...).
    pub type_name: String,
    pub ranges: Vec<LetterRange>,
    pub span: Span,
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct LetterRange {
    pub start: char,
    pub end: char,
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum FileTransferKind {
    Get,
    Put,
}

#[derive(Debug, Clone, PartialEq)]
pub struct FileTransferStmt {
    pub kind: FileTransferKind,
    pub file_number: Expr,
    /// Omitted in `Get #n, , value` and `Put #n, , value`.
    pub record_number: Option<Expr>,
    pub value: Expr,
    pub span: Span,
}

#[derive(Debug, Clone, PartialEq)]
pub struct FileSeekStmt {
    pub file_number: Expr,
    pub position: Expr,
    pub span: Span,
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum FileSystemUnaryKind {
    Kill,
    MkDir,
    RmDir,
    ChDir,
    ChDrive,
}

#[derive(Debug, Clone, PartialEq)]
pub enum FileSystemStmt {
    Rename {
        source: Expr,
        destination: Expr,
        span: Span,
    },
    Copy {
        source: Expr,
        destination: Expr,
        span: Span,
    },
    Unary {
        kind: FileSystemUnaryKind,
        path: Expr,
        span: Span,
    },
    SetAttr {
        path: Expr,
        attributes: Expr,
        span: Span,
    },
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum FileRecordLockKind {
    Lock,
    Unlock,
}

#[derive(Debug, Clone, PartialEq)]
pub struct FileRecordLockStmt {
    pub kind: FileRecordLockKind,
    pub file_number: Expr,
    /// A single record uses `start` with `end = None`; `start To end` uses both.
    pub start: Option<Expr>,
    pub end: Option<Expr>,
    pub span: Span,
}

#[derive(Debug, Clone, PartialEq)]
pub struct IfStmt {
    pub condition: Expr,
    pub then_body: Vec<Statement>,
    pub else_ifs: Vec<(Expr, Vec<Statement>)>,
    pub else_body: Option<Vec<Statement>>,
    /// A one-line `If x Then y` has no `End If` and cannot be extended.
    pub single_line: bool,
    pub span: Span,
}

#[derive(Debug, Clone, PartialEq)]
pub struct SelectCaseStmt {
    pub subject: Expr,
    pub cases: Vec<CaseClause>,
    pub case_else: Option<Vec<Statement>>,
    pub span: Span,
}

#[derive(Debug, Clone, PartialEq)]
pub struct CaseClause {
    pub labels: Vec<CaseLabel>,
    pub body: Vec<Statement>,
}

#[derive(Debug, Clone, PartialEq)]
pub enum CaseLabel {
    Value(Expr),
    /// `Case 1 To 10`.
    Range(Expr, Expr),
    /// `Case Is >= 10`.
    Compare(BinaryOp, Expr),
}

#[derive(Debug, Clone, PartialEq)]
pub struct ForStmt {
    pub counter: Expr,
    pub from: Expr,
    pub to: Expr,
    pub step: Option<Expr>,
    pub body: Vec<Statement>,
    pub span: Span,
}

#[derive(Debug, Clone, PartialEq)]
pub struct ForEachStmt {
    pub item: Expr,
    pub collection: Expr,
    pub body: Vec<Statement>,
    pub span: Span,
}

#[derive(Debug, Clone, PartialEq)]
pub struct DoStmt {
    /// `None` for a bare `Do ... Loop`.
    pub pre: Option<LoopTest>,
    pub post: Option<LoopTest>,
    pub body: Vec<Statement>,
    pub span: Span,
}

#[derive(Debug, Clone, PartialEq)]
pub struct LoopTest {
    pub until: bool,
    pub condition: Expr,
}

#[derive(Debug, Clone, PartialEq)]
pub enum OnError {
    /// `On Error GoTo <label>`.
    Goto { label: String, span: Span },
    /// `On Error GoTo 0`: stop handling.
    Disable { span: Span },
    /// `On Error Resume Next`: swallow everything. Worth finding.
    ResumeNext { span: Span },
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum OnBranchKind {
    GoTo,
    GoSub,
}

#[derive(Debug, Clone, PartialEq)]
pub struct OnBranchStmt {
    pub selector: Expr,
    pub kind: OnBranchKind,
    /// Choice 1 selects index 0; choice 0 or greater than the list falls through.
    pub labels: Vec<String>,
    pub span: Span,
}

#[derive(Debug, Clone, PartialEq)]
pub struct MidAssignStmt {
    pub target: Expr,
    pub start: Expr,
    pub length: Option<Expr>,
    pub value: Expr,
    pub span: Span,
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum AlignmentKind {
    Left,
    Right,
}

#[derive(Debug, Clone, PartialEq)]
pub struct AlignedAssignStmt {
    pub kind: AlignmentKind,
    pub target: Expr,
    pub value: Expr,
    pub span: Span,
}

#[derive(Debug, Clone, PartialEq)]
pub struct RaiseEventStmt {
    pub name: String,
    pub args: Vec<Argument>,
    pub span: Span,
}

#[derive(Debug, Clone, PartialEq)]
pub enum ResumeTarget {
    Same,
    Next,
    Label(String),
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum ExitKind {
    Sub,
    Function,
    Property,
    For,
    Do,
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum UnaryOp {
    Neg,
    Plus,
    Not,
}

/// Binary operators, grouped by VBA's precedence levels.
///
/// The two that catch people out:
/// - `\` (integer division) and `Mod` sit on their own levels *between* `*`/`/`
///   and `+`/`-`, so `a + b \ c` is `a + (b \ c)`.
/// - `Not` binds *looser* than comparison, so `Not a = b` is `Not (a = b)`.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum BinaryOp {
    Pow,
    Mul,
    Div,
    IntDiv,
    Mod,
    Add,
    Sub,
    Concat,
    Eq,
    Ne,
    Lt,
    Le,
    Gt,
    Ge,
    Is,
    Like,
    And,
    Or,
    Xor,
    Eqv,
    Imp,
}

#[derive(Debug, Clone, PartialEq)]
pub enum Literal {
    Number(f64),
    Str(String),
    /// Kept as written; interpretation needs a locale.
    Date(String),
    Bool(bool),
    /// `Empty`, `Null`, `Nothing`.
    Empty,
    Null,
    Nothing,
}

#[derive(Debug, Clone, PartialEq)]
pub enum Expr {
    Literal(Literal, Span),
    /// A bare name. Whether it is a variable, a call, or a property is not
    /// decidable without type information, so it is left as a name.
    Ident(String, Span),
    /// `.Value` inside a `With` block.
    WithMember(String, Span),
    /// `a.b`
    Member {
        object: Box<Expr>,
        name: String,
        span: Span,
    },
    /// `a(1)` — indexing and calling are the same syntax in VBA.
    Index {
        target: Box<Expr>,
        args: Vec<Argument>,
        span: Span,
    },
    /// `a!b`, the dictionary-access operator.
    Bang {
        object: Box<Expr>,
        name: String,
        span: Span,
    },
    New {
        type_name: String,
        span: Span,
    },
    Unary {
        op: UnaryOp,
        operand: Box<Expr>,
        span: Span,
    },
    Binary {
        op: BinaryOp,
        lhs: Box<Expr>,
        rhs: Box<Expr>,
        span: Span,
    },
    /// `TypeOf x Is Worksheet`.
    TypeOf {
        operand: Box<Expr>,
        type_name: String,
        span: Span,
    },
}

/// An argument, which may be named, and may be omitted entirely.
#[derive(Debug, Clone, PartialEq)]
pub struct Argument {
    pub name: Option<String>,
    /// `None` for an omitted positional argument: `Foo a, , c`.
    pub value: Option<Expr>,
}

impl Expr {
    pub fn span(&self) -> Span {
        match self {
            Expr::Literal(_, s)
            | Expr::Ident(_, s)
            | Expr::WithMember(_, s)
            | Expr::Member { span: s, .. }
            | Expr::Index { span: s, .. }
            | Expr::Bang { span: s, .. }
            | Expr::New { span: s, .. }
            | Expr::Unary { span: s, .. }
            | Expr::Binary { span: s, .. }
            | Expr::TypeOf { span: s, .. } => *s,
        }
    }

    /// Walk the expression, parents before children.
    pub fn visit(&self, f: &mut impl FnMut(&Expr)) {
        f(self);
        match self {
            Expr::Member { object, .. } | Expr::Bang { object, .. } => object.visit(f),
            Expr::Index { target, args, .. } => {
                target.visit(f);
                for arg in args {
                    if let Some(value) = &arg.value {
                        value.visit(f);
                    }
                }
            }
            Expr::Unary { operand, .. } | Expr::TypeOf { operand, .. } => operand.visit(f),
            Expr::Binary { lhs, rhs, .. } => {
                lhs.visit(f);
                rhs.visit(f);
            }
            Expr::Literal(..) | Expr::Ident(..) | Expr::WithMember(..) | Expr::New { .. } => {}
        }
    }

    /// The dotted name this expression reads, if it is a plain member chain.
    ///
    /// `Application.WorksheetFunction.Sum` yields
    /// `"Application.WorksheetFunction.Sum"`. Used by the diagnostics to spot
    /// API use without needing to resolve types.
    pub fn dotted_name(&self) -> Option<String> {
        match self {
            Expr::Ident(name, _) => Some(name.clone()),
            Expr::Member { object, name, .. } => {
                Some(format!("{}.{}", object.dotted_name()?, name))
            }
            Expr::Index { target, .. } => target.dotted_name(),
            _ => None,
        }
    }
}
