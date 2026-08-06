// This Source Code Form is subject to the terms of the Mozilla Public
// License, v. 2.0. If a copy of the MPL was not distributed with this
// file, You can obtain one at https://mozilla.org/MPL/2.0/.

//! Static diagnostics over a parsed module.
//!
//! Everything here runs on the syntax tree alone: no Excel, no COM, no Windows,
//! and no interpreter. That is the whole claim being made: the useful part
//! of "can this macro be migrated?" is answerable from source text.
//!
//! # The classification rules are data, not code
//!
//! [`RULES`] is a table. It is meant to be read by someone deciding whether they
//! agree with a verdict, and edited when they do not. A rule buried in a
//! `match` arm is a rule nobody can audit, and the verdict is the product.

use std::collections::{BTreeMap, BTreeSet};

use crate::ast::*;

/// How migratable a module looks.
///
/// A module is classified by its *worst* finding: one `Shell` call is enough to
/// keep the whole thing off a browser. "Worst" is [`Class::severity`], not the
/// declaration order. `A` and `B` are categories rather than grades, and a
/// macro that both transforms data and formats it is a report generator.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum Class {
    /// Report generation: formatting, printing, laying out a sheet.
    A,
    /// Data transformation: reading, computing, writing values.
    B,
    /// Reaches outside Excel: files, databases, other applications, the shell.
    C,
    /// Has a user interface of its own. Not a migration; a rewrite.
    D,
}

impl Class {
    /// How far from migratable, used to pick a winner when several rules match.
    ///
    /// `B` is below `A` deliberately: reading and writing cells is what a report
    /// generator does on its way to formatting, so formatting is the more
    /// specific description of the two.
    pub fn severity(self) -> u8 {
        match self {
            Class::B => 0,
            Class::A => 1,
            Class::C => 2,
            Class::D => 3,
        }
    }

    pub fn as_str(self) -> &'static str {
        match self {
            Class::A => "A",
            Class::B => "B",
            Class::C => "C",
            Class::D => "D",
        }
    }

    pub fn description(self) -> &'static str {
        match self {
            Class::A => "report generation",
            Class::B => "data transformation",
            Class::C => "out of scope: reaches outside Excel",
            Class::D => "out of scope: has its own user interface",
        }
    }
}

/// What a rule looks for.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum Match {
    /// The dotted name equals this, case-insensitively.
    Exact,
    /// The first segment equals this, so `ADODB.Connection` matches `ADODB`.
    Root,
    /// The first segment *starts with* this. Needed for the names the VBE
    /// generates by numbering: `UserForm1`, `UserForm2`.
    RootPrefix,
    /// Any segment equals this, so that `ws.Range.Font` matches a rule on
    /// `Font`.
    Segment,
}

pub struct Rule {
    pub pattern: &'static str,
    pub how: Match,
    pub class: Class,
    /// Shown to whoever reads the verdict. This is the justification, and it has
    /// to stand on its own.
    pub reason: &'static str,
}

/// The classification table.
///
/// Grouped so that it reads as a checklist. Position carries no weight: the
/// winner among several matches is chosen by [`Class::severity`].
pub const RULES: &[Rule] = &[
    // -- D: has its own UI ------------------------------------------------
    Rule {
        pattern: "UserForm",
        how: Match::RootPrefix,
        class: Class::D,
        reason: "drives a UserForm; the target platform supplies its own UI",
    },
    Rule {
        pattern: "MSForms",
        how: Match::Root,
        class: Class::D,
        reason: "uses MSForms controls",
    },
    Rule {
        pattern: "Load",
        how: Match::Exact,
        class: Class::D,
        reason: "loads a form",
    },
    Rule {
        pattern: "Unload",
        how: Match::Exact,
        class: Class::D,
        reason: "unloads a form",
    },
    // -- C: leaves Excel ---------------------------------------------------
    Rule {
        pattern: "Shell",
        how: Match::Exact,
        class: Class::C,
        reason: "starts an external process",
    },
    Rule {
        pattern: "CreateObject",
        how: Match::Exact,
        class: Class::C,
        reason: "late-binds an external COM object; the target cannot be determined statically",
    },
    Rule {
        pattern: "GetObject",
        how: Match::Exact,
        class: Class::C,
        reason: "attaches to an external COM object",
    },
    Rule {
        pattern: "Scripting.FileSystemObject",
        how: Match::Exact,
        class: Class::C,
        reason: "manipulates the file system",
    },
    Rule {
        pattern: "ADODB",
        how: Match::Root,
        class: Class::C,
        reason: "connects to a database",
    },
    Rule {
        pattern: "DAO",
        how: Match::Root,
        class: Class::C,
        reason: "connects to a database",
    },
    Rule {
        pattern: "Outlook",
        how: Match::Root,
        class: Class::C,
        reason: "drives another Office application",
    },
    Rule {
        pattern: "Word",
        how: Match::Root,
        class: Class::C,
        reason: "drives another Office application",
    },
    Rule {
        pattern: "PowerPoint",
        how: Match::Root,
        class: Class::C,
        reason: "drives another Office application",
    },
    Rule {
        pattern: "WScript",
        how: Match::Root,
        class: Class::C,
        reason: "uses Windows Script Host",
    },
    Rule {
        pattern: "SendKeys",
        how: Match::Exact,
        class: Class::C,
        reason: "synthesises keystrokes",
    },
    Rule {
        pattern: "Kill",
        how: Match::Exact,
        class: Class::C,
        reason: "deletes a file",
    },
    Rule {
        pattern: "MkDir",
        how: Match::Exact,
        class: Class::C,
        reason: "creates a directory",
    },
    Rule {
        pattern: "RmDir",
        how: Match::Exact,
        class: Class::C,
        reason: "removes a directory",
    },
    Rule {
        pattern: "FileCopy",
        how: Match::Exact,
        class: Class::C,
        reason: "copies a file",
    },
    Rule {
        pattern: "ChDir",
        how: Match::Exact,
        class: Class::C,
        reason: "changes the process working directory",
    },
    Rule {
        pattern: "ChDrive",
        how: Match::Exact,
        class: Class::C,
        reason: "changes the process working drive",
    },
    Rule {
        pattern: "SetAttr",
        how: Match::Exact,
        class: Class::C,
        reason: "changes file attributes",
    },
    Rule {
        pattern: "GetAttr",
        how: Match::Exact,
        class: Class::C,
        reason: "reads file attributes",
    },
    Rule {
        pattern: "CurDir",
        how: Match::Exact,
        class: Class::C,
        reason: "reads the process working directory",
    },
    Rule {
        pattern: "Dir",
        how: Match::Exact,
        class: Class::C,
        reason: "enumerates the file system",
    },
    Rule {
        pattern: "EOF",
        how: Match::Exact,
        class: Class::C,
        reason: "queries an external file",
    },
    Rule {
        pattern: "Input",
        how: Match::Exact,
        class: Class::C,
        reason: "reads bytes or characters from an external file",
    },
    Rule {
        pattern: "InputB",
        how: Match::Exact,
        class: Class::C,
        reason: "reads bytes from an external file",
    },
    Rule {
        pattern: "LOF",
        how: Match::Exact,
        class: Class::C,
        reason: "queries an external file length",
    },
    Rule {
        pattern: "Loc",
        how: Match::Exact,
        class: Class::C,
        reason: "queries an external file position",
    },
    Rule {
        pattern: "Seek",
        how: Match::Exact,
        class: Class::C,
        reason: "queries an external file position",
    },
    Rule {
        pattern: "FileLen",
        how: Match::Exact,
        class: Class::C,
        reason: "queries an external file length",
    },
    Rule {
        pattern: "FreeFile",
        how: Match::Exact,
        class: Class::C,
        reason: "allocates a native file number",
    },
    Rule {
        pattern: "Environ",
        how: Match::Exact,
        class: Class::C,
        reason: "reads the process environment",
    },
    Rule {
        pattern: "GetSetting",
        how: Match::Exact,
        class: Class::C,
        reason: "reads the current user's Windows registry settings",
    },
    Rule {
        pattern: "SaveSetting",
        how: Match::Exact,
        class: Class::C,
        reason: "writes the current user's Windows registry settings",
    },
    Rule {
        pattern: "DeleteSetting",
        how: Match::Exact,
        class: Class::C,
        reason: "deletes the current user's Windows registry settings",
    },
    Rule {
        pattern: "SaveAs",
        how: Match::Segment,
        class: Class::C,
        reason: "writes a file to a path",
    },
    Rule {
        pattern: "OpenText",
        how: Match::Segment,
        class: Class::C,
        reason: "reads an external file",
    },
    Rule {
        pattern: "QueryTables",
        how: Match::Segment,
        class: Class::C,
        reason: "pulls data from an external source",
    },
    // -- A: report generation ---------------------------------------------
    Rule {
        pattern: "PrintOut",
        how: Match::Segment,
        class: Class::A,
        reason: "prints",
    },
    Rule {
        pattern: "PageSetup",
        how: Match::Segment,
        class: Class::A,
        reason: "configures a printed page",
    },
    Rule {
        pattern: "Interior",
        how: Match::Segment,
        class: Class::A,
        reason: "sets cell fill",
    },
    Rule {
        pattern: "Font",
        how: Match::Segment,
        class: Class::A,
        reason: "sets fonts",
    },
    Rule {
        pattern: "Borders",
        how: Match::Segment,
        class: Class::A,
        reason: "sets borders",
    },
    Rule {
        pattern: "NumberFormat",
        how: Match::Segment,
        class: Class::A,
        reason: "sets number formats",
    },
    Rule {
        pattern: "MergeCells",
        how: Match::Segment,
        class: Class::A,
        reason: "merges cells",
    },
    Rule {
        pattern: "ColumnWidth",
        how: Match::Segment,
        class: Class::A,
        reason: "adjusts layout",
    },
    Rule {
        pattern: "RowHeight",
        how: Match::Segment,
        class: Class::A,
        reason: "adjusts layout",
    },
    // -- B: data transformation --------------------------------------------
    Rule {
        pattern: "Range",
        how: Match::Segment,
        class: Class::B,
        reason: "reads or writes cells",
    },
    Rule {
        pattern: "Cells",
        how: Match::Segment,
        class: Class::B,
        reason: "reads or writes cells",
    },
    Rule {
        pattern: "Evaluate",
        how: Match::Exact,
        class: Class::B,
        reason: "evaluates an Excel name or formula",
    },
    Rule {
        pattern: "Value",
        how: Match::Segment,
        class: Class::B,
        reason: "reads or writes cell values",
    },
    Rule {
        pattern: "Value2",
        how: Match::Segment,
        class: Class::B,
        reason: "reads or writes cell values",
    },
];

/// Names that indicate a dependency on Excel's calculation engine.
///
/// This is a separate axis from the class: a class B macro that writes formulas
/// and reads the results back needs a recalculation engine, and one that does
/// its arithmetic in VBA does not. The second is far cheaper to migrate.
const FORMULA_ENGINE_MARKERS: &[&str] = &[
    "Formula",
    "FormulaR1C1",
    "FormulaLocal",
    "FormulaArray",
    "WorksheetFunction",
    "Evaluate",
    "Calculate",
];

/// A construct worth telling the reader about, with the reason attached.
#[derive(Debug, Clone, PartialEq)]
pub struct Finding {
    pub what: String,
    pub reason: String,
    pub class: Option<Class>,
    pub line: u32,
}

#[derive(Debug, Clone, Default, PartialEq)]
pub struct Metrics {
    pub procedures: usize,
    pub statements: usize,
    /// Deepest nesting of control-flow blocks in any procedure.
    pub max_nesting: usize,
    pub longest_procedure: usize,
    /// Lines the parser could not interpret. Reported rather than hidden.
    pub unparsed: usize,
}

#[derive(Debug, Clone, PartialEq)]
pub struct ProcedureFacts {
    pub name: String,
    pub kind: ProcKind,
    pub visibility: Visibility,
    pub statements: usize,
    pub max_nesting: usize,
    /// Names this procedure mentions, for the call graph.
    pub calls: BTreeSet<String>,
    pub line: u32,
}

#[derive(Debug, Clone, PartialEq)]
pub struct Analysis {
    pub metrics: Metrics,
    pub procedures: Vec<ProcedureFacts>,
    /// Every dotted name the module mentions, with how often.
    pub api_names: BTreeMap<String, usize>,
    pub findings: Vec<Finding>,
    /// The worst class any rule matched. `None` when nothing matched at all.
    pub class: Option<Class>,
    pub needs_formula_engine: bool,
    /// Procedures nothing in this module calls. Not proof they are unused:
    /// see [`Analysis::dead_procedures`].
    pub uncalled_procedures: Vec<String>,
    /// `On Error Resume Next` swallows every error, so it is worth surfacing.
    pub blanket_error_handlers: usize,
    pub has_option_explicit: bool,
    pub external_declares: Vec<String>,
}

pub fn analyse(module: &Module) -> Analysis {
    let mut walker = Walker::default();
    walker.walk_module(module);
    walker.finish()
}

#[derive(Default)]
struct Walker {
    metrics: Metrics,
    procedures: Vec<ProcedureFacts>,
    api_names: BTreeMap<String, usize>,
    findings: Vec<Finding>,
    needs_formula_engine: bool,
    blanket_error_handlers: usize,
    has_option_explicit: bool,
    external_declares: Vec<String>,
    defined_procedures: Vec<(String, Visibility, ProcKind)>,
    module_names: BTreeSet<String>,

    // Per-procedure scratch.
    current_locals: BTreeSet<String>,
    current_calls: BTreeSet<String>,
    current_statements: usize,
    current_depth: usize,
    current_max_depth: usize,
}

impl Walker {
    fn walk_module(&mut self, module: &Module) {
        self.module_names = module
            .items
            .iter()
            .filter_map(|item| match item {
                ModuleItem::Variables(decl) => Some(decl),
                _ => None,
            })
            .flat_map(|decl| decl.items.iter())
            .map(|item| item.name.to_ascii_lowercase())
            .collect();
        for item in &module.items {
            match item {
                ModuleItem::Option(ModuleOption::Explicit, _) => self.has_option_explicit = true,
                ModuleItem::Option(..) => {}
                ModuleItem::DefType(_) => {}
                ModuleItem::Attribute { .. } => {}
                ModuleItem::Comment { .. } => {}
                ModuleItem::Directive { text, span } => {
                    self.findings.push(Finding {
                        what: text.clone(),
                        reason: "conditional compilation; the source differs by build".to_string(),
                        class: None,
                        line: span.line,
                    });
                }
                ModuleItem::Implements { interface, span } => {
                    self.record_type_name(interface, span.line);
                }
                ModuleItem::Unknown { .. } => self.metrics.unparsed += 1,
                ModuleItem::ExternalProc(d) => {
                    self.external_declares.push(d.name.clone());
                    self.findings.push(Finding {
                        what: format!("Declare {} Lib \"{}\"", d.name, d.lib),
                        reason: "calls a Windows API directly; unavailable outside Windows"
                            .to_string(),
                        class: Some(Class::C),
                        line: d.span.line,
                    });
                    for param in &d.params {
                        self.walk_type_name(&param.type_name, d.span.line);
                        if let Some(default) = &param.default {
                            self.walk_expr(default);
                        }
                    }
                    if let Some(return_type) = &d.return_type {
                        self.walk_type_name(return_type, d.span.line);
                    }
                }
                ModuleItem::Variables(v) => self.walk_var_decl(v),
                ModuleItem::Type(t) => {
                    for field in &t.fields {
                        self.walk_var_item(field);
                    }
                }
                ModuleItem::Enum(e) => {
                    for (_, value) in &e.members {
                        if let Some(value) = value {
                            self.walk_expr(value);
                        }
                    }
                }
                ModuleItem::Event { params, span, .. } => {
                    for param in params {
                        self.walk_type_name(&param.type_name, span.line);
                        if let Some(default) = &param.default {
                            self.walk_expr(default);
                        }
                    }
                }
                ModuleItem::Procedure(p) => self.walk_procedure(p),
            }
        }
    }

    fn walk_procedure(&mut self, proc: &Procedure) {
        self.current_locals = self.module_names.clone();
        // A Function or Property Get assigns its return value through the
        // procedure name.  Treating that bare name as a call makes every such
        // private procedure appear used by itself.  The same rule also keeps
        // self-recursion from hiding an otherwise unreachable procedure.
        self.current_locals.insert(proc.name.to_ascii_lowercase());
        self.current_locals.extend(
            proc.params
                .iter()
                .map(|param| param.name.to_ascii_lowercase()),
        );
        collect_declared_names(&proc.body, &mut self.current_locals);
        self.current_calls = BTreeSet::new();
        self.current_statements = 0;
        self.current_depth = 0;
        self.current_max_depth = 0;

        for param in &proc.params {
            self.walk_type_name(&param.type_name, proc.span.line);
            if let Some(default) = &param.default {
                self.walk_expr(default);
            }
        }
        if let Some(return_type) = &proc.return_type {
            self.walk_type_name(return_type, proc.span.line);
        }
        self.walk_body(&proc.body);

        self.defined_procedures
            .push((proc.name.clone(), proc.visibility, proc.kind));
        self.metrics.procedures += 1;
        self.metrics.statements += self.current_statements;
        self.metrics.max_nesting = self.metrics.max_nesting.max(self.current_max_depth);
        self.metrics.longest_procedure =
            self.metrics.longest_procedure.max(self.current_statements);

        self.procedures.push(ProcedureFacts {
            name: proc.name.clone(),
            kind: proc.kind,
            visibility: proc.visibility,
            statements: self.current_statements,
            max_nesting: self.current_max_depth,
            calls: std::mem::take(&mut self.current_calls),
            line: proc.span.line,
        });
    }

    fn nested(&mut self, body: &[Statement]) {
        self.current_depth += 1;
        self.current_max_depth = self.current_max_depth.max(self.current_depth);
        self.walk_body(body);
        self.current_depth -= 1;
    }

    fn walk_body(&mut self, body: &[Statement]) {
        for stmt in body {
            self.walk_statement(stmt);
        }
    }

    fn walk_statement(&mut self, stmt: &Statement) {
        self.current_statements += 1;
        match stmt {
            Statement::Assign { target, value, .. }
            | Statement::SetAssign { target, value, .. } => {
                self.walk_expr(target);
                self.walk_expr(value);
            }
            Statement::MidAssign(mid) => {
                self.walk_expr(&mid.target);
                self.walk_expr(&mid.start);
                if let Some(length) = &mid.length {
                    self.walk_expr(length);
                }
                self.walk_expr(&mid.value);
            }
            Statement::AlignedAssign(aligned) => {
                self.walk_expr(&aligned.target);
                self.walk_expr(&aligned.value);
            }
            Statement::Call { target, .. } => self.walk_expr(target),
            Statement::Dim(v) => self.walk_var_decl(v),
            Statement::ReDim { items, .. } => {
                for item in items {
                    self.walk_var_item(item);
                }
            }
            Statement::Erase { targets, .. } => {
                for t in targets {
                    self.walk_expr(t);
                }
            }
            Statement::Open(open) => {
                self.walk_expr(&open.path);
                self.walk_expr(&open.file_number);
                if let Some(record_len) = &open.record_len {
                    self.walk_expr(record_len);
                }
                self.record_file_io("Open", open.span.line, "opens an external file");
            }
            Statement::Close { files, span } => {
                for file in files {
                    self.walk_expr(file);
                }
                self.record_file_io("Close", span.line, "operates on external files");
            }
            Statement::FileOutput(output) => {
                self.walk_expr(&output.file_number);
                for item in &output.items {
                    if let Some(value) = &item.value {
                        self.walk_expr(value);
                    }
                }
                let what = match output.kind {
                    FileOutputKind::Print => "Print #",
                    FileOutputKind::Write => "Write #",
                };
                self.record_file_io(what, output.span.line, "writes to an external file");
            }
            Statement::FileInput(input) => {
                self.walk_expr(&input.file_number);
                for target in &input.targets {
                    self.walk_expr(target);
                }
                let what = if input.line {
                    "Line Input #"
                } else {
                    "Input #"
                };
                self.record_file_io(what, input.span.line, "reads from an external file");
            }
            Statement::FileTransfer(transfer) => {
                self.walk_expr(&transfer.file_number);
                if let Some(record_number) = &transfer.record_number {
                    self.walk_expr(record_number);
                }
                self.walk_expr(&transfer.value);
                let (what, reason) = match transfer.kind {
                    FileTransferKind::Get => ("Get #", "reads from an external file"),
                    FileTransferKind::Put => ("Put #", "writes to an external file"),
                };
                self.record_file_io(what, transfer.span.line, reason);
            }
            Statement::FileSeek(seek) => {
                self.walk_expr(&seek.file_number);
                self.walk_expr(&seek.position);
                self.record_file_io(
                    "Seek #",
                    seek.span.line,
                    "changes an external file position",
                );
            }
            Statement::FileSystem(operation) => match operation {
                FileSystemStmt::Rename {
                    source,
                    destination,
                    span,
                } => {
                    self.walk_expr(source);
                    self.walk_expr(destination);
                    self.record_file_io("Name", span.line, "renames an external file or directory");
                }
                FileSystemStmt::Copy {
                    source,
                    destination,
                    span,
                } => {
                    self.walk_expr(source);
                    self.walk_expr(destination);
                    self.record_file_io("FileCopy", span.line, "copies an external file");
                }
                FileSystemStmt::Unary { kind, path, span } => {
                    self.walk_expr(path);
                    let (what, reason) = match kind {
                        FileSystemUnaryKind::Kill => ("Kill", "deletes an external file"),
                        FileSystemUnaryKind::MkDir => ("MkDir", "creates an external directory"),
                        FileSystemUnaryKind::RmDir => ("RmDir", "removes an external directory"),
                        FileSystemUnaryKind::ChDir => {
                            ("ChDir", "changes the process working directory")
                        }
                        FileSystemUnaryKind::ChDrive => {
                            ("ChDrive", "changes the process working drive")
                        }
                    };
                    self.record_file_io(what, span.line, reason);
                }
                FileSystemStmt::SetAttr {
                    path,
                    attributes,
                    span,
                } => {
                    self.walk_expr(path);
                    self.walk_expr(attributes);
                    self.record_file_io("SetAttr", span.line, "changes external file attributes");
                }
            },
            Statement::FileRecordLock(lock) => {
                self.walk_expr(&lock.file_number);
                if let Some(start) = &lock.start {
                    self.walk_expr(start);
                }
                if let Some(end) = &lock.end {
                    self.walk_expr(end);
                }
                let what = match lock.kind {
                    FileRecordLockKind::Lock => "Lock",
                    FileRecordLockKind::Unlock => "Unlock",
                };
                self.record_file_io(
                    what,
                    lock.span.line,
                    "controls external file record locking",
                );
            }
            Statement::FileWidth {
                file_number,
                width,
                span,
            } => {
                self.walk_expr(file_number);
                self.walk_expr(width);
                self.record_file_io("Width #", span.line, "configures external file output");
            }
            Statement::FileReset { span } => {
                self.record_file_io("Reset", span.line, "closes every open external disk file");
            }
            Statement::If(s) => {
                self.walk_expr(&s.condition);
                self.nested(&s.then_body);
                for (cond, body) in &s.else_ifs {
                    self.walk_expr(cond);
                    self.nested(body);
                }
                if let Some(body) = &s.else_body {
                    self.nested(body);
                }
            }
            Statement::SelectCase(s) => {
                self.walk_expr(&s.subject);
                for case in &s.cases {
                    for label in &case.labels {
                        match label {
                            CaseLabel::Value(e) | CaseLabel::Compare(_, e) => self.walk_expr(e),
                            CaseLabel::Range(a, b) => {
                                self.walk_expr(a);
                                self.walk_expr(b);
                            }
                        }
                    }
                    self.nested(&case.body);
                }
                if let Some(body) = &s.case_else {
                    self.nested(body);
                }
            }
            Statement::For(s) => {
                self.walk_expr(&s.counter);
                self.walk_expr(&s.from);
                self.walk_expr(&s.to);
                if let Some(step) = &s.step {
                    self.walk_expr(step);
                }
                self.nested(&s.body);
            }
            Statement::ForEach(s) => {
                self.walk_expr(&s.item);
                self.walk_expr(&s.collection);
                self.nested(&s.body);
            }
            Statement::Do(s) => {
                for test in [&s.pre, &s.post].into_iter().flatten() {
                    self.walk_expr(&test.condition);
                }
                self.nested(&s.body);
            }
            Statement::While {
                condition, body, ..
            } => {
                self.walk_expr(condition);
                self.nested(body);
            }
            Statement::With { subject, body, .. } => {
                self.walk_expr(subject);
                self.nested(body);
            }
            Statement::OnError(OnError::ResumeNext { span }) => {
                self.blanket_error_handlers += 1;
                self.findings.push(Finding {
                    what: "On Error Resume Next".to_string(),
                    reason: "swallows every error from here on; failures become silent wrong \
                             answers rather than stops"
                        .to_string(),
                    class: None,
                    line: span.line,
                });
            }
            Statement::OnBranch(branch) => {
                self.walk_expr(&branch.selector);
            }
            Statement::RaiseEvent(event) => {
                for arg in &event.args {
                    if let Some(value) = &arg.value {
                        self.walk_expr(value);
                    }
                }
            }
            Statement::Unknown { text, span } => {
                self.metrics.unparsed += 1;
                self.findings.push(Finding {
                    what: text.clone(),
                    reason: "not understood by the parser; excluded from every other count"
                        .to_string(),
                    class: None,
                    line: span.line,
                });
            }
            Statement::Directive { text, span } => {
                self.findings.push(Finding {
                    what: text.clone(),
                    reason: "conditional compilation; the source differs by build".to_string(),
                    class: None,
                    line: span.line,
                });
            }
            _ => {}
        }
    }

    fn walk_var_decl(&mut self, decl: &VarDecl) {
        for item in &decl.items {
            self.walk_var_item(item);
        }
    }

    fn walk_var_item(&mut self, item: &VarItem) {
        if let Some(bounds) = &item.array_bounds {
            for bound in bounds {
                if let Some(lower) = &bound.lower {
                    self.walk_expr(lower);
                }
                self.walk_expr(&bound.upper);
            }
        }
        if let Some(value) = &item.value {
            self.walk_expr(value);
        }
        // `Dim x As Scripting.FileSystemObject` is just as much a dependency as
        // calling `CreateObject`, and early binding is the common spelling.
        self.walk_type_name(&item.type_name, 0);
    }

    fn walk_type_name(&mut self, type_name: &TypeName, line: u32) {
        self.record_type_name(&type_name.name, line);
        if let Some(length) = &type_name.fixed_length {
            self.walk_expr(length);
        }
    }

    fn record_type_name(&mut self, name: &str, line: u32) {
        if is_intrinsic_type(name) {
            return;
        }
        self.record_dependency_name(name, line, false);
    }

    fn walk_expr(&mut self, expr: &Expr) {
        let mut names = Vec::new();
        collect_names(expr, &mut names);
        for (name, line) in names {
            self.record_name(&name, line);
        }
    }

    fn record_name(&mut self, name: &str, line: u32) {
        if name.is_empty() {
            return;
        }
        if !name.contains('.') && self.current_locals.contains(&name.to_ascii_lowercase()) {
            return;
        }
        self.record_dependency_name(name, line, true);
    }

    fn record_dependency_name(&mut self, name: &str, line: u32, counts_as_call: bool) {
        if name.is_empty() {
            return;
        }
        *self.api_names.entry(name.to_string()).or_default() += 1;

        if counts_as_call {
            if let Some(root) = name.split('.').next() {
                self.current_calls.insert(root.to_string());
            }
            self.current_calls.insert(name.to_string());
        }

        if segments(name).any(|s| {
            FORMULA_ENGINE_MARKERS
                .iter()
                .any(|m| s.eq_ignore_ascii_case(m))
        }) {
            self.needs_formula_engine = true;
        }

        if let Some(rule) = match_rule(name) {
            self.findings.push(Finding {
                what: name.to_string(),
                reason: rule.reason.to_string(),
                class: Some(rule.class),
                line,
            });
        }
    }

    fn record_file_io(&mut self, name: &str, line: u32, reason: &str) {
        *self.api_names.entry(name.to_string()).or_default() += 1;
        self.findings.push(Finding {
            what: name.to_string(),
            reason: reason.to_string(),
            class: Some(Class::C),
            line,
        });
    }

    fn finish(mut self) -> Analysis {
        let class = self
            .findings
            .iter()
            .filter_map(|f| f.class)
            .max_by_key(|c| c.severity());

        let mut procedures_by_name: BTreeMap<String, Vec<usize>> = BTreeMap::new();
        for (index, procedure) in self.procedures.iter().enumerate() {
            procedures_by_name
                .entry(procedure.name.to_ascii_lowercase())
                .or_default()
                .push(index);
        }

        // A mention in dead code does not make its callee live.  Start with
        // procedures the host or another module may call, then follow only
        // calls made by those reachable procedures.
        let mut reachable = BTreeSet::new();
        let mut pending = Vec::new();
        for procedure in &self.procedures {
            if is_externally_reachable(&procedure.name, procedure.visibility, procedure.kind) {
                let name = procedure.name.to_ascii_lowercase();
                if reachable.insert(name.clone()) {
                    pending.push(name);
                }
            }
        }
        while let Some(name) = pending.pop() {
            let Some(indices) = procedures_by_name.get(&name) else {
                continue;
            };
            for &index in indices {
                for call in &self.procedures[index].calls {
                    let callee = call.to_ascii_lowercase();
                    if procedures_by_name.contains_key(&callee) && reachable.insert(callee.clone())
                    {
                        pending.push(callee);
                    }
                }
            }
        }

        let uncalled = self
            .defined_procedures
            .iter()
            .filter(|(name, _, _)| !reachable.contains(&name.to_ascii_lowercase()))
            .map(|(name, _, _)| name.clone())
            .collect();

        self.findings.sort_by(|a, b| a.line.cmp(&b.line));

        Analysis {
            metrics: self.metrics,
            procedures: self.procedures,
            api_names: self.api_names,
            findings: self.findings,
            class,
            needs_formula_engine: self.needs_formula_engine,
            uncalled_procedures: uncalled,
            blanket_error_handlers: self.blanket_error_handlers,
            has_option_explicit: self.has_option_explicit,
            external_declares: self.external_declares,
        }
    }
}

fn collect_declared_names(body: &[Statement], names: &mut BTreeSet<String>) {
    for statement in body {
        match statement {
            Statement::Dim(decl) => {
                names.extend(decl.items.iter().map(|item| item.name.to_ascii_lowercase()))
            }
            Statement::ReDim { items, .. } => {
                names.extend(items.iter().map(|item| item.name.to_ascii_lowercase()));
            }
            Statement::If(if_statement) => {
                collect_declared_names(&if_statement.then_body, names);
                for (_, branch) in &if_statement.else_ifs {
                    collect_declared_names(branch, names);
                }
                if let Some(branch) = &if_statement.else_body {
                    collect_declared_names(branch, names);
                }
            }
            Statement::SelectCase(select) => {
                for case in &select.cases {
                    collect_declared_names(&case.body, names);
                }
                if let Some(branch) = &select.case_else {
                    collect_declared_names(branch, names);
                }
            }
            Statement::For(for_statement) => {
                if let Expr::Ident(name, _) = &for_statement.counter {
                    names.insert(name.to_ascii_lowercase());
                }
                collect_declared_names(&for_statement.body, names);
            }
            Statement::ForEach(for_each) => {
                if let Expr::Ident(name, _) = &for_each.item {
                    names.insert(name.to_ascii_lowercase());
                }
                collect_declared_names(&for_each.body, names);
            }
            Statement::Do(do_statement) => collect_declared_names(&do_statement.body, names),
            Statement::While { body, .. } | Statement::With { body, .. } => {
                collect_declared_names(body, names);
            }
            _ => {}
        }
    }
}

/// Whether something outside this module can reach a procedure.
///
/// Getting this wrong is how a diagnosis tool tells someone to delete a button
/// handler. Public procedures are callable from anywhere, event handlers are
/// called by the host, and property accessors are reached through their name.
fn is_externally_reachable(name: &str, visibility: Visibility, kind: ProcKind) -> bool {
    if matches!(
        visibility,
        Visibility::Public | Visibility::Global | Visibility::Friend | Visibility::Default
    ) {
        return true;
    }
    if !matches!(kind, ProcKind::Sub | ProcKind::Function) {
        return true;
    }
    // Worksheet_Change, Workbook_Open, CommandButton1_Click, ...
    name.contains('_')
}

fn segments(name: &str) -> impl Iterator<Item = &str> {
    name.split('.')
}

fn match_rule(name: &str) -> Option<&'static Rule> {
    let mut best: Option<&'static Rule> = None;
    for rule in RULES {
        let root = segments(name).next().unwrap_or("");
        let hit = match rule.how {
            Match::Exact => name.eq_ignore_ascii_case(rule.pattern),
            Match::Root => root.eq_ignore_ascii_case(rule.pattern),
            Match::RootPrefix => {
                root.len() >= rule.pattern.len()
                    && root[..rule.pattern.len()].eq_ignore_ascii_case(rule.pattern)
            }
            Match::Segment => segments(name).any(|s| s.eq_ignore_ascii_case(rule.pattern)),
        };
        if !hit {
            continue;
        }
        // Worst class wins: one Shell call outranks any amount of formatting.
        if best.is_none_or(|b| rule.class.severity() > b.class.severity()) {
            best = Some(rule);
        }
    }
    best
}

fn is_intrinsic_type(name: &str) -> bool {
    [
        "Any", "Boolean", "Byte", "Currency", "Date", "Decimal", "Double", "Integer", "Long",
        "LongLong", "LongPtr", "Object", "Single", "String", "Variant",
    ]
    .iter()
    .any(|intrinsic| name.eq_ignore_ascii_case(intrinsic))
}

/// Collect the longest dotted name of each member chain, so that
/// `Application.WorksheetFunction.Sum` is recorded once rather than three times.
fn collect_names(expr: &Expr, out: &mut Vec<(String, u32)>) {
    match expr {
        Expr::EvaluateShortcut { span, .. } => {
            out.push(("Evaluate".to_string(), span.line));
        }
        Expr::Ident(..) | Expr::Member { .. } => {
            if let Some(name) = expr.dotted_name() {
                out.push((name, expr.span().line));
                // Still descend into any arguments hidden inside the chain.
            }
            if let Expr::Member { object, .. } = expr {
                descend_arguments(object, out);
            }
        }
        Expr::Index { target, args, .. } => {
            if let Some(name) = expr.dotted_name() {
                out.push((name, expr.span().line));
            } else {
                collect_names(target, out);
            }
            descend_arguments(target, out);
            for arg in args {
                if let Some(value) = &arg.value {
                    collect_names(value, out);
                }
            }
        }
        Expr::Bang { object, name, span } => {
            out.push((name.clone(), span.line));
            collect_names(object, out);
        }
        Expr::New { type_name, span } => out.push((type_name.clone(), span.line)),
        Expr::AddressOf { procedure, span } => out.push((procedure.clone(), span.line)),
        Expr::Unary { operand, .. } | Expr::TypeOf { operand, .. } => collect_names(operand, out),
        Expr::Binary { lhs, rhs, .. } => {
            collect_names(lhs, out);
            collect_names(rhs, out);
        }
        Expr::Literal(..) | Expr::WithMember(..) => {}
    }
}

/// A call can hide inside a chain: `Sheets(Name).Range(Addr)`. The chain's own
/// name is recorded by the caller; this picks up the arguments along it.
fn descend_arguments(expr: &Expr, out: &mut Vec<(String, u32)>) {
    match expr {
        Expr::Index { target, args, .. } => {
            descend_arguments(target, out);
            for arg in args {
                if let Some(value) = &arg.value {
                    collect_names(value, out);
                }
            }
        }
        Expr::Member { object, .. } | Expr::Bang { object, .. } => descend_arguments(object, out),
        _ => {}
    }
}

impl Analysis {
    /// Procedures that look unused, with the caveat spelled out.
    ///
    /// Never a deletion instruction: a `Public` procedure can be called from any
    /// other module, from a ribbon button, or from `Application.Run` with a name
    /// built at run time. Only procedures private to this module and not
    /// obviously an event handler are listed at all.
    pub fn dead_procedures(&self) -> &[String] {
        &self.uncalled_procedures
    }

    /// One-line verdict, with the reason that decided it.
    pub fn verdict(&self) -> String {
        match self.class {
            Some(class) => {
                let reason = self
                    .findings
                    .iter()
                    .find(|f| f.class == Some(class))
                    .map(|f| f.reason.as_str())
                    .unwrap_or("");
                format!("{} ({}): {}", class.as_str(), class.description(), reason)
            }
            None => "unclassified: no recognised Excel or external API used".to_string(),
        }
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::parse_module;

    fn analyse_src(src: &str) -> Analysis {
        analyse(&parse_module(src).expect("should parse"))
    }

    #[test]
    fn plain_cell_arithmetic_is_class_b() {
        let a =
            analyse_src("Sub Total()\n  Range(\"A1\").Value = Range(\"A2\").Value + 1\nEnd Sub");
        assert_eq!(a.class, Some(Class::B));
        assert!(!a.needs_formula_engine);
    }

    #[test]
    fn excel_bracket_expressions_require_evaluation() {
        let a = analyse_src("Sub Total()\n  value = [A1] + [SUM(A1:A2)]\nEnd Sub");
        assert_eq!(a.class, Some(Class::B));
        assert!(a.needs_formula_engine);
        assert_eq!(a.api_names.get("Evaluate"), Some(&2));
        assert_eq!(a.metrics.unparsed, 0);
    }

    #[test]
    fn formatting_lifts_it_to_class_a() {
        let a = analyse_src(
            "Sub Report()\n  Range(\"A1\").Value = 1\n  Range(\"A1\").Interior.Color = 255\nEnd Sub",
        );
        assert_eq!(a.class, Some(Class::A));
    }

    #[test]
    fn one_shell_call_outranks_everything_else() {
        let a = analyse_src(
            "Sub Mixed()\n  Range(\"A1\").Interior.Color = 255\n  Shell \"cmd.exe\"\nEnd Sub",
        );
        assert_eq!(a.class, Some(Class::C));
        assert!(a.verdict().contains("external process"));
    }

    #[test]
    fn a_userform_makes_it_a_rewrite() {
        let a = analyse_src("Sub Show()\n  UserForm1.Show\n  Shell \"x\"\nEnd Sub");
        assert_eq!(a.class, Some(Class::D));
    }

    #[test]
    fn early_bound_external_types_count_as_much_as_createobject() {
        let early = analyse_src("Sub A()\n  Dim c As ADODB.Connection\nEnd Sub");
        let late = analyse_src("Sub B()\n  Set c = CreateObject(\"ADODB.Connection\")\nEnd Sub");
        assert_eq!(early.class, Some(Class::C));
        assert_eq!(late.class, Some(Class::C));
    }

    #[test]
    fn declaration_signature_types_are_dependencies() {
        let a = analyse_src(
            "Implements ADODB.Connection\n\
             Public Event Ready(ByVal rows As ADODB.Recordset)\n\
             Private Function Export(ByVal fs As Scripting.FileSystemObject) As Outlook.MailItem\n\
             End Function",
        );
        for name in [
            "ADODB.Connection",
            "ADODB.Recordset",
            "Scripting.FileSystemObject",
            "Outlook.MailItem",
        ] {
            assert_eq!(a.api_names.get(name), Some(&1), "missing {name}");
        }
        assert_eq!(a.class, Some(Class::C));
        assert_eq!(a.metrics.unparsed, 0);
    }

    #[test]
    fn intrinsic_declaration_types_are_not_api_dependencies() {
        let a = analyse_src(
            "Private Const Width As Long = 4\n\
             Private Function Convert(ByVal number As Long) As String\n\
               Dim ready As Boolean\n\
               Dim legacy&\n\
               Dim padded As String * Width\n\
             End Function",
        );
        assert!(a.api_names.is_empty(), "unexpected APIs: {:?}", a.api_names);
        assert_eq!(a.metrics.unparsed, 0);
    }

    #[test]
    fn declare_statements_are_flagged_without_being_called() {
        let a = analyse_src(
            "Private Declare PtrSafe Function Sleep Lib \"kernel32\" (ByVal ms As Long) As Long\n",
        );
        assert_eq!(a.class, Some(Class::C));
        assert_eq!(a.external_declares, vec!["Sleep".to_string()]);
    }

    #[test]
    fn addressof_callback_counts_as_a_procedure_reference() {
        let a = analyse_src(
            "Public Sub Hook()\n\
               timerId = SetTimer(0, 0, 1000, AddressOf TimerProc)\n\
             End Sub\n\
             Private Sub TimerProc()\n\
             End Sub",
        );
        assert!(!a.uncalled_procedures.contains(&"Hook".to_string()));
        assert!(!a.uncalled_procedures.contains(&"TimerProc".to_string()));
        assert_eq!(a.metrics.unparsed, 0);
    }

    #[test]
    fn addressof_from_dead_code_does_not_make_the_callback_reachable() {
        let a = analyse_src(
            "Private Sub DeadHook()\n\
               timerId = SetTimer(0, 0, 1000, AddressOf DeadTimerProc)\n\
             End Sub\n\
             Private Sub DeadTimerProc()\n\
             End Sub",
        );
        assert_eq!(
            a.dead_procedures(),
            ["DeadHook".to_string(), "DeadTimerProc".to_string()]
        );
    }

    #[test]
    fn the_formula_engine_axis_is_independent_of_the_class() {
        let plain = analyse_src("Sub A()\n  Range(\"A1\").Value = 1\nEnd Sub");
        let needs = analyse_src("Sub B()\n  Range(\"A1\").Formula = \"=SUM(B:B)\"\nEnd Sub");
        let wsf = analyse_src(
            "Sub C()\n  x = Application.WorksheetFunction.Sum(Range(\"A:A\"))\nEnd Sub",
        );
        assert_eq!(plain.class, Some(Class::B));
        assert!(!plain.needs_formula_engine);
        assert!(needs.needs_formula_engine);
        assert!(wsf.needs_formula_engine);
    }

    #[test]
    fn member_chains_are_recorded_once_at_full_length() {
        let a = analyse_src(
            "Sub T()\n  x = Application.WorksheetFunction.Sum(Range(\"A:A\"))\nEnd Sub",
        );
        assert_eq!(
            a.api_names.get("Application.WorksheetFunction.Sum"),
            Some(&1)
        );
        // The arguments inside the chain are still seen.
        assert!(a.api_names.contains_key("Range"));
    }

    #[test]
    fn declared_names_do_not_impersonate_excel_api_members() {
        let a = analyse_src(
            "Private Value As Long\n\
             Sub T(ByVal Font As Long)\n\
               Dim Range As Long\n\
               Value = Font + Range\n\
             End Sub",
        );
        assert_eq!(a.class, None);
        assert!(!a.api_names.contains_key("Value"));
        assert!(!a.api_names.contains_key("Font"));
        assert!(!a.api_names.contains_key("Range"));
    }

    #[test]
    fn declared_object_roots_keep_their_member_api_chain() {
        let a = analyse_src(
            "Sub T()\n\
               Dim ws As Worksheet\n\
               ws.Range(\"A1\").Value = 1\n\
             End Sub",
        );
        assert_eq!(a.class, Some(Class::B));
        assert!(a
            .api_names
            .keys()
            .any(|name| name.eq_ignore_ascii_case("ws.Range.Value")));
    }

    #[test]
    fn blanket_error_handlers_are_surfaced() {
        let a = analyse_src("Sub T()\n  On Error Resume Next\n  x = 1\nEnd Sub");
        assert_eq!(a.blanket_error_handlers, 1);
        assert!(a
            .findings
            .iter()
            .any(|f| f.reason.contains("silent wrong answers")));
    }

    #[test]
    fn private_uncalled_procedures_are_listed_but_public_ones_are_not() {
        let a = analyse_src(
            "Public Sub Entry()\n  HelperUsed\nEnd Sub\n\
             Private Sub HelperUsed()\nEnd Sub\n\
             Private Sub HelperUnused()\nEnd Sub\n",
        );
        assert_eq!(a.dead_procedures(), ["HelperUnused".to_string()]);
    }

    #[test]
    fn local_labels_do_not_impersonate_procedure_calls() {
        let a = analyse_src(
            "Public Sub Entry()\n\
             GoTo LabelCollision\n\
             LabelCollision:\n\
             On 1 GoSub LabelCollision\n\
             End Sub\n\
             Private Sub LabelCollision()\n\
             End Sub\n",
        );
        assert_eq!(a.dead_procedures(), ["LabelCollision".to_string()]);
    }

    #[test]
    fn procedure_call_matching_is_case_insensitive() {
        let a = analyse_src(
            "Public Sub Entry()\n\
             helperused\n\
             End Sub\n\
             Private Sub HelperUsed()\n\
             End Sub\n",
        );
        assert!(a.dead_procedures().is_empty());
    }

    #[test]
    fn function_result_assignment_is_not_a_self_call() {
        let a = analyse_src(
            "Private Function HiddenValue() As Long\n\
             HiddenValue = 7\n\
             End Function\n",
        );
        assert_eq!(a.dead_procedures(), ["HiddenValue".to_string()]);
    }

    #[test]
    fn self_recursion_does_not_make_a_procedure_reachable() {
        let a = analyse_src(
            "Private Sub RecursiveOnly()\n\
             RecursiveOnly\n\
             End Sub\n",
        );
        assert_eq!(a.dead_procedures(), ["RecursiveOnly".to_string()]);
    }

    #[test]
    fn calls_from_dead_code_do_not_make_their_callees_reachable() {
        let a = analyse_src(
            "Private Sub DeadCaller()\n\
             DeadCallee\n\
             End Sub\n\
             Private Sub DeadCallee()\n\
             End Sub\n",
        );
        assert_eq!(
            a.dead_procedures(),
            ["DeadCaller".to_string(), "DeadCallee".to_string()]
        );
    }

    #[test]
    fn calls_from_public_entry_points_are_followed_transitively() {
        let a = analyse_src(
            "Public Sub Entry()\n\
             FirstStep\n\
             End Sub\n\
             Private Sub FirstStep()\n\
             LastStep\n\
             End Sub\n\
             Private Sub LastStep()\n\
             End Sub\n",
        );
        assert!(a.dead_procedures().is_empty());
    }

    #[test]
    fn event_handlers_are_never_reported_as_dead() {
        // Nothing in the module calls this; the host does.
        let a = analyse_src("Private Sub Worksheet_Change(ByVal Target As Range)\nEnd Sub");
        assert!(a.dead_procedures().is_empty());
    }

    #[test]
    fn metrics_measure_nesting_and_size() {
        let a = analyse_src(
            "Sub T()\n  For i = 1 To 3\n    If x Then\n      Do\n        y = 1\n      Loop\n    End If\n  Next i\nEnd Sub",
        );
        assert_eq!(a.metrics.procedures, 1);
        assert_eq!(a.metrics.max_nesting, 3);
        assert!(a.metrics.statements >= 4);
    }

    #[test]
    fn unparsed_lines_are_counted_not_hidden() {
        let a = analyse_src("Sub T()\n  Get #1\n  x = 1\nEnd Sub");
        assert_eq!(a.metrics.unparsed, 1);
        assert!(a
            .findings
            .iter()
            .any(|f| f.reason.contains("not understood")));
    }

    #[test]
    fn native_file_io_is_parsed_and_classified_as_external() {
        let a = analyse_src(
            "Sub T()\n\
             Open \"f.txt\" For Output As #1\n\
             Print #1, Range(\"A1\").Value\n\
             Close #1\n\
             End Sub",
        );
        assert_eq!(a.metrics.unparsed, 0);
        assert_eq!(a.class, Some(Class::C));
        assert_eq!(a.api_names.get("Open"), Some(&1));
        assert_eq!(a.api_names.get("Print #"), Some(&1));
        assert_eq!(a.api_names.get("Close"), Some(&1));
        assert!(a.findings.iter().any(|f| f.what == "Open" && f.line == 2));
        assert!(a
            .api_names
            .keys()
            .any(|name| name.eq_ignore_ascii_case("Range.Value")));
    }

    #[test]
    fn binary_file_io_and_position_functions_are_class_c() {
        let a = analyse_src(
            "Sub T()\n\
             Put #1, 1, value\n\
             Seek #1, 1\n\
             Get #1, , value\n\
             n = LOF(1) + Loc(1) + Seek(1) + FileLen(path) + FreeFile\n\
             done = EOF(1)\n\
             End Sub",
        );
        assert_eq!(a.metrics.unparsed, 0);
        assert_eq!(a.class, Some(Class::C));
        for name in [
            "Put #", "Seek #", "Get #", "LOF", "Loc", "Seek", "FileLen", "FreeFile", "EOF",
        ] {
            assert!(
                a.api_names.contains_key(name),
                "missing {name}: {:#?}",
                a.api_names
            );
        }
    }

    #[test]
    fn input_function_with_hash_file_number_is_class_c() {
        let a = analyse_src("Sub T()\nvalue = Input$(2, #handle)\nEnd Sub");
        assert_eq!(a.class, Some(Class::C));
        assert_eq!(a.api_names.get("Input"), Some(&1));
        assert_eq!(a.metrics.unparsed, 0);
    }

    #[test]
    fn filesystem_statements_and_queries_are_class_c() {
        let a = analyse_src(
            "Sub T()\n\
             FileCopy sourcePath, copyPath\n\
             Name copyPath As renamedPath\n\
             SetAttr renamedPath, vbHidden\n\
             Kill renamedPath\n\
             MkDir directoryPath\n\
             ChDrive driveName\n\
             ChDir directoryPath\n\
             RmDir directoryPath\n\
             attr = GetAttr(sourcePath)\n\
             cwd = CurDir$\n\
             End Sub",
        );
        assert_eq!(a.metrics.unparsed, 0);
        assert_eq!(a.class, Some(Class::C));
        for name in [
            "FileCopy", "Name", "SetAttr", "Kill", "MkDir", "ChDrive", "ChDir", "RmDir", "GetAttr",
            "CurDir",
        ] {
            assert!(a.api_names.contains_key(name), "missing {name}");
        }
    }

    #[test]
    fn registry_setting_apis_are_class_c() {
        let a = analyse_src(
            "Sub T()\n\
             value = GetSetting(\"Oxi\", \"Probe\", \"Value\", \"missing\")\n\
             SaveSetting \"Oxi\", \"Probe\", \"Value\", value\n\
             DeleteSetting \"Oxi\", \"Probe\"\n\
             End Sub",
        );
        assert_eq!(a.class, Some(Class::C));
        assert_eq!(a.api_names.get("GetSetting"), Some(&1));
        assert_eq!(a.api_names.get("SaveSetting"), Some(&1));
        assert_eq!(a.api_names.get("DeleteSetting"), Some(&1));
        assert_eq!(a.metrics.unparsed, 0);
    }

    #[test]
    fn lock_width_and_reset_are_class_c() {
        let a = analyse_src(
            "Sub T()\n\
             Lock #1, 1 To 4\n\
             Unlock #1, 1 To 4\n\
             Width #1, 80\n\
             Reset\n\
             End Sub",
        );
        assert_eq!(a.metrics.unparsed, 0);
        assert_eq!(a.class, Some(Class::C));
        for name in ["Lock", "Unlock", "Width #", "Reset"] {
            assert_eq!(a.api_names.get(name), Some(&1), "missing {name}");
        }
    }

    #[test]
    fn computed_on_branch_is_parsed_and_walks_its_selector() {
        let a = analyse_src(
            "Sub T()\n\
             On Range(\"A1\").Value GoTo First, Second\n\
             First:\n\
             Second:\n\
             End Sub",
        );
        assert_eq!(a.metrics.unparsed, 0);
        assert!(a.api_names.contains_key("Range.Value"));
    }

    #[test]
    fn special_string_assignments_walk_all_expressions() {
        let a = analyse_src(
            "Sub T()\n\
             Mid$(Range(\"A1\").Value, startAt, fieldLength) = replacement\n\
             LSet target = Range(\"B1\").Value\n\
             End Sub",
        );
        assert_eq!(a.metrics.unparsed, 0);
        assert_eq!(a.api_names.get("Range.Value"), Some(&2));
    }

    #[test]
    fn raiseevent_walks_argument_expressions() {
        let a = analyse_src(
            "Sub Fire()\n\
             RaiseEvent Fired(42, Range(\"A1\").Value)\n\
             RaiseEvent Ping\n\
             End Sub",
        );
        assert_eq!(a.metrics.unparsed, 0);
        assert_eq!(a.api_names.get("Range.Value"), Some(&1));
    }

    #[test]
    fn option_explicit_is_noticed() {
        assert!(analyse_src("Option Explicit\nSub T()\nEnd Sub").has_option_explicit);
        assert!(!analyse_src("Sub T()\nEnd Sub").has_option_explicit);
    }

    #[test]
    fn deftype_is_understood_module_context() {
        let a = analyse_src("DefInt A-C\nDefStr S\nSub T()\nDim apple, sample\nEnd Sub");
        assert_eq!(a.metrics.unparsed, 0);
        assert_eq!(a.metrics.procedures, 1);
    }

    #[test]
    fn module_options_are_understood_without_becoming_api_calls() {
        let a = analyse_src(
            "Option Explicit\nOption Base 1\nOption Compare Text\nOption Private Module\nSub T()\nEnd Sub",
        );
        assert_eq!(a.metrics.unparsed, 0);
        assert!(a.has_option_explicit);
        assert!(a.api_names.is_empty());
    }

    #[test]
    fn module_comments_and_directives_are_not_unparsed() {
        let a = analyse_src(
            "' platform declaration\n\
             #If Win64 Then\n\
             Private value As LongLong\n\
             #End If\n",
        );
        assert_eq!(a.metrics.unparsed, 0);
        assert!(a
            .findings
            .iter()
            .any(|finding| finding.what == "#If Win64 Then"));
        assert!(a.findings.iter().any(|finding| finding.what == "#End If"));
    }

    #[test]
    fn a_module_with_no_excel_api_is_unclassified_rather_than_guessed() {
        let a = analyse_src("Sub T()\n  x = 1 + 2\nEnd Sub");
        assert_eq!(a.class, None);
        assert!(a.verdict().contains("unclassified"));
    }
}
