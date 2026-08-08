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
    Rule {
        pattern: "FileDialog",
        how: Match::Segment,
        class: Class::D,
        reason: "opens an Office file-selection user interface",
    },
    Rule {
        pattern: "GetOpenFilename",
        how: Match::Segment,
        class: Class::D,
        reason: "opens Excel's file-open selection interface",
    },
    Rule {
        pattern: "GetSaveAsFilename",
        how: Match::Segment,
        class: Class::D,
        reason: "opens Excel's save-as selection interface",
    },
    Rule {
        pattern: "InputBox",
        how: Match::Segment,
        class: Class::D,
        reason: "prompts the user for interactive input",
    },
    Rule {
        pattern: "MsgBox",
        how: Match::Segment,
        class: Class::D,
        reason: "shows an interactive message box",
    },
    Rule {
        pattern: "StatusBar",
        how: Match::Segment,
        class: Class::D,
        reason: "writes progress or status text to Excel's user interface",
    },
    Rule {
        pattern: "DisplayStatusBar",
        how: Match::Segment,
        class: Class::D,
        reason: "shows or hides Excel's status-bar user interface",
    },
    Rule {
        pattern: "DisplayFormulaBar",
        how: Match::Segment,
        class: Class::D,
        reason: "shows or hides Excel's formula-bar user interface",
    },
    Rule {
        pattern: "DisplayFullScreen",
        how: Match::Segment,
        class: Class::D,
        reason: "switches Excel's application window into or out of full-screen mode",
    },
    Rule {
        pattern: "Application.Cursor",
        how: Match::Exact,
        class: Class::D,
        reason: "changes Excel's process-global mouse-pointer user interface",
    },
    Rule {
        pattern: "Application.ShowWindowsInTaskbar",
        how: Match::Exact,
        class: Class::D,
        reason: "requests showing or hiding Excel workbook windows in the Windows taskbar; modern Excel may ignore it",
    },
    Rule {
        pattern: "Application.WindowState",
        how: Match::Exact,
        class: Class::D,
        reason: "changes Excel's application-window minimized, normal, or maximized UI state",
    },
    Rule {
        pattern: "DisplayGridlines",
        how: Match::Segment,
        class: Class::D,
        reason: "shows or hides worksheet gridlines in an Excel window",
    },
    Rule {
        pattern: "DisplayHeadings",
        how: Match::Segment,
        class: Class::D,
        reason: "shows or hides worksheet row and column headings in an Excel window",
    },
    Rule {
        pattern: "DisplayWorkbookTabs",
        how: Match::Segment,
        class: Class::D,
        reason: "shows or hides workbook sheet tabs in an Excel window",
    },
    Rule {
        pattern: "DisplayHorizontalScrollBar",
        how: Match::Segment,
        class: Class::D,
        reason: "shows or hides the horizontal scroll bar in an Excel window",
    },
    Rule {
        pattern: "DisplayVerticalScrollBar",
        how: Match::Segment,
        class: Class::D,
        reason: "shows or hides the vertical scroll bar in an Excel window",
    },
    Rule {
        pattern: "ActiveWindow.Zoom",
        how: Match::Exact,
        class: Class::D,
        reason: "changes the zoom level of Excel's active window",
    },
    Rule {
        pattern: "Application.ActiveWindow.Zoom",
        how: Match::Exact,
        class: Class::D,
        reason: "changes the zoom level of Excel's active window",
    },
    Rule {
        pattern: "ActiveWindow.View",
        how: Match::Exact,
        class: Class::D,
        reason: "changes Excel's active-window worksheet view mode",
    },
    Rule {
        pattern: "Application.ActiveWindow.View",
        how: Match::Exact,
        class: Class::D,
        reason: "changes Excel's active-window worksheet view mode",
    },
    Rule {
        pattern: "SplitRow",
        how: Match::Segment,
        class: Class::D,
        reason: "changes the pane layout of an Excel window",
    },
    Rule {
        pattern: "SplitColumn",
        how: Match::Segment,
        class: Class::D,
        reason: "changes the pane layout of an Excel window",
    },
    Rule {
        pattern: "FreezePanes",
        how: Match::Segment,
        class: Class::D,
        reason: "changes the pane layout of an Excel window",
    },
    Rule {
        pattern: "ScrollRow",
        how: Match::Segment,
        class: Class::D,
        reason: "changes the first visible row in an Excel window",
    },
    Rule {
        pattern: "ScrollColumn",
        how: Match::Segment,
        class: Class::D,
        reason: "changes the first visible column in an Excel window",
    },
    Rule {
        pattern: "Activate",
        how: Match::Segment,
        class: Class::D,
        reason: "changes Excel's active workbook, sheet, or object UI context",
    },
    Rule {
        pattern: "Select",
        how: Match::Segment,
        class: Class::D,
        reason: "changes Excel's active selection UI context",
    },
    Rule {
        pattern: "Goto",
        how: Match::Segment,
        class: Class::D,
        reason: "activates and scrolls to an Excel range or object",
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
        pattern: "AppActivate",
        how: Match::Exact,
        class: Class::C,
        reason: "activates a desktop application window by title",
    },
    Rule {
        pattern: "SendKeys",
        how: Match::Exact,
        class: Class::C,
        reason: "injects keystrokes into the active desktop application",
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
        pattern: "Workbooks.Open",
        how: Match::Exact,
        class: Class::C,
        reason: "opens an external workbook from a path",
    },
    Rule {
        pattern: "Application.Workbooks.Open",
        how: Match::Exact,
        class: Class::C,
        reason: "opens an external workbook from a path",
    },
    Rule {
        pattern: "LinkSources",
        how: Match::Segment,
        class: Class::C,
        reason: "enumerates external workbook links",
    },
    Rule {
        pattern: "UpdateLink",
        how: Match::Segment,
        class: Class::C,
        reason: "refreshes data from an external workbook link",
    },
    Rule {
        pattern: "ChangeLink",
        how: Match::Segment,
        class: Class::C,
        reason: "retargets an external workbook link",
    },
    Rule {
        pattern: "BreakLink",
        how: Match::Segment,
        class: Class::C,
        reason: "removes an external workbook link and replaces formulas with values",
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
        pattern: "Formula2",
        how: Match::Segment,
        class: Class::B,
        reason: "reads or writes dynamic-array formulas",
    },
    Rule {
        pattern: "Formula2R1C1",
        how: Match::Segment,
        class: Class::B,
        reason: "reads or writes dynamic-array formulas in R1C1 notation",
    },
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
        pattern: "ExecuteExcel4Macro",
        how: Match::Segment,
        class: Class::B,
        reason: "executes a legacy Excel 4.0 macro string",
    },
    Rule {
        pattern: "ConvertFormula",
        how: Match::Segment,
        class: Class::B,
        reason: "converts formula references between A1 and R1C1 notation",
    },
    Rule {
        pattern: "Volatile",
        how: Match::Segment,
        class: Class::B,
        reason: "marks a VBA function for execution on every Excel recalculation",
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
    "Formula2",
    "Formula2R1C1",
    "WorksheetFunction",
    "Evaluate",
    "ExecuteExcel4Macro",
    "ConvertFormula",
    "Volatile",
    "Calculate",
    "CalculateFull",
    "CalculateFullRebuild",
    "Calculation",
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
    with_subjects: Vec<Option<String>>,
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
        self.with_subjects.clear();

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
                let parent = self.with_subjects.last().and_then(|name| name.as_deref());
                let resolved = resolve_expr_name(subject, parent);
                self.with_subjects.push(resolved);
                self.nested(body);
                self.with_subjects.pop();
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
        let with_subject = self.with_subjects.last().and_then(|name| name.as_deref());
        collect_names(expr, with_subject, &mut names);
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

        if name.eq_ignore_ascii_case("Application.Run") {
            self.findings.push(Finding {
                what: name.to_string(),
                reason: "dispatches a macro by name; target resolution requires workbook context"
                    .to_string(),
                class: None,
                line,
            });
        }
        if name.eq_ignore_ascii_case("CallByName") || name.eq_ignore_ascii_case("VBA.CallByName") {
            self.findings.push(Finding {
                what: name.to_string(),
                reason:
                    "dispatches an object member by name; target resolution requires runtime type context"
                        .to_string(),
                class: None,
                line,
            });
        }
        if name.eq_ignore_ascii_case("Application.OnTime") {
            self.findings.push(Finding {
                what: name.to_string(),
                reason:
                    "schedules a macro by name; execution depends on Excel's application event loop"
                        .to_string(),
                class: None,
                line,
            });
        }
        if name.eq_ignore_ascii_case("Application.Wait") {
            self.findings.push(Finding {
                what: name.to_string(),
                reason:
                    "blocks Excel until a wall-clock deadline and suspends application activity"
                        .to_string(),
                class: None,
                line,
            });
        }
        if name.eq_ignore_ascii_case("Application.International") {
            self.findings.push(Finding {
                what: name.to_string(),
                reason: "reads Excel locale settings; behavior can vary by machine".to_string(),
                class: None,
                line,
            });
        }
        if name.eq_ignore_ascii_case("Application.Caller")
            || name
                .get(.."Application.Caller.".len())
                .is_some_and(|prefix| prefix.eq_ignore_ascii_case("Application.Caller."))
        {
            self.findings.push(Finding {
                what: name.to_string(),
                reason: "reads the Excel invocation context; behavior depends on the calling cell or object"
                    .to_string(),
                class: None,
                line,
            });
        }
        if [
            "ActiveWorkbook",
            "ActiveSheet",
            "ActiveCell",
            "ActiveWindow",
            "Selection",
            "Application.ActiveWorkbook",
            "Application.ActiveSheet",
            "Application.ActiveCell",
            "Application.ActiveWindow",
            "Application.Selection",
        ]
        .iter()
        .any(|context| {
            name.eq_ignore_ascii_case(context)
                || (name
                    .get(..context.len())
                    .is_some_and(|prefix| prefix.eq_ignore_ascii_case(context))
                    && name.as_bytes().get(context.len()) == Some(&b'.'))
        }) {
            self.findings.push(Finding {
                what: name.to_string(),
                reason: "reads Excel's active UI context; result depends on the current window, workbook, sheet, cell, or selection"
                    .to_string(),
                class: None,
                line,
            });
        }
        let terminal = name.rsplit('.').next().unwrap_or(name);
        if segments(name).any(|segment| segment.eq_ignore_ascii_case("Range"))
            && terminal.eq_ignore_ascii_case("Find")
        {
            self.findings.push(Finding {
                what: name.to_string(),
                reason: "uses Excel's stateful Range.Find settings; omitted options can inherit previous UI or VBA choices"
                    .to_string(),
                class: None,
                line,
            });
        }
        if segments(name).any(|segment| segment.eq_ignore_ascii_case("Range"))
            && ["FindNext", "FindPrevious"]
                .iter()
                .any(|member| terminal.eq_ignore_ascii_case(member))
        {
            self.findings.push(Finding {
                what: name.to_string(),
                reason: "continues Excel's stateful preceding Range.Find operation".to_string(),
                class: None,
                line,
            });
        }
        if segments(name).any(|segment| segment.eq_ignore_ascii_case("Range"))
            && terminal.eq_ignore_ascii_case("Replace")
        {
            self.findings.push(Finding {
                what: name.to_string(),
                reason: "uses Excel's stateful Range.Replace settings; omitted options can inherit previous UI or VBA choices"
                    .to_string(),
                class: None,
                line,
            });
        }
        for (format_state, reason) in [
            (
                "Application.FindFormat",
                "reads or changes Excel's process-global find-format criteria",
            ),
            (
                "Application.ReplaceFormat",
                "reads or changes Excel's process-global replace-format criteria",
            ),
        ] {
            if name.eq_ignore_ascii_case(format_state)
                || (name
                    .get(..format_state.len())
                    .is_some_and(|prefix| prefix.eq_ignore_ascii_case(format_state))
                    && name.as_bytes().get(format_state.len()) == Some(&b'.'))
            {
                self.findings.push(Finding {
                    what: name.to_string(),
                    reason: reason.to_string(),
                    class: None,
                    line,
                });
            }
        }
        if name.eq_ignore_ascii_case("Application.CutCopyMode") {
            self.findings.push(Finding {
                what: name.to_string(),
                reason: "reads or changes Excel's process-global clipboard and cut/copy mode"
                    .to_string(),
                class: None,
                line,
            });
        }
        if segments(name).any(|segment| segment.eq_ignore_ascii_case("Range"))
            && ["Copy", "Cut"]
                .iter()
                .any(|member| terminal.eq_ignore_ascii_case(member))
        {
            self.findings.push(Finding {
                what: name.to_string(),
                reason: "changes Excel's process-global clipboard and cut/copy mode".to_string(),
                class: None,
                line,
            });
        }
        if name.eq_ignore_ascii_case("Application.CalculateFull") {
            self.findings.push(Finding {
                what: name.to_string(),
                reason: "forces a full Excel workbook recalculation".to_string(),
                class: None,
                line,
            });
        }
        if name.eq_ignore_ascii_case("Application.CalculateFullRebuild") {
            self.findings.push(Finding {
                what: name.to_string(),
                reason: "rebuilds Excel's formula dependencies and recalculates every workbook"
                    .to_string(),
                class: None,
                line,
            });
        }
        if name.eq_ignore_ascii_case("DoEvents") || name.eq_ignore_ascii_case("VBA.DoEvents") {
            self.findings.push(Finding {
                what: name.to_string(),
                reason: "yields to the Windows/Office event loop and permits reentrant execution"
                    .to_string(),
                class: None,
                line,
            });
        }
        if name.eq_ignore_ascii_case("Err.Raise") {
            self.findings.push(Finding {
                what: name.to_string(),
                reason: "raises a VBA runtime error whose number, source, and description are observable"
                    .to_string(),
                class: None,
                line,
            });
        }
        if [
            "Now",
            "Date",
            "Time",
            "Timer",
            "VBA.Now",
            "VBA.Date",
            "VBA.Time",
            "VBA.Timer",
        ]
        .iter()
        .any(|clock| name.eq_ignore_ascii_case(clock))
        {
            self.findings.push(Finding {
                what: name.to_string(),
                reason:
                    "reads the system clock; results depend on execution time and local time zone"
                        .to_string(),
                class: None,
                line,
            });
        }
        if ["Rnd", "Randomize", "VBA.Rnd", "VBA.Randomize"]
            .iter()
            .any(|random| name.eq_ignore_ascii_case(random))
        {
            self.findings.push(Finding {
                what: name.to_string(),
                reason: "uses VBA's process-global pseudorandom generator; results depend on seed and call order"
                    .to_string(),
                class: None,
                line,
            });
        }
        if name.eq_ignore_ascii_case("Application.EnableEvents") {
            self.findings.push(Finding {
                what: name.to_string(),
                reason: "changes Excel's process-global event delivery state".to_string(),
                class: None,
                line,
            });
        }
        if name.eq_ignore_ascii_case("Application.Calculation") {
            self.findings.push(Finding {
                what: name.to_string(),
                reason: "changes Excel's process-global automatic calculation mode".to_string(),
                class: None,
                line,
            });
        }
        if name.eq_ignore_ascii_case("Application.DisplayAlerts") {
            self.findings.push(Finding {
                what: name.to_string(),
                reason:
                    "changes Excel's process-global alert handling and automatic default responses"
                        .to_string(),
                class: None,
                line,
            });
        }
        if name.eq_ignore_ascii_case("Application.AutomationSecurity") {
            self.findings.push(Finding {
                what: name.to_string(),
                reason:
                    "changes Excel's process-global macro policy for programmatically opened files"
                        .to_string(),
                class: None,
                line,
            });
        }
        if name.eq_ignore_ascii_case("Application.ScreenUpdating") {
            self.findings.push(Finding {
                what: name.to_string(),
                reason: "changes Excel's process-global screen redraw state".to_string(),
                class: None,
                line,
            });
        }
        if name.eq_ignore_ascii_case("Application.AskToUpdateLinks") {
            self.findings.push(Finding {
                what: name.to_string(),
                reason: "changes Excel's process-global prompt policy for updating external links"
                    .to_string(),
                class: None,
                line,
            });
        }
        if name.eq_ignore_ascii_case("Application.Interactive") {
            self.findings.push(Finding {
                what: name.to_string(),
                reason: "changes Excel's process-global keyboard and mouse input state".to_string(),
                class: None,
                line,
            });
        }

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
                if let Expr::Ident(name, _) | Expr::TypedIdent { name, .. } = &for_statement.counter
                {
                    names.insert(name.to_ascii_lowercase());
                }
                collect_declared_names(&for_statement.body, names);
            }
            Statement::ForEach(for_each) => {
                if let Expr::Ident(name, _) | Expr::TypedIdent { name, .. } = &for_each.item {
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
fn resolve_expr_name(expr: &Expr, with_subject: Option<&str>) -> Option<String> {
    match expr {
        Expr::Ident(name, _) | Expr::TypedIdent { name, .. } => Some(name.clone()),
        Expr::WithMember(name, _) | Expr::WithBangMember(name, _) => {
            Some(format!("{}.{}", with_subject?, name))
        }
        Expr::Member { object, name, .. } | Expr::Bang { object, name, .. } => Some(format!(
            "{}.{}",
            resolve_expr_name(object, with_subject)?,
            name
        )),
        Expr::Index { target, .. } => resolve_expr_name(target, with_subject),
        _ => None,
    }
}

fn collect_names(expr: &Expr, with_subject: Option<&str>, out: &mut Vec<(String, u32)>) {
    match expr {
        Expr::EvaluateShortcut { span, .. } => {
            out.push(("Evaluate".to_string(), span.line));
        }
        Expr::Ident(..)
        | Expr::TypedIdent { .. }
        | Expr::WithMember(..)
        | Expr::WithBangMember(..)
        | Expr::Member { .. }
        | Expr::Bang { .. } => {
            if let Some(name) = resolve_expr_name(expr, with_subject) {
                out.push((name, expr.span().line));
                // Still descend into any arguments hidden inside the chain.
            }
            if let Expr::Member { object, .. } | Expr::Bang { object, .. } = expr {
                descend_arguments(object, with_subject, out);
            }
        }
        Expr::Index { target, args, .. } => {
            if let Some(name) = resolve_expr_name(expr, with_subject) {
                out.push((name, expr.span().line));
            } else {
                collect_names(target, with_subject, out);
            }
            descend_arguments(target, with_subject, out);
            for arg in args {
                if let Some(value) = &arg.value {
                    collect_names(value, with_subject, out);
                }
            }
        }
        Expr::New { type_name, span } => out.push((type_name.clone(), span.line)),
        Expr::AddressOf { procedure, span } => out.push((procedure.clone(), span.line)),
        Expr::Unary { operand, .. } | Expr::TypeOf { operand, .. } => {
            collect_names(operand, with_subject, out)
        }
        Expr::Binary { lhs, rhs, .. } => {
            collect_names(lhs, with_subject, out);
            collect_names(rhs, with_subject, out);
        }
        Expr::Literal(..) => {}
    }
}

/// A call can hide inside a chain: `Sheets(Name).Range(Addr)`. The chain's own
/// name is recorded by the caller; this picks up the arguments along it.
fn descend_arguments(expr: &Expr, with_subject: Option<&str>, out: &mut Vec<(String, u32)>) {
    match expr {
        Expr::Index { target, args, .. } => {
            descend_arguments(target, with_subject, out);
            for arg in args {
                if let Some(value) = &arg.value {
                    collect_names(value, with_subject, out);
                }
            }
        }
        Expr::Member { object, .. } | Expr::Bang { object, .. } => {
            descend_arguments(object, with_subject, out)
        }
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
    fn bang_default_members_are_walked_without_unparsed_input() {
        let a = analyse_src(
            "Sub ReadField()\n  Dim record As Object\n  value = record!Answer + record![Display Name]\nEnd Sub",
        );
        assert_eq!(a.metrics.unparsed, 0);
        assert_eq!(a.api_names.get("record.Answer"), Some(&1));
        assert_eq!(a.api_names.get("record.Display Name"), Some(&1));
    }

    #[test]
    fn bang_and_dot_member_chains_are_recorded_once_at_full_length() {
        let a = analyse_src(
            "Sub ReadCell()\n  Dim book As Object\n  Dim value As Variant\n  value = book!Sheets!Data.Range(\"A1\").Value\nEnd Sub",
        );
        assert_eq!(a.metrics.unparsed, 0);
        assert_eq!(a.api_names.get("book.Sheets.Data.Range.Value"), Some(&1));
        assert_eq!(a.api_names.len(), 1);
        assert_eq!(a.class, Some(Class::B));
    }

    #[test]
    fn with_relative_members_resolve_against_the_subject() {
        let a = analyse_src(
            "Sub FormatCell()\n  Dim cell As Object\n  Set cell = Range(\"A1\")\n  With cell\n    .Value = 1\n    .Font.Bold = True\n  End With\nEnd Sub",
        );
        assert_eq!(a.metrics.unparsed, 0);
        assert_eq!(a.api_names.get("cell.Value"), Some(&1));
        assert_eq!(a.api_names.get("cell.Font.Bold"), Some(&1));
        assert_eq!(a.class, Some(Class::A));
    }

    #[test]
    fn nested_with_and_bang_subjects_keep_the_full_chain() {
        let a = analyse_src(
            "Sub ReadCell()\n  Dim book As Object\n  Dim value As Variant\n  With book!Sheets!Data\n    With .Range(\"A1\")\n      value = .Value\n    End With\n  End With\nEnd Sub",
        );
        assert_eq!(a.metrics.unparsed, 0);
        assert_eq!(a.api_names.get("book.Sheets.Data.Range.Value"), Some(&1));
        assert_eq!(a.class, Some(Class::B));
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
    fn formula2_variants_require_the_formula_engine() {
        let a = analyse_src(
            "Sub T()\n  Range(\"A1\").Formula2 = \"=SEQUENCE(2)\"\n  Range(\"C1\").Formula2R1C1 = \"=RC[-1]*2\"\nEnd Sub",
        );
        assert_eq!(a.class, Some(Class::B));
        assert!(a.needs_formula_engine);
        assert_eq!(a.api_names.get("Range.Formula2"), Some(&1));
        assert_eq!(a.api_names.get("Range.Formula2R1C1"), Some(&1));
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
    fn application_run_is_reported_without_guessing_a_private_target() {
        let a = analyse_src(
            "Public Sub Entry()\n\
             Application.Run \"Module1.HiddenValue\"\n\
             End Sub\n\
             Private Function HiddenValue() As Long\n\
             HiddenValue = 42\n\
             End Function\n",
        );
        assert_eq!(a.dead_procedures(), ["HiddenValue".to_string()]);
        assert_eq!(a.api_names.get("Application.Run"), Some(&1));
        assert!(a.findings.iter().any(|finding| {
            finding.what.eq_ignore_ascii_case("Application.Run")
                && finding.reason.contains("workbook context")
                && finding.class.is_none()
        }));
    }

    #[test]
    fn callbyname_is_reported_as_runtime_member_dispatch() {
        let a = analyse_src(
            "Public Function ReadValue(ByVal target As Object) As Variant\n\
             CallByName target, \"Value\", VbLet, 40\n\
             ReadValue = VBA.CallByName(target, \"Value\", VbGet)\n\
             End Function\n",
        );
        assert_eq!(a.api_names.get("CallByName"), Some(&1));
        assert_eq!(a.api_names.get("VBA.CallByName"), Some(&1));
        assert_eq!(
            a.findings
                .iter()
                .filter(|finding| finding.reason.contains("runtime type context"))
                .count(),
            2
        );
        assert!(a
            .findings
            .iter()
            .filter(|finding| finding.reason.contains("runtime type context"))
            .all(|finding| finding.class.is_none()));
    }

    #[test]
    fn application_ontime_is_reported_as_scheduled_macro_dispatch() {
        let a = analyse_src(
            "Public Sub ScheduleReport()\n\
             Application.OnTime EarliestTime:=Now + TimeSerial(0, 0, 1), Procedure:=\"ReportModule.RefreshReport\"\n\
             End Sub\n",
        );
        assert_eq!(a.api_names.get("Application.OnTime"), Some(&1));
        assert!(a.findings.iter().any(|finding| {
            finding.what.eq_ignore_ascii_case("Application.OnTime")
                && finding.reason.contains("application event loop")
                && finding.class.is_none()
        }));
    }

    #[test]
    fn application_wait_is_reported_as_blocking_clock_dependency() {
        let a = analyse_src(
            "Public Function WaitOneSecond() As Boolean\n\
             WaitOneSecond = Application.Wait(Now + TimeSerial(0, 0, 1))\n\
             End Function\n",
        );
        assert_eq!(a.metrics.unparsed, 0);
        assert_eq!(a.api_names.get("Application.Wait"), Some(&1));
        assert!(a.findings.iter().any(|finding| {
            finding.what.eq_ignore_ascii_case("Application.Wait")
                && finding.reason.contains("wall-clock deadline")
                && finding.class.is_none()
        }));
    }

    #[test]
    fn execute_excel4_macro_is_a_formula_engine_dependency() {
        let a = analyse_src(
            "Public Function LegacyFormula() As Double\n\
             LegacyFormula = Application.ExecuteExcel4Macro(\"SUM(40,2)\")\n\
             End Function\n",
        );
        assert_eq!(a.class, Some(Class::B));
        assert!(a.needs_formula_engine);
        assert_eq!(a.api_names.get("Application.ExecuteExcel4Macro"), Some(&1));
        assert!(a.findings.iter().any(|finding| {
            finding
                .what
                .eq_ignore_ascii_case("Application.ExecuteExcel4Macro")
                && finding.reason.contains("legacy Excel 4.0 macro")
        }));
    }

    #[test]
    fn application_international_is_reported_as_locale_dependent() {
        let a = analyse_src(
            "Public Function Separators() As String\n\
             Separators = Application.International(xlDecimalSeparator) & Application.International(xlListSeparator)\n\
             End Function\n",
        );
        assert_eq!(a.api_names.get("Application.International"), Some(&2));
        assert_eq!(
            a.findings
                .iter()
                .filter(|finding| finding.reason.contains("locale settings"))
                .count(),
            2
        );
        assert!(a
            .findings
            .iter()
            .filter(|finding| finding.reason.contains("locale settings"))
            .all(|finding| finding.class.is_none()));
    }

    #[test]
    fn convert_formula_is_a_formula_engine_dependency() {
        let a = analyse_src(
            "Public Function ConvertedAddress() As String\n\
             ConvertedAddress = Application.ConvertFormula(\"=R2C2\", xlR1C1, xlA1)\n\
             End Function\n",
        );
        assert_eq!(a.class, Some(Class::B));
        assert!(a.needs_formula_engine);
        assert_eq!(a.api_names.get("Application.ConvertFormula"), Some(&1));
        assert!(a.findings.iter().any(|finding| {
            finding
                .what
                .eq_ignore_ascii_case("Application.ConvertFormula")
                && finding.reason.contains("A1 and R1C1")
        }));
    }

    #[test]
    fn application_caller_is_reported_through_member_chains() {
        let a = analyse_src(
            "Public Function CallerAddress() As String\n\
             CallerAddress = Application.Caller.Address(False, False)\n\
             End Function\n",
        );
        assert_eq!(a.api_names.get("Application.Caller.Address"), Some(&1));
        assert!(a.findings.iter().any(|finding| {
            finding
                .what
                .eq_ignore_ascii_case("Application.Caller.Address")
                && finding.reason.contains("calling cell or object")
                && finding.class.is_none()
        }));
    }

    #[test]
    fn active_excel_context_is_reported_through_member_chains() {
        let a = analyse_src(
            "Public Function ActiveContext() As String\n\
             Sheet1.Activate\n\
             Sheet1.Range(\"B2\").Select\n\
             ActiveContext = Application.ActiveWorkbook.Name & \"|\" & Application.ActiveSheet.Name & \"|\" & Application.ActiveCell.Address(False, False) & \"|\" & Application.Selection.Address(False, False)\n\
             End Function\n",
        );
        assert_eq!(a.metrics.unparsed, 0);
        assert_eq!(a.class, Some(Class::D));
        assert!(a.findings.iter().any(|finding| {
            finding.what.eq_ignore_ascii_case("Sheet1.Activate")
                && finding.reason.contains("active workbook")
                && finding.class == Some(Class::D)
        }));
        assert!(a.findings.iter().any(|finding| {
            finding.what.eq_ignore_ascii_case("Sheet1.Range.Select")
                && finding.reason.contains("active selection")
                && finding.class == Some(Class::D)
        }));
        for api in [
            "Application.ActiveWorkbook.Name",
            "Application.ActiveSheet.Name",
            "Application.ActiveCell.Address",
            "Application.Selection.Address",
        ] {
            assert_eq!(a.api_names.get(api), Some(&1), "missing {api}");
            assert!(a.findings.iter().any(|finding| {
                finding.what.eq_ignore_ascii_case(api)
                    && finding.reason.contains("active UI context")
                    && finding.class.is_none()
            }));
        }
    }

    #[test]
    fn application_goto_is_a_user_interface_dependency() {
        let a = analyse_src(
            "Public Function GoToCell() As String\n\
             Application.Goto Reference:=Sheet1.Range(\"C3\"), Scroll:=True\n\
             GoToCell = Application.ActiveCell.Address(False, False)\n\
             End Function\n",
        );
        assert_eq!(a.metrics.unparsed, 0);
        assert_eq!(a.class, Some(Class::D));
        assert_eq!(a.api_names.get("Application.Goto"), Some(&1));
        assert!(a.findings.iter().any(|finding| {
            finding.what.eq_ignore_ascii_case("Application.Goto")
                && finding.reason.contains("activates and scrolls")
                && finding.class == Some(Class::D)
        }));
        assert!(a.findings.iter().any(|finding| {
            finding
                .what
                .eq_ignore_ascii_case("Application.ActiveCell.Address")
                && finding.reason.contains("active UI context")
                && finding.class.is_none()
        }));
    }

    #[test]
    fn range_find_is_reported_as_stateful_search() {
        let a = analyse_src(
            "Public Function FindValues() As String\n\
             Dim found As Range\n\
             Dim following As Range\n\
             Set found = Sheet1.Range(\"A1:A3\").Find(What:=42, After:=Sheet1.Range(\"A1\"), LookIn:=xlValues, LookAt:=xlWhole, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:=False)\n\
             Set following = Sheet1.Range(\"A1:A3\").FindNext(After:=found)\n\
             FindValues = found.Address(False, False) & \"|\" & following.Address(False, False)\n\
             End Function\n",
        );
        assert_eq!(a.metrics.unparsed, 0);
        assert_eq!(a.class, Some(Class::B));
        assert_eq!(a.api_names.get("Sheet1.Range.Find"), Some(&1));
        assert_eq!(a.api_names.get("Sheet1.Range.FindNext"), Some(&1));
        assert!(a.findings.iter().any(|finding| {
            finding.what.eq_ignore_ascii_case("Sheet1.Range.Find")
                && finding.reason.contains("omitted options")
                && finding.class.is_none()
        }));
        assert!(a.findings.iter().any(|finding| {
            finding.what.eq_ignore_ascii_case("Sheet1.Range.FindNext")
                && finding.reason.contains("preceding Range.Find")
                && finding.class.is_none()
        }));
    }

    #[test]
    fn range_replace_is_reported_as_stateful_operation() {
        let a = analyse_src(
            "Public Function ReplaceValues() As String\n\
             Sheet1.Range(\"A1:A3\").Replace What:=\"foo\", Replacement:=\"baz\", LookAt:=xlWhole, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False\n\
             ReplaceValues = Sheet1.Range(\"A1\").Value2 & \"|\" & Sheet1.Range(\"A2\").Value2 & \"|\" & Sheet1.Range(\"A3\").Value2\n\
             End Function\n",
        );
        assert_eq!(a.metrics.unparsed, 0);
        assert_eq!(a.class, Some(Class::B));
        assert_eq!(a.api_names.get("Sheet1.Range.Replace"), Some(&1));
        assert!(a.findings.iter().any(|finding| {
            finding.what.eq_ignore_ascii_case("Sheet1.Range.Replace")
                && finding.reason.contains("omitted options")
                && finding.class.is_none()
        }));
    }

    #[test]
    fn application_search_formats_are_global_state() {
        let a = analyse_src(
            "Public Function ExerciseSearchFormats() As Long\n\
             Application.FindFormat.Clear\n\
             Application.ReplaceFormat.Clear\n\
             Application.FindFormat.Font.Bold = True\n\
             Application.ReplaceFormat.Font.Italic = True\n\
             ExerciseSearchFormats = CLng(Application.FindFormat.Font.Bold) + CLng(Application.ReplaceFormat.Font.Italic)\n\
             Application.FindFormat.Clear\n\
             Application.ReplaceFormat.Clear\n\
             End Function\n",
        );
        assert_eq!(a.metrics.unparsed, 0);
        assert_eq!(a.class, Some(Class::A));
        assert_eq!(a.api_names.get("Application.FindFormat.Clear"), Some(&2));
        assert_eq!(a.api_names.get("Application.ReplaceFormat.Clear"), Some(&2));
        assert_eq!(
            a.api_names.get("Application.FindFormat.Font.Bold"),
            Some(&2)
        );
        assert_eq!(
            a.api_names.get("Application.ReplaceFormat.Font.Italic"),
            Some(&2)
        );
        assert_eq!(
            a.findings
                .iter()
                .filter(|finding| finding.reason.contains("find-format criteria"))
                .count(),
            4
        );
        assert_eq!(
            a.findings
                .iter()
                .filter(|finding| finding.reason.contains("replace-format criteria"))
                .count(),
            4
        );
    }

    #[test]
    fn cut_copy_operations_are_global_clipboard_state() {
        let a = analyse_src(
            "Public Function ExerciseCutCopyMode() As String\n\
             Sheet1.Range(\"A1\").Value2 = \"copied\"\n\
             Sheet1.Range(\"A1\").Copy\n\
             ExerciseCutCopyMode = CStr(CLng(Application.CutCopyMode))\n\
             Application.CutCopyMode = False\n\
             ExerciseCutCopyMode = ExerciseCutCopyMode & \"|\" & CStr(CLng(Application.CutCopyMode))\n\
             Sheet1.Range(\"A1\").Cut Destination:=Sheet1.Range(\"B1\")\n\
             End Function\n",
        );
        assert_eq!(a.metrics.unparsed, 0);
        assert_eq!(a.class, Some(Class::B));
        assert_eq!(a.api_names.get("Application.CutCopyMode"), Some(&3));
        assert_eq!(a.api_names.get("Sheet1.Range.Copy"), Some(&1));
        assert_eq!(a.api_names.get("Sheet1.Range.Cut"), Some(&1));
        assert_eq!(
            a.findings
                .iter()
                .filter(|finding| finding.reason.contains("clipboard and cut/copy mode"))
                .count(),
            5
        );
    }

    #[test]
    fn application_volatile_is_a_recalculation_dependency() {
        let a = analyse_src(
            "Public Function RecalculatedValue() As Long\n\
             Application.Volatile\n\
             RecalculatedValue = 42\n\
             End Function\n",
        );
        assert_eq!(a.class, Some(Class::B));
        assert!(a.needs_formula_engine);
        assert_eq!(a.api_names.get("Application.Volatile"), Some(&1));
        assert!(a.findings.iter().any(|finding| {
            finding.what.eq_ignore_ascii_case("Application.Volatile")
                && finding.reason.contains("every Excel recalculation")
        }));
    }

    #[test]
    fn full_recalculation_variants_are_formula_engine_dependencies() {
        let a = analyse_src(
            "Public Sub RecalculateEverything()\n\
             Application.CalculateFull\n\
             Application.CalculateFullRebuild\n\
             End Sub\n",
        );
        assert_eq!(a.class, None);
        assert!(a.needs_formula_engine);
        assert_eq!(a.api_names.get("Application.CalculateFull"), Some(&1));
        assert_eq!(
            a.api_names.get("Application.CalculateFullRebuild"),
            Some(&1)
        );
        assert!(a
            .findings
            .iter()
            .any(|finding| finding.reason.contains("full Excel workbook")));
        assert!(a
            .findings
            .iter()
            .any(|finding| finding.reason.contains("formula dependencies")));
    }

    #[test]
    fn desktop_activation_and_sendkeys_are_external_dependencies() {
        let a = analyse_src(
            "Public Sub DriveDesktop()\n\
             AppActivate Application.Caption\n\
             SendKeys \"{F15}\", True\n\
             End Sub\n",
        );
        assert_eq!(a.class, Some(Class::C));
        assert_eq!(a.api_names.get("AppActivate"), Some(&1));
        assert_eq!(a.api_names.get("SendKeys"), Some(&1));
        assert!(a
            .findings
            .iter()
            .any(|finding| finding.reason.contains("window by title")));
        assert!(a
            .findings
            .iter()
            .any(|finding| finding.reason.contains("injects keystrokes")));
    }

    #[test]
    fn office_dialog_apis_are_user_interface_dependencies() {
        let a = analyse_src(
            "Public Sub PromptUser()\n\
             Set picker = Application.FileDialog(msoFileDialogFilePicker)\n\
             path = Application.GetOpenFilename()\n\
             savePath = Application.GetSaveAsFilename()\n\
             answer = InputBox(\"Value?\")\n\
             MsgBox \"Done\"\n\
             End Sub\n",
        );
        assert_eq!(a.class, Some(Class::D));
        for name in [
            "Application.FileDialog",
            "Application.GetOpenFilename",
            "Application.GetSaveAsFilename",
            "InputBox",
            "MsgBox",
        ] {
            assert_eq!(a.api_names.get(name), Some(&1), "missing {name}");
        }
        assert_eq!(
            a.findings
                .iter()
                .filter(|finding| finding.class == Some(Class::D))
                .count(),
            5
        );
    }

    #[test]
    fn application_status_bar_is_a_user_interface_dependency() {
        let a = analyse_src(
            "Public Function ShowProgress() As String\n\
             Application.StatusBar = \"Oxi 42\"\n\
             ShowProgress = CStr(Application.StatusBar)\n\
             Application.StatusBar = False\n\
             End Function\n",
        );
        assert_eq!(a.metrics.unparsed, 0);
        assert_eq!(a.class, Some(Class::D));
        assert_eq!(a.api_names.get("Application.StatusBar"), Some(&3));
        assert_eq!(
            a.findings
                .iter()
                .filter(|finding| {
                    finding.what.eq_ignore_ascii_case("Application.StatusBar")
                        && finding.class == Some(Class::D)
                })
                .count(),
            3
        );
    }

    #[test]
    fn display_status_bar_is_a_user_interface_dependency() {
        let a = analyse_src(
            "Public Sub ToggleStatusBar()\n\
             Application.DisplayStatusBar = False\n\
             Application.DisplayStatusBar = True\n\
             End Sub\n",
        );
        assert_eq!(a.metrics.unparsed, 0);
        assert_eq!(a.class, Some(Class::D));
        assert_eq!(a.api_names.get("Application.DisplayStatusBar"), Some(&2));
        assert_eq!(
            a.findings
                .iter()
                .filter(|finding| {
                    finding
                        .what
                        .eq_ignore_ascii_case("Application.DisplayStatusBar")
                        && finding.class == Some(Class::D)
                })
                .count(),
            2
        );
    }

    #[test]
    fn display_formula_bar_is_a_user_interface_dependency() {
        let a = analyse_src(
            "Public Sub ExerciseFormulaBar()\n\
             Application.DisplayFormulaBar = False\n\
             Application.DisplayFormulaBar = True\n\
             End Sub\n",
        );
        assert_eq!(a.metrics.unparsed, 0);
        assert_eq!(a.class, Some(Class::D));
        assert_eq!(a.api_names.get("Application.DisplayFormulaBar"), Some(&2));
        assert_eq!(
            a.findings
                .iter()
                .filter(|finding| {
                    finding
                        .what
                        .eq_ignore_ascii_case("Application.DisplayFormulaBar")
                        && finding.reason.contains("formula-bar user interface")
                })
                .count(),
            2
        );
    }

    #[test]
    fn display_full_screen_is_a_user_interface_dependency() {
        let a = analyse_src(
            "Public Sub ExerciseFullScreen()\n\
             Application.DisplayFullScreen = True\n\
             Application.DisplayFullScreen = False\n\
             End Sub\n",
        );
        assert_eq!(a.metrics.unparsed, 0);
        assert_eq!(a.class, Some(Class::D));
        assert_eq!(a.api_names.get("Application.DisplayFullScreen"), Some(&2));
        assert_eq!(
            a.findings
                .iter()
                .filter(|finding| {
                    finding
                        .what
                        .eq_ignore_ascii_case("Application.DisplayFullScreen")
                        && finding.reason.contains("full-screen mode")
                })
                .count(),
            2
        );
    }

    #[test]
    fn application_cursor_is_a_user_interface_dependency() {
        let a = analyse_src(
            "Public Sub ExerciseCursor()\n\
             Application.Cursor = xlWait\n\
             Application.Cursor = xlDefault\n\
             End Sub\n",
        );
        assert_eq!(a.metrics.unparsed, 0);
        assert_eq!(a.class, Some(Class::D));
        assert_eq!(a.api_names.get("Application.Cursor"), Some(&2));
        assert_eq!(
            a.findings
                .iter()
                .filter(|finding| {
                    finding.what.eq_ignore_ascii_case("Application.Cursor")
                        && finding.reason.contains("mouse-pointer user interface")
                })
                .count(),
            2
        );
    }

    #[test]
    fn taskbar_window_visibility_is_a_user_interface_dependency() {
        let a = analyse_src(
            "Public Sub ExerciseTaskbarWindows()\n\
             Application.ShowWindowsInTaskbar = False\n\
             Application.ShowWindowsInTaskbar = True\n\
             End Sub\n",
        );
        assert_eq!(a.metrics.unparsed, 0);
        assert_eq!(a.class, Some(Class::D));
        assert_eq!(
            a.api_names.get("Application.ShowWindowsInTaskbar"),
            Some(&2)
        );
        assert_eq!(
            a.findings
                .iter()
                .filter(|finding| {
                    finding
                        .what
                        .eq_ignore_ascii_case("Application.ShowWindowsInTaskbar")
                        && finding.reason.contains("Windows taskbar")
                })
                .count(),
            2
        );
    }

    #[test]
    fn application_window_state_is_a_user_interface_dependency() {
        let a = analyse_src(
            "Public Sub ExerciseWindowState()\n\
             Application.WindowState = xlMaximized\n\
             Application.WindowState = xlNormal\n\
             End Sub\n",
        );
        assert_eq!(a.metrics.unparsed, 0);
        assert_eq!(a.class, Some(Class::D));
        assert_eq!(a.api_names.get("Application.WindowState"), Some(&2));
        assert_eq!(
            a.findings
                .iter()
                .filter(|finding| {
                    finding.what.eq_ignore_ascii_case("Application.WindowState")
                        && finding.reason.contains("application-window")
                })
                .count(),
            2
        );
    }

    #[test]
    fn active_window_gridlines_are_a_user_interface_dependency() {
        let a = analyse_src(
            "Public Sub ExerciseGridlines()\n\
             Application.ActiveWindow.DisplayGridlines = False\n\
             Application.ActiveWindow.DisplayGridlines = True\n\
             End Sub\n",
        );
        assert_eq!(a.metrics.unparsed, 0);
        assert_eq!(a.class, Some(Class::D));
        assert_eq!(
            a.api_names.get("Application.ActiveWindow.DisplayGridlines"),
            Some(&2)
        );
        assert_eq!(
            a.findings
                .iter()
                .filter(|finding| finding.reason.contains("active UI context"))
                .count(),
            2
        );
        assert_eq!(
            a.findings
                .iter()
                .filter(|finding| finding.reason.contains("worksheet gridlines"))
                .count(),
            2
        );
    }

    #[test]
    fn active_window_headings_are_a_user_interface_dependency() {
        let a = analyse_src(
            "Public Sub ExerciseHeadings()\n\
             Application.ActiveWindow.DisplayHeadings = False\n\
             Application.ActiveWindow.DisplayHeadings = True\n\
             End Sub\n",
        );
        assert_eq!(a.metrics.unparsed, 0);
        assert_eq!(a.class, Some(Class::D));
        assert_eq!(
            a.api_names.get("Application.ActiveWindow.DisplayHeadings"),
            Some(&2)
        );
        assert_eq!(
            a.findings
                .iter()
                .filter(|finding| finding.reason.contains("active UI context"))
                .count(),
            2
        );
        assert_eq!(
            a.findings
                .iter()
                .filter(|finding| finding.reason.contains("row and column headings"))
                .count(),
            2
        );
    }

    #[test]
    fn active_window_workbook_tabs_are_a_user_interface_dependency() {
        let a = analyse_src(
            "Public Sub ExerciseWorkbookTabs()\n\
             Application.ActiveWindow.DisplayWorkbookTabs = False\n\
             Application.ActiveWindow.DisplayWorkbookTabs = True\n\
             End Sub\n",
        );
        assert_eq!(a.metrics.unparsed, 0);
        assert_eq!(a.class, Some(Class::D));
        assert_eq!(
            a.api_names
                .get("Application.ActiveWindow.DisplayWorkbookTabs"),
            Some(&2)
        );
        assert_eq!(
            a.findings
                .iter()
                .filter(|finding| finding.reason.contains("active UI context"))
                .count(),
            2
        );
        assert_eq!(
            a.findings
                .iter()
                .filter(|finding| finding.reason.contains("workbook sheet tabs"))
                .count(),
            2
        );
    }

    #[test]
    fn active_window_scroll_bars_are_user_interface_dependencies() {
        let a = analyse_src(
            "Public Sub ExerciseScrollBars()\n\
             Application.ActiveWindow.DisplayHorizontalScrollBar = False\n\
             Application.ActiveWindow.DisplayVerticalScrollBar = False\n\
             Application.ActiveWindow.DisplayHorizontalScrollBar = True\n\
             Application.ActiveWindow.DisplayVerticalScrollBar = True\n\
             End Sub\n",
        );
        assert_eq!(a.metrics.unparsed, 0);
        assert_eq!(a.class, Some(Class::D));
        assert_eq!(
            a.api_names
                .get("Application.ActiveWindow.DisplayHorizontalScrollBar"),
            Some(&2)
        );
        assert_eq!(
            a.api_names
                .get("Application.ActiveWindow.DisplayVerticalScrollBar"),
            Some(&2)
        );
        assert_eq!(
            a.findings
                .iter()
                .filter(|finding| finding.reason.contains("active UI context"))
                .count(),
            4
        );
        assert_eq!(
            a.findings
                .iter()
                .filter(|finding| finding.reason.contains("scroll bar in an Excel window"))
                .count(),
            4
        );
    }

    #[test]
    fn active_window_zoom_is_a_user_interface_dependency() {
        let a = analyse_src(
            "Public Sub ExerciseZoom()\n\
             ActiveWindow.Zoom = 75\n\
             Application.ActiveWindow.Zoom = 100\n\
             End Sub\n",
        );
        assert_eq!(a.metrics.unparsed, 0);
        assert_eq!(a.class, Some(Class::D));
        assert_eq!(a.api_names.get("ActiveWindow.Zoom"), Some(&1));
        assert_eq!(a.api_names.get("Application.ActiveWindow.Zoom"), Some(&1));
        assert_eq!(
            a.findings
                .iter()
                .filter(|finding| finding.reason.contains("active UI context"))
                .count(),
            2
        );
        assert_eq!(
            a.findings
                .iter()
                .filter(|finding| finding.reason.contains("zoom level"))
                .count(),
            2
        );
    }

    #[test]
    fn active_window_view_mode_is_a_user_interface_dependency() {
        let a = analyse_src(
            "Public Sub ExerciseViewMode()\n\
             ActiveWindow.View = xlPageBreakPreview\n\
             Application.ActiveWindow.View = xlNormalView\n\
             End Sub\n",
        );
        assert_eq!(a.metrics.unparsed, 0);
        assert_eq!(a.class, Some(Class::D));
        assert_eq!(a.api_names.get("ActiveWindow.View"), Some(&1));
        assert_eq!(a.api_names.get("Application.ActiveWindow.View"), Some(&1));
        assert_eq!(
            a.findings
                .iter()
                .filter(|finding| finding.reason.contains("active UI context"))
                .count(),
            2
        );
        assert_eq!(
            a.findings
                .iter()
                .filter(|finding| finding.reason.contains("worksheet view mode"))
                .count(),
            2
        );
    }

    #[test]
    fn active_window_frozen_panes_are_user_interface_dependencies() {
        let a = analyse_src(
            "Public Sub ExerciseFrozenPanes()\n\
             ActiveWindow.SplitRow = 1\n\
             ActiveWindow.SplitColumn = 1\n\
             ActiveWindow.FreezePanes = True\n\
             Application.ActiveWindow.FreezePanes = False\n\
             Application.ActiveWindow.SplitRow = 0\n\
             Application.ActiveWindow.SplitColumn = 0\n\
             End Sub\n",
        );
        assert_eq!(a.metrics.unparsed, 0);
        assert_eq!(a.class, Some(Class::D));
        for name in [
            "ActiveWindow.SplitRow",
            "ActiveWindow.SplitColumn",
            "ActiveWindow.FreezePanes",
            "Application.ActiveWindow.FreezePanes",
            "Application.ActiveWindow.SplitRow",
            "Application.ActiveWindow.SplitColumn",
        ] {
            assert_eq!(a.api_names.get(name), Some(&1), "{name}");
        }
        assert_eq!(
            a.findings
                .iter()
                .filter(|finding| finding.reason.contains("active UI context"))
                .count(),
            6
        );
        assert_eq!(
            a.findings
                .iter()
                .filter(|finding| finding.reason.contains("pane layout"))
                .count(),
            6
        );
    }

    #[test]
    fn active_window_scroll_position_is_a_user_interface_dependency() {
        let a = analyse_src(
            "Public Sub ExerciseScrollPosition()\n\
             ActiveWindow.ScrollRow = 10\n\
             ActiveWindow.ScrollColumn = 5\n\
             Application.ActiveWindow.ScrollRow = 1\n\
             Application.ActiveWindow.ScrollColumn = 1\n\
             End Sub\n",
        );
        assert_eq!(a.metrics.unparsed, 0);
        assert_eq!(a.class, Some(Class::D));
        for name in [
            "ActiveWindow.ScrollRow",
            "ActiveWindow.ScrollColumn",
            "Application.ActiveWindow.ScrollRow",
            "Application.ActiveWindow.ScrollColumn",
        ] {
            assert_eq!(a.api_names.get(name), Some(&1), "{name}");
        }
        assert_eq!(
            a.findings
                .iter()
                .filter(|finding| finding.reason.contains("active UI context"))
                .count(),
            4
        );
        assert_eq!(
            a.findings
                .iter()
                .filter(|finding| finding.reason.contains("first visible"))
                .count(),
            4
        );
    }

    #[test]
    fn doevents_is_reported_as_event_loop_reentrancy() {
        let a = analyse_src(
            "Public Function PumpMessages() As Long\n\
             DoEvents\n\
             PumpMessages = VBA.DoEvents()\n\
             End Function\n",
        );
        assert_eq!(a.api_names.get("DoEvents"), Some(&1));
        assert_eq!(a.api_names.get("VBA.DoEvents"), Some(&1));
        assert_eq!(
            a.findings
                .iter()
                .filter(|finding| finding.reason.contains("reentrant execution"))
                .count(),
            2
        );
        assert!(a
            .findings
            .iter()
            .filter(|finding| finding.reason.contains("reentrant execution"))
            .all(|finding| finding.class.is_none()));
    }

    #[test]
    fn err_raise_is_reported_as_observable_error_semantics() {
        let a = analyse_src(
            "Public Sub FailWithContext()\n\
             Err.Raise 513, \"Probe\", \"boom\"\n\
             End Sub\n",
        );
        assert_eq!(a.api_names.get("Err.Raise"), Some(&1));
        assert!(a.findings.iter().any(|finding| {
            finding.what.eq_ignore_ascii_case("Err.Raise")
                && finding.reason.contains("number, source, and description")
                && finding.class.is_none()
        }));
    }

    #[test]
    fn clock_functions_are_reported_as_time_dependent() {
        let a = analyse_src(
            "Public Function ClockSnapshot() As Variant\n\
             ClockSnapshot = Array(Now, Date, Time, Timer, VBA.Now)\n\
             End Function\n",
        );
        for name in ["Now", "Date", "Time", "Timer", "VBA.Now"] {
            assert_eq!(a.api_names.get(name), Some(&1), "missing {name}");
        }
        assert_eq!(
            a.findings
                .iter()
                .filter(|finding| finding.reason.contains("local time zone"))
                .count(),
            5
        );
        assert!(a
            .findings
            .iter()
            .filter(|finding| finding.reason.contains("local time zone"))
            .all(|finding| finding.class.is_none()));
    }

    #[test]
    fn random_functions_are_reported_as_generator_state_dependent() {
        let a = analyse_src(
            "Public Function Sample() As Single\n\
             Randomize 42\n\
             Sample = Rnd() + VBA.Rnd()\n\
             End Function\n",
        );
        for name in ["Randomize", "Rnd", "VBA.Rnd"] {
            assert_eq!(a.api_names.get(name), Some(&1), "missing {name}");
        }
        assert_eq!(
            a.findings
                .iter()
                .filter(|finding| finding.reason.contains("seed and call order"))
                .count(),
            3
        );
        assert!(a
            .findings
            .iter()
            .filter(|finding| finding.reason.contains("seed and call order"))
            .all(|finding| finding.class.is_none()));
    }

    #[test]
    fn enable_events_is_reported_as_global_excel_state() {
        let a = analyse_src(
            "Public Sub WriteWithoutRecursion()\n\
             Application.EnableEvents = False\n\
             Range(\"A1\").Value2 = 1\n\
             Application.EnableEvents = True\n\
             End Sub\n",
        );
        assert_eq!(a.api_names.get("Application.EnableEvents"), Some(&2));
        assert_eq!(
            a.findings
                .iter()
                .filter(|finding| finding.reason.contains("event delivery state"))
                .count(),
            2
        );
        assert!(a
            .findings
            .iter()
            .filter(|finding| finding.reason.contains("event delivery state"))
            .all(|finding| finding.class.is_none()));
    }

    #[test]
    fn calculation_mode_is_a_global_formula_engine_dependency() {
        let a = analyse_src(
            "Public Sub ToggleCalculation()\n\
             Application.Calculation = xlCalculationManual\n\
             Application.Calculation = xlCalculationAutomatic\n\
             End Sub\n",
        );
        assert!(a.needs_formula_engine);
        assert_eq!(a.api_names.get("Application.Calculation"), Some(&2));
        assert_eq!(
            a.findings
                .iter()
                .filter(|finding| finding.reason.contains("automatic calculation mode"))
                .count(),
            2
        );
        assert!(a
            .findings
            .iter()
            .filter(|finding| finding.reason.contains("automatic calculation mode"))
            .all(|finding| finding.class.is_none()));
    }

    #[test]
    fn display_alerts_is_reported_as_global_excel_state() {
        let a = analyse_src(
            "Public Sub SuppressAlerts()\n\
             Application.DisplayAlerts = False\n\
             Application.DisplayAlerts = True\n\
             End Sub\n",
        );
        assert_eq!(a.api_names.get("Application.DisplayAlerts"), Some(&2));
        assert_eq!(
            a.findings
                .iter()
                .filter(|finding| finding.reason.contains("automatic default responses"))
                .count(),
            2
        );
        assert!(a
            .findings
            .iter()
            .filter(|finding| finding.reason.contains("automatic default responses"))
            .all(|finding| finding.class.is_none()));
    }

    #[test]
    fn automation_security_is_reported_as_global_excel_state() {
        let a = analyse_src(
            "Public Sub SetMacroPolicy()\n\
             Application.AutomationSecurity = msoAutomationSecurityForceDisable\n\
             Application.AutomationSecurity = msoAutomationSecurityLow\n\
             End Sub\n",
        );
        assert_eq!(a.api_names.get("Application.AutomationSecurity"), Some(&2));
        assert_eq!(
            a.findings
                .iter()
                .filter(|finding| finding.reason.contains("programmatically opened files"))
                .count(),
            2
        );
        assert!(a
            .findings
            .iter()
            .filter(|finding| finding.reason.contains("programmatically opened files"))
            .all(|finding| finding.class.is_none()));
    }

    #[test]
    fn screen_updating_is_reported_as_global_excel_state() {
        let a = analyse_src(
            "Public Sub ToggleRedraw()\n\
             Application.ScreenUpdating = False\n\
             Application.ScreenUpdating = True\n\
             End Sub\n",
        );
        assert_eq!(a.api_names.get("Application.ScreenUpdating"), Some(&2));
        assert_eq!(
            a.findings
                .iter()
                .filter(|finding| finding.reason.contains("screen redraw state"))
                .count(),
            2
        );
        assert!(a
            .findings
            .iter()
            .filter(|finding| finding.reason.contains("screen redraw state"))
            .all(|finding| finding.class.is_none()));
    }

    #[test]
    fn ask_to_update_links_is_reported_as_global_excel_state() {
        let a = analyse_src(
            "Public Sub ToggleLinkPrompts()\n\
             Application.AskToUpdateLinks = False\n\
             Application.AskToUpdateLinks = True\n\
             End Sub\n",
        );
        assert_eq!(a.api_names.get("Application.AskToUpdateLinks"), Some(&2));
        assert_eq!(
            a.findings
                .iter()
                .filter(|finding| finding.reason.contains("prompt policy"))
                .count(),
            2
        );
        assert!(a
            .findings
            .iter()
            .filter(|finding| finding.reason.contains("prompt policy"))
            .all(|finding| finding.class.is_none()));
    }

    #[test]
    fn interactive_is_reported_as_global_excel_state() {
        let a = analyse_src(
            "Public Sub ToggleUserInput()\n\
             Application.Interactive = False\n\
             Application.Interactive = True\n\
             End Sub\n",
        );
        assert_eq!(a.api_names.get("Application.Interactive"), Some(&2));
        assert_eq!(
            a.findings
                .iter()
                .filter(|finding| finding.reason.contains("keyboard and mouse input state"))
                .count(),
            2
        );
        assert!(a
            .findings
            .iter()
            .filter(|finding| finding.reason.contains("keyboard and mouse input state"))
            .all(|finding| finding.class.is_none()));
    }

    #[test]
    fn opening_an_external_workbook_is_class_c() {
        let a = analyse_src(
            "Public Function ReadExternalWorkbook(ByVal path As String) As Double\n\
             Dim opened As Workbook\n\
             Set opened = Application.Workbooks.Open(Filename:=path, ReadOnly:=True)\n\
             ReadExternalWorkbook = opened.Worksheets(1).Range(\"A1\").Value2\n\
             opened.Close SaveChanges:=False\n\
             End Function\n",
        );
        assert_eq!(a.metrics.unparsed, 0);
        assert_eq!(a.class, Some(Class::C));
        assert_eq!(a.api_names.get("Application.Workbooks.Open"), Some(&1));
        assert!(a.findings.iter().any(|finding| {
            finding
                .what
                .eq_ignore_ascii_case("Application.Workbooks.Open")
                && finding.reason.contains("external workbook")
                && finding.class == Some(Class::C)
        }));
    }

    #[test]
    fn managing_external_workbook_links_is_class_c() {
        let a = analyse_src(
            "Public Sub ManageLinks()\n\
             Dim links As Variant\n\
             links = ThisWorkbook.LinkSources(xlExcelLinks)\n\
             ThisWorkbook.UpdateLink Name:=links(1), Type:=xlExcelLinks\n\
             ThisWorkbook.ChangeLink Name:=links(1), NewName:=\"next.xlsx\", Type:=xlExcelLinks\n\
             ThisWorkbook.BreakLink Name:=links(1), Type:=xlLinkTypeExcelLinks\n\
             End Sub\n",
        );
        assert_eq!(a.metrics.unparsed, 0);
        assert_eq!(a.class, Some(Class::C));
        for api in [
            "ThisWorkbook.LinkSources",
            "ThisWorkbook.UpdateLink",
            "ThisWorkbook.ChangeLink",
            "ThisWorkbook.BreakLink",
        ] {
            assert_eq!(a.api_names.get(api), Some(&1), "missing {api}");
            assert!(a.findings.iter().any(|finding| {
                finding.what.eq_ignore_ascii_case(api)
                    && finding.reason.contains("external workbook link")
                    && finding.class == Some(Class::C)
            }));
        }
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
