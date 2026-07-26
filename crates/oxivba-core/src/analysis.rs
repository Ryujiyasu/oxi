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
    Rule { pattern: "UserForm",   how: Match::RootPrefix, class: Class::D, reason: "drives a UserForm; the target platform supplies its own UI" },
    Rule { pattern: "MSForms",    how: Match::Root,       class: Class::D, reason: "uses MSForms controls" },
    Rule { pattern: "Load",       how: Match::Exact,   class: Class::D, reason: "loads a form" },
    Rule { pattern: "Unload",     how: Match::Exact,   class: Class::D, reason: "unloads a form" },

    // -- C: leaves Excel ---------------------------------------------------
    Rule { pattern: "Shell",           how: Match::Exact,   class: Class::C, reason: "starts an external process" },
    Rule { pattern: "CreateObject",    how: Match::Exact,   class: Class::C, reason: "late-binds an external COM object; the target cannot be determined statically" },
    Rule { pattern: "GetObject",       how: Match::Exact,   class: Class::C, reason: "attaches to an external COM object" },
    Rule { pattern: "Scripting.FileSystemObject", how: Match::Exact, class: Class::C, reason: "manipulates the file system" },
    Rule { pattern: "ADODB",           how: Match::Root,    class: Class::C, reason: "connects to a database" },
    Rule { pattern: "DAO",             how: Match::Root,    class: Class::C, reason: "connects to a database" },
    Rule { pattern: "Outlook",         how: Match::Root,    class: Class::C, reason: "drives another Office application" },
    Rule { pattern: "Word",            how: Match::Root,    class: Class::C, reason: "drives another Office application" },
    Rule { pattern: "PowerPoint",      how: Match::Root,    class: Class::C, reason: "drives another Office application" },
    Rule { pattern: "WScript",         how: Match::Root,    class: Class::C, reason: "uses Windows Script Host" },
    Rule { pattern: "SendKeys",        how: Match::Exact,   class: Class::C, reason: "synthesises keystrokes" },
    Rule { pattern: "Kill",            how: Match::Exact,   class: Class::C, reason: "deletes a file" },
    Rule { pattern: "MkDir",           how: Match::Exact,   class: Class::C, reason: "creates a directory" },
    Rule { pattern: "RmDir",           how: Match::Exact,   class: Class::C, reason: "removes a directory" },
    Rule { pattern: "FileCopy",        how: Match::Exact,   class: Class::C, reason: "copies a file" },
    Rule { pattern: "Dir",             how: Match::Exact,   class: Class::C, reason: "enumerates the file system" },
    Rule { pattern: "Environ",         how: Match::Exact,   class: Class::C, reason: "reads the process environment" },
    Rule { pattern: "SaveAs",          how: Match::Segment, class: Class::C, reason: "writes a file to a path" },
    Rule { pattern: "OpenText",        how: Match::Segment, class: Class::C, reason: "reads an external file" },
    Rule { pattern: "QueryTables",     how: Match::Segment, class: Class::C, reason: "pulls data from an external source" },

    // -- A: report generation ---------------------------------------------
    Rule { pattern: "PrintOut",      how: Match::Segment, class: Class::A, reason: "prints" },
    Rule { pattern: "PageSetup",     how: Match::Segment, class: Class::A, reason: "configures a printed page" },
    Rule { pattern: "Interior",      how: Match::Segment, class: Class::A, reason: "sets cell fill" },
    Rule { pattern: "Font",          how: Match::Segment, class: Class::A, reason: "sets fonts" },
    Rule { pattern: "Borders",       how: Match::Segment, class: Class::A, reason: "sets borders" },
    Rule { pattern: "NumberFormat",  how: Match::Segment, class: Class::A, reason: "sets number formats" },
    Rule { pattern: "MergeCells",    how: Match::Segment, class: Class::A, reason: "merges cells" },
    Rule { pattern: "ColumnWidth",   how: Match::Segment, class: Class::A, reason: "adjusts layout" },
    Rule { pattern: "RowHeight",     how: Match::Segment, class: Class::A, reason: "adjusts layout" },

    // -- B: data transformation --------------------------------------------
    Rule { pattern: "Range",   how: Match::Segment, class: Class::B, reason: "reads or writes cells" },
    Rule { pattern: "Cells",   how: Match::Segment, class: Class::B, reason: "reads or writes cells" },
    Rule { pattern: "Value",   how: Match::Segment, class: Class::B, reason: "reads or writes cell values" },
    Rule { pattern: "Value2",  how: Match::Segment, class: Class::B, reason: "reads or writes cell values" },
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

    // Per-procedure scratch.
    current_calls: BTreeSet<String>,
    current_statements: usize,
    current_depth: usize,
    current_max_depth: usize,
}

impl Walker {
    fn walk_module(&mut self, module: &Module) {
        for item in &module.items {
            match item {
                ModuleItem::Option(ModuleOption::Explicit, _) => self.has_option_explicit = true,
                ModuleItem::Option(..) => {}
                ModuleItem::Attribute { .. } | ModuleItem::Implements { .. } => {}
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
                        if let Some(default) = &param.default {
                            self.walk_expr(default);
                        }
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
                ModuleItem::Event { params, .. } => {
                    for param in params {
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
        self.current_calls = BTreeSet::new();
        self.current_statements = 0;
        self.current_depth = 0;
        self.current_max_depth = 0;

        for param in &proc.params {
            if let Some(default) = &param.default {
                self.walk_expr(default);
            }
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
            Statement::GoTo { label, .. } | Statement::GoSub { label, .. } => {
                self.current_calls.insert(label.clone());
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
        self.record_name(&item.type_name.name, 0);
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
        *self.api_names.entry(name.to_string()).or_default() += 1;

        if let Some(root) = name.split('.').next() {
            self.current_calls.insert(root.to_string());
        }
        self.current_calls.insert(name.to_string());

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

    fn finish(mut self) -> Analysis {
        let class = self
            .findings
            .iter()
            .filter_map(|f| f.class)
            .max_by_key(|c| c.severity());

        let called: BTreeSet<String> = self
            .procedures
            .iter()
            .flat_map(|p| p.calls.iter().cloned())
            .collect();

        let uncalled = self
            .defined_procedures
            .iter()
            .filter(|(name, visibility, kind)| {
                !called.contains(name)
                    && !is_externally_reachable(name, *visibility, *kind)
            })
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
            Match::RootPrefix => root.len() >= rule.pattern.len()
                && root[..rule.pattern.len()].eq_ignore_ascii_case(rule.pattern),
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

/// Collect the longest dotted name of each member chain, so that
/// `Application.WorksheetFunction.Sum` is recorded once rather than three times.
fn collect_names(expr: &Expr, out: &mut Vec<(String, u32)>) {
    match expr {
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
        let a = analyse_src(
            "Sub Total()\n  Range(\"A1\").Value = Range(\"A2\").Value + 1\nEnd Sub",
        );
        assert_eq!(a.class, Some(Class::B));
        assert!(!a.needs_formula_engine);
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
    fn declare_statements_are_flagged_without_being_called() {
        let a = analyse_src(
            "Private Declare PtrSafe Function Sleep Lib \"kernel32\" (ByVal ms As Long) As Long\n",
        );
        assert_eq!(a.class, Some(Class::C));
        assert_eq!(a.external_declares, vec!["Sleep".to_string()]);
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
        let a = analyse_src("Sub T()\n  Open \"f.txt\" For Input As #1\n  x = 1\nEnd Sub");
        assert_eq!(a.metrics.unparsed, 1);
        assert!(a.findings.iter().any(|f| f.reason.contains("not understood")));
    }

    #[test]
    fn option_explicit_is_noticed() {
        assert!(analyse_src("Option Explicit\nSub T()\nEnd Sub").has_option_explicit);
        assert!(!analyse_src("Sub T()\nEnd Sub").has_option_explicit);
    }

    #[test]
    fn a_module_with_no_excel_api_is_unclassified_rather_than_guessed() {
        let a = analyse_src("Sub T()\n  x = 1 + 2\nEnd Sub");
        assert_eq!(a.class, None);
        assert!(a.verdict().contains("unclassified"));
    }
}
