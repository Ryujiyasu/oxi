// This Source Code Form is subject to the terms of the Mozilla Public
// License, v. 2.0. If a copy of the MPL was not distributed with this
// file, You can obtain one at https://mozilla.org/MPL/2.0/.

//! Structural fingerprints, for finding which macros are copies of which.
//!
//! # Why not hash the file
//!
//! An `.xlsm` is a zip: opening and saving it changes every byte. Hashing the
//! extracted source is not much better, because the copies that matter differ —
//! that is what makes them copies rather than duplicates. A file whose only
//! change is the fiscal year in a caption has to come out equal.
//!
//! So the hash is taken over the syntax tree with the parts that carry no
//! meaning removed, and how much counts as "no meaning" is adjustable
//! ([`Normalization`]).
//!
//! # Why per procedure
//!
//! Copies are made and then edited: eight of ten procedures survive untouched,
//! two are changed, one is added. Comparing whole files calls that "different".
//! Comparing procedures gives the overlap, and the overlap is the answer.
//!
//! # What is never normalised away
//!
//! - **Procedure names.** Two procedures with the same name and different
//!   bodies is the most dangerous thing a corpus can contain, and erasing names
//!   would hide it.
//! - **Member and API names.** `.Value` and `.Formula` are not interchangeable,
//!   and the classification depends on them.

use std::collections::{BTreeMap, BTreeSet};
use std::fmt::Write as _;

use crate::ast::*;

/// Which parts of the source are treated as carrying no meaning.
///
/// Deliberately three independent switches rather than one dial. "Ignore
/// literals" is safe for a caption and unsafe for a tax rate, and whoever runs
/// the tool is the one who knows which they are looking at.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct Normalization {
    /// Rename locals to `v1`, `v2`, ... in order of first appearance, so that
    /// renaming `i` to `cnt` does not make a copy look original.
    pub rename_locals: bool,
    /// Replace literal values with their type. Absorbs the year, the department
    /// name, the file path — and also the tax rate, which is why this is off by
    /// default at `Standard`.
    pub erase_literals: bool,
    /// Rename procedures too, to catch a copy that was renamed wholesale.
    pub rename_procedures: bool,
}

/// Named presets. Report which one produced a result, always.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum Strength {
    /// Structure, names and values all significant.
    Strict,
    /// Local variable names ignored.
    Standard,
    /// Local names and literal values ignored.
    Loose,
    /// Also ignores procedure names.
    Loosest,
}

impl Strength {
    pub fn normalization(self) -> Normalization {
        match self {
            Strength::Strict => Normalization {
                rename_locals: false,
                erase_literals: false,
                rename_procedures: false,
            },
            Strength::Standard => Normalization {
                rename_locals: true,
                erase_literals: false,
                rename_procedures: false,
            },
            Strength::Loose => Normalization {
                rename_locals: true,
                erase_literals: true,
                rename_procedures: false,
            },
            Strength::Loosest => Normalization {
                rename_locals: true,
                erase_literals: true,
                rename_procedures: true,
            },
        }
    }

    pub fn as_str(self) -> &'static str {
        match self {
            Strength::Strict => "strict",
            Strength::Standard => "standard",
            Strength::Loose => "loose",
            Strength::Loosest => "loosest",
        }
    }
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct ProcedureFingerprint {
    pub name: String,
    pub hash: u128,
    pub statements: usize,
    pub line: u32,
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct ModuleFingerprint {
    pub procedures: Vec<ProcedureFingerprint>,
    /// Order-independent hash of the whole module, so that moving a procedure
    /// does not change it.
    pub combined: u128,
    /// Hash of declarations that affect exact module identity without changing
    /// every procedure body (Declare, module variables, Type, Enum, events).
    pub declarations: u128,
    pub strength: Strength,
}

impl ModuleFingerprint {
    fn hashes(&self) -> BTreeSet<u128> {
        self.procedures.iter().map(|p| p.hash).collect()
    }
}

pub fn fingerprint_module(module: &Module, strength: Strength) -> ModuleFingerprint {
    let norm = strength.normalization();
    let mut procedures = Vec::new();
    let module_context = canonical_procedure_context(module);
    let declarations = canonical_module_declarations(module, norm);

    for item in &module.items {
        if let ModuleItem::Procedure(proc) = item {
            let procedure = canonical_procedure(proc, norm);
            let canonical = if module_context.is_empty() {
                procedure
            } else {
                format!("{module_context}{procedure}")
            };
            procedures.push(ProcedureFingerprint {
                name: proc.name.clone(),
                hash: hash128(&canonical),
                statements: count_statements(&proc.body),
                line: proc.span.line,
            });
        }
    }

    // XOR folds order away; the count is mixed in so that a module with a
    // procedure duplicated is not equal to one with it appearing once.
    let mut combined = 0u128;
    for p in &procedures {
        combined ^= p.hash;
    }
    combined ^= hash128(&format!("#{}", procedures.len()));
    if !module_context.is_empty() {
        combined ^= hash128(&module_context);
    }
    let declarations = if declarations.is_empty() {
        0
    } else {
        hash128(&declarations)
    };
    combined ^= declarations;

    ModuleFingerprint {
        procedures,
        combined,
        declarations,
        strength,
    }
}

/// Render one procedure to the canonical text its hash is taken over.
///
/// Exposed because a fingerprint that disagrees is useless without being able
/// to see *what* it compared.
pub fn canonical_procedure(proc: &Procedure, norm: Normalization) -> String {
    let mut locals = LocalNames::new(norm);
    if norm.rename_locals {
        locals.collect_from_procedure(proc);
    }

    let mut out = String::new();
    let name = if norm.rename_procedures {
        "proc".to_string()
    } else {
        proc.name.to_ascii_lowercase()
    };
    let _ = write!(
        out,
        "{:?}{} {:?} {} (",
        proc.visibility,
        if proc.is_static { " static" } else { "" },
        proc.kind,
        name
    );
    for param in &proc.params {
        render_param(param, &mut locals, &mut out);
        out.push(';');
    }
    out.push(')');
    if let Some(return_type) = &proc.return_type {
        out.push_str(" as ");
        render_type_name(return_type, &mut locals, &mut out);
    }
    out.push('\n');
    render_body(&proc.body, &mut locals, 1, &mut out);
    out
}

fn render_param(param: &Param, locals: &mut LocalNames, out: &mut String) {
    let _ = write!(
        out,
        "{:?}{}{} {}:",
        param.mode,
        if param.optional { "?" } else { "" },
        if param.is_array { "[]" } else { "" },
        locals.render(&param.name)
    );
    render_type_name(&param.type_name, locals, out);
    if let Some(default) = &param.default {
        out.push('=');
        render_expr(default, locals, out);
    }
}

fn render_type_name(type_name: &TypeName, locals: &mut LocalNames, out: &mut String) {
    out.push_str(&type_name.name.to_ascii_lowercase());
    if let Some(length) = &type_name.fixed_length {
        out.push('*');
        render_expr(length, locals, out);
    }
}

fn render_array_bounds(
    bounds: &Option<Vec<ArrayBound>>,
    locals: &mut LocalNames,
    out: &mut String,
) {
    let Some(bounds) = bounds else {
        return;
    };
    out.push('[');
    for bound in bounds {
        if let Some(lower) = &bound.lower {
            render_expr(lower, locals, out);
            out.push_str(" to ");
        }
        render_expr(&bound.upper, locals, out);
        out.push(';');
    }
    out.push(']');
}

fn count_statements(body: &[Statement]) -> usize {
    let mut n = 0;
    for stmt in body {
        n += 1;
        for nested in child_bodies(stmt) {
            n += count_statements(nested);
        }
    }
    n
}

fn child_bodies(stmt: &Statement) -> Vec<&[Statement]> {
    match stmt {
        Statement::If(s) => {
            let mut v: Vec<&[Statement]> = vec![&s.then_body];
            v.extend(s.else_ifs.iter().map(|(_, b)| b.as_slice()));
            if let Some(b) = &s.else_body {
                v.push(b);
            }
            v
        }
        Statement::SelectCase(s) => {
            let mut v: Vec<&[Statement]> = s.cases.iter().map(|c| c.body.as_slice()).collect();
            if let Some(b) = &s.case_else {
                v.push(b);
            }
            v
        }
        Statement::For(s) => vec![&s.body],
        Statement::ForEach(s) => vec![&s.body],
        Statement::Do(s) => vec![&s.body],
        Statement::While { body, .. } | Statement::With { body, .. } => vec![body],
        _ => Vec::new(),
    }
}

/// Maps declared local names onto `v1`, `v2`, ... in order of first appearance.
///
/// Only names *declared inside the procedure* are renamed. Anything else is a
/// module-level variable, a call, or an API name, and those carry meaning.
struct LocalNames {
    norm: Normalization,
    declared: BTreeSet<String>,
    assigned: BTreeMap<String, String>,
    next: usize,
}

impl LocalNames {
    fn new(norm: Normalization) -> LocalNames {
        LocalNames {
            norm,
            declared: BTreeSet::new(),
            assigned: BTreeMap::new(),
            next: 1,
        }
    }

    fn collect_from_procedure(&mut self, proc: &Procedure) {
        for param in &proc.params {
            self.declared.insert(param.name.to_ascii_lowercase());
        }
        self.collect_body(&proc.body);
    }

    fn collect_body(&mut self, body: &[Statement]) {
        for stmt in body {
            match stmt {
                Statement::Dim(decl) => {
                    for item in &decl.items {
                        self.declared.insert(item.name.to_ascii_lowercase());
                    }
                }
                Statement::ReDim { items, .. } => {
                    for item in items {
                        self.declared.insert(item.name.to_ascii_lowercase());
                    }
                }
                Statement::For(s) => {
                    if let Some(name) = s.counter.dotted_name() {
                        self.declared.insert(name.to_ascii_lowercase());
                    }
                }
                Statement::ForEach(s) => {
                    if let Some(name) = s.item.dotted_name() {
                        self.declared.insert(name.to_ascii_lowercase());
                    }
                }
                _ => {}
            }
            for nested in child_bodies(stmt) {
                self.collect_body(nested);
            }
        }
    }

    fn render(&mut self, name: &str) -> String {
        let lower = name.to_ascii_lowercase();
        if !self.norm.rename_locals || !self.declared.contains(&lower) {
            return lower;
        }
        if let Some(existing) = self.assigned.get(&lower) {
            return existing.clone();
        }
        let placeholder = format!("v{}", self.next);
        self.next += 1;
        self.assigned.insert(lower, placeholder.clone());
        placeholder
    }
}

fn indent(depth: usize, out: &mut String) {
    for _ in 0..depth {
        out.push(' ');
    }
}

fn render_body(body: &[Statement], locals: &mut LocalNames, depth: usize, out: &mut String) {
    for stmt in body {
        render_statement(stmt, locals, depth, out);
    }
}

fn render_statement(stmt: &Statement, locals: &mut LocalNames, depth: usize, out: &mut String) {
    // Comments carry no behaviour. Dropping them is the whole point of
    // comparing trees rather than text, and they have to be skipped before
    // anything is written or the indentation alone would make copies differ.
    if matches!(stmt, Statement::Comment { .. }) {
        return;
    }

    indent(depth, out);
    match stmt {
        Statement::Comment { .. } => unreachable!("handled above"),
        Statement::Assign { target, value, .. } => {
            out.push_str("let ");
            render_expr(target, locals, out);
            out.push('=');
            render_expr(value, locals, out);
        }
        Statement::SetAssign { target, value, .. } => {
            out.push_str("set ");
            render_expr(target, locals, out);
            out.push('=');
            render_expr(value, locals, out);
        }
        Statement::MidAssign(mid) => {
            out.push_str("mid ");
            render_expr(&mid.target, locals, out);
            out.push(',');
            render_expr(&mid.start, locals, out);
            if let Some(length) = &mid.length {
                out.push(',');
                render_expr(length, locals, out);
            }
            out.push('=');
            render_expr(&mid.value, locals, out);
        }
        Statement::AlignedAssign(aligned) => {
            out.push_str(match aligned.kind {
                AlignmentKind::Left => "lset ",
                AlignmentKind::Right => "rset ",
            });
            render_expr(&aligned.target, locals, out);
            out.push('=');
            render_expr(&aligned.value, locals, out);
        }
        Statement::Call { target, .. } => {
            // `Call Foo(x)` and `Foo x` are the same call; the spelling is not
            // part of the structure.
            out.push_str("call ");
            render_expr(target, locals, out);
        }
        Statement::Dim(decl) => {
            let _ = write!(
                out,
                "dim{}{}",
                if decl.is_static { " static" } else { "" },
                if decl.is_const { " const" } else { "" }
            );
            for item in &decl.items {
                let name = locals.render(&item.name);
                let _ = write!(out, " {name}:");
                render_type_name(&item.type_name, locals, out);
                render_array_bounds(&item.array_bounds, locals, out);
                if let Some(value) = &item.value {
                    out.push('=');
                    render_expr(value, locals, out);
                }
            }
        }
        Statement::ReDim {
            preserve, items, ..
        } => {
            let _ = write!(out, "redim{}", if *preserve { " preserve" } else { "" });
            for item in items {
                let name = locals.render(&item.name);
                let _ = write!(out, " {name}:");
                render_type_name(&item.type_name, locals, out);
                render_array_bounds(&item.array_bounds, locals, out);
            }
        }
        Statement::Erase { targets, .. } => {
            out.push_str("erase");
            for t in targets {
                out.push(' ');
                render_expr(t, locals, out);
            }
        }
        Statement::Open(open) => {
            out.push_str("open ");
            render_expr(&open.path, locals, out);
            let _ = write!(out, " for {:?}", open.mode);
            if let Some(access) = open.access {
                let _ = write!(out, " access {access:?}");
            }
            if let Some(lock) = open.lock {
                let _ = write!(out, " lock {lock:?}");
            }
            out.push_str(" as #");
            render_expr(&open.file_number, locals, out);
            if let Some(record_len) = &open.record_len {
                out.push_str(" len=");
                render_expr(record_len, locals, out);
            }
        }
        Statement::Close { files, .. } => {
            out.push_str("close");
            for file in files {
                out.push_str(" #");
                render_expr(file, locals, out);
            }
        }
        Statement::FileOutput(output) => {
            let keyword = match output.kind {
                FileOutputKind::Print => "print",
                FileOutputKind::Write => "write",
            };
            let _ = write!(out, "{keyword} #");
            render_expr(&output.file_number, locals, out);
            for item in &output.items {
                out.push(match item.separator {
                    FileOutputSeparator::Comma => ',',
                    FileOutputSeparator::Semicolon => ';',
                });
                if let Some(value) = &item.value {
                    render_expr(value, locals, out);
                }
            }
        }
        Statement::FileInput(input) => {
            out.push_str(if input.line {
                "line input #"
            } else {
                "input #"
            });
            render_expr(&input.file_number, locals, out);
            for target in &input.targets {
                out.push(',');
                render_expr(target, locals, out);
            }
        }
        Statement::FileTransfer(transfer) => {
            out.push_str(match transfer.kind {
                FileTransferKind::Get => "get #",
                FileTransferKind::Put => "put #",
            });
            render_expr(&transfer.file_number, locals, out);
            out.push(',');
            if let Some(record_number) = &transfer.record_number {
                render_expr(record_number, locals, out);
            }
            out.push(',');
            render_expr(&transfer.value, locals, out);
        }
        Statement::FileSeek(seek) => {
            out.push_str("seek #");
            render_expr(&seek.file_number, locals, out);
            out.push(',');
            render_expr(&seek.position, locals, out);
        }
        Statement::FileSystem(operation) => match operation {
            FileSystemStmt::Rename {
                source,
                destination,
                ..
            } => {
                out.push_str("name ");
                render_expr(source, locals, out);
                out.push_str(" as ");
                render_expr(destination, locals, out);
            }
            FileSystemStmt::Copy {
                source,
                destination,
                ..
            } => {
                out.push_str("filecopy ");
                render_expr(source, locals, out);
                out.push(',');
                render_expr(destination, locals, out);
            }
            FileSystemStmt::Unary { kind, path, .. } => {
                let _ = write!(out, "{kind:?} ");
                render_expr(path, locals, out);
            }
            FileSystemStmt::SetAttr {
                path, attributes, ..
            } => {
                out.push_str("setattr ");
                render_expr(path, locals, out);
                out.push(',');
                render_expr(attributes, locals, out);
            }
        },
        Statement::FileRecordLock(lock) => {
            out.push_str(match lock.kind {
                FileRecordLockKind::Lock => "lock #",
                FileRecordLockKind::Unlock => "unlock #",
            });
            render_expr(&lock.file_number, locals, out);
            if lock.start.is_some() || lock.end.is_some() {
                out.push(',');
                if let Some(start) = &lock.start {
                    render_expr(start, locals, out);
                }
                if let Some(end) = &lock.end {
                    out.push_str(" to ");
                    render_expr(end, locals, out);
                }
            }
        }
        Statement::FileWidth {
            file_number, width, ..
        } => {
            out.push_str("width #");
            render_expr(file_number, locals, out);
            out.push(',');
            render_expr(width, locals, out);
        }
        Statement::FileReset { .. } => out.push_str("reset"),
        Statement::If(s) => {
            out.push_str("if ");
            render_expr(&s.condition, locals, out);
            out.push('\n');
            render_body(&s.then_body, locals, depth + 1, out);
            for (cond, body) in &s.else_ifs {
                indent(depth, out);
                out.push_str("elseif ");
                render_expr(cond, locals, out);
                out.push('\n');
                render_body(body, locals, depth + 1, out);
            }
            if let Some(body) = &s.else_body {
                indent(depth, out);
                out.push_str("else\n");
                render_body(body, locals, depth + 1, out);
            }
            indent(depth, out);
            out.push_str("endif");
        }
        Statement::SelectCase(s) => {
            out.push_str("select ");
            render_expr(&s.subject, locals, out);
            out.push('\n');
            for case in &s.cases {
                indent(depth, out);
                out.push_str("case");
                for label in &case.labels {
                    out.push(' ');
                    match label {
                        CaseLabel::Value(e) => render_expr(e, locals, out),
                        CaseLabel::Range(a, b) => {
                            render_expr(a, locals, out);
                            out.push_str("..");
                            render_expr(b, locals, out);
                        }
                        CaseLabel::Compare(op, e) => {
                            let _ = write!(out, "{op:?}");
                            render_expr(e, locals, out);
                        }
                    }
                }
                out.push('\n');
                render_body(&case.body, locals, depth + 1, out);
            }
            if let Some(body) = &s.case_else {
                indent(depth, out);
                out.push_str("caseelse\n");
                render_body(body, locals, depth + 1, out);
            }
            indent(depth, out);
            out.push_str("endselect");
        }
        Statement::For(s) => {
            out.push_str("for ");
            render_expr(&s.counter, locals, out);
            out.push('=');
            render_expr(&s.from, locals, out);
            out.push_str("..");
            render_expr(&s.to, locals, out);
            if let Some(step) = &s.step {
                out.push_str(" step ");
                render_expr(step, locals, out);
            }
            out.push('\n');
            render_body(&s.body, locals, depth + 1, out);
            indent(depth, out);
            out.push_str("next");
        }
        Statement::ForEach(s) => {
            out.push_str("foreach ");
            render_expr(&s.item, locals, out);
            out.push_str(" in ");
            render_expr(&s.collection, locals, out);
            out.push('\n');
            render_body(&s.body, locals, depth + 1, out);
            indent(depth, out);
            out.push_str("next");
        }
        Statement::Do(s) => {
            out.push_str("do");
            for (label, test) in [("pre", &s.pre), ("post", &s.post)] {
                if let Some(test) = test {
                    let _ = write!(
                        out,
                        " {label}{}",
                        if test.until { "until" } else { "while" }
                    );
                    render_expr(&test.condition, locals, out);
                }
            }
            out.push('\n');
            render_body(&s.body, locals, depth + 1, out);
            indent(depth, out);
            out.push_str("loop");
        }
        Statement::While {
            condition, body, ..
        } => {
            out.push_str("while ");
            render_expr(condition, locals, out);
            out.push('\n');
            render_body(body, locals, depth + 1, out);
            indent(depth, out);
            out.push_str("wend");
        }
        Statement::With { subject, body, .. } => {
            out.push_str("with ");
            render_expr(subject, locals, out);
            out.push('\n');
            render_body(body, locals, depth + 1, out);
            indent(depth, out);
            out.push_str("endwith");
        }
        Statement::OnError(kind) => {
            let text = match kind {
                OnError::Goto { label, .. } => {
                    format!("onerror goto {}", label.to_ascii_lowercase())
                }
                OnError::Disable { .. } => "onerror off".to_string(),
                OnError::ResumeNext { .. } => "onerror resumenext".to_string(),
            };
            out.push_str(&text);
        }
        Statement::OnBranch(branch) => {
            out.push_str("on ");
            render_expr(&branch.selector, locals, out);
            out.push_str(match branch.kind {
                OnBranchKind::GoTo => " goto",
                OnBranchKind::GoSub => " gosub",
            });
            for label in &branch.labels {
                out.push(' ');
                out.push_str(&label.to_ascii_lowercase());
            }
        }
        Statement::RaiseEvent(event) => {
            out.push_str("raiseevent ");
            out.push_str(&event.name.to_ascii_lowercase());
            out.push('(');
            for (index, arg) in event.args.iter().enumerate() {
                if index > 0 {
                    out.push(',');
                }
                if let Some(name) = &arg.name {
                    out.push_str(&name.to_ascii_lowercase());
                    out.push_str(":=");
                }
                if let Some(value) = &arg.value {
                    render_expr(value, locals, out);
                }
            }
            out.push(')');
        }
        Statement::Resume { target, .. } => {
            let _ = write!(out, "resume {target:?}");
        }
        Statement::GoTo { label, .. } => {
            let _ = write!(out, "goto {}", label.to_ascii_lowercase());
        }
        Statement::GoSub { label, .. } => {
            let _ = write!(out, "gosub {}", label.to_ascii_lowercase());
        }
        Statement::Return { .. } => out.push_str("return"),
        Statement::Exit { what, .. } => {
            let _ = write!(out, "exit {what:?}");
        }
        Statement::Label { name, .. } => {
            let _ = write!(out, "label {}", name.to_ascii_lowercase());
        }
        Statement::LineNumber { value, .. } => {
            let _ = write!(out, "label {value}");
        }
        Statement::End { .. } => out.push_str("end"),
        Statement::Stop { .. } => out.push_str("stop"),
        Statement::Directive { text, .. } => {
            let _ = write!(out, "directive {}", text.to_ascii_lowercase());
        }
        // Kept in the hash: two files differ if one has a line the parser could
        // not read and the other does not.
        Statement::Unknown { text, .. } => {
            let _ = write!(out, "unparsed {}", text.to_ascii_lowercase());
        }
    }
    out.push('\n');
}

fn render_expr(expr: &Expr, locals: &mut LocalNames, out: &mut String) {
    match expr {
        Expr::Literal(lit, _) => render_literal(lit, locals.norm, out),
        Expr::Ident(name, _) => out.push_str(&locals.render(name)),
        Expr::WithMember(name, _) => {
            out.push('.');
            out.push_str(&name.to_ascii_lowercase());
        }
        Expr::Member { object, name, .. } => {
            render_expr(object, locals, out);
            out.push('.');
            // Member names are API surface; never renamed.
            out.push_str(&name.to_ascii_lowercase());
        }
        Expr::Index { target, args, .. } => {
            render_expr(target, locals, out);
            out.push('(');
            for arg in args {
                if let Some(name) = &arg.name {
                    let _ = write!(out, "{}:=", name.to_ascii_lowercase());
                }
                match &arg.value {
                    Some(value) => render_expr(value, locals, out),
                    None => out.push('_'),
                }
                out.push(',');
            }
            out.push(')');
        }
        Expr::Bang { object, name, .. } => {
            render_expr(object, locals, out);
            out.push('!');
            out.push_str(&name.to_ascii_lowercase());
        }
        Expr::New { type_name, .. } => {
            let _ = write!(out, "new {}", type_name.to_ascii_lowercase());
        }
        Expr::AddressOf { procedure, .. } => {
            let _ = write!(out, "addressof {}", procedure.to_ascii_lowercase());
        }
        Expr::Unary { op, operand, .. } => {
            let _ = write!(out, "{op:?}(");
            render_expr(operand, locals, out);
            out.push(')');
        }
        Expr::Binary { op, lhs, rhs, .. } => {
            let _ = write!(out, "{op:?}(");
            render_expr(lhs, locals, out);
            out.push(',');
            render_expr(rhs, locals, out);
            out.push(')');
        }
        Expr::TypeOf {
            operand, type_name, ..
        } => {
            out.push_str("typeof(");
            render_expr(operand, locals, out);
            let _ = write!(out, ",{})", type_name.to_ascii_lowercase());
        }
    }
}

fn render_literal(lit: &Literal, norm: Normalization, out: &mut String) {
    if norm.erase_literals {
        let tag = match lit {
            Literal::Number(_) => "<num>",
            Literal::Str(_) => "<str>",
            Literal::Date(_) => "<date>",
            Literal::Bool(_) => "<bool>",
            Literal::Empty => "<empty>",
            Literal::Null => "<null>",
            Literal::Nothing => "<nothing>",
        };
        out.push_str(tag);
        return;
    }
    match lit {
        Literal::Number(n) => {
            let _ = write!(out, "{n}");
        }
        Literal::Str(s) => {
            let _ = write!(out, "{s:?}");
        }
        Literal::Date(s) => {
            let _ = write!(out, "#{s}#");
        }
        Literal::Bool(b) => {
            let _ = write!(out, "{b}");
        }
        Literal::Empty => out.push_str("empty"),
        Literal::Null => out.push_str("null"),
        Literal::Nothing => out.push_str("nothing"),
    }
}

/// How two modules relate.
#[derive(Debug, Clone, PartialEq)]
pub struct Similarity {
    pub shared: usize,
    pub only_a: usize,
    pub only_b: usize,
    /// Shared / union. `1.0` means every procedure matched.
    pub jaccard: f64,
    /// Same procedure name, different body.
    ///
    /// The most important thing this comparison finds: two files that look
    /// interchangeable but are not, which is exactly how a copy quietly
    /// diverges from its original.
    pub diverged: Vec<String>,
    /// Module-level declarations differ even if individual procedures match.
    pub declarations_differ: bool,
}

pub fn compare(a: &ModuleFingerprint, b: &ModuleFingerprint) -> Similarity {
    let (ha, hb) = (a.hashes(), b.hashes());
    let shared = ha.intersection(&hb).count();
    let union = ha.union(&hb).count();

    let by_name_a: BTreeMap<String, u128> = a
        .procedures
        .iter()
        .map(|p| (p.name.to_ascii_lowercase(), p.hash))
        .collect();
    let mut diverged = Vec::new();
    for p in &b.procedures {
        if let Some(other) = by_name_a.get(&p.name.to_ascii_lowercase()) {
            if *other != p.hash {
                diverged.push(p.name.clone());
            }
        }
    }

    Similarity {
        shared,
        only_a: ha.len() - shared,
        only_b: hb.len() - shared,
        jaccard: if union == 0 {
            1.0
        } else {
            shared as f64 / union as f64
        },
        diverged,
        declarations_differ: a.declarations != b.declarations,
    }
}

/// FNV-1a, run twice with different offsets to make a 128-bit value.
///
/// A 64-bit hash would collide somewhere around a few billion procedures by the
/// birthday bound, which is comfortably beyond any corpus. Widening it anyway is
/// nearly free, and a tool that says "these 47 files are the same" should not be
/// wrong because of arithmetic.
fn hash128(text: &str) -> u128 {
    const PRIME: u64 = 0x0000_0100_0000_01b3;
    let mut lo: u64 = 0xcbf2_9ce4_8422_2325;
    let mut hi: u64 = 0x9dcf_16f7_0d8c_5b21;
    for byte in text.as_bytes() {
        lo = (lo ^ *byte as u64).wrapping_mul(PRIME);
        hi = (hi ^ (*byte as u64).rotate_left(17)).wrapping_mul(PRIME);
    }
    ((hi as u128) << 64) | lo as u128
}

fn canonical_procedure_context(module: &Module) -> String {
    let mut entries = Vec::new();
    for item in &module.items {
        match item {
            ModuleItem::Option(option, _) => {
                let text = match option {
                    ModuleOption::Explicit => "option explicit".to_string(),
                    ModuleOption::Base(base) => format!("option base {base}"),
                    ModuleOption::Compare(mode) => {
                        format!("option compare {}", mode.to_ascii_lowercase())
                    }
                    ModuleOption::PrivateModule => "option private module".to_string(),
                };
                entries.push(text);
            }
            ModuleItem::DefType(def) => {
                for range in &def.ranges {
                    entries.push(format!(
                        "deftype {} {}-{}",
                        def.type_name.to_ascii_lowercase(),
                        range.start,
                        range.end
                    ));
                }
            }
            _ => {}
        }
    }
    entries.sort_unstable();
    entries.dedup();

    let mut out = String::new();
    for entry in entries {
        let _ = writeln!(out, "{entry}");
    }
    out
}

fn canonical_module_declarations(module: &Module, norm: Normalization) -> String {
    let mut entries = Vec::new();
    let module_norm = Normalization {
        rename_locals: false,
        ..norm
    };
    for item in &module.items {
        match item {
            ModuleItem::Option(..) | ModuleItem::DefType(_) => {}
            ModuleItem::Variables(decl) => {
                let mut text = format!(
                    "modulevar {:?}{}{}",
                    decl.visibility,
                    if decl.is_static { " static" } else { "" },
                    if decl.is_const { " const" } else { "" }
                );
                let mut locals = LocalNames::new(module_norm);
                for item in &decl.items {
                    text.push(' ');
                    render_module_var_item(item, &mut locals, &mut text);
                }
                entries.push(text);
            }
            ModuleItem::Type(def) => {
                let mut text = format!(
                    "type {:?} {}",
                    def.visibility,
                    def.name.to_ascii_lowercase()
                );
                let mut locals = LocalNames::new(module_norm);
                for field in &def.fields {
                    text.push(' ');
                    render_module_var_item(field, &mut locals, &mut text);
                }
                entries.push(text);
            }
            ModuleItem::Enum(def) => {
                let mut text = format!(
                    "enum {:?} {}",
                    def.visibility,
                    def.name.to_ascii_lowercase()
                );
                let mut locals = LocalNames::new(module_norm);
                for (name, value) in &def.members {
                    let _ = write!(text, " {}", name.to_ascii_lowercase());
                    if let Some(value) = value {
                        text.push('=');
                        render_expr(value, &mut locals, &mut text);
                    }
                }
                entries.push(text);
            }
            ModuleItem::ExternalProc(proc) => {
                let mut text = format!(
                    "declare {:?}{} {} {} lib ",
                    proc.visibility,
                    if proc.ptr_safe { " ptrsafe" } else { "" },
                    if proc.is_function { "function" } else { "sub" },
                    proc.name.to_ascii_lowercase()
                );
                render_literal(&Literal::Str(proc.lib.clone()), module_norm, &mut text);
                if let Some(alias) = &proc.alias {
                    text.push_str(" alias ");
                    render_literal(&Literal::Str(alias.clone()), module_norm, &mut text);
                }
                text.push('(');
                let mut locals = LocalNames::new(module_norm);
                for param in &proc.params {
                    render_param(param, &mut locals, &mut text);
                    text.push(';');
                }
                text.push(')');
                if let Some(return_type) = &proc.return_type {
                    text.push_str(" as ");
                    render_type_name(return_type, &mut locals, &mut text);
                }
                entries.push(text);
            }
            ModuleItem::Implements { interface, .. } => {
                entries.push(format!("implements {}", interface.to_ascii_lowercase()));
            }
            ModuleItem::Event { name, params, .. } => {
                let mut text = format!("event {}(", name.to_ascii_lowercase());
                let mut locals = LocalNames::new(module_norm);
                for param in params {
                    render_param(param, &mut locals, &mut text);
                    text.push(';');
                }
                text.push(')');
                entries.push(text);
            }
            ModuleItem::Unknown { text, .. } if !text.trim_start().starts_with('\'') => {
                entries.push(format!("unparsed module {}", text.to_ascii_lowercase()));
            }
            ModuleItem::Attribute { .. }
            | ModuleItem::Procedure(_)
            | ModuleItem::Unknown { .. } => {}
        }
    }
    entries.sort_unstable();
    entries.dedup();

    let mut out = String::new();
    for entry in entries {
        let _ = writeln!(out, "{entry}");
    }
    out
}

fn render_module_var_item(item: &VarItem, locals: &mut LocalNames, out: &mut String) {
    if item.with_events {
        out.push_str("withevents ");
    }
    let _ = write!(out, "{}:", item.name.to_ascii_lowercase());
    render_type_name(&item.type_name, locals, out);
    render_array_bounds(&item.array_bounds, locals, out);
    if let Some(value) = &item.value {
        out.push('=');
        render_expr(value, locals, out);
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::parse_module;

    fn fp(src: &str, strength: Strength) -> ModuleFingerprint {
        fingerprint_module(&parse_module(src).expect("should parse"), strength)
    }

    const ORIGINAL: &str = "\
Sub Summarise()
    Dim i As Long
    For i = 1 To 10
        Range(\"A\" & i).Value = i * 2
    Next i
End Sub";

    #[test]
    fn comments_and_layout_never_matter() {
        let commented = "\
Sub Summarise()
    ' count the rows
    Dim i As Long
    For i = 1 To 10
        ' double it
        Range(\"A\" & i).Value = i * 2
    Next i
End Sub";
        assert_eq!(
            fp(ORIGINAL, Strength::Strict).combined,
            fp(commented, Strength::Strict).combined
        );
    }

    #[test]
    fn renaming_a_variable_is_invisible_at_standard_but_not_at_strict() {
        let renamed = ORIGINAL
            .replace(" i ", " cnt ")
            .replace("(i", "(cnt")
            .replace("i *", "cnt *")
            .replace("& i", "& cnt");
        assert_ne!(
            fp(ORIGINAL, Strength::Strict).combined,
            fp(&renamed, Strength::Strict).combined
        );
        assert_eq!(
            fp(ORIGINAL, Strength::Standard).combined,
            fp(&renamed, Strength::Standard).combined
        );
    }

    #[test]
    fn changing_a_literal_shows_at_standard_and_hides_at_loose() {
        let next_year = ORIGINAL.replace("To 10", "To 20");
        assert_ne!(
            fp(ORIGINAL, Strength::Standard).combined,
            fp(&next_year, Strength::Standard).combined
        );
        assert_eq!(
            fp(ORIGINAL, Strength::Loose).combined,
            fp(&next_year, Strength::Loose).combined
        );
    }

    #[test]
    fn a_renamed_procedure_only_matches_at_loosest() {
        let renamed = ORIGINAL.replace("Summarise", "Aggregate");
        assert_ne!(
            fp(ORIGINAL, Strength::Loose).combined,
            fp(&renamed, Strength::Loose).combined
        );
        assert_eq!(
            fp(ORIGINAL, Strength::Loosest).combined,
            fp(&renamed, Strength::Loosest).combined
        );
    }

    #[test]
    fn member_names_are_never_normalised_away() {
        // .Value and .Formula are not interchangeable, at any strength.
        let formula = ORIGINAL.replace(".Value", ".Formula");
        assert_ne!(
            fp(ORIGINAL, Strength::Loosest).combined,
            fp(&formula, Strength::Loosest).combined
        );
    }

    #[test]
    fn reordering_procedures_does_not_change_the_module() {
        let a = "Sub One()\nx = 1\nEnd Sub\nSub Two()\ny = 2\nEnd Sub";
        let b = "Sub Two()\ny = 2\nEnd Sub\nSub One()\nx = 1\nEnd Sub";
        assert_eq!(
            fp(a, Strength::Standard).combined,
            fp(b, Strength::Standard).combined
        );
    }

    #[test]
    fn call_spelling_is_not_structure() {
        let bare = "Sub T()\nFoo 1, 2\nEnd Sub";
        let explicit = "Sub T()\nCall Foo(1, 2)\nEnd Sub";
        assert_eq!(
            fp(bare, Strength::Standard).combined,
            fp(explicit, Strength::Standard).combined
        );
    }

    #[test]
    fn partial_edits_show_up_as_partial_overlap() {
        // Eight untouched, one edited, one added: the shape of a real copy.
        let original = "\
Sub A()\nx = 1\nEnd Sub
Sub B()\nx = 2\nEnd Sub
Sub C()\nx = 3\nEnd Sub";
        let variant = "\
Sub A()\nx = 1\nEnd Sub
Sub B()\nx = 99\nEnd Sub
Sub C()\nx = 3\nEnd Sub
Sub D()\nx = 4\nEnd Sub";

        let s = compare(
            &fp(original, Strength::Standard),
            &fp(variant, Strength::Standard),
        );
        assert_eq!(s.shared, 2);
        assert_eq!(s.only_a, 1);
        assert_eq!(s.only_b, 2);
        assert!(s.jaccard > 0.3 && s.jaccard < 0.5);
    }

    #[test]
    fn same_name_different_body_is_called_out() {
        let a = "Sub Calc()\nx = 1\nEnd Sub";
        let b = "Sub Calc()\nx = 2\nEnd Sub";
        let s = compare(&fp(a, Strength::Standard), &fp(b, Strength::Standard));
        assert_eq!(s.diverged, vec!["Calc".to_string()]);
        assert_eq!(s.shared, 0);
    }

    #[test]
    fn identical_modules_compare_as_identical() {
        let s = compare(
            &fp(ORIGINAL, Strength::Standard),
            &fp(ORIGINAL, Strength::Standard),
        );
        assert_eq!(s.jaccard, 1.0);
        assert!(s.diverged.is_empty());
        assert_eq!(s.only_a, 0);
    }

    #[test]
    fn the_canonical_form_is_inspectable() {
        // A fingerprint that disagrees is useless without seeing what it read.
        let module = parse_module(ORIGINAL).unwrap();
        let proc = module
            .items
            .iter()
            .find_map(|i| match i {
                ModuleItem::Procedure(p) => Some(p),
                _ => None,
            })
            .unwrap();
        let text = canonical_procedure(proc, Strength::Standard.normalization());
        assert!(text.contains("v1"), "locals should be renamed:\n{text}");
        assert!(text.contains("range"), "API names should survive:\n{text}");
        assert!(!text.contains("count the rows"));
    }

    #[test]
    fn unparsed_lines_still_count_as_a_difference() {
        let clean = "Sub T()\nx = 1\nEnd Sub";
        let with_io = "Sub T()\nGet #1\nx = 1\nEnd Sub";
        assert_ne!(
            fp(clean, Strength::Loosest).combined,
            fp(with_io, Strength::Loosest).combined
        );
    }

    #[test]
    fn file_io_shape_and_print_separators_are_fingerprinted() {
        let comma = "Sub T()\nPrint #1, a, b\nEnd Sub";
        let semicolon = "Sub T()\nPrint #1, a; b\nEnd Sub";
        let write = "Sub T()\nWrite #1, a, b\nEnd Sub";
        assert_ne!(
            fp(comma, Strength::Loosest).combined,
            fp(semicolon, Strength::Loosest).combined
        );
        assert_ne!(
            fp(comma, Strength::Loosest).combined,
            fp(write, Strength::Loosest).combined
        );

        let get_next = "Sub T()\nGet #1, , value\nEnd Sub";
        let get_at = "Sub T()\nGet #1, 10, value\nEnd Sub";
        let put_at = "Sub T()\nPut #1, 10, value\nEnd Sub";
        assert_ne!(
            fp(get_next, Strength::Loosest).combined,
            fp(get_at, Strength::Loosest).combined
        );
        assert_ne!(
            fp(get_at, Strength::Loosest).combined,
            fp(put_at, Strength::Loosest).combined
        );
    }

    #[test]
    fn addressof_callback_name_is_fingerprinted() {
        let a = fp(
            "Sub Hook()\nx = SetTimer(0, 0, 1, AddressOf TimerA)\nEnd Sub",
            Strength::Standard,
        );
        let b = fp(
            "Sub Hook()\nx = SetTimer(0, 0, 1, AddressOf TimerB)\nEnd Sub",
            Strength::Standard,
        );
        assert_ne!(a.combined, b.combined);
    }

    #[test]
    fn filesystem_operation_kinds_are_fingerprinted() {
        let rename = "Sub T()\nName sourcePath As destinationPath\nEnd Sub";
        let copy = "Sub T()\nFileCopy sourcePath, destinationPath\nEnd Sub";
        let kill = "Sub T()\nKill sourcePath\nEnd Sub";
        let mkdir = "Sub T()\nMkDir sourcePath\nEnd Sub";
        assert_ne!(
            fp(rename, Strength::Loosest).combined,
            fp(copy, Strength::Loosest).combined
        );
        assert_ne!(
            fp(kill, Strength::Loosest).combined,
            fp(mkdir, Strength::Loosest).combined
        );
    }

    #[test]
    fn file_lock_ranges_and_width_are_fingerprinted() {
        let whole = "Sub T()\nLock #1\nEnd Sub";
        let single = "Sub T()\nLock #1, 2\nEnd Sub";
        let range = "Sub T()\nLock #1, 1 To 4\nEnd Sub";
        let unlock = "Sub T()\nUnlock #1, 1 To 4\nEnd Sub";
        let width = "Sub T()\nWidth #1, 4\nEnd Sub";
        for different in [single, range, unlock, width] {
            assert_ne!(
                fp(whole, Strength::Loosest).combined,
                fp(different, Strength::Loosest).combined
            );
        }
    }

    #[test]
    fn computed_branch_kind_and_label_order_are_fingerprinted() {
        let goto = "Sub T()\nOn choice GoTo First, Second\nEnd Sub";
        let gosub = "Sub T()\nOn choice GoSub First, Second\nEnd Sub";
        let reordered = "Sub T()\nOn choice GoTo Second, First\nEnd Sub";
        assert_ne!(
            fp(goto, Strength::Loosest).combined,
            fp(gosub, Strength::Loosest).combined
        );
        assert_ne!(
            fp(goto, Strength::Loosest).combined,
            fp(reordered, Strength::Loosest).combined
        );
    }

    #[test]
    fn string_assignment_kind_and_mid_length_are_fingerprinted() {
        let mid_short = "Sub T()\nMid$(value, 2) = replacement\nEnd Sub";
        let mid_bounded = "Sub T()\nMid$(value, 2, 3) = replacement\nEnd Sub";
        let lset = "Sub T()\nLSet value = replacement\nEnd Sub";
        let rset = "Sub T()\nRSet value = replacement\nEnd Sub";
        assert_ne!(
            fp(mid_short, Strength::Loosest).combined,
            fp(mid_bounded, Strength::Loosest).combined
        );
        assert_ne!(
            fp(lset, Strength::Loosest).combined,
            fp(rset, Strength::Loosest).combined
        );
    }

    #[test]
    fn raised_event_name_and_arguments_are_fingerprinted() {
        let fired = "Sub T()\nRaiseEvent Fired(value)\nEnd Sub";
        let ping = "Sub T()\nRaiseEvent Ping(value)\nEnd Sub";
        let no_args = "Sub T()\nRaiseEvent Fired\nEnd Sub";
        assert_ne!(
            fp(fired, Strength::Loosest).combined,
            fp(ping, Strength::Loosest).combined
        );
        assert_ne!(
            fp(fired, Strength::Loosest).combined,
            fp(no_args, Strength::Loosest).combined
        );
    }

    #[test]
    fn deftype_module_context_changes_procedure_and_combined_hashes() {
        let integer = "DefInt A-C\nSub T()\nDim apple\nEnd Sub";
        let string = "DefStr A-C\nSub T()\nDim apple\nEnd Sub";
        let integer_fp = fp(integer, Strength::Loosest);
        let string_fp = fp(string, Strength::Loosest);
        assert_ne!(integer_fp.combined, string_fp.combined);
        assert_ne!(integer_fp.procedures[0].hash, string_fp.procedures[0].hash);
    }

    #[test]
    fn equivalent_deftype_range_order_has_the_same_fingerprint() {
        let first = "DefInt A-C, X-Z\nSub T()\nEnd Sub";
        let second = "DefInt X-Z, A-C\nSub T()\nEnd Sub";
        assert_eq!(
            fp(first, Strength::Loosest).combined,
            fp(second, Strength::Loosest).combined
        );
    }

    #[test]
    fn module_options_change_procedure_and_combined_hashes() {
        let base_zero = "Option Base 0\nSub T()\nDim values(3)\nEnd Sub";
        let base_one = "Option Base 1\nSub T()\nDim values(3)\nEnd Sub";
        let binary = "Option Compare Binary\nSub T()\nx = \"a\" = \"A\"\nEnd Sub";
        let text = "Option Compare Text\nSub T()\nx = \"a\" = \"A\"\nEnd Sub";
        for (left, right) in [(base_zero, base_one), (binary, text)] {
            let left = fp(left, Strength::Loosest);
            let right = fp(right, Strength::Loosest);
            assert_ne!(left.combined, right.combined);
            assert_ne!(left.procedures[0].hash, right.procedures[0].hash);
        }
    }

    #[test]
    fn procedure_signatures_are_fingerprinted() {
        let base =
            "Private Function Convert(Optional ByVal value As Long = 1) As Long\nEnd Function";
        for different in [
            "Private Function Convert(Optional ByVal value As String = 1) As Long\nEnd Function",
            "Private Function Convert(Optional ByVal value As Long = 1) As String\nEnd Function",
            "Public Function Convert(Optional ByVal value As Long = 1) As Long\nEnd Function",
            "Private Static Function Convert(Optional ByVal value As Long = 1) As Long\nEnd Function",
        ] {
            assert_ne!(
                fp(base, Strength::Standard).combined,
                fp(different, Strength::Standard).combined
            );
        }

        let changed_default =
            "Private Function Convert(Optional ByVal value As Long = 2) As Long\nEnd Function";
        assert_ne!(
            fp(base, Strength::Standard).combined,
            fp(changed_default, Strength::Standard).combined
        );
        assert_eq!(
            fp(base, Strength::Loose).combined,
            fp(changed_default, Strength::Loose).combined
        );
    }

    #[test]
    fn type_suffixes_match_equivalent_as_types() {
        let suffix = "Private Function Convert$(ByVal value&)\nDim ratio!\nEnd Function";
        let explicit = "Private Function Convert(ByVal value As Long) As String\n\
                        Dim ratio As Single\n\
                        End Function";
        let different = "Private Function Convert#(ByVal value%)\nDim ratio@\nEnd Function";
        assert_eq!(
            fp(suffix, Strength::Standard).combined,
            fp(explicit, Strength::Standard).combined
        );
        assert_ne!(
            fp(suffix, Strength::Standard).combined,
            fp(different, Strength::Standard).combined
        );
    }

    #[test]
    fn fixed_string_lengths_follow_literal_normalization() {
        let four = "Sub T()\nDim value As String * 4\nEnd Sub";
        let eight = "Sub T()\nDim value As String * 8\nEnd Sub";
        assert_ne!(
            fp(four, Strength::Standard).combined,
            fp(eight, Strength::Standard).combined
        );
        assert_eq!(
            fp(four, Strength::Loose).combined,
            fp(eight, Strength::Loose).combined
        );
    }

    #[test]
    fn local_array_shapes_and_static_storage_are_fingerprinted() {
        let scalar = "Sub T()\nDim values As Long\nEnd Sub";
        let dynamic = "Sub T()\nDim values() As Long\nEnd Sub";
        let bounded = "Sub T()\nDim values(1 To 4) As Long\nEnd Sub";
        let larger = "Sub T()\nDim values(1 To 8) As Long\nEnd Sub";
        let static_local = "Sub T()\nStatic values As Long\nEnd Sub";
        for different in [dynamic, bounded, larger, static_local] {
            assert_ne!(
                fp(scalar, Strength::Standard).combined,
                fp(different, Strength::Standard).combined
            );
        }
        assert_ne!(
            fp(bounded, Strength::Standard).combined,
            fp(larger, Strength::Standard).combined
        );
    }

    #[test]
    fn redim_bounds_and_element_types_are_fingerprinted() {
        let base = "Sub T()\nReDim values(1 To 4) As Long\nEnd Sub";
        let larger = "Sub T()\nReDim values(1 To 8) As Long\nEnd Sub";
        let strings = "Sub T()\nReDim values(1 To 4) As String\nEnd Sub";
        for different in [larger, strings] {
            assert_ne!(
                fp(base, Strength::Standard).combined,
                fp(different, Strength::Standard).combined
            );
        }
    }

    #[test]
    fn module_declarations_are_fingerprinted() {
        let procedure = "\nSub T()\nEnd Sub";
        for (left, right) in [
            (
                "Private Const Limit As Long = 1",
                "Private Const Limit As Long = 2",
            ),
            (
                "Private Declare PtrSafe Sub Sleep Lib \"kernel32\" (ByVal ms As Long)",
                "Private Declare PtrSafe Sub Sleep Lib \"user32\" (ByVal ms As Long)",
            ),
            (
                "Private Type Pair\nLeft As Long\nEnd Type",
                "Private Type Pair\nLeft As String\nEnd Type",
            ),
            (
                "Private Enum State\nReady = 1\nEnd Enum",
                "Private Enum State\nReady = 2\nEnd Enum",
            ),
            ("Implements First.Contract", "Implements Second.Contract"),
            (
                "Public Event Ready(ByVal value As Long)",
                "Public Event Ready(ByVal value As String)",
            ),
        ] {
            let left = fp(&format!("{left}{procedure}"), Strength::Standard);
            let right = fp(&format!("{right}{procedure}"), Strength::Standard);
            assert_ne!(
                left.combined, right.combined,
                "declaration difference was erased"
            );
            assert_eq!(
                left.procedures[0].hash, right.procedures[0].hash,
                "module declarations should not erase procedure overlap"
            );
            assert!(compare(&left, &right).declarations_differ);
        }
    }

    #[test]
    fn module_comments_do_not_change_fingerprints() {
        let plain = "Option Explicit\nSub T()\nEnd Sub";
        let commented = "' module documentation\nOption Explicit\nSub T()\nEnd Sub";
        assert_eq!(
            fp(plain, Strength::Strict).combined,
            fp(commented, Strength::Strict).combined
        );
    }

    #[test]
    fn equivalent_module_option_order_has_the_same_fingerprint() {
        let first = "Option Explicit\nOption Base 1\nOption Compare Text\nSub T()\nEnd Sub";
        let second = "Option Compare Text\nOption Explicit\nOption Base 1\nSub T()\nEnd Sub";
        assert_eq!(
            fp(first, Strength::Loosest).combined,
            fp(second, Strength::Loosest).combined
        );
    }
}
