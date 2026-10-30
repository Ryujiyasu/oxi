// This Source Code Form is subject to the terms of the Mozilla Public
// License, v. 2.0. If a copy of the MPL was not distributed with this
// file, You can obtain one at https://mozilla.org/MPL/2.0/.

//! Formula syntax tree.

use crate::reference::Reference;
use crate::value::Value;

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum UnaryOp {
    /// Unary minus. Binds tighter than `^` in Excel, so `-2^2` is `4`.
    Neg,
    /// Unary plus preserves the operand, including text and logical values.
    Plus,
    /// Postfix `%`.
    Percent,
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum BinaryOp {
    Add,
    Sub,
    Mul,
    Div,
    Pow,
    Concat,
    Eq,
    Ne,
    Lt,
    Le,
    Gt,
    Ge,
}

impl BinaryOp {
    pub fn is_comparison(self) -> bool {
        matches!(
            self,
            BinaryOp::Eq | BinaryOp::Ne | BinaryOp::Lt | BinaryOp::Le | BinaryOp::Gt | BinaryOp::Ge
        )
    }
}

#[derive(Debug, Clone, PartialEq)]
pub enum Expr {
    Literal(Value),
    /// A cell or range reference, possibly sheet-qualified.
    Ref(Reference),
    /// A defined name. Resolution is deferred to evaluation time.
    Name(String),
    /// A table's column, named rather than pointed at:
    /// `tblNomina[[#This Row],[DATE]]`. Resolved against the workbook's tables
    /// when it is evaluated, since which cells it means depends on where the
    /// table is and — for `#This Row` — on which cell is asking.
    Table { name: String, asked: String },
    Unary {
        op: UnaryOp,
        operand: Box<Expr>,
    },
    Binary {
        op: BinaryOp,
        lhs: Box<Expr>,
        rhs: Box<Expr>,
    },
    Function {
        /// Always upper-cased; Excel function names are case-insensitive.
        name: String,
        args: Vec<Expr>,
    },
    /// An array written out where a range could go: `{1,2;3,4}`. Rows first,
    /// each the same width -- Excel refuses a ragged one outright. Only
    /// constants may appear inside, which is why this holds values rather
    /// than expressions.
    Array(Vec<Vec<Value>>),
}

impl Expr {
    /// Instantiate references in a defined-name body at its caller. Names
    /// inside the body remain symbolic and instantiate their own definitions.
    pub fn for_defined_name(&self, at: (u32, u32)) -> Self {
        match self {
            Self::Ref(reference) => {
                let mut reference = reference.clone();
                reference.range = reference.range.for_defined_name(at);
                Self::Ref(reference)
            }
            Self::Unary { op, operand } => Self::Unary { op: *op, operand: Box::new(operand.for_defined_name(at)) },
            Self::Binary { op, lhs, rhs } => Self::Binary { op: *op, lhs: Box::new(lhs.for_defined_name(at)), rhs: Box::new(rhs.for_defined_name(at)) },
            Self::Function { name, args } => Self::Function { name: name.clone(), args: args.iter().map(|a| a.for_defined_name(at)).collect() },
            _ => self.clone(),
        }
    }
}

// An alias retains two dependency views without copying its syntax tree:
// consuming its values, and obtaining the address of its result.
#[derive(Clone, Default)]
struct ReferenceReads {
    values: Vec<Reference>,
    addresses: Vec<Reference>,
}

impl ReferenceReads {
    fn append(to: &mut Vec<Reference>, from: &[Reference]) {
        for reference in from {
            if !to.contains(reference) { to.push(reference.clone()); }
        }
    }

    fn value(values: Vec<Reference>) -> Self {
        Self { addresses: values.clone(), values }
    }
}

impl Expr {
    /// Walk every node in the tree, parents before children.
    pub fn visit(&self, f: &mut impl FnMut(&Expr)) {
        f(self);
        match self {
            Expr::Unary { operand, .. } => operand.visit(f),
            Expr::Binary { lhs, rhs, .. } => {
                lhs.visit(f);
                rhs.visit(f);
            }
            // A table reference names its cells rather than pointing at them,
            // so it has no children to walk.
            Expr::Table { .. } => {}
            Expr::Function { args, .. } => {
                for arg in args {
                    arg.visit(f);
                }
            }
            // An array holds constants, so like a table reference it has no
            // children to walk.
            Expr::Literal(_) | Expr::Ref(_) | Expr::Name(_) | Expr::Array(_) => {}
        }
    }

    /// Whether this function is asking about a range rather than reading it.
    ///
    /// These four want a shape or a position, so the range they name is not
    /// something they wait for. Only a plain reference counts: anything that
    /// has to be worked out first is a dependency like any other, and the
    /// evaluator draws the line in the same place.
    pub fn asks_only_the_shape(name: &str, args: &[Expr]) -> bool {
        matches!(name, "ROW" | "COLUMN" | "ROWS" | "COLUMNS")
            && matches!(args.first(), Some(Expr::Ref(_)) | Some(Expr::Table { .. }))
    }

    fn metadata_reference_argument(name: &str, args: &[Expr]) -> Option<usize> {
        if matches!(name, "ROW" | "COLUMN" | "ROWS" | "COLUMNS" | "ISREF" | "AREAS" | "SHEET" | "SHEETS" | "ISFORMULA") && args.len() == 1 { return Some(0); }
        if name == "CELL" && args.len() == 2
            && matches!(&args[0], Expr::Literal(Value::Text(info)) if ["row", "col", "address"].iter().any(|which| info.eq_ignore_ascii_case(which))) { return Some(1); }
        None
    }

    /// Whether an expression can supply a reference argument in a stored formula.
    /// User functions are resolved at calculation time, so their return type
    /// cannot be rejected while the formula is being entered.
    pub fn can_return_reference(&self) -> bool {
        match self {
            Expr::Ref(_) | Expr::Table { .. } | Expr::Name(_) => true,
            Expr::Function { name, .. } => matches!(name.as_str(),
                "IF" | "CHOOSE" | "IFS" | "SWITCH" | "LET" | "_INVOKE" |
                "INDEX" | "OFFSET" | "INDIRECT" | "TAKE" | "DROP" | "XLOOKUP" |
                "TRIMRANGE" | "_UNION" | "_ISECT" | "_SPILL") || !crate::functions::is_known_function(name),
            _ => false,
        }
    }

    // Selectors consume values; result branches and range sources can supply
    // addresses. Summaries keep repeated aliases bounded by distinct references.
    fn walk_reference_values(&self, found: &mut Vec<Reference>) {
        ReferenceReads::append(found, &self.reference_reads(&[]).addresses);
    }

    fn reference_reads(&self, aliases: &[(String, ReferenceReads)]) -> ReferenceReads {
        let append = ReferenceReads::append;
        match self {
            Expr::Ref(reference) => ReferenceReads { values: vec![reference.clone()], addresses: Vec::new() },
            Expr::Name(label) => aliases.iter().rev().find(|(held, _)| held.eq_ignore_ascii_case(label)).map(|(_, reads)| reads.clone()).unwrap_or_default(),
            Expr::Unary { operand, .. } => ReferenceReads::value(operand.reference_reads(aliases).values),
            Expr::Binary { lhs, rhs, .. } => {
                let mut values = lhs.reference_reads(aliases).values;
                append(&mut values, &rhs.reference_reads(aliases).values);
                ReferenceReads::value(values)
            },
            Expr::Function { name, args } if name == "LET" && args.len() >= 3 && args.len() % 2 == 1 => {
                let mut scope = aliases.to_vec();
                let mut result = ReferenceReads::default();
                for pair in args[..args.len()-1].chunks_exact(2) {
                    let Expr::Name(label) = &pair[0] else { return ReferenceReads::value(self.references()); };
                    // The initializer still sees the outer binding. Shadow it
                    // only after both dependency views have been collected.
                    let bound = pair[1].reference_reads(&scope);
                    append(&mut result.values, &bound.values);
                    append(&mut result.addresses, &bound.addresses);
                    scope.push((label.clone(), bound));
                }
                let body = args.last().unwrap().reference_reads(&scope);
                append(&mut result.values, &body.values);
                append(&mut result.addresses, &body.addresses);
                result
            },
            Expr::Function { name, args } => {
                if name == "_INVOKE" && !args.is_empty() {
                    if let Expr::Function { name: lambda, args: body } = &args[0] {
                        if lambda == "LAMBDA" && body.len() == args.len() && !body.is_empty() {
                            let mut scope = aliases.to_vec();
                            let mut result = ReferenceReads::default();
                            for (label, value) in body[..body.len()-1].iter().zip(&args[1..]) {
                                let Expr::Name(label) = label else { return ReferenceReads::value(self.references()); };
                                let bound = value.reference_reads(aliases);
                                append(&mut result.values, &bound.values);
                                append(&mut result.addresses, &bound.addresses);
                                scope.push((label.clone(), bound));
                            }
                            let body = body.last().unwrap().reference_reads(&scope);
                            append(&mut result.values, &body.values);
                            append(&mut result.addresses, &body.addresses);
                            return result;
                        }
                    }
                }
                let reads: Vec<_> = args.iter().map(|arg| arg.reference_reads(aliases)).collect();
                let mut values = Vec::new();
                for read in &reads { append(&mut values, &read.values); }
                if name == "INDEX" {
                    if let Some(selected) = Self::indexed_references(args) {
                        values = selected;
                        for read in reads.iter().skip(1) { append(&mut values, &read.values); }
                    }
                }
                if name == "SUMIF" && args.len() == 3 {
                    if let (Expr::Ref(criteria), Expr::Ref(sum)) = (&args[0], &args[2]) {
                        let right = sum.range.start.col.checked_add(criteria.range.width() - 1);
                        let bottom = sum.range.start.row.checked_add(criteria.range.height() - 1);
                        if let (Some(right), Some(bottom)) = (right, bottom) {
                            if right <= crate::reference::MAX_COL && bottom <= crate::reference::MAX_ROW {
                                values = reads[0].values.clone();
                                append(&mut values, &reads[1].values);
                                let mut sum = sum.clone();
                                sum.range.end = crate::reference::CellRef::new(right, bottom);
                                append(&mut values, &[sum]);
                            }
                        }
                    }
                }
                if name == "SUMIF" { return ReferenceReads::value(values); }
                if let Some(index) = Self::metadata_reference_argument(name, args) {
                    let mut measured = Vec::new();
                    for (i, read) in reads.iter().enumerate() {
                        append(&mut measured, if i == index { &read.addresses } else { &read.values });
                    }
                    return ReferenceReads::value(measured);
                }
                let mut addresses = Vec::new();
                for (i, read) in reads.iter().enumerate() {
                    let reference = match name.as_str() {
                        "IF" | "CHOOSE" => i > 0,
                        "IFS" => i % 2 == 1,
                        "SWITCH" => i > 0 && (i % 2 == 0 || i + 1 == args.len()),
                        "INDEX" | "TAKE" | "DROP" | "OFFSET" | "TRIMRANGE" => i == 0,
                        "XLOOKUP" => i == 2,
                        "_UNION" | "_ISECT" | "_SPILL" => true,
                        _ => false,
                    };
                    append(&mut addresses, if reference { &read.addresses } else { &read.values });
                }
                ReferenceReads { values, addresses }
            },
            Expr::Table { .. } | Expr::Literal(_) | Expr::Array(_) => ReferenceReads::default(),
        }
    }

    /// Address and shape references invalidate formulas when their cells
    /// change, even though their contents impose no calculation ordering.
    pub fn invalidation_references(&self) -> Vec<Reference> {
        let mut found = self.value_references();
        self.visit(&mut |node| {
            if let Expr::Function { name, args } = node {
                if let Some(at) = Self::metadata_reference_argument(name, args) {
                    found.extend(args[at].references());
                }
            }
        });
        found
    }

    /// Every reference this expression reads. Used to build the dependency graph.
    pub fn references(&self) -> Vec<Reference> {
        let mut found = Vec::new();
        self.visit(&mut |node| {
            if let Expr::Ref(reference) = node {
                found.push(reference.clone());
            }
        });
        found
    }

    /// Every reference whose CONTENTS this expression depends on.
    ///
    /// The same as `references`, less the ranges that are only being measured.
    /// This is what the dependency graph wants: a formula that asks how tall a
    /// block is does not have to wait for the block to be worked out, and
    /// treating it as though it did makes a cycle out of the commonest way to
    /// number a series.
    pub fn value_references(&self) -> Vec<Reference> {
        let mut found = Vec::new();
        self.walk_values(&mut found);
        found
    }

    fn walk_values(&self, found: &mut Vec<Reference>) {
        match self {
            Expr::Ref(reference) => found.push(reference.clone()),
            Expr::Unary { operand, .. } => operand.walk_values(found),
            Expr::Binary { lhs, rhs, .. } => {
                lhs.walk_values(found);
                rhs.walk_values(found);
            }
            Expr::Function { name, args } => {
                if name == "SUMIF" && args.len() == 3 {
                    if let (Expr::Ref(criteria), Expr::Ref(sum)) = (&args[0], &args[2]) {
                        let right = sum.range.start.col.checked_add(criteria.range.width() - 1);
                        let bottom = sum.range.start.row.checked_add(criteria.range.height() - 1);
                        if let (Some(right), Some(bottom)) = (right, bottom) {
                            if right <= crate::reference::MAX_COL && bottom <= crate::reference::MAX_ROW {
                                args[0].walk_values(found);
                                args[1].walk_values(found);
                                let mut sum = sum.clone();
                                sum.range.end = crate::reference::CellRef::new(right, bottom);
                                found.push(sum);
                                return;
                            }
                        }
                    }
                }
                if name == "INDEX" {
                    if let Some(selected) = Self::indexed_references(args) {
                        found.extend(selected);
                        for index in args.iter().skip(1) { index.walk_values(found); }
                        return;
                    }
                }
                let measured = Self::metadata_reference_argument(name, args);
                for (at, arg) in args.iter().enumerate() {
                    if measured == Some(at) { arg.walk_reference_values(found); }
                    else { arg.walk_values(found); }
                }
            }
            // An array points at no cells, so it contributes no references
            // for a dependency to be built from.
            Expr::Table { .. } | Expr::Literal(_) | Expr::Name(_) | Expr::Array(_) => {}
        }
    }

    /// Literal INDEX selectors read selected cells, rather than every cell
    /// in the source rectangle. In particular a whole-row source may include
    /// the formula itself without the selected cells being circular.
    /// Unsupported selectors conservatively retain the original dependency.
    fn indexed_references(args: &[Expr]) -> Option<Vec<Reference>> {
        let Expr::Ref(source) = args.first()? else { return None };
        if !(2..=3).contains(&args.len()) { return None; }
        let numbers = |expr: &Expr| -> Option<(Vec<u32>, usize, usize)> {
            let (values, width, height): (Vec<&Value>, _, _) = match expr {
                Expr::Literal(value) => (vec![value], 1, 1),
                Expr::Array(rows) => (rows.iter().flatten().collect(), rows.first()?.len(), rows.len()),
                _ => return None,
            };
            if width == 0 || width.checked_mul(height)? != values.len() { return None; }
            let values = values.into_iter().map(|value| {
                let n = value.to_number().ok()?;
                (n.is_finite() && n >= 0.0 && n <= u32::MAX as f64).then_some(n.trunc() as u32)
            }).collect::<Option<Vec<_>>>()?;
            Some((values, width, height))
        };
        let first = numbers(&args[1])?;
        let (rows, cols) = if args.len() == 3 {
            (first, numbers(&args[2])?)
        } else if source.range.height() == 1 {
            ((vec![1], 1, 1), first)
        } else if source.range.width() == 1 {
            (first, (vec![1], 1, 1))
        } else { return None; };
        // Retain broadcasting pairs, rather than a Cartesian product that
        // could introduce a circular edge for a cell never selected.
        let width = rows.1.max(cols.1); let height = rows.2.max(cols.2);
        if width.checked_mul(height)? > 1_048_576
            || (rows.1 != 1 && rows.1 != width) || (cols.1 != 1 && cols.1 != width)
            || (rows.2 != 1 && rows.2 != height) || (cols.2 != 1 && cols.2 != height) { return None; }
        let mut selected = Vec::new();
        for y in 0..height {
            for x in 0..width {
                let row = rows.0[(if rows.2 == 1 { 0 } else { y }) * rows.1 + if rows.1 == 1 { 0 } else { x }];
                let col = cols.0[(if cols.2 == 1 { 0 } else { y }) * cols.1 + if cols.1 == 1 { 0 } else { x }];
                if row > source.range.height() || col > source.range.width() { continue; }
                let mut reference = source.clone();
                if row != 0 {
                    reference.range.start.row += row - 1;
                    reference.range.end.row = reference.range.start.row;
                }
                if col != 0 {
                    reference.range.start.col += col - 1;
                    reference.range.end.col = reference.range.start.col;
                }
                selected.push(reference);
            }
        }
        Some(selected)
    }
}

#[cfg(test)]
mod reference_dependency_tests {
    use super::*;

    #[test]
    fn reference_let_dependencies_respect_shadowing() {
        let outer = crate::parser::parse("ISREF(LET(x_,A1,LET(x_,x_+1,x_)))").unwrap();
        let references = outer.value_references();
        assert_eq!(references.len(), 1);
        assert_eq!(references[0].range.start.to_a1(), "A1");
        let shadowed = crate::parser::parse("ISREF(LET(x_,A1,LET(x_,B1,x_)))").unwrap();
        assert!(shadowed.value_references().is_empty());
        assert_eq!(shadowed.invalidation_references().len(), 2);
        let sum = crate::parser::parse("ISREF(LET(x_,SUMIF(A1:A5,1,B1),x_))").unwrap();
        let ranges: Vec<_> = sum.value_references().iter().map(|r| r.range.to_string()).collect();
        assert_eq!(ranges, ["A1:A5", "B1:B5"]);
    }

    #[test]
    fn repeated_reference_aliases_do_not_expand_the_expression() {
        let mut formula = "ISREF(LET(x_0,A1".to_string();
        for i in 1..=48 { formula.push_str(&format!(",x_{i},x_{}+x_{}", i-1, i-1)); }
        formula.push_str(",x_48))");
        let parsed = crate::parser::parse(&formula).unwrap();
        let references = parsed.value_references();
        assert_eq!(references.len(), 1);
        assert_eq!(references[0].range.start.to_a1(), "A1");
    }
}
