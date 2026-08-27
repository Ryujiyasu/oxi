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
    /// Unary plus. Kept in the tree because it still forces numeric coercion.
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
            Expr::Literal(_) | Expr::Ref(_) | Expr::Name(_) => {}
        }
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
}
