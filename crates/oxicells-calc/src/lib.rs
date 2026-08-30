// This Source Code Form is subject to the terms of the Mozilla Public
// License, v. 2.0. If a copy of the MPL was not distributed with this
// file, You can obtain one at https://mozilla.org/MPL/2.0/.

//! An Excel-compatible formula engine: tokeniser, parser, dependency graph, and
//! dependency-ordered recalculation.
//!
//! # Why this is separate from `oxicells-core`
//!
//! `oxicells-core` carries a small evaluator that is correct for its own job —
//! rendering a workbook Excel already calculated, where the cached values in the
//! file are the right answer. This crate exists for the other job: recalculating
//! after something writes a new value, which is what happens when a macro drives
//! the sheet. That needs a dependency graph, and a graph cannot be retrofitted
//! onto snapshot-then-evaluate.
//!
//! Dependency direction is `oxicells-core` → `oxicells-calc`, never the reverse.
//!
//! # Determinism
//!
//! Every result must be reproducible run to run, because the engine is verified
//! by diffing its output against Excel driven over COM. Two consequences:
//!
//! - `BTreeMap` everywhere, so iteration order never varies.
//! - No volatile functions (`NOW`, `TODAY`, `RAND`). They need an injected clock
//!   and seed before they can be added; see `functions.rs`.
//!
//! # Example
//!
//! ```
//! use oxicells_calc::{Value, Workbook};
//!
//! let mut wb = Workbook::new();
//! wb.add_sheet("Sheet1");
//! wb.set_value("Sheet1", "A1", Value::Number(1000.0)).unwrap();
//! wb.set_formula("Sheet1", "B1", "=A1*0.1").unwrap();
//! wb.set_formula("Sheet1", "C1", r#"=IF(B1>50,"high","low")"#).unwrap();
//! wb.recalculate();
//!
//! assert_eq!(wb.value("Sheet1", "B1"), Value::Number(100.0));
//! assert_eq!(wb.value("Sheet1", "C1"), Value::text("high"));
//! ```

pub mod ast;
pub mod datetime;
pub mod engine;
pub mod functions;
pub mod lexer;
pub mod parser;
pub mod numfmt;
pub mod r1c1;
pub mod reference;
pub mod value;

pub use ast::{BinaryOp, Expr, UnaryOp};
pub use engine::{CalcError, RecalcReport, Workbook};
pub use numfmt::format_number;
pub use lexer::{
    shift_formula_references, translate_formula_references, ReferenceShift, ShiftAxis,
};
pub use lexer::ParseError;
pub use parser::parse;
pub use r1c1::{formula_from_r1c1, formula_to_r1c1};
pub use reference::{CellRef, RangeRef, Reference};
pub use value::{ExcelError, Value};
