// This Source Code Form is subject to the terms of the Mozilla Public
// License, v. 2.0. If a copy of the MPL was not distributed with this
// file, You can obtain one at https://mozilla.org/MPL/2.0/.

//! Workbook model, evaluator, and dependency-ordered recalculation.
//!
//! The distinguishing feature versus a display-oriented evaluator is that
//! formulas are recalculated in **dependency order**. Snapshotting every value
//! and then evaluating each formula against that snapshot is adequate when the
//! workbook was written by Excel (the cached values are already correct), but it
//! produces stale results the moment anything writes a new value — which is
//! exactly what happens when a macro drives the sheet.

use std::collections::{BTreeMap, BTreeSet, VecDeque};
use std::fmt;

use crate::ast::{BinaryOp, Expr, UnaryOp};
use crate::functions::{self, block_of, reach, Arg, RangeData};

/// Which cell is being worked out, when that is known. `ROW()` with no
/// argument is the only thing that needs it, and a formula evaluated on its
/// own rather than in a cell has no answer for it.
type At = Option<(u32, u32)>;
use crate::lexer::ParseError;
use crate::parser::parse;
use crate::reference::{parse_a1, CellRef, RangeRef};
use crate::value::{compare, ExcelError, Value};

/// Guards against defined names that refer to one another in a loop.
const MAX_EVAL_DEPTH: u32 = 64;

#[derive(Debug, Clone, PartialEq, Eq)]
pub enum CalcError {
    UnknownSheet(String),
    BadReference(String),
    Parse(ParseError),
}

impl fmt::Display for CalcError {
    fn fmt(&self, f: &mut fmt::Formatter<'_>) -> fmt::Result {
        match self {
            CalcError::UnknownSheet(s) => write!(f, "unknown sheet {s:?}"),
            CalcError::BadReference(s) => write!(f, "not a cell reference: {s:?}"),
            CalcError::Parse(e) => write!(f, "{e}"),
        }
    }
}

impl std::error::Error for CalcError {}

impl From<ParseError> for CalcError {
    fn from(e: ParseError) -> CalcError {
        CalcError::Parse(e)
    }
}

#[derive(Debug, Clone, PartialEq)]
enum Cell {
    Literal(Value),
    Formula {
        source: String,
        expr: Expr,
        cached: Value,
    },
}

impl Cell {
    fn value(&self) -> &Value {
        match self {
            Cell::Literal(v) => v,
            Cell::Formula { cached, .. } => cached,
        }
    }
}

#[derive(Debug, Clone, Default)]
struct Sheet {
    /// Keyed by `(col, row)`, both 0-based. `BTreeMap` so that iteration — and
    /// therefore recalculation order among independent cells — is deterministic.
    cells: BTreeMap<(u32, u32), Cell>,
}

/// What a recalculation did.
#[derive(Debug, Clone, Default, PartialEq, Eq)]
pub struct RecalcReport {
    pub evaluated: usize,
    /// Cells that take part in a reference cycle, as `Sheet!A1`.
    ///
    /// With iterative calculation off, Excel leaves these at `0` and raises a
    /// warning rather than producing an error value; that is reproduced here,
    /// and the cells are listed so the caller can surface the warning.
    pub circular: Vec<String>,
}

#[derive(Debug, Clone, Default)]
pub struct Workbook {
    sheets: BTreeMap<String, Sheet>,
    names: BTreeMap<String, Expr>,
}

impl Workbook {
    pub fn new() -> Workbook {
        Workbook::default()
    }

    pub fn add_sheet(&mut self, name: &str) {
        self.sheets.entry(name.to_string()).or_default();
    }

    pub fn sheet_names(&self) -> impl Iterator<Item = &str> {
        self.sheets.keys().map(String::as_str)
    }

    /// Define a workbook-scoped name, e.g. `TAX_RATE` → `0.1`.
    pub fn define_name(&mut self, name: &str, formula: &str) -> Result<(), CalcError> {
        let expr = parse(formula)?;
        self.names.insert(name.to_uppercase(), expr);
        Ok(())
    }

    pub fn set_value(&mut self, sheet: &str, a1: &str, value: Value) -> Result<(), CalcError> {
        let cell = self.cell_ref(a1)?;
        self.sheet_mut(sheet)?
            .cells
            .insert(cell.coord(), Cell::Literal(value));
        Ok(())
    }

    pub fn set_formula(&mut self, sheet: &str, a1: &str, formula: &str) -> Result<(), CalcError> {
        let cell = self.cell_ref(a1)?;
        let expr = parse(formula)?;
        self.sheet_mut(sheet)?.cells.insert(
            cell.coord(),
            Cell::Formula {
                source: formula.to_string(),
                expr,
                cached: Value::Blank,
            },
        );
        Ok(())
    }

    pub fn clear(&mut self, sheet: &str, a1: &str) -> Result<(), CalcError> {
        let cell = self.cell_ref(a1)?;
        self.sheet_mut(sheet)?.cells.remove(&cell.coord());
        Ok(())
    }

    /// Current value of a cell. Reads the cached result for formulas, so call
    /// [`Workbook::recalculate`] after any mutation.
    pub fn value(&self, sheet: &str, a1: &str) -> Value {
        let Ok(cell) = self.cell_ref(a1) else {
            return Value::Error(ExcelError::Ref);
        };
        self.value_at(sheet, cell.col, cell.row)
    }

    /// The formula source of a cell, if it has one.
    pub fn formula(&self, sheet: &str, a1: &str) -> Option<&str> {
        let cell = self.cell_ref(a1).ok()?;
        match self.sheets.get(sheet)?.cells.get(&cell.coord())? {
            Cell::Formula { source, .. } => Some(source),
            Cell::Literal(_) => None,
        }
    }

    /// Evaluate a formula against the workbook without storing it anywhere.
    pub fn evaluate(&self, sheet: &str, formula: &str) -> Result<Value, CalcError> {
        let expr = parse(formula)?;
        // Nowhere in particular, so `ROW()` with no argument has no answer.
        Ok(formula_result(self.eval(&expr, sheet, 0, None)))
    }

    // -- recalculation ----------------------------------------------------

    /// Recalculate every formula in dependency order.
    pub fn recalculate(&mut self) -> RecalcReport {
        let keys = self.formula_keys();
        let index: BTreeMap<&(String, (u32, u32)), usize> =
            keys.iter().enumerate().map(|(i, k)| (k, i)).collect();

        // Edge dep -> dependent, plus in-degree, for Kahn's algorithm.
        let mut dependents: Vec<BTreeSet<usize>> = vec![BTreeSet::new(); keys.len()];
        let mut indegree = vec![0usize; keys.len()];

        for (i, (sheet, _)) in keys.iter().enumerate() {
            for dep in self.dependencies_of(sheet, &keys[i].1) {
                if let Some(&j) = index.get(&dep) {
                    if j != i && dependents[j].insert(i) {
                        indegree[i] += 1;
                    }
                }
            }
        }

        let mut queue: VecDeque<usize> = (0..keys.len()).filter(|&i| indegree[i] == 0).collect();
        let mut order = Vec::with_capacity(keys.len());
        while let Some(i) = queue.pop_front() {
            order.push(i);
            for &j in &dependents[i] {
                indegree[j] -= 1;
                if indegree[j] == 0 {
                    queue.push_back(j);
                }
            }
        }

        let mut report = RecalcReport {
            evaluated: order.len(),
            circular: Vec::new(),
        };

        let emitted: BTreeSet<usize> = order.iter().copied().collect();

        for &i in &order {
            let (sheet, coord) = &keys[i];
            let Some(expr) = self.expr_at(sheet, coord) else {
                continue;
            };
            let value = formula_result(self.eval(&expr, sheet, 0, Some(*coord)));
            self.store_cached(sheet, coord, value);
        }

        // Anything Kahn could not emit sits in a cycle.
        for (i, (sheet, coord)) in keys.iter().enumerate() {
            if !emitted.contains(&i) {
                report
                    .circular
                    .push(format!("{}!{}", sheet, CellRef::new(coord.0, coord.1).to_a1()));
                self.store_cached(sheet, coord, Value::Number(0.0));
            }
        }

        report
    }

    fn formula_keys(&self) -> Vec<(String, (u32, u32))> {
        let mut keys = Vec::new();
        for (name, sheet) in &self.sheets {
            for (coord, cell) in &sheet.cells {
                if matches!(cell, Cell::Formula { .. }) {
                    keys.push((name.clone(), *coord));
                }
            }
        }
        keys
    }

    /// Cells this formula reads.
    ///
    /// A range dependency is expanded against the cells that actually exist,
    /// not against every address in the rectangle: `SUM(A:A)` must not
    /// enumerate a million cells.
    fn dependencies_of(&self, sheet: &str, coord: &(u32, u32)) -> Vec<(String, (u32, u32))> {
        let Some(expr) = self.expr_at(sheet, coord) else {
            return Vec::new();
        };
        let mut deps = Vec::new();
        for reference in expr.references() {
            let target = reference.sheet.as_deref().unwrap_or(sheet);
            let Some(target_sheet) = self.sheets.get(target) else {
                continue;
            };
            if reference.range.is_single() {
                let c = reference.range.start.coord();
                if target_sheet.cells.contains_key(&c) {
                    deps.push((target.to_string(), c));
                }
            } else {
                for c in target_sheet.cells.keys() {
                    if reference.range.contains(c.0, c.1) {
                        deps.push((target.to_string(), *c));
                    }
                }
            }
        }
        deps
    }

    fn expr_at(&self, sheet: &str, coord: &(u32, u32)) -> Option<Expr> {
        match self.sheets.get(sheet)?.cells.get(coord)? {
            Cell::Formula { expr, .. } => Some(expr.clone()),
            Cell::Literal(_) => None,
        }
    }

    fn store_cached(&mut self, sheet: &str, coord: &(u32, u32), value: Value) {
        if let Some(Cell::Formula { cached, .. }) = self
            .sheets
            .get_mut(sheet)
            .and_then(|s| s.cells.get_mut(coord))
        {
            *cached = value;
        }
    }

    // -- evaluation -------------------------------------------------------

    fn value_at(&self, sheet: &str, col: u32, row: u32) -> Value {
        self.sheets
            .get(sheet)
            .and_then(|s| s.cells.get(&(col, row)))
            .map(|c| c.value().clone())
            .unwrap_or(Value::Blank)
    }

    fn eval(&self, expr: &Expr, sheet: &str, depth: u32, at: At) -> Value {
        self.eval_arg(expr, sheet, depth, at).scalar()
    }

    fn eval_arg(&self, expr: &Expr, sheet: &str, depth: u32, at: At) -> Arg {
        self.eval_arg_inner(expr, sheet, depth, false, at)
    }

    /// `skip_subtotals` drops cells that are themselves `SUBTOTAL` formulas when
    /// materialising a range. See the `SUBTOTAL` arm below for why.
    fn eval_arg_inner(
        &self,
        expr: &Expr,
        sheet: &str,
        depth: u32,
        skip_subtotals: bool,
        at: At,
    ) -> Arg {
        if depth > MAX_EVAL_DEPTH {
            return Arg::Value(Value::Error(ExcelError::Num));
        }

        match expr {
            Expr::Literal(v) => Arg::Value(v.clone()),

            Expr::Ref(reference) => {
                let target = reference.sheet.as_deref().unwrap_or(sheet);
                if !self.sheets.contains_key(target) {
                    return Arg::Value(Value::Error(ExcelError::Ref));
                }
                // Even a single cell is handed over as a range. It is still a
                // *reference*, not a literal, and the aggregate functions treat
                // the two differently: `SUM("5")` is 5, but `SUM(A1)` where A1
                // holds text is 0. `Arg::scalar` unwraps a 1x1 range
                // transparently, so nothing else has to care.
                Arg::Range(self.materialise(target, &reference.range, skip_subtotals))
            }

            Expr::Name(name) => match self.names.get(name) {
                Some(bound) => self.eval_arg(&bound.clone(), sheet, depth + 1, at),
                None => Arg::Value(Value::Error(ExcelError::Name)),
            },

            Expr::Unary { op, operand } => {
                let operand = self.eval_arg(operand, sheet, depth + 1, at);
                match operand {
                    Arg::Value(v) => Arg::Value(apply_unary(*op, v)),
                    Arg::Range(block) => Arg::Range(RangeData {
                        width: block.width,
                        height: block.height,
                        cells: block
                            .cells
                            .iter()
                            .map(|one| apply_unary(*op, one.clone()))
                            .collect(),
                    }),
                }
            }

            Expr::Binary { op, lhs, rhs } => {
                let a = self.eval_arg(lhs, sheet, depth + 1, at);
                let b = self.eval_arg(rhs, sheet, depth + 1, at);
                across(*op, &a, &b)
            }

            Expr::Function { name, args } if name == "ROW" || name == "COLUMN" => {
                self.which_line(name, args, at)
            }

            Expr::Function { name, args } => {
                // SUBTOTAL ignores any cell in its range that is itself a
                // SUBTOTAL, which is how a column of group subtotals can be
                // summed by a grand total without double counting. That cannot
                // be decided from the values alone, so the exclusion has to
                // happen here, while the range is still a reference.
                let nested = name == "SUBTOTAL";
                let evaluated: Vec<Arg> = args
                    .iter()
                    .map(|a| self.eval_arg_inner(a, sheet, depth + 1, nested, at))
                    .collect();
                functions::call_arg(name, &evaluated)
            }
        }
    }

    /// `ROW` and `COLUMN`, answered from the reference rather than its contents.
    ///
    /// With no argument they mean the cell being worked out. With a reference
    /// they mean its rows or its columns — all of them, as a block, which is
    /// what makes `SMALL(IF(range = x, ROW(range)), n)` pick out the nth row
    /// where something is true. Answering only the first would give one number
    /// where five hundred were wanted.
    fn which_line(&self, name: &str, args: &[Expr], at: At) -> Arg {
        let down = name == "ROW";
        let Some(first) = args.first() else {
            // No argument: wherever we are. Outside a cell there is no answer.
            // Counted from one, where everything inside here counts from zero.
            return match at {
                Some((col, row)) => Arg::Value(Value::Number(if down {
                    row as f64 + 1.0
                } else {
                    col as f64 + 1.0
                })),
                None => Arg::Value(Value::Error(ExcelError::Value)),
            };
        };
        let Expr::Ref(reference) = first else {
            return Arg::Value(Value::Error(ExcelError::Value));
        };
        let range = &reference.range;
        let (from, to) = if down {
            (range.start.row, range.end.row)
        } else {
            (range.start.col, range.end.col)
        };
        let many = (to.saturating_sub(from) + 1) as usize;
        let cells: Vec<Value> = (from..=to)
            .map(|one| Value::Number(one as f64 + 1.0))
            .collect();
        Arg::Range(RangeData {
            width: if down { 1 } else { many },
            height: if down { many } else { 1 },
            cells,
        })
    }

    fn materialise(&self, sheet: &str, range: &RangeRef, skip_subtotals: bool) -> RangeData {
        RangeData::from_range(range, |col, row| {
            if skip_subtotals && self.is_subtotal_cell(sheet, col, row) {
                Value::Blank
            } else {
                self.value_at(sheet, col, row)
            }
        })
    }

    fn is_subtotal_cell(&self, sheet: &str, col: u32, row: u32) -> bool {
        let Some(Cell::Formula { expr, .. }) =
            self.sheets.get(sheet).and_then(|s| s.cells.get(&(col, row)))
        else {
            return false;
        };
        matches!(expr, Expr::Function { name, .. } if name == "SUBTOTAL")
    }

    // -- helpers ----------------------------------------------------------

    fn cell_ref(&self, a1: &str) -> Result<CellRef, CalcError> {
        parse_a1(a1).ok_or_else(|| CalcError::BadReference(a1.to_string()))
    }

    fn sheet_mut(&mut self, name: &str) -> Result<&mut Sheet, CalcError> {
        self.sheets
            .get_mut(name)
            .ok_or_else(|| CalcError::UnknownSheet(name.to_string()))
    }
}

/// A formula never results in a blank cell.
///
/// `=A1` where `A1` is empty shows `0`, not an empty cell, and so does an
/// omitted branch such as `=IF(TRUE,,)`. Excel caches `0` in both cases. The
/// blank only exists while a reference is being read; once it becomes a
/// formula's answer it is a number.
fn formula_result(value: Value) -> Value {
    if value.is_blank() {
        Value::Number(0.0)
    } else {
        value
    }
}

fn apply_unary(op: UnaryOp, v: Value) -> Value {
    if let Some(e) = v.err() {
        return Value::Error(e);
    }
    let n = match v.to_number() {
        Ok(n) => n,
        Err(e) => return Value::Error(e),
    };
    match op {
        UnaryOp::Neg => Value::Number(-n),
        UnaryOp::Plus => Value::Number(n),
        UnaryOp::Percent => Value::Number(n / 100.0),
    }
}

/// Apply `op` to every pair of cells, stretching each side to fit the other.
///
/// Two single values are one value, as they always were. Anything wider or
/// taller becomes a block of answers as big as the larger of the two in each
/// direction: a column of 72 against a row of 14 makes a block of 14 by 72,
/// which is the shape the cross-tab idiom depends on.
///
/// A side that is one cell across is read for every column, and one cell down
/// for every row. Where the two disagree and neither is one, the cells with
/// nothing on one side are `#N/A` — Excel's answer, and the reason it is worth
/// stretching rather than refusing.
fn across(op: BinaryOp, a: &Arg, b: &Arg) -> Arg {
    let (left, right) = match (a, b) {
        (Arg::Value(x), Arg::Value(y)) => return Arg::Value(apply_binary(op, x.clone(), y.clone())),
        _ => (block_of(a), block_of(b)),
    };
    if left.cells.len() == 1 && right.cells.len() == 1 {
        return Arg::Value(apply_binary(op, left.cells[0].clone(), right.cells[0].clone()));
    }
    let width = left.width.max(right.width);
    let height = left.height.max(right.height);
    let mut cells = Vec::with_capacity(width * height);
    for row in 0..height {
        for col in 0..width {
            match (reach(&left, col, row), reach(&right, col, row)) {
                (Some(x), Some(y)) => cells.push(apply_binary(op, x, y)),
                _ => cells.push(Value::Error(ExcelError::NA)),
            }
        }
    }
    Arg::Range(RangeData {
        width,
        height,
        cells,
    })
}

fn apply_binary(op: BinaryOp, a: Value, b: Value) -> Value {
    if let Some(e) = a.err() {
        return Value::Error(e);
    }
    if let Some(e) = b.err() {
        return Value::Error(e);
    }

    if op.is_comparison() {
        return match compare(&a, &b) {
            Ok(ord) => Value::Logical(match op {
                BinaryOp::Eq => ord.is_eq(),
                BinaryOp::Ne => !ord.is_eq(),
                BinaryOp::Lt => ord.is_lt(),
                BinaryOp::Le => ord.is_le(),
                BinaryOp::Gt => ord.is_gt(),
                _ => ord.is_ge(),
            }),
            Err(e) => Value::Error(e),
        };
    }

    if op == BinaryOp::Concat {
        return match (a.to_text(), b.to_text()) {
            (Ok(x), Ok(y)) => Value::Text(x + &y),
            (Err(e), _) | (_, Err(e)) => Value::Error(e),
        };
    }

    let (x, y) = match (a.to_number(), b.to_number()) {
        (Ok(x), Ok(y)) => (x, y),
        (Err(e), _) | (_, Err(e)) => return Value::Error(e),
    };

    match op {
        BinaryOp::Add => Value::Number(x + y),
        BinaryOp::Sub => Value::Number(x - y),
        BinaryOp::Mul => Value::Number(x * y),
        BinaryOp::Div => {
            if y == 0.0 {
                Value::Error(ExcelError::DivZero)
            } else {
                Value::Number(x / y)
            }
        }
        BinaryOp::Pow => {
            let r = x.powf(y);
            if r.is_nan() {
                Value::Error(ExcelError::Num)
            } else {
                Value::Number(r)
            }
        }
        _ => unreachable!("comparison and concat handled above"),
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    fn book() -> Workbook {
        let mut wb = Workbook::new();
        wb.add_sheet("Sheet1");
        wb
    }

    /// A column of values in A, and a second column in B.
    fn two_columns() -> Workbook {
        let mut wb = book();
        for (at, (a, b)) in [(1.0, 10.0), (2.0, 20.0), (3.0, 30.0)].iter().enumerate() {
            wb.set_value("Sheet1", &format!("A{}", at + 1), Value::Number(*a))
                .unwrap();
            wb.set_value("Sheet1", &format!("B{}", at + 1), Value::Number(*b))
                .unwrap();
        }
        wb
    }

    #[test]
    fn row_and_column_are_answered_from_the_reference() {
        // Neither can be answered from the VALUES of the arguments, which is
        // all the function library is ever shown: by the time `$A$2:$A$5`
        // reaches it, it is four cell contents with no idea where they came
        // from. So these two are settled in the engine, where the reference is
        // still a reference.
        let mut wb = book();
        wb.set_formula("Sheet1", "C7", "=ROW()").unwrap();
        wb.set_formula("Sheet1", "D8", "=COLUMN()").unwrap();
        wb.set_formula("Sheet1", "E1", "=ROW(A5)").unwrap();
        wb.set_formula("Sheet1", "E2", "=COLUMN(C1)").unwrap();
        wb.recalculate();
        assert_eq!(wb.value("Sheet1", "C7"), Value::Number(7.0));
        assert_eq!(wb.value("Sheet1", "D8"), Value::Number(4.0));
        assert_eq!(wb.value("Sheet1", "E1"), Value::Number(5.0));
        assert_eq!(wb.value("Sheet1", "E2"), Value::Number(3.0));
    }

    #[test]
    fn row_of_a_range_is_every_row_in_it() {
        // What makes `SMALL(IF(range = x, ROW(range)), n)` pick out the nth
        // row where something is true. Answering only the first would give one
        // number where five hundred were wanted, and the whole family of
        // formulas built on it came back empty.
        let mut wb = book();
        for at in 2..=5 {
            wb.set_value("Sheet1", &format!("A{at}"), Value::Number(at as f64))
                .unwrap();
        }
        // 2 + 3 + 4 + 5
        wb.set_formula("Sheet1", "C1", "=SUM(ROW(A2:A5))").unwrap();
        // The smallest row where A holds 4, which is row 4.
        wb.set_formula("Sheet1", "C2", "=SMALL(IF(A2:A5=4,ROW(A2:A5)),1)")
            .unwrap();
        wb.recalculate();
        assert_eq!(wb.value("Sheet1", "C1"), Value::Number(14.0));
        assert_eq!(wb.value("Sheet1", "C2"), Value::Number(4.0));
    }

    #[test]
    fn a_formula_evaluated_outside_a_cell_has_no_row() {
        let wb = book();
        assert_eq!(
            wb.evaluate("Sheet1", "=ROW()"),
            Ok(Value::Error(ExcelError::Value))
        );
        // But one with a reference still knows its own answer.
        assert_eq!(wb.evaluate("Sheet1", "=ROW(B9)"), Ok(Value::Number(9.0)));
    }

    #[test]
    fn comparing_a_range_to_a_value_answers_for_every_cell() {
        // The single largest thing this could not do. Both sides of the `=`
        // were collapsed to one value first, and a column of three does not
        // collapse, so the whole formula came back #VALUE!.
        let mut wb = two_columns();
        wb.set_formula("Sheet1", "D1", "=SUMPRODUCT((A1:A3=2)*B1:B3)")
            .unwrap();
        wb.recalculate();
        assert_eq!(wb.value("Sheet1", "D1"), Value::Number(20.0));
    }

    #[test]
    fn two_conditions_multiply_together() {
        let mut wb = two_columns();
        wb.set_formula("Sheet1", "D1", "=SUMPRODUCT((A1:A3>=2)*(B1:B3<30)*B1:B3)")
            .unwrap();
        wb.recalculate();
        assert_eq!(wb.value("Sheet1", "D1"), Value::Number(20.0));
    }

    #[test]
    fn a_column_meeting_a_row_makes_a_block_of_both() {
        // The cross-tab idiom, and the reason shapes have to STRETCH rather
        // than merely match: one wide by three tall against three wide by one
        // tall is a three-by-three block, and a rule that only handled equal
        // shapes would miss exactly the case this is written for.
        let mut wb = book();
        for at in 1..=3 {
            wb.set_value("Sheet1", &format!("A{at}"), Value::Number(at as f64))
                .unwrap();
            wb.set_value(
                "Sheet1",
                &format!("{}1", (b'C' + at as u8 - 1) as char),
                Value::Number(at as f64 * 10.0),
            )
            .unwrap();
        }
        // Each of A1:A3 against each of C1:E1, so nine products, summing to
        // (1+2+3) * (10+20+30) = 360.
        wb.set_formula("Sheet1", "A5", "=SUMPRODUCT(A1:A3*C1:E1)").unwrap();
        wb.recalculate();
        assert_eq!(wb.value("Sheet1", "A5"), Value::Number(360.0));
    }

    #[test]
    fn a_function_meant_for_one_value_answers_for_every_cell() {
        // `ISNUMBER(SEARCH(x, A1:A3))` is a column of three yes-or-nos. SEARCH
        // wants a piece of text and is handed three, and the answer is three
        // answers.
        let mut wb = book();
        for (at, word) in ["alpha", "beta", "alphabet"].iter().enumerate() {
            wb.set_value("Sheet1", &format!("A{}", at + 1), Value::text(*word))
                .unwrap();
            wb.set_value(
                "Sheet1",
                &format!("B{}", at + 1),
                Value::Number((at + 1) as f64),
            )
            .unwrap();
        }
        wb.set_formula(
            "Sheet1",
            "D1",
            "=SUMPRODUCT(ISNUMBER(SEARCH(\"alpha\",A1:A3))*B1:B3)",
        )
        .unwrap();
        wb.recalculate();
        // "alpha" is in the first and the third, so 1 + 3.
        assert_eq!(wb.value("Sheet1", "D1"), Value::Number(4.0));
    }

    #[test]
    fn a_function_handed_ranges_on_purpose_is_left_alone() {
        // DGET takes three ranges and means something by all of them. It is
        // not implemented, so it answers `#NAME?` — and a great many real
        // sheets are written to swallow exactly that: `IF(ISERR(DGET(...)),,
        // DGET(...))` gives nothing whether DGET works or not.
        //
        // The first version of this asked "is it a known aggregate?" rather
        // than "is it known to take one value?", so DGET fell through, was
        // applied to each of six hundred cells in turn, and answered with six
        // hundred `#NAME?`s. A block of them does not fit in a cell, so the
        // swallow stopped working and the formula turned `#VALUE!` — seventy
        // eight cells of a corpus that had been right for months.
        //
        // Anything not named as one-at-a-time keeps the behaviour it had.
        let mut wb = book();
        wb.set_value("Sheet1", "A1", Value::text("name")).unwrap();
        wb.set_value("Sheet1", "A2", Value::text("ann")).unwrap();
        wb.set_value("Sheet1", "C1", Value::text("name")).unwrap();
        wb.set_value("Sheet1", "C2", Value::text("bob")).unwrap();
        wb.set_formula("Sheet1", "E1", "=DGET(A1:A2,1,C1:C2)").unwrap();
        wb.set_formula("Sheet1", "E2", "=IF(ISERR(DGET(A1:A2,1,C1:C2)),,1)")
            .unwrap();
        wb.recalculate();
        // One answer, not a block of them.
        assert_eq!(wb.value("Sheet1", "E1"), Value::Error(ExcelError::Name));
        // And so the sheet's own way of swallowing it still works.
        assert_eq!(wb.value("Sheet1", "E2"), Value::Number(0.0));
    }

    #[test]
    fn an_array_that_reaches_a_cell_still_will_not_fit_in_one() {
        // Nothing about this changed, and it is worth pinning that it did not:
        // a block of answers written into a single cell is #VALUE!, as it was
        // before any of this and as it is in Excel without dynamic arrays.
        let mut wb = two_columns();
        wb.set_formula("Sheet1", "D1", "=A1:A3+1").unwrap();
        wb.recalculate();
        assert_eq!(
            wb.value("Sheet1", "D1"),
            Value::Error(ExcelError::Value)
        );
    }

    #[test]
    fn where_two_blocks_disagree_and_neither_is_one_the_rest_is_not_available() {
        let mut wb = book();
        for at in 1..=3 {
            wb.set_value("Sheet1", &format!("A{at}"), Value::Number(1.0))
                .unwrap();
        }
        wb.set_value("Sheet1", "B1", Value::Number(1.0)).unwrap();
        wb.set_value("Sheet1", "B2", Value::Number(1.0)).unwrap();
        // Three against two: the third pair has nothing on one side. SUM skips
        // nothing, so the #N/A comes through.
        wb.set_formula("Sheet1", "D1", "=SUM(A1:A3+B1:B2)").unwrap();
        wb.recalculate();
        assert_eq!(wb.value("Sheet1", "D1"), Value::Error(ExcelError::NA));
    }

    #[test]
    fn formula_chains_resolve_in_dependency_order() {
        // C1 depends on B1 which depends on A1. A snapshot-then-evaluate engine
        // leaves C1 stale; this is the case that motivated the rewrite.
        let mut wb = book();
        wb.set_value("Sheet1", "A1", Value::Number(1.0)).unwrap();
        wb.set_formula("Sheet1", "B1", "=A1+1").unwrap();
        wb.set_formula("Sheet1", "C1", "=B1+1").unwrap();
        wb.recalculate();
        assert_eq!(wb.value("Sheet1", "C1"), Value::Number(3.0));

        wb.set_value("Sheet1", "A1", Value::Number(10.0)).unwrap();
        wb.recalculate();
        assert_eq!(wb.value("Sheet1", "B1"), Value::Number(11.0));
        assert_eq!(wb.value("Sheet1", "C1"), Value::Number(12.0));
    }

    #[test]
    fn dependency_order_holds_regardless_of_insertion_order() {
        let mut wb = book();
        // Deliberately define the dependent cell first.
        wb.set_formula("Sheet1", "Z1", "=Y1*2").unwrap();
        wb.set_formula("Sheet1", "Y1", "=X1*2").unwrap();
        wb.set_value("Sheet1", "X1", Value::Number(3.0)).unwrap();
        wb.recalculate();
        assert_eq!(wb.value("Sheet1", "Z1"), Value::Number(12.0));
    }

    #[test]
    fn circular_references_are_reported_not_hung() {
        let mut wb = book();
        wb.set_formula("Sheet1", "A1", "=B1+1").unwrap();
        wb.set_formula("Sheet1", "B1", "=A1+1").unwrap();
        let report = wb.recalculate();
        assert_eq!(report.circular.len(), 2);
        assert!(report.circular.contains(&"Sheet1!A1".to_string()));
        assert_eq!(wb.value("Sheet1", "A1"), Value::Number(0.0));
    }

    #[test]
    fn ranges_aggregate_across_cells() {
        let mut wb = book();
        for (i, a1) in ["A1", "A2", "A3"].iter().enumerate() {
            wb.set_value("Sheet1", a1, Value::Number(i as f64 + 1.0)).unwrap();
        }
        wb.set_formula("Sheet1", "B1", "=SUM(A1:A3)").unwrap();
        wb.recalculate();
        assert_eq!(wb.value("Sheet1", "B1"), Value::Number(6.0));
    }

    #[test]
    fn cross_sheet_references_resolve() {
        let mut wb = book();
        wb.add_sheet("Data");
        wb.set_value("Data", "A1", Value::Number(42.0)).unwrap();
        wb.set_formula("Sheet1", "A1", "=Data!A1*2").unwrap();
        wb.recalculate();
        assert_eq!(wb.value("Sheet1", "A1"), Value::Number(84.0));
    }

    #[test]
    fn division_by_zero_produces_the_excel_error() {
        let mut wb = book();
        wb.set_formula("Sheet1", "A1", "=1/0").unwrap();
        wb.recalculate();
        assert_eq!(wb.value("Sheet1", "A1"), Value::Error(ExcelError::DivZero));
    }

    #[test]
    fn comparison_in_if_now_works() {
        // The prototype could not parse `A1>10` and returned #VALUE! here.
        let mut wb = book();
        wb.set_value("Sheet1", "A1", Value::Number(20.0)).unwrap();
        wb.set_formula("Sheet1", "B1", r#"=IF(A1>10,"big","small")"#).unwrap();
        wb.recalculate();
        assert_eq!(wb.value("Sheet1", "B1"), Value::text("big"));
    }

    #[test]
    fn a_reference_to_an_empty_cell_results_in_zero() {
        // Excel shows 0, not an empty cell, and caches 0 in the file.
        let mut wb = book();
        wb.set_formula("Sheet1", "B1", "=A1").unwrap();
        wb.set_formula("Sheet1", "B2", "=IF(TRUE,,)").unwrap();
        wb.recalculate();
        assert_eq!(wb.value("Sheet1", "B1"), Value::Number(0.0));
        assert_eq!(wb.value("Sheet1", "B2"), Value::Number(0.0));
        // An empty *literal* cell is still blank; only formula results convert.
        assert_eq!(wb.value("Sheet1", "A1"), Value::Blank);
    }

    #[test]
    fn subtotal_skips_nested_subtotals() {
        // A grand total over a column that already contains group subtotals
        // must not count the members twice.
        let mut wb = book();
        for (a1, n) in [("A1", 1.0), ("A2", 2.0), ("A4", 3.0), ("A5", 4.0)] {
            wb.set_value("Sheet1", a1, Value::Number(n)).unwrap();
        }
        wb.set_formula("Sheet1", "A3", "=SUBTOTAL(9,A1:A2)").unwrap();
        wb.set_formula("Sheet1", "A6", "=SUBTOTAL(9,A4:A5)").unwrap();
        wb.set_formula("Sheet1", "A7", "=SUBTOTAL(9,A1:A6)").unwrap();
        wb.recalculate();
        assert_eq!(wb.value("Sheet1", "A3"), Value::Number(3.0));
        assert_eq!(wb.value("Sheet1", "A6"), Value::Number(7.0));
        // 10, not 20.
        assert_eq!(wb.value("Sheet1", "A7"), Value::Number(10.0));
        // A plain SUM over the same range does double count, by design.
        wb.set_formula("Sheet1", "A8", "=SUM(A1:A6)").unwrap();
        wb.recalculate();
        assert_eq!(wb.value("Sheet1", "A8"), Value::Number(20.0));
    }

    #[test]
    fn absolute_references_evaluate_identically_to_relative() {
        let mut wb = book();
        wb.set_value("Sheet1", "A1", Value::Number(7.0)).unwrap();
        wb.set_formula("Sheet1", "B1", "=$A$1").unwrap();
        wb.recalculate();
        assert_eq!(wb.value("Sheet1", "B1"), Value::Number(7.0));
    }

    #[test]
    fn unary_minus_binds_tighter_than_power_when_evaluated() {
        let wb = book();
        assert_eq!(wb.evaluate("Sheet1", "=-2^2").unwrap(), Value::Number(4.0));
        assert_eq!(wb.evaluate("Sheet1", "=2^3^2").unwrap(), Value::Number(64.0));
    }

    #[test]
    fn percent_and_concat_evaluate() {
        let wb = book();
        assert_eq!(wb.evaluate("Sheet1", "=50%").unwrap(), Value::Number(0.5));
        assert_eq!(
            wb.evaluate("Sheet1", r#"="a"&1&TRUE"#).unwrap(),
            Value::text("a1TRUE")
        );
    }

    #[test]
    fn defined_names_resolve() {
        let mut wb = book();
        wb.define_name("TAX_RATE", "=0.1").unwrap();
        wb.set_value("Sheet1", "A1", Value::Number(1000.0)).unwrap();
        wb.set_formula("Sheet1", "B1", "=A1*TAX_RATE").unwrap();
        wb.recalculate();
        assert_eq!(wb.value("Sheet1", "B1"), Value::Number(100.0));
    }

    #[test]
    fn unknown_names_and_sheets_report_errors() {
        let mut wb = book();
        wb.set_formula("Sheet1", "A1", "=NOPE").unwrap();
        wb.set_formula("Sheet1", "A2", "=Missing!A1").unwrap();
        wb.recalculate();
        assert_eq!(wb.value("Sheet1", "A1"), Value::Error(ExcelError::Name));
        assert_eq!(wb.value("Sheet1", "A2"), Value::Error(ExcelError::Ref));
    }

    #[test]
    fn recalculation_is_reproducible() {
        let build = || {
            let mut wb = book();
            wb.set_value("Sheet1", "A1", Value::Number(2.0)).unwrap();
            wb.set_formula("Sheet1", "B1", "=A1*3").unwrap();
            wb.set_formula("Sheet1", "C1", "=B1+A1").unwrap();
            wb.recalculate();
            wb.value("Sheet1", "C1")
        };
        assert_eq!(build(), build());
    }
}
