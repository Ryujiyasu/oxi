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

/// Which part of a table a structured reference names.
enum TablePart {
    /// The rows below the heading, which is what a bare column name means.
    Body,
    Headers,
    Totals,
    All,
    /// The one row the asking cell is on.
    ThisRow,
}

/// Where a table is and what its columns are called.
#[derive(Debug, Clone)]
struct TableRef {
    sheet: String,
    first_row: u32,
    last_row: u32,
    first_col: u32,
    last_col: u32,
    header_rows: u32,
    headings: Vec<String>,
}

impl TableRef {
    /// The sheet column a heading names, counted as the sheet counts.
    fn column(&self, heading: &str) -> Option<u32> {
        let wanted = heading.trim().to_uppercase();
        self.headings
            .iter()
            .position(|one| *one == wanted)
            .map(|at| self.first_col + at as u32)
    }
}

/// Split on the commas that are not inside brackets, so a heading with a comma
/// in it stays one piece.
fn outside_brackets(asked: &str) -> Vec<String> {
    let mut parts = Vec::new();
    let mut depth = 0usize;
    let mut held = String::new();
    for ch in asked.chars() {
        match ch {
            '[' => {
                depth += 1;
                held.push(ch);
            }
            ']' => {
                depth = depth.saturating_sub(1);
                held.push(ch);
            }
            ',' if depth == 0 => parts.push(std::mem::take(&mut held)),
            _ => held.push(ch),
        }
    }
    parts.push(held);
    parts
}

/// Which cell is being worked out, when that is known. `ROW()` with no
/// argument is the only thing that needs it, and a formula evaluated on its
/// own rather than in a cell has no answer for it.
type At = Option<(u32, u32)>;
use crate::lexer::ParseError;
use crate::parser::parse;
use crate::reference::{parse_a1, CellRef, RangeRef, MAX_COL, MAX_ROW};
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
    /// The tables, by the name a formula calls them, so that
    /// `tblNomina[[#This Row],[DATE]]` can be told which cells it means.
    tables: BTreeMap<String, TableRef>,
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

    /// Tell the workbook about a table, so that formulas may name its columns.
    ///
    /// `rows` and `cols` are the whole table INCLUDING its heading, counted
    /// from zero, and `headings` names the columns left to right. A table with
    /// no headings can still be named as a whole; only its columns become
    /// unreachable.
    pub fn add_table(
        &mut self,
        sheet: &str,
        name: &str,
        rows: (u32, u32),
        cols: (u32, u32),
        header_rows: u32,
        headings: Vec<String>,
    ) {
        self.tables.insert(
            name.to_uppercase(),
            TableRef {
                sheet: sheet.to_string(),
                first_row: rows.0,
                last_row: rows.1,
                first_col: cols.0,
                last_col: cols.1,
                header_rows,
                headings: headings.into_iter().map(|one| one.to_uppercase()).collect(),
            },
        );
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

            Expr::Table { name, asked } => self.a_table_column(name, asked, at),

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

    /// The cells a structured reference names.
    ///
    /// `tblNomina[[#This Row],[DATE]]` is the DATE column of the row asking;
    /// `Table1[Description]` is the whole of that column below its heading;
    /// `Suppliers1[]` is every data cell. Which rows are meant depends on the
    /// part named — the body, the heading, this row — and which columns on the
    /// heading named, so the two are worked out separately and put together.
    fn a_table_column(&self, name: &str, asked: &str, at: At) -> Arg {
        match self.table_range(name, asked, at) {
            Ok((sheet, range)) => Arg::Range(self.materialise(&sheet, &range, false)),
            Err(why) => Arg::Value(Value::Error(why)),
        }
    }

    /// The same, as a reference rather than as its contents.
    ///
    /// `ROW(tbl[[#Headers],[ID]])` asks where the heading IS, not what it
    /// says, so the two callers need different halves of the same answer.
    fn table_range(&self, name: &str, asked: &str, at: At)
        -> Result<(String, RangeRef), ExcelError>
    {
        let Some(table) = self.tables.get(&name.to_uppercase()) else {
            return Err(ExcelError::Name);
        };
        // `[[#This Row],[DATE]]` arrives as `[#This Row],[DATE]`; a lone
        // `[DATE]` arrives as `DATE`. Splitting on the commas OUTSIDE any
        // brackets keeps a heading with a comma in it whole.
        let parts = outside_brackets(asked);
        let mut part = TablePart::Body;
        let mut wanted: Vec<&str> = Vec::new();
        for piece in &parts {
            let bare = piece.trim().trim_start_matches('[').trim_end_matches(']');
            match bare.trim().to_ascii_uppercase().as_str() {
                "#THIS ROW" | "@" => part = TablePart::ThisRow,
                "#HEADERS" => part = TablePart::Headers,
                "#TOTALS" => part = TablePart::Totals,
                "#ALL" => part = TablePart::All,
                "#DATA" => part = TablePart::Body,
                "" => {}
                _ => wanted.push(bare),
            }
        }

        // Which columns. A span written `[A]:[B]` reaches here as one piece
        // with a colon in it.
        let (first_col, last_col) = match wanted.len() {
            0 => (table.first_col, table.last_col),
            _ => {
                let mut edges = Vec::new();
                for one in &wanted {
                    for side in one.split(':') {
                        let heading = side.trim().trim_start_matches('[').trim_end_matches(']');
                        match table.column(heading) {
                            Some(col) => edges.push(col),
                            None => return Err(ExcelError::Name),
                        }
                    }
                }
                (
                    *edges.iter().min().unwrap_or(&table.first_col),
                    *edges.iter().max().unwrap_or(&table.last_col),
                )
            }
        };

        // Which rows.
        let body_first = table.first_row + table.header_rows;
        let (first_row, last_row) = match part {
            TablePart::Headers => (table.first_row, table.first_row + table.header_rows - 1),
            TablePart::All => (table.first_row, table.last_row),
            TablePart::Totals => (table.last_row, table.last_row),
            TablePart::Body => (body_first, table.last_row),
            TablePart::ThisRow => match at {
                // Outside the table is `#VALUE!` in Excel, which is what a
                // stray `[#This Row]` deserves.
                Some((_, row)) if row >= table.first_row && row <= table.last_row => (row, row),
                _ => return Err(ExcelError::Value),
            },
        };
        if first_row > last_row || first_col > last_col {
            return Err(ExcelError::Ref);
        }
        Ok((
            table.sheet.clone(),
            RangeRef::normalised(
                CellRef::new(first_col, first_row),
                CellRef::new(last_col, last_row),
            ),
        ))
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
        // `ROW(tbl[[#Headers],[ID]])` asks where the heading is, which is a
        // reference like any other once the table has been looked up.
        let range = &match first {
            Expr::Ref(reference) => reference.range,
            Expr::Table { name, asked } => match self.table_range(name, asked, at) {
                Ok((_, range)) => range,
                Err(why) => return Arg::Value(Value::Error(why)),
            },
            _ => return Arg::Value(Value::Error(ExcelError::Value)),
        };
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

    /// Every value in `range`, as a block.
    ///
    /// A whole-column reference names 1,048,576 cells, and building all of
    /// them to add up the nine that are there is a way to run out of memory
    /// rather than a way to answer. So a range reaching past what the sheet
    /// holds is cut back to what it holds — which is what makes `SUM(A:A)` an
    /// ordinary sum.
    ///
    /// The shape it comes back as is therefore the sheet's, not the
    /// reference's, and `ROWS(A:A)` can see the difference: Excel says
    /// 1,048,576 and this says however many rows the sheet has. That is a
    /// wrong answer to a rare question in exchange for a right answer to a
    /// common one.
    fn materialise(&self, sheet: &str, range: &RangeRef, skip_subtotals: bool) -> RangeData {
        let range = &self.cut_to_fit(sheet, range);
        RangeData::from_range(range, |col, row| {
            if skip_subtotals && self.is_subtotal_cell(sheet, col, row) {
                Value::Blank
            } else {
                self.value_at(sheet, col, row)
            }
        })
    }

    /// `range`, with any part reaching past the sheet's own contents removed.
    ///
    /// Only ever cuts back, never extends: a range wholly inside the sheet
    /// comes back untouched, and a range naming a sheet that holds nothing
    /// comes back as one cell rather than as nothing, since a block of no
    /// cells is not a shape anything else here knows what to do with.
    fn cut_to_fit(&self, sheet: &str, range: &RangeRef) -> RangeRef {
        let Some(held) = self.sheets.get(sheet) else {
            return *range;
        };
        // Only worth the walk when the range actually reaches past the end.
        // A whole column is the case this is for and it is unmistakable.
        if range.end.row < MAX_ROW && range.end.col < MAX_COL {
            return *range;
        }
        let (mut last_col, mut last_row) = (0u32, 0u32);
        for (col, row) in held.cells.keys() {
            last_col = last_col.max(*col);
            last_row = last_row.max(*row);
        }
        let mut cut = *range;
        cut.end.row = cut.end.row.min(last_row.max(cut.start.row));
        cut.end.col = cut.end.col.min(last_col.max(cut.start.col));
        cut
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

    /// A table on sheet S: a heading row and three rows under it.
    fn a_table() -> Workbook {
        let mut wb = book();
        for (at, heading) in ["ID", "NAME", "PAY"].iter().enumerate() {
            wb.set_value(
                "Sheet1",
                &format!("{}1", (b'A' + at as u8) as char),
                Value::text(*heading),
            )
            .unwrap();
        }
        for row in 2..=4 {
            wb.set_value("Sheet1", &format!("A{row}"), Value::Number(row as f64 - 1.0))
                .unwrap();
            wb.set_value("Sheet1", &format!("B{row}"), Value::text(format!("w{row}")))
                .unwrap();
            wb.set_value("Sheet1", &format!("C{row}"), Value::Number(row as f64 * 100.0))
                .unwrap();
        }
        wb.add_table(
            "Sheet1",
            "tbl",
            (0, 3),
            (0, 2),
            1,
            vec!["ID".into(), "NAME".into(), "PAY".into()],
        );
        wb
    }

    /// Three rows of three, so a whole column and a whole row are told apart
    /// by their sums.
    fn a_block() -> Workbook {
        let mut wb = book();
        for row in 1..=3u32 {
            for (at, col) in ["A", "B", "C"].iter().enumerate() {
                wb.set_value(
                    "Sheet1",
                    &format!("{col}{row}"),
                    Value::Number(((row - 1) * 3 + at as u32 + 1) as f64),
                )
                .unwrap();
            }
        }
        wb
    }

    #[test]
    fn an_index_that_leaves_out_a_row_is_the_whole_column() {
        // An omission and an explicit zero ask the same thing, and a zero in
        // the other place asks for a row instead. Both zeros is everything.
        let mut wb = a_block();
        wb.set_formula("Sheet1", "E1", "=SUM(INDEX(A1:C3,,2))").unwrap();
        wb.set_formula("Sheet1", "E2", "=SUM(INDEX(A1:C3,0,2))").unwrap();
        wb.set_formula("Sheet1", "E3", "=SUM(INDEX(A1:C3,2,0))").unwrap();
        wb.set_formula("Sheet1", "E4", "=SUM(INDEX(A1:C3,0,0))").unwrap();
        wb.recalculate();
        assert_eq!(wb.value("Sheet1", "E1"), Value::Number(15.0), "2 + 5 + 8");
        assert_eq!(wb.value("Sheet1", "E2"), Value::Number(15.0));
        assert_eq!(wb.value("Sheet1", "E3"), Value::Number(15.0), "4 + 5 + 6");
        assert_eq!(wb.value("Sheet1", "E4"), Value::Number(45.0));
    }

    #[test]
    fn a_whole_line_index_is_an_array_the_rest_can_work_on() {
        // The point of returning a line rather than a cell: it is compared
        // against a value to make a column of trues, and those pick out row
        // numbers. `SMALL(IF(INDEX(range,,n)=x, ROW(range)), k)` — the k-th row
        // that matches — is how the corpus asks its commonest question, and
        // every part of it needs the array.
        let mut wb = a_block();
        wb.set_formula("Sheet1", "E1", "=SUM(IF(INDEX(A1:C3,,2)>2,1,0))").unwrap();
        wb.set_formula("Sheet1", "E2", "=SMALL(IF(INDEX(A1:C3,,2)>2,ROW(A1:A3)),1)")
            .unwrap();
        wb.recalculate();
        assert_eq!(wb.value("Sheet1", "E1"), Value::Number(2.0), "5 and 8");
        assert_eq!(wb.value("Sheet1", "E2"), Value::Number(2.0), "the first is row 2");
    }

    #[test]
    fn an_index_still_addresses_one_cell_when_it_is_asked_to() {
        let mut wb = a_block();
        wb.set_formula("Sheet1", "E1", "=INDEX(A1:C3,2,3)").unwrap();
        // One index into a single column is still counted along it.
        wb.set_formula("Sheet1", "E2", "=INDEX(A1:A3,2)").unwrap();
        // A line outside the range is the ordinary refusal.
        wb.set_formula("Sheet1", "E3", "=SUM(INDEX(A1:C3,,9))").unwrap();
        wb.recalculate();
        assert_eq!(wb.value("Sheet1", "E1"), Value::Number(6.0));
        assert_eq!(wb.value("Sheet1", "E2"), Value::Number(4.0));
        assert_eq!(wb.value("Sheet1", "E3"), Value::Error(ExcelError::Ref));
    }

    #[test]
    fn a_table_column_is_the_cells_under_its_heading() {
        // `Table1[Description]` means the data, not the heading — which is why
        // the heading row has to be taken off the top rather than the range
        // being used whole.
        let mut wb = a_table();
        wb.set_formula("Sheet1", "E1", "=SUM(tbl[PAY])").unwrap();
        wb.set_formula("Sheet1", "E2", "=COUNT(tbl[])").unwrap();
        wb.recalculate();
        assert_eq!(wb.value("Sheet1", "E1"), Value::Number(900.0));
        // Six numbers in the body: three IDs and three PAYs. The names are not
        // numbers and the heading row is not the body.
        assert_eq!(wb.value("Sheet1", "E2"), Value::Number(6.0));
    }

    #[test]
    fn this_row_means_the_row_that_is_asking() {
        let mut wb = a_table();
        wb.set_formula("Sheet1", "E3", "=tbl[[#This Row],[PAY]]").unwrap();
        wb.set_formula("Sheet1", "E4", "=tbl[[#This Row],[PAY]]").unwrap();
        // Outside the table there is no such row, and Excel says so.
        wb.set_formula("Sheet1", "E9", "=tbl[[#This Row],[PAY]]").unwrap();
        wb.recalculate();
        assert_eq!(wb.value("Sheet1", "E3"), Value::Number(300.0));
        assert_eq!(wb.value("Sheet1", "E4"), Value::Number(400.0));
        assert_eq!(wb.value("Sheet1", "E9"), Value::Error(ExcelError::Value));
    }

    #[test]
    fn the_headers_can_be_named_too() {
        let mut wb = a_table();
        wb.set_formula("Sheet1", "E1", "=tbl[[#Headers],[PAY]]").unwrap();
        // `ROW(tbl[[#Headers],[ID]])` asks WHERE the heading is, not what it
        // says, which is a different question of the same reference.
        wb.set_formula("Sheet1", "E2", "=ROW(tbl[[#Headers],[ID]])").unwrap();
        wb.recalculate();
        assert_eq!(wb.value("Sheet1", "E1"), Value::text("PAY"));
        assert_eq!(wb.value("Sheet1", "E2"), Value::Number(1.0));
    }

    #[test]
    fn a_span_runs_from_one_column_to_another() {
        // `tblEmpleados[[NOMBRE]:[FECHA INGRESO]]` is a VLOOKUP table written
        // by the names of its ends. The colon inside the brackets means
        // columns, not cells, which is why the whole bracket group is kept out
        // of the ordinary parser's way.
        let mut wb = a_table();
        wb.set_formula("Sheet1", "E1", "=COUNT(tbl[[ID]:[PAY]])").unwrap();
        wb.set_formula("Sheet1", "E2", "=SUM(tbl[[ID]:[ID]])").unwrap();
        wb.recalculate();
        assert_eq!(wb.value("Sheet1", "E1"), Value::Number(6.0));
        assert_eq!(wb.value("Sheet1", "E2"), Value::Number(6.0), "1 + 2 + 3");
    }

    #[test]
    fn a_name_nobody_gave_is_not_a_table() {
        let mut wb = a_table();
        wb.set_formula("Sheet1", "E1", "=SUM(tbl[NOPE])").unwrap();
        wb.set_formula("Sheet1", "E2", "=SUM(other[PAY])").unwrap();
        wb.recalculate();
        assert_eq!(wb.value("Sheet1", "E1"), Value::Error(ExcelError::Name));
        assert_eq!(wb.value("Sheet1", "E2"), Value::Error(ExcelError::Name));
    }

    #[test]
    fn a_table_is_named_without_regard_to_capitals() {
        let mut wb = a_table();
        wb.set_formula("Sheet1", "E1", "=SUM(TBL[pay])").unwrap();
        wb.recalculate();
        assert_eq!(wb.value("Sheet1", "E1"), Value::Number(900.0));
    }

    #[test]
    fn a_whole_column_is_every_cell_in_it() {
        let mut wb = book();
        for at in 1..=3 {
            wb.set_value("Sheet1", &format!("A{at}"), Value::Number(at as f64))
                .unwrap();
            wb.set_value("Sheet1", &format!("B{at}"), Value::Number(at as f64 * 10.0))
                .unwrap();
        }
        // Kept clear of rows 1 to 3 and columns A and B: an answer that sits
        // inside the range it is summing is part of its own sum, and the first
        // version of this put `=SUM(1:1)` in D1 and then wondered why row one
        // came to more than row one.
        wb.set_formula("Sheet1", "E5", "=SUM(A:A)").unwrap();
        wb.set_formula("Sheet1", "E6", "=SUM($A:$A)").unwrap();
        wb.set_formula("Sheet1", "E7", "=SUM(A:B)").unwrap();
        wb.set_formula("Sheet1", "E8", "=SUM(1:1)").unwrap();
        wb.recalculate();
        assert_eq!(wb.value("Sheet1", "E5"), Value::Number(6.0));
        assert_eq!(wb.value("Sheet1", "E6"), Value::Number(6.0), "dollars mean nothing here");
        assert_eq!(wb.value("Sheet1", "E7"), Value::Number(66.0));
        assert_eq!(wb.value("Sheet1", "E8"), Value::Number(11.0), "a whole row");
    }

    #[test]
    fn a_whole_column_keeps_the_sheet_it_was_given() {
        // The first attempt read the range off the parsed atoms, by which
        // point the sheet was gone: `Data!$D` becomes a bare name and the
        // qualifier is dropped on purpose, a defined name belonging to the
        // workbook rather than to any sheet. So `SUMIFS(Data!$D:$D, ...)` was
        // built against whatever sheet the formula sat on, found nothing, and
        // answered a confident nought — worse than the parse error it
        // replaced, because a formula that will not parse at least keeps the
        // value the file was saved with.
        let mut wb = book();
        wb.add_sheet("Data");
        for at in 1..=3 {
            wb.set_value("Data", &format!("A{at}"), Value::Number(at as f64))
                .unwrap();
            wb.set_value("Data", &format!("D{at}"), Value::Number(at as f64 * 100.0))
                .unwrap();
        }
        // Something else entirely on the sheet the formula lives on, so a
        // range that lost its sheet would be visibly wrong rather than empty.
        wb.set_value("Sheet1", "A1", Value::Number(99.0)).unwrap();
        wb.set_formula("Sheet1", "C1", "=SUM(Data!A:A)").unwrap();
        wb.set_formula("Sheet1", "C2", "=SUM(Data!$D:$D)").unwrap();
        wb.set_formula("Sheet1", "C3", "=SUMIFS(Data!$D:$D,Data!$A:$A,2)")
            .unwrap();
        wb.set_formula("Sheet1", "C4", "=SUM(Data!1:1)").unwrap();
        wb.set_formula("Sheet1", "C5", "=SUM(A:A)").unwrap();
        wb.recalculate();
        assert_eq!(wb.value("Sheet1", "C1"), Value::Number(6.0));
        assert_eq!(wb.value("Sheet1", "C2"), Value::Number(600.0));
        assert_eq!(wb.value("Sheet1", "C3"), Value::Number(200.0));
        assert_eq!(wb.value("Sheet1", "C4"), Value::Number(101.0), "a qualified row");
        assert_eq!(
            wb.value("Sheet1", "C5"),
            Value::Number(99.0),
            "and an unqualified one is still this sheet's",
        );
    }

    #[test]
    fn a_whole_column_costs_what_the_sheet_holds_not_what_it_could() {
        // A column names 1,048,576 cells. Building all of them to add up the
        // three that are there is a way to run out of memory rather than a way
        // to answer, so a range reaching past the sheet is cut back to it.
        let mut wb = book();
        for at in 1..=3 {
            wb.set_value("Sheet1", &format!("A{at}"), Value::Number(1.0))
                .unwrap();
        }
        wb.set_formula("Sheet1", "C1", "=COUNT(A:A)").unwrap();
        let started = std::time::Instant::now();
        wb.recalculate();
        assert_eq!(wb.value("Sheet1", "C1"), Value::Number(3.0));
        assert!(
            started.elapsed().as_millis() < 500,
            "a whole column took {:?}, which means it was materialised in full",
            started.elapsed(),
        );
    }

    #[test]
    fn a_name_may_stand_for_a_whole_column() {
        // Which is what several of the blind corpus's names do — `cats` is
        // `'Exp-DB'!$B:$B` — and why reading the names was not enough on its
        // own.
        let mut wb = book();
        for at in 1..=3 {
            wb.set_value("Sheet1", &format!("B{at}"), Value::Number(at as f64))
                .unwrap();
        }
        wb.define_name("col", "Sheet1!$B:$B").expect("the name parses");
        wb.set_formula("Sheet1", "D1", "=SUM(col)").unwrap();
        wb.recalculate();
        assert_eq!(wb.value("Sheet1", "D1"), Value::Number(6.0));
    }

    #[test]
    fn an_empty_criterion_asks_for_the_empty_ones() {
        // `COUNTIFS(B:B, x, D:D, "")` — count where D has nothing in it — is
        // how anyone counts what is still outstanding, and a rule that says a
        // blank matches nothing answers nought to every one of them.
        let mut wb = book();
        for (at, who) in ["ann", "bob", "ann"].iter().enumerate() {
            wb.set_value("Sheet1", &format!("B{}", at + 1), Value::text(*who))
                .unwrap();
        }
        // D2 is filled in; D1 and D3 are not.
        wb.set_value("Sheet1", "D2", Value::text("done")).unwrap();
        wb.set_formula("Sheet1", "F1", "=COUNTIFS(B:B,\"ann\",D:D,\"\")")
            .unwrap();
        wb.set_formula("Sheet1", "F2", "=COUNTIF(D:D,\"\")").unwrap();
        wb.recalculate();
        assert_eq!(wb.value("Sheet1", "F1"), Value::Number(2.0));
        assert_eq!(wb.value("Sheet1", "F2"), Value::Number(2.0));
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
