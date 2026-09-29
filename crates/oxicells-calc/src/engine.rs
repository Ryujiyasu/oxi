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
    totals_rows: u32,
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

/// How much of a workbook a recalculation covers.
enum Extent<'a> {
    /// Every formula there is.
    Everything,
    /// Only what a change to these cells can reach.
    ReachedBy(&'a [(String, (u32, u32))]),
    /// Only these, and nothing they lead to.
    JustThese(&'a [(String, (u32, u32))]),
}

/// Which cell is being worked out, when that is known. `ROW()` with no
/// argument is the only thing that needs it, and a formula evaluated on its
/// own rather than in a cell has no answer for it.
fn parse_range_string(address: &str) -> Option<RangeRef> {
    match address.split_once(':') {
        Some((a, b)) => {
            let start = parse_a1(a.trim())?;
            let end = parse_a1(b.trim())?;
            Some(RangeRef::normalised(start, end))
        }
        None => Some(RangeRef::single(parse_a1(address.trim())?)),
    }
}

type At = Option<(u32, u32)>;

/// What a range read inside SUBTOTAL passes over.
#[derive(Clone, Copy)]
enum Skip {
    Nothing,
    /// Other SUBTOTALs and the filtered rows; with `by_hand`, every hidden row.
    Subtotals { by_hand: bool },
}
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
        /// Where this cell sits in the block its ARRAY formula is dealt
        /// across, as (columns, rows) from the block's top-left; None for a
        /// formula of its own.
        share: Option<(u32, u32)>,
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
    /// Rows out of sight, 0-based, and which of them a filter hid. SUBTOTAL
    /// passes over the filtered ones always and the rest from 101 up.
    hidden_rows: BTreeSet<u32>,
    filtered_rows: BTreeSet<u32>,
    /// The formulas written as dynamic arrays, which spill.
    dynamic: BTreeSet<(u32, u32)>,
    /// Each spilling formula's block, as (width, height), by its cell.
    spills: BTreeMap<(u32, u32), (u32, u32)>,
    /// The values spilled into the cells beside them, with the formula that
    /// spilled each -- None for one remembered from before and not yet
    /// worked out again.
    spilled: BTreeMap<(u32, u32), (Value, Option<(u32, u32)>)>,
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
    /// What NOW answers: the moment this workbook is being worked out, as a
    /// serial with the time of day after the point.
    ///
    /// Read once at the start of a recalculation so that every TODAY in a
    /// workbook agrees with every other — a sheet where one column thought it
    /// was Monday and the next thought it was Tuesday would be worse than one
    /// that could not tell the time at all. `set_now` pins it, which is what
    /// makes a test of a function that means "now" possible.
    now: f64,
    /// True once someone has pinned the moment, so the clock is left alone.
    now_pinned: bool,
    /// What this workbook remembers of the workbooks it links to, by the
    /// number `[1]`, `[2]` … names them and the sheet inside.
    ///
    /// The other workbook is not open — in a browser it never is — so a
    /// formula reading one is answered from the copy the file keeps, which is
    /// the same copy Excel reads from when the source is closed.
    linked: BTreeMap<(u32, String), BTreeMap<(u32, u32), Value>>,
    /// The names a LET is binding while its last argument is worked out,
    /// innermost last. Asked before the workbook's own names.
    let_scope: std::cell::RefCell<Vec<(String, Arg)>>,
    /// The state RAND draws from: seeded from the moment on first use, and
    /// moved on by every draw.
    rand_state: std::cell::Cell<u64>,
    /// The sheets in the order the workbook has them, which SHEET counts by.
    sheet_order: Vec<String>,
}

/// `range`, with any part reaching past what is remembered of a linked sheet
/// removed. Without this a whole-column reference into a link would ask for a
/// million rows of nothing.
fn cut_to_remembered(range: &RangeRef, cells: &BTreeMap<(u32, u32), Value>) -> RangeRef {
    let Some(last_col) = cells.keys().map(|(col, _)| *col).max() else {
        return *range;
    };
    let Some(last_row) = cells.keys().map(|(_, row)| *row).max() else {
        return *range;
    };
    RangeRef::normalised(
        CellRef::new(range.start.col, range.start.row),
        CellRef::new(range.end.col.min(last_col), range.end.row.min(last_row)),
    )
}

impl Workbook {
    /// Remember one cell of a linked workbook: which link, which sheet of it,
    /// and where in that sheet, counting a column and a row from zero as the
    /// rest of the engine does.
    pub fn add_linked_cell(&mut self, book: u32, sheet: &str, col: u32, row: u32, value: Value) {
        self.linked
            .entry((book, sheet.to_string()))
            .or_default()
            .insert((col, row), value);
    }

    /// Whether anything is remembered of a link at all. A link Excel could not
    /// refresh leaves an empty copy behind, and a formula reading one has no
    /// answer to be had — better to leave such a cell showing what it was
    /// saved with than to overwrite it with `#REF!`.
    pub fn remembers_link(&self, book: u32) -> bool {
        self.linked.keys().any(|(at, _)| *at == book)
    }

    pub fn new() -> Workbook {
        Workbook::default()
    }

    pub fn add_sheet(&mut self, name: &str) {
        self.sheets.entry(name.to_string()).or_default();
        if !self.sheet_order.iter().any(|held| held == name) {
            self.sheet_order.push(name.to_string());
        }
    }

    /// Say that a row (0-based) is hidden, and whether a filter hid it.
    pub fn hide_row(&mut self, sheet: &str, row: u32, by_filter: bool) {
        let held = self.sheets.entry(sheet.to_string()).or_default();
        held.hidden_rows.insert(row);
        if by_filter {
            held.filtered_rows.insert(row);
        }
    }

    /// Fix what NOW and TODAY answer, instead of asking the clock.
    ///
    /// `serial` is counted the way every other date here is: whole days since
    /// the last day of 1899, with the time of day after the point.
    pub fn set_now(&mut self, serial: f64) {
        self.now = serial;
        self.now_pinned = true;
    }

    /// The moment by the system clock, as a serial, where there is one.
    ///
    /// There is not always one: `wasm32-unknown-unknown` has no clock at all,
    /// and asking it for the time does not fail politely — it panics, which
    /// took down the browser editor on the first thing anyone typed. So this
    /// answers `None` there and the host says instead. Anything that knows it
    /// is running in a browser knows what day it is.
    fn read_the_clock() -> Option<f64> {
        #[cfg(target_arch = "wasm32")]
        {
            None
        }
        #[cfg(not(target_arch = "wasm32"))]
        {
            use std::time::{SystemTime, UNIX_EPOCH};
            let since = SystemTime::now()
                .duration_since(UNIX_EPOCH)
                .map(|held| held.as_secs_f64())
                .unwrap_or(0.0);
            // 25,569 is 1970-01-01 counted from the last day of 1899.
            Some(25_569.0 + since / 86_400.0)
        }
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
        totals_rows: u32,
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
                totals_rows,
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
                share: None,
            },
        );
        Ok(())
    }

    /// Give a formula cell the answer it already holds, so that a formula
    /// worked out on its own -- `evaluate` -- reads it without the whole
    /// book being worked out first. A cell holding no formula is left alone.
    pub fn set_cached(&mut self, sheet: &str, a1: &str, value: Value) -> Result<(), CalcError> {
        let cell = self.cell_ref(a1)?;
        if let Some(Cell::Formula { cached, .. }) = self.sheet_mut(sheet)?.cells.get_mut(&cell.coord()) {
            *cached = value;
        }
        Ok(())
    }

    /// Mark a formula as a dynamic array, whose block of answers spills into
    /// the cells beside it.
    pub fn set_dynamic(&mut self, sheet: &str, a1: &str) -> Result<(), CalcError> {
        let cell = self.cell_ref(a1)?;
        self.sheet_mut(sheet)?.dynamic.insert(cell.coord());
        Ok(())
    }

    /// A value a spill left in a cell, remembered from before so that a
    /// formula worked out on its own can read it. A recalculation works the
    /// spills out afresh.
    pub fn set_spilled_value(&mut self, sheet: &str, a1: &str, value: Value) -> Result<(), CalcError> {
        let cell = self.cell_ref(a1)?;
        self.sheet_mut(sheet)?.spilled.insert(cell.coord(), (value, None));
        Ok(())
    }

    /// Every spill: its sheet, its formula's cell and its (width, height).
    pub fn spill_blocks(&self) -> Vec<(String, (u32, u32), (u32, u32))> {
        self.sheets
            .iter()
            .flat_map(|(name, sheet)| sheet.spills.iter().map(move |(at, size)| (name.clone(), *at, *size)))
            .collect()
    }

    /// The values spilled into a sheet's cells, by (col, row).
    pub fn spilled_values(&self, sheet: &str) -> Vec<((u32, u32), Value)> {
        self.sheets.get(sheet).map_or_else(Vec::new, |held| {
            held.spilled
                .iter()
                .filter(|(_, (_, owner))| owner.is_some())
                .map(|(at, (value, _))| (*at, value.clone()))
                .collect()
        })
    }

    /// The block a spilling formula's cell spilled, as a range.
    fn spill_range(&self, sheet: &str, anchor: CellRef) -> Option<RangeRef> {
        let (width, height) = *self.sheets.get(sheet)?.spills.get(&(anchor.col, anchor.row))?;
        Some(RangeRef::normalised(anchor, CellRef::new(anchor.col + width - 1, anchor.row + height - 1)))
    }

    /// One cell of an ARRAY formula: the formula is worked out once as a
    /// block and this cell shows the element `offset` (columns, rows) into
    /// it. Every member is given the whole formula, which is what Excel
    /// answers for `Formula` on any of them.
    ///
    /// A block with one column is dealt down every column of the cells, and
    /// one with one row across every row -- a scalar reaches them all -- and
    /// a cell past a block that has more than one is `#N/A`: measured,
    /// `=A1:A2*2` dealt across three cells gives the third `#N/A`.
    pub fn set_array_member(
        &mut self,
        sheet: &str,
        a1: &str,
        formula: &str,
        offset: (u32, u32),
    ) -> Result<(), CalcError> {
        let cell = self.cell_ref(a1)?;
        let expr = parse(formula)?;
        self.sheet_mut(sheet)?.cells.insert(
            cell.coord(),
            Cell::Formula {
                source: formula.to_string(),
                expr,
                cached: Value::Blank,
                share: Some(offset),
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

    /// The same, keeping a block of answers whole: `{1,2,3}*2` is three
    /// numbers, as (width, height, cells row by row).
    pub fn evaluate_block(&self, sheet: &str, formula: &str) -> Result<(usize, usize, Vec<Value>), CalcError> {
        let expr = parse(formula)?;
        Ok(match self.eval_arg(&expr, sheet, 0, None) {
            Arg::Range(block) => (block.width, block.height, block.cells.into_iter().map(formula_result).collect()),
            Arg::Value(value) => (1, 1, vec![formula_result(value)]),
        })
    }

    /// Evaluate a formula as though it stood in one particular cell.
    ///
    /// A conditional rule's formula is one formula written for the top-left of
    /// its range and read at every cell in it, and the commonest of them all —
    /// `MOD(ROW(),2)=0`, banded rows — asks where it is standing. Evaluated
    /// nowhere that answers `#VALUE!` and the rule catches nothing, so the
    /// caller says where.
    ///
    /// `at` is `(column, row)`, counted from zero as everything here counts.
    pub fn evaluate_at(
        &self,
        sheet: &str,
        formula: &str,
        at: (u32, u32),
    ) -> Result<Value, CalcError> {
        let expr = parse(formula)?;
        Ok(formula_result(self.eval(&expr, sheet, 0, Some(at))))
    }

    // -- recalculation ----------------------------------------------------

    /// Recalculate every formula in dependency order.
    pub fn recalculate(&mut self) -> RecalcReport {
        // What was remembered of the spills goes: every one is worked out
        // afresh.
        for sheet in self.sheets.values_mut() {
            sheet.spilled.retain(|_, (_, owner)| owner.is_some());
        }
        let mut report = self.work_out(Extent::Everything);
        // A formula reading a cell a spill fills is waiting on the spilling
        // formula, which the order can only know once the spill is known --
        // so go round again while the spills are still changing.
        let mut layout = self.spill_blocks();
        for _ in 0..3 {
            if layout.is_empty() {
                break;
            }
            report = self.work_out(Extent::Everything);
            let now = self.spill_blocks();
            if now == layout {
                break;
            }
            layout = now;
        }
        report
    }

    /// Recalculate only the formulas a change to these cells can reach.
    ///
    /// A change reaches whatever reads the cell, and whatever reads that, and
    /// nothing else — which for an ordinary edit is a handful of cells out of
    /// however many the workbook holds. Each entry is a sheet name and a
    /// (column, row) counted from zero, as everything here counts.
    ///
    /// The order is the same order a full recalculation would use: the whole
    /// graph is still built, since a cell that has just been given a formula
    /// may read anything. What is skipped is the working out.
    pub fn recalculate_after(&mut self, changed: &[(String, (u32, u32))]) -> RecalcReport {
        self.work_out(Extent::ReachedBy(changed))
    }

    /// Work out these formulas and no others, in the order a whole pass would
    /// have used.
    ///
    /// This is for filling gaps rather than following a change: whatever these
    /// read is either already known or is another of them, and a cell that is
    /// waited for comes earlier in that order.
    pub fn recalculate_these(&mut self, wanted: &[(String, (u32, u32))]) -> RecalcReport {
        self.work_out(Extent::JustThese(wanted))
    }

    /// The whole of it, or some part.
    fn work_out(&mut self, extent: Extent<'_>) -> RecalcReport {
        if !self.now_pinned {
            if let Some(moment) = Workbook::read_the_clock() {
                self.now = moment;
            }
        }
        let keys = self.formula_keys();
        // Where the formulas are, sheet by sheet, ordered by column and then
        // row. A range can ask this directly for the formulas it covers; the
        // alternative is walking every cell of the sheet once per range, which
        // is what made a 22,864-formula workbook take 37 seconds.
        let mut at: BTreeMap<&str, BTreeMap<(u32, u32), usize>> = BTreeMap::new();
        for (i, (sheet, coord)) in keys.iter().enumerate() {
            at.entry(sheet.as_str()).or_default().insert(*coord, i);
        }

        // Edge dep -> dependent, plus in-degree, for Kahn's algorithm.
        let mut dependents: Vec<BTreeSet<usize>> = vec![BTreeSet::new(); keys.len()];
        let mut indegree = vec![0usize; keys.len()];

        // A formula that reads its own cell -- `=SUM(5:6)` sitting in row 5
        // -- is a cycle of one: measured, Excel leaves it at 0.
        let mut reads_itself = BTreeSet::new();
        for (i, (sheet, coord)) in keys.iter().enumerate() {
            self.each_dependency(sheet, coord, &at, &mut |j| {
                if j == i {
                    reads_itself.insert(i);
                } else if dependents[j].insert(i) {
                    indegree[i] += 1;
                }
            });
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

        let emitted: BTreeSet<usize> = order.iter().copied().collect();

        match extent {
            Extent::Everything => {}
            // Which of them a change can reach.
            Extent::ReachedBy(changed) => {
                let reached = self.reached_by(changed, &keys, &at, &dependents);
                order.retain(|i| reached.contains(i));
            }
            // Just these, wherever they fall in the order.
            Extent::JustThese(wanted) => {
                let asked: BTreeSet<usize> = wanted
                    .iter()
                    .filter_map(|(sheet, coord)| {
                        at.get(sheet.as_str()).and_then(|held| held.get(coord)).copied()
                    })
                    .collect();
                order.retain(|i| asked.contains(i));
            }
        }

        let mut report = RecalcReport {
            evaluated: order.len(),
            circular: Vec::new(),
        };

        for &i in &order {
            let (sheet, coord) = &keys[i];
            if reads_itself.contains(&i) {
                self.store_cached(sheet, coord, Value::Number(0.0));
                continue;
            }
            let Some(expr) = self.expr_at(sheet, coord) else {
                continue;
            };
            let dynamic = self.sheets.get(sheet.as_str()).is_some_and(|held| held.dynamic.contains(coord));
            let value = match self.share_at(sheet, coord) {
                Some(offset) => {
                    formula_result(element_of(self.eval_arg(&expr, sheet, 0, Some(*coord)), offset))
                }
                None if dynamic => self.spill_out(sheet, *coord, &expr),
                None => formula_result(self.eval(&expr, sheet, 0, Some(*coord))),
            };
            self.store_cached(sheet, coord, value);
        }

        // Anything Kahn could not emit sits in a cycle. A partial pass says
        // nothing about the cells it did not look at, so it leaves them be.
        if !matches!(extent, Extent::Everything) {
            return report;
        }
        for (i, (sheet, coord)) in keys.iter().enumerate() {
            if reads_itself.contains(&i) {
                report
                    .circular
                    .push(format!("{}!{}", sheet, CellRef::new(coord.0, coord.1).to_a1()));
            }
            if !emitted.contains(&i) {
                report
                    .circular
                    .push(format!("{}!{}", sheet, CellRef::new(coord.0, coord.1).to_a1()));
                self.store_cached(sheet, coord, Value::Number(0.0));
            }
        }

        report
    }

    /// Work out a dynamic array and spill its block into the cells beside
    /// it, answering what its own cell shows. Measured in Excel: `=A1:A3*2`
    /// in D1 fills D1:D3; a cell in the way holding anything makes it
    /// #SPILL! and fills nothing; a block of one is simply that value.
    fn spill_out(&mut self, sheet: &str, anchor: (u32, u32), expr: &Expr) -> Value {
        if let Some(held) = self.sheets.get_mut(sheet) {
            held.spills.remove(&anchor);
            held.spilled.retain(|_, (_, owner)| *owner != Some(anchor));
        }
        let worked = self.eval_arg(expr, sheet, 0, Some(anchor));
        let block = match worked {
            Arg::Range(block) if block.cells.len() > 1 => block,
            Arg::Range(block) => return formula_result(block.cells.first().cloned().unwrap_or(Value::Blank)),
            Arg::Value(value) => return formula_result(value),
        };
        let (width, height) = (block.width as u32, block.height as u32);
        let (col, row) = anchor;
        if col + width - 1 > MAX_COL || row + height - 1 > MAX_ROW {
            return Value::Error(ExcelError::Spill);
        }
        let Some(held) = self.sheets.get_mut(sheet) else {
            return Value::Error(ExcelError::Ref);
        };
        for down in 0..height {
            for across in 0..width {
                let at = (col + across, row + down);
                if at == anchor {
                    continue;
                }
                let in_the_way = match held.cells.get(&at) {
                    Some(Cell::Literal(Value::Blank)) | None => false,
                    Some(_) => true,
                };
                let taken = held.spilled.get(&at).is_some_and(|(_, owner)| owner.is_some());
                if in_the_way || taken {
                    return Value::Error(ExcelError::Spill);
                }
            }
        }
        for down in 0..height {
            for across in 0..width {
                let at = (col + across, row + down);
                if at == anchor {
                    continue;
                }
                let value = block.cells[(down * width + across) as usize].clone();
                held.spilled.insert(at, (value, Some(anchor)));
            }
        }
        held.spills.insert(anchor, (width, height));
        formula_result(block.cells[0].clone())
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

    /// Every formula a change to `changed` can reach, itself included.
    ///
    /// A formula is reached when it READS one of the changed cells — which is
    /// asked of its references, since a changed cell need not be a formula and
    /// so need not be in the graph at all — or when it reads something already
    /// reached.
    fn reached_by(
        &self,
        changed: &[(String, (u32, u32))],
        keys: &[(String, (u32, u32))],
        at: &BTreeMap<&str, BTreeMap<(u32, u32), usize>>,
        dependents: &[BTreeSet<usize>],
    ) -> BTreeSet<usize> {
        let mut reached = BTreeSet::new();
        let mut walking: VecDeque<usize> = VecDeque::new();

        // A changed cell that is itself a formula has to be worked out again.
        for (sheet, coord) in changed {
            if let Some(&i) = at.get(sheet.as_str()).and_then(|held| held.get(coord)) {
                if reached.insert(i) {
                    walking.push_back(i);
                }
            }
        }
        // And every formula that reads one of them.
        for (i, (sheet, coord)) in keys.iter().enumerate() {
            if reached.contains(&i) {
                continue;
            }
            let Some(expr) = self.expr_at(sheet, coord) else {
                continue;
            };
            let reads_a_change = expr.value_references().iter().any(|reference| {
                let target = reference.sheet.as_deref().unwrap_or(sheet.as_str());
                changed.iter().any(|(where_, (col, row))| {
                    where_ == target && reference.range.contains(*col, *row)
                })
            });
            if reads_a_change && reached.insert(i) {
                walking.push_back(i);
            }
        }
        // Then outwards, to whatever reads those.
        while let Some(i) = walking.pop_front() {
            for &j in &dependents[i] {
                if reached.insert(j) {
                    walking.push_back(j);
                }
            }
        }
        reached
    }

    /// Hands every formula this one waits for to `found`, as a graph node.
    ///
    /// `at` says where the formulas are. Only they can be waited for — nothing
    /// waits for a literal — so a range asks that rather than the sheet, which
    /// is a few thousand entries instead of a few hundred thousand cells.
    fn each_dependency(
        &self,
        sheet: &str,
        coord: &(u32, u32),
        at: &BTreeMap<&str, BTreeMap<(u32, u32), usize>>,
        found: &mut dyn FnMut(usize),
    ) {
        let Some(expr) = self.expr_at(sheet, coord) else {
            return;
        };
        // Not `references`: a range that is only being measured is not
        // something this cell waits for.
        for reference in expr.value_references() {
            let target = reference.sheet.as_deref().unwrap_or(sheet);
            let Some(formulas) = at.get(target) else {
                continue;
            };
            if reference.range.is_single() {
                if let Some(&j) = formulas.get(&reference.range.start.coord()) {
                    found(j);
                }
                continue;
            }
            // Ordered by column and then row, so the columns the range spans
            // are one contiguous stretch of the map and the rows are checked
            // as they come past.
            let (from, to) = (reference.range.start, reference.range.end);
            for (coord, &j) in formulas.range((from.col, 0)..=(to.col, u32::MAX)) {
                if coord.1 >= from.row && coord.1 <= to.row {
                    found(j);
                }
            }
        }
        // A cell a spill fills is waiting on the formula that spills it.
        for reference in expr.value_references() {
            let target = reference.sheet.as_deref().unwrap_or(sheet);
            let (Some(formulas), Some(held)) = (at.get(target), self.sheets.get(target)) else {
                continue;
            };
            let (from, to) = (reference.range.start, reference.range.end);
            for (anchor, (width, height)) in &held.spills {
                let (last_col, last_row) = (anchor.0 + width - 1, anchor.1 + height - 1);
                let meets = anchor.0 <= to.col && last_col >= from.col && anchor.1 <= to.row && last_row >= from.row;
                if meets {
                    if let Some(&j) = formulas.get(anchor) {
                        found(j);
                    }
                }
            }
        }
    }

    fn expr_at(&self, sheet: &str, coord: &(u32, u32)) -> Option<Expr> {
        match self.sheets.get(sheet)?.cells.get(coord)? {
            Cell::Formula { expr, .. } => Some(expr.clone()),
            Cell::Literal(_) => None,
        }
    }

    /// Where the cell sits in its array formula's block, if it is in one.
    fn share_at(&self, sheet: &str, coord: &(u32, u32)) -> Option<(u32, u32)> {
        match self.sheets.get(sheet)?.cells.get(coord)? {
            Cell::Formula { share, .. } => *share,
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
        let Some(held) = self.sheets.get(sheet) else {
            return Value::Blank;
        };
        match held.cells.get(&(col, row)) {
            Some(Cell::Literal(Value::Blank)) | None => held
                .spilled
                .get(&(col, row))
                .map(|(value, _)| value.clone())
                .unwrap_or(Value::Blank),
            Some(cell) => cell.value().clone(),
        }
    }

    /// A draw in [0, 1), from a xorshift seeded by the moment.
    fn next_random(&self) -> f64 {
        let mut state = self.rand_state.get();
        if state == 0 {
            state = (self.now.to_bits() ^ 0x9E37_79B9_7F4A_7C15) | 1;
        }
        state ^= state << 13;
        state ^= state >> 7;
        state ^= state << 17;
        self.rand_state.set(state);
        (state >> 11) as f64 / (1u64 << 53) as f64
    }

    fn eval(&self, expr: &Expr, sheet: &str, depth: u32, at: At) -> Value {
        let worked = self.eval_arg(expr, sheet, depth, at);
        match &worked {
            // A block that was worked out spills: this cell shows the first of
            // it and the rest fills the cells below, which is why a file holds
            // the formula in that cell alone. A bare range reference is not
            // that — `=A1:A5` in one cell is the old implicit intersection —
            // so it goes on being refused.
            Arg::Range(block) if block.cells.len() > 1 => block
                .cells
                .first()
                .cloned()
                .unwrap_or(Value::Error(ExcelError::NA)),
            _ => worked.scalar(),
        }
    }

    fn eval_arg(&self, expr: &Expr, sheet: &str, depth: u32, at: At) -> Arg {
        self.eval_arg_inner(expr, sheet, depth, Skip::Nothing, at)
    }

    /// `skip_subtotals` drops cells that are themselves `SUBTOTAL` formulas when
    /// materialising a range. See the `SUBTOTAL` arm below for why.
    fn eval_arg_inner(
        &self,
        expr: &Expr,
        sheet: &str,
        depth: u32,
        skip: Skip,
        at: At,
    ) -> Arg {
        if depth > MAX_EVAL_DEPTH {
            return Arg::Value(Value::Error(ExcelError::Num));
        }

        match expr {
            Expr::Literal(v) => Arg::Value(v.clone()),

            // An array written out IS a block of values -- the same thing a
            // range materialises to, with nothing behind it.
            Expr::Array(rows) => {
                let height = rows.len();
                let width = rows.first().map_or(0, |row| row.len());
                Arg::Range(RangeData {
                    width,
                    height,
                    cells: rows.iter().flatten().cloned().collect(),
                })
            }

            Expr::Ref(reference) if reference.book.is_some() => {
                let book = reference.book.unwrap_or_default();
                let Some(name) = reference.sheet.as_deref() else {
                    return Arg::Value(Value::Error(ExcelError::Ref));
                };
                let Some(cells) = self.linked.get(&(book, name.to_string())) else {
                    return Arg::Value(Value::Error(ExcelError::Ref));
                };
                // A linked range is often written as a whole column --
                // `VLOOKUP(C2,[1]Lookup!A:B,2,FALSE)` -- so it is cut to what
                // is actually remembered before being materialised, exactly as
                // a range in this workbook is cut to the sheet's contents.
                let range = cut_to_remembered(&reference.range, cells);
                Arg::Range(RangeData::from_range(&range, |col, row| {
                    cells.get(&(col, row)).cloned().unwrap_or(Value::Blank)
                }))
            }

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
                Arg::Range(self.materialise(target, &reference.range, skip))
            }

            Expr::Table { name, asked } => self.a_table_column(name, asked, sheet, at),

            Expr::Name(name) => {
                let bound_by_let = self
                    .let_scope
                    .borrow()
                    .iter()
                    .rev()
                    .find(|(held, _)| held == name)
                    .map(|(_, value)| value.clone());
                if let Some(value) = bound_by_let {
                    return value;
                }
                match self.names.get(name) {
                    Some(bound) => self.eval_arg(&bound.clone(), sheet, depth + 1, at),
                    // A table's bare name is its data: `COLUMNS(tbl)`.
                    None if self.tables.contains_key(&name.to_uppercase()) => {
                        self.a_table_column(name, "", sheet, at)
                    }
                    None => Arg::Value(Value::Error(ExcelError::Name)),
                }
            }

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

            // CELL(info, [reference]): what Excel says of the first cell.
            // Measured: `CELL("address",B2)` is $B$2 and `CELL("col",C5)` 3.
            Expr::Function { name, args } if name == "CELL" && (1..=2).contains(&args.len()) => {
                let info = match self.eval_arg(&args[0], sheet, depth + 1, at).scalar() {
                    Value::Text(text) => text.to_ascii_lowercase(),
                    Value::Error(why) => return Arg::Value(Value::Error(why)),
                    _ => return Arg::Value(Value::Error(ExcelError::Value)),
                };
                let (target, cell) = match args.get(1) {
                    Some(reference) => match self.reference_of(reference, sheet, depth + 1, at) {
                        Some((target, range)) => (target, range.start),
                        None => return Arg::Value(Value::Error(ExcelError::Value)),
                    },
                    None => match at {
                        Some((col, row)) => (sheet.to_string(), CellRef::new(col, row)),
                        None => return Arg::Value(Value::Error(ExcelError::Value)),
                    },
                };
                let held = self.value_at(&target, cell.col, cell.row);
                Arg::Value(match info.as_str() {
                    "address" => {
                        let mut letters = String::new();
                        let mut left = cell.col + 1;
                        while left > 0 {
                            letters.insert(0, (b'A' + ((left - 1) % 26) as u8) as char);
                            left = (left - 1) / 26;
                        }
                        let local = format!("${letters}${}", cell.row + 1);
                        if target == sheet {
                            Value::Text(local)
                        } else {
                            Value::Text(format!("'{}'!{local}", target.replace('\'', "''")))
                        }
                    }
                    "col" => Value::Number(f64::from(cell.col) + 1.0),
                    "row" => Value::Number(f64::from(cell.row) + 1.0),
                    "contents" => held,
                    "type" => Value::Text(
                        match held {
                            Value::Blank => "b",
                            Value::Text(_) => "l",
                            _ => "v",
                        }
                        .to_string(),
                    ),
                    _ => Value::Error(ExcelError::Value),
                })
            }

            // Whether the argument names cells, rather than what they hold.
            // A table part that is not there -- `[#Totals]` with no totals
            // row -- is not a reference.
            // `@`: of a range, the one cell in line with the asking cell --
            // its row in a column, its column in a row, #VALUE! where there
            // is none; of a block of answers, the first.
            Expr::Function { name, args } if name == "_AT" && args.len() == 1 => {
                if let Some((target, range)) = self.reference_of(&args[0], sheet, depth + 1, at) {
                    let (start, end) = (range.start, range.end);
                    let cell = if start == end {
                        Some(start)
                    } else {
                        match at {
                            Some((col, row)) if start.col == end.col && (start.row..=end.row).contains(&row) => {
                                Some(CellRef::new(start.col, row))
                            }
                            Some((col, row)) if start.row == end.row && (start.col..=end.col).contains(&col) => {
                                Some(CellRef::new(col, start.row))
                            }
                            _ => None,
                        }
                    };
                    return match cell {
                        Some(cell) => Arg::Range(self.materialise(&target, &RangeRef::single(cell), skip)),
                        None => Arg::Value(Value::Error(ExcelError::Value)),
                    };
                }
                match self.eval_arg(&args[0], sheet, depth + 1, at) {
                    Arg::Range(block) => Arg::Value(block.cells.into_iter().next().unwrap_or(Value::Blank)),
                    value => value,
                }
            }

            Expr::Function { name, args } if name == "INDEX" && self.index_range(args, sheet, depth + 1, at).is_some() => {
                let (target, range) = self.index_range(args, sheet, depth + 1, at).expect("just asked");
                Arg::Range(self.materialise(&target, &range, skip))
            }

            // Two references that share no cell: #NULL!, measured.
            Expr::Function { name, .. } if name == "_ISECT" => match self.reference_of(expr, sheet, depth + 1, at) {
                Some((target, range)) => Arg::Range(self.materialise(&target, &range, skip)),
                None => Arg::Value(Value::Error(ExcelError::Null)),
            },

            // `D1#`: all that D1 spilled, or #REF! where it spills nothing.
            Expr::Function { name, .. } if name == "_SPILL" => match self.reference_of(expr, sheet, depth + 1, at) {
                Some((target, range)) => Arg::Range(self.materialise(&target, &range, skip)),
                None => Arg::Value(Value::Error(ExcelError::Ref)),
            },

            Expr::Function { name, args } if name == "ISREF" => match args.as_slice() {
                [only] => Arg::Value(Value::Logical(self.reference_of(only, sheet, depth + 1, at).is_some())),
                _ => Arg::Value(Value::Error(ExcelError::Value)),
            },

            Expr::Function { name, args } if name == "ROW" || name == "COLUMN" => {
                self.which_line(name, args, sheet, at)
            }

            // The clock belongs to the workbook, not to the function library,
            // which is handed values and has no way to ask what day it is.
            Expr::Function { name, args } if matches!(name.as_str(), "TODAY" | "NOW") => {
                if !args.is_empty() {
                    return Arg::Value(Value::Error(ExcelError::Value));
                }
                Arg::Value(Value::Number(if name == "TODAY" {
                    self.now.trunc()
                } else {
                    self.now
                }))
            }

            // How tall or how wide, answered from the reference. Counting the
            // range by evaluating it would pull in every cell it names — and
            // the commonest use writes the formula INTO that block, so the
            // count would be a circular reference where Excel sees no
            // difficulty at all. Only a plain reference is taken this way; an
            // array arriving from somewhere else still goes the ordinary road.
            Expr::Function { name, args }
                if matches!(name.as_str(), "ROWS" | "COLUMNS")
                    && Expr::asks_only_the_shape(name, args) =>
            {
                self.how_many_lines(name, args, sheet, at)
            }

            // OFFSET and INDIRECT hand back a *reference*, computed here where
            // the sheet grid is in reach, then materialised like any other.
            Expr::Function { name, args } if name == "OFFSET" => {
                self.offset_reference(args, sheet, depth, at)
            }
            Expr::Function { name, args } if name == "INDIRECT" => {
                self.indirect_reference(args, sheet, depth, at)
            }

            // LET names what it works out, pair by pair, and answers with its
            // last argument: `LET(x, 2, y, x*3, x+y)` is 8. A later pair may
            // use an earlier name.
            // RAND and RANDBETWEEN draw afresh at every working-out, as
            // Excel's do; a browser has no other source that is not the clock.
            // RANDARRAY([rows], [columns], [min], [max], [whole]): measured,
            // RANDARRAY(2,2,1,1,TRUE) is a block of ones.
            Expr::Function { name, args } if name == "RANDARRAY" => {
                let given = |index: usize, default: f64| match args.get(index) {
                    Some(expr) => match self.eval_arg_inner(expr, sheet, depth + 1, skip, at).scalar() {
                        Value::Blank => Ok(default),
                        value => value.to_number(),
                    },
                    None => Ok(default),
                };
                let whole = match args.get(4) {
                    Some(expr) => self.eval_arg_inner(expr, sheet, depth + 1, skip, at).scalar().to_logical().unwrap_or(false),
                    None => false,
                };
                let (rows, cols, low, high) = match (given(0, 1.0), given(1, 1.0), given(2, 0.0), given(3, 1.0)) {
                    (Ok(r), Ok(c), Ok(l), Ok(h)) => (r, c, l, h),
                    (Err(why), ..) | (_, Err(why), ..) | (_, _, Err(why), _) | (_, _, _, Err(why)) => {
                        return Arg::Value(Value::Error(why))
                    }
                };
                if rows < 1.0 || cols < 1.0 || low > high {
                    return Arg::Value(Value::Error(ExcelError::Value));
                }
                let (height, width) = (rows as usize, cols as usize);
                let cells = (0..height * width)
                    .map(|_| {
                        let draw = self.next_random();
                        Value::Number(if whole {
                            low.ceil() + (draw * (high.floor() - low.ceil() + 1.0)).floor()
                        } else {
                            low + draw * (high - low)
                        })
                    })
                    .collect();
                Arg::Range(RangeData { width, height, cells })
            }

            Expr::Function { name, args } if name == "RAND" || name == "RANDBETWEEN" => {
                let draw = self.next_random();
                if name == "RAND" {
                    return if args.is_empty() {
                        Arg::Value(Value::Number(draw))
                    } else {
                        Arg::Value(Value::Error(ExcelError::Value))
                    };
                }
                let [low, high] = args.as_slice() else {
                    return Arg::Value(Value::Error(ExcelError::Value));
                };
                let bound = |expr: &Expr| {
                    self.eval_arg_inner(expr, sheet, depth + 1, skip, at).scalar().to_number()
                };
                match (bound(low), bound(high)) {
                    (Ok(low), Ok(high)) => {
                        // Both bounds go up to a whole number: measured,
                        // RANDBETWEEN(2.5, 2.9) is 3.
                        let (low, high) = (low.ceil(), high.ceil());
                        if low > high {
                            Arg::Value(Value::Error(ExcelError::Num))
                        } else {
                            Arg::Value(Value::Number(low + (draw * (high - low + 1.0)).floor()))
                        }
                    }
                    (Err(why), _) | (_, Err(why)) => Arg::Value(Value::Error(why)),
                }
            }
            // What a formula asks about the workbook rather than the values:
            // measured, SHEETS() is the count of sheets, SHEET() the place of
            // the formula's own sheet, SHEET(Two!A1) and SHEET("Two") that
            // sheet's place (an unknown name #N/A), ISFORMULA of a cell
            // whether it holds one, and AREAS of a plain range 1.
            Expr::Function { name, args }
                if matches!(name.as_str(), "SHEET" | "SHEETS" | "ISFORMULA" | "AREAS") =>
            {
                let place = |named: &str| {
                    self.sheet_order
                        .iter()
                        .position(|held| held.eq_ignore_ascii_case(named))
                        .map(|at| Arg::Value(Value::Number((at + 1) as f64)))
                        .unwrap_or(Arg::Value(Value::Error(ExcelError::NA)))
                };
                match (name.as_str(), args.as_slice()) {
                    ("AREAS", [Expr::Function { name: union, args: areas }]) if union == "_UNION" => {
                        Arg::Value(Value::Number(areas.len() as f64))
                    }
                    ("SHEETS", []) => Arg::Value(Value::Number(self.sheet_order.len().max(1) as f64)),
                    ("SHEETS", [Expr::Ref(_)]) | ("AREAS", [Expr::Ref(_)]) => Arg::Value(Value::Number(1.0)),
                    ("SHEET", []) => place(sheet),
                    ("SHEET", [Expr::Ref(reference)]) => place(reference.sheet.as_deref().unwrap_or(sheet)),
                    ("SHEET", [other]) => {
                        match self.eval_arg_inner(other, sheet, depth + 1, skip, at).scalar() {
                            Value::Text(named) => place(&named),
                            Value::Error(why) => Arg::Value(Value::Error(why)),
                            _ => Arg::Value(Value::Error(ExcelError::NA)),
                        }
                    }
                    ("ISFORMULA", [Expr::Ref(reference)]) => {
                        let target = reference.sheet.as_deref().unwrap_or(sheet);
                        let corner = (
                            reference.range.start.col.min(reference.range.end.col),
                            reference.range.start.row.min(reference.range.end.row),
                        );
                        let held = self
                            .sheets
                            .get(target)
                            .and_then(|held| held.cells.get(&corner))
                            .is_some_and(|cell| matches!(cell, Cell::Formula { .. }));
                        Arg::Value(Value::Logical(held))
                    }
                    _ => Arg::Value(Value::Error(ExcelError::Value)),
                }
            }
            Expr::Function { name, args } if name == "LET" => {
                if args.len() < 3 || args.len() % 2 == 0 {
                    return Arg::Value(Value::Error(ExcelError::Value));
                }
                let mut bound = 0;
                let mut answer = None;
                for pair in args[..args.len() - 1].chunks(2) {
                    // A name given twice is refused: measured, LET(x,1,x,2,x)
                    // is #VALUE!.
                    let duplicate = matches!(&pair[0], Expr::Name(label)
                        if self.let_scope.borrow()[self.let_scope.borrow().len() - bound..]
                            .iter()
                            .any(|(held, _)| held == label));
                    let (Expr::Name(label), false) = (&pair[0], duplicate) else {
                        answer = Some(Arg::Value(Value::Error(ExcelError::Value)));
                        break;
                    };
                    let value = self.eval_arg_inner(&pair[1], sheet, depth + 1, skip, at);
                    self.let_scope.borrow_mut().push((label.clone(), value));
                    bound += 1;
                }
                let answer = answer.unwrap_or_else(|| {
                    self.eval_arg_inner(&args[args.len() - 1], sheet, depth + 1, skip, at)
                });
                let mut scope = self.let_scope.borrow_mut();
                let kept = scope.len() - bound;
                scope.truncate(kept);
                answer
            }
            Expr::Function { name, args } => {
                // SUBTOTAL ignores any cell in its range that is itself a
                // SUBTOTAL, which is how a column of group subtotals can be
                // summed by a grand total without double counting. That cannot
                // be decided from the values alone, so the exclusion has to
                // happen here, while the range is still a reference.
                //
                // It also passes over the rows a filter hid, and from function
                // 101 up the rows hidden by hand as well: measured through VBA,
                // SUBTOTAL(9, ...) under a filter adds only what is showing.
                let skip = if name == "SUBTOTAL" {
                    let kind = args
                        .first()
                        .map(|a| self.eval_arg_inner(a, sheet, depth + 1, Skip::Nothing, at).scalar());
                    let by_hand = matches!(kind, Some(Value::Number(n)) if n >= 100.0);
                    Skip::Subtotals { by_hand }
                } else {
                    Skip::Nothing
                };
                // A reference of several areas, measured against Excel:
                // a function that takes any number of arguments takes each
                // area as one (`SUM((A1:A3,C1:C3))` is `SUM(A1:A3,C1:C3)`);
                // one that reads a list reads every area as one list (LARGE,
                // RANK); INDEX picks the area its fourth argument names; and
                // anything else -- COUNTIF, SUMPRODUCT, a cell of its own --
                // is #VALUE!.
                if name == "_UNION" {
                    return Arg::Value(Value::Error(ExcelError::Value));
                }
                let spliced = matches!(
                    name.as_str(),
                    "SUM" | "COUNT" | "COUNTA" | "AVERAGE" | "AVERAGEA" | "MAX" | "MAXA" | "MIN"
                        | "MINA" | "PRODUCT" | "STDEV" | "STDEV.S" | "STDEV.P" | "STDEVP" | "STDEVA"
                        | "VAR" | "VAR.S" | "VAR.P" | "VARP" | "MEDIAN" | "MODE" | "MODE.SNGL"
                        | "SUMSQ" | "GEOMEAN" | "HARMEAN" | "AVEDEV" | "DEVSQ" | "SUBTOTAL"
                );
                let listed = matches!(
                    name.as_str(),
                    "LARGE" | "SMALL" | "RANK" | "RANK.EQ" | "RANK.AVG" | "PERCENTILE" | "QUARTILE"
                        | "PERCENTILE.INC" | "PERCENTILE.EXC" | "QUARTILE.INC" | "QUARTILE.EXC"
                        | "PERCENTRANK"
                );
                let mut evaluated: Vec<Arg> = Vec::with_capacity(args.len());
                for (at_arg, a) in args.iter().enumerate() {
                    // `Jan:Mar!A1` is that cell on each sheet from Jan to Mar
                    // as they stand in the book, taken as so many areas.
                    let expanded;
                    let a = match a {
                        Expr::Function { name: sheets, args: parts } if sheets == "_SHEETS" => {
                            expanded = self.across_sheets(parts);
                            &expanded
                        }
                        _ => a,
                    };
                    let Expr::Function { name: union, args: areas } = a else {
                        evaluated.push(self.eval_arg_inner(a, sheet, depth + 1, skip, at));
                        continue;
                    };
                    if union != "_UNION" {
                        evaluated.push(self.eval_arg_inner(a, sheet, depth + 1, skip, at));
                        continue;
                    }
                    if spliced {
                        for area in areas {
                            evaluated.push(self.eval_arg_inner(area, sheet, depth + 1, skip, at));
                        }
                    } else if listed {
                        let cells: Vec<Value> = areas
                            .iter()
                            .flat_map(|area| self.eval_arg_inner(area, sheet, depth + 1, skip, at).flatten())
                            .collect();
                        evaluated.push(Arg::Range(RangeData { width: 1, height: cells.len(), cells }));
                    } else if name == "INDEX" && at_arg == 0 {
                        let chosen = args
                            .get(3)
                            .map(|n| self.eval_arg_inner(n, sheet, depth + 1, skip, at).scalar().to_number())
                            .unwrap_or(Ok(1.0));
                        match chosen {
                            Ok(n) if n >= 1.0 && (n as usize) <= areas.len() => evaluated
                                .push(self.eval_arg_inner(&areas[n as usize - 1], sheet, depth + 1, skip, at)),
                            _ => return Arg::Value(Value::Error(ExcelError::Ref)),
                        }
                    } else {
                        return Arg::Value(Value::Error(ExcelError::Value));
                    }
                }
                // INDEX's area number has done its work once the area is picked.
                if name == "INDEX"
                    && matches!(args.first(), Some(Expr::Function { name: union, .. }) if union == "_UNION")
                {
                    evaluated.truncate(3);
                }
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
    fn a_table_column(&self, name: &str, asked: &str, sheet: &str, at: At) -> Arg {
        match self.table_range(name, asked, sheet, at) {
            Ok((sheet, range)) => Arg::Range(self.materialise(&sheet, &range, Skip::Nothing)),
            Err(why) => Arg::Value(Value::Error(why)),
        }
    }

    /// The same, as a reference rather than as its contents.
    ///
    /// `ROW(tbl[[#Headers],[ID]])` asks where the heading IS, not what it
    /// says, so the two callers need different halves of the same answer.
    fn table_range(&self, name: &str, asked: &str, sheet: &str, at: At)
        -> Result<(String, RangeRef), ExcelError>
    {
        // `[@Qty]` names no table: it is the one the asking cell sits in.
        let found = if name.is_empty() {
            at.and_then(|(col, row)| {
                self.tables.values().find(|table| {
                    table.sheet == sheet
                        && (table.first_row..=table.last_row).contains(&row)
                        && (table.first_col..=table.last_col).contains(&col)
                })
            })
        } else {
            self.tables.get(&name.to_uppercase())
        };
        let Some(table) = found else {
            return Err(ExcelError::Name);
        };
        // `[[#This Row],[DATE]]` arrives as `[#This Row],[DATE]`; a lone
        // `[DATE]` arrives as `DATE`. Splitting on the commas OUTSIDE any
        // brackets keeps a heading with a comma in it whole.
        let parts = outside_brackets(asked);
        let mut part = TablePart::Body;
        let mut wanted: Vec<&str> = Vec::new();
        for piece in &parts {
            let mut bare = piece.trim().trim_start_matches('[').trim_end_matches(']');
            // `[@Qty]` and `[@[Unit Price]]` are this row's cell of a column.
            if let Some(column) = bare.strip_prefix('@').filter(|rest| !rest.trim().is_empty()) {
                part = TablePart::ThisRow;
                bare = column.trim().trim_start_matches('[').trim_end_matches(']');
            }
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
            // A table with no totals row has no `[#Totals]`: Excel says #REF!.
            TablePart::Totals if table.totals_rows == 0 => return Err(ExcelError::Ref),
            TablePart::Totals => (table.last_row, table.last_row),
            TablePart::Body => (body_first, table.last_row - table.totals_rows),
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
    fn which_line(&self, name: &str, args: &[Expr], sheet: &str, at: At) -> Arg {
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
            Expr::Table { name, asked } => match self.table_range(name, asked, sheet, at) {
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

    /// `ROWS` and `COLUMNS`, answered from the reference rather than its
    /// contents.
    ///
    /// A whole-column `ROWS(A:A)` is therefore 1,048,576, which is what Excel
    /// says — `materialise` cuts a range back to what the sheet holds, and
    /// counting the cut-down block would answer with the sheet's height
    /// instead of the reference's.
    fn how_many_lines(&self, name: &str, args: &[Expr], sheet: &str, at: At) -> Arg {
        let down = name == "ROWS";
        let range = match args.first() {
            Some(Expr::Ref(reference)) => reference.range,
            Some(Expr::Table { name, asked }) => match self.table_range(name, asked, sheet, at) {
                Ok((_, range)) => range,
                Err(why) => return Arg::Value(Value::Error(why)),
            },
            Some(other) => match self.reference_of(other, sheet, 0, at) {
                Some((_, range)) => range,
                None => return Arg::Value(Value::Error(ExcelError::Value)),
            },
            _ => return Arg::Value(Value::Error(ExcelError::Value)),
        };
        let (from, to) = if down {
            (range.start.row, range.end.row)
        } else {
            (range.start.col, range.end.col)
        };
        Arg::Value(Value::Number((to.saturating_sub(from) + 1) as f64))
    }

    /// The sheet and range a reference-shaped expression names, for OFFSET.
    fn reference_of(&self, expr: &Expr, sheet: &str, depth: u32, at: At) -> Option<(String, RangeRef)> {
        match expr {
            Expr::Ref(reference) if reference.book.is_none() => Some((
                reference.sheet.clone().unwrap_or_else(|| sheet.to_string()),
                reference.range,
            )),
            Expr::Table { name, asked } => self.table_range(name, asked, sheet, at).ok(),
            Expr::Name(name) => match self.names.get(name) {
                Some(bound) => {
                    let bound = bound.clone();
                    self.reference_of(&bound, sheet, depth, at)
                }
                None if self.tables.contains_key(&name.to_uppercase()) => {
                    self.table_range(name, "", sheet, at).ok()
                }
                None => None,
            },
            // A reference can be built by another OFFSET or INDIRECT --
            // `OFFSET(INDIRECT("A1"), 4, 0)` -- so those are followed to their
            // own reference rather than to their materialised values.
            Expr::Function { name, args } if name == "OFFSET" => {
                self.offset_range(args, sheet, depth, at).ok()
            }
            Expr::Function { name, args } if name == "INDIRECT" => {
                self.indirect_range(args, sheet, depth, at).ok()
            }
            Expr::Function { name, args } if name == "INDEX" => self.index_range(args, sheet, depth, at),
            Expr::Function { name, args } if name == "_ISECT" && args.len() == 2 => {
                let (left_sheet, left) = self.reference_of(&args[0], sheet, depth, at)?;
                let (right_sheet, right) = self.reference_of(&args[1], sheet, depth, at)?;
                if !left_sheet.eq_ignore_ascii_case(&right_sheet) {
                    return None;
                }
                let start = CellRef::new(left.start.col.max(right.start.col), left.start.row.max(right.start.row));
                let end = CellRef::new(left.end.col.min(right.end.col), left.end.row.min(right.end.row));
                (start.col <= end.col && start.row <= end.row).then(|| (left_sheet, RangeRef::normalised(start, end)))
            }
            Expr::Function { name, args } if name == "_SPILL" && args.len() == 1 => {
                let (target, range) = self.reference_of(&args[0], sheet, depth, at)?;
                let range = self.spill_range(&target, range.start).or_else(|| {
                    // A formula that is not spilling is its own cell.
                    self.sheets
                        .get(&target)
                        .is_some_and(|held| held.dynamic.contains(&range.start.coord()))
                        .then_some(RangeRef::single(range.start))
                })?;
                Some((target, range))
            }
            _ => None,
        }
    }

    /// `Jan:Mar!A1` as the one reference on each sheet from Jan to Mar, in
    /// the order the book has them now.
    fn across_sheets(&self, parts: &[Expr]) -> Expr {
        let (Some(Expr::Literal(Value::Text(first))), Some(Expr::Ref(reference))) = (parts.first(), parts.get(1)) else {
            return Expr::Literal(Value::Error(ExcelError::Ref));
        };
        let Some(last) = reference.sheet.as_deref() else {
            return Expr::Literal(Value::Error(ExcelError::Ref));
        };
        let place = |wanted: &str| self.sheet_order.iter().position(|held| held.eq_ignore_ascii_case(wanted));
        let (Some(from), Some(to)) = (place(first), place(last)) else {
            return Expr::Literal(Value::Error(ExcelError::Ref));
        };
        let (from, to) = (from.min(to), from.max(to));
        let areas = self.sheet_order[from..=to]
            .iter()
            .map(|sheet| Expr::Ref(crate::Reference { sheet: Some(sheet.clone()), ..reference.clone() }))
            .collect();
        Expr::Function { name: "_UNION".to_string(), args: areas }
    }

    /// `INDEX(reference, row, [col])` as the cells it names: a row or column
    /// of 0, or left empty, is all of them. Measured: `=@INDEX(A1:A3,0)` in
    /// row 1 is A1, and `=@INDEX(A1:B3,2,0)` in column Z is #VALUE!, the
    /// answer being A2:B2. Only a plain range is taken; anything else is
    /// left to the function.
    fn index_range(&self, args: &[Expr], sheet: &str, depth: u32, at: At) -> Option<(String, RangeRef)> {
        if !(2..=3).contains(&args.len()) {
            return None;
        }
        let (target, range) = self.reference_of(&args[0], sheet, depth, at)?;
        let number = |expr: &Expr| -> Option<u32> {
            let value = self.eval(expr, sheet, depth + 1, at).to_number().ok()?;
            (value >= 0.0).then_some(value as u32)
        };
        let rows = range.end.row - range.start.row + 1;
        let cols = range.end.col - range.start.col + 1;
        let (row, col) = match args.get(2) {
            Some(col) => (number(&args[1])?, number(col)?),
            None if rows == 1 => (1, number(&args[1])?),
            None if cols == 1 => (number(&args[1])?, 1),
            None => return None,
        };
        if row > rows || col > cols {
            return None;
        }
        let (first_row, last_row) = if row == 0 { (range.start.row, range.end.row) } else { (range.start.row + row - 1, range.start.row + row - 1) };
        let (first_col, last_col) = if col == 0 { (range.start.col, range.end.col) } else { (range.start.col + col - 1, range.start.col + col - 1) };
        Some((target, RangeRef::normalised(CellRef::new(first_col, first_row), CellRef::new(last_col, last_row))))
    }

    /// `OFFSET(reference, rows, cols, [height], [width])`: the base range moved
    /// by rows/cols and, if asked, resized. Off the sheet is #REF!.
    fn offset_range(
        &self,
        args: &[Expr],
        sheet: &str,
        depth: u32,
        at: At,
    ) -> Result<(String, RangeRef), ExcelError> {
        if !(3..=5).contains(&args.len()) {
            return Err(ExcelError::Value);
        }
        let (base_sheet, base) = self
            .reference_of(&args[0], sheet, depth, at)
            .ok_or(ExcelError::Value)?;
        let whole = |index: usize, default: i64| -> Result<i64, ExcelError> {
            match args.get(index) {
                None => Ok(default),
                Some(expr) => {
                    let value = self.eval_arg(expr, sheet, depth + 1, at).scalar();
                    if let Some(why) = value.err() {
                        return Err(why);
                    }
                    Ok(value.to_number()?.trunc() as i64)
                }
            }
        };
        let rows = whole(1, 0)?;
        let cols = whole(2, 0)?;
        let height = whole(3, base.height() as i64)?;
        let width = whole(4, base.width() as i64)?;
        if height <= 0 || width <= 0 {
            return Err(ExcelError::Ref);
        }
        let start_col = base.start.col as i64 + cols;
        let start_row = base.start.row as i64 + rows;
        let end_col = start_col + width - 1;
        let end_row = start_row + height - 1;
        if start_col < 0
            || start_row < 0
            || end_col > MAX_COL as i64
            || end_row > MAX_ROW as i64
        {
            return Err(ExcelError::Ref);
        }
        if !self.sheets.contains_key(&base_sheet) {
            return Err(ExcelError::Ref);
        }
        Ok((
            base_sheet,
            RangeRef::normalised(
                CellRef::new(start_col as u32, start_row as u32),
                CellRef::new(end_col as u32, end_row as u32),
            ),
        ))
    }

    fn offset_reference(&self, args: &[Expr], sheet: &str, depth: u32, at: At) -> Arg {
        match self.offset_range(args, sheet, depth, at) {
            Ok((s, range)) => Arg::Range(self.materialise(&s, &range, Skip::Nothing)),
            Err(why) => Arg::Value(Value::Error(why)),
        }
    }

    /// `INDIRECT(ref_text, [a1])`: read a reference out of text. Only A1-style
    /// text is understood; R1C1 (a1 = FALSE) is #REF! here.
    fn indirect_range(
        &self,
        args: &[Expr],
        sheet: &str,
        depth: u32,
        at: At,
    ) -> Result<(String, RangeRef), ExcelError> {
        if args.is_empty() || args.len() > 2 {
            return Err(ExcelError::Value);
        }
        let text = match self.eval_arg(&args[0], sheet, depth + 1, at).scalar() {
            Value::Text(s) => s,
            Value::Error(why) => return Err(why),
            _ => return Err(ExcelError::Ref),
        };
        let a1 = match args.get(1) {
            None => true,
            Some(expr) => self
                .eval_arg(expr, sheet, depth + 1, at)
                .scalar()
                .to_logical()
                .unwrap_or(true),
        };
        let (target, address) = match text.rfind('!') {
            Some(pos) => {
                let raw = text[..pos].trim();
                let name = if raw.len() >= 2 && raw.starts_with('\'') && raw.ends_with('\'') {
                    raw[1..raw.len() - 1].replace("''", "'")
                } else {
                    raw.to_string()
                };
                (name, text[pos + 1..].trim().to_string())
            }
            None => (sheet.to_string(), text.trim().to_string()),
        };
        if !self.sheets.contains_key(&target) {
            return Err(ExcelError::Ref);
        }
        // R1C1 text is read from where the formula stands: measured,
        // `INDIRECT("R2C1",FALSE)` is A2.
        let address = if a1 {
            address
        } else {
            let (col, row) = at.unwrap_or((0, 0));
            crate::formula_from_r1c1(&format!("={address}"), row, col)
                .map_err(|_| ExcelError::Ref)?
                .trim_start_matches('=')
                .to_string()
        };
        let range = parse_range_string(&address).ok_or(ExcelError::Ref)?;
        Ok((target, range))
    }

    fn indirect_reference(&self, args: &[Expr], sheet: &str, depth: u32, at: At) -> Arg {
        match self.indirect_range(args, sheet, depth, at) {
            Ok((s, range)) => Arg::Range(self.materialise(&s, &range, Skip::Nothing)),
            Err(why) => Arg::Value(Value::Error(why)),
        }
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
    fn materialise(&self, sheet: &str, range: &RangeRef, skip: Skip) -> RangeData {
        let range = &self.cut_to_fit(sheet, range);
        let held = self.sheets.get(sheet);
        RangeData::from_range(range, |col, row| {
            let passed_over = match skip {
                Skip::Nothing => false,
                Skip::Subtotals { by_hand } => {
                    self.is_subtotal_cell(sheet, col, row)
                        || held.is_some_and(|held| {
                            held.filtered_rows.contains(&row)
                                || (by_hand && held.hidden_rows.contains(&row))
                        })
                }
            };
            if passed_over {
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
/// The element of a worked-out block that a member of an array formula
/// shows, with a one-column or one-row block dealt along the other side.
fn element_of(worked: Arg, (dx, dy): (u32, u32)) -> Value {
    match worked {
        Arg::Value(value) => value,
        Arg::Range(block) => {
            let col = if block.width == 1 { 0 } else { dx as usize };
            let row = if block.height == 1 { 0 } else { dy as usize };
            if col >= block.width || row >= block.height {
                Value::Error(ExcelError::NA)
            } else {
                block.at(col, row)
            }
        }
    }
}

fn formula_result(value: Value) -> Value {
    match value {
        Value::Blank => Value::Number(0.0),
        // A cell holds no negative zero: `=-A1` over a blank, or a SUM of
        // nothing, is 0 in Excel, and a macro reading it writes 0.
        Value::Number(number) if number == 0.0 => Value::Number(0.0),
        other => other,
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
            0,
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

    /// A1 and A2 hold numbers. B1 reads A1, C1 reads B1, D1 reads A2, and E1
    /// sums both.
    fn a_little_chain() -> Workbook {
        let mut wb = book();
        wb.set_value("Sheet1", "A1", Value::Number(1.0)).unwrap();
        wb.set_value("Sheet1", "A2", Value::Number(2.0)).unwrap();
        wb.set_formula("Sheet1", "B1", "=A1*10").unwrap();
        wb.set_formula("Sheet1", "C1", "=B1+1").unwrap();
        wb.set_formula("Sheet1", "D1", "=A2*100").unwrap();
        wb.set_formula("Sheet1", "E1", "=SUM(A1:A2)").unwrap();
        wb.recalculate();
        wb
    }

    #[test]
    fn a_change_reaches_what_reads_it_and_what_reads_that() {
        // Typing into one cell of a 936,000-cell workbook made the editor work
        // out all 22,864 of its formulas — twenty seconds. Almost none of it
        // was needed.
        let mut wb = a_little_chain();
        wb.set_value("Sheet1", "A1", Value::Number(5.0)).unwrap();
        // A1 is column 0, row 0: everything here counts from zero.
        let report = wb.recalculate_after(&[("Sheet1".to_string(), (0, 0))]);
        assert_eq!(wb.value("Sheet1", "B1"), Value::Number(50.0), "reads A1");
        assert_eq!(wb.value("Sheet1", "C1"), Value::Number(51.0), "reads B1");
        assert_eq!(wb.value("Sheet1", "E1"), Value::Number(7.0), "its range holds A1");
        assert_eq!(wb.value("Sheet1", "D1"), Value::Number(200.0), "reads only A2");
        assert_eq!(report.evaluated, 3, "B1, C1 and E1, and not D1");
    }

    #[test]
    fn what_a_change_does_not_reach_is_left_exactly_as_it_was() {
        // The saving is only real if the untouched cells are genuinely not
        // worked out — so this makes one of them WRONG first, without saying
        // so, and then checks it stays wrong. A pass that quietly recomputed
        // everything would tidy it up and look like a pass.
        let mut wb = a_little_chain();
        wb.set_value("Sheet1", "A2", Value::Number(9.0)).unwrap();
        let report = wb.recalculate_after(&[("Sheet1".to_string(), (0, 0))]);
        assert_eq!(
            wb.value("Sheet1", "D1"),
            Value::Number(200.0),
            "still the old answer, because nobody said A2 had changed",
        );
        assert_eq!(report.evaluated, 3);
        // Told about it, or asked for the whole thing, it catches up.
        wb.recalculate();
        assert_eq!(wb.value("Sheet1", "D1"), Value::Number(900.0));
    }

    #[test]
    fn a_cell_that_has_just_been_given_a_formula_is_worked_out() {
        // The changed cell may be the formula itself rather than something it
        // reads, which is what typing one into an empty cell looks like.
        let mut wb = a_little_chain();
        wb.set_formula("Sheet1", "F1", "=A1+A2").unwrap();
        let report = wb.recalculate_after(&[("Sheet1".to_string(), (5, 0))]);
        assert_eq!(wb.value("Sheet1", "F1"), Value::Number(3.0));
        assert_eq!(report.evaluated, 1);
    }

    #[test]
    fn a_partial_pass_agrees_with_a_whole_one() {
        // Whatever the saving, the answers have to be the answers.
        let mut partly = a_little_chain();
        let mut wholly = a_little_chain();
        for (at, value) in [("A1", 7.0), ("A2", 11.0)] {
            partly.set_value("Sheet1", at, Value::Number(value)).unwrap();
            wholly.set_value("Sheet1", at, Value::Number(value)).unwrap();
        }
        partly.recalculate_after(&[
            ("Sheet1".to_string(), (0, 0)),
            ("Sheet1".to_string(), (0, 1)),
        ]);
        wholly.recalculate();
        for at in ["B1", "C1", "D1", "E1"] {
            assert_eq!(
                partly.value("Sheet1", at),
                wholly.value("Sheet1", at),
                "at {at}",
            );
        }
    }

    #[test]
    fn a_change_on_one_sheet_reaches_a_formula_on_another() {
        let mut wb = book();
        wb.add_sheet("Other");
        wb.set_value("Sheet1", "A1", Value::Number(2.0)).unwrap();
        wb.set_formula("Other", "A1", "=Sheet1!A1*3").unwrap();
        wb.recalculate();
        assert_eq!(wb.value("Other", "A1"), Value::Number(6.0));
        wb.set_value("Sheet1", "A1", Value::Number(5.0)).unwrap();
        let report = wb.recalculate_after(&[("Sheet1".to_string(), (0, 0))]);
        assert_eq!(wb.value("Other", "A1"), Value::Number(15.0));
        assert_eq!(report.evaluated, 1);
        // And a change on the sheet it does NOT read leaves it alone.
        wb.set_value("Sheet1", "A1", Value::Number(6.0)).unwrap();
        wb.recalculate_after(&[("Other".to_string(), (9, 9))]);
        assert_eq!(wb.value("Other", "A1"), Value::Number(15.0));
    }

    #[test]
    fn a_workbook_can_be_told_what_time_it_is() {
        // Pinning the moment is what makes a function meaning "now" testable
        // at all. 45297.75 is 2024-01-06 at six in the evening.
        let mut wb = book();
        wb.set_now(45297.75);
        wb.set_formula("Sheet1", "A1", "=TODAY()").unwrap();
        wb.set_formula("Sheet1", "A2", "=NOW()").unwrap();
        wb.set_formula("Sheet1", "A3", "=DAY(TODAY())").unwrap();
        wb.set_formula("Sheet1", "A4", "=TEXT(TODAY(),\"yyyy-mm-dd\")").unwrap();
        // Neither takes an argument.
        wb.set_formula("Sheet1", "A5", "=TODAY(1)").unwrap();
        wb.recalculate();
        assert_eq!(wb.value("Sheet1", "A1"), Value::Number(45297.0), "the day only");
        assert_eq!(wb.value("Sheet1", "A2"), Value::Number(45297.75), "and the hour");
        assert_eq!(wb.value("Sheet1", "A3"), Value::Number(6.0));
        assert_eq!(wb.value("Sheet1", "A4"), Value::text("2024-01-06"));
        assert_eq!(wb.value("Sheet1", "A5"), Value::Error(ExcelError::Value));
    }

    #[test]
    fn every_today_in_one_working_out_is_the_same_day() {
        // The clock is read once for a whole recalculation. A sheet where one
        // column thought it was Monday and the next thought it was Tuesday
        // would be worse than one that could not tell the time — and midnight
        // falls in the middle of some recalculation eventually.
        let mut wb = book();
        wb.set_formula("Sheet1", "A1", "=TODAY()").unwrap();
        wb.set_formula("Sheet1", "A2", "=TODAY()").unwrap();
        wb.set_formula("Sheet1", "A3", "=A1=A2").unwrap();
        // And it is a real date, not zero: well past the year 2000.
        wb.set_formula("Sheet1", "A4", "=TODAY()>36526").unwrap();
        wb.recalculate();
        assert_eq!(wb.value("Sheet1", "A3"), Value::Logical(true));
        assert_eq!(wb.value("Sheet1", "A4"), Value::Logical(true));
    }

    #[test]
    fn a_range_that_is_only_measured_is_not_waited_for() {
        // A date series that numbers itself by how far down the block it has
        // got: the same formula in every cell of B7:B10, each asking how tall
        // B7:B10 is. Excel answers without hesitating, because ROWS never
        // looks inside the range.
        //
        // Counting that range as a dependency makes the block a cycle, and
        // then NOTHING in it is evaluated — including, as the last assertion
        // shows, a cell standing outside it, since the cells it names are the
        // ones stuck.
        let mut wb = book();
        wb.set_value("Sheet1", "B3", Value::Number(44440.0)).unwrap();
        wb.set_value("Sheet1", "F3", Value::Number(44530.0)).unwrap();
        let counting = "=IF($B$3+ROWS($B$7:$B10)-1<=$F$3,$B$3+ROWS($B$7:$B10)-1,\"\")";
        for row in 7..=10 {
            wb.set_formula("Sheet1", &format!("B{row}"), counting).unwrap();
        }
        wb.set_formula("Sheet1", "Z1", counting).unwrap();
        wb.recalculate();
        for at in ["B7", "B8", "B10", "Z1"] {
            assert_eq!(wb.value("Sheet1", at), Value::Number(44443.0), "at {at}");
        }
    }

    #[test]
    fn how_tall_a_range_is_comes_from_the_reference() {
        // `materialise` cuts a range back to what the sheet holds, so counting
        // the block it returns would answer with the sheet's height rather
        // than the reference's. Excel says a whole column is 1,048,576 rows
        // however few of them are filled in.
        let mut wb = book();
        wb.set_value("Sheet1", "A1", Value::Number(1.0)).unwrap();
        wb.set_value("Sheet1", "A2", Value::Number(2.0)).unwrap();
        wb.set_formula("Sheet1", "C1", "=ROWS(A:A)").unwrap();
        wb.set_formula("Sheet1", "C2", "=COLUMNS(A1:D9)").unwrap();
        wb.set_formula("Sheet1", "C3", "=ROWS(A1:B9)").unwrap();
        wb.recalculate();
        assert_eq!(wb.value("Sheet1", "C1"), Value::Number(1_048_576.0));
        assert_eq!(wb.value("Sheet1", "C2"), Value::Number(4.0));
        assert_eq!(wb.value("Sheet1", "C3"), Value::Number(9.0));
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
        // DGET takes three ranges and means something by all of them. It now
        // works, and with no row matching the criteria it answers `#VALUE!`,
        // as Excel does -- and a great many real sheets swallow exactly that:
        // `IF(ISERR(DGET(...)),, DGET(...))` gives nothing either way.
        //
        // What this pins is that DGET is worked out ONCE, as one answer, not
        // applied to each of six hundred cells in turn (a block of `#NAME?`s
        // does not fit in a cell, and turned seventy-eight cells `#VALUE!`
        // when it was). Anything not named one-at-a-time keeps that.
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
        assert_eq!(wb.value("Sheet1", "E1"), Value::Error(ExcelError::Value));
        // And so the sheet's own way of swallowing it still works.
        assert_eq!(wb.value("Sheet1", "E2"), Value::Number(0.0));
    }

    /// An ARRAY formula -- `{=A1:A3*2}` dealt across D1:D3 -- is one formula
    /// whose block of answers is shared out, each member showing its own.
    /// Measured in Excel: `=A1:A2*2` across C1:C2 gives 2 and 4; across
    /// three cells the third is #N/A; `=7` across two cells is 7 and 7;
    /// `=ROW()` across R1:R2 is 1 and 2; and `=TRANSPOSE(A1:A2)` across a
    /// row F1:G1 is 1 and 2.
    #[test]
    fn an_array_formula_deals_its_block_across_its_members() {
        let mut wb = two_columns();
        for (at, offset) in [("D1", (0, 0)), ("D2", (0, 1)), ("D3", (0, 2)), ("D4", (0, 3))] {
            wb.set_array_member("Sheet1", at, "=A1:A3*2", offset).unwrap();
        }
        wb.set_array_member("Sheet1", "E1", "=7", (0, 0)).unwrap();
        wb.set_array_member("Sheet1", "E2", "=7", (0, 1)).unwrap();
        wb.set_array_member("Sheet1", "F1", "=TRANSPOSE(A1:A2)", (0, 0)).unwrap();
        wb.set_array_member("Sheet1", "G1", "=TRANSPOSE(A1:A2)", (1, 0)).unwrap();
        wb.set_array_member("Sheet1", "H1", "=ROW()", (0, 0)).unwrap();
        wb.set_array_member("Sheet1", "H2", "=ROW()", (0, 1)).unwrap();
        // A formula reading a member sees that member's share.
        wb.set_formula("Sheet1", "I1", "=D2+D3").unwrap();
        wb.recalculate();
        assert_eq!(wb.value("Sheet1", "D1"), Value::Number(2.0));
        assert_eq!(wb.value("Sheet1", "D2"), Value::Number(4.0));
        assert_eq!(wb.value("Sheet1", "D3"), Value::Number(6.0));
        assert_eq!(wb.value("Sheet1", "D4"), Value::Error(ExcelError::NA));
        assert_eq!(wb.value("Sheet1", "E1"), Value::Number(7.0));
        assert_eq!(wb.value("Sheet1", "E2"), Value::Number(7.0));
        assert_eq!(wb.value("Sheet1", "F1"), Value::Number(1.0));
        assert_eq!(wb.value("Sheet1", "G1"), Value::Number(2.0));
        assert_eq!(wb.value("Sheet1", "H1"), Value::Number(1.0));
        assert_eq!(wb.value("Sheet1", "H2"), Value::Number(2.0));
        assert_eq!(wb.value("Sheet1", "I1"), Value::Number(10.0));
        // Every member answers with the whole formula.
        assert_eq!(wb.formula("Sheet1", "D3"), Some("=A1:A3*2"));
    }

    #[test]
    fn a_reference_of_several_areas_goes_where_excel_takes_it() {
        // Every answer here is Excel's.
        let mut wb = Workbook::new();
        wb.add_sheet("Sheet1");
        for (row, (a, c)) in [(1, (1.0, 10.0)), (2, (2.0, 20.0)), (3, (3.0, 30.0))] {
            wb.set_value("Sheet1", &format!("A{row}"), Value::Number(a)).unwrap();
            wb.set_value("Sheet1", &format!("C{row}"), Value::Number(c)).unwrap();
        }
        for (cell, formula) in [
            ("E1", "=SUM((A1:A3,C1:C3))"),
            ("E2", "=AREAS((A1,B2,C3))"),
            ("E3", "=LARGE((A1:A3,C1:C3),2)"),
            ("E4", "=INDEX((A1:A3,C1:C3),2,1,2)"),
            ("E5", "=(A1,C1)"),
            ("E6", "=COUNTIF((A1:A3,C1:C3),\">2\")"),
        ] {
            wb.set_formula("Sheet1", cell, formula).unwrap();
        }
        wb.recalculate();
        assert_eq!(wb.value("Sheet1", "E1"), Value::Number(66.0));
        assert_eq!(wb.value("Sheet1", "E2"), Value::Number(3.0));
        assert_eq!(wb.value("Sheet1", "E3"), Value::Number(20.0));
        assert_eq!(wb.value("Sheet1", "E4"), Value::Number(20.0));
        assert_eq!(wb.value("Sheet1", "E5"), Value::Error(ExcelError::Value));
        assert_eq!(wb.value("Sheet1", "E6"), Value::Error(ExcelError::Value));
    }

    #[test]
    fn sheet_questions_and_trend_answer_as_excel_does() {
        // Every answer here is Excel's.
        let mut wb = Workbook::new();
        wb.add_sheet("Sheet1");
        wb.add_sheet("Two");
        wb.set_value("Sheet1", "A1", Value::Number(5.0)).unwrap();
        wb.set_formula("Sheet1", "A2", "=A1*2").unwrap();
        for (row, (x, y)) in [(1, (1.0, 2.0)), (2, (2.0, 4.0)), (3, (3.0, 7.0))] {
            wb.set_value("Sheet1", &format!("B{row}"), Value::Number(x)).unwrap();
            wb.set_value("Sheet1", &format!("C{row}"), Value::Number(y)).unwrap();
        }
        for (cell, formula) in [
            ("D1", "=SHEETS()"),
            ("D2", "=SHEET()"),
            ("D3", "=SHEET(Two!A1)"),
            ("D4", "=SHEET(\"Two\")"),
            ("D5", "=ISFORMULA(A2)"),
            ("D6", "=ISFORMULA(A1)"),
            ("D7", "=AREAS(A1:B2)"),
            ("D8", "=INDEX(TREND(C1:C3,B1:B3,{4;5},FALSE),2)"),
            ("D9", "=SUM(TREND(C1:C3,B1:B3))"),
            ("D10", "=SHEET(\"nope\")"),
        ] {
            wb.set_formula("Sheet1", cell, formula).unwrap();
        }
        wb.recalculate();
        let at = |cell: &str| wb.value("Sheet1", cell);
        assert_eq!(at("D1"), Value::Number(2.0));
        assert_eq!(at("D2"), Value::Number(1.0));
        assert_eq!(at("D3"), Value::Number(2.0));
        assert_eq!(at("D4"), Value::Number(2.0));
        assert_eq!(at("D5"), Value::Logical(true));
        assert_eq!(at("D6"), Value::Logical(false));
        assert_eq!(at("D7"), Value::Number(1.0));
        match at("D8") {
            Value::Number(n) => assert!((n - 11.0714285714286).abs() < 1e-12),
            other => panic!("{other:?}"),
        }
        match at("D9") {
            Value::Number(n) => assert!((n - 13.0).abs() < 1e-12),
            other => panic!("{other:?}"),
        }
        assert_eq!(at("D10"), Value::Error(ExcelError::NA));
    }

    #[test]
    fn let_binds_names_and_randbetween_rounds_both_bounds_up() {
        // Every answer here is Excel's.
        let mut wb = Workbook::new();
        wb.add_sheet("Sheet1");
        for (row, value) in [(1, 2.0), (2, 3.0), (3, 4.0)] {
            wb.set_value("Sheet1", &format!("A{row}"), Value::Number(value)).unwrap();
        }
        wb.set_formula("Sheet1", "B1", "=LET(x,2,y,x*3,x+y)").unwrap();
        wb.set_formula("Sheet1", "B2", "=LET(total,SUM(A1:A3),total/COUNT(A1:A3))").unwrap();
        wb.set_formula("Sheet1", "B3", "=LET(x,1,x,2,x)").unwrap();
        wb.set_formula("Sheet1", "B4", "=RANDBETWEEN(2.5,2.9)").unwrap();
        wb.set_formula("Sheet1", "B5", "=RANDBETWEEN(5,3)").unwrap();
        wb.recalculate();
        assert_eq!(wb.value("Sheet1", "B1"), Value::Number(8.0));
        assert_eq!(wb.value("Sheet1", "B2"), Value::Number(3.0));
        assert_eq!(wb.value("Sheet1", "B3"), Value::Error(ExcelError::Value));
        assert_eq!(wb.value("Sheet1", "B4"), Value::Number(3.0));
        assert_eq!(wb.value("Sheet1", "B5"), Value::Error(ExcelError::Num));
    }

    #[test]
    fn subtotal_passes_over_filtered_rows_and_from_101_hidden_ones() {
        let mut wb = Workbook::new();
        wb.add_sheet("Sheet1");
        for (row, value) in [(1, 1.0), (2, 10.0), (3, 100.0), (4, 1000.0)] {
            wb.set_value("Sheet1", &format!("A{row}"), Value::Number(value)).unwrap();
        }
        // Row 2 hidden by a filter, row 3 hidden by hand (0-based rows).
        wb.hide_row("Sheet1", 1, true);
        wb.hide_row("Sheet1", 2, false);
        wb.set_formula("Sheet1", "B1", "=SUBTOTAL(9,A1:A4)").unwrap();
        wb.set_formula("Sheet1", "B2", "=SUBTOTAL(109,A1:A4)").unwrap();
        wb.set_formula("Sheet1", "B3", "=SUM(A1:A4)").unwrap();
        wb.recalculate();
        assert_eq!(wb.value("Sheet1", "B1"), Value::Number(1101.0));
        assert_eq!(wb.value("Sheet1", "B2"), Value::Number(1001.0));
        assert_eq!(wb.value("Sheet1", "B3"), Value::Number(1111.0));
    }

    /// A space between references is the cells they share, #NULL! where they
    /// share none; MMULT is the matrix product. Measured in Excel.
    #[test]
    fn intersection_and_mmult() {
        let mut wb = Workbook::new();
        wb.add_sheet("Sheet1");
        for (at, value) in ["A1", "B1", "C1", "A2", "B2", "C2", "A3", "B3", "C3"].iter().zip(1..) {
            wb.set_value("Sheet1", at, Value::Number(value as f64)).unwrap();
        }
        wb.set_formula("Sheet1", "E1", "=SUM(A1:C3 B2:C3)").unwrap();
        wb.set_formula("Sheet1", "E2", "=SUM(A1:A3 C1:C3)").unwrap();
        wb.set_formula("Sheet1", "E3", "=ROWS(A1:C3 B2:C3)").unwrap();
        wb.set_formula("Sheet1", "E4", "=ISREF(A1 B1)").unwrap();
        wb.set_formula("Sheet1", "E5", "=SUM(MMULT(A1:C1,A1:A3))").unwrap();
        wb.set_formula("Sheet1", "E6", "=A1:A3,C1").unwrap();
        wb.recalculate();
        assert_eq!(wb.value("Sheet1", "E1"), Value::Number(28.0));
        assert_eq!(wb.value("Sheet1", "E2"), Value::Error(ExcelError::Null));
        assert_eq!(wb.value("Sheet1", "E3"), Value::Number(2.0));
        assert_eq!(wb.value("Sheet1", "E4"), Value::Logical(false));
        assert_eq!(wb.value("Sheet1", "E5"), Value::Number(30.0));
        assert_eq!(wb.value("Sheet1", "E6"), Value::Error(ExcelError::Value));
    }

    /// A formula reading its own cell is a cycle and stays 0, as Excel
    /// leaves it: measured, `=SUM(5:6)` in D5 is 0. RATE finds a rate Newton
    /// from the guess misses: `RATE(360,-1073.64,200000)` is 0.0041666445...
    #[test]
    fn a_formula_reading_itself_is_zero_and_rate_is_found() {
        let mut wb = Workbook::new();
        wb.add_sheet("Sheet1");
        wb.set_value("Sheet1", "A5", Value::Number(5.0)).unwrap();
        wb.set_value("Sheet1", "A6", Value::Number(7.0)).unwrap();
        wb.set_formula("Sheet1", "D5", "=SUM(5:6)").unwrap();
        wb.set_formula("Sheet1", "E1", "=RATE(360,-1073.64,200000)").unwrap();
        let report = wb.recalculate();
        assert_eq!(wb.value("Sheet1", "D5"), Value::Number(0.0));
        assert_eq!(report.circular, vec!["Sheet1!D5".to_string()]);
        match wb.value("Sheet1", "E1") {
            Value::Number(rate) => assert!((rate - 0.00416664453634559).abs() < 1e-12, "{rate}"),
            other => panic!("{other:?}"),
        }
    }

    /// An error as the criterion counts the cells holding that error.
    /// Measured: over 1, 2, 3, #N/A, #DIV/0!, `COUNTIF(A1:A5,C1)` with C1
    /// #N/A is 1, and `SUMIF` of it over 10s is 10.
    #[test]
    fn an_error_criterion_finds_that_error() {
        let mut wb = Workbook::new();
        wb.add_sheet("Sheet1");
        for row in 1..=3 {
            wb.set_value("Sheet1", &format!("A{row}"), Value::Number(row as f64)).unwrap();
        }
        for row in 1..=5 {
            wb.set_value("Sheet1", &format!("B{row}"), Value::Number(10.0)).unwrap();
        }
        wb.set_formula("Sheet1", "A4", "=NA()").unwrap();
        wb.set_formula("Sheet1", "A5", "=1/0").unwrap();
        wb.set_formula("Sheet1", "C1", "=NA()").unwrap();
        wb.set_formula("Sheet1", "C2", "=1/0").unwrap();
        wb.set_formula("Sheet1", "D1", "=COUNTIF(A1:A5,C1)").unwrap();
        wb.set_formula("Sheet1", "D2", "=COUNTIF(A1:A5,\"#N/A\")").unwrap();
        wb.set_formula("Sheet1", "D3", "=SUMIF(A1:A5,C1,B1:B5)").unwrap();
        wb.set_formula("Sheet1", "D4", "=SUMIFS(B1:B5,A1:A5,C2)").unwrap();
        wb.set_formula("Sheet1", "D5", "=COUNTIFS(A1:A5,#DIV/0!)").unwrap();
        wb.set_formula("Sheet1", "D6", "=AVERAGEIF(A1:A5,C1,B1:B5)").unwrap();
        wb.set_formula("Sheet1", "D7", "=AVERAGEIF(A1:A3,C1,B1:B3)").unwrap();
        wb.recalculate();
        assert_eq!(wb.value("Sheet1", "D1"), Value::Number(1.0));
        assert_eq!(wb.value("Sheet1", "D2"), Value::Number(1.0));
        assert_eq!(wb.value("Sheet1", "D3"), Value::Number(10.0));
        assert_eq!(wb.value("Sheet1", "D4"), Value::Number(10.0));
        assert_eq!(wb.value("Sheet1", "D5"), Value::Number(1.0));
        assert_eq!(wb.value("Sheet1", "D6"), Value::Number(10.0));
        assert_eq!(wb.value("Sheet1", "D7"), Value::Error(ExcelError::DivZero));
    }

    /// `@` takes the cell of a range in line with the asking cell, and
    /// #VALUE! where none is. Measured: `=@A1:A3` in D2 is A2's 2, in M5 is
    /// #VALUE!; `=@A1:A3` beside a row range takes its column.
    #[test]
    fn an_at_takes_the_cell_in_line() {
        let mut wb = two_columns();
        wb.set_formula("Sheet1", "D2", "=@A1:A3").unwrap();
        wb.set_formula("Sheet1", "M5", "=@A1:A3").unwrap();
        wb.set_formula("Sheet1", "B7", "=@A1:C1").unwrap();
        wb.set_formula("Sheet1", "E3", "=@A1:A3+1").unwrap();
        wb.recalculate();
        assert_eq!(wb.value("Sheet1", "D2"), Value::Number(2.0));
        assert_eq!(wb.value("Sheet1", "M5"), Value::Error(ExcelError::Value));
        assert_eq!(wb.value("Sheet1", "B7"), Value::Number(10.0));
        assert_eq!(wb.value("Sheet1", "E3"), Value::Number(4.0));
    }

    #[test]
    fn a_block_of_answers_spills_and_the_cell_shows_the_first() {
        // This test used to assert the opposite — that a block written into
        // one cell is `#VALUE!`, "as it is in Excel without dynamic arrays".
        // Excel was asked, and with A1:A3 holding 1, 2, 3 it fills D1:D3 with
        // 2, 3, 4: the block spills and D1 shows the first of it. A bare range
        // reference does the same, `=A1:A3` in D1 giving 1.
        //
        // It matters here because a file stores such a formula in the anchor
        // cell alone, so the first element is the value the workbook was saved
        // holding.
        let mut wb = two_columns();
        wb.set_formula("Sheet1", "D1", "=A1:A3+1").unwrap();
        wb.set_formula("Sheet1", "D2", "=A1:A3").unwrap();
        wb.set_formula("Sheet1", "D3", "=IF(A1:A3>1,10,0)").unwrap();
        wb.recalculate();
        assert_eq!(wb.value("Sheet1", "D1"), Value::Number(2.0));
        assert_eq!(wb.value("Sheet1", "D2"), Value::Number(1.0));
        assert_eq!(wb.value("Sheet1", "D3"), Value::Number(0.0), "1 is not > 1");
    }

    #[test]
    fn offset_and_indirect_resolve_references() {
        // A1:A3 = 1,2,3 and B1:B3 = 10,20,30.
        let mut wb = two_columns();
        wb.set_formula("Sheet1", "D1", "=OFFSET(A1,2,0)").unwrap();
        wb.set_formula("Sheet1", "D2", "=OFFSET(A1,0,1)").unwrap();
        wb.set_formula("Sheet1", "D3", "=SUM(OFFSET(A1,0,0,3,1))").unwrap();
        wb.set_formula("Sheet1", "D4", "=IFERROR(OFFSET(A1,-1,0),\"REF\")").unwrap();
        wb.set_formula("Sheet1", "D5", "=INDIRECT(\"B2\")").unwrap();
        wb.set_formula("Sheet1", "D6", "=SUM(INDIRECT(\"A1:A3\"))").unwrap();
        // A reference built by a nested INDIRECT is followed, not materialised.
        wb.set_formula("Sheet1", "D7", "=OFFSET(INDIRECT(\"A1\"),2,1)").unwrap();
        wb.set_formula("Sheet1", "D8", "=IFERROR(INDIRECT(\"nope\"),\"REF\")").unwrap();
        wb.recalculate();
        assert_eq!(wb.value("Sheet1", "D1"), Value::Number(3.0), "A3");
        assert_eq!(wb.value("Sheet1", "D2"), Value::Number(10.0), "B1");
        assert_eq!(wb.value("Sheet1", "D3"), Value::Number(6.0), "A1:A3");
        assert_eq!(wb.value("Sheet1", "D4"), Value::text("REF"), "off the sheet");
        assert_eq!(wb.value("Sheet1", "D5"), Value::Number(20.0), "B2");
        assert_eq!(wb.value("Sheet1", "D6"), Value::Number(6.0), "A1:A3 again");
        assert_eq!(wb.value("Sheet1", "D7"), Value::Number(30.0), "B3 via nesting");
        assert_eq!(wb.value("Sheet1", "D8"), Value::text("REF"), "unparseable");
    }

    #[test]
    fn sortby_orders_one_block_by_the_numbers_in_another() {
        // The block being ordered and the numbers ordering it are separate, so
        // the numbers never appear in the answer — which is the whole point of
        // it over SORT.
        let mut wb = book();
        for (at, (name, score)) in [("Belgium", 3.0), ("Afghanistan", 9.0),
                                    ("Chad", 5.0), ("Denmark", 7.0)].iter().enumerate() {
            wb.set_value("Sheet1", &format!("E{}", at + 4), Value::text(*name)).unwrap();
            wb.set_value("Sheet1", &format!("F{}", at + 4), Value::Number(*score)).unwrap();
        }
        wb.set_formula("Sheet1", "A1", "=INDEX(_xlfn.SORTBY(E4:F7,F4:F7,-1),1,1)").unwrap();
        wb.set_formula("Sheet1", "A2", "=INDEX(_xlfn.SORTBY(E4:F7,F4:F7,1),1,1)").unwrap();
        wb.set_formula("Sheet1", "A3", "=INDEX(_xlfn.SORTBY(E4:E7,F4:F7,-1),2,1)").unwrap();
        // One column in, one column out: the ordering column is not carried
        // through into the answer.
        wb.set_formula("Sheet1", "A4", "=COLUMNS(_xlfn.SORTBY(E4:E7,F4:F7,-1))").unwrap();
        wb.recalculate();
        assert_eq!(wb.value("Sheet1", "A1"), Value::text("Afghanistan"), "9 is the most");
        assert_eq!(wb.value("Sheet1", "A2"), Value::text("Belgium"), "3 is the least");
        assert_eq!(wb.value("Sheet1", "A3"), Value::text("Denmark"), "7 is next");
        assert_eq!(wb.value("Sheet1", "A4"), Value::Number(1.0));
    }

    #[test]
    fn the_three_that_hand_back_a_block() {
        // UNIQUE keeps the first of each distinct row, SORT puts the rows in
        // order, FILTER keeps the ones a second block says to keep. Each shows
        // the first of its block in the cell it is written in, and each is
        // still a whole block to whatever is wrapped around it.
        let mut wb = book();
        for (at, (name, pay)) in [("pear", 1.0), ("apple", 2.0), ("pear", 3.0),
                                  ("plum", 4.0), ("apple", 5.0)].iter().enumerate() {
            wb.set_value("Sheet1", &format!("A{}", at + 1), Value::text(*name)).unwrap();
            wb.set_value("Sheet1", &format!("B{}", at + 1), Value::Number(*pay)).unwrap();
        }
        for (at, formula) in [
            "=UNIQUE(A1:A5)",
            "=COUNTA(UNIQUE(A1:A5))",
            "=COUNTA(UNIQUE(A1:A5,FALSE,TRUE))",
            "=SORT(A1:A5)",
            "=SORT(A1:A5,1,-1)",
            "=FILTER(A1:A5,B1:B5>3)",
            "=COUNTA(FILTER(A1:A5,B1:B5>3))",
            "=FILTER(A1:A5,B1:B5>99,\"none\")",
            "=FILTER(A1:A5,B1:B5>99)",
            "=_xlfn._xlws.SORT(_xlfn.UNIQUE(A1:A5))",
        ].iter().enumerate() {
            wb.set_formula("Sheet1", &format!("D{}", at + 1), formula).unwrap();
        }
        wb.recalculate();
        let shown = |at: &str| wb.value("Sheet1", at);
        assert_eq!(shown("D1"), Value::text("pear"), "the first as it stands");
        assert_eq!(shown("D2"), Value::Number(3.0), "pear, apple, plum");
        assert_eq!(shown("D3"), Value::Number(1.0), "only plum appears once");
        assert_eq!(shown("D4"), Value::text("apple"));
        assert_eq!(shown("D5"), Value::text("plum"), "the last one, first");
        assert_eq!(shown("D6"), Value::text("plum"), "rows 4 and 5 survive");
        assert_eq!(shown("D7"), Value::Number(2.0));
        assert_eq!(shown("D8"), Value::text("none"), "nothing kept, so the spare");
        assert_eq!(shown("D9"), Value::Error(ExcelError::NA), "and nothing to say");
        assert_eq!(shown("D10"), Value::text("apple"), "the prefixes come off");
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

    /// An array written out where a range could go. Measured in Excel by
    /// putting each formula in a cell and reading what it showed.
    #[test]
    fn reads_an_array_written_out() {
        let mut book = Workbook::new();
        book.add_sheet("Sheet1");

        let measured = [
            ("SUM({1,2,3})", Value::Number(6.0)),
            ("SUM({1,2;3,4})", Value::Number(10.0)),
            ("INDEX({1,2;3,4},2,1)", Value::Number(3.0)),
            ("ROWS({1,2;3,4})", Value::Number(2.0)),
            ("COLUMNS({1,2;3,4})", Value::Number(2.0)),
            // SUM steps over text and logicals in an array, as it does in a
            // range, and COUNT counts only the numbers.
            ("SUM({1,\"a\",TRUE})", Value::Number(1.0)),
            ("COUNTA({1,\"a\"})", Value::Number(2.0)),
            ("COUNT({1,\"a\"})", Value::Number(1.0)),
            ("MATCH(2,{1,2,3},0)", Value::Number(2.0)),
            ("SUM({-1,2})", Value::Number(1.0)),
            ("SUM({1.5,2.5})", Value::Number(4.0)),
            // An error inside is passed on.
            ("SUM({#N/A,1})", Value::Error(ExcelError::NA)),
        ];
        for (formula, want) in measured {
            assert_eq!(
                book.evaluate("Sheet1", formula).unwrap(),
                want,
                "for {formula}"
            );
        }
    }

    /// Excel refuses a ragged one outright rather than padding it, so this
    /// refuses to parse rather than guessing a shape.
    #[test]
    fn refuses_an_array_whose_rows_are_not_the_same_width() {
        assert!(crate::parse("SUM({1,2;3})").is_err());
        assert!(crate::parse("{1,2;3}").is_err());
    }

    /// Only constants go inside one.
    #[test]
    fn refuses_a_reference_inside_an_array() {
        assert!(crate::parse("SUM({A1,2})").is_err());
        assert!(crate::parse("SUM({SUM(1),2})").is_err());
    }

    /// Which ARGUMENTS a function takes one at a time, measured in Excel with
    /// A1:A3 holding x, y, x and B1:B3 holding 10, 20, 30.
    ///
    /// A lookup reads its haystack whole and its needle one at a time; a
    /// conditional aggregate is the same shape reversed. Getting the grain
    /// wrong either way answers one thing where Excel answers several.
    #[test]
    fn a_function_takes_some_of_its_arguments_one_at_a_time() {
        let mut book = Workbook::new();
        book.add_sheet("Sheet1");
        let _ = book.set_value("Sheet1", "A1", Value::text("x"));
        let _ = book.set_value("Sheet1", "A2", Value::text("y"));
        let _ = book.set_value("Sheet1", "A3", Value::text("x"));
        let _ = book.set_value("Sheet1", "B1", Value::Number(10.0));
        let _ = book.set_value("Sheet1", "B2", Value::Number(20.0));
        let _ = book.set_value("Sheet1", "B3", Value::Number(30.0));

        let measured = [
            // The criterion is taken one at a time, so a cell shows the first
            // answer and an aggregate over it sees both.
            ("COUNTIF(A1:A3,{\"x\",\"y\"})", 2.0),
            ("SUM(COUNTIF(A1:A3,{\"x\",\"y\"}))", 3.0),
            ("SUMPRODUCT(COUNTIF(A1:A3,{\"x\",\"y\"}))", 3.0),
            ("SUM(COUNTIFS(A1:A3,{\"x\",\"y\"}))", 3.0),
            ("SUM(SUMIF(A1:A3,{\"x\",\"y\"},B1:B3))", 60.0),
            ("SUM(SUMIFS(B1:B3,A1:A3,{\"x\",\"y\"}))", 60.0),
            // MATCH takes its needle one at a time — one answer per needle.
            ("MATCH({\"y\",\"x\"},A1:A3,0)", 2.0),
            ("SUM(MATCH({\"y\",\"x\"},A1:A3,0))", 3.0),
            ("SUMPRODUCT(--ISNUMBER(MATCH({\"x\",\"z\"},A1:A3,0)))", 1.0),
            // Every argument, for a function that only knows single values.
            ("SUM(LEN({\"ab\",\"cde\"}))", 5.0),
            // ★ And one that does NOT: a lookup answers for the first needle
            // alone. `SUM(VLOOKUP({"x","y"},A1:B3,2,FALSE))` is 10 in Excel,
            // not 30 — it looks like MATCH and behaves the other way.
            ("SUM(VLOOKUP({\"x\",\"y\"},A1:B3,2,FALSE))", 10.0),
        ];
        for (formula, want) in measured {
            assert_eq!(
                book.evaluate("Sheet1", formula).unwrap(),
                Value::Number(want),
                "for {formula}"
            );
        }
    }
}
