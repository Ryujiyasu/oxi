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

#[derive(Debug, Clone)]
enum LocalBinding {
    Value(Arg),
    Reference { sheet: String, range: RangeRef },
}

impl LocalBinding {
    fn reference(&self) -> Option<(String, RangeRef)> {
        match self {
            Self::Reference { sheet, range } => Some((sheet.clone(), *range)),
            Self::Value(Arg::Range(block)) => block.cell_reference(),
            _ => None,
        }
    }
}

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
        let wanted = crate::case_fold::table_key(heading.trim());
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

/// Whether a formula body has a comma outside every bracket and quote: the
/// list of blocks a name can stand for.
fn has_top_level_comma(body: &str) -> bool {
    let mut depth = 0usize;
    let mut quote: Option<char> = None;
    for ch in body.chars() {
        match (quote, ch) {
            (Some(open), c) if c == open => quote = None,
            (Some(_), _) => {}
            (None, '"' | '\'') => quote = Some(ch),
            (None, '(' | '{' | '[') => depth += 1,
            (None, ')' | '}' | ']') => depth = depth.saturating_sub(1),
            (None, ',') if depth == 0 => return true,
            _ => {}
        }
    }
    false
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

/// An argument handed to a function the workbook's macros define: the cells
/// themselves where a reference was written, so the function can walk them.
#[derive(Debug, Clone)]
pub enum UserArg {
    Value(Value),
    Block(RangeData),
    /// What the cells hold is carried too, so an answer worked out for
    /// them is not reused once they change.
    Cells { sheet: String, range: RangeRef, values: Vec<Value> },
}

/// Where a formula calling one of them stands: sheet, column and row, the
/// last two counted from zero.
pub type UserCaller = Option<(String, u32, u32)>;

pub type FormulaCell = (String, (u32, u32));

#[derive(Debug, Clone, Copy)]
pub struct UserFunctionCallContext { pub occurrence: u64, pub provisional: bool, pub provisional_indirect_text_value: bool, pub dynamic_reference_argument_mask: u64 }

/// Published array extents at the start of a host calculation operation.
pub type SpillProbeLayout = Vec<(FormulaCell, (u32, u32))>;

type UserFunctionHook = Box<dyn FnMut(&str, Vec<UserArg>, UserCaller) -> Option<Arg>>;
type UserFunctionReferenceRetryHook = Box<dyn FnMut(&str, &[UserArg], &UserCaller, UserFunctionCallContext, bool, bool) -> bool>;

/// Cell values at the instant a callback was requested. Later calculation
/// steps can publish other values before the host dispatches the callback.
#[derive(Debug, Clone)]
pub struct UserFunctionReadView {
    epoch: u64,
    history: std::rc::Rc<std::cell::RefCell<UserReadHistory>>,
}

#[derive(Debug, Default)]
struct UserReadHistory {
    epoch: u64,
    cells: BTreeMap<FormulaCell, (Value, Vec<(u64, Value)>)>,
}

impl UserReadHistory {
    fn change(&mut self, sheet: &str, coord: (u32, u32), before: &Value, after: &Value) {
        let unchanged = match (before, after) {
            (Value::Number(before), Value::Number(after)) => before.to_bits() == after.to_bits(),
            _ => before == after,
        };
        if unchanged { return; }
        self.epoch += 1;
        let entry = self.cells.entry((sheet.to_string(), coord)).or_insert_with(|| (before.clone(), Vec::new()));
        entry.1.push((self.epoch, after.clone()));
    }
}

impl UserFunctionReadView {
    pub fn value(&self, sheet: &str, col: u32, row: u32) -> Option<Value> {
        let history = self.history.borrow();
        let (original, changes) = history.cells.get(&(sheet.to_string(), (col, row)))?;
        let at = changes.partition_point(|(epoch, _)| *epoch <= self.epoch);
        Some(if at == 0 { original.clone() } else { changes[at - 1].1.clone() })
    }
}

thread_local! {
    static USER_REFERENCE_ARGUMENTS: std::cell::RefCell<BTreeMap<String, BTreeSet<usize>>> = const { std::cell::RefCell::new(BTreeMap::new()) };
    static USER_VARIANT_ARGUMENTS: std::cell::RefCell<BTreeMap<String, BTreeSet<usize>>> = const { std::cell::RefCell::new(BTreeMap::new()) };
    static USER_VALUE_READ_ARGUMENTS: std::cell::RefCell<BTreeMap<String, BTreeSet<usize>>> = const { std::cell::RefCell::new(BTreeMap::new()) };
    static USER_ADDRESS_ONLY_ARGUMENTS: std::cell::RefCell<BTreeMap<String, BTreeSet<usize>>> = const { std::cell::RefCell::new(BTreeMap::new()) };
    static USER_FUNCTION_ARGUMENT_CONTEXT: std::cell::RefCell<Option<(String, usize)>> = const { std::cell::RefCell::new(None) };
    static USER_FORMULA_READS: std::cell::RefCell<BTreeSet<FormulaCell>> = const { std::cell::RefCell::new(BTreeSet::new()) };
    static USER_FUNCTION_PENDING: std::cell::Cell<bool> = const { std::cell::Cell::new(false) };
    static USER_FUNCTION_WAIT_FOR_DEPENDENCIES: std::cell::Cell<bool> = const { std::cell::Cell::new(false) };
    static USER_FUNCTION_SPILL_PROBE_LAYOUT: std::cell::RefCell<Option<SpillProbeLayout>> = const { std::cell::RefCell::new(None) };
    static USER_FUNCTION_WAIT_FOR_DYNAMIC_LOOKUPS: std::cell::Cell<bool> = const { std::cell::Cell::new(false) };
    static USER_FUNCTION_PROVISIONAL_DYNAMIC_TEXT_VALUE: std::cell::Cell<bool> = const { std::cell::Cell::new(false) };
    static USER_FUNCTION_PROVISIONAL_ZERO_TEXT_PROBE: std::cell::Cell<bool> = const { std::cell::Cell::new(false) };
    static USER_FUNCTION_PRIORITIZE_DYNAMIC_INDIRECT_CONSUMERS: std::cell::Cell<bool> = const { std::cell::Cell::new(false) };
    static USER_FUNCTION_PROVISIONAL_TEXT_PROBE_TARGET: std::cell::RefCell<Option<FormulaCell>> = const { std::cell::RefCell::new(None) };
    static USER_FUNCTION_FULL_REBUILD: std::cell::Cell<bool> = const { std::cell::Cell::new(false) };
    static USER_FUNCTION_FULL_CALCULATION: std::cell::Cell<bool> = const { std::cell::Cell::new(false) };
    static USER_FUNCTION_DISTINCT_PROVISIONAL_CALLS: std::cell::Cell<bool> = const { std::cell::Cell::new(false) };
    static USER_FUNCTION_AT_ROOT: std::cell::Cell<bool> = const { std::cell::Cell::new(false) };
    static USER_FUNCTION_OCCURRENCE: std::cell::Cell<u64> = const { std::cell::Cell::new(0) };
    static USER_FUNCTION_CONTEXT: std::cell::Cell<UserFunctionCallContext> = const { std::cell::Cell::new(UserFunctionCallContext { occurrence: 0, provisional: false, provisional_indirect_text_value: false, dynamic_reference_argument_mask: 0 }) };
    static USER_PROVISIONAL_INDIRECT_TEXT_VALUE: std::cell::Cell<bool> = const { std::cell::Cell::new(false) };
    static USER_FUNCTIONS: std::cell::RefCell<Option<UserFunctionHook>> = const { std::cell::RefCell::new(None) };
    static USER_FUNCTION_REFERENCE_RETRY: std::cell::RefCell<Option<UserFunctionReferenceRetryHook>> = const { std::cell::RefCell::new(None) };
    static USER_FUNCTION_READ_VIEW: std::cell::RefCell<Option<UserFunctionReadView>> = const { std::cell::RefCell::new(None) };
    static USER_FUNCTION_WRITTEN_DEPENDENCIES: std::cell::RefCell<Vec<(String, RangeRef)>> = const { std::cell::RefCell::new(Vec::new()) };
}

/// Whether a direct scalar answer may satisfy this formula replay.
/// Geometry probes remain distinct calls even when their values agree.
pub fn user_function_scalar_replay_allowed() -> bool { USER_FUNCTION_AT_ROOT.with(std::cell::Cell::get) }

pub fn user_function_call_context() -> UserFunctionCallContext { USER_FUNCTION_CONTEXT.with(std::cell::Cell::get) }

pub fn user_function_read_view() -> Option<UserFunctionReadView> {
    USER_FUNCTION_READ_VIEW.with(|held| held.borrow().clone())
}

/// Formula references written into the call, before dynamic reference functions
/// select additional ranges. Names and tables retain their written provenance.
pub fn user_function_written_dependencies() -> Vec<(String, RangeRef)> {
    USER_FUNCTION_WRITTEN_DEPENDENCIES.with(|held| held.borrow().clone())
}

/// Let formulas call the workbook's own functions while `hook` is set: it
/// answers `None` for a name that is not one of them.
pub fn set_user_function_hook(hook: Option<UserFunctionHook>) {
    USER_FUNCTION_REFERENCE_RETRY.with(|held| *held.borrow_mut() = None);
    USER_FUNCTION_PENDING.with(|pending| pending.set(false));
    USER_FUNCTION_OCCURRENCE.with(|occurrence| occurrence.set(0));
    if hook.is_none() {
        USER_FUNCTION_SPILL_PROBE_LAYOUT.with(|held| held.replace(None));
        USER_FORMULA_READS.with(|held| held.borrow_mut().clear());
        USER_REFERENCE_ARGUMENTS.with(|held| held.borrow_mut().clear());
        USER_VARIANT_ARGUMENTS.with(|held| held.borrow_mut().clear());
        USER_VALUE_READ_ARGUMENTS.with(|held| held.borrow_mut().clear());
        USER_ADDRESS_ONLY_ARGUMENTS.with(|held| held.borrow_mut().clear());
        USER_FUNCTION_ARGUMENT_CONTEXT.with(|held| held.replace(None));
    }
    USER_FUNCTIONS.with(|held| *held.borrow_mut() = hook);
}

/// Observe a nested callback reference used as a provisional value. The final argument
/// says whether its selected cells are ready. Return true to discard a cached
/// provisional reference and request the callback again after publication.
/// Replacing the main function hook clears this observer.
pub fn set_user_function_reference_retry_hook(hook: Option<UserFunctionReferenceRetryHook>) {
    USER_FUNCTION_REFERENCE_RETRY.with(|held| *held.borrow_mut() = hook);
}

/// Scope provisional spill reads to ownership known before callback replay.
/// Return the previous layout so nested host operations can restore it.
pub fn replace_user_function_spill_probe_layout(layout: Option<SpillProbeLayout>) -> Option<SpillProbeLayout> {
    USER_FUNCTION_SPILL_PROBE_LAYOUT.with(|held| held.replace(layout))
}

/// Install the current project's object parameter declarations for sheet calls.
pub fn set_user_function_reference_arguments(functions: &[(String, Vec<usize>)]) {
    USER_REFERENCE_ARGUMENTS.with(|held| {
        *held.borrow_mut() = functions.iter().map(|(name, indexes)|
            (name.to_ascii_uppercase(), indexes.iter().copied().collect())).collect();
    });
}

pub fn set_user_function_variant_arguments(functions: &[(String, Vec<usize>)]) {
    USER_VARIANT_ARGUMENTS.with(|held| {
        *held.borrow_mut() = functions.iter().map(|(name, indexes)|
            (name.to_ascii_uppercase(), indexes.iter().copied().collect())).collect();
    });
}

pub fn set_user_function_value_read_arguments(functions: &[(String, Vec<usize>)]) {
    USER_VALUE_READ_ARGUMENTS.with(|held| {
        *held.borrow_mut() = functions.iter().map(|(name, indexes)|
            (name.to_ascii_uppercase(), indexes.iter().copied().collect())).collect();
    });
}

pub fn set_user_function_address_only_arguments(functions: &[(String, Vec<usize>)]) {
    USER_ADDRESS_ONLY_ARGUMENTS.with(|held| {
        *held.borrow_mut() = functions.iter().map(|(name, indexes)|
            (name.to_ascii_uppercase(), indexes.iter().copied().collect())).collect();
    });
}

fn user_function_current_argument_is_address_only() -> bool {
    USER_FUNCTION_ARGUMENT_CONTEXT.with(|context| {
        let Some((name, index)) = context.borrow().clone() else { return false; };
        USER_ADDRESS_ONLY_ARGUMENTS.with(|arguments| arguments.borrow().get(&name)
            .is_some_and(|indexes| indexes.contains(&index)))
    })
}

fn user_reference_argument(name: &str, index: usize) -> bool {
    USER_REFERENCE_ARGUMENTS.with(|held| held.borrow().get(name)
        .is_some_and(|indexes| indexes.contains(&index)))
}

fn user_value_read_argument(name: &str, index: usize) -> bool {
    USER_VALUE_READ_ARGUMENTS.with(|held| held.borrow().get(name)
        .is_some_and(|indexes| indexes.contains(&index)))
}

fn user_function_address_only_argument(name: &str, index: usize) -> bool {
    USER_ADDRESS_ONLY_ARGUMENTS.with(|held| held.borrow().get(name)
        .is_some_and(|indexes| indexes.contains(&index)))
}

fn user_variant_argument(name: &str, index: usize) -> bool {
    USER_VARIANT_ARGUMENTS.with(|held| held.borrow().get(name)
        .is_some_and(|indexes| indexes.contains(&index)))
}

/// The ranges actually read by earlier callbacks, used only to retry a provisional view.
/// These reads do not add workbook dirty dependencies.
pub fn record_user_function_reads(reads: impl IntoIterator<Item=FormulaCell>) {
    USER_FORMULA_READS.with(|held| held.borrow_mut().extend(reads));
}

fn clear_user_function_reads() {
    USER_FORMULA_READS.with(|held| held.borrow_mut().clear());
}

/// The host queued a callback; suspend the rest of this evaluation pass.
pub fn defer_user_function() { USER_FUNCTION_PENDING.with(|pending| pending.set(true)); }
pub fn clear_user_function_pending() { USER_FUNCTION_PENDING.with(|pending| pending.set(false)); }
pub fn set_user_function_wait_for_dependencies(wait: bool) { USER_FUNCTION_WAIT_FOR_DEPENDENCIES.with(|held| held.set(wait)); }
pub fn user_function_wait_for_dependencies() -> bool { USER_FUNCTION_WAIT_FOR_DEPENDENCIES.with(std::cell::Cell::get) }
pub fn set_user_function_wait_for_dynamic_lookups(wait: bool) { USER_FUNCTION_WAIT_FOR_DYNAMIC_LOOKUPS.with(|held| held.set(wait)); }
pub fn user_function_wait_for_dynamic_lookups() -> bool { USER_FUNCTION_WAIT_FOR_DYNAMIC_LOOKUPS.with(std::cell::Cell::get) }
pub fn set_user_function_provisional_dynamic_text_value(allow: bool) { USER_FUNCTION_PROVISIONAL_DYNAMIC_TEXT_VALUE.with(|held| held.set(allow)); }
pub fn user_function_provisional_dynamic_text_value() -> bool { USER_FUNCTION_PROVISIONAL_DYNAMIC_TEXT_VALUE.with(std::cell::Cell::get) }
pub fn set_user_function_provisional_zero_text_probe(allow: bool) { USER_FUNCTION_PROVISIONAL_ZERO_TEXT_PROBE.with(|held| held.set(allow)); }
pub fn user_function_provisional_zero_text_probe() -> bool { USER_FUNCTION_PROVISIONAL_ZERO_TEXT_PROBE.with(std::cell::Cell::get) }
pub fn set_user_function_prioritize_dynamic_indirect_consumers(prioritize: bool) { USER_FUNCTION_PRIORITIZE_DYNAMIC_INDIRECT_CONSUMERS.with(|held| held.set(prioritize)); }
pub fn user_function_prioritize_dynamic_indirect_consumers() -> bool { USER_FUNCTION_PRIORITIZE_DYNAMIC_INDIRECT_CONSUMERS.with(std::cell::Cell::get) }
pub fn set_user_function_provisional_text_probe_target(target: Option<FormulaCell>) { USER_FUNCTION_PROVISIONAL_TEXT_PROBE_TARGET.with(|held| *held.borrow_mut() = target); }
pub fn user_function_provisional_text_probe_target() -> Option<FormulaCell> { USER_FUNCTION_PROVISIONAL_TEXT_PROBE_TARGET.with(|held| held.borrow().clone()) }
pub fn set_user_function_full_rebuild(rebuild: bool) { USER_FUNCTION_FULL_REBUILD.with(|held| held.set(rebuild)); }
pub fn user_function_full_rebuild() -> bool { USER_FUNCTION_FULL_REBUILD.with(std::cell::Cell::get) }
pub fn set_user_function_full_calculation(full: bool) { USER_FUNCTION_FULL_CALCULATION.with(|held| held.set(full)); }
pub fn user_function_full_calculation() -> bool { USER_FUNCTION_FULL_CALCULATION.with(std::cell::Cell::get) }
pub fn set_user_function_distinct_provisional_calls(distinct: bool) { USER_FUNCTION_DISTINCT_PROVISIONAL_CALLS.with(|held| held.set(distinct)); }
pub fn user_function_distinct_provisional_calls() -> bool { USER_FUNCTION_DISTINCT_PROVISIONAL_CALLS.with(std::cell::Cell::get) }
fn user_function_pending() -> bool { USER_FUNCTION_PENDING.with(|pending| pending.get()) }

fn reset_user_function_occurrence() {
    USER_FUNCTION_OCCURRENCE.with(|occurrence| occurrence.set(0));
    USER_PROVISIONAL_INDIRECT_TEXT_VALUE.with(|held| held.set(false));
}

fn next_user_function_context(provisional: bool) -> UserFunctionCallContext {
    let occurrence = USER_FUNCTION_OCCURRENCE.with(|held| {
        let occurrence = held.get();
        held.set(occurrence.saturating_add(1));
        occurrence
    });
    UserFunctionCallContext {
        occurrence,
        provisional: provisional && USER_FUNCTION_DISTINCT_PROVISIONAL_CALLS.with(std::cell::Cell::get),
        provisional_indirect_text_value: USER_PROVISIONAL_INDIRECT_TEXT_VALUE.with(|held| held.replace(false)),
        dynamic_reference_argument_mask: 0,
    }
}

fn omit_xlookup_return_arrays(expr: &Expr) -> Expr {
    match expr {
        Expr::Unary { op, operand } => Expr::Unary { op: *op, operand: Box::new(omit_xlookup_return_arrays(operand)) },
        Expr::Binary { op, lhs, rhs } => Expr::Binary {
            op: *op,
            lhs: Box::new(omit_xlookup_return_arrays(lhs)),
            rhs: Box::new(omit_xlookup_return_arrays(rhs)),
        },
        Expr::Function { name, args } => Expr::Function {
            name: name.clone(),
            args: args.iter().enumerate().map(|(index, arg)| {
                if name == "XLOOKUP" && index == 2 { Expr::Literal(Value::Blank) }
                else { omit_xlookup_return_arrays(arg) }
            }).collect(),
        },
        _ => expr.clone(),
    }
}

fn ask_user_function(name: &str, args: Vec<UserArg>, caller: UserCaller, context: UserFunctionCallContext, view: Option<UserFunctionReadView>, root: bool, written: Vec<(String, RangeRef)>) -> Option<Arg> {
    struct Restore(Option<UserFunctionReadView>, bool, Vec<(String, RangeRef)>, UserFunctionCallContext);
    impl Drop for Restore {
        fn drop(&mut self) {
            USER_FUNCTION_READ_VIEW.with(|held| *held.borrow_mut() = self.0.take());
            USER_FUNCTION_AT_ROOT.with(|held| held.set(self.1));
            USER_FUNCTION_WRITTEN_DEPENDENCIES.with(|held| *held.borrow_mut() = std::mem::take(&mut self.2));
            USER_FUNCTION_CONTEXT.with(|held| held.set(self.3));
        }
    }
    let previous = USER_FUNCTION_READ_VIEW.with(|held| held.replace(view));
    let previous_root = USER_FUNCTION_AT_ROOT.with(|held| held.replace(root));
    let previous_written = USER_FUNCTION_WRITTEN_DEPENDENCIES.with(|held| held.replace(written));
    let previous_context = USER_FUNCTION_CONTEXT.with(|held| held.replace(context));
    let _restore = Restore(previous, previous_root, previous_written, previous_context);
    USER_FUNCTIONS.with(|held| match held.try_borrow_mut() {
        Ok(mut hook) => hook.as_mut().and_then(|hook| hook(name, args, caller)),
        Err(_) => None,
    })
}

/// What a range read inside SUBTOTAL passes over.
#[derive(Clone, Copy)]
enum Skip {
    Nothing,
    /// Other SUBTOTALs and the filtered rows; with `by_hand`, every hidden row.
    Subtotals { by_hand: bool },
    /// AGGREGATE's options: 0 to 3 pass over nested SUBTOTALs, and 1, 3, 5
    /// and 7 over every hidden row.
    Aggregate { nested: bool, hidden: bool },
}
use crate::lexer::ParseError;
use crate::parser::parse;
mod recalculation;
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
        source_utf16: Option<Vec<u16>>,
        expr: Expr,
        cached: Value,
        /// Where this cell sits in the block its ARRAY formula is dealt
        /// across, as (columns, rows) from the block's top-left; None for a
        /// formula of its own.
        share: Option<(u32, u32)>,
    },
}

fn formula_source_value(source: &str, source_utf16: Option<&[u16]>) -> Value {
    if let Some(units) = source_utf16 {
        let shown = crate::parser::transform_utf16_formula(units, crate::lexer::formula_from_ooxml).unwrap_or_else(|_| units.to_vec());
        let units = shown.as_slice();
        let body = units.strip_prefix(&[b'=' as u16]).unwrap_or(units);
        let mut formula = Vec::with_capacity(body.len() + 1);
        formula.push(b'=' as u16);
        formula.extend_from_slice(body);
        Value::from_utf16(formula)
    } else {
        let shown = crate::lexer::formula_from_ooxml(source).unwrap_or_else(|_| source.to_string());
        Value::Text(format!("={}", shown.strip_prefix('=').unwrap_or(&shown)))
    }
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
    blocked_spills: BTreeMap<(u32, u32), (u32, u32)>,
    /// The values spilled into the cells beside them, with the formula that
    /// spilled each -- None for one remembered from before and not yet
    /// worked out again.
    spilled: BTreeMap<(u32, u32), (Value, Option<(u32, u32)>)>,
    /// Each dressed cell's number format and horizontal alignment, by
    /// `(col, row)`, for CELL to read.
    dress: BTreeMap<(u32, u32), (Option<String>, Option<String>)>,
    /// The file's column widths, zero-based, and the width of any other.
    widths: (Vec<f32>, f32),
    custom_widths: BTreeMap<u32, bool>,
    default_width_explicit: bool,
    hidden_columns: BTreeSet<u32>,
}

thread_local! {
    static DIGIT_WIDTH: std::cell::Cell<f64> = const { std::cell::Cell::new(7.0) };
}

/// The width in pixels of a digit of the workbook's Normal font, which
/// CELL("width") needs to turn a file's column width back into characters.
pub fn set_digit_width(pixels: f64) {
    if pixels > 0.0 {
        DIGIT_WIDTH.with(|held| held.set(pixels));
    }
}

/// What a recalculation did.
#[derive(Debug, Clone, Default, PartialEq, Eq)]
pub struct RecalcReport {
    pub evaluated: usize,
    /// Formula cells visited by the pass, including cells retaining a cyclic cache.
    pub calculated: Vec<(String, (u32, u32))>,
    /// First accepted visits, before consumers refreshed after a spill changed.
    /// Normal calculation learns this traversal, rather than the refresh order.
    pub traversal_order: Vec<FormulaCell>,
    /// Dependencies learned by scalar value reads or waits for spill owners.
    /// Hosts retain these local chain repairs, not the whole completion order.
    pub dependency_order_repairs: Vec<(FormulaCell, FormulaCell)>,
    pub blocked_spills: Vec<((String, (u32, u32)), (u32, u32))>,
    pub deferred: bool,
    /// Cells that take part in a reference cycle, as `Sheet!A1`.
    ///
    /// With iterative calculation off, Excel leaves these at `0` and raises a
    /// warning rather than producing an error value; that is reproduced here,
    /// and the cells are listed so the caller can surface the warning.
    pub circular: Vec<String>,
}

#[derive(Debug, Clone, Copy)]
struct ArrayGroup {
    anchor: (u32, u32),
    width: u32,
    height: u32,
}

#[derive(Debug, Clone)]
enum ArrayResult {
    Shared(std::rc::Rc<Arg>),
    Lifted { occurrence_start: u64 },
}

#[derive(Debug, Clone, Default)]
pub struct Workbook {
    calculation_order: BTreeMap<(String, (u32, u32)), usize>,
    chain_probes: bool,
    read_history: std::cell::RefCell<Option<std::rc::Rc<std::cell::RefCell<UserReadHistory>>>>,
    prepared_spills: BTreeMap<FormulaCell, Arg>,
    array_groups: BTreeMap<FormulaCell, ArrayGroup>,
    array_results: BTreeMap<FormulaCell, ArrayResult>,
    array_context: std::cell::Cell<Option<ArrayGroup>>,
    array_materialization: std::cell::Cell<u32>,
    array_lifted: std::cell::Cell<bool>,
    array_projection: std::cell::Cell<Option<(u32, u32)>>,
    array_entry: bool,
    /// Consumers whose arguments cross a changing spill footprint.
    /// Their provisional and published reads require separate callbacks.
    scalar_replay_exclusions: BTreeSet<FormulaCell>,
    // A chain probe may hand an unready reference to a UDF, but consuming
    // its value (or that UDF's result) suspends the rest of the expression.
    provisional_value: std::cell::Cell<bool>,
    value_deferred: std::cell::Cell<bool>,
    dynamic_reference_wait: std::cell::Cell<bool>,
    formula_values_before_pass: std::cell::RefCell<BTreeMap<FormulaCell, Value>>,
    provisional_text_formula_values_before_pass: std::cell::RefCell<BTreeMap<FormulaCell, Value>>,
    variant_value_read_context: std::cell::Cell<bool>,
    /// Distinguish waiting before a callback from retrying a completed callback.
    callback_answer_observed: std::cell::Cell<bool>,
    reference_context: std::cell::Cell<bool>,
    unready_cells: std::cell::RefCell<BTreeSet<FormulaCell>>,
    /// Formula cells selected for this pass, including completed precedents.
    pass_cells: std::cell::RefCell<BTreeSet<FormulaCell>>,
    compatibility_version: u8,
    dbcs_language: Option<bool>,
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
    let_scope: std::cell::RefCell<Vec<(String, LocalBinding)>>,
    /// The LAMBDAs a LET has given a name, innermost last, so that
    /// `LET(f,LAMBDA(x,x*2),f(4))` can call one by it.
    lambda_scope: std::cell::RefCell<Vec<(String, Expr)>>,
    /// The optional LAMBDA parameters (`[b]`) a call left out, innermost
    /// last, for ISOMITTED to find.
    omitted: std::cell::RefCell<Vec<String>>,
    /// The state RAND draws from: seeded from the moment on first use, and
    /// moved on by every draw.
    rand_state: std::cell::Cell<u64>,
    /// A cell being watched while its own formula is worked out, and whether
    /// that formula read it: a formula naming its own cell is only circular
    /// when the reading actually happens.
    watched: std::cell::RefCell<Option<(String, (u32, u32))>>,
    watched_read: std::cell::Cell<bool>,
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
    pub fn set_dbcs_language(&mut self, enabled: bool) { self.dbcs_language=Some(enabled); }
    pub fn set_compatibility_version(&mut self, version: u8) {
        self.compatibility_version = version.max(1);
    }
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

    /// Enable calculation-chain attempts; formula entry keeps dependency order.
    pub fn set_chain_probes(&mut self, enabled: bool) { self.chain_probes = enabled; }

    pub fn new() -> Workbook {
        Workbook::default()
    }

    pub fn add_sheet(&mut self, name: &str) {
        self.sheets.entry(name.to_string()).or_default();
        if !self.sheet_order.iter().any(|held| held == name) {
            self.sheet_order.push(name.to_string());
        }
    }

    /// Say how a cell (0-based) is dressed: its number format and its
    /// horizontal alignment, which CELL("format") and CELL("prefix") read.
    pub fn set_cell_dress(&mut self, sheet: &str, col: u32, row: u32, format: Option<&str>, align: Option<&str>) {
        let held = self.sheets.entry(sheet.to_string()).or_default();
        held.dress.insert((col, row), (format.map(str::to_string), align.map(str::to_string)));
    }

    /// The sheet's column widths as the file keeps them, zero-based, and the
    /// width of a column it does not list.
    pub fn set_column_widths(&mut self, sheet: &str, widths: Vec<f32>, default: f32) {
        let held = self.sheets.entry(sheet.to_string()).or_default();
        held.widths = (widths, default);
    }

    pub fn set_custom_column_widths(&mut self, sheet: &str, custom: BTreeMap<u32, bool>) {
        self.sheets.entry(sheet.to_string()).or_default().custom_widths = custom;
    }
    pub fn set_column_width_state(&mut self, sheet: &str, explicit_default: bool, hidden: BTreeSet<u32>) {
        let held = self.sheets.entry(sheet.to_string()).or_default();
        held.default_width_explicit = explicit_default;
        held.hidden_columns = hidden;
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
            crate::case_fold::table_key(name),
            TableRef {
                sheet: sheet.to_string(),
                first_row: rows.0,
                last_row: rows.1,
                first_col: cols.0,
                last_col: cols.1,
                header_rows,
                totals_rows,
                headings: headings.iter().map(|one| crate::case_fold::table_key(one)).collect(),
            },
        );
    }

    /// Define a workbook-scoped name, e.g. `TAX_RATE` → `0.1`.
    /// What a name stands for, read from the sheet asking: that sheet's own
    /// name of that spelling first, then the workbook's. Measured on a copied
    /// sheet: `=loc+mine` reads the copy's own `loc` and `mine`.
    fn name_bound(&self, name: &str, sheet: &str, at: At) -> Option<Expr> {
        self.name_definition(name, sheet).map(|(_, body)| body.for_defined_name(at.unwrap_or((0, 0))))
    }

    fn name_definition(&self, name: &str, sheet: &str) -> Option<(&String, &Expr)> {
        let explicit = name.rsplit_once('!').map(|(owner, leaf)| {
            (owner.strip_prefix('\'').and_then(|s| s.strip_suffix('\''))
                .map(|s| s.replace("''", "'")).unwrap_or_else(|| owner.to_string()), leaf)
        });
        let upper = crate::case_fold::name_key(explicit.as_ref().map(|(_, leaf)| *leaf).unwrap_or(name));
        let scope = crate::case_fold::name_key(explicit.as_ref().map(|(owner, _)| owner.as_str()).unwrap_or(sheet));
        self.names
            .get_key_value(&format!("{scope}!{upper}"))
            .or_else(|| self.names.get_key_value(&format!("'{}'!{upper}", scope.replace('\'', "''"))))
            .or_else(|| self.names.get_key_value(&upper))
    }

    pub fn define_name(&mut self, name: &str, formula: &str) -> Result<(), CalcError> {
        // A name may stand for several blocks written as a list,
        // `=Sheet1!$C$1,Sheet1!$C$3:$C$4`, which a formula would have to
        // bracket to read as the union it is: measured, SUM of such a name
        // adds both blocks.
        let body = formula.trim();
        let body = body.strip_prefix('=').unwrap_or(body);
        let expr = if has_top_level_comma(body) {
            parse(&format!("=({body})"))?
        } else {
            parse(formula)?
        };
        self.names.insert(crate::case_fold::name_key(name), expr);
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
                source_utf16: None,
                expr,
                cached: Value::Blank,
                share: None,
            },
        );
        Ok(())
    }

    /// Place a formula whose source may contain unpaired UTF-16 units.
    pub fn set_formula_utf16(&mut self, sheet: &str, a1: &str, formula: &[u16]) -> Result<(), CalcError> {
        if let Ok(text) = String::from_utf16(formula) { return self.set_formula(sheet, a1, &text); }
        let cell = self.cell_ref(a1)?;
        let expr = crate::parser::parse_utf16(formula)?;
        self.sheet_mut(sheet)?.cells.insert(cell.coord(), Cell::Formula {
            source: String::from_utf16_lossy(formula),
            source_utf16: Some(formula.to_vec()),
            expr,
            cached: Value::Blank,
            share: None,
        });
        Ok(())
    }

    /// Retain the displayed source after compiling an evaluation expression
    /// with implicit coercions. The expression and cached value are unchanged.
    pub fn remember_formula_source(&mut self, sheet: &str, a1: &str, original: &str, original_utf16: Option<&[u16]>) -> Result<(), CalcError> {
        let cell = self.cell_ref(a1)?;
        if let Some(Cell::Formula { source, source_utf16, .. }) = self.sheet_mut(sheet)?.cells.get_mut(&cell.coord()) {
            *source = original.to_string();
            *source_utf16 = original_utf16.map(<[u16]>::to_vec);
        }
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

    /// Restore a cached spill's geometry and ownership without evaluating
    /// its formula. Partial calculation can then read and preserve it.
    pub fn set_cached_spill(&mut self, sheet: &str, a1: &str, width: u32, height: u32) -> Result<(), CalcError> {
        let anchor = self.cell_ref(a1)?;
        if width == 0 || height == 0 || width - 1 > MAX_COL - anchor.col || height - 1 > MAX_ROW - anchor.row {
            return Err(CalcError::BadReference(a1.to_string()));
        }
        let held = self.sheet_mut(sheet)?;
        held.spills.insert(anchor.coord(), (width, height));
        for ((col, row), (_, owner)) in &mut held.spilled {
            if *col >= anchor.col && *col - anchor.col < width && *row >= anchor.row && *row - anchor.row < height {
                *owner = Some(anchor.coord());
            }
        }
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
                source_utf16: None,
                expr,
                cached: Value::Blank,
                share: Some(offset),
            },
        );
        Ok(())
    }

    /// Place a member of an array formula with exact UTF-16 source.
    pub fn set_array_member_utf16(&mut self, sheet: &str, a1: &str, formula: &[u16], offset: (u32, u32)) -> Result<(), CalcError> {
        if let Ok(text) = String::from_utf16(formula) { return self.set_array_member(sheet, a1, &text, offset); }
        let cell = self.cell_ref(a1)?;
        let expr = crate::parser::parse_utf16(formula)?;
        self.sheet_mut(sheet)?.cells.insert(cell.coord(), Cell::Formula {
            source: String::from_utf16_lossy(formula),
            source_utf16: Some(formula.to_vec()),
            expr,
            cached: Value::Blank,
            share: Some(offset),
        });
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

    /// The exact UTF-16 source of a formula, including unpaired units.
    pub fn formula_utf16(&self, sheet: &str, a1: &str) -> Option<Vec<u16>> {
        let cell = self.cell_ref(a1).ok()?;
        match self.sheets.get(sheet)?.cells.get(&cell.coord())? {
            Cell::Formula { source, source_utf16, .. } => Some(source_utf16.clone().unwrap_or_else(|| source.encode_utf16().collect())),
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
    /// The last of the four says whether the answer came as a block -- an
    /// array, however small -- rather than as one value.
    pub fn evaluate_block(&self, sheet: &str, formula: &str) -> Result<(usize, usize, Vec<Value>, bool), CalcError> {
        Ok(match self.evaluate_argument(sheet, formula)? {
            Arg::Range(block) => { let array = block.is_array(); (block.width, block.height, block.cells, array) },
            Arg::Value(value) => (1, 1, vec![value], false),
        })
    }

    pub fn evaluate_argument(&self, sheet: &str, formula: &str) -> Result<Arg, CalcError> {
        let expr = parse(formula)?;
        let mut answer = self.eval_arg(&expr, sheet, 0, None);
        match &mut answer {
            Arg::Range(block) => for value in &mut block.cells { *value = formula_result(value.clone()); },
            Arg::Value(value) => *value = formula_result(value.clone()),
        }
        Ok(answer)
    }

    /// The cells a formula names when what it works out to is a reference
    /// -- `INDEX(A1:A6,2)`, `OFFSET(A1,1,0)`, `INDIRECT("B2")` -- as the
    /// sheet and range; None for anything else.
    pub fn evaluate_reference(&self, sheet: &str, formula: &str) -> Option<(String, RangeRef)> {
        match self.evaluate_argument(sheet, formula).ok()? {
            Arg::Range(block) => block.cell_reference(),
            _ => None,
        }
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
        let layout = self.spill_blocks();
        let report = self.work_out(Extent::Everything);
        self.settle_changed_spills(layout, report, None)
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
        let layout = self.spill_blocks();
        let report = self.work_out(Extent::ReachedBy(changed));
        self.settle_changed_spills(layout, report, None)
    }

    /// Normal dirty calculation also follows cells gained or lost by spills.
    /// An explicit range calculation instead keeps its requested boundary.
    pub fn recalculate_dirty(&mut self, wanted: &[(String, (u32, u32))], sheet_scope: Option<&str>) -> RecalcReport {
        self.recalculate_dirty_with_dynamic_reference_wait(wanted, sheet_scope, false)
    }

    pub fn recalculate_dirty_with_dynamic_reference_wait(&mut self, wanted: &[(String, (u32, u32))], sheet_scope: Option<&str>, wait: bool) -> RecalcReport {
        let layout = self.spill_blocks();
        let previous = self.dynamic_reference_wait.replace(wait);
        let report = self.work_out(Extent::JustThese(wanted));
        self.dynamic_reference_wait.set(previous);
        self.settle_changed_spills(layout, report, sheet_scope)
    }

    fn settle_changed_spills(&mut self, mut layout: Vec<(String, (u32, u32), (u32, u32))>, mut report: RecalcReport, sheet_scope: Option<&str>) -> RecalcReport {
        for _ in 0..3 {
            let now = self.spill_blocks();
            if now == layout { break; }
            let before: BTreeMap<_, _> = layout.iter().map(|(sheet, at, size)| ((sheet.clone(), *at), *size)).collect();
            let after: BTreeMap<_, _> = now.iter().map(|(sheet, at, size)| ((sheet.clone(), *at), *size)).collect();
            let anchors: BTreeSet<_> = before.keys().chain(after.keys()).cloned().collect();
            let positions: BTreeMap<_, _> = report.calculated.iter().enumerate().map(|(index, cell)| (cell.clone(), index)).collect();
            let mut stale = BTreeSet::new();
            for anchor in anchors {
                let old = before.get(&anchor).copied().unwrap_or((1, 1));
                let new = after.get(&anchor).copied().unwrap_or((1, 1));
                if old == new { continue; }
                // Only cells gained or lost by this owner changed geometry.
                // Their intersection, including the anchor, is unchanged.
                let mut strips = Vec::new();
                for (outer, inner) in [(old, new), (new, old)] {
                    if outer.0 > inner.0 { strips.push((anchor.1.0 + inner.0, anchor.1.1, anchor.1.0 + outer.0 - 1, anchor.1.1 + outer.1 - 1)); }
                    if outer.1 > inner.1 { strips.push((anchor.1.0, anchor.1.1 + inner.1, anchor.1.0 + outer.0.min(inner.0) - 1, anchor.1.1 + outer.1 - 1)); }
                }
                for cell in self.formula_keys() {
                    if sheet_scope.is_some_and(|name| cell.0 != name) { continue; }
                    // A consumer already accepted after this producer saw
                    // the new values. Repeating it would repeat VBA effects.
                    if positions.get(&cell).zip(positions.get(&anchor)).is_some_and(|(consumer, producer)| consumer > producer) { continue; }
                    let Some(expr) = self.expr_at(&cell.0, &cell.1) else { continue; };
                    let reads = self.dependency_expression(&expr, &cell.0, Some(cell.1)).invalidation_references();
                    if reads.iter().any(|reference| reference.sheet.as_deref().unwrap_or(&cell.0) == anchor.0
                        && strips.iter().any(|(left, top, right, bottom)| reference.range.start.col <= *right && reference.range.end.col >= *left && reference.range.start.row <= *bottom && reference.range.end.row >= *top)) {
                        stale.insert(cell);
                    }
                }
            }
            layout = now;
            if stale.is_empty() { break; }
            let stale: Vec<_> = stale.into_iter().collect();
            let changed: Vec<_> = stale.iter().map(|(sheet, (col, row))|
                (sheet.clone(), RangeRef::single(CellRef::new(*col, *row)))).collect();
            let mut reached = self.dependent_formula_ranges(&changed);
            if let Some(name) = sheet_scope { reached.retain(|(sheet, _)| sheet == name); }
            if report.deferred {
                // Retain newly required work across the host's callback
                // suspension. No values are consumed from unfinished calls.
                for cell in reached {
                    if !report.calculated.contains(&cell) { report.calculated.push(cell); report.evaluated += 1; }
                }
                break;
            }
            let next = self.work_out(Extent::JustThese(&reached));
            let visited: BTreeSet<_> = next.calculated.iter().cloned().collect();
            report.calculated.retain(|cell| !visited.contains(cell));
            report.calculated.extend(next.calculated);
            let mut seen: BTreeSet<_> = report.traversal_order.iter().cloned().collect();
            for cell in next.traversal_order {
                if seen.insert(cell.clone()) { report.traversal_order.push(cell); }
            }
            for repair in next.dependency_order_repairs {
                if !report.dependency_order_repairs.contains(&repair) { report.dependency_order_repairs.push(repair); }
            }
            report.evaluated += next.evaluated;
            report.blocked_spills = next.blocked_spills;
            report.deferred = next.deferred;
            for circular in next.circular { if !report.circular.contains(&circular) { report.circular.push(circular); } }
        }
        report
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

    /// Evaluate newly assigned legacy array formulas as complete array results.
    pub fn recalculate_array_entry(&mut self, wanted: &[(String, (u32, u32))]) -> RecalcReport {
        let previous = std::mem::replace(&mut self.array_entry, true);
        let report = self.work_out(Extent::JustThese(wanted));
        self.array_entry = previous;
        report
    }

    /// The whole of it, or some part.
    fn work_out(&mut self, extent: Extent<'_>) -> RecalcReport {
        *self.read_history.get_mut() = None;
        self.prepared_spills.clear();
        self.array_results.clear();
        self.array_context.set(None);
        self.array_materialization.set(0);
        self.array_lifted.set(false);
        self.array_projection.set(None);
        self.array_groups = self.collect_array_groups();
        self.scalar_replay_exclusions.clear();
        if !self.now_pinned {
            if let Some(moment) = Workbook::read_the_clock() {
                self.now = moment;
            }
        }
        if let Ok(today) = crate::datetime::date_from_serial(self.now.floor() as i64) {
            crate::datetime::set_this_year(today.year);
        }
        let keys = self.formula_keys();
        *self.formula_values_before_pass.borrow_mut() = if self.dynamic_reference_wait.get() || user_function_wait_for_dynamic_lookups() {
            keys.iter().map(|(sheet, coord)| ((sheet.clone(), *coord), self.value_at(sheet, coord.0, coord.1))).collect()
        } else { BTreeMap::new() };
        *self.provisional_text_formula_values_before_pass.borrow_mut() = if user_function_provisional_dynamic_text_value() {
            keys.iter().map(|(sheet, coord)| ((sheet.clone(), *coord), self.value_at(sheet, coord.0, coord.1))).collect()
        } else { BTreeMap::new() };
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
            self.each_dependency(sheet, coord, &at, true, &mut |j| {
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
                let mut invalidation = vec![BTreeSet::new(); keys.len()];
                for (i, (sheet, coord)) in keys.iter().enumerate() {
                    self.each_dependency(sheet, coord, &at, false, &mut |j| { invalidation[j].insert(i); });
                }
                let reached = self.reached_by(changed, &keys, &at, &invalidation);
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
            calculated: order.iter().map(|&i| keys[i].clone()).collect(),
            traversal_order: Vec::new(),
            dependency_order_repairs: Vec::new(),
            blocked_spills: Vec::new(),
            deferred: false,
            circular: Vec::new(),
        };


        *self.pass_cells.borrow_mut() = order.iter().map(|&i| keys[i].clone()).collect();
        let chain = self.chain_probes && !self.calculation_order.is_empty();
        let mut accepted = Vec::new();
        let mut dependency_order_repairs = BTreeSet::new();
        let mut temporary_value_dependencies = false;
        if chain {
            let selected: BTreeSet<usize> = order.iter().copied().collect();
            let mut requirements = vec![BTreeSet::new(); keys.len()];
            let mut direct_requirements = vec![BTreeSet::new(); keys.len()];
            for &i in &selected {
                let (sheet, coord) = &keys[i];
                self.each_dependency(sheet, coord, &at, true, &mut |j| {
                    if j != i && selected.contains(&j) { requirements[i].insert(j); }
                });
                self.each_dependency_with_spills(sheet, coord, &at, true, false, &mut |j| {
                    if j != i && selected.contains(&j) { direct_requirements[i].insert(j); }
                });
                if let Some(expr) = self.expr_at(sheet, coord) {
                    expr.visit(&mut |node| {
                        let Expr::Function { name, args } = node else { return; };
                        for (index, argument) in args.iter().enumerate() {
                            let Some(reference) = Self::scalar_cell_reference(name, index, argument) else { continue; };
                            let target = reference.sheet.as_deref().unwrap_or(sheet);
                            for &j in &direct_requirements[i] {
                                if crate::case_fold::name_key(&keys[j].0) == crate::case_fold::name_key(target)
                                    && reference.range.contains(keys[j].1.0, keys[j].1.1) {
                                    // Keep already ordered scalar links too:
                                    // delaying one consumer must not move its
                                    // own consumers ahead of it.
                                    dependency_order_repairs.insert((j, i));
                                }
                            }
                        }
                    });
                }
            }
            for &i in &selected {
                if !reads_itself.contains(&i) {
                    let (sheet, coord) = &keys[i]; self.store_cached(sheet, coord, Value::Blank);
                }
            }
            // An array whose result waits on VBA has no published member
            // values in the provisional view. Keep its known ownership for
            // ordering, while callers probing those values observe Empty.
            let mut unpublished = Vec::new();
            for &i in &selected {
                let (sheet, coord) = &keys[i];
                if !self.sheets.get(sheet).is_some_and(|held| held.spills.contains_key(coord)) { continue; }
                let mut walking = vec![i];
                let mut seen = BTreeSet::new();
                let mut user = false;
                while let Some(j) = walking.pop() {
                    if !seen.insert(j) { continue; }
                    let (target, cell) = &keys[j];
                    if let Some(expr) = self.expr_at(target, cell) {
                        self.dependency_expression(&expr, target, Some(*cell)).visit(&mut |node| {
                            if let Expr::Function { name, .. } = node {
                                if !functions::is_known_function(name) && !matches!(name.as_str(), "_UNION" | "_ISECT" | "_SPILL" | "_SHEETS" | "_INVOKE") { user = true; }
                            }
                        });
                    }
                    if user { break; }
                    walking.extend(requirements[j].iter().copied());
                }
                if user { unpublished.push((sheet.clone(), *coord)); }
            }
            for (sheet, anchor) in unpublished {
                let history = self.read_history.borrow().clone();
                if let Some(held) = self.sheets.get_mut(&sheet) {
                    for (coord, (value, owner)) in &mut held.spilled {
                        if *owner == Some(anchor) {
                            if let Some(history) = &history { history.borrow_mut().change(&sheet, *coord, value, &Value::Blank); }
                            *value = Value::Blank;
                        }
                    }
                }
            }
            let mut read_cells = vec![BTreeSet::new(); keys.len()];
            *self.unready_cells.borrow_mut() = selected.iter().map(|&i| keys[i].clone()).collect();
            let mut remaining = selected;
            let mut completed = BTreeSet::new();
            let mut attempted: BTreeMap<usize, BTreeSet<usize>> = BTreeMap::new();
            let mut waiting = BTreeSet::new();
            let mut held_back = BTreeSet::new();
            let mut deferred = false;
            // A deterministic array can discover its new footprint before
            // publishing values. Cells gained or lost by that footprint are
            // probed against Empty, then evaluated after publication. This
            // remains two actual calls even when both arguments coerce to 0.
            let mut geometry_probes = BTreeSet::new();
            for &producer in &remaining {
                let (sheet, anchor) = &keys[producer];
                if !requirements[producer].is_empty() || !self.sheets.get(sheet).is_some_and(|held| held.dynamic.contains(anchor)) { continue; }
                let Some(expr) = self.expr_at(sheet, anchor) else { continue; };
                if !self.dependency_value_is_fixed(&expr, sheet) { continue; }
                let old = self.sheets.get(sheet).and_then(|held| held.spills.get(anchor)).copied().unwrap_or((1, 1));
                reset_user_function_occurrence();
                let worked = self.eval_arg(&expr, sheet, 0, Some(*anchor));
                let mut shadow = worked.clone();
                match &mut shadow {
                    Arg::Value(value) => *value = Value::Blank,
                    Arg::Range(block) => block.cells.fill(Value::Blank),
                }
                self.clear_spill(sheet, *anchor);
                let _ = self.publish_spill(sheet, *anchor, shadow, true);
                self.prepared_spills.insert((sheet.clone(), *anchor), worked);
                let new = self.sheets.get(sheet).and_then(|held| held.spills.get(anchor)).copied().unwrap_or((1, 1));
                if old == new { continue; }
                let mut strips = Vec::new();
                for (outer, inner) in [(old, new), (new, old)] {
                    if outer.0 > inner.0 { strips.push((anchor.0 + inner.0, anchor.1, anchor.0 + outer.0 - 1, anchor.1 + outer.1 - 1)); }
                    if outer.1 > inner.1 { strips.push((anchor.0, anchor.1 + inner.1, anchor.0 + outer.0.min(inner.0) - 1, anchor.1 + outer.1 - 1)); }
                }
                for &consumer in &remaining {
                    if consumer == producer { continue; }
                    let (target, cell) = &keys[consumer];
                    let Some(expr) = self.expr_at(target, cell) else { continue; };
                    let resolved = self.dependency_expression(&expr, target, Some(*cell));
                    let reads_changed = resolved.invalidation_references().iter().any(|reference| reference.sheet.as_deref().unwrap_or(target) == sheet
                        && strips.iter().any(|(left, top, right, bottom)| reference.range.start.col <= *right && reference.range.end.col >= *left && reference.range.start.row <= *bottom && reference.range.end.row >= *top));
                    let mut reads_extent = false;
                    resolved.visit(&mut |node| {
                        if let Expr::Function { name, args } = node {
                            if name == "_SPILL" && args.iter().flat_map(Expr::invalidation_references).any(|reference| reference.sheet.as_deref().unwrap_or(target) == sheet && reference.range.contains(anchor.0, anchor.1)) { reads_extent = true; }
                        }
                    });
                    if reads_changed || reads_extent {
                        requirements[consumer].insert(producer);
                        dependents[producer].insert(consumer);
                        geometry_probes.insert(consumer);
                    }
                }
            }
            self.scalar_replay_exclusions.extend(geometry_probes.iter().map(|i| keys[*i].clone()));
            for i in geometry_probes {
                for &producer in requirements[i].difference(&direct_requirements[i]) {
                    if producer > i && self.spill_dependency_was_published(&keys[i], &keys[producer]) {
                        dependency_order_repairs.insert((producer, i));
                    }
                }
                let (sheet, coord) = &keys[i];
                clear_user_function_reads();
                self.provisional_value.set(true); self.value_deferred.set(false);
                reset_user_function_occurrence();
                if let Some(expr) = self.expr_at(sheet, coord) { let _ = self.eval_arg(&expr, sheet, 0, Some(*coord)); }
                self.provisional_value.set(false); self.value_deferred.set(false);
                let observed = USER_FORMULA_READS.with(|held| std::mem::take(&mut *held.borrow_mut()));
                for (read_sheet, read_coord) in observed {
                    if let Some(&j) = at.get(read_sheet.as_str()).and_then(|cells| cells.get(&read_coord)) {
                        if remaining.contains(&j) { read_cells[i].insert(j); }
                    }
                }
                attempted.insert(i, BTreeSet::new());
                if user_function_pending() { waiting.insert(i); deferred = true; }
                USER_FUNCTION_PENDING.with(|pending| pending.set(false));
            }
            loop {
                let refreshed = remaining.iter().copied().find(|i| {
                    if waiting.contains(i) || held_back.contains(i) || requirements[*i].is_subset(&completed) { return false; }
                    let view: BTreeSet<_> = requirements[*i].union(&read_cells[*i]).filter(|j| completed.contains(j)).copied().collect();
                    attempted.get(i).is_some_and(|previous| previous != &view)
                });
                // If a scalar UDF's INDIRECT text comes from a later formula
                // cell, Excel can call it once with the prior selected value
                // before publishing a dirty target. Probe that consumer first;
                // its read is still provisional and the normal dependency pass
                // resumes it after the target is committed.
                let lagging_text_consumer = if user_function_provisional_dynamic_text_value() {
                    remaining.iter().copied().find(|i| {
                        if waiting.contains(i) || held_back.contains(i) { return false; }
                        let view: BTreeSet<_> = requirements[*i].union(&read_cells[*i]).filter(|j| completed.contains(j)).copied().collect();
                        if !requirements[*i].is_subset(&completed) && attempted.get(i) == Some(&view) { return false; }
                        let (sheet, coord) = &keys[*i];
                        self.expr_at(sheet, coord).is_some_and(|expr| self.expression_has_lagging_indirect_text(&expr, sheet, Some(*coord)))
                    })
                } else { None };
                // A suspended scalar UDF selected through INDIRECT waits for
                // the producer it selected. Finish that producer before
                // advancing an independent formula from the preferred chain.
                let waiting_precedent = remaining.iter().copied().find(|i| {
                    !waiting.contains(i) && !held_back.contains(i)
                        && waiting.iter().any(|consumer| {
                            if !requirements[*consumer].contains(i) { return false; }
                            let Some(expr) = self.expr_at(&keys[*consumer].0, &keys[*consumer].1) else { return false; };
                            let mut dynamic_reference = false;
                            expr.visit(&mut |node| {
                                if matches!(node, Expr::Function { name, .. } if name == "INDIRECT") {
                                    dynamic_reference = true;
                                }
                            });
                            dynamic_reference
                        })
                        && !completed.contains(i)
                        && {
                            let view: BTreeSet<_> = requirements[*i].union(&read_cells[*i])
                                .filter(|j| completed.contains(j)).copied().collect();
                            requirements[*i].is_subset(&completed) || attempted.get(i) != Some(&view)
                        }
                });
                let provisional_target = user_function_provisional_text_probe_target().and_then(|target| {
                    remaining.iter().copied().find(|i| !waiting.contains(i) && !held_back.contains(i) && keys[*i] == target)
                });
                let independent_of_dynamic_lookup = if user_function_prioritize_dynamic_indirect_consumers() {
                    remaining.iter().copied().find(|i| {
                        if waiting.contains(i) || held_back.contains(i) { return false; }
                        let Some((sheet, coord)) = keys.get(*i) else { return false; };
                        let Some(expr) = self.expr_at(sheet, coord) else { return false; };
                        let mut dynamic_lookup = false;
                        expr.visit(&mut |node| {
                            if matches!(node, Expr::Function { name, .. } if name == "INDIRECT") { dynamic_lookup = true; }
                        });
                        if dynamic_lookup { return false; }
                        let view: BTreeSet<_> = requirements[*i].union(&read_cells[*i]).filter(|j| completed.contains(j)).copied().collect();
                        requirements[*i].is_subset(&completed) || attempted.get(i) != Some(&view)
                    })
                } else { None };
                let candidate = provisional_target.or(independent_of_dynamic_lookup).or(lagging_text_consumer).or(refreshed).or(waiting_precedent).or_else(|| remaining.iter().copied().find(|i| {
                    if waiting.contains(i) || held_back.contains(i) { return false; }
                    let view: BTreeSet<_> = requirements[*i].union(&read_cells[*i]).filter(|j| completed.contains(j)).copied().collect();
                    requirements[*i].is_subset(&completed) || attempted.get(i) != Some(&view)
                }));
                let Some(mut i) = candidate else { break; };
                // In a forced pass, resume an earlier chain consumer after
                // its queued precedent commits before dispatching callbacks
                // farther down that chain. Those independent calls otherwise
                // overtake the suspended consumer and change its host reads.
                if matches!(extent, Extent::Everything) && deferred
                    && held_back.range(..i).any(|j| remaining.contains(j)) {
                    break;
                }
                // A ready builtin precedent can finish without dispatching VBA.
                // Finish it before probing a reference to its uncomputed value.
                let mut walking: Vec<_> = requirements[i].iter().copied().collect();
                let mut seen = BTreeSet::new();
                while let Some(j) = walking.pop() {
                    if !seen.insert(j) || !remaining.contains(&j) || waiting.contains(&j) || held_back.contains(&j) { continue; }
                    if requirements[j].is_subset(&completed) {
                        let (target, cell) = &keys[j];
                        if self.expr_at(target, cell).is_some_and(|expr| {
                            let resolved = self.dependency_expression(&expr, target, Some(*cell));
                            let mut udf = false;
                            resolved.visit(&mut |node| {
                                if let Expr::Function { name, .. } = node {
                                    if name == "INDIRECT" || (!functions::is_known_function(name) && !matches!(name.as_str(), "_UNION" | "_ISECT" | "_SPILL" | "_SHEETS" | "_INVOKE")) { udf = true; }
                                }
                            });
                            !udf
                        }) { i = j; break; }
                    } else { walking.extend(requirements[j].iter().copied()); }
                }
                let ready = requirements[i].is_subset(&completed);
                let (sheet, coord) = &keys[i];
                if ready {
                    let value = self.calculate_cell(sheet, coord, reads_itself.contains(&i));
                    self.provisional_value.set(false);
                    if self.value_deferred.get() && !user_function_pending() {
                        // A returned reference can reveal a prerequisite that
                        // is absent from the formula's written arguments. Keep
                        // it in this pass only: it is not a dirty dependency.
                        let observed = USER_FORMULA_READS.with(|held| std::mem::take(&mut *held.borrow_mut()));
                        for (read_sheet, read_coord) in observed {
                            if let Some(&j) = at.get(read_sheet.as_str()).and_then(|cells| cells.get(&read_coord)) {
                                if j != i && remaining.contains(&j) { let discovered = requirements[i].insert(j); temporary_value_dependencies |= discovered && !self.callback_answer_observed.get(); }
                            }
                        }
                        if !requirements[i].is_subset(&completed) {
                            attempted.insert(i, requirements[i].intersection(&completed).copied().collect());
                            continue;
                        }
                    }
                    if user_function_pending() {
                        waiting.insert(i);
                        // Cells after an already queued precedent wait for its result.
                        // A valid learned chain must not acquire fresh zero probes.
                        let mut downstream: Vec<_> = dependents[i].iter().copied().collect();
                        while let Some(j) = downstream.pop() {
                            if held_back.insert(j) { downstream.extend(dependents[j].iter().copied()); }
                        }
                    }
                    else { self.store_cached(sheet, coord, value); self.unready_cells.borrow_mut().remove(&(sheet.clone(), *coord)); completed.insert(i); remaining.remove(&i); accepted.push(i); }
                } else {
                    // A spill child reveals its owner when the chain reaches
                    // the consumer first. Retain that owner before this
                    // consumer on later passes, even if the spill shrinks.
                    // A ready builtin handled above never needed this wait.
                    for &producer in requirements[i].difference(&direct_requirements[i]) {
                        if producer > i && !completed.contains(&producer)
                            && self.spill_dependency_was_published(&keys[i], &keys[producer]) {
                            dependency_order_repairs.insert((producer, i));
                        }
                    }
                    clear_user_function_reads();
                    self.provisional_value.set(true); self.value_deferred.set(false);
                    reset_user_function_occurrence();
                    if let Some(expr) = self.expr_at(sheet, coord) { let _ = self.eval_arg(&expr, sheet, 0, Some(*coord)); }
                    self.provisional_value.set(false);
                    self.value_deferred.set(false);
                    let observed = USER_FORMULA_READS.with(|held| std::mem::take(&mut *held.borrow_mut()));
                    read_cells[i].clear();
                    for (read_sheet, read_coord) in observed {
                        if let Some(&j) = at.get(read_sheet.as_str()).and_then(|cells| cells.get(&read_coord)) {
                            if remaining.contains(&j) || completed.contains(&j) { read_cells[i].insert(j); }
                        }
                    }
                    let view = requirements[i].union(&read_cells[i]).filter(|j| completed.contains(j)).copied().collect();
                    attempted.insert(i, view);
                    if user_function_pending() { waiting.insert(i); }
                }
                deferred |= user_function_pending();
                // Queue independent callbacks against the same uncommitted view.
                USER_FUNCTION_PENDING.with(|pending| pending.set(false));
            }
            USER_FUNCTION_PENDING.with(|pending| pending.set(deferred));
            self.unready_cells.borrow_mut().clear();
        } else {
            for &i in &order {
                let (sheet, coord) = &keys[i];
                let value = self.calculate_cell(sheet, coord, reads_itself.contains(&i));
                if user_function_pending() { break; }
                self.store_cached(sheet, coord, value); accepted.push(i);
            }
        }
        if !user_function_pending() { report.calculated = accepted.iter().map(|&i|keys[i].clone()).collect(); }

        self.pass_cells.borrow_mut().clear();
        self.scalar_replay_exclusions.clear();
        report.blocked_spills = self.blocked_spill_ranges();
        report.deferred = user_function_pending();
        report.traversal_order = report.calculated.clone();
        report.dependency_order_repairs = dependency_order_repairs.into_iter()
            .map(|(producer, consumer)| (keys[producer].clone(), keys[consumer].clone())).collect();
        if temporary_value_dependencies && !user_function_pending() && !matches!(extent, Extent::Everything) {
            // Waiting for a computed reference before dispatching its callback
            // does not move that consumer behind unrelated completed calls.
            // Retain the incoming positions subject to the written graph;
            // the published values and actual completion report stay intact.
            let selected: BTreeSet<_> = accepted.iter().copied().collect();
            let mut indegree = vec![0usize; keys.len()];
            for &producer in &selected {
                for &consumer in &dependents[producer] {
                    if selected.contains(&consumer) { indegree[consumer] += 1; }
                }
            }
            let mut ready: BTreeSet<_> = selected.iter().copied().filter(|&i| indegree[i] == 0).collect();
            let mut learned = Vec::with_capacity(selected.len());
            while let Some(producer) = ready.pop_first() {
                learned.push(producer);
                for &consumer in &dependents[producer] {
                    if selected.contains(&consumer) {
                        indegree[consumer] -= 1;
                        if indegree[consumer] == 0 { ready.insert(consumer); }
                    }
                }
            }
            if learned.len() == accepted.len() {
                report.traversal_order = learned.iter().map(|&i| keys[i].clone()).collect();
            }
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
                report.calculated.push((sheet.clone(), *coord));
                report
                    .circular
                    .push(format!("{}!{}", sheet, CellRef::new(coord.0, coord.1).to_a1()));
                self.keep_or_zero(sheet, coord);
            }
        }

        report.traversal_order = report.calculated.clone();
        report
    }

    fn calculate_cell(&mut self, sheet: &str, coord: &(u32,u32), reads_itself: bool) -> Value {
        reset_user_function_occurrence();
        clear_user_function_reads();
        self.callback_answer_observed.set(false);
        self.provisional_value.set(false); self.value_deferred.set(false);
            // A formula naming its own cell is worked out all the same, and is
            // circular only if it really read itself: measured,
            // `=IF(D1>0,5,G1)` in G1 is 5 and `=CELL("address",L1)` $L$1,
            // while `=E1` in E1 is 0.
            if reads_itself {
                if let Some(expr) = self.expr_at(sheet, coord) {
                    *self.watched.borrow_mut() = Some((sheet.to_string(), *coord));
                    self.watched_read.set(false);
                    let worked = self.eval_arg(&expr, sheet, 0, Some(*coord));
                    *self.watched.borrow_mut() = None;
                    // A dynamic array can contain a self-reference in one
                    // element and independent answers in its other elements.
                    // Publishing only the anchor loses those answers.
                    if self.sheets.get(sheet).is_some_and(|held| held.dynamic.contains(coord))
                        && matches!(&worked, Arg::Range(block) if block.cells.len() > 1) {
                        // A circular spill keeps its first completed result,
                        // including its members, on subsequent calculations.
                        let cached_spill = self.sheets.get(sheet).is_some_and(|held|
                            held.spills.contains_key(coord)
                                && held.cells.get(coord).is_some_and(|cell| !matches!(cell.value(), Value::Blank)));
                        if self.watched_read.get() && cached_spill {
                            self.keep_or_zero(sheet, coord);
                            return self.value(sheet, &CellRef::new(coord.0, coord.1).to_a1());
                        }
                        self.clear_spill(sheet, *coord);
                        return self.publish_spill(sheet, *coord, worked, false);
                    }
                    if !self.watched_read.get() {
                        return formula_result(worked.first());
                    }
                }
                self.keep_or_zero(sheet, coord);
                return self.value(sheet, &CellRef::new(coord.0, coord.1).to_a1());
            }
            let Some(expr) = self.expr_at(sheet, coord) else {
                return Value::Blank;
            };
            let dynamic = self.sheets.get(sheet).is_some_and(|held| held.dynamic.contains(coord));
            let value = match self.share_at(sheet, coord) {
                Some(offset) => {
                    self.array_formula_value(sheet, *coord, &expr, offset)
                }
                None if dynamic => self.spill_out(sheet, *coord, &expr),
                None => formula_result(self.eval(&expr, sheet, 0, Some(*coord))),
            };
            value
    }

    fn collect_array_groups(&self) -> BTreeMap<FormulaCell, ArrayGroup> {
        let mut groups: BTreeMap<FormulaCell, (ArrayGroup, u64, bool)> = BTreeMap::new();
        for (sheet, held) in &self.sheets {
            for (&coord, cell) in &held.cells {
                let Cell::Formula { source, source_utf16, share: Some((dx, dy)), .. } = cell else { continue; };
                let (Some(col), Some(row)) = (coord.0.checked_sub(*dx), coord.1.checked_sub(*dy)) else { continue; };
                let anchor = (col, row);
                let same = matches!(held.cells.get(&anchor),
                    Some(Cell::Formula { source: root, source_utf16: units, share: Some((0, 0)), .. })
                        if source == root && source_utf16 == units);
                let entry = groups.entry((sheet.clone(), anchor))
                    .or_insert((ArrayGroup { anchor, width: 1, height: 1 }, 0, true));
                entry.0.width = entry.0.width.max(dx + 1);
                entry.0.height = entry.0.height.max(dy + 1);
                entry.1 += 1;
                entry.2 &= same;
            }
        }
        groups.into_iter().filter_map(|(key, (group, cells, valid))|
            (valid && cells == u64::from(group.width) * u64::from(group.height)).then_some((key, group))).collect()
    }

    fn array_formula_value(&mut self, sheet: &str, coord: (u32, u32), expr: &Expr, offset: (u32, u32)) -> Value {
        let anchor = (coord.0.saturating_sub(offset.0), coord.1.saturating_sub(offset.1));
        let key = (sheet.to_string(), anchor);
        let Some(group) = self.array_groups.get(&key).copied() else {
            return formula_result(element_of(&self.eval_arg(expr, sheet, 0, Some(coord)), offset));
        };
        let projected_start = match self.array_results.get(&key) {
            Some(ArrayResult::Shared(answer)) => return formula_result(element_of(answer.as_ref(), offset)),
            Some(ArrayResult::Lifted { occurrence_start }) => Some(*occurrence_start),
            None => None,
        };
        let previous = self.array_context.replace(Some(group));
        let previous_projection = self.array_projection.replace(projected_start.map(|_| offset));
        let previous_lifted = self.array_lifted.replace(false);
        if let Some(start) = projected_start { USER_FUNCTION_OCCURRENCE.with(|held| held.set(start)); }
        let mut worked = self.eval_arg(expr, sheet, 0, Some(if projected_start.is_some() { coord } else { anchor }));
        if projected_start.is_none() && !user_function_pending() && !self.value_deferred.get() {
            if self.array_lifted.get() {
                // Reserve probe occurrences even when another member is first.
                let occurrence_start = USER_FUNCTION_OCCURRENCE.with(std::cell::Cell::get);
                self.array_results.insert(key, ArrayResult::Lifted { occurrence_start });
                self.array_projection.set(Some(offset));
                USER_FUNCTION_OCCURRENCE.with(|held| held.set(occurrence_start));
                worked = self.eval_arg(expr, sheet, 0, Some(coord));
            } else {
                self.array_results.insert(key, ArrayResult::Shared(std::rc::Rc::new(worked.clone())));
            }
        }
        self.array_lifted.set(previous_lifted);
        self.array_projection.set(previous_projection);
        self.array_context.set(previous);
        formula_result(element_of(&worked, offset))
    }

    fn materialized_index_argument(expr: &Expr) -> bool {
        match expr {
            Expr::Unary { operand, .. } => Self::materialized_index_argument(operand),
            Expr::Binary { lhs, rhs, .. } =>
                Self::materialized_index_argument(lhs) || Self::materialized_index_argument(rhs),
            Expr::Function { name, args } if name == "N" && args.len() == 1 =>
                matches!(&args[0], Expr::Ref(_) | Expr::Name(_) | Expr::Table { .. })
                    || matches!(&args[0], Expr::Function { name, .. }
                        if matches!(name.as_str(), "IF" | "CHOOSE" | "INDEX" | "TAKE" | "DROP" | "XLOOKUP" | "LET" | "OFFSET" | "INDIRECT")),
            _ => false,
        }
    }

    fn prepare_array_indexes(&self, indexes: &mut [Arg], expressions: &[Expr]) -> bool {
        // The conversion belongs to this expression. LET bindings and other
        // function results are fresh selector values, even with the same cells.
        let materialized = expressions.iter().zip(indexes.iter()).any(|(expr, value)|
            Self::materialized_index_argument(expr) && matches!(value, Arg::Range(block) if block.is_array()));
        if materialized || self.array_entry || self.array_context.get().is_none() || self.array_materialization.get() != 0 {
            return materialized;
        }
        for index in indexes {
            if matches!(index, Arg::Range(block) if block.is_array() || block.cells.len() > 1) {
                self.array_lifted.set(true);
                if let Some(offset) = self.array_projection.get() {
                    *index = Arg::Value(element_of(index, offset));
                }
            }
        }
        false
    }

    fn eval_materialized_arg(&self, expr: &Expr, sheet: &str, depth: u32, skip: Skip, at: At) -> Arg {
        if self.array_context.get().is_none() { return self.eval_arg_inner(expr, sheet, depth, skip, at); }
        let previous = self.array_materialization.get();
        self.array_materialization.set(previous + 1);
        let answer = self.eval_arg_inner(expr, sheet, depth, skip, at);
        self.array_materialization.set(previous);
        answer
    }

    /// Work out a dynamic array and spill its block into the cells beside
    /// it, answering what its own cell shows. Measured in Excel: `=A1:A3*2`
    /// in D1 fills D1:D3; a cell in the way holding anything makes it
    /// #SPILL! and fills nothing; a block of one is simply that value.
    fn spill_out(&mut self, sheet: &str, anchor: (u32, u32), expr: &Expr) -> Value {
        let worked = match self.prepared_spills.remove(&(sheet.to_string(), anchor)) {
            Some(worked) => worked,
            None => self.eval_arg(expr, sheet, 0, Some(anchor)),
        };
        // A suspended callback has not replaced the array. Keep its known
        // ownership so resumed consumers still wait for the same producer.
        if user_function_pending() || !self.require_ready_value(&worked) {
            return Value::Error(ExcelError::Value);
        }
        self.clear_spill(sheet, anchor);
        self.publish_spill(sheet, anchor, worked, false)
    }

    fn clear_spill(&mut self, sheet: &str, anchor: (u32, u32)) {
        let history = self.read_history.borrow().clone();
        if let Some(held) = self.sheets.get_mut(sheet) {
            held.spills.remove(&anchor);
            if let Some(history) = &history {
                let mut history = history.borrow_mut();
                for (coord, (value, owner)) in &held.spilled {
                    if *owner == Some(anchor) { history.change(sheet, *coord, value, &Value::Blank); }
                }
            }
            held.spilled.retain(|_, (_, owner)| *owner != Some(anchor));
        }
    }

    fn publish_spill(&mut self, sheet: &str, anchor: (u32, u32), worked: Arg, provisional: bool) -> Value {
        let history = self.read_history.borrow().clone();
        if !self.require_ready_value(&worked) { return Value::Error(ExcelError::Value); }
        let block = match worked {
            Arg::Range(block) if block.cells.len() > 1 => block,
            answer => {
                if !user_function_pending() { if let Some(held) = self.sheets.get_mut(sheet) { held.blocked_spills.remove(&anchor); } }
                return match answer { Arg::Range(block) => formula_result(block.cells.first().cloned().unwrap_or(Value::Blank)), Arg::Value(value) => formula_result(value) };
            }
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
                    held.blocked_spills.insert(anchor, (width, height));
                    return Value::Error(ExcelError::Spill);
                }
            }
        }
        held.blocked_spills.remove(&anchor);
        for down in 0..height {
            for across in 0..width {
                let at = (col + across, row + down);
                if at == anchor {
                    continue;
                }
                let value = if provisional { Value::Blank } else { formula_result(block.cells[(down * width + across) as usize].clone()) };
                if let Some(history) = &history {
                    let before = held.spilled.get(&at).map(|(value, _)| value).unwrap_or(&Value::Blank);
                    history.borrow_mut().change(sheet, at, before, &value);
                }
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
        if !self.calculation_order.is_empty() {
            keys.sort_by_key(|cell| self.calculation_order.get(cell).copied().unwrap_or(usize::MAX));
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
            let obstruction_changed = self.sheets.get(sheet).and_then(|held| held.blocked_spills.get(coord)).is_some_and(|(width, height)| {
                changed.iter().any(|(target, (col, row))| target == sheet && *col >= coord.0 && *row >= coord.1 && *col - coord.0 < *width && *row - coord.1 < *height)
            });
            let reads_a_change = obstruction_changed || self.dependency_expression(&expr, sheet, Some(*coord)).invalidation_references().iter().any(|reference| {
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
        contents_only: bool,
        found: &mut dyn FnMut(usize),
    ) {
        self.each_dependency_with_spills(sheet, coord, at, contents_only, true, found);
    }

    fn each_dependency_with_spills(
        &self,
        sheet: &str,
        coord: &(u32, u32),
        at: &BTreeMap<&str, BTreeMap<(u32, u32), usize>>,
        contents_only: bool,
        include_spills: bool,
        found: &mut dyn FnMut(usize),
    ) {
        let Some(expr) = self.expr_at(sheet, coord) else {
            return;
        };
        let expr = self.dependency_expression(&expr, sheet, Some(*coord));
        let references = if contents_only { self.ordering_expression(&expr, sheet, Some(*coord)).value_references() } else { expr.invalidation_references() };
        for reference in &references {
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
        if !include_spills { return; }
        // A cell a spill fills is waiting on the formula that spills it.
        for reference in &references {
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

    fn eval_reference_argument(&self, expr: &Expr, sheet: &str, depth: u32, at: At) -> Arg {
        let previous = self.reference_context.replace(true);
        let answer = self.eval_arg(expr, sheet, depth, at);
        self.reference_context.set(previous);
        answer
    }

    fn metadata_ranges(&self, expr: &Expr, sheet: &str, depth: u32, at: At) -> Option<Vec<(String, RangeRef)>> {
        let ranges = match expr {
            Expr::Function { name, args } if name == "_UNION" => args.iter().map(|area| self.reference_of(area, sheet, depth + 1, at)).collect::<Option<Vec<_>>>()?,
            _ => vec![self.reference_of(expr, sheet, depth + 1, at)?],
        };
        let first = ranges.first()?;
        ranges.iter().all(|(target, _)| target.eq_ignore_ascii_case(&first.0)).then_some(ranges)
    }

    fn dependency_value_is_fixed(&self, expr: &Expr, sheet: &str) -> bool {
        let mut fixed = true;
        expr.visit(&mut |node| match node {
            Expr::Ref(reference) => {
                let target = reference.sheet.as_deref().unwrap_or(sheet);
                if reference.book.is_some() { fixed = false; return; }
                if let Some(held) = self.sheets.get(target) {
                    let (from, to) = (reference.range.start, reference.range.end);
                    if held.cells.range((from.col, 0)..=(to.col, u32::MAX)).any(|((_, row), cell)| *row >= from.row && *row <= to.row && matches!(cell, Cell::Formula { .. })) { fixed = false; }
                }
            }
            Expr::Function { name, .. } if !functions::is_known_function(name)
                || matches!(name.as_str(), "RAND" | "RANDBETWEEN" | "RANDARRAY" | "NOW" | "TODAY" | "OFFSET" | "INDIRECT" | "CELL" | "INFO" | "SHEET" | "SHEETS" | "FORMULATEXT") => fixed = false,
            Expr::Name(_) | Expr::Table { .. } => fixed = false,
            _ => {},
        });
        fixed
    }

    // Reduce only lookup inputs whose values cannot change during this pass.
    // Invalidation still uses the original expression and all its dependencies.
    fn ordering_expression(&self, expr: &Expr, sheet: &str, at: At) -> Expr {
        match expr {
            Expr::Unary { op, operand } => Expr::Unary { op: *op, operand: Box::new(self.ordering_expression(operand, sheet, at)) },
            Expr::Binary { op, lhs, rhs } => Expr::Binary { op: *op, lhs: Box::new(self.ordering_expression(lhs, sheet, at)), rhs: Box::new(self.ordering_expression(rhs, sheet, at)) },
            Expr::Function { name, args } => {
                let mut args: Vec<_> = args.iter().map(|arg| self.ordering_expression(arg, sheet, at)).collect();
                if name == "INDEX" && !self.array_formula_context(sheet, at)
                    && !args.iter().skip(1).any(|arg| Self::materialized_index_argument(arg)
                        && (!self.dependency_value_is_fixed(arg, sheet)
                            || matches!(self.eval_arg(arg, sheet, 0, at), Arg::Range(block) if block.is_array()))) {
                    // Legacy formulas consume only the first literal selector.
                    // Unused selector elements must not introduce graph cycles.
                    for arg in args.iter_mut().skip(1) {
                        if let Expr::Array(rows) = arg {
                            if let Some(first) = rows.first().and_then(|row| row.first()) {
                                *arg = Expr::Literal(first.clone());
                            }
                        }
                    }
                }
                if name == "IF" && (2..=3).contains(&args.len()) && self.dependency_value_is_fixed(&args[0], sheet) {
                    let condition = self.eval_arg(&args[0], sheet, 0, at);
                    if !matches!(&condition, Arg::Range(block) if block.is_array() || block.cells.len() != 1) {
                        if let Ok(taken) = condition.scalar().to_logical() {
                            let branch = args.get(if taken { 1 } else { 2 }).cloned().unwrap_or(Expr::Literal(Value::Logical(false)));
                            // Only the selected value orders this pass. Changes
                            // to either branch still invalidate the original IF.
                            return Expr::Function { name: "_UNION".into(), args: vec![args[0].clone(), branch] };
                        }
                    }
                }
                if name == "FORMULATEXT" && args.len() == 1 {
                    let first = match &args[0] {
                        Expr::Ref(reference) => Some(reference),
                        Expr::Function { name: union, args } if union == "_UNION" => match args.first() { Some(Expr::Ref(reference)) => Some(reference), _ => None },
                        _ => None,
                    };
                    if let Some(reference) = first {
                        let mut reference = reference.clone();reference.range = RangeRef::single(reference.range.start);
                        let target = reference.sheet.as_deref().unwrap_or(sheet);
                        if target == sheet && at == Some(reference.range.start.coord()) {
                            if let Some(Cell::Formula { source, source_utf16, .. }) = self.sheets.get(target).and_then(|held| held.cells.get(&reference.range.start.coord())) {
                                return Expr::Literal(formula_source_value(source, source_utf16.as_deref()));
                            }
                        }
                        return Expr::Function { name: name.clone(), args: vec![Expr::Ref(reference)] };
                    }
                }
                if name == "XLOOKUP" && (3..=6).contains(&args.len()) && [0, 1, 3, 4, 5].iter().all(|i| args.get(*i).is_none_or(|arg| self.dependency_value_is_fixed(arg, sheet))) {
                    if let Expr::Ref(returned) = &args[2] {
                        let lookup = self.eval_arg(&args[1], sheet, 0, at);
                        let asked = [self.eval_arg(&args[0], sheet, 0, at), lookup.clone(), args.get(4).map(|arg| self.eval_arg(arg, sheet, 0, at)).unwrap_or(Arg::Value(Value::Number(0.0))), args.get(5).map(|arg| self.eval_arg(arg, sheet, 0, at)).unwrap_or(Arg::Value(Value::Number(1.0)))];
                        match functions::call("XMATCH", &asked) {
                            Value::Number(position) if position >= 1.0 => {
                                let offset = position as u32 - 1;let mut selected = returned.clone();
                                let lookup = lookup.as_range();
                                match functions::xlookup_axis((lookup.width, lookup.height), (selected.range.width() as usize, selected.range.height() as usize)) {
                                    Some(true) if offset < selected.range.height() => { selected.range.start.row += offset;selected.range.end.row = selected.range.start.row;return Expr::Ref(selected); }
                                    Some(false) if offset < selected.range.width() => { selected.range.start.col += offset;selected.range.end.col = selected.range.start.col;return Expr::Ref(selected); }
                                    _ => {}
                                }
                            }
                            Value::Error(ExcelError::NA) => return args.get(3).cloned().unwrap_or(Expr::Literal(Value::Error(ExcelError::NA))),
                            Value::Error(why) => return Expr::Literal(Value::Error(why)),
                            _ => {},
                        }
                    }
                }
                Expr::Function { name: name.clone(), args }
            }
            _ => expr.clone(),
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

    /// A cell caught in a cycle keeps what it last showed; one never worked
    /// out shows 0. Measured: `=B1+1` in A1 reads 1, and then `=A1+1` in B1
    /// closes the loop and reads 0 while A1 stays 1 -- as does a formula
    /// entered downstream of the loop (`=A1*2+D1` reads 0).
    fn keep_or_zero(&mut self, sheet: &str, coord: &(u32, u32)) {
        let held = self
            .sheets
            .get(sheet)
            .and_then(|s| s.cells.get(coord))
            .map(|cell| cell.value().clone())
            .unwrap_or(Value::Blank);
        if matches!(held, Value::Blank) {
            self.store_cached(sheet, coord, Value::Number(0.0));
        }
    }

    fn store_cached(&mut self, sheet: &str, coord: &(u32, u32), value: Value) {
        let history = self.read_history.borrow().clone();
        if let Some(Cell::Formula { cached, .. }) = self
            .sheets
            .get_mut(sheet)
            .and_then(|s| s.cells.get_mut(coord))
        {
            if let Some(history) = history { history.borrow_mut().change(sheet, *coord, cached, &value); }
            *cached = value;
        }
    }

    // -- evaluation -------------------------------------------------------

    fn callback_read_view(&self) -> Option<UserFunctionReadView> {
        if !USER_FUNCTIONS.with(|held| held.try_borrow().is_ok_and(|hook| hook.is_some())) { return None; }
        let mut slot = self.read_history.borrow_mut();
        let history = slot.get_or_insert_with(|| std::rc::Rc::new(std::cell::RefCell::new(UserReadHistory::default())));
        let epoch = history.borrow().epoch;
        Some(UserFunctionReadView { epoch, history: history.clone() })
    }

    fn value_at(&self, sheet: &str, col: u32, row: u32) -> Value {
        if let Some((watched_sheet, coord)) = self.watched.borrow().as_ref() {
            if *coord == (col, row) && watched_sheet.eq_ignore_ascii_case(sheet) {
                self.watched_read.set(true);
            }
        }
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
        if !self.require_ready_value(&worked) { return Value::Error(ExcelError::Value); }
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

    fn reference_arg(&self, sheet: &str, range: &RangeRef, skip: Skip) -> Arg {
        let mut data = self.materialise(sheet, range, skip);
        data.origin = crate::functions::BlockOrigin::Cells { sheet: sheet.to_string(), range: *range };
        Arg::Range(data)
    }

    fn array_formula_context(&self, sheet: &str, at: At) -> bool {
        at.is_some_and(|coord| self.share_at(sheet, &coord).is_some()
            || self.sheets.get(sheet).is_some_and(|held| held.dynamic.contains(&coord)))
    }

    /// Resolve stored references without reading their cell contents or
    /// evaluating a source expression twice when it is not a reference.
    fn stored_reference(&self, expr: &Expr, sheet: &str, depth: u32, at: At) -> Option<(String, RangeRef)> {
        if depth > 100 { return None; }
        match expr {
            Expr::Ref(reference) if reference.book.is_none() => {
                let target = reference.sheet.as_deref().unwrap_or(sheet);
                self.sheets.contains_key(target).then(|| (target.to_string(), reference.range))
            }
            Expr::Table { name, asked } => self.table_range(name, asked, sheet, at).ok(),
            Expr::Name(name) => {
                if let Some((_, value)) = self.let_scope.borrow().iter().rev().find(|(label, _)| label == name) {
                    return value.reference();
                }
                self.name_bound(name, sheet, at).and_then(|bound| self.stored_reference(&bound, sheet, depth + 1, at))
            }
            Expr::Function { name, args } if name == "CHOOSE" => {
                let Expr::Literal(index) = args.first()? else { return None; };
                let index = index.to_number().ok()?;
                if !index.is_finite() || index < 1.0 || index >= args.len() as f64 { return None; }
                self.stored_reference(&args[index as usize], sheet, depth + 1, at)
            }
            _ => None,
        }
    }

    /// INDEX reads its selection, not every cell in the source rectangle.
    /// In particular, unselected cells can include the formula itself.
    fn index_reference_arg(&self, sheet: &str, range: &RangeRef, indexes: &[Arg], skip: Skip) -> Arg {
        let error = |why| Arg::Value(Value::Error(why));
        if !(1..=2).contains(&indexes.len()) { return error(ExcelError::Value); }
        if indexes.iter().any(|arg| matches!(arg, Arg::Range(block) if block.is_array() || block.cells.len() > 1)) {
            let blocks: Vec<_> = indexes.iter().map(Arg::as_range).collect();
            let width = blocks.iter().map(|block| block.width).max().unwrap_or(1);
            let height = blocks.iter().map(|block| block.height).max().unwrap_or(1);
            let Some(size) = width.checked_mul(height) else { return error(ExcelError::Num); };
            let mut cells = Vec::with_capacity(size);
            for row in 0..height {
                for col in 0..width {
                    let selected: Vec<_> = blocks.iter().map(|block| {
                        let x = if block.width == 1 { 0 } else { col };
                        let y = if block.height == 1 { 0 } else { row };
                        Arg::Value(if x < block.width && y < block.height { block.at(x, y) } else { Value::Error(ExcelError::NA) })
                    }).collect();
                    cells.push(self.index_reference_arg(sheet, range, &selected, skip).first());
                }
            }
            return Arg::Range(RangeData { origin: crate::functions::BlockOrigin::Array, width, height, cells });
        }
        let number = |arg: &Arg| -> Result<u32, ExcelError> {
            let n = arg.scalar().to_number()?;
            if n < 0.0 { return Err(ExcelError::Value); }
            if !n.is_finite() || n > u32::MAX as f64 { return Err(ExcelError::Ref); }
            Ok(n.trunc() as u32)
        };
        let first = match number(&indexes[0]) { Ok(n) => n, Err(why) => return error(why) };
        let (row, col) = if let Some(index) = indexes.get(1) {
            match number(index) { Ok(col) => (first, col), Err(why) => return error(why) }
        } else if range.height() == 1 { (1, first) }
        else if range.width() == 1 { (first, 1) }
        else { return error(ExcelError::Ref); };
        if row > range.height() || col > range.width() { return error(ExcelError::Ref); }
        let mut selected = *range;
        if row != 0 { selected.start.row += row - 1; selected.end.row = selected.start.row; }
        if col != 0 { selected.start.col += col - 1; selected.end.col = selected.start.col; }
        self.reference_arg(sheet, &selected, skip)
    }

    fn scalar_cell_reference<'a>(name: &str, index: usize, argument: &'a Expr) -> Option<&'a crate::reference::Reference> {
        if functions::is_known_function(name) || user_reference_argument(name, index)
            || user_variant_argument(name, index) { return None; }
        match argument {
            Expr::Ref(reference) if reference.book.is_none() && reference.range.is_single() => Some(reference),
            _ => None,
        }
    }

    fn spill_dependency_was_published(&self, consumer: &FormulaCell, producer: &FormulaCell) -> bool {
        let original = USER_FUNCTION_SPILL_PROBE_LAYOUT.with(|held| held.borrow().as_ref().map(|layout|
            layout.iter().find(|(owner, _)| owner == producer).map(|(_, size)| *size)));
        let Some(original) = original else { return true; };
        let Some((width, height)) = original else { return false; };
        let Some(expr) = self.expr_at(&consumer.0, &consumer.1) else { return false; };
        let expr = self.dependency_expression(&expr, &consumer.0, Some(consumer.1));
        self.ordering_expression(&expr, &consumer.0, Some(consumer.1)).value_references().iter().any(|reference|
            crate::case_fold::name_key(reference.sheet.as_deref().unwrap_or(&consumer.0)) == crate::case_fold::name_key(&producer.0)
                && reference.range.start.col < producer.1.0 + width && reference.range.end.col >= producer.1.0
                && reference.range.start.row < producer.1.1 + height && reference.range.end.row >= producer.1.1)
    }

    fn reference_waits_for_formula(&self, source: &str, range: &RangeRef, target: &str, coord: (u32, u32)) -> bool {
        if crate::case_fold::name_key(source) != crate::case_fold::name_key(target) { return false; }
        if range.contains(coord.0, coord.1) { return true; }
        // A staged deterministic array intentionally exposes its new footprint
        // to geometry probes. Other spill members wait for their owner.
        if self.prepared_spills.contains_key(&(target.to_string(), coord)) { return false; }
        self.sheets.get(target).and_then(|sheet| sheet.spills.get(&coord)).is_some_and(|(width, height)|
            range.start.col < coord.0 + width && range.end.col >= coord.0
                && range.start.row < coord.1 + height && range.end.row >= coord.1)
    }

    fn unready_argument(&self, argument: &Arg) -> bool {
        let Arg::Range(block) = argument else { return false; };
        let Some((sheet, range)) = block.cell_reference() else { return false; };
        self.unready_reference(&sheet, &range)
    }

    fn unready_reference(&self, sheet: &str, range: &RangeRef) -> bool {
        if !self.provisional_value.get() { return false; }
        let cells = self.unready_cells.borrow();
        cells.iter().any(|(target, coord)| self.reference_waits_for_formula(sheet, range, target, *coord))
    }

    fn require_ready_reference(&self, sheet: &str, range: &RangeRef) -> bool {
        if self.unready_reference(sheet, range) { self.value_deferred.set(true); }
        !self.value_deferred.get()
    }

    fn require_ready_value(&self, argument: &Arg) -> bool {
        if self.unready_argument(argument) { self.value_deferred.set(true); }
        !self.value_deferred.get()
    }

    fn eval_function_argument(&self, (name, index): (&str, usize), expr: &Expr, sheet: &str, depth: u32, skip: Skip, at: At) -> Arg {
        let user_function_argument = !functions::is_known_function(name);
        let previous_argument_context = if user_function_argument {
            USER_FUNCTION_ARGUMENT_CONTEXT.with(|held| held.replace(Some((name.to_ascii_uppercase(), index))))
        } else { None };
        if user_function_argument && user_reference_argument(name, index) {
            let previous = self.reference_context.replace(true);
            let previous_value_read = self.variant_value_read_context.replace(self.variant_value_read_context.get()
                || (user_variant_argument(name, index) && user_value_read_argument(name, index)));
            let value = self.eval_arg_inner(expr, sheet, depth, skip, at);
            self.variant_value_read_context.set(previous_value_read);
            self.reference_context.set(previous);
            USER_FUNCTION_ARGUMENT_CONTEXT.with(|held| { held.replace(previous_argument_context); });
            return value;
        }
        let value = if name == "SUMPRODUCT" {
            self.eval_materialized_arg(expr, sheet, depth, skip, at)
        } else { self.eval_arg_inner(expr, sheet, depth, skip, at) };
        if user_function_argument {
            USER_FUNCTION_ARGUMENT_CONTEXT.with(|held| { held.replace(previous_argument_context); });
        }
        if (name == "SWITCH" && index == 0) || (name == "IFS" && index % 2 == 0) {
            self.require_ready_value(&value);
        }
        value
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
        if user_function_pending() || self.value_deferred.get() { return Arg::Value(Value::Error(ExcelError::Value)); }
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
                Arg::Range(RangeData {origin: crate::functions::BlockOrigin::Array,
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
                self.reference_arg(target, &reference.range, skip)
            }

            Expr::Table { name, asked } => self.a_table_column(name, asked, sheet, skip, at),

            Expr::Name(name) => {
                let bound_by_let = self
                    .let_scope
                    .borrow()
                    .iter()
                    .rev()
                    .find(|(held, _)| held == name)
                    .map(|(_, value)| value.clone());
                if let Some(value) = bound_by_let {
                    return match value {
                        LocalBinding::Value(value) => value,
                        LocalBinding::Reference { sheet, range } => self.reference_arg(&sheet, &range, skip),
                    };
                }
                match self.name_bound(name, sheet, at) {
                    Some(bound) => self.eval_arg_inner(&bound, sheet, depth + 1, skip, at),
                    // A table's bare name is its data: `COLUMNS(tbl)`.
                    None if self.tables.contains_key(&crate::case_fold::table_key(name)) => {
                        self.a_table_column(name, "", sheet, skip, at)
                    }
                    None => Arg::Value(Value::Error(ExcelError::Name)),
                }
            }

            Expr::Unary { op, operand } => {
                let operand = self.eval_arg(operand, sheet, depth + 1, at);
                if !self.require_ready_value(&operand) { return Arg::Value(Value::Error(ExcelError::Value)); }
                match operand {
                    Arg::Value(v) => Arg::Value(apply_unary(*op, v)),
                    Arg::Range(block) if block.cells.len() == 1 && !block.is_array() => Arg::Value(apply_unary(*op, block.cells[0].clone())),
                    Arg::Range(block) => Arg::Range(RangeData {origin: RangeData::value_origin([&block]),
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
                let a = self.eval_materialized_arg(lhs, sheet, depth + 1, Skip::Nothing, at);
                if !self.require_ready_value(&a) { return Arg::Value(Value::Error(ExcelError::Value)); }
                let b = self.eval_materialized_arg(rhs, sheet, depth + 1, Skip::Nothing, at);
                if !self.require_ready_value(&b) { return Arg::Value(Value::Error(ExcelError::Value)); }
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
                // The value is read only for the kinds that show it, so asking where a
                // cell is does not read it: measured, `=CELL("address",L1)` in L1
                // is $L$1, not a circular 0.
                let held = || self.value_at(&target, cell.col, cell.row);
                if info == "width" {
                    let held = self.sheets.get(&target);
                    let stated = held.and_then(|sheet| sheet.widths.0.get(cell.col as usize)).copied().filter(|width| *width > 0.0);
                    let default = held.map_or(0.0, |sheet| sheet.widths.1);
                    let file = stated.or(Some(default).filter(|width| *width > 0.0 || held.is_some_and(|sheet| sheet.default_width_explicit)));
                    let digit = DIGIT_WIDTH.with(|held| held.get());
                    let characters = if held.is_some_and(|sheet| sheet.hidden_columns.contains(&cell.col)) { 0.0 }
                        else { file.map_or(8.43, |file| crate::column_width::pixels_to_characters(crate::column_width::stored_width_to_pixels(f64::from(file), digit), digit)) };
                    let custom = held.and_then(|sheet| sheet.custom_widths.get(&cell.col)).copied().unwrap_or(stated.is_some());
                    return Arg::Range(RangeData {
                        origin: crate::functions::BlockOrigin::Vector,
                        width: 2, height: 1,
                        cells: vec![Value::Number(characters.max(0.0).round()), Value::Logical(!custom)],
                    });
                }
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
                    // An unsaved book has no file name: measured, "".
                    "filename" => Value::Text(String::new()),
                    "protect" => Value::Number(1.0),
                    "format" | "parentheses" | "color" => {
                        let format = self.dress_of(&target, cell.col, cell.row).0.unwrap_or_default();
                        let (code, parentheses, color) = cell_format_code(&format);
                        match info.as_str() {
                            "format" => Value::Text(code),
                            "parentheses" => Value::Number(f64::from(u8::from(parentheses))),
                            _ => Value::Number(f64::from(u8::from(color))),
                        }
                    }
                    // Text says how it is aligned: measured, ' for left, "
                    // for right and ^ for centred; anything else says nothing.
                    "prefix" => Value::Text(match held() {
                        Value::Text(_) => match self.dress_of(&target, cell.col, cell.row).1.as_deref() {
                            Some("right") => "\"",
                            Some("center") => "^",
                            Some("fill") => "\\",
                            _ => "'",
                        }
                        .to_string(),
                        _ => String::new(),
                    }),
                    "contents" => held(),
                    "type" => Value::Text(
                        match held() {
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
                // A UDF may return either a reference or a scalar. Retain its
                // result rather than calling it once to test for a reference
                // and again to obtain its value.
                let worked = self.eval_reference_argument(&args[0], sheet, depth + 1, at);
                let reference = match &worked { Arg::Range(block) => block.cell_reference(), _ => None };
                if let Some((target, range)) = reference {
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
                        Some(cell) => self.reference_arg(&target, &RangeRef::single(cell), skip),
                        None => Arg::Value(Value::Error(ExcelError::Value)),
                    };
                }
                match worked {
                    Arg::Range(block) => Arg::Value(block.cells.into_iter().next().unwrap_or(Value::Blank)),
                    value => value,
                }
            }

            Expr::Function { name, args } if name == "TRIMRANGE" => match self.trimmed_range(args, sheet, depth + 1, at) {
                Some((target, range)) => self.reference_arg(&target, &range, skip),
                None => Arg::Value(Value::Error(ExcelError::Ref)),
            },
            // INFO: facts about the environment, as Excel 16 on Windows gives
            // them. Measured: numfile 1 (the sheets), release "16.0", system
            // "pcdos", origin "$A:$A$1", osversion "Windows (64-bit) NT
            // 10.00"; an unknown type is #VALUE!.
            Expr::Function { name, args } if name == "INFO" && args.len() == 1 => {
                let asked = match self.eval_arg_inner(&args[0], sheet, depth + 1, skip, at).scalar() {
                    Value::Text(text) => text.to_ascii_lowercase(),
                    Value::Error(why) => return Arg::Value(Value::Error(why)),
                    _ => return Arg::Value(Value::Error(ExcelError::Value)),
                };
                Arg::Value(match asked.as_str() {
                    "numfile" => Value::Number(self.sheet_order.len().max(1) as f64),
                    "release" => Value::Text("16.0".to_string()),
                    "system" => Value::Text("pcdos".to_string()),
                    "origin" => Value::Text("$A:$A$1".to_string()),
                    "osversion" => Value::Text("Windows (64-bit) NT 10.00".to_string()),
                    "directory" => Value::Text(String::new()),
                    // Measured: recalc 自動 in Japanese Excel (Automatic);
                    // the memory types went away with Excel 2007: #N/A.
                    "recalc" => Value::Text("Automatic".to_string()),
                    "memavail" | "memused" | "totmem" => Value::Error(ExcelError::NA),
                    _ => Value::Error(ExcelError::Value),
                })
            }
            Expr::Function { name, args } if name == "CHOOSE" && args.len() >= 2 => {
                let index = self.eval_arg_inner(&args[0], sheet, depth + 1, skip, at);
                if !self.require_ready_value(&index) { return Arg::Value(Value::Error(ExcelError::Value)); }
                if user_function_pending() || self.value_deferred.get() { return Arg::Value(Value::Error(ExcelError::Value)); }
                if matches!(&index, Arg::Range(block) if block.is_array() || block.cells.len() > 1) {
                    let mut evaluated = vec![index];
                    evaluated.extend(args.iter().skip(1).map(|arg| self.eval_arg_inner(arg, sheet, depth + 1, skip, at)));
                    functions::call_arg_with_text_context(name, &evaluated, self.compatibility_version, self.dbcs_language.unwrap_or(true))
                } else {
                    match index.scalar().to_number() {
                        Ok(n) if n >= 1.0 && (n.trunc() as usize) < args.len() => self.eval_arg_inner(&args[n.trunc() as usize], sheet, depth + 1, skip, at),
                        Ok(_) => Arg::Value(Value::Error(ExcelError::Value)),
                        Err(why) => Arg::Value(Value::Error(why)),
                    }
                }
            }
            Expr::Function { name, args } if matches!(name.as_str(), "IFERROR" | "IFNA") && args.len() == 2 => {
                let first = self.eval_arg_inner(&args[0], sheet, depth + 1, skip, at);
                if !self.require_ready_value(&first) { return Arg::Value(Value::Error(ExcelError::Value)); }
                if user_function_pending() || self.value_deferred.get() { return Arg::Value(Value::Error(ExcelError::Value)); }
                let needs_fallback = first.flatten().iter().any(|value| matches!(value, Value::Error(why) if name == "IFERROR" || *why == ExcelError::NA));
                let fallback = if needs_fallback { self.eval_arg_inner(&args[1], sheet, depth + 1, skip, at) } else { Arg::Value(Value::Blank) };
                functions::call_arg_with_text_context(name, &[first, fallback], self.compatibility_version, self.dbcs_language.unwrap_or(true))
            }

            // Resolve selections from their already evaluated arguments.
            // A failed reference selection must not retry a volatile selector
            // when falling back to the value-function path.
            Expr::Function { name, args }
                if matches!(name.as_str(), "INDEX" | "TAKE" | "DROP" | "XLOOKUP")
                    && !(name == "INDEX" && matches!(args.first(), Some(Expr::Function { name, .. }) if name == "_UNION")) =>
            {
                if name == "INDEX" && (2..=3).contains(&args.len()) {
                    let source = args.first().and_then(|source| match source {
                        Expr::Function { name, args } if name == "OFFSET" => Some(self.offset_range(args, sheet, depth + 1, at)),
                        Expr::Function { name, args } if name == "INDIRECT" => Some(self.indirect_range(args, sheet, depth + 1, at)),
                        _ => self.stored_reference(source, sheet, depth + 1, at).map(Ok),
                    });
                    if let Some(source) = source {
                        let mut indexes: Vec<_> = args.iter().skip(1).map(|index| self.eval_arg_inner(index, sheet, depth + 1, skip, at)).collect();
                        if indexes.iter().any(|index| !self.require_ready_value(index)) { return Arg::Value(Value::Error(ExcelError::Value)); }
                        let materialized = self.prepare_array_indexes(&mut indexes, &args[1..]);
                        if !self.array_formula_context(sheet, at) && !materialized {
                            for index in &mut indexes { *index = Arg::Value(index.first()); }
                        }
                        return match source {
                            Ok((target, range)) => self.index_reference_arg(&target, &range, &indexes, skip),
                            Err(why) => {
                                let mut evaluated = vec![Arg::Value(Value::Error(why))];
                                evaluated.extend(indexes);
                                functions::call_arg_with_text_context(name, &evaluated, self.compatibility_version, self.dbcs_language.unwrap_or(true))
                            }
                        };
                    }
                }
                let mut evaluated: Vec<Arg> = args.iter().map(|arg| self.eval_arg_inner(arg, sheet, depth + 1, skip, at)).collect();
                let selectors: Vec<_> = if name == "XLOOKUP" { [0, 1, 4, 5].iter().filter_map(|i| evaluated.get(*i)).collect() } else { evaluated.iter().skip(1).collect() };
                if selectors.iter().any(|arg| !self.require_ready_value(arg)) { return Arg::Value(Value::Error(ExcelError::Value)); }
                if name == "INDEX" {
                    let materialized = self.prepare_array_indexes(&mut evaluated[1..], &args[1..]);
                    // Excel Evaluate uses the first index in an index block;
                    // it does not produce one selection per index element.
                    let array_context = self.array_formula_context(sheet, at);
                    if !array_context && !materialized {
                        for arg in evaluated.iter_mut().skip(1) {
                            *arg = Arg::Value(match &*arg { Arg::Range(block) => block.cells.first().cloned().unwrap_or(Value::Blank), other => other.scalar() });
                        }
                    }
                    let has_array_index = evaluated.iter().skip(1).any(|arg| matches!(arg, Arg::Range(block) if block.is_array() || block.cells.len() > 1));
                    if !has_array_index && evaluated.iter().skip(1).any(|arg| arg.scalar().to_number().is_ok_and(|n| n < 0.0)) {
                        return Arg::Value(Value::Error(ExcelError::Value));
                    }
                }
                let array_index = name == "INDEX" && evaluated.iter().skip(1).any(|arg| matches!(arg, Arg::Range(block) if block.is_array() || block.cells.len() > 1));
                if let Some((target, range)) = (!array_index).then(|| self.selection_reference(name, &evaluated)).flatten() {
                    self.reference_arg(&target, &range, skip)
                } else if name == "INDEX" && !array_index && evaluated.len() == 2
                    && matches!(evaluated.first(), Some(Arg::Range(block)) if block.cell_reference().is_some_and(|(_, range)| range.width() > 1 && range.height() > 1)) {
                    Arg::Value(Value::Error(ExcelError::Ref))
                } else {
                    if evaluated.iter().any(|argument| !self.require_ready_value(argument)) { return Arg::Value(Value::Error(ExcelError::Value)); }
                    functions::call_arg_with_text_context(name, &evaluated, self.compatibility_version, self.dbcs_language.unwrap_or(true))
                }
            }

            // Two references that share no cell: #NULL!, measured.
            Expr::Function { name, .. } if name == "_ISECT" => match self.reference_of(expr, sheet, depth + 1, at) {
                Some((target, range)) => self.reference_arg(&target, &range, skip),
                None => Arg::Value(Value::Error(ExcelError::Null)),
            },

            // `D1#`: all that D1 spilled, or #REF! where it spills nothing.
            Expr::Function { name, .. } if name == "_SPILL" => match self.reference_of(expr, sheet, depth + 1, at) {
                Some((target, range)) => self.reference_arg(&target, &range, skip),
                None => Arg::Value(Value::Error(ExcelError::Ref)),
            },

            Expr::Function { name, args } if name == "ISREF" => match args.as_slice() {
                [only] => Arg::Value(Value::Logical(self.metadata_ranges(only, sheet, depth + 1, at).is_some())),
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
            // difficulty at all. Computed references retain their full shape;
            // value arrays and scalars use the function library.
            Expr::Function { name, args }
                if matches!(name.as_str(), "ROWS" | "COLUMNS") =>
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
                // No bigger than SEQUENCE will make.
                if rows < 1.0 || cols < 1.0 || low > high || rows.trunc() * cols.trunc() > 1_048_576.0 {
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
                Arg::Range(RangeData {origin: crate::functions::BlockOrigin::Array,  width, height, cells })
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
            // A multi-area reference is still one PHONETIC argument. Its
            // areas must not be flattened as separate scalar arguments.
            Expr::Function { name, args } if name == "PHONETIC" && args.len() == 1
                && matches!(&args[0], Expr::Function { name, .. } if name == "_UNION") =>
            {
                Arg::Value(Value::Error(if self.metadata_ranges(&args[0], sheet, depth + 1, at).is_some() {
                    ExcelError::NA
                } else {
                    ExcelError::Value
                }))
            }
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
                    ("SHEET", [only @ Expr::Function { name: union, .. }]) if union == "_UNION" => {
                        match self.metadata_ranges(only, sheet, depth + 1, at) {
                            Some(ranges) => place(&ranges[0].0),
                            None => Arg::Value(Value::Error(ExcelError::Value)),
                        }
                    }
                    ("SHEETS", [only @ Expr::Function { name: union, .. }]) if union == "_UNION" => {
                        match self.metadata_ranges(only, sheet, depth + 1, at) {
                            Some(_) => Arg::Value(Value::Number(1.0)),
                            None => Arg::Value(Value::Error(ExcelError::Value)),
                        }
                    }
                    ("SHEETS", []) => Arg::Value(Value::Number(self.sheet_order.len().max(1) as f64)),
                    ("SHEETS", [Expr::Ref(_)]) | ("AREAS", [Expr::Ref(_)]) => Arg::Value(Value::Number(1.0)),
                    ("SHEET", []) => place(sheet),
                    ("SHEET", [Expr::Ref(reference)]) => place(reference.sheet.as_deref().unwrap_or(sheet)),
                    ("SHEETS", [other]) | ("AREAS", [other]) | ("SHEET", [other]) => {
                        let answer = self.eval_reference_argument(other, sheet, depth + 1, at);
                        if let Arg::Range(block) = &answer {
                            if let Some((target, _)) = block.cell_reference() { return if name == "SHEET" { place(&target) } else { Arg::Value(Value::Number(1.0)) }; }
                        }
                        if name != "SHEET" { return Arg::Value(Value::Error(if name == "SHEETS" { ExcelError::NA } else { ExcelError::Value })); }
                        match answer.scalar() {
                            Value::Text(named) => place(&named),
                            Value::Error(why) => Arg::Value(Value::Error(why)),
                            _ => Arg::Value(Value::Error(ExcelError::NA)),
                        }
                    }
                    ("ISFORMULA", [only]) => {
                        let Some(ranges) = self.metadata_ranges(only, sheet, depth + 1, at) else { return Arg::Value(Value::Error(ExcelError::Value)); };
                        let (target, reference) = &ranges[0];
                        let target = target.as_str();
                        let corner = (reference.start.col, reference.start.row);
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
            // FORMULATEXT: the formula a cell holds, as written; #N/A for a
            // cell holding none. Measured: over `=A1+1` it is "=A1+1".
            Expr::Function { name, args } if name == "FORMULATEXT" && args.len() == 1 => {
                let Some(ranges) = self.metadata_ranges(&args[0], sheet, depth + 1, at) else {
                    return Arg::Value(Value::Error(ExcelError::Value));
                };
                let (target, range) = &ranges[0];
                let corner = (range.start.col, range.start.row);
                if self.provisional_value.get() && !(target == sheet && at == Some(corner)) {
                    let reference = self.reference_arg(target, &RangeRef::single(range.start), skip);
                    if !self.require_ready_value(&reference) { return Arg::Value(Value::Error(ExcelError::Value)); }
                }
                match self.sheets.get(target).and_then(|held| held.cells.get(&corner)) {
                    Some(Cell::Formula { source, source_utf16, .. }) => Arg::Value(formula_source_value(source, source_utf16.as_deref())),
                    _ => Arg::Value(Value::Error(ExcelError::NA)),
                }
            }
            // The helpers that hand values to a LAMBDA written in place:
            // MAP, REDUCE, SCAN, BYROW, BYCOL and MAKEARRAY. Measured:
            // `MAP(A1:A3,LAMBDA(v,v*10))` over 3,1,2 is 30,10,20,
            // `REDUCE(0,A1:A4,LAMBDA(a,v,a+v))` 7, `SCAN` 3,4,6,7,
            // `BYROW(A1:C2,LAMBDA(r,SUM(r)))` 4,3 (text is not summed) and
            // `MAKEARRAY(2,2,LAMBDA(r,c,r*c))` 1,2,2,4.
            // The ranking functions rank within a reference and nothing else:
            // measured, RANK.AVG(5,{1,5,5,9}) is #VALUE!.
            Expr::Function { name, args }
                if matches!(name.as_str(), "RANK" | "RANK.EQ" | "RANK.AVG") && matches!(args.get(1), Some(Expr::Array(_))) =>
            {
                Arg::Value(Value::Error(ExcelError::Value))
            }
            // AGGREGATE's first thirteen take references only: measured,
            // AGGREGATE(1,6,{1,2,#N/A}) is #VALUE!, where 14 to 19 take an
            // array.
            Expr::Function { name, args }
                if name == "AGGREGATE"
                    && matches!(args.first(), Some(Expr::Literal(Value::Number(n))) if (1.0..14.0).contains(n))
                    && args.iter().skip(2).any(|arg| matches!(arg, Expr::Array(_))) =>
            {
                Arg::Value(Value::Error(ExcelError::Value))
            }
            Expr::Function { name, args } if name == "GROUPBY" => self.group_by(args, sheet, depth, skip, at),
            Expr::Function { name, args } if name == "PIVOTBY" => self.pivot_by(args, sheet, depth, skip, at),
            Expr::Function { name, args }
                if matches!(name.as_str(), "MAP" | "REDUCE" | "SCAN" | "BYROW" | "BYCOL" | "MAKEARRAY") =>
            {
                self.lambda_helper(name, args, sheet, depth, skip, at)
            }
            // A LAMBDA called where it is written: measured,
            // `LAMBDA(x,x+1)(5)` is 6.
            Expr::Function { name, args } if name == "_INVOKE" && !args.is_empty() => {
                let values = args[1..]
                    .iter()
                    .map(|expr| self.eval_arg_inner(expr, sheet, depth + 1, skip, at))
                    .collect();
                self.call_lambda(&args[0], values, sheet, depth, skip, at)
            }
            // Outside a LAMBDA nothing is left out: measured, ISOMITTED(1)
            // is FALSE.
            Expr::Function { name, args } if name == "ISOMITTED" && args.len() == 1 => {
                let left_out = matches!(&args[0], Expr::Name(label)
                    if self.omitted.borrow().iter().any(|held| held.eq_ignore_ascii_case(label)));
                Arg::Value(Value::Logical(left_out))
            }
            Expr::Function { name, args } if name == "LET" => {
                if args.len() < 3 || args.len() % 2 == 0 {
                    return Arg::Value(Value::Error(ExcelError::Value));
                }
                let mut bound = 0;
                let mut named_lambdas = 0;
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
                    if matches!(&pair[1], Expr::Function { name, .. } if name == "LAMBDA") {
                        self.lambda_scope.borrow_mut().push((label.clone(), pair[1].clone()));
                        named_lambdas += 1;
                    }
                    let value = if let Some((sheet, range)) = self.stored_reference(&pair[1], sheet, depth + 1, at) {
                        // LET still waits for a binding's precedents even when
                        // its final expression does not consume that binding.
                        // Deferring its cell reads must not run a UDF early.
                        if !self.reference_context.get() && !self.require_ready_reference(&sheet, &range) {
                            answer = Some(Arg::Value(Value::Error(ExcelError::Value)));
                            break;
                        }
                        LocalBinding::Reference { sheet, range }
                    } else {
                        let value = self.eval_arg_inner(&pair[1], sheet, depth + 1, skip, at);
                        if !self.reference_context.get() && !self.require_ready_value(&value) { answer = Some(Arg::Value(Value::Error(ExcelError::Value))); break; }
                        LocalBinding::Value(value)
                    };
                    self.let_scope.borrow_mut().push((label.clone(), value));
                    bound += 1;
                }
                let answer = answer.unwrap_or_else(|| {
                    self.eval_arg_inner(&args[args.len() - 1], sheet, depth + 1, skip, at)
                });
                let mut scope = self.let_scope.borrow_mut();
                let kept = scope.len() - bound;
                scope.truncate(kept);
                let mut lambdas = self.lambda_scope.borrow_mut();
                let kept = lambdas.len() - named_lambdas;
                lambdas.truncate(kept);
                answer
            }
            // A LAMBDA nobody calls is #CALC!: measured, `=LAMBDA(x,x)`.
            Expr::Function { name, .. } if name == "LAMBDA" => Arg::Value(Value::Error(ExcelError::Calc)),
            // A LAMBDA a LET named, called by that name.
            Expr::Function { name, args } if self.named_lambda(name, sheet, at).is_some() => {
                let lambda = self.named_lambda(name, sheet, at).expect("just asked");
                let values = args
                    .iter()
                    .map(|expr| self.eval_arg_inner(expr, sheet, depth + 1, skip, at))
                    .collect();
                self.call_lambda(&lambda, values, sheet, depth, skip, at)
            }
            // IF works out only the branch it takes, so a branch naming the
            // formula's own cell is no circle unless it is taken: measured,
            // `=IF(D1>0,5,G1)` in G1 is 5. A block for a condition is decided
            // cell by cell, and then every branch is worked out as before.
            Expr::Function { name, args } if name == "IF" && (2..=3).contains(&args.len()) => {
                let test = self.eval_arg_inner(&args[0], sheet, depth + 1, skip, at);
                if !self.require_ready_value(&test) { return Arg::Value(Value::Error(ExcelError::Value)); }
                let single = match &test {
                    Arg::Value(value) => Some(value.clone()),
                    Arg::Range(block) if block.cells.len() == 1 => Some(block.cells[0].clone()),
                    Arg::Range(_) => None,
                };
                let taken = single.as_ref().and_then(|value| match value {
                    Value::Number(n) => Some(*n != 0.0),
                    Value::Logical(state) => Some(*state),
                    Value::Blank => Some(false),
                    Value::Text(text) if text.eq_ignore_ascii_case("TRUE") => Some(true),
                    Value::Text(text) if text.eq_ignore_ascii_case("FALSE") => Some(false),
                    _ => None,
                });
                if !matches!(&test, Arg::Range(block) if block.is_array()) {
                    if let Some(selected) = taken {
                        return match args.get(if selected { 1 } else { 2 }) {
                            Some(branch) => self.eval_arg_inner(branch, sheet, depth + 1, skip, at),
                            None => Arg::Value(Value::Logical(false)),
                        };
                    }
                }
                let mut evaluated = vec![test];
                for (index, arg) in args.iter().enumerate().skip(1) {
                    let wanted = match taken {
                        Some(true) => index == 1,
                        Some(false) => index == 2,
                        None => true,
                    };
                    evaluated.push(if wanted {
                        self.eval_arg_inner(arg, sheet, depth + 1, skip, at)
                    } else {
                        Arg::Value(Value::Blank)
                    });
                }
                functions::call_arg_with_text_context(name, &evaluated, self.compatibility_version, self.dbcs_language.unwrap_or(true))
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
                } else if name == "AGGREGATE" {
                    // Measured: `AGGREGATE(9,5,B2:B6)` with row 3 hidden
                    // leaves row 3 out.
                    let option = args
                        .get(1)
                        .map(|a| self.eval_arg_inner(a, sheet, depth + 1, Skip::Nothing, at).scalar());
                    let option = match option {
                        Some(Value::Number(n)) => n.trunc() as i64,
                        _ => 0,
                    };
                    Skip::Aggregate {
                        nested: (0..=3).contains(&option),
                        hidden: matches!(option, 1 | 3 | 5 | 7),
                    }
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
                        // A name standing for several blocks is those blocks:
                        // measured, SUM of a name for A1:B2,D4 adds both.
                        Expr::Name(label)
                            if matches!(
                                self.name_bound(label, sheet, at),
                                Some(Expr::Function { name: union, .. }) if union == "_UNION"
                            ) =>
                        {
                            expanded = self.name_bound(label, sheet, at).unwrap_or_else(|| a.clone());
                            &expanded
                        }
                        _ => a,
                    };
                    let Expr::Function { name: union, args: areas } = a else {
                        evaluated.push(self.eval_function_argument((name, at_arg), a, sheet, depth + 1, skip, at));
                        continue;
                    };
                    if union != "_UNION" {
                        evaluated.push(self.eval_function_argument((name, at_arg), a, sheet, depth + 1, skip, at));
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
                        evaluated.push(Arg::Range(RangeData {origin: crate::functions::BlockOrigin::Array,  width: 1, height: cells.len(), cells }));
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
                // SUMIF uses the sum range's top-left and the criteria range's
                // dimensions, even when the written sum reference is shorter.
                if name == "SUMIF" && evaluated.len() == 3 {
                    if let (Some(criteria), Some((target, sum))) = (
                        match &evaluated[0] { Arg::Range(block) => block.cell_reference().map(|(_, range)| range), _ => None },
                        match &evaluated[2] { Arg::Range(block) => block.cell_reference(), _ => None },
                    ) {
                        let right = sum.start.col.checked_add(criteria.width() - 1);
                        let bottom = sum.start.row.checked_add(criteria.height() - 1);
                        evaluated[2] = match (right, bottom) {
                            (Some(right), Some(bottom)) if right <= MAX_COL && bottom <= MAX_ROW => {
                                self.reference_arg(&target, &RangeRef::normalised(sum.start, CellRef::new(right, bottom)), skip)
                            }
                            _ => Arg::Value(Value::Error(ExcelError::Ref)),
                        };
                    }
                }
                // INDEX's area number has done its work once the area is picked.
                if name == "INDEX"
                    && matches!(args.first(), Some(Expr::Function { name: union, .. }) if union == "_UNION")
                {
                    evaluated.truncate(3);
                }
                // INDEX with one index into CELLS two ways wide and tall has
                // no one cell to give: measured, `=INDEX(A1:B2,2)` is #REF!,
                // where the same into an array is its whole row.
                if name == "INDEX"
                    && args.len() == 2
                    && matches!(evaluated.first(), Some(Arg::Range(block)) if block.width > 1 && block.height > 1)
                    && self.reference_of(&args[0], sheet, depth + 1, at).is_some()
                {
                    return Arg::Value(Value::Error(ExcelError::Ref));
                }
                // EDATE and EOMONTH take an array a value at a time but refuse
                // a block of CELLS: measured, `EOMONTH(A1:A3,0)` is #VALUE!
                // wherever it stands, while `EDATE(A1,{1,2})` reads 32.
                if matches!(name.as_str(), "EDATE" | "EOMONTH")
                    && args.iter().zip(&evaluated).any(|(expr, arg)| {
                        matches!(expr, Expr::Ref(_) | Expr::Name(_))
                            && matches!(arg, Arg::Range(block) if block.cells.len() > 1)
                    })
                {
                    return Arg::Value(Value::Error(ExcelError::Value));
                }
                if user_function_pending() || self.value_deferred.get() { return Arg::Value(Value::Error(ExcelError::Value)); }
                if name == "IFS" && evaluated.len() >= 2 && evaluated.len() % 2 == 0 {
                    for pair in evaluated.chunks_exact(2) {
                        if matches!(&pair[0], Arg::Range(block) if block.is_array() || block.cells.len() > 1) { break; }
                        match pair[0].scalar().to_logical() {
                            Ok(true) => return pair[1].clone(),
                            Ok(false) => {},
                            Err(why) => return Arg::Value(Value::Error(why)),
                        }
                    }
                }
                if name == "SWITCH" && evaluated.len() >= 3 && !matches!(&evaluated[0], Arg::Range(block) if block.is_array() || block.cells.len() > 1) {
                    let subject = evaluated[0].scalar();
                    if let Value::Error(why) = subject { return Arg::Value(Value::Error(why)); }
                    let mut i = 1;
                    while i + 1 < evaluated.len() {
                        if matches!(&evaluated[i], Arg::Range(block) if block.is_array() || block.cells.len() > 1) { break; }
                        let candidate = evaluated[i].scalar();
                        if let Value::Error(why) = candidate { return Arg::Value(Value::Error(why)); }
                        if crate::value::compare(&subject, &candidate) == Ok(std::cmp::Ordering::Equal) { return evaluated[i + 1].clone(); }
                        i += 2;
                    }
                    if i + 1 == evaluated.len() { return evaluated[i].clone(); }
                }
                // A function the sheet does not know may be one the workbook's
                // macros define: `=Twice(A1)` asks the module's own Twice.
                if !functions::is_known_function(name) {
                    let legacy = at.is_some_and(|cell| self.sheets.get(sheet).is_some_and(|held| {
                        !held.dynamic.contains(&cell) && matches!(held.cells.get(&cell), Some(Cell::Formula { share: None, .. }))
                    }));
                    let handed: Vec<UserArg> = args
                        .iter()
                        .zip(&evaluated)
                        .enumerate()
                        .map(|(index, (expr, arg))| match (expr, arg) {
                            (Expr::Ref(reference), arg) if reference.book.is_none() => UserArg::Cells {
                                sheet: reference.sheet.clone().unwrap_or_else(|| sheet.to_string()),
                                range: reference.range,
                                values: arg.flatten(),
                            },
                            (_, Arg::Range(block)) if legacy && block.is_vector() && !(user_reference_argument(name, index) && block.cell_reference().is_some()) => UserArg::Value(block.cells.first().cloned().unwrap_or(Value::Blank)),
                            (_, Arg::Range(block)) => UserArg::Block(block.clone()),
                            (_, Arg::Value(value)) => UserArg::Value(value.clone()),
                        })
                        .collect();
                    let written = self.written_dependency_expression(expr, sheet, at);
                    let written: Vec<(String, RangeRef)> = self.ordering_expression(&written, sheet, at).value_references().into_iter()
                        .filter(|reference| reference.book.is_none())
                        .map(|reference| (reference.sheet.unwrap_or_else(|| sheet.to_string()), reference.range)).collect();
                    let caller = at.map(|(col, row)| (sheet.to_string(), col, row));
                    let lagging_indirect_arguments: BTreeSet<usize> = if self.provisional_value.get() {
                        args.iter().enumerate().filter_map(|(index, argument)|
                            ((self.expression_has_lagging_indirect_text(argument, sheet, at)
                                && (user_function_provisional_dynamic_text_value()
                                    || user_function_address_only_argument(name, index)))
                                || (user_function_wait_for_dynamic_lookups() && !user_function_full_rebuild()
                                    && self.expression_has_shared_indirect_precedent(argument, sheet, at)))
                                .then_some(index)
                        ).collect()
                    } else { BTreeSet::new() };
                    let value_argument_ranges: Vec<_> = handed.iter().enumerate().filter_map(|(index, argument)| {
                        if user_reference_argument(name, index) { return None; }
                        let selected_lookup = !user_function_wait_for_dynamic_lookups() && args.get(index).is_some_and(|argument| {
                            let mut xlookup = false;
                            argument.visit(&mut |node| if matches!(node, Expr::Function { name, .. } if name == "XLOOKUP") { xlookup = true; });
                            xlookup
                        });
                        if selected_lookup { return None; }
                        match argument {
                            UserArg::Cells { sheet, range, .. } => Some((index, sheet.clone(), *range)),
                            UserArg::Block(block) => match &block.origin {
                                functions::BlockOrigin::Cells { sheet, range } => Some((index, sheet.clone(), *range)),
                                _ => None,
                            },
                            UserArg::Value(_) => None,
                        }
                    }).collect();
                    let mut unready_formula_values: Vec<_> = if user_function_wait_for_dependencies() && self.provisional_value.get() {
                        self.unready_cells.borrow().iter().filter(|(target, (col, row))| {
                            value_argument_ranges.iter().any(|(_, source, range)|
                                self.reference_waits_for_formula(source, range, target, (*col, *row)))
                        }).cloned().collect()
                    } else { Vec::new() };
                    if user_function_wait_for_dependencies() && self.provisional_value.get() {
                        if let Some(coord) = at {
                            if let Some(formula) = self.expr_at(sheet, &coord) {
                                let formula = if user_function_wait_for_dynamic_lookups() { formula } else { omit_xlookup_return_arrays(&formula) };
                                let ordered_formula = self.ordering_expression(&self.dependency_expression(&formula, sheet, Some(coord)), sheet, Some(coord));
                                let references = ordered_formula.value_references();
                                unready_formula_values.extend(self.unready_cells.borrow().iter().filter(|(target, (col, row))| {
                                    references.iter().any(|reference| {
                                        self.reference_waits_for_formula(reference.sheet.as_deref().unwrap_or(sheet), &reference.range, target, (*col, *row))
                                    })
                                }).cloned());
                            }
                        }
                        if !user_function_wait_for_dynamic_lookups() {
                            unready_formula_values.retain(|(target, coord)| {
                                self.expr_at(target, coord).is_none_or(|formula| {
                                    let mut xlookup = false;
                                    formula.visit(&mut |node| if matches!(node, Expr::Function { name, .. } if name == "XLOOKUP") { xlookup = true; });
                                    !xlookup
                                })
                            });
                        }
                        unready_formula_values.sort();
                        unready_formula_values.dedup();
                        unready_formula_values.retain(|(target, coord)| {
                            let precedes_owner = at.and_then(|cell| self.calculation_order.get(&(sheet.to_string(), cell)))
                                .zip(self.calculation_order.get(&(target.clone(), *coord)))
                                .is_some_and(|(consumer, producer)| consumer < producer);
                            if !precedes_owner { return true; }
                            let probes_written_scalar = value_argument_ranges.iter().any(|(index, source, range)|
                                crate::case_fold::name_key(source) == crate::case_fold::name_key(target)
                                    && range.contains(coord.0, coord.1)
                                    && args.get(*index).is_some_and(|argument|
                                        Self::scalar_cell_reference(name, *index, argument).is_some()));
                            if probes_written_scalar { return false; }
                            // A chain consumer may probe an unpublished spill
                            // member before its later owner. A direct read of
                            // the formula anchor still waits for that value.
                            let reads_anchor = value_argument_ranges.iter().any(|(_, source, range)|
                                crate::case_fold::name_key(source) == crate::case_fold::name_key(target)
                                    && range.contains(coord.0, coord.1));
                            let reads_spill = value_argument_ranges.iter().any(|(_, source, range)|
                                !range.contains(coord.0, coord.1)
                                    && self.reference_waits_for_formula(source, range, target, *coord)
                                    && USER_FUNCTION_SPILL_PROBE_LAYOUT.with(|held| held.borrow().as_ref().is_none_or(|layout|
                                        layout.iter().any(|((owner_sheet, owner), (width, height))|
                                            owner_sheet == target && owner == coord
                                                && range.start.col < owner.0 + width && range.end.col >= owner.0
                                                && range.start.row < owner.1 + height && range.end.row >= owner.1))));
                            reads_anchor || !reads_spill
                        });
                    }
                    if !lagging_indirect_arguments.is_empty() {
                        unready_formula_values.retain(|(target, coord)| !value_argument_ranges.iter().any(|(index, source, range)| {
                            if !lagging_indirect_arguments.contains(index)
                                || crate::case_fold::name_key(source) != crate::case_fold::name_key(target)
                                || !range.contains(coord.0, coord.1) { return false; }
                            if !user_function_provisional_dynamic_text_value() { return true; }
                            let scalar_value_argument = !user_reference_argument(name, *index)
                                && !user_variant_argument(name, *index)
                                && !user_value_read_argument(name, *index);
                            if !scalar_value_argument { return true; }
                            let previous = self.provisional_text_formula_values_before_pass.borrow().get(&(target.clone(), *coord)).cloned();
                            let previous_nonzero = previous.clone().is_some_and(|value| !matches!(value, Value::Blank | Value::Number(0.0)));
                            if user_function_provisional_zero_text_probe() && matches!(previous, Some(Value::Number(0.0))) {
                                true
                            } else if user_function_full_rebuild() { !previous_nonzero } else { previous_nonzero }
                        }));
                    }
                    if !unready_formula_values.is_empty() {
                        // A scalar formula argument is only a provisional
                        // value until its formula precedent is published.
                        // Wait for that value before entering workbook VBA.
                        self.value_deferred.set(true);
                        record_user_function_reads(unready_formula_values);
                        return Arg::Value(Value::Error(ExcelError::Value));
                    }
                    if !lagging_indirect_arguments.is_empty() {
                        USER_PROVISIONAL_INDIRECT_TEXT_VALUE.with(|held| held.set(true));
                    }
                    let mut context = next_user_function_context(self.provisional_value.get());
                    context.dynamic_reference_argument_mask = args.iter().take(64).enumerate().fold(0u64, |mask, (index, argument)| {
                        let mut has_indirect = false;
                        argument.visit(&mut |node| {
                            if matches!(node, Expr::Function { name, .. } if name == "INDIRECT") { has_indirect = true; }
                        });
                        if has_indirect { mask | (1u64 << index) } else { mask }
                    });
                    let reference_retry_arguments = (!self.reference_context.get()
                        && ((depth > 0 && !user_function_wait_for_dependencies())
                            || (depth == 0 && user_function_full_calculation() && !user_function_full_rebuild()))
                        && USER_FUNCTION_REFERENCE_RETRY.with(|held| held.borrow().is_some())).then(|| handed.clone());
                    let root = depth == 0 && at.is_none_or(|coord| !self.scalar_replay_exclusions.contains(&(sheet.to_string(), coord)));
                    if let Some(mut answer) = ask_user_function(name, handed, caller.clone(), context, self.callback_read_view(), root, written.clone()) {
                        if let (Some(arguments), Arg::Range(block)) = (&reference_retry_arguments, &answer) {
                            if let Some((target, range)) = block.cell_reference() {
                                let ready = !self.unready_cells.borrow().iter().any(|(source, (col, row))|
                                    crate::case_fold::name_key(source) == crate::case_fold::name_key(&target)
                                        && range.contains(*col, *row));
                                let retry = USER_FUNCTION_REFERENCE_RETRY.with(|held| held.try_borrow_mut().ok()
                                    .and_then(|mut hook| hook.as_mut().map(|hook| hook(name, arguments, &caller, context, ready, depth == 0)))
                                    .unwrap_or(false));
                                if retry {
                                    answer = ask_user_function(name, arguments.clone(), caller.clone(), context,
                                        self.callback_read_view(), root, written.clone())
                                        .unwrap_or(Arg::Value(Value::Error(ExcelError::Value)));
                                }
                            }
                        }
                        self.callback_answer_observed.set(true);
                        if evaluated.iter().enumerate().any(|(index, argument)|
                            !user_reference_argument(name, index) && self.unready_argument(argument))
                            || USER_FORMULA_READS.with(|held| held.borrow().iter().any(|cell|
                                self.unready_cells.borrow().contains(cell))) {
                            self.value_deferred.set(true);
                        }
                        return match answer {
                            Arg::Range(block) => match block.cell_reference() {
                                Some((target, range)) => {
                                    if !self.reference_context.get() {
                                        // Its address was supplied by the callback, so the
                                        // written formula could not order this value read.
                                        self.provisional_value.set(true);
                                        let cells = self.unready_cells.borrow();
                                        let reads = cells.range((target.clone(), (range.start.col, 0))..=(target.clone(), (range.end.col, u32::MAX)))
                                            .filter(|(_, (col, row))| *col <= range.end.col && *row >= range.start.row && *row <= range.end.row).cloned();
                                        record_user_function_reads(reads);
                                    }
                                    self.reference_arg(&target, &range, skip)
                                },
                                None => Arg::Range(block),
                            },
                            other => other,
                        };
                    }
                }
                let reference_only = matches!(name.as_str(), "ROW" | "COLUMN" | "ROWS" | "COLUMNS" | "AREAS" | "ISREF" | "SHEET" | "SHEETS" | "FORMULATEXT" | "ISFORMULA");
                let arguments = if name == "INDEX" { evaluated.get(1..).unwrap_or(&[]) } else if reference_only { &evaluated[0..0] } else { &evaluated[..] };
                if arguments.iter().any(|argument| !self.require_ready_value(argument)) { return Arg::Value(Value::Error(ExcelError::Value)); }
                functions::call_arg_with_text_context(name, &evaluated, self.compatibility_version, self.dbcs_language.unwrap_or(true))
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
    fn a_table_column(&self, name: &str, asked: &str, sheet: &str, skip: Skip, at: At) -> Arg {
        match self.table_range(name, asked, sheet, at) {
            Ok((sheet, range)) => self.reference_arg(&sheet, &range, skip),
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
            self.tables.get(&crate::case_fold::table_key(name))
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
            TablePart::Headers if table.header_rows == 0 => return Err(ExcelError::Ref),
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
            if let Some(group) = self.array_context.get() {
                let (from, many) = if down { (group.anchor.1, group.height) } else { (group.anchor.0, group.width) };
                let cells = (0..many).map(|n| Value::Number(f64::from(from + n) + 1.0)).collect();
                return Arg::Range(RangeData {
                    origin: crate::functions::BlockOrigin::Array,
                    width: if down { 1 } else { many as usize },
                    height: if down { many as usize } else { 1 },
                    cells,
                });
            }
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
            _ => match self.eval_reference_argument(first, sheet, 0, at) {
                Arg::Range(block) => match block.cell_reference() { Some((_, range)) => range, None => return Arg::Value(Value::Error(ExcelError::Value)) },
                Arg::Value(Value::Error(why)) => return Arg::Value(Value::Error(why)),
                _ => return Arg::Value(Value::Error(ExcelError::Value)),
            },
        };
        let (from, to) = if down {
            (range.start.row, range.end.row)
        } else {
            (range.start.col, range.end.col)
        };
        let many = (to.saturating_sub(from) + 1) as usize;
        // Evaluate and legacy array formulas retain array identity even when
        // the reference measures one row or column. Ordinary cell formulas
        // pass a scalar result to their callers.
        let scalar_context = at.is_some_and(|coord| self.share_at(sheet, &coord).is_none());
        if many == 1 && scalar_context { return Arg::Value(Value::Number(from as f64 + 1.0)); }
        let cells: Vec<Value> = (from..=to)
            .map(|one| Value::Number(one as f64 + 1.0))
            .collect();
        Arg::Range(RangeData {origin: crate::functions::BlockOrigin::Array,
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
        if args.len() != 1 { return Arg::Value(Value::Error(ExcelError::Value)); }
        let range = match args.first() {
            Some(Expr::Ref(reference)) => reference.range,
            Some(Expr::Table { name, asked }) => match self.table_range(name, asked, sheet, at) {
                Ok((_, range)) => range,
                Err(why) => return Arg::Value(Value::Error(why)),
            },
            Some(other) => {
                let answer = self.eval_reference_argument(other, sheet, 0, at);
                match &answer {
                    Arg::Range(block) if block.cell_reference().is_some() => block.cell_reference().unwrap().1,
                    _ => {
                        if !self.require_ready_value(&answer) { return Arg::Value(Value::Error(ExcelError::Value)); }
                        return functions::call_arg_with_text_context(name, &[answer], self.compatibility_version, self.dbcs_language.unwrap_or(true));
                    },
                }
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

    /// GROUPBY(row_fields, values, function, [field_headers], [total_depth],
    /// [sort_order], [filter_array]). Measured: groups in ascending order of
    /// their fields, one aggregate column per value column, a "Total" row
    /// (total_depth 1, the default) aggregating every row kept, subtotals
    /// plus a "Grand Total" at depth 2, none at 0; headers 3 shows the first
    /// row as headings; sort_order n sorts by output column |n|, descending
    /// when negative; filter_array keeps the TRUE rows.
    fn group_by(&self, args: &[Expr], sheet: &str, depth: u32, skip: Skip, at: At) -> Arg {
        let bad = |why: ExcelError| Arg::Value(Value::Error(why));
        if args.len() < 3 {
            return bad(ExcelError::Value);
        }
        let fields = block_of(&self.eval_arg_inner(&args[0], sheet, depth + 1, skip, at));
        let values = block_of(&self.eval_arg_inner(&args[1], sheet, depth + 1, skip, at));
        if fields.height != values.height || fields.height == 0 {
            return bad(ExcelError::Value);
        }
        let number = |i: usize, default: f64| -> Result<f64, ExcelError> {
            match args.get(i) {
                None => Ok(default),
                Some(expr) => match self.eval_arg_inner(expr, sheet, depth + 1, skip, at).scalar() {
                    Value::Blank => Ok(default),
                    value => value.to_number(),
                },
            }
        };
        let headers_given = args.get(3).map(|expr| self.eval_arg_inner(expr, sheet, depth + 1, skip, at).scalar());
        let headers = match headers_given {
            None | Some(Value::Blank) => {
                // Left out, the first row is taken as headings, not shown,
                // when the values' first cell is text (measured on PIVOTBY).
                if matches!(values.at(0, 0), Value::Text(_)) { 1 } else { 0 }
            }
            Some(value) => match value.to_number() {
                Ok(n) => n as i64,
                Err(why) => return bad(why),
            },
        };
        let total_depth = match number(4, 1.0) {
            Ok(n) => n as i64,
            Err(why) => return bad(why),
        };
        let sort_by = match number(5, 0.0) {
            Ok(n) => n as i64,
            Err(why) => return bad(why),
        };
        let first_row = if headers == 1 || headers == 3 { 1 } else { 0 };
        let keep: Vec<bool> = match args.get(6) {
            None => vec![true; fields.height],
            Some(expr) => {
                let filter = block_of(&self.eval_arg_inner(expr, sheet, depth + 1, skip, at));
                (0..fields.height).map(|row| filter.at(0, row).to_logical().unwrap_or(false)).collect()
            }
        };
        let rows: Vec<usize> = (first_row..fields.height).filter(|row| keep.get(*row).copied().unwrap_or(false)).collect();
        let aggregate = |picked: &[usize], column: usize| -> Value {
            let cells: Vec<Value> = picked.iter().map(|row| values.at(column, *row)).collect();
            let block = Arg::Range(RangeData {origin: crate::functions::BlockOrigin::Array,  width: 1, height: cells.len(), cells });
            match &args[2] {
                Expr::Name(function) if function.eq_ignore_ascii_case("PERCENTOF") => {
                    let all: Vec<Value> = rows.iter().map(|row| values.at(column, *row)).collect();
                    let all = Arg::Range(RangeData {origin: crate::functions::BlockOrigin::Array,  width: 1, height: all.len(), cells: all });
                    crate::functions::call("PERCENTOF", &[block, all])
                }
                Expr::Name(function) => crate::functions::call(&function.to_uppercase(), &[block]),
                lambda => match self.call_lambda(lambda, vec![block], sheet, depth, skip, at) {
                    Arg::Value(value) => value,
                    Arg::Range(range) => range.cells.first().cloned().unwrap_or(Value::Blank),
                },
            }
        };
        let key_of = |row: usize| -> Vec<Value> { (0..fields.width).map(|col| fields.at(col, row)).collect() };
        let order = |a: &Value, b: &Value| crate::value::compare(a, b).unwrap_or(std::cmp::Ordering::Equal);
        // Distinct keys, each with its rows.
        let mut groups: Vec<(Vec<Value>, Vec<usize>)> = Vec::new();
        for row in &rows {
            let key = key_of(*row);
            match groups.iter_mut().find(|(held, _)| held.iter().zip(&key).all(|(x, y)| order(x, y).is_eq())) {
                Some((_, members)) => members.push(*row),
                None => groups.push((key, vec![*row])),
            }
        }
        groups.sort_by(|(a, _), (b, _)| {
            for (x, y) in a.iter().zip(b) {
                let o = order(x, y);
                if !o.is_eq() {
                    return o;
                }
            }
            std::cmp::Ordering::Equal
        });
        let width = fields.width + values.width;
        let line = |key: &[Value], picked: &[usize]| -> Vec<Value> {
            let mut out: Vec<Value> = key.to_vec();
            for column in 0..values.width {
                out.push(aggregate(picked, column));
            }
            out
        };
        let mut body: Vec<Vec<Value>> = groups.iter().map(|(key, members)| line(key, members)).collect();
        if sort_by != 0 {
            let column = (sort_by.unsigned_abs() as usize).saturating_sub(1).min(width - 1);
            body.sort_by(|a, b| {
                let o = order(&a[column], &b[column]);
                if sort_by < 0 { o.reverse() } else { o }
            });
        }
        let mut out: Vec<Vec<Value>> = Vec::new();
        if headers == 3 {
            let mut heading: Vec<Value> = (0..fields.width).map(|col| fields.at(col, 0)).collect();
            heading.extend((0..values.width).map(|col| values.at(col, 0)));
            out.push(heading);
        }
        let blank_key = |label: &str, first: Option<Value>| -> Vec<Value> {
            let mut key = vec![Value::Text(String::new()); fields.width];
            key[0] = first.unwrap_or_else(|| Value::Text(label.to_string()));
            key
        };
        if total_depth.abs() >= 2 && fields.width >= 2 {
            // Subtotals after each group of the first field.
            let mut i = 0;
            while i < body.len() {
                let lead = body[i][0].clone();
                let mut j = i;
                while j < body.len() && order(&body[j][0], &lead).is_eq() {
                    out.push(body[j].clone());
                    j += 1;
                }
                let members: Vec<usize> = groups
                    .iter()
                    .filter(|(key, _)| order(&key[0], &lead).is_eq())
                    .flat_map(|(_, members)| members.clone())
                    .collect();
                out.push(line(&blank_key("", Some(lead)), &members));
                i = j;
            }
            out.push(line(&blank_key("Grand Total", None), &rows));
        } else {
            out.extend(body);
            if total_depth != 0 {
                let total = line(&blank_key("Total", None), &rows);
                if total_depth < 0 {
                    let at_top = if headers == 3 { 1 } else { 0 };
                    out.insert(at_top, total);
                } else {
                    out.push(total);
                }
            }
        }
        let height = out.len();
        Arg::Range(RangeData {origin: crate::functions::BlockOrigin::Array,  width, height, cells: out.into_iter().flatten().collect() })
    }

    /// PIVOTBY(row_fields, col_fields, values, function, [field_headers],
    /// [row_total_depth], [row_sort_order], [col_total_depth],
    /// [col_sort_order], [filter_array], [relative_to]). Measured: a first
    /// row of the column keys (ascending) after a blank corner, then a row per
    /// row key (ascending); a "Total" column and row by default, first when
    /// the depth is negative, none at 0; a missing pairing is blank; headers
    /// 3 adds a row naming the column field and, below the keys, one naming
    /// the row field and the value under each column; row_sort n sorts the
    /// rows by output column |n|, col_sort n the columns by their key;
    /// PERCENTOF measures each cell against its column's total.
    fn pivot_by(&self, args: &[Expr], sheet: &str, depth: u32, skip: Skip, at: At) -> Arg {
        use std::cmp::Ordering;
        let bad = |why: ExcelError| Arg::Value(Value::Error(why));
        if args.len() < 4 {
            return bad(ExcelError::Value);
        }
        let block = |i: usize| block_of(&self.eval_arg_inner(&args[i], sheet, depth + 1, skip, at));
        let (row_fields, col_fields, values) = (block(0), block(1), block(2));
        let height = values.height;
        if row_fields.height != height || col_fields.height != height || height == 0 {
            return bad(ExcelError::Value);
        }
        let setting = |i: usize, default: i64| -> Result<i64, ExcelError> {
            match args.get(i) {
                None => Ok(default),
                Some(expr) => match self.eval_arg_inner(expr, sheet, depth + 1, skip, at).scalar() {
                    Value::Blank => Ok(default),
                    value => value.to_number().map(|n| n as i64),
                },
            }
        };
        let headers_default = if matches!(values.at(0, 0), Value::Text(_)) { 1 } else { 0 };
        let settings = (|| -> Result<[i64; 5], ExcelError> {
            Ok([setting(4, headers_default)?, setting(5, 1)?, setting(6, 0)?, setting(7, 1)?, setting(8, 0)?])
        })();
        let [headers, row_total, row_sort, col_total, col_sort] = match settings {
            Ok(held) => held,
            Err(why) => return bad(why),
        };
        let first_row = if headers == 1 || headers == 3 { 1 } else { 0 };
        let keep: Vec<bool> = match args.get(9) {
            None => vec![true; height],
            Some(expr) => {
                let filter = block_of(&self.eval_arg_inner(expr, sheet, depth + 1, skip, at));
                (0..height).map(|row| filter.at(0, row).to_logical().unwrap_or(false)).collect()
            }
        };
        let rows: Vec<usize> = (first_row..height).filter(|row| keep[*row]).collect();
        let order = |a: &Value, b: &Value| crate::value::compare(a, b).unwrap_or(Ordering::Equal);
        let same = |a: &[Value], b: &[Value]| a.iter().zip(b).all(|(x, y)| order(x, y).is_eq());
        let keys_of = |fields: &RangeData| -> Vec<(Vec<Value>, Vec<usize>)> {
            let mut groups: Vec<(Vec<Value>, Vec<usize>)> = Vec::new();
            for row in &rows {
                let key: Vec<Value> = (0..fields.width).map(|col| fields.at(col, *row)).collect();
                match groups.iter_mut().find(|(held, _)| same(held, &key)) {
                    Some((_, members)) => members.push(*row),
                    None => groups.push((key, vec![*row])),
                }
            }
            groups.sort_by(|(a, _), (b, _)| {
                a.iter().zip(b).map(|(x, y)| order(x, y)).find(|o| !o.is_eq()).unwrap_or(Ordering::Equal)
            });
            groups
        };
        let row_groups = keys_of(&row_fields);
        let mut col_groups = keys_of(&col_fields);
        if col_sort != 0 {
            let field = (col_sort.unsigned_abs() as usize).saturating_sub(1).min(col_fields.width - 1);
            col_groups.sort_by(|(a, _), (b, _)| {
                let o = order(&a[field], &b[field]);
                if col_sort < 0 { o.reverse() } else { o }
            });
        }
        let blank = || Value::Text(String::new());
        let aggregate = |picked: &[usize], relative: &[usize], column: usize| -> Value {
            if picked.is_empty() {
                return blank();
            }
            let range = |members: &[usize]| {
                let cells: Vec<Value> = members.iter().map(|row| values.at(column, *row)).collect();
                Arg::Range(RangeData {origin: crate::functions::BlockOrigin::Array,  width: 1, height: cells.len(), cells })
            };
            match &args[3] {
                Expr::Name(function) if function.eq_ignore_ascii_case("PERCENTOF") => {
                    crate::functions::call("PERCENTOF", &[range(picked), range(relative)])
                }
                Expr::Name(function) => crate::functions::call(&function.to_uppercase(), &[range(picked)]),
                lambda => match self.call_lambda(lambda, vec![range(picked)], sheet, depth, skip, at) {
                    Arg::Value(value) => value,
                    Arg::Range(range) => range.cells.first().cloned().unwrap_or(Value::Blank),
                },
            }
        };
        // The columns in order: each key's rows, the total's (all of them).
        let mut columns: Vec<(Option<Vec<Value>>, Vec<usize>)> =
            col_groups.iter().map(|(key, members)| (Some(key.clone()), members.clone())).collect();
        if col_total != 0 {
            let total = (None, rows.clone());
            if col_total < 0 {
                columns.insert(0, total);
            } else {
                columns.push(total);
            }
        }
        let cells_for = |row_members: &[usize]| -> Vec<Value> {
            let mut out = Vec::new();
            for (_, col_members) in &columns {
                let picked: Vec<usize> = row_members.iter().copied().filter(|row| col_members.contains(row)).collect();
                for column in 0..values.width {
                    out.push(aggregate(&picked, col_members, column));
                }
            }
            out
        };
        let mut body: Vec<Vec<Value>> = row_groups
            .iter()
            .map(|(key, members)| {
                let mut line = key.clone();
                line.extend(cells_for(members));
                line
            })
            .collect();
        let width = row_fields.width + columns.len() * values.width;
        if row_sort != 0 {
            let column = (row_sort.unsigned_abs() as usize).saturating_sub(1).min(width - 1);
            // A blank cell sorts last either way (measured).
            let empty = |v: &Value| matches!(v, Value::Blank) || matches!(v, Value::Text(t) if t.is_empty());
            body.sort_by(|a, b| match (empty(&a[column]), empty(&b[column])) {
                (true, false) => Ordering::Greater,
                (false, true) => Ordering::Less,
                _ => {
                    let o = order(&a[column], &b[column]);
                    if row_sort < 0 { o.reverse() } else { o }
                }
            });
        }
        let label = |fields: usize| -> Vec<Value> {
            let mut key = vec![blank(); fields];
            key[0] = Value::Text("Total".to_string());
            key
        };
        if row_total != 0 {
            let mut total = label(row_fields.width);
            total.extend(cells_for(&rows));
            if row_total < 0 {
                body.insert(0, total);
            } else {
                body.push(total);
            }
        }
        let mut out: Vec<Vec<Value>> = Vec::new();
        if headers == 3 {
            let mut line = vec![blank(); width];
            for level in 0..col_fields.width {
                line[row_fields.width + level] = col_fields.at(level, 0);
            }
            out.push(line);
        }
        for level in 0..col_fields.width {
            let mut line = vec![blank(); row_fields.width];
            for (key, _) in &columns {
                let name = match key {
                    Some(key) => key[level].clone(),
                    None if level == 0 => Value::Text("Total".to_string()),
                    None => blank(),
                };
                line.extend(std::iter::repeat_n(name, values.width));
            }
            out.push(line);
        }
        if headers == 3 {
            let mut line: Vec<Value> = (0..row_fields.width).map(|col| row_fields.at(col, 0)).collect();
            for _ in &columns {
                line.extend((0..values.width).map(|col| values.at(col, 0)));
            }
            out.push(line);
        }
        out.extend(body);
        let tall = out.len();
        Arg::Range(RangeData {origin: crate::functions::BlockOrigin::Array,  width, height: tall, cells: out.into_iter().flatten().collect() })
    }

    /// Call a LAMBDA written in place with these values for its
    /// parameters, each bound the way LET binds a name.
    /// A LAMBDA by name: one a LET named, else a defined name that holds
    /// one -- measured, Names.Add "Twice", "=LAMBDA(n,n*2)" makes
    /// `=Twice(8)` 16.
    fn dress_of(&self, sheet: &str, col: u32, row: u32) -> (Option<String>, Option<String>) {
        self.sheets
            .get(sheet)
            .and_then(|held| held.dress.get(&(col, row)).cloned())
            .unwrap_or_default()
    }

    fn named_lambda(&self, name: &str, sheet: &str, at: At) -> Option<Expr> {
        self.lambda_scope
            .borrow()
            .iter()
            .rev()
            .find(|(held, _)| held.eq_ignore_ascii_case(name))
            .map(|(_, lambda)| lambda.clone())
            .or_else(|| {
                self.name_bound(name, sheet, at)
                    .filter(|held| matches!(held, Expr::Function { name, .. } if name == "LAMBDA"))
            })
    }

    fn call_lambda(&self, lambda: &Expr, values: Vec<Arg>, sheet: &str, depth: u32, skip: Skip, at: At) -> Arg {
        if let Expr::Name(label) = lambda {
            return match self.named_lambda(label, sheet, at) {
                Some(named) => self.call_lambda(&named, values, sheet, depth + 1, skip, at),
                None => Arg::Value(Value::Error(ExcelError::Value)),
            };
        }
        let Expr::Function { name, args } = lambda else {
            return Arg::Value(Value::Error(ExcelError::Value));
        };
        if name != "LAMBDA" || args.is_empty() || values.len() > args.len() - 1 {
            return Arg::Value(Value::Error(ExcelError::Value));
        }
        // A parameter in brackets, `[b]`, may be left out: measured,
        // LAMBDA(a,[b],IF(ISOMITTED(b),"no b",b))(1) is "no b".
        let mut labels = Vec::with_capacity(args.len() - 1);
        for parameter in &args[..args.len() - 1] {
            match parameter {
                Expr::Name(label) => labels.push((label.clone(), false)),
                Expr::Table { name, asked } if name.is_empty() => labels.push((asked.to_uppercase(), true)),
                _ => return Arg::Value(Value::Error(ExcelError::Value)),
            }
        }
        if labels[values.len()..].iter().any(|(_, optional)| !optional) {
            return Arg::Value(Value::Error(ExcelError::Value));
        }
        let bound = labels.len();
        let mut left_out = 0;
        {
            let mut scope = self.let_scope.borrow_mut();
            let mut omitted = self.omitted.borrow_mut();
            let mut values = values.into_iter();
            for (label, _) in labels {
                match values.next() {
                    Some(value) => scope.push((label, LocalBinding::Value(value))),
                    None => {
                        omitted.push(label.clone());
                        left_out += 1;
                        scope.push((label, LocalBinding::Value(Arg::Value(Value::Logical(false)))));
                    }
                }
            }
        }
        let answer = self.eval_arg_inner(&args[args.len() - 1], sheet, depth + 1, skip, at);
        let mut scope = self.let_scope.borrow_mut();
        let kept = scope.len() - bound;
        scope.truncate(kept);
        let mut omitted = self.omitted.borrow_mut();
        let kept = omitted.len() - left_out;
        omitted.truncate(kept);
        answer
    }

    fn lambda_helper(&self, name: &str, args: &[Expr], sheet: &str, depth: u32, skip: Skip, at: At) -> Arg {
        let bad = || Arg::Value(Value::Error(ExcelError::Value));
        let Some((lambda, given)) = args.split_last() else {
            return bad();
        };
        let one = |answer: Arg| match answer {
            Arg::Value(value) => value,
            Arg::Range(block) if block.cells.len() == 1 => block.cells[0].clone(),
            Arg::Range(_) => Value::Error(ExcelError::Value),
        };
        let blocks: Vec<RangeData> = given
            .iter()
            .map(|expr| block_of(&self.eval_arg_inner(expr, sheet, depth + 1, skip, at)))
            .collect();
        match name {
            "MAP" => {
                let Some(first) = blocks.first() else {
                    return bad();
                };
                let (width, height) = (first.width, first.height);
                let mut cells = Vec::with_capacity(width * height);
                for row in 0..height {
                    for col in 0..width {
                        let values = blocks
                            .iter()
                            .map(|block| Arg::Value(reach(block, col, row).unwrap_or(Value::Error(ExcelError::NA))))
                            .collect();
                        cells.push(one(self.call_lambda(lambda, values, sheet, depth, skip, at)));
                    }
                }
                Arg::Range(RangeData {origin: crate::functions::BlockOrigin::Array,  width, height, cells })
            }
            "REDUCE" | "SCAN" => {
                let (mut total, data) = match blocks.as_slice() {
                    [data] => (Value::Blank, data),
                    [start, data] => (start.cells.first().cloned().unwrap_or(Value::Blank), data),
                    _ => return bad(),
                };
                let mut steps = Vec::with_capacity(data.cells.len());
                for value in &data.cells {
                    let values = vec![Arg::Value(total), Arg::Value(value.clone())];
                    total = one(self.call_lambda(lambda, values, sheet, depth, skip, at));
                    steps.push(total.clone());
                }
                if name == "REDUCE" {
                    Arg::Value(total)
                } else {
                    Arg::Range(RangeData {origin: crate::functions::BlockOrigin::Array,  width: data.width, height: data.height, cells: steps })
                }
            }
            "BYROW" | "BYCOL" => {
                let [data] = blocks.as_slice() else {
                    return bad();
                };
                let by_row = name == "BYROW";
                let lines = if by_row { data.height } else { data.width };
                let mut cells = Vec::with_capacity(lines);
                for line in 0..lines {
                    let piece = if by_row {
                        RangeData {origin: Default::default(),  width: data.width, height: 1, cells: (0..data.width).map(|col| data.at(col, line)).collect() }
                    } else {
                        RangeData {origin: Default::default(),  width: 1, height: data.height, cells: (0..data.height).map(|row| data.at(line, row)).collect() }
                    };
                    // One value per line: measured, BYROW(A1:B3,LAMBDA(r,r))
                    // is a single #CALC!.
                    match self.call_lambda(lambda, vec![Arg::Range(piece)], sheet, depth, skip, at) {
                        Arg::Range(block) if block.cells.len() > 1 => {
                            return Arg::Value(Value::Error(ExcelError::Calc));
                        }
                        answer => cells.push(one(answer)),
                    }
                }
                if by_row {
                    Arg::Range(RangeData {origin: crate::functions::BlockOrigin::Array,  width: 1, height: lines, cells })
                } else {
                    Arg::Range(RangeData {origin: crate::functions::BlockOrigin::Array,  width: lines, height: 1, cells })
                }
            }
            _ => {
                // MAKEARRAY(rows, columns, LAMBDA(r, c, ...)), counted from one.
                let [rows, columns] = blocks.as_slice() else {
                    return bad();
                };
                let count = |block: &RangeData| match block.cells.first().map(|value| value.to_number()) {
                    Some(Ok(n)) if n >= 1.0 => Some(n.trunc() as usize),
                    _ => None,
                };
                let (Some(height), Some(width)) = (count(rows), count(columns)) else {
                    return bad();
                };
                // No bigger than SEQUENCE will make.
                if height as f64 * width as f64 > 1_048_576.0 {
                    return bad();
                }
                let mut cells = Vec::with_capacity(width * height);
                for row in 1..=height {
                    for col in 1..=width {
                        let values = vec![Arg::Value(Value::Number(row as f64)), Arg::Value(Value::Number(col as f64))];
                        cells.push(one(self.call_lambda(lambda, values, sheet, depth, skip, at)));
                    }
                }
                Arg::Range(RangeData {origin: crate::functions::BlockOrigin::Array,  width, height, cells })
            }
        }
    }

    /// The sheet and range a reference-shaped expression names, for OFFSET.
    fn reference_of(&self, expr: &Expr, sheet: &str, depth: u32, at: At) -> Option<(String, RangeRef)> {
        match expr {
            Expr::Ref(reference) if reference.book.is_none() => Some((
                reference.sheet.clone().unwrap_or_else(|| sheet.to_string()),
                reference.range,
            )),
            Expr::Table { name, asked } => self.table_range(name, asked, sheet, at).ok(),
            Expr::Name(name) if self.let_scope.borrow().iter().any(|(label, _)| label == name) => {
                let scope = self.let_scope.borrow();
                scope.iter().rev().find(|(label, _)| label == name)?.1.reference()
            },
            Expr::Name(name) => match self.name_bound(name, sheet, at) {
                Some(bound) => {
                    let bound = bound.clone();
                    self.reference_of(&bound, sheet, depth, at)
                }
                None if self.tables.contains_key(&crate::case_fold::table_key(name)) => {
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
            Expr::Function { name, .. } if matches!(name.as_str(), "INDEX" | "TAKE" | "DROP" | "XLOOKUP") => {
                match self.eval_reference_argument(expr, sheet, depth + 1, at) { Arg::Range(block) => block.cell_reference(), _ => None }
            }
            Expr::Function { name, args } if name == "TRIMRANGE" => self.trimmed_range(args, sheet, depth, at),
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
            Expr::Function { .. } => match self.eval_reference_argument(expr, sheet, depth + 1, at) { Arg::Range(block) => block.cell_reference(), _ => None },
            _ => None,
        }
    }

    /// Select cells without evaluating any expression a second time.
    fn selection_reference(&self, name: &str, args: &[Arg]) -> Option<(String, RangeRef)> {
        let reference = |arg: &Arg| match arg { Arg::Range(block) => block.cell_reference(), _ => None };
        let whole = |at: usize| args.get(at)?.scalar().to_number().ok().map(|n| n.trunc() as i64);
        if name == "XLOOKUP" {
            let (target, returned) = reference(args.get(2)?)?;
            let looked = args.get(1)?.as_range();
            let asked = [args.first()?.clone(), args.get(1)?.clone(), args.get(4).cloned().unwrap_or(Arg::Value(Value::Number(0.0))), args.get(5).cloned().unwrap_or(Arg::Value(Value::Number(1.0)))];
            let Value::Number(position) = functions::call("XMATCH", &asked) else { return None };
            if position < 1.0 { return None; }
            let offset = position as u32 - 1;
            let down = functions::xlookup_axis((looked.width, looked.height), (returned.width() as usize, returned.height() as usize))?;
            let range = if down {
                if offset >= returned.height() { return None; }
                RangeRef::normalised(CellRef::new(returned.start.col, returned.start.row + offset), CellRef::new(returned.end.col, returned.start.row + offset))
            } else {
                if offset >= returned.width() { return None; }
                RangeRef::normalised(CellRef::new(returned.start.col + offset, returned.start.row), CellRef::new(returned.start.col + offset, returned.end.row))
            };
            return Some((target, range));
        }
        let (target, range) = reference(args.first()?)?;
        if name == "INDEX" {
            if !(2..=3).contains(&args.len()) { return None; }
            let (rows, cols) = (i64::from(range.height()), i64::from(range.width()));
            let (row, col) = match args.get(2) {
                Some(_) => (whole(1)?, whole(2)?),
                None if rows == 1 => (1, whole(1)?),
                None if cols == 1 => (whole(1)?, 1),
                None => return None,
            };
            if row < 0 || col < 0 || row > rows || col > cols { return None; }
            let (r0, r1) = if row == 0 { (range.start.row, range.end.row) } else { let r = range.start.row + row as u32 - 1; (r, r) };
            let (c0, c1) = if col == 0 { (range.start.col, range.end.col) } else { let c = range.start.col + col as u32 - 1; (c, c) };
            return Some((target, RangeRef::normalised(CellRef::new(c0, r0), CellRef::new(c1, r1))));
        }
        if !(2..=3).contains(&args.len()) { return None; }
        let count = |at| match args.get(at).map(Arg::scalar) { None | Some(Value::Blank) => Some(None), Some(value) => value.to_number().ok().map(|n| Some(n.trunc() as i64)) };
        let cut = |start: u32, end: u32, asked: Option<i64>| -> Option<(u32, u32)> {
            let Some(asked) = asked else { return Some((start, end)) };
            let length = u64::from(end - start + 1);
            let n = asked.unsigned_abs();
            let (lo, hi) = if name == "TAKE" {
                let n = n.min(length); if n == 0 { return None; }
                if asked > 0 { (0, n - 1) } else { (length - n, length - 1) }
            } else {
                if n >= length { return None; }
                if asked > 0 { (n, length - 1) } else { (0, length - n - 1) }
            };
            Some((start + lo as u32, start + hi as u32))
        };
        let (r0, r1) = cut(range.start.row, range.end.row, count(1)?)?;
        let (c0, c1) = cut(range.start.col, range.end.col, count(2)?)?;
        Some((target, RangeRef::normalised(CellRef::new(c0, r0), CellRef::new(c1, r1))))
    }

    /// TRIMRANGE(range, [rows], [columns]): the range less its empty
    /// outer rows and columns -- 0 keeps them, 1 trims the leading ones, 2
    /// the trailing, 3 (the default) both. Measured over B3:C5 filled inside
    /// A1:F10: $B$3:$C$5; rows 1 → $B$3:$C$10, rows 2 → $B$1:$C$5, rows 0
    /// → $B$1:$C$10; all empty is #REF!.
    fn trimmed_range(&self, args: &[Expr], sheet: &str, depth: u32, at: At) -> Option<(String, RangeRef)> {
        let (target, range) = self.reference_of(args.first()?, sheet, depth + 1, at)?;
        let mode = |i: usize| -> Option<u32> {
            match args.get(i) {
                None => Some(3),
                Some(expr) => match self.eval_arg(expr, sheet, depth + 1, at).scalar() {
                    Value::Blank => Some(3),
                    value => value.to_number().ok().map(|n| n as u32).filter(|n| *n <= 3),
                },
            }
        };
        let (rows, columns) = (mode(1)?, mode(2)?);
        let held = self.sheets.get(&target)?;
        let (mut low_row, mut high_row, mut low_col, mut high_col) = (u32::MAX, 0u32, u32::MAX, 0u32);
        for ((col, row), cell) in &held.cells {
            if *col < range.start.col || *col > range.end.col || *row < range.start.row || *row > range.end.row {
                continue;
            }
            let empty = matches!(cell, Cell::Literal(Value::Blank)) || matches!(cell, Cell::Literal(Value::Text(t)) if t.is_empty());
            if empty {
                continue;
            }
            low_row = low_row.min(*row);
            high_row = high_row.max(*row);
            low_col = low_col.min(*col);
            high_col = high_col.max(*col);
        }
        if low_row == u32::MAX {
            return None;
        }
        let pick = |mode: u32, start: u32, end: u32, low: u32, high: u32| match mode {
            1 => (low, end),
            2 => (start, high),
            3 => (low, high),
            _ => (start, end),
        };
        let (r0, r1) = pick(rows, range.start.row, range.end.row, low_row, high_row);
        let (c0, c1) = pick(columns, range.start.col, range.end.col, low_col, high_col);
        Some((target, RangeRef::normalised(CellRef::new(c0, r0), CellRef::new(c1, r1))))
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
                    let argument = self.eval_arg(expr, sheet, depth + 1, at);
                    if !self.require_ready_value(&argument) { return Err(ExcelError::Value); }
                    let value = argument.scalar();
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
            Ok((s, range)) => self.reference_arg(&s, &range, Skip::Nothing),
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
        let argument = self.eval_arg(&args[0], sheet, depth + 1, at);
        if !self.require_ready_value(&argument) { return Err(ExcelError::Value); }
        let text = match argument.scalar() {
            Value::Text(s) => s,
            Value::Error(why) => return Err(why),
            _ => return Err(ExcelError::Ref),
        };
        let a1 = match args.get(1) {
            None => true,
            Some(expr) => {
                let argument = self.eval_arg(expr, sheet, depth + 1, at);
                if !self.require_ready_value(&argument) { return Err(ExcelError::Value); }
                argument.scalar().to_logical().unwrap_or(true)
            },
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

    fn indirect_text_precedent_is_later(&self, args: &[Expr], sheet: &str, at: At) -> bool {
        let selected = self.pass_cells.borrow();
        let caller_order = at.and_then(|coord| self.calculation_order.get(&(sheet.to_string(), coord))).copied();
        args.iter().any(|arg| {
            let resolved = self.dependency_expression(arg, sheet, at);
            let ordered = self.ordering_expression(&resolved, sheet, at);
            ordered.value_references().iter().any(|reference| {
                let target = reference.sheet.as_deref().unwrap_or(sheet);
                let computed_scalar = reference.range.is_single()
                    && self.expr_at(target, &reference.range.start.coord()).is_some()
                    && (caller_order.zip(self.calculation_order.get(&(target.to_string(), reference.range.start.coord())).copied())
                        .is_some_and(|(caller, text)| text > caller)
                        || (target.eq_ignore_ascii_case(sheet)
                            && at.is_some_and(|caller| reference.range.start.coord() > caller)));
                computed_scalar || selected.range((target.to_string(), (reference.range.start.col, 0))..=(target.to_string(), (reference.range.end.col, u32::MAX)))
                    .filter(|(_, (col, row))| *col <= reference.range.end.col && *row >= reference.range.start.row && *row <= reference.range.end.row)
                    .any(|cell| caller_order.zip(self.calculation_order.get(cell).copied())
                        .is_none_or(|(caller, text)| text > caller))
            })
        })
    }

    fn expression_has_lagging_indirect_text(&self, expr: &Expr, sheet: &str, at: At) -> bool {
        let mut found = false;
        expr.visit(&mut |node| {
            if let Expr::Function { name, args } = node {
                if name == "INDIRECT" && self.indirect_text_precedent_is_later(args, sheet, at) { found = true; }
            }
        });
        found
    }

    fn expression_has_shared_indirect_precedent(&self, expr: &Expr, sheet: &str, at: At) -> bool {
        let mut found = false;
        expr.visit(&mut |node| {
            let Expr::Function { name, args } = node else { return; };
            if name != "INDIRECT" || args.is_empty() { return; }
            let Some(Expr::Ref(selector)) = args.first() else { return; };
            if selector.book.is_some() || !selector.range.is_single() { return; }
            let selector_sheet = selector.sheet.as_deref().unwrap_or(sheet);
            let selector_coord = selector.range.start.coord();
            let Some(selector_expr) = self.expr_at(selector_sheet, &selector_coord) else { return; };
            let Ok((target_sheet, target_range)) = self.indirect_range(args, sheet, 0, at) else { return; };
            if !target_range.is_single() { return; }
            let Some(target_expr) = self.expr_at(&target_sheet, &target_range.start.coord()) else { return; };
            let previous_target = self.formula_values_before_pass.borrow()
                .get(&(target_sheet.clone(), target_range.start.coord())).cloned();
            if previous_target.is_none_or(|value| matches!(value, Value::Blank | Value::Number(0.0))) { return; }
            let selector_reads = self.dependency_expression(&selector_expr, selector_sheet, Some(selector_coord)).value_references();
            let target_reads = self.dependency_expression(&target_expr, &target_sheet, Some(target_range.start.coord())).value_references();
            if selector_reads.iter().any(|left| target_reads.iter().any(|right| {
                left.book == right.book
                    && left.sheet.as_deref().unwrap_or(selector_sheet).eq_ignore_ascii_case(right.sheet.as_deref().unwrap_or(&target_sheet))
                    && left.range.intersect(&right.range).is_some()
            })) { found = true; }
        });
        found
    }

    fn indirect_reference(&self, args: &[Expr], sheet: &str, depth: u32, at: At) -> Arg {
        match self.indirect_range(args, sheet, depth, at) {
            Ok((target, range)) => {
                let formula_selector = args.first().and_then(|selector| match selector {
                    Expr::Ref(reference) if reference.book.is_none() && reference.range.is_single() => {
                        let selector_sheet = reference.sheet.as_deref().unwrap_or(sheet);
                        self.expr_at(selector_sheet, &reference.range.start.coord())
                            .map(|_| (selector_sheet.to_string(), reference.range.start.coord()))
                    }
                    _ => None,
                });
                if !self.reference_context.get() {
                    self.provisional_value.set(true);
                    let cells = self.unready_cells.borrow();
                    let reads = cells.range((target.clone(), (range.start.col, 0))..=(target.clone(), (range.end.col, u32::MAX)))
                        .filter(|(_, (col, row))| *col <= range.end.col && *row >= range.start.row && *row <= range.end.row).cloned();
                    record_user_function_reads(reads);
                } else if self.dynamic_reference_wait.get() {
                    if let Some((selector_sheet, selector_at)) = &formula_selector {
                    let cells = self.unready_cells.borrow();
                    let reads: Vec<_> = cells.range((target.clone(), (range.start.col, 0))..=(target.clone(), (range.end.col, u32::MAX)))
                        .filter(|(_, (col, row))| *col <= range.end.col && *row >= range.start.row && *row <= range.end.row).cloned().collect();
                    let caller_order = at.and_then(|coord| self.calculation_order.get(&(sheet.to_string(), coord))).copied();
                    let follows_caller = reads.iter().any(|cell| caller_order.zip(self.calculation_order.get(cell).copied())
                        .is_none_or(|(caller, producer)| producer > caller));
                    let selector_expr = self.expr_at(selector_sheet, selector_at);
                    let target_expr = range.is_single().then(|| self.expr_at(&target, &range.start.coord())).flatten();
                    let shared_precedent = selector_expr.as_ref().zip(target_expr.as_ref()).is_some_and(|(selector, producer)| {
                        let selector_reads = self.dependency_expression(selector, selector_sheet, Some(*selector_at)).value_references();
                        let producer_reads = self.dependency_expression(producer, &target, Some(range.start.coord())).value_references();
                        selector_reads.iter().any(|left| producer_reads.iter().any(|right| {
                            left.book == right.book
                                && left.sheet.as_deref().unwrap_or(selector_sheet).eq_ignore_ascii_case(right.sheet.as_deref().unwrap_or(&target))
                                && left.range.intersect(&right.range).is_some()
                        }))
                    });
                    let selector_unchanged = self.formula_values_before_pass.borrow()
                        .get(&(selector_sheet.clone(), *selector_at)).cloned()
                        .is_some_and(|previous| previous == self.eval_arg(&args[0], sheet, depth + 1, at).scalar().clone());
                    if follows_caller && !(self.variant_value_read_context.get() && shared_precedent && selector_unchanged) {
                        record_user_function_reads(reads);
                        self.value_deferred.set(true);
                    }
                    }
                }
                let answer = self.reference_arg(&target, &range, Skip::Nothing);
                if !self.reference_context.get() {
                    let calculated_text = self.indirect_text_precedent_is_later(args, sheet, at);
                    let shared_scalar_probe = user_function_wait_for_dynamic_lookups()
                        && !user_function_full_rebuild()
                        && self.provisional_value.get()
                        && self.expression_has_shared_indirect_precedent(&Expr::Function { name: "INDIRECT".into(), args: args.to_vec() }, sheet, at)
                        && USER_FUNCTION_ARGUMENT_CONTEXT.with(|held| held.borrow().as_ref().is_some_and(|(function, index)| !user_reference_argument(function, *index)));
                    // A text precedent still behind its consumer starts a
                    // separate readiness phase. An already placed precedent
                    // retains the normal provisional callback value probes.
                    if calculated_text && !((user_function_provisional_dynamic_text_value()
                        || (user_function_wait_for_dependencies() && user_function_current_argument_is_address_only())
                        || shared_scalar_probe) && self.provisional_value.get())
                        && !self.require_ready_value(&answer)
                    { return Arg::Value(Value::Error(ExcelError::Value)); }
                }
                answer
            },
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
                Skip::Aggregate { nested, hidden } => {
                    (nested && self.is_subtotal_cell(sheet, col, row))
                        || (hidden && held.is_some_and(|held| held.hidden_rows.contains(&row)))
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
fn element_of(worked: &Arg, (dx, dy): (u32, u32)) -> Value {
    match worked {
        Arg::Value(value) => value.clone(),
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
        // A number past what a cell holds is #NUM!: measured, `=10^400`
        // and `=FACT(171)`.
        Value::Number(number) if !number.is_finite() => Value::Error(ExcelError::Num),
        other => other,
    }
}

fn apply_unary(op: UnaryOp, v: Value) -> Value {
    if op == UnaryOp::Plus { return v; }
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
    if left.cells.len() == 1 && right.cells.len() == 1 && !left.is_array() && !right.is_array() {
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
    Arg::Range(RangeData {origin: RangeData::value_origin([&left, &right]),
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
        // Two numbers are compared at fifteen significant digits, as Excel
        // keeps them: measured, `=0.1+0.2=0.3` is TRUE.
        if let (Value::Number(x), Value::Number(y)) = (&a, &b) {
            let (x, y) = (fifteen_digits(*x), fifteen_digits(*y));
            let ord = x.partial_cmp(&y).unwrap_or(std::cmp::Ordering::Equal);
            return Value::Logical(match op {
                BinaryOp::Eq => ord.is_eq(),
                BinaryOp::Ne => !ord.is_eq(),
                BinaryOp::Lt => ord.is_lt(),
                BinaryOp::Le => ord.is_le(),
                BinaryOp::Gt => ord.is_gt(),
                _ => ord.is_ge(),
            });
        }
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
        return match (a.to_utf16(), b.to_utf16()) {
            (Ok(mut x), Ok(y)) => { x.extend(y); Value::from_utf16(x) },
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
            let r = crate::functions::excel_power(x, y);
            if r.is_nan() {
                Value::Error(ExcelError::Num)
            } else {
                Value::Number(r)
            }
        }
        _ => unreachable!("comparison and concat handled above"),
    }
}

/// A number as its fifteen significant digits.
fn fifteen_digits(x: f64) -> f64 {
    if x == 0.0 || !x.is_finite() {
        return x;
    }
    format!("{x:.14e}").parse().unwrap_or(x)
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
    #[test]
    fn utf16_formula_calculation_and_formulatext_preserve_units() {
        let mut workbook = book();
        let formula = vec![b'=' as u16, b'"' as u16, b'a' as u16, 0xD83D, b'b' as u16, b'"' as u16];
        workbook.set_formula_utf16("Sheet1", "A1", &formula).unwrap();
        assert_eq!(workbook.formula_utf16("Sheet1", "A1"), Some(formula.clone()));
        workbook.set_formula("Sheet1", "B1", "=FORMULATEXT(A1)").unwrap();
        workbook.set_formula("Sheet1", "C1", "=A1 & \"!\"").unwrap();
        workbook.recalculate();
        assert_eq!(workbook.value("Sheet1", "A1"), Value::Utf16(vec![97, 0xD83D, 98]));
        assert_eq!(workbook.value("Sheet1", "B1"), Value::Utf16(formula));
        assert_eq!(workbook.value("Sheet1", "C1"), Value::Utf16(vec![97, 0xD83D, 98, 33]));
    }

    #[test]
    fn formula_display_keeps_original_source_when_evaluation_adds_coercions() {
        let mut workbook = book();
        for row in 1..=3 { workbook.set_value("Sheet1",&format!("A{row}"),Value::Number(row as f64)).unwrap(); }
        workbook.define_name("Changing","=Sheet1!$A$1:$A$3").unwrap();
        workbook.set_formula("Sheet1","B2","=@Changing").unwrap();
        workbook.remember_formula_source("Sheet1","B2","Changing",None).unwrap();
        workbook.set_formula("Sheet1","C2","=FORMULATEXT(B2)").unwrap();
        workbook.set_formula("Sheet1","A4","=_xlfn.SINGLE(A1:A3)").unwrap();
        workbook.set_formula("Sheet1","B4","=FORMULATEXT(A4)").unwrap();
        workbook.recalculate();
        assert_eq!(workbook.value("Sheet1","B2"),Value::Number(2.));
        assert_eq!(workbook.formula("Sheet1","B2"),Some("Changing"));
        assert_eq!(workbook.value("Sheet1","C2"),Value::text("=Changing"));
        assert_eq!(workbook.value("Sheet1","B4"),Value::text("=@A1:A3"));
    }

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
    fn queued_callbacks_keep_the_values_from_their_request() {
        let mut wb = book();
        wb.set_formula("Sheet1", "A1", "=CAPTURE()").unwrap();
        wb.set_formula("Sheet1", "C1", "=SEQUENCE(2)").unwrap();
        wb.set_dynamic("Sheet1", "C1").unwrap();
        wb.set_formula("Sheet1", "Z1", "=CAPTURE()").unwrap();
        let views = std::rc::Rc::new(std::cell::RefCell::new(Vec::new()));
        let captured = views.clone();
        set_user_function_hook(Some(Box::new(move |name, _, _| {
            assert_eq!(name, "CAPTURE");
            captured.borrow_mut().push(user_function_read_view().unwrap());
            Some(Arg::Value(Value::Number(0.0)))
        })));
        wb.recalculate();
        set_user_function_hook(None);
        assert_eq!(wb.value("Sheet1", "C2"), Value::Number(2.0));
        drop(wb);
        let views = views.borrow();
        assert_eq!(views.len(), 2);
        assert_eq!(views[0].value("Sheet1", 2, 1), Some(Value::Blank));
        assert_eq!(views[1].value("Sheet1", 2, 1), Some(Value::Number(2.0)));
        assert!(user_function_read_view().is_none());
    }

    #[test]
    fn rebuilt_order_separates_written_references_from_cached_spill_members() {
        let mut wb = book();
        wb.set_formula("Sheet1", "M1", "=PRODUCER()").unwrap();
        wb.set_formula("Sheet1", "N1", "=CONSUMER(M1)").unwrap();
        wb.set_formula("Sheet1", "X1", "=ARRAY_PRODUCER()").unwrap();
        wb.set_dynamic("Sheet1", "X1").unwrap();
        wb.set_spilled_value("Sheet1", "Y1", Value::Number(2.0)).unwrap();
        wb.set_cached_spill("Sheet1", "X1", 2, 1).unwrap();
        wb.set_formula("Sheet1", "Z1", "=CONSUMER(Y1)").unwrap();
        wb.set_calculation_order(&[("Sheet1".into(), (13, 0)), ("Sheet1".into(), (25, 0)), ("Sheet1".into(), (12, 0)), ("Sheet1".into(), (23, 0))]);
        let order = wb.rebuilt_calculation_order();
        let position = |col| order.iter().position(|(sheet, at)| sheet == "Sheet1" && *at == (col, 0)).unwrap();
        assert!(position(12) < position(13), "a written formula reference establishes the rebuilt traversal");
        assert!(position(25) < position(23), "cached spill members remain work discovered during calculation");
    }

    #[test]
    fn spill_refresh_keeps_the_first_accepted_traversal() {
        let mut wb = book();
        wb.set_formula("Sheet1", "N1", "=ARRAY_SOURCE()").unwrap();
        wb.set_dynamic("Sheet1", "N1").unwrap();
        wb.set_cached("Sheet1", "N1", Value::Number(1.0)).unwrap();
        wb.set_formula("Sheet1", "Q1", "=OBSERVE(N2)").unwrap();
        let wanted = vec![("Sheet1".into(), (16, 0)), ("Sheet1".into(), (13, 0))];
        wb.set_calculation_order(&wanted);wb.set_chain_probes(true);
        let calls = std::rc::Rc::new(std::cell::RefCell::new(Vec::new()));
        let seen = calls.clone();
        set_user_function_hook(Some(Box::new(move |name, args, _| {
            if name == "ARRAY_SOURCE" {
                seen.borrow_mut().push("array".to_string());
                return Some(Arg::Range(RangeData { origin: crate::functions::BlockOrigin::Array, width: 1, height: 2, cells: vec![Value::Number(1.0), Value::Number(2.0)] }));
            }
            assert_eq!(name, "OBSERVE");
            let value = match &args[0] {
                UserArg::Cells { values, .. } => values.first().cloned().unwrap_or(Value::Blank),
                UserArg::Value(value) => value.clone(),
                _ => panic!("expected one cell"),
            };
            let value = if value == Value::Blank { Value::Number(0.0) } else { value };
            seen.borrow_mut().push(format!("observe:{value:?}"));
            Some(Arg::Value(value))
        })));
        let report = wb.recalculate_dirty(&wanted, None);
        set_user_function_hook(None);
        assert_eq!(wb.value("Sheet1", "Q1"), Value::Number(2.0));
        assert_eq!(&*calls.borrow(), &["observe:Number(0.0)", "array", "observe:Number(2.0)"]);
        assert_eq!(report.traversal_order, wanted);
        assert_eq!(report.calculated, vec![("Sheet1".into(), (13, 0)), ("Sheet1".into(), (16, 0))]);
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
    fn hidden_table_headers_are_ref_errors_including_on_the_first_sheet_row() {
        for first_row in [0, 1] {
            let mut wb = book();
            wb.add_table("Sheet1", "tbl", (first_row, first_row + 1), (0, 1), 0, 0, vec!["X".into(), "Y".into()]);
            wb.set_value("Sheet1", &format!("A{}", first_row + 1), Value::Number(7.0)).unwrap();
            for (address, formula) in [
                ("E1", "=ROWS(tbl[#Headers])"),
                ("E2", "=ROW(tbl[[#Headers],[X]])"),
                ("E3", "=SUM(tbl[X])"),
                ("E4", "=ROWS(tbl[#All])"),
            ] { wb.set_formula("Sheet1", address, formula).unwrap(); }
            wb.recalculate();
            assert_eq!(wb.value("Sheet1", "E1"), Value::Error(ExcelError::Ref));
            assert_eq!(wb.value("Sheet1", "E2"), Value::Error(ExcelError::Ref));
            assert_eq!(wb.value("Sheet1", "E3"), Value::Number(7.0));
            assert_eq!(wb.value("Sheet1", "E4"), Value::Number(2.0));
        }
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

    /// COUNTIF over blanks, errors and logicals, every answer Excel's:
    /// `"<>x"` counts a blank and an error, `""` counts neither an error nor
    /// text of spaces, and `"TRUE"` finds the logical TRUE.
    #[test]
    fn countif_reads_blanks_errors_and_logicals_as_excel_does() {
        let mut wb = Workbook::new();
        wb.add_sheet("Sheet1");
        wb.set_value("Sheet1", "A1", Value::Text("apple".into())).unwrap();
        wb.set_formula("Sheet1", "A8", "=x").unwrap();
        wb.set_value("Sheet1", "A9", Value::Text("  ".into())).unwrap();
        wb.set_value("Sheet1", "A10", Value::Logical(true)).unwrap();
        wb.set_formula("Sheet1", "B1", "=COUNTIF(A8,\"\")").unwrap();
        wb.set_formula("Sheet1", "B2", "=COUNTIF(A2,\"<>apple\")").unwrap();
        wb.set_formula("Sheet1", "B3", "=COUNTIF(A8,\"<>apple\")").unwrap();
        wb.set_formula("Sheet1", "B4", "=COUNTIF(A1:A10,\"TRUE\")").unwrap();
        wb.set_formula("Sheet1", "B5", "=COUNTIF(A1:A10,\"\")").unwrap();
        wb.set_formula("Sheet1", "B6", "=COUNTIF(A1:A10,\"<>apple\")").unwrap();
        wb.recalculate();
        assert_eq!(wb.value("Sheet1", "B1"), Value::Number(0.0));
        assert_eq!(wb.value("Sheet1", "B2"), Value::Number(1.0));
        assert_eq!(wb.value("Sheet1", "B3"), Value::Number(1.0));
        assert_eq!(wb.value("Sheet1", "B4"), Value::Number(1.0));
        assert_eq!(wb.value("Sheet1", "B5"), Value::Number(6.0));
        assert_eq!(wb.value("Sheet1", "B6"), Value::Number(9.0));
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
    fn table_and_named_references_keep_subtotal_and_aggregate_exclusions() {
        for filtered in [false, true] {
            let mut wb = book();
            wb.add_table("Sheet1", "T", (0, 5), (0, 0), 1, 0, vec!["X".into()]);
            for (row, value) in [(2, 1.0), (3, 10.0), (4, 100.0), (6, 1000.0)] {
                wb.set_value("Sheet1", &format!("A{row}"), Value::Number(value)).unwrap();
            }
            wb.set_formula("Sheet1", "A5", "=SUBTOTAL(9,A2:A4)").unwrap();
            wb.hide_row("Sheet1", 2, filtered);
            wb.hide_row("Sheet1", 3, filtered);
            wb.define_name("Held", "=T[X]").unwrap();
            for (row, formula) in [
                (1, "=SUBTOTAL(9,T[X])"), (2, "=SUBTOTAL(109,T[X])"),
                (3, "=SUBTOTAL(9,T)"), (4, "=SUBTOTAL(109,T)"),
                (5, "=AGGREGATE(9,0,T[X])"), (6, "=AGGREGATE(9,5,T[X])"),
                (7, "=SUBTOTAL(9,Held)"), (8, "=SUBTOTAL(109,Held)"),
            ] { wb.set_formula("Sheet1", &format!("E{row}"), formula).unwrap(); }
            wb.recalculate();
            let subtotal = if filtered { 1001.0 } else { 1111.0 };
            for (row, expected) in [(1, subtotal), (2, 1001.0), (3, subtotal), (4, 1001.0), (5, 1111.0), (6, if filtered { 1002.0 } else { 1112.0 }), (7, subtotal), (8, 1001.0)] {
                assert_eq!(wb.value("Sheet1", &format!("E{row}")), Value::Number(expected), "filtered={filtered} E{row}");
            }
        }
    }

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
        // Measured: a FILTER keeping nothing, with no stand-in, is #CALC!
        // (Error 2050 in the cell).
        assert_eq!(shown("D9"), Value::Error(ExcelError::Calc), "and nothing to say");
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

/// CELL("format")'s code for a number format, with whether it puts the
/// positive numbers in brackets and colours the negative ones. Measured on
/// a Japanese Excel: General, fractions, text and elapsed times are G; 0 F0,
/// 0.00 F2, #,##0 ,0, 0% P0, 0.00E+00 S2, $#,##0.00 C2; a date with year,
/// month and day D1 (time or not), a month and year D2, a month and day D3,
/// an era D4; h:mm:ss AM/PM D6, h:mm AM/PM D7, h:mm:ss D8, h:mm D9; "-" after
/// when the negative section has a colour, "()" when the positive section
/// has a bracket.
fn cell_format_code(format: &str) -> (String, bool, bool) {
    let sections = split_format_sections(format);
    let first = sections.first().cloned().unwrap_or_default();
    let parentheses = first.contains('(');
    let color = sections.get(1).is_some_and(|negative| {
        let lower = negative.to_lowercase();
        ["[black]", "[blue]", "[cyan]", "[green]", "[magenta]", "[red]", "[white]", "[yellow]", "[color"]
            .iter()
            .any(|name| lower.contains(name))
    });
    let lower = first.to_lowercase();
    // Brackets other than an elapsed time's say nothing of the kind.
    let elapsed = lower.contains("[h") || lower.contains("[m") || lower.contains("[s");
    let bare: String = {
        let mut out = String::new();
        let mut depth = 0;
        for character in lower.chars() {
            match character {
                '[' => depth += 1,
                ']' => depth -= 1,
                _ if depth > 0 => {}
                _ => out.push(character),
            }
        }
        out
    };
    let era = lower.contains("$-411") && (bare.contains('g') || bare.contains('e'));
    let has = |c: char| bare.contains(c);
    // Minutes and seconds with no hour are no time of day: measured, mm:ss is G.
    let minutes_only = has('s') && !has('h') && !has('d') && !has('y');
    let mut code = if lower.is_empty() || lower == "general" || lower == "@" || has('?') || elapsed || minutes_only {
        "G".to_string()
    } else if era {
        "D4".to_string()
    } else if has('y') || has('d') || (has('m') && !has('0') && !has('#')) || has('h') || has('s') {
        let ampm = bare.contains("am/pm") || bare.contains("a/p");
        let (year, day, hour, second) = (has('y'), has('d'), has('h'), has('s'));
        let month = bare.replace("am/pm", "").contains('m') && !hour || (hour && bare.matches('m').count() > 1);
        if year && day {
            "D1".to_string()
        } else if year && month {
            "D2".to_string()
        } else if day || (month && !hour) {
            if !hour && !second && bare.contains('m') && !day && !year { "G".to_string() } else { "D3".to_string() }
        } else if hour {
            match (second, ampm) {
                (true, true) => "D6".to_string(),
                (false, true) => "D7".to_string(),
                (true, false) => "D8".to_string(),
                (false, false) => "D9".to_string(),
            }
        } else {
            "G".to_string()
        }
    } else {
        let decimals = bare
            .split_once('.')
            .map_or(0, |(_, after)| after.chars().take_while(|c| matches!(c, '0' | '#' | '?')).count());
        let grouped = {
            let chars: Vec<char> = bare.chars().collect();
            chars.windows(3).any(|w| matches!(w[0], '0' | '#') && w[1] == ',' && matches!(w[2], '0' | '#'))
        };
        let kind = if bare.contains('%') {
            "P"
        } else if bare.contains("e+") || bare.contains("e-") {
            "S"
        } else if bare.contains('$') {
            "C"
        } else if grouped {
            ","
        } else {
            "F"
        };
        format!("{kind}{decimals}")
    };
    if color {
        code.push('-');
    }
    if parentheses {
        code.push_str("()");
    }
    (code, parentheses, color)
}

/// A number format's sections, with quoted text and escaped characters
/// taken out so that what is left is the pattern.
fn split_format_sections(format: &str) -> Vec<String> {
    let mut sections = vec![String::new()];
    let mut chars = format.chars();
    while let Some(character) = chars.next() {
        match character {
            '"' => {
                for inner in chars.by_ref() {
                    if inner == '"' {
                        break;
                    }
                }
            }
            '\\' => {
                chars.next();
            }
            ';' => sections.push(String::new()),
            other => sections.last_mut().expect("one section at least").push(other),
        }
    }
    sections
}

#[cfg(test)]
mod name_union {
    use super::*;

    /// A name may stand for a list of blocks, and SUM of it adds them all.
    #[test]
    fn a_name_of_several_blocks_sums_them_all() {
        let mut wb = Workbook::new();
        wb.add_sheet("Sheet1");
        wb.set_value("Sheet1", "A1", Value::Number(4.0)).unwrap();
        wb.set_value("Sheet1", "D4", Value::Number(4.0)).unwrap();
        wb.define_name("Twin", "Sheet1!$A$1:$B$2,Sheet1!$D$4").unwrap();
        assert_eq!(wb.evaluate("Sheet1", "=SUM(Twin)").unwrap(), Value::Number(8.0));
    }
    #[test]
    fn defined_names_use_relative_mixed_whole_line_and_wrapped_references() {
        let mut wb = Workbook::new();wb.add_sheet("Caller");wb.add_sheet("Other");
        for (cell, value) in [("E1",41.0),("E2",42.0),("F1",51.0),("F2",52.0),("F3",53.0),("G1",61.0),("G3",63.0),("D4",14.0),("B4",40.0),("B6",26.0),("C5",20.0),("E5",45.0)] {
            wb.set_value("Other",cell,Value::Number(value)).unwrap();
        }
        for (name, body) in [("Rel","Other!A1"),("Block","Other!A1:B2"),("Mix","SUM(Other!A1,Other!$B2,Other!C$3,Other!$D$4)"),
            ("Prev","Other!XFD1048576"),("FixedCol","Other!$B:$B"),("RelCol","Other!A:A"),("RelRows","Other!1:2"),
            ("Scale","LAMBDA(x,x+Other!A1)"),("Other!Scale","LAMBDA(x,x+Other!A1)")] {
            wb.define_name(name,body).unwrap();
        }
        for (cell, formula) in [("E1","=Rel"),("E2","=SUM(Block)"),("C5","=Prev"),("E5","=Mix"),("C8","=SUM(FixedCol)"),("C9","=SUM(RelCol)"),("E4","=SUM(RelRows)"),("F1","=Scale(1)"),("G1","=Other!Scale(1)")] {
            wb.set_formula("Caller",cell,formula).unwrap();
        }
        wb.recalculate();
        for (cell, expected) in [("E1",41.0),("E2",147.0),("C5",40.0),("E5",148.0),("C8",66.0),("C9",20.0),("E4",119.0),("F1",52.0),("G1",62.0)] {
            assert_eq!(wb.value("Caller",cell),Value::Number(expected),"{cell}");
        }
        let affected=wb.dependent_formula_ranges(&[("Other".into(),RangeRef::single(CellRef::new(4,0)))]);
        assert!(affected.contains(&("Caller".into(),(4,0))));
        assert!(!affected.contains(&("Caller".into(),(4,1))));
        wb.set_value("Other","E1",Value::Number(99.0)).unwrap();
        wb.recalculate_after(&[("Other".into(),(4,0))]);
        assert_eq!(wb.value("Caller","E1"),Value::Number(99.0));
        assert_eq!(wb.value("Caller","E2"),Value::Number(147.0));
    }

}
