// This Source Code Form is subject to the terms of the Mozilla Public
// License, v. 2.0. If a copy of the MPL was not distributed with this
// file, You can obtain one at https://mozilla.org/MPL/2.0/.

//! Worksheet function library.
//!
//! Deliberately excluded for now: the volatile functions (`NOW`, `TODAY`,
//! `RAND`, `RANDBETWEEN`). They read the wall clock or an RNG, which would make
//! recalculation non-reproducible and therefore impossible to diff against an
//! Excel oracle. They need an injected clock and seed before they can be added.

use crate::datetime;
use crate::reference::RangeRef;
use crate::value::{compare, ExcelError, Value};
use std::cmp::Ordering;

/// A materialised rectangular block of values, row-major.
#[derive(Debug, Clone, PartialEq)]
pub struct RangeData {
    pub width: usize,
    pub height: usize,
    pub cells: Vec<Value>,
}

impl RangeData {
    pub fn at(&self, col: usize, row: usize) -> Value {
        self.cells
            .get(row * self.width + col)
            .cloned()
            .unwrap_or(Value::Blank)
    }

    pub fn from_range(range: &RangeRef, mut read: impl FnMut(u32, u32) -> Value) -> RangeData {
        let width = range.width() as usize;
        let height = range.height() as usize;
        let mut cells = Vec::with_capacity(width * height);
        for (col, row) in range.iter() {
            cells.push(read(col, row));
        }
        RangeData {
            width,
            height,
            cells,
        }
    }
}

/// One evaluated argument: either a scalar or a whole range.
///
/// The distinction matters because `SUM(A1:A3)` must see three cells while
/// `LEN(A1:A3)` must fail; collapsing ranges to scalars too early loses that.
#[derive(Debug, Clone, PartialEq)]
pub enum Arg {
    Value(Value),
    Range(RangeData),
}

impl Arg {
    /// Collapse to a single value for scalar contexts.
    ///
    /// A 1×1 range is transparently a scalar. Anything larger is `#VALUE!`,
    /// which is what Excel produces without dynamic arrays.
    pub fn scalar(&self) -> Value {
        match self {
            Arg::Value(v) => v.clone(),
            Arg::Range(r) if r.cells.len() == 1 => r.cells[0].clone(),
            Arg::Range(_) => Value::Error(ExcelError::Value),
        }
    }

    /// The first value of a block, whatever its size.
    ///
    /// `scalar` refuses a block larger than one cell, which is right where a
    /// block is a mistake. It is wrong for a lookup's needle: asked of Excel,
    /// `SUM(VLOOKUP({"x","y"},A1:B3,2,FALSE))` is 10 -- the first needle's
    /// answer alone, not one answer per needle, which is what MATCH gives.
    /// The two look alike and behave the other way round.
    pub fn first(&self) -> Value {
        match self {
            Arg::Value(v) => v.clone(),
            Arg::Range(r) => r.cells.first().cloned().unwrap_or(Value::Blank),
        }
    }

    /// Every value, flattened. Used by the aggregate functions.
    pub fn flatten(&self) -> Vec<Value> {
        match self {
            Arg::Value(v) => vec![v.clone()],
            Arg::Range(r) => r.cells.clone(),
        }
    }

    fn as_range(&self) -> RangeData {
        match self {
            Arg::Range(r) => r.clone(),
            Arg::Value(v) => RangeData {
                width: 1,
                height: 1,
                cells: vec![v.clone()],
            },
        }
    }
}

/// Does `text` match `pattern`, reading `?` as one character, `*` as any run
/// of them, and `~` in front of either as the character itself?
///
/// Excel matches this way in the exact-match forms of VLOOKUP, HLOOKUP and
/// MATCH, and in every criteria argument. Comparing the pattern as literal
/// text instead means `VLOOKUP(D1 & "*", ...)` — the ordinary way to look
/// something up by its beginning — finds nothing at all.
pub(crate) fn wildcard_match(text: &str, pattern: &str) -> bool {
    let text: Vec<char> = text.to_lowercase().chars().collect();
    let pattern: Vec<char> = pattern.to_lowercase().chars().collect();
    // Walked rather than recursed, remembering the last `*` so a dead end can
    // be backed out of: `a*b` against `aXbY` has to try the second `b` too.
    let (mut at, mut against) = (0usize, 0usize);
    let (mut star, mut after_star) = (None, 0usize);
    while at < text.len() {
        let here = pattern.get(against).copied();
        match here {
            Some('~') if against + 1 < pattern.len() => {
                if pattern[against + 1] == text[at] {
                    at += 1;
                    against += 2;
                    continue;
                }
            }
            Some('?') => {
                at += 1;
                against += 1;
                continue;
            }
            Some('*') => {
                star = Some(against);
                against += 1;
                after_star = at;
                continue;
            }
            Some(one) if one == text[at] => {
                at += 1;
                against += 1;
                continue;
            }
            _ => {}
        }
        match star {
            Some(back) => {
                against = back + 1;
                after_star += 1;
                at = after_star;
            }
            None => return false,
        }
    }
    while pattern.get(against) == Some(&'*') {
        against += 1;
    }
    against == pattern.len()
}

/// The ISO week: weeks start on Monday and week one is the one holding the
/// year's first Thursday, so a date in early January can belong to the year
/// before it.
fn weeknum_iso(serial: i64) -> Result<Value, ExcelError> {
    // The Thursday of this date's week settles which year the week belongs to.
    let weekday = weekday_with_type(serial, 2)?; // Monday = 1
    let thursday = serial - (weekday - 1) + 3;
    let year = datetime::date_from_serial(thursday)?.year;
    let first = datetime::serial_from_date(year, 1, 1)?;
    let first_weekday = weekday_with_type(first, 2)?;
    let first_thursday = first - (first_weekday - 1) + 3;
    Ok(Value::Number(((thursday - first_thursday) / 7 + 1) as f64))
}

/// Does this candidate answer to this key, the way an exact lookup asks?
///
/// Text against text with a `*` or a `?` in it is a pattern; everything else
/// is ordinary equality. A blank never answers, not even to a blank key —
/// Excel reports `#N/A` rather than pairing two empty cells, and without that
/// a lookup of an unfilled cell quietly returns whatever sits beside the first
/// gap in the table.
fn answers_to(candidate: &Value, key: &Value) -> bool {
    if candidate.is_blank() {
        return false;
    }
    if let (Value::Text(held), Value::Text(pattern)) = (candidate, key) {
        if has_wildcards(pattern) {
            return wildcard_match(held, pattern);
        }
        // An exact lookup asks for the same letters, case aside -- not the
        // collation's likeness: measured, MATCH("ss",{..."ß","ss"},0) finds
        // "ss", and MATCH("ｱ",{"ア",...,"ｱ"},0) the half-width one.
        return held.to_lowercase() == pattern.to_lowercase();
    }
    compare(candidate, key) == Ok(Ordering::Equal)
}

/// Whether `pattern` has anything in it that wildcard matching would read.
pub(crate) fn has_wildcards(pattern: &str) -> bool {
    pattern.contains('*') || pattern.contains('?')
}

/// Return the first error among the arguments, so that errors propagate the
/// way Excel propagates them (before the function body runs).
fn first_error(args: &[Arg]) -> Option<ExcelError> {
    args.iter()
        .flat_map(|a| a.flatten())
        .find_map(|v| v.err())
}

/// The first error handed over as an argument in its own right, rather than
/// found among the values of a block.
///
/// The difference is the difference between a cell of a range holding `#N/A`
/// and a range that is not there at all.
fn bare_error(args: &[Arg]) -> Option<ExcelError> {
    args.iter().find_map(|one| match one {
        Arg::Value(held) => held.err(),
        // Anything that came from the sheet is a block, however small, and its
        // contents are values. `CHOOSE(1,A3,A2)` is 30 with an error in A2,
        // and `COUNT(A4)` is 0 rather than that error — each of those knows
        // what to do with one, and neither is missing a block.
        Arg::Range(_) => None,
    })
}

pub(crate) fn num(arg: &Arg) -> Result<f64, ExcelError> {
    arg.scalar().to_number()
}

fn text(arg: &Arg) -> Result<String, ExcelError> {
    arg.scalar().to_text()
}

/// Numbers only, the way `SUM` and `AVERAGE` see a range: text and logicals
/// inside a *range* are skipped, but a directly supplied argument is coerced.
pub(crate) fn numeric_operands(args: &[Arg]) -> Result<Vec<f64>, ExcelError> {
    let mut out = Vec::new();
    for arg in args {
        match arg {
            Arg::Value(v) => match v {
                Value::Error(e) => return Err(*e),
                Value::Blank => {}
                Value::Text(_) | Value::Logical(_) | Value::Number(_) => out.push(v.to_number()?),
            },
            Arg::Range(r) => {
                for v in &r.cells {
                    match v {
                        Value::Error(e) => return Err(*e),
                        Value::Number(n) => out.push(*n),
                        // Text and logicals inside a range are ignored, not coerced.
                        _ => {}
                    }
                }
            }
        }
    }
    Ok(out)
}

/// UTF-16 code units, because that is the unit Excel's text functions count.
///
/// `LEN("あ")` is 1 and `LEN("𠮷")` is 2 in Excel. Counting `char`s would give
/// 1 for both; counting bytes (as the previous prototype did) would give 3 and 4.
fn utf16(s: &str) -> Vec<u16> {
    s.encode_utf16().collect()
}

fn from_utf16(units: &[u16]) -> String {
    String::from_utf16_lossy(units)
}

/// Evaluate a worksheet function. Returns `#NAME?` for anything unimplemented,
/// which is the same thing Excel reports for a function it does not know.
/// One cell of a block, stretching a single row or column across the rest.
///
/// A side one cell wide is read for every column and one cell tall for every
/// row, which is what lets a column of 72 meet a row of 14 and make a block of
/// both. `None` is a cell the block does not reach.
pub(crate) fn reach(block: &RangeData, col: usize, row: usize) -> Option<Value> {
    let col = if block.width == 1 { 0 } else { col };
    let row = if block.height == 1 { 0 } else { row };
    if col >= block.width || row >= block.height {
        return None;
    }
    Some(block.at(col, row))
}

pub(crate) fn block_of(arg: &Arg) -> RangeData {
    arg.as_range()
}

/// Whether every argument this function takes is a single value.
///
/// Named one by one, and deliberately so. The first version of this asked the
/// question the other way round — everything is one-at-a-time unless it is a
/// known aggregate — and `DGET(A$3:C$200, 4, P11:P12)` was quietly applied to
/// each of six hundred cells in turn. It takes three ranges and was in no list
/// of aggregates because nobody had thought of it.
///
/// A name left off this list keeps the behaviour it always had. A name wrongly
/// on it is applied hundreds of times to the wrong things and says nothing
/// about it, so the cost of the two mistakes is not remotely equal.
fn one_at_a_time(name: &str) -> bool {
    matches!(
        name,
        // arithmetic on one number
        "ABS" | "INT" | "MOD" | "POWER" | "SQRT" | "ROUND" | "ROUNDDOWN"
            | "ROUNDUP" | "CEILING" | "FLOOR" | "CEILING.MATH" | "FLOOR.MATH"
            | "EXP" | "LN" | "LOG" | "LOG10" | "SIN" | "COS" | "TAN" | "ASIN"
            | "ACOS" | "ATAN" | "ATAN2" | "SINH" | "COSH" | "TANH" | "ASINH"
            | "ACOSH" | "ATANH" | "SQRTPI"
            | "DEC2HEX" | "DEC2BIN" | "DEC2OCT" | "HEX2DEC" | "BIN2DEC"
            | "OCT2DEC" | "BIN2HEX" | "HEX2BIN" | "BIN2OCT" | "OCT2BIN"
            | "HEX2OCT" | "OCT2HEX"
        // one piece of text
            | "LEN" | "LEFT" | "RIGHT" | "MID" | "LOWER" | "UPPER" | "TRIM"
            | "FIND" | "SEARCH" | "SUBSTITUTE" | "REPLACE" | "REPT" | "EXACT"
            | "CHAR" | "CODE" | "UNICODE" | "TEXT" | "VALUE" | "PROPER" | "T"
        // one date
            // (EDATE and EOMONTH take arrays so, but a block of cells is
            // #VALUE! -- the engine refuses that before it gets here)
            | "DATE" | "DATEDIF" | "DAY" | "DAYS" | "EDATE" | "EOMONTH"
            | "HOUR" | "MINUTE" | "MONTH" | "SECOND" | "TIME" | "WEEKDAY"
            | "YEAR"
        // one thing, tested
            | "ISBLANK" | "ISERR" | "ISERROR" | "ISLOGICAL" | "ISNA"
            | "ISNUMBER" | "ISTEXT" | "NOT"
        // and the ones that pick between two answers, which is what makes
        // `IF(A1:A10>5, 1, 0)` a column of ten rather than one #VALUE!
            | "IF" | "IFERROR" | "IFNA"
    )
}

/// Whether argument `at` is one this function takes a cell at a time.
///
/// `one_at_a_time` says a function works on single values; that is enough for
/// `LEN` or `IF`, where EVERY argument is a single value. It is not enough for
/// a lookup: `MATCH($S$1:$S$20,$A2,0)` hands a block as the thing to look FOR
/// and a single value as the place to look IN, and Excel answers a block —
/// one match per thing looked for. `COUNTIF(range,{"LONG","SHORT"})` is the
/// same shape the other way round: the range is read whole and the criterion
/// is taken one at a time, which is why `SUMPRODUCT` of it is 3 and not 2.
///
/// So the question is per argument. A function absent from both lists takes
/// everything whole.
fn taken_one_at_a_time(name: &str, at: usize) -> bool {
    if one_at_a_time(name) {
        return true;
    }
    match name {
        // The needle, not the haystack. MATCH does this and the LOOKUPs do
        // NOT: asked of Excel, `SUM(MATCH({"y","x"},A1:A3,0))` is 3 — one
        // answer per needle — while `SUM(VLOOKUP({"x","y"},A1:B3,2,FALSE))` is
        // 10, the first needle's answer alone. They look alike and they are
        // not, so only the one that was measured is here.
        "MATCH" => at == 0,
        // range, criteria[, sum range] — the criteria only.
        "COUNTIF" | "SUMIF" | "AVERAGEIF" => at == 1,
        // range, criteria, range, criteria, … — every second one.
        "COUNTIFS" => at % 2 == 1,
        // sum range, then the pairs, so the criteria are the even ones.
        "SUMIFS" | "AVERAGEIFS" => at >= 2 && at.is_multiple_of(2),
        _ => false,
    }
}

/// Call `name`, applying it a cell at a time when it has been handed a block
/// and only knows what to do with one value.
pub fn call_arg(name: &str, args: &[Arg]) -> Arg {
    let name = plain(name);
    // An INDEX missing one of its two indexes means a whole line rather than
    // one cell, and a whole line is an array — which `call` has no way to
    // return.
    if name == "INDEX" {
        if let Some(line) = a_whole_line(args) {
            return line;
        }
        // Indexes handed as an array pick one cell each: measured,
        // Evaluate("INDEX({1;2},{1;2})") is an array.
        if let [table, Arg::Range(rows)] = args {
            if rows.cells.len() > 1 {
                let cells = rows
                    .cells
                    .iter()
                    .map(|row| match call_arg("INDEX", &[table.clone(), Arg::Value(row.clone()), Arg::Value(Value::Number(1.0))]) {
                        Arg::Value(value) => value,
                        Arg::Range(block) => block.cells.first().cloned().unwrap_or(Value::Blank),
                    })
                    .collect();
                return Arg::Range(RangeData { width: rows.width, height: rows.height, cells });
            }
        }
        // One index into an array of several rows is that whole row, a
        // column's one cell included: measured, `SUM(INDEX({1,2;3,4},2))`
        // is 7 and `Evaluate("INDEX({1;2;3},1)")` an array of one. (A
        // reference of two dimensions is #REF! instead, which the engine
        // settles before it gets here.)
        if let [table, row] = args {
            let table = table.as_range();
            if table.height > 1 {
                return match row.scalar().to_number() {
                    Ok(row) if row >= 1.0 && (row as usize) <= table.height => {
                        let at = row as usize - 1;
                        Arg::Range(RangeData {
                            width: table.width,
                            height: 1,
                            cells: (0..table.width).map(|col| table.at(col, at)).collect(),
                        })
                    }
                    Ok(_) => Arg::Value(Value::Error(ExcelError::Ref)),
                    Err(why) => Arg::Value(Value::Error(why)),
                };
            }
        }
    }
    // The functions that cut, join and reshape blocks.
    if matches!(
        name,
        "SEQUENCE" | "TAKE" | "DROP" | "CHOOSEROWS" | "CHOOSECOLS" | "VSTACK" | "HSTACK"
            | "TOCOL" | "TOROW" | "WRAPROWS" | "WRAPCOLS" | "EXPAND" | "TEXTSPLIT"
    ) {
        return match reshaped(name, args) {
            Ok(block) => Arg::Range(block),
            Err(why) => Arg::Value(Value::Error(why)),
        };
    }
    // TREND answers along a straight line fitted to the known points, one
    // answer for each new x, in the new x's shape.
    if name == "TREND" {
        if let Some(answer) = crate::functions_more::trend_many(args) {
            return match answer {
                Ok(block) => Arg::Range(block),
                Err(why) => Arg::Value(Value::Error(why)),
            };
        }
        return match trend(args) {
            Ok(block) => Arg::Range(block),
            Err(why) => Arg::Value(Value::Error(why)),
        };
    }
    // MODE.MULT: every value turning up most often (twice at least), in the
    // order each first turns up, down a column. Measured: over 3,1,1,3,2,5
    // it spills 3 then 1.
    if name == "MODE.MULT" {
        let held = match numeric_operands(args) {
            Ok(held) => held,
            Err(why) => return Arg::Value(Value::Error(why)),
        };
        let times = |one: &f64| held.iter().filter(|other| *other == one).count();
        let most = held.iter().map(times).max().unwrap_or(0);
        if most < 2 {
            return Arg::Value(Value::Error(ExcelError::NA));
        }
        let mut cells: Vec<Value> = Vec::new();
        for one in &held {
            if times(one) == most && !cells.contains(&Value::Number(*one)) {
                cells.push(Value::Number(*one));
            }
        }
        return Arg::Range(RangeData { width: 1, height: cells.len(), cells });
    }
    if let Some(block) = crate::functions_more::call_block(name, args) {
        return block;
    }
    // FREQUENCY counts into the bins and one more, down a column.
    if name == "FREQUENCY" {
        return match frequency(args) {
            Ok(block) => Arg::Range(block),
            Err(why) => Arg::Value(Value::Error(why)),
        };
    }
    // The ones that hand back a block rather than a value.
    if matches!(name, "UNIQUE" | "SORT" | "FILTER" | "SORTBY") {
        return match a_block_of_rows(name, args) {
            Ok(block) => block,
            Err(why) => Arg::Value(Value::Error(why)),
        };
    }
    // MMULT: the matrix product, rows of the first by columns of the second.
    // Measured: `@MMULT(A1:A3,1)` over 1,2,3 is 1 -- a lone number is a
    // block of one. A block that is not all numbers, or shapes that do not
    // meet, is #VALUE!.
    if name == "MMULT" {
        let [left, right] = args else {
            return Arg::Value(Value::Error(ExcelError::Value));
        };
        let (a, b) = (left.as_range(), right.as_range());
        if a.width != b.height || a.cells.is_empty() || b.cells.is_empty() {
            return Arg::Value(Value::Error(ExcelError::Value));
        }
        let number = |value: Value| match value {
            Value::Number(n) => Some(n),
            _ => None,
        };
        let mut cells = Vec::with_capacity(a.height * b.width);
        for row in 0..a.height {
            for col in 0..b.width {
                let mut sum = 0.0;
                for k in 0..a.width {
                    match (number(a.at(k, row)), number(b.at(col, k))) {
                        (Some(x), Some(y)) => sum += x * y,
                        _ => return Arg::Value(Value::Error(ExcelError::Value)),
                    }
                }
                cells.push(Value::Number(sum));
            }
        }
        return Arg::Range(RangeData { width: b.width, height: a.height, cells });
    }
    // TRANSPOSE turns its one block on its side, which is the whole of it:
    // `{=TRANSPOSE(A1:A2)}` across F1:G1 is 1 and 2.
    if name == "TRANSPOSE" {
        return match args {
            [block] => Arg::Range(on_its_side(&block.as_range())),
            _ => Arg::Value(Value::Error(ExcelError::Value)),
        };
    }
    // The shape of the answer comes from the arguments taken a cell at a time;
    // the ones read whole say nothing about it.
    let spread: Vec<bool> = (0..args.len())
        .map(|at| taken_one_at_a_time(name, at))
        .collect();
    if spread.iter().any(|one| *one) {
        let sized = |pick: fn(&RangeData) -> usize| {
            args.iter()
                .zip(&spread)
                .filter(|(_, taken)| **taken)
                .map(|(one, _)| pick(&block_of(one)))
                .max()
                .unwrap_or(1)
        };
        let width = sized(|block| block.width);
        let height = sized(|block| block.height);
        if width * height > 1 {
            let blocks: Vec<RangeData> = args.iter().map(block_of).collect();
            let mut cells = Vec::with_capacity(width * height);
            for row in 0..height {
                for col in 0..width {
                    let picked: Vec<Arg> = args
                        .iter()
                        .zip(&blocks)
                        .zip(&spread)
                        .map(|((whole, block), taken)| {
                            if *taken {
                                Arg::Value(
                                    reach(block, col, row).unwrap_or(Value::Error(ExcelError::NA)),
                                )
                            } else {
                                whole.clone()
                            }
                        })
                        .collect();
                    cells.push(call(name, &picked));
                }
            }
            return Arg::Range(RangeData {
                width,
                height,
                cells,
            });
        }
    }
    Arg::Value(call(name, args))
}

/// The name without the prefixes a file writes and Excel does not show.
///
/// A function newer than the format's own version is stored as `_xlfn.NAME`,
/// and one that only works on a worksheet as `_xlfn._xlws.NAME`. They are a
/// note to older readers, not part of the name.
pub(crate) fn plain(name: &str) -> &str {
    // The parser upper-cases every function name, so the prefix arrives as
    // `_XLFN.` however the file spelled it. Stripping only the lower-case form
    // matched nothing at all, silently.
    let name = strip_either(name, "_xlfn.");
    strip_either(name, "_xlws.")
}

fn strip_either<'a>(name: &'a str, prefix: &str) -> &'a str {
    if name.len() >= prefix.len() && name[..prefix.len()].eq_ignore_ascii_case(prefix) {
        &name[prefix.len()..]
    } else {
        name
    }
}

/// Every function this build knows the name of: those the library answers
/// and those the workbook works out itself. Kept in step with the match
/// arms by `every_function_the_library_answers_is_known`.
const KNOWN_FUNCTIONS: &[&str] = &[
    "ABS", "ACCRINT", "ACCRINTM", "ACOS", "ACOSH", "ACOT", "ACOTH", "ADDRESS", "AGGREGATE",
    "AMORDEGRC", "AMORLINC", "AND", "ARABIC", "AREAS", "ARRAYTOTEXT", "ASC", "ASIN", "ASINH",
    "ATAN", "ATAN2", "ATANH", "AVEDEV", "AVERAGE", "AVERAGEA", "AVERAGEIF", "AVERAGEIFS",
    "BAHTTEXT", "BASE", "BESSELI", "BESSELJ", "BESSELK", "BESSELY", "BETA.DIST", "BETA.INV",
    "BETADIST", "BETAINV", "BIN2DEC", "BIN2HEX", "BIN2OCT", "BINOM.DIST", "BINOM.DIST.RANGE",
    "BINOM.INV", "BINOMDIST", "BITAND", "BITLSHIFT", "BITOR", "BITRSHIFT", "BITXOR", "BYCOL",
    "BYROW", "CEILING", "CEILING.MATH", "CEILING.PRECISE", "CELL", "CHAR", "CHIDIST", "CHIINV",
    "CHISQ.DIST", "CHISQ.DIST.RT", "CHISQ.INV", "CHISQ.INV.RT", "CHISQ.TEST", "CHITEST", "CHOOSE",
    "CHOOSECOLS", "CHOOSEROWS", "CLEAN", "CODE", "COLUMN", "COLUMNS", "COMBIN", "COMBINA",
    "COMPLEX", "CONCAT", "CONCATENATE", "CONFIDENCE", "CONFIDENCE.NORM", "CONFIDENCE.T", "CONVERT",
    "CORREL", "COS", "COSH", "COT", "COTH", "COUNT", "COUNTA", "COUNTBLANK", "COUNTIF", "COUNTIFS",
    "COUPDAYBS", "COUPDAYS", "COUPDAYSNC", "COUPNCD", "COUPNUM", "COUPPCD", "COVAR", "COVARIANCE.P",
    "COVARIANCE.S", "CRITBINOM", "CSC", "CSCH", "CUMIPMT", "CUMPRINC", "D", "DATE", "DATEDIF",
    "DATEVALUE", "DAVERAGE", "DAY", "DAYS", "DAYS360", "DB", "DBCS", "DCOUNT", "DCOUNTA", "DDB",
    "DEC2BIN", "DEC2HEX", "DEC2OCT", "DECIMAL", "DEGREES", "DELTA", "DEVSQ", "DGET", "DISC", "DMAX",
    "DMIN", "DOLLAR", "DOLLARDE", "DOLLARFR", "DPRODUCT", "DROP", "DSTDEV", "DSTDEVP", "DSUM",
    "DURATION", "DVAR", "DVARP", "ECMA.CEILING", "EDATE", "EFFECT", "ENCODEURL", "EOMONTH", "ERF",
    "ERF.PRECISE", "ERFC", "ERFC.PRECISE", "ERROR.TYPE", "EVEN", "EXACT", "EXP", "EXPAND",
    "EXPON.DIST", "EXPONDIST", "F.DIST", "F.DIST.RT", "F.INV", "F.INV.RT", "F.TEST", "FACT",
    "FACTDOUBLE", "FALSE", "FDIST", "FIND", "FINDB", "FINV", "FISHER", "FISHERINV", "FIXED",
    "FLOOR", "FLOOR.MATH", "FLOOR.PRECISE", "FORECAST", "FORECAST.LINEAR", "FORMULATEXT", "FTEST",
    "FV", "FVSCHEDULE", "GAMMA", "GAMMA.DIST", "GAMMA.INV", "GAMMADIST", "GAMMAINV", "GAMMALN",
    "GAMMALN.PRECISE", "GAUSS", "GCD", "GEOMEAN", "GESTEP", "GROUPBY", "GROWTH", "HARMEAN",
    "HEX2BIN", "HEX2DEC", "HEX2OCT", "HLOOKUP", "HOUR", "HSTACK", "HYPERLINK", "HYPGEOM.DIST",
    "HYPGEOMDIST", "IF", "IFERROR", "IFNA", "IFS", "IMABS", "IMAGINARY", "IMARGUMENT",
    "IMCONJUGATE", "IMCOS", "IMCOSH", "IMCOT", "IMCSC", "IMCSCH", "IMDIV", "IMEXP", "IMLN",
    "IMLOG10", "IMLOG2", "IMPOWER", "IMPRODUCT", "IMREAL", "IMSEC", "IMSECH", "IMSIN", "IMSINH",
    "IMSQRT", "IMSUB", "IMSUM", "IMTAN", "INDEX", "INDIRECT", "INFO", "INT", "INTERCEPT", "INTRATE",
    "IPMT", "IRR", "ISBLANK", "ISERR", "ISERROR", "ISEVEN", "ISFORMULA", "ISLOGICAL", "ISNA",
    "ISNONTEXT", "ISNUMBER", "ISO.CEILING", "ISODD", "ISOMITTED", "ISOWEEKNUM", "ISPMT", "ISREF",
    "ISTEXT", "KURT", "LAMBDA", "LARGE", "LCM", "LEFT", "LEFTB", "LEN", "LENB", "LET", "LINEST",
    "LN", "LOG", "LOG10", "LOGEST", "LOGINV", "LOGNORM.DIST", "LOGNORM.INV", "LOGNORMDIST",
    "LOOKUP", "LOWER", "M", "MAKEARRAY", "MAP", "MATCH", "MAX", "MAXA", "MAXIFS", "MD", "MDETERM",
    "MDURATION", "MEDIAN", "MID", "MIDB", "MIN", "MINA", "MINIFS", "MINUTE", "MINVERSE", "MIRR",
    "MMULT", "MOD", "MODE", "MODE.MULT", "MODE.SNGL", "MONTH", "MROUND", "MULTINOMIAL", "MUNIT",
    "N", "NA", "NEGBINOM.DIST", "NEGBINOMDIST", "NETWORKDAYS", "NETWORKDAYS.INTL", "NOMINAL",
    "NORM.DIST", "NORM.INV", "NORM.S.DIST", "NORM.S.INV", "NORMDIST", "NORMINV", "NORMSDIST",
    "NORMSINV", "NOT", "NOW", "NPER", "NPV", "NUMBERVALUE", "OCT2BIN", "OCT2DEC", "OCT2HEX", "ODD",
    "ODDFPRICE", "ODDFYIELD", "ODDLPRICE", "ODDLYIELD", "OFFSET", "OR", "PDURATION", "PEARSON",
    "PERCENTILE", "PERCENTILE.EXC", "PERCENTILE.INC", "PERCENTOF", "PERCENTRANK", "PERCENTRANK.EXC",
    "PERCENTRANK.INC", "PERMUT", "PERMUTATIONA", "PHI", "PHONETIC", "PI", "PIVOTBY", "PMT",
    "POISSON", "POISSON.DIST", "POWER", "PPMT", "PRICE", "PRICEDISC", "PRICEMAT", "PROB", "PRODUCT",
    "PROPER", "PV", "QUARTILE", "QUARTILE.EXC", "QUARTILE.INC", "QUOTIENT", "RADIANS", "RAND",
    "RANDARRAY", "RANDBETWEEN", "RANK", "RANK.AVG", "RANK.EQ", "RATE", "RECEIVED", "REDUCE",
    "REGEXEXTRACT", "REGEXREPLACE", "REGEXTEST", "REPLACE", "REPLACEB", "REPT", "RIGHT", "RIGHTB",
    "ROMAN", "ROUND", "ROUNDDOWN", "ROUNDUP", "ROW", "ROWS", "RRI", "RSQ", "SCAN", "SEARCH",
    "SEARCHB", "SEC", "SECH", "SECOND", "SEQUENCE", "SERIESSUM", "SHEET", "SHEETS", "SIGN", "SIN",
    "SINH", "SKEW", "SKEW.P", "SLN", "SLOPE", "SMALL", "SORT", "SORTBY", "SQRT", "SQRTPI",
    "STANDARDIZE", "STDEV", "STDEV.P", "STDEV.S", "STDEVA", "STDEVP", "STDEVPA", "STEYX",
    "SUBSTITUTE", "SUBTOTAL", "SUM", "SUMIF", "SUMIFS", "SUMPRODUCT", "SUMSQ", "SUMX2MY2",
    "SUMX2PY2", "SUMXMY2", "SWITCH", "SYD", "T", "T.DIST", "T.DIST.2T", "T.DIST.RT", "T.INV",
    "T.INV.2T", "T.TEST", "TAKE", "TAN", "TANH", "TBILLEQ", "TBILLPRICE", "TBILLYIELD", "TDIST",
    "TEXT", "TEXTAFTER", "TEXTBEFORE", "TEXTJOIN", "TEXTSPLIT", "TIME", "TIMEVALUE", "TINV",
    "TOCOL", "TODAY", "TOROW", "TRIM", "TRIMMEAN", "TRIMRANGE", "TRUE", "TRUNC", "TTEST", "TYPE",
    "UNICHAR", "UNICODE", "UNIQUE", "UPPER", "VALUE", "VALUETOTEXT", "VAR", "VAR.P", "VAR.S",
    "VARA", "VARP", "VARPA", "VDB", "VLOOKUP", "VSTACK", "WEEKDAY", "WEEKNUM", "WEIBULL",
    "WEIBULL.DIST", "WORKDAY", "WORKDAY.INTL", "WRAPCOLS", "WRAPROWS", "XIRR", "XLOOKUP", "XMATCH",
    "XNPV", "XOR", "Y", "YD", "YEAR", "YEARFRAC", "YIELD", "YIELDDISC", "YIELDMAT", "YM", "Z.TEST",
    "ZTEST",
];

/// Whether `name` is a function this build knows.
pub fn is_known_function(name: &str) -> bool {
    let upper = name.to_ascii_uppercase();
    KNOWN_FUNCTIONS.binary_search(&upper.as_str()).is_ok()
}

pub fn call(name: &str, args: &[Arg]) -> Value {
    let name = plain(name);
    match dispatch(name, args) {
        Ok(v) => v,
        Err(e) => Value::Error(e),
    }
}

fn dispatch(name: &str, args: &[Arg]) -> Result<Value, ExcelError> {
    // Some functions have to SEE an error rather than pass it on. For IFERROR
    // and IFNA that is the whole point of them; for the IS* family it is too —
    // `ISNUMBER(#VALUE!)` is FALSE in Excel, not `#VALUE!`, and a function that
    // cannot say "no, that is not a number" about an error is no use for the
    // one thing it exists to do. `ISNUMBER(SEARCH(x, range))` is the commonest
    // way anyone asks "does this text appear in that list", and it only works
    // because the misses come back as errors and ISNUMBER calls them false.
    let error_transparent = matches!(
        name,
        "IFERROR" | "IFNA" | "IF" | "IFS" | "SWITCH" | "TYPE" | "ERROR.TYPE"
            | "ISERROR" | "ISNA" | "ISERR" | "ISNUMBER" | "ISTEXT"
            | "ISBLANK" | "ISLOGICAL" | "ISREF"
        // The conditional sums look at one row at a time, so an error in a
        // range is a fact about that row and not about the answer. Each of
        // them settles for itself what to do with one.
            | "SUMIF" | "COUNTIF" | "AVERAGEIF"
            | "SUMIFS" | "COUNTIFS" | "AVERAGEIFS"
        // AGGREGATE's whole second argument is about what to do with them.
            | "AGGREGATE"
        // These three never mind an error, wherever it came from: COUNT is
        // asking how many NUMBERS there are and an error is not one, COUNTA
        // how many cells are not empty and an error fills a cell, and CHOOSE
        // only ever looks at the one it is told to. Excel: `COUNT(#REF!)` is
        // 0, `COUNTA(#REF!)` is 1, `CHOOSE(1,30,#REF!)` is 30.
            | "COUNT" | "COUNTA" | "CHOOSE"
        // COUNTBLANK looks only for the empty: measured, an #NAME? among
        // the cells is simply not one.
            | "COUNTBLANK"
        // And these write an error as its word: measured,
        // `ARRAYTOTEXT(EXPAND(A1:A2,3))` is "3, 1, #N/A".
            | "VALUETOTEXT" | "ARRAYTOTEXT"
    );
    // And some mind only the errors handed to them DIRECTLY.
    //
    // These pick one thing out of a block, search it, or count it, so an error
    // among the other values is nothing to do with the answer — and where it
    // IS the answer, as `INDEX(A1:A4,2)` over an error at 2, it comes back on
    // its own account.
    //
    // An error handed over WHERE THE BLOCK SHOULD BE is a different thing
    // entirely. `INDEX(#REF!,MATCH(x,#REF!,0))` is what Excel writes into a
    // formula whose external workbook has gone, and it answers `#REF!`: there
    // is no block to pick from. Ignoring that gave `#N/A` — MATCH searching a
    // nothing and finding nothing — for 185 cells of one workbook. `ROWS` and
    // `COLUMNS` are here for that case alone: `ROWS(#REF!)` is `#REF!`, and
    // there is nothing else in a range for them to mind.
    let minds_only_bare_errors = matches!(
        name,
        "INDEX" | "MATCH" | "VLOOKUP" | "HLOOKUP" | "XLOOKUP" | "ROWS" | "COLUMNS"
    );
    if !error_transparent {
        let found = if minds_only_bare_errors {
            bare_error(args)
        } else {
            first_error(args)
        };
        if let Some(e) = found {
            return Err(e);
        }
    }

    match name {
        // ---- aggregates -------------------------------------------------
        "SUM" => Ok(Value::Number(numeric_operands(args)?.iter().sum())),
        "PRODUCT" => Ok(Value::Number(
            numeric_operands(args)?.iter().product::<f64>(),
        )),
        "AVERAGE" => {
            let v = numeric_operands(args)?;
            if v.is_empty() {
                return Err(ExcelError::DivZero);
            }
            Ok(Value::Number(v.iter().sum::<f64>() / v.len() as f64))
        }
        // Like AVERAGE, but text counts as zero and logicals as 0/1, in a
        // range as well as an argument; only truly empty cells are skipped.
        "AVERAGEA" => {
            let mut sum = 0.0;
            let mut count = 0u64;
            for value in args.iter().flat_map(|a| a.flatten()) {
                match value {
                    Value::Error(e) => return Err(e),
                    Value::Number(n) => {
                        sum += n;
                        count += 1;
                    }
                    Value::Logical(b) => {
                        sum += f64::from(b);
                        count += 1;
                    }
                    Value::Text(_) => count += 1,
                    Value::Blank => {}
                }
            }
            if count == 0 {
                return Err(ExcelError::DivZero);
            }
            Ok(Value::Number(sum / count as f64))
        }
        // The nth root of the product; every number must be positive.
        "GEOMEAN" => {
            let numbers = numeric_operands(args)?;
            if numbers.is_empty() {
                return Err(ExcelError::Num);
            }
            let mut product = 1.0;
            for n in &numbers {
                if *n <= 0.0 {
                    return Err(ExcelError::Num);
                }
                product *= n;
            }
            Ok(Value::Number(product.powf(1.0 / numbers.len() as f64)))
        }
        // MIN/MAX over nothing is 0 in Excel, not an error.
        "MIN" => {
            let m = numeric_operands(args)?
                .into_iter()
                .fold(f64::INFINITY, f64::min);
            Ok(Value::Number(if m.is_infinite() { 0.0 } else { m }))
        }
        "MAX" => {
            let m = numeric_operands(args)?
                .into_iter()
                .fold(f64::NEG_INFINITY, f64::max);
            Ok(Value::Number(if m.is_infinite() { 0.0 } else { m }))
        }
        // How many numbers there are. An error is not one, so a range holding
        // one is counted as though it were not there — `numeric_operands`
        // handed the error back instead of counting.
        //
        // Excel asks a different question of an argument given DIRECTLY than
        // of a value found inside a range. Measured:
        //
        //     COUNT(A1:A5) 2   over 1, TRUE, 2, #N/A, "text"
        //     COUNT(A2)    0   a logical in a reference is not a number
        //     COUNT(TRUE)  1   but written out it counts
        //     COUNT("2")   1   and so does text that reads as one
        //     COUNT(1,"x") 1   where text that does not, does not
        //     COUNT(NA())  0   and an error never does
        "COUNT" => Ok(Value::Number(
            args.iter()
                .map(|one| match one {
                    // Written out: anything that reads as a number.
                    Arg::Value(held) => usize::from(held.to_number().is_ok()),
                    // Found in a range: only what IS a number.
                    Arg::Range(block) => block
                        .cells
                        .iter()
                        .filter(|held| matches!(held, Value::Number(_)))
                        .count(),
                })
                .sum::<usize>() as f64,
        )),
        "COUNTA" => Ok(Value::Number(
            args.iter()
                .flat_map(|a| a.flatten())
                .filter(|v| !v.is_blank())
                .count() as f64,
        )),
        "COUNTBLANK" => Ok(Value::Number(
            args.iter()
                .flat_map(|a| a.flatten())
                // Empty text counts as blank: measured, a cell holding
                // ="" is counted.
                .filter(|v| v.is_blank() || matches!(v, Value::Text(text) if text.is_empty()))
                .count() as f64,
        )),

        // ---- arithmetic --------------------------------------------------
        "ABS" => Ok(Value::Number(one(args)?.abs())),
        // Logarithms, powers of e and the circle, with Excel's errors: a
        // logarithm of nought or less is #NUM!, and of base 1 #DIV/0!.
        "EXP" => fin(one(args)?.exp()),
        "LN" => {
            let n = one(args)?;
            if n <= 0.0 {
                return Err(ExcelError::Num);
            }
            Ok(Value::Number(n.ln()))
        }
        "LOG10" => {
            let n = one(args)?;
            if n <= 0.0 {
                return Err(ExcelError::Num);
            }
            Ok(Value::Number(n.log10()))
        }
        "LOG" => {
            let n = one(args)?;
            let base = match args.get(1) {
                Some(a) => num(a)?,
                None => 10.0,
            };
            if n <= 0.0 || base <= 0.0 {
                return Err(ExcelError::Num);
            }
            if base == 1.0 {
                return Err(ExcelError::DivZero);
            }
            Ok(Value::Number(if base == 10.0 { n.log10() } else { n.ln() / base.ln() }))
        }
        "PI" => Ok(Value::Number(std::f64::consts::PI)),
        "SQRTPI" => {
            let n = one(args)?;
            if n < 0.0 {
                return Err(ExcelError::Num);
            }
            Ok(Value::Number((n * std::f64::consts::PI).sqrt()))
        }
        "SIN" => fin(one(args)?.sin()),
        "COS" => fin(one(args)?.cos()),
        "TAN" => fin(one(args)?.tan()),
        "SINH" => fin(one(args)?.sinh()),
        "COSH" => fin(one(args)?.cosh()),
        "TANH" => fin(one(args)?.tanh()),
        "ATAN" => fin(one(args)?.atan()),
        "ASINH" => fin(one(args)?.asinh()),
        "ASIN" | "ACOS" => {
            let n = one(args)?;
            if !(-1.0..=1.0).contains(&n) {
                return Err(ExcelError::Num);
            }
            Ok(Value::Number(if name == "ASIN" { n.asin() } else { n.acos() }))
        }
        "ACOSH" => {
            let n = one(args)?;
            if n < 1.0 {
                return Err(ExcelError::Num);
            }
            Ok(Value::Number(n.acosh()))
        }
        "ATANH" => {
            let n = one(args)?;
            if n <= -1.0 || n >= 1.0 {
                return Err(ExcelError::Num);
            }
            Ok(Value::Number(n.atanh()))
        }
        // ATAN2 takes x first, the other way round from most languages.
        "ATAN2" => {
            expect(args, 2)?;
            let (x, y) = (num(&args[0])?, num(&args[1])?);
            if x == 0.0 && y == 0.0 {
                return Err(ExcelError::DivZero);
            }
            Ok(Value::Number(y.atan2(x)))
        }
        "DEC2HEX" | "DEC2BIN" | "DEC2OCT" => {
            let radix = match name { "DEC2HEX" => 16, "DEC2BIN" => 2, _ => 8 };
            let n = num(one_arg(args)?)?.trunc() as i64;
            to_base(n, radix, args.get(1))
        }
        "HEX2DEC" | "BIN2DEC" | "OCT2DEC" => {
            let radix = match name { "HEX2DEC" => 16, "BIN2DEC" => 2, _ => 8 };
            Ok(Value::Number(from_base(&text(one_arg(args)?)?, radix)? as f64))
        }
        "BIN2HEX" | "BIN2OCT" | "HEX2BIN" | "HEX2OCT" | "OCT2BIN" | "OCT2HEX" => {
            let from = match &name[..3] { "BIN" => 2, "HEX" => 16, _ => 8 };
            let to = match &name[4..] { "BIN" => 2, "HEX" => 16, _ => 8 };
            let n = from_base(&text(one_arg(args)?)?, from)?;
            to_base(n, to, args.get(1))
        }
        // Paired data: pairs where either side is not a number are passed
        // over, and the two ranges must be the same size.
        "CORREL" | "PEARSON" | "RSQ" | "SLOPE" | "INTERCEPT" | "COVAR" | "COVARIANCE.P"
        | "COVARIANCE.S" | "STEYX" => {
            expect(args, 2)?;
            let fit = Fit::of(&args[0], &args[1])?;
            match name {
                "CORREL" | "PEARSON" => fit.correl().map(Value::Number),
                "RSQ" => fit.correl().map(|r| Value::Number(r * r)),
                "SLOPE" => fit.slope().map(Value::Number),
                "INTERCEPT" => fit.slope().map(|b| Value::Number(fit.mean_y - b * fit.mean_x)),
                "COVAR" | "COVARIANCE.P" => Ok(Value::Number(fit.sxy / fit.n)),
                "COVARIANCE.S" => {
                    if fit.n < 2.0 {
                        return Err(ExcelError::DivZero);
                    }
                    Ok(Value::Number(fit.sxy / (fit.n - 1.0)))
                }
                _ => {
                    if fit.n < 3.0 || fit.sxx == 0.0 {
                        return Err(ExcelError::DivZero);
                    }
                    Ok(Value::Number(((fit.syy - fit.sxy * fit.sxy / fit.sxx) / (fit.n - 2.0)).sqrt()))
                }
            }
        }
        "FORECAST" | "FORECAST.LINEAR" => {
            expect(args, 3)?;
            let x = num(&args[0])?;
            let fit = Fit::of(&args[1], &args[2])?;
            let slope = fit.slope()?;
            Ok(Value::Number(fit.mean_y + slope * (x - fit.mean_x)))
        }
        "HARMEAN" => {
            let numbers = numeric_operands(args)?;
            if numbers.is_empty() || numbers.iter().any(|n| *n <= 0.0) {
                return Err(ExcelError::Num);
            }
            // One over the mean of the reciprocals, in that order: measured,
            // HARMEAN(1,3,2,5,4,6) is 2.44897959183674 (n / sum gives ...673).
            let mean = numbers.iter().map(|n| 1.0 / n).sum::<f64>() / numbers.len() as f64;
            Ok(Value::Number(1.0 / mean))
        }
        "SKEW" | "KURT" => {
            let numbers = numeric_operands(args)?;
            let n = numbers.len() as f64;
            let least = if name == "SKEW" { 3.0 } else { 4.0 };
            if n < least {
                return Err(ExcelError::DivZero);
            }
            let mean = numbers.iter().sum::<f64>() / n;
            let sd = (numbers.iter().map(|x| (x - mean).powi(2)).sum::<f64>() / (n - 1.0)).sqrt();
            if sd == 0.0 {
                return Err(ExcelError::DivZero);
            }
            Ok(Value::Number(if name == "SKEW" {
                n / ((n - 1.0) * (n - 2.0)) * numbers.iter().map(|x| ((x - mean) / sd).powi(3)).sum::<f64>()
            } else {
                n * (n + 1.0) / ((n - 1.0) * (n - 2.0) * (n - 3.0))
                    * numbers.iter().map(|x| ((x - mean) / sd).powi(4)).sum::<f64>()
                    - 3.0 * (n - 1.0).powi(2) / ((n - 2.0) * (n - 3.0))
            }))
        }
        "SQRT" => {
            let n = one(args)?;
            if n < 0.0 {
                return Err(ExcelError::Num);
            }
            Ok(Value::Number(n.sqrt()))
        }
        "POWER" => {
            expect(args, 2)?;
            let answer = excel_power(num(&args[0])?, num(&args[1])?);
            if answer.is_nan() {
                return Err(ExcelError::Num);
            }
            Ok(Value::Number(answer))
        }
        // Excel's INT floors toward negative infinity: INT(-1.5) is -2.
        "INT" => Ok(Value::Number(one(args)?.floor())),
        // Excel's MOD takes the sign of the divisor: MOD(-3,2) is 1, where
        // Rust's `%` would give -1.
        "MOD" => {
            expect(args, 2)?;
            let (n, d) = (num(&args[0])?, num(&args[1])?);
            if d == 0.0 {
                return Err(ExcelError::DivZero);
            }
            Ok(Value::Number(n - d * (n / d).floor()))
        }
        // Rounds to the nearest multiple, halves away from zero. Excel: the
        // number and the multiple must share a sign, else #NUM!; a zero
        // multiple gives zero. MRound(17,5)=15, MRound(-17,-5)=-15,
        // MRound(2.5,1)=3.
        "MROUND" => {
            expect(args, 2)?;
            let (n, m) = (num(&args[0])?, num(&args[1])?);
            if m == 0.0 {
                return Ok(Value::Number(0.0));
            }
            if n != 0.0 && n.signum() != m.signum() {
                return Err(ExcelError::Num);
            }
            Ok(Value::Number((n / m).round() * m))
        }
        // The integer part of a division, truncated toward zero:
        // Quotient(17,5)=3, Quotient(-17,5)=-3. A zero divisor is #DIV/0!.
        "QUOTIENT" => {
            expect(args, 2)?;
            let (n, d) = (num(&args[0])?, num(&args[1])?);
            if d == 0.0 {
                return Err(ExcelError::DivZero);
            }
            Ok(Value::Number((n / d).trunc()))
        }
        // Truncate toward zero at a number of decimal places (0 by default):
        // Trunc(3.78)=3, Trunc(-3.78,1)=-3.7.
        "TRUNC" => {
            let n = num(args.first().ok_or(ExcelError::Value)?)?;
            let digits = match args.get(1) {
                Some(a) => num(a)? as i32,
                None => 0,
            };
            let factor = 10f64.powi(digits);
            Ok(Value::Number((n * factor).trunc() / factor))
        }
        "SIGN" => {
            let n = num(args.first().ok_or(ExcelError::Value)?)?;
            Ok(Value::Number(if n > 0.0 {
                1.0
            } else if n < 0.0 {
                -1.0
            } else {
                0.0
            }))
        }
        "DEGREES" => Ok(Value::Number(one(args)?.to_degrees())),
        "RADIANS" => Ok(Value::Number(one(args)?.to_radians())),
        // COMBIN and PERMUT truncate their arguments; both need 0 <= k <= n.
        "COMBIN" => Ok(Value::Number(combin(
            num(args.first().ok_or(ExcelError::Value)?)?,
            num(args.get(1).ok_or(ExcelError::Value)?)?,
        )?)),
        "PERMUT" => Ok(Value::Number(permut(
            num(args.first().ok_or(ExcelError::Value)?)?,
            num(args.get(1).ok_or(ExcelError::Value)?)?,
        )?)),
        "FACT" => Ok(Value::Number(factorial(one(args)?)?)),
        "FACTDOUBLE" => Ok(Value::Number(factdouble(one(args)?)?)),
        // Read a string of digits in a base, and write a number in one.
        "DECIMAL" => {
            expect(args, 2)?;
            let digits = text(&args[0])?;
            let radix = num(&args[1])?.trunc();
            if !(2.0..=36.0).contains(&radix) {
                return Err(ExcelError::Num);
            }
            let radix = radix as u32;
            let mut acc: i64 = 0;
            for ch in digits.trim().chars() {
                let d = ch.to_digit(radix).ok_or(ExcelError::Num)?;
                acc = acc
                    .checked_mul(radix as i64)
                    .and_then(|a| a.checked_add(d as i64))
                    .ok_or(ExcelError::Num)?;
            }
            Ok(Value::Number(acc as f64))
        }
        "BASE" => {
            if !(2..=3).contains(&args.len()) {
                return Err(ExcelError::Value);
            }
            let value = num(&args[0])?.trunc();
            let radix = num(&args[1])?.trunc();
            let min_len = match args.get(2) {
                Some(a) => num(a)?.trunc().max(0.0) as usize,
                None => 0,
            };
            if value < 0.0 || !(2.0..=36.0).contains(&radix) {
                return Err(ExcelError::Num);
            }
            let radix = radix as u64;
            let mut digits = Vec::new();
            let mut n = value as u64;
            if n == 0 {
                digits.push(b'0');
            }
            while n > 0 {
                let d = (n % radix) as u8;
                digits.push(if d < 10 { b'0' + d } else { b'A' + d - 10 });
                n /= radix;
            }
            while digits.len() < min_len {
                digits.push(b'0');
            }
            digits.reverse();
            Ok(Value::text(String::from_utf8(digits).unwrap_or_default()))
        }
        // Bitwise, on non-negative integers below 2^48.
        "BITAND" | "BITOR" => {
            expect(args, 2)?;
            let a = num(&args[0])?.trunc();
            let b = num(&args[1])?.trunc();
            let limit = 281_474_976_710_655.0;
            if a < 0.0 || b < 0.0 || a > limit || b > limit {
                return Err(ExcelError::Num);
            }
            let (a, b) = (a as u64, b as u64);
            Ok(Value::Number(
                (if name == "BITAND" { a & b } else { a | b }) as f64,
            ))
        }
        "ROUND" | "ROUNDUP" | "ROUNDDOWN" => {
            let n = num(&args.first().ok_or(ExcelError::Value)?.clone())?;
            let digits = match args.get(1) {
                Some(a) => num(a)? as i32,
                None => 0,
            };
            let factor = 10f64.powi(digits);
            // Scaled as the fifteen-digit decimal Excel keeps, so a binary
            // shortfall does not decide it: measured, ROUND(0.285,2) is 0.29.
            let scaled = n * factor;
            let scaled = if scaled.is_finite() && scaled != 0.0 {
                format!("{scaled:.14e}").parse::<f64>().unwrap_or(scaled)
            } else {
                scaled
            };
            let rounded = match name {
                // Excel rounds halves away from zero, which is what f64::round does.
                "ROUND" => scaled.round(),
                "ROUNDUP" => scaled.abs().ceil() * scaled.signum(),
                _ => scaled.abs().floor() * scaled.signum(),
            };
            Ok(Value::Number(rounded / factor))
        }

        // ---- logical -----------------------------------------------------
        "IF" => {
            if args.len() < 2 {
                return Err(ExcelError::Value);
            }
            let cond = args[0].scalar();
            if let Some(e) = cond.err() {
                return Err(e);
            }
            if cond.to_logical()? {
                Ok(args[1].scalar())
            } else {
                Ok(args.get(2).map(|a| a.scalar()).unwrap_or(Value::Logical(false)))
            }
        }
        "IFERROR" => {
            expect(args, 2)?;
            let v = args[0].scalar();
            Ok(if v.is_error() { args[1].scalar() } else { v })
        }
        "IFNA" => {
            expect(args, 2)?;
            let v = args[0].scalar();
            Ok(if v.err() == Some(ExcelError::NA) {
                args[1].scalar()
            } else {
                v
            })
        }
        "AND" | "OR" => {
            let mut seen = false;
            let mut acc = name == "AND";
            for v in args.iter().flat_map(|a| a.flatten()) {
                if let Some(e) = v.err() {
                    return Err(e);
                }
                // Blanks and text inside ranges are skipped by AND/OR.
                let b = match v {
                    Value::Logical(b) => b,
                    Value::Number(n) => n != 0.0,
                    _ => continue,
                };
                seen = true;
                acc = if name == "AND" { acc && b } else { acc || b };
            }
            if !seen {
                return Err(ExcelError::Value);
            }
            Ok(Value::Logical(acc))
        }
        // True when an ODD number of its arguments are true. Blanks and text
        // inside ranges are skipped, as AND/OR skip them. Xor(True,True,True)
        // is True; Xor(True,False,True) is False.
        "XOR" => {
            let mut seen = false;
            let mut trues = 0u64;
            for v in args.iter().flat_map(|a| a.flatten()) {
                if let Some(e) = v.err() {
                    return Err(e);
                }
                let b = match v {
                    Value::Logical(b) => b,
                    Value::Number(n) => n != 0.0,
                    _ => continue,
                };
                seen = true;
                if b {
                    trues += 1;
                }
            }
            if !seen {
                return Err(ExcelError::Value);
            }
            Ok(Value::Logical(trues % 2 == 1))
        }
        "NOT" => {
            expect(args, 1)?;
            Ok(Value::Logical(!args[0].scalar().to_logical()?))
        }
        "TRUE" => Ok(Value::Logical(true)),
        "FALSE" => Ok(Value::Logical(false)),
        "NA" => Err(ExcelError::NA),

        // ---- information -------------------------------------------------
        "ISBLANK" => Ok(Value::Logical(args.first().map(|a| a.scalar().is_blank()).unwrap_or(false))),
        "ISNUMBER" => Ok(Value::Logical(matches!(one_value(args), Value::Number(_)))),
        "ISTEXT" => Ok(Value::Logical(matches!(one_value(args), Value::Text(_)))),
        "ISLOGICAL" => Ok(Value::Logical(matches!(one_value(args), Value::Logical(_)))),
        "ISERROR" => Ok(Value::Logical(one_value(args).is_error())),
        "ISERR" => Ok(Value::Logical(matches!(
            one_value(args).err(),
            Some(e) if e != ExcelError::NA
        ))),
        "ISNA" => Ok(Value::Logical(one_value(args).err() == Some(ExcelError::NA))),
        // ODD/EVEN test the truncated whole number.
        "ISODD" => Ok(Value::Logical(one(args)?.trunc() as i64 % 2 != 0)),
        "ISEVEN" => Ok(Value::Logical(one(args)?.trunc() as i64 % 2 == 0)),
        // N turns a value into a number: a number stays, a logical is 0/1, a
        // date is its serial (already a number here), and text is 0. An error
        // arrives already handed back by the error check above.
        // Of a range, the first cell: measured, `=N(A1:A3)` over 1,2,3 is 1.
        "N" => Ok(match match args.first() {
            Some(Arg::Range(block)) => block.cells.first().cloned().unwrap_or(Value::Blank),
            _ => one_value(args),
        } {
            Value::Number(n) => Value::Number(n),
            Value::Logical(b) => Value::Number(f64::from(b)),
            _ => Value::Number(0.0),
        }),
        // 1 number, 2 text, 4 logical, 16 error, 64 an array of more than one.
        "TYPE" => {
            let arg = one_arg(args)?;
            Ok(Value::Number(match arg {
                Arg::Range(block) if block.cells.len() > 1 => 64.0,
                _ => match arg.scalar() {
                    Value::Text(_) => 2.0,
                    Value::Logical(_) => 4.0,
                    Value::Error(_) => 16.0,
                    _ => 1.0,
                },
            }))
        }
        // The number Excel gives each error kind, else #N/A.
        "ERROR.TYPE" => match one_value(args) {
            Value::Error(e) => Ok(Value::Number(match e {
                ExcelError::Null => 1.0,
                ExcelError::DivZero => 2.0,
                ExcelError::Value => 3.0,
                ExcelError::Ref => 4.0,
                ExcelError::Name => 5.0,
                ExcelError::Num => 6.0,
                ExcelError::NA => 7.0,
                ExcelError::Spill => 9.0,
                ExcelError::Calc => 14.0,
            })),
            _ => Err(ExcelError::NA),
        },

        // ---- text --------------------------------------------------------
        "LEN" => Ok(Value::Number(utf16(&text(one_arg(args)?)?).len() as f64)),
        "LEFT" | "RIGHT" => {
            let s = utf16(&text(one_arg(args)?)?);
            let n = match args.get(1) {
                Some(a) => num(a)?,
                None => 1.0,
            };
            if n < 0.0 {
                return Err(ExcelError::Value);
            }
            let n = (n as usize).min(s.len());
            let slice = if name == "LEFT" {
                &s[..n]
            } else {
                &s[s.len() - n..]
            };
            Ok(Value::Text(from_utf16(slice)))
        }
        "MID" => {
            expect(args, 3)?;
            let s = utf16(&text(&args[0])?);
            let start = num(&args[1])?;
            let len = num(&args[2])?;
            if start < 1.0 || len < 0.0 {
                return Err(ExcelError::Value);
            }
            let start = (start as usize - 1).min(s.len());
            let end = (start + len as usize).min(s.len());
            Ok(Value::Text(from_utf16(&s[start..end])))
        }
        // Full-width letters, digits and katakana to their half-width forms;
        // kanji and already-half-width text are left as they are. On an en-US
        // Excel this is available where JIS, the reverse, is #NAME?.
        // DBCS: ASC turned round. Measured: ASCII goes to its full-width
        // form except \ (to U+FFE5), ' and " (to the closing quotes U+2019
        // and U+201D) and the space (U+3000); half-width katakana take a
        // following sound mark into one letter (ｶﾞ is ガ) -- but ｳﾞ stays
        // ウ and ゛.
        "DBCS" => {
            let source = text(&args[0])?;
            let wide = |half: &str| -> Option<char> {
                ('\u{3000}'..='\u{30FF}').find(|full| *full != '\u{30F4}' && asc_halfwidth(*full) == Some(half))
            };
            let chars: Vec<char> = source.chars().collect();
            let mut out = String::with_capacity(source.len() * 3);
            let mut i = 0;
            while i < chars.len() {
                let ch = chars[i];
                if let Some(mark) = chars.get(i + 1).filter(|m| matches!(m, '\u{FF9E}' | '\u{FF9F}')) {
                    let pair: String = [ch, *mark].iter().collect();
                    if let Some(full) = wide(&pair) {
                        out.push(full);
                        i += 2;
                        continue;
                    }
                }
                let code = ch as u32;
                match ch {
                    ' ' => out.push('\u{3000}'),
                    '\\' => out.push('\u{FFE5}'),
                    '\'' => out.push('\u{2019}'),
                    '"' => out.push('\u{201D}'),
                    _ if (0x21..=0x7E).contains(&code) => out.push(char::from_u32(code + 0xFEE0).unwrap_or(ch)),
                    _ => match wide(&ch.to_string()) {
                        Some(full) => out.push(full),
                        None => out.push(ch),
                    },
                }
                i += 1;
            }
            Ok(Value::text(out))
        }
        "ASC" => {
            let source = text(&args[0])?;
            let mut out = String::with_capacity(source.len());
            for ch in source.chars() {
                let code = ch as u32;
                if (0xFF01..=0xFF5E).contains(&code) {
                    out.push(char::from_u32(code - 0xFEE0).unwrap_or(ch));
                } else if let Some(half) = asc_halfwidth(ch) {
                    out.push_str(half);
                } else {
                    out.push(ch);
                }
            }
            Ok(Value::text(out))
        }
        // The text on one side of the nth occurrence of a delimiter. A
        // negative instance counts occurrences from the end; an instance past
        // the last is #N/A, or `if_not_found` when given. The delimiter may be
        // a list of them; match_mode 1 ignores case; match_end 1 counts the
        // end of the text (its start, counting back) as one more delimiter.
        // A blank argument is its default: measured,
        // `TEXTAFTER("x","-",,,,"none")` is "none" and
        // `TEXTBEFORE("a-b-c","B",1,1)` "a-".
        "TEXTBEFORE" | "TEXTAFTER" => {
            if args.len() < 2 || args.len() > 6 {
                return Err(ExcelError::Value);
            }
            let given = |at: usize| args.get(at).filter(|one| !matches!(one.scalar(), Value::Blank));
            let hay = text(&args[0])?;
            let mut needles = Vec::new();
            for one in args[1].flatten() {
                if let Value::Error(why) = one {
                    return Err(why);
                }
                needles.push(one.to_text()?);
            }
            let instance = match given(2) {
                Some(a) => num(a)?.trunc() as i64,
                None => 1,
            };
            let ignore_case = match given(3) {
                Some(a) => num(a)? != 0.0,
                None => false,
            };
            let match_end = match given(4) {
                Some(a) => num(a)? != 0.0,
                None => false,
            };
            // Measured: an instance of 0, or further than the text is long
            // (`TEXTBEFORE("a-b","-",5)`), is #VALUE!.
            // An empty delimiter sits before the first character, whatever
            // the instance: measured, TEXTAFTER("ab","",3) is "ab" and
            // TEXTBEFORE("","") is "".
            if instance > 0 && needles.iter().all(|one| one.is_empty()) {
                return Ok(Value::Text(if name == "TEXTBEFORE" { String::new() } else { hay }));
            }
            // Empty text is searched when no instance is named, and so not
            // found: measured, TEXTAFTER("","-") is #N/A where
            // TEXTAFTER("","-",1) is #VALUE!.
            let unnamed_on_empty = hay.is_empty() && given(2).is_none();
            if instance == 0 || (!unnamed_on_empty && instance.unsigned_abs() as usize > hay.chars().count()) {
                return Err(ExcelError::Value);
            }
            let fold = |t: &str| if ignore_case { t.to_lowercase() } else { t.to_string() };
            let folded = fold(&hay);
            // Case folding that changes a length would put the cuts in the
            // wrong place; such text is matched as written.
            let (search, needles): (String, Vec<String>) = if folded.len() == hay.len() {
                (folded, needles.iter().map(|one| fold(one)).collect())
            } else {
                (hay.clone(), needles)
            };
            // Each occurrence as (start, length), left to right, never
            // overlapping; the first delimiter listed wins at a place.
            let mut found: Vec<(usize, usize)> = Vec::new();
            let mut at = 0;
            while at <= search.len() {
                if !search.is_char_boundary(at) {
                    at += 1;
                    continue;
                }
                match needles.iter().find(|one| !one.is_empty() && search[at..].starts_with(one.as_str())) {
                    Some(one) => {
                        found.push((at, one.len()));
                        at += one.len();
                    }
                    None => at += 1,
                }
            }
            // An empty delimiter meets the text at its very start: measured,
            // TEXTBEFORE of "abc" by "" is "" and TEXTAFTER "abc".
            if needles.iter().all(String::is_empty) {
                found = vec![(0, 0)];
            }
            if match_end {
                if instance > 0 {
                    found.push((hay.len(), 0));
                } else {
                    found.insert(0, (0, 0));
                }
            }
            let count = found.len() as i64;
            let index = if instance > 0 { instance - 1 } else { count + instance };
            if index < 0 || index >= count {
                return match given(5) {
                    Some(fallback) => Ok(fallback.scalar()),
                    None => Err(ExcelError::NA),
                };
            }
            let (cut, length) = found[index as usize];
            Ok(Value::text(if name == "TEXTBEFORE" {
                &hay[..cut]
            } else {
                &hay[cut + length..]
            }))
        }
        // A number written with its own separators: group separators are
        // dropped, the decimal separator is honoured, and each trailing % (or
        // Japanese/full-width percents count too, but a plain one here) divides
        // by a hundred.
        "NUMBERVALUE" => {
            let s = text(&args[0])?;
            let decimal = match args.get(1) {
                Some(a) => text(a)?,
                None => ".".to_string(),
            };
            // With no group mark given, the default one gives way to a decimal
            // mark that is the same: measured, NUMBERVALUE("1,2",",") is 1.2.
            let group = match args.get(2) {
                Some(a) => text(a)?,
                None if decimal.starts_with(',') => String::new(),
                None => ",".to_string(),
            };
            let decimal = decimal.chars().next().unwrap_or('.');
            let mut body = s.trim().to_string();
            let mut percents = 0u32;
            while body.trim_end().ends_with('%') {
                let trimmed = body.trim_end();
                body = trimmed[..trimmed.len() - 1].to_string();
                percents += 1;
            }
            let mut cleaned = String::new();
            for ch in body.chars() {
                if group.contains(ch) || ch.is_whitespace() {
                    continue;
                }
                cleaned.push(if ch == decimal { '.' } else { ch });
            }
            if cleaned.is_empty() {
                return Ok(Value::Number(0.0));
            }
            let mut value: f64 = cleaned.parse().map_err(|_| ExcelError::Value)?;
            for _ in 0..percents {
                value /= 100.0;
            }
            Ok(Value::Number(value))
        }
        "TRIM" => {
            // Excel's TRIM also collapses runs of interior spaces to one.
            let s = text(one_arg(args)?)?;
            // Only the plain space: measured, CHAR(160) stays.
            let collapsed = s.split(' ').filter(|part| !part.is_empty()).collect::<Vec<_>>().join(" ");
            Ok(Value::Text(collapsed))
        }
        // A letter whose capital is two letters keeps itself: measured,
        // UPPER("ß") is ß.
        "UPPER" => Ok(Value::Text(
            text(one_arg(args)?)?
                .chars()
                .map(|one| {
                    let mut upper = one.to_uppercase();
                    match (upper.next(), upper.next()) {
                        (Some(single), None) => single,
                        _ => one,
                    }
                })
                .collect(),
        )),
        "LOWER" => Ok(Value::Text(text(one_arg(args)?)?.to_lowercase())),
        "CONCATENATE" | "CONCAT" => {
            let mut out = String::new();
            for v in args.iter().flat_map(|a| a.flatten()) {
                out.push_str(&v.to_text()?);
            }
            Ok(Value::Text(out))
        }
        // A value, or every value of an array, as text. Concise (0) writes
        // text as it is; strict (1) quotes it. An error is its own word.
        // Measured: `VALUETOTEXT(12.5)` is 12.5, `VALUETOTEXT("a-b-c",1)`
        // "a-b-c" in quotes, `ARRAYTOTEXT(A1:A4)` "a-b-c, x, , 12.5" and
        // `ARRAYTOTEXT(A1:A2,1)` {"a-b-c";"x"}.
        "VALUETOTEXT" | "ARRAYTOTEXT" => {
            if args.is_empty() || args.len() > 2 {
                return Err(ExcelError::Value);
            }
            let strict = match args.get(1).map(Arg::scalar) {
                None | Some(Value::Blank) => false,
                Some(one) => match one.to_number()? {
                    0.0 => false,
                    1.0 => true,
                    _ => return Err(ExcelError::Value),
                },
            };
            let written = |one: &Value| match one {
                Value::Text(t) if strict => format!("\"{}\"", t.replace('"', "\"\"")),
                Value::Error(why) => why.as_str().to_string(),
                other => other.to_text().unwrap_or_default(),
            };
            if name == "VALUETOTEXT" {
                return Ok(Value::Text(written(&args[0].scalar())));
            }
            let block = args[0].as_range();
            if !strict {
                return Ok(Value::Text(block.cells.iter().map(written).collect::<Vec<_>>().join(", ")));
            }
            let rows: Vec<String> = (0..block.height)
                .map(|row| (0..block.width).map(|col| written(&block.at(col, row))).collect::<Vec<_>>().join(","))
                .collect();
            Ok(Value::Text(format!("{{{}}}", rows.join(";"))))
        }
        "REPT" => {
            expect(args, 2)?;
            let n = num(&args[1])?;
            if n < 0.0 {
                return Err(ExcelError::Value);
            }
            // No longer than a cell holds: measured, REPT("a",32768) is #VALUE!.
            let unit = text(&args[0])?;
            if unit.encode_utf16().count() * (n as usize) > 32_767 {
                return Err(ExcelError::Value);
            }
            Ok(Value::Text(unit.repeat(n as usize)))
        }
        "SUBSTITUTE" => {
            if args.len() < 3 {
                return Err(ExcelError::Value);
            }
            let (s, old, new) = (text(&args[0])?, text(&args[1])?, text(&args[2])?);
            if old.is_empty() {
                return Ok(Value::Text(s));
            }
            // With a fourth argument only that one occurrence, counted from
            // one, is replaced: measured, `SUBSTITUTE("aaa","a","b",2)` is
            // "aba". Without it every occurrence goes.
            match args.get(3) {
                None => Ok(Value::Text(s.replace(&old, &new))),
                Some(which) => {
                    let which = num(which)?;
                    if which < 1.0 {
                        return Err(ExcelError::Value);
                    }
                    let which = which as usize;
                    let mut seen = 0usize;
                    let mut out = String::new();
                    let mut rest = s.as_str();
                    while let Some(at) = rest.find(&old) {
                        seen += 1;
                        out.push_str(&rest[..at]);
                        if seen == which {
                            out.push_str(&new);
                        } else {
                            out.push_str(&old);
                        }
                        rest = &rest[at + old.len()..];
                    }
                    out.push_str(rest);
                    Ok(Value::Text(out))
                }
            }
        }
        // FIND is case-sensitive, SEARCH is not. Both are 1-based in UTF-16 units.
        "FIND" | "SEARCH" => {
            if args.len() < 2 {
                return Err(ExcelError::Value);
            }
            let needle = text(&args[0])?;
            let haystack = text(&args[1])?;
            let (needle, haystack) = if name == "SEARCH" {
                (needle.to_lowercase(), haystack.to_lowercase())
            } else {
                (needle, haystack)
            };
            // A start before the first character is #VALUE!: measured,
            // FIND("a","abc",0).
            let start = match args.get(2) {
                Some(a) => {
                    let asked = num(a)?.trunc();
                    if asked < 1.0 {
                        return Err(ExcelError::Value);
                    }
                    asked as usize - 1
                }
                None => 0,
            };
            let units = utf16(&haystack);
            let needle_units = utf16(&needle);
            if start > units.len() {
                return Err(ExcelError::Value);
            }
            // SEARCH reads `*`, `?` and `~` as FIND does not: measured,
            // `SEARCH("*na","banana")` is 1 -- the first place the pattern
            // matches the text from.
            if name == "SEARCH" && has_wildcards(&needle) {
                let chars: Vec<char> = haystack.chars().collect();
                let pattern = format!("{needle}*");
                let mut unit_at = 0usize;
                for (index, _) in chars.iter().enumerate() {
                    if unit_at >= start {
                        let rest: String = chars[index..].iter().collect();
                        if wildcard_match(&rest, &pattern) {
                            return Ok(Value::Number((unit_at + 1) as f64));
                        }
                    }
                    unit_at += chars[index].len_utf16();
                }
                return Err(ExcelError::Value);
            }
            let found = units[start..]
                .windows(needle_units.len().max(1))
                .position(|w| w == needle_units.as_slice());
            match found {
                Some(idx) => Ok(Value::Number((start + idx + 1) as f64)),
                None if needle_units.is_empty() => Ok(Value::Number((start + 1) as f64)),
                None => Err(ExcelError::Value),
            }
        }
        "VALUE" => Ok(Value::Number(
            Value::Text(text(one_arg(args)?)?).to_number()?,
        )),
        // Measured: BINOM.DIST(3,10,0.5,FALSE) is 0.117188 as shown,
        // POISSON.DIST(2,3,TRUE) 0.42319, EXPON.DIST(1,2,TRUE) 0.864665.
        "BINOM.DIST" | "BINOMDIST" => {
            expect(args, 4)?;
            let (k, n, p) = (num(&args[0])?.trunc(), num(&args[1])?.trunc(), num(&args[2])?);
            let cumulative = args[3].scalar().to_logical()?;
            if k < 0.0 || k > n || !(0.0..=1.0).contains(&p) {
                return Err(ExcelError::Num);
            }
            // Exact products while they fit, so a tidy answer stays tidy:
            // 120/1024 is 0.1171875 and shows as 0.117188.
            let mass = |j: f64| {
                if n <= 1000.0 {
                    let mut choose = 1.0;
                    for i in 0..j as i64 {
                        choose = choose * (n - i as f64) / (i as f64 + 1.0);
                    }
                    choose * p.powi(j as i32) * (1.0 - p).powi((n - j) as i32)
                } else {
                    (ln_choose(n, j) + j * p.ln() + (n - j) * (1.0 - p).ln()).exp()
                }
            };
            let answer = if cumulative { (0..=k as i64).map(|j| mass(j as f64)).sum() } else { mass(k) };
            Ok(Value::Number(answer))
        }
        "POISSON.DIST" | "POISSON" => {
            expect(args, 3)?;
            let (x, mean) = (num(&args[0])?.trunc(), num(&args[1])?);
            let cumulative = args[2].scalar().to_logical()?;
            if x < 0.0 || mean < 0.0 {
                return Err(ExcelError::Num);
            }
            let mass = |j: f64| {
                if j <= 170.0 {
                    let mut term = (-mean).exp();
                    for i in 1..=j as i64 {
                        term = term * mean / i as f64;
                    }
                    term
                } else {
                    (j * mean.ln() - mean - ln_gamma(j + 1.0)).exp()
                }
            };
            // Measured to the last digit: the point mass by exact products
            // (POISSON.DIST(4,2.5,FALSE) 0.133601885781085), the running sum
            // by logarithms (POISSON.DIST(2,3,TRUE) 0.423190081126843).
            let by_logs = |j: f64| (j * mean.ln() - mean - lanczos_ln_gamma(j + 1.0)).exp();
            let answer = if cumulative { (0..=x as i64).map(|j| by_logs(j as f64)).sum() } else { mass(x) };
            Ok(Value::Number(answer))
        }
        "EXPON.DIST" | "EXPONDIST" => {
            expect(args, 3)?;
            let (x, lambda) = (num(&args[0])?, num(&args[1])?);
            let cumulative = args[2].scalar().to_logical()?;
            if x < 0.0 || lambda <= 0.0 {
                return Err(ExcelError::Num);
            }
            Ok(Value::Number(if cumulative { 1.0 - (-lambda * x).exp() } else { lambda * (-lambda * x).exp() }))
        }
        // T.TEST(a, b, tails, type): paired (1), equal variance (2) or
        // unequal (3). Measured: a sample paired with itself is #DIV/0!.
        "T.TEST" | "TTEST" => {
            expect(args, 4)?;
            let numbers = |arg: &Arg| arg.flatten().into_iter().filter_map(|v| match v {
                Value::Number(n) => Some(n),
                _ => None,
            }).collect::<Vec<f64>>();
            let (a, b) = (numbers(&args[0]), numbers(&args[1]));
            let tails = num(&args[2])?.trunc();
            let kind = num(&args[3])?.trunc();
            if !(tails == 1.0 || tails == 2.0) || !(1.0..=3.0).contains(&kind) {
                return Err(ExcelError::Num);
            }
            let mean = |xs: &[f64]| xs.iter().sum::<f64>() / xs.len() as f64;
            let variance = |xs: &[f64]| {
                let m = mean(xs);
                xs.iter().map(|x| (x - m) * (x - m)).sum::<f64>() / (xs.len() as f64 - 1.0)
            };
            let (t, df) = if kind == 1.0 {
                if a.len() != b.len() {
                    return Err(ExcelError::NA);
                }
                let d: Vec<f64> = a.iter().zip(&b).map(|(x, y)| x - y).collect();
                if d.len() < 2 {
                    return Err(ExcelError::DivZero);
                }
                let spread = variance(&d);
                if spread == 0.0 {
                    return Err(ExcelError::DivZero);
                }
                (mean(&d) / (spread / d.len() as f64).sqrt(), d.len() as f64 - 1.0)
            } else {
                if a.len() < 2 || b.len() < 2 {
                    return Err(ExcelError::DivZero);
                }
                let (na, nb) = (a.len() as f64, b.len() as f64);
                let (va, vb) = (variance(&a), variance(&b));
                if kind == 2.0 {
                    let squares = |xs: &[f64]| {
                        let m = mean(xs);
                        xs.iter().map(|x| (x - m) * (x - m)).sum::<f64>()
                    };
                    let pooled = (squares(&a) + squares(&b)) / (na + nb - 2.0);
                    let se = (pooled * (1.0 / na + 1.0 / nb)).sqrt();
                    if se == 0.0 {
                        return Err(ExcelError::DivZero);
                    }
                    ((mean(&a) - mean(&b)) / se, na + nb - 2.0)
                } else {
                    let (sa, sb) = (va / na, vb / nb);
                    let se = (sa + sb).sqrt();
                    if se == 0.0 {
                        return Err(ExcelError::DivZero);
                    }
                    let df = (sa + sb).powi(2) / (sa * sa / (na - 1.0) + sb * sb / (nb - 1.0));
                    ((mean(&a) - mean(&b)) / se, df)
                }
            };
            let upper = 0.5 * regularized_beta(df / (df + t * t), df / 2.0, 0.5);
            Ok(Value::Number(tails * upper))
        }

        // ---- conditional aggregates ---------------------------------------
        "COUNTIF" => {
            expect(args, 2)?;
            let criteria = Criteria::parse(&args[1].scalar());
            let count = args[0]
                .flatten()
                .iter()
                .filter(|v| criteria.matches(v))
                .count();
            Ok(Value::Number(count as f64))
        }
        // ---- several conditions at once --------------------------------
        //
        // SUMIFS reads its ranges the other way round from SUMIF: the range to
        // add comes FIRST, and the pairs to test follow it. Getting that the
        // wrong way round is the classic way to write one of these.
        "SUMIFS" | "COUNTIFS" | "AVERAGEIFS" => {
            let counting = name == "COUNTIFS";
            let pairs = if counting { &args[0..] } else { &args[1..] };
            if pairs.len() < 2 || pairs.len() % 2 != 0 {
                return Err(ExcelError::Value);
            }
            // For SUMIFS and AVERAGEIFS this is the range to add up; for
            // COUNTIFS it is the first range to test, and is only used for its
            // length.
            let over = args[0].flatten();
            // An error for a criterion looks for that error, as COUNTIF's
            // does: measured, `SUMIFS(B1:B5,A1:A5,C2)` with C2 #DIV/0! adds
            // the row holding #DIV/0!.
            let mut total = 0.0;
            let mut seen = 0.0;
            for at in 0..over.len() {
                let mut all = true;
                for pair in pairs.chunks(2) {
                    let tested = pair[0].flatten();
                    let criteria = Criteria::parse(&pair[1].scalar());
                    // A row is only counted when every range has something to
                    // say about it; ranges of different lengths are Excel's
                    // #VALUE!, but a short one simply fails to match here.
                    match tested.get(at) {
                        Some(value) if criteria.matches(value) => {}
                        _ => {
                            all = false;
                            break;
                        }
                    }
                }
                if !all {
                    continue;
                }
                seen += 1.0;
                if !counting {
                    match over.get(at) {
                        // An error on a row that matched is being added up.
                        Some(Value::Error(why)) => return Err(*why),
                        Some(Value::Number(n)) => total += n,
                        _ => {}
                    }
                }
            }
            Ok(match name {
                "COUNTIFS" => Value::Number(seen),
                "SUMIFS" => Value::Number(total),
                _ if seen == 0.0 => Value::Error(ExcelError::DivZero),
                _ => Value::Number(total / seen),
            })
        }
        // The largest or smallest of a range where every criterion holds; no
        // match is 0, as Excel gives.
        "MAXIFS" | "MINIFS" => {
            let pairs = &args[1..];
            if pairs.is_empty() || !pairs.len().is_multiple_of(2) {
                return Err(ExcelError::Value);
            }
            let over = args[0].flatten();
            for pair in pairs.chunks(2) {
                if let Some(why) = pair[1].scalar().err() {
                    return Err(why);
                }
            }
            let want_max = name == "MAXIFS";
            let mut best: Option<f64> = None;
            for at in 0..over.len() {
                let mut all = true;
                for pair in pairs.chunks(2) {
                    let tested = pair[0].flatten();
                    let criteria = Criteria::parse(&pair[1].scalar());
                    match tested.get(at) {
                        Some(value) if criteria.matches(value) => {}
                        _ => {
                            all = false;
                            break;
                        }
                    }
                }
                if !all {
                    continue;
                }
                match over.get(at) {
                    Some(Value::Error(why)) => return Err(*why),
                    Some(Value::Number(n)) => {
                        best = Some(match best {
                            None => *n,
                            Some(b) => if want_max { b.max(*n) } else { b.min(*n) },
                        });
                    }
                    _ => {}
                }
            }
            Ok(Value::Number(best.unwrap_or(0.0)))
        }
        // The A-suffix max/min: text counts as zero and logicals as 0/1, in a
        // range as well as an argument.
        "MAXA" | "MINA" => {
            let want_max = name == "MAXA";
            let mut best: Option<f64> = None;
            for value in args.iter().flat_map(|a| a.flatten()) {
                let x = match value {
                    Value::Error(e) => return Err(e),
                    Value::Number(n) => n,
                    Value::Logical(b) => f64::from(b),
                    Value::Text(_) => 0.0,
                    Value::Blank => continue,
                };
                best = Some(match best {
                    None => x,
                    Some(b) => if want_max { b.max(x) } else { b.min(x) },
                });
            }
            Ok(Value::Number(best.unwrap_or(0.0)))
        }
        // The A-suffix sample spread, with the same reading of text and logicals.
        "STDEVA" | "VARA" => {
            let mut values = Vec::new();
            for value in args.iter().flat_map(|a| a.flatten()) {
                match value {
                    Value::Error(e) => return Err(e),
                    Value::Number(n) => values.push(n),
                    Value::Logical(b) => values.push(f64::from(b)),
                    Value::Text(_) => values.push(0.0),
                    Value::Blank => {}
                }
            }
            if values.len() < 2 {
                return Err(ExcelError::DivZero);
            }
            let mean = values.iter().sum::<f64>() / values.len() as f64;
            let variance =
                values.iter().map(|x| (x - mean) * (x - mean)).sum::<f64>() / (values.len() - 1) as f64;
            Ok(Value::Number(if name == "VARA" { variance } else { variance.sqrt() }))
        }
        // Combinations with repetition, C(n+k-1, k).
        "COMBINA" => {
            let n = num(args.first().ok_or(ExcelError::Value)?)?.trunc();
            let k = num(args.get(1).ok_or(ExcelError::Value)?)?.trunc();
            if n < 0.0 || k < 0.0 {
                return Err(ExcelError::Num);
            }
            if n == 0.0 {
                return if k == 0.0 { Ok(Value::Number(1.0)) } else { Err(ExcelError::Num) };
            }
            Ok(Value::Number(combin(n + k - 1.0, k)?))
        }
        // The character at a Unicode code point.
        "UNICHAR" => {
            let n = one(args)?.trunc();
            if !(1.0..=1_114_111.0).contains(&n) {
                return Err(ExcelError::Value);
            }
            match char::from_u32(n as u32) {
                Some(c) => Ok(Value::text(c.to_string())),
                None => Err(ExcelError::Value),
            }
        }

        // ---- the normal distribution -------------------------------------
        "NORM.DIST" | "NORMDIST" => {
            expect(args, 4)?;
            let sd = num(&args[2])?;
            if sd <= 0.0 {
                return Err(ExcelError::Num);
            }
            let z = (num(&args[0])? - num(&args[1])?) / sd;
            let cumulative = args[3].scalar().to_logical()?;
            Ok(Value::Number(if cumulative {
                norm_cdf(z)
            } else {
                norm_pdf(z) / sd
            }))
        }
        "NORM.S.DIST" => {
            expect(args, 2)?;
            let z = num(&args[0])?;
            let cumulative = args[1].scalar().to_logical()?;
            Ok(Value::Number(if cumulative { norm_cdf(z) } else { norm_pdf(z) }))
        }
        // The legacy one-argument form is always cumulative.
        "NORMSDIST" => Ok(Value::Number(norm_cdf(one(args)?))),
        "NORM.INV" | "NORMINV" => {
            expect(args, 3)?;
            let p = num(&args[0])?;
            let sd = num(&args[2])?;
            if sd <= 0.0 || p <= 0.0 || p >= 1.0 {
                return Err(ExcelError::Num);
            }
            Ok(Value::Number(num(&args[1])? + sd * norm_s_inv(p)))
        }
        "NORM.S.INV" | "NORMSINV" => {
            let p = one(args)?;
            if p <= 0.0 || p >= 1.0 {
                return Err(ExcelError::Num);
            }
            Ok(Value::Number(norm_s_inv(p)))
        }
        "STANDARDIZE" => {
            expect(args, 3)?;
            let sd = num(&args[2])?;
            if sd <= 0.0 {
                return Err(ExcelError::Num);
            }
            Ok(Value::Number((num(&args[0])? - num(&args[1])?) / sd))
        }
        "GAUSS" => Ok(Value::Number(norm_cdf(one(args)?) - 0.5)),
        "PHI" => Ok(Value::Number(norm_pdf(one(args)?))),
        "CONFIDENCE.NORM" | "CONFIDENCE" => {
            expect(args, 3)?;
            let alpha = num(&args[0])?;
            let sd = num(&args[1])?;
            let size = num(&args[2])?.trunc();
            if alpha <= 0.0 || alpha >= 1.0 || sd <= 0.0 || size < 1.0 {
                return Err(ExcelError::Num);
            }
            Ok(Value::Number(norm_s_inv(1.0 - alpha / 2.0) * sd / size.sqrt()))
        }

        // ---- the other continuous and discrete distributions -------------
        // Measured against Excel: T.DIST(1.5,10,TRUE) 0.91774633677728,
        // CHISQ.DIST.RT(3,4) 0.557825400371075, F.DIST(2,3,10,TRUE)
        // 0.821992592624824, GAMMA.DIST(2,3,1,TRUE) 0.323323583816937,
        // BETA.DIST(0.4,2,3,TRUE) 0.5248, ERF(1) 0.842700792949715.
        "T.DIST" | "T.DIST.RT" | "T.DIST.2T" | "TDIST" | "T.INV" | "T.INV.2T" | "TINV"
        | "CHISQ.DIST" | "CHISQ.DIST.RT" | "CHIDIST" | "CHISQ.INV" | "CHISQ.INV.RT" | "CHIINV"
        | "F.DIST" | "F.DIST.RT" | "FDIST" | "F.INV" | "F.INV.RT" | "FINV"
        | "GAMMA.DIST" | "GAMMADIST" | "GAMMA.INV" | "GAMMAINV" | "GAMMALN" | "GAMMALN.PRECISE" | "GAMMA"
        | "BETA.DIST" | "BETADIST" | "BETA.INV" | "BETAINV"
        | "LOGNORM.DIST" | "LOGNORMDIST" | "LOGNORM.INV" | "LOGINV"
        | "HYPGEOM.DIST" | "HYPGEOMDIST" | "NEGBINOM.DIST" | "NEGBINOMDIST"
        | "WEIBULL.DIST" | "WEIBULL" | "FISHER" | "FISHERINV"
        | "ERF" | "ERF.PRECISE" | "ERFC" | "ERFC.PRECISE" | "BINOM.INV" | "CRITBINOM" | "CONFIDENCE.T" => {
            distribution(name, args).map(Value::Number)
        }

        // Multiply the arrays together elementwise and add up the lot. Text and
        // blanks count as nothing rather than spoiling the sum, which is what
        // makes `SUMPRODUCT((A=x)*(B=y), C)` work at all — the comparisons come
        // through as TRUE and FALSE and have to weigh one and nothing.
        "SUMPRODUCT" => {
            if args.is_empty() {
                return Err(ExcelError::Value);
            }
            let columns: Vec<Vec<Value>> = args.iter().map(|one| one.flatten()).collect();
            let reach = columns.iter().map(|one| one.len()).max().unwrap_or(0);
            let mut total = 0.0;
            for at in 0..reach {
                let mut running = 1.0;
                for column in &columns {
                    // Arrays of different lengths are #VALUE! in Excel, and a
                    // missing cell here is treated as one.
                    if column.len() != reach && column.len() != 1 {
                        return Err(ExcelError::Value);
                    }
                    let value = if column.len() == 1 { &column[0] } else { &column[at] };
                    running *= match value {
                        Value::Number(n) => *n,
                        // A Boolean counts for nothing, as text does:
                        // measured, TRUE among the cells adds 0.
                        Value::Logical(_) | Value::Blank | Value::Text(_) => 0.0,
                        Value::Error(e) => return Err(*e),
                    };
                    if running == 0.0 {
                        break;
                    }
                }
                total += running;
            }
            Ok(Value::Number(total))
        }

        // ---- how big is this ---------------------------------------------
        "ROWS" | "COLUMNS" => {
            let shape = one_arg(args)?.as_range();
            Ok(Value::Number(if name == "ROWS" {
                shape.height as f64
            } else {
                shape.width as f64
            }))
        }

        // ---- the nth smallest, and the nth largest ------------------------
        "SMALL" | "LARGE" => {
            if args.len() < 2 {
                return Err(ExcelError::Value);
            }
            let mut numbers: Vec<f64> = args[0]
                .flatten()
                .iter()
                .filter_map(|one| match one {
                    Value::Number(n) => Some(*n),
                    _ => None,
                })
                .collect();
            if numbers.is_empty() {
                return Err(ExcelError::Num);
            }
            numbers.sort_by(|a, b| a.partial_cmp(b).unwrap_or(Ordering::Equal));
            let nth = num(&args[1])?;
            if nth < 1.0 || nth as usize > numbers.len() {
                return Err(ExcelError::Num);
            }
            let at = nth as usize - 1;
            Ok(Value::Number(if name == "SMALL" {
                numbers[at]
            } else {
                numbers[numbers.len() - 1 - at]
            }))
        }

        // Where a number comes in a list, counting from the largest unless
        // told otherwise. Equal numbers share the higher place, and the places
        // after them are skipped — two firsts are followed by a third.
        "RANK" | "RANK.EQ" => {
            if args.len() < 2 {
                return Err(ExcelError::Value);
            }
            let wanted = num(&args[0])?;
            let numbers: Vec<f64> = args[1]
                .flatten()
                .iter()
                .filter_map(|one| match one {
                    Value::Number(n) => Some(*n),
                    _ => None,
                })
                .collect();
            let up = match args.get(2) {
                Some(one) => num(one)? != 0.0,
                None => false,
            };
            if !numbers.contains(&wanted) {
                return Err(ExcelError::NA);
            }
            let ahead = numbers
                .iter()
                .filter(|one| if up { **one < wanted } else { **one > wanted })
                .count();
            Ok(Value::Number(ahead as f64 + 1.0))
        }

        // ---- rounding away from zero to a multiple ------------------------
        "CEILING" | "FLOOR" | "CEILING.MATH" | "FLOOR.MATH" => {
            let value = num(&args[0])?;
            let step = match args.get(1) {
                Some(one) => num(one)?,
                // The .MATH forms take a step of one when none is given; the
                // older ones insist on being told.
                None if name.ends_with(".MATH") => 1.0,
                None => return Err(ExcelError::Value),
            };
            // FLOOR by nothing divides by nothing: measured, FLOOR(5,0) is
            // #DIV/0! where CEILING(0,0) is 0.
            if step == 0.0 && name == "FLOOR" && value != 0.0 {
                return Err(ExcelError::DivZero);
            }
            if step == 0.0 {
                return Ok(Value::Number(0.0));
            }
            // Excel refuses a positive number rounded to a negative step.
            if value > 0.0 && step < 0.0 && !name.ends_with(".MATH") {
                return Err(ExcelError::Num);
            }
            let up = name.starts_with("CEILING");
            // The .MATH forms' third argument sends a negative number the
            // other way: measured, `FLOOR.MATH(-4.5,2,1)` is -4, toward
            // zero, where `FLOOR.MATH(-4.5,2)` is -6.
            if name.ends_with(".MATH") {
                let step = step.abs();
                let mode = match args.get(2) {
                    Some(one) => num(one)?,
                    None => 0.0,
                };
                let steps = value / step;
                let toward_zero = value < 0.0 && mode != 0.0;
                let rounded = match (up, toward_zero) {
                    (true, false) => steps.ceil(),
                    (true, true) => steps.floor(),
                    (false, false) => steps.floor(),
                    (false, true) => steps.ceil(),
                };
                return Ok(Value::Number(step * rounded));
            }
            let steps = value / step;
            Ok(Value::Number(
                step * if up { steps.ceil() } else { steps.floor() },
            ))
        }

        // ---- letters and their numbers ------------------------------------
        "CHAR" => {
            let code = one(args)?;
            if !(1.0..=255.0).contains(&code) {
                return Err(ExcelError::Value);
            }
            // Excel's CHAR is the Windows codepage, which agrees with Latin-1
            // over the whole range it accepts.
            Ok(Value::Text(
                char::from_u32(code as u32).map(String::from).unwrap_or_default(),
            ))
        }
        "CODE" | "UNICODE" => {
            let letters = text(one_arg(args)?)?;
            match letters.chars().next() {
                Some(one) => Ok(Value::Number(u32::from(one) as f64)),
                None => Err(ExcelError::Value),
            }
        }

        // Whether two pieces of text are the same, letter case and all —
        // which is exactly what `=` does not ask.
        "EXACT" => {
            expect(args, 2)?;
            Ok(Value::Logical(text(&args[0])? == text(&args[1])?))
        }

        // Put something in the middle of some text, over what was there.
        "REPLACE" => {
            expect(args, 4)?;
            let held = utf16(&text(&args[0])?);
            let from = num(&args[1])?;
            let many = num(&args[2])?;
            if from < 1.0 || many < 0.0 {
                return Err(ExcelError::Value);
            }
            let from = (from as usize - 1).min(held.len());
            let to = (from + many as usize).min(held.len());
            let mut out = from_utf16(&held[..from]);
            out.push_str(&text(&args[3])?);
            out.push_str(&from_utf16(&held[to..]));
            Ok(Value::Text(out))
        }

        // A number written the way a cell would show it under `format`.
        "TEXT" => {
            expect(args, 2)?;
            let format = text(&args[1])?;
            // A section with two percent signs is refused: measured,
            // TEXT(0.123,"0.0%%") and TEXT(0.123,"0%%") are #VALUE!, while
            // "0%0" and "%0" are fine.
            if crate::numfmt::sections_with_percents(&format).any(|count| count > 1) {
                return Err(ExcelError::Value);
            }
            match args[0].scalar() {
                // A date picture has no way to write a number off the
                // calendar: measured, TEXT(-0.5,"h:m:s") and
                // TEXT(1E+15,"yyyy/mm/dd") are #VALUE!.
                Value::Number(n)
                    if crate::numfmt::looks_like_a_date(&format) && !(0.0..2_958_466.0).contains(&n) =>
                {
                    Err(ExcelError::Value)
                }
                Value::Number(n) => Ok(Value::Text(crate::numfmt::format_number(n, &format))),
                // Text that reads as a number is formatted as that number,
                // and other text goes through the text section. Measured:
                // `TEXT("12","0.0")` is 12.0, `TEXT("1/2/2020","yyyy")` 2020,
                // `TEXT("abc","@@")` abcabc, `TEXT("abc","0;0;0;<@>")` <abc>
                // and `TEXT("abc","0.0")` abc.
                Value::Text(t) => Ok(Value::Text(match Value::Text(t.clone()).to_number() {
                    Ok(n) if !t.trim().is_empty() => crate::numfmt::format_number(n, &format),
                    _ => crate::numfmt::format_text(&t, &format),
                })),
                // A Boolean is text to TEXT: measured, under
                // `0;-0;"zero";"text:"@` TRUE is text:TRUE.
                Value::Logical(b) => Ok(Value::Text(crate::numfmt::format_text(
                    if b { "TRUE" } else { "FALSE" },
                    &format,
                ))),
                Value::Blank => Ok(Value::Text(String::new())),
                Value::Error(e) => Err(e),
            }
        }

        "SUMIF" | "AVERAGEIF" => {
            if args.len() < 2 {
                return Err(ExcelError::Value);
            }
            // An error for the criterion looks for that error: measured,
            // SUMIF over a range holding one #N/A, asked for #N/A, adds up
            // that row.
            let asked = args[1].scalar();
            let criteria = Criteria::parse(&asked);
            let tested = args[0].flatten();
            let summed = match args.get(2) {
                Some(a) => a.flatten(),
                None => tested.clone(),
            };
            let mut total = 0.0;
            let mut seen = 0.0;
            for (i, v) in tested.iter().enumerate() {
                if !criteria.matches(v) {
                    continue;
                }
                seen += 1.0;
                match summed.get(i) {
                    // An error on a row that MATCHED is being added up, and an
                    // error cannot be added up.
                    Some(Value::Error(why)) => return Err(*why),
                    Some(Value::Number(n)) => total += n,
                    _ => {}
                }
            }
            if name == "SUMIF" {
                return Ok(Value::Number(total));
            }
            if seen == 0.0 {
                return Err(ExcelError::DivZero);
            }
            Ok(Value::Number(total / seen))
        }

        // ---- lookup --------------------------------------------------------
        // Look through one list, take from another. No counting of columns,
        // which is what it was made to get rid of.
        "XLOOKUP" => {
            if args.len() < 3 {
                return Err(ExcelError::Value);
            }
            let key = args[0].scalar();
            if let Some(why) = key.err() {
                return Err(why);
            }
            let looked = args[1].flatten();
            let taken = args[2].flatten();
            let how = match args.get(4) {
                Some(one) => num(one)? as i32,
                None => 0,
            };
            let downwards = match args.get(5) {
                Some(one) => num(one)? >= 0.0,
                None => true,
            };
            let order: Vec<usize> = if downwards {
                (0..looked.len()).collect()
            } else {
                (0..looked.len()).rev().collect()
            };
            let found = match how {
                // Exact, and 2 is exact with wildcards — which `answers_to`
                // already reads when the key carries one.
                0 | 2 => order.into_iter().find(|at| answers_to(&looked[*at], &key)),
                // Exact or the nearest one under it, and 1 the nearest over.
                -1 | 1 => {
                    let mut best: Option<(usize, Value)> = None;
                    for at in order {
                        let candidate = &looked[at];
                        if candidate.is_blank() {
                            continue;
                        }
                        let Ok(side) = compare(candidate, &key) else {
                            continue;
                        };
                        let usable = if how == -1 {
                            side != Ordering::Greater
                        } else {
                            side != Ordering::Less
                        };
                        if !usable {
                            continue;
                        }
                        if side == Ordering::Equal {
                            best = Some((at, candidate.clone()));
                            break;
                        }
                        // The nearest so far on the right side of the key.
                        let nearer = match &best {
                            None => true,
                            Some((_, held)) => match compare(candidate, held) {
                                Ok(Ordering::Greater) => how == -1,
                                Ok(Ordering::Less) => how == 1,
                                _ => false,
                            },
                        };
                        if nearer {
                            best = Some((at, candidate.clone()));
                        }
                    }
                    best.map(|(at, _)| at)
                }
                _ => return Err(ExcelError::Value),
            };
            match found.and_then(|at| taken.get(at).cloned()) {
                Some(value) => Ok(value),
                // The fourth argument is what to say when there is nothing,
                // and without one it is #N/A as any lookup would be.
                // An argument left empty is no answer: measured,
                // XLOOKUP(99, D1:D6, E1:E6, , 1) is #N/A.
                None => match args.get(3) {
                    Some(one) if !matches!(one.scalar(), Value::Blank) => Ok(one.scalar()),
                    _ => Err(ExcelError::NA),
                },
            }
        }

        // The date a given number of WORKING days away: weekends are stepped
        // over, and so is any day named in the third argument.
        //
        // The corpus writes `WORKDAY(date,"")`, which Excel refuses — a text
        // second argument is `#VALUE!` — so getting the refusal right is as
        // much of the answer as getting the arithmetic right.
        "WORKDAY" => {
            let start = serial(&args[0])?;
            let days = num(args.get(1).ok_or(ExcelError::Value)?)? as i64;
            let mut holidays: Vec<i64> = Vec::new();
            if let Some(given) = args.get(2) {
                for one in given.flatten() {
                    if one.is_blank() {
                        continue;
                    }
                    holidays.push(serial(&Arg::Value(one))?);
                }
            }
            let step = if days < 0 { -1 } else { 1 };
            let mut at = start;
            let mut left = days.abs();
            while left > 0 {
                at += step;
                if at < 0 {
                    return Err(ExcelError::Num);
                }
                // Saturday and Sunday are 6 and 7 when Monday is 1.
                if weekday_with_type(at, 2)? >= 6 || holidays.contains(&at) {
                    continue;
                }
                left -= 1;
            }
            Ok(Value::Number(at as f64))
        }

        // Join what you are given with a separator between. Unlike CONCAT it
        // can be told to leave out the blanks, which is the whole point of it:
        // a list of five cells of which two are empty joins with two
        // separators, not four.
        "TEXTJOIN" => {
            if args.len() < 3 {
                return Err(ExcelError::Value);
            }
            let between = text(&args[0])?;
            let skip_blanks = args[1].scalar().to_logical()?;
            let mut pieces: Vec<String> = Vec::new();
            for one in &args[2..] {
                for cell in one.flatten() {
                    if let Value::Error(why) = cell {
                        return Err(why);
                    }
                    let piece = text(&Arg::Value(cell))?;
                    if skip_blanks && piece.is_empty() {
                        continue;
                    }
                    pieces.push(piece);
                }
            }
            Ok(Value::text(pieces.join(&between)))
        }

        // Each word's first letter made a capital and the rest small. A word
        // starts wherever a letter follows something that is not a letter, so
        // `o'neill-smith` becomes `O'Neill-Smith`, which is Excel's answer
        // whatever one thinks of the name.
        "PROPER" => {
            let source = text(one_arg(args)?)?;
            let mut out = String::with_capacity(source.len());
            let mut starting = true;
            for character in source.chars() {
                if starting {
                    out.extend(character.to_uppercase());
                } else {
                    out.extend(character.to_lowercase());
                }
                // A word runs on only through LETTERS. Excel makes
                // "ANNA MARIA 3rd" into "Anna Maria 3Rd" — the r after the
                // digit starts a word as surely as the one after a space.
                starting = !character.is_alphabetic();
            }
            Ok(Value::text(out))
        }

        // The text of what you are given, and nothing at all if it is not
        // text. A number is not text, and neither is a logical.
        "T" => Ok(match one_arg(args)?.scalar() {
            Value::Text(held) => Value::Text(held),
            Value::Error(why) => return Err(why),
            _ => Value::text(""),
        }),

        // Which week of the year a date falls in. The second argument says
        // which day starts a week; 1 (or nothing) is Sunday, 2 is Monday.
        "WEEKNUM" => {
            let serial = serial(&args[0])?;
            let starts = match args.get(1) {
                Some(one) => num(one)? as i64,
                None => 1,
            };
            // Excel's 11 to 17 are Monday through Sunday; 1 and 2 are Sunday
            // and Monday. Everything becomes "how far into the week is Sunday".
            let shift = match starts {
                1 | 17 => 0,
                2 | 11 => 1,
                12 => 2,
                13 => 3,
                14 => 4,
                15 => 5,
                16 => 6,
                21 => return weeknum_iso(serial),
                _ => return Err(ExcelError::Num),
            };
            let year = datetime::date_from_serial(serial)?.year;
            let first = datetime::serial_from_date(year, 1, 1)?;
            // Which day of the week the year opened on, counted from the day
            // the week is taken to start.
            let opened = (weekday_with_type(first, 1)? - 1 - shift).rem_euclid(7);
            Ok(Value::Number(
                ((serial - first + opened) / 7 + 1) as f64,
            ))
        }
        "ISOWEEKNUM" => weeknum_iso(serial(&args[0])?),

        // ---- financial ---------------------------------------------------
        // Straight-line, sum-of-years and the two declining-balance
        // depreciations.
        "SLN" => {
            expect(args, 3)?;
            fin_sln(num(&args[0])?, num(&args[1])?, num(&args[2])?)
        }
        "SYD" => {
            expect(args, 4)?;
            fin_syd(num(&args[0])?, num(&args[1])?, num(&args[2])?, num(&args[3])?)
        }
        "DDB" => {
            if !(4..=5).contains(&args.len()) {
                return Err(ExcelError::Value);
            }
            let factor = match args.get(4) {
                Some(a) => num(a)?,
                None => 2.0,
            };
            fin_ddb(
                num(&args[0])?,
                num(&args[1])?,
                num(&args[2])?,
                num(&args[3])?,
                factor,
            )
        }
        "DB" => {
            if !(4..=5).contains(&args.len()) {
                return Err(ExcelError::Value);
            }
            let month = match args.get(4) {
                Some(a) => num(a)?,
                None => 12.0,
            };
            fin_db(
                num(&args[0])?,
                num(&args[1])?,
                num(&args[2])?,
                num(&args[3])?,
                month,
            )
        }
        // The annuity family. Payment type (end/start of period) is 0 or 1.
        "PMT" => {
            if !(3..=5).contains(&args.len()) {
                return Err(ExcelError::Value);
            }
            fin_pmt(
                num(&args[0])?,
                num(&args[1])?,
                num(&args[2])?,
                fin_optional(args, 3)?,
                fin_kind(args, 4)?,
            )
        }
        "FV" => {
            if !(3..=5).contains(&args.len()) {
                return Err(ExcelError::Value);
            }
            fin_fv(
                num(&args[0])?,
                num(&args[1])?,
                num(&args[2])?,
                fin_optional(args, 3)?,
                fin_kind(args, 4)?,
            )
        }
        "PV" => {
            if !(3..=5).contains(&args.len()) {
                return Err(ExcelError::Value);
            }
            fin_pv(
                num(&args[0])?,
                num(&args[1])?,
                num(&args[2])?,
                fin_optional(args, 3)?,
                fin_kind(args, 4)?,
            )
        }
        // NPER's arguments are (rate, pmt, pv, [fv], [type]).
        "NPER" => {
            if !(3..=5).contains(&args.len()) {
                return Err(ExcelError::Value);
            }
            fin_nper(
                num(&args[0])?,
                num(&args[1])?,
                num(&args[2])?,
                fin_optional(args, 3)?,
                fin_kind(args, 4)?,
            )
        }
        // The interest and principal parts of one payment.
        "IPMT" | "PPMT" => {
            if !(4..=6).contains(&args.len()) {
                return Err(ExcelError::Value);
            }
            let rate = num(&args[0])?;
            let period = num(&args[1])?;
            let periods = num(&args[2])?;
            if period < 1.0 || period > periods {
                return Err(ExcelError::Num);
            }
            let present = num(&args[3])?;
            let future = fin_optional(args, 4)?;
            let kind = fin_kind(args, 5)?;
            let payment = fin_pmt_raw(rate, periods, present, future, kind)?;
            let interest = if kind == 1.0 && period == 1.0 {
                0.0
            } else {
                let balance = fin_fv_raw(rate, period - 1.0, payment, present, kind)?;
                let raw = balance * rate;
                if kind == 1.0 { raw / (1.0 + rate) } else { raw }
            };
            let answer = if name == "IPMT" { interest } else { payment - interest };
            fin_finite(answer)
        }
        // RATE's arguments are (nper, pmt, pv, [fv], [type], [guess]).
        "RATE" => {
            if !(3..=6).contains(&args.len()) {
                return Err(ExcelError::Value);
            }
            let guess = match args.get(5) {
                Some(a) => num(a)?,
                None => 0.1,
            };
            fin_rate(
                num(&args[0])?,
                num(&args[1])?,
                num(&args[2])?,
                fin_optional(args, 3)?,
                fin_kind(args, 4)?,
                guess,
            )
        }
        // NPV discounts a series that follows the first period; unlike the VBA
        // one it does not require both a positive and a negative flow.
        "NPV" => {
            if args.len() < 2 {
                return Err(ExcelError::Value);
            }
            let rate = num(&args[0])?;
            if rate == -1.0 {
                return Err(ExcelError::DivZero);
            }
            let values = numeric_operands(&args[1..])?;
            let base = 1.0 + rate;
            fin_finite(
                values
                    .iter()
                    .enumerate()
                    .map(|(period, flow)| flow / base.powf((period + 1) as f64))
                    .sum(),
            )
        }
        // IRR needs at least one inflow and one outflow, or it is #NUM!.
        "IRR" => {
            if args.is_empty() {
                return Err(ExcelError::Value);
            }
            let values = numeric_operands(&args[..1])?;
            let guess = match args.get(1) {
                Some(a) => num(a)?,
                None => 0.1,
            };
            fin_irr(&values, guess)
        }

        // ---- more lookup / statistics ------------------------------------
        // XMATCH: exact by default (0), or the next smaller (-1) / larger (1),
        // or a wildcard (2); searched forward (1) or backward (-1). Returns
        // the 1-based position.
        "XMATCH" => {
            if args.len() < 2 {
                return Err(ExcelError::Value);
            }
            let key = args[0].scalar();
            if let Some(why) = key.err() {
                return Err(why);
            }
            let hay = args[1].flatten();
            let match_mode = match args.get(2) {
                Some(a) => num(a)? as i32,
                None => 0,
            };
            let search_mode = match args.get(3) {
                Some(a) => num(a)? as i32,
                None => 1,
            };
            let order: Vec<usize> = if search_mode < 0 {
                (0..hay.len()).rev().collect()
            } else {
                (0..hay.len()).collect()
            };
            let found = match match_mode {
                0 => order
                    .into_iter()
                    .find(|&i| !hay[i].is_blank() && crate::value::same_ignoring_width(&hay[i], &key)),
                2 => {
                    let pattern = text(&args[0])?;
                    order.into_iter().find(|&i| match &hay[i] {
                        Value::Text(s) => wildcard_match(s, &pattern),
                        other => compare(other, &key) == Ok(Ordering::Equal),
                    })
                }
                -1 => xmatch_nearest(&hay, &key, &order, true),
                1 => xmatch_nearest(&hay, &key, &order, false),
                _ => return Err(ExcelError::Value),
            };
            found
                .map(|i| Value::Number((i + 1) as f64))
                .ok_or(ExcelError::NA)
        }
        // Build an A1 or R1C1 reference string; abs_num 1..4 fixes row/column.
        "ADDRESS" => {
            if args.len() < 2 {
                return Err(ExcelError::Value);
            }
            let row = num(&args[0])? as i64;
            let col = num(&args[1])? as i64;
            if row < 1 || col < 1 {
                return Err(ExcelError::Value);
            }
            let abs = match args.get(2) {
                Some(a) => num(a)? as i64,
                None => 1,
            };
            let a1 = match args.get(3) {
                Some(a) => a.scalar().to_logical()?,
                None => true,
            };
            let core = if a1 {
                let letters = column_letters(col as u64);
                match abs {
                    1 => format!("${letters}${row}"),
                    2 => format!("{letters}${row}"),
                    3 => format!("${letters}{row}"),
                    4 => format!("{letters}{row}"),
                    _ => return Err(ExcelError::Value),
                }
            } else {
                match abs {
                    1 => format!("R{row}C{col}"),
                    2 => format!("R{row}C[{col}]"),
                    3 => format!("R[{row}]C{col}"),
                    4 => format!("R[{row}]C[{col}]"),
                    _ => return Err(ExcelError::Value),
                }
            };
            let result = match args.get(4) {
                Some(a) => {
                    let sheet = text(a)?;
                    if sheet.is_empty() {
                        core
                    } else if sheet.chars().any(|c| !c.is_alphanumeric() && c != '_')
                        || sheet.chars().next().is_some_and(|c| c.is_ascii_digit())
                    {
                        format!("'{}'!{core}", sheet.replace('\'', "''"))
                    } else {
                        format!("{sheet}!{core}")
                    }
                }
                None => core,
            };
            Ok(Value::text(result))
        }
        // Paired-array sums.
        "SUMXMY2" | "SUMX2MY2" | "SUMX2PY2" => {
            expect(args, 2)?;
            let xs = args[0].flatten();
            let ys = args[1].flatten();
            if let Some(e) = first_error(&args[0..2]) {
                return Err(e);
            }
            if xs.len() != ys.len() {
                return Err(ExcelError::NA);
            }
            let mut sum = 0.0;
            for (x, y) in xs.iter().zip(&ys) {
                if let (Value::Number(a), Value::Number(b)) = (x, y) {
                    sum += match name {
                        "SUMXMY2" => (a - b) * (a - b),
                        "SUMX2MY2" => a * a - b * b,
                        _ => a * a + b * b,
                    };
                }
            }
            Ok(Value::Number(sum))
        }
        // Like RANK, but tied numbers share the average of the places they fill.
        "RANK.AVG" => {
            if args.len() < 2 {
                return Err(ExcelError::Value);
            }
            let wanted = num(&args[0])?;
            let numbers: Vec<f64> = args[1]
                .flatten()
                .iter()
                .filter_map(|one| match one {
                    Value::Number(n) => Some(*n),
                    _ => None,
                })
                .collect();
            if !numbers.contains(&wanted) {
                return Err(ExcelError::NA);
            }
            let up = match args.get(2) {
                Some(one) => num(one)? != 0.0,
                None => false,
            };
            let ahead = numbers
                .iter()
                .filter(|one| if up { **one < wanted } else { **one > wanted })
                .count();
            let ties = numbers.iter().filter(|one| **one == wanted).count();
            Ok(Value::Number(ahead as f64 + 1.0 + (ties as f64 - 1.0) / 2.0))
        }
        // Mean absolute deviation, and the sum of squared deviations.
        "AVEDEV" | "DEVSQ" => {
            let numbers = numeric_operands(args)?;
            if numbers.is_empty() {
                return Err(ExcelError::Num);
            }
            let mean = numbers.iter().sum::<f64>() / numbers.len() as f64;
            if name == "AVEDEV" {
                let total: f64 = numbers.iter().map(|n| (n - mean).abs()).sum();
                Ok(Value::Number(total / numbers.len() as f64))
            } else {
                Ok(Value::Number(numbers.iter().map(|n| (n - mean) * (n - mean)).sum()))
            }
        }
        // (a+b+...)! / (a! b! ...), on the truncated arguments.
        "MULTINOMIAL" => {
            let numbers = numeric_operands(args)?;
            let mut sum = 0.0;
            let mut denominator = 1.0;
            for n in &numbers {
                if *n < 0.0 {
                    return Err(ExcelError::Num);
                }
                let whole = n.trunc();
                sum += whole;
                denominator *= factorial(whole)?;
            }
            Ok(Value::Number(factorial(sum)? / denominator))
        }
        // Where a value falls in a set, 0..1, truncated to some significant
        // digits (3 by default).
        "PERCENTRANK" | "PERCENTRANK.INC" => {
            if args.len() < 2 {
                return Err(ExcelError::Value);
            }
            let mut numbers: Vec<f64> = args[0]
                .flatten()
                .iter()
                .filter_map(|one| match one {
                    Value::Number(n) => Some(*n),
                    _ => None,
                })
                .collect();
            if numbers.is_empty() {
                return Err(ExcelError::Num);
            }
            numbers.sort_by(|a, b| a.partial_cmp(b).unwrap_or(Ordering::Equal));
            let x = num(&args[1])?;
            if x < numbers[0] || x > numbers[numbers.len() - 1] {
                return Err(ExcelError::NA);
            }
            let significance = match args.get(2) {
                Some(a) => num(a)? as i32,
                None => 3,
            };
            if significance < 1 {
                return Err(ExcelError::Num);
            }
            let n = numbers.len();
            let rank = if n == 1 {
                1.0
            } else {
                let mut position = None;
                for i in 0..n - 1 {
                    if (numbers[i] - x).abs() < f64::EPSILON {
                        position = Some(i as f64);
                        break;
                    }
                    if numbers[i] < x && x < numbers[i + 1] {
                        position = Some(i as f64 + (x - numbers[i]) / (numbers[i + 1] - numbers[i]));
                        break;
                    }
                }
                let position = position.unwrap_or((n - 1) as f64);
                position / (n as f64 - 1.0)
            };
            Ok(Value::Number(truncate_significant(rank, significance)))
        }
        // The mean after trimming a fraction from both ends (rounded down to a
        // whole pair excluded).
        "TRIMMEAN" => {
            expect(args, 2)?;
            let percent = num(&args[1])?;
            if !(0.0..1.0).contains(&percent) {
                return Err(ExcelError::Num);
            }
            let mut numbers: Vec<f64> = args[0]
                .flatten()
                .iter()
                .filter_map(|one| match one {
                    Value::Number(n) => Some(*n),
                    _ => None,
                })
                .collect();
            if numbers.is_empty() {
                return Err(ExcelError::Num);
            }
            numbers.sort_by(|a, b| a.partial_cmp(b).unwrap_or(Ordering::Equal));
            let excluded = ((numbers.len() as f64 * percent).floor() as usize / 2) * 2;
            let each = excluded / 2;
            let kept = &numbers[each..numbers.len() - each];
            Ok(Value::Number(kept.iter().sum::<f64>() / kept.len() as f64))
        }

        "VLOOKUP" | "HLOOKUP" => {
            if args.len() < 3 {
                return Err(ExcelError::Value);
            }
            // The needle is taken as the FIRST of whatever it was given rather
            // than one answer per needle -- see `Arg::first`. MATCH, below,
            // does the opposite.
            let key = args[0].first();
            // Looking for an error finds nothing: the error is the answer.
            if let Some(why) = key.err() {
                return Err(why);
            }
            let table = args[1].as_range();
            let index = num(&args[2])? as usize;
            if index < 1 {
                return Err(ExcelError::Value);
            }
            let approximate = match args.get(3) {
                Some(a) => a.scalar().to_logical().unwrap_or(true),
                None => true,
            };
            let vertical = name == "VLOOKUP";
            let lanes = if vertical { table.height } else { table.width };
            let depth = if vertical { table.width } else { table.height };
            if index > depth {
                return Err(ExcelError::Ref);
            }

            let probe = |i: usize| {
                if vertical {
                    table.at(0, i)
                } else {
                    table.at(i, 0)
                }
            };
            let fetch = |i: usize| {
                if vertical {
                    table.at(index - 1, i)
                } else {
                    table.at(i, index - 1)
                }
            };

            if approximate {
                sorted_position(lanes, false, probe, &key).map(fetch).ok_or(ExcelError::NA)
            } else {
                // An empty cell in the lookup column never matches, not even an
                // empty lookup value: Excel reports #N/A rather than pairing two
                // blanks. Without this, looking up an unfilled cell silently
                // returns whatever sits beside the first gap in the table.
                (0..lanes)
                    .find(|&i| answers_to(&probe(i), &key))
                    .map(fetch)
                    .ok_or(ExcelError::NA)
            }
        }
        "MATCH" => {
            if args.len() < 2 {
                return Err(ExcelError::Value);
            }
            let key = args[0].scalar();
            // Looking for an error finds nothing: the error is the answer.
            if let Some(why) = key.err() {
                return Err(why);
            }
            let haystack = args[1].flatten();
            let mode = match args.get(2) {
                Some(a) => num(a)? as i32,
                None => 1,
            };
            let found = match mode {
                0 => haystack.iter().position(|v| answers_to(v, &key)),
                m => sorted_position(haystack.len(), m < 0, |i| haystack[i].clone(), &key),
            };
            found
                .map(|i| Value::Number((i + 1) as f64))
                .ok_or(ExcelError::NA)
        }
        "INDEX" => {
            if args.len() < 2 {
                return Err(ExcelError::Value);
            }
            let table = args[0].as_range();
            let row = num(&args[1])? as usize;
            let col = match args.get(2) {
                Some(a) => num(a)? as usize,
                None => {
                    // With one index, a single row or column is addressed linearly.
                    if table.height == 1 {
                        return index_at(&table, 1, row);
                    }
                    if table.width == 1 {
                        return index_at(&table, row, 1);
                    }
                    return Err(ExcelError::Ref);
                }
            };
            index_at(&table, row, col)
        }

        // ---- misc ---------------------------------------------------------
        // The link target is metadata; the value of the cell is what it shows.
        "HYPERLINK" => {
            let link = text(one_arg(args)?)?;
            Ok(match args.get(1) {
                Some(friendly) => friendly.scalar(),
                None => Value::Text(link),
            })
        }
        "CHOOSE" => {
            if args.len() < 2 {
                return Err(ExcelError::Value);
            }
            // `num` hands an error straight back, so an error for the number
            // saying which one to take is the answer.
            let index = num(&args[0])? as usize;
            if index < 1 || index >= args.len() {
                return Err(ExcelError::Value);
            }
            Ok(args[index].scalar())
        }
        // The first condition that holds gives its paired result; a condition
        // that errors is the answer; nothing true is #N/A. A non-taken
        // result's error is ignored (IFS is error-transparent, as IF is).
        "IFS" => {
            if args.is_empty() {
                return Err(ExcelError::Value);
            }
            let mut i = 0;
            while i + 1 < args.len() {
                let cond = args[i].scalar();
                if let Some(e) = cond.err() {
                    return Err(e);
                }
                if cond.to_logical()? {
                    return Ok(args[i + 1].scalar());
                }
                i += 2;
            }
            Err(ExcelError::NA)
        }
        // The first value equal to the subject gives its paired result; a
        // leftover final argument is the default; no match and no default is
        // #N/A. Equality is Excel's `=`, so text matches case-insensitively.
        "SWITCH" => {
            if args.len() < 3 {
                return Err(ExcelError::Value);
            }
            let subject = args[0].scalar();
            if let Some(e) = subject.err() {
                return Err(e);
            }
            let mut i = 1;
            while i + 1 < args.len() {
                let candidate = args[i].scalar();
                if let Some(e) = candidate.err() {
                    return Err(e);
                }
                if compare(&subject, &candidate) == Ok(Ordering::Equal) {
                    return Ok(args[i + 1].scalar());
                }
                i += 2;
            }
            if i < args.len() {
                Ok(args[i].scalar())
            } else {
                Err(ExcelError::NA)
            }
        }
        // Codes 1..=11 include manually hidden rows, 101..=111 exclude them.
        // Row visibility is not modelled here, so both behave the same; the
        // difference only shows up on a sheet with hidden rows.
        "SUBTOTAL" => {
            if args.len() < 2 {
                return Err(ExcelError::Value);
            }
            let inner = match num(&args[0])? as i64 % 100 {
                1 => "AVERAGE",
                2 => "COUNT",
                3 => "COUNTA",
                4 => "MAX",
                5 => "MIN",
                6 => "PRODUCT",
                7 => "STDEV.S",
                8 => "STDEV.P",
                9 => "SUM",
                10 => "VAR.S",
                11 => "VAR.P",
                _ => return Err(ExcelError::Value),
            };
            dispatch(inner, &args[1..])
        }

        // ---- statistics -------------------------------------------------
        // The middle value, or the mean of the two in the middle when there is
        // no single one.
        "MEDIAN" => {
            let mut held = numeric_operands(args)?;
            if held.is_empty() {
                return Err(ExcelError::Num);
            }
            held.sort_by(|a, b| a.partial_cmp(b).unwrap_or(Ordering::Equal));
            let middle = held.len() / 2;
            Ok(Value::Number(if held.len() % 2 == 1 {
                held[middle]
            } else {
                (held[middle - 1] + held[middle]) / 2.0
            }))
        }

        // How far the values lie from their mean. The `.S` forms divide by one
        // less than the count, taking the values for a sample of something
        // larger; the `.P` forms divide by the count, taking them for the whole
        // of it.
        "STDEV" | "STDEV.S" | "STDEVP" | "STDEV.P" | "VAR" | "VAR.S" | "VARP" | "VAR.P" => {
            let held = numeric_operands(args)?;
            let whole = matches!(name, "STDEVP" | "STDEV.P" | "VARP" | "VAR.P");
            let divisor = if whole {
                held.len() as f64
            } else {
                held.len() as f64 - 1.0
            };
            if divisor <= 0.0 {
                return Err(ExcelError::DivZero);
            }
            let mean = held.iter().sum::<f64>() / held.len() as f64;
            let spread = held.iter().map(|one| (one - mean).powi(2)).sum::<f64>() / divisor;
            Ok(Value::Number(if name.starts_with("STDEV") {
                spread.sqrt()
            } else {
                spread
            }))
        }

        // The value that turns up most often. One that turns up no more often
        // than any other is not a mode at all.
        "MODE" | "MODE.SNGL" => {
            let held = numeric_operands(args)?;
            let mut best: Option<(f64, usize)> = None;
            for one in &held {
                let times = held.iter().filter(|other| *other == one).count();
                if times < 2 {
                    continue;
                }
                // Walking in order and refusing to replace on a tie keeps the
                // earliest of the values that turn up equally often.
                match best {
                    Some((_, seen)) if seen >= times => {}
                    _ => best = Some((*one, times)),
                }
            }
            match best {
                Some((one, _)) => Ok(Value::Number(one)),
                None => Err(ExcelError::NA),
            }
        }

        // The value a given way along the sorted list. The two families differ
        // in where they start counting: INC from the first value, EXC from
        // before it, which is why the same quarter comes out differently.
        "PERCENTILE" | "PERCENTILE.INC" | "PERCENTILE.EXC" | "QUARTILE" | "QUARTILE.INC"
        | "QUARTILE.EXC" => {
            if args.len() < 2 {
                return Err(ExcelError::Value);
            }
            let mut held = numeric_operands(&args[..args.len() - 1])?;
            if held.is_empty() {
                return Err(ExcelError::Num);
            }
            held.sort_by(|a, b| a.partial_cmp(b).unwrap_or(Ordering::Equal));
            let asked = num(&args[args.len() - 1])?;
            // A quartile is a percentile in quarters.
            let part = if name.starts_with("QUARTILE") {
                if !(0.0..=4.0).contains(&asked) {
                    return Err(ExcelError::Num);
                }
                asked.trunc() / 4.0
            } else {
                asked
            };
            let excluding = name.ends_with(".EXC");
            let count = held.len() as f64;
            let place = if excluding {
                part * (count + 1.0) - 1.0
            } else {
                part * (count - 1.0)
            };
            if !(0.0..=count - 1.0).contains(&place) {
                return Err(ExcelError::Num);
            }
            let below = place.floor() as usize;
            let above = (below + 1).min(held.len() - 1);
            let along = place - below as f64;
            Ok(Value::Number(held[below] + (held[above] - held[below]) * along))
        }

        // SUBTOTAL's successor: the same aggregations, and a second argument
        // saying what to leave out of them.
        "AGGREGATE" => {
            if args.len() < 2 {
                return Err(ExcelError::Value);
            }
            let which = num(&args[0])? as i64;
            let leaving_out = num(&args[1])? as i64;
            if !(0..=7).contains(&leaving_out) {
                return Err(ExcelError::Value);
            }
            let inner = match which {
                1 => "AVERAGE",
                2 => "COUNT",
                3 => "COUNTA",
                4 => "MAX",
                5 => "MIN",
                6 => "PRODUCT",
                7 => "STDEV.S",
                8 => "STDEV.P",
                9 => "SUM",
                10 => "VAR.S",
                11 => "VAR.P",
                12 => "MEDIAN",
                13 => "MODE.SNGL",
                14 => "LARGE",
                15 => "SMALL",
                16 => "PERCENTILE.INC",
                17 => "QUARTILE.INC",
                18 => "PERCENTILE.EXC",
                19 => "QUARTILE.EXC",
                _ => return Err(ExcelError::Value),
            };
            // 14 to 19 want a k, or a fraction, after the values.
            let wants_k = (14..=19).contains(&which);
            let rest = &args[2..];
            if rest.is_empty() || (wants_k && rest.len() < 2) {
                return Err(ExcelError::Value);
            }
            let (values, k) = if wants_k {
                rest.split_at(rest.len() - 1)
            } else {
                (rest, &rest[..0])
            };
            // Options 2, 3, 6 and 7 pass over the errors. The rest do not, and
            // an error in the values is then the answer, as it would be for the
            // aggregation on its own.
            let mut passed: Vec<Arg> = if matches!(leaving_out, 2 | 3 | 6 | 7) {
                let kept: Vec<Value> = values
                    .iter()
                    .flat_map(|one| one.flatten())
                    .filter(|held| !held.is_error())
                    .collect();
                vec![Arg::Range(RangeData {
                    width: 1,
                    height: kept.len(),
                    cells: kept,
                })]
            } else {
                if let Some(why) = first_error(values) {
                    return Err(why);
                }
                values.to_vec()
            };
            passed.extend(k.iter().cloned());
            dispatch(inner, &passed)
        }

        // ---- date and time ---------------------------------------------
        "DATE" => {
            expect(args, 3)?;
            let s = datetime::serial_from_date(
                num(&args[0])? as i64,
                num(&args[1])? as i64,
                num(&args[2])? as i64,
            )?;
            Ok(Value::Number(s as f64))
        }
        // Each part is cut to a whole number no greater than 32767, and the
        // clock they add up to may not run backwards: measured, TIME(0,-1,0)
        // and TIME(32768,0,0) are #NUM! where TIME(25,0,0) is 1/24.
        "TIME" => {
            expect(args, 3)?;
            let (hours, minutes, seconds) = (num(&args[0])?.trunc(), num(&args[1])?.trunc(), num(&args[2])?.trunc());
            if [hours, minutes, seconds].iter().any(|part| *part > 32_767.0)
                || hours * 3600.0 + minutes * 60.0 + seconds < 0.0
            {
                return Err(ExcelError::Num);
            }
            Ok(Value::Number(datetime::fraction_from_time(hours, minutes, seconds)))
        }
        "YEAR" => Ok(Value::Number(
            datetime::date_from_serial(serial(one_arg(args)?)?)?.year as f64,
        )),
        "MONTH" => Ok(Value::Number(
            datetime::date_from_serial(serial(one_arg(args)?)?)?.month as f64,
        )),
        "DAY" => Ok(Value::Number(
            datetime::date_from_serial(serial(one_arg(args)?)?)?.day as f64,
        )),
        "HOUR" | "MINUTE" | "SECOND" => {
            // Measured: MINUTE(-0.5) is #NUM!.
            let moment = one(args)?;
            if moment < 0.0 {
                return Err(ExcelError::Num);
            }
            let (h, m, s) = datetime::time_from_fraction(moment);
            Ok(Value::Number(match name {
                "HOUR" => h as f64,
                "MINUTE" => m as f64,
                _ => s as f64,
            }))
        }
        "WEEKDAY" => {
            let kind = match args.get(1) {
                Some(a) => num(a)? as i64,
                None => 1,
            };
            Ok(Value::Number(
                weekday_with_type(serial(one_arg(args)?)?, kind)? as f64,
            ))
        }
        "EDATE" => {
            expect(args, 2)?;
            Ok(Value::Number(
                datetime::add_months(serial(&args[0])?, num(&args[1])? as i64)? as f64,
            ))
        }
        "EOMONTH" => {
            expect(args, 2)?;
            Ok(Value::Number(
                datetime::end_of_month(serial(&args[0])?, num(&args[1])? as i64)? as f64,
            ))
        }
        "DAYS" => {
            expect(args, 2)?;
            Ok(Value::Number((serial(&args[0])? - serial(&args[1])?) as f64))
        }
        // The 360-day count. In the US (NASD) form a last-of-February start
        // counts as the 30th, then a 31st start becomes the 30th, and a 31st
        // end becomes the 30th only once the start is on the 30th. The
        // European form just pulls any 31 down to 30. Measured: 29 Feb 2024
        // -> 31 Mar 2024 is 30 in the US form and 31 in the European.
        "DAYS360" => {
            if !(2..=3).contains(&args.len()) {
                return Err(ExcelError::Value);
            }
            let start = serial(&args[0])?;
            let end = serial(&args[1])?;
            let european = match args.get(2) {
                Some(a) => a.scalar().to_logical()?,
                None => false,
            };
            Ok(Value::Number(days360(start, end, european)? as f64))
        }
        // Read a date out of text; the time part, if any, is dropped.
        "DATEVALUE" => {
            let s = text(&args[0])?;
            match datetime::text_as_datetime(&s) {
                // A time with no date is day 0: measured, DATEVALUE("12:00")
                // is 0.
                Some(serial) if serial >= 0.0 => Ok(Value::Number(serial.floor())),
                _ => Err(ExcelError::Value),
            }
        }
        // The time of day a text names, with any date dropped and a clock
        // past 24 hours taken round again. Measured: "2024/1/1 6:00" is 0.25,
        // "25:00" is 1/24, a date alone is 0, and a number is #VALUE!.
        "TIMEVALUE" => {
            if matches!(args[0].scalar(), Value::Number(_)) {
                return Err(ExcelError::Value);
            }
            let s = text(&args[0])?;
            match datetime::text_as_datetime(&s) {
                // Round the clock in seconds, not days, so "25:00" is exactly
                // an hour rather than a day's worth of rounding short of one.
                // With a date in front the time is what is left of the
                // whole once the day is taken off, rounding and all:
                // measured, "2024/3/5 25:00" is 0.0416666666642413.
                // So is a time past a day: measured, "25:00:00.5" is
                // 0.0416724537037036.
                Some(serial) if serial >= 1.0 => Ok(Value::Number(serial - serial.floor())),
                Some(serial) => Ok(Value::Number((serial * 86_400.0).rem_euclid(86_400.0) / 86_400.0)),
                None => Err(ExcelError::Value),
            }
        }
        // The fraction of a year between two dates, on one of five day-count
        // bases. Symmetric in its dates. All five verified against Excel.
        "YEARFRAC" => {
            if !(2..=3).contains(&args.len()) {
                return Err(ExcelError::Value);
            }
            let basis = match args.get(2) {
                Some(a) => num(a)? as i64,
                None => 0,
            };
            Ok(Value::Number(yearfrac(serial(&args[0])?, serial(&args[1])?, basis)?))
        }
        "DATEDIF" => {
            expect(args, 3)?;
            let unit = text(&args[2])?;
            Ok(Value::Number(datedif(
                serial(&args[0])?,
                serial(&args[1])?,
                &unit,
            )?))
        }

        // ---- more aggregates ---------------------------------------------
        "SUMSQ" => Ok(Value::Number(numeric_operands(args)?.iter().map(|n| n * n).sum())),
        "GCD" => {
            let mut g = 0i64;
            for n in numeric_operands(args)? {
                if n < 0.0 {
                    return Err(ExcelError::Num);
                }
                g = gcd(g, n.trunc() as i64);
            }
            Ok(Value::Number(g as f64))
        }
        "LCM" => {
            let mut l = 1i64;
            for n in numeric_operands(args)? {
                if n < 0.0 {
                    return Err(ExcelError::Num);
                }
                let n = n.trunc() as i64;
                if n == 0 {
                    return Ok(Value::Number(0.0));
                }
                let g = gcd(l, n);
                l = l / g * n;
            }
            Ok(Value::Number(l as f64))
        }
        // ---- more one-number maths --------------------------------------
        "EVEN" | "ODD" => {
            let n = num(&args.first().ok_or(ExcelError::Value)?.clone())?;
            let mut up = n.abs().ceil() as i64;
            let want_even = name == "EVEN";
            if (up % 2 == 0) != want_even {
                up += 1;
            }
            Ok(Value::Number(if n < 0.0 { -up } else { up } as f64))
        }
        "ROMAN" => {
            let n = num(&args.first().ok_or(ExcelError::Value)?.clone())? as i64;
            if !(0..=3999).contains(&n) {
                return Err(ExcelError::Value);
            }
            Ok(Value::Text(roman_numeral(n)))
        }
        "ARABIC" => {
            let s = text(&args.first().ok_or(ExcelError::Value)?.clone())?;
            arabic_number(&s).map(|n| Value::Number(n as f64)).ok_or(ExcelError::Value)
        }
        // ---- more text ---------------------------------------------------
        "CLEAN" => {
            let s = text(&args.first().ok_or(ExcelError::Value)?.clone())?;
            Ok(Value::Text(s.chars().filter(|c| (*c as u32) >= 32).collect()))
        }
        "FIXED" | "DOLLAR" => {
            let n = num(&args.first().ok_or(ExcelError::Value)?.clone())?;
            let digits = match args.get(1) {
                Some(a) => num(a)? as i32,
                None => 2,
            };
            // A negative digit count rounds to the left of the point, the way
            // ROUND does; the shown number carries no decimals then.
            let places = digits.max(0) as usize;
            let no_commas = name == "FIXED"
                && matches!(args.get(2), Some(a) if a.scalar().to_logical().unwrap_or(false));
            let rounded = {
                let factor = 10f64.powi(digits);
                (n * factor).round() / factor
            };
            let dp = if name == "DOLLAR" { places.max(0) } else { places };
            let group = if no_commas { "" } else { "#,##" };
            let body = if dp == 0 {
                format!("{group}0")
            } else {
                format!("{group}0.{}", "0".repeat(dp))
            };
            let format = if name == "DOLLAR" {
                format!("\"$\"{body};(\"$\"{body})")
            } else {
                body
            };
            Ok(Value::Text(crate::numfmt::format_number(rounded, &format)))
        }
        // ---- lookup ------------------------------------------------------
        "LOOKUP" => {
            // The vector form: find the largest value not over the needle in
            // an ascending vector, and answer the matching cell of the
            // result vector -- or of the same vector, with none given.
            let needle = args.first().ok_or(ExcelError::Value)?.scalar();
            // The array form, a block and no result vector: looked up down
            // its first column and answered from its last when it is at least
            // as tall as it is wide, else along its first row and answered
            // from its last. Measured: LOOKUP(99, D1:E6) is E6.
            let block = args.get(1).ok_or(ExcelError::Value)?.as_range();
            let (vector, result) = match args.get(2) {
                Some(a) => (args[1].flatten(), a.flatten()),
                None if block.width > 1 && block.height > 1 => {
                    let (width, height) = (block.width, block.height);
                    let cell = |row: usize, col: usize| block.cells[row * width + col].clone();
                    if height >= width {
                        ((0..height).map(|row| cell(row, 0)).collect::<Vec<_>>(), (0..height).map(|row| cell(row, width - 1)).collect())
                    } else {
                        ((0..width).map(|col| cell(0, col)).collect::<Vec<_>>(), (0..width).map(|col| cell(height - 1, col)).collect())
                    }
                }
                None => {
                    let vector = args[1].flatten();
                    (vector.clone(), vector)
                }
            };
            let found = sorted_position(vector.len(), false, |i| vector[i].clone(), &needle);
            match found {
                Some(at) => Ok(result.get(at).cloned().unwrap_or(Value::Error(ExcelError::NA))),
                None => Err(ExcelError::NA),
            }
        }
        // ---- database ----------------------------------------------------
        "DSUM" | "DAVERAGE" | "DCOUNT" | "DCOUNTA" | "DMAX" | "DMIN" | "DGET" | "DPRODUCT" | "DSTDEV" | "DSTDEVP"
        | "DVAR" | "DVARP" => database_function(name, args),
        name if crate::functions_more::NAMES.contains(&name) => crate::functions_more::call(name, args),
        name if crate::complex::NAMES.contains(&name) => crate::complex::call(name, args),
        name if crate::bonds::NAMES.contains(&name) => crate::bonds::call(name, args),
        "BESSELJ" | "BESSELY" | "BESSELI" | "BESSELK" => {
            if args.len() != 2 {
                return Err(ExcelError::Value);
            }
            let (x, order) = (num(&args[0])?, num(&args[1])?.trunc());
            if order < 0.0 || (matches!(name, "BESSELY" | "BESSELK") && x <= 0.0) {
                return Err(ExcelError::Num);
            }
            let order = order as i64;
            let answer = match name {
                "BESSELJ" => crate::bessel::bessel_j(order, x),
                "BESSELY" => crate::bessel::bessel_y(order, x),
                "BESSELI" => crate::bessel::bessel_i(order, x),
                _ => crate::bessel::bessel_k(order, x),
            };
            if answer.is_finite() { Ok(Value::Number(answer)) } else { Err(ExcelError::Num) }
        }
        "CONVERT" => {
            if args.len() != 3 {
                return Err(ExcelError::Value);
            }
            crate::convert::convert(num(&args[0])?, &text(&args[1])?, &text(&args[2])?)
        }
        // ---- more dates --------------------------------------------------
        // The working days between two dates, both counted, less the weekend
        // and less the holidays given -- which the plain form used to ignore.
        "NETWORKDAYS" | "NETWORKDAYS.INTL" => {
            let intl = name == "NETWORKDAYS.INTL";
            let (start, end) = (serial(&args[0])?, serial(&args[1])?);
            let weekend = weekend_days(if intl { args.get(2) } else { None })?;
            let holidays = holiday_serials(args.get(if intl { 3 } else { 2 }))?;
            let (lo, hi) = if start <= end { (start, end) } else { (end, start) };
            let mut days = 0i64;
            for day in lo..=hi {
                if !weekend[monday_zero(day)?] && !holidays.contains(&day) {
                    days += 1;
                }
            }
            Ok(Value::Number(if start <= end { days } else { -days } as f64))
        }
        "WORKDAY.INTL" => {
            let start = serial(&args[0])?;
            let days = num(args.get(1).ok_or(ExcelError::Value)?)? as i64;
            let weekend = weekend_days(args.get(2))?;
            // With no working day in the week there is no day to land on.
            if weekend.iter().all(|off| *off) {
                return Err(ExcelError::Value);
            }
            let holidays = holiday_serials(args.get(3))?;
            let step = if days < 0 { -1 } else { 1 };
            let mut at = start;
            let mut left = days.abs();
            while left > 0 {
                at += step;
                if at < 0 {
                    return Err(ExcelError::Num);
                }
                if weekend[monday_zero(at)?] || holidays.contains(&at) {
                    continue;
                }
                left -= 1;
            }
            Ok(Value::Number(at as f64))
        }
        // Byte-counting text functions: a character is one byte when it is
        // ASCII or half-width katakana and two otherwise, as in Shift_JIS;
        // half of a two-byte character cut off is a space. Measured:
        // LENB("東京abc") is 7, LEFTB(,3) "東 ", RIGHTB(,4) " abc",
        // MIDB(,2,3) " 京", FINDB("a",) 5, SEARCHB("B",) 6.
        "LENB" => Ok(Value::Number(bytes_of(&text(one_arg(args)?)?) as f64)),
        "LEFTB" | "RIGHTB" => {
            let t = text(one_arg(args)?)?;
            let n = match args.get(1) {
                Some(a) => num(a)?.trunc(),
                None => 1.0,
            };
            if n < 0.0 {
                return Err(ExcelError::Value);
            }
            let total = bytes_of(&t);
            let n = (n as usize).min(total);
            Ok(Value::Text(if name == "LEFTB" {
                bytes_between(&t, 1, n)
            } else {
                bytes_between(&t, total - n + 1, n)
            }))
        }
        "MIDB" => {
            expect(args, 3)?;
            let t = text(&args[0])?;
            let (start, n) = (num(&args[1])?.trunc(), num(&args[2])?.trunc());
            if start < 1.0 || n < 0.0 {
                return Err(ExcelError::Value);
            }
            Ok(Value::Text(bytes_between(&t, start as usize, n as usize)))
        }
        "REPLACEB" => {
            expect(args, 4)?;
            let t = text(&args[0])?;
            let (start, n) = (num(&args[1])?.trunc(), num(&args[2])?.trunc());
            if start < 1.0 || n < 0.0 {
                return Err(ExcelError::Value);
            }
            let (start, n) = (start as usize, n as usize);
            let total = bytes_of(&t);
            let head = bytes_between(&t, 1, start - 1);
            let tail_from = start + n;
            let tail = if tail_from > total {
                String::new()
            } else {
                bytes_between(&t, tail_from, total - tail_from + 1)
            };
            Ok(Value::Text(format!("{head}{}{tail}", text(&args[3])?)))
        }
        "FINDB" | "SEARCHB" => {
            expect(args, 2)?;
            let within = text(&args[1])?;
            let start_byte = match args.get(2) {
                Some(a) => num(a)?.trunc(),
                None => 1.0,
            };
            if start_byte < 1.0 {
                return Err(ExcelError::Value);
            }
            // The character the starting byte falls in or before.
            let mut seen = 0usize;
            let mut start_char = within.chars().count() + 1;
            for (at, c) in within.chars().enumerate() {
                if seen + 1 >= start_byte as usize {
                    start_char = at + 1;
                    break;
                }
                seen += byte_width(c);
            }
            let found = dispatch(
                if name == "FINDB" { "FIND" } else { "SEARCH" },
                &[args[0].clone(), args[1].clone(), Arg::Value(Value::Number(start_char as f64))],
            )?;
            let Value::Number(at) = found else {
                return Ok(found);
            };
            let before: String = within.chars().take(at as usize - 1).collect();
            Ok(Value::Number((bytes_of(&before) + 1) as f64))
        }
        "ISNONTEXT" => Ok(Value::Logical(!matches!(one_value(args), Value::Text(_)))),
        "DELTA" => {
            let a = one(args)?;
            let b = match args.get(1) {
                Some(x) => num(x)?,
                None => 0.0,
            };
            Ok(Value::Number(if a == b { 1.0 } else { 0.0 }))
        }
        "GESTEP" => {
            let a = one(args)?;
            let step = match args.get(1) {
                Some(x) => num(x)?,
                None => 0.0,
            };
            Ok(Value::Number(if a >= step { 1.0 } else { 0.0 }))
        }
        "BITXOR" => {
            expect(args, 2)?;
            let a = num(&args[0])?.trunc();
            let b = num(&args[1])?.trunc();
            let limit = 281_474_976_710_655.0;
            if a < 0.0 || b < 0.0 || a > limit || b > limit {
                return Err(ExcelError::Num);
            }
            Ok(Value::Number(((a as u64) ^ (b as u64)) as f64))
        }
        // MIRR: the positive flows carried forward at the reinvestment rate,
        // the negative ones brought back at the finance rate.
        "MIRR" => {
            expect(args, 3)?;
            let flows: Vec<f64> = args[0]
                .flatten()
                .into_iter()
                .filter_map(|v| if let Value::Number(n) = v { Some(n) } else { None })
                .collect();
            let (finance, reinvest) = (num(&args[1])?, num(&args[2])?);
            let n = flows.len();
            if n < 2 {
                return Err(ExcelError::DivZero);
            }
            let mut future = 0.0;
            let mut present = 0.0;
            for (at, flow) in flows.iter().enumerate() {
                if *flow > 0.0 {
                    future += flow * (1.0 + reinvest).powi((n - 1 - at) as i32);
                } else {
                    present += flow / (1.0 + finance).powi(at as i32);
                }
            }
            if future == 0.0 || present == 0.0 {
                return Err(ExcelError::DivZero);
            }
            fin_finite((future / -present).powf(1.0 / (n as f64 - 1.0)) - 1.0)
        }
        _ => Err(ExcelError::Name),
    }
}

// ---- financial helpers ----------------------------------------------------
// Ported from the VBA runtime's annuity and depreciation math, which was
// measured against Excel; failure is #NUM! here rather than a raised error.

fn fin_finite(value: f64) -> Result<Value, ExcelError> {
    if value.is_finite() {
        Ok(Value::Number(if value == 0.0 { 0.0 } else { value }))
    } else {
        Err(ExcelError::Num)
    }
}

fn fin_optional(args: &[Arg], index: usize) -> Result<f64, ExcelError> {
    match args.get(index) {
        None => Ok(0.0),
        Some(a) => num(a),
    }
}

fn fin_kind(args: &[Arg], index: usize) -> Result<f64, ExcelError> {
    let value = fin_optional(args, index)?;
    if value == 0.0 || value == 1.0 {
        Ok(value)
    } else {
        Err(ExcelError::Num)
    }
}

fn annuity_factor(rate: f64, periods: f64) -> Result<f64, ExcelError> {
    if rate <= -1.0 {
        return Err(ExcelError::Num);
    }
    // A whole number of periods is raised by squaring and multiplying, the
    // way Excel's own figures come out to the last digit: measured,
    // `FV(0.004,120,-300,-5000,1)` is 54346.5852341411 that way, where
    // `powf` gives ...408; PV, PMT and FV agreed on every case tried.
    let factor = if periods.fract() == 0.0 && periods.abs() < 2_147_483_648.0 {
        let mut base = 1.0 + rate;
        let mut left = periods.abs() as u64;
        let mut answer = 1.0;
        while left > 0 {
            if left & 1 == 1 {
                answer *= base;
            }
            base *= base;
            left >>= 1;
        }
        if periods < 0.0 { 1.0 / answer } else { answer }
    } else {
        (1.0 + rate).powf(periods)
    };
    if factor.is_finite() { Ok(factor) } else { Err(ExcelError::Num) }
}

pub(crate) fn fin_fv_raw(rate: f64, periods: f64, payment: f64, present: f64, kind: f64) -> Result<f64, ExcelError> {
    if rate == 0.0 {
        return Ok(-(present + payment * periods));
    }
    // At -100% FV still has an answer, everything gone after the first
    // period: measured, FV(-1,2,1) is -1 (while PMT(-1,10,1000) is #NUM!).
    let factor = if rate == -1.0 && periods > 0.0 { 0.0 } else { annuity_factor(rate, periods)? };
    Ok(-(present * factor + payment * (1.0 + rate * kind) * (factor - 1.0) / rate))
}

pub(crate) fn fin_pmt_raw(rate: f64, periods: f64, present: f64, future: f64, kind: f64) -> Result<f64, ExcelError> {
    if periods == 0.0 {
        return Err(ExcelError::Num);
    }
    if rate == 0.0 {
        return Ok(-(future + present) / periods);
    }
    let factor = annuity_factor(rate, periods)?;
    let denominator = (1.0 + rate * kind) * (factor - 1.0);
    if denominator == 0.0 {
        return Err(ExcelError::Num);
    }
    Ok(-(future + present * factor) * rate / denominator)
}

fn fin_fv(rate: f64, periods: f64, payment: f64, present: f64, kind: f64) -> Result<Value, ExcelError> {
    fin_finite(fin_fv_raw(rate, periods, payment, present, kind)?)
}

fn fin_pv(rate: f64, periods: f64, payment: f64, future: f64, kind: f64) -> Result<Value, ExcelError> {
    let value = if rate == 0.0 {
        -(future + payment * periods)
    } else {
        let factor = annuity_factor(rate, periods)?;
        -(future + payment * (1.0 + rate * kind) * (factor - 1.0) / rate) / factor
    };
    fin_finite(value)
}

fn fin_pmt(rate: f64, periods: f64, present: f64, future: f64, kind: f64) -> Result<Value, ExcelError> {
    fin_finite(fin_pmt_raw(rate, periods, present, future, kind)?)
}

fn fin_nper(rate: f64, payment: f64, present: f64, future: f64, kind: f64) -> Result<Value, ExcelError> {
    if rate == 0.0 {
        if payment == 0.0 {
            return Err(ExcelError::Num);
        }
        return fin_finite(-(present + future) / payment);
    }
    if rate <= -1.0 {
        return Err(ExcelError::Num);
    }
    let adjusted = payment * (1.0 + rate * kind);
    let denominator = present * rate + adjusted;
    if denominator == 0.0 {
        return Err(ExcelError::Num);
    }
    let ratio = (adjusted - future * rate) / denominator;
    if ratio <= 0.0 {
        return Err(ExcelError::Num);
    }
    fin_finite(ratio.ln() / (1.0 + rate).ln())
}

fn fin_equation(rate: f64, periods: f64, payment: f64, present: f64, future: f64, kind: f64) -> Result<f64, ExcelError> {
    if rate.abs() < 1e-12 {
        return Ok(present + payment * periods + future);
    }
    let factor = annuity_factor(rate, periods)?;
    Ok(present * factor + payment * (1.0 + rate * kind) * (factor - 1.0) / rate + future)
}

fn fin_rate(periods: f64, payment: f64, present: f64, future: f64, kind: f64, guess: f64) -> Result<Value, ExcelError> {
    if periods <= 0.0 || !guess.is_finite() || guess <= -1.0 {
        return Err(ExcelError::Num);
    }
    let mut rate = guess;
    for _ in 0..20 {
        let value = fin_equation(rate, periods, payment, present, future, kind)?;
        let step = (rate.abs() * 1e-6).max(1e-7);
        let lower = (rate - step).max(-0.999_999_999);
        let upper = rate + step;
        let derivative = (fin_equation(upper, periods, payment, present, future, kind)?
            - fin_equation(lower, periods, payment, present, future, kind)?)
            / (upper - lower);
        if derivative == 0.0 || !derivative.is_finite() {
            break;
        }
        let mut next = rate - value / derivative;
        if next <= -1.0 {
            next = (rate - 1.0) / 2.0;
        }
        if !next.is_finite() {
            break;
        }
        if (next - rate).abs() <= 1e-7 {
            return fin_finite(next);
        }
        rate = next;
    }
    // Newton from the guess can run away where a rate is still there to be
    // found: measured, `RATE(360, -1073.64, 200000)` is 0.00416664453634559,
    // which Newton from 0.1 overshoots. Find where the balance changes sign
    // and close in on it.
    let equation = |rate: f64| fin_equation(rate, periods, payment, present, future, kind).ok().filter(|v| v.is_finite());
    const MARKS: [f64; 22] = [
        -0.99, -0.9, -0.5, -0.2, -0.1, -0.05, -0.01, -0.001, 1e-6, 0.0005, 0.001, 0.002, 0.005, 0.01, 0.02,
        0.05, 0.1, 0.2, 0.5, 1.0, 2.0, 10.0,
    ];
    let mut brackets: Vec<(f64, f64)> = Vec::new();
    for pair in MARKS.windows(2) {
        if let (Some(low), Some(high)) = (equation(pair[0]), equation(pair[1])) {
            if low == 0.0 {
                return fin_finite(pair[0]);
            }
            if low.signum() != high.signum() {
                brackets.push((pair[0], pair[1]));
            }
        }
    }
    let Some(&(mut low, mut high)) = brackets
        .iter()
        .min_by(|a, b| (a.0 - guess).abs().total_cmp(&(b.0 - guess).abs()))
    else {
        return Err(ExcelError::Num);
    };
    let low_sign = equation(low).map(f64::signum).ok_or(ExcelError::Num)?;
    for _ in 0..200 {
        let middle = (low + high) / 2.0;
        if middle == low || middle == high {
            break;
        }
        match equation(middle) {
            Some(value) if value == 0.0 => return fin_finite(middle),
            Some(value) if value.signum() == low_sign => low = middle,
            Some(_) => high = middle,
            None => return Err(ExcelError::Num),
        }
    }
    fin_finite((low + high) / 2.0)
}

fn fin_irr(values: &[f64], guess: f64) -> Result<Value, ExcelError> {
    if values.len() < 2
        || !values.iter().any(|v| *v < 0.0)
        || !values.iter().any(|v| *v > 0.0)
        || !guess.is_finite()
        || guess <= -1.0
    {
        return Err(ExcelError::Num);
    }
    let mut rate = guess;
    for _ in 0..20 {
        let base = 1.0 + rate;
        let mut value = 0.0;
        let mut derivative = 0.0;
        for (period, flow) in values.iter().enumerate() {
            let period = period as f64;
            value += flow / base.powf(period);
            if period != 0.0 {
                derivative -= period * flow / base.powf(period + 1.0);
            }
        }
        if !value.is_finite() || !derivative.is_finite() || derivative == 0.0 {
            break;
        }
        let mut next = rate - value / derivative;
        if next <= -1.0 {
            next = (rate - 1.0) / 2.0;
        }
        if !next.is_finite() {
            break;
        }
        if (next - rate).abs() <= 1e-7 {
            return fin_finite(next);
        }
        rate = next;
    }
    Err(ExcelError::Num)
}

fn fin_sln(cost: f64, salvage: f64, life: f64) -> Result<Value, ExcelError> {
    // Measured: SLN(1000,100,0) is #DIV/0!.
    if life == 0.0 {
        return Err(ExcelError::DivZero);
    }
    if cost < 0.0 || salvage < 0.0 || life < 0.0 {
        return Err(ExcelError::Num);
    }
    fin_finite((cost - salvage) / life)
}

fn fin_syd(cost: f64, salvage: f64, life: f64, period: f64) -> Result<Value, ExcelError> {
    if cost < 0.0 || salvage < 0.0 || life <= 0.0 || period <= 0.0 || period > life {
        return Err(ExcelError::Num);
    }
    fin_finite((cost - salvage) * (life - period + 1.0) * 2.0 / (life * (life + 1.0)))
}

fn fin_ddb(cost: f64, salvage: f64, life: f64, period: f64, factor: f64) -> Result<Value, ExcelError> {
    if cost < 0.0 || salvage < 0.0 || life <= 0.0 || period <= 0.0 || factor <= 0.0 {
        return Err(ExcelError::Num);
    }
    if period > life {
        return Err(ExcelError::Num);
    }
    if cost <= salvage {
        return Ok(Value::Number(0.0));
    }
    let rate = (factor / life).min(1.0);
    let book = cost * (1.0 - rate).powf(period - 1.0);
    fin_finite((book * rate).min((book - salvage).max(0.0)))
}

/// Fixed-declining-balance depreciation for one period. The rate is rounded to
/// three decimals, and the first and last periods are prorated by `month`.
/// Measured: DB(10000,1000,5,2) = 2328.39, DB(10000,1000,5,6,6) = 238.53.
fn fin_db(cost: f64, salvage: f64, life: f64, period: f64, month: f64) -> Result<Value, ExcelError> {
    if cost < 0.0 || salvage < 0.0 || life <= 0.0 || period <= 0.0 || !(1.0..=12.0).contains(&month) {
        return Err(ExcelError::Num);
    }
    let last = life + if month < 12.0 { 1.0 } else { 0.0 };
    if period > last {
        return Err(ExcelError::Num);
    }
    // The whole rate is rounded to three decimals, not the ratio inside it.
    let rate = ((1.0 - (salvage / cost).powf(1.0 / life)) * 1000.0).round() / 1000.0;
    let target = period.floor() as i64;
    let mut total = 0.0;
    let mut answer = 0.0;
    for p in 1..=target {
        let book = cost - total;
        let dep = if p == 1 {
            book * rate * month / 12.0
        } else if (p as f64) == last && month < 12.0 {
            book * rate * (12.0 - month) / 12.0
        } else {
            book * rate
        };
        answer = dep;
        total += dep;
    }
    fin_finite(answer)
}

fn column_letters(mut col: u64) -> String {
    let mut letters = Vec::new();
    while col > 0 {
        col -= 1;
        letters.push(b'A' + (col % 26) as u8);
        col /= 26;
    }
    letters.reverse();
    String::from_utf8(letters).unwrap_or_default()
}

/// For XMATCH -1/1: the position of an exact match, or of the nearest value
/// below (`smaller`) or above the key.
fn xmatch_nearest(hay: &[Value], key: &Value, order: &[usize], smaller: bool) -> Option<usize> {
    let mut best: Option<usize> = None;
    for &i in order {
        match compare(&hay[i], key) {
            Ok(Ordering::Equal) => return Some(i),
            Ok(side) => {
                let candidate = if smaller {
                    side == Ordering::Less
                } else {
                    side == Ordering::Greater
                };
                if candidate {
                    let better = match best {
                        None => true,
                        Some(b) => {
                            let against = compare(&hay[i], &hay[b]);
                            if smaller {
                                against == Ok(Ordering::Greater)
                            } else {
                                against == Ok(Ordering::Less)
                            }
                        }
                    };
                    if better {
                        best = Some(i);
                    }
                }
            }
            Err(_) => continue,
        }
    }
    best
}

/// Truncate (not round) to a number of significant digits, as PERCENTRANK does.
pub(crate) fn truncate_significant(value: f64, digits: i32) -> f64 {
    if value == 0.0 {
        return 0.0;
    }
    let magnitude = value.abs().log10().floor() as i32;
    let places = digits - 1 - magnitude;
    let factor = 10f64.powi(places);
    (value * factor).trunc() / factor
}

/// The standard normal density.
fn norm_pdf(x: f64) -> f64 {
    (-0.5 * x * x).exp() / (2.0 * std::f64::consts::PI).sqrt()
}

/// The standard normal cumulative distribution, by West's (2004) rational
/// approximation -- accurate to about 1e-16, which is what Excel's eight
/// printed digits need even out in the tails.
pub(crate) fn norm_cdf(x: f64) -> f64 {
    let z = x.abs();
    if z > 37.0 {
        return if x > 0.0 { 1.0 } else { 0.0 };
    }
    let e = (-0.5 * z * z).exp();
    let tail = if z < 7.071_067_811_865_47 {
        let n = (((((3.526_249_659_989_11e-2 * z + 0.700_383_064_443_688) * z
            + 6.373_962_203_531_65) * z
            + 33.912_866_078_383) * z
            + 112.079_291_497_871) * z
            + 221.213_596_169_931) * z
            + 220.206_867_912_376;
        let d = ((((((8.838_834_764_831_84e-2 * z + 1.755_667_163_182_64) * z
            + 16.064_177_579_207) * z
            + 86.780_732_202_946_1) * z
            + 296.564_248_779_674) * z
            + 637.333_633_378_831) * z
            + 793.826_512_519_948) * z
            + 440.413_735_824_752;
        e * n / d
    } else {
        let f = z + 1.0 / (z + 2.0 / (z + 3.0 / (z + 4.0 / (z + 0.65))));
        e / (2.506_628_274_631_000_2 * f)
    };
    if x <= 0.0 { tail } else { 1.0 - tail }
}

/// The inverse standard normal, by Acklam's rational approximation with one
/// Halley step against `norm_cdf`, which brings it to full double precision.
fn norm_s_inv(p: f64) -> f64 {
    const A: [f64; 6] = [
        -3.969_683_028_665_376e1, 2.209_460_984_245_205e2, -2.759_285_104_469_687e2,
        1.383_577_518_672_69e2, -3.066_479_806_614_716e1, 2.506_628_277_459_239e0,
    ];
    const B: [f64; 5] = [
        -5.447_609_879_822_406e1, 1.615_858_368_580_409e2, -1.556_989_798_598_866e2,
        6.680_131_188_771_972e1, -1.328_068_155_288_572e1,
    ];
    const C: [f64; 6] = [
        -7.784_894_002_430_293e-3, -3.223_964_580_411_365e-1, -2.400_758_277_161_838e0,
        -2.549_732_539_343_734e0, 4.374_664_141_464_968e0, 2.938_163_982_698_783e0,
    ];
    const D: [f64; 4] = [
        7.784_695_709_041_462e-3, 3.224_671_290_700_398e-1, 2.445_134_137_142_996e0,
        3.754_408_661_907_416e0,
    ];
    let plow = 0.024_25;
    let phigh = 1.0 - plow;
    let mut x = if p < plow {
        let q = (-2.0 * p.ln()).sqrt();
        (((((C[0] * q + C[1]) * q + C[2]) * q + C[3]) * q + C[4]) * q + C[5])
            / ((((D[0] * q + D[1]) * q + D[2]) * q + D[3]) * q + 1.0)
    } else if p <= phigh {
        let q = p - 0.5;
        let r = q * q;
        (((((A[0] * r + A[1]) * r + A[2]) * r + A[3]) * r + A[4]) * r + A[5]) * q
            / (((((B[0] * r + B[1]) * r + B[2]) * r + B[3]) * r + B[4]) * r + 1.0)
    } else {
        let q = (-2.0 * (1.0 - p).ln()).sqrt();
        -(((((C[0] * q + C[1]) * q + C[2]) * q + C[3]) * q + C[4]) * q + C[5])
            / ((((D[0] * q + D[1]) * q + D[2]) * q + D[3]) * q + 1.0)
    };
    // One Halley step: x -= (cdf(x)-p) / phi(x) corrected for curvature.
    let error = norm_cdf(x) - p;
    let u = error / norm_pdf(x);
    x -= u / (1.0 + x * u / 2.0);
    x
}

/// The distributions of the arm above, each checked the way Excel checks
/// its arguments: degrees of freedom are cut to whole numbers and must be
/// at least one, a probability must lie in (0, 1], and so on, #NUM!
/// otherwise.
fn distribution(name: &str, args: &[Arg]) -> Result<f64, ExcelError> {
    use crate::distributions as d;
    let at = |i: usize| -> Result<f64, ExcelError> {
        match args.get(i) {
            Some(arg) => num(arg),
            None => Err(ExcelError::Value),
        }
    };
    let flag = |i: usize| -> Result<bool, ExcelError> {
        match args.get(i) {
            Some(arg) => arg.scalar().to_logical(),
            None => Err(ExcelError::Value),
        }
    };
    let count = |low: usize, high: usize| {
        if args.len() < low || args.len() > high {
            Err(ExcelError::Value)
        } else {
            Ok(())
        }
    };
    let freedom = |value: f64| {
        let value = value.trunc();
        if value < 1.0 {
            Err(ExcelError::Num)
        } else {
            Ok(value)
        }
    };
    let chance = |p: f64| if p <= 0.0 || p > 1.0 { Err(ExcelError::Num) } else { Ok(p) };
    let answer = match name {
        "T.DIST" => {
            count(3, 3)?;
            let (x, df) = (at(0)?, freedom(at(1)?)?);
            if flag(2)? { d::t_cdf(x, df) } else { d::t_pdf(x, df) }
        }
        "T.DIST.RT" => {
            count(2, 2)?;
            1.0 - d::t_cdf(at(0)?, freedom(at(1)?)?)
        }
        "T.DIST.2T" => {
            count(2, 2)?;
            let x = at(0)?;
            if x < 0.0 {
                return Err(ExcelError::Num);
            }
            d::t_two_tailed(x, freedom(at(1)?)?)
        }
        "TDIST" => {
            count(3, 3)?;
            let (x, df, tails) = (at(0)?, freedom(at(1)?)?, at(2)?.trunc());
            if x < 0.0 {
                return Err(ExcelError::Num);
            }
            match tails as i64 {
                1 => 0.5 * d::t_two_tailed(x, df),
                2 => d::t_two_tailed(x, df),
                _ => return Err(ExcelError::Num),
            }
        }
        "T.INV" => {
            count(2, 2)?;
            let p = at(0)?;
            if p <= 0.0 || p >= 1.0 {
                return Err(ExcelError::Num);
            }
            d::t_inv(p, freedom(at(1)?)?)
        }
        "T.INV.2T" | "TINV" => {
            count(2, 2)?;
            let p = chance(at(0)?)?;
            d::t_inv(1.0 - p / 2.0, freedom(at(1)?)?).abs()
        }
        "CHISQ.DIST" => {
            count(3, 3)?;
            let (x, k) = (at(0)?, freedom(at(1)?)?);
            if x < 0.0 {
                return Err(ExcelError::Num);
            }
            if flag(2)? { d::regularized_gamma_p(k / 2.0, x / 2.0) } else { d::chisq_pdf(x, k) }
        }
        "CHISQ.DIST.RT" | "CHIDIST" => {
            count(2, 2)?;
            let (x, k) = (at(0)?, freedom(at(1)?)?);
            if x < 0.0 {
                return Err(ExcelError::Num);
            }
            d::regularized_gamma_q(k / 2.0, x / 2.0)
        }
        "CHISQ.INV" => {
            count(2, 2)?;
            let (p, k) = (at(0)?, freedom(at(1)?)?);
            if !(0.0..1.0).contains(&p) {
                return Err(ExcelError::Num);
            }
            d::invert(p, 0.0, k.max(1.0), |x| d::regularized_gamma_p(k / 2.0, x / 2.0))
        }
        "CHISQ.INV.RT" | "CHIINV" => {
            count(2, 2)?;
            let (p, k) = (chance(at(0)?)?, freedom(at(1)?)?);
            d::invert_upper(p, 0.0, k.max(1.0), |x| d::regularized_gamma_q(k / 2.0, x / 2.0))
        }
        "F.DIST" => {
            count(4, 4)?;
            let (x, d1, d2) = (at(0)?, freedom(at(1)?)?, freedom(at(2)?)?);
            if x < 0.0 {
                return Err(ExcelError::Num);
            }
            if flag(3)? { d::f_cdf(x, d1, d2) } else { d::f_pdf(x, d1, d2) }
        }
        "F.DIST.RT" | "FDIST" => {
            count(3, 3)?;
            let (x, d1, d2) = (at(0)?, freedom(at(1)?)?, freedom(at(2)?)?);
            if x < 0.0 {
                return Err(ExcelError::Num);
            }
            d::f_upper(x, d1, d2)
        }
        "F.INV" => {
            count(3, 3)?;
            let (p, d1, d2) = (at(0)?, freedom(at(1)?)?, freedom(at(2)?)?);
            if !(0.0..1.0).contains(&p) {
                return Err(ExcelError::Num);
            }
            d::invert(p, 0.0, 1.0, |x| d::f_cdf(x, d1, d2))
        }
        "F.INV.RT" | "FINV" => {
            count(3, 3)?;
            let (p, d1, d2) = (chance(at(0)?)?, freedom(at(1)?)?, freedom(at(2)?)?);
            d::invert_upper(p, 0.0, 1.0, |x| d::f_upper(x, d1, d2))
        }
        "GAMMA.DIST" | "GAMMADIST" => {
            count(4, 4)?;
            let (x, alpha, beta) = (at(0)?, at(1)?, at(2)?);
            if x < 0.0 || alpha <= 0.0 || beta <= 0.0 {
                return Err(ExcelError::Num);
            }
            if flag(3)? { d::regularized_gamma_p(alpha, x / beta) } else { d::gamma_pdf(x, alpha, beta) }
        }
        "GAMMA.INV" | "GAMMAINV" => {
            count(3, 3)?;
            let (p, alpha, beta) = (at(0)?, at(1)?, at(2)?);
            if !(0.0..1.0).contains(&p) || alpha <= 0.0 || beta <= 0.0 {
                return Err(ExcelError::Num);
            }
            beta * d::invert(p, 0.0, alpha.max(1.0), |x| d::regularized_gamma_p(alpha, x))
        }
        "GAMMALN" | "GAMMALN.PRECISE" => {
            count(1, 1)?;
            let x = at(0)?;
            if x <= 0.0 {
                return Err(ExcelError::Num);
            }
            ln_gamma(x)
        }
        "GAMMA" => {
            count(1, 1)?;
            let x = at(0)?;
            if x <= 0.0 && x.fract() == 0.0 {
                return Err(ExcelError::Num);
            }
            let value = d::gamma(x);
            if !value.is_finite() {
                return Err(ExcelError::Num);
            }
            value
        }
        "BETA.DIST" | "BETADIST" => {
            let legacy = name == "BETADIST";
            if legacy { count(3, 5)? } else { count(4, 6)? }
            let (x, a, b) = (at(0)?, at(1)?, at(2)?);
            let bounds_from = if legacy { 3 } else { 4 };
            let low = match args.get(bounds_from) { Some(arg) => num(arg)?, None => 0.0 };
            let high = match args.get(bounds_from + 1) { Some(arg) => num(arg)?, None => 1.0 };
            if a <= 0.0 || b <= 0.0 || x < low || x > high || low == high {
                return Err(ExcelError::Num);
            }
            let scaled = (x - low) / (high - low);
            if legacy || flag(3)? {
                regularized_beta(scaled, a, b)
            } else {
                d::beta_pdf(scaled, a, b) / (high - low)
            }
        }
        "BETA.INV" | "BETAINV" => {
            count(3, 5)?;
            let (p, a, b) = (at(0)?, at(1)?, at(2)?);
            let low = match args.get(3) { Some(arg) => num(arg)?, None => 0.0 };
            let high = match args.get(4) { Some(arg) => num(arg)?, None => 1.0 };
            if p <= 0.0 || p > 1.0 || a <= 0.0 || b <= 0.0 || low >= high {
                return Err(ExcelError::Num);
            }
            let mut lo = 0.0;
            let mut hi = 1.0;
            for _ in 0..200 {
                let middle = 0.5 * (lo + hi);
                if middle <= lo || middle >= hi {
                    break;
                }
                if regularized_beta(middle, a, b) < p { lo = middle } else { hi = middle }
            }
            low + 0.5 * (lo + hi) * (high - low)
        }
        "LOGNORM.DIST" | "LOGNORMDIST" => {
            let legacy = name == "LOGNORMDIST";
            if legacy { count(3, 3)? } else { count(4, 4)? }
            let (x, mean, sd) = (at(0)?, at(1)?, at(2)?);
            if x <= 0.0 || sd <= 0.0 {
                return Err(ExcelError::Num);
            }
            let z = (x.ln() - mean) / sd;
            if legacy || flag(3)? { norm_cdf(z) } else { norm_pdf(z) / (x * sd) }
        }
        "LOGNORM.INV" | "LOGINV" => {
            count(3, 3)?;
            let (p, mean, sd) = (at(0)?, at(1)?, at(2)?);
            if p <= 0.0 || p >= 1.0 || sd <= 0.0 {
                return Err(ExcelError::Num);
            }
            (mean + sd * norm_s_inv(p)).exp()
        }
        "HYPGEOM.DIST" | "HYPGEOMDIST" => {
            let legacy = name == "HYPGEOMDIST";
            if legacy { count(4, 4)? } else { count(5, 5)? }
            let (s, n, k, total) = (at(0)?.trunc(), at(1)?.trunc(), at(2)?.trunc(), at(3)?.trunc());
            if s < 0.0 || s > n || s > k || n > total || k > total || n <= 0.0 || k <= 0.0 || total <= 0.0 || s < n - total + k {
                return Err(ExcelError::Num);
            }
            let mass = |s: f64| (ln_choose(k, s) + ln_choose(total - k, n - s) - ln_choose(total, n)).exp();
            if !legacy && flag(4)? {
                let first = (n - total + k).max(0.0) as i64;
                (first..=s as i64).map(|one| mass(one as f64)).sum()
            } else {
                mass(s)
            }
        }
        "NEGBINOM.DIST" | "NEGBINOMDIST" => {
            let legacy = name == "NEGBINOMDIST";
            if legacy { count(3, 3)? } else { count(4, 4)? }
            let (f, s, p) = (at(0)?.trunc(), at(1)?.trunc(), at(2)?);
            if !(0.0..=1.0).contains(&p) || f < 0.0 || s < 1.0 {
                return Err(ExcelError::Num);
            }
            if !legacy && flag(3)? {
                regularized_beta(p, s, f + 1.0)
            } else {
                (ln_choose(f + s - 1.0, s - 1.0) + s * p.ln() + f * (1.0 - p).ln()).exp()
            }
        }
        "WEIBULL.DIST" | "WEIBULL" => {
            count(4, 4)?;
            let (x, alpha, beta) = (at(0)?, at(1)?, at(2)?);
            if x < 0.0 || alpha <= 0.0 || beta <= 0.0 {
                return Err(ExcelError::Num);
            }
            let power = (x / beta).powf(alpha);
            if flag(3)? {
                -(-power).exp_m1()
            } else {
                alpha / beta.powf(alpha) * x.powf(alpha - 1.0) * (-power).exp()
            }
        }
        "FISHER" => {
            count(1, 1)?;
            let x = at(0)?;
            if x <= -1.0 || x >= 1.0 {
                return Err(ExcelError::Num);
            }
            0.5 * ((1.0 + x) / (1.0 - x)).ln()
        }
        "FISHERINV" => {
            count(1, 1)?;
            at(0)?.tanh()
        }
        "ERF" => {
            count(1, 2)?;
            let low = at(0)?;
            match args.get(1) {
                Some(arg) => d::erf(num(arg)?) - d::erf(low),
                None => d::erf(low),
            }
        }
        "ERF.PRECISE" => {
            count(1, 1)?;
            d::erf(at(0)?)
        }
        "ERFC" | "ERFC.PRECISE" => {
            count(1, 1)?;
            d::erfc(at(0)?)
        }
        "BINOM.INV" | "CRITBINOM" => {
            count(3, 3)?;
            let (n, p, alpha) = (at(0)?.trunc(), at(1)?, at(2)?);
            if n < 0.0 || !(0.0..=1.0).contains(&p) || !(0.0..=1.0).contains(&alpha) {
                return Err(ExcelError::Num);
            }
            let mut total = 0.0;
            let mut k = 0.0;
            loop {
                total += (ln_choose(n, k) + k * p.ln() + (n - k) * (1.0 - p).ln()).exp();
                if total >= alpha || k >= n {
                    break k;
                }
                k += 1.0;
            }
        }
        "CONFIDENCE.T" => {
            count(3, 3)?;
            let (alpha, sd, size) = (at(0)?, at(1)?, at(2)?.trunc());
            if alpha <= 0.0 || alpha >= 1.0 || sd <= 0.0 || size < 1.0 {
                return Err(ExcelError::Num);
            }
            if size == 1.0 {
                return Err(ExcelError::DivZero);
            }
            d::t_inv(1.0 - alpha / 2.0, size - 1.0) * sd / size.sqrt()
        }
        _ => return Err(ExcelError::Name),
    };
    if answer.is_finite() {
        Ok(answer)
    } else {
        Err(ExcelError::Num)
    }
}

/// n!, refusing a negative and overflowing to #NUM!.
fn factorial(n: f64) -> Result<f64, ExcelError> {
    if n < 0.0 {
        return Err(ExcelError::Num);
    }
    // Multiplied from the top down: measured, FACT(170) is
    // 7.257415615308E+306, where counting up gives ...7994 (…799).
    let mut acc = 1.0f64;
    for i in (2..=(n.trunc() as u64)).rev() {
        acc *= i as f64;
        if !acc.is_finite() {
            return Err(ExcelError::Num);
        }
    }
    Ok(acc)
}

/// n!!: the product going down by twos. FactDouble(7)=105, and 0 and -1 are 1.
fn factdouble(n: f64) -> Result<f64, ExcelError> {
    let mut i = n.trunc() as i64;
    if i < -1 {
        return Err(ExcelError::Num);
    }
    let mut acc = 1.0f64;
    while i > 1 {
        acc *= i as f64;
        if !acc.is_finite() {
            return Err(ExcelError::Num);
        }
        i -= 2;
    }
    Ok(acc)
}

/// n choose k, built up multiplicatively so it stays exact for the sizes that
/// fit; #NUM! unless 0 <= k <= n.
fn combin(n: f64, k: f64) -> Result<f64, ExcelError> {
    let (n, k) = (n.trunc(), k.trunc());
    if n < 0.0 || k < 0.0 || k > n {
        return Err(ExcelError::Num);
    }
    let (n, k) = (n as u64, k as u64);
    let k = k.min(n - k);
    let mut acc = 1.0f64;
    for i in 0..k {
        acc = acc * (n - i) as f64 / (i + 1) as f64;
    }
    Ok(acc.round())
}

/// The number of ordered arrangements, n!/(n-k)!.
fn permut(n: f64, k: f64) -> Result<f64, ExcelError> {
    let (n, k) = (n.trunc(), k.trunc());
    if n < 0.0 || k < 0.0 || k > n {
        return Err(ExcelError::Num);
    }
    let (n, k) = (n as u64, k as u64);
    let mut acc = 1.0f64;
    for i in 0..k {
        acc *= (n - i) as f64;
        if !acc.is_finite() {
            return Err(ExcelError::Num);
        }
    }
    Ok(acc)
}

/// The greatest common divisor, for GCD/LCM.
fn gcd(mut a: i64, mut b: i64) -> i64 {
    while b != 0 {
        let t = b;
        b = a % b;
        a = t;
    }
    a.abs()
}

/// Whether the largest-not-over test of LOOKUP holds for a candidate against
/// the needle: numbers compare as numbers, text as text without case.
/// `x ^ y` as Excel works it: a negative number to one over an odd whole
/// number is its real root -- measured, `=POWER(-8,1/3)` is -2.
pub(crate) fn excel_power(x: f64, y: f64) -> f64 {
    if x < 0.0 && y.fract() != 0.0 {
        let inverse = 1.0 / y;
        let whole = inverse.round();
        if (inverse - whole).abs() < 1e-9 && whole as i64 % 2 != 0 {
            return -(-x).powf(y);
        }
    }
    x.powf(y)
}

/// ln Γ(x): the log of the factorial for a whole number, else the
/// Stirling series once x has been walked up past 15, which holds every
/// digit a double has (Lanczos below one half, through the reflection).
pub(crate) fn ln_gamma(x: f64) -> f64 {
    if x >= 0.5 {
        if x.fract() == 0.0 && x <= 171.0 {
            let mut product = 1.0f64;
            let mut k = 2.0;
            while k < x {
                product *= k;
                k += 1.0;
            }
            return product.ln();
        }
        let mut shift = 0.0f64;
        let mut z = x;
        let mut product = 1.0f64;
        while z < 15.0 {
            product *= z;
            z += 1.0;
        }
        if product != 1.0 {
            shift = product.ln();
        }
        let inverse = 1.0 / z;
        let square = inverse * inverse;
        let series = inverse
            * (1.0 / 12.0
                - square
                    * (1.0 / 360.0
                        - square * (1.0 / 1260.0 - square * (1.0 / 1680.0 - square * (1.0 / 1188.0 - square * 691.0 / 360_360.0)))));
        return (z - 0.5) * z.ln() - z + 0.5 * (2.0 * std::f64::consts::PI).ln() + series - shift;
    }
    lanczos_ln_gamma(x)
}

fn lanczos_ln_gamma(x: f64) -> f64 {
    const G: [f64; 9] = [
        0.999_999_999_999_809_9,
        676.520_368_121_885_1,
        -1_259.139_216_722_402_8,
        771.323_428_777_653_1,
        -176.615_029_162_140_6,
        12.507_343_278_686_905,
        -0.138_571_095_265_720_12,
        9.984_369_578_019_572e-6,
        1.505_632_735_149_311_6e-7,
    ];
    if x < 0.5 {
        return (std::f64::consts::PI / (std::f64::consts::PI * x).sin()).ln() - ln_gamma(1.0 - x);
    }
    let x = x - 1.0;
    let mut sum = G[0];
    for (i, g) in G.iter().enumerate().skip(1) {
        sum += g / (x + i as f64);
    }
    let t = x + 7.5;
    0.5 * (2.0 * std::f64::consts::PI).ln() + (x + 0.5) * t.ln() - t + sum.ln()
}

fn ln_choose(n: f64, k: f64) -> f64 {
    ln_gamma(n + 1.0) - ln_gamma(k + 1.0) - ln_gamma(n - k + 1.0)
}

/// The regularised incomplete beta function I_x(a, b), by continued fraction.
pub(crate) fn regularized_beta(x: f64, a: f64, b: f64) -> f64 {
    if x <= 0.0 {
        return 0.0;
    }
    if x >= 1.0 {
        return 1.0;
    }
    // A whole-numbered b makes it a finite sum, which keeps every digit:
    // I_x(a, b) = x^a * sum_{j<b} Γ(a+j)/(Γ(a) j!) (1-x)^j. A whole a is the
    // same from the other side.
    let finite = |x: f64, a: f64, b: f64| {
        let mut term = 1.0f64;
        let mut sum = 1.0f64;
        let mut j = 1.0;
        while j < b {
            term *= (a + j - 1.0) / j * (1.0 - x);
            sum += term;
            j += 1.0;
        }
        x.powf(a) * sum
    };
    if b.fract() == 0.0 && b <= 1000.0 {
        return finite(x, a, b);
    }
    if a.fract() == 0.0 && a <= 1000.0 {
        return 1.0 - finite(1.0 - x, b, a);
    }
    let front = (ln_gamma(a + b) - ln_gamma(a) - ln_gamma(b) + a * x.ln() + b * (1.0 - x).ln()).exp();
    let fraction = |x: f64, a: f64, b: f64| {
        let tiny = 1e-300;
        let (mut c, mut d) = (1.0, 1.0 - (a + b) * x / (a + 1.0));
        if d.abs() < tiny {
            d = tiny;
        }
        d = 1.0 / d;
        let mut h = d;
        for m in 1..300 {
            let m = m as f64;
            let even = m * (b - m) * x / ((a + 2.0 * m - 1.0) * (a + 2.0 * m));
            d = 1.0 + even * d;
            if d.abs() < tiny { d = tiny; }
            c = 1.0 + even / c;
            if c.abs() < tiny { c = tiny; }
            d = 1.0 / d;
            h *= d * c;
            let odd = -(a + m) * (a + b + m) * x / ((a + 2.0 * m) * (a + 2.0 * m + 1.0));
            d = 1.0 + odd * d;
            if d.abs() < tiny { d = tiny; }
            c = 1.0 + odd / c;
            if c.abs() < tiny { c = tiny; }
            d = 1.0 / d;
            let step = d * c;
            h *= step;
            if (step - 1.0).abs() < 1e-15 {
                break;
            }
        }
        h
    };
    if x < (a + 1.0) / (a + b + 2.0) {
        front * fraction(x, a, b) / a
    } else {
        1.0 - front * fraction(1.0 - x, b, a) / b
    }
}

/// Excel's search over a table it assumes is sorted, zero-based: a plain
/// binary search that stops on the first equal probe and then walks to the
/// far end of that run of equals -- forward ascending, backward descending --
/// and a descending search refuses a needle above the leading value. The same
/// search `WorksheetFunction.Match` was measured to make (5,040 differential
/// cases); on unsorted data it is why `LOOKUP(4.5, {1,3,4,6,2,5})` is 2's
/// neighbour 2 and not 4.
fn sorted_position(count: usize, descending: bool, value_at: impl Fn(usize) -> Value, needle: &Value) -> Option<usize> {
    if count == 0 {
        return None;
    }
    // A cell of another kind than the needle -- text beside a number, a
    // Boolean -- is passed over: the search looks on from it for one of the
    // needle's kind, forward first and then back. Measured: MATCH(9, {1, 3,
    // 5, "apple", "Banana", "b*n", TRUE, 7}, 1) is 8.
    let kind = |value: &Value| match value {
        Value::Number(_) => 0,
        Value::Text(_) => 1,
        Value::Logical(_) => 2,
        _ => 3,
    };
    let wanted = kind(needle);
    let same_kind = |i: usize| kind(&value_at(i)) == wanted;
    let order = |i: usize| compare(&value_at(i), needle).ok();
    if descending {
        match order(0) {
            Some(Ordering::Less) => return None,
            Some(Ordering::Equal) => return Some(0),
            _ => {}
        }
    }
    let (mut low, mut high) = (1usize, count);
    let mut found = None;
    while low <= high {
        let mut middle = (low + high) / 2;
        if !same_kind(middle - 1) {
            let ahead = (middle..=high).find(|at| same_kind(at - 1));
            let behind = (low..middle).rev().find(|at| same_kind(at - 1));
            match ahead.or(behind) {
                Some(at) => middle = at,
                None => break,
            }
        }
        match order(middle - 1) {
            Some(Ordering::Equal) => {
                found = Some(middle);
                break;
            }
            Some(side) if (side == Ordering::Less) != descending => {
                found = Some(middle);
                low = middle + 1;
            }
            _ => high = middle - 1,
        }
    }
    let mut position = found?;
    if descending {
        while position > 1 && order(position - 2) == Some(Ordering::Equal) {
            position -= 1;
        }
    } else {
        while position < count && order(position) == Some(Ordering::Equal) {
            position += 1;
        }
    }
    Some(position - 1)
}

#[allow(dead_code)]
fn lookup_le(cell: &Value, needle: &Value) -> bool {
    match (cell, needle) {
        (Value::Number(a), Value::Number(b)) => a <= b,
        (Value::Text(a), Value::Text(b)) => a.to_lowercase() <= b.to_lowercase(),
        (Value::Number(_), _) => true,
        _ => false,
    }
}

/// The classic subtractive Roman numeral for 0..=3999 (0 is empty, as Excel).
fn roman_numeral(mut n: i64) -> String {
    const TABLE: [(i64, &str); 13] = [
        (1000, "M"), (900, "CM"), (500, "D"), (400, "CD"), (100, "C"),
        (90, "XC"), (50, "L"), (40, "XL"), (10, "X"), (9, "IX"),
        (5, "V"), (4, "IV"), (1, "I"),
    ];
    let mut out = String::new();
    for (value, sign) in TABLE {
        while n >= value {
            out.push_str(sign);
            n -= value;
        }
    }
    out
}

/// The number a classic Roman numeral stands for, for ARABIC.
fn arabic_number(text: &str) -> Option<i64> {
    let text = text.trim().to_uppercase();
    let (negative, text) = match text.strip_prefix('-') {
        Some(rest) => (true, rest.to_string()),
        None => (false, text),
    };
    let value = |c: char| match c {
        'I' => Some(1),
        'V' => Some(5),
        'X' => Some(10),
        'L' => Some(50),
        'C' => Some(100),
        'D' => Some(500),
        'M' => Some(1000),
        _ => None,
    };
    let mut total = 0i64;
    let mut prev = 0i64;
    for c in text.chars().rev() {
        let v = value(c)?;
        if v < prev {
            total -= v;
        } else {
            total += v;
            prev = v;
        }
    }
    Some(if negative { -total } else { total })
}

/// The database functions -- `DSUM(database, field, criteria)` and its kin.
/// The database's first row names its columns; the field is a column by name
/// or by one-based number; the criteria range's first row names columns and
/// each row under it is one set of conditions, joined across a row by AND and
/// down the rows by OR.
fn database_function(name: &str, args: &[Arg]) -> Result<Value, ExcelError> {
    let database = args.first().ok_or(ExcelError::Value)?.as_range();
    let criteria = args.get(2).ok_or(ExcelError::Value)?.as_range();
    if database.height < 1 || criteria.height < 1 {
        return Err(ExcelError::Value);
    }
    let headers: Vec<String> = (0..database.width)
        .map(|c| database.at(c, 0).to_number().map(|n| n.to_string()).unwrap_or_else(|_| match database.at(c, 0) {
            Value::Text(t) => t,
            other => other.to_number().map(|n| n.to_string()).unwrap_or_default(),
        }))
        .collect();
    let field = match args.get(1).map(|a| a.scalar()) {
        Some(Value::Number(n)) => (n as usize).checked_sub(1).ok_or(ExcelError::Value)?,
        Some(Value::Text(t)) => headers
            .iter()
            .position(|h| h.eq_ignore_ascii_case(&t))
            .ok_or(ExcelError::Value)?,
        _ => return Err(ExcelError::Value),
    };
    if field >= database.width {
        return Err(ExcelError::Value);
    }
    let crit_headers: Vec<String> = (0..criteria.width)
        .map(|c| match criteria.at(c, 0) {
            Value::Text(t) => t,
            other => other.to_number().map(|n| n.to_string()).unwrap_or_default(),
        })
        .collect();
    let mut chosen: Vec<Value> = Vec::new();
    for row in 1..database.height {
        let mut any = false;
        for crow in 1..criteria.height {
            let mut all = true;
            for ccol in 0..criteria.width {
                let cond = criteria.at(ccol, crow);
                if matches!(cond, Value::Blank) {
                    continue;
                }
                let Some(dcol) = headers.iter().position(|h| h.eq_ignore_ascii_case(&crit_headers[ccol])) else {
                    all = false;
                    break;
                };
                if !criterion_matches(&database.at(dcol, row), &cond) {
                    all = false;
                    break;
                }
            }
            if all {
                any = true;
                break;
            }
        }
        if any {
            chosen.push(database.at(field, row));
        }
    }
    let numbers: Vec<f64> = chosen.iter().filter_map(|v| match v {
        Value::Number(n) => Some(*n),
        _ => None,
    }).collect();
    match name {
        "DSUM" => Ok(Value::Number(numbers.iter().sum())),
        "DPRODUCT" => Ok(Value::Number(numbers.iter().product())),
        "DAVERAGE" => {
            if numbers.is_empty() {
                return Err(ExcelError::DivZero);
            }
            Ok(Value::Number(numbers.iter().sum::<f64>() / numbers.len() as f64))
        }
        "DMAX" => Ok(Value::Number(numbers.iter().cloned().fold(f64::MIN, f64::max))),
        "DMIN" => Ok(Value::Number(numbers.iter().cloned().fold(f64::MAX, f64::min))),
        "DCOUNT" => Ok(Value::Number(numbers.len() as f64)),
        // The spreads of the chosen numbers, sample and whole.
        "DSTDEV" | "DSTDEVP" | "DVAR" | "DVARP" => {
            let whole = name.ends_with('P');
            let n = numbers.len() as f64;
            let divisor = if whole { n } else { n - 1.0 };
            if divisor <= 0.0 {
                return Err(ExcelError::DivZero);
            }
            let mean = numbers.iter().sum::<f64>() / n;
            let spread = numbers.iter().map(|x| (x - mean).powi(2)).sum::<f64>() / divisor;
            Ok(Value::Number(if name.starts_with("DSTDEV") { spread.sqrt() } else { spread }))
        }
        "DCOUNTA" => Ok(Value::Number(chosen.iter().filter(|v| !matches!(v, Value::Blank)).count() as f64)),
        "DGET" => match numbers.len() {
            1 => Ok(chosen[0].clone()),
            0 => Err(ExcelError::Value),
            _ => Err(ExcelError::Num),
        },
        _ => Err(ExcelError::Name),
    }
}

/// Whether a database cell meets a criterion cell: a bare value is an equal
/// test, and a leading `>`, `<`, `>=`, `<=`, `<>` or `=` a compare.
fn criterion_matches(cell: &Value, criterion: &Value) -> bool {
    let text = match criterion {
        Value::Text(t) => t.clone(),
        Value::Number(n) => return matches!(cell, Value::Number(c) if (c - n).abs() < 1e-9),
        _ => return true,
    };
    let text = text.trim();
    for op in ["<=", ">=", "<>", ">", "<", "="] {
        if let Some(rest) = text.strip_prefix(op) {
            let rest = rest.trim();
            if let Ok(threshold) = rest.parse::<f64>() {
                if let Value::Number(c) = cell {
                    return match op {
                        ">" => *c > threshold,
                        "<" => *c < threshold,
                        ">=" => *c >= threshold,
                        "<=" => *c <= threshold,
                        "<>" => (*c - threshold).abs() >= 1e-9,
                        _ => (*c - threshold).abs() < 1e-9,
                    };
                }
                return false;
            }
            let held = match cell {
                Value::Text(t) => t.clone(),
                other => other.to_number().map(|n| n.to_string()).unwrap_or_default(),
            };
            return match op {
                "<>" => !held.eq_ignore_ascii_case(rest),
                _ => held.eq_ignore_ascii_case(rest),
            };
        }
    }
    // A bare value: text matches text without case, a number matches a number.
    match cell {
        Value::Text(t) => t.eq_ignore_ascii_case(text),
        other => other.to_number().ok().map(|n| n.to_string()).as_deref() == Some(text),
    }
}

/// UNIQUE, SORT and FILTER: a block in, a block out.
fn a_block_of_rows(name: &str, args: &[Arg]) -> Result<Arg, ExcelError> {
    if args.is_empty() {
        return Err(ExcelError::Value);
    }
    let table = args[0].as_range();
    // `by_col` says to do the whole thing sideways. Turning the block on its
    // side, working on rows as usual, and turning it back is the same answer
    // with none of the second implementation.
    let sideways = match name {
        "UNIQUE" => reads_true(args.get(1)),
        "SORT" => reads_true(args.get(3)),
        _ => false,
    };
    // SORTBY orders one block by the values in ANOTHER, which do not appear in
    // the answer. Carrying that second block alongside as an extra column, and
    // taking it off again afterwards, makes it the same sort as any other.
    let (table, sort_by) = match name {
        "SORTBY" => {
            let beside = args.get(1).ok_or(ExcelError::Value)?.flatten();
            if beside.len() != table.height {
                return Err(ExcelError::Value);
            }
            (with_a_column(&table, &beside), Some(table.width))
        }
        _ => (table, None),
    };
    let table = if sideways { on_its_side(&table) } else { table };
    let mut rows: Vec<Vec<Value>> = (0..table.height)
        .map(|row| (0..table.width).map(|col| table.at(col, row)).collect())
        .collect();

    match name {
        "UNIQUE" => {
            // The third argument asks for the rows that appear EXACTLY once,
            // which is a different question from the distinct rows.
            let once_only = reads_true(args.get(2));
            let mut kept: Vec<Vec<Value>> = Vec::new();
            for row in &rows {
                let seen = rows.iter().filter(|other| same_row(other, row)).count();
                let already = kept.iter().any(|other| same_row(other, row));
                if already || (once_only && seen > 1) {
                    continue;
                }
                kept.push(row.clone());
            }
            rows = kept;
        }
        "SORT" | "SORTBY" => {
            // Which column to order by, counted from one, and which way.
            let by = match sort_by {
                Some(added) => added + 1,
                // Left empty, as `SORT(A1:A3,,-1)` leaves it, it is the
                // first column.
                None => match args.get(1) {
                    Some(one) if !one.scalar().is_blank() => num(one)? as usize,
                    _ => 1,
                },
            };
            // Both spell the direction third: SORT(block, by, order) and
            // SORTBY(block, ordered_by, order).
            let descending = match args.get(2) {
                Some(one) => num(one)? < 0.0,
                None => false,
            };
            if by < 1 || by > table.width {
                return Err(ExcelError::Value);
            }
            rows.sort_by(|left, right| in_order(&left[by - 1], &right[by - 1], descending));
        }
        _ => {
            // FILTER: a second block, as tall as this one, saying which rows
            // to keep.
            let asked = args.get(1).ok_or(ExcelError::Value)?.flatten();
            if asked.len() != rows.len() {
                return Err(ExcelError::Value);
            }
            let mut kept = Vec::new();
            for (row, wanted) in rows.into_iter().zip(asked) {
                if let Value::Error(why) = wanted {
                    return Err(why);
                }
                if wanted.to_logical()? {
                    kept.push(row);
                }
            }
            rows = kept;
        }
    }

    if rows.is_empty() {
        // Nothing left. FILTER's third argument says what to show instead;
        // without one there is no answer to give.
        return match args.get(2) {
            Some(instead) if name == "FILTER" => Ok(Arg::Value(instead.scalar())),
            // Measured: FILTER with nothing kept and no stand-in is #CALC!.
            None if name == "FILTER" => Err(ExcelError::Calc),
            _ => Err(ExcelError::NA),
        };
    }
    // The column SORTBY was ordering by is not part of the answer.
    if let Some(added) = sort_by {
        for row in &mut rows {
            row.truncate(added);
        }
    }
    let width = rows[0].len();
    let height = rows.len();
    let block = RangeData {
        width,
        height,
        cells: rows.into_iter().flatten().collect(),
    };
    Ok(Arg::Range(if sideways { on_its_side(&block) } else { block }))
}

/// Which of two values comes first when a block is being put in order.
///
/// Excel ranks the KINDS before it compares within one — numbers, then text,
/// then the logicals, then the errors — so an error in the column being sorted
/// by is something to place rather than something to refuse.
///
/// A blank goes last whichever way round the sort is, which is why it cannot
/// simply be given the highest rank: it takes no part in the reversal.
fn in_order(left: &Value, right: &Value, descending: bool) -> Ordering {
    match (left.is_blank(), right.is_blank()) {
        (true, true) => return Ordering::Equal,
        (true, false) => return Ordering::Greater,
        (false, true) => return Ordering::Less,
        _ => {}
    }
    let side = sorting_rank(left)
        .cmp(&sorting_rank(right))
        // Within one kind, the ordinary comparison. Two errors are left as
        // they were: a stable sort keeps them in the order they arrived, and
        // whether Excel puts one error above another was not measured.
        .then_with(|| compare(left, right).unwrap_or(Ordering::Equal));
    if descending {
        side.reverse()
    } else {
        side
    }
}

/// Which kind of value this is, for the purpose of ordering a block.
fn sorting_rank(value: &Value) -> u8 {
    match value {
        Value::Number(_) => 0,
        Value::Text(_) => 1,
        Value::Logical(_) => 2,
        Value::Error(_) => 3,
        // Handled before the rank is asked for.
        Value::Blank => 4,
    }
}

/// The block with one more column on the end, a value to each row.
fn with_a_column(block: &RangeData, beside: &[Value]) -> RangeData {
    let mut cells = Vec::with_capacity(block.cells.len() + beside.len());
    for (row, alongside) in beside.iter().enumerate().take(block.height) {
        for col in 0..block.width {
            cells.push(block.at(col, row));
        }
        cells.push(alongside.clone());
    }
    RangeData {
        width: block.width + 1,
        height: block.height,
        cells,
    }
}

/// The same block with its rows and columns exchanged.
fn on_its_side(block: &RangeData) -> RangeData {
    let mut cells = Vec::with_capacity(block.cells.len());
    for col in 0..block.width {
        for row in 0..block.height {
            cells.push(block.at(col, row));
        }
    }
    RangeData {
        width: block.height,
        height: block.width,
        cells,
    }
}

/// An optional argument that has to be true to count, and is false when it is
/// not there.
fn reads_true(arg: Option<&Arg>) -> bool {
    arg.map(|one| one.scalar().to_logical().unwrap_or(false))
        .unwrap_or(false)
}

/// Two rows holding the same things. UNIQUE compares whole rows, so two rows
/// alike in every column are one row twice.
fn same_row(a: &[Value], b: &[Value]) -> bool {
    a.len() == b.len()
        && a.iter().zip(b).all(|(one, other)| match (one, other) {
            // Comparing two errors is not a comparison, but two of the SAME
            // error are plainly the same value, and UNIQUE has to see that.
            (Value::Error(why), Value::Error(also)) => why == also,
            _ => crate::value::same_ignoring_width(one, other),
        })
}

/// The line an INDEX asks for when it leaves out a row or a column, or `None`
/// when it is addressing one cell after all.
///
/// `INDEX(range,,3)` and `INDEX(range,0,3)` are the same request: the third
/// column entire. `INDEX(range,3,0)` is the third row. `INDEX(range,0,0)` is
/// everything. Anything else is one cell and belongs to `index_at`.
fn a_whole_line(args: &[Arg]) -> Option<Arg> {
    if args.len() < 3 {
        return None;
    }
    let table = args[0].as_range();
    let (row, col) = (a_missing_index(&args[1])?, a_missing_index(&args[2])?);
    let taken = |rows: std::ops::Range<usize>, cols: std::ops::Range<usize>| {
        let (width, height) = (cols.len(), rows.len());
        let mut cells = Vec::with_capacity(width * height);
        for r in rows {
            for c in cols.clone() {
                cells.push(table.at(c, r));
            }
        }
        Arg::Range(RangeData {
            width,
            height,
            cells,
        })
    };
    match (row, col) {
        (0, 0) => Some(taken(0..table.height, 0..table.width)),
        (0, col) if col <= table.width => Some(taken(0..table.height, col - 1..col)),
        (row, 0) if row <= table.height => Some(taken(row - 1..row, 0..table.width)),
        // A line outside the range is the ordinary `#REF!`, which `index_at`
        // already says.
        (0, _) | (_, 0) => Some(Arg::Value(Value::Error(ExcelError::Ref))),
        _ => None,
    }
}

/// What an index argument says, when it says a whole line: an omitted argument
/// and an explicit zero both do. A number, a range, or anything unreadable
/// does not, and `None` leaves the ordinary path to deal with it.
fn a_missing_index(arg: &Arg) -> Option<usize> {
    match arg {
        Arg::Value(Value::Blank) => Some(0),
        Arg::Value(Value::Number(n)) if *n >= 0.0 => Some(*n as usize),
        _ => None,
    }
}

fn index_at(table: &RangeData, row: usize, col: usize) -> Result<Value, ExcelError> {
    if row < 1 || col < 1 || row > table.height || col > table.width {
        return Err(ExcelError::Ref);
    }
    Ok(table.at(col - 1, row - 1))
}

/// A block built from rows.
fn block_from_rows(rows: Vec<Vec<Value>>) -> Result<RangeData, ExcelError> {
    let height = rows.len();
    let width = rows.first().map_or(0, Vec::len);
    if height == 0 || width == 0 {
        return Err(ExcelError::Value);
    }
    Ok(RangeData { width, height, cells: rows.into_iter().flatten().collect() })
}

fn rows_of(block: &RangeData) -> Vec<Vec<Value>> {
    (0..block.height)
        .map(|row| (0..block.width).map(|col| block.at(col, row)).collect())
        .collect()
}

/// An optional whole-number argument, None when omitted or blank.
fn optional_count(args: &[Arg], at: usize) -> Result<Option<i64>, ExcelError> {
    match args.get(at) {
        None => Ok(None),
        Some(arg) => match arg.scalar() {
            Value::Blank => Ok(None),
            other => Ok(Some(other.to_number()?.trunc() as i64)),
        },
    }
}

fn optional_pad(args: &[Arg], at: usize) -> Value {
    match args.get(at).map(Arg::scalar) {
        None | Some(Value::Blank) => Value::Error(ExcelError::NA),
        Some(value) => value,
    }
}

/// The first or last `n` of `len` indexes; None for all.
fn kept(len: usize, n: Option<i64>, take: bool) -> Result<Vec<usize>, ExcelError> {
    let Some(n) = n else {
        return Ok((0..len).collect());
    };
    let count = (n.unsigned_abs() as usize).min(len);
    let indexes: Vec<usize> = match (take, n >= 0) {
        (true, true) => (0..count).collect(),
        (true, false) => (len - count..len).collect(),
        (false, true) => (count..len).collect(),
        (false, false) => (0..len - count).collect(),
    };
    if n == 0 && take || indexes.is_empty() {
        return Err(ExcelError::Value);
    }
    Ok(indexes)
}

/// Picks by 1-based index, negative counting from the end.
fn picked(len: usize, wanted: &[Arg]) -> Result<Vec<usize>, ExcelError> {
    let mut out = Vec::new();
    for arg in wanted {
        for value in arg.flatten() {
            let at = value.to_number()?.trunc() as i64;
            let index = if at > 0 { at - 1 } else { len as i64 + at };
            if at == 0 || index < 0 || index >= len as i64 {
                return Err(ExcelError::Value);
            }
            out.push(index as usize);
        }
    }
    if out.is_empty() {
        return Err(ExcelError::Value);
    }
    Ok(out)
}

/// SEQUENCE, TAKE, DROP, CHOOSEROWS, CHOOSECOLS, VSTACK, HSTACK, TOCOL, TOROW,
/// WRAPROWS, WRAPCOLS, EXPAND and TEXTSPLIT. Measured against Excel: a short
/// block in a stack or a wrap is padded with #N/A.
fn reshaped(name: &str, args: &[Arg]) -> Result<RangeData, ExcelError> {
    match name {
        "SEQUENCE" => {
            expect(args, 1)?;
            let rows = num(&args[0])?.trunc();
            let cols = optional_count(args, 1)?.unwrap_or(1) as f64;
            let start = match args.get(2).map(Arg::scalar) {
                None | Some(Value::Blank) => 1.0,
                Some(v) => v.to_number()?,
            };
            let step = match args.get(3).map(Arg::scalar) {
                None | Some(Value::Blank) => 1.0,
                Some(v) => v.to_number()?,
            };
            if rows < 1.0 || cols < 1.0 || rows * cols > 1_048_576.0 {
                return Err(ExcelError::Value);
            }
            let (rows, cols) = (rows as usize, cols as usize);
            Ok(RangeData {
                width: cols,
                height: rows,
                cells: (0..rows * cols).map(|at| Value::Number(start + step * at as f64)).collect(),
            })
        }
        "TAKE" | "DROP" => {
            expect(args, 2)?;
            let block = args[0].as_range();
            let take = name == "TAKE";
            let rows = kept(block.height, optional_count(args, 1)?, take)?;
            let cols = kept(block.width, optional_count(args, 2)?, take)?;
            block_from_rows(
                rows.iter()
                    .map(|row| cols.iter().map(|col| block.at(*col, *row)).collect())
                    .collect(),
            )
        }
        "CHOOSEROWS" | "CHOOSECOLS" => {
            expect(args, 2)?;
            let block = args[0].as_range();
            if name == "CHOOSEROWS" {
                let rows = picked(block.height, &args[1..])?;
                block_from_rows(
                    rows.iter()
                        .map(|row| (0..block.width).map(|col| block.at(col, *row)).collect())
                        .collect(),
                )
            } else {
                let cols = picked(block.width, &args[1..])?;
                block_from_rows(
                    (0..block.height)
                        .map(|row| cols.iter().map(|col| block.at(*col, row)).collect())
                        .collect(),
                )
            }
        }
        "VSTACK" | "HSTACK" => {
            expect(args, 1)?;
            let blocks: Vec<RangeData> = args.iter().map(Arg::as_range).collect();
            let na = Value::Error(ExcelError::NA);
            if name == "VSTACK" {
                let width = blocks.iter().map(|b| b.width).max().unwrap_or(0);
                let mut rows = Vec::new();
                for block in &blocks {
                    for row in rows_of(block) {
                        let mut row = row;
                        row.resize(width, na.clone());
                        rows.push(row);
                    }
                }
                block_from_rows(rows)
            } else {
                let height = blocks.iter().map(|b| b.height).max().unwrap_or(0);
                let rows = (0..height)
                    .map(|row| {
                        blocks
                            .iter()
                            .flat_map(|block| {
                                (0..block.width).map(move |col| {
                                    if row < block.height { block.at(col, row) } else { Value::Error(ExcelError::NA) }
                                })
                            })
                            .collect()
                    })
                    .collect();
                block_from_rows(rows)
            }
        }
        "TOCOL" | "TOROW" => {
            expect(args, 1)?;
            let block = args[0].as_range();
            let ignore = optional_count(args, 1)?.unwrap_or(0);
            let by_column = match args.get(2).map(Arg::scalar) {
                None | Some(Value::Blank) => false,
                Some(v) => v.to_logical()?,
            };
            let mut values = Vec::new();
            let (outer, inner) = if by_column { (block.width, block.height) } else { (block.height, block.width) };
            for a in 0..outer {
                for b in 0..inner {
                    let value = if by_column { block.at(a, b) } else { block.at(b, a) };
                    let skip = match value {
                        Value::Blank => ignore == 1 || ignore == 3,
                        Value::Error(_) => ignore == 2 || ignore == 3,
                        _ => false,
                    };
                    if !skip {
                        values.push(value);
                    }
                }
            }
            if values.is_empty() {
                return Err(ExcelError::Value);
            }
            let n = values.len();
            Ok(if name == "TOCOL" {
                RangeData { width: 1, height: n, cells: values }
            } else {
                RangeData { width: n, height: 1, cells: values }
            })
        }
        "WRAPROWS" | "WRAPCOLS" => {
            expect(args, 2)?;
            let values = args[0].flatten();
            let size = num(&args[1])?.trunc();
            if size < 1.0 {
                return Err(ExcelError::Num);
            }
            let size = size as usize;
            let pad = optional_pad(args, 2);
            let lines = values.len().div_ceil(size);
            let line = |at: usize| -> Vec<Value> {
                (0..size).map(|k| values.get(at * size + k).cloned().unwrap_or_else(|| pad.clone())).collect()
            };
            if name == "WRAPROWS" {
                block_from_rows((0..lines).map(line).collect())
            } else {
                let columns: Vec<Vec<Value>> = (0..lines).map(line).collect();
                block_from_rows((0..size).map(|row| columns.iter().map(|c| c[row].clone()).collect()).collect())
            }
        }
        "EXPAND" => {
            expect(args, 2)?;
            let block = args[0].as_range();
            let rows = optional_count(args, 1)?.unwrap_or(block.height as i64);
            let cols = optional_count(args, 2)?.unwrap_or(block.width as i64);
            if rows < block.height as i64 || cols < block.width as i64 {
                return Err(ExcelError::Value);
            }
            let pad = optional_pad(args, 3);
            block_from_rows(
                (0..rows as usize)
                    .map(|row| {
                        (0..cols as usize)
                            .map(|col| {
                                if row < block.height && col < block.width { block.at(col, row) } else { pad.clone() }
                            })
                            .collect()
                    })
                    .collect(),
            )
        }
        "TEXTSPLIT" => {
            expect(args, 2)?;
            let whole = text(&args[0])?;
            let marks = |at: usize| -> Result<Vec<String>, ExcelError> {
                match args.get(at) {
                    None => Ok(Vec::new()),
                    Some(arg) => arg
                        .flatten()
                        .into_iter()
                        .filter(|v| !v.is_blank())
                        .map(|v| v.to_text())
                        .collect(),
                }
            };
            let (columns, rows) = (marks(1)?, marks(2)?);
            let ignore_empty = match args.get(3).map(Arg::scalar) {
                None | Some(Value::Blank) => false,
                Some(v) => v.to_logical()?,
            };
            let pad = optional_pad(args, 5);
            let split = |text: &str, by: &[String]| -> Vec<String> {
                if by.iter().all(String::is_empty) {
                    return vec![text.to_string()];
                }
                let mut pieces = Vec::new();
                let mut rest = text;
                loop {
                    let next = by
                        .iter()
                        .filter(|d| !d.is_empty())
                        .filter_map(|d| rest.find(d.as_str()).map(|at| (at, d.len())))
                        .min();
                    match next {
                        Some((at, len)) => {
                            pieces.push(rest[..at].to_string());
                            rest = &rest[at + len..];
                        }
                        None => {
                            pieces.push(rest.to_string());
                            break;
                        }
                    }
                }
                if ignore_empty {
                    pieces.retain(|p| !p.is_empty());
                }
                pieces
            };
            let lines: Vec<Vec<String>> = split(&whole, &rows).iter().map(|line| split(line, &columns)).collect();
            let width = lines.iter().map(Vec::len).max().unwrap_or(0);
            block_from_rows(
                lines
                    .into_iter()
                    .map(|line| {
                        let mut row: Vec<Value> = line.into_iter().map(Value::Text).collect();
                        row.resize(width, pad.clone());
                        row
                    })
                    .collect(),
            )
        }
        _ => Err(ExcelError::Name),
    }
}

/// A finite answer, or #NUM! for one that ran off to infinity.
fn fin(n: f64) -> Result<Value, ExcelError> {
    if n.is_finite() {
        Ok(Value::Number(n))
    } else {
        Err(ExcelError::Num)
    }
}

/// How many digits a base's ten-digit two's complement spans, and so where a
/// negative number starts.
fn base_bits(radix: u32) -> u32 {
    match radix {
        2 => 10,
        8 => 30,
        _ => 40,
    }
}

/// DEC2HEX and its kin: a negative number is written in ten digits of two's
/// complement and ignores `places`; a positive one is padded to `places`,
/// which may not be fewer than it needs.
fn to_base(n: i64, radix: u32, places: Option<&Arg>) -> Result<Value, ExcelError> {
    let bits = base_bits(radix);
    let half = 1i64 << (bits - 1);
    if n < -half || n >= half {
        return Err(ExcelError::Num);
    }
    if n < 0 {
        let held = (n + (1i64 << bits)) as u64;
        return Ok(Value::Text(format_radix(held, radix)));
    }
    let mut written = format_radix(n as u64, radix);
    if let Some(places) = places {
        let places = num(places)?.trunc();
        if places < written.len() as f64 || places > 10.0 {
            return Err(ExcelError::Num);
        }
        while written.len() < places as usize {
            written.insert(0, '0');
        }
    }
    Ok(Value::Text(written))
}

fn format_radix(mut n: u64, radix: u32) -> String {
    if n == 0 {
        return "0".to_string();
    }
    let mut digits = Vec::new();
    while n > 0 {
        digits.push(std::char::from_digit((n % radix as u64) as u32, radix).unwrap().to_ascii_uppercase());
        n /= radix as u64;
    }
    digits.iter().rev().collect()
}

/// HEX2DEC and its kin: up to ten digits, the tenth-digit sign bit read as
/// two's complement.
fn from_base(text: &str, radix: u32) -> Result<i64, ExcelError> {
    let text = text.trim();
    if text.len() > 10 {
        return Err(ExcelError::Num);
    }
    if text.is_empty() {
        return Ok(0);
    }
    let n = i64::from_str_radix(text, radix).map_err(|_| ExcelError::Num)?;
    let bits = base_bits(radix);
    Ok(if n >= 1i64 << (bits - 1) { n - (1i64 << bits) } else { n })
}

/// Sums of squares for two paired samples.
struct Fit {
    n: f64,
    mean_x: f64,
    mean_y: f64,
    sxx: f64,
    syy: f64,
    sxy: f64,
}

impl Fit {
    /// `ys` then `xs`, the order SLOPE and CORREL are written in.
    fn of(ys: &Arg, xs: &Arg) -> Result<Fit, ExcelError> {
        let (ys, xs) = (ys.flatten(), xs.flatten());
        if ys.len() != xs.len() {
            return Err(ExcelError::NA);
        }
        let mut pairs = Vec::new();
        for (y, x) in ys.iter().zip(&xs) {
            if let Value::Error(e) = y {
                return Err(*e);
            }
            if let Value::Error(e) = x {
                return Err(*e);
            }
            if let (Value::Number(y), Value::Number(x)) = (y, x) {
                pairs.push((*x, *y));
            }
        }
        if pairs.is_empty() {
            return Err(ExcelError::DivZero);
        }
        let n = pairs.len() as f64;
        let mean_x = pairs.iter().map(|p| p.0).sum::<f64>() / n;
        let mean_y = pairs.iter().map(|p| p.1).sum::<f64>() / n;
        let (mut sxx, mut syy, mut sxy) = (0.0, 0.0, 0.0);
        for (x, y) in pairs {
            sxx += (x - mean_x) * (x - mean_x);
            syy += (y - mean_y) * (y - mean_y);
            sxy += (x - mean_x) * (y - mean_y);
        }
        Ok(Fit { n, mean_x, mean_y, sxx, syy, sxy })
    }

    fn slope(&self) -> Result<f64, ExcelError> {
        if self.sxx == 0.0 {
            return Err(ExcelError::DivZero);
        }
        Ok(self.sxy / self.sxx)
    }

    fn correl(&self) -> Result<f64, ExcelError> {
        if self.sxx == 0.0 || self.syy == 0.0 {
            return Err(ExcelError::DivZero);
        }
        Ok(self.sxy / (self.sxx * self.syy).sqrt())
    }
}

/// TREND(known_y, [known_x], [new_x], [const]): a least-squares line with one
/// x, through the origin when const is FALSE. Measured: y 2,4,7 over x 1,2,3
/// gives 9.333 at 4, and 11.071 at 5 through the origin.
fn trend(args: &[Arg]) -> Result<RangeData, ExcelError> {
    expect(args, 1)?;
    let numbers = |arg: &Arg| -> Result<Vec<f64>, ExcelError> {
        arg.flatten().into_iter().map(|v| v.to_number()).collect()
    };
    let ys = numbers(&args[0])?;
    let given_x = |at: usize| match args.get(at) {
        Some(Arg::Value(Value::Blank)) | None => None,
        Some(arg) => Some(arg),
    };
    let xs = match given_x(1) {
        Some(arg) => numbers(arg)?,
        None => (1..=ys.len()).map(|n| n as f64).collect(),
    };
    if xs.len() != ys.len() || ys.is_empty() {
        return Err(ExcelError::Ref);
    }
    let with_constant = match args.get(3).map(Arg::scalar) {
        None | Some(Value::Blank) => true,
        Some(v) => v.to_logical()?,
    };
    let n = ys.len() as f64;
    let (slope, intercept) = if with_constant {
        let (mx, my) = (xs.iter().sum::<f64>() / n, ys.iter().sum::<f64>() / n);
        let sxx: f64 = xs.iter().map(|x| (x - mx) * (x - mx)).sum();
        let sxy: f64 = xs.iter().zip(&ys).map(|(x, y)| (x - mx) * (y - my)).sum();
        if sxx == 0.0 {
            return Err(ExcelError::DivZero);
        }
        (sxy / sxx, my - sxy / sxx * mx)
    } else {
        let sxx: f64 = xs.iter().map(|x| x * x).sum();
        if sxx == 0.0 {
            return Err(ExcelError::DivZero);
        }
        (xs.iter().zip(&ys).map(|(x, y)| x * y).sum::<f64>() / sxx, 0.0)
    };
    let (width, height, new_xs) = match given_x(2) {
        Some(arg) => {
            let block = arg.as_range();
            (block.width, block.height, numbers(arg)?)
        }
        None => match given_x(1) {
            Some(arg) => {
                let block = arg.as_range();
                (block.width, block.height, xs.clone())
            }
            None => {
                let block = args[0].as_range();
                (block.width, block.height, xs.clone())
            }
        },
    };
    Ok(RangeData {
        width,
        height,
        cells: new_xs.iter().map(|x| Value::Number(intercept + slope * x)).collect(),
    })
}

/// FREQUENCY: how many of the data fall at or under each bin and above the
/// one before, then how many are above them all. Bins are counted in rising
/// order and answered in the order given.
fn frequency(args: &[Arg]) -> Result<RangeData, ExcelError> {
    expect(args, 2)?;
    let data: Vec<f64> = args[0]
        .flatten()
        .into_iter()
        .filter_map(|v| if let Value::Number(n) = v { Some(n) } else { None })
        .collect();
    let bins: Vec<f64> = args[1]
        .flatten()
        .into_iter()
        .filter_map(|v| if let Value::Number(n) = v { Some(n) } else { None })
        .collect();
    let mut order: Vec<usize> = (0..bins.len()).collect();
    order.sort_by(|a, b| bins[*a].partial_cmp(&bins[*b]).unwrap_or(std::cmp::Ordering::Equal));
    let mut counts = vec![0.0; bins.len() + 1];
    for value in data {
        match order.iter().find(|at| value <= bins[**at]) {
            Some(at) => counts[*at] += 1.0,
            None => counts[bins.len()] += 1.0,
        }
    }
    Ok(RangeData {
        width: 1,
        height: counts.len(),
        cells: counts.into_iter().map(Value::Number).collect(),
    })
}

/// Monday 0 .. Sunday 6.
fn monday_zero(day: i64) -> Result<usize, ExcelError> {
    Ok((weekday_with_type(day, 2)? - 1) as usize)
}

/// Which days of the week are the weekend, Monday first: a code 1-7 for a
/// pair (1 Saturday and Sunday, 2 Sunday and Monday ...), 11-17 for one day
/// (11 Sunday, 12 Monday ...), or seven 0s and 1s from Monday.
fn weekend_days(arg: Option<&Arg>) -> Result<[bool; 7], ExcelError> {
    let mut days = [false; 7];
    let Some(arg) = arg else {
        days[5] = true;
        days[6] = true;
        return Ok(days);
    };
    match arg.scalar() {
        Value::Blank => {
            days[5] = true;
            days[6] = true;
        }
        Value::Text(mask) => {
            // Every day a weekend is allowed: measured, NETWORKDAYS.INTL
            // with "1111111" is 0.
            if mask.len() != 7 || !mask.chars().all(|c| c == '0' || c == '1') {
                return Err(ExcelError::Value);
            }
            for (at, c) in mask.chars().enumerate() {
                days[at] = c == '1';
            }
        }
        other => {
            let code = other.to_number()? as i64;
            match code {
                // 1 is Saturday-Sunday; each code after moves the pair a day on.
                1..=7 => {
                    let first = (code as usize + 4) % 7;
                    days[first] = true;
                    days[(first + 1) % 7] = true;
                }
                11..=17 => days[(code as usize + 2) % 7] = true,
                _ => return Err(ExcelError::Num),
            }
        }
    }
    Ok(days)
}

fn holiday_serials(arg: Option<&Arg>) -> Result<Vec<i64>, ExcelError> {
    let mut held = Vec::new();
    if let Some(given) = arg {
        for one in given.flatten() {
            if one.is_blank() {
                continue;
            }
            held.push(serial(&Arg::Value(one))?);
        }
    }
    Ok(held)
}

/// A character's width in Shift_JIS bytes.
fn byte_width(c: char) -> usize {
    let code = c as u32;
    if code < 0x80 || (0xFF61..=0xFF9F).contains(&code) {
        1
    } else {
        2
    }
}

fn bytes_of(t: &str) -> usize {
    t.chars().map(byte_width).sum()
}

/// The bytes `from` (1-based) onward for `count`, a character cut in two
/// written as a space for the half inside.
fn bytes_between(t: &str, from: usize, count: usize) -> String {
    if count == 0 || from == 0 {
        return String::new();
    }
    let last = from + count - 1;
    let mut out = String::new();
    let mut at = 1usize;
    for c in t.chars() {
        let width = byte_width(c);
        let end = at + width - 1;
        if end >= from && at <= last {
            if at >= from && end <= last {
                out.push(c);
            } else {
                out.push(' ');
            }
        }
        at += width;
        if at > last {
            break;
        }
    }
    out
}

fn expect(args: &[Arg], n: usize) -> Result<(), ExcelError> {
    if args.len() < n {
        Err(ExcelError::Value)
    } else {
        Ok(())
    }
}

fn one_arg(args: &[Arg]) -> Result<&Arg, ExcelError> {
    args.first().ok_or(ExcelError::Value)
}

fn one(args: &[Arg]) -> Result<f64, ExcelError> {
    num(one_arg(args)?)
}

fn one_value(args: &[Arg]) -> Value {
    args.first().map(|a| a.scalar()).unwrap_or(Value::Blank)
}

/// The integer day part of a date argument. Excel truncates toward zero.
fn serial(arg: &Arg) -> Result<i64, ExcelError> {
    let n = num(arg)?;
    if n < 0.0 {
        return Err(ExcelError::Num);
    }
    Ok(n.floor() as i64)
}

/// YEARFRAC on Excel's five day-count bases. The dates are put in order first,
/// so it is symmetric as Excel is.
pub(crate) fn yearfrac(start: i64, end: i64, basis: i64) -> Result<f64, ExcelError> {
    let (start, end) = if start <= end { (start, end) } else { (end, start) };
    if start == end {
        return Ok(0.0);
    }
    match basis {
        // 30/360, US (NASD) and European. The two differ only in how they
        // pull a day of 31 -- and whether the last day of February counts.
        0 => Ok(days_30_360(start, end, false)? as f64 / 360.0),
        4 => Ok(days_30_360(start, end, true)? as f64 / 360.0),
        2 => Ok((end - start) as f64 / 360.0),
        3 => Ok((end - start) as f64 / 365.0),
        1 => {
            let (d1, d2) = (datetime::date_from_serial(start)?, datetime::date_from_serial(end)?);
            let denom = if d1.year == d2.year {
                days_in_year(d1.year)? as f64
            } else if d2.year == d1.year + 1
                && (d1.month > d2.month || (d1.month == d2.month && d1.day >= d2.day))
            {
                // A span of a year or less that crosses one year boundary: the
                // denominator is 366 when a 29 February falls within it, else
                // 365. Measured: 1 Mar 2023 -> 1 Mar 2024 is exactly 1.
                if feb29_within(start, end)? { 366.0 } else { 365.0 }
            } else {
                // A longer span: the average length of the years it touches.
                let touched = datetime::serial_from_date(d2.year + 1, 1, 1)?
                    - datetime::serial_from_date(d1.year, 1, 1)?;
                touched as f64 / (d2.year - d1.year + 1) as f64
            };
            Ok((end - start) as f64 / denom)
        }
        _ => Err(ExcelError::Num),
    }
}

fn days_in_year(year: i64) -> Result<i64, ExcelError> {
    Ok(datetime::serial_from_date(year + 1, 1, 1)? - datetime::serial_from_date(year, 1, 1)?)
}

fn feb29_within(start: i64, end: i64) -> Result<bool, ExcelError> {
    let (y1, y2) = (
        datetime::date_from_serial(start)?.year,
        datetime::date_from_serial(end)?.year,
    );
    for year in y1..=y2 {
        // A 29 February exists only in a leap year; serial_from_date refuses
        // it otherwise, which is the leap-year test.
        if let Ok(leap_day) = datetime::serial_from_date(year, 2, 29) {
            if (start..=end).contains(&leap_day) {
                return Ok(true);
            }
        }
    }
    Ok(false)
}

/// The last day of February -- the day whose next day leaves the month.
fn is_last_of_feb(serial: i64) -> Result<bool, ExcelError> {
    let date = datetime::date_from_serial(serial)?;
    Ok(date.month == 2 && datetime::date_from_serial(serial + 1)?.month != 2)
}

/// The half-width form of a full-width space or katakana character, for ASC.
/// Full-width ASCII is a plain offset the caller handles. Generated from
/// Unicode's compatibility mapping and checked against Excel; the combining
/// voiced marks U+3099/U+309A are deliberately absent, as Excel leaves them.
fn asc_halfwidth(c: char) -> Option<&'static str> {
    Some(match c {
        '\u{3000}' => "\u{20}",
        '\u{3001}' => "\u{FF64}",
        '\u{3002}' => "\u{FF61}",
        '\u{300C}' => "\u{FF62}",
        '\u{300D}' => "\u{FF63}",
        '\u{309B}' => "\u{FF9E}",
        '\u{309C}' => "\u{FF9F}",
        '\u{30A1}' => "\u{FF67}",
        '\u{30A2}' => "\u{FF71}",
        '\u{30A3}' => "\u{FF68}",
        '\u{30A4}' => "\u{FF72}",
        '\u{30A5}' => "\u{FF69}",
        '\u{30A6}' => "\u{FF73}",
        '\u{30A7}' => "\u{FF6A}",
        '\u{30A8}' => "\u{FF74}",
        '\u{30A9}' => "\u{FF6B}",
        '\u{30AA}' => "\u{FF75}",
        '\u{30AB}' => "\u{FF76}",
        '\u{30AC}' => "\u{FF76}\u{FF9E}",
        '\u{30AD}' => "\u{FF77}",
        '\u{30AE}' => "\u{FF77}\u{FF9E}",
        '\u{30AF}' => "\u{FF78}",
        '\u{30B0}' => "\u{FF78}\u{FF9E}",
        '\u{30B1}' => "\u{FF79}",
        '\u{30B2}' => "\u{FF79}\u{FF9E}",
        '\u{30B3}' => "\u{FF7A}",
        '\u{30B4}' => "\u{FF7A}\u{FF9E}",
        '\u{30B5}' => "\u{FF7B}",
        '\u{30B6}' => "\u{FF7B}\u{FF9E}",
        '\u{30B7}' => "\u{FF7C}",
        '\u{30B8}' => "\u{FF7C}\u{FF9E}",
        '\u{30B9}' => "\u{FF7D}",
        '\u{30BA}' => "\u{FF7D}\u{FF9E}",
        '\u{30BB}' => "\u{FF7E}",
        '\u{30BC}' => "\u{FF7E}\u{FF9E}",
        '\u{30BD}' => "\u{FF7F}",
        '\u{30BE}' => "\u{FF7F}\u{FF9E}",
        '\u{30BF}' => "\u{FF80}",
        '\u{30C0}' => "\u{FF80}\u{FF9E}",
        '\u{30C1}' => "\u{FF81}",
        '\u{30C2}' => "\u{FF81}\u{FF9E}",
        '\u{30C3}' => "\u{FF6F}",
        '\u{30C4}' => "\u{FF82}",
        '\u{30C5}' => "\u{FF82}\u{FF9E}",
        '\u{30C6}' => "\u{FF83}",
        '\u{30C7}' => "\u{FF83}\u{FF9E}",
        '\u{30C8}' => "\u{FF84}",
        '\u{30C9}' => "\u{FF84}\u{FF9E}",
        '\u{30CA}' => "\u{FF85}",
        '\u{30CB}' => "\u{FF86}",
        '\u{30CC}' => "\u{FF87}",
        '\u{30CD}' => "\u{FF88}",
        '\u{30CE}' => "\u{FF89}",
        '\u{30CF}' => "\u{FF8A}",
        '\u{30D0}' => "\u{FF8A}\u{FF9E}",
        '\u{30D1}' => "\u{FF8A}\u{FF9F}",
        '\u{30D2}' => "\u{FF8B}",
        '\u{30D3}' => "\u{FF8B}\u{FF9E}",
        '\u{30D4}' => "\u{FF8B}\u{FF9F}",
        '\u{30D5}' => "\u{FF8C}",
        '\u{30D6}' => "\u{FF8C}\u{FF9E}",
        '\u{30D7}' => "\u{FF8C}\u{FF9F}",
        '\u{30D8}' => "\u{FF8D}",
        '\u{30D9}' => "\u{FF8D}\u{FF9E}",
        '\u{30DA}' => "\u{FF8D}\u{FF9F}",
        '\u{30DB}' => "\u{FF8E}",
        '\u{30DC}' => "\u{FF8E}\u{FF9E}",
        '\u{30DD}' => "\u{FF8E}\u{FF9F}",
        '\u{30DE}' => "\u{FF8F}",
        '\u{30DF}' => "\u{FF90}",
        '\u{30E0}' => "\u{FF91}",
        '\u{30E1}' => "\u{FF92}",
        '\u{30E2}' => "\u{FF93}",
        '\u{30E3}' => "\u{FF6C}",
        '\u{30E4}' => "\u{FF94}",
        '\u{30E5}' => "\u{FF6D}",
        '\u{30E6}' => "\u{FF95}",
        '\u{30E7}' => "\u{FF6E}",
        '\u{30E8}' => "\u{FF96}",
        '\u{30E9}' => "\u{FF97}",
        '\u{30EA}' => "\u{FF98}",
        '\u{30EB}' => "\u{FF99}",
        '\u{30EC}' => "\u{FF9A}",
        '\u{30ED}' => "\u{FF9B}",
        '\u{30EF}' => "\u{FF9C}",
        '\u{30F2}' => "\u{FF66}",
        '\u{30F3}' => "\u{FF9D}",
        '\u{30F4}' => "\u{FF73}\u{FF9E}",
        '\u{30F7}' => "\u{FF9C}\u{FF9E}",
        '\u{30FA}' => "\u{FF66}\u{FF9E}",
        '\u{30FB}' => "\u{FF65}",
        '\u{30FC}' => "\u{FF70}",
        _ => return None,
    })
}

/// The 360-day count, US (NASD) or European. Distinct from YEARFRAC basis 0:
/// here the D2 test runs AFTER the last-of-February start is pulled to 30, so
/// 29 Feb -> 31 Mar is 30 days, where YEARFRAC basis 0 makes it 31.
fn days360(start: i64, end: i64, european: bool) -> Result<i64, ExcelError> {
    let (a, b) = (
        datetime::date_from_serial(start)?,
        datetime::date_from_serial(end)?,
    );
    let (mut d1, mut d2) = (a.day, b.day);
    if european {
        if d1 == 31 {
            d1 = 30;
        }
        if d2 == 31 {
            d2 = 30;
        }
    } else {
        if is_last_of_feb(start)? {
            d1 = 30;
        }
        if d1 == 31 {
            d1 = 30;
        }
        if d2 == 31 && d1 == 30 {
            d2 = 30;
        }
    }
    Ok((b.year - a.year) * 360 + (b.month - a.month) * 30 + (d2 - d1))
}

/// The 30/360 day count. In the US (NASD) form the order matters and the D2
/// test reads the ORIGINAL D1 -- measured: 29 Feb 2024 -> 31 Mar 2024 is 31
/// days, not 30. The European form just pulls any 31 down to 30.
pub(crate) fn days_30_360(start: i64, end: i64, european: bool) -> Result<i64, ExcelError> {
    let (a, b) = (datetime::date_from_serial(start)?, datetime::date_from_serial(end)?);
    let (mut d1, mut d2) = (a.day, b.day);
    if european {
        if d1 == 31 {
            d1 = 30;
        }
        if d2 == 31 {
            d2 = 30;
        }
    } else {
        if is_last_of_feb(start)? && is_last_of_feb(end)? {
            d2 = 30;
        }
        if d2 == 31 && (d1 == 30 || d1 == 31) {
            d2 = 30;
        }
        if d1 == 31 {
            d1 = 30;
        }
        if is_last_of_feb(start)? {
            d1 = 30;
        }
    }
    Ok((b.year - a.year) * 360 + (b.month - a.month) * 30 + (d2 - d1))
}

/// Map Excel's `WEEKDAY` return-type codes onto a day number.
///
/// Types 1 and 17 start the week on Sunday, 2 and 11 on Monday, 12..=16 walk
/// the start day forward, and type 3 is the only zero-based variant.
fn weekday_with_type(serial: i64, kind: i64) -> Result<i64, ExcelError> {
    let sunday_zero = datetime::weekday_sunday_one(serial) - 1;
    let shifted = |start: i64| (sunday_zero + 7 - start).rem_euclid(7) + 1;
    match kind {
        1 | 17 => Ok(shifted(0)),
        2 | 11 => Ok(shifted(1)),
        3 => Ok(shifted(1) - 1),
        12 => Ok(shifted(2)),
        13 => Ok(shifted(3)),
        14 => Ok(shifted(4)),
        15 => Ok(shifted(5)),
        16 => Ok(shifted(6)),
        _ => Err(ExcelError::Num),
    }
}

/// `DATEDIF` unit handling. Excel never documented this function, but Japanese
/// workbooks use it constantly for ages and years of service.
fn datedif(start: i64, end: i64, unit: &str) -> Result<f64, ExcelError> {
    if end < start {
        return Err(ExcelError::Num);
    }
    let a = datetime::date_from_serial(start)?;
    let b = datetime::date_from_serial(end)?;

    // Whole months elapsed, backing off one if the day of month has not arrived.
    let mut months = (b.year - a.year) * 12 + (b.month - a.month);
    if b.day < a.day {
        months -= 1;
    }

    match unit.to_uppercase().as_str() {
        "D" => Ok((end - start) as f64),
        "M" => Ok(months as f64),
        "Y" => Ok((months / 12) as f64),
        "YM" => Ok((months % 12) as f64),
        // The days past the last whole month, counted Excel's way: when the
        // end's day is the smaller, the month BEFORE the end lends its days.
        // Measured: 31 Jan 2024 to 1 Mar 2024 is -1 (1 + 29 - 31).
        "MD" => {
            if b.day >= a.day {
                return Ok((b.day - a.day) as f64);
            }
            let (year, month) = if b.month == 1 { (b.year - 1, 12) } else { (b.year, b.month - 1) };
            Ok((b.day + datetime::days_in_month(year, month) - a.day) as f64)
        }
        // The days since the start's month and day last came round, the
        // start's date moved into the end's year as DATE would move it:
        // measured, 29 Feb 2024 to 1 Mar 2025 is 0 (DATE(2025,2,29) is
        // 1 Mar) and to 28 Feb 2025 is 365.
        "YD" => {
            let mut anchor = datetime::serial_from_date(b.year, a.month, a.day)?;
            if anchor > end {
                anchor = datetime::serial_from_date(b.year - 1, a.month, a.day)?;
            }
            Ok((end - anchor) as f64)
        }
        _ => Err(ExcelError::Num),
    }
}

/// A `COUNTIF`/`SUMIF` criterion such as `">5"`, `"<>x"`, or a bare value.
struct Criteria {
    op: BinaryPredicate,
    operand: Value,
}

enum BinaryPredicate {
    Eq,
    Ne,
    Lt,
    Le,
    Gt,
    Ge,
}

/// The error a criterion spells out, if it spells one.
fn an_error_named(text: &str) -> Option<ExcelError> {
    const NAMED: &[(&str, ExcelError)] = &[
        ("#DIV/0!", ExcelError::DivZero),
        ("#VALUE!", ExcelError::Value),
        ("#NAME?", ExcelError::Name),
        ("#NULL!", ExcelError::Null),
        ("#REF!", ExcelError::Ref),
        ("#NUM!", ExcelError::Num),
        ("#N/A", ExcelError::NA),
    ];
    NAMED
        .iter()
        .find(|(spelled, _)| text.eq_ignore_ascii_case(spelled))
        .map(|(_, why)| *why)
}

impl Criteria {
    fn parse(v: &Value) -> Criteria {
        let text = match v {
            Value::Text(s) => s.clone(),
            other => {
                return Criteria {
                    op: BinaryPredicate::Eq,
                    operand: other.clone(),
                }
            }
        };
        let (op, rest) = if let Some(r) = text.strip_prefix(">=") {
            (BinaryPredicate::Ge, r)
        } else if let Some(r) = text.strip_prefix("<=") {
            (BinaryPredicate::Le, r)
        } else if let Some(r) = text.strip_prefix("<>") {
            (BinaryPredicate::Ne, r)
        } else if let Some(r) = text.strip_prefix('>') {
            (BinaryPredicate::Gt, r)
        } else if let Some(r) = text.strip_prefix('<') {
            (BinaryPredicate::Lt, r)
        } else if let Some(r) = text.strip_prefix('=') {
            (BinaryPredicate::Eq, r)
        } else {
            (BinaryPredicate::Eq, text.as_str())
        };

        // A criterion reads a date the way a typed cell does: measured,
        // `">1/1/2024"` over a cell holding 45292 counts nothing and
        // `">=1/1/2024"` counts it.
        let as_typed = || {
            (!rest.trim().is_empty())
                .then(|| Value::Text(rest.to_string()).to_number().ok())
                .flatten()
        };
        // TRUE and FALSE name the logicals: measured, `COUNTIF(..,"TRUE")`
        // counts a cell holding TRUE.
        if rest.eq_ignore_ascii_case("true") || rest.eq_ignore_ascii_case("false") {
            return Criteria { op, operand: Value::Logical(rest.eq_ignore_ascii_case("true")) };
        }
        let operand = match rest.parse::<f64>().ok().or_else(as_typed) {
            Some(n) => Value::Number(n),
            // `"<>#N/A"` names the error, not the four characters of it.
            None => match an_error_named(rest) {
                Some(why) => Value::Error(why),
                None => Value::Text(rest.to_string()),
            },
        };
        Criteria { op, operand }
    }

    fn matches(&self, v: &Value) -> bool {
        // An error is a value of its own kind: equal to itself, equal to
        // nothing else, and beyond comparing for greater or less. So a
        // NOT-equal criterion IS satisfied by one — `COUNTIF(range,"<>0")`
        // counts an `#N/A` — unless the criterion names that same error.
        let held = match v {
            Value::Error(why) => Some(*why),
            _ => None,
        };
        let wanted = match &self.operand {
            Value::Error(why) => Some(*why),
            _ => None,
        };
        if held.is_some() || wanted.is_some() {
            let alike = held.is_some() && held == wanted;
            return match self.op {
                BinaryPredicate::Eq => alike,
                BinaryPredicate::Ne => !alike,
                _ => false,
            };
        }
        // `""` asks for the empty ones. `COUNTIFS(B:B, x, D:D, "")` — count
        // where D has nothing in it — is how anyone counts what is still
        // outstanding, and a rule that says a blank never matches anything
        // answers nought to all of them.
        if let Value::Text(wanted) = &self.operand {
            if wanted.is_empty() {
                let empty = v.is_blank() || matches!(v, Value::Text(t) if t.is_empty());
                return match self.op {
                    BinaryPredicate::Eq => empty,
                    BinaryPredicate::Ne => !empty,
                    _ => false,
                };
            }
        }
        // Otherwise a blank satisfies no comparison but not-equal: measured,
        // COUNTIF(A2,"<>apple") over an empty A2 is 1.
        if v.is_blank() {
            return matches!(self.op, BinaryPredicate::Ne);
        }
        // `"a*"` asked of COUNTIF means "starting with a", not the two
        // characters. Only equality and inequality read wildcards; `>a*` is
        // a comparison against the literal text.
        if let Value::Text(pattern) = &self.operand {
            if has_wildcards(pattern) {
                if let Value::Text(held) = v {
                    let hit = wildcard_match(held, pattern);
                    return match self.op {
                        BinaryPredicate::Eq => hit,
                        BinaryPredicate::Ne => !hit,
                        _ => false,
                    };
                }
                return matches!(self.op, BinaryPredicate::Ne);
            }
        }
        // A comparison for greater or less reads only its own kind: measured,
        // `COUNTIF(.., ">1/1/2024")` counts no text cell, where the sheet's
        // own `>` would put every text above every number.
        let kind = |value: &Value| match value {
            Value::Number(_) => 0,
            Value::Text(_) => 1,
            Value::Logical(_) => 2,
            _ => 3,
        };
        if !matches!(self.op, BinaryPredicate::Eq | BinaryPredicate::Ne)
            && kind(v) != kind(&self.operand)
        {
            return false;
        }
        match compare(v, &self.operand) {
            Ok(ord) => match self.op {
                BinaryPredicate::Eq => ord == Ordering::Equal,
                BinaryPredicate::Ne => ord != Ordering::Equal,
                BinaryPredicate::Lt => ord == Ordering::Less,
                BinaryPredicate::Le => ord != Ordering::Greater,
                BinaryPredicate::Gt => ord == Ordering::Greater,
                BinaryPredicate::Ge => ord != Ordering::Less,
            },
            Err(_) => false,
        }
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    /// Every quoted name a match arm of this file answers to is in
    /// KNOWN_FUNCTIONS, which is kept sorted for its binary search.
    #[test]
    fn every_function_the_library_answers_is_known() {
        let mut sorted = KNOWN_FUNCTIONS.to_vec();
        sorted.sort_unstable();
        assert_eq!(sorted, KNOWN_FUNCTIONS, "KNOWN_FUNCTIONS must stay sorted");
        let source = include_str!("functions.rs");
        for line in source.lines() {
            let trimmed = line.trim_start().trim_start_matches('|').trim_start();
            if !trimmed.starts_with('"') || !(trimmed.trim_end().ends_with("=>") || trimmed.trim_end().ends_with('|') || trimmed.contains("\" =>") || trimmed.contains("\" |")) {
                continue;
            }
            for piece in trimmed.split('"').skip(1).step_by(2) {
                let is_name = piece.chars().next().is_some_and(|ch| ch.is_ascii_uppercase())
                    && piece.chars().all(|ch| ch.is_ascii_uppercase() || ch.is_ascii_digit() || ch == '.' || ch == '_');
                if is_name {
                    assert!(is_known_function(piece), "{piece} is answered but not in KNOWN_FUNCTIONS");
                }
            }
        }
    }

    fn v(n: f64) -> Arg {
        Arg::Value(Value::Number(n))
    }
    fn t(s: &str) -> Arg {
        Arg::Value(Value::text(s))
    }
    fn range(values: &[Value], width: usize) -> Arg {
        Arg::Range(RangeData {
            width,
            height: values.len() / width,
            cells: values.to_vec(),
        })
    }

    /// The functions added 2026-09-05 (flows10), measured against Excel.
    #[test]
    fn the_later_worksheet_functions_agree_with_excel() {
        assert_eq!(call("SUMSQ", &[v(3.0), v(4.0)]), Value::Number(25.0));
        assert_eq!(call("GCD", &[v(12.0), v(18.0)]), Value::Number(6.0));
        assert_eq!(call("LCM", &[v(4.0), v(6.0)]), Value::Number(12.0));
        assert_eq!(call("EVEN", &[v(3.0)]), Value::Number(4.0));
        assert_eq!(call("ODD", &[v(4.0)]), Value::Number(5.0));
        assert_eq!(call("EVEN", &[v(-1.0)]), Value::Number(-2.0));
        assert_eq!(call("ROMAN", &[v(49.0)]), Value::text("XLIX"));
        assert_eq!(call("ROMAN", &[v(2024.0)]), Value::text("MMXXIV"));
        assert_eq!(call("ARABIC", &[t("XLIX")]), Value::Number(49.0));
        assert_eq!(call("CLEAN", &[t("a	b")]), Value::text("ab"));
        assert_eq!(call("FIXED", &[v(1234.567), v(2.0)]), Value::text("1,234.57"));
        assert_eq!(call("FIXED", &[v(1234.5), v(0.0), Arg::Value(Value::Logical(true))]), Value::text("1235"));
        assert_eq!(call("DOLLAR", &[v(1234.5), v(0.0)]), Value::text("$1,235"));
        assert_eq!(call("SUBSTITUTE", &[t("aaa"), t("a"), t("b"), v(2.0)]), Value::text("aba"));
        // LOOKUP's vector form: the largest not over the needle.
        let col = range(&[n(10.0), n(20.0), n(30.0), n(40.0)], 1);
        assert_eq!(call("LOOKUP", &[v(25.0), col.clone()]), Value::Number(20.0));
        // NETWORKDAYS counts the weekdays, both ends in: Mon 1 Jan to Sun 7.
        let mon = Arg::Value(Value::Number(datetime::serial_from_date(2024, 1, 1).unwrap() as f64));
        let sun = Arg::Value(Value::Number(datetime::serial_from_date(2024, 1, 7).unwrap() as f64));
        assert_eq!(call("NETWORKDAYS", &[mon, sun]), Value::Number(5.0));
        // DAVERAGE: the field's average over rows matching the criteria.
        let db = range(
            &[
                Value::text("k"), Value::text("v"),
                Value::text("x"), n(10.0),
                Value::text("y"), n(20.0),
                Value::text("x"), n(30.0),
            ],
            2,
        );
        let crit = range(&[Value::text("k"), Value::text("x")], 1);
        assert_eq!(call("DAVERAGE", &[db, v(2.0), crit]), Value::Number(20.0));
    }

    /// The functions added 2026-09-05 (flows12), measured against Excel.
    #[test]
    fn the_flows12_worksheet_functions_agree_with_excel() {
        // MROUND rounds to the nearest multiple, halves away from zero, and
        // asks the number and multiple to share a sign.
        assert_eq!(call("MROUND", &[v(17.0), v(5.0)]), Value::Number(15.0));
        assert_eq!(call("MROUND", &[v(-17.0), v(-5.0)]), Value::Number(-15.0));
        assert_eq!(call("MROUND", &[v(2.5), v(1.0)]), Value::Number(3.0));
        assert_eq!(call("MROUND", &[v(3.0), v(0.0)]), Value::Number(0.0));
        assert_eq!(call("MROUND", &[v(5.0), v(-2.0)]), Value::Error(ExcelError::Num));
        // QUOTIENT truncates toward zero; a zero divisor is #DIV/0!.
        assert_eq!(call("QUOTIENT", &[v(17.0), v(5.0)]), Value::Number(3.0));
        assert_eq!(call("QUOTIENT", &[v(-17.0), v(5.0)]), Value::Number(-3.0));
        assert_eq!(call("QUOTIENT", &[v(5.0), v(0.0)]), Value::Error(ExcelError::DivZero));
        // XOR is true for an odd count of truths.
        assert_eq!(call("XOR", &[l(true), l(false), l(true)]), Value::Logical(false));
        assert_eq!(call("XOR", &[l(true), l(true), l(true)]), Value::Logical(true));
        // IFS takes the first true condition; nothing true is #N/A.
        assert_eq!(
            call("IFS", &[l(false), t("a"), l(true), t("b")]),
            Value::text("b")
        );
        assert_eq!(call("IFS", &[l(false), t("a")]), Value::Error(ExcelError::NA));
        // SWITCH takes the first value equal to the subject, then a lone final
        // default, else #N/A.
        assert_eq!(
            call("SWITCH", &[v(3.0), v(1.0), t("a"), v(3.0), t("c"), t("def")]),
            Value::text("c")
        );
        assert_eq!(
            call("SWITCH", &[v(9.0), v(1.0), t("a"), t("def")]),
            Value::text("def")
        );
        assert_eq!(
            call("SWITCH", &[v(9.0), v(1.0), t("a"), v(3.0), t("c")]),
            Value::Error(ExcelError::NA)
        );
        // YEARFRAC on all five bases, 1 Jan 2024 -> 1 Jul 2024 (a leap year).
        let jan1 = Arg::Value(n(datetime::serial_from_date(2024, 1, 1).unwrap() as f64));
        let jul1 = Arg::Value(n(datetime::serial_from_date(2024, 7, 1).unwrap() as f64));
        let round4 = |a: Arg, b: Arg, basis: f64| match call("YEARFRAC", &[a, b, v(basis)]) {
            Value::Number(x) => (x * 1_000_000.0).round() / 1_000_000.0,
            other => panic!("YEARFRAC gave {other:?}"),
        };
        assert_eq!(round4(jan1.clone(), jul1.clone(), 0.0), 0.5);
        assert_eq!(round4(jan1.clone(), jul1.clone(), 1.0), 0.497268);
        assert_eq!(round4(jan1.clone(), jul1.clone(), 2.0), 0.505556);
        assert_eq!(round4(jan1.clone(), jul1.clone(), 3.0), 0.49863);
        assert_eq!(round4(jan1, jul1, 4.0), 0.5);
        // The 30/360 US Feb edge: 29 Feb 2024 -> 31 Mar 2024 is 31 days, and
        // its D2 test reads the day before February pulls D1 to 30.
        let feb29 = Arg::Value(n(datetime::serial_from_date(2024, 2, 29).unwrap() as f64));
        let mar31 = Arg::Value(n(datetime::serial_from_date(2024, 3, 31).unwrap() as f64));
        assert_eq!(round4(feb29, mar31, 0.0), 0.086111);
        // Symmetric in its dates.
        let a = Arg::Value(n(datetime::serial_from_date(2024, 1, 1).unwrap() as f64));
        let b = Arg::Value(n(datetime::serial_from_date(2024, 7, 1).unwrap() as f64));
        assert_eq!(call("YEARFRAC", &[b, a]), Value::Number(0.5));
    }

    fn l(value: bool) -> Arg {
        Arg::Value(Value::Logical(value))
    }

    /// The normal-distribution family added 2026-09-05, measured against Excel
    /// to eight decimals (the CDF/inverse kernels must hold in the tails too).
    #[test]
    fn the_normal_distribution_agrees_with_excel() {
        let near = |value: Value, want: f64| match value {
            Value::Number(x) => assert!((x - want).abs() < 5e-8, "got {x}, want {want}"),
            other => panic!("expected a number, got {other:?}"),
        };
        near(call("NORM.S.DIST", &[v(1.0), l(true)]), 0.84134475);
        near(call("NORM.S.DIST", &[v(1.0), l(false)]), 0.24197072);
        near(call("NORM.S.DIST", &[v(-2.5), l(true)]), 0.00620967);
        near(call("NORM.DIST", &[v(8.0), v(10.0), v(2.0), l(true)]), 0.15865525);
        near(call("NORM.DIST", &[v(8.0), v(10.0), v(2.0), l(false)]), 0.12098536);
        near(call("NORM.S.INV", &[v(0.975)]), 1.95996398);
        near(call("NORM.S.INV", &[v(0.001)]), -3.09023231);
        near(call("NORM.S.INV", &[v(0.5)]), 0.0);
        near(call("NORM.INV", &[v(0.95), v(100.0), v(15.0)]), 124.6728044);
        near(call("STANDARDIZE", &[v(85.0), v(100.0), v(15.0)]), -1.0);
        near(call("GAUSS", &[v(1.0)]), 0.34134475);
        near(call("PHI", &[v(1.0)]), 0.24197072);
        near(call("NORMSDIST", &[v(1.0)]), 0.84134475);
        near(call("NORMINV", &[v(0.95), v(100.0), v(15.0)]), 124.6728044);
        near(call("NORMSINV", &[v(0.975)]), 1.95996398);
        near(call("CONFIDENCE.NORM", &[v(0.05), v(2.5), v(50.0)]), 0.69295191);
        // Bad standard deviation or probability is #NUM!.
        assert_eq!(
            call("NORM.DIST", &[v(1.0), v(0.0), v(-1.0), l(true)]),
            Value::Error(ExcelError::Num)
        );
        assert_eq!(call("NORM.S.INV", &[v(0.0)]), Value::Error(ExcelError::Num));
    }

    /// The conditional and A-suffix aggregates added 2026-09-05 (flows15),
    /// measured against Excel.
    #[test]
    fn the_flows15_aggregates_agree_with_excel() {
        let a = range(&[n(3.0), n(1.0), n(4.0), n(1.0), n(5.0), n(9.0), n(2.0), n(6.0)], 1);
        let b = range(&[n(10.0), n(20.0), n(10.0), n(20.0), n(10.0), n(20.0), n(10.0), n(20.0)], 1);
        assert_eq!(call("MAXIFS", &[a.clone(), b.clone(), v(10.0)]), Value::Number(5.0));
        assert_eq!(call("MINIFS", &[a.clone(), b.clone(), v(20.0)]), Value::Number(1.0));
        // Two criteria pairs.
        assert_eq!(
            call("MAXIFS", &[a.clone(), b, v(10.0), a, t(">1")]),
            Value::Number(5.0)
        );
        // MAXA/MINA read text as 0 -- so the max of {-5, "abc", -3} is 0.
        let mixed = range(&[n(-5.0), Value::text("abc"), n(-3.0)], 1);
        assert_eq!(call("MAXA", &[mixed.clone()]), Value::Number(0.0));
        assert_eq!(call("MINA", &[mixed]), Value::Number(-5.0));
        // A logical in a range counts as 1.
        assert_eq!(call("MAXA", &[range(&[Value::Logical(true), n(0.5)], 1)]), Value::Number(1.0));
        // STDEVA/VARA, sample spread.
        let four = range(&[n(3.0), n(1.0), n(4.0), n(1.0)], 1);
        assert_eq!(call("STDEVA", &[four.clone()]), Value::Number(1.5));
        assert_eq!(call("VARA", &[four]), Value::Number(2.25));
        // SQRTPI, COMBINA, UNICHAR.
        match call("SQRTPI", &[v(2.0)]) {
            Value::Number(x) => assert!((x - 2.5066282746).abs() < 1e-6, "SQRTPI {x}"),
            other => panic!("SQRTPI {other:?}"),
        }
        assert_eq!(call("COMBINA", &[v(4.0), v(2.0)]), Value::Number(10.0));
        assert_eq!(call("COMBINA", &[v(3.0), v(0.0)]), Value::Number(1.0));
        assert_eq!(call("UNICHAR", &[v(65.0)]), Value::text("A"));
        assert_eq!(call("UNICHAR", &[v(0.0)]), Value::Error(ExcelError::Value));
    }

    /// The lookup, reference and statistics functions added 2026-09-05
    /// (flows14), measured against Excel.
    #[test]
    fn the_flows14_lookup_and_stat_functions_agree_with_excel() {
        // XMATCH over a sorted list, exact and the two approximate modes.
        let sorted = range(&[n(10.0), n(20.0), n(30.0), n(40.0), n(50.0)], 1);
        assert_eq!(call("XMATCH", &[v(30.0), sorted.clone()]), Value::Number(3.0));
        assert_eq!(call("XMATCH", &[v(35.0), sorted.clone(), v(-1.0)]), Value::Number(3.0));
        assert_eq!(call("XMATCH", &[v(35.0), sorted.clone(), v(1.0)]), Value::Number(4.0));
        assert_eq!(call("XMATCH", &[v(20.0), sorted.clone(), v(0.0), v(-1.0)]), Value::Number(2.0));
        assert_eq!(call("XMATCH", &[v(35.0), sorted]), Value::Error(ExcelError::NA));
        // ADDRESS in its four absolute forms, R1C1, with a sheet, and a wide column.
        assert_eq!(call("ADDRESS", &[v(3.0), v(2.0)]), Value::text("$B$3"));
        assert_eq!(call("ADDRESS", &[v(3.0), v(2.0), v(2.0)]), Value::text("B$3"));
        assert_eq!(call("ADDRESS", &[v(3.0), v(2.0), v(3.0)]), Value::text("$B3"));
        assert_eq!(call("ADDRESS", &[v(3.0), v(2.0), v(4.0)]), Value::text("B3"));
        assert_eq!(call("ADDRESS", &[v(3.0), v(2.0), v(1.0), l(false)]), Value::text("R3C2"));
        assert_eq!(
            call("ADDRESS", &[v(3.0), v(2.0), v(1.0), l(true), t("Sheet1")]),
            Value::text("Sheet1!$B$3")
        );
        assert_eq!(call("ADDRESS", &[v(1.0), v(27.0)]), Value::text("$AA$1"));
        // Paired-array sums.
        let xs = range(&[n(3.0), n(1.0), n(4.0)], 1);
        let ys = range(&[n(100.0), n(200.0), n(300.0)], 1);
        assert_eq!(call("SUMXMY2", &[xs.clone(), ys.clone()]), Value::Number(136626.0));
        assert_eq!(call("SUMX2MY2", &[xs.clone(), ys.clone()]), Value::Number(-139974.0));
        assert_eq!(call("SUMX2PY2", &[xs, ys]), Value::Number(140026.0));
        // RANK.AVG averages tied places.
        let data = range(&[n(3.0), n(1.0), n(4.0), n(1.0), n(5.0), n(9.0), n(2.0), n(6.0)], 1);
        assert_eq!(call("RANK.AVG", &[v(1.0), data.clone()]), Value::Number(7.5));
        // Deviations.
        let four = range(&[n(3.0), n(1.0), n(4.0), n(1.0)], 1);
        assert_eq!(call("AVEDEV", &[four.clone()]), Value::Number(1.25));
        assert_eq!(call("DEVSQ", &[four]), Value::Number(6.75));
        assert_eq!(call("MULTINOMIAL", &[v(2.0), v(3.0), v(4.0)]), Value::Number(1260.0));
        // PERCENTRANK truncated to three significant digits, and TRIMMEAN.
        assert_eq!(call("PERCENTRANK", &[data.clone(), v(4.0)]), Value::Number(0.571));
        assert_eq!(call("TRIMMEAN", &[data, v(0.25)]), Value::Number(3.5));
    }

    /// The financial functions added 2026-09-05 (flows14), measured against
    /// Excel. The iterative ones (IRR, RATE) are checked to a tolerance.
    #[test]
    fn the_financial_functions_agree_with_excel() {
        let round2 = |value: Value| match value {
            Value::Number(x) => (x * 100.0).round() / 100.0,
            other => panic!("expected a number, got {other:?}"),
        };
        assert_eq!(round2(call("PMT", &[v(0.05 / 12.0), v(60.0), v(-10000.0)])), 188.71);
        assert_eq!(round2(call("FV", &[v(0.05 / 12.0), v(60.0), v(-100.0)])), 6800.61);
        assert_eq!(round2(call("PV", &[v(0.05 / 12.0), v(60.0), v(-100.0)])), 5299.07);
        assert_eq!(round2(call("IPMT", &[v(0.05 / 12.0), v(1.0), v(60.0), v(-10000.0)])), 41.67);
        assert_eq!(round2(call("PPMT", &[v(0.05 / 12.0), v(1.0), v(60.0), v(-10000.0)])), 147.05);
        assert_eq!(call("SLN", &[v(10000.0), v(1000.0), v(5.0)]), Value::Number(1800.0));
        assert_eq!(call("SYD", &[v(10000.0), v(1000.0), v(5.0), v(1.0)]), Value::Number(3000.0));
        assert_eq!(call("DDB", &[v(10000.0), v(1000.0), v(5.0), v(1.0)]), Value::Number(4000.0));
        assert_eq!(round2(call("DB", &[v(10000.0), v(1000.0), v(5.0), v(1.0)])), 3690.0);
        assert_eq!(round2(call("DB", &[v(10000.0), v(1000.0), v(5.0), v(2.0)])), 2328.39);
        assert_eq!(round2(call("DB", &[v(10000.0), v(1000.0), v(5.0), v(6.0), v(6.0)])), 238.53);
        // NPV takes several values and needs no sign mix; IRR does.
        let flows = range(&[n(100.0), n(200.0), n(300.0)], 1);
        assert_eq!(round2(call("NPV", &[v(0.1), flows])), 481.59);
        let cash = range(&[n(-1000.0), n(300.0), n(400.0), n(500.0)], 1);
        match call("IRR", &[cash]) {
            Value::Number(x) => assert!((x - 0.0891).abs() < 5e-4, "IRR {x}"),
            other => panic!("IRR {other:?}"),
        }
        match call("RATE", &[v(60.0), v(-100.0), v(5000.0)]) {
            Value::Number(x) => assert!((x - 0.006183).abs() < 1e-5, "RATE {x}"),
            other => panic!("RATE {other:?}"),
        }
        match call("NPER", &[v(0.05 / 12.0), v(-100.0), v(5000.0)]) {
            Value::Number(x) => assert!((x - 56.1843).abs() < 1e-3, "NPER {x}"),
            other => panic!("NPER {other:?}"),
        }
        // A zero-rate annuity and a bad payment type.
        assert_eq!(call("PMT", &[v(0.0), v(10.0), v(-1000.0)]), Value::Number(100.0));
        assert_eq!(
            call("PMT", &[v(0.05), v(10.0), v(-1000.0), v(0.0), v(2.0)]),
            Value::Error(ExcelError::Num)
        );
    }

    /// The math and information functions added 2026-09-05 (flows13), measured
    /// against Excel.
    #[test]
    fn the_flows13_math_and_info_functions_agree_with_excel() {
        assert_eq!(call("TRUNC", &[v(3.78)]), Value::Number(3.0));
        assert_eq!(call("TRUNC", &[v(-3.78), v(1.0)]), Value::Number(-3.7));
        assert_eq!(call("SIGN", &[v(-5.0)]), Value::Number(-1.0));
        assert_eq!(call("SIGN", &[v(0.0)]), Value::Number(0.0));
        assert_eq!(call("SIGN", &[v(7.0)]), Value::Number(1.0));
        assert_eq!(call("COMBIN", &[v(10.0), v(3.0)]), Value::Number(120.0));
        assert_eq!(call("PERMUT", &[v(10.0), v(3.0)]), Value::Number(720.0));
        assert_eq!(call("FACT", &[v(6.0)]), Value::Number(720.0));
        assert_eq!(call("FACTDOUBLE", &[v(7.0)]), Value::Number(105.0));
        assert_eq!(call("FACTDOUBLE", &[v(0.0)]), Value::Number(1.0));
        assert_eq!(call("FACT", &[v(-1.0)]), Value::Error(ExcelError::Num));
        match call("DEGREES", &[v(std::f64::consts::PI)]) {
            Value::Number(x) => assert!((x - 180.0).abs() < 1e-9),
            other => panic!("DEGREES gave {other:?}"),
        }
        match call("RADIANS", &[v(180.0)]) {
            Value::Number(x) => assert!((x - std::f64::consts::PI).abs() < 1e-9),
            other => panic!("RADIANS gave {other:?}"),
        }
        assert_eq!(call("ISODD", &[v(7.0)]), Value::Logical(true));
        assert_eq!(call("ISEVEN", &[v(7.0)]), Value::Logical(false));
        assert_eq!(call("N", &[v(42.0)]), Value::Number(42.0));
        assert_eq!(call("N", &[t("x")]), Value::Number(0.0));
        assert_eq!(call("N", &[l(true)]), Value::Number(1.0));
        assert_eq!(call("TYPE", &[v(5.0)]), Value::Number(1.0));
        assert_eq!(call("TYPE", &[t("a")]), Value::Number(2.0));
        assert_eq!(call("TYPE", &[l(true)]), Value::Number(4.0));
        assert_eq!(
            call("TYPE", &[Arg::Value(Value::Error(ExcelError::NA))]),
            Value::Number(16.0)
        );
        assert_eq!(
            call("TYPE", &[range(&[n(1.0), n(2.0), n(3.0)], 1)]),
            Value::Number(64.0)
        );
        assert_eq!(
            call("ERROR.TYPE", &[Arg::Value(Value::Error(ExcelError::DivZero))]),
            Value::Number(2.0)
        );
        assert_eq!(
            call("ERROR.TYPE", &[Arg::Value(Value::Error(ExcelError::NA))]),
            Value::Number(7.0)
        );
        assert_eq!(call("ERROR.TYPE", &[v(5.0)]), Value::Error(ExcelError::NA));
        assert_eq!(call("DECIMAL", &[t("FF"), v(16.0)]), Value::Number(255.0));
        assert_eq!(call("BASE", &[v(255.0), v(16.0)]), Value::text("FF"));
        assert_eq!(call("BASE", &[v(5.0), v(2.0), v(8.0)]), Value::text("00000101"));
        assert_eq!(call("BITAND", &[v(12.0), v(10.0)]), Value::Number(8.0));
        assert_eq!(call("BITOR", &[v(12.0), v(10.0)]), Value::Number(14.0));
        let data = range(&[n(3.0), n(1.0), n(4.0), n(1.0), n(5.0), n(9.0), n(2.0), n(6.0)], 1);
        assert_eq!(call("AVERAGEA", &[data]), Value::Number(3.875));
        match call("GEOMEAN", &[v(3.0), v(1.0), v(4.0), v(1.0)]) {
            Value::Number(x) => assert!((x - 1.861209).abs() < 1e-4),
            other => panic!("GEOMEAN gave {other:?}"),
        }
        let jan1 = Arg::Value(n(datetime::serial_from_date(2024, 1, 1).unwrap() as f64));
        assert_eq!(call("ISOWEEKNUM", &[jan1]), Value::Number(1.0));
    }

    /// The date and text functions added 2026-09-05 (flows13), measured
    /// against Excel.
    #[test]
    fn the_flows13_date_and_text_functions_agree_with_excel() {
        let date = |y, m, d| Arg::Value(n(datetime::serial_from_date(y, m, d).unwrap() as f64));
        // DAYS360, US and European. The US form differs from YEARFRAC basis 0.
        assert_eq!(call("DAYS360", &[date(2024, 2, 29), date(2024, 3, 31)]), Value::Number(30.0));
        assert_eq!(
            call("DAYS360", &[date(2024, 2, 29), date(2024, 3, 31), l(true)]),
            Value::Number(31.0)
        );
        assert_eq!(call("DAYS360", &[date(2024, 1, 31), date(2024, 3, 31)]), Value::Number(60.0));
        assert_eq!(call("DAYS360", &[date(2024, 4, 15), date(2024, 5, 31)]), Value::Number(46.0));
        assert_eq!(call("DAYS360", &[date(2024, 1, 15), date(2024, 2, 29)]), Value::Number(44.0));
        // DATEVALUE reads a date out of text and drops any time.
        let serial = datetime::serial_from_date(2024, 3, 15).unwrap() as f64;
        assert_eq!(call("DATEVALUE", &[t("2024-03-15")]), Value::Number(serial));
        assert_eq!(call("DATEVALUE", &[t("3/15/2024")]), Value::Number(serial));
        assert_eq!(call("DATEVALUE", &[t("15-Mar-2024")]), Value::Number(serial));
        assert_eq!(call("DATEVALUE", &[t("not a date")]), Value::Error(ExcelError::Value));
        // Logarithms, the circle, bases and paired data -- every answer Excel's.
        let t2 = |a: &str| Arg::Value(Value::text(a));
        for (name, args, want) in [
            ("LN", vec![v(10.0)], 10f64.ln()),
            ("LOG10", vec![v(1000.0)], 3.0),
            ("LOG", vec![v(8.0), v(2.0)], 3.0),
            ("LOG", vec![v(100.0)], 2.0),
            ("ATAN2", vec![v(1.0), v(1.0)], std::f64::consts::FRAC_PI_4),
            ("HEX2DEC", vec![t2("FFFFFFFFFF")], -1.0),
            ("BIN2DEC", vec![t2("1010")], 10.0),
        ] {
            assert_eq!(call(name, &args), Value::Number(want), "{name}");
        }
        for (name, args) in [
            ("LN", vec![v(0.0)]),
            ("ASIN", vec![v(2.0)]),
            ("ACOSH", vec![v(0.5)]),
            ("ATANH", vec![v(1.0)]),
            ("DEC2HEX", vec![v(255.0), v(1.0)]),
            ("DEC2BIN", vec![v(512.0)]),
            ("HEX2BIN", vec![t2("200")]),
        ] {
            assert_eq!(call(name, &args), Value::Error(ExcelError::Num), "{name}");
        }
        assert_eq!(call("LOG", &[v(10.0), v(1.0)]), Value::Error(ExcelError::DivZero));
        assert_eq!(call("ATAN2", &[v(0.0), v(0.0)]), Value::Error(ExcelError::DivZero));
        for (name, args, want) in [
            ("DEC2HEX", vec![v(-1.0)], "FFFFFFFFFF"),
            ("DEC2BIN", vec![v(-1.0)], "1111111111"),
            ("DEC2HEX", vec![v(255.0), v(4.0)], "00FF"),
            ("DEC2OCT", vec![v(-8.0)], "7777777770"),
            ("BIN2HEX", vec![t2("1111111111")], "FFFFFFFFFF"),
            ("HEX2BIN", vec![t2("1FF")], "111111111"),
        ] {
            assert_eq!(call(name, &args), Value::text(want), "{name}");
        }
        let a = Arg::Range(RangeData {
            width: 1,
            height: 6,
            cells: [3.0, 7.0, 7.0, 1.0, 9.0, 4.0].map(Value::Number).to_vec(),
        });
        let b = Arg::Range(RangeData {
            width: 1,
            height: 6,
            cells: vec![v(2.0).scalar(), v(5.0).scalar(), Value::text("x"), v(1.0).scalar(), v(8.0).scalar(), v(3.0).scalar()],
        });
        let close = |got: Value, want: f64| match got {
            Value::Number(n) => assert!((n - want).abs() < 1e-12, "{n} vs {want}"),
            other => panic!("{other:?}"),
        };
        close(call("INTERCEPT", &[a.clone(), b.clone()]), 0.506493506493507);
        close(call("RSQ", &[a.clone(), b.clone()]), 0.963712757830405);
        close(call("STEYX", &[a.clone(), b.clone()]), 0.702500173314211);
        close(call("COVAR", &[a.clone(), b.clone()]), 6.96);
        close(call("COVARIANCE.S", &[a.clone(), b]), 8.7);
        close(call("SKEW", &[a.clone()]), -0.172562731898406);
        close(call("KURT", &[a.clone()]), -1.34119207860588);
        close(call("HARMEAN", &[a]), 3.03006012024048);
        // The reshaping functions, every answer Excel's (A1:C4 = 1..12).
        let grid = Arg::Range(RangeData {
            width: 3,
            height: 4,
            cells: (1..=12).map(|n| Value::Number(n as f64)).collect(),
        });
        let flat = |arg: Arg| -> (usize, usize, Vec<Value>) {
            let Arg::Range(block) = arg else { panic!("{arg:?}") };
            (block.height, block.width, block.cells)
        };
        let nums = |xs: &[f64]| xs.iter().map(|x| Value::Number(*x)).collect::<Vec<_>>();
        assert_eq!(flat(call_arg("SEQUENCE", &[v(2.0), v(2.0), v(10.0), v(5.0)])), (2, 2, nums(&[10.0, 15.0, 20.0, 25.0])));
        assert_eq!(flat(call_arg("TAKE", &[grid.clone(), v(2.0), v(-2.0)])), (2, 2, nums(&[2.0, 3.0, 5.0, 6.0])));
        assert_eq!(flat(call_arg("DROP", &[grid.clone(), v(-2.0), v(1.0)])), (2, 2, nums(&[2.0, 3.0, 5.0, 6.0])));
        assert_eq!(flat(call_arg("CHOOSECOLS", &[grid.clone(), v(3.0), v(1.0)])).2[..2], nums(&[3.0, 1.0])[..]);
        assert_eq!(flat(call_arg("TOCOL", &[grid.clone(), v(0.0), l(true)])).2[..3], nums(&[1.0, 4.0, 7.0])[..]);
        let wrapped = flat(call_arg("WRAPCOLS", &[
            Arg::Range(RangeData { width: 1, height: 6, cells: nums(&[1.0, 2.0, 3.0, 4.0, 5.0, 6.0]) }),
            v(4.0),
            Arg::Value(Value::text("-")),
        ]));
        assert_eq!((wrapped.0, wrapped.1), (4, 2));
        assert_eq!(wrapped.2[5], Value::text("-"));
        let split = flat(call_arg("TEXTSPLIT", &[
            Arg::Value(Value::text("a,b;c")),
            Arg::Value(Value::text(",")),
            Arg::Value(Value::text(";")),
        ]));
        assert_eq!(split, (2, 2, vec![Value::text("a"), Value::text("b"), Value::text("c"), Value::Error(ExcelError::NA)]));
        // Shift_JIS byte functions and the weekend patterns, every answer Excel's.
        let tokyo = || Arg::Value(Value::text("東京abc"));
        assert_eq!(call("LENB", &[tokyo()]), Value::Number(7.0));
        assert_eq!(call("LENB", &[Arg::Value(Value::text("ｱｲｳ"))]), Value::Number(3.0));
        assert_eq!(call("LEFTB", &[tokyo(), v(3.0)]), Value::text("東 "));
        assert_eq!(call("RIGHTB", &[tokyo(), v(4.0)]), Value::text(" abc"));
        assert_eq!(call("MIDB", &[tokyo(), v(2.0), v(3.0)]), Value::text(" 京"));
        assert_eq!(call("FINDB", &[Arg::Value(Value::text("a")), tokyo()]), Value::Number(5.0));
        assert_eq!(call("SEARCHB", &[Arg::Value(Value::text("B")), tokyo()]), Value::Number(6.0));
        assert_eq!(
            call("REPLACEB", &[tokyo(), v(1.0), v(2.0), Arg::Value(Value::text("x"))]),
            Value::text("x京abc")
        );
        let (jan1, jan31) = (v(45292.0), v(45322.0));
        for (weekend, want) in [(1.0, 23.0), (11.0, 27.0)] {
            assert_eq!(
                call("NETWORKDAYS.INTL", &[jan1.clone(), jan31.clone(), v(weekend)]),
                Value::Number(want)
            );
        }
        assert_eq!(
            call("NETWORKDAYS.INTL", &[jan1.clone(), jan31.clone(), Arg::Value(Value::text("0000011"))]),
            Value::Number(23.0)
        );
        assert_eq!(
            call("NETWORKDAYS.INTL", &[jan1, jan31, Arg::Value(Value::text("1111111"))]),
            Value::Number(0.0)
        );
        assert_eq!(call("WORKDAY.INTL", &[v(45296.0), v(1.0), v(7.0)]), Value::Number(45298.0));
        assert_eq!(call("BITXOR", &[v(12.0), v(10.0)]), Value::Number(6.0));
        // TIMEVALUE, every answer Excel's.
        for (text, want) in [
            ("1:00", 1.0 / 24.0),
            ("1:00 PM", 13.0 / 24.0),
            ("2024/1/1 6:00", 0.25),
            // What is left of 25 hours once the day is taken off, rounding
            // and all: measured, (TIMEVALUE("25:00")*24-1)*1E15 is 1.776.
            ("25:00", 25.0 / 24.0 - 1.0),
            ("12:00 AM", 0.0),
            ("2024/1/1", 0.0),
            ("1:60", 2.0 / 24.0),
            ("10 PM", 22.0 / 24.0),
            ("0 AM", 0.0),
        ] {
            assert_eq!(call("TIMEVALUE", &[t(text)]), Value::Number(want), "{text}");
        }
        for text in ["abc", "10PM", "13 PM", "1:60 PM", "10000:00"] {
            assert_eq!(call("TIMEVALUE", &[t(text)]), Value::Error(ExcelError::Value), "{text}");
        }
        assert_eq!(call("TIMEVALUE", &[v(0.5)]), Value::Error(ExcelError::Value));
        // TEXTBEFORE / TEXTAFTER on the nth delimiter, from either end.
        assert_eq!(call("TEXTBEFORE", &[t("a-b-c"), t("-")]), Value::text("a"));
        assert_eq!(call("TEXTAFTER", &[t("a-b-c"), t("-"), v(2.0)]), Value::text("c"));
        assert_eq!(call("TEXTBEFORE", &[t("a-b-c"), t("-"), v(-1.0)]), Value::text("a-b"));
        assert_eq!(call("TEXTAFTER", &[t("a-b-c"), t("-"), v(-1.0)]), Value::text("c"));
        assert_eq!(
            call("TEXTAFTER", &[t("a-b-c"), t("-"), v(5.0)]),
            Value::Error(ExcelError::NA)
        );
        // NUMBERVALUE with its own separators, a space group, and a percent.
        assert_eq!(
            call("NUMBERVALUE", &[t("1,234.5"), t("."), t(",")]),
            Value::Number(1234.5)
        );
        assert_eq!(
            call("NUMBERVALUE", &[t("1 234,5"), t(","), t(" ")]),
            Value::Number(1234.5)
        );
        assert_eq!(call("NUMBERVALUE", &[t("50%")]), Value::Number(0.5));
        assert_eq!(call("NUMBERVALUE", &[t("")]), Value::Number(0.0));
    }

    /// ASC, full-width to half-width, measured against Excel by code point.
    #[test]
    fn asc_narrows_full_width_letters_and_kana() {
        // Full-width ABC12 -> ABC12.
        assert_eq!(
            call("ASC", &[t("\u{FF21}\u{FF22}\u{FF23}\u{FF11}\u{FF12}")]),
            Value::text("ABC12")
        );
        // Plain katakana ア イ ウ -> ｱ ｲ ｳ.
        assert_eq!(
            call("ASC", &[t("\u{30A2}\u{30A4}\u{30A6}")]),
            Value::text("\u{FF71}\u{FF72}\u{FF73}")
        );
        // Voiced and semi-voiced split into base + mark: ガ パ -> ｶﾞ ﾊﾟ.
        assert_eq!(
            call("ASC", &[t("\u{30AC}\u{30D1}")]),
            Value::text("\u{FF76}\u{FF9E}\u{FF8A}\u{FF9F}")
        );
        // ヴ -> ｳﾞ.
        assert_eq!(call("ASC", &[t("\u{30F4}")]), Value::text("\u{FF73}\u{FF9E}"));
        // Kanji and half-width digits are left alone; the full-width A narrows.
        assert_eq!(call("ASC", &[t("\u{FF21}\u{611B}1")]), Value::text("A\u{611B}1"));
        // The full-width space, and the katakana punctuation.
        assert_eq!(call("ASC", &[t("\u{3000}")]), Value::text(" "));
        assert_eq!(
            call("ASC", &[t("\u{30FC}\u{30FB}\u{3001}\u{3002}\u{300C}\u{300D}")]),
            Value::text("\u{FF70}\u{FF65}\u{FF64}\u{FF61}\u{FF62}\u{FF63}")
        );
        // A standalone combining voiced mark stays as it is, as Excel leaves it.
        assert_eq!(call("ASC", &[t("\u{3099}")]), Value::text("\u{3099}"));
    }

    fn n(value: f64) -> Value {
        Value::Number(value)
    }

    #[test]
    fn the_prefix_a_file_writes_is_not_part_of_the_name() {
        // A function newer than the format's own version is stored as
        // `_xlfn.NAME`, and Excel shows it without. The parser upper-cases
        // every name it reads, so the prefix arrives as `_XLFN.` however the
        // file spelled it — and stripping only the lower-case form matched
        // nothing at all, silently, which is how IFNA came back `#NAME?`
        // despite being implemented all along.
        assert_eq!(plain("_xlfn.IFNA"), "IFNA");
        assert_eq!(plain("_XLFN.IFNA"), "IFNA");
        assert_eq!(plain("_xlfn._xlws.SORT"), "SORT");
        assert_eq!(plain("_XLFN._XLWS.FILTER"), "FILTER");
        assert_eq!(plain("SUM"), "SUM");
        // A name that merely starts with an underscore is left alone.
        assert_eq!(plain("_MYNAME"), "_MYNAME");
        assert_eq!(call("_XLFN.IFNA", &[v(1.0), v(2.0)]), Value::Number(1.0));
    }

    #[test]
    fn a_lookup_takes_from_one_list_what_it_found_in_another() {
        // XLOOKUP is VLOOKUP with the column-counting taken out: the list to
        // search and the list to fetch from are two separate arguments, so
        // nothing depends on which column happens to be third.
        let keys = range(&[Value::text("apple"), Value::text("pear"), Value::text("plum")], 1);
        let pay = range(&[n(10.0), n(20.0), n(30.0)], 1);
        assert_eq!(call("XLOOKUP", &[t("pear"), keys.clone(), pay.clone()]), n(20.0));
        // Missing is #N/A, as any lookup would be, unless a fourth argument
        // says what to put there instead.
        assert_eq!(
            call("XLOOKUP", &[t("fig"), keys.clone(), pay.clone()]),
            Value::Error(ExcelError::NA)
        );
        assert_eq!(
            call("XLOOKUP", &[t("fig"), keys.clone(), pay.clone(), t("none")]),
            Value::text("none")
        );
    }

    #[test]
    fn a_lookup_can_settle_for_the_nearest_on_one_side() {
        let sizes = range(&[n(10.0), n(20.0), n(30.0)], 1);
        let names = range(&[Value::text("S"), Value::text("M"), Value::text("L")], 1);
        // -1 takes the nearest at or under the key, 1 the nearest at or over.
        // Nothing is sorted first, unlike VLOOKUP's approximate match.
        assert_eq!(
            call("XLOOKUP", &[v(25.0), sizes.clone(), names.clone(), v(0.0), v(-1.0)]),
            Value::text("M")
        );
        assert_eq!(
            call("XLOOKUP", &[v(25.0), sizes.clone(), names.clone(), v(0.0), v(1.0)]),
            Value::text("L")
        );
        // An exact hit is still preferred over either neighbour.
        assert_eq!(
            call("XLOOKUP", &[v(20.0), sizes.clone(), names.clone(), v(0.0), v(-1.0)]),
            Value::text("M")
        );
    }

    #[test]
    fn a_lookup_may_be_asked_to_start_at_the_bottom() {
        // Two rows answer; which one is returned is the whole point of the
        // sixth argument.
        let keys = range(&[Value::text("a"), Value::text("b"), Value::text("a")], 1);
        let pay = range(&[n(1.0), n(2.0), n(3.0)], 1);
        assert_eq!(call("XLOOKUP", &[t("a"), keys.clone(), pay.clone()]), n(1.0));
        assert_eq!(
            call("XLOOKUP", &[t("a"), keys, pay, v(0.0), v(0.0), v(-1.0)]),
            n(3.0)
        );
    }

    #[test]
    fn a_week_number_depends_on_which_day_opens_the_week() {
        // 2024-01-07 is a Sunday. It opens week 2 when weeks start on Sunday
        // and closes week 1 when they start on Monday, so a WEEKNUM that
        // ignored its second argument would still pass a Monday test.
        assert_eq!(call("WEEKNUM", &[v(45298.0)]), n(2.0));
        assert_eq!(call("WEEKNUM", &[v(45298.0), v(2.0)]), n(1.0));
        // 11 to 17 are Monday through Sunday, so 11 says what 2 says.
        assert_eq!(call("WEEKNUM", &[v(45298.0), v(11.0)]), n(1.0));
        assert_eq!(call("WEEKNUM", &[v(45298.0), v(17.0)]), n(2.0));
        assert_eq!(
            call("WEEKNUM", &[v(45292.0), v(9.0)]),
            Value::Error(ExcelError::Num)
        );
    }

    #[test]
    fn the_iso_week_can_belong_to_the_year_before_it() {
        // ISO weeks start on Monday and week one is the one holding the year's
        // first Thursday. 2021-01-01 was a Friday, so its week's Thursday fell
        // in 2020 and the date is in week 53 of that year — while the ordinary
        // count calls it week 1 of 2021.
        assert_eq!(call("WEEKNUM", &[v(44197.0), v(21.0)]), n(53.0));
        assert_eq!(call("WEEKNUM", &[v(44197.0)]), n(1.0));
        // 2024-01-01 was itself a Monday, so both counts agree.
        assert_eq!(call("WEEKNUM", &[v(45292.0), v(21.0)]), n(1.0));
    }

    /// Every expectation here is what Excel 16 returned for that formula.
    #[test]
    fn a_working_day_steps_over_the_weekend_and_over_the_holidays() {
        assert_eq!(call("WORKDAY", &[v(45292.0), v(5.0)]), n(45299.0));
        // 2024-01-01 was a Monday, so five working days on is the next Monday.
        assert_eq!(call("WORKDAY", &[v(45292.0), v(-3.0)]), n(45287.0));
        assert_eq!(call("WORKDAY", &[v(45292.0), v(0.0)]), n(45292.0));
        assert_eq!(call("WORKDAY", &[v(45293.0), v(1.0)]), n(45294.0));
        // A day named as a holiday is stepped over like a Saturday.
        assert_eq!(
            call("WORKDAY", &[v(45292.0), v(5.0), v(45294.0)]),
            n(45300.0)
        );
        // The corpus writes `WORKDAY(date,"")`, and Excel refuses it.
        assert_eq!(
            call("WORKDAY", &[v(45292.0), t("")]),
            Value::Error(ExcelError::Value)
        );
    }

    #[test]
    fn a_join_can_be_told_to_leave_the_blanks_out() {
        // Four cells of which two hold nothing. Leaving them out gives one
        // separator; keeping them gives three, one of them trailing.
        let cells = range(
            &[
                Value::text("one"),
                Value::text(""),
                Value::text("three"),
                Value::Blank,
            ],
            1,
        );
        assert_eq!(
            call("TEXTJOIN", &[t(", "), Arg::Value(Value::Logical(true)), cells.clone()]),
            Value::text("one, three")
        );
        assert_eq!(
            call("TEXTJOIN", &[t(", "), Arg::Value(Value::Logical(false)), cells]),
            Value::text("one, , three, ")
        );
        assert_eq!(
            call("TEXTJOIN", &[t("-"), Arg::Value(Value::Logical(true)), t("a"), t("b")]),
            Value::text("a-b")
        );
    }

    #[test]
    fn a_word_runs_on_through_letters_and_nothing_else() {
        assert_eq!(call("PROPER", &[t("o'neill-smith jr")]), Value::text("O'Neill-Smith Jr"));
        // The digit ends the word, so the r after it is a capital.
        assert_eq!(call("PROPER", &[t("ANNA MARIA 3rd")]), Value::text("Anna Maria 3Rd"));
    }

    #[test]
    fn the_text_of_a_thing_that_is_not_text_is_nothing_at_all() {
        assert_eq!(call("T", &[t("one")]), Value::text("one"));
        assert_eq!(call("T", &[v(7.0)]), Value::text(""));
        assert_eq!(call("T", &[Arg::Value(Value::Logical(true))]), Value::text(""));
    }

    /// A column of 10, <an error>, 30 — the shape every one of these is asked
    /// about. Each expectation is what Excel 16 answered.
    fn with_an_error(why: ExcelError) -> Arg {
        range(&[n(10.0), Value::Error(why), n(30.0)], 1)
    }

    #[test]
    fn an_error_in_a_range_being_tested_is_not_a_match() {
        // The guard that hands back the first error found anywhere in any
        // argument is right for SUM — a sum of an error IS an error — and
        // wrong for this whole family, where an error is a fact about one row.
        // `COUNTIF(range,"yes")` used to answer #N/A because one cell held one.
        let column = range(
            &[Value::text("yes"), Value::Error(ExcelError::NA), Value::text("yes")],
            1,
        );
        let amounts = range(&[n(10.0), n(20.0), n(30.0)], 1);
        assert_eq!(call("COUNTIF", &[column.clone(), t("yes")]), n(2.0));
        assert_eq!(
            call("SUMIF", &[column.clone(), t("yes"), amounts.clone()]),
            n(40.0),
        );
        assert_eq!(
            call("SUMIFS", &[amounts.clone(), column.clone(), t("yes")]),
            n(40.0),
        );
        assert_eq!(call("AVERAGEIF", &[column, t("yes"), amounts]), n(20.0));
    }

    #[test]
    fn an_error_on_a_row_that_matched_is_being_added_up() {
        // The other way round: the error is in the range being SUMMED, on a
        // row the criterion picked. There is no adding that up.
        let names = range(
            &[Value::text("yes"), Value::text("yes"), Value::text("no")],
            1,
        );
        let amounts = range(&[n(10.0), Value::Error(ExcelError::NA), n(30.0)], 1);
        assert_eq!(
            call("SUMIF", &[names.clone(), t("yes"), amounts.clone()]),
            Value::Error(ExcelError::NA),
        );
        // And on a row it did NOT pick, the error is simply not reached.
        assert_eq!(call("SUMIF", &[names, t("no"), amounts]), n(30.0));
    }

    #[test]
    fn a_criterion_that_spells_an_error_means_that_error() {
        // `"#N/A"` is the error, not the four characters. An error equals
        // itself, equals no other error, and equals no number — so a NOT-equal
        // criterion IS satisfied by one unless it names that same error.
        let na = with_an_error(ExcelError::NA);
        let bad_ref = with_an_error(ExcelError::Ref);
        let by_zero = with_an_error(ExcelError::DivZero);

        assert_eq!(call("COUNTIF", &[na.clone(), t("#N/A")]), n(1.0));
        assert_eq!(call("COUNTIF", &[na.clone(), t("<>#N/A")]), n(2.0));
        assert_eq!(call("SUMIF", &[na.clone(), t("<>#N/A")]), n(40.0));
        assert_eq!(
            call("SUMIF", &[na.clone(), t("#N/A")]),
            Value::Error(ExcelError::NA),
            "the row it picked holds an error",
        );

        // A DIFFERENT error is "not #N/A", so it matches — and then it is
        // being added up. This is what one corpus workbook does down a whole
        // column of #REF!, and fifty cells turned on getting it right.
        assert_eq!(call("COUNTIF", &[bad_ref.clone(), t("#N/A")]), n(0.0));
        assert_eq!(call("COUNTIF", &[bad_ref.clone(), t("<>#N/A")]), n(3.0));
        assert_eq!(
            call("SUMIF", &[bad_ref.clone(), t("<>#N/A")]),
            Value::Error(ExcelError::Ref),
        );
        assert_eq!(
            call("SUMIF", &[by_zero.clone(), t("<>#N/A")]),
            Value::Error(ExcelError::DivZero),
        );
        // Naming its own error excludes it again.
        assert_eq!(call("COUNTIF", &[bad_ref.clone(), t("<>#REF!")]), n(2.0));
        assert_eq!(call("SUMIF", &[bad_ref.clone(), t("<>#REF!")]), n(40.0));
    }

    #[test]
    fn an_error_is_past_comparing_for_greater_or_less() {
        // Beyond equality there is nothing to say about an error, so it falls
        // out of every comparison — but `"<>0"` is an equality, and an error
        // is indeed not zero.
        let na = with_an_error(ExcelError::NA);
        let bad_ref = with_an_error(ExcelError::Ref);
        assert_eq!(call("COUNTIF", &[na.clone(), t(">5")]), n(2.0));
        assert_eq!(call("SUMIF", &[na.clone(), t(">5")]), n(40.0));
        assert_eq!(call("COUNTIF", &[bad_ref.clone(), t(">5")]), n(2.0));
        assert_eq!(call("SUMIF", &[bad_ref, t(">5")]), n(40.0));
        assert_eq!(call("COUNTIF", &[na.clone(), t("<>0")]), n(3.0));
        assert_eq!(
            call("SUMIF", &[na, t("<>0")]),
            Value::Error(ExcelError::NA),
            "all three matched, and one of them is an error",
        );
    }

    #[test]
    fn a_criterion_that_is_an_error_looks_for_that_error() {
        // Measured in Excel: an error for the criterion counts and adds the
        // rows holding that same error, and here there are none.
        let amounts = range(&[n(10.0), n(20.0), n(30.0)], 1);
        let broken = Arg::Value(Value::Error(ExcelError::Value));
        assert_eq!(call("SUMIF", &[amounts.clone(), broken.clone()]), n(0.0));
        assert_eq!(call("SUMIFS", &[amounts.clone(), amounts, broken]), n(0.0));
    }

    /// 2 4 4 4 5 5 7 — a set chosen so the mean, the median and the mode are
    /// three different questions with three different answers.
    fn a_spread() -> Arg {
        range(&[n(2.0), n(4.0), n(4.0), n(4.0), n(5.0), n(5.0), n(7.0)], 1)
    }

    /// 1 2 3 4 — an even count, so the median and the quartiles have to land
    /// between two values rather than on one.
    fn four_in_a_row() -> Arg {
        range(&[n(1.0), n(2.0), n(3.0), n(4.0)], 1)
    }

    fn close_to(got: Value, want: f64, what: &str) {
        match got {
            Value::Number(held) => assert!(
                (held - want).abs() < 1e-9,
                "{what}: {held} is not {want}",
            ),
            other => panic!("{what}: {other:?} is not a number"),
        }
    }

    /// Every expectation is Excel 16's answer.
    #[test]
    fn the_middle_and_the_spread_of_a_set_of_numbers() {
        close_to(call("MEDIAN", &[a_spread()]), 4.0, "MEDIAN");
        // No single middle: the mean of the two there are.
        close_to(call("MEDIAN", &[four_in_a_row()]), 2.5, "MEDIAN of four");
        // `.S` divides by one less than the count, taking the values for a
        // sample; `.P` divides by the count, taking them for the whole.
        close_to(call("STDEV.S", &[a_spread()]), 1.511_857_892_036_909, "STDEV.S");
        close_to(call("STDEV.P", &[a_spread()]), 1.399_708_424_447_929_5, "STDEV.P");
        close_to(call("VAR.S", &[a_spread()]), 2.285_714_285_714_285_5, "VAR.S");
        close_to(call("VAR.P", &[a_spread()]), 1.959_183_673_469_387_7, "VAR.P");
        // The old names mean the sample forms.
        close_to(call("STDEV", &[a_spread()]), 1.511_857_892_036_909, "STDEV");
        close_to(call("VAR", &[a_spread()]), 2.285_714_285_714_285_5, "VAR");
        // One value is no sample at all.
        assert_eq!(
            call("STDEV.S", &[range(&[n(1.0)], 1)]),
            Value::Error(ExcelError::DivZero),
        );
    }

    #[test]
    fn a_mode_has_to_turn_up_more_than_once() {
        assert_eq!(call("MODE.SNGL", &[a_spread()]), n(4.0));
        assert_eq!(call("MODE", &[a_spread()]), n(4.0));
        assert_eq!(
            call("MODE.SNGL", &[range(&[n(1.0), n(2.0), n(3.0)], 1)]),
            Value::Error(ExcelError::NA),
            "nothing turned up twice",
        );
    }

    #[test]
    fn the_two_percentile_families_start_counting_in_different_places() {
        // Over 1 2 3 4: INC puts the rank at `p x (n-1)` from the first value,
        // so a quarter of the way is 0.75 along and reads 1.75. EXC puts it at
        // `p x (n+1)` counted from before the first, so a quarter is 1.25.
        close_to(call("PERCENTILE.INC", &[four_in_a_row(), v(0.25)]), 1.75, "INC");
        close_to(call("PERCENTILE.EXC", &[four_in_a_row(), v(0.25)]), 1.25, "EXC");
        close_to(call("PERCENTILE.INC", &[four_in_a_row(), v(0.0)]), 1.0, "the least");
        close_to(call("PERCENTILE.INC", &[four_in_a_row(), v(1.0)]), 4.0, "the most");
        close_to(call("PERCENTILE", &[four_in_a_row(), v(0.9)]), 3.7, "the old name");
        // A quartile is a percentile in quarters.
        close_to(call("QUARTILE.INC", &[four_in_a_row(), v(1.0)]), 1.75, "Q1");
        close_to(call("QUARTILE.INC", &[four_in_a_row(), v(2.0)]), 2.5, "Q2");
        close_to(call("QUARTILE.INC", &[four_in_a_row(), v(3.0)]), 3.25, "Q3");
        close_to(call("QUARTILE", &[four_in_a_row(), v(0.0)]), 1.0, "Q0");
        close_to(call("QUARTILE.EXC", &[four_in_a_row(), v(1.0)]), 1.25, "Q1 exclusive");
    }

    #[test]
    fn aggregate_can_be_told_to_pass_over_the_errors() {
        // 10, #N/A, 30, 40. The second argument is what to leave out: 2, 3, 6
        // and 7 leave out errors, and the corpus writes 6.
        let with_a_gap = range(&[n(10.0), Value::Error(ExcelError::NA), n(30.0), n(40.0)], 1);
        assert_eq!(call("AGGREGATE", &[v(15.0), v(6.0), with_a_gap.clone(), v(1.0)]), n(10.0));
        assert_eq!(call("AGGREGATE", &[v(15.0), v(6.0), with_a_gap.clone(), v(3.0)]), n(40.0));
        assert_eq!(call("AGGREGATE", &[v(14.0), v(6.0), with_a_gap.clone(), v(1.0)]), n(40.0));
        assert_eq!(call("AGGREGATE", &[v(9.0), v(6.0), with_a_gap.clone()]), n(80.0));
        assert_eq!(call("AGGREGATE", &[v(4.0), v(6.0), with_a_gap.clone()]), n(40.0));
        assert_eq!(call("AGGREGATE", &[v(12.0), v(6.0), with_a_gap.clone()]), n(30.0));
        close_to(
            call("AGGREGATE", &[v(1.0), v(6.0), with_a_gap.clone()]),
            26.666_666_666_666_668,
            "the mean of what is left",
        );
        // Option 0 leaves nothing out, so the error is the answer — as it is
        // for the aggregation on its own.
        assert_eq!(
            call("AGGREGATE", &[v(9.0), v(0.0), with_a_gap.clone()]),
            Value::Error(ExcelError::NA),
        );
        assert_eq!(call("SUM", &[with_a_gap.clone()]), Value::Error(ExcelError::NA));
        // There is no ninth of four.
        assert_eq!(
            call("AGGREGATE", &[v(15.0), v(6.0), with_a_gap, v(9.0)]),
            Value::Error(ExcelError::Num),
        );
    }

    /// 10, <an error>, 30, 40 — a block with one bad cell in the middle.
    fn a_block_with_a_gap() -> Arg {
        range(&[n(10.0), Value::Error(ExcelError::NA), n(30.0), n(40.0)], 1)
    }

    /// Excel 16's answers over that block, and over w x y z beside it.
    #[test]
    fn a_function_that_picks_does_not_mind_what_it_is_not_looking_at() {
        let gap = a_block_with_a_gap();
        let letters = range(
            &[Value::text("w"), Value::text("x"), Value::text("y"), Value::text("z")],
            1,
        );
        assert_eq!(call("INDEX", &[gap.clone(), v(1.0)]), n(10.0));
        assert_eq!(call("INDEX", &[gap.clone(), v(3.0)]), n(30.0));
        // At the cell picked, the error IS the answer.
        assert_eq!(
            call("INDEX", &[gap.clone(), v(2.0)]),
            Value::Error(ExcelError::NA),
        );
        assert_eq!(call("MATCH", &[v(30.0), gap.clone(), v(0.0)]), n(3.0));
        assert_eq!(
            call("MATCH", &[v(99.0), gap.clone(), v(0.0)]),
            Value::Error(ExcelError::NA),
            "not there is still not there",
        );
        assert_eq!(call("COUNT", &[gap.clone()]), n(3.0));
        assert_eq!(call("COUNTA", &[gap.clone()]), n(4.0));
        assert_eq!(
            call("VLOOKUP", &[t("y"), letters, v(1.0), Arg::Value(Value::Logical(false))]),
            Value::text("y"),
        );
        // The unchosen one is not looked at either.
        assert_eq!(
            call("CHOOSE", &[v(1.0), v(30.0), Arg::Value(Value::Error(ExcelError::NA))]),
            n(30.0),
        );
        // And the ones that must total or order the whole lot still mind it.
        assert_eq!(call("SUM", &[gap.clone()]), Value::Error(ExcelError::NA));
        assert_eq!(call("MAX", &[gap.clone()]), Value::Error(ExcelError::NA));
        assert_eq!(
            call("SMALL", &[gap, v(1.0)]),
            Value::Error(ExcelError::NA),
        );
    }

    #[test]
    fn an_error_where_the_block_should_be_is_a_block_that_is_not_there() {
        // `INDEX(#REF!,MATCH(x,#REF!,0))` is what Excel writes into a formula
        // whose external workbook has gone, and it answers `#REF!`. Treating
        // that as "an error to step over" left MATCH searching a nothing,
        // finding nothing, and answering `#N/A` — 185 cells of one workbook.
        //
        // The difference from the test above is the whole rule: a `#REF!`
        // among the values of a block is a value; a `#REF!` WHERE THE BLOCK
        // SHOULD BE is not.
        let missing = Arg::Value(Value::Error(ExcelError::Ref));
        assert_eq!(
            call("INDEX", &[missing.clone(), v(2.0)]),
            Value::Error(ExcelError::Ref),
        );
        assert_eq!(
            call("MATCH", &[v(30.0), missing.clone(), v(0.0)]),
            Value::Error(ExcelError::Ref),
        );
        assert_eq!(call("ROWS", &[missing]), Value::Error(ExcelError::Ref));
        // But COUNT, COUNTA and CHOOSE never mind one, however it arrives:
        // Excel gives 0, 1 and 30 for these.
        let gone = Arg::Value(Value::Error(ExcelError::Ref));
        assert_eq!(call("COUNT", &[gone.clone()]), n(0.0));
        assert_eq!(call("COUNTA", &[gone.clone()]), n(1.0));
        assert_eq!(call("CHOOSE", &[v(1.0), v(30.0), gone]), n(30.0));
        // Which holds for #N/A written out just the same, and the one CHOOSE
        // does pick still comes back whatever it is.
        let missing_value = Arg::Value(Value::Error(ExcelError::NA));
        assert_eq!(call("COUNT", &[missing_value.clone()]), n(0.0));
        assert_eq!(call("COUNTA", &[missing_value.clone()]), n(1.0));
        assert_eq!(call("CHOOSE", &[v(1.0), v(30.0), missing_value.clone()]), n(30.0));
        assert_eq!(
            call("CHOOSE", &[v(2.0), v(30.0), missing_value.clone()]),
            Value::Error(ExcelError::NA),
        );
        assert_eq!(call("COUNT", &[v(1.0), missing_value]), n(1.0));
    }

    #[test]
    fn looking_for_an_error_finds_nothing() {
        // The thing being searched FOR is not one of the values searched.
        let gap = a_block_with_a_gap();
        let broken = Arg::Value(Value::Error(ExcelError::NA));
        assert_eq!(
            call("MATCH", &[broken.clone(), gap.clone(), v(0.0)]),
            Value::Error(ExcelError::NA),
        );
        assert_eq!(
            call("INDEX", &[gap, broken]),
            Value::Error(ExcelError::NA),
            "nor is the number saying which one",
        );
    }

    /// How many numbers there are — a question Excel asks differently of an
    /// argument written out than of a value found inside a block.
    #[test]
    fn count_asks_two_questions() {
        let block = range(
            &[
                n(1.0),
                Value::Logical(true),
                n(2.0),
                Value::Error(ExcelError::NA),
                Value::text("text"),
            ],
            1,
        );
        // In a block: only what IS a number. The logical, the error and the
        // text are all not.
        assert_eq!(call("COUNT", &[block.clone()]), n(2.0));
        assert_eq!(call("COUNTA", &[block.clone()]), n(5.0));
        // Written out: anything that READS as a number.
        assert_eq!(call("COUNT", &[Arg::Value(Value::Logical(true))]), n(1.0));
        assert_eq!(call("COUNT", &[v(1.0), Arg::Value(Value::Logical(true))]), n(2.0));
        assert_eq!(call("COUNT", &[t("2")]), n(1.0), "text that reads as one");
        assert_eq!(call("COUNT", &[v(1.0), t("x")]), n(1.0), "and text that does not");
        assert_eq!(
            call("COUNT", &[Arg::Value(Value::Error(ExcelError::NA))]),
            n(0.0),
            "an error never reads as one",
        );
        assert_eq!(call("COUNT", &[block, Arg::Value(Value::Logical(true))]), n(3.0));
    }

    /// A sort ranks the KINDS of value before comparing within one, so an
    /// error is something to place rather than something to refuse.
    ///
    /// Excel 16, over a column holding 3, #N/A, 5, "zz", TRUE and a blank:
    /// ascending gives 3, 5, zz, TRUE, #N/A, blank; descending gives #N/A,
    /// TRUE, zz, 5, 3, blank. The blank is last BOTH ways — it takes no part
    /// in the reversal, which is what shows this to be a ranking.
    #[test]
    fn a_sort_puts_the_kinds_in_order_and_the_blanks_last() {
        let mixed = range(
            &[
                n(3.0),
                Value::Error(ExcelError::NA),
                n(5.0),
                Value::text("zz"),
                Value::Logical(true),
                Value::Blank,
            ],
            1,
        );
        let up = call_arg("SORT", &[mixed.clone(), Arg::Value(Value::Number(1.0)), v(1.0)]);
        let down = call_arg("SORT", &[mixed, Arg::Value(Value::Number(1.0)), v(-1.0)]);
        assert_eq!(
            up.flatten(),
            vec![
                n(3.0),
                n(5.0),
                Value::text("zz"),
                Value::Logical(true),
                Value::Error(ExcelError::NA),
                Value::Blank,
            ],
        );
        assert_eq!(
            down.flatten(),
            vec![
                Value::Error(ExcelError::NA),
                Value::Logical(true),
                Value::text("zz"),
                n(5.0),
                n(3.0),
                Value::Blank,
            ],
        );
    }

    #[test]
    fn a_star_stands_for_any_run_of_characters() {
        assert!(wildcard_match("life insurance", "life*"));
        assert!(wildcard_match("life insurance", "*insurance"));
        assert!(wildcard_match("life insurance", "*e i*"));
        assert!(wildcard_match("anything", "*"));
        assert!(!wildcard_match("car insurance", "life*"));
        // Backing out of a dead end: the first `b` does not lead anywhere, so
        // the second has to be tried.
        assert!(wildcard_match("aXbY", "a*bY"));
        assert!(!wildcard_match("aXbY", "a*bZ"));
    }

    #[test]
    fn a_question_mark_stands_for_exactly_one() {
        assert!(wildcard_match("cat", "c?t"));
        assert!(!wildcard_match("coat", "c?t"));
        assert!(wildcard_match("coat", "c??t"));
    }

    #[test]
    fn a_tilde_means_the_character_itself() {
        assert!(wildcard_match("10%*", "10%~*"));
        assert!(!wildcard_match("10%x", "10%~*"));
        assert!(wildcard_match("what?", "what~?"));
    }

    #[test]
    fn matching_pays_no_attention_to_capitals() {
        // Excel's text comparison does not, and neither does this.
        assert!(wildcard_match("Life Insurance", "life*"));
        assert!(wildcard_match("life insurance", "LIFE*"));
    }

    #[test]
    fn the_lookups_read_wildcards_when_asked_for_an_exact_match() {
        // `VLOOKUP(D1 & "*", ...)` is the ordinary way to look something up by
        // its beginning, and comparing the pattern as literal text finds
        // nothing at all.
        let table = range(
            &[
                Value::text("life insurance"),
                Value::text("yes"),
                Value::text("car"),
                Value::text("no"),
            ],
            2,
        );
        assert_eq!(
            call("VLOOKUP", &[t("life*"), table.clone(), v(2.0), v(0.0)]),
            Value::text("yes")
        );
        let column = range(&[Value::text("life insurance"), Value::text("car")], 1);
        assert_eq!(
            call("MATCH", &[t("*insurance"), column.clone(), v(0.0)]),
            Value::Number(1.0)
        );
        assert_eq!(call("COUNTIF", &[column.clone(), t("*i*")]), Value::Number(1.0));
        // The approximate form sorts rather than matches, so the star there is
        // just a character. Against this unsorted pair that lands on "car" and
        // answers "no" — which is the point: whatever it is, it is not the
        // pattern match, and a star must not turn the sorted form into one.
        assert_eq!(
            call("VLOOKUP", &[t("life*"), table, v(2.0), v(1.0)]),
            Value::text("no")
        );
    }

    #[test]
    fn a_criterion_with_no_wildcard_in_it_is_still_plain_equality() {
        let column = range(&[Value::text("a"), Value::text("ab")], 1);
        assert_eq!(call("COUNTIF", &[column.clone(), t("a")]), Value::Number(1.0));
        assert_eq!(call("COUNTIF", &[column, t("a*")]), Value::Number(2.0));
    }

    #[test]
    fn sumproduct_multiplies_across_and_adds_up() {
        let a = range(&[n(1.0), n(2.0), n(3.0)], 1);
        let b = range(&[n(4.0), n(5.0), n(6.0)], 1);
        // 1*4 + 2*5 + 3*6
        assert_eq!(call("SUMPRODUCT", &[a.clone(), b]), Value::Number(32.0));
        // One array on its own is just its sum.
        assert_eq!(call("SUMPRODUCT", &[a]), Value::Number(6.0));
    }

    #[test]
    fn sumproduct_weighs_a_condition_as_one_or_nothing() {
        // A column of TRUE and FALSE handed over as it is weighs NOTHING --
        // it has to be made numbers first with `--` or `*`: measured,
        // SUMPRODUCT((A1:A3>0),B1:B3) and SUMPRODUCT({TRUE,FALSE,TRUE},{1,2,3})
        // are 0 where SUMPRODUCT(--(A1:A3>0),B1:B3) is 40. Text and blanks
        // weigh nothing too rather than spoil the sum.
        let flags = range(
            &[Value::Logical(true), Value::Logical(false), Value::Logical(true)],
            1,
        );
        let amounts = range(&[n(10.0), n(20.0), n(30.0)], 1);
        assert_eq!(call("SUMPRODUCT", &[flags, amounts]), Value::Number(0.0));
        let mixed = range(&[n(2.0), Value::text("x"), Value::Blank], 1);
        let ones = range(&[n(1.0), n(1.0), n(1.0)], 1);
        assert_eq!(call("SUMPRODUCT", &[mixed, ones]), Value::Number(2.0));
    }

    #[test]
    fn sumproduct_refuses_arrays_of_different_lengths() {
        let three = range(&[n(1.0), n(2.0), n(3.0)], 1);
        let two = range(&[n(1.0), n(2.0)], 1);
        assert_eq!(
            call("SUMPRODUCT", &[three, two]),
            Value::Error(ExcelError::Value)
        );
    }

    #[test]
    fn sumifs_reads_its_ranges_the_other_way_round_from_sumif() {
        // SUMIF puts the range to test first and the range to add last; SUMIFS
        // puts the range to add FIRST. Writing one as the other is the classic
        // way to get a plausible wrong answer.
        let amounts = range(&[n(10.0), n(20.0), n(30.0)], 1);
        let region = range(
            &[Value::text("N"), Value::text("S"), Value::text("N")],
            1,
        );
        assert_eq!(
            call("SUMIFS", &[amounts.clone(), region.clone(), t("N")]),
            Value::Number(40.0)
        );
        assert_eq!(
            call("COUNTIFS", &[region.clone(), t("N")]),
            Value::Number(2.0)
        );
        assert_eq!(
            call("AVERAGEIFS", &[amounts.clone(), region.clone(), t("N")]),
            Value::Number(20.0)
        );
        // Two conditions, both of which must hold.
        let size = range(&[n(1.0), n(1.0), n(2.0)], 1);
        assert_eq!(
            call("SUMIFS", &[amounts, region, t("N"), size, t("1")]),
            Value::Number(10.0)
        );
    }

    #[test]
    fn averageifs_of_nothing_is_a_division_by_zero() {
        let amounts = range(&[n(10.0)], 1);
        let region = range(&[Value::text("N")], 1);
        assert_eq!(
            call("AVERAGEIFS", &[amounts, region, t("S")]),
            Value::Error(ExcelError::DivZero)
        );
    }

    #[test]
    fn rows_and_columns_report_the_shape_of_what_they_are_given() {
        let block = range(&[n(1.0), n(2.0), n(3.0), n(4.0), n(5.0), n(6.0)], 3);
        assert_eq!(call("ROWS", &[block.clone()]), Value::Number(2.0));
        assert_eq!(call("COLUMNS", &[block]), Value::Number(3.0));
        // A single value is a block one by one.
        assert_eq!(call("ROWS", &[v(5.0)]), Value::Number(1.0));
    }

    #[test]
    fn small_and_large_count_from_opposite_ends() {
        let data = range(&[n(5.0), n(1.0), n(9.0), n(3.0)], 1);
        assert_eq!(call("SMALL", &[data.clone(), v(1.0)]), Value::Number(1.0));
        assert_eq!(call("SMALL", &[data.clone(), v(3.0)]), Value::Number(5.0));
        assert_eq!(call("LARGE", &[data.clone(), v(1.0)]), Value::Number(9.0));
        assert_eq!(call("LARGE", &[data.clone(), v(2.0)]), Value::Number(5.0));
        // Past the end of the list is #NUM!, not the last one.
        assert_eq!(
            call("SMALL", &[data, v(9.0)]),
            Value::Error(ExcelError::Num)
        );
    }

    #[test]
    fn small_ignores_what_is_not_a_number() {
        let data = range(&[n(5.0), Value::text("x"), Value::Blank, n(1.0)], 1);
        assert_eq!(call("SMALL", &[data.clone(), v(1.0)]), Value::Number(1.0));
        assert_eq!(call("SMALL", &[data, v(2.0)]), Value::Number(5.0));
    }

    #[test]
    fn rank_gives_equal_numbers_the_same_place_and_skips_the_next() {
        let data = range(&[n(9.0), n(9.0), n(5.0)], 1);
        assert_eq!(call("RANK", &[v(9.0), data.clone()]), Value::Number(1.0));
        // Two firsts are followed by a third, not a second.
        assert_eq!(call("RANK", &[v(5.0), data.clone()]), Value::Number(3.0));
        // Counting up instead of down.
        assert_eq!(
            call("RANK", &[v(5.0), data.clone(), v(1.0)]),
            Value::Number(1.0)
        );
        assert_eq!(
            call("RANK", &[v(7.0), data]),
            Value::Error(ExcelError::NA)
        );
    }

    #[test]
    fn ceiling_and_floor_move_to_a_multiple() {
        assert_eq!(call("CEILING", &[v(4.2), v(1.0)]), Value::Number(5.0));
        assert_eq!(call("CEILING", &[v(4.2), v(0.5)]), Value::Number(4.5));
        assert_eq!(call("FLOOR", &[v(4.8), v(0.5)]), Value::Number(4.5));
        assert_eq!(call("CEILING", &[v(-4.2), v(1.0)]), Value::Number(-4.0));
        // The older CEILING refuses a positive number and a negative step;
        // the .MATH form takes one.
        assert_eq!(
            call("CEILING", &[v(4.2), v(-1.0)]),
            Value::Error(ExcelError::Num)
        );
        assert_eq!(call("CEILING.MATH", &[v(4.2)]), Value::Number(5.0));
    }

    #[test]
    fn exact_is_the_comparison_that_notices_capitals() {
        assert_eq!(call("EXACT", &[t("Word"), t("Word")]), Value::Logical(true));
        assert_eq!(call("EXACT", &[t("Word"), t("word")]), Value::Logical(false));
    }

    #[test]
    fn char_and_code_are_each_others_undoing() {
        assert_eq!(call("CHAR", &[v(65.0)]), Value::text("A"));
        assert_eq!(call("CODE", &[t("A")]), Value::Number(65.0));
        assert_eq!(call("CODE", &[t("Apple")]), Value::Number(65.0));
        assert_eq!(call("CHAR", &[v(0.0)]), Value::Error(ExcelError::Value));
        assert_eq!(call("CHAR", &[v(300.0)]), Value::Error(ExcelError::Value));
    }

    #[test]
    fn replace_puts_something_over_what_was_there() {
        assert_eq!(
            call("REPLACE", &[t("abcdef"), v(2.0), v(3.0), t("XY")]),
            Value::text("aXYef")
        );
        // Nothing taken out is an insertion.
        assert_eq!(
            call("REPLACE", &[t("abc"), v(2.0), v(0.0), t("-")]),
            Value::text("a-bc")
        );
        // Counted in the same units as LEN, so a surrogate pair is two.
        assert_eq!(
            call("REPLACE", &[t("𠮷野"), v(1.0), v(2.0), t("Y")]),
            Value::text("Y野")
        );
    }

    #[test]
    fn text_writes_a_number_the_way_a_cell_would_show_it() {
        assert_eq!(call("TEXT", &[v(1234.5), t("0.00")]), Value::text("1234.50"));
        // Text handed to it comes back untouched: a number format has nothing
        // to say about it.
        assert_eq!(call("TEXT", &[t("already"), t("0.00")]), Value::text("already"));
    }

    #[test]
    fn len_counts_utf16_units_like_excel() {
        // The prototype this replaces returned byte length, so LEN("あ") was 3.
        assert_eq!(call("LEN", &[t("あ")]), Value::Number(1.0));
        assert_eq!(call("LEN", &[t("単価")]), Value::Number(2.0));
        // Surrogate pair: Excel counts 2, and so do we.
        assert_eq!(call("LEN", &[t("𠮷")]), Value::Number(2.0));
    }

    #[test]
    fn left_and_mid_slice_by_utf16_units() {
        assert_eq!(call("LEFT", &[t("東京都港区"), v(3.0)]), Value::text("東京都"));
        assert_eq!(call("MID", &[t("東京都港区"), v(4.0), v(2.0)]), Value::text("港区"));
        assert_eq!(call("RIGHT", &[t("東京都港区"), v(2.0)]), Value::text("港区"));
    }

    #[test]
    fn int_floors_toward_negative_infinity() {
        assert_eq!(call("INT", &[v(-1.5)]), Value::Number(-2.0));
        assert_eq!(call("INT", &[v(1.5)]), Value::Number(1.0));
    }

    #[test]
    fn mod_takes_the_sign_of_the_divisor() {
        // Rust's `%` would give -1 here.
        assert_eq!(call("MOD", &[v(-3.0), v(2.0)]), Value::Number(1.0));
        assert_eq!(call("MOD", &[v(3.0), v(-2.0)]), Value::Number(-1.0));
        assert_eq!(call("MOD", &[v(3.0), v(0.0)]), Value::Error(ExcelError::DivZero));
    }

    #[test]
    fn round_family_behaves_like_excel() {
        assert_eq!(call("ROUND", &[v(2.5), v(0.0)]), Value::Number(3.0));
        assert_eq!(call("ROUND", &[v(-2.5), v(0.0)]), Value::Number(-3.0));
        assert_eq!(call("ROUNDUP", &[v(1.1), v(0.0)]), Value::Number(2.0));
        assert_eq!(call("ROUNDDOWN", &[v(1.9), v(0.0)]), Value::Number(1.0));
        assert_eq!(call("ROUNDDOWN", &[v(-1.9), v(0.0)]), Value::Number(-1.0));
    }

    #[test]
    fn aggregates_ignore_text_inside_ranges() {
        let data = range(
            &[Value::Number(1.0), Value::text("x"), Value::Number(2.0)],
            3,
        );
        assert_eq!(call("SUM", std::slice::from_ref(&data)), Value::Number(3.0));
        assert_eq!(call("COUNT", std::slice::from_ref(&data)), Value::Number(2.0));
        assert_eq!(call("COUNTA", &[data]), Value::Number(3.0));
    }

    #[test]
    fn errors_propagate_out_of_aggregates() {
        let data = range(&[Value::Number(1.0), Value::Error(ExcelError::DivZero)], 2);
        assert_eq!(call("SUM", &[data]), Value::Error(ExcelError::DivZero));
    }

    #[test]
    fn iferror_sees_the_error_instead_of_propagating_it() {
        let bad = Arg::Value(Value::Error(ExcelError::DivZero));
        assert_eq!(call("IFERROR", &[bad, t("fallback")]), Value::text("fallback"));
    }

    #[test]
    fn vlookup_exact_and_approximate() {
        let table = range(
            &[
                Value::Number(1.0),
                Value::text("one"),
                Value::Number(5.0),
                Value::text("five"),
                Value::Number(9.0),
                Value::text("nine"),
            ],
            2,
        );
        // Approximate: 7 falls into the 5 bucket.
        assert_eq!(
            call("VLOOKUP", &[v(7.0), table.clone(), v(2.0)]),
            Value::text("five")
        );
        // Exact: 7 is absent.
        assert_eq!(
            call("VLOOKUP", &[v(7.0), table.clone(), v(2.0), Arg::Value(Value::Logical(false))]),
            Value::Error(ExcelError::NA)
        );
        assert_eq!(
            call("VLOOKUP", &[v(5.0), table, v(2.0), Arg::Value(Value::Logical(false))]),
            Value::text("five")
        );
    }

    #[test]
    fn countif_and_sumif_parse_comparison_criteria() {
        let data = range(
            &[Value::Number(1.0), Value::Number(5.0), Value::Number(9.0)],
            3,
        );
        assert_eq!(call("COUNTIF", &[data.clone(), t(">4")]), Value::Number(2.0));
        assert_eq!(call("SUMIF", &[data.clone(), t(">=5")]), Value::Number(14.0));
        assert_eq!(call("COUNTIF", &[data, t("<>5")]), Value::Number(2.0));
    }

    #[test]
    fn index_and_match_address_one_based() {
        let data = range(
            &[Value::text("a"), Value::text("b"), Value::text("c")],
            1,
        );
        assert_eq!(call("MATCH", &[t("b"), data.clone(), v(0.0)]), Value::Number(2.0));
        assert_eq!(call("INDEX", &[data, v(2.0)]), Value::text("b"));
    }

    #[test]
    fn unknown_functions_report_name_error() {
        assert_eq!(call("NOTAFUNCTION", &[v(1.0)]), Value::Error(ExcelError::Name));
    }

    #[test]
    fn date_parts_round_trip() {
        let serial = call("DATE", &[v(2026.0), v(7.0), v(26.0)]);
        assert_eq!(serial, Value::Number(46229.0));
        let s = Arg::Value(serial);
        assert_eq!(call("YEAR", std::slice::from_ref(&s)), Value::Number(2026.0));
        assert_eq!(call("MONTH", std::slice::from_ref(&s)), Value::Number(7.0));
        assert_eq!(call("DAY", &[s]), Value::Number(26.0));
    }

    #[test]
    fn weekday_types_agree_with_excel() {
        // 2026-07-26 is a Sunday.
        let sunday = Arg::Value(call("DATE", &[v(2026.0), v(7.0), v(26.0)]));
        assert_eq!(call("WEEKDAY", std::slice::from_ref(&sunday)), Value::Number(1.0));
        assert_eq!(call("WEEKDAY", &[sunday.clone(), v(2.0)]), Value::Number(7.0));
        assert_eq!(call("WEEKDAY", &[sunday.clone(), v(3.0)]), Value::Number(6.0));
        assert_eq!(call("WEEKDAY", &[sunday, v(99.0)]), Value::Error(ExcelError::Num));
    }

    #[test]
    fn edate_and_eomonth_handle_the_fiscal_year() {
        let apr1 = Arg::Value(call("DATE", &[v(2026.0), v(4.0), v(1.0)]));
        let year_end = call("EOMONTH", &[apr1.clone(), v(11.0)]);
        assert_eq!(year_end, call("DATE", &[v(2027.0), v(3.0), v(31.0)]));
        // EDATE clamps rather than overflowing into the next month.
        let jan31 = Arg::Value(call("DATE", &[v(2026.0), v(1.0), v(31.0)]));
        assert_eq!(
            call("EDATE", &[jan31, v(1.0)]),
            call("DATE", &[v(2026.0), v(2.0), v(28.0)])
        );
    }

    #[test]
    fn datedif_computes_whole_years() {
        let birth = Arg::Value(call("DATE", &[v(1990.0), v(8.0), v(1.0)]));
        let today = Arg::Value(call("DATE", &[v(2026.0), v(7.0), v(26.0)]));
        // Birthday has not arrived yet this year.
        assert_eq!(
            call("DATEDIF", &[birth.clone(), today.clone(), t("Y")]),
            Value::Number(35.0)
        );
        assert_eq!(
            call("DATEDIF", &[birth, today, t("YM")]),
            Value::Number(11.0)
        );
    }

    #[test]
    fn time_components_extract_cleanly() {
        let noonish = Arg::Value(call("TIME", &[v(11.0), v(59.0), v(59.0)]));
        assert_eq!(call("HOUR", std::slice::from_ref(&noonish)), Value::Number(11.0));
        assert_eq!(call("MINUTE", std::slice::from_ref(&noonish)), Value::Number(59.0));
        assert_eq!(call("SECOND", &[noonish]), Value::Number(59.0));
    }

    #[test]
    fn scalar_context_rejects_multi_cell_ranges() {
        let data = range(&[Value::text("ab"), Value::text("cd")], 2);
        assert_eq!(call("LEN", &[data]), Value::Error(ExcelError::Value));
    }
}
