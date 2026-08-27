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

fn num(arg: &Arg) -> Result<f64, ExcelError> {
    arg.scalar().to_number()
}

fn text(arg: &Arg) -> Result<String, ExcelError> {
    arg.scalar().to_text()
}

/// Numbers only, the way `SUM` and `AVERAGE` see a range: text and logicals
/// inside a *range* are skipped, but a directly supplied argument is coerced.
fn numeric_operands(args: &[Arg]) -> Result<Vec<f64>, ExcelError> {
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
        // one piece of text
            | "LEN" | "LEFT" | "RIGHT" | "MID" | "LOWER" | "UPPER" | "TRIM"
            | "FIND" | "SEARCH" | "SUBSTITUTE" | "REPLACE" | "REPT" | "EXACT"
            | "CHAR" | "CODE" | "UNICODE" | "TEXT" | "VALUE" | "PROPER" | "T"
        // one date
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
    }
    // The three that hand back a block rather than a value.
    if matches!(name, "UNIQUE" | "SORT" | "FILTER" | "SORTBY") {
        return match a_block_of_rows(name, args) {
            Ok(block) => block,
            Err(why) => Arg::Value(Value::Error(why)),
        };
    }
    if one_at_a_time(name) {
        let width = args.iter().map(|one| block_of(one).width).max().unwrap_or(1);
        let height = args.iter().map(|one| block_of(one).height).max().unwrap_or(1);
        if width * height > 1 {
            let blocks: Vec<RangeData> = args.iter().map(block_of).collect();
            let mut cells = Vec::with_capacity(width * height);
            for row in 0..height {
                for col in 0..width {
                    let picked: Vec<Arg> = blocks
                        .iter()
                        .map(|block| {
                            Arg::Value(
                                reach(block, col, row)
                                    .unwrap_or(Value::Error(ExcelError::NA)),
                            )
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
        "IFERROR" | "IFNA" | "IF"
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
                .filter(|v| v.is_blank())
                .count() as f64,
        )),

        // ---- arithmetic --------------------------------------------------
        "ABS" => Ok(Value::Number(one(args)?.abs())),
        "SQRT" => {
            let n = one(args)?;
            if n < 0.0 {
                return Err(ExcelError::Num);
            }
            Ok(Value::Number(n.sqrt()))
        }
        "POWER" => {
            expect(args, 2)?;
            Ok(Value::Number(num(&args[0])?.powf(num(&args[1])?)))
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
        "ROUND" | "ROUNDUP" | "ROUNDDOWN" => {
            let n = num(&args.first().ok_or(ExcelError::Value)?.clone())?;
            let digits = match args.get(1) {
                Some(a) => num(a)? as i32,
                None => 0,
            };
            let factor = 10f64.powi(digits);
            let scaled = n * factor;
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
        "TRIM" => {
            // Excel's TRIM also collapses runs of interior spaces to one.
            let s = text(one_arg(args)?)?;
            let collapsed = s.split_whitespace().collect::<Vec<_>>().join(" ");
            Ok(Value::Text(collapsed))
        }
        "UPPER" => Ok(Value::Text(text(one_arg(args)?)?.to_uppercase())),
        "LOWER" => Ok(Value::Text(text(one_arg(args)?)?.to_lowercase())),
        "CONCATENATE" | "CONCAT" => {
            let mut out = String::new();
            for v in args.iter().flat_map(|a| a.flatten()) {
                out.push_str(&v.to_text()?);
            }
            Ok(Value::Text(out))
        }
        "REPT" => {
            expect(args, 2)?;
            let n = num(&args[1])?;
            if n < 0.0 {
                return Err(ExcelError::Value);
            }
            Ok(Value::Text(text(&args[0])?.repeat(n as usize)))
        }
        "SUBSTITUTE" => {
            if args.len() < 3 {
                return Err(ExcelError::Value);
            }
            let (s, old, new) = (text(&args[0])?, text(&args[1])?, text(&args[2])?);
            if old.is_empty() {
                return Ok(Value::Text(s));
            }
            Ok(Value::Text(s.replace(&old, &new)))
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
            let start = match args.get(2) {
                Some(a) => num(a)?.max(1.0) as usize - 1,
                None => 0,
            };
            let units = utf16(&haystack);
            let needle_units = utf16(&needle);
            if start > units.len() {
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
            for pair in pairs.chunks(2) {
                if let Some(why) = pair[1].scalar().err() {
                    return Err(why);
                }
            }
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
                        Value::Logical(true) => 1.0,
                        Value::Logical(false) | Value::Blank | Value::Text(_) => 0.0,
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
            if step == 0.0 {
                return Ok(Value::Number(0.0));
            }
            // Excel refuses a positive number rounded to a negative step.
            if value > 0.0 && step < 0.0 && !name.ends_with(".MATH") {
                return Err(ExcelError::Num);
            }
            let up = name.starts_with("CEILING");
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
            match args[0].scalar() {
                Value::Number(n) => Ok(Value::Text(crate::numfmt::format_number(n, &format))),
                // Text handed to TEXT comes back as it was: there is nothing
                // for a number format to do to it.
                Value::Text(t) => Ok(Value::Text(t)),
                Value::Logical(b) => Ok(Value::Text(
                    if b { "TRUE" } else { "FALSE" }.to_string(),
                )),
                Value::Blank => Ok(Value::Text(String::new())),
                Value::Error(e) => Err(e),
            }
        }

        "SUMIF" | "AVERAGEIF" => {
            if args.len() < 2 {
                return Err(ExcelError::Value);
            }
            let asked = args[1].scalar();
            // The criterion itself being an error is a different matter from a
            // range holding one: there is nothing to test against.
            if let Some(why) = asked.err() {
                return Err(why);
            }
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
                None => match args.get(3) {
                    Some(one) => Ok(one.scalar()),
                    None => Err(ExcelError::NA),
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

        "VLOOKUP" | "HLOOKUP" => {
            if args.len() < 3 {
                return Err(ExcelError::Value);
            }
            let key = args[0].scalar();
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
                let mut best = None;
                for i in 0..lanes {
                    match compare(&probe(i), &key) {
                        Ok(Ordering::Greater) => break,
                        Ok(_) => best = Some(i),
                        Err(_) => continue,
                    }
                }
                best.map(fetch).ok_or(ExcelError::NA)
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
                m if m > 0 => {
                    let mut best = None;
                    for (i, v) in haystack.iter().enumerate() {
                        match compare(v, &key) {
                            Ok(Ordering::Greater) => break,
                            Ok(_) => best = Some(i),
                            Err(_) => continue,
                        }
                    }
                    best
                }
                _ => {
                    let mut best = None;
                    for (i, v) in haystack.iter().enumerate() {
                        match compare(v, &key) {
                            Ok(Ordering::Less) => break,
                            Ok(_) => best = Some(i),
                            Err(_) => continue,
                        }
                    }
                    best
                }
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
        "TIME" => {
            expect(args, 3)?;
            Ok(Value::Number(datetime::fraction_from_time(
                num(&args[0])?,
                num(&args[1])?,
                num(&args[2])?,
            )))
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
            let (h, m, s) = datetime::time_from_fraction(one(args)?);
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
        "DATEDIF" => {
            expect(args, 3)?;
            let unit = text(&args[2])?;
            Ok(Value::Number(datedif(
                serial(&args[0])?,
                serial(&args[1])?,
                &unit,
            )?))
        }

        _ => Err(ExcelError::Name),
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
                None => match args.get(1) {
                    Some(one) => num(one)? as usize,
                    None => 1,
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
            _ => matches!(compare(one, other), Ok(Ordering::Equal)),
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
        "MD" => {
            let anchor = datetime::add_months(start, months)?;
            Ok((end - anchor) as f64)
        }
        "YD" => {
            let anchor = datetime::add_months(start, (months / 12) * 12)?;
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

        let operand = match rest.parse::<f64>() {
            Ok(n) => Value::Number(n),
            // `"<>#N/A"` names the error, not the four characters of it.
            Err(_) => match an_error_named(rest) {
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
        // Otherwise a blank satisfies no comparison.
        if v.is_blank() {
            return false;
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
    fn the_criterion_itself_being_an_error_is_a_different_matter() {
        // A range holding an error is a fact about a row. A CRITERION that is
        // an error leaves nothing to test against at all.
        let amounts = range(&[n(10.0), n(20.0), n(30.0)], 1);
        let broken = Arg::Value(Value::Error(ExcelError::Value));
        assert_eq!(
            call("SUMIF", &[amounts.clone(), broken.clone()]),
            Value::Error(ExcelError::Value),
        );
        assert_eq!(
            call("SUMIFS", &[amounts.clone(), amounts, broken]),
            Value::Error(ExcelError::Value),
        );
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
        // This is what the function is nearly always for: a column of TRUE and
        // FALSE picking out which of another column to add. Text and blanks
        // have to weigh nothing rather than spoil the sum.
        let flags = range(
            &[Value::Logical(true), Value::Logical(false), Value::Logical(true)],
            1,
        );
        let amounts = range(&[n(10.0), n(20.0), n(30.0)], 1);
        assert_eq!(call("SUMPRODUCT", &[flags, amounts]), Value::Number(40.0));
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
