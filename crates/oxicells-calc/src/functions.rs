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

/// Return the first error among the arguments, so that errors propagate the
/// way Excel propagates them (before the function body runs).
fn first_error(args: &[Arg]) -> Option<ExcelError> {
    args.iter()
        .flat_map(|a| a.flatten())
        .find_map(|v| v.err())
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
pub fn call(name: &str, args: &[Arg]) -> Value {
    match dispatch(name, args) {
        Ok(v) => v,
        Err(e) => Value::Error(e),
    }
}

fn dispatch(name: &str, args: &[Arg]) -> Result<Value, ExcelError> {
    // IFERROR/IFNA/ISERROR/ISNA must see errors rather than propagate them.
    let error_transparent = matches!(name, "IFERROR" | "IFNA" | "ISERROR" | "ISNA" | "ISERR" | "IF");
    if !error_transparent {
        if let Some(e) = first_error(args) {
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
        "COUNT" => Ok(Value::Number(numeric_operands(args)?.len() as f64)),
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
        "SUMIF" => {
            if args.len() < 2 {
                return Err(ExcelError::Value);
            }
            let criteria = Criteria::parse(&args[1].scalar());
            let tested = args[0].flatten();
            let summed = match args.get(2) {
                Some(a) => a.flatten(),
                None => tested.clone(),
            };
            let mut total = 0.0;
            for (i, v) in tested.iter().enumerate() {
                if criteria.matches(v) {
                    if let Some(Value::Number(n)) = summed.get(i) {
                        total += n;
                    }
                }
            }
            Ok(Value::Number(total))
        }

        // ---- lookup --------------------------------------------------------
        "VLOOKUP" | "HLOOKUP" => {
            if args.len() < 3 {
                return Err(ExcelError::Value);
            }
            let key = args[0].scalar();
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
                    .find(|&i| {
                        let candidate = probe(i);
                        !candidate.is_blank() && compare(&candidate, &key) == Ok(Ordering::Equal)
                    })
                    .map(fetch)
                    .ok_or(ExcelError::NA)
            }
        }
        "MATCH" => {
            if args.len() < 2 {
                return Err(ExcelError::Value);
            }
            let key = args[0].scalar();
            let haystack = args[1].flatten();
            let mode = match args.get(2) {
                Some(a) => num(a)? as i32,
                None => 1,
            };
            let found = match mode {
                0 => haystack
                    .iter()
                    .position(|v| !v.is_blank() && compare(v, &key) == Ok(Ordering::Equal)),
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
                9 => "SUM",
                // 7, 8, 10, 11 are STDEV/STDEVP/VAR/VARP, not implemented yet.
                _ => return Err(ExcelError::Value),
            };
            dispatch(inner, &args[1..])
        }

        // ---- date and time ---------------------------------------------
        // NOW/TODAY are absent on purpose: see the module docs.
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
            Err(_) => Value::Text(rest.to_string()),
        };
        Criteria { op, operand }
    }

    fn matches(&self, v: &Value) -> bool {
        // A blank never satisfies a comparison criterion.
        if v.is_blank() {
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
        assert_eq!(call("XLOOKUP", &[v(1.0)]), Value::Error(ExcelError::Name));
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
