// This Source Code Form is subject to the terms of the Mozilla Public
// License, v. 2.0. If a copy of the MPL was not distributed with this
// file, You can obtain one at https://mozilla.org/MPL/2.0/.

//! Excel's value model, error values, and coercion rules.
//!
//! The coercion rules here are the part most often gotten wrong by naive
//! implementations, and they are exactly the part that shows up as a divergence
//! when comparing against Excel via COM. Each non-obvious rule is documented
//! with the behaviour it reproduces.

use std::cmp::Ordering;
use std::fmt;

/// The seven Excel error values.
#[derive(Debug, Clone, Copy, PartialEq, Eq, PartialOrd, Ord, Hash)]
pub enum ExcelError {
    /// `#NULL!` — intersection of two ranges that do not intersect.
    Null,
    /// `#DIV/0!`
    DivZero,
    /// `#VALUE!` — wrong type of argument.
    Value,
    /// `#REF!` — reference to a cell that does not exist.
    Ref,
    /// `#NAME?` — unrecognised function or defined name.
    Name,
    /// `#NUM!` — numeric overflow or invalid numeric argument.
    Num,
    /// `#N/A` — value not available (lookup miss).
    NA,
}

impl ExcelError {
    pub fn as_str(self) -> &'static str {
        match self {
            ExcelError::Null => "#NULL!",
            ExcelError::DivZero => "#DIV/0!",
            ExcelError::Value => "#VALUE!",
            ExcelError::Ref => "#REF!",
            ExcelError::Name => "#NAME?",
            ExcelError::Num => "#NUM!",
            ExcelError::NA => "#N/A",
        }
    }
}

impl fmt::Display for ExcelError {
    fn fmt(&self, f: &mut fmt::Formatter<'_>) -> fmt::Result {
        f.write_str(self.as_str())
    }
}

/// A scalar cell value.
///
/// `Blank` is deliberately distinct from `Number(0.0)` and `Text("")`: Excel
/// treats an empty cell differently from a cell containing zero in `COUNT`,
/// `ISBLANK`, and lookup functions, even though it coerces to `0` in arithmetic.
#[derive(Debug, Clone, PartialEq)]
pub enum Value {
    Blank,
    Number(f64),
    Text(String),
    Logical(bool),
    Error(ExcelError),
}

impl Value {
    pub fn text(s: impl Into<String>) -> Value {
        Value::Text(s.into())
    }

    pub fn is_error(&self) -> bool {
        matches!(self, Value::Error(_))
    }

    pub fn is_blank(&self) -> bool {
        matches!(self, Value::Blank)
    }

    /// Propagate an error out of a value, if it is one.
    pub fn err(&self) -> Option<ExcelError> {
        match self {
            Value::Error(e) => Some(*e),
            _ => None,
        }
    }

    /// Coerce to a number following Excel's rules.
    ///
    /// - `Blank` → `0`
    /// - `Logical` → `1` / `0`
    /// - `Text` → parsed if it looks numeric, otherwise `#VALUE!`
    ///   (Excel really does evaluate `="5"+1` to `6`.)
    pub fn to_number(&self) -> Result<f64, ExcelError> {
        match self {
            Value::Blank => Ok(0.0),
            Value::Number(n) => Ok(*n),
            Value::Logical(b) => Ok(if *b { 1.0 } else { 0.0 }),
            Value::Text(s) => parse_numeric_text(s).ok_or(ExcelError::Value),
            Value::Error(e) => Err(*e),
        }
    }

    /// Coerce to text following Excel's rules.
    ///
    /// Logicals render as the uppercase words `TRUE` / `FALSE`, which is what
    /// `=A1&""` produces when `A1` holds a boolean.
    pub fn to_text(&self) -> Result<String, ExcelError> {
        match self {
            Value::Blank => Ok(String::new()),
            Value::Number(n) => Ok(number_to_text(*n)),
            Value::Text(s) => Ok(s.clone()),
            Value::Logical(b) => Ok(if *b { "TRUE".into() } else { "FALSE".into() }),
            Value::Error(e) => Err(*e),
        }
    }

    /// Coerce to a boolean following Excel's rules.
    ///
    /// Only the literal words `TRUE` / `FALSE` convert from text; any other
    /// text is `#VALUE!` (unlike the numeric coercion, which parses digits).
    pub fn to_logical(&self) -> Result<bool, ExcelError> {
        match self {
            Value::Blank => Ok(false),
            Value::Number(n) => Ok(*n != 0.0),
            Value::Logical(b) => Ok(*b),
            Value::Text(s) => {
                if s.eq_ignore_ascii_case("TRUE") {
                    Ok(true)
                } else if s.eq_ignore_ascii_case("FALSE") {
                    Ok(false)
                } else {
                    Err(ExcelError::Value)
                }
            }
            Value::Error(e) => Err(*e),
        }
    }
}

impl From<f64> for Value {
    fn from(n: f64) -> Value {
        Value::Number(n)
    }
}

impl From<bool> for Value {
    fn from(b: bool) -> Value {
        Value::Logical(b)
    }
}

impl From<&str> for Value {
    fn from(s: &str) -> Value {
        Value::Text(s.to_string())
    }
}

impl From<ExcelError> for Value {
    fn from(e: ExcelError) -> Value {
        Value::Error(e)
    }
}

impl fmt::Display for Value {
    fn fmt(&self, f: &mut fmt::Formatter<'_>) -> fmt::Result {
        match self {
            Value::Blank => Ok(()),
            Value::Number(n) => f.write_str(&number_to_text(*n)),
            Value::Text(s) => f.write_str(s),
            Value::Logical(b) => f.write_str(if *b { "TRUE" } else { "FALSE" }),
            Value::Error(e) => f.write_str(e.as_str()),
        }
    }
}

/// Parse text that Excel would accept as a number in an arithmetic context.
fn parse_numeric_text(s: &str) -> Option<f64> {
    let t = s.trim();
    if t.is_empty() {
        return None;
    }
    if let Some(stripped) = t.strip_suffix('%') {
        return parse_numeric_text(stripped).map(|n| n / 100.0);
    }
    // Brackets round a number are how an accountant writes a minus sign.
    if let Some(inside) = t.strip_prefix('(').and_then(|held| held.strip_suffix(')')) {
        return parse_numeric_text(inside).map(|n| -n);
    }
    if let Ok(n) = t.parse::<f64>() {
        return Some(n);
    }
    // A currency sign in front, and separators between the thousands. Both are
    // how the number was written down rather than part of it.
    let plain: String = t
        .chars()
        .filter(|held| !matches!(held, '$' | '\u{a5}' | '\u{20ac}' | '\u{a3}' | '\u{ffe5}' | ','))
        .collect();
    if plain != t {
        if let Ok(n) = plain.trim().parse::<f64>() {
            return Some(n);
        }
    }
    // A date or a time is a number too — `="2004-08-15"+1` is the next day.
    crate::datetime::text_as_datetime(t)
}

/// Render a number the way Excel's General format does.
///
/// Excel carries 15 significant decimal digits, not 17. Rendering with Rust's
/// default `{}` would surface the 16th and 17th digits of the binary
/// representation (`0.1 + 0.2` → `0.30000000000000004`), which Excel never
/// shows and which would appear as a spurious divergence against the oracle.
pub fn number_to_text(n: f64) -> String {
    if n == 0.0 {
        // Also normalises -0.0, which Excel displays as "0".
        return "0".to_string();
    }
    if n.is_nan() || n.is_infinite() {
        return ExcelError::Num.as_str().to_string();
    }

    let abs = n.abs();
    if !(1e-4..1e15).contains(&abs) {
        return scientific_to_text(n);
    }

    let exponent = abs.log10().floor() as i32;
    let decimals = (15 - 1 - exponent).clamp(0, 17) as usize;
    let rendered = format!("{:.*}", decimals, n);
    trim_trailing_zeros(&rendered)
}

fn scientific_to_text(n: f64) -> String {
    let formatted = format!("{:E}", n);
    let (mantissa, exponent) = match formatted.split_once('E') {
        Some(parts) => parts,
        None => return formatted,
    };
    let mantissa = trim_trailing_zeros(mantissa);
    let exp: i32 = exponent.parse().unwrap_or(0);
    format!(
        "{}E{}{:02}",
        mantissa,
        if exp < 0 { '-' } else { '+' },
        exp.abs()
    )
}

fn trim_trailing_zeros(s: &str) -> String {
    if !s.contains('.') {
        return s.to_string();
    }
    let trimmed = s.trim_end_matches('0');
    trimmed.strip_suffix('.').unwrap_or(trimmed).to_string()
}

/// Rank used when comparing values of different types.
///
/// Excel does **not** coerce across types when comparing; it orders them by
/// type first. Every number sorts before every text, and every text before
/// every logical, so `=1>"zzz"` is `FALSE` and `="zzz">TRUE` is `FALSE`.
fn type_rank(v: &Value) -> u8 {
    match v {
        Value::Number(_) | Value::Blank => 0,
        Value::Text(_) => 1,
        Value::Logical(_) => 2,
        Value::Error(_) => 3,
    }
}

/// Compare two values using Excel's comparison semantics.
///
/// A `Blank` operand adopts the type of the other side: it compares as `0`
/// against a number, as `""` against text, and as `FALSE` against a logical.
pub fn compare(a: &Value, b: &Value) -> Result<Ordering, ExcelError> {
    if let Some(e) = a.err() {
        return Err(e);
    }
    if let Some(e) = b.err() {
        return Err(e);
    }

    match (a, b) {
        (Value::Blank, Value::Blank) => return Ok(Ordering::Equal),
        (Value::Blank, Value::Text(s)) => return Ok(compare_text("", s)),
        (Value::Text(s), Value::Blank) => return Ok(compare_text(s, "")),
        (Value::Blank, Value::Logical(b)) => return Ok(false.cmp(b)),
        (Value::Logical(a), Value::Blank) => return Ok(a.cmp(&false)),
        _ => {}
    }

    let (ra, rb) = (type_rank(a), type_rank(b));
    if ra != rb {
        return Ok(ra.cmp(&rb));
    }

    match (a, b) {
        (Value::Text(x), Value::Text(y)) => Ok(compare_text(x, y)),
        (Value::Logical(x), Value::Logical(y)) => Ok(x.cmp(y)),
        _ => {
            let x = a.to_number()?;
            let y = b.to_number()?;
            Ok(x.partial_cmp(&y).unwrap_or(Ordering::Equal))
        }
    }
}

/// Excel's text comparison is case-insensitive: `="a"="A"` is `TRUE`.
fn compare_text(a: &str, b: &str) -> Ordering {
    let mut left = a.chars().flat_map(char::to_lowercase);
    let mut right = b.chars().flat_map(char::to_lowercase);
    loop {
        match (left.next(), right.next()) {
            (None, None) => return Ordering::Equal,
            (None, Some(_)) => return Ordering::Less,
            (Some(_), None) => return Ordering::Greater,
            (Some(x), Some(y)) => match x.cmp(&y) {
                Ordering::Equal => continue,
                other => return other,
            },
        }
    }
}

#[cfg(test)]
mod tests {

    /// Every expectation is what Excel 16 gave for `VALUE` of that text, on a
    /// machine that puts the month first (country 81).
    ///
    /// The two day-first dates are the exception and are marked as such: that
    /// Excel refuses them, and the one that wrote the corpus workbook — where
    /// the day comes first — answers as here. One rule covers both: the first
    /// number is the month if it CAN be, and the day if it cannot.
    #[test]
    fn text_that_names_a_moment_or_a_sum_of_money_is_a_number() {
        let read = |text: &str| Value::text(text).to_number();
        for (text, want) in [
            ("2004-08-15", 38214.0),
            ("2004/08/15", 38214.0),
            ("8/15/2004", 38214.0),
            ("15-Aug-2004", 38214.0),
            ("Aug 15, 2004", 38214.0),
            ("15 August 2004", 38214.0),
            ("01/02/2004", 37988.0),
            // Day-first: refused by a month-first Excel, and the only reading
            // there is. This is the corpus workbook's own date.
            ("15/08/2004", 38214.0),
            ("16/01/2009", 39829.0),
        ] {
            assert_eq!(read(text), Ok(want), "{text}");
        }
        for (text, want) in [
            ("12:30", 0.520_833_333_333_333_4),
            ("12:30:45", 0.521_354_166_666_666_7),
            ("1:00 PM", 0.541_666_666_666_666_6),
            ("2004-08-15 12:30", 38_214.520_833_333_336),
        ] {
            assert_eq!(read(text), Ok(want), "{text}");
        }
        for (text, want) in [
            ("1,234.5", 1234.5),
            ("  42  ", 42.0),
            ("42%", 0.42),
            ("-3.5", -3.5),
            ("$100", 100.0),
            ("(5)", -5.0),
            ("1E3", 1000.0),
        ] {
            assert_eq!(read(text), Ok(want), "{text}");
        }
        for text in ["", "not a date", "2004-13-01", "31/02/2004", "25:00", "12:60"] {
            assert!(read(text).is_err(), "{text} is not a number");
        }
    }

    #[test]
    fn a_date_written_out_is_a_number_wherever_it_appears() {
        // Excel keeps this in the coercion rather than in VALUE, so a date in
        // quotes can be added to.
        assert_eq!(
            (Value::text("2004-08-15").to_number().unwrap() + 1.0),
            38215.0,
        );
    }
    use super::*;

    #[test]
    fn blank_is_not_zero_but_coerces_to_zero() {
        assert_ne!(Value::Blank, Value::Number(0.0));
        assert_eq!(Value::Blank.to_number(), Ok(0.0));
        assert_eq!(Value::Blank.to_text(), Ok(String::new()));
    }

    #[test]
    fn numeric_text_coerces_in_arithmetic_context() {
        assert_eq!(Value::text("5").to_number(), Ok(5.0));
        assert_eq!(Value::text(" 5.5 ").to_number(), Ok(5.5));
        assert_eq!(Value::text("50%").to_number(), Ok(0.5));
        assert_eq!(Value::text("abc").to_number(), Err(ExcelError::Value));
    }

    #[test]
    fn only_the_words_true_and_false_coerce_to_logical() {
        assert_eq!(Value::text("TRUE").to_logical(), Ok(true));
        assert_eq!(Value::text("false").to_logical(), Ok(false));
        assert_eq!(Value::text("1").to_logical(), Err(ExcelError::Value));
    }

    #[test]
    fn general_format_carries_fifteen_significant_digits() {
        // Rust's default Display would print 0.30000000000000004 here.
        assert_eq!(number_to_text(0.1 + 0.2), "0.3");
        assert_eq!(number_to_text(1.0), "1");
        assert_eq!(number_to_text(-0.0), "0");
        assert_eq!(number_to_text(1.0 / 3.0), "0.333333333333333");
    }

    #[test]
    fn comparison_ranks_by_type_before_value() {
        // Every number sorts below every text, regardless of magnitude.
        assert_eq!(
            compare(&Value::Number(1e300), &Value::text("a")),
            Ok(Ordering::Less)
        );
        // Every text sorts below every logical.
        assert_eq!(
            compare(&Value::text("zzz"), &Value::Logical(false)),
            Ok(Ordering::Less)
        );
    }

    #[test]
    fn text_comparison_is_case_insensitive() {
        assert_eq!(compare(&Value::text("a"), &Value::text("A")), Ok(Ordering::Equal));
    }

    #[test]
    fn errors_propagate_through_comparison() {
        let err = Value::Error(ExcelError::NA);
        assert_eq!(compare(&err, &Value::Number(1.0)), Err(ExcelError::NA));
    }
}
