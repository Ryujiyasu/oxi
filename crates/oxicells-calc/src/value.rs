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
    /// `#SPILL!` — a dynamic array's answer has nowhere to spill.
    Spill,
    /// `#CALC!` — a LAMBDA left uncalled, among other things.
    Calc,
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
            ExcelError::Spill => "#SPILL!",
            ExcelError::Calc => "#CALC!",
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
    if plain != t && thousands_well_placed(t) {
        if let Ok(n) = plain.trim().parse::<f64>() {
            return Some(n);
        }
    }
    // Digits and commas that fail as thousands are no date either: measured,
    // VALUE("1,23,456") is #VALUE!.
    if t.contains(',') && t.chars().all(|held| held.is_ascii_digit() || matches!(held, ',' | '.' | '+' | '-' | ' ')) {
        return None;
    }
    // A date or a time is a number too — `="2004-08-15"+1` is the next day.
    crate::datetime::text_as_datetime(t)
}

/// Whether the text's thousands separators stand where thousands are: every
/// group after a comma carries three digits or more before any point, and
/// none is empty. Measured: VALUE of "1,234", "1,2345" and "12,345,678"
/// reads, of "1,2", "12,34", "1,234,5", "1,23,456" and ",123" is #VALUE!.
fn thousands_well_placed(text: &str) -> bool {
    let digits: String = text
        .chars()
        .filter(|held| !matches!(held, '$' | '\u{a5}' | '\u{20ac}' | '\u{a3}' | '\u{ffe5}' | '+' | '-' | ' '))
        .collect();
    if !digits.contains(',') {
        return true;
    }
    let groups: Vec<&str> = digits.split(',').collect();
    groups.iter().enumerate().all(|(at, group)| {
        let whole = group.split('.').next().unwrap_or(group);
        !whole.is_empty() && (at == 0 || whole.len() >= 3) && (at + 1 == groups.len() || !group.contains('.'))
    })
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

    // Fifteen significant figures, written out where that takes twenty
    // characters or fewer and in exponent form where it would take more.
    // Measured through ="" & x: 1E+19 is 10000000000000000000 and 1E+20
    // 1E+20; 123456789012345678 is 123456789012345000; 0.000000001 and
    // -0.00001234567890123 are written out, 1/3*1E-5 is 3.33333333333333E-06.
    let formatted = format!("{:.14e}", n.abs());
    let (mantissa, exponent) = formatted.split_once('e').unwrap_or((&formatted, "0"));
    let exponent: i32 = exponent.parse().unwrap_or(0);
    let digits: String = mantissa.replace('.', "");
    let digits = digits.trim_end_matches('0');
    let digits = if digits.is_empty() { "0" } else { digits };
    let written = if exponent >= 0 {
        let whole_len = exponent as usize + 1;
        if digits.len() <= whole_len {
            format!("{digits}{}", "0".repeat(whole_len - digits.len()))
        } else {
            format!("{}.{}", &digits[..whole_len], &digits[whole_len..])
        }
    } else {
        format!("0.{}{digits}", "0".repeat((-exponent - 1) as usize))
    };
    let sign = if n < 0.0 { "-" } else { "" };
    if written.len() <= 20 {
        return format!("{sign}{written}");
    }
    scientific_to_text(n)
}

fn scientific_to_text(n: f64) -> String {
    // Fifteen significant digits, as everywhere else: measured,
    // TIMEVALUE("0:0:0.5")&"" is 5.78703703703704E-06 -- fourteen where
    // the exponent reaches 99: 1/3*1E100 is 3.3333333333333E+99.
    let exponent = n.abs().log10().floor() as i32;
    let formatted = if exponent.abs() >= 99 { format!("{:.13E}", n) } else { format!("{:.14E}", n) };
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

/// Excel's text comparison, a collation rather than a comparison of codes.
/// Case never counts: `="a"="A"` is TRUE. Past that, level by level:
///
/// 1. the letters themselves, with ligatures spelled out, width, kana kind
///    and accents set aside, and hyphens and apostrophes passed over --
///    measured, `="ß"="ss"` is TRUE, `="é"<"f"` TRUE, `="a-b"<"ab"` FALSE,
///    and `="Ⅰ"<"H"` TRUE (Roman numerals before the letters);
/// 2. accents: `="é"="e"` is FALSE;
/// 3. width and kana kind: `="A"="ａ"` and `="ア"="ｱ"` are FALSE, and SORT
///    puts Ａ before a, and ァ, ア, ｱ, あ in that order;
/// 4. the hyphens and apostrophes: SORT puts coop before co-op.
fn compare_text(a: &str, b: &str) -> Ordering {
    if a.is_ascii() && b.is_ascii() && !a.contains(['-', '\'']) && !b.contains(['-', '\'']) {
        let left = a.bytes().map(|byte| byte.to_ascii_lowercase());
        let right = b.bytes().map(|byte| byte.to_ascii_lowercase());
        return left.cmp(right);
    }
    let (left, right) = (collation_units(a), collation_units(b));
    let passed = |unit: &CollationUnit| unit.passed;
    let level = |units: &[CollationUnit], pick: fn(&CollationUnit) -> char| -> Vec<char> {
        units.iter().filter(|unit| !passed(unit)).map(pick).collect()
    };
    level(&left, |unit| unit.base)
        .cmp(&level(&right, |unit| unit.base))
        .then_with(|| level(&left, |unit| unit.accented).cmp(&level(&right, |unit| unit.accented)))
        .then_with(|| {
            let kinds = |units: &[CollationUnit]| units.iter().filter(|unit| !unit.passed).map(|unit| unit.kind).collect::<Vec<u8>>();
            kinds(&left).cmp(&kinds(&right))
        })
        .then_with(|| left.iter().filter(|unit| unit.passed).count().cmp(&right.iter().filter(|unit| unit.passed).count()))
}

/// One letter as the collation sees it: its base, its base with any accent
/// kept, and its kind (width, kana script, size) for the third level.
struct CollationUnit {
    base: char,
    accented: char,
    kind: u8,
    /// The kind with width set aside: a full-width letter, a half-width
    /// kana and a ringed digit are their plain selves here.
    widthless: u8,
    passed: bool,
}

/// Two values the same when width is set aside as well as case: UNIQUE,
/// XMATCH and XLOOKUP ask this. Measured: UNIQUE merges "a", "A" and "ａ",
/// "ア" and "ｱ", "①" and "1", "ß" and "ss", while "あ", "ァ", "é", "a-b" and
/// "Ⅰ" each stand apart; XMATCH("ａ",{"a","ａ"}) is 1.
pub fn same_ignoring_width(a: &Value, b: &Value) -> bool {
    match (a, b) {
        (Value::Text(x), Value::Text(y)) => {
            let (left, right) = (collation_units(x), collation_units(y));
            let key = |units: &[CollationUnit]| -> Vec<(char, char, u8, bool)> {
                units.iter().map(|unit| (unit.base, unit.accented, unit.widthless, unit.passed)).collect()
            };
            key(&left) == key(&right)
        }
        (Value::Error(x), Value::Error(y)) => x == y,
        _ => compare(a, b) == Ok(Ordering::Equal),
    }
}

fn collation_units(text: &str) -> Vec<CollationUnit> {
    let mut units = Vec::with_capacity(text.len());
    for character in text.chars().flat_map(char::to_lowercase) {
        let push = |units: &mut Vec<CollationUnit>, base: char, accented: char, kind: u8| {
            units.push(CollationUnit { base, accented, kind, widthless: kind, passed: false });
        };
        match character {
            '-' | '\'' => units.push(CollationUnit { base: character, accented: character, kind: 1, widthless: 1, passed: true }),
            'ß' => {
                push(&mut units, 's', 's', 1);
                push(&mut units, 's', 's', 1);
            }
            'æ' | 'œ' | '\u{FB01}' | '\u{FB02}' | '\u{0133}' => {
                let (first, second) = match character {
                    'æ' => ('a', 'e'),
                    'œ' => ('o', 'e'),
                    '\u{FB01}' => ('f', 'i'),
                    '\u{FB02}' => ('f', 'l'),
                    _ => ('i', 'j'),
                };
                push(&mut units, first, first, 1);
                push(&mut units, second, second, 1);
            }
            // Roman numerals come before every letter.
            '\u{2170}'..='\u{217F}' => {
                let low = char::from_u32(character as u32 - 0x2170 + 1).unwrap_or(character);
                push(&mut units, low, low, 1);
            }
            // Full-width Latin is the Latin letter, dressed wider.
            '\u{FF01}'..='\u{FF5E}' => {
                let narrow = char::from_u32(character as u32 - 0xFEE0).unwrap_or(character);
                push(&mut units, narrow, narrow, 0);
                if let Some(last) = units.last_mut() {
                    last.widthless = 1;
                }
            }
            // A ringed digit is its digit, drawn differently.
            '\u{2460}'..='\u{2468}' => {
                let digit = char::from_u32(character as u32 - 0x2460 + u32::from(b'1')).unwrap_or(character);
                push(&mut units, digit, digit, 4);
                if let Some(last) = units.last_mut() {
                    last.widthless = 1;
                }
            }
            _ => {
                let (katakana, kind) = kana_collation(character);
                let base = accent_base(katakana);
                push(&mut units, base, katakana, kind);
                if kind == 2 {
                    if let Some(last) = units.last_mut() {
                        last.widthless = 1;
                    }
                }
            }
        }
    }
    units
}

/// A kana as full-size katakana, with its kind: 0 small, 1 katakana,
/// 2 half-width, 3 hiragana. Anything else is itself, of kind 1.
fn kana_collation(character: char) -> (char, u8) {
    let code = character as u32;
    let (code, kind) = match code {
        0x3041..=0x3096 => (code + 0x60, 3),
        0xFF66..=0xFF9D => (HALF_WIDTH_KATAKANA[(code - 0xFF66) as usize] as u32, 2),
        _ => (code, 1),
    };
    // The small kana stand beside their full-size letter.
    let (code, kind) = match code {
        0x30A1 | 0x30A3 | 0x30A5 | 0x30A7 | 0x30A9 | 0x30C3 | 0x30E3 | 0x30E5 | 0x30E7 | 0x30EE => {
            (code + 1, if kind == 1 { 0 } else { kind })
        }
        0x30F5 => (0x30AB, 0),
        0x30F6 => (0x30B1, 0),
        _ => (code, kind),
    };
    (char::from_u32(code).unwrap_or(character), kind)
}

const HALF_WIDTH_KATAKANA: [char; 56] = [
    'ヲ', 'ァ', 'ィ', 'ゥ', 'ェ', 'ォ', 'ャ', 'ュ', 'ョ', 'ッ', 'ー', 'ア', 'イ', 'ウ', 'エ', 'オ', 'カ', 'キ', 'ク', 'ケ', 'コ',
    'サ', 'シ', 'ス', 'セ', 'ソ', 'タ', 'チ', 'ツ', 'テ', 'ト', 'ナ', 'ニ', 'ヌ', 'ネ', 'ノ', 'ハ', 'ヒ', 'フ', 'ヘ', 'ホ', 'マ',
    'ミ', 'ム', 'メ', 'モ', 'ヤ', 'ユ', 'ヨ', 'ラ', 'リ', 'ル', 'レ', 'ロ', 'ワ', 'ン',
];

/// A Latin letter with its accent taken off.
fn accent_base(character: char) -> char {
    match character {
        'à'..='å' => 'a',
        'ç' => 'c',
        'è'..='ë' => 'e',
        'ì'..='ï' => 'i',
        'ñ' => 'n',
        'ò'..='ö' | 'ø' => 'o',
        'ù'..='ü' => 'u',
        'ý' | 'ÿ' => 'y',
        other => other,
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
            // Measured: `"25:00"+0` is a day and an hour, and "10 PM" is 22:00.
            ("25:00", 25.0 / 24.0),
            ("10 PM", 22.0 / 24.0),
            ("12:60", 13.0 / 24.0),
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
        for text in ["", "not a date", "2004-13-01", "31/02/2004", "10000:00", "0:10000", "1:60 PM", "10PM", "13 PM"] {
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
