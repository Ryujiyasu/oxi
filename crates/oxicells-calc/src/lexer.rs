// This Source Code Form is subject to the terms of the Mozilla Public
// License, v. 2.0. If a copy of the MPL was not distributed with this
// file, You can obtain one at https://mozilla.org/MPL/2.0/.

//! Tokeniser for Excel formula text.
//!
//! Names are lexed without deciding what they are. `SUM`, `A1` and `TAX_RATE`
//! all arrive as [`Token::Name`]; the parser classifies them by looking at what
//! follows and by trying [`crate::reference::parse_a1`]. Deciding in the lexer
//! would misread `LOG10` as a cell reference.

use crate::reference::{parse_a1, CellRef, MAX_COL, MAX_ROW};
use crate::value::ExcelError;
use std::fmt;

#[derive(Debug, Clone, PartialEq)]
pub enum Token {
    Number(f64),
    Text(String),
    ErrorLit(ExcelError),
    /// A bare name, optionally qualified by a sheet: function name, cell
    /// reference, or defined name. Classified during parsing.
    Name {
        sheet: Option<String>,
        name: String,
    },

    Plus,
    Minus,
    Star,
    Slash,
    Caret,
    Percent,
    Amp,

    Eq,
    Ne,
    Lt,
    Le,
    Gt,
    Ge,

    Colon,
    Comma,
    LParen,
    RParen,
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub enum ParseError {
    UnexpectedChar(char, usize),
    UnterminatedString,
    UnterminatedSheetName,
    InvalidNumber(String),
    UnexpectedToken(String),
    UnexpectedEnd,
    TrailingInput(String),
}

impl fmt::Display for ParseError {
    fn fmt(&self, f: &mut fmt::Formatter<'_>) -> fmt::Result {
        match self {
            ParseError::UnexpectedChar(c, at) => write!(f, "unexpected character {c:?} at byte {at}"),
            ParseError::UnterminatedString => f.write_str("unterminated string literal"),
            ParseError::UnterminatedSheetName => f.write_str("unterminated quoted sheet name"),
            ParseError::InvalidNumber(s) => write!(f, "invalid number literal {s:?}"),
            ParseError::UnexpectedToken(s) => write!(f, "unexpected token {s}"),
            ParseError::UnexpectedEnd => f.write_str("unexpected end of formula"),
            ParseError::TrailingInput(s) => write!(f, "trailing input after formula: {s}"),
        }
    }
}

impl std::error::Error for ParseError {}

/// Error literals, longest first so that `#N/A` cannot shadow a longer match.
const ERROR_LITERALS: &[(&str, ExcelError)] = &[
    ("#DIV/0!", ExcelError::DivZero),
    ("#VALUE!", ExcelError::Value),
    ("#NAME?", ExcelError::Name),
    ("#NULL!", ExcelError::Null),
    ("#REF!", ExcelError::Ref),
    ("#NUM!", ExcelError::Num),
    ("#N/A", ExcelError::NA),
];

pub fn tokenize(input: &str) -> Result<Vec<Token>, ParseError> {
    // A leading '=' is how a formula is stored in a cell; accept it either way.
    let src = input.trim();
    let src = src.strip_prefix('=').unwrap_or(src);

    let bytes = src.as_bytes();
    let mut tokens = Vec::new();
    let mut i = 0usize;

    while i < bytes.len() {
        let c = bytes[i] as char;

        if c.is_ascii_whitespace() {
            i += 1;
            continue;
        }

        // Two-character comparison operators must be matched before the
        // one-character forms, or `<>` lexes as `<` followed by `>`.
        if let Some(rest) = src.get(i..) {
            if let Some(op) = rest.strip_prefix("<>").map(|_| Token::Ne) {
                tokens.push(op);
                i += 2;
                continue;
            }
            if rest.starts_with("<=") {
                tokens.push(Token::Le);
                i += 2;
                continue;
            }
            if rest.starts_with(">=") {
                tokens.push(Token::Ge);
                i += 2;
                continue;
            }
        }

        let single = match c {
            '+' => Some(Token::Plus),
            '-' => Some(Token::Minus),
            '*' => Some(Token::Star),
            '/' => Some(Token::Slash),
            '^' => Some(Token::Caret),
            '%' => Some(Token::Percent),
            '&' => Some(Token::Amp),
            '=' => Some(Token::Eq),
            '<' => Some(Token::Lt),
            '>' => Some(Token::Gt),
            ':' => Some(Token::Colon),
            ',' => Some(Token::Comma),
            '(' => Some(Token::LParen),
            ')' => Some(Token::RParen),
            _ => None,
        };
        if let Some(tok) = single {
            tokens.push(tok);
            i += 1;
            continue;
        }

        if c == '"' {
            let (text, next) = lex_string(src, i)?;
            tokens.push(Token::Text(text));
            i = next;
            continue;
        }

        if c == '#' {
            let rest = &src[i..];
            let matched = ERROR_LITERALS
                .iter()
                .find(|(lit, _)| rest.len() >= lit.len() && rest[..lit.len()].eq_ignore_ascii_case(lit));
            match matched {
                Some((lit, err)) => {
                    tokens.push(Token::ErrorLit(*err));
                    i += lit.len();
                    continue;
                }
                None => return Err(ParseError::UnexpectedChar('#', i)),
            }
        }

        if c.is_ascii_digit() || (c == '.' && matches!(bytes.get(i + 1), Some(d) if d.is_ascii_digit())) {
            let (n, next) = lex_number(src, i)?;
            tokens.push(Token::Number(n));
            i = next;
            continue;
        }

        if c == '\'' || c.is_alphabetic() || c == '_' || c == '$' {
            let (tok, next) = lex_name(src, i)?;
            tokens.push(tok);
            i = next;
            continue;
        }

        return Err(ParseError::UnexpectedChar(c, i));
    }

    Ok(tokens)
}

/// Move relative A1 references as Excel does when a formula cell is copied.
///
/// Only formulas understood by this crate are translated. Rejecting an
/// unsupported formula is deliberate: copying it verbatim would silently
/// preserve relative references that Excel would have moved.
pub fn translate_formula_references(
    input: &str,
    row_offset: i64,
    column_offset: i64,
) -> Result<String, String> {
    crate::parser::parse(input).map_err(|error| error.to_string())?;
    let had_equals = input.trim_start().starts_with('=');
    let mut tokens = tokenize(input).map_err(|error| error.to_string())?;
    for index in 0..tokens.len() {
        let is_function = matches!(tokens.get(index + 1), Some(Token::LParen));
        let Token::Name { name, .. } = &mut tokens[index] else {
            continue;
        };
        if is_function {
            continue;
        }
        let Some(mut reference) = parse_a1(name) else {
            continue;
        };
        if !reference.row_absolute {
            reference.row = shifted_coordinate(reference.row, row_offset, MAX_ROW)?;
        }
        if !reference.col_absolute {
            reference.col = shifted_coordinate(reference.col, column_offset, MAX_COL)?;
        }
        *name = reference.to_a1();
    }

    let mut output = String::new();
    if had_equals {
        output.push('=');
    }
    for token in tokens {
        render_token(&mut output, token);
    }
    Ok(output)
}

/// Which way a band of inserted or removed cells runs.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum ShiftAxis {
    Rows,
    Columns,
}

/// Move A1 references across rows or columns put in above them or taken out
/// from under them, the way Excel rewrites formulas after an insert or delete.
///
/// `at` is the one-based first index of the band and `count` how many were put
/// in (positive) or taken out (negative). Unlike a copy, this moves absolute
/// references too: inserting a row above `$A$2` leaves `$A$3`. A reference to
/// something removed becomes `#REF!`, while a range only partly overlapped
/// shrinks, and one an insertion lands inside grows.
pub fn shift_formula_references(
    input: &str,
    axis: ShiftAxis,
    at: u32,
    count: i64,
) -> Result<String, String> {
    crate::parser::parse(input).map_err(|error| error.to_string())?;
    let had_equals = input.trim_start().starts_with('=');
    let tokens = tokenize(input).map_err(|error| error.to_string())?;

    // Callers count rows and columns the way a worksheet does; a CellRef counts
    // from zero.
    let at = at.saturating_sub(1);
    let maximum = match axis {
        ShiftAxis::Rows => MAX_ROW,
        ShiftAxis::Columns => MAX_COL,
    };
    let coordinate = |reference: &CellRef| match axis {
        ShiftAxis::Rows => reference.row,
        ShiftAxis::Columns => reference.col,
    };
    let with_coordinate = |mut reference: CellRef, value: u32| {
        match axis {
            ShiftAxis::Rows => reference.row = value,
            ShiftAxis::Columns => reference.col = value,
        }
        reference
    };

    let mut shifted = Vec::with_capacity(tokens.len());
    let mut index = 0;
    while index < tokens.len() {
        let is_function = matches!(tokens.get(index + 1), Some(Token::LParen));
        let Token::Name { sheet, name } = &tokens[index] else {
            shifted.push(tokens[index].clone());
            index += 1;
            continue;
        };
        if is_function {
            shifted.push(tokens[index].clone());
            index += 1;
            continue;
        }
        let Some(start) = parse_a1(name) else {
            shifted.push(tokens[index].clone());
            index += 1;
            continue;
        };

        // A range moves as a whole, so its ends are decided together.
        let range_end = match (tokens.get(index + 1), tokens.get(index + 2)) {
            (Some(Token::Colon), Some(Token::Name { name, .. })) => parse_a1(name),
            _ => None,
        };
        if let Some(end) = range_end {
            let (low, high) = (coordinate(&start), coordinate(&end));
            match shifted_range(low, high, at, count, maximum)? {
                Some((low, high)) => {
                    shifted.push(Token::Name {
                        sheet: sheet.clone(),
                        name: with_coordinate(start, low).to_a1(),
                    });
                    shifted.push(Token::Colon);
                    let end_sheet = match &tokens[index + 2] {
                        Token::Name { sheet, .. } => sheet.clone(),
                        _ => None,
                    };
                    shifted.push(Token::Name {
                        sheet: end_sheet,
                        name: with_coordinate(end, high).to_a1(),
                    });
                }
                None => shifted.push(Token::ErrorLit(ExcelError::Ref)),
            }
            index += 3;
            continue;
        }

        match shifted_cell(coordinate(&start), at, count, maximum)? {
            Some(value) => shifted.push(Token::Name {
                sheet: sheet.clone(),
                name: with_coordinate(start, value).to_a1(),
            }),
            None => shifted.push(Token::ErrorLit(ExcelError::Ref)),
        }
        index += 1;
    }

    let mut output = String::new();
    if had_equals {
        output.push('=');
    }
    for token in shifted {
        render_token(&mut output, token);
    }
    Ok(output)
}

/// Where one coordinate lands, or `None` once it has been taken out.
fn shifted_cell(value: u32, at: u32, count: i64, maximum: u32) -> Result<Option<u32>, String> {
    if value < at {
        return Ok(Some(value));
    }
    if count >= 0 {
        return shifted_coordinate(value, count, maximum).map(Some);
    }
    let removed = count.unsigned_abs() as u32;
    if value < at.saturating_add(removed) {
        return Ok(None);
    }
    shifted_coordinate(value, count, maximum).map(Some)
}

/// Where a range's ends land. A range wholly taken out answers `None`; one only
/// partly overlapped closes up to the edge of what went, and an insertion
/// landing inside pushes the far end out.
fn shifted_range(
    low: u32,
    high: u32,
    at: u32,
    count: i64,
    maximum: u32,
) -> Result<Option<(u32, u32)>, String> {
    if count >= 0 {
        let low = if low >= at {
            shifted_coordinate(low, count, maximum)?
        } else {
            low
        };
        let high = if high >= at {
            shifted_coordinate(high, count, maximum)?
        } else {
            high
        };
        return Ok(Some((low, high)));
    }
    let removed = count.unsigned_abs() as u32;
    let past = at.saturating_add(removed);
    if low >= at && high < past {
        return Ok(None);
    }
    let low = if low >= past {
        shifted_coordinate(low, count, maximum)?
    } else if low >= at {
        at
    } else {
        low
    };
    let high = if high >= past {
        shifted_coordinate(high, count, maximum)?
    } else if high >= at {
        at.saturating_sub(1)
    } else {
        high
    };
    Ok(Some((low, high)))
}

fn shifted_coordinate(value: u32, offset: i64, maximum: u32) -> Result<u32, String> {
    i64::from(value)
        .checked_add(offset)
        .and_then(|value| u32::try_from(value).ok())
        .filter(|value| *value <= maximum)
        .ok_or_else(|| "copied formula reference moves outside the worksheet".to_string())
}

fn render_token(output: &mut String, token: Token) {
    match token {
        Token::Number(value) => output.push_str(&value.to_string()),
        Token::Text(value) => {
            output.push('"');
            output.push_str(&value.replace('"', "\"\""));
            output.push('"');
        }
        Token::ErrorLit(value) => output.push_str(value.as_str()),
        Token::Name { sheet, name } => {
            if let Some(sheet) = sheet {
                if sheet_needs_quotes(&sheet) {
                    output.push('\'');
                    output.push_str(&sheet.replace('\'', "''"));
                    output.push('\'');
                } else {
                    output.push_str(&sheet);
                }
                output.push('!');
            }
            output.push_str(&name);
        }
        Token::Plus => output.push('+'),
        Token::Minus => output.push('-'),
        Token::Star => output.push('*'),
        Token::Slash => output.push('/'),
        Token::Caret => output.push('^'),
        Token::Percent => output.push('%'),
        Token::Amp => output.push('&'),
        Token::Eq => output.push('='),
        Token::Ne => output.push_str("<>"),
        Token::Lt => output.push('<'),
        Token::Le => output.push_str("<="),
        Token::Gt => output.push('>'),
        Token::Ge => output.push_str(">="),
        Token::Colon => output.push(':'),
        Token::Comma => output.push(','),
        Token::LParen => output.push('('),
        Token::RParen => output.push(')'),
    }
}

fn sheet_needs_quotes(name: &str) -> bool {
    name.is_empty()
        || name
            .chars()
            .any(|character| !(character.is_alphanumeric() || character == '_' || character == '.'))
        || name
            .chars()
            .next()
            .is_some_and(|character| character.is_ascii_digit())
}

fn lex_string(src: &str, start: usize) -> Result<(String, usize), ParseError> {
    let bytes = src.as_bytes();
    let mut i = start + 1;
    let mut out = String::new();
    while i < bytes.len() {
        if bytes[i] == b'"' {
            // A doubled quote is an escaped quote, not the end of the literal.
            if bytes.get(i + 1) == Some(&b'"') {
                out.push('"');
                i += 2;
                continue;
            }
            return Ok((out, i + 1));
        }
        let ch = src[i..].chars().next().expect("in bounds");
        out.push(ch);
        i += ch.len_utf8();
    }
    Err(ParseError::UnterminatedString)
}

fn lex_number(src: &str, start: usize) -> Result<(f64, usize), ParseError> {
    let bytes = src.as_bytes();
    let mut i = start;

    while i < bytes.len() && bytes[i].is_ascii_digit() {
        i += 1;
    }
    if bytes.get(i) == Some(&b'.') {
        i += 1;
        while i < bytes.len() && bytes[i].is_ascii_digit() {
            i += 1;
        }
    }
    // An exponent only counts when digits actually follow, so that `1E` stays
    // a number followed by a name rather than a malformed literal.
    if matches!(bytes.get(i), Some(b'e') | Some(b'E')) {
        let mut j = i + 1;
        if matches!(bytes.get(j), Some(b'+') | Some(b'-')) {
            j += 1;
        }
        if matches!(bytes.get(j), Some(d) if d.is_ascii_digit()) {
            j += 1;
            while j < bytes.len() && bytes[j].is_ascii_digit() {
                j += 1;
            }
            i = j;
        }
    }

    let text = &src[start..i];
    text.parse::<f64>()
        .map(|n| (n, i))
        .map_err(|_| ParseError::InvalidNumber(text.to_string()))
}

fn lex_name(src: &str, start: usize) -> Result<(Token, usize), ParseError> {
    let mut i = start;
    let mut sheet = None;

    // Quoted sheet prefix: 'My Sheet'!  — an embedded quote is doubled.
    if src.as_bytes()[i] == b'\'' {
        let (name, next) = lex_quoted_sheet(src, i)?;
        if src.as_bytes().get(next) != Some(&b'!') {
            return Err(ParseError::UnterminatedSheetName);
        }
        sheet = Some(name);
        i = next + 1;
    } else {
        // Bare sheet prefix: Sheet1!
        let end = scan_sheet_word(src, i);
        if src.as_bytes().get(end) == Some(&b'!') && end > i {
            sheet = Some(src[i..end].to_string());
            i = end + 1;
        }
    }

    // A reference can be sheet-qualified and still broken: `'Sheet'!#REF!` is
    // what Excel writes after the target is deleted. The sheet no longer means
    // anything, so it collapses to the error value.
    if src.as_bytes().get(i) == Some(&b'#') {
        let rest = &src[i..];
        if let Some((lit, err)) = ERROR_LITERALS
            .iter()
            .find(|(lit, _)| rest.len() >= lit.len() && rest[..lit.len()].eq_ignore_ascii_case(lit))
        {
            return Ok((Token::ErrorLit(*err), i + lit.len()));
        }
    }

    let end = scan_word(src, i);
    if end == i {
        return Err(ParseError::UnexpectedChar(
            src[i..].chars().next().unwrap_or('!'),
            i,
        ));
    }

    Ok((
        Token::Name {
            sheet,
            name: src[i..end].to_string(),
        },
        end,
    ))
}

fn lex_quoted_sheet(src: &str, start: usize) -> Result<(String, usize), ParseError> {
    let bytes = src.as_bytes();
    let mut i = start + 1;
    let mut out = String::new();
    while i < bytes.len() {
        if bytes[i] == b'\'' {
            if bytes.get(i + 1) == Some(&b'\'') {
                out.push('\'');
                i += 2;
                continue;
            }
            return Ok((out, i + 1));
        }
        let ch = src[i..].chars().next().expect("in bounds");
        out.push(ch);
        i += ch.len_utf8();
    }
    Err(ParseError::UnterminatedSheetName)
}

/// Scan the run of characters that can make up a name or an A1 reference.
fn scan_word(src: &str, start: usize) -> usize {
    let mut i = start;
    for ch in src[start..].chars() {
        if ch.is_alphanumeric() || ch == '_' || ch == '.' || ch == '$' {
            i += ch.len_utf8();
        } else {
            break;
        }
    }
    i
}

/// Scan an *unquoted* sheet name, which is far more permissive than a defined
/// name.
///
/// Excel does not quote a sheet name unless it has to, and its rules for "has
/// to" do not cover characters like the katakana middle dot `・`. Real files
/// contain `前月比・前年同月比計算!AD7` written bare, so anything that is not a
/// formula operator or delimiter has to be accepted here.
fn scan_sheet_word(src: &str, start: usize) -> usize {
    const DELIMITERS: &[char] = &[
        '!', '\'', '"', '(', ')', '[', ']', ':', ',', ';', '+', '-', '*', '/', '^', '&', '<', '>',
        '=', '%', '{', '}',
    ];
    let mut i = start;
    for ch in src[start..].chars() {
        if ch.is_whitespace() || DELIMITERS.contains(&ch) {
            break;
        }
        i += ch.len_utf8();
    }
    i
}

#[cfg(test)]
mod tests {
    use super::*;

    fn lex(s: &str) -> Vec<Token> {
        tokenize(s).expect("should lex")
    }

    #[test]
    fn translates_relative_and_absolute_formula_references() {
        assert_eq!(
            translate_formula_references(
                "=A1+$B2+C$3+$D$4+SUM(E5:F6)+\"A1\"+'Data Sheet'!G7",
                2,
                1,
            )
            .unwrap(),
            "=B3+$B4+D$3+$D$4+SUM(F7:G8)+\"A1\"+'Data Sheet'!H9"
        );
    }

    #[test]
    fn formula_translation_rejects_out_of_bounds_and_unsupported_formulas() {
        assert!(translate_formula_references("=A1", -1, 0)
            .unwrap_err()
            .contains("outside the worksheet"));
        assert!(translate_formula_references("=A1;B1", 1, 1).is_err());
    }

    #[test]
    fn formula_translation_does_not_treat_function_names_as_cells() {
        assert_eq!(
            translate_formula_references("LOG10(A1)", 1, 1).unwrap(),
            "LOG10(B2)"
        );
    }

    fn name(n: &str) -> Token {
        Token::Name {
            sheet: None,
            name: n.to_string(),
        }
    }

    #[test]
    fn leading_equals_is_optional() {
        assert_eq!(lex("=1+1"), lex("1+1"));
    }

    #[test]
    fn two_character_operators_win_over_one() {
        assert_eq!(lex("1<>2"), vec![Token::Number(1.0), Token::Ne, Token::Number(2.0)]);
        assert_eq!(lex("1<=2"), vec![Token::Number(1.0), Token::Le, Token::Number(2.0)]);
        assert_eq!(lex("1>=2"), vec![Token::Number(1.0), Token::Ge, Token::Number(2.0)]);
    }

    #[test]
    fn names_are_not_classified_during_lexing() {
        // LOG10 must not be split into a name and a number, and A1 must not be
        // resolved to a reference yet.
        assert_eq!(lex("LOG10(A1)"), vec![name("LOG10"), Token::LParen, name("A1"), Token::RParen]);
    }

    #[test]
    fn dollar_signs_stay_attached_to_the_reference() {
        assert_eq!(lex("$A$1"), vec![name("$A$1")]);
    }

    #[test]
    fn sheet_prefixes_are_captured() {
        assert_eq!(
            lex("Sheet1!A1"),
            vec![Token::Name {
                sheet: Some("Sheet1".into()),
                name: "A1".into()
            }]
        );
        assert_eq!(
            lex("'My Sheet'!A1"),
            vec![Token::Name {
                sheet: Some("My Sheet".into()),
                name: "A1".into()
            }]
        );
    }

    #[test]
    fn doubled_quotes_are_escapes() {
        assert_eq!(lex(r#""a""b""#), vec![Token::Text(r#"a"b"#.to_string())]);
        assert_eq!(tokenize(r#""oops"#), Err(ParseError::UnterminatedString));
    }

    #[test]
    fn exponents_need_digits_to_count() {
        assert_eq!(lex("1E3"), vec![Token::Number(1000.0)]);
        assert_eq!(lex("1E-3"), vec![Token::Number(0.001)]);
        // No digits after E: the number ends at `1` and `E3` is a name.
        assert_eq!(lex("1E"), vec![Token::Number(1.0), name("E")]);
    }

    #[test]
    fn error_literals_lex_as_values() {
        assert_eq!(lex("#DIV/0!"), vec![Token::ErrorLit(ExcelError::DivZero)]);
        assert_eq!(lex("#N/A"), vec![Token::ErrorLit(ExcelError::NA)]);
        assert_eq!(lex("#REF!"), vec![Token::ErrorLit(ExcelError::Ref)]);
    }

    #[test]
    fn japanese_text_and_names_survive() {
        assert_eq!(lex(r#""単価""#), vec![Token::Text("単価".to_string())]);
        assert_eq!(lex("税率"), vec![name("税率")]);
    }
}

#[cfg(test)]
mod shift_tests {
    use super::{shift_formula_references, ShiftAxis};

    /// Every case here is what Excel 16 left in the cell after the operation.
    #[test]
    fn references_follow_inserted_and_removed_rows() {
        for (formula, at, count, expected) in [
            // A row put in above a reference pushes it down, absolute or not.
            ("=$A$2*2", 1, 1, "=$A$3*2"),
            ("=A$2+$A3", 1, 1, "=A$3+$A4"),
            ("=A1*2", 3, 1, "=A1*2"),
            ("=A2*2", 2, 2, "=A4*2"),
            // Taking rows out pulls what is below them up.
            ("=$A$4*2", 1, -1, "=$A$3*2"),
            ("=A4*2", 2, -2, "=A2*2"),
            // A reference to a row that went becomes #REF!.
            ("=A2*2", 2, -1, "=#REF!*2"),
        ] {
            assert_eq!(
                shift_formula_references(formula, ShiftAxis::Rows, at, count).unwrap(),
                expected,
                "{formula} with {count} at row {at}"
            );
        }
    }

    #[test]
    fn a_range_grows_and_shrinks_around_the_change() {
        for (formula, at, count, expected) in [
            // An insertion inside a range stretches it.
            ("=SUM(A1:A3)", 2, 1, "=SUM(A1:A4)"),
            // A range straddling what went closes up.
            ("=SUM(A1:A3)", 2, -1, "=SUM(A1:A2)"),
            ("=SUM(A1:A3)", 1, -1, "=SUM(A1:A2)"),
            // One wholly inside what went is left with nothing to point at.
            ("=SUM(A2:A2)", 2, -1, "=SUM(#REF!)"),
            ("=SUM(A2:A3)", 2, -2, "=SUM(#REF!)"),
        ] {
            assert_eq!(
                shift_formula_references(formula, ShiftAxis::Rows, at, count).unwrap(),
                expected,
                "{formula} with {count} at row {at}"
            );
        }
    }

    #[test]
    fn columns_move_the_same_way_rows_do() {
        for (formula, at, count, expected) in [
            ("=B1*2", 1, 1, "=C1*2"),
            ("=B1*2", 1, -1, "=A1*2"),
            ("=B1*2", 2, -1, "=#REF!*2"),
            ("=SUM(A1:C1)", 2, -1, "=SUM(A1:B1)"),
        ] {
            assert_eq!(
                shift_formula_references(formula, ShiftAxis::Columns, at, count).unwrap(),
                expected,
                "{formula} with {count} at column {at}"
            );
        }
    }

    /// A function's name is not a reference, however much it reads like one.
    #[test]
    fn a_function_name_is_left_alone() {
        assert_eq!(
            shift_formula_references("=LOG10(A2)", ShiftAxis::Rows, 1, 1).unwrap(),
            "=LOG10(A3)"
        );
    }
}
