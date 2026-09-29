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
    /// A structured reference: a table's name and whatever was asked of it,
    /// as the raw text between the brackets.
    ///
    /// Kept whole because what is inside those brackets is not the ordinary
    /// language: `[#This Row]` has a space in it and `[[A]:[B]]` uses a colon
    /// that means columns rather than cells. Letting either through to the
    /// ordinary parser would make quite different sense of them.
    Table {
        name: String,
        asked: String,
    },

    Plus,
    Minus,
    Star,
    Slash,
    Caret,
    Percent,
    Amp,
    /// `@`, implicit intersection: `=@A1:A3` in row 2 is A2.
    At,
    /// `#` after a cell: the whole of what that cell's formula spilled,
    /// `=SUM(D1#)`.
    Hash,
    /// A space between two references: the cells both share,
    /// `=SUM(A1:C3 B2:C3)`.
    Intersect,

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
    /// The braces an array constant is written between, and the `;` that
    /// separates its rows -- `{1,2;3,4}` is two rows of two.
    LBrace,
    RBrace,
    Semicolon,
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub enum ParseError {
    UnexpectedChar(char, usize),
    UnterminatedString,
    UnterminatedSheetName,
    /// A sheet living in another workbook: `[1]Sales!A1`, or quoted as
    /// `'[1]May 2021'!A1`. There is nothing here to resolve it against, and
    /// answering `#REF!` would throw away the value the file was saved with.
    AnotherWorkbook(String),
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
            ParseError::AnotherWorkbook(name) => {
                write!(f, "sheet {name:?} is in a workbook this one only links to")
            }
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
    ("#SPILL!", ExcelError::Spill),
    ("#CALC!", ExcelError::Calc),
];

pub fn tokenize(input: &str) -> Result<Vec<Token>, ParseError> {
    tokenize_spanned(input).map(|spanned| spanned.into_iter().map(|(token, _)| token).collect())
}

/// The tokens, each with the byte offset in `input` it starts at.
fn tokenize_spanned(input: &str) -> Result<Vec<(Token, usize)>, ParseError> {
    // A leading '=' is how a formula is stored in a cell; accept it either way.
    let src = input.trim();
    let base = input.len() - input.trim_start().len() + usize::from(src.starts_with('='));
    let src = src.strip_prefix('=').unwrap_or(src);

    let bytes = src.as_bytes();
    let mut tokens = Vec::new();
    let mut starts: Vec<usize> = Vec::new();
    let mut i = 0usize;
    let mut last_start = 0usize;

    while i < bytes.len() {
        while starts.len() < tokens.len() {
            starts.push(base + last_start);
        }
        last_start = i;
        let c = bytes[i] as char;

        if c.is_ascii_whitespace() {
            // A space with a reference on either side is the intersection
            // operator; anywhere else it is only a space.
            let ends_a_reference = matches!(
                tokens.last(),
                Some(Token::Name { .. } | Token::RParen | Token::Table { .. } | Token::Hash)
            );
            let mut next = i;
            while next < bytes.len() && (bytes[next] as char).is_ascii_whitespace() {
                next += 1;
            }
            let starts_a_reference = next < bytes.len()
                && matches!(bytes[next] as char, 'A'..='Z' | 'a'..='z' | '$' | '\'' | '(' | '_');
            if ends_a_reference && starts_a_reference && !matches!(tokens.last(), Some(Token::Name { name, .. }) if src[next..].starts_with('(') && name.is_empty()) {
                tokens.push(Token::Intersect);
            }
            i = next;
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
            '@' => Some(Token::At),
            '=' => Some(Token::Eq),
            '<' => Some(Token::Lt),
            '>' => Some(Token::Gt),
            ':' => Some(Token::Colon),
            ',' => Some(Token::Comma),
            '(' => Some(Token::LParen),
            ')' => Some(Token::RParen),
            '{' => Some(Token::LBrace),
            '}' => Some(Token::RBrace),
            ';' => Some(Token::Semicolon),
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
                // `D1#`: the spill of the formula in D1.
                None if matches!(tokens.last(), Some(Token::Name { .. })) => {
                    tokens.push(Token::Hash);
                    i += 1;
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

        // `[1]Assistente!R6` starts with the link number rather than with a
        // letter, so a bracketed number in front of a sheet name is a name
        // start too. A bracket that is not that is still nothing here.
        let starts_a_link = c == '['
            && src[i + 1..].split_once(']').is_some_and(|(digits, _)| {
                !digits.is_empty() && digits.bytes().all(|b| b.is_ascii_digit())
            });
        if c == '\'' || c.is_alphabetic() || c == '_' || c == '$' || starts_a_link {
            // A name followed by `[` is a table being asked for one of its
            // columns, and the whole bracket group belongs to it.
            if let Some((tok, next)) = lex_table(src, i) {
                tokens.push(tok);
                i = next;
                continue;
            }
            let (tok, next) = lex_name(src, i)?;
            tokens.push(tok);
            i = next;
            continue;
        }

        // A bracket group on its own, `[@Qty]`, names a column of the table
        // the formula sits in.
        if c == '[' {
            if let Some(Token::Table { asked, .. }) = lex_table(&format!("_{}", &src[i..]), 0).map(|(tok, _)| tok) {
                let next = i + asked.len() + 2;
                tokens.push(Token::Table { name: String::new(), asked });
                i = next;
                continue;
            }
        }

        return Err(ParseError::UnexpectedChar(c, i));
    }
    while starts.len() < tokens.len() {
        starts.push(base + last_start);
    }

    Ok(tokens.into_iter().zip(starts).collect())
}

/// A formula written through `.Formula` as Excel keeps it: with `@` wherever
/// a legacy formula would have taken one cell out of a range. `multi_name`
/// says whether a defined name stands for more than one cell.
///
/// Measured through `.Formula2` after `.Formula`: a range, or a name for one,
/// where one value is wanted takes `@` (`=@A1:A3`, `=LEN(@A1:A3)`,
/// `=@A1:A3+1`, `=IF(@A1:A3>1,1,0)`); a function that may answer with a
/// block takes it in front (`=@INDEX(A1:A3,0)`, `=@OFFSET(A1,0,0,2,1)`,
/// `=@INDIRECT("A1")`); a parameter that takes a range leaves it be
/// (`=SUM(A1:A3)`, `=VLOOKUP(@A1:A3,A1:B3,2,0)`), though an expression in it
/// is still worked one value at a time (`=SUM(@A1:A3*2)`); and an array
/// parameter leaves everything in it be (`=SUMPRODUCT((A1:A3>1)*B1:B3)`).
/// A function not in the tables is left untouched, as is a formula that
/// will not read.
pub fn implied_intersections(input: &str, multi_name: &dyn Fn(&str) -> bool) -> String {
    let Ok(tokens) = tokenize_spanned(input) else {
        return input.to_string();
    };
    let mut reader = AtReader { tokens: &tokens, pos: 0, multi_name };
    let Some(tree) = reader.expr() else {
        return input.to_string();
    };
    if reader.pos != tokens.len() {
        return input.to_string();
    }
    let mut marks = Vec::new();
    mark_intersections(&tree, AtContext::Value, &mut marks);
    if marks.is_empty() {
        return input.to_string();
    }
    marks.sort_unstable();
    marks.dedup();
    let mut output = String::with_capacity(input.len() + marks.len());
    let mut from = 0;
    for at in marks {
        output.push_str(&input[from..at]);
        output.push('@');
        from = at;
    }
    output.push_str(&input[from..]);
    output
}

/// A formula as Excel keeps it once written: cell references, function
/// names, TRUE and FALSE upper-cased; a sheet and a defined name spelt as
/// they were defined (`sheet_case` and `name_case` say how, or None where
/// there is no such thing); numbers written plainly; and no space before a
/// comma. Measured through `.Formula`: `=sum( a1 , b1 )` reads back
/// `=SUM( A1, B1 )`, `=sheet1!a1` `=Sheet1!A1`, `=myname` `=MyName`,
/// `=1.50+0.0` `=1.5+0`, `=1e3` `=1000`, `=.5` `=0.5`, `=#n/a` `=#N/A`.
/// Other spacing is kept. A formula that will not tokenize comes back as it
/// was.
pub fn canonical_formula(
    input: &str,
    sheet_case: &dyn Fn(&str) -> Option<String>,
    name_case: &dyn Fn(&str) -> Option<String>,
) -> String {
    let Ok(tokens) = tokenize_spanned(input) else {
        return input.to_string();
    };
    if tokens.is_empty() {
        return input.to_string();
    }
    let mut output = String::with_capacity(input.len());
    output.push_str(&input[..tokens[0].1]);
    for (index, (token, start)) in tokens.iter().enumerate() {
        let next = tokens.get(index + 1).map_or(input.len(), |(_, at)| *at);
        let written = &input[*start..next];
        let text = written.trim_end();
        let gap = &written[text.len()..];
        let calls = matches!(tokens.get(index + 1), Some((Token::LParen, _)));
        let beside_colon = matches!(tokens.get(index + 1), Some((Token::Colon, _)))
            || (index > 0 && matches!(tokens.get(index - 1), Some((Token::Colon, _))));
        // A sign written straight onto a nought is dropped: measured, `=-0`
        // reads back `=0`, `=1--0` `=1-0`, `=--0` `=-0` and `=+0` `=0`, while
        // `=- 0`, `=-5` and `=+A1` keep theirs.
        if matches!(token, Token::Minus | Token::Plus)
            && gap.is_empty()
            && matches!(tokens.get(index + 1), Some((Token::Number(value), _)) if *value == 0.0)
            && !matches!(
                index.checked_sub(1).and_then(|before| tokens.get(before)),
                Some((
                    Token::Number(_)
                        | Token::Name { .. }
                        | Token::Text(_)
                        | Token::ErrorLit(_)
                        | Token::RParen
                        | Token::Percent
                        | Token::Table { .. },
                    _
                ))
            )
        {
            continue;
        }
        match token {
            // The first sheet of `Qa:Qc!A1` is a sheet, spelt as it was named.
            Token::Name { sheet: None, name }
                if matches!(tokens.get(index + 1), Some((Token::Colon, _)))
                    && matches!(tokens.get(index + 2), Some((Token::Name { sheet: Some(_), .. }, _))) =>
            {
                let name = sheet_case(name).unwrap_or_else(|| name.clone());
                render_token(&mut output, Token::Name { sheet: None, name });
            }
            Token::Name { sheet, name } => {
                // A function this build does not know keeps its spelling:
                // measured, `=nosuchfn()` reads back as written.
                let name = if calls {
                    if crate::functions::is_known_function(name) {
                        name.to_ascii_uppercase()
                    } else {
                        // One of the workbook's own functions is spelt as it
                        // was declared: measured, `=twice("x")` reads back
                        // `=Twice("x")`.
                        name_case(name).unwrap_or_else(|| name.clone())
                    }
                } else if let Some(cell) = absolute_r1c1(name) {
                    cell
                } else if parse_a1(name).is_some()
                    || name.eq_ignore_ascii_case("TRUE")
                    || name.eq_ignore_ascii_case("FALSE")
                    || (beside_colon && name.trim_start_matches('$').chars().all(|ch| ch.is_ascii_alphabetic()))
                {
                    name.to_ascii_uppercase()
                } else if sheet.is_none() {
                    name_case(name).unwrap_or_else(|| name.clone())
                } else {
                    name.clone()
                };
                let sheet = sheet.as_ref().map(|sheet| sheet_case(sheet).unwrap_or_else(|| sheet.clone()));
                render_token(&mut output, Token::Name { sheet, name });
            }
            Token::Number(value) if value.is_finite() => {
                output.push_str(&formula_number_text(text).unwrap_or_else(|| formula_number(*value)))
            }
            Token::ErrorLit(_) => render_token(&mut output, token.clone()),
            _ => output.push_str(text),
        }
        if !matches!(tokens.get(index + 1), Some((Token::Comma, _))) {
            output.push_str(gap);
        }
    }
    output
}

/// A number as Excel writes it in a formula: its fifteen significant digits,
/// plainly while that takes at most 21 characters and in exponent form past
/// that. Measured: 1E+20 is 100000000000000000000, 1.5E+21 stays 1.5E+21,
/// 123456789012345678 is 123456789012345000, 1E-10 is 0.0000000001, 1.5E-20
/// stays 1.5E-20, and pi is 3.14159265358979.
fn formula_number(value: f64) -> String {
    if value == 0.0 {
        return "0".to_string();
    }
    let written = format!("{:.14e}", value.abs());
    let (mantissa, exponent) = written.split_once('e').unwrap_or((&written, "0"));
    let exponent: i32 = exponent.parse().unwrap_or(0);
    let digits: String = mantissa.chars().filter(char::is_ascii_digit).collect();
    let digits = digits.trim_end_matches('0');
    let digits = if digits.is_empty() { "0" } else { digits };
    let plain = if exponent >= 0 {
        let whole_len = exponent as usize + 1;
        if digits.len() <= whole_len {
            format!("{digits}{}", "0".repeat(whole_len - digits.len()))
        } else {
            format!("{}.{}", &digits[..whole_len], &digits[whole_len..])
        }
    } else {
        format!("0.{}{digits}", "0".repeat((-exponent - 1) as usize))
    };
    let sign = if value < 0.0 { "-" } else { "" };
    if plain.len() <= 21 {
        return format!("{sign}{plain}");
    }
    let lead = &digits[..1];
    let rest = &digits[1..];
    let mantissa = if rest.is_empty() { lead.to_string() } else { format!("{lead}.{rest}") };
    format!("{sign}{mantissa}E{}{:02}", if exponent < 0 { '-' } else { '+' }, exponent.abs())
}

/// A number as written in a formula, cut -- not rounded -- to fifteen
/// significant digits, then written as `formula_number` writes it. Measured:
/// 123456789012345678 is 123456789012345000 and 0.1234567890123456789 is
/// 0.123456789012345, the cell's value cut the same way.
fn formula_number_text(text: &str) -> Option<String> {
    let (mantissa, exponent) = match text.find(['e', 'E']) {
        Some(at) => (&text[..at], text[at + 1..].parse::<i32>().ok()?),
        None => (text, 0),
    };
    let (whole, fraction) = mantissa.split_once('.').unwrap_or((mantissa, ""));
    if !whole.chars().chain(fraction.chars()).all(|ch| ch.is_ascii_digit()) {
        return None;
    }
    let all: String = format!("{whole}{fraction}");
    let leading = all.chars().take_while(|ch| *ch == '0').count();
    if leading == all.len() {
        return Some("0".to_string());
    }
    let significant: String = all[leading..].chars().take(15).collect();
    // The power of ten of the first significant digit.
    let power = whole.len() as i32 - 1 - leading as i32 + exponent;
    let rebuilt = format!("{}.{}e{}", &significant[..1], &significant[1..], power);
    let value: f64 = rebuilt.replace(".e", "e").parse().ok()?;
    Some(formula_number(value))
}

/// `R2C3` written in an A1 formula is the cell it names, absolutely:
/// measured, `=R2C3+r10c1` reads back `=$C$2+$A$10`.
fn absolute_r1c1(name: &str) -> Option<String> {
    let upper = name.to_ascii_uppercase();
    let rest = upper.strip_prefix('R')?;
    let (row, col) = rest.split_once('C')?;
    if row.is_empty() || col.is_empty() || !row.bytes().all(|b| b.is_ascii_digit()) || !col.bytes().all(|b| b.is_ascii_digit()) {
        return None;
    }
    let (row, col): (u32, u32) = (row.parse().ok()?, col.parse().ok()?);
    if !(1..=MAX_ROW + 1).contains(&row) || !(1..=MAX_COL + 1).contains(&col) {
        return None;
    }
    let mut letters = String::new();
    let mut left = col;
    while left > 0 {
        let digit = (left - 1) % 26;
        letters.insert(0, (b'A' + digit as u8) as char);
        left = (left - 1) / 26;
    }
    Some(format!("${letters}${row}"))
}

/// A formula as `.Formula` shows it: every `@` taken out. Measured, `=@A1:A3`
/// written through `.Formula` reads back `=A1:A3`.
pub fn without_intersections(input: &str) -> String {
    let Ok(tokens) = tokenize_spanned(input) else {
        return input.to_string();
    };
    let cuts: Vec<usize> = tokens.iter().filter(|(token, _)| *token == Token::At).map(|(_, at)| *at).collect();
    if cuts.is_empty() {
        return input.to_string();
    }
    let mut output = String::with_capacity(input.len());
    for (index, ch) in input.char_indices() {
        if !cuts.contains(&index) {
            output.push(ch);
        }
    }
    output
}

/// What a place in a formula wants: one value, a range (which an expression
/// in it still works out one value at a time), or anything at all.
#[derive(Clone, Copy, PartialEq)]
enum AtContext {
    Value,
    Range,
    Array,
}

enum AtNode {
    /// A number, text, a single cell, or anything else that is one value.
    Plain { number: Option<f64>, cell: bool },
    /// A range, or a name standing for one.
    Block { start: usize },
    /// Operators, whose operands are worked one value at a time.
    Operation(Vec<AtNode>),
    /// A `@` already written.
    Written,
    Group(Vec<AtNode>),
    Call { start: usize, name: String, args: Vec<Option<AtNode>> },
}

struct AtReader<'a> {
    tokens: &'a [(Token, usize)],
    pos: usize,
    multi_name: &'a dyn Fn(&str) -> bool,
}

impl AtReader<'_> {
    fn peek(&self) -> Option<&Token> {
        self.tokens.get(self.pos).map(|(token, _)| token)
    }

    fn start(&self) -> usize {
        self.tokens.get(self.pos).map_or(0, |(_, at)| *at)
    }

    fn expr(&mut self) -> Option<AtNode> {
        let mut parts = vec![self.unary()?];
        while matches!(
            self.peek(),
            Some(
                Token::Plus
                    | Token::Minus
                    | Token::Star
                    | Token::Slash
                    | Token::Caret
                    | Token::Amp
                    | Token::Eq
                    | Token::Ne
                    | Token::Lt
                    | Token::Le
                    | Token::Gt
                    | Token::Ge
            )
        ) {
            self.pos += 1;
            parts.push(self.unary()?);
        }
        Some(if parts.len() == 1 { parts.pop()? } else { AtNode::Operation(parts) })
    }

    fn unary(&mut self) -> Option<AtNode> {
        match self.peek()? {
            Token::Minus | Token::Plus => {
                self.pos += 1;
                let operand = self.unary()?;
                Some(AtNode::Operation(vec![operand]))
            }
            Token::At => {
                self.pos += 1;
                self.unary()?;
                Some(AtNode::Written)
            }
            _ => {
                let start = self.start();
                let mut node = self.primary()?;
                if self.peek() == Some(&Token::Colon) {
                    while self.peek() == Some(&Token::Colon) {
                        self.pos += 1;
                        self.primary()?;
                    }
                    node = AtNode::Block { start };
                }
                if self.peek() == Some(&Token::Hash) {
                    self.pos += 1;
                    node = AtNode::Block { start };
                }
                while self.peek() == Some(&Token::Intersect) {
                    self.pos += 1;
                    self.primary()?;
                    while self.peek() == Some(&Token::Colon) {
                        self.pos += 1;
                        self.primary()?;
                    }
                    node = AtNode::Block { start };
                }
                if self.peek() == Some(&Token::Percent) {
                    while self.peek() == Some(&Token::Percent) {
                        self.pos += 1;
                    }
                    node = AtNode::Operation(vec![node]);
                }
                Some(node)
            }
        }
    }

    fn primary(&mut self) -> Option<AtNode> {
        let start = self.start();
        let token = self.peek()?.clone();
        self.pos += 1;
        match token {
            Token::Number(number) => Some(AtNode::Plain { number: Some(number), cell: false }),
            Token::Text(_) | Token::ErrorLit(_) | Token::Table { .. } => Some(AtNode::Plain { number: None, cell: false }),
            Token::Name { name, sheet } => {
                if self.peek() == Some(&Token::LParen) && sheet.is_none() {
                    self.pos += 1;
                    let mut args = Vec::new();
                    if self.peek() == Some(&Token::RParen) {
                        self.pos += 1;
                        return Some(AtNode::Call { start, name: name.to_ascii_uppercase(), args });
                    }
                    loop {
                        if matches!(self.peek(), Some(Token::Comma | Token::RParen)) {
                            args.push(None);
                        } else {
                            args.push(Some(self.expr()?));
                        }
                        match self.peek()? {
                            Token::Comma => self.pos += 1,
                            Token::RParen => {
                                self.pos += 1;
                                break;
                            }
                            _ => return None,
                        }
                    }
                    return Some(AtNode::Call { start, name: name.to_ascii_uppercase(), args });
                }
                if parse_a1(&name).is_some() {
                    return Some(AtNode::Plain { number: None, cell: true });
                }
                if sheet.is_none() && (self.multi_name)(&name) {
                    return Some(AtNode::Block { start });
                }
                Some(AtNode::Plain { number: None, cell: false })
            }
            Token::LParen => {
                let mut inner = vec![self.expr()?];
                while self.peek() == Some(&Token::Comma) {
                    self.pos += 1;
                    inner.push(self.expr()?);
                }
                if self.peek() != Some(&Token::RParen) {
                    return None;
                }
                self.pos += 1;
                Some(AtNode::Group(inner))
            }
            Token::LBrace => {
                while self.peek()? != &Token::RBrace {
                    self.pos += 1;
                }
                self.pos += 1;
                Some(AtNode::Plain { number: None, cell: false })
            }
            _ => None,
        }
    }
}

/// How each parameter of a function takes what it is given: `V` one value,
/// `R` a range (an expression in it still one value at a time), `A` an
/// array. The last letter repeats. Measured through `.Formula2`; a function
/// not listed is left as written.
fn parameter_kinds(name: &str) -> Option<&'static str> {
    Some(match name {
        "LEN" | "LEFT" | "RIGHT" | "MID" | "LOWER" | "UPPER" | "PROPER" | "TRIM" | "TEXT" | "VALUE"
        | "ROUND" | "ROUNDUP" | "ROUNDDOWN" | "INT" | "ABS" | "MOD" | "ISNUMBER" | "ISTEXT"
        | "ISBLANK" | "ISERROR" | "ISNA" | "IFERROR" | "IFNA" | "DATE" | "YEAR" | "MONTH" | "DAY"
        | "WEEKDAY" | "FIND" | "SEARCH" | "SUBSTITUTE" | "REPT" | "EXACT" | "NOT" | "SQRT" | "CHAR"
        | "CONCATENATE" | "TRANSPOSE" => "V",
        "IF" | "CHOOSE" => "VR",
        "VLOOKUP" | "HLOOKUP" => "VAV",
        "MATCH" | "XLOOKUP" => "VA",
        "LOOKUP" => "VA",
        "COUNTIF" | "AVERAGEIF" | "SUMIF" => "RVR",
        "LARGE" | "SMALL" => "RV",
        "RANK" => "VRV",
        "INDEX" => "AV",
        "OFFSET" => "RV",
        "ROW" | "COLUMN" | "SUM" | "MAX" | "MIN" | "AVERAGE" | "COUNT" | "COUNTA" | "PRODUCT"
        | "MEDIAN" | "STDEV" | "AND" | "OR" | "N" => "R",
        "SUMPRODUCT" | "MMULT" | "CONCAT" | "TEXTJOIN" | "EOMONTH" | "EDATE" | "ROWS" | "COLUMNS"
        | "INDIRECT" | "FILTER" | "SORT" | "SORTBY" | "UNIQUE" | "SEQUENCE" => "A",
        _ => return None,
    })
}

fn parameter_kind(name: &str, index: usize) -> AtContext {
    // The pairs of COUNTIFS / SUMIFS: a range, then what to look for.
    match name {
        "COUNTIFS" => {
            return if index % 2 == 0 { AtContext::Range } else { AtContext::Value };
        }
        "SUMIFS" | "AVERAGEIFS" | "MAXIFS" | "MINIFS" => {
            return if index == 0 || index % 2 == 1 { AtContext::Range } else { AtContext::Value };
        }
        _ => {}
    }
    let Some(kinds) = parameter_kinds(name) else {
        return AtContext::Array;
    };
    let letter = kinds.as_bytes().get(index).or_else(|| kinds.as_bytes().last()).copied();
    match letter {
        Some(b'V') => AtContext::Value,
        Some(b'R') => AtContext::Range,
        _ => AtContext::Array,
    }
}

/// Whether a function may answer with more than one cell where it stands.
fn may_answer_a_block(name: &str, args: &[Option<AtNode>]) -> bool {
    let is_block = |node: &Option<AtNode>| matches!(node, Some(AtNode::Block { .. }));
    let loose_index = |node: &Option<AtNode>| match node {
        None => true,
        Some(AtNode::Plain { number: Some(number), .. }) => *number == 0.0,
        Some(AtNode::Plain { cell: true, .. }) | Some(AtNode::Block { .. }) => true,
        _ => false,
    };
    match name {
        "INDIRECT" | "FILTER" | "SORT" | "SORTBY" | "UNIQUE" | "SEQUENCE" | "TRANSPOSE" | "MMULT" => true,
        "OFFSET" => args.iter().skip(3).take(2).any(|size| match size {
            Some(AtNode::Plain { number: Some(number), .. }) => *number != 1.0,
            None => false,
            Some(_) => true,
        }),
        "INDEX" => args.iter().skip(1).take(2).any(loose_index),
        "IF" => args.iter().skip(1).any(is_block),
        "CHOOSE" => args.iter().skip(1).any(is_block),
        "ROW" | "COLUMN" => args.first().is_some_and(is_block),
        _ => false,
    }
}

fn mark_intersections(node: &AtNode, context: AtContext, marks: &mut Vec<usize>) {
    match node {
        AtNode::Plain { .. } | AtNode::Written => {}
        AtNode::Block { start } => {
            if context == AtContext::Value {
                marks.push(*start);
            }
        }
        AtNode::Operation(parts) => {
            let inner = if context == AtContext::Array { AtContext::Array } else { AtContext::Value };
            for part in parts {
                mark_intersections(part, inner, marks);
            }
        }
        AtNode::Group(inner) => {
            if let [only] = inner.as_slice() {
                mark_intersections(only, context, marks);
            }
        }
        AtNode::Call { start, name, args } => {
            if context == AtContext::Value && may_answer_a_block(name, args) {
                marks.push(*start);
            }
            for (index, arg) in args.iter().enumerate() {
                if let Some(arg) = arg {
                    mark_intersections(arg, parameter_kind(name, index), marks);
                }
            }
        }
    }
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
        // A reference carried off the sheet's edge is `#REF!`, the way Excel
        // writes it: measured, `=$A$1+B1` filled one column to the left reads
        // `=$A$1+#REF!`.
        let row = if reference.row_absolute {
            Ok(reference.row)
        } else {
            shifted_coordinate(reference.row, row_offset, MAX_ROW)
        };
        let col = if reference.col_absolute {
            Ok(reference.col)
        } else {
            shifted_coordinate(reference.col, column_offset, MAX_COL)
        };
        *name = match (row, col) {
            (Ok(row), Ok(col)) => {
                reference.row = row;
                reference.col = col;
                reference.to_a1()
            }
            _ => "#REF!".to_string(),
        };
    }
    // A range with either end carried off the sheet is #REF! as a whole:
    // measured, `=SUM(R[-2]C:R[-1]C)` filled from B2 to C3 reads
    // `=SUM(#REF!)` there.
    let lost = |token: &Token| matches!(token, Token::Name { name, .. } if name == "#REF!");
    let mut kept: Vec<Token> = Vec::with_capacity(tokens.len());
    let mut index = 0;
    while index < tokens.len() {
        if let (Some(near), Some(Token::Colon), Some(far)) = (tokens.get(index), tokens.get(index + 1), tokens.get(index + 2)) {
            if (lost(near) || lost(far)) && matches!(near, Token::Name { .. }) && matches!(far, Token::Name { .. }) {
                kept.push(Token::Name { sheet: None, name: "#REF!".to_string() });
                index += 3;
                continue;
            }
        }
        kept.push(tokens[index].clone());
        index += 1;
    }
    let tokens = kept;

    let mut output = String::new();
    if had_equals {
        output.push('=');
    }
    for token in tokens {
        render_token(&mut output, token);
    }
    Ok(output)
}

/// Write every range top-left first, the way Excel stores it.
///
/// Measured through VBA's `.Formula`: `SUM(B6:A5)` is kept as `SUM(A5:B6)`,
/// `C:A` as `A:C` and `6:5` as `5:6`. Rows and columns are put in order each
/// on its own, and a `$` goes with the coordinate it was written on:
/// `B$6:$A5` becomes `$A5:B$6`. A formula with nothing out of order, or one
/// that will not tokenize, comes back exactly as it was given. A table asked
/// for nothing, `tblP[]`, is written as the table's bare name.
pub fn normalise_formula_ranges(input: &str) -> String {
    let Ok(tokens) = tokenize(input) else {
        return input.to_string();
    };
    let mut changed = false;
    let mut written: Vec<Token> = Vec::with_capacity(tokens.len());
    let mut index = 0;
    while index < tokens.len() {
        // `tblP[]` is stored as `tblP`: measured, `=SUM(tblP[])` reads back
        // `=SUM(tblP)`.
        if let Token::Table { name, asked } = &tokens[index] {
            if asked.trim().is_empty() && !name.is_empty() {
                changed = true;
                written.push(Token::Name { sheet: None, name: name.clone() });
                index += 1;
                continue;
            }
        }
        if let (Some(near), Some(Token::Colon), Some(far)) =
            (tokens.get(index), tokens.get(index + 1), tokens.get(index + 2))
        {
            if let Some((first, second)) = ordered_range(near, far) {
                changed = true;
                written.push(first);
                written.push(Token::Colon);
                written.push(second);
                index += 3;
                continue;
            }
        }
        written.push(tokens[index].clone());
        index += 1;
    }
    if !changed {
        return input.to_string();
    }
    let mut output = String::new();
    if input.trim_start().starts_with('=') {
        output.push('=');
    }
    for token in written {
        render_token(&mut output, token);
    }
    output
}

/// A formula as read from a cell inside the table `table`: a reference to
/// ONE of that table's columns drops the table's name. Measured through
/// VBA's `.Formula`: `tblP[Qty]` reads `[Qty]`, `tblP[@Qty]` and
/// `tblP[[#This Row],[Qty]]` read `[@Qty]`, while `tblP[@[Qty]:[X]]`,
/// `tblP[[Qty]:[X]]`, `tblP[[#Headers],[Qty]]`, `tblP[#Data]` and a bare
/// `tblP` keep it. A formula that will not tokenize comes back as it was.
pub fn drop_own_table_name(input: &str, table: &str) -> String {
    let Ok(tokens) = tokenize(input) else {
        return input.to_string();
    };
    // The one column a specifier names, as it should be written back, and
    // whether it is this row's cell of it.
    fn one_column(asked: &str) -> Option<(String, bool)> {
        let column = |text: &str| -> Option<String> {
            let text = text.trim();
            let bare = text.strip_prefix('[').and_then(|rest| rest.strip_suffix(']')).unwrap_or(text);
            (!bare.is_empty() && !bare.starts_with('#') && !bare.contains(['[', ']', ':'])).then(|| {
                if text.starts_with('[') { text.to_string() } else { bare.to_string() }
            })
        };
        if let Some(rest) = asked.strip_prefix('@') {
            return column(rest).map(|name| (name, true));
        }
        if let Some(rest) = asked.strip_prefix("[#This Row],") {
            let name = column(rest)?;
            let bare = name.trim_start_matches('[').trim_end_matches(']').to_string();
            let plain = bare.chars().all(|ch| ch.is_alphanumeric() || ch == '_' || ch == '.');
            return Some((if plain { bare } else { format!("[{bare}]") }, true));
        }
        column(asked).map(|name| (name, false))
    }
    let mut changed = false;
    let written: Vec<Token> = tokens
        .into_iter()
        .map(|token| match token {
            Token::Table { name, asked } if name.eq_ignore_ascii_case(table) => match one_column(&asked) {
                Some((column, this_row)) => {
                    changed = true;
                    let asked = if this_row { format!("@{column}") } else { column };
                    Token::Table { name: String::new(), asked }
                }
                None => Token::Table { name, asked },
            },
            other => other,
        })
        .collect();
    if !changed {
        return input.to_string();
    }
    let mut output = String::new();
    if input.trim_start().starts_with('=') {
        output.push('=');
    }
    for token in written {
        render_token(&mut output, token);
    }
    output
}

/// The two ends of a range put in order, or None when they already are (or
/// are not the ends of a range at all).
fn ordered_range(near: &Token, far: &Token) -> Option<(Token, Token)> {
    let (sheet, start, far_sheet, end) = (&far_sheet(near), far_text(near)?, far_sheet(far), far_text(far)?);
    let start = start.as_str();
    let rebuilt = |sheet: &Option<String>, name: String| Token::Name { sheet: sheet.clone(), name };
    if let (Some(mut low), Some(mut high)) = (parse_a1(start), parse_a1(&end)) {
        // A block reaching from the first row to the last is the whole
        // column, and one from the first column to the last the whole row:
        // measured, B1:B1048576 reads back B:B.
        let whole = |low: &crate::reference::CellRef, high: &crate::reference::CellRef| -> Option<(Token, Token)> {
            let dollar = |absolute: bool| if absolute { "$" } else { "" };
            if low.row == 0 && high.row == 1_048_575 && low.row_absolute == high.row_absolute {
                return Some((
                    rebuilt(sheet, format!("{}{}", dollar(low.col_absolute), crate::reference::col_to_letters(low.col))),
                    rebuilt(&far_sheet, format!("{}{}", dollar(high.col_absolute), crate::reference::col_to_letters(high.col))),
                ));
            }
            if low.col == 0 && high.col == 16_383 && low.col_absolute == high.col_absolute {
                return Some((
                    rebuilt(sheet, format!("{}{}", dollar(low.row_absolute), low.row + 1)),
                    rebuilt(&far_sheet, format!("{}{}", dollar(high.row_absolute), high.row + 1)),
                ));
            }
            None
        };
        if low.row <= high.row && low.col <= high.col {
            return whole(&low, &high);
        }
        if low.row > high.row {
            std::mem::swap(&mut low.row, &mut high.row);
            std::mem::swap(&mut low.row_absolute, &mut high.row_absolute);
        }
        if low.col > high.col {
            std::mem::swap(&mut low.col, &mut high.col);
            std::mem::swap(&mut low.col_absolute, &mut high.col_absolute);
        }
        if let Some(whole) = whole(&low, &high) {
            return Some(whole);
        }
        return Some((rebuilt(sheet, low.to_a1()), rebuilt(&far_sheet, high.to_a1())));
    }
    // A whole column or a whole row: `C:A`, `$6:5`.
    let line = |text: &str| -> Option<(bool, u32, bool)> {
        let (absolute, rest) = match text.strip_prefix('$') {
            Some(rest) => (true, rest),
            None => (false, text),
        };
        if !rest.is_empty() && rest.bytes().all(|b| b.is_ascii_digit()) {
            return Some((false, rest.parse().ok()?, absolute));
        }
        if !rest.is_empty() && rest.bytes().all(|b| b.is_ascii_alphabetic()) {
            let column = rest
                .bytes()
                .fold(0u32, |held, b| held * 26 + u32::from(b.to_ascii_uppercase() - b'A' + 1));
            return Some((true, column, absolute));
        }
        None
    };
    let (Some(one), Some(other)) = (line(start), line(&end)) else {
        return None;
    };
    if one.0 != other.0 || one.1 <= other.1 {
        return None;
    }
    let text_of = |held: &str| held.to_string();
    Some((rebuilt(sheet, text_of(&end)), rebuilt(&far_sheet, text_of(start))))
}

fn far_sheet(far: &Token) -> Option<String> {
    match far {
        Token::Name { sheet, .. } => sheet.clone(),
        _ => None,
    }
}

fn far_text(far: &Token) -> Option<String> {
    match far {
        Token::Name { name, .. } => Some(name.clone()),
        Token::Number(value) if value.fract() == 0.0 && *value >= 1.0 => Some(format!("{value}")),
        _ => None,
    }
}

/// Rewrite every reference to the sheet named `old` so it names `new` instead,
/// leaving string literals, other sheets and unqualified references untouched.
///
/// The match is case-insensitive on ASCII, as Excel treats sheet names, and
/// exact otherwise (a Japanese name has no case to fold). A formula that will
/// not tokenize -- one that links to another workbook, say -- is returned
/// unchanged rather than mangled.
pub fn rename_sheet_in_formula(input: &str, old: &str, new: &str) -> String {
    let had_equals = input.trim_start().starts_with('=');
    let Ok(mut tokens) = tokenize(input) else {
        return input.to_string();
    };
    let mut touched = false;
    let count = tokens.len();
    for at in 0..count {
        // The first sheet of `Jan:Mar!A1` is written bare before the colon.
        let opens_a_span = matches!(tokens.get(at + 1), Some(Token::Colon))
            && matches!(tokens.get(at + 2), Some(Token::Name { sheet: Some(_), .. }));
        match &mut tokens[at] {
            Token::Name { sheet: Some(sheet), .. } if sheet.eq_ignore_ascii_case(old) => {
                *sheet = new.to_string();
                touched = true;
            }
            Token::Name { sheet: None, name } if opens_a_span && name.eq_ignore_ascii_case(old) => {
                *name = new.to_string();
                touched = true;
            }
            _ => {}
        }
    }
    if !touched {
        return input.to_string();
    }
    let mut output = String::new();
    if had_equals {
        output.push('=');
    }
    for token in tokens {
        render_token(&mut output, token);
    }
    output
}

/// A formula once the sheet `gone` is deleted, `order` being the sheets as
/// they stood. Measured through `.Formula`: `=Jan!A1+Data3!A1` becomes
/// `=Jan!A1+#REF!A1`, and `=SUM(Jan:Data3!A1)`, whose last sheet went,
/// `=SUM(Jan:Data2!A1)` -- the span pulls in to the sheet beside it.
pub fn drop_sheet_in_formula(input: &str, gone: &str, order: &[String]) -> String {
    let had_equals = input.trim_start().starts_with('=');
    let Ok(mut tokens) = tokenize(input) else {
        return input.to_string();
    };
    let place = |wanted: &str| order.iter().position(|held| held.eq_ignore_ascii_case(wanted));
    let mut touched = false;
    let count = tokens.len();
    for at in 0..count {
        let first_of_span = match (tokens.get(at + 1), tokens.get(at + 2)) {
            (Some(Token::Colon), Some(Token::Name { sheet: Some(last), .. })) => Some(last.clone()),
            _ => None,
        };
        let last_of_span = match (at.checked_sub(2).and_then(|before| tokens.get(before)), at.checked_sub(1).and_then(|before| tokens.get(before))) {
            (Some(Token::Name { sheet: None, name: first }), Some(Token::Colon)) => Some(first.clone()),
            _ => None,
        };
        match tokens[at].clone() {
            Token::Name { sheet: None, name } if first_of_span.is_some() && name.eq_ignore_ascii_case(gone) => {
                let last = first_of_span.unwrap_or_default();
                if let (Some(from), Some(to)) = (place(&name), place(&last)) {
                    let step = if to >= from { from + 1 } else { from.saturating_sub(1) };
                    if let Some(next) = order.get(step) {
                        tokens[at] = Token::Name { sheet: None, name: next.clone() };
                        touched = true;
                    }
                }
            }
            Token::Name { sheet: Some(sheet), name } if sheet.eq_ignore_ascii_case(gone) => {
                touched = true;
                tokens[at] = match last_of_span.as_deref().and_then(|first| Some((place(first)?, place(&sheet)?))) {
                    Some((from, to)) if from != to => {
                        let step = if to > from { to - 1 } else { to + 1 };
                        Token::Name { sheet: Some(order[step].clone()), name }
                    }
                    _ => Token::Name { sheet: None, name: format!("#REF!{name}") },
                };
            }
            _ => {}
        }
    }
    if !touched {
        return input.to_string();
    }
    let mut output = String::new();
    if had_equals {
        output.push('=');
    }
    for token in tokens {
        render_token(&mut output, token);
    }
    output
}

/// A formula once the sheet `moved` has moved from where it stood in
/// `before` to where it stands in `after`. A span `Qa:Qc!A1` follows its end
/// sheets wherever they go -- measured, moving Qc past the last sheet widens
/// it -- until one is taken past the other end, when that end becomes the
/// sheet that stood beside it inside the span: measured, `Qa:Qc` with Qc
/// moved to the front reads `Qa:Qb`.
pub fn move_sheet_in_formula(input: &str, moved: &str, before: &[String], after: &[String]) -> String {
    let had_equals = input.trim_start().starts_with('=');
    let Ok(mut tokens) = tokenize(input) else {
        return input.to_string();
    };
    let at_in = |order: &[String], wanted: &str| order.iter().position(|held| held.eq_ignore_ascii_case(wanted));
    let mut touched = false;
    let count = tokens.len();
    for at in 0..count {
        let (Token::Name { sheet: None, name: first }, Some(Token::Colon), Some(Token::Name { sheet: Some(last), name: cell })) =
            (tokens[at].clone(), tokens.get(at + 1).cloned(), tokens.get(at + 2).cloned())
        else {
            continue;
        };
        let (Some(first_now), Some(last_now), Some(first_was), Some(last_was)) =
            (at_in(after, &first), at_in(after, &last), at_in(before, &first), at_in(before, &last))
        else {
            continue;
        };
        let forward = last_was >= first_was;
        if first.eq_ignore_ascii_case(moved) && (if forward { first_now > last_now } else { first_now < last_now }) {
            let inner = if forward { first_was + 1 } else { first_was.saturating_sub(1) };
            if let Some(next) = before.get(inner) {
                tokens[at] = Token::Name { sheet: None, name: next.clone() };
                touched = true;
            }
        } else if last.eq_ignore_ascii_case(moved) && (if forward { last_now < first_now } else { last_now > first_now }) {
            let inner = if forward { last_was.saturating_sub(1) } else { last_was + 1 };
            if let Some(next) = before.get(inner) {
                tokens[at + 2] = Token::Name { sheet: Some(next.clone()), name: cell };
                touched = true;
            }
        }
    }
    if !touched {
        return input.to_string();
    }
    let mut output = String::new();
    if had_equals {
        output.push('=');
    }
    for token in tokens {
        render_token(&mut output, token);
    }
    output
}

/// A band of rows or columns that was inserted or removed, and how far its
/// effect reaches.
///
/// `across` is how far the band reaches along the other axis, one-based and
/// inclusive — the columns a row band spans, or the rows a column band spans.
/// Only a reference lying wholly within it moves, which is what makes a partial
/// insert leave neighbouring columns alone: shifting `B2` down rewrites `B3`
/// but not `C3`, and leaves `SUM(A1:C3)` alone because it reaches past B. Use
/// the sheet's full extent for a whole-row or whole-column band.
#[derive(Debug, Clone, Copy)]
pub struct ReferenceShift<'a> {
    pub axis: ShiftAxis,
    /// One-based first index of the band.
    pub at: u32,
    /// How many were put in (positive) or taken out (negative).
    pub count: i64,
    /// One-based inclusive reach along the other axis.
    pub across: (u32, u32),
    /// The sheet whose cells moved. A reference naming a different sheet is
    /// left alone, while one naming this sheet moves even from another sheet's
    /// formula, which is how `=Data!A5` follows a row inserted on `Data`.
    pub sheet: Option<&'a str>,
    /// The sheet the formula being rewritten is written ON, when that is
    /// known.
    ///
    /// An unqualified `A1` means this sheet, so it moves only when this is the
    /// sheet the cells moved on. Without it, rewriting another sheet's
    /// formulas moves references that never pointed at the change: a row put
    /// into `Data` would drag `=A3` on `Summary` along with it.
    ///
    /// `None` means it is not known, and then an unqualified reference is
    /// taken to be on the moved sheet — which is what a caller rewriting only
    /// that sheet wants.
    pub on_sheet: Option<&'a str>,
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
///
/// See [`ReferenceShift`] for what the band covers and which sheet it moved.
pub fn shift_formula_references(
    input: &str,
    shift: &ReferenceShift<'_>,
) -> Result<String, String> {
    let ReferenceShift {
        axis,
        at,
        count,
        across,
        sheet: moved_sheet,
        on_sheet,
    } = *shift;
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
    // The other axis, where the band's reach decides whether a reference moves.
    let crossing = |reference: &CellRef| match axis {
        ShiftAxis::Rows => reference.col,
        ShiftAxis::Columns => reference.row,
    };
    let (first_across, last_across) = (across.0.saturating_sub(1), across.1.saturating_sub(1));
    let within = |low: u32, high: u32| low >= first_across && high <= last_across;
    let with_coordinate = |mut reference: CellRef, value: u32| {
        match axis {
            ShiftAxis::Rows => reference.row = value,
            ShiftAxis::Columns => reference.col = value,
        }
        reference
    };

    // A whole column or a whole row, `B:B` or `$2:3`: which it is, its
    // one-based number and whether it is pinned.
    let line_of = |token: &Token| -> Option<(bool, u32, bool)> {
        let text = match token {
            Token::Name { name, .. } => name.clone(),
            Token::Number(value) if value.fract() == 0.0 && *value >= 1.0 => format!("{value}"),
            _ => return None,
        };
        let (absolute, rest) = match text.strip_prefix('$') {
            Some(rest) => (true, rest.to_string()),
            None => (false, text.clone()),
        };
        if !rest.is_empty() && rest.bytes().all(|b| b.is_ascii_digit()) {
            return Some((false, rest.parse().ok()?, absolute));
        }
        if !rest.is_empty() && rest.len() <= 3 && rest.bytes().all(|b| b.is_ascii_alphabetic()) {
            let column = rest
                .bytes()
                .fold(0u32, |held, b| held * 26 + u32::from(b.to_ascii_uppercase() - b'A' + 1));
            return Some((true, column, absolute));
        }
        None
    };
    let write_line = |column: bool, number: u32, absolute: bool| -> String {
        let pin = if absolute { "$" } else { "" };
        if column {
            let mut letters = String::new();
            let mut left = number;
            while left > 0 {
                letters.insert(0, (b'A' + ((left - 1) % 26) as u8) as char);
                left = (left - 1) / 26;
            }
            format!("{pin}{letters}")
        } else {
            format!("{pin}{number}")
        }
    };
    let across_everything = match axis {
        ShiftAxis::Rows => first_across == 0 && last_across >= MAX_COL,
        ShiftAxis::Columns => first_across == 0 && last_across >= MAX_ROW,
    };

    let mut shifted = Vec::with_capacity(tokens.len());
    let mut index = 0;
    while index < tokens.len() {
        let is_function = matches!(tokens.get(index + 1), Some(Token::LParen));
        // A whole column moves with columns put in or taken out across every
        // row, and a whole row with rows: measured, `SUM(B:B)` is `SUM(C:C)`
        // after a column goes in at A and `SUM(#REF!)` once B is deleted,
        // `SUM(1:3)` `SUM(2:4)` after a row goes in above.
        if !is_function && matches!(tokens.get(index + 1), Some(Token::Colon)) {
            if let (Some(first), Some(last)) = (
                line_of(&tokens[index]),
                tokens.get(index + 2).and_then(|token| line_of(token)),
            ) {
                let sheet = match &tokens[index] {
                    Token::Name { sheet, .. } => sheet.clone(),
                    _ => None,
                };
                let names_moved_sheet = match (sheet.as_deref(), moved_sheet) {
                    (None, Some(moved)) => on_sheet.is_none_or(|own| own.eq_ignore_ascii_case(moved)),
                    (None, None) => true,
                    (Some(named), Some(moved)) => named.eq_ignore_ascii_case(moved),
                    (Some(_), None) => false,
                };
                if first.0 == last.0 {
                    let moves = names_moved_sheet
                        && across_everything
                        && (first.0 == (axis == ShiftAxis::Columns));
                    if moves {
                        match shifted_range(first.1 - 1, last.1 - 1, at, count, maximum)? {
                            Some((low, high)) => {
                                shifted.push(Token::Name { sheet, name: write_line(first.0, low + 1, first.2) });
                                shifted.push(Token::Colon);
                                shifted.push(Token::Name { sheet: None, name: write_line(last.0, high + 1, last.2) });
                            }
                            None => shifted.push(Token::ErrorLit(ExcelError::Ref)),
                        }
                    } else {
                        shifted.extend(tokens[index..index + 3].iter().cloned());
                    }
                    index += 3;
                    continue;
                }
            }
        }
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
        // A reference naming another sheet points at cells this change never
        // touched, so it stays as it is.
        let names_moved_sheet = match (sheet.as_deref(), moved_sheet) {
            // Unqualified: it means the sheet the formula is written on, so it
            // moves only when that is the sheet the cells moved on.
            (None, Some(moved)) => {
                on_sheet.is_none_or(|own| own.eq_ignore_ascii_case(moved))
            }
            (None, None) => true,
            (Some(named), Some(moved)) => named.eq_ignore_ascii_case(moved),
            (Some(_), None) => false,
        };
        let Some(start) = parse_a1(name).filter(|_| names_moved_sheet) else {
            // A range whose near end names another sheet has to be stepped
            // over WHOLE. Its far end is written without a sheet of its own,
            // so reading that one alone takes it for this sheet's cell and
            // moves half the range: `SUM(Other!$B$3:$B$5)` came out as
            // `SUM(Other!$B$3:$B$6)`.
            let width = match (parse_a1(name), tokens.get(index + 1), tokens.get(index + 2)) {
                (Some(_), Some(Token::Colon), Some(Token::Name { name, .. }))
                    if parse_a1(name).is_some() =>
                {
                    3
                }
                _ => 1,
            };
            for step in 0..width {
                shifted.push(tokens[index + step].clone());
            }
            index += width;
            continue;
        };

        // A range moves as a whole, so its ends are decided together.
        let range_end = match (tokens.get(index + 1), tokens.get(index + 2)) {
            (Some(Token::Colon), Some(Token::Name { name, .. })) => parse_a1(name),
            _ => None,
        };
        if let Some(end) = range_end {
            let (near, far) = (crossing(&start), crossing(&end));
            if !within(near.min(far), near.max(far)) {
                shifted.push(tokens[index].clone());
                shifted.push(tokens[index + 1].clone());
                shifted.push(tokens[index + 2].clone());
                index += 3;
                continue;
            }
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

        let side = crossing(&start);
        if !within(side, side) {
            shifted.push(tokens[index].clone());
            index += 1;
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

/// A block of cells a cut moved, for rewriting the references that pointed at
/// them.
///
/// A cut is not a copy. The references FOLLOW the cells, absolute ones
/// included: asked of Excel, cutting `A2:B3` onto `D2` leaves `=SUM(A2:B3)`
/// reading `=SUM(D2:E3)` and `=$A$2` reading `=$D$2`, from any sheet. Only a
/// reference lying WHOLLY inside the block follows it, so `=SUM(A1:B4)`, which
/// reaches past the block, is left where it is, and so is `=SUM(A:A)`.
#[derive(Debug, Clone, Copy)]
pub struct CellMove<'a> {
    /// The block that moved, zero-based and inclusive.
    pub first_row: u32,
    pub first_column: u32,
    pub last_row: u32,
    pub last_column: u32,
    /// How far it went.
    pub down: i64,
    pub across: i64,
    /// The sheet the cells moved OFF. A reference naming another sheet points
    /// at cells this cut never touched.
    pub from_sheet: Option<&'a str>,
    /// The sheet they landed ON. `None` says they stayed where they were.
    pub to_sheet: Option<&'a str>,
    /// What an unqualified reference in this formula means.
    ///
    /// For a formula that TRAVELLED with the block this is the sheet it came
    /// from, not the one it now sits on: its references were written against
    /// the old home and have to be read there.
    pub read_as: Option<&'a str>,
    /// The sheet the formula now sits on, which decides whether a rewritten
    /// reference has to name its sheet. Asked of Excel, a formula carried to
    /// another sheet keeps `=D2*10` for a cell that came with it and gains
    /// `=Sheet3!G9` for one that stayed behind.
    pub written_on: Option<&'a str>,
}

impl CellMove<'_> {
    fn covers(&self, span: (u32, u32, u32, u32)) -> bool {
        let (low_row, low_column, high_row, high_column) = span;
        low_row >= self.first_row
            && high_row <= self.last_row
            && low_column >= self.first_column
            && high_column <= self.last_column
    }

    /// Where the block came to rest — the cells it overwrote on the way.
    pub(crate) fn landing(&self) -> Option<(u32, u32, u32, u32)> {
        let moved = |value: u32| u32::try_from(i64::from(value) + self.down).ok();
        let across = |value: u32| u32::try_from(i64::from(value) + self.across).ok();
        Some((
            moved(self.first_row)?,
            across(self.first_column)?,
            moved(self.last_row)?,
            across(self.last_column)?,
        ))
    }
}

/// Move the A1 references that pointed at cells a cut took away, the way Excel
/// rewrites formulas after a cut-and-paste.
///
/// A reference wholly inside the moved block follows it. A reference wholly
/// inside the cells the block LANDED on becomes `#REF!`, since what it named
/// was overwritten. Everything else — a range only partly overlapping either,
/// a whole column, another sheet's cells — keeps pointing where it did.
///
/// Where the block changed sheet, so does everything that followed it, and a
/// reference then has to say which sheet it means whenever that is no longer
/// the one the formula sits on. That is why a formula carried across says
/// `=Sheet3!G9` about a neighbour it left behind.
///
/// A range the block landed on the END of closes up to just before it, but
/// only where the block reaches PAST that end: `SUM(D1:D2)` becomes
/// `SUM(D1:D1)` when D2:E3 is landed on, and `SUM(D1:D3)` is left alone when
/// the block stops at row 3. A block landing on a range's near end, or inside
/// it, changes nothing — what is written there is somebody else's number now,
/// but the range still names the same cells.
pub fn move_formula_references(input: &str, moved: &CellMove<'_>) -> Result<String, String> {
    crate::parser::parse(input).map_err(|error| error.to_string())?;
    let had_equals = input.trim_start().starts_with('=');
    let tokens = tokenize(input).map_err(|error| error.to_string())?;
    let landing = moved.landing();
    let landed_on = moved.to_sheet.or(moved.from_sheet);

    let same = |one: Option<&str>, other: Option<&str>| match (one, other) {
        (Some(one), Some(other)) => one.eq_ignore_ascii_case(other),
        (None, None) => true,
        _ => false,
    };
    let travelled = |reference: CellRef| -> Result<CellRef, String> {
        Ok(CellRef {
            row: shifted_coordinate(reference.row, moved.down, MAX_ROW)?,
            col: shifted_coordinate(reference.col, moved.across, MAX_COL)?,
            ..reference
        })
    };

    let mut written = Vec::with_capacity(tokens.len());
    let mut index = 0;
    while index < tokens.len() {
        let is_function = matches!(tokens.get(index + 1), Some(Token::LParen));
        let Token::Name { sheet, name } = &tokens[index] else {
            written.push(tokens[index].clone());
            index += 1;
            continue;
        };
        if is_function {
            written.push(tokens[index].clone());
            index += 1;
            continue;
        }
        let Some(start) = parse_a1(name) else {
            written.push(tokens[index].clone());
            index += 1;
            continue;
        };
        // A range is judged as one thing, since it follows the cut as one.
        let end = match (tokens.get(index + 1), tokens.get(index + 2)) {
            (Some(Token::Colon), Some(Token::Name { name, .. })) => parse_a1(name),
            _ => None,
        };
        let width = if end.is_some() { 3 } else { 1 };
        let far = end.unwrap_or(start);
        let span = (
            start.row.min(far.row),
            start.col.min(far.col),
            start.row.max(far.row),
            start.col.max(far.col),
        );
        // An unqualified reference means whichever sheet this formula's
        // references were written against.
        let points_at = sheet.as_deref().or(moved.read_as);

        let follows = same(points_at, moved.from_sheet) && moved.covers(span);
        // What the block landed on it also overwrote, leaving nothing there to
        // name — unless the block itself brought it.
        let overwritten = !follows
            && same(points_at, landed_on)
            && landing.is_some_and(|(first_row, first_column, last_row, last_column)| {
                span.0 >= first_row
                    && span.2 <= last_row
                    && span.1 >= first_column
                    && span.3 <= last_column
            });
        if overwritten {
            written.push(Token::ErrorLit(ExcelError::Ref));
            index += width;
            continue;
        }

        // A range the block landed on the END of closes up to just before it.
        // The block has to reach PAST that end — where it stops exactly at the
        // end, or starts at it, or sits in the middle, Excel leaves the range
        // as it was.
        let closes_up = |(first_row, first_column, last_row, last_column): (u32, u32, u32, u32)| {
            let across_inside = span.1 >= first_column && span.3 <= last_column;
            let down_inside = span.0 >= first_row && span.2 <= last_row;
            if across_inside && first_row > span.0 && first_row <= span.2 && last_row > span.2 {
                return Some((span.0, span.1, first_row - 1, span.3));
            }
            if down_inside
                && first_column > span.1
                && first_column <= span.3
                && last_column > span.3
            {
                return Some((span.0, span.1, span.2, first_column - 1));
            }
            None
        };
        let closed = (!follows && same(points_at, landed_on) && end.is_some())
            .then(|| landing.and_then(closes_up))
            .flatten();
        if let Some((first_row, first_column, last_row, last_column)) = closed {
            let keep = |reference: CellRef, row: u32, col: u32| CellRef {
                row,
                col,
                ..reference
            };
            written.push(Token::Name {
                sheet: sheet.clone(),
                name: keep(start, first_row, first_column).to_a1(),
            });
            written.push(Token::Colon);
            let end_sheet = match &tokens[index + 2] {
                Token::Name { sheet, .. } => sheet.clone(),
                _ => None,
            };
            written.push(Token::Name {
                sheet: end_sheet,
                name: keep(far, last_row, last_column).to_a1(),
            });
            index += width;
            continue;
        }

        let now_at = if follows { landed_on } else { points_at };
        // It has to name its sheet when that is not the one it sits on, and
        // one that already named a sheet goes on naming it.
        let named = if sheet.is_some() || !same(now_at, moved.written_on) {
            now_at.map(str::to_string)
        } else {
            None
        };
        if !follows && (sheet.is_some() || named.is_none()) {
            // Nothing to say about it that it does not already say.
            for step in 0..width {
                written.push(tokens[index + step].clone());
            }
            index += width;
            continue;
        }

        let (near, far) = if follows {
            (travelled(start)?, travelled(far)?)
        } else {
            (start, far)
        };
        written.push(Token::Name {
            sheet: named,
            name: near.to_a1(),
        });
        if end.is_some() {
            let end_sheet = match &tokens[index + 2] {
                Token::Name { sheet, .. } => sheet.clone(),
                _ => None,
            };
            written.push(Token::Colon);
            written.push(Token::Name {
                sheet: end_sheet,
                name: far.to_a1(),
            });
        }
        index += width;
    }

    let mut output = String::new();
    if had_equals {
        output.push('=');
    }
    for token in written {
        render_token(&mut output, token);
    }
    Ok(output)
}

/// Turn a formula's references a quarter turn, as Excel does when a block is
/// pasted transposed.
///
/// A relative reference is a distance from the cell that holds it, and
/// transposing swaps the two halves of that distance: `=B3*2` written in C3
/// looks one to the LEFT, so pasted transposed it looks one ABOVE. Asked of
/// Excel, `=C2*2` in C3 pasted onto F2 reads `=E2*2`, `=Z9` reads `=L30`, and
/// `=SUM(A3:B3)` — a row — comes out as the column `=SUM(F4:F5)`. An absolute
/// reference names a fixed cell and does not turn.
///
/// A MIXED reference is left as it stands. Asked of Excel, `=B$3` and `=$A1`
/// both come out of a transposed paste unchanged — though one written as the
/// end of a RANGE does turn, `$A1:B2` becoming `F$6:G7`, which is a second
/// rule this does not follow.
///
/// `from` and `to` are the cell the formula was written in and the cell it is
/// being put down at, both zero-based as (row, column).
pub fn transpose_formula_references(
    input: &str,
    from: (u32, u32),
    to: (u32, u32),
) -> Result<String, String> {
    crate::parser::parse(input).map_err(|error| error.to_string())?;
    let had_equals = input.trim_start().starts_with('=');
    let tokens = tokenize(input).map_err(|error| error.to_string())?;

    let turned = |reference: CellRef| -> Option<CellRef> {
        if reference.row_absolute != reference.col_absolute {
            return Some(reference);
        }
        if reference.row_absolute {
            return Some(reference);
        }
        let down = i64::from(reference.col) - i64::from(from.1);
        let across = i64::from(reference.row) - i64::from(from.0);
        let row = i64::from(to.0) + down;
        let col = i64::from(to.1) + across;
        if row < 0 || col < 0 || row > i64::from(MAX_ROW) || col > i64::from(MAX_COL) {
            return None;
        }
        Some(CellRef {
            row: row as u32,
            col: col as u32,
            ..reference
        })
    };

    let mut written = Vec::with_capacity(tokens.len());
    for (index, token) in tokens.iter().enumerate() {
        let is_function = matches!(tokens.get(index + 1), Some(Token::LParen));
        let Token::Name { sheet, name } = token else {
            written.push(token.clone());
            continue;
        };
        match parse_a1(name).filter(|_| !is_function) {
            Some(reference) => match turned(reference) {
                Some(turned) => written.push(Token::Name {
                    sheet: sheet.clone(),
                    name: turned.to_a1(),
                }),
                None => written.push(Token::ErrorLit(ExcelError::Ref)),
            },
            None => written.push(token.clone()),
        }
    }

    let mut output = String::new();
    if had_equals {
        output.push('=');
    }
    for token in written {
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

pub(crate) fn render_token(output: &mut String, token: Token) {
    match token {
        Token::Number(value) => output.push_str(&value.to_string()),
        Token::Text(value) => {
            output.push('"');
            output.push_str(&value.replace('"', "\"\""));
            output.push('"');
        }
        Token::ErrorLit(value) => output.push_str(value.as_str()),
        Token::LBrace => output.push('{'),
        Token::At => output.push('@'),
        Token::Hash => output.push('#'),
        Token::Intersect => output.push(' '),
        Token::RBrace => output.push('}'),
        Token::Semicolon => output.push(';'),
        Token::Table { name, asked } => {
            output.push_str(&name);
            output.push('[');
            output.push_str(&asked);
            output.push(']');
        }
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
    // Past fifteen significant digits a number is cut, not rounded: measured,
    // `=COMPLEX(123456789012345678,0)` reads 123456789012345000.
    let mut kept = String::with_capacity(text.len());
    let mut significant = 0;
    let mut in_exponent = false;
    for c in text.chars() {
        if matches!(c, 'e' | 'E') {
            in_exponent = true;
        }
        if !in_exponent && c.is_ascii_digit() && (significant > 0 || c != '0') {
            significant += 1;
            if significant > 15 {
                kept.push('0');
                continue;
            }
        }
        kept.push(c);
    }
    kept.parse::<f64>()
        .map(|n| (n, i))
        .map_err(|_| ParseError::InvalidNumber(text.to_string()))
}

/// A table's name and the bracket group after it, or `None` when what is here
/// is not one.
///
/// The brackets nest — `tbl[[#This Row],[DATE]]` has two levels — so they are
/// counted rather than scanned to the first `]`. An unclosed group is not a
/// table reference at all, and is left for the ordinary lexer to complain
/// about wherever it actually goes wrong.
fn lex_table(src: &str, start: usize) -> Option<(Token, usize)> {
    let bytes = src.as_bytes();
    let mut at = start;
    // A table's name is a plain word: no sheet, no dollars.
    while at < bytes.len() {
        let ch = src[at..].chars().next()?;
        if ch.is_alphanumeric() || ch == '_' || ch == '.' {
            at += ch.len_utf8();
        } else {
            break;
        }
    }
    if at == start || bytes.get(at) != Some(&b'[') {
        return None;
    }
    let name = src[start..at].to_string();
    // A name that spells a cell is read as the cell, and a bracket after a
    // cell is no formula: measured, `=SUM(T1[n])` over a table named T1 is
    // refused (1004) and Evaluate of it is an error.
    if crate::reference::parse_a1(&name).is_some() {
        return None;
    }
    let inside = at + 1;
    let mut depth = 1usize;
    let mut end = inside;
    while end < bytes.len() {
        match bytes[end] {
            b'[' => depth += 1,
            b']' => {
                depth -= 1;
                if depth == 0 {
                    return Some((
                        Token::Table {
                            name,
                            asked: src[inside..end].to_string(),
                        },
                        end + 1,
                    ));
                }
            }
            _ => {}
        }
        end += 1;
    }
    None
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
        // `'[1]May 2021'!A1` names a sheet in a workbook this one only links
        // to. The bracket is left where it was written and the parser takes it
        // apart, the same way a table's brackets are kept whole here and read
        // there.
        sheet = Some(name);
        i = next + 1;
    } else {
        // Bare sheet prefix: `Sheet1!`, and `[1]Sheet1!` for a sheet in a
        // workbook this one links to. The link number is not part of a sheet
        // word, so it is stepped over here and kept in front of the name.
        let mut from = i;
        if src.as_bytes().get(from) == Some(&b'[') {
            if let Some(close) = src[from..].find(']') {
                let digits = &src[from + 1..from + close];
                if !digits.is_empty() && digits.bytes().all(|b| b.is_ascii_digit()) {
                    from += close + 1;
                }
            }
        }
        let end = scan_sheet_word(src, from);
        if src.as_bytes().get(end) == Some(&b'!') && end > from {
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
    fn rename_sheet_rewrites_only_that_sheets_references() {
        assert_eq!(
            rename_sheet_in_formula("=Data!A1+Summary!B2", "Data", "Sales"),
            "=Sales!A1+Summary!B2"
        );
        // The old name matches without regard to case; the new name is kept as
        // given.
        assert_eq!(rename_sheet_in_formula("=data!A1", "Data", "X"), "=X!A1");
        // A new name that needs quoting gets it; an old quoted name that no
        // longer needs quoting loses it.
        assert_eq!(rename_sheet_in_formula("=Data!A1", "Data", "My Sheet"), "='My Sheet'!A1");
        assert_eq!(rename_sheet_in_formula("='Old Name'!A1", "Old Name", "New"), "=New!A1");
        // A string literal that merely contains the name is left alone.
        assert_eq!(
            rename_sheet_in_formula("=\"Data!x\"&Data!A1", "Data", "Z"),
            "=\"Data!x\"&Z!A1"
        );
        // Nothing to rename comes back unchanged.
        assert_eq!(rename_sheet_in_formula("=A1+B2", "Data", "Z"), "=A1+B2");
    }

    /// A sheet renamed, deleted or moved is written into formulas as Excel
    /// writes it. Every answer measured through `.Formula`.
    #[test]
    fn sheets_renamed_deleted_and_moved_are_written_in() {
        assert_eq!(rename_sheet_in_formula("=SUM(Data1:Data3!A1)", "Data1", "Jan"), "=SUM(Jan:Data3!A1)");
        assert_eq!(rename_sheet_in_formula("=Data1!A1+Data3!A1", "Data1", "Jan"), "=Jan!A1+Data3!A1");
        let order: Vec<String> = ["Main", "Jan", "Data2", "Data3"].iter().map(|one| one.to_string()).collect();
        assert_eq!(drop_sheet_in_formula("=Jan!A1+Data3!A1", "Data3", &order), "=Jan!A1+#REF!A1");
        assert_eq!(drop_sheet_in_formula("=SUM(Jan:Data3!A1)", "Data3", &order), "=SUM(Jan:Data2!A1)");
        let names = |list: &[&str]| list.iter().map(|one| one.to_string()).collect::<Vec<_>>();
        let before = names(&["Base", "Qa", "Qb", "Qc", "Qd"]);
        assert_eq!(
            move_sheet_in_formula("=SUM(Qa:Qc!A1)", "Qc", &before, &names(&["Qc", "Base", "Qa", "Qb", "Qd"])),
            "=SUM(Qa:Qb!A1)"
        );
        assert_eq!(
            move_sheet_in_formula("=SUM(Qa:Qb!A1)", "Qb", &names(&["Base", "Qa", "Qb", "Qc", "Qd"]), &names(&["Base", "Qa", "Qc", "Qd", "Qb"])),
            "=SUM(Qa:Qb!A1)"
        );
        assert_eq!(
            move_sheet_in_formula("=SUM(Qa:Qb!A1)", "Qa", &names(&["Base", "Qa", "Qd", "Qc", "Qb"]), &names(&["Base", "Qd", "Qc", "Qb", "Qa"])),
            "=SUM(Qd:Qb!A1)"
        );
    }

    /// A written formula reads back as Excel keeps it. Every pair measured
    /// through `.Formula`, with a sheet Sheet1 and a name MyName.
    #[test]
    fn a_formula_reads_back_as_excel_keeps_it() {
        let sheet = |asked: &str| asked.eq_ignore_ascii_case("Sheet1").then(|| "Sheet1".to_string());
        let name = |asked: &str| asked.eq_ignore_ascii_case("MyName").then(|| "MyName".to_string());
        for (written, kept) in [
            ("=a1+b1", "=A1+B1"),
            ("=sum(a1:b2)", "=SUM(A1:B2)"),
            ("=$a$1*2", "=$A$1*2"),
            ("=sheet1!a1", "=Sheet1!A1"),
            ("=SHEET1!A1", "=Sheet1!A1"),
            ("=Sum( a1 , b1 )", "=SUM( A1, B1 )"),
            ("=a:a", "=A:A"),
            ("=if(true,1,0)", "=IF(TRUE,1,0)"),
            ("=vlookup(a1,a1:b3,2,false)", "=VLOOKUP(A1,A1:B3,2,FALSE)"),
            ("=myname", "=MyName"),
            ("=A1 + B1", "=A1 + B1"),
            ("=sum(A1:B2)*1.50", "=SUM(A1:B2)*1.5"),
            ("=1.50+0.0", "=1.5+0"),
            ("=1e3", "=1000"),
            ("=.5", "=0.5"),
            ("=#n/a", "=#N/A"),
            ("=iferror(1/0,#div/0!)", "=IFERROR(1/0,#DIV/0!)"),
            ("=sheet1!$b$2:$c$3", "=Sheet1!$B$2:$C$3"),
            ("=R2C3+r10c1", "=$C$2+$A$10"),
            ("=a1&\"abc\"", "=A1&\"abc\""),
        ] {
            assert_eq!(canonical_formula(written, &sheet, &name), kept, "{written}");
        }
    }

    /// A legacy formula takes `@` where Excel's `.Formula2` shows it. Every
    /// pair here was measured: written through `.Formula`, read back
    /// through `.Formula2`, with Nm naming B1:B3.
    #[test]
    fn a_legacy_formula_takes_an_at_where_excel_puts_one() {
        let multi = |name: &str| name.eq_ignore_ascii_case("Nm");
        for (written, kept) in [
            ("=A1:A3", "=@A1:A3"),
            ("=A1:A3+1", "=@A1:A3+1"),
            ("=SUM(A1:A3)", "=SUM(A1:A3)"),
            ("=Nm", "=@Nm"),
            ("=INDEX(A1:A3,0)", "=@INDEX(A1:A3,0)"),
            ("=INDEX(A1:A3,2)", "=INDEX(A1:A3,2)"),
            ("=INDEX(A1:A3,MATCH(2,A1:A3,0))", "=INDEX(A1:A3,MATCH(2,A1:A3,0))"),
            ("=SUMPRODUCT(A1:A3*2)", "=SUMPRODUCT(A1:A3*2)"),
            ("=IF(A1:A3>1,1,0)", "=IF(@A1:A3>1,1,0)"),
            ("=IF(A1:A3>1,B1:B3,0)", "=@IF(@A1:A3>1,B1:B3,0)"),
            ("=VLOOKUP(A1:A3,A1:B3,2,0)", "=VLOOKUP(@A1:A3,A1:B3,2,0)"),
            ("=SUM(A1:A3*2)", "=SUM(@A1:A3*2)"),
            ("=COUNTIF(A1:A3,A1:A3)", "=COUNTIF(A1:A3,@A1:A3)"),
            ("=SUMIFS(B1:B3,A1:A3,A1:A3)", "=SUMIFS(B1:B3,A1:A3,@A1:A3)"),
            ("=ROW(A1)", "=ROW(A1)"),
            ("=COLUMN(A1:B1)", "=@COLUMN(A1:B1)"),
            ("=INDIRECT(\"A1\")", "=@INDIRECT(\"A1\")"),
            ("=OFFSET(A1,1,0)", "=OFFSET(A1,1,0)"),
            ("=OFFSET(A1,0,0,2,1)", "=@OFFSET(A1,0,0,2,1)"),
            ("=TRANSPOSE(A1:A3)", "=@TRANSPOSE(@A1:A3)"),
            ("=MAX(IF(A1:A3>1,A1:A3))", "=MAX(IF(@A1:A3>1,A1:A3))"),
            ("=(A1:A3)*1", "=(@A1:A3)*1"),
            ("=$A:$A", "=@$A:$A"),
            ("=1:1", "=@1:1"),
            ("=A1:A1", "=@A1:A1"),
            ("=A1+B1", "=A1+B1"),
            ("=MYOWN(A1:A3)", "=MYOWN(A1:A3)"),
        ] {
            assert_eq!(implied_intersections(written, &multi), kept, "{written}");
            assert_eq!(without_intersections(kept), written, "{kept}");
        }
    }

    /// Read from inside its own table, a formula drops the table's name only
    /// where it names one column. Every answer here is Excel's.
    #[test]
    fn a_formula_in_its_table_drops_the_name_for_one_column() {
        let read = |formula: &str| drop_own_table_name(formula, "tblP");
        assert_eq!(read("=SUM(tblP[Qty])"), "=SUM([Qty])");
        assert_eq!(read("=tblP[[#This Row],[Qty]]"), "=[@Qty]");
        assert_eq!(read("=tblP[@[Qty]:[X]]"), "=tblP[@[Qty]:[X]]");
        assert_eq!(read("=tblP[[#Headers],[Qty]]"), "=tblP[[#Headers],[Qty]]");
        assert_eq!(read("=SUM(tblP[[Qty]:[X]])"), "=SUM(tblP[[Qty]:[X]])");
        assert_eq!(read("=SUM(tblP[#Data])"), "=SUM(tblP[#Data])");
        assert_eq!(read("=SUM(tblP)"), "=SUM(tblP)");
        assert_eq!(normalise_formula_ranges("=SUM(tblP[])"), "=SUM(tblP)");
    }

    /// A structured reference names no cell, so moving it changes nothing,
    /// the table's name included.
    #[test]
    fn a_structured_reference_moves_unchanged() {
        assert_eq!(
            translate_formula_references("=SUM(tblP[@[Unit Price]:[Qty]])+[@Qty]*B2", 1, 0).unwrap(),
            "=SUM(tblP[@[Unit Price]:[Qty]])+[@Qty]*B3"
        );
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

    /// A reference carried off the sheet becomes `#REF!`, as Excel writes
    /// it (measured: `=$A$1+B1` filled two columns to the left is
    /// `=$A$1+#REF!`); a formula that cannot be read is refused.
    #[test]
    fn formula_translation_marks_out_of_bounds_and_rejects_unsupported_formulas() {
        assert_eq!(translate_formula_references("=A1", -1, 0).unwrap(), "=#REF!");
        assert_eq!(translate_formula_references("=$A$1+B1", 0, -2).unwrap(), "=$A$1+#REF!");
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
    use super::{
        move_formula_references, shift_formula_references, transpose_formula_references, CellMove,
        ReferenceShift, ShiftAxis,
    };
    use crate::reference::{MAX_COL, MAX_ROW};

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
                shift_formula_references(
                    formula,
                    &ReferenceShift {
                        axis: ShiftAxis::Rows,
                        at,
                        count,
                        across: (1, MAX_COL + 1),
                        sheet: None,
                        on_sheet: None,
                    },
                )
                .unwrap(),
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
                shift_formula_references(
                    formula,
                    &ReferenceShift {
                        axis: ShiftAxis::Rows,
                        at,
                        count,
                        across: (1, MAX_COL + 1),
                        sheet: None,
                        on_sheet: None,
                    },
                )
                .unwrap(),
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
                shift_formula_references(
                    formula,
                    &ReferenceShift {
                        axis: ShiftAxis::Columns,
                        at,
                        count,
                        across: (1, MAX_ROW + 1),
                        sheet: None,
                        on_sheet: None,
                    },
                )
                .unwrap(),
                expected,
                "{formula} with {count} at column {at}"
            );
        }
    }

    fn whole_rows(at: u32, count: i64, sheet: Option<&str>) -> ReferenceShift<'_> {
        ReferenceShift {
            axis: ShiftAxis::Rows,
            at,
            count,
            across: (1, MAX_COL + 1),
            sheet,
            // These tests rewrite the moved sheet's own formulas, which is
            // what an unspoken `on_sheet` already means.
            on_sheet: None,
        }
    }

    /// An unqualified reference means the sheet the formula is written on.
    ///
    /// Without saying which sheet that is, every unqualified reference moved —
    /// so putting a row into `Data` dragged `=A3` on `Summary` along with it,
    /// pointing it at a row nothing had touched. Found by a test that inserted
    /// into a sheet that did not exist and watched another sheet's formulas
    /// change anyway.
    #[test]
    fn an_unqualified_reference_belongs_to_its_own_sheet() {
        let moving_data = |on: Option<&str>, formula: &str| {
            let mut shift = whole_rows(2, 1, Some("Data"));
            shift.on_sheet = on;
            shift_formula_references(formula, &shift).unwrap()
        };
        // Written on Data: the unqualified one is about Data, so it moves.
        assert_eq!(moving_data(Some("Data"), "=A3+Data!A3+Other!A3"), "=A4+Data!A4+Other!A3");
        // Written anywhere else: the unqualified one is about that sheet.
        assert_eq!(moving_data(Some("Other"), "=A3+Data!A3+Other!A3"), "=A3+Data!A4+Other!A3");
        // Sheet names are matched without regard to capitals, as Excel does.
        assert_eq!(moving_data(Some("DATA"), "=A3"), "=A4");
        // Saying nothing keeps the old meaning: the caller is rewriting the
        // moved sheet's own formulas.
        assert_eq!(moving_data(None, "=A3"), "=A4");
    }

    /// A formula on another sheet follows the rows of the sheet it names.
    /// Measured against Excel after inserting a row on a sheet called Data.
    #[test]
    fn a_reference_follows_the_sheet_it_names() {
        for (formula, expected) in [
            ("=Data!A5*2", "=Data!A6*2"),
            ("=SUM(Data!A1:A6)", "=SUM(Data!A2:A7)"),
            // A sheet the change never touched keeps its references.
            ("=Report!A5*2", "=Report!A5*2"),
            // An unqualified reference belongs to whichever sheet holds it.
            ("=A5*2", "=A6*2"),
        ] {
            assert_eq!(
                shift_formula_references(formula, &whole_rows(1, 1, Some("Data"))).unwrap(),
                expected,
                "{formula} after a row went into Data"
            );
        }
        // Excel writes `=Data!#REF!*2` here, keeping a sheet name that no longer
        // points anywhere. This crate collapses a broken reference to the error
        // value itself, which reads the same when evaluated.
        assert_eq!(
            shift_formula_references("=Data!A5*2", &whole_rows(5, -1, Some("Data"))).unwrap(),
            "=#REF!*2"
        );
    }

    #[test]
    fn a_reference_to_another_sheet_stays_put() {
        assert_eq!(
            shift_formula_references("=Sheet2!A2*2", &whole_rows(1, 1, None)).unwrap(),
            "=Sheet2!A2*2"
        );
    }

    /// Shifting part of a column leaves its neighbours alone. Every expectation
    /// is what Excel 16 left after `Range("B2").Insert` or `.Delete`, with the
    /// band one column wide.
    #[test]
    fn only_references_inside_the_band_move() {
        let column_b = (2, 2);
        for (formula, count, expected) in [
            ("=B3*2", 1, "=B4*2"),
            ("=B2*2", 1, "=B3*2"),
            ("=B1*2", 1, "=B1*2"),
            // A different column is untouched, however close.
            ("=C3*2", 1, "=C3*2"),
            // A range inside the band still grows and shrinks.
            ("=SUM(B1:B4)", 1, "=SUM(B1:B5)"),
            ("=SUM(B1:B4)", -1, "=SUM(B1:B3)"),
            // One reaching past the band is left as it stands.
            ("=SUM(A1:C3)", 1, "=SUM(A1:C3)"),
            ("=SUM(A1:C3)", -1, "=SUM(A1:C3)"),
            ("=B3*2", -1, "=B2*2"),
            ("=B2*2", -1, "=#REF!*2"),
        ] {
            assert_eq!(
                shift_formula_references(
                    formula,
                    &ReferenceShift {
                        axis: ShiftAxis::Rows,
                        at: 2,
                        count,
                        across: column_b,
                        sheet: None,
                        on_sheet: None,
                    },
                )
                .unwrap(),
                expected,
                "{formula} with {count} at row 2 across column B"
            );
        }
    }

    /// The same rule the other way round: shifting part of a row moves what
    /// shares that row and nothing else.
    #[test]
    fn only_references_inside_a_row_band_move() {
        let row_2 = (2, 2);
        for (formula, count, expected) in [
            ("=C2*2", 1, "=D2*2"),
            ("=C3*2", 1, "=C3*2"),
            ("=C2*2", -1, "=B2*2"),
        ] {
            assert_eq!(
                shift_formula_references(
                    formula,
                    &ReferenceShift {
                        axis: ShiftAxis::Columns,
                        at: 2,
                        count,
                        across: row_2,
                        sheet: None,
                        on_sheet: None,
                    },
                )
                .unwrap(),
                expected,
                "{formula} with {count} at column 2 across row 2"
            );
        }
    }

    /// A range on another sheet is stepped over whole.
    ///
    /// Its far end carries no sheet of its own, so judging that end alone
    /// takes it for this sheet's cell and moves half the range.
    #[test]
    fn another_sheets_range_is_left_alone_at_both_ends() {
        let shift = ReferenceShift {
            axis: ShiftAxis::Rows,
            at: 1,
            count: 1,
            across: (1, u32::MAX),
            sheet: Some("Sheet1"),
            on_sheet: Some("Sheet1"),
        };
        assert_eq!(
            shift_formula_references("=SUM(Other!$B$3:$B$5)", &shift).unwrap(),
            "=SUM(Other!$B$3:$B$5)"
        );
        assert_eq!(
            shift_formula_references("=SUM(Other!B3:B5)+SUM(B3:B5)", &shift).unwrap(),
            "=SUM(Other!B3:B5)+SUM(B4:B6)"
        );
        // The sheet that did move still moves, named or not.
        assert_eq!(
            shift_formula_references("=SUM(Sheet1!$B$3:$B$5)", &shift).unwrap(),
            "=SUM(Sheet1!$B$4:$B$6)"
        );
    }

    /// A function's name is not a reference, however much it reads like one.
    #[test]
    fn a_function_name_is_left_alone() {
        assert_eq!(
            shift_formula_references("=LOG10(A2)", &whole_rows(1, 1, None)).unwrap(),
            "=LOG10(A3)"
        );
    }

    /// `A2:B3` cut onto `D2`, which is where every answer below was measured.
    fn cut_a2b3_onto_d2(written_on: Option<&'static str>) -> CellMove<'static> {
        CellMove {
            first_row: 1,
            first_column: 0,
            last_row: 2,
            last_column: 1,
            down: 0,
            across: 3,
            from_sheet: Some("Sheet1"),
            to_sheet: Some("Sheet1"),
            read_as: written_on,
            written_on,
        }
    }

    /// A reference follows the cells a cut took, and one aimed at what the cut
    /// landed on is left with nothing to name. Asked of Excel.
    #[test]
    fn references_follow_the_cells_a_cut_moved() {
        let moved = cut_a2b3_onto_d2(Some("Sheet1"));
        let said = |formula: &str| move_formula_references(formula, &moved).unwrap();

        // Wholly inside the block: it follows, absolute halves included.
        assert_eq!(said("=SUM(A2:B3)"), "=SUM(D2:E3)");
        assert_eq!(said("=SUM(A2:A3)"), "=SUM(D2:D3)");
        assert_eq!(said("=A2+B3"), "=D2+E3");
        assert_eq!(said("=$A$2"), "=$D$2");
        assert_eq!(said("=A2*10"), "=D2*10");
        // Reaching past the block, or naming a whole line: left alone.
        assert_eq!(said("=SUM(A1:B4)"), "=SUM(A1:B4)");
        assert_eq!(said("=SUM(A2:B5)"), "=SUM(A2:B5)");
        assert_eq!(said("=SUM(A:A)"), "=SUM(A:A)");
        assert_eq!(said("=G9"), "=G9");
        // Aimed at what the block landed on.
        assert_eq!(said("=D2"), "=#REF!");
        assert_eq!(said("=SUM(D2:E3)"), "=SUM(#REF!)");
        assert_eq!(said("=D2+D4"), "=#REF!+D4");
        assert_eq!(said("=$D$3"), "=#REF!");
        assert_eq!(said("=SUM(D4:D6)"), "=SUM(D4:D6)");
    }

    /// A range the block landed on the END of closes up to just before it.
    ///
    /// Every answer was asked of Excel, cutting a block onto D2 so that
    /// D2:E3 is what gets written over. It closes up only where the block
    /// reaches PAST the range's end: `D1:D2` becomes `D1:D1`, while `D1:D3` —
    /// which ends where the block does — is left alone.
    #[test]
    fn a_range_the_block_landed_on_the_end_of_closes_up() {
        let moved = cut_a2b3_onto_d2(Some("Sheet1"));
        let said = |formula: &str| move_formula_references(formula, &moved).unwrap();

        assert_eq!(said("=SUM(D1:D2)"), "=SUM(D1:D1)");
        assert_eq!(said("=SUM(C2:D3)"), "=SUM(C2:C3)");
        // The block has to reach past the end, not merely up to it.
        assert_eq!(said("=SUM(D1:D3)"), "=SUM(D1:D3)");
        assert_eq!(said("=SUM(D1:E3)"), "=SUM(D1:E3)");
        // Landing on the near end, or in the middle, changes nothing.
        assert_eq!(said("=SUM(D2:D5)"), "=SUM(D2:D5)");
        assert_eq!(said("=SUM(D1:D4)"), "=SUM(D1:D4)");
        assert_eq!(said("=SUM(D2:F3)"), "=SUM(D2:F3)");
        // And what it covers entirely still has nothing left to name.
        assert_eq!(said("=SUM(D2:E3)"), "=SUM(#REF!)");
    }

    /// The cut reaches another sheet's formulas, but only where they name the
    /// sheet the cells moved on.
    #[test]
    fn a_cut_reaches_the_formulas_on_other_sheets() {
        let elsewhere = cut_a2b3_onto_d2(Some("Second"));
        assert_eq!(
            move_formula_references("=Sheet1!A2", &elsewhere).unwrap(),
            "=Sheet1!D2"
        );
        assert_eq!(
            move_formula_references("=SUM(Sheet1!A2:B3)", &elsewhere).unwrap(),
            "=SUM(Sheet1!D2:E3)"
        );
        // Unqualified on another sheet means that sheet's own A2, untouched.
        assert_eq!(move_formula_references("=A2", &elsewhere).unwrap(), "=A2");
    }

    /// A formula turned a quarter turn, as Excel turns one.
    ///
    /// Every answer is what Excel left in the cell after a transposed paste of
    /// a formula written in C3, which is (2, 2) counting from zero.
    #[test]
    fn a_transposed_formula_looks_the_other_way() {
        let from = (2, 2);
        let turned = |formula: &str, to| transpose_formula_references(formula, from, to).unwrap();

        // One to the left becomes one above, and one above becomes one left.
        assert_eq!(turned("=C2*2", (1, 5)), "=E2*2");
        assert_eq!(turned("=B3*2", (0, 5)), "=#REF!*2");
        // A row of cells comes out as a column of them.
        assert_eq!(turned("=SUM(A3:B3)", (5, 5)), "=SUM(F4:F5)");
        assert_eq!(turned("=SUM(A1:B2)", (8, 8)), "=SUM(G7:H8)");
        // Far away turns as far.
        assert_eq!(turned("=Z9", (6, 5)), "=L30");
        // The cell itself stays the cell itself.
        assert_eq!(turned("=C3", (9, 9)), "=J10");
        // What names a fixed cell, or half of one, does not turn.
        assert_eq!(turned("=$A$1", (2, 5)), "=$A$1");
        assert_eq!(turned("=B$3", (3, 5)), "=B$3");
        assert_eq!(turned("=$B3", (4, 5)), "=$B3");
        // And what names no cell at all is left alone.
        assert_eq!(turned("=1+1", (7, 5)), "=1+1");
        assert_eq!(
            turned("=SUM(A3:B3)+LOG10(100)", (5, 5)),
            "=SUM(F4:F5)+LOG10(100)"
        );
    }

    /// A cut onto another sheet takes the references there too, and they have
    /// to say so wherever that is no longer the sheet they sit on.
    ///
    /// Measured with `Sheet3!A2:B3` cut onto `Sheet2!D2`: the watcher on
    /// Sheet3 reads `=Sheet2!D2`, a third sheet's `=Sheet3!A2` reads
    /// `=Sheet2!D2`, and of the formulas that travelled, one naming a cell
    /// that came with them reads `=D2*10` while one naming a neighbour left
    /// behind reads `=Sheet3!G9`.
    #[test]
    fn a_cut_onto_another_sheet_carries_the_sheet_name_too() {
        let across_sheets = |read_as, written_on| CellMove {
            first_row: 1,
            first_column: 0,
            last_row: 2,
            last_column: 1,
            down: 0,
            across: 3,
            from_sheet: Some("Sheet3"),
            to_sheet: Some("Sheet2"),
            read_as: Some(read_as),
            written_on: Some(written_on),
        };

        // Watching from the sheet the cells left.
        let watcher = across_sheets("Sheet3", "Sheet3");
        assert_eq!(
            move_formula_references("=A2", &watcher).unwrap(),
            "=Sheet2!D2"
        );
        assert_eq!(
            move_formula_references("=$A$2", &watcher).unwrap(),
            "=Sheet2!$D$2"
        );
        assert_eq!(
            move_formula_references("=SUM(A2:B3)", &watcher).unwrap(),
            "=SUM(Sheet2!D2:E3)"
        );

        // Watching from a third sheet, and from the sheet they landed on.
        let bystander = across_sheets("Sheet1", "Sheet1");
        assert_eq!(
            move_formula_references("=Sheet3!A2", &bystander).unwrap(),
            "=Sheet2!D2"
        );
        let landed = across_sheets("Sheet2", "Sheet2");
        assert_eq!(
            move_formula_references("=Sheet3!A2", &landed).unwrap(),
            "=Sheet2!D2"
        );
        // A cell the block landed on has nothing left to name.
        assert_eq!(move_formula_references("=D2", &landed).unwrap(), "=#REF!");

        // The formulas that travelled: read against the sheet they came from,
        // written against the one they sit on now.
        let carried = across_sheets("Sheet3", "Sheet2");
        assert_eq!(
            move_formula_references("=A2*10", &carried).unwrap(),
            "=D2*10"
        );
        assert_eq!(
            move_formula_references("=G9", &carried).unwrap(),
            "=Sheet3!G9"
        );
        assert_eq!(
            move_formula_references("=Sheet1!A1", &carried).unwrap(),
            "=Sheet1!A1"
        );
    }

    /// Every pair here is what Excel stored for what VBA's `.Formula` wrote.
    #[test]
    fn a_range_is_stored_top_left_first() {
        for (written, stored) in [
            ("=SUM(B6:A5)", "=SUM(A5:B6)"),
            ("=SUM($A6:A$5)", "=SUM($A$5:A6)"),
            ("=SUM(A6:B5)", "=SUM(A5:B6)"),
            ("=SUM(Sheet1!A6:A5)", "=SUM(Sheet1!A5:A6)"),
            ("=SUM(6:5)", "=SUM(5:6)"),
            ("=SUM(C:A)", "=SUM(A:C)"),
            ("=SUM(B$6:$A5)", "=SUM($A5:B$6)"),
            ("=SUM(A1:B2)+1.50", "=SUM(A1:B2)+1.50"),
        ] {
            assert_eq!(super::normalise_formula_ranges(written), stored, "{written}");
        }
    }
}
