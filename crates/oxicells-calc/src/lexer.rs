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
    text.parse::<f64>()
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
}
