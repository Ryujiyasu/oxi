// This Source Code Form is subject to the terms of the Mozilla Public
// License, v. 2.0. If a copy of the MPL was not distributed with this
// file, You can obtain one at https://mozilla.org/MPL/2.0/.

//! Excel's other way of writing a reference, and the way back.
//!
//! `A1` names a cell by where it is; `R1C1` names it by where it is *from the
//! formula*, so `=RC[-1]*2` means the same thing in every cell of a column.
//! That is why a macro writes it: one string fills a block correctly.
//!
//! The two directions are not mirror images, and Excel was asked about each:
//!
//! - Reading, an offset is the plain difference. A formula in A20 pointing at
//!   XFD20 reads `=RC[16383]`, not `=RC[-1]`.
//! - Writing, an offset wraps around the sheet. `=RC[-1]` written into A1
//!   lands on XFD1, and `=RC[1]` written into XFD5 lands on A5. An offset as
//!   long as the sheet is refused.
//! - A whole-line range collapses when its two ends agree: `$A:$A` reads as
//!   `C1`, `A:A` as `C[-1]`, and `2:2` as `R[-6]` from row 8. An ordinary
//!   range does not — `A1:A1` reads as `R[-1]C[-1]:R[-1]C[-1]`.

use crate::lexer::{render_token, tokenize, Token};
use crate::reference::{col_to_letters, letters_to_col, parse_a1, CellRef, MAX_COL, MAX_ROW};

/// How many rows and columns a sheet has, which is what an offset wraps by.
const ROW_COUNT: i64 = MAX_ROW as i64 + 1;
const COLUMN_COUNT: i64 = MAX_COL as i64 + 1;

/// One end of a whole-line range: a bare column or a bare row, and whether it
/// was written with a `$`.
#[derive(Clone, Copy, PartialEq, Eq)]
enum Line {
    Column { index: u32, absolute: bool },
    Row { index: u32, absolute: bool },
}

/// The R1C1 halves of one reference, before they are placed on a sheet.
#[derive(Clone, Copy)]
struct R1C1 {
    row: Option<Axis>,
    column: Option<Axis>,
}

#[derive(Clone, Copy)]
enum Axis {
    /// A bare number: the line itself, counted from one.
    Absolute(u32),
    /// A bracketed number: how far from the formula's own line.
    Relative(i64),
}

/// Write a formula the way Excel shows it in R1C1 style, as seen from the cell
/// holding it. `row` and `column` are zero-based.
pub fn formula_to_r1c1(formula: &str, row: u32, column: u32) -> Result<String, String> {
    let tokens = tokenize(formula).map_err(|error| error.to_string())?;
    let mut output = String::new();
    if formula.trim_start().starts_with('=') {
        output.push('=');
    }

    let mut index = 0;
    while index < tokens.len() {
        if tokens.get(index + 1) == Some(&Token::Colon) {
            if let (Some(first), Some(last)) = (
                line_of(&tokens[index]),
                tokens.get(index + 2).and_then(line_of),
            ) {
                if let (Some(left), Some(right)) = (
                    render_line(first.1, row, column),
                    render_line(last.1, row, column),
                ) {
                    let sheets_agree = first.0 == last.0;
                    render_token(
                        &mut output,
                        Token::Name {
                            sheet: first.0,
                            name: left.clone(),
                        },
                    );
                    if left != right || !sheets_agree {
                        output.push(':');
                        render_token(
                            &mut output,
                            Token::Name {
                                sheet: last.0,
                                name: right,
                            },
                        );
                    }
                    index += 3;
                    continue;
                }
            }
        }

        let token = tokens[index].clone();
        match &token {
            // A name followed by `(` is a function, which is how `LOG10` keeps
            // its own shape.
            Token::Name { sheet, name } if tokens.get(index + 1) != Some(&Token::LParen) => {
                match parse_a1(name) {
                    Some(cell) => render_token(
                        &mut output,
                        Token::Name {
                            sheet: sheet.clone(),
                            name: cell_to_r1c1(cell, row, column),
                        },
                    ),
                    None => render_token(&mut output, token),
                }
            }
            _ => render_token(&mut output, token),
        }
        index += 1;
    }
    Ok(output)
}

/// Read a formula written in R1C1 style into the A1 style the file keeps.
/// `row` and `column` are zero-based.
pub fn formula_from_r1c1(formula: &str, row: u32, column: u32) -> Result<String, String> {
    let bytes = formula.as_bytes();
    let mut output = String::new();
    let mut index = 0;

    while index < bytes.len() {
        let character = bytes[index] as char;

        if character == '"' || character == '\'' {
            index = copy_quoted(formula, index, &mut output).ok_or_else(|| {
                format!("unterminated {character} in the R1C1 formula {formula:?}")
            })?;
            continue;
        }

        if !is_name_start(character) {
            output.push(character);
            index += 1;
            continue;
        }

        if let Some((reference, end)) = match_reference(formula, index) {
            if ends_a_reference(bytes.get(end).copied()) {
                let (text, end) = render_reference(formula, reference, end, row, column)?;
                output.push_str(&text);
                index = end;
                continue;
            }
        }

        // Not a reference: a function name, a defined name, a sheet name.
        let start = index;
        while index < bytes.len() && is_name_body(bytes[index] as char) {
            index += 1;
        }
        output.push_str(&formula[start..index]);
    }
    Ok(output)
}

fn cell_to_r1c1(cell: CellRef, row: u32, column: u32) -> String {
    format!(
        "{}{}",
        axis_to_r1c1('R', cell.row, cell.row_absolute, row),
        axis_to_r1c1('C', cell.col, cell.col_absolute, column)
    )
}

fn axis_to_r1c1(letter: char, target: u32, absolute: bool, current: u32) -> String {
    if absolute {
        return format!("{letter}{}", target + 1);
    }
    match i64::from(target) - i64::from(current) {
        0 => letter.to_string(),
        offset => format!("{letter}[{offset}]"),
    }
}

/// A bare column letter or row number, as one end of `A:C` or `1:3`. A bare row
/// number reaches the tokeniser as a number and a bare column as a name, and
/// either may carry a `$` that means nothing to a whole line but must survive.
fn line_of(token: &Token) -> Option<(Option<String>, Line)> {
    match token {
        Token::Number(one) => {
            let rounded = *one as u32;
            (one.fract() == 0.0 && *one >= 1.0 && i64::from(rounded) <= ROW_COUNT).then_some((
                None,
                Line::Row {
                    index: rounded - 1,
                    absolute: false,
                },
            ))
        }
        Token::Name { sheet, name } => {
            let absolute = name.starts_with('$');
            let bare = name.strip_prefix('$').unwrap_or(name);
            if bare.is_empty() {
                return None;
            }
            let line = if bare
                .chars()
                .all(|character| character.is_ascii_alphabetic())
            {
                Line::Column {
                    index: letters_to_col(bare)?,
                    absolute,
                }
            } else if bare.chars().all(|character| character.is_ascii_digit()) {
                let number: u32 = bare.parse().ok()?;
                if number == 0 || i64::from(number) > ROW_COUNT {
                    return None;
                }
                Line::Row {
                    index: number - 1,
                    absolute,
                }
            } else {
                return None;
            };
            Some((sheet.clone(), line))
        }
        _ => None,
    }
}

fn render_line(line: Line, row: u32, column: u32) -> Option<String> {
    Some(match line {
        Line::Column { index, absolute } => axis_to_r1c1('C', index, absolute, column),
        Line::Row { index, absolute } => axis_to_r1c1('R', index, absolute, row),
    })
}

fn is_name_start(character: char) -> bool {
    character.is_ascii_alphabetic() || character == '_' || character == '$'
}

fn is_name_body(character: char) -> bool {
    character.is_ascii_alphanumeric() || character == '_' || character == '.' || character == '$'
}

/// What may stand right after a reference. A letter or digit would make it part
/// of a longer name, a `[` would make it a table, a `!` would make it a sheet,
/// and a `(` would make it a function.
fn ends_a_reference(next: Option<u8>) -> bool {
    match next {
        None => true,
        Some(byte) => {
            let character = byte as char;
            !(is_name_body(character) || character == '[' || character == '!' || character == '(')
        }
    }
}

/// Copy a string literal or a quoted sheet name through untouched, doubled
/// quote marks and all. Returns where it ended.
fn copy_quoted(source: &str, start: usize, output: &mut String) -> Option<usize> {
    let bytes = source.as_bytes();
    let mark = bytes[start];
    let mut index = start + 1;
    output.push(mark as char);
    while index < bytes.len() {
        if bytes[index] == mark {
            if bytes.get(index + 1) == Some(&mark) {
                output.push(mark as char);
                output.push(mark as char);
                index += 2;
                continue;
            }
            output.push(mark as char);
            return Some(index + 1);
        }
        let character = source[index..].chars().next()?;
        output.push(character);
        index += character.len_utf8();
    }
    None
}

fn match_reference(source: &str, start: usize) -> Option<(R1C1, usize)> {
    let bytes = source.as_bytes();
    let mut index = start;
    let mut row = None;
    if matches!(bytes.get(index), Some(b'R' | b'r')) {
        index += 1;
        row = Some(read_axis(bytes, &mut index)?);
    }
    let mut column = None;
    if matches!(bytes.get(index), Some(b'C' | b'c')) {
        index += 1;
        column = Some(read_axis(bytes, &mut index)?);
    }
    if row.is_none() && column.is_none() {
        return None;
    }
    Some((R1C1 { row, column }, index))
}

fn read_axis(bytes: &[u8], index: &mut usize) -> Option<Axis> {
    if bytes.get(*index) == Some(&b'[') {
        let mut scan = *index + 1;
        let negative = bytes.get(scan) == Some(&b'-');
        if negative || bytes.get(scan) == Some(&b'+') {
            scan += 1;
        }
        let digits = scan;
        while matches!(bytes.get(scan), Some(byte) if byte.is_ascii_digit()) {
            scan += 1;
        }
        if scan == digits || bytes.get(scan) != Some(&b']') {
            return None;
        }
        let magnitude: i64 = std::str::from_utf8(&bytes[digits..scan])
            .ok()?
            .parse()
            .ok()?;
        *index = scan + 1;
        return Some(Axis::Relative(if negative {
            -magnitude
        } else {
            magnitude
        }));
    }
    if matches!(bytes.get(*index), Some(byte) if byte.is_ascii_digit()) {
        let digits = *index;
        while matches!(bytes.get(*index), Some(byte) if byte.is_ascii_digit()) {
            *index += 1;
        }
        let number: u32 = std::str::from_utf8(&bytes[digits..*index])
            .ok()?
            .parse()
            .ok()?;
        if number == 0 {
            return None;
        }
        return Some(Axis::Absolute(number));
    }
    Some(Axis::Relative(0))
}

/// Place one matched reference on the sheet, joining it with the one after the
/// colon when both name a whole line.
fn render_reference(
    source: &str,
    reference: R1C1,
    end: usize,
    row: u32,
    column: u32,
) -> Result<(String, usize), String> {
    match (reference.row, reference.column) {
        (Some(down), Some(across)) => {
            let cell = CellRef {
                col: place(across, column, COLUMN_COUNT, "column")?,
                row: place(down, row, ROW_COUNT, "row")?,
                col_absolute: matches!(across, Axis::Absolute(_)),
                row_absolute: matches!(down, Axis::Absolute(_)),
            };
            Ok((cell.to_a1(), end))
        }
        (Some(down), None) => {
            let (second, end) = paired_line(source, end, |other| {
                other.row.filter(|_| other.column.is_none())
            });
            let first = row_line(down, row)?;
            let last = match second {
                Some(other) => row_line(other, row)?,
                None => first.clone(),
            };
            Ok((format!("{first}:{last}"), end))
        }
        (None, Some(across)) => {
            let (second, end) = paired_line(source, end, |other| {
                other.column.filter(|_| other.row.is_none())
            });
            let first = column_line(across, column)?;
            let last = match second {
                Some(other) => column_line(other, column)?,
                None => first.clone(),
            };
            Ok((format!("{first}:{last}"), end))
        }
        (None, None) => Err("an R1C1 reference names neither a row nor a column".to_string()),
    }
}

/// A whole-line reference reaches to a second one across a colon: `R1:R3` is
/// one range, not two rows standing side by side.
fn paired_line(
    source: &str,
    end: usize,
    pick: impl Fn(&R1C1) -> Option<Axis>,
) -> (Option<Axis>, usize) {
    if source.as_bytes().get(end) != Some(&b':') {
        return (None, end);
    }
    let Some((next, next_end)) = match_reference(source, end + 1) else {
        return (None, end);
    };
    if !ends_a_reference(source.as_bytes().get(next_end).copied()) {
        return (None, end);
    }
    match pick(&next) {
        Some(axis) => (Some(axis), next_end),
        None => (None, end),
    }
}

fn row_line(axis: Axis, row: u32) -> Result<String, String> {
    let index = place(axis, row, ROW_COUNT, "row")?;
    Ok(format!(
        "{}{}",
        if matches!(axis, Axis::Absolute(_)) {
            "$"
        } else {
            ""
        },
        index + 1
    ))
}

fn column_line(axis: Axis, column: u32) -> Result<String, String> {
    let index = place(axis, column, COLUMN_COUNT, "column")?;
    Ok(format!(
        "{}{}",
        if matches!(axis, Axis::Absolute(_)) {
            "$"
        } else {
            ""
        },
        col_to_letters(index)
    ))
}

/// Where one half of an R1C1 reference lands, zero-based.
///
/// A relative offset wraps: Excel accepts `=RC[-1]` in A1 and puts XFD1 there.
/// An offset as long as the sheet is refused, which is what Excel does with
/// `=R[1048576]C`.
fn place(axis: Axis, current: u32, count: i64, what: &str) -> Result<u32, String> {
    match axis {
        Axis::Absolute(number) => {
            let index = i64::from(number) - 1;
            if index >= count {
                return Err(format!("R1C1 {what} {number} is outside the worksheet"));
            }
            Ok(index as u32)
        }
        Axis::Relative(offset) => {
            if offset.abs() >= count {
                return Err(format!(
                    "an R1C1 {what} offset of {offset} is longer than the worksheet"
                ));
            }
            Ok((i64::from(current) + offset).rem_euclid(count) as u32)
        }
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    /// Every answer here was read off Excel through COM.
    #[test]
    fn writes_a1_formulas_the_way_excel_shows_them() {
        // B2, B3, B4, B5 all point at the top-left corner in different ways.
        assert_eq!(formula_to_r1c1("=A1*2", 1, 1).unwrap(), "=R[-1]C[-1]*2");
        assert_eq!(formula_to_r1c1("=$A$1", 2, 1).unwrap(), "=R1C1");
        assert_eq!(formula_to_r1c1("=$A1", 3, 1).unwrap(), "=R[-3]C1");
        assert_eq!(formula_to_r1c1("=A$1", 4, 1).unwrap(), "=R1C[-1]");
        assert_eq!(
            formula_to_r1c1("=SUM(A1:A3)", 5, 1).unwrap(),
            "=SUM(R[-5]C[-1]:R[-3]C[-1])"
        );
        assert_eq!(
            formula_to_r1c1("=SUM($A1:A$3)", 8, 1).unwrap(),
            "=SUM(R[-8]C1:R3C[-1])"
        );
        // A whole line collapses when its ends agree, and only then.
        assert_eq!(formula_to_r1c1("=SUM(A:A)", 6, 1).unwrap(), "=SUM(C[-1])");
        assert_eq!(
            formula_to_r1c1("=SUM(A:C)", 6, 1).unwrap(),
            "=SUM(C[-1]:C[1])"
        );
        assert_eq!(formula_to_r1c1("=SUM(2:2)", 7, 1).unwrap(), "=SUM(R[-6])");
        assert_eq!(
            formula_to_r1c1("=SUM(1:3)", 7, 1).unwrap(),
            "=SUM(R[-7]:R[-5])"
        );
        assert_eq!(formula_to_r1c1("=SUM($A:$A)", 13, 1).unwrap(), "=SUM(C1)");
        assert_eq!(formula_to_r1c1("=SUM($1:$1)", 14, 1).unwrap(), "=SUM(R1)");
        assert_eq!(formula_to_r1c1("=SUM($A:B)", 4, 1).unwrap(), "=SUM(C1:C)");
        assert_eq!(
            formula_to_r1c1("=SUM(A1:A1)", 1, 1).unwrap(),
            "=SUM(R[-1]C[-1]:R[-1]C[-1])"
        );
        // Sheets, function names, text and error values keep their shape.
        assert_eq!(
            formula_to_r1c1("=Second!A1", 8, 1).unwrap(),
            "=Second!R[-8]C[-1]"
        );
        assert_eq!(
            formula_to_r1c1("='My Sheet'!A1", 9, 1).unwrap(),
            "='My Sheet'!R[-9]C[-1]"
        );
        assert_eq!(formula_to_r1c1("=LOG10(100)", 9, 1).unwrap(), "=LOG10(100)");
        assert_eq!(
            formula_to_r1c1("=\"A1 is \"&A1", 10, 1).unwrap(),
            "=\"A1 is \"&R[-10]C[-1]"
        );
        assert_eq!(formula_to_r1c1("=1/0", 11, 1).unwrap(), "=1/0");
        // Reading never wraps: the offset is the plain difference.
        assert_eq!(formula_to_r1c1("=XFD20", 19, 0).unwrap(), "=RC[16383]");
        assert_eq!(formula_to_r1c1("=A1048576", 20, 0).unwrap(), "=R[1048555]C");
    }

    #[test]
    fn reads_r1c1_formulas_the_way_excel_places_them() {
        assert_eq!(formula_from_r1c1("=RC[-3]*2", 0, 3).unwrap(), "=A1*2");
        assert_eq!(formula_from_r1c1("=R1C1", 4, 3).unwrap(), "=$A$1");
        assert_eq!(formula_from_r1c1("=R1C[-3]", 5, 3).unwrap(), "=A$1");
        assert_eq!(formula_from_r1c1("=R[-6]C1", 6, 3).unwrap(), "=$A1");
        assert_eq!(
            formula_from_r1c1("=SUM(R[-7]C[-3]:R[-5]C[-3])", 7, 3).unwrap(),
            "=SUM(A1:A3)"
        );
        assert_eq!(formula_from_r1c1("=SUM(C[-3])", 8, 3).unwrap(), "=SUM(A:A)");
        assert_eq!(formula_from_r1c1("=SUM(R[-9])", 9, 3).unwrap(), "=SUM(1:1)");
        assert_eq!(
            formula_from_r1c1("=SUM(C[-3]:C[-1])", 0, 3).unwrap(),
            "=SUM(A:C)"
        );
        assert_eq!(
            formula_from_r1c1("=SUM(C1:C3)", 1, 3).unwrap(),
            "=SUM($A:$C)"
        );
        assert_eq!(
            formula_from_r1c1("=SUM(R[-2]:R[-1])", 2, 3).unwrap(),
            "=SUM(1:2)"
        );
        assert_eq!(
            formula_from_r1c1("=SUM(R1:R3)", 3, 3).unwrap(),
            "=SUM($1:$3)"
        );
        assert_eq!(
            formula_from_r1c1("=SUM(R1C[-3]:R[-1]C)", 4, 3).unwrap(),
            "=SUM(A$1:D4)"
        );
        // A bare R or C is the formula's own line.
        assert_eq!(formula_from_r1c1("=SUM(R)", 8, 0).unwrap(), "=SUM(9:9)");
        assert_eq!(formula_from_r1c1("=SUM(C)", 9, 0).unwrap(), "=SUM(A:A)");
        assert_eq!(formula_from_r1c1("=RC", 15, 3).unwrap(), "=D16");
        assert_eq!(formula_from_r1c1("=R[0]C[0]", 13, 3).unwrap(), "=D14");
        assert_eq!(formula_from_r1c1("=rc[-3]*2", 14, 3).unwrap(), "=A15*2");
        assert_eq!(
            formula_from_r1c1("='My Sheet'!R1C1", 5, 3).unwrap(),
            "='My Sheet'!$A$1"
        );
        assert_eq!(
            formula_from_r1c1("=Sheet1!R1C1", 17, 3).unwrap(),
            "=Sheet1!$A$1"
        );
        // Writing wraps around the sheet.
        assert_eq!(formula_from_r1c1("=RC[-1]", 0, 0).unwrap(), "=XFD1");
        assert_eq!(formula_from_r1c1("=RC[-2]", 1, 0).unwrap(), "=XFC2");
        assert_eq!(formula_from_r1c1("=R[-3]C", 2, 0).unwrap(), "=A1048576");
        assert_eq!(formula_from_r1c1("=RC[1]", 4, 16383).unwrap(), "=A5");
        assert_eq!(formula_from_r1c1("=RC[16383]", 0, 0).unwrap(), "=XFD1");
        assert_eq!(formula_from_r1c1("=R[1048575]C", 2, 0).unwrap(), "=A2");
        // An offset as long as the sheet is refused, as Excel refuses it.
        assert!(formula_from_r1c1("=RC[16384]", 1, 0).is_err());
        assert!(formula_from_r1c1("=R[1048576]C", 3, 0).is_err());
        // Names that merely start with R or C are left alone.
        assert_eq!(
            formula_from_r1c1("=ROUND(RC[-1],2)", 0, 1).unwrap(),
            "=ROUND(A1,2)"
        );
        assert_eq!(
            formula_from_r1c1("=LOG10(100)", 0, 1).unwrap(),
            "=LOG10(100)"
        );
        assert_eq!(
            formula_from_r1c1("=\"RC[-1]\"&RC[-1]", 0, 1).unwrap(),
            "=\"RC[-1]\"&A1"
        );
    }

    #[test]
    fn carries_a_formula_out_and_back() {
        for formula in [
            "=A1*2",
            "=$A$1+B2",
            "=SUM(A1:B3)",
            "=SUM(A:A)",
            "=SUM(A:C)",
            "=SUM(3:5)",
            "=Second!A1",
            "=IF(A1>2,\"yes\",B1)",
        ] {
            let there = formula_to_r1c1(formula, 7, 4).unwrap();
            let back = formula_from_r1c1(&there, 7, 4).unwrap();
            assert_eq!(back, formula, "round trip of {formula} through {there}");
        }
    }
}
