// This Source Code Form is subject to the terms of the Mozilla Public
// License, v. 2.0. If a copy of the MPL was not distributed with this
// file, You can obtain one at https://mozilla.org/MPL/2.0/.

//! Recursive-descent parser using Excel's operator precedence.
//!
//! Excel's precedence has two traps that a conventional expression parser gets
//! wrong, and both are reproduced here deliberately:
//!
//! 1. **Unary minus binds tighter than `^`.** `-2^2` is `4` in Excel, not `-4`.
//! 2. **`^` is left-associative.** `2^3^2` is `64` in Excel, not `512`.
//!
//! Precedence, tightest first:
//! `:` › unary `-`/`+` › postfix `%` › `^` › `*` `/` › `+` `-` › `&` › comparisons

use crate::ast::{BinaryOp, Expr, UnaryOp};
use crate::lexer::{tokenize, ParseError, Token};
use crate::reference::{
    letters_to_col, parse_a1, CellRef, RangeRef, Reference, MAX_COL, MAX_ROW,
};
use crate::value::Value;

/// Parse a formula, with or without its leading `=`.
pub fn parse(input: &str) -> Result<Expr, ParseError> {
    let tokens = tokenize(input)?;
    let mut parser = Parser { tokens, pos: 0 };
    let expr = parser.parse_comparison()?;
    if parser.pos < parser.tokens.len() {
        return Err(ParseError::TrailingInput(format!("{:?}", parser.peek())));
    }
    Ok(expr)
}

/// One side of a whole-line range: a column named by its letters, or a row
/// named by its number. A `$` means nothing here — a whole column is the whole
/// column however it was written down.
enum Line {
    Column(u32),
    Row(u32),
}

impl Line {
    fn of(name: &str) -> Option<Line> {
        let bare = name.trim_start_matches('$');
        // A row number reaches the parser as a NAME when the range names a
        // sheet — `Data!1:1` lexes as a name of "1" — and as a number when it
        // does not. Both are the same row.
        if let Some(col) = letters_to_col(bare) {
            return Some(Line::Column(col));
        }
        bare.parse::<f64>().ok().and_then(Line::of_row)
    }

    fn of_row(one: f64) -> Option<Line> {
        (one >= 1.0 && one.fract() == 0.0 && one <= MAX_ROW as f64 + 1.0)
            .then(|| Line::Row(one as u32 - 1))
    }
}

struct Parser {
    tokens: Vec<Token>,
    pos: usize,
}

impl Parser {
    fn peek(&self) -> Option<&Token> {
        self.tokens.get(self.pos)
    }

    fn bump(&mut self) -> Option<Token> {
        let tok = self.tokens.get(self.pos).cloned();
        if tok.is_some() {
            self.pos += 1;
        }
        tok
    }

    fn eat(&mut self, want: &Token) -> bool {
        if self.peek() == Some(want) {
            self.pos += 1;
            true
        } else {
            false
        }
    }

    fn parse_comparison(&mut self) -> Result<Expr, ParseError> {
        let mut lhs = self.parse_concat()?;
        loop {
            let op = match self.peek() {
                Some(Token::Eq) => BinaryOp::Eq,
                Some(Token::Ne) => BinaryOp::Ne,
                Some(Token::Lt) => BinaryOp::Lt,
                Some(Token::Le) => BinaryOp::Le,
                Some(Token::Gt) => BinaryOp::Gt,
                Some(Token::Ge) => BinaryOp::Ge,
                _ => return Ok(lhs),
            };
            self.pos += 1;
            let rhs = self.parse_concat()?;
            lhs = binary(op, lhs, rhs);
        }
    }

    fn parse_concat(&mut self) -> Result<Expr, ParseError> {
        let mut lhs = self.parse_additive()?;
        while self.eat(&Token::Amp) {
            let rhs = self.parse_additive()?;
            lhs = binary(BinaryOp::Concat, lhs, rhs);
        }
        Ok(lhs)
    }

    fn parse_additive(&mut self) -> Result<Expr, ParseError> {
        let mut lhs = self.parse_multiplicative()?;
        loop {
            let op = match self.peek() {
                Some(Token::Plus) => BinaryOp::Add,
                Some(Token::Minus) => BinaryOp::Sub,
                _ => return Ok(lhs),
            };
            self.pos += 1;
            let rhs = self.parse_multiplicative()?;
            lhs = binary(op, lhs, rhs);
        }
    }

    fn parse_multiplicative(&mut self) -> Result<Expr, ParseError> {
        let mut lhs = self.parse_power()?;
        loop {
            let op = match self.peek() {
                Some(Token::Star) => BinaryOp::Mul,
                Some(Token::Slash) => BinaryOp::Div,
                _ => return Ok(lhs),
            };
            self.pos += 1;
            let rhs = self.parse_power()?;
            lhs = binary(op, lhs, rhs);
        }
    }

    /// Left-associative, matching Excel: `2^3^2` == `(2^3)^2` == 64.
    fn parse_power(&mut self) -> Result<Expr, ParseError> {
        let mut lhs = self.parse_percent()?;
        while self.eat(&Token::Caret) {
            let rhs = self.parse_percent()?;
            lhs = binary(BinaryOp::Pow, lhs, rhs);
        }
        Ok(lhs)
    }

    fn parse_percent(&mut self) -> Result<Expr, ParseError> {
        let mut expr = self.parse_unary()?;
        while self.eat(&Token::Percent) {
            expr = Expr::Unary {
                op: UnaryOp::Percent,
                operand: Box::new(expr),
            };
        }
        Ok(expr)
    }

    /// Sits *inside* `parse_percent` and *below* `parse_power`, which is what
    /// makes `-2^2` evaluate to `4`.
    fn parse_unary(&mut self) -> Result<Expr, ParseError> {
        match self.peek() {
            Some(Token::Minus) => {
                self.pos += 1;
                Ok(Expr::Unary {
                    op: UnaryOp::Neg,
                    operand: Box::new(self.parse_unary()?),
                })
            }
            Some(Token::Plus) => {
                self.pos += 1;
                Ok(Expr::Unary {
                    op: UnaryOp::Plus,
                    operand: Box::new(self.parse_unary()?),
                })
            }
            _ => self.parse_range(),
        }
    }

    /// `A:A` and `3:3`, as a range covering the whole column or the whole row.
    ///
    /// A bare column letter reaches the parser as a name, since `A` on its own
    /// is not a cell; a bare row number reaches it as a number. Excel reads
    /// `A:A` as the column whether or not something called `A` is also
    /// defined, so the shape decides and the name does not get a say.
    /// `A:A` and `3:3`, read straight off the tokens.
    ///
    /// It has to happen before the atoms are parsed, because that is where the
    /// sheet is lost: `Sheet1!$D` becomes a bare name and the qualifier is
    /// dropped on purpose, a defined name belonging to the workbook rather
    /// than to any sheet. Reading it afterwards built `Sheet1!$D:$D` as a
    /// range on whatever sheet the formula sat on, which found nothing and
    /// said so with a confident nought.
    ///
    /// Excel reads `A:A` as the column whether or not something called `A` is
    /// also defined, so the shape decides and the name gets no say.
    fn whole_line(&mut self) -> Option<Expr> {
        let (sheet, first) = match self.peek()? {
            Token::Name { sheet, name } => (sheet.clone(), Line::of(name)?),
            Token::Number(one) => (None, Line::of_row(*one)?),
            _ => return None,
        };
        if self.tokens.get(self.pos + 1) != Some(&Token::Colon) {
            return None;
        }
        let last = match self.tokens.get(self.pos + 2)? {
            // `Sheet1!A:Sheet1!C` is legal and means the same as `Sheet1!A:C`;
            // a different sheet on the right is not, and Excel rejects it.
            Token::Name { sheet: other, name } => {
                if other.is_some() && other != &sheet {
                    return None;
                }
                Line::of(name)?
            }
            Token::Number(one) => Line::of_row(*one)?,
            _ => return None,
        };
        if std::mem::discriminant(&first) != std::mem::discriminant(&last) {
            return None;
        }
        self.pos += 3;
        let range = match (first, last) {
            (Line::Column(a), Line::Column(b)) => RangeRef::normalised(
                CellRef::new(a, 0),
                CellRef::new(b, MAX_ROW),
            ),
            (Line::Row(a), Line::Row(b)) => RangeRef::normalised(
                CellRef::new(0, a),
                CellRef::new(MAX_COL, b),
            ),
            _ => return None,
        };
        Some(Expr::Ref(Reference { sheet, range }))
    }

    fn parse_range(&mut self) -> Result<Expr, ParseError> {
        if let Some(whole) = self.whole_line() {
            return Ok(whole);
        }
        let lhs = self.parse_atom()?;
        if !self.eat(&Token::Colon) {
            return Ok(lhs);
        }
        let rhs = self.parse_atom()?;

        // `Sheet1!A1:B2` qualifies the whole range with the left sheet.
        match (&lhs, &rhs) {
            (Expr::Ref(a), Expr::Ref(b)) if a.range.is_single() && b.range.is_single() => {
                Ok(Expr::Ref(Reference {
                    sheet: a.sheet.clone().or_else(|| b.sheet.clone()),
                    range: RangeRef::normalised(a.range.start, b.range.start),
                }))
            }
            _ => Err(ParseError::UnexpectedToken(
                "`:` requires a cell reference on both sides".to_string(),
            )),
        }
    }

    fn parse_atom(&mut self) -> Result<Expr, ParseError> {
        let token = self.bump().ok_or(ParseError::UnexpectedEnd)?;
        match token {
            Token::Number(n) => Ok(Expr::Literal(Value::Number(n))),
            Token::Text(s) => Ok(Expr::Literal(Value::Text(s))),
            Token::ErrorLit(e) => Ok(Expr::Literal(Value::Error(e))),
            Token::LParen => {
                let inner = self.parse_comparison()?;
                if !self.eat(&Token::RParen) {
                    return Err(ParseError::UnexpectedToken("expected `)`".to_string()));
                }
                Ok(inner)
            }
            Token::LBrace => self.parse_array(),
            Token::Table { name, asked } => Ok(Expr::Table { name, asked }),
            Token::Name { sheet, name } => self.finish_name(sheet, name),
            other => Err(ParseError::UnexpectedToken(format!("{other:?}"))),
        }
    }

    /// An array written out: `{1,2;3,4}`, rows separated by `;` and columns by
    /// `,`. Only constants go inside -- a reference or a call there is not a
    /// thing Excel accepts -- and every row must be the same width, which
    /// Excel refuses the formula outright for.
    fn parse_array(&mut self) -> Result<Expr, ParseError> {
        let mut rows: Vec<Vec<Value>> = Vec::new();
        let mut row: Vec<Value> = Vec::new();
        loop {
            row.push(self.parse_array_element()?);
            match self.bump() {
                Some(Token::Comma) => {}
                Some(Token::Semicolon) => {
                    rows.push(std::mem::take(&mut row));
                }
                Some(Token::RBrace) => {
                    rows.push(row);
                    break;
                }
                Some(other) => {
                    return Err(ParseError::UnexpectedToken(format!("{other:?} in an array")))
                }
                None => return Err(ParseError::UnexpectedEnd),
            }
        }
        let width = rows[0].len();
        if rows.iter().any(|row| row.len() != width) {
            return Err(ParseError::UnexpectedToken(
                "every row of an array must be the same width".to_string(),
            ));
        }
        Ok(Expr::Array(rows))
    }

    /// One constant inside an array. A sign may lead it; nothing else may.
    fn parse_array_element(&mut self) -> Result<Value, ParseError> {
        let negative = match self.peek() {
            Some(Token::Minus) => {
                self.pos += 1;
                true
            }
            Some(Token::Plus) => {
                self.pos += 1;
                false
            }
            _ => false,
        };
        let token = self.bump().ok_or(ParseError::UnexpectedEnd)?;
        let value = match token {
            Token::Number(n) => Value::Number(if negative { -n } else { n }),
            Token::Text(s) if !negative => Value::Text(s),
            Token::ErrorLit(e) if !negative => Value::Error(e),
            Token::Name { sheet: None, name } if !negative && name.eq_ignore_ascii_case("TRUE") => {
                Value::Logical(true)
            }
            Token::Name { sheet: None, name } if !negative && name.eq_ignore_ascii_case("FALSE") => {
                Value::Logical(false)
            }
            other => {
                return Err(ParseError::UnexpectedToken(format!(
                    "{other:?} is not a constant an array can hold"
                )))
            }
        };
        Ok(value)
    }

    /// Decide what a bare name meant, now that we can see what follows it.
    fn finish_name(&mut self, sheet: Option<String>, name: String) -> Result<Expr, ParseError> {
        // A following `(` makes it a call, even for TRUE()/FALSE().
        if self.peek() == Some(&Token::LParen) {
            self.pos += 1;
            let args = self.parse_args()?;
            return Ok(Expr::Function {
                name: name.to_uppercase(),
                args,
            });
        }

        if sheet.is_none() {
            if name.eq_ignore_ascii_case("TRUE") {
                return Ok(Expr::Literal(Value::Logical(true)));
            }
            if name.eq_ignore_ascii_case("FALSE") {
                return Ok(Expr::Literal(Value::Logical(false)));
            }
        }

        if let Some(cell) = parse_a1(&name) {
            return Ok(Expr::Ref(Reference {
                sheet,
                range: RangeRef::single(cell),
            }));
        }

        // Not a reference and not a call: a defined name. A sheet-qualified
        // defined name keeps only its local part; scoping is resolved later.
        Ok(Expr::Name(name.to_uppercase()))
    }

    fn parse_args(&mut self) -> Result<Vec<Expr>, ParseError> {
        // `FOO()` — no arguments at all.
        if self.eat(&Token::RParen) {
            return Ok(Vec::new());
        }

        let mut args = Vec::new();
        loop {
            // An omitted argument is legal and meaningful: `IF(A1,,"x")` passes
            // a blank as the second argument rather than shifting the third.
            if matches!(self.peek(), Some(Token::Comma) | Some(Token::RParen)) {
                args.push(Expr::Literal(Value::Blank));
            } else {
                args.push(self.parse_comparison()?);
            }

            if self.eat(&Token::Comma) {
                continue;
            }
            if self.eat(&Token::RParen) {
                return Ok(args);
            }
            return Err(ParseError::UnexpectedToken(
                "expected `,` or `)` in argument list".to_string(),
            ));
        }
    }
}

fn binary(op: BinaryOp, lhs: Expr, rhs: Expr) -> Expr {
    Expr::Binary {
        op,
        lhs: Box::new(lhs),
        rhs: Box::new(rhs),
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::value::ExcelError;

    fn p(s: &str) -> Expr {
        parse(s).unwrap_or_else(|e| panic!("failed to parse {s:?}: {e}"))
    }

    #[test]
    fn comparison_operators_parse() {
        // The prototype this replaces could not parse these at all, so
        // `IF(A1>10,...)` evaluated to #VALUE!.
        let expr = p("A1>10");
        assert!(matches!(
            expr,
            Expr::Binary {
                op: BinaryOp::Gt,
                ..
            }
        ));
    }

    #[test]
    fn unary_minus_binds_tighter_than_power() {
        // Excel: -2^2 = 4
        assert_eq!(
            p("-2^2"),
            binary(
                BinaryOp::Pow,
                Expr::Unary {
                    op: UnaryOp::Neg,
                    operand: Box::new(Expr::Literal(Value::Number(2.0)))
                },
                Expr::Literal(Value::Number(2.0))
            )
        );
    }

    /// Which references a formula has to WAIT for, as against which it merely
    /// measures.
    ///
    /// `ROWS(A1:A9)` asks how tall that block is and never looks inside it, so
    /// counting the block as a dependency is what turns a self-numbering
    /// series into a cycle. The line is drawn at "a plain reference", because
    /// that is the only shape the evaluator answers without evaluating — and
    /// the graph has to draw it in exactly the same place, or a formula waits
    /// for something that is never computed, or is computed too early.
    #[test]
    fn a_range_that_is_only_measured_is_not_a_dependency() {
        let waits_for = |s: &str| {
            let mut names: Vec<String> = p(s)
                .value_references()
                .iter()
                .map(|one| one.range.to_string())
                .collect();
            names.sort();
            names
        };
        assert_eq!(waits_for("=SUM(A1:A9)"), ["A1:A9"]);
        assert!(waits_for("=ROWS(A1:A9)").is_empty());
        assert!(waits_for("=COLUMNS(A1:D1)").is_empty());
        assert!(waits_for("=ROW(A1:A9)").is_empty());
        // The other arguments still count, and so does a second use of the
        // same range that IS reading it.
        assert_eq!(waits_for("=SUM(A1:A9)+ROWS(A1:A9)"), ["A1:A9"]);
        assert_eq!(waits_for("=INDEX(B1:B9,ROWS(A1:A9))"), ["B1:B9"]);
        // An argument that has to be worked out first is a dependency like any
        // other: what it reads, the count reads.
        assert_eq!(waits_for("=ROWS(INDEX(A1:C5,,B1))"), ["A1:C5", "B1"]);
    }

    /// A sheet in another workbook cannot be read, and saying so is better
    /// than answering.
    ///
    /// `[1]` names a workbook this file links to and does not contain. The
    /// bare form never lexed; the quoted form did, because brackets are
    /// ordinary characters inside a sheet name, and it went on to resolve
    /// against a sheet that is not there. A formula left unread keeps the
    /// value the file was saved with, which is the best answer available.
    /// One that resolves to `#REF!` replaces that value with an error.
    #[test]
    fn a_sheet_in_another_workbook_is_not_read() {
        for formula in [
            "'[1]May 2021'!$L$2:$L$29",
            "INDEX('[1]May 2021'!$L$2:$L$29,2)",
            "SUM('[2]Q3'!A1:A9)",
        ] {
            assert!(
                parse(formula).is_err(),
                "{formula} names a workbook that is not here",
            );
        }
        // A sheet whose name merely needs quoting is read as it always was.
        assert!(parse("'May 2021'!$L$2:$L$29").is_ok());
        assert!(parse("'HR '!$A3").is_ok());
    }

    #[test]
    fn power_is_left_associative() {
        // Excel: 2^3^2 = 64, not 512.
        let expr = p("2^3^2");
        match expr {
            Expr::Binary {
                op: BinaryOp::Pow,
                lhs,
                ..
            } => assert!(matches!(
                *lhs,
                Expr::Binary {
                    op: BinaryOp::Pow,
                    ..
                }
            )),
            other => panic!("unexpected tree: {other:?}"),
        }
    }

    #[test]
    fn percent_is_postfix() {
        assert_eq!(
            p("50%"),
            Expr::Unary {
                op: UnaryOp::Percent,
                operand: Box::new(Expr::Literal(Value::Number(50.0)))
            }
        );
    }

    #[test]
    fn absolute_and_sheet_qualified_references_parse() {
        assert_eq!(p("$A$1").references().len(), 1);
        let expr = p("Sheet1!A1:B2");
        let refs = expr.references();
        assert_eq!(refs.len(), 1);
        assert_eq!(refs[0].sheet.as_deref(), Some("Sheet1"));
        assert_eq!(refs[0].range.to_string(), "A1:B2");
    }

    #[test]
    fn functions_are_case_insensitive_and_nest() {
        let expr = p("sum(A1:A3, max(B1,B2))");
        match expr {
            Expr::Function { name, args } => {
                assert_eq!(name, "SUM");
                assert_eq!(args.len(), 2);
            }
            other => panic!("unexpected tree: {other:?}"),
        }
    }

    #[test]
    fn log10_is_a_function_not_a_reference() {
        assert!(matches!(p("LOG10(100)"), Expr::Function { .. }));
    }

    #[test]
    fn bare_names_become_defined_names() {
        assert_eq!(p("TAX_RATE"), Expr::Name("TAX_RATE".to_string()));
        assert_eq!(p("税率"), Expr::Name("税率".to_string()));
    }

    #[test]
    fn true_and_false_are_literals_but_can_be_called() {
        assert_eq!(p("TRUE"), Expr::Literal(Value::Logical(true)));
        assert!(matches!(p("TRUE()"), Expr::Function { .. }));
    }

    #[test]
    fn omitted_arguments_hold_their_position() {
        match p(r#"IF(A1,,"x")"#) {
            Expr::Function { args, .. } => {
                assert_eq!(args.len(), 3);
                assert_eq!(args[1], Expr::Literal(Value::Blank));
            }
            other => panic!("unexpected tree: {other:?}"),
        }
    }

    #[test]
    fn error_literals_parse() {
        assert_eq!(p("#N/A"), Expr::Literal(Value::Error(ExcelError::NA)));
    }

    #[test]
    fn concatenation_is_looser_than_arithmetic() {
        // 1&2+3 parses as 1 & (2+3)
        match p("1&2+3") {
            Expr::Binary {
                op: BinaryOp::Concat,
                rhs,
                ..
            } => assert!(matches!(
                *rhs,
                Expr::Binary {
                    op: BinaryOp::Add,
                    ..
                }
            )),
            other => panic!("unexpected tree: {other:?}"),
        }
    }

    #[test]
    fn junk_is_rejected_rather_than_silently_accepted() {
        assert!(parse("1+").is_err());
        assert!(parse("(1").is_err());
        assert!(parse("1 2").is_err());
    }
}
