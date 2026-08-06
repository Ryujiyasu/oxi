// This Source Code Form is subject to the terms of the Mozilla Public
// License, v. 2.0. If a copy of the MPL was not distributed with this
// file, You can obtain one at https://mozilla.org/MPL/2.0/.

//! Tokeniser for VBA source.
//!
//! VBA is line-oriented, so the end of a line is a token rather than
//! whitespace. Several of its lexical rules have no analogue in a modern
//! language and are the usual source of wrong results in ad-hoc parsers:
//!
//! - A trailing `_` continues a logical line onto the next physical one.
//! - `Rem` starts a comment, exactly like `'`.
//! - A leading integer on a line is a *line number*, not an expression.
//! - `:` separates statements on one line, but also ends a label.
//! - `$ % & ! # @` after an identifier are type suffixes, not operators.
//! - `#...#` is a date literal, and `#If` starts a conditional-compilation
//!   directive. Both begin with the same character.
//! - Keywords are case-insensitive, and the VBE rewrites their casing, so the
//!   original casing carries no meaning.

use std::fmt;

/// Where a token came from, for diagnostics and for source-preserving output.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct Span {
    /// 1-based physical line.
    pub line: u32,
    /// 0-based byte offset within the source.
    pub start: usize,
    pub end: usize,
}

#[derive(Debug, Clone, PartialEq)]
pub enum TokenKind {
    /// A name: identifier or keyword. Classification happens in the parser,
    /// because VBA lets you name a variable `Line` or `Name`.
    Ident(String),
    /// Numeric literal, with the type suffix stripped.
    Number(f64),
    /// String literal, with `""` unescaped to `"`.
    Str(String),
    /// `#1999-12-31#` or `#12/31/1999#`.
    DateLit(String),
    /// A line number at the start of a line, kept because `Erl` reports it and
    /// `GoTo 100` targets it.
    LineNumber(u32),
    /// `'` or `Rem` comment, text only.
    Comment(String),
    /// `#If` / `#Else` / `#End If` and friends, verbatim.
    Directive(String),
    /// A type suffix attached to the preceding identifier.
    TypeSuffix(char),

    Punct(Punct),
    /// End of a logical line. Line continuations do not produce one.
    Eol,
    Eof,
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum Punct {
    Plus,
    Minus,
    Star,
    Slash,
    BackSlash,
    Caret,
    Amp,
    Eq,
    Lt,
    Gt,
    Le,
    Ge,
    Ne,
    LParen,
    RParen,
    Comma,
    Dot,
    Colon,
    Semicolon,
    /// `:=` in a named argument.
    Assign,
    /// A bare `#`, as in the file handle of `Print #1, x`.
    Hash,
}

impl Punct {
    pub fn as_str(self) -> &'static str {
        match self {
            Punct::Plus => "+",
            Punct::Minus => "-",
            Punct::Star => "*",
            Punct::Slash => "/",
            Punct::BackSlash => "\\",
            Punct::Caret => "^",
            Punct::Amp => "&",
            Punct::Eq => "=",
            Punct::Lt => "<",
            Punct::Gt => ">",
            Punct::Le => "<=",
            Punct::Ge => ">=",
            Punct::Ne => "<>",
            Punct::LParen => "(",
            Punct::RParen => ")",
            Punct::Comma => ",",
            Punct::Dot => ".",
            Punct::Colon => ":",
            Punct::Semicolon => ";",
            Punct::Assign => ":=",
            Punct::Hash => "#",
        }
    }
}

#[derive(Debug, Clone, PartialEq)]
pub struct Token {
    pub kind: TokenKind,
    pub span: Span,
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub enum LexError {
    UnterminatedString { line: u32 },
    UnterminatedDate { line: u32 },
    UnexpectedChar { ch: char, line: u32 },
}

impl fmt::Display for LexError {
    fn fmt(&self, f: &mut fmt::Formatter<'_>) -> fmt::Result {
        match self {
            LexError::UnterminatedString { line } => {
                write!(f, "unterminated string literal on line {line}")
            }
            LexError::UnterminatedDate { line } => {
                write!(f, "unterminated date literal on line {line}")
            }
            LexError::UnexpectedChar { ch, line } => {
                write!(f, "unexpected character {ch:?} on line {line}")
            }
        }
    }
}

impl std::error::Error for LexError {}

pub fn tokenize(source: &str) -> Result<Vec<Token>, LexError> {
    Lexer::new(source).run()
}

struct Lexer<'a> {
    src: &'a str,
    bytes: &'a [u8],
    pos: usize,
    line: u32,
    /// True until the first non-trivia token of a logical line is emitted.
    /// A number seen while this holds is a line number, not a literal.
    at_line_start: bool,
    out: Vec<Token>,
}

impl<'a> Lexer<'a> {
    fn new(src: &'a str) -> Lexer<'a> {
        Lexer {
            src,
            bytes: src.as_bytes(),
            pos: 0,
            line: 1,
            at_line_start: true,
            out: Vec::new(),
        }
    }

    fn run(mut self) -> Result<Vec<Token>, LexError> {
        while self.pos < self.bytes.len() {
            let c = self.bytes[self.pos];

            if c == b'\r' {
                self.pos += 1;
                continue;
            }
            if c == b'\n' {
                self.push_here(TokenKind::Eol, self.pos, self.pos + 1);
                self.pos += 1;
                self.line += 1;
                self.at_line_start = true;
                continue;
            }
            if c == b' ' || c == b'\t' {
                self.pos += 1;
                continue;
            }

            // A `_` at end of line joins this line to the next. Only whitespace
            // and an optional comment may follow it.
            if c == b'_' && self.is_line_continuation() {
                self.skip_to_next_physical_line();
                continue;
            }

            if c == b'\'' {
                self.lex_comment_body(self.pos, 1);
                continue;
            }

            if c == b'#' {
                self.lex_hash()?;
                continue;
            }

            if c.is_ascii_digit() || (c == b'.' && self.peek_digit(1)) {
                if self.at_line_start && c.is_ascii_digit() && self.line_number_follows() {
                    self.lex_line_number();
                } else {
                    self.lex_number();
                }
                self.at_line_start = false;
                continue;
            }

            if c == b'"' {
                self.lex_string()?;
                self.at_line_start = false;
                continue;
            }

            if c == b'&' && self.lex_radix_number() {
                self.at_line_start = false;
                continue;
            }

            if is_ident_start(self.char_at(self.pos)) {
                self.lex_ident_or_rem();
                continue;
            }

            match self.lex_punct() {
                Some(()) => {
                    self.at_line_start = false;
                }
                None => {
                    return Err(LexError::UnexpectedChar {
                        ch: self.char_at(self.pos),
                        line: self.line,
                    })
                }
            }
        }

        self.push_here(TokenKind::Eof, self.pos, self.pos);
        Ok(self.out)
    }

    // -- helpers ---------------------------------------------------------

    fn char_at(&self, at: usize) -> char {
        self.src[at..].chars().next().unwrap_or('\0')
    }

    fn peek_digit(&self, ahead: usize) -> bool {
        matches!(self.bytes.get(self.pos + ahead), Some(b) if b.is_ascii_digit())
    }

    fn push_here(&mut self, kind: TokenKind, start: usize, end: usize) {
        self.out.push(Token {
            kind,
            span: Span {
                line: self.line,
                start,
                end,
            },
        });
    }

    /// A `_` continues the line only when nothing but whitespace (and possibly
    /// a comment) separates it from the newline.
    fn is_line_continuation(&self) -> bool {
        let mut i = self.pos + 1;
        while matches!(self.bytes.get(i), Some(b' ') | Some(b'\t') | Some(b'\r')) {
            i += 1;
        }
        matches!(self.bytes.get(i), Some(b'\n') | None)
    }

    fn skip_to_next_physical_line(&mut self) {
        while self.pos < self.bytes.len() && self.bytes[self.pos] != b'\n' {
            self.pos += 1;
        }
        if self.pos < self.bytes.len() {
            self.pos += 1;
            self.line += 1;
        }
    }

    /// A leading integer is a line number only when it is followed by
    /// whitespace or a colon; `1 + 2` on its own line is an expression, but
    /// VBA does not allow that as a statement anyway.
    fn line_number_follows(&self) -> bool {
        let mut i = self.pos;
        while matches!(self.bytes.get(i), Some(b) if b.is_ascii_digit()) {
            i += 1;
        }
        matches!(
            self.bytes.get(i),
            Some(b' ') | Some(b'\t') | Some(b':') | Some(b'\r') | Some(b'\n') | None
        )
    }

    fn lex_line_number(&mut self) {
        let start = self.pos;
        while matches!(self.bytes.get(self.pos), Some(b) if b.is_ascii_digit()) {
            self.pos += 1;
        }
        let value = self.src[start..self.pos].parse().unwrap_or(0);
        self.push_here(TokenKind::LineNumber(value), start, self.pos);
    }

    fn lex_comment_body(&mut self, start: usize, skip: usize) {
        let text_start = start + skip;
        let mut end = text_start;
        while end < self.bytes.len() && self.bytes[end] != b'\n' {
            end += 1;
        }
        let text = self.src[text_start..end].trim_end_matches('\r').to_string();
        self.push_here(TokenKind::Comment(text), start, end);
        self.pos = end;
    }

    fn lex_string(&mut self) -> Result<(), LexError> {
        let start = self.pos;
        let mut i = self.pos + 1;
        let mut value = String::new();
        loop {
            match self.bytes.get(i) {
                None | Some(b'\n') => return Err(LexError::UnterminatedString { line: self.line }),
                Some(b'"') => {
                    // A doubled quote is an escaped quote.
                    if self.bytes.get(i + 1) == Some(&b'"') {
                        value.push('"');
                        i += 2;
                        continue;
                    }
                    i += 1;
                    break;
                }
                Some(_) => {
                    let ch = self.char_at(i);
                    value.push(ch);
                    i += ch.len_utf8();
                }
            }
        }
        self.push_here(TokenKind::Str(value), start, i);
        self.pos = i;
        Ok(())
    }

    /// `#` starts either a date literal or a compiler directive.
    fn lex_hash(&mut self) -> Result<(), LexError> {
        let start = self.pos;
        let rest = &self.src[self.pos + 1..];
        let is_directive = ["if", "else", "elseif", "end", "const"]
            .iter()
            .any(|kw| rest.len() >= kw.len() && rest[..kw.len()].eq_ignore_ascii_case(kw));

        if is_directive {
            let mut end = self.pos;
            while end < self.bytes.len() && self.bytes[end] != b'\n' {
                end += 1;
            }
            let text = self.src[start..end].trim_end_matches('\r').to_string();
            self.push_here(TokenKind::Directive(text), start, end);
            self.pos = end;
            self.at_line_start = false;
            return Ok(());
        }

        let mut i = self.pos + 1;
        while let Some(&b) = self.bytes.get(i) {
            if b == b'#' {
                let text = self.src[self.pos + 1..i].to_string();
                // In `Close #1, #2` (and similar file-I/O statements), the
                // second handle's opening hash is not the first handle's date
                // terminator. A separator immediately before it distinguishes
                // that shape without narrowing VBA's accepted date spellings.
                if !text.trim_end().ends_with([',', ';']) {
                    self.push_here(TokenKind::DateLit(text), start, i + 1);
                    self.pos = i + 1;
                    self.at_line_start = false;
                    return Ok(());
                }
                break;
            }
            if b == b'\n' {
                break;
            }
            i += 1;
        }

        // No closing `#` on this line, so it was never a date. The file I/O
        // statements spell a file handle `#1`, and refusing to lex those would
        // fail the whole module over a construct that is out of scope anyway.
        self.push_here(TokenKind::Punct(Punct::Hash), start, start + 1);
        self.pos = start + 1;
        self.at_line_start = false;
        Ok(())
    }

    /// `&H1F` and `&O17` are hex and octal literals. Without this they lex as
    /// the concatenation operator followed by an identifier, which parses
    /// cleanly and computes the wrong thing — the worst kind of failure.
    fn lex_radix_number(&mut self) -> bool {
        let Some(&marker) = self.bytes.get(self.pos + 1) else {
            return false;
        };
        let radix = match marker {
            b'h' | b'H' => 16,
            b'o' | b'O' => 8,
            _ => return false,
        };

        let digits_start = self.pos + 2;
        let mut i = digits_start;
        while matches!(self.bytes.get(i), Some(b) if (*b as char).is_digit(radix)) {
            i += 1;
        }
        if i == digits_start {
            return false;
        }

        let value = u64::from_str_radix(&self.src[digits_start..i], radix).unwrap_or(0);
        let start = self.pos;
        self.pos = i;
        let suffix = self
            .bytes
            .get(self.pos)
            .copied()
            .filter(|b| is_type_suffix(*b) || *b == b'^')
            .map(char::from);
        // VBA sign-extends an unsuffixed radix literal from the narrowest
        // Integer or Long bit pattern. Explicit `&` and `^` instead select
        // 32- and 64-bit interpretation.
        let number = match suffix {
            Some('%') => (value as u16 as i16) as f64,
            Some('&') => (value as u32 as i32) as f64,
            Some('^') => (value as i64) as f64,
            _ if value <= u16::MAX as u64 => (value as u16 as i16) as f64,
            _ if value <= u32::MAX as u64 => (value as u32 as i32) as f64,
            _ => value as f64,
        };
        self.push_here(TokenKind::Number(number), start, i);

        if let Some(suffix) = suffix {
            let kind = if suffix == '^' {
                TokenKind::Punct(Punct::Caret)
            } else {
                TokenKind::TypeSuffix(suffix)
            };
            self.push_here(kind, self.pos, self.pos + 1);
            self.pos += 1;
        }
        true
    }

    fn lex_number(&mut self) {
        let start = self.pos;

        while matches!(self.bytes.get(self.pos), Some(b) if b.is_ascii_digit()) {
            self.pos += 1;
        }
        if self.bytes.get(self.pos) == Some(&b'.') && self.peek_digit(1) {
            self.pos += 1;
            while matches!(self.bytes.get(self.pos), Some(b) if b.is_ascii_digit()) {
                self.pos += 1;
            }
        }
        // Exponent, with `D` accepted alongside `E` for Double literals.
        if matches!(
            self.bytes.get(self.pos),
            Some(b'e') | Some(b'E') | Some(b'd') | Some(b'D')
        ) {
            let mut j = self.pos + 1;
            if matches!(self.bytes.get(j), Some(b'+') | Some(b'-')) {
                j += 1;
            }
            if matches!(self.bytes.get(j), Some(b) if b.is_ascii_digit()) {
                j += 1;
                while matches!(self.bytes.get(j), Some(b) if b.is_ascii_digit()) {
                    j += 1;
                }
                self.pos = j;
            }
        }

        let text = self.src[start..self.pos].replace(['d', 'D'], "e");
        let value = text.parse::<f64>().unwrap_or(0.0);
        self.push_here(TokenKind::Number(value), start, self.pos);

        // A type suffix binds to the literal.
        if let Some(&b) = self.bytes.get(self.pos) {
            if is_type_suffix(b) {
                self.push_here(TokenKind::TypeSuffix(b as char), self.pos, self.pos + 1);
                self.pos += 1;
            }
        }
    }

    fn lex_ident_or_rem(&mut self) {
        let start = self.pos;
        while self.pos < self.bytes.len() && is_ident_continue(self.char_at(self.pos)) {
            self.pos += self.char_at(self.pos).len_utf8();
        }
        let word = &self.src[start..self.pos];

        // `Rem` is a comment keyword, not an identifier, but only when it is a
        // whole word: `Remark` is a perfectly good variable name.
        if word.eq_ignore_ascii_case("rem") {
            self.lex_comment_body(start, 3);
            self.at_line_start = false;
            return;
        }

        self.push_here(TokenKind::Ident(word.to_string()), start, self.pos);
        self.at_line_start = false;

        if let Some(&b) = self.bytes.get(self.pos) {
            // `!` is also the dictionary-access operator, but as a suffix it is
            // only ambiguous in contexts the parser resolves.
            if is_type_suffix(b) {
                self.push_here(TokenKind::TypeSuffix(b as char), self.pos, self.pos + 1);
                self.pos += 1;
            }
        }
    }

    fn lex_punct(&mut self) -> Option<()> {
        let start = self.pos;
        let two = self.src.get(start..start + 2);
        let (punct, len) = match two {
            Some("<=") => (Punct::Le, 2),
            Some(">=") => (Punct::Ge, 2),
            Some("<>") => (Punct::Ne, 2),
            Some("=<") => (Punct::Le, 2),
            Some("=>") => (Punct::Ge, 2),
            Some(":=") => (Punct::Assign, 2),
            _ => {
                let one = match self.bytes[start] {
                    b'+' => Punct::Plus,
                    b'-' => Punct::Minus,
                    b'*' => Punct::Star,
                    b'/' => Punct::Slash,
                    b'\\' => Punct::BackSlash,
                    b'^' => Punct::Caret,
                    b'&' => Punct::Amp,
                    b'=' => Punct::Eq,
                    b'<' => Punct::Lt,
                    b'>' => Punct::Gt,
                    b'(' => Punct::LParen,
                    b')' => Punct::RParen,
                    b',' => Punct::Comma,
                    b'.' => Punct::Dot,
                    b':' => Punct::Colon,
                    b';' => Punct::Semicolon,
                    _ => return None,
                };
                (one, 1)
            }
        };
        self.push_here(TokenKind::Punct(punct), start, start + len);
        self.pos += len;
        Some(())
    }
}

fn is_ident_start(c: char) -> bool {
    c.is_alphabetic() || c == '_' || !c.is_ascii()
}

fn is_ident_continue(c: char) -> bool {
    c.is_alphanumeric() || c == '_' || !c.is_ascii()
}

fn is_type_suffix(b: u8) -> bool {
    matches!(b, b'$' | b'%' | b'&' | b'!' | b'#' | b'@')
}

#[cfg(test)]
mod tests {
    use super::*;

    fn kinds(src: &str) -> Vec<TokenKind> {
        tokenize(src)
            .expect("should lex")
            .into_iter()
            .map(|t| t.kind)
            .filter(|k| !matches!(k, TokenKind::Eof))
            .collect()
    }

    fn ident(name: &str) -> TokenKind {
        TokenKind::Ident(name.to_string())
    }

    #[test]
    fn line_endings_are_tokens() {
        assert_eq!(kinds("a\nb"), vec![ident("a"), TokenKind::Eol, ident("b")]);
    }

    #[test]
    fn trailing_underscore_joins_lines() {
        // One logical line, so no Eol in the middle.
        assert_eq!(
            kinds("Debug.Print a, _\n    b"),
            vec![
                ident("Debug"),
                TokenKind::Punct(Punct::Dot),
                ident("Print"),
                ident("a"),
                TokenKind::Punct(Punct::Comma),
                ident("b"),
            ]
        );
    }

    #[test]
    fn rem_is_a_comment_but_remark_is_a_name() {
        assert_eq!(
            kinds("Rem hello"),
            vec![TokenKind::Comment(" hello".to_string())]
        );
        assert_eq!(kinds("Remark = 1")[0], ident("Remark"));
    }

    #[test]
    fn apostrophe_comments_run_to_end_of_line() {
        assert_eq!(
            kinds("x = 1 ' set x\ny"),
            vec![
                ident("x"),
                TokenKind::Punct(Punct::Eq),
                TokenKind::Number(1.0),
                TokenKind::Comment(" set x".to_string()),
                TokenKind::Eol,
                ident("y"),
            ]
        );
    }

    #[test]
    fn leading_integers_are_line_numbers() {
        assert_eq!(
            kinds("10 x = 1"),
            vec![
                TokenKind::LineNumber(10),
                ident("x"),
                TokenKind::Punct(Punct::Eq),
                TokenKind::Number(1.0),
            ]
        );
        // Not at the start of a line, so it is just a number.
        assert_eq!(kinds("x = 10")[2], TokenKind::Number(10.0));
    }

    #[test]
    fn type_suffixes_attach_to_the_preceding_token() {
        assert_eq!(
            kinds("s$ = 1&"),
            vec![
                ident("s"),
                TokenKind::TypeSuffix('$'),
                TokenKind::Punct(Punct::Eq),
                TokenKind::Number(1.0),
                TokenKind::TypeSuffix('&'),
            ]
        );
    }

    #[test]
    fn doubled_quotes_escape_inside_strings() {
        assert_eq!(
            kinds(r#""a""b""#),
            vec![TokenKind::Str(r#"a"b"#.to_string())]
        );
        assert!(matches!(
            tokenize("\"oops\n"),
            Err(LexError::UnterminatedString { line: 1 })
        ));
    }

    #[test]
    fn hash_starts_a_date_or_a_directive() {
        assert_eq!(
            kinds("d = #1999-12-31#"),
            vec![
                ident("d"),
                TokenKind::Punct(Punct::Eq),
                TokenKind::DateLit("1999-12-31".to_string()),
            ]
        );
        assert_eq!(
            kinds("#If Win64 Then"),
            vec![TokenKind::Directive("#If Win64 Then".to_string())]
        );
    }

    #[test]
    fn multiple_file_handles_do_not_merge_into_a_date_literal() {
        assert_eq!(
            kinds("Close #7, #8"),
            vec![
                ident("Close"),
                TokenKind::Punct(Punct::Hash),
                TokenKind::Number(7.0),
                TokenKind::Punct(Punct::Comma),
                TokenKind::Punct(Punct::Hash),
                TokenKind::Number(8.0),
            ]
        );
    }

    #[test]
    fn colon_separates_statements_and_assign_is_distinct() {
        assert_eq!(
            kinds("a = 1: b = 2"),
            vec![
                ident("a"),
                TokenKind::Punct(Punct::Eq),
                TokenKind::Number(1.0),
                TokenKind::Punct(Punct::Colon),
                ident("b"),
                TokenKind::Punct(Punct::Eq),
                TokenKind::Number(2.0),
            ]
        );
        assert_eq!(kinds("f a:=1")[2], TokenKind::Punct(Punct::Assign));
    }

    #[test]
    fn double_precision_exponent_uses_d() {
        assert_eq!(kinds("1.5D3"), vec![TokenKind::Number(1500.0)]);
        assert_eq!(kinds("1.5E3"), vec![TokenKind::Number(1500.0)]);
    }

    #[test]
    fn ampersand_h_is_a_hex_literal_not_concatenation() {
        assert_eq!(kinds("&HFF"), vec![TokenKind::Number(255.0)]);
        assert_eq!(kinds("&O17"), vec![TokenKind::Number(15.0)]);
        assert_eq!(
            kinds("&H1F&"),
            vec![TokenKind::Number(31.0), TokenKind::TypeSuffix('&')]
        );
        // A bare & is still concatenation.
        assert_eq!(
            kinds(r#"a & "x""#),
            vec![
                ident("a"),
                TokenKind::Punct(Punct::Amp),
                TokenKind::Str("x".to_string())
            ]
        );
    }

    #[test]
    fn radix_literals_follow_vba_signed_widths() {
        assert_eq!(kinds("&HFFFF"), vec![TokenKind::Number(-1.0)]);
        assert_eq!(kinds("&H8000"), vec![TokenKind::Number(-32768.0)]);
        assert_eq!(kinds("&HFFFFFFFF"), vec![TokenKind::Number(-1.0)]);
        assert_eq!(kinds("&O37777777777"), vec![TokenKind::Number(-1.0)]);
        assert_eq!(
            kinds("&HFFFF&"),
            vec![TokenKind::Number(65535.0), TokenKind::TypeSuffix('&')]
        );
        assert_eq!(
            kinds("&HFFFFFFFF^"),
            vec![
                TokenKind::Number(4294967295.0),
                TokenKind::Punct(Punct::Caret)
            ]
        );
    }

    #[test]
    fn japanese_identifiers_are_allowed() {
        // The VBE accepts these, and Japanese workbooks use them heavily.
        assert_eq!(
            kinds("税率 = 0.1"),
            vec![
                ident("税率"),
                TokenKind::Punct(Punct::Eq),
                TokenKind::Number(0.1),
            ]
        );
    }

    #[test]
    fn spans_track_physical_lines_across_continuations() {
        let tokens = tokenize("a = 1 + _\n    2\nb = 3").expect("should lex");
        let last = tokens
            .iter()
            .find(|t| t.kind == ident("b"))
            .expect("b present");
        assert_eq!(last.span.line, 3);
    }
}
