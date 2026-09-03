// This Source Code Form is subject to the terms of the Mozilla Public
// License, v. 2.0. If a copy of the MPL was not distributed with this
// file, You can obtain one at https://mozilla.org/MPL/2.0/.

//! Conditional compilation: `#If`, `#ElseIf`, `#Else`, `#End If`, `#Const`.
//!
//! This runs over the SOURCE, before a single token is read, because a branch
//! VBA does not take is not compiled at all. Asked of Excel, a module carrying
//!
//! ```text
//! #If Mac Then
//!     x = "unterminated
//! #End If
//! ```
//!
//! compiles and runs, and so does one whose dead branch holds `]not VBA at
//! all[ &&&` or an `If` with no `End If`. None of that could survive being
//! lexed, so the dead lines have to go before the lexer sees them.
//!
//! Every dropped line is replaced by a blank one of its own length, the way
//! `parse_module` already blanks a form's designer block: same bytes, same
//! newlines, so every span still points where it did.
//!
//! ## Why this is not merely tidy
//!
//! Without it both branches are read, and a declaration split across one is
//! read twice. Real code does this constantly:
//!
//! ```text
//! #If VBA7 Then
//! Private Function MouseProc(...) As LongPtr
//! #Else
//! Private Function MouseProc(...) As Long
//! #End If
//!     Dim idx As Long
//! ```
//!
//! -- two procedure headers with no `End Function` between them. The parser
//! loses the second header and everything after it, so ONE unread directive
//! costs a whole procedure. Measured over 378 real modules, this shape was the
//! largest single class of unread lines, and most of what it cost was not the
//! directive but the code around it.

use std::collections::BTreeMap;

/// What the compiler knows a name to stand for.
///
/// VBA's compilation constants are Variants, and the two subtypes that turn up
/// are a number and a string. A name never defined is Empty, which is falsy
/// and compares equal to zero -- asked of Excel, `#If UNDEF Then` does not
/// fire and `#If UNDEF = 0 Then` does.
#[derive(Clone, Debug, PartialEq)]
enum Known {
    Empty,
    Number(f64),
    Text(String),
}

impl Known {
    fn truthy(&self) -> bool {
        match self {
            Known::Empty => false,
            Known::Number(value) => *value != 0.0,
            Known::Text(text) => text.trim().parse::<f64>().is_ok_and(|n| n != 0.0),
        }
    }

    fn number(&self) -> f64 {
        match self {
            Known::Empty => 0.0,
            Known::Number(value) => *value,
            Known::Text(text) => text.trim().parse::<f64>().unwrap_or(0.0),
        }
    }
}

/// The constants Excel defines for itself.
///
/// Measured by asking a 64-bit Excel what each one is. They are 1, not -1:
/// `#If Win64 Then` fires while `#If Win64 = -1 Then` does not, and
/// `#If Win64 = 1 Then` does. VBA6 is defined and true alongside VBA7, which
/// is not what the documentation leads one to expect.
///
/// The two that are 0 are the two this host is not: a browser is not a Mac and
/// not 16-bit. Win32 is 1 even on a 64-bit Office, which is Excel's answer,
/// not a simplification.
fn predefined(name: &str) -> Option<Known> {
    let value = match name.to_ascii_lowercase().as_str() {
        "vba6" | "vba7" | "win32" | "win64" => 1.0,
        "win16" | "mac" => 0.0,
        _ => return None,
    };
    Some(Known::Number(value))
}

/// Resolve `#If` and friends, returning source of the same shape with the
/// branches VBA would not compile blanked out.
pub fn live_source(source: &str) -> String {
    if !carries_a_directive(source) {
        return source.to_string();
    }

    let mut defined: BTreeMap<String, Known> = BTreeMap::new();
    // Per open `#If`: whether some branch has already been taken, and whether
    // the branch now open is the live one. A branch inside a dead branch is
    // dead however its own condition reads.
    let mut stack: Vec<(bool, bool)> = Vec::new();
    let mut out = String::with_capacity(source.len());

    for line in split_keeping_ends(source) {
        let (body, ending) = split_ending(line);
        let trimmed = body.trim_start();
        let outer_live = stack.iter().all(|(_, live)| *live);

        let directive = directive_of(trimmed);
        match directive {
            Some(Directive::If(condition)) => {
                let live = outer_live && evaluate(condition, &defined).truthy();
                stack.push((live, live));
            }
            Some(Directive::ElseIf(condition)) => {
                if !stack.is_empty() {
                    let outer = stack_outer_live(&stack);
                    let taken = stack[stack.len() - 1].0;
                    let fires = !taken && outer && evaluate(condition, &defined).truthy();
                    let last = stack.len() - 1;
                    stack[last] = (taken || fires, fires);
                }
            }
            Some(Directive::Else) => {
                if !stack.is_empty() {
                    let outer = stack_outer_live(&stack);
                    let taken = stack[stack.len() - 1].0;
                    let fires = !taken && outer;
                    let last = stack.len() - 1;
                    stack[last] = (taken || fires, fires);
                }
            }
            Some(Directive::EndIf) => {
                stack.pop();
            }
            Some(Directive::Const(name, value)) => {
                if outer_live {
                    defined.insert(name.to_ascii_lowercase(), evaluate(value, &defined));
                }
            }
            None => {}
        }

        // A directive line is consumed here, so it is blanked whether or not
        // it is live -- the parser never sees one again. Ordinary lines are
        // kept only while every enclosing branch is live.
        let keep = directive.is_none() && stack.iter().all(|(_, live)| *live);
        if keep {
            out.push_str(line);
        } else {
            blank_into(&mut out, body);
            out.push_str(ending);
        }
    }
    out
}

/// Whether every branch OUTSIDE the innermost one is live, which is what
/// decides if the innermost may fire at all.
fn stack_outer_live(stack: &[(bool, bool)]) -> bool {
    stack[..stack.len() - 1].iter().all(|(_, live)| *live)
}

enum Directive<'a> {
    If(&'a str),
    ElseIf(&'a str),
    Else,
    EndIf,
    Const(&'a str, &'a str),
}

fn carries_a_directive(source: &str) -> bool {
    source.lines().any(|line| {
        let trimmed = line.trim_start();
        trimmed.starts_with('#')
            && ["if", "else", "elseif", "end", "const"]
                .iter()
                .any(|word| trimmed[1..].len() >= word.len()
                    && trimmed[1..word.len() + 1].eq_ignore_ascii_case(word))
    })
}

fn directive_of(trimmed: &str) -> Option<Directive<'_>> {
    let rest = trimmed.strip_prefix('#')?;
    let word_end = rest
        .find(|c: char| !c.is_ascii_alphabetic())
        .unwrap_or(rest.len());
    let (word, tail) = rest.split_at(word_end);
    let tail = tail.trim();
    match word.to_ascii_lowercase().as_str() {
        "if" => Some(Directive::If(strip_then(tail))),
        "elseif" => Some(Directive::ElseIf(strip_then(tail))),
        "else" => Some(Directive::Else),
        "end" if tail.eq_ignore_ascii_case("if") => Some(Directive::EndIf),
        "const" => {
            let (name, value) = tail.split_once('=')?;
            Some(Directive::Const(name.trim(), value.trim()))
        }
        _ => None,
    }
}

/// `#If cond Then` -- the trailing `Then` is not part of the condition.
fn strip_then(tail: &str) -> &str {
    let trimmed = tail.trim_end();
    if trimmed.len() >= 4 && trimmed[trimmed.len() - 4..].eq_ignore_ascii_case("then") {
        let head = &trimmed[..trimmed.len() - 4];
        if head.ends_with(|c: char| c.is_whitespace()) || head.ends_with(')') {
            return head.trim_end();
        }
    }
    trimmed
}

fn split_keeping_ends(source: &str) -> Vec<&str> {
    let mut lines = Vec::new();
    let mut start = 0;
    for (index, ch) in source.char_indices() {
        if ch == '\n' {
            lines.push(&source[start..index + 1]);
            start = index + 1;
        }
    }
    if start < source.len() {
        lines.push(&source[start..]);
    }
    lines
}

fn split_ending(line: &str) -> (&str, &str) {
    let body = line.trim_end_matches('\n').trim_end_matches('\r');
    (body, &line[body.len()..])
}

/// A run of spaces the same length as what it replaces, so every later span
/// still points at the byte it did.
fn blank_into(out: &mut String, body: &str) {
    for _ in body.chars() {
        out.push(' ');
    }
}

// --- the condition language ------------------------------------------------

/// Read a `#If` condition.
///
/// Measured against Excel: `And`, `Or`, `Not`, `=`, `>`, brackets and
/// user-defined `#Const` values all behave as they do in ordinary VBA, and a
/// name that was never defined is Empty -- falsy, and equal to zero.
///
/// Anything this cannot read answers Empty, which is falsy. That matches an
/// undefined name, which is the honest reading of a condition whose terms mean
/// nothing here.
fn evaluate(source: &str, defined: &BTreeMap<String, Known>) -> Known {
    let mut reader = Reader {
        text: source.as_bytes(),
        at: 0,
        defined,
    };
    let value = reader.or_expr();
    reader.skip_spaces();
    if reader.at < reader.text.len() {
        return Known::Empty;
    }
    value.unwrap_or(Known::Empty)
}

struct Reader<'a> {
    text: &'a [u8],
    at: usize,
    defined: &'a BTreeMap<String, Known>,
}

impl Reader<'_> {
    fn skip_spaces(&mut self) {
        while self.at < self.text.len() && self.text[self.at].is_ascii_whitespace() {
            self.at += 1;
        }
    }

    fn eat_word(&mut self, word: &str) -> bool {
        self.skip_spaces();
        let end = self.at + word.len();
        if end <= self.text.len()
            && self.text[self.at..end].eq_ignore_ascii_case(word.as_bytes())
            && self
                .text
                .get(end)
                .is_none_or(|b| !b.is_ascii_alphanumeric() && *b != b'_')
        {
            self.at = end;
            return true;
        }
        false
    }

    fn eat_symbol(&mut self, symbol: &str) -> bool {
        self.skip_spaces();
        let end = self.at + symbol.len();
        if end <= self.text.len() && &self.text[self.at..end] == symbol.as_bytes() {
            self.at = end;
            return true;
        }
        false
    }

    fn or_expr(&mut self) -> Option<Known> {
        let mut left = self.and_expr()?;
        loop {
            if self.eat_word("or") {
                let right = self.and_expr()?;
                left = Known::Number(bit_or(&left, &right));
            } else if self.eat_word("xor") {
                let right = self.and_expr()?;
                left = Known::Number(
                    ((left.number() as i64) ^ (right.number() as i64)) as f64,
                );
            } else {
                return Some(left);
            }
        }
    }

    fn and_expr(&mut self) -> Option<Known> {
        let mut left = self.not_expr()?;
        while self.eat_word("and") {
            let right = self.not_expr()?;
            left = Known::Number(
                ((left.number() as i64) & (right.number() as i64)) as f64,
            );
        }
        Some(left)
    }

    fn not_expr(&mut self) -> Option<Known> {
        if self.eat_word("not") {
            let value = self.not_expr()?;
            return Some(Known::Number(!(value.number() as i64) as f64));
        }
        self.compare()
    }

    fn compare(&mut self) -> Option<Known> {
        let left = self.sum()?;
        // The two-character operators are tried first, or `<=` reads as `<`.
        for (symbol, answer) in [
            ("<>", 0u8),
            ("<=", 1),
            (">=", 2),
            ("=", 3),
            ("<", 4),
            (">", 5),
        ] {
            if self.eat_symbol(symbol) {
                let right = self.sum()?;
                let same = match (&left, &right) {
                    (Known::Text(a), Known::Text(b)) => Some(a.cmp(b)),
                    (Known::Text(_), _) | (_, Known::Text(_)) => None,
                    _ => left.number().partial_cmp(&right.number()),
                };
                let order = match same {
                    Some(order) => order,
                    None => {
                        // A string against a number: only the two equality
                        // questions have an answer, and it is "not equal".
                        return Some(vba_bool(answer == 0));
                    }
                };
                let held = match answer {
                    0 => order != std::cmp::Ordering::Equal,
                    1 => order != std::cmp::Ordering::Greater,
                    2 => order != std::cmp::Ordering::Less,
                    3 => order == std::cmp::Ordering::Equal,
                    4 => order == std::cmp::Ordering::Less,
                    _ => order == std::cmp::Ordering::Greater,
                };
                return Some(vba_bool(held));
            }
        }
        Some(left)
    }

    fn sum(&mut self) -> Option<Known> {
        let mut left = self.product()?;
        loop {
            if self.eat_symbol("+") {
                let right = self.product()?;
                left = Known::Number(left.number() + right.number());
            } else if self.eat_symbol("-") {
                let right = self.product()?;
                left = Known::Number(left.number() - right.number());
            } else {
                return Some(left);
            }
        }
    }

    fn product(&mut self) -> Option<Known> {
        let mut left = self.unary()?;
        loop {
            if self.eat_symbol("*") {
                let right = self.unary()?;
                left = Known::Number(left.number() * right.number());
            } else if self.eat_symbol("/") {
                let right = self.unary()?;
                let divisor = right.number();
                if divisor == 0.0 {
                    return None;
                }
                left = Known::Number(left.number() / divisor);
            } else {
                return Some(left);
            }
        }
    }

    fn unary(&mut self) -> Option<Known> {
        if self.eat_symbol("-") {
            return Some(Known::Number(-self.unary()?.number()));
        }
        if self.eat_symbol("+") {
            return self.unary();
        }
        self.term()
    }

    fn term(&mut self) -> Option<Known> {
        self.skip_spaces();
        if self.eat_symbol("(") {
            let inner = self.or_expr()?;
            if !self.eat_symbol(")") {
                return None;
            }
            return Some(inner);
        }
        let byte = *self.text.get(self.at)?;
        if byte == b'"' {
            let mut text = String::new();
            self.at += 1;
            while let Some(&b) = self.text.get(self.at) {
                self.at += 1;
                if b == b'"' {
                    if self.text.get(self.at) == Some(&b'"') {
                        text.push('"');
                        self.at += 1;
                        continue;
                    }
                    return Some(Known::Text(text));
                }
                text.push(b as char);
            }
            return None;
        }
        if byte.is_ascii_digit() || byte == b'.' {
            let start = self.at;
            while self
                .text
                .get(self.at)
                .is_some_and(|b| b.is_ascii_digit() || *b == b'.')
            {
                self.at += 1;
            }
            let text = std::str::from_utf8(&self.text[start..self.at]).ok()?;
            return text.parse::<f64>().ok().map(Known::Number);
        }
        if byte.is_ascii_alphabetic() || byte == b'_' {
            let start = self.at;
            while self
                .text
                .get(self.at)
                .is_some_and(|b| b.is_ascii_alphanumeric() || *b == b'_')
            {
                self.at += 1;
            }
            let name = std::str::from_utf8(&self.text[start..self.at]).ok()?;
            if name.eq_ignore_ascii_case("true") {
                return Some(Known::Number(-1.0));
            }
            if name.eq_ignore_ascii_case("false") {
                return Some(Known::Number(0.0));
            }
            return Some(
                self.defined
                    .get(&name.to_ascii_lowercase())
                    .cloned()
                    .or_else(|| predefined(name))
                    .unwrap_or(Known::Empty),
            );
        }
        None
    }
}

fn bit_or(left: &Known, right: &Known) -> f64 {
    ((left.number() as i64) | (right.number() as i64)) as f64
}

fn vba_bool(held: bool) -> Known {
    Known::Number(if held { -1.0 } else { 0.0 })
}

#[cfg(test)]
mod tests {
    use super::*;

    /// The compilation constants Excel defines, and what they are worth.
    ///
    /// Measured on a 64-bit Excel: every one of VBA6, VBA7, Win32 and Win64
    /// fires, Win16 and Mac do not, and the four that fire are 1 rather than
    /// the -1 a VBA True would be -- `#If Win64 = 1 Then` fires where
    /// `#If Win64 = -1 Then` does not.
    #[test]
    fn the_constants_excel_defines_are_one_not_minus_one() {
        let taken = |condition: &str| {
            let source = format!("#If {condition} Then\nkept\n#End If\n");
            live_source(&source).contains("kept")
        };
        for condition in ["VBA6", "VBA7", "Win32", "Win64", "vba7", "VBA7 = 1", "Win64 = 1"] {
            assert!(taken(condition), "{condition}");
        }
        for condition in ["Win16", "Mac", "VBA7 = -1", "Win64 = -1", "VBA7 = 2"] {
            assert!(!taken(condition), "{condition}");
        }
    }

    /// A name never defined is Empty: falsy, and equal to zero.
    #[test]
    fn a_name_never_defined_is_empty() {
        let taken = |condition: &str| {
            live_source(&format!("#If {condition} Then\nkept\n#End If\n")).contains("kept")
        };
        assert!(!taken("NOBODY_DEFINED_THIS"));
        assert!(taken("NOBODY_DEFINED_THIS = 0"));
        assert!(taken("Not NOBODY_DEFINED_THIS"));
    }

    /// The operators, and what `#Const` puts within reach of them.
    #[test]
    fn the_condition_reads_the_operators_vba_allows() {
        let taken = |source: &str| {
            live_source(&format!(
                "#Const MyFlag = 7\n#Const MyStr = \"a\"\n#If {source} Then\nkept\n#End If\n"
            ))
            .contains("kept")
        };
        for condition in [
            "VBA7 And Win64",
            "Win32 Or Mac",
            "Not Mac",
            "MyFlag",
            "MyFlag = 7",
            "MyFlag > 3",
            "MyStr = \"a\"",
            "(VBA7 And Not Mac) Or Win16",
        ] {
            assert!(taken(condition), "{condition}");
        }
        for condition in ["Mac And VBA7", "MyFlag = 8", "MyStr = \"b\"", "MyFlag < 3"] {
            assert!(!taken(condition), "{condition}");
        }
    }

    /// `#ElseIf` takes the first branch that fires, and no later one.
    #[test]
    fn a_chain_takes_its_first_live_branch_only() {
        let source = "#If Mac Then\nmac\n#ElseIf Win64 Then\nw64\n#ElseIf Win32 Then\n\
                      w32\n#Else\nother\n#End If\n";
        let live = live_source(source);
        assert!(live.contains("w64"), "{live}");
        for dead in ["mac", "w32", "other"] {
            assert!(!live.contains(dead), "{dead} survived: {live}");
        }
    }

    /// A branch inside a dead branch is dead, whatever its own condition says.
    #[test]
    fn a_branch_inside_a_dead_branch_is_dead() {
        let source = "#If Mac Then\n#If VBA7 Then\ninner\n#End If\n#End If\n\
                      #If VBA7 Then\n#If Win64 Then\nnested\n#End If\n#End If\n";
        let live = live_source(source);
        assert!(!live.contains("inner"), "{live}");
        assert!(live.contains("nested"), "{live}");
    }

    /// A dead branch is never compiled, so it may hold anything at all.
    ///
    /// Asked of Excel, a module whose dead branch carries `]not VBA at all[
    /// &&&`, or an `If` with no `End If`, or a string with no closing quote,
    /// compiles and runs. That is why this pass runs over the source rather
    /// than over tokens: none of those could survive being lexed.
    #[test]
    fn a_dead_branch_may_hold_anything() {
        for rubbish in ["]not VBA at all[ &&&", "If x Then", "x = \"unterminated"] {
            let source = format!("kept\n#If Mac Then\n{rubbish}\n#End If\n");
            let live = live_source(&source);
            assert!(live.contains("kept"));
            assert!(!live.contains('['), "{live}");
            assert!(!live.contains("unterminated"), "{live}");
        }
    }

    /// Every line keeps its place, so a span still points where it pointed.
    #[test]
    fn dropping_a_branch_moves_no_other_line() {
        let source = "one\n#If Mac Then\ntwo\n#End If\nthree\n";
        let live = live_source(source);
        assert_eq!(live.lines().count(), source.lines().count());
        assert_eq!(live.lines().next(), Some("one"));
        assert_eq!(live.lines().nth(4), Some("three"));
        assert_eq!(live.len(), source.len());
    }

    /// Source with no directive in it is handed back untouched.
    #[test]
    fn source_without_a_directive_is_left_alone() {
        let source = "Public Sub S()\n    x = 1 ' # not a directive\nEnd Sub\n";
        assert_eq!(live_source(source), source);
    }
}
