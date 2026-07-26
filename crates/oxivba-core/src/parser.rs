// This Source Code Form is subject to the terms of the Mozilla Public
// License, v. 2.0. If a copy of the MPL was not distributed with this
// file, You can obtain one at https://mozilla.org/MPL/2.0/.

//! Recursive-descent parser for VBA.
//!
//! # Precedence
//!
//! Loosest to tightest:
//! `Imp` › `Eqv` › `Xor` › `Or` › `And` › `Not` › comparisons › `&` › `+ -` ›
//! `Mod` › `\` › `* /` › unary `-` › `^` › member access
//!
//! Two levels are routinely got wrong:
//!
//! - **`Not` is looser than comparison**, so `Not a = b` means `Not (a = b)`.
//! - **`\` and `Mod` are their own levels between `*` and `+`**, so
//!   `a + b \ c` is `a + (b \ c)`.
//!
//! And one that differs from the *same expression in a worksheet cell*:
//! `^` binds tighter than unary minus here, so **`-2^2` is `-4` in VBA** while
//! the Excel formula `=-2^2` is `4`. A migration that moves an expression
//! between the two has to account for it.
//!
//! # Recovery
//!
//! Nothing is ever dropped. A construct the parser does not understand becomes
//! [`Statement::Unknown`] holding the original text, so that unparsed input can
//! be counted and reported rather than silently vanishing.

use crate::ast::*;
use crate::lexer::{tokenize, LexError, Punct, Span, Token, TokenKind};

pub fn parse_module(source: &str) -> Result<Module, LexError> {
    let tokens = tokenize(source)?;
    let mut parser = Parser {
        src: source,
        tokens,
        pos: 0,
        terminated: true,
    };
    Ok(parser.parse_module())
}

struct Parser<'a> {
    src: &'a str,
    tokens: Vec<Token>,
    pos: usize,
    /// Whether the statement just parsed ended where a statement may end.
    ///
    /// If it did not, the parser stopped in the middle of a line it did not
    /// fully understand, and the whole line is recorded as unparsed rather than
    /// half-interpreted. Without this, leftover tokens get read as the start of
    /// a new statement — a stray `For` can then swallow the rest of a procedure.
    terminated: bool,
}

impl<'a> Parser<'a> {
    // -- token access -----------------------------------------------------

    fn kind(&self) -> &TokenKind {
        self.tokens
            .get(self.pos)
            .map(|t| &t.kind)
            .unwrap_or(&TokenKind::Eof)
    }

    fn kind_at(&self, ahead: usize) -> &TokenKind {
        self.tokens
            .get(self.pos + ahead)
            .map(|t| &t.kind)
            .unwrap_or(&TokenKind::Eof)
    }

    fn span(&self) -> Span {
        self.tokens
            .get(self.pos)
            .map(|t| t.span)
            .unwrap_or(Span {
                line: 0,
                start: self.src.len(),
                end: self.src.len(),
            })
    }

    fn at_eof(&self) -> bool {
        matches!(self.kind(), TokenKind::Eof)
    }

    /// Identifier text at an offset, lower-cased. VBA keywords are
    /// case-insensitive and the VBE rewrites their casing anyway.
    fn word_at(&self, ahead: usize) -> Option<String> {
        match self.kind_at(ahead) {
            TokenKind::Ident(s) => Some(s.to_ascii_lowercase()),
            _ => None,
        }
    }

    fn at_kw(&self, kw: &str) -> bool {
        self.word_at(0).as_deref() == Some(kw)
    }

    fn eat_kw(&mut self, kw: &str) -> bool {
        if self.at_kw(kw) {
            self.pos += 1;
            true
        } else {
            false
        }
    }

    fn at_punct(&self, p: Punct) -> bool {
        matches!(self.kind(), TokenKind::Punct(x) if *x == p)
    }

    fn eat_punct(&mut self, p: Punct) -> bool {
        if self.at_punct(p) {
            self.pos += 1;
            true
        } else {
            false
        }
    }

    fn at_eol(&self) -> bool {
        matches!(self.kind(), TokenKind::Eol) || self.at_eof()
    }

    /// Consume a statement terminator if one is next, and nothing else.
    ///
    /// Deliberately not greedy. Skipping ahead to the newline would swallow the
    /// `Else` of a one-line `If`, and would hide unparsed input instead of
    /// leaving it for the caller to record.
    fn end_statement(&mut self) {
        self.terminated = matches!(
            self.kind(),
            TokenKind::Eol | TokenKind::Punct(Punct::Colon) | TokenKind::Eof | TokenKind::Comment(_)
        );
        if matches!(
            self.kind(),
            TokenKind::Eol | TokenKind::Punct(Punct::Colon)
        ) {
            self.pos += 1;
        }
    }

    fn skip_blank(&mut self) {
        while matches!(self.kind(), TokenKind::Eol)
            || matches!(self.kind(), TokenKind::Punct(Punct::Colon))
        {
            self.pos += 1;
        }
    }

    /// The one or two leading keywords of the statement at the cursor,
    /// lower-cased and joined, e.g. `"end if"`. Used for block terminators.
    fn stmt_head(&self) -> String {
        let Some(first) = self.word_at(0) else {
            return String::new();
        };
        if matches!(first.as_str(), "end" | "exit" | "on" | "for" | "select") {
            if let Some(second) = self.word_at(1) {
                return format!("{first} {second}");
            }
        }
        first
    }

    fn at_block_end(&self, terminators: &[&str]) -> bool {
        if self.at_eof() {
            return true;
        }
        let head = self.stmt_head();
        terminators.iter().any(|t| {
            head == *t
                || (head.starts_with(*t) && t.contains(' '))
                || head.split(' ').next() == Some(*t)
        })
    }

    fn text_of(&self, span: Span) -> String {
        self.src
            .get(span.start..span.end)
            .unwrap_or_default()
            .trim()
            .to_string()
    }

    /// Skip to the end of the current physical line and return its text.
    fn consume_line(&mut self) -> (String, Span) {
        let start = self.span();
        let mut end = start;
        while !self.at_eol() {
            end = self.span();
            self.pos += 1;
        }
        let span = Span {
            line: start.line,
            start: start.start,
            end: end.end.max(start.end),
        };
        let text = self.text_of(span);
        self.end_statement();
        (text, span)
    }

    // -- module -----------------------------------------------------------

    fn parse_module(&mut self) -> Module {
        let mut items = Vec::new();
        loop {
            self.skip_blank();
            if self.at_eof() {
                break;
            }
            match self.parse_module_item() {
                Some(item) => items.push(item),
                None => {
                    let (text, span) = self.consume_line();
                    if !text.is_empty() {
                        items.push(ModuleItem::Unknown { text, span });
                    }
                }
            }
        }
        Module { items }
    }

    fn parse_module_item(&mut self) -> Option<ModuleItem> {
        let span = self.span();

        if let TokenKind::Comment(text) = self.kind() {
            let text = text.clone();
            self.pos += 1;
            self.end_statement();
            return Some(ModuleItem::Unknown {
                text: format!("'{text}"),
                span,
            });
        }
        if let TokenKind::Directive(text) = self.kind() {
            let text = text.clone();
            self.pos += 1;
            self.end_statement();
            return Some(ModuleItem::Unknown { text, span });
        }

        if self.at_kw("attribute") {
            return self.parse_attribute(span);
        }
        if self.at_kw("option") {
            return self.parse_option(span);
        }
        if self.at_kw("implements") {
            self.pos += 1;
            let interface = self.parse_qualified_name()?;
            self.end_statement();
            return Some(ModuleItem::Implements { interface, span });
        }

        let visibility = self.parse_visibility();
        let is_static = self.eat_kw("static");

        if self.at_kw("declare") {
            return self.parse_declare(visibility, span);
        }
        if self.at_kw("type") {
            return self.parse_type_def(visibility, span);
        }
        if self.at_kw("enum") {
            return self.parse_enum_def(visibility, span);
        }
        if self.at_kw("event") {
            self.pos += 1;
            let name = self.parse_ident()?;
            let params = self.parse_param_list();
            self.end_statement();
            return Some(ModuleItem::Event { name, params, span });
        }
        if self.at_kw("sub")
            || self.at_kw("function")
            || self.at_kw("property")
        {
            return self
                .parse_procedure(visibility, is_static, span)
                .map(ModuleItem::Procedure);
        }
        if self.at_kw("dim") || self.at_kw("const") || visibility != Visibility::Default {
            return self
                .parse_var_decl(visibility, is_static, span)
                .map(ModuleItem::Variables);
        }

        None
    }

    fn parse_attribute(&mut self, span: Span) -> Option<ModuleItem> {
        self.pos += 1;
        let name = self.parse_qualified_name()?;
        if !self.eat_punct(Punct::Eq) {
            return None;
        }
        let value = match self.kind() {
            TokenKind::Str(s) => {
                let s = s.clone();
                self.pos += 1;
                s
            }
            _ => {
                let (text, _) = self.consume_line();
                return Some(ModuleItem::Attribute {
                    name,
                    value: text,
                    span,
                });
            }
        };
        self.end_statement();
        Some(ModuleItem::Attribute { name, value, span })
    }

    fn parse_option(&mut self, span: Span) -> Option<ModuleItem> {
        self.pos += 1;
        let option = match self.word_at(0)?.as_str() {
            "explicit" => {
                self.pos += 1;
                ModuleOption::Explicit
            }
            "base" => {
                self.pos += 1;
                let n = match self.kind() {
                    TokenKind::Number(n) => *n as u32,
                    _ => 0,
                };
                self.pos += 1;
                ModuleOption::Base(n)
            }
            "compare" => {
                self.pos += 1;
                let mode = self.parse_ident().unwrap_or_default();
                ModuleOption::Compare(mode)
            }
            "private" => {
                self.pos += 1;
                self.eat_kw("module");
                ModuleOption::PrivateModule
            }
            _ => return None,
        };
        self.end_statement();
        Some(ModuleItem::Option(option, span))
    }

    fn parse_visibility(&mut self) -> Visibility {
        for (kw, vis) in [
            ("private", Visibility::Private),
            ("public", Visibility::Public),
            ("friend", Visibility::Friend),
            ("global", Visibility::Global),
        ] {
            if self.at_kw(kw) {
                self.pos += 1;
                return vis;
            }
        }
        Visibility::Default
    }

    fn parse_declare(&mut self, visibility: Visibility, span: Span) -> Option<ModuleItem> {
        self.pos += 1;
        let ptr_safe = self.eat_kw("ptrsafe");
        let is_function = if self.eat_kw("function") {
            true
        } else {
            self.eat_kw("sub");
            false
        };
        let name = self.parse_ident()?;
        self.eat_kw("lib");
        let lib = match self.kind() {
            TokenKind::Str(s) => {
                let s = s.clone();
                self.pos += 1;
                s
            }
            _ => String::new(),
        };
        let alias = if self.eat_kw("alias") {
            match self.kind() {
                TokenKind::Str(s) => {
                    let s = s.clone();
                    self.pos += 1;
                    Some(s)
                }
                _ => None,
            }
        } else {
            None
        };
        let params = self.parse_param_list();
        let return_type = self.parse_as_type();
        self.end_statement();
        Some(ModuleItem::ExternalProc(ExternalProc {
            visibility,
            is_function,
            ptr_safe,
            name,
            lib,
            alias,
            params,
            return_type,
            span,
        }))
    }

    fn parse_type_def(&mut self, visibility: Visibility, span: Span) -> Option<ModuleItem> {
        self.pos += 1;
        let name = self.parse_ident()?;
        self.end_statement();
        let mut fields = Vec::new();
        loop {
            self.skip_blank();
            if self.at_eof() || self.at_block_end(&["end type"]) {
                break;
            }
            if let Some(item) = self.parse_var_item() {
                fields.push(item);
            }
            self.end_statement();
        }
        self.consume_line();
        Some(ModuleItem::Type(TypeDef {
            visibility,
            name,
            fields,
            span,
        }))
    }

    fn parse_enum_def(&mut self, visibility: Visibility, span: Span) -> Option<ModuleItem> {
        self.pos += 1;
        let name = self.parse_ident()?;
        self.end_statement();
        let mut members = Vec::new();
        loop {
            self.skip_blank();
            if self.at_eof() || self.at_block_end(&["end enum"]) {
                break;
            }
            let Some(member) = self.parse_ident() else {
                self.consume_line();
                continue;
            };
            let value = if self.eat_punct(Punct::Eq) {
                self.parse_expr()
            } else {
                None
            };
            members.push((member, value));
            self.end_statement();
        }
        self.consume_line();
        Some(ModuleItem::Enum(EnumDef {
            visibility,
            name,
            members,
            span,
        }))
    }

    fn parse_procedure(
        &mut self,
        visibility: Visibility,
        is_static: bool,
        span: Span,
    ) -> Option<Procedure> {
        let kind = if self.eat_kw("sub") {
            ProcKind::Sub
        } else if self.eat_kw("function") {
            ProcKind::Function
        } else {
            self.pos += 1; // property
            if self.eat_kw("get") {
                ProcKind::PropertyGet
            } else if self.eat_kw("let") {
                ProcKind::PropertyLet
            } else {
                self.eat_kw("set");
                ProcKind::PropertySet
            }
        };
        let name = self.parse_ident()?;
        let params = self.parse_param_list();
        let return_type = self.parse_as_type();
        self.end_statement();

        let terminator = match kind {
            ProcKind::Sub => "end sub",
            ProcKind::Function => "end function",
            _ => "end property",
        };
        let body = self.parse_block(&[terminator]);
        self.consume_line();

        Some(Procedure {
            kind,
            visibility,
            is_static,
            name,
            params,
            return_type,
            body,
            span,
        })
    }

    fn parse_param_list(&mut self) -> Vec<Param> {
        let mut params = Vec::new();
        if !self.eat_punct(Punct::LParen) {
            return params;
        }
        if self.eat_punct(Punct::RParen) {
            return params;
        }
        loop {
            let optional = self.eat_kw("optional");
            let mode = if self.eat_kw("byval") {
                ParamMode::ByVal
            } else if self.eat_kw("byref") {
                ParamMode::ByRef
            } else if self.eat_kw("paramarray") {
                ParamMode::ParamArray
            } else {
                // The default is ByRef, which is why a callee can quietly
                // reassign the caller's variable.
                ParamMode::ByRef
            };
            let Some(name) = self.parse_ident() else {
                break;
            };
            let is_array = if self.at_punct(Punct::LParen)
                && matches!(self.kind_at(1), TokenKind::Punct(Punct::RParen))
            {
                self.pos += 2;
                true
            } else {
                false
            };
            let type_name = self.parse_as_type().unwrap_or_else(TypeName::implicit);
            let default = if self.eat_punct(Punct::Eq) {
                self.parse_expr()
            } else {
                None
            };
            params.push(Param {
                mode,
                optional,
                name,
                is_array,
                type_name,
                default,
            });
            if self.eat_punct(Punct::Comma) {
                continue;
            }
            self.eat_punct(Punct::RParen);
            break;
        }
        params
    }

    fn parse_as_type(&mut self) -> Option<TypeName> {
        if !self.eat_kw("as") {
            return None;
        }
        self.eat_kw("new");
        let name = self.parse_qualified_name()?;
        Some(TypeName { name, suffix: None })
    }

    fn parse_ident(&mut self) -> Option<String> {
        let name = match self.kind() {
            TokenKind::Ident(s) => s.clone(),
            _ => return None,
        };
        self.pos += 1;
        if let TokenKind::TypeSuffix(_) = self.kind() {
            self.pos += 1;
        }
        Some(name)
    }

    fn parse_qualified_name(&mut self) -> Option<String> {
        let mut name = self.parse_ident()?;
        while self.at_punct(Punct::Dot) {
            if let TokenKind::Ident(next) = self.kind_at(1).clone() {
                self.pos += 2;
                name.push('.');
                name.push_str(&next);
            } else {
                break;
            }
        }
        Some(name)
    }

    fn parse_var_decl(
        &mut self,
        visibility: Visibility,
        is_static: bool,
        span: Span,
    ) -> Option<VarDecl> {
        let is_const = self.eat_kw("const");
        if !is_const {
            self.eat_kw("dim");
        }
        let mut items = Vec::new();
        loop {
            let Some(item) = self.parse_var_item() else {
                break;
            };
            items.push(item);
            if !self.eat_punct(Punct::Comma) {
                break;
            }
        }
        self.end_statement();
        if items.is_empty() {
            return None;
        }
        Some(VarDecl {
            visibility,
            is_const,
            is_static,
            items,
            span,
        })
    }

    fn parse_var_item(&mut self) -> Option<VarItem> {
        let with_events = self.eat_kw("withevents");
        let name = self.parse_ident()?;
        let array_bounds = if self.at_punct(Punct::LParen) {
            Some(self.parse_array_bounds())
        } else {
            None
        };
        let type_name = self.parse_as_type().unwrap_or_else(TypeName::implicit);
        let value = if self.eat_punct(Punct::Eq) {
            self.parse_expr()
        } else {
            None
        };
        Some(VarItem {
            name,
            array_bounds,
            type_name,
            with_events,
            value,
        })
    }

    fn parse_array_bounds(&mut self) -> Vec<ArrayBound> {
        let mut bounds = Vec::new();
        self.eat_punct(Punct::LParen);
        if self.eat_punct(Punct::RParen) {
            return bounds; // dynamic array: Dim a()
        }
        loop {
            let Some(first) = self.parse_expr() else {
                break;
            };
            let bound = if self.eat_kw("to") {
                match self.parse_expr() {
                    Some(upper) => ArrayBound {
                        lower: Some(first),
                        upper,
                    },
                    None => break,
                }
            } else {
                ArrayBound {
                    lower: None,
                    upper: first,
                }
            };
            bounds.push(bound);
            if self.eat_punct(Punct::Comma) {
                continue;
            }
            self.eat_punct(Punct::RParen);
            break;
        }
        bounds
    }

    // -- statements --------------------------------------------------------

    fn parse_block(&mut self, terminators: &[&str]) -> Vec<Statement> {
        let mut body = Vec::new();
        loop {
            self.skip_blank();
            if self.at_eof() || self.at_block_end(terminators) {
                break;
            }
            let line_start = self.span();
            let before = self.pos;
            self.terminated = false;
            let stmt = self.parse_statement();

            if self.pos == before {
                // Nothing consumed: force progress rather than spin.
                let (text, span) = self.consume_line();
                if !text.is_empty() {
                    body.push(Statement::Unknown { text, span });
                }
                continue;
            }

            if self.terminated || self.at_block_end(terminators) {
                body.push(stmt);
                continue;
            }

            // The line was only partly understood. Record it whole rather than
            // keeping a fragment and re-reading the rest as new statements.
            let (_, tail) = self.consume_line();
            let span = Span {
                line: line_start.line,
                start: line_start.start,
                end: tail.end.max(line_start.end),
            };
            body.push(Statement::Unknown {
                text: self.text_of(span),
                span,
            });
        }
        body
    }

    fn parse_statement(&mut self) -> Statement {
        let span = self.span();

        if let TokenKind::Comment(text) = self.kind() {
            let text = text.clone();
            self.pos += 1;
            self.end_statement();
            return Statement::Comment { text, span };
        }
        if let TokenKind::Directive(text) = self.kind() {
            let text = text.clone();
            self.pos += 1;
            self.end_statement();
            return Statement::Directive { text, span };
        }
        if let TokenKind::LineNumber(value) = *self.kind() {
            self.pos += 1;
            // A line number labels the statement that follows it on the same
            // line, so it is not itself terminated.
            self.terminated = true;
            return Statement::LineNumber { value, span };
        }

        // A label: `foo:` at the head of a line, distinguished from `foo := x`
        // and from `foo: bar` only by the colon coming straight after the name.
        if matches!(self.kind(), TokenKind::Ident(_))
            && matches!(self.kind_at(1), TokenKind::Punct(Punct::Colon))
        {
            let name = self.parse_ident().unwrap_or_default();
            self.pos += 1;
            self.terminated = true;
            return Statement::Label { name, span };
        }

        match self.word_at(0).as_deref() {
            Some("dim") | Some("const") | Some("static") => {
                let is_static = self.eat_kw("static");
                if let Some(decl) = self.parse_var_decl(Visibility::Default, is_static, span) {
                    return Statement::Dim(decl);
                }
            }
            Some("redim") => return self.parse_redim(span),
            Some("erase") => {
                self.pos += 1;
                let mut targets = Vec::new();
                while let Some(e) = self.parse_expr() {
                    targets.push(e);
                    if !self.eat_punct(Punct::Comma) {
                        break;
                    }
                }
                self.end_statement();
                return Statement::Erase { targets, span };
            }
            Some("if") => return self.parse_if(span),
            Some("select") => return self.parse_select_case(span),
            Some("for") => return self.parse_for(span),
            Some("do") => return self.parse_do(span),
            Some("while") => return self.parse_while(span),
            Some("with") => return self.parse_with(span),
            Some("on") => return self.parse_on_error(span),
            Some("resume") => {
                self.pos += 1;
                let target = if self.eat_kw("next") {
                    ResumeTarget::Next
                } else if let Some(label) = self.parse_label_ref() {
                    ResumeTarget::Label(label)
                } else {
                    ResumeTarget::Same
                };
                self.end_statement();
                return Statement::Resume { target, span };
            }
            Some("goto") => {
                self.pos += 1;
                let label = self.parse_label_ref().unwrap_or_default();
                self.end_statement();
                return Statement::GoTo { label, span };
            }
            Some("gosub") => {
                self.pos += 1;
                let label = self.parse_label_ref().unwrap_or_default();
                self.end_statement();
                return Statement::GoSub { label, span };
            }
            Some("return") => {
                self.pos += 1;
                self.end_statement();
                return Statement::Return { span };
            }
            Some("exit") => {
                self.pos += 1;
                let what = match self.word_at(0).as_deref() {
                    Some("sub") => ExitKind::Sub,
                    Some("function") => ExitKind::Function,
                    Some("property") => ExitKind::Property,
                    Some("for") => ExitKind::For,
                    _ => ExitKind::Do,
                };
                self.pos += 1;
                self.end_statement();
                return Statement::Exit { what, span };
            }
            Some("stop") => {
                self.pos += 1;
                self.end_statement();
                return Statement::Stop { span };
            }
            Some("end") if matches!(self.kind_at(1), TokenKind::Eol | TokenKind::Eof) => {
                self.pos += 1;
                self.end_statement();
                return Statement::End { span };
            }
            Some("set") => {
                self.pos += 1;
                let target = self.parse_postfix();
                self.eat_punct(Punct::Eq);
                let value = self.parse_expr();
                self.end_statement();
                if let (Some(target), Some(value)) = (target, value) {
                    return Statement::SetAssign {
                        target,
                        value,
                        span,
                    };
                }
                return Statement::Unknown {
                    text: self.text_of(span),
                    span,
                };
            }
            Some("let") => {
                self.pos += 1;
            }
            Some("call") => {
                self.pos += 1;
                let target = self.parse_expr();
                self.end_statement();
                return match target {
                    Some(target) => Statement::Call {
                        target,
                        explicit_call: true,
                        span,
                    },
                    None => Statement::Unknown {
                        text: self.text_of(span),
                        span,
                    },
                };
            }
            _ => {}
        }

        self.parse_assign_or_call(span)
    }

    fn parse_label_ref(&mut self) -> Option<String> {
        match self.kind().clone() {
            TokenKind::Ident(name) => {
                self.pos += 1;
                Some(name)
            }
            TokenKind::Number(n) => {
                self.pos += 1;
                Some(format!("{}", n as i64))
            }
            _ => None,
        }
    }

    fn parse_assign_or_call(&mut self, span: Span) -> Statement {
        // `=` is both assignment and equality in VBA, told apart only by
        // position. Parsing the target as a full expression would let the
        // comparison level eat the `=` and turn `x = 1` into a comparison that
        // is then thrown away. An assignment target is always a postfix
        // expression, so stopping there keeps the `=` for us.
        let Some(first) = self.parse_postfix() else {
            let (text, span) = self.consume_line();
            return Statement::Unknown { text, span };
        };

        if self.eat_punct(Punct::Eq) {
            let value = self.parse_expr();
            self.end_statement();
            return match value {
                Some(value) => Statement::Assign {
                    target: first,
                    value,
                    span,
                },
                None => Statement::Unknown {
                    text: self.text_of(span),
                    span,
                },
            };
        }

        // A bare call with unparenthesised arguments: `MsgBox "hi", vbOK`.
        if !self.at_eol() && !self.at_punct(Punct::Colon) {
            let mut args = vec![Argument {
                name: None,
                value: None,
            }];
            args.clear();
            let target = self.parse_bare_call_args(first, &mut args, span);
            self.end_statement();
            return Statement::Call {
                target,
                explicit_call: false,
                span,
            };
        }

        self.end_statement();
        Statement::Call {
            target: first,
            explicit_call: false,
            span,
        }
    }

    fn parse_bare_call_args(
        &mut self,
        callee: Expr,
        args: &mut Vec<Argument>,
        span: Span,
    ) -> Expr {
        loop {
            if self.at_eol() || self.at_punct(Punct::Colon) {
                break;
            }
            // `Print a; b` separates with a semicolon.
            if self.eat_punct(Punct::Semicolon) {
                continue;
            }
            if self.at_punct(Punct::Comma) {
                self.pos += 1;
                args.push(Argument {
                    name: None,
                    value: None,
                });
                continue;
            }
            match self.parse_argument() {
                Some(arg) => args.push(arg),
                None => break,
            }
            if !self.eat_punct(Punct::Comma) && !self.at_punct(Punct::Semicolon) {
                break;
            }
        }
        Expr::Index {
            target: Box::new(callee),
            args: std::mem::take(args),
            span,
        }
    }

    fn parse_redim(&mut self, span: Span) -> Statement {
        self.pos += 1;
        let preserve = self.eat_kw("preserve");
        let mut items = Vec::new();
        loop {
            let Some(item) = self.parse_var_item() else {
                break;
            };
            items.push(item);
            if !self.eat_punct(Punct::Comma) {
                break;
            }
        }
        self.end_statement();
        Statement::ReDim {
            preserve,
            items,
            span,
        }
    }

    fn parse_if(&mut self, span: Span) -> Statement {
        self.pos += 1;
        let condition = self.parse_expr().unwrap_or(Expr::Literal(Literal::Empty, span));
        self.eat_kw("then");

        // A one-line `If` has its body on the same line and no `End If`.
        if !self.at_eol() {
            let mut then_body = Vec::new();
            let mut else_body = None;
            while !self.at_eol() && !self.at_kw("else") {
                let before = self.pos;
                then_body.push(self.parse_statement());
                if self.pos == before {
                    break;
                }
            }
            if self.eat_kw("else") {
                let mut body = Vec::new();
                while !self.at_eol() {
                    let before = self.pos;
                    body.push(self.parse_statement());
                    if self.pos == before {
                        break;
                    }
                }
                else_body = Some(body);
            }
            self.end_statement();
            return Statement::If(IfStmt {
                condition,
                then_body,
                else_ifs: Vec::new(),
                else_body,
                single_line: true,
                span,
            });
        }

        self.end_statement();
        let then_body = self.parse_block(&["elseif", "else", "end if"]);
        let mut else_ifs = Vec::new();
        let mut else_body = None;

        loop {
            if self.at_kw("elseif") {
                self.pos += 1;
                let cond = self.parse_expr().unwrap_or(Expr::Literal(Literal::Empty, span));
                self.eat_kw("then");
                self.end_statement();
                let body = self.parse_block(&["elseif", "else", "end if"]);
                else_ifs.push((cond, body));
                continue;
            }
            if self.at_kw("else") {
                self.pos += 1;
                self.end_statement();
                else_body = Some(self.parse_block(&["end if"]));
                continue;
            }
            break;
        }
        self.consume_line();

        Statement::If(IfStmt {
            condition,
            then_body,
            else_ifs,
            else_body,
            single_line: false,
            span,
        })
    }

    fn parse_select_case(&mut self, span: Span) -> Statement {
        self.pos += 2; // Select Case
        let subject = self.parse_expr().unwrap_or(Expr::Literal(Literal::Empty, span));
        self.end_statement();

        let mut cases = Vec::new();
        let mut case_else = None;
        loop {
            self.skip_blank();
            if self.at_eof() || self.at_block_end(&["end select"]) {
                break;
            }
            if !self.eat_kw("case") {
                let (text, span) = self.consume_line();
                if !text.is_empty() {
                    cases.push(CaseClause {
                        labels: Vec::new(),
                        body: vec![Statement::Unknown { text, span }],
                    });
                }
                continue;
            }
            if self.eat_kw("else") {
                self.end_statement();
                case_else = Some(self.parse_block(&["case", "end select"]));
                continue;
            }
            let labels = self.parse_case_labels();
            self.end_statement();
            let body = self.parse_block(&["case", "end select"]);
            cases.push(CaseClause { labels, body });
        }
        self.consume_line();

        Statement::SelectCase(SelectCaseStmt {
            subject,
            cases,
            case_else,
            span,
        })
    }

    fn parse_case_labels(&mut self) -> Vec<CaseLabel> {
        let mut labels = Vec::new();
        loop {
            if self.eat_kw("is") {
                let op = self.parse_comparison_op().unwrap_or(BinaryOp::Eq);
                if let Some(value) = self.parse_expr() {
                    labels.push(CaseLabel::Compare(op, value));
                }
            } else {
                let Some(first) = self.parse_expr() else {
                    break;
                };
                if self.eat_kw("to") {
                    match self.parse_expr() {
                        Some(upper) => labels.push(CaseLabel::Range(first, upper)),
                        None => break,
                    }
                } else {
                    labels.push(CaseLabel::Value(first));
                }
            }
            if !self.eat_punct(Punct::Comma) {
                break;
            }
        }
        labels
    }

    fn parse_comparison_op(&mut self) -> Option<BinaryOp> {
        let op = match self.kind() {
            TokenKind::Punct(Punct::Eq) => BinaryOp::Eq,
            TokenKind::Punct(Punct::Ne) => BinaryOp::Ne,
            TokenKind::Punct(Punct::Lt) => BinaryOp::Lt,
            TokenKind::Punct(Punct::Le) => BinaryOp::Le,
            TokenKind::Punct(Punct::Gt) => BinaryOp::Gt,
            TokenKind::Punct(Punct::Ge) => BinaryOp::Ge,
            _ => return None,
        };
        self.pos += 1;
        Some(op)
    }

    fn parse_for(&mut self, span: Span) -> Statement {
        self.pos += 1;
        if self.eat_kw("each") {
            let item = self.parse_expr().unwrap_or(Expr::Literal(Literal::Empty, span));
            self.eat_kw("in");
            let collection = self.parse_expr().unwrap_or(Expr::Literal(Literal::Empty, span));
            self.end_statement();
            let body = self.parse_block(&["next"]);
            self.consume_line();
            return Statement::ForEach(ForEachStmt {
                item,
                collection,
                body,
                span,
            });
        }

        // Same reason as in an assignment: the counter must not eat the `=`.
        let counter = self.parse_postfix().unwrap_or(Expr::Literal(Literal::Empty, span));
        self.eat_punct(Punct::Eq);
        let from = self.parse_expr().unwrap_or(Expr::Literal(Literal::Empty, span));
        self.eat_kw("to");
        let to = self.parse_expr().unwrap_or(Expr::Literal(Literal::Empty, span));
        let step = if self.eat_kw("step") {
            self.parse_expr()
        } else {
            None
        };
        self.end_statement();
        let body = self.parse_block(&["next"]);
        self.consume_line();

        Statement::For(ForStmt {
            counter,
            from,
            to,
            step,
            body,
            span,
        })
    }

    fn parse_do(&mut self, span: Span) -> Statement {
        self.pos += 1;
        let pre = self.parse_loop_test();
        self.end_statement();
        let body = self.parse_block(&["loop"]);
        self.eat_kw("loop");
        let post = self.parse_loop_test();
        self.end_statement();
        Statement::Do(DoStmt {
            pre,
            post,
            body,
            span,
        })
    }

    fn parse_loop_test(&mut self) -> Option<LoopTest> {
        let until = if self.eat_kw("until") {
            true
        } else if self.eat_kw("while") {
            false
        } else {
            return None;
        };
        let condition = self.parse_expr()?;
        Some(LoopTest { until, condition })
    }

    fn parse_while(&mut self, span: Span) -> Statement {
        self.pos += 1;
        let condition = self.parse_expr().unwrap_or(Expr::Literal(Literal::Empty, span));
        self.end_statement();
        let body = self.parse_block(&["wend"]);
        self.consume_line();
        Statement::While {
            condition,
            body,
            span,
        }
    }

    fn parse_with(&mut self, span: Span) -> Statement {
        self.pos += 1;
        let subject = self.parse_expr().unwrap_or(Expr::Literal(Literal::Empty, span));
        self.end_statement();
        let body = self.parse_block(&["end with"]);
        self.consume_line();
        Statement::With {
            subject,
            body,
            span,
        }
    }

    fn parse_on_error(&mut self, span: Span) -> Statement {
        self.pos += 1; // On
        if !self.eat_kw("error") {
            let (text, span) = self.consume_line();
            return Statement::Unknown { text, span };
        }
        if self.eat_kw("resume") {
            self.eat_kw("next");
            self.end_statement();
            return Statement::OnError(OnError::ResumeNext { span });
        }
        self.eat_kw("goto");
        let target = self.parse_label_ref().unwrap_or_default();
        self.end_statement();
        if target == "0" {
            Statement::OnError(OnError::Disable { span })
        } else {
            Statement::OnError(OnError::Goto {
                label: target,
                span,
            })
        }
    }

    // -- expressions -------------------------------------------------------

    pub fn parse_expr(&mut self) -> Option<Expr> {
        self.parse_imp()
    }

    fn parse_imp(&mut self) -> Option<Expr> {
        let mut lhs = self.parse_eqv()?;
        while self.at_kw("imp") {
            self.pos += 1;
            let rhs = self.parse_eqv()?;
            lhs = binary(BinaryOp::Imp, lhs, rhs);
        }
        Some(lhs)
    }

    fn parse_eqv(&mut self) -> Option<Expr> {
        let mut lhs = self.parse_xor()?;
        while self.at_kw("eqv") {
            self.pos += 1;
            let rhs = self.parse_xor()?;
            lhs = binary(BinaryOp::Eqv, lhs, rhs);
        }
        Some(lhs)
    }

    fn parse_xor(&mut self) -> Option<Expr> {
        let mut lhs = self.parse_or()?;
        while self.at_kw("xor") {
            self.pos += 1;
            let rhs = self.parse_or()?;
            lhs = binary(BinaryOp::Xor, lhs, rhs);
        }
        Some(lhs)
    }

    fn parse_or(&mut self) -> Option<Expr> {
        let mut lhs = self.parse_and()?;
        while self.at_kw("or") {
            self.pos += 1;
            let rhs = self.parse_and()?;
            lhs = binary(BinaryOp::Or, lhs, rhs);
        }
        Some(lhs)
    }

    fn parse_and(&mut self) -> Option<Expr> {
        let mut lhs = self.parse_not()?;
        while self.at_kw("and") {
            self.pos += 1;
            let rhs = self.parse_not()?;
            lhs = binary(BinaryOp::And, lhs, rhs);
        }
        Some(lhs)
    }

    /// `Not` sits *below* comparison, so `Not a = b` is `Not (a = b)`.
    fn parse_not(&mut self) -> Option<Expr> {
        if self.at_kw("not") {
            let span = self.span();
            self.pos += 1;
            let operand = self.parse_not()?;
            return Some(Expr::Unary {
                op: UnaryOp::Not,
                operand: Box::new(operand),
                span,
            });
        }
        self.parse_comparison()
    }

    fn parse_comparison(&mut self) -> Option<Expr> {
        let mut lhs = self.parse_concat()?;
        loop {
            let op = if self.at_kw("is") {
                self.pos += 1;
                BinaryOp::Is
            } else if self.at_kw("like") {
                self.pos += 1;
                BinaryOp::Like
            } else if let Some(op) = self.parse_comparison_op() {
                op
            } else {
                return Some(lhs);
            };
            let rhs = self.parse_concat()?;
            lhs = binary(op, lhs, rhs);
        }
    }

    fn parse_concat(&mut self) -> Option<Expr> {
        let mut lhs = self.parse_additive()?;
        while self.at_punct(Punct::Amp) {
            self.pos += 1;
            let rhs = self.parse_additive()?;
            lhs = binary(BinaryOp::Concat, lhs, rhs);
        }
        Some(lhs)
    }

    fn parse_additive(&mut self) -> Option<Expr> {
        let mut lhs = self.parse_mod()?;
        loop {
            let op = if self.at_punct(Punct::Plus) {
                BinaryOp::Add
            } else if self.at_punct(Punct::Minus) {
                BinaryOp::Sub
            } else {
                return Some(lhs);
            };
            self.pos += 1;
            let rhs = self.parse_mod()?;
            lhs = binary(op, lhs, rhs);
        }
    }

    fn parse_mod(&mut self) -> Option<Expr> {
        let mut lhs = self.parse_int_div()?;
        while self.at_kw("mod") {
            self.pos += 1;
            let rhs = self.parse_int_div()?;
            lhs = binary(BinaryOp::Mod, lhs, rhs);
        }
        Some(lhs)
    }

    fn parse_int_div(&mut self) -> Option<Expr> {
        let mut lhs = self.parse_multiplicative()?;
        while self.at_punct(Punct::BackSlash) {
            self.pos += 1;
            let rhs = self.parse_multiplicative()?;
            lhs = binary(BinaryOp::IntDiv, lhs, rhs);
        }
        Some(lhs)
    }

    fn parse_multiplicative(&mut self) -> Option<Expr> {
        let mut lhs = self.parse_unary()?;
        loop {
            let op = if self.at_punct(Punct::Star) {
                BinaryOp::Mul
            } else if self.at_punct(Punct::Slash) {
                BinaryOp::Div
            } else {
                return Some(lhs);
            };
            self.pos += 1;
            let rhs = self.parse_unary()?;
            lhs = binary(op, lhs, rhs);
        }
    }

    /// Unary minus binds *looser* than `^` here, so `-2^2` is `-4`. The same
    /// text in a worksheet cell evaluates to `4`.
    fn parse_unary(&mut self) -> Option<Expr> {
        let span = self.span();
        if self.at_punct(Punct::Minus) {
            self.pos += 1;
            let operand = self.parse_unary()?;
            return Some(Expr::Unary {
                op: UnaryOp::Neg,
                operand: Box::new(operand),
                span,
            });
        }
        if self.at_punct(Punct::Plus) {
            self.pos += 1;
            let operand = self.parse_unary()?;
            return Some(Expr::Unary {
                op: UnaryOp::Plus,
                operand: Box::new(operand),
                span,
            });
        }
        self.parse_power()
    }

    fn parse_power(&mut self) -> Option<Expr> {
        let mut lhs = self.parse_postfix()?;
        while self.at_punct(Punct::Caret) {
            self.pos += 1;
            let rhs = self.parse_postfix()?;
            lhs = binary(BinaryOp::Pow, lhs, rhs);
        }
        Some(lhs)
    }

    fn parse_postfix(&mut self) -> Option<Expr> {
        let mut expr = self.parse_primary()?;
        loop {
            let span = self.span();
            if self.at_punct(Punct::Dot) {
                if let TokenKind::Ident(name) = self.kind_at(1).clone() {
                    self.pos += 2;
                    if matches!(self.kind(), TokenKind::TypeSuffix(_)) {
                        self.pos += 1;
                    }
                    expr = Expr::Member {
                        object: Box::new(expr),
                        name,
                        span,
                    };
                    continue;
                }
                break;
            }
            if self.at_punct(Punct::LParen) {
                self.pos += 1;
                let args = self.parse_arguments();
                expr = Expr::Index {
                    target: Box::new(expr),
                    args,
                    span,
                };
                continue;
            }
            break;
        }
        Some(expr)
    }

    fn parse_arguments(&mut self) -> Vec<Argument> {
        let mut args = Vec::new();
        if self.eat_punct(Punct::RParen) {
            return args;
        }
        loop {
            if self.at_punct(Punct::Comma) {
                // An omitted positional argument keeps its slot.
                self.pos += 1;
                args.push(Argument {
                    name: None,
                    value: None,
                });
                continue;
            }
            match self.parse_argument() {
                Some(arg) => args.push(arg),
                None => break,
            }
            if self.eat_punct(Punct::Comma) {
                continue;
            }
            break;
        }
        self.eat_punct(Punct::RParen);
        args
    }

    fn parse_argument(&mut self) -> Option<Argument> {
        // `name:=value`
        if let (TokenKind::Ident(name), TokenKind::Punct(Punct::Assign)) =
            (self.kind().clone(), self.kind_at(1))
        {
            self.pos += 2;
            let value = self.parse_expr()?;
            return Some(Argument {
                name: Some(name),
                value: Some(value),
            });
        }
        let value = self.parse_expr()?;
        Some(Argument {
            name: None,
            value: Some(value),
        })
    }

    fn parse_primary(&mut self) -> Option<Expr> {
        let span = self.span();
        match self.kind().clone() {
            TokenKind::Number(n) => {
                self.pos += 1;
                if matches!(self.kind(), TokenKind::TypeSuffix(_)) {
                    self.pos += 1;
                }
                Some(Expr::Literal(Literal::Number(n), span))
            }
            TokenKind::Str(s) => {
                self.pos += 1;
                Some(Expr::Literal(Literal::Str(s), span))
            }
            TokenKind::DateLit(s) => {
                self.pos += 1;
                Some(Expr::Literal(Literal::Date(s), span))
            }
            TokenKind::Punct(Punct::LParen) => {
                self.pos += 1;
                let inner = self.parse_expr()?;
                self.eat_punct(Punct::RParen);
                Some(inner)
            }
            // `.Value` with no object: a member of the enclosing `With`.
            TokenKind::Punct(Punct::Dot) => {
                if let TokenKind::Ident(name) = self.kind_at(1).clone() {
                    self.pos += 2;
                    Some(Expr::WithMember(name, span))
                } else {
                    None
                }
            }
            TokenKind::Ident(name) => {
                let lower = name.to_ascii_lowercase();
                match lower.as_str() {
                    "true" => {
                        self.pos += 1;
                        return Some(Expr::Literal(Literal::Bool(true), span));
                    }
                    "false" => {
                        self.pos += 1;
                        return Some(Expr::Literal(Literal::Bool(false), span));
                    }
                    "empty" => {
                        self.pos += 1;
                        return Some(Expr::Literal(Literal::Empty, span));
                    }
                    "null" => {
                        self.pos += 1;
                        return Some(Expr::Literal(Literal::Null, span));
                    }
                    "nothing" => {
                        self.pos += 1;
                        return Some(Expr::Literal(Literal::Nothing, span));
                    }
                    "new" => {
                        self.pos += 1;
                        let type_name = self.parse_qualified_name()?;
                        return Some(Expr::New { type_name, span });
                    }
                    "typeof" => {
                        self.pos += 1;
                        let operand = self.parse_postfix()?;
                        self.eat_kw("is");
                        let type_name = self.parse_qualified_name().unwrap_or_default();
                        return Some(Expr::TypeOf {
                            operand: Box::new(operand),
                            type_name,
                            span,
                        });
                    }
                    // Words that can only start a statement; stop the expression
                    // so the caller can see them.
                    "then" | "else" | "elseif" | "to" | "step" | "in" | "as" | "end" | "next"
                    | "loop" | "wend" | "case" | "until" | "while" => return None,
                    _ => {}
                }
                self.pos += 1;
                if matches!(self.kind(), TokenKind::TypeSuffix(_)) {
                    self.pos += 1;
                }
                Some(Expr::Ident(name, span))
            }
            _ => None,
        }
    }
}

fn binary(op: BinaryOp, lhs: Expr, rhs: Expr) -> Expr {
    let span = lhs.span();
    Expr::Binary {
        op,
        lhs: Box::new(lhs),
        rhs: Box::new(rhs),
        span,
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    fn module(src: &str) -> Module {
        parse_module(src).expect("should parse")
    }

    fn only_proc(src: &str) -> Procedure {
        let m = module(src);
        m.items
            .into_iter()
            .find_map(|i| match i {
                ModuleItem::Procedure(p) => Some(p),
                _ => None,
            })
            .expect("a procedure")
    }

    fn expr(src: &str) -> Expr {
        let proc = only_proc(&format!("Sub T()\nx = {src}\nEnd Sub"));
        match proc.body.into_iter().find(|s| matches!(s, Statement::Assign { .. })) {
            Some(Statement::Assign { value, .. }) => value,
            other => panic!("expected an assignment, got {other:?}"),
        }
    }

    fn op_of(e: &Expr) -> BinaryOp {
        match e {
            Expr::Binary { op, .. } => *op,
            other => panic!("expected a binary expression, got {other:?}"),
        }
    }

    #[test]
    fn not_binds_looser_than_comparison() {
        // Not a = b  ==  Not (a = b)
        let e = expr("Not a = b");
        match e {
            Expr::Unary {
                op: UnaryOp::Not,
                operand,
                ..
            } => assert_eq!(op_of(&operand), BinaryOp::Eq),
            other => panic!("expected Not at the top, got {other:?}"),
        }
    }

    #[test]
    fn integer_division_and_mod_sit_between_mul_and_add() {
        // a + b \ c  ==  a + (b \ c)
        let e = expr("a + b \\ c");
        assert_eq!(op_of(&e), BinaryOp::Add);
        match &e {
            Expr::Binary { rhs, .. } => assert_eq!(op_of(rhs), BinaryOp::IntDiv),
            _ => unreachable!(),
        }
        // a \ b Mod c  ==  (a \ b) Mod c
        assert_eq!(op_of(&expr("a \\ b Mod c")), BinaryOp::Mod);
    }

    #[test]
    fn power_binds_tighter_than_unary_minus() {
        // -2 ^ 2 is -4 in VBA. The same text as a worksheet formula is 4.
        let e = expr("-2 ^ 2");
        match e {
            Expr::Unary {
                op: UnaryOp::Neg,
                operand,
                ..
            } => assert_eq!(op_of(&operand), BinaryOp::Pow),
            other => panic!("expected negation of a power, got {other:?}"),
        }
    }

    #[test]
    fn logical_operators_ladder_correctly() {
        // a And b Or c  ==  (a And b) Or c
        assert_eq!(op_of(&expr("a And b Or c")), BinaryOp::Or);
        // a Or b Xor c  ==  (a Or b) Xor c
        assert_eq!(op_of(&expr("a Or b Xor c")), BinaryOp::Xor);
        assert_eq!(op_of(&expr("a Xor b Eqv c")), BinaryOp::Eqv);
        assert_eq!(op_of(&expr("a Eqv b Imp c")), BinaryOp::Imp);
    }

    #[test]
    fn member_chains_and_calls_parse() {
        let e = expr("Application.WorksheetFunction.Sum(A, B)");
        assert_eq!(
            e.dotted_name().as_deref(),
            Some("Application.WorksheetFunction.Sum")
        );
    }

    #[test]
    fn named_and_omitted_arguments_keep_their_slots() {
        let e = expr("Foo(1, , key:=2)");
        match e {
            Expr::Index { args, .. } => {
                assert_eq!(args.len(), 3);
                assert!(args[1].value.is_none());
                assert_eq!(args[2].name.as_deref(), Some("key"));
            }
            other => panic!("expected a call, got {other:?}"),
        }
    }

    #[test]
    fn procedure_signature_defaults_to_byref() {
        let p = only_proc("Private Sub Foo(a As Long, ByVal b As String, Optional c = 1)\nEnd Sub");
        assert_eq!(p.visibility, Visibility::Private);
        assert_eq!(p.params.len(), 3);
        assert_eq!(p.params[0].mode, ParamMode::ByRef);
        assert_eq!(p.params[1].mode, ParamMode::ByVal);
        assert!(p.params[2].optional);
        assert!(p.params[2].default.is_some());
    }

    #[test]
    fn block_and_single_line_if_stay_distinguishable() {
        let block = only_proc("Sub T()\nIf a Then\nb = 1\nElseIf c Then\nb = 2\nElse\nb = 3\nEnd If\nEnd Sub");
        match &block.body[0] {
            Statement::If(s) => {
                assert!(!s.single_line);
                assert_eq!(s.else_ifs.len(), 1);
                assert!(s.else_body.is_some());
            }
            other => panic!("expected If, got {other:?}"),
        }

        let inline = only_proc("Sub T()\nIf a Then b = 1 Else b = 2\nEnd Sub");
        match &inline.body[0] {
            Statement::If(s) => {
                assert!(s.single_line);
                assert_eq!(s.then_body.len(), 1);
                assert!(s.else_body.is_some());
            }
            other => panic!("expected If, got {other:?}"),
        }
    }

    #[test]
    fn select_case_labels_cover_ranges_and_comparisons() {
        let p = only_proc(
            "Sub T()\nSelect Case x\nCase 1, 2\nCase 3 To 5\nCase Is >= 10\nCase Else\nEnd Select\nEnd Sub",
        );
        match &p.body[0] {
            Statement::SelectCase(s) => {
                assert_eq!(s.cases.len(), 3);
                assert_eq!(s.cases[0].labels.len(), 2);
                assert!(matches!(s.cases[1].labels[0], CaseLabel::Range(..)));
                assert!(matches!(
                    s.cases[2].labels[0],
                    CaseLabel::Compare(BinaryOp::Ge, _)
                ));
                assert!(s.case_else.is_some());
            }
            other => panic!("expected Select Case, got {other:?}"),
        }
    }

    #[test]
    fn loops_record_where_the_test_sits() {
        let p = only_proc("Sub T()\nDo Until x > 3\nx = x + 1\nLoop\nEnd Sub");
        match &p.body[0] {
            Statement::Do(d) => {
                assert!(d.pre.as_ref().is_some_and(|t| t.until));
                assert!(d.post.is_none());
            }
            other => panic!("expected Do, got {other:?}"),
        }

        let post = only_proc("Sub T()\nDo\nx = 1\nLoop While x < 3\nEnd Sub");
        match &post.body[0] {
            Statement::Do(d) => {
                assert!(d.pre.is_none());
                assert!(post_test_is_while(d));
            }
            other => panic!("expected Do, got {other:?}"),
        }
    }

    fn post_test_is_while(d: &DoStmt) -> bool {
        d.post.as_ref().is_some_and(|t| !t.until)
    }

    #[test]
    fn for_and_for_each_are_separate_shapes() {
        let p = only_proc("Sub T()\nFor i = 1 To 10 Step 2\nNext i\nEnd Sub");
        assert!(matches!(&p.body[0], Statement::For(f) if f.step.is_some()));

        let each = only_proc("Sub T()\nFor Each c In Range(\"A1:A3\")\nNext c\nEnd Sub");
        assert!(matches!(&each.body[0], Statement::ForEach(_)));
    }

    #[test]
    fn with_blocks_capture_leading_dot_members() {
        let p = only_proc("Sub T()\nWith Range(\"A1\")\n.Value = 1\nEnd With\nEnd Sub");
        match &p.body[0] {
            Statement::With { body, .. } => match &body[0] {
                Statement::Assign { target, .. } => {
                    assert!(matches!(target, Expr::WithMember(name, _) if name == "Value"))
                }
                other => panic!("expected assignment, got {other:?}"),
            },
            other => panic!("expected With, got {other:?}"),
        }
    }

    #[test]
    fn on_error_variants_are_distinct() {
        let p = only_proc(
            "Sub T()\nOn Error Resume Next\nOn Error GoTo Handler\nOn Error GoTo 0\nEnd Sub",
        );
        assert!(matches!(
            &p.body[0],
            Statement::OnError(OnError::ResumeNext { .. })
        ));
        assert!(matches!(
            &p.body[1],
            Statement::OnError(OnError::Goto { label, .. }) if label == "Handler"
        ));
        assert!(matches!(
            &p.body[2],
            Statement::OnError(OnError::Disable { .. })
        ));
    }

    #[test]
    fn set_assignment_is_not_ordinary_assignment() {
        let p = only_proc("Sub T()\nSet ws = Worksheets(1)\nx = 1\nEnd Sub");
        assert!(matches!(&p.body[0], Statement::SetAssign { .. }));
        assert!(matches!(&p.body[1], Statement::Assign { .. }));
    }

    #[test]
    fn bare_and_explicit_calls_stay_distinguishable() {
        let p = only_proc("Sub T()\nMsgBox \"hi\", vbOKOnly\nCall Foo(1)\nEnd Sub");
        match &p.body[0] {
            Statement::Call {
                target,
                explicit_call,
                ..
            } => {
                assert!(!explicit_call);
                assert_eq!(target.dotted_name().as_deref(), Some("MsgBox"));
                match target {
                    Expr::Index { args, .. } => assert_eq!(args.len(), 2),
                    other => panic!("expected call args, got {other:?}"),
                }
            }
            other => panic!("expected Call, got {other:?}"),
        }
        assert!(matches!(
            &p.body[1],
            Statement::Call { explicit_call: true, .. }
        ));
    }

    #[test]
    fn labels_and_line_numbers_are_kept() {
        let p = only_proc("Sub T()\n10 x = 1\nHandler:\nResume Next\nEnd Sub");
        assert!(matches!(&p.body[0], Statement::LineNumber { value: 10, .. }));
        assert!(matches!(&p.body[2], Statement::Label { name, .. } if name == "Handler"));
        assert!(matches!(
            &p.body[3],
            Statement::Resume {
                target: ResumeTarget::Next,
                ..
            }
        ));
    }

    #[test]
    fn module_level_declarations_are_ordered_and_typed() {
        let m = module(
            "Attribute VB_Name = \"Module1\"\nOption Explicit\nPrivate Const TAX As Double = 0.1\nPublic ws As Worksheet\n",
        );
        assert!(matches!(&m.items[0], ModuleItem::Attribute { name, .. } if name == "VB_Name"));
        assert!(matches!(
            &m.items[1],
            ModuleItem::Option(ModuleOption::Explicit, _)
        ));
        match &m.items[2] {
            ModuleItem::Variables(v) => {
                assert!(v.is_const);
                assert_eq!(v.visibility, Visibility::Private);
                assert_eq!(v.items[0].type_name.name, "Double");
            }
            other => panic!("expected a const, got {other:?}"),
        }
    }

    #[test]
    fn type_and_enum_blocks_parse() {
        let m = module("Private Type Point\n  X As Long\n  Y As Long\nEnd Type\n");
        match &m.items[0] {
            ModuleItem::Type(t) => {
                assert_eq!(t.name, "Point");
                assert_eq!(t.fields.len(), 2);
            }
            other => panic!("expected Type, got {other:?}"),
        }

        let e = module("Public Enum Colour\n  Red = 1\n  Green\nEnd Enum\n");
        match &e.items[0] {
            ModuleItem::Enum(en) => {
                assert_eq!(en.members.len(), 2);
                assert!(en.members[0].1.is_some());
                assert!(en.members[1].1.is_none());
            }
            other => panic!("expected Enum, got {other:?}"),
        }
    }

    #[test]
    fn declare_statements_are_recognised_as_external() {
        let m = module(
            "Private Declare PtrSafe Function Sleep Lib \"kernel32\" (ByVal ms As Long) As Long\n",
        );
        match &m.items[0] {
            ModuleItem::ExternalProc(d) => {
                assert!(d.ptr_safe);
                assert!(d.is_function);
                assert_eq!(d.lib, "kernel32");
                assert_eq!(d.params.len(), 1);
            }
            other => panic!("expected a Declare, got {other:?}"),
        }
    }

    #[test]
    fn redim_preserve_is_recorded() {
        let p = only_proc("Sub T()\nReDim Preserve a(1 To 10)\nEnd Sub");
        assert!(matches!(&p.body[0], Statement::ReDim { preserve: true, .. }));
    }

    #[test]
    fn japanese_identifiers_survive_parsing() {
        let p = only_proc("Sub 集計()\n税率 = 0.1\nEnd Sub");
        assert_eq!(p.name, "集計");
        match &p.body[0] {
            Statement::Assign { target, .. } => {
                assert_eq!(target.dotted_name().as_deref(), Some("税率"))
            }
            other => panic!("expected assignment, got {other:?}"),
        }
    }

    #[test]
    fn unparsed_input_is_preserved_rather_than_dropped() {
        let p = only_proc("Sub T()\nOpen \"f.txt\" For Input As #1\nx = 1\nEnd Sub");
        // Whatever the first line becomes, the next statement still parses.
        assert!(p.body.iter().any(|s| matches!(s, Statement::Assign { .. })));
    }
}
