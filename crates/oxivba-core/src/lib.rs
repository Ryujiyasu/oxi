// This Source Code Form is subject to the terms of the Mozilla Public
// License, v. 2.0. If a copy of the MPL was not distributed with this
// file, You can obtain one at https://mozilla.org/MPL/2.0/.

//! VBA source analysis.
//!
//! The first goal is to answer, for a body of macros nobody has read in years,
//! three questions:
//!
//! 1. Which of these files are copies or variants of one another?
//! 2. What does each one actually use?
//! 3. Which of them can be migrated, and which cannot?
//!
//! All three are answerable from source text alone — no Excel, no COM, no
//! Windows. That is the whole reason this crate takes VBA *text* as its input
//! rather than an `.xlsm` file: extraction has several existing solutions,
//! while the analysis layer has none. The same syntax tree now also feeds a
//! small host-independent interpreter for pure VBA code. Excel's object model
//! will be supplied separately by a browser host adapter.
//!
//! # Status
//!
//! Lexer and parser. See `FEATURE-MAP.md` for the language surface this is
//! working towards and for where the line between "in scope" and "out of scope"
//! is currently drawn.
//!
//! Nothing is ever silently dropped: input the parser does not understand is
//! preserved verbatim as `Unknown`, so unparsed input can be counted.

pub mod analysis;
pub mod ast;
pub mod fingerprint;
pub mod lexer;
pub mod parser;
pub mod runtime;

pub use analysis::{analyse, Analysis, Class, Finding};
pub use ast::{Expr, Module, ModuleItem, Procedure, Statement};
pub use fingerprint::{compare, fingerprint_module, ModuleFingerprint, Similarity, Strength};
pub use lexer::{tokenize, LexError, Punct, Span, Token, TokenKind};
pub use parser::parse_module;
pub use runtime::{execute, Runtime, RuntimeError, RuntimeErrorKind, Value};
