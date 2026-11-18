// This Source Code Form is subject to the terms of the Mozilla Public
// License, v. 2.0. If a copy of the MPL was not distributed with this
// file, You can obtain one at https://mozilla.org/MPL/2.0/.

use super::*;
use crate::extended::{Extended, Fault};

impl Runtime<'_> {
    /// Library-qualified calls receive stored numeric arguments, including
    /// VBA.CInt. The unqualified compiler conversion consumes the intermediate.
    pub(super) fn qualified_stored_numeric_call(
        &mut self,
        bare: &Expr,
        frame: &mut Frame,
    ) -> Option<Result<Value, RuntimeError>> {
        let Expr::Index {
            target, args, span, ..
        } = bare
        else {
            return None;
        };
        let name = expr_name(target)?;
        if !matches!(
            key(name).as_str(),
            "abs" | "sgn" | "int" | "fix" | "cint" | "cdbl" | "csng" | "cbool" | "clng" | "cbyte"
        ) {
            return None;
        }
        let [Argument {
            value: Some(operand),
            name: None,
            ..
        }] = args.as_slice()
        else {
            return None;
        };
        Some((|| {
            let value = self.eval_expr(operand, frame)?;
            self.call_named(name, vec![value], &[None], Some(span.line), frame)
        })())
    }
    pub(super) fn floating_condition_value(
        &mut self,
        expr: &Expr,
        frame: &mut Frame,
    ) -> Result<Value, RuntimeError> {
        match self.floating_conversion(expr, "Boolean", frame) {
            Some(result) => result,
            None => self.eval_scalar_expr(expr, frame),
        }
    }

    pub(super) fn floating_integer_operand(
        &mut self,
        expr: &Expr,
        frame: &mut Frame,
    ) -> Result<Value, RuntimeError> {
        match self.floating_conversion(expr, "Long", frame) {
            Some(result) => result,
            None => self.eval_expr(expr, frame),
        }
    }
    /// Determine eligibility without executing a getter or a function twice.
    fn intermediate_rank(&self, expr: &Expr, frame: &Frame) -> Option<NumRank> {
        if let Some(bare) = self.vba_qualified(expr, frame) {
            return self.intermediate_rank(&bare, frame);
        }
        let rank = match expr {
            Expr::Group { expression, .. } => return self.intermediate_rank(expression, frame),
            Expr::Literal(value, _) => NumRank::of(&literal_value(value))?,
            Expr::Unary {
                op: UnaryOp::Plus | UnaryOp::Neg,
                operand,
                ..
            } => return self.intermediate_rank(operand, frame),
            Expr::Binary { op, lhs, rhs, .. }
                if matches!(
                    op,
                    BinaryOp::Add | BinaryOp::Sub | BinaryOp::Mul | BinaryOp::Div
                ) =>
            {
                let a = self.intermediate_rank(lhs, frame)?;
                let b = self.intermediate_rank(rhs, frame)?;
                let single = [a, b].contains(&NumRank::Single);
                let narrow = |r| matches!(r, NumRank::Byte | NumRank::Int16 | NumRank::Single);
                if *op == BinaryOp::Div {
                    if single && narrow(a) && narrow(b) {
                        NumRank::Single
                    } else {
                        NumRank::Double
                    }
                } else if single && matches!(a, NumRank::Long | NumRank::LongLong)
                    || single && matches!(b, NumRank::Long | NumRank::LongLong)
                {
                    NumRank::Double
                } else {
                    a.max(b)
                }
            }
            _ => {
                let declared = self
                    .source_expression_type(expr, frame)
                    .map(|t| t.name)
                    .or_else(|| {
                        let Expr::Index { target, args, .. } = expr else {
                            return None;
                        };
                        let name = expr_name(target)?;
                        if self.lookup_slot(frame, name).is_some()
                            || self.has_visible_procedure(name)
                        {
                            return None;
                        }
                        match key(name).as_str() {
                            "cdbl" => Some("Double".into()),
                            "csng" => Some("Single".into()),
                            "cint" => Some("Integer".into()),
                            "clng" => Some("Long".into()),
                            "cbyte" => Some("Byte".into()),
                            "sqr" | "val" => Some("Double".into()),
                            "abs" | "int" | "fix" => {
                                let operand = args.first()?.value.as_ref()?;
                                match self.intermediate_rank(operand, frame)? {
                                    NumRank::Single => Some("Single".into()),
                                    NumRank::Double => Some("Double".into()),
                                    _ => None,
                                }
                            }
                            _ => None,
                        }
                    })?;
                match key(&declared).as_str() {
                    "byte" => NumRank::Byte,
                    "integer" | "boolean" => NumRank::Int16,
                    "long" => NumRank::Long,
                    "longlong" => NumRank::LongLong,
                    "single" => NumRank::Single,
                    "double" => NumRank::Double,
                    _ => return None,
                }
            }
        };
        (rank != NumRank::Currency).then_some(rank)
    }

    pub(super) fn floating_chain(&self, expr: &Expr, frame: &Frame) -> Option<NumRank> {
        if !matches!(
            expr,
            Expr::Binary {
                op: BinaryOp::Add | BinaryOp::Sub | BinaryOp::Mul | BinaryOp::Div,
                ..
            } | Expr::Unary {
                op: UnaryOp::Plus | UnaryOp::Neg,
                ..
            }
        ) {
            return None;
        }
        self.intermediate_rank(expr, frame)
            .filter(|r| matches!(r, NumRank::Single | NumRank::Double))
    }

    fn floating_fault(fault: Fault, line: u32) -> RuntimeError {
        match fault {
            Fault::Overflow => error(RuntimeErrorKind::Overflow, "overflow", Some(line)),
            Fault::DivisionByZero => error(
                RuntimeErrorKind::DivisionByZero,
                "division by zero",
                Some(line),
            ),
        }
    }

    fn floating_intermediate(
        &mut self,
        expr: &Expr,
        frame: &mut Frame,
    ) -> Result<Extended, RuntimeError> {
        let line = expr.span().line;
        if self.prepared_source_argument_value(expr, frame).is_none() {
            if let Some(bare) = self.vba_qualified(expr, frame) {
                let mut aliases = Vec::new();
                self.alias_prepared_source_tree(expr, &bare, frame, &mut aliases);
                let result = match self.qualified_stored_numeric_call(&bare, frame) {
                    Some(result) => result.and_then(|value| {
                        any_number(&value)
                            .and_then(Extended::from_f64)
                            .ok_or_else(|| {
                                error(RuntimeErrorKind::TypeMismatch, "type mismatch", Some(line))
                            })
                    }),
                    None => self.floating_intermediate(&bare, frame),
                };
                for (scope, id, value, reference, grouped) in aliases.into_iter().rev() {
                    if value {
                        self.source_argument_preparations[scope].values.remove(&id);
                    }
                    if reference {
                        self.source_argument_preparations[scope]
                            .references
                            .remove(&id);
                    }
                    if grouped {
                        self.source_argument_preparations[scope]
                            .grouped_values
                            .remove(&id);
                    }
                }
                return result;
            }
            if let Expr::Group { expression, .. } = expr {
                return self.floating_intermediate(expression, frame);
            }
            if let Expr::Index { target, args, .. } = expr {
                if let Some(name) = expr_name(target).filter(|name| {
                    name.eq_ignore_ascii_case("abs")
                        && self.lookup_slot(frame, name).is_none()
                        && !self.has_visible_procedure(name)
                }) {
                    let _ = name;
                    if let [Argument {
                        value: Some(operand),
                        name: None,
                        ..
                    }] = args.as_slice()
                    {
                        if self
                            .intermediate_rank(operand, frame)
                            .is_some_and(|r| matches!(r, NumRank::Single | NumRank::Double))
                        {
                            return self
                                .floating_intermediate(operand, frame)
                                .map(Extended::abs);
                        }
                    }
                }
            }
            if self.floating_chain(expr, frame).is_some() {
                return match expr {
                    Expr::Binary { op, lhs, rhs, .. } => {
                        let a = self.floating_intermediate(lhs, frame)?;
                        let b = self.floating_intermediate(rhs, frame)?;
                        match op {
                            BinaryOp::Add => a.add(b),
                            BinaryOp::Sub => a.sub(b),
                            BinaryOp::Mul => a.mul(b),
                            BinaryOp::Div if a.is_zero() && b.is_zero() => Err(Fault::Overflow),
                            BinaryOp::Div => a.div(b),
                            _ => unreachable!(),
                        }
                        .map_err(|fault| Self::floating_fault(fault, line))
                    }
                    Expr::Unary { op, operand, .. } => {
                        let value = self.floating_intermediate(operand, frame)?;
                        Ok(if *op == UnaryOp::Neg {
                            value.negated()
                        } else {
                            value
                        })
                    }
                    _ => unreachable!(),
                };
            }
        }
        let value = self.eval_scalar_expr(expr, frame)?;
        let wide = match value {
            Value::LongLong(n) | Value::Integer(n) => Some(Extended::from_i64(n)),
            Value::Boolean(value) => Some(Extended::from_i64(if value { -1 } else { 0 })),
            other => any_number(&other).and_then(Extended::from_f64),
        };
        wide.ok_or_else(|| error(RuntimeErrorKind::TypeMismatch, "type mismatch", Some(line)))
    }

    pub(super) fn floating_conversion(
        &mut self,
        expr: &Expr,
        kind: &str,
        frame: &mut Frame,
    ) -> Option<Result<Value, RuntimeError>> {
        if !self
            .intermediate_rank(expr, frame)
            .is_some_and(|r| matches!(r, NumRank::Single | NumRank::Double))
        {
            return None;
        }
        let target = key(kind);
        if !matches!(
            target.as_str(),
            "cdbl"
                | "double"
                | "csng"
                | "single"
                | "cbool"
                | "boolean"
                | "cint"
                | "integer"
                | "clng"
                | "long"
                | "cbyte"
                | "byte"
                | "sgn"
                | "int"
        ) {
            return None;
        }
        Some((|| {
            let wide = self.floating_intermediate(expr, frame)?;
            let line = expr.span().line;
            match target.as_str() {
                "cbool" | "boolean" => Ok(Value::Boolean(!wide.is_zero())),
                "sgn" => Ok(Value::Int16(if wide.is_zero() {
                    0
                } else if wide.is_negative() {
                    -1
                } else {
                    1
                })),
                "cdbl" | "double" | "csng" | "single" | "int" => {
                    let wide = if target == "int" { wide.floor() } else { wide };
                    let single = target == "csng"
                        || target == "single"
                        || target == "int"
                            && self.intermediate_rank(expr, frame) == Some(NumRank::Single);
                    let value = if single {
                        Value::Single(wide.to_f32())
                    } else {
                        Value::Double(wide.to_f64())
                    };
                    if any_number(&value).is_some_and(f64::is_finite) {
                        Ok(value)
                    } else {
                        Err(Self::floating_fault(Fault::Overflow, line))
                    }
                }
                _ => {
                    let value = wide
                        .round_i64()
                        .ok_or_else(|| Self::floating_fault(Fault::Overflow, line))?;
                    let name = match target.as_str() {
                        "cint" | "integer" => "Integer",
                        "clng" | "long" => "Long",
                        _ => "Byte",
                    };
                    self.coerce_runtime(Value::LongLong(value), name, line, self.this_year())
                }
            }
        })())
    }

    pub(super) fn boxed_floating_chain(
        &mut self,
        expr: &Expr,
        frame: &mut Frame,
        rank: NumRank,
    ) -> Result<Value, RuntimeError> {
        let wide = self.floating_intermediate(expr, frame)?;
        let answer = if rank == NumRank::Single {
            Value::Single(wide.to_f32())
        } else {
            Value::Double(wide.to_f64())
        };
        if any_number(&answer).is_some_and(f64::is_finite) {
            Ok(answer)
        } else {
            Err(Self::floating_fault(Fault::Overflow, expr.span().line))
        }
    }
}
