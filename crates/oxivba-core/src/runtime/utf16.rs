// This Source Code Form is subject to the terms of the Mozilla Public
// License, v. 2.0. If a copy of the MPL was not distributed with this
// file, You can obtain one at https://mozilla.org/MPL/2.0/.

//! String operations on UTF-16 units, including unpaired surrogates.
use super::*;

pub(super) fn units(value: &Value, line: Option<u32>) -> Result<Vec<u16>, RuntimeError> {
    match value {
        Value::Null => return Err(invalid_null(line)),
        Value::Error(_) => {
            return Err(error(
                RuntimeErrorKind::TypeMismatch,
                "Error is not string data",
                line,
            ))
        }
        Value::Nothing => {
            return Err(error(
                RuntimeErrorKind::ObjectVariableNotSet,
                "Nothing is not string data",
                line,
            ))
        }
        _ => {}
    }
    if let Some(s) = value.string_units() {
        return Ok(s.into_owned());
    }
    text(value)
        .map(|s| s.encode_utf16().collect())
        .map_err(|m| error(RuntimeErrorKind::TypeMismatch, m, line))
}
pub(super) fn has_null_string(value: &Value) -> bool {
    match value {
        Value::NullString => true,
        Value::Array(a) => {
            a.values.iter().any(has_null_string) || has_null_string(&a.element_default)
        }
        Value::Record(r) => r.fields.iter().any(|(_, v)| has_null_string(v)),
        _ => false,
    }
}
pub(super) fn host_value(value: &Value) -> Value {
    match value {
        Value::NullString => Value::String(String::new()),
        Value::Array(a) => {
            let mut a = a.clone();
            a.values = a.values.iter().map(host_value).collect();
            a.element_default = Box::new(host_value(&a.element_default));
            Value::Array(a)
        }
        Value::Record(r) => {
            let mut r = r.clone();
            r.fields = r
                .fields
                .into_iter()
                .map(|(n, v)| (n, host_value(&v)))
                .collect();
            Value::Record(r)
        }
        _ => value.clone(),
    }
}
pub(super) fn fold(units: &[u16], upper: bool) -> Vec<u16> {
    let mut out = Vec::new();
    for c in char::decode_utf16(units.iter().copied()) {
        match c {
            Ok(c) => {
                let s = c.to_string();
                let mapped = if upper {
                    crate::case_table::upper(&s)
                } else {
                    crate::case_table::lower(&s)
                };
                out.extend(mapped.encode_utf16());
            }
            Err(e) => out.push(e.unpaired_surrogate()),
        }
    }
    out
}
pub(super) fn collation(units: &[u16]) -> Vec<u32> {
    let mut out = Vec::new();
    let mut run = String::new();
    for c in char::decode_utf16(units.iter().copied()) {
        match c {
            Ok(c) => run.push(c),
            Err(e) => {
                out.extend(collation_spelling(&run, false).chars().map(u32::from));
                run.clear();
                out.push(0x110000 + u32::from(e.unpaired_surrogate()));
            }
        }
    }
    out.extend(collation_spelling(&run, false).chars().map(u32::from));
    out
}
pub(super) fn find_span(source: &[u16], needle: &[u16], start: usize) -> Option<(usize, usize)> {
    if start > source.len() {
        return None;
    }
    if needle.is_empty() {
        return Some((start, start));
    }
    let wanted = collation(needle);
    for at in start..source.len() {
        for end in at + 1..=source.len() {
            let spelled = collation(&source[at..end]);
            if spelled == wanted {
                return Some((at, end));
            }
            let pending_pair = (0xd800..=0xdbff).contains(&source[end - 1])
                && end < source.len()
                && (0xdc00..=0xdfff).contains(&source[end]);
            if !pending_pair && !wanted.starts_with(&spelled) {
                break;
            }
        }
    }
    None
}
fn wide(value: &Value) -> bool {
    match value {
        Value::Utf16(_) => true,
        Value::Array(a) => a.values.iter().any(wide),
        _ => false,
    }
}
pub(super) fn builtin(
    name: &str,
    args: &[Value],
    line: Option<u32>,
    compare_text: bool,
) -> Option<Result<Value, RuntimeError>> {
    let name = name.trim_end_matches('$');
    let reverse_pair = name == "strreverse"
        && args.iter().any(|v| {
            v.string_units()
                .is_some_and(|u| u.iter().any(|c| (0xd800..=0xdfff).contains(c)))
        });
    if !args.iter().any(wide) && !reverse_pair {
        return None;
    }
    if !matches!(
        name,
        "cstr"
            | "cvar"
            | "len"
            | "lenb"
            | "ascw"
            | "asc"
            | "left"
            | "right"
            | "mid"
            | "strreverse"
            | "trim"
            | "ltrim"
            | "rtrim"
            | "ucase"
            | "lcase"
            | "strcomp"
            | "instr"
            | "instrrev"
            | "replace"
            | "split"
            | "join"
            | "string"
            | "filter"
    ) {
        return None;
    }
    let count = |min, max| {
        if args.len() < min || args.len() > max {
            Err(error(
                RuntimeErrorKind::ArgumentCount,
                "incorrect string argument count",
                line,
            ))
        } else {
            Ok(())
        }
    };
    let integer = |i: usize| integer_argument(&args[i], line);
    let fail = |m: &str| invalid_procedure_call(m.to_string(), line);
    let mode = |i: usize| compare_mode(args.get(i), compare_text, line, true);
    Some((|| match name {
        "cstr" | "cvar" => {
            count(1, 1)?;
            if args[0].is_string() {
                Ok(args[0].clone())
            } else {
                Err(error(RuntimeErrorKind::TypeMismatch, "not a string", line))
            }
        }
        "len" | "lenb" | "ascw" | "asc" => {
            count(1, 1)?;
            let u = units(&args[0], line)?;
            if name == "len" || name == "lenb" {
                return Ok(Value::Integer(
                    u.len() as i64 * if name == "lenb" { 2 } else { 1 },
                ));
            }
            let first = *u.first().ok_or_else(|| fail("empty string"))?;
            Ok(Value::Int16(if name == "ascw" {
                first as i16
            } else if first < 256 {
                first as i16
            } else {
                63
            }))
        }
        "left" | "right" | "mid" => {
            count(2, if name == "mid" { 3 } else { 2 })?;
            if matches!(args[0], Value::Null) {
                return Ok(Value::Null);
            }
            let u = units(&args[0], line)?;
            let n = integer(1)?;
            if n < if name == "mid" { 1 } else { 0 } {
                return Err(fail("negative string position or length"));
            }
            let (start, length) = if name == "mid" {
                let start = usize::try_from(n - 1).unwrap_or(usize::MAX).min(u.len());
                let length = if args.len() == 3 {
                    let n = integer(2)?;
                    if n < 0 {
                        return Err(fail("negative string length"));
                    }
                    usize::try_from(n).unwrap_or(usize::MAX)
                } else {
                    u.len()
                };
                (start, length.min(u.len() - start))
            } else {
                let n = usize::try_from(n).unwrap_or(usize::MAX).min(u.len());
                (if name == "left" { 0 } else { u.len() - n }, n)
            };
            Ok(Value::from_utf16(u[start..start + length].to_vec()))
        }
        "strreverse" | "trim" | "ltrim" | "rtrim" | "ucase" | "lcase" => {
            count(1, 1)?;
            let mut u = units(&args[0], line)?;
            match name {
                "strreverse" => u.reverse(),
                "ucase" => u = fold(&u, true),
                "lcase" => u = fold(&u, false),
                _ => {
                    let first = if name != "rtrim" {
                        u.iter().position(|c| *c != 32).unwrap_or(u.len())
                    } else {
                        0
                    };
                    let end = if name != "ltrim" {
                        u.iter().rposition(|c| *c != 32).map_or(0, |i| i + 1)
                    } else {
                        u.len()
                    };
                    u = u[first..end.max(first)].to_vec();
                }
            }
            Ok(Value::from_utf16(u))
        }
        "strcomp" => {
            count(2, 3)?;
            if matches!(args[0], Value::Null) || matches!(args[1], Value::Null) {
                return Ok(Value::Null);
            }
            let a = units(&args[0], line)?;
            let b = units(&args[1], line)?;
            let order = if mode(2)? {
                collation(&a).cmp(&collation(&b))
            } else {
                a.cmp(&b)
            };
            Ok(Value::Int16(match order {
                std::cmp::Ordering::Less => -1,
                std::cmp::Ordering::Equal => 0,
                std::cmp::Ordering::Greater => 1,
            }))
        }
        "instr" | "instrrev" => {
            count(2, 4)?;
            let reverse = name == "instrrev";
            let (a, b) = if !reverse && args.len() > 2 {
                (1, 2)
            } else {
                (0, 1)
            };
            if matches!(args[a], Value::Null) || matches!(args[b], Value::Null) {
                return Ok(Value::Null);
            }
            let source = units(&args[a], line)?;
            let needle = units(&args[b], line)?;
            let text_mode = if args.len() == 4 {
                mode(3)?
            } else {
                compare_text
            };
            let start = if reverse {
                if args.len() > 2 {
                    integer(2)?
                } else {
                    -1
                }
            } else if args.len() > 2 {
                integer(0)?
            } else {
                1
            };
            if source.is_empty() {
                return Ok(Value::Integer(0));
            }
            if !reverse && needle.is_empty() {
                return Ok(Value::Integer(start));
            }
            if start < 1 && !(reverse && start == -1) {
                return Err(fail("invalid string position"));
            }
            let found = if reverse {
                utf16_rfind(
                    &source,
                    &needle,
                    if start == -1 {
                        source.len()
                    } else {
                        start as usize
                    },
                    text_mode,
                )
            } else {
                utf16_find(&source, &needle, (start - 1) as usize, text_mode)
            };
            Ok(Value::Integer(found.map_or(0, |i| i as i64 + 1)))
        }
        "replace" => {
            count(3, 6)?;
            let source = units(&args[0], line)?;
            let needle = units(&args[1], line)?;
            let replacement = units(&args[2], line)?;
            let start = if args.len() > 3 { integer(3)? } else { 1 };
            let limit = if args.len() > 4 { integer(4)? } else { -1 };
            if start < 1 || limit < -1 {
                return Err(fail("invalid replace position/count"));
            }
            let start = (start as usize - 1).min(source.len());
            let mut out = Vec::new();
            let mut cursor = start;
            let mut changed = 0_i64;
            if !needle.is_empty() {
                while limit < 0 || changed < limit {
                    let Some((at, end)) = utf16_find_span(&source, &needle, cursor, mode(5)?)
                    else {
                        break;
                    };
                    out.extend_from_slice(&source[cursor..at]);
                    out.extend_from_slice(&replacement);
                    cursor = end;
                    changed += 1;
                }
            }
            out.extend_from_slice(&source[cursor..]);
            Ok(Value::from_utf16(out))
        }
        "split" => {
            count(1, 4)?;
            let source = units(&args[0], line)?;
            let delimiter = if args.len() > 1 {
                units(&args[1], line)?
            } else {
                vec![32]
            };
            let limit = if args.len() > 2 { integer(2)? } else { -1 };
            if limit < -1 {
                return Err(fail("invalid split limit"));
            }
            let mut values = Vec::new();
            let mut cursor = 0;
            if limit != 0 && !source.is_empty() {
                if !delimiter.is_empty() {
                    while limit < 0 || values.len() + 1 < limit as usize {
                        let Some((at, end)) =
                            utf16_find_span(&source, &delimiter, cursor, mode(3)?)
                        else {
                            break;
                        };
                        values.push(Value::from_utf16(source[cursor..at].to_vec()));
                        cursor = end;
                    }
                }
                values.push(Value::from_utf16(source[cursor..].to_vec()));
            }
            let length = values.len();
            Ok(Value::Array(ArrayValue {
                dimensions: vec![ArrayDimension {
                    lower_bound: 0,
                    length,
                }],
                values,
                element_default: Box::new(Value::NullString),
                resizable: true,
            }))
        }
        "join" => {
            count(1, 2)?;
            let Value::Array(a) = &args[0] else {
                return Err(error(
                    RuntimeErrorKind::TypeMismatch,
                    "Join requires an array",
                    line,
                ));
            };
            if a.dimensions.len() != 1 {
                return Err(error(
                    RuntimeErrorKind::TypeMismatch,
                    "Join requires one dimension",
                    line,
                ));
            }
            let delimiter = if args.len() > 1 {
                units(&args[1], line)?
            } else {
                vec![32]
            };
            let mut out = Vec::new();
            for (i, v) in a.values.iter().enumerate() {
                if i > 0 {
                    out.extend_from_slice(&delimiter);
                }
                out.extend(units(v, line)?);
            }
            Ok(Value::from_utf16(out))
        }
        "filter" => {
            count(2, 4)?;
            let Value::Array(a) = &args[0] else {
                return Err(error(
                    RuntimeErrorKind::TypeMismatch,
                    "Filter requires an array",
                    line,
                ));
            };
            if a.dimensions.len() != 1 {
                return Err(error(
                    RuntimeErrorKind::TypeMismatch,
                    "Filter requires one dimension",
                    line,
                ));
            }
            let needle = units(&args[1], line)?;
            let include = match args.get(2) {
                None | Some(Value::Missing) => true,
                Some(v) => truthy(v).map_err(|m| error(RuntimeErrorKind::TypeMismatch, m, line))?,
            };
            let mut values = Vec::new();
            for v in &a.values {
                let u = units(v, line)?;
                if utf16_find(&u, &needle, 0, mode(3)?).is_some() == include {
                    values.push(Value::from_utf16(u));
                }
            }
            let length = values.len();
            Ok(Value::Array(ArrayValue {
                dimensions: vec![ArrayDimension {
                    lower_bound: 0,
                    length,
                }],
                values,
                element_default: Box::new(Value::NullString),
                resizable: true,
            }))
        }
        "string" => {
            count(2, 2)?;
            let count = integer(0)?;
            if count < 0 {
                return Err(fail("negative string length"));
            }
            if count > 1_000_000 {
                return Err(error(
                    RuntimeErrorKind::Overflow,
                    "string length exceeds runtime limit",
                    line,
                ));
            }
            let u = units(&args[1], line)?;
            let c = *u.first().ok_or_else(|| fail("empty character"))?;
            Ok(Value::from_utf16(vec![c; count as usize]))
        }
        _ => unreachable!("filtered string operation"),
    })())
}
