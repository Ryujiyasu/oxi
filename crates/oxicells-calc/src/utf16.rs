// This Source Code Form is subject to the terms of the Mozilla Public
// License, v. 2.0. If a copy of the MPL was not distributed with this
// file, You can obtain one at https://mozilla.org/MPL/2.0/.

//! Worksheet text operations. VBA and Characters use different indexing rules.
use crate::{
    functions::Arg,
    value::{ExcelError, Value},
};
use std::collections::BTreeSet;

fn text(args: &[Arg], at: usize) -> Result<Vec<u16>, ExcelError> {
    args.get(at).ok_or(ExcelError::Value)?.scalar().to_utf16()
}
fn number(args: &[Arg], at: usize, default: Option<f64>) -> Result<f64, ExcelError> {
    match args.get(at) {
        Some(value) => value.scalar().to_number(),
        None => default.ok_or(ExcelError::Value),
    }
}
fn high(unit: u16) -> bool {
    (0xD800..=0xDBFF).contains(&unit)
}
fn low(unit: u16) -> bool {
    (0xDC00..=0xDFFF).contains(&unit)
}
fn spans(units: &[u16], omit_trailing_high: bool) -> Vec<(usize, usize)> {
    let mut result = Vec::new();
    let mut at = 0;
    while at < units.len() {
        if omit_trailing_high && at + 1 == units.len() && high(units[at]) {
            break;
        }
        let end = at
            + if high(units[at]) && units.get(at + 1).is_some_and(|u| low(*u)) {
                2
            } else {
                1
            };
        result.push((at, end));
        at = end;
    }
    result
}
pub(crate) fn end_piece(source: &[u16], count: f64, left: bool) -> Result<Vec<u16>, ExcelError> {
    if !count.is_finite() || count < 0.0 {
        return Err(ExcelError::Value);
    }
    let count = count as usize;
    // Excel returns the original text when the requested count reaches its
    // unit length. Otherwise LEFT's forward scan omits a trailing high unit.
    if count >= source.len() {
        return Ok(source.to_vec());
    }
    let chars = spans(source, left);
    if left && source.last().is_some_and(|u| high(*u)) && count > chars.len() {
        // When the forward scan ends on an unpaired high surrogate, Excel
        // falls back to the requested raw prefix, which can split a pair.
        return Ok(source[..count].to_vec());
    }
    let count = count.min(chars.len());
    Ok(if count == 0 {
        vec![]
    } else if left {
        source[..chars[count - 1].1].to_vec()
    } else {
        source[chars[chars.len() - count].0..].to_vec()
    })
}
fn mapped(units: &[u16], uppercase: bool) -> Vec<u16> {
    let mut out = Vec::new();
    for decoded in char::decode_utf16(units.iter().copied()) {
        match decoded {
            Ok(ch) if uppercase => {
                let mut upper = ch.to_uppercase();
                let first = upper.next().unwrap();
                let chosen = if upper.next().is_none() { first } else { ch };
                out.extend(chosen.encode_utf16(&mut [0; 2]).iter().copied());
            }
            Ok(ch) => {
                for lower in ch.to_lowercase() {
                    out.extend(lower.encode_utf16(&mut [0; 2]).iter().copied());
                }
            }
            Err(error) => out.push(error.unpaired_surrogate()),
        }
    }
    out
}

fn search_at(pattern: &[u16], haystack: &[u16], start: usize, modern: bool) -> bool {
    let end = if modern && haystack.last().is_some_and(|u| high(*u)) {
        haystack.len() - 1
    } else {
        haystack.len()
    };
    let mut work = vec![(0, start)];
    let mut seen = BTreeSet::new();
    while let Some((p, at)) = work.pop() {
        if !seen.insert((p, at)) {
            continue;
        }
        let Some(&unit) = pattern.get(p) else {
            return true;
        };
        let next = || {
            if at >= end {
                None
            } else {
                Some(
                    at + if modern
                        && high(haystack[at])
                        && haystack.get(at + 1).is_some_and(|u| low(*u))
                    {
                        2
                    } else {
                        1
                    },
                )
            }
        };
        match unit {
            42 => {
                work.push((p + 1, at));
                if let Some(next) = next() {
                    work.push((p, next));
                }
            }
            63 => {
                if let Some(next) = next() {
                    work.push((p + 1, next));
                }
            }
            126 => {
                if let Some(escaped) = pattern.get(p + 1) {
                    if at < end && haystack.get(at) == Some(escaped) {
                        work.push((p + 2, at + 1));
                    }
                } else {
                    // Excel ignores a trailing escape, even after a literal.
                    work.push((p + 1, at));
                }
            }
            _ => {
                if at < end && haystack.get(at) == Some(&unit) {
                    work.push((p + 1, at + 1));
                }
            }
        }
    }
    false
}

pub(crate) fn call(name: &str, args: &[Arg], version: u8) -> Option<Result<Value, ExcelError>> {
    let (least, most) = match name {
        "LEN" | "UPPER" | "LOWER" | "TRIM" | "T" => (1, 1),
        "LEFT" | "RIGHT" => (1, 2),
        "MID" => (3, 3),
        "REPLACE" => (4, 4),
        "EXACT" => (2, 2),
        "SUBSTITUTE" => (3, 4),
        "FIND" | "SEARCH" => (2, 3),
        _ => return None,
    };
    Some((|| {
        if args.len() < least || args.len() > most {
            return Err(ExcelError::Value);
        }
        if name == "T" {
            return match args[0].scalar() {
                value @ (Value::Text(_) | Value::Utf16(_)) => Ok(value),
                Value::Error(error) => Err(error),
                _ => Ok(Value::Text(String::new())),
            };
        }
        let source = text(args, 0)?;
        let modern = version >= 2;
        match name {
            "LEN" => Ok(Value::Number(if modern {
                spans(&source, true).len()
            } else {
                source.len()
            } as f64)),
            "EXACT" => Ok(Value::Logical(source == text(args, 1)?)),
            "UPPER" | "LOWER" => Ok(Value::from_utf16(mapped(&source, name == "UPPER"))),
            "TRIM" => {
                let mut out = Vec::new();
                let mut space = false;
                for unit in source {
                    if unit == 32 {
                        space = !out.is_empty();
                    } else {
                        if space {
                            out.push(32);
                        }
                        out.push(unit);
                        space = false;
                    }
                }
                Ok(Value::from_utf16(out))
            }
            "LEFT" | "RIGHT" => {
                let count = number(args, 1, Some(1.0))?;
                Ok(Value::from_utf16(end_piece(
                    &source,
                    count,
                    name == "LEFT",
                )?))
            }
            "MID" | "REPLACE" => {
                let start = number(args, 1, None)?;
                let count = number(args, 2, None)?;
                let limit = if name == "REPLACE" {
                    32_767.0
                } else {
                    i32::MAX as f64
                };
                if !start.is_finite()
                    || !count.is_finite()
                    || start < 1.0
                    || count < 0.0
                    || start.trunc() > limit
                    || count.trunc() > limit
                {
                    return Err(ExcelError::Value);
                }
                let index = (start as usize).saturating_sub(1);
                let (from, to) = if modern {
                    let chars = spans(&source, name == "MID");
                    let from = chars
                        .get(index)
                        .map(|span| span.0)
                        .unwrap_or_else(|| chars.last().map_or(0, |span| span.1));
                    let end = index.saturating_add(count as usize).min(chars.len());
                    let to = if end > index { chars[end - 1].1 } else { from };
                    (from, to)
                } else {
                    let from = index.min(source.len());
                    (from, from.saturating_add(count as usize).min(source.len()))
                };
                if name == "MID" {
                    Ok(Value::from_utf16(source[from..to].to_vec()))
                } else {
                    let mut out = source[..from].to_vec();
                    out.extend(text(args, 3)?);
                    out.extend_from_slice(&source[to..]);
                    if out.len() > 32_767 {
                        return Err(ExcelError::Value);
                    }
                    Ok(Value::from_utf16(out))
                }
            }
            "SUBSTITUTE" => {
                let old = text(args, 1)?;
                let new = text(args, 2)?;
                let which = number(args, 3, Some(f64::INFINITY))?;
                if which < 1.0 {
                    return Err(ExcelError::Value);
                }
                if old.is_empty() {
                    return Ok(Value::from_utf16(source));
                }
                let mut out = Vec::new();
                let mut at = 0;
                let mut seen = 0;
                while at < source.len() {
                    if source[at..].starts_with(&old) {
                        seen += 1;
                        if which.is_infinite() || seen == which as usize {
                            out.extend_from_slice(&new);
                        } else {
                            out.extend_from_slice(&old);
                        }
                        at += old.len();
                    } else {
                        out.push(source[at]);
                        at += 1;
                    }
                }
                Ok(Value::from_utf16(out))
            }
            "FIND" | "SEARCH" => {
                let mut needle = source;
                let mut haystack = text(args, 1)?;
                let start = number(args, 2, Some(1.0))?;
                if !start.is_finite() || start < 1.0 {
                    return Err(ExcelError::Value);
                }
                let start = (start as usize).saturating_sub(1);
                if name == "FIND" && modern && needle.last().is_some_and(|u| high(*u)) {
                    needle.pop();
                }
                if name == "SEARCH" {
                    needle = mapped(&needle, false);
                    haystack = mapped(&haystack, false);
                }
                let mut positions: Vec<usize> = if modern {
                    spans(&haystack, true).iter().map(|span| span.0).collect()
                } else {
                    (0..haystack.len()).collect()
                };
                positions.push(if modern {
                    spans(&haystack, true).last().map_or(0, |span| span.1)
                } else {
                    haystack.len()
                });
                // SEARCH requires a character at the starting position;
                // FIND also accepts the end position for an empty needle.
                if name == "SEARCH" && start >= positions.len().saturating_sub(1) {
                    return Err(ExcelError::Value);
                }
                for (index, at) in positions.into_iter().enumerate().skip(start) {
                    let matched = if name == "SEARCH" {
                        search_at(&needle, &haystack, at, modern)
                    } else {
                        haystack[at..].starts_with(&needle)
                    };
                    if matched {
                        return Ok(Value::Number((index + 1) as f64));
                    }
                }
                Err(ExcelError::Value)
            }
            _ => unreachable!(),
        }
    })())
}
