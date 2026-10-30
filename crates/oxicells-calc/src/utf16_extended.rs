// This Source Code Form is subject to the terms of the Mozilla Public
// License, v. 2.0. If a copy of the MPL was not distributed with this
// file, You can obtain one at https://mozilla.org/MPL/2.0/.

use crate::{
    functions::Arg,
    value::{ExcelError, Value},
};
const LIMIT: usize = 32_767;
fn units(args: &[Arg], at: usize) -> Result<Vec<u16>, ExcelError> {
    args.get(at).ok_or(ExcelError::Value)?.scalar().to_utf16()
}
fn append(out: &mut Vec<u16>, piece: &[u16]) -> Result<(), ExcelError> {
    if out.len().saturating_add(piece.len()) > LIMIT {
        return Err(ExcelError::Value);
    }
    out.extend_from_slice(piece);
    Ok(())
}
fn written(value: &Value, strict: bool) -> Vec<u16> {
    match value {
        Value::Error(error) => error.as_str().encode_utf16().collect(),
        Value::Text(_) | Value::Utf16(_) if strict => {
            let mut out = vec![34];
            for unit in value.to_utf16().unwrap() {
                out.push(unit);
                if unit == 34 {
                    out.push(34);
                }
            }
            out.push(34);
            out
        }
        _ => value.to_utf16().unwrap_or_default(),
    }
}
fn fold(units: &[u16]) -> Vec<u16> {
    let mut out = Vec::new();
    let mut text = String::new();
    let flush = |text: &mut String, out: &mut Vec<u16>| {
        out.extend(crate::case_fold::fold_text(text).encode_utf16());
        text.clear();
    };
    for ch in char::decode_utf16(units.iter().copied()) {
        match ch {
            Ok(ch) => text.push(ch),
            Err(error) => {
                flush(&mut text, &mut out);
                out.push(error.unpaired_surrogate());
            }
        }
    }
    flush(&mut text, &mut out);
    out
}

pub(crate) fn call(
    name: &str,
    args: &[Arg],
    dbcs_language: bool,
) -> Option<Result<Value, ExcelError>> {
    if !matches!(
        name,
        "CONCAT"
            | "CONCATENATE"
            | "TEXTJOIN"
            | "REPT"
            | "CLEAN"
            | "PROPER"
            | "UNICODE"
            | "VALUETOTEXT"
            | "ARRAYTOTEXT"
            | "TEXTBEFORE"
            | "TEXTAFTER"
            | "TEXT"
            | "LENB"
            | "LEFTB"
            | "RIGHTB"
            | "MIDB"
            | "FINDB"
            | "SEARCHB"
            | "REPLACEB"
    ) {
        return None;
    }
    if name == "TEXT" && !matches!(args.first().map(Arg::scalar), Some(Value::Utf16(_))) {
        return None;
    }
    if dbcs_language
        && matches!(
            name,
            "LENB" | "LEFTB" | "RIGHTB" | "MIDB" | "FINDB" | "SEARCHB" | "REPLACEB"
        )
    {
        // Keep the established DBCS implementation until its invalid UTF-16
        // behavior has been measured with a DBCS calculation language.
        return None;
    }
    // In an SBCS calculation language the byte functions retain the legacy
    // character-function indexing, including in compatibility version 2.
    let legacy = match name {
        "LEFTB" => Some("LEFT"),
        "RIGHTB" => Some("RIGHT"),
        "MIDB" => Some("MID"),
        "FINDB" => Some("FIND"),
        "SEARCHB" => Some("SEARCH"),
        "REPLACEB" => Some("REPLACE"),
        _ => None,
    };
    if let Some(legacy) = legacy {
        return crate::utf16::call(legacy, args, 1);
    }
    Some((|| match name {
        "UNICODE" => {
            if args.len() != 1 {
                return Err(ExcelError::Value);
            }
            let source = units(args, 0)?;
            let first = *source.first().ok_or(ExcelError::Value)?;
            let code = if (0xD800..=0xDBFF).contains(&first) {
                let low = *source
                    .get(1)
                    .filter(|low| (0xDC00..=0xDFFF).contains(*low))
                    .ok_or(ExcelError::Value)?;
                0x10000 + ((first as u32 - 0xD800) << 10) + (low as u32 - 0xDC00)
            } else {
                first as u32
            };
            Ok(Value::Number(code as f64))
        }
        "LENB" => {
            let source = units(args, 0)?;
            if args.len() != 1 {
                return Err(ExcelError::Value);
            }
            Ok(Value::Number(source.len() as f64))
        }
        "CONCAT" | "CONCATENATE" => {
            if args.is_empty() {
                return Err(ExcelError::Value);
            }
            let mut out = Vec::new();
            for value in args.iter().flat_map(Arg::flatten) {
                append(&mut out, &value.to_utf16()?)?;
            }
            Ok(Value::from_utf16(out))
        }
        "TEXTJOIN" => {
            if args.len() < 3 {
                return Err(ExcelError::Value);
            }
            let between = units(args, 0)?;
            let skip = args[1].scalar().to_logical()?;
            let mut out = Vec::new();
            let mut seen = false;
            for value in args[2..].iter().flat_map(Arg::flatten) {
                let piece = value.to_utf16()?;
                if skip && piece.is_empty() {
                    continue;
                }
                if seen {
                    append(&mut out, &between)?;
                }
                append(&mut out, &piece)?;
                seen = true;
            }
            Ok(Value::from_utf16(out))
        }
        "REPT" => {
            if args.len() != 2 {
                return Err(ExcelError::Value);
            }
            let piece = units(args, 0)?;
            let count = args[1].scalar().to_number()?.trunc();
            if count < 0.0 || piece.len() as f64 * count > LIMIT as f64 {
                return Err(ExcelError::Value);
            }
            if piece.is_empty() {
                return Ok(Value::Text(String::new()));
            }
            Ok(Value::from_utf16(piece.repeat(count as usize)))
        }
        "CLEAN" => {
            if args.len() != 1 {
                return Err(ExcelError::Value);
            }
            Ok(Value::from_utf16(
                units(args, 0)?
                    .into_iter()
                    .filter(|unit| *unit >= 32)
                    .collect(),
            ))
        }
        "PROPER" => {
            if args.len() != 1 {
                return Err(ExcelError::Value);
            }
            let source = units(args, 0)?;
            let mut out = Vec::new();
            let mut starting = true;
            for ch in char::decode_utf16(source) {
                match ch {
                    Err(error) => {
                        out.push(error.unpaired_surrogate());
                        starting = true;
                    }
                    Ok(ch) => {
                        if starting {
                            let mut upper = ch.to_uppercase();
                            let first = upper.next().unwrap();
                            let chosen = if upper.next().is_none() { first } else { ch };
                            out.extend_from_slice(chosen.encode_utf16(&mut [0; 2]));
                        } else {
                            for lower in ch.to_lowercase() {
                                out.extend_from_slice(lower.encode_utf16(&mut [0; 2]));
                            }
                        }
                        starting = !ch.is_alphabetic();
                    }
                }
            }
            Ok(Value::from_utf16(out))
        }
        "VALUETOTEXT" | "ARRAYTOTEXT" => {
            if args.is_empty() || args.len() > 2 {
                return Err(ExcelError::Value);
            }
            let strict = match args.get(1).map(Arg::scalar) {
                None | Some(Value::Blank) => false,
                Some(value) => match value.to_number()? {
                    0.0 => false,
                    1.0 => true,
                    _ => return Err(ExcelError::Value),
                },
            };
            if name == "VALUETOTEXT" {
                return Ok(Value::from_utf16(written(&args[0].scalar(), strict)));
            }
            let block = args[0].as_range();
            let mut out = Vec::new();
            if strict {
                out.push(123);
            }
            for row in 0..block.height {
                for col in 0..block.width {
                    if col > 0 {
                        out.push(44);
                        if !strict {
                            out.push(32);
                        }
                    } else if row > 0 {
                        out.push(if strict { 59 } else { 44 });
                        if !strict {
                            out.push(32);
                        }
                    }
                    out.extend(written(&block.at(col, row), strict));
                }
            }
            if strict {
                out.push(125);
            }
            Ok(Value::from_utf16(out))
        }
        "TEXT" => {
            if args.len() != 2 {
                return Err(ExcelError::Value);
            }
            let format = args[1].scalar().to_text()?;
            Ok(Value::from_utf16(crate::numfmt::format_text_utf16(
                &units(args, 0)?,
                &format,
            )))
        }
        "TEXTBEFORE" | "TEXTAFTER" => {
            if !(2..=6).contains(&args.len()) {
                return Err(ExcelError::Value);
            }
            let given = |at| {
                args.get(at)
                    .filter(|arg: &&Arg| !matches!(arg.scalar(), Value::Blank))
            };
            let hay = units(args, 0)?;
            let mut needles = args[1]
                .flatten()
                .iter()
                .map(Value::to_utf16)
                .collect::<Result<Vec<_>, _>>()?;
            let instance = given(2)
                .map(|arg| arg.scalar().to_number())
                .transpose()?
                .unwrap_or(1.0)
                .trunc() as i64;
            let ignore = given(3)
                .map(|arg| arg.scalar().to_number())
                .transpose()?
                .unwrap_or(0.0)
                != 0.0;
            let match_end = given(4)
                .map(|arg| arg.scalar().to_number())
                .transpose()?
                .unwrap_or(0.0)
                != 0.0;
            if instance > 0 && needles.iter().all(Vec::is_empty) {
                return Ok(Value::from_utf16(if name == "TEXTBEFORE" {
                    vec![]
                } else {
                    hay
                }));
            }
            let logical_length = char::decode_utf16(hay.iter().copied()).count();
            if instance == 0
                || (!(hay.is_empty() && given(2).is_none())
                    && instance.unsigned_abs() as usize > logical_length)
            {
                return Err(ExcelError::Value);
            }
            let folded = if ignore { fold(&hay) } else { hay.clone() };
            let search = if folded.len() == hay.len() {
                if ignore {
                    needles = needles.iter().map(|needle| fold(needle)).collect();
                }
                folded
            } else {
                hay.clone()
            };
            let mut found = Vec::new();
            let mut at = 0;
            while at <= search.len() {
                if let Some(needle) = needles
                    .iter()
                    .find(|needle| !needle.is_empty() && search[at..].starts_with(needle))
                {
                    found.push((at, needle.len()));
                    at += needle.len();
                } else {
                    at += 1;
                }
            }
            if needles.iter().all(Vec::is_empty) {
                found = vec![(0, 0)];
            }
            if match_end {
                if instance > 0 {
                    found.push((hay.len(), 0));
                } else {
                    found.insert(0, (0, 0));
                }
            }
            let index = if instance > 0 {
                instance - 1
            } else {
                found.len() as i64 + instance
            };
            if index < 0 || index >= found.len() as i64 {
                return given(5)
                    .map(|arg| Ok(arg.scalar()))
                    .unwrap_or(Err(ExcelError::NA));
            }
            let (cut, length) = found[index as usize];
            Ok(Value::from_utf16(if name == "TEXTBEFORE" {
                hay[..cut].to_vec()
            } else {
                hay[cut + length..].to_vec()
            }))
        }
        _ => unreachable!(),
    })())
}
