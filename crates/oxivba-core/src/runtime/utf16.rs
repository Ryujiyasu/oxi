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

#[derive(Default, PartialEq, Eq, PartialOrd, Ord)]
struct TextOrderKey {
    primary: Vec<u32>,
    secondary: Vec<u32>,
    punctuation: usize,
    full: Vec<u32>,
}

fn append_order_run(run: &mut String, key: &mut TextOrderKey) {
    for character in collation_spelling(run, true).chars().filter(|c| *c != '\0') {
        let mut buffer = [0; 2];
        key.full.extend(character.encode_utf16(&mut buffer).iter().map(|u| u32::from(*u)));
        if matches!(character, '-' | '\'') {
            key.punctuation += 1;
        } else {
            key.secondary.extend(character.encode_utf16(&mut buffer).iter().map(|u| u32::from(*u)));
            key.primary.extend(latin_base_letter(character).encode_utf16(&mut buffer).iter().map(|u| u32::from(*u)));
        }
    }
    run.clear();
}

#[inline(never)]
fn text_order_key(units: &[u16]) -> TextOrderKey {
    let mut key = TextOrderKey::default();
    let mut run = String::new();
    for decoded in char::decode_utf16(units.iter().copied()) {
        match decoded {
            Ok(character) => run.push(character),
            Err(error) => {
                append_order_run(&mut run, &mut key);
                let unit = 0x110000 + u32::from(error.unpaired_surrogate());
                key.primary.push(unit);
                key.secondary.push(unit);
                key.full.push(unit);
            }
        }
    }
    append_order_run(&mut run, &mut key);
    key
}

pub(super) fn text_order(left: &[u16], right: &[u16]) -> std::cmp::Ordering {
    if left.is_empty() || right.is_empty() {
        return left.len().cmp(&right.len());
    }
    text_order_key(left).cmp(&text_order_key(right))
}

pub(super) fn text_value_order(
    left: &Value,
    right: &Value,
    line: Option<u32>,
) -> Result<std::cmp::Ordering, RuntimeError> {
    let a = units(left, line)?;
    let b = units(right, line)?;
    // Excel checks an actually empty BSTR before linguistic comparison.
    // A one-byte BSTR has no complete character but bypasses that check:
    // empty < NUL, while odd-one-byte = empty and odd-one-byte = NUL.
    let empty = |value: &Value, units: &[u16]| match value {
        Value::ByteString(bytes) => bytes.is_empty(),
        _ => units.is_empty(),
    };
    Ok(if empty(left, &a) || empty(right, &b) {
        a.len().cmp(&b.len())
    } else {
        text_order_key(&a).cmp(&text_order_key(&b))
    })
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
        Value::Utf16(_) | Value::ByteString(_) => true,
        Value::Array(a) => a.values.iter().any(wide),
        _ => false,
    }
}

#[inline(never)]
pub(super) fn array_storage(value: Value, byte_array: bool, line: Option<u32>) -> Result<Value, RuntimeError> {
    if byte_array {
        if let Some(bytes) = value.string_bytes() {
            return Ok(Value::Array(ArrayValue {
                dimensions: vec![ArrayDimension { lower_bound: 0, length: bytes.len() }],
                values: bytes.iter().map(|b| Value::Byte(*b)).collect(),
                element_default: Box::new(Value::Byte(0)),
                resizable: true,
            }));
        }
    }
    Err(error(RuntimeErrorKind::TypeMismatch, "an array variable requires a compatible array value", line))
}

pub(super) fn decode_ansi(host: Option<&dyn Host>, bytes: &[u8], locale: Option<i64>, line: Option<u32>) -> Result<Vec<u16>, RuntimeError> {
    if let Some(value) = host.and_then(|h| h.decode_ansi(bytes, locale)) { return Ok(value); }
    crate::AnsiCodec::for_locale(locale, 1252, 1252).map(|codec| codec.decode(bytes))
        .ok_or_else(|| invalid_procedure_call("unsupported ANSI conversion locale".into(), line))
}

pub(super) fn encode_ansi(host: Option<&dyn Host>, units: &[u16], locale: Option<i64>, line: Option<u32>) -> Result<Vec<u8>, RuntimeError> {
    if let Some(value) = host.and_then(|h| h.encode_ansi(units, locale)) { return Ok(value); }
    crate::AnsiCodec::for_locale(locale, 1252, 1252).map(|codec| codec.encode(units))
        .ok_or_else(|| invalid_procedure_call("unsupported ANSI conversion locale".into(), line))
}

#[inline(never)]
pub(super) fn ansi_builtin(name: &str, args: &[Value], line: Option<u32>, host: Option<&dyn Host>) -> Option<Result<Value, RuntimeError>> {
    if !["chr", "asc", "string", "strconv"].iter().any(|n| name.trim_end_matches('$').eq_ignore_ascii_case(n)) { return None; }
    let name = name.to_ascii_lowercase();
    let name = name.trim_end_matches('$');
    if !matches!(name, "chr" | "asc" | "string" | "strconv") { return None; }
    if name == "strconv" { return Some(call_text_conversion_builtin(name, args, line, host)); }
    Some((|| {
        let required = if name=="string" {2} else {1};
        if args.len()!=required { return Err(error(RuntimeErrorKind::ArgumentCount,"incorrect character argument count",line)); }
        if args.iter().any(|v| matches!(v, Value::Error(_))) {
            return Err(error(RuntimeErrorKind::TypeMismatch,"Error is not character data",line));
        }
        if args.iter().any(|v| matches!(v, Value::Nothing)) {
            return Err(error(RuntimeErrorKind::ObjectVariableNotSet,"Nothing is not character data",line));
        }
        if name=="asc" {
            let value=byte_array_string(&args[0],line).transpose()?.unwrap_or_else(||args[0].clone());
            let u=units(&value,line)?;
            let first=*u.first().ok_or_else(||invalid_procedure_call("Asc requires a character".into(),line))?;
            let bytes=encode_ansi(host,&[first],None,line)?;
            let code=if bytes.len()>1 {u16::from_be_bytes([bytes[0],bytes[1]])} else {bytes[0] as u16};
            return Ok(Value::Int16(code as i16));
        }
        if name=="chr" {
            let code=integer_argument(&args[0],line)?;
            if !(0..=255).contains(&code) {return Err(invalid_procedure_call("Chr takes a byte".into(),line));}
            return decode_ansi(host,&[code as u8],None,line).map(Value::from_utf16);
        }
        let length=nonnegative_length(&args[0],line)?;
        if length>1_000_000 {return Err(error(RuntimeErrorKind::Overflow,"String result exceeds runtime limit",line));}
        if matches!(args[1],Value::Null) {return Ok(Value::Null);}
        let unit=if args[1].is_string() {
            *units(&args[1],line)?.first().ok_or_else(||invalid_procedure_call("String requires a character".into(),line))?
        } else {
            let code=integer_argument(&args[1],line)?;
            if !(-32768..=32767).contains(&code) {return Err(error(RuntimeErrorKind::Overflow,"overflow",line));}
            decode_ansi(host,&[(code & 255) as u8],None,line)?[0]
        };
        Ok(Value::from_utf16(vec![unit;length]))
    })())
}

fn string_argument_positions(name: &str, count: usize) -> &'static [usize] {
    match name.trim_end_matches('$') {
        "len" | "lenb" | "asc" | "ascb" | "ascw" | "left" | "leftb"
        | "right" | "rightb" | "mid" | "midb" | "strreverse" | "trim"
        | "ltrim" | "rtrim" | "ucase" | "lcase" | "cstr" | "val" => &[0],
        "instr" | "instrb" if count == 2 => &[0, 1],
        "instr" | "instrb" => &[1, 2],
        "instrrev" | "strcomp" | "split" => &[0, 1],
        "replace" => &[0, 1, 2],
        "join" | "filter" => &[1],
        _ => &[],
    }
}

#[inline(never)]
pub(super) fn byte_array_arguments(
    name: &str,
    args: &[Value],
    line: Option<u32>,
) -> Option<Result<Vec<Value>, RuntimeError>> {
    let positions = string_argument_positions(name, args.len());
    let byte_array = |value: &Value| matches!(value, Value::Array(array)
        if array.dimensions.len() <= 1 && matches!(array.element_default.as_ref(), Value::Byte(_)));
    if !positions.iter().any(|i| args.get(*i).is_some_and(byte_array)) {
        return None;
    }
    // Convert only String argument positions. Join/Filter keep their input
    // arrays, and numeric arguments retain their own conversion rules.
    Some(args.iter().enumerate().map(|(index, value)| {
        if positions.contains(&index) {
            if let Some(converted) = byte_array_string(value, line) {
                return converted;
            }
        }
        Ok(value.clone())
    }).collect())
}

fn byte_builtin(name: &str, args: &[Value], line: Option<u32>, compare_text: bool) -> Option<Result<Value, RuntimeError>> {
    let byte_form = matches!(name, "lenb" | "leftb" | "rightb" | "midb" | "instrb" | "ascb");
    let odd_char_slice = matches!(name, "left" | "right" | "mid") && matches!(args.first(), Some(Value::ByteString(_)));
    if !odd_char_slice && !byte_form { return None; }
    let fail = |m: &str| invalid_procedure_call(m.to_owned(), line);
    let count = |min,max| if args.len() < min || args.len() > max {Err(error(RuntimeErrorKind::ArgumentCount,"incorrect byte string argument count",line))} else {Ok(())};
    let bytes = |v: &Value| -> Result<Vec<u8>, RuntimeError> {
        if let Some(bytes) = v.string_bytes() { return Ok(bytes.into_owned()); }
        units(v,line).map(|u|u.into_iter().flat_map(u16::to_le_bytes).collect())
    };
    Some((|| {
        if name == "instrb" {
            count(2,4)?;
            let (start,source,needle,mode) = if args.len() == 2 {(1,&args[0],&args[1],None)} else {(integer_argument(&args[0],line)?,&args[1],&args[2],args.get(3))};
            if start < 1 { return Err(fail("String position must be positive")); }
            if matches!(source,Value::Null) || matches!(needle,Value::Null) {return Ok(Value::Null);}
            let mut hay=bytes(source)?;let mut find=bytes(needle)?;
            if compare_mode(mode,compare_text,line,true)? {
                let folded=|raw:&[u8]| {let u=raw.chunks_exact(2).map(|b|u16::from_le_bytes([b[0],b[1]])).collect::<Vec<_>>();let mut result=fold(&u,true).into_iter().flat_map(u16::to_le_bytes).collect::<Vec<_>>();if raw.len()%2==1 {result.push(raw[raw.len()-1]);}result};
                hay=folded(&hay);find=folded(&find);
            }
            let from=usize::try_from(start-1).unwrap_or(usize::MAX);
            if find.is_empty() {return Ok(Value::Integer(if from<=hay.len(){start}else{0}));}
            let found=hay.get(from..).and_then(|suffix|suffix.windows(find.len()).position(|w|w==find));
            return Ok(Value::Integer(found.map_or(0,|i|(from+i+1) as i64)));
        }
        count(if name=="lenb"||name=="ascb"{1}else{2},if name=="midb"||name=="mid"{3}else if name=="lenb"||name=="ascb"{1}else{2})?;
        if matches!(args[0],Value::Null) {return if name=="ascb" {Err(invalid_null(line))} else {Ok(Value::Null)};}
        let raw=bytes(&args[0])?;
        if name=="lenb" {return Ok(Value::Integer(raw.len() as i64));}
        if name=="ascb" {return raw.first().map(|b|Value::Byte(*b)).ok_or_else(||fail("AscB requires a non-empty String"));}
        let scale=if byte_form {1usize}else{2usize};
        let amount=integer_argument(&args[1],line)?;
        let is_mid=name=="midb"||name=="mid";
        if amount < if is_mid {1}else{0} {return Err(fail("negative string position or length"));}
        let start=if is_mid {usize::try_from(amount-1).unwrap_or(usize::MAX).saturating_mul(scale).min(raw.len())}
            else if name=="rightb"||name=="right" {raw.len().saturating_sub(usize::try_from(amount).unwrap_or(usize::MAX).saturating_mul(scale))} else {0};
        let length=if is_mid {match args.get(2) {Some(v)=>{let n=integer_argument(v,line)?;if n<0{return Err(fail("negative string length"));}usize::try_from(n).unwrap_or(usize::MAX).saturating_mul(scale)},None=>raw.len()}}
            else {usize::try_from(amount).unwrap_or(usize::MAX).saturating_mul(scale)};
        let end=start.saturating_add(length).min(raw.len());
        Ok(Value::from_bstr_bytes(raw[start..end].to_vec()))
    })())
}

pub(super) fn builtin(
    name: &str,
    args: &[Value],
    line: Option<u32>,
    compare_text: bool,
) -> Option<Result<Value, RuntimeError>> {
    let name = name.trim_end_matches('$');
    if matches!(name,"lenb"|"leftb"|"rightb"|"midb"|"instrb"|"ascb")
        || (matches!(name,"left"|"right"|"mid") && matches!(args.first(),Some(Value::ByteString(_)))) {
        if let Some(value)=byte_builtin(name,args,line,compare_text){return Some(value);}
    }
    let reverse_pair = name == "strreverse"
        && args.iter().any(|v| {
            v.string_units()
                .is_some_and(|u| u.iter().any(|c| (0xd800..=0xdfff).contains(c)))
        });
    let scalar_error = !matches!(name, "cstr" | "cvar")
        && string_argument_positions(name, args.len()).iter().any(|i|
            args.get(*i).is_some_and(|v| matches!(v, Value::Nothing | Value::Error(_))));
    if !args.iter().any(wide) && !reverse_pair && !scalar_error && name != "len" {
        return None;
    }
    if !matches!(
        name,
        "cstr"
            | "cvar"
            | "len"
            | "lenb"
            | "ascw"
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
        "len" | "lenb" | "ascw" => {
            count(1, 1)?;
            if matches!(name, "len" | "lenb") && matches!(args[0], Value::Null) {
                return Ok(Value::Null);
            }
            let u = units(&args[0], line)?;
            if name == "len" || name == "lenb" {
                return Ok(Value::Integer(
                    u.len() as i64 * if name == "lenb" { 2 } else { 1 },
                ));
            }
            let first = *u.first().ok_or_else(|| fail("empty string"))?;
            Ok(Value::Int16(first as i16))
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
            let order = if mode(2)? {
                text_value_order(&args[0], &args[1], line)?
            } else {
                let a = units(&args[0], line)?;
                let b = units(&args[1], line)?;
                args[0].binary_string_cmp(&args[1]).unwrap_or_else(||a.cmp(&b))
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
                return if reverse { Err(invalid_null(line)) } else { Ok(Value::Null) };
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
            if matches!(a.element_default.as_ref(), Value::Byte(_)) {
                return Err(fail("Join does not accept a Byte array"));
            }
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
        _ => unreachable!("filtered string operation"),
    })())
}
