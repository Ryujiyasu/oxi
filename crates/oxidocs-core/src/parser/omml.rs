// This Source Code Form is subject to the terms of the Mozilla Public
// License, v. 2.0. If a copy of the MPL was not distributed with this
// file, You can obtain one at https://mozilla.org/MPL/2.0/.

//! OMML (Office Math Markup Language) parser — XML → `MathBlock` tree.
//!
//! Consumes `quick_xml` events starting from an already-opened
//! `<m:oMath>` or `<m:oMathPara>` element and builds the IR tree.
//!
//! This is a Phase 2 skeleton: leaf runs, sequences, and 6 common
//! primitives (Fraction, Superscript, Subscript, SubSuperscript,
//! Radical, Delimiter) are fully parsed. Other primitives (Nary,
//! Matrix, Accent, Bar, Limit, etc.) are recognized but their children
//! collapse to a flat `Seq` — Phase 3 extends to full tree fidelity.

use quick_xml::events::Event;
use quick_xml::reader::Reader;

use crate::ir::{
    MathBlock, MathExpr, MathAlignment, FracBarType,
};
use crate::parser::ParseError;

fn local(name: &[u8]) -> String {
    let s = std::str::from_utf8(name).unwrap_or("");
    match s.rfind(':') {
        Some(pos) => s[pos + 1..].to_string(),
        None => s.to_string(),
    }
}

/// Parse `<m:oMath>` (inline) content. Reader should have just consumed
/// the opening tag; reads until matching `</m:oMath>`.
pub fn parse_omath_inline(
    reader: &mut Reader<&[u8]>,
) -> Result<MathBlock, ParseError> {
    let exprs = parse_expr_sequence(reader, "oMath")?;
    Ok(MathBlock::Inline(exprs))
}

/// Parse `<m:oMathPara>` (display) content. Reader should have just
/// consumed the opening tag; reads until matching `</m:oMathPara>`.
pub fn parse_omath_para(
    reader: &mut Reader<&[u8]>,
) -> Result<MathBlock, ParseError> {
    // oMathPara contains optional <m:oMathParaPr> + one or more <m:oMath>.
    // We currently merge all child <m:oMath> into one display block.
    let mut jc = MathAlignment::Center;
    let mut content: Vec<MathExpr> = Vec::new();
    loop {
        match reader.read_event() {
            Ok(Event::Start(e)) => {
                let tag = local(e.name().as_ref());
                if tag == "oMathParaPr" {
                    // Read until </oMathParaPr>, looking for <m:jc>.
                    loop {
                        match reader.read_event() {
                            Ok(Event::Empty(ee)) | Ok(Event::Start(ee)) => {
                                if local(ee.name().as_ref()) == "jc" {
                                    for attr in ee.attributes().flatten() {
                                        if local(attr.key.as_ref()) == "val" {
                                            let v = String::from_utf8_lossy(&attr.value).to_string();
                                            jc = match v.as_str() {
                                                "left" => MathAlignment::Left,
                                                "right" => MathAlignment::Right,
                                                "centerGroup" => MathAlignment::CenterGroup,
                                                _ => MathAlignment::Center,
                                            };
                                        }
                                    }
                                }
                            }
                            Ok(Event::End(ee)) if local(ee.name().as_ref()) == "oMathParaPr" => break,
                            Ok(Event::Eof) => return Err(ParseError::MissingPart(
                                "unexpected EOF in oMathParaPr".to_string())),
                            _ => {}
                        }
                    }
                } else if tag == "oMath" {
                    // Collect inner expressions directly into our content.
                    let inner = parse_expr_sequence(reader, "oMath")?;
                    content.extend(inner);
                } else {
                    // Unknown — best-effort, try to parse as a single expr.
                    if let Some(expr) = parse_single_expr(reader, &tag, &e)? {
                        content.push(expr);
                    }
                }
            }
            Ok(Event::End(e)) if local(e.name().as_ref()) == "oMathPara" => break,
            Ok(Event::Eof) => return Err(ParseError::MissingPart(
                "unexpected EOF in oMathPara".to_string())),
            _ => {}
        }
    }
    Ok(MathBlock::Display { content, jc })
}

/// Parse a sequence of math expressions, reading events until the
/// matching closing tag (e.g., `</m:oMath>`, `</m:e>`, `</m:num>`).
fn parse_expr_sequence(
    reader: &mut Reader<&[u8]>,
    closing_tag: &str,
) -> Result<Vec<MathExpr>, ParseError> {
    let mut out: Vec<MathExpr> = Vec::new();
    loop {
        match reader.read_event() {
            Ok(Event::Start(e)) => {
                let tag = local(e.name().as_ref());
                if let Some(expr) = parse_single_expr(reader, &tag, &e)? {
                    out.push(expr);
                }
            }
            Ok(Event::Empty(_)) => {
                // Empty tags like <m:deg/> inside radicals — skip here.
            }
            Ok(Event::End(e)) if local(e.name().as_ref()) == closing_tag => break,
            Ok(Event::Eof) => return Err(ParseError::MissingPart(
                format!("unexpected EOF, expected </m:{closing_tag}>"))),
            _ => {}
        }
    }
    Ok(out)
}

/// Parse a single OMML element (dispatcher). Returns `None` if the tag
/// should be skipped (e.g., properties containers).
fn parse_single_expr(
    reader: &mut Reader<&[u8]>,
    tag: &str,
    _open: &quick_xml::events::BytesStart,
) -> Result<Option<MathExpr>, ParseError> {
    match tag {
        // Leaf: math run (text)
        "r" => Ok(Some(parse_run(reader)?)),

        // Fraction
        "f" => Ok(Some(parse_fraction(reader)?)),

        // Scripts
        "sSup" => Ok(Some(parse_ssup(reader)?)),
        "sSub" => Ok(Some(parse_ssub(reader)?)),
        "sSubSup" => Ok(Some(parse_ssubsup(reader)?)),
        "sPre" => Ok(Some(parse_spre(reader)?)),

        // Radical
        "rad" => Ok(Some(parse_radical(reader)?)),

        // Delimiter (brackets around content)
        "d" => Ok(Some(parse_delimiter(reader)?)),

        // Matrix
        "m" => Ok(Some(parse_matrix(reader)?)),

        // N-ary operator
        "nary" => Ok(Some(parse_nary(reader)?)),

        // Accent (hat, tilde, macron, vector)
        "acc" => Ok(Some(parse_accent(reader)?)),

        // Bar (overline/underline)
        "bar" => Ok(Some(parse_bar(reader)?)),

        // Limits (lim_{x→0} / limⁿ)
        "limLow" => Ok(Some(parse_limit(reader, "limLow", crate::ir::LimitPos::Lower)?)),
        "limUpp" => Ok(Some(parse_limit(reader, "limUpp", crate::ir::LimitPos::Upper)?)),

        // Function (sin x, cos y, log z)
        "func" => Ok(Some(parse_func(reader)?)),

        // Group character (underbrace, overbrace)
        "groupChr" => Ok(Some(parse_group_chr(reader)?)),

        // Equation array (stacked equations)
        "eqArr" => Ok(Some(parse_eq_arr(reader)?)),

        // Box (visual grouping, no bar)
        "box" => Ok(Some(parse_box(reader)?)),

        // Bordered box (rectangle border around expr)
        "borderBox" => Ok(Some(parse_border_box(reader)?)),

        // Phantom (reserves space without ink)
        "phant" => Ok(Some(parse_phantom(reader)?)),

        // Properties containers — skip (read through the closing tag)
        "rPr" | "fPr" | "sSubPr" | "sSupPr" | "sSubSupPr" | "sPrePr"
        | "radPr" | "naryPr" | "mPr" | "mcs" | "mc" | "mcPr" | "dPr"
        | "accPr" | "barPr" | "boxPr" | "borderBoxPr" | "limLowPr"
        | "limUppPr" | "phantPr" | "funcPr" | "groupChrPr" | "eqArrPr"
        | "ctrlPr" => {
            skip_until_end(reader, tag)?;
            Ok(None)
        }

        // Fallback: primitives not yet implemented. Read their children
        // as a flat Seq (loses tree structure; Phase 3 will extend).
        _ => {
            let children = parse_expr_sequence(reader, tag)?;
            if children.is_empty() {
                Ok(None)
            } else if children.len() == 1 {
                Ok(Some(children.into_iter().next().unwrap()))
            } else {
                Ok(Some(MathExpr::Seq(children)))
            }
        }
    }
}

/// Parse `<m:r>` (math run). Concatenates all `<m:t>` text children.
fn parse_run(reader: &mut Reader<&[u8]>) -> Result<MathExpr, ParseError> {
    let mut text = String::new();
    loop {
        match reader.read_event() {
            Ok(Event::Start(e)) => {
                let tag = local(e.name().as_ref());
                if tag == "t" {
                    // Read text until </m:t>
                    loop {
                        match reader.read_event() {
                            Ok(Event::Text(t)) => {
                                // unescape_and_decode? use unescape() for entities.
                                let raw = t.unescape().unwrap_or_default();
                                text.push_str(&raw);
                            }
                            Ok(Event::End(end)) if local(end.name().as_ref()) == "t" => break,
                            Ok(Event::Eof) => return Err(ParseError::MissingPart(
                                "EOF in m:t".to_string())),
                            _ => {}
                        }
                    }
                } else if tag == "rPr" {
                    skip_until_end(reader, "rPr")?;
                } else {
                    // Unknown inner element — skip.
                    skip_until_end(reader, &tag)?;
                }
            }
            Ok(Event::End(e)) if local(e.name().as_ref()) == "r" => break,
            Ok(Event::Eof) => return Err(ParseError::MissingPart(
                "EOF in m:r".to_string())),
            _ => {}
        }
    }
    if text.is_empty() {
        Ok(MathExpr::Text(String::new()))
    } else {
        Ok(MathExpr::Text(text))
    }
}

/// Parse `<m:f>` fraction (num over den with bar).
fn parse_fraction(reader: &mut Reader<&[u8]>) -> Result<MathExpr, ParseError> {
    let mut num: Option<MathExpr> = None;
    let mut den: Option<MathExpr> = None;
    let mut bar_type = FracBarType::Bar;
    loop {
        match reader.read_event() {
            Ok(Event::Start(e)) => {
                let tag = local(e.name().as_ref());
                match tag.as_str() {
                    "num" => {
                        let children = parse_expr_sequence(reader, "num")?;
                        num = Some(wrap_seq(children));
                    }
                    "den" => {
                        let children = parse_expr_sequence(reader, "den")?;
                        den = Some(wrap_seq(children));
                    }
                    "fPr" => {
                        // Look inside fPr for <m:type m:val="bar|noBar|lin|skw"/>.
                        loop {
                            match reader.read_event() {
                                Ok(Event::Empty(ee)) | Ok(Event::Start(ee)) => {
                                    if local(ee.name().as_ref()) == "type" {
                                        for attr in ee.attributes().flatten() {
                                            if local(attr.key.as_ref()) == "val" {
                                                let v = String::from_utf8_lossy(&attr.value);
                                                bar_type = match v.as_ref() {
                                                    "noBar" => FracBarType::NoBar,
                                                    "lin" => FracBarType::Linear,
                                                    "skw" => FracBarType::Skewed,
                                                    _ => FracBarType::Bar,
                                                };
                                            }
                                        }
                                    }
                                }
                                Ok(Event::End(ee)) if local(ee.name().as_ref()) == "fPr" => break,
                                Ok(Event::Eof) => return Err(ParseError::MissingPart(
                                    "EOF in fPr".to_string())),
                                _ => {}
                            }
                        }
                    }
                    _ => { skip_until_end(reader, &tag)?; }
                }
            }
            Ok(Event::End(e)) if local(e.name().as_ref()) == "f" => break,
            Ok(Event::Eof) => return Err(ParseError::MissingPart(
                "EOF in m:f".to_string())),
            _ => {}
        }
    }
    Ok(MathExpr::Fraction {
        num: Box::new(num.unwrap_or(MathExpr::Text(String::new()))),
        den: Box::new(den.unwrap_or(MathExpr::Text(String::new()))),
        bar_type,
    })
}

/// Parse `<m:sSup>` (base^sup).
fn parse_ssup(reader: &mut Reader<&[u8]>) -> Result<MathExpr, ParseError> {
    let (base, sup) = parse_base_and_script(reader, "sSup", "e", "sup")?;
    Ok(MathExpr::Superscript {
        base: Box::new(base),
        sup: Box::new(sup),
    })
}

/// Parse `<m:sSub>` (base_sub).
fn parse_ssub(reader: &mut Reader<&[u8]>) -> Result<MathExpr, ParseError> {
    let (base, sub) = parse_base_and_script(reader, "sSub", "e", "sub")?;
    Ok(MathExpr::Subscript {
        base: Box::new(base),
        sub: Box::new(sub),
    })
}

/// Parse `<m:sSubSup>` (base_sub^sup). Order in XML: e → sub → sup.
fn parse_ssubsup(reader: &mut Reader<&[u8]>) -> Result<MathExpr, ParseError> {
    let mut base: Option<MathExpr> = None;
    let mut sub: Option<MathExpr> = None;
    let mut sup: Option<MathExpr> = None;
    loop {
        match reader.read_event() {
            Ok(Event::Start(e)) => {
                let tag = local(e.name().as_ref());
                match tag.as_str() {
                    "e" => base = Some(wrap_seq(parse_expr_sequence(reader, "e")?)),
                    "sub" => sub = Some(wrap_seq(parse_expr_sequence(reader, "sub")?)),
                    "sup" => sup = Some(wrap_seq(parse_expr_sequence(reader, "sup")?)),
                    _ => { skip_until_end(reader, &tag)?; }
                }
            }
            Ok(Event::End(e)) if local(e.name().as_ref()) == "sSubSup" => break,
            Ok(Event::Eof) => return Err(ParseError::MissingPart(
                "EOF in m:sSubSup".to_string())),
            _ => {}
        }
    }
    Ok(MathExpr::SubSuperscript {
        base: Box::new(base.unwrap_or(MathExpr::Text(String::new()))),
        sub: Box::new(sub.unwrap_or(MathExpr::Text(String::new()))),
        sup: Box::new(sup.unwrap_or(MathExpr::Text(String::new()))),
    })
}

/// Parse `<m:sPre>` (pre-sub ^pre-sup base). XML order: sub → sup → e.
fn parse_spre(reader: &mut Reader<&[u8]>) -> Result<MathExpr, ParseError> {
    let mut base: Option<MathExpr> = None;
    let mut sub: Option<MathExpr> = None;
    let mut sup: Option<MathExpr> = None;
    loop {
        match reader.read_event() {
            Ok(Event::Start(e)) => {
                let tag = local(e.name().as_ref());
                match tag.as_str() {
                    "sub" => sub = Some(wrap_seq(parse_expr_sequence(reader, "sub")?)),
                    "sup" => sup = Some(wrap_seq(parse_expr_sequence(reader, "sup")?)),
                    "e" => base = Some(wrap_seq(parse_expr_sequence(reader, "e")?)),
                    _ => { skip_until_end(reader, &tag)?; }
                }
            }
            Ok(Event::End(e)) if local(e.name().as_ref()) == "sPre" => break,
            Ok(Event::Eof) => return Err(ParseError::MissingPart(
                "EOF in m:sPre".to_string())),
            _ => {}
        }
    }
    Ok(MathExpr::PreScript {
        base: Box::new(base.unwrap_or(MathExpr::Text(String::new()))),
        sub: Box::new(sub.unwrap_or(MathExpr::Text(String::new()))),
        sup: Box::new(sup.unwrap_or(MathExpr::Text(String::new()))),
    })
}

/// Parse `<m:rad>` (nth root or sqrt). XML: optional `<m:deg/>` + `<m:e>`.
fn parse_radical(reader: &mut Reader<&[u8]>) -> Result<MathExpr, ParseError> {
    let mut degree: Option<Box<MathExpr>> = None;
    let mut radicand: Option<MathExpr> = None;
    loop {
        match reader.read_event() {
            Ok(Event::Start(e)) => {
                let tag = local(e.name().as_ref());
                match tag.as_str() {
                    "deg" => {
                        let children = parse_expr_sequence(reader, "deg")?;
                        if !children.is_empty() {
                            degree = Some(Box::new(wrap_seq(children)));
                        }
                    }
                    "e" => {
                        radicand = Some(wrap_seq(parse_expr_sequence(reader, "e")?));
                    }
                    _ => { skip_until_end(reader, &tag)?; }
                }
            }
            Ok(Event::Empty(_)) => {} // <m:deg/> empty tag means sqrt (no degree)
            Ok(Event::End(e)) if local(e.name().as_ref()) == "rad" => break,
            Ok(Event::Eof) => return Err(ParseError::MissingPart(
                "EOF in m:rad".to_string())),
            _ => {}
        }
    }
    Ok(MathExpr::Radical {
        degree,
        radicand: Box::new(radicand.unwrap_or(MathExpr::Text(String::new()))),
    })
}

/// Helper: parse `<parent>` with `<base_tag>...</base_tag>` and
/// `<script_tag>...</script_tag>` children.
fn parse_base_and_script(
    reader: &mut Reader<&[u8]>,
    parent: &str,
    base_tag: &str,
    script_tag: &str,
) -> Result<(MathExpr, MathExpr), ParseError> {
    let mut base: Option<MathExpr> = None;
    let mut script: Option<MathExpr> = None;
    loop {
        match reader.read_event() {
            Ok(Event::Start(e)) => {
                let tag = local(e.name().as_ref());
                if tag == base_tag {
                    base = Some(wrap_seq(parse_expr_sequence(reader, base_tag)?));
                } else if tag == script_tag {
                    script = Some(wrap_seq(parse_expr_sequence(reader, script_tag)?));
                } else {
                    skip_until_end(reader, &tag)?;
                }
            }
            Ok(Event::End(e)) if local(e.name().as_ref()) == parent => break,
            Ok(Event::Eof) => return Err(ParseError::MissingPart(
                format!("EOF in m:{parent}"))),
            _ => {}
        }
    }
    Ok((
        base.unwrap_or(MathExpr::Text(String::new())),
        script.unwrap_or(MathExpr::Text(String::new())),
    ))
}

/// Wrap a sequence of expressions into a single MathExpr.
/// - Empty → Text("")
/// - Single → that expression
/// - Multiple → Seq(...)
fn wrap_seq(mut children: Vec<MathExpr>) -> MathExpr {
    match children.len() {
        0 => MathExpr::Text(String::new()),
        1 => children.pop().unwrap(),
        _ => MathExpr::Seq(children),
    }
}

/// Parse `<m:d>` delimiter: begChr/endChr/sepChr from dPr, content from e.
fn parse_delimiter(reader: &mut Reader<&[u8]>) -> Result<MathExpr, ParseError> {
    let mut beg: char = '(';
    let mut end: char = ')';
    let mut sep: Option<char> = None;
    let mut content: Option<MathExpr> = None;
    loop {
        match reader.read_event() {
            Ok(Event::Start(e)) => {
                let tag = local(e.name().as_ref());
                match tag.as_str() {
                    "dPr" => {
                        loop {
                            match reader.read_event() {
                                Ok(Event::Empty(ee)) | Ok(Event::Start(ee)) => {
                                    let t = local(ee.name().as_ref());
                                    if t == "begChr" || t == "endChr" || t == "sepChr" {
                                        for attr in ee.attributes().flatten() {
                                            if local(attr.key.as_ref()) == "val" {
                                                let v = String::from_utf8_lossy(&attr.value);
                                                if let Some(c) = v.chars().next() {
                                                    match t.as_str() {
                                                        "begChr" => beg = c,
                                                        "endChr" => end = c,
                                                        "sepChr" => sep = Some(c),
                                                        _ => {}
                                                    }
                                                }
                                            }
                                        }
                                    }
                                }
                                Ok(Event::End(ee)) if local(ee.name().as_ref()) == "dPr" => break,
                                Ok(Event::Eof) => return Err(ParseError::MissingPart(
                                    "EOF in dPr".to_string())),
                                _ => {}
                            }
                        }
                    }
                    "e" => {
                        content = Some(wrap_seq(parse_expr_sequence(reader, "e")?));
                    }
                    _ => { skip_until_end(reader, &tag)?; }
                }
            }
            Ok(Event::End(e)) if local(e.name().as_ref()) == "d" => break,
            Ok(Event::Eof) => return Err(ParseError::MissingPart(
                "EOF in m:d".to_string())),
            _ => {}
        }
    }
    Ok(MathExpr::Delimiter {
        beg,
        end,
        sep,
        content: Box::new(content.unwrap_or(MathExpr::Text(String::new()))),
    })
}

/// Parse `<m:m>` matrix: rows from mr, cells from e within row.
fn parse_matrix(reader: &mut Reader<&[u8]>) -> Result<MathExpr, ParseError> {
    use crate::ir::MathAlignment;
    let mut rows: Vec<Vec<MathExpr>> = Vec::new();
    let mut col_align = MathAlignment::Center;
    loop {
        match reader.read_event() {
            Ok(Event::Start(e)) => {
                let tag = local(e.name().as_ref());
                match tag.as_str() {
                    "mPr" => {
                        // Parse mcJc for column alignment
                        loop {
                            match reader.read_event() {
                                Ok(Event::Empty(ee)) | Ok(Event::Start(ee)) => {
                                    if local(ee.name().as_ref()) == "mcJc" {
                                        for attr in ee.attributes().flatten() {
                                            if local(attr.key.as_ref()) == "val" {
                                                let v = String::from_utf8_lossy(&attr.value);
                                                col_align = match v.as_ref() {
                                                    "left" => MathAlignment::Left,
                                                    "right" => MathAlignment::Right,
                                                    "centerGroup" => MathAlignment::CenterGroup,
                                                    _ => MathAlignment::Center,
                                                };
                                            }
                                        }
                                    }
                                }
                                Ok(Event::End(ee)) if local(ee.name().as_ref()) == "mPr" => break,
                                Ok(Event::Eof) => return Err(ParseError::MissingPart(
                                    "EOF in mPr".to_string())),
                                _ => {}
                            }
                        }
                    }
                    "mr" => {
                        // Parse a row: sequence of <m:e> cells
                        let mut row_cells: Vec<MathExpr> = Vec::new();
                        loop {
                            match reader.read_event() {
                                Ok(Event::Start(ee)) => {
                                    let t = local(ee.name().as_ref());
                                    if t == "e" {
                                        row_cells.push(wrap_seq(parse_expr_sequence(reader, "e")?));
                                    } else {
                                        skip_until_end(reader, &t)?;
                                    }
                                }
                                Ok(Event::End(ee)) if local(ee.name().as_ref()) == "mr" => break,
                                Ok(Event::Eof) => return Err(ParseError::MissingPart(
                                    "EOF in mr".to_string())),
                                _ => {}
                            }
                        }
                        rows.push(row_cells);
                    }
                    _ => { skip_until_end(reader, &tag)?; }
                }
            }
            Ok(Event::End(e)) if local(e.name().as_ref()) == "m" => break,
            Ok(Event::Eof) => return Err(ParseError::MissingPart(
                "EOF in m:m".to_string())),
            _ => {}
        }
    }
    let cols = rows.iter().map(|r| r.len()).max().unwrap_or(0);
    Ok(MathExpr::Matrix { rows, cols, col_align })
}

/// Parse `<m:nary>`: n-ary operator (sum, integral, product).
fn parse_nary(reader: &mut Reader<&[u8]>) -> Result<MathExpr, ParseError> {
    use crate::ir::LimLoc;
    let mut op: char = '\u{2211}'; // default ∑
    let mut lim_loc = LimLoc::SubSup;
    let mut grow = false;
    let mut sub: Option<Box<MathExpr>> = None;
    let mut sup: Option<Box<MathExpr>> = None;
    let mut operand: Option<MathExpr> = None;
    loop {
        match reader.read_event() {
            Ok(Event::Start(e)) => {
                let tag = local(e.name().as_ref());
                match tag.as_str() {
                    "naryPr" => {
                        loop {
                            match reader.read_event() {
                                Ok(Event::Empty(ee)) | Ok(Event::Start(ee)) => {
                                    let t = local(ee.name().as_ref());
                                    match t.as_str() {
                                        "chr" => {
                                            for attr in ee.attributes().flatten() {
                                                if local(attr.key.as_ref()) == "val" {
                                                    let v = String::from_utf8_lossy(&attr.value);
                                                    if let Some(c) = v.chars().next() { op = c; }
                                                }
                                            }
                                        }
                                        "limLoc" => {
                                            for attr in ee.attributes().flatten() {
                                                if local(attr.key.as_ref()) == "val" {
                                                    let v = String::from_utf8_lossy(&attr.value);
                                                    lim_loc = match v.as_ref() {
                                                        "undOvr" => LimLoc::UndOvr,
                                                        _ => LimLoc::SubSup,
                                                    };
                                                }
                                            }
                                        }
                                        "grow" => {
                                            for attr in ee.attributes().flatten() {
                                                if local(attr.key.as_ref()) == "val" {
                                                    let v = String::from_utf8_lossy(&attr.value);
                                                    grow = v.as_ref() == "1" || v.as_ref() == "true" || v.as_ref() == "on";
                                                }
                                            }
                                        }
                                        _ => {}
                                    }
                                }
                                Ok(Event::End(ee)) if local(ee.name().as_ref()) == "naryPr" => break,
                                Ok(Event::Eof) => return Err(ParseError::MissingPart(
                                    "EOF in naryPr".to_string())),
                                _ => {}
                            }
                        }
                    }
                    "sub" => sub = Some(Box::new(wrap_seq(parse_expr_sequence(reader, "sub")?))),
                    "sup" => sup = Some(Box::new(wrap_seq(parse_expr_sequence(reader, "sup")?))),
                    "e" => operand = Some(wrap_seq(parse_expr_sequence(reader, "e")?)),
                    _ => { skip_until_end(reader, &tag)?; }
                }
            }
            Ok(Event::End(e)) if local(e.name().as_ref()) == "nary" => break,
            Ok(Event::Eof) => return Err(ParseError::MissingPart(
                "EOF in m:nary".to_string())),
            _ => {}
        }
    }
    Ok(MathExpr::Nary {
        op,
        sub,
        sup,
        operand: Box::new(operand.unwrap_or(MathExpr::Text(String::new()))),
        lim_loc,
        grow,
    })
}

/// Parse `<m:acc>` accent: chr from accPr, base from e.
fn parse_accent(reader: &mut Reader<&[u8]>) -> Result<MathExpr, ParseError> {
    let mut accent: char = '\u{0302}'; // default ̂ hat
    let mut base: Option<MathExpr> = None;
    loop {
        match reader.read_event() {
            Ok(Event::Start(e)) => {
                let tag = local(e.name().as_ref());
                match tag.as_str() {
                    "accPr" => {
                        loop {
                            match reader.read_event() {
                                Ok(Event::Empty(ee)) | Ok(Event::Start(ee)) => {
                                    if local(ee.name().as_ref()) == "chr" {
                                        for attr in ee.attributes().flatten() {
                                            if local(attr.key.as_ref()) == "val" {
                                                let v = String::from_utf8_lossy(&attr.value);
                                                if let Some(c) = v.chars().next() { accent = c; }
                                            }
                                        }
                                    }
                                }
                                Ok(Event::End(ee)) if local(ee.name().as_ref()) == "accPr" => break,
                                Ok(Event::Eof) => return Err(ParseError::MissingPart(
                                    "EOF in accPr".to_string())),
                                _ => {}
                            }
                        }
                    }
                    "e" => base = Some(wrap_seq(parse_expr_sequence(reader, "e")?)),
                    _ => { skip_until_end(reader, &tag)?; }
                }
            }
            Ok(Event::End(e)) if local(e.name().as_ref()) == "acc" => break,
            Ok(Event::Eof) => return Err(ParseError::MissingPart("EOF in m:acc".to_string())),
            _ => {}
        }
    }
    Ok(MathExpr::Accent {
        accent,
        base: Box::new(base.unwrap_or(MathExpr::Text(String::new()))),
    })
}

/// Parse `<m:bar>`: overline (pos=top) or underline (pos=bot), base from e.
fn parse_bar(reader: &mut Reader<&[u8]>) -> Result<MathExpr, ParseError> {
    use crate::ir::BarPos;
    let mut pos = BarPos::Top;
    let mut base: Option<MathExpr> = None;
    loop {
        match reader.read_event() {
            Ok(Event::Start(e)) => {
                let tag = local(e.name().as_ref());
                match tag.as_str() {
                    "barPr" => {
                        loop {
                            match reader.read_event() {
                                Ok(Event::Empty(ee)) | Ok(Event::Start(ee)) => {
                                    if local(ee.name().as_ref()) == "pos" {
                                        for attr in ee.attributes().flatten() {
                                            if local(attr.key.as_ref()) == "val" {
                                                let v = String::from_utf8_lossy(&attr.value);
                                                pos = match v.as_ref() {
                                                    "bot" => BarPos::Bot,
                                                    _ => BarPos::Top,
                                                };
                                            }
                                        }
                                    }
                                }
                                Ok(Event::End(ee)) if local(ee.name().as_ref()) == "barPr" => break,
                                Ok(Event::Eof) => return Err(ParseError::MissingPart(
                                    "EOF in barPr".to_string())),
                                _ => {}
                            }
                        }
                    }
                    "e" => base = Some(wrap_seq(parse_expr_sequence(reader, "e")?)),
                    _ => { skip_until_end(reader, &tag)?; }
                }
            }
            Ok(Event::End(e)) if local(e.name().as_ref()) == "bar" => break,
            Ok(Event::Eof) => return Err(ParseError::MissingPart("EOF in m:bar".to_string())),
            _ => {}
        }
    }
    Ok(MathExpr::Bar {
        pos,
        base: Box::new(base.unwrap_or(MathExpr::Text(String::new()))),
    })
}

/// Parse `<m:limLow>` or `<m:limUpp>`: base from e, limit expr from lim.
fn parse_limit(
    reader: &mut Reader<&[u8]>,
    tag_name: &str,
    pos: crate::ir::LimitPos,
) -> Result<MathExpr, ParseError> {
    let mut base: Option<MathExpr> = None;
    let mut lim: Option<MathExpr> = None;
    loop {
        match reader.read_event() {
            Ok(Event::Start(e)) => {
                let tag = local(e.name().as_ref());
                match tag.as_str() {
                    "e" => base = Some(wrap_seq(parse_expr_sequence(reader, "e")?)),
                    "lim" => lim = Some(wrap_seq(parse_expr_sequence(reader, "lim")?)),
                    _ => { skip_until_end(reader, &tag)?; }
                }
            }
            Ok(Event::End(e)) if local(e.name().as_ref()) == tag_name => break,
            Ok(Event::Eof) => return Err(ParseError::MissingPart(
                format!("EOF in m:{tag_name}"))),
            _ => {}
        }
    }
    Ok(MathExpr::Limit {
        base: Box::new(base.unwrap_or(MathExpr::Text(String::new()))),
        lim: Box::new(lim.unwrap_or(MathExpr::Text(String::new()))),
        pos,
    })
}

/// Parse `<m:func>` function: fName + e argument.
fn parse_func(reader: &mut Reader<&[u8]>) -> Result<MathExpr, ParseError> {
    let mut name: Option<MathExpr> = None;
    let mut arg: Option<MathExpr> = None;
    loop {
        match reader.read_event() {
            Ok(Event::Start(e)) => {
                let tag = local(e.name().as_ref());
                match tag.as_str() {
                    "fName" => name = Some(wrap_seq(parse_expr_sequence(reader, "fName")?)),
                    "e" => arg = Some(wrap_seq(parse_expr_sequence(reader, "e")?)),
                    _ => { skip_until_end(reader, &tag)?; }
                }
            }
            Ok(Event::End(e)) if local(e.name().as_ref()) == "func" => break,
            Ok(Event::Eof) => return Err(ParseError::MissingPart("EOF in m:func".to_string())),
            _ => {}
        }
    }
    Ok(MathExpr::Function {
        name: Box::new(name.unwrap_or(MathExpr::Text(String::new()))),
        arg: Box::new(arg.unwrap_or(MathExpr::Text(String::new()))),
    })
}

/// Parse `<m:groupChr>` underbrace/overbrace: chr + pos from groupChrPr, base from e.
fn parse_group_chr(reader: &mut Reader<&[u8]>) -> Result<MathExpr, ParseError> {
    use crate::ir::BarPos;
    let mut chr: char = '\u{23DF}'; // default bottom curly bracket
    let mut pos = BarPos::Bot;
    let mut base: Option<MathExpr> = None;
    loop {
        match reader.read_event() {
            Ok(Event::Start(e)) => {
                let tag = local(e.name().as_ref());
                match tag.as_str() {
                    "groupChrPr" => {
                        loop {
                            match reader.read_event() {
                                Ok(Event::Empty(ee)) | Ok(Event::Start(ee)) => {
                                    let t = local(ee.name().as_ref());
                                    if t == "chr" {
                                        for attr in ee.attributes().flatten() {
                                            if local(attr.key.as_ref()) == "val" {
                                                let v = String::from_utf8_lossy(&attr.value);
                                                if let Some(c) = v.chars().next() { chr = c; }
                                            }
                                        }
                                    } else if t == "pos" {
                                        for attr in ee.attributes().flatten() {
                                            if local(attr.key.as_ref()) == "val" {
                                                let v = String::from_utf8_lossy(&attr.value);
                                                pos = match v.as_ref() {
                                                    "top" => BarPos::Top,
                                                    _ => BarPos::Bot,
                                                };
                                            }
                                        }
                                    }
                                }
                                Ok(Event::End(ee)) if local(ee.name().as_ref()) == "groupChrPr" => break,
                                Ok(Event::Eof) => return Err(ParseError::MissingPart(
                                    "EOF in groupChrPr".to_string())),
                                _ => {}
                            }
                        }
                    }
                    "e" => base = Some(wrap_seq(parse_expr_sequence(reader, "e")?)),
                    _ => { skip_until_end(reader, &tag)?; }
                }
            }
            Ok(Event::End(e)) if local(e.name().as_ref()) == "groupChr" => break,
            Ok(Event::Eof) => return Err(ParseError::MissingPart(
                "EOF in m:groupChr".to_string())),
            _ => {}
        }
    }
    Ok(MathExpr::GroupChar {
        chr,
        pos,
        base: Box::new(base.unwrap_or(MathExpr::Text(String::new()))),
    })
}

/// Parse `<m:eqArr>` equation array: stacked expressions (one per e).
fn parse_eq_arr(reader: &mut Reader<&[u8]>) -> Result<MathExpr, ParseError> {
    let mut items: Vec<MathExpr> = Vec::new();
    loop {
        match reader.read_event() {
            Ok(Event::Start(e)) => {
                let tag = local(e.name().as_ref());
                match tag.as_str() {
                    "e" => items.push(wrap_seq(parse_expr_sequence(reader, "e")?)),
                    _ => { skip_until_end(reader, &tag)?; }
                }
            }
            Ok(Event::End(e)) if local(e.name().as_ref()) == "eqArr" => break,
            Ok(Event::Eof) => return Err(ParseError::MissingPart("EOF in m:eqArr".to_string())),
            _ => {}
        }
    }
    Ok(MathExpr::EqArray(items))
}

/// Parse `<m:box>` visual grouping: content from e.
fn parse_box(reader: &mut Reader<&[u8]>) -> Result<MathExpr, ParseError> {
    let mut inner: Option<MathExpr> = None;
    loop {
        match reader.read_event() {
            Ok(Event::Start(e)) => {
                let tag = local(e.name().as_ref());
                match tag.as_str() {
                    "e" => inner = Some(wrap_seq(parse_expr_sequence(reader, "e")?)),
                    _ => { skip_until_end(reader, &tag)?; }
                }
            }
            Ok(Event::End(e)) if local(e.name().as_ref()) == "box" => break,
            Ok(Event::Eof) => return Err(ParseError::MissingPart("EOF in m:box".to_string())),
            _ => {}
        }
    }
    Ok(MathExpr::BoxExpr(Box::new(
        inner.unwrap_or(MathExpr::Text(String::new())),
    )))
}

/// Parse `<m:borderBox>`: box with border, sides from borderBoxPr.
fn parse_border_box(reader: &mut Reader<&[u8]>) -> Result<MathExpr, ParseError> {
    use crate::ir::BoxBorders;
    let mut sides = BoxBorders::default();
    let mut base: Option<MathExpr> = None;
    // Default: all 4 sides drawn unless borderBoxPr specifies hideTop/hideBot/etc.
    sides.top = true; sides.bot = true; sides.left = true; sides.right = true;
    loop {
        match reader.read_event() {
            Ok(Event::Start(e)) => {
                let tag = local(e.name().as_ref());
                match tag.as_str() {
                    "borderBoxPr" => {
                        loop {
                            match reader.read_event() {
                                Ok(Event::Empty(ee)) | Ok(Event::Start(ee)) => {
                                    let t = local(ee.name().as_ref());
                                    match t.as_str() {
                                        "hideTop" => sides.top = false,
                                        "hideBot" => sides.bot = false,
                                        "hideLeft" => sides.left = false,
                                        "hideRight" => sides.right = false,
                                        "strikeH" => sides.strikeh = true,
                                        "strikeV" => sides.strikev = true,
                                        "strikeBLTR" => sides.strikebltr = true,
                                        "strikeTLBR" => sides.striketlbr = true,
                                        _ => {}
                                    }
                                }
                                Ok(Event::End(ee)) if local(ee.name().as_ref()) == "borderBoxPr" => break,
                                Ok(Event::Eof) => return Err(ParseError::MissingPart(
                                    "EOF in borderBoxPr".to_string())),
                                _ => {}
                            }
                        }
                    }
                    "e" => base = Some(wrap_seq(parse_expr_sequence(reader, "e")?)),
                    _ => { skip_until_end(reader, &tag)?; }
                }
            }
            Ok(Event::End(e)) if local(e.name().as_ref()) == "borderBox" => break,
            Ok(Event::Eof) => return Err(ParseError::MissingPart(
                "EOF in m:borderBox".to_string())),
            _ => {}
        }
    }
    Ok(MathExpr::BorderBox {
        base: Box::new(base.unwrap_or(MathExpr::Text(String::new()))),
        sides,
    })
}

/// Parse `<m:phant>` phantom: reserves space without ink.
fn parse_phantom(reader: &mut Reader<&[u8]>) -> Result<MathExpr, ParseError> {
    let mut inner: Option<MathExpr> = None;
    loop {
        match reader.read_event() {
            Ok(Event::Start(e)) => {
                let tag = local(e.name().as_ref());
                match tag.as_str() {
                    "e" => inner = Some(wrap_seq(parse_expr_sequence(reader, "e")?)),
                    _ => { skip_until_end(reader, &tag)?; }
                }
            }
            Ok(Event::End(e)) if local(e.name().as_ref()) == "phant" => break,
            Ok(Event::Eof) => return Err(ParseError::MissingPart("EOF in m:phant".to_string())),
            _ => {}
        }
    }
    Ok(MathExpr::Phantom(Box::new(
        inner.unwrap_or(MathExpr::Text(String::new())),
    )))
}

/// Advance the reader past the closing tag matching `end_tag`, balancing
/// nested opens/closes of the same name.
fn skip_until_end(
    reader: &mut Reader<&[u8]>,
    end_tag: &str,
) -> Result<(), ParseError> {
    let mut depth = 1_i32;
    while depth > 0 {
        match reader.read_event() {
            Ok(Event::Start(e)) => {
                if local(e.name().as_ref()) == end_tag {
                    depth += 1;
                }
            }
            Ok(Event::End(e)) => {
                if local(e.name().as_ref()) == end_tag {
                    depth -= 1;
                }
            }
            Ok(Event::Eof) => return Err(ParseError::MissingPart(
                format!("EOF before </m:{end_tag}>"))),
            _ => {}
        }
    }
    Ok(())
}

#[cfg(test)]
mod tests {
    use super::*;

    fn parse(xml: &str) -> MathBlock {
        let mut r = Reader::from_str(xml);
        // Consume opening event
        loop {
            match r.read_event() {
                Ok(Event::Start(e)) => {
                    match local(e.name().as_ref()).as_str() {
                        "oMath" => return parse_omath_inline(&mut r).unwrap(),
                        "oMathPara" => return parse_omath_para(&mut r).unwrap(),
                        _ => continue,
                    }
                }
                Ok(Event::Eof) => panic!("no oMath/oMathPara in XML"),
                _ => {}
            }
        }
    }

    #[test]
    fn parse_simple_run() {
        let block = parse(r#"<m:oMath xmlns:m="x"><m:r><m:t>hello</m:t></m:r></m:oMath>"#);
        match block {
            MathBlock::Inline(exprs) => {
                assert_eq!(exprs.len(), 1);
                match &exprs[0] {
                    MathExpr::Text(s) => assert_eq!(s, "hello"),
                    _ => panic!("expected Text"),
                }
            }
            _ => panic!("expected Inline"),
        }
    }

    #[test]
    fn parse_fraction_a_b() {
        let xml = r#"<m:oMath xmlns:m="x"><m:f>
            <m:num><m:r><m:t>a</m:t></m:r></m:num>
            <m:den><m:r><m:t>b</m:t></m:r></m:den>
        </m:f></m:oMath>"#;
        let block = parse(xml);
        let exprs = match block {
            MathBlock::Inline(v) => v,
            _ => panic!("expected inline"),
        };
        assert_eq!(exprs.len(), 1);
        match &exprs[0] {
            MathExpr::Fraction { num, den, bar_type } => {
                assert!(matches!(**num, MathExpr::Text(ref s) if s == "a"));
                assert!(matches!(**den, MathExpr::Text(ref s) if s == "b"));
                assert_eq!(*bar_type, FracBarType::Bar);
            }
            other => panic!("expected Fraction, got {:?}", other),
        }
    }

    #[test]
    fn parse_superscript_x_2() {
        let xml = r#"<m:oMath xmlns:m="x"><m:sSup>
            <m:e><m:r><m:t>x</m:t></m:r></m:e>
            <m:sup><m:r><m:t>2</m:t></m:r></m:sup>
        </m:sSup></m:oMath>"#;
        let block = parse(xml);
        let exprs = match block {
            MathBlock::Inline(v) => v,
            _ => panic!(),
        };
        assert_eq!(exprs.len(), 1);
        match &exprs[0] {
            MathExpr::Superscript { base, sup } => {
                assert!(matches!(**base, MathExpr::Text(ref s) if s == "x"));
                assert!(matches!(**sup, MathExpr::Text(ref s) if s == "2"));
            }
            _ => panic!("expected Superscript"),
        }
    }

    #[test]
    fn parse_subscript_y_1() {
        let xml = r#"<m:oMath xmlns:m="x"><m:sSub>
            <m:e><m:r><m:t>y</m:t></m:r></m:e>
            <m:sub><m:r><m:t>1</m:t></m:r></m:sub>
        </m:sSub></m:oMath>"#;
        let block = parse(xml);
        let exprs = match block {
            MathBlock::Inline(v) => v,
            _ => panic!(),
        };
        match &exprs[0] {
            MathExpr::Subscript { base, sub } => {
                assert!(matches!(**base, MathExpr::Text(ref s) if s == "y"));
                assert!(matches!(**sub, MathExpr::Text(ref s) if s == "1"));
            }
            _ => panic!("expected Subscript"),
        }
    }

    #[test]
    fn parse_sub_superscript() {
        let xml = r#"<m:oMath xmlns:m="x"><m:sSubSup>
            <m:e><m:r><m:t>x</m:t></m:r></m:e>
            <m:sub><m:r><m:t>1</m:t></m:r></m:sub>
            <m:sup><m:r><m:t>2</m:t></m:r></m:sup>
        </m:sSubSup></m:oMath>"#;
        let block = parse(xml);
        let exprs = match block { MathBlock::Inline(v) => v, _ => panic!() };
        match &exprs[0] {
            MathExpr::SubSuperscript { base, sub, sup } => {
                assert!(matches!(**base, MathExpr::Text(ref s) if s == "x"));
                assert!(matches!(**sub, MathExpr::Text(ref s) if s == "1"));
                assert!(matches!(**sup, MathExpr::Text(ref s) if s == "2"));
            }
            _ => panic!("expected SubSuperscript"),
        }
    }

    #[test]
    fn parse_radical_sqrt() {
        let xml = r#"<m:oMath xmlns:m="x"><m:rad>
            <m:radPr><m:degHide m:val="1"/></m:radPr>
            <m:deg/>
            <m:e><m:r><m:t>x</m:t></m:r></m:e>
        </m:rad></m:oMath>"#;
        let block = parse(xml);
        let exprs = match block { MathBlock::Inline(v) => v, _ => panic!() };
        match &exprs[0] {
            MathExpr::Radical { degree, radicand } => {
                assert!(degree.is_none());
                assert!(matches!(**radicand, MathExpr::Text(ref s) if s == "x"));
            }
            _ => panic!("expected Radical"),
        }
    }

    #[test]
    fn parse_display_math_para_with_center_jc() {
        let xml = r#"<m:oMathPara xmlns:m="x">
            <m:oMathParaPr><m:jc m:val="center"/></m:oMathParaPr>
            <m:oMath><m:r><m:t>E=mc</m:t></m:r></m:oMath>
        </m:oMathPara>"#;
        let block = parse(xml);
        match block {
            MathBlock::Display { content, jc } => {
                assert_eq!(jc, MathAlignment::Center);
                assert_eq!(content.len(), 1);
                assert!(matches!(&content[0], MathExpr::Text(s) if s == "E=mc"));
            }
            _ => panic!("expected Display"),
        }
    }

    #[test]
    fn parse_display_math_para_with_left_jc() {
        let xml = r#"<m:oMathPara xmlns:m="x">
            <m:oMathParaPr><m:jc m:val="left"/></m:oMathParaPr>
            <m:oMath><m:r><m:t>x</m:t></m:r></m:oMath>
        </m:oMathPara>"#;
        let block = parse(xml);
        match block {
            MathBlock::Display { jc, .. } => assert_eq!(jc, MathAlignment::Left),
            _ => panic!(),
        }
    }

    #[test]
    fn parse_multiple_runs_concatenates_as_seq() {
        let xml = r#"<m:oMath xmlns:m="x">
            <m:r><m:t>a</m:t></m:r>
            <m:r><m:t>+</m:t></m:r>
            <m:r><m:t>b</m:t></m:r>
        </m:oMath>"#;
        let block = parse(xml);
        let exprs = match block { MathBlock::Inline(v) => v, _ => panic!() };
        assert_eq!(exprs.len(), 3);
    }

    #[test]
    fn parse_nested_fraction() {
        // (a/b) / c — fraction whose numerator is a fraction
        let xml = r#"<m:oMath xmlns:m="x"><m:f>
            <m:num><m:f>
                <m:num><m:r><m:t>a</m:t></m:r></m:num>
                <m:den><m:r><m:t>b</m:t></m:r></m:den>
            </m:f></m:num>
            <m:den><m:r><m:t>c</m:t></m:r></m:den>
        </m:f></m:oMath>"#;
        let block = parse(xml);
        let exprs = match block { MathBlock::Inline(v) => v, _ => panic!() };
        match &exprs[0] {
            MathExpr::Fraction { num, den, .. } => {
                // Inner num should itself be a fraction
                match &**num {
                    MathExpr::Fraction { num: inner_num, den: inner_den, .. } => {
                        assert!(matches!(**inner_num, MathExpr::Text(ref s) if s == "a"));
                        assert!(matches!(**inner_den, MathExpr::Text(ref s) if s == "b"));
                    }
                    _ => panic!("expected nested Fraction in num"),
                }
                assert!(matches!(**den, MathExpr::Text(ref s) if s == "c"));
            }
            _ => panic!("expected outer Fraction"),
        }
    }

    #[test]
    fn parse_accent_hat() {
        use crate::ir::MathExpr;
        let xml = r#"<m:oMath xmlns:m="x"><m:acc>
            <m:accPr><m:chr m:val="̂"/></m:accPr>
            <m:e><m:r><m:t>x</m:t></m:r></m:e>
        </m:acc></m:oMath>"#;
        let block = parse(xml);
        let exprs = match block { MathBlock::Inline(v) => v, _ => panic!() };
        match &exprs[0] {
            MathExpr::Accent { accent, base } => {
                assert_eq!(*accent, '\u{0302}');
                assert!(matches!(**base, MathExpr::Text(ref s) if s == "x"));
            }
            _ => panic!("expected Accent"),
        }
    }

    #[test]
    fn parse_bar_top() {
        use crate::ir::{MathExpr, BarPos};
        let xml = r#"<m:oMath xmlns:m="x"><m:bar>
            <m:e><m:r><m:t>x</m:t></m:r></m:e>
        </m:bar></m:oMath>"#;
        let block = parse(xml);
        let exprs = match block { MathBlock::Inline(v) => v, _ => panic!() };
        match &exprs[0] {
            MathExpr::Bar { pos, base } => {
                assert_eq!(*pos, BarPos::Top);
                assert!(matches!(**base, MathExpr::Text(ref s) if s == "x"));
            }
            _ => panic!("expected Bar"),
        }
    }

    #[test]
    fn parse_matrix_2x2() {
        use crate::ir::MathExpr;
        let xml = r#"<m:oMath xmlns:m="x"><m:m>
            <m:mPr><m:mcs><m:mc><m:mcPr><m:count m:val="2"/></m:mcPr></m:mc></m:mcs></m:mPr>
            <m:mr>
                <m:e><m:r><m:t>a</m:t></m:r></m:e>
                <m:e><m:r><m:t>b</m:t></m:r></m:e>
            </m:mr>
            <m:mr>
                <m:e><m:r><m:t>c</m:t></m:r></m:e>
                <m:e><m:r><m:t>d</m:t></m:r></m:e>
            </m:mr>
        </m:m></m:oMath>"#;
        let block = parse(xml);
        let exprs = match block { MathBlock::Inline(v) => v, _ => panic!() };
        match &exprs[0] {
            MathExpr::Matrix { rows, cols, .. } => {
                assert_eq!(*cols, 2);
                assert_eq!(rows.len(), 2);
                assert_eq!(rows[0].len(), 2);
                assert!(matches!(&rows[0][0], MathExpr::Text(s) if s == "a"));
                assert!(matches!(&rows[1][1], MathExpr::Text(s) if s == "d"));
            }
            _ => panic!("expected Matrix"),
        }
    }

    #[test]
    fn parse_delimiter_square_brackets() {
        use crate::ir::MathExpr;
        let xml = r#"<m:oMath xmlns:m="x"><m:d>
            <m:dPr><m:begChr m:val="["/><m:endChr m:val="]"/></m:dPr>
            <m:e><m:r><m:t>x</m:t></m:r></m:e>
        </m:d></m:oMath>"#;
        let block = parse(xml);
        let exprs = match block { MathBlock::Inline(v) => v, _ => panic!() };
        match &exprs[0] {
            MathExpr::Delimiter { beg, end, content, .. } => {
                assert_eq!(*beg, '[');
                assert_eq!(*end, ']');
                assert!(matches!(**content, MathExpr::Text(ref s) if s == "x"));
            }
            _ => panic!("expected Delimiter"),
        }
    }

    #[test]
    fn parse_nary_sum() {
        use crate::ir::MathExpr;
        let xml = r#"<m:oMath xmlns:m="x"><m:nary>
            <m:naryPr>
                <m:chr m:val="∑"/>
                <m:limLoc m:val="undOvr"/>
                <m:grow m:val="1"/>
            </m:naryPr>
            <m:sub><m:r><m:t>i=1</m:t></m:r></m:sub>
            <m:sup><m:r><m:t>n</m:t></m:r></m:sup>
            <m:e><m:r><m:t>i</m:t></m:r></m:e>
        </m:nary></m:oMath>"#;
        let block = parse(xml);
        let exprs = match block { MathBlock::Inline(v) => v, _ => panic!() };
        match &exprs[0] {
            MathExpr::Nary { op, grow, operand, .. } => {
                assert_eq!(*op, '∑');
                assert!(*grow);
                assert!(matches!(**operand, MathExpr::Text(ref s) if s == "i"));
            }
            _ => panic!("expected Nary"),
        }
    }

    #[test]
    fn parse_fraction_nobar() {
        let xml = r#"<m:oMath xmlns:m="x"><m:f>
            <m:fPr><m:type m:val="noBar"/></m:fPr>
            <m:num><m:r><m:t>a</m:t></m:r></m:num>
            <m:den><m:r><m:t>b</m:t></m:r></m:den>
        </m:f></m:oMath>"#;
        let block = parse(xml);
        let exprs = match block { MathBlock::Inline(v) => v, _ => panic!() };
        match &exprs[0] {
            MathExpr::Fraction { bar_type, .. } => assert_eq!(*bar_type, FracBarType::NoBar),
            _ => panic!(),
        }
    }
}
