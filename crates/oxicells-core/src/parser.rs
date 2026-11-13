// This Source Code Form is subject to the terms of the Mozilla Public
// License, v. 2.0. If a copy of the MPL was not distributed with this
// file, You can obtain one at https://mozilla.org/MPL/2.0/.

use std::collections::HashMap;

use quick_xml::events::Event;
use quick_xml::reader::Reader;
use thiserror::Error;

use oxidocs_common::archive::OoxmlArchive;
use oxidocs_common::relationships::{parse_relationships, Relationship};

fn parse_header_footer(xml: &str) -> Result<Option<crate::ir::HeaderFooter>, XlsxError> {
    let mut reader = Reader::from_str(xml);
    reader.config_mut().trim_text(false);
    let mut footer = None;
    let mut field = None;
    let mut text = String::new();
    loop {
        let event = reader.read_event().map_err(XlsxError::Xml)?;
        let empty = matches!(event, Event::Empty(_));
        match event {
            Event::Start(start) | Event::Empty(start) if start.local_name().as_ref() == b"headerFooter" => {
                footer = Some(crate::ir::HeaderFooter {
                    different_odd_even: is_true(get_attr(&start, "differentOddEven").as_deref()),
                    different_first: is_true(get_attr(&start, "differentFirst").as_deref()),
                    scale_with_doc: get_attr(&start, "scaleWithDoc").as_deref().is_none_or(|s| is_true(Some(s))),
                    align_with_margins: get_attr(&start, "alignWithMargins").as_deref().is_none_or(|s| is_true(Some(s))),
                    ..Default::default()
                });
                if empty { break; }
            }
            Event::Start(start) if footer.is_some() => {
                field = match start.local_name().as_ref() {
                    b"oddHeader" => Some(0), b"oddFooter" => Some(1),
                    b"evenHeader" => Some(2), b"evenFooter" => Some(3),
                    b"firstHeader" => Some(4), b"firstFooter" => Some(5), _ => None,
                };
                if field.is_some() { text.clear(); }
            }
            Event::Text(value) if field.is_some() => text.push_str(&value.unescape()?),
            Event::CData(value) if field.is_some() => text.push_str(std::str::from_utf8(value.as_ref())
                .map_err(|e| XlsxError::InvalidData(e.to_string()))?),
            Event::End(end) if end.local_name().as_ref() == b"headerFooter" => break,
            Event::End(_) if field.is_some() => {
                let f = footer.as_mut().expect("header/footer is open");
                let value = match field.take().unwrap() { 0 => &mut f.odd_header, 1 => &mut f.odd_footer,
                    2 => &mut f.even_header, 3 => &mut f.even_footer,
                    4 => &mut f.first_header, _ => &mut f.first_footer };
                *value = crate::string_value::decode(&text);
            }
            Event::Eof => break,
            _ => {}
        }
    }
    Ok(footer)
}

/// Turns an OOXML custom-filter operator into the prefix VBA states it with.
fn filter_operator(operator: &str) -> &'static str {
    match operator {
        "greaterThan" => ">",
        "greaterThanOrEqual" => ">=",
        "lessThan" => "<",
        "lessThanOrEqual" => "<=",
        "notEqual" => "<>",
        _ => "",
    }
}

/// Records a worksheet feature this parser walks past, so a caller can say
/// what it could not show. The part itself is left alone when the workbook is
/// saved, so nothing is lost 遯ｶ繝ｻit simply does not reach the IR.
fn note_unsupported(name: &str, noted: &mut Vec<String>) {
    let feature = match name {
        "conditionalFormatting" => "Conditional formatting",
        "dataValidation" | "dataValidations" => "Data validation",
        "hyperlink" | "hyperlinks" => "Hyperlinks",
        "pane" => "Split panes",
        "drawing" => "Drawings",
        "legacyDrawing" => "Comments",
        "tableParts" => "Tables",
        "pivotSelection" => "Pivot tables",
        _ => return,
    };
    if !noted.iter().any(|held| held == feature) {
        noted.push(feature.to_string());
    }
}

/// OOXML writes a boolean attribute as `1`/`0` or `true`/`false`.
fn is_true(value: Option<&str>) -> bool {
    matches!(value, Some("1") | Some("true"))
}

fn sheet_protection_from(start: &quick_xml::events::BytesStart<'_>) -> crate::ir::SheetProtection {
    let allowed = |name| !get_attr(start, name).as_deref().is_none_or(|value| is_true(Some(value)));
    let password = crate::ir::PasswordVerifier {
        legacy_hash: get_attr(start, "password").and_then(|value| u16::from_str_radix(&value, 16).ok()),
        algorithm: get_attr(start, "algorithmName"),
        salt: get_attr(start, "saltValue"),
        hash: get_attr(start, "hashValue"),
        iterations: get_attr(start, "spinCount").and_then(|value| value.parse().ok()),
    };
    let has_password = password.legacy_hash.is_some() || password.algorithm.is_some()
        || password.salt.is_some() || password.hash.is_some() || password.iterations.is_some();
    crate::ir::SheetProtection {
        drawing_objects: is_true(get_attr(start, "objects").as_deref()),
        scenarios: is_true(get_attr(start, "scenarios").as_deref()),
        allow_formatting_cells: allowed("formatCells"),
        allow_formatting_columns: allowed("formatColumns"),
        allow_formatting_rows: allowed("formatRows"),
        allow_inserting_columns: allowed("insertColumns"),
        allow_inserting_rows: allowed("insertRows"),
        allow_inserting_hyperlinks: allowed("insertHyperlinks"),
        allow_deleting_columns: allowed("deleteColumns"),
        allow_deleting_rows: allowed("deleteRows"),
        allow_sorting: allowed("sort"),
        allow_filtering: allowed("autoFilter"),
        allow_using_pivot_tables: allowed("pivotTables"),
        password: has_password.then_some(password),
    }
}



use oxidocs_common::xml_utils::{get_attr, local_name};

use crate::ir::{
    BorderLine, Cell, CellStyle, CellValue, ExternalBook, ExternalSheet, MergeCell, Row, Sheet,
    Workbook,
};

#[derive(Error, Debug)]
pub enum XlsxError {
    #[error("Archive error: {0}")]
    Archive(#[from] oxidocs_common::OxiError),

    #[error("XML error: {0}")]
    Xml(#[from] quick_xml::Error),

    #[error("Invalid cell reference: {0}")]
    InvalidCellRef(String),

    #[error("Invalid data: {0}")]
    InvalidData(String),
}

/// Parse a cell reference like "A1" into (col, row) as 0-based indices.
/// "A1" -> (0, 0), "B2" -> (1, 1), "AA1" -> (26, 0), "AZ3" -> (51, 2)
pub fn parse_cell_ref(s: &str) -> (u32, u32) {
    let mut col: u32 = 0;
    let mut row_str = String::new();
    let mut found_digit = false;

    for ch in s.chars() {
        if ch.is_ascii_alphabetic() && !found_digit {
            col = col * 26 + (ch.to_ascii_uppercase() as u32 - b'A' as u32 + 1);
        } else {
            found_digit = true;
            row_str.push(ch);
        }
    }

    let col = if col > 0 { col - 1 } else { 0 };
    let row = row_str.parse::<u32>().unwrap_or(1).saturating_sub(1);

    (col, row)
}

/// Parse a range reference like "A1:C3" into (start_col, start_row, end_col, end_row).
/// Columns are 0-based, rows are 1-based.
fn parse_range_ref(s: &str) -> Option<(u32, u32, u32, u32)> {
    let parts: Vec<&str> = s.split(':').collect();
    if parts.len() != 2 {
        return None;
    }
    let (start_col, start_row_0) = parse_cell_ref(parts[0]);
    let (end_col, end_row_0) = parse_cell_ref(parts[1]);
    // Convert to 1-based rows for MergeCell
    Some((start_col, start_row_0 + 1, end_col, end_row_0 + 1))
}

/// Parse the shared strings table (xl/sharedStrings.xml).
/// Returns a Vec of strings indexed by position.
/// A shared string, and how its parts are dressed when they are not all dressed
/// alike. `text` always holds the whole of it.
#[derive(Debug, Clone, Default)]
struct SharedString {
    text: String,
    units: Option<Vec<u16>>,
    runs: Vec<crate::ir::TextRun>,
}

/// Text as the reader sees it, with the `_xHHHH_` escapes spelled out.
///
/// A character a writer cannot put in XML 遯ｶ繝ｻa carriage return, a control code 遯ｶ繝ｻ
/// is written as its code point between `_x` and `_`, and a literal `_x` that
/// would otherwise begin one of those is itself written `_x005F_`
/// (ECMA-376 ・ゑｽｧ22.9.2.19). Left to stand, the escape is drawn:
/// `procurement_contractor_list_02` shows 隴ｬ・ｪ陟台ｸ茨ｽｼ螟ゑｽ､・ｾ繝ｻ・ｮ繝ｻ・ｳ繝ｻ・､_x000D_ where Excel shows the
/// name alone, and the seven strings it carries make the widest black band in
/// that book's picture.
///
/// One pass, left to right, is what keeps `_x005F_` honest: it becomes an
/// underscore, and the `x000D_` behind it is then ordinary text rather than an
/// escape read a second time.
#[cfg(test)]
fn unescape_wide(text: &str) -> String {
    if !text.contains("_x") {
        return text.to_string();
    }
    let letters: Vec<char> = text.chars().collect();
    let mut out = String::with_capacity(text.len());
    let mut at = 0;
    while at < letters.len() {
        let escape = letters[at] == '_'
            && letters.get(at + 1) == Some(&'x')
            && letters.get(at + 6) == Some(&'_')
            && letters[at + 2..at + 6].iter().all(char::is_ascii_hexdigit);
        if escape {
            let code: String = letters[at + 2..at + 6].iter().collect();
            if let Some(held) = u32::from_str_radix(&code, 16).ok().and_then(char::from_u32) {
                out.push(held);
                at += 7;
                continue;
            }
        }
        out.push(letters[at]);
        at += 1;
    }
    out
}

fn parse_shared_strings(xml: &str) -> Result<Vec<SharedString>, XlsxError> {
    let mut reader = Reader::from_str(xml);
    let mut strings: Vec<SharedString> = Vec::new();
    let mut current = SharedString::default();
    let mut in_si = false;
    let mut in_t = false;
    // A phonetic guide (furigana) is stored as an <rPh> element containing its
    // own <t>. It is not part of the cell's text: Excel shows "陋ｹ・ｺ陋ｻ繝ｻ, not
    // "陋ｹ・ｺ陋ｻ繝ｻ縺醍ｹ晄じﾎｦ". Japanese workbooks carry these on names and addresses
    // constantly, so failing to skip them corrupts a large share of real files.
    let mut in_phonetic = false;
    // A string can be built from <r> runs, each optionally carrying its own
    // <rPr>: an 8pt aside inside an 11pt cell, a raised footnote marker. The
    // run's text goes into the whole string as well, so a reader that ignores
    // the dressing still sees everything.
    let mut in_run = false;
    let mut in_run_props = false;
    let mut current_run = crate::ir::TextRun::default();
    let mut text_part = String::new();

    loop {
        match reader.read_event()? {
            Event::Start(e) | Event::Empty(e) => {
                let name = local_name(e.name().as_ref());
                match name.as_str() {
                    "si" => {
                        in_si = true;
                        in_phonetic = false;
                        in_run = false;
                        current = SharedString::default();
                        current.units = Some(Vec::new());
                    }
                    "rPh" => in_phonetic = true,
                    "r" if in_si && !in_phonetic => {
                        in_run = true;
                        current_run = crate::ir::TextRun::default();
                    }
                    "rPr" if in_run => {
                        in_run_props = true;
                        current_run.dressed = true;
                    }
                    "sz" if in_run_props => {
                        current_run.size = get_attr(&e, "val").and_then(|v| v.parse().ok());
                    }
                    "rFont" | "font" if in_run_props => {
                        current_run.font = get_attr(&e, "val");
                    }
                    "b" if in_run_props => {
                        current_run.bold = get_attr(&e, "val").as_deref() != Some("0");
                    }
                    "i" if in_run_props => {
                        current_run.italic = get_attr(&e, "val").as_deref() != Some("0");
                    }
                    "strike" if in_run_props => {
                        current_run.strikethrough = is_true(get_attr(&e, "val").as_deref().or(Some("1")));
                    }
                    "u" if in_run_props => {
                        use crate::ir::CellUnderlineKind;
                        let value=get_attr(&e,"val");
                        current_run.underline = !matches!(value.as_deref(),Some("0")|Some("none"));
                        current_run.underline_kind=match value.as_deref() {
                            Some("double")=>Some(CellUnderlineKind::Double),
                            Some("singleAccounting")=>Some(CellUnderlineKind::SingleAccounting),
                            Some("doubleAccounting")=>Some(CellUnderlineKind::DoubleAccounting),
                            _=>None,
                        };
                    }
                    "vertAlign" if in_run_props => {
                        current_run.vert_align = get_attr(&e, "val");
                    }
                    "family" if in_run_props => {
                        current_run.font_family=get_attr(&e,"val").and_then(|v|v.parse().ok());
                    }
                    "charset" if in_run_props => {
                        current_run.font_charset=get_attr(&e,"val").and_then(|v|v.parse().ok());
                    }
                    "scheme" if in_run_props => {
                        current_run.font_reference.get_or_insert_with(Default::default).scheme=
                            get_attr(&e,"val").filter(|v|v!="none");
                    }
                    "color" if in_run_props => {
                        let reference=current_run.font_reference.get_or_insert_with(Default::default);
                        reference.rgb=get_attr(&e,"rgb");
                        reference.implicit_automatic=false;
                        reference.theme=get_attr(&e,"theme").and_then(|v|v.parse().ok());
                        reference.indexed=get_attr(&e,"indexed").and_then(|v|v.parse().ok());
                        reference.automatic=is_true(get_attr(&e,"auto").as_deref());
                        reference.tint=get_attr(&e,"tint").and_then(|v|v.parse::<f64>().ok()).filter(|v|v.is_finite());
                        current_run.color = get_attr(&e, "rgb").filter(|rgb| matches!(rgb.len(),6|8)&&rgb.bytes().all(|byte|byte.is_ascii_hexdigit())).map(|rgb| {
                            let rgb = rgb.trim_start_matches('#');
                            if rgb.len()==8 { rgb[2..].to_ascii_uppercase() } else { rgb.to_ascii_uppercase() }
                        });
                    }
                    "t" if in_si => { in_t = true; text_part.clear(); },
                    _ => {}
                }
            }
            Event::End(e) => {
                let name = local_name(e.name().as_ref());
                match name.as_str() {
                    "si" => {
                        in_si = false;
                        // A string whose runs are all dressed the same as the
                        // cell is no richer than a plain one.
                        if current.runs.iter().all(|run| {
                            !run.dressed
                                && run.size.is_none()
                                && run.font.is_none()
                                && run.vert_align.is_none()
                                && !run.bold
                                && !run.italic
                                && !run.underline
                                && !run.strikethrough
                                && run.color.is_none()
                        }) {
                            current.runs.clear();
                        }
                        strings.push(std::mem::take(&mut current));
                    }
                    "rPh" => in_phonetic = false,
                    "rPr" => in_run_props = false,
                    "r" if in_run => {
                        in_run = false;
                        if current_run.font_reference.as_deref().is_some_and(crate::ir::FontReference::is_empty) {
                            current_run.font_reference=None;
                        }
                        if !current_run.text.is_empty() {
                            current.runs.push(std::mem::take(&mut current_run));
                        }
                    }
                    "t" => {
                        if in_t && in_si && !in_phonetic {
                            let units = crate::string_value::decode(&text_part);
                            current.text.push_str(&text_part);
                            current.units.as_mut().unwrap().extend_from_slice(&units);
                            if in_run {
                                current_run.text.push_str(&String::from_utf16_lossy(&units).replace("\r\n", "\n"));
                            }
                        }
                        in_t = false;
                    },
                    _ => {}
                }
            }
            Event::Text(e) => {
                if in_t && in_si && !in_phonetic {
                    // A break inside a cell is written as a pair of
                    // characters by some writers and one by others; the IR
                    // holds one, so the pieces of a string add up to the
                    // string. They did not before: `001290291`'s title came
                    // out 85 characters in its five pieces against 62 in the
                    // cell, and every one of the 23 was the other half of a
                    // break.
                    text_part.push_str(&e.unescape()?);
                }
            }
            Event::Eof => break,
            _ => {}
        }
    }

    Ok(strings)
}

/// Information about a sheet from workbook.xml
struct SheetInfo {
    name: String,
    visibility: crate::ir::Visibility,
    r_id: String,
}

fn parse_calculation(xml: &str) -> crate::ir::CalculationSettings {
    use crate::ir::{CalculationMode, CalculationSettings};
    let compatibility_version = crate::compatibility::read_version(xml);
    let mut reader = Reader::from_str(xml);
    loop {
        match reader.read_event() {
            Ok(Event::Start(e) | Event::Empty(e)) if local_name(e.name().as_ref()) == "calcPr" => {
                let mode = match get_attr(&e, "calcMode").as_deref() {
                    Some("manual") => CalculationMode::Manual,
                    Some("autoNoTable") => CalculationMode::AutomaticExceptTables,
                    _ => CalculationMode::Automatic,
                };
                let boolean = |name| match get_attr(&e, name).as_deref() {
                    Some("0" | "false") => Some(false),
                    Some("1" | "true") => Some(true),
                    _ => None,
                };
                return CalculationSettings { mode, on_save: boolean("calcOnSave").unwrap_or(true), completed: boolean("calcCompleted"), compatibility_version, dbcs_language: None };
            }
            Ok(Event::Eof) | Err(_) => return CalculationSettings { compatibility_version, ..Default::default() },
            _ => {}
        }
    }
}

/// Saved array anchors Excel queues when opening in manual calculation mode.
/// The array's aca flag is independent of calcCompleted and its cached value.
fn pending_array_calculations(xml: &str) -> Result<Vec<(u32, u32)>, XlsxError> {
    if !xml.contains("aca") { return Ok(Vec::new()); }
    let mut reader = Reader::from_str(xml);
    let mut path = Vec::<String>::new();
    let mut active = None;
    let mut pending = Vec::new();
    loop {
        let event = reader.read_event().map_err(XlsxError::Xml)?;
        let empty = matches!(&event, Event::Empty(_));
        match event {
            Event::Start(e) | Event::Empty(e) => {
                let name = local_name(e.name().as_ref());
                if name == "c" && path.last().map(String::as_str) == Some("row")
                    && path.iter().rev().nth(1).map(String::as_str) == Some("sheetData") {
                    active = get_attr(&e, "r").and_then(|address| oxicells_calc::reference::parse_a1(&address))
                        .map(|cell| (cell.col, cell.row));
                } else if name == "f" && path.last().map(String::as_str) == Some("c")
                    && get_attr(&e, "t").as_deref() == Some("array")
                    && is_true(get_attr(&e, "aca").as_deref()) {
                    if let Some(cell) = active { pending.push(cell); }
                }
                if !empty { path.push(name); } else if name == "c" { active = None; }
            },
            Event::End(e) => {
                if local_name(e.name().as_ref()) == "c" { active = None; }
                path.pop();
            },
            Event::Eof => break,
            _ => {},
        }
    }
    Ok(pending)
}

/// Workbook and sheet-local names, with their original per-name attributes.
fn parse_defined_names(xml: &str) -> (Vec<(String, String)>, crate::ir::DefinedNameAttributes) {
    let sheets = parse_workbook_sheets(xml).unwrap_or_default();
    let mut reader = Reader::from_str(xml);
    let mut names = Vec::new();
    let mut attributes = std::collections::BTreeMap::new();
    let mut wanted: Option<String> = None;
    let mut refers = String::new();
    loop {
        match reader.read_event() {
            Ok(event @ (Event::Start(_) | Event::Empty(_))) if matches!(&event, Event::Start(e) | Event::Empty(e) if local_name(e.name().as_ref()) == "definedName") => {
                let empty = matches!(&event, Event::Empty(_));
                let (Event::Start(e) | Event::Empty(e)) = event else { unreachable!() };
                let held = e.attributes().flatten().filter_map(|a| {
                    let key = std::str::from_utf8(a.key.as_ref()).ok()?.to_string();
                    let value = a.unescape_value().ok()?.into_owned();
                    Some((key, value))
                }).collect::<std::collections::BTreeMap<_, _>>();
                let name = held.get("name").cloned().unwrap_or_default();
                let leaf = name.strip_prefix("_xlnm.").unwrap_or(&name);
                wanted = if leaf.is_empty() { None } else if let Some(scope) = held.get("localSheetId") {
                    scope.parse::<usize>().ok().and_then(|index| sheets.get(index))
                        .map(|sheet| crate::ir::scoped_defined_name(&sheet.name, leaf))
                } else { Some(leaf.to_string()) };
                if let Some(name) = &wanted { attributes.insert(name.clone(), held); }
                refers.clear();
                if empty {
                    if let Some(name) = wanted.take() { names.push((name, String::new())); }
                }
            }
            Ok(Event::Text(e)) if wanted.is_some() => {
                if let Ok(text) = e.unescape() { refers.push_str(&text); }
            }
            Ok(Event::CData(e)) if wanted.is_some() => {
                refers.push_str(&String::from_utf8_lossy(e.as_ref()));
            }
            Ok(Event::End(e)) if local_name(e.name().as_ref()) == "definedName" => {
                if let Some(name) = wanted.take() { names.push((name, refers.trim().to_string())); }
            }
            Ok(Event::Eof) | Err(_) => break,
            _ => {}
        }
    }
    (names, attributes)
}

/// The relationship ids of the workbooks this one links to, in the order
/// `[1]`, `[2]` 遯ｶ・ｦ name them.
fn parse_table_part_references(xml: &str) -> Result<Option<Vec<String>>,XlsxError> {
    use quick_xml::name::ResolveResult;
    let mut reader=quick_xml::reader::NsReader::from_str(xml);
    let mut path=Vec::<(Vec<u8>,bool)>::new();let mut ids=Vec::new();let mut found=false;
    loop {
        let (namespace,event)=reader.read_resolved_event().map_err(XlsxError::Xml)?;
        let standard=matches!(namespace,ResolveResult::Bound(uri) if matches!(uri.as_ref(),
            b"http://schemas.openxmlformats.org/spreadsheetml/2006/main"|b"http://purl.oclc.org/ooxml/spreadsheetml/main"));
        match event {
            Event::Start(tag)|Event::Empty(tag)=>{
                let name=tag.local_name().as_ref().to_vec();
                if standard && name==b"tableParts" && path.len()==1 && path[0].0==b"worksheet" && path[0].1 {found=true;}
                if standard && name==b"tablePart" && path.len()==2
                    && path[0].0==b"worksheet" && path[0].1 && path[1].0==b"tableParts" && path[1].1 {
                    for attribute in tag.attributes() {
                        let attribute=attribute.map_err(|error|XlsxError::InvalidData(error.to_string()))?;
                        let (namespace,local)=reader.resolve_attribute(attribute.key);
                        if local.as_ref()==b"id" && matches!(namespace,ResolveResult::Bound(uri) if matches!(uri.as_ref(),
                            b"http://schemas.openxmlformats.org/officeDocument/2006/relationships"|b"http://purl.oclc.org/ooxml/officeDocument/relationships")) {
                            ids.push(attribute.decode_and_unescape_value(reader.decoder()).map_err(XlsxError::Xml)?.into_owned());
                        }
                    }
                }
                let end=oxidocs_common::xml_utils::xml_byte_position(xml,reader.buffer_position());
                if !xml[..end].ends_with("/>") {path.push((name,standard));}
            }
            Event::End(_)=>{path.pop();}
            Event::Eof=>break,
            _=>{},
        }
    }
    Ok(found.then_some(ids))
}

fn parse_external_references(xml: &str) -> Vec<String> {
    let mut reader = Reader::from_str(xml);
    let mut ids = Vec::new();
    loop {
        match reader.read_event() {
            Ok(Event::Start(e)) | Ok(Event::Empty(e))
                if local_name(e.name().as_ref()) == "externalReference" =>
            {
                // The id attribute is namespaced `r:id`; `get_attr` matches on
                // the local name, so both spellings are found.
                if let Some(id) = get_attr(&e, "id") {
                    ids.push(id);
                }
            }
            Ok(Event::Eof) | Err(_) => break,
            _ => {}
        }
    }
    ids
}

/// One `xl/externalLinks/externalLinkN.xml`: the sheets the linked workbook
/// has, and whatever this file remembers of their cells.
///
/// A cached cell says nothing about its type when it held a number, `t="str"`
/// when it held text and `t="b"` for a Boolean 遯ｶ繝ｻthe same shorthand a
/// worksheet uses, minus the shared-string table, since the text is written
/// out here rather than pooled.
fn parse_external_link_with_identity(archive: &mut OoxmlArchive, part: &str, xml: &str) -> Result<ExternalBook,XlsxError> {
    use quick_xml::name::ResolveResult;
    let mut book=parse_external_link(xml);book.origin_part=Some(part.to_owned());
    let (parent,leaf)=part.rsplit_once('/').unwrap_or(("",part));
    let rels=if parent.is_empty(){format!("_rels/{leaf}.rels")}else{format!("{parent}/_rels/{leaf}.rels")};
    let Some(rels)=archive.try_read_part(&rels)? else {return Ok(book)};
    let relationships=oxidocs_common::relationships::parse_relationships(&rels)?;
    let mut primary=None;let mut absolute=None;
    let mut reader=quick_xml::reader::NsReader::from_str(xml);
    loop {
        let event=reader.read_event().map_err(XlsxError::Xml)?.into_owned();
        match event {
            Event::Start(element)|Event::Empty(element)=>{
                let (namespace,local)=reader.resolve_element(element.name());
                let standard=matches!(namespace,ResolveResult::Bound(uri)if matches!(uri.as_ref(),b"http://schemas.openxmlformats.org/spreadsheetml/2006/main"|b"http://purl.oclc.org/ooxml/spreadsheetml/main"));
                let alternate=matches!(namespace,ResolveResult::Bound(uri)if uri.as_ref()==b"http://schemas.microsoft.com/office/spreadsheetml/2021/extlinks2021");
                let destination=if standard && local.as_ref()==b"externalBook"{&mut primary}
                    else if alternate && local.as_ref()==b"absoluteUrl"{&mut absolute}else{continue};
                for attribute in element.attributes() {
                    let attribute=attribute.map_err(|error|XlsxError::InvalidData(error.to_string()))?;
                    let (namespace,local)=reader.resolve_attribute(attribute.key);
                    let relationship=matches!(namespace,ResolveResult::Bound(uri)if matches!(uri.as_ref(),b"http://schemas.openxmlformats.org/officeDocument/2006/relationships"|b"http://purl.oclc.org/ooxml/officeDocument/relationships"));
                    if relationship && local.as_ref()==b"id" {
                        if destination.is_some(){return Err(XlsxError::InvalidData("duplicate external-book relationship identity".into()))}
                        *destination=Some(attribute.decode_and_unescape_value(reader.decoder()).map_err(XlsxError::Xml)?.into_owned());
                    }
                }
            },
            Event::Eof=>break,
            _=>{},
        }
    }
    let target=|id:Option<String>|id.and_then(|id|relationships.get(&id)).filter(|relationship|relationship.rel_type.ends_with("/externalLinkPath") && relationship.target_mode.as_deref()==Some("External")).map(|relationship|relationship.target.clone());
    book.target=target(primary);book.absolute_target=target(absolute);Ok(book)
}

fn parse_external_link(xml: &str) -> ExternalBook {
    let mut reader = Reader::from_str(xml);
    let mut sheets: Vec<ExternalSheet> = Vec::new();
    let mut at_sheet: Option<usize> = None;
    let mut cell: Option<(u32, u32, String)> = None;
    let mut text = String::new();
    loop {
        match reader.read_event() {
            Ok(Event::Start(e)) | Ok(Event::Empty(e))
                if local_name(e.name().as_ref()) == "sheetName" =>
            {
                sheets.push(ExternalSheet {
                    name: get_attr(&e, "val").unwrap_or_default(),
                    cells: Vec::new(),
                });
            }
            Ok(Event::Start(e)) | Ok(Event::Empty(e))
                if local_name(e.name().as_ref()) == "sheetData" =>
            {
                at_sheet = get_attr(&e, "sheetId").and_then(|id| id.parse::<usize>().ok());
            }
            Ok(Event::Start(e)) if local_name(e.name().as_ref()) == "cell" => {
                let kind = get_attr(&e, "t").unwrap_or_default();
                // `parse_cell_ref` answers (COLUMN, row) and counts both from
                // zero; the IR counts a row from one.
                cell = get_attr(&e, "r")
                    .map(|reference| parse_cell_ref(&reference))
                    .map(|(column, row)| (row + 1, column, kind));
                text.clear();
            }
            Ok(Event::Text(e)) if cell.is_some() => {
                if let Ok(read) = e.unescape() {
                    text.push_str(&read);
                }
            }
            Ok(Event::End(e)) if local_name(e.name().as_ref()) == "cell" => {
                if let (Some((row, column, kind)), Some(at)) = (cell.take(), at_sheet) {
                    if let Some(sheet) = sheets.get_mut(at) {
                        sheet.cells.push((row, column, cached_cell_value(&kind, &text)));
                    }
                }
            }
            Ok(Event::Eof) | Err(_) => break,
            _ => {}
        }
    }
    ExternalBook { sheets, ..Default::default() }
}

/// What a cached cell held, from the shorthand the link part writes.
fn cached_cell_value(kind: &str, text: &str) -> CellValue {
    match kind {
        "str" | "s" => CellValue::from_utf16(crate::string_value::decode(text)),
        "b" => CellValue::Boolean(text != "0"),
        "e" => CellValue::Error(text.to_string()),
        _ => match text.parse::<f64>() {
            Ok(number) => CellValue::Number(number),
            Err(_) if text.is_empty() => CellValue::Empty,
            Err(_) => CellValue::String(text.to_string()),
        },
    }
}

/// Parse workbook.xml to extract sheet names and their relationship IDs.
fn parse_workbook_sheets(xml: &str) -> Result<Vec<SheetInfo>, XlsxError> {
    let mut reader = Reader::from_str(xml);
    let mut sheets = Vec::new();

    loop {
        match reader.read_event()? {
            Event::Start(e) | Event::Empty(e) => {
                let name = local_name(e.name().as_ref());
                if name == "sheet" {
                    let sheet_name = get_attr(&e, "name").unwrap_or_default();
                    // r:id attribute 遯ｶ繝ｻtry both namespaced and raw forms
                    let r_id = get_attr(&e, "id")
                        .or_else(|| {
                            // Try raw attribute key "r:id"
                            for attr in e.attributes().flatten() {
                                let key =
                                    std::str::from_utf8(attr.key.as_ref()).unwrap_or("");
                                if key == "r:id" {
                                    return Some(
                                        String::from_utf8_lossy(&attr.value).to_string(),
                                    );
                                }
                            }
                            None
                        })
                        .unwrap_or_default();

                    sheets.push(SheetInfo {
                        name: sheet_name,
                        visibility: crate::ir::Visibility::of(
                            get_attr(&e, "state").unwrap_or_default().as_str(),
                        ),
                        r_id,
                    });
                }
            }
            Event::Eof => break,
            _ => {}
        }
    }

    Ok(sheets)
}

// =====================================================================
// styles.xml parsing
// =====================================================================

#[derive(Debug, Clone, Default)]
struct FontInfo {
    bold: bool,
    italic: bool,
    underline: bool,
    underline_kind: Option<crate::ir::CellUnderlineKind>,
    font_position: Option<crate::ir::CellFontPosition>,
    strikethrough: bool,
    size: Option<f32>,
    color: Option<String>,
    name: Option<String>,
    latin_name: Option<String>,
    reference: crate::ir::FontReference,
    charset: Option<i32>,
    family: Option<i32>,
}

#[derive(Debug, Clone, Default)]
struct FillInfo {
    bg_color: Option<String>,
    pattern_fill: crate::ir::CellFill,
}

#[derive(Debug, Clone, Default)]
struct BorderInfo {
    left: Option<BorderLine>,
    right: Option<BorderLine>,
    top: Option<BorderLine>,
    bottom: Option<BorderLine>,
    diagonal: Option<BorderLine>,
    up: bool,
    down: bool,
}

fn xf_format_reference(element: &quick_xml::events::BytesStart<'_>) -> crate::ir::CellFormatReference {
    let flag=|name| get_attr(element,name).map(|value|matches!(value.as_str(),"1"|"true"));
    crate::ir::CellFormatReference {
        style_id:get_attr(element,"xfId").and_then(|value|value.parse().ok()),
        number_format:flag("applyNumberFormat"),font:flag("applyFont"),
        fill:flag("applyFill"),border:flag("applyBorder"),
        alignment:flag("applyAlignment"),protection:flag("applyProtection"),
    }
}

#[derive(Debug, Clone, Default)]
struct XfRecord {
    format_reference: crate::ir::CellFormatReference,
    num_fmt_id: u32,
    font_id: usize,
    fill_id: usize,
    border_id: usize,
    horizontal_align: Option<String>,
    vertical_align: Option<String>,
    indent: u32,
    wrap_text: bool,
    stacked_text: bool,
    text_rotation: Option<i32>,
    justify_last_line: bool,
    shrink_to_fit: bool,
    reading_order: crate::ir::CellReadingOrder,
    /// `quotePrefix="1"`: the entry was typed behind an apostrophe.
    quote_prefix: bool,
    /// Whether the cell may still be edited once its sheet is protected.
    /// Excel locks every cell by default and unlocks the few a form wants
    /// typed into, so the negative is the one worth holding: it keeps this
    /// struct right when it is built from its own default.
    unlocked: bool,
    formula_hidden: bool,
}

#[derive(Debug, Clone, Default)]
struct StyleSheet {
    num_fmts: HashMap<u32, String>,
    fonts: Vec<FontInfo>,
    fills: Vec<FillInfo>,
    borders: Vec<BorderInfo>,
    /// The borders the differential formats state, in the order they are
    /// written. A table names the rule it draws round itself by an index into
    /// these.
    dxf_borders: Vec<BorderInfo>,
    /// The whole of each differential format 遯ｶ繝ｻcolour, weight, fill 遯ｶ繝ｻin the
    /// order a `dxfId` counts them. A conditional rule names one of these.
    dxfs: Vec<crate::ir::DiffStyle>,
    cell_xfs: Vec<XfRecord>,
    cell_style_xfs: Vec<XfRecord>,
    named_styles: Vec<(String,usize,Option<u32>)>,
}

/// Built-in number format strings for well-known IDs.
/// The formats OOXML gives a number rather than spelling them out, from
/// ECMA-376 ・ゑｽｧ18.8.30. A workbook that uses one of these writes only its id:
/// the accounting formats 37 to 40 are how a Japanese statistical table asks
/// for thousands separators, and reading them as "General" prints 24493 where
/// Excel shows 24,493.
fn builtin_number_format(id: u32) -> Option<&'static str> {
    match id {
        0 => Some("General"),
        1 => Some("0"),
        2 => Some("0.00"),
        3 => Some("#,##0"),
        4 => Some("#,##0.00"),
        5 => Some("$#,##0_);($#,##0)"),
        6 => Some("$#,##0_);[Red]($#,##0)"),
        7 => Some("$#,##0.00_);($#,##0.00)"),
        8 => Some("$#,##0.00_);[Red]($#,##0.00)"),
        9 => Some("0%"),
        10 => Some("0.00%"),
        11 => Some("0.00E+00"),
        12 => Some("# ?/?"),
        13 => Some("# ??/??"),
        14 => Some("mm-dd-yy"),
        15 => Some("d-mmm-yy"),
        16 => Some("d-mmm"),
        17 => Some("mmm-yy"),
        18 => Some("h:mm AM/PM"),
        19 => Some("h:mm:ss AM/PM"),
        20 => Some("h:mm"),
        21 => Some("h:mm:ss"),
        22 => Some("m/d/yy h:mm"),
        // Excel's own `NumberFormat` for these four, asked of a cell that
        // wears one: the room after the number is the width of a bracket,
        // not of a space, which is two pixels in 繝ｻ・ｭ繝ｻ・ｳ 11 and shows up as a
        // right-aligned number sitting two pixels off.
        37 => Some("#,##0_);(#,##0)"),
        38 => Some("#,##0_);[Red](#,##0)"),
        39 => Some("#,##0.00_);(#,##0.00)"),
        40 => Some("#,##0.00_);[Red](#,##0.00)"),
        // The Japanese locale fills in 27 to 36 and 50 to 58, which is what a
        // government workbook's dates are written with. The era forms 遯ｶ繝ｻggge,
        // 闔会ｽ､陷･繝ｻ遯ｶ繝ｻare given their Gregorian equivalent here, because the
        // formatter has no calendar to name an era from; a date in the right
        // shape is nearer Excel's ink than the serial number that reading
        // none of them at all leaves behind.
        27 | 36 | 50 | 57 => Some("yyyy\".\"m\".\"d"),
        28 | 29 | 51 | 54 | 58 => Some("yyyy\"\u{5e74}\"m\"\u{6708}\"d\"\u{65e5}\""),
        30 => Some("m/d/yy"),
        31 => Some("yyyy\"\u{5e74}\"m\"\u{6708}\"d\"\u{65e5}\""),
        32 => Some("h\"\u{6642}\"mm\"\u{5206}\""),
        33 => Some("h\"\u{6642}\"mm\"\u{5206}\"ss\"\u{79d2}\""),
        34 | 52 | 55 => Some("yyyy\"\u{5e74}\"m\"\u{6708}\""),
        35 | 53 | 56 => Some("m\"\u{6708}\"d\"\u{65e5}\""),
        45 => Some("mm:ss"),
        46 => Some("[h]:mm:ss"),
        47 => Some("mmss.0"),
        48 => Some("##0.0E+0"),
        49 => Some("@"),
        _ => None,
    }
}

/// The colours a workbook's theme names, in the order the theme states them:
/// dk1, lt1, dk2, lt2, accent1-6, hlink, folHlink.
#[derive(Debug, Clone)]
pub(crate) struct Theme {
    colours: Vec<String>,
    /// The palette the workbook states in place of Excel's own, if it does.
    /// It lives in the styles part rather than the theme, and is read before
    /// the styles that use it.
    indexed: Vec<String>,
    /// The faces the theme's major and minor schemes name for this script.
    /// A font that says `<scheme val="minor"/>` wears one of these and its
    /// own `<name>` counts for nothing.
    major_face: Option<String>,
    minor_face: Option<String>,
    major_latin_face: Option<String>,
    minor_latin_face: Option<String>,
    /// The widths the theme's own line styles are drawn at, in EMU. A shape
    /// that names one through `lnRef` and states no width of its own is ruled
    /// at the width the theme gives that slot.
    rules: Vec<i64>,
}

impl Default for Theme {
    fn default() -> Self {
        // Excel uses its built-in Office theme when an otherwise valid
        // workbook has no theme part. Keep the DrawingML slots available so
        // theme-based shape text still resolves on load and after SaveCopyAs.
        Self {
            colours: [
                "000000", "FFFFFF", "0E2841", "E8E8E8", "156082", "E97132",
                "196B24", "0F9ED5", "A02B93", "4EA72E", "467886", "96607D",
            ].into_iter().map(str::to_string).collect(),
            indexed: Vec::new(),
            major_face: None,
            minor_face: None,
            major_latin_face: None,
            minor_latin_face: None,
            rules: Vec::new(),
        }
    }
}

impl Theme {
    /// A `theme="N"` is not an index into that order. Excel counts the first
    /// two the other way round, and the second two as well, so 0 is lt1 and 1
    /// is dk1.
    fn colour(&self, index: usize) -> Option<&str> {
        const ORDER: [usize; 12] = [1, 0, 3, 2, 4, 5, 6, 7, 8, 9, 10, 11];
        let slot = *ORDER.get(index)?;
        self.colours.get(slot).map(String::as_str)
    }
}

/// The script whose face a theme font resolves to. Excel picks it by the
/// language it is running as; measured against a Japanese Excel, which is
/// what the corpus is compared to.
const THEME_SCRIPT: &str = "Jpan";

/// The face a shape's run means when it names the theme's scheme instead of a
/// face of its own. `+mn-ea` is "whatever the minor scheme says for East
/// Asian text", `+mj-lt` the major scheme's, and so on 遯ｶ繝ｻthe same resolution
/// a cell's font gets from `scheme="minor"`.
///
/// Left as it stands, `+mn-ea` reaches the device as a face name nothing has,
/// and the missing-face rule answers 繝ｻ・ｭ繝ｻ・ｳ 郢ｧ・ｴ郢ｧ・ｷ郢昴・縺・for it. `tb_r8_youshiki`'s
/// banner then steps its lines 21 pixels where Excel, in the theme's
/// 雋ゑｽｸ郢ｧ・ｴ郢ｧ・ｷ郢昴・縺・ steps 27. 38 shapes across five workbooks name a theme font.
fn scheme_face(face: &str, theme: &Theme) -> Option<String> {
    match face.strip_prefix('+')?.get(..2)? {
        "mj" => theme.major_face.clone(),
        "mn" => theme.minor_face.clone(),
        _ => None,
    }
}

fn parse_theme_xml(xml: &str) -> Theme {
    let mut reader = Reader::from_str(xml);
    let mut theme = Theme::default();
    let mut in_scheme = false;
    let mut colour_slot: Option<usize> = None;
    // Which of the two font schemes is being read, and how good the face
    // held for each is: the entry for this script beats the East Asian
    // fallback, which beats the Latin one.
    let mut in_font_scheme: Option<bool> = None;
    let mut in_rules = false;
    let mut face_rank = [0u8; 2];
    let mut buf = Vec::new();
    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(e)) => match local_name(e.name().as_ref()).as_str() {
                "clrScheme" => in_scheme = true,
                name if in_scheme => {
                    colour_slot = match name {
                        "dk1" => Some(0), "lt1" => Some(1),
                        "dk2" => Some(2), "lt2" => Some(3),
                        "accent1" => Some(4), "accent2" => Some(5),
                        "accent3" => Some(6), "accent4" => Some(7),
                        "accent5" => Some(8), "accent6" => Some(9),
                        "hlink" => Some(10), "folHlink" => Some(11),
                        _ => colour_slot,
                    };
                    let colour = match name {
                        "srgbClr" => get_attr(&e, "val"),
                        "sysClr" => get_attr(&e, "lastClr"),
                        _ => None,
                    };
                    if let (Some(slot), Some(colour)) = (colour_slot, colour) {
                        theme.colours[slot] = colour;
                    }
                }
                "majorFont" => in_font_scheme = Some(true),
                "minorFont" => in_font_scheme = Some(false),
                "lnStyleLst" => in_rules = true,
                "ln" if in_rules => {
                    theme
                        .rules
                        .push(get_attr(&e, "w").and_then(|w| w.parse().ok()).unwrap_or(9525));
                }
                _ => {}
            },
            Ok(Event::End(e)) => match local_name(e.name().as_ref()).as_str() {
                // The colour scheme comes first; reading on gathers the
                // fonts that follow it.
                "clrScheme" => { in_scheme = false; colour_slot = None; },
                "dk1" | "lt1" | "dk2" | "lt2" | "accent1" | "accent2" |
                "accent3" | "accent4" | "accent5" | "accent6" | "hlink" | "folHlink"
                    if in_scheme => colour_slot = None,
                "lnStyleLst" => in_rules = false,
                "majorFont" | "minorFont" => in_font_scheme = None,
                "theme" | "themeElements" => break,
                _ => {}
            },
            Ok(Event::Empty(e)) if in_font_scheme.is_some() => {
                let major = in_font_scheme == Some(true);
                let name = local_name(e.name().as_ref());
                // Rank: the script's own entry, then <a:ea>, then <a:latin>.
                let rank = match name.as_str() {
                    "font" if get_attr(&e, "script").as_deref() == Some(THEME_SCRIPT) => 3,
                    "ea" => 2,
                    "latin" => 1,
                    _ => 0,
                };
                let face = get_attr(&e, "typeface").filter(|face| !face.is_empty());
                if name == "latin" {
                    if major { theme.major_latin_face = face.clone(); }
                    else { theme.minor_latin_face = face.clone(); }
                }
                let held = &mut face_rank[usize::from(major)];
                if let (Some(face), true) = (face, rank > *held) {
                    *held = rank;
                    if major {
                        theme.major_face = Some(face);
                    } else {
                        theme.minor_face = Some(face);
                    }
                }
            }
            Ok(Event::Empty(e)) if in_scheme => {
                // A scheme colour is either stated outright or taken from the
                // system, which records what it last resolved to.
                let colour = match local_name(e.name().as_ref()).as_str() {
                    "srgbClr" => get_attr(&e, "val"),
                    "sysClr" => get_attr(&e, "lastClr"),
                    _ => None,
                };
                if let (Some(slot), Some(colour)) = (colour_slot, colour) {
                    theme.colours[slot] = colour;
                }
            }
            Ok(Event::Eof) | Err(_) => break,
            _ => {}
        }
        buf.clear();
    }
    theme
}

/// Convert a tint to the color step stored by spreadsheet applications.
pub fn quantise_cell_tint(tint: f64) -> f64 {
    if !tint.is_finite() {return 0.0;}
    let scaled=tint.clamp(-1.0,1.0)*32767.0;
    let step=if (scaled-scaled.round()).abs()<1.0e-8 {scaled.round()} else {scaled.trunc()};
    step/32767.0
}

/// Resolve a cell RGB color using its stored tint precision and HLS rounding.
pub fn tint_cell_color(hex: &str, tint: f64) -> String {
    let tint=quantise_cell_tint(tint);
    if tint==0.0 {return hex.to_string();}
    let Ok(rgb)=i64::from_str_radix(hex,16) else {return hex.to_string();};
    let packed=((rgb & 0xff)<<16) | (rgb & 0xff00) | ((rgb>>16)&0xff);
    let color=tinted_cell_color_value(packed,tint);
    format!("{:02X}{:02X}{:02X}",color & 0xff,(color>>8)&0xff,(color>>16)&0xff)
}

fn tinted_cell_color_value(colour: i64, tint: f64) -> i64 {
    const HLSMAX: i64 = 240;
    const RGBMAX: i64 = 255;
    // No tint, no round trip: the conversion is lossy, and Excel hands the
    // theme colour back exactly.
    if tint == 0.0 {
        return colour;
    }
    let (r, g, b) = (colour & 0xff, (colour >> 8) & 0xff, (colour >> 16) & 0xff);
    let cmax = r.max(g).max(b);
    let cmin = r.min(g).min(b);
    let lightness = ((cmax + cmin) * HLSMAX + RGBMAX) / (2 * RGBMAX);
    let (hue, saturation) = if cmax == cmin {
        (0, 0)
    } else {
        let saturation = if lightness <= HLSMAX / 2 {
            ((cmax - cmin) * HLSMAX + (cmax + cmin) / 2) / (cmax + cmin)
        } else {
            ((cmax - cmin) * HLSMAX + (2 * RGBMAX - cmax - cmin) / 2) / (2 * RGBMAX - cmax - cmin)
        };
        let delta = |channel: i64| ((cmax - channel) * (HLSMAX / 6) + (cmax - cmin) / 2) / (cmax - cmin);
        let (red, green, blue) = (delta(r), delta(g), delta(b));
        let mut hue = if r == cmax {
            blue - green
        } else if g == cmax {
            HLSMAX / 3 + red - blue
        } else {
            2 * HLSMAX / 3 + green - red
        };
        if hue < 0 {
            hue += HLSMAX;
        }
        if hue > HLSMAX {
            hue -= HLSMAX;
        }
        (hue, saturation)
    };
    let moved = if tint < 0.0 {
        lightness as f64 * (1.0 + tint)
    } else {
        (lightness as f64 * (1.0 - tint)).floor()
            + HLSMAX as f64 - (HLSMAX as f64 * (1.0 - tint)).floor()
    };
    let lightness = (moved.floor() as i64).clamp(0, HLSMAX);
    if saturation == 0 {
        let grey = ((lightness * RGBMAX + HLSMAX / 2) / HLSMAX).clamp(0, RGBMAX);
        return grey | (grey << 8) | (grey << 16);
    }
    let magic2 = if lightness <= HLSMAX / 2 {
        (lightness * (HLSMAX + saturation) + HLSMAX / 2) / HLSMAX
    } else {
        lightness + saturation - (lightness * saturation + HLSMAX / 2) / HLSMAX
    };
    let magic1 = 2 * lightness - magic2;
    let hue_to_rgb = |mut hue: i64| -> i64 {
        if hue < 0 {
            hue += HLSMAX;
        }
        if hue > HLSMAX {
            hue -= HLSMAX;
        }
        if hue < HLSMAX / 6 {
            magic1 + ((magic2 - magic1) * hue + HLSMAX / 12) / (HLSMAX / 6)
        } else if hue < HLSMAX / 2 {
            magic2
        } else if hue < 2 * HLSMAX / 3 {
            magic1 + ((magic2 - magic1) * (2 * HLSMAX / 3 - hue) + HLSMAX / 12) / (HLSMAX / 6)
        } else {
            magic1
        }
    };
    let channel = |hue: i64| ((hue_to_rgb(hue) * RGBMAX + HLSMAX / 2) / HLSMAX).clamp(0, RGBMAX);
    let (red, green, blue) = (channel(hue + HLSMAX / 3), channel(hue), channel(hue - HLSMAX / 3));
    red | (green << 8) | (blue << 16)
}

fn tinted(hex: &str, tint: f32) -> String {
    tint_cell_color(hex,f64::from(tint))
}

/// The colours a workbook may name by number rather than by value.
///
/// Excel kept a palette of 56 from the days when a file could hold no more,
/// and files still write `<color indexed="12"/>` for blue. A workbook may
/// state its own in `<indexedColors>`, which is why `Theme` carries one.
/// 64 and 65 are not in the palette at all: they mean the system's own
/// foreground and background, which the caller decides.
const PALETTE: [&str; 56] = [
    "000000", "FFFFFF", "FF0000", "00FF00", "0000FF", "FFFF00", "FF00FF", "00FFFF",
    "000000", "FFFFFF", "FF0000", "00FF00", "0000FF", "FFFF00", "FF00FF", "00FFFF",
    "800000", "008000", "000080", "808000", "800080", "008080", "C0C0C0", "808080",
    "9999FF", "993366", "FFFFCC", "CCFFFF", "660066", "FF8080", "0066CC", "CCCCFF",
    "000080", "FF00FF", "FFFF00", "00FFFF", "800080", "800000", "008080", "0000FF",
    "00CCFF", "CCFFFF", "CCFFCC", "FFFF99", "99CCFF", "FF99CC", "CC99FF", "FFCC99",
    "3366FF", "33CCCC", "99CC00", "FFCC00", "FF9900", "FF6600", "666699", "969696",
];

/// The colour a number stands for, from the workbook's own palette where it
/// states one and Excel's otherwise.
fn indexed_colour(index: usize, theme: &Theme) -> Option<String> {
    if let Some(own) = theme.indexed.get(index) {
        return Some(own.clone());
    }
    PALETTE.get(index).map(|hex| hex.to_string())
}

fn resolve_run_font_references(runs: &mut [crate::ir::TextRun], theme: &Theme) {
    for run in runs {
        let Some(reference)=run.font_reference.as_mut() else {continue;};
        let mut element=quick_xml::events::BytesStart::new("color");
        let color=reference.theme.map(|v|v.to_string());
        let indexed=reference.indexed.map(|v|v.to_string());
        let tint=reference.tint.map(|v|v.to_string());
        if let Some(v)=reference.rgb.as_deref() {element.push_attribute(("rgb",v));}
        if let Some(v)=color.as_deref() {element.push_attribute(("theme",v));}
        if let Some(v)=indexed.as_deref() {element.push_attribute(("indexed",v));}
        if let Some(v)=tint.as_deref() {element.push_attribute(("tint",v));}
        if let Some(color)=parse_color_attr(&element,theme) {run.color=Some(color);}
        let render=match reference.scheme.as_deref() {
            Some("major")=>theme.major_face.as_ref(),
            Some("minor")=>theme.minor_face.as_ref(),
            _=>None,
        };
        if let (Some(logical),Some(render))=(run.font.as_ref(),render) {
            reference.latin_name=(logical!=render).then(||logical.clone());
            run.font=Some(render.clone());
        }
    }
}

fn cell_color_reference(e: &quick_xml::events::BytesStart) -> crate::ir::CellColorReference {
    crate::ir::CellColorReference {
        rgb:get_attr(e,"rgb"),theme:get_attr(e,"theme").and_then(|v|v.parse().ok()),
        indexed:get_attr(e,"indexed").and_then(|v|v.parse().ok()),
        tint:get_attr(e,"tint").and_then(|v|v.parse::<f64>().ok()).filter(|v|v.is_finite()),
        automatic:is_true(get_attr(e,"auto").as_deref()),
    }
}

/// Resolve the rendering color with the same theme and palette rules as import.
/// The original color reference remains authoritative when writing the file.
pub fn resolve_cell_color_reference(reference: &crate::ir::CellColorReference,
    theme_colors: &[String], indexed_colors: &[String]) -> Option<String> {
    let mut theme=Theme::default();
    if !theme_colors.is_empty() {theme.colours=theme_colors.to_vec();}
    theme.indexed=indexed_colors.to_vec();
    let mut element=quick_xml::events::BytesStart::new("color");
    let rgb=reference.rgb.as_deref();
    let named=reference.theme.map(|value|value.to_string());
    let indexed=reference.indexed.map(|value|value.to_string());
    let tint=reference.tint.map(|value|value.to_string());
    if let Some(rgb)=rgb {element.push_attribute(("rgb",rgb));}
    if let Some(named)=named.as_deref() {element.push_attribute(("theme",named));}
    if let Some(indexed)=indexed.as_deref() {element.push_attribute(("indexed",indexed));}
    if let Some(tint)=tint.as_deref() {element.push_attribute(("tint",tint));}
    if reference.automatic {element.push_attribute(("auto","1"));}
    parse_color_attr(&element,&theme)
}

fn parse_color_attr(e: &quick_xml::events::BytesStart, theme: &Theme) -> Option<String> {
    let color=if let Some(rgb)=get_attr(e,"rgb") {
        if rgb.len()==8 {rgb[2..].to_string()} else {rgb}
    } else if let Some(index)=get_attr(e,"indexed").and_then(|value|value.parse::<usize>().ok()) {
        indexed_colour(index,theme)?
    } else {
        theme.colour(get_attr(e,"theme")?.parse::<usize>().ok()?)?.to_string()
    };
    let tint=get_attr(e,"tint").and_then(|value|value.parse::<f64>().ok())
        .filter(|value|value.is_finite()).unwrap_or(0.0);
    Some(tint_cell_color(&color,tint))
}

/// The palette a workbook states in place of Excel's own, empty when it
/// states none. `<rgbColor rgb="00RRGGBB"/>`, in order from index zero.
fn indexed_palette(xml: &str) -> Vec<String> {
    let mut reader = Reader::from_str(xml);
    let mut buf = Vec::new();
    let mut inside = false;
    let mut held = Vec::new();
    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(ref e)) | Ok(Event::Empty(ref e)) => {
                match local_name(e.name().as_ref()).as_str() {
                    "indexedColors" => inside = true,
                    "rgbColor" if inside => {
                        if let Some(rgb) = get_attr(e, "rgb") {
                            let hex = if rgb.len() == 8 { &rgb[2..] } else { &rgb };
                            held.push(hex.to_string());
                        }
                    }
                    _ => {}
                }
            }
            Ok(Event::End(ref e)) => {
                if local_name(e.name().as_ref()) == "indexedColors" {
                    break;
                }
            }
            Ok(Event::Eof) | Err(_) => break,
            _ => {}
        }
        buf.clear();
    }
    held
}

/// The differential formats a workbook holds, in the order a `dxfId` counts
/// them.
///
/// A dxf states only what it changes, which is why every part of the answer is
/// optional: a rule that turns text red says nothing about the fill, and the
/// cell keeps the one it had. Excel writes a dxf's fill colour in `<bgColor>`
/// where an ordinary fill puts it in `<fgColor>`, so the background is read
/// first and the foreground only stands in for it.
fn parse_dxfs(xml: &str, theme: &Theme) -> Vec<crate::ir::DiffStyle> {
    let Some(block) = xml
        .split("<dxfs")
        .nth(1)
        .and_then(|rest| rest.split("</dxfs>").next())
    else {
        return Vec::new();
    };
    let mut held = Vec::new();
    let mut reader = Reader::from_str(block);
    let mut current: Option<crate::ir::DiffStyle> = None;
    let mut dxf_start = None;
    let (mut in_font, mut in_fill) = (false, false);
    loop {
        let offset = oxidocs_common::xml_utils::xml_byte_position(block, reader.buffer_position());
        let event = reader.read_event();
        let event_end = oxidocs_common::xml_utils::xml_byte_position(block, reader.buffer_position());
        let empty = matches!(&event, Ok(Event::Empty(_)));
        match event {
            Ok(Event::Start(ref e)) | Ok(Event::Empty(ref e)) => {
                let name = local_name(e.name().as_ref());
                match name.as_str() {
                    "dxf" => {
                        let mut style = crate::ir::DiffStyle::default();
                        if empty {
                            style.source_xml = Some(block[offset..event_end].to_string());
                            held.push(style);
                        } else {
                            current = Some(style);
                            dxf_start = Some(offset);
                        }
                        continue;
                    }
                    "font" => {
                        in_font = true;
                        continue;
                    }
                    "fill" => {
                        in_fill = true;
                        continue;
                    }
                    _ => {}
                }
                let Some(style) = current.as_mut() else { continue };
                // `<b/>` on its own is on; only `val="0"` turns one off.
                let on = !matches!(get_attr(e, "val").as_deref(), Some("0") | Some("false"));
                match name.as_str() {
                    "b" if in_font => style.bold = Some(on),
                    "i" if in_font => style.italic = Some(on),
                    "u" if in_font => style.underline = Some(on),
                    "strike" if in_font => style.strikethrough = Some(on),
                    "color" if in_font => style.font_color = parse_color_attr(e, theme),
                    "bgColor" if in_fill => style.bg_color = parse_color_attr(e, theme),
                    "fgColor" if in_fill && style.bg_color.is_none() => {
                        style.bg_color = parse_color_attr(e, theme);
                    }
                    "numFmt" => style.number_format = get_attr(e, "formatCode")
                        .map(|code| oxicells_calc::numfmt::number_format_from_storage(&code, false)),
                    _ => {}
                }
            }
            Ok(Event::End(ref e)) => match local_name(e.name().as_ref()).as_str() {
                "font" => in_font = false,
                "fill" => in_fill = false,
                "dxf" => {
                    if let Some(mut style) = current.take() {
                        if let Some(start) = dxf_start.take() { style.source_xml = Some(block[start..event_end].to_string()); }
                        held.push(style);
                    }
                }
                _ => {}
            },
            Ok(Event::Eof) | Err(_) => break,
            _ => {}
        }
    }
    held
}

// Importing a workbook reconciles stored application bits with the component
// definitions. Later explicit edits keep their own application bits.
#[derive(Clone, Debug, Default)]
struct ImportedStyleElement {
    start: usize,
    end: usize,
    name: String,
    attributes: std::collections::BTreeMap<String, String>,
    children: Vec<ImportedStyleElement>,
    text: String,
}

impl ImportedStyleElement {
    fn child(&self, name: &str) -> Option<&Self> {
        self.children.iter().find(|child| child.name == name)
    }

    fn id(&self, name: &str) -> usize {
        self.attributes.get(name).and_then(|value| value.parse().ok()).unwrap_or(0)
    }

    fn component(&self) -> ImportedStyleComponent {
        let mut attributes: Vec<_> = self.attributes.iter().filter_map(|(name, value)| {
            if name == "xmlns" || name.starts_with("xmlns:") {
                return None;
            }
            let value = match (self.name.as_str(), name.as_str(), value.as_str()) {
                ("alignment", "horizontal", "general")
                | ("alignment", "vertical", "bottom")
                | ("alignment", "wrapText" | "shrinkToFit" | "justifyLastLine", "0" | "false")
                | ("alignment", "indent" | "relativeIndent" | "readingOrder" | "textRotation", "0")
                | ("protection", "locked", "1" | "true")
                | ("protection", "hidden", "0" | "false") => return None,
                (_, "val", "true") if self.boolean_font_property() => "1",
                (_, "val", "false") if self.boolean_font_property() => "0",
                _ => value,
            };
            Some((name.clone(), value.to_string()))
        }).collect();
        if self.boolean_font_property() && !self.attributes.contains_key("val") {
            attributes.push(("val".into(), "1".into()));
        }
        if self.name == "u" && !self.attributes.contains_key("val") {
            attributes.push(("val".into(), "single".into()));
        }
        attributes.sort();
        let mut children: Vec<_> = self.children.iter().map(Self::component).collect();
        children.sort();
        ImportedStyleComponent {
            name: self.name.clone(),
            attributes,
            children,
            text: self.text.clone(),
        }
    }

    fn boolean_font_property(&self) -> bool {
        matches!(self.name.as_str(), "b" | "i" | "strike" | "outline" | "shadow" | "condense" | "extend")
    }
}

#[derive(Clone, Debug, PartialEq, Eq, PartialOrd, Ord)]
struct ImportedStyleComponent {
    name: String,
    attributes: Vec<(String, String)>,
    children: Vec<ImportedStyleComponent>,
    text: String,
}

fn imported_style_elements(events: &[Event<'_>]) -> Result<Option<ImportedStyleElement>, XlsxError> {
    let mut stack: Vec<ImportedStyleElement> = Vec::new();
    let mut root = None;
    for (index, event) in events.iter().enumerate() {
        match event {
            Event::Start(tag) | Event::Empty(tag) => {
                let attributes = tag.attributes().map(|attribute| {
                    let attribute = attribute.map_err(|error| XlsxError::InvalidData(error.to_string()))?;
                    let name = String::from_utf8_lossy(attribute.key.as_ref()).into_owned();
                    let value = attribute.unescape_value()?.into_owned();
                    Ok((name, value))
                }).collect::<Result<_, XlsxError>>()?;
                let node = ImportedStyleElement {
                    start: index,
                    end: index,
                    name: String::from_utf8_lossy(tag.local_name().as_ref()).into_owned(),
                    attributes,
                    ..Default::default()
                };
                if matches!(event, Event::Start(_)) {
                    stack.push(node);
                } else if let Some(parent) = stack.last_mut() {
                    parent.children.push(node);
                } else {
                    root = Some(node);
                }
            }
            Event::End(_) => {
                let Some(mut node) = stack.pop() else {
                    return Err(XlsxError::InvalidData("Unexpected styles closing element".into()));
                };
                node.end = index;
                if let Some(parent) = stack.last_mut() {
                    parent.children.push(node);
                } else {
                    root = Some(node);
                }
            }
            Event::Text(text) => {
                if let Some(node) = stack.last_mut() {
                    let text = text.unescape()?;
                    if !text.trim().is_empty() {
                        node.text.push_str(&text);
                    }
                }
            }
            Event::CData(text) => {
                if let Some(node) = stack.last_mut() {
                    node.text.push_str(&String::from_utf8_lossy(text.as_ref()));
                }
            }
            _ => {}
        }
    }
    Ok(root)
}

/// Reconcile application flags when a package is first imported.
///
/// Comparison uses the complete XML definitions, including properties the
/// rendering IR does not yet expose. Runtime setters must not call this:
/// explicitly assigning even an unchanged value has separate save semantics.
pub fn import_styles_xml(xml: &str) -> Result<String, XlsxError> {
    reconcile_imported_styles_xml(xml, true)
}

fn reconcile_imported_styles_xml(xml: &str, normalize_normal_font: bool) -> Result<String, XlsxError> {
    use quick_xml::Writer;
    const FLAGS: [&str; 6] = ["applyNumberFormat", "applyFont", "applyFill", "applyBorder", "applyAlignment", "applyProtection"];
    let mut reader = Reader::from_str(xml);
    let mut events = Vec::new();
    loop {
        let event = reader.read_event()?.into_owned();
        if matches!(event, Event::Eof) { break; }
        events.push(event);
    }
    let Some(root) = imported_style_elements(&events)? else { return Ok(xml.to_string()); };
    let Some(cell_xfs) = root.child("cellXfs") else { return Ok(xml.to_string()); };
    let Some(style_xfs) = root.child("cellStyleXfs") else { return Ok(xml.to_string()); };
    let mut parents = style_xfs.children.clone();
    let mut replacements = HashMap::new();
    // The default cell font supplies Normal's font on import. Other Normal
    // components remain those of its named-style definition.
    if let (Some(normal), Some(default)) = (parents.first_mut(), cell_xfs.children.first()) {
        let font_id = default.attributes.get("fontId").map(String::as_str).unwrap_or("0");
        normal.attributes.insert("fontId".into(), font_id.to_string());
        if let (true, Event::Start(tag) | Event::Empty(tag)) = (normalize_normal_font, &events[normal.start]) {
            let mut updated = tag.clone();
            updated.clear_attributes();
            for attribute in tag.attributes() {
                let attribute = attribute.map_err(|error| XlsxError::InvalidData(error.to_string()))?;
                if attribute.key.as_ref() != b"fontId" { updated.push_attribute(attribute); }
            }
            updated.push_attribute(("fontId", font_id));
            replacements.insert(normal.start, if matches!(events[normal.start], Event::Empty(_)) {
                Event::Empty(updated)
            } else { Event::Start(updated) });
        }
    }
    let definitions = |section: &str, index: usize| {
        root.child(section).and_then(|section| section.children.get(index)).map(ImportedStyleElement::component)
    };
    let different_definition = |section: &str, left: usize, right: usize| {
        if left == right { return false; }
        match (definitions(section, left), definitions(section, right)) {
            (Some(left), Some(right)) => left != right,
            _ => true,
        }
    };
    let component = |xf: &ImportedStyleElement, name: &str| {
        xf.child(name).map(ImportedStyleElement::component).unwrap_or_else(|| ImportedStyleComponent {
            name: name.into(), attributes: Vec::new(), children: Vec::new(), text: String::new(),
        })
    };
    let number_format = |xf: &ImportedStyleElement| {
        let id = xf.id("numFmtId") as u32;
        root.child("numFmts").and_then(|section| section.children.iter().find(|format| format.id("numFmtId") == id as usize))
            .and_then(|format| format.attributes.get("formatCode"))
            .map(|code| oxicells_calc::numfmt::number_format_from_storage(code, false))
            .or_else(|| builtin_number_format(id).map(str::to_string))
            .unwrap_or_else(|| format!("#{id}"))
    };
    for xf in &cell_xfs.children {
        let Some(parent) = parents.get(xf.id("xfId")) else { continue; };
        let differs = [
            number_format(xf) != number_format(parent),
            different_definition("fonts", xf.id("fontId"), parent.id("fontId")),
            different_definition("fills", xf.id("fillId"), parent.id("fillId")),
            different_definition("borders", xf.id("borderId"), parent.id("borderId")),
            component(xf, "alignment") != component(parent, "alignment"),
            component(xf, "protection") != component(parent, "protection"),
        ];
        let (Event::Start(tag) | Event::Empty(tag)) = &events[xf.start] else { continue; };
        let mut updated = tag.clone();
        updated.clear_attributes();
        for attribute in tag.attributes() {
            let attribute = attribute.map_err(|error| XlsxError::InvalidData(error.to_string()))?;
            if !FLAGS.iter().any(|name| attribute.key.as_ref() == name.as_bytes()) {
                updated.push_attribute(attribute);
            }
        }
        for (flag, differs) in FLAGS.into_iter().zip(differs) {
            if differs { updated.push_attribute((flag, "1")); }
        }
        replacements.insert(xf.start, if matches!(events[xf.start], Event::Empty(_)) {
            Event::Empty(updated)
        } else { Event::Start(updated) });
    }
    let mut writer = Writer::new(Vec::new());
    for (index, event) in events.into_iter().enumerate() {
        writer.write_event(replacements.remove(&index).unwrap_or(event))
            .map_err(|error| XlsxError::InvalidData(error.to_string()))?;
    }
    String::from_utf8(writer.into_inner()).map_err(|error| XlsxError::InvalidData(error.to_string()))
}

fn parse_styles_xml(xml: &str, theme: &Theme) -> Result<StyleSheet, XlsxError> {
    // Keep the source Normal font for column metrics until serialization.
    // Imported application flags still compare against the effective font.
    let imported = reconcile_imported_styles_xml(xml, false)?;
    let xml = imported.as_str();
    let mut reader = quick_xml::reader::NsReader::from_str(xml);
    let mut root_seen = false;
    let mut allow_unqualified = false;
    let mut ignored_depth = 0usize;
    let mut ss = StyleSheet::default();

    // Parsing state
    #[derive(PartialEq)]
    enum Section {
        None,
        NumFmts,
        Fonts,
        Fills,
        Borders,
        CellXfs,
        CellStyleXfs,
        /// The differential formats. A table states the rule it draws round
        /// itself as an index into these, not into the ordinary borders.
        Dxfs,
    }
    let mut section = Section::None;
    let mut in_font = false;
    let mut current_font = FontInfo::default();
    let mut current_font_scheme: Option<String> = None;
    let mut in_fill = false;
    let mut current_fill = FillInfo::default();
    let mut in_border = false;
    // Which edge of a border is open, so its colour knows where it belongs.
    let mut border_edge: Option<Edge> = None;
    let mut current_border = BorderInfo::default();
    // How many differential formats have been closed.
    let mut dxf_seen = 0usize;
    let mut in_xf = false;
    let mut current_xf = XfRecord::default();

    loop {
        use quick_xml::name::ResolveResult;
        let (namespace, event) = reader.read_resolved_event()?;
        if !root_seen {
            if let Event::Start(tag) | Event::Empty(tag) = &event {
                allow_unqualified = tag.name().as_ref() == b"styleSheet"
                    && matches!(&namespace, ResolveResult::Unbound);
                root_seen = true;
            }
        }
        let standard = matches!(&namespace, ResolveResult::Bound(ns)
            if ns.as_ref() == b"http://schemas.openxmlformats.org/spreadsheetml/2006/main")
            || (allow_unqualified && matches!(&namespace, ResolveResult::Unbound));
        if ignored_depth != 0 {
            match event {
                Event::Start(_) => ignored_depth += 1,
                Event::End(_) => ignored_depth -= 1,
                Event::Eof => break,
                _ => {}
            }
            continue;
        }
        if !standard {
            match event {
                Event::Start(_) => ignored_depth = 1,
                Event::Eof => break,
                _ => {}
            }
            continue;
        }
        match event {
            Event::Start(e) => {
                let name = local_name(e.name().as_ref());
                match name.as_str() {
                    "cellStyle" => {
                        if let Some(name)=get_attr(&e,"name") {
                            let index=get_attr(&e,"xfId").and_then(|v|v.parse::<usize>().ok()).unwrap_or(0);
                            let builtin=get_attr(&e,"builtinId").and_then(|v|v.parse::<u32>().ok());
                            ss.named_styles.push((name,index,builtin));
                        }
                    }

                    "numFmts" => section = Section::NumFmts,
                    "fonts" => section = Section::Fonts,
                    "fills" => section = Section::Fills,
                    "borders" => section = Section::Borders,
                    "dxfs" => section = Section::Dxfs,
                    "cellXfs" => section = Section::CellXfs,
                    "cellStyleXfs" => section = Section::CellStyleXfs,

                    "font" if section == Section::Fonts => {
                        in_font = true;
                        current_font = FontInfo::default();
                    }
                    "u" if in_font => {
                        let val = get_attr(&e, "val");
                        current_font.underline = !matches!(val.as_deref(),Some("0")|Some("none"));
                        current_font.underline_kind = match val.as_deref() {
                            Some("double") => Some(crate::ir::CellUnderlineKind::Double),
                            Some("singleAccounting") => Some(crate::ir::CellUnderlineKind::SingleAccounting),
                            Some("doubleAccounting") => Some(crate::ir::CellUnderlineKind::DoubleAccounting),
                            _ => None,
                        };
                    }
                    "vertAlign" if in_font => {
                        current_font.font_position = match get_attr(&e,"val").as_deref() {
                            Some("superscript") => Some(crate::ir::CellFontPosition::Superscript),
                            Some("subscript") => Some(crate::ir::CellFontPosition::Subscript),
                            _ => None,
                        };
                    }
                    "fill" if section == Section::Fills => {
                        in_fill = true;
                        current_fill = FillInfo::default();
                    }
                    "border" if section == Section::Borders || section == Section::Dxfs => {
                        in_border = true;
                        current_border = BorderInfo::default();
                        // Which way the corner-to-corner rule runs is stated
                        // on the border itself, and a cell can carry both.
                        current_border.up = is_true(get_attr(&e, "diagonalUp").as_deref());
                        current_border.down = is_true(get_attr(&e, "diagonalDown").as_deref());
                    }
                    "xf" if section == Section::CellXfs
                        || section == Section::CellStyleXfs =>
                    {
                        in_xf = true;
                        current_xf = XfRecord {
                            num_fmt_id: get_attr(&e, "numFmtId")
                                .and_then(|v| v.parse().ok())
                                .unwrap_or(0),
                            font_id: get_attr(&e, "fontId")
                                .and_then(|v| v.parse().ok())
                                .unwrap_or(0),
                            fill_id: get_attr(&e, "fillId")
                                .and_then(|v| v.parse().ok())
                                .unwrap_or(0),
                            border_id: get_attr(&e, "borderId")
                                .and_then(|v| v.parse().ok())
                                .unwrap_or(0),
                            horizontal_align: None,
                            indent: 0,
                            vertical_align: None,
                            wrap_text: false,
                            stacked_text: false,
                            text_rotation: None,
                            justify_last_line: false,
                            shrink_to_fit: false,
                            reading_order: crate::ir::CellReadingOrder::Contextual,
                            quote_prefix: is_true(get_attr(&e, "quotePrefix").as_deref()),
                            format_reference: xf_format_reference(&e),
                            unlocked: false,
                            formula_hidden: false,
                        };
                    }
                    "alignment" if in_xf => {
                        current_xf.horizontal_align = get_attr(&e, "horizontal");
                        current_xf.vertical_align = get_attr(&e, "vertical");
                        current_xf.indent = get_attr(&e, "indent")
                            .and_then(|level| level.parse().ok())
                            .unwrap_or(0);
                        current_xf.wrap_text =
                            matches!(get_attr(&e, "wrapText").as_deref(), Some("1") | Some("true"));
                        // 255 is not an angle: it is Excel's way of saying the
                        // characters stand one above the next.
                        current_xf.stacked_text =
                            get_attr(&e, "textRotation").as_deref() == Some("255");
                        current_xf.text_rotation = match get_attr(&e, "textRotation").and_then(|value|value.parse::<i32>().ok()) {
                            Some(angle @ 1..=90) => Some(angle),
                            Some(angle @ 91..=180) => Some(90-angle),
                            _ => None,
                        };
                        current_xf.justify_last_line = is_true(get_attr(&e, "justifyLastLine").as_deref());
                        current_xf.shrink_to_fit =
                            is_true(get_attr(&e, "shrinkToFit").as_deref());
                        current_xf.reading_order = match get_attr(&e, "readingOrder").as_deref() {
                            Some("1") => crate::ir::CellReadingOrder::LeftToRight,
                            Some("2") => crate::ir::CellReadingOrder::RightToLeft,
                            _ => crate::ir::CellReadingOrder::Contextual,
                        };
                    }
                    // <protection locked="0"/> 遯ｶ繝ｻthe cells a protected sheet
                    // still lets a person type into.
                    "protection" if in_xf => {
                        current_xf.unlocked =
                            matches!(get_attr(&e, "locked").as_deref(), Some("0") | Some("false"));
                        current_xf.formula_hidden = is_true(get_attr(&e, "hidden").as_deref());
                    }

                    // Inside a border element, parse child elements with style attr
                    //
                    // An edge that states a colour states it as a CHILD, so it
                    // arrives as a Start and its `<color>` follows. Which edge
                    // is open has to be remembered for the colour to know
                    // where it belongs.
                    "left" if in_border => {
                        current_border.left = border_line(&e);
                        border_edge = Some(Edge::Left);
                    }
                    "right" if in_border => {
                        current_border.right = border_line(&e);
                        border_edge = Some(Edge::Right);
                    }
                    "top" if in_border => {
                        current_border.top = border_line(&e);
                        border_edge = Some(Edge::Top);
                    }
                    "bottom" if in_border => {
                        current_border.bottom = border_line(&e);
                        border_edge = Some(Edge::Bottom);
                    }
                    // A diagonal that states a colour has a child, so it
                    // arrives here and not among the self-closing sides.
                    "diagonal" if in_border => {
                        current_border.diagonal = border_line(&e);
                        border_edge = Some(Edge::Diagonal);
                    }

                    // The colour of the edge that is open. `indexed="64"` is
                    // Excel's "automatic" and is past the end of the palette,
                    // so it reads back as no colour 遯ｶ繝ｻwhich is what it means.
                    "color" if border_edge.is_some() => {
                        let held = parse_color_attr(&e, theme);
                        if let Some(line) = match border_edge {
                            Some(Edge::Left) => current_border.left.as_mut(),
                            Some(Edge::Right) => current_border.right.as_mut(),
                            Some(Edge::Top) => current_border.top.as_mut(),
                            Some(Edge::Bottom) => current_border.bottom.as_mut(),
                            Some(Edge::Diagonal) => current_border.diagonal.as_mut(),
                            None => None,
                        } {
                            line.color = held;
                            let reference=cell_color_reference(&e);
                            let default=if matches!(border_edge,Some(Edge::Diagonal)) {
                                crate::ir::CellColorReference {automatic:true,..Default::default()}
                            } else {crate::ir::CellColorReference {indexed:Some(64),..Default::default()}};
                            let rgb=reference.rgb.as_deref().is_some_and(|rgb|rgb.len()==8 && rgb.get(..2).is_some_and(|prefix|prefix.eq_ignore_ascii_case("FF")))
                                && reference.theme.is_none() && reference.indexed.is_none()
                                && reference.tint.is_none() && !reference.automatic;
                            let canonical=rgb || reference==default || reference==crate::ir::CellColorReference::default();
                            line.color_reference=(!canonical).then(||Box::new(reference));
                        }
                    }

                    // Font color
                    "color" if in_font => {
                        if let Some(c) = parse_color_attr(&e, theme) {
                            current_font.color = Some(c);
                        }
                        current_font.reference.rgb = get_attr(&e, "rgb");
                        current_font.reference.implicit_automatic = false;
                        current_font.reference.theme = get_attr(&e, "theme").and_then(|v| v.parse().ok());
                        current_font.reference.indexed = get_attr(&e, "indexed").and_then(|v| v.parse().ok());
                        current_font.reference.automatic = is_true(get_attr(&e, "auto").as_deref());
                        current_font.reference.tint = get_attr(&e, "tint").and_then(|v| v.parse::<f64>().ok()).filter(|v| v.is_finite());
                        if get_attr(&e, "rgb").is_some() {
                            current_font.reference.theme = None;
                            current_font.reference.indexed = None;
                            current_font.reference.automatic = false;
                        }
                    }

                    // Fill color 遯ｶ繝ｻlook for fgColor inside patternFill
                    "patternFill" if in_fill => {
                        current_fill.pattern_fill.pattern=get_attr(&e,"patternType").unwrap_or_else(||"none".into());
                    }
                    "fgColor" | "bgColor" if in_fill => {
                        let foreground=name=="fgColor";
                        if foreground {current_fill.pattern_fill.foreground=Some(cell_color_reference(&e));}
                        else {current_fill.pattern_fill.background=Some(cell_color_reference(&e));}
                        if foreground==(current_fill.pattern_fill.pattern=="solid") {
                            current_fill.bg_color=parse_color_attr(&e,theme);
                        }
                    }

                    _ => {}
                }
            }
            Event::Empty(e) => {
                let name = local_name(e.name().as_ref());
                match name.as_str() {
                    "font" if section == Section::Fonts => ss.fonts.push(FontInfo::default()),
                    "fill" if section == Section::Fills => ss.fills.push(FillInfo::default()),
                    "border" if section == Section::Borders => ss.borders.push(BorderInfo {
                        up: is_true(get_attr(&e, "diagonalUp").as_deref()),
                        down: is_true(get_attr(&e, "diagonalDown").as_deref()),
                        ..BorderInfo::default()
                    }),
                    "cellStyle" => {
                        if let Some(name)=get_attr(&e,"name") {
                            let index=get_attr(&e,"xfId").and_then(|v|v.parse::<usize>().ok()).unwrap_or(0);
                            let builtin=get_attr(&e,"builtinId").and_then(|v|v.parse::<u32>().ok());
                            ss.named_styles.push((name,index,builtin));
                        }
                    }

                    "numFmt" if section == Section::NumFmts => {
                        if let (Some(id_str), Some(code)) =
                            (get_attr(&e, "numFmtId"), get_attr(&e, "formatCode"))
                        {
                            if let Ok(id) = id_str.parse::<u32>() {
                                ss.num_fmts.insert(id, oxicells_calc::numfmt::number_format_from_storage(&code, false));
                            }
                        }
                    }
                    // Self-closing <b/>, <i/>, <sz val="..."/>
                    "b" if in_font => {
                        // <b/> means bold=true, <b val="0"/> means false
                        let val = get_attr(&e, "val");
                        current_font.bold = val.as_deref() != Some("0");
                    }
                    "i" if in_font => {
                        let val = get_attr(&e, "val");
                        current_font.italic = val.as_deref() != Some("0");
                    }
                    "u" if in_font => {
                        let val = get_attr(&e, "val");
                        current_font.underline = !matches!(val.as_deref(),Some("0")|Some("none"));
                        current_font.underline_kind = match val.as_deref() {
                            Some("double") => Some(crate::ir::CellUnderlineKind::Double),
                            Some("singleAccounting") => Some(crate::ir::CellUnderlineKind::SingleAccounting),
                            Some("doubleAccounting") => Some(crate::ir::CellUnderlineKind::DoubleAccounting),
                            _ => None,
                        };
                    }
                    "vertAlign" if in_font => {
                        current_font.font_position = match get_attr(&e,"val").as_deref() {
                            Some("superscript") => Some(crate::ir::CellFontPosition::Superscript),
                            Some("subscript") => Some(crate::ir::CellFontPosition::Subscript),
                            _ => None,
                        };
                    }
                    "strike" if in_font => {
                        // <strike/> strikes the writing through;
                        // <strike val="0"/> does not.
                        let val = get_attr(&e, "val");
                        current_font.strikethrough = val.as_deref() != Some("0");
                    }
                    "sz" if in_font => {
                        current_font.size =
                            get_attr(&e, "val").and_then(|v| v.parse().ok());
                    }
                    "name" | "rFont" if in_font => {
                        if let Some(encoded) = get_attr(&e, "val") {
                            let units = crate::string_value::decode(&encoded);
                            current_font.reference.latin_name_utf16 = String::from_utf16(&units).err().map(|_| units.clone());
                            current_font.name = Some(String::from_utf16_lossy(&units));
                        }
                    }
                    "charset" if in_font => {
                        current_font.charset =
                            get_attr(&e, "val").and_then(|v| v.parse().ok());
                    }
                    "family" if in_font => {
                        current_font.family =
                            get_attr(&e, "val").and_then(|v| v.parse().ok());
                    }
                    "scheme" if in_font => {
                        current_font_scheme = get_attr(&e, "val");
                    }
                    // A rule's colour is written self-closing, so it lands
                    // here rather than among the Start events 遯ｶ繝ｻwhich is why
                    // it read back black for as long as it did.
                    "color" if border_edge.is_some() => {
                        let held = parse_color_attr(&e, theme);
                        if let Some(line) = match border_edge {
                            Some(Edge::Left) => current_border.left.as_mut(),
                            Some(Edge::Right) => current_border.right.as_mut(),
                            Some(Edge::Top) => current_border.top.as_mut(),
                            Some(Edge::Bottom) => current_border.bottom.as_mut(),
                            Some(Edge::Diagonal) => current_border.diagonal.as_mut(),
                            None => None,
                        } {
                            line.color = held;
                            let reference=cell_color_reference(&e);
                            let default=if matches!(border_edge,Some(Edge::Diagonal)) {
                                crate::ir::CellColorReference {automatic:true,..Default::default()}
                            } else {crate::ir::CellColorReference {indexed:Some(64),..Default::default()}};
                            let rgb=reference.rgb.as_deref().is_some_and(|rgb|rgb.len()==8 && rgb.get(..2).is_some_and(|prefix|prefix.eq_ignore_ascii_case("FF")))
                                && reference.theme.is_none() && reference.indexed.is_none()
                                && reference.tint.is_none() && !reference.automatic;
                            let canonical=rgb || reference==default || reference==crate::ir::CellColorReference::default();
                            line.color_reference=(!canonical).then(||Box::new(reference));
                        }
                    }
                    "color" if in_font => {
                        if let Some(c) = parse_color_attr(&e, theme) {
                            current_font.color = Some(c);
                        }
                        current_font.reference.rgb = get_attr(&e, "rgb");
                        current_font.reference.implicit_automatic = false;
                        current_font.reference.theme = get_attr(&e, "theme").and_then(|v| v.parse().ok());
                        current_font.reference.indexed = get_attr(&e, "indexed").and_then(|v| v.parse().ok());
                        current_font.reference.automatic = is_true(get_attr(&e, "auto").as_deref());
                        current_font.reference.tint = get_attr(&e, "tint").and_then(|v| v.parse::<f64>().ok()).filter(|v| v.is_finite());
                        if get_attr(&e, "rgb").is_some() {
                            current_font.reference.theme = None;
                            current_font.reference.indexed = None;
                            current_font.reference.automatic = false;
                        }
                    }
                    "patternFill" if in_fill => {
                        current_fill.pattern_fill.pattern=get_attr(&e,"patternType").unwrap_or_else(||"none".into());
                    }
                    "fgColor" | "bgColor" if in_fill => {
                        let foreground=name=="fgColor";
                        if foreground {current_fill.pattern_fill.foreground=Some(cell_color_reference(&e));}
                        else {current_fill.pattern_fill.background=Some(cell_color_reference(&e));}
                        if foreground==(current_fill.pattern_fill.pattern=="solid") {
                            current_fill.bg_color=parse_color_attr(&e,theme);
                        }
                    }
                    // Self-closing border sides: <left style="thin"/>
                    // An edge written self-closing states no colour, so it
                    // leaves nothing open behind it.
                    "left" if in_border => {
                        current_border.left = border_line(&e);
                        border_edge = None;
                    }
                    "right" if in_border => {
                        current_border.right = border_line(&e);
                        border_edge = None;
                    }
                    "top" if in_border => {
                        current_border.top = border_line(&e);
                        border_edge = None;
                    }
                    "bottom" if in_border => {
                        current_border.bottom = border_line(&e);
                        border_edge = None;
                    }
                    "diagonal" if in_border => {
                        current_border.diagonal = border_line(&e);
                        border_edge = None;
                    }
                    "alignment" if in_xf => {
                        current_xf.horizontal_align = get_attr(&e, "horizontal");
                        current_xf.vertical_align = get_attr(&e, "vertical");
                        current_xf.indent = get_attr(&e, "indent")
                            .and_then(|level| level.parse().ok())
                            .unwrap_or(0);
                        current_xf.wrap_text =
                            matches!(get_attr(&e, "wrapText").as_deref(), Some("1") | Some("true"));
                        // 255 is not an angle: it is Excel's way of saying the
                        // characters stand one above the next.
                        current_xf.stacked_text =
                            get_attr(&e, "textRotation").as_deref() == Some("255");
                        current_xf.text_rotation = match get_attr(&e, "textRotation").and_then(|value|value.parse::<i32>().ok()) {
                            Some(angle @ 1..=90) => Some(angle),
                            Some(angle @ 91..=180) => Some(90-angle),
                            _ => None,
                        };
                        current_xf.justify_last_line = is_true(get_attr(&e, "justifyLastLine").as_deref());
                        current_xf.shrink_to_fit =
                            is_true(get_attr(&e, "shrinkToFit").as_deref());
                        current_xf.reading_order = match get_attr(&e, "readingOrder").as_deref() {
                            Some("1") => crate::ir::CellReadingOrder::LeftToRight,
                            Some("2") => crate::ir::CellReadingOrder::RightToLeft,
                            _ => crate::ir::CellReadingOrder::Contextual,
                        };
                    }
                    // <protection locked="0"/> 遯ｶ繝ｻthe cells a protected sheet
                    // still lets a person type into.
                    "protection" if in_xf => {
                        current_xf.unlocked =
                            matches!(get_attr(&e, "locked").as_deref(), Some("0") | Some("false"));
                        current_xf.formula_hidden = is_true(get_attr(&e, "hidden").as_deref());
                    }
                    "xf" if section == Section::CellXfs
                        || section == Section::CellStyleXfs =>
                    {
                        // Self-closing <xf ... />
                        let xf = XfRecord {
                            num_fmt_id: get_attr(&e, "numFmtId")
                                .and_then(|v| v.parse().ok())
                                .unwrap_or(0),
                            font_id: get_attr(&e, "fontId")
                                .and_then(|v| v.parse().ok())
                                .unwrap_or(0),
                            fill_id: get_attr(&e, "fillId")
                                .and_then(|v| v.parse().ok())
                                .unwrap_or(0),
                            border_id: get_attr(&e, "borderId")
                                .and_then(|v| v.parse().ok())
                                .unwrap_or(0),
                            horizontal_align: None,
                            indent: 0,
                            vertical_align: None,
                            wrap_text: false,
                            stacked_text: false,
                            text_rotation: None,
                            justify_last_line: false,
                            shrink_to_fit: false,
                            reading_order: crate::ir::CellReadingOrder::Contextual,
                            quote_prefix: is_true(get_attr(&e, "quotePrefix").as_deref()),
                            format_reference: xf_format_reference(&e),
                            unlocked: false,
                            formula_hidden: false,
                        };
                        if section == Section::CellStyleXfs {
                            ss.cell_style_xfs.push(xf);
                        } else {
                            ss.cell_xfs.push(xf);
                        }
                    }

                    _ => {}
                }
            }
            Event::End(e) => {
                let name = local_name(e.name().as_ref());
                match name.as_str() {
                    "numFmts" | "fonts" | "fills" | "borders" | "cellXfs"
                    | "cellStyleXfs" => {
                        section = Section::None;
                    }
                    "font" if in_font => {
                        // A font that names a theme scheme wears the face
                        // that scheme states, and its own <name> counts for
                        // nothing: openpyxl writes Calibri with
                        // `scheme="minor"` and Excel opens the workbook in
                        // the theme's 繝ｻ・ｭ繝ｻ・ｳ 繝ｻ・ｰ郢ｧ・ｴ郢ｧ・ｷ郢昴・縺・ rows and all.
                        let scheme = current_font_scheme.take();
                        current_font.latin_name = match scheme.as_deref() {
                            Some("major") => theme.major_latin_face.clone(),
                            Some("minor") => theme.minor_latin_face.clone(),
                            _ => None,
                        }.or_else(|| current_font.name.clone());
                        let face = match scheme.as_deref() {
                            Some("major") => theme.major_face.clone(),
                            Some("minor") => theme.minor_face.clone(),
                            _ => None,
                        };
                        if let Some(face) = face {
                            current_font.name = Some(face);
                            current_font.reference.latin_name_utf16 = None;
                        }
                        current_font.reference.scheme = scheme;
                        current_font.reference.latin_name = current_font.latin_name.clone()
                            .filter(|name| Some(name) != current_font.name.as_ref());
                        ss.fonts.push(std::mem::take(&mut current_font));
                        in_font = false;
                    }
                    "fill" if in_fill => {
                        ss.fills.push(std::mem::take(&mut current_fill));
                        in_fill = false;
                    }
                    // Each side closes before the next opens. Forgetting to
                    // let go of it would hand the next `<color>` in the file 遯ｶ繝ｻ
                    // a font's, most likely 遯ｶ繝ｻto a rule that had finished.
                    "left" | "right" | "top" | "bottom" | "diagonal" if in_border => {
                        border_edge = None;
                    }
                    "border" if in_border => {
                        border_edge = None;
                        let held = std::mem::take(&mut current_border);
                        if section == Section::Dxfs {
                            ss.dxf_borders.push(held);
                        } else {
                            ss.borders.push(held);
                        }
                        in_border = false;
                    }
                    // Every differential format takes a place in the list,
                    // whether or not it states a border 遯ｶ繝ｻa table's index
                    // counts them all, so one that only changes a font must
                    // still hold its place.
                    "dxf" if section == Section::Dxfs => {
                        dxf_seen += 1;
                        while ss.dxf_borders.len() < dxf_seen {
                            ss.dxf_borders.push(BorderInfo::default());
                        }
                    }
                    "dxfs" => section = Section::None,
                    "xf" if in_xf
                        && (section == Section::CellXfs
                            || section == Section::CellStyleXfs) =>
                    {
                        let xf = std::mem::take(&mut current_xf);
                        if section == Section::CellStyleXfs {
                            ss.cell_style_xfs.push(xf);
                        } else {
                            ss.cell_xfs.push(xf);
                        }
                        in_xf = false;
                    }
                    _ => {}
                }
            }
            Event::Eof => break,
            _ => {}
        }
    }

    // The dxfs are read in one pass of their own: they are a handful of
    // small, self-contained blocks, and threading them through the streaming
    // reader above would tangle them with the ordinary fonts and fills.
    ss.dxfs = parse_dxfs(xml, theme);
    Ok(ss)
}

/// Which side of a border is being read, so that a `<color>` child can be put
/// on the right one.
#[derive(Clone, Copy)]
enum Edge {
    Left,
    Right,
    Top,
    Bottom,
    Diagonal,
}

/// An edge that names no style is not drawn; one that does carries how.
fn border_line(e: &quick_xml::events::BytesStart) -> Option<BorderLine> {
    let style = get_attr(e, "style")?;
    if style == "none" {
        return None;
    }
    Some(BorderLine { style, color: None, color_reference: None })
}

/// Build a CellStyle from a style index referencing the StyleSheet.
fn resolve_cell_style(style_index: usize, stylesheet: &StyleSheet) -> CellStyle {
    let xf = match stylesheet.cell_xfs.get(style_index) {
        Some(xf) => xf,
        None => return CellStyle::default(),
    };
    let mut style=resolve_xf_style(xf,stylesheet);
    let normal=stylesheet.named_styles.iter().find(|(_,_,builtin)|*builtin==Some(0))
        .or_else(||stylesheet.named_styles.iter().find(|(name,_,_)|name.eq_ignore_ascii_case("Normal")))
        .map_or(0,|(_,index,_)|*index);
    let named=xf.format_reference.style_id.map(|id|id as usize).filter(|id|*id!=normal)
        .and_then(|id|stylesheet.named_styles.iter().find(|(_,index,_)|*index==id)
            .and_then(|(_,_,builtin)|stylesheet.cell_style_xfs.get(id).map(|parent|(parent,*builtin))));
    if let Some((parent,builtin))=named {
        let attached=xf.format_reference.font!=Some(true)&&xf.font_id==parent.font_id;
        style.named_font_link=Some(attached);
        // Numeric builtins exclude font and keep the cell's default font.
        // An included/custom named font can share resource0 while retaining
        // independent ownership across a later Normal edit.
        let named_owns_default=attached&&(builtin.is_none()||parent.format_reference.font!=Some(false));
        style.normal_font_link=Some(xf.font_id==0&&!named_owns_default);
    } else {style.normal_font_link=Some(xf.font_id==0);}
    style
}

fn resolve_xf_style(xf: &XfRecord, stylesheet: &StyleSheet) -> CellStyle {
    // Component identifiers select the stored definitions. Keep application
    // flags separately for later style edits and serialization.
    let font_id = xf.font_id;
    let fill_id = xf.fill_id;
    let border_id = xf.border_id;

    let mut font = stylesheet.fonts.get(font_id).cloned().unwrap_or_default();
    // A font record that names no face, or no size, is not asking for the
    // reader's own default: it is saying "the workbook's font, in bold" 遯ｶ繝ｻ
    // which is how a generated header row is written. The parts it leaves out
    // come from the Normal style's font.
    if font.name.is_none() || font.size.is_none() {
        let normal = stylesheet
            .cell_style_xfs
            .first()
            .map(|xf| xf.font_id)
            .and_then(|id| stylesheet.fonts.get(id));
        if let Some(normal) = normal {
            font.name = font.name.or_else(|| normal.name.clone());
            font.size = font.size.or(normal.size);
        }
    }
    if font.color.is_none() && font.reference.color_reference().is_none() {
        font.reference.implicit_automatic=true;
    }
    let fill = stylesheet.fills.get(fill_id).cloned().unwrap_or_default();
    let border = stylesheet
        .borders
        .get(border_id)
        .cloned()
        .unwrap_or_default();

    // Resolve number format
    let num_fmt_id = xf.num_fmt_id;
    let number_format = if num_fmt_id == 0 {
        None // General 遯ｶ繝ｻno explicit format needed
    } else if let Some(custom) = stylesheet.num_fmts.get(&num_fmt_id) {
        Some(custom.clone())
    } else {
        builtin_number_format(num_fmt_id).map(|s| s.to_string())
    };

    CellStyle {
        normal_font_link: None,
        named_font_link: None,
        format_reference: Some(Box::new(xf.format_reference.clone())),
        bold: font.bold,
        italic: font.italic,
        underline: font.underline,
        underline_kind: font.underline_kind,
        font_position: font.font_position,
        strikethrough: font.strikethrough,
        font_size: font.size,
        font_name: font.name.clone(),
        font_reference: (!font.reference.is_empty()).then(|| Box::new(font.reference.clone())),
        font_charset: font.charset,
        font_family: font.family,
        font_color: font.color,
        bg_color: fill.bg_color,
        pattern_fill: (fill.pattern_fill.pattern!="none" || fill.pattern_fill.foreground.is_some() || fill.pattern_fill.background.is_some()).then(||Box::new(fill.pattern_fill)),
        number_format,
        horizontal_align: xf.horizontal_align.clone(),
        indent: xf.indent,
        vertical_align: xf.vertical_align.clone(),
        wrap_text: xf.wrap_text,
        stacked_text: xf.stacked_text,
        text_rotation: xf.text_rotation,
        justify_last_line: xf.justify_last_line,
        shrink_to_fit: xf.shrink_to_fit,
        reading_order: xf.reading_order,
        quote_prefix: xf.quote_prefix,
        unlocked: xf.unlocked,
        formula_hidden: xf.formula_hidden,
        border_top: border.top.clone(),
        border_bottom: border.bottom.clone(),
        border_left: border.left.clone(),
        border_right: border.right.clone(),
        border_diagonal: border.diagonal.clone(),
        diagonal_up: border.up,
        diagonal_down: border.down,
    }
}

/// The font a row's own format puts on it, when the row carries one
/// (`s=` with customFormat="1").
fn row_style_font(
    e: &quick_xml::events::BytesStart,
    stylesheet: &StyleSheet,
) -> Option<(String, f32)> {
    if !is_true(get_attr(e, "customFormat").as_deref()) {
        return None;
    }
    let si = get_attr(e, "s")?.parse::<usize>().ok()?;
    let style = resolve_cell_style(si, stylesheet);
    Some((style.font_name?, style.font_size?))
}

/// Parse a single worksheet XML into a Sheet.
/// Where a relationship target lands, given the part that names it. Targets
/// are written relative to the naming part's own folder, so `../tables/t.xml`
/// beside `xl/worksheets/sheet1.xml` is `xl/tables/t.xml`.
fn part_beside(from: &str, target: &str) -> String {
    if target.starts_with('/') {
        return target.trim_start_matches('/').to_string();
    }
    let mut parts: Vec<&str> = from.rsplit_once('/').map_or("", |(dir, _)| dir).split('/').collect();
    parts.retain(|part| !part.is_empty());
    for step in target.split('/') {
        match step {
            "." | "" => {}
            ".." => {
                parts.pop();
            }
            step => parts.push(step),
        }
    }
    parts.join("/")
}

/// The ranges an `sqref` names, each as `(start_row, start_col, end_row,
/// end_col)` counted from zero. Excel writes them separated by spaces.
fn parse_sqref(sqref: &str) -> Vec<(u32, u32, u32, u32)> {
    sqref
        .split_whitespace()
        .map(|part| {
            let (first, last) = part.split_once(':').unwrap_or((part, part));
            let (start_col, start_row) = parse_cell_ref(first);
            let (end_col, end_row) = parse_cell_ref(last);
            (
                start_row.min(end_row),
                start_col.min(end_col),
                start_row.max(end_row),
                start_col.max(end_col),
            )
        })
        .collect()
}

/// A validation rule as its own element states it, before its formulas 遯ｶ繝ｻwhich
/// are child elements 遯ｶ繝ｻhave been read.
fn validation_from(e: &quick_xml::events::BytesStart) -> crate::ir::DataValidation {
    crate::ir::DataValidation {
        // No `type` at all is Excel's `any`, which allows everything.
        kind: get_attr(e, "type").unwrap_or_else(|| "any".to_string()),
        operator: get_attr(e, "operator"),
        ranges: get_attr(e, "sqref")
            .map(|held| parse_sqref(&held))
            .unwrap_or_default(),
        formula1: None,
        formula2: None,
        allow_blank: is_true(get_attr(e, "allowBlank").as_deref()),
        // `showDropDown` is written when the in-cell arrow was turned OFF, so
        // the flag is turned the right way round on the way in.
        in_cell_dropdown: !is_true(get_attr(e, "showDropDown").as_deref()),
        error_style: get_attr(e, "errorStyle"),
        show_input: is_true(get_attr(e, "showInputMessage").as_deref()),
        show_error: is_true(get_attr(e, "showErrorMessage").as_deref()),
        prompt_title: get_attr(e, "promptTitle"),
        error_title: get_attr(e, "errorTitle"),
        error: get_attr(e, "error"),
        prompt: get_attr(e, "prompt"),
    }
}

/// A conditional rule as its own element states it, before its formulas 遯ｶ繝ｻ
/// which are child elements 遯ｶ繝ｻhave been read. The look it puts on is resolved
/// here, because a `dxfId` means nothing once the stylesheet is out of reach.
fn conditional_from(
    e: &quick_xml::events::BytesStart,
    ranges: &[(u32, u32, u32, u32)],
    stylesheet: &StyleSheet,
) -> crate::ir::ConditionalRule {
    crate::ir::ConditionalRule {
        kind: get_attr(e, "type").unwrap_or_else(|| "expression".to_string()),
        operator: get_attr(e, "operator"),
        text: get_attr(e, "text"),
        rank: get_attr(e, "rank").and_then(|held| held.parse().ok()),
        bottom: is_true(get_attr(e, "bottom").as_deref()),
        percent: is_true(get_attr(e, "percent").as_deref()),
        above_average: get_attr(e, "aboveAverage").as_deref().map(|value| is_true(Some(value))).unwrap_or(true),
        equal_average: is_true(get_attr(e, "equalAverage").as_deref()),
        standard_deviations: get_attr(e, "stdDev").and_then(|held| held.parse().ok()),
        formulas: Vec::new(),
        color_scale_points: None,
        visual_xml: None,
        visual_extension_id: None,
        visual_extension_xml: None,
        visual_extension_only: false,
        priority: get_attr(e, "priority")
            .and_then(|held| held.parse().ok())
            .unwrap_or(0),
        origin_priority: get_attr(e, "priority").and_then(|held| held.parse().ok()),
        stop_if_true: is_true(get_attr(e, "stopIfTrue").as_deref()),
        ranges: ranges.to_vec(),
        style: get_attr(e, "dxfId")
            .and_then(|held| held.parse::<usize>().ok())
            .and_then(|at| stylesheet.dxfs.get(at))
            .cloned(),
    }
}

/// The buttons and tick boxes a sheet's legacy drawing carries.
///
/// The same part holds the notes, which is why the two are read together: a
/// sheet can name a legacy drawing and have no notes in it at all.
fn parse_form_controls(vml: &str) -> Vec<crate::ir::FormControl> {
    use crate::ir::{Anchor, FormControl};
    let mut held = Vec::new();
    // The space matters, the same way it does for notes: `<v:shapetype>`
    // starts with `<v:shape` too.
    for shape in vml.split("<v:shape ").skip(1) {
        let shape = shape.split("</v:shape>").next().unwrap_or(shape);
        let Some(kind) = shape
            .split("ObjectType=\"")
            .nth(1)
            .and_then(|rest| rest.split('"').next())
        else {
            continue;
        };
        // A note is the one thing in here that is not a control, and it is
        // read by the note parser instead.
        if kind == "Note" {
            continue;
        }
        let tagged = |name: &str| -> Option<String> {
            let open = format!("<x:{name}>");
            let at = shape.find(&open)? + open.len();
            let rest = &shape[at..];
            Some(rest[..rest.find('<')?].trim().to_string())
        };
        let Some(anchor) = tagged("Anchor") else { continue };
        let numbers: Vec<i64> = anchor
            .split(',')
            .filter_map(|part| part.trim().parse().ok())
            .collect();
        let [left, dx, top, dy, right, dx2, bottom, dy2] = numbers[..] else {
            continue;
        };
        let corner = |column: i64, x: i64, row: i64, y: i64| Anchor {
            col: column.max(0) as u32,
            col_off: x * 9525,
            row: row.max(0) as u32,
            row_off: y * 9525,
        };
        // The caption is HTML inside the shape's text box: a div, sometimes a
        // font, and the words between them.
        let text = shape
            .split("<v:textbox")
            .nth(1)
            .and_then(|rest| rest.split("</v:textbox>").next())
            .map(|held| {
                let mut out = String::new();
                let mut inside_tag = false;
                for letter in held.chars() {
                    match letter {
                        '<' => inside_tag = true,
                        '>' => inside_tag = false,
                        _ if !inside_tag => out.push(letter),
                        _ => {}
                    }
                }
                out.split_whitespace().collect::<Vec<_>>().join(" ")
            })
            .filter(|held| !held.is_empty());
        held.push(FormControl {
            kind: kind.to_string(),
            from: corner(left, dx, top, dy),
            to: Some(corner(right, dx2, bottom, dy2)),
            text,
            checked: tagged("Checked").is_some_and(|held| held != "0"),
        });
    }
    held
}

/// The links a sheet's cells carry, resolved against the sheet's own
/// relationships.
///
/// A link that leaves the workbook keeps its URL in a relationship and names
/// it by id; one that stays inside names its place outright in `location`.
/// Neither reaches a cell's style, which is why a sheet full of blue
/// underlined text can still have nowhere to go.
fn parse_hyperlinks(
    sheet_xml: &str,
    rels: &std::collections::HashMap<String, oxidocs_common::relationships::Relationship>,
) -> Vec<crate::ir::Hyperlink> {
    let mut held = Vec::new();
    let mut reader = Reader::from_str(sheet_xml);
    loop {
        match reader.read_event() {
            Ok(Event::Start(ref e)) | Ok(Event::Empty(ref e)) => {
                if local_name(e.name().as_ref()) != "hyperlink" {
                    continue;
                }
                let Some(at) = get_attr(e, "ref") else { continue };
                // Excel may put one hyperlink on a whole cell block. Keep both
                // corners so the VBA host can return the original Range.
                let first = at.split(':').next().unwrap_or(&at);
                let (col, row) = parse_cell_ref(first);
                let end_cell = at.split_once(':').map(|(_, last)| {
                    let (end_col, end_row) = parse_cell_ref(last);
                    (end_row, end_col)
                }).filter(|end| *end != (row, col));
                let location = get_attr(e, "location");
                let relationship = get_attr(e, "r:id")
                    .or_else(|| get_attr(e, "id"))
                    .and_then(|id| rels.get(&id));
                let internal = relationship.is_none();
                let target = if internal {
                    let Some(place) = location.as_ref() else { continue };
                    place.clone()
                } else {
                    relationship.unwrap().target.clone()
                };
                held.push(crate::ir::Hyperlink {
                    cell: (row, col),
                    end_cell,
                    target,
                    internal,
                    sub_address: (!internal).then_some(location).flatten(),
                    tooltip: get_attr(e, "tooltip"),
                });
            }
            Ok(Event::Eof) | Err(_) => break,
            _ => {}
        }
    }
    held
}

/// Reads one `xl/tables/*.xml` part into the range and dress it describes.
/// What a sheet's drawing part holds, with each picture's relationship id.
///
/// The anchors come in three kinds: hung from two cells, hung from one with a
/// size of its own, or placed at a fixed spot on the sheet. Only the first two
/// are common in the corpus, and an absolute one is read as a one-cell anchor
/// at the top-left, which is where its offsets are measured from anyway.
#[derive(Debug, Clone, Default, PartialEq, Eq)]
struct DrawingImageRelationships {
    raster: Option<String>,
    vector: Option<String>,
    shape_id: Option<u32>,
}

fn svg_intrinsic_size_points(svg: &[u8]) -> Option<(f64, f64)> {
    fn svg_length_points(value: &str) -> Option<f64> {
        let value = value.trim();
        let numeric_end = value
            .find(|ch: char| !(ch.is_ascii_digit() || matches!(ch, '+' | '-' | '.' | 'e' | 'E')))
            .unwrap_or(value.len());
        let number = value.get(..numeric_end)?.parse::<f64>().ok()?;
        let unit = value.get(numeric_end..)?.trim();
        let points = match unit {
            "" | "px" => number * 72.0 / 96.0,
            "pt" => number,
            "pc" => number * 12.0,
            "in" => number * 72.0,
            "cm" => number * 72.0 / 2.54,
            "mm" => number * 72.0 / 25.4,
            _ => return None,
        };
        (points.is_finite() && points > 0.0).then_some(points)
    }

    let mut reader = Reader::from_reader(svg);
    reader.config_mut().trim_text(true);
    let mut buf = Vec::new();
    loop {
        match reader.read_event_into(&mut buf).ok()? {
            Event::Start(element) | Event::Empty(element)
                if local_name(element.name().as_ref()) == "svg" =>
            {
                let mut width = None;
                let mut height = None;
                let mut view_box = None;
                for attribute in element.attributes().flatten() {
                    match local_name(attribute.key.as_ref()).as_str() {
                        "width" => width = std::str::from_utf8(&attribute.value).ok().and_then(svg_length_points),
                        "height" => height = std::str::from_utf8(&attribute.value).ok().and_then(svg_length_points),
                        "viewBox" => {
                            let values = std::str::from_utf8(&attribute.value).ok()?.split([',', ' ', '\t', '\r', '\n'])
                                .filter(|item| !item.is_empty()).map(str::parse::<f64>).collect::<Result<Vec<_>, _>>().ok()?;
                            if values.len() == 4 && values[2].is_finite() && values[3].is_finite() && values[2] > 0.0 && values[3] > 0.0 {
                                view_box = Some((values[2], values[3]));
                            }
                        }
                        _ => {}
                    }
                }
                if let Some((view_width, view_height)) = view_box {
                    match (width, height) {
                        (Some(width), Some(height)) => return Some((width, height)),
                        (Some(width), None) => return Some((width, width * view_height / view_width)),
                        (None, Some(height)) => return Some((height * view_width / view_height, height)),
                        (None, None) => return Some((view_width * 72.0 / 96.0, view_height * 72.0 / 96.0)),
                    }
                }
                return Some((width?, height?));
            }
            Event::Eof => return None,
            _ => {}
        }
        buf.clear();
    }
}

#[derive(Clone)]
struct OleObjectMetadata {
    prog_id: String,
    display_as_icon: bool,
    placement: i64,
    ole_type: i64,
    auto_load: Option<bool>,
    enabled: Option<bool>,
    locked: Option<bool>,
    print_object: Option<bool>,
    on_action: Option<String>,
}

struct PendingOleObject {
    id: u32,
    prog_id: String,
    relationship_id: Option<String>,
    display_as_icon: bool,
    auto_load: Option<bool>,
    move_with_cells: bool,
    size_with_cells: bool,
    enabled: Option<bool>,
    locked: Option<bool>,
    print_object: Option<bool>,
    on_action: Option<String>,
}

fn xml_bool(value: Option<String>) -> Option<bool> {
    match value.as_deref() {
        Some("1" | "true" | "True") => Some(true),
        Some("0" | "false" | "False") => Some(false),
        _ => None,
    }
}

fn parse_ole_object_metadata(xml: &str, relationships: &HashMap<String, Relationship>) -> HashMap<u32, OleObjectMetadata> {
    let mut result = HashMap::new();
    let mut reader = Reader::from_str(xml);
    let mut buf = Vec::new();
    let mut current: Option<PendingOleObject> = None;
    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(element)) if local_name(element.name().as_ref()) == "oleObject" => {
                if let (Some(id), Some(prog_id)) = (
                    get_attr(&element, "shapeId").and_then(|value| value.parse::<u32>().ok()),
                    get_attr(&element, "progId"),
                ) {
                    current = Some(PendingOleObject {
                        id,
                        prog_id,
                        relationship_id: get_attr(&element, "id"),
                        display_as_icon: get_attr(&element, "dvAspect").as_deref() == Some("DVASPECT_ICON"),
                        auto_load: xml_bool(get_attr(&element, "autoLoad")),
                        move_with_cells: false,
                        size_with_cells: false,
                        enabled: None,
                        locked: None,
                        print_object: None,
                        on_action: None,
                    });
                }
            }
            Ok(Event::Empty(element)) if local_name(element.name().as_ref()) == "oleObject" => {
                if let (Some(id), Some(prog_id)) = (
                    get_attr(&element, "shapeId").and_then(|value| value.parse::<u32>().ok()),
                    get_attr(&element, "progId"),
                ) {
                    let metadata = OleObjectMetadata {
                        prog_id,
                        display_as_icon: get_attr(&element, "dvAspect").as_deref() == Some("DVASPECT_ICON"),
                        placement: 3,
                        ole_type: get_attr(&element, "id").and_then(|rid| relationships.get(&rid))
                            .is_some_and(|relationship| relationship.target_mode.as_deref() == Some("External")) as i64 + 1,
                        auto_load: xml_bool(get_attr(&element, "autoLoad")),
                        enabled: None,
                        locked: None,
                        print_object: None,
                        on_action: None,
                    };
                    result.entry(id).or_insert(metadata);
                }
            }
            Ok(Event::Start(element) | Event::Empty(element)) if local_name(element.name().as_ref()) == "anchor" => {
                if let Some(pending) = current.as_mut() {
                    pending.move_with_cells = xml_bool(get_attr(&element, "moveWithCells")).unwrap_or(false);
                    pending.size_with_cells = xml_bool(get_attr(&element, "sizeWithCells")).unwrap_or(false);
                }
            }
            Ok(Event::Start(element) | Event::Empty(element)) if local_name(element.name().as_ref()) == "objectPr" => {
                if let Some(pending) = current.as_mut() {
                    pending.enabled = xml_bool(get_attr(&element, "disabled")).map(|disabled| !disabled);
                    pending.locked = xml_bool(get_attr(&element, "locked"));
                    pending.print_object = xml_bool(get_attr(&element, "print"));
                    pending.on_action = get_attr(&element, "macro");
                }
            }
            Ok(Event::End(element)) if local_name(element.name().as_ref()) == "oleObject" => {
                if let Some(pending) = current.take() {
                    let placement = if pending.size_with_cells { 1 } else if pending.move_with_cells { 2 } else { 3 };
                    let ole_type = pending.relationship_id.as_ref().and_then(|rid| relationships.get(rid))
                        .is_some_and(|relationship| relationship.target_mode.as_deref() == Some("External")) as i64 + 1;
                    result.insert(pending.id, OleObjectMetadata {
                        prog_id: pending.prog_id,
                        display_as_icon: pending.display_as_icon,
                        placement,
                        ole_type,
                        auto_load: pending.auto_load,
                        enabled: pending.enabled,
                        locked: pending.locked,
                        print_object: pending.print_object,
                        on_action: pending.on_action,
                    });
                }
            }
            Ok(Event::Eof) | Err(_) => break,
            _ => {}
        }
        buf.clear();
    }
    result
}

#[derive(Default, Clone)]
struct VmlOleMetadata {
    shadow: Option<bool>,
    border_color: Option<String>,
    border_weight: Option<i64>,
    interior_color: Option<String>,
}

fn vml_rgb(value: &str) -> Option<String> {
    let value = value.trim();
    if let Some(hex) = value.strip_prefix('#').filter(|hex| hex.len() == 6 && hex.chars().all(|c| c.is_ascii_hexdigit())) {
        return Some(hex.to_ascii_uppercase());
    }
    let name = value.split_whitespace().next()?.to_ascii_lowercase();
    Some(match name.as_str() {
        "black" => "000000", "white" => "FFFFFF", "red" => "FF0000", "green" => "008000",
        "lime" => "00FF00", "blue" => "0000FF", "yellow" => "FFFF00", "cyan" => "00FFFF",
        "magenta" => "FF00FF", "gray" | "grey" => "808080", "maroon" => "800000",
        "navy" => "000080", "olive" => "808000", "purple" => "800080", "teal" => "008080",
        _ => return None,
    }.to_string())
}

fn parse_vml_ole_shadows(xml: &str) -> HashMap<u32, VmlOleMetadata> {
    let mut result = HashMap::new();
    let mut reader = Reader::from_str(xml);
    let mut buf = Vec::new();
    let mut current_shape = None;
    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(element)) if local_name(element.name().as_ref()) == "shape" => {
                current_shape = get_attr(&element, "id").and_then(|name| name.rsplit_once("_s")?.1.parse::<u32>().ok());
                if let Some(id) = current_shape {
                    let border_weight = get_attr(&element, "strokeweight").and_then(|weight| {
                        let number = weight.trim().trim_end_matches("pt").parse::<f64>().ok()?;
                        Some(if number >= 1.75 { 4 } else if number >= 1.1 { -4138 } else if number >= 0.5 { 2 } else { 1 })
                    });
                    result.insert(id, VmlOleMetadata {
                        shadow: Some(false),
                        border_color: get_attr(&element, "strokecolor").as_deref().and_then(vml_rgb),
                        border_weight,
                        interior_color: get_attr(&element, "fillcolor").as_deref().and_then(vml_rgb),
                    });
                }
            }
            Ok(Event::Empty(element)) if local_name(element.name().as_ref()) == "shape" => {
                if let Some(id) = get_attr(&element, "id").and_then(|name| name.rsplit_once("_s")?.1.parse::<u32>().ok()) {
                    result.insert(id, VmlOleMetadata { shadow: Some(false), ..VmlOleMetadata::default() });
                }
            }
            Ok(Event::Start(element) | Event::Empty(element)) if local_name(element.name().as_ref()) == "shadow" => {
                if let Some(id) = current_shape {
                    let enabled = xml_bool(get_attr(&element, "on")).unwrap_or(true);
                    result.entry(id).or_insert_with(VmlOleMetadata::default).shadow = Some(enabled);
                }
            }
            Ok(Event::End(element)) if local_name(element.name().as_ref()) == "shape" => current_shape = None,
            Ok(Event::Eof) | Err(_) => break,
            _ => {}
        }
        buf.clear();
    }
    result
}

/// XML normalizes literal attribute whitespace before expanding references.
/// Drawing descriptions are plain XML strings, not OOXML escaped cell text.
fn drawing_metadata_attribute(element: &quick_xml::events::BytesStart<'_>, key: &[u8]) -> Option<String> {
    let attribute = element.attributes().flatten().find(|attribute| attribute.key.as_ref() == key)?;
    let raw = std::str::from_utf8(attribute.value.as_ref()).ok()?;
    let normalized = raw.replace("\r\n", "\n").chars().map(|ch| if matches!(ch, '\t' | '\n' | '\r') { ' ' } else { ch }).collect::<String>();
    quick_xml::escape::unescape(&normalized).ok().map(|text| text.into_owned())
}

fn parse_drawing_xml(xml: &str, theme: &Theme) -> Vec<(crate::ir::Drawing, DrawingImageRelationships)> {
    use crate::ir::{
        Anchor, Drawing, DrawingKind, Frame, PathStep, Shape, ShapeLine, ShapeParagraph,
        ShapePath, ShapeRun, ShapeText,
    };

    /// Which part of the shape a colour being read belongs to.
    #[derive(Clone, Copy, PartialEq)]
    enum Paints {
        Fill,
        Line,
        LinePatternForeground,
        LinePatternBackground,
        Text,
    }

    let mut drawing_name=None;
    let mut drawing_id: Option<u32> = None;
    let mut drawing_metadata: Option<crate::ir::DrawingMetadata> = None;
    let mut group_name: Option<String> = None;
    let mut group_non_visual_id: Option<u32> = None;
    let mut reader = quick_xml::reader::NsReader::from_str(xml);
    let mut buf = Vec::new();
    let mut found: Vec<(Drawing, DrawingImageRelationships)> = Vec::new();
    // Where a chart's own shape sits, as fractions of the chart's box.
    let mut frame: Option<(f64, f64)> = None;
    let mut frame_to: Option<(f64, f64)> = None;

    // Whether the geometry being read is a preset, so its adjust handles are
    // told apart from a custom geometry's formula list.
    let mut in_preset = false;

    // The outline a shape draws itself with, and the step being read. A
    // verb's points are children of it, so which verb is open says what to do
    // with the points that arrive.
    let mut path: Option<ShapePath> = None;
    let mut step: Option<String> = None;
    let mut points: Vec<(i64, i64)> = Vec::new();

    // Explicit properties on the run being read. Missing values inherit the
    // paragraph style, while explicit false values remain distinct.
    let mut run_bold = false;
    let mut run_italic: Option<bool> = None;
    let mut run_underline = false;
    let mut run_color: Option<String> = None;
    let mut run_color_theme: Option<usize> = None;
    let mut run_color_brightness: Option<f32> = None;
    let mut run_legacy_color_index: Option<u8> = None;
    let mut run_face: Option<String> = None;
    let mut run_size: Option<f32> = None;

    let blank = Anchor { col: 0, col_off: 0, row: 0, row_off: 0 };
    let mut from = blank;
    let mut to: Option<Anchor> = None;
    let mut extent: Option<(i64, i64)> = None;
    let mut kind: Option<DrawingKind> = None;
    let blank_shape = Shape {
        transform_offset: None,
        geometry: String::new(),
        text_box: false,
        adjusts: Vec::new(),
        path: None,
        fill: None,
        fill_theme: None,
        fill_legacy_color_index: None,
        fill_brightness: None,
        line: None,
        flip_h: false,
        flip_v: false,
        rotation: 0,
        text: None,
    };
    let mut shape = blank_shape.clone();
    let mut line_width: i64 = 9525;
    let mut width_stated = false;
    let mut line_cap: Option<String> = None;
    let mut dash: Option<String> = None;
    let (mut head_end, mut tail_end): (Option<String>, Option<String>) = (None, None);
    let (
        mut head_end_length,
        mut head_end_width,
        mut tail_end_length,
        mut tail_end_width,
    ): (Option<String>, Option<String>, Option<String>, Option<String>) = (None, None, None, None);
    let mut embed: Option<String> = None;
    let mut vector_embed: Option<String> = None;
    let (mut in_picture_clr_change, mut in_picture_clr_from) = (false, false);
    // Which corner the col/row elements belong to, and whether we are inside
    // a shape 遯ｶ繝ｻwhere `<ext>` is the shape's own size, not the anchor's.
    let mut corner: Option<bool> = None;
    let mut depth_in_shape = 0usize;
    let mut number = String::new();
    // Where a colour is being read from, and what is being read into.
    let (mut in_sp_pr, mut in_ln, mut in_style) = (false, false, false);
    // An extension list holds what a shape would be painted with if it were
    // painted at all: Excel writes the fill a shape had before it was set to
    // none into an `a14:hiddenFill`, and reading that paints white boxes over
    // the sheet. Nothing inside one is a shape's own paint.
    let mut in_ext_lst = 0usize;
    // Whether the shape itself said what it is painted and ruled with, in
    // which case its style's references are not consulted.
    let (mut fill_stated, mut line_stated) = (false, false);
    let mut paints: Option<Paints> = None;
    let mut colour: Option<(String, Vec<(String, f32)>, Option<usize>, Option<u8>)> = None;
    // What the shape says: the body's own settings, the paragraph being read,
    // and whether a run's text is what the reader is collecting.
    let mut said = ShapeText {
        paragraphs: Vec::new(),
        anchor: None,
        insets: (91440, 45720, 91440, 45720),
        wrap: true,
        clip: false,
    };
    let mut paragraph: Option<ShapeParagraph> = None;
    let mut in_run_props = false;
    let mut in_text = false;
    let mut in_line_spacing = false;
    let mut text = String::new();
    // A group states where it sits and what its children's coordinates mean;
    // a child inside one is placed by mapping its own box through that.
    let mut group: Option<((i64, i64), (i64, i64), (i64, i64), (i64, i64))> = None;
    // Nested DrawingML groups each add another child-to-parent coordinate
    // transform. Excel exposes the leaves as one flat GroupItems collection.
    let mut parent_groups: Vec<((i64, i64), (i64, i64), (i64, i64), (i64, i64))> = Vec::new();
    let mut parent_group_rotations: Vec<i32> = Vec::new();
    let mut current_group_rotation = 0i32;
    let mut group_rotation = 0i32;
    let mut in_group_props = false;
    let mut own_off: Option<(i64, i64)> = None;
    let mut own_ext: Option<(i64, i64)> = None;

    loop {
        let event = reader.read_event_into(&mut buf);
        let (start, empty) = match &event {
            Ok(Event::Start(e)) => (Some(e), false),
            Ok(Event::Empty(e)) => (Some(e), true),
            _ => (None, false),
        };
        if let Some(e) = start {
            let name = local_name(e.name().as_ref());
            match name.as_str() {
                "srcRect" => {
                    if let Some(DrawingKind::Picture { format, .. }) = kind.as_mut() {
                        let read = |attribute: &str| get_attr(e, attribute).and_then(|value| value.parse::<i32>().ok()).unwrap_or(0);
                        format.crop = (read("l"), read("t"), read("r"), read("b"));
                    }
                }
                "lum" => {
                    if let Some(DrawingKind::Picture { format, .. }) = kind.as_mut() {
                        format.brightness = get_attr(e, "bright").and_then(|value| value.parse().ok()).unwrap_or(0);
                        format.contrast = get_attr(e, "contrast").and_then(|value| value.parse().ok()).unwrap_or(0);
                        if format.brightness == 70_000 && format.contrast == -70_000 {
                            format.color_type = 4;
                        }
                    }
                }
                "grayscl" => {
                    if let Some(DrawingKind::Picture { format, .. }) = kind.as_mut() { format.color_type = 2; }
                }
                "biLevel" => {
                    if let Some(DrawingKind::Picture { format, .. }) = kind.as_mut() { format.color_type = 3; }
                }
                "clrChange" => {
                    if let Some(DrawingKind::Picture { .. }) = kind.as_mut() {
                        in_picture_clr_change = true;
                    }
                }
                "clrFrom" if in_picture_clr_change => in_picture_clr_from = true,
                "srgbClr" if in_picture_clr_from => {
                    if let Some(DrawingKind::Picture { format, .. }) = kind.as_mut() {
                        format.transparency_color = get_attr(e, "val");
                    }
                }
                "alpha" if in_sp_pr && get_attr(e, "val").as_deref() == Some("0") => {
                    if let Some(DrawingKind::Picture { format, .. }) = kind.as_mut() { format.transparent_background = true; }
                }
                "cNvPr" if matches!(reader.resolve_element(e.name()).0,
                    quick_xml::name::ResolveResult::Bound(namespace) if matches!(namespace.as_ref(),
                    b"http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing" |
                    b"http://purl.oclc.org/ooxml/drawingml/spreadsheetDrawing" |
                    b"http://schemas.openxmlformats.org/drawingml/2006/chartDrawing" |
                    b"http://purl.oclc.org/ooxml/drawingml/chartDrawing")) => {
                    let metadata = crate::ir::DrawingMetadata {
                        description: drawing_metadata_attribute(e, b"descr"), utf16_description: None,
                        title: drawing_metadata_attribute(e, b"title"), hidden: drawing_metadata_attribute(e, b"hidden"),
                    };
                    drawing_metadata = (!metadata.is_empty()).then_some(metadata);
                    let name = get_attr(e, "name");
                    drawing_id = get_attr(e, "id").and_then(|value| value.parse().ok());
                    if depth_in_shape == 1 && group.is_none() && parent_groups.is_empty() {
                        // Excel exposes nested group leaves through the outer
                        // group's flat GroupItems collection and ParentGroup.
                        group_name = name.clone();
                        group_non_visual_id = drawing_id;
                    }
                    drawing_name = name;
                }
                "cNvSpPr" if get_attr(e, "txBox").as_deref() == Some("1") => shape.text_box = true,
                "twoCellAnchor" | "oneCellAnchor" | "absoluteAnchor"
                // A shape that hangs on a chart rather than on the grid
                // is anchored by fractions of the chart's own box.
                | "relSizeAnchor" | "absSizeAnchor" => {
                    frame = None;
                    frame_to = None;
                    drawing_metadata = None;
                    group_name = None;
                    group_non_visual_id = None;
                    group_rotation = 0;
                    parent_groups.clear();
                    parent_group_rotations.clear();
                    current_group_rotation = 0;
                    from = blank;
                    to = None;
                    extent = None;
                    kind = None;
                    embed = None;
                    vector_embed = None;
                    in_picture_clr_change = false;
                    in_picture_clr_from = false;
                    depth_in_shape = 0;
                    shape = blank_shape.clone();
                    line_width = 9525;
                    width_stated = false;
                    dash = None;
                    head_end = None;
                    tail_end = None;
                    head_end_length = None;
                    head_end_width = None;
                    tail_end_length = None;
                    tail_end_width = None;
                    fill_stated = false;
                    line_stated = false;
                    said = ShapeText {
                        paragraphs: Vec::new(),
                        anchor: None,
                        insets: (91440, 45720, 91440, 45720),
                        wrap: true,
                        clip: false,
                    };
                    paragraph = None;
                    group = None;
                    own_off = None;
                    own_ext = None;
                }
                "from" => corner = Some(true),
                "to" => {
                    corner = Some(false);
                    to = Some(blank);
                }
                "col" | "colOff" | "row" | "rowOff" => number.clear(),
                "x" | "y" if corner.is_some() => number.clear(),
                "ext" => {
                    let cx = get_attr(e, "cx").and_then(|v| v.parse().ok()).unwrap_or(0);
                    let cy = get_attr(e, "cy").and_then(|v| v.parse().ok()).unwrap_or(0);
                    if depth_in_shape == 0 && !in_group_props {
                        extent = Some((cx, cy));
                    } else if in_group_props {
                        group.get_or_insert(((0, 0), (0, 0), (0, 0), (0, 0))).1 = (cx, cy);
                    } else if in_sp_pr {
                        own_ext = Some((cx, cy));
                    }
                }
                "pic" | "sp" | "cxnSp" => {
                    depth_in_shape += 1;
                    // A child of a group starts its own shape; one that is
                    // not carries the anchor's.
                    if group.is_some() {
                        shape = blank_shape.clone();
                        said.paragraphs.clear();
                        paragraph = None;
                        fill_stated = false;
                        line_stated = false;
                        head_end = None;
                        tail_end = None;
                        head_end_length = None;
                        head_end_width = None;
                        tail_end_length = None;
                        tail_end_width = None;
                        own_off = None;
                        own_ext = None;
                        embed = None;
                        vector_embed = None;
                    }
                    if name == "pic" {
                        kind = Some(DrawingKind::Picture { bytes: Vec::new(), format: Default::default() });
                    } else if kind.is_none() || group.is_some() {
                        kind = Some(DrawingKind::Shape(shape.clone()));
                    }
                }
                "graphicFrame" => {
                    depth_in_shape += 1;
                    kind = Some(DrawingKind::Chart(Default::default()));
                }
                // The frame holds no picture of its own: it names the part
                // that does, the way a picture names its bytes.
                "chart" if embed.is_none() => {
                    embed = get_attr(e, "id");
                }
                "grpSp" => {
                    depth_in_shape += 1;
                    if let Some(parent) = group.take() {
                        parent_groups.push(parent);
                        parent_group_rotations.push(current_group_rotation);
                    } else {
                        current_group_rotation = 0;
                    }
                    if kind.is_none() {
                        kind = Some(DrawingKind::Other);
                    }
                }
                "grpSpPr" => in_group_props = true,
                "xfrm" if in_group_props => {
                    let stated_rotation = get_attr(e, "rot")
                        .and_then(|held| held.parse().ok())
                        .unwrap_or(0);
                    // Shape.Rotation and child geometry follow the outermost
                    // group that Excel exposes as the parent. Nested groups
                    // are flattened by Excel's GroupItems collection.
                    if depth_in_shape == 1 {
                        group_rotation = stated_rotation;
                    }
                    current_group_rotation = stated_rotation;
                }
                "chOff" | "chExt" | "off" if in_group_props || in_sp_pr => {
                    let x = get_attr(e, "x").or_else(|| get_attr(e, "cx"));
                    let y = get_attr(e, "y").or_else(|| get_attr(e, "cy"));
                    let pair = (
                        x.and_then(|v| v.parse().ok()).unwrap_or(0),
                        y.and_then(|v| v.parse().ok()).unwrap_or(0),
                    );
                    if in_group_props {
                        let held = group.get_or_insert(((0, 0), (0, 0), (0, 0), (0, 0)));
                        match name.as_str() {
                            "off" => held.0 = pair,
                            "chOff" => held.2 = pair,
                            _ => held.3 = pair,
                        }
                    } else if name == "off" {
                        own_off = Some(pair);
                    }
                }
                "spPr" => in_sp_pr = true,
                "style" => in_style = true,
                "extLst" => in_ext_lst += 1,
                "ln" if in_sp_pr && in_ext_lst == 0 => {
                    in_ln = true;
                    if let Some(stated) = get_attr(e, "w").and_then(|w| w.parse().ok()) {
                        line_width = stated;
                        width_stated = true;
                    }
                    line_cap = get_attr(e, "cap");
                }
                "xfrm" if in_sp_pr => {
                    shape.flip_h = is_true(get_attr(e, "flipH").as_deref());
                    shape.flip_v = is_true(get_attr(e, "flipV").as_deref());
                    shape.rotation = get_attr(e, "rot")
                        .and_then(|held| held.parse().ok())
                        .unwrap_or(0);
                }
                // An adjust handle, which a preset reads its own shape from.
                // Only the ones inside a `prstGeom` count: a `custGeom` states
                // a `gdLst` full of formulas under the same element name.
                "gd" if in_preset => {
                    if let (Some(name), Some(value)) = (
                        get_attr(e, "name"),
                        get_attr(e, "fmla")
                            .and_then(|held| held.strip_prefix("val ").map(str::to_string))
                            .and_then(|held| held.trim().parse::<i64>().ok()),
                    ) {
                        shape.adjusts.push((name, value));
                    }
                }
                "path" => {
                    let number = |name: &str| {
                        get_attr(e, name).and_then(|held| held.parse::<i64>().ok()).unwrap_or(0)
                    };
                    let (across, down) = (number("w"), number("h"));
                    if across > 0 && down > 0 {
                        path = Some(ShapePath { across, down, steps: Vec::new() });
                    }
                }
                "moveTo" | "lnTo" | "cubicBezTo" if path.is_some() => {
                    step = Some(name.clone());
                    points.clear();
                }
                "close" if path.is_some() => {
                    if let Some(held) = path.as_mut() {
                        held.steps.push(PathStep::Close);
                    }
                }
                "pt" if step.is_some() => {
                    let number = |name: &str| {
                        get_attr(e, name).and_then(|held| held.parse::<i64>().ok()).unwrap_or(0)
                    };
                    points.push((number("x"), number("y")));
                }
                "prstGeom" => {
                    in_preset = true;
                    if let Some(preset) = get_attr(e, "prst") {
                        shape.geometry = preset;
                    }
                }
                "prstDash" if in_ln => dash = get_attr(e, "val"),
                "pattFill" if in_ln => {
                    if let Some(line) = shape.line.as_mut() {
                        line.pattern = get_attr(e, "prst");
                    } else {
                        shape.line = Some(ShapeLine {
                            color: "000000".into(),
                            color_theme: None,
                            legacy_color_index: None,
                            color_brightness: None, color_shade: None,
                            width: line_width,
                            dash: dash.clone(),
                            cap: line_cap.clone(),
                            head_end: None,
                            tail_end: None,
                            head_end_length: None,
                            head_end_width: None,
                            tail_end_length: None,
                            tail_end_width: None,
                            pattern: get_attr(e, "prst"),
                            back_color: None,
                            back_color_theme: None,
                            back_legacy_color_index: None,
                            back_color_brightness: None, back_color_shade: None,
                        });
                    }
                }
                "fgClr" if in_ln => paints = Some(Paints::LinePatternForeground),
                "bgClr" if in_ln => paints = Some(Paints::LinePatternBackground),
                // What the rule wears at its ends. `type="none"` is most of
                // them and means bare.
                "headEnd" | "tailEnd" if in_ln => {
                    let worn = get_attr(e, "type").filter(|kind| kind != "none");
                    if name == "headEnd" {
                        head_end = worn;
                        head_end_length = get_attr(e, "len");
                        head_end_width = get_attr(e, "w");
                    } else {
                        tail_end = worn;
                        tail_end_length = get_attr(e, "len");
                        tail_end_width = get_attr(e, "w");
                    }
                }
                "noFill" if in_sp_pr && in_ext_lst == 0 => {
                    if in_ln {
                        line_stated = true;
                        shape.line = None;
                    } else {
                        fill_stated = true;
                        shape.fill = None;
                    }
                }
                // A shape that states no fill or line of its own wears the
                // one its style names, which is a theme colour.
                "solidFill" if in_sp_pr && in_ext_lst == 0 => {
                    if let Some(DrawingKind::Picture { format, .. }) = kind.as_mut() {
                        format.transparent_background = false;
                    }
                    if in_ln {
                        line_stated = true;
                        paints = Some(Paints::Line);
                    } else {
                        fill_stated = true;
                        paints = Some(Paints::Fill);
                    }
                }
                // A style reference at index 0 is the theme's empty slot: it
                // names a colour but asks for nothing to be drawn with it.
                // Excel leaves such a shape unpainted or unruled, and the
                // corpus carries 559 `fillRef idx="0"` and 49 `lnRef idx="0"`.
                // A freeform built through Excel's own BuildFreeform with its
                // fill turned off states no fill at all in `spPr` and says so
                // only this way 遯ｶ繝ｻdrawn from the reference regardless, it came
                // out filled in the theme's accent blue where Excel draws bare
                // strokes (`_xlsx_custgeom_repro.py`).
                "fillRef" if in_style && !fill_stated => {
                    paints = (get_attr(e, "idx").as_deref() != Some("0"))
                        .then_some(Paints::Fill);
                }
                "lnRef" if in_style && !line_stated => {
                    paints = (get_attr(e, "idx").as_deref() != Some("0"))
                        .then_some(Paints::Line);
                    // And the width the theme draws that slot at, where the
                    // shape has not stated one. `bunya_taikeizu_point`'s
                    // headings name slot 2 of a theme whose styles are 0.75,
                    // 2 and 3 point, and Excel rules them three pixels wide
                    // where the 0.75pt default rules one.
                    if !width_stated {
                        if let Some(slot) = get_attr(e, "idx")
                            .and_then(|idx| idx.parse::<usize>().ok())
                            .filter(|idx| *idx > 0)
                            .and_then(|idx| theme.rules.get(idx - 1))
                        {
                            line_width = *slot;
                        }
                    }
                }
                "solidFill" if in_run_props => paints = Some(Paints::Text),
                "srgbClr" | "schemeClr" | "sysClr" if paints.is_some() => {
                    if in_run_props && name == "schemeClr" && paints == Some(Paints::Text) {
                        run_color_theme = get_attr(e, "val").and_then(|val| drawing_theme_index(&val));
                    }
                    if in_run_props && name == "srgbClr" && paints == Some(Paints::Text) {
                        run_legacy_color_index = get_attr(e, "legacySpreadsheetColorIndex")
                            .and_then(|value| value.parse::<u8>().ok());
                    }
                    let named = match name.as_str() {
                        "sysClr" => get_attr(e, "lastClr"),
                        "srgbClr" => get_attr(e, "val"),
                        _ => get_attr(e, "val").and_then(|val| scheme_colour(&val, theme)),
                    };
                    let color_theme = if name == "schemeClr" {
                        get_attr(e, "val").and_then(|value| drawing_theme_index(&value))
                    } else { None };
                    let legacy_color_index = if name == "srgbClr" {
                        get_attr(e, "legacySpreadsheetColorIndex").and_then(|value| value.parse::<u8>().ok())
                    } else { None };
                    colour = named.map(|hex| (hex, Vec::new(), color_theme, legacy_color_index));
                }
                "lumMod" | "lumOff" | "shade" | "tint" => {
                    if let Some(value) = get_attr(e, "val").and_then(|v| v.parse::<f32>().ok()) {
                        if let Some((_, mods, _, _)) = colour.as_mut() {
                            mods.push((name.clone(), value / 100_000.0));
                        }
                        if in_run_props && paints == Some(Paints::Text) {
                            run_color_brightness = match name.as_str() {
                                "tint" => Some(value / 100_000.0),
                                "shade" => Some(-value / 100_000.0),
                                "lumMod" => Some(value / 100_000.0 - 1.0),
                                "lumOff" => Some(value / 100_000.0),
                                _ => run_color_brightness,
                            };
                        }
                    }
                }
                // What the shape says, and how it is laid in its box.
                "bodyPr" => {
                    let emu = |name: &str, fallback: i64| {
                        get_attr(e, name).and_then(|v| v.parse().ok()).unwrap_or(fallback)
                    };
                    said.insets = (
                        emu("lIns", 91440),
                        emu("tIns", 45720),
                        emu("rIns", 91440),
                        emu("bIns", 45720),
                    );
                    said.anchor = get_attr(e, "anchor");
                    said.wrap = get_attr(e, "wrap").as_deref() != Some("none");
                    said.clip = get_attr(e, "vertOverflow").as_deref() == Some("clip");
                }
                "p" => {
                    run_bold = false;
                    run_italic = None;
                    run_underline = false;
                    run_color = None;
                    run_color_theme = None;
                    run_color_brightness = None;
                    run_legacy_color_index = None;
                    run_face = None;
                    run_size = None;
                    paragraph = Some(ShapeParagraph {
                        text: String::new(),
                        runs: Vec::new(),
                        align: None,
                        size: 18.0,
                        bold: false,
                        italic: false,
                        face: None,
                        charset: None,
                        pitch_family: None,
                        color: None,
                        line_pitch: None,
                        line_scale: None,
                    });
                }
                "pPr" if paragraph.is_some() => {
                    if let (Some(held), Some(align)) = (paragraph.as_mut(), get_attr(e, "algn")) {
                        held.align = Some(align);
                    }
                }
                // A break inside a paragraph starts a line without starting a
                // paragraph, which is how `sanko_tool` sets a heading over the
                // text that follows it.
                "br" => {
                    if let Some(held) = paragraph.as_mut() {
                        held.text.push('\n');
                        match held.runs.last_mut() {
                            Some(last) => last.text.push('\n'),
                            None => held.runs.push(ShapeRun {
                                text: "\n".to_string(),
                                bold: run_bold,
                                italic: run_italic,
                                underline: run_underline,
                                face: run_face.clone(),
                                size: run_size,
                                color: run_color.clone(),
                                color_theme: run_color_theme,
                                color_brightness: run_color_brightness,
                                legacy_color_index: run_legacy_color_index,
                            }),
                        }
                    }
                }
                "lnSpc" => in_line_spacing = true,
                // A paragraph can pin its line pitch outright, or ask for a
                // share of the font's own. Both are stated the same way
                // inside `<a:lnSpc>` as they are inside the space before and
                // after a paragraph, so only what is inside that counts.
                "spcPts" if in_line_spacing => {
                    if let (Some(held), Some(points)) = (
                        paragraph.as_mut(),
                        get_attr(e, "val").and_then(|v| v.parse::<f32>().ok()),
                    ) {
                        if held.line_pitch.is_none() && points > 0.0 {
                            held.line_pitch = Some(points / 100.0);
                        }
                    }
                }
                "spcPct" if in_line_spacing => {
                    if let (Some(held), Some(share)) = (
                        paragraph.as_mut(),
                        get_attr(e, "val").and_then(|v| v.parse::<f32>().ok()),
                    ) {
                        if held.line_scale.is_none() && share > 0.0 {
                            held.line_scale = Some(share / 100_000.0);
                        }
                    }
                }
                // Only the first run of a paragraph dresses it: the corpus
                // dresses every run of a shape's paragraph alike.
                // An empty paragraph states its own size in `endParaRPr`
                // rather than `rPr` 遯ｶ繝ｻit has no run to dress 遯ｶ繝ｻand Excel gives
                // it a line of that size. Read as one and the same thing:
                // measured on `_xlsx_shape_block.py`, where an empty paragraph
                // between two blocks is worth exactly one line of the size it
                // states, not of the size a paragraph defaults to.
                "rPr" | "endParaRPr" => {
                    in_run_props = true;
                    run_bold = is_true(get_attr(e, "b").as_deref());
                    run_italic = get_attr(e, "i").map(|value| is_true(Some(&value)));
                    // A shape names the underline it wants rather than
                    // flagging it, so anything but "none" is one.
                    run_underline = get_attr(e, "u")
                        .is_some_and(|kind| !kind.is_empty() && kind != "none");
                    run_color = None;
                    run_color_theme = None;
                    run_color_brightness = None;
                    run_legacy_color_index = None;
                    run_face = None;
                    run_size = get_attr(e, "sz").and_then(|size| size.parse::<f32>().ok()).map(|size| size / 100.0);
                    if let Some(held) = paragraph.as_mut() {
                        if held.text.is_empty() {
                            if let Some(size) = get_attr(e, "sz").and_then(|v| v.parse::<f32>().ok())
                            {
                                held.size = size / 100.0;
                            }
                            held.bold = is_true(get_attr(e, "b").as_deref());
                            held.italic = is_true(get_attr(e, "i").as_deref());
                        }
                    }
                }
                // The East Asian face wins over the Latin one for the sheets
                // this is measured against.
                "latin" | "ea" if in_run_props => {
                    let face = get_attr(e, "typeface")
                        .filter(|face| !face.is_empty())
                        .map(|face| scheme_face(&face, theme).unwrap_or(face));
                    if face.is_some() && (name == "ea" || run_face.is_none()) {
                        run_face = face.clone();
                    }
                    if let Some(held) = paragraph.as_mut() {
                        if held.text.is_empty() {
                            if face.is_some() && (name == "ea" || held.face.is_none()) {
                                held.face = face;
                                // The charset travels with the face it is
                                // written beside: it is what Excel falls back
                                // on when the face is not installed.
                                held.charset = get_attr(e, "charset")
                                    .and_then(|held| held.parse::<i32>().ok());
                                held.pitch_family = get_attr(e, "pitchFamily")
                                    .and_then(|held| held.parse::<i32>().ok());
                            }
                        }
                    }
                }
                "t" => {
                    in_text = true;
                    text.clear();
                }
                // The picture's own part is named by a relationship.
                "blip" if embed.is_none() => {
                    embed = e
                        .attributes()
                        .flatten()
                        .find(|attr| local_name(attr.key.as_ref()) == "embed")
                        .and_then(|attr| String::from_utf8(attr.value.to_vec()).ok());
                }
                "svgBlip" => {
                    vector_embed = e
                        .attributes()
                        .flatten()
                        .find(|attr| local_name(attr.key.as_ref()) == "embed")
                        .and_then(|attr| String::from_utf8(attr.value.to_vec()).ok());
                }
                _ => {}
            }
            // An element with no children ends where it starts.
            if empty && matches!(name.as_str(), "srgbClr" | "schemeClr" | "sysClr") {
                if let (Some((hex, mods, color_theme, legacy_color_index)), Some(part)) = (colour.take(), paints) {
                    let painted = shaded(&hex, &mods);
                    match part {
                        Paints::Fill => {
                            shape.fill = Some(painted);
                            shape.fill_theme = color_theme;
                            shape.fill_legacy_color_index = legacy_color_index;
                            shape.fill_brightness = drawing_color_brightness(&mods);
                        }
                        Paints::Line => {
                            shape.line = Some(ShapeLine {
                                color: painted,
                                color_theme,
                                legacy_color_index,
                                color_brightness: drawing_color_brightness(&mods),
                                color_shade: drawing_color_shade(&mods),
                                width: line_width,
                                dash: dash.clone(),
                                cap: line_cap.clone(),
                                head_end: head_end.clone(),
                                tail_end: tail_end.clone(),
                                head_end_length: head_end_length.clone(),
                                head_end_width: head_end_width.clone(),
                                tail_end_length: tail_end_length.clone(),
                                tail_end_width: tail_end_width.clone(),
                                pattern: None,
                                back_color: None,
                                back_color_theme: None,
                                back_legacy_color_index: None,
                                back_color_brightness: None, back_color_shade: None,
                            })
                        }
                        Paints::LinePatternForeground => {
                            let line = shape.line.get_or_insert_with(|| ShapeLine {
                                color: "000000".into(), color_theme: None, legacy_color_index: None, color_brightness: None, color_shade: None, back_color_brightness: None, back_color_shade: None,
                                width: line_width, dash: dash.clone(), cap: line_cap.clone(),
                                head_end: None, tail_end: None, head_end_length: None, head_end_width: None,
                                tail_end_length: None, tail_end_width: None, pattern: None, back_color: None, back_color_theme: None, back_legacy_color_index: None,
                            });
                            line.color = painted;
                            line.color_theme = color_theme;
                            line.legacy_color_index = legacy_color_index;
                            line.color_brightness = drawing_color_brightness(&mods);
                            line.color_shade = drawing_color_shade(&mods);
                        }
                        Paints::LinePatternBackground => {
                            let line = shape.line.get_or_insert_with(|| ShapeLine {
                                color: "000000".into(), color_theme: None, legacy_color_index: None, color_brightness: None, color_shade: None, back_color_brightness: None, back_color_shade: None,
                                width: line_width, dash: dash.clone(), cap: line_cap.clone(),
                                head_end: None, tail_end: None, head_end_length: None, head_end_width: None,
                                tail_end_length: None, tail_end_width: None, pattern: None, back_color: None, back_color_theme: None, back_legacy_color_index: None,
                            });
                            line.back_color = Some(painted);
                            line.back_color_theme = color_theme;
                            line.back_legacy_color_index = legacy_color_index;
                            line.back_color_brightness = drawing_color_brightness(&mods);
                            line.back_color_shade = drawing_color_shade(&mods);
                        }
                        Paints::Text => {
                            if let Some(held) = paragraph.as_mut() {
                                if held.text.is_empty() {
                                    held.color = Some(painted.clone());
                                }
                            }
                            run_color = Some(painted);
                        }
                    }
                }
            }
            if empty && matches!(name.as_str(), "srgbClr" | "schemeClr" | "sysClr") {
                paints = None;
            }
        }
        match &event {
            Ok(Event::Text(e)) => {
                if let Ok(held) = e.unescape() {
                    if in_text {
                        text.push_str(&held);
                    } else {
                        number.push_str(&held);
                    }
                }
            }
            Ok(Event::End(e)) => {
                let name = local_name(e.name().as_ref());
                let value = || number.trim().parse::<i64>().unwrap_or(0);
                match name.as_str() {
                    "clrFrom" => in_picture_clr_from = false,
                    "clrChange" => {
                        in_picture_clr_change = false;
                        in_picture_clr_from = false;
                    }
                    "col" | "colOff" | "row" | "rowOff" => {
                        if let Some(which) = corner {
                            let corner = if which { &mut from } else { to.get_or_insert(blank) };
                            match name.as_str() {
                                "col" => corner.col = value().max(0) as u32,
                                "colOff" => corner.col_off = value(),
                                "row" => corner.row = value().max(0) as u32,
                                _ => corner.row_off = value(),
                            }
                        }
                    }
                    "x" | "y" if corner.is_some() => {
                        let held = number.trim().parse::<f64>().unwrap_or(0.0);
                        let corner = if corner == Some(true) {
                            frame.get_or_insert((0.0, 0.0))
                        } else {
                            frame_to.get_or_insert((0.0, 0.0))
                        };
                        if name == "x" {
                            corner.0 = held;
                        } else {
                            corner.1 = held;
                        }
                    }
                    "from" | "to" => corner = None,
                    "srgbClr" | "schemeClr" | "sysClr" => {
                        if let (Some((hex, mods, color_theme, legacy_color_index)), Some(part)) = (colour.take(), paints) {
                            let painted = shaded(&hex, &mods);
                            match part {
                                Paints::Fill => {
                                    shape.fill = Some(painted);
                                    shape.fill_theme = color_theme;
                                    shape.fill_legacy_color_index = legacy_color_index;
                                    shape.fill_brightness = drawing_color_brightness(&mods);
                                }
                                Paints::Line => {
                                    shape.line = Some(ShapeLine {
                                        color: painted,
                                        color_theme,
                                        legacy_color_index,
                                        color_brightness: drawing_color_brightness(&mods),
                                        color_shade: drawing_color_shade(&mods),
                                        width: line_width,
                                        dash: dash.clone(),
                                        cap: line_cap.clone(),
                                        head_end: head_end.clone(),
                                        tail_end: tail_end.clone(),
                                        head_end_length: head_end_length.clone(),
                                        head_end_width: head_end_width.clone(),
                                        tail_end_length: tail_end_length.clone(),
                                        tail_end_width: tail_end_width.clone(),
                                        pattern: None,
                                        back_color: None,
                                        back_color_theme: None,
                                        back_legacy_color_index: None,
                                        back_color_brightness: None, back_color_shade: None,
                                    })
                                }
                                Paints::LinePatternForeground => {
                                    let line = shape.line.get_or_insert_with(|| ShapeLine {
                                        color: "000000".into(), color_theme: None, legacy_color_index: None, color_brightness: None, color_shade: None, back_color_brightness: None, back_color_shade: None,
                                        width: line_width, dash: dash.clone(), cap: line_cap.clone(),
                                        head_end: None, tail_end: None, head_end_length: None, head_end_width: None,
                                        tail_end_length: None, tail_end_width: None, pattern: None, back_color: None, back_color_theme: None, back_legacy_color_index: None,
                                    });
                                    line.color = painted;
                                    line.color_theme = color_theme;
                                    line.legacy_color_index = legacy_color_index;
                                    line.color_brightness = drawing_color_brightness(&mods);
                                    line.color_shade = drawing_color_shade(&mods);
                                }
                                Paints::LinePatternBackground => {
                                    let line = shape.line.get_or_insert_with(|| ShapeLine {
                                        color: "000000".into(), color_theme: None, legacy_color_index: None, color_brightness: None, color_shade: None, back_color_brightness: None, back_color_shade: None,
                                        width: line_width, dash: dash.clone(), cap: line_cap.clone(),
                                        head_end: None, tail_end: None, head_end_length: None, head_end_width: None,
                                        tail_end_length: None, tail_end_width: None, pattern: None, back_color: None, back_color_theme: None, back_legacy_color_index: None,
                                    });
                                    line.back_color = Some(painted);
                                    line.back_color_theme = color_theme;
                                    line.back_legacy_color_index = legacy_color_index;
                                    line.back_color_brightness = drawing_color_brightness(&mods);
                                    line.back_color_shade = drawing_color_shade(&mods);
                                }
                                Paints::Text => {
                                    if let Some(held) = paragraph.as_mut() {
                                        if held.text.is_empty() {
                                            held.color = Some(painted.clone());
                                        }
                                    }
                                    run_color = Some(painted);
                                }
                            }
                        }
                        paints = None;
                    }
                    "t" => {
                        in_text = false;
                        if let Some(held) = paragraph.as_mut() {
                            held.text.push_str(&text);
                            held.runs.push(ShapeRun {
                                text: text.clone(),
                                bold: run_bold,
                                italic: run_italic,
                                underline: run_underline,
                                face: run_face.clone(),
                                size: run_size,
                                color: run_color.clone(),
                                color_theme: run_color_theme,
                                color_brightness: run_color_brightness,
                                legacy_color_index: run_legacy_color_index,
                            });
                        }
                        text.clear();
                    }
                    "lnSpc" => in_line_spacing = false,
                    "rPr" | "endParaRPr" => in_run_props = false,
                    "p" => {
                        // A paragraph with nothing in it still spends a line,
                        // wherever it sits: `_xlsx_shape_block.py` anchors the
                        // same four paragraphs to the top, the middle and the
                        // foot of a box, and an empty one at the front, at the
                        // back or in the middle moves the ink by exactly one
                        // line every time. Both ends used to be dropped.
                        if let Some(held) = paragraph.take() {
                            said.paragraphs.push(held);
                        }
                    }
                    "txBody" => {
                        if said.paragraphs.iter().any(|held| !held.text.is_empty()) {
                            shape.text = Some(said.clone());
                        }
                    }
                    "solidFill" | "fillRef" | "lnRef" => paints = None,
                    // A rule states its colour first and how it is broken
                    // after, so the dash can only be put on the line once the
                    // whole of `<a:ln>` has been read.
                    "ln" => {
                        in_ln = false;
                        // The ends are stated AFTER the fill inside `<a:ln>`,
                        // and the fill is what builds the rule 遯ｶ繝ｻso they are
                        // hung on it here rather than when it is made.
                        // `glossary_05` states every one of its twelve arrows
                        // this way, and reading them at build time found none.
                        if let Some(ruled) = shape.line.as_mut() {
                            if ruled.head_end.is_none() {
                                ruled.head_end = head_end.clone();
                            }
                            if ruled.tail_end.is_none() {
                                ruled.tail_end = tail_end.clone();
                            }
                        }
                        if let Some(line) = shape.line.as_mut() {
                            line.width = line_width;
                            line.dash = dash.take();
                        }
                    }
                    "spPr" => in_sp_pr = false,
                    "style" => in_style = false,
                    "extLst" => in_ext_lst = in_ext_lst.saturating_sub(1),
                    "moveTo" | "lnTo" | "cubicBezTo" => {
                        if let (Some(held), Some(which)) = (path.as_mut(), step.take()) {
                            match (which.as_str(), points.as_slice()) {
                                ("moveTo", [(x, y)]) => held.steps.push(PathStep::MoveTo(*x, *y)),
                                ("lnTo", [(x, y)]) => held.steps.push(PathStep::LineTo(*x, *y)),
                                ("cubicBezTo", [(ax, ay), (bx, by), (cx, cy)]) => held
                                    .steps
                                    .push(PathStep::CurveTo(*ax, *ay, *bx, *by, *cx, *cy)),
                                _ => {}
                            }
                        }
                        points.clear();
                    }
                    "prstGeom" => in_preset = false,
                    "custGeom" => {
                        shape.path = path.take();
                        if shape.geometry.is_empty() {
                            shape.geometry = "custGeom".to_string();
                        }
                    }
                    "grpSpPr" => in_group_props = false,
                    "pic" | "sp" | "cxnSp" | "grpSp" | "graphicFrame" => {
                        depth_in_shape = depth_in_shape.saturating_sub(1);
                        // A child of a group is placed by mapping its own box
                        // through the group's transform, and hangs from the
                        // anchor the group hangs from with the difference as
                        // its offset. `002`'s callout is one of these.
                        if let (Some((off, ext, ch_off, ch_ext)), Some(own), Some(size), true) =
                            (group, own_off, own_ext, name != "grpSp")
                        {
                            // A child sits at a FRACTION of the group's box,
                            // and the box is the one the group is DRAWN in 遯ｶ繝ｻ
                            // the rectangle its two cell anchors give, not the
                            // `<a:ext>` it states. The two are the same number
                            // in a file Excel has just written and part company
                            // as soon as the rows move underneath: Excel keeps
                            // `ext` as a cache and does not rewrite it unless
                            // it saves the shape again.
                            //
                            // Measured by `_xlsx_group_stale.py`, which builds
                            // the group with Excel and then rewrites the row
                            // heights in the sheet XML without touching the
                            // drawing. Over six heights from 0.7x to 2x the
                            // children move with the ROWS, to the pixel, where
                            // `ext` would have held them still. `glossary_05`
                            // is such a workbook and its grouped labels were
                            // the only text in it out by more than a pixel.
                            //
                            // The renderer works the box out because only it
                            // knows what the rows and columns come to; a group
                            // hung on one cell has no second anchor to differ
                            // from, so its stated extent stands and the two
                            // readings agree again.
                            let fraction = |part: f64, whole: i64| {
                                if whole != 0 { part / whole as f64 } else { 0.0 }
                            };
                            // Bring the leaf's rectangle back through every
                            // nested group, then express it in the root group's
                            // child-coordinate system. The root's off/ext are
                            // deliberately not used: its two-cell anchor is
                            // the live geometry, as for a single-level group.
                            let map = |position: f64, size: f64, transform: ((i64, i64), (i64, i64), (i64, i64), (i64, i64)), axis: usize| {
                                let (off, ext, ch_off, ch_ext) = transform;
                                let (off, ext, ch_off, ch_ext) = if axis == 0 {
                                    (off.0, ext.0, ch_off.0, ch_ext.0)
                                } else {
                                    (off.1, ext.1, ch_off.1, ch_ext.1)
                                };
                                if ch_ext == 0 {
                                    (position, size)
                                } else {
                                    (
                                        off as f64 + (position - ch_off as f64) * ext as f64 / ch_ext as f64,
                                        size * ext as f64 / ch_ext as f64,
                                    )
                                }
                            };
                            let root = parent_groups.first().copied().unwrap_or((off, ext, ch_off, ch_ext));
                            let mut x = own.0 as f64;
                            let mut y = own.1 as f64;
                            let mut w = size.0 as f64;
                            let mut h = size.1 as f64;
                            if !parent_groups.is_empty() {
                                let transforms = std::iter::once((
                                    (off, ext, ch_off, ch_ext),
                                    current_group_rotation,
                                ))
                                .chain(
                                    parent_groups
                                        .iter()
                                        .copied()
                                        .zip(parent_group_rotations.iter().copied())
                                        .skip(1)
                                        .rev(),
                                );
                                let mut nested_rotation = 0i32;
                                for (transform, rotation) in transforms {
                                    (x, w) = map(x, w, transform, 0);
                                    (y, h) = map(y, h, transform, 1);
                                    if rotation != 0 {
                                        let angle = rotation as f64 / 60_000.0
                                            * std::f64::consts::PI
                                            / 180.0;
                                        let (sin, cos) = angle.sin_cos();
                                        let center_x = transform.0.0 as f64 + transform.1.0 as f64 / 2.0;
                                        let center_y = transform.0.1 as f64 + transform.1.1 as f64 / 2.0;
                                        let dx = x + w / 2.0 - center_x;
                                        let dy = y + h / 2.0 - center_y;
                                        let rotated_x = center_x + dx * cos - dy * sin;
                                        let rotated_y = center_y + dx * sin + dy * cos;
                                        x = rotated_x - w / 2.0;
                                        y = rotated_y - h / 2.0;
                                    }
                                    nested_rotation = nested_rotation.saturating_add(rotation);
                                }
                                shape.rotation = shape.rotation.saturating_add(nested_rotation);
                            }
                            let root_ch_off = root.2;
                            let root_ch_ext = root.3;
                            let mut kind = kind.take().unwrap_or(DrawingKind::Other);
                            if let DrawingKind::Shape(held) = &mut kind {
                                shape.transform_offset = own_off;
                                *held = shape.clone();
                            }
                            let picture = matches!(kind, DrawingKind::Picture { .. });
                            found.push((
                                Drawing {
                                    name: drawing_name.take(),
                                    non_visual_id: drawing_id,
                                    metadata: drawing_metadata.clone(),
                                    ole_prog_id: None,
                                    ole_source_name: None,
                                    ole_clone_position: None,
            ole_embedded_package: None,
            ole_package_is_empty: None,
                                    ole_clone_visible: None,
                                    ole_placement: None,
                                    ole_type: None,
                                    ole_display_as_icon: None,
                                    ole_auto_load: None,
                                    ole_enabled: None,
                                    ole_locked: None,
                                    ole_print_object: None,
                                    ole_shadow: None,
                                    ole_on_action: None,
                                    ole_border_color: None,
                                    ole_border_weight: None,
                                    ole_interior_color: None,
                                    from,
                                    to,
                                    extent,
                                    kind,
                                    frame: Some(Frame {
                                        x: fraction(x - root_ch_off.0 as f64, root_ch_ext.0),
                                        y: fraction(y - root_ch_off.1 as f64, root_ch_ext.1),
                                        w: fraction(w, root_ch_ext.0),
                                        h: fraction(h, root_ch_ext.1),
                                    }),
                                    grouped: true,
                                    group_name: group_name.clone(),
                                    group_non_visual_id,
                                    group_rotation,
                                    group_transform: Some(crate::ir::GroupTransform {
                                        offset: root.0,
                                        extent: root.1,
                                        child_offset: root.2,
                                        child_extent: root.3,
                                    }),
                                    group_child_transform: None,
                                    group_nested: !parent_groups.is_empty(),
                                },
                                DrawingImageRelationships {
                                    raster: if picture { embed.take() } else { None },
                                    vector: if picture { vector_embed.take() } else { None },
                                    shape_id: drawing_id.take(),
                                },
                            ));
                            let _ = (off, ext);
                        }
                        if name == "grpSp" {
                            if let Some(parent) = parent_groups.pop() {
                                group = Some(parent);
                                current_group_rotation = parent_group_rotations.pop().unwrap_or(0);
                            }
                        }
                    }
                    "twoCellAnchor" | "oneCellAnchor" | "absoluteAnchor"
                    | "relSizeAnchor" | "absSizeAnchor" => {
                        // A group's children have been pushed one by one; the
                        // group itself draws nothing.
                        if group.is_none() {
                            let mut kind = kind.take().unwrap_or(DrawingKind::Other);
                            if let DrawingKind::Shape(held) = &mut kind {
                                shape.transform_offset = own_off;
                                *held = shape.clone();
                            }
                            let named = matches!(
                                kind,
                                DrawingKind::Picture { .. } | DrawingKind::Chart(_)
                            );
                            let picture = matches!(kind, DrawingKind::Picture { .. });
                            let held = frame.take().map(|(x, y)| {
                                let (to_x, to_y) = frame_to.take().unwrap_or((x, y));
                                crate::ir::Frame {
                                    x,
                                    y,
                                    w: (to_x - x).max(0.0),
                                    h: (to_y - y).max(0.0),
                                }
                            });
                            found.push((
                                Drawing {
                                    name: drawing_name.take(),
                                    non_visual_id: drawing_id,
                                    metadata: drawing_metadata.clone(),
                                    ole_prog_id: None,
                                    ole_source_name: None,
                                    ole_clone_position: None,
            ole_embedded_package: None,
            ole_package_is_empty: None,
                                    ole_clone_visible: None,
                                    ole_placement: None,
                                    ole_type: None,
                                    ole_display_as_icon: None,
                                    ole_auto_load: None,
                                    ole_enabled: None,
                                    ole_locked: None,
                                    ole_print_object: None,
                                    ole_shadow: None,
                                    ole_on_action: None,
                                    ole_border_color: None,
                                    ole_border_weight: None,
                                    ole_interior_color: None,
                                    from,
                                    to: to.take(),
                                    extent: extent.take(),
                                    kind,
                                    frame: held,
                                    grouped: false,
                                    group_name: None,
                                    group_non_visual_id: None,
                                    group_rotation: 0,
                                    group_transform: None,
                                    group_child_transform: None,
                                    group_nested: false,
                                },
                                DrawingImageRelationships {
                                    raster: if named { embed.take() } else { None },
                                    vector: if picture { vector_embed.take() } else { None },
                                    shape_id: drawing_id.take(),
                                },
                            ));
                        }
                        embed = None;
                        vector_embed = None;
                    }
                    _ => {}
                }
            }
            Ok(Event::Eof) | Err(_) => break,
            _ => {}
        }
        buf.clear();
    }
    found
}

/// The notes a sheet keeps pinned open, from the pair of parts that hold
/// them: the text in `xl/comments{n}.xml`, keyed by cell, and the box in the
/// VML beside it, which states it in points from the sheet's corner and says
/// whether Excel shows it.
/// A VML length, in points, taken from the start of `said`.
///
/// VML writes a size in whatever unit suits it 遯ｶ繝ｻ`175.5pt`, `2.5mm`, `8in` 遯ｶ繝ｻ
/// and a bare number is points. Anything after the number that is not a unit
/// it knows is ignored, so the same reader takes `width:175.5pt;` and the
/// `2.3mm` of an inset list.
fn vml_length(said: &str) -> Option<f32> {
    let said = said.trim_start();
    let end = said
        .find(|c: char| !(c.is_ascii_digit() || c == '.' || c == '-'))
        .unwrap_or(said.len());
    let number: f32 = said[..end].parse().ok()?;
    Some(match said[end..].chars().take(2).collect::<String>().as_str() {
        held if held.starts_with("in") => number * 72.0,
        held if held.starts_with("mm") => number * 72.0 / 25.4,
        held if held.starts_with("cm") => number * 72.0 / 2.54,
        held if held.starts_with("px") => number * 72.0 / 96.0,
        _ => number,
    })
}

#[cfg(test)]
fn parse_comments(comments_xml: &str, vml: &str) -> Vec<crate::ir::Comment> {
    parse_comments_with_theme(comments_xml, vml, &Theme::default())
}

fn vml_shape_attribute(shape: &str, name: &str) -> Option<String> {
    let head = shape.split('>').next()?;
    let xml = format!("<v:shape {head}/>");
    let mut reader = Reader::from_str(&xml);
    match reader.read_event().ok()? {
        Event::Start(element) | Event::Empty(element) => get_attr(&element, name),
        _ => None,
    }
}

fn parse_drawing_id_blocks(xml: &str) -> Vec<u32> {
    let mut reader = Reader::from_str(xml);
    loop {
        match reader.read_event() {
            Ok(Event::Start(element) | Event::Empty(element)) if local_name(element.name().as_ref()) == "idmap" => {
                return get_attr(&element, "data").unwrap_or_default().split(',')
                    .filter_map(|value| value.trim().parse().ok()).filter(|block| *block != 0).collect();
            }
            Ok(Event::Eof) | Err(_) => return Vec::new(),
            _ => {}
        }
    }
}

fn parse_comments_with_theme(comments_xml: &str, vml: &str, theme: &Theme) -> Vec<crate::ir::Comment> {
    use crate::ir::{Anchor, Comment, ShapeParagraph, ShapeRun, ShapeText};

    // The text of each note, by the cell it belongs to.
    let mut said: HashMap<(u32, u32), Vec<ShapeParagraph>> = HashMap::new();
    let mut utf16_by_cell: HashMap<(u32, u32), Vec<u16>> = HashMap::new();
    let mut author_names = Vec::new();
    let mut authors_by_cell: HashMap<(u32, u32), String> = HashMap::new();
    let mut reader = Reader::from_str(comments_xml);
    let mut buf = Vec::new();
    let mut at: Option<(u32, u32)> = None;
    let mut run = ShapeParagraph {
        text: String::new(),
        runs: Vec::new(),
        align: None,
        size: 9.0,
        bold: false,
        italic: false,
        face: None,
        charset: None,
        pitch_family: None,
        color: None,
        line_pitch: None,
        line_scale: None,
    };
    let blank_run = run.clone();
    let mut part = blank_run.clone();
    let mut pieces = Vec::<ShapeRun>::new();
    let mut in_rich_run = false;
    let mut underline = false;
    let (mut in_run_props, mut in_text, mut first) = (false, false, true);
    let mut in_author = false;
    let mut author_text = String::new();
    let mut comment_author_id = None;
    let mut text = String::new();
    let mut text_units = Vec::<u16>::new();
    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(ref e)) | Ok(Event::Empty(ref e)) => {
                match local_name(e.name().as_ref()).as_str() {
                    "comment" => {
                        let cell = get_attr(e, "ref").map(|held| parse_cell_ref(&held));
                        at = cell.map(|(column, row)| (row, column));
                        comment_author_id = get_attr(e, "authorId").and_then(|id| id.parse::<usize>().ok());
                        run = blank_run.clone();
                        part = blank_run.clone();
                        pieces.clear();
                        text_units.clear();
                        first = true;
                    }
                    "author" => {
                        in_author = true;
                        author_text.clear();
                    }
                    "r" => { in_rich_run = true; part = blank_run.clone(); underline = false; }
                    "rPr" => in_run_props = true,
                    "sz" if in_run_props => {
                        if let Some(points) = get_attr(e, "val").and_then(|v| v.parse().ok()) {
                            part.size = points;
                        }
                    }
                    "b" if in_run_props => part.bold = !matches!(get_attr(e, "val").as_deref(), Some("0" | "false")),
                    "i" if in_run_props => part.italic = !matches!(get_attr(e, "val").as_deref(), Some("0" | "false")),
                    "u" if in_run_props => underline = !matches!(get_attr(e, "val").as_deref(), Some("0" | "false" | "none")),
                    "rFont" if in_run_props => part.face = get_attr(e, "val"),
                    "charset" if in_run_props => part.charset = get_attr(e, "val").and_then(|value| value.parse().ok()),
                    "color" if in_run_props => part.color = parse_color_attr(e, theme),
                    "t" => {
                        in_text = true;
                        text.clear();
                    }
                    _ => {}
                }
            }
            Ok(Event::Text(ref e)) if in_author => {
                if let Ok(held) = e.unescape() {
                    author_text.push_str(&held);
                }
            }
            Ok(Event::Text(ref e)) if in_text => {
                if let Ok(held) = e.unescape() {
                    text.push_str(&held);
                }
            }
            Ok(Event::End(ref e)) => match local_name(e.name().as_ref()).as_str() {
                "author" => {
                    author_names.push(std::mem::take(&mut author_text));
                    in_author = false;
                }
                "rPr" => in_run_props = false,
                "t" => {
                    in_text = false;
                    if first && in_rich_run {
                        run.size = part.size; run.bold = part.bold; run.italic = part.italic;
                        run.face = part.face.clone(); run.charset = part.charset; run.color = part.color.clone();
                    }
                    let units = crate::string_value::decode(&text);
                    if in_rich_run {
                        pieces.push(ShapeRun { text: String::from_utf16_lossy(&units),
                            bold: part.bold, italic: Some(part.italic), underline, face: part.face.clone(),
                            size: Some(part.size), color: part.color.clone(), color_theme: None, color_brightness: None, legacy_color_index: None });
                    }
                    text_units.extend(units);
                    first = false;
                }
                "r" => in_rich_run = false,
                "comment" => {
                    if let Some(key) = at.take() {
                        let units = std::mem::take(&mut text_units);
                        run.text = match String::from_utf16(&units) {
                            Ok(text) => text,
                            Err(_) => {
                                utf16_by_cell.insert(key, units.clone());
                                String::from_utf16_lossy(&units)
                            }
                        };
                        // A note is one run of text with newlines in it; the
                        // paragraphs are what the drawing splits it into.
                        let mut paragraphs: Vec<ShapeParagraph> = run
                            .text
                            .split('\n')
                            .map(|line| ShapeParagraph {
                                text: line.trim_end_matches('\r').to_string(),
                                ..run.clone()
                            })
                            .collect();
                        if pieces.len() > 1 || pieces.first().is_some_and(|piece| piece.underline) {
                            let mut paragraph = 0usize;
                            for piece in &pieces {
                                let mut segments = piece.text.split('\n').peekable();
                                while let Some(segment) = segments.next() {
                                    if let Some(target) = paragraphs.get_mut(paragraph) {
                                        if !segment.is_empty() { let mut current = piece.clone(); current.text = segment.to_string(); target.runs.push(current); }
                                    }
                                    if segments.peek().is_some() { paragraph += 1; }
                                }
                            }
                        }
                        said.insert(key, paragraphs);
                        if let Some(author) = comment_author_id
                            .and_then(|index| author_names.get(index))
                            .cloned()
                        {
                            authors_by_cell.insert(key, author);
                        }
                    }
                    comment_author_id = None;
                }
                _ => {}
            },
            Ok(Event::Eof) | Err(_) => break,
            _ => {}
        }
        buf.clear();
    }

    // The box of each note, and whether it is shown at all.
    let mut held = Vec::new();
    // The space matters: `<v:shapetype>` starts with `<v:shape` too, so
    // splitting without it makes the first piece straddle the shapetype AND
    // the first real note 遯ｶ繝ｻand every size after that can pair with the wrong
    // anchor.
    for shape in vml.split("<v:shape ").skip(1) {
        let shape = shape.split("</v:shape>").next().unwrap_or(shape);
        // A note the workbook pins open is drawn as its box. One it leaves
        // closed is still drawn 遯ｶ繝ｻas the red corner Excel puts on any cell
        // that carries a note 遯ｶ繝ｻso it is carried too, marked for what it is
        // rather than dropped.
        let visible = shape.contains("<x:Visible/>");
        let tagged = |name: &str| -> Option<String> {
            let open = format!("<x:{name}>");
            let at = shape.find(&open)? + open.len();
            let rest = &shape[at..];
            Some(rest[..rest.find('<')?].trim().to_string())
        };
        let (Some(row), Some(column)) = (
            tagged("Row").and_then(|v| v.parse::<u32>().ok()),
            tagged("Column").and_then(|v| v.parse::<u32>().ok()),
        ) else {
            continue;
        };
        // `<x:Row>` and `<x:Column>` count from zero, the way the cell
        // reference in the comment part does once it is parsed.
        let Some(paragraphs) = said.get(&(row, column)) else {
            continue;
        };
        // The anchor is eight numbers: the cell each corner hangs from and
        // how far into it, in pixels at 96 dpi. Measured against Excel's own
        // picture of `002`, whose note lands at column 64 plus 12 pixels 遯ｶ繝ｻ
        // where the margin the same shape states is 13 pixels short.
        let Some(anchor) = tagged("Anchor") else { continue };
        let numbers: Vec<i64> = anchor
            .split(',')
            .filter_map(|part| part.trim().parse().ok())
            .collect();
        let [left, dx, top, dy, right, dx2, bottom, dy2] = numbers[..] else {
            continue;
        };
        let corner = |column: i64, x: i64, row: i64, y: i64| Anchor {
            col: column.max(0) as u32,
            col_off: x * 9525,
            row: row.max(0) as u32,
            row_off: y * 9525,
        };
        // The far corner is the box. A note's `<x:Anchor>` names both, and
        // the offsets are pixels within the cell each names.
        let far = corner(right, dx2, bottom, dy2);
        // How big the box is comes from the style, in whatever unit it names:
        // Excel sizes a note to its text and writes the answer there.
        let measure = |name: &str| -> Option<f32> {
            let at = shape.find(name)? + name.len();
            vml_length(&shape[at..])
        };
        let (Some(wide), Some(tall)) = (measure("width:"), measure("height:")) else {
            continue;
        };
        // A note is `#ffffe1` unless the shape says otherwise; the three-digit
        // form is the one Excel writes.
        let fill = shape
            .split("fillcolor=\"")
            .nth(1)
            .and_then(|rest| rest.split('"').next())
            .map(|held| match held {
                "infoBackground [80]" => "FFFFE1".to_string(),
                value => value.strip_prefix('#').unwrap_or(value).to_string(),
            })
            .filter(|held| held.len() == 3 || held.len() == 6)
            .map(|held| {
                if held.len() == 3 {
                    held.chars().flat_map(|part| [part, part]).collect()
                } else {
                    held
                }
            })
            .unwrap_or_else(|| "FFFFE1".to_string());
        let fill_visible = shape
            .split("filled=\"")
            .nth(1)
            .and_then(|rest| rest.split('\"').next())
            .map(|value| !matches!(value.trim().to_ascii_lowercase().as_str(), "f" | "false" | "0"))
            .unwrap_or(true);
        let vml_attribute = |name: &str| -> Option<String> {
            let marker = format!("{name}=\"");
            let start = shape.find(&marker)? + marker.len();
            Some(shape[start..].split('"').next()?.to_string())
        };
        let line_color = vml_attribute("strokecolor").and_then(|value| {
            let value = value.trim();
            if value.starts_with('#') && value.len() == 7 {
                Some(value[1..].to_uppercase())
            } else {
                let name = value.split_whitespace().next()?.to_ascii_lowercase();
                Some(match name.as_str() {
                    "black" => "000000",
                    "white" => "FFFFFF",
                    "red" => "FF0000",
                    "green" => "008000",
                    "lime" => "00FF00",
                    "blue" => "0000FF",
                    "yellow" => "FFFF00",
                    "cyan" => "00FFFF",
                    "magenta" => "FF00FF",
                    "gray" | "grey" => "808080",
                    "maroon" => "800000",
                    "navy" => "000080",
                    "olive" => "808000",
                    "purple" => "800080",
                    "teal" => "008080",
                    _ => return None,
                }.to_string())
            }
        });
        let line_visible = vml_attribute("stroked")
            .map(|value| !matches!(value.trim().to_ascii_lowercase().as_str(), "f" | "false" | "0"))
            .unwrap_or(true);
        let line_weight = vml_attribute("strokeweight")
            .and_then(|value| vml_length(&value))
            .unwrap_or(0.75);
        let fill_transparency = vml_opacity_transparency(shape, "fill").unwrap_or(0.0);
        let line_transparency = vml_opacity_transparency(shape, "stroke").unwrap_or(0.0);
        // What a note keeps between its box and its text. One that states
        // nothing keeps 7.2 points either side and 3.6 above and below, which
        // is what Excel makes a note with 遯ｶ繝ｻasked of COM of a fresh one.
        //
        // Excel honours the vertical part exactly. `_xlsx_note_002.py` writes
        // the corpus floor's own note with top insets of 0, 3.6, 7.2 and 20
        // points and Excel puts its first line 5, 10, 15 and 32 pixels below
        // the box's rule: one pixel of text for one pixel of inset. Only one
        // workbook of the corpus states an inset at all (50 notes, all
        // `2.5mm,2.3mm`); the other 280 notes state none, and were being given
        // that workbook's numbers.
        let inset = shape
            .split("inset=\"")
            .nth(1)
            .and_then(|rest| rest.split('"').next())
            .map(|held| {
                let mut parts = held.split(',');
                let mut next = |fallback: f32| -> f32 {
                    match parts.next().map(str::trim) {
                        Some(one) if !one.is_empty() => {
                            vml_length(one).unwrap_or(fallback)
                        }
                        _ => fallback,
                    }
                };
                (next(7.2), next(3.6), next(7.2), next(3.6))
            })
            .unwrap_or((7.2, 3.6, 7.2, 3.6));
        let emu = |points: f32| (points * 12700.0).round() as i64;
        held.push(Comment {
            position: measure("margin-left:").zip(measure("margin-top:"))
                .map(|(left, top)| (f64::from(left), f64::from(top))),
            shape_id: vml_shape_attribute(shape, "spid").or_else(|| vml_shape_attribute(shape, "id"))
                .and_then(|id| id.strip_prefix("_x0000_s").unwrap_or(&id).parse().ok()),
            shape_name: vml_shape_attribute(shape, "id").filter(|_| vml_shape_attribute(shape, "spid").is_some())
                .map(|name| String::from_utf16_lossy(&crate::string_value::decode(&name))),
            author: authors_by_cell.get(&(row, column)).cloned().unwrap_or_default(),
            utf16_text: utf16_by_cell.get(&(row, column)).cloned(),
            from: corner(left, dx, top, dy),
            size: (wide, tall),
            cell: (row, column),
            to: Some(far),
            text: ShapeText {
                paragraphs: paragraphs.clone(),
                anchor: Some("t".to_string()),
                insets: (emu(inset.0), emu(inset.1), emu(inset.2), emu(inset.3)),
                wrap: true,
                // A note keeps its text inside its box, and the file never
                // says so 遯ｶ繝ｻa note has no body properties at all, unlike a
                // shape, which states `vertOverflow`. Measured instead:
                // `_xlsx_note_clip.py` holds a box and grows the text, and
                // Excel draws 3 lines in a 53-pixel box, 4 in an 80 and 6 in
                // a 107, whether the text runs to four lines or eight. It is
                // the same rule a clipped shape follows.
                //
                // The corpus's lowest-scoring workbook is where this shows:
                // `002`'s worst tiles are all one column of yellow notes, in
                // which Excel stops after three lines and this drew a fourth.
                clip: true,
            },
            fill: Some(fill.to_uppercase()),
            fill_visible,
            fill_transparency,
            line_color,
            line_visible,
            line_weight,
            line_transparency,
            visible,
        });
    }
    held
}

fn vml_opacity_transparency(shape: &str, element: &str) -> Option<f32> {
    let marker = format!("<v:{element} ");
    let attributes = shape.split(&marker).nth(1)?.split('>').next()?;
    let encoded = attributes
        .split("opacity=\"")
        .nth(1)?
        .split('\"')
        .next()?;
    let opacity = if let Some(fixed) = encoded.strip_suffix('f') {
        fixed.parse::<f32>().ok()? / 65_536.0
    } else if let Some(percent) = encoded.strip_suffix('%') {
        percent.parse::<f32>().ok()? / 100.0
    } else {
        encoded.parse::<f32>().ok()?
    };
    Some((1.0 - opacity).clamp(0.0, 1.0))
}

fn parse_threaded_comments(xml: &str) -> Vec<crate::ir::ThreadedComment> {
    use crate::ir::ThreadedComment;

    let mut reader = Reader::from_str(xml);
    let mut buf = Vec::new();
    let mut stack: Vec<(usize, bool)> = Vec::new();
    let mut comments = Vec::<Option<ThreadedComment>>::new();
    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(ref event)) => {
                match local_name(event.name().as_ref()).as_str() {
                    "threadedComment" => {
                        let Some((column, row)) = get_attr(event, "ref").map(|cell| parse_cell_ref(&cell)) else {
                            buf.clear();
                            continue;
                        };
                        let id = get_attr(event, "id").unwrap_or_default();
                        let parent_id = get_attr(event, "parentId").or_else(|| {
                            stack.last().and_then(|(parent, _)| {
                                comments[*parent].as_ref().map(|parent| parent.id.clone())
                            })
                        });
                        let index = comments.len();
                        comments.push(None);
                        comments[index] = Some(ThreadedComment {
                            cell: (row, column),
                            id,
                            parent_id,
                            person_id: get_attr(event, "personId").unwrap_or_default(),
                            person_name: String::new(),
                            person_provider_id: String::new(),
                            person_user_id: String::new(),
                            date: get_attr(event, "dT").unwrap_or_default(),
                            text: String::new(),
                            utf16_text: None,
                            resolved: matches!(get_attr(event, "done").as_deref(), Some("1" | "true")),
                        });
                        stack.push((index, false));
                    }
                    "text" => {
                        if let Some((_, in_text)) = stack.last_mut() {
                            *in_text = true;
                        }
                    }
                    _ => {}
                }
            }
            Ok(Event::Empty(ref event))
                if local_name(event.name().as_ref()) == "threadedComment" =>
            {
                if let Some((column, row)) = get_attr(event, "ref").map(|cell| parse_cell_ref(&cell)) {
                    let parent_id = get_attr(event, "parentId").or_else(|| {
                        stack.last().and_then(|(parent, _)| {
                            comments[*parent].as_ref().map(|parent| parent.id.clone())
                        })
                    });
                    comments.push(Some(ThreadedComment {
                        cell: (row, column),
                        id: get_attr(event, "id").unwrap_or_default(),
                        parent_id,
                        person_id: get_attr(event, "personId").unwrap_or_default(),
                        person_name: String::new(),
                        person_provider_id: String::new(),
                        person_user_id: String::new(),
                        date: get_attr(event, "dT").unwrap_or_default(),
                        text: String::new(),
                            utf16_text: None,
                        resolved: matches!(get_attr(event, "done").as_deref(), Some("1" | "true")),
                    }));
                }
            }
            Ok(Event::Text(event)) if stack.last().is_some_and(|(_, in_text)| *in_text) => {
                if let (Some((index, _)), Ok(text)) = (stack.last(), event.unescape()) {
                    if let Some(comment) = comments[*index].as_mut() {
                        comment.text.push_str(&text);
                    }
                }
            }
            Ok(Event::End(ref event)) => match local_name(event.name().as_ref()).as_str() {
                "text" => {
                    if let Some((_, in_text)) = stack.last_mut() {
                        *in_text = false;
                    }
                }
                "threadedComment" => {
                    stack.pop();
                }
                _ => {}
            },
            Ok(Event::Eof) | Err(_) => break,
            _ => {}
        }
        buf.clear();
    }
    comments.into_iter().flatten().map(|mut comment| {
        let units = crate::string_value::decode(&comment.text);
        comment.text = String::from_utf16_lossy(&units);
        if String::from_utf16(&units).is_err() { comment.utf16_text = Some(units); }
        comment
    }).collect()
}

#[derive(Default)]
struct ThreadedCommentPerson {
    display_name: String,
    provider_id: String,
    user_id: String,
}

fn parse_person_list(xml: &str) -> HashMap<String, ThreadedCommentPerson> {
    let mut reader = Reader::from_str(xml);
    let mut people = HashMap::new();
    loop {
        match reader.read_event() {
            Ok(Event::Start(person) | Event::Empty(person))
                if local_name(person.name().as_ref()) == "person" =>
            {
                if let Some(id) = get_attr(&person, "id") {
                    people.insert(id, ThreadedCommentPerson {
                        display_name: get_attr(&person, "displayName").unwrap_or_default(),
                        provider_id: get_attr(&person, "providerId").unwrap_or_default(),
                        user_id: get_attr(&person, "userId").unwrap_or_default(),
                    });
                }
            }
            Ok(Event::Eof) | Err(_) => break,
            _ => {}
        }
    }
    people
}

/// The theme slot a DrawingML colour names. `tx1` and `bg1` are the same two
/// colours as `dk1` and `lt1` under other names.
pub(crate) fn scheme_colour(name: &str, theme: &Theme) -> Option<String> {
    let slot = match name {
        "dk1" | "tx1" => 0,
        "lt1" | "bg1" => 1,
        "dk2" | "tx2" => 2,
        "lt2" | "bg2" => 3,
        "accent1" => 4,
        "accent2" => 5,
        "accent3" => 6,
        "accent4" => 7,
        "accent5" => 8,
        "accent6" => 9,
        "hlink" => 10,
        "folHlink" => 11,
        _ => return None,
    };
    theme.colours.get(slot).cloned()
}

/// The VBA `ObjectThemeColor` value corresponding to a DrawingML colour name.
fn drawing_theme_index(name: &str) -> Option<usize> {
    match name {
        "dk1" | "tx1" => Some(1),
        "lt1" | "bg1" => Some(2),
        "dk2" | "tx2" => Some(3),
        "lt2" | "bg2" => Some(4),
        "accent1" => Some(5),
        "accent2" => Some(6),
        "accent3" => Some(7),
        "accent4" => Some(8),
        "accent5" => Some(9),
        "accent6" => Some(10),
        "hlink" => Some(11),
        "folHlink" => Some(12),
        _ => None,
    }
}

/// A colour with the modifiers DrawingML hangs off it: `shade` and `tint`
/// move every channel toward black or white, `lumMod` and `lumOff` scale and
/// raise how light it is.
pub(crate) fn shaded(hex: &str, mods: &[(String, f32)]) -> String {
    let Ok(value) = u32::from_str_radix(hex, 16) else {
        return hex.to_string();
    };
    let mut channels = [
        ((value >> 16) & 0xFF) as f32 / 255.0,
        ((value >> 8) & 0xFF) as f32 / 255.0,
        (value & 0xFF) as f32 / 255.0,
    ];
    // `shade` and `tint` scale the LIGHT, not the byte. sRGB is not linear,
    // so halving a byte and halving the light are different colours 遯ｶ繝ｻand it
    // is the light Excel halves. `cas-r02-ippan-4hyou` rules a box in `lt1`
    // under `shade 50000` and Excel draws 188, where halving the byte gives
    // 128. `_xlsx_shade_curve.py` sweeps both modifiers over four bases and
    // three amounts: all twelve land on the light, several to the integer 遯ｶ繝ｻ
    // white shaded a quarter is 137 and not 64, `4472C4` shaded a quarter is
    // 32,58,104 and not 17,28,49.
    let straight = |part: f32| {
        if part <= 0.040_45 {
            part / 12.92
        } else {
            ((part + 0.055) / 1.055).powf(2.4)
        }
    };
    let encoded = |light: f32| {
        if light <= 0.003_130_8 {
            12.92 * light
        } else {
            1.055 * light.powf(1.0 / 2.4) - 0.055
        }
    };
    for (name, amount) in mods {
        match name.as_str() {
            "shade" => channels
                .iter_mut()
                .for_each(|part| *part = encoded(straight(*part) * amount)),
            "tint" => channels.iter_mut().for_each(|part| {
                *part = encoded(straight(*part) * amount + (1.0 - amount))
            }),
            // Lightness is HSL's, and only it moves: the hue and the
            // saturation stay where they are. Measured on `002`'s banner,
            // whose fill is accent1 under `lumMod 20% lumOff 80%` 遯ｶ繝ｻ5B9BD5
            // becomes DEEBF7 exactly, where shifting the bytes by hand lands
            // at CCEDFF.
            "lumMod" | "lumOff" => {
                let high = channels[0].max(channels[1]).max(channels[2]);
                let low = channels[0].min(channels[1]).min(channels[2]);
                let lum = (high + low) / 2.0;
                let span = high - low;
                let wanted = if name == "lumMod" {
                    lum * amount
                } else {
                    (lum + amount).min(1.0)
                };
                if span <= f32::EPSILON {
                    channels = [wanted; 3];
                } else {
                    let sat = span / (1.0 - (2.0 * lum - 1.0).abs()).max(f32::EPSILON);
                    let reach = (1.0 - (2.0 * wanted - 1.0).abs()) * sat.min(1.0);
                    let scale = reach / span;
                    for part in channels.iter_mut() {
                        *part = (wanted + (*part - lum) * scale).clamp(0.0, 1.0);
                    }
                }
            }
            _ => {}
        }
    }
    format!(
        "{:02X}{:02X}{:02X}",
        (channels[0] * 255.0).round() as u8,
        (channels[1] * 255.0).round() as u8,
        (channels[2] * 255.0).round() as u8
    )
}

fn drawing_color_brightness(mods: &[(String, f32)]) -> Option<f32> {
    let mut brightness = None;
    for (name, value) in mods {
        brightness = match name.as_str() {
            "tint" => Some(*value),
            "lumMod" => Some(*value - 1.0),
            "lumOff" => Some(*value),
            _ => brightness,
        };
    }
    brightness.filter(|value| value.abs() > f32::EPSILON)
}

fn drawing_color_shade(mods: &[(String, f32)]) -> Option<f32> {
    mods.iter().find_map(|(name, value)| (name == "shade").then_some(*value))
}

fn table_filter_xml_with_namespaces(xml: &str, namespaces: &[(Vec<u8>, Vec<u8>)]) -> Option<String> {
    let mut reader = Reader::from_str(xml);
    loop {
        let before = oxidocs_common::xml_utils::xml_byte_position(xml, reader.buffer_position());
        let event = reader.read_event().ok()?;
        let after = oxidocs_common::xml_utils::xml_byte_position(xml, reader.buffer_position());
        if let Event::Start(element) | Event::Empty(element) = &event {
            let mut root = element.to_owned();
            for (key, value) in namespaces {
                if !element.attributes().flatten().any(|attribute| attribute.key.as_ref() == key.as_slice()) {
                    root.push_attribute((key.as_slice(), value.as_slice()));
                }
            }
            let mut writer = quick_xml::Writer::new(Vec::new());
            writer.write_event(if matches!(event, Event::Empty(_)) { Event::Empty(root) } else { Event::Start(root) }).ok()?;
            let root = String::from_utf8(writer.into_inner()).ok()?;
            return Some(format!("{}{root}{}", &xml[..before], &xml[after..]));
        }
        if matches!(event, Event::Eof) { return None; }
    }
}

fn parse_table_xml(
    xml: &str,
    theme: &Theme,
    dxf_borders: &[BorderInfo],
) -> Option<crate::ir::Table> {
    let mut reader = Reader::from_str(xml);
    let mut range = None;
    let mut origin_id = None;
    let mut header_rows = 1;
    let mut totals_rows = 0;
    let mut totals = Vec::<crate::ir::TableTotals>::new();
    let mut totals_column = None;
    let mut totals_formula = None::<String>;
    let mut insert_row = false;
    let mut style = None;
    let mut banded_rows = false;
    let mut banded_columns = false;
    let mut emphasize_first_column = false;
    let mut emphasize_last_column = false;
    let mut show_auto_filter = false;
    let mut auto_filter_xml = None;
    let mut filter_start = None;
    let mut namespaces = Vec::new();
    let mut hidden_filter_buttons = Vec::new();
    let mut outline_at: Option<usize> = None;
    let mut name = String::new();
    let mut columns: Vec<String> = Vec::new();
    let mut buf = Vec::new();
    loop {
        let before = oxidocs_common::xml_utils::xml_byte_position(xml, reader.buffer_position());
        let event = reader.read_event_into(&mut buf);
        match event {
            Ok(Event::Start(ref e)) | Ok(Event::Empty(ref e)) => {
                match local_name(e.name().as_ref()).as_str() {
                    "table" => {
                        origin_id = get_attr(e,"id").and_then(|value|value.parse().ok());
                        namespaces = e.attributes().flatten().filter(|attribute| attribute.key.as_ref() == b"xmlns" || attribute.key.as_ref().starts_with(b"xmlns:"))
                            .map(|attribute| (attribute.key.as_ref().to_vec(), attribute.value.to_vec())).collect();
                        insert_row = is_true(get_attr(e, "insertRow").as_deref());
                        // `displayName` is what a formula writes; `name` is the
                        // same thing unless someone has renamed the table, and
                        // then only `displayName` follows.
                        name = get_attr(e, "displayName")
                            .or_else(|| get_attr(e, "name"))
                            .unwrap_or_default();
                        range = get_attr(e, "ref").as_deref().and_then(parse_range_ref);
                        if let Some(count) = get_attr(e, "headerRowCount") {
                            header_rows = count.parse().unwrap_or(1);
                        }
                        if let Some(count) = get_attr(e, "totalsRowCount") {
                            totals_rows = count.parse().unwrap_or(0);
                        }
                        outline_at = get_attr(e, "tableBorderDxfId")
                            .and_then(|held| held.parse::<usize>().ok());
                    }
                    "tableColumn" => {
                        if let Some(heading) = get_attr(e, "name") {
                            columns.push(heading);
                            totals.push(crate::ir::TableTotals {
                                function: get_attr(e, "totalsRowFunction"),
                                label: get_attr(e, "totalsRowLabel").map(|text| crate::string_value::decode(&text)),
                                formula: None,
                            });
                            totals_column = (!matches!(event, Ok(Event::Empty(_)))).then_some(totals.len() - 1);
                        }
                    }
                    "totalsRowFormula" => totals_formula = Some(String::new()),
                    "tableStyleInfo" => {
                        style = get_attr(e, "name");
                        banded_rows = is_true(get_attr(e, "showRowStripes").as_deref());
                        banded_columns = is_true(get_attr(e, "showColumnStripes").as_deref());
                        emphasize_first_column = is_true(get_attr(e, "showFirstColumn").as_deref());
                        emphasize_last_column = is_true(get_attr(e, "showLastColumn").as_deref());
                    }
                    "autoFilter" => {
                        show_auto_filter = true;
                        if matches!(event, Ok(Event::Empty(_))) {
                            auto_filter_xml = table_filter_xml_with_namespaces(&xml[before..oxidocs_common::xml_utils::xml_byte_position(xml, reader.buffer_position())], &namespaces);
                        } else { filter_start = Some(before); }
                    }
                    "filterColumn" if is_true(get_attr(e, "hiddenButton").as_deref()) || matches!(get_attr(e, "showButton").as_deref(), Some("0" | "false")) => {
                        if let Some(column) = get_attr(e, "colId").and_then(|value| value.parse().ok()) { hidden_filter_buttons.push(column); }
                    }
                    _ => {}
                }
            }
            Ok(Event::Text(ref text)) if totals_formula.is_some() => {
                totals_formula.as_mut().unwrap().push_str(&text.unescape().ok()?);
            }
            Ok(Event::CData(ref text)) if totals_formula.is_some() => {
                totals_formula.as_mut().unwrap().push_str(&String::from_utf8_lossy(text));
            }
            Ok(Event::End(ref element)) if local_name(element.name().as_ref()) == "totalsRowFormula" => {
                if let (Some(column), Some(formula)) = (totals_column, totals_formula.take()) {
                    totals[column].formula = Some(crate::string_value::decode(&formula));
                }
            }
            Ok(Event::End(ref element)) if local_name(element.name().as_ref()) == "tableColumn" => totals_column = None,
            Ok(Event::End(ref element)) if local_name(element.name().as_ref()) == "autoFilter" => {
                if let Some(start) = filter_start.take() {
                    auto_filter_xml = table_filter_xml_with_namespaces(&xml[start..oxidocs_common::xml_utils::xml_byte_position(xml, reader.buffer_position())], &namespaces);
                }
            }
            Ok(Event::Eof) | Err(_) => break,
            _ => {}
        }
        buf.clear();
    }
    // The outline is one rule, not four: every one of the corpus's fifteen
    // states the same thin auto-coloured line on all four sides, so the first
    // side that names a style is the rule.
    let outline = outline_at
        .and_then(|at| dxf_borders.get(at))
        .and_then(|held| {
            held.top
                .clone()
                .or_else(|| held.bottom.clone())
                .or_else(|| held.left.clone())
                .or_else(|| held.right.clone())
        });
    let (start_col, start_row, end_col, end_row) = range?;
    // A built-in style is named for the theme colour it uses: TableStyleMedium2
    // dresses the table in accent1, Medium7 in accent6. Measured on a worksheet
    // holding one table per style, header and band both exact.
    let accent = style.as_deref().and_then(|name| {
        let number: u32 = name
            .strip_prefix("TableStyleMedium")
            .or_else(|| name.strip_prefix("TableStyleLight"))
            .or_else(|| name.strip_prefix("TableStyleDark"))?
            .parse()
            .ok()?;
        // 1 is the greyscale one; 2 onward walk accent1..accent6 and repeat.
        if number < 2 {
            return None;
        }
        theme.colour(4 + ((number - 2) % 6) as usize).map(str::to_string)
    });
    let band = accent.as_deref().map(|colour| tinted(colour, 0.8));
    // A Medium table is ruled along every row in the accent under a lighter
    // tint: `doi-list` wears TableStyleMedium7 over accent6 4EA72E, and Excel
    // draws its rules 8ED973 遯ｶ繝ｻwhich is that colour with its lightness moved
    // to `0.6 L + 0.4`, the tint below. The Light and Dark styles rule
    // themselves differently and are left alone until they are measured.
    let rule = accent
        .as_deref()
        .filter(|_| style.as_deref().is_some_and(|name| name.starts_with("TableStyleMedium")))
        .map(|colour| tinted(colour, 0.4));
    Some(crate::ir::Table {
        origin_id,
        origin_part: None,
        name,
        columns,
        start_row,
        start_col,
        end_row,
        end_col,
        style,
        outline,
        header_rows,
        totals_rows,
        totals,
        insert_row,
        banded_rows,
        banded_columns,
        emphasize_first_column,
        emphasize_last_column,
        show_auto_filter,
        auto_filter_xml,
        hidden_filter_buttons,
        accent,
        band,
        rule,
    })
}

fn worksheet_visual_extensions(xml: &str) -> Result<std::collections::BTreeMap<String, String>, XlsxError> {
    let mut reader = Reader::from_str(xml);
    let mut depth = 0usize;
    let mut block = None;
    let mut ids = Vec::new();
    let mut result = std::collections::BTreeMap::new();
    loop {
        let start = oxidocs_common::xml_utils::xml_byte_position(xml, reader.buffer_position());
        match reader.read_event()? {
            Event::Start(ref e) => {
                if e.name().as_ref() == b"x14:conditionalFormatting" {
                    block = Some((depth, start)); ids.clear();
                } else if block.is_some() && e.name().as_ref() == b"x14:cfRule" {
                    if let Some(id) = get_attr(e, "id") { ids.push(id); }
                }
                depth += 1;
            }
            Event::Empty(ref e) if block.is_some() && e.name().as_ref() == b"x14:cfRule" => {
                if let Some(id) = get_attr(e, "id") { ids.push(id); }
            }
            Event::End(ref e) => {
                depth = depth.saturating_sub(1);
                if e.name().as_ref() == b"x14:conditionalFormatting" && block.is_some_and(|(held, _)| held == depth) {
                    let (_, first) = block.take().unwrap();
                    if ids.len() == 1 { result.insert(ids[0].clone(), xml[first..oxidocs_common::xml_utils::xml_byte_position(xml, reader.buffer_position())].to_string()); }
                }
            }
            Event::Eof => break,
            _ => {}
        }
    }
    Ok(result)
}

/// Read an icon-set rule that exists only in an x14 extension. Keep the
/// original block for lossless preservation and expose a base-shaped view to
/// consumers that read threshold properties.
fn extended_only_icon_rule(id: &str, xml: &str, stylesheet: &StyleSheet) -> Result<Option<crate::ir::ConditionalRule>, XlsxError> {
    let mut reader = Reader::from_str(xml);
    let mut rule = None;
    let mut visual = String::new();
    let mut threshold: Option<(String, Option<String>, String)> = None;
    let mut in_formula = false;
    let mut in_range = false;
    let mut range = String::new();
    let escaped = |value: &str| quick_xml::escape::escape(value).into_owned();
    loop {
        match reader.read_event()? {
            Event::Start(ref element) => match element.name().as_ref() {
                b"x14:cfRule" if get_attr(element, "id").as_deref() == Some(id) => {
                    if get_attr(element, "type").as_deref() != Some("iconSet") { return Ok(None); }
                    rule = Some(conditional_from(element, &[], stylesheet));
                }
                b"x14:iconSet" => {
                    visual.push_str("<iconSet");
                    for name in ["iconSet", "showValue", "reverse"] {
                        if let Some(value) = get_attr(element, name) { visual.push_str(&format!(" {name}=\"{}\"", escaped(&value))); }
                    }
                    visual.push('>');
                }
                b"x14:cfvo" => threshold = Some((get_attr(element, "type").unwrap_or_default(), get_attr(element, "gte"), String::new())),
                b"xm:f" if threshold.is_some() => in_formula = true,
                b"xm:sqref" => in_range = true,
                _ => {}
            },
            Event::Empty(ref element) if element.name().as_ref() == b"x14:cfvo" => {
                let kind = get_attr(element, "type").unwrap_or_default();
                let gte = get_attr(element, "gte").map(|value| format!(" gte=\"{}\"", escaped(&value))).unwrap_or_default();
                visual.push_str(&format!("<cfvo type=\"{}\"{gte}/>", escaped(&kind)));
            }
            Event::Text(ref text) if in_formula => if let Some((_, _, formula)) = threshold.as_mut() { formula.push_str(&text.unescape()?); },
            Event::Text(ref text) if in_range => range.push_str(&text.unescape()?),
            Event::End(ref element) => match element.name().as_ref() {
                b"xm:f" => in_formula = false,
                b"xm:sqref" => in_range = false,
                b"x14:cfvo" => if let Some((kind, gte, formula)) = threshold.take() {
                    let gte = gte.map(|value| format!(" gte=\"{}\"", escaped(&value))).unwrap_or_default();
                    let value = if formula.is_empty() { String::new() } else { format!(" val=\"{}\"", escaped(&formula)) };
                    visual.push_str(&format!("<cfvo type=\"{}\"{value}{gte}/>", escaped(&kind)));
                },
                b"x14:iconSet" => visual.push_str("</iconSet>"),
                _ => {}
            },
            Event::Eof => break,
            _ => {}
        }
    }
    let Some(mut rule) = rule else { return Ok(None); };
    rule.ranges = parse_sqref(&range);
    if rule.ranges.is_empty() || !visual.starts_with("<iconSet") || !visual.ends_with("</iconSet>") { return Ok(None); }
    rule.visual_xml = Some(visual);
    rule.visual_extension_id = Some(id.to_string());
    rule.visual_extension_xml = Some(xml.to_string());
    rule.visual_extension_only = true;
    Ok(Some(rule))
}

fn parse_worksheet(
    xml: &str,
    sheet_name: &str,
    shared_strings: &[SharedString],
    stylesheet: &StyleSheet,
    theme: &Theme,
) -> Result<Sheet, XlsxError> {
    let mut reader = Reader::from_str(xml);
    let mut rows: Vec<Row> = Vec::new();
    let mut max_col: u32 = 0;

    // Column widths: index is 0-based col number
    let mut col_widths: Vec<f32> = Vec::new();
    let mut col_custom_widths = std::collections::BTreeMap::new();
    let mut best_fit_cols = std::collections::BTreeSet::new();
    let mut hidden_cols: Vec<u32> = Vec::new();
    let mut auto_filter: Option<crate::ir::AutoFilter> = None;
    let mut declared_range: Option<(u32, u32, u32, u32)> = None;
    let mut unsupported: Vec<String> = Vec::new();
    // Whether the sheet is protected, which decides nothing about how it is
    // drawn and everything about whether it can be typed into.
    let mut protected = false;
    let mut protection = None;
    // What the sheet says its ranges will accept, and the rule being read.
    let mut validations: Vec<crate::ir::DataValidation> = Vec::new();
    let mut rule: Option<crate::ir::DataValidation> = None;
    let mut rule_formula: u8 = 0;
    let mut rule_text = String::new();
    // The rules that change how a range looks, the ranges the block being read
    // covers, and the rule inside it.
    let mut conditional_rules: Vec<crate::ir::ConditionalRule> = Vec::new();
    let mut cf_ranges: Vec<(u32, u32, u32, u32)> = Vec::new();
    let mut cf_rule: Option<crate::ir::ConditionalRule> = None;
    let mut in_cf_formula = false;
    let mut cf_text = String::new();
    let mut in_color_scale = false;
    let mut visual_start: Option<(usize, String)> = None;
    let mut in_visual_extension_id = false;
    // Extended conditional-format payloads contain their own cfRule and
    // conditionalFormatting tags. They supplement the base rule and must not
    // be mistaken for another rule with no priority or ranges.
    let mut conditional_extension_depth = 0usize;
    let mut filter_field: Option<u32> = None;
    let mut filter_criteria: Vec<String> = Vec::new();
    let mut filter_either = false;
    // Zero until the sheet states one. A stated default is measured the same
    // way a <col> width is; Excel's own default, for a sheet that states none,
    // is a plain count of characters instead, so the two cannot share a value.
    let mut default_col_width: f32 = 0.0;
    let mut default_col_width_explicit = false;
    let mut default_row_height: f32 = 15.0;
    let mut frozen_rows: u32 = 0;
    let mut frozen_cols: u32 = 0;
    let mut default_row_custom = false;
    // The fonts <col> styles put on their columns. A row's height is
    // measured from these, not from the number the sheet or the row states.
    let mut col_fonts: Vec<(u32, u32, String, f32)> = Vec::new();
    let mut band_styles = crate::ir::BandStyles::default();
    let mut merge_cells: Vec<MergeCell> = Vec::new();
    // The colour on the sheet's tab, from <sheetPr><tabColor>.
    let mut tab_color: Option<String> = None;

    // State tracking
    let mut current_row_index: u32 = 0;
    let mut current_row_height: Option<f32> = None;
    let mut current_row_custom = false;
    let mut current_row_format_recorded: Option<bool> = None;
    let mut current_row_font: Option<(String, f32)> = None;
    let mut current_row_thick_top = false;
    let mut current_row_thick_bottom = false;
    let mut current_row_hidden = false;
    let mut current_cells: Vec<Cell> = Vec::new();
    let mut in_row = false;

    // Cell state
    let mut cell_col: u32 = 0;
    let mut cell_type: Option<String> = None;
    let mut cell_style_index: Option<usize> = None;
    let mut in_cell = false;
    let mut in_value = false;
    let mut value_text = String::new();
    let mut in_formula = false;
    let mut formula_text = String::new();
    // An ARRAY formula is written on its top-left cell alone, with the block
    // it is dealt across in `ref`; the other cells of the block hold only
    // their values. Each anchor is noted as it is read and the block filled
    // in once every cell has been, since the members come later in the file.
    let mut cell_array_ref: Option<String> = None;
    let mut array_anchors: Vec<(u32, u32, String, String, Option<Vec<u16>>)> = Vec::new();
    // A cell may carry its text itself, in an <is>, instead of pointing into
    // the shared table. Anything that writes a sheet without building that
    // table does it this way. Its text sits in <t> elements, and its phonetic
    // guides in <rPh><t>, which are no more part of the cell's text here than
    // they are in a shared string.
    let mut in_inline = false;
    let mut inline_start = None;
    let mut inline_runs = Vec::new();
    let mut inline_units = None;
    let mut in_inline_text = false;
    let mut in_inline_phonetic = false;

    // Section tracking
    let mut in_merge_cells = false;

    loop {
        let event_start = oxidocs_common::xml_utils::xml_byte_position(xml, reader.buffer_position());
        match reader.read_event()? {
            Event::Start(e) => {
                let name = local_name(e.name().as_ref());
                note_unsupported(&name, &mut unsupported);
                match name.as_str() {
                    "sheetProtection" => {
                        protected = is_true(get_attr(&e, "sheet").as_deref());
                        protection = Some(sheet_protection_from(&e));
                    }
                    "extLst" => conditional_extension_depth += 1,
                    "id" if conditional_extension_depth > 0 && cf_rule.is_some() && e.name().as_ref() == b"x14:id" => in_visual_extension_id = true,
                    "dataValidation" => rule = Some(validation_from(&e)),
                    "conditionalFormatting" if conditional_extension_depth == 0 => {
                        cf_ranges = get_attr(&e, "sqref")
                            .map(|held| parse_sqref(&held))
                            .unwrap_or_default();
                    }
                    "cfRule" if conditional_extension_depth == 0 => cf_rule = Some(conditional_from(&e, &cf_ranges, stylesheet)),
                    "colorScale" if conditional_extension_depth == 0 && cf_rule.as_ref().is_some_and(|r| r.kind == "colorScale") => {
                        in_color_scale = true;
                        visual_start = Some((event_start, name.clone()));
                    }
                    "dataBar" | "iconSet" if conditional_extension_depth == 0 && cf_rule.as_ref().is_some_and(|r| r.kind == name) => {
                        visual_start = Some((event_start, name.clone()));
                    }
                    "cfvo" if in_color_scale && conditional_extension_depth == 0 => {
                        if let Some(held) = cf_rule.as_mut() { held.color_scale_points = Some(held.color_scale_points.unwrap_or(0).saturating_add(1)); }
                    }
                    "formula" if conditional_extension_depth == 0 && cf_rule.is_some() => {
                        in_cf_formula = true;
                        cf_text.clear();
                    }
                    "formula1" if rule.is_some() => {
                        rule_formula = 1;
                        rule_text.clear();
                    }
                    "formula2" if rule.is_some() => {
                        rule_formula = 2;
                        rule_text.clear();
                    }
                    "tabColor" => {
                        tab_color = parse_color_attr(&e, theme);
                    }
                    "dimension" => {
                        declared_range = get_attr(&e, "ref")
                            .as_deref()
                            .and_then(parse_range_ref)
                            .map(|(start_col, start_row, end_col, end_row)| {
                                (start_row, start_col, end_row, end_col)
                            });
                    }
                    "autoFilter" => {
                        if let Some(reference) = get_attr(&e, "ref") {
                            if let Some((start_col, start_row, end_col, end_row)) =
                                parse_range_ref(&reference)
                            {
                                auto_filter = Some(crate::ir::AutoFilter {
                                    start_row,
                                    start_col,
                                    end_row,
                                    end_col,
                                    columns: Vec::new(),
                                });
                            }
                        }
                    }
                    "filterColumn" => {
                        filter_field = get_attr(&e, "colId")
                            .and_then(|value| value.parse::<u32>().ok())
                            .map(|col| col + 1);
                        filter_criteria.clear();
                        filter_either = false;
                    }
                    "filters" => {
                        filter_either = true;
                    }
                    "customFilters" => {
                        filter_either = get_attr(&e, "and").as_deref() != Some("1");
                    }
                    "row" => {
                        in_row = true;
                        current_cells.clear();
                        let row_num = get_attr(&e, "r")
                            .and_then(|v| v.parse::<u32>().ok())
                            .unwrap_or(current_row_index + 1);
                        current_row_index = row_num;
                        if is_true(get_attr(&e, "customFormat").as_deref()) {
                            if let Some(si) = get_attr(&e, "s").and_then(|value| value.parse::<usize>().ok()) {
                                band_styles.rows.insert(row_num, resolve_cell_style(si, stylesheet));
                            }
                        }

                        // A row states its height whenever it is not the
                        // default one. customHeight says whether that number
                        // is pinned; without it the number is only a cache
                        // from the machine that wrote the file, and Excel
                        // works the height out again from the row's content.
                        current_row_height =
                            get_attr(&e, "ht").and_then(|v| v.parse::<f64>().ok())
                                .map(crate::row_height::file_points);
                        current_row_custom =
                            is_true(get_attr(&e, "customHeight").as_deref());
                        current_row_format_recorded = get_attr(&e, "customFormat").map(|value| is_true(Some(value.as_str())));
                        current_row_font = row_style_font(&e, stylesheet);
                        current_row_thick_top =
                            is_true(get_attr(&e, "thickTop").as_deref());
                        current_row_thick_bottom =
                            is_true(get_attr(&e, "thickBot").as_deref());
                        current_row_hidden = is_true(get_attr(&e, "hidden").as_deref());
                    }
                    "c" if in_row => {
                        in_cell = true;
                        value_text.clear();
                        formula_text.clear();
                        in_formula = false;
                        cell_array_ref = None;
                        cell_type = get_attr(&e, "t");
                        cell_style_index =
                            get_attr(&e, "s").and_then(|v| v.parse::<usize>().ok());
                        let cell_ref = get_attr(&e, "r").unwrap_or_default();
                        let (col, _) = parse_cell_ref(&cell_ref);
                        cell_col = col;
                        if col + 1 > max_col {
                            max_col = col + 1;
                        }
                    }
                    "f" if in_cell => {
                        in_formula = true;
                        formula_text.clear();
                        if get_attr(&e, "t").as_deref() == Some("array") {
                            cell_array_ref = get_attr(&e, "ref");
                        }
                    }
                    "v" if in_cell => {
                        in_value = true;
                        value_text.clear();
                    }
                    "is" if in_cell => {
                        in_inline = true;
                        in_inline_phonetic = false;
                        inline_start = Some(oxidocs_common::xml_utils::xml_byte_position(xml, reader.buffer_position()));
                        inline_runs.clear();
                        inline_units = None;
                        value_text.clear();
                    }
                    "rPh" if in_inline => in_inline_phonetic = true,
                    "t" if in_inline && !in_inline_phonetic => in_inline_text = true,
                    "cols" => {
                        // We'll handle col elements inside
                    }
                    "mergeCells" => {
                        in_merge_cells = true;
                    }
                    _ => {}
                }
            }
            Event::End(e) => {
                let name = local_name(e.name().as_ref());
                match name.as_str() {
                    "extLst" => conditional_extension_depth = conditional_extension_depth.saturating_sub(1),
                    "id" if in_visual_extension_id && e.name().as_ref() == b"x14:id" => in_visual_extension_id = false,
                    "colorScale" | "dataBar" | "iconSet" if conditional_extension_depth == 0 => {
                        if name == "colorScale" { in_color_scale = false; }
                        if visual_start.as_ref().is_some_and(|(_, tag)| *tag == name) {
                            let (start, _) = visual_start.take().unwrap();
                            if let Some(held) = cf_rule.as_mut() { held.visual_xml = Some(xml[start..oxidocs_common::xml_utils::xml_byte_position(xml, reader.buffer_position())].to_string()); }
                        }
                    }
                    "formula" if conditional_extension_depth == 0 => {
                        if let Some(held) = cf_rule.as_mut() {
                            held.formulas.push(std::mem::take(&mut cf_text));
                        }
                        in_cf_formula = false;
                    }
                    "cfRule" if conditional_extension_depth == 0 => {
                        if let Some(held) = cf_rule.take() {
                            conditional_rules.push(held);
                        }
                    }
                    "conditionalFormatting" if conditional_extension_depth == 0 => cf_ranges.clear(),
                    "formula1" | "formula2" => {
                        if let Some(rule) = rule.as_mut() {
                            let held = std::mem::take(&mut rule_text);
                            if rule_formula == 1 {
                                rule.formula1 = Some(held);
                            } else if rule_formula == 2 {
                                rule.formula2 = Some(held);
                            }
                        }
                        rule_formula = 0;
                    }
                    "dataValidation" => {
                        if let Some(rule) = rule.take() {
                            validations.push(rule);
                        }
                    }
                    "filterColumn" => {
                        if let (Some(field), Some(filter)) =
                            (filter_field.take(), auto_filter.as_mut())
                        {
                            if !filter_criteria.is_empty() {
                                filter.columns.push(crate::ir::AutoFilterColumn {
                                    field,
                                    criteria: std::mem::take(&mut filter_criteria),
                                    either: filter_either,
                                });
                            }
                        }
                    }
                    "row" => {
                        in_row = false;
                        rows.push(Row {
                            index: current_row_index,
                            cells: std::mem::take(&mut current_cells),
                            height: current_row_height,
                            custom_height: current_row_custom,
                            format_recorded: current_row_format_recorded, style_font: current_row_font.take(),
                            thick_top: current_row_thick_top,
                            thick_bottom: current_row_thick_bottom,
                            hidden: current_row_hidden,
                        });
                        current_row_height = None;
                        current_row_custom = false;
                        current_row_format_recorded = None;
                        current_row_thick_top = false;
                        current_row_thick_bottom = false;
                        current_row_hidden = false;
                    }
                    "c" => {
                        if in_cell {
                            let cell_value = if cell_type.as_deref() == Some("inlineStr") {
                                inline_units.take().map(CellValue::from_utf16)
                                    .unwrap_or_else(|| resolve_cell_value(&value_text, &cell_type, shared_strings))
                            } else {
                                resolve_cell_value(&value_text, &cell_type, shared_strings)
                            };
                            // A cell without an s wears the workbook's default
                            // format, which is cellXfs[0] 遯ｶ繝ｻnot no format at all.
                            let style =
                                resolve_cell_style(cell_style_index.unwrap_or(0), stylesheet);
                            let units = crate::string_value::decode(&formula_text);
                            let formula_utf16 = String::from_utf16(&units).err().map(|_| units.clone());
                            let formula = (!units.is_empty()).then(|| String::from_utf16_lossy(&units));
                            if let (Some(reference), Some(formula)) =
                                (cell_array_ref.take(), formula.as_ref())
                            {
                                array_anchors.push((
                                    current_row_index,
                                    cell_col,
                                    reference,
                                    formula.clone(),
                                    formula_utf16.clone(),
                                ));
                            }
                            current_cells.push(Cell {
                                formula_utf16,
                                array_block: None,
                                spill: Default::default(),
                                col: cell_col,
                                value: cell_value,
                                style,
                                formula,
                                runs: if cell_type.as_deref()==Some("inlineStr") { std::mem::take(&mut inline_runs) } else { runs_of(&value_text, &cell_type, shared_strings) },
                            });
                            in_cell = false;
                            in_formula = false;
                            cell_type = None;
                            cell_style_index = None;
                        }
                    }
                    "f" => {
                        in_formula = false;
                    }
                    "v" => {
                        in_value = false;
                    }
                    "is" => {
                        if let Some(start)=inline_start.take() {
                            let through_end=&xml[start..oxidocs_common::xml_utils::xml_byte_position(xml, reader.buffer_position())];
                            if let Some(end)=through_end.rfind('<') {
                                let fragment=format!("<sst><si>{}</si></sst>",&through_end[..end]);
                                if let Some(mut held) = parse_shared_strings(&fragment)?.into_iter().next() {
                                    resolve_run_font_references(&mut held.runs,theme);
                                    inline_runs = held.runs;
                                    inline_units = held.units;
                                }
                            }
                        }
                        in_inline = false;
                        in_inline_text = false;
                    }
                    "rPh" => in_inline_phonetic = false,
                    "t" => in_inline_text = false,
                    "mergeCells" => {
                        in_merge_cells = false;
                    }
                    _ => {}
                }
            }
            Event::Text(e) => {
                if in_visual_extension_id {
                    if let Some(held) = cf_rule.as_mut() { held.visual_extension_id = Some(e.unescape()?.into_owned()); }
                } else if in_cf_formula {
                    cf_text.push_str(&e.unescape()?);
                } else if rule_formula > 0 {
                    rule_text.push_str(&e.unescape()?);
                } else if in_formula {
                    formula_text.push_str(&e.unescape()?);
                } else if in_value || in_inline_text {
                    let text = e.unescape()?;
                    value_text.push_str(&text);
                }
            }
            Event::Empty(e) => {
                let name = local_name(e.name().as_ref());
                // A frozen pane IS read; only a draggable split is not.
                if name != "pane"
                    || !matches!(
                        get_attr(&e, "state").as_deref(),
                        Some("frozen") | Some("frozenSplit")
                    )
                {
                    note_unsupported(&name, &mut unsupported);
                }
                match name.as_str() {
                    "dataValidation" => validations.push(validation_from(&e)),
                    "cfRule" if conditional_extension_depth == 0 => {
                        conditional_rules.push(conditional_from(&e, &cf_ranges, stylesheet));
                    }
                    "cfvo" if in_color_scale && conditional_extension_depth == 0 => {
                        if let Some(held) = cf_rule.as_mut() { held.color_scale_points = Some(held.color_scale_points.unwrap_or(0).saturating_add(1)); }
                    }
                    "colorScale" | "dataBar" | "iconSet" if conditional_extension_depth == 0 && cf_rule.as_ref().is_some_and(|r| r.kind == name) => {
                        if let Some(held) = cf_rule.as_mut() { held.visual_xml = Some(xml[event_start..oxidocs_common::xml_utils::xml_byte_position(xml, reader.buffer_position())].to_string()); }
                    }
                    // <sheetProtection sheet="1"> is a sheet that is protected.
                    // The same element without it only records the options of
                    // one that is not, which Excel writes just as readily.
                    "sheetProtection" => {
                        protected = is_true(get_attr(&e, "sheet").as_deref());
                        protection = Some(sheet_protection_from(&e));
                    }
                    // <tabColor rgb="FF00B050"/> -- the sheet tab's colour.
                    "tabColor" => {
                        tab_color = parse_color_attr(&e, theme);
                    }
                    // <pane xSplit="2" ySplit="3" topLeftCell="C4" state="frozen"/>
                    "pane" => {
                        if matches!(
                            get_attr(&e, "state").as_deref(),
                            Some("frozen") | Some("frozenSplit")
                        ) {
                            frozen_cols = get_attr(&e, "xSplit")
                                .and_then(|v| v.parse::<f32>().ok())
                                .map(|v| v as u32)
                                .unwrap_or(0);
                            frozen_rows = get_attr(&e, "ySplit")
                                .and_then(|v| v.parse::<f32>().ok())
                                .map(|v| v as u32)
                                .unwrap_or(0);
                        }
                    }
                    // Handle self-closing <c .../> (cell with no value)
                    "c" if in_row => {
                        let cell_ref = get_attr(&e, "r").unwrap_or_default();
                        let (col, _) = parse_cell_ref(&cell_ref);
                        if col + 1 > max_col {
                            max_col = col + 1;
                        }
                        let si =
                            get_attr(&e, "s").and_then(|v| v.parse::<usize>().ok());
                        let style = resolve_cell_style(si.unwrap_or(0), stylesheet);
                        current_cells.push(Cell {
                            formula_utf16: None,
                            array_block: None,
                            spill: Default::default(),
                            col,
                            value: CellValue::Empty,
                            style,
                            formula: None,
                            runs: Vec::new(),
                        });
                    }

                    // <sheetFormatPr defaultRowHeight="15" defaultColWidth="8.43" ... />
                    "sheetFormatPr" => {
                        if let Some(v) = get_attr(&e, "defaultRowHeight") {
                            if let Ok(h) = v.parse::<f32>() {
                                default_row_height = h;
                            }
                        }
                        default_row_custom =
                            is_true(get_attr(&e, "customHeight").as_deref());
                        if let Some(v) = get_attr(&e, "defaultColWidth") {
                            if let Ok(w) = v.parse::<f32>() {
                                default_col_width = w;
                                default_col_width_explicit = true;
                            }
                        }
                    }

                    // <col min="1" max="3" width="12.5" ... />
                    "dimension" => {
                        declared_range = get_attr(&e, "ref")
                            .as_deref()
                            .and_then(parse_range_ref)
                            .map(|(start_col, start_row, end_col, end_row)| {
                                (start_row, start_col, end_row, end_col)
                            });
                    }
                    "autoFilter" => {
                        if let Some(reference) = get_attr(&e, "ref") {
                            if let Some((start_col, start_row, end_col, end_row)) =
                                parse_range_ref(&reference)
                            {
                                auto_filter = Some(crate::ir::AutoFilter {
                                    start_row,
                                    start_col,
                                    end_row,
                                    end_col,
                                    columns: Vec::new(),
                                });
                            }
                        }
                    }
                    "filterColumn" => {
                        // colId counts from zero within the filtered range;
                        // Field counts from one.
                        filter_field = get_attr(&e, "colId")
                            .and_then(|value| value.parse::<u32>().ok())
                            .map(|col| col + 1);
                        filter_criteria.clear();
                        filter_either = false;
                    }
                    "filters" => {
                        // A list of values is an "is one of these" test.
                        filter_either = true;
                    }
                    "filter" => {
                        if let Some(value) = get_attr(&e, "val") {
                            filter_criteria.push(value);
                        }
                    }
                    "customFilter" => {
                        let operator = get_attr(&e, "operator").unwrap_or_default();
                        let value = get_attr(&e, "val").unwrap_or_default();
                        filter_criteria.push(format!("{}{value}", filter_operator(&operator)));
                    }
                    "customFilters" => {
                        filter_either = get_attr(&e, "and").as_deref() != Some("1");
                    }
                    "col" => {
                        let min_col = get_attr(&e, "min")
                            .and_then(|v| v.parse::<u32>().ok())
                            .unwrap_or(1);
                        let max_col_attr = get_attr(&e, "max")
                            .and_then(|v| v.parse::<u32>().ok())
                            .unwrap_or(min_col);
                        let width = get_attr(&e, "width")
                            .and_then(|v| v.parse::<f32>().ok())
                            .unwrap_or(default_col_width);

                        // Ensure col_widths vec is large enough (0-based)
                        let needed = max_col_attr as usize;
                        if col_widths.len() < needed {
                            col_widths.resize(needed, 0.0);
                        }
                        for c in min_col..=max_col_attr {
                            col_widths[(c - 1) as usize] = width;
                            col_custom_widths.insert(c - 1, is_true(get_attr(&e, "customWidth").as_deref()));
                            if is_true(get_attr(&e, "bestFit").as_deref()) { best_fit_cols.insert(c - 1); }
                        }
                        if is_true(get_attr(&e, "hidden").as_deref()) {
                            for c in min_col..=max_col_attr {
                                hidden_cols.push(c - 1);
                            }
                        }
                        // A column's style puts a font on every cell in it,
                        // blank ones included, and that is what a row's
                        // height is measured from when nothing in the row
                        // is taller.
                        if let Some(si) =
                            get_attr(&e, "style").and_then(|v| v.parse::<usize>().ok())
                        {
                            let style = resolve_cell_style(si, stylesheet);
                            for column in min_col..=max_col_attr {
                                band_styles.columns.insert(column - 1, style.clone());
                            }
                            if let (Some(name), Some(size)) =
                                (style.font_name, style.font_size)
                            {
                                col_fonts.push((min_col - 1, max_col_attr - 1, name, size));
                            }
                        }
                    }

                    // <mergeCell ref="A1:C3"/>
                    "mergeCell" if in_merge_cells => {
                        if let Some(ref_str) = get_attr(&e, "ref") {
                            if let Some((sc, sr, ec, er)) = parse_range_ref(&ref_str) {
                                merge_cells.push(MergeCell {
                                    start_row: sr,
                                    start_col: sc,
                                    end_row: er,
                                    end_col: ec,
                                });
                            }
                        }
                    }

                    // Self-closing <row ... /> (empty row with attributes)
                    "row" => {
                        let row_num = get_attr(&e, "r")
                            .and_then(|v| v.parse::<u32>().ok())
                            .unwrap_or(current_row_index + 1);
                        current_row_index = row_num;
                        if is_true(get_attr(&e, "customFormat").as_deref()) {
                            if let Some(si) = get_attr(&e, "s").and_then(|value| value.parse::<usize>().ok()) {
                                band_styles.rows.insert(row_num, resolve_cell_style(si, stylesheet));
                            }
                        }
                        rows.push(Row {
                            index: row_num,
                            cells: Vec::new(),
                            height: get_attr(&e, "ht")
                                .and_then(|v| v.parse::<f64>().ok())
                                .map(crate::row_height::file_points),
                            custom_height:
                                is_true(get_attr(&e, "customHeight").as_deref()),
                            format_recorded: get_attr(&e, "customFormat").map(|value| is_true(Some(value.as_str()))), style_font: row_style_font(&e, stylesheet),
                            thick_top: is_true(get_attr(&e, "thickTop").as_deref()),
                            thick_bottom: is_true(get_attr(&e, "thickBot").as_deref()),
                            hidden: is_true(get_attr(&e, "hidden").as_deref()),
                        });
                    }

                    _ => {}
                }
            }
            Event::Eof => break,
            _ => {}
        }
    }

    // Physical empty cells still contribute to the imported used extent.
    // The host refreshes this extent when UsedRange is explicitly requested.
    let pending_used_bounds = rows.iter().flat_map(|row| row.cells.iter().map(move |cell| (cell.col, row.index.saturating_sub(1))))
        .filter(|(col, row)| *col <= oxicells_calc::reference::MAX_COL && *row <= oxicells_calc::reference::MAX_ROW)
        .fold(None, |bounds: Option<(u32, u32, u32, u32)>, (col, row)| Some(match bounds {
            Some((left, top, right, bottom)) => (left.min(col), top.min(row), right.max(col), bottom.max(row)),
            None => (col, row, col, row),
        }));

    // Ensure col_widths covers all columns
    let col_count = max_col as usize;
    if col_widths.len() < col_count {
        // A column with no <col> entry of its own keeps a width of zero. The
        // sheet's default is measured in plain characters while a stated width
        // carries the cell's gutter, so a reader has to be able to tell them
        // apart.
        col_widths.resize(col_count, 0.0);
    }

    // The Normal font, which every column no <col> dresses is written in.
    let normal_font = {
        let font_id = stylesheet
            .cell_style_xfs
            .first()
            .map(|xf| xf.font_id)
            .unwrap_or(0);
        stylesheet.fonts.get(font_id).and_then(|font| {
            match (font.name.clone(), font.size) {
                (Some(name), Some(size)) => Some((name, size)),
                _ => None,
            }
        })
    };
    let column_width_font = stylesheet.fonts.get(stylesheet.cell_style_xfs.first().map_or(0, |xf| xf.font_id))
        .and_then(|font| font.latin_name.clone().or_else(|| font.name.clone()));

    // Deal every array formula across its block: each member gets the
    // formula's text and the block, and a member the file left out -- a cell
    // with nothing to say for itself -- is made so it can be worked out.
    for (anchor_row, anchor_col, reference, formula, formula_utf16) in array_anchors {
        let block = match parse_range_ref(&reference) {
            Some((start_col, start_row, end_col, end_row)) => {
                (start_row, start_col, end_row, end_col)
            }
            // A single cell's array formula, whose `ref` is the cell itself.
            None => {
                let (col, row) = parse_cell_ref(&reference);
                (row + 1, col, row + 1, col)
            }
        };
        let (start_row, start_col, end_row, end_col) = block;
        if end_row < start_row || end_col < start_col || (end_row - start_row) * (end_col - start_col) > 1_000_000 {
            continue;
        }
        for row_index in start_row..=end_row {
            let row = match rows.iter().position(|row| row.index == row_index) {
                Some(at) => &mut rows[at],
                None => {
                    rows.push(Row {
                        index: row_index,
                        cells: Vec::new(),
                        height: None,
                        custom_height: false,
                        format_recorded: None, style_font: None,
                        thick_top: false,
                        thick_bottom: false,
                        hidden: false,
                    });
                    rows.sort_by_key(|row| row.index);
                    rows.iter_mut().find(|row| row.index == row_index).unwrap()
                }
            };
            for col in start_col..=end_col {
                let is_anchor = row_index == anchor_row && col == anchor_col;
                match row.cells.iter_mut().find(|cell| cell.col == col) {
                    Some(cell) => {
                        cell.array_block = Some(block);
                        if !is_anchor {
                            cell.formula = Some(formula.clone());
                            cell.formula_utf16 = formula_utf16.clone();
                        }
                    }
                    None => {
                        row.cells.push(Cell {
                            formula_utf16: formula_utf16.clone(),
                            col,
                            value: CellValue::Empty,
                            style: resolve_cell_style(0, stylesheet),
                            formula: Some(formula.clone()),
                            runs: Vec::new(),
                            array_block: Some(block),
                            spill: Default::default(),
                        });
                        row.cells.sort_by_key(|cell| cell.col);
                    }
                }
            }
        }
    }

    // A style explicitly covering every column is equivalent to a sheet-wide
    // base plus exceptions. Keep that base compact: otherwise reopening a
    // sheet-wide format creates 16,384 templates and applying one row format
    // materializes 16,384 unnecessary intersection cells.
    if band_styles.columns.len() == 16_384
        && band_styles.columns.first_key_value().map(|(key, _)| *key) == Some(0)
        && band_styles.columns.last_key_value().map(|(key, _)| *key) == Some(16_383)
    {
        // A majority candidate gives the common style when one exists; even
        // without a majority, retaining all different entries is equivalent.
        let mut candidate: Option<&CellStyle> = None;
        let mut balance = 0usize;
        for style in band_styles.columns.values() {
            if balance == 0 { candidate = Some(style); balance = 1; }
            else if candidate == Some(style) { balance += 1; }
            else { balance -= 1; }
        }
        let base = candidate.expect("a complete set is not empty").clone();
        band_styles.columns.retain(|_, style| *style != base);
        band_styles.sheet = Some(base);
    }

    // The first font in the list, which is what an indent is measured in.
    let first_font = stylesheet.fonts.first().and_then(|font| {
        match (font.name.clone(), font.size) {
            (Some(name), Some(size)) => Some((name, size)),
            _ => None,
        }
    });

    let visual_extensions = worksheet_visual_extensions(xml)?;
    for rule in &mut conditional_rules {
        if let Some(id) = &rule.visual_extension_id { rule.visual_extension_xml = visual_extensions.get(id).cloned(); }
    }
    for (id, raw) in &visual_extensions {
        if conditional_rules.iter().any(|rule| rule.visual_extension_id.as_deref() == Some(id)) { continue; }
        if let Some(rule) = extended_only_icon_rule(id, raw, stylesheet)? { conditional_rules.push(rule); }
    }
    Ok(Sheet {
        kind: Default::default(),
        tables: Vec::new(),
        drawings: Vec::new(),
        comments: Vec::new(),
        threaded_comments: Vec::new(),
        name: sheet_name.to_string(),
        // The tab strip's business, which the workbook states rather than the
        // sheet: the caller fills it in from there.
        visibility: crate::ir::Visibility::Visible,
        rows,
        col_count,
        col_widths,
        col_custom_widths,
        best_fit_cols,
        cached_spills: None,
        pending_used_bounds,
        column_width_font,
        default_col_width,
        default_col_width_explicit,
        default_row_height,
        default_row_custom,
        col_fonts,
        band_styles,
        normal_font,
        first_font,
        frozen_rows,
        frozen_cols,
        merge_cells,
        auto_filter,
        declared_range,
        hidden_cols: {
            // The order columns appear in says nothing; keep the list tidy.
            let mut hidden_cols = hidden_cols;
            hidden_cols.sort_unstable();
            hidden_cols.dedup();
            hidden_cols
        },
        unsupported_elements: unsupported,
        drawing_id_blocks: Vec::new(),
        header_footer: parse_header_footer(xml)?,
        protected,
        protection,
        validations,
        conditional_rules: {
            let mut held = conditional_rules;
            held.sort_by_key(|rule| rule.priority);
            held
        },
        // Resolved once the sheet's relationships are to hand, which is a
        // level up from here.
        hyperlinks: Vec::new(),
        form_controls: Vec::new(),
        tab_color,
        origin_id: None,
    })
}

/// Resolve a cell's raw value text + type attribute into a CellValue.
/// The dressing of the shared string a cell holds, if it holds one that is not
/// all dressed alike.
fn runs_of(
    value_text: &str,
    cell_type: &Option<String>,
    shared_strings: &[SharedString],
) -> Vec<crate::ir::TextRun> {
    if cell_type.as_deref() != Some("s") {
        return Vec::new();
    }
    value_text
        .parse::<usize>()
        .ok()
        .and_then(|at| shared_strings.get(at))
        .map(|held| held.runs.clone())
        .unwrap_or_default()
}

fn resolve_cell_value(
    value_text: &str,
    cell_type: &Option<String>,
    shared_strings: &[SharedString],
) -> CellValue {
    if value_text.is_empty() && cell_type.is_none() {
        return CellValue::Empty;
    }

    match cell_type.as_deref() {
        Some("s") => {
            // Shared string index
            if let Ok(idx) = value_text.parse::<usize>() {
                if idx < shared_strings.len() {
                    CellValue::from_utf16(shared_strings[idx].units.clone().unwrap_or_else(|| crate::string_value::decode(&shared_strings[idx].text)))
                } else {
                    CellValue::Error(format!("Invalid SST index: {}", idx))
                }
            } else {
                CellValue::Error(format!("Non-numeric SST index: {}", value_text))
            }
        }
        Some("b") => {
            CellValue::Boolean(value_text == "1" || value_text.eq_ignore_ascii_case("true"))
        }
        Some("e") => CellValue::Error(value_text.to_string()),
        Some("str") | Some("inlineStr") => {
            // Inline string or formula string result
            CellValue::from_utf16(crate::string_value::decode(value_text))
        }
        _ => {
            // No type attribute means number
            if value_text.is_empty() {
                CellValue::Empty
            } else if let Ok(n) = value_text.parse::<f64>() {
                CellValue::Number(n)
            } else {
                CellValue::from_utf16(crate::string_value::decode(value_text))
            }
        }
    }
}

/// Direct XML children, including each child's original bytes. Metadata
/// indices count records, including records whose payload we do not support.
pub(crate) fn metadata_elements(xml: &str) -> Result<Vec<(quick_xml::events::BytesStart<'static>, &str)>, XlsxError> {
    let mut reader = Reader::from_str(xml);
    let mut depth = 0usize;
    let mut current = None;
    let mut result = Vec::new();
    loop {
        let start = oxidocs_common::xml_utils::xml_byte_position(xml, reader.buffer_position());
        match reader.read_event().map_err(XlsxError::Xml)? {
            Event::Start(e) => {
                if depth == 1 { current = Some((start, e.into_owned())); }
                depth += 1;
            }
            Event::Empty(e) if depth == 1 => result.push((e.into_owned(), &xml[start..oxidocs_common::xml_utils::xml_byte_position(xml, reader.buffer_position())])),
            Event::End(_) => {
                depth = depth.saturating_sub(1);
                if depth == 1 {
                    if let Some((start, tag)) = current.take() { result.push((tag, &xml[start..oxidocs_common::xml_utils::xml_byte_position(xml, reader.buffer_position())])); }
                }
            }
            Event::Eof => break,
            _ => {}
        }
    }
    Ok(result)
}

const RICH_ERRORS: [&str; 6] = ["#SPILL!", "#CONNECT!", "#BLOCKED!", "#UNKNOWN!", "#FIELD!", "#CALC!"];

pub(crate) fn rich_error_type(error: &str) -> Option<u32> {
    RICH_ERRORS.iter().position(|held| *held == error).map(|index| index as u32 + 8)
}

#[derive(Default)]
pub(crate) struct RichErrorMetadata {
    pub type_index: Option<usize>,
    pub type_count: usize,
    pub future_count: usize,
    pub blocks: Vec<String>,
    /// One-based value-metadata indices, not rich-value record indices.
    pub errors: std::collections::BTreeMap<usize, &'static str>,
}

pub(crate) fn rich_error_metadata(xml: &str, structures: &str, values: &str) -> Result<RichErrorMetadata, XlsxError> {
    let sections = metadata_elements(xml)?;
    let mut result = RichErrorMetadata::default();
    if let Some((_, types)) = sections.iter().find(|(tag, _)| local_name(tag.name().as_ref()) == "metadataTypes") {
        for (tag, _) in metadata_elements(types)?.into_iter().filter(|(tag, _)| local_name(tag.name().as_ref()) == "metadataType") {
            result.type_count += 1;
            if get_attr(&tag, "name").as_deref() == Some("XLRICHVALUE") { result.type_index = Some(result.type_count); }
        }
    }
    // An error's fields are named by its structure. Offsets and subtype can
    // precede errorType, and other rich-value types use the same record table.
    let mut error_fields = Vec::new();
    for (tag, structure) in metadata_elements(structures)?.into_iter().filter(|(tag, _)| local_name(tag.name().as_ref()) == "s") {
        let keys = metadata_elements(structure)?;
        error_fields.push((get_attr(&tag, "t").as_deref() == Some("_error"))
            .then(|| keys.iter().filter(|(tag, _)| local_name(tag.name().as_ref()) == "k")
                .position(|(tag, _)| get_attr(tag, "n").as_deref() == Some("errorType") && get_attr(tag, "t").as_deref() == Some("i")))
            .flatten());
    }
    let mut errors = Vec::new();
    for (tag, value) in metadata_elements(values)?.into_iter().filter(|(tag, _)| local_name(tag.name().as_ref()) == "rv") {
        let field = get_attr(&tag, "s").and_then(|s| s.parse::<usize>().ok()).and_then(|index| error_fields.get(index)).copied().flatten();
        let fields = metadata_elements(value)?;
        let error = field.and_then(|field| fields.iter().filter(|(tag, _)| local_name(tag.name().as_ref()) == "v").nth(field))
            .and_then(|(_, xml)| {
                let mut reader = Reader::from_str(xml);
                match reader.read_event().ok()? {
                    Event::Start(tag) => reader.read_text(tag.name()).ok()?.trim().parse::<usize>().ok(),
                    _ => None,
                }
            }).and_then(|kind| kind.checked_sub(8)).and_then(|index| RICH_ERRORS.get(index)).copied();
        errors.push(error);
    }
    let mut future = Vec::new();
    for (_, section) in sections.iter().filter(|(tag, _)| local_name(tag.name().as_ref()) == "futureMetadata" && get_attr(tag, "name").as_deref() == Some("XLRICHVALUE")) {
        for (_, record) in metadata_elements(section)?.into_iter().filter(|(tag, _)| local_name(tag.name().as_ref()) == "bk") {
            let mut reader = Reader::from_str(record);
            let mut value = None;
            loop {
                match reader.read_event().map_err(XlsxError::Xml)? {
                    Event::Start(tag) | Event::Empty(tag) if local_name(tag.name().as_ref()) == "rvb" => {
                        value = get_attr(&tag, "i").and_then(|s| s.parse::<usize>().ok()).and_then(|index| errors.get(index)).copied().flatten();
                    }
                    Event::Eof => break,
                    _ => {}
                }
            }
            future.push(value);
        }
    }
    result.future_count = future.len();
    if let Some((_, section)) = sections.iter().find(|(tag, _)| local_name(tag.name().as_ref()) == "valueMetadata") {
        for (_, record) in metadata_elements(section)?.into_iter().filter(|(tag, _)| local_name(tag.name().as_ref()) == "bk") {
            result.blocks.push(record.to_string());
            for (tag, _) in metadata_elements(record)? {
                if result.type_index.is_some() && local_name(tag.name().as_ref()) == "rc" && get_attr(&tag, "t").and_then(|s| s.parse().ok()) == result.type_index {
                    if let Some(error) = get_attr(&tag, "v").and_then(|s| s.parse::<usize>().ok()).and_then(|index| future.get(index)).copied().flatten() {
                        result.errors.insert(result.blocks.len(), error);
                    }
                }
            }
        }
    }
    Ok(result)
}

fn restore_rich_errors(sheet: &mut Sheet, xml: &str, metadata: &RichErrorMetadata) -> Result<(), XlsxError> {
    if metadata.errors.is_empty() { return Ok(()); }
    let mut reader = Reader::from_str(xml);
    let mut path = Vec::new();
    let mut errors = HashMap::new();
    loop {
        let event = reader.read_event().map_err(XlsxError::Xml)?;
        match &event {
            Event::Start(tag) | Event::Empty(tag) => {
                let name = local_name(tag.name().as_ref());
                if name == "c" && path.last().map(String::as_str) == Some("row")
                    && path.iter().rev().nth(1).map(String::as_str) == Some("sheetData") && get_attr(tag, "t").as_deref() == Some("e") {
                    if let Some(error) = get_attr(tag, "vm").and_then(|s| s.parse::<usize>().ok()).and_then(|index| metadata.errors.get(&index)) {
                        if let Some(at) = get_attr(tag, "r").and_then(|s| oxicells_calc::reference::parse_a1(&s)) { errors.insert((at.col, at.row), *error); }
                    }
                }
                if matches!(&event, Event::Start(_)) { path.push(name); }
            }
            Event::End(_) => { path.pop(); }
            Event::Eof => break,
            _ => {}
        }
    }
    for row in &mut sheet.rows {
        for cell in &mut row.cells {
            if let Some(error) = errors.get(&(cell.col, row.index.saturating_sub(1))) { cell.value = CellValue::Error((*error).to_string()); }
        }
    }
    Ok(())
}

#[derive(Default)]
pub(crate) struct DynamicMetadata {
    pub type_index: Option<usize>,
    pub type_count: usize,
    pub future_count: usize,
    pub cell_count: usize,
    pub cells: std::collections::BTreeSet<usize>,
    pub cell_blocks: Vec<String>,
}

/// Resolve cell metadata records through their type and future-record indices.
/// Other metadata types can precede XLDAPR and share a cell's metadata block.
pub(crate) fn dynamic_metadata(xml: &str) -> Result<DynamicMetadata, XlsxError> {
    let mut reader = Reader::from_str(xml);
    let mut path = Vec::<String>::new();
    let mut types = Vec::new();
    let mut future = Vec::new();
    let mut records = Vec::<Vec<(usize, usize)>>::new();
    let mut in_future = false;
    let mut future_record = None;
    let mut cell_record = None::<Vec<(usize, usize)>>;
    let mut cell_start = 0;
    let mut cell_blocks = Vec::new();
    loop {
        let start = oxidocs_common::xml_utils::xml_byte_position(xml, reader.buffer_position());
        let event = reader.read_event().map_err(XlsxError::Xml)?;
        match &event {
            Event::Start(e) | Event::Empty(e) => {
                let name = local_name(e.name().as_ref());
                let empty = matches!(&event, Event::Empty(_));
                match (path.last().map(String::as_str), name.as_str()) {
                    (Some("metadataTypes"), "metadataType") => types.push(get_attr(e, "name").unwrap_or_default()),
                    (Some("metadata"), "futureMetadata") => in_future = get_attr(e, "name").as_deref() == Some("XLDAPR"),
                    (Some("futureMetadata"), "bk") if in_future => {
                        if empty { future.push(false); } else { future_record = Some(false); }
                    }
                    (Some("cellMetadata"), "bk") => {
                        if empty {
                            records.push(Vec::new());
                            cell_blocks.push(xml[start..oxidocs_common::xml_utils::xml_byte_position(xml, reader.buffer_position())].to_string());
                        } else {
                            cell_record = Some(Vec::new());
                            cell_start = start;
                        }
                    }
                    (Some("bk"), "rc") if cell_record.is_some() => {
                        if let (Some(t), Some(v)) = (get_attr(e, "t").and_then(|s| s.parse().ok()), get_attr(e, "v").and_then(|s| s.parse().ok())) {
                            cell_record.as_mut().unwrap().push((t, v));
                        }
                    }
                    (_, "dynamicArrayProperties") if future_record.is_some() => {
                        future_record = Some(is_true(get_attr(e, "fDynamic").as_deref()));
                    }
                    _ => {}
                }
                if !empty { path.push(name); }
                else if name == "futureMetadata" { in_future = false; }
            }
            Event::End(e) => {
                let name = local_name(e.name().as_ref());
                if name == "bk" {
                    if path.iter().rev().nth(1).map(String::as_str) == Some("cellMetadata") {
                        if let Some(record) = cell_record.take() {
                            records.push(record);
                            cell_blocks.push(xml[cell_start..oxidocs_common::xml_utils::xml_byte_position(xml, reader.buffer_position())].to_string());
                        }
                    } else if path.iter().rev().nth(1).map(String::as_str) == Some("futureMetadata") {
                        if let Some(dynamic) = future_record.take() { future.push(dynamic); }
                    }
                }
                if name == "futureMetadata" { in_future = false; }
                path.pop();
            }
            Event::Eof => break,
            _ => {}
        }
    }
    let type_index = types.iter().position(|s| s == "XLDAPR").map(|i| i + 1);
    let cells = records.iter().enumerate().filter_map(|(i, record)| {
        record.iter().any(|(t, v)| Some(*t) == type_index && future.get(*v) == Some(&true)).then_some(i + 1)
    }).collect();
    Ok(DynamicMetadata { type_index, type_count: types.len(), future_count: future.len(), cell_count: records.len(), cells, cell_blocks })
}

fn restore_dynamic_arrays(sheet: &mut Sheet, xml: &str, metadata: &DynamicMetadata) -> Result<(), XlsxError> {
    if metadata.cells.is_empty() { return Ok(()); }
    let mut reader = Reader::from_str(xml);
    let mut anchor = None;
    let mut blocks = Vec::new();
    loop {
        match reader.read_event().map_err(XlsxError::Xml)? {
            Event::Start(e) | Event::Empty(e) => match local_name(e.name().as_ref()).as_str() {
                "c" => {
                    anchor = get_attr(&e, "cm").and_then(|s| s.parse().ok()).filter(|id| metadata.cells.contains(id))
                        .and_then(|_| get_attr(&e, "r")).and_then(|s| oxicells_calc::reference::parse_a1(&s));
                }
                "f" if anchor.is_some() && get_attr(&e, "t").as_deref() == Some("array") => {
                    if let Some(reference) = get_attr(&e, "ref") {
                        let block = parse_range_ref(&reference).or_else(|| oxicells_calc::reference::parse_a1(&reference)
                            .map(|r| (r.col, r.row + 1, r.col, r.row + 1)));
                        if let Some((left, top, right, bottom)) = block {
                            let at = anchor.as_ref().unwrap();
                            if left <= right && top >= 1 && top <= bottom && right <= oxicells_calc::reference::MAX_COL
                                && bottom <= oxicells_calc::reference::MAX_ROW + 1 && at.col == left && at.row + 1 == top {
                                blocks.push(((at.col, at.row), (left, top, right, bottom)));
                            }
                        }
                    }
                }
                _ => {}
            },
            Event::End(e) if local_name(e.name().as_ref()) == "c" => anchor = None,
            Event::Eof => break,
            _ => {}
        }
    }
    if blocks.is_empty() { return Ok(()); }
    let mut cached = Vec::new();
    for (anchor, (left, top, right, bottom)) in blocks {
        cached.push((anchor, (right - left + 1, bottom - top + 1)));
        for row in sheet.rows.iter_mut().filter(|row| (top..=bottom).contains(&row.index)) {
            for cell in row.cells.iter_mut().filter(|cell| (left..=right).contains(&cell.col)) {
                cell.array_block = None;
                if (cell.col, row.index - 1) == anchor {
                    cell.spill = crate::ir::Spill::Dynamic;
                } else {
                    cell.spill = crate::ir::Spill::Member;
                    cell.formula = None;
                    cell.formula_utf16 = None;
                }
            }
        }
    }
    sheet.cached_spills = Some(cached);
    Ok(())
}

/// Parse an .xlsx file from raw bytes into a Workbook IR.
///
/// The values Excel cached in the file are **kept**. For display they are the
/// correct answer 遯ｶ繝ｻrecalculating a freshly loaded workbook can only introduce
/// divergence from what the user sees in Excel. Only formula cells that arrive
/// without a cached value are computed.
///
/// Use [`crate::formula::evaluate_workbook_formulas`] after editing a workbook,
/// and [`parse_xlsx_preserving_values`] when even the gap filling is unwanted.
pub fn parse_xlsx(data: &[u8]) -> Result<Workbook, XlsxError> {
    let mut workbook = parse_xlsx_preserving_values(data)?;
    crate::formula::fill_missing_formula_values(&mut workbook);
    Ok(workbook)
}

// The default-theme version can select Excel's built-in font scheme even
// when a stale theme part still contains different faces. Version202300 and
// explicit custom version0 are compared against saved Native workbooks.
fn apply_workbook_default_font_scheme(theme: &mut Theme, workbook_xml: &str) {
    let mut reader = Reader::from_str(workbook_xml);
    loop {
        match reader.read_event() {
            Ok(Event::Start(e) | Event::Empty(e)) if e.local_name().as_ref() == b"workbookPr" => {
                if get_attr(&e, "defaultThemeVersion").as_deref() == Some("202300") {
                    theme.major_face = Some("游ゴシック Light".into());
                    theme.minor_face = Some("游ゴシック".into());
                    theme.major_latin_face = Some("Aptos Display".into());
                    theme.minor_latin_face = Some("Aptos Narrow".into());
                }
                break;
            }
            Ok(Event::Eof) | Err(_) => break,
            _ => {}
        }
    }
}

/// Parse an .xlsx file without recalculating anything.
///
/// Every formula cell keeps the value Excel last computed for it, alongside its
/// formula text. That pair is what makes a workbook usable as a test oracle.
pub fn parse_xlsx_preserving_values(data: &[u8]) -> Result<Workbook, XlsxError> {
    let mut archive = OoxmlArchive::new(data)?;

    // 1. Parse shared strings (optional 遯ｶ繝ｻsome xlsx files have none)
    let mut shared_strings = match archive.try_read_part("xl/sharedStrings.xml")? {
        Some(xml) => parse_shared_strings(&xml)?,
        None => Vec::new(),
    };

    let workbook_xml = archive.read_part("xl/workbook.xml")?;

    // 2. Parse styles.xml (optional 遯ｶ繝ｻsome simple xlsx have none)
    // The theme has to be read first: a style may name one of its colours
    // rather than state one of its own.
    let mut theme = match archive.try_read_part("xl/theme/theme1.xml")? {
        Some(xml) => parse_theme_xml(&xml),
        None => Theme::default(),
    };
    apply_workbook_default_font_scheme(&mut theme, &workbook_xml);
    let stylesheet = match archive.try_read_part("xl/styles.xml")? {
        Some(xml) => {
            // A hundred of the corpus's workbooks state a palette of their
            // own, and it sits at the foot of the same part as the styles
            // that name it 遯ｶ繝ｻso it is read in a pass of its own first.
            theme.indexed = indexed_palette(&xml);
            parse_styles_xml(&xml, &theme)?
        }
        None => StyleSheet::default(),
    };

    for string in &mut shared_strings {resolve_run_font_references(&mut string.runs,&theme);}

    // 3. Parse workbook.xml to get sheet names and rIds
    let sheet_infos = parse_workbook_sheets(&workbook_xml)?;

    // 4. Parse workbook relationships to map rIds to sheet file paths
    let rels_xml = archive.read_part("xl/_rels/workbook.xml.rels")?;
    let rels = parse_relationships(&rels_xml)?;
    let mut related_part = |kind: &str| -> Result<String, XlsxError> {
        match rels.values().find(|rel| rel.rel_type.ends_with(kind) && rel.target_mode.as_deref() != Some("External")) {
            Some(relation) => Ok(archive.try_read_part(&part_beside("xl/workbook.xml", &relation.target))?.unwrap_or_default()),
            None => Ok(String::new()),
        }
    };
    let metadata = related_part("/sheetMetadata")?;
    let dynamic = dynamic_metadata(&metadata)?;
    let structures = related_part("/rdRichValueStructure")?;
    let values = related_part("/rdRichValue")?;
    let rich_errors = rich_error_metadata(&metadata, &structures, &values)?;
    let persons = match rels.values().find(|rel| rel.rel_type.ends_with("/person")) {
        Some(relation) => {
            let path = part_beside("xl/workbook.xml", &relation.target);
            archive.try_read_part(&path)?.map(|xml| parse_person_list(&xml)).unwrap_or_default()
        }
        None => HashMap::new(),
    };

    // Build rId -> target path map
    let sheet_kinds: HashMap<String, crate::ir::SheetKind> = rels.iter().map(|(id,rel)| {
        let kind = if rel.rel_type.ends_with("/chartsheet") { crate::ir::SheetKind::Chart }
            else if rel.rel_type.ends_with("/xlIntlMacrosheet") { crate::ir::SheetKind::InternationalMacroSheet }
            else if rel.rel_type.ends_with("/xlMacrosheet") { crate::ir::SheetKind::MacroSheet }
            else { crate::ir::SheetKind::Worksheet };
        (id.clone(), kind)
    }).collect();
    let rid_to_path: HashMap<String, String> = rels
        .into_iter()
        .map(|(id, rel)| (id, rel.target))
        .collect();

    // 4b. The workbooks this one only links to, with the values it remembers
    // of them. Their order is what `[1]`, `[2]` 遯ｶ・ｦ mean in a formula.
    let mut external_books = Vec::new();
    for id in parse_external_references(&workbook_xml) {
        let Some(target) = rid_to_path.get(&id) else {
            continue;
        };
        let part = if let Some(rest) = target.strip_prefix('/') {
            rest.to_string()
        } else {
            format!("xl/{target}")
        };
        match archive.try_read_part(&part)? {
            Some(xml) => external_books.push(parse_external_link_with_identity(&mut archive,&part,&xml)?),
            None => external_books.push(ExternalBook::default()),
        }
    }

    // 5. Parse each worksheet
    let calculation = parse_calculation(&workbook_xml);
    let mut recalculation = crate::ir::RecalculationState::default();
    let mut sheets = Vec::new();
    let chart_sources=crate::chart_edit::chart_part_names(&mut archive)?;
    for info in &sheet_infos {
        let sheet_path = match rid_to_path.get(&info.r_id) {
            Some(target) => {
                // Target is relative to xl/, e.g. "worksheets/sheet1.xml"
                if target.starts_with('/') {
                    // Absolute path within archive (strip leading /)
                    target.trim_start_matches('/').to_string()
                } else {
                    format!("xl/{}", target)
                }
            }
            None => {
                log::warn!(
                    "No relationship found for sheet '{}' (rId={}), skipping",
                    info.name,
                    info.r_id
                );
                continue;
            }
        };

        match archive.try_read_part(&sheet_path)? {
            Some(sheet_xml) => {
                let mut sheet =
                    parse_worksheet(&sheet_xml, &info.name, &shared_strings, &stylesheet, &theme)?;
                restore_dynamic_arrays(&mut sheet, &sheet_xml, &dynamic)?;
                if calculation.mode == crate::ir::CalculationMode::Manual {
                    recalculation.dirty.extend(pending_array_calculations(&sheet_xml)?.into_iter()
                        .map(|cell| (info.name.clone(), cell)));
                }
                restore_rich_errors(&mut sheet, &sheet_xml, &rich_errors)?;
                sheet.kind=sheet_kinds.get(&info.r_id).copied().unwrap_or_default();
                // What the sheet's own parts turn out to hold, which is only
                // known once they have been read 遯ｶ繝ｻand which decides what the
                // walk past their names in the sheet XML really meant.
                let mut had_notes = false;
                let mut has_controls = false;
                // Whether any drawing part the sheet points at held an anchor
                // at all, which is what tells an unread drawing from no
                // drawing.
                let mut anchored = false;
                // Whether the tab is shown is the workbook's business, not the
                // worksheet part's.
                sheet.visibility = info.visibility;
                // Its position in the file is the identity an editor round-trip
                // keeps, so a rename is not mistaken for an add-and-delete.
                sheet.origin_id = Some(sheets.len() as u32);
                // A table lives in its own part, named by the sheet's own
                // relationships. Excel dresses the range from there, so no cell
                // inside carries the header fill or the banding.
                let rels_path = sheet_path
                    .rsplit_once('/')
                    .map(|(dir, file)| format!("{dir}/_rels/{file}.rels"))
                    .unwrap_or_default();
                if let Some(rels_xml) = archive.try_read_part(&rels_path)? {
                    let rels = parse_relationships(&rels_xml)?;
                    let ole_prog_ids = parse_ole_object_metadata(&sheet_xml, &rels);
                    // A note is two parts: its text, and the VML that says
                    // where its box is and whether the sheet shows it.
                    let beside = |ending: &str| {
                        rels.values()
                            .find(|rel| rel.rel_type.ends_with(ending))
                            .map(|rel| part_beside(&sheet_path, &rel.target))
                    };
                    // The legacy drawing is read whether or not notes come with
                    // it. The same part carries a sheet's buttons and tick
                    // boxes, and a sheet can hold a hundred of those and not a
                    // single note 遯ｶ繝ｻtwo of the four hundred workbooks swept do
                    // exactly that, and were being reported as unread notes.
                    let vml = match beside("Drawing") {
                        Some(part) => archive.try_read_part(&part)?,
                        None => None,

                    };
                    sheet.drawing_id_blocks = vml.as_deref().map(parse_drawing_id_blocks).unwrap_or_default();
                    let ole_shadows = vml.as_deref().map(parse_vml_ole_shadows).unwrap_or_default();

                    let notes = match beside("/comments") {
                        Some(part) => archive.try_read_part(&part)?,
                        None => None,
                    };
                    let threaded_comments = match beside("/threadedComment") {
                        Some(part) => archive.try_read_part(&part)?,
                        None => None,
                    };
                    had_notes = notes.is_some();
                    if let (Some(notes), Some(vml)) = (&notes, &vml) {
                        sheet.comments = parse_comments_with_theme(notes, vml, &theme);
                    }
                    if let Some(threaded_comments) = &threaded_comments {
                        sheet.threaded_comments = parse_threaded_comments(threaded_comments);
                        for comment in &mut sheet.threaded_comments {
                            if let Some(person) = persons.get(&comment.person_id) {
                                comment.person_name = person.display_name.clone();
                                comment.person_provider_id = person.provider_id.clone();
                                comment.person_user_id = person.user_id.clone();
                            }
                        }
                    }
                    if let Some(vml) = &vml {
                        sheet.form_controls = parse_form_controls(vml);
                    }
                    has_controls = vml.as_deref().is_some_and(|vml| {
                        [
                            "Button", "Checkbox", "Radio", "List", "Drop", "Spin", "Scroll",
                            "GBox", "Label", "EditBox", "Dialog",
                        ]
                        .iter()
                        .any(|kind| vml.contains(&format!("ObjectType=\"{kind}\"")))
                    });
                    // ListObjects follows the worksheet's tableParts order,
                    // independent of relationship IDs and hash-map iteration.
                    let table_ids=match parse_table_part_references(&sheet_xml)? {
                        Some(ids)=>ids,
                        None=>{
                            // Retain support for older relation-only inputs.
                            let mut held=rels.iter().filter(|(_,rel)|rel.rel_type.ends_with("/table")).collect::<Vec<_>>();
                            held.sort_by(|(a,left),(b,right)|left.target.cmp(&right.target).then_with(||a.cmp(b)));
                            held.into_iter().map(|(id,_)|id.clone()).collect()
                        }
                    };
                    for id in table_ids {
                        let Some(rel)=rels.get(&id).filter(|rel|rel.rel_type.ends_with("/table")) else {continue};
                        let part=part_beside(&sheet_path,&rel.target);
                        if let Some(xml)=archive.try_read_part(&part)? {
                            if let Some(mut table)=parse_table_xml(&xml,&theme,&stylesheet.dxf_borders) {
                                table.origin_part=Some(part);sheet.tables.push(table);
                            }
                        }
                    }
                    for rel in rels.values() {
                        // What is drawn over the grid lives in a part of its
                        // own, and each picture inside names its bytes through
                        // that part's own relationships.
                        if rel.rel_type.ends_with("/drawing") {
                            let part = part_beside(&sheet_path, &rel.target);
                            let Some(drawing_xml) = archive.try_read_part(&part)? else {
                                continue;
                            };
                            // Excel leaves the part behind when the last shape
                            // on a sheet is deleted: a bare `<xdr:wsDr/>` with
                            // nothing but its namespaces. Three of the four
                            // hundred workbooks swept carry one, and reporting
                            // them as unread drawings was reporting an absence.
                            if drawing_xml.contains("Anchor") {
                                anchored = true;
                            }
                            let inside = part
                                .rsplit_once('/')
                                .map(|(dir, file)| format!("{dir}/_rels/{file}.rels"))
                                .unwrap_or_default();
                            let media = match archive.try_read_part(&inside)? {
                                Some(xml) => parse_relationships(&xml)?,
                                None => Default::default(),
                            };
                            for (mut drawn, image_relationships) in parse_drawing_xml(&drawing_xml, &theme) {
                                if let Some(metadata) = image_relationships.shape_id.and_then(|shape_id| ole_prog_ids.get(&shape_id)) {
                                    drawn.ole_prog_id = Some(metadata.prog_id.clone());
                                    drawn.ole_placement = Some(metadata.placement);
                                    drawn.ole_type = Some(metadata.ole_type);
                                    drawn.ole_display_as_icon = Some(metadata.display_as_icon);
                                    drawn.ole_auto_load = metadata.auto_load;
                                    drawn.ole_enabled = metadata.enabled;
                                    drawn.ole_locked = metadata.locked;
                                    drawn.ole_print_object = metadata.print_object;
                                    if let Some(metadata) = image_relationships.shape_id.and_then(|shape_id| ole_shadows.get(&shape_id)) {
                                        drawn.ole_shadow = metadata.shadow;
                                        drawn.ole_border_color = metadata.border_color.clone();
                                        drawn.ole_border_weight = metadata.border_weight;
                                        drawn.ole_interior_color = metadata.interior_color.clone();
                                    }
                                    drawn.ole_on_action = image_relationships.shape_id.and_then(|shape_id| ole_prog_ids.get(&shape_id).and_then(|metadata| metadata.on_action.clone()));
                                }
                                let beside = image_relationships.raster
                                    .and_then(|named| media.get(&named))
                                    .map(|rel| part_beside(&part, &rel.target));
                                let vector_part = image_relationships.vector
                                    .and_then(|named| media.get(&named))
                                    .map(|rel| part_beside(&part, &rel.target));
                                match (&mut drawn.kind, beside) {
                                    (
                                        crate::ir::DrawingKind::Picture { bytes, .. },
                                        Some(image),
                                    ) => {
                                        if let Some(found) = archive.try_read_bytes(&image)? {
                                            *bytes = found;
                                        }
                                    }
                                    (crate::ir::DrawingKind::Chart(chart), Some(graph)) => {
                                        match archive
                                            .try_read_part(&graph)?
                                            .as_deref()
                                            .and_then(|xml| {
                                                crate::chart::parse_chart_xml(xml, &theme)
                                            }) {
                                            Some(mut read) => {
                                                read.origin_id=chart_sources.binary_search(&graph).ok().map(|index|index as u32);
                                                *chart = read;
                                            },
                                            // A chart nothing can draw is left
                                            // out rather than half-drawn.
                                            None => continue,
                                        }
                                        // The boxes a chart is annotated with
                                        // hang from a part of its own.
                                        let beside = graph
                                            .rsplit_once('/')
                                            .map(|(dir, file)| {
                                                format!("{dir}/_rels/{file}.rels")
                                            })
                                            .unwrap_or_default();
                                        if let Some(xml) = archive.try_read_part(&beside)? {
                                            for rel in parse_relationships(&xml)?.values() {
                                                if !rel.rel_type.ends_with("/chartUserShapes") {
                                                    continue;
                                                }
                                                let held = part_beside(&graph, &rel.target);
                                                if let Some(body) =
                                                    archive.try_read_part(&held)?
                                                {
                                                    chart.shapes.extend(
                                                        parse_drawing_xml(&body, &theme)
                                                            .into_iter()
                                                            .map(|(drawn, _)| drawn)
                                                            .filter(|drawn| {
                                                                drawn.frame.is_some()
                                                            }),
                                                    );
                                                }
                                            }
                                        }
                                    }
                                    _ => {}
                                }
                                if let Some(vector_part) = vector_part {
                                    if let Some(svg) = archive.try_read_bytes(&vector_part)? {
                                        if let (Some(size), crate::ir::DrawingKind::Picture { format, .. }) =
                                            (svg_intrinsic_size_points(&svg), &mut drawn.kind)
                                        {
                                            format.original_size_points = Some(size);
                                        }
                                    }
                                }
                                sheet.drawings.push(drawn);
                            }
                        }
                    }
                }
                // A sheet names its drawing, its tables and its notes in its own
                // XML, and the walk past those names was recorded before the
                // parts they point at had been read. They have now been read,
                // so anything that did arrive stops being a gap: a report of
                // what cannot be shown is worth less than nothing when it lists
                // what can.
                let hyperlink_rels = match archive.try_read_part(&rels_path)? {
                    Some(xml) => parse_relationships(&xml)?,
                    None => Default::default(),
                };
                sheet.hyperlinks = parse_hyperlinks(&sheet_xml, &hyperlink_rels);
                let drew = !sheet.drawings.is_empty();
                let tabled = !sheet.tables.is_empty();
                let noted = !sheet.comments.is_empty();
                let linked = !sheet.hyperlinks.is_empty();
                sheet.unsupported_elements.retain(|feature| match feature.as_str() {
                    "Drawings" => anchored && !drew,
                    "Tables" => !tabled,
                    // A legacy drawing is only a note when notes came with it.
                    "Comments" => had_notes && !noted,
                    "Hyperlinks" => !linked,
                    "Data validation" => sheet.validations.is_empty(),
                    "Conditional formatting" => sheet.conditional_rules.is_empty(),
                    _ => true,
                });
                // Only what could not be taken is reported: a sheet whose
                // controls all reached the IR has nothing outstanding.
                if has_controls && sheet.form_controls.is_empty() {
                    sheet.unsupported_elements.push("Form controls".to_string());
                }
                sheets.push(sheet);
            }
            None => {
                log::warn!("Sheet file '{}' not found in archive, skipping", sheet_path);
            }
        }
    }

    let (defined_names, defined_name_attributes) = parse_defined_names(&workbook_xml);
    Ok(Workbook {
        sheets,
        copied_sheet_origins: Default::default(),
        copied_sheet_code_names: Default::default(),
        theme_colors: theme.colours.clone(),
        theme_fonts: crate::ir::ThemeFonts {
            major: theme.major_face.clone(), minor: theme.minor_face.clone(),
            major_latin: theme.major_latin_face.clone(), minor_latin: theme.minor_latin_face.clone(),
        },
        indexed_colors: theme.indexed.clone(),
        recalculation,
        calculation,
        default_style: resolve_cell_style(0, &stylesheet),
        source_named_style_xf_count: stylesheet.cell_style_xfs.len() as u32,
        named_styles: {
            let records=if stylesheet.named_styles.is_empty() && !stylesheet.cell_style_xfs.is_empty() {
                vec![("Normal".to_string(),0,Some(0))]
            } else {stylesheet.named_styles.clone()};
            records.into_iter().filter_map(|(name,index,builtin_id)| {
                let xf=stylesheet.cell_style_xfs.get(index)?;
                Some(crate::ir::NamedCellStyle {name,builtin_id,number_format_id:xf.num_fmt_id,source_id:index as u32,style:resolve_xf_style(xf,&stylesheet)})
            }).collect()
        },
        defined_names,
        defined_name_attributes,
        compatibility_names: Default::default(),
        external_books,
    })
}

#[cfg(test)]
pub(crate) fn native_external_link_corpus() -> Vec<(Vec<u8>,serde_json::Value)> {
        use std::io::Read;
        // Self-authored complete Excel conformance corpus: CC BY-SA 4.0.
        const PACKED: &str = concat!(
            "504b03041400000008006d996d5d96375a7ebc1d00007a280000110000006f72646572302d6d6f6465302e786c7378ed",
            "7a65541cdbb66e37eeee103c48700f10209086600182bbbbbbd3b884e0eeee1624b83b8404d7e0c1a1710b7ec9dee7be",
            "bd3727778cf77edd1fef7c3daabaaac6986bf4eafae69cdf5c73c94941c3e000e00008000080024007bd6a09060200f9",
            "b000003600014ae38d8db5a391b5a38ea29bad918316b3ab9525790e0c145d16000af01ffc7f8db2b8111928368cbe43",
            "9ab6ee02951a8d23dd30bd551f7776494ac1d756f075d9f3dce97e4717c378d43d6596281e7305de9b37e0f4f461d3ea",
            "f27775be3edfe762503f02873a240b3f6b25400ceed3966b50722b25094cb20df95f7313eae8846ff2a327bdffdc1ba3",
            "c4760d3ffc5e590e99310d5e8a1f19df1cd38c286101b9114b6aa5debff359e2b058255b64381c7124a91ec5348db1ac",
            "86fa522c0ceb40f5cebdee9a12fb6aeb983c887283b814dcb80e1fec47394158f6b99fb84bb7652a7bee6a7e05e352ac",
            "18515d981b9ba22d3485f5e7f4110bbe72e67079e82925539c75d528565cec5e0abb5c163a398e32a08a5b94e8a3bae9",
            "4db7579fd033b18c20bd11f26f7731b6f0f0a3e8c8ca49f8562a03b2265ed8e9e62eaf455f88acec78f08ac6ecb0f634",
            "70f878cd8860eddb578000e7a4d002334b40995afe9779acf6eaef172fd202720f78bddda56b565ec96b201e79fb64b6",
            "8d81b63202c841a8fdc363175091119331cc6d760a73cabee6c79bc58311ed8fcef9f0000d90fb87dbd62ab1529d3d5e",
            "493f7a25f2a3dbead81b593ab030ff3affc75dff835f288b7927dbcd8a1274287e1b3c7039b518852122e5222cde4ca1",
            "08f6915840af0dcfa5265e6ca725c15024f2877eebdd0abefd32bff82a7c7780e2e20767d9472c1c7eaa3a97f2a0c46b",
            "3b27f81734f609627665ab3d98ce53e75317c96f25dde89c5b235f943614654b0d2e53894d1c83aa304698558810b1f3",
            "4982f004927dd5fbdea92de9ffac004e8928ec970755d6f364edc68ea47acf2414eca00a85f04a43c69371071cf44de6",
            "2a3c072c29ceaf8894144c79cd2af533ecf928dd740c2d3b4e41a9edd8779e4c326903f2949c583fd947def63fcf3af7",
            "56ffcc793c94eed6c383d85d3a7034e8f791ac68e9b74e029d378db0050d00ac230200e88f4f5c2d595c6cec2df46d6c",
            "2c7ee5b532652b9bdc9718f7e0f13bf81e1716eea4d77213b89f9a24f74bcc6bf6d371cc36fc266c61bf124ee4955cb4",
            "1f26059c3d2f55a8f078afaf3ff2f3b8e515c7e6fa694145d85be666794e8e58cfae82ce544d1c2307218ef6618ef554",
            "b14255854a494e4f63c1e3e15bcfeb1b1b5a0b0ab28300575640d68ef1a66431776519b3ce718a53266888388a31b02c",
            "d9afb8a5be49d95ed9916158494b8dbb0fbbd4d8c8c01a13ca0d25489579288d96f2d062a98829641d5741175cff62de",
            "1a547a6ab0624e43768db791630dfc51512679cee99322960e6f72191e05a980aae7e9e1e2c66ac3fe8989e22f900193",
            "e39fbbcb9e98832d758e98555e3e067f8f6d8979420393da4d40120c621a24934aef6e55b23f112ea5f1281b4426e86c",
            "0f13f3518756814258821f5b68b5aa7373af666400d7e457747ebf402e752a783151a45a389ea3d892944c2add9bd2ae",
            "16e5457c6e042b0fdf5c67a25ee8c1a72399bb2235ad628e06225d352908b782372847ec4160e835eac68914cdf509e6",
            "6f07c73b403bb76075d8ae853e4c24c6bcd103b677a810a89156361fc414c78f0d991bf775feb0803f60595778706944",
            "98e7d60b651a161d2ee94d9bd917b1943c73ed1f4d71e2fa42184703254403c36f781cda06a954be87f655abc8d51756",
            "ad2bb0796847c77425f78fce4b94243a7a60e40d0669d5bc15278eca9d553a275cef7db8801832d364af65233a2980ba",
            "07c4ee279f8734b5ed0554ed2b8cdca0690f4844214f974e4520113918cb0fdcf86efb598921760aebe0b16c09ad2dc2",
            "ef6c35ebc579cd64ec16b8289c385d8a10097071d71b29849c7f05f5c0d9642cc9ccf81fd5314e220d70f6e228d65978",
            "1f9d8061f1b067266eb25e7e36e238259465972cec79437fce9a231197358ae815c72b03e9470c1b34ac1c70371a1fd5",
            "eefcbc6ac6f94642f3a3f626c97583752b5f530fa761d64894092c61358ff55aaaefdc14ae8e3d223e10bec8bf36f1db",
            "45e59549aa8f3fef640c18855606b1ae509d386e849feb956fdff39f12746802aff2110a233e1609058214fac5d8de41",
            "26cfc7a9c335d34dbd46601eded63f2ba1276061b5da0895e04a64f4f6b140d78c54858f5b108383785959ae5f633a77",
            "8052a251deb3f7d285d2e0344a6c36904afc98fca61d3675c02234bbe76931fb2d79371b7d1b8ff52bb48071f3d4dd0f",
            "fbe64aa99fa9bba81558cf2e170fbc155ae00e19c079f6ce0e79fab71bfa60bbbc0337be07e7d467a524da2009d716b0",
            "ab13ee3d97ec0349642598486e4fab0b5ba47056e20dccb918e6393bfa15a173f41b01a2664e4944d2efa309e9537cde",
            "b7aa6d56e6e161613aed1190a694e1763874e74373b1fe2919983b43c637dbad065efb6332de80dfc584ba220249fcc7",
            "6831f71817880008c0c798f067eefc7b64f8ef3c0a7ccca3c0ffdd28febf8fb25861a91e568cc043bf078aaf8b055a4e",
            "98738ee8a1952b12336418e366c42934cfb19d6e0e5171e79919cfc7d73cde2c79b2ba35df1573419005cc4c600d2590",
            "50caed03ebe27d22d786be41b3228de38dbffc08e54e3dbb35a3d2ccc3c9def356712f4255ba8a3a0b2ddb633d4d31bc",
            "3cafc7c9742dcfb12d9e4252e758a6444c7980c7ac69039c01bc52a3e1c781bacd6e80758634375c3d58aa24a8afd628",
            "64327db7d583458f4636c83ac38503fadf72d550fb41b57fad548b05ca025be33795e68d45929b7bb6cca319356221f3",
            "5741e63dac7ae178c77ad508355ad86212b919e6c5fb301cea52f54d3ad87ee63b55aee882d3be34511c1e889bd64a49",
            "a3a9b36dc797e3590fd3a735229b4280dfd12a4088c48ce55172e13c96510480ff936a1c4c8d8c1c1d58fef862fbc5ac",
            "b4b8f9d079568ca043642f78a5cb2acf2f9d7dc731cfd59ae24476760ace44a2fb3f8458114af0569d2df2db131d72eb",
            "5130ceeaba19b75de57b432007722a86f0dfc5a7a4cd56aee25e532952a44ec5ddffbc674d92d367b441d22e862b3843",
            "085ada385c36f2c8e73146d545c422e2c8b7b0988f63a39a9673582bc6230a8b939c59f821ff29c98ceb2887a18401ef",
            "fa3e644a93fb0c247ece84a29420dfa1f9c10a355c4098537121075c9350ce3b616e875cb1e95e63052f32c54499501d",
            "73312b4c5f2c7add601eb7398474e985a0cf3c20bc86fbe359fabce3e127be29795323453301ea9ad8cdbd7d5a463dcd",
            "b3bcbdbbe50044fd023dda74df91a223ab183fbd11a44cbd3162bd5cd2cd08a9420d35e99d0957bcebd85e3a12ea2a21",
            "929e3157ee864c94586192befb971262d164610fa3af18a8f920cea29c710a7cc8e3a220b50ca7d35385aea516579a04",
            "967661d82ea80b043abbb54645f3352cb2e3f843c939a5fd373c193a9cf459ebf35fc04a3791e42a161fa33ef090e658",
            "9b78aa936ce9c2c890c4b43424b26d256c252b7bf86deaa94cd85fd2c3ec5804bad168eea238601ad9a6362f8ccca2d5",
            "7ef0aeba49ac5a7998a5ec08657fc0bdbd8729bf63edb6ddc1f8c974fe0165e32c65b437954bcffa137dcb288eb3a4bf",
            "4a53562b8f4be7429b94e3d8f3dda680442d99baf4e1a60d869f5fcf25b98314e59cc161f463b4239ee899e5e6e952c3",
            "2df9d5ca535b53440f823e85871e74398a5f32f0ad39af1fd4eef17f47b27da5c9815e78000097f257adfe07c91c4d8d",
            "ac8cfe3cffc12f88faa2057523a637faba37c69d20ffa2c0b2b149e495e672c53ed7de1cea25dd886941ad723d918481",
            "0702bc1e2afc251682a68836d6494230867d943f094594505fa594fd1595bcfca574fdab94f973b06da2ee4dbec669e9",
            "f8a75249a529f9e5877576e3ed340409b7f1733f8fb66071f60d8e14c6a6a59ac186f18b28c597c97a7dd8e3e7ca20f7",
            "cc573f77640fd60470a393c8a364d7490fdca3e577b46b5f14b2866d0864c4c080531d300d60ebf5dfa7237a02dd0d94",
            "51bddad5c2d9debd8a5b2198144c1ed989cbda2b65a9f8746fe775ee71d8e9ee87e9af86f5fcbc23ee03767986aa3c1c",
            "3f3f32a748c5ba0fa26384da8c7fa3f1adbe08737c3041021278dbbdedeab404fd40e5589bb967126d2e6661f676c3be",
            "b103c2c8e910dabb632cef8d53b7d729d60e64ad0d27ae3d7f10cb04aeece06470aec94f1372a680f9bf563d9773de19",
            "5d3483e522d619d799ff41e298725769bdc39325eded7543d89f3e56a9b78a4ad7cd53772420047131c75bf261f77440",
            "d6213b575f1de82296eb31b5af0bb7aefbbe6f1cb1f47a336d9f7f35dd78c206d5c1d5d7cf74169498a13f8c5dcb30ad",
            "7febf7910dc019a5278a9da74c9fbc460fb26751d2cca188a7ef6640cf7fb98a1574c78a407eb60b6eb8e2c1ebf78049",
            "cce362d1a5ac0fe224a32ef42b28d913802ad7463cd99db7f9a40376b83bb1995151996179f043dba6944f7e21654d46",
            "fbb17ac1c5fb76ca1b7c897c76d87ebfdafc09bcdf37c47e7f9ed9ee7977e90dce7457350fa4d7bedbdc9b59f6bcdbad",
            "6fbf9dc96cb3373d44f4540f923595a6e6a3c4dfa68f5c0586d35bdeaf2b2ec8e863e9994bac2038f22b462a1a022395",
            "931a6dbdf4538f1a7d18d7a00435db2a1b3240fac545069bfdef2d81e8bc265ecc6970ae655b2fd45caba5f10e067bab",
            "c33c428f1b2d0e80328cf09109f4792abbf2bc2121a2a749aba055588910916b1eb058c8da8e60ee64f51e453a12f395",
            "66f608ddaec12ee7604086baaf22c2fe04c16db53d268c8a461daf67c4755a6386bbf8b0153523f2579275da003563cf",
            "9044a47ee6786c929aea4ef4ca504d6a59657df5bef03c2859168435534e2af926c30c04d7a598093203714b997ab26a",
            "254baac01c13e674f8a45376fb811138cc08055a93149011a4f004450427d7a042f3e88df5270c2f7470c7a08c8b5061",
            "8ffd652a762787b8ad91a4a8cfb64e42d7ea5070c0fc9e5def665243b5b46c8fb2f33ee26e58c4bdbd5ee00590024567",
            "8590627bae100dce2b439d9b455c2894bc1482ada22bd0e55cb67b1320ee27ad53934e895ef459e42280f1279a607029",
            "f10793c058828f5a29911d7bcadb1948aa61115b8f558589d10a66201ec68f5eb3385a2cc7383f8fcf527b7601e5dff3",
            "0ff3f9b555ba4e96e7cbd8541359ac9b99858316d4e1b06e34e68dfd92b9843159b1a5c77c1916fd4236ee4540334a54",
            "758e5fbeed8b6a340089a84a973f989d146c6e212077b2809396b6729d04ac7b0f270e215f186e789ad2b41a1e3a978c",
            "431207d97cedabd31b0ed4539ba7fb6eda6dac9ccb379a83362bb21b2d4c353fe992e567d3bb1398da849693d61c631f",
            "2e5c262bf4c94c8646439bd581a95ba6bff6b19a5ee47ded5b3726630d5f91174297d65dc5832226a46eb826541a9308",
            "50982543facabc29b259828f95afd955174dc8ed51b21e645d73f5390cfd8a71fe80f3460a6c2cf24189ed8047778d9e",
            "38b8fc457c766ab9a7dca986028f25be71a9d3bbd7434b412033f654d037e19c5a524f68c456870b3d554712c6efea77",
            "df95340a00e6c49f3a3c78fceacf2b668a4cc274dfa2b99f255f286ceed6bad43a081615ba10c6c65301eaf0ea0516c7",
            "345d0433c04d8c934b5d1378886b718c6596587d2a059db1982da85b636a4b2e0f6d3ecde59ff803a598dcbfe3ec4dc9",
            "8d9ad60a057d72c26a8c2656951fc66114c6afec7ba3ea489af2e3f3de37bd6502e6351cbfb7bb2e2c5ff16c011c4a94",
            "1ac78731eaf21453184374350d837cc9defa2521afc3fab464851e684854f563c4c52aa0a52efb259927910ffac5d05a",
            "ac3d92bf6e19967c01b3849b4a99d447fda8315cbd9d0a69e757953a79de70439da295a5508723cb16a09a8b6b7a385c",
            "14b0e64fee701ed6e5364b138b7bec92cbbe31b7ccb7fee284cdd82ce034467e78937001270ef599437d95f00eb3b731",
            "eb761da22c7e53579bcb9c9743f97402c3ae41215d8ab2fa98d144fa3b21ed7896aa172545bd2f8616791cac4dd221e4",
            "72e9463fee2cbc3fa8b3258bbee3581aeee633dbcae5ae459fc58c2b9175085ee416ed32fbe0c27447eec261ee14a8ca",
            "934073f46dd5edca3235b3cd679d2d40a1e7f9904cb31a944b41ad7464e0f13e7d66f989144d0d28f64a93275fed85b7",
            "e466c83073b3f6f62ec969f75a0bb12e62d6bb1164064baa6dc885a5b016c47eccb2d33aa086fb42acdd1cd9867ababf",
            "8c58afa6763ba79c798676b7498587c322d848d66abcb20a47d34a0333e20136689a97ffeb377b96f564a97897efb681",
            "097ecead9937e1da5e18dc70b8ed8b044abbb54d3a663ceeb5929fc68a6073c556d06f56b75ecb8ada8d0eb48109bacd",
            "303436965f271d971f92ff38f79d69a58f2c5df0dc92e35ae2591686c9dee30ff9703fd009fd16aadc1667b25bfb6511",
            "4a90708d3c7af037c966a53d25383d07828b2c539efed6110147eb0e61ef972491c47af9bc00cdc543ae17ad602b39b3",
            "bab777b62461bb6fc83adcae1a03ea09d21d1c343af257cf57da7f2b2da386d4e31b1fa5a5281c0080fa67d6777074b3",
            "3472f895eff314bf99cfbfc668c3bf461b5a8a8fc719aa0dcf93ebc78491a7a30ed49e19ffa25d2d645c2bb391a01de3",
            "427786eabe973f3a3ac3b6cecbc770f596269f6950949f9fff7a33f5ca9847b038ff35cf5745cd5218f8b78262d21b4d",
            "2be6092b2e2ac93c1d45d58ab88aa6bd582dfab882123a6d210ccd5810fbe12451407609cc48b9561842f1096fec67be",
            "71e48295bc493179640d0eaaa92e5675c5035a7af1ed1b2f4a7a95065182c15befd7d45c33438189168d5909b28d2bce",
            "53ba7ad4a76357d2666f0aa18aa68523378bd0ae6940b57334662d9127e219c45a3be3a440e631a341878ba4832814c4",
            "6f97e23eb79c79e4946057390fab8c9049aaa0455b5a32993579068d6c4d96785836df228bce6df51df4170e6f87e7c4",
            "f2edce3117e580198058cee1da5138b52834be19579cf8b270d786769de57bd5e57bcde5c5729bf6e6a3184966e57dbb",
            "6d929efc2546dc34070a679048223689194aa89ab1221cf1e7aa4fcac80345ba0dea49340764631c6f5484099cc0e4ef",
            "ba3833cc153a580f352a9df2ab74f36173824636ad6bdd81cbc5757830b26b7e13c9ca8a0df2ea51e9c08d75e910299e",
            "a3c321684813c89c9ec36ce4234cf794ddbbd914b9813ecbe3efe1dd8739b5b32341a24e6e21e821cfbec3e7b9ec44c7",
            "beec445f6e13bcbd86eaaef67d3df11c3fa1540803ee6bce527d870b92aacb7cca26a4fa352d92f2787ea54a7337cce1",
            "3e4fd98a14e141ced209b11571c8e33d72a1e2a80bafaa15cb0b0216d5b1f123186f3ee2762299e1ea51e29f6bbe33fc",
            "c6d13c617dd82cbbe911b4b196dbc1f73b5bc3e9de7de13619459c41cd9c25302dc8c60a90829ea15696cef47752edc9",
            "efb3163704516eeef7e88e5a21b3e82dda99682dc20f2f4faae91068ef967e327ec941bbdcfec0e22feb97e565407265",
            "6ad868e8575aaeac11d8a48d8158e74def67194f7e73297487fa3b8ef3f45c0a553c5e1d3f1e38ffe2b8a99ebd91a182",
            "a3bd99b5c91f54d71c7e87d4c58a01bbd2baf2ac0bc2a50402c5e0606a364e42f7387ff63596e757f0b6b601816cafda",
            "dca047e222674b882f48b0d492de21cb892b95406b3192a0370f22e17c2cfc5eff61b40b9b50928794d0f8fd55e2ac08",
            "d57879ed6649e7b65b0a5f9e276fa1a0fb6055bbe981fb19174f089686004b687e842046d10b546a3cbc2c38a41960fd",
            "320405dea467ff0e5123b797f02e08ddfa47f3c527b765a141813bf1adc103d8dfcd6f75f799aa171000a081060028ff",
            "9c9f91aba391bdb59ea5b499b585c33feefe2c14633a7f355b820e69dade4cf03bce11e4123b7f084ac208c73f074aba",
            "f804d5d5982133db1c7ea60b7525d7aee398189ef17ed8711dd7e8ae5bf11753391dab5c4149eb526e36fab6ac4d06c7",
            "931648a7c580ca5774850695c999886fb04de54afb2dd6d9420c03dfbc2db4aa3400c11c563988c279c5d4c477e98b78",
            "b14ee8a85b94952a10640152c399e992dcff9921a7e684d623d9bee4081d5817fb7ea476cc2e512d4d4de95b025a6b1a",
            "31d67c6485ea72aafd596f61d720ac32432410df4efcb5efc74065db9253e50552fa6fb30f555fa2c94f51aef0bef87d",
            "e6826739c4146a2dbbe51583f4cabcb7626c68a89ff0a750180d8c3f8032691bb7fa7657f6eed2702c9ed3cd1e73a905",
            "31f8c7abacb9f5ef2eb283a21c78dece1b988252c8a0973d960cefbee9d8aee1e64deb10478ef64abf32cbf842a7430e",
            "76ce0bdbe0e018d72611ef3f3f7b43ba4cc81b2f02db97f793f9f97c50c6b4aaa58457c98408ed4e6a36e229c7307e61",
            "f4fa659af1eb66a14c97d3df2e239f809d10e21faf8a1f0f8c3f5f9e819ea5c11b533d33eb5fafea558430d2e3ab8239",
            "f4f34670703199f10fa19ef3df463e0352f0a7bc665763173ebbecfb98189493dd747e3b2cd15a24c968a94b2426d7d3",
            "c6189d097750bf77802099b2d69bdf65b8f206aa66d1fab3cb60a6ffc7b226dfded7a6560214a1e871e92846cf0fbee5",
            "4e42d4fbaf416a0c8da250ef87102fd75c3b1660756b55a63a8ca0706fe0e576124db9eedf427e5b391214849f843d4e",
            "6eec318f6002108086360672f636b60e2c0636f646ffeaf1fe67b10bf047d344aa9b95a4efd06f397b7d297f31497ce7",
            "4511248ffce5fb96c0f543c63a33899c4711954d8eb0abf42ac4f50816e76c1d199c6c796a91e4ec207346bd14f34a16",
            "fcea81f6d8c4796de03b90a23b40254ac3360a368d8671fd1c49bd5869dfcc5042f425d5fbae6cbf14f2316956ded589",
            "c9dc0f3ecfa3f3c81511c6d9a999055294470be48b83abf25231b0d558d1ac7bc2413da0e4f2a2b5b8512679784dc734",
            "7a127fca303a3f8eacf48810d31043c700e97ad13868825e0c61144d09089ee6ed8612b9fed55ea026a1435fb6798171",
            "81de821d772eb4ddfa52fceac0ac292752b83c69f7bacbf1f71f5df68225cf4eca1890e88e08cf8b00395898dbd3760f",
            "b70597ee935e025eed9e13fb0c8bbc33fbb48b6d23d3fbc88b6e63e9168bbc13e5d8a52e2ea3deae90fdc31de5bdbb8b",
            "b21bcfed353655654f509e3aa7cd674ca854ab7be807da89f5913d2214e057765857ff26c1f8a9f105cff250bf1c85b4",
            "6aad196729ae99e3b46a15e7e99a81e503934014b41f7a4742d84240df486fcfae15e8df5158d0753721fd919f84d0bf",
            "bcf06f14d6b3b5fd0f83ff425a4c95f5af15c64b8017bcaf87342faa5f4716018ecc47147afd0e19e704be39107d0cf9",
            "88df229bf6093761cdbad52d25af440855617ad39dcbc3ce4c3f55ef54419d1a341be1211237e9062374a197d942b5ee",
            "e547b2c24446a922cde23217b94a266ae7e7ed69216a67404a227142d106c54a5cfd24db8c3324f37831437efb176416",
            "818a7dcf2f958c38353d5a2b3a2984864e91d188615444dc88d07ccb64d3387fb667e7cf30e9b8ab5a4203a58e99602c",
            "3e8bf0e5df9a399e0daa24792d9b56faa1bf13767418dd33b56570e80cce4fb230d293ecac96e7f65d2143b5832ed1fa",
            "2e9d9c94676a3bee23465663783ffa7d58f542eab9a91dd2749169dc9b2ef3feeb9c096eea30fb768fa543edcbc8919b",
            "fe446548064c61e71264fd28359b7bcb7e0135ea25a925b43b397bc47b7c38a6bc66f302ca7b7fcf57d48704ee07a426",
            "0733ef2728058f9dde160a797dc9b8e6a3ab7fbe96651d7e937ed0dda5c31070594bcad3eb994f19e7881d5c4b46f423",
            "889bc88a985d8bbe3ee0c46bd326d70b06fd4e2a71ede3fb7baec29b6767300fae1327bf553d73b1e3bd9c8f84cd7c8c",
            "ca4c80dfa8823f3b13ffa60dfe684fb4c42c5b3ccaa1fe65f81b32c92583713af59e1706d6e42f09de2c5fdb17cc70d7",
            "bc13c1aabd3bc4ef83c30e68bd7225b9bcba5fc12b6db388e5e0fb14362cc1f7835a98e95db1637aecdd41bc31f531d0",
            "11b17c2c291628a0c644468647362b12f2b6edfd346d748e974c0a06248d8d8518d2111926fa8c27deee59c0cba614b5",
            "a8b745d188c3855fef8ca2367039e6e278ed52e5a30c79668b6e85e615e3631ab84f26b3d178a8092e80e941cea567c7",
            "cf44b0e756cc51788b153f8ef17a64e07b90edd12dd85e2bf6c6d65a20c8598607d4cdea99204f30b80b025a7d0e8b02",
            "08e291980be8c90a031b4d19dbebc7677defbfe473c8c02acbe5079b2947d062b3119b681a1a60840c98d0452bd8c8c8",
            "506ee4b69ddf7daae251e8db6611aa436c5e7cef9e19711015a1d53a7582fcdf7f3f108a09f03f6d7e7a8adf6c857a3a",
            "c03fb761fc1df58f61ea6f9b329e1a3e6d4dff8505b87f6b543f35fe670febef3842fac3f87fe8683d1de769d3e22fbc",
            "45fd9f5b184f4779ba2afd17cad07fbb46fd7480a705ee5fc8c17f52ee3e357d5a37fc8520a2df57114f4778aaccff82",
            "31f1ff8d4e7f3ade53b1f817a89efdbb747c6afd4f9df6776c3e5aff9b6a7b6afecf1cf9771493ffcdfc5f19f3a9f5d3",
            "70f4173ca9fe1f83939c142cdc2f43d4c78fd9e38fb9a1fe75f75f504b03041400000008006d996d5d81f19222c01f00",
            "00e12a0000110000006f72646572302d6d6f6465312e786c7378ed7a65541cdbb66ee3ee1edc35b8844020b8bbbb05b7",
            "046b9cd0c103c1dd9d20c1ddddddb5f1e0040f810e97ec7ddfdd7b7372c678efd7fdf1ced7a3aaab6a8c3947ad5adf9a",
            "df5c6b2e653918585c003c001100005002cee2bd7c3f400100cfe101001c0022b49e98a3838bb9838b91bac75b736703",
            "56777b3b8a1c5868862c0034e03ff8ff1a3d71a336d01cb88137816ddd51f6ef725865193fb1da35339529df01ac9d5c",
            "41c934db92258e60813304c63911923d0b51bac6afde8d3496ecba5a747a22af8fed3a3d3d12cd8df3d596c650aef97d",
            "f6052b50732b64092db38d895ff33e5b5b432d20c64852a9ed05c573fc901a534950461a49479063e62260c2b27eb6bb",
            "82d288e9bc0e71b8c990b3a45e5808ad852746cf30a64ba0b318b7fdd14a8b641c65be75dcb75462dce5959b1f223ea4",
            "a50971da3f22712a96f7aa98213b45fc70168b7e3504866f6e02da284c205785229e29a80774ce07f935577c777288ce",
            "24e017d9b7e07a2ec2df585a04c3ff03206a8284c91c5bc8511e50e775c2f5d39aeb9cfb5217fef3786200dbf6de030e",
            "c2442f926692963dc140a4a53d7057bb399b3f9a0ce6db656b06d52037a1e90a6147dde91b6c7b275609c015164cafe4",
            "21a07f098d4495909f751f52b312f0e184dfcf7326765d50b505ae81c438d3c8d5b53f1633a06823cab830667d5d7ad2",
            "fd2af430d51eaf08674d3010f20d01007878800128ff63ecd66ab0535f3e5ec93f0e4d94c7b16be4646ee7ccc6faebfc",
            "9f31fb1ffc42698c8252373b6ad0a9f47df0c0cdec6a14a6a81c5044ba9952ddf7bdcc0a466d442e0df16a3b3d09a6fa",
            "b30f30527eadbef743cbab8211070394d79bdca51fb17105a8eb806541893fdeb92230d1392548be2bdde8c1729bbd9a",
            "bd4e9692f560706b8d642a6928ca961b04534b4e9f8957628eb36a3d43c2c92709c27f95ecafdba7a0b366fafd0bd4ac",
            "a8da715950453d5fd641ec78aadf7c42c13e9a7008bffcd15432de80b3a9e5e217ef013bcaabdb671a6a56fcd615a619",
            "4e2fa93c8cded8755c88a7b6e340bc9f2ba60da85271637fe71c97eaa7cdbaf2d3ade13e1b4ef7e8e143ea2e19f83608",
            "fa485eb406f5bb416204e39bf5f81040810c00603c3e71b763033a3ad99a3a3adafe12b7524d7bc7cd17983f7da720c1",
            "0ce0d6ca0bd99c293c8f690d1b2d8de13a20f6d4104dc1764a0f7f3f54b1cdc3dd2718285d4dfa2f5e88d87691e4c25c",
            "17233404997c1a9fcb92ac500c235dcf910e9f957065105c4d6b3fbc5498949ed795394c4b68b1308304d6fef83a0cf6",
            "6bd6763e13fc7aaf03f3be89f4508e31b5837dc96e9069c97713adc4c4bc53428216b53776921734d962c0495732c5ce",
            "e11611865fdb2681d8c34fe85dc9c0897f515aa2fdbc9e0214af0e0d71d3e13e7eaf69feb543435d108836e43e0be869",
            "6140b99dfa70a01b0fa7a5c443e062085a26deb422863222f085edee255104e5c7e61653e1b4eca0dc2265453cdfc779",
            "40e245f4c0018d5aa38e8955a1756a060e6d64c4bf151e185a356b446007f91145bcaff12c83c56cc1c8e1f77228f0f5",
            "a86490bdae88ff2299ba8b52a65e206bc06ccd1cdb12c5427e48867a9bd29a8c6980e326815482ef26afc4c77ae3d486",
            "a73b865fc3bc831b20d811e98a938adf37841dc0c1488bc313cfb2590b4b7a72ff4363f46a6132e64357adc0b5921622",
            "7c20dc0f68a30dd9eaf763cb7a4c89baf443956470d5e5cf1b8471d856c1bd7653dc1cf12e6ea35cba6305a53616de6a",
            "59b369393cf6cc712a0c0ab0be0db07a6fdc688c3c75991f443ab8d554382858f40e961f4ec560c9428c31f7cb709fdd",
            "b7879553aadcb8c059c3bcdb6e2e6e03b584b7b5d2b1ca7893bc1c08a34658cc26c899a87d2e2e2aea3d1ffa9cefbafd",
            "0341171aa7a240e38a8c88c92757e273df2a1731be6ee44e11adede00ba3b3930f33905573edbb2f821a0aa7d9c687d7",
            "14cf7866782325d22faeb4a57a601d0dd7d0e7219b66d1c3b42183fdbc49765f7f1af74e6313e2f4abdd59d17dc44b9d",
            "e21f612090ef752e6da76324cab74ea2e29347aa4a9709e8e2565421c43e4a5dc05020b8b90fa706f954c3e5400c4aee",
            "a09355cf7a685c5c6743344a3708b18ab87a3414e382fbe76e5a708804aa82daeaeb79ef4630822b148cf07648f5d8a1",
            "30b0e01132474879f520bb99f6d0100be5e083b98b0b245e1f73b4bef9b2a353656838f1c005f74d13b0e48bcad734d4",
            "6071ea654e9537fd169bee3269019fb0433d25e11ea40a4f983908c97cc3766290f9f1b8bd82a4e81ac8a6f0b757fa27",
            "016dfd6fc42c714d5118961865a215bc51179dedb77b3eefb8d37cc2a67f15a571c226b4402f25d99eaab22882a1729c",
            "4b8e709072627113fbf1b84791cc9149b8e1a3c60ab0f961f41b590bdb6272df7e7772d06462eafe6872ddfce0e5b299",
            "e797b36dba70faafcb4a2727747e0ab1c2c1d4e199f0973e5d66a6770bb274035554c48b8607d92a61d9b98b19a4fc6e",
            "499847a4d8865faa67fbcfed97fab206f6be6b58cc6369fcc473aff01a9fd6a503ca4d7c0ce20c687b35334f767cfea8",
            "a7bf8915878a5d88548f51c4021600780640847a8c157f6aeadf23c6ffd157a8477d85fadf8deefffb6889eb94ed61c7",
            "0c3a059d2aef9e14943585d9352148687fcb3b840970590a289019145f6d6bcd0e7643916bc565f34f96f9f9601633e8",
            "e067a3701e9151534d52a5808ca4bb2f6b6f8b9ed93f9f2a25cd80a887bcc5e2b4176d6ca46f74bc36198289113d454a",
            "abd28b8d80277ad34328135896157138d127bbe3cd2c8aea33fa929258857a78da65f8be111e976b745bd9e95bf0397e",
            "8c1958e26b7bad818675289ba26fad55371c4a50e1fb35e48b027b1cd7c0728eb72fc6a7ec94ea42bf77b0c6be2a95a1",
            "0b8c1e97b5123089f84c1be2ad1042fe955835c40444a8d61324db061b0be11ad9bcb29b0fa0a4154c20efec6fccfa2c",
            "507da96a79a21588d41a6af6a2cde8d926ed5d33cfb1188943da4a7014daa13bb9ddb5006e3be0770c13a8ccce687864",
            "4dc2a3221102fe478d9cadcccd5d9cd9fef8e3f845b2b45805dbad5fdf150582210ad4fb395018aa96d264e2a5cd30ec",
            "701a265827c132282d1b5b71e407ae145b09d417ed848fe0f1dc251d3d9cdd92a9dd86298d95fcca8d4e2dfb1a454dd6",
            "e0ebe54db97d45568fc717b9b56768fcaf3ee4461e063b96e4567acd03dc41b0c74e4bee60f370fec9f79bacdca4562b",
            "9fc33556ccac6c775fa2b490a07f4ea055e1360b9353600b1859eb66d7a085c39524c1b5f2f3d7b66a94827c894e6f9e",
            "8a6b40286c58256f7a4d384b694aafb73f0e13874d5b35af4b3c0c7c7eed1346af1f15db4bdfc9d79a40d406565755d5",
            "de92df11a3a9feb8bb7c9d6837d84028530f010786d28f74051f217c0f956da363ccf854ba8d93e08bc54b6c02715d3c",
            "a2184efbb968b7e9a227ffd1386164b7f34d2c2bc85c5710812d75ef96563677d39a0322375fa2ba1970f7252490a950",
            "89c3fc70d63803882475034eb9345c21aa3ee533ac949c29ac4a5d395eefe593932dfa9e7d253c7a2630d07466752ff1",
            "223367c07748c783b503a98620a66b50bb6e3d83ba7e8be77ebd2e61291abf98d25a31b4f57d0ee6fc9a0b6432a3816b",
            "8dd835e2d9a5f07a5a2ea6631e679fd107c3d3ef52650fbfcd3d5c02e7a1381f73705cb87fdbdb9c7ff4769ca3632a3b",
            "661f586909f8ed5e88507a0a5fbc2826ca6a96db7aa99d6fdaf0c5ae742940b22e96bbf87a151814e816ca2411ffba6b",
            "65f594ec8e5ffa84bc24e64884f8f301129a60267651e090796bc6043938f183806cf88eeb989c46770059e7c4dd25a4",
            "793beb72ce0ec022226a4fe762668fc365d103cfe42038fb3a225a743f4c4fff80872b023ea35e42635014dc6e8297c8",
            "738c58d916941d268a3c10b7a0a1bc8c9e93bd8d7f951a613712aa196afb166f7ed9800bb3d806b33226bfa61f5f5686",
            "7633d8b58cdb28bee1b4dfb807ce63c63dd01d43c97ee986209d89b186b3b8822c46a7f0d6e90eeb2321c1ab7da7072d",
            "1310d67445e25887a5fa775e99ce06cbc0f18a0d6213be5723a4899f809c991385bda64470d50ccfde6ced58b1a7585e",
            "88d4a05a7ef04888e84a2667a5f2fb1e5c20c33a6952cf9c939d34618dc91614972f23202c4c4b32f2d1d2f4f680388e",
            "82c538f23d365f06737e1deb1994d1ed97388319d6e5b3fb1c19743870fc75c94b6fd8c635ed80158f7ab399e3f0c80b",
            "3378c9f51eca41e22b4e4e078f142b570a2b32371e3460a62866bb66125d57be2207810a7c7da31224e05825f53b416a",
            "6b38dc433f3e9426ee7ef9fa0dbdd5513063104a7087e367a5370c85f301ac866548ce5e0a1bc905151acf5aa6e08e5b",
            "9b5076af0d822c0d82ec13bfb5c859ea01e7d08cba2dbfcc57f1e605e45c6d2560af7aec7a7db7d3d8d277da2d8bab2d",
            "7e59c6fbfd064ad0072c8615a87a5df33677f2617ceaf751e5586366a0f751d2f0a87e2de2fcc133172b737bf33fcf7f",
            "049423dd555b9a462c3f8c6d3f4c8890c0ea2bb08565e4ad3ef8cb31cfe122da0dc3b85541ad66fd3319332f44041334",
            "841b6c447d5143ecf384604ca7a80f249451c27d15724eb7d4aaaa37f2f58229cb57be6f138deff2f52e4aa6ca4b6435",
            "6655c10fdb9c167b6988321e535720afb66069ce1dae1496a6b5eac186a9eb28f517c9267d3853579ae29e9982dff795",
            "4eb65ee145275144296d939d7846abee1bd63215b287efbcca8881f54d75c63283ab37554947f286f234d344f369d789",
            "e050108c5b279c114a1edf8fcb3a2c61fb52fef39dcf95d769a72708eb830e36ed55475c284e5986b62abc80000ab7e8",
            "97edf7482e9f74e63f345adc9b8ab2c607132620fbee79b6dd5e14639c689d19b2f6cca02fc6ac2cdcef3835761cb170",
            "3b87f5ee5ba8fae1d61d764ab643b1d74610d75e3d486642adefe366706fa9ce1171a7f80a8c54d22abbed4fac5ac3f1",
            "101b4d192d6f92b8a4402a1cf6f9b2e4fd7cee88fad3272b4c36d018baf9eabebd123e02dae0afbde7f476463122bfd2",
            "dd18e82256eeb172aa8b70a85b3ab6f8b4f67a37ed586023dd62da11cdd9dd1f64b5209e98613a8653cb3c677a0ffac8",
            "01e08e3291c0c9d3644cde62147762d3d0cfa18c67ec66c6c87fb1811d046147a4b83cf06db8e5c3eff7824dcce36133",
            "a6aa0fe226a7290415141fbe822e33443a3f58762c37f275869c3bce6b69cdb33d80d0f7a8549399e41cc8e93f56ad00",
            "fdee67fd7c6f502e4fdb7f6e3497fb1ef70d73febcca6cf786dcf8f9667a6adb04321a42760fe7c1de9083faf6fbf9cc",
            "3627ab53246fdd20252b799a9754047b8c911b50118c763fb7d557144db14d6c64d6115d04d423d5df40456a2635bef5",
            "314dfdd6f89e650b5a48bfada22143dcf47391d96ebf8a1d1406bfa50f6b1abc7be957261df72a79fc93c1deaa70afb0",
            "b346db1328451684c804c63cad0355fe9010898ba40df10d389910d11f7cbe92215bfb42b933558794e9c8acb7fad9e3",
            "0c076607dc830119bafeea88c7d384f7554e58b05a7a75fcde9f7ea43566784a8fd9d3b0a08c906cd307e85878872422",
            "f7b3c6e390545775625484e9d328699aeaf645e4412bb1216e597153ab36bdc940745f8b99263793b653ac27afd2b0a3",
            "0eccb1644d4748bae0741a1887c7faa4466f99226e7e54788e2a8a9b6bf645ff9b984339a60f866fc7a02250b8b0c7e9",
            "2615a7934bfaadb9acc4fb3da384ae8de1e080e5c377bdbb490d55f24a3d9a6ec74807e19f7e3a99045e4351a21aad13",
            "51ee2d16a2c3fb64e8f2b2490b8751941ce16819bfea722b3db80b9006c91b55a7536114d5885e07b07c47170a2e210e",
            "b50c8c25fc689012d971a8b99781ac1dfee9ab775781a5f93a56203ee666af751c3db64b1cc8ab46eef05d40d952fe69",
            "be80a156d73978b994433b91cda19955246845171efb4e6fd90294cc2382c58e233fe9cfbc0a0ad9f9292a3eaf415de7",
            "32347a2ca1d700f58cba041c6a7d5eb0fb1511a593cd3769ed6baeeb2b87ded3e9d3a321e63bbea634838687ce358b90",
            "c4410e7fa7aaf48613ddd4e6b9bebb76477bb7b29de6a0dd2fd98db656fae5c6e4f9d98c9e84568e616564d56738a72b",
            "37c96a7d8a3361d130d675be342d73237dec56d779237ddb16e4ec11ebaac218f2c61bf8d0c444340d3f8834266502d4",
            "16c89147587745778b09b0f3f5bbeaa28978bd8ab7831caa6f6bc2316e59964fb8efe47c2d444335384ef88cb7188983",
            "cb98e2b353cbbc952ff4d4f8ec082c4a5c155e0faf05895b73a68a8f8ae4d49279c320b53a5f9b68bb90b02ce9429634",
            "f40a0036c4e51d5e7ca0faab2ff34596e1c652e89e97c9d76abb07b5c05a67a1a24220516c3c35a00ebffed5eaa43e50",
            "28c3b7896566ad6b1a1f692b8ea5d40ebb4faba03316ab05edeba4ce1af0a1ed7d7359b940a0dc73cf25dcc359e509ab",
            "5ae1a07257ecc668626dd5315c1611828a3e316d17b294cd9ac351133021eb162e48ea00c83682ff16c0a541a577761a",
            "a3ab4a398b39cc50dd30f832d9cfb438e475789f8192f0031d89b6698cb4e41718b99b7e59d61994937e49f41607afe4",
            "91af6f8a877cd922ace42ceba336abdf6cdccf86b40b68cb9dd336dcd1a41864a9d5e12a710468e7e2599d8e15056c7d",
            "a070be0aeff258a08bc53b03e672ee2c825f6e339d735858075cc4a88eed12ade0c6a1913ad7578aecb3fa59b0efd521",
            "29113475b501177d9ccbe612980fcc0a1952347527cda7d315840de3d92a998a8b7a998657f99c1d2cd38f2894d3cd37",
            "21b67ea1ba1cc9120a5c6b63dd2fadbfe6f2d6622c60c5152b3907aff24a745987029f4328805c36ae81da7c0974df46",
            "373c6eed5233dbde6f7304a8f5d00e2b36eb40030b6ae52303cf8e1933cbcee5e8aac5636ff5f9f27598fc647743c658",
            "9b0df70e482ebab75a888d91b214c65198eda8f78eaeed440c8e9c26ed3a1d02aa79af25db6d501c69e6fa4b894daa6b",
            "f772ca58e7e90f9ab4f8b86c83cd95eca72a2a71f5edf5b03e3dc005cdf10b8c8c3ab16d27cbc50397de062680dc5a33",
            "ef220c7d3079e1f1da5709350e6a9b8cacf93c6b65cb278be07225d731ee36bebe5692783731d0e64bd86d8da9b7037e",
            "9d7456764ab179e53fdfca1859b2e2fd5599678d0f2c029b7d281012fa73a013460abaec2dee4cb7e18b22d420916a55",
            "8ce051d9668d430d781367c2eb2c2bbefed6f1572e0e1d227e2f4822894df2f901faaba73c4cadbef6cad6755290b724",
            "e10762e41d1eb78d01f584e9cece7a1df91b57ebbf57fda861ddf8c6c7ec52021e0040fb53f59d5d3ceccc9d7fe97d9e",
            "faa8cdf26bcc36821fe8c36bf1f1b8c3b51179cafd58b0aa0c348186f353438655c216b58a3b09863140864b34cfc3fc",
            "8989798e6dfe97ccb75274f9cf07250404047eeca6de5af0097dce7fcd37a2ae5f028b20252429bfd3b46e93b00ed44a",
            "e6eb28aa52c753b7eac56e31c51392316a0b616ec63e721a4b92006417c38e971984237e3ee78fad79398552b09e3723",
            "a98aa2c7453ddbc5aeab7e42cf28bd77e743c5a8d520413878eff79a86677e3830d1b6312b41a971dd6dd6d884e662f2",
            "56de5aac10ba684e2472b708fd079d78ed229d754be4b97406b1c1fe141914eba4f9a0f375d249142ad2e88df4fb7bee",
            "3c0a2a5f77652ffb8c9019eaa0d5b7f4e48a5baacc7ad9fa6cf1701cfe45b69d7bbafb184cce52638b92f9efaeb05695",
            "a13200b1dc63b513f03a51e82fe7dd71e34b23dc1bda8dc03fb5c13ff5c1ab658eedcddf626459358fdfed91f4e4afb1",
            "e0a53953ba898b26e29058a386e958a8c313d754966ba20c141937e826d19d904f7289698910bafa5228747167d8a875",
            "b09fea55b8e6571ae7c3e5048def3ad47a42813fd7e1c32a6d81a69335d51b5475a3d2a176b6e543e4f8be9d0ec31c35",
            "89db3072598f7f84ed9e7da7b090a23cd06777b614d17d9a53bb301e24e1ea11821142ba849007dc8f8e7dd189016e13",
            "baff01dd5de5ff7a9a9620a14418137e2467adbe0388ac0d5c4ed93daa7a4d8fac39955fa1d5dc0d7b7acc57ba2e4774",
            "92b3764e6c4f1cf2788f52a83e01e4d7b667632264d39e9cfa06ebf792b8fd99e258d504f1f72dff79018b68bef03e1c",
            "b683f44ff4b1767bc13ff7bf8ea5fbf5453866147107357317c3b6a058a81d15f40cb7b275a62bc8b527ab64adee08a1",
            "defd3c64f8d67ab480d1629889de22f2f0e2bc8a01911eb2f69d652807fd662f94ed831228cbc78ce4d6ea4de31b5049",
            "99a65e60932126529d1f23c82e9ee2ee461882f63b8ef3f5dc087f79bc3a7b3c70ff9be356264ee66fd45c9cac1d2cff",
            "a0bafe980272173b26dc7aeb3a69d7118f86b8780c2e967ee30c4c8f5b8dbf85aa809a9f83a3b8f8dbdb360f98f1b8c8",
            "8562e26b126c9d24051465698d62180316128ce64164dc8f854bf5a1135d3844b27c6444162ab7890ba2d45365b5bbc5",
            "9d7b1e292ff3bcf90b853c072bdbad4e3c2f79f842b0f55eb185e57f12c22c6242a3c1c7cf82479e87aa071fa12258f6",
            "1c4390f4727b892041180e9bcdd7e51e60e1c15710e9af832770bf6bdfc601a9b6cfe3dc910e0600a0fab37de6ee2ee6",
            "4e0e2676f2d60eb6ceffb8fb736520a653119a0333e894ae4d6c5ac065913097d82d3428093382e00a4a16f83ea8aeda",
            "1a85d5f1b48621cc9dc2b08e6b7a6cdeef61df7d4aafbb6efd83a4d6c564c53a6a5a9766b3f928d8901c9e2f2d90c180",
            "19ed65d12d3a7426772281d91eb53bfd68ac9bad2426814d5b58654900a20d9c6610a5dbba95a5ffda90f467a3b0098f",
            "287b6d28715b711ddcf92ed9e3ef19ca3aaee83db2ed6b2e308175b12ae3b593ef1275d274344613d05bd388b19723bf",
            "6883539d2e7b0bbb06e1349923a108de49bff6ff18a8f9b6f84273858c7174e1a172289ae202f5167f0854c383c0768a",
            "25dc5a7acf2f79d4aba862cfd2d0503ffd81526d2230fe04dab26dca7e1452aa70f366329edbc3096bad05297853306b",
            "717b09a83428c185efe7b68325248722fea2c78e5961d4e8ed165ede9c1171e444afbca075c610831185af5b5ef80e17",
            "d794218974ffd5a5181998883f5e14ae2fef3b2bed7250c69cb69d8c4ff1b428fd7e6a36d205d7184161f4f64d9ac5eb",
            "66e14ce0c56fa7f755d3681cf18f579f1f0fcc3f3bcfccc4ce4ccccac4dae15757097e12f9454c891b1010b1ba7eda8d",
            "924aa609f3953f19455e2e124d6c0c524e3d1f736224c6a970abb793aa6331360a715210166e259bed76a27f3fc4ed85",
            "7fccaebbeb6776a2ef311da63702c706e789eb54250ec61c98bacb63219b30e6f4628b2177e72c3ccd73956dca75b20e",
            "f1af679cba3c4822939a01a5d7c5cf1a14b372fc89d40f68df6b6e3df4fcfcad861016449c873f366ef25147b0008850",
            "6f1ccd949d1cdf3ab399393a99ff77f1ff3f0b9d803f0a6972ddec247da72070f6f65afe6a92f43e53d1511ec50b9596",
            "c0ed53963a6b999cc7242a9b02f1404330c4fd1b1ceee5368a6fb2dd856d929bb3e225cd5a8ca092afe003fd99a5dbd6",
            "c012146577805694dedb28b8343a96ed2b64ddcf1ac7d66f64245e50ab746583522826e5d9f937a6677243dfd346e751",
            "a8234e71d2b0be4ad19c2850fd1c5c99978a89a3c38eeed01321de239e5c56b41537f15c1541df258d91e40355380388",
            "2b2bfd538855c81b9700f97a893818c25e4c11547d99237cfdfb1d0d0ad3dbc3407d22e7be6c9b028b02939577bcb930",
            "efb6d7e2370616acb8912354c9bab781674b9b5d4e42c5a4e7a5ccc80cdf88ae8a0039d8587b73ef1eee0b6e3c677c5e",
            "f9b47b4f1f33aff2cf1fd3afb68dcf1da3ac7a4ca6dbaef24f97e1940081137eee47c7a7fb9a8790ebd23befbd2d0e6d",
            "4d6ff13c5d6ec71a2ce854fb9f300ff4d3dbe387cf50a14638e1dc3f3409c5cf4ead7897858172d4d2aa0ce6dde478e6",
            "cfd2aab4dce6aa07c0279681a8e89b26df847184a1fc23fdbcbbd6617e476149c67189bc477e72c3fc1a857fa3b0c9db",
            "b7ff61f05f488bad745866c70cba81f641f0f792ef7d0eeaca22c455fc88fadca643d62de1e5a238630cc5386895c3f0",
            "072f51f3a5fd3d15bf4c0855a198bfc3c3aadf8fab7de9d0431fb36a7c53ce35bc23a7846ad4b4ebf877a9a0d5f84c36",
            "a10a966ca7f8aa35914ef4a8855aef3bc9d21c0cccc9d7430c5df9207bac6ed0455efd441137f572b4dadd02431ea743",
            "43b96dee46dfa1010cfaade078f98b232c7db859094298c5a806ab333f31e67281361723326828b9330e58db1ad11fb2",
            "b7b37b7bb1e51c7735f3c6709929a0bdb92ce0e4bac82e4849ae726ea82b06645e6607dd69f46c135fbdbe269e934366",
            "7233072e42c862f03eaf7654e7da8076ef2df25cb1559c58d7415f50ce3c2f4df8b98bb720e9431a19991246bbc0bb7b",
            "da1768e8166c85e3f2b51e6c5f2180d08b4e566300d516492db715a54f8916d366fb805790bf5707a3c60a40f15672aa",
            "162ca120e019b8e6184861d8eab69f9ab5e21faee2c37122d58492dfae82384e2cec1876a52951b417d5874344bc9990",
            "406a4f50b5c254ef7fe0b1eb98db4c7ed3ad44c1566e71db9cf6c9770ae77ea769981ef0dbe447214e88e791b7998fc1",
            "f939e037c9c19fc5a97f4911fea850b5c4806d1fc5a71f8c70472ebb6636c5a0dbc364e640f182500cfcc3a9609eb75a",
            "4114bb16724ad0078f13d07aeb4e7273fb731dbfa4cd3696eb6579f898cccb4d1a91e70a9f5dd2632127f1163467502e",
            "48659349b150af749e9393e3932f888648b5a9ccd147e7f828a6601ea571b0111f7544864b90f2c5bf230d78d194a213",
            "2555148d34563802318fdac1e35a8ce37f97aa1af5866fa1e85e78593d3ea681f77c261b9d8f86f01a2a3dc8ade4f28c",
            "54146771dd0695ffb3fac7497eaf0c022ff2438695b73fd47b636b6d1195ed2202ea164c2c51a6993d8500adef4f8b02",
            "08e391590b18c90b031bad58daeba716fc7f0ee57329c2692ae7075b6b7ea2c7e120b6d47f6386193260c910ade6a8a8",
            "48a520e573786764e658103e69e46b87e3d69c7b2944ba4a49daec59b243f03fdf1f0afa39e0dfed8e7b8adfec957bea",
            "e09f5b74fe8e96c770f5b70d3b4f0d9f6e5bf80b1bf0ffb289e1a9f13feb987f071aca1fc6ffa6aaf9d4cfd36ad55f70",
            "44fbf7b5aba75e9e5641fe8201c6bfaf893cf5f2748dfb2ff462fd76c5fba983a7d3e5bfd04cf464f2fcd4f4e92ce42f",
            "6492fc7e4ef2d4c3d33cff2ff891fedf64fd4ffd3d4d3dff8220f9bf26a24fadff99f5fd1d0f8fd6ff92033e35ffa7e2",
            "fe1d5d547f33ff6ffdfd978ff924aafd8534daffc718a72c0707ffcb10fdf157fef83224bfe228e0bf00504b03041400",
            "000008006d996d5d2fc7eafe092700006e330000110000006f72646572302d6d6f6465322e786c7378ed7a65509c5bd6",
            "6e37eeee9040709740682468700feeeeee16dcdd3db83b04d268700f16dc2d10dc35f8e59c33dfe40c93a9faeeaffbe3",
            "ced3f5f6bbdfee5aab6b572f79d6da4b4e12120a0b000380030000af007181238ca9400020150600c004c04168bcb3b1",
            "7634b276d45174b33572d06274b5b224c98582a0ce064000fe8bffaf519e282dd3c58c167825b6f8ea14e70d233d9aa0",
            "743375b5a00f14916e3555a4525fdd5b98f9055616865239b3803469a2becc2bdec721c3d71ae0ba9fed62bdcd12e9c9",
            "06efb315b409eb8bbbf70697b625b012dacdaa97f528a6c5f4d9797852bf9443ca778c45d18ee5635516904fe9222ade",
            "f844899a136b882d4622e2f333e8f99eeabf6d4df8da1cc018295ff4aa309eb07c6da9f0bb734f569b38a93a46cfdd07",
            "fd2d2578b80e02e850f11fc1136d9a2284c171646744e652d9acbd68f655ef40a7f51d9487b1f59404913d98dde70006",
            "83fec4c42c5ec6952f3a49c1c29772374dec0860952386237ef157a4a6e6680a16838b490c78f53f7857b3a8e28786af",
            "db49966e6a6c51614f7d85612f971b28069696bf0935cff3c88532bdfa7af7369374f01c4f5e5d233fe74810e7c0c19a",
            "24d74b2676c2c9d647ebe025ca8bb5da77311f93d7af3cc926d7cb52586287867533d343d4aa4709cbd340e96590c45f",
            "e6775f9e6e419c6279ae16dc863cc2ce4e7f6574c752853378fd5381e14a29835894b14bb4aa61d0b8e8fdd149d275df",
            "ebfb27977c7c8404c8fd8b47839598c9ce9f56524fdf223e79b48ebd91a50313e31feffff5e4ffe20f94c749cb763123",
            "051d89dd05f75f4d2dc5a0094aba088835bf52f4f2115f440547e691132eb55111a12912f8438a7ab778dd7d5d587a1b",
            "b9dbffeaf23b5b7938061637599d4b4550f28d9d132c2da57d92885df95a37baf3d4c5d465aaa8841bb5734b346d5943",
            "718ee4c00a99c8c489d027b45146150278cc02a2201c9e545ff55e69b565fd9f95c0294185838aa0ea7a50f66efce847",
            "ef99a4c21d64be104ea9fdf154ec7e077d93b94a8f7ecb5717d7044a0aa69c66d5fa99f65ca46e3a8696ed67421fdb30",
            "ef3d1864d2fbe549d9307ebe1e15eda3c8bef056ffcc763298e1d60d82ef2aeb3f1ef00b272e5e06fece493e8af8b36e",
            "4102001bf00000ead327ae964c2e36f616fa3636167fa4bc72652b9b3c0eb407aff107d86e175c422992e29db0e3dcb2",
            "c9ca52f68565f15865460e6495f67416e952f3c7db2871b17963a7af4e2009a97d7cbecc6493de742f6d9582936e238d",
            "1077fd384c4b822f2342da93a898677bf4c2ad669fc3a638e98487a936596d86474db21e38b4e0af5d99baf38000276e",
            "96648937be46e0b93895721e74e42a811ab41a7d84c2d30e73c201b565464e1525b339bce8f2306c70a300dc060b9efe",
            "27ead72fc3cbcb5c2bd849e292a57caf0f6a4ce781155fa37ccb523cf60912424b8121b7b9f4930ea8e6044a38e97c93",
            "8235b7188bda9833aa923fa3564599912e6d314250a24ac3a4c294e927e1f93508f350efa8664937a4e1a68359385042",
            "d5e3f8b0b982b9cb36b6fd95239d55f378f03ada22447cd4795420e09661c71633aceadc9c9ba825bc9a732a3bd23611",
            "cb1c0b2526a8d5cae2bec8bdcd4a7929d513d6a616a3c57a1086918fe26c35fa79c4d9b1354e6404b79c6e0a29e46dbb",
            "ac14e1024aaf3a66087a52d8601acdcb5075206abdd73ded96fbbe298dcf5aef9bc77a3af2e04e082fdf2ae61a7addab",
            "036c05c9bcf81fdd117ed733388b59835b8759fa2fd9a587b03eea8444249ee157a535aae26eef64e6605bd025b0504b",
            "437945c039f7df617bcee955dca142d44a89f6f8555b4aaddc7913931ef3c3408a5725b0ad3eba7900439bf46463d849",
            "445b2b761dfac31f0f8cd245d8a368444a8e14840e3804d643f2e1162a515bba4e3a99af51b4fbdfc7c04f97fc484220",
            "70d097efbff5dd8eab8085ef1050d9b83dd339bef43f3f5b62ee6fb590759a6866f859ba142f221b79a1f57e4bff0d0b",
            "faa78e8f5bf5134a4c5fe8ecc150af2359a01387e7b3beb8aafaf5745f9cf352a358cd757bd8097ea72297782f484638",
            "d69f43dd6d16963b4b58d06ab530c6b1cb50a39d3a59afc25ce14de0a48301ca8dc70fe6add4a88a626441c6df66dd70",
            "93b9d9475ef8d964e2ec368cabd929dc0edcd5b56a9b3b043d5a62da32fbab8f473f74b3745caa5af89f0a97dd507d6b",
            "00f0c1f7e404f761ad75dba15ff4b19f9618e485f43a7f77174f97d2ca30f51c856acb5bb89cc242c24da9b31fc84f13",
            "4fe2e904c356caaa21262c0ac1ec7b5aed4f5e14cd0231cd1034f30d23c61cb972fb2315fa2ec48828ac9b40cacd6d3f",
            "75abce2abc4026a5febcd91afc7ca8e729d62a2ee1e1c39dec6b36a5a08628a7c5fae647c913982f4cd7b15f770e62c3",
            "c69207e77ec436e10eec2f68b997dbbb724676ad5f74eded1d1fa99dced8065c4cf8a81f61bcc4e4ea95c120ed0bcd46",
            "762b09bec36aeecf7919a9c9250d777909cfa43237d1776aa560a16b787ee364b830e5f8f0c6b5da538555bdb28e749b",
            "ea1dbee0d18b5d8b60e74bc8dfc584ba623c09dca76831f7f42d01000ef81413feca9d7f8f0cff9347814f7914f8ff38",
            "8cff3f4779bc8064f713233cf27b7c35b454a8e5843ee7881a56bd2a3e438c366e4698464981e9747b848cbdc0487f31",
            "befee1ddb207b35bf37dc99b7d441e331368437104a40afbc0ba449fe8f5c161486684719c718e700877f2d9ad199566",
            "10dbeb6e51c5bd2855a94fe4d928391f36d215232bf2bb9d4cd7f31d5b135f49e89cc8948a28f783cc9a7e786502afd5",
            "28b9b120ee721aa09df79b1bae1f2d5592d4d76a15b218e66df5a05163110db2cfb16180fe776f6ac9fd20da86aad5e2",
            "81b2c096c44da50563c1d4e6ee2df3587a8df8fd85eb20f36e66bd489c13bd1ab85a2d4c11f1bc4cf39203285675c9fa",
            "261d4c3ff39d4faea8bcd3be9431ac1fe037ad9552be7d9c6d3db91acf7e9c3eab15dce403fcceac9060d5bef03f512e",
            "5a6800000ff0cf54e3606a64e4e8c0f4e78de50fcb4a4f58085b60460b3a42f48455bafa44bf3047de7b1217438f5c20",
            "e7de44d322b761473a5d17914f945470b41c3be15a4f88d3290891f4e66c98efa6efab75b493f0a56ef9f870c229493e",
            "144109e29465e179df5573470ac69e9042880488b9005577c7eae141cbf9dbe6b235904840d0a0c8714ff3e36be31f38",
            "e67a904b82fedc05ec60db063642630dc8dcedb8dd30613dd72ef31c0277d2fc4d822046490abd9c1f9541bd192430b4",
            "6b726d4ba5d39966e02f302a880e6d16b0f2330ca4d3a5f28739fef2f2838e60b3cfe87b08b7bca74262f66c3fc38f89",
            "b3061d8f9868a72408ac9c7620c86be337072fa92c3b03f6c593ef7502d7e059ab5596dba914cfd9c53d666b73a954ef",
            "4466a551cffa8cc633680bcfedd8d5e9bb984361bb060c5f8b0549f2a644c75febf24a049d0d706cd61a5e066faae57c",
            "d35bc1669193d989281a4ca979bbb7b955e4e7ef792995aeeaed8f900f59de4173b2de387db30c5b79acb4ad30277ca8",
            "05cad441a2c9c66afeea39151c9dcdae111c63c74c906b6de3a1fe96be5d229d831c5a9cc9e0557a3f6b1e2ad4705769",
            "1df683c2ea8e45a81ba5e7369201bad1a761cf97c696d1f31b6d2dd7673c0f8fcb710622a98fdfdbf804f678876c9bcf",
            "e1ee926ef883ecdda5b7bf4fa531d7cd60b57cc37292b8a653fab65c56a66ba943e1f59ee2b05488a92cdd923bc3898e",
            "fd582d25d29cbe247bff0eefbdbe4c662090977bc93df98b9bc5417ae56819c332af4f11310c75eee47c062e0741203a",
            "ec97b310f1fb9834a58e66df1d9bc63ba6df5add81d2647f0f2c00804dfa475dffa7d5399a1a5919fdf5fea7c1edab2f",
            "599037a27ba36e78a3dd3ffd26cf8ab149f4b5e64ae5c19bbd39e42bea51d342b0723d81b8c10738583d64d82b0c384d",
            "416d8cd3a46034fb187fa257317cbdd592f6d764f2f25752f56fd3162ebc6c93756f0b34cecac6abca2494a6e4571e37",
            "5e1b6fa7c389bb8d5ff87d680d167bfd83358dbe69b976a061fc3246912355af1773fc4259c83debedcf1dd9c3751eec",
            "d8149218d98d9787eeb1f23bda60da22e6881f3c9971505e1f1dd00da0ebf5df67c07b00dd0d94913ddbd42259a4df26",
            "ace24df2a68eee2464ef953155563dd8795e7c38ea70f743f757c3a0b8684f08c5acc8549587e1e6466413acdcf08177",
            "8c529bf16f34bed317644c0cc64b42f0da766fbd3e2b453d5439d166ec9e44998b5b9cbdfb61dfd8be4fcfe610d6b363",
            "2cef8d55b7d721d20664064712822f1e45b280ab3b58996cebf2d3f86c695edc439f28e49c77be2d9941bf21d419d759",
            "f84ee498765f6dbd03ca96f2f6bcc5efcb18abd65b43a6ee02d51df3f0edbb98e32cfbbcf67040d421be505febef2494",
            "eb36b5af8bb4ae9b3f308e5ae6df4c3fe05ecb309eb0417670f5f5339d154aced41fc104d34debdff985b300d862f484",
            "31f3956952d76984ec999434735f25d274d1a11670ac6104dd33c3919cef7a355c8370fa3e4025e7bf61d225ad0f6223",
            "262ff22b2cdde381a8d0863fdd5db0a9d2f172b83fb599515199617af443d926954fa595b426a60aaf5974f1be9bf2f6",
            "ba423c3f6a7b586baef23ae81d7cfd7091d5e6717fe5ed95e5ae6a1e48a37dbfb937b3e271bf5bdf763793d56a6f7a04",
            "efa11e246b2a45ce458abb4d13bd068ca4b17cd8505c94d1c7d033175f8573e4568c563404462ba734da7aea7f3c6ef4",
            "a15f87e0d56cad6ec814d22f2936d8ec7b6f0944e534f1644c87712ddfa25573ad91c2391ce8a989f81076d268710894",
            "a1878d4ea2c957d995e70c09113e4b59135a83160f11bc01798984aceff0e64dd6ecbdca4060bcd6cc19a5de35d8651b",
            "08c854f755843b98c0bbabb1478752d1a8e3f488ba496fcc74171bb122a7471c22daa00a5033f6084946e8634cc424aa",
            "ade940ad0ed3249755d657ef8dcc879065825b376523936f32cc84735d8e9b203610b394a927ae51b2240bcc3561cc80",
            "4d397b6ddf3f0a831ea54065922664b45f748a2488956750a979fcceba0acd13d5ab7d40c685afa8dbfeea236607ab98",
            "ad9184b0cfb64e52e7da6070c0c29e5dcf664a438d946cb7b2f301fc6e44d483bd5ee025f01592ce2afeabedb9221418",
            "cf4c75762631be3092b27d4c155d9e4ee7f2dddb00313f299dda0c52d4e2cf829701f43f517883cb08434d02e3f1c2b5",
            "d2a2dbf794b733115423a2b63c3a0b4d8c56d10371d0bef7982550613826f87df82cb9671750315f7054c0adadd279ba",
            "b250cea29acc64ddcc2810b4a80e8371abb160ec97fa46009d19536acc976ec92fe4c783a0d08c12599de3d7e103618d",
            "06200159d94aa8d969e1e6161c62079357caf2569e138f75cfd1c4d1fe57ba5b5053ba56c363c7b27148f2008baf7d4d",
            "46c3a1fac7e6e9dedb361b2be78a1fcd419b95398d16a69a55bac4053934ee78a63661152f6b4f308f16af52157a6526",
            "c36221cdeabcc8bf4c0ff5329b5ee60ff56e18133347aecaf3a14ae9aee14010e29337dce02b8d890728cc12230c316e",
            "0a6e96e262146876d6c5e2b37f28dd08b2aebdfe1c817a4dbf70c8762be9652c18aac47208d25da7210caea04dccf958",
            "e12177a6a100b2c4352e7392e61f5c0e12327bfd5168582017fcd20312bec5e1524fd591887e5efd7e5e49a310604e58",
            "d5fe01e4577f5139536c12a12b8ae27e9e7aa9b0b90b76013bf01617b9e0c7279201ea70ea7996c6345d7833bd9ae827",
            "973b2770e0d713e8cb2d317a550a3be2d1bf206f8da92dbb3cb6fa34575471074a32b8cf63ed4dc97d3305f305553961",
            "34c612aaca8f60d10be056f7be53757c99f6fdf3deb0de0a1ee33a969fe8ae0bd3108e2d80558954e3e4284e5dfed514",
            "da20756dc30057aab77e69087f44af962cdf232591aa7e9c984825a4e4559f04e324e2619f08ca17eb0fa9435b86a55f",
            "bd98224d254dea63bed71aaedd4d85b471ab4a9e5234dc92a769652bd461c9b204a8e6619b1e8d1407acfb93385c4474",
            "bacd52c6639fb8e4bdfe31b7c2b5417bca626c167016273fb289bf889580fcc2a1fe93c00ea3b731f3761dbc2c6e5367",
            "abcb9ca743c57412ddae4111759ab2fa98d14486349f7622d327dad2e21edac1259083b549c63e895c86d1f77b0bef50",
            "7596546169d6e5912e2eb3ad3c7630ea2c7a42a9ac43f012bb70a759a80bc33d890babb953a02a2889f27878cdeddaf2",
            "6356abcf064b804237c5a04cb31a844b21582a3af0e48026abe25492b25628fe5a1354a046eb2db11932c2d8acbdbd4b",
            "74d6b5fe8550173e5b7a1491ce926c7bffd252406bdf7eccb2c33aa096fd52a4cd1cd1867cbaaf9c50af16bc9d5bc138",
            "43b5dba40262b5083692b51aaffe84a569a5811ef5081d34cdc93d346ccfb4912a99e8326f1b98e4e7dc92751ba9ed89",
            "c60e83ddb684a7b40b6ed23103b98325aac68aa1f34456516fd7b6f86585edbef5b77ae17599a169fc58e14f39a93822",
            "f97ee13bd342135db6e8b125f76619b4220095b3c71d12fad0df01290a51618b35d9a5cd518c1424502b8f1a3c2cd1ac",
            "b4a704a3e78077996d0aea6b19e571b46e17f0e6208a26d42be004682e1dbda16df1b29233ab13bdb7258ad87d47dcee",
            "76dd18508f97e1e0a0d15eb076b1daf6dbac1f33a89ed8f8c43585610000e4bfb2be83a39ba591c31ff93e5f71d87c81",
            "1fad15f70665703931116b101c992fd7870e254f4d1ea83d33fe55bb86cf182cf323493bce85fa1cd97dafe0dbb71996",
            "0d4e2eba6b51ca028601616e6eee9bcd8fd7c620de92027ed090a2661914ac28af88d48fa655f3a455179554507b718d",
            "22b6a2690fc6177d6c5e719dd610ba668c7dfb911461404e29d4688556045cc92967fc67ae71c4c2d5fc491179440d56",
            "b2a94e6675c5432a1ab1ed5b4f521a950661bc813b6f7ef237338381c9168dd949b28dabce53ba7ae46763d75266ef8a",
            "208aa705a2378b516e2885c07394665fa24fc53209b576c65f0219c78c061c2e530e6390e087afc47ceed8f24948bd5c",
            "e53e5865864c92052dd95211cbaccbd369e468322542b3f8165b746cabefa0d23a888ecc8914d85da02fc9013301f16c",
            "23e06f306a31285c33ae5889e591ae0d6d3a2b0faa2b0f9a2b4b15366dcdc771128cca0776db44dd05cbf4d8e90eaf9c",
            "8504933189cc90c2d48c1561083f7faa5246ec2fd66d504fa13c241e637da72280e7e44522ddc99669aed0ce7ca451ed",
            "54f049b7003a376874d31aec0e5c29a9c381925df79b4855566c90578fc900fed8900a91041d1f0d42ee370999d3b09a",
            "8d8643754dd949cfa6c9f5f75a9ecc47761de582674783849ddc4250435eccc3e6bbecc4c67374a0aeb4f2dedd4074d5",
            "f8f24f50e02695f1a1c10ce52ed7b7bb20a8ba2ca46deed7f05321288f1754ab3477411d1d80ca5725f10f73974f09ad",
            "08439e9e118b14bfb970aa5a31d1e231a98e8d1f43797311b611c88cd47c23fcb9ee3bc36d1c0b8ae8c564dacd88a28a",
            "b7dc0e7ed8d91ac9f0ee8db4c92c660b6a662b85fa8268acb05fd83dd8c2d491212dd996fa3e7be9072fd2edc31ef571",
            "cbfe2cea17ed2c942f028f1ca735d47054f7cb3fe9bfe6a25c6d8732f9cbfa657b1a105d9b1a361afa9555286b043669",
            "a3c1d779d3f8592692dc5ef1dd23ffcec641dd577c954fab93a70beb1f366eaa676f64a8e0686f666df2a7a96b8e4823",
            "7432a341afb6acbee8dc7fa32424148785aed93809d9edfcd9d7589e5bc1dbda4648c8f6bad50d7234217ab694f09208",
            "432d451a514e4ca914528b9e08b57900012bbc68be3ef45b2726be04e825bef1fbebe45941b2f10af06669c7b65b1a57",
            "be076711affbc0a736d343f7f337a0100c0d1ea6b082285eb4625a64721c9c6c18841960fdca3e12ac49f7c13dbc465e",
            "0ffe7d10aaf5f7e6cb2ab715be019e7bb1ad8143e8dfed4fc01ee8f1e6c987217ed58b86f67a2e7f6ced7f167fd58b8a",
            "0bd679cc689e3a668f1da7dc4022410cace234c427e6bd107c08abda0bf57936ca1224ae74b132c34e19ca5e31c980ca",
            "3ce43aa239e28a03129df6e5492cc57e1b7c42ede7634f5a9a64e3f0f3b611c9188d3ad22d90bbc30c45a7c40fd4334a",
            "1c6e4051ebaf2ea90b68f83a64834258e1a286473b2e173425b64f0d9a56170b025f272e54c7fdd4cd56c3cda5967807",
            "8ec75229cd51acfa8078eefa93d014a99e088791d5138d6edea93e9d4df0b2541d615a52c8115a3da00e828e0b96eced",
            "b413baac1c8630d54725ce5aac60d2c3296522c7110a81da9d0d430746945b22b970075def535ae634912c68193e9488",
            "21d693c6b4fc11df605b8d5017694989507c68244c1afbeddbaeab36c608d956e10dbd4772018d1f6802ed912a285914",
            "9ba585f653ba2591ac23dbef57e80e288e4917885a1329d465d88d30124410e24c53a106a55fa34ab0655a7dfbac9963",
            "7763ddb2ce4616646f55785049501c354157a36651a3943b4d46aa4abddb44e188f79295cb0e83ab46abd2a6de8e2b7a",
            "5cdcbc53da44e07554728707b695dfc957f3998f736c996d5dd5aa018cb76b7ba5db45edd30734ee0dbb48238c846e4c",
            "fdc4469f167f6a6639df442f7c282f48e54f4fb56f6709de2c6d99af073ab801c7e906b29de4fce3473b322fbe13582c",
            "5e3db61dbd38b18c1b12cf843b8de49e2173f361aaf0c91c02933045d82f1d01c5b94198a5662e0810a973f8092e4afb",
            "6acdc5b39f5ccc832cc62878327b2adfccd609402a3cd2436d99cca89ce0feced4de5ade184e43010020a47fba92c193",
            "2b393afc75fbd3cc8694356c16d8b1bc88ab1fb3e13d95c8d56a8c083f216ac8845937e60c70e9da7c165ab0a496a7a0",
            "e9d31e7b6ced0ac79e1bae339eed3a5d8d798b9b31e8c27b7a775099bb8c8328cef1f1e789dac68714ec2fe4963c27de",
            "9e0d4841ad54d8fa14b9e96ba16d7de3dedcf727c167fa9a963172576f68ab616552887031bdf3147950dea959da066e",
            "e057b140ce858523d0cc8639a9229c438dd0115ce69aa21150b53af83c04fa23a03b57e3ae188caad0c9af4233a521b0",
            "1ae5a1a5904263c710ca0ffad793be0c5906bda4938763bc96c8c0fe825051fb4e7e95fe30d84a59815a3e142b4af4fd",
            "ae157ac6cf58e2dd7761b3f32f3d043b465d2ee58e83121b861a75bf45457929459828ea75b42b8da4fa156bb5941e95",
            "ba06761b28c0a259238b33c25398db8c5f166dad72afd816075f01bdad43daedbfe3dadd68fa9d83a1436a6cf12e3575",
            "5ecd71bc900417cd526289f26458cd7acd04d886526be57e7f3f5398b758ca3b085e3363f99ce81bab9d2d72029b4bad",
            "ad1b86cd18bdabad50690c41b6e0cf639360d68711aa69a046dce52e1000a6e50057a8bdb33107547017460638af8e23",
            "d48bed3bc85cbf1f2541a5ffcc277f1825b89adb3dd20209a9c7ecb57bbf0b5395c3a3a1424d451f132a417b02acfadc",
            "36759e09ca6de76b496dd9895ca3e244bd282949ef4747cad49ea185e5a0b9355310f5dc146023085037e00e749baed3",
            "300dc651672dc3e0c6bee88e56e2c7d4b7cbfa44888947af55ece37010fa660a843eead47d26d39d3d2962c949ba4591",
            "d9dac50c0c49eb46724899381aa28d7aa029590e85e97e200a5922b2b71f58cd688a2be4af3e2510a07c487e2b1ef43e",
            "b348c09ba5016edc3401f62c3c1f57d3b045d82681e29e7cb4cdb978634a28ced9d894d91aef7a71f50a655487f40a37",
            "dfff9a5e8ce9c68d99b0bbfa6b9c2547c8d6bde3d8a3396505debba1c29e607ac7057a8ec1328a44182ec207c7c09994",
            "c4b6d57b746d0a315ed20652aef38251fb7c6b392724654fafe89ef02d173cfe0f807cc8fae92cd02045cfd54674255d",
            "4835b255e4ee8a58097b14dfcd08ce010a582cce956ddb282ff6459f8940d2e5b86582486ad93b972160796dc65a78db",
            "2919a290de061805174a10841b10994bba4f7090be10a56fc6b155ac6cead2f72295c9cac7073c28e136ea9af513e5d8",
            "76226dae95741f3dd99650c31fc1f3fb775face819fe8d694ee6e47874675a03e619e5e27618ae459876001075f1eb39",
            "05a549d4c7285ae8f872a359e10af80b17faf4d4e5469986324b7113c7612256dab1cce54a650c6e8830b29937bd5584",
            "a5fd5349829a195c6a37d52f80f919759be4ad8fad6fa2b6a40b56e0e505b73eb1b9743d566b50f93c7eb4c61f232aa3",
            "8ad00a4e725ecfde93a8d869fbc04a40caec92af06fe6937fcc356e7fd6401bc2dbf06d9c12e596c7d041736ec7a5cbf",
            "3b0a539628a7c337bfb13adb31a1ea6f9a8e88957a355194db14a707a89b79e16fb5a106aa5137088bba74b17a9205d8",
            "3591f28a938dd45e3a0da8905715aed047c43142860bb5fdf80651e3b3616371a619012b47e94ae9873462f34d2a0d89",
            "1b1383eb24bc50ae1eb7bfd17370d28fd9293a11e2e56bb6960e386dd4e4f4aeb3aeec5a110ba4f26a4a2977c986b9d1",
            "9ad6a31d974db5cd0fc704ec8f5302bac3178c27bede9f719bcd3c3bfad2d4db4601e87dc43398f3ad5745debf3d32c7",
            "ae2a7ddbf2b4a203020064ffd6a2fdeb10e057a3f6cf238080e8e13f1846c091df11daa5f5b83b5a6ca1b6d031ed0530",
            "c9f0d6df0aa3e6a4c8c686d1cf6fd5ca24f9deb8709eafe8fd22cde96a46903c0e493fbe0dec8c59035bdf140e149523",
            "8ee891589714d2ad7b63f1089bfe57d945545fe1428ee96f4558ce129f76cc049378630693a7fbe015923e290087b1b9",
            "e5f772678b7f0ef8dc84d3c828bf7b3b8721c41abc7f9775b0e0ccec78bcc91e2bc9479f10f9b37ff8358681a8b87fc2",
            "d18bb660f9cdfeec5d094fa5729ddf160d2842e72b5f9e5614bf76ff4fc2f1d7deff4e3bfed3ee5b42be954dc2aa8e1f",
            "a0d689f78b7cb2b1c9c685ee98a8e6de07710b5fa4519de4b7e7d512d8a7fa92586872cf6b775534ee479170e45ad9d9",
            "27b1d177cb252e48db15c01b707dd73ba7c7e8a9097a25e22d1f36cccf5ae3986969cf856b1c6186d4726a7a381b3f6f",
            "322710280b93a023f5be9e4070699324ada768e3d172c6292fa796353ad5f20e3d67c8e715254950662e0ccd0c7255dc",
            "e0397d718d9cce8fafa0df6ede1052babde21f873ea47f6ddec8d5d1c8de5acf52caccdac2e15f9efee25df11dd6eb24",
            "68414774172fd0ae6016b42aea94f74b959468e6cd999678011df150d87888217bad9eb0714ee3d2f915ee84c629d1de",
            "5eba5eacea5d30c7027916ae9c741c9045d513d64bc75b7ca8c97c9c88ef9b64081cbf1d83e0ee0a5b7eb207b3f8b98c",
            "bf35bfea84fb2ed80562b904a30958b57f65ee7a49a4a4bc5d11f1f9dc02dd1db50691d0c734b238ec7b70b53c772bd9",
            "a4b2278c7e81f7dc65328215539eb679deba86b19db1528b2af7b9bd99d8beaa97c9ba69e92953160615ac6ac55b18cd",
            "1e7c2074a5b27a47c98885ddbd8059af57a452a2df10eb864631c2f8b48cf6f2abc396f27b9388d9d016513078f5fb88",
            "0c6a9094002e7b136afab5ba0bf255386346bcf0ccf8c6c7f0e37de994b8411fe38f0ec19961e1ce8c23ce3fe8ee1818",
            "d0f4c2e648a4bbb26cf7b0f3a7bd11dab2a3535dc73d98f23cfc1a8ec4e892aee8c53e0874f2b4beb6cfcf95f365a169",
            "ade00847461d53b037ddb11be9e13757b8afde63042359a911e142af516e4eb19ced42b2b67e573336c79fdf4552309c",
            "127f9c7b3ffe625f547c8d4d429d7cd53b04d6fd92b625f70e7aa2366720dadbf016e5777ff6a997135ce2d3aae4e942",
            "fb07dfd1b3347867aa6766fdc75ffb364a000182050deac8cf1bcec1c564c63f847cce7f1bf11cf88a3b8dffb5da6b81",
            "f3abdef0e4a0dc9ca68bbb11f19662097a4b5d0211b9ee56fad82c98c3fabd433889b4f59e824ec3d57710b54bd69f5d",
            "06b2fcc3cb9b7c7bf84dad785e85a126642019511c0ee74deeabf7dd08a9d1350a43bc1f84bf5a776d5f84d605ab4cb5",
            "1b4160dfc2caed249bbe7910ddff2d65e3e2aca98878b2d7b1a740860e80031ada18c8d9dbd83e31361b7ba37f4cebfd",
            "f76c12f0e78c8b64173351ef91df4acec672c19c5cbe23b6e25b1a38192c9b57d7dce156ea71db8b5696d4707eeed8a8",
            "a42f3d80d03cd77e3a94326d4166a02d21d9ef0c03671b7d67de71baf6dca1190c31b412c2d8b171c14468f43f628b55",
            "c6e4ada633f8a3c6156d098a204a8be111a483f2c880831e4912aa316630f2a2d94492185f4dd4b02b2c32aef1e8ea50",
            "e09528f6d09435258d04c14b5fb946a817b29dda8d4609fbe04dbf961472c036cdf984dc888dc0f45b2af442fbb022b1",
            "90937ceecc173a4e56f1979aacc24530ea4f156a4a9d4bbdb10b9e4bbfe8f6fda0210c69903746fd865618e963d66ae2",
            "75638fbe2208c34d05490ecf7693155eb1e8355b4557fd44d9e159c7ded9c9ed75b8994eca4889195fcaa555ab7556ea",
            "48fc8e66da48a8d90bcea58f2df5cdb787875166e0e52a596de22bc8b75083e1d125e3679791d3a4ab0ac3deab8f86f3",
            "77c25a2daf7cb8eb006d2a3345359aa99925c3b9d5fb897d5a829af225e8a357bd772c3569a20fce7c4cf92e31e7fe8f",
            "a93fe1e120efd631838801bf33615ed7dda48c27fbc487fcc30bff66c27ab6b6ffb5e05f488ffb64fdc781f015c013d6",
            "f7831427b25f7b361e964c38128d7ebb8c7312d79c104d1cc9a8df128bf6293b7eed86d51d29a77808595146d3bdcbe3",
            "ce4c1f59cf54619d1a240bfe1102fbcb1ff490459e668b35ba57e1c445c9f492c59a25e52e72d50ce4ce146de9216ae7",
            "405202317ce106c56a6cfd14dbcc7304f34411436e7b5a628b40c55e8a2b252336cd0f2d951daff806cf105108a15404",
            "dd08507ccb65d3d97eb6e514cc30e8b8ab5a4202254f18a02c3e0b7215dc99399e0fa8a478ae9856fba14a0b383a7cdb",
            "33b5a573e8082e48b130d293e8a89167f75d2546b6832cd59a974a4dc937b51df71121ae357cf8363fa27a2949616a87",
            "305d6c9af0aed3bcef2677829d3cc2beedc3f291f655f4e86d5fb2f27e265451c7f2fec6f1c71cf62dfb45e4188e9796",
            "90ee24afa3dee3c230e4379b17923ef87bbc253fc2733f7c697238f37e8294f7c449b488cff36be60d17753dc57ab675",
            "e46dc66157a70e5dc015f825a8c7a38034c11133184c4cf03d889dc08af0b5164d7dc0a9e7a64d9e2714eabd64f27af8",
            "fb873745b72fcea11e5d274e7fdb9312a266dc667b32d8aca7a8cc00f80d8bf88b47fd1b97f8934c7d895bb17822537d",
            "2bb0b7c412cb06e3d4eaddb406d6241c78ef566eec0b67d86ba50531c0f747b8bd3098012dd7ae4457d70fab3865ad16",
            "f1ac5c551123e25cdfc90518a44b1c33e2ef0f138dc94f808ef0156329f1401e350662621ce259c110d1d6f7d354b1b9",
            "9e326968fbe92c4c84fbedd111c22f4089762f02389ad2d462448b63e1478a86ee8d627e60b3ce2570da7d948f3104cd",
            "16dff12d2826c635b09f4ee6a080c8f12e811941ce65e7272f0431e756cd91384b14c3c7383f64e27e20dea35eb4bd51",
            "ec89075bc0c9594606d4cdea99204ed0b9f3025a7c8e8a03f01211180b69888b021b4de9dbeac7677d1fbe16b0ca402b",
            "cb15049b29475161b2109a681a1aa085f49b50c72ad8c8c8884c17db783e5aec0f27b0dc0c465b755abb242c17e571e7",
            "1ba8d880effec9df81100c80ff34c6fe1cbf196a7faee05fa766ff8ef1a730f5b719dae782cf27097fe10ae6dfe60a9f",
            "0bffebc8d1df8188f8a7f07f18407aaee7f98cc92fe823ffe78993e75a9ecf0cfcc22cea6f27089e2b787efcf00b23b8",
            "cf0e239e8b3eefeafe420dc1ef7bbccf353cef9bfe421ce17feea23ed7f2bc25f60bf72f7edf207baee17931f90bd064",
            "ffabd2f2b9c2e7f5d92f4092ffafaab5e70a9fd73c7f03c5ffa6027aaeef39adfe8503ca7f27d9cfa5ff95d1fe1dd554",
            "00c0bff1dbe7e2ffca26fe0e07dabf89ff835b3c977e1eb87f4184e1ff328ccb4942c3fc2188fef48a7fb2a01ec63f9e",
            "fe0f504b03041400000008006d996d5d6a04c3a40c290000d6350000110000006f72646572302d6d6f6465332e786c73",
            "78ed7b65509c5bb676e34ee31220b8bb0497407077777769dc837b487082bb077777081a5c82bb078204bfe49c99c919",
            "26a7ea7ebfee8f6f9eae7e7bbfddb556d57ef7b364afbd5a5e0a021203000d80050000c4804cf2bbf04cb0a7110c0080",
            "0e8005d77a6367eb6462eba4a7ec6e6fe2a8c3e066634d940d094e95090007fc17ff5fa3277e4c168218a5ef84a2ad3b",
            "b8646fe2523f726f26c0575654660a06c76081dcc48fa9f0689127a2cc42fa287f8f4237c1c4a0eda49f0e31b6c4c428",
            "021869dc52d253881f0e9e76b6d0e04eaeab770c8d9d9d09291b70814c5b2984a63639193b38444384897f5488e62cb9",
            "a628a9f88afd751e70d7cd78db0822acd98101a2106a58e15d48b763518bd4ee0451ed0b43eec64edc9645e5af0cd7df",
            "392484a5b21b1ed2e73c54c0852a44050d15afd7ce19133b17b2360778432cd177239903bfee8dbfbbb3a978e7b93925",
            "9c53e9e69ff432d300b1b1e66cb2bf37e52a02d43e802c2990821bb4507215cc2aaf180afbd1b2676a8ea5de04a2cfa6",
            "e1eac7ce209923eb3e6eae8e1e16ff7ba33b39fc013c4aee8aab3e7a71dbab9bd1e13a1731101a2174dcb53baf48c21e",
            "536e4d95b8f7c920ddc296f61be1fb5492c4a865b0ba5a1e4ea7c0312cc14f7aefc66e350bbbd5988d48925396791b89",
            "8c3b19ec2af69bb63f7ef91ead86b7e4fda03d909e0d6c71223cb57f2b33629b03e317c4441120a0e2a97671ebb079b0",
            "ba8b595f306f31c5d362d2b69aaff7b4368f8f1000f97f33ec5a1526d2f3a791f493dd223c19b61ec8c4da9191e1e7f5",
            "bf06fd5ffc44698c8c5c371362f089f85dc8c0d5f4d27b1421295741f16662659fb7125f81b5513964784bed94f828ca",
            "2f0220c47c5b7dee3e2f2ef146ed0f105faeb39546a061f090d6b9960527de3838c3d0508012441d4ad77a505da62fa6",
            "2f93c524dda95c5aa3694a1a0ab3a406574845274f852b51c618d45ec0a1e7e10763f125fb69f6c9682c1bfef804362d",
            "a47454165c51cf91b91f3b96e23b9b90bf872410ca257d38918c39e0686836ffc96bc09af8e2fa858a9239974585613a",
            "889bc45dcfd8bae3bb704a3bfabd17bd6cea8022091bda0f9631b17ef2cc0b5fcd1ab6d3a134f71e0eb8ee92816f83fe",
            "118485cb60bf3312327ae4e17b080080081e00003e7de366cde86a07b232b4b3b3fa19f94a556decd639511e7c26ee43",
            "a8565a2bd177820a9d44969de21754a6ab8ed2026b58c2a81bb78ff0610a8bef4f789021cd35d21a8e89c52bf87cee86",
            "875963b5bd96e382ead6b3470e2c4b63e5d467dde2673474f9a22eca2721dd9b6624554d4a0b69f967f4dc4a7cdbeeee",
            "eb239ce9d3779cef4bc15e4fbfbbc88b8bb2c3b169da24b71298462ad73751888891c009a7ccb18134d368ded44e0832",
            "303cc4211fa2f58e588d7064daef259173572ed019a46d84426dd147dbad1f6ab180a961e183f964bcde6825af9988f4",
            "7af83b4922235fe1a7f822df0be569ff11047bcf37d636a4ab071dd1fa0d49dd12154e20e5e4826cdb83a03bc8f1be84",
            "f3b04748bcc016b8ce6d36e81d628537df543e811accc8538155a04bf551b0c8553d5152a2325d5a7ec89300f3cbf494",
            "223d5dba337f01ebbc92cafceb2ec9f822a159f899cf9c43e4637d4d08dd1e232bd2a19fe09a5adeaa202ce697e0257a",
            "72adc0d36f20cc6082e0855eaef63785c9201899c0752561c444844bbc9c89872438bebb5119b9985bab820eb272bdcb",
            "60c0c6874738876aeb8acfe34fa9af9766a895c530e23b37d3d3b2f7a5d06d6aff3e9f6b4142317fd824abc19c725695",
            "3c9d825fb68d58a6a1951d276841f2ad9b788ea49af4ba6a52f31a086e5a141e0e6630afb8f2e04b48f2edf58f0071bb",
            "bcf1d5471527a4b026335a71bc50d245950bcccddedbcb703864091cae812c3261b62e30782aa40f785ae932f62c6146",
            "67042d28ef32664a6387d23030b6d50bde5fc5e6e1d5e1aac25e36f3f2a43e86bb12e61e9f30742fb86e15597a00b9f4",
            "ea61be1466d5b90f9794ea463b0b4b7fdbde3ab25f62e4df4ed6f64f124df21f743906a60bd0fb3157ead446e7b1e89d",
            "52a77447823a9579fb6facf00d13bbb0dd5fd6a146516240e5ef647dd8e19400770cdbb8b88b24f5f7ae86cabed729b9",
            "014f6639ed2177729e0e55295dc3592b61ed51118d0be99fb972a654f43146f2b6b3b3b94c26445e4b69a309e7695b05",
            "7f49828fcc9fcc635d773322b008792393853c5db6c34132874f6dc5d4666be21b8236404c13b81e162186b136e75fe8",
            "468aee3ec5a2204eca5ab7c4190ea11526cc9773f798a37415bf0b040af482060353e113f8bb49d1b581ea08a0e30d73",
            "c08a5195df11d41c78381b315cb671a0020a27c6bc82e13c4358d79c2cb271fc6d04f052626cd7d591826797b8d8fbdc",
            "484f2776e9d00754993637b821b375ebfa21fcd6f6aa1d7739e00a4a6092d39d7392e90ec4daca3579c99deee3293dfa",
            "2afd2be4974dafdbc1a5a5cd728bbb1dfbc0d3c9b76d726ec2bbb05aa7901aaae2a8cee8d8ee70cae5223a2f16d842e1",
            "5c80708c337bfb46d7ec45cc542cbc3ef1f3cec2f1be50c727cbdf6b128696d0ceb3b310b3ec8017b562cd77bff51407",
            "b25db0244fbf984202002f00b0604f9ee2cf88fa577ff1cfe80af6145dc1fe8f9dfbff395ae23a257b9850824ffc4fe4",
            "b78ff3cb9ac2ad9b6044d4bfe51e40043a2d04e64b0c0a2fb5b56685b82048b56230fa254b3c3c1ac50cdafa5aca9c45",
            "a5d754e357c9c0c369ee49da582167f4cfa6888953c16ac16fd081763fe8eb69eb1d2d7f0945017e98202057e84583c1",
            "14baeac191082acb8c3a18ef93dcf2a21542f41ee126c653201d9a741aba6b84c6601dd994077d0b39c38a315a11d969",
            "afd551b1086394f5a935ef8642082e78bb0cff3ddf06dd39a89cd99e736cc25aae2eec4707432c5fa90445d087314973",
            "1e83a822f2502f9950c21d3cc550037f1ca59e60c936c8d87bd6e1f50bebd9406272de04c2cefec6cc229eea7345b363",
            "b520b8d63023ce36bd17ebe4b7cdaf8edee0dba67e0d798f74e046687dc983d10ef81dc3ae96c9f31b9e5893f0148f70",
            "00ff8a458ee626264e8e8c7f7c30ff24596a6c67f8e2cfe71ad41630be9470d42361bd4e37989fe5491537b1d44c126d",
            "832a654466122b7f731c4fb04ebf8986653dae9e63f7f0e54e973f4d456ab9533dac29d30147015c539fda59677c76a5",
            "595f30cc593bb1450409aaf56d9e5c49b95e02251bcc2280731d6e01abc6e1d3106e753658af4d288179bb5296c2d7eb",
            "f1a9611768372e7c457aa96cf33e8cc48fba708dee817e2ae450d60af898e6ae4138e68ddf6f6c48793c73a84f61651d",
            "1c395dfc5e4fbbf7e8d64de84053914b6e32d68a765f959ddce2cb3512996d64f4b4c76a25afddea9093070946c423f0",
            "25ebbc5aa9cfa88c74608e39bef58545974bf6739d077e8388f792554c7d57ba09c17e4d66cdde7573c072ad900075cf",
            "2a553de8186322f8c22e3d60af9e0e6966de0b4fd8a229c22cf6c8b45af72eafc484fa3ef0f653a85f71e64152d53488",
            "692c2530f6de4ec683f1181578df17518e7a5174c8e0cae1765a404cae7c37d7a4e776835c659bcaf140925aa06b2cb0",
            "d5d262dd096fc728075f60344feffad59147b71b686b59446f91f522663b2e89960afe7ec16096f028ccd309799adeb2",
            "e5b1738c0bfe3e4a39f418ebdca35fc0f202f8bbd5a6ad9887157ccac069a0fe76b559fe58ed381dab45268ce013985b",
            "e0c763238dba45b23e0789ef9d9f7122bd972275a8ce1c482cb544156382ab1edb48cd1c9a8385a8c8f435767877f83c",
            "a2e39778a730e7fd2a2539d03f07bc4733cf6550ae4e1f275c490c40259f09cb4789b7d51702eaefc93c3ce8d4cd6e2f",
            "db02f005858c289d36b453584cb7a02d0d2096840278f2d86bed1bd8ac4cb520b27763f6c3450cdc462cb35e7890e56e",
            "a3053348911b7cdaaa0bee232086a659936f5f2a9c4937af6d815443706cb781519ca527512d563c1e17a42e32f1a8b7",
            "ad412d81bff5590d7bbfd9b9cad8e5eb4533d92a50a24c6b51339d87f07ef8ddc7b145c26a66f42acab42b5fd460cae1",
            "0e5977880cf92f8b64a725e9987c9deb1c5d9f3c930418626d5598b3a20743c7fa07dd7893180b22044178a3d039a7bb",
            "962848b122554963ebc6588f59eeaa544e73ae228871e8dc5b7b24f974634dd7d745fa76197ce894aa2b99e07b002b7b",
            "c93bb74f356d577eae3f68ce8b0ee95ce7a376b0dfc053f5d681f42e1e480b2ffa888d3110cbe5eb39f8e9911d2bd028",
            "9125042786921028cf8677f904bf5714db043d62b8dd681aa64b1072f7aa0b6b583b2686e02e66b15fdb9d5c5ebf7cf4",
            "b5eb64ce19f63d5dc9c87426cc6fd0f180f491f2267ac3d54ae1746a2883b36815ae6b1ece9d77174f637ed2abc25457",
            "8ed65e88e6aa929da8f2ae8e60ac09d3ea07ddb0e6ab046579171f446a66b1b12e40469a6dabd4724bbc8bec6cc134b4",
            "1d61c724100243c1f24816013faa2b946fd9bd27d7477c44a7b3d96fcfaef18ef1b74ee648656aa0170600c024f959ee",
            "f983764ee62636267f5efff02f879a4b56648da8bec04d5f947b7e9e25be1553b3e86bed954f47af0ee691aea8c6ccf3",
            "6b55eb5f481879c2c21820c15ca1c16a0be9a29d2584a080de07e013bf17e8ab90025d932a2a5e49d7f37e5cbcf0b14f",
            "d4bfcdd3fa5e32515e22a932adb8f2b8c962ba9b0a2be13e71e1efd91622ceb2c5fa91ae69b97ab061e2f2bd3267b241",
            "1ffac485aab04706ef8f3db9e30d3ecc0f4944efe5365f1e7b7c50dcd3ada529608adce24b8f81f44971443582aa3754",
            "4883f302f3305245f26ed7886296e18d5bc599e24f1edb8bcb3c2861fc54fee0e07de179d2e9e18f1aa081467ed11117",
            "865e96aeae08cdc383c026f469f32d9cd33b8dd98046d33b432186f8109c04789f5d8fb6ebefc5c063b5535d869e29e4",
            "f998af73775ba0c68e433a36c7f0de3d53455f8cba834ed17630a6da28bcda8b47d10cb0d53d8c74b60dc5195cb68f3e",
            "3cc395e4f22e7be34b1650aff0f426f416d7f19d3ede57d8ee71644afb7adfe2f6a77da9305843a2eae6a8fbc62770e8",
            "6a89b5fc96c5cb11418ff042736da00b4fbec71c5417655bb77064fa6ef9f576ea11cf5a9ae9a41d92a39b9fbff99c70",
            "62bae1287a2ded8ce19d7f043380edbd81087aae2a75f206b5308851453b9b389eba9b1698c7b986167ccf044b74beef",
            "d370cd81d5ef099998fb8a519fa43e988d90acc03fbff8800fbc4c17ee6c7fd1ae5ccfc7f1fecc6e564d6d96f1d11f79",
            "9744319946ca969032a2eaababefddb4afcf15c2f949fbc35a73b9cf51df10cbc34546bbd7fd95af4f8687ba6510b5ee",
            "fdf6c1ec8ad7fd7e7dfbdd6c461bc8fc04ce4b3358ce5c9a8c9b047b973a7a0d2c8adafa6153f9abac219a81a5c42aac",
            "138f72b4b23158b46a52a3bdb761cab7c6b7741be0fcda6d150de9c286458546dbfd0ad660402e336f865468b7d21d1a",
            "0db72a69ace3c1deaa48cff0d346ab6330593a98e804ea5cb57d45aed05091ef496bc26b5012a142371c3ea2a11b7bfc",
            "39535507c469f00cd7da596354fb46fb6c8381e99a7ecab0479338775520544835ad3a2eaf7737a98de91ee2a3366474",
            "08c3f89b94811aa65ea189f0fd0cf1e8f8d5559dc08a706d32395543cdbea85c703946d80d733652c526e37458b7e598",
            "494223716bd97ac22a156bd2a06c33863498a4ef2ca0813168d4774a94661f854d0e0bce108530728c3e697f7b635b8e",
            "e20df4e91894751528e8015da5a077b28adb9b488abcddd54be85a1b0a095c3c70e8dd4e6aa89296eb51753982db8f7c",
            "f7003208ba042346d45bc525de9d2f4086f64ed76467141708272a394457d3e7eb7229ddbf0d14f797d6ab4e230116d6",
            "085d06d2fd40e60f29c10b330b8ac589d0f918dd71a0ba9b0eaf1ef96ec7ab2bdfcc641535080b65bdd7228e12cd29ce",
            "dfb346eac021b06c21ef248f4757adeb6c65b194593d91d1b6994130f8ab2634daadd6a2a97ff22b41542674e92f7eb4",
            "4bfea15b0f42c2b32aa4754e9f478e44b41ac05e9096ac84599ce56fefc0227432fa242defe438f3d9f69e4c9e1c7ea6",
            "bde5684ad56978ec5c360d4d1c64f603553d6d8935539a67fa6edbed6c5ccab69a83b73f65355a996b97eb13e665517b",
            "e098db8597bdac3e453ff97a95acd4273b15fe01c2a2ce87ac6566b88fc9fc3277b86fd39490296a55510028adbf8605",
            "8e874bd67083abf2452250698e107e98615b68bb181b2d4fbbabee032ebb67f166b06df5754d24f09a6ef198ed56cac7",
            "54284c85f998437f831a2fa48c263e2ba5cc4bfebb96128735b66989b3cceba1e560610b9614e111c1ecda975e1070ad",
            "8e9706ea4ef8740b9af70b2a5af9004bbcf20e4f0efffa8b4fb3856691fa62c81ee7c9974adbfbb5aeb58efc8505aeb8",
            "b1f1a4803aac7abea52fdaaefce93e4d7453cb5d9358701b7174a5d6687d6af99db1a82d483b5f34965d1fdbde369795",
            "f30449d17b2c601c4ccb8f9bd70a04973ba3357ec053571cc5a013c4aee87ba3eef4f2e37acdc188c10a0ec30686bfd8",
            "be2be330963d80558544ebf424465391781a6588aaba61903bd9d7b038f475649f8e9cc02305beba618cb8e82708a9ab",
            "7e49862984e37e51e4165bcfe4e11de3e2cf3e8c51e65266f5efd7ab8dd7eea643db79d4a5cec81b6ec93eea642ad561",
            "c83107aae7609a9f8c16066e0410395e4476b9cf51c4629ebae6b06ccdaf706fd29c319b5a047e8f511cddc6fd8a1187",
            "44e0585f29b8c7e06bcab45b072787ddd4d5e63aefed58369340bb6f5440f55155f38bc9649a8c806e3c63254d71612f",
            "cdd01287a3ad59da21917c9ac9fabd956f982673b2880cebf26837b7c54e0e7b2d700e35ae58ce3164895da4cb22cc95",
            "fe9ec895d5d239489d2381e2dbc89afbb5754a46dbdb4de640a51ef221d9660d70d7fc5ae9e8a0d323ea8cb233298a6a",
            "e1d86b6d8e3c0d1a5fc9edd0518666dddd7dfcefdd1b2d78fa7099326308b4d6a4bb8797d6823a87a02fd69db681d5ec",
            "97a2ed9608766433fda57806d5b5bbd9650cb394fb4d6a1cac562126723613159518da365aa8ef1ea18267b878864740",
            "8c9bc952f1ae0bf64109fe2ead19b751bade28ecd098ed4b382afbb54d7a161c1eb592e55f0aa172445781b76b3bafe5",
            "441cc607da7c70ba2d50b4b6565e279d969d10ad5ff8cdb65247977cf5da917fb5ccb122089975c0131af630d0092106",
            "5e668f31d5adcb5988182c58ad080c19916c563950813670c4b9cc34e7e86f1de373b2ed10f4e5c48fc633c8e302682f",
            "9dbca269f5b191b7a813bbb7c78fdc7f43d8e17edd18588f93e6e8a8d591b776b1fafbadc5fb21cdf8c6a76453041a00",
            "40fa33ea3b3ab95b9b38fe8cf7b9ca23968baf51dab06f908796e3e331866aa372e5fb512115a9c8827467273eeb5609",
            "98d6ca6e25e8c6b8529d23791ce48d8fcf326f7271d35e8b51e4d10f8af0f0f0dc6ca75c9b72f017e5bde61856d62e81",
            "8411e31795de6a5ab54c5875554be6e828ac52c65436ef456b31c4e497d06b0ba56d463b048d268900b28a21c7ca7422",
            "618bceb8626bb82710f25773a7441511b45849a7bb9834958f29a9c5776fbd49a8d51a447006ef7c5f93bd9a1d0a4ab4",
            "6acc4c906b5c7599d63720fbfee55adae24d0178e18c60f47621f20d8570ed3c85454bf499783a9ecedec44b30862f26",
            "838e9749c7ef11e146aec4dfdeb1e51291f8b8c97bdaa4874e91062fd95312ca6e28d26a656933c64331fb155a75ee6a",
            "ee01691cc546e745f31c2e5097e4c1d201b16ca3b5e3d01aef91b967dd30e24ba3dc1adaf5561ed4571eb45796caecda",
            "9bbfc54832a81e39ece2f7e42dd361a63a12bb080b25a2e35b20866b982a43e3d55496ab220c14ea376826511c137e61",
            "7da32688e3ec4324d3c5966ea9d4c174a255e19c57a99f07951d3cb66d5beb01b65254870529b7e13f99acaadca0a8f9",
            "3e0d6c6b533a548ae3dbc910c46193b02535abc5580464f7b483ccdc47f9813eebd385a8ee93ecdab9b1601167f75060",
            "28c1024caeebde8758ce4ee04a1bffdd0d787795dfeb4972ec84120114e8e1ece5fa0e577875d7c58fdb8755af29e155",
            "27f22ad49abb214f8e384a57a5708fb397cff06cf0429fee110a94c75db9d46d18697018d5bf4c7c83f4e5c66b7f213b",
            "5a358ef763c36f96c7f40347641f3ae37eda3bca58ebdd9087bd9dd134dfbe28bbf442b6e066b662c8160453a5c3fc9e",
            "a156c6ce3419a9f66485cca52d7ec4db8703aa6fad8773c016dd0ce416c147ceb32a2a58cafbe51f749fb391af76c318",
            "03e4fc33bd8df0afcd8d1b8dfd4bca54b5829a7451e0ea7ca9fdade3896eaf04ee917ec7718e9e2b814f4fa3d3a737c6",
            "3f386e6e00323156720259d89afd4175ed5119f82e2614a8d5d65582aec3572ac2c23118a8da8d53103d2e357ea68a3c",
            "4abeb676c2c2f6d76dee106371d173c57897f8681a493208f2e22ac5103a74f8c0e641788c888285fab0f12e745c498e",
            "97b8a60ad7897342a41365b5dbc59dbbee1fb973bdb80af83d062bdbcd8f3dce5f7184a269f13186e7bde34729a44122",
            "c3c2ca84869f05ab5f39448431eb39ba87d3cae9c5bd0f06daae375f96bbaf080cf2dd8bef0c1e43fd6e7e61e4e79bec",
            "4f360cfe6bc3680c3270fd39b57f0efe2c0f282fdae630a178eb593c825da7e923366240155304aae4151fbdb982a033",
            "02d7b416ab4bc92dbe39b1c28b154ef954a3f5523fe934b531f594ce2eba1478a95c2cfb0dbf071d16fc7358214fdb48",
            "e758228df6dbd716c3565f11641dbf896236d0b7f96abdb44a2650a68680e85031f5711cb76c1dca16ff90c0524a6128",
            "d658c92ec4523a52fe55b4b6070662e406569033624a5651bb24415107196bf5f41bf638ce312795b6e90dd14ccdc909",
            "a74510be721486f5fbd05d849a486b285527645137cb8f24e9d9185db123c5a91ae16f508196b4efb87762b27bbd3caa",
            "b8b043bc910b7bb92a04dc317146d656c0c7f88191c3ac178d1f7347fdfaedb5c25c6dc468c27091702814373bae424e",
            "7cb57132bc2306361e34a075d22020c16d5b40df4898d590e20fb0a46c521d44a36f042d18a1caab863a3d55b2a69d94",
            "f20c4c7299824ad8b46d26522a84060d03b6a75f9057030f08c6580ca9664d178e0f3eecbbb988b01f48d749b2790841",
            "46084f3a05beeb3612a5942515ab1bba680dfceecccd68846eab14b497fd31b88425780da9490a2dae0785620f2634dc",
            "64b38ab0fd53b801bcf5527a767fc9e8ea61d507a0ff9d22d650d4d88545c1634b6bbfe847b89e646a4aceed0eed5a8e",
            "5b767d86f2b788922c960453bb9b8e1566adea8651210f8f57ada4e932315c83e8826df63bd064c06f72b4fa44d5189b",
            "f10a9f42f8882c4e4c856a9764fd5e9bba406f791378ce30d0b81ca67d5cd374928836ab3c41acb3191f982244e8bc9e",
            "bf50b983fe1dd73e3972534c4302001c88ffb225a3275b7272fcf3e30f9e0dab6ad92db263f810563c66c279ab3855d3",
            "2f86eacc81b49d86906466c03c73c8174cdf7ca091e0ba0c7fb822d02d1db0a87bd5fd3201466cb3f19b76dbf70f3b76",
            "5798172eeae25df33dd163242bdda5abe40629618f07f72c3837e9caa1f1a1a060ca0730a9c7a3f37ec2d108874a11a1",
            "a5e9c28e979c06206de9bb32c50b4e4493ca5ecc18b3963e44632b069a5c7d6a7603ea7152ce72f324e12f446cb25e3b",
            "807b04184c924373c6b6fef4f2d9dc0e045d364cd3cfa2c44c22f03444f5b909627622bc28831c2f69156119ae35d330",
            "5be0cbaa8d1457e98e436c5495a814c330de89298c2d12739f920a4ca56054462dfd08f4cfb03bca5f45a2754876eccc",
            "2420b82fc6192becf2f72b4965855268f650595139850f0bc94726fafa22461b5572aa33f7486eacd3b56d4b1e7802f1",
            "f015c56fab87717d5909cea39cfe755c503793837790f1c6b27a31442f7b16a7a4539d81e7273883302a9dac75c1d9fc",
            "9cafc5fc83b56bdf0dcd692148effd5ef49e8465bbf951ab36f01679ca347e848bb40622b728e413d09281a24c5a9156",
            "10484a553a8b6c46a03d9420cc97d86dc86a2012551bdeacbf77fc2ce701edae9f75ff01e48b068e4dee724c80031b70",
            "aa74323beb067134df28c465e027adac1806742bbb2f2b706d41857ad81b39946109ce38e79d6ace07918831b9dd34ab",
            "727548cce7c8105ca4c37c3147b18c5ac2d8d6583cce24d0341f29953c549e221550019309ed2530c2a660fae194053e",
            "5a4b39a3bc22cc3b8b3ac6831aa08857b9159a9ae87683dc0d496489c63a6ce0d2ce5ac27953901b63f602e55689c8c5",
            "6178903ce06a5ff1237867e22c2ccb69fba60577e4590bf6397640442676d499becf078a4e9104ba188584cb49dec785",
            "ac1835748972d637010edd033bd0edae6e0f688dedfa6cda5d7bf6fa4c87f45544293bf3a51948a683d6d7d3d29accc1",
            "493211d11abb751af60989b434484e7577a768764cb4be9d21e82d92efbd84ed459c76a54eb6c47572f6314a3d1e9123",
            "b0afec98df5ec28ac3d89609d8a9d2869f0c10eaa8a118e02d404fb77f50b421f03d4fa35f20b462969e7ca0da33f3e6",
            "7927e75f9a78bed12030a04dd06484ac9f7684a477174fa841044f5aeedc4084d2d448648d8169a0667ef2d50543aa9b",
            "8bb256e856eb78c0cb0402b2949e768b6fcd7b9da7a35bc6caf436a60246bbcbc61a89e63c2cc25391571bec83d64e87",
            "b6ace89c236e40639d1014fbf81a3b001ab66147359f9199cbf1ab93227592317d3783a54c61bbd1daf0f967e72e1d13",
            "5aee1cbced248530972bd3d684395479fcc64fcc3ed3f2220cae18ebfb5383132147b04447ee56e9bd572bf57da72d45",
            "e73c8292bbbc09dbb63f085d0e70571ff3b544d9faf7518d66b3e27538779b14ad6e4e9623836187092cda34ef1ab641",
            "0de554fb13909c99651f5c2e3e0ce089a68507384818f385248e05a47299779a6b359893a89b333481cd326950a29fa3",
            "5d5fc27cff22f292d1df690ee626c2b35b1fa37788088f215821ce5823414ce70c73af648f0321a766a88364da8eabe6",
            "2d1dd175ca022ff606581c5c7a1851c0bd155b2c45f76860a0ec5d36f5d41114c87685a596487ac920ef6d08a28c1db3",
            "ff6c778ca39b0f8ed38c79801f267ba0841d14c72d6e4566c14ab259f249a7d5dc9d9a98dfabd9922998872fba87d917",
            "5021af6a7d32c0486481b3e80b6d5705bee0bf738398e5c5bcad4f235a300080f43f6ab47f1efffcaad4fe71f813183d",
            "f233c3083cf13f41b9b49df040f990af2bfc8de6022cc1f836c006adeab4c0ce8ec1df7fd5c62cf1de347f41a040e12b",
            "f5d96a5ab02216d100ae1dccac45035bff341624a51396d88978b734e2ad4763e1289be167b9af403f917cce99f1020c",
            "17c9ca3d0ba104fef7438933fd704a09954a6023983c8a07d973853f06dfde4450cbaabee19d4713660d39bccb385a74",
            "6172fab6cdfe414a802e2eeac7c0080b9a91984440dc09417b88e2f640e6bea4b74ae9ef1b489085cf575a9e46e4bf66",
            "ffaf84e3cfb9ff35edf8bbd9b7868e3f3d75f58923609dc48068a59d5d263654e76405cf21078fc8c547cad3dc8e9cea",
            "17a0643f222b6d9e05ddeeb2c6c377449cd9360ea00436ba1ef9f84519873c3823ee7583733ab4deaa6062515fc5f091",
            "d7ac554ee9d6206e6cd3480bc4d633f3e3b9d805b379c12039e8383d6985fa17424bdb441f7b0b361fad679d73b2aa59",
            "a393adef50b386df12531005a7674353cf2295c70c9dd31556c9eb6d7de6f8ede48d21643aca9e263e0f010090fc3979",
            "1337271390ad81b5b485ad95e3bfddfdf358c676830825f884f68200e50a7a51a7ac4ef5b05845857ac19271891fd019",
            "0b898983107ad0e60d13e33c21935be681679a14edeba3efc3aad90dfd4d30c7ca8d8b9613a2a062d276e9db8e003051",
            "800b41a149f685d3f8370ed8bbfcd61fec21ccfeae13bc96575db0eb42dd1ccc97b52882361d9f99ba5fe2aba8ee9645",
            "d69c5ba17a00ab10f0de9a471586af875428f2b4914ea97a431be6f9ce5f26c2db30e6e85ae66c68993a98aab4aaf39c",
            "832cc40fd57dcc36cc8bcf1833d02861d4cb78a1b57b71c1a03ea96a76168d5a39dc0b5af4f944a9c4fb0fb36e6a15c2",
            "4fccc8ea2e131fb796de9b45ce85b58ad5d6aeae8fca0283a505b1d99b80a9d79aae4857110c69b122b3139b2911df0e",
            "65926286de9aa63886a48747b8308cba6cd1ded1d3a31884cf13c97467d81f60e6cef8c2b7674627bb4d7831e678f937",
            "9c88d3265cd1897b0a76f1b5b18072b3e5fd98a9dbca382390805f9440e67b0ea3bdaf2d95ee2b0e186a116d34f0b1a1",
            "d628b6a799bfef43b0b6ad6b985ae22eec232a194f4b3cce2b4c101c8a49acb1496a92adfa86c2785cd2b466df414d56",
            "670d46fb1adf22ff6eb1ab269198e39f46454f6f947fa43b06d6466fcc0d2c6c7f2e2def3bc19fb416b9f27785adae9f",
            "742126916842e1f37b49949b0347161b03975dcf419b180d3c1168f50229da15a321e02505a36254325a6d26faf5dfbb",
            "70fac56cbb391731e1fe88e930bce239d2394b5c2529b1d56746d15c1c0d5d8730a17c331f7a7b46f7aa79a6b24dbe4e",
            "d236fef514a8cb1d3f3aa919507a59fca2413633db0f57799ffcadeac663cfc36ff99af4ee5636fa89af5f9e1c192a00",
            "16ccd8ce481e6467ff94b0d9814cfed1c4f9df5369c01f3d4f52dd4cf87d27fe2b599bcb79f3f2b94e98cabcd4b0b218",
            "76c4d73c11369a31bb5f6daca960fd3d3081242fbdc0a0f8aefdf52864db832d38768419fdebc9c66209c71e25fc078f",
            "c85cea5e175361aa084bbe04bdae8813cbd61591aad57005e79254eab448852fcca6a062c37df70e92eebb068d81603e",
            "ae38811f884e822ac344a5e693f3b651e91c3799727ce91b7507fa78b86917da03ce1ceb8019b0443e3b7cca4cea828f",
            "ebf8bba61024fba45c7844d532b8d88051a69ef76fc7e163d13af574915814da75a9e288cdd0f64c4dcdf6fb094c1d67",
            "d0618e6d309122a34473a60a3198ad7dfcd5b677b0221496c43775985f1b74248f91c9a77d1eaf415e912d6f1a83291f",
            "eb3fd961ccf736e02ec8bfd75fa8bf6ef236e596cb77b4e0a6cc77de74b2bb5c3e3e6e6ee6c99ff668e777f7fe0afc86",
            "42c7c89f2f39b6c8a119f156eed5a3df63acc5776ce78b60c0f11cf89dae4ec6679089473e7b445795da4b573807893c",
            "51de15bceffd03962faf17ef5bde2f0b4e001f280ddd437fc0c142dc6da00713027ec76251ea3191dc278ab241fc34c4",
            "bfb0d8c0defebf24fe85d4d84adb9f2d0057e0de307e9ed2bdf4fe5d993818b21188f4961d922e09dcf3c2d4314463fe",
            "4bccba37ecb8cde73677245c12a124056ffc6c1f977c6f2ef6c4c30ebc8daab10c5996310f4109d588a997f10e29fe4b",
            "f1198cfc157459a0f8aa65c14ee4f773b55eb7a2a5d940942faf3f5375e5f9dba076fb7fcfad1f2f64235dfca0743b47",
            "95cb62db506e95b3d677a003817ccd3b56ce7988aa0d352d820331ffbec1fcd4f70d6d394f9b93de4b7030a9536648ab",
            "1aa11bc9ebe9ddddd872e6db9a597da88c8ffebb3399ae5f5605b7fde5a42a673e77c5f89b94598377eabd58c752aeaf",
            "89676196f8b29e0d15c56f3a78975b3ba271a943be6b0f3f536c1ef7a66bbf2f387b969d2cf2ccc98b97e031f5e54b39",
            "603b8fc31d392712b22963c19874ad3be3ce3d20ec7b27833e806403bf96cd9cd8bb448d66bd7dc033d8cfb3835ae52b",
            "40f65a74a276454486c72368d92e8848b7d5652f25f3ab5fa48237f3b15813425ebb02ec189e805df885aa48e1eefb3e",
            "745cbcf58404021becaaaf34f57efbeedb7639cd8457dd72448ce5a6d7cda9ef7c26d0efb69a862801bf6376f1a539fa",
            "ab27de663cf9677ac06ff2893f33aaffc82afe48ab5a6256ac9ee24fff0acc2da1e4b2d184b2720f8d912d1127ce9b95",
            "1b50fe2c7bb58c105aedfd09761f347a60ebb51bfed5f5c32a56499b552c2b7779e4a804f73a9920bd4c91535aecfd71",
            "fca8f83ac42e9a660e3325844795263f3f1d7f45104ab4574ea91c89c86d2a1bf11c97912efb9c1f216e04b72dcd0637",
            "ac83d3ab6aa2687952f454b9a4db21a2215ad32a6ac775f63ca241db0a856bdfba221a72079bcda23704b6e24c8710dc",
            "4847aa7b1b3c8114559dd33820c522ac1cc74b5ec64bfe198586b5f3a2704aab5994ccb917703695dda358f96ae75e00",
            "0fb0150538661a4c6d69c5f4bca04673baf6fa8939bf87cf794ab250aaf2792183a5c37d46cc23a6dac32c28a1030554",
            "1f94ec6465494644bd0f6ef58cecf223bfe8f958a3bb34e79cf3132c1113347b946c61ffebf98381d303feee8f0ecff1",
            "9bbf3d3c57f0ef0dd57fc5cf24f12fedd5cf059f3799fec203f47fb49c3e17fef7beb3bf820ae10fe1bfe9427baee779",
            "77d12f8423fd7dafd1732dcfbb567ec10df8f73d2ccfb53c6f42f88513d4dfb6243c57f0fc3ce317b6709f9d6e3c177d",
            "5e26fe8541fcdf178d9f6b785e88fd852282bf2fcb3ed7f2bcc4f60b9844bf2fb83dd7f07c77fa0bf8e4ffabbdea7385",
            "cf377cbff082e27fb5fd7baef0f926ea17b029ff375baae7fa9ee7e9bf0049fd9f59fb73e97f4f91ff8afe27e9ff4898",
            "9f8bff7b6ef25744d1fd45fc1f99ca73e9e7feff17cc18ff1fa381bc1414f44f41b4a7d7fc138376987edefd0f504b03",
            "041400000008006d996d5d12159c09d21d00008e280000110000006f72646572302d6d6f6465342e786c7378ed7a6554",
            "1cdbb66ee34ee3c13d78b040081034407008eeee344e706d3478707787c6dd5d8243022478706834480876c9dee7bebd",
            "3727778cf77edd1fef7c3daabaaac6986bf4eafae69cdf5c7329cac0c1e3011001c80000800ac000b706f28101000a10",
            "00005c0032acf66b3b5b67535b677d15777b53275d56371b10652e3c2c43360016f01ffc7f8df2f80979580eac8123ba",
            "8ede42f55aed638308c3353f0f4e696a41611ba4869c851719e0e38b7102dabe7210bae79742dfad6b9f8c8c718b1a88",
            "5c83bfdfd72fb118e130235dd24575ba8950e3bbf4955af4bc2a6942f31c137ee11744fafa915bfcc0e4b775fdb1aa1c",
            "3f91c6dfaa29a2b1a423c9f0a33db1c2b6244e5c446bc691596d0ce8264b1a97a8e2888e4424892637a49aa53353d0d6",
            "5a8e83671faad9bd335857e55c6b9f5212a3de2429f369de400a0153cf1095d70d92f418b47dcef972b5b08a75295182",
            "a225f20297aa232c95fdc7ec31db13b5ac7148d877ea67f1b6d59338f171fba99c8ad9404a3c3540f50b71e2702d8beb",
            "5eef01213289cc60c309cab1db587b24a449209a5af2131bf52105736fdc0c2b57617126d1d55d4f5ef1d85df6bea6e7",
            "7ede73a238078e15628073723881b96518f97afe97f9ec8e5a6f972ed203f30e797d3d646b575f2969a31cfbfa65754c",
            "896d6706528a610c8e4f5dc046477d8a65ed7050fea2e66f75b255321cd5f9e09cf7f77000c57fb86dbd2a3bcdd9c395",
            "ec8357a23db8adbea329c8898df5d7f93feefa1ffc4279ac9c422f3b7af091e44dc8d0e5e7a5182c51195711c9562a15",
            "1f3fa945607d641e2dc952273d29960a7100dc1bdf769f9b8f0b4baf22f786a82ebe719587e3e0f1d334b84282937e3a",
            "bc4362a2734c9470285febc376f97cfef922e58db43b834b7b34535953718eccf00a8dc4cc895835d604ab3a310a6e01",
            "693081408abfd6809ce6b2d18f0a98cfa2ca0790e0aa469eecbdb88934dfb9c4c25d0ca1505e59e8740afe909391f997",
            "0aaf2110d5f915b1aab205af659551a6231fb5bbbe09a8ebbb585a27eeadd733f9f421256a2e9c1f9c136f069f669ffb",
            "6ad5719d8c64b8f7f1a0f4960d1d0f83c3298a977feb248018a1cb5d380060070500003e3c7103b1b9da395a1bd9d959",
            "ffca6be56a3676792fb1ee7ca66f91fa5c9f90c88a2acee057b6481fe87d365b10609cfe8850b4869afcbc504af5c0b7",
            "9303456aa15cadf950fcfd7b81f5c1efa7a3c9330afba55c994d20b4b70a728e60b980b994a1f09f86563e290543cf23",
            "d5aaf9cc0b8aaf3e796d41ae76a1972bf2eaa282cb68271600e125f9b102a5483b4d9b860d7c6b05b15eeb1896c0f264",
            "708957d55709c7c44699715555ad1703346546e236b6585eeee8c11a2c611183d4fb4c7a6518a11b4f950d7c5a59525d",
            "84cb6a8c5799e8285a114d73ad00df3a18d0ce1703522512aeb681f531d00ef8c6b93eee599c0e9a1fd8e80102e9f0b9",
            "0179a5d471b9f868ef6027670776e0ee511a90dd71c1e786e851af7313371482c6d7c699787f66abe691cd5a00258ed9",
            "f2b0b3d526f029c097225caee669369919ac61b12bb6c5d6558557a34c1925d35cd2b3a5b46c7813455b848546fd2b2c",
            "431528adb9f02501ba9e202d737e3e7d299a6db4d9042b4431f2d5a2c2481b2463084a1fb274bfe9185eb418c42f44bb",
            "d327c149c0a50eafcb7e3df57e063ff6b5214c67973aa1e656edf5455849c1d4b0a395c4dd8d0cdae1c4a6d3a52b0730",
            "95ce2c6c8c4d6ce0e35584fa984e39da99f3c464583dfe47c2783a582109787e93bbb00e6895da1d9cbf66619e9188c6",
            "8b42bbbb4e20b61be5ad7fbeba34f1f1fdab0e3fb1a52a7a4992189a15b573de8dfeeb0b4ec79e219a9cf062fe0b2c32",
            "e530303b8a224e116e120e4f57e677b2b6181a21dc72e59167f85c5b9ff2632ee1a69e6a0371c0b7154397db9dab0748",
            "bbdbad8631de73997b85aedac7aaadf49199c4fbadb9e3fd5656d446013c137645657a6eea1b562826a906a84c69f542",
            "6e2755f0e1eff967afb35fd6992a7fefafed0d45ac889c76c0a141a541fd8697c720d7eb1788dd2030ad6dd2ce596257",
            "69a075cfb5d7416b2b46efafe43b4bead16cdbc9d7d3c7955f39f1619b77f7f94bd0e67ee6d226b7be3d4a120c525340",
            "fdad59ebca4d9e5c572ffa9e94d04f5c7af8055c08c6c72790bd2ce14d58afa77848c6af7290e9314d16d3e687f2a222",
            "a98917dea3ddd1e794c5418e88efef779abcc5af7a9f164274384bcc624da3edcffdf7a9acfd13d386fcda2b233174d9",
            "fa3093259563a9a5e90883d642a979e4c95c65bf3138ccb4577fbded47bace303f6fe555263b835d42fcd95c39e8dcda",
            "a2ad03226890f255eebe7b57bedeb174fbbc9fe4b863e30d5351e31b26892c36f3a4377b7ab4b576aebbd51fbb4d0991",
            "06f791e6e6baf56b37f6ed03cf7bfdb48e70c871f9ba31cfc807df33d4e53a47b4e1b50ef957e03573c922931fe1ea55",
            "d414456f2e14ce760ded9c330f37d4eede58dbfb5e7798d5e9d8884f2a077306def17f9ac33c80ea037e17151a8a09a5",
            "9f3cc48b2f0f918118800cf31015fecc9e7f8f0dff9d49611e3229ccff6e1cffdf47799c884c1f3b56d011f89e6a74a9",
            "50f71df617676058d5aad41c05d6b425492add53dc77d74718f80bac2ce7d3eb9eaf97bdd8dd5b6f4bb8a1680296e608",
            "2652a8e810c7a08604bfe8f591313876d46982e997e1b01eb4f3db73eaad3c5c9c7d6f54f6a33464ab69b331733c37d2",
            "552221f97def2cd6f39d3b12a8a4f54fe44b25d486782c5b367d3261ae34e9f8f1606f729a105ca0ad4d57f720f544ad",
            "b55ae5ac675fed0d11801fd08cb3cff01161026eb86b69c1b09da3559a71300a30ed095baa0b66a229ad7ddb561f58b4",
            "e3a00b57c1567dec869104278635c8b5bab812527999562507f0cfb5641a5bf471c156bbd56e40c1597fba98e79e285b",
            "b6aac99369f31d2797d3d9f7b3df6b45b78400bfa3d5616f5004cf83e8227e28a40801ff27d93859989a3a3bb1fdf1c5",
            "f18b59e9f10b610bec58c14768de48aa3c9ef82321503a8ea82493b8c982f8c44c5b69363cf7f4d0f9c40ff1b80547cb",
            "1f96dc1a51d00cbb116552afb628ee8de92f436a28f927254aad64e1268ae099723835210367d78de27692efabdb3eba",
            "e53011ced0107a647578c741139693b9fd9e4a168f26574020c603111aaf87238bac8a8f69d55b3608923cfa6521fee6",
            "724a9a559397471ac3a6e2daa2e59586c5af9260cda44132c520b42945fbf09f635ab35b62ccfddc4db1fb630d8728aa",
            "b3a11a1f45bf1ac315e4e39dac972aa5b2a141afb1ce9a299122d5435d9d9966a0fa2a4c05931f8b173d24cd205b2317",
            "6f407d4d3b9249b7fa41283f657be5f9e0b214d7f669111ac998bdc0e295814befd82f5b9cdb12be24df918f81df77ba",
            "119e6d855bccee2b5c03db77412d3be779cd42fded6b5dd177935cc3b47cd0166ae2538e0cf89357b92c72a957575486",
            "769902b82349fa390043186bf8f8a6131da67a07ffe8a6cd5615970fad1603df6be98836bcb4e7d9aceec937ac6cb039",
            "cf101b8b178828cc078908fd7143682b5786931d4666c4ca0444be579596345fe3817f14f47161b33a83ab4339e7b758",
            "cf711b5edb7c5bf97276a5797cffe55557c4abf62ba29b9f80faa7594c4dee3cddd98bfec7e17b3d44617bf13949eacc",
            "bdf3f29a6e852ca03344361591b8ca0a70b56ec24eb9f427159489b2baad7abbb62de61f43e7d26ac178313c15e80c53",
            "10a78e9009b5cfe3cce66b4a9a651f67ac22ef2985815e40f10de92f333d189297ea32f7b0bf23db81eaa7a17e240000",
            "9ffa57d5fe07d99c2d4c6d4cff3cffc133a8d692356d33b62f70c317eb56907f4960c5cc3cfa4a67a5e2807bff0bc625",
            "c3844561bd5a23b194b12732922106d2250eb28ea81ece69620896634c0029558cd040958ce3158d92d2a56ce3abd485",
            "731ffb2483eb02edef65d39565d2aa9f9556ee3738cd76d291a5dca7cfc19e1d21929c9bcf53595a966b879ba62f6254",
            "5ea6180ee04e9fab897964bdfab1ab70b82e80ff219932466183fcd0e383d2ae5e3d53117bc4a640662cbc4f9a13b631",
            "42a3d1db0c142f180f63350cef4ecd480eb957f1ab849f04532676e3b3f7cbd82a2aef1cbccf3d8fba3dc0d8019a384f",
            "cfbbe2dfe34232359410f9f9d1b8442b36fc509ca334e7029acd6e8c445913420813517d763c3aaebe97020fd54ff458",
            "fb3e617e895d9cbfd9746cee82b2703985f5ef9a29f9e235ec774b74c2b0d74792d49fdf4b64c1aceee26572ad2bcd12",
            "71a5faf08f563f5574d99d5cb244e026d19fd65ff846ea9c7a5b65bbcb932debeb7d4d3498315565b886c1d0cbd3702c",
            "200475b52258f6e3f47242d3a738d75a1bea2151ecb3706c88b46df87a6016b52cbc957ec0bf9661366387e1e4e60fb6",
            "98174bca341ac7ad679e35ba01877300b8620cc571f3d51853d619c51cd9547572a912187b9981052fd770826fd99129",
            "cff67c9aae7808063de193f2b9d90ca81b83b928688bc085a5fb02b0103d94d3bd05bb4a7d1fa7db53bb3975f539b67b",
            "30e60eb5520a938c2d057d78cda2abefcd675f9f4bb4b3a3cebbb5d64a9f838111cebbf3ac4eafdb4b5f9f2c0f0dab20",
            "46bddbadfdb915afdbbdc6ce9bb9ac0e478b23142fad60050b595a3eea273b8cd16b30918ca0bb0d954579231c432ba9",
            "5564677e956815139868b5e4667b6fa3b4e3663f967558419d8eaaa64c31a39262e3adc1b7201820afb9376b3aa25bf9",
            "3693a65b8d2cc1e1707f4d8467d849b3f5218c3c0b52742263befa9e126f68a8f8f7e435b13504a950d19f3c3e12a1eb",
            "bb82799f6af6a9325059af74722618f68cf7b8860333b5fc55900f66086f6a1cb1e1d5b51b78bda27ea637677a488edb",
            "d0b2a08d926ed0076a9a798526a10eb226e092d6d67403abc2746815d48cb40622f36115d890d72db868945a4c3291dd",
            "966367288c2541f28d1435aa209aa05c73d60ca4e4ef9c8e431388d851caf4e6a962a6d0a2537451bc3ce30a9de3d7b6",
            "9558de409fae617957a1a23ec7cb34dceee792f6a6d2e27e3bfa893d6b2321810bfb0efd5bc94d35b20a7d6a2e07287b",
            "1151778e8641173054e8faab44543b5f8a3011bd33b55eb0490a8551964171d50d047a5ccaf7ae0325c1b2fab519d4c0",
            "e23ad18b40961f9882216524efcd83e208c37553a3bbf6d57632513522a2b6bd7a0acd4d57b18308b0bef55bc6d3e338",
            "c7833deb64f61d02215f0b8e0af8f5d47b4e5716ca393492d86c5b59458217b51071aeb517ccc029dc22d8ecb8b253fe",
            "cc4be0d0cd3b51b139559a06e78f6307e2da4d30c434652bef2d4f0bb7b691d1bad97c9297b7f3de09d8f61fcd1c413f",
            "325ff3b4a4eb36dd772f9b85260d73f83bd664341d6aa5b5ce0e5c77dad9b840365b83b72a729aad2d742a0d280a7218",
            "3d082decc220e4b527b8478b9729ca03f29fc23ec05936f8d0b6cd8e0eb05b5ce48f0e6c9851b047ae2a0901650dd608",
            "604988689b7e12a94e49052acf53a08eb26e896e953ec129d0e969f840f4c2b37423d8b6f6aa2e0278c5b270c8752de3",
            "6326fa5e95e390c7609d912404c294909306f152fcaeadcc037a6256f64e4e78643958cc92334d6c4c24b79edc0b0ea5",
            "dde9c250c39994e5abd6ed5755ed428015496597270fb8f1bc62aed83cc2e00da6c759ca85f2d65ebd6bbd936071912b",
            "515c020da081a05160694ac75530d3a785e5d372cf0c01ca7a3c4b39086740bdb03b0ebb0d637b4a73d9f5bec3af1552",
            "c91f24f3cce32bdefe67c5498b7aa1e0ca7738cd1f483494c6f158449e540dbcd670264ffd56b73f66b842c8ba8e077e",
            "b3e7ca364a600f78ae4aad7d7214aba544f5196b84a1b669982fc5d7a834543862405741e89e8e54c3285652a2024ee6",
            "72509af513dae1a004669bad67cae8b649e9471fb6480b19f3c6986fb5266b379f433bf935644e9f365dd3a6ea662b37",
            "e02970046ae4e15b1c8d1707ae07503a9d47f4b8cfd3c5e19fb8e6716e7e59e1db603ae530b30cfc1eab34be45b48817",
            "8f41e6d4582db2cbea6bc6bed380a2f0a4a5a7c3f58bb71364369179cfb88821554d6bca7426434e482f81ad9aa9b4b8",
            "9f696489c7c9d63c034aa99861faedd6daf7bd16478ab8dcf3e5f15e3ecbedbc17f5c079ecf85205a790a517e23d96ef",
            "5d9fdd52ba3eb77a17a4c19348773cb6e67e054acbeaf0dbe00854ee7b3a22dfaa09eb5a582f1b1d7472c098053995a1",
            "ab158bbbd2e129d064f295de0a1d676dd5dbd923fddebbde466280922d3781c60ca2d9815e804474a18e53a06edbc0da",
            "1717129d566876b4b383e52486b5f53bb910d639fabd16759ee7d621a60a36d355d5783a36dad851f708c1b3bcfca363",
            "8e6c1b293209ae5fed8312c12eed59d7917ade582f10f13b970855f7ea5bf42d793ceaa52ba78a11f2245681d76bdbc2",
            "0ae20e93431d3e84bd9658da9b2bc2c9279023ca6fe7fe73ed8cd1658b5edb8adccb3c2b22f039fbfca1efef86bae1de",
            "c042ecf13ef5eabd2c460f16a95502868c49b7aaeeab221a3a115e645bf00cb64f0838db7689f8be248d26312ce005e8",
            "2c1d7133b5fbd8285a36bcb9b5278dd87b4dd1e57ed51cd84898e1e4a4dd55b076bedaf95b891933a295d0fca007c411",
            "01008c3fb3be93b33bc8d4e957becf5719b35a10c6ea78f21373643921016fa43e325f71101b5e898136486f6efaa35e",
            "8d9059bdfc66a25eac2bc31986c77ec1e4e41cc7062f1ff3d51bba8267c3e2fcfcfc3fb7d2aecc78044b0a8479465574",
            "cae091de084ac86eb6ac5a25aebaaaa7f07415d7a8e0ab58f4e3b419e10b4ae9778432b7e2401dc793c50139a5f01310",
            "dd08e49253deb83abe69b4c2d5fc4f124a68dacf693ef7b06ba91cd2334aee5c7b5333aa3789130edff80ad372cf8d04",
            "25593767272a34afba7c3630a4fd3e75256bf9ba08b67856247aab18f3279d58fd173acbb6e853c94c12dddd697218d6",
            "29d361a78be4c3187494b14b49bf1bae7c4a6a1f37454f9bccd04f34c14bf6f414f2eb4accda393a6c09081cfec5d6dd",
            "3b5abb4026a737e35f240a1cceb19714613201715ce3f593889a31987c736e7809e5916e4d9dfa2b771a2b773a2b4b10",
            "bbced6e3586956b503871dd2be826516fc74272a1731d1245c524bf4304d33154492baea4a35b4a1628326ad64ba438a",
            "a9e7afd54508dff950caf570655a2977b11f6957bd2ba8362840c80d9ed8b2adf7805929692080575807cfa4a8a93429",
            "69c564c06c6ec886caf01c1f8dc0415bc4ac189f5b4e84c3f77e76909b4f551c1a009d7c8dec3dcaad9f9f08167fe71e",
            "0a0c25fb8a94efbafb21ee653770a543f0e6276c6f8dbff0ccd3278965425888a3b9cb8d5daea81aae0ba95bd01a617a",
            "54b5e9822af5d65ef8a3039ef25519a2c3dce553121b92d0877bb4229549575e0d1b362642368da9e963785f3e924e62",
            "f9f19a49921febfe73fc661f78220670d9f632a2e8e3403b2177bbdbe319be03917699c55cc1ad5ca5f06d6866cad0c2",
            "be9176b6ee0c3999ce94b7d94b9b82e8d777fb0cc7edd079609b5e16669bc8fdcbd31a0664fadbe51f2c1f73312f77de",
            "b3052880b3bd8d49af2c4c9a4dc0651035eda0163d2c94065f46302881f2fa52e816e3771ce7e9bb14aa78b83a7938f0",
            "fec5710b434753136567474b5bf33fa8ae332e87dac38e85b0dabe4ad603e55615138bc5c3d669fe04d7e752e76fa6c4",
            "afec6b6b2726667fd5e10e37111f3d5f4a72418aa3992c87a628a95a0aa7cb420a6c1d46c50b2ffadaf87eb20797489a",
            "879cc8eced55d2bc28cd34a47eabb47bc73d952fdf8bb748d063b8bad3e2d0e38c9b2714475b802dac204a10ab980983",
            "9680201b11750ea671058a8e64de77708ba29dd74f741b0cb4fdd67a51e9be22342c702bb93d7c88f0bbf9dd3c796109",
            "86010058e10000ea3fe767eae66cea686b0892b5b4b576fac7dd9f0563ec987c2f3bfae09164dbeb787e19752af8cf5a",
            "f0a9317458a47e48434a76e192df988625dc3dd3a101910ec94d755164f746f70753367618dafe4062ddb19000f0737b",
            "54ed832d6e5f5fcfa6a62d16b15d2dcb4991b14ecae385ad1b10bc37fc51befbf46100765f502a9b1194890aace59760",
            "e847f6326c74529b6877679bda99dc8484003045fc96a03fa14a89bf834e452d03b67ada97a5152da821eead42fdd409",
            "73edcbdad2d451320f5e6e1590509b89b7ddc7b54485a83ca4723572b8f16f7e7eb0e6cf21ab4de3240e4e8a960377d1",
            "aa9808efcd251b150954b46b674ec91df24c3715b08d74a1849fa7c1e1d6bc02843808ba1a3bacf65bcc4bb38b0a5fcf",
            "e4595b62e914587b5a790cb7d311b946a0d7c38ccd1612df26ca5ff5985bfa23f3a741f24bae9859db612fbe5116dca3",
            "0a8a3e69a47280271a61d0a7b2ad997ebfc98fbb97333f85795bc2119d15812753086b1ceb9d845916ac06a255fb247e",
            "3a136d3bcdf31ad76c651c5834b5d12e5f9bbd7397cab3f2db287ceaf30e39e1e1aae4e1c0faf30d1a1b828c5f5b185a",
            "dafe7a5fafa244506139b0e08fc0bec84eaee67301a1b45f0276d0ce60a8f8538539353945ce2e07c2938273735ace6f",
            "c6a5da8ba5594006c4128a7d1d2c1fb2100f1bf70f91a553d7fb0b7a4c565fc3d62ed9d6b90e67058497b7f8f70b5bd8",
            "08508501e333d04d9f1e8ee57d826a0dfe14d3646e16877d3b8272b9eed6b5886050affeb9cb1416ff1a497137c982fb",
            "ee0df4c9efa6a0937c2a1bf940c2a98764820d408631b1335674b4b3776233b67334fd57cbf73f2b5f803f7a2832bdec",
            "a40347e0959c8de582a564c95da662683ee5cbb76d411b472c0d9652b90f4a2a8712794ff555a8db3102ded9069a4f0a",
            "e8bb75b28b93fc19ed72ec2b059f57f7f427e62eeb435f61a87a03d563b4ed6310d2e95836ce51b54a540f2c4da4c45f",
            "d2bcedc901a7524ec9b2f3aecd7cca7beff7f4433ea50af234272dab40aada64a1524948757e1a16ae263ba66d5fa458",
            "9f580aa4783d7ef29912928e733a236900750403f879764654a845a88973a06ca3783c1c613f9608ba8e149440e76653",
            "95d2e86a3f4887c86920c7aad0acd070d1e1451e9cc3c672c2dad0bc05176aa41279ef86ebc9d76f3d8e82a564a7e5cc",
            "a80cc744e7c5805c1cec9d5987fb9bc24b8f4fde02de9d5e3307cc4bbc7307f44b1d13b307684bee5319d64bbc3310dc",
            "3257d7495f37e8c1d1aedafeed45f9b5d7ce3a87869a9758be16975d1d366c9acd1ddc3dfdccc6c43e313acc2827825b",
            "408b60c2e7e9452f4818385739bd4677ce45867bee24bd469db5c2fc43e7d2042a11695fefaa2f9d0f2c2cc5fd8fc06e",
            "24c0ef385c34bf859cfd405032b85f6ef8370e1bdadbff87c27f213db6daf6d77ae325c01bc9df5396f719b82b9b104f",
            "3e1c1dc5a84bde2598ef8b18632ce504788943efe70ba2da0d9b1b6a5ea9509aa28c965bd7fbddb9419afecdc2064d38",
            "0ea223d417e49b2c7045de968b350697e159454f58648a754aca5d15ab9ed1ba3076a6876a9ec150134b128937a954e1",
            "1b25db679ea15a254898f03b32515807a90c3cbd5435e5d2716fafe8a6121af9ee844902af2eea4e8ce95fae90cef5a3",
            "33a760ee99be8706080e46e64400deba4e94afa08d9befe706f3a80f884ba33ba490b28577e71d575342f36aafd2687d",
            "0d7b818126639af071747013e267569b42595925ae9796c279d19a9c9d9636e62c9e2a285ae9606e26e2cd98aaafd51d",
            "3b751828798b479997ad73b733dfef760bf05b85c03306762e8e3fc614d3ec1c1731625e9283e03c2839a3de3e417c96",
            "df6a55487d17e0f58af688d0e390dcfc70eeed6b976d0ba22cc2d6f59915c14d056fb2f8855807459e815be94bc3b54a",
            "bc9e6babe834a30e256ce996c8de3d2086839c766403464a45bcb6e895e0f7c629c1c090bbfa91932b86fbd41932a007",
            "d8b7cdfa0af177ac5edf60cee47ee06cd643647e06f88d3cf8b355f16f22e18f7e455bec8af5832e1a5c41baa6905e36",
            "9e56a90b551d5814b6670fee38df94adb4314b0ba4b5ba3d7a3280881bd87ee5467a7975b7ba359273148bcb5719c12a",
            "c9f78d56e4995c81b372ccd245c216ed098c3342c694460c9a478dce2bc151a1aa00ac68afb272056af16b701dd53c91",
            "b15ec37c000551389fadd5ba07b28333462d25709212375d21f97a84329cd9aca602f5db6c02e5f08551dee93d48794e",
            "fadbc721e368bec5188397f0cec42d91d321aec8362601d6ec78f96a6412db5077b543aff23ce7a4f522a6ac725d113f",
            "47635ae74d024fed5491a919f21d222f85a77dc52af61c52ee89895be5d53f97a58dd9ced459331c99aa82f996144e99",
            "48a2fbb9469a879e63850e15317c50b68b92a7268b0e1e07760e2ccbe2e775deccd3415ddeec7af3b5bce67339570b67",
            "03fcf7ff0f03fb0cf03f6d887a8cdf6c8f7a3cc03fb766fc1d8d0fa1ea6f1b351e1b3e6e57ff8545c47f6b5e3f36fe67",
            "57ebef3845fdc3f87fe8713d1ee7711be32f4861fccf4d8dc7a33c5e9ffe0b0dc0dfae563f1ee071a9fb17ca9f3c2a7c",
            "1f9b3eae20fe420cf1efeb89c7233cd6e87fc19ee4ff46b13f1eefb162fc0b9c64ffae1f1f5bff53acfd1de70fd6ff26",
            "dd1e9bff334ffe1dcd947f33ff57d67c6cfd381efd85689affc7e8a4288380f8cb10e3e163f9f063804f7fddfd17504b",
            "03041400000008006d996d5d9d28e5b5cb1f0000f12a0000110000006f72646572302d6d6f6465352e786c7378ed7a65",
            "545ccbb66ee3ee924642700f122058208100c1dddd5d1bd7860e104820b8bb13dcdddd824b02341edc35f8257bdffbf6",
            "de9c9c31defb757fbcf3f558abab568f3947d7aaafe69c5573ca4bc1c1e3031001c8000080027014e3e9f31e0600788e",
            "0800e0019061b5dedad93a99d83ae929bbdb9b38eab0b8d9589367c2c3d2a7036001ffc1ffd7e88afe6a09cb8e1f7011",
            "d0d2196ee390c922c9f099c5ba91b158fe1a600172862450af8915da2df21f21314c0b916e9a0ad3d66f78d5539bb169",
            "aad16a09bdd9b36ef7708f33d1cf519a1d463be7f5de7a55869e55260934cbd02779f3927861013d97042b5ea1ba1b12",
            "c37ef56e5821561e652805498a89e309238e05f1fa3c5a3db6e3d2aded45aa9419d5b76fc1d5882498a9fab4b1b4a623",
            "5657cd3428fae126ab7b3db385fa1d9e59394122036aaab7a0ad5d525081b467d924d901f2fba328ccb38145c4c60657",
            "4b9951d48a60e4231965fff699407063d94f906d44da135ea12d538ee742bcf545f970bc57006103146ca6a83cf652ff",
            "1acf7d8e3b0b8e63ce534dc42f2371feac6b9bf77848a3dd28aaf16a364ffac2cc6c5cd7d51b337823c8e00e4f9b5329",
            "fb398186f3c0b69a03635c1b108b28e00c07ae5b6c07d03b8b41aa08e465d9baad9af77fbfcf0bf6988c5a7aa5d88450",
            "47aa9fa6e7ecdc1b85ed9fbf1cae9f17b9b4243ee67616bc936443908fb7f02ae0f6100900b8bf8703c8ff63ed56abb0",
            "519d3eb4a41f9626dac3dad50399583bb2b2fcbaff67cdfe07bf50142923d7c9861e78207ef3a1ef620a1a8e2d2ce52a",
            "24de48a1ece32b318f551d9a454d026da523c556267e0ff70edcec733330077d15badd4771bec259f409179f9faac6b5",
            "3830eecac119899116142be650b4dc85e3327536759ef04ed29ddea5398cb1b02e3f43aa7f914a6ce248a41c7b84458d",
            "18052f8734905020c14fb3474663c1f06709cc94b0d25e7160592d77fa76d44812782636770be37510aff4ee7802419f",
            "a3a1d9f712af3e6b8ab34b621525735e8b32c354101fa5bb9eb175db8948522bdeadd773d9e43e454a4edc9f2f46def5",
            "d2a49f8135ab388f0653dcbbb8513a0bfb0efb219f9ee52fc0fc6e91f406a78101f00000352a0080f5f0c4cd9ad5d50e",
            "6465686767f5cbb915a9dad8adf060dff98cdf7ea05f6c7658400e9654739f52b1bc2888b539a0ce8e91903ee20a76fc",
            "0ca73c75774182045b5524ab7d814d5d2ee0733334c481c7226857a0a8a530f64257cee9046f9a4b73a349ff6c42e49e",
            "efeba8fa8ca6c43624ba2917ef36a0e66a63500f5cafee78f46ae34603ceb72165538a21a98d6dd64b8269dc6705a3d8",
            "c0a45d54941abd3b6aec252432534791b170bc8c8d8c88c0a0ba4514b98b17e855aecf2be3b6a05215ab852d2445e9e7",
            "d3682cc39fae72c96b4f8033a2f31ea4cb056037a3896215c82f5171b73f5376145aa7a96b66d7b1a15ae287a4052f1b",
            "62d0b58ff72b100c22d85d0afbc09a1b495d7581f148722f69e1932680c91fd6809092f881e5d418a077f8947918179f",
            "6fc2cfbc31e970d5d21800e44288d375e4eb382b2b61513696d69665b6454bb8a449c1647564d58be7f1d9dd2217fa7d",
            "33a75ffd3713f38d688a66bf55bf82d6cd0f5eaa9a23e88497442b611f035e8f3db7fe0e6492c216539764b4a4b418b2",
            "42e19bbf3cab4cdc295fa5410ab0723d3df88c821a807005abb72c79ec3b34a74519a74937104c865459fabcee351e2b",
            "54cf9af0a58c307582ce49727425437122d1225b55225e15ad4d00083627b82d2c22d8c6d1855a0fa4c9742fd446ada6",
            "c04eceacb53db77141018b25c21674aa8a8a915a1a69253e29a02134fb86b4bba4a0949c75d5aed462adfb5b95581fb1",
            "2705488c45cc67f3273d86c01c31afd12eddb9eba1700824394a96c555513e1e0fead794c60c0dff5a66b8926ed4bd7c",
            "a3d6744dee0c76259a6e61925399f0d4778292896405ceae474c21e8c939196b2e7d1d9c5b6ed0f3320ed086c47bedfa",
            "f6bbec730fa4c6b67d7312a80e5b8b2cde8c1e3210abd76748c5bbe522e625eb6909f0eaadc67e754d5f51817fa9e0df",
            "cc53356fbd7b7e134205f1ae44c8bcd529bc864dc03d12a171729e0a52295a06e2e47374a918306285151fb8d3bdf9dc",
            "3a736a8e74da728b593872678f0d3d8d515818432038a4ccb0ffd93f6f5009b5bb1ba31db5c7acceb71578b7d8155b5a",
            "d4feb321e6f6e980af90b134d2076c4ca4e54997088e4060e8d339e1f228f52f3b3912c442c6810b27af9d765a88df2f",
            "23e4abb304872b091d9bb9d94076702c84c6bd9931963ca55f827a716afc32a3128840d2cfb290153fcd5357ea8c48af",
            "4c7ec5348eb9fe84753e503e76902be232fa4e196dabd7afc1b4f1fbed9a217f79545fd236d63e2ec105741fccd18474",
            "20e591eb706e9f6b70f3c3d4c33177cf99eff63c3ec58a744144c2adc5dbcdf5c91dafcc3df0339b0f70f402454ebab7",
            "d7859872eef3d3fcb9a1b7abe88d7d2eafe85c52e8b9ca5dc5ddea87d59e0dd6fe00bea75c5d6de5682cc93f2e6d7b0f",
            "050b14e7806ae16fcde9df6e42df7aef8db1dcffd658ecc87620533efc62fa60308801c8300fc6e24fa7fa7793f13f0e",
            "16e6c1c1c2fc2fdbf7ff753445b74b76b161071e400ee4d7f7738b1b3e5a372089aa1f66efc0f93bcdfae74af48b405b",
            "9a333eb8a04935e3b3fa2548dcdd1b45f6db822d658e4353ab2a492b64505134b7246dac30d37a6792de89d3236ba1ae",
            "32833623f4f5b4f5f616c682b0b122c69fd22874e32211085f740125028ad34377467b247f783109a37b7fe5a32051a0",
            "1a9c701abca947c4e7f8ba260f3afc704c1869b428bad15aada36211cc2aeb536dde89801698e7bb807a926b83e71c50",
            "ca6ecf33326e2d5713fcb38d254aa0488236206244d29cdf20f40b4d90974cd0b30d12c520030850a92b50b2053eea96",
            "6368e5cc7ac69f82e655ecb3f6defaf42ffc95a78a66fb6a0128cdc1463c2d7ac42b34d78d5c7b6f496d93e73f8463ec",
            "b83db33ee7c76f05fc8e61f7dcd22c750fac898503008080ffe38e1ccd4d4c9c1c59fff862ff45b2e428994f73bfde6b",
            "c01d5246e373f05a27bb436cfceeb800fec0f879499f9156d73850e23b75f4ddc500e611210805eda579d118c781b96e",
            "69da7001d5817e9188ce284f083dccf3b2c80686cdd2345621b160676db5a6b71808cdbe397249a57ad3749c4873b03c",
            "2b28b38455ce252344959908dd364102e6ad4a590af397a39c43e7886e5c648a4fa532cddbf1e3127551eaddfdfd9c9f",
            "23ecca931298bb060037eb4eae6c983c3db3188e90651d1c79cefdde4c7976e9ba8ceb88d1d348ae95868a755ea41c5c",
            "938ed4919bada675b56e150ffdbc2d898cec48ef9e0d481d2a493c98d32b275835a4845edc63d3a6b2c31d80804b488c",
            "97c9d96e82e59158ea5712df0c21183326c75945a85766bcac83a284cfdfbb3784bcd7b8bc21359eb3c6319688352dac",
            "d2a6f971ce0a2d62daa27be5248a3e8900ed89b699ce72f791e2f97e3f6d7cb84276c466dd74453c163eb00fd21363bd",
            "63cb86243ee3bf08ccb3234f9a765f355b06ed8b622a0d170e62161d9e1f67dbbf8c813af40ad46e60786de446cc4662",
            "6eb35b4804d74332474b9f9e778ef9d5bddb9087664e5f3fbb49ce1a5d140fddb46d4f77bdca7a79ffdbd9f6ff2edcc0",
            "f310841323fcdbd97ef1c76c47eb58cdb1e1071e205d6325da7a100c82c61bd83f4bcd5217859497972695a3890aca8a",
            "580f8a0fbc23ba3ee0af78526b45ef2f0c1beb3c7cd2d25bba7bba2e1f638cd4c3a823478f252c991ecba7766629dc0a",
            "d6422f6128979daca777c6c0bb02c42e18db2d667bdabc74c15842c6b5e4cfb0da48f467679a914721d944b49c67a8fa",
            "e2df99a3c13e07c4769f538e6146ed6452455199f81c1eef086ca08ff64f889d43b59c11428feec80037aa860a8c33ba",
            "6396e09a56da202aaa6150c61628ef6704282af6aff9a85851c5d20c7b2f5151f4b72fb176dc1fd24f2cbc562888e69c",
            "ddee4010374dd99838eab51ef2df95c83b490b58a632d5cff3f20dcbdf748efd6932f291cca053ada3c4231eac4dfbbd",
            "c04b18dcabf669f0a7ad173ff044811378b473a9dfd4f6c3be0179d5ada12ea93beb35a7946300caf597203f43e1bc57",
            "09580c83e3e73737f265b57225c17883a5f210275fe2645cf0ab384d3baf658105546f26fe2d9daade934a0a5e7c43ad",
            "06568f7bb249cb3d6cd0d3aeb8cd39a4b4352c8e325c39d24888829d11ed43b8997501dff956650ef99ef1702da78739",
            "a6c5e963755005cf3a0b9a52cd5bbbe3c585bd0981bbfb8564a377113a376bad826f56fbf288591664caf30fc9bdb2ae",
            "4081b91e66f13fbe0f97d5b8120a381256fb5ca231028fbf33aacf15a2a77d41d78ad9ac977ca27bd145520e09eb2852",
            "b69140114d92940ede8a44836b726938d2b8448e9be5d579aa66758308638a8d689e393afbf4c99defa95ef52eeb6fd9",
            "b6a732d9d7fdb05324a0fc7596f307db9ccc4d6c4cfebcff61567635a156d4f53860ac3530f6ad203f5460d1d42cec52",
            "7bb1648f6be73bc605fd88796eb56a2db184912732920106d2052eb2b6b02eee71ec076c50f87b528af0d73d6552a04b",
            "2a45c50be9da578973673ef671fad7395a2785e3a585922a538a8bf76b2f4c37939125dcc7cf209e2d1fc45ffce04864",
            "6e58a8ecaf1b3f0f57e64930e8c11b3f5315f1487bf5734b6e7f558020229e3c5c6e8d6cdf2342714bb79a318f2de487",
            "406a24bc4f92238e1142ada1420a8a178c87912a8677ab4628bbccabe825e0a460c2c85674fa4e216b49e99d83f799e7",
            "41bb0704e7bd062ecd595b74305e71aaba22223f3f1aa770c99a2f8ad3678d99f7f5a63786c22c311f80b1a83e9b1e2d",
            "97270558fb6a47ba2c5d9398df23e7bfddfc00d5b7ed32733a7eecde325504e3d7ecb48bb5c2b0558792549fdd8ba5c1",
            "2c6de1a772ae2a4e137126faf00f95d3c8bb6c8d422d10b848f4c6f5e656489d126fcb6cb7b8d3a5c1ded744bd296365",
            "06cb18f49ddc358702af775d2d09177c5f7839a2e93d3bd35ceeeb2091ef3207d584dad6ccee997e5e78b39ebcc7bf9c",
            "623a6187e1e8e60731ff2612976a388c57cd346d7803f9c40ee00c3710c5cb56654858651001b1aa686752c430743261",
            "e5f02ce306deb221939f6efbd45d7213f67ac2c76573b1ea53d606723ea3ce83e416ec08c016eba21c6fcfd995eaf938",
            "de1edbcda8a9cdb0de43303729151318a56c9fd17daa987705df4c817d2ed04e0f5aef961b4b7df67a065fdc9da5b57a",
            "dd5e807dd23cd42d0318746fd7776616bd6eb76b5b6f66d25a40e607285e9a8172e6d2d47c944f3619c296614219acef",
            "d694e7650d710d2c2596909df895c3948d61c254e3ebedbd0d930eeb7d99576105b55bcaea52450cbfe41badf72a58c3",
            "60f19a79b32423ba156d306ab8554813eef7775784787e3caab7da879165460a8b65c856db56e40d0a123d895f165946",
            "900812bee2f6110b5add12cc9aacd8a1484165b9d4ce18a1df36dae6ecf74fd5f45346de9b00de548070e0d5b46a78bd",
            "3e5f25d7a77a880fdb5033a30d91aed1f96b987a05c5a1f6b2c4e0915656b463957dd4a6965335d4ec09cd869563455e",
            "35e7a4526c304e45765b889c7866246e2d5bfbac42c59a2a20d38c250529fee405a86f0411e7b3129d59a288c96ede31",
            "ba307e965189f6e15bdb526c6f2c9fb67e59d7d7795da08b24bc760e717b134951df4dbdd88ee5c10ffe733b0eddebf1",
            "7515d2725daa2e7b28db219fef400601e73014e87a4b44149bdff33011bd53355fb28abffe485eb88ba7a62fd0e152b4",
            "7ded2f0e91d6ab4ca1c4caaf123ef767fe8929f8a19024d82c200af8492731ac6d47753315553de4f3865747ae99c912",
            "4e0021f64ab745341dae5334c4b34a6ac7c1bf7836e720875f57ade37871ae885d3d8ed5b6914528705e1311f75a6bce",
            "1492c02584c386273de6c7048504fdb813169951a1aa711af8ba27aa5507434c55b8186c719cbbbe818cd6ceea13bfb0",
            "91e52c60db7d3071b03bc074cddd90ac5377dfbe601a14d7cfee07aa48a9dbd74c6a9ceeb96eb5b37129fed118b85e92",
            "516f65ae5daaff2c2783c103686ef7b198acf208ef60fe2241a94776f263049c458d0f75d3f4500f9bf979f650cf9ae9",
            "33b6d025c5d758d2facb84b02444d47557442a6312fe4adf9ea10eb1ac0baf173cc1cdd1eea889207ae959b016685b79",
            "59158275c93cb7cf792de5632a1cacc2becfadbfca40f2a198312623a9d84bfe444b89dbfa8969a1b3cc9bc18540118b",
            "1749225f8532abc9bce0509a1dcf0dd49d489967356f6755b472019624a56d9edc90dab392997cb310fd77981ea709e7",
            "4aebdbd5aed58e82f979ae44513154801ac25a01e898b6ab60aa4f03f3e442c70421ca6a347391356e8f5a6e7b144e13",
            "c6c698c682eb7d8b6f6371297f80d4738f59fc9d29f951f3ead781a5ceb8f51124ea8ac3f8cc424fca7adeaa3b9125ae",
            "54ed7c355804b2ace243de6dbbb20e11da03385428b58e0e22351529a6b007e92bebfaf912c08605416f427a74e45edf",
            "d392aa1b468a8b95c0495df44ab24ca2edf78a6136d97a260c6d18170cf8b0869a4b99d586af541a2fdf4c05b5f2ab4b",
            "1dd3d45d5327eaa42bd5e0cbb1fbab6711981f0ce7fbafbe27773c0be970ff461b4570e49af5e2c7f745be35c6637653",
            "0bff9348c5e175a279fc688ca78eb5e5425b2c6053b6cd1a14b9270d1d2daedfbd1d8ba76399b68df2e8135535c74c26",
            "52645eebc6b0963316e477330e42b91d6dcd5276c9e5534c566eadc0c19aec09a2321c0bc39d7c161b592fabb1bee144",
            "17c8397e80be14edb008767d7e4beeca61e91ca0ce1d4b7bf875d9fdd23a29adc5778ddd5fa98b6650b65103d635b75a",
            "3a2ce0688f21adf8588ab65224ea529b3b4783112cb91e34ccd2a8bbb94d7ad2b9da44a28f922e3382c6644db5b97b6e",
            "2da4b30b1ab36eb7f5af7c792ed66a8966473ddd5b44625059bd9959cc3243b7dda0c6cd61f5c144ce66bcac1c5fdb46",
            "0be7f33d42e0342fffd05710eb5a82548cebac7d402cc4a539ed3a54d71bfb2522412b14a8b25ddda067c1ed512d593a",
            "968f9025b68475bdbcf1464ed461b4afc507d86981adf563f14dfc51f101f9ca99df4c334358e1bcd7863cd702f7a210",
            "7cc60e7f50f05d5f3bdc3bd8627bfcc94e5d9e7cf440a14a45ac0f5f251b557654100d1c81e7e9e6dcbdcd23024eb66d",
            "42601ed23012831c5e8036f4808bb1d9c746dea2e6ddad3d69c8f6db676dee97f5feb5c0144747adb69ce5b3a5dfef28",
            "c2073563ea1f624c51440000e34fafefe8e46e6de2f8cbdf672b7fb59c7b83ddf2e40a73702126067fb03a345bbe1707",
            "5e919e3a4077667c40b7e2b569b5ec8f58dd4857fa530c8f9d9cd1d119f6355e3ea6cb77b439cffb45f9f9f9afd6932e",
            "4db905bfe4bce11e52d62e84477a272826fda361c93276c9552d81bb2dbf429940d9bc1bb7c990405042af2588a91177",
            "17341c2f0ac828801f29d60941fe72cc1b55c5378e96bb943d29a688a6c54135d5c1a6a9bc4fc720be79ed4dc9a05627",
            "0aecbf01bfa1e69a190c88b3aa4f8f95ab5f7299d237a03e19bb94b6789b079b3f2d14b69e8f79452b52fd9dd6a229ec",
            "583c9544676b9c0c8665cca4dff13c7e3f1c1de5eb85b8ef0d673639a58f9bbca74d6ad0245520d49eee99ecaa229356",
            "86366b0c02bb5fbe55fba6e61616a3e3bbe1ef62390e6738507998544014e770f528a2463826df8c1b7e4c51a85b5dab",
            "dee29dfae29df622b4d8aeb5f13052924575cf6193b42b67819920d991c24544380e8fd402fda386a93222495579a92a",
            "5a5fbe7e9d663cedfeb3318eb76a4240671f72990ece544ba536b603ad32e79c72fd1c84ccc09175db6a0f98c52f3584",
            "f072ab90890455e53a45cdf014981f6bd24152dc87078370bb0d22960c1c16239fe03ba71c64be25caf7f5581fcd8676",
            "1e64567f1b091475760fc20a7a3a8b94edba1511c5d38eb5d822787305db59e1f76682e6496ce16b6cc4a1cc85da3657",
            "5475d7b9c4f5dd8a3774a8aae339656a8d9df0077bdc454b5244fb990bc7243624410f7db43ce551575e751b564620ab",
            "fad8f8213c988fa4955876b86294e4e7aadf0cbf690477480f1eeb76ca67ba28ebcd0f775b1bc329e09e50bbd47ccec0",
            "46ce02f8263453a5dddcaec166d6f61419a9d6048574e80f41f4ebbb1dfac3e6dd6f584dba69984d42f73cc715f4c874",
            "b70b3f990732312f368359dfcb41d2bd8d482fcd8deb8d2185c5aa5a010dbad82835600688750cf9f5c5eb5b8cdf719c",
            "bbebe275c943ebe8e1c2ff6f8e9b1b804c8c959c4016b6667f505d7b5806b5830d1b61a979e969c72e978a8848243e8e",
            "76fd245c974b959fa922bf12d8d64e44c4feb2c51d6e243aec5b01c93929ae46bc0c9abcb84a019c0e332956633f2afe",
            "a7bcd9dae0d10e3c22496e32225385cbb86fc254e3c5d5eb05ed9bee897cd95ebc79821efde5ade6fb1ea75cdc41b85a",
            "02ac1f733e0b62e73362501312a623a2cec0d42eeea2239975eddda268657513dd0662d9ae349e97ba2fbeee17b815df",
            "e8df47f8ddf86e9ebcb480c000002c700000e59fe33371733201d91a584b5bd85a39fea3f7e7f940e457d94e36f4de03",
            "f1a6b7d1fc526a14f0539af089e1b4d8a4be487d8a769fc45718fbc5dc3d9377df873ac4d7557d7e7a6f78bf37666387",
            "a1e58745acf3f5c37b08873daad6de3a1718ec5957b7ce2cb2a569312af4b595fc706efdc61ade1bfe20db7d7cff3d4e",
            "574022abe12e230544d337c6c0f729cfc7a1512da2adcd0d4a2732631242c018b10261774c99227f0badb26a0a6cf938",
            "98b9112da0264a41ae7aec88a992a7b22071e8a9072f97b2f5eb26636fbb81e558b9cf594845aa6470c32bbebeb0661c",
            "c54b75c3240e8ef2163d77612a9808c166e2b5f284ca5a9513c7640e59263fe4700c75768153e3904f56bc02405c041d",
            "f54d16fb7526e8f4bcdceca92c4b4324ad1c4b4723b7c1463222e7e0ee753f43bdb9d8ca48d1ab0e330b3f64fea4e2ec",
            "2f974c2ccdb0e72be439f7a882c24f6a291ce08906e9f5286c2bc6837ff0e36d677c1bc3bcfdc21e9616822f950b6b14",
            "e91d875918a86a4dad3a297a3c11663bcefd16cf7471182b6f6cad59b6327df32e917bf1b756b86202833de6a1f5e5e1",
            "c2fe73068d0cac8dde9a1b58d8fe9aaf579f857eb153f402e28a5c593be1424129d1802de047469e9d85421d15899259",
            "cbcd14178675f0bad90ba46857808b46121f88835fce6ab516e7d77bebc2e317b9eee6fc858de867649be105ff9ece71",
            "dc1265a1ad3e3bb6e6dc70d00a9c09dddbef41d7c7cc5c8dd3e52df23592b6316f26411deea461f18d80a2f302e23ad9",
            "f44c3f22e56d1a5fd5d5fbaebbdf0e413bfe583af48184630fce0407800c636c67240fb2b3776435b20399fc7721c07f",
            "ce3c017f24d5a43ad9487b0e208b196b0b39d078f12dc6fcdd6c721e85a680b503e61a0b89cc87482a831c795be55590",
            "db2102fee91a9a4f82f58955bc8ba3ec29f542e42b399f57f77447662eab7db330149dfe6ae15af6e108c9b4cc6b67a8",
            "9a5f54f62c8c254479a8143a322089e463d26cbccb139359c1be3411d9e4cac8e32fa859041255477315bf7c28cf4ec2",
            "c6d360c3b4ed0a15e9124928ce5f8d1e7dae88a4ed94cc40fa9e32841ec2919ef239c83cc8d8c95fba56341a0ed88d2d",
            "84ae2db14ba87df34385dcf07227409bc8b127c332d734d760dee165169cc3da42cc72df37734ed45045b2ce35d7a3d9",
            "950e9060c1d3e3222654fa43a2b37c40262ecee6b4c3fd4dee85c7a4b78077abd7c41e139477668f0eda3232bd870675",
            "1f4bb182f24e14e315baba8e82dd76f70eb654776ecf8baebd3657d9d555bd44b23539edaa7060936ceee0eee926d646",
            "7688d161865e20b8bd6f108c991a9ff72afe08c9544aaed0997191e29a394aae506329318b68858ea0129176752e8169",
            "7d60619fddfff46f4702fc8ec3b3eff85ee53d109407eed732fc1b870decedff43e1bf901c556efbeb78f902d61bc9cf",
            "53bafb39a4231d882ffb09fdb9619ba44b2cdf77118648f21108945df7ea2551e3a9cd0da580441065de5b3fdb7b28f8",
            "ea6c4b3c78c7dba892d0f0c502c12e28b6123df93cc62109028d4b63152c63ce00c5542c08b563867fabf6ba162bcac4",
            "c21e7b3340df9103b1c1e9849c64d78ee67352cd45285d7fa3cf7e615b576a95b5dcb3a3038779f96aa4946717471b61",
            "4a1408f73dbccefc08fc96a994bfe5549e0c1646ea881ddeaa4af84af2726a7333aa94fdba6a461f212d11b2399dee3a",
            "b624b40e91932a9f1ee8888498145bc3b6eb11af102ad756c5bc6097185bc944081534edbf49affeaa712e43b3698f3a",
            "5d601efdb663bb273073e62575c8b19317df53701a599a6c582b9fc30d030f06a6296bde8874b53bebc62d20f8a49d45",
            "1f40b94a5acd694e019d622238d25b13e86c13588a66b405e479516dcd1e6406ee0e8fa5ea2e5f245184ab95ea5ed1d2",
            "bb921b29eb855ca7eebfddb2c3ef38991e49326c56c161e4cb3f52ee09accb09783717c05c3b6093ee917a39b7b9dfd7",
            "7948d53758a871c792b79a760a0fae3f3efe95a8fe577263f0d998703d5037edc1403f07fc264af83357f52fb1c21f09",
            "aba648192b58f63f120a1f682ee2ada7b7294b142dd120018aa73571ecfcd70aae51e122d0db0b8c1c61840ceeb04eb6",
            "91ebf6a138ef1931baaa923146aaa3a028e49849d1ac8290dd1a3b4e9a88937454f3ea8e10f4cd3ed0d14fb6fb2e422f",
            "1e90b391363efd8049f6f9c210d0b1a207735e9c44d940c972ce0f2bab381ec6dd173561c9ab63570f51439b3dcb3902",
            "a791053eb8e63dcac07da5b3c64b5c354d327ee5bb40f8ce852c95515b62e79f003da497df50bdc50672bb97929b878e",
            "53d09cd75c787cb82e2bd9990ec3edaafc5f2e1a9a81a0dca755481e88330a28148c04dad286a95df5824ccc3e358a49",
            "7eba6b2b0c79b052f4abb7d4cf238fd4a5a3a23f71c2f805797dc4e575be32a719cf53586cb833dcd5794271b5536ad3",
            "61eb8abf9897c59f6da46657fd13f17fde3f0cec73c0bfab967b8cdfd4ce3d56f0cf929dbfa3e9c162fdad80e7b1e0e3",
            "3286bfb08cf82f450d8f85ff99d6fc3b30d1fe10fe3749cec77a1e27affe823dc6bf4f653dd6f23829f21774b1fe7d8a",
            "e4b196c787dd7f6108e7b747df8f153cde37ff850ea247bbe8c7a28fb7237f2197f4f79b93c71a1e07fc7f21f0e9ff4d",
            "f8ff58dfe3f0f32f883ffbd760f4b1f43f23bfbf038d1c00f89738f0b1f83f9deedf3142f937f1ff76c18fa51f5bb5bf",
            "5044f3ff68e3e4a51010ff60f5c3a7f4e1cf30d3fdeafd17504b03041400000008006d996d5d6c08b5492a2700008e33",
            "0000110000006f72646572302d6d6f6465362e786c7378ed7a55701cd9b26d8b99c1228b653158cccc60313333333358",
            "b2182c6666468b192c268b9999594f3373eff51c8f4fc47d5fefe39dd55155bbba22b3637765e65e993b652440405101",
            "e000480000400088091ca54902020092c00100140024b0ba80b59583a19583b682ab8da1bd268d8ba5057e3628305926",
            "0018f01ffc7f8dd278a94fdd748881b7a24b0417e84c345488fc522d6495fc3ea0383a95a4e18afdf59ce00b8b0c1fa9",
            "8b654c0392a570fad36fb95f470ce8d5ebeaef3a44fb5ac453bfeacb66ca6b613714f61c0e2fef89a3c6759856aee892",
            "cc8aea317371257d2b0591eb9c88a098c8452dcf239ed1815178f0891031c353175d0a87c1e4a5d6f5bdd0e36c8b1b6a",
            "09a009972b20c88fc52e5d5fcedf70eacd6817235443ee7d72d7db558482ecc402fb2cb61d3cd5ae218c1d1c43748963",
            "2699c9d087685721c072d1d0f9e124bae10356782f4acf15805a7f203e3e839b66f59b7642b0d08dcc433333749df229",
            "f529af1801a18919a2bcf9f052023546c336f75a0669ecc8f7fb0efce5876a1b04880b5f21889b954692c1e59571c196",
            "052e99cfb404434f9ce984c35718726aeab959a7fce8c7f656f8d95e9fa2a71c6d7c348f71e1dfafd70844a57cddbcf5",
            "249ade2c49fc183df25d273d3544b5720cbb349925b50404efdbc201eec52ef005aae75ade63c82bc4fcec108d1baa0a",
            "a43efd9d3cf5ad621a9e084db74845e3b05181ece979c27d3ffdf39b4bbebe820064fec5a3eb14e988aede46926f4f61",
            "de3c5adbced0c29e96e68ff37f3cf93ff803a53152d2dd74b041a7a24fc103b733cb5188fc12ce7ca22d040a5e3e624b",
            "0875e139c4d8cbeda438880a58fe2022dead5e4f438bcb9ce1070304371b8ca561c8a81c44f5ce65415f1f6c1d21283e",
            "d82508db96aef72039cd5ccfdc248988bb9239b5465294341666490cae12094f9d0b56218ed1286341a1e4e104a17325",
            "f9aaf549a9aee8dd9503cdf0cb1f97055536b0641ec48ea578cf25e4efc3f184b0491e4d26a10dd8eb19ff28f718b020",
            "b8bec75294376133add44bb7632774d536b0e8b8144c694779f6a0fe943a2047c8887c473f26d24f9279edad56cb783e",
            "9ce6dac302d55d327036e8178657b802f43b27891aef933f040100f6a0000084b76f5c2c689dadedccf5acadcdff58f2",
            "4a952cad7358115fbc269f217a9cdf618f48c84c29ba7c2e976b51aefbe1241d5a6bd769da08bcf3c5544ef171d55a92",
            "7f92b944b1dc5d564f6feceefc1b27c3ced55433453829cd22ffb47e6caaced49ae3878b6c2f8c146d2379d444e11295",
            "b41ab9690f0beecbd927879b47eb024a02bcd340174640a655ff8e5cd187da5266f58bd03ad2ec752651747e25c94e05",
            "6e154b22b6e14592dd620a6a6cfdb16215d914ec00195c7f02f430c2cdcf03cda173662084ae83a889ab3461f9efc78b",
            "8f301a4307fbcbf9920ab17c1a6987f8b96e33f3650dd7ef05a8895b683b8f996cc3eb022a2078fc742b8363791d6c14",
            "140de46af5b23e011616303cd7783ad109acf92bb9bf74da8a4ec73c158db32da992b422a86d1f577ee3c2e86cff22ec",
            "a3c6a5cc0fb90231b1946659efead44c26eed59255de99bcc350e2902f3e45a65a32f94d86132f855bb237b45d354a13",
            "c551084a0ea2a5de5acd98835d5b8c68176636de0c5c1077ad3f3fdc1242bf0caa27312e7438991c374c0d08a1c1eb99",
            "62d7edc885dc67bd8fe9b5809238b80bd8cbb782ae9acae0f6184d5e2227767bfe4bf7fd57f4a58ce1dd93533d6ed3e4",
            "0072e676444cca5d062d26fb2a9af1292e01ca5925327d9964c8a73348a7812734cf1fba654f10c0359222bd7e951692",
            "ab4fde788467bc41206215718c6bafae76409f9b75a5a398f145daca0eec0724ce8e7bbbc9c3090fb27a9cc8b2230eb3",
            "c81ce550fb81d264e6c2cfc89ee1b5446947400fa7ed24c0c2d998c9b79e791d0667fd03d62a292f9eef695d1ef91eee",
            "ad300c56eb0a9af797d13d664ec4720afa1eca0bedabb05242e6d7be5f2eee17bf5b416b52034e124ef41b323a46d0e2",
            "a4ead4db7870c3430daeb7b06873314dce65758a7a770795cd879333117047869b4753e59f28ec34ae6ed04a5f645da1",
            "a3f6ca78d0466c2548ea2be73d8be3d664d5ceded5c368503816bdcbb6cfc06ab17d98bebccda46d03f51508a2d1bfee",
            "d9a865f52947aaa31bf6408ce70145047491a80c6ee85dd94106ef36b00709aab40767366441044c5b8405aa4cbff047",
            "a9a3e96b02c4e892d8b2d3e4b3d7bd464fa1c06e92fc320dfa22a318c3489b6bdf432473df8494019fd68af0f025fc2e",
            "f844d1028350324c5d54215c212b7a478d0c62db834fb77d66b7eff08e9c7d674e77241dc7611cb0f776b3700b166796",
            "37f46a2aa51e52a0e12b50dedf1e9f6490aff8dc92f0c8b9bab9e7ea3ddddb39b9e69e08b13f3ab17029b9da6192c46e",
            "5e93cecdc5b61bedcfad431e0603d5ad9070123804a439f8e19006296d497ba24a5685e11c4ccde7c98a57b64bbb340d",
            "a5732758a7a881106dce55102f8ee65f5674761f3f264fd8d5a1f149630e213e9cca6873d4d678ff362ed4176288bf7b",
            "7bf2e32d3660012081dee2c25febe7dfa3c37fafa5406f6b29d0ffe350feff1ca5b17c123d6facf0d4ef956064395fd3",
            "11e987034268e59ad81c1ee2a42976f2071214c7c75338b4451aaaebc94d7781150f3ad796e722a623182e5363300331",
            "68d832bbc0fa789fc8cde1ef2074d093e893ac61c06ec4f3bb73ca2d2c8cf43d220a87112a9255c499f059ee5ba90ae1",
            "65b93d8e269bb90e6df104e2dae79f8a859506584c9bb7bdd281ee553f70a0023f653582391db534debf5a2827a8add7",
            "c867502fd8e8822144c3e8675ea18103f93f31d510fb01b78f54aac6024903b5c6ef282e1af127b5f4ec9a4553a9c71e",
            "2dde0799f5d0e986a39feb5643d668a2088be5a49b151d8332a84934346ba3f899ed57b92070cffa7e88627087dab152",
            "4c1c4f996f3bbf9dcc7c9dbdace1dfe101fcceacc630508745de68171d18008001f89fe5c6dec4d0d0c19ef6cfcbc73f",
            "2c2b356ed53a850eb17f95f4d8fbec69b24007d634059fbf3759302cdf74caab55d2dd66245a05e88b05a9a9e2b1f76a",
            "1255adb902793128729f73c623ee6af7fe774dcde8f718250e0c599c90e4bec207fc4c2550cfe79edc18b27a6926d0b2",
            "fbe01fae70a95a0757dac46e172c8eb03a9088199df8e746eba0f495cab3d099c7a9191be5cde46182256af4ea99f043",
            "698aca555083944bd18bb320f8f54675f7aa1590f3d5accc4d9680f43291e1af17ebcc8ea39586ed6deae8522c39a29a",
            "ad90aa14e2ed6edfc7c5a87d0853cc0957519abf7d8f072d7602ab85b4725e2579fb922b57963ba1f49586979ec384db",
            "d27ba44f7220dd61e1551583db00dc6908fb2e24548bf4b3f64de91e127a3bb2d34cf7d32efda0121aca0362e4c5e221",
            "2b3c44ceba0b3bc6a8514de79e593ac2c622276e708e89d3e0250fa3044c20fc637988af1dbf7dfa88de50756d4a4606",
            "f4cdcd2083adde3d1c45147c2578414fe4aedc613882c65ac9784175ec922b4bba362c7926eae290e7cc2342a6927970",
            "94ac8de7e7e6d673315b1e304adb8d285458774cf975c104ac5652135ca5bab2b37660692df111b97d01fb2f4271bdcc",
            "ceaa89d5efcd625f6f567896375f4fbfb31ac5fde8a57d7924aa8df0b05adbd08ce0cb8358c79b0966409da114a01d53",
            "0aaa6aadde4baa2e9c1ad62ef4fb90a109a6ff48715e4361ab4ce2a6b5adb3b21ed6a0079611a05ad417033d341a6f2e",
            "3ed2d944d7615ca07e1adf60539436756d5bd7c70039ee4fc78b94377bd2d80144fe657e735044724364860a7a255cca",
            "7bd6f0f731ed58717aa01702004023fc23c5ffd3f81c4c0c2d0dff3aff6977476acbe6c44d48de085bde88cfdc1ccb5c",
            "ab46c691f71aabe5c74c873fe06ec9c64cf2eb941ab0c4f4dd212174e1206e912135f8b5902f128211eda2fc7108a278",
            "fa2a25ecee89e4e46e251b389317afbd6cbeea3ce6a95f964c5694882bcec8adbe6ed11beda5428ab94e5efbb9b7058b",
            "d26f33245335afd40c364ede4429b026e9f6a14c5e2b09ba6570deed4b9f6c72a14527e247496fe19eb845cbed6bd551",
            "14d07dd9e64a8f01f54ab147d2076bd0934d83f20072d35782f36c570dff28c519b78631cd9d34b61f977958425b5ef1",
            "62eb79ed7edae9e687e4af8a4c72dd11f719a52c5d450e9c83038691bf7ccb07ca214275cebfc9e8498f9f263e182301",
            "da6bcfadedfeb218e144f95c8ba6671afe47ccd2fcd3b65d53c71115a37d68efbe919c376afd61a7703b105d5d3876dd",
            "f5ab7006d0da3e6a3ae3a6dc2c2663b217c74815898cd3fef8b2291813b6f6a4f6e2068e43f273a5d53e4ba6a4b7e723",
            "667fda44a5ee3a1c59374bfd1917cf91b319fa8a0fbd873d8c36deb5dafa4017b64c8f895d7db855fdc2b151c40aef4e",
            "ea31c77a9ad194359cbd8baf9fc9bce0d774bd51943aca59bd27bfb08f00c6285d21945c25f2a44d72413b5a458d6c82",
            "78f26e4a843cd675e4a0673a48fcab03afc67b16f47e77d0afb94cb43a840d418c78c4057ef9c5875cc0655a5017078b",
            "d615da5ef6cf17d673caca73b4af7ef07b84724914125678a461d54bcede4f33de5eb73057a7ed2feb2d155ec77dc3f4",
            "2fd719ed1ecfb7de5e196e2a6681e45acf3b8773ab1ecf070ded4f73196d7626a7501e6a41d22692c4ec84eff6c823d7",
            "81c2c92d5eb614963ee921eb9a89ad413a7028442a1800452a2536d978eaa59c35f9506d02736bb45536a60bea1515ea",
            "eff4cb5a0021b0197bd2a482bb94ee52a8ba544ba29f0cf6567f710f3d6f323f01fa44051199409eab7c20c716122274",
            "99b82eb80e2616c2ffc0e2251cb2b9cf9d335d7d4890064d73af91354676a07fc0381890aee6ab00793c85f1546d8704",
            "aaac5ecfe611f190da94ee263a6a494c053382b3451aa06ae411f215ba9f261e05a7a6ba13a1325483585a494fad2f3c",
            "17589a1672d3849148aed9201dd26525660a4f5fd4e253035eb5a2055160b6314d1a44e225bdddc0183852843ca971b2",
            "a0e151c1052c3f6a8e7eb9c699805505a2278257c7e027679e821ebbdb14944e06511b4371219f3ded84aef5e1e080c5",
            "43dbde9dc4c66a49e91e25a763a8832f112f76ba81374004b0da6b98047b3f0ae0c13dd3d59869457942f14b8e509475",
            "b8ba9c4a0f1e0344fd24b56bd208110a6bf96f02a8eee0b9834bb03f1b07c66284692647761c2aeda543ab7c89d8f5e8",
            "ca37365c430a4447dce8358d23457688f373af9538b40d285bc83bcde3d052eeba585d2cfda8f295d6aa85862f68490d",
            "1cf9517dd1c82f89890f890e4572c29772d92f64fb855f704e91a8de61e8fbb1907a23101651c9ea67d38bfc9d5d4898",
            "4e5aafc495dd1c472eabded3a9d3a321ca4796e654cdc6d7ce15a390af831f7dedaad31a4fd4525a66fb1edbad2d9dca",
            "b65b8276cab39acc4d342a74f0f2b2c8dd304cac43cb706bce514e976e93e4fb3e4d87468398d67b117f9b1de9a333b9",
            "c91de9db32c2a30b5f93e34190d4594707c6c6246e7cc0549c100b909fc7831ea1d9e1df297e879ca7d1551f8dc9ec5e",
            "bc156455735ffb05e19e6af184f151c2cb88ffb3e2c713169d4d72ece0328af8ac94320f994b7579168b7746258e52bc",
            "c32b4182a6f42982dff9b2eb703d40a05aed6f74551c70a816d49e1714d5f30166d8151dee2c7e0dd7e57385c65f7444",
            "e0ddae926ee4770eea9cebecb90b0b9c3163e38900f5e80d5ccb131acedce95ecd54d32b5d53e8509b7154a516c87dca",
            "f99db148dfe0762754579c5fdb7c5aca2a380225a8dd16500f6764c64dea78822a1c919ba2b155e44651a9f8de55f609",
            "a838e0266fd41e7ed75dc5a0d944f5133970a61d41b701302812aa9f9fc6a8c911cc200e93d5340eb22779eb1587f07e",
            "e9d394e679fd80a3a217232a5c0e2271db2f4e330d73d22f0cffcdca3d6964d7a078c88b36dc44c2b8216aa3c660fd69",
            "26a49d4345e282a4f19138593353be1e55fa63804a0e9ac9e96861c0a63fbefdf5972ed7f90fb168e7ce39f4db3f56d9",
            "b7282e3e1a99065cc6c88dee602ea1c6c1bdb76fa8e2dba7f136a2dbab87927ed7dcd5e6fcc3d3be6c3681f240bf802c",
            "59496dc2702a4d8a472b9eb68aa2b8b097627899c5deca38ed085f26cd70e3d9dcfbb3dac72421298695d16e76d3dd1c",
            "e63a8479a4b86269fbe06566a12ed3cfced4cff8ce0c668e812a2c091fcebeafbbde5ba464b4f96c7d0c90ef2119fed4",
            "a20aec9c5f271919787e4c9e517621f1a14630f65e83254f95c25b7c276494a6456bef00e7b27bf31bb60e54a6d4180c",
            "a505d1ded18d059fe691dd8445a755400df38d70bb198c35f16c7f29b66e4ddd5e7619cd1ce941b3320b8379b0a1b4e5",
            "646515aa86a53a52c42b58d02c1bc7c8773bdaad248978e7059bc0043fa7d68cc7702d4f446670b4f6650cc583ba666d",
            "5316b73af18a8942b01ce13584c7f55d5e6921dbf181362f8c6e5344f5ed55dec4f3b253fc8d6bdfb956f2c892258f5d",
            "19a61596553ed0ac438e90cf2f039d2022c06536a8d3dd5aac85b0417c357208c1dfc55b140f15c175ed316e324d58fa",
            "5bc7b81cac3af8bc597122b175f3d8001acba74c14ad5e9632a6f522cf36385f0e04f03a5cef9b021a30d2ecedd53bf2",
            "d6afd7da7f4b39a386d5e29bde28a710380000f7d7aa6fefe06a6168ffc77a9fabf0dd6c9117b1eddd03fcf04a7c3cea",
            "705d78ae4c3f12a81c1971a0d6dce49056358f51dda7ed04ad1867b22b38b7c3bcf1f1b98f5b6cec94f7221ff2a80785",
            "3838381e7652ee8d58b88bf278594614344a402144b88525b79bd7cc12d69c9593583a0aab15d0144c7a91bfe9a1718b",
            "69b78550b6201fd98d260a01b28a41c7ca34bf40165db0c5d6b24fc2e4afe54e0bcbc1a83310cd74d1a9299c90928bee",
            "3d7a12922b370a610c3e79f31233cd0d077e356fca4c906e5a739ad1d125be9cb897341528002e9ce58bdc29847ff820",
            "58f7e383e9b7c80bd1746ccdfd495c209a09c341fb9bc4932858a8efb7a23e4f8cb9f8845e2e32ee96e921d34441cb36",
            "a4789f36e528d5b33468e3c13efa169a77eea9ed2350d88b8cfe10ceb3bd465a96014a07c4328ed68d83ab46c1b3cfb9",
            "a0c69786bb34b66bafbea8acbe68ac2e9759b7b79cc588d3281ddbeee1f4e4ad50a1a5da133809f27f45c131850d5535",
            "5200c7aeadaa50821928d469544bfc708237c120a0cc87e1e8852fd5c5986e26df4177aa5ee99857a59307961d34b663",
            "55e706b45a548f0e2abde93795a4a4d028a7169506b4bd251922c172763a0c72d42c6846ce603a1606da3d632b359f2c",
            "33d06771be10de7d9a5d373f1624e4e81a8210f27e0122d7793f3a96b51361b58dfbe901b8bbda97778ae45d42090f22",
            "f848f64a438733b48af362f2ce51352f29b4d2645ea5724b37e8e9314be99a04e649f6ca05b62576c8db3d4c81c2b833",
            "9b8a252d0506adcac4e419a8373b763bd6a7d1ea71ecbb4ddf390ea368962f7d28b4076911a4b1167bc12ffbbba369de",
            "7de1d6e9858c412d8cc5a0df608ce48ff27b865b693bd3a424da93643397b7b9611f5f0ec9ce5a8fe611be6965c07fe3",
            "7b65bda82683247d5eb9a31aca86bfddfb4ceb2fed97e9a98f736f62d064e05752a6a41ed8ac850855ef4dee67118fff",
            "78cbf30cf73b1b67e9b9e5297f1b9dbf1da8ff65e326ba768606f20e76a656c67f9abac6a81474171d22d85aebdafbae",
            "23264541c11854248da669901ea75a5f23390e796f2b6b41419bfb365790b1b8c8f962ec1b1c64d54429181951c56210",
            "4d2a1c84964168d4b0828586cfe35d2898e22cb89846b2f75fe7f98926cbea768a3bf75c93d9733dd80ab8dd06abda4d",
            "4edcae98584290d5b96843f322b8110b29e088d1d133c1a1e7801a568f60218c7b8e9fa1d4737a319f8310ac365a6e2a",
            "5c577906b99e4577074fc07e37bfbc3c025996371f06fe99361ad8e93aff31b5ff1efc95362a2c5ba5d0a1b665d43af3",
            "b615f8744b07012b829c2b141db84d70f9184af28270542f132bce3867584a2253a7d5c6c3b0a66e5ea6659cde955a45",
            "6eae1d17457119838558f1afc90abb4cd5580279711379437116d1999f7560527750c6afd53f3d183a5a087a45d59d31",
            "02629b5ea8790ba6691080c5d8d9f55cbb9de34ccd877b43e25fd0163391ee32b3fa097c7413c1ed48c9ac0c448a8a49",
            "616c5cee504c884e30c92c8d0e3a28d9619dbf386848162b9dd9c640e28937424b8209d7c2abdd4dd41bc675b0c24a33",
            "e7541b752218321685b31e964ac07e55b5ab567d17e6095f78c756c9e18a8661b4be0a5c800711f420a651ce5e64bfc6",
            "b6a4b4a63987427c26da8a84d209b3f6d07db34afb45ba5c68cbf695f85b1b390014774a0bcc053d713021b4b89bdfd4",
            "38752ac72d6a486c256c0fe62c085183141bf55154c20ba75cdfbbf573fb3d244d0276621634cda4f3e6fb3cb689384a",
            "2ecbdbef928e715ad15c37ecea83ea8e02771867f5c75069c122919f5021122c520e173ff0e9c650cbf384d16ff51f4a",
            "6dad96096c75b397a22b864f0ad5b940f035415bdc5cce297d94f22d5a94bd7299c14cd566bef08dd59ed4f971dec46d",
            "798dd7703b0fcde5cb3ab7edd20fbc577263b68235b2b31d2d26665416989ca601bf538a625e7a36fdfadae6dc532566",
            "47c94638407ab214b0057831059334984dfdf625aec81401927d3e294ab33db603a52989cf765530a22e2c94dabdc068",
            "1e5d73725305c72ddc128b91e1ba9fbfd2cbc833befdb769a27f950ee52a28002000fb3fcea4ffe64c0ef67f5dfe34b4",
            "112575eb4566542fbccad73528b087dc3db36166bd77f547a8cb7602b18e1d2b269feb7fc8e48acb45b466bf7a06bfe3",
            "a8fe6a355819bcdd49e04ac3fef5d463ebbec55288039b8a78dd624bc8386baeb3748d44b724faf6f03970f5965d8e40",
            "89700813d79583aad57913cf9b9536a8170bca513907d8b9032ab65b6fab2cbe38980f3f8ca93c8a7393c9085fdd4e96",
            "004cbd0a5adc17b7d355c89f15d6ddd776830d624f9bb51051d7ddfd1cd9bd4d230678cc654221911c3bb00395bcd047",
            "2356eda4154d56bac7fb5c259adf66f0eb41e99aa24eda712cb0b12c0bab01531061bce8ba09627e1a5c3611ed15b0a7",
            "2892fa0fffb181d7fb595ce8954ae4af6cef601726262b42d70328a8301a87d8e411eb4a94afb34395453919eb10d667",
            "a3b6fc64b29d3e585a7bc0b75d3205bb39706fd0c221a9447c2314cf108f6aab07e8a25ff9cf5fa3c5844f662970845a",
            "f2cb621c5af8773b4b41be0b9904a3c075b0541b8298216d62e6a40eaaa7a79b849926f39246ed6590aceada965d361f",
            "9a11d43170afdc0ca73f56c18f36a73ad457190ba954022ba867a2dc9382cb74ecaf8d1e87d7505b07ffac394ace7380",
            "b9e51b782332742c0ba91137dfe4ac830fdb27eda9a2b589566bbc09bd1d00462f2b46c2bd9177d380d3deba19ee13b4",
            "4cf77453c012cb728ee81573c93b2cf92c18a1f374d5acc4e623fe0320c974ad0aa13729338d95f251b2ae7e9957be76",
            "3c59357ea0a850c4ab14381931faf1cb72e299ce3e3078adb97469511bcaad423f97652f5c717023b5f6d6b79ecd5b8c",
            "7c20d9d8745d24fb87aea5ecb30cb93069467cff02de23a38c240928eb993c46433fc50224832dcfefb1b9783bae07ce",
            "c6ad4d1941fe4b19a3b1f319c6e754b806622c8397d6d46aae4c45b965ae0770d8c26d8c77fedd3773e3cff7b6c9d24f",
            "b529bb2f6b7accb79756ac87a305b6e8f9eb2c0c8ebd887acce12bc65929c6be6a5a34858c00b5cb7e896f5da54f60a7",
            "1195947c17509d8e13d9f49e09448ee1c5fa6b07971c54e95cedbd3eab5c54efc7eb5e6a55a9c8da88bc6b35f90282c7",
            "950a5fc2e3cc8e1b730f2d269cb4c3883f83b7b28a353b316680aa62153cf32d8c69bc5f2042884ab38073f6188ba442",
            "0c3c9f63d740845cd640670e0dada4ac75b4520fcd5bcfc851cb3aa676ee69ab04029ea5a19871ed7a1cb0f709569ffb",
            "724fa28bd01890f2034b618fe1dd6dae4b23b05cc3a86f6fea974a794cfffd7a7e30e6f17501002ca2cd5899a4b689dd",
            "86cb41a11cdd985a90e674b286777f57d27eda6140cb2cba03bff248552e72f054cab69c9e8880ec5e73df9da2005ad3",
            "924c488922a904613d914fedd46c6acfed6a439f6d5bd1953b9ff81e67c24efa1d9ed321e899f7143551b2de7548d5e1",
            "7897158e8366d135c9781b091c52e2e36429fc9dedf76d1b67d9e97c281b5e04e1851952426b0c470af83e92e86b542d",
            "3e04bb8d4cb01ccbf5eccf3af94d7be01add06efa5c6edb7168fa5202e3993efc5da61b6b0f33cbd1185df79cd665c77",
            "d38d5653413c2cf2c8ae4b7dfd96ec55832fb9004b85fc5c2ec030001cd35ef47deae3015f52ed63d6e8099654ee6025",
            "13812f0a8f39b4a37b502ac269b21a62fe49a55c874050fe1c46a7c4869fbd4b2886c3ac49a02f9a52c0e028f0c96d4d",
            "e9a1e355ca25d525fb193df3eb8c96cef76b9a731fefda772d6c777e914c755e254084113ce5a457673778af50bf0b88",
            "6815c59cad6f23cab77049f48fa2ed5fdb023f4bb77f6e0a04447eff836c049cfa9d22de584dba2146e76b099e515c03",
            "25183cfa5b22579f17585bd3f8f9ad591a7f7d36ca5fe029905d22bf584b0b9243c71fc0b48698336d64ec9f41072575",
            "40173915ed96847d746b2a1c65d41b925e42f015ca679d1d2f407512afda37e54fe08e1afe3adb0f259f50250ff41d8d",
            "43ee307bbef06ed0e7218cfc939200e70f644186e0a3a78ce345273a87b31de668091eaab8f0bb81eff4c8fa2262fe71",
            "a7efdb83e57606320fc43d154bb57f9b3fc00b5ead7e7b1b91fc9cfdff708fbfe6fe7706f2ef66df1a325e320da13279",
            "8c502f36205c656d9df90eac73aa92e3888543e83a99f43cb723a706cb2ec917df5c836341abbbace928029f35dbd2d6",
            "2e8191aa47267e51ca360f4a9f7d43f78a0ab9b73a8840d85b2ef43b2f43b543ba851dfb3ba32fa6b0ad172627f3b10b",
            "c63ff802a5c1e3b425651bb0f89777f0937b0bb65e2de61c73b26a1822932c9e90b2467c083ee007a5678393cfc155c4",
            "0c5f511556cb686f0fb1fcbe5edf929956fd36f175100080f0afc91bba3818da59e95a489a5a99dbffcbdd5f142c56d6",
            "7a11ffcfca3d07fe0ab68573b112a5535d71326ad9016eb317c08f168ac2061d71eef9864bf420572257d19de3a411ef",
            "f511e43265028bd907014bb37b58b513933530cde93088e7f9def53cad0631ef30970f02ce3bb0c37aa1427b874ef8ec",
            "53b2fb713a269f86c9dd87be3674c24238154ca379ff9c21e8d6f8e8042f6afe533e275d10c4322445590d28caba4d0e",
            "7ecb19637ca5875d4befacf11bfa833c66f636237463a599b03a01daaf413d11db2ccc11f914abf771ac6c1d66db1fd7",
            "284621c96cea26a13f396a851d228833dcbcaa28920f700ab40516e35b1e6f6d536fac0c2f78f785a4cd7328a8c3d09c",
            "1c5de90ca03ae8e63c76245538ced9b4ef4db2a93a504bb535414aaf049037c19ad4b339ef15359aec6e5828b7754198",
            "fa42ca8f809b62bd84a857f83c3b452dbc52bcf871b77ded81dafa247c0f69bb1aaf4475aa3a0e0b06b7761e18cd8016",
            "08fd45ec3c32b466fe0c918b243801775f0336977d06a1c54a7e783ca1d75616d14a4f0b978196876e1326dd303f21ab",
            "6c94ce3996e4d95844ccb90a8f24708acc34dae7151f77ef47c955f285b2e458f0079e97a463ecdfbef10b2f47c8f8b7",
            "51d1db81f85fec47d7425fc044d7d4ea8ff7cb19c1070dfc1111f4d4cf1bd2ded978ce3f84f887ff1ecc15100147322f",
            "bd2a3ddfd56d5fd8d7a0ecace6eba751b1d642712a0b1d2c61999e36aae80cf09386c31348f1e4cddebc2e833501e09a",
            "65ab5ae7c10cffb0d266df5e5e134b2e825084b83458439293ef39d3476afd0f82aa944d42c0b2c350b79b2e1d4b603a",
            "75ca331d86c0688f1032fb5f4d985e448edefd6e0a5e10df51c2df8c76e22d63400240021958ebcbd859dbbcf1376b3b",
            "c3ff6ae4fbcf9625e0cff617896e3a9cbe53bfd5acad95bce544d17d8ac2a35c7c56d96f815ba754f5a662d9f60d5559",
            "f890078a9c212e6760a8575b305e491697e6894ef69fae88576238a5bd385f49cf8d9d3607168008ba0394a3d46da2c0",
            "523f506d5d43ab15291e9b1a8809b112c97665f925e34f48d2b1ad4f4de77cf62189cec557809ca427a6e14a561acf97",
            "2b0aaeca4d414451a583b7ea0917ec114c2a2bdc8c1ba79683d0704825c7f127fc42e6c790991611621262e01020d920",
            "140782d18bc807ab217684aef1b4ad88af777f18a88169df9765966f94afbb64cb9c0362bbb512bf3e306fc2081d2e87",
            "dbbde57cbeb0d165c75dfcfea294129aec0cf3ba10908d8cb4376bfbfa947feb36edc9e5d9ee31754cb9cc36774cbadc",
            "36367b0cb3ec3a9166becc36558652e2ec3ceeed72747cbaaf74f87c53fae8b1b7f95145c94330578dd1ba160938c5f2",
            "05e495746a6bec100b1668841eccc5bf993b7e6672c9a32cd42f5b3eb55a73ce498269ee3cb55a99a6cc38ba7d790c1a",
            "13a7a77bcdfb83173030deeb5d402704e077365c30bf0399f966a0ef41fe70c3bfd9b0ae8dcd7f4cf8275263aaac16e9",
            "10836e019e10beee926cd47e1d9918a89fc260a1f43a3e3905b1ff10248fc11ff35bfea8f5c08c59b365f944c8261642",
            "5490d6fcecfcba3fd74fd4bb9d5faf0af211f3149a19779b0aa4c0d374a95ae7362ca3e01d9544a14651a9b34c2535b1",
            "13797b6a88ea1510219628a650a342259a5ea24dfa15b459bcb001871d059e79a0421fc9ada221a3866b6b792701cff0",
            "a53d3c36a832bf2b16bc6fa9742ae35d7b56de1cb5b69b8a050890c43917a8792d3f7bde3726f6872dca112f0b4695ce",
            "e07cfc66b63d47c6c6f8a6b56eb991ba6aba3c1d55f214deb3c8a046f0191acb7c494939465653de9c4855fa76534b63",
            "2a770528b5543f260aaceda8ca85aab6bd5a38a89ca5d3f4db96b9e7a9cb97fd3cb4161ebf291d6b27bbbbef3229d676",
            "4b7051acb816206ef8f411b2efc0a9735bccf2095ffc3d38894f31dc4e708d4fe664059c764d3033305a36a756b9b7a5",
            "3ddfc72dc6d8cab0f43d8bdfeaae57a0763d9a45a6e8b5c921893787771f20c0d94aa987d7c32595c7a9f3df735f364c",
            "700704bfd40d9fdf93bd264fbd4770f3f3fe667e0ffedbc8ac3d57c3f466b3196f91991af01b3af117a1fa07a9f89355",
            "7d8bd134efa2430d3a85784420b9fd684a5117a2d8b7c46b4317b472bd9398ee35c20c2536fd47d103ee0364cbe5850d",
            "ebd863a76a69505bd8a8896689e2a70d2852248aa261a519b61fa75aa6c284bb7cf1e3a6816c9813d13317e7bae7beef",
            "ef6cb692671bc964e286c5365b9318760cfcb91a3f331776e54fd643e009ab33805c4032d177ba0e1d6983ab6a7dcc91",
            "4a42692107fbec940335f45ae3a858873ca796eea6f01204dadd99a3386ea9eee801b10aca3c8fe5f9a9b8be7b2d956b",
            "2cac15c611d589d3b3be37f6a3f9998c455d40f9bcae353457ed9507c00d684e1e8a80024d23f7a3f406bf366ac2aa7a",
            "d17751efed0dcc0260cac20d81e4f841070a23b13ea1041f5e416e81cfa9ec9ed5c8a6ec98ec2c4f58ee5386212f17fa",
            "29958bef7b6ff7f7e0f7e767d4a97fd6c28180a901ffaecdfd57fca6e9fd5705ffda55fb774cbe85aabff5d8fe2af86b",
            "a7e14fdc82ffa3eff057e17f6d47fa3b1060fe14fe37cd49bfeaf9b5ffe4278ce0fe7d37caaf5a7e6d24f8894d84dfb6",
            "15fce3bff8654fe2277ebcfb6587e257d15f4bbd3fd18af5fbc2efaf1a7e2da6fe441af6bf2fadfeaae5d72ad94f40e3",
            "febe66f6ab865fd3ca9fc026fa5f2599bf2afc3553fb094ce2ff55def60f6bf925fbf9097492ff4d2ef4abbe5fb9f54f",
            "4092fe9369ff2afdafb4f6ef187e93fe07c9fd55fc5f19c5df1149f137f1ffe217fff8f15f22f74f1852ff5fc6711909",
            "30f03f0491de3eb16f16b44df3c7ddff01504b03041400000008006d996d5d97321c6d28290000f3350000110000006f",
            "72646572302d6d6f6465372e786c7378ed7b656c9c5bb26d3b6666766266668e999999993166662766666666668a1d3b",
            "3173cc1033335ee79cb993339e1c69deaff7a437cbeaeefd59aa92aa7755ed55b5ab6525804150006000080000800fc8",
            "227ef89805f4b20207009001106f34046cac1d8dac1d75145d6d8d1cb4685cac2cf17240de906501de00fe8bffafd117",
            "3f2a0d8c8f30704cd2d11b54ba3b71a51bb63be3ef2d2d2c35058ea1b7406ce44b5774b8c8f9a9dc4cf2b06097443bc1",
            "48afe3f833154c6ca991c127f830c3b6d2be229c8f6fd2cf169a5c89b5758ec0d073b240a4fd2fe128abf99154262763",
            "878629f050710e8b903e88afc98bcb33b3f0e6c3efb8186e1b0087b6dad10017818ec84504f73a14b749ec4ce0d563e9",
            "73347763b62d2a7ea7b93d67151394c8697aca9873537ac35f25cca72f7fbb76419bd8bd90bd39c8156c8ebc13461ff0",
            "7d772ce2c1aa2ac27d734a30b7dac52fe95d961e4c73ddd9e4e7fed4eb4ff69d8370e23ca998810ba5d7418cb2f22110",
            "29e67d53730c8d46c003564dd7373f86881c18f730f3b474d0de47193cc8e00c629372545d0f508b5a5fdf7d1b697012",
            "b147c2058bbb75e5124ad8a5cbabab11f53c1ea25ad8d214107c4c23480c5f066aa8e764730c1845e3abd08918bd572f",
            "ea55a13720484e5de66ac633eca6b1a9da6bd94e193f8f54c15ef27cd21cccc8816f73c43db5f591fa6a9d0bee1b4847",
            "e2cfa3e4ae72796fb7b9bfba83da58386f36c5d966d4b15aa0f3b237cfcfc000d97f09ec7a253ac28b9795e44bdc42bf",
            "04b68ebd91a5032dcdcff7ff06f47ff11365315232bd743041c7a20fc183d7d34b5108fc12ce7ca2adf88a5e3e62dfe1",
            "ebc37389b0973a49711014b1fc8145bcdbbd1ebe2c2e7185ef0de25fad33957d4242e1246c702e0f4abcb3fb004e4162",
            "9f206c57b6d687e8347d397d952c22ee4ae6d41e4951da54942d31b442283c792a588d304aa38205899c8f1384c69dec",
            "ab3e20a5b6ac7f530134cdaf70581e54d5c89ab5173b9aea3d9b50b00bcb13c22e7930918c3ae8a06f325fe13168897f",
            "798ba5a460ca6e56a59f61cf41e0aa6368d9752e98da89fce8412d9d36284fc08474c3302af29938ebd25bbd8ee97438",
            "ddb58f15b2b774f064c8ef136ed132d0ef82e4c6f1f40e0c040020820200e05ffee36249eb6c636fa16f6363f1f3e42b",
            "53b6b2c9654378f29a7804ef7346c79614959d445d1352295e2a5db06c251552675f65d2047218202a2e793c6625c197",
            "4a9da6986d27f39d7bc7e365bc9dec7b2e935a43a7acc9d25c3440c2de5bd05daa68e87ac9c6d22989ccc4142e5c4d61",
            "10587433e7b11d7ebbbb7ebdf26d96fffdd74017064096f9c0b67cf142631989d6590cb349ce366a19550055b25fb647",
            "5d8aba7d02b3c4372535559601e0524357236b440f57a82055aab33811022b09ed62ea904d5005de4eb5ed0f93b9f1bb",
            "e4b6c4b13256e1f2723300fad14f9adcce590572054f5b06c444a9e8ab57414d03c23dce001ebf35fd6bd2ee897139a3",
            "0cf1f190400e60530b89bb416f3f6afee5005d071cffc9a822e25b8a59f6efb531fef0355807d56095ba20cf0c584066",
            "1fd44ff15d70051daebf5b9eaf7b44a33f78516af94f5f9a28ef49a216c85509f7ab55b8b0b070637c64f666c45f749b",
            "d8062de398abe79e6c991bbe15c6c2d4a96a905540b087e761b1b2e4c4b09040109625a3a12c7021a300e5f87e7b599a",
            "b25fd59b0f08c46e7cc86449c48182be00ede889cf7f9fcaca2a491dae8c66c07861a28365eb4d92b9d879ae263b8ba8",
            "60e9442caf3c52d05666a21f9b3b9d96cd6c211ec7402605e20503e13478c0db71d0a4fce4efa35692abcfa7da30d5f8",
            "bc120cc115e5453b814d0c7be38da9c5977334073501b63ee8a4e409e98ae9e56ef7164b286323ab2b9c40d05fd72768",
            "ca215ac2921673dc14213ce6783510e7a658b75cfd282f8bfe3349601b2eb556ad54277fd5caf6fcfe2dadcb35dfddd3",
            "d144a5e76cc65e8173e6f987730a2c6e261678d7f8d0cb39c13e309b8c15b8d9a575818c62e210b300f224cb1f4fabb7",
            "1b4818c8b393f7aa2567a89935ec23a6e8965f1d2c3b9aec780d7038c8ae4282c9a4f9515ca4c3dac6c408181cfdf1ce",
            "38bc6680ba2be6df8c1f592b6af94a267b6c42b5b458f771ccd8865d5c629e31102ef40aec738e8b476d64a2f2cc01b5",
            "db74f61ed5ace0f6b16a19a1347a74137389b2e178315db7ccefb23c9b229bb205aa4f35466c84e8bae79577df94c82d",
            "bdfd82173c88f011183e92aca77a838003671e0fe5b3307ef8c1dce53a5946933ad91dfbd633ee524ba970a737c03c41",
            "82645ab1e602f25deaea5d0033de2e4d02da093711a75d30f2014028250a448ea1c794015376533874c38d08caa051e1",
            "32a1ccc9fbb45b9bbdae6529ad18ed2ec41a7f696e76d690b53553d392d409c955e1f1c92a6fa363f191b197f3a46395",
            "148bf49a142b28b3b270967441bb7fd2067ea878a0fb92084aba11dedac6bfb3e074f609623f18c833f324640cd1691d",
            "c02e2d8a7889040f8b7c4010ecc6e4b84b4482b1d7a062522dad823bdc2855ae1dbdb6b63f5241b9f76061eb7d5d61c2",
            "a4799890857b6904d1de5a59c9bdd0df09f85dcad897ee8120784926c62f69030b0001f49232fe3c5aff9a38fef79805",
            "7a396681feafe6f8ff07d016d72dde47871074ec772cbb7d5450def2d1b2055c48f5246f1f38c07121a0406c4870a9a3",
            "3d3bd8095aa21d85d63759ece9d92066c8dadb5cea2c3ca3ae16a7460a0a527d57dcca022ef3f36caa8828198406d406",
            "95fd4eb4ae8ea6cee1f27808027cf4c45b62b97e247054feeb3e0cb1c0f2acf0fdb101f12d0f4a7e18cfaf1cf8d87284",
            "c3938ec30fcd60288c5f3765ed4f82cfd0620c56847e74d66b299985d24a7bd59bf682420715fa2c439d1758217f08ac",
            "a4b7651b9db0946908bde9a289e52e1323098c1e1537e5d40b7f09580fa910dc1fd8f2217a7e180a7d41e21d20b18f8c",
            "23eb9796b301f8c45c09b8dd9f9bb38a396b2fe44d8e540221db430dd83a74b0d689ef5b990f0570acd3be0747c1eebb",
            "e05a5e71a2fcdec3e4bfd46735bd784d0230008001f8e7a1e4606a64e4e840fbc707fd4f274b8bed8e48fdf9bdc6dfc1",
            "251cebea0e21c90e7e8c0b6f0916579a5e6a3500abf2b332e44c15545a6967919b136f16a4f4ddecdeb0793a7d28cb90",
            "d6e73dce8255a9e9b3c3d00389290e1aaee8abec6cec46223428b6dcc9228abb8101e329eff4527b1bbf9cfaae8b4024",
            "aaa0afbdde802f448f1bc29c9fd9cf91b37cc24f1d7d2f22490ed8aa40401f4f0c851b89afb90f9d3ffb58afa8610bc6",
            "bc1fee4d5e3b2f13cc14c763c18fdce609905670d2a265dc011f727a8ebe08f8ef19a464c27ea8b4e98a294760b3d7e4",
            "d2e402a487197dedf0b30c93d7a9e27e8ce561e6563b839b39b70d8f699b0135c98e8bcf5a496cc458cbcdb8abe014b7",
            "69792ef4b531f073bda47a918c233a8934706de4a79942f2dcb499ac6ab2fd6a6ff316fb1b75150a95648d5c35459490",
            "6cce1e68b52bb18e480e5af9b1169054e849f6a70d3358e144ce57b05c698fe216eec3d15508cf8a37499c395d4f0b85",
            "61eb2ee8e674f66edebacc6a983c182551a789f26ce5b0cfea26d5261206ee6a361424701c8972def07df89ccf075ca7",
            "201f14dde19333565e8ef52827f39ddc22be49b8fd9ca7895d24bb355a99dc1db66df93652fd65a37fb3dbc7df9f5044",
            "5ea8381de8dfee36c31fbb1da765934a8712744c7ac8e3d2b1ab444712d010226b743b5662b65057996a861ec14ea40c",
            "44b8785b54f278fd3636d0299c42289eb7e726ed475bf0f5ea81b18e8a21f400c55218393cbf7856748bec5773496f0f",
            "1d980ab26a69a6c02247600337f884e54d9b15434eaba303ac2e44222627c9d96ff5e406ca15d9682c63d44c4d0ae60a",
            "d0c112b5fa0dcc781f698a2b545182204b03cd777ca3521c305aa8e203921352b1b09c7d308a02041e9dcad5ddf3a6cf",
            "7034a947960ade8a6a3586a8968837dfc3898b514f2e2ae586ab2acf5d677ce68f69efa6ec7177a19a6ce72956a434ab",
            "9bce878e1a814b9159ec349e951ccc705c7856c378c708acb105eb21285419db27c24de91600ab0329e196f4ac49b254",
            "a42f898b3629c6902a1dcbd7b3b18bf7eeeb46f7f4660a7cccad29c7c70d177a991e1ea651a840b8fb8a105f7b4952ae",
            "7478f2e15a6b2f2f8ab979d2d1bed02dd6225e8e2eac34249e0c89f2638f55ea0e284f0ace212dc78ddbfea837e2d5c5",
            "df55daef01e6f54e10cd503422458bd89968d78c5548a470627cdf87932049cf4676def29d579558053ea3badcc9eb67",
            "88e17274ab0aaa593aa406632ee6b0dcda1c3f4e723f3d2fa7190827eb3cac77f2f0ef0ffe00a35996a22bbac07bc8bd",
            "b30dca7133dad99aff42d7a082c6ad80f6c1ed165a696cb3b654d752c7e8a6ec4b55291277a53dddf149bf46b59f4c4f",
            "99a28118a8d18f2fc4c4dbab9a18abb593e5d75f349aa6beedde35a96f6082ecf0616485c85b386baefaa0081f9c6c12",
            "925ce69a5b432f874b79cf6c79ff96ff1e2a4d0df6830300a8043fbb3f7f389fa3a99195d19fef7f649903f5250ba266",
            "446ff84d6f84c7f79c4bdc2bc62691b79a2b1587ccfbf3b0d764a3a605f5ca8d586206ee10e07ab0e0d748109afcda48",
            "6709c108f651fe38f8513c035512f6b784f2f2d7928d5c298b975eb689baf7f91ae7a51395a52f894a7ee57993c17827",
            "0d42cc75e2d2cfbd235894618b3185aa65b976a869e22a4a912d596f0079e25259d02d93eb6657e668831b353a092f4a",
            "66f3dd915bb4fcae763d45215dd81677460c8857aa03a20168a3be5c3aa407909b8132ac67a75a38bd1457dc2ac6d4fb",
            "e4d1ddb8acfd52da8aca273bcf4bf7e36e373f447f3524e2cbaeb850e4f20c5579304e4e6826fe8a4d1f48c708b559ff",
            "66e3077d7e9af8608c0428af1db78edbf312f82395536d9abe29b8f998ef730f5bf6cd5d07544c0e1ffb778de5bd511a",
            "f6bb853b81e8eac3b1eb2f9f853381567751329836e467309952bc3847aa89659d76c796cc4099b175267416d7711c53",
            "1eabac7759b324bd3def313fa78f57e9adc192f5b2369c70f31c389ba32dfb30783840ebe05eaaaf0df660cbf699da37",
            "845b372c1c1a472cf36ea71d72aea51b4fdac03ab8f8fa99ce092666e87f43aea79cd17ff0fb440f608ad21342ce5326",
            "4fde2017b4a755d2ccc18f27efa584cf675b430a7aa483c0bbd8f36aba6545fbec0e9298c74cab4bd018c4844b54e857",
            "50b2cffda65c1bf26c6fd1a652c7cbe1f1cc6656456596f6d90f6e87403e9942c21a97f453cd7767ef87696faf6be88b",
            "e3cea7b5d64aafc3816186a7cbcc4e8fc76b6faf4c3755f34072edc7edfdd9158fc7bdc6ce87d9cc0e7bd363480ff520",
            "195349220e02f41df2c835a07072cba74dc5efd2fa487ae662ab108e9c8a918a864091ca49cdb69efaa927cd3e541b6f",
            "de6b7654356508ea1717196c7f96b304826737f1a449037329fb41a1e65223897634d45f13e6fef1b4d9e208489a0a3c",
            "32813c4f654f9e3d2444e83c694d700d542c84ff8ed54b386463f77dee54cd3e7e3a14cdad66f628d99ec11ed3504086",
            "baaf22c4e124c6438d3d22888a4603bb47c45d5a73869be8372b222ae8119c4dd20035638f9044a8cf34f1c838b535dd",
            "f0551f35896494f5d507c2f3dec8d0426c983211cab7186640b82cc74ce21a885a4a37e2d628591206e698d0a483279d",
            "33d80f8e82214628909aa4081a1d149ec1f0a3e41a54689e0858572278c27b750d493bf314f6d95fa72277338ada1a89",
            "0bf9ece824f4ac0d07072ceedbf56f2735d548caf4293b1d42ee85453cd9eb055e01e1c3e8ac62e2efcc17c2817966a8",
            "b3d08af27cc42b3d4056d1e5ee712adbbb0f10f593d4a94d27802faae3bf0aa0ba817b1f5c8a1d6a12188bf1492b25b2",
            "6b5f7927034a352ce287474f8189d12a62201ac27abf591c2992639c9f7b9dc4be5d40f942fe713ea7b64acfd9ca6219",
            "bd6a22ad752b0d5fd0777530a47b8d4563bf64663e443a64c9715fca25bf90ad277ec15925c206c72f5f0f85349a80b0",
            "084b5742cdce0ab67f404077d37a252dffc8fdc06ddd7f3c797cf085f29eb5254dabe9b97bd938247188ded7be26bde9",
            "483db57566e0bed3c6caa97cab3568bb22bbd9c254b35217373f9bdc0dc3d4e663f9bbda53e4e3efd7c90a03d2531fa3",
            "81cd1abc88da664606e84caff24606368d71e9c257e579e02575d7d0de60631235dd612a8d8b0528cce1428dd06cf36f",
            "97a023e56bf6344463b2b8976c0659d7ded685c1df522d1e31dd4b7819f3872ad11fb1ea6e9063079753c467a7967bc8",
            "9e6b28b05aa21b977e90e21d5e0e1234634815fcca9753ffce0318b2dde14a4fd511876a41fd714149a300608e5dd9e5",
            "ceead77859315b6412a62b02e776917ca5b0bd57ef5ceff0bea8d01933369e10d080d6c8bd34aee9fc3ec3ab856a6ab9",
            "67120d72238eaacc126940a5a03b16b10df6c7b8dab2f373874f6b792567a004b5db02cafeb4ec98693d4f50e507a4e6",
            "686c55f96f28547ce8550302aa8eef52d6ebf6bfeaad60d06ca0f889ec39d38ea0d902189508344e8f63d4e5f1a71186",
            "c96a9b863892bdf54b4278c306b464789e497054f56344852b8025ae3f8bd34c411f7d16866bb3764f1ef96158f2c58b",
            "36dc54c2a4316abdd670ed613aa4935355e28cb8e99e28452b4ba10145863e403517d5f4f85b51c0863f9ec365588feb",
            "1c492ceaa9732ec3d6fc0ac726c519bdb159c0798cfcb76dccef2871b06f1d1aabf97669bc8de9761a2065d05b7a3a9c",
            "e73d1dca671228f70c0ac95294d5c78d26d3a578b4e369ab294a8afa298697581dac4dd20ff064d38dd61f2dbc43d5e9",
            "9385a41897bff57298fdc865a9879f438c2b9171085e6211ea310b75a67ec4736634ff10a8ca9a4072f275cdf5d63235",
            "b3c367933e40a18f7858ba55ed8d7341bd6464e0e9217966f9990449ad60ecad266bbe1a85b7f876c8379a56ed9d3d9c",
            "f3de8d366c5dc82ca951684a4bc29d832b4b3ead03fb71cb6eeb805a962be14e73681ba299cf65d87ab5f53b39e534b3",
            "a47b2d2aac8c16c14632561355d5289a561a8811cfa04133ec9c235fed69379325e29d176c0313fc9cda33efc3b53d11",
            "58c0503b973094f6ea5b74cc58ddeac52bc78b40738557e1efd77ef0ca08d98d0d767861f49a21686cadf0269d961fe3",
            "ad5ffaceb69347967ef7f821cbbcccbac20792bdcf1912fa34d80d2cf2a6dc1665aa579bad082688af561e3ef8ab78ab",
            "d2be12989e03c6559629ebe7f6516e47eb2e3e6f369c486cbd7c7680e6d2313345bb9795ac5983c8a32d4ed89e006e97",
            "eb6d73402346ba83834657fedae5eaef0b8ca861f5f8e617ca29040600c0fe79ea3b38ba5a1a39fc3ceff314bf9a2ff2",
            "2274a0dfc10d2fc7c7a30cd787e7c97e46049127230ad49e9df8a25dc3635c2fbd95a01de34c7601ebb69f3f36364bbf",
            "c9ce41792b42924f3d24c4c9c979b79d7a6bccfabe389f97754451b314045ce4bdb0e456cbaa79c2aab34a326b57518d",
            "22aaa2693f529b3eea7b319d8e10ca56a403fb6f494280ec1290d172ad3088e233f6d83a8e09e882d5bc296179680d46",
            "c2e91e3a75c5235272d19d7b4f02729526218ca1076f5e22e6d9e1c0448be6ac0499e655a7695d3da2f3f15b493381c2",
            "3745337c91db4570772482f5f324666d9167a219d85abb13ef8068c68d861cae928ea26020bf5e8bfa3c30e5e11178b9",
            "c8ba5b65844c11062dd992e24a6fc8536a646bd2c683d2fb165974efa8efc25338887c9b17ceb7bb445c9205ca00c432",
            "7dab1f03538b82e3987541892f0b7769ead45979525d79d25c592ab7e96c3d8911a7513eb4dbc1e9cb5fa6424d73c077",
            "12e44f44c63183f9a866ac08865d575da90c3d58a4dba49e4472843bce28a0c287f1c10b4faa8729c35ca18bee58a3ea",
            "437eb56e3e684ed0e8b675bd1bd04a71031a88cc86df64b2b26293bc7a543ad0d6a6648804ebc9f130f0418ba03939a3",
            "d9e82790de693ba9b914d9c101cbd385f0dee39cfab9d120a10fae21f0216f17c0f39c77a363d9bae1573ade3fdcbde9",
            "adf1e59d24464f28e541001bc9596eec728652755e4cd93ea8e12585529ec8af5269ed05393e642d5b95c03cca593ec3",
            "b6c20e7979862e541c736657b5a2a5c0a0551d9f3801f1e6c0eec492fe5633867db3e13bcb691ccd1a36804cbb971e41",
            "1a6bb913fcb4fbe35bbaf740b84d461153502b5309481bb4b1c24141df703b6d77ba944467b25cd6d2d67b98fba77db2",
            "93f68339f836ed4cb836be67b6b31a3208d2c7e51baa2f3970d73ba1b4fe327e599e0638b7a686cd867ea5e5ca1a812d",
            "da08900ddee47e96f178f7d73c8fb0bff371d6be6b9e8a97d5e9cb0be51f3e6eaa676f64a8e0686f666df287ab6b7e93",
            "82eaa143005d6d5f7ddb73c0ac2428188382a8d93c05dce754e76b2ccfa9e06d6d2328687bdbe10a3c1a173957827d85",
            "83a49624052d2baa5402ac458503df3a0485f2a970a13174ac0719539cf51da6b1dc6de21c3fe14479fd7649f78e6b0a",
            "479e077be17bb7a1ea4ed323b70b66d610240d6eda8ff911ef118a286089d0d0b2c0a066811a570e60c04dfa0e1f2135",
            "72fb311f83e0add75baf2a5d577886b81f457f0c1d81fecebed0d859669697187ef3ab6c34b4d773fe69daff2efe6c12",
            "287eb5cea543e8a8ac3bf6f120f58126e2f32d0970512adebb1af3040c27f9a01d0c3ac52a4e77dc2f0c108d1750d66f",
            "017ac37ea466da70952e6e8cb9bb53187e5683cb5183e0fa0251c9d9f0b5bbd26493d68714d598f9883d97dd251fe93b",
            "da4a677944fd205c0d5206982ec36427bb69fc82111f9e9804389589874872038b268ceaae4e35acb92f3c96db0e3046",
            "90199eb3dfba7a435c89791d2038b8d3876535c2b7c7b800ce43eafe29223ae193dc34c4fb30a4354f2015dfe2ae9019",
            "3e689ed48210248ce0b52163c519aa9bec10e17a44d2f40fa16bfa19699633fdfa99c0145e7007d2daa041e727debe93",
            "d7c0028f44342ae953ec2728d69427158710482e59cb3ea1abecb68f7af7b8e803a4b304674dde48da2a3f006f22cd2a",
            "71b8a0470c05fa5bfaf9cd4cd2c673dda2be882d7fdab13f0942a08a85fde89947cc03373c47cb7a3a2172625cedeb50",
            "17c4ba8bbbf1361f733c8e92dbeafa52f0439c7634b72387c6103633a10784cba46380746fae08e947608986af97d452",
            "5fa012cefb44c4daae368414ef9c8d5aae58eddac22802be1815e582f75354cdcbe868f4b38116588b8c6f2899d83acd",
            "0f824b7b49f81b7dfe715737ddb9b26c407801b22e99b8f1d621ab72a1d4d5822e3551449838b90ecddcde9e3da956ba",
            "acaac7e099c7f338c420d650d95e308af67b33e2109087296898a1725a1b4c5cb11d3cc4c262729455636c177272329f",
            "9d8e4844fda78fd4ee85c673685a131baa079ce156584cf497a3fc555ec69ea8bfaf1219d905259640000001987fc692",
            "c14b2c393afcf9f1879f8d286bd82cb2a078e1563daf4282dee5ed980fb3e8a337e0a02fd90bc47ee85a360d6d9897cd",
            "13978f68cf79f60c46e7ac49b41eaa0adeeac677a5e1483cf6d8bc6db512e2c4a6225ab3dc1432c99eed2e5b25d62b8d",
            "bede7ffc7eecfeb6188112710b26b2ed6dc272eb69244f1a7aaf7e78404ba99c2fe76ac0901efdd98cf8742f5e543f96",
            "8a68c629166314757351945fb92a683e6fe46a5b0e1f4e507b17db691af0052d8e6248557bbb4b28b70e8d98af499b29",
            "c5485c38bfee471445de0a31eaebe58f45a4eb9d2e65a2f84d9b121f944f28aabe5d89f9d61661e3d46009448be7d932",
            "414c7e031b8f4677f37d9f17527ed06db2f9ec61010f6aa31a3a9c060d7668baabd26fcb8f8a06a3b995461eb2305dea",
            "26d74f49f42e6521d8d642d49e8f4c8e33b681f53df8ca6d8a003707ee6962100495883662be0cb1a88e860f1df41ddf",
            "c11dd490f0ae6cc9db7e966cb2b00f0ddd6bad2401ec398ebe1fc1b55213cef82ec52a60653e6597c3845dbc61feb882",
            "2c740aa35066ebaa789d62cf8c8831ffeec0255792134d74136bf80e638691a04ca9bbd8ea2b949b02b0e912bd15511d",
            "70a272139f7ea2c39748373057ddecc7687b779437e8c44e476f3120fc4f157013665d400fe79bf8d9f57c2515e54323",
            "5d8a1fa93fdb1cbb1066e9b1c2743eee7edd94e20ac8dce41a6350e08926c4e069af5082cb926701284879cef82631a1",
            "d40f18c328945bb18ea150d7192b13db53ec45e396fa90086ebe376025e1526df1013a5c40b62839eee39eb693d1b0f3",
            "dfc3b71e4667a73dbde10e32e39dca3351ed4beba06b247b289c5a47861d5c2547689d37d996efd1f8402e95ac3b6d06",
            "82e9aa734eb41bec9a71b7a476dc23d3cb7724431b33d739b8aec63518631df36345bdeecc5c5e4ce37c0a9b25ccc3ec",
            "043cf8b032fbfedcfdbcec47a16bdaadff0dc135b8994bf465616cd34aaa37a3820a0cbef1e48e275a8b2b5ad1d6969f",
            "387016b3775b1fe5c8179d9323885624a24c044d746a8a491376f9d9220eb08af79dbdebf3e71a30374bbd13405afb32",
            "a9fd7eeb9c67228ba8822e1056817b2ba2c53311122ddfd0acaeeb31f2c4d03ec4f2198a9c8b45e5db6139bb9b6def8d",
            "bd73f53930b468da70f732f4250beb9c59f3e3ddd9aea992530cd0ca2e775894a7827dc69986b2b26926bc0a8445a862",
            "a954fc01794c380af4ce5a8b713b31807b15c9dad842367b54f37becfa1a94e81207389a03438fb4114cc12214ea1c6f",
            "303c101d6f1777238367f650e84671ae78379705a6939c85fde35b689302e70ae423758d31482dfacf31e238d2360912",
            "a16f6185b2f7b13654609536ed05f02b94d3a942e7220013d21a0e482bde2b589f6f0b268e41c9b3cf5553dcd7bde7ca",
            "91bb9ea51ac34c3f4e130e540840b7da8d7dc89f49e4bd4f6012d43d92a8e558fdcc9cd8bb2424974ad205c48d61a692",
            "19d2a63755cc86912838a3e002ca3d18ea8f16d610f889c42f793b07db1636f60313df68b44d23eb5bcdcdf111993b2d",
            "bc15b2c37b7fba947b35ef19f85eb6e11a2ecbd84a516da94e3504122b30b450afd96c4c43a0a195a9ed5d8ccb2cc913",
            "4310cc349ddc5026cc061222db7513a96cebd1a8f3352d05e89ea9b7ca502e83631758842e6223ecc185217c26e1ca18",
            "aa7a42490754ab5e3cc32dda229deca30c95c0b406e73a87a0339ac748f03722eed5a9feb437aa0afb8d5878eb371b6e",
            "fcf036cfa799bfa5cfa895255ced2f2bca976449f86f1ddb3f2f837ef56dffb80a0a88fcfa9369041cfb1d235c594fb8",
            "214417680b9e505c022518defb5b21d59c16dad8d0f8f9ad5a99243e1a172cf014ca7d273f5b4d0f9247c31bc4b4019f",
            "356b62fa3c8d0642ea8826722cda2b0973efd65cf48d49ff8bcc77785fa102b699b1421427f1ea5d33fe84f751c38933",
            "9f211512aa1580bea272caefe7cc15dd0cf9dc7d22975616e09a4712640c3e78c83c5c74a2733cd9668996e0a18a0bbf",
            "19fcca80642022e61f77fcb633587e7b306b4fdc53a9ecf773257082172b6d2f2be25fd6ff9378fc69fb5fe9c7df59df",
            "1e32563a05ae3a7108df2036285c6d6393850eda3d59c579c0ca297499427a9ad7955b8b659fec8b67a1c9b9a0dd5bde",
            "7c1081c7966365679fc044d5271bbf2865970f69c0b1ae774185d45f13842fec2dfff12b2f638d6386a53d07ba719819",
            "4cfb99e9d15cec82c93c5fa00c589c8ea45c2316ffd2365e4a7fe1e6b3e5ec87dcec5ac6c864cb07c4ec111f7c12bca0",
            "8c1c30f259d8ca98e10baaa21a599dad2facbf357eb4352bbde6c5f035600080e04fe38d5c1c8decadf52c25cdac2d1c",
            "fee5e91f973472368b78089f57480f39f196b12d9d4b94299dea4b5250caf7deb57801fc6821296cd110661fafb845f7",
            "f224f294dc398f9a709fef81cf53c7b1587ce0b1b47a87d5ba31d902d39df683781e6f5d4fd36b11f2f7f3f8c061bd03",
            "bb6c162a75b6e9844fa453dc0f3330f9344d6f48063ad0088a6055318de7fc73bf40b5c747277851f31ff339e9012394",
            "232ac9694252d66f70f25b4d9be029dffdb0f2ce1ebb62d8cb67e1e8304633519efe542f409b18d417b1c5ca125140b1",
            "721bc7c6de65be45bf4af10d82ccb67e024afa83f6a77d7871c6ab675525f2412e818ec0123cabc3cd2deaf5e5e105ef",
            "8190f4394e450d689aa3830bdd411447bddcfbaee4ca0fb3b69d3b13ec6a8ed4521dcd1032cb01e4cd30a60decce3bc5",
            "4da63fd62d553a7ac0cd7c211446c0ccb09e42342a7d1e9da2169e299efcde7724f6416e4a0bdf42d8adc42b531dab8d",
            "c180c2ae9e064633a2064285899d467eac9d3b41e0260e4e78b7ab0993c7310ddf6aad303c96d06f278760adaffd8e91",
            "96876e033ac3a82021bbfc1b9d732cf1a389889873352e71e0249959b4cf33debb9df9d28b943315c9d160129ea7e443",
            "ecdfee78cd242c7dfccbaaf8e585f00feea367692060aa6766fd737fb922f87efab6d0b59f33446de3a4133e81580b02",
            "b7ef3bbcbc5c48a2d818c89c4656cac448f8639e760f7b799b122468eca42044946a5a8bcd44dfcf8f4e6cbe31db2e1f",
            "8ae9306f62baf4af390fb5ce1257094aad75e911d417bf85ac031b910acc87dc9f5131b7ce5477c836885bc7f34ed9f7",
            "b8e24426b502caae4ab09aa4b3727c3115f7887d94379efb9e7e6b8217f857e4f017a71d7f291710011040863606b2f6",
            "36b62feccdc6dee81f039effbda806fc310f25d14b873370ecb792bdb99cbf9424ba4b51749087c726d716b8794cd560",
            "2696e3d0589d8d07b1a7c415e272028a72b109ed956c796e91e4e4207d41b41cc325e3c5f54c7a6ae2b431b80084df1b",
            "a012a5611b059a4642b57909a55eac7468662826c44628d793ed9782372e49c7be3639951bea431c9d87a70831c14044",
            "c39da23c56205f1c5c9d978a80ac460767dd172ed827985c5eb41137462d0faee998468ee34f1046e6c798951e11621a",
            "62e81820d92814078cd18fc007a3297680a6f9b0a584a77fbb1fa889e930906d5e605ca0f7dd8e2517d86e73397e6d70",
            "ce94092a5cfe5defa6f3e9c27a8ffdfb92b7676594506427989745801c24c49d19bbe787826bb7294f6ecf4e8fc943ca",
            "25f6d943d2a58ed19943e825d7f1748b25f6c972e45267e7316f9783c3e35de5fdc7abb27b8f9d0d7a55650fc13c7526",
            "9b3ac437a9564fc0cfa4939ba3fb583040230ca02efe2defe3a727be7b947ff4cb5148abd19a7592609e3d4dab51a129",
            "3789ee5c1a85c2c4e9eb5df526f17af306f7f926a01b1cf03b1f5e10e1e02a7c715036e09f61f8171fd6b3b5fdaf0bff",
            "425a6cb5f5221d42d0f51b4f705f77c97e6abf9e2c0c14e94f30d4fa5de24e091cf382e43178a37e4bf4da772c98ad17",
            "560f04dc6221048502bed6cf4bde7797bba2a1fb9e06b568fa0ccba807f609b5306957f176a97e4b8999b4efaba8b2ed",
            "e36b96f9bae1a2e6ea3dee85cb72e011c679bf90f5e4fb5921f6fa9de7358e1531112e462bdccf91e5315837555ae4ae",
            "0dec6b01c3dd728d56b21d206a824e0b6100cf4735999e7a0b505672765cc8be7b0324714a0f6251c77f277e3bbdb313",
            "5b497f5f37ab0b9a99e2b73393e53cbecab7ed2723513df3a527c6cfa8dcf24db70ed63a9a62635d3c03bdd8f87a0e68",
            "f87be3a187acfaaf6a5752c43bb6503325a671023d7b034139b32c4461678e1e1c6fbd33df654a477672d83d90b3c1c2",
            "19d3168e4ad6bbd2fe7804849e77d3e802083670ea994cf197a629514f7536b97bbbb857e328ac01851e84bb0bc73941",
            "07dfc633b4d7ae53f1a3542ab5ef48c89cf10c1475c2ee338e04766d507ace674653f5db951029388a4e1507829af203",
            "451603a91abf5865b965dc2eee1c0df69e100e0e97aa3dd1146e645e8078379f9dfd9c31fc4d3fd6375a94f9c575335f",
            "123435e037ace24f5ef56fdce20f72d516a365d1f3732e00fc1e9ef89ade8ca2214469e03baf2d5dd0f2e5765286d708",
            "0ba4d8d4d50a6d182c0944ebf9992ddbe87db75a5950c7a76fa65aa54ad2eb90a48814c5c3cad3ecf3c7da66c2043ff8",
            "e2c7cc02d931c7a3a7cf4ef54e7ddfded86ea6cc3491c9c60d8b6db427336e1bfa733785b214f5144c3480e30a6b3002",
            "9f41303374df86cd77c2d4b4e90ba732903a2940211d0a23c73d9bec97981357d4735d163dc04206f98b94662d36ecdd",
            "c075405a5631dfa5295a0577b3bb67a0bba3ef911db8dc597e22359859cd9e3387d4aaee5d467133ddbd055cbca92844",
            "e257a4684aeae70a85f69453f5b2c967217c8cffa4200daa2cdb1faca64c32592c1a1bf6890ac837844b28348dc3b306",
            "c98c039383f501cb7dd228e4e9cc20b56af16dfff5ee0edceedcb406f5af7e38d01b6ac0dffd0ae2357ef39b88d70afe",
            "75dafaaf987fc9587f99bd7e2df87a02f5179ec0fe6d1ef5b5f0bfcea2fd15f4d07f08ffcd64da6b3daf278e7e2116f6",
            "efe78f5e6b793dc9f20bbef07f3fd7f25acbeb9184bf7c2188bf1d5078ade0f5edc62f5c60bebaeb782dfaba69fc0b73",
            "38bf6f21bfd6f0ba2dfb0b4d6fffbe49fb5acbeb86db2f10e1fdbefdf65ac3eb1af5171889ffa38af5b5c2d765df5f5c",
            "8de43f2a025f2b7c5d4afd020de97f5258bdd6f79aa8ff021ef9bfd3f6d7d2ffca91ff8a8d17e97f63ccafc5ff959efc",
            "1585547f11ff0759f9378f7d95ff7f2180f6fff03490950005fb2988f4f237ffe24120f43f9ffe07504b030414000000",
            "08006d996d5d1ae087de28290000f3350000110000006f72646572312d6d6f6465372e786c7378ed7b65509cdbb2f6e0",
            "ee123cb887c15d82bb04770d047777081adc83067777770d09ee047777b74bf63ef7661f7676d5fd7edd1fdf79a89959",
            "2f55dd553deb6959bd7ae424c1c0d101900068000040044827bbff940ef2bc820200d000d0a01a825696f68696f63a8a",
            "2ed686765a74ce16e68499e0a094e90050c07ff0ff35bae38665c088907b8fc85bbb028a76c62e754377a63e7ac98848",
            "4f4061e9cd9119fad0e71fcc73859498481de4ee906bc71beab51ef5d1c2c714191a842085be6f2eeacec7fb049a7a3a",
            "57ef42a6ad7308f92a331d5ce6e305224d8500aacaf878ccc020352106de413eaa83c48abc843c0bebdb1ca46de7f79b",
            "0660c14d367460f910dfde850776d915344b6e8f11d6e0e87336746037cf2bfea0bb3963131792ccac7f4c9b71550215",
            "2817e1d797bf5939072674cc65acf773079aa26d8732f8fdd81909bfb7280f775b9f10caaa70f6fd4c90ae07df507d3a",
            "ded7937c1562dbd68f28c1978ced3f577415c024271f049d64da3d31c3586708d66b517f75bd35406ac7b48b9dada583",
            "c91b69702f8bd78f4bc1597ed5fb46ccf2ea76e85bada3a82dea6bc8d81b176ee1f81dfaecea4a318fa301dab90d4d41",
            "a18714e284b04590da1a2e767bbf614cfe529df0e13bf5fc2e150603e2c4e445ee06c2f71d7456e5bb8d9b49a367112a",
            "b80b1e8f9afd699948cdf6af4facbda5bf5b6641f9f8d3937fe4537253b9b8b359df5bdec6a8cb9b3599e06a366c5dce",
            "d579de9ba7273080dcbf39768d123dc9f9f34aead96fe19e1d5bc7d6d0dc0e48f7f3fd3f0efd1ffc4471b4b46c173d7c",
            "c091d87d60ffd5e44224b280a413bf581391a2a7b7f80fa49ab02c52dc85360a3c64459c8f60a25e2d9ef75fe717b8c3",
            "76fb892e57998b4350d1b9486a9d4a02126e6d1ca0a8c96de3456c8a57ba511c272f262f1345255c281d5b22a88beaf3",
            "332407964844c64f842a9087e9547060d072f0023079127dd47ba5d516f5af4b412605140e4a02caebd8d277638693bd",
            "a6e3737710f88238a4f6c71231faedf43fcc96baf79b135ddce0282918739894eba7d97212bbe8bc376f3f134a6e437b",
            "707f2393d22f4fcc8c7acd382cda47967ee1a55ecd7c3298ead2cd06d355d47f3ce01bf23a7f11e4774e8287fac00b06",
            "0e0090c2020048cfff7136073a59d99ae95b5999fdcc7cc5ca1656abecc88f9e630f81944b0b33c167028a8d5fe14695",
            "4caf0ae32dae48b3e3c4a50658acb6d8df2a1d78b5d910c6c64f7eda23f0059f5feb1bfec6cd34709327ad664c4337d7",
            "5cdacbc9d1d5d85b543d74ba6b53ec9548de3f5ea35c496d20ac7552f1f84de361bce7a82db5ccec84b96c8f020150e8",
            "580ccc3177ef321e2703aab929b335fbc609bf9120d7a32408ab4549313c88b653c30c08ae3088a52d3b7fe30b94e9c8",
            "7f15c2bd6193ea4856b68e2974ec2b9ebe4477968b3ff2f51cbd3e3a38af94df448b15c0301ca2c97b95dea89cfbb861",
            "a004c70a6c3f40a8a809f32e03e3f35dd147a4e8181b7d6798f6a9b1e4d5297c7ad89b1d92271856681734dfef36f032",
            "829901cb459a1bd6c312ecafdf6fec57a4a8d0fb7a61877957bb965c23372365ceba59e67adab4f6c379b6c595f6266d",
            "32d128e5c26ad1a8d250a8d1b9b04e69614d993d1a09d84bcf6c4014abcfb06a2dd4cf0cde881863eb9497462a209f02",
            "f8a6a4cdd5b168249145e4c4e8a8c95d145550edebcff65458a60d440771d04b174ebd1a35ad2969475e5dc3aac4df9a",
            "2f2ed2374c70be0bbde8ffe261d2f190e87cf838f4817fbc5baa725629b5529f75c390a9888d43fd3bb67a55ad10d5c7",
            "31a1e540c1726123a853c3be9ad3d7e0dfe43130c03a2b735a1f9e7885967d8ea1a216c5323a9e06af61511dfa942359",
            "91456b8b7625fb43ae0efc038cf0207d825171e783a24704b348c590efa576c1d9e105eeb89c89f0f84ad42812b8a828",
            "bf55c8101d894a4e9b331523ef3bb95cba3ec22d7a4436b635c19b2dc5e515ee2c6cb92fb47e1241a8f92ea86d78e15a",
            "4ca2efcb366c15b8e7794234bc8edacddc892e556be7757c6f03131a7c50e8411b77f3497616cf28162e7795634a879b",
            "a99d405a3f56ce9468971d0527ca9b9022c8c4af252a04ac6ceb031ef27ef21be48323f5c957fdb1c1a27419454522c5",
            "fa087dd03aa944744c99dcdcd5ad455c20b6f755ab505ce957074e9f3d2e033f5ee79685aec603ebbda16002bb2236b5",
            "d9341e365f5b8022e4c985ec4ed9e1a19887c7b260e9d454b60962af23630bd82cafa2083bebfbc8354705739b265266",
            "5670542c8873cfc20af3e4dd3f65695c4d6e2321bdedb11df04b818de7ed3040d34452838b3e3430012c0956661cf8cf",
            "807e622682c977f57b87c5893e6ba33f4b179c159c4aa96435407df5eaf57e5cdae8d1a694e388a822dc4e5f6b2379d3",
            "ecc3ba3e57454c7ff2ee959501c6d5c2e117ec459f2b08be51f606be51fa5b779e16ded14b53c687cbc814533ccd401b",
            "eb460f6b37dc7b965477fc707a4fac912b9808187c3d3354c2c14f949996abc4409f00a3770922ac0814e04d6530c0a9",
            "9d5d831bd602064a464ef7d45907e1382f88c3a3c58beaf8c105d4f3ac0cf80c27a48b1ad1a6eb57bf8b177b329dd0c4",
            "cf91c4e83966e000a0419ee3c59f79f5af51e3bf732cc8738e05f93f0ef1ffe7688eed90e8a6470e38f23d92db3ccc2d",
            "69fc64de0825ac7a9cbd07e6673fe7972b3e20b4d0da9211e80827d9820ef449147f7c32881eb0f432953e0d4babaec2",
            "ab94868551df91b03043fcd2379d2c2a4609ad01bb466bbb1da5aba3a973b0381a848c1435864ff6ae07150a43e0aa1b",
            "4bdcbf243d6c6fa45762c39d4600dee33b2711ee3b92c171fbc1fb064874a6efeb72b6c781a798d1064bc25b6d355a4a",
            "26c14019cf1ae32e08b8803cef45d8b35c0b3407ff32066bf6e13173d9dae0eb76ba189e627172ffa86109632ebdb002",
            "b22077e9a0d75bb8f2417abe580add0112ade0310f4cdf562fcca7fd88c8b8e35f77f435a41770559dcb7f3854f18769",
            "0936606fd5c15925bb6b623910c4b34cf9111889b0e7fcdafc920bbd0df03b86d9714832d63fb3261e0c00c002fc4f46",
            "b2333634b4b703fef1c1f0936429311de1c93fbfd7b84b7cdaa58ef6e8b772fdd416acbb04d425858e47e1d8edb09784",
            "f3b3a485b7872c427a341ba825e0d150116d377d6755ee9c05e1cd1f2b501dc4d7e833c1abdb257673b3dbbc2c3ea6a0",
            "eed5973a07e248b8424a7e5146d2891765769b07655f85998baf76280562556542f45804f118b72964bdfb7133c2fced",
            "12d29985401e5f32d378083d21491bdabf21d0c7e10d84b91c1ebab1930fd676fdd9ad058d9b5b0e950d928c4d15dfb9",
            "cfdb49b76e19c7b154514a3209746098489753ead11d7a9e6664cc89ecea529f59e2c99d3419993f7f480d1c4fa216f3",
            "526d9b2e5c6f8f50e3c1139d540a23d8912dd63214f54d4ab6336b4534d2cc1a85de190cab6602156225f5ce1752b7a1",
            "f1035a844aca59f01dfd34f09bb3ba32fea4ef67b45ac35321ec8f922a8312b730485a28600ee2d30c86451bc6696c30",
            "a2774b49f0da3f880e03c35ee990e0c8cf4a1f3a9eaf1611614cdc8f347c71be47ac344965f324306202da7eb67d97ca",
            "a6eb5f9740e1bfa3e8d8857ff5238dedf863b288a924677e5f7628aa1d8870860ed77e6016984de05046a348c915df35",
            "8768465b14cea8d647fea373d192c7dfd61f359898ef449feb707a887fdc6dc63f763b76fecfdda6b9c0d75cda57a247",
            "ae1447d68bcb6a2c2c54e659506cb1fe16a90a126abe55febcdff8825d966cea54ea30a4a11f5a6f02bdac6d9da8e16c",
            "0116d985e4ccc7ee12e9a8d4319346026d5e1a1722b1588345a1af0a7d613c008247637556d159635c5510e5e0307189",
            "2325350edf3ea36850369c2822601cae4b24fc5895b7fa5cc0725c2c4e0d8779fbf886a6a1d85832e882053e6e4bbe5d",
            "20d80326ce820807d6f25ddbecd7b914939c7a4815db96f9699e8caf3828666a6277236fa9a884dddb0a6237f740ef78",
            "21505ead679fc09e1c7e19643e0252c7c91b1b2a5af29056299f19de19b01a69da670f3e967582a38c5b470eb58b2a5e",
            "738a77a41bf847943fd2e8d5e17ff3a0c33f3264cd68fb30d863ccf029bec5057dc86d549773f7bbf29aa7153dfe5d5d",
            "9771eb67c42e3e6669c660c4a6d2a00077813cfee67d8578f2eff7f791d6ccb269c11b7b6572dee5de2ca0123f9c4be3",
            "e61adadddc379a141da39a4c7bcfaac8b1d7dd756680bbd704eb4666288ca701601307a111558110583e6881a4c0ddf7",
            "b4f51b26999351e92eaa93133f1e60f9efe5f593a05f71f6a80a619b9f45dec270a5b39e58595dda03afbdacb07545d2",
            "16efe13d3dbce7a265317e3427ebcb1dbebd167558811776f83a325814a7679e47eb9e4b6b7e0e0954e48f292bf5add0",
            "8a7baca27ec895b86e9630d3beeb316bf589d0b47a6f2806e2b2f63512758173cbedda5ce9a0d916b334aea690bbacd6",
            "1912c475d98d9258e56bda2b76c2c44d51ee9ba0f085fcdd3b5ae3de5caebdb1df879a03a589fe1e28000083f867ebe7",
            "0ff2d91b1b5a18fef9fe4794d9575f30236d40f1425af7427ee0e55ae05932fa1071a3b9547ac0b2378b7045396c9c5b",
            "a35c87236ee0060da5870075850aad29a08d7a1a1f886c1bf9118f2892afb75cd2f686445efe4aaa8e3b69fec2d33a41",
            "f72e47e3ac68acac484269527ee9699dd1683b055adc65ecc2d7ad35508c71832989b671b16aa07eec3252913d51af17",
            "6dec4259c8f50bf7f58eece11a0f46d467c248d9758243d728f91ded1aea3cfad00d9eb46870cf643b1403883afd77a9",
            "30ee20ae06ca081e6d6a610cd2dcb1cb5813bc89c33bb1e97b45c0d2b2471b8f0bb7a30e575f948f6aa86417edb1c168",
            "2569aaf2905c5c70cc02a5ebde30f6e16ad31f1b8ceef505e8e202b1e2613db75d5b6fce0a910e554eb4e9ba271067a3",
            "7fccdc6fd836b4efd332db7dead93192f742afddeb106903a1af09c3adb97812f902b2bc839ec6bc263f85cd9ce4c9f5",
            "ad824cce716764c104820557674c677e15cf3ee9a1dc72872d5dcacbe30ebb2f75b45c6f0581b28badf698876fdfc914",
            "73d19bd1dd0e4ee7f585fa4a7f27ae5cb7b16d6d9865eddc8151f8e2dbcd9403ae9554a3712b043b671f5fe319a18434",
            "fd21b41a9a29fd7bdf10060073a49e305ab63255e21a95902d50493393288eaa8b0629877d0535e0811e9af07cd7b3fe",
            "860db3cf0d3c219b05a84b5c17c0fc9a34cf37b7708f07b4441be67477deaa4cc7d3eee1d46a5a45651af8e48bb84d2c",
            "9f482d69f99a22a4f28793d7fda497e715dcf951dbe34a5399e741ef20e3e3c59736f7872b2fcf2faeaaa6fe54da0f9b",
            "7bd34bee0fbb756df7d35f5a6d8d8f60dcd503648da54839895f6d5345ac808451993fae2bfe90d147d533155f86b6e7",
            "528c507c0f12a1fcb9c1da433ff9b8c19b760d9457b3b5bc3e4d48bf20df60b3ef9d390812c7070fba1448e7e22d6a35",
            "e74a29ccc3819eca50b74f270d66872032b45011f154d92abbf21c4141c2679f57845620c483046ed93c4582d67678b3",
            "262af7885261e96e34338629770d769907fcd2d47d14a10fc6b1ee2b6d51c055346a39dcc36f531ad25cc5862c4869e1",
            "bee1ad53f8a919b90725c0f6d1c5a1e155557620957fd2249555d657ef0dcb06950542af193393c837be4f83765e8c1e",
            "7f6d20662e53f7ba52c99cc43ff3035d2ad4e73346dbfe61489470058a0f494286fb79a7f002e85906a59ac7829665c8",
            "1e489eed03324e7c79ddb657c9681d4c62d68612c2dedb3af19d2b83817ef37b363d9b9feb2ba564bb951d0f607643c3",
            "1f6df5fc2f4188e07596b189b667f310213dd2d45981627c9f088bf6d15474793a1d8b77effcc47ca574aa528991f2ab",
            "052efd68af1179038b70833ff8c76085682545b4ef296fa7c1aa86866fb977e67e305c46f1c7445eed3189a540b58ff5",
            "75ab96dcb3f12b99cb39cae1d256e93c5d9a2f66504d005a36d1f107fc508744bdd39837f24d64e147a147931af5a159",
            "f00dda7814109a5622a9b5fffafd4058a31e0487a46829d8e43477730b1aae03e8f979712bcb81c7b2e768fc68ff2bcd",
            "1d5b638a56fd53c7a25150c200838f6d656afda17a72d354ef5d9b958563c94653c06669468399b16699eeeb9c0c2a57",
            "2c63ab4f2504552768473fae12157a65263e458199d47a92364f7deba537beccfed6bb6ef49a3e6c599e0f494a770513",
            "14179bb4fe165b6954dc4f61e635ec37ba4d81cdc257a8399a9db551d8ac6e85eb01965537d5a14837b4f387cc77929e",
            "4602c14a0c876cba6b54b88125d47119c925ee72671a0a6ce6af8c8a1ca4df0e2e06089930260b7de7cfac2170078369",
            "b1bbd453b5c7a39d537f9853d2c80598e296b5bbb1f9d65d944ee77f08d51545743d4fbc54d8dcad71aab1e3cdcf73c2",
            "8e892301d462d6f12c8c6a3af1a67936d24e2c768e63c2acc5d2169ba3f6aae476c4a034236c8daa2d3a3db57a379594",
            "71f94bbe719d43df9b941b31aee10b2873406d88c255951f42a7e57f55de2ba86a4f90b45abdf75d6f098b6e0ddd5774",
            "d709f80dd31ac0a444ac717214ad2e4f34893c4859553fc099e8a55f18f436b4574b96ef891c4f553f5a4ca4144cf2aa",
            "4f826e02eeb04f04b1d9d22df1dbd6fbc2af9ec03063c90f7591ab55ef57ee2783dab854254fc9eaef4893b4d2156ad1",
            "6519fc54b3308c8f86f2fdd63e12da5d8476bacc90c7609c3865316ecc2e71ae539f321899f89d45cb0f6d62ff408f45",
            "c0b7ababe0dfa1f332a2dfae85917dd5d8d9ea34eb615732154fb36b904799a4ac3e6a389e2acda71d07aca02eccefa1",
            "1e5c60b3b3fc90ba4f28976ab8fa60e615acce90282ccdb438d4c569b295c55a833483125b286b17b8c02adc6912ecf4",
            "e681d089c9d4c15f952d9efcf8fb8acb8d79f29756ef75063f856eb24199263550a7dc1aa908ff9303aa2f25a792e455",
            "4231379a6c396ad45e129b4143744ddadbbb78675d6bcdb8ba30e9d2c37034e624dbfb97e6fc5afbb6a3e61d967e55ac",
            "97226da67056a4537dc5b87a5535db992574d314bb8d2a6c4c668186b21663e515e89a161a28e14f1001531c5cdfbedb",
            "02d71325e39ce6acfde37d1d5bbedc85697b20b34262b42d6029edd634ea98b0b9d648948de64364892c23ddad6cbd95",
            "15b619e96ff5c4ea3241d6d8587afbf9a4e48870f5c267ba852aa2e887fb961ccb22db123f78c61e5750f0637f079828",
            "688935fa4497367b3e7c007f953c52e0778926a53d25483d3baccb7463b6be96611e7bcb767e2f76bc085cbd1c0e80e6",
            "c2110b758ba7859c49ade883355ee8aee0eb76979b06bf3aac543b3b8df69c958be5df67fdc841f5b886e79253181200",
            "40f833ebdbd9bb981bdafdccf7d98adf4de7df22b7beba451c5c8c8b431fac09cb96eb430197a724f5d79e1efbaa5dc9",
            "675423b311af1ded44798ee0ba97333232cdb0cec14973234a9ef36640988b8beb7633f9c6888db720e72ddb3745cd22",
            "7028515e11a98dc665d3f865279544b6f6fc4a450c45e31ed4667d0c5e719dd6209a26d47ddba1cfc2808c42f0e112ad",
            "50e882538e986ace31b8dce5ec091179380d2692c94e7a75c5430a2ab1ed3b0f622a957a61ac817bafb7a42cd383fe09",
            "660de9f1b20dcb8e93ba7aa467a33752268279a0f953fc119bf988b7e44235b3e426cd11a76269b85a3b63042074a386",
            "0376979f0f23e161be5f8979df336713127b3acbb959a4054d90042c5853bc965993a7d1c8d004c64130f8e49b756cab",
            "ef2051db890ecd8ae4d85ca02cc881a4016298876a4620d5221139a79dd1e38ac39cebdb74961e55971e3597164aacda",
            "9a8ea325e8940f6cb6f1ba7316693152ec881c850412d0f04ce03fa9192942e256579429c3f5e7ebd6ab7f263f7c3dca",
            "24a8c28fe5e04928ddc99c66aad04e7fa451ee9053a19b03911930bc6959e30ab254508b092ebbe63b9ea8ac582faf1e",
            "990ab2b12e1524c9767c3408b6df28644ac564321c02de3569233d9324d7df6b7e3217d675945933331c20ece0128414",
            "843f0795edb41315c3de81b4d4ca7b7f0bda55e9f3769cec557c111f32e4b7ccc5ba76275855a7f9a4cdfdcab714b0ca",
            "6339e52a4d5de047076cc5cb92d887998ba7b816b841cfcf70798a234e1caa16406a2ca0eae8d831b817276e1b8ecc50",
            "e508eef59acf349751145b682f1a7037359c22c67c3bf071676b28d5ab37cc2a2d9f39a089b910bc19ce48613fb77bb0",
            "05d8912a2dd996f82e7d618317feee718ff2b8657f06a959fb0b6233ff13fb69252534c5c3e235edd74cc4abed60e047",
            "59df740f03bc1be3f70def7d8b4a9435fc1bb591616abda87ccde308efaef81e107ec771b6ee2bbed2e7d5c9f30bfd5f",
            "1c37d6b3357caf606f6b62f9e10faa6b0e49c376d223432cb72ce377eeb328090945a3a368364c80753b56fb18c97329",
            "78595a090959dfb4ba800dc746cc14e25ee2a1aa7d96869313532a04d3a2c5436a1a80450fc99bab0b1ee944c3966023",
            "c0367a7793302340325652b359d8b1ed92c499edce91c7eb3a50d1667ce87acec21684aac103fc9413ce8b9c4f8d408a",
            "89990e093b0d52b7b40f0ff5a1fbe0014623ab07fb2100c972b5e9b2cc65896f80e7416c6be010e277f6755ea269b13c",
            "fb30e8af63e37b5b3da79fa6fdf7e2cf2681e28265323d7aeb976aa77277590014aa2fe824b45ea1a243cb36413b2379",
            "ba0ffeac6574e16ed3302b7970bcccdc577f0edebe2d2ebea58a69a7bc0c2828eaf77d6a905d6a001043bf69fc79c3e3",
            "2dbb0d3264317fb530b72163489d023fa79b69e09ec29bb12c5a88118054ae6b72877281c7ee5d74b57dbb6def55ac89",
            "59578f77dce2a7541d38f78911c1a32c42d56b5b2065a9a16401b5289cbdf3359a31d821d6370ba3dd7eb6eaa3368631",
            "5b86429c8ec077c4cdc3a3736c1858a39cc62ba2e3eb7aa19984ecf3a15f9320f83657f25957b3abe85024e24f668db4",
            "682fc9d25713da0fd7a575abe11f319d4f6c74b3db66360a631ac1d33a0fc017ebf00bc00928b1e851b83f96f2ed1f3f",
            "7cbd5e571722409205b5778044019772333c56d325c7a5667d8329c1b5169b36f1515106149839b8aca19441bca81ca5",
            "623210e24b7cfac9a2fa949a59b6d658cd62c4fdc4b2b94b588279434276de83e71d3e9d653f01676ac97a58120aef75",
            "b309a740deca68169a482bd9fce66dad400735ddaa18d4e2764e0a94b67cbc7fa0b2b012ed28ab90f831e283fbc1fc7e",
            "5e5b897e0aa8425dd6f6492e137b8be53a18455b7647e59a9ddbdc2e9fe551018a2ba84de120481efc6ed1a5a93282d1",
            "19e87bc56871b30c728a848f1c7bdda834dc974f6d47f836e6d11312e1d0a7c15c6cc41cde5eaea0228c9232da905f27",
            "ae11a1e7e613232dea12908326bea5d79709f6cd29f5c4b7d831998794eed8d05ea6f6b222247dbe5de32ff7dcbcf9ea",
            "05f63baaf10ace0017c1010071f8ff7125836757b2b7fbf3e30f9a7d53d1b09a6745f77c5dfe940e03719bbdbd991c3a",
            "e3ff665eb8e9c7c87bfc7236f54c0b56f482e85883b26dafa5959ef0945c0dce7d3d1731b1b490a10daed789676c4d9f",
            "5c35684957ccd7fb3e441de9291f47975ff43f1d58a2c0bbcb52774b087174917845e43c19dd9fc96ef594cf1066589a",
            "29ea22a5d073d08a3c8a14b8620554ce2cc30e30691b405662bcc2902fc7d82b47df0649516196cacc7f8543d16ae7fd",
            "e8ff1116c5b1fc15b062ab0883ea38f495b47f62e53b68a9285fd87e84828d0e0d942ba1d3b0be4fe8be92adf21f20ca",
            "04f6f623d1d9439a56188a18d00b8460fbd4264e6a83c2effb912e8884cd2d907833961fc278d09bf5c8e8ece8b0aab3",
            "727430ba609454e9ade3ec72094d55ca768531ca22ebad76019d25440364194207d2738fd778ee5bb88817c7bccbc308",
            "28aae1f7c487dcf1e967591041b72b84529a3a44b3f54cb1a6b215d294e1a8a916cef7a5307a3e945a99abdad3b939ee",
            "8a1eb1665d793d23ca9022ae3e383d474199ce0154caf5dc056ed20d49d0a1e64897cdef72f125496dc30c3de0f7b051",
            "a60a97276a8d105a0a40aa0f3fb3a2cef948d2fc20ae18e6d8a0e0f9c85d3ef228d6c005e31382ca758508efd7715de0",
            "5444d31c7069ae959152fe969cbaa01baa79ec51796dbaae1965a4dc2c40e7e9d4f4941abbb3ec9c7b8251812f8a048b",
            "afa55409315d9e15a020edb1c92f85f349dd802ba060aa56c3f813a63a53112a17c645bbacd25b3462db2f15b86804b4",
            "5aedde8e07c12c934eab69bbf57925f5fc1758751b212909cef7985dcc84a6634cb4fa323a9f35f01e144c07d010fa97",
            "19909b38de6fca777239104927ea6eb0f886de035d921033b98fc017d4ac08ec33699172cf9861361fdec7bf7fa84ede",
            "bc99c62cb9a72fb0573f9bc0650d75b73ce60a3c0046ddc1707e54836a2590f6351f4d861407ca8c69b37ec93ddb5161",
            "72e841d6670de37d9f91fcc1475d9bb6801930771c08d7dc5b7c8f71245b4ec37f8ad2e1309ac9e8114fe2e05fa8b0bc",
            "7ba6419bc6d3d6e1bd04468b3f52f358ad4647d54ae655abc5efd7373295ff981a6bba569fbd673e7a2817c22ee3d622",
            "83f92131f57b2de940657e0269419190db3648754caa5ec4916b3444663b403d180e7434f9ae13521ec531d871c84ab2",
            "8219dd5e9eaeba69257c3e8155c87b3e537e91645db4b58b5bb3d9680c08ddd447a63934952ef0a5be9ea4e781ecb8d1",
            "015168f3bd1f4f1c0396358da22e98342f643b00dce9c7d7f354728af0c2d57a2d9acf266fbae9cca5cbbd7457f6cf87",
            "1cbab42c243933cc363fcb07afed5c9892556465b2dba61adc57667cd234a7029fa28e25439d477b3b5ff79d03e1b8cd",
            "2a9cc45361afce3776c48df6bbe149dbfe34d5ee53ae8e0873cf2e88c17456ce02fb79a3fe5e51ba870a0e277d5dae1a",
            "3ed1fadce00a6456be34e6b23741587551849036d3ae225eb424f11e55b3ffe146b70f1ce40f5091f10e29a63196fa80",
            "154eeccc2dcafaf964a333ee34889d7438743c27487ecc9e9074becf0899943b57772e13b9ba518ff430575d941667e5",
            "62177b88df09e1c41e1b5991faeba10c7d0349fa97ef61a5b307ca594832d0f8cc601ddc023223ca92d591732d2be4df",
            "0a0714b36175b8e97770708bc1739b55f7be8537ebccaef3d3782c5ebae07938e5debb23688ef8aee380ac6ad722e991",
            "4e78d06319deba6c205da263ee1d341eb187cdf30845ff08f7bb78885156c8ddf2bca201010048fed6b1fdf332e857df",
            "f68fab20bf88ef3f2b0dbf23df23e44bcb3157e4a85c6da163ea0b90f8f7771f2d502b4ff2acace87c7d972d3e243c18",
            "e5cef1e5bdfb4175ba9c1a208f49d88f6d05356d52cfdc3789094e618f297a24d625057fe7da903fc4acff55f607928f",
            "702efbd4481ebaa344c58e89403c6fe460c2541f8c427c8502c8770c2ef9bdcc99fceb01efdb102a196541ee595421a6",
            "c0fdfb2f07f38ef4f6c79bac51927cb4b161d7fddf19510d44c53fc61ee1b705ca6ff6a7ef4a782815ff7ea80451e87c",
            "a9f97945f6cbfaff293cfeb4fdafe5c73f59df1234523401a53a7680542bde2f52616595fe0aa263bc9c6b9f8d4bf822",
            "89e224bb3dab0ac736d187d04c936b4ebbaba4613f9c903dd3c2c6369e99b65b2e6e5eda2607c6807355ef9c16b5a732",
            "8048c44bfed3f7b74c95f669e6b69caf8c424de05b4e8d0f6762e63eccf2fbcb42c6ea48bdabc31158d8244ceac95b7f",
            "329f76c8caa8628a4834bf47c9f8e64d444e1890960949358d50163d784e9b5f29a7b3f195edb7c60f37a5a7563e1bbe",
            "f29c2889ff34ded0d9ded0d652cf5ccac4d2cceedf9efe7549f3ce6a9e10b96f89e2808b7011d7dca95099c6b1a63009",
            "bd6497a0d113e00b84a1b6c6449e7eb8e411dbcd96cc5672e33aac7ffd740776963c8ac3ea8d84a3d535a8d681cdee9f",
            "eab817c0f770e372925a859cb397cd0f85e0e5df6e3557a6b3492f722c93e4769086cdaf697c4ddedb8a499c8fa08a6d",
            "34f331eb2b6c4b5c54bce71b81237e473d30e41214a5779a3034356b5c0216931f08956fb72cbc32462e19777358395b",
            "8d303f284f86d408021302bac337d858c373a9976e62d939da4d371896a987a029ad6bc660651cb443f69024982e9f54",
            "95a8fab9055bfd0b092d0ed637deac2e0ece79f506a5ce70296ac0d11dee9febf6a3dbeb65ddb52796394c5bb76d8f71",
            "a8d9bf916e6d80965df4a36a8037aee570da2ea837de5a355769ed8432f18156f8066982f318a451e6fde01839f744fd",
            "e8cbdb9ad00db32e2372036db314a74c7ba436020f81b07ce21fc584e10f1b2a7e12f1a96ae61899872c309e6047133e",
            "9b7312a9c952617024bec7e61db2a5be360113908f7e0d2ecd30373ea36488de2986ece183a8b853c56b32ff714a9328",
            "ef274282edd9a2f3a45315a9e14072bec7c403dcdfee78e5380243dcf3aae0f985fcafe247cfdc40d058cfc4f2e7fe72",
            "87f3ffe4b6f095af137455ddb82311b17823328f0f016176160c694c344c661d1b4d4204d2115f8bbbadbc55212a1cee",
            "e70014f40aa0d97a824fdf8323bb4ff4a6b343013df67574bbfe15d781d669c2327191a52e03b2fafc50d02a982185e0",
            "6cd0dd292d4bd35445ab5cad8465dcdb09db4e17bc88cf4d80e2cb429c7a99f44c1f6cc55d326fe5b5a7eec7df9af01a",
            "6bd83cf299b4a3cfc705140034c87b2b03395b2bebe7f2cdcad6f05fd39dffb9a806fc310c25d9458fd77be4bb94b1be",
            "98b370ec3fa62477924dc8feaed97ffd48d97c3c5a68cbd240901f79a2689f70b5e395dc480ce73da3cd2d8b54130785",
            "2b4a33994bdabdf3a3eceaf07e6f540d9840000c90a46e99108e5d5ab5bfc8b05a5dad3c2f8a0c632554082690d1d45b",
            "88a9e3fab3ac42043968bc70b4af1c91d86064fdb9b1ba084b8e3c9281b819329a1a3da265f777a16ea1c492fcd5d891",
            "37f2109af629e2781f8943297d99d25311838dbddfdb0749d509c78261f520f3c36b8aef636ade6f2811eadfecf96b62",
            "dbf46698661be5f2ff78c79ae561b3bd18b732a25a360d77214fd0b5ee74723bb62ac15b887f5a4c034b798c7d910fc8",
            "4445d9de3c999f7528733b9fe3f168731e2d915ce01d2f115d681c9c2ab15bf09c3aa05938dcde0c9e6ce2716ef2da3e",
            "38da313a7bc83dbbb5df5e4351557617ca5667b4aac6064db6b8077ba2185f1fdedb4502f9c608e1fcd0c81b5733f6c3",
            "dd100442388fdda4a9f4209eba748d036db17b4204f366fea138f2919f1bfa9162f51cfa040df9faba8f029e17f03b1a",
            "cf897272e73d73941deca727fe85c67ad6d6ff61f12fa4c45458ced323075c817a40f9b849f5bcf1ed4cc7429709817f",
            "a3df2ee118cf392b44154d38ecbbc0a07dcb8add746e714fcc231e449c27e863f9b4e0757bb12316bce7615085a9cfb8",
            "88b16f1b5f059f72196793ecbb90f005c85b4e9b611b57b9c8df81183953e37e27529c89843cfaf62b65678eaf054a97",
            "ef5976dd483e33c97c94c2dd0c6536a3657d9959d64aef9e1618e20df77019fb3e8a26c4a43016d86c64bdf18997204d",
            "1957ebb91c012888e40903b859b5c0adc4cde4f6764c19c35df5b42ec49724dfeda974a7d165fe4d5f59c98aa9af9dd1",
            "be8625e6a01d3a38ab988a75d5718c0ce2a3ab991061bc4603f7e935dfd52ea5c9b6ad61a70a8d63053b777b0332a759",
            "49434feddd39f1bdbe107c918968e3b4b9a76247403402e60d4bd5b800b71e00c1671d74ba00e235bc1a6663a285491a",
            "8c139d759eae769ee5586a4b409e3bc9cedc5166c0fed0689af6ca553251a44a99f62d39a513a181a24ee85ddaa1e08e",
            "157ae7d9d470b27e8b120a3567fe89626f407d8ebfe8bc3f6ddd578b74d7b49bf9edc3feae6392fec122b547babcb52f",
            "e7e05e0da7a73f670c7fd392f589126379a6ee97e718fd06f09bc2e2cfd2ea6fe5c51ff55573b49659273d7ac011d41d",
            "12d9158309756d9052ef8fb7d6f4018b179b9fd33cbfb1c2884f5c2e014311c8a19bce4eadd987ef3ad48a035a43868c",
            "b58a946456612850a80b06952739668fb44d4488b7f8e3464cfc39b047a3264f4ff44e7cf0afadd793a6ea29e56207c5",
            "d75a129936df7fe4a90f66cdefcc1dab857a2da2c104760acdc2d871133adb065fd9ac2f92cc48e1a8008b7a208216fb",
            "f461afd094acb486fb22ff1e0126e0a36851fa7cedee35622b8c7939cb6d8aa2456007875bda2bb757bb94fbceb7e621",
            "140653cb1933a6305a155d8be8aec63b378073d0d23c540145eafacf3ddcc1701eef543dad7258491ee24214642094e5",
            "7a02d594c9c70bc462424368417c82b8858353383d2a514d38b139d9ee71dcc60d831e4f0d92cbe7f17bae76b6117766",
            "2635defc6a898380be01fcd3af205ee237bf8978a9e0dfa7adff8ad9e788f597d9eb97822f27507fe111f26ff3a82f85",
            "ff7d1cedafa083fb43f81f86d35eea793974f40b9108ff3c82f4b7efe0c530cb2f7822fdf368cb4b2d2fa7127ee10ee5",
            "b7330a2f15bcbce0f88513ec17d71d2f455ff68d7f6112eff75de4971a5e76667fa106ff9ffbb42fb5bc6cbafd0221e1",
            "ef5b702f35bc3ca6fe023dd9ffead0fa52e1cb93df5fa846febf3a07be54f8f234f50b3414ff9bb3d54b7d2f6bf55fc0",
            "a7fa7be5fe52fadfcbe4bf62f959fa6f45f34bf17f2f4ffe8a3cdabf88ffab58f91b635fc4ff5ff003fe3f6603394908",
            "c89f82a8cf7fb3cf0c0267f8f9f45f504b03041400000008006d996d5d46759602222700008a330000110000006f7264",
            "6572312d6d6f6465362e786c7378ed7a65509cdbb2f6e0ee2ec143707782067777f7c1dddd93e01adc25103cb8bbbbbb",
            "bb0582eb65ef7deecd3e9c9caafbfdba3fbef35033b35ea6baabbae6e95ecfead5729260e0e80048003400002006c404",
            "8dd025810000499000001a001a54f3838db5a3b1b5a3ae929badb183369dab95255136382845260014f01ffc7f8d6ff1",
            "d2329d0cc841d7624bc4e798ac7434c882d28d14e582bee0f87ae5efc2957b6bde432e2c3233d216c9010393a5f17bd3",
            "af799f878c9834ab6b6edac47a1a2552130de5331575f06a0bbb0e0797f724d0e3da80e52bfa6f67c40cd87878929abe",
            "8129b48f47508de7a297e6914debc129ddf946889a136a8a2d85c3e1f0d3eafb9d1bbc6f891b680ca40b572820ce8fc5",
            "fbb6be9cbfe1dc9dd12a4ea281dafde061b0ab0c03dd8e0bf1517c3b64b2554b042f2486f427beb95426730fb27dd907",
            "f6f3da76f293e85a72dcf06eb4ae0b00ad615f7c7c062fdd6a936e4288f095dc5d031b6cb5ea29ed29bf3831899939b2",
            "a2c5e052022d76ed36ef5ac6bbd8a1e1db36a2e5bb4a5b24a8733f61a8ab95bab7fdcb2b63428d0b3c721fe989071ede",
            "a7930c5e602b6868e6669d0a621e3b5813657bcb444f3ad9fa6a1f1320be59affa109592b879ed453ab559fc85317a68",
            "582f3d3554bd7c14ef5b327b6a311861d3c201c1f92ee839bad75ade7de833d4dccc009d3bba1ab421d38d22edb5721a",
            "a1285da76859dda04981fce959c26d2fd3e34b4a3e3f8301e4fe29a3ab9519482f5e56522fdfc2bd64b4aebdb1a5033d",
            "dd1fefffc9e4ffe00f7c8b9196ed64800f3e157b08e9bb9e5e8e4216947411106b2456f2f6155f42aa0ecf21c35b6e7d",
            "878fac841b0026ead3ecfd30b0b8fc3efca08ff86a83e5db6754746ed21a9792e0c43b3b27282a72fb0411bb6feb5d28",
            "ced397d35749a2126e14cecd9154c575855992fdaba42293674215c8a374aab8306879f8c1983c497e1a3dd2ea2b0637",
            "a520d3828ac725c1e5b5ec9907b1a3293eb309f9fb087ca19c52471349187d0e06a6f3a59e7d96c497b7b8ca8a669cc0",
            "7283747b2e12375d23cbb69f4229ad688f9eb432a97d0a242ca8374ca3a2bd6f332f7d34beb39c0da6b975b1c37416f7",
            "fde8f7ff4c58b802f2bb248954c50939000300f6600000a497ffb85ad2bbd8d85b18d8d858fcb1e57d53b1b2d9e0407e",
            "f29e780ca1585d9e732326a3d2de5656fdba5c4cb178222e46cdcef876966c8413514ef97ed5061ffabb71f54643485b",
            "7bfad9a69b9b49ea4fd994a22915a9ef1065efb8383b1b7a0ec276851f39c65b87145866aa5526b80ade96cd0bf99894",
            "3ced6f5caf8ecc6abde7b17dca05e15f96b9c88b0bb7c1b6aa31c2b01a75ed348fa2a92b2c0cfceacab220c2905c2bd9",
            "63a48c4bde439a932d5c6283e269070ba94bf329acfbfd2195ca206da8dedb783def469a6467fee24ac5352aaaf40648",
            "e32673c0461605dced52c08146fc6d62242bd64516782d7bd7381b680be90d0a7c00672a78764e4e11095ad3369c1be4",
            "58cfcc1ee9330c1bb41b9affb01dbccc874a847695a081f5152ac6bb4ce51c9e993d80c80ffa1c944c95510c42ff6b01",
            "16539b94afbaa9f49f6296adbe168fe7df4a51c77f559895581820a3271b7cd700a7d862628329341ba82def3fd511b1",
            "7fd0d39d3ec22e3838efef9c970a120d75f3d1da60b2430d08cae8afc024afd84fa223b9b63e52a6bbad15f986969341",
            "2ba5d2a741908182660deb069674d8f36e6c45ae7e8a4b5ee6a22fc307d8fe58ef7af244b72330d9a534718b316ac699",
            "515194f453da4f0509bd04a89915260824f96128542ebc53b6271969b1cd87fcc5848a0a32d0d8e2eeeaf91ed90faa83",
            "50cc9d3408f2997e935a6c8a565b60c117ffb6b4889e987ed3869e74f2aca3af0a5966037bded95fde0c021ca2144b63",
            "e58b6b3e8617147d57df28c7f196bd284410e7b855df7df45eb235244cbbb7ea56c84fdd1f2701533f9719fa5aa86595",
            "271ba9e71a78a244a4ab87bde4773d79c65118db53766b279de89be8edebc199dce620e38f1708bfb96af97777d94df3",
            "52201abfb0e26c08b9321e354e8443b026088d0551cb5fcae5a807bb31c2b0a759b8873af6ee0a36ff52b6e4447223c1",
            "0594e2d2743833b1f8aa9844ca51fae600df2e299637a053a5667f5e5b66de1fa60e1ebbed4dc91bb694c71d422b5002",
            "1339d983f18c3d144b7709bcb8328d793ebf01881b094bfcbab92ae2f98480a30ad3ec0aab248a4a4366bb25321941d1",
            "b79acbac4fcfe376f79fce4324a5d4eb730547c087d3d70ec0663f2c80d2d6468378e8e2e1716d226bea65a7c5a0cf77",
            "81990493e271da9973d27dcce91247eaa1bac6223c7231cc591d2e3ccac42a641b1bf562cd6f9c7edc30682c97bc4b81",
            "bd2aa8f8746f7d3d8a6ed37e1f935698ba925688edd592ae9d51e8113fd4b25c50400535ab8f9f5a4a2855d0ae5bb575",
            "681b78d6e98b7b8a4a80b6d8a9824a34f8c9d2a8f1ebe71e4ae728e346e7ad397dc1720f31d7fa1115c2c15a201e1869",
            "ce6c1959e348fe7959dbe3f27df20c533d9a802cce00f2dda99caec7f7cf3ebf2d0b3585d812582fdfccbf94065c0034",
            "c84b59f86bfbfc7b71f8efad14e4652b05f93faee4ffe7f8162b20d9f5220a4ffd9f898796f3b59d50e61d913e95af89",
            "cf12224f00f192c9dfa239dd9f22602cd2d15c4e6c7a7c58f164706b7cfcca7a04c7033485301287852fb10faa89f78d",
            "dc1c1c0663809dc09ce0f80cea4e36b73babdac8cec2d425aa7418a1265541968998e5b195aa145e92dbe564b699ebd8",
            "124f2ca17b265324a2d2c70e6cd8f64e07b95527e746077dc8aa83703e6aacbb7db6544dd058af52cca05db0d587408a",
            "8633ccbcc08004097860ad22f3076d1d2a578f059105698edf515e34114c6aecda358fa6d18c3d5abc0d36ef62d00fc7",
            "3cd3af84aed2461311cf4937ff7a0cceac2159dba08be66fbe5fe18ac43be3471ec5ec01b363adfc652c65aee5ec7a22",
            "f379e66795e00e1fe077b4fa72f38543e4457531400000d880ffd96d1ccc8c8d1d1de8fffc60fc8359a9718b9f161990",
            "834fe1bca094d93d30069d538d663974ec73b3e212d2ad25e8d1dd5243e712a2e3d0f24e57a2275d6b61e0f4db210793",
            "6f77089f0ddf46865412718f8914998b818d1680536531a997045fdcd7baeda6a31abac59f092a33e48a325c307add93",
            "cf5337335a0024a3e49288b535357bc3702a8263f064a7e5d6654a1dfb69e92e3f2569808d26e756ea659d5eabf71b0b",
            "cb06aaad76c8b9d14148c5cf25c8cda3097d68a73d60b7504f612e41af5d9328655f6c2055522735a084034622c4c528",
            "747d56ce63d1c59cbb260ead23820a570d75b9d6ca77f351fe1a27489d637f489c94c6a25bfbdc46808d45bbdff09cc2",
            "20d0c57294cbfc0306a3454661cc427d1deddbbd845a6b50fd68de66c950f0c5f55ce48e215f18d764fdeaa150c07ca6",
            "031ff6257773ffd4a94c7eac3af6ad8f59b211ea1b676d1484dba8087fd7993c1ae9e4db5b627d9b741eb4c144dd2c80",
            "3e8805785cdd991655b59d1f9ce776a392733cb74dcfcf2a729c2d4fcd397af367822d732b14fb5304f7c2451c42d35e",
            "4a6c3fb410327a7363a9fa6d60f6747aa69bdaf4d4d223acc083824132341657879a108ea543e21dcc6216db998dcd83",
            "27fd8d8f0da69e48faca23bcb797ef428c2cc652339ebedc09ff8d68f53abc70f5c0d8604abcbe65314d931c269b7b00",
            "5651663fbb6afb5c8dc475b644ed57541d15ce2497bb0d56dd76d152d54aa65cdf263b63fe8fbc9cbb1e3796cac74df6",
            "9874f1d545efcb665c2141dcd73c28485407d2b1388892763cdedf86462c171e0863b2184ea63bedebfe967bc7ca537d",
            "dd50000006c91f07fc3fb9e768666c65fcd7fb9fb43bd258b620ab47f141daf2417ee4e55ee65935318dbcd55a2d3d66",
            "3d9c47b8a61835cbaf56a9c51537f48086d24780ba4685d612d4413d4f0841b68f0ac0278ee2eb2997b4bf255550b896",
            "aa7d9fbc78e96d9ba8779fa7f9b378a2ac5842795a61f5798bc9642f155adc6de2d2dfa325448c699b3999a661a5aabf",
            "6ee22a4a892349bf076de25245c83de3fdcdbeecc9260f46f417a228d92d8213f768857d9d6aaa0286b06d9ef41870ef",
            "14071443885a03f934184f1077431504af56f57046e9f7716bd853bc49a3fb719987c5f4a5654f765e971ea7edeefe28",
            "01eaa86f2fdbe23ea295a4ab29407273c3b108966ef9c23846a8cf06d49b3c1808d2c5876027c07aefb9b7dcfe2c423a",
            "513dd3a1eb9a429c8f599a7bd8b6af6f3ba26171f8d4bd6fa2e0835e73d82ed20ac2501d8e577df92c9201b2b68f9ece",
            "b2a93083c392eccd3d54f156ce797f6c1908c18aa73ba1bbb881ef98fc586ebdcf9e29e5e3758fd39b365eaebf8e40d1",
            "c95ef38387efc8c51c73c597c9d3014e97f05263bdaf034faecbccbe26dcba66e1d82462857f27f5987b3dcd64d206c1",
            "c1d5cfdf6c4e2831dd6004ad9a7ac6e0c1ff332380254a5f182d578532699352c89e5e592b9b389eb2931a298f631d35",
            "f891019ae8e2c0bbee961db3d7033c3197955e8fa436988590acc03fbfe89007b44407e6fc60d1a64cd7dbe1f1dc6656",
            "557596fed91f718f4421894ad29af0dde7ca25179f87691fef6bb88bd3d6a7f5c632efe39e41a6a7cb8c56cfc76b1fef",
            "0c7735f3204a9dc79dc3d955cfc783dad687d98c167bb353184f8d60593329322e12ac3dcac87590704acba72da52519",
            "03547d73f13568476ea54825239048952ff5b65e06293fea7d69364179b55acaebd2850cbe161aeef4ca5b8220719a7a",
            "d1a542ba7edba55277ad94c23ce9efae0cf3f874566f710222430315994099ab7aa0c0191a2afcf3cbbad03a8478a8e0",
            "1dbbb748e8e63e6fce54e521711a2cddad56d628c581e1014b7f60ba869f12f4f124f643a53d0ab8aa660da767c45d6a",
            "7dbabbd88815190ddc10fed6bb407513cfd044d85eba7834fcaaca76a4f24f5a64b22a061a3de1b9a0b2f4d09b662ca4",
            "0a0d46e9d0ae2b31938486629632b58495ca96a441d9a67469505f7e32d9f78d42a24428be334d16323e2a38871744cf",
            "312cd5faf1c1ba0cd90bc9bbad5fc685afa0cbfe3a05ad9d59ccd65842d8774f37a1637d302470f1d0ae7be74b5da594",
            "6c978af331cc4158c493bd7ed0150831bcee1a0ef1de7c0122a457ba061bbd18df27a2e22334553d9e0ee76f07f78162",
            "fe52ba5569244885df05af02696e1079438af13e9a06c5627fd64e8e6c3b54d94b87550b8bd8f5ecc837355e4309c244",
            "dee806c6bd43758cf3f7f82e79681758b290779ac7ada3da71bebaf88d512d91deba914e2078490312f55e73d1c43f89",
            "550085014d6adc8f7ad93f74fb4950685699b4c67160f85858b30e0497b478f523f03c7f67171aae9ddefbcbca6e8e13",
            "8f75f7e9e4e9d100f53d7b43aa76dd73fb8a4968623fa39f7d655add89464ae34ccf7dab8d9573c97663f04e6956bd85",
            "9956991e615e16a53bb699cda71282aa33b4d3a5eb24c51e99a94fd160c01a6fb2a699a11e06b3abdca19e2d134286f0",
            "35053e2429bd754c503c1cb2ba3b1ce571f140c53942d821ba1dc19d222cd43cad8e9a681c368fa2ad60ebaadbef6148",
            "b7348b272cf792de26821f95194fd8f53629f1424aa8e2b3524a3ce57e6a2ab25b6299143b49f30fae040b0199528486",
            "05b2ab093cc1609a1daef4d51cf16916341e179435f301e678656d1eecfeb597a5b385a6617aa288ee1749578a3b07d5",
            "2ed50ebc85052e38b1f1a4801acc5a9ee5712d17de74ef069aa9958e494c98cd389a6f96a83daaf9edb1284d08bbe3ea",
            "2b2ecf2dbe8d2565dc4192b4ee0be887d3726366d57cc1654ea8f5d1786a0a23e8340258e53d1fd41c099237be1f0eeb",
            "af62d36da2fb8b1eb8d00f61da0298954934cf4e63341488a7910729aaeafab9927c0c8a42f9c37ab465f99ec9f1d50c",
            "62c4444ac124af7b25e8a6e04e7a45109bac3d9286768d8a06bce9c3cd244d6ba336aa8cd61fa6435bb9d524cfdfd6dd",
            "93256b672ad6a0cb3206aae560989d8e14066e0610395c8675b8cd91c7629cb9e4306dcfaf726d519d339a00037fc628",
            "8cece02ca1c721bc71a8ad10d8a7f33161d8ab8191c56ae8687199f772289949a03e302ca04856d118379e4c93e6d389",
            "a7afa02a2aeca61a5c6677b0364d3b22924b33de78b4f0f9a8c198242ccdbc32d2c905dccd61ab469a43892b92750859",
            "6613ee007e74a17d24726136770a52634f20ff31bcee766b9992d1e2bbc518a8d8f57650a6511dd425bf5a2a32e8ec98",
            "32a3e45c92bc4a28f6568b3d4f9dca4762277484ae5167ef00ff67e766139e1e4ca6f4281cb525e9ded195a580f691fd",
            "b865bb756015db9548ab399c0dd94cef373cfdaaeabdec12bad977070daaeccc1621c6b25613e515e85a569a2811cf10",
            "c1339cdc43c3f6f45b4992f12e0bb64109fececd19f7e13a5ec86c9018adcbd8ca07d50dba4076f76a89b2f142881c91",
            "35a4fbf55d7e5961bbb1be166fec4e20b2e6f62aff97b39253a28d4bbfd966cac8e225cf5d39d615f65501f0ac43eed0",
            "8f4f7ded60a2a025b6e8539d3a1c85f0c102550a4821c3128dca87ca90fa0ed8579966ecbdcda33c8ed66d023e1cf891",
            "78fa799c00ade55356aa666f2b39608de8a32d7ed8c107c236b7dbfac05aec340707cdb6bcf5cbb5d6dfeefa51831af1",
            "f52f8a5318120040f86bd7777074b33476f863bfcf551a365fe4476ec1ba431c5c898f471fac0ecf95eb450157a0200b",
            "d2999d18d0a9e433a996d94ed08971a1b840703fcc1b1b9b65dce2e4a2be1525cfa3ed17e6e6e6bedb49b93561e7fd9a",
            "c7cf3ea4a4550c0e25ca2b22b5ddb0669eb0e6a29ac4de5658a984a164d68dda6480c12baedb124add887a643ff24518",
            "9055043e5aa21d06fdf59c33f63bd7045cfe5aee9488029c2633e974078386d2c93b4ab1bd7b2f124ad53a61ecfe071f",
            "7e32d6d9c1a0448bfacc04d9fa35e7693d7db29fe3b752c00f05a0853302913b858877e442d5f3e4c0a6c873b1743ced",
            "fd090210ba71e37e87ab2f2751f030c3d762be0f2cb94424deae721e56e9a153a4c1cbb6ef08653615a835b3b4e8e321",
            "18fd0a2ddaf734f691a81c4447e645f2ec2e5196e540d201b12c23d56390ea51885cb3aee8f1dfc25deb5a75579fd456",
            "9fb456974b6c5a1b7fc448d0a91cdbede177e5add060a43a103b0b0926a2e103e13fa99b2841e27daf285381eb2bd4ab",
            "d3f8427e4238cefc415500dbc99b48ba8325dd5cb18de154b3dc29af422f0f223b7874c7bada1d64f56b0d26b8eca6ff",
            "64928a529d8246541ac8f69654a824fb8fd341b0a30621734a66e0e867f0ce693be9b964b9be1ecbb385f0ced3eceab9",
            "d1606127b750a4d0370b50b92efbd1b11ced48ab2dbc0f77a09d957efc936fb1128af9902187b2576adb5c60d55c1693",
            "778e2af9dfc1aa4ce495ab3676829f1eb37f5b93c439c95e39c7b3c20b7d79862b501a73e154b3a2a7c2a6571b9ff801",
            "eec385d78a2b3352398677b3e937cb6d12cd1ed683467f9016f12ed6722fe4697f7724cda727dc26bd9025b891a508bc",
            "09ce44f128bf6bb099be3d4d5ab235493e73799b17fefee990e247f3d11c52934e066293c033c7792505f4bbc7951b9a",
            "816cc4ebbd8ff401b2fe995e86f8b76646f546fec5252a9a410d3ac830353e94fe96f144f7d77c8f08bfe3387bd7355f",
            "e9cbeaece585fe0f8e9be9db1b1b293ada03ad4dffa4bad688346c070332c45af3da9b8e23566521a1187414adfa29b0",
            "2ee7ef7e260adc8a3ed6364242b6b72d6e60a3719173457857f8a8ea5fa4e1e4c4948bc0b469f0911afb61d13f172cd4",
            "7e1ceb40c3916027c03191bf4d9c13249d28a9de296adf734be6caf5e42ce075efaf68353b71bf60650f45d5e4a1ff94",
            "17c18b5c48854086899909093b0b52bb7a040f65da75fc08a399d38df3188c64bdd17855e6b6cad7cff328b6db7f02f1",
            "bbf83efa1874b0bde430e8af53a391bdbecb1fa1fdf7e2af53a3d2b0750e03724bd9f753bd5b19c02c310cac9234f48b",
            "f25efc700d426308aa514eba28ada0ec755d1d2e9e2dadba400b856df27ef5c173537d85e01ef12849a9640d27ec1d0a",
            "7898a27cb5674a7b06bd3143083fb0837a026684e9a607bd69b8d54793803a894089d80f648d8ab515c7f2fb4a9e7ea1",
            "59a77b36b9416f5dc58ce095508a6e3d02b5382264c3a920bab078a1734cb4046bd8be41397fa4688bef4c6c42e79271",
            "3e92c170c349c4a4489f80b2e438122a867644425649dba791766859a7b45ec5c09c3da32a2ad3848da8de80854f3923",
            "f57a2c19d2b26cbb1be5a320d51c5057efca4073d42f6adbfa77fda2ddf2795563a7387f605b53f797028194e2a3fcb2",
            "85e838178febf784583da6ab2477753e0999254ca1446b3dd401f431b6caccf5d6a1533d49b94785847157a8aef15690",
            "ab83a825f929a85079d1a0f825b6e50b6bc272d929b790c303ac0de2bcdf5c5273a7e44a87bf2fab924edbd90b3baf26",
            "5bc6175e067ee245775a5660438cc449c30e89b04c3cf40823f9fec5103d0e7b1320fb99496d788343cd325b994913dd",
            "4558deb38d413798eefeecb0083bbf6de698d2bde9005e9c1eef96688fdeb862e9462bc3f96e74d1e35b5e127f4cb27d",
            "1b63c04e51f37c2d88831b0817c57ea1935c4eec687bfaa581000bceddf3ea75843b1bf0804c16fcb1230d2fea7d1b5f",
            "89afdf90e0d86cba1dcf11483f173b5a11d0051634690127ee5af948bdb170aec2c53cd862fc2d4f7a7729eb5c8d00a9",
            "a20f0df8aee9acea13deefb886e23fe3bd020e0088c0ff4f2e19bee492a3c35f1f7ff26c4845d366910ddd9bb0fc790d",
            "06e22e57acd26890cd00abc6126bd9fe43ac53db8ad9c79a79b95c098588e6ec67af102c6afb04ebfef290ed7662373a",
            "aec453cfaddb462f716e6e1ab275cb2361d3acd9769a5d08fd948aebc3e50942d7375f91a991b7e1239bc112561afb23",
            "f95277733c3f073694c84570af05f6cbf35ed54828e71045dd62268aa59fc1324795d45344fb97cc8107f123ad35e50a",
            "b0c3dffbd96d7042ede97250a0e88b7a9ca17ab4e8c4e28fbae9247ea1c40b6a43a72cf45d8a912068de96d7757cdea8",
            "20ab5f8b299b561f196ce3e2268718c9d6193362f523c1105b37c3264f7b936d487f01ea2586b2381730daf77cdb4300",
            "ab53058a858305bf303e41ff733c8e8a06bb6e804311b9ba58f536fb53a2d8d26c3cd2ba4ad4d627b96c44722b1b7bc4",
            "962ada107747de0d7a5614b588e6f7120492514d35007d5c78ffd2436a32bcb7a45f11d11703b3cb8bab10924f939113",
            "422730064e9dac3406c20f72b4f07869836b981826e094a8bd45d1bb1da42a3abac597a907a695f48c3c6cd7aa998ed5",
            "88a3813487866aa31fcb55380b6a58f1f6a421329d7abf278c216aa9ad237efcd043c9b78fb3e5177425fae9511c5a0b",
            "7daede458f08be47d64b4d771de1bbe92eec76202793bcf85bde8deaabdaf4a6e6ed6adf606de687ab02f649f6659487",
            "ae5dbf84a4bb8ff0ed37fb4e794c80f84a902fe60755885db95985acce685fccad1a03f2cdb258aba8a215e5911e9420",
            "588849ae711a19d7db2720e0cd66d3d5b4bddf9d244578cc7cc2534aa9a3d5356a6a8bbdc6ce07918f4dd74771b8eb58",
            "cafe91a170c83bee0b4f005691626842f7b649156b220e46254d23aae29af5f3a6c017e38fab29d758dea888cc3e35a1",
            "8d4f24c7238d9d2c27ddea5725c2d465e4f1b9acb28ff2c37d89b83eb6734fb80d7c3e38f4f9de1329c14f3412c1dee8",
            "41cddd99ad6a3dfd1b29490d9ee08c2922845563c355fcb4742a4ab800abdb4e88153d95077fa47773f144b7816b0d66",
            "f249bc83510d82db94b697b7b498a3e9ab7a6da72809887b1a3e2a652ae8ab317c0ba5441de4df118abc498316b8e8f6",
            "3935bff3cd90844f08a006e99b8bfe9447165e5e67808ff53e821d4f776b3b72b4e9dbf488308aa0186a9d59f7e73fda",
            "a929d78f0bd4ce6e765854a081074943873bb6a933bc8a0f476eaba151de45f961384a4c50aacdbc8b15b8bc095a6a62",
            "21973daab514fdf084aab6cc0585ebc0d4913e801d568b4155ee87c90ec2e0dbf6a696c92ba7ff2b69d1a442bb9005b1",
            "b3be85fbe31b38d3029752ac1303cd71ca6586de58099682d9c3eaae88d04c3e3a376f7545613a6774500b8c8189502b",
            "18bd63a04d3ae4cd2a7bef78cbd74b0f01d83d9e849da21bc2ab43989de77c6d1196ce0310c3d9ecf8658e1f0d1a871e",
            "995ef2b8740caaf9eaacd0f6cbdbb65ef25326b0b6fc78a4c70768d1b3615c18505bfe7de69048591e9c4e73fefb3575",
            "635a0cb38e8e08a5fa3532267b9c4fc797a69dc298c38fae63604fe1317e7cf0397919883b010c47f9c0e420ab278cb3",
            "621c0e848e3ae90e9219596e96101aa2dbc38126821cde38189d70a2809fd42cf1e49d33a1e0320fdf28a78e71bcadd3",
            "d9e289a4960df27c3fc22bda30faf776c538b886633baa9805f861b105f68fc096de547ddf75baa8fa4973c9f58383ed",
            "795a5e6ff892eecad7ef3b5623d0b3bdf740826f16805240384bbaf0fef99dcf6f2fd131ca8ade37bfaca8410000d27f",
            "e9d8fe7527f0ab6ffbe78d4060e4f01f5223f0d4ff14f9ca7ac21d393a5f47e807d5254882d17d80156ae559818d0d9d",
            "bfff9a9569e2a349fe025f81fc12e5f95a5ab00226511f8e0dd42cb08ea5771a13fc9d23a6e8a958a714fcbd7b7de108",
            "8bc180ec12929f703ec7cc5801bab344c53e503081376a3071a6174631a142116418835be1307baef0a6dff7ee33a58c",
            "ca87f7f3a842cc21470f19c78bce0c8e3f76d8a225f968e2c26ffa8699500d45c503e24edfb48628ecf4651e4878297f",
            "fb7dcf1051e862b5e965f5f657f4ffa33cfe8afdeffae3df45df1c3a563c05a536718c5423de27526163938905d13e59",
            "ce7dc4ce2d7c99fcee2cb72da70ad73ec98fc8428b7b41a7b3a4fe28828823dbcace3e8185a64b2e7e51da2e0fc6906b",
            "43ff8206b5bb329858c447e1d3303f73a563baa53d1796491810bef9dcec642e76c1745e204816324e574abe16577079",
            "8728b9bb60ebd972d62927ab8a3932c9f201256bc897989c28383d1b927216a12c66f082a6b0524e777b80fdb7c18f36",
            "66a655be04be0e060090fc15bcb1aba3b1bdb5bea514d0dac2e19f9efe1260b1f2368b44c8bdabef8eb98956f02c5d8a",
            "54a89dab8b92d14b0e081abc01fef43054b698c8b38f573c6207b992b9ca1edc277584cff7603f53c671d97c9170b53b",
            "07d5db713882d29c0f83f91e6fddced2aa90f30e7305a0107c82da6c16ca747718447ec8247b1ca7e3086899dd90f7b4",
            "60921422a8e198cc05e40cc036c7472778d30a9e0a38eb832197a028cb6bc150576f720b5a4d9b12a9dced5af9648d5d",
            "311de4b171b598609aaa4c7faefe409f18dc15b1cdce16914fb57a1bc7c1d966becdb84635024d615b3d012be3a4f3f9",
            "104982f9ea594d99b2effd8796a02222abe3ad6dda8d95c1059f9ed0b4396e254d38ba93a30bbd3e7447fd9cfbb6a432",
            "a759dbd6bd094e75475ae9967a68d99540ca7a78b31a4e97bdaf7566bb1b96aa2d1d50403f68c5214820ee53a86699ef",
            "a373d4c233d5933f6f4b6217cc968cc82db4dd6abc0acda9fa183c04c2da5950343346106c98f859e4a7aab91fc83c6f",
            "431208f6b5e073b9a6911aad1507c712baede491ad0d740898e9f91836e1d28df35f84e408834becdb47535171970ac2",
            "b7419314c068df672282bdf9e28be47355a9d11072bea7a463bcdffee2e7de4ed0f12fabaf2f2fe47f881f7d4bc30f66",
            "fa40eb3f7edff71102b0a08cc8e0a7fe3ed00e2ea6b301a164f3017b701720c4dcc9fc4cea4c0217d73d9f1383b3b31a",
            "2e1f46c49b0b25682cf57045e4ba5a68a233204f6a0f4fa0259237bbf33a8cd63e80562d5b7f77e9cf08f8fcadc1af9b",
            "dfcc8a87f813525c1abcf1db93e19ca9238dde3b2175ea7a6150f94198eb4dd7b62508bd6ad5e93663508c7b28b9fd44",
            "33d627d123acdf8530a8bc6112fe42daf1976a8602800631b23194b3b7b17d916f36f6c6ff18e2fbcf7d25e0cfd117c9",
            "4e06fc9e53ffd5acad95bce52f6113ca72f3e202b6d9ae70fdabea96933142bbd6861f0490a78a5d91cfd6e028f6fab1",
            "1e98ecee58a51a39dffd245b89792febfdb26b9c15b86cf42d40127706aa466bda4641a492d36c85536994ce1ec156e6",
            "92a44653accb09d34699717d49fd013c18176acb165340a60207e2a0261248179b7df83ad5a9368110fa91061b8a7dc3",
            "9dc4207b4885c556628f1c08a0e31a418f104011416fc71993091162e167e212cad724190009330825822c257386a37b",
            "baa242ae78b8142c0de3d894ab56609e243aa148196fe1703cedb335b3644b0932a18cd43f6f7475d2f7a6249c46749b",
            "c18040bf09735702c87f8476bcdf7b3c6ab8b6bfb9b23fb73b2be24e6cf0c8e78e6c70482ce1de6db8ffd6a8dae051b4",
            "d8ab72eafcfcf894555b5bbc781c9932595b004618d956b9414a32edef9673893c27754878b32915cb5756cc07ce3088",
            "be767978a7a02251e570890bfb49c66ec0597b91765abbd76140a7be3241f8c98e8f3ad70719c9ff39e90602daf3fd15",
            "4c802ce0771c2e98db81ce7c21e81bb03fd2f06f1cd6b7b5fd0f857f2135a6c2fa8f5be26b8017949f871427ad7f5b26",
            "36bacc677818833619e760ae7921ca18a251ff65469d3b369caa2dab07124ef150d282b4864797e7fdd95ed2eeedfc1a",
            "7530469c535836826d1ab0022fe052a5def5e78c022c1ac942adafdf5ce4ca69c99c295b5343d52f404870c57084eb94",
            "ca310cbed8a65fc09ac78b1871db53115a0429f5bcbd563666d1726b2e6d27e61bfce9808807ae2ae8868be8f74d3695",
            "e5a6352b6f9656d75dcd120c44f28c07dce2bb20575e132bd7dd16f590b7258b5a7b483e5103e79e134b5d7cfd5aa7c2",
            "507525439e9e3a650aff8fc8e03ac8693aab7c292905160e207f4ea43a532bd0ca94c643094623d59f950a773baa7ca1",
            "a265ef3b024cced269fa75e3ece3e4cfa7fd3c8c463eff493d1b67fb9b61b9141bfb2584280e024b307722a608792c48",
            "dadc46f37c92a700cff764a7d8ee2704a627b3f21f9c77cd7032b01b37275779b765bddec42dc6d8c9b1f73c4a5cebaf",
            "97a177dc9b47a618b428a0483484771e2021d8496b86d7202495c6690adef2feac1de70d0c79aa1e3cbba5784e9e7c83",
            "e4eeefd364710bf93b567bebce56b1be7036e3a532d3027e2327fe1254ff222afe54554d31da161d0ce8c1a750f7486f",
            "af198154d5a1ca3d4bfcb60cc12b973b5fd2bd87d860c4a7ae56e9c310c8a11b7f9edb728cdeb7ab7f0b6ef93c62a65d",
            "ac2cb301f30e85eaeba0ca34e7fca90e50846457207e0c18c489331e3d7d7ea67fe6f7e6c6762b79a68e422e6e507cb3",
            "398979c72880a7ee235b6147fe440d14a1882633d839342b53bbdbc0912ea4ba0e638e74125a2325c447e71c9881e72a",
            "27e56ad4598d7477a5a760f0cef61ce5312b4d274fa85570b6395c2f99a29aceb5549ed1cfcd704ee8ceefbd6aba6319",
            "2d7ec859560796cee9dbc0f27cbff004b883cc2ac210536168e532ca6e08eaa227ac6a7e1d16f3d9dec02900a52edcf8",
            "901cdfef486522de239ce0cb2fc4fbe1632a9757252a900b878bfd01d763d238f4e9dc30a57cf14df7f5fe1ee2fedcb4",
            "26edaf4e3808282de0df8db8bfc66f06de5f3bf8e789dabf63e2a554fd6dbef6b5e1eb29c35fb886fc9799c3d7c6ff3c",
            "8bf47720c2fd69fc6f26935efb793d7cf20bc608ff7e14e5b597d76304bfb08ef4dba182d70e5edf48fcc22cd6abfb89",
            "d7a6af1bbdbfd088fbfbb6ef6b0faf5ba9bf9082f7ef1babafbdbc6e92fd0214c1ef5b66af3dbc3e56fe0216e9ffea90",
            "f9dae1eb93da2f6090fdafce6daf1dbe3efdfc02eadbffcd59e8b5bfd7dafa17c0dffdabd27e6dfdcfb2f6efe87db1fe",
            "1791fbdafc9f15c5dff199ea6fe6ffd017afad5f57ee5fd0a7fd7face3729210907f18a2bcfcc5be30689dee8fa7ff02",
            "504b03041400000008006d996d5d9018dd50d61f0000f32a0000110000006f72646572312d6d6f6465352e786c7378ed",
            "7a65541cdbb66ee316683c587087e0ae8100c1217870090e8dbb85c69a04b7e0eeee1a9c10dc9de004770d7ec9def7be",
            "bd3727678cf77edd1fef7c3daabaaac7987374d5fad637e75a732acac0c06203e001880000800c701cebe1ed070500bc",
            "840700b00088d05aaf6d408ec620473d15375b63071d26576b2bd24c5868da740034e03ff8ff1a5d318316d0acd80197",
            "012d9d11d676994cd274614c564df4258a3700737b277002e5ba4491cd12ff3102dd9408d196892875c3a66703a5298b",
            "a63ab596c8ab7dab7677b7cfc6fa39ca73432817bc5edb02e5cfb2caa5f14c33f4095f71112c2e3ecb2504c6bfadf90a",
            "8e65bd7e33f4364e11692005418681fd393d8639c1c67794067487e53bd065aa8c29c5cc4c700d3c215aaa3e751cb5c9",
            "b0e5753315927e84f1da7ef75c917e8747560e44ac4f5dedce7e7b8fc8be50d6a37c82f810d1ef381aedbc6f09bea9d1",
            "c5426e04b93218f1584ec5bf7d3ad0a7a9fca73d2832ed39afc8b609fb4b11de86e27c18de6b80a801123a43741e6b99",
            "7fadc701fbbd39fb09c799267cc1f0677fe6f5ad072c8491af486af1ead6cf7bc24dad5d36349a32782389618ece9a53",
            "c97b39f00cbfe3b5d51ebec7b4b66712079c63c07c95d8057c9b432552c2e365dabeabfeeeef77c0ebe33e11bd2ca0f4",
            "05ae9e483f4dcfc9e95b34ba7ffe4a847e5ed4f2b2e4a8eb79f06e92354e3ed6a240c0dd110200f0f0000350fcc7dcad",
            "5165a1387bbc927d9c9a288f7357cfded8ca8199e9d7f93f73f63ff885e22839854e9667818792b7413d97930b11e8a2",
            "322e22924d642ade1fa4be036b42b32809175a6988d05508fc60def8347bdff6cd2f0884eef4905dac72147fc2c4e6a7",
            "a8752909fc7c6de784404f6d1f276157bcd285e13c793e7991f046da8dd6b9399cbea83e3f43a677894262fc58ac027d",
            "98499d00092b8728105730c157b35beedda2e1cf52a84951e5fd92c0f23aeef49de8e1249fe9b8dc6d546108afecde58",
            "024e8f83a1e96ca9678f15d9f91581aab219af79b961aa3d1fb99bde7babb653b1a456ac3bcf97f2c93d4ae41c983fd9",
            "86df7ca34a3ff7d1ace638ee4f71ebe246ea2cea39ea057f22c95f84fadd2461df70e78086050028910100e0e32fae56",
            "cc2e36f69686363696bf825bb19ab54d160ffabdf7d81d42970b33179ea4e2b8ea87f4a2890355ebdae668f212225b02",
            "3d377ea32895c9fb4bc2e888ecc19d985ddd985756e124c2eca7036d5f37e52209199808dfc652447b76e4b6f757f789",
            "5d71b0b5ca6371b0844a54b019b9a99e445e75109e9efe286bd52e66b812282b348501f499bc678a2d84e9b718a362d6",
            "382d28744362f8248d2bdbae28f44eee256f7f22d56661cf4b42a361fa78dac197a2a4c3edf9b8289127f5d2f3e0cf16",
            "28e46e43d82336719920f76a72418452fd938692b79a389caf06c4992664102c08c409b4056363df9f127d5f449f9e91",
            "fce97e4481c7e4ec83c99842a40a515193ab514c03cccd45091da7e97f8ae0269b49ee5e16a0b4906ac1e40bf2501bbb",
            "2165b009fff1a20df1de9c037a6ca71af3d5911064e3a06176e3c757d1f55b03061dbfc9f3e76a5bb23f7223ab24e4d8",
            "325cb9d8bcf0d42cef4d446b9d663e12a9bdd89b6b31d739f8e1c9101ac2ac91544487d8104492c055fb22a4860ad128",
            "99814926d7552516d9b1fe74d79473dae04d3f3236e8fb71ab16192f36ee168a6700830c0c970d284193a064e08d09f6",
            "400bd7f2bd6c4bb54ff06c7da951eecc5ea3f43ba1a493cc5455425f3587c0a9623ac5be740e8c9f2b187b11b3919eb3",
            "e625b70850552a6fbe82cbad0a6c1e5a8118aea47ebad975d204470ff42d81620b5538d9a85d91fb85d7c8276c3f2ff8",
            "baa538246dfdb509022068295004ae2030397f6e57969d21d621bf2b5933698946511c3f9dd940a414d9218d0cb7d578",
            "193a5135397ff2c3471792ecc343a6bb3997fd02903b105c96e46776363223f091ad688a020b83f08b9b1bbfb6cba770",
            "726519f471746596d9a1bbeeafe3987858d3d3371a855d1f1566894caa516a3778d9f474a6f16ffc292181defef96962",
            "b46781e8682eaf02b1a266b0c7c3d29cc4ca2d6b5f5537111aab932a53a7d5834b4b092bc238892860bc02bff4412e8e",
            "66ae5b2ca87a8e2fa348f63f2cd62d19b8cc1e0257934ad94ef85b9631892989d08413f8ad9c1b0785e76009e43370cf",
            "4bd6782867882d2c595a402a2e416ce8e8f8a2c733df5e7f3ceab2c378f1cc9a233d2c949c601e6b3aad2b7a64527474",
            "b860f9616b4a8713fe0e46b14a9bad20372a67f8680676faf5d487ec260dcdb6cb78ab2d34d13d8078e29676149984c1",
            "d6dbf5eee0b56bca306a903d69d1829e67b93a5a3188dbb4304028534b921bed8c7376f080e6f97c708ad8a28a8fedf3",
            "ddfa43e787a41541d7b453f6d0e82076d42c66eef12476eb329a719bf0de86eeb52416e2341be19626c207de94db5112",
            "833b96f4432c622cbeaea1f7214434998d202f589d4171b473f3f5315bd1721f49d716797592fe3ab912ddc8d1e3dd81",
            "52869d5bcb23fdcb52530eedfdb874927363c4e6a6b232c1b91fc280dfc9c5ae7c0722f9a390983c4a06010011ea512e",
            "fe0cab7f178dff09b1508f2116ea7f57e0fff7f125a65dba8b053df0107ca8b871905bd2f8d1aa11415ce3287b17c6df",
            "71ce3f57aa576ca1a53923c81945a6199bd93741eafec128aa17e4632177129a5a5d455429878ca4b92d6d6d8996f66d",
            "3ae98d242da216f21aa3fd56a4be9eb6defee228041d1839f682eaed574c041cd1cb2e3ca98092f4d0dd916ee9478111",
            "7de635c84746f896a27fdcb1ffb6011e9b7d705dd1fe28e80437ca68497cb3b54647d53c9859debbc6ac130e2530efc3",
            "22f269ae3596534019ab2dcff09895426df0cf36a668c16229ea80c86169337e83d0022a88a71c84649350096200c653",
            "ee0a946e818dbe631f583db79af627a312882369ffd6905ec05f75a6647aa01e80d41c6cc4d3a247b04a75d3c4b9ff9a",
            "0894fc3d280275d795c4ea821fbb15f03b8681fd9a3beb1f5913070300e001fe4f4072303336767460fee38bf517c992",
            "a3e5c2937ebdd7d86b12c6a5f6365121c51e8b2aba1d627a555597c330fc36e40bd2f959cac2eb254b48875a1f952612",
            "259058f8346cd3f8dc55e9851b8211d50eb8db400cc9cc3776424e44cfe7733b1873b7bed4358800d91d5e86645aafb5",
            "38cceca216ceb60babfaa5d9ae761abb8938cac779748f11ef3c89acfad30cf3f87dd4230b816cbe38b1d1103a26ce66",
            "747b0720ccae26f24c260fedc80134eb88dde6f9bcdac585b4f22a71f2aa89cf168c6fe12524652f9b3f5c513aa6478f",
            "2034f0906fe9924ebe8194662dadab1538cdb67699242d8dec876bf9dc9dadc9acc5daa7372ee453b0e3fc834e3c8f31",
            "fce20f9636a0ca09af941d57451470a693c6a00a894bfb732fb0927e2b8d43266dcd2c1aae3fcf0c37dbee1d7cd25767",
            "ac81b8de53c6d00d635c9f1b99e922e96b24a00e068c93489a61df54d1a8800d66d8452b27fe97fad7df8184cc2d121c",
            "d99549dfedcf56155ed3e75f6538b41efd7cf17e9ccfe6569c67f84b8f419f9bd32218675137153b5779f6a5cb77078f",
            "c5362088a0e0e5763a71f4464c3c0a2df2dd9c8130d10b08bf23da24a745f343fb302ff25de804e480e0997b9eb0c51d",
            "f077a37db83bdccef3988613c0fddbd166fb63b463742ce759b0030f116e808920779cfeca680ed63099318ee2908a8a",
            "b2a40a1471217931eb7ec9bee7460f5e6f86569d51031425db2a87dc162f7fc6a62cee154559b56be4ec22fd1091449a",
            "50da2cd495dbbab4f9b408c1dc154bf9a434cd9281507e46ed7143bd607eaa3f0d0823cda8c56a742e95c2abb2478d2d",
            "f705a577a7882649f36827e254854245d6c9a9f666b210554ae72712230537f8677c12958ab76c2848777e03f155e239",
            "b3992eaca0d1a4626e57ed10b04a4c9d50f583e4d0c7f2572b4ad3277c2fb3b0eeaaec6ab8fcc24c394b782bfe335c81",
            "6dff8e3d475a2aab8449fbc37be7334fa1fbf8d0848817db9ef79f58be0c3fb38a051e61e0340f619238166f61e0b662",
            "72ede9dfdb25bccf3a41ba133c1cceed8c38d51508b7f1b8f3c63eefae5b6b5a34e3223c824b1b2a7d730b1c1cd336e8",
            "f0a6374c8f1fa1533830ecabac161416c6dedb8bb65ced3a6188217d910e9fd745329c33056c3922b60968894d39679a",
            "5bbb9a8d2042aea05b78deec01b4a87146c4879b8639dbf75338e9242aeec20aca4e5f9a8d47a1a99cd87ad17ea438c9",
            "d5d18a2ce056c02a93c39c62f8329331ac61b023c67a34e9b68edba3f1dbc352d3daccdbf7eac25765c0119e9382c022",
            "1765467a01c4a5fcd2e48c82217c13b70af6e2fd64f1e164f112bd7b3072a06705f29433fd317b6cbd6a706a115127bf",
            "d0b144f172160b83aa9692e868aebd5c302ff57378d77dc7b5aa130ca61a8746be92a99fa8d003a4cf46c4326af85fde",
            "426db758cceafd565bf655277abe3eae1571c87fede6fcc1364733636be33fcf7fc8ca9ee682256503860f70dd07fd4e",
            "887f4170c9c434fc4a7ba9749f737716f59276d82cb746ad8e40cac80311c10015e11213515b5417f3242e08dd3ec28f",
            "882c42b8bb5cc6fe8a4249e952b64e2071fedcdbf6b3fe4d8ed669d1585991b4eaa4d2d2c33a9bc95632a294dbd839d8",
            "a3254892ed077b2263e362556ffdd845840a4f824137d6d8b99a987b9ac0cf6d858335419cc878d2088575e203f748a5",
            "6ddd1afa3c96901f82a951b0de490e18467075866f53903ca1dc8dd450bd5adf85b2ca09c42ce34d08250c6fc7a4ef16",
            "319796dddb799d7b1cb6bb8331fcde61529db7c5046395a46a28c1f3f3a3708896ae7f40720c7b37edd760726b28ca14",
            "1b841787ecbde5de72755a083c503fd665ea9a409b8dfa3e73fbc3bea16d8f91c3e1e3d76d13251fecdadd7689562896",
            "9a50c29af3078934a8e56dec548e35a5297c8e446ffe810a2a45e7ed910573384e42bd31bdf95522c7c4bb72d03677ba",
            "ac8fd70dfeb794d172831554da4eeeda2341e13d170bdcc50f6c9e0e287a24e79a2b3d1d848a5d66f6b5a1a0dab97d93",
            "b0c5571bc9fbfc2b2926e336a80eaebe60b319b1cfa9864358350c5386b7e04fac008e080371ac6c35ba84353a317b66",
            "55ed4cb258ba4e06600ecf0a66e01d0b22e9d98e77fd1537ee370fd8cfd99cccfae475811c249479e0dcc25d41e8125d",
            "a4939d799b323d6f87bb139b6975f569e60730da16b95202bd0c8884e653e577179fdb491fef4b94b3c3d6fb95a632ef",
            "fdee7eb6fbf3b456cfbb4b1fef34770d8b003addbb8ddde925cfbb9dbad6dbe9b4167bb343244fcd400533594a3ef2e7",
            "5b74e12b50a17456f7eb2adfe50d310d2ca496111df955c255de4385abc537d87a19261d357c605c8316d26e29af4f15",
            "332cc837daf8f6d60a0ac86beac5940cef5abc49ffceb55216f7a0f76b6588c7c7e306cb0328794684f038ba6cf51d25",
            "5e0844fc347e456c054e0a227acded2d0159db16ca9aa8dc254b4166bad2ce18a6dd31dae1e8f54fd5f45541dc1fc7bb",
            "adb4c78055d7aae5f50cbb4e6e4875971cb2a6644419205aa7f17f67e209f98cfc8d29168ba8aab21d58fe519b5241cd",
            "50b33b341b5a811971cd8c8342a9f17d2aa2eb62d4388991a4957c1d49a5aa154540a629530a42fc299b7dcf303c4698",
            "328d69a298f15edec93351ec2ca352eda3d7a032742fa0775bafbc8b705e97fd6512563bbba4adb1b4f8872dbdb88e95",
            "fe20fff95dbbaf1bf1f595b20a5d6acefb483b2161f7f60601175064cff496f1c9b666f3d0e0bd5235b99825853f9216",
            "ed61a9eb0b763817efdcf84b8265f5aa52c881f9d5a217fe8c3fd184828a08834d03a2f13ee92486b7edaa6da5226b84",
            "846d3eae384d8d97310270d157bf9ac7d0603ac6803daa6576edfc4be6720e73f875d53b4e96e68b59353e33839a9844",
            "02bf6bc263de68cd9b801338453058b064477d1916c0901ff7a262d3aa14b58e7d83fbe25af5500414454bc1e627b91b",
            "9b8828edccdef18b9b594e82a0af87e3877b7d0c37dc8dc93af50fed8b2690cfbdacbef69529f5079a494d53dd37ad36",
            "d6ce253f9a02374a331a2ccdb4cbf4497232e8dcf1cc6c3e9610571d631d7ebf4c50ee969ff81809635eeb4df9656aa0",
            "9bc5ec227ba07bdd8484257459491828abbf820b4d884f597f8daf3a2ae5af3c43823cc0b421ba51f81c3347bba33612",
            "9fcba3703d105475551d02bc629c3fe0b891f136110d56653de0d65fa3230c2aa18fcd482af1543cd552e6b67a6e52e4",
            "24f7aa7f3150cc9c2d496c5024b386d81306a9d9e1c240c39188714ef36e4e552b17604158d6e6c10dae3b2f9dce370d",
            "d17f83e67e9670a1bcb153e352e320949fe7821f1d4b01a8c5ad135c18d576114af56e649c58ec18c7455a8b612cb6c2",
            "ec56cf6d8fc6f882ba39fa6ed1e5a1e5435349197f80cc4bf739ecdd49c511b31ae1c03227cc8648420da5216c4691e7",
            "e5ddaf351c891357ab77070d96f098d6b0c16f765c9807706d01ecaae45ac787519a4a6493e8fdb455f5bd7c093e8685",
            "905721dd3a0ac20fd4441a86519212a5303297dfa49926500ebe49a07d0179240c6cbe2fecf3660e359331ad8b58ad7a",
            "bf723b0969e5d79039a1aabfa14cd44957aec55660f5d7c8c2313b1ccaf75ff32375380fe9709ba18ec63976c962fb31",
            "bbc4b74e7fc26a62ee7f1aa534b481ff1d3b06f585435d85c836938f09cb562d92c2f3c68e1697592f8792a938861da3",
            "3cda4435cd51e3f1143961dd58e60afac2fcaff4fd0bdc0e20d3943d52c514e3d53b4b9f604dd6047139f6c5a14e3ef3",
            "cd2cae1ae00c464ca18243d00297788779b0cbcb3b5217760ba7000dee38eaa3c115b72baba4b4960febacfeca5d54fd",
            "f24defa05d726b64c3038ef7e9d24a4e64a8abc4a2afb4b973ded1fb486f4086989a74b776884e3bd7be10ea23a5cb0d",
            "a33058516ced5d5889e8ecd98f5ab583fcabb82e245a2d506c28a7be15131a54d56c6596304dd3ec34aa73b35b06192b",
            "588f9557606b5b6b61843dc0054ef1f20f0cda33af27c8c4baccd906c4819d9bd36e4275bdd0b9e0715a17f054776a1a",
            "f5ccb9dd6ba4cb46f3e1b2249681372b9baf14c4ed467a5abcf13acdd1b57e2cbd8a3f2e39245d3df79d6ea60b2ffaee",
            "b9a9c8b9c8bd24029bb1cb0f09beef698779035d628b3dd1a9cb93ff2c50a44a09183428dda4baab0a6fe08077916ec6",
            "fdad7958d011d426e2c343144e6890c30bd05e38e4a46ff6b65634af7d73674b14b2f39aa4cdedaac1bf0e2fc5c141ab",
            "2d67e57cf9f72b8a887ecdd886c71c531c1e0040fd33ea3b38ba59193bfc8af7d92a8316f3afd05b9e5fa3f52fc6c662",
            "f7d784662b7ec38055a2a50cd09d1eebd3ad1436a991ff11a71be5427b86eabe9b333232cdbacecbc770f5863ae765af",
            "383f3ffff546d29509b75041ce2bee0115ed225884374212b23f1a972de2965dd413b8dbf22b557054ccbe627e31c411",
            "92d26b81303461eed90fc58b03320a61874b7442100b4e78a3abf9c6507297b327249450b4d829263b5834550e68e824",
            "b76ebcc8e9d4ebc5f17a6f7d5e51724ef7077cb66c488f536858769ed437a03c1dbd92357f9d079d3f2512be918f764d",
            "2d56334b6dfe25fc44329550677b8c188a69d4b8d7e122fe20e219d2e0a5e4875b8e6c52726f57450feb54c80445e082",
            "2d0d89fc9a1283568636732c1cab6fbe65fb96e63690dee1cdd0ac448edd39c68222542a209a63a86604fe5d041adfb4",
            "2b766c71a86b7dabded2bdc6d2bdf6d242894d6bd351943493dabedd165157ce22234eb20399b398e8672c22f3671fdf",
            "99a8c013565794a9a1f4e4ebd76bc6531f908cb2bf5617c173f22695ebe048b5506e6339d42a77caa9d0cf81cb0c1cde",
            "00d5b8432d15d4e2c22aac81c713d454ea95342352a07eaccb4264b88f0efb61f61ac52ce8d8cd873fc1764edac9cd24",
            "2af6745b1dcf85761e66d6cc0c078a3bb9418090177308d92edb91d13cedc0a516a1db6be8ce4adf57e354cfe38a84d1",
            "e1073217ebda5c90355ce61337f62a5fd120ab8de594ab3775c21eee73172fcbe01f642e9e105a13421eef51f254465c",
            "7835ac99e9f1983546c78e607df8085b09e4872a47087faef94ef39b447287746331efa484d1445b6d05dd6f6f0ea5f8",
            "7487daa4e67304367114c27e413151decbedea6f666e4f9193694d789bbef043e8d9cdfd2eed51f3de0cf08b6e1ada17",
            "91079e934a5a449abbc59f8c7d9968975bc1cc7e0ae0742f23a22bb3f70defc145256a5a018dbae848b53e7460ab58d2",
            "9b4be13bd4df719cbbeb52b8f4f1eaf8f1c0fe6f8e9b19d81bbf5776b4370799fe4175ed2139e40e1674b8e5e6e5171d",
            "7b9caa626251d818da0d13305dced5be264afcca3e201b3131dbab163798e198f09942c20b22cc77f172288a92aa8530",
            "3a8c44c0a65e64ec4f797375c1231d58f8d2dcc4f8266faf3ecf88528c95d46c14b66fb925f2657bf2e609b9f756b49a",
            "1db89f71724330b504993fe68409a1e7d3a352e2e2a6c3234f43d52ded3d4330eddabf43d2cafa8a7f170804ad365d94",
            "b92d09f70ade496ef61ec0fdeef96e9f739983a1000026180080fccfe733767534b6071958c99a832c1dfe71f7e7fe40",
            "d4a07c27cbb36f87925f5ec7f0cba893c14e6ac2264650a3137d40e851b2f924b94adf2be1e691bce7176a175f5f1df6",
            "e2c1f0617fd4da0655cb1748a03318e40766b745d6dadfe0f4f1f1a8afdf6014dbd6341f11196c253d9adfb8b582f582",
            "3dcc761b3bf0c3e80a486436dca327036b7e8835f8f082e7e3c08816fef6d626b923f17b425cc028c15bdcafb1e54afc",
            "2dd42a6a29d015633e8c4d2801b5d16f156a468f19aa78aa0a13075eb8f372aa58097f79ef65d3b712a710968550ac46",
            "0c33b4fae103b4297bc972fd10a19d83a279f77db82a1a5cb0a9649d22ae8a56d5f809b15d96f10f050c439d3dbcc931",
            "f0274b5e413c4c381d8d2d26db0d8685a9ef0a7367f24c8d51d40a4c1d4ddc069bc9f01cfd7b37bd740d6612abc3c502",
            "1da6e6be88fc4925d905570c4ccdd017aba4390fc842a2cfebc8ec60f1fb69f5c8409563c13ff8b176326646d1ee0a58",
            "c3d342b06572a18da2bc3ea31505aa5951aa4d889f8c8783c6b85f63992c0d01f346d79be5abd2b7ee13b9977eabc295",
            "e3a8acb18f57058f07fa9f2368646065f4daccc01cf46bbc04c2447eb153fc12ec82585537ee4c462ed5882ee84b4c9a",
            "9d8544191d859459c7cdf0391c7828dcec69af6453888942181f88815dc16cb9fed9f7db9d338f6fd486ab53010bfecf",
            "a836c34bfe7d9d93cfcbe445207d5674cdf921c82a8c31cdeb59c8cd092367d354458b62ad3428f6d5847d871b51787c",
            "13a0f8a290a05e3e3dd3175f6587ea83dada43d7fd6f1f61bb6f8137f49184a38fc104038008f5dec648d1dec6d681d9",
            "c8c6def8bf5b01feb3e709f8a3ac26d3c942d47d085eca585fcc5950ac76c4517156424fa65d143d71c19dafa61c6d98",
            "9f5144873da7e787b81ec1619fada37827f0dc11e42e7cc2120caea31c8e26197e90ba4d58e0a9747ed5c3b262bc8e4b",
            "f41e2a0f09c5ad2444ad10c703f99d12c690648c2d6d16bc24075648f815c7b919b9be14251d220eb8fa59b04c78ae71",
            "7581ca8e01e3367597186e881fa11d0506eb5b13060b1e2ac74f1380e72f4c31bfa523527c5c7e6696f79a7c8702ff05",
            "f9b835c60f8004ab183a8c3a6613dc731f371c88c6bd3b99ba44b25d3e41414201894bcca0990251a38b310fc7d4eeb9",
            "6873cc37962b9d2fb5f5e5c9e1164102bb1fa1913d2504e93fe48bc136badd7a3b2f78f1ba7bf0ded5bb4e96c82c088d",
            "97bc5968ec9f2a7158f09eda675838d8da089e6c16746df5d93a38dc36395dcb3dbd70dc5ac3d050f314cbd6e4b0a9c6",
            "804eb2be8579a0195f1fdedd41851a6083736d6e84efb33007b5688a888cd2c9cf955a36e514eeba0ef973b336a78adf",
            "3b0b3364bba05ffa3d24f0006085ee8f91fc1400bfe3f0dc1b3e81bc4782f2c0fc9a867fe3b081aded7f28fc1792a32b",
            "40f32ce88197d05e08be1eb25f5f823bd2f1b0e53f3d7b69d826ed1cc7372b4617453a0c5e60d5bde6c26f3ab3be2517",
            "948290e7bdf6053d2cf85c9f6f4b06ef7a1955e11ab22de2ecd9c7553d4bbe88b54b022f7c4e63162a67ccb08fad5c14",
            "69478b98a9f1bc9128ce04a28fbeeaa3edc8015b6374824fb3eb46f23928e623956f6668b3d940f56596592bddbb3a30",
            "685702c3653c7b18da7093e27830b311f566c73eaf19caf85bce1489a1a1648e59612dab45afa5af26b7b6a2cb586faa",
            "a7f5e1d212c15b53e92ea3cb221b6005998aa9be8e28b071891574bb1ec12aae4a5d752c1babd4e86a265ca89049ef6d",
            "7acde0bb0b39aa2d5be4a942b398d71d3bdd8199d35c9421278e9e7c2f7cd288d3e4c35bf9ec6ee97850d14c98f38665",
            "6bdc9837ef00c1a7ed4cfa00f235a21a0e33b28549069c63bd75c1ce36c1e5187a1020cf93627bee3033706f68345577",
            "e532892c42bd4cf79a9ad685d448452fe426f5e0f5b60d76c7e9d4709261b32a063d5ffeb14a77607d4ec09bf900c6ba",
            "3eeb74f7d4abf9ad839ece238a9efea277f74c796b6967b03e0d2727bf4ad5ff4a6e543e6b63ce47eaa63d0af44bc06f",
            "b2843f6b55ff922bfc51b0fa12256709cd8a1e7818701f4475196f35b5435eaa6481020e503aabfdccca7ff3d6253a42",
            "6ce1ee123547142e833bbc9365f8a67de0b3d7b4044d75e9283dc531241a3176423cab3064afd686832af2341dd9aca6",
            "23e4d9568ffdf14f96872e5c4f1e7b27236d6cda3ee3ec8bc5013c87ca6eb4ef92842a06ca16f3bec0ac927828b70fc8",
            "09cb9e1d7b7af0efb459b39c22319a9860836bfd90fa1eaa9cde7161aa6b12f3abdc07c2762e66a98e80089c7e02f410",
            "b86690bd24fa72bf2e27370f9ca4a038ad3bf378735e55b1321c45d854fb732d199ada2f709f5523b8c34fbf4522a3c7",
            "d196354ced6a106260f4ae554af2d55d5fa5cb8396a15dbba37c1975ac211b1df38903ca17e2f91193d7e9da8c6a2cef",
            "ed52e3bde19ece73b2ebdd32eb0e900bf6525e167fb691ba4dcd4ff8ff79ff50d02f01ffae5fee297ed33df7d4c13f9b",
            "76fe8e2f8f8af5b7169ea7864f1b19fec20afcbfb4353c35fe6759f3ef4047f9c3f8df14399ffa795abcfa0bf6a8ffbe",
            "94f5d4cbd3a2c85fd007fefb12c9532f4f37bbffc210c66fb7be9f3a78ba6efe0b5df84f56d14f4d9f2e47fe423ed1ef",
            "17274f3d3c4df8ff02e4c5ff4dfaffd4dfd3f4f32f4893fc6b32fad4fa9f99dfdf814a0a00fc4b1ef8d4fc9f41f7ef18",
            "25ff9bf97f87e0a7d64f55ed2f9452fd3f6a9ca20c1cfc2f43b4c74fd9e39f61a2f975f75f504b03041400000008006d",
            "996d5d52296b20d31d000090280000110000006f72646572312d6d6f6465342e786c7378ed7a65541c59b776e31668dc",
            "091a3c58208d040d101c8234eeee1a5c1b77820597e0eeee12245870126824b83b8460979979ef3733bc79d7faeeaffb",
            "e3ced3abaaabaad73eababead97b3ffbeca3200d078f0b4004200300004a003ddc8a95370c0090870000e0009061355f",
            "dbda3819db38e92abbd9193b6ab3b85a5b517c8487a5cf02c002fec1ff6994c68fc9c1b263f61dd2b6f7e4abd6681ee9",
            "45e8aff8ba73485109085923d567cfbf4c871c5d8ce2d3f4965a3df1f89aefb371ed9d9e3e6a565d265befe7fbed6b2c",
            "7a38cc50a75441ad76e29ee15dda52cd939c4a2942d36c233ea19744baba911b7cc0a4b7b59f6255d87f228dbe052ba0",
            "31a72149f3a11158609913272ea035614b2f37f877917d1815af648f8e4424897eaa4f39436b22afa9b11807cf3650bd",
            "7da7b7aac2b1d236ae284ab54e52e2ddb486140ca19a242aaded27e9d66b9dcefe7a35bf8c79295e84a221fc1287b23d",
            "2c85edc7cc112b013873b42cec94ea79bc4dd517ecf8b8dd140e852c20052e1850f5528c385cc3ecbac7ab4f904c3c23",
            "487f8c62e436d60e09e90b100d9c4460ad3a206fea85936ee12224c628b2bcedc12316bbcdd6dbf8c2d76b56047bdfa1",
            "5c1470fe148e7f761146ae8e8f3b97cd41e32df4222d20e780c7c75da666f995a226ca918f6f66fbb8e866460085287a",
            "ffe8f8056c74d4542c4bbbbdd257b09fc5f146d16054c78373dedfc30114fee6b6752a6cd4670f47320f5e89f6e0b6ba",
            "0ec6568eac2cbfedff71d77ff01b4a6365e57bd89e041d4adc040f5c4e43633045a45d84255a2895bd7d251780759139",
            "3424d00e3a524c65627fb8373e6dde379fe7a1af227706282fbe73968663e3f251d7bb94057df869ff0e8991d62151dc",
            "be74a517cb79fa7cfa22f98d941bbd735b3463496361b6f4e012b5f8e4b16815e6188b2a310a4e1e69103e7fb29f469f",
            "acfaa2c18f72986911a5fdb2a0ca0650d64edc58aacf6c62fe36ba60088fccde4432de80a381e9d772cf012bcaf32b62",
            "1525331ef34a830c075e2a375d23abce53d1d40e9c5bcfe77269038a549cd83f38c6def43fcb3af7d1a8e53c1e4a77eb",
            "05a1f4940c1c0d42c2c90b177fe924a23ba304bb7000c0160a00007cb8e26ac5ea62eb6069606b6bf95b5e2b055bdbe6",
            "7063de794fdc22f5ba10903c819f724a38fa583275a0625d0f8da32ad3b2232ef7e5314c519ebebb24898ba91db148d8",
            "d5894f9a7be1fd8a752c82e7b6b65c6a3574a81ef37c8e866e8ed3353db4b590bc6fe3837a9878fe382378b314978b3c",
            "75ac954b30fdf66e7e5b6aa0df83f2293aa070beaa49b918e579c94862abc17ea8992b7a6c600c9e125b02103c2f317a",
            "e23cd8542c35d43757c3089eda40a48cefd70bf9f471c095b9bafc7bdd375f091d0e848d86a18fe648b58c7c88e50627",
            "55960ab4c5881407e1d50a63fc85e6caeaae770a688307d2761eb25a86cf5776bba2f508f7566830c94855f09271cbd4",
            "0a717ceb5f8a93ff90af141b48c52c051ada9162d7d1f8303a61b880250ffdc16e7cccb640f123d61cac2cf0580226e4",
            "5298d3c534d53a239d252c76c9a6d0b232ff6a9831bd6882536aa698861577ac3f1fa854f5dd16df6856a6e5237c91bf",
            "f63ca8ac862cbd22f7fd29c2ae14171c15f0c870b2afdebf6a06de4046ea93f1086eb468996fb06687778223bf732d6e",
            "a7dd6acafd245eec6b7d988e4e554275fceaeb8bb0a2bcf141072bf1951b69b483b175c74b17f64c4ed961dc54dd9088",
            "8453a28a94263582aded8c6c3c4ba678767a59786f57783ea3bbb0f6bd4af01d9c9f7a7e8e81b0dacb7cdbbb0e20962b",
            "458f5faeaa14f1d1fdab765f5168259d04490cf512f89c67cdf86cff5326af7d9a384618553495542b2e4228193564df",
            "68adef8b00f60d596b8c4e12ca4cd17a222ab1a395e2c0b5df566c19124a97beeadaf5a9eed185ffd929946da0dd52fe",
            "dd64cbf31fc5d03871f9c873edb79b065c2fb1aaba52371b2655585b991ceae0391092209f4df6813aaf98bb0cbeff74",
            "27c70daeb7b26a87175e7b2629f59e40f695a876ff1bbf37dab8bd6fe37f62c8d23ddd7f5ead933cd5a0ca56e643fc4e",
            "171bf4314ea848d012a34dab41974cdf8098636a73e034ed8c88b4fee4ddd8c1490a6b63c0902fdc42d737ef1aa8eddd",
            "38edb23ee2792ef96d284c2ecb976f599ee3c57728ca2c9daa22c7bdfd3062dcf65aa26ffc22c5a7aade8a1af2d9dfd2",
            "c8e597cbd5b75b76dfe7ac5ee0e1617c48aa7718cc4d914ce4ef9e462a975747ab5d1045dcf3b2de9b2293ad82a39d88",
            "cfed1c06e46ce372eff415adfea489c20641494aa0ba9e954b91a5de20d3627f9dac792141e079928daa4b78f848f7cb",
            "6d5b46eed6a8770b0d2d3ef9ae708b1f6e24d6cf2e243e6d9daeef9d4840c38d2ef659f8c16e0e3c9174fdbbf2e51574",
            "3e23b1153d583340388b3669571933df742308291d04bc5620409c372786716ebe36612e52d921efdaf4394320d13655",
            "038e3a67b682667e34ffb4a2ab67ff3a65dca10e4f589ee833e6cf43055dbeda1a9f5fc685fa42422982875fbe3ec406",
            "620032cc435cf8237ffe353afc772e8579c8a530ffcba1fc7f1da571c2d2bd6c988187907bca6168bef63bacaf4ec0b0",
            "ca65c95972cc09739214da6738efae0fd1f1e65998cf27563d5e2f7ab2b9b5dc1671eda1f19b9b221849a23e297308ac",
            "4ff08d5e1d1a8163439dc09fe00e8775a799db9c556d017172f4be51de8d5293a9a2c9c2c8f6584b538e2ccbed7d67b6",
            "9aebd49e4029a57b2c572c0e1e009937af7b67c05ca9d3f2e1c2de64372238efb5345edd5ba9266aacd428653eff66a7",
            "8f007c8f669875868708e37fc355430381ed18ae548f839187694bd850993711496ee9ddb478cfac19b7377f1564d1cb",
            "a61f897fac5f8d5ca38d232e99936151b40fff4243baa159170762b15de50a1498f1a38d79e181b261a392f42575aefd",
            "f87222eb7ee6b446644310f02b5abd0fb3c1023dc82ee287528a10f0ffd28da399b1b19323ebef5fecbf312b2d5e337c",
            "fee1612ea1dd028b752e18128d4d41a1f6ad3434d21f1435dc3d18bc14d6eca90c34c5156de2ef2ea3d2fa5a64d590d1",
            "50d466865d467f0cc81d5eaa48d6eb954e9c279d50e4c25b1439aab0f69f5d37042fc660ef8aa6872bceb2e523559e09",
            "785cd342cd6bea99006f6214e6d9b579666922d49b20527342208555b96cd967a42c23b15ffb7c07e6f38b8dde122e56",
            "4915a63c45096d0aca0e17914c326f2aca727e13e2a7c87d66ab5a5c45a7f1aca24ba5bb14242bbf532cdc2b92f0994c",
            "83fe5982ccd30f752f0dde7e5ff0582488dd43bfa3bae9a74bbeb80b54cccb1d0713b00871f0b9b793faf48acb0c4439",
            "f1dcab102e9a22d67f46fa1110a6631abae45cba8582df81ed3cdb75b3ca3128be6ef8d38337f43b7d3bd9c5c6e4d7d6",
            "d9cc6137db5ea4b7ac18483977561cc4724839829cd28e8128d745217e063274af92fbe38dcd3dbcbde9e7be9acef686",
            "7e52a517e61542970b15cc402b9b6d3fc2b20d6c4fd03e13a8ed3fada1e4c135506a6675bf7c3ae5b08ec9f104b1616b",
            "1e29732d98482d940e4352b8a3861d47cea8689c0fb2925d6c1de083e77aa2689828a5cbdba521aa4eea981ac830ff11",
            "74656bfbd333fabe03b439f7d60824785501fcc27d521494ea82c3ccf80a79a9b03c2dbb6894c8c4adee45e97e9a98e9",
            "e8c719561f6184d7ed1a08967c093f649ed9a87c744c3074eda75b0c79e6896b5264212d45ff4d19313ec9d041df7756",
            "0e24c044be22c25750326911b844d1f9a41b208babd132d94d1e735092788bf02bb2edab4c0d7c420200f0a87eabdb7f",
            "279b9399b1b5f11ffbdf79b6a701b5a469c2f201aef960de0af041f9974c4ca3afb496caf7b976bfa25fd28f99e5d781",
            "1b88250d3d9091f4d1912eb191b54474b04f1283311d62fc49296304fb2aa51daea815152f651a5ea5cc9f7bdb7dd0bb",
            "ced33c2d99a828915299565cba5fe330d94a4396749b388778b4074b70acbf48616e5eac196c9cb88851e64ed6efc399",
            "38078bba67befab12d7fb0ca8ff73e8922467eede981fb7bc56d9d3ac602b68875fe8c5878ef54472c43840683b7e928",
            "9e30ee866074af0ef54876d957f1cb845302c963dbf159bb25ace51577f65ee71e875dee102c7f75ec67e79df1a13865",
            "196a8a887c7c689c22e56bbe284e51eab3fe4d263706222c09c18489a8de5beeed57a7c5c003d5631d96de298cafb10b",
            "7337eb0e4d9d7bcc9c8e619fb64d147d70eb77bbc43b60d8ea2249eaceefc5336196b7713338571567883853bcf986ab",
            "9e29386f7f819a237091e84ee8ce7f27754ab9adb4d90665c9f8785d13f5a78f57eaafa0d3f780ea8ff805f75c2cf017",
            "7d393c1dd174c9cf355606ba49147acd1cea236deabfed9b442d0a6da4edf3ada49b4cdaa23bbafa41cce6443f64188c",
            "e2d431cd18dc40c2d9019c31fa6238b96086e4550651075615ad8f94090c3d4cc03cee15eca05b36648ab31defc62b10",
            "7ebf07fc875c2e563daa86204e729a02487ef12e3f6c990ecac9cebc6d85aeb7e3ed89edacaaea2ceb3d04638b4a3199",
            "51da869c2ebc7ac1c5e766dac7fb12edecb0e36ea5a5c27bbf6f88e3ee3cb3c3f3f6d2c73bd35dcd229041e776637776",
            "c9f376a7a1e36636b3ddc1ec10c5532348de4c8686978a608b217a052692c1ea6e4d7941ce005bdf427219d9894f395a",
            "d908261a9cd464e765907ad4e4cbbc0a2ba0d55ed998216a505468b8d1ffd60a06c863eac59286e85abac9a8ee5a2d83",
            "7f30f8a93ac223ecb8c9f200468e19293a912157754791272444ec3469457405413244e427c85b3c64755b2067aa7a97",
            "321d95e54a2b7b8c7ec7708773302043c34f19797f92f0a6da010b5e55b39ec733ea675a5386bbc4a8350d33da30e91a",
            "5d80ba8967c807d47e96041cd29aea2e606598168d3cd840a32f3217569e1579d58c935ab1d92803d975317692dc50c2",
            "4aae81bc5ac58a3af0a3294b3a52d22987c3c018225694129d698aa8f15ec1c91311dc1cc372ada3d73615985e40efce",
            "413917c1825e87cb549cae171276c65262be5bba89dd2b43c101f3bbf69f36921aab65e47bc1cefb283b1151770efa81",
            "1730944f74978928b7be1660207a6568bc6495100ca328d9c351d5e3ef762eddb90e9080c8e8d6a453010b6b452e0298",
            "7f6008049790849a06c611866ba74477ee82b73250d522a2363dbbf34d8d97b102f131bf7f328fa7c3768a8778d44aef",
            "da07947dcb3bcce3d351ed3e599a2f6557fbc06ad3c2221cb4a081887dad396f0249e612c662c39119f763824242d6ef",
            "44446755a8eb9d3e8fec8b6936c21053972c859a9fe46f6c22a375b17a272d6ee6bce3b7f9743879b8f799e91ad49ca6",
            "dd78dfb56812f26190ddcfa13abdf14023b565a6efbac3d6dab96cbd2568a33cbbc9d24cab428f3c2f9bc19dd0cc36ac",
            "ec69cd31cee1c265b2529fdc54d87b38f37a6f9ad699e13e36b38bdce1be351372b6c8654541a08cde0a3e2c09114de3",
            "4f229571c900a53972d461960d918d6202ec3cadeefaf7442f3d8ad7826c6aae6a238057ccf3079cd7d2de2622a12aec",
            "0720bd550692e032c684ecd4324f85534d2590158149c93b59a1a1c52051738e54d111e18f754f3de150da1c2ff4d59c",
            "4899bf69dc7e53d1cc07589054747a80200de7e5b385a6117a6f30dccf922f943676ea5cea1c050a0b5c88e212a801f5",
            "f80dfcd0712d17810cef66e6a9c5ee497c94d578e6522bec3ed5fcae38ac56f4cd71f54597fb76df96b20abe40e9e7ee",
            "df7077a715be98d5090655bcc36e7a4fa2a6388acb2c4c50d9f75acde969caf7dadd11fd254296555cc89b1d17d6617c",
            "3bc00b152acde3c3580d45ca69cc21fa9ac641de641f83e210a1883e6d79c17b5a5235835809f17238e9cb7e299629b4",
            "837e718c561b8fe4e14da3e2cfdeac9166d2a60d31df6b8c566ea6433af8d4a44f9e355ed3a4686729d5e3cab307a8e5",
            "e0991d8e1606acfa53389e4774bbcdd1c6e11dbbe470ac7f5de25d633c6137310f388d551cdd205ac08d4727736ca812",
            "de66f13161dbaa47912768ee6e77f9eae5583693c8b46358409f02d618379e4c9715d44960ad622c2efcc438040539da",
            "98a6ef5128a41b7fbfb5f409d5604f16937db138dac36bbe99f3b20e3887155f2cef180c7d29d66d1eeaf2fc96c2e585",
            "c5bb40355022edd1c88adb95556a66bbef1a7b8052efb321b916755897fc3a99e8c0e37d86ccb21369da1ad1b82b2d50",
            "9e3aa38fd446c8284b8bced60ee969cf6a2b891e4a96ec181a9315f5d6de8595b0f69ec3b855974d40cdcb0bf10e0b34",
            "5b9a99fe5212fd9abaad8f652cb3743bcdaaa01796c1c6f2d6139555b85ad69a5851f70841333c7cc3230eac6bc9d209",
            "2edfec021321ce6d99d7913a5e982f11f13aa0842a3b75cdbae620f73aa98af142841cf165e0f5caa690bc98fd978176",
            "6fc21e734ccdf525a1a4e3b2438aefe77eb36d0cd1250b9e9b0a5c8ba02561f8ec5dbe90d0bb812eb837b06576b8533d",
            "3adc854f82846b1481c123522d2abb2a88fa8e84175966a0feb6317e279b4e611f6ed26812fd3c1e8016f4908bb1cddb",
            "5ac1bcfecdad1d69c4ce6bf24eb7aba68006c2744747cdcebc95f3e58e5f4acc98218d84a6078929860800a0ff91f51d",
            "9ddcac8c1d7fcbf7b9ca2316f34298ed043f31861613127087ea227315fab1e015e969027566273eeb540b9ad4c9ad27",
            "eac4bad09fa1bbefe67df932cbbec6c3cb74f58636eff9a0181f1fdfcf8dd42b139040519e10685859ab041ee98d80b8",
            "cc7af3b245e2b28b6a32a8b3b05a194fd9ec1376ab019e80a46e7b08530bf69ec368921820bb187eac4c3b02b9e88427",
            "ae9677022d7f39774a5c114df305f574379b86f2011d83c4d6b51715836aa318e1e08d8f100dd7ec50e007cba6ac44f9",
            "a665e7693d7d9ad3f12b19f3d705b08533c2d11b85183f6945ebbed29ab7469f486490686f4f3c856119371e74bc483a",
            "887982327229e17bc3994b41e5edaae0619d1132451d04b5a323975b5564d2ccd6624d4060f72bb4ecdad2d806323abe",
            "19fd2a9e677f8e055580c900c4718ed67d41548fc1e09d75c54d288d746decd05dba535bbad35a8296d976b41cc54ab1",
            "80f7edb7487bf31699f1d21c299d45453ee0909a3f095337514624a9adaa00a30d14ea356a24d11e908fbf78ad2a4cf8",
            "ce9b42b69b33c342a993ed50b3f25d5e955e1ec2c7a0b10d9b3a7798a5a27a7c78f955c8643258b9515123261d667d4d",
            "26441a74743804b7d72c6ac1f0c27c2c1cbe67da5e762e4561a0cfeaf85b64cfe1c7bab9b120b1776e21c010b26f48b9",
            "2edbefe3b8bb804bed02373f617baafd84269f1124960862220e7f5c6ce874415573994fd9d8ab16a243054fe455aab6",
            "f4c01fee834a97a5890e3e2e9e905893843c9ca315287f71e151b366652464551b9f3882f7e125e920961badfe42f263",
            "d56f96cfe43d28a20f8775273d8a2ece6a2bf86e7b7334dda72fd236a39033a885b318be15cd44692fbf77a88db52b5d",
            "56ba23f96d16745de0c9f5dd2efd51dbde1cb0552713a355f89efba49a1e99ee76f107f3e78f18975ba1acfef2902c2f",
            "43d22b33a326234849195833b0590713a5de8701629540717d29788bfe2b8e837a2f05cb1f8e8e1f36dc7f71dc4cdfc1",
            "d848c9c9c1dcc6f477aa6b8dcaa276b361222cb72d9375ef71a9888ac6e26269354dc1f53ad7fa9928f229f9d8d88a8a",
            "da5db5bbc18dc547cf15935c9062ab27c9a22948a814c3693393025b065171c30bbe35847ee9c62192023d2532797bf5",
            "614e847aa2ac6ea3b86bcb2d8537d793a740c07db0aac3ecc0fd8c0b1482adc9cf1a9617258059c8884e838f9f85883a",
            "0bd3b0b4f704c9b477ff164533e713d16d10d0e67bcb4585db92e020ffadc4e6e0c12f95fb0dc14b73080c00c0020700",
            "50fd717fc6ae4ec60e36fa5632e636968e7f3bfba3608c1d91eb617bd27f28d1fa3a9e4f5a95127e5a033e25861693d4",
            "176940d1365ce23be3a0b89b47da9e7fa47d52636d14d9bdc1fdfeb8b52dbaa61f90587b24d81ff2c20e55737f83cbc7",
            "c7a3b1718359745bc3fc8bf04807c5d1fcc68d15bc17fc61aedbc4813f566f600aabc11e232544c33741df978c3b6cf8",
            "8b26d1f6d62695d35323127cc038f15bfc4f09958a7cedb4cae074d8aa091fe616b4c0fab8b7f275e3c74c35dc35c529",
            "c364ee3c5cca5682ad465eb69f5712e5a372904ac14fe146bffbfac29abe285b6e1c25b1775430efbb8b56c140083595",
            "6850c057d6ac993c796a9f63bc2e8f65a0bd47383d0109b7e4e127c446d056db62b1db6082ce2cc87f3b9363698ea595",
            "67e96e01e96fa621720eed5d0f323499897f1f2b7dd56d6aee87cc975a965b74c5c4d2067bf19d22ef1e554084a081d2",
            "1e9e68885e97d2a67a22749d0f67277b6e1ce3b6883d3a3302573a1fd630d6eb03464910d88a063c257632196d33017a",
            "8d63b2340a2c185f6b93abc9daba4b012dfd320a9f78bf434e78382a7ad830ff788386fa5686afcdf4cd6d7e7b5fafa2",
            "845161d931e10f213ec88e2ea6b3fe21345ffdb7d0ce6028f9528438d43984cf2efbc23f047dcc6e3ebf19956c2b9462",
            "b6d2231657e86d677e9f8978d0b07b802c95b2fa29afdb68f9356c0dd4a6d66530d33fbcb4d9ef939099353f6518303e",
            "fd89f1b383919ca93d8dfe9fa2ea4c4d62b06f87502e575d3b1710f4ea54a73b8d61f1ae9114b63f9871ddbdd923f8d5",
            "2d6c7f86f2443e9070fc2199600190618c6c0d151c6ced1c590d6d1d8cffd5f4fd67e60bf07b1745ba878db4ef10b294",
            "bdb6980755a875c2537656c44ca35f143971c19fafa5196f9a9f53c0843f67e40b713d42c03d5b43f34ee6be25ce8786",
            "e3f08736d08cc5918fdd4bde2443b9ab9d8506d8568cd7f0498d600a50d0dcca22c0c5781ea8ea8a58a312f176f43988",
            "129c3811d1579ce766547a92340cc87890da27a1d2d1f9c6b545ca3bfaccdbb4bda2f811fe24f6d458ec6f4d982cb89f",
            "39854f0108c84cb1fbb390a9c3969f9815bca6daa12622a39ab4c65a0788b38b62c2a962b72010f8b8e185a8ddb953aa",
            "8aa7d9171217251791bbc48f98c99336bb187373ceec9e8bb4c5f7b35d69b7d63756a6455b04bfda0d8345f514e767f4",
            "2d14856f76bbf176867af1b87bf0dc36ba4e9749430526cbde409b8766ca1ca1de33fb4cd083ad8dd0e9367ed70e9fad",
            "83c36d93d3d5fcd30ba7ad552c35b0a768ae06a76d2d166caaf50ddc3ddde4dad8ee0e3acc3007826b5b33e2670b739b",
            "760d61e17106b96fe5962d79c5bbaea30120f6b60cb13b6741a65c17cc4bfffb646e00bcc0dd318abf3ce0571c2e98db",
            "40ce7a202819dc6f6ef8170eebdbd9fd43e13f91165b6533cf86197409f042f2f390e1790ee9cc22c4950b7f8262d029",
            "e71cc4fb55942196620c0265d7f9f992a866cdfa868a473284ba20bdf9d6e57e7bb69ffad37a7ebd3a1c3bd121eacba7",
            "ebcc70055ee60bd57a97e1990504ccd2855a45a52e0a95cf699c193ad242d4cf60a8882588c41a952bf10c92ec32ce50",
            "2d12c48df81c18c92d0395fb9e5daa18736ab9b59577510a0e9d3a6290c0ab8ab81163f895caa771fee8c8ce9b7daeeb",
            "ae660507237dcc0f6f592bc29bd7cac5fb738d69d8db8a53ad2b389fa29967eb1d676342d3728fe2705d355b9e9e3a43",
            "aad051745023e2348b75be8c8c2227b7b9504eb43a4787b9b529b387328a461a848b91783da6f25b55fb562d3a4acec2",
            "61c665cbecede4e9dd761e5e8b206452cfd6d9e1c78842aaadc3027a0cf7532b38770a8ea8b70488cf735b2cf2a9eefc",
            "3d5fd11c12ba1f3c353d987dfbda79d38c2893b065757249605dde8b2c7e3ed65e01d4772b75a9bf5281db7d6d119d6a",
            "d0ae8825d51cd9b30344b797d58cac474f2e8fd714b912386d18170808beab1b3abea2bf4f992403ba437c5a2daf107f",
            "c5ead535a60cae07ce663e44e6e7805fc8833f5a15ff26127eef57b4c62e593ee8a2fe25a46b72a945c309e5da1095be",
            "05213bb6a0f6f375990a6b93d4001a8bdb43823e449c80b62b57d2cbabbbe58da1ecc3581cde8a081609deef34c2cf65",
            "f39c9462a017091b34c7304e08e9e36a3168eed55aaf0486052bfd31a33d4b4ae5a9c4ae21b5947344863af573fee444",
            "e1bc3616abeec8f64ee83514c02f143869f249d74314e14c26d5e5a8df671228062f0c724eeead9466a5be7f1e308ce6",
            "5d88d1e38677226e8e9c087641b636f2b764c3cd0593896feeb9810f3c4b739c3eac163066966a0bfb3a18d238ade37b",
            "68a6088f4f3edd22f2947fd65ba86cc72ee99698b8515af57351ca90f54c9525dd81b13288172a7fc24812fd8973a869",
            "e00566c84001fd7b25db28392ab2e8a0516047dfa20c5e4ec7cd1ced9ef39b6d2fdee6d7bccee7e07056c07f3f7f18d8",
            "e780ffb424ea317eb140eaf1007f5f9cf157343c84aabf2cd5786cf8b861fd271610ffad7dfdd8f8ef5dadbfe20cf577",
            "e3ffd0e37a3ccee336c69f9046ffcf4d8dc7a33c9e9ffe138dc05fce563f1ee071a9fb27ca091e15be8f4d1f57107f22",
            "96f8d7f5c4e3111e6bf43fe140f2ffa3d81f8ff75831fe094eb27fd78f8fadff2ed6fe8acb07eb7f936e8fcdff9e27ff",
            "8a168abf98ff2b6b3eb67e1c8ffec47beaff617452904640fccd10fde163fef067b09efd76f65f504b03041400000008",
            "006d996d5d78a0e3470e290000d5350000110000006f72646572312d6d6f6465332e786c7378ed7b65541c4ab675e3ee",
            "eeee040f10dcdddd25b835a1b1e0d0b83bc1836b700f2e414270081624b85b9010ec917b67267798dcb5def7ebfdf866",
            "b3babbbad73a67515dbbced975eab48a3c1434360016000f000028003934b71139108f233800000b000fa92fee0874b1",
            "00ba186b78bcb270366479ed604f9e070d499f038004fc17ff5fa32f794c098a02bdff98b6a337a47c77f2d2246af773",
            "a09f9294e2341cbee9028d45005bc9e1227f64858dc261d12ead518a8569c7f100337262b98559245a94795b795f0971",
            "0464d6d942b3078d91f1112c5e5e0eb452e0052a538d18a6f6d454e2d03023390ef16109a6abdc573539b5e7dc228568",
            "3bafcdb7cca0c25b9d58a04a603ea9c684f63a97b6c9ef4c923710bee46be926685bd4f8c272fd8d4756423eaff93e7b",
            "ce531352ac5a4af4a5daf5d773d637dd0bb91b8302a1b6583b51ec415f76c7636e1daa63bc36a625f26b5e835349734c",
            "915beacfa6063e645c45823a0751e58433088217caaf423855d4c2e0d36dfba6e7389a2ca0fa1d9aafbe6f0f513b73ee",
            "1114181ae30ac599dd2a130f12d1f1555ff53f93015efd18fdd4e8260dc224834dbaf610904cd9652ba8af95f1391e62",
            "5ed8341097b8cba47c13bd0cd1d8c0cfeb1234862b5a691c3376a357d2abcd6e469996b12cd0426edecde258bdf77e2b",
            "7de25bac36d192cfbdc160761e5a9b0bd9e92b7fc511603e5c40301b6da0b0a697f6c58dd3c6feea0e4e53f1bccd347f",
            "9b45c76a91f1e3da3c3c400154fe6d633768b2519d3f8e141ef72dd2e3c6360659d83bb3b2fc7cfeef86fe2f7ee25d82",
            "a2722f1b72c8b1cc6de8e0d5cc521cba98bcbba84c2b8586afbfec17b486e87c6aa2a54e3a62740dc2402869bf76dfdb",
            "8f8b4b02d17b8314976b5cef2231b1f9a91add2b42defc70728563a405a54839bdfbda87e13673317399262de741efd6",
            "1ecb58de5c922b3fb4422535752a51833ec6a24d888055481c822b9816a0d7afa8bbfcf27b25c48c98fa61454875134f",
            "ce5ee25886df6c4ad12e8a70d80b8583c9349c41e79756f395de83f61417d7849aead62f6caa5f6683f8283d8ccdedbb",
            "be49647462dd793f53ca1c54a3e4c2fcce31263d409373e1a757cf753a9ce5d1c783d05b3e7832048e242b5986f8dd26",
            "0917019c3c7e0820470400d01e3f796dcfeaee08b27be9e868f733f3bdd372705ce345bff79dbc0ba55f69aff986a8ba",
            "cb58f55eeeb0cc76fe52494c2fad2fc92091779b5744f3d06fc5895c6478afcf154d147c78ba6e65999db6b1f6ad88b1",
            "0107cba17b3f5c49c8d4d0ccb5ff5b9e6ffa8e713103e1ae1e934dd654c2ac7d5e67b5d683f7f58d239e1d4bf676a6b0",
            "1d40a509cb87c1629d482ae399aee4a2729be9f377cccd252541b9af3916a4d8d29be4fbcd350969fb2b5527f23e37c1",
            "77f0c205b2e24af6bfcc76c529dfc495a84e90cb5961892c22192f3b206e8e88d8ae1425a822f44fc5b3a030f654312c",
            "e778d82bdd028fcabff2122776a05addef8935694eed95ad760169689a839bbc298421ea6b687cd684bb61291cc5aa5f",
            "44777bc83452cc95a9bcf852877904a7dbe2c57c05903a61cdc788d3aac221035f8972598d8d4cb33efe0f43a48d9333",
            "d653deb9c91f354aed681c2c12f01236b19602298ceb78b0f2ec7a2ae94d677a627677fb378d469bc486e7c16e8599e0",
            "38b8efe1c097533d3aba6036b01a87aafa10a591fcf8d7d12ae319835892672fe80d5acd1e5c824d547472584f702ab9",
            "beedb6b517824a5df33277e2856f16c0379ca79377869fc045e1a5057576d913206e0b4dcef50cc88a2dec0a1bfddc28",
            "311bca935e9a39ca3aaaebc929bd6b34c8a2b4880808d379b5957b3f32ca1391ef81328e85e3ab0f9ade28e1ef8b9964",
            "88c2a856342f70362cce0f679ca4e2f951a924ec39a51a2df4e89127f1d362b64d88a92d5fc67404e2f55c702af1b77c",
            "9115355572213af3ad7511e7e945ecf6d5de08fd667c7a14387db764a1735329a0a9789c6bb27f494ef87c9a3b5632eb",
            "db8584741fb4a3d132eaecdd9a59fc304d984d9403dbfce85d7744310d1b6d4cf1e584022e1377a173aa2a4b6ac4b684",
            "9fe23bf246c13151ee7ee8794b3631517fa8a4d75a49fd6607983c47d704d2d03796481237ad5a57301c856be1b23b7b",
            "6518e53abdec3dda5661e599f468b115c71e741c6ee6283e8e8e0e976964a8a6191d8c1ff91b4d204929915185d2f8ed",
            "1b7f8c082f42ff18c9419da9d8e6a19d2366b263eb005af88562d29333f6f507840e70f784c3e412442b724196e60cc6",
            "8cf49d26f02290d0605e449d3dc42a2d69e7090bdc620fd3e08018ad2f9163155e09163c1b6c89cd87f08e1d0a19a5ef",
            "c355a9574fa0cc886247149f9552cad48b915ea3779a84d2ddbc5f0bf401999fb71224939ca3cf08fda8a9aadc05ba09",
            "b7cc291d6c9c169dddb5ea7cf069ba59413eeebe49c8b2c914cab4c1f769495bceb4f1ca52e8142cda2684e33125ce74",
            "24fbb204e3a798281c4a15f516f6dca7c7ec656c4d32f4602d25d1bcd15eae6a0489f97c36c90bb754f403124ca3caba",
            "9981338785fe9cc19def9a96b3189af738cd9d5e63537ab4eef2e391211c411d82d3b3a487678f52f837b1625fa9079e",
            "f2318a5842030084007888c758f1674efd6bc4f8677e8578ccaf10ffb7d1fdff1e6d49dd727d6ce821c7e06395ada3a2",
            "8af711f6efe124754e0af6a1825c16828a648724963ada7343dd90e4dbb15903d264ef1fcc1286807eb68a67d1d9f575",
            "c4b58a88087abb720e76a86f076633a465e8e1f511d799413bf126c606c687cb1361e868f1932434aa1f30e170c4aefa",
            "f065832b72a2f7c7fbe536bd99c4907d46f8288854a986a75c866f5b60b1394736544027a167b809662b92db9d0d869a",
            "36e1ac4abe0dd6bd304821c5fecb88df8a1cb05c83abd85ff18e4dda2b37867fef6249147c274b1b1c3f2667cd6f1a5d",
            "4a13e6ad1846b64da416660ac657ef0b91eb804ebce3fcb476613f1b4441239042d63dd09253ca5f77ae6675a41d8cd0",
            "1e6ec6db614cb84673d3fafc509c1898f925340e65ff3599fd253f7627e0770c93abea4b6b7e644dca6346c207fc2b1b",
            "395b5b58b838b3fef1c2fe93649989dd31193fbfd7e44b12e695eeae1011954146c5f77ba48c15656ec731045d888df1",
            "87f6989a3e57a812a64c9b9815d00970b19dd703dfeabcf94a63da025d277712fa4d2510ac4d194b3312eeef2f883792",
            "78c75c3ff0e7e71658bd3ab7b938b1f0e0723bd60328c7e518caea1fdf80d9e9f454be0417c145f337d44478213bf41a",
            "380668af53ccd450476c293895ee06abcc79e1175516fa5760c0c12c1893e3f21da976da591c5772d11c753300af4285",
            "d0a52ec2e686947e7c832f94085f2b9ce83f330cf579bb4680b7617e2a7be2d39642d0b9a2a1a6a6b3ae70284e5d17b9",
            "b5788967dfdd9c2ad374b7121c6ef5a9a79518f67b809c101d83129dd61984d72db5bd5077c87eed07512d9abbba219e",
            "a12666e5406a63fb405d6c8b6e3ee3826e5a7dba6fd48a1dc0cff3c6eff5b430e9045c2499a7690c07922a52f23d96a9",
            "8997fcec385e3b597df29d7bdd4d3811f711087a2bcddaa1a7263a8296f56013b322ca5d21bdbe6d9a923768fa51d275",
            "198c65cd4a1d5c577a90e215f3259be7247051ca569aaf64a0200ad30b4292a2b27d37341fcac9aa8d678fb061c7afeb",
            "8534c5b28c4354276ce0f26eacde1dd4ef569b84f73c44e8518333c2fced6a73fcb1da498b118b3f571bc9074ef3aac6",
            "7edb9e98f5c56421b69c324f215e84809284fdb0ccc7f0c2e3e5f8a5d74d0848a61410c3e9d75b6477661fd08a29e9e0",
            "c6a5ca40625063c5d072b96aba40daf3812b4d2b89f09a36058f4ab9d4714efc1f716d7e7864c9cb69cffc6964e84782",
            "80151594fd517362ddd1c58cf44e74eaef5771df78f62a140614bbaac5d78d5fc1bd5bad2d40a198bc62d348918540a2",
            "b984496a0cb01e6c91bcb3646272cec5592335e8be001abb0631324abcb360396c0951530bdb7152671c997d71e983e8",
            "691097d86f1cf67054a870775fa44611ac2565f1599cc38e6ba551d8d2ac5f52fefde1038bc2b81921ea24c1f728c415",
            "5a869da67727049f7da8ecc77bbe38ed5ab34e21f81056c00ecfbfe6de67deaab3299e450acddab5e53d3ed7eedd7c3d",
            "4d32b23dbc4996de572a116aa82d21ce9d3b45aee7c620bfe6e9eb2bc79b8e37d67776c13328aad3450841031498953b",
            "3458a5ea3833646c1d329c59bf964cc4e6e8183bd476bd8507d535c01378f40aedb9f54bef7a4048adf6510ced3770e0",
            "7e39a3febc27332ef0ce96f1482898dcf3197b783052d64be63ce4460f99db5412958cef4d3c5e2d78b79d4dc86c54f2",
            "be309dc2a216c20a4efccef0b7290b2221203dc59db5a574b6c6196c231b44bbc26e26cdf1e5d672137bfd422fcd427d",
            "6d84aa72aec6e75686a728de5b9e54f2e2a0415be190a49d7ece955eab72c6ed044ec78a779a1f2c1a3e2cc29d9cfba0",
            "499acbcd0f434a53cddc86b9934ec413a9dda5c25e74d537199ce1fd8e73879ad3831f1ed31b0ee5cf6acf1f9c73b1b6",
            "70b0f8f3f98fe072a0b76447dd82e187b6e1877e27c4bf24b86269157b6db05279f87c7f1ee58a7eccbaa841ab8950d6",
            "cc0b1ece1405ee0a13de40cc08f32c25141d1417484c1127dc5f2d0fbaa65253bb526812485fbcf07df5c6e4a650ff5b",
            "f96455b99ce68cdacac30687e54e26bcacc7e405d8ab2354866393339df9fd72dd50f3e4659c066f9a693fd6e4859684",
            "e75b81efbbca47eb8238f1a9e471ca1ba4479ef16abb460d8cc56c519b82d909d0be19ce1866304d2f55b310bc213ccd",
            "b4507c3a75a3d915059256f1a785d2c6769372f6cb592babee9d7c2ebc8ebb3dc11881ba9834175d49e15815d93a6ab0",
            "fcfc485c62951bfe082e31bab3812d96b72fc5589243f153107d773c3baebf95a11d699f1ab1f44da3ce277c99bbdd04",
            "b5741d307339477cd8b554f3c36edcef96ea84606b88266ab878907a0bb1ba8b9dcdb5aef699802bdd97ff530d8d8adb",
            "eef8920dcc7322e349e3c5356297f4bb6ae02e4f8e829fcf0dc140d644b5e95714fa5e9ec61341e103775bdc657f0e6f",
            "672463b20bbdaf833d442a7dd6a0c66860e3c2a165ccb2c856e621ffd72ccb294714e7d70160eb398937d92f47b11a98",
            "3ebfbc0547b203b8e24c25b10ab418d2d6192440ac9a067914c90cbd4c6885bc5f3143eed8e0c9cff77c9baf797007bc",
            "a0df143c6735a16c0ae122a32e061795ed0b425618219ced2d3a5619fb3adf9d39ce6a6bcfb23e80517728d5d218e581",
            "647491b55fdcfd6e67fc7caf90ce8f3befbfb656f91ef60f73dc5fbcedf4bebbf2f37deba9631bcc6074b7b53fbbe27d",
            "b7d7d4793bfbb603647d8ce0ad17a26cad40cd4789b7c310fb15229ac1fe7e43e38bd24b4c535bd95578177e8d580d73",
            "8858add496573e2f334e5afc99d721850c3aaa9bb3255e9696986d0da8da43a0bdb0f261c9847dfd6e9b51f775ad02ee",
            "d1d087da28af88d316bb23082566b8d8148602ed3db517616192df52bf4a7c85910d13fbc1e32b15b6be2b943f5dbb4f",
            "9185c8726d903b46bf67b6c7351494ad17a0017f38857f5b0bc280d6d66f7ce11df323b325db5366d4819a19e913f106",
            "5d90aea577d81bc40196642ce2bada6eb4ea08036a65ad977afdd10590caacf0ebd65c546aefcdb3e15f2f274c9199c9",
            "d82b3591d56ada5305e759b164c1a57ee3000d8ec162c4a8d359a54b581c149f218b61e79b551a9c8803abd07dd07cbb",
            "8694dc858bfb40571958dd9c32af2ce424fd778c537abe0e87062dee3b7dd84a6dae5550eed3723b44d88b8ab9079906",
            "5f4250201baf1250eccc17a3c2fa64eb71b3ca084790971f60699b08f6b8bddbbb0992012b18d76551a295d48b5d0631",
            "7f47150a2d270ab70a4ec48f344c8feddad7dac946d4898ad9f6ee29b2b258c508c6455ffb60934487e99204f6aa97df",
            "770aaa58283c2ee437d2ee395b597cc7aef38615d8ca221af2450f16f3467fd1129cf65c14830d4b61228069091cb679",
            "2f2631ab49d5e8f271e45052bf198290aa7c25dce6ac686b1b1ea99bd53775793bdf5510f8e178eaf8e023d30dcffb4c",
            "c3e687ee65cbb03743ec01a0daace623bd8cd6cffd379d8e0e6e159bad215b95b92d76d60655266485b90c9ef8d68e11",
            "15a475a758c75faed2d4fb95a623e2a16c1a7da9db3e7fea67b3be2cf8d4bf6149c616bdaa268ca660f21517928880ba",
            "f90781e6846c90fa1c19e227962db1ad323ccc42839ec678026eafb28d1060dd757d14da35f3e211d78dbcafa558b826",
            "fb118fc93a0351680563726e4685b7ca377d751e7b3ccb72574591e1e510091b8e0c8911d1bc06526f288476e74b531d",
            "1762e605bdbb054dfd22802d515597170fb8e9a272b6c42aca441ad5f33ced527d6bafc1bdc159a8a4d89d2031990ad0",
            "88db24b83461e02e94edfb9e797ab9670a17613d89f99d3d66bf76517722461bcaf684eeb2fb43877f6b45157fb0fc33",
            "cf05ecfd199571eb06e1902a57cc9678221db5516c6651bcea7e711d17d2f4b5fafd11d3157c96756cb0f49e3beb27dc",
            "57004e4d4afdd3e3043d358a19f461fabae621be34bf9765612251fd86cac20fb4c43a2f1364a42aa1e4af06e458a691",
            "8e06a450db805e699fb6cdcb3efab2465bcb5b35c5add5997fbd9d09ebe4d7913fa369bea14e37cc516fc456660fd2c9",
            "c7b13e1e2d095a0f2477be88eaf198a34dc43975cfe7d89c5fe1db603c63b7b409fa96a036ba45f0053b0985c4b9a946",
            "7497c5cf926da7114119ef7d4f87fbbc8f73c5e714a63db362fa742dbd098ba92c4561a364d61ac6b2920f8cc34b3cce",
            "40abac0372952c8bb53b3bbf703df6344945cee5d15e3e9bed7cee06b4398ca43265e7d0256ec91e9b70f76777e4ee9c",
            "b6aec13a3c29b427235f3daeed33de76f86fb007a9f7d10c2bb5ea42ba173528c4069f1e32bcad3893a7ad9348bc36e0",
            "29d465f493db0a1b656935dad923fed6bbde46648290a33886c4644fb57370692f6a78009ab0ef0606d5715f4a75da22",
            "39527f1e7847645ad7b09357c1324bb7f75e9b87d32ed442d961b2ba06dbc0411f23e60126e4f30bfe4f2320d68d34f9",
            "64f78557c12960b7f6b737d1463ee8dcb0389d4bf89a7b0def8d6d783c1be4aa264a60f2a556d16ebe6e8b284b3a8d0f",
            "76f8e2f7daa0eb6fae88a49e561c93af5d04ccb633c4967ff1de5679beccb3220a9dbbcf1f167e3fd80d250d59f10a7b",
            "bad788b7043944b44e0d2d7444ae55735f13d6d419ff32c79a67a07d4cd005d825eac74b1c4b645af80260b074fc9cb1",
            "ddd741c5a651faee1571d49e385997c7754b50137e96b3b37e57e1d78bd5df9f2be286f5925b1e95a6242c0080f267d6",
            "7776f1b0b770fe99ef0b34466c1745d03bf07ea00e2f2727630f374417a80c6040abd153071bcd4e7e34aa15b66c50da",
            "4c314a70a73f47f1dc2f1c1f9f65df78c1c7742d4d5bf86c48929f9fffc756c6b5258f5069a108cf270d837268386921",
            "2985cdf7abb629abeeda693c5d25b51a381ad61f30db5ee208c91a778431b5621e8046532501b965d063158651f0a567",
            "2f12ebf926918a560ba6a5d490f439a9667ad8f4348ee81864766e7c2819b49b25f1876efd44a89fcf0e07bfb16bc949",
            "516e59759b3131a5fe3671ad60235e0c59f2593476ab04f507ad44c33cad4d5bec994c3691e1ee242904cb84c590f365",
            "ea511c32c2c8958cff2d570139a5ef6b152f87ecb069aa90a55774644aeb6a4cfab906acc930ec012576dd3b7abb688c",
            "ced2a3f352854e17184b2a10d98044aed1867158dd3854bed9d7d8c9efa25f37771aafdcebacdc1bac2c553876b69e24",
            "c8b1681d3aed10f7152e33e3643a53b84988bdc122b6418ed0b5d48025aaafa9d2421a2c3169d64ba53d229be014d716",
            "c577f52557ece1cab655ef623bd6af762dac312984c90b19db02367842ac9436e2422baf83a7d2b4349ad5f4e2b22036",
            "3714c2e4794e8e87a10ede4bd83270da8c4542f7ce3829cea5ab0cf6db9f2e44f71ee735cc8d8548ba7a84a185912cc0",
            "15b8efc627f276a3ad7408ddfe80ecad0d1099a2c14b29174687fd94b7dcd4e58ea8e3be98be75502b4287a8355958ad",
            "ddda0b7d7cc8f36e559ee0286ff98cc88128ecf13d52b1c6b8fb0b1d0756467c569d89c913683f3ea24e42a5d1da71a2",
            "efeb01b3fc96f13c51fd58ac7b59317489f63ba1f7bbdba3597efdd18ed9255c21ad5c65d06d4896ea07457dc3edacdd",
            "598af29d69aa394b9b42c837f7fbf427ed0773686d466f51db441f78cf6ae9e1e9ee96bf337fcc43bdda09670d5406e7",
            "f898115f5b9bb79883cb2bb4f483df1ba12334fa3180ed93c96fae84ef507ec7719ebe2be1cac7d1e9e303fb1f1cb736",
            "055998abbb806c80567f50dd605411b1870d1d66b57d95a4e7e0b9a6844402368641cb34549f5b7d80a51abfba1fd051",
            "42e2d5758707d45852ec5c19d12531a66eaa22928a8c6619942133315aeb10227664f14253f8780f16811c0f2981a5ea",
            "f59b3931aac98a86adb2ee1d8f74be02ef17c5429e43359dd6479ee7cf79c230f50559230a6384d04b1851a871717360",
            "1167219a560e90e1acfa0eef10f4f33f10dc85a001d75a2fab3c56848704ef64b6878e607e373f0d08d7d5e78f7b18f2",
            "d769d11c64eafe736aff1cfc591bd0580266b06177bcad7717e9a0f387920d81d4f43dd128ddf39c10f4b75010e97533",
            "6f952d2b3b165a4ca5d6e3b361c27bc53bb0cd2fbc72aabb9ca543465ca291b50a2343cb0e4f4f9777f4bd1220fc56b5",
            "191030993adb8c9789bb9accf84adbf7ae56608643b8ac021919a0dd72fd4ca4fc4228965252ef63757a98e066ff02fc",
            "50357727086f06930c856b9f0a7fa04ee690e643832dc3cc6237a400e93d34612f095ce1b002d98790cbbb8844ad41d1",
            "19996ed8f0b08ea0e1412d13c96d93203245755a0285f5eff6758c56296b3ab4810b11f9a32d05afd8c746b9b7d7aa59",
            "211a9aa12fe91c217b3d6e9bbaea62a1c4bf990b69274ed39ce003e94e3b0ee1135227fc4d292521bcfdbe7fbd493e1d",
            "6a0a1320cd361f9f81f1275fed2f078c613acfa8b4d8d44cf7136c9f1790255daa9d50b642b65140cfea064fa2133028",
            "acbb1256969feb87adbaa876bf58e8da3d58f2a4abcbc4ffd890d678e359c097b61437c2975881199d8dde4d39c14592",
            "bb7e923a213b39187f08f229c7d56d00cf4c4bc4d6b165f56a36cba8e1e745e7c25a58bf971ce20ffb68b1514bd65919",
            "618a68bf949d379031ba7afb3e1e0d7cab863b1c3d766853fcd0dc5e2c958ed0d7c240c78bd365d0a07dc36dc252e58f",
            "2cc7614b32bdb3e15c6dd5defc8110edeee1d84b5a2083da3d810eecbd368a2243bafa56ab4b644b2b0e4faa85f4d2bf",
            "a03d0361c6cd0b06326d9120e9ca754db7b56caec61d186237f15c30fb4307e75ca32895ba1f33f4b6d5acf635d1efa8",
            "36f2b58c671e1a001044fed756327bdc4a2ece7fbefc41b34f5afa8e8bdcd8be64d50f39083e9ad4bab5cf886a90f4dd",
            "22802db9437c268ef5128bf6f46a340c03ba130f1dbd91382f861acde77acf56e304f0b286dd85ceceef3af296519164",
            "7933be9fca6e78a5e2b451db739efab50b2271ac2422d660aa8e9ef4f9754dfa09df9d864ebd34b08f57e179ce580da7",
            "944a8c87e557ae21882aae6bef14bc8150c50e359f1e89c83017a1a883e88236ca447899678d4e48d7e1ec7f1f1c8888",
            "e1568d5755b3bdff9ce1048ca7189c56ab0aaf100f461c7c56bad9ad8f8126719c3610810d96efd0a283a912db3f88c3",
            "e68dbc31cd98c18ed0a00c5e574df2e497a0f3dd08fd1147d9580367ac72f2ad4110bbcd9486c59905bf3e3fdf18679d",
            "405387ed55b27311b9ad76d59e264e555cb3dd1ea0a782620827574230b5fe6e97e1fb6813f1de579f1e6322caeaccef",
            "92994229e43e761079a6a9c182b3da32dc4a036a822af65811f9d9a3dcf69d7662bce249952a4ec57653aa404db2cd85",
            "131bf6fae48078a31ca953fd3c7a2393081c96d83d23f54a4b48aac540b865faa9181ccc96fe5aa1900b30bc9de62b5b",
            "fdba80e569809ebb0211c2615721ce51dcebc1179bd2828102d5e3f7322dc288019198fc57a8c841dddf4b63f7e5db02",
            "2e0f0c7233ab4568194bfbe8da3edf7f9e225a0e837c38075e56f284bcdde09a2e2e02514ab3bdf6342c275e2d382856",
            "24b948839bb046b78d5ec2ded25d3cca21d1b31e792777a0312d28ac05a118f6c1a71fa86052ed920b39da40f759638c",
            "327b3f6c3b2b4cb4347a0e24312a2f70e76f0a89fe3c9c608badb8ea1363cc3d43e932164ad81d3b7aeb8b8d8d846e2f",
            "5716c50093cde760855b478f0473d5ec62513f761d581bae8f81d71cbb30c3a68e13442889be09db2bfc1a679694b2fc",
            "9c5cf83c51de0e27d2a1db8ef137483fea3cb00bee5471cd27400412c35c6917dbbba038f39591021cb3b0f9faef71c1",
            "c00c182b8f8ebc4911abe1e3b153d8ba33bcdb5741da080963f09561a83ebb75c4a5f21a24819a6f574c4ed7ae9fc38e",
            "930276bb2a2eb6476c254f51aef3f71129edc517737ef0a89555acc776d449de883f5be325409e5415209cf73890f3cc",
            "6796876befd3cce3df01ccce2fb3f51b7b26c0e4b15de88523414e64f0f7c0aa61b8055c8e3a4ad770d53b6bb0d4776e",
            "c58c69734bf82fea16b6490f2a7b877a34db8fbc7e8be5162adf1c57211adcf1ca493ea2156915b44f3c383c04efc568",
            "c6b6a055d2059bfe05b60b00ddf4e5e3b920ad554cf65ab33e53aae5b33e167bc5b24eb3af07e7a3aebd86564c7cf944",
            "5b73aae197e7b6ed2973f12ac42da3ecbe132a92a3eed890fbd81fd5c20e11720e3d9c63a06e8f33d60797a77f0492d3",
            "9c93ca3b8bdeb2deb846fde8dc7d833152f37e60eec2da8207b5bd72ee07e6ce726c08ea27451bad6741065e2fbee861",
            "bb0e41128f7f9669fda1b08193aff481a285ba36bbf7665b6994b8d9a4eb9909170933179561801d7e653e856738f826",
            "f0da3a6d20927ce88748b7533a4c3ebd1052c83c6564def22bde4ff92c30911b23d402c8f905c3bd9413ac7cf5fecce4",
            "d7e50b0284eb104908a491e481dfe4b912697b47a382946ef318a60f612280c65c0de40a4b2685fe11c8ba8eec609e3e",
            "82ed1378f6937713b050ecf6f089cb88475789a62b57e776dfcc77bc4e68b91f66444d462e584e03c0f578ad363edd03",
            "e97a9d93008c01b259ce858e1fc57e90bf0b8538556502ed8f2326080080ea3f6ab47f5efffcaad4fe71f913143bf253",
            "64041d838fd12f81939ee8f1454612278c171029e637810e98b5a7c58e8e2c60f0aa83d59b3bcba205e162d52f0c67ab",
            "59216ab8e483048e70b336cd5c0333b8d0742eb8d2c732bd0ac8379e2d25a35c2f3f2a7f410b902ce2fd3c5e8ced2657",
            "b36b2396221437fce6f300827a4a8d3ac4080ebfda7ede5cc9f721ff1f910c4a5ae202f398129ca107b76f0f17ddd85c",
            "4eb6b8e3e5859993a2bf0f8e70609a49cb06261d937486aa6d0de6ecc9f968befb7d0b09aac4f94adbe388e6d7ecffa5",
            "39fe9cfb5f95c7dfcdbe3d6cbc7c1a4e67f210ad517650aac6d131070fa67baa9aff80875ff2229deeb4a02bbf8e1094",
            "16406e67c0bf60d45bd1721043ce9be7e0044ae162ee53495e54742a4430e35b333d67c6fc501b4221e5a7163122c259",
            "eb926d0fe2c3b38cb2416e3fb33e9a4b5cb09a170d56864d3256506d22145bda224fff50bcf1603feb9a9f5bc7199b66",
            "7f8b91fbc99f82963c243b0f966116a52a61f89cb9a456c578f323cf6f276f0ea5d855f138f179280080f2cfc95bbc76",
            "b100014ded156c8076cefff6ee9fd732c07572f49063a60b12f42bd845c38a46ad83324d4d86055bd62521407722340e",
            "3e52d87e870f5c82eba46241852791656aac9faf892fa75e2fec8968beddeb174cbc50c5d553c0a5936d61b437c22f90",
            "54df2b11ba8c9ff0c0df16b57fe70e6507bb4f0ad85ef5c0af89f5f2b05f36a08b3a747d64eb2525d6d4daa988aa3fb7",
            "c3f044ab4522f2b78e2e89580bad56e3efa09ad6f2817d59e8377ff906d18135dfc8367f5ddfd2c952b35d87ff1c6423",
            "73a0e36bb56e5d76c6fa16930e4ea74200d6e00301044ca5965e77e9a89dd39da84dbf6fb46632f813e7867e09e2e467",
            "25a3658aa3f67777565173e1edd20d0dab6ba34a68210aa278dcefd132aff5dc51ae2259b2122567273732224f0e1453",
            "1386fd2d339c43b32322dd5846dd36996e9f3d43378d982757ec7dfb6a1fa7e0b31f62674e6cdaeb496fd67c6f70f3b1",
            "0c53ca15b38c97688f600707a8204f25809da1a3823712056d421d64bdeb34fa41c456fdae7a9fa501d94197180fe62b",
            "edd60cfbb73d28ce8e355d4b5b82853d6475f319d98779d549920369d9af5c727ad4ab7e61709e978ced79b7305375b9",
            "43b17ee637a8bf5becda2914f6e4c751e9e303fd1f92c7d4de4cdcdad406f873690562447fd25af20aec0e5fd734e546",
            "4129fb1e5d308094bc201f813a310121af8987e94d2cdab170bb3748cdb10c138928350403bb86d56ee34dc0c09d1b6f",
            "40c2d66bd7523682ef095d2faff80f0dcfdeac5296034dd8d1f51647c3d6a02ce8c4e7c36ece989fb77eaee950699403",
            "268b4c837a3c8863535b01ef2ecb089b9572f2020834f668fcb5d61ffaee7fcb57292d1e99a847be4e3c06320c003c84",
            "b9a3990ac8f1d5a368730459fca38df3bfb7d2803fba9ee47bd988fb8fc12bb91bcb854b2632bb8c250705e4bcaa6dc1",
            "1bc7cc8d36b2798b40337151f4e9f2d7e8a7ab48f43b4378b71c4e4776a96eac4ae7d4cb0902cabe020f74a756194ee6",
            "87e038d31ea6c167bc83904a58b8679e30159fcb2e916b0b2833e3e9bfaa883e8bb3e64ecd3eb3d99b90e8aa95514367",
            "84b621c07c43aac8642d5f3addab33f92c2c9c191f8e67cd92f265de27ad6927b91dee5228163e2582811c78aa885564",
            "eb62334a424a0212aa29128c4d8014bb043a9436662b0c9e9f074e98cebd2785366da653895e699a0dbe7bc2c88e0871",
            "6b0607311151b454aeea47b3d45b1623633d8321e906f16c57490818a17c521c110d4ac825a17bdecedd76d2f6ceec4b",
            "d615976d39f731d76d59f7a33ad76d317763972b3df731979bbebd25126f0181cbcb2b97d6a54e412f3fa1f353686626",
            "214ab567693c0bc801230df7010f893667dbae282100cb34c83651432a8b0aae4e95f2bc9c1dece1c6fd0560e9165fb6",
            "fa25b11555e4eda1afd6e02a3929ec03a71334bcf7ed3a16ec5bc06f29cc302659f0c84f2ea89fbbf02f14367df5eabf",
            "0cfe85ccc41ae0cf1be12b481fb8002f850fcfc03d39f8d84a91c8cf6cbbe4dc52f8e6251812c8c7c04bec463fb8095a",
            "cf1d6e295fc88651168b07001f96fc7e5cecca84effb98d5e1bee458c63900a5d421675e263b65809792dfb20a5533e7",
            "82926b9745bb51e3e61abc6fa4dee5a1a14f887ca4ef29043b60f482bf15348d9770512dc6abdfccd11770009babecf2",
            "bff6ef1b42a15e0b8c55f11e6018c0cc48e243cdc7355b9ffa893355f177b818934242c89fb243dbd58bfd90bb9ed9d9",
            "49ac62bfa99f3581799b0edef99ce33eb12aba055696aff9fcb127016c51610fd96d4cb886abd1549fccc12e3bb19607",
            "132d6439745bd030a27b6948b3f30af17399759278cf5e7f48de2c3775d4998bb700c9432629a9325a27bfd32d0d2f0a",
            "aa256bf198428307ebf61d20fc5b378b0980729db881cb9ac2a75c9b71ad73d02b24c0ab8b41f30b40e95a6ab2614552",
            "91df3378d93198dca8dd6d3723e74b4094aa0ffb91f47ba4c24e55f8312261c7880b2dc9929db87e2c02a2b594141207",
            "bcda2f8c4d017b1e5b8ef9ad6457bdcae4ac5596d7ad9931be9358b79bef87e900bf6376d9a535d6f347debe7d0ccecf",
            "00bf11137fcaa9ff90147f68aab68415bbc7e433b002774326b76c36a9a1d1c7680624e7c5175ff9012a9ae5ae5314c3",
            "6cb83bc6eb87c50a6abf7e4d7c757dbf8a5bde6197c8c95715352acbb7462dfa4cb1d4252bf1ee287954660d6a07532f",
            "9f9d0ecab3564f488859a83a183dd63bff9d32a5e44d2617c5dc0b3323eeb9003282483e20e33a1fbc93cbf33af25815",
            "2aac4ce5d49b61f26126cb5a06e735ee42f22160b5eab55f6329238d93c346a938095086ed008a0fe5506b779d3f88b6",
            "b67b061fa4568a9bef7c29c07a29f459b5f9eb7969049ddd2c7ace1c2182434def286e91f6b937c013624515819d11c7",
            "40412dbb30b8c59ab9b369722ee0fe63a1ba128c964a61e8d0bb4ffd66ec2396069f38d0c3068be9e3d51d95942847a4",
            "7cf66f8ccd1c8ba2268c7dedb1dc5af3cf85489628485a3dcb37f1fef5fd43403e03fcddef1c9ee237bf7a78eae0dffb",
            "a9ff8a9f0af12fddd54f0d9ff698fec23dec7f749c3e35fef7a6b3bf8211e90fe3bf69417beae7696bd12f44a1fc7da3",
            "d1532f4f5b567ec113edef1b589e7a79da84f00b4718bf6d4978eae0e97dc62f6c103cb9dd786afab44cfc0b83c4bf2f",
            "1a3ff5f0b410fb0b25247f5f967deae5698ded17b0c87f5f717beae1e9d1f4174869fe5707d5a70e9f9ef67e8198f67f",
            "75f67beaf0e909ea1708e8fe37e7a9a7fe9e8af45f8065f84fc9fed4fadff5f15f31f468fd1f6af93fccff4d9bfc1551",
            "cc7f31ff8752796afd34feff8225ebff633650918781fd6988f9f8f793415b6c3fdffd0f504b03041400000008006d99",
            "6d5de9388bd8012700006b330000110000006f72646572312d6d6f6465322e786c7378ed7a65505c6bd66ee3ee2ec1dd",
            "1d824370770feeee04770fc13db83b81e0c1ddddddadd12021d8e59c33dfe40c93a9faeeaffbe3ced3d57bbfbbbbd6aa",
            "7a6b2f79d67a97823418383a0012000d0000880171c1630ca92000402a240080068006d57e6767eb6c62ebaca7e26e6f",
            "e2f49ee1838d35512e382855360014f05ffc7f8df24459b96e26e4e05b8915e24b4c76063a6461d916aa6a613f707cfd",
            "6aca28d5fe7a5ec8a5655666fa52058ba03459fcfecc5bfee7116316edbafa1f1d127d2d529f938d14b39575f11a8a7b",
            "8e87570fa4d0133a2caad70cc8e7240c39f8f852bf958329754e7ea299cc47af2c209bd58753f9e9f749dc92505b6225",
            "0a0e4790dec0ffd290b72d61a82588214aa988b8301eaf7c73b570cbb537ab5d92440bb5f7c1d3705f1506ba1317225c",
            "723774ba5d470c2f348ef43bbea54c366b1fb263d53bcecb864e8ad3d8060adca85eb49e2b00bdd1406262163fc3fa37",
            "bda450d11b859fcd1cb075ea67f4678292c424e696c8ca56c32b49f4d80dbbfc1b5994f123a3771d44ab3f6bec91a02e",
            "fd45a16ed61ac90757d726445a96f814c21989871e78334986afb095b4b4f373ce84314f9c6c89727de462a75decfdde",
            "9f1020bed9ac7d17939ebc7deb4d3ab35d96c21c3b32aa9ff9394cb37a1caf3c8df3731918e1b7a52382cb7dd04b74ef",
            "8d82fbb067a885b921060f740d6823961fcaf4b7aa1984e20cdde2558dc3a6458a67174977fd2c8f2f2ef9fc0c0650f8",
            "178fae536522bd7a59c9bcfc0bf7e2d17a8e26d64e8c0c7f5cffebc9ffc51f288f9395ef66820f399378081db89d5d8d",
            "411696761392682156f1f1935c41aa8bca23c35b6da7c44756c10d0413f76df579185a5ee58d3a1a20bed9622b8f4445",
            "e721ad77ab0849fee9e0024543e19824e650bed983e23a7b3d7b932a2ee54ee5da1a4d53d6589c233db84e2a367d21f2",
            "05799c411d1706ad003f04932fd55fab4f5673cdf04725c8acb0f2494548750367f651fc78baef7c52e1218240d85b19",
            "e0542ac68093a1d962a5d78035f1f51daeaab2f95b8b6ac34c476e12773d63eb8eef22e9ed688f5ef4729f079448d850",
            "7fb08c8bf793675ffb6a7d65bb18ce70efe184e92e1b381f0c88242c5e03f99d93742d81bddf07030076600000a4975f",
            "3e5833bad9395a19dad959fd91f2cad56cecf2b8909f7ca69ea07adcb0f0e0c1679c7bcb2a955a4a97ac3d2945b59a36",
            "d874dc798cd254669f6edfc4c77c1d3d4a38b64d485960f5e1651cfff8f6f15ba5d4497864a5f951899c5c75e505b7d5",
            "65ae4fca919e1975da8116a566466d4dc99d8aef5ec6d31d709ab0a84c8c778d73340f12e0c2399bac54df66e0bc18a7",
            "5ece63a1a51b683c19162199120195670561e6deb0a3a356606008c4c6a0b5f40edc8834624aebcd1df8b062ecdd6bd5",
            "0482f2cd1875e42c497439ca5cd5155ea7b7c7683db7bc44cbaf917148981f29bb59cde4e9eedd2ab4ab5ec0b295a195",
            "15881e940f78b7fe274aff838de2d2019169763877d089bebe03b067180e6877b4806b03f84fef724336e4838736c768",
            "36bd2668cc4277819dd04f266ca033b88d5befcef9c34ccfac59af6ba745152f97115cd8e7eaf5340a58548fb406b953",
            "778277ee54dc0c96ca7fa4434c3618e82510d555db34373bd4e77d4f9c706cd3f7a63613dc41f1294d5ff8caa49626f8",
            "313b416746e6b2b884fcdcfe78becad1b6e3531202934ecbdef34df086824607e3398658d21db0adc5d0b1e43af7f351",
            "acc0f392ffbdd3c5cab3f248607d54aaf1f7480aa35dafbe32d7b572b82be7f18908178c21ec040a5001fc007a165f51",
            "3d570ddae7c00e3a654566223abc69cee7f550685ee42789295a6ab61fbe047a82b9b60b68f90803b176653ff12f59ee",
            "6e84250c851cad8d8dc685122011610307441d58c2338822a6167ad63be1367f68767b34ad280919c8b9c0ba0b20bfaf",
            "b9a6225e388bf4f27d825bf389f168c3135b74db2db1f46880d073f19c2a2ea96fa33dfaa629eb222273be7777e2cc0e",
            "7d16a1c3e02f4350985da3791f8f85245d9db250b904d59b43977ec448e162cd6acf87ee4120c696064ee5fc195b148d",
            "b2a7f0905c121b1113a74a067789368583b7ce832a3cbd100b2794bf52ab2a8261927db38cb5c75d2e0418b5fcc343df",
            "d57d144ba743b403e448dfa67d119fe3d91acd9e29504b297a53245a7f5a7929db6b727013868da1435ed861c3309a04",
            "1e46dc792bd55c981e4d948d036d59768becab6aee2cdff6f3275b1e9578015dfda1217234470a8d1be8798246821a44",
            "aaa670f885fba58589cca00550340b29251118630af8371140ae3aebe41e45c3224fa3c4fb48a49be67066e09ecccd04",
            "940a1cf65e1beef4c991adc317cd058a07d9eb9ff35f22ee5d6f7da9d7fc6ec90594deb6f22931ff747cdbc2ab744b92",
            "eadd221dad8ac8d0cdc5a52bd0d886f89449ee1b258e2d80507cc1b0c99c43501f07d05413417151c69255b2268d6c66",
            "abb14841069ea2e956d61445ef2ee7f70d0c1c5c1f0fd497953cd0dafbde5716bae8d88a4ec8f7b104adf1cccd239e00",
            "a100bf8b09f5c5d852582fd162f1252ee002a0415e62c25fb9f3ef91e17ff228c84b1e05f97f1bc5ffdfa33c5e48bae7",
            "85119e053c138fac16be77415974468aa8de909c27449eb2c04ba3204773b93f43c05866a0bb9edaf67cb7e6c5e4def2",
            "58c20e84e3b3308330968485af700cae4ff48bde1e1e0563829dc29ce28a04f5205bd89f576fe16463e9115739fea421",
            "f3852c1b31c773e7b34a54457e8f8bf976be735b22b194de855ca998da00a745f3ae4f26c89d26050f3ae8434e23842b",
            "b0a5f1eed95a3d496bb356398b7ec9de00022916ce28fb0a031224f081bd962c00b47da45a331e441ea435714f75d954",
            "38b5a567df32964e3b1eb87c1762d9c3641085796150035dfb1e4d4c322fd3b2e4049c554bbaa1590f2dc0f2f0cb0724",
            "fe397f8a18564f983d5bd59489f485b68bdba9ece7b9efb5c27b0280df999556eb5926cf0be5a2810000b001ff4c354e",
            "e62626ce4e8c7fde98ffb0accf2acb11cb4cc8de8c57842aadfd359fad19fa2ee262824294a85a5b13f8a92eb950bed0",
            "935273263d798acb6daf52d0050583938d8ebc3573ee9322883611bad12f9f9a4db924ca07a7296155ade8bfeabf1535",
            "930877c91d12c9a1c59e46d3f899d2e69b8cff757109c18f344c619179954bbdffa36693bfd282e0bc8201d97c213a62",
            "83e9a8b58da0b1aca2663515fc7d397a69ae59d0c24a60ce771dc9148b26da1c473364b060fbe2755b151b46f3826f50",
            "95e8a62d653c4ac9ecc2e52a392b252005f909178eef2d48d5206ebc77f88807273718bb7c6f85a73faf33254ae1dab8",
            "9c828a989a8d369c442f84c12c92acdef862c19f318379ee22788544ac51840b7cd59864d5b88dd27fc41cd14b0e77a3",
            "c19ba6e43f43b8a382e7d93464035e1506ba9bb1a48ef34c1f969f8d3bac3ebf9b26533ea83e8bc01ee877eeda764fae",
            "61b3d8db2f0a08f4fe422157de26043205e2682ff5f0a1f2ebdacf1e7282e0b684d51986b48b074309c8a6058cd3327e",
            "2ff0ba352e701cc840a89993a8e8da69f8f29ef077f9d9f343293094533d076042df6366e7ba7c61797d2a9864f3e533",
            "f4198435ad9db2426096f338ee1839bc71b69ffbf7f0692fafb09e1f612c6f0d023657de6d34cc50b0244c5b0a639dcd",
            "053555141e4829cf50ed565567b7bad5eadb0e4fb51a1e72d04db20f2fd67f7f3f6f889015ac593614077ad9170d42be",
            "07d4c4de583a5c6ad919d5b6d877765fb439bf06f9f6c82488d234775ad33939f0916b235d5c06183a43c7ba062f733f",
            "54e0fb5b6e73a23a33d0fb12e23048fe28e9ff34386773131b93bfae7fda1a506bd58aac09c51769c717f9919f67956f",
            "ddd42cfa4e67bdf284fd7811e1966adcbcb04ead0157d2c8131aca0001ea16155a475817f5322914d93126109f3846a0",
            "af5adaf18e5449e956a681376df9dac73e59ffbe40fb7bd954559994eaacd2faf30e8be9c1676849f7a9eb00cfb65009",
            "965dd634bae6b5dac1c6a99b1815ae54833eb4a96b35118f2cde1f87f2a7db7c18b1294431f23b04a71eb14a87ba7534",
            "454c1f77f932e3c07dd29d508c201a0c153360bc403c8cd410bcdb35a39865791336b067f853c70f13b28fcb182bab9e",
            "1cbcaf3dcf3a3d0250023551c9af3b12c2d12a323594207978e0d8842b77fc609c3f69ce0736993e180a3324866227c1",
            "fa1c78b4dd7d2f453a55bfd065e899415c8c5b5978d8756cea00d2b13945f41e9a2af9a2d71f778ab58330d545e1d55d",
            "3f8b65816c1ca267b26d2bcde1b0a5f9f08c7c2157703d9c58b58060c7d39bd25bdec2774e7bacb63de4cc96f1f5bec7",
            "e9cf98ac36d844a0eae6ac3fe71300ba5962aef9b17839c1e9115e6b6d0e74e129f4983bd647d9d62f9d987e5a13dcfb",
            "7cc2b399613a6d87e0f4c13fc07c412439d3700cad8e76cef021209219c01663208a96af469dba4d2de2c8a8aa934b9c",
            "48dd4d8b54c0b5891af2c8044d7475e4d378c789d9ef099e9ccfcea84fd210c24648561450587acc075aa10b7379b46c",
            "57a5e7e3f4786937afae3ecff81c807840a2944a236d4b481959b3e2e6fb30ebeb730b7775d6feb4d952e573d237ccf2",
            "749dd5eef578ebeb93e5a161194cadfbb8773cbfeef578d4d0fe309fd5e6687e06e3a515226f2e43c64d8275401dbd09",
            "12456dfdb4a3b22267886a6029b901edcca312ad620c12ad96d264ef6d987edee447b70dcaafd356dd98296258526cb4",
            "d7af680d82f4d6cc9be133e487f27d1acd0f353298a783bd351f3d232e9aac4e41e4e8a0a293a8f3d58f94de8685897e",
            "4fd914d984900c13fec9e92316b67dc89f3753734c9c01cb70a793334e756474c4361894a9e5af027d328dfd50e38802",
            "aeae5dffd6ebd3cfcf4d991e12633664747023f83b94419aa65e61c9b0fd0c8968f8b5359d48d5113a64f26a865a7d51",
            "f9a0f28cd0dbe66ca44acdc699d01fd6e2a6098d24ace51a086b54ad498373cd1832a052beb3380e8c43a27c52a6344b",
            "133101165dc20ba3e71955ea9cbfb3ad42f646f2e91894731328ea71bc4d47eb6495b0379112f53bd04beada1c0e0d5a",
            "3e76e8dd4b69ac9191ef51733d8139faf8e9c9d120f80684185e6f0387f860b10811d23b538b8351422082a80c88a6ae",
            "cfd7e55a7e741f241120a3579b418254fc55f82688ee07227f68195eb859703c76e4fbb4e88e63b5834c588d8f9ff6bd",
            "ba0acd4c3650823191b77a2d1228519d13023cbf4a1f3b04552c159c15f0e8aa775dae2f97336b2433dab6300885ac68",
            "41a2de6b2f9b06a4b20ba130a1c94cfad3ae0684ed3e098bccab92d63b0f8d9e886a3782e09296ad875b5c16eeed43c3",
            "7532faa4acede7b9f0d9f69e4d9f018768ef399b3fbf6f7cee5c330d4b1e64f677acc9683cd54a6f99ebbb6fb7b371ad",
            "d86d09d9abcc69b232d7a9d2272cc8a1f6c036b78ba820a8bd403b5bb94d55ee939b898805b3a8f721fb3637d2c7647e",
            "933fd2b7634ac814b5a1248024a3bf89098a8743d6f81347755232487981107684614f78af140bb540a7ab3e1687c3b3",
            "7427c4b6f6eeeb47a43bbae553b67b691f53e17055e6534efd6d6abcd00a9ac49cf40a2f85efdaca9cd658a6652eb282",
            "c36b2122162ce922a342b975045e6030ad4e37061acef8744b5a8f4baada85004bbcaa0e4fce8086ebcaf962b38ffae2",
            "881e57a937ca7b47756e754efcc5456e38f189a4807acc06bed5491d37fe4c9f66ba99b5ae694c98ed04ba726bd43ef5",
            "c2ce78946f08fb939a6b6ecf6d7e2d15553cc1d2f41e4be8c7b30a13e6750221552ea84db1781a4a63e8744258d57def",
            "349c09d2b6be1e8f1aac63336ca307881fb9318e60da03585549b42fcee2b49488679187a96a1b5faa285fc3d230c18f",
            "7defe5059e29f0350ce324c42ac1a46ffba51866e04efbc510bfd97aa68eec1b970ef93046994b9b35c46cd51a6f3ecc",
            "86b5f368485f9237de93a5bdcf56ae4797670ed2c8c3303f1b2b0eda0e2472bafed8e5be40118f71e196c7b2bbb8cebd",
            "4373c96c6a11f43d4e696c0f67053d01e18d53c317a143065f53a6837a1879ace6ae36b7456fa78ab924da23a322aa34",
            "35ad4993e90c5901dd44c62f34a5c5bd34c3ab9c4eb666194022850c93ad472bdf702de6545159d6b5b16e6e8bfd3c8e",
            "3aa40594845279a7d0550ed12e8b7037fa472237564b97600dce248af3d14df73bebf4ac36bf1de620e51ef261b9164d",
            "50b7c23a99e8e08b13eaac8a4b698a5a91f83b1dce024d1a5fa9bdb0318616dd8323fcefdddbdff0f461b265c7e168ad",
            "490f8037d642ef818e93d69db641b51c3762ed9670766473fde57806b57507b9150cf39447cdea9cac56a126f23653d5",
            "5fd0756cb4513e3d4384ccbde519197564dc49954e745bb20f4e0a706dcdba8fd2f546e680c4685fc5563daa6bd6b3e0",
            "f4a893aa9a2c86c813db40badfdc17941775981868f3c1eeb640d6de5d174cb9a83823dabaf69f6fa58e2e5bf1da5760",
            "5fe35c1702cf39e6090b7f1ae8041307adb0479fe9d6e52a860f11aa55420a1d956a513d56853470c2bec936e7ec6f1d",
            "e773b6ed10f2e5c28fc63328780bd0593d63a769f5b151b0a8177fb4c7ff78f48eb0c3fdae29a8013bc3c949bba360f3",
            "7aa3fdb7343366582bb1e985668a420200087f657d2767776b13a73ff27dbecaa8e5b220721bd64fc4e1b5c444f4e1ba",
            "a87c857e1470252ab260ddf9a921dd1a01d33ab9dd24dd3837aa2b048fe382898979e69db7dcb477e21405f483a23c3c",
            "3c3ff7d2ef4c39f94b0a0439475474cac0a1c4f9c564769b372c9336dcd453393b8a6b543054cc7b51bf1962f04beab5",
            "85d1b6a0021dc752440139a5e0e315ef3f42975cbe8dffca3d0557b8913f23a604a7cd4a3adbc5a4a5724a492d7170ef",
            "4d42adde288a3df8e02b48c63e3f1c9c6cd5949d24dfb4e13aab6f40f67df24ec6e25d1168f19c50f45e31e24f0a91ba",
            "450a8b6fd197129978ef0fa7084018264d069d6e524e63e061466f25fc1ed8f289487c3e2878da6486cd9086acda5312",
            "ca6d2bd16ae7e830264230fb175b751e681d22d138898f2d8a15385ca3ac2a806402e2d9c6ea2620356310b9e73fa027",
            "96477d686cd75b7fd2587fd2595fadb06b6f398f9362503b7138c0ef2958a3c3f8ec44ec2a229c8c866f011fa169aa02",
            "89f7f54b951adc40b17ea3560ac529e124eb3b75216c171f22d92eb64c4be50ea633ed6a97822ffa0510b921e37bb675",
            "1e20eb25f598e0f2db01d3a96a2a8d4a5a311920bb3b3261d29ce767c360c066114b6a568bf148f0ee5907d985348581",
            "3eeb8ba5a8eeb3dcba85f1105117f730a4b0374b50f96e87b1f15c9d48eb6dfc0f3f41bb6bfc05a7c9b192ca04902147",
            "72d71a3adc6035dc96d3f680358294b06a5305d5ea2ddde067279ce51bd238a7b96b97783678612fcf70452a136e6f35",
            "6c1869b0193526a7cec17db9f1da71e5c66a26f07e6cfbcff398c6727eec43633ccaf844196f7d10fa74b83f96e1db17",
            "659759cc16d2c2560afe0dce541958d833dccad899212bdd9eaa98bdbacb0f7fff744c75de0a5c40faa69b85f84de899",
            "ebb2860a9af271ed07dd502ee2ed413863a07c40b6b711fe9db97193714059859a7670b32e324cbd2f75807522d1fdad",
            "c023c2ef6c9cb3e756a0f26575f1f245ff878d9b1b389a182b3b3b5ad89afd69ea3a63b2b05d4cc8101bad1b6fba80ec",
            "aa222271e8283a4d33603dae5ffd4d9578947d6ded4444ecefdadcc1c613a2174af16ef051355364e11424544bc1ded3",
            "e123b50cc2a247162d35844f74a1e1487112e0982ade252f08934e55d4ed95761eb8a771e77bbd2de2f718fcd26e7eea",
            "71c5ce1986aacdc71851f0891fb99806810c13331b12761ea4611d080f65d673f208a39dd78bf3188264bbd57253e5be",
            "2e30c8f728b13f780af1bbfd51efc49971bcf830e8af52d1d8d1c0ed8fadfdcfe21fa5e2a86d1e13725bd5d7333f2f4a",
            "3f2b094c8852f0ae52a5d9939c7b40123308fd42ec09454969dbfd521fd9646162f32ea0fbe37e7a961deff0f2f684a7",
            "278d71bf2664ae2634ef1074154ffd68c759d1901e409e66c472c5498cf55c98bc71b4dd579bc02a96c0989a17a1c3b8",
            "d4977582a64e3a50302e095163ecc727161b9b46eceaa97643bc195042b6ab092826f878c153f0de3a6b6ad55d731044",
            "245f109b8e37604da3326f3e08ddfbf44a0cf7165b82138a859ff3870e77a8e88b1c6707ebc9c68b338b187ca8c14964",
            "87f319c853b4e991347b8fac33b21d3fdaac017412ec1df75f1d2cebb027f041e77217609a91073ace4a6680c96fd683",
            "9e2d7d5201459ac16342e105ab14b07efedef5e3404b84e026baebf8085e7083f924d19fd23fd5ba984a8b2ec1b58f9a",
            "afd4b3e4f30e5f240358952ac4fc45a1b04f013961b0d54e15e741a85e677a2fe45819e81a995d1f8f94522eb5bac7d2",
            "e951f811551389e781737d1ca7a5d80fe4f3c243e88c10f168b92050e9fad1eb3ac14851f4ceba76b12f9fda22e07c55",
            "7387a20dc273cd33830c34bff913e94172dc1f5a95611775cc9c5079341ec1cb31e2b98f0f125a54037fe864d5ffbc5d",
            "f6b40eca14fc7ce4f8cd08f24ad593d90ec1f4fbe5403159fc9cb0ac1c0394d35c38b9da07e0b3ef3acfd6025949223e",
            "ca0e9c9bad8823204b0b8434959c721cccc472052ce0c695c278f984ec5cc4d254616536d6e824d77078ad04a75e34f9",
            "8a0bf35ece281d5256e6b116b93c4b53ef4eefb7f9e240c88f691e1c00e086ffa72f19bdf892b3d35fb73fed6c444ddb",
            "6e9903dd87b0fa7903c65b954cb3c604ef0b9ca65c846d53ce20b7bedd5791656b2a2572ea7eddc9e7b6ee488cc5d17a",
            "d385eecb8d185eac8c6137fecb8793cadc567acc387b8eebdcda896303b5f3b8eaeb99fbebcd294277be12645a945d78",
            "f956bea4b5e5c56881cf58f1868141cd558afe9e1b418392ccd71c52b3f14431bd90ea129917b8ac31158dc53101151a",
            "500582f21bdf728504e043fd1db6de421de87115231b40795ef0dd78d59140e89ee796302973c0fa532b29803411d7f1",
            "7ea0cdcd39f0ddfe4212bc115735a3d959d6c1cd430631962b3461ca1e489228b1698e4fd14a9fabcc7805ea2d81a2bd",
            "1c283ff07c374600bb568d9afc160b7e6972b22a42520c03135e67f73375d8d26cd9cfca087509de1756b33917b313a7",
            "907b436163e78ad8f61d21d4e38a7f631c014543ae95442a533ae65b3dc000f76340db312d191e395ac90dfa72506e75",
            "552dc2d2591a32034a81b40aefd1927382f6ac5913072f7d483d0bd314dc0cd593207a2f9ccc97ae5dc555ab21550aa6",
            "1abe85ef62a9ae74b103ec49d7cce558221a6501cacb69ec57859013cd464b12e6500c15f6913dfcfb0984d7f0971dc2",
            "f7a4bba754e00c52c0151e2664c8eda2b52adb73dca55a9760c72ed0d4e27c54c273a97bedd1537752d887e974a0ac2f",
            "a8c0857bb6bcecaec827fdf3ab66b5b71d928bf2e9dc47b504227942c5ecf5ac2966362dc1aee639ecb534b1ca8af20f",
            "1a106cc498b7382dccfe9d53c1b05cf3991acade68173ae2bc6722f052363d49a48eba5b8b3cf0687ec592234ca16f1f",
            "c7ca73cfb3e28eb2cc0581e3fe0ba3c494c9f06b25b416b0e065baa40333f771fd175423df7b08e942b07c503f39f9fa",
            "a0dc3e467e6d3f889efe80bc7f378f98f82d4d226e0c1e811281340bef616df34efae0e9c19e80c207b57bfc8144ea91",
            "6a544c33cd38b097e4de92a2224570a4b03794ce79998e6bb88c3c11921befc939783e25b17de31145975c829fa49184",
            "fbaa60dc31df56c1055ecddb27ba3772df0d5bd013900fd63097c5394cde7bbb135d491b568d601375b42e51c2f14900",
            "720cf304b14e22cf3921c880bb7efb075d58eed307f4d00b2ad7e083b1b7d9eccb385d3f68c56b04d14d758fed85885c",
            "9dde2d30d174554e8e359d502585387f8a0876755bf972372506addf52ae728673ae3c4e4c50a9c5baa21ec4b7855a69",
            "64a59023afc31bb375819bb2c20d75e4c8d22567020fb90c8bb12018b90cc016dcf0d74e6deb35de1b386c54611ad60a",
            "7bbf90a6e4dbdfb57898711cdc3292c49667b5d7a8187e73e55a17b7a0a880dedcc6ecb3aa205a598f0e7a8c3e841a66",
            "03a37ff295b308e4e976d441c5cef2512e06d523baa089e029f491bbf361fdeb70d89e069f0150ebeb573c285d752b6f",
            "306746f16e88bd34b6324ca1ca16fc4aba886619104473abfcf7de6497498a455b7e3aa89ae39b17dfa95843e9536c23",
            "b14621e064c51b36704d8fcd4374b7d63bedf9db06f28209a3213bda15592efb0d9c0f8bf27aa9e0ef2dd2f60c927d45",
            "1f2c21e502994ee3aac3ae9d10e989e0249eea3918fbcca14d02924906361e204da68c1770623bc2f5d8829b5b3bc64f",
            "9c0a31834a3ccbd199620d8e10c4cefd1a77f075892033d8b58c3682d90d6a1a82861fe9de7ce07b9ce1956e25e00d1b",
            "f57541d6706a8de4d3273ae9b5d5d23bff523827cf26d873d17f1d562fffb498f5db40885155cadbfab2a205010048ff",
            "ad3ffbd709c0af2eed9ffdffa0e8d13f3846d059c019f28ded9407726ca1aec839cd354892f17da00d6acd45919d1d43",
            "40c0868d59f2a369e1924091e20af5e54646881226d1008e1dd4bc45235bff2c2638a533a6f89944b70cfcbd4753f118",
            "9be190fc0a92bf6821d7dc4411baabd497430be124fe98e1e4b97e18e5a42fca20a3183c4ac7b90bc53f06fd7e4652cb",
            "a9bde35d4415610d053e649d2cbb32399fef71c44a0bd02544fd181865413512970c4c387bd31eaab437907d24e5ad5a",
            "fefb34802872b5feed6545fe6bf7ffa41c7fedfdefc4e33fedbe356ca26c064a63ea04a95e7240ec8b9d5d361644e774",
            "350f909347f43a8df222bf23af16d731d59fc84a876749b7bba209f889882bd7c6c131898dae47217159d6a100c6887b",
            "cbe08a0eb5b7268458cc5729625490b5c639d3da911bcbf4a3057ceba5f9e942fc92d9a250b03c64829e8c6203aef0ea",
            "1e515a6fd1ceb3f5bc4b5e4e2d6b74aaf5034ace881f31055148662e24f53c4255dcf0155d718d82deee10e76f376f0c",
            "26db51f18f131f92bf366ff2c1d9c4d1d6c05ac6c2d6cae95f9efe625ef19db6db44c82167b4d76f906f2197df57d4ab",
            "014b5555a9972c1957f9019df1e018d87061c76dde50712e53b2f9151e78a629d1be3efa3eac5add90e74279561fded2",
            "728115554fdbae9eef0b20250bbc85536c96c3759e38e7847e286cfdc111ca1ce036c56b79db05bd25dccdc97c53872c",
            "64d331c4d44d80afaa7650f1f1eb95158a07520d1c9e9f795471c45668b5124f1be98c9a37a46181efe24d32ac0d639e",
            "ae65deb6b6a983a96aab06cf95a3850450c3c76cdbbcf492310b95124aa3821752a7170704a2524dabb364cccae151c8",
            "a2cf274a353160847547bb18766a4e4e778df8b4b5fcd1ece34278ab785dddc6d6981c52888c10164733d2e73b2d3784",
            "db48868c78d1f9a99df4c873a06c4adcb09f69ba53686644a42bc398eb2eed033d3db241c422916c7796fd3146fe9c2f",
            "6c7b7674ea87292fc63caf80c63309daa45b3a094fa12ebe3616c7fc5c057f66eab60aae4804a4496547f34387b15e41",
            "4be5c7ea63863a781b4d7c2c884d8abd59e6ef4760ac6d5b9aa696384b47f0cac6b392cf8b8a536f80e2929b6c525a64",
            "1bbe61501e3734adb90f10d3b53983d1bec6f788bf7bd9973e2ed0892fab92972ff23f088f81b5d13b73030bdb3f5e2d",
            "ef272158506664f0b3005f682737b3f9c030b2c5c003b82b10629e3441164d16a1abdbbec8e490dc9ce6eb8731c9d662",
            "293a6b7d5c31859e36bad82cc8d386e35368a9b4edde822ee38d77a0b5abb65fdd06b30223cb9bfd7b05cd6df8882390",
            "1232e04dc84f47f366805afd3f4534699b44411587616eb73f74ac40e8d7a9cf76988062dc43291c269bb33f8903b17e",
            "b70531354e898f2ff63af95223a000a0418ced8c141cedec5f289b9da3c93f46f5fe7b3009f873c045ba9b09bfef2c60",
            "3d6767ad60555fe290a618984fc4a5f82d78e78caede423277d9d6e89d10f24cd907e48b0d38aa8341ac07168753ab14",
            "5746b92bb2b5385e791fde67ca0bb37407e3938018832eda017aae01503934cc4b0f888ab9d21bf89a7c92cfb1549b0a",
            "42f431e61c299997164793221d35124ac834e01638a8c904b2b4e6d22533dd1a53f461e174d8509c5ba62486b9236a33",
            "0e52071c25600cdc7238fdd9d0a4111bf0e6454624b824386f48a7dfa0ec02c4984590c1d4515b20b07cdd31c2349e3c",
            "88d5293e3b146b95a45a60bbc58d1e08e2b7a4b3e0e3e14589e5280e19a53c30e8ea69e90c8ad7bdcb74110581e0cf23",
            "c0105421015de57fe26a3f6c25686dcfbc615c77de97721b77d997743bad75d91776d373bed5721b77beef395a7de3c5",
            "cb7b7373ebdcb2dacee7e9cb7f75014e47cb4fa2449fcab904ef3f5af7e4ff1c6f71b9ef821002304d05fd26f49ed4a4",
            "82ad5da12c37fb007db8fe78c9b6648f3b53f906df8c34f2e1c4476d60838800f299d5011cdaeb611b0df2f7d996ffc3",
            "5152c68b7de280fde1857f3361037bfbff5af02f7c8efb62bbcc841c720bf086f2f794798b10d0918d8d2e17094f6dd8",
            "21e79ac4bd28421d47341eb0caac7bc98153bb63f340f256328cb428a3f9d1edf970be9fb477b6b05e138c19e70c9683",
            "60970eacc8db62a546ff3692b028994eba58a7a4dc4da19a9ecc95bcfd7398e6150809ae048e68a34a3586618a7de615",
            "ac65a298318f230da155b04a1ff9adaa099b8e676b6527b1c0f07738443c707561775c44ff72f9cf6c3fda730ae6e9f5",
            "3c34acc140a42fe8c1adbe0a73173c58385f0daaa778af9b570720c90a393b4d1c9bdbd33a758616a45899184875d628",
            "71f86f1022388095be5f92494dc937b79ff21323ac357e9a581ad3b891263777809d2b364f78d765d9ff33779a83eca3",
            "63bbe7da99ee6df4f87d7fb21a3013bca8730db8739e9ec3b1efb88210c345600de641c4f249110b923ebfc5b290e429",
            "d08b97ec0cdbe394c0ec745e719a84ffc245bc48c07b28f327375503f976b66dd47dc66977971e6dd06d1d0167af5701",
            "4982335a681d21ee560807ae0d1ecb7bea86a04bef3dbb3c6f70a447e9e4ed48c527f6a2fb3757e0cf1fa62f7fdb9512",
            "a16238607b31d8ac97a84c0ff80d8bf88b47fd1b97f8934c7d8b5bb77a2153fdeb50f784526b4653545a3d3446b6445c",
            "d8efd67f3a16ce73d4ca0aa3d63d9e61f541a205b5de7dc0bfbd7bdac02c6bb38a67e5aefa3826c9bd4526442f5be29c",
            "11ff789a684a7601e20c533199120fc2a7494f488849b8201c26dea63847199beb2d97860cfccccc8807ec88fe28fa86",
            "33d1e14d1057739a668c78712ccc58d1c8a349cc2e06eb62c25b8774a51863ce85e207816595c4b8468ecb991c444e32",
            "ec1b908c10d7b2ab8b37c2688b1b96f06f4b542227df7a666279121e53add8ff54e98dafb38256b08e0aaa5f3030839b",
            "a6f5e007b4fa9d15076127c23214521316053799d3b5374c2df83f0d15b0ca41a82914845aa87da24463c633d3313642",
            "0e1b30a38a55b69393139b2bb6f37eb6028e2630ff1c8eb6e9b2754b582bcae3c93752b7ab7bf8e7bc2b08283de03fcd",
            "b0bfc66f26da5f2bf8d791d9bf63ea254cfd6d80f6b5e0eb31c25fb885fcb7a1c2d7c2ff3a6ff477c0c1fd29fc1fa68f",
            "5eeb793d60f20b8608ff79dce4b596d75303bf308df4db1982d70a5e1f40fcc220d6abe388d7a2affbbabf5081fbfb2e",
            "ef6b0daf3ba7bf1085f79ffba8afb5bcee89fdc2dd9bdf77c85e6b785d4cfe0228e9ffaab47cadf0757df63790fdafaa",
            "b5d70a5fd73cbff048f6bfa9805eeb7b4dab7fe188e2df49f66be97f65b4fff2e62901807fe3b7afc5ff954dfc1df634",
            "7f13ff07b7782dfd3a70ffc23bfaffcb30ae200d01f98720cacb27fec582ba18fe78fa3f504b03041400000008006d99",
            "6d5d26453c14b81f0000de2a0000110000006f72646572312d6d6f6465312e786c7378ed7a65541cdbb66ee3ee6ec19d",
            "d01008eeee10dcdd9dc63d1034c19da0c13510a0d1e0ee6ec1dd2140f00097ec7ddfdd7b7372c678efd7fdf1ced7a3aa",
            "57558f394757ad6fda9a4b49160a1a1b000b80070000e480d3446fbf771000c04b5800000b000fa923ea60ef6266ef62",
            "a0eae968e6acc7ec61674b96030d49970d8004fc07ff5fa32b61d81a92153bf82ab8a533c6ce298759863e8ad9b689a1",
            "5ce90e6005720d4ca5da94287558e13d85a39f1126de3517a169d8f169a0b260d1d6a0d111163ab26df7f24c3633cc57",
            "f9368274c9edbbc757899c5b29836ff1c990488883707919b980082de50db83b3091f55672e44d9212c250069c2c231b",
            "1e038615e1f6225203baf3eabdfd55a6ac05e5dc5c3818960835d3902689c67cd4e6b6991ac130c66ce3a8e75ba96187",
            "776e7e98d88086fa3d68ef90185422e75d39f5e204fedd693ceac5c00a6c53a3bbb5fc18625538fca9bc6a50fb6c887f",
            "53e535c83e360b8f5b78cf9ceda5307743591114f72d40c408019d31be90b522a8d6fb98edc18aed8cfd87366cf16872",
            "107073f7110b6eac1b413d45c30eaf2fdac2ce7d5bb3e91377ec0ba8ef3f9a3329fad9f18d17f1db6a4f4c31ed40cce2",
            "800b0ca86e890340ef371462657c6ee6bdfb9ac5a077c7dcfe5e53f1ab7cca5f61ea890db30c5c5d7be3d1838ad6620c",
            "0be35657a5c63d2ec20fd2ec708ab096f982efbfc301008f8f5000a57fd82e588d85f2c7d348eec934919e6cd7006466",
            "eb0c64fe75fe8fcdfe07bf501627afd8c9821c7222f533b4ef6a7a29065d44d65d58aa895cd5efadf4221a3832978a68",
            "a99596185d95f01d94a47fb3dfcf818525bec8fd3ef2cb75f6b20f98d8bc94b5eee521c9b74eae700c34a02409a7b2b5",
            "2e0cb7e98be9cb5449194f3ab7e66886d2faa24fb2fd2b941293a7625fd04799350811b0f2894370f95303b47be4b596",
            "8daf3f434c8ba81c958754d67166efc78fa6f9cf2615eca1088671cb1d4ea4e2f4391b5bcc7ff6e9b325bfb8215453b1",
            "e4b6aa34ce04f150781a98dab69d8ba5b562ddfbbc5448ef53a660c7bc7e352ad94b9d7de1af5dc37e3a98e1d9c589d0",
            "59daf7bd3ff00369d132c4ef8ca40d2e410d000d0090210200684f773c6c81ee0e201b6307079b5fc1ad4cddce21970b",
            "fdc16fe201aecb1d8f488eac688f017a4ccd5a432dcdee244f8a112ebd47637924854a0dfce8db4b956755cdb3d54844",
            "8dbeb0d13b3ac4c7e6f8e0a9b2d9f37e504275ab8fca628edd23553d958e54727bd62ad067d2129b912d49839a3475f4",
            "2baba081cfed868f5e7fa1a71a3f01408983d2b9589544533d9ed9b3fd6375c93675191383160b8ce88de56bf026934d",
            "527a29459d6d049c7abb98953d06a427a20458332c4152ed1bee57551df47ee782b77ed5d8ae93b9387bf48e30f18a76",
            "9113254510c6cca04f99bc4a937413ab675fcaa8d276562f43ea7b243adc0182816bc657b4ed13e36f82a5cca655915c",
            "61c766927ec43e42eb047d4568dfe6822d247f237b9638c6bd581df70ead8af0f00b6c8521f4e32b42082b576d04320f",
            "5231e7ab45dbf3751f72905f6be2e7f6fd5b24469702d0649f56eeba1c5d66a41c5ac8f4c7d6d4183b925a8a5035b8a6",
            "5a0b6d3c6fee1529ca0da4996d6b4491e335c982483b38133b840efc9df78371f402c532303c8b3717a51f0f2a690711",
            "b1ed17cffc75c8b9b17177917c821965a1381cec53b409cb8724b7a9527ceddaee537dec1f31aa9df424e5beccab2556",
            "f9706c89f3977240966f23976ba96427085b517c3711aa1435873b33eb059f09420f2993e0207554e5b7dc3f0a88ad06",
            "207bc52ed359b53d0eee2262baf6aac770a04bd696eecbf67d783c0a0eea72eae813f7c00c099b88c9410f45e820a78e",
            "c5325a896c3e590d2686b6ead6062f4a0b9b44c9217a0a6a379aa7f905af66e3cc1fdc003dae846f1f8e272a7c6733f7",
            "0bdc81d78d9e38912f08d3e0de984f7b995218bfe31c75083db03f1536b54630263446a45968106cf0d0853459f7f222",
            "954dd43ca54feb8c706b9fe193dcfe4012cdbf72c14258df499e63b44e416a428a1148eb24129ee6e1d63bca4336f75a",
            "87ecdb526415630c7d7cb4aeb0c66ca4662f61546c8080a8beb998b787a5af06f8d516dcccd8d5e58c5bf2cd037acb5e",
            "0945eb1612f31a79c7da6a6ec3ecb042dd8d459a502bfdd4a716d79f21d00ba276283d1a91fbc74a90903922742fbc3d",
            "a430a248f0def85b5266f1707c1ceb8bd20346ad8f4d538c8f16adbabbec7fc67cf7b056a8c59caaaad23f41eb611a78",
            "406e2334a247fc72d55b8ac3e51e9d572896864a38019915b11342dc3e81ebb3648ef4214d8529f8ee03daa5f4cabe9b",
            "7381f62e79a3afcb869041fcd2a11fe84bc65cff86fc165c5d6cd49dfd953fc172c015b7e0385783e0b8e3d99697976f",
            "ce911adffd654cba355f0b6aeee9dee5c6c2c2e9cae475c503fc4121846fd6f7b03172b77556503a82f8942cdf26ed9c",
            "38eac56e57be919ce36c47edfed0ecdebec94d6df134230874cd507d10c6f348c837d7bc4380a2d144bd2bee4b20b272",
            "7c6013ed760af53b6771a0d0014ff1e446cc9f1c0621001ee2c959fc1954ffee32fe4f8085780ab010ffcbfefd7f1d5f",
            "13da65ba58d0434e024f94b68f0bca1bdfdb36c2896b7ecf3b800a72f9165420dd2fb6d4d2fc29d40d49b6191b18902a",
            "fdf06812d76fef6f2d7f169959534d5c258f88a0bd276367839ad53b9b26294507af83b8c104da8d3534d035385a1e0f",
            "43478b9d20a17ed38d09872372d5852f1d5c9e1d7930d623b3e5c32882ec3bcc434ef4867270d265f067032c36dbf0a6",
            "12e87be8196e9cc98af84e2b584fcd2a1ca8e007b6ec84410a297cbb8c785e6087e51a5cc1eac8353a61ab581b7eddc6",
            "1ccf5f264d131c3b2a63c96b14594c1de6231f46ba43a41c661488afd21522d3021d7fcf36b47e613b1b444ecd9744da",
            "dedb905dcc5bfd43d9e2582318a139dc84abc580709dfaaee9f59128b17dfa62680cca8107a9ed252f762be0770c9b9f",
            "4447a97d624dd213fff001ff138e9c2dcdcc5c9c817f7cb1fe22597afc9b0fb92ce8bd2b6897bc415f6da6ac4b959b5c",
            "4e078e3adbbf2c2e79134aa2bc0ea7daabddaa3ff10bcf3f163618108124b0384b4b1d5ab938d742de25a9599351e742",
            "a58c107a35ed38af763d3baa176362ca94884292108468d0b6ba333c997c21bdf76e0f8209046dc5b0a8c39b14843ff3",
            "2e9b4695526ac72a11778a47c5bd2b82d0924d0f134a3ac7521a7a0bd30021df2fe8d3b724b6221c5ec4984bc47116a5",
            "837d4e7ed2e1489a394a2ee18b352e58b9f071f24b9fd5cb57c2ac98653e4779132e82f52da34e32043c367598757875",
            "ec93574d72816ce511d676bb7d9bdae7be7730f4f0666c7b2d8fa12c4e32ec5f9d05dbd054cf1ae2d66dcda9492bbb69",
            "8d0609b6b36493259755af159b07262bb4aa50f6bbed09f2c45fc0019bcae0271ae9d67dd23c770cbe6a712ecacfb907",
            "b3d7e02cf7a4d8cec47bf83973a53e8ef4d938554cb5f3eaeb45f585c635cd2de5759312be1e9bd8305c99da23fafeb9",
            "4ee1dab37e4037d66fc0ccb9b89d8516288938c96d27ffd27531d6ed7bc67e94ec3b58cbc26c1c895bc4f722eeb5d63f",
            "34f89d2acad4d6cd2f141e57b789838fc5b5f30f87bf5f8cf84f6d93fe76ae2bd90262599f52706c987f3bd7affe98eb",
            "8485f70bbf6c08c9174eedea8b6def7c57cf69dc700073be9257237db3d2a61385b18e84b20293ef55547a4f93bc263c",
            "12876666aa83df1e296fb3a02af57c0087f2b409a18700b51252e2a0d7eb02c116a017a54cc496eb88ac5a67ca8bf6b1",
            "bbc3fb26978a1f3d65104cc222763d2e7676f12ce65d840cf67cd34291b1227bef7574f771d8226133ebc4d5fa45565a",
            "4d70925ff36360fa2289e20463b3a818eb08db29a2a350d715ad9499998cd4c82c0696d62fdba6652aa9bd0bcb67cae7",
            "af55926188e5fb6c9d7f53d5e54b1a1886041af308f6b81ab5fb76a591c1405ff3aaa4f245dcbcf4f9c5a5092e1b33cf",
            "c4c583430734c66465fc489b85ea3587747bb649f068e50391913bc91049b242ed2bd2c94223f623b853977122855c69",
            "e07b3c2760548893432dfecdbb62419fb0c273417659e770d4263d0c94ee148540d3e54226d9859b1b7a23077d3eacc1",
            "e4956c6823086fe884fa535d06b05378f4f5807bfe6189db78d40e9e3c88eebab6c660fa81b26ffa1b85e9260a5cd102",
            "41d6662f017e0056281570df74a87e4bab643a33db736e9a7dd11751f84cdb28093199a75d578c60719bf95ca556d46e",
            "dde070cd8ce95a9003ad32a7cf0fa65550d844508e45970bda4f568f42943b83c6e5d45c3e65c10a11c81e5c3eb19ac8",
            "6c7932acc7527730d1ba3be1d01494aca7509b31d2b8c6783d7421331392a054e3874c3f4e3bea8396556e9d2e3bf235",
            "bf4a5d6d679af051e06de189375d8ed240269e3d7be3a3d625deef6876a436d5d7fd5420e250fc5ac2f983662e966676",
            "667f9efff02687da4b36540d18fe689bfee8f702bc4bfc2be616d137ba2b9f8f5e1fcca35cd18d5a1680d5eb08a54dbc",
            "e1e18c50e0ae30e17545f431cf9242d14131ef88c963047b2a65413794caca5772757c1f172efc1c930deff275ce4b27",
            "2a4a65d4a695571e375f99efa6c34b7b4e5c047ab7844abdda62fbc8d4b85cdd5f3f7119a3ca956ad4833571a12ee695",
            "c577bda778bcc18f139b4216a3b8f9e2d82b56794f1fcc50c812b1c59f1907ed97e68c61025367fc2603c107c2cb441d",
            "c5b7552b92559e2f61157f4a2075742f21fba014f8b9e2c1c9f7c2fba4dd2b10e39d1626f5455b42385679a6a6322c2f",
            "2f12bbc8e7cdb7082e515ab3ef1acc7f1a8b302786e22721faed7ab5dc9c97a01d6b9cea33774da1cec72dcefddc0235",
            "b41d32b13bbfefde3357f6c7ae3d6897688560014712812f1e25b22056f7b033d937946708d83ffaf10e7da15672db1b",
            "5bb282794d643061b0b04eecf2f1bed27e8f335bcedff78ea03763bcd2680d85ae93b3f63bbfe0a1bb35eef2db573ece",
            "4806a417da6b7d1d444a5d96a0da48fbda6f47e651cb42dbe947bc6b19e6930e28ce1e0181967362c999c6235860c619",
            "e39f811f5801ec3146e25879eaf4a91bf46220a09a6e0e79227d27235a3ed71a66c83d0b3cd98f7dbffa1b4edc5e6fe8",
            "e4bcd740438aba107652aac2c08292037ec8727d84b3fd05870a033fe7fb3387590d8d59e06320ea2e85722a83ac3d29",
            "ed87aa4577ff9fd3fe7e57483f4e5a1fd69a2afc8e7a065f3d5c64b5fadc5ff9fb6579695a07d3ebdf6f1fccaef8dcef",
            "d7b5fe9ccd6a01599e20f86887285aca51f150e0edd247af4144d2db3e6caa2e2a18631a594bafc2bbf0aa46ab9a4244",
            "aba73438fa1aa77d6f78cbb40129a0db52599f29665c5c64b2ddfbc616028ddbc297391dd6a36c8741cba34a0ef7b8bf",
            "bb2ac2fbfd6983cd318402135c74127d9ec6be32775898f879ca9ad81a8c7498c82da79f44d8c69e40ee54d501790622",
            "f38deea751ba7d937df6fea04ced0055f8a349fc9f55200c680d9d5a6e9fa8dbf4864c2fa9113b2a26a421e24dda202d",
            "739fb064c45ee6442ce2eaaa76b4caf7ba548aeac6da3d9179908a40f80d4b764ae546d34c788fe5b8495213295b853a",
            "d22a355bcae01c0be60cb894f357a0be51588c28155a8b8f6266878567c822d8b9269f75bf8bda57a0fba2f9b5f52bb8",
            "0b167681aed2b0dad9a41ccd64c4dfee1a2475ac0d86062d1c38756fa7d457c92976a9bb1d21ec47443d808c822f21c8",
            "910d5609c877e70b51617d33b539805282efc94a0fb1340cf93bdccaf6ef82a402e50caa3328d08a6a442e8398ae5105",
            "424b89c22d82e3f13fe87d8c6e3b50dfcd44d48c88daf1e928b0305bc508c6455fefb64aa0c5744908f4ae913d700a2a",
            "ff967f92cfabafd171b6b250c6aa990cb46f62160e59d486c5bcd359300f4c7d2d8cc18225371ec0b81418b6f5202236",
            "ab4659eb32307c24ae530f414859ba126e7556b0bd038fd40ef44b59dec975e5b7ef3e993c391c60bce36c4cd7ab7f6c",
            "5f360f4bee670d005565d41f6ba735cdf4dcb53ad8b9956f35856c7ffed46063a95b61489aff89de0bdfd2e17df98bea",
            "53ac93c5ab54951e85a9f7b15056b57e545f67867a582c2ff3867a36cd495922579505d1e40cd770218908a8ea6f09d4",
            "c6a58354e648118798b745b64bf030f3753b6a630938bc4b3643ecab6f6a22d06e98168ed9ef64fdcc45c2d5588f390d",
            "37e88942cb19123fa595fb289deba870dae29997baca0b0d2e878859bd4a131b16ce01bff081426876be34d2742166fa",
            "a67dff4d4da700604d54d1e6cd195877f179b6c822c25012d5eb47eaa5caf63ed81dec2c5054e84e109f4809a8c5ade3",
            "5f1ad77517c8f46b649a5aee98c445d848602ab3c5ecd128688fc7f88ab233aeb5ecfed8f2b6a9bc823758f6a5d737ec",
            "8369a5314bb06048852b66432c91a6f2083693305e658fa8a6cb8b8feb3507c3462bf8cc1bd88192fbeec0215c47009b",
            "1a85cee9499cb632f934fa205d757d3f4faabf71499850448f9ea2e0230db1a6719c94c46728d9ab5e19e629a4e35e09",
            "d4aff6dea9433ba625037ec0484b598bba98f56ad3b59fd361adbc9ab267d4f577541ff5b2556ab11559833473712c4f",
            "468a8236de91395f447478ced1c4e39cbae7beda9a5fe1d964386335b70a3a8f531ed92658c44e402171aefb22bcc7ec",
            "6fceb25b8ba088d7d8d1e23eefeb5c3e93c4b86f5248f7515d7bdc6c32435e503f11f885a1a4a89b617089d3d9de22e3",
            "904c29c36cfddec63f5c9b35555c9e6d79a493c76a2797038c36879150a2e81cbac421de6115eefef29ecc9dcdda3558",
            "933389e6fbf09ae78d6d5a56cbdb4dd620952eea4185262d48f702b05c74f0e9117d56f9992c4db558fc8d2e67be1683",
            "bfcc76d8087393feee3ef179e7c6572243846cf95124465bcaddc34b5b61bd43d0b86dbb7d5035c7a544ab359203d54c",
            "6f199151357837a79c799676bf518393cd26d44cd16ea2f20bb6ae9d0e46d4234cc80c37efd03008b8992a9be8fecd31",
            "3829d0ad39eb2e52df179d0316a775095f6d1fdc6860c5e90596a9182f82c9955845bb5bdb115214771aeb6bf1c3efb4",
            "42d7d95a114a392d3f215bbf08986da68f2e5df4d9517abdccb9220cfde980372cfca1af1d4a12b2dc117baa539fab08",
            "3944b85a192d7458a649ed400dd6c819ff32db92b3b77994dfc5be4dd89f8b389ac8289f1ba0bb74f29aa1d9cf4ec9aa",
            "56f2de9138625f94b4cdf3a621a80e3fc3d959a72d7fed62f5f78544cca07662c35372290e0b00a0fc19f59d5d3c6dcd",
            "9c7fc5fb3cd561eb0521f416bc5bd4c1e5c444ec4170649e522f06b4321d55b0feecc4807e95a03958612b493fce9dee",
            "078ad741fed8d82ceb26370fe38d244dfecb7e715e5ededbedb41b734e81e27c21ce2155dd526838490109b9adc655eb",
            "a455778d54ceb6a22a551c55cb6eccafc63802d2062d618c4d9887a0911471c0a712e8d172bd08f8e233eef81a9e09a4",
            "82d5bc290965241d36cae90e166dd5635a7aa9dd3b5f0a7a8d7a71fcfe9ffe4254af670783936d1ab293141b56dda60d",
            "8da8cec76fe4ac440b218b6684a3b78b506f69c4c0f334565fa3cfa43289f4f6265e40308f9bf53b5fa61cc720230c5f",
            "49bdfdc99e4746e1e7a1e46d9719364519b2e4484baab0a1cca8f349179808c31a5064d3beabbd87c6e02c39322f91ef",
            "7481b1a404910988671f018fc16ac5a0f2cc7a602796457ad4b71aac3c68ae3ce8ae2c953bb4367d8f9361563f72da25",
            "eeca5f66c249772677131349c622b6427eaf65ae0a4b54f3a5421da9afc8b05e3b85e698749c4d544318dfd58f4cbe83",
            "3dd35aa58de544a7d235ff8b613e4c4ec8e8b63dd80b62a5b816175a71237032555db55e593b2603626b532e4c96f3fb",
            "c920d461a398353d9bd5e807e8ce6927f9b98f4a7d3db6a7df223b4f72c073a321e2ae9e61686124dfe0f2dcf762e3b9",
            "dad1565a047ede42765605084d52e325950aa2c30ee52cd7b5b9236aba2f7cdc3eac12a245549fc8afd468ea843e39e2",
            "2c5b952538ce593e23b2230a7bba462a541d73e7d6b40332e00335c727be43fbf310b5122a8c548d115d6f04ccf29ac7",
            "7246f46001f733a268e36d77431ff6764632fc7b221d328bd8439ad84ba0bf2299ab1c16740d3603db33e4655b53df64",
            "2f6d0920df3d1cd07d6f3e9c43fbaa9f85fa55f891ebac8a0e9ef67ef99a692007f56a371cf84e3130dbd784f8c6d2b4",
            "c134b0b45c5d27b8511f1da1d69f3ed03691ecee4af01ee5771ce7ecba12fcfc343a7d3ab0ff9be39646203353151790",
            "95bdc51f54d71d9147ec604187596d5e25e9387cad262616878da1db3005d5e5561360aeccabe26fef2026e678d3e209",
            "359a103d574274498ca995228fa424a55602a5c7448cd6d48f88fda1f05b5df8580716810ce70b02f33737c973229413",
            "e5e0ed92f65dcf8f3c793edc85025efd5f5a2d8fbd7ebce60cc3d4e107becf8f12402f6240a1c2c5cd86459c85a85b39",
            "4486b3e83aba47d0c9ed26b80f41b35f6fbaacf05c11ece7bf97dae93f86f9ddf3aded9368fa42000034500000c59fcf",
            "67e6e16206b237b295b3b2b771fec7d59fcb0271ed0a90ac4fa5224d8be824afcb3c7e2e915b78480a7a24de05848cfb",
            "db90da6a2b246687931abaf71e64fab56c9323b3fe8f7b1e133a9db5abef2434cec72b5791d33bd49bcc8657f4496139",
            "d383e9f41851788a6e5021b3d893f14c76293d6887e3dd6c24d0f1ac5bde7f290d82b786510f21775bb5b408581e902a",
            "36783fe61963a709216623a6853ddb2173749da9a4e58ada25d3baec02155c1bff66143cee94ac95aea5369c84da9c4e",
            "84b910fd5973250df4a3bbb0a31f469d311a02cf494a28e043b0ba63c9b9fae20bfae1b9c72f03b164e7c837b8038135",
            "afe180271882cd653fb9250ebb15ded831d5d7d74dbe2357190b4e3c86b46899b01bbe2f93bf321d4f64f704612c7f45",
            "085de7cb9edffce6aed82fce86ebefb68521208b24c6d565cb283f6ce0b88193376340143dd62dc7679539406740e6e7",
            "9617b1c5c636a14f2cd57bf143f4c50a0177a2084c4fde3533f54248e68ca6adb46fc9a408ed5eda278473b611bcc2d8",
            "cdab7473a126c12cf7f3df3616aa265158139f46c54f07fa9f936762646b226a696465ff6baaf8a2847f1153fc2ad01d",
            "beba6ed28d9c42ba119d3fe005595e2e02557c1c424e1d27637234da8960b30f48d9a104138928250403fb0bd0663339",
            "a0f7de8d2b206edbc3b59885e03aaecdf88af748ef2c7995a2d4de90155d7b61246c1dca8c56743eecee8ce975d3cc97",
            "16a55a19fb44a12950872771744a13a0ecb284b05e213b278040759ffaadfac663d7c36f6388693f3cde87a7871b7f8a",
            "2318007808530713259083a333d0c40164f6dfadffffac7202fe68a3c976b210f79c04ae7cda5cce9f3794da6328f250",
            "464fa75b163973c75da8a11a6f5898534287be607821c6770d89e87506d32a97ee8732693f129ed9a513b7d31fbdf348",
            "d5bee986c9ab43a696f09e90821a8d9b4c738072bce853fe6d196f00717ef12abb02cc941212f647e4de5ec8a41ba6c4",
            "4ab249e4bc2821ee24aa84d12a42ed599e7356755b12cc129919f2b2faa48120eba64494f437b5c2446f073335ba112c",
            "078a0bb8e01ae763c260a54642fa6c557a60deb221b35291d574e4897d4fd678273755818768d6952ad698389f0aeb24",
            "399f7ed1fdc95b475cd4a47a9ee935a338725ad66ae24d43b7b12a27a6a706b212bee3361b826ae12bf6a3cebac9d2e5",
            "f3f683f3d3bb9b0f7bba29c7c57b7c2923762dc7e9a9c7f17bfa1f8fc3ad48b99bd39aeb9aee968fa3acea562a14f549",
            "afa0f820073f44174f9c5f46ce50acaa94fbaf3e9a7efb29ae774ffe96b716d06a305b58a59b9a593c9a53799858ca4c",
            "ce5c3c15bee3adb64259eab9e970752e720bd74f769d159ef596a54d50c0e8fab74628413f2a9ef7f40b3bd42f23fc1b",
            "838d1c1dff43e0bf901effc5fed712e315a42f5c80b75cf7cbc08e6c7c6c850fc82faddb64dc9278e6c5e8e3c8460397",
            "58f56f39089a7ed8fda4e0960ea328140db07f5cf2bfbdd8930a3ff035a9c6357eb58c73084aaa464ebf4c744a0b5c4a",
            "cc020a54327d0225562d0bb7a3c6cc817dee24ca72d0d0c78506e83af203ed303a03cff3eac68ad829176255eee6e8f2",
            "5ed9d757d8e4aef51ce841a1def08d56701d62e8c24c8be343cdc7d45b9efa8b3256f0b6b818bc8084903d6585b6a911",
            "b995b999dedd8daf60bdab993584c9fa18b83b93ed3ebe2abc1da828fb6566a0232ed0acdc16b2dd80701d57b5ae26f1",
            "15abf4f87a0e4ca48079ffcf3cf0b0d6a51ef5ae23e24c89658268c77e4f48ce2c0755c4998b0f1fc963fa8b178a68ad",
            "bc4e3fa9b95050cd8185a372604fe0ce3d20fcbc9dd91040b1410c66b724f72dd560586fedf30e09f06ea3575b0428dc",
            "484c8057c4e579bd82971d82c9f49bddf6d2b2170322def8b21e4b3622e5b7be811f251274787fa12e5eb41bd3834540",
            "b49e9444628757b5c85017b0efb9ed90db447ad5a94806ac30bf694a8ff29bc0fab9d538480bf86dee239f20f0fa89b7",
            "594fbef925e037b9c19f8da97fc910fee84e7d8d5bb1e9f8d53d80bb2395593699a0d3ee6230b127e3c2175db90515cc",
            "7254cb8b6082ef4ff07a60b1829a6f3c88af6e1e56714b5b6ce2d9782a2246a479d6a9845fca17bb64c4df1f279a539d",
            "42b820948fa7c443f06bbd2425c5259d1309936c7933431b9be3abf011fd309d154874d8161d214ec299e84412c4d5f8",
            "512b46b2281661a470e8de2c660b876d3e81db294d39c69473aee8a7e0826a625c3dc7d9d427544e2afc4b888c10b7d2",
            "1fa7242258f3abd6c8dcc5aa1fc6b9bd33f1bc490fe8161d6f55bbe3c136f04ab69141b573461648938c5e0280e6b727",
            "4541f88988cc05f4a485c10d964cad751373010f03f96c0a30ea4af9a156ea51b458ac4416baa626e8617d1674b12a0e",
            "0a0a14f292be077706260e0511e3067eb6586e4db93f044896c8499abc4ab7f0fee7fd4340be04fcbbad71cff19b8d72",
            "cf15fc737fcedff1f5c95dfd6db7ce73c1e77b16fec21aecbfec60782efccf1ee6df818af487f0bfe9683ed7f3bc53f5",
            "171c51fe7ddfeab996e73d90bfa08bf6ef3b22cfb53c5fe2fe0b9d18bf5df07eaee079b5fc171a089ed5cecf459f1721",
            "7f2183f8f725c9730dcfd3fcbfe04bf27f93f43fd7f73cf3fc0b3ca4ff9a873e97fe67d2f777dc3f49ff4b0af85cfc9f",
            "11f7ef68a5f89bf87fc7df7f7999cfbcda5f48a5fe7ff4714ab230b07fb0fae953f1f467087ef951c07f01504b030414",
            "00000008006d996d5d3eef9253b41d000079280000110000006f72646572312d6d6f6465302e786c7378ed7a65505cdb",
            "b66ee3ee044fd0e00ea12104d7e0eeeed238c15d1a82bbbbbb3bc11d12080e21587077087ec9dee7ddbd3727b7eabd5f",
            "f7c73b5fd75a3dd7ea1ab37ad61cf28d31a69c241434360016000f00004801d450ab202f0800201f0600c002c0436a0a",
            "d9583b1a5b3bea2ab9da1a3b6833ba58814872a021a9b3009080ffe0ff6b94c58dc940b2a0f71f51b6f714a8d66a1eeb",
            "85e9affab8b14a90f1f25bc135642f70a4fb1d5f8ee250f4968190dde70bbc376fbdd2d347cd6acaa51b7c7dbecdc7a0",
            "84428c744a14d66927ec1b3ea42dd722e75649e099661b71f373e0ebea866f72a325c9d7f5c528b3dcc08dcaabc821d1",
            "a7c1497223e15a609813247c476ac6945c69f4ef7a99382a5ac512190e4b18f94a9f7486d24456536329169a79b066e7",
            "416f4d9975b56d5c41986c83b0d4ab791d2ec88f6c12bfac6e80b05befd374f6fcf5c20afa9568318286000716697b48",
            "0af3cf9963265c95ccd1f29033328638ebeaaf9871b17b29ac72596824d82a806a0e1182500db3db1ecf7ebe97a21960",
            "fd31922ff731b670705fd190549270ad5407654d3db1d22d9cf94568055776dcb9446276987b9bd87c3c6705310fec2b",
            "840117afa07866972064eab939f398ed35e4172fd302720fb9bcdda46a57de2968221c7bfb64b68f0b6f65049008a30c",
            "8c8e5f4246464cc530b6db29ceabf85a9c6c160f45743c19e7e3231440ee1f665bafcc4c7efe34927ab24aa427b3d5b5",
            "3706393031febaffc75cff835f288b9196ed6146061f89df050d5e4d2f46a10b4a3a0b88b7922a79f9bcff8e561f9e4b",
            "41b8d8414584ae44e00f25e6dde67537bcb0f82e7c7790f4f2077b59282636377983733938f1c6ee031c2da57d82a85d",
            "d96a2f86d3f4c5f465b298842bb5535b246d695351b6e4d032b9e8e4897035fa18a32a0102563e11188727d957a35f5a",
            "7dc9e06705c4b4a0e24139b8aa1198b51b3b96ea3d9b50b083c217cc25b53f91fc62d0c1c074bec26310447a714da0ac",
            "68c6655e659061ff96cc55d708d479269cda8175efc1209336a840c68ef993754c6ce075d685b7461dfbc948ba6b2f10",
            "a1a774f078c82f94b868e9b74662d5dd96b0050500ac230000684f6f5c404cce36f69606363696bfe25a998a954d2e27",
            "fa83d7c4035caf336eaab3dca4b2cba4b285aa72aad5519e381d6e1a984e6f348942b9fed17380228f493ddd7e311c33",
            "f8e064cdd424039ff3fe53319d559fb0eaa9494e2c569965d990ab3c071ad6fd1ebd507bc587180e9872971eca1bacbd",
            "f331d3cc074e6d871397b1eb5c588034703c51a1215bdff13a66ba8e7b95c152cfb84b44fa35729fff78bddf788e262b",
            "e5de30338b53784a4278bb087c1f0b9e4735b543c2f1a74a75f246123f6525c87b6bf586c56095ad81ce69a577970cc3",
            "ae8b90bd9eeaf1a7e680e9fab76863c4048cdb9e30d636c1b956d01eda3f3190711dde2124a046948c681d4730ad23f0",
            "136a4ec03d6081d04f29a1537bf088828c08fdca25e27b78e856bda2e25b07ec753be11f8cd9212708ea30058f798337",
            "f79b936fe67772cae60e545a08ac1698ea54924b1ccb73d2a542156a0c72dc99d72d2f4d04cfada6125f9469ccb1f02c",
            "722e8c5c4b98c1687e6d8f61433f85e71b970669e0d149a2f7d7c548b12bd52401311d9bcef654df00f506245111672b",
            "e63b3c69f0b071b6913c02e9be08b92d6bcb374fb9c94b9f73a6791f74de37bbec3ed26c4d364817e59d21654c5af01a",
            "7ee0de4ac52c3fa4d198ac10e50a9c105e0992ac12d9ac5c9744b3dcf0864fdaa4a5850d3096f4bc7df40cee84c28820",
            "6b8b11f47fdc1a82a1d805579058914418a84e6f45e13e2e384855c181acd6a912c2f7d923e48d6e14b19b2898dfc761",
            "f0912f7721ad82b90adbb40ef3b298a95a505cf9d45af252ed0257aa984213cf324f0efda7ee178db56e2bde294b1f35",
            "18ed1d8ab05b14591d0e17acec7d8c0846587ad7f6b222a8afb5264a125dd1c78a797ef47e6838e83533256a53a9786a",
            "c8bcb0f76595a9710616f972bd213c00412e8a8bb17b92fd38672e94350d252c4c350e7c8938056e9d417e6dd7e31ecd",
            "b6bd567251b28da98d95fd56280dd76870a929dc0ad6587516d405fd0e1e8fffe5e7361047d017be056882cfd9095356",
            "9c2c2473c8716ff0741ded9bfcc3248f63e75c76573b4539eb898c6e7d3f7a2cb0ca83a84c765dcc63442324eadc4603",
            "3a721750a7b123d098fbecdf07a621a4f0751ba269459ac1c57c1785fdd94e49c8e515c09065664a21c481c01c42362b",
            "2a13a31e2116c3f98184865c7de3e1877fdb2810bc1836fceac6c702a287536916046ccdd402513921782bb6ad7ec8f9",
            "d1be70af0b06ae78f645b08f5a47b28775688d3246d6b5899b2e1fd618861726cab3f22c79b979f2dcbb273cbc79a57f",
            "c79ed5df1a40d58b5a152f136ddcf715fce9b39017f8a5d17cc64b6ea738c154b7f72ef5a3aac4238de684b9646bb395",
            "e4ad1505a7959d3d8bad29332ccd9a7e5eac09a4e74bd5ed9766a67c80dff98486223c09dc276f31ffe4170800f0104f",
            "3ee1cfd8f977cff07fe228c4531c85f8dff5e2fffb288b1590ec65460f3cf27b24fdbc58a0fd0163de112da46ae5fd2c",
            "31fa8439610ae56bac0fb747282f1618e92f26d6dc85963c985d5bef8bdfec23f1989bc218bd47442eb70f6c88f7895c",
            "1bf902c58c388133c1190ae94631b735abda0a6467ed1553da8b5093aaa6c842cd765f4f530a2fcfebfd60b696e7d81e",
            "4f2aa17b225322aa3208346fd9f0ca80b856a7e4c686bccb6e8271da6f6dba7e04a92668acd62a66327cb3d587418b46",
            "32cc3a7f010be17ff7a696c20fb2e373957a2c842c445bfca6f2828960726bef964534bd66ecfec235d8a297593f1ce7",
            "44bf06be561b4bf47d6e8645f101349b8664638b2e969fc54eb50b1aef8c2f65149b3bc2a6b572d2d7d4b9f693ab89ac",
            "c799b35ac1cddfab55bb85e008d313e5c27e4aa3f000ff1d6a1ccc8c8d1d1d98fef862f9a55969710b210bcce8e02324",
            "4f38e5ab6afa2d1011d3908950a20235cf629c0ef52927463503390d5542bbbb98ccda22257d000c21fdd86760c77964",
            "faf797c37efb9d6af9a5081b02e2084a0aa6251ad2db9147a1a3c1987bc2f11f055599b3c5aace193d1f19e72d6bab2d",
            "01625172f308dadcaa0361eacd608539fe59397d8ad9026cd444933e9015bf91b4bc7a1535325c1976490e9520c85a7f",
            "fb9d2ae6a686b564d17704832c74d48b85fa92836495614ddb7ce654abc5e896328c6aa540fb015889f7346b03caf9e1",
            "6a384d47504708ef51612be18f0e9ba5cf1e4a148bf3c655b618f959b9cddb971e92c293a31876dc1e81fa2e91f87334",
            "362b6cb46ea9144b36eae36c65b7a2735981d76bac272ac65837bd5c68ee6cbcfbf4a2a1b732950b68dcb510ba8eb057",
            "b97d1e38118550bc3f85262914b7ab965fb044c96ec70da0690f4fe4256766225e5ef6b1d9195ca32046a156c115f646",
            "6e29ec85cb7af4aa663319abd6879abcebc08b11606c08bc2db92216289f4125cdd59b135b5cfac9c93dda2591c64901",
            "d3a6d2afc5194d297a849012a46dc57057d8b9368b38297c3345b34936fc83c3d93ca9eaa545eca3416d22cb83f7beaf",
            "a148c6c38f0e3e81035e29963477682f492b3221ae724ac71313e9a4058b1046769170850709dcbaab4dc6a405a71d1d",
            "c71d0eed9c1146aa051953edd0171e3826f9b4928883847ce0b8a458537d9f3199d9a5d7364b0adc85236716e1dec49d",
            "936850d8d19607a3487c04078ff4f748bf53b203e5a9c13e3800e005d9af5cfd0f25733433b632fef3fe877eed6b2c5a",
            "52346378a3ad7ba3dff3722ff22c9b98465e6b2d571cbcd99b47b9a21e332ba8576924786fe80e0fa78f02778509af25",
            "a883799a10846e1fe54f441ac5d75f25697f4daea07025d5f82e65e1c2cb3651ef365ff3ac74a2b25442795a61f9719d",
            "d5643b0dfebdebc4859f7b7b9038eb065b0a7dcb52ed50d3c465941267b27e3fd6c4858ab05be6bb9f3bb2876b3c2fa2",
            "9348a264d75f1dba452bece8d4d31632876df064c4407ba53a6018c2341ac8a7237840b819aaa07876a887b348bf8b5b",
            "c19be24d1edb89cbda2b65aaa87cb0f3bc703fea72f3c3f057c77c7dd119f711ab3c434d01969b1b895db062dd07c131",
            "427dd6bfd9e4ce4090313e082f01d16bdbadfdfaac04ed50f54487b1770a753ee6fbdcdd867d73e73e3dbb4348df8e89",
            "823776c35e97680704737d3861fdc5a36826c4ca0e7606fb9ac20c3e7b8a17f7e7ead7724e3b5f17cd61de10ea4ee82e",
            "fc20724cb9afb2de016649797bdee20fa48f57e9afa250f7001b8e79f8f69d2d70967c583d1c9074892f345607bb09e5",
            "7acdec1bc2ad1bbe1d98442cf16fa61d70afa69b4cdaa038b8f8fa99cd092766188c62d5d3cd18dcf985b200d8a3f445",
            "b0f2546892d76884ed9994b57248e3697ae8d0f2395731c1f7ccf024e7bb5e4dd7409c0177e8c4bc374c7a648d607662",
            "8a42bf82923d1ec8721d84d3dd059b4a5d2f87fb539b5955d559a6473fd46d3285645a496b62aad09aefcede77d3de5e",
            "5748e7471d0fabad955e07fd23ac0f17991d1ef757de5e996e6a1681343af79b7bb3cb1ef7bb8d1d77b399edf6664708",
            "1e1a605933298ab764b8db3491ab10e134a08775a5ef320698fa16ef57e01db99522958c202255929a6d3d0d528f9b7d",
            "e8d72079b5daab9a32840d8a8b0c3707e44110685ca69e8c69b02e655bb4ea2e35523887437d3561ee2127cd96871032",
            "f47091093479aabb0a5cc1c1226749abc2ab30ef83056f805ea2c16b3bbcb953357ba4e9888cd75ad963d4bb86bbec43",
            "01191abe4af00793787735f618d0aa9a0d5c1e113769cd196ee2a35614f4489f89d6a902d44d3c8213110718e3b1886a",
            "6bbad0aa42b42864550c34fac3f3206599e0d7ccd8c9155a8c32e05d966226890dc541328dc435ca20f2c01c53c674b8",
            "a43356fbc131588c08452ad31461e3fdc2536441ec5cc30aad6321eb4a744f34afce211967bec25efbab54ac2e36715b",
            "6309119f6ddd84eed591a080853dbbbecda4a61a29d95e15a70384ddb088077bfdc04b085264dd157cd2edf9425458cf",
            "0c0d0e2671be1092d27d2c553d9e6ea7b2dddb00713f29ddda7432b4a23ac1cb00fa9fa8bc41a5841f4d0363f142b553",
            "223bf754b63310d5c222b63cba0b4c8d57300271d07ff499c751613ac6f9b9d749eed905947fcb3fcae7d651ed3e5d5e",
            "2863514b64b26e6514007fd780c5bcd55c30f14b7e2380c18c2535ee4bb7e817bcf120283cab4cdee038fce54044b309",
            "8280bc74f9a3f969c1e6163c52179357d2d256ee071eebbea3c9a3fd61ba5b604b9a76d363d7924970e2108baf7d4d7a",
            "d3a1466aeb4cff6d878d9553f9462b78b322bbd9d24cab528f383f9bc60dcfcc26a4fc55ed09d6d1f7ab64c57e99a990",
            "6828f3062f8a4f339ffb99cd2ef33ef7af9b103387af28f0a149e9ade24012e25334dde02b8fbf0f509c2346fcccb829",
            "b859828b99afd5dd108dcfe15eb20eb6aebdae0b43bba65f3864bf95f43211fca8cc7208d45ba3210c2aa78dcf4e2df7",
            "903bd3540482704d4a3f48f38f2c8185cd595385bf08e4d4bff280426873b8d4577324a2ffa671ff4d59b300604158d9",
            "e90ef46bbca8982d320dd31343753b4fbe54dcdcad77ae77e02d2a74c68f8d270734e034f22c8e6b39f36678b5d04f2d",
            "754fe220acc5d1978130fb550bba62313ea16c8dab2f393fb6fbb4965772074a32b87dc3de9b96fb6a56cf07aefc80d9",
            "1c4da8a6308a4d2f805bd52fa4e6f82ae547ddde17fd653cc6356c3fb15d67a6cf38b600366532cd93a3180d05d269f4",
            "11eadaa6a1b7c9de0625c1fc61fddab27c8f94446a0631e2a21550925703128c53488703a2a89facdd933f6f19950c7b",
            "31859b499a3646fda8355abd9b0eeee056933c7ddd744b91a29da5d8802dcb12a096fbc2ec68b42860cd9fc4e122acdb",
            "758e32f6c589732eebc6fcf2db75da531613f380b31885d14dfcefd871282f1d1aab057618bd4d98b71b1064715bbadb",
            "9de73d1dca6712e8760d0ba9535434c68d27d3a5f974e299aa694b8afa684716810ed6a6e9fb2472e9c63fee2dbd3f6a",
            "b0248b48b32d8df6bc35dfcae5a8479bc3882b9175085ae410e936ffe8cc704fe2cc66f121500d984079fc65d5f51a94",
            "9ad9eeb3ce12a0d8fb7a44a6551dd2b9a05e2a32f0e48026b3fc5492b25638f65a0b98af4eeb2db1193ccad8aab3bd4b",
            "74d6b3f689500f214b7a0c890e44bebd7f0912d0deb71f07755907d4725c8a765820d950cc0c9411ead7d66fe79433ce",
            "52edb6a802d92c838c65ad26aaaab1b5ac3431221e61c0335cdc9fbfd833ad274bc63b7fb30d4cf0736acbbc0dd7f144",
            "e7807dd1b188a7bc5bdfa26b0e74ab97a81c2f82c9155d41bb5ddde29715b1fb3ad8ee85d7638eaeb9b1cc9f74527e44",
            "f2e3c277b68d26b2f4bbc796dc9b25e0b20074f61e77f0c787c12e2831c8725beca91e1dce2264b040ad025ad0178956",
            "e53d65587d07bccb2c33e040db188fa375a78037275124a17e3e17406bf1e80d6d9b97959c7983d8bd2d51d8ae1071a7",
            "eb757340235eba83836667feeac54ac76fa965d488467cf313b51481050050fe8cfa0e8eae2063875ff13e4fe98bc502",
            "3f7a3bee0deac8527c3cf6487d789edc0006b4023545a0ceecc4b04e0d9f49bdcc46824e8c33f5398adb5efed7afb32c",
            "eb5c6fe9aec528f3198644b8b9b96f3653af4d80bcc5f9fcc0cf4a5aa5d07062bca2521b2d2b16092bceaac9c0cea21a",
            "a5174a667d989f0c5ef0bed76d0fa66bc5dcb71f4d12016497408f956b87c1179f72c5d6bd9d402a58c99b125540d264",
            "239fee66d6503aa4a211dfbef524a3516d12c11bbaf3e6a778333b129868d99c9520dbbce234ada74f71367e2d652e54",
            "0859342310b959847a43295c3f4f69fe29f2543c83507b67e21504e3b8f190c365d2611432c2972b719f3bf63c12322f",
            "173977ab8ce02972f0a22d15b1cc9a029d66b616533c0c8b6f9165d7b6c60e1aad83d8e8bc68bedd05c6a21c44062096",
            "7db4fe2bac7a14eadb5917ecf8b27097a60edde507b5e507ade5c5729b8ed6e31809469503bb6da2defc25fa17690ea4",
            "4ec282895844e6c821ea264ab08475d5952a4883457a4d1a499487c4e36c42aa02781fbc48a4bbd9332c143b998f34ab",
            "3ee457ebe5c3e480c736adebdd20968b1b70a065d7fc269355949a1434a2d22136d6a5822581c7472350fb2dc216346c",
            "e663a1d03dd376d273297283fda0936fe13d4739f5736360910faec168c12fbfc1e539ef44c77276a12db7f3dedd40f6",
            "d4f8f24fbec64d28e54387fd9cb3d4d8e98ca8e6bc90b2b95fc34f85a832915fa5dada037d74002c5b91c43fcc593a25",
            "b4220c7e7a462a54faeacca566c5448bc7a4363e710cedfd96b0834066b4e62be1cf35df596e936860583f16d36e7a04",
            "552c683be861676b34ddbb3fdc26a3881ddcca5e02fd09c94471bfa077a48da92b5d5ab223593e6b718317f9f6618ffa",
            "b86d7f0eed934e26ea278147ced31a6a78aafba59ff4c339a857db1f99fc65fdb23c0d89aecd8c9a8dfc4acb5534035b",
            "74d0111abc69fc40f124b7577cf728bfd37160ef155fc5d3e8e4e9c2fe978e9be9db1b1b293ada9b5b9bfea1ea5aa3d2",
            "88ddcce8302b6d2b2fbbf7df280b0bc7606368354f41f53ad5f99a28702b7a5bdb080bdb5eb7bb428dc545ce95105e12",
            "61aa274923c9892b974069d313a1b50e216287167e6bfcf8b51b0b5f02f80adf44fe3a714e907ca2bc7eb3a46bdb35e5",
            "6d9e075721afdb507587d9a1dbf91b6030a6260f53487e042f7a112d0a050e4e162ce22c44e3f23e329c69efc13d8266",
            "6e1ffe3d18cdfa47eb65a5eb32df10cfbdf8d6d021ccefd6b7bafb52cd130200a0840200c8fe5c9fb18ba3b1bdb53e48",
            "cadcdad2e11f4f7f268a315dbf9a2de023ca76a1496ec779bc5c42a78fe024f470dc0b0809671f7043ad3912a3cd511d",
            "75880b894e03dbe4e8acf7e38ecb84664fc38abfa8ead978d50a725ab74aabf197651d6258605a20b5361dcadba26b54",
            "c84cf6445cc36d7217aa2fb14e96a2e8b816ed21d5a501f016302a6052a7153353dfa561f162dd90afae51566a10c296",
            "c2ead8b3dd12073f33e4d43fa0f64a742c39420536c4ca8fd58fdb25aaa7a92b7f49406d4b23c45c88ac505b4eb53fef",
            "2bec1e8251a18b84c0b513e7f70d0d54b12d3953f9fe8ae6cbdc63f57034c919f235ceb05fdd1b38a6230cbeb6b23b2e",
            "d1fd3e19792bfaa6a6c6497f52c5af81f18790a6ed13565feecba4af8cc6e3d95ded31963e2104fd789735bffecd5976",
            "48840dc7db69038357124998b3174427fd45d776ed45de8c2e61e4d73ea977e619c3d4ba245e4e79611b6c6c133a44e2",
            "0317e742af96f1b9e20561faf37e32be5e0067cca881de7b964c0a52eda466239cb18de21646af5fa599f0b7f2653a9f",
            "fdb68c7ceaf5013efe6954fc74a1ffb97986fa204321337d73eb5f5bf52e4200f169aba08ffcbce11d9c4d67fd8329e6",
            "fdb791ce2148b953f859d55905ceaffa4313c139d92d1777a3efdb8a24e8417a04a272bdedf4d199b0878d7b87f01229",
            "6b7df9dd462b4290b58bd675ce4399fea1652dbe7dfc66563ca4216871e9c8c6af0fbfe44eed6b0cdc08abd3358b40ca",
            "8f205cadb9747e87d1ab579dee34867c710b27b79368f6e6416c1ff7774b301a82c70d7d5adcf8531cc100c04318d918",
            "cad9dbd83a3019dad81bffabc7fb9f6217e08fa689640f3351ff91df72f6fa52febc9ef80e6d918b027a1af592e0a933",
            "ce421dc578f3c29c1c3af405ed2be1773f2111dd4e613aa4d2bc5026ad473f66f46ac66c0d456e3d5274ad3b61726b92",
            "28c7851090bd46e322511b261f2fcacebf29e3f625ca2f5e619781999243c24e411e18804cb8a68faf229944ce8be0e7",
            "4aa0881bab21d0987d7bc6a2027a89592231435ad694301c60d11a8f9226df2040e83392a1da8760365c5cc009d7321f",
            "150c2b3e0a1e0429f6c3f8b021b35090d475e7091f27aafa4b4d55e2221af7260bb7c4cf27c3da89cda75df465bb6b8a",
            "0819d6ced3bfa113414ecd5c89bf6eee33500262baaa22cbe1d96eb2212815b2b21ff4344e962e9d75ed9d9ddc5e87ee",
            "68251d16efbc4b1ab56a3f4c4b3e8cddd14939fc684eccd596dad6d87abb741861deb85c29ab437c05f50e722434b278",
            "e2ec327c866c45b1dc7be5d1e8db9d88f63da90f7703a04377b6b0462b39a3782ca76a3fbe949194b178eae396bbf232",
            "79a9ebbacdd599e00ddc10c9cfcc8f993ecc9d7cbcfa3f7f6b84bc2ebb09e94fbfe043fd32c2bf69b0beaded7f14f82f",
            "a4c5545bff2a305e013ce17cdda5b850fc3ab3f0b0654291690c3a659c12dece0bd3c4908cf92db2e89c72e0d7ae5bdd",
            "9171bd0f262f4c6fb9777edc991d20ef9b2e68508762c13f42e478b5410f55e869febd46ef2a94b830915eb248abb8cc",
            "59ae8a81c2e975475ab0fa3904198138be489352d50b8324db8c73448b7851236e7b5a62cb40a5fed757cac6ec5aee6d",
            "155da47c236748a884d0aa82ae04a8be65b269ec3f3bb2f3671974ddd440501092270cd09675826ff3efcc1dcf875493",
            "3c97cdaafcd0a4051c1dbeee99d9d2397405e527591aeb4b74d52870f8ae10a3d84195687f934a4eca33b39df01125ae",
            "357af8fa6d54ed52f2b5991de24c91599c50b7c5c04dce240745987d87fbd291ce55e4d8ed40a2ca7e067461d7d2fefa",
            "716a36c796fd779428ce5720283712d608795c5886bc568b02b2077f8f771447786e87af4c0f67e527c9784f3e8815f2",
            "790e67dcbca56e7cbd96651d7e9b7ed8d3ad4b177055ff0ad8e7914f16e78815544f4cf003cc416045c8aa4dd31870ea",
            "b96993eb098d762f99b8162afff0a6f0f6e539f4a3cbe4e96f49cf7cec441ffb93c2663e396506c06f48c19f8d897fa3",
            "067f74273ec52c5b3eb1a18165b85b628925c3096a8d5e5a436b124e3ca1e51bfb82598e5a6941ccfafb23dc7e58ac80",
            "b66b17a2abeb87159cd276cb58b6b79561a3efdffea01060902e764c8fbd3f8c37a138817044281f4f8a85e051672026",
            "c6219e130c166b979fa18acef1944941df4f636122dcef8c0c1379098cb77b19c0d992a21e2556148d305af8f9de386a",
            "e305db7c1c975daa42941170aee88e6f41293ea689e3742a1b1548817709910e762a3d3f79298835bf6281cc55ac143a",
            "cee59e81eb4ebc47fdddf646a92fb6de125e0e141ed030a76f8a3449e7c60b68f3392a0ac08b47642ca0212e0c6c36a3",
            "ef689c98f37d18ce6793815191cb0f325789a0c2622134d53232440f1e34a58e56b4919121dbc86dbfb8afac062af66f",
            "33f13520b42ecabb65461c464568b74d9ffe773515029201f03f9d7d7a8edf9c847a3ec13f4f61fc1d8d4f6eea6f6732",
            "9e0b3eef4cff85efb0ffd6a77e2efccf16d6df718cf887f0ffd0d07a3ecff39ec55f1043f99f3b18cf67795e94fe0b65",
            "68bf2d513f9fe0797efb1772709f65bbcf459fa70d7f014cf0fb24e2f90ccf89f95f3021fcbfa1e9cfe77bce15ff02f9",
            "cb7f678ecfa5ff49d3fe8ecd27e97f236dcfc5ff1923ff8e2292bf89ff2b623e977eee8efe823bf9ffa37392938481fd",
            "2588f2f4317ffa333714bf9efe0b504b03041400000008006d996d5d7d965c3a2f020000c02500000d000000636f6e74",
            "726f6c732e6a736f6eed9a416f9b301480ff4ae473696c4320ca2ddb6eabb4c3b6531d4d0ebcacac0e46181aaaaaff7d",
            "c0d28696e26a2d12b8f2cde6d9d643df27193f7c798776b100b49a2199459061672f23c0e7a550253a3b3eaca2b86ad7",
            "916333e5595ecf29c51cca1cb2848b8b38b9564f7ae4bcdc8b7a919c67bfa1192f0f0944cecd963ba14c6f9d87d124c0",
            "0e851ddf91adbf8c5ccfa381bbc474e9720a00388010b0a3649185e09c72e35b254591c3afd3faf5bbace6f3f9e71563",
            "3f15648ab1ecb6f81333b64ed32f3ce78c5dc8900bc67ec03e654c967193cd81abbd13f2f00a1c8aa94fb04baae090c9",
            "aa2b805c55395ea2ef4dac7ef815d21c6daa46084234c1bb7f03eb776982c758dd5f93ba771d2751d54b0a21aade0d17",
            "45038f628cee37f767b31779121d4f322e4f62124f320d9e54c7938ecb939ac493be83e769d82351da228a549ea13652",
            "99006a806a56f844354e10fceafcb5abcb203fc8d73370b5563e9b3fa0d3aece69775ca75d939c76add35371dad339ed",
            "8debb46792d35ebfd3b36fc9a3d74d5bb7ff3e0c1e80ed42c776312edb85496c17d363ebebd8fae3b2f54d62eb0fc0f6",
            "e9f0f7ed4bcf5779cbded4c9e44dfb532793ffdea306f43dd0f91e8ceb7b6092ef81f57ddabe931edf89f5ddfafe617d",
            "ef7ecfb47cb7df33f67be683f9de3d9bb57cb7673343cf66a4a7a6d2626b6b2a86d654484f0db8c5d6d6806d0dd8a81a",
            "30e9f957d772dafeabb3ffea0c74ba7b9fa0e5b4bd4f60d07d02d273dfa7c5d3def799ee7d9fcd5f504b010214001400",
            "000008006d996d5d96375a7ebc1d00007a2800001100000000000000000000008001000000006f72646572302d6d6f64",
            "65302e786c7378504b010214001400000008006d996d5d81f19222c01f0000e12a000011000000000000000000000080",
            "01eb1d00006f72646572302d6d6f6465312e786c7378504b010214001400000008006d996d5d2fc7eafe092700006e33",
            "00001100000000000000000000008001da3d00006f72646572302d6d6f6465322e786c7378504b010214001400000008",
            "006d996d5d6a04c3a40c290000d63500001100000000000000000000008001126500006f72646572302d6d6f6465332e",
            "786c7378504b010214001400000008006d996d5d12159c09d21d00008e28000011000000000000000000000080014d8e",
            "00006f72646572302d6d6f6465342e786c7378504b010214001400000008006d996d5d9d28e5b5cb1f0000f12a000011",
            "000000000000000000000080014eac00006f72646572302d6d6f6465352e786c7378504b010214001400000008006d99",
            "6d5d6c08b5492a2700008e330000110000000000000000000000800148cc00006f72646572302d6d6f6465362e786c73",
            "78504b010214001400000008006d996d5d97321c6d28290000f33500001100000000000000000000008001a1f300006f",
            "72646572302d6d6f6465372e786c7378504b010214001400000008006d996d5d1ae087de28290000f335000011000000",
            "00000000000000008001f81c01006f72646572312d6d6f6465372e786c7378504b010214001400000008006d996d5d46",
            "759602222700008a33000011000000000000000000000080014f4601006f72646572312d6d6f6465362e786c7378504b",
            "010214001400000008006d996d5d9018dd50d61f0000f32a00001100000000000000000000008001a06d01006f726465",
            "72312d6d6f6465352e786c7378504b010214001400000008006d996d5d52296b20d31d00009028000011000000000000",
            "00000000008001a58d01006f72646572312d6d6f6465342e786c7378504b010214001400000008006d996d5d78a0e347",
            "0e290000d53500001100000000000000000000008001a7ab01006f72646572312d6d6f6465332e786c7378504b010214",
            "001400000008006d996d5de9388bd8012700006b3300001100000000000000000000008001e4d401006f72646572312d",
            "6d6f6465322e786c7378504b010214001400000008006d996d5d26453c14b81f0000de2a000011000000000000000000",
            "0000800114fc01006f72646572312d6d6f6465312e786c7378504b010214001400000008006d996d5d3eef9253b41d00",
            "00792800001100000000000000000000008001fb1b02006f72646572312d6d6f6465302e786c7378504b010214001400",
            "000008006d996d5d7d965c3a2f020000c02500000d00000000000000000000008001de390200636f6e74726f6c732e6a",
            "736f6e504b050600000000110011002b040000383c02000000",
        );
        let bytes=PACKED.as_bytes().chunks_exact(2).map(|pair|u8::from_str_radix(std::str::from_utf8(pair).unwrap(),16).unwrap()).collect::<Vec<_>>();
        let mut archive=zip::ZipArchive::new(std::io::Cursor::new(bytes)).unwrap();
        let mut text=String::new();archive.by_name("controls.json").unwrap().read_to_string(&mut text).unwrap();
        let rows:serde_json::Value=serde_json::from_str(&text).unwrap();assert_eq!(rows.as_array().unwrap().len(),16);
        rows.as_array().unwrap().iter().map(|row|{
            let mut bytes=Vec::new();archive.by_name(row["file"].as_str().unwrap()).unwrap().read_to_end(&mut bytes).unwrap();
            (bytes,row.clone())
        }).collect()
}

#[cfg(test)]
mod tests {
    #[test]
    fn table_reference_order_distinguishes_absent_and_explicit_empty_sections() {
        let xml=format!(r#"<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheetData><row><c t="inlineStr"><is><t>{}</t></is></c></row></sheetData><tableParts count="2"><tablePart r:id="z"/><tablePart r:id="a"/></tableParts></worksheet>"#,"\u{65e5}\u{672c}");
        assert_eq!(super::parse_table_part_references(&xml).unwrap(),Some(vec!["z".into(),"a".into()]));
        assert_eq!(super::parse_table_part_references(r#"<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><tableParts count="0"/></worksheet>"#).unwrap(),Some(Vec::new()));
        assert_eq!(super::parse_table_part_references(r#"<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"/>"#).unwrap(),None);
    }



    #[test]
    fn native_external_link_identities_and_unrounded_caches_survive_parse_and_recalculation() {
        for (bytes,row)in native_external_link_corpus() {
            let mut workbook=crate::parse_xlsx_preserving_values(&bytes).unwrap();assert_eq!(workbook.external_books.len(),1);
            let book=&workbook.external_books[0];
            assert_eq!(book.origin_part.as_deref(),row["part"].as_str());
            assert_eq!(book.target.as_deref(),row["target"].as_str());assert_eq!(book.absolute_target.as_deref(),row["absolute_target"].as_str());
            let expected_names=row["sheets"].as_array().unwrap().iter().map(|name|name.as_str().unwrap()).collect::<Vec<_>>();
            assert_eq!(book.sheets.iter().map(|sheet|sheet.name.as_str()).collect::<Vec<_>>(),expected_names);
            assert_eq!(book.sheets.iter().map(|sheet|sheet.cells.len()).sum::<usize>(),row["cells"].as_array().unwrap().len());
            for expected in row["cells"].as_array().unwrap() {
                let sheet=book.sheets.iter().find(|sheet|sheet.name==expected["sheet"].as_str().unwrap()).unwrap();
                let (col,index)=parse_cell_ref(expected["cell"].as_str().unwrap());
                let cell=&sheet.cells.iter().find(|(r,c,_)|*r==index+1 && *c==col).unwrap().2;
                match expected["kind"].as_str() {
                    Some("str")=>assert!(matches!(cell,CellValue::String(value)if value==expected["value"].as_str().unwrap())),
                    None=>assert!(matches!(cell,CellValue::Number(value)if value.to_bits()==expected["value"].as_str().unwrap().parse::<f64>().unwrap().to_bits())),
                    other=>panic!("unexpected full Native cache kind {other:?}"),
                }
            }
            let json=serde_json::to_vec(&workbook).unwrap();let restored:Workbook=serde_json::from_slice(&json).unwrap();
            assert_eq!(serde_json::to_value(&restored.external_books).unwrap(),serde_json::to_value(&workbook.external_books).unwrap());
            assert_eq!(crate::editor::XlsxEditor::new(&bytes).unwrap().save().unwrap(),bytes);
            crate::formula::evaluate_workbook_formulas(&mut workbook);
            let copied=workbook.sheets.iter().find(|sheet|sheet.name==expected_names[0]).unwrap();
            let e1=&copied.rows.iter().find(|row|row.index==1).unwrap().cells.iter().find(|cell|cell.col==4).unwrap().value;
            assert!(matches!(e1,CellValue::Number(value)if value.to_bits()==200.0_f64.to_bits()));
        }
        let legacy:ExternalBook=serde_json::from_str(r#"{"sheets":[]}"#).unwrap();
        assert!(legacy.origin_part.is_none() && legacy.target.is_none() && legacy.absolute_target.is_none());
    }

    #[test]
    fn saved_array_recalculation_is_driven_by_aca_not_cached_results() {
        let xml = r#"<worksheet><extLst><c r="Z1"><f t="array" aca="1">1</f></c></extLst>
            <sheetData><row r="1">
            <c r="A1"><f t="array" aca="true" ref="A1">1</f><v>1</v></c>
            <c r="B1"><f t="array" aca="0" ref="B1">1</f><v>1</v></c>
            <c r="C1"><f t="array" ca="1" ref="C1">1</f><v>1</v></c>
            <c r="D1"><f aca="1">1</f><v>1</v></c>
            </row></sheetData></worksheet>"#;
        assert_eq!(super::pending_array_calculations(xml).unwrap(), vec![(0, 0)]);
        assert!(super::pending_array_calculations("<worksheet><sheetData/></worksheet>").unwrap().is_empty());
    }

    #[test]
    fn physical_empty_cells_preserve_the_imported_used_extent() {
        let xml = r#"<worksheet><dimension ref="A1:Z100"/><sheetData>
            <row r="1"><c r="A1"><v></v></c></row>
            <row r="3"><c r="C3"><f t="array" ref="C3:C8">ROW(A1:A6)</f><v>1</v></c></row>
            <row r="5"><c r="E5"/></row>
        </sheetData></worksheet>"#;
        let sheet = super::parse_worksheet(xml, "Sheet1", &[], &super::StyleSheet::default(), &super::Theme::default()).unwrap();
        assert_eq!(sheet.pending_used_bounds, Some((0, 0, 4, 4)));
        assert!(sheet.rows.iter().find(|row| row.index == 8).unwrap().cells.iter().any(|cell| cell.col == 2));
        let empty = super::parse_worksheet("<worksheet><dimension ref=\"A1:Z100\"/><sheetData/></worksheet>",
            "Sheet1", &[], &super::StyleSheet::default(), &super::Theme::default()).unwrap();
        assert_eq!(empty.pending_used_bounds, None);
    }

    #[test]
    fn extended_conditional_rule_does_not_create_a_second_base_rule() {
        let xml = r#"<worksheet><conditionalFormatting sqref="B2:B6"><cfRule type="dataBar" priority="2"><dataBar/><extLst><ext><x14:id>{ABC}</x14:id></ext></extLst></cfRule></conditionalFormatting><extLst><ext><x14:conditionalFormattings><x14:conditionalFormatting><x14:cfRule type="dataBar" id="{ABC}"><x14:dataBar/></x14:cfRule><xm:sqref>B2:B6</xm:sqref></x14:conditionalFormatting></x14:conditionalFormattings></ext></extLst></worksheet>"#;
        let sheet = super::parse_worksheet(xml, "Sheet1", &[], &super::StyleSheet::default(), &super::Theme::default()).unwrap();
        assert_eq!(sheet.conditional_rules.len(), 1);
        assert_eq!(sheet.conditional_rules[0].kind, "dataBar");
        assert_eq!(sheet.conditional_rules[0].priority, 2);
        assert_eq!(sheet.conditional_rules[0].ranges, vec![(1, 1, 5, 1)]);
    }

    #[test]
    fn average_conditional_rule_reads_excel_attributes() {
        let mut element = quick_xml::events::BytesStart::new("cfRule");
        element.push_attribute(("type", "aboveAverage"));
        element.push_attribute(("aboveAverage", "0"));
        element.push_attribute(("equalAverage", "1"));
        element.push_attribute(("stdDev", "2"));
        let rule = super::conditional_from(&element, &[(0, 0, 2, 0)], &super::StyleSheet::default());
        assert!(!rule.above_average);
        assert!(rule.equal_average);
        assert_eq!(rule.standard_deviations, Some(2));
        let default = super::conditional_from(&quick_xml::events::BytesStart::new("cfRule"), &[], &super::StyleSheet::default());
        assert!(default.above_average);
        assert!(!default.equal_average);
        assert_eq!(default.standard_deviations, None);
    }
    use super::*;

    #[test]
    fn test_parse_cell_ref_simple() {
        assert_eq!(parse_cell_ref("A1"), (0, 0));
        assert_eq!(parse_cell_ref("B2"), (1, 1));
        assert_eq!(parse_cell_ref("Z1"), (25, 0));
    }

    /// Furigana lives in `<rPh>` and is not part of the cell's text. Excel shows
    /// "陋ｹ・ｺ陋ｻ繝ｻ; appending the reading would give "陋ｹ・ｺ陋ｻ繝ｻ縺醍ｹ晄じﾎｦ" and silently corrupt
    /// most Japanese workbooks, which carry phonetic guides on names and
    /// addresses as a matter of course.
    #[test]
    fn a_wide_escape_is_spelled_out_once() {
        assert_eq!(unescape_wide("plain"), "plain");
        assert_eq!(unescape_wide("NSD_x000D_"), "NSD\r");
        // The escape for an underscore leaves what follows it alone: this is
        // how a writer spells a literal `_x000D_`.
        assert_eq!(unescape_wide("_x005F_x000D_"), "_x000D_");
        // Not four hex digits, so not an escape.
        assert_eq!(unescape_wide("_xZZZZ_"), "_xZZZZ_");
        assert_eq!(unescape_wide("a_x0041_b"), "aAb");
    }

    #[test]
    fn shared_strings_exclude_phonetic_guides() {
        let xml = r#"<?xml version="1.0"?>
<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="3">
  <si><t>&#x533A;&#x5206;</t><rPh sb="0" eb="2"><t>&#x30AF;&#x30D6;&#x30F3;</t></rPh><phoneticPr fontId="1"/></si>
  <si><r><t>&#x5C71;&#x7530;</t></r><r><t>&#x592A;&#x90CE;</t></r><rPh sb="0" eb="2"><t>&#x30E4;&#x30DE;&#x30C0;</t></rPh><rPh sb="2" eb="4"><t>&#x30BF;&#x30ED;&#x30A6;</t></rPh></si>
  <si><t>plain</t></si>
</sst>"#;
        let strings = parse_shared_strings(xml).expect("should parse");
        let held: Vec<&str> = strings.iter().map(|one| one.text.as_str()).collect();
        assert_eq!(held, vec!["\u{533a}\u{5206}", "\u{5c71}\u{7530}\u{592a}\u{90ce}", "plain"]);
    }

    #[test]
    fn cell_string_escapes_are_decoded_within_each_text_run() {
        let rich = "<r><t>_x0</t></r><r><t>041_ _x005F_x0041_</t></r><r><t>_xDE00_</t></r>";
        let strings = parse_shared_strings(&format!("<sst><si>{rich}</si></sst>")).unwrap();
        let expected: Vec<u16> = "_x0041_ _x0041_".encode_utf16().chain([0xDE00]).collect();
        assert_eq!(strings[0].units.as_ref(), Some(&expected));
        assert_eq!(resolve_cell_value("0", &Some("s".into()), &strings).string_units().unwrap().as_ref(), expected);
        let xml = format!("<worksheet><sheetData><row r=\"1\"><c r=\"A1\" t=\"inlineStr\"><is>{rich}</is></c></row></sheetData></worksheet>");
        let sheet = parse_worksheet(&xml, "probe", &[], &StyleSheet::default(), &Theme::default()).unwrap();
        assert_eq!(sheet.rows[0].cells[0].value.string_units().unwrap().as_ref(), expected);
    }

    /// A cell can hold its own text instead of pointing into the shared table.
    /// Whole sheets are written that way by anything that streams a workbook
    /// out without building the table first, and Excel shows them like any
    /// other text 遯ｶ繝ｻso a reader that only looks in <v> draws an empty sheet.
    #[test]
    fn inline_strings_are_read_from_the_cell() {
        let xml = r#"<?xml version="1.0"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData>
    <row r="1">
      <c r="A1" t="inlineStr"><is><t>Agreed</t></is></c>
      <c r="B1" t="inlineStr"><is><r><t>half </t></r><r><t>and half</t></r></is></c>
      <c r="C1" t="inlineStr"><is><t>&#x533A;&#x5206;</t><rPh sb="0" eb="2"><t>&#x30AF;&#x30D6;&#x30F3;</t></rPh></is></c>
      <c r="D1"><v>42</v></c>
    </row>
  </sheetData>
</worksheet>"#;
        let sheet = parse_worksheet(xml, "probe", &[], &StyleSheet::default(), &Theme::default())
            .expect("should parse");
        let cells = &sheet.rows[0].cells;
        let text_of = |cell: &Cell| match &cell.value {
            CellValue::String(held) => held.clone(),
            other => panic!("wanted text, found {other:?}"),
        };
        assert_eq!(text_of(&cells[0]), "Agreed");
        assert_eq!(text_of(&cells[1]), "half and half");
        // A phonetic guide is no more part of the text here than in a shared
        // string: Excel shows 陋ｹ・ｺ陋ｻ繝ｻ not 陋ｹ・ｺ陋ｻ繝ｻ縺醍ｹ晄じﾎｦ.
        assert_eq!(text_of(&cells[2]), "\u{533a}\u{5206}");
        assert!(matches!(cells[3].value, CellValue::Number(n) if n == 42.0));
    }

    #[test]
    fn complete_column_styles_are_compacted_without_losing_exceptions() {
        let styles = r#"<styleSheet><fonts count="2"><font><sz val="11"/></font><font><b/><sz val="11"/></font></fonts>
          <cellStyleXfs count="1"><xf fontId="0"/></cellStyleXfs>
          <cellXfs count="2"><xf fontId="0" xfId="0"/><xf fontId="1" xfId="0" applyFont="1"/></cellXfs>
        </styleSheet>"#;
        let stylesheet = parse_styles_xml(styles, &Theme::default()).unwrap();
        let full = r#"<worksheet><cols><col min="1" max="16384" style="1"/>
          <col min="1" max="1" style="0"/></cols><sheetData/></worksheet>"#;
        let sheet = parse_worksheet(full, "probe", &[], &stylesheet, &Theme::default()).unwrap();
        assert!(sheet.band_styles.sheet.as_ref().unwrap().bold);
        assert_eq!(sheet.band_styles.columns.len(), 1);
        assert!(!sheet.band_styles.columns.get(&0).unwrap().bold);
        let partial = r#"<worksheet><cols><col min="1" max="16383" style="1"/></cols><sheetData/></worksheet>"#;
        let sheet = parse_worksheet(partial, "probe", &[], &stylesheet, &Theme::default()).unwrap();
        assert!(sheet.band_styles.sheet.is_none());
        assert_eq!(sheet.band_styles.columns.len(), 16_383);
    }

    #[test]
    fn raw_native_font_names_survive_parse_serialize_and_reparse() {
        let cases: &[(&str, &[u16])] = &[(r####"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" mc:Ignorable="x14ac x16r2 xr" xmlns:x14ac="http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac" xmlns:x16r2="http://schemas.microsoft.com/office/spreadsheetml/2015/02/main" xmlns:xr="http://schemas.microsoft.com/office/spreadsheetml/2014/revision"><fonts count="1"><font><sz val="13"/><color theme="1"/><name val="_xd800_"/><family val="2"/></font></fonts><fills count="2"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill></fills><borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders><cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs><cellXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/></cellXfs><cellStyles count="1"><cellStyle name="標準" xfId="0" builtinId="0"/></cellStyles><dxfs count="0"/><tableStyles count="0" defaultTableStyle="TableStyleMedium2" defaultPivotStyle="PivotStyleLight16"/><extLst><ext uri="{EB79DEF2-80B8-43e5-95BD-54CBDDF9020C}" xmlns:x14="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main"><x14:slicerStyles defaultSlicerStyle="SlicerStyleLight1"/></ext><ext uri="{9260A510-F301-46a8-8635-F512D64BE5F5}" xmlns:x15="http://schemas.microsoft.com/office/spreadsheetml/2010/11/main"><x15:timelineStyles defaultTimelineStyle="TimeSlicerStyleLight1"/></ext></extLst></styleSheet>"####,&[55296]),(r####"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" mc:Ignorable="x14ac x16r2 xr" xmlns:x14ac="http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac" xmlns:x16r2="http://schemas.microsoft.com/office/spreadsheetml/2015/02/main" xmlns:xr="http://schemas.microsoft.com/office/spreadsheetml/2014/revision"><fonts count="1"><font><sz val="13"/><color theme="1"/><name val="_xdc00_"/><family val="2"/></font></fonts><fills count="2"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill></fills><borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders><cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs><cellXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/></cellXfs><cellStyles count="1"><cellStyle name="標準" xfId="0" builtinId="0"/></cellStyles><dxfs count="0"/><tableStyles count="0" defaultTableStyle="TableStyleMedium2" defaultPivotStyle="PivotStyleLight16"/><extLst><ext uri="{EB79DEF2-80B8-43e5-95BD-54CBDDF9020C}" xmlns:x14="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main"><x14:slicerStyles defaultSlicerStyle="SlicerStyleLight1"/></ext><ext uri="{9260A510-F301-46a8-8635-F512D64BE5F5}" xmlns:x15="http://schemas.microsoft.com/office/spreadsheetml/2010/11/main"><x15:timelineStyles defaultTimelineStyle="TimeSlicerStyleLight1"/></ext></extLst></styleSheet>"####,&[56320]),(r####"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" mc:Ignorable="x14ac x16r2 xr" xmlns:x14ac="http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac" xmlns:x16r2="http://schemas.microsoft.com/office/spreadsheetml/2015/02/main" xmlns:xr="http://schemas.microsoft.com/office/spreadsheetml/2014/revision"><fonts count="1"><font><sz val="13"/><color theme="1"/><name val="😀"/><family val="2"/></font></fonts><fills count="2"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill></fills><borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders><cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs><cellXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/></cellXfs><cellStyles count="1"><cellStyle name="標準" xfId="0" builtinId="0"/></cellStyles><dxfs count="0"/><tableStyles count="0" defaultTableStyle="TableStyleMedium2" defaultPivotStyle="PivotStyleLight16"/><extLst><ext uri="{EB79DEF2-80B8-43e5-95BD-54CBDDF9020C}" xmlns:x14="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main"><x14:slicerStyles defaultSlicerStyle="SlicerStyleLight1"/></ext><ext uri="{9260A510-F301-46a8-8635-F512D64BE5F5}" xmlns:x15="http://schemas.microsoft.com/office/spreadsheetml/2010/11/main"><x15:timelineStyles defaultTimelineStyle="TimeSlicerStyleLight1"/></ext></extLst></styleSheet>"####,&[55357,56832]),(r####"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" mc:Ignorable="x14ac x16r2 xr" xmlns:x14ac="http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac" xmlns:x16r2="http://schemas.microsoft.com/office/spreadsheetml/2015/02/main" xmlns:xr="http://schemas.microsoft.com/office/spreadsheetml/2014/revision"><fonts count="1"><font><sz val="13"/><color theme="1"/><name val="_x005f_x0041_"/><family val="2"/></font></fonts><fills count="2"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill></fills><borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders><cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs><cellXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/></cellXfs><cellStyles count="1"><cellStyle name="標準" xfId="0" builtinId="0"/></cellStyles><dxfs count="0"/><tableStyles count="0" defaultTableStyle="TableStyleMedium2" defaultPivotStyle="PivotStyleLight16"/><extLst><ext uri="{EB79DEF2-80B8-43e5-95BD-54CBDDF9020C}" xmlns:x14="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main"><x14:slicerStyles defaultSlicerStyle="SlicerStyleLight1"/></ext><ext uri="{9260A510-F301-46a8-8635-F512D64BE5F5}" xmlns:x15="http://schemas.microsoft.com/office/spreadsheetml/2010/11/main"><x15:timelineStyles defaultTimelineStyle="TimeSlicerStyleLight1"/></ext></extLst></styleSheet>"####,&[95,120,48,48,52,49,95])];
        for &(xml, expected) in cases {
            let styles = parse_styles_xml(xml, &Theme::default()).unwrap();
            let style = resolve_cell_style(0, &styles);
            assert_eq!(style.latin_font_units().as_deref(), Some(expected));
            let serialized = crate::editor::serialize_cell_font(&style);
            let changed = format!("<styleSheet><fonts>{serialized}</fonts><cellXfs><xf fontId=\"0\"/></cellXfs></styleSheet>");
            let reparsed = parse_styles_xml(&changed, &Theme::default()).unwrap();
            assert_eq!(resolve_cell_style(0, &reparsed).latin_font_units().as_deref(), Some(expected));
        }
    }

    #[test]
    fn workbook_default_font_scheme_matches_native_theme_versions() {
        for (version, expected) in [("202300", "Aptos Display"), ("0", "Courier New")] {
            let mut theme = Theme {major_face:Some("メイリオ".into()), major_latin_face:Some("Courier New".into()), ..Default::default()};
            let xml = format!(r#"<workbook><workbookPr defaultThemeVersion="{version}"/></workbook>"#);
            apply_workbook_default_font_scheme(&mut theme, &xml);
            assert_eq!(theme.major_latin_face.as_deref(), Some(expected));
            assert_eq!(theme.major_face.as_deref(), Some(if version=="0" {"メイリオ"} else {"游ゴシック Light"}));
        }
    }

    #[test]
    fn theme_colours_override_defaults_by_named_slot() {
        let xml = r#"<theme><themeElements><clrScheme>
          <hlink><srgbClr val="0563C1"/></hlink>
          <dk1><sysClr val="windowText" lastClr="112233"/></dk1>
          <accent2><srgbClr val="AABBCC"><alpha val="100000"/></srgbClr></accent2>
        </clrScheme></themeElements></theme>"#;
        let theme = parse_theme_xml(xml);
        assert_eq!(theme.colours.len(), 12);
        assert_eq!(theme.colour(10), Some("0563C1"));
        assert_eq!(theme.colour(1), Some("112233"));
        assert_eq!(theme.colour(5), Some("AABBCC"));
        assert_eq!(theme.colour(0), Some("FFFFFF"));
    }

    #[test]
    fn a_sheet_tab_colour_is_read_from_sheet_pr() {
        // An explicit rgb, with the leading alpha stripped like every other colour.
        let coloured = r#"<?xml version="1.0"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetPr><tabColor rgb="FF00B050"/></sheetPr>
  <sheetData/>
</worksheet>"#;
        let sheet = parse_worksheet(coloured, "probe", &[], &StyleSheet::default(), &Theme::default())
            .expect("should parse");
        assert_eq!(sheet.tab_color.as_deref(), Some("00B050"));

        // No <tabColor> leaves it None.
        let plain = r#"<?xml version="1.0"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData/>
</worksheet>"#;
        let sheet = parse_worksheet(plain, "probe", &[], &StyleSheet::default(), &Theme::default())
            .expect("should parse");
        assert_eq!(sheet.tab_color, None);

        // A theme colour is resolved against the workbook's palette.
        let mut theme = Theme::default();
        theme.colours = vec!["112233".to_string(), "445566".to_string()];
        let themed = r#"<?xml version="1.0"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetPr><tabColor theme="1"/></sheetPr>
  <sheetData/>
</worksheet>"#;
        let sheet = parse_worksheet(themed, "probe", &[], &StyleSheet::default(), &theme)
            .expect("should parse");
        assert_eq!(sheet.tab_color.as_deref(), Some("112233"));
    }

    /// Shared strings that are all dressed alike, which is what most tests want.
    fn shared(texts: &[&str]) -> Vec<SharedString> {
        texts
            .iter()
            .map(|text| SharedString {
                text: (*text).to_string(),
                units: None,
                runs: Vec::new(),
            })
            .collect()
    }

    #[test]
    fn test_parse_cell_ref_multi_letter() {
        assert_eq!(parse_cell_ref("AA1"), (26, 0));
        assert_eq!(parse_cell_ref("AB1"), (27, 0));
        assert_eq!(parse_cell_ref("AZ3"), (51, 2));
    }

    #[test]
    fn test_parse_cell_ref_large_row() {
        assert_eq!(parse_cell_ref("A100"), (0, 99));
        assert_eq!(parse_cell_ref("C65536"), (2, 65535));
    }

    #[test]
    fn test_resolve_cell_value_number() {
        let sst: Vec<SharedString> = vec![];
        assert!(matches!(
            resolve_cell_value("42", &None, &sst),
            CellValue::Number(n) if (n - 42.0).abs() < f64::EPSILON
        ));
    }

    #[test]
    fn test_resolve_cell_value_shared_string() {
        let sst = shared(&["Hello", "World"]);
        let t = Some("s".to_string());
        assert!(matches!(
            resolve_cell_value("0", &t, &sst),
            CellValue::String(ref s) if s == "Hello"
        ));
        assert!(matches!(
            resolve_cell_value("1", &t, &sst),
            CellValue::String(ref s) if s == "World"
        ));
    }

    #[test]
    fn test_resolve_cell_value_boolean() {
        let sst: Vec<SharedString> = vec![];
        let t = Some("b".to_string());
        assert!(matches!(
            resolve_cell_value("1", &t, &sst),
            CellValue::Boolean(true)
        ));
        assert!(matches!(
            resolve_cell_value("0", &t, &sst),
            CellValue::Boolean(false)
        ));
    }

    #[test]
    fn test_resolve_cell_value_error() {
        let sst: Vec<SharedString> = vec![];
        let t = Some("e".to_string());
        assert!(matches!(
            resolve_cell_value("#REF!", &t, &sst),
            CellValue::Error(ref s) if s == "#REF!"
        ));
    }

    #[test]
    fn test_resolve_cell_value_empty() {
        let sst: Vec<SharedString> = vec![];
        assert!(matches!(
            resolve_cell_value("", &None, &sst),
            CellValue::Empty
        ));
    }

    #[test]
    fn test_cell_value_display() {
        assert_eq!(CellValue::Empty.display(), "");
        assert_eq!(CellValue::String("hello".into()).display(), "hello");
        assert_eq!(CellValue::Number(42.0).display(), "42");
        assert_eq!(CellValue::Number(2.75).display(), "2.75");
        assert_eq!(CellValue::Boolean(true).display(), "TRUE");
        assert_eq!(CellValue::Boolean(false).display(), "FALSE");
        assert_eq!(CellValue::Error("#N/A".into()).display(), "#N/A");
    }

    #[test]
    fn test_parse_shared_strings() {
        let xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="3" uniqueCount="3">
  <si><t>Hello</t></si>
  <si><t>World</t></si>
  <si><r><t>Rich</t></r><r><t> Text</t></r></si>
</sst>"#;
        let result = parse_shared_strings(xml).unwrap();
        assert_eq!(result.len(), 3);
        assert_eq!(result[0].text, "Hello");
        assert_eq!(result[1].text, "World");
        assert_eq!(result[2].text, "Rich Text");
    }

    #[test]
    fn test_parse_workbook_sheets() {
        let xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
          xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheets>
    <sheet name="Sheet1" sheetId="1" r:id="rId1"/>
    <sheet name="Data" sheetId="2" r:id="rId2"/>
  </sheets>
</workbook>"#;
        let result = parse_workbook_sheets(xml).unwrap();
        assert_eq!(result.len(), 2);
        assert_eq!(result[0].name, "Sheet1");
        assert_eq!(result[0].r_id, "rId1");
        assert_eq!(result[1].name, "Data");
        assert_eq!(result[1].r_id, "rId2");
    }

    #[test]
    fn test_parse_range_ref() {
        assert_eq!(parse_range_ref("A1:C3"), Some((0, 1, 2, 3)));
        assert_eq!(parse_range_ref("B2:D5"), Some((1, 2, 3, 5)));
        assert_eq!(parse_range_ref("A1"), None);
    }

    /// Cell font selection and saved application flags remain independent.
    #[test]
    fn a_cell_font_identifier_is_used_even_when_application_is_disabled() {
        let xml = r##"<?xml version="1.0"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <fonts count="3">
    <font><sz val="11"/><name val="Calibri"/></font>
    <font><sz val="10"/><name val="ＭＳ ゴシック"/></font>
    <font><u/><sz val="11"/><color rgb="FF0563C1"/><name val="Calibri"/></font>
  </fonts>
  <cellStyleXfs count="2">
    <xf numFmtId="0" fontId="0" fillId="0" borderId="0"/>
    <xf numFmtId="0" fontId="2" fillId="0" borderId="0"/>
  </cellStyleXfs>
  <cellXfs count="3">
    <xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>
    <xf numFmtId="0" fontId="1" fillId="0" borderId="0" xfId="0"/>
    <xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="1" applyFont="0"/>
  </cellXfs>
</styleSheet>"##;
        let sheet = parse_styles_xml(xml, &Theme::default()).unwrap();
        let named = resolve_cell_style(1, &sheet);
        assert_eq!(named.font_name.as_deref(), Some("ＭＳ ゴシック"));
        assert_eq!(named.font_size, Some(10.0));
        let denied = resolve_cell_style(2, &sheet);
        assert!(!denied.underline);
        assert_eq!(denied.font_name.as_deref(), Some("Calibri"));
        assert_eq!(denied.font_color, None);
        let reference = denied.format_reference.as_ref().unwrap();
        assert_eq!(reference.style_id, Some(1));
        assert_eq!(reference.font, Some(true));
    }

    #[test]
    fn empty_style_components_keep_their_indices() {
        let xml = r##"<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
<fonts count="2"><font/><font><name val="Consolas"/><sz val="12"/></font></fonts>
<fills count="2"><fill/><fill><patternFill patternType="solid"><fgColor rgb="FFFF0000"/></patternFill></fill></fills>
<borders count="2"><border/><border><left style="thin"><color rgb="FF0000FF"/></left></border></borders>
<cellStyleXfs count="1"><xf fontId="0" fillId="0" borderId="0" numFmtId="0"/></cellStyleXfs>
<cellXfs count="1"><xf fontId="1" fillId="1" borderId="1" numFmtId="0" xfId="0"/></cellXfs>
</styleSheet>"##;
        let sheet = parse_styles_xml(xml, &Theme::default()).unwrap();
        assert_eq!(sheet.fonts.len(), 2);
        assert_eq!(sheet.fills.len(), 2);
        assert_eq!(sheet.borders.len(), 2);
        let cell = resolve_cell_style(0, &sheet);
        assert_eq!(cell.font_name.as_deref(), Some("Consolas"));
        assert_eq!(cell.font_size, Some(12.0));
        assert_eq!(cell.bg_color.as_deref(), Some("FF0000"));
        assert_eq!(cell.border_left.as_ref().map(|line| line.style.as_str()), Some("thin"));
        assert_eq!(cell.border_left.as_ref().and_then(|line| line.color.as_deref()), Some("0000FF"));
    }

    #[test]
    fn imported_application_flags_compare_complete_components() {
        let flags = ["applyFont", "applyFill", "applyBorder", "applyNumberFormat", "applyAlignment", "applyProtection"];
        for normal in [false, true] {
            for raw_flag in [None, Some("0"), Some("1")] {
                for kind in 0..8 {
                    let changed = |component| kind == component || kind == 7;
                    let applications = raw_flag.map(|value| flags.iter()
                        .map(|name| format!(r#"{name}="{value}""#)).collect::<Vec<_>>().join(" ")).unwrap_or_default();
                    let parent_id = usize::from(!normal);
                    let parent = format!(r#"<xf fontId="{}" fillId="{}" borderId="{}" numFmtId="{}">{}{}</xf>"#,
                        usize::from(changed(1)), usize::from(changed(2)), usize::from(changed(3)),
                        if changed(4) {2} else {0},
                        if changed(5) {r#"<alignment horizontal="right" vertical="top" wrapText="1" indent="2"/>"#} else {""},
                        if changed(6) {r#"<protection locked="0" hidden="1"/>"#} else {""});
                    let parents = if normal { parent } else {
                        format!(r#"<xf fontId="0" fillId="0" borderId="0" numFmtId="0"/>{parent}"#)
                    };
                    let xml = format!(r#"<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
<fonts><font><name val="Calibri"/><sz val="11"/></font><font><name val="Arial"/><sz val="15"/></font></fonts>
<fills><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="solid"><fgColor rgb="FF00FF00"/></patternFill></fill></fills>
<borders><border/><border><left style="double"><color rgb="FFFF0000"/></left></border></borders>
<cellStyleXfs>{parents}</cellStyleXfs>
<cellXfs><xf fontId="0" fillId="0" borderId="0" numFmtId="0" xfId="0"/>
<xf fontId="0" fillId="0" borderId="0" numFmtId="0" xfId="{parent_id}" {applications}/></cellXfs></styleSheet>"#);
                    let imported = import_styles_xml(&xml).unwrap();
                    assert_eq!(import_styles_xml(&imported).unwrap(), imported);
                    let sheet = parse_styles_xml(&xml, &Theme::default()).unwrap();
                    let reference = &sheet.cell_xfs[1].format_reference;
                    let actual = [reference.font, reference.fill, reference.border, reference.number_format, reference.alignment, reference.protection];
                    for (index, actual) in actual.into_iter().enumerate() {
                        let expected = (changed(index + 1) && !(normal && index == 0)).then_some(true);
                        assert_eq!(actual, expected, "normal={normal}, raw={raw_flag:?}, kind={kind}, flag={}", flags[index]);
                    }
                }
            }
        }
        // These properties are not all present in CellStyle. Import must
        // compare their XML rather than silently dropping their differences.
        let xml = r#"<styleSheet><fonts><font/></fonts><fills><fill/></fills><borders><border/></borders>
<cellStyleXfs><xf/></cellStyleXfs><cellXfs><xf/><xf><alignment textRotation="45" readingOrder="2"/><protection hidden="1"/></xf></cellXfs></styleSheet>"#;
        let styles = parse_styles_xml(xml, &Theme::default()).unwrap();
        assert_eq!(styles.cell_xfs[1].format_reference.alignment, Some(true));
        assert_eq!(styles.cell_xfs[1].format_reference.protection, Some(true));
    }

    #[test]
    fn cell_components_use_their_own_identifiers_across_application_flags() {
        let flags = ["applyFont", "applyFill", "applyBorder", "applyNumberFormat", "applyAlignment", "applyProtection"];
        let identifiers = ["fontId", "fillId", "borderId", "numFmtId"];
        for missing in [false, true] {
            for kind in 0..8 {
                let omitted = missing.then(|| identifiers[kind / 2]);
                let ids = identifiers.iter().filter(|name| Some(**name) != omitted)
                    .map(|name| format!(r#"{name}="0""#)).collect::<Vec<_>>().join(" ");
                let applications = flags.iter().enumerate().map(|(index, name)| {
                    let enabled = if missing { index != kind / 2 || kind % 2 == 1 } else { kind != 6 && kind != index };
                    format!(r#"{name}="{}""#, u8::from(enabled))
                }).collect::<Vec<_>>().join(" ");
                let xml = format!(r##"<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
<fonts count="2"><font><name val="Calibri"/><sz val="11"/></font><font><name val="Arial"/><sz val="15"/><b/><color rgb="FF0000FF"/></font></fonts>
<fills count="2"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="solid"><fgColor rgb="FF00FF00"/></patternFill></fill></fills>
<borders count="2"><border/><border><left style="double"><color rgb="FFFF0000"/></left></border></borders>
<cellStyleXfs count="2"><xf fontId="0" fillId="0" borderId="0" numFmtId="0"/><xf fontId="1" fillId="1" borderId="1" numFmtId="2"><alignment horizontal="right" vertical="top" wrapText="1" indent="2"/><protection locked="0" hidden="1"/></xf></cellStyleXfs>
<cellXfs count="1"><xf {ids} xfId="1" {applications}><alignment horizontal="left" vertical="bottom" wrapText="0"/><protection locked="1" hidden="0"/></xf></cellXfs>
<cellStyles count="2"><cellStyle name="Normal" xfId="0" builtinId="0"/><cellStyle name="ProbeInheritance" xfId="1"/></cellStyles>
</styleSheet>"##);
                let styles = parse_styles_xml(&xml, &Theme::default()).unwrap();
                let actual = resolve_cell_style(0, &styles);
                assert_eq!(actual.font_name.as_deref(), Some("Calibri"), "missing={missing}, kind={kind}");
                assert_eq!(actual.font_size, Some(11.0));
                assert!(!actual.bold);
                assert_eq!(actual.font_color, None);
                assert_eq!(actual.bg_color, None);
                assert_eq!(actual.border_left, None);
                assert_eq!(actual.number_format, None);
                assert_eq!(actual.horizontal_align.as_deref(), Some("left"));
                assert_eq!(actual.vertical_align.as_deref(), Some("bottom"));
                assert!(!actual.wrap_text);
                assert!(!actual.unlocked);
                let reference = actual.format_reference.as_ref().unwrap();
                assert_eq!(reference.style_id, Some(1));
                let stored = [reference.font, reference.fill, reference.border, reference.number_format, reference.alignment, reference.protection];
                assert_eq!(stored, [Some(true); 6]);
            }
        }
    }

    /// A face the machine has not got is answered from the two numbers beside
    /// the name, not from the name: `sanko_tool` asks for `AR P闕ｳ・ｸ郢ｧ・ｴ郢ｧ・ｷ郢昴・縺題抄鞫倭
    /// with `family 3` (fixed pitch) and `charset 128` (Japanese), and Excel
    /// draws 繝ｻ・ｭ繝ｻ・ｳ 郢ｧ・ｴ郢ｧ・ｷ郢昴・縺・ Drop either number in the reader and the renderer
    /// has nothing to answer with.
    #[test]
    fn a_font_keeps_the_family_and_charset_beside_its_name() {
        let xml = r##"<?xml version="1.0"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <fonts count="2">
    <font><sz val="11"/><name val="Calibri"/></font>
    <font><sz val="11"/><color rgb="FFFF0000"/><name val="AR P丸ゴシック体E"/><family val="3"/><charset val="128"/></font>
  </fonts>
  <cellStyleXfs count="1">
    <xf numFmtId="0" fontId="0" fillId="0" borderId="0"/>
  </cellStyleXfs>
  <cellXfs count="2">
    <xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>
    <xf numFmtId="0" fontId="1" fillId="0" borderId="0" xfId="0" applyFont="1"/>
  </cellXfs>
</styleSheet>"##;
        let sheet = parse_styles_xml(xml, &Theme::default()).unwrap();
        let asked = resolve_cell_style(1, &sheet);
        assert_eq!(asked.font_name.as_deref(), Some("AR P丸ゴシック体E"));
        assert_eq!(asked.font_family, Some(3));
        assert_eq!(asked.font_charset, Some(128));
        let plain = resolve_cell_style(0, &sheet);
        assert_eq!(plain.font_family, None);
        assert_eq!(plain.font_charset, None);
    }

    /// A generated header row is written as `<font><b/></font>` 遯ｶ繝ｻbold, and
    /// nothing else. The face and size it leaves out are the workbook's, not
    /// the reader's idea of a default: read otherwise, every such row is drawn
    /// in the wrong face and stands the wrong height.
    #[test]
    fn a_font_that_names_no_face_wears_the_workbook_s() {
        let xml = r##"<?xml version="1.0"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <fonts count="2">
    <font><name val="ＭＳ Ｐゴシック"/><sz val="9"/></font>
    <font><b val="1"/><color rgb="00FFFFFF"/></font>
  </fonts>
  <cellStyleXfs count="1">
    <xf numFmtId="0" fontId="0" fillId="0" borderId="0"/>
  </cellStyleXfs>
  <cellXfs count="2">
    <xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>
    <xf numFmtId="0" fontId="1" fillId="0" borderId="0" xfId="0"/>
  </cellXfs>
</styleSheet>"##;
        let sheet = parse_styles_xml(xml, &Theme::default()).unwrap();
        let header = resolve_cell_style(1, &sheet);
        assert!(header.bold);
        assert_eq!(header.font_name.as_deref(), Some("ＭＳ Ｐゴシック"));
        assert_eq!(header.font_size, Some(9.0));
    }

    /// A drawing part states where a shape hangs, what it is painted with,
    /// and 遯ｶ繝ｻfor a picture 遯ｶ繝ｻwhich part holds its bytes.
    #[test]
    fn a_drawing_part_gives_up_its_anchors_and_its_paint() {
        let xml = r##"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing"
          xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
          xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <xdr:twoCellAnchor>
    <xdr:from><xdr:col>5</xdr:col><xdr:colOff>12699</xdr:colOff>
      <xdr:row>8</xdr:row><xdr:rowOff>85725</xdr:rowOff></xdr:from>
    <xdr:to><xdr:col>7</xdr:col><xdr:colOff>590550</xdr:colOff>
      <xdr:row>9</xdr:row><xdr:rowOff>123825</xdr:rowOff></xdr:to>
    <xdr:sp macro="" textlink="">
      <xdr:nvSpPr><xdr:cNvPr id="77" name="Imported"/><xdr:cNvSpPr/></xdr:nvSpPr>
      <xdr:spPr>
        <a:prstGeom prst="rect"><a:avLst/></a:prstGeom>
        <a:noFill/>
        <a:ln w="28575"><a:headEnd type="diamond" w="lg" len="sm"/><a:tailEnd type="oval" w="sm" len="lg"/><a:solidFill><a:srgbClr val="FF0000"/></a:solidFill></a:ln>
      </xdr:spPr>
    </xdr:sp>
    <xdr:clientData/>
  </xdr:twoCellAnchor>
  <xdr:oneCellAnchor>
    <xdr:from><xdr:col>0</xdr:col><xdr:colOff>0</xdr:colOff>
      <xdr:row>0</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from>
    <xdr:ext cx="952500" cy="476250"/>
    <xdr:pic>
      <xdr:blipFill><a:blip r:embed="rId7"/></xdr:blipFill>
      <xdr:spPr><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></xdr:spPr>
    </xdr:pic>
    <xdr:clientData/>
  </xdr:oneCellAnchor>
</xdr:wsDr>"##;
        let found = parse_drawing_xml(xml, &Theme::default());
        assert_eq!(found[0].0.non_visual_id, Some(77));
        assert_eq!(found.len(), 2);

        let (rectangle, embeds) = &found[0];
        assert_eq!(embeds.raster, None);
        assert_eq!(embeds.vector, None);
        assert_eq!(rectangle.from.col, 5);
        assert_eq!(rectangle.from.col_off, 12699);
        assert_eq!(rectangle.to.map(|to| (to.col, to.row)), Some((7, 9)));
        let crate::ir::DrawingKind::Shape(shape) = &rectangle.kind else {
            panic!("the first anchor holds a shape");
        };
        assert_eq!(shape.geometry, "rect");
        assert_eq!(shape.fill, None, "a:noFill leaves the shape unpainted");
        let line = shape.line.as_ref().expect("the shape is ruled");
        assert_eq!(line.color, "FF0000");
        assert_eq!(line.width, 28575, "three pixels at 96 dpi");
        assert_eq!(line.dash, None);
        assert_eq!(line.head_end.as_deref(), Some("diamond"));
        assert_eq!(line.head_end_width.as_deref(), Some("lg"));
        assert_eq!(line.head_end_length.as_deref(), Some("sm"));
        assert_eq!(line.tail_end.as_deref(), Some("oval"));
        assert_eq!(line.tail_end_width.as_deref(), Some("sm"));
        assert_eq!(line.tail_end_length.as_deref(), Some("lg"));

        let (picture, embeds) = &found[1];
        assert_eq!(embeds.raster.as_deref(), Some("rId7"));
        assert_eq!(embeds.vector, None);
        assert_eq!(picture.extent, Some((952500, 476250)));
        assert!(matches!(picture.kind, crate::ir::DrawingKind::Picture { .. }));
    }

    #[test]
    fn captures_svg_picture_relationship_and_intrinsic_extent() {
        let xml = r#"<xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:asvg="http://schemas.microsoft.com/office/drawing/2016/SVG/main"><xdr:oneCellAnchor><xdr:from><xdr:col>0</xdr:col><xdr:row>0</xdr:row></xdr:from><xdr:ext cx="1" cy="1"/><xdr:pic><xdr:blipFill><a:blip r:embed="rId1"><a:extLst><a:ext><asvg:svgBlip r:embed="rId2"/></a:ext></a:extLst></a:blip></xdr:blipFill></xdr:pic><xdr:clientData/></xdr:oneCellAnchor></xdr:wsDr>"#;
        let found = parse_drawing_xml(xml, &Theme::default());
        assert_eq!(found.len(), 1);
        assert_eq!(found[0].1.raster.as_deref(), Some("rId1"));
        assert_eq!(found[0].1.vector.as_deref(), Some("rId2"));
        assert_eq!(svg_intrinsic_size_points(br#"<svg width="120" height="80" viewBox="0 0 120 80"/>"#), Some((90.0, 60.0)));
        assert_eq!(svg_intrinsic_size_points(br#"<svg width="1in" viewBox="0 0 2 1"/>"#), Some((72.0, 36.0)));
    }

    #[test]
    fn excel_svg_picture_uses_vector_size_over_raster_preview() {
        let book = crate::parse_xlsx(include_bytes!("../../../tests/fixtures/vba/scale_picture_original_svg.xlsx")).unwrap();
        let format = book.sheets[0].drawings.iter().find_map(|drawing| match &drawing.kind {
            crate::ir::DrawingKind::Picture { format, .. } => Some(format),
            _ => None,
        }).expect("picture format");
        assert_eq!(format.original_size_points, Some((90.0, 60.0)));
    }

    #[test]
    fn excel_ole_objects_keep_their_drawing_frames() {
        let book = crate::parse_xlsx(include_bytes!("../../../tests/fixtures/vba/scale_picture_original_ole.xlsx")).unwrap();
        assert_eq!(book.sheets[0].drawings.len(), 2);
        assert_eq!(book.sheets[0].drawings[0].name.as_deref(), Some("Object 1"));
        assert_eq!(book.sheets[0].drawings[1].name.as_deref(), Some("Object 2"));
        assert!(book.sheets[0].drawings.iter().all(|drawing| matches!(drawing.kind, crate::ir::DrawingKind::Shape(_))));
        assert_eq!(book.sheets[0].drawings[0].ole_prog_id.as_deref(), Some("Package"));
        assert_eq!(book.sheets[0].drawings[1].ole_prog_id.as_deref(), Some("Word.Document.12"));
        assert_eq!(book.sheets[0].drawings[0].ole_placement, Some(2));
        assert_eq!(book.sheets[0].drawings[1].ole_placement, Some(2));
        assert_eq!(book.sheets[0].drawings[0].ole_type, Some(1));
        assert_eq!(book.sheets[0].drawings[0].ole_auto_load, None);
        assert_eq!(book.sheets[0].drawings[0].ole_enabled, None);
        assert_eq!(book.sheets[0].drawings[0].ole_locked, None);
        assert_eq!(book.sheets[0].drawings[0].ole_print_object, None);
    }

    #[test]
    fn ole_object_placement_reads_move_and_size_anchor_flags() {
        let xml = r#"<worksheet><oleObjects>
          <oleObject shapeId="1" id="rId1" autoLoad="1" progId="A"><objectPr locked="0" print="0" disabled="1"><anchor moveWithCells="1" sizeWithCells="1"/></objectPr></oleObject>
          <mc:AlternateContent><mc:Choice><oleObject shapeId="2" id="rId2" progId="B"><objectPr><anchor moveWithCells="1"/></objectPr></oleObject></mc:Choice><mc:Fallback><oleObject shapeId="2" id="rId2" progId="B"/></mc:Fallback></mc:AlternateContent>
          <oleObject shapeId="3" progId="C"><objectPr><anchor/></objectPr></oleObject>
        </oleObjects></worksheet>"#;
        let relationships = HashMap::from([
            ("rId1".into(), Relationship { id: "rId1".into(), rel_type: String::new(), target: "embeddings/a.bin".into(), target_mode: None }),
            ("rId2".into(), Relationship { id: "rId2".into(), rel_type: String::new(), target: "C:/linked.bin".into(), target_mode: Some("External".into()) }),
        ]);
        let metadata = parse_ole_object_metadata(xml, &relationships);
        assert_eq!(metadata.get(&1).map(|item| item.placement), Some(1));
        assert_eq!(metadata.get(&2).map(|item| item.placement), Some(2));
        assert_eq!(metadata.get(&3).map(|item| item.placement), Some(3));
        assert_eq!(metadata.get(&1).map(|item| item.ole_type), Some(1));
        assert_eq!(metadata.get(&2).map(|item| item.ole_type), Some(2));
        assert_eq!(metadata.get(&1).and_then(|item| item.auto_load), Some(true));
        assert_eq!(metadata.get(&1).and_then(|item| item.enabled), Some(false));
        assert_eq!(metadata.get(&1).and_then(|item| item.locked), Some(false));
        assert_eq!(metadata.get(&1).and_then(|item| item.print_object), Some(false));
    }

    /// A shape's own text, with the dressing of the run that starts each
    /// paragraph and the insets its body states.
    #[test]
    fn a_shape_gives_up_what_it_says() {
        let xml = r##"<xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing"
          xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
  <xdr:twoCellAnchor>
    <xdr:from><xdr:col>0</xdr:col><xdr:colOff>0</xdr:colOff>
      <xdr:row>0</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from>
    <xdr:to><xdr:col>8</xdr:col><xdr:colOff>0</xdr:colOff>
      <xdr:row>2</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:to>
    <xdr:sp>
      <xdr:spPr><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></xdr:spPr>
      <xdr:txBody>
        <a:bodyPr lIns="288000" tIns="72000" rIns="288000" bIns="72000" anchor="ctr"/>
        <a:p>
          <a:pPr algn="ctr"><a:lnSpc><a:spcPts val="3000"/></a:lnSpc></a:pPr>
          <a:r>
            <a:rPr sz="2000" b="1">
              <a:solidFill><a:srgbClr val="203864"/></a:solidFill>
              <a:latin typeface="Calibri"/><a:ea typeface="メイリオ"/>
            </a:rPr>
            <a:t>この計算書は、</a:t>
          </a:r>
          <a:r><a:rPr sz="2000" b="1"/><a:t>所得を計算するものです。</a:t></a:r>
        </a:p>
        <a:p><a:endParaRPr sz="2000"/></a:p>
      </xdr:txBody>
    </xdr:sp>
    <xdr:clientData/>
  </xdr:twoCellAnchor>
</xdr:wsDr>"##;
        let found = parse_drawing_xml(xml, &Theme::default());
        let crate::ir::DrawingKind::Shape(shape) = &found[0].0.kind else {
            panic!("the anchor holds a shape");
        };
        let said = shape.text.as_ref().expect("the shape says something");
        assert_eq!(said.anchor.as_deref(), Some("ctr"));
        assert_eq!(said.insets, (288000, 72000, 288000, 72000));
        // The trailing empty paragraph is a line, not a marker: Excel moves
        // the ink of a middle-anchored block half a line up and a
        // bottom-anchored one a whole line up when it is there
        // (`_xlsx_shape_block.py`). It states its own size in `endParaRPr`.
        assert_eq!(said.paragraphs.len(), 2);
        assert_eq!(said.paragraphs[1].text, "");
        assert_eq!(said.paragraphs[1].size, 20.0);
        let first = &said.paragraphs[0];
        assert_eq!(first.text, "この計算書は、所得を計算するものです。");
        assert_eq!(first.align.as_deref(), Some("ctr"));
        assert_eq!(first.size, 20.0);
        assert!(first.bold);
        assert_eq!(first.face.as_deref(), Some("メイリオ"), "the East Asian face wins");
        assert_eq!(first.color.as_deref(), Some("203864"));
        assert_eq!(first.line_pitch, Some(30.0));
    }

    /// 162 of the corpus's workbooks name a colour by number rather than by
    /// value, and a hundred of them state a palette of their own.
    #[test]
    fn a_colour_named_by_number_comes_from_the_palette() {
        let xml = r##"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <fonts count="3">
    <font><sz val="11"/><name val="Calibri"/></font>
    <font><u/><sz val="11"/><color indexed="12"/><name val="Calibri"/></font>
    <font><sz val="11"/><color indexed="64"/><name val="Calibri"/></font>
  </fonts>
  <cellStyleXfs count="1">
    <xf numFmtId="0" fontId="0" fillId="0" borderId="0"/>
  </cellStyleXfs>
  <cellXfs count="3">
    <xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>
    <xf numFmtId="0" fontId="1" fillId="0" borderId="0" xfId="0" applyFont="1"/>
    <xf numFmtId="0" fontId="2" fillId="0" borderId="0" xfId="0" applyFont="1"/>
  </cellXfs>
</styleSheet>"##;
        let sheet = parse_styles_xml(xml, &Theme::default()).unwrap();
        // 12 is the blue Excel has kept since it had only 56 colours to keep.
        assert_eq!(resolve_cell_style(1, &sheet).font_color.as_deref(), Some("0000FF"));
        // 64 is the system's own foreground, which is not in the palette at
        // all: the caller decides, as it did before any of this was read.
        assert_eq!(resolve_cell_style(2, &sheet).font_color, None);

        // A workbook that states its own palette is taken at its word.
        let own = xml.replace(
            "</styleSheet>",
            "<colors><indexedColors><rgbColor rgb=\"00000000\"/><rgbColor rgb=\"00FFFFFF\"/><rgbColor rgb=\"00FF0000\"/><rgbColor rgb=\"0000FF00\"/><rgbColor rgb=\"000000FF\"/><rgbColor rgb=\"00FFFF00\"/><rgbColor rgb=\"00FF00FF\"/><rgbColor rgb=\"0000FFFF\"/><rgbColor rgb=\"00000000\"/><rgbColor rgb=\"00FFFFFF\"/><rgbColor rgb=\"00FF0000\"/><rgbColor rgb=\"0000FF00\"/><rgbColor rgb=\"00123456\"/></indexedColors></colors></styleSheet>",
        );
        let mut theme = Theme::default();
        theme.indexed = indexed_palette(&own);
        assert_eq!(theme.indexed.len(), 13);
        let sheet = parse_styles_xml(&own, &theme).unwrap();
        assert_eq!(resolve_cell_style(1, &sheet).font_color.as_deref(), Some("123456"));
    }

    #[test]
    fn workbook_ir_retains_a_custom_indexed_color_palette() {
        use std::io::Write;
        let palette=(0..64).map(|index| {
            let rgb=if index==12 {"0000D4".to_string()} else {format!("{index:06X}")};
            format!(r#"<rgbColor rgb="FF{rgb}"/>"#)
        }).collect::<String>();
        let styles=format!(r#"<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
<fonts count="1"><font><sz val="11"/><color indexed="12"/><name val="Calibri"/></font></fonts>
<fills count="2"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill></fills>
<borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders>
<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>
<cellXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/></cellXfs>
<cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles>
<colors><indexedColors>{palette}</indexedColors></colors></styleSheet>"#);
        let parts=[
            ("[Content_Types].xml",r#"<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/><Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/></Types>"#.to_string()),
            ("_rels/.rels",r#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>"#.to_string()),
            ("xl/workbook.xml",r#"<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="Sheet1" sheetId="1" r:id="rId1"/></sheets></workbook>"#.to_string()),
            ("xl/_rels/workbook.xml.rels",r#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/></Relationships>"#.to_string()),
            ("xl/worksheets/sheet1.xml",r#"<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData><row r="1"><c r="A1"><v>1</v></c></row></sheetData></worksheet>"#.to_string()),
            ("xl/styles.xml",styles),
        ];
        let mut zip=zip::ZipWriter::new(std::io::Cursor::new(Vec::new()));
        for (name,xml) in parts {
            zip.start_file(name,zip::write::SimpleFileOptions::default()).unwrap();
            zip.write_all(xml.as_bytes()).unwrap();
        }
        let workbook=parse_xlsx(&zip.finish().unwrap().into_inner()).unwrap();
        assert_eq!(workbook.indexed_colors.len(),64);
        assert_eq!(workbook.indexed_colors[12],"0000D4");
        assert_eq!(workbook.indexed_colors[63],"00003F");
        assert_eq!(workbook.default_style.font_color.as_deref(),Some("0000D4"));
        assert_eq!(workbook.default_style.font_reference.as_ref().and_then(|r|r.indexed),Some(12));
    }

    /// `<a:br/>` starts a line without starting a paragraph. Without it
    /// `sanko_tool`'s heading ran into the sentence under it 遯ｶ繝ｻone line where
    /// Excel draws two.
    #[test]
    fn a_break_inside_a_paragraph_starts_a_line() {
        let xml = r##"<xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing"
          xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
  <xdr:twoCellAnchor>
    <xdr:from><xdr:col>0</xdr:col><xdr:colOff>0</xdr:colOff>
      <xdr:row>0</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from>
    <xdr:to><xdr:col>8</xdr:col><xdr:colOff>0</xdr:colOff>
      <xdr:row>2</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:to>
    <xdr:sp>
      <xdr:spPr><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></xdr:spPr>
      <xdr:txBody>
        <a:bodyPr/>
        <a:p>
          <a:r><a:rPr sz="1200" u="sng"/><a:t>郢ｧ・､郢晢ｽ｡郢晢ｽｼ郢ｧ・ｸ</a:t></a:r>
          <a:br><a:rPr sz="1200"/></a:br>
          <a:br><a:rPr sz="1200"/></a:br>
          <a:r><a:rPr sz="1200"/><a:t>繝ｻ謇假ｽｼ螳茨ｽ｢・ｺ髫ｱ髦ｪ・邵ｺ貅假ｼ櫁惓竏ｫ蟯ｼ</a:t></a:r>
        </a:p>
        <a:p><a:r><a:rPr sz="1200"/><a:t>繝ｻ謚ｵ・ｼ蠑ｱﾎ懃ｹ晢ｽｳ郢ｧ・ｯ闖ｫ繧育・</a:t></a:r></a:p>
      </xdr:txBody>
    </xdr:sp>
    <xdr:clientData/>
  </xdr:twoCellAnchor>
</xdr:wsDr>"##;
        let found = parse_drawing_xml(xml, &Theme::default());
        let crate::ir::DrawingKind::Shape(shape) = &found[0].0.kind else {
            panic!("the anchor holds a shape");
        };
        let said = shape.text.as_ref().expect("the shape says something");
        assert_eq!(said.paragraphs.len(), 2);
        assert_eq!(said.paragraphs[0].text, "郢ｧ・､郢晢ｽ｡郢晢ｽｼ郢ｧ・ｸ\n\n繝ｻ謇假ｽｼ螳茨ｽ｢・ｺ髫ｱ髦ｪ・邵ｺ貅假ｼ櫁惓竏ｫ蟯ｼ");
        assert_eq!(said.paragraphs[1].text, "繝ｻ謚ｵ・ｼ蠑ｱﾎ懃ｹ晢ｽｳ郢ｧ・ｯ闖ｫ繧育・");
    }

    /// A rule states its colour before it says how it is broken, so the dash
    /// has to be put on the line at the end of `<a:ln>` rather than when the
    /// colour is read 遯ｶ繝ｻ`glossary_05`'s dashed frame came out solid until it
    /// was.
    #[test]
    fn a_dash_stated_after_the_colour_still_reaches_the_line() {
        let xml = r##"<xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing"
          xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
  <xdr:twoCellAnchor>
    <xdr:from><xdr:col>0</xdr:col><xdr:colOff>0</xdr:colOff>
      <xdr:row>0</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from>
    <xdr:to><xdr:col>4</xdr:col><xdr:colOff>0</xdr:colOff>
      <xdr:row>4</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:to>
    <xdr:sp>
      <xdr:spPr>
        <a:prstGeom prst="rect"><a:avLst/></a:prstGeom>
        <a:noFill/>
        <a:ln w="19050">
          <a:solidFill><a:srgbClr val="000000"/></a:solidFill>
          <a:prstDash val="dash"/>
        </a:ln>
      </xdr:spPr>
    </xdr:sp>
    <xdr:clientData/>
  </xdr:twoCellAnchor>
</xdr:wsDr>"##;
        let found = parse_drawing_xml(xml, &Theme::default());
        let crate::ir::DrawingKind::Shape(shape) = &found[0].0.kind else {
            panic!("the anchor holds a shape");
        };
        let line = shape.line.as_ref().expect("the shape is ruled");
        assert_eq!(line.dash.as_deref(), Some("dash"));
        assert_eq!(line.width, 19050);
    }

    #[test]
    fn a_patterned_shape_line_keeps_both_colours() {
        let xml = r##"<xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing"
          xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
  <xdr:twoCellAnchor>
    <xdr:from><xdr:col>0</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>0</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from>
    <xdr:to><xdr:col>2</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>2</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:to>
    <xdr:sp><xdr:nvSpPr><xdr:cNvPr id="2" name="Patterned"/><xdr:cNvSpPr/></xdr:nvSpPr>
      <xdr:spPr><a:prstGeom prst="rect"><a:avLst/></a:prstGeom>
        <a:ln><a:pattFill prst="pct10"><a:fgClr><a:srgbClr val="FF0000"/></a:fgClr>
          <a:bgClr><a:srgbClr val="00FF00"/></a:bgClr></a:pattFill></a:ln>
      </xdr:spPr>
    </xdr:sp><xdr:clientData/>
  </xdr:twoCellAnchor>
</xdr:wsDr>"##;
        let found = parse_drawing_xml(xml, &Theme::default());
        let drawing = &found[0].0;
        assert_eq!(drawing.name.as_deref(), Some("Patterned"));
        let crate::ir::DrawingKind::Shape(shape) = &drawing.kind else { panic!("shape expected") };
        let line = shape.line.as_ref().expect("patterned rule");
        assert_eq!(line.pattern.as_deref(), Some("pct10"));
        assert_eq!(line.color, "FF0000");
        assert_eq!(line.back_color.as_deref(), Some("00FF00"));
    }

    /// A shape's run can name the theme's scheme instead of a face.
    #[test]
    fn a_run_that_names_the_theme_wears_the_face_the_theme_names() {
        let theme = Theme {
            major_face: Some("游ゴシック Light".into()),
            minor_face: Some("游ゴシック".into()),
            ..Theme::default()
        };
        let xml = r##"<xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing"
          xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
  <xdr:twoCellAnchor>
    <xdr:from><xdr:col>0</xdr:col><xdr:colOff>0</xdr:colOff>
      <xdr:row>0</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from>
    <xdr:to><xdr:col>8</xdr:col><xdr:colOff>0</xdr:colOff>
      <xdr:row>2</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:to>
    <xdr:sp>
      <xdr:spPr><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></xdr:spPr>
      <xdr:txBody>
        <a:bodyPr/>
        <a:p><a:r>
          <a:rPr sz="1200"><a:latin typeface="+mn-lt"/><a:ea typeface="+mn-ea"/></a:rPr>
          <a:t>昨年からの変更点</a:t>
        </a:r></a:p>
        <a:p><a:r>
          <a:rPr sz="1200"><a:ea typeface="+mj-ea"/></a:rPr>
          <a:t>見出し</a:t>
        </a:r></a:p>
        <a:p><a:r>
          <a:rPr sz="1200"><a:ea typeface="メイリオ"/></a:rPr>
          <a:t>名指し</a:t>
        </a:r></a:p>
      </xdr:txBody>
    </xdr:sp>
    <xdr:clientData/>
  </xdr:twoCellAnchor>
</xdr:wsDr>"##;
        let found = parse_drawing_xml(xml, &theme);
        let crate::ir::DrawingKind::Shape(shape) = &found[0].0.kind else {
            panic!("the anchor holds a shape");
        };
        let said = shape.text.as_ref().expect("the shape says something");
        assert_eq!(said.paragraphs[0].face.as_deref(), Some("游ゴシック"));
        assert_eq!(said.paragraphs[1].face.as_deref(), Some("游ゴシック Light"));
        assert_eq!(
            said.paragraphs[2].face.as_deref(),
            Some("メイリオ"),
            "a face named outright is left alone"
        );
    }

    /// A shape can name a theme colour and shade it rather than state one.
    #[test]
    fn a_scheme_colour_is_resolved_and_shaded() {
        let theme = Theme {
            colours: vec![
                "000000".into(), "FFFFFF".into(), "44546A".into(), "E7E6E6".into(),
                "4472C4".into(),
            ],
            ..Theme::default()
        };
        let xml = r##"<xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing"
          xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
  <xdr:twoCellAnchor>
    <xdr:from><xdr:col>0</xdr:col><xdr:colOff>0</xdr:colOff>
      <xdr:row>0</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from>
    <xdr:to><xdr:col>1</xdr:col><xdr:colOff>0</xdr:colOff>
      <xdr:row>1</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:to>
    <xdr:sp>
      <xdr:spPr>
        <a:prstGeom prst="line"><a:avLst/></a:prstGeom>
        <a:solidFill><a:schemeClr val="accent1"/></a:solidFill>
        <a:ln><a:solidFill><a:schemeClr val="accent1"><a:shade val="50000"/></a:schemeClr></a:solidFill></a:ln>
      </xdr:spPr>
    </xdr:sp>
    <xdr:clientData/>
  </xdr:twoCellAnchor>
</xdr:wsDr>"##;
        let found = parse_drawing_xml(xml, &theme);
        let crate::ir::DrawingKind::Shape(shape) = &found[0].0.kind else {
            panic!("the anchor holds a shape");
        };
        assert_eq!(shape.fill.as_deref(), Some("4472C4"));
        assert_eq!(
            shape.line.as_ref().map(|line| line.color.as_str()),
            Some("2F528F"),
            "a 50% shade halves the LIGHT, not the byte 遯ｶ繝ｻhalving the byte \
             would give 223962, and Excel draws neither that nor anything \
             near it (see `_xlsx_shade_curve.py`)"
        );
    }

    #[test]
    fn a_text_run_retains_theme_colour_brightness() {
        let xml = r#"<xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><xdr:twoCellAnchor><xdr:from><xdr:col>0</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>0</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from><xdr:to><xdr:col>1</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>1</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:to><xdr:sp><xdr:txBody><a:bodyPr/><a:p><a:r><a:rPr><a:solidFill><a:schemeClr val="accent1"><a:lumMod val="60000"/><a:lumOff val="40000"/></a:schemeClr></a:solidFill></a:rPr><a:t>x</a:t></a:r></a:p></xdr:txBody></xdr:sp></xdr:twoCellAnchor></xdr:wsDr>"#;
        let found = parse_drawing_xml(xml, &Theme::default());
        let crate::ir::DrawingKind::Shape(shape) = &found[0].0.kind else { panic!("shape expected") };
        let run = &shape.text.as_ref().unwrap().paragraphs[0].runs[0];
        assert_eq!(run.color_theme, Some(5));
        assert_eq!(run.color_brightness, Some(0.4));
    }

    /// How light a colour is moves through HSL, and Excel's own picture is
    /// what says so: `002`'s banner is accent1 遯ｶ繝ｻ5B9BD5 in that workbook's
    /// theme 遯ｶ繝ｻunder `lumMod 20% lumOff 80%`, and Excel paints it DEEBF7.
    #[test]
    fn lightness_is_moved_the_way_excel_moves_it() {
        assert_eq!(
            shaded("5B9BD5", &[("lumMod".into(), 0.2), ("lumOff".into(), 0.8)]),
            "DEEBF7"
        );
        // tx1 at 65% with 35% added is the grey Excel rules a table with.
        assert_eq!(
            shaded("000000", &[("lumMod".into(), 0.65), ("lumOff".into(), 0.35)]),
            "595959"
        );
    }

    /// An indent is a level, not a measurement, and it survives whichever way
    /// the alignment element is written: quick-xml hands an empty element to a
    /// different arm than one with children, and only one of them used to be
    /// read.
    #[test]
    fn an_indent_is_read_either_way_the_alignment_is_written() {
        let xml = r##"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <fonts count="1"><font><sz val="11"/><name val="Calibri"/></font></fonts>
  <cellStyleXfs count="1">
    <xf numFmtId="0" fontId="0" fillId="0" borderId="0"/>
  </cellStyleXfs>
  <cellXfs count="3">
    <xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>
    <xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0" applyAlignment="1">
      <alignment horizontal="left" indent="2"/>
    </xf>
    <xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0" applyAlignment="1">
      <alignment horizontal="distributed" indent="3" vertical="center"/>
    </xf>
  </cellXfs>
</styleSheet>"##;
        let sheet = parse_styles_xml(xml, &Theme::default()).unwrap();
        assert_eq!(resolve_cell_style(0, &sheet).indent, 0);
        assert_eq!(resolve_cell_style(1, &sheet).indent, 2);
        assert_eq!(resolve_cell_style(2, &sheet).indent, 3);
    }

    /// The font an indent is measured in is the first one in the list, not
    /// the one the Normal style points at. The two are the same in a workbook
    /// Excel writes and different in one openpyxl does, and Excel's indent
    /// follows the first entry in both (`_xlsx_indent_bisect.py`).
    #[test]
    fn the_first_font_is_kept_apart_from_the_normal_style_font() {
        let styles = r##"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <fonts count="2">
    <font><sz val="11"/><name val="Calibri"/></font>
    <font><sz val="9"/><name val="Meiryo UI"/></font>
  </fonts>
  <cellStyleXfs count="1">
    <xf numFmtId="0" fontId="1" fillId="0" borderId="0"/>
  </cellStyleXfs>
  <cellXfs count="1">
    <xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>
  </cellXfs>
</styleSheet>"##;
        let stylesheet = parse_styles_xml(styles, &Theme::default()).unwrap();
        let xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
      <c r="A1" t="inlineStr"><is><t>Agreed</t></is></c>
</worksheet>"#;
        let sheet = parse_worksheet(xml, "probe", &[], &stylesheet, &Theme::default()).expect("should parse");
        assert_eq!(sheet.first_font, Some(("Calibri".to_string(), 11.0)));
        assert_eq!(sheet.normal_font, Some(("Meiryo UI".to_string(), 9.0)));
    }

    /// The pieces a shared string is dressed in add up to the string: a
    /// reader that draws them one after another must not be handed more text
    /// The pieces of a shared string add up to the string, and a
    /// of several rich-text pieces, while phonetic text remains separate.
    /// shows.
    #[test]
    fn the_pieces_of_a_shared_string_add_up_to_it() {
        let xml = concat!(
            r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>"#,
            r#"<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">"#,
            r#"<si><r><rPr><sz val="20"/><rFont val="Meiryo UI"/></rPr>"#,
            "<t xml:space=\"preserve\">\n\n邵ｺ・ｯ邵ｺ蛟･・・</t></r>",
            r#"<r><rPr><sz val="14"/><rFont val="Meiryo UI"/></rPr>"#,
            "<t xml:space=\"preserve\">\n\n</t></r>",
            r#"<r><rPr><sz val="18"/><rFont val="Meiryo UI"/></rPr><t>鬮ｮ繝ｻ・ｨ驛・ｽ｡・ｨ</t></r>"#,
            r#"<rPh sb="25" eb="28"><t>郢ｧ・ｷ郢晢ｽ･郢ｧ・ｦ郢ｧ・ｱ郢ｧ・､郢晏・ﾎ咏ｹｧ・ｦ</t></rPh>"#,
            r#"<phoneticPr fontId="9"/></si></sst>"#,
        );
        let held = parse_shared_strings(xml).expect("should parse");
        let string = &held[0];
        let pieces: usize = string.runs.iter().map(|run| run.text.chars().count()).sum();
        assert_eq!(string.text.chars().count(), pieces, "text {:?} runs {:?}",
                   string.text, string.runs.iter().map(|r| &r.text).collect::<Vec<_>>());
        assert_eq!(string.runs.len(), 3);
        // The phonetic guide is in neither.
        assert!(!string.text.contains('\u{30b7}'));
    }

    #[test]
    fn a_diagonal_is_read_whether_or_not_it_states_a_colour() {
        // A `<diagonal/>` that names a colour has a child and so arrives as a
        // different event than one that closes itself. Both spellings rule
        // the cell corner to corner, so both are pinned here.
        let xml = r##"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <fonts count="1"><font><sz val="11"/><name val="Calibri"/></font></fonts>
  <borders count="3">
    <border><left/><right/><top/><bottom/><diagonal/></border>
    <border diagonalUp="1">
      <left/><right/><top/><bottom/>
      <diagonal style="thin"><color indexed="64"/></diagonal>
    </border>
    <border diagonalDown="1">
      <left/><right/><top/><bottom/><diagonal style="medium"/>
    </border>
  </borders>
  <cellStyleXfs count="1">
    <xf numFmtId="0" fontId="0" fillId="0" borderId="0"/>
  </cellStyleXfs>
  <cellXfs count="3">
    <xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>
    <xf numFmtId="0" fontId="0" fillId="0" borderId="1" xfId="0" applyBorder="1"/>
    <xf numFmtId="0" fontId="0" fillId="0" borderId="2" xfId="0" applyBorder="1"/>
  </cellXfs>
</styleSheet>"##;
        let sheet = parse_styles_xml(xml, &Theme::default()).unwrap();

        let plain = resolve_cell_style(0, &sheet);
        assert!(plain.border_diagonal.is_none());

        let up = resolve_cell_style(1, &sheet);
        assert_eq!(up.border_diagonal.as_ref().map(|l| l.style.as_str()), Some("thin"));
        assert!(up.diagonal_up);
        assert!(!up.diagonal_down);

        let down = resolve_cell_style(2, &sheet);
        assert_eq!(down.border_diagonal.as_ref().map(|l| l.style.as_str()), Some("medium"));
        assert!(down.diagonal_down);
        assert!(!down.diagonal_up);
    }

    #[test]
    fn test_parse_styles_xml() {
        let xml = r##"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <numFmts count="1">
    <numFmt numFmtId="164" formatCode="#,##0.00_ "/>
  </numFmts>
  <fonts count="2">
    <font><sz val="11"/><color rgb="FF000000"/><name val="Calibri"/></font>
    <font><b/><sz val="14"/><color rgb="FFFF0000"/><name val="Calibri"/></font>
  </fonts>
  <fills count="2">
    <fill><patternFill patternType="none"/></fill>
    <fill><patternFill patternType="solid"><fgColor rgb="FFFFFF00"/></patternFill></fill>
  </fills>
  <borders count="2">
    <border><left/><right/><top/><bottom/></border>
    <border><left style="thin"/><right style="thin"/><top style="thin"/><bottom style="thin"/></border>
  </borders>
  <cellXfs count="2">
    <xf numFmtId="0" fontId="0" fillId="0" borderId="0"/>
    <xf numFmtId="164" fontId="1" fillId="1" borderId="1"><alignment horizontal="center"/></xf>
  </cellXfs>
</styleSheet>"##;
        let ss = parse_styles_xml(xml, &Theme::default()).unwrap();
        assert_eq!(ss.num_fmts.len(), 1);
        assert_eq!(ss.num_fmts.get(&164).unwrap(), "#,##0.00_ ");
        assert_eq!(ss.fonts.len(), 2);
        assert!(ss.fonts[1].bold);
        assert_eq!(ss.fonts[1].size, Some(14.0));
        assert_eq!(ss.fonts[1].color.as_deref(), Some("FF0000"));
        assert_eq!(ss.fills.len(), 2);
        assert_eq!(ss.fills[1].bg_color.as_deref(), Some("FFFF00"));
        assert_eq!(ss.borders.len(), 2);
        assert!(ss.borders[0].left.is_none());
        assert!(ss.borders[1].left.is_some());
        assert!(ss.borders[1].right.is_some());
        assert!(ss.borders[1].top.is_some());
        assert!(ss.borders[1].bottom.is_some());
        assert_eq!(ss.cell_xfs.len(), 2);
        assert_eq!(ss.cell_xfs[1].num_fmt_id, 164);
        assert_eq!(
            ss.cell_xfs[1].horizontal_align.as_deref(),
            Some("center")
        );

        // Test resolve_cell_style
        let style = resolve_cell_style(1, &ss);
        assert!(style.bold);
        assert_eq!(style.font_color.as_deref(), Some("FF0000"));
        assert_eq!(style.bg_color.as_deref(), Some("FFFF00"));
        assert_eq!(style.number_format.as_deref(), Some("#,##0.00_ "));
        assert_eq!(style.horizontal_align.as_deref(), Some("center"));
        assert!(style.border_top.is_some());
        assert!(style.border_bottom.is_some());
        assert!(style.border_left.is_some());
        assert!(style.border_right.is_some());
    }

    #[test]
    fn test_builtin_number_formats() {
        assert_eq!(builtin_number_format(0), Some("General"));
        assert_eq!(builtin_number_format(3), Some("#,##0"));
        assert_eq!(builtin_number_format(14), Some("mm-dd-yy"));
        assert_eq!(builtin_number_format(99), None);
    }
}

#[cfg(test)]
mod threaded_comments {
    use super::{parse_person_list, parse_threaded_comments};

    #[test]
    fn parses_person_display_and_identity_metadata() {
        let people = parse_person_list(r#"<personList><person id="p1" displayName="Excel User" userId="user-123" providerId="Windows Live"/></personList>"#);
        let person = people.get("p1").unwrap();
        assert_eq!(person.display_name, "Excel User");
        assert_eq!(person.user_id, "user-123");
        assert_eq!(person.provider_id, "Windows Live");
    }

    #[test]
    fn parses_threaded_comment_text_and_nested_reply_identity() {
        let xml = r#"<ThreadedComments xmlns="http://schemas.microsoft.com/office/spreadsheetml/2018/threadedcomments">
          <threadedComment ref="C5" dT="2026-10-23T13:55:31.50" personId="person-a" id="thread-a">
            <text>thread &amp; root</text>
            <threadedComment ref="C5" dT="2026-10-23T13:56:00Z" personId="person-b" id="reply-b" parentId="thread-a">
              <text>reply text</text>
            </threadedComment>
          </threadedComment>
        </ThreadedComments>"#;
        let comments = parse_threaded_comments(xml);
        assert_eq!(comments.len(), 2);
        assert_eq!(comments[0].cell, (4, 2));
        assert_eq!(comments[0].id, "thread-a");
        assert_eq!(comments[0].person_id, "person-a");
        assert_eq!(comments[0].text, "thread & root");
        assert_eq!(comments[0].parent_id, None);
        assert_eq!(comments[1].id, "reply-b");
        assert_eq!(comments[1].person_id, "person-b");
        assert_eq!(comments[1].parent_id.as_deref(), Some("thread-a"));
        assert_eq!(comments[1].text, "reply text");
    }
}

#[cfg(test)]
mod note_room {
    #[test]
    fn escape_sequences_do_not_cross_rich_text_boundaries() {
        let comments = r#"<comments><authors><author>a</author></authors><commentList><comment ref="B2" authorId="0"><text><r><t>_x0</t></r><r><t>041_ and _x005F_x0041_</t></r><r><t>_xDE00_</t></r></text></comment></commentList></comments>"#;
        let vml = r#"<v:shape style='position:absolute;width:100pt;height:50pt'><x:ClientData><x:Anchor>1, 0, 1, 0, 3, 0, 3, 0</x:Anchor><x:Row>1</x:Row><x:Column>1</x:Column></x:ClientData></v:shape>"#;
        let notes = super::parse_comments(comments, vml);
        let expected: Vec<u16> = "_x0041_ and _x0041_".encode_utf16().chain([0xDE00]).collect();
        assert_eq!(notes[0].utf16_text.as_ref(), Some(&expected));
        assert_eq!(notes[0].text.paragraphs[0].text, "_x0041_ and _x0041_\u{fffd}");
    }

    #[test]
    fn note_reads_each_rich_text_run() {
        let comments = r#"<comments><authors><author>a</author></authors><commentList><comment ref="B2" authorId="0"><text><r><rPr><b/><sz val="9"/></rPr><t>ab</t></r><r><rPr><b/><i/><sz val="9"/></rPr><t>cd</t></r><r><rPr><color indexed="10"/></rPr><t>ef</t></r></text></comment></commentList></comments>"#;
        let vml = r#"<v:shape style='position:absolute;width:100pt;height:50pt'><x:ClientData><x:Anchor>1, 0, 1, 0, 3, 0, 3, 0</x:Anchor><x:Row>1</x:Row><x:Column>1</x:Column></x:ClientData></v:shape>"#;
        let notes = super::parse_comments(comments, vml);
        let runs = &notes[0].text.paragraphs[0].runs;
        assert_eq!(runs.iter().map(|run| run.text.as_str()).collect::<Vec<_>>(), ["ab", "cd", "ef"]);
        assert_eq!(runs.iter().map(|run| run.italic).collect::<Vec<_>>(), [Some(false), Some(true), Some(false)]);
        assert_eq!(runs[2].color.as_deref(), Some("FF0000"));
    }
    /// A note keeps 7.2 points either side and 3.6 above and below unless it
    /// says otherwise, which is what Excel makes one with. The corpus's floor
    /// workbook states 2.5mm and 2.3mm instead, and those were being given to
    /// every note in every workbook.
    #[test]
    fn a_note_keeps_the_room_it_states_and_excels_own_when_it_states_none() {
        let vml = |textbox: &str| {
            format!(
                "<v:shape style='position:absolute;width:100pt;height:50pt'                 fillcolor=\"#ffffe1\"><v:textbox {textbox}/><x:ClientData>                 <x:Anchor>1, 0, 1, 0, 3, 0, 3, 0</x:Anchor>                 <x:Row>1</x:Row><x:Column>1</x:Column><x:Visible/>                 </x:ClientData></v:shape>"
            )
        };
        let comments = "<comments><authors><author>a</author></authors>            <commentList><comment ref=\"B2\" authorId=\"0\"><text><r>            <t>said</t></r></text></comment></commentList></comments>";
        let bare = super::parse_comments(comments, &vml("style='mso-direction-alt:auto'"));
        assert_eq!(bare.len(), 1);
        assert_eq!(bare[0].author, "a");
        assert_eq!(bare[0].text.insets, (91440, 45720, 91440, 45720));
        let stated = super::parse_comments(
            comments,
            &vml("style='mso-direction-alt:auto' inset=\"2.5mm,2.3mm,2.5mm,2.3mm\""),
        );
        assert_eq!(stated[0].text.insets, (90000, 82800, 90000, 82800));
        // Excel writes only the part it changed, leaving the rest to default.
        let partial = super::parse_comments(
            comments,
            &vml("style='mso-direction-alt:auto' inset=\",0\""),
        );
        assert_eq!(partial[0].text.insets, (91440, 0, 91440, 45720));
    }
}

#[cfg(test)]
mod note_line_style {
    #[test]
    fn reads_vml_outline_color_visibility_and_width() {
        let comments = r#"<comments><authors><author>a</author></authors><commentList><comment ref="B2" authorId="0"><text><t>said</t></text></comment></commentList></comments>"#;
        let vml = r##"<v:shape style='position:absolute;width:100pt;height:50pt' fillcolor="#ffffe1" strokecolor="#FF0000" strokeweight="2.25pt" stroked="f"><v:textbox/><x:ClientData><x:Anchor>1, 0, 1, 0, 3, 0, 3, 0</x:Anchor><x:Row>1</x:Row><x:Column>1</x:Column></x:ClientData></v:shape>"##;
        let parsed = super::parse_comments(comments, vml);
        assert_eq!(parsed.len(), 1);
        assert_eq!(parsed[0].line_color.as_deref(), Some("FF0000"));
        assert!(!parsed[0].line_visible);
        assert_eq!(parsed[0].line_weight, 2.25);

        let default_vml = r##"<v:shape style='position:absolute;width:100pt;height:50pt' fillcolor="#ffffe1"><v:textbox/><x:ClientData><x:Anchor>1, 0, 1, 0, 3, 0, 3, 0</x:Anchor><x:Row>1</x:Row><x:Column>1</x:Column></x:ClientData></v:shape>"##;
        let defaulted = super::parse_comments(comments, default_vml);
        assert!(defaulted[0].line_visible);
        assert_eq!(defaulted[0].line_weight, 0.75);
        let hidden_fill_vml = default_vml.replace("<v:shape ", "<v:shape filled=\"f\" ");
        let hidden_fill = super::parse_comments(comments, &hidden_fill_vml);
        assert!(!hidden_fill[0].fill_visible);
        let transparency_vml = default_vml.replace(
            "<v:textbox/>",
            "<v:fill opacity=\"39322f\"/><v:stroke opacity=\"39322f\"/><v:textbox/>",
        );
        let transparent = super::parse_comments(comments, &transparency_vml);
        assert!((transparent[0].fill_transparency - 0.399_993_9).abs() < 0.000_001);
        assert!((transparent[0].line_transparency - 0.399_993_9).abs() < 0.000_001);

        let system_fill_vml = r##"<v:shape style='position:absolute;width:100pt;height:50pt' fillcolor="infoBackground [80]"><v:textbox/><x:ClientData><x:Anchor>1, 0, 1, 0, 3, 0, 3, 0</x:Anchor><x:Row>1</x:Row><x:Column>1</x:Column></x:ClientData></v:shape>"##;
        let system_filled = super::parse_comments(comments, system_fill_vml);
        assert_eq!(system_filled[0].fill.as_deref(), Some("FFFFE1"));
    }
}

#[cfg(test)]
mod theme_tints {
    use super::tinted;

    fn close_to(got: &str, want: &str) {
        let byte = |hex: &str, at: usize| {
            u8::from_str_radix(&hex[at..at + 2], 16).expect("two hex digits") as i32
        };
        for at in [0, 2, 4] {
            let apart = (byte(got, at) - byte(want, at)).abs();
            assert!(apart <= 1, "{got} is not within a shade of {want}");
        }
    }

    /// Both colours were read off a worksheet Excel drew: a table's banded row
    /// is its header colour under a tint of 0.8. Rounding puts us within one
    /// step of Excel on each channel.
    #[test]
    fn a_tint_lightens_without_changing_the_hue() {
        close_to(&tinted("156082", 0.8), "C0E6F5");
        close_to(&tinted("4EA72E", 0.8), "DAF2D0");
    }

    #[test]
    fn a_negative_tint_darkens() {
        close_to(&tinted("FFFFFF", -0.5), "808080");
        assert_eq!(tinted("156082", 0.0), "156082");
    }

}

#[cfg(test)]
mod utf8_bom_table_fragment_tests {
    #[test]
    fn inherited_filter_namespaces_preserve_bom_and_complete_tag(){
        let xml="<autoFilter ref='A1:B2'/>";
        let namespaces=vec![(b"xmlns:x".to_vec(),b"urn:extra".to_vec())];
        let plain=super::table_filter_xml_with_namespaces(xml,&namespaces).unwrap();
        let bom=super::table_filter_xml_with_namespaces(&format!("\u{feff}{xml}"),&namespaces).unwrap();
        assert_eq!(bom,format!("\u{feff}{plain}"));
        assert!(bom.contains("ref=\"A1:B2\"")||bom.contains("ref='A1:B2'"));
    }
}
