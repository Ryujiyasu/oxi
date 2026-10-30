// This Source Code Form is subject to the terms of the Mozilla Public
// License, v. 2.0. If a copy of the MPL was not distributed with this
// file, You can obtain one at https://mozilla.org/MPL/2.0/.

use std::collections::HashMap;

use quick_xml::events::Event;
use quick_xml::reader::Reader;
use thiserror::Error;

use oxidocs_common::archive::OoxmlArchive;
use oxidocs_common::relationships::{parse_relationships, Relationship};

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

/// Whether a cell format applies a part of itself, rather than taking that
/// part from the named style it is built on.
///
/// Absent is not the same as `0` here. Excel leaves the flag off when the
/// format's own value is the one to use, and writes `applyFont="0"` when it
/// is not 遯ｶ繝ｻa cell that draws its blue underline from the Hyperlink style
/// says so explicitly. Writers that are not Excel commonly leave every flag
/// off while still naming a font on each cell, and reading absent as "do not
/// apply" throws that font away and draws the whole sheet in the workbook's
/// default face.
fn unless_denied(value: Option<&str>) -> bool {
    !matches!(value, Some("0") | Some("false"))
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
                    "u" if in_run_props => {
                        current_run.underline = !matches!(
                            get_attr(&e, "val").as_deref(),
                            Some("0") | Some("none")
                        );
                    }
                    "vertAlign" if in_run_props => {
                        current_run.vert_align = get_attr(&e, "val");
                    }
                    "color" if in_run_props => {
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
    ExternalBook { sheets }
}

/// What a cached cell held, from the shorthand the link part writes.
fn cached_cell_value(kind: &str, text: &str) -> CellValue {
    match kind {
        "str" | "s" => CellValue::String(text.to_string()),
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
    strikethrough: bool,
    size: Option<f32>,
    color: Option<String>,
    name: Option<String>,
    latin_name: Option<String>,
    charset: Option<i32>,
    family: Option<i32>,
}

#[derive(Debug, Clone, Default)]
struct FillInfo {
    bg_color: Option<String>,
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

#[derive(Debug, Clone, Default)]
struct XfRecord {
    num_fmt_id: u32,
    font_id: usize,
    fill_id: usize,
    border_id: usize,
    /// The named style this format is built on, and whether it overrides each
    /// part of it. A cell format that does not apply its own font wears the
    /// font of the style it names 遯ｶ繝ｻthat is how Excel dresses a hyperlink.
    style_id: Option<usize>,
    applies_font: bool,
    applies_fill: bool,
    applies_border: bool,
    applies_number_format: bool,
    horizontal_align: Option<String>,
    vertical_align: Option<String>,
    indent: u32,
    wrap_text: bool,
    stacked_text: bool,
    shrink_to_fit: bool,
    /// `quotePrefix="1"`: the entry was typed behind an apostrophe.
    quote_prefix: bool,
    /// Whether the cell may still be edited once its sheet is protected.
    /// Excel locks every cell by default and unlocks the few a form wants
    /// typed into, so the negative is the one worth holding: it keeps this
    /// struct right when it is built from its own default.
    unlocked: bool,
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

/// Moves a colour toward white or black, the way a theme tint does. Only how
/// light it is changes; its hue and how saturated it is do not, which is why
/// this goes through HSL rather than nudging each channel.
fn tinted(hex: &str, tint: f32) -> String {
    let Ok(value) = u32::from_str_radix(hex, 16) else {
        return hex.to_string();
    };
    let channel = |shift: u32| ((value >> shift) & 0xFF) as f32 / 255.0;
    let (red, green, blue) = (channel(16), channel(8), channel(0));
    let high = red.max(green).max(blue);
    let low = red.min(green).min(blue);
    let lum = (high + low) / 2.0;
    let span = high - low;
    let sat = if span == 0.0 {
        0.0
    } else if lum < 0.5 {
        span / (high + low)
    } else {
        span / (2.0 - high - low)
    };
    let hue = if span == 0.0 {
        0.0
    } else if high == red {
        (((green - blue) / span) % 6.0 + 6.0) % 6.0
    } else if high == green {
        (blue - red) / span + 2.0
    } else {
        (red - green) / span + 4.0
    };

    let moved = if tint < 0.0 {
        lum * (1.0 + tint)
    } else {
        lum * (1.0 - tint) + tint
    };

    let chroma = (1.0 - (2.0 * moved - 1.0).abs()) * sat;
    let second = chroma * (1.0 - ((hue % 2.0) - 1.0).abs());
    let (red, green, blue) = match hue as u32 {
        0 => (chroma, second, 0.0),
        1 => (second, chroma, 0.0),
        2 => (0.0, chroma, second),
        3 => (0.0, second, chroma),
        4 => (second, 0.0, chroma),
        _ => (chroma, 0.0, second),
    };
    let base = moved - chroma / 2.0;
    let byte = |part: f32| (((part + base).clamp(0.0, 1.0) * 255.0).round() as u32).min(255);
    format!("{:02X}{:02X}{:02X}", byte(red), byte(green), byte(blue))
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

fn parse_color_attr(e: &quick_xml::events::BytesStart, theme: &Theme) -> Option<String> {
    if let Some(rgb) = get_attr(e, "rgb") {
        // Strip leading alpha if 8-char hex
        let hex = if rgb.len() == 8 { &rgb[2..] } else { &rgb };
        return Some(hex.to_string());
    }
    if let Some(index) = get_attr(e, "indexed").and_then(|at| at.parse::<usize>().ok()) {
        return indexed_colour(index, theme);
    }
    let index = get_attr(e, "theme")?.parse::<usize>().ok()?;
    let named = theme.colour(index)?;
    let tint = get_attr(e, "tint")
        .and_then(|value| value.parse::<f32>().ok())
        .unwrap_or(0.0);
    Some(if tint == 0.0 {
        named.to_string()
    } else {
        tinted(named, tint)
    })
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
        let offset = reader.buffer_position() as usize;
        let event = reader.read_event();
        let event_end = reader.buffer_position() as usize;
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
                    "numFmt" => style.number_format = get_attr(e, "formatCode"),
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

fn parse_styles_xml(xml: &str, theme: &Theme) -> Result<StyleSheet, XlsxError> {
    let mut reader = Reader::from_str(xml);
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
        match reader.read_event()? {
            Event::Start(e) => {
                let name = local_name(e.name().as_ref());
                match name.as_str() {
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
                            shrink_to_fit: false,
                            quote_prefix: is_true(get_attr(&e, "quotePrefix").as_deref()),
                            style_id: get_attr(&e, "xfId").and_then(|v| v.parse().ok()),
                            applies_font: unless_denied(
                                get_attr(&e, "applyFont").as_deref(),
                            ),
                            applies_fill: unless_denied(
                                get_attr(&e, "applyFill").as_deref(),
                            ),
                            applies_border: unless_denied(
                                get_attr(&e, "applyBorder").as_deref(),
                            ),
                            applies_number_format: unless_denied(
                                get_attr(&e, "applyNumberFormat").as_deref(),
                            ),
                            unlocked: false,
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
                        current_xf.shrink_to_fit =
                            is_true(get_attr(&e, "shrinkToFit").as_deref());
                    }
                    // <protection locked="0"/> 遯ｶ繝ｻthe cells a protected sheet
                    // still lets a person type into.
                    "protection" if in_xf => {
                        current_xf.unlocked =
                            matches!(get_attr(&e, "locked").as_deref(), Some("0") | Some("false"));
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
                        }
                    }

                    // Font color
                    "color" if in_font => {
                        if let Some(c) = parse_color_attr(&e, theme) {
                            current_font.color = Some(c);
                        }
                    }

                    // Fill color 遯ｶ繝ｻlook for fgColor inside patternFill
                    "fgColor" if in_fill => {
                        if let Some(c) = parse_color_attr(&e, theme) {
                            current_fill.bg_color = Some(c);
                        }
                    }

                    _ => {}
                }
            }
            Event::Empty(e) => {
                let name = local_name(e.name().as_ref());
                match name.as_str() {
                    "numFmt" if section == Section::NumFmts => {
                        if let (Some(id_str), Some(code)) =
                            (get_attr(&e, "numFmtId"), get_attr(&e, "formatCode"))
                        {
                            if let Ok(id) = id_str.parse::<u32>() {
                                ss.num_fmts.insert(id, code);
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
                        // <u/> underlines; <u val="none"/> does not.
                        let val = get_attr(&e, "val");
                        current_font.underline = !matches!(
                            val.as_deref(),
                            Some("0") | Some("none")
                        );
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
                        current_font.name = get_attr(&e, "val");
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
                        }
                    }
                    "color" if in_font => {
                        if let Some(c) = parse_color_attr(&e, theme) {
                            current_font.color = Some(c);
                        }
                    }
                    "fgColor" if in_fill => {
                        if let Some(c) = parse_color_attr(&e, theme) {
                            current_fill.bg_color = Some(c);
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
                        current_xf.shrink_to_fit =
                            is_true(get_attr(&e, "shrinkToFit").as_deref());
                    }
                    // <protection locked="0"/> 遯ｶ繝ｻthe cells a protected sheet
                    // still lets a person type into.
                    "protection" if in_xf => {
                        current_xf.unlocked =
                            matches!(get_attr(&e, "locked").as_deref(), Some("0") | Some("false"));
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
                            shrink_to_fit: false,
                            quote_prefix: is_true(get_attr(&e, "quotePrefix").as_deref()),
                            style_id: get_attr(&e, "xfId").and_then(|v| v.parse().ok()),
                            applies_font: unless_denied(
                                get_attr(&e, "applyFont").as_deref(),
                            ),
                            applies_fill: unless_denied(
                                get_attr(&e, "applyFill").as_deref(),
                            ),
                            applies_border: unless_denied(
                                get_attr(&e, "applyBorder").as_deref(),
                            ),
                            applies_number_format: unless_denied(
                                get_attr(&e, "applyNumberFormat").as_deref(),
                            ),
                            unlocked: false,
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
                        }
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
    Some(BorderLine { style, color: None })
}

/// Build a CellStyle from a style index referencing the StyleSheet.
fn resolve_cell_style(style_index: usize, stylesheet: &StyleSheet) -> CellStyle {
    let xf = match stylesheet.cell_xfs.get(style_index) {
        Some(xf) => xf,
        None => return CellStyle::default(),
    };

    // A cell format that does not apply a part of itself wears that part from
    // the named style it is built on. This is how a hyperlink gets its blue
    // underline: the cell's own format names no font at all.
    let parent = xf
        .style_id
        .and_then(|id| stylesheet.cell_style_xfs.get(id));
    let inherited = |applies: bool, own: usize, from: fn(&XfRecord) -> usize| {
        match (applies, parent) {
            (false, Some(parent)) => from(parent),
            _ => own,
        }
    };
    let font_id = inherited(xf.applies_font, xf.font_id, |xf| xf.font_id);
    let fill_id = inherited(xf.applies_fill, xf.fill_id, |xf| xf.fill_id);
    let border_id = inherited(xf.applies_border, xf.border_id, |xf| xf.border_id);

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
    let fill = stylesheet.fills.get(fill_id).cloned().unwrap_or_default();
    let border = stylesheet
        .borders
        .get(border_id)
        .cloned()
        .unwrap_or_default();

    // Resolve number format
    let num_fmt_id = match (xf.applies_number_format, parent) {
        (false, Some(parent)) => parent.num_fmt_id,
        _ => xf.num_fmt_id,
    };
    let number_format = if num_fmt_id == 0 {
        None // General 遯ｶ繝ｻno explicit format needed
    } else if let Some(custom) = stylesheet.num_fmts.get(&num_fmt_id) {
        Some(custom.clone())
    } else {
        builtin_number_format(num_fmt_id).map(|s| s.to_string())
    };

    CellStyle {
        bold: font.bold,
        italic: font.italic,
        underline: font.underline,
        strikethrough: font.strikethrough,
        font_size: font.size,
        font_name: font.name.clone(),
        font_charset: font.charset,
        font_family: font.family,
        font_color: font.color,
        bg_color: fill.bg_color,
        number_format,
        horizontal_align: xf.horizontal_align.clone(),
        indent: xf.indent,
        vertical_align: xf.vertical_align.clone(),
        wrap_text: xf.wrap_text,
        stacked_text: xf.stacked_text,
        shrink_to_fit: xf.shrink_to_fit,
        quote_prefix: xf.quote_prefix,
        unlocked: xf.unlocked,
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
    let mut group_name: Option<String> = None;
    let mut reader = Reader::from_str(xml);
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
                "cNvPr"=> {
                    let name = get_attr(e, "name");
                    drawing_id = get_attr(e, "id").and_then(|value| value.parse().ok());
                    if depth_in_shape == 1 && group.is_none() && parent_groups.is_empty() {
                        // Excel exposes nested group leaves through the outer
                        // group's flat GroupItems collection and ParentGroup.
                        group_name = name.clone();
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
                    group_name = None;
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
                                *held = shape.clone();
                            }
                            let picture = matches!(kind, DrawingKind::Picture { .. });
                            found.push((
                                Drawing {
                                    name: drawing_name.take(),
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
        let before = reader.buffer_position() as usize;
        let event = reader.read_event().ok()?;
        let after = reader.buffer_position() as usize;
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
        let before = reader.buffer_position() as usize;
        let event = reader.read_event_into(&mut buf);
        match event {
            Ok(Event::Start(ref e)) | Ok(Event::Empty(ref e)) => {
                match local_name(e.name().as_ref()).as_str() {
                    "table" => {
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
                            auto_filter_xml = table_filter_xml_with_namespaces(&xml[before..reader.buffer_position() as usize], &namespaces);
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
                    auto_filter_xml = table_filter_xml_with_namespaces(&xml[start..reader.buffer_position() as usize], &namespaces);
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
        let start = reader.buffer_position() as usize;
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
                    if ids.len() == 1 { result.insert(ids[0].clone(), xml[first..reader.buffer_position() as usize].to_string()); }
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
    let mut hidden_cols: Vec<u32> = Vec::new();
    let mut auto_filter: Option<crate::ir::AutoFilter> = None;
    let mut declared_range: Option<(u32, u32, u32, u32)> = None;
    let mut unsupported: Vec<String> = Vec::new();
    // Whether the sheet is protected, which decides nothing about how it is
    // drawn and everything about whether it can be typed into.
    let mut protected = false;
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
        let event_start = reader.buffer_position() as usize;
        match reader.read_event()? {
            Event::Start(e) => {
                let name = local_name(e.name().as_ref());
                note_unsupported(&name, &mut unsupported);
                match name.as_str() {
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
                        inline_start = Some(reader.buffer_position() as usize);
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
                            if let Some(held) = cf_rule.as_mut() { held.visual_xml = Some(xml[start..reader.buffer_position() as usize].to_string()); }
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
                            style_font: current_row_font.take(),
                            thick_top: current_row_thick_top,
                            thick_bottom: current_row_thick_bottom,
                            hidden: current_row_hidden,
                        });
                        current_row_height = None;
                        current_row_custom = false;
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
                            let through_end=&xml[start..reader.buffer_position() as usize];
                            if let Some(end)=through_end.rfind('<') {
                                let fragment=format!("<sst><si>{}</si></sst>",&through_end[..end]);
                                if let Some(held) = parse_shared_strings(&fragment)?.into_iter().next() {
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
                        if let Some(held) = cf_rule.as_mut() { held.visual_xml = Some(xml[event_start..reader.buffer_position() as usize].to_string()); }
                    }
                    // <sheetProtection sheet="1"> is a sheet that is protected.
                    // The same element without it only records the options of
                    // one that is not, which Excel writes just as readily.
                    "sheetProtection" => {
                        protected = is_true(get_attr(&e, "sheet").as_deref());
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
                            style_font: row_style_font(&e, stylesheet),
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
                        style_font: None,
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
        cached_spills: None,
        pending_used_bounds: None,
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
        protected,
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

/// Parse an .xlsx file without recalculating anything.
///
/// Every formula cell keeps the value Excel last computed for it, alongside its
/// formula text. That pair is what makes a workbook usable as a test oracle.
pub fn parse_xlsx_preserving_values(data: &[u8]) -> Result<Workbook, XlsxError> {
    let mut archive = OoxmlArchive::new(data)?;

    // 1. Parse shared strings (optional 遯ｶ繝ｻsome xlsx files have none)
    let shared_strings = match archive.try_read_part("xl/sharedStrings.xml")? {
        Some(xml) => parse_shared_strings(&xml)?,
        None => Vec::new(),
    };

    // 2. Parse styles.xml (optional 遯ｶ繝ｻsome simple xlsx have none)
    // The theme has to be read first: a style may name one of its colours
    // rather than state one of its own.
    let mut theme = match archive.try_read_part("xl/theme/theme1.xml")? {
        Some(xml) => parse_theme_xml(&xml),
        None => Theme::default(),
    };
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

    // 3. Parse workbook.xml to get sheet names and rIds
    let workbook_xml = archive.read_part("xl/workbook.xml")?;
    let sheet_infos = parse_workbook_sheets(&workbook_xml)?;

    // 4. Parse workbook relationships to map rIds to sheet file paths
    let rels_xml = archive.read_part("xl/_rels/workbook.xml.rels")?;
    let rels = parse_relationships(&rels_xml)?;
    let persons = match rels.values().find(|rel| rel.rel_type.ends_with("/person")) {
        Some(relation) => {
            let path = part_beside("xl/workbook.xml", &relation.target);
            archive.try_read_part(&path)?.map(|xml| parse_person_list(&xml)).unwrap_or_default()
        }
        None => HashMap::new(),
    };

    // Build rId -> target path map
    let sheet_kinds: HashMap<String, crate::ir::SheetKind> = rels.iter().map(|(id,rel)| {
        (id.clone(), if rel.rel_type.ends_with("/chartsheet") { crate::ir::SheetKind::Chart } else { crate::ir::SheetKind::Worksheet })
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
            Some(xml) => external_books.push(parse_external_link(&xml)),
            None => external_books.push(ExternalBook::default()),
        }
    }

    // 5. Parse each worksheet
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
                    for rel in rels.values() {
                        if rel.rel_type.ends_with("/table") {
                            let part = part_beside(&sheet_path, &rel.target);
                            if let Some(table_xml) = archive.try_read_part(&part)? {
                                if let Some(mut table) =
                                    parse_table_xml(&table_xml, &theme, &stylesheet.dxf_borders)
                                {
                                    table.origin_part = Some(part);
                                    sheet.tables.push(table);
                                }
                            }
                        }
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
        theme_colors: theme.colours.clone(),
        indexed_colors: theme.indexed.clone(),
        recalculation: Default::default(),
        calculation: parse_calculation(&workbook_xml),
        default_style: resolve_cell_style(0, &stylesheet),
        defined_names,
        defined_name_attributes,
        compatibility_names: Default::default(),
        external_books,
    })
}

#[cfg(test)]
mod tests {
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

    /// A cell format that names a font and says nothing about applying it is
    /// using that font. Only an explicit `applyFont="0"` sends the reader to
    /// the named style the format is built on 遯ｶ繝ｻwhich is how a hyperlink keeps
    /// its blue underline. Files written by anything other than Excel
    /// routinely leave every flag off, so reading absent as "do not apply"
    /// draws whole workbooks in the wrong face.
    #[test]
    fn a_font_is_applied_unless_the_format_denies_it() {
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
        assert!(denied.underline, "should wear the style's own font");
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
        let workbook = parse_xlsx(include_bytes!(
            "../../../tools/golden-test/documents/xlsx/e1c631fac30a_bunya_taikeizu_point.xlsx"
        ))
        .unwrap();
        assert_eq!(workbook.indexed_colors.len(), 64);
        assert_eq!(workbook.indexed_colors[12], "0000D4");
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
