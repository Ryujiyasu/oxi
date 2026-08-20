// This Source Code Form is subject to the terms of the Mozilla Public
// License, v. 2.0. If a copy of the MPL was not distributed with this
// file, You can obtain one at https://mozilla.org/MPL/2.0/.

use std::collections::HashMap;

use quick_xml::events::Event;
use quick_xml::reader::Reader;
use thiserror::Error;

use oxidocs_common::archive::OoxmlArchive;
use oxidocs_common::relationships::parse_relationships;

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
/// saved, so nothing is lost — it simply does not reach the IR.
fn note_unsupported(name: &str, noted: &mut Vec<String>) {
    let feature = match name {
        "conditionalFormatting" => "Conditional formatting",
        "dataValidation" | "dataValidations" => "Data validation",
        "hyperlink" | "hyperlinks" => "Hyperlinks",
        "pane" => "Frozen panes",
        "sheetProtection" => "Sheet protection",
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
use oxidocs_common::xml_utils::{get_attr, local_name};

use crate::ir::{BorderLine, Cell, CellStyle, CellValue, MergeCell, Row, Sheet, Workbook};

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
fn parse_shared_strings(xml: &str) -> Result<Vec<String>, XlsxError> {
    let mut reader = Reader::from_str(xml);
    let mut strings = Vec::new();
    let mut current_string = String::new();
    let mut in_si = false;
    let mut in_t = false;
    // A phonetic guide (furigana) is stored as an <rPh> element containing its
    // own <t>. It is not part of the cell's text: Excel shows "区分", not
    // "区分クブン". Japanese workbooks carry these on names and addresses
    // constantly, so failing to skip them corrupts a large share of real files.
    let mut in_phonetic = false;

    loop {
        match reader.read_event()? {
            Event::Start(e) => {
                let name = local_name(e.name().as_ref());
                match name.as_str() {
                    "si" => {
                        in_si = true;
                        in_phonetic = false;
                        current_string.clear();
                    }
                    "rPh" => {
                        in_phonetic = true;
                    }
                    "t" if in_si => {
                        in_t = true;
                    }
                    _ => {}
                }
            }
            Event::End(e) => {
                let name = local_name(e.name().as_ref());
                match name.as_str() {
                    "si" => {
                        in_si = false;
                        strings.push(std::mem::take(&mut current_string));
                    }
                    "rPh" => {
                        in_phonetic = false;
                    }
                    "t" => {
                        in_t = false;
                    }
                    _ => {}
                }
            }
            Event::Text(e) => {
                if in_t && in_si && !in_phonetic {
                    let text = e.unescape()?.to_string();
                    current_string.push_str(&text);
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
    r_id: String,
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
                    // r:id attribute — try both namespaced and raw forms
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
    size: Option<f32>,
    color: Option<String>,
    name: Option<String>,
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
}

#[derive(Debug, Clone, Default)]
struct XfRecord {
    num_fmt_id: u32,
    font_id: usize,
    fill_id: usize,
    border_id: usize,
    /// The named style this format is built on, and whether it overrides each
    /// part of it. A cell format that does not apply its own font wears the
    /// font of the style it names — that is how Excel dresses a hyperlink.
    style_id: Option<usize>,
    applies_font: bool,
    applies_fill: bool,
    applies_border: bool,
    applies_number_format: bool,
    horizontal_align: Option<String>,
    vertical_align: Option<String>,
    wrap_text: bool,
}

#[derive(Debug, Clone, Default)]
struct StyleSheet {
    num_fmts: HashMap<u32, String>,
    fonts: Vec<FontInfo>,
    fills: Vec<FillInfo>,
    borders: Vec<BorderInfo>,
    cell_xfs: Vec<XfRecord>,
    cell_style_xfs: Vec<XfRecord>,
}

/// Built-in number format strings for well-known IDs.
fn builtin_number_format(id: u32) -> Option<&'static str> {
    match id {
        0 => Some("General"),
        1 => Some("0"),
        2 => Some("0.00"),
        3 => Some("#,##0"),
        4 => Some("#,##0.00"),
        9 => Some("0%"),
        10 => Some("0.00%"),
        11 => Some("0.00E+00"),
        14 => Some("mm-dd-yy"),
        22 => Some("m/d/yy h:mm"),
        _ => None,
    }
}

/// The colours a workbook's theme names, in the order the theme states them:
/// dk1, lt1, dk2, lt2, accent1-6, hlink, folHlink.
#[derive(Debug, Clone, Default)]
struct Theme {
    colours: Vec<String>,
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

fn parse_theme_xml(xml: &str) -> Theme {
    let mut reader = Reader::from_str(xml);
    let mut theme = Theme::default();
    let mut in_scheme = false;
    let mut buf = Vec::new();
    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(e)) => {
                if local_name(e.name().as_ref()) == "clrScheme" {
                    in_scheme = true;
                }
            }
            Ok(Event::End(e)) => {
                if local_name(e.name().as_ref()) == "clrScheme" {
                    break;
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
                if let Some(colour) = colour {
                    theme.colours.push(colour);
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

fn parse_color_attr(e: &quick_xml::events::BytesStart, theme: &Theme) -> Option<String> {
    if let Some(rgb) = get_attr(e, "rgb") {
        // Strip leading alpha if 8-char hex
        let hex = if rgb.len() == 8 { &rgb[2..] } else { &rgb };
        return Some(hex.to_string());
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
    }
    let mut section = Section::None;
    let mut in_font = false;
    let mut current_font = FontInfo::default();
    let mut in_fill = false;
    let mut current_fill = FillInfo::default();
    let mut in_border = false;
    let mut current_border = BorderInfo::default();
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
                    "border" if section == Section::Borders => {
                        in_border = true;
                        current_border = BorderInfo::default();
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
                            vertical_align: None,
                            wrap_text: false,
                            style_id: get_attr(&e, "xfId").and_then(|v| v.parse().ok()),
                            applies_font: is_true(get_attr(&e, "applyFont").as_deref()),
                            applies_fill: is_true(get_attr(&e, "applyFill").as_deref()),
                            applies_border: is_true(get_attr(&e, "applyBorder").as_deref()),
                            applies_number_format: is_true(
                                get_attr(&e, "applyNumberFormat").as_deref(),
                            ),
                        };
                    }
                    "alignment" if in_xf => {
                        current_xf.horizontal_align = get_attr(&e, "horizontal");
                        current_xf.vertical_align = get_attr(&e, "vertical");
                        current_xf.wrap_text =
                            matches!(get_attr(&e, "wrapText").as_deref(), Some("1") | Some("true"));
                    }

                    // Inside a border element, parse child elements with style attr
                    "left" if in_border => {
                        current_border.left = border_line(&e);
                    }
                    "right" if in_border => {
                        current_border.right = border_line(&e);
                    }
                    "top" if in_border => {
                        current_border.top = border_line(&e);
                    }
                    "bottom" if in_border => {
                        current_border.bottom = border_line(&e);
                    }

                    // Font color
                    "color" if in_font => {
                        if let Some(c) = parse_color_attr(&e, theme) {
                            current_font.color = Some(c);
                        }
                    }

                    // Fill color — look for fgColor inside patternFill
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
                    "sz" if in_font => {
                        current_font.size =
                            get_attr(&e, "val").and_then(|v| v.parse().ok());
                    }
                    "name" | "rFont" if in_font => {
                        current_font.name = get_attr(&e, "val");
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
                    "left" if in_border => {
                        current_border.left = border_line(&e);
                    }
                    "right" if in_border => {
                        current_border.right = border_line(&e);
                    }
                    "top" if in_border => {
                        current_border.top = border_line(&e);
                    }
                    "bottom" if in_border => {
                        current_border.bottom = border_line(&e);
                    }
                    "alignment" if in_xf => {
                        current_xf.horizontal_align = get_attr(&e, "horizontal");
                        current_xf.vertical_align = get_attr(&e, "vertical");
                        current_xf.wrap_text =
                            matches!(get_attr(&e, "wrapText").as_deref(), Some("1") | Some("true"));
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
                            vertical_align: None,
                            wrap_text: false,
                            style_id: get_attr(&e, "xfId").and_then(|v| v.parse().ok()),
                            applies_font: is_true(get_attr(&e, "applyFont").as_deref()),
                            applies_fill: is_true(get_attr(&e, "applyFill").as_deref()),
                            applies_border: is_true(get_attr(&e, "applyBorder").as_deref()),
                            applies_number_format: is_true(
                                get_attr(&e, "applyNumberFormat").as_deref(),
                            ),
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
                        ss.fonts.push(std::mem::take(&mut current_font));
                        in_font = false;
                    }
                    "fill" if in_fill => {
                        ss.fills.push(std::mem::take(&mut current_fill));
                        in_fill = false;
                    }
                    "border" if in_border => {
                        ss.borders.push(std::mem::take(&mut current_border));
                        in_border = false;
                    }
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

    Ok(ss)
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

    let font = stylesheet.fonts.get(font_id).cloned().unwrap_or_default();
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
        None // General — no explicit format needed
    } else if let Some(custom) = stylesheet.num_fmts.get(&num_fmt_id) {
        Some(custom.clone())
    } else {
        builtin_number_format(num_fmt_id).map(|s| s.to_string())
    };

    CellStyle {
        bold: font.bold,
        italic: font.italic,
        underline: font.underline,
        font_size: font.size,
        font_name: font.name.clone(),
        font_color: font.color,
        bg_color: fill.bg_color,
        number_format,
        horizontal_align: xf.horizontal_align.clone(),
        vertical_align: xf.vertical_align.clone(),
        wrap_text: xf.wrap_text,
        border_top: border.top.clone(),
        border_bottom: border.bottom.clone(),
        border_left: border.left.clone(),
        border_right: border.right.clone(),
    }
}

/// Parse a single worksheet XML into a Sheet.
fn parse_worksheet(
    xml: &str,
    sheet_name: &str,
    shared_strings: &[String],
    stylesheet: &StyleSheet,
) -> Result<Sheet, XlsxError> {
    let mut reader = Reader::from_str(xml);
    let mut rows: Vec<Row> = Vec::new();
    let mut max_col: u32 = 0;

    // Column widths: index is 0-based col number
    let mut col_widths: Vec<f32> = Vec::new();
    let mut hidden_cols: Vec<u32> = Vec::new();
    let mut auto_filter: Option<crate::ir::AutoFilter> = None;
    let mut declared_range: Option<(u32, u32, u32, u32)> = None;
    let mut unsupported: Vec<String> = Vec::new();
    let mut filter_field: Option<u32> = None;
    let mut filter_criteria: Vec<String> = Vec::new();
    let mut filter_either = false;
    let mut default_col_width: f32 = 8.43;
    let mut default_row_height: f32 = 15.0;
    let mut merge_cells: Vec<MergeCell> = Vec::new();

    // State tracking
    let mut current_row_index: u32 = 0;
    let mut current_row_height: Option<f32> = None;
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

    // Section tracking
    let mut in_merge_cells = false;

    loop {
        match reader.read_event()? {
            Event::Start(e) => {
                let name = local_name(e.name().as_ref());
                note_unsupported(&name, &mut unsupported);
                match name.as_str() {
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

                        // A row states its height whenever it is not the
                        // default one, whether a person set it or Excel worked
                        // it out from what the row holds. customHeight only
                        // says which of the two it was, so it does not decide
                        // whether the height counts.
                        current_row_height =
                            get_attr(&e, "ht").and_then(|v| v.parse::<f32>().ok());
                        current_row_hidden = is_true(get_attr(&e, "hidden").as_deref());
                    }
                    "c" if in_row => {
                        in_cell = true;
                        value_text.clear();
                        formula_text.clear();
                        in_formula = false;
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
                    }
                    "v" if in_cell => {
                        in_value = true;
                        value_text.clear();
                    }
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
                            hidden: current_row_hidden,
                        });
                        current_row_height = None;
                        current_row_hidden = false;
                    }
                    "c" => {
                        if in_cell {
                            let cell_value =
                                resolve_cell_value(&value_text, &cell_type, shared_strings);
                            // A cell without an s wears the workbook's default
                            // format, which is cellXfs[0] — not no format at all.
                            let style =
                                resolve_cell_style(cell_style_index.unwrap_or(0), stylesheet);
                            let formula = if formula_text.is_empty() {
                                None
                            } else {
                                Some(formula_text.clone())
                            };
                            current_cells.push(Cell {
                                col: cell_col,
                                value: cell_value,
                                style,
                                formula,
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
                    "mergeCells" => {
                        in_merge_cells = false;
                    }
                    _ => {}
                }
            }
            Event::Text(e) => {
                if in_formula {
                    let text = e.unescape()?.to_string();
                    formula_text.push_str(&text);
                } else if in_value {
                    let text = e.unescape()?.to_string();
                    value_text.push_str(&text);
                }
            }
            Event::Empty(e) => {
                let name = local_name(e.name().as_ref());
                note_unsupported(&name, &mut unsupported);
                match name.as_str() {
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
                            col,
                            value: CellValue::Empty,
                            style,
                            formula: None,
                        });
                    }

                    // <sheetFormatPr defaultRowHeight="15" defaultColWidth="8.43" ... />
                    "sheetFormatPr" => {
                        if let Some(v) = get_attr(&e, "defaultRowHeight") {
                            if let Ok(h) = v.parse::<f32>() {
                                default_row_height = h;
                            }
                        }
                        if let Some(v) = get_attr(&e, "defaultColWidth") {
                            if let Ok(w) = v.parse::<f32>() {
                                default_col_width = w;
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
                            col_widths.resize(needed, default_col_width);
                        }
                        for c in min_col..=max_col_attr {
                            col_widths[(c - 1) as usize] = width;
                        }
                        if is_true(get_attr(&e, "hidden").as_deref()) {
                            for c in min_col..=max_col_attr {
                                hidden_cols.push(c - 1);
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
                        let mut rh: Option<f32> = None;
                        let custom_height = get_attr(&e, "customHeight");
                        if custom_height.as_deref() == Some("1") || custom_height.as_deref() == Some("true") {
                            rh = get_attr(&e, "ht").and_then(|v| v.parse::<f32>().ok());
                        }
                        rows.push(Row {
                            index: row_num,
                            cells: Vec::new(),
                            height: rh,
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
        col_widths.resize(col_count, default_col_width);
    }

    Ok(Sheet {
        name: sheet_name.to_string(),
        rows,
        col_count,
        col_widths,
        default_col_width,
        default_row_height,
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
    })
}

/// Resolve a cell's raw value text + type attribute into a CellValue.
fn resolve_cell_value(
    value_text: &str,
    cell_type: &Option<String>,
    shared_strings: &[String],
) -> CellValue {
    if value_text.is_empty() && cell_type.is_none() {
        return CellValue::Empty;
    }

    match cell_type.as_deref() {
        Some("s") => {
            // Shared string index
            if let Ok(idx) = value_text.parse::<usize>() {
                if idx < shared_strings.len() {
                    CellValue::String(shared_strings[idx].clone())
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
            CellValue::String(value_text.to_string())
        }
        _ => {
            // No type attribute means number
            if value_text.is_empty() {
                CellValue::Empty
            } else if let Ok(n) = value_text.parse::<f64>() {
                CellValue::Number(n)
            } else {
                CellValue::String(value_text.to_string())
            }
        }
    }
}

/// Parse an .xlsx file from raw bytes into a Workbook IR.
///
/// The values Excel cached in the file are **kept**. For display they are the
/// correct answer — recalculating a freshly loaded workbook can only introduce
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

    // 1. Parse shared strings (optional — some xlsx files have none)
    let shared_strings = match archive.try_read_part("xl/sharedStrings.xml")? {
        Some(xml) => parse_shared_strings(&xml)?,
        None => Vec::new(),
    };

    // 2. Parse styles.xml (optional — some simple xlsx have none)
    // The theme has to be read first: a style may name one of its colours
    // rather than state one of its own.
    let theme = match archive.try_read_part("xl/theme/theme1.xml")? {
        Some(xml) => parse_theme_xml(&xml),
        None => Theme::default(),
    };
    let stylesheet = match archive.try_read_part("xl/styles.xml")? {
        Some(xml) => parse_styles_xml(&xml, &theme)?,
        None => StyleSheet::default(),
    };

    // 3. Parse workbook.xml to get sheet names and rIds
    let workbook_xml = archive.read_part("xl/workbook.xml")?;
    let sheet_infos = parse_workbook_sheets(&workbook_xml)?;

    // 4. Parse workbook relationships to map rIds to sheet file paths
    let rels_xml = archive.read_part("xl/_rels/workbook.xml.rels")?;
    let rels = parse_relationships(&rels_xml)?;

    // Build rId -> target path map
    let rid_to_path: HashMap<String, String> = rels
        .into_iter()
        .map(|(id, rel)| (id, rel.target))
        .collect();

    // 5. Parse each worksheet
    let mut sheets = Vec::new();
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
                let sheet =
                    parse_worksheet(&sheet_xml, &info.name, &shared_strings, &stylesheet)?;
                sheets.push(sheet);
            }
            None => {
                log::warn!("Sheet file '{}' not found in archive, skipping", sheet_path);
            }
        }
    }

    Ok(Workbook {
        sheets,
        default_style: resolve_cell_style(0, &stylesheet),
    })
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_parse_cell_ref_simple() {
        assert_eq!(parse_cell_ref("A1"), (0, 0));
        assert_eq!(parse_cell_ref("B2"), (1, 1));
        assert_eq!(parse_cell_ref("Z1"), (25, 0));
    }

    /// Furigana lives in `<rPh>` and is not part of the cell's text. Excel shows
    /// "区分"; appending the reading would give "区分クブン" and silently corrupt
    /// most Japanese workbooks, which carry phonetic guides on names and
    /// addresses as a matter of course.
    #[test]
    fn shared_strings_exclude_phonetic_guides() {
        let xml = r#"<?xml version="1.0"?>
<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="3">
  <si><t>区分</t><rPh sb="0" eb="2"><t>クブン</t></rPh><phoneticPr fontId="1"/></si>
  <si><r><t>山田</t></r><r><t>太郎</t></r><rPh sb="0" eb="2"><t>ヤマダ</t></rPh><rPh sb="2" eb="4"><t>タロウ</t></rPh></si>
  <si><t>plain</t></si>
</sst>"#;
        let strings = parse_shared_strings(xml).expect("should parse");
        assert_eq!(strings, vec!["区分", "山田太郎", "plain"]);
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
        let sst: Vec<String> = vec![];
        assert!(matches!(
            resolve_cell_value("42", &None, &sst),
            CellValue::Number(n) if (n - 42.0).abs() < f64::EPSILON
        ));
    }

    #[test]
    fn test_resolve_cell_value_shared_string() {
        let sst = vec!["Hello".to_string(), "World".to_string()];
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
        let sst: Vec<String> = vec![];
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
        let sst: Vec<String> = vec![];
        let t = Some("e".to_string());
        assert!(matches!(
            resolve_cell_value("#REF!", &t, &sst),
            CellValue::Error(ref s) if s == "#REF!"
        ));
    }

    #[test]
    fn test_resolve_cell_value_empty() {
        let sst: Vec<String> = vec![];
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
        assert_eq!(result[0], "Hello");
        assert_eq!(result[1], "World");
        assert_eq!(result[2], "Rich Text");
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
