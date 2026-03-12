use std::collections::HashMap;

use quick_xml::events::Event;
use quick_xml::reader::Reader;
use thiserror::Error;

use oxi_common::archive::OoxmlArchive;
use oxi_common::relationships::parse_relationships;
use oxi_common::xml_utils::{get_attr, local_name};

use crate::ir::{Cell, CellStyle, CellValue, MergeCell, Row, Sheet, Workbook};

#[derive(Error, Debug)]
pub enum XlsxError {
    #[error("Archive error: {0}")]
    Archive(#[from] oxi_common::OxiError),

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

    loop {
        match reader.read_event()? {
            Event::Start(e) => {
                let name = local_name(e.name().as_ref());
                match name.as_str() {
                    "si" => {
                        in_si = true;
                        current_string.clear();
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
                    "t" => {
                        in_t = false;
                    }
                    _ => {}
                }
            }
            Event::Text(e) => {
                if in_t && in_si {
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
    size: Option<f32>,
    color: Option<String>,
}

#[derive(Debug, Clone, Default)]
struct FillInfo {
    bg_color: Option<String>,
}

#[derive(Debug, Clone, Default)]
struct BorderInfo {
    left: bool,
    right: bool,
    top: bool,
    bottom: bool,
}

#[derive(Debug, Clone, Default)]
struct XfRecord {
    num_fmt_id: u32,
    font_id: usize,
    fill_id: usize,
    border_id: usize,
    horizontal_align: Option<String>,
}

#[derive(Debug, Clone, Default)]
struct StyleSheet {
    num_fmts: HashMap<u32, String>,
    fonts: Vec<FontInfo>,
    fills: Vec<FillInfo>,
    borders: Vec<BorderInfo>,
    cell_xfs: Vec<XfRecord>,
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

fn parse_color_attr(e: &quick_xml::events::BytesStart) -> Option<String> {
    // Try "rgb" first (e.g., "FFFF0000"), then "theme" (ignored for simplicity),
    // then "indexed" (ignored).
    if let Some(rgb) = get_attr(e, "rgb") {
        // Strip leading alpha if 8-char hex
        let hex = if rgb.len() == 8 { &rgb[2..] } else { &rgb };
        return Some(hex.to_string());
    }
    None
}

fn parse_styles_xml(xml: &str) -> Result<StyleSheet, XlsxError> {
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
                    "xf" if section == Section::CellXfs => {
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
                        };
                    }
                    "alignment" if in_xf => {
                        current_xf.horizontal_align = get_attr(&e, "horizontal");
                    }

                    // Inside a border element, parse child elements with style attr
                    "left" if in_border => {
                        if get_attr(&e, "style").is_some() {
                            current_border.left = true;
                        }
                    }
                    "right" if in_border => {
                        if get_attr(&e, "style").is_some() {
                            current_border.right = true;
                        }
                    }
                    "top" if in_border => {
                        if get_attr(&e, "style").is_some() {
                            current_border.top = true;
                        }
                    }
                    "bottom" if in_border => {
                        if get_attr(&e, "style").is_some() {
                            current_border.bottom = true;
                        }
                    }

                    // Font color
                    "color" if in_font => {
                        if let Some(c) = parse_color_attr(&e) {
                            current_font.color = Some(c);
                        }
                    }

                    // Fill color — look for fgColor inside patternFill
                    "fgColor" if in_fill => {
                        if let Some(c) = parse_color_attr(&e) {
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
                    "sz" if in_font => {
                        current_font.size =
                            get_attr(&e, "val").and_then(|v| v.parse().ok());
                    }
                    "color" if in_font => {
                        if let Some(c) = parse_color_attr(&e) {
                            current_font.color = Some(c);
                        }
                    }
                    "fgColor" if in_fill => {
                        if let Some(c) = parse_color_attr(&e) {
                            current_fill.bg_color = Some(c);
                        }
                    }
                    // Self-closing border sides: <left style="thin"/>
                    "left" if in_border => {
                        if get_attr(&e, "style").is_some() {
                            current_border.left = true;
                        }
                    }
                    "right" if in_border => {
                        if get_attr(&e, "style").is_some() {
                            current_border.right = true;
                        }
                    }
                    "top" if in_border => {
                        if get_attr(&e, "style").is_some() {
                            current_border.top = true;
                        }
                    }
                    "bottom" if in_border => {
                        if get_attr(&e, "style").is_some() {
                            current_border.bottom = true;
                        }
                    }
                    "alignment" if in_xf => {
                        current_xf.horizontal_align = get_attr(&e, "horizontal");
                    }
                    "xf" if section == Section::CellXfs => {
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
                        };
                        ss.cell_xfs.push(xf);
                    }

                    _ => {}
                }
            }
            Event::End(e) => {
                let name = local_name(e.name().as_ref());
                match name.as_str() {
                    "numFmts" | "fonts" | "fills" | "borders" | "cellXfs" => {
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
                    "xf" if in_xf && section == Section::CellXfs => {
                        ss.cell_xfs.push(std::mem::take(&mut current_xf));
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

/// Build a CellStyle from a style index referencing the StyleSheet.
fn resolve_cell_style(style_index: usize, stylesheet: &StyleSheet) -> CellStyle {
    let xf = match stylesheet.cell_xfs.get(style_index) {
        Some(xf) => xf,
        None => return CellStyle::default(),
    };

    let font = stylesheet.fonts.get(xf.font_id).cloned().unwrap_or_default();
    let fill = stylesheet.fills.get(xf.fill_id).cloned().unwrap_or_default();
    let border = stylesheet
        .borders
        .get(xf.border_id)
        .cloned()
        .unwrap_or_default();

    // Resolve number format
    let number_format = if xf.num_fmt_id == 0 {
        None // General — no explicit format needed
    } else if let Some(custom) = stylesheet.num_fmts.get(&xf.num_fmt_id) {
        Some(custom.clone())
    } else {
        builtin_number_format(xf.num_fmt_id).map(|s| s.to_string())
    };

    CellStyle {
        bold: font.bold,
        italic: font.italic,
        font_size: font.size,
        font_color: font.color,
        bg_color: fill.bg_color,
        number_format,
        horizontal_align: xf.horizontal_align.clone(),
        border_top: border.top,
        border_bottom: border.bottom,
        border_left: border.left,
        border_right: border.right,
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
    let mut default_col_width: f32 = 8.43;
    let mut default_row_height: f32 = 15.0;
    let mut merge_cells: Vec<MergeCell> = Vec::new();

    // State tracking
    let mut current_row_index: u32 = 0;
    let mut current_row_height: Option<f32> = None;
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
                match name.as_str() {
                    "row" => {
                        in_row = true;
                        current_cells.clear();
                        let row_num = get_attr(&e, "r")
                            .and_then(|v| v.parse::<u32>().ok())
                            .unwrap_or(current_row_index + 1);
                        current_row_index = row_num;

                        // Parse custom row height
                        current_row_height = None;
                        let custom_height = get_attr(&e, "customHeight");
                        if custom_height.as_deref() == Some("1") || custom_height.as_deref() == Some("true") {
                            current_row_height =
                                get_attr(&e, "ht").and_then(|v| v.parse::<f32>().ok());
                        }
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
                    "row" => {
                        in_row = false;
                        rows.push(Row {
                            index: current_row_index,
                            cells: std::mem::take(&mut current_cells),
                            height: current_row_height,
                        });
                        current_row_height = None;
                    }
                    "c" => {
                        if in_cell {
                            let cell_value =
                                resolve_cell_value(&value_text, &cell_type, shared_strings);
                            let style = match cell_style_index {
                                Some(idx) => resolve_cell_style(idx, stylesheet),
                                None => CellStyle::default(),
                            };
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
                        let style = match si {
                            Some(idx) => resolve_cell_style(idx, stylesheet),
                            None => CellStyle::default(),
                        };
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
        unsupported_elements: Vec::new(),
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
pub fn parse_xlsx(data: &[u8]) -> Result<Workbook, XlsxError> {
    let mut archive = OoxmlArchive::new(data)?;

    // 1. Parse shared strings (optional — some xlsx files have none)
    let shared_strings = match archive.try_read_part("xl/sharedStrings.xml")? {
        Some(xml) => parse_shared_strings(&xml)?,
        None => Vec::new(),
    };

    // 2. Parse styles.xml (optional — some simple xlsx have none)
    let stylesheet = match archive.try_read_part("xl/styles.xml")? {
        Some(xml) => parse_styles_xml(&xml)?,
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

    // 6. Evaluate formulas in each sheet
    for sheet in &mut sheets {
        crate::formula::evaluate_sheet_formulas(sheet);
    }

    Ok(Workbook { sheets })
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
        assert_eq!(CellValue::Number(3.14).display(), "3.14");
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
        let ss = parse_styles_xml(xml).unwrap();
        assert_eq!(ss.num_fmts.len(), 1);
        assert_eq!(ss.num_fmts.get(&164).unwrap(), "#,##0.00_ ");
        assert_eq!(ss.fonts.len(), 2);
        assert!(ss.fonts[1].bold);
        assert_eq!(ss.fonts[1].size, Some(14.0));
        assert_eq!(ss.fonts[1].color.as_deref(), Some("FF0000"));
        assert_eq!(ss.fills.len(), 2);
        assert_eq!(ss.fills[1].bg_color.as_deref(), Some("FFFF00"));
        assert_eq!(ss.borders.len(), 2);
        assert!(!ss.borders[0].left);
        assert!(ss.borders[1].left);
        assert!(ss.borders[1].right);
        assert!(ss.borders[1].top);
        assert!(ss.borders[1].bottom);
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
        assert!(style.border_top);
        assert!(style.border_bottom);
        assert!(style.border_left);
        assert!(style.border_right);
    }

    #[test]
    fn test_builtin_number_formats() {
        assert_eq!(builtin_number_format(0), Some("General"));
        assert_eq!(builtin_number_format(3), Some("#,##0"));
        assert_eq!(builtin_number_format(14), Some("mm-dd-yy"));
        assert_eq!(builtin_number_format(99), None);
    }
}
