use std::collections::HashMap;

use quick_xml::events::Event;
use quick_xml::reader::Reader;
use thiserror::Error;

use oxi_common::archive::OoxmlArchive;
use oxi_common::relationships::parse_relationships;
use oxi_common::xml_utils::{get_attr, local_name};

use crate::ir::{Cell, CellStyle, CellValue, Row, Sheet, Workbook};

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

/// Parse a single worksheet XML into a Sheet.
fn parse_worksheet(
    xml: &str,
    sheet_name: &str,
    shared_strings: &[String],
) -> Result<Sheet, XlsxError> {
    let mut reader = Reader::from_str(xml);
    let mut rows: Vec<Row> = Vec::new();
    let mut max_col: u32 = 0;

    // State tracking
    let mut current_row_index: u32 = 0;
    let mut current_cells: Vec<Cell> = Vec::new();
    let mut in_row = false;

    // Cell state
    let mut cell_col: u32 = 0;
    let mut cell_type: Option<String> = None;
    let mut in_cell = false;
    let mut in_value = false;
    let mut value_text = String::new();

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
                    }
                    "c" if in_row => {
                        in_cell = true;
                        value_text.clear();
                        cell_type = get_attr(&e, "t");
                        let cell_ref = get_attr(&e, "r").unwrap_or_default();
                        let (col, _) = parse_cell_ref(&cell_ref);
                        cell_col = col;
                        if col + 1 > max_col {
                            max_col = col + 1;
                        }
                    }
                    "v" if in_cell => {
                        in_value = true;
                        value_text.clear();
                    }
                    _ => {}
                }
            }
            Event::End(e) => {
                let name = local_name(e.name().as_ref());
                match name.as_str() {
                    "row" => {
                        in_row = false;
                        if !current_cells.is_empty() {
                            rows.push(Row {
                                index: current_row_index,
                                cells: std::mem::take(&mut current_cells),
                            });
                        }
                    }
                    "c" => {
                        if in_cell {
                            let cell_value =
                                resolve_cell_value(&value_text, &cell_type, shared_strings);
                            current_cells.push(Cell {
                                col: cell_col,
                                value: cell_value,
                                style: CellStyle::default(),
                            });
                            in_cell = false;
                            cell_type = None;
                        }
                    }
                    "v" => {
                        in_value = false;
                    }
                    _ => {}
                }
            }
            Event::Text(e) => {
                if in_value {
                    let text = e.unescape()?.to_string();
                    value_text.push_str(&text);
                }
            }
            Event::Empty(e) => {
                let name = local_name(e.name().as_ref());
                // Handle self-closing <c .../> (cell with no value)
                if name == "c" && in_row {
                    let cell_ref = get_attr(&e, "r").unwrap_or_default();
                    let (col, _) = parse_cell_ref(&cell_ref);
                    if col + 1 > max_col {
                        max_col = col + 1;
                    }
                    current_cells.push(Cell {
                        col,
                        value: CellValue::Empty,
                        style: CellStyle::default(),
                    });
                }
            }
            Event::Eof => break,
            _ => {}
        }
    }

    Ok(Sheet {
        name: sheet_name.to_string(),
        rows,
        col_count: max_col as usize,
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

    // 2. Parse workbook.xml to get sheet names and rIds
    let workbook_xml = archive.read_part("xl/workbook.xml")?;
    let sheet_infos = parse_workbook_sheets(&workbook_xml)?;

    // 3. Parse workbook relationships to map rIds to sheet file paths
    let rels_xml = archive.read_part("xl/_rels/workbook.xml.rels")?;
    let rels = parse_relationships(&rels_xml)?;

    // Build rId -> target path map
    let rid_to_path: HashMap<String, String> = rels
        .into_iter()
        .map(|(id, rel)| (id, rel.target))
        .collect();

    // 4. Parse each worksheet
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
                let sheet = parse_worksheet(&sheet_xml, &info.name, &shared_strings)?;
                sheets.push(sheet);
            }
            None => {
                log::warn!("Sheet file '{}' not found in archive, skipping", sheet_path);
            }
        }
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
}
