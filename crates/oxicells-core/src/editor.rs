// This Source Code Form is subject to the terms of the Mozilla Public
// License, v. 2.0. If a copy of the MPL was not distributed with this
// file, You can obtain one at https://mozilla.org/MPL/2.0/.

//! Round-trip xlsx editor.
//!
//! Preserves the original ZIP archive. Patches cell values in worksheet XML,
//! inserting missing cells and rows when necessary. Cells edited to string
//! values are written as inline strings (t="str") to avoid rewriting the
//! shared-string table.

use std::collections::{BTreeMap, HashMap};
use std::io::{Cursor, Read, Write};

use quick_xml::events::{BytesEnd, BytesStart, BytesText, Event};
use quick_xml::reader::Reader;
use quick_xml::writer::Writer;
use zip::write::SimpleFileOptions;
use zip::{ZipArchive, ZipWriter};

use crate::ir::Workbook;
use crate::parser::{parse_xlsx, XlsxError};
use oxidocs_common::archive::OoxmlArchive;
use oxidocs_common::relationships::parse_relationships;
use oxidocs_common::xml_utils::{get_attr, local_name};

/// A cell edit operation.
#[derive(Debug, Clone)]
pub struct CellEdit {
    /// 0-based sheet index
    pub sheet_index: usize,
    /// 1-based row number (as in OOXML)
    pub row: u32,
    /// 0-based column index
    pub col: u32,
    /// New display value (written as inline string)
    pub new_value: String,
}

/// Round-trip xlsx editor.
pub struct XlsxEditor {
    original_data: Vec<u8>,
    workbook: Workbook,
    edits: HashMap<(usize, u32, u32), String>, // (sheet_idx, row, col) -> value
}

/// Convert 0-based column to letter reference (0->A, 25->Z, 26->AA).
pub fn col_to_letter(mut col: u32) -> String {
    let mut result = String::new();
    loop {
        result.insert(0, (b'A' + (col % 26) as u8) as char);
        if col < 26 {
            break;
        }
        col = col / 26 - 1;
    }
    result
}

impl XlsxEditor {
    pub fn new(data: &[u8]) -> Result<Self, XlsxError> {
        let workbook = parse_xlsx(data)?;
        Ok(Self {
            original_data: data.to_vec(),
            workbook,
            edits: HashMap::new(),
        })
    }

    pub fn workbook(&self) -> &Workbook {
        &self.workbook
    }

    pub fn set_cell(&mut self, sheet_index: usize, row: u32, col: u32, value: String) {
        self.edits.insert((sheet_index, row, col), value);
    }

    pub fn apply_edits(&mut self, edits: &[CellEdit]) {
        for e in edits {
            self.set_cell(e.sheet_index, e.row, e.col, e.new_value.clone());
        }
    }

    pub fn has_edits(&self) -> bool {
        !self.edits.is_empty()
    }

    /// Save edited xlsx.
    pub fn save(&self) -> Result<Vec<u8>, XlsxError> {
        if self.edits.is_empty() {
            return Ok(self.original_data.clone());
        }

        // Determine which sheet files need patching
        let sheet_paths = self.resolve_sheet_paths()?;

        let cursor = Cursor::new(&self.original_data);
        let mut archive = ZipArchive::new(cursor)
            .map_err(|e| XlsxError::InvalidData(e.to_string()))?;

        // Group edits by sheet index
        let mut edits_by_sheet: HashMap<usize, HashMap<(u32, u32), &String>> = HashMap::new();
        for ((si, row, col), val) in &self.edits {
            edits_by_sheet
                .entry(*si)
                .or_default()
                .insert((*row, *col), val);
        }

        // Map sheet path -> edits for that sheet
        let mut path_edits: HashMap<String, &HashMap<(u32, u32), &String>> = HashMap::new();
        for (si, edits) in &edits_by_sheet {
            if let Some(path) = sheet_paths.get(*si) {
                path_edits.insert(path.clone(), edits);
            }
        }

        let mut output = Vec::new();
        {
            let mut writer = ZipWriter::new(Cursor::new(&mut output));

            for i in 0..archive.len() {
                let mut entry = archive.by_index(i)
                    .map_err(|e| XlsxError::InvalidData(e.to_string()))?;
                let name = entry.name().to_string();
                let options = SimpleFileOptions::default()
                    .compression_method(entry.compression());

                writer.start_file(&name, options)
                    .map_err(|e| XlsxError::InvalidData(e.to_string()))?;

                if let Some(cell_edits) = path_edits.get(&name) {
                    let mut xml = String::new();
                    entry.read_to_string(&mut xml)
                        .map_err(|e| XlsxError::InvalidData(e.to_string()))?;
                    let patched = patch_worksheet_xml(&xml, cell_edits)?;
                    writer.write_all(patched.as_bytes())
                        .map_err(|e| XlsxError::InvalidData(e.to_string()))?;
                } else {
                    let mut buf = Vec::new();
                    entry.read_to_end(&mut buf)
                        .map_err(|e| XlsxError::InvalidData(e.to_string()))?;
                    writer.write_all(&buf)
                        .map_err(|e| XlsxError::InvalidData(e.to_string()))?;
                }
            }

            writer.finish().map_err(|e| XlsxError::InvalidData(e.to_string()))?;
        }

        Ok(output)
    }

    /// Resolve sheet index -> ZIP path for each sheet.
    fn resolve_sheet_paths(&self) -> Result<Vec<String>, XlsxError> {
        let mut archive = OoxmlArchive::new(&self.original_data)?;
        let workbook_xml = archive.read_part("xl/workbook.xml")?;
        let rels_xml = archive.read_part("xl/_rels/workbook.xml.rels")?;

        // Parse sheet rIds
        let mut reader = Reader::from_str(&workbook_xml);
        let mut r_ids = Vec::new();
        loop {
            match reader.read_event().map_err(XlsxError::Xml)? {
                Event::Start(e) | Event::Empty(e) => {
                    if local_name(e.name().as_ref()) == "sheet" {
                        let r_id = get_attr(&e, "id")
                            .or_else(|| {
                                for attr in e.attributes().flatten() {
                                    let key = std::str::from_utf8(attr.key.as_ref()).unwrap_or("");
                                    if key == "r:id" {
                                        return Some(String::from_utf8_lossy(&attr.value).to_string());
                                    }
                                }
                                None
                            })
                            .unwrap_or_default();
                        r_ids.push(r_id);
                    }
                }
                Event::Eof => break,
                _ => {}
            }
        }

        let rels = parse_relationships(&rels_xml)?;
        let rid_to_path: HashMap<String, String> = rels
            .into_iter()
            .map(|(id, rel)| (id, rel.target))
            .collect();

        let mut paths = Vec::new();
        for r_id in &r_ids {
            if let Some(target) = rid_to_path.get(r_id) {
                let path = oxidocs_common::security::sanitize_rel_target("xl", target)
                    .unwrap_or_default();
                paths.push(path);
            } else {
                paths.push(String::new());
            }
        }

        Ok(paths)
    }
}

fn write_inline_string_cell(
    writer: &mut Writer<Cursor<Vec<u8>>>,
    row: u32,
    col: u32,
    value: &str,
) -> Result<(), XlsxError> {
    let reference = format!("{}{row}", col_to_letter(col));
    let mut cell = BytesStart::new("c");
    cell.push_attribute(("r", reference.as_str()));
    cell.push_attribute(("t", "str"));
    writer
        .write_event(Event::Start(cell))
        .map_err(|error| XlsxError::InvalidData(error.to_string()))?;
    writer
        .write_event(Event::Start(BytesStart::new("v")))
        .map_err(|error| XlsxError::InvalidData(error.to_string()))?;
    writer
        .write_event(Event::Text(BytesText::new(value)))
        .map_err(|error| XlsxError::InvalidData(error.to_string()))?;
    writer
        .write_event(Event::End(BytesEnd::new("v")))
        .map_err(|error| XlsxError::InvalidData(error.to_string()))?;
    writer
        .write_event(Event::End(BytesEnd::new("c")))
        .map_err(|error| XlsxError::InvalidData(error.to_string()))
}

fn write_inserted_row(
    writer: &mut Writer<Cursor<Vec<u8>>>,
    row: u32,
    cells: BTreeMap<u32, String>,
) -> Result<(), XlsxError> {
    let row_text = row.to_string();
    let mut row_start = BytesStart::new("row");
    row_start.push_attribute(("r", row_text.as_str()));
    writer
        .write_event(Event::Start(row_start))
        .map_err(|error| XlsxError::InvalidData(error.to_string()))?;
    for (col, value) in cells {
        write_inline_string_cell(writer, row, col, &value)?;
    }
    writer
        .write_event(Event::End(BytesEnd::new("row")))
        .map_err(|error| XlsxError::InvalidData(error.to_string()))
}

/// Patch worksheet XML, replacing or inserting cells at specified positions.
/// Edited cells are written as inline string type (t="str") with `<v>` containing the text.
fn patch_worksheet_xml(
    xml: &str,
    edits: &HashMap<(u32, u32), &String>,
) -> Result<String, XlsxError> {
    let mut reader = Reader::from_str(xml);
    let mut writer = Writer::new(Cursor::new(Vec::new()));
    let mut pending: BTreeMap<u32, BTreeMap<u32, String>> = BTreeMap::new();
    for (&(row, col), value) in edits {
        pending
            .entry(row)
            .or_default()
            .insert(col, (*value).clone());
    }

    let mut current_row: u32 = 0;
    let mut in_row = false;
    let mut in_cell = false;
    let mut cell_col: u32;
    let mut cell_row: u32;
    let mut in_value = false;
    let mut current_edit: Option<String> = None;
    let mut skip_value_text = false;
    let mut skip_replaced_content_depth = 0_u32;

    loop {
        match reader.read_event().map_err(XlsxError::Xml)? {
            Event::Eof => break,
            Event::Start(ref e) => {
                if skip_replaced_content_depth > 0 {
                    skip_replaced_content_depth += 1;
                    continue;
                }
                let name = local_name(e.name().as_ref());
                if in_cell && current_edit.is_some() && (name == "f" || name == "is") {
                    skip_replaced_content_depth = 1;
                    continue;
                }
                match name.as_str() {
                    "row" => {
                        let next_row = get_attr(e, "r")
                            .and_then(|value| value.parse().ok())
                            .unwrap_or(0);
                        let missing_rows: Vec<u32> =
                            pending.range(..next_row).map(|(&row, _)| row).collect();
                        for row in missing_rows {
                            if let Some(cells) = pending.remove(&row) {
                                write_inserted_row(&mut writer, row, cells)?;
                            }
                        }
                        in_row = true;
                        current_row = next_row;
                    }
                    "c" if in_row => {
                        in_cell = true;
                        let cell_ref = get_attr(e, "r").unwrap_or_default();
                        let (col, row) = crate::parser::parse_cell_ref(&cell_ref);
                        cell_col = col;
                        cell_row = if row > 0 { row + 1 } else { current_row };

                        if let Some(cells) = pending.get_mut(&cell_row) {
                            let missing_cols: Vec<u32> =
                                cells.range(..cell_col).map(|(&col, _)| col).collect();
                            for col in missing_cols {
                                if let Some(value) = cells.remove(&col) {
                                    write_inline_string_cell(
                                        &mut writer,
                                        cell_row,
                                        col,
                                        &value,
                                    )?;
                                }
                            }
                            current_edit = cells.remove(&cell_col);
                        } else {
                            current_edit = None;
                        }

                        if current_edit.is_some() {
                            // Rewrite the <c> element with t="str"
                            let mut new_start = BytesStart::new("c");
                            for attr in e.attributes().flatten() {
                                let key = std::str::from_utf8(attr.key.as_ref()).unwrap_or("");
                                if key == "t" {
                                    continue; // we'll set our own type
                                }
                                new_start.push_attribute((key, std::str::from_utf8(&attr.value).unwrap_or("")));
                            }
                            new_start.push_attribute(("t", "str"));
                            writer.write_event(Event::Start(new_start)).map_err(|e| XlsxError::InvalidData(e.to_string()))?;
                        } else {
                            writer.write_event(Event::Start(e.clone())).map_err(|e| XlsxError::InvalidData(e.to_string()))?;
                        }
                        continue;
                    }
                    "v" if in_cell && current_edit.is_some() => {
                        in_value = true;
                        skip_value_text = true;
                        // Write <v> start
                        writer.write_event(Event::Start(e.clone())).map_err(|e| XlsxError::InvalidData(e.to_string()))?;
                        // Write new value text
                        if let Some(val) = &current_edit {
                            writer.write_event(Event::Text(BytesText::new(val))).map_err(|e| XlsxError::InvalidData(e.to_string()))?;
                        }
                        continue;
                    }
                    _ => {}
                }
                writer.write_event(Event::Start(e.clone())).map_err(|e| XlsxError::InvalidData(e.to_string()))?;
            }
            Event::End(ref e) => {
                if skip_replaced_content_depth > 0 {
                    skip_replaced_content_depth -= 1;
                    continue;
                }
                let name = local_name(e.name().as_ref());
                match name.as_str() {
                    "row" => {
                        if let Some(cells) = pending.remove(&current_row) {
                            for (col, value) in cells {
                                write_inline_string_cell(
                                    &mut writer,
                                    current_row,
                                    col,
                                    &value,
                                )?;
                            }
                        }
                        in_row = false;
                    }
                    "c" => {
                        if in_cell && current_edit.is_some() {
                            // If the original cell had no <v>, we need to add one
                            if !in_value {
                                if let Some(val) = &current_edit {
                                    writer.write_event(Event::Start(BytesStart::new("v"))).map_err(|e| XlsxError::InvalidData(e.to_string()))?;
                                    writer.write_event(Event::Text(BytesText::new(val))).map_err(|e| XlsxError::InvalidData(e.to_string()))?;
                                    writer.write_event(Event::End(BytesEnd::new("v"))).map_err(|e| XlsxError::InvalidData(e.to_string()))?;
                                }
                            }
                        }
                        in_cell = false;
                        in_value = false;
                        current_edit = None;
                        skip_value_text = false;
                    }
                    "v" => {
                        in_value = false;
                        skip_value_text = false;
                    }
                    "sheetData" => {
                        let remaining = std::mem::take(&mut pending);
                        for (row, cells) in remaining {
                            write_inserted_row(&mut writer, row, cells)?;
                        }
                    }
                    _ => {}
                }
                writer.write_event(Event::End(e.clone())).map_err(|e| XlsxError::InvalidData(e.to_string()))?;
            }
            Event::Text(ref e) => {
                if skip_replaced_content_depth > 0 || (skip_value_text && in_value) {
                    // Already wrote the new value, skip the original
                    continue;
                }
                writer.write_event(Event::Text(e.clone())).map_err(|e| XlsxError::InvalidData(e.to_string()))?;
            }
            Event::Empty(ref e) => {
                if skip_replaced_content_depth > 0 {
                    continue;
                }
                let name = local_name(e.name().as_ref());
                if in_cell && current_edit.is_some() && (name == "f" || name == "is") {
                    continue;
                }
                if name == "sheetData" && !pending.is_empty() {
                    writer
                        .write_event(Event::Start(e.clone()))
                        .map_err(|error| XlsxError::InvalidData(error.to_string()))?;
                    let remaining = std::mem::take(&mut pending);
                    for (row, cells) in remaining {
                        write_inserted_row(&mut writer, row, cells)?;
                    }
                    writer
                        .write_event(Event::End(BytesEnd::new("sheetData")))
                        .map_err(|error| XlsxError::InvalidData(error.to_string()))?;
                    continue;
                }
                if name == "row" {
                    let row_num = get_attr(e, "r")
                        .and_then(|value| value.parse().ok())
                        .unwrap_or(0);
                    let missing_rows: Vec<u32> =
                        pending.range(..row_num).map(|(&row, _)| row).collect();
                    for row in missing_rows {
                        if let Some(cells) = pending.remove(&row) {
                            write_inserted_row(&mut writer, row, cells)?;
                        }
                    }
                    if let Some(cells) = pending.remove(&row_num) {
                        writer.write_event(Event::Start(e.clone())).map_err(|error| {
                            XlsxError::InvalidData(error.to_string())
                        })?;
                        for (col, value) in cells {
                            write_inline_string_cell(&mut writer, row_num, col, &value)?;
                        }
                        writer.write_event(Event::End(BytesEnd::new("row"))).map_err(
                            |error| XlsxError::InvalidData(error.to_string()),
                        )?;
                        continue;
                    }
                }
                if name == "c" && in_row {
                    let cell_ref = get_attr(e, "r").unwrap_or_default();
                    let (col, row) = crate::parser::parse_cell_ref(&cell_ref);
                    let row_num = if row > 0 { row + 1 } else { current_row };

                    let mut edit = None;
                    if let Some(cells) = pending.get_mut(&row_num) {
                        let missing_cols: Vec<u32> =
                            cells.range(..col).map(|(&column, _)| column).collect();
                        for missing_col in missing_cols {
                            if let Some(value) = cells.remove(&missing_col) {
                                write_inline_string_cell(
                                    &mut writer,
                                    row_num,
                                    missing_col,
                                    &value,
                                )?;
                            }
                        }
                        edit = cells.remove(&col);
                    }

                    if let Some(val) = edit {
                        // Convert empty cell to cell with value
                        let mut new_start = BytesStart::new("c");
                        for attr in e.attributes().flatten() {
                            let key = std::str::from_utf8(attr.key.as_ref()).unwrap_or("");
                            if key == "t" { continue; }
                            new_start.push_attribute((key, std::str::from_utf8(&attr.value).unwrap_or("")));
                        }
                        new_start.push_attribute(("t", "str"));
                        writer.write_event(Event::Start(new_start)).map_err(|e| XlsxError::InvalidData(e.to_string()))?;
                        writer.write_event(Event::Start(BytesStart::new("v"))).map_err(|e| XlsxError::InvalidData(e.to_string()))?;
                        writer.write_event(Event::Text(BytesText::new(&val))).map_err(|e| XlsxError::InvalidData(e.to_string()))?;
                        writer.write_event(Event::End(BytesEnd::new("v"))).map_err(|e| XlsxError::InvalidData(e.to_string()))?;
                        writer.write_event(Event::End(BytesEnd::new("c"))).map_err(|e| XlsxError::InvalidData(e.to_string()))?;
                        continue;
                    }
                }
                writer.write_event(Event::Empty(e.clone())).map_err(|e| XlsxError::InvalidData(e.to_string()))?;
            }
            event => {
                if skip_replaced_content_depth > 0 {
                    continue;
                }
                writer.write_event(event).map_err(|e| XlsxError::InvalidData(e.to_string()))?;
            }
        }
    }

    let result = writer.into_inner().into_inner();
    String::from_utf8(result).map_err(|_| XlsxError::InvalidData("UTF-8 error".to_string()))
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_col_to_letter() {
        assert_eq!(col_to_letter(0), "A");
        assert_eq!(col_to_letter(1), "B");
        assert_eq!(col_to_letter(25), "Z");
        assert_eq!(col_to_letter(26), "AA");
        assert_eq!(col_to_letter(27), "AB");
    }

    #[test]
    fn test_editor_round_trip() {
        let data = include_bytes!("../../../tests/fixtures/basic_test.xlsx");
        let editor = XlsxEditor::new(data).expect("should open");
        let saved = editor.save().expect("should save");
        let wb = parse_xlsx(&saved).expect("should parse");
        assert_eq!(wb.sheets.len(), 1);
        assert_eq!(wb.sheets[0].name, "Sales");
    }

    #[test]
    fn test_editor_change_cell() {
        let data = include_bytes!("../../../tests/fixtures/basic_test.xlsx");
        let mut editor = XlsxEditor::new(data).expect("should open");

        // Change cell A1 (row 1, col 0) — "Product" header
        editor.set_cell(0, 1, 0, "Item".to_string());

        let saved = editor.save().expect("should save");
        let wb = parse_xlsx(&saved).expect("should parse");

        let row1 = &wb.sheets[0].rows[0];
        assert!(matches!(&row1.cells[0].value, crate::ir::CellValue::String(s) if s == "Item"));
    }

    #[test]
    fn test_editor_inserts_missing_cells_and_rows() {
        let data = include_bytes!("../../../tests/fixtures/basic_test.xlsx");
        let mut editor = XlsxEditor::new(data).expect("should open");

        editor.set_cell(0, 1, 25, "New column".to_string());
        editor.set_cell(0, 4, 1, "New row".to_string());

        let saved = editor.save().expect("should save");
        let wb = parse_xlsx(&saved).expect("should parse");
        let sheet = &wb.sheets[0];
        let inserted_cell = sheet.rows[0]
            .cells
            .iter()
            .find(|cell| cell.col == 25)
            .expect("Z1 should be inserted");
        assert!(matches!(
            &inserted_cell.value,
            crate::ir::CellValue::String(value) if value == "New column"
        ));
        let inserted_row = sheet
            .rows
            .iter()
            .find(|row| row.index == 4)
            .expect("row 4 should be inserted");
        assert!(matches!(
            &inserted_row.cells[0].value,
            crate::ir::CellValue::String(value) if value == "New row"
        ));
    }

    #[test]
    fn worksheet_patch_orders_insertions_and_replaces_formula_content() {
        let xml = r#"<worksheet><sheetData><row r="1"><c r="A1"><f>1+1</f><v>2</v></c><c r="C1" t="inlineStr"><is><t>old</t></is></c></row><row r="3"/></sheetData></worksheet>"#;
        let formula_value = "9".to_string();
        let middle_value = "middle".to_string();
        let inline_value = "new".to_string();
        let new_row_value = "row two".to_string();
        let edits = HashMap::from([
            ((1, 0), &formula_value),
            ((1, 1), &middle_value),
            ((1, 2), &inline_value),
            ((2, 1), &new_row_value),
        ]);

        let patched = patch_worksheet_xml(xml, &edits).expect("should patch worksheet");
        assert!(!patched.contains("<f>"));
        assert!(!patched.contains("<is>"));
        let a1 = patched.find("r=\"A1\"").unwrap();
        let b1 = patched.find("r=\"B1\"").unwrap();
        let c1 = patched.find("r=\"C1\"").unwrap();
        assert!(a1 < b1 && b1 < c1);
        let row1 = patched.find("<row r=\"1\"").unwrap();
        let row2 = patched.find("<row r=\"2\"").unwrap();
        let row3 = patched.find("<row r=\"3\"").unwrap();
        assert!(row1 < row2 && row2 < row3);
        assert!(patched.contains("<v>9</v>"));
        assert!(patched.contains("<v>middle</v>"));
        assert!(patched.contains("<v>new</v>"));
        assert!(patched.contains("<v>row two</v>"));
    }

    #[test]
    fn worksheet_patch_populates_an_empty_sheet_data_element() {
        let xml = r#"<worksheet><sheetData/></worksheet>"#;
        let value = "first value".to_string();
        let edits = HashMap::from([((1, 0), &value)]);

        let patched = patch_worksheet_xml(xml, &edits).expect("should patch empty worksheet");
        let expected = "<sheetData><row r=\"1\"><c r=\"A1\" t=\"str\">\
                        <v>first value</v></c></row></sheetData>";
        assert!(patched.contains(expected));
    }
}
