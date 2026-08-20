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

/// The OOXML value type to write for a cell edit.
#[derive(Debug, Clone)]
pub enum CellEditValue {
    String(String),
    Number(f64),
    Boolean(bool),
    Formula(String),
    Empty,
}

/// Round-trip xlsx editor.
pub struct XlsxEditor {
    original_data: Vec<u8>,
    workbook: Workbook,
    edits: HashMap<(usize, u32, u32), CellEditValue>, // (sheet_idx, row, col) -> value
    /// (sheet_idx, 1-based row) -> hidden
    row_hidden: HashMap<(usize, u32), bool>,
    /// (sheet_idx, 0-based column) -> hidden
    col_hidden: HashMap<(usize, u32), bool>,
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
            row_hidden: HashMap::new(),
            col_hidden: HashMap::new(),
        })
    }

    pub fn workbook(&self) -> &Workbook {
        &self.workbook
    }

    pub fn set_cell(&mut self, sheet_index: usize, row: u32, col: u32, value: String) {
        self.set_cell_value(sheet_index, row, col, CellEditValue::String(value));
    }

    pub fn set_cell_value(
        &mut self,
        sheet_index: usize,
        row: u32,
        col: u32,
        value: CellEditValue,
    ) {
        self.edits.insert((sheet_index, row, col), value);
    }

    pub fn apply_edits(&mut self, edits: &[CellEdit]) {
        for e in edits {
            self.set_cell(e.sheet_index, e.row, e.col, e.new_value.clone());
        }
    }

    /// Works out what changed between the workbook this editor opened and the
    /// one it is handed, and records those changes as edits.
    ///
    /// This is how a VBA run reaches the file: the runtime rewrites the IR in
    /// place, and the difference against the original is what has to be
    /// written back.
    ///
    /// Only what the editor can write is compared — cell values, cell formulas,
    /// and which rows and columns are hidden. Styling and merged cells travel
    /// with the original XML untouched, so a change to those is not saved.
    pub fn apply_workbook(&mut self, edited: &Workbook) -> Result<(), XlsxError> {
        if edited.sheets.len() != self.workbook.sheets.len() {
            return Err(XlsxError::InvalidData(
                "the edited workbook has a different number of sheets".to_string(),
            ));
        }

        for (index, (before, after)) in self
            .workbook
            .sheets
            .iter()
            .zip(&edited.sheets)
            .enumerate()
            .collect::<Vec<_>>()
        {
            let mut changes: Vec<((usize, u32, u32), CellEditValue)> = Vec::new();
            let held = cells_of(before);
            let now = cells_of(after);
            for (&(row, col), cell) in &now {
                let same = held
                    .get(&(row, col))
                    .is_some_and(|before| same_content(before, cell));
                if !same {
                    if let Some(value) = edit_for(cell) {
                        changes.push(((index, row, col), value));
                    }
                }
            }
            // A cell the run emptied has to be written as empty, not left alone.
            for &(row, col) in held.keys() {
                if !now.contains_key(&(row, col)) {
                    changes.push(((index, row, col), CellEditValue::Empty));
                }
            }
            for (key, value) in changes {
                self.edits.insert(key, value);
            }

            let mut rows: Vec<(u32, bool)> = Vec::new();
            for row in &after.rows {
                let was = before
                    .rows
                    .iter()
                    .find(|held| held.index == row.index)
                    .is_some_and(|held| held.hidden);
                if was != row.hidden {
                    rows.push((row.index, row.hidden));
                }
            }
            for row in &before.rows {
                let gone = !after.rows.iter().any(|held| held.index == row.index);
                if gone && row.hidden {
                    rows.push((row.index, false));
                }
            }
            for (row, hidden) in rows {
                self.row_hidden.insert((index, row), hidden);
            }

            let mut cols: Vec<(u32, bool)> = Vec::new();
            for col in &after.hidden_cols {
                if !before.hidden_cols.contains(col) {
                    cols.push((*col, true));
                }
            }
            for col in &before.hidden_cols {
                if !after.hidden_cols.contains(col) {
                    cols.push((*col, false));
                }
            }
            for (col, hidden) in cols {
                self.col_hidden.insert((index, col), hidden);
            }
        }
        Ok(())
    }

    /// Hide or reveal a whole row. `row` is one-based, as OOXML counts them.
    pub fn set_row_hidden(&mut self, sheet_index: usize, row: u32, hidden: bool) {
        self.row_hidden.insert((sheet_index, row), hidden);
    }

    /// Hide or reveal a whole column. `col` is zero-based, as the IR counts it.
    pub fn set_col_hidden(&mut self, sheet_index: usize, col: u32, hidden: bool) {
        self.col_hidden.insert((sheet_index, col), hidden);
    }

    pub fn has_edits(&self) -> bool {
        !self.edits.is_empty() || !self.row_hidden.is_empty() || !self.col_hidden.is_empty()
    }

    /// Save edited xlsx.
    pub fn save(&self) -> Result<Vec<u8>, XlsxError> {
        if !self.has_edits() {
            return Ok(self.original_data.clone());
        }

        // Determine which sheet files need patching
        let sheet_paths = self.resolve_sheet_paths()?;

        let cursor = Cursor::new(&self.original_data);
        let mut archive = ZipArchive::new(cursor)
            .map_err(|e| XlsxError::InvalidData(e.to_string()))?;

        // Group edits by sheet index
        let mut edits_by_sheet: HashMap<usize, HashMap<(u32, u32), &CellEditValue>> =
            HashMap::new();
        for ((si, row, col), val) in &self.edits {
            edits_by_sheet
                .entry(*si)
                .or_default()
                .insert((*row, *col), val);
        }

        let mut rows_by_sheet: HashMap<usize, BTreeMap<u32, bool>> = HashMap::new();
        for ((sheet, row), hidden) in &self.row_hidden {
            rows_by_sheet.entry(*sheet).or_default().insert(*row, *hidden);
        }
        let mut cols_by_sheet: HashMap<usize, BTreeMap<u32, bool>> = HashMap::new();
        for ((sheet, col), hidden) in &self.col_hidden {
            cols_by_sheet.entry(*sheet).or_default().insert(*col, *hidden);
        }

        // Map sheet path -> everything to change in that sheet
        let empty_cells: HashMap<(u32, u32), &CellEditValue> = HashMap::new();
        let empty_lines: BTreeMap<u32, bool> = BTreeMap::new();
        let mut path_edits: HashMap<String, SheetEdits<'_>> = HashMap::new();
        for sheet in edits_by_sheet
            .keys()
            .chain(rows_by_sheet.keys())
            .chain(cols_by_sheet.keys())
            .copied()
            .collect::<std::collections::BTreeSet<_>>()
        {
            if let Some(path) = sheet_paths.get(sheet) {
                path_edits.insert(
                    path.clone(),
                    SheetEdits {
                        cells: edits_by_sheet.get(&sheet).unwrap_or(&empty_cells),
                        rows: rows_by_sheet.get(&sheet).unwrap_or(&empty_lines),
                        cols: cols_by_sheet.get(&sheet).unwrap_or(&empty_lines),
                    },
                );
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

                if let Some(sheet_edits) = path_edits.get(&name) {
                    let mut xml = String::new();
                    entry.read_to_string(&mut xml)
                        .map_err(|e| XlsxError::InvalidData(e.to_string()))?;
                    let patched = patch_worksheet_xml(&xml, sheet_edits)?;
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

impl CellEditValue {
    fn cell_type(&self) -> Option<&'static str> {
        match self {
            Self::String(_) => Some("str"),
            Self::Boolean(_) => Some("b"),
            Self::Number(_) | Self::Formula(_) | Self::Empty => None,
        }
    }

    fn value_text(&self) -> Result<Option<String>, XlsxError> {
        match self {
            Self::String(value) => Ok(Some(value.clone())),
            Self::Number(value) if value.is_finite() => Ok(Some(value.to_string())),
            Self::Number(_) => Err(XlsxError::InvalidData(
                "spreadsheet numbers must be finite".to_string(),
            )),
            Self::Boolean(value) => Ok(Some(if *value { "1" } else { "0" }.to_string())),
            Self::Formula(_) | Self::Empty => Ok(None),
        }
    }

    fn formula_text(&self) -> Option<&str> {
        match self {
            Self::Formula(value) => Some(value.strip_prefix('=').unwrap_or(value)),
            _ => None,
        }
    }
}

fn write_formula(
    writer: &mut Writer<Cursor<Vec<u8>>>,
    formula: &str,
) -> Result<(), XlsxError> {
    writer
        .write_event(Event::Start(BytesStart::new("f")))
        .map_err(|error| XlsxError::InvalidData(error.to_string()))?;
    writer
        .write_event(Event::Text(BytesText::new(formula)))
        .map_err(|error| XlsxError::InvalidData(error.to_string()))?;
    writer
        .write_event(Event::End(BytesEnd::new("f")))
        .map_err(|error| XlsxError::InvalidData(error.to_string()))
}

fn write_cell_value(
    writer: &mut Writer<Cursor<Vec<u8>>>,
    row: u32,
    col: u32,
    value: &CellEditValue,
) -> Result<(), XlsxError> {
    let reference = format!("{}{row}", col_to_letter(col));
    let mut cell = BytesStart::new("c");
    cell.push_attribute(("r", reference.as_str()));
    if let Some(cell_type) = value.cell_type() {
        cell.push_attribute(("t", cell_type));
    }
    writer
        .write_event(Event::Start(cell))
        .map_err(|error| XlsxError::InvalidData(error.to_string()))?;
    if let Some(formula) = value.formula_text() {
        write_formula(writer, formula)?;
    }
    if let Some(text) = value.value_text()? {
        writer
            .write_event(Event::Start(BytesStart::new("v")))
            .map_err(|error| XlsxError::InvalidData(error.to_string()))?;
        writer
            .write_event(Event::Text(BytesText::new(&text)))
            .map_err(|error| XlsxError::InvalidData(error.to_string()))?;
        writer
            .write_event(Event::End(BytesEnd::new("v")))
            .map_err(|error| XlsxError::InvalidData(error.to_string()))?;
    }
    writer
        .write_event(Event::End(BytesEnd::new("c")))
        .map_err(|error| XlsxError::InvalidData(error.to_string()))
}

/// Every cell a sheet holds, keyed by its one-based row and zero-based column.
fn cells_of(sheet: &crate::ir::Sheet) -> BTreeMap<(u32, u32), &crate::ir::Cell> {
    let mut cells = BTreeMap::new();
    for row in &sheet.rows {
        for cell in &row.cells {
            cells.insert((row.index, cell.col), cell);
        }
    }
    cells
}

fn same_content(before: &crate::ir::Cell, after: &crate::ir::Cell) -> bool {
    before.formula == after.formula && same_value(&before.value, &after.value)
}

fn same_value(before: &crate::ir::CellValue, after: &crate::ir::CellValue) -> bool {
    use crate::ir::CellValue::*;
    match (before, after) {
        (Empty, Empty) => true,
        (String(before), String(after)) => before == after,
        (Number(before), Number(after)) => before == after,
        (Boolean(before), Boolean(after)) => before == after,
        (Error(before), Error(after)) => before == after,
        _ => false,
    }
}

/// How a cell should be written back, or `None` for one this editor cannot
/// express — an error value has no edit of its own.
fn edit_for(cell: &crate::ir::Cell) -> Option<CellEditValue> {
    if let Some(formula) = cell.formula.as_ref() {
        return Some(CellEditValue::Formula(formula.clone()));
    }
    match &cell.value {
        crate::ir::CellValue::Empty => Some(CellEditValue::Empty),
        crate::ir::CellValue::String(value) => Some(CellEditValue::String(value.clone())),
        crate::ir::CellValue::Number(value) => Some(CellEditValue::Number(*value)),
        crate::ir::CellValue::Boolean(value) => Some(CellEditValue::Boolean(*value)),
        crate::ir::CellValue::Error(_) => None,
    }
}

/// Writes one `<col>` span, splitting it where a change covers only part of it.
///
/// A span reads `<col min="1" max="3" .../>` and covers three columns at once,
/// so hiding the middle one means writing three spans in its place.
fn write_col_span(
    writer: &mut Writer<Cursor<Vec<u8>>>,
    start: &BytesStart<'_>,
    pending: &mut BTreeMap<u32, bool>,
) -> Result<(), XlsxError> {
    let min = get_attr(start, "min")
        .and_then(|value| value.parse::<u32>().ok())
        .unwrap_or(1);
    let max = get_attr(start, "max")
        .and_then(|value| value.parse::<u32>().ok())
        .unwrap_or(min);
    let was_hidden = matches!(get_attr(start, "hidden").as_deref(), Some("1") | Some("true"));

    // Group the columns of this span into runs that share an answer.
    let mut runs: Vec<(u32, u32, bool)> = Vec::new();
    for column in min..=max {
        let hidden = pending
            .remove(&(column - 1))
            .unwrap_or(was_hidden);
        match runs.last_mut() {
            Some((_, last, held)) if *held == hidden && *last + 1 == column => *last = column,
            _ => runs.push((column, column, hidden)),
        }
    }

    for (first, last, hidden) in runs {
        let mut span = with_hidden(start, hidden);
        let mut rebuilt = BytesStart::new("col");
        for attribute in span.attributes().flatten() {
            let key = String::from_utf8_lossy(attribute.key.as_ref()).into_owned();
            if key == "min" || key == "max" {
                continue;
            }
            let value = String::from_utf8_lossy(&attribute.value).into_owned();
            rebuilt.push_attribute((key.as_str(), value.as_str()));
        }
        rebuilt.push_attribute(("min", first.to_string().as_str()));
        rebuilt.push_attribute(("max", last.to_string().as_str()));
        span = rebuilt;
        writer
            .write_event(Event::Empty(span))
            .map_err(|error| XlsxError::InvalidData(error.to_string()))?;
    }
    Ok(())
}

/// Writes `<col>` spans for columns the sheet never described.
fn write_new_cols(
    writer: &mut Writer<Cursor<Vec<u8>>>,
    pending: &mut BTreeMap<u32, bool>,
) -> Result<(), XlsxError> {
    for (column, hidden) in std::mem::take(pending) {
        if !hidden {
            continue;
        }
        let reference = (column + 1).to_string();
        let mut span = BytesStart::new("col");
        span.push_attribute(("min", reference.as_str()));
        span.push_attribute(("max", reference.as_str()));
        span.push_attribute(("hidden", "1"));
        writer
            .write_event(Event::Empty(span))
            .map_err(|error| XlsxError::InvalidData(error.to_string()))?;
    }
    Ok(())
}

fn open_cols(writer: &mut Writer<Cursor<Vec<u8>>>) -> Result<(), XlsxError> {
    writer
        .write_event(Event::Start(BytesStart::new("cols")))
        .map_err(|error| XlsxError::InvalidData(error.to_string()))
}

fn close_cols(writer: &mut Writer<Cursor<Vec<u8>>>) -> Result<(), XlsxError> {
    writer
        .write_event(Event::End(BytesEnd::new("cols")))
        .map_err(|error| XlsxError::InvalidData(error.to_string()))
}

/// Copies a start tag, replacing whatever it said about being hidden.
fn with_hidden(start: &BytesStart<'_>, hidden: bool) -> BytesStart<'static> {
    let name = String::from_utf8_lossy(start.name().as_ref()).into_owned();
    let mut rewritten = BytesStart::new(name);
    for attribute in start.attributes().flatten() {
        let key = String::from_utf8_lossy(attribute.key.as_ref()).into_owned();
        if key == "hidden" {
            continue;
        }
        let value = String::from_utf8_lossy(&attribute.value).into_owned();
        rewritten.push_attribute((key.as_str(), value.as_str()));
    }
    if hidden {
        rewritten.push_attribute(("hidden", "1"));
    }
    rewritten
}

fn write_inserted_row(
    writer: &mut Writer<Cursor<Vec<u8>>>,
    row: u32,
    cells: BTreeMap<u32, CellEditValue>,
    hidden: Option<bool>,
) -> Result<(), XlsxError> {
    let row_text = row.to_string();
    let mut row_start = BytesStart::new("row");
    row_start.push_attribute(("r", row_text.as_str()));
    if hidden == Some(true) {
        row_start.push_attribute(("hidden", "1"));
    }
    writer
        .write_event(Event::Start(row_start))
        .map_err(|error| XlsxError::InvalidData(error.to_string()))?;
    for (col, value) in cells {
        write_cell_value(writer, row, col, &value)?;
    }
    writer
        .write_event(Event::End(BytesEnd::new("row")))
        .map_err(|error| XlsxError::InvalidData(error.to_string()))
}

/// Patch worksheet XML, replacing or inserting cells at specified positions.
/// Each edit is written with its corresponding OOXML scalar type.
/// Everything one worksheet has to change.
struct SheetEdits<'a> {
    cells: &'a HashMap<(u32, u32), &'a CellEditValue>,
    /// One-based row -> hidden.
    rows: &'a BTreeMap<u32, bool>,
    /// Zero-based column -> hidden.
    cols: &'a BTreeMap<u32, bool>,
}

fn patch_worksheet_xml(xml: &str, sheet_edits: &SheetEdits<'_>) -> Result<String, XlsxError> {
    let edits = sheet_edits.cells;
    let mut reader = Reader::from_str(xml);
    let mut writer = Writer::new(Cursor::new(Vec::new()));
    let mut pending: BTreeMap<u32, BTreeMap<u32, CellEditValue>> = BTreeMap::new();
    for (&(row, col), value) in edits {
        pending
            .entry(row)
            .or_default()
            .insert(col, (*value).clone());
    }

    let mut pending_rows: std::collections::BTreeSet<u32> =
        sheet_edits.rows.keys().copied().collect();
    let mut pending_cols: BTreeMap<u32, bool> = sheet_edits.cols.clone();
    let mut seen_cols = false;

    let mut current_row: u32 = 0;
    let mut in_row = false;
    let mut in_cell = false;
    let mut cell_col: u32;
    let mut cell_row: u32;
    let mut in_value = false;
    let mut current_edit: Option<CellEditValue> = None;
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
                if name == "cols" {
                    seen_cols = true;
                }
                if name == "sheetData" && !seen_cols && !pending_cols.is_empty() {
                    open_cols(&mut writer)?;
                    write_new_cols(&mut writer, &mut pending_cols)?;
                    close_cols(&mut writer)?;
                    seen_cols = true;
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
                                let hidden = sheet_edits.rows.get(&row).copied();
                                write_inserted_row(&mut writer, row, cells, hidden)?;
                            }
                        }
                        in_row = true;
                        current_row = next_row;
                        pending_rows.remove(&next_row);
                        if let Some(hidden) = sheet_edits.rows.get(&next_row) {
                            writer
                                .write_event(Event::Start(with_hidden(e, *hidden)))
                                .map_err(|e| XlsxError::InvalidData(e.to_string()))?;
                            continue;
                        }
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
                                    write_cell_value(
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

                        if let Some(value) = &current_edit {
                            // Rewrite the cell type while preserving its reference and style.
                            let mut new_start = BytesStart::new("c");
                            for attr in e.attributes().flatten() {
                                let key = std::str::from_utf8(attr.key.as_ref()).unwrap_or("");
                                if key == "t" {
                                    continue;
                                }
                                new_start.push_attribute((key, std::str::from_utf8(&attr.value).unwrap_or("")));
                            }
                            if let Some(cell_type) = value.cell_type() {
                                new_start.push_attribute(("t", cell_type));
                            }
                            writer.write_event(Event::Start(new_start)).map_err(|e| XlsxError::InvalidData(e.to_string()))?;
                            if let Some(formula) = value.formula_text() {
                                write_formula(&mut writer, formula)?;
                            }
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
                        if let Some(text) = current_edit
                            .as_ref()
                            .map(CellEditValue::value_text)
                            .transpose()?
                            .flatten()
                        {
                            writer.write_event(Event::Text(BytesText::new(&text))).map_err(|e| XlsxError::InvalidData(e.to_string()))?;
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
                if name == "cols" {
                    write_new_cols(&mut writer, &mut pending_cols)?;
                }
                match name.as_str() {
                    "row" => {
                        if let Some(cells) = pending.remove(&current_row) {
                            for (col, value) in cells {
                                write_cell_value(
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
                                if let Some(text) = current_edit
                                    .as_ref()
                                    .map(CellEditValue::value_text)
                                    .transpose()?
                                    .flatten()
                                {
                                    writer.write_event(Event::Start(BytesStart::new("v"))).map_err(|e| XlsxError::InvalidData(e.to_string()))?;
                                    writer.write_event(Event::Text(BytesText::new(&text))).map_err(|e| XlsxError::InvalidData(e.to_string()))?;
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
                            let hidden = sheet_edits.rows.get(&row).copied();
                            write_inserted_row(&mut writer, row, cells, hidden)?;
                        }
                        // A row with nothing in it still has to be written for
                        // its hidden flag to survive.
                        let rows_left = std::mem::take(&mut pending_rows);
                        for row in rows_left {
                            if let Some(true) = sheet_edits.rows.get(&row) {
                                write_inserted_row(
                                    &mut writer,
                                    row,
                                    BTreeMap::new(),
                                    Some(true),
                                )?;
                            }
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
                if name == "col" {
                    seen_cols = true;
                    write_col_span(&mut writer, e, &mut pending_cols)?;
                    continue;
                }
                if in_cell && current_edit.is_some() && (name == "f" || name == "is") {
                    continue;
                }
                if name == "sheetData" && !seen_cols && !pending_cols.is_empty() {
                    open_cols(&mut writer)?;
                    write_new_cols(&mut writer, &mut pending_cols)?;
                    close_cols(&mut writer)?;
                    seen_cols = true;
                }
                if name == "sheetData" && !pending.is_empty() {
                    writer
                        .write_event(Event::Start(e.clone()))
                        .map_err(|error| XlsxError::InvalidData(error.to_string()))?;
                    let remaining = std::mem::take(&mut pending);
                    for (row, cells) in remaining {
                        let hidden = sheet_edits.rows.get(&row).copied();
                        write_inserted_row(&mut writer, row, cells, hidden)?;
                    }
                    for row in std::mem::take(&mut pending_rows) {
                        if let Some(true) = sheet_edits.rows.get(&row) {
                            write_inserted_row(&mut writer, row, BTreeMap::new(), Some(true))?;
                        }
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
                            let hidden = sheet_edits.rows.get(&row).copied();
                            write_inserted_row(&mut writer, row, cells, hidden)?;
                        }
                    }
                    pending_rows.remove(&row_num);
                    if let Some(cells) = pending.remove(&row_num) {
                        writer.write_event(Event::Start(e.clone())).map_err(|error| {
                            XlsxError::InvalidData(error.to_string())
                        })?;
                        for (col, value) in cells {
                            write_cell_value(&mut writer, row_num, col, &value)?;
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
                                write_cell_value(
                                    &mut writer,
                                    row_num,
                                    missing_col,
                                    &value,
                                )?;
                            }
                        }
                        edit = cells.remove(&col);
                    }

                    if let Some(value) = edit {
                        // Convert the empty cell while preserving its reference and style.
                        let mut new_start = BytesStart::new("c");
                        for attr in e.attributes().flatten() {
                            let key = std::str::from_utf8(attr.key.as_ref()).unwrap_or("");
                            if key == "t" { continue; }
                            new_start.push_attribute((key, std::str::from_utf8(&attr.value).unwrap_or("")));
                        }
                        if let Some(cell_type) = value.cell_type() {
                            new_start.push_attribute(("t", cell_type));
                        }
                        if let Some(formula) = value.formula_text() {
                            writer.write_event(Event::Start(new_start)).map_err(|e| XlsxError::InvalidData(e.to_string()))?;
                            write_formula(&mut writer, formula)?;
                            writer.write_event(Event::End(BytesEnd::new("c"))).map_err(|e| XlsxError::InvalidData(e.to_string()))?;
                        } else if let Some(text) = value.value_text()? {
                            writer.write_event(Event::Start(new_start)).map_err(|e| XlsxError::InvalidData(e.to_string()))?;
                            writer.write_event(Event::Start(BytesStart::new("v"))).map_err(|e| XlsxError::InvalidData(e.to_string()))?;
                            writer.write_event(Event::Text(BytesText::new(&text))).map_err(|e| XlsxError::InvalidData(e.to_string()))?;
                            writer.write_event(Event::End(BytesEnd::new("v"))).map_err(|e| XlsxError::InvalidData(e.to_string()))?;
                            writer.write_event(Event::End(BytesEnd::new("c"))).map_err(|e| XlsxError::InvalidData(e.to_string()))?;
                        } else {
                            writer.write_event(Event::Empty(new_start)).map_err(|e| XlsxError::InvalidData(e.to_string()))?;
                        }
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
    fn test_editor_preserves_typed_cell_edits() {
        let data = include_bytes!("../../../tests/fixtures/basic_test.xlsx");
        let mut editor = XlsxEditor::new(data).expect("should open");
        editor.set_cell_value(0, 1, 20, CellEditValue::Number(12.5));
        editor.set_cell_value(0, 1, 21, CellEditValue::Boolean(true));
        editor.set_cell_value(0, 1, 22, CellEditValue::Empty);

        let saved = editor.save().expect("should save");
        let workbook = parse_xlsx(&saved).expect("should parse");
        let row = &workbook.sheets[0].rows[0];
        assert!(matches!(row.cells.iter().find(|cell| cell.col == 20).unwrap().value, crate::ir::CellValue::Number(12.5)));
        assert!(matches!(row.cells.iter().find(|cell| cell.col == 21).unwrap().value, crate::ir::CellValue::Boolean(true)));
        assert!(matches!(row.cells.iter().find(|cell| cell.col == 22).unwrap().value, crate::ir::CellValue::Empty));
    }

    #[test]
    fn test_editor_preserves_formula_edits() {
        let data = include_bytes!("../../../tests/fixtures/basic_test.xlsx");
        let mut editor = XlsxEditor::new(data).expect("should open");
        editor.set_cell_value(
            0,
            1,
            23,
            CellEditValue::Formula("=SUM(A2:A3)".to_string()),
        );

        let saved = editor.save().expect("should save");
        let workbook = parse_xlsx(&saved).expect("should parse");
        let cell = workbook.sheets[0].rows[0]
            .cells
            .iter()
            .find(|cell| cell.col == 23)
            .unwrap();
        assert_eq!(cell.formula.as_deref(), Some("SUM(A2:A3)"));
    }

    /// The patcher takes everything a sheet changes; these tests only change cells.
    fn cells_only<'a>(
        cells: &'a HashMap<(u32, u32), &'a CellEditValue>,
    ) -> SheetEdits<'a> {
        static NOTHING: std::sync::OnceLock<BTreeMap<u32, bool>> = std::sync::OnceLock::new();
        let nothing = NOTHING.get_or_init(BTreeMap::new);
        SheetEdits {
            cells,
            rows: nothing,
            cols: nothing,
        }
    }

    #[test]
    fn worksheet_patch_replaces_an_existing_formula() {
        let xml = r#"<worksheet><sheetData><row r="1"><c r="A1"><f>OLD()</f><v>1</v></c></row></sheetData></worksheet>"#;
        let formula = CellEditValue::Formula("=SUM(B1:B2)".to_string());
        let edits = HashMap::from([((1, 0), &formula)]);

        let patched = patch_worksheet_xml(xml, &cells_only(&edits)).expect("should replace formula");
        assert!(patched.contains("<f>SUM(B1:B2)</f>"));
        assert!(!patched.contains("OLD()"));
        assert_eq!(patched.matches("<f>").count(), 1);
    }

    #[test]
    fn worksheet_patch_orders_insertions_and_replaces_formula_content() {
        let xml = r#"<worksheet><sheetData><row r="1"><c r="A1"><f>1+1</f><v>2</v></c><c r="C1" t="inlineStr"><is><t>old</t></is></c></row><row r="3"/></sheetData></worksheet>"#;
        let formula_value = CellEditValue::String("9".to_string());
        let middle_value = CellEditValue::String("middle".to_string());
        let inline_value = CellEditValue::String("new".to_string());
        let new_row_value = CellEditValue::String("row two".to_string());
        let edits = HashMap::from([
            ((1, 0), &formula_value),
            ((1, 1), &middle_value),
            ((1, 2), &inline_value),
            ((2, 1), &new_row_value),
        ]);

        let patched = patch_worksheet_xml(xml, &cells_only(&edits)).expect("should patch worksheet");
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
        let value = CellEditValue::String("first value".to_string());
        let edits = HashMap::from([((1, 0), &value)]);

        let patched = patch_worksheet_xml(xml, &cells_only(&edits)).expect("should patch empty worksheet");
        let expected = "<sheetData><row r=\"1\"><c r=\"A1\" t=\"str\">\
                        <v>first value</v></c></row></sheetData>";
        assert!(patched.contains(expected));
    }
}
