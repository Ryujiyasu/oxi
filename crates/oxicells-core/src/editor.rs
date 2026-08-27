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

use crate::ir::{AutoFilter, BorderLine, CellStyle, MergeCell, Workbook};
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
    /// (sheet_idx, 0-based column) -> width, in characters of the workbook's
    /// default font, gutter included. This is the number the file states, not
    /// the one a person types into Excel: typing 10 stores 10.625.
    col_width: HashMap<(usize, u32), f32>,
    /// Sheets whose merged cells are being replaced wholesale. A merge is not
    /// edited in place, so the whole list travels together.
    merges: HashMap<usize, Vec<MergeCell>>,
    /// (sheet_idx, 1-based row, 0-based column) -> the style it should carry.
    styles: HashMap<(usize, u32, u32), CellStyle>,
    /// Sheets whose filter is being replaced. `None` takes the filter away.
    filters: HashMap<usize, Option<AutoFilter>>,
    /// What the workbook's sheets should end up as, when they are changing.
    sheet_plan: Option<Vec<PlannedSheet>>,
}

/// One sheet of a rearranged workbook: either one the file already holds, or a
/// new one to be written from scratch.
enum PlannedSheet {
    /// The sheet at this position in the original file, under this name.
    Held { origin: usize, name: String },
    /// A sheet the file has never held.
    Added(crate::ir::Sheet),
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
            col_width: HashMap::new(),
            merges: HashMap::new(),
            styles: HashMap::new(),
            filters: HashMap::new(),
            sheet_plan: None,
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
    /// Only what the editor can write is compared: cell values and formulas,
    /// which rows and columns are hidden, which cells are merged, the styling
    /// each cell wears, the filter the sheet is under, and which sheets the
    /// workbook holds.
    ///
    /// Anything the IR does not model — conditional formatting, data
    /// validation, charts, drawings — rides along in the original XML
    /// untouched. A run cannot change those, and cannot lose them either.
    pub fn apply_workbook(&mut self, edited: &Workbook) -> Result<(), XlsxError> {
        if edited.sheets.is_empty() {
            return Err(XlsxError::InvalidData(
                "a workbook must keep at least one sheet".to_string(),
            ));
        }

        // Sheets are matched by position when there are as many as before, so a
        // rename is a rename. Otherwise they are matched by name, and whatever
        // is left over has been added or taken away.
        let same_count = edited.sheets.len() == self.workbook.sheets.len();
        let mut origins: Vec<Option<usize>> = Vec::with_capacity(edited.sheets.len());
        for (index, sheet) in edited.sheets.iter().enumerate() {
            let origin = if same_count {
                Some(index)
            } else {
                self.workbook
                    .sheets
                    .iter()
                    .position(|held| held.name.eq_ignore_ascii_case(&sheet.name))
            };
            origins.push(origin);
        }

        let rearranged = !same_count
            || origins.iter().enumerate().any(|(index, origin)| {
                *origin != Some(index)
            })
            || edited
                .sheets
                .iter()
                .zip(&self.workbook.sheets)
                .any(|(after, before)| after.name != before.name);
        if rearranged {
            self.sheet_plan = Some(
                edited
                    .sheets
                    .iter()
                    .zip(&origins)
                    .map(|(sheet, origin)| match origin {
                        Some(origin) => PlannedSheet::Held {
                            origin: *origin,
                            name: sheet.name.clone(),
                        },
                        None => PlannedSheet::Added(sheet.clone()),
                    })
                    .collect(),
            );
        }

        let pairs: Vec<(usize, (&crate::ir::Sheet, &crate::ir::Sheet))> = edited
            .sheets
            .iter()
            .zip(&origins)
            .filter_map(|(after, origin)| {
                origin.map(|origin| (origin, (&self.workbook.sheets[origin], after)))
            })
            .collect();

        for (index, (before, after)) in pairs {
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

            // A width of zero is how the parser records "this column had no
            // `<col>` entry of its own", so it means nothing to write and is
            // not a change worth chasing.
            for (col, width) in after.col_widths.iter().enumerate() {
                let was = before.col_widths.get(col).copied().unwrap_or(0.0);
                if *width != was && *width != 0.0 {
                    self.col_width.insert((index, col as u32), *width);
                }
            }

            if !same_merges(&before.merge_cells, &after.merge_cells) {
                self.merges.insert(index, after.merge_cells.clone());
            }

            if before.auto_filter != after.auto_filter {
                self.filters.insert(index, after.auto_filter.clone());
            }

            for (&(row, col), cell) in &now {
                let was = held.get(&(row, col)).map(|before| &before.style);
                if was != Some(&cell.style) {
                    self.styles
                        .insert((index, row, col), cell.style.clone());
                }
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

    /// Sets how wide a column is.
    ///
    /// The width is in characters of the workbook's default font with the
    /// gutter either side counted in, which is what the file states and not
    /// what Excel's own box shows: typing 10 there stores 10.625.
    pub fn set_col_width(&mut self, sheet_index: usize, col: u32, width: f32) {
        self.col_width.insert((sheet_index, col), width);
    }

    /// Replaces every merged cell on a sheet.
    pub fn set_merges(&mut self, sheet_index: usize, merges: Vec<MergeCell>) {
        self.merges.insert(sheet_index, merges);
    }

    /// Gives a cell a style of its own.
    pub fn set_cell_style(
        &mut self,
        sheet_index: usize,
        row: u32,
        col: u32,
        style: CellStyle,
    ) {
        self.styles.insert((sheet_index, row, col), style);
    }

    /// Puts a sheet under a filter, or takes its filter away with `None`.
    pub fn set_auto_filter(&mut self, sheet_index: usize, filter: Option<AutoFilter>) {
        self.filters.insert(sheet_index, filter);
    }

    pub fn has_edits(&self) -> bool {
        !self.edits.is_empty()
            || !self.row_hidden.is_empty()
            || !self.col_hidden.is_empty()
            || !self.col_width.is_empty()
            || !self.merges.is_empty()
            || !self.styles.is_empty()
            || !self.filters.is_empty()
            || self.sheet_plan.is_some()
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

        // One xf per distinct style, so a run that bolds a hundred cells adds
        // one entry rather than a hundred.
        let mut distinct: Vec<CellStyle> = Vec::new();
        let mut style_slot: HashMap<(usize, u32, u32), usize> = HashMap::new();
        for (key, style) in &self.styles {
            let slot = match distinct.iter().position(|held| held == style) {
                Some(slot) => slot,
                None => {
                    distinct.push(style.clone());
                    distinct.len() - 1
                }
            };
            style_slot.insert(*key, slot);
        }

        // The style sheet is read ahead of the walk below, since a worksheet
        // needs the indices and may come before it in the archive.
        let mut patched_styles: Option<String> = None;
        // What each planned sheet is called in the package, and which of the
        // original parts it comes from.
        let sheet_paths_by_index = self.resolve_sheet_paths()?;
        let mut planned_parts: Vec<(String, Option<usize>)> = Vec::new();
        if let Some(plan) = self.sheet_plan.as_ref() {
            let mut next_part = sheet_paths_by_index.len() + 1;
            for planned in plan {
                match planned {
                    PlannedSheet::Held { origin, .. } => {
                        let path = sheet_paths_by_index.get(*origin).cloned().ok_or_else(|| {
                            XlsxError::InvalidData(
                                "a kept sheet is not in the original workbook".to_string(),
                            )
                        })?;
                        planned_parts.push((path, Some(*origin)));
                    }
                    PlannedSheet::Added(_) => {
                        // Take a part name nothing is using.
                        let mut path = format!("xl/worksheets/sheet{next_part}.xml");
                        while sheet_paths_by_index.contains(&path) {
                            next_part += 1;
                            path = format!("xl/worksheets/sheet{next_part}.xml");
                        }
                        next_part += 1;
                        planned_parts.push((path, None));
                    }
                }
            }
        }
        let dropped: Vec<String> = match self.sheet_plan.as_ref() {
            None => Vec::new(),
            Some(_) => sheet_paths_by_index
                .iter()
                .filter(|path| !planned_parts.iter().any(|(held, _)| held == *path))
                .cloned()
                .collect(),
        };

        let mut style_indices: Vec<u32> = Vec::new();
        if !distinct.is_empty() {
            let mut xml = String::new();
            archive
                .by_name("xl/styles.xml")
                .map_err(|_| {
                    XlsxError::InvalidData(
                        "the workbook has no style sheet to add styles to".to_string(),
                    )
                })?
                .read_to_string(&mut xml)
                .map_err(|error| XlsxError::InvalidData(error.to_string()))?;
            let (patched, indices) = patch_styles_xml(&xml, &distinct)?;
            patched_styles = Some(patched);
            style_indices = indices;
        }

        // Where each cell's style ended up, ready for the sheet writer.
        let mut cell_styles: HashMap<usize, BTreeMap<(u32, u32), u32>> = HashMap::new();
        for ((sheet, row, col), slot) in &style_slot {
            if let Some(index) = style_indices.get(*slot) {
                cell_styles
                    .entry(*sheet)
                    .or_default()
                    .insert((*row, *col), *index);
            }
        }

        let mut rows_by_sheet: HashMap<usize, BTreeMap<u32, bool>> = HashMap::new();
        for ((sheet, row), hidden) in &self.row_hidden {
            rows_by_sheet.entry(*sheet).or_default().insert(*row, *hidden);
        }
        let mut cols_by_sheet: HashMap<usize, BTreeMap<u32, ColumnWant>> = HashMap::new();
        for ((sheet, col), hidden) in &self.col_hidden {
            cols_by_sheet
                .entry(*sheet)
                .or_default()
                .entry(*col)
                .or_default()
                .hidden = Some(*hidden);
        }
        for ((sheet, col), width) in &self.col_width {
            cols_by_sheet
                .entry(*sheet)
                .or_default()
                .entry(*col)
                .or_default()
                .width = Some(*width);
        }


        // Map sheet path -> everything to change in that sheet
        let empty_cells: HashMap<(u32, u32), &CellEditValue> = HashMap::new();
        let empty_lines: BTreeMap<u32, bool> = BTreeMap::new();
        let empty_cols: BTreeMap<u32, ColumnWant> = BTreeMap::new();
        let empty_styles: BTreeMap<(u32, u32), u32> = BTreeMap::new();
        let mut path_edits: HashMap<String, SheetEdits<'_>> = HashMap::new();
        for sheet in edits_by_sheet
            .keys()
            .chain(rows_by_sheet.keys())
            .chain(cols_by_sheet.keys())
            .chain(self.merges.keys())
            .chain(cell_styles.keys())
            .chain(self.filters.keys())
            .copied()
            .collect::<std::collections::BTreeSet<_>>()
        {
            if let Some(path) = sheet_paths.get(sheet) {
                path_edits.insert(
                    path.clone(),
                    SheetEdits {
                        cells: edits_by_sheet.get(&sheet).unwrap_or(&empty_cells),
                        rows: rows_by_sheet.get(&sheet).unwrap_or(&empty_lines),
                        cols: cols_by_sheet.get(&sheet).unwrap_or(&empty_cols),
                        merges: self.merges.get(&sheet).map(Vec::as_slice),
                        styles: cell_styles.get(&sheet).unwrap_or(&empty_styles),
                        filter: self.filters.get(&sheet),
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

                if dropped.contains(&name) {
                    // The sheet is gone; leave its part out of the package.
                    writer
                        .abort_file()
                        .map_err(|e| XlsxError::InvalidData(e.to_string()))?;
                    continue;
                }
                // The calculation chain names the cells that hold formulas and
                // the order Excel works them out in. Writing a value over a
                // formula takes the formula away, and a chain still pointing at
                // that cell makes Excel refuse the WHOLE workbook — `fies_t2`
                // has one entry, `C1`, and replacing C1's `<f>C9</f>` with a
                // value lost the file. The chain is only a cache: Excel builds
                // it again on open, so the safe answer is to leave it out
                // whenever a cell has been touched.
                if name == "xl/calcChain.xml" {
                    writer
                        .abort_file()
                        .map_err(|e| XlsxError::InvalidData(e.to_string()))?;
                    continue;
                }
                if let (Some(plan), "xl/workbook.xml") = (self.sheet_plan.as_ref(), name.as_str())
                {
                    let mut xml = String::new();
                    entry
                        .read_to_string(&mut xml)
                        .map_err(|e| XlsxError::InvalidData(e.to_string()))?;
                    let names: Vec<String> = plan
                        .iter()
                        .map(|planned| match planned {
                            PlannedSheet::Held { name, .. } => name.clone(),
                            PlannedSheet::Added(sheet) => sheet.name.clone(),
                        })
                        .collect();
                    let patched = patch_workbook_sheets(&xml, &names)?;
                    writer
                        .write_all(patched.as_bytes())
                        .map_err(|e| XlsxError::InvalidData(e.to_string()))?;
                    continue;
                }
                if self.sheet_plan.is_some() && name == "xl/_rels/workbook.xml.rels" {
                    let mut xml = String::new();
                    entry
                        .read_to_string(&mut xml)
                        .map_err(|e| XlsxError::InvalidData(e.to_string()))?;
                    let targets: Vec<String> = planned_parts
                        .iter()
                        .map(|(path, _)| {
                            path.strip_prefix("xl/").unwrap_or(path).to_string()
                        })
                        .collect();
                    let patched = patch_workbook_rels(&xml, &targets)?;
                    writer
                        .write_all(patched.as_bytes())
                        .map_err(|e| XlsxError::InvalidData(e.to_string()))?;
                    continue;
                }
                if self.sheet_plan.is_some() && name == "[Content_Types].xml" {
                    let mut xml = String::new();
                    entry
                        .read_to_string(&mut xml)
                        .map_err(|e| XlsxError::InvalidData(e.to_string()))?;
                    let parts: Vec<String> =
                        planned_parts.iter().map(|(path, _)| path.clone()).collect();
                    let patched = patch_content_types(&xml, &parts)?;
                    writer
                        .write_all(patched.as_bytes())
                        .map_err(|e| XlsxError::InvalidData(e.to_string()))?;
                    continue;
                }
                // The package must not name a part that is no longer in it:
                // the chain's own relationship and content type go with it.
                // Excel opens the workbook either way, but a relationship
                // pointing at nothing is not a package, and "invalid but
                // tolerated" is what this whole exercise is meant to catch.
                if name == "xl/_rels/workbook.xml.rels" || name == "[Content_Types].xml" {
                    let mut xml = String::new();
                    entry
                        .read_to_string(&mut xml)
                        .map_err(|e| XlsxError::InvalidData(e.to_string()))?;
                    let patched = without_calc_chain(&xml);
                    writer
                        .write_all(patched.as_bytes())
                        .map_err(|e| XlsxError::InvalidData(e.to_string()))?;
                    continue;
                }
                if name == "xl/styles.xml" {
                    if let Some(patched) = patched_styles.as_deref() {
                        writer
                            .write_all(patched.as_bytes())
                            .map_err(|e| XlsxError::InvalidData(e.to_string()))?;
                        continue;
                    }
                }
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

            // Sheets the file never held are written after everything it did.
            if let Some(plan) = self.sheet_plan.as_ref() {
                for (planned, (path, origin)) in plan.iter().zip(&planned_parts) {
                    if origin.is_some() {
                        continue;
                    }
                    let PlannedSheet::Added(sheet) = planned else {
                        continue;
                    };
                    writer
                        .start_file(path, SimpleFileOptions::default())
                        .map_err(|e| XlsxError::InvalidData(e.to_string()))?;
                    writer
                        .write_all(&write_new_sheet(sheet)?)
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

/// Write one cell.
///
/// `styles` says which cellXfs index each cell should carry. A cell being
/// written for the first time has to be given its style here: nothing later
/// puts one on it, and without it the cell comes out under General however it
/// was styled — which is how a date lands in the file as a bare serial number.
fn write_cell_value(
    writer: &mut Writer<Cursor<Vec<u8>>>,
    row: u32,
    col: u32,
    value: &CellEditValue,
    styles: &BTreeMap<(u32, u32), u32>,
) -> Result<(), XlsxError> {
    let reference = format!("{}{row}", col_to_letter(col));
    let style_text = styles.get(&(row, col)).map(|index| index.to_string());
    let mut cell = BytesStart::new("c");
    cell.push_attribute(("r", reference.as_str()));
    if let Some(style) = style_text.as_deref() {
        cell.push_attribute(("s", style));
    }
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

/// Merges compare as a set: the order a sheet lists them in means nothing.
fn same_merges(before: &[MergeCell], after: &[MergeCell]) -> bool {
    if before.len() != after.len() {
        return false;
    }
    let key = |merge: &MergeCell| {
        (merge.start_row, merge.start_col, merge.end_row, merge.end_col)
    };
    let mut before: Vec<_> = before.iter().map(key).collect();
    let mut after: Vec<_> = after.iter().map(key).collect();
    before.sort_unstable();
    after.sort_unstable();
    before == after
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

/// Adds the styles a save needs to `styles.xml`, and says which `cellXfs`
/// index each one ended up at.
///
/// Styles are appended rather than matched against what is already there. A
/// style read back out of a file is only as complete as the parser that read
/// it, so reusing an existing entry risks carrying over something the parser
/// never saw. Appending costs a few bytes and keeps every original entry as it
/// was for the cells still using it.
fn patch_styles_xml(
    xml: &str,
    wanted: &[CellStyle],
) -> Result<(String, Vec<u32>), XlsxError> {
    if wanted.is_empty() {
        return Ok((xml.to_string(), Vec::new()));
    }

    let fonts_at = count_of(xml, "fonts");
    let fills_at = count_of(xml, "fills");
    let borders_at = count_of(xml, "borders");
    let xfs_at = count_of(xml, "cellXfs");
    let mut next_number_format = 164_u32.max(
        // Custom formats are numbered from 164; step past any the file already has.
        highest_custom_number_format(xml).map_or(164, |highest| highest + 1),
    );

    let mut fonts = String::new();
    let mut fills = String::new();
    let mut borders = String::new();
    let mut xfs = String::new();
    let mut number_formats = String::new();
    let mut indices = Vec::with_capacity(wanted.len());

    for (offset, style) in wanted.iter().enumerate() {
        let offset = offset as u32;
        fonts.push_str(&font_xml(style));
        fills.push_str(&fill_xml(style));
        borders.push_str(&border_xml(style));

        let number_format_id = match style.number_format.as_deref() {
            None => 0,
            Some(format) => match builtin_number_format_id(format) {
                Some(id) => id,
                None => {
                    number_formats.push_str(&format!(
                        "<numFmt numFmtId=\"{next_number_format}\" formatCode=\"{}\"/>",
                        escape(format)
                    ));
                    let id = next_number_format;
                    next_number_format += 1;
                    id
                }
            },
        };

        xfs.push_str(&xf_xml(
            style,
            number_format_id,
            fonts_at + offset,
            fills_at + offset,
            borders_at + offset,
        ));
        indices.push(xfs_at + offset);
    }

    let added = wanted.len() as u32;
    let mut patched = xml.to_string();
    patched = append_section(&patched, "numFmts", &number_formats, true)?;
    patched = append_section(&patched, "fonts", &fonts, false)?;
    patched = append_section(&patched, "fills", &fills, false)?;
    patched = append_section(&patched, "borders", &borders, false)?;
    patched = append_section(&patched, "cellXfs", &xfs, false)?;
    let _ = added;
    Ok((patched, indices))
}

/// How many entries a section already holds, read from its `count` attribute.
fn count_of(xml: &str, section: &str) -> u32 {
    let open = format!("<{section} ");
    let Some(start) = xml.find(&open).or_else(|| xml.find(&format!("<{section}>"))) else {
        return 0;
    };
    let Some(end) = xml[start..].find('>') else {
        return 0;
    };
    let tag = &xml[start..start + end];
    tag.find("count=\"")
        .and_then(|at| {
            let rest = &tag[at + 7..];
            rest.find('"').and_then(|close| rest[..close].parse().ok())
        })
        .unwrap_or(0)
}

fn highest_custom_number_format(xml: &str) -> Option<u32> {
    let mut highest = None;
    let mut rest = xml;
    while let Some(at) = rest.find("numFmtId=\"") {
        rest = &rest[at + 10..];
        let Some(close) = rest.find('"') else { break };
        if let Ok(id) = rest[..close].parse::<u32>() {
            if id >= 164 {
                highest = Some(highest.map_or(id, |held: u32| held.max(id)));
            }
        }
    }
    highest
}

/// Puts new entries at the end of a section, growing its `count`. A section
/// that is not there is created when `create` is set, which only numFmts needs.
fn append_section(
    xml: &str,
    section: &str,
    entries: &str,
    create: bool,
) -> Result<String, XlsxError> {
    if entries.is_empty() {
        return Ok(xml.to_string());
    }
    let added = entries.matches("<numFmt ").count().max(
        entries
            .matches(&format!("<{}", singular(section)))
            .count(),
    ) as u32;

    let closing = format!("</{section}>");
    if let Some(at) = xml.find(&closing) {
        let mut patched = String::with_capacity(xml.len() + entries.len());
        patched.push_str(&xml[..at]);
        patched.push_str(entries);
        patched.push_str(&xml[at..]);
        return Ok(bump_count(&patched, section, added));
    }

    // A self-closing section, such as <numFmts count="0"/>.
    let empty = format!("<{section} ");
    if let Some(at) = xml.find(&empty) {
        if let Some(end) = xml[at..].find("/>") {
            let head = &xml[..at + end];
            let mut patched = String::with_capacity(xml.len() + entries.len());
            patched.push_str(head);
            patched.push('>');
            patched.push_str(entries);
            patched.push_str(&closing);
            patched.push_str(&xml[at + end + 2..]);
            return Ok(bump_count(&patched, section, added));
        }
    }

    if !create {
        return Err(XlsxError::InvalidData(format!(
            "the workbook's styles have no {section} to add to"
        )));
    }
    // numFmts belongs at the head of the style sheet, ahead of the fonts.
    let at = xml.find("<fonts").ok_or_else(|| {
        XlsxError::InvalidData("the workbook's styles have no fonts".to_string())
    })?;
    let mut patched = String::with_capacity(xml.len() + entries.len() + 40);
    patched.push_str(&xml[..at]);
    patched.push_str(&format!("<{section} count=\"{added}\">"));
    patched.push_str(entries);
    patched.push_str(&closing);
    patched.push_str(&xml[at..]);
    Ok(patched)
}

fn singular(section: &str) -> &str {
    match section {
        "fonts" => "font",
        "fills" => "fill",
        "borders" => "border",
        "cellXfs" => "xf",
        other => other,
    }
}

fn bump_count(xml: &str, section: &str, added: u32) -> String {
    let open = format!("<{section} ");
    let Some(start) = xml.find(&open) else {
        return xml.to_string();
    };
    let Some(end) = xml[start..].find('>') else {
        return xml.to_string();
    };
    let tag = &xml[start..start + end];
    let Some(at) = tag.find("count=\"") else {
        return xml.to_string();
    };
    let rest = &tag[at + 7..];
    let Some(close) = rest.find('"') else {
        return xml.to_string();
    };
    let held: u32 = rest[..close].parse().unwrap_or(0);
    let replaced = format!("{}count=\"{}\"{}", &tag[..at], held + added, &rest[close + 1..]);
    format!("{}{}{}", &xml[..start], replaced, &xml[start + end..])
}

fn escape(value: &str) -> String {
    value
        .replace('&', "&amp;")
        .replace('<', "&lt;")
        .replace('>', "&gt;")
        .replace('"', "&quot;")
}

fn font_xml(style: &CellStyle) -> String {
    let mut font = String::from("<font>");
    if style.bold {
        font.push_str("<b/>");
    }
    if style.italic {
        font.push_str("<i/>");
    }
    if style.underline {
        font.push_str("<u/>");
    }
    if let Some(size) = style.font_size {
        font.push_str(&format!("<sz val=\"{size}\"/>"));
    }
    if let Some(color) = style.font_color.as_deref() {
        font.push_str(&format!("<color rgb=\"FF{}\"/>", escape(color)));
    }
    if let Some(name) = style.font_name.as_deref() {
        font.push_str(&format!("<name val=\"{}\"/>", escape(name)));
    }
    font.push_str("</font>");
    font
}

fn fill_xml(style: &CellStyle) -> String {
    match style.bg_color.as_deref() {
        Some(color) => format!(
            "<fill><patternFill patternType=\"solid\"><fgColor rgb=\"FF{}\"/><bgColor indexed=\"64\"/></patternFill></fill>",
            escape(color)
        ),
        None => "<fill><patternFill patternType=\"none\"/></fill>".to_string(),
    }
}

fn border_xml(style: &CellStyle) -> String {
    let edge = |name: &str, line: &Option<BorderLine>| match line {
        Some(line) => {
            let colour = match line.color.as_deref() {
                Some(colour) => format!("<color rgb=\"FF{}\"/>", escape(colour)),
                None => "<color indexed=\"64\"/>".to_string(),
            };
            format!(
                "<{name} style=\"{}\">{colour}</{name}>",
                escape(&line.style)
            )
        }
        None => format!("<{name}/>"),
    };
    format!(
        "<border>{}{}{}{}<diagonal/></border>",
        edge("left", &style.border_left),
        edge("right", &style.border_right),
        edge("top", &style.border_top),
        edge("bottom", &style.border_bottom)
    )
}

fn xf_xml(
    style: &CellStyle,
    number_format_id: u32,
    font_id: u32,
    fill_id: u32,
    border_id: u32,
) -> String {
    let mut xf = format!(
        "<xf numFmtId=\"{number_format_id}\" fontId=\"{font_id}\" fillId=\"{fill_id}\" borderId=\"{border_id}\" xfId=\"0\" applyFont=\"1\" applyFill=\"1\" applyBorder=\"1\""
    );
    if number_format_id != 0 {
        xf.push_str(" applyNumberFormat=\"1\"");
    }
    match style.horizontal_align.as_deref() {
        Some(alignment) => {
            xf.push_str(" applyAlignment=\"1\">");
            xf.push_str(&format!("<alignment horizontal=\"{}\"/>", escape(alignment)));
            xf.push_str("</xf>");
        }
        None => xf.push_str("/>"),
    }
    xf
}

/// The well-known number formats, so a common one does not need a new entry.
fn builtin_number_format_id(format: &str) -> Option<u32> {
    match format {
        "General" => Some(0),
        "0" => Some(1),
        "0.00" => Some(2),
        "#,##0" => Some(3),
        "#,##0.00" => Some(4),
        "0%" => Some(9),
        "0.00%" => Some(10),
        "0.00E+00" => Some(11),
        "mm-dd-yy" => Some(14),
        "m/d/yy h:mm" => Some(22),
        _ => None,
    }
}

/// Rewrites the `<sheets>` list of `xl/workbook.xml`.
///
/// Each sheet is named in order and given a relationship id matching its
/// position, which is what `patch_workbook_rels` writes on the other side.
fn patch_workbook_sheets(xml: &str, names: &[String]) -> Result<String, XlsxError> {
    let start = xml.find("<sheets").ok_or_else(|| {
        XlsxError::InvalidData("the workbook does not list its sheets".to_string())
    })?;
    let end = xml[start..]
        .find("</sheets>")
        .map(|at| start + at + "</sheets>".len())
        .or_else(|| {
            // A self-closing <sheets/> holds nothing, but is still the place.
            xml[start..].find("/>").map(|at| start + at + 2)
        })
        .ok_or_else(|| {
            XlsxError::InvalidData("the workbook's sheet list has no end".to_string())
        })?;

    let mut listed = String::from("<sheets>");
    for (position, name) in names.iter().enumerate() {
        listed.push_str(&format!(
            "<sheet name=\"{}\" sheetId=\"{}\" r:id=\"rIdSheet{}\"/>",
            escape(name),
            position + 1,
            position + 1
        ));
    }
    listed.push_str("</sheets>");
    Ok(format!("{}{}{}", &xml[..start], listed, &xml[end..]))
}

/// The same XML with any element naming the calculation chain taken out.
///
/// Used on `[Content_Types].xml` and on the workbook's relationships, where
/// the chain is declared. Both state it as a single self-closing element, so
/// cutting from its `<` to the `>` that closes it is enough and leaves the
/// rest of the part byte for byte as it was.
fn without_calc_chain(xml: &str) -> String {
    let mut out = String::with_capacity(xml.len());
    let mut rest = xml;
    while let Some(at) = rest.find("calcChain.xml") {
        let Some(open) = rest[..at].rfind('<') else {
            break;
        };
        let Some(close) = rest[at..].find('>') else {
            break;
        };
        out.push_str(&rest[..open]);
        rest = &rest[at + close + 1..];
    }
    out.push_str(rest);
    out
}

/// Rewrites the worksheet relationships of `xl/_rels/workbook.xml.rels`,
/// leaving every other relationship — styles, theme, shared strings — as it is.
fn patch_workbook_rels(xml: &str, targets: &[String]) -> Result<String, XlsxError> {
    const WORKSHEET: &str =
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet";

    let mut kept = String::new();
    let mut rest = xml;
    let end = rest.find("</Relationships>").ok_or_else(|| {
        XlsxError::InvalidData("the workbook has no relationships".to_string())
    })?;
    let tail = &rest[end..];
    rest = &rest[..end];

    // Keep everything that is not a worksheet, in the order it was written.
    let head_end = rest.find('>').map(|at| at + 1).unwrap_or(0);
    let mut head = rest[..head_end].to_string();
    if !head.contains("<Relationships") {
        // The declaration came first; find the real opening tag.
        let at = rest.find("<Relationships").ok_or_else(|| {
            XlsxError::InvalidData("the workbook has no relationships".to_string())
        })?;
        let close = rest[at..].find('>').map(|offset| at + offset + 1).unwrap();
        head = rest[..close].to_string();
        rest = &rest[close..];
    } else {
        rest = &rest[head_end..];
    }

    for piece in rest.split_inclusive("/>") {
        if piece.contains(WORKSHEET) {
            continue;
        }
        kept.push_str(piece);
    }

    let mut added = String::new();
    for (position, target) in targets.iter().enumerate() {
        added.push_str(&format!(
            "<Relationship Id=\"rIdSheet{}\" Type=\"{WORKSHEET}\" Target=\"{}\"/>",
            position + 1,
            escape(target)
        ));
    }
    Ok(format!("{head}{added}{kept}{tail}"))
}

/// Rewrites the worksheet overrides of `[Content_Types].xml`.
fn patch_content_types(xml: &str, parts: &[String]) -> Result<String, XlsxError> {
    const WORKSHEET: &str =
        "application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml";

    let end = xml.find("</Types>").ok_or_else(|| {
        XlsxError::InvalidData("the package has no content types".to_string())
    })?;
    let (body, tail) = xml.split_at(end);

    let mut kept = String::new();
    for piece in body.split_inclusive("/>") {
        if piece.contains(WORKSHEET) {
            continue;
        }
        kept.push_str(piece);
    }

    let mut added = String::new();
    for part in parts {
        added.push_str(&format!(
            "<Override PartName=\"/{}\" ContentType=\"{WORKSHEET}\"/>",
            escape(part)
        ));
    }
    Ok(format!("{kept}{added}{tail}"))
}

/// Writes a worksheet part from scratch, for a sheet the file never held.
fn write_new_sheet(sheet: &crate::ir::Sheet) -> Result<Vec<u8>, XlsxError> {
    let mut writer = Writer::new(Cursor::new(Vec::new()));
    let mut root = BytesStart::new("worksheet");
    root.push_attribute((
        "xmlns",
        "http://schemas.openxmlformats.org/spreadsheetml/2006/main",
    ));
    writer
        .write_event(Event::Start(root))
        .map_err(|error| XlsxError::InvalidData(error.to_string()))?;

    if sheet.rows.iter().all(|row| row.cells.is_empty() && !row.hidden) {
        writer
            .write_event(Event::Empty(BytesStart::new("sheetData")))
            .map_err(|error| XlsxError::InvalidData(error.to_string()))?;
    } else {
        writer
            .write_event(Event::Start(BytesStart::new("sheetData")))
            .map_err(|error| XlsxError::InvalidData(error.to_string()))?;
        for row in &sheet.rows {
            let mut cells = BTreeMap::new();
            for cell in &row.cells {
                if let Some(value) = edit_for(cell) {
                    cells.insert(cell.col, value);
                }
            }
            if cells.is_empty() && !row.hidden {
                continue;
            }
            write_inserted_row(&mut writer, row.index, cells, Some(row.hidden), &BTreeMap::new())?;
        }
        writer
            .write_event(Event::End(BytesEnd::new("sheetData")))
            .map_err(|error| XlsxError::InvalidData(error.to_string()))?;
    }

    write_merges(&mut writer, &sheet.merge_cells)?;
    writer
        .write_event(Event::End(BytesEnd::new("worksheet")))
        .map_err(|error| XlsxError::InvalidData(error.to_string()))?;

    let mut part = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>"#.to_vec();
    part.extend(writer.into_inner().into_inner());
    Ok(part)
}

/// Writes an `<autoFilter>` block, or nothing when the filter has been taken
/// away.
///
/// A criterion is stated the way VBA states it — `"apple"`, `">15"` — and is
/// written as a plain value list when it names values, or as a custom filter
/// when it compares.
fn write_auto_filter(
    writer: &mut Writer<Cursor<Vec<u8>>>,
    filter: Option<&AutoFilter>,
) -> Result<(), XlsxError> {
    let Some(filter) = filter else {
        return Ok(());
    };
    let reference = format!(
        "{}{}:{}{}",
        col_to_letter(filter.start_col),
        filter.start_row,
        col_to_letter(filter.end_col),
        filter.end_row
    );
    let mut block = BytesStart::new("autoFilter");
    block.push_attribute(("ref", reference.as_str()));
    if filter.columns.is_empty() {
        return writer
            .write_event(Event::Empty(block))
            .map_err(|error| XlsxError::InvalidData(error.to_string()));
    }
    writer
        .write_event(Event::Start(block))
        .map_err(|error| XlsxError::InvalidData(error.to_string()))?;

    for column in &filter.columns {
        let mut tag = BytesStart::new("filterColumn");
        tag.push_attribute(("colId", (column.field.saturating_sub(1)).to_string().as_str()));
        writer
            .write_event(Event::Start(tag))
            .map_err(|error| XlsxError::InvalidData(error.to_string()))?;

        let comparing = column
            .criteria
            .iter()
            .any(|criterion| filter_comparison(criterion).is_some());
        if comparing {
            let mut customs = BytesStart::new("customFilters");
            if !column.either {
                customs.push_attribute(("and", "1"));
            }
            writer
                .write_event(Event::Start(customs))
                .map_err(|error| XlsxError::InvalidData(error.to_string()))?;
            for criterion in &column.criteria {
                let (operator, value) = match filter_comparison(criterion) {
                    Some(split) => split,
                    None => ("equal", criterion.as_str()),
                };
                let mut custom = BytesStart::new("customFilter");
                custom.push_attribute(("operator", operator));
                custom.push_attribute(("val", value));
                writer
                    .write_event(Event::Empty(custom))
                    .map_err(|error| XlsxError::InvalidData(error.to_string()))?;
            }
            writer
                .write_event(Event::End(BytesEnd::new("customFilters")))
                .map_err(|error| XlsxError::InvalidData(error.to_string()))?;
        } else {
            writer
                .write_event(Event::Start(BytesStart::new("filters")))
                .map_err(|error| XlsxError::InvalidData(error.to_string()))?;
            for criterion in &column.criteria {
                let mut value = BytesStart::new("filter");
                value.push_attribute(("val", criterion.as_str()));
                writer
                    .write_event(Event::Empty(value))
                    .map_err(|error| XlsxError::InvalidData(error.to_string()))?;
            }
            writer
                .write_event(Event::End(BytesEnd::new("filters")))
                .map_err(|error| XlsxError::InvalidData(error.to_string()))?;
        }

        writer
            .write_event(Event::End(BytesEnd::new("filterColumn")))
            .map_err(|error| XlsxError::InvalidData(error.to_string()))?;
    }

    writer
        .write_event(Event::End(BytesEnd::new("autoFilter")))
        .map_err(|error| XlsxError::InvalidData(error.to_string()))
}

/// Splits a comparing criterion into the operator OOXML names it by and the
/// value it compares against. A criterion that only names a value answers None.
fn filter_comparison(criterion: &str) -> Option<(&'static str, &str)> {
    for (prefix, operator) in [
        ("<>", "notEqual"),
        (">=", "greaterThanOrEqual"),
        ("<=", "lessThanOrEqual"),
        (">", "greaterThan"),
        ("<", "lessThan"),
    ] {
        if let Some(rest) = criterion.strip_prefix(prefix) {
            return Some((operator, rest));
        }
    }
    None
}

/// Writes a `<mergeCells>` block, or nothing at all when there is none left.
fn write_merges(
    writer: &mut Writer<Cursor<Vec<u8>>>,
    merges: &[MergeCell],
) -> Result<(), XlsxError> {
    if merges.is_empty() {
        return Ok(());
    }
    let mut block = BytesStart::new("mergeCells");
    block.push_attribute(("count", merges.len().to_string().as_str()));
    writer
        .write_event(Event::Start(block))
        .map_err(|error| XlsxError::InvalidData(error.to_string()))?;
    for merge in merges {
        let reference = format!(
            "{}{}:{}{}",
            col_to_letter(merge.start_col),
            merge.start_row,
            col_to_letter(merge.end_col),
            merge.end_row
        );
        let mut span = BytesStart::new("mergeCell");
        span.push_attribute(("ref", reference.as_str()));
        writer
            .write_event(Event::Empty(span))
            .map_err(|error| XlsxError::InvalidData(error.to_string()))?;
    }
    writer
        .write_event(Event::End(BytesEnd::new("mergeCells")))
        .map_err(|error| XlsxError::InvalidData(error.to_string()))
}

/// What a column should end up as. A field left `None` keeps whatever the
/// sheet already says about it, so hiding a column does not disturb its width
/// and widening one does not un-hide it.
#[derive(Clone, Copy, Default, PartialEq)]
struct ColumnWant {
    hidden: Option<bool>,
    /// In characters of the workbook's default font, gutter included.
    width: Option<f32>,
}

/// Writes one `<col>` span, splitting it where a change covers only part of it.
///
/// A span reads `<col min="1" max="3" .../>` and covers three columns at once,
/// so hiding the middle one means writing three spans in its place.
fn write_col_span(
    writer: &mut Writer<Cursor<Vec<u8>>>,
    start: &BytesStart<'_>,
    pending: &mut BTreeMap<u32, ColumnWant>,
) -> Result<(), XlsxError> {
    let min = get_attr(start, "min")
        .and_then(|value| value.parse::<u32>().ok())
        .unwrap_or(1);
    let max = get_attr(start, "max")
        .and_then(|value| value.parse::<u32>().ok())
        .unwrap_or(min);
    let was_hidden = matches!(get_attr(start, "hidden").as_deref(), Some("1") | Some("true"));
    let was_width = get_attr(start, "width");
    let was_custom = matches!(
        get_attr(start, "customWidth").as_deref(),
        Some("1") | Some("true")
    );

    // Group the columns of this span into runs that share an answer.
    let mut runs: Vec<(u32, u32, ColumnFace)> = Vec::new();
    for column in min..=max {
        let want = pending.remove(&(column - 1)).unwrap_or_default();
        let face = ColumnFace {
            hidden: want.hidden.unwrap_or(was_hidden),
            width: want
                .width
                .map(width_text)
                .or_else(|| was_width.clone()),
            resized: want.width.is_some(),
            custom: want.width.is_some() || was_custom,
        };
        match runs.last_mut() {
            Some((_, last, held)) if *held == face && *last + 1 == column => *last = column,
            _ => runs.push((column, column, face)),
        }
    }

    for (first, last, face) in runs {
        write_col(writer, Some(start), first, last, &face)?;
    }
    Ok(())
}

/// What one column should look like once everything has been applied.
#[derive(Clone, PartialEq)]
struct ColumnFace {
    hidden: bool,
    width: Option<String>,
    /// True when this run's width was set here rather than read off the file.
    resized: bool,
    custom: bool,
}

/// A width as the file states it.
///
/// The shortest form that still reads back as the same number, which is what
/// Excel writes: 10.625 for a column someone typed 10 into, and a bare `9` for
/// one left at the default. Rounding to a fixed number of places would write
/// 10.63, a different width.
fn width_text(width: f32) -> String {
    format!("{width}")
}

/// Writes one `<col>` span, keeping whatever attributes `from` carried that
/// this span does not decide for itself — a style, an outline level, a best
/// fit that still stands.
fn write_col(
    writer: &mut Writer<Cursor<Vec<u8>>>,
    from: Option<&BytesStart<'_>>,
    first: u32,
    last: u32,
    face: &ColumnFace,
) -> Result<(), XlsxError> {
    let mut span = BytesStart::new("col");
    span.push_attribute(("min", first.to_string().as_str()));
    span.push_attribute(("max", last.to_string().as_str()));
    if let Some(width) = &face.width {
        span.push_attribute(("width", width.as_str()));
    }
    if let Some(start) = from {
        for attribute in start.attributes().flatten() {
            let key = String::from_utf8_lossy(attribute.key.as_ref()).into_owned();
            // Everything this span states for itself, and — when a width has
            // been set by hand — the claim that the old width was a fit to
            // something. Excel drops bestFit on a resize, measured.
            if matches!(
                key.as_str(),
                "min" | "max" | "width" | "hidden" | "customWidth"
            ) || (face.resized && key == "bestFit")
            {
                continue;
            }
            let value = String::from_utf8_lossy(&attribute.value).into_owned();
            span.push_attribute((key.as_str(), value.as_str()));
        }
    }
    if face.hidden {
        span.push_attribute(("hidden", "1"));
    }
    if face.custom {
        span.push_attribute(("customWidth", "1"));
    }
    writer
        .write_event(Event::Empty(span))
        .map_err(|error| XlsxError::InvalidData(error.to_string()))
}

/// Writes `<col>` spans for columns the sheet never described.
fn write_new_cols(
    writer: &mut Writer<Cursor<Vec<u8>>>,
    pending: &mut BTreeMap<u32, ColumnWant>,
) -> Result<(), XlsxError> {
    for (column, want) in std::mem::take(pending) {
        let hidden = want.hidden.unwrap_or(false);
        if !hidden && want.width.is_none() {
            continue;
        }
        let face = ColumnFace {
            hidden,
            width: want.width.map(width_text),
            resized: want.width.is_some(),
            custom: want.width.is_some(),
        };
        write_col(writer, None, column + 1, column + 1, &face)?;
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

/// Copies a start tag, pointing it at a different entry in the style sheet.
fn with_style(start: &BytesStart<'_>, style: u32) -> BytesStart<'static> {
    let name = String::from_utf8_lossy(start.name().as_ref()).into_owned();
    let mut rewritten = BytesStart::new(name);
    for attribute in start.attributes().flatten() {
        let key = String::from_utf8_lossy(attribute.key.as_ref()).into_owned();
        if key == "s" {
            continue;
        }
        let value = String::from_utf8_lossy(&attribute.value).into_owned();
        rewritten.push_attribute((key.as_str(), value.as_str()));
    }
    rewritten.push_attribute(("s", style.to_string().as_str()));
    rewritten
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
    styles: &BTreeMap<(u32, u32), u32>,
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
        write_cell_value(writer, row, col, &value, styles)?;
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
    /// Zero-based column -> what it should become.
    cols: &'a BTreeMap<u32, ColumnWant>,
    /// Every merge the sheet should end up with, when they are being replaced.
    merges: Option<&'a [MergeCell]>,
    /// (one-based row, zero-based column) -> the cellXfs index it should carry.
    styles: &'a BTreeMap<(u32, u32), u32>,
    /// The filter the sheet should end up under, when it is being replaced.
    /// `Some(None)` takes an existing filter away.
    filter: Option<&'a Option<AutoFilter>>,
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
    let mut pending_cols: BTreeMap<u32, ColumnWant> = sheet_edits.cols.clone();
    let mut seen_cols = false;
    let mut skip_merges_depth = 0_u32;
    // Whether the sheet already describes its merges decides where the new
    // block goes: in place of the old one, or straight after the data.
    let has_merge_block = xml.contains("<mergeCells");
    let mut wrote_merges = false;
    let has_filter_block = xml.contains("<autoFilter");
    let mut wrote_filter = false;
    let mut skip_filter_depth = 0_u32;

    let mut current_row: u32 = 0;
    let mut in_row = false;
    let mut in_cell = false;
    let mut cell_col: u32;
    let mut cell_row: u32;
    let mut in_value = false;
    let mut value_written = false;
    let mut current_edit: Option<CellEditValue> = None;
    let mut skip_value_text = false;
    let mut skip_replaced_content_depth = 0_u32;

    loop {
        match reader.read_event().map_err(XlsxError::Xml)? {
            Event::Eof => break,
            Event::Start(ref e) => {
                if skip_filter_depth > 0 {
                    skip_filter_depth += 1;
                    continue;
                }
                if skip_merges_depth > 0 {
                    skip_merges_depth += 1;
                    continue;
                }
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
                if name == "autoFilter" && sheet_edits.filter.is_some() {
                    skip_filter_depth = 1;
                    continue;
                }
                if name == "mergeCells" && sheet_edits.merges.is_some() {
                    skip_merges_depth = 1;
                    continue;
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
                                write_inserted_row(&mut writer, row, cells, hidden, sheet_edits.styles)?;
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
                        // Whether this cell has been GIVEN its new value yet,
                        // which is not the same question as whether we are
                        // inside its `<v>` right now. Asking `in_value` at
                        // `</c>` — by which point `</v>` has already cleared
                        // it — put a SECOND `<v>` after the one just written,
                        // and a cell holding two of them is a workbook Excel
                        // refuses to open. It is the same conflation that
                        // wrote a docx run's text once per `<w:t>`.
                        value_written = false;
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
                                        sheet_edits.styles,
                                    )?;
                                }
                            }
                            current_edit = cells.remove(&cell_col);
                        } else {
                            current_edit = None;
                        }

                        let restyle = sheet_edits.styles.get(&(cell_row, cell_col)).copied();
                        if current_edit.is_none() {
                            if let Some(style) = restyle {
                                writer
                                    .write_event(Event::Start(with_style(e, style)))
                                    .map_err(|e| XlsxError::InvalidData(e.to_string()))?;
                                continue;
                            }
                        }
                        if let Some(value) = &current_edit {
                            // Rewrite the cell type while preserving its reference and style.
                            let mut new_start = BytesStart::new("c");
                            for attr in e.attributes().flatten() {
                                let key = std::str::from_utf8(attr.key.as_ref()).unwrap_or("");
                                if key == "t" || (key == "s" && restyle.is_some()) {
                                    continue;
                                }
                                new_start.push_attribute((key, std::str::from_utf8(&attr.value).unwrap_or("")));
                            }
                            if let Some(style) = restyle {
                                new_start.push_attribute(("s", style.to_string().as_str()));
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
                        value_written = true;
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
                if skip_filter_depth > 0 {
                    skip_filter_depth -= 1;
                    if skip_filter_depth == 0 {
                        if let Some(filter) = sheet_edits.filter {
                            write_auto_filter(&mut writer, filter.as_ref())?;
                            wrote_filter = true;
                        }
                    }
                    continue;
                }
                if skip_merges_depth > 0 {
                    skip_merges_depth -= 1;
                    if skip_merges_depth == 0 {
                        if let Some(merges) = sheet_edits.merges {
                            write_merges(&mut writer, merges)?;
                            wrote_merges = true;
                        }
                    }
                    continue;
                }
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
                                        sheet_edits.styles,
                                )?;
                            }
                        }
                        in_row = false;
                    }
                    "c" => {
                        if in_cell && current_edit.is_some() {
                            // A cell that carried no `<v>` of its own still
                            // needs one to hold the new value.
                            if !value_written {
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
                            write_inserted_row(&mut writer, row, cells, hidden, sheet_edits.styles)?;
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
                                        sheet_edits.styles,
                                )?;
                            }
                        }
                    }
                    _ => {}
                }
                writer.write_event(Event::End(e.clone())).map_err(|e| XlsxError::InvalidData(e.to_string()))?;
                // A filter sits between the data and the merges.
                if name == "sheetData" && !wrote_filter && !has_filter_block {
                    if let Some(filter) = sheet_edits.filter {
                        write_auto_filter(&mut writer, filter.as_ref())?;
                        wrote_filter = true;
                    }
                }
                // A sheet that never had a mergeCells element needs one, and it
                // belongs directly after the data.
                if name == "sheetData" && !wrote_merges && !has_merge_block {
                    if let Some(merges) = sheet_edits.merges {
                        write_merges(&mut writer, merges)?;
                        wrote_merges = true;
                    }
                }
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
                if name == "autoFilter" && sheet_edits.filter.is_some() {
                    if let Some(filter) = sheet_edits.filter {
                        write_auto_filter(&mut writer, filter.as_ref())?;
                        wrote_filter = true;
                    }
                    continue;
                }
                if name == "mergeCells" && sheet_edits.merges.is_some() {
                    if let Some(merges) = sheet_edits.merges {
                        write_merges(&mut writer, merges)?;
                        wrote_merges = true;
                    }
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
                        write_inserted_row(&mut writer, row, cells, hidden, sheet_edits.styles)?;
                    }
                    for row in std::mem::take(&mut pending_rows) {
                        if let Some(true) = sheet_edits.rows.get(&row) {
                            write_inserted_row(&mut writer, row, BTreeMap::new(), Some(true), sheet_edits.styles)?;
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
                            write_inserted_row(&mut writer, row, cells, hidden, sheet_edits.styles)?;
                        }
                    }
                    pending_rows.remove(&row_num);
                    if let Some(cells) = pending.remove(&row_num) {
                        writer.write_event(Event::Start(e.clone())).map_err(|error| {
                            XlsxError::InvalidData(error.to_string())
                        })?;
                        for (col, value) in cells {
                            write_cell_value(&mut writer, row_num, col, &value, sheet_edits.styles)?;
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
                                    sheet_edits.styles,
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

    const STYLES: &str = concat!(
        r#"<styleSheet><numFmts count="0"/><fonts count="1"><font><sz val="11"/></font></fonts>"#,
        r#"<fills count="1"><fill><patternFill patternType="none"/></fill></fills>"#,
        r#"<borders count="1"><border/></borders>"#,
        r#"<cellXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/></cellXfs>"#,
        r#"</styleSheet>"#
    );

    #[test]
    fn styles_are_appended_and_their_indices_reported() {
        let bold = CellStyle {
            bold: true,
            ..CellStyle::default()
        };
        let (patched, indices) = patch_styles_xml(STYLES, &[bold]).expect("appends");

        assert_eq!(indices, vec![1]);
        assert!(patched.contains("<font><b/></font>"));
        assert!(patched.contains(r#"<fonts count="2">"#));
        assert!(patched.contains(r#"<cellXfs count="2">"#));
        assert!(patched.contains(r#"fontId="1" fillId="1" borderId="1""#));
        // The entry that was already there is left alone.
        assert!(patched.contains(r#"<font><sz val="11"/></font>"#));
    }

    #[test]
    fn a_custom_number_format_gets_an_id_of_its_own() {
        let styled = CellStyle {
            number_format: Some("0.000".to_string()),
            ..CellStyle::default()
        };
        let (patched, _) = patch_styles_xml(STYLES, &[styled]).expect("appends");
        assert!(patched.contains(r#"<numFmt numFmtId="164" formatCode="0.000"/>"#));
        assert!(patched.contains(r#"numFmtId="164""#));
        assert!(patched.contains(r#"<numFmts count="1">"#));
    }

    #[test]
    fn a_well_known_number_format_reuses_its_id() {
        let styled = CellStyle {
            number_format: Some("0.00".to_string()),
            ..CellStyle::default()
        };
        let (patched, _) = patch_styles_xml(STYLES, &[styled]).expect("appends");
        assert!(!patched.contains("<numFmt "));
        assert!(patched.contains(r#"numFmtId="2""#));
    }

    #[test]
    fn alignment_and_colour_reach_the_style_sheet() {
        let styled = CellStyle {
            horizontal_align: Some("center".to_string()),
            bg_color: Some("FFFF00".to_string()),
            font_color: Some("FF0000".to_string()),
            border_top: Some(BorderLine { style: "thin".to_string(), color: None }),
            ..CellStyle::default()
        };
        let (patched, _) = patch_styles_xml(STYLES, &[styled]).expect("appends");
        assert!(patched.contains(r#"<alignment horizontal="center"/>"#));
        assert!(patched.contains(r#"<fgColor rgb="FFFFFF00"/>"#));
        assert!(patched.contains(r#"<color rgb="FFFF0000"/>"#));
        assert!(patched.contains(r#"<top style="thin">"#));
    }

    #[test]
    fn asking_for_no_styles_changes_nothing() {
        let (patched, indices) = patch_styles_xml(STYLES, &[]).expect("does nothing");
        assert_eq!(patched, STYLES);
        assert!(indices.is_empty());
    }

    /// The patcher takes everything a sheet changes; these tests only change cells.
    fn cells_only<'a>(
        cells: &'a HashMap<(u32, u32), &'a CellEditValue>,
    ) -> SheetEdits<'a> {
        static NOTHING: std::sync::OnceLock<BTreeMap<u32, bool>> = std::sync::OnceLock::new();
        let nothing = NOTHING.get_or_init(BTreeMap::new);
        static NO_STYLES: std::sync::OnceLock<BTreeMap<(u32, u32), u32>> =
            std::sync::OnceLock::new();
        static NO_COLS: std::sync::OnceLock<BTreeMap<u32, ColumnWant>> =
            std::sync::OnceLock::new();
        SheetEdits {
            cells,
            rows: nothing,
            cols: NO_COLS.get_or_init(BTreeMap::new),
            merges: None,
            styles: NO_STYLES.get_or_init(BTreeMap::new),
            filter: None,
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
    fn the_calculation_chain_is_not_left_pointing_at_nothing() {
        // Writing a value over a formula takes the formula away. A chain still
        // naming that cell makes Excel refuse the WHOLE workbook — `fies_t2`
        // has exactly one entry, C1, and replacing C1's `<f>C9</f>` lost the
        // file. The part is dropped, and so is anything that declares it.
        let types = concat!(
            r#"<Types><Override PartName="/xl/workbook.xml" ContentType="a"/>"#,
            r#"<Override PartName="/xl/calcChain.xml" ContentType="b"/>"#,
            r#"<Override PartName="/xl/styles.xml" ContentType="c"/></Types>"#,
        );
        let out = without_calc_chain(types);
        assert!(!out.contains("calcChain"), "{out}");
        assert!(out.contains("workbook.xml") && out.contains("styles.xml"), "{out}");
        let rels = concat!(
            r#"<Relationships><Relationship Id="rId1" Target="worksheets/sheet1.xml"/>"#,
            r#"<Relationship Id="rId8" Target="calcChain.xml"/></Relationships>"#,
        );
        let out = without_calc_chain(rels);
        assert!(!out.contains("calcChain"), "{out}");
        assert!(out.contains("rId1"), "{out}");
    }

    #[test]
    fn a_replaced_cell_keeps_one_value_element() {
        // The check for "this cell had no `<v>` of its own" was reading the
        // flag that says "we are inside a `<v>` right now", which `</v>` had
        // already cleared — so the new value went in twice, and Excel refuses
        // to open a workbook whose cell holds two of them. `001250007` was
        // that workbook: 83 sheets, and the whole file would not load.
        let xml = concat!(
            r#"<worksheet><sheetData><row r="2">"#,
            r#"<c r="B2" s="1" t="s"><v>877</v></c>"#,
            r#"</row></sheetData></worksheet>"#,
        );
        let value = CellEditValue::String("fresh".to_string());
        let edits = HashMap::from([((2, 1), &value)]);
        let patched = patch_worksheet_xml(xml, &cells_only(&edits)).expect("should patch");
        assert_eq!(patched.matches("<v>").count(), 1, "{patched}");
        assert!(patched.contains("<v>fresh</v>"), "{patched}");
        assert!(!patched.contains("877"), "{patched}");
    }

    /// The edits for one sheet's columns and nothing else.
    fn cols_only(cols: &BTreeMap<u32, ColumnWant>) -> SheetEdits<'_> {
        static NOTHING: std::sync::OnceLock<BTreeMap<u32, bool>> = std::sync::OnceLock::new();
        static NO_CELLS: std::sync::OnceLock<HashMap<(u32, u32), &CellEditValue>> =
            std::sync::OnceLock::new();
        static NO_STYLES: std::sync::OnceLock<BTreeMap<(u32, u32), u32>> =
            std::sync::OnceLock::new();
        SheetEdits {
            cells: NO_CELLS.get_or_init(HashMap::new),
            rows: NOTHING.get_or_init(BTreeMap::new),
            cols,
            merges: None,
            styles: NO_STYLES.get_or_init(BTreeMap::new),
            filter: None,
        }
    }

    /// The `<cols>` element of a patched sheet, or what stands in its place.
    fn cols_of(xml: &str, cols: &BTreeMap<u32, ColumnWant>) -> String {
        let patched = patch_worksheet_xml(xml, &cols_only(cols)).expect("should patch");
        match (patched.find("<cols>"), patched.find("</cols>")) {
            (Some(from), Some(to)) => patched[from..to + "</cols>".len()].to_string(),
            _ => format!("(no cols in {patched})"),
        }
    }

    const BARE: &str =
        r#"<worksheet><sheetData><row r="1"><c r="A1"><v>1</v></c></row></sheetData></worksheet>"#;

    #[test]
    fn a_width_is_written_where_the_sheet_had_no_columns_at_all() {
        // A sheet nobody has resized carries no `<cols>` element — measured on
        // a freshly saved workbook — so one has to be put in, before
        // `<sheetData>` where the schema wants it.
        let want = BTreeMap::from([(1, ColumnWant { hidden: None, width: Some(10.625) })]);
        assert_eq!(
            cols_of(BARE, &want),
            r#"<cols><col min="2" max="2" width="10.625" customWidth="1"/></cols>"#,
        );
    }

    #[test]
    fn a_width_always_brings_the_flag_that_says_it_was_chosen() {
        // Excel writes `customWidth="1"` beside every width it stores. Without
        // it the number is there and Excel ignores it.
        let want = BTreeMap::from([(0, ColumnWant { hidden: None, width: Some(30.625) })]);
        let out = cols_of(BARE, &want);
        assert!(out.contains(r#"width="30.625""#), "{out}");
        assert!(out.contains(r#"customWidth="1""#), "{out}");
    }

    #[test]
    fn resizing_one_column_of_a_run_splits_the_run() {
        // `A:E` at one width, then C widened, comes back as three spans —
        // 1-2, 3-3, 4-5 — which is what Excel itself writes.
        let xml = concat!(
            r#"<worksheet><cols><col min="1" max="5" width="12.625" customWidth="1"/></cols>"#,
            r#"<sheetData/></worksheet>"#,
        );
        let want = BTreeMap::from([(2, ColumnWant { hidden: None, width: Some(30.625) })]);
        assert_eq!(
            cols_of(xml, &want),
            concat!(
                r#"<cols><col min="1" max="2" width="12.625" customWidth="1"/>"#,
                r#"<col min="3" max="3" width="30.625" customWidth="1"/>"#,
                r#"<col min="4" max="5" width="12.625" customWidth="1"/></cols>"#,
            ),
        );
    }

    #[test]
    fn resizing_a_column_drops_the_claim_that_it_was_a_best_fit() {
        // An autofitted column carries `bestFit="1"`. Excel takes that off
        // when a width is set by hand, which is the sort of thing that has to
        // be looked at rather than reasoned about — measured.
        let xml = concat!(
            r#"<worksheet><cols>"#,
            r#"<col min="2" max="2" width="7.5" bestFit="1" customWidth="1"/>"#,
            r#"</cols><sheetData/></worksheet>"#,
        );
        let want = BTreeMap::from([(1, ColumnWant { hidden: None, width: Some(20.625) })]);
        let out = cols_of(xml, &want);
        assert!(!out.contains("bestFit"), "{out}");
        assert!(out.contains(r#"width="20.625""#), "{out}");
    }

    #[test]
    fn hiding_a_column_leaves_its_width_alone_and_the_other_way_round() {
        let xml = concat!(
            r#"<worksheet><cols>"#,
            r#"<col min="2" max="2" width="18.625" customWidth="1"/>"#,
            r#"</cols><sheetData/></worksheet>"#,
        );
        let hide = BTreeMap::from([(1, ColumnWant { hidden: Some(true), width: None })]);
        let out = cols_of(xml, &hide);
        assert!(out.contains(r#"width="18.625""#), "{out}");
        assert!(out.contains(r#"hidden="1""#), "{out}");

        let widen = BTreeMap::from([(1, ColumnWant { hidden: None, width: Some(9.0) })]);
        let out = cols_of(
            &xml.replace(r#"width="18.625""#, r#"width="18.625" hidden="1""#),
            &widen,
        );
        assert!(out.contains(r#"hidden="1""#), "{out}");
        assert!(out.contains(r#"width="9""#), "{out}");
    }

    #[test]
    fn a_column_keeps_whatever_else_its_span_was_carrying() {
        // A `<col>` can dress its column in a style and set an outline level.
        // Neither is this edit's business, and losing them would change how
        // every cell in the column is drawn.
        let xml = concat!(
            r#"<worksheet><cols>"#,
            r#"<col min="3" max="3" width="8" style="4" outlineLevel="2"/>"#,
            r#"</cols><sheetData/></worksheet>"#,
        );
        let want = BTreeMap::from([(2, ColumnWant { hidden: None, width: Some(15.5) })]);
        let out = cols_of(xml, &want);
        assert!(out.contains(r#"style="4""#), "{out}");
        assert!(out.contains(r#"outlineLevel="2""#), "{out}");
        assert!(out.contains(r#"width="15.5""#), "{out}");
    }

    #[test]
    fn columns_asked_for_the_same_width_are_written_as_one_span() {
        let want = BTreeMap::from([
            (0, ColumnWant { hidden: None, width: Some(12.0) }),
            (1, ColumnWant { hidden: None, width: Some(12.0) }),
            (2, ColumnWant { hidden: None, width: Some(12.0) }),
        ]);
        let out = cols_of(BARE, &want);
        assert_eq!(out.matches("<col ").count(), 3, "{out}");
        // Columns the sheet never described are written one at a time; the
        // joining-up happens where a span already covers them.
        let xml = r#"<worksheet><cols><col min="1" max="3" width="8"/></cols><sheetData/></worksheet>"#;
        let out = cols_of(xml, &want);
        assert_eq!(
            out,
            r#"<cols><col min="1" max="3" width="12" customWidth="1"/></cols>"#,
        );
    }

    #[test]
    fn a_cell_written_for_the_first_time_wears_the_style_it_was_given() {
        // A cell the sheet already had gets its style put on as it is copied
        // through. A cell that did not exist goes down the insert path, which
        // did not consult the style map at all — so it came out under General
        // however it had been styled, and a date filled into an empty cell
        // landed in the file as a bare serial number.
        let xml = r#"<worksheet><sheetData><row r="1"><c r="A1"><v>1</v></c></row></sheetData></worksheet>"#;
        let fresh = CellEditValue::Number(46053.0);
        let over = CellEditValue::Number(7.0);
        let edits = HashMap::from([((1, 1), &fresh), ((1, 0), &over)]);
        let rows = BTreeMap::new();
        let styles = BTreeMap::from([((1, 1), 4u32), ((1, 0), 5u32)]);
        let patched = patch_worksheet_xml(xml, &SheetEdits {
            cells: &edits,
            rows: &rows,
            cols: &BTreeMap::new(),
            merges: None,
            styles: &styles,
            filter: None,
        }).expect("should patch");
        assert!(patched.contains(r#"<c r="B1" s="4""#), "{patched}");
        assert!(patched.contains(r#"<c r="A1" s="5""#), "{patched}");
    }

    #[test]
    fn a_row_written_for_the_first_time_carries_its_cells_styles() {
        let xml = r#"<worksheet><sheetData><row r="1"><c r="A1"><v>1</v></c></row></sheetData></worksheet>"#;
        let fresh = CellEditValue::Number(46054.0);
        let edits = HashMap::from([((3, 0), &fresh)]);
        let rows = BTreeMap::new();
        let styles = BTreeMap::from([((3, 0), 2u32)]);
        let patched = patch_worksheet_xml(xml, &SheetEdits {
            cells: &edits,
            rows: &rows,
            cols: &BTreeMap::new(),
            merges: None,
            styles: &styles,
            filter: None,
        }).expect("should patch");
        assert!(patched.contains(r#"<c r="A3" s="2""#), "{patched}");
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
