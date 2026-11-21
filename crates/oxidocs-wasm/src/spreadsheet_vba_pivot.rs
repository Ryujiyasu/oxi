// SPDX-License-Identifier: MIT OR Apache-2.0

//! A macro's pivot tables: `Worksheet.PivotTables`, the PivotTable, its
//! PivotCache and its fields, `Range.PivotTable` and `Workbook.RefreshAll`.
//!
//! The tables are the ones the package defines (`Workbook::pivots`), laid out
//! by `oxicells_core::pivot::compute` over the source as it stands. What each
//! member answers was read off Excel 16.0 on a Japanese Office through COM
//! (`tools/metrics/_xlsx_vba_pivot_probe.py`, 2026-11-21): field orientation
//! xlRowField 1 / xlColumnField 2 / xlPageField 3 / xlDataField 4 /
//! xlHidden 0, Position counted from 1 along its axis, a value field's
//! Function -4157 for a sum, CurrentPage "(All)", RefreshTable True and
//! PivotCache.Refresh Empty, SourceData "D!R1C1:R5C5", RecordCount the data
//! rows, TableRange2 the page fields included and TableRange1 not.

use super::*;
use oxicells_core::pivot::{self, Aggregate, PivotSpec};

/// Which fields a field collection holds.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub(super) enum PivotAxis {
    All,
    Rows,
    Columns,
    Pages,
    Data,
}

/// A field of a table: one of the source's columns, or one of its value
/// fields.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub(super) enum PivotFieldRef {
    Source(usize),
    Data(usize),
}

const XL_HIDDEN: i32 = 0;
const XL_ROW_FIELD: i32 = 1;
const XL_COLUMN_FIELD: i32 = 2;
const XL_PAGE_FIELD: i32 = 3;
const XL_DATA_FIELD: i32 = 4;

fn function_of(aggregate: Aggregate) -> i32 {
    match aggregate {
        Aggregate::Sum => -4157,
        Aggregate::Count => -4112,
        Aggregate::Average => -4106,
        Aggregate::Max => -4136,
        Aggregate::Min => -4139,
        Aggregate::Product => -4149,
        Aggregate::CountNums => -4113,
        Aggregate::StdDev => -4155,
        Aggregate::StdDevP => -4156,
        Aggregate::Var => -4164,
        Aggregate::VarP => -4165,
    }
}

fn r1c1_range(sheet: &str, (r0, c0, r1, c1): (u32, u32, u32, u32)) -> String {
    format!("{}!R{}C{}:R{}C{}", quoted_sheet(sheet), r0 + 1, c0 + 1, r1 + 1, c1 + 1)
}

impl WorkbookHost<'_> {
    pub(super) fn pivot_object(&self, object: &ObjectRef) -> Option<HostObject> {
        match self.objects.get(host_object_index(object.handle)) {
            Some(
                held @ (HostObject::PivotTables(_)
                | HostObject::PivotTable(_)
                | HostObject::PivotCache(_)
                | HostObject::PivotFields(..)
                | HostObject::PivotField(..)
                | HostObject::PivotCaches
                | HostObject::NewPivotCache(_)),
            ) => Some(*held),
            _ => None,
        }
    }

    fn sheet_index_named(&self, name: &str) -> Option<usize> {
        self.workbook.sheets.iter().position(|sheet| sheet.name.eq_ignore_ascii_case(name))
    }

    /// The tables standing on a sheet, as indices into the workbook's list.
    fn pivots_on(&self, sheet: usize) -> Vec<usize> {
        let name = &self.workbook.sheets[sheet].name;
        self.workbook
            .pivots
            .iter()
            .enumerate()
            .filter(|(_, spec)| spec.sheet.eq_ignore_ascii_case(name))
            .map(|(index, _)| index)
            .collect()
    }

    /// The cells a table covers, zero-based (row, col, row, col), with its page
    /// fields (`TableRange2`) or without them (`TableRange1`).
    fn pivot_extent(spec: &PivotSpec, with_pages: bool) -> (u32, u32, u32, u32) {
        let pages = spec.pages.len() as u32;
        let top = if with_pages && pages > 0 { spec.origin.0.saturating_sub(pages + 1) } else { spec.origin.0 };
        (top, spec.origin.1, spec.extent.0.max(spec.origin.0), spec.extent.1.max(spec.origin.1))
    }

    /// The table a cell stands in, if any.
    pub(super) fn pivot_at(&self, at: CellAddress) -> Option<usize> {
        let (row, col) = (at.row.saturating_sub(1), at.column);
        self.pivots_on(at.sheet).into_iter().find(|index| {
            let (r0, c0, r1, c1) = Self::pivot_extent(&self.workbook.pivots[*index], true);
            (r0..=r1).contains(&row) && (c0..=c1).contains(&col)
        })
    }

    fn pivot_range(&mut self, index: usize, with_pages: bool) -> Result<Value, String> {
        let spec = &self.workbook.pivots[index];
        let sheet = self.sheet_index_named(&spec.sheet).ok_or_else(|| host_error(1004, "the pivot table's sheet is gone"))?;
        let (r0, c0, r1, c1) = Self::pivot_extent(spec, with_pages);
        Ok(self.object(HostObject::Range(CellRange {
            sheet,
            start_row: r0 + 1,
            start_column: c0,
            end_row: r1 + 1,
            end_column: c1,
        })))
    }

    /// Lay a table out again over its source as it stands now and write it to
    /// its sheet, clearing what it covered before.
    pub(super) fn refresh_pivot(&mut self, index: usize) -> Result<(), String> {
        let spec = self.workbook.pivots[index].clone();
        let target = self.sheet_index_named(&spec.sheet).ok_or_else(|| host_error(1004, "the pivot table's sheet is gone"))?;
        let grid = {
            let get = pivot::source_reader(&*self.workbook, &spec.source_sheet);
            pivot::compute(&spec, &get, pivot::Ui::Japanese)
        };
        let (r0, c0, r1, c1) = Self::pivot_extent(&spec, true);
        for row in r0..=r1 {
            for column in c0..=c1 {
                self.set_cell_value_with_revision(CellAddress { sheet: target, row: row + 1, column }, CellValue::Empty, true)?;
            }
        }
        let mut bottom = grid.origin.0;
        let mut right = grid.origin.1;
        for (i, line) in grid.cells.iter().enumerate() {
            for (j, value) in line.iter().enumerate() {
                let (row, column) = (grid.origin.0 + i as u32, grid.origin.1 + j as u32);
                bottom = bottom.max(row);
                right = right.max(column);
                if !matches!(value, CellValue::Empty) {
                    self.set_cell_value_with_revision(CellAddress { sheet: target, row: row + 1, column }, value.clone(), true)?;
                }
            }
        }
        self.workbook.pivots[index].extent = (bottom, right);
        Ok(())
    }

    /// Every table in the book laid out again (`RefreshAll`).
    pub(super) fn refresh_all_pivots(&mut self) -> Result<(), String> {
        for index in 0..self.workbook.pivots.len() {
            self.refresh_pivot(index)?;
        }
        Ok(())
    }

    fn pivot_item(&mut self, sheet: usize, wanted: &Value) -> Result<Value, String> {
        let on = self.pivots_on(sheet);
        let found = match wanted {
            Value::String(name) => on.into_iter().find(|index| self.workbook.pivots[*index].name.eq_ignore_ascii_case(name)),
            other => {
                let n = any_number(other).ok_or_else(|| host_error(13, "Type mismatch"))?;
                (n >= 1.0).then(|| on.get(n as usize - 1).copied()).flatten()
            }
        };
        match found {
            Some(index) => Ok(self.object(HostObject::PivotTable(index))),
            None => Err(host_error(1004, "Unable to get the PivotTables property of the Worksheet class")),
        }
    }

    /// The fields an axis holds, in order.
    fn pivot_fields_of(spec: &PivotSpec, axis: PivotAxis) -> Vec<PivotFieldRef> {
        let source = |list: &[Option<usize>]| list.iter().flatten().map(|f| PivotFieldRef::Source(*f)).collect();
        match axis {
            PivotAxis::All => (0..spec.field_names.len()).map(PivotFieldRef::Source).collect(),
            PivotAxis::Rows => source(&spec.rows),
            PivotAxis::Columns => source(&spec.cols),
            PivotAxis::Pages => spec.pages.iter().map(|(f, _)| PivotFieldRef::Source(*f)).collect(),
            PivotAxis::Data => (0..spec.data.len()).map(PivotFieldRef::Data).collect(),
        }
    }

    fn pivot_field_name(spec: &PivotSpec, field: PivotFieldRef) -> String {
        match field {
            PivotFieldRef::Source(f) => spec.field_names.get(f).cloned().unwrap_or_default(),
            PivotFieldRef::Data(d) => spec.data.get(d).map(|data| data.name.clone()).unwrap_or_default(),
        }
    }

    fn pivot_field_item(&mut self, index: usize, axis: PivotAxis, wanted: &Value) -> Result<Value, String> {
        let spec = &self.workbook.pivots[index];
        let fields = Self::pivot_fields_of(spec, axis);
        let found = match wanted {
            Value::String(name) => {
                let named = |field: &PivotFieldRef| Self::pivot_field_name(spec, *field).eq_ignore_ascii_case(name);
                fields.iter().copied().find(named).or_else(|| {
                    // PivotFields("合計 / Sales") finds a value field as well.
                    (axis == PivotAxis::All)
                        .then(|| Self::pivot_fields_of(spec, PivotAxis::Data).into_iter().find(named))
                        .flatten()
                })
            }
            other => {
                let n = any_number(other).ok_or_else(|| host_error(13, "Type mismatch"))?;
                (n >= 1.0).then(|| fields.get(n as usize - 1).copied()).flatten()
            }
        };
        match found {
            Some(field) => Ok(self.object(HostObject::PivotField(index, field))),
            None => Err(host_error(1004, "Unable to get the PivotFields property of the PivotTable class")),
        }
    }

    fn pivot_orientation(spec: &PivotSpec, field: PivotFieldRef) -> (i32, i32) {
        match field {
            PivotFieldRef::Data(d) => (XL_DATA_FIELD, d as i32 + 1),
            PivotFieldRef::Source(f) => {
                let at = |list: &[Option<usize>]| list.iter().flatten().position(|g| *g == f);
                if let Some(p) = at(&spec.rows) {
                    (XL_ROW_FIELD, p as i32 + 1)
                } else if let Some(p) = at(&spec.cols) {
                    (XL_COLUMN_FIELD, p as i32 + 1)
                } else if let Some(p) = spec.pages.iter().position(|(g, _)| *g == f) {
                    (XL_PAGE_FIELD, p as i32 + 1)
                } else {
                    (XL_HIDDEN, 0)
                }
            }
        }
    }

    /// Get or call a member of a pivot object; `None` for a name it lacks.
    pub(super) fn pivot_member(&mut self, held: HostObject, name: &str, args: &[Value]) -> Result<Option<Value>, String> {
        let lower = name.to_ascii_lowercase();
        match held {
            HostObject::PivotTables(sheet) => match (lower.as_str(), args) {
                ("count", []) => Ok(Some(Value::Integer(self.pivots_on(sheet).len() as i64))),
                ("item", [wanted]) | ("_default", [wanted]) => self.pivot_item(sheet, wanted).map(Some),
                ("parent", []) => Ok(Some(self.object(HostObject::Worksheet(sheet)))),
                _ => Ok(None),
            },
            HostObject::PivotTable(index) => {
                let spec = self.workbook.pivots[index].clone();
                let axis = match lower.as_str() {
                    "pivotfields" => Some(PivotAxis::All),
                    "rowfields" => Some(PivotAxis::Rows),
                    "columnfields" => Some(PivotAxis::Columns),
                    "pagefields" => Some(PivotAxis::Pages),
                    "datafields" => Some(PivotAxis::Data),
                    _ => None,
                };
                if let Some(axis) = axis {
                    return match args {
                        [] => Ok(Some(self.object(HostObject::PivotFields(index, axis)))),
                        [wanted] => self.pivot_field_item(index, axis, wanted).map(Some),
                        _ => Err(host_error(450, "Wrong number of arguments")),
                    };
                }
                Ok(Some(match lower.as_str() {
                    "name" | "_default" => Value::String(spec.name.clone()),
                    "tablerange1" => self.pivot_range(index, false)?,
                    "tablerange2" => self.pivot_range(index, true)?,
                    "refreshtable" => {
                        self.refresh_pivot(index)?;
                        self.recalculate();
                        Value::Boolean(true)
                    }
                    "update" => {
                        self.refresh_pivot(index)?;
                        self.recalculate();
                        Value::Empty
                    }
                    "pivotcache" => self.object(HostObject::PivotCache(index)),
                    // AddDataField(Field[, Caption][, Function]): a value field
                    // over a source column; its caption Excel's own
                    // ("合計 / Sales") unless one is given.
                    "adddatafield" => {
                        let Some(field) = args.first().and_then(|held| match held {
                            Value::Object(object) => match self.pivot_object(object) {
                                Some(HostObject::PivotField(i, PivotFieldRef::Source(f))) if i == index => Some(f),
                                _ => None,
                            },
                            _ => None,
                        }) else {
                            return Err(host_error(1004, "AddDataField takes a field of this table"));
                        };
                        let aggregate = match args.get(2).and_then(any_number).map(|n| n as i32) {
                            Some(-4112) => Aggregate::Count,
                            Some(-4106) => Aggregate::Average,
                            Some(-4136) => Aggregate::Max,
                            Some(-4139) => Aggregate::Min,
                            Some(-4149) => Aggregate::Product,
                            Some(-4113) => Aggregate::CountNums,
                            Some(-4155) => Aggregate::StdDev,
                            Some(-4156) => Aggregate::StdDevP,
                            Some(-4164) => Aggregate::Var,
                            Some(-4165) => Aggregate::VarP,
                            _ => Aggregate::Sum,
                        };
                        let caption = match args.get(1) {
                            Some(Value::String(text)) if !text.is_empty() => text.clone(),
                            _ => aggregate.caption(pivot::Ui::Japanese, &spec.field_names[field]),
                        };
                        let spec = &mut self.workbook.pivots[index];
                        spec.data.push(pivot::DataField { name: caption, field, aggregate });
                        // A second value field puts the values across the
                        // columns, as Excel's own Values field does.
                        if spec.data.len() == 2 && !spec.rows.contains(&None) && !spec.cols.contains(&None) {
                            spec.cols.push(None);
                        }
                        let d = self.workbook.pivots[index].data.len() - 1;
                        self.refresh_pivot(index)?;
                        self.recalculate();
                        self.object(HostObject::PivotField(index, PivotFieldRef::Data(d)))
                    }
                    "sourcedata" => Value::String(r1c1_range(&spec.source_sheet, spec.source)),
                    "rowgrand" => Value::Boolean(spec.row_grand_totals),
                    "columngrand" => Value::Boolean(spec.col_grand_totals),
                    "parent" => match self.sheet_index_named(&spec.sheet) {
                        Some(sheet) => self.object(HostObject::Worksheet(sheet)),
                        None => return Err(host_error(1004, "the pivot table's sheet is gone")),
                    },
                    _ => return Ok(None),
                }))
            }
            HostObject::PivotCaches => match (lower.as_str(), args) {
                ("count", []) => Ok(Some(Value::Integer(self.workbook.pivots.len() as i64))),
                ("item", [wanted]) | ("_default", [wanted]) => {
                    let n = any_number(wanted).ok_or_else(|| host_error(13, "Type mismatch"))?;
                    if n >= 1.0 && (n as usize) <= self.workbook.pivots.len() {
                        Ok(Some(self.object(HostObject::PivotCache(n as usize - 1))))
                    } else {
                        Err(host_error(1004, "Unable to get the Item property of the PivotCaches class"))
                    }
                }
                // Create(SourceType, SourceData[, Version]): a range of this
                // workbook (xlDatabase = 1), as a Range or as "Sheet!R1C1:R5C5" /
                // "Sheet!A1:E5".
                ("create", [_, source, ..]) => {
                    let (sheet, range) = self.pivot_source_of(source)?;
                    self.pivot_caches.push((sheet, range));
                    Ok(Some(self.object(HostObject::NewPivotCache(self.pivot_caches.len() - 1))))
                }
                _ => Ok(None),
            },
            HostObject::NewPivotCache(id) => match lower.as_str() {
                // CreatePivotTable(TableDestination[, TableName]): an empty
                // table whose fields a macro then places.
                "createpivottable" => {
                    let [destination, rest @ ..] = args else {
                        return Err(host_error(450, "CreatePivotTable takes a destination"));
                    };
                    let (sheet_name, (row, column, _, _)) = self.pivot_source_of(destination)?;
                    let count = self.workbook.pivots.len();
                    let name = match rest.first() {
                        Some(Value::String(name)) if !name.is_empty() => name.clone(),
                        _ => format!("ピボットテーブル{}", count + 1),
                    };
                    let (source_sheet, source) = self.pivot_caches[id].clone();
                    let new = pivot::NewPivot {
                        name,
                        source_sheet: source_sheet.clone(),
                        source,
                        target_sheet: sheet_name,
                        at: (row, column),
                        rows: Vec::new(),
                        cols: Vec::new(),
                        pages: Vec::new(),
                        data: Vec::new(),
                    };
                    let mut spec = {
                        let get = pivot::source_reader(&*self.workbook, &source_sheet);
                        pivot::spec_for_new(&new, &get, pivot::Ui::Japanese)
                    };
                    spec.created = true;
                    spec.extent = spec.origin;
                    self.workbook.pivots.push(spec);
                    Ok(Some(self.object(HostObject::PivotTable(count))))
                }
                "sourcedata" => {
                    let (sheet, range) = self.pivot_caches[id].clone();
                    Ok(Some(Value::String(r1c1_range(&sheet, range))))
                }
                "recordcount" => Ok(Some(Value::Integer(i64::from(self.pivot_caches[id].1 .2.saturating_sub(self.pivot_caches[id].1 .0))))),
                _ => Ok(None),
            },
            HostObject::PivotCache(index) => {
                let spec = self.workbook.pivots[index].clone();
                Ok(Some(match lower.as_str() {
                    "refresh" => {
                        // Every table the cache feeds: here, the one it came with.
                        self.refresh_pivot(index)?;
                        self.recalculate();
                        Value::Empty
                    }
                    "sourcedata" => Value::String(r1c1_range(&spec.source_sheet, spec.source)),
                    "recordcount" => Value::Integer(i64::from(spec.source.2.saturating_sub(spec.source.0))),
                    _ => return Ok(None),
                }))
            }
            HostObject::PivotFields(index, axis) => match (lower.as_str(), args) {
                ("count", []) => Ok(Some(Value::Integer(Self::pivot_fields_of(&self.workbook.pivots[index], axis).len() as i64))),
                ("item", [wanted]) | ("_default", [wanted]) => self.pivot_field_item(index, axis, wanted).map(Some),
                ("parent", []) => Ok(Some(self.object(HostObject::PivotTable(index)))),
                _ => Ok(None),
            },
            HostObject::PivotField(index, field) => {
                let spec = self.workbook.pivots[index].clone();
                let (orientation, position) = Self::pivot_orientation(&spec, field);
                Ok(Some(match lower.as_str() {
                    "name" | "caption" | "value" | "_default" => Value::String(Self::pivot_field_name(&spec, field)),
                    "sourcename" => Value::String(match field {
                        PivotFieldRef::Data(d) => spec.field_names.get(spec.data[d].field).cloned().unwrap_or_default(),
                        PivotFieldRef::Source(f) => spec.field_names.get(f).cloned().unwrap_or_default(),
                    }),
                    "orientation" => Value::Integer(i64::from(orientation)),
                    "position" if orientation != XL_HIDDEN => Value::Integer(i64::from(position)),
                    "function" => match field {
                        PivotFieldRef::Data(d) => Value::Integer(i64::from(function_of(spec.data[d].aggregate))),
                        PivotFieldRef::Source(_) => return Err(host_error(1004, "Unable to get the Function property of the PivotField class")),
                    },
                    "currentpage" => match field {
                        PivotFieldRef::Source(f) => match spec.pages.iter().find(|(g, _)| *g == f) {
                            Some((_, Some(item))) => from_cell_value(item),
                            Some((_, None)) => Value::String("(All)".into()),
                            None => return Err(host_error(1004, "Unable to get the CurrentPage property of the PivotField class")),
                        },
                        PivotFieldRef::Data(_) => return Err(host_error(1004, "Unable to get the CurrentPage property of the PivotField class")),
                    },
                    "parent" => self.object(HostObject::PivotTable(index)),
                    _ => return Ok(None),
                }))
            }
            _ => Ok(None),
        }
    }

    /// Set a member of a pivot object; `false` for a name it does not take.
    ///
    /// Moving a field or picking a page item lays the table out again at once,
    /// as Excel does with manual update off.
    pub(super) fn set_pivot_member(&mut self, held: HostObject, name: &str, value: &Value) -> Result<bool, String> {
        let HostObject::PivotField(index, PivotFieldRef::Source(f)) = held else { return Ok(false) };
        let lower = name.to_ascii_lowercase();
        match lower.as_str() {
            "orientation" => {
                let wanted = any_number(value).ok_or_else(|| host_error(13, "Type mismatch"))? as i32;
                let spec = &mut self.workbook.pivots[index];
                spec.rows.retain(|g| *g != Some(f));
                spec.cols.retain(|g| *g != Some(f));
                spec.pages.retain(|(g, _)| *g != f);
                // A field placed for the first time lists its items as a new
                // table does: ascending (Excel saves them so, and refreshes
                // keep that order).
                if wanted != XL_HIDDEN && self.workbook.pivots[index].item_order[f].is_empty() {
                    let (source_sheet, source) = {
                        let spec = &self.workbook.pivots[index];
                        (spec.source_sheet.clone(), spec.source)
                    };
                    let items = {
                        let get = pivot::source_reader(&*self.workbook, &source_sheet);
                        pivot::ascending_items(source, f, &get)
                    };
                    self.workbook.pivots[index].item_order[f] = items;
                }
                let spec = &mut self.workbook.pivots[index];
                match wanted {
                    XL_ROW_FIELD => spec.rows.push(Some(f)),
                    XL_COLUMN_FIELD => spec.cols.push(Some(f)),
                    XL_PAGE_FIELD => spec.pages.push((f, None)),
                    XL_HIDDEN => {}
                    _ => return Err(host_error(1004, "Unable to set the Orientation property of the PivotField class")),
                }
            }
            "currentpage" => {
                let spec = &mut self.workbook.pivots[index];
                let Some(page) = spec.pages.iter_mut().find(|(g, _)| *g == f) else {
                    return Err(host_error(1004, "Unable to set the CurrentPage property of the PivotField class"));
                };
                page.1 = match value {
                    Value::String(text) if text.eq_ignore_ascii_case("(All)") => None,
                    Value::String(text) => Some(CellValue::String(text.clone())),
                    other => any_number(other).map(CellValue::Number),
                };
            }
            _ => return Ok(false),
        }
        self.refresh_pivot(index)?;
        self.recalculate();
        Ok(true)
    }

    /// Where a macro points a cache or a table: a Range, or text naming one
    /// ("D!R1C1:R5C5", "D!A1:E5", "P!R4C1"). Zero-based (row, col, row, col).
    fn pivot_source_of(&mut self, held: &Value) -> Result<(String, (u32, u32, u32, u32)), String> {
        if let Value::Object(object) = held {
            if let Some(range) = self.range(object) {
                let sheet = self.workbook.sheets[range.sheet].name.clone();
                return Ok((sheet, (range.start_row - 1, range.start_column, range.end_row - 1, range.end_column)));
            }
        }
        let Value::String(text) = held else { return Err(host_error(13, "Type mismatch")) };
        let (sheet, reference) = match text.rsplit_once('!') {
            Some((sheet, reference)) => (sheet.trim_matches('\'').replace("''", "'"), reference),
            None => (self.workbook.sheets[self.active_sheet].name.clone(), text.as_str()),
        };
        let cell = |part: &str| -> Option<(u32, u32)> {
            let part = part.trim().replace('$', "");
            let upper = part.to_ascii_uppercase();
            if let Some(rest) = upper.strip_prefix('R') {
                if let Some((r, c)) = rest.split_once('C') {
                    if let (Ok(r), Ok(c)) = (r.parse::<u32>(), c.parse::<u32>()) {
                        return Some((r.checked_sub(1)?, c.checked_sub(1)?));
                    }
                }
            }
            let letters: String = upper.chars().take_while(char::is_ascii_alphabetic).collect();
            let digits = &upper[letters.len()..];
            let row: u32 = digits.parse().ok()?;
            let col = letters.bytes().try_fold(0u32, |n, b| Some(n * 26 + u32::from(b - b'A' + 1)))?;
            Some((row.checked_sub(1)?, col.checked_sub(1)?))
        };
        let (first, last) = reference.split_once(':').unwrap_or((reference, reference));
        match (cell(first), cell(last)) {
            (Some((r0, c0)), Some((r1, c1))) => Ok((sheet, (r0, c0, r1, c1))),
            _ => Err(host_error(1004, format!("not a range: {text}"))),
        }
    }

    pub(super) fn pivot_enumerate(&mut self, held: HostObject) -> Option<Vec<Value>> {
        match held {
            HostObject::PivotTables(sheet) => {
                let on = self.pivots_on(sheet);
                Some(on.into_iter().map(|index| self.object(HostObject::PivotTable(index))).collect())
            }
            HostObject::PivotFields(index, axis) => {
                let fields = Self::pivot_fields_of(&self.workbook.pivots[index], axis);
                Some(fields.into_iter().map(|field| self.object(HostObject::PivotField(index, field))).collect())
            }
            _ => None,
        }
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    /// The table `_xlsx_vba_pivot_probe.py` builds in Excel, made here through
    /// the editor, and the same members asked of it.
    #[test]
    fn pivot_tables_answer_as_excel_does() {
        let original = include_bytes!("../../../tests/fixtures/basic_test.xlsx");
        let mut book = oxicells_core::parse_xlsx(original).unwrap();
        let setup = parse_module(r#"Public Sub Setup()
            Dim d As Object, p As Object
            Set d = Worksheets(1)
            d.Cells.Clear
            d.Name = "D"
            d.Range("A1:E1").Value = Array("Region", "Product", "Month", "Kind", "Sales")
            d.Range("A2:E2").Value = Array("West", "Apple", "Feb", "A", 5)
            d.Range("A3:E3").Value = Array("East", "Apple", "Jan", "B", 2)
            d.Range("A4:E4").Value = Array("East", "Pear", "Feb", "A", 3)
            d.Range("A5:E5").Value = Array("West", "Apple", "Jan", "A", 1)
            Set p = Worksheets.Add(After:=d)
            p.Name = "P"
        End Sub"#).unwrap();
        execute_with_host(&setup, "Setup", vec![], &mut WorkbookHost::new(&mut book, 0).unwrap()).unwrap();
        let mut editor = oxicells_core::XlsxEditor::new(original).unwrap();
        editor.apply_workbook(&book).unwrap();
        let bytes = editor.save().unwrap();
        let mut editor = oxicells_core::XlsxEditor::new(&bytes).unwrap();
        editor
            .add_pivot(
                &oxicells_core::pivot::NewPivot {
                    name: "PT".into(),
                    source_sheet: "D".into(),
                    source: (0, 0, 4, 4),
                    target_sheet: "P".into(),
                    at: (3, 0),
                    rows: vec![0, 1],
                    cols: vec![2],
                    pages: vec![3],
                    data: vec![(4, oxicells_core::pivot::Aggregate::Sum)],
                },
                oxicells_core::pivot::Ui::Japanese,
            )
            .unwrap();
        let mut book = oxicells_core::parse_xlsx(&editor.save().unwrap()).unwrap();
        let sheet = book.sheets.iter().position(|s| s.name == "P").unwrap();
        let ask = parse_module(r#"Public Function Ask() As String
            Dim pt As Object, f As Object, s As String
            Set pt = ActiveSheet.PivotTables(1)
            s = ActiveSheet.PivotTables.Count & "|" & pt.Name & "|" & pt.TableRange1.Address & "|" & pt.TableRange2.Address
            s = s & "|" & pt.PivotFields.Count
            For Each f In pt.RowFields
                s = s & "|" & f.Name & ":" & f.Orientation & ":" & f.Position
            Next
            s = s & "|" & pt.ColumnFields(1).Name & ":" & pt.ColumnFields(1).Orientation
            s = s & "|" & pt.PageFields(1).Name & ":" & pt.PageFields(1).Orientation & ":" & pt.PageFields(1).CurrentPage
            Set f = pt.DataFields(1)
            s = s & "|" & f.Name & ":" & f.SourceName & ":" & f.Function & ":" & f.Orientation
            s = s & "|" & pt.PivotFields("Sales").Orientation & "|" & pt.SourceData & "|" & pt.PivotCache.RecordCount
            s = s & "|" & pt.RowGrand & "|" & Range("C7").PivotTable.Name & "|" & ActiveSheet.PivotTables("pt").Name
            s = s & "|" & TypeName(pt) & "|" & Range("B7").Value
            Worksheets("D").Range("E3").Value = 20
            s = s & "|" & Range("B7").Value & "|" & pt.RefreshTable & "|" & Range("B7").Value
            Worksheets("D").Range("E3").Value = 30
            pt.PivotCache.Refresh
            s = s & "|" & Range("B7").Value
            Worksheets("D").Range("E3").Value = 40
            ThisWorkbook.RefreshAll
            Ask = s & "|" & Range("B7").Value & "|" & Range("D11").Value
        End Function"#).unwrap();
        let answer = execute_with_host(&ask, "Ask", vec![], &mut WorkbookHost::new(&mut book, sheet).unwrap()).unwrap();
        assert_eq!(
            answer,
            Value::String(
                "1|PT|$A$4:$D$11|$A$2:$D$11|5|Region:1:1|Product:1:2|Month:2|Kind:3:(All)\
                 |合計 / Sales:Sales:-4157:4|0|D!R1C1:R5C5|4|True|PT|PT|PivotTable|2|2|True|20|30|40|49"
                    .into()
            )
        );
    }

    /// A table made from nothing by a macro, the way the probe makes it in
    /// Excel: PivotCaches.Create, CreatePivotTable, fields placed one by one.
    /// Saved, it opens as a pivot table of the same name and grid.
    #[test]
    fn a_macro_creates_a_pivot_table_that_saves() {
        let original = include_bytes!("../../../tests/fixtures/basic_test.xlsx");
        let mut book = oxicells_core::parse_xlsx(original).unwrap();
        let make = parse_module(r#"Public Function Make() As String
            Dim d As Object, p As Object, pc As Object, pt As Object, r As Long, c As Long, s As String
            Set d = Worksheets(1)
            d.Cells.Clear
            d.Name = "D"
            d.Range("A1:E1").Value = Array("Region", "Product", "Month", "Kind", "Sales")
            d.Range("A2:E2").Value = Array("West", "Apple", "Feb", "A", 5)
            d.Range("A3:E3").Value = Array("East", "Apple", "Jan", "B", 2)
            d.Range("A4:E4").Value = Array("East", "Pear", "Feb", "A", 3)
            d.Range("A5:E5").Value = Array("West", "Apple", "Jan", "A", 1)
            Set p = Worksheets.Add(After:=d)
            p.Name = "P"
            Set pc = ThisWorkbook.PivotCaches.Create(1, "D!R1C1:R5C5")
            Set pt = pc.CreatePivotTable("P!R4C1", "PT")
            pt.PivotFields("Region").Orientation = 1
            pt.PivotFields("Product").Orientation = 1
            pt.PivotFields("Month").Orientation = 2
            pt.PivotFields("Kind").Orientation = 3
            pt.AddDataField pt.PivotFields("Sales")
            For r = 2 To 11
                For c = 1 To 4
                    s = s & p.Cells(r, c).Value & ","
                Next
                s = s & "/"
            Next
            Make = TypeName(pc) & "|" & pt.Name & "|" & pt.TableRange2.Address & "|" & pt.DataFields(1).Name & "|" & s
        End Function"#).unwrap();
        let answer = execute_with_host(&make, "Make", vec![], &mut WorkbookHost::new(&mut book, 0).unwrap()).unwrap();
        // Excel's grid for the same table (the probe, before its edits).
        assert_eq!(
            answer,
            Value::String(
                "PivotCache|PT|$A$2:$D$11|合計 / Sales|Kind,(すべて),,,/,,,,/合計 / Sales,列ラベル,,,/行ラベル,Jan,Feb,総計,/\
                 East,2,3,5,/Apple,2,,2,/Pear,,3,3,/West,1,5,6,/Apple,1,5,6,/総計,3,8,11,/"
                    .into()
            )
        );
        let mut editor = oxicells_core::XlsxEditor::new(original).unwrap();
        editor.apply_workbook(&book).unwrap();
        let saved = editor.save().unwrap();
        if let Ok(path) = std::env::var("OXI_VBA_PIVOT_OUT") {
            std::fs::write(path, &saved).unwrap();
        }
        let reopened = oxicells_core::parse_xlsx(&saved).unwrap();
        assert_eq!(reopened.pivots.len(), 1);
        let spec = &reopened.pivots[0];
        assert_eq!((spec.name.as_str(), spec.sheet.as_str(), spec.origin, spec.created), ("PT", "P", (3, 0), false));
        assert_eq!((spec.rows.len(), spec.cols.len(), spec.pages.len(), spec.data.len()), (2, 1, 1, 1));
    }
}
