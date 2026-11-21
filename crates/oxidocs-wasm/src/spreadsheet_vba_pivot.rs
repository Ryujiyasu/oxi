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

/// Which of a field's items a collection holds: `PivotItems`,
/// `VisibleItems` or `HiddenItems`.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub(super) enum PivotItemSet {
    All,
    Visible,
    Hidden,
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

/// A sheet block, zero-based (row, col, row, col).
type Block = (u32, u32, u32, u32);

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
                | HostObject::NewPivotCache(_)
                | HostObject::PivotItems(..)
                | HostObject::PivotItem(..)),
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
            let get = pivot::source_reader(&self.workbook, &spec.source_sheet);
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
                    "getpivotdata" => self.get_pivot_data(index, args)?,
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
                        let get = pivot::source_reader(&self.workbook, &source_sheet);
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
                let set = match lower.as_str() {
                    "pivotitems" => Some(PivotItemSet::All),
                    "visibleitems" => Some(PivotItemSet::Visible),
                    "hiddenitems" => Some(PivotItemSet::Hidden),
                    _ => None,
                };
                if let Some(set) = set {
                    return match args {
                        [] | [Value::Missing] => Ok(Some(self.object(HostObject::PivotItems(index, field, set)))),
                        [wanted] => self.pivot_items_item(index, field, set, wanted).map(Some),
                        _ => Err(host_error(450, "Wrong number of arguments")),
                    };
                }
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
            HostObject::PivotItems(index, field, set) => match (lower.as_str(), args) {
                ("count", []) => Ok(Some(Value::Integer(self.pivot_item_list(index, field, set).len() as i64))),
                ("item", [wanted]) | ("_default", [wanted]) => self.pivot_items_item(index, field, set, wanted).map(Some),
                ("parent", []) => Ok(Some(self.object(HostObject::PivotField(index, field)))),
                _ => Ok(None),
            },
            HostObject::PivotItem(index, f, at) => self.pivot_item_member(index, f, at, &lower).map(Some).or_else(|e| {
                if e.is_empty() { Ok(None) } else { Err(e) }
            }),
            _ => Ok(None),
        }
    }

    /// Set a member of a pivot object; `false` for a name it does not take.
    ///
    /// Moving a field or picking a page item lays the table out again at once,
    /// as Excel does with manual update off.
    pub(super) fn set_pivot_member(&mut self, held: HostObject, name: &str, value: &Value) -> Result<bool, String> {
        if let HostObject::PivotItem(index, f, at) = held {
            return self.set_pivot_item(index, f, at, name, value);
        }
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
                        let get = pivot::source_reader(&self.workbook, &source_sheet);
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
    fn pivot_source_of(&mut self, held: &Value) -> Result<(String, Block), String> {
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

    /// A field's items in the table's order, hidden ones included.
    fn pivot_field_items(&self, index: usize, f: usize) -> Vec<CellValue> {
        let spec = &self.workbook.pivots[index];
        let get = pivot::source_reader(&self.workbook, &spec.source_sheet);
        pivot::field_items(spec, f, &get)
    }

    fn pivot_item_hidden(spec: &PivotSpec, f: usize, item: &CellValue) -> bool {
        spec.hidden.get(f).is_some_and(|hidden| hidden.iter().any(|h| pivot::same_item(h, item)))
    }

    /// Whether an item shows. Measured: a page field showing "(All)" or
    /// "(複数のアイテム)" calls every item not visible, one showing an item
    /// only that one; a field on no axis calls them all visible.
    fn pivot_item_visible(spec: &PivotSpec, f: usize, item: &CellValue) -> bool {
        match spec.pages.iter().find(|(g, _)| *g == f) {
            Some((_, Some(pick))) => pivot::same_item(pick, item),
            Some((_, None)) => false,
            None => !Self::pivot_item_hidden(spec, f, item),
        }
    }

    /// The items a collection holds, as places in the field's item list.
    fn pivot_item_list(&self, index: usize, field: PivotFieldRef, set: PivotItemSet) -> Vec<usize> {
        // A value field has no items (measured: DataFields(1).PivotItems.Count 0).
        let PivotFieldRef::Source(f) = field else { return Vec::new() };
        let spec = &self.workbook.pivots[index];
        let items = self.pivot_field_items(index, f);
        (0..items.len())
            .filter(|at| match set {
                PivotItemSet::All => true,
                PivotItemSet::Visible => Self::pivot_item_visible(spec, f, &items[*at]),
                PivotItemSet::Hidden => !Self::pivot_item_visible(spec, f, &items[*at]),
            })
            .collect()
    }

    /// `PivotItems(n)` or `PivotItems("west")`: by place or by name without
    /// regard to case; anything else raises 1004 (measured).
    fn pivot_items_item(&mut self, index: usize, field: PivotFieldRef, set: PivotItemSet, wanted: &Value) -> Result<Value, String> {
        let list = self.pivot_item_list(index, field, set);
        let failed = || host_error(1004, "Unable to get the PivotItems property of the PivotField class");
        let PivotFieldRef::Source(f) = field else { return Err(failed()) };
        let items = self.pivot_field_items(index, f);
        let found = match wanted {
            Value::String(name) => list
                .iter()
                .copied()
                .find(|at| pivot::item_name(&items[*at], pivot::Ui::Japanese).eq_ignore_ascii_case(name.trim())),
            other => {
                let n = any_number(other).ok_or_else(|| host_error(13, "Type mismatch"))?;
                (n >= 1.0).then(|| list.get(n as usize - 1).copied()).flatten()
            }
        };
        found.map(|at| self.object(HostObject::PivotItem(index, f, at))).ok_or_else(failed)
    }

    /// The cells of a table's grid, as blocks of the sheet: every run of
    /// rows by every run of columns, both zero-based within the grid.
    fn pivot_blocks(&mut self, index: usize, origin: (u32, u32), rows: &[usize], cols: &[usize], what: &str) -> Result<Value, String> {
        let spec = &self.workbook.pivots[index];
        let sheet = self.sheet_index_named(&spec.sheet).ok_or_else(|| host_error(1004, "the pivot table's sheet is gone"))?;
        let runs = |mut list: Vec<usize>| -> Vec<(usize, usize)> {
            list.sort_unstable();
            list.dedup();
            let mut out: Vec<(usize, usize)> = Vec::new();
            for n in list {
                match out.last_mut() {
                    Some(run) if run.1 + 1 == n => run.1 = n,
                    _ => out.push((n, n)),
                }
            }
            out
        };
        let mut blocks = Vec::new();
        for (r0, r1) in runs(rows.to_vec()) {
            for (c0, c1) in runs(cols.to_vec()) {
                blocks.push(CellRange {
                    sheet,
                    start_row: origin.0 + r0 as u32 + 1,
                    start_column: origin.1 + c0 as u32,
                    end_row: origin.0 + r1 as u32 + 1,
                    end_column: origin.1 + c1 as u32,
                });
            }
        }
        match blocks.len() {
            0 => Err(host_error(1004, format!("Unable to get the {what} property of the PivotItem class"))),
            1 => Ok(self.object(HostObject::Range(blocks[0]))),
            _ => self.blocks_object(blocks),
        }
    }

    /// A member of a PivotItem; `Err("")` for a name it lacks.
    ///
    /// Measured on Region/Product rows, Month columns, Kind page: Name, Value
    /// and Caption the item as shown; Position its place among the field's
    /// visible items, #N/A (2042) when hidden or off the row and column
    /// axes; RecordCount its source records, hidden or not; LabelRange the
    /// cells that show it (an inner row item under two outer ones: two
    /// areas); DataRange the value cells of its own rows (an outer compact
    /// item: its own row only) or columns, grand totals left out, a page
    /// item's the whole body.
    fn pivot_item_member(&mut self, index: usize, f: usize, at: usize, lower: &str) -> Result<Value, String> {
        let items = self.pivot_field_items(index, f);
        let Some(item) = items.get(at).cloned() else {
            return Err(host_error(1004, "the pivot item is gone"));
        };
        let spec = self.workbook.pivots[index].clone();
        let name = pivot::item_name(&item, pivot::Ui::Japanese);
        let (orientation, _) = Self::pivot_orientation(&spec, PivotFieldRef::Source(f));
        Ok(match lower {
            "name" | "value" | "caption" | "sourcename" | "_default" => Value::String(name),
            "visible" => Value::Boolean(Self::pivot_item_visible(&spec, f, &item)),
            "position" => {
                if matches!(orientation, XL_ROW_FIELD | XL_COLUMN_FIELD) && Self::pivot_item_visible(&spec, f, &item) {
                    let before = items[..at].iter().filter(|v| Self::pivot_item_visible(&spec, f, v)).count();
                    Value::Integer(before as i64 + 1)
                } else {
                    Value::Error(2042)
                }
            }
            "recordcount" => {
                let get = pivot::source_reader(&self.workbook, &spec.source_sheet);
                let (r0, c0, r1, _) = spec.source;
                let n = (r0 + 1..=r1).filter(|r| pivot::same_item(&get(*r, c0 + f as u32), &item)).count();
                Value::Integer(n as i64)
            }
            "parent" => self.object(HostObject::PivotField(index, PivotFieldRef::Source(f))),
            "labelrange" | "datarange" => {
                let grid = {
                    let get = pivot::source_reader(&self.workbook, &spec.source_sheet);
                    pivot::compute(&spec, &get, pivot::Ui::Japanese)
                };
                let wanted = name.to_lowercase();
                let lc = grid.label_columns;
                let row_real: Vec<usize> = spec.rows.iter().flatten().copied().collect();
                let col_real: Vec<usize> = spec.cols.iter().flatten().copied().collect();
                let body_rows: Vec<usize> = grid
                    .row_paths
                    .iter()
                    .enumerate()
                    .filter(|(_, p)| p.as_ref().is_some_and(|(p, _)| !p.is_empty() || row_real.is_empty()))
                    .map(|(r, _)| r)
                    .collect();
                let body_cols: Vec<usize> = grid
                    .col_paths
                    .iter()
                    .enumerate()
                    .filter(|(_, (p, _))| !p.is_empty() || col_real.is_empty())
                    .map(|(j, _)| lc + j)
                    .collect();
                let what = if lower == "labelrange" { "LabelRange" } else { "DataRange" };
                let (rows, cols): (Vec<usize>, Vec<usize>) = match orientation {
                    XL_ROW_FIELD => {
                        let level = row_real.iter().position(|g| *g == f).unwrap_or(0);
                        let mine = |r: usize, exact: bool| {
                            grid.row_paths[r].as_ref().is_some_and(|(p, _)| {
                                p.get(level) == Some(&wanted) && (!exact || p.len() == level + 1)
                            })
                        };
                        if lower == "labelrange" {
                            let compact = spec.row_layout == pivot::RowLayout::Compact;
                            let column = if compact { 0 } else { level };
                            let rows: Vec<usize> = (0..grid.cells.len())
                                .filter(|r| match grid.kinds[*r] {
                                    pivot::RowKind::Item { level: l, .. } => {
                                        (!compact || l == level)
                                            && mine(*r, false)
                                            && grid.cells[*r].get(column).is_some_and(|v| !matches!(v, CellValue::Empty))
                                    }
                                    _ => false,
                                })
                                .collect();
                            (rows, vec![column])
                        } else {
                            let rows: Vec<usize> = (0..grid.cells.len()).filter(|r| mine(*r, true)).collect();
                            (rows, body_cols)
                        }
                    }
                    XL_COLUMN_FIELD => {
                        let level = spec.cols.iter().position(|g| *g == Some(f)).unwrap_or(0);
                        let real_level = col_real.iter().position(|g| *g == f).unwrap_or(0);
                        let mine = |j: usize, exact: bool| {
                            let p = &grid.col_paths[j].0;
                            p.get(real_level) == Some(&wanted) && (!exact || p.len() == real_level + 1)
                        };
                        let pages = if spec.pages.is_empty() { 0 } else { spec.pages.len() + 1 };
                        if lower == "labelrange" {
                            let header = pages + 1 + level;
                            let cols: Vec<usize> = (0..grid.col_paths.len())
                                .filter(|j| {
                                    mine(*j, false)
                                        && grid.cells.get(header).and_then(|row| row.get(lc + j)).is_some_and(|v| !matches!(v, CellValue::Empty))
                                })
                                .map(|j| lc + j)
                                .collect();
                            (vec![header], cols)
                        } else {
                            let cols: Vec<usize> = (0..grid.col_paths.len()).filter(|j| mine(*j, true)).map(|j| lc + j).collect();
                            (body_rows, cols)
                        }
                    }
                    XL_PAGE_FIELD if lower == "datarange" => (body_rows, body_cols),
                    _ => (Vec::new(), Vec::new()),
                };
                self.pivot_blocks(index, grid.origin, &rows, &cols, what)?
            }
            _ => return Err(String::new()),
        })
    }

    /// `PivotItem.Visible` and `.Position`. Measured: hiding the last visible
    /// item of a row field raises 1004; Position takes 1 to the visible count
    /// (5 of 2 raises 1004) and moves the item there among the visible ones.
    fn set_pivot_item(&mut self, index: usize, f: usize, at: usize, name: &str, value: &Value) -> Result<bool, String> {
        let items = self.pivot_field_items(index, f);
        let Some(item) = items.get(at).cloned() else {
            return Err(host_error(1004, "the pivot item is gone"));
        };
        let lower = name.to_ascii_lowercase();
        match lower.as_str() {
            "visible" => {
                let show = match value {
                    Value::Boolean(b) => *b,
                    other => any_number(other).ok_or_else(|| host_error(13, "Type mismatch"))? != 0.0,
                };
                let spec = &mut self.workbook.pivots[index];
                if show {
                    spec.hidden[f].retain(|h| !pivot::same_item(h, &item));
                } else if !Self::pivot_item_hidden(spec, f, &item) {
                    let shown = items.iter().filter(|v| !Self::pivot_item_hidden(spec, f, v)).count();
                    if shown <= 1 {
                        return Err(host_error(1004, "Unable to set the Visible property of the PivotItem class"));
                    }
                    spec.hidden[f].push(item);
                }
            }
            "position" => {
                let spec = &self.workbook.pivots[index];
                let (orientation, _) = Self::pivot_orientation(spec, PivotFieldRef::Source(f));
                let visible: Vec<CellValue> = items.iter().filter(|v| Self::pivot_item_visible(spec, f, v)).cloned().collect();
                let n = any_number(value).ok_or_else(|| host_error(13, "Type mismatch"))?;
                if !matches!(orientation, XL_ROW_FIELD | XL_COLUMN_FIELD)
                    || !Self::pivot_item_visible(spec, f, &item)
                    || n < 1.0
                    || n as usize > visible.len()
                {
                    return Err(host_error(1004, "Unable to set the Position property of the PivotItem class"));
                }
                let mut order: Vec<CellValue> = items.into_iter().filter(|v| !pivot::same_item(v, &item)).collect();
                let others: Vec<&CellValue> = visible.iter().filter(|v| !pivot::same_item(v, &item)).collect();
                let at = match others.get(n as usize - 1) {
                    Some(next) => order.iter().position(|v| pivot::same_item(v, next)).unwrap_or(order.len()),
                    None => others
                        .last()
                        .and_then(|last| order.iter().position(|v| pivot::same_item(v, last)))
                        .map_or(order.len(), |p| p + 1),
                };
                order.insert(at, item);
                self.workbook.pivots[index].item_order[f] = order;
            }
            _ => return Ok(false),
        }
        self.refresh_pivot(index)?;
        self.recalculate();
        Ok(true)
    }

    /// `PivotTable.GetPivotData([DataField], Field1, Item1, ...)`: the cell
    /// that shows the value. Measured: the value field by its caption or its
    /// source column's name, any case; row (column) fields must be the
    /// outermost ones of their axis; a page field passes only while it shows
    /// that item; no pairs is the grand total; a hidden item, an unknown name
    /// or an inner field alone raises 1004; a shown empty cell is returned.
    fn get_pivot_data(&mut self, index: usize, args: &[Value]) -> Result<Value, String> {
        let failed = || host_error(1004, "Unable to get the GetPivotData property of the PivotTable class");
        let spec = self.workbook.pivots[index].clone();
        let text = |v: &Value| -> Option<String> {
            match v {
                Value::String(s) => Some(s.trim().to_lowercase()),
                Value::Missing | Value::Empty => None,
                other => any_number(other).map(|n| pivot::item_name(&CellValue::Number(n), pivot::Ui::Japanese).to_lowercase()),
            }
        };
        let data = match args.first().and_then(text) {
            None => 0,
            Some(wanted) => spec
                .data
                .iter()
                .position(|d| {
                    d.name.to_lowercase() == wanted
                        || spec.field_names.get(d.field).is_some_and(|n| n.to_lowercase() == wanted)
                })
                .ok_or_else(failed)?,
        };
        let row_real: Vec<usize> = spec.rows.iter().flatten().copied().collect();
        let col_real: Vec<usize> = spec.cols.iter().flatten().copied().collect();
        let mut row_pairs: Vec<(usize, String)> = Vec::new();
        let mut col_pairs: Vec<(usize, String)> = Vec::new();
        for pair in args.get(1..).unwrap_or(&[]).chunks(2) {
            let [field, item] = pair else { return Err(failed()) };
            let (Some(field), Some(item)) = (text(field), text(item)) else { return Err(failed()) };
            let f = spec.field_names.iter().position(|n| n.to_lowercase() == field).ok_or_else(failed)?;
            if let Some(p) = row_real.iter().position(|g| *g == f) {
                row_pairs.push((p, item));
            } else if let Some(p) = col_real.iter().position(|g| *g == f) {
                col_pairs.push((p, item));
            } else {
                match spec.pages.iter().find(|(g, _)| *g == f) {
                    Some((_, Some(pick))) if pivot::item_name(pick, pivot::Ui::Japanese).to_lowercase() == item => {}
                    _ => return Err(failed()),
                }
            }
        }
        let path = |mut pairs: Vec<(usize, String)>| -> Option<Vec<String>> {
            pairs.sort();
            pairs.iter().enumerate().all(|(i, (p, _))| *p == i).then(|| pairs.into_iter().map(|(_, s)| s).collect())
        };
        let (Some(row_path), Some(col_path)) = (path(row_pairs), path(col_pairs)) else { return Err(failed()) };
        let grid = {
            let get = pivot::source_reader(&self.workbook, &spec.source_sheet);
            pivot::compute(&spec, &get, pivot::Ui::Japanese)
        };
        for (r, held) in grid.row_paths.iter().enumerate() {
            let Some((p, row_data)) = held else { continue };
            if *p != row_path {
                continue;
            }
            for (j, (q, col_data)) in grid.col_paths.iter().enumerate() {
                if *q == col_path && row_data.or(*col_data).unwrap_or(0) == data {
                    return self.pivot_blocks(index, grid.origin, &[r], &[grid.label_columns + j], "GetPivotData");
                }
            }
        }
        Err(failed())
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
            HostObject::PivotItems(index, field, set) => {
                let items = self.pivot_item_list(index, field, set);
                let PivotFieldRef::Source(f) = field else { return Some(Vec::new()) };
                Some(items.into_iter().map(|at| self.object(HostObject::PivotItem(index, f, at))).collect())
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
    /// `_xlsx_vba_pivot_items_probe.py`'s table and questions: each field's
    /// items, GetPivotData's cell, hiding, moving and page picks, against
    /// what Excel 16.0 answered.
    #[test]
    fn pivot_items_and_get_pivot_data_answer_as_excel_does() {
        let original = include_bytes!("../../../tests/fixtures/basic_test.xlsx");
        let mut book = oxicells_core::parse_xlsx(original).unwrap();
        let module = parse_module(r##"Private Function Grid(p As Object) As String
            Dim r As Long, c As Long, s As String
            For r = 2 To 11
                For c = 1 To 4
                    s = s & p.Cells(r, c).Value & ","
                Next
                s = s & "/"
            Next
            Grid = s
        End Function

        Private Function GP(pt As Object, n As Long, Optional d, Optional f1, Optional i1, Optional f2, Optional i2, Optional f3, Optional i3) As String
            On Error GoTo bad
            If n = 0 Then
                GP = pt.GetPivotData().Address
            ElseIf n = 1 Then
                GP = pt.GetPivotData(d).Address
            ElseIf n = 3 Then
                GP = pt.GetPivotData(d, f1, i1).Address
            ElseIf n = 5 Then
                GP = pt.GetPivotData(d, f1, i1, f2, i2).Address
            Else
                GP = pt.GetPivotData(d, f1, i1, f2, i2, f3, i3).Address
            End If
            Exit Function
        bad:
            GP = "ERR" & Err.Number
        End Function

        Private Function Items(pf As Object) As String
            Dim it As Object, s As String
            For Each it In pf.PivotItems
                s = s & it.Name & ":" & it.Visible & ":"
                If IsError(it.Position) Then s = s & "NA" Else s = s & it.Position
                s = s & ":" & it.RecordCount & ":"
                On Error Resume Next
                s = s & it.LabelRange.Address
                If Err.Number <> 0 Then s = s & "ERR"
                Err.Clear
                s = s & ":" & it.DataRange.Address
                If Err.Number <> 0 Then s = s & "ERR"
                Err.Clear
                On Error GoTo 0
                s = s & ";"
            Next
            Items = pf.PivotItems.Count & "=" & s
        End Function

        Private Function Try(o As Object, name As String, v) As String
            On Error GoTo bad
            If name = "Visible" Then
                o.Visible = v
            Else
                o.Position = v
            End If
            Try = "ok"
            Exit Function
        bad:
            Try = "ERR" & Err.Number
        End Function

        Public Function Probe() As String
            Dim d As Object, p As Object, pc As Object, pt As Object, s As String, f As Variant
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
            For Each f In Array("Region", "Product", "Month", "Kind", "Sales")
                s = s & f & "#" & Items(pt.PivotFields(f)) & vbLf
            Next
            s = s & pt.PivotFields("Region").PivotItems("west").Name & "|" & pt.DataFields(1).PivotItems.Count & vbLf
            s = s & GP(pt, 1, "Sales") & "," & GP(pt, 3, "Sales", "Region", "West") & "," & GP(pt, 5, "Sales", "Region", "West", "Product", "Apple") & ","
            s = s & GP(pt, 3, "Sales", "Month", "Jan") & "," & GP(pt, 5, "Sales", "Region", "East", "Month", "Feb") & ","
            s = s & GP(pt, 7, "Sales", "Region", "East", "Product", "Pear", "Month", "Feb") & "," & GP(pt, 7, "Sales", "Region", "East", "Product", "Pear", "Month", "Jan") & ","
            s = s & GP(pt, 3, "Sales", "Product", "Apple") & "," & GP(pt, 3, "Sales", "Region", "Nope") & "," & GP(pt, 3, "合計 / Sales", "Region", "West") & ","
            s = s & GP(pt, 3, "sales", "region", "west") & "," & GP(pt, 3, "Sales", "Kind", "A") & "," & GP(pt, 0) & vbLf
            s = s & Try(pt.PivotFields("Region").PivotItems("East"), "Visible", False) & "|" & Grid(p) & vbLf
            s = s & pt.PivotFields("Region").PivotItems("East").Visible & "|" & IsError(pt.PivotFields("Region").PivotItems("East").Position) & "|"
            s = s & pt.PivotFields("Region").PivotItems("West").Position & "|" & pt.PivotFields("Region").VisibleItems.Count & "|" & pt.PivotFields("Region").HiddenItems(1).Name & "|"
            s = s & pt.PivotFields("Region").PivotItems.Count & "|" & GP(pt, 3, "Sales", "Region", "East") & "|"
            s = s & Try(pt.PivotFields("Region").PivotItems("West"), "Visible", False) & "|" & Try(pt.PivotFields("Region").PivotItems("East"), "Visible", True) & "|" & Grid(p) & vbLf
            s = s & Try(pt.PivotFields("Product").PivotItems("Pear"), "Position", 1) & "|" & Grid(p) & "|"
            s = s & pt.PivotFields("Product").PivotItems(1).Name & "|" & Try(pt.PivotFields("Product").PivotItems("Pear"), "Position", 5) & "|"
            s = s & Try(pt.PivotFields("Product").PivotItems("Pear"), "Position", 2) & "|" & Grid(p) & vbLf
            pt.PivotFields("Kind").CurrentPage = "A"
            s = s & pt.PivotFields("Kind").PivotItems("A").Visible & "|" & pt.PivotFields("Kind").PivotItems("B").Visible & "|"
            s = s & pt.PivotFields("Kind").PivotItems("A").DataRange.Address & "|" & GP(pt, 3, "Sales", "Kind", "A") & "|"
            pt.PivotFields("Kind").CurrentPage = "(All)"
            s = s & pt.PivotFields("Kind").PivotItems("A").Visible & "|" & Try(pt.PivotFields("Kind").PivotItems("B"), "Visible", False) & "|" & Grid(p)
            Probe = s
        End Function"##).unwrap();
        let answer = execute_with_host(&module, "Probe", vec![], &mut WorkbookHost::new(&mut book, 0).unwrap()).unwrap();
        let Value::String(answer) = answer else { panic!("{answer:?}") };
        let full = "Kind,(すべて),,,/,,,,/合計 / Sales,列ラベル,,,/行ラベル,Jan,Feb,総計,/\
                    East,2,3,5,/Apple,2,,2,/Pear,,3,3,/West,1,5,6,/Apple,1,5,6,/総計,3,8,11,/";
        let expected = [
            "Region#2=East:True:1:2:$A$6:$B$6:$C$6;West:True:2:2:$A$9:$B$9:$C$9;".to_string(),
            "Product#2=Apple:True:1:3:$A$7,$A$10:$B$7:$C$7,$B$10:$C$10;Pear:True:2:1:$A$8:$B$8:$C$8;".into(),
            "Month#2=Jan:True:1:2:$B$5:$B$6:$B$10;Feb:True:2:2:$C$5:$C$6:$C$10;".into(),
            "Kind#2=A:False:NA:3:ERR:$B$6:$C$10;B:False:NA:1:ERR:$B$6:$C$10;".into(),
            "Sales#4=1:True:NA:1:ERRERR;2:True:NA:1:ERRERR;3:True:NA:1:ERRERR;5:True:NA:1:ERRERR;".into(),
            "West|0".into(),
            "$D$11,$D$9,$D$10,$B$11,$C$6,$C$8,$B$8,ERR1004,ERR1004,$D$9,$D$9,ERR1004,$D$11".into(),
            "ok|Kind,(すべて),,,/,,,,/合計 / Sales,列ラベル,,,/行ラベル,Jan,Feb,総計,/\
             West,1,5,6,/Apple,1,5,6,/総計,1,5,6,/,,,,/,,,,/,,,,/"
                .into(),
            format!("False|True|1|1|East|2|ERR1004|ERR1004|ok|{full}"),
            "ok|Kind,(すべて),,,/,,,,/合計 / Sales,列ラベル,,,/行ラベル,Jan,Feb,総計,/\
             East,2,3,5,/Pear,,3,3,/Apple,2,,2,/West,1,5,6,/Apple,1,5,6,/総計,3,8,11,/|Pear|ERR1004|ok|"
                .to_string()
                + full,
            "True|False|$B$6:$C$9|$D$10|False|ok|Kind,(複数のアイテム),,,/,,,,/合計 / Sales,列ラベル,,,/行ラベル,Jan,Feb,総計,/\
             East,,3,3,/Pear,,3,3,/West,1,5,6,/Apple,1,5,6,/総計,1,8,9,/,,,,/"
                .into(),
        ];
        let got: Vec<&str> = answer.split('\n').collect();
        assert_eq!(got.len(), expected.len(), "{answer}");
        for (got, want) in got.iter().zip(&expected) {
            assert_eq!(got, want);
        }
        // The hidden page item is saved and read back.
        let mut editor = oxicells_core::XlsxEditor::new(original).unwrap();
        editor.apply_workbook(&book).unwrap();
        let saved = editor.save().unwrap();
        if let Ok(path) = std::env::var("OXI_VBA_PIVOT_ITEMS_OUT") {
            std::fs::write(path, &saved).unwrap();
        }
        let reopened = oxicells_core::parse_xlsx(&saved).unwrap();
        let spec = &reopened.pivots[0];
        assert_eq!(spec.hidden[3], vec![CellValue::String("B".into())]);
        assert_eq!(spec.pages, vec![(3, None)]);
    }

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
