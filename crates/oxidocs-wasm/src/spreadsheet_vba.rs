// SPDX-License-Identifier: MIT OR Apache-2.0

use std::collections::{BTreeMap, BTreeSet};

use oxicells_core::ir::{Cell, CellStyle, CellValue, MergeCell, Row, Workbook};
use oxivba_core::ast::{ParamMode, ProcKind, Visibility};
#[cfg(test)]
use oxivba_core::execute_with_host;
use oxivba_core::{
    parse_module, ArrayDimension, ArrayValue, Host, ModuleItem, ObjectRef, Runtime, Value,
};
use serde::{Deserialize, Serialize};
use wasm_bindgen::prelude::*;

#[derive(Debug, Clone, Copy)]
struct CellAddress {
    sheet: usize,
    row: u32,
    column: u32,
}

#[derive(Debug, Clone, Copy)]
struct CellRange {
    sheet: usize,
    start_row: u32,
    start_column: u32,
    end_row: u32,
    end_column: u32,
}

#[derive(Debug, Clone, Copy)]
enum RangeAxis {
    Rows,
    Columns,
}

#[derive(Debug, Clone, Copy)]
enum EndDirection {
    Up,
    Down,
    Left,
    Right,
}

#[derive(Debug, Clone, Copy)]
enum BorderSelection {
    All,
    EdgeLeft,
    EdgeTop,
    EdgeBottom,
    EdgeRight,
    InsideVertical,
    InsideHorizontal,
}

const MAX_WORKSHEET_ROW: u32 = 1_048_576;
const MAX_WORKSHEET_COLUMN: u32 = 16_383;

impl CellRange {
    fn single(address: CellAddress) -> Self {
        Self {
            sheet: address.sheet,
            start_row: address.row,
            start_column: address.column,
            end_row: address.row,
            end_column: address.column,
        }
    }

    fn is_single(self) -> bool {
        self.start_row == self.end_row && self.start_column == self.end_column
    }

    fn addresses(self) -> impl Iterator<Item = CellAddress> {
        (self.start_row..=self.end_row).flat_map(move |row| {
            (self.start_column..=self.end_column).map(move |column| CellAddress {
                sheet: self.sheet,
                row,
                column,
            })
        })
    }
}

#[derive(Debug, Clone, Copy)]
enum HostObject {
    Range(CellRange),
    RangeFont(CellRange),
    RangeInterior(CellRange),
    RangeBorders(CellRange, BorderSelection),
    RangeCollection(CellRange, RangeAxis),
    Worksheet(usize),
    Worksheets,
    Workbook,
    Application,
    DebugConsole,
}

struct WorkbookHost<'a> {
    workbook: &'a mut Workbook,
    active_sheet: usize,
    selection: CellRange,
    objects: Vec<HostObject>,
    debug_output: Vec<String>,
    messages: Vec<BrowserMessage>,
}

impl<'a> WorkbookHost<'a> {
    fn new(workbook: &'a mut Workbook, active_sheet: usize) -> Result<Self, String> {
        if active_sheet >= workbook.sheets.len() {
            return Err(format!(
                "active sheet index is out of range: {active_sheet}"
            ));
        }
        Ok(Self {
            workbook,
            active_sheet,
            selection: CellRange::single(CellAddress {
                sheet: active_sheet,
                row: 1,
                column: 0,
            }),
            objects: Vec::new(),
            debug_output: Vec::new(),
            messages: Vec::new(),
        })
    }

    fn object(&mut self, object: HostObject) -> Value {
        let handle = self.objects.len() as u64;
        self.objects.push(object);
        Value::Object(ObjectRef {
            handle,
            kind: match object {
                HostObject::Range(_) => "Range",
                HostObject::RangeFont(_) => "Font",
                HostObject::RangeInterior(_) => "Interior",
                HostObject::RangeBorders(_, _) => "Borders",
                HostObject::RangeCollection(_, RangeAxis::Rows) => "Rows",
                HostObject::RangeCollection(_, RangeAxis::Columns) => "Columns",
                HostObject::Worksheet(_) => "Worksheet",
                HostObject::Worksheets => "Worksheets",
                HostObject::Workbook => "Workbook",
                HostObject::Application => "Application",
                HostObject::DebugConsole => "Debug",
            }
            .to_string(),
        })
    }

    fn range(&self, object: &ObjectRef) -> Option<CellRange> {
        match self.objects.get(object.handle as usize) {
            Some(HostObject::Range(range)) => Some(*range),
            _ => None,
        }
    }

    fn range_font(&self, object: &ObjectRef) -> Option<CellRange> {
        match self.objects.get(object.handle as usize) {
            Some(HostObject::RangeFont(range)) => Some(*range),
            _ => None,
        }
    }

    fn range_interior(&self, object: &ObjectRef) -> Option<CellRange> {
        match self.objects.get(object.handle as usize) {
            Some(HostObject::RangeInterior(range)) => Some(*range),
            _ => None,
        }
    }

    fn range_borders(&self, object: &ObjectRef) -> Option<(CellRange, BorderSelection)> {
        match self.objects.get(object.handle as usize) {
            Some(HostObject::RangeBorders(range, selection)) => Some((*range, *selection)),
            _ => None,
        }
    }

    fn borders_object(&mut self, range: CellRange, args: &[Value]) -> Result<Value, String> {
        let selection = match args {
            [] => BorderSelection::All,
            [value] => border_selection(value)?,
            _ => return Err("Range.Borders expects zero or one border index".to_string()),
        };
        Ok(self.object(HostObject::RangeBorders(range, selection)))
    }

    fn range_collection(&self, object: &ObjectRef) -> Option<(CellRange, RangeAxis)> {
        match self.objects.get(object.handle as usize) {
            Some(HostObject::RangeCollection(range, axis)) => Some((*range, *axis)),
            _ => None,
        }
    }

    fn worksheet(&self, object: &ObjectRef) -> Option<usize> {
        match self.objects.get(object.handle as usize) {
            Some(HostObject::Worksheet(sheet)) => Some(*sheet),
            _ => None,
        }
    }

    fn is_workbook(&self, object: &ObjectRef) -> bool {
        matches!(
            self.objects.get(object.handle as usize),
            Some(HostObject::Workbook)
        )
    }

    fn is_worksheets(&self, object: &ObjectRef) -> bool {
        matches!(
            self.objects.get(object.handle as usize),
            Some(HostObject::Worksheets)
        )
    }

    fn is_application(&self, object: &ObjectRef) -> bool {
        matches!(
            self.objects.get(object.handle as usize),
            Some(HostObject::Application)
        )
    }

    fn is_debug_console(&self, object: &ObjectRef) -> bool {
        matches!(
            self.objects.get(object.handle as usize),
            Some(HostObject::DebugConsole)
        )
    }

    fn take_debug_output(&mut self) -> Vec<String> {
        std::mem::take(&mut self.debug_output)
    }

    fn take_messages(&mut self) -> Vec<BrowserMessage> {
        std::mem::take(&mut self.messages)
    }

    fn show_message_box(&mut self, args: &[Value]) -> Result<Value, String> {
        if !(1..=5).contains(&args.len()) {
            return Err(format!(
                "MsgBox expects 1 to 5 arguments, received {}",
                args.len()
            ));
        }
        let buttons = match args.get(1) {
            None | Some(Value::Missing) | Some(Value::Empty) => 0,
            Some(Value::Integer(value)) => *value,
            Some(Value::Double(value)) if value.is_finite() && value.fract() == 0.0 => {
                *value as i64
            }
            Some(_) => return Err("MsgBox buttons must be an integer style".to_string()),
        };
        if buttons & 0x0f != 0 {
            return Err(
                "browser MsgBox supports vbOKOnly; interactive button styles are unavailable"
                    .to_string(),
            );
        }
        let title = match args.get(2) {
            None | Some(Value::Missing) | Some(Value::Empty) => "Oxi VBA".to_string(),
            Some(value) => format_debug_value(value),
        };
        self.messages.push(BrowserMessage {
            prompt: format_debug_value(&args[0]),
            title,
        });
        Ok(Value::Integer(1))
    }

    fn worksheet_object(&mut self, value: &Value) -> Result<Value, String> {
        let sheet = self.worksheet_from_value(value)?;
        Ok(self.object(HostObject::Worksheet(sheet)))
    }

    fn worksheets_object_or_item(&mut self, args: &[Value]) -> Result<Value, String> {
        match args {
            [] => Ok(self.object(HostObject::Worksheets)),
            [value] => self.worksheet_object(value),
            _ => Err("Worksheets expects zero or one argument".to_string()),
        }
    }

    fn worksheet_from_value(&self, value: &Value) -> Result<usize, String> {
        match value {
            Value::String(name) => self
                .workbook
                .sheets
                .iter()
                .position(|sheet| sheet.name.eq_ignore_ascii_case(name))
                .ok_or_else(|| format!("worksheet not found: {name}")),
            Value::Integer(index) => self.worksheet_from_number(*index as f64),
            Value::Double(index) => self.worksheet_from_number(*index),
            _ => Err("Worksheets expects a sheet name or one-based index".to_string()),
        }
    }

    fn worksheet_from_number(&self, index: f64) -> Result<usize, String> {
        if !index.is_finite() || index.fract() != 0.0 || index < 1.0 {
            return Err("worksheet index must be a positive integer".to_string());
        }
        let index = index as usize - 1;
        (index < self.workbook.sheets.len())
            .then_some(index)
            .ok_or_else(|| format!("worksheet index is out of range: {}", index + 1))
    }

    fn range_object(&mut self, sheet: usize, args: &[Value]) -> Result<Value, String> {
        let (start, end) = match args {
            [Value::String(reference)] => parse_range_reference(reference)?,
            [Value::String(start), Value::String(end)] => {
                (parse_a1_reference(start)?, parse_a1_reference(end)?)
            }
            _ => return Err("Range expects one range reference or two cell references".to_string()),
        };
        let (start_column, start_row) = start;
        let (end_column, end_row) = end;
        Ok(self.object(HostObject::Range(CellRange {
            sheet,
            start_row: start_row.min(end_row),
            start_column: start_column.min(end_column),
            end_row: start_row.max(end_row),
            end_column: start_column.max(end_column),
        })))
    }

    fn evaluate_object(&mut self, sheet: usize, args: &[Value]) -> Result<Value, String> {
        let [Value::String(expression)] = args else {
            return Err("Evaluate expects one String expression".to_string());
        };
        let expression = expression
            .trim()
            .strip_prefix('=')
            .unwrap_or(expression.trim());
        let expression = expression
            .strip_prefix('[')
            .and_then(|value| value.strip_suffix(']'))
            .unwrap_or(expression)
            .trim();
        let (sheet, reference) = match expression.rsplit_once('!') {
            Some((sheet_name, reference)) => {
                let mut sheet_name = sheet_name.trim();
                if let Some(workbook_end) = sheet_name.rfind(']') {
                    sheet_name = &sheet_name[workbook_end + 1..];
                }
                let sheet_name = if sheet_name.starts_with('\'')
                    && sheet_name.ends_with('\'')
                    && sheet_name.len() >= 2
                {
                    sheet_name[1..sheet_name.len() - 1].replace("''", "'")
                } else {
                    sheet_name.to_string()
                };
                let sheet = self
                    .workbook
                    .sheets
                    .iter()
                    .position(|candidate| candidate.name.eq_ignore_ascii_case(&sheet_name))
                    .ok_or_else(|| format!("worksheet not found: {sheet_name}"))?;
                (sheet, reference.trim())
            }
            None => (sheet, expression),
        };
        self.range_object(sheet, &[Value::String(reference.to_string())])
    }

    fn used_range_object(&mut self, sheet: usize) -> Result<Value, String> {
        let worksheet = self
            .workbook
            .sheets
            .get(sheet)
            .ok_or_else(|| "worksheet no longer exists".to_string())?;
        let mut bounds = None::<(u32, u32, u32, u32)>;
        for row in &worksheet.rows {
            for cell in &row.cells {
                bounds = Some(match bounds {
                    Some((start_row, start_column, end_row, end_column)) => (
                        start_row.min(row.index),
                        start_column.min(cell.col),
                        end_row.max(row.index),
                        end_column.max(cell.col),
                    ),
                    None => (row.index, cell.col, row.index, cell.col),
                });
            }
        }
        let (start_row, start_column, end_row, end_column) = bounds.unwrap_or((1, 0, 1, 0));
        Ok(self.object(HostObject::Range(CellRange {
            sheet,
            start_row,
            start_column,
            end_row,
            end_column,
        })))
    }

    fn worksheet_axis_object_or_item(
        &mut self,
        sheet: usize,
        axis: RangeAxis,
        args: &[Value],
    ) -> Result<Value, String> {
        let range = CellRange {
            sheet,
            start_row: 1,
            start_column: 0,
            end_row: MAX_WORKSHEET_ROW,
            end_column: MAX_WORKSHEET_COLUMN,
        };
        self.range_collection_object_or_item(range, axis, args)
    }

    fn cells_object(&mut self, sheet: usize, args: &[Value]) -> Result<Value, String> {
        let [row, column] = args else {
            return Err("Cells expects row and column".to_string());
        };
        let row = positive_index(row, "row")?;
        let column = positive_index(column, "column")? - 1;
        Ok(
            self.object(HostObject::Range(CellRange::single(CellAddress {
                sheet,
                row,
                column,
            }))),
        )
    }

    fn range_cells_object(&mut self, range: CellRange, args: &[Value]) -> Result<Value, String> {
        let [row, column] = args else {
            return Err("Range.Cells expects row and column".to_string());
        };
        let row = positive_index(row, "row")? - 1;
        let column = positive_index(column, "column")? - 1;
        let row = range
            .start_row
            .checked_add(row)
            .ok_or_else(|| "Range.Cells row is too large".to_string())?;
        let column = range
            .start_column
            .checked_add(column)
            .ok_or_else(|| "Range.Cells column is too large".to_string())?;
        Ok(
            self.object(HostObject::Range(CellRange::single(CellAddress {
                sheet: range.sheet,
                row,
                column,
            }))),
        )
    }

    fn range_collection_object_or_item(
        &mut self,
        range: CellRange,
        axis: RangeAxis,
        args: &[Value],
    ) -> Result<Value, String> {
        match args {
            [] => Ok(self.object(HostObject::RangeCollection(range, axis))),
            [index] => self.range_collection_item(range, axis, index),
            _ => Err(format!(
                "Range.{} expects zero or one argument",
                range_axis_name(axis)
            )),
        }
    }

    fn range_collection_item(
        &mut self,
        range: CellRange,
        axis: RangeAxis,
        index: &Value,
    ) -> Result<Value, String> {
        let offset = positive_index(index, range_axis_name(axis))? - 1;
        let item = match axis {
            RangeAxis::Rows => {
                let row = range
                    .start_row
                    .checked_add(offset)
                    .ok_or_else(|| "Range.Rows index is too large".to_string())?;
                CellRange {
                    start_row: row,
                    end_row: row,
                    ..range
                }
            }
            RangeAxis::Columns => {
                let column = range
                    .start_column
                    .checked_add(offset)
                    .ok_or_else(|| "Range.Columns index is too large".to_string())?;
                CellRange {
                    start_column: column,
                    end_column: column,
                    ..range
                }
            }
        };
        Ok(self.object(HostObject::Range(item)))
    }

    fn cell_value(&self, address: CellAddress) -> Value {
        self.workbook
            .sheets
            .get(address.sheet)
            .and_then(|sheet| sheet.rows.iter().find(|row| row.index == address.row))
            .and_then(|row| row.cells.iter().find(|cell| cell.col == address.column))
            .map(|cell| from_cell_value(&cell.value))
            .unwrap_or(Value::Empty)
    }

    fn range_cell_count(range: CellRange) -> Result<usize, String> {
        let count = Self::range_cell_count_large(range)?;
        if count > 1_000_000 {
            return Err("cell range exceeds the 1,000,000-cell execution limit".to_string());
        }
        Ok(count as usize)
    }

    fn range_cell_count_large(range: CellRange) -> Result<u64, String> {
        let rows = u64::from(range.end_row - range.start_row) + 1;
        let columns = u64::from(range.end_column - range.start_column) + 1;
        rows.checked_mul(columns)
            .ok_or_else(|| "cell range is too large".to_string())
    }

    fn range_value(&self, range: CellRange) -> Result<Value, String> {
        Self::range_cell_count(range)?;
        if range.is_single() {
            return Ok(self.cell_value(range.addresses().next().unwrap()));
        }
        Ok(Value::Array(ArrayValue {
            dimensions: vec![
                ArrayDimension {
                    lower_bound: 1,
                    length: (range.end_row - range.start_row + 1) as usize,
                },
                ArrayDimension {
                    lower_bound: 1,
                    length: (range.end_column - range.start_column + 1) as usize,
                },
            ],
            values: range
                .addresses()
                .map(|address| self.cell_value(address))
                .collect(),
            element_default: Box::new(Value::Empty),
            resizable: true,
        }))
    }

    fn cell_formula(&self, address: CellAddress) -> Value {
        self.workbook
            .sheets
            .get(address.sheet)
            .and_then(|sheet| sheet.rows.iter().find(|row| row.index == address.row))
            .and_then(|row| row.cells.iter().find(|cell| cell.col == address.column))
            .and_then(|cell| cell.formula.as_deref())
            .map(|formula| Value::String(format!("={formula}")))
            .unwrap_or_else(|| Value::String(String::new()))
    }

    fn range_formula(&self, range: CellRange) -> Result<Value, String> {
        Self::range_cell_count(range)?;
        if range.is_single() {
            return Ok(self.cell_formula(range.addresses().next().unwrap()));
        }
        Ok(Value::Array(ArrayValue {
            dimensions: vec![
                ArrayDimension {
                    lower_bound: 1,
                    length: (range.end_row - range.start_row + 1) as usize,
                },
                ArrayDimension {
                    lower_bound: 1,
                    length: (range.end_column - range.start_column + 1) as usize,
                },
            ],
            values: range
                .addresses()
                .map(|address| self.cell_formula(address))
                .collect(),
            element_default: Box::new(Value::String(String::new())),
            resizable: true,
        }))
    }

    fn range_has_formula(&self, range: CellRange) -> Result<Value, String> {
        let total = Self::range_cell_count_large(range)?;
        let sheet = self
            .workbook
            .sheets
            .get(range.sheet)
            .ok_or_else(|| "worksheet no longer exists".to_string())?;
        let formulas = sheet
            .rows
            .iter()
            .filter(|row| (range.start_row..=range.end_row).contains(&row.index))
            .flat_map(|row| &row.cells)
            .filter(|cell| {
                (range.start_column..=range.end_column).contains(&cell.col)
                    && cell.formula.is_some()
            })
            .count() as u64;
        Ok(if formulas == 0 {
            Value::Boolean(false)
        } else if formulas == total {
            Value::Boolean(true)
        } else {
            Value::Null
        })
    }

    fn range_end(&mut self, range: CellRange, args: &[Value]) -> Result<Value, String> {
        let [direction] = args else {
            return Err("Range.End expects one direction".to_string());
        };
        let direction = end_direction(direction)?;
        let worksheet = self
            .workbook
            .sheets
            .get(range.sheet)
            .ok_or_else(|| "worksheet no longer exists".to_string())?;
        let start = CellAddress {
            sheet: range.sheet,
            row: range.start_row,
            column: range.start_column,
        };
        let destination = match direction {
            EndDirection::Up | EndDirection::Down => {
                let mut occupied = worksheet
                    .rows
                    .iter()
                    .filter(|row| {
                        row.cells
                            .iter()
                            .any(|cell| cell.col == start.column && cell_has_content(cell))
                    })
                    .map(|row| row.index)
                    .collect::<Vec<_>>();
                occupied.sort_unstable();
                occupied.dedup();
                CellAddress {
                    row: ctrl_arrow_destination(
                        &occupied,
                        start.row,
                        1,
                        MAX_WORKSHEET_ROW,
                        matches!(direction, EndDirection::Down),
                    ),
                    ..start
                }
            }
            EndDirection::Left | EndDirection::Right => {
                let mut occupied = worksheet
                    .rows
                    .iter()
                    .find(|row| row.index == start.row)
                    .map(|row| {
                        row.cells
                            .iter()
                            .filter(|cell| cell_has_content(cell))
                            .map(|cell| cell.col)
                            .collect::<Vec<_>>()
                    })
                    .unwrap_or_default();
                occupied.sort_unstable();
                occupied.dedup();
                CellAddress {
                    column: ctrl_arrow_destination(
                        &occupied,
                        start.column,
                        0,
                        MAX_WORKSHEET_COLUMN,
                        matches!(direction, EndDirection::Right),
                    ),
                    ..start
                }
            }
        };
        Ok(self.object(HostObject::Range(CellRange::single(destination))))
    }

    fn current_region_object(&mut self, range: CellRange) -> Result<Value, String> {
        let worksheet = self
            .workbook
            .sheets
            .get(range.sheet)
            .ok_or_else(|| "worksheet no longer exists".to_string())?;
        let mut occupied_rows = BTreeMap::<u32, BTreeSet<u32>>::new();
        let mut occupied_columns = BTreeMap::<u32, BTreeSet<u32>>::new();
        for row in &worksheet.rows {
            for cell in row.cells.iter().filter(|cell| cell_has_content(cell)) {
                occupied_rows.entry(row.index).or_default().insert(cell.col);
                occupied_columns
                    .entry(cell.col)
                    .or_default()
                    .insert(row.index);
            }
        }
        let mut region = range;
        loop {
            let outer_start_row = region.start_row.saturating_sub(1).max(1);
            let outer_start_column = region.start_column.saturating_sub(1);
            let outer_end_row = region.end_row.saturating_add(1).min(MAX_WORKSHEET_ROW);
            let outer_end_column = region
                .end_column
                .saturating_add(1)
                .min(MAX_WORKSHEET_COLUMN);
            let row_has_content = |row, start_column, end_column| {
                occupied_rows.get(&row).is_some_and(|columns| {
                    columns.range(start_column..=end_column).next().is_some()
                })
            };
            let column_has_content = |column, start_row, end_row| {
                occupied_columns
                    .get(&column)
                    .is_some_and(|rows| rows.range(start_row..=end_row).next().is_some())
            };
            let expand_top = region.start_row > 1
                && row_has_content(region.start_row - 1, outer_start_column, outer_end_column);
            let expand_bottom = region.end_row < MAX_WORKSHEET_ROW
                && row_has_content(region.end_row + 1, outer_start_column, outer_end_column);
            let expand_left = region.start_column > 0
                && column_has_content(region.start_column - 1, outer_start_row, outer_end_row);
            let expand_right = region.end_column < MAX_WORKSHEET_COLUMN
                && column_has_content(region.end_column + 1, outer_start_row, outer_end_row);
            if !expand_top && !expand_bottom && !expand_left && !expand_right {
                break;
            }
            region.start_row -= u32::from(expand_top);
            region.end_row += u32::from(expand_bottom);
            region.start_column -= u32::from(expand_left);
            region.end_column += u32::from(expand_right);
        }
        Ok(self.object(HostObject::Range(region)))
    }

    fn offset_range(&mut self, range: CellRange, args: &[Value]) -> Result<Value, String> {
        let (row_offset, column_offset) = match args {
            [] => (0, 0),
            [row] => (integer_offset(row, "row")?, 0),
            [row, column] => (
                integer_offset(row, "row")?,
                integer_offset(column, "column")?,
            ),
            _ => return Err("Range.Offset expects zero, one, or two offsets".to_string()),
        };
        let shift = |value: u32, offset: i64, label: &str| {
            i64::from(value)
                .checked_add(offset)
                .and_then(|value| u32::try_from(value).ok())
                .ok_or_else(|| format!("Range.Offset moves the {label} outside the worksheet"))
        };
        let start_row = shift(range.start_row, row_offset, "row")?;
        let end_row = shift(range.end_row, row_offset, "row")?;
        let start_column = shift(range.start_column, column_offset, "column")?;
        let end_column = shift(range.end_column, column_offset, "column")?;
        if start_row == 0 {
            return Err("Range.Offset moves the row outside the worksheet".to_string());
        }
        Ok(self.object(HostObject::Range(CellRange {
            sheet: range.sheet,
            start_row,
            start_column,
            end_row,
            end_column,
        })))
    }

    fn resize_range(&mut self, range: CellRange, args: &[Value]) -> Result<Value, String> {
        let current_columns = range.end_column - range.start_column + 1;
        let (rows, columns) = match args {
            [] => (range.end_row - range.start_row + 1, current_columns),
            [rows] => (positive_index(rows, "row size")?, current_columns),
            [rows, columns] => (
                positive_index(rows, "row size")?,
                positive_index(columns, "column size")?,
            ),
            _ => return Err("Range.Resize expects zero, one, or two sizes".to_string()),
        };
        let end_row = range
            .start_row
            .checked_add(rows - 1)
            .ok_or_else(|| "Range.Resize row size is too large".to_string())?;
        let end_column = range
            .start_column
            .checked_add(columns - 1)
            .ok_or_else(|| "Range.Resize column size is too large".to_string())?;
        Ok(self.object(HostObject::Range(CellRange {
            sheet: range.sheet,
            start_row: range.start_row,
            start_column: range.start_column,
            end_row,
            end_column,
        })))
    }

    fn set_cell_value(&mut self, address: CellAddress, value: CellValue) -> Result<(), String> {
        let sheet = self
            .workbook
            .sheets
            .get_mut(address.sheet)
            .ok_or_else(|| "worksheet no longer exists".to_string())?;
        sheet.col_count = sheet.col_count.max(address.column as usize + 1);
        let row_position = sheet.rows.iter().position(|row| row.index == address.row);
        let row = match row_position {
            Some(position) => &mut sheet.rows[position],
            None => {
                sheet.rows.push(Row {
                    index: address.row,
                    cells: Vec::new(),
                    height: None,
                });
                sheet.rows.sort_by_key(|row| row.index);
                sheet
                    .rows
                    .iter_mut()
                    .find(|row| row.index == address.row)
                    .unwrap()
            }
        };
        match row.cells.iter_mut().find(|cell| cell.col == address.column) {
            Some(cell) => {
                cell.value = value;
                cell.formula = None;
            }
            None => {
                row.cells.push(Cell {
                    col: address.column,
                    value,
                    style: CellStyle::default(),
                    formula: None,
                });
                row.cells.sort_by_key(|cell| cell.col);
            }
        }
        Ok(())
    }

    fn set_cell_formula(&mut self, address: CellAddress, formula: String) -> Result<(), String> {
        let sheet = self
            .workbook
            .sheets
            .get_mut(address.sheet)
            .ok_or_else(|| "worksheet no longer exists".to_string())?;
        sheet.col_count = sheet.col_count.max(address.column as usize + 1);
        let row = match sheet.rows.iter().position(|row| row.index == address.row) {
            Some(position) => &mut sheet.rows[position],
            None => {
                sheet.rows.push(Row {
                    index: address.row,
                    cells: Vec::new(),
                    height: None,
                });
                sheet.rows.sort_by_key(|row| row.index);
                sheet
                    .rows
                    .iter_mut()
                    .find(|row| row.index == address.row)
                    .unwrap()
            }
        };
        let formula = formula.strip_prefix('=').unwrap_or(&formula).to_string();
        let formula = (!formula.is_empty()).then_some(formula);
        match row.cells.iter_mut().find(|cell| cell.col == address.column) {
            Some(cell) => {
                cell.value = CellValue::Empty;
                cell.formula = formula;
            }
            None => {
                row.cells.push(Cell {
                    col: address.column,
                    value: CellValue::Empty,
                    style: CellStyle::default(),
                    formula,
                });
                row.cells.sort_by_key(|cell| cell.col);
            }
        }
        Ok(())
    }

    fn set_range_value(&mut self, range: CellRange, value: Value) -> Result<(), String> {
        let count = Self::range_cell_count(range)?;
        match value {
            Value::Array(array) => {
                validate_range_array_shape(range, &array, "range assignment")?;
                if array.values.len() != count {
                    return Err(format!(
                        "range assignment needs {count} values, but the array contains {}",
                        array.values.len()
                    ));
                }
                for (address, value) in range.addresses().zip(array.values) {
                    self.set_cell_value(address, to_cell_value(value)?)?;
                }
            }
            value => {
                let value = to_cell_value(value)?;
                for address in range.addresses() {
                    self.set_cell_value(address, value.clone())?;
                }
            }
        }
        Ok(())
    }

    fn set_range_formula(&mut self, range: CellRange, value: Value) -> Result<(), String> {
        let count = Self::range_cell_count(range)?;
        let formulas = match value {
            Value::Array(array) => {
                validate_range_array_shape(range, &array, "range formula assignment")?;
                if array.values.len() != count {
                    return Err(format!(
                        "range formula assignment needs {count} values, but the array contains {}",
                        array.values.len()
                    ));
                }
                array
                    .values
                    .into_iter()
                    .map(to_formula)
                    .collect::<Result<Vec<_>, _>>()?
            }
            value => vec![to_formula(value)?; count],
        };
        for (address, formula) in range.addresses().zip(formulas) {
            self.set_cell_formula(address, formula)?;
        }
        Ok(())
    }

    fn set_range_style(
        &mut self,
        range: CellRange,
        mut update: impl FnMut(CellAddress, &mut CellStyle),
    ) -> Result<(), String> {
        Self::range_cell_count(range)?;
        for address in range.addresses() {
            let sheet = self
                .workbook
                .sheets
                .get_mut(address.sheet)
                .ok_or_else(|| "worksheet no longer exists".to_string())?;
            sheet.col_count = sheet.col_count.max(address.column as usize + 1);
            if !sheet.rows.iter().any(|row| row.index == address.row) {
                sheet.rows.push(Row {
                    index: address.row,
                    cells: Vec::new(),
                    height: None,
                });
                sheet.rows.sort_by_key(|row| row.index);
            }
            let row = sheet
                .rows
                .iter_mut()
                .find(|row| row.index == address.row)
                .expect("the destination row was created");
            if !row.cells.iter().any(|cell| cell.col == address.column) {
                row.cells.push(Cell {
                    col: address.column,
                    value: CellValue::Empty,
                    style: CellStyle::default(),
                    formula: None,
                });
                row.cells.sort_by_key(|cell| cell.col);
            }
            let cell = row
                .cells
                .iter_mut()
                .find(|cell| cell.col == address.column)
                .expect("the destination cell was created");
            update(address, &mut cell.style);
        }
        Ok(())
    }

    fn uniform_style<T: Clone + PartialEq>(
        &self,
        range: CellRange,
        read: impl Fn(&CellStyle) -> T,
    ) -> Result<Option<T>, String> {
        Self::range_cell_count(range)?;
        let default_style = CellStyle::default();
        let mut first = None;
        for address in range.addresses() {
            let style = self
                .workbook
                .sheets
                .get(address.sheet)
                .and_then(|sheet| sheet.rows.iter().find(|row| row.index == address.row))
                .and_then(|row| row.cells.iter().find(|cell| cell.col == address.column))
                .map(|cell| &cell.style)
                .unwrap_or(&default_style);
            let value = read(style);
            if first.as_ref().is_some_and(|first| first != &value) {
                return Ok(None);
            }
            first = Some(value);
        }
        Ok(first)
    }

    fn uniform_border(
        &self,
        range: CellRange,
        selection: BorderSelection,
    ) -> Result<Option<bool>, String> {
        Self::range_cell_count(range)?;
        let default_style = CellStyle::default();
        let mut first = None;
        for address in range.addresses() {
            let style = self
                .workbook
                .sheets
                .get(address.sheet)
                .and_then(|sheet| sheet.rows.iter().find(|row| row.index == address.row))
                .and_then(|row| row.cells.iter().find(|cell| cell.col == address.column))
                .map(|cell| &cell.style)
                .unwrap_or(&default_style);
            for value in selected_borders(style, address, range, selection) {
                if first.is_some_and(|first| first != value) {
                    return Ok(None);
                }
                first = Some(value);
            }
        }
        Ok(Some(first.unwrap_or(false)))
    }

    fn range_column_width(&self, range: CellRange) -> Result<Value, String> {
        let sheet = self
            .workbook
            .sheets
            .get(range.sheet)
            .ok_or_else(|| "worksheet no longer exists".to_string())?;
        let mut first = None;
        for column in range.start_column..=range.end_column {
            let width = sheet
                .col_widths
                .get(column as usize)
                .copied()
                .unwrap_or(sheet.default_col_width);
            if first.is_some_and(|first| first != width) {
                return Ok(Value::Null);
            }
            first = Some(width);
        }
        Ok(Value::Double(f64::from(
            first.unwrap_or(sheet.default_col_width),
        )))
    }

    fn set_range_column_width(&mut self, range: CellRange, value: Value) -> Result<(), String> {
        let sheet = self
            .workbook
            .sheets
            .get_mut(range.sheet)
            .ok_or_else(|| "worksheet no longer exists".to_string())?;
        let width = optional_dimension(value, "Range.ColumnWidth", 255.0)?;
        let required = range.end_column as usize + 1;
        if sheet.col_widths.len() < required {
            sheet.col_widths.resize(required, sheet.default_col_width);
        }
        let width = width.unwrap_or(sheet.default_col_width);
        for column in range.start_column..=range.end_column {
            sheet.col_widths[column as usize] = width;
        }
        sheet.col_count = sheet.col_count.max(required);
        Ok(())
    }

    fn range_row_height(&self, range: CellRange) -> Result<Value, String> {
        let sheet = self
            .workbook
            .sheets
            .get(range.sheet)
            .ok_or_else(|| "worksheet no longer exists".to_string())?;
        let selected_rows = u64::from(range.end_row - range.start_row) + 1;
        let rows = sheet
            .rows
            .iter()
            .filter(|row| (range.start_row..=range.end_row).contains(&row.index));
        let existing_rows = rows.clone().count() as u64;
        let mut first = (existing_rows < selected_rows).then_some(sheet.default_row_height);
        for row in rows {
            let height = row.height.unwrap_or(sheet.default_row_height);
            if first.is_some_and(|first| first != height) {
                return Ok(Value::Null);
            }
            first = Some(height);
        }
        Ok(Value::Double(f64::from(
            first.unwrap_or(sheet.default_row_height),
        )))
    }

    fn set_range_row_height(&mut self, range: CellRange, value: Value) -> Result<(), String> {
        let sheet = self
            .workbook
            .sheets
            .get_mut(range.sheet)
            .ok_or_else(|| "worksheet no longer exists".to_string())?;
        let height = optional_dimension(value, "Range.RowHeight", 409.5)?;
        if range.start_row == 1 && range.end_row == MAX_WORKSHEET_ROW {
            if let Some(height) = height {
                sheet.default_row_height = height;
            }
            for row in &mut sheet.rows {
                row.height = None;
            }
            return Ok(());
        }
        for row_index in range.start_row..=range.end_row {
            if !sheet.rows.iter().any(|row| row.index == row_index) {
                sheet.rows.push(Row {
                    index: row_index,
                    cells: Vec::new(),
                    height: None,
                });
            }
            let row = sheet
                .rows
                .iter_mut()
                .find(|row| row.index == row_index)
                .expect("the resized row was created");
            row.height = height;
        }
        sheet.rows.sort_by_key(|row| row.index);
        Ok(())
    }

    fn clear_range(
        &mut self,
        range: CellRange,
        clear_contents: bool,
        clear_formats: bool,
    ) -> Result<(), String> {
        Self::range_cell_count(range)?;
        let sheet = self
            .workbook
            .sheets
            .get_mut(range.sheet)
            .ok_or_else(|| "worksheet no longer exists".to_string())?;
        for row in &mut sheet.rows {
            if !(range.start_row..=range.end_row).contains(&row.index) {
                continue;
            }
            for cell in &mut row.cells {
                if !(range.start_column..=range.end_column).contains(&cell.col) {
                    continue;
                }
                if clear_contents {
                    cell.value = CellValue::Empty;
                    cell.formula = None;
                }
                if clear_formats {
                    cell.style = CellStyle::default();
                }
            }
        }
        Ok(())
    }

    fn merge_range(&mut self, range: CellRange) -> Result<(), String> {
        if range.is_single() {
            return Ok(());
        }
        let sheet = self
            .workbook
            .sheets
            .get_mut(range.sheet)
            .ok_or_else(|| "worksheet no longer exists".to_string())?;
        for existing in &sheet.merge_cells {
            let existing = merge_range(range.sheet, existing);
            if ranges_overlap(range, existing) {
                if ranges_equal(range, existing) {
                    return Ok(());
                }
                return Err("Range.Merge overlaps an existing merged range".to_string());
            }
        }
        for row in &mut sheet.rows {
            if !(range.start_row..=range.end_row).contains(&row.index) {
                continue;
            }
            for cell in &mut row.cells {
                if (range.start_column..=range.end_column).contains(&cell.col)
                    && !(row.index == range.start_row && cell.col == range.start_column)
                {
                    cell.value = CellValue::Empty;
                    cell.formula = None;
                }
            }
        }
        sheet.merge_cells.push(MergeCell {
            start_row: range.start_row,
            start_col: range.start_column,
            end_row: range.end_row,
            end_col: range.end_column,
        });
        sheet.merge_cells.sort_by_key(|merge| {
            (
                merge.start_row,
                merge.start_col,
                merge.end_row,
                merge.end_col,
            )
        });
        Ok(())
    }

    fn unmerge_range(&mut self, range: CellRange) -> Result<(), String> {
        let sheet = self
            .workbook
            .sheets
            .get_mut(range.sheet)
            .ok_or_else(|| "worksheet no longer exists".to_string())?;
        sheet
            .merge_cells
            .retain(|merge| !ranges_overlap(range, merge_range(range.sheet, merge)));
        Ok(())
    }

    fn range_merge_state(&self, range: CellRange) -> Result<Value, String> {
        let sheet = self
            .workbook
            .sheets
            .get(range.sheet)
            .ok_or_else(|| "worksheet no longer exists".to_string())?;
        let mut overlaps = sheet
            .merge_cells
            .iter()
            .map(|merge| merge_range(range.sheet, merge))
            .filter(|merge| ranges_overlap(range, *merge));
        let Some(merge) = overlaps.next() else {
            return Ok(Value::Boolean(false));
        };
        if overlaps.next().is_none() && range_contains(merge, range) {
            Ok(Value::Boolean(true))
        } else {
            Ok(Value::Null)
        }
    }

    fn copy_range(&mut self, source: CellRange, args: &[Value]) -> Result<Value, String> {
        let [Value::Object(destination)] = args else {
            return Err(
                "Range.Copy expects one destination Range; the browser clipboard is unavailable"
                    .to_string(),
            );
        };
        let destination = self
            .range(destination)
            .ok_or_else(|| "Range.Copy destination must be a Range".to_string())?;
        Self::range_cell_count(source)?;
        let row_count = source.end_row - source.start_row + 1;
        let column_count = source.end_column - source.start_column + 1;
        let end_row = destination
            .start_row
            .checked_add(row_count - 1)
            .filter(|row| *row <= MAX_WORKSHEET_ROW)
            .ok_or_else(|| {
                "Range.Copy destination extends beyond the worksheet rows".to_string()
            })?;
        let end_column = destination
            .start_column
            .checked_add(column_count - 1)
            .filter(|column| *column <= MAX_WORKSHEET_COLUMN)
            .ok_or_else(|| {
                "Range.Copy destination extends beyond the worksheet columns".to_string()
            })?;
        let destination = CellRange {
            sheet: destination.sheet,
            start_row: destination.start_row,
            start_column: destination.start_column,
            end_row,
            end_column,
        };
        let row_offset = i64::from(destination.start_row) - i64::from(source.start_row);
        let column_offset = i64::from(destination.start_column) - i64::from(source.start_column);
        let worksheet = self
            .workbook
            .sheets
            .get(source.sheet)
            .ok_or_else(|| "worksheet no longer exists".to_string())?;
        let copied = source
            .addresses()
            .map(|address| {
                worksheet
                    .rows
                    .iter()
                    .find(|row| row.index == address.row)
                    .and_then(|row| row.cells.iter().find(|cell| cell.col == address.column))
                    .map(|cell| {
                        let formula = cell
                            .formula
                            .as_deref()
                            .map(|formula| {
                                oxicells_core::translate_formula_references(
                                    formula,
                                    row_offset,
                                    column_offset,
                                )
                                .map_err(|error| {
                                    format!("Range.Copy cannot adjust formula {formula:?}: {error}")
                                })
                            })
                            .transpose()?;
                        Ok::<_, String>((cell.value.clone(), cell.style.clone(), formula))
                    })
                    .unwrap_or_else(|| Ok((CellValue::Empty, CellStyle::default(), None)))
            })
            .collect::<Result<Vec<_>, _>>()?;
        for (address, (value, style, formula)) in destination.addresses().zip(copied) {
            self.set_cell_value(address, value)?;
            let sheet = &mut self.workbook.sheets[address.sheet];
            let cell = sheet
                .rows
                .iter_mut()
                .find(|row| row.index == address.row)
                .and_then(|row| row.cells.iter_mut().find(|cell| cell.col == address.column))
                .expect("set_cell_value creates the destination cell");
            cell.style = style;
            if formula.is_some() {
                cell.value = CellValue::Empty;
            }
            cell.formula = formula;
        }
        Ok(Value::Empty)
    }
}

fn validate_range_array_shape(
    range: CellRange,
    array: &ArrayValue,
    operation: &str,
) -> Result<(), String> {
    if array.dimensions.len() == 1 {
        return Ok(());
    }
    let rows = (range.end_row - range.start_row + 1) as usize;
    let columns = (range.end_column - range.start_column + 1) as usize;
    if array.dimensions.len() == 2
        && array.dimensions[0].length == rows
        && array.dimensions[1].length == columns
    {
        return Ok(());
    }
    let shape = array
        .dimensions
        .iter()
        .map(|dimension| dimension.length.to_string())
        .collect::<Vec<_>>()
        .join(" x ");
    Err(format!(
        "{operation} needs a {rows} x {columns} array, but received {shape}"
    ))
}

impl Host for WorkbookHost<'_> {
    fn call(
        &mut self,
        receiver: Option<&ObjectRef>,
        name: &str,
        args: &[Value],
    ) -> Result<Option<Value>, String> {
        if let Some(receiver) = receiver {
            if self.is_debug_console(receiver) && name.eq_ignore_ascii_case("print") {
                self.debug_output.push(
                    args.iter()
                        .map(format_debug_value)
                        .collect::<Vec<_>>()
                        .join("\t"),
                );
                return Ok(Some(Value::Empty));
            }
            if let Some(sheet) = self.worksheet(receiver) {
                if name.eq_ignore_ascii_case("evaluate") {
                    return self.evaluate_object(sheet, args).map(Some);
                }
                if name.eq_ignore_ascii_case("range") {
                    return self.range_object(sheet, args).map(Some);
                }
                if name.eq_ignore_ascii_case("cells") {
                    return self.cells_object(sheet, args).map(Some);
                }
                if name.eq_ignore_ascii_case("activate") {
                    if !args.is_empty() {
                        return Err("Worksheet.Activate does not accept arguments".to_string());
                    }
                    self.active_sheet = sheet;
                    self.selection = CellRange::single(CellAddress {
                        sheet,
                        row: 1,
                        column: 0,
                    });
                    return Ok(Some(Value::Empty));
                }
                if name.eq_ignore_ascii_case("usedrange") {
                    if !args.is_empty() {
                        return Err("Worksheet.UsedRange does not accept arguments".to_string());
                    }
                    return self.used_range_object(sheet).map(Some);
                }
                if name.eq_ignore_ascii_case("rows") {
                    return self
                        .worksheet_axis_object_or_item(sheet, RangeAxis::Rows, args)
                        .map(Some);
                }
                if name.eq_ignore_ascii_case("columns") {
                    return self
                        .worksheet_axis_object_or_item(sheet, RangeAxis::Columns, args)
                        .map(Some);
                }
                return Ok(None);
            }
            if (self.is_workbook(receiver) || self.is_application(receiver))
                && (name.eq_ignore_ascii_case("worksheets") || name.eq_ignore_ascii_case("sheets"))
            {
                return self.worksheets_object_or_item(args).map(Some);
            }
            if self.is_application(receiver) && name.eq_ignore_ascii_case("evaluate") {
                return self.evaluate_object(self.active_sheet, args).map(Some);
            }
            if self.is_application(receiver)
                && (name.eq_ignore_ascii_case("selection")
                    || name.eq_ignore_ascii_case("activecell"))
            {
                if !args.is_empty() {
                    return Err(format!("Application.{name} does not accept arguments"));
                }
                let range = if name.eq_ignore_ascii_case("activecell") {
                    CellRange::single(CellAddress {
                        sheet: self.selection.sheet,
                        row: self.selection.start_row,
                        column: self.selection.start_column,
                    })
                } else {
                    self.selection
                };
                return Ok(Some(self.object(HostObject::Range(range))));
            }
            if self.is_application(receiver) && name.eq_ignore_ascii_case("rows") {
                return self
                    .worksheet_axis_object_or_item(self.active_sheet, RangeAxis::Rows, args)
                    .map(Some);
            }
            if self.is_application(receiver) && name.eq_ignore_ascii_case("columns") {
                return self
                    .worksheet_axis_object_or_item(self.active_sheet, RangeAxis::Columns, args)
                    .map(Some);
            }
            if self.is_worksheets(receiver) && name.eq_ignore_ascii_case("item") {
                let [value] = args else {
                    return Err("Worksheets.Item expects one sheet name or index".to_string());
                };
                return self.worksheet_object(value).map(Some);
            }
            if let Some((range, axis)) = self.range_collection(receiver) {
                if name.eq_ignore_ascii_case("item") {
                    let [index] = args else {
                        return Err(format!(
                            "Range.{}.Item expects one index",
                            range_axis_name(axis)
                        ));
                    };
                    return self.range_collection_item(range, axis, index).map(Some);
                }
                return Ok(None);
            }
            if let Some(range) = self.range(receiver) {
                if name.eq_ignore_ascii_case("borders") {
                    return self.borders_object(range, args).map(Some);
                }
                if name.eq_ignore_ascii_case("copy") {
                    return self.copy_range(range, args).map(Some);
                }
                if name.eq_ignore_ascii_case("select") {
                    if !args.is_empty() {
                        return Err("Range.Select does not accept arguments".to_string());
                    }
                    if range.sheet != self.active_sheet {
                        return Err("Range.Select requires its worksheet to be active".to_string());
                    }
                    self.selection = range;
                    return Ok(Some(Value::Empty));
                }
                if name.eq_ignore_ascii_case("cells") {
                    return self.range_cells_object(range, args).map(Some);
                }
                if name.eq_ignore_ascii_case("offset") {
                    return self.offset_range(range, args).map(Some);
                }
                if name.eq_ignore_ascii_case("resize") {
                    return self.resize_range(range, args).map(Some);
                }
                if name.eq_ignore_ascii_case("end") {
                    return self.range_end(range, args).map(Some);
                }
                if name.eq_ignore_ascii_case("currentregion") {
                    if !args.is_empty() {
                        return Err("Range.CurrentRegion does not accept arguments".to_string());
                    }
                    return self.current_region_object(range).map(Some);
                }
                if name.eq_ignore_ascii_case("address") {
                    return range_address_from_args(range, args)
                        .map(Value::String)
                        .map(Some);
                }
                if name.eq_ignore_ascii_case("rows") {
                    return self
                        .range_collection_object_or_item(range, RangeAxis::Rows, args)
                        .map(Some);
                }
                if name.eq_ignore_ascii_case("columns") {
                    return self
                        .range_collection_object_or_item(range, RangeAxis::Columns, args)
                        .map(Some);
                }
                if name.eq_ignore_ascii_case("clearcontents") {
                    if !args.is_empty() {
                        return Err("Range.ClearContents does not accept arguments".to_string());
                    }
                    self.set_range_value(range, Value::Empty)?;
                    return Ok(Some(Value::Empty));
                }
                if name.eq_ignore_ascii_case("clearformats") {
                    if !args.is_empty() {
                        return Err("Range.ClearFormats does not accept arguments".to_string());
                    }
                    self.clear_range(range, false, true)?;
                    return Ok(Some(Value::Empty));
                }
                if name.eq_ignore_ascii_case("clear") {
                    if !args.is_empty() {
                        return Err("Range.Clear does not accept arguments".to_string());
                    }
                    self.clear_range(range, true, true)?;
                    return Ok(Some(Value::Empty));
                }
                if name.eq_ignore_ascii_case("merge") {
                    match args {
                        [] => {}
                        [across] => {
                            if style_boolean(across, "Range.Merge Across")? {
                                return Err(
                                    "Range.Merge Across:=True is not supported in the browser"
                                        .to_string(),
                                );
                            }
                        }
                        _ => return Err("Range.Merge expects zero or one argument".to_string()),
                    }
                    self.merge_range(range)?;
                    return Ok(Some(Value::Empty));
                }
                if name.eq_ignore_ascii_case("unmerge") {
                    if !args.is_empty() {
                        return Err("Range.UnMerge does not accept arguments".to_string());
                    }
                    self.unmerge_range(range)?;
                    return Ok(Some(Value::Empty));
                }
            }
            return Ok(None);
        }
        if args.is_empty() {
            if let Some(value) = host_constant(name) {
                return Ok(Some(value));
            }
        }
        if name.eq_ignore_ascii_case("msgbox") {
            return self.show_message_box(args).map(Some);
        }
        if name.eq_ignore_ascii_case("rgb") {
            return rgb_value(args).map(Some);
        }
        if name.eq_ignore_ascii_case("range") {
            return self.range_object(self.active_sheet, args).map(Some);
        }
        if name.eq_ignore_ascii_case("evaluate") {
            return self.evaluate_object(self.active_sheet, args).map(Some);
        }
        if name.eq_ignore_ascii_case("cells") {
            return self.cells_object(self.active_sheet, args).map(Some);
        }
        if name.eq_ignore_ascii_case("worksheets") || name.eq_ignore_ascii_case("sheets") {
            return self.worksheets_object_or_item(args).map(Some);
        }
        if name.eq_ignore_ascii_case("activesheet") {
            return Ok(Some(self.object(HostObject::Worksheet(self.active_sheet))));
        }
        if name.eq_ignore_ascii_case("thisworkbook") || name.eq_ignore_ascii_case("activeworkbook")
        {
            return Ok(Some(self.object(HostObject::Workbook)));
        }
        if name.eq_ignore_ascii_case("application") {
            return Ok(Some(self.object(HostObject::Application)));
        }
        if name.eq_ignore_ascii_case("selection") || name.eq_ignore_ascii_case("activecell") {
            if !args.is_empty() {
                return Err(format!("{name} does not accept arguments"));
            }
            let range = if name.eq_ignore_ascii_case("activecell") {
                CellRange::single(CellAddress {
                    sheet: self.selection.sheet,
                    row: self.selection.start_row,
                    column: self.selection.start_column,
                })
            } else {
                self.selection
            };
            return Ok(Some(self.object(HostObject::Range(range))));
        }
        if name.eq_ignore_ascii_case("debug") {
            if !args.is_empty() {
                return Err("Debug does not accept arguments".to_string());
            }
            return Ok(Some(self.object(HostObject::DebugConsole)));
        }
        if name.eq_ignore_ascii_case("usedrange") {
            if !args.is_empty() {
                return Err("UsedRange does not accept arguments".to_string());
            }
            return self.used_range_object(self.active_sheet).map(Some);
        }
        if name.eq_ignore_ascii_case("rows") {
            return self
                .worksheet_axis_object_or_item(self.active_sheet, RangeAxis::Rows, args)
                .map(Some);
        }
        if name.eq_ignore_ascii_case("columns") {
            return self
                .worksheet_axis_object_or_item(self.active_sheet, RangeAxis::Columns, args)
                .map(Some);
        }
        Ok(None)
    }

    fn get(&mut self, receiver: &ObjectRef, name: &str) -> Result<Option<Value>, String> {
        if let Some((range, selection)) = self.range_borders(receiver) {
            if name.eq_ignore_ascii_case("linestyle") {
                return self.uniform_border(range, selection).map(|value| {
                    Some(match value {
                        Some(true) => Value::Integer(1),
                        Some(false) => Value::Integer(-4142),
                        None => Value::Null,
                    })
                });
            }
            return Ok(None);
        }
        if let Some(range) = self.range_font(receiver) {
            if name.eq_ignore_ascii_case("bold") {
                return self
                    .uniform_style(range, |style| style.bold)
                    .map(|value| Some(value.map(Value::Boolean).unwrap_or(Value::Null)));
            }
            if name.eq_ignore_ascii_case("italic") {
                return self
                    .uniform_style(range, |style| style.italic)
                    .map(|value| Some(value.map(Value::Boolean).unwrap_or(Value::Null)));
            }
            if name.eq_ignore_ascii_case("size") {
                return self
                    .uniform_style(range, |style| style.font_size)
                    .map(|value| {
                        Some(match value {
                            Some(Some(value)) => Value::Double(f64::from(value)),
                            Some(None) => Value::Empty,
                            None => Value::Null,
                        })
                    });
            }
            if name.eq_ignore_ascii_case("color") {
                return self
                    .uniform_style(range, |style| style.font_color.clone())
                    .map(|value| Some(style_color_value(value)));
            }
            return Ok(None);
        }
        if let Some(range) = self.range_interior(receiver) {
            if name.eq_ignore_ascii_case("color") {
                return self
                    .uniform_style(range, |style| style.bg_color.clone())
                    .map(|value| Some(style_color_value(value)));
            }
            return Ok(None);
        }
        if self.is_application(receiver) {
            if name.eq_ignore_ascii_case("activesheet") {
                return Ok(Some(self.object(HostObject::Worksheet(self.active_sheet))));
            }
            if name.eq_ignore_ascii_case("activeworkbook")
                || name.eq_ignore_ascii_case("thisworkbook")
            {
                return Ok(Some(self.object(HostObject::Workbook)));
            }
            if name.eq_ignore_ascii_case("worksheets") || name.eq_ignore_ascii_case("sheets") {
                return Ok(Some(self.object(HostObject::Worksheets)));
            }
            if name.eq_ignore_ascii_case("selection") || name.eq_ignore_ascii_case("activecell") {
                let range = if name.eq_ignore_ascii_case("activecell") {
                    CellRange::single(CellAddress {
                        sheet: self.selection.sheet,
                        row: self.selection.start_row,
                        column: self.selection.start_column,
                    })
                } else {
                    self.selection
                };
                return Ok(Some(self.object(HostObject::Range(range))));
            }
            if name.eq_ignore_ascii_case("rows") {
                return self
                    .worksheet_axis_object_or_item(self.active_sheet, RangeAxis::Rows, &[])
                    .map(Some);
            }
            if name.eq_ignore_ascii_case("columns") {
                return self
                    .worksheet_axis_object_or_item(self.active_sheet, RangeAxis::Columns, &[])
                    .map(Some);
            }
            return Ok(None);
        }
        if self.is_workbook(receiver)
            && (name.eq_ignore_ascii_case("worksheets") || name.eq_ignore_ascii_case("sheets"))
        {
            return Ok(Some(self.object(HostObject::Worksheets)));
        }
        if self.is_worksheets(receiver) {
            if name.eq_ignore_ascii_case("count") {
                return Ok(Some(Value::Integer(self.workbook.sheets.len() as i64)));
            }
            return Ok(None);
        }
        if let Some((range, axis)) = self.range_collection(receiver) {
            if name.eq_ignore_ascii_case("count") {
                let count = match axis {
                    RangeAxis::Rows => range.end_row - range.start_row + 1,
                    RangeAxis::Columns => range.end_column - range.start_column + 1,
                };
                return Ok(Some(Value::Integer(i64::from(count))));
            }
            return Ok(None);
        }
        if let Some(sheet) = self.worksheet(receiver) {
            if name.eq_ignore_ascii_case("name") {
                return Ok(Some(Value::String(
                    self.workbook.sheets[sheet].name.clone(),
                )));
            }
            if name.eq_ignore_ascii_case("index") {
                return Ok(Some(Value::Integer(sheet as i64 + 1)));
            }
            if name.eq_ignore_ascii_case("usedrange") {
                return self.used_range_object(sheet).map(Some);
            }
            if name.eq_ignore_ascii_case("rows") {
                return self
                    .worksheet_axis_object_or_item(sheet, RangeAxis::Rows, &[])
                    .map(Some);
            }
            if name.eq_ignore_ascii_case("columns") {
                return self
                    .worksheet_axis_object_or_item(sheet, RangeAxis::Columns, &[])
                    .map(Some);
            }
            return Ok(None);
        }
        let Some(range) = self.range(receiver) else {
            return Ok(None);
        };
        if name.eq_ignore_ascii_case("value") || name.eq_ignore_ascii_case("value2") {
            return self.range_value(range).map(Some);
        }
        if name.eq_ignore_ascii_case("formula") || name.eq_ignore_ascii_case("formula2") {
            return self.range_formula(range).map(Some);
        }
        if name.eq_ignore_ascii_case("hasformula") {
            return self.range_has_formula(range).map(Some);
        }
        if name.eq_ignore_ascii_case("parent") || name.eq_ignore_ascii_case("worksheet") {
            return Ok(Some(self.object(HostObject::Worksheet(range.sheet))));
        }
        if name.eq_ignore_ascii_case("entirerow") {
            return Ok(Some(self.object(HostObject::Range(CellRange {
                start_column: 0,
                end_column: MAX_WORKSHEET_COLUMN,
                ..range
            }))));
        }
        if name.eq_ignore_ascii_case("entirecolumn") {
            return Ok(Some(self.object(HostObject::Range(CellRange {
                start_row: 1,
                end_row: MAX_WORKSHEET_ROW,
                ..range
            }))));
        }
        if name.eq_ignore_ascii_case("font") {
            return Ok(Some(self.object(HostObject::RangeFont(range))));
        }
        if name.eq_ignore_ascii_case("interior") {
            return Ok(Some(self.object(HostObject::RangeInterior(range))));
        }
        if name.eq_ignore_ascii_case("borders") {
            return Ok(Some(
                self.object(HostObject::RangeBorders(range, BorderSelection::All)),
            ));
        }
        if name.eq_ignore_ascii_case("numberformat") {
            return self
                .uniform_style(range, |style| style.number_format.clone())
                .map(|value| {
                    Some(match value {
                        Some(Some(value)) => Value::String(value),
                        Some(None) => Value::String("General".to_string()),
                        None => Value::Null,
                    })
                });
        }
        if name.eq_ignore_ascii_case("horizontalalignment") {
            return self
                .uniform_style(range, |style| style.horizontal_align.clone())
                .map(|value| Some(horizontal_alignment_value(value)));
        }
        if name.eq_ignore_ascii_case("columnwidth") {
            return self.range_column_width(range).map(Some);
        }
        if name.eq_ignore_ascii_case("rowheight") {
            return self.range_row_height(range).map(Some);
        }
        if name.eq_ignore_ascii_case("mergecells") {
            return self.range_merge_state(range).map(Some);
        }
        if name.eq_ignore_ascii_case("row") {
            return Ok(Some(Value::Integer(i64::from(range.start_row))));
        }
        if name.eq_ignore_ascii_case("column") {
            return Ok(Some(Value::Integer(i64::from(range.start_column) + 1)));
        }
        if name.eq_ignore_ascii_case("count") || name.eq_ignore_ascii_case("countlarge") {
            return Self::range_cell_count_large(range)
                .map(|count| Some(Value::Integer(count as i64)));
        }
        if name.eq_ignore_ascii_case("address") {
            return Ok(Some(Value::String(format_range_address(range, true, true))));
        }
        if name.eq_ignore_ascii_case("currentregion") {
            return self.current_region_object(range).map(Some);
        }
        if name.eq_ignore_ascii_case("rows") {
            return Ok(Some(
                self.object(HostObject::RangeCollection(range, RangeAxis::Rows)),
            ));
        }
        if name.eq_ignore_ascii_case("columns") {
            return Ok(Some(
                self.object(HostObject::RangeCollection(range, RangeAxis::Columns)),
            ));
        }
        Ok(None)
    }

    fn set(&mut self, receiver: &ObjectRef, name: &str, value: Value) -> Result<bool, String> {
        if let Some((range, selection)) = self.range_borders(receiver) {
            if name.eq_ignore_ascii_case("linestyle") {
                let enabled = border_line_style(&value)?;
                self.set_range_style(range, |address, style| {
                    set_selected_borders(style, address, range, selection, enabled);
                })?;
                return Ok(true);
            }
            return Ok(false);
        }
        if let Some(range) = self.range_font(receiver) {
            if name.eq_ignore_ascii_case("bold") {
                let value = style_boolean(&value, "Font.Bold")?;
                self.set_range_style(range, |_, style| style.bold = value)?;
                return Ok(true);
            }
            if name.eq_ignore_ascii_case("italic") {
                let value = style_boolean(&value, "Font.Italic")?;
                self.set_range_style(range, |_, style| style.italic = value)?;
                return Ok(true);
            }
            if name.eq_ignore_ascii_case("size") {
                let value = font_size(&value)?;
                self.set_range_style(range, |_, style| style.font_size = value)?;
                return Ok(true);
            }
            if name.eq_ignore_ascii_case("color") {
                let value = style_color(&value, "Font.Color")?;
                self.set_range_style(range, |_, style| style.font_color = value.clone())?;
                return Ok(true);
            }
            return Ok(false);
        }
        if let Some(range) = self.range_interior(receiver) {
            if name.eq_ignore_ascii_case("color") {
                let value = style_color(&value, "Interior.Color")?;
                self.set_range_style(range, |_, style| style.bg_color = value.clone())?;
                return Ok(true);
            }
            return Ok(false);
        }
        let Some(range) = self.range(receiver) else {
            return Ok(false);
        };
        if name.eq_ignore_ascii_case("value") || name.eq_ignore_ascii_case("value2") {
            self.set_range_value(range, value)?;
            return Ok(true);
        }
        if name.eq_ignore_ascii_case("formula") || name.eq_ignore_ascii_case("formula2") {
            self.set_range_formula(range, value)?;
            return Ok(true);
        }
        if name.eq_ignore_ascii_case("numberformat") {
            let value = match value {
                Value::Empty => None,
                Value::String(value) if value.eq_ignore_ascii_case("general") => None,
                Value::String(value) => Some(value),
                _ => return Err("Range.NumberFormat must be a string".to_string()),
            };
            self.set_range_style(range, |_, style| style.number_format = value.clone())?;
            return Ok(true);
        }
        if name.eq_ignore_ascii_case("horizontalalignment") {
            let value = horizontal_alignment(&value)?;
            self.set_range_style(range, |_, style| style.horizontal_align = value.clone())?;
            return Ok(true);
        }
        if name.eq_ignore_ascii_case("columnwidth") {
            self.set_range_column_width(range, value)?;
            return Ok(true);
        }
        if name.eq_ignore_ascii_case("rowheight") {
            self.set_range_row_height(range, value)?;
            return Ok(true);
        }
        if name.eq_ignore_ascii_case("mergecells") {
            if style_boolean(&value, "Range.MergeCells")? {
                self.merge_range(range)?;
            } else {
                self.unmerge_range(range)?;
            }
            return Ok(true);
        }
        Ok(false)
    }

    fn enumerate(&mut self, receiver: &ObjectRef) -> Result<Option<Vec<Value>>, String> {
        if self.is_worksheets(receiver) {
            let mut worksheets = Vec::with_capacity(self.workbook.sheets.len());
            for sheet in 0..self.workbook.sheets.len() {
                worksheets.push(self.object(HostObject::Worksheet(sheet)));
            }
            return Ok(Some(worksheets));
        }
        if let Some((range, axis)) = self.range_collection(receiver) {
            Self::range_cell_count(range)?;
            let count = match axis {
                RangeAxis::Rows => range.end_row - range.start_row + 1,
                RangeAxis::Columns => range.end_column - range.start_column + 1,
            };
            let mut items = Vec::with_capacity(count as usize);
            for index in 1..=count {
                items.push(self.range_collection_item(
                    range,
                    axis,
                    &Value::Integer(i64::from(index)),
                )?);
            }
            return Ok(Some(items));
        }
        let Some(range) = self.range(receiver) else {
            return Ok(None);
        };
        Self::range_cell_count(range)?;
        let mut cells = Vec::new();
        for address in range.addresses() {
            cells.push(self.object(HostObject::Range(CellRange::single(address))));
        }
        Ok(Some(cells))
    }
}

fn format_debug_value(value: &Value) -> String {
    match value {
        Value::Empty => String::new(),
        Value::Missing => "Missing".to_string(),
        Value::Nothing => "Nothing".to_string(),
        Value::Null => "Null".to_string(),
        Value::Boolean(value) => if *value { "True" } else { "False" }.to_string(),
        Value::Integer(value) => value.to_string(),
        Value::Double(value) => value.to_string(),
        Value::Error(value) => format!("Error {value}"),
        Value::String(value) => value.clone(),
        Value::Array(_) => "<Array>".to_string(),
        Value::Object(value) => format!("<{}>", value.kind),
    }
}

fn style_boolean(value: &Value, property: &str) -> Result<bool, String> {
    match value {
        Value::Boolean(value) => Ok(*value),
        Value::Integer(value) => Ok(*value != 0),
        Value::Double(value) if value.is_finite() => Ok(*value != 0.0),
        _ => Err(format!("{property} must be Boolean")),
    }
}

fn optional_dimension(value: Value, property: &str, maximum: f64) -> Result<Option<f32>, String> {
    let number = match value {
        Value::Empty => return Ok(None),
        Value::Integer(value) => value as f64,
        Value::Double(value) => value,
        _ => return Err(format!("{property} must be numeric")),
    };
    if !number.is_finite() || !(0.0..=maximum).contains(&number) {
        return Err(format!("{property} must be between 0 and {maximum}"));
    }
    Ok(Some(number as f32))
}

fn font_size(value: &Value) -> Result<Option<f32>, String> {
    let number = match value {
        Value::Empty => return Ok(None),
        Value::Integer(value) => *value as f64,
        Value::Double(value) => *value,
        _ => return Err("Font.Size must be numeric".to_string()),
    };
    if !number.is_finite() || number <= 0.0 || number > f32::MAX as f64 {
        return Err("Font.Size must be a positive finite number".to_string());
    }
    Ok(Some(number as f32))
}

fn color_number(value: &Value, property: &str) -> Result<Option<u32>, String> {
    let number = match value {
        Value::Empty => return Ok(None),
        Value::Integer(value) => *value as f64,
        Value::Double(value) => *value,
        _ => return Err(format!("{property} must be an RGB color number")),
    };
    if !number.is_finite() || number.fract() != 0.0 || !(0.0..=16_777_215.0).contains(&number) {
        return Err(format!(
            "{property} must be an RGB color number from 0 to 16777215"
        ));
    }
    Ok(Some(number as u32))
}

fn style_color(value: &Value, property: &str) -> Result<Option<String>, String> {
    Ok(color_number(value, property)?.map(|color| {
        let red = color & 0xff;
        let green = (color >> 8) & 0xff;
        let blue = (color >> 16) & 0xff;
        format!("#{red:02x}{green:02x}{blue:02x}")
    }))
}

fn style_color_value(value: Option<Option<String>>) -> Value {
    let Some(value) = value else {
        return Value::Null;
    };
    let Some(value) = value else {
        return Value::Empty;
    };
    let Some(hex) = value.strip_prefix('#').filter(|hex| hex.len() == 6) else {
        return Value::Empty;
    };
    let Ok(rgb) = u32::from_str_radix(hex, 16) else {
        return Value::Empty;
    };
    let red = (rgb >> 16) & 0xff;
    let green = (rgb >> 8) & 0xff;
    let blue = rgb & 0xff;
    Value::Integer(i64::from(red | (green << 8) | (blue << 16)))
}

fn horizontal_alignment(value: &Value) -> Result<Option<String>, String> {
    let value = match value {
        Value::Empty => return Ok(None),
        Value::Integer(value) => *value,
        Value::Double(value) if value.is_finite() && value.fract() == 0.0 => *value as i64,
        _ => {
            return Err("Range.HorizontalAlignment must be an Excel alignment constant".to_string())
        }
    };
    match value {
        1 => Ok(None),
        -4131 => Ok(Some("left".to_string())),
        -4108 => Ok(Some("center".to_string())),
        -4152 => Ok(Some("right".to_string())),
        5 => Ok(Some("fill".to_string())),
        -4130 => Ok(Some("justify".to_string())),
        7 => Ok(Some("centerContinuous".to_string())),
        -4117 => Ok(Some("distributed".to_string())),
        _ => Err(format!(
            "unsupported Range.HorizontalAlignment constant: {value}"
        )),
    }
}

fn horizontal_alignment_value(value: Option<Option<String>>) -> Value {
    let Some(value) = value else {
        return Value::Null;
    };
    let Some(value) = value else {
        return Value::Integer(1);
    };
    let constant = match value.to_ascii_lowercase().as_str() {
        "left" => -4131,
        "center" => -4108,
        "right" => -4152,
        "fill" => 5,
        "justify" => -4130,
        "centercontinuous" => 7,
        "distributed" => -4117,
        _ => return Value::Empty,
    };
    Value::Integer(constant)
}

fn border_selection(value: &Value) -> Result<BorderSelection, String> {
    let value = match value {
        Value::Integer(value) => *value,
        Value::Double(value) if value.is_finite() && value.fract() == 0.0 => *value as i64,
        _ => return Err("Range.Borders index must be an Excel border constant".to_string()),
    };
    match value {
        7 => Ok(BorderSelection::EdgeLeft),
        8 => Ok(BorderSelection::EdgeTop),
        9 => Ok(BorderSelection::EdgeBottom),
        10 => Ok(BorderSelection::EdgeRight),
        11 => Ok(BorderSelection::InsideVertical),
        12 => Ok(BorderSelection::InsideHorizontal),
        _ => Err(format!("unsupported Range.Borders index: {value}")),
    }
}

fn border_line_style(value: &Value) -> Result<bool, String> {
    let value = match value {
        Value::Empty => return Ok(false),
        Value::Integer(value) => *value,
        Value::Double(value) if value.is_finite() && value.fract() == 0.0 => *value as i64,
        _ => return Err("Borders.LineStyle must be an Excel line-style constant".to_string()),
    };
    match value {
        1 => Ok(true),
        -4142 => Ok(false),
        _ => Err(format!("unsupported Borders.LineStyle constant: {value}")),
    }
}

fn selected_borders(
    style: &CellStyle,
    address: CellAddress,
    range: CellRange,
    selection: BorderSelection,
) -> Vec<bool> {
    match selection {
        BorderSelection::All => vec![
            style.border_top,
            style.border_bottom,
            style.border_left,
            style.border_right,
        ],
        BorderSelection::EdgeLeft if address.column == range.start_column => {
            vec![style.border_left]
        }
        BorderSelection::EdgeTop if address.row == range.start_row => vec![style.border_top],
        BorderSelection::EdgeBottom if address.row == range.end_row => vec![style.border_bottom],
        BorderSelection::EdgeRight if address.column == range.end_column => {
            vec![style.border_right]
        }
        BorderSelection::InsideVertical => {
            let mut values = Vec::with_capacity(2);
            if address.column > range.start_column {
                values.push(style.border_left);
            }
            if address.column < range.end_column {
                values.push(style.border_right);
            }
            values
        }
        BorderSelection::InsideHorizontal => {
            let mut values = Vec::with_capacity(2);
            if address.row > range.start_row {
                values.push(style.border_top);
            }
            if address.row < range.end_row {
                values.push(style.border_bottom);
            }
            values
        }
        _ => Vec::new(),
    }
}

fn set_selected_borders(
    style: &mut CellStyle,
    address: CellAddress,
    range: CellRange,
    selection: BorderSelection,
    enabled: bool,
) {
    match selection {
        BorderSelection::All => {
            style.border_top = enabled;
            style.border_bottom = enabled;
            style.border_left = enabled;
            style.border_right = enabled;
        }
        BorderSelection::EdgeLeft if address.column == range.start_column => {
            style.border_left = enabled;
        }
        BorderSelection::EdgeTop if address.row == range.start_row => {
            style.border_top = enabled;
        }
        BorderSelection::EdgeBottom if address.row == range.end_row => {
            style.border_bottom = enabled;
        }
        BorderSelection::EdgeRight if address.column == range.end_column => {
            style.border_right = enabled;
        }
        BorderSelection::InsideVertical => {
            if address.column > range.start_column {
                style.border_left = enabled;
            }
            if address.column < range.end_column {
                style.border_right = enabled;
            }
        }
        BorderSelection::InsideHorizontal => {
            if address.row > range.start_row {
                style.border_top = enabled;
            }
            if address.row < range.end_row {
                style.border_bottom = enabled;
            }
        }
        _ => {}
    }
}

fn rgb_value(args: &[Value]) -> Result<Value, String> {
    let [red, green, blue] = args else {
        return Err("RGB expects red, green, and blue arguments".to_string());
    };
    let component = |value: &Value, name: &str| -> Result<u32, String> {
        let value = match value {
            Value::Integer(value) => *value as f64,
            Value::Double(value) => *value,
            _ => return Err(format!("RGB {name} must be numeric")),
        };
        if !value.is_finite() || value.fract() != 0.0 || !(0.0..=255.0).contains(&value) {
            return Err(format!("RGB {name} must be an integer from 0 to 255"));
        }
        Ok(value as u32)
    };
    let red = component(red, "red")?;
    let green = component(green, "green")?;
    let blue = component(blue, "blue")?;
    Ok(Value::Integer(i64::from(red | (green << 8) | (blue << 16))))
}

fn cell_has_content(cell: &Cell) -> bool {
    cell.formula.is_some() || !matches!(&cell.value, CellValue::Empty)
}

fn merge_range(sheet: usize, merge: &MergeCell) -> CellRange {
    CellRange {
        sheet,
        start_row: merge.start_row,
        start_column: merge.start_col,
        end_row: merge.end_row,
        end_column: merge.end_col,
    }
}

fn ranges_overlap(left: CellRange, right: CellRange) -> bool {
    left.sheet == right.sheet
        && left.start_row <= right.end_row
        && right.start_row <= left.end_row
        && left.start_column <= right.end_column
        && right.start_column <= left.end_column
}

fn ranges_equal(left: CellRange, right: CellRange) -> bool {
    left.sheet == right.sheet
        && left.start_row == right.start_row
        && left.start_column == right.start_column
        && left.end_row == right.end_row
        && left.end_column == right.end_column
}

fn range_contains(outer: CellRange, inner: CellRange) -> bool {
    outer.sheet == inner.sheet
        && outer.start_row <= inner.start_row
        && outer.start_column <= inner.start_column
        && outer.end_row >= inner.end_row
        && outer.end_column >= inner.end_column
}

fn ctrl_arrow_destination(
    occupied: &[u32],
    start: u32,
    minimum: u32,
    maximum: u32,
    forward: bool,
) -> u32 {
    if start == if forward { maximum } else { minimum } {
        return start;
    }
    let start_index = occupied.binary_search(&start).ok();
    let neighbour = if forward { start + 1 } else { start - 1 };
    let neighbour_index = occupied.binary_search(&neighbour).ok();
    if let (Some(_), Some(mut index)) = (start_index, neighbour_index) {
        let mut destination = neighbour;
        if forward {
            while occupied.get(index + 1) == Some(&(destination + 1)) {
                index += 1;
                destination += 1;
            }
        } else {
            while index > 0 && occupied.get(index - 1) == Some(&(destination - 1)) {
                index -= 1;
                destination -= 1;
            }
        }
        return destination;
    }
    if forward {
        occupied
            .iter()
            .copied()
            .find(|position| *position > start)
            .unwrap_or(maximum)
    } else {
        occupied
            .iter()
            .rev()
            .copied()
            .find(|position| *position < start)
            .unwrap_or(minimum)
    }
}

fn end_direction(value: &Value) -> Result<EndDirection, String> {
    let value = match value {
        Value::Integer(value) => *value,
        Value::Double(value) if value.is_finite() && value.fract() == 0.0 => *value as i64,
        _ => return Err("Range.End direction must be an Excel direction constant".to_string()),
    };
    match value {
        -4162 => Ok(EndDirection::Up),
        -4121 => Ok(EndDirection::Down),
        -4159 => Ok(EndDirection::Left),
        -4161 => Ok(EndDirection::Right),
        _ => Err(format!("unsupported Range.End direction: {value}")),
    }
}

fn host_constant(name: &str) -> Option<Value> {
    let value = match name.to_ascii_lowercase().as_str() {
        "xlup" => -4162,
        "xldown" => -4121,
        "xltoleft" => -4159,
        "xltoright" => -4161,
        "xlgeneral" => 1,
        "xlleft" => -4131,
        "xlcenter" => -4108,
        "xlright" => -4152,
        "xlfill" => 5,
        "xljustify" => -4130,
        "xlcenteracrossselection" => 7,
        "xldistributed" => -4117,
        "xlcontinuous" => 1,
        "xllinestylenone" => -4142,
        "xledgeleft" => 7,
        "xledgetop" => 8,
        "xledgebottom" => 9,
        "xledgeright" => 10,
        "xlinsidevertical" => 11,
        "xlinsidehorizontal" => 12,
        "vbokonly" | "vbapplicationmodal" | "vbdefaultbutton1" => 0,
        "vbokcancel" | "vbok" => 1,
        "vbabortretryignore" | "vbcancel" => 2,
        "vbyesnocancel" | "vbabort" => 3,
        "vbyesno" | "vbretry" => 4,
        "vbretrycancel" | "vbignore" => 5,
        "vbyes" => 6,
        "vbno" => 7,
        "vbcritical" => 16,
        "vbquestion" => 32,
        "vbexclamation" => 48,
        "vbinformation" => 64,
        "vbdefaultbutton2" => 256,
        "vbdefaultbutton3" => 512,
        "vbdefaultbutton4" => 768,
        "vbsystemmodal" => 4096,
        "vbmsgboxhelpbutton" => 16_384,
        "vbmsgboxsetforeground" => 65_536,
        "vbmsgboxright" => 524_288,
        "vbmsgboxrtlreading" => 1_048_576,
        "vbblack" => 0,
        "vbred" => 255,
        "vbgreen" => 65_280,
        "vbyellow" => 65_535,
        "vbblue" => 16_711_680,
        "vbmagenta" => 16_711_935,
        "vbcyan" => 16_776_960,
        "vbwhite" => 16_777_215,
        _ => return None,
    };
    Some(Value::Integer(value))
}

fn parse_range_reference(reference: &str) -> Result<((u32, u32), (u32, u32)), String> {
    let mut parts = reference.split(':');
    let start = parts.next().unwrap_or_default();
    let end = parts.next();
    if parts.next().is_some() {
        return Err(format!("invalid range reference: {reference}"));
    }
    let start = parse_a1_reference(start)?;
    let end = match end {
        Some(end) => parse_a1_reference(end)?,
        None => start,
    };
    Ok((start, end))
}

fn parse_a1_reference(reference: &str) -> Result<(u32, u32), String> {
    let reference = reference.trim().replace('$', "");
    let reference = reference.as_str();
    if reference.is_empty() || reference.contains(':') || reference.contains('!') {
        return Err(format!(
            "only a single-sheet cell reference is supported: {reference}"
        ));
    }
    let split = reference
        .find(|ch: char| ch.is_ascii_digit())
        .ok_or_else(|| format!("invalid cell reference: {reference}"))?;
    let (letters, digits) = reference.split_at(split);
    if letters.is_empty()
        || digits.is_empty()
        || !letters.chars().all(|ch| ch.is_ascii_alphabetic())
        || !digits.chars().all(|ch| ch.is_ascii_digit())
    {
        return Err(format!("invalid cell reference: {reference}"));
    }
    let mut column = 0_u32;
    for letter in letters.bytes() {
        column = column
            .checked_mul(26)
            .and_then(|value| value.checked_add((letter.to_ascii_uppercase() - b'A' + 1) as u32))
            .ok_or_else(|| format!("cell column is too large: {reference}"))?;
    }
    let row = digits
        .parse::<u32>()
        .map_err(|_| format!("cell row is too large: {reference}"))?;
    if row == 0 || row > MAX_WORKSHEET_ROW {
        return Err(format!("cell row is outside the worksheet: {reference}"));
    }
    if column == 0 || column - 1 > MAX_WORKSHEET_COLUMN {
        return Err(format!("cell column is outside the worksheet: {reference}"));
    }
    Ok((column - 1, row))
}

fn positive_index(value: &Value, label: &str) -> Result<u32, String> {
    let number = match value {
        Value::Integer(value) => *value as f64,
        Value::Double(value) => *value,
        _ => return Err(format!("Cells {label} must be numeric")),
    };
    if !number.is_finite() || number.fract() != 0.0 || !(1.0..=u32::MAX as f64).contains(&number) {
        return Err(format!("Cells {label} must be a positive integer"));
    }
    Ok(number as u32)
}

fn range_axis_name(axis: RangeAxis) -> &'static str {
    match axis {
        RangeAxis::Rows => "Rows",
        RangeAxis::Columns => "Columns",
    }
}

fn integer_offset(value: &Value, label: &str) -> Result<i64, String> {
    let number = match value {
        Value::Integer(value) => *value as f64,
        Value::Double(value) => *value,
        _ => return Err(format!("Range.Offset {label} offset must be numeric")),
    };
    if !number.is_finite()
        || number.fract() != 0.0
        || !(i64::MIN as f64..=i64::MAX as f64).contains(&number)
    {
        return Err(format!(
            "Range.Offset {label} offset must be a whole number"
        ));
    }
    Ok(number as i64)
}

fn range_address_from_args(range: CellRange, args: &[Value]) -> Result<String, String> {
    let (row_absolute, column_absolute) = match args {
        [] => (true, true),
        [row_absolute] => (boolean_argument(row_absolute, "row absolute")?, true),
        [row_absolute, column_absolute] => (
            boolean_argument(row_absolute, "row absolute")?,
            boolean_argument(column_absolute, "column absolute")?,
        ),
        _ => return Err("Range.Address supports up to two arguments".to_string()),
    };
    Ok(format_range_address(range, row_absolute, column_absolute))
}

fn boolean_argument(value: &Value, label: &str) -> Result<bool, String> {
    match value {
        Value::Boolean(value) => Ok(*value),
        Value::Integer(value) => Ok(*value != 0),
        Value::Double(value) if value.is_finite() => Ok(*value != 0.0),
        _ => Err(format!("Range.Address {label} must be Boolean")),
    }
}

fn format_range_address(range: CellRange, row_absolute: bool, column_absolute: bool) -> String {
    let format_cell = |row: u32, column: u32| {
        format!(
            "{}{}{}{}",
            if column_absolute { "$" } else { "" },
            oxicells_core::editor::col_to_letter(column),
            if row_absolute { "$" } else { "" },
            row
        )
    };
    let start = format_cell(range.start_row, range.start_column);
    if range.is_single() {
        start
    } else {
        format!("{start}:{}", format_cell(range.end_row, range.end_column))
    }
}

fn from_cell_value(value: &CellValue) -> Value {
    match value {
        CellValue::Empty => Value::Empty,
        CellValue::String(value) => Value::String(value.clone()),
        CellValue::Error(value) => Value::Error(spreadsheet_error_number(value)),
        CellValue::Number(value)
            if value.fract() == 0.0 && *value >= i64::MIN as f64 && *value <= i64::MAX as f64 =>
        {
            Value::Integer(*value as i64)
        }
        CellValue::Number(value) => Value::Double(*value),
        CellValue::Boolean(value) => Value::Boolean(*value),
    }
}

fn to_cell_value(value: Value) -> Result<CellValue, String> {
    match value {
        Value::Empty | Value::Null => Ok(CellValue::Empty),
        Value::Missing => Err("an omitted VBA argument cannot be assigned to a cell".to_string()),
        Value::Nothing => Err("Nothing cannot be assigned to a spreadsheet cell".to_string()),
        Value::Boolean(value) => Ok(CellValue::Boolean(value)),
        Value::Integer(value) => Ok(CellValue::Number(value as f64)),
        Value::Double(value) => Ok(CellValue::Number(value)),
        Value::Error(value) => Ok(CellValue::Error(spreadsheet_error_text(value).to_string())),
        Value::String(value) => Ok(CellValue::String(value)),
        Value::Array(_) => Err("a VBA array cannot be assigned to one cell".to_string()),
        Value::Object(_) => Err("a VBA object cannot be assigned to one cell".to_string()),
    }
}

fn to_formula(value: Value) -> Result<String, String> {
    match value {
        Value::Empty | Value::Null => Ok(String::new()),
        Value::Missing => Err("an omitted VBA argument cannot be used as a formula".to_string()),
        Value::String(value) => Ok(value),
        _ => Err("a spreadsheet formula must be a String".to_string()),
    }
}

#[derive(Deserialize)]
#[serde(untagged)]
enum InputValue {
    Null(()),
    Boolean(bool),
    Number(f64),
    String(String),
}

impl From<InputValue> for Value {
    fn from(value: InputValue) -> Self {
        match value {
            InputValue::Null(()) => Value::Null,
            InputValue::Boolean(value) => Value::Boolean(value),
            InputValue::Number(value)
                if value.fract() == 0.0 && value >= i64::MIN as f64 && value <= i64::MAX as f64 =>
            {
                Value::Integer(value as i64)
            }
            InputValue::Number(value) => Value::Double(value),
            InputValue::String(value) => Value::String(value),
        }
    }
}

#[derive(Serialize)]
#[serde(tag = "type", content = "value", rename_all = "snake_case")]
enum OutputValue {
    Empty,
    Missing,
    Nothing,
    Null,
    Boolean(bool),
    Integer(i64),
    Double(f64),
    Error(i64),
    String(String),
    Array {
        lower_bound: i64,
        dimensions: Vec<OutputArrayDimension>,
        resizable: bool,
        values: Vec<OutputValue>,
    },
    Object {
        handle: u64,
        kind: String,
    },
}

impl From<Value> for OutputValue {
    fn from(value: Value) -> Self {
        match value {
            Value::Empty => Self::Empty,
            Value::Missing => Self::Missing,
            Value::Nothing => Self::Nothing,
            Value::Null => Self::Null,
            Value::Boolean(value) => Self::Boolean(value),
            Value::Integer(value) => Self::Integer(value),
            Value::Double(value) => Self::Double(value),
            Value::Error(value) => Self::Error(value),
            Value::String(value) => Self::String(value),
            Value::Array(value) => Self::Array {
                lower_bound: value.lower_bound(1).unwrap_or(0),
                dimensions: value
                    .dimensions
                    .iter()
                    .map(|dimension| OutputArrayDimension {
                        lower_bound: dimension.lower_bound,
                        length: dimension.length,
                    })
                    .collect(),
                resizable: value.resizable,
                values: value.values.into_iter().map(OutputValue::from).collect(),
            },
            Value::Object(value) => Self::Object {
                handle: value.handle,
                kind: value.kind,
            },
        }
    }
}

fn spreadsheet_error_number(value: &str) -> i64 {
    match value.to_ascii_uppercase().as_str() {
        "#NULL!" => 2000,
        "#DIV/0!" => 2007,
        "#VALUE!" => 2015,
        "#REF!" => 2023,
        "#NAME?" => 2029,
        "#NUM!" => 2036,
        "#N/A" => 2042,
        "#GETTING_DATA" => 2043,
        "#SPILL!" => 2045,
        "#CONNECT!" => 2046,
        "#BLOCKED!" => 2047,
        "#UNKNOWN!" => 2048,
        "#FIELD!" => 2049,
        "#CALC!" => 2050,
        _ => 2015,
    }
}

fn spreadsheet_error_text(value: i64) -> &'static str {
    match value {
        2000 => "#NULL!",
        2007 => "#DIV/0!",
        2015 => "#VALUE!",
        2023 => "#REF!",
        2029 => "#NAME?",
        2036 => "#NUM!",
        2042 => "#N/A",
        2043 => "#GETTING_DATA",
        2045 => "#SPILL!",
        2046 => "#CONNECT!",
        2047 => "#BLOCKED!",
        2048 => "#UNKNOWN!",
        2049 => "#FIELD!",
        2050 => "#CALC!",
        _ => "#VALUE!",
    }
}

#[derive(Serialize)]
struct OutputArrayDimension {
    lower_bound: i64,
    length: usize,
}

#[derive(Serialize)]
struct RunResult {
    workbook: Workbook,
    result: OutputValue,
    debug_output: Vec<String>,
    messages: Vec<BrowserMessage>,
}

#[derive(Debug, PartialEq, Serialize)]
struct BrowserMessage {
    prompt: String,
    title: String,
}

#[derive(Debug, PartialEq, Serialize)]
struct ProcedureSummary {
    name: String,
    kind: &'static str,
    visibility: &'static str,
    parameter_count: usize,
    required_parameter_count: usize,
    line: u32,
}

fn spreadsheet_vba_procedures(source: &str) -> Result<Vec<ProcedureSummary>, String> {
    let module = parse_module(source).map_err(|error| error.to_string())?;
    Ok(module
        .items
        .iter()
        .filter_map(|item| {
            let ModuleItem::Procedure(procedure) = item else {
                return None;
            };
            let kind = match procedure.kind {
                ProcKind::Sub => "sub",
                ProcKind::Function => "function",
                ProcKind::PropertyGet | ProcKind::PropertyLet | ProcKind::PropertySet => {
                    return None
                }
            };
            let visibility = match procedure.visibility {
                Visibility::Default | Visibility::Public | Visibility::Global => "public",
                Visibility::Private => "private",
                Visibility::Friend => "friend",
            };
            Some(ProcedureSummary {
                name: procedure.name.clone(),
                kind,
                visibility,
                parameter_count: procedure.params.len(),
                required_parameter_count: procedure
                    .params
                    .iter()
                    .filter(|parameter| {
                        !parameter.optional && parameter.mode != ParamMode::ParamArray
                    })
                    .count(),
                line: procedure.span.line,
            })
        })
        .collect())
}

/// List executable Sub and Function entry points in VBA source.
#[wasm_bindgen]
pub fn list_spreadsheet_vba_procedures(source: &str) -> Result<JsValue, JsError> {
    let procedures = spreadsheet_vba_procedures(source).map_err(|error| JsError::new(&error))?;
    serde_wasm_bindgen::to_value(&procedures).map_err(|error| JsError::new(&error.to_string()))
}

/// Execute VBA source against an OxiCells workbook IR.
#[wasm_bindgen]
pub fn run_spreadsheet_vba(
    workbook: JsValue,
    source: &str,
    procedure: &str,
    args: JsValue,
    active_sheet: usize,
) -> Result<JsValue, JsError> {
    let mut workbook: Workbook = serde_wasm_bindgen::from_value(workbook)
        .map_err(|error| JsError::new(&format!("invalid workbook: {error}")))?;
    let args: Vec<InputValue> = serde_wasm_bindgen::from_value(args)
        .map_err(|error| JsError::new(&format!("invalid VBA arguments: {error}")))?;
    let module = parse_module(source).map_err(|error| JsError::new(&error.to_string()))?;
    let mut host =
        WorkbookHost::new(&mut workbook, active_sheet).map_err(|error| JsError::new(&error))?;
    let random_seed =
        js_sys::Date::now().to_bits() ^ js_sys::Math::random().to_bits().rotate_left(17);
    let browser_now = js_sys::Date::new_0();
    let local_millis = browser_now.get_time() - browser_now.get_timezone_offset() * 60_000.0;
    let current_time = local_millis / 86_400_000.0 + 25_569.0;
    let result = Runtime::new(&module)
        .with_host(&mut host)
        .with_random_seed(random_seed)
        .with_current_time(current_time)
        .call(procedure, args.into_iter().map(Value::from).collect())
        .map_err(|error| JsError::new(&error.to_string()))?;
    let debug_output = host.take_debug_output();
    let messages = host.take_messages();
    drop(host);
    serde_wasm_bindgen::to_value(&RunResult {
        workbook,
        result: result.into(),
        debug_output,
        messages,
    })
    .map_err(|error| JsError::new(&error.to_string()))
}

#[cfg(test)]
mod tests {
    use super::*;
    use oxicells_core::ir::Sheet;

    fn workbook() -> Workbook {
        Workbook {
            sheets: vec![Sheet {
                name: "Sheet1".to_string(),
                rows: Vec::new(),
                col_count: 0,
                col_widths: Vec::new(),
                default_col_width: 8.43,
                default_row_height: 15.0,
                merge_cells: Vec::new(),
                unsupported_elements: Vec::new(),
            }],
        }
    }

    #[test]
    fn vba_updates_real_workbook_cells_without_losing_value_types() {
        let mut workbook = workbook();
        let module = parse_module(
            "Public Function Fill() As Double\n\
               Range(\"A1\").Value = 40\n\
               Cells(2, 1).Value2 = 2.5\n\
               Fill = Range(\"A1\").Value + Range(\"A2\").Value2\n\
             End Function\n",
        )
        .unwrap();
        let result = {
            let mut host = WorkbookHost::new(&mut workbook, 0).unwrap();
            execute_with_host(&module, "Fill", vec![], &mut host).unwrap()
        };
        assert_eq!(result, Value::Double(42.5));
        assert!(matches!(
            workbook.sheets[0].rows[0].cells[0].value,
            CellValue::Number(value) if value == 40.0
        ));
        assert!(matches!(
            workbook.sheets[0].rows[1].cells[0].value,
            CellValue::Number(value) if value == 2.5
        ));
    }

    #[test]
    fn vba_collects_debug_print_output() {
        let mut workbook = workbook();
        let module = parse_module(
            "Public Sub TraceValues()\n\
               Debug.Print \"before\", 42, True\n\
               Debug.Print\n\
               Range(\"A1\").Value = 7\n\
             End Sub\n",
        )
        .unwrap();
        let output = {
            let mut host = WorkbookHost::new(&mut workbook, 0).unwrap();
            execute_with_host(&module, "TraceValues", vec![], &mut host).unwrap();
            host.take_debug_output()
        };

        assert_eq!(output, vec!["before\t42\tTrue", ""]);
        assert!(matches!(
            workbook.sheets[0].rows[0].cells[0].value,
            CellValue::Number(7.0)
        ));
    }

    #[test]
    fn vba_collects_ok_only_message_boxes() {
        let mut workbook = workbook();
        let module = parse_module(
            "Public Function Notify() As Long\n\
               Notify = MsgBox(\"Finished\", vbOKOnly + vbInformation, \"Report\")\n\
             End Function\n",
        )
        .unwrap();
        let (result, messages) = {
            let mut host = WorkbookHost::new(&mut workbook, 0).unwrap();
            let result = execute_with_host(&module, "Notify", vec![], &mut host).unwrap();
            (result, host.take_messages())
        };

        assert_eq!(result, Value::Integer(1));
        assert_eq!(
            messages,
            vec![BrowserMessage {
                prompt: "Finished".to_string(),
                title: "Report".to_string(),
            }]
        );
    }

    #[test]
    fn vba_rejects_interactive_message_boxes() {
        let mut workbook = workbook();
        let module = parse_module(
            "Public Sub AskUser()\n\
               MsgBox \"Continue?\", vbYesNo\n\
             End Sub\n",
        )
        .unwrap();
        let failure = {
            let mut host = WorkbookHost::new(&mut workbook, 0).unwrap();
            execute_with_host(&module, "AskUser", vec![], &mut host).unwrap_err()
        };

        assert!(failure.message.contains("interactive button styles"));
    }

    #[test]
    fn lists_vba_entry_points_and_required_arguments() {
        let procedures = spreadsheet_vba_procedures(
            "Private Sub Prepare(Optional label As String = \"x\")\n\
             End Sub\n\
             Public Function Add(left As Long, ByVal right As Long) As Long\n\
             End Function\n\
             Public Property Get Caption() As String\n\
             End Property\n",
        )
        .unwrap();

        assert_eq!(
            procedures,
            vec![
                ProcedureSummary {
                    name: "Prepare".to_string(),
                    kind: "sub",
                    visibility: "private",
                    parameter_count: 1,
                    required_parameter_count: 0,
                    line: 1,
                },
                ProcedureSummary {
                    name: "Add".to_string(),
                    kind: "function",
                    visibility: "public",
                    parameter_count: 2,
                    required_parameter_count: 2,
                    line: 3,
                },
            ]
        );
    }

    #[test]
    fn validates_cell_references_and_one_based_cells_indices() {
        assert_eq!(parse_a1_reference("AA12").unwrap(), (26, 12));
        assert_eq!(parse_range_reference("B2:A1").unwrap(), ((1, 2), (0, 1)));
        assert!(parse_a1_reference("A0").is_err());
        assert!(parse_a1_reference("A1048577").is_err());
        assert!(parse_a1_reference("XFE1").is_err());
        assert!(parse_a1_reference("A1:B2").is_err());
        assert!(parse_range_reference("A1:B2:C3").is_err());
        assert!(positive_index(&Value::Integer(0), "row").is_err());
        assert!(positive_index(&Value::Double(1.5), "row").is_err());
    }

    #[test]
    fn preserves_spreadsheet_error_values_as_vba_error_variants() {
        assert_eq!(
            from_cell_value(&CellValue::Error("#N/A".to_string())),
            Value::Error(2042)
        );
        assert!(matches!(
            to_cell_value(Value::Error(2007)).unwrap(),
            CellValue::Error(value) if value == "#DIV/0!"
        ));
    }

    #[test]
    fn vba_reads_and_writes_rectangular_ranges_in_row_major_order() {
        let mut workbook = workbook();
        let module = parse_module(
            "Public Function FillRange() As Long\n\
               Dim values As Variant\n\
               Dim item As Variant\n\
               Dim total As Long\n\
               Range(\"A1:B2\").Value = 5\n\
               values = Range(\"A1\", \"B2\").Value\n\
               For Each item In values\n\
                 total = total + item\n\
               Next item\n\
               FillRange = total + values(1, 1) + values(2, 2) + UBound(values, 2)\n\
             End Function\n\
             Public Sub WriteArray()\n\
               Dim values(1 To 2, 1 To 2) As Long\n\
               values(1, 1) = 10\n\
               values(1, 2) = 20\n\
               values(2, 1) = 30\n\
               values(2, 2) = 40\n\
               Range(\"C1:D2\").Value = values\n\
             End Sub\n",
        )
        .unwrap();
        {
            let mut host = WorkbookHost::new(&mut workbook, 0).unwrap();
            let mut runtime = oxivba_core::Runtime::new(&module).with_host(&mut host);
            assert_eq!(
                runtime.call("FillRange", vec![]).unwrap(),
                Value::Integer(32)
            );
            assert_eq!(runtime.call("WriteArray", vec![]).unwrap(), Value::Empty);
        }
        let rows = &workbook.sheets[0].rows;
        assert!(matches!(rows[0].cells[2].value, CellValue::Number(10.0)));
        assert!(matches!(rows[0].cells[3].value, CellValue::Number(20.0)));
        assert!(matches!(rows[1].cells[2].value, CellValue::Number(30.0)));
        assert!(matches!(rows[1].cells[3].value, CellValue::Number(40.0)));
    }

    #[test]
    fn vba_accesses_worksheets_by_name_and_one_based_index() {
        let mut workbook = workbook();
        workbook.sheets.push(Sheet {
            name: "Data".to_string(),
            rows: Vec::new(),
            col_count: 0,
            col_widths: Vec::new(),
            default_col_width: 8.43,
            default_row_height: 15.0,
            merge_cells: Vec::new(),
            unsupported_elements: Vec::new(),
        });
        let module = parse_module(
            "Public Function FillSheets() As Long\n\
               Dim ws As Worksheet\n\
               Set ws = Worksheets(\"Data\")\n\
               ws.Range(\"A1\").Value = 40\n\
               Worksheets(2).Cells(2, 1).Value = 2\n\
               FillSheets = ws.Range(\"A1\").Value + ws.Cells(2, 1).Value\n\
             End Function\n\
             Public Function SheetIdentity() As String\n\
               Dim ws As Worksheet\n\
               Set ws = Sheets(2)\n\
               SheetIdentity = ws.Name & \"|\" & ws.Index\n\
             End Function\n",
        )
        .unwrap();
        let (total, identity) = {
            let mut host = WorkbookHost::new(&mut workbook, 0).unwrap();
            let mut runtime = oxivba_core::Runtime::new(&module).with_host(&mut host);
            let total = runtime.call("FillSheets", vec![]).unwrap();
            let identity = runtime.call("SheetIdentity", vec![]).unwrap();
            (total, identity)
        };
        assert_eq!(total, Value::Integer(42));
        assert_eq!(identity, Value::String("Data|2".to_string()));
        assert!(matches!(
            workbook.sheets[1].rows[0].cells[0].value,
            CellValue::Number(value) if value == 40.0
        ));
        assert!(matches!(
            workbook.sheets[1].rows[1].cells[0].value,
            CellValue::Number(value) if value == 2.0
        ));
    }

    #[test]
    fn vba_evaluates_bracket_and_explicit_a1_references() {
        let mut workbook = workbook();
        workbook.sheets.push(Sheet {
            name: "Data Sheet".to_string(),
            rows: Vec::new(),
            col_count: 0,
            col_widths: Vec::new(),
            default_col_width: 8.43,
            default_row_height: 15.0,
            merge_cells: Vec::new(),
            unsupported_elements: Vec::new(),
        });
        let module = parse_module(
            "Public Function EvaluateReferences() As String\n\
               Dim target As Range\n\
               Dim item As Variant\n\
               Dim total As Long\n\
               [A1:B2] = 5\n\
               Evaluate(\"$C$1\").Value = 20\n\
               ['Data Sheet'!A1] = 30\n\
               Set target = ['Data Sheet'!$B$2]\n\
               target.Value = 40\n\
               For Each item In [A1:B2]\n\
                 total = total + item\n\
               Next item\n\
               EvaluateReferences = total & \"|\" & [C1] & \"|\" & Worksheets(\"Data Sheet\").Range(\"A1\").Value & \"|\" & target.Value\n\
             End Function\n",
        )
        .unwrap();

        let result = {
            let mut host = WorkbookHost::new(&mut workbook, 0).unwrap();
            execute_with_host(&module, "EvaluateReferences", vec![], &mut host).unwrap()
        };

        assert_eq!(result, Value::String("20|20|30|40".to_string()));
    }

    #[test]
    fn vba_offsets_resizes_and_clears_ranges() {
        let mut workbook = workbook();
        let module = parse_module(
            "Public Sub TransformRange()\n\
               Range(\"B2:C3\").Value = 7\n\
               Range(\"B2:C3\").Offset(1, -1).Resize(1, 2).Value = 9\n\
               Range(\"B2:C2\").ClearContents\n\
             End Sub\n",
        )
        .unwrap();
        {
            let mut host = WorkbookHost::new(&mut workbook, 0).unwrap();
            execute_with_host(&module, "TransformRange", vec![], &mut host).unwrap();
        }

        assert!(matches!(
            workbook.sheets[0].rows[0].cells[0].value,
            CellValue::Empty
        ));
        assert!(matches!(
            workbook.sheets[0].rows[0].cells[1].value,
            CellValue::Empty
        ));
        assert!(matches!(
            workbook.sheets[0].rows[1].cells[0].value,
            CellValue::Number(9.0)
        ));
        assert!(matches!(
            workbook.sheets[0].rows[1].cells[1].value,
            CellValue::Number(9.0)
        ));
        assert!(matches!(
            workbook.sheets[0].rows[1].cells[2].value,
            CellValue::Number(7.0)
        ));
    }

    #[test]
    fn vba_copies_range_values_and_styles_to_a_destination() {
        let mut workbook = workbook();
        workbook.sheets[0].rows.push(Row {
            index: 1,
            height: None,
            cells: vec![
                Cell {
                    col: 0,
                    value: CellValue::Number(10.0),
                    style: CellStyle {
                        bold: true,
                        ..CellStyle::default()
                    },
                    formula: None,
                },
                Cell {
                    col: 1,
                    value: CellValue::String("copied".to_string()),
                    style: CellStyle {
                        bg_color: Some("#ff0000".to_string()),
                        ..CellStyle::default()
                    },
                    formula: None,
                },
            ],
        });
        let module = parse_module(
            "Public Sub CopyValues()\n\
               Range(\"A1:B1\").Copy Destination:=Range(\"C2\")\n\
             End Sub\n",
        )
        .unwrap();
        {
            let mut host = WorkbookHost::new(&mut workbook, 0).unwrap();
            execute_with_host(&module, "CopyValues", vec![], &mut host).unwrap();
        }

        let destination = &workbook.sheets[0].rows[1].cells;
        assert!(matches!(destination[0].value, CellValue::Number(10.0)));
        assert!(destination[0].style.bold);
        assert!(matches!(
            &destination[1].value,
            CellValue::String(value) if value == "copied"
        ));
        assert_eq!(destination[1].style.bg_color.as_deref(), Some("#ff0000"));
    }

    #[test]
    fn vba_copies_formulas_with_relative_reference_adjustment() {
        let mut workbook = workbook();
        let module = parse_module(
            "Public Sub CopyFormula()\n\
               Range(\"A1\").Formula = \"=B1+$C$1+D$2\"\n\
               Range(\"B1\").Formula = \"=SUM(A1:B1)\"\n\
               Range(\"A1:B1\").Copy Destination:=Range(\"C2\")\n\
             End Sub\n",
        )
        .unwrap();
        {
            let mut host = WorkbookHost::new(&mut workbook, 0).unwrap();
            execute_with_host(&module, "CopyFormula", vec![], &mut host).unwrap();
        }

        let row = &workbook.sheets[0].rows[1];
        assert_eq!(row.cells[0].formula.as_deref(), Some("D2+$C$1+F$2"));
        assert_eq!(row.cells[1].formula.as_deref(), Some("SUM(C2:D2)"));
        assert!(matches!(row.cells[0].value, CellValue::Empty));
        assert!(matches!(row.cells[1].value, CellValue::Empty));
    }

    #[test]
    fn vba_formats_ranges_without_destroying_cell_contents() {
        let mut workbook = workbook();
        let module = parse_module(
            "Public Sub FormatCells()\n\
               Range(\"A1\").Value = 12\n\
               Range(\"B1\").Formula = \"=A1*2\"\n\
               Range(\"A1:B2\").Font.Bold = True\n\
               Range(\"A1:B2\").Font.Italic = True\n\
               Range(\"A1:B2\").Font.Size = 14\n\
               Range(\"A1:B2\").Font.Color = RGB(10, 20, 30)\n\
               Range(\"A1:B2\").Interior.Color = vbYellow\n\
               Range(\"A1:B2\").NumberFormat = \"0.00\"\n\
               Debug.Print Range(\"A1:B2\").Font.Bold, Range(\"A1:B2\").Font.Size, Range(\"A1:B2\").Font.Color, Range(\"A1:B2\").Interior.Color, Range(\"A1:B2\").NumberFormat\n\
             End Sub\n",
        )
        .unwrap();
        let debug_output = {
            let mut host = WorkbookHost::new(&mut workbook, 0).unwrap();
            execute_with_host(&module, "FormatCells", vec![], &mut host).unwrap();
            host.take_debug_output()
        };

        assert!(matches!(
            workbook.sheets[0].rows[0].cells[0].value,
            CellValue::Number(12.0)
        ));
        assert_eq!(
            workbook.sheets[0].rows[0].cells[1].formula.as_deref(),
            Some("A1*2")
        );
        assert_eq!(workbook.sheets[0].rows.len(), 2);
        for row in &workbook.sheets[0].rows {
            assert_eq!(row.cells.len(), 2);
            for cell in &row.cells {
                assert!(cell.style.bold);
                assert!(cell.style.italic);
                assert_eq!(cell.style.font_size, Some(14.0));
                assert_eq!(cell.style.font_color.as_deref(), Some("#0a141e"));
                assert_eq!(cell.style.bg_color.as_deref(), Some("#ffff00"));
                assert_eq!(cell.style.number_format.as_deref(), Some("0.00"));
            }
        }
        assert_eq!(
            debug_output,
            vec!["True\t14\t1971210\t65535\t0.00".to_string()]
        );
    }

    #[test]
    fn vba_aligns_and_clears_contents_or_formats_independently() {
        let mut workbook = workbook();
        let module = parse_module(
            "Public Sub ClearCells()\n\
               Range(\"A1\").Value = 1\n\
               Range(\"B1\").Formula = \"=A1*2\"\n\
               Range(\"C1\").Value = 3\n\
               Range(\"A1:C1\").Font.Bold = True\n\
               Range(\"A1:C1\").NumberFormat = \"0.00\"\n\
               Range(\"A1:C1\").HorizontalAlignment = xlCenter\n\
               Range(\"A1\").ClearFormats\n\
               Range(\"B1\").ClearContents\n\
               Range(\"C1\").Clear\n\
               Debug.Print Range(\"A1\").Value, Range(\"A1\").HorizontalAlignment, Range(\"B1\").Font.Bold, Range(\"B1\").HorizontalAlignment, Range(\"B1\").NumberFormat\n\
             End Sub\n",
        )
        .unwrap();
        let debug_output = {
            let mut host = WorkbookHost::new(&mut workbook, 0).unwrap();
            execute_with_host(&module, "ClearCells", vec![], &mut host).unwrap();
            host.take_debug_output()
        };

        let cells = &workbook.sheets[0].rows[0].cells;
        assert!(matches!(cells[0].value, CellValue::Number(1.0)));
        assert!(!cells[0].style.bold);
        assert_eq!(cells[0].style.number_format, None);
        assert_eq!(cells[0].style.horizontal_align, None);

        assert!(matches!(cells[1].value, CellValue::Empty));
        assert_eq!(cells[1].formula, None);
        assert!(cells[1].style.bold);
        assert_eq!(cells[1].style.number_format.as_deref(), Some("0.00"));
        assert_eq!(cells[1].style.horizontal_align.as_deref(), Some("center"));

        assert!(matches!(cells[2].value, CellValue::Empty));
        assert_eq!(cells[2].formula, None);
        assert!(!cells[2].style.bold);
        assert_eq!(cells[2].style.number_format, None);
        assert_eq!(cells[2].style.horizontal_align, None);
        assert_eq!(debug_output, vec!["1\t1\tTrue\t-4108\t0.00".to_string()]);
    }

    #[test]
    fn vba_sets_all_or_indexed_range_borders() {
        let mut workbook = workbook();
        let module = parse_module(
            "Public Sub DrawBorders()\n\
               Range(\"A1\").Value = 5\n\
               Range(\"B1\").Formula = \"=A1*2\"\n\
               Range(\"A1:B2\").Borders.LineStyle = xlContinuous\n\
               Range(\"A1:B2\").Borders(xlEdgeBottom).LineStyle = xlLineStyleNone\n\
               Debug.Print Range(\"A1:B2\").Borders.LineStyle, Range(\"A1:B2\").Borders(xlEdgeBottom).LineStyle, Range(\"A1:B2\").Borders(xlEdgeTop).LineStyle\n\
             End Sub\n",
        )
        .unwrap();
        let debug_output = {
            let mut host = WorkbookHost::new(&mut workbook, 0).unwrap();
            execute_with_host(&module, "DrawBorders", vec![], &mut host).unwrap();
            host.take_debug_output()
        };

        assert!(matches!(
            workbook.sheets[0].rows[0].cells[0].value,
            CellValue::Number(5.0)
        ));
        assert_eq!(
            workbook.sheets[0].rows[0].cells[1].formula.as_deref(),
            Some("A1*2")
        );
        for (row_index, row) in workbook.sheets[0].rows.iter().enumerate() {
            for cell in &row.cells {
                assert!(cell.style.border_top);
                assert_eq!(cell.style.border_bottom, row_index == 0);
                assert!(cell.style.border_left);
                assert!(cell.style.border_right);
            }
        }
        assert_eq!(debug_output, vec!["Null\t-4142\t1".to_string()]);
    }

    #[test]
    fn vba_reads_and_writes_column_widths_and_row_heights() {
        let mut workbook = workbook();
        let module = parse_module(
            "Public Sub ResizeSheet()\n\
               Range(\"A1\").Value = 7\n\
               Columns(2).ColumnWidth = 12.5\n\
               Range(\"C1:D1\").ColumnWidth = 9\n\
               Columns(4).ColumnWidth = 11\n\
               Rows(3).RowHeight = 24\n\
               Range(\"A4\").RowHeight = 18\n\
               Debug.Print Columns(2).ColumnWidth, Range(\"C1:D1\").ColumnWidth, Rows(3).RowHeight, Rows(4).RowHeight\n\
             End Sub\n",
        )
        .unwrap();
        let debug_output = {
            let mut host = WorkbookHost::new(&mut workbook, 0).unwrap();
            execute_with_host(&module, "ResizeSheet", vec![], &mut host).unwrap();
            host.take_debug_output()
        };

        let sheet = &workbook.sheets[0];
        assert_eq!(sheet.col_count, 4);
        assert_eq!(
            sheet.col_widths,
            vec![sheet.default_col_width, 12.5, 9.0, 11.0]
        );
        assert!(matches!(
            sheet.rows[0].cells[0].value,
            CellValue::Number(7.0)
        ));
        assert_eq!(sheet.rows[1].index, 3);
        assert_eq!(sheet.rows[1].height, Some(24.0));
        assert!(sheet.rows[1].cells.is_empty());
        assert_eq!(sheet.rows[2].index, 4);
        assert_eq!(sheet.rows[2].height, Some(18.0));
        assert!(sheet.rows[2].cells.is_empty());
        assert_eq!(debug_output, vec!["12.5\tNull\t24\t18".to_string()]);
    }

    #[test]
    fn vba_inspects_formulas_and_expands_ranges_to_whole_axes() {
        let mut workbook = workbook();
        let module = parse_module(
            "Public Sub InspectRanges()\n\
               Range(\"A1\").Formula = \"=1+1\"\n\
               Range(\"A2\").Formula = \"=A1+1\"\n\
               Range(\"B1\").Value = 3\n\
               Range(\"B3\").EntireRow.RowHeight = 22\n\
               Range(\"C1\").EntireColumn.ColumnWidth = 13\n\
               Debug.Print Range(\"A1:A2\").HasFormula, Range(\"B1:B2\").HasFormula, Range(\"A1:B1\").HasFormula\n\
               Debug.Print Range(\"B3\").EntireRow.Address, Range(\"C1\").EntireColumn.Address, Range(\"C1\").EntireColumn.CountLarge\n\
               Debug.Print Range(\"A1\").Parent.Name, Range(\"A1\").Worksheet.Index\n\
             End Sub\n",
        )
        .unwrap();
        let debug_output = {
            let mut host = WorkbookHost::new(&mut workbook, 0).unwrap();
            execute_with_host(&module, "InspectRanges", vec![], &mut host).unwrap();
            host.take_debug_output()
        };

        let sheet = &workbook.sheets[0];
        assert_eq!(sheet.col_count, 3);
        assert_eq!(sheet.col_widths.len(), 3);
        assert_eq!(sheet.col_widths[2], 13.0);
        let row = sheet.rows.iter().find(|row| row.index == 3).unwrap();
        assert_eq!(row.height, Some(22.0));
        assert!(row.cells.is_empty());
        assert_eq!(
            debug_output,
            vec![
                "True\tFalse\tNull".to_string(),
                "$A$3:$XFD$3\t$C$1:$C$1048576\t1048576".to_string(),
                "Sheet1\t1".to_string(),
            ]
        );
    }

    #[test]
    fn vba_merges_and_unmerges_ranges_without_keeping_hidden_contents() {
        let mut workbook = workbook();
        let module = parse_module(
            "Public Sub MergeRanges()\n\
               Range(\"A1\").Value = \"heading\"\n\
               Range(\"B1\").Value = \"discarded\"\n\
               Range(\"A2\").Value = 3\n\
               Range(\"B2\").Formula = \"=A2*2\"\n\
               Range(\"A1:B2\").Merge\n\
               Debug.Print Range(\"A1\").MergeCells, Range(\"A1:B2\").MergeCells, Range(\"A1:C2\").MergeCells, Range(\"C1\").MergeCells\n\
               Range(\"A1:B2\").UnMerge\n\
               Range(\"C1:D1\").MergeCells = True\n\
               Range(\"E1:F1\").MergeCells = True\n\
               Range(\"E1\").MergeCells = False\n\
             End Sub\n",
        )
        .unwrap();
        let debug_output = {
            let mut host = WorkbookHost::new(&mut workbook, 0).unwrap();
            execute_with_host(&module, "MergeRanges", vec![], &mut host).unwrap();
            host.take_debug_output()
        };

        let sheet = &workbook.sheets[0];
        assert!(matches!(
            &sheet.rows[0].cells[0].value,
            CellValue::String(value) if value == "heading"
        ));
        assert!(matches!(sheet.rows[0].cells[1].value, CellValue::Empty));
        assert!(matches!(sheet.rows[1].cells[0].value, CellValue::Empty));
        assert!(matches!(sheet.rows[1].cells[1].value, CellValue::Empty));
        assert_eq!(sheet.rows[1].cells[1].formula, None);
        assert_eq!(sheet.merge_cells.len(), 1);
        let merge = &sheet.merge_cells[0];
        assert_eq!(
            (
                merge.start_row,
                merge.start_col,
                merge.end_row,
                merge.end_col
            ),
            (1, 2, 1, 3)
        );
        assert_eq!(debug_output, vec!["True\tTrue\tNull\tFalse".to_string()]);

        let overlap = parse_module(
            "Public Sub OverlapMerge()\n\
               Range(\"D1:E1\").Merge\n\
             End Sub\n",
        )
        .unwrap();
        let failure = {
            let mut host = WorkbookHost::new(&mut workbook, 0).unwrap();
            execute_with_host(&overlap, "OverlapMerge", vec![], &mut host).unwrap_err()
        };
        assert!(failure
            .message
            .contains("overlaps an existing merged range"));
        assert_eq!(workbook.sheets[0].merge_cells.len(), 1);
    }

    #[test]
    fn vba_selects_ranges_and_exposes_the_active_cell() {
        let mut workbook = workbook();
        let module = parse_module(
            "Public Function FillSelection() As String\n\
               Range(\"B2:C3\").Select\n\
               Selection.Value = 5\n\
               ActiveCell.Value = 7\n\
               FillSelection = Selection.Address & \"|\" & Application.Selection.Count & \"|\" & Application.ActiveCell.Address\n\
             End Function\n",
        )
        .unwrap();
        let result = {
            let mut host = WorkbookHost::new(&mut workbook, 0).unwrap();
            execute_with_host(&module, "FillSelection", vec![], &mut host).unwrap()
        };

        assert_eq!(result, Value::String("$B$2:$C$3|4|$B$2".to_string()));
        let rows = &workbook.sheets[0].rows;
        assert!(matches!(rows[0].cells[0].value, CellValue::Number(7.0)));
        assert!(matches!(rows[0].cells[1].value, CellValue::Number(5.0)));
        assert!(matches!(rows[1].cells[0].value, CellValue::Number(5.0)));
        assert!(matches!(rows[1].cells[1].value, CellValue::Number(5.0)));
    }

    #[test]
    fn vba_requires_a_worksheet_to_be_active_before_selecting_its_range() {
        let mut workbook = workbook();
        workbook.sheets.push(Sheet {
            name: "Data".to_string(),
            rows: Vec::new(),
            col_count: 0,
            col_widths: Vec::new(),
            default_col_width: 8.43,
            default_row_height: 15.0,
            merge_cells: Vec::new(),
            unsupported_elements: Vec::new(),
        });
        let invalid = parse_module(
            "Public Sub InvalidSelection()\n\
               Worksheets(\"Data\").Range(\"A1\").Select\n\
             End Sub\n",
        )
        .unwrap();
        let failure = {
            let mut host = WorkbookHost::new(&mut workbook, 0).unwrap();
            execute_with_host(&invalid, "InvalidSelection", vec![], &mut host).unwrap_err()
        };
        assert!(failure.message.contains("worksheet to be active"));

        let valid = parse_module(
            "Public Sub ValidSelection()\n\
               Worksheets(\"Data\").Activate\n\
               Range(\"A1\").Select\n\
               Selection.Value = 42\n\
             End Sub\n",
        )
        .unwrap();
        {
            let mut host = WorkbookHost::new(&mut workbook, 0).unwrap();
            execute_with_host(&valid, "ValidSelection", vec![], &mut host).unwrap();
        }
        assert!(matches!(
            workbook.sheets[1].rows[0].cells[0].value,
            CellValue::Number(42.0)
        ));
    }

    #[test]
    fn vba_uses_nested_with_blocks_for_range_chains() {
        let mut workbook = workbook();
        let module = parse_module(
            "Public Sub TransformWith()\n\
               With Range(\"B2:C3\")\n\
                 .Value = 7\n\
                 With .Offset(1, -1).Resize(1, 2)\n\
                   .Value = 9\n\
                 End With\n\
                 .Resize(1, 2).ClearContents\n\
               End With\n\
               With Range(\"D1\")\n\
                 .Value = 5\n\
                 .ClearContents\n\
               End With\n\
             End Sub\n",
        )
        .unwrap();
        {
            let mut host = WorkbookHost::new(&mut workbook, 0).unwrap();
            execute_with_host(&module, "TransformWith", vec![], &mut host).unwrap();
        }

        let cell = |row, column| {
            workbook.sheets[0]
                .rows
                .iter()
                .find(|item| item.index == row)
                .and_then(|item| item.cells.iter().find(|cell| cell.col == column))
                .map(|cell| &cell.value)
        };
        assert!(matches!(cell(1, 3), Some(CellValue::Empty)));
        assert!(matches!(cell(2, 1), Some(CellValue::Empty)));
        assert!(matches!(cell(2, 2), Some(CellValue::Empty)));
        assert!(matches!(cell(3, 0), Some(CellValue::Number(9.0))));
        assert!(matches!(cell(3, 1), Some(CellValue::Number(9.0))));
        assert!(matches!(cell(3, 2), Some(CellValue::Number(7.0))));
    }

    #[test]
    fn vba_recovers_from_spreadsheet_host_errors() {
        let mut workbook = workbook();
        let module = parse_module(
            "Public Function RecoverRange() As String\n\
               Dim failure As Long\n\
               On Error Resume Next\n\
               Range(\"not-an-address\").Value = 1\n\
               failure = Err.Number\n\
               On Error GoTo 0\n\
               Range(\"A1\").Value = 42\n\
               RecoverRange = failure & \"|\" & Range(\"A1\").Value\n\
             End Function\n",
        )
        .unwrap();
        let result = {
            let mut host = WorkbookHost::new(&mut workbook, 0).unwrap();
            execute_with_host(&module, "RecoverRange", vec![], &mut host).unwrap()
        };

        assert_eq!(result, Value::String("1004|42".to_string()));
        assert!(matches!(
            workbook.sheets[0].rows[0].cells[0].value,
            CellValue::Number(42.0)
        ));
    }

    #[test]
    fn vba_preserves_module_state_between_entry_point_calls() {
        let mut workbook = workbook();
        let module = parse_module(
            "Private runCount As Long\n\
             Private target As Worksheet\n\
             Public Sub InitializeState()\n\
               Set target = Worksheets(1)\n\
               runCount = 40\n\
             End Sub\n\
             Public Sub ApplyState()\n\
               runCount = runCount + 1\n\
               target.Range(\"A1\").Value = runCount\n\
             End Sub\n\
             Public Function ReadState() As Long\n\
               ReadState = runCount + target.Range(\"A1\").Value\n\
             End Function\n",
        )
        .unwrap();
        let result = {
            let mut host = WorkbookHost::new(&mut workbook, 0).unwrap();
            let mut runtime = oxivba_core::Runtime::new(&module).with_host(&mut host);
            runtime.call("InitializeState", vec![]).unwrap();
            runtime.call("ApplyState", vec![]).unwrap();
            runtime.call("ApplyState", vec![]).unwrap();
            runtime.call("ReadState", vec![]).unwrap()
        };

        assert_eq!(result, Value::Integer(84));
        assert!(matches!(
            workbook.sheets[0].rows[0].cells[0].value,
            CellValue::Number(42.0)
        ));
    }

    #[test]
    fn vba_tracks_nothing_and_range_object_identity() {
        let mut workbook = workbook();
        let module = parse_module(
            "Public Function ObjectState() As String\n\
               Dim target As Range\n\
               Dim aliasValue As Object\n\
               ObjectState = (target Is Nothing) & \"|\"\n\
               Set target = Range(\"A1\")\n\
               Set aliasValue = target\n\
               target.Value = 42\n\
               ObjectState = ObjectState & (target Is aliasValue) & \"|\" & (TypeOf target Is Range) & \"|\" & TypeName(target) & \"|\"\n\
               Set aliasValue = Nothing\n\
               ObjectState = ObjectState & (aliasValue Is Nothing) & \"|\" & IsObject(aliasValue) & \"|\"\n\
               On Error Resume Next\n\
               ObjectState = aliasValue.Value\n\
               ObjectState = ObjectState & Err.Number\n\
             End Function\n",
        )
        .unwrap();
        let result = {
            let mut host = WorkbookHost::new(&mut workbook, 0).unwrap();
            execute_with_host(&module, "ObjectState", vec![], &mut host).unwrap()
        };

        assert_eq!(
            result,
            Value::String("True|True|True|Range|True|True|91".to_string())
        );
        assert!(matches!(
            workbook.sheets[0].rows[0].cells[0].value,
            CellValue::Number(42.0)
        ));
    }

    #[test]
    fn vba_collections_store_key_and_enumerate_range_objects() {
        let mut workbook = workbook();
        let module = parse_module(
            "Public Function CollectionRanges() As String\n\
               Dim ranges As New Collection\n\
               Dim cell As Variant\n\
               Dim value As Long\n\
               ranges.Add Range(\"A1\"), \"first\"\n\
               ranges.Add Range(\"A2\"), \"second\"\n\
               For Each cell In ranges\n\
                 value = value + 10\n\
                 cell.Value = value\n\
               Next\n\
               CollectionRanges = ranges.Count & \"|\" & ranges(\"first\").Value & \"|\" & ranges.Item(2).Value\n\
             End Function\n",
        )
        .unwrap();
        let result = {
            let mut host = WorkbookHost::new(&mut workbook, 0).unwrap();
            execute_with_host(&module, "CollectionRanges", vec![], &mut host).unwrap()
        };

        assert_eq!(result, Value::String("2|10|20".to_string()));
        assert!(matches!(
            workbook.sheets[0].rows[0].cells[0].value,
            CellValue::Number(10.0)
        ));
        assert!(matches!(
            workbook.sheets[0].rows[1].cells[0].value,
            CellValue::Number(20.0)
        ));
    }

    #[test]
    fn vba_inspects_ranges_and_uses_relative_cells() {
        let mut workbook = workbook();
        let module = parse_module(
            "Public Function InspectRange() As String\n\
               Range(\"B2:C3\").Cells(2, 2).Value = 11\n\
               InspectRange = Range(\"B2:C3\").Row & \"|\" & Range(\"B2:C3\").Column & \"|\" & Range(\"B2:C3\").Count & \"|\" & Range(\"B2:C3\").Address & \"|\" & Range(\"B2:C3\").Address(False, False)\n\
             End Function\n",
        )
        .unwrap();
        let result = {
            let mut host = WorkbookHost::new(&mut workbook, 0).unwrap();
            execute_with_host(&module, "InspectRange", vec![], &mut host).unwrap()
        };

        assert_eq!(result, Value::String("2|2|4|$B$2:$C$3|B2:C3".to_string()));
        assert!(matches!(
            workbook.sheets[0].rows[0].cells[0].value,
            CellValue::Number(11.0)
        ));
        assert_eq!(workbook.sheets[0].rows[0].index, 3);
        assert_eq!(workbook.sheets[0].rows[0].cells[0].col, 2);
    }

    #[test]
    fn vba_uses_active_sheet_and_workbook_context() {
        let mut workbook = workbook();
        workbook.sheets.push(Sheet {
            name: "Data".to_string(),
            rows: Vec::new(),
            col_count: 0,
            col_widths: Vec::new(),
            default_col_width: 8.43,
            default_row_height: 15.0,
            merge_cells: Vec::new(),
            unsupported_elements: Vec::new(),
        });
        let module = parse_module(
            "Public Function UseContext() As String\n\
               ThisWorkbook.Worksheets(\"Data\").Activate\n\
               Range(\"A1\").Value = 40\n\
               Application.ActiveSheet.Cells(2, 1).Value = 2\n\
               UseContext = ActiveSheet.Name & \"|\" & ActiveWorkbook.Worksheets(2).Range(\"A1\").Value + ActiveWorkbook.Sheets(2).Cells(2, 1).Value\n\
             End Function\n",
        )
        .unwrap();
        let result = {
            let mut host = WorkbookHost::new(&mut workbook, 0).unwrap();
            execute_with_host(&module, "UseContext", vec![], &mut host).unwrap()
        };

        assert_eq!(result, Value::String("Data|42".to_string()));
        assert!(workbook.sheets[0].rows.is_empty());
        assert!(matches!(
            workbook.sheets[1].rows[0].cells[0].value,
            CellValue::Number(40.0)
        ));
        assert!(matches!(
            workbook.sheets[1].rows[1].cells[0].value,
            CellValue::Number(2.0)
        ));
    }

    #[test]
    fn vba_discovers_the_used_range_on_worksheets() {
        let mut workbook = workbook();
        workbook.sheets.push(Sheet {
            name: "Empty".to_string(),
            rows: Vec::new(),
            col_count: 0,
            col_widths: Vec::new(),
            default_col_width: 8.43,
            default_row_height: 15.0,
            merge_cells: Vec::new(),
            unsupported_elements: Vec::new(),
        });
        let module = parse_module(
            "Public Function InspectUsedRange() As String\n\
               Range(\"D4\").Value = 1\n\
               Range(\"F7\").Formula = \"=1+1\"\n\
               InspectUsedRange = Worksheets(1).UsedRange.Address(False, False) & \"|\" & ActiveSheet.UsedRange.Rows.Count & \"|\" & UsedRange.Columns.Count & \"|\" & Worksheets(2).UsedRange.Address(False, False)\n\
             End Function\n",
        )
        .unwrap();
        let result = {
            let mut host = WorkbookHost::new(&mut workbook, 0).unwrap();
            execute_with_host(&module, "InspectUsedRange", vec![], &mut host).unwrap()
        };

        assert_eq!(result, Value::String("D4:F7|4|3|A1".to_string()));
    }

    #[test]
    fn vba_discovers_current_regions_and_worksheet_dimensions() {
        let mut workbook = workbook();
        let module = parse_module(
            "Public Function InspectRegions() As String\n\
               Range(\"B2\").Value = 1\n\
               Range(\"C3\").Formula = \"=1+1\"\n\
               Range(\"D1\").Value = 1\n\
               Range(\"H2\").Value = 1\n\
               Range(\"J2\").Value = 1\n\
               Range(\"H6\").Value = 1\n\
               Range(\"J7\").Value = 1\n\
               InspectRegions = Rows.Count & \"|\" & Columns.Count & \"|\" & ActiveSheet.Rows.Count & \"|\" & Worksheets(1).Columns.Count & \"|\" & Application.Rows.Count & \"|\" & Application.Columns.Count & \"|\" & Range(\"B2\").CurrentRegion.Address(False, False) & \"|\" & Range(\"H2\").CurrentRegion.Address(False, False) & \"|\" & Range(\"I2\").CurrentRegion.Address(False, False) & \"|\" & Range(\"H6\").CurrentRegion.Address(False, False) & \"|\" & Range(\"I6\").CurrentRegion.Address(False, False)\n\
             End Function\n",
        )
        .unwrap();
        let result = {
            let mut host = WorkbookHost::new(&mut workbook, 0).unwrap();
            execute_with_host(&module, "InspectRegions", vec![], &mut host).unwrap()
        };

        assert_eq!(
            result,
            Value::String(
                "1048576|16384|1048576|16384|1048576|16384|B1:D3|H2|H2:J2|H6|H6:J7".to_string()
            )
        );
    }

    #[test]
    fn vba_moves_to_range_edges_in_all_four_directions() {
        let mut workbook = workbook();
        let module = parse_module(
            "Public Function InspectEnds() As String\n\
               Range(\"A1:A3\").Value = 1\n\
               Range(\"A5\").Value = 1\n\
               Range(\"B7\").Formula = \"=1+1\"\n\
               Range(\"C10:E10\").Value = 1\n\
               Range(\"G10\").Value = 1\n\
               InspectEnds = Range(\"A1\").End(xlDown).Address(False, False) & \"|\" & Range(\"A3\").End(xlDown).Address(False, False) & \"|\" & Range(\"A5\").End(xlDown).Address(False, False) & \"|\" & Range(\"A5\").End(xlUp).Address(False, False) & \"|\" & Range(\"A3\").End(xlUp).Address(False, False) & \"|\" & Range(\"C10\").End(xlToRight).Address(False, False) & \"|\" & Range(\"E10\").End(xlToRight).Address(False, False) & \"|\" & Range(\"G10\").End(xlToRight).Address(False, False) & \"|\" & Range(\"G10\").End(xlToLeft).Address(False, False) & \"|\" & Range(\"C10\").End(xlToLeft).Address(False, False) & \"|\" & Range(\"B1\").End(xlDown).Address(False, False)\n\
             End Function\n",
        )
        .unwrap();
        let result = {
            let mut host = WorkbookHost::new(&mut workbook, 0).unwrap();
            execute_with_host(&module, "InspectEnds", vec![], &mut host).unwrap()
        };

        assert_eq!(
            result,
            Value::String("A3|A5|A1048576|A3|A1|E10|G10|XFD10|E10|A10|B7".to_string())
        );
    }

    #[test]
    fn vba_iterates_range_cells_as_objects() {
        let mut workbook = workbook();
        let module = parse_module(
            "Public Function UpdateCells() As Long\n\
               Dim cell As Range\n\
               Range(\"A1:A3\").Value = 1\n\
               For Each cell In Range(\"A1:A3\")\n\
                 cell.Value = cell.Value + cell.Row\n\
                 UpdateCells = UpdateCells + cell.Value\n\
               Next cell\n\
             End Function\n",
        )
        .unwrap();
        let result = {
            let mut host = WorkbookHost::new(&mut workbook, 0).unwrap();
            execute_with_host(&module, "UpdateCells", vec![], &mut host).unwrap()
        };

        assert_eq!(result, Value::Integer(9));
        for (row, expected) in workbook.sheets[0].rows.iter().zip([2.0, 3.0, 4.0]) {
            assert!(matches!(
                row.cells[0].value,
                CellValue::Number(value) if value == expected
            ));
        }
    }

    #[test]
    fn vba_iterates_and_indexes_worksheet_collections() {
        let mut workbook = workbook();
        workbook.sheets.push(Sheet {
            name: "Data".to_string(),
            rows: Vec::new(),
            col_count: 0,
            col_widths: Vec::new(),
            default_col_width: 8.43,
            default_row_height: 15.0,
            merge_cells: Vec::new(),
            unsupported_elements: Vec::new(),
        });
        let module = parse_module(
            "Public Function FillWorksheets() As String\n\
               Dim ws As Worksheet\n\
               For Each ws In ThisWorkbook.Worksheets\n\
                 ws.Cells(1, 1).Value = ws.Index * 10\n\
               Next ws\n\
               FillWorksheets = Worksheets.Count & \"|\" & Application.Sheets.Item(2).Name & \"|\" & Worksheets(1).Range(\"A1\").Value + Worksheets.Item(2).Range(\"A1\").Value\n\
             End Function\n",
        )
        .unwrap();
        let result = {
            let mut host = WorkbookHost::new(&mut workbook, 0).unwrap();
            execute_with_host(&module, "FillWorksheets", vec![], &mut host).unwrap()
        };

        assert_eq!(result, Value::String("2|Data|30".to_string()));
        assert!(matches!(
            workbook.sheets[0].rows[0].cells[0].value,
            CellValue::Number(10.0)
        ));
        assert!(matches!(
            workbook.sheets[1].rows[0].cells[0].value,
            CellValue::Number(20.0)
        ));
    }

    #[test]
    fn vba_iterates_and_indexes_range_rows_and_columns() {
        let mut workbook = workbook();
        let module = parse_module(
            "Public Function FillBands() As String\n\
               Dim band As Range\n\
               Range(\"A1:C2\").Value = 1\n\
               For Each band In Range(\"A1:C2\").Rows\n\
                 band.Value = band.Row * 10\n\
               Next band\n\
               For Each band In Range(\"A1:C2\").Columns\n\
                 band.Cells(1, 1).Value = band.Column\n\
               Next band\n\
               FillBands = Range(\"A1:C2\").Rows.Count & \"|\" & Range(\"A1:C2\").Columns.Count & \"|\" & Range(\"A1:C2\").Rows.Item(2).Address(False, False)\n\
             End Function\n",
        )
        .unwrap();
        let result = {
            let mut host = WorkbookHost::new(&mut workbook, 0).unwrap();
            execute_with_host(&module, "FillBands", vec![], &mut host).unwrap()
        };

        assert_eq!(result, Value::String("2|3|A2:C2".to_string()));
        let first_row = &workbook.sheets[0].rows[0];
        assert!(matches!(first_row.cells[0].value, CellValue::Number(1.0)));
        assert!(matches!(first_row.cells[1].value, CellValue::Number(2.0)));
        assert!(matches!(first_row.cells[2].value, CellValue::Number(3.0)));
        for cell in &workbook.sheets[0].rows[1].cells {
            assert!(matches!(cell.value, CellValue::Number(20.0)));
        }
    }

    #[test]
    fn vba_reads_and_writes_range_formulas() {
        let mut workbook = workbook();
        let module = parse_module(
            "Public Function WriteFormulas() As String\n\
               Range(\"A1\").Value = 10\n\
               Range(\"A2\").Value = 20\n\
               Range(\"A3\").Formula = \"=SUM(A1:A2)\"\n\
               Range(\"B1:B2\").Formula2 = \"=A1*2\"\n\
               WriteFormulas = Range(\"A3\").Formula & \"|\" & Range(\"B1\").Formula2\n\
             End Function\n",
        )
        .unwrap();
        let result = {
            let mut host = WorkbookHost::new(&mut workbook, 0).unwrap();
            execute_with_host(&module, "WriteFormulas", vec![], &mut host).unwrap()
        };

        assert_eq!(result, Value::String("=SUM(A1:A2)|=A1*2".to_string()));
        assert_eq!(
            workbook.sheets[0].rows[2].cells[0].formula.as_deref(),
            Some("SUM(A1:A2)")
        );
        assert_eq!(
            workbook.sheets[0].rows[0].cells[1].formula.as_deref(),
            Some("A1*2")
        );
    }
}
