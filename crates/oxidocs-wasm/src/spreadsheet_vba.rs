// SPDX-License-Identifier: MIT OR Apache-2.0

use std::cmp::Ordering;
use std::collections::{BTreeMap, BTreeSet};

use oxicells_core::ir::{Cell, CellStyle, CellValue, MergeCell, Row, Sheet, Workbook};
use oxicells_core::{ReferenceShift, ShiftAxis};
use oxivba_core::ast::{ParamMode, ProcKind, Visibility};
#[cfg(test)]
use oxivba_core::execute_with_host;
use oxivba_core::{
    parse_module, ArrayDimension, ArrayValue, Host, ModuleItem, ObjectRef, Runtime, Value,
};
use serde::{Deserialize, Serialize};
use wasm_bindgen::prelude::*;

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
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
enum LookupOrientation {
    Vertical,
    Horizontal,
}

impl LookupOrientation {
    fn depth_name(self) -> &'static str {
        match self {
            Self::Vertical => "column",
            Self::Horizontal => "row",
        }
    }
}

/// A rectangular block of values a lookup searches, built either from a cell
/// range or from a VBA array. Excel treats a one-dimensional VBA array as a
/// single row, so `Array(5, 15, 25)` has one row and three columns.
struct LookupTable {
    rows: usize,
    columns: usize,
    values: Vec<Value>,
}

impl LookupTable {
    fn get(&self, row: usize, column: usize) -> Value {
        self.values
            .get(row * self.columns + column)
            .cloned()
            .unwrap_or(Value::Empty)
    }
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
    WorksheetFunction,
    DebugConsole,
}

/// What `Range.Copy` set aside, as it stood when it was copied. Excel keeps a
/// live link to the cells instead, so changing them before pasting changes what
/// arrives; this build pastes what was there at the time.
struct Clipboard {
    /// Row-major, one entry per cell of the copied block.
    cells: Vec<Option<Cell>>,
    rows: u32,
    columns: u32,
    /// Where it came from, so a formula can be moved by the right offset.
    origin: CellAddress,
}

#[derive(Clone)]
struct FindState {
    range: CellRange,
    args: Vec<Value>,
    last_found: Option<CellAddress>,
}

struct WorkbookHost<'a> {
    workbook: &'a mut Workbook,
    active_sheet: usize,
    clipboard: Option<Clipboard>,
    selection: CellRange,
    screen_updating: bool,
    enable_events: bool,
    display_alerts: bool,
    calculation: i64,
    last_find: Option<FindState>,
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
            clipboard: None,
            selection: CellRange::single(CellAddress {
                sheet: active_sheet,
                row: 1,
                column: 0,
            }),
            screen_updating: true,
            enable_events: true,
            display_alerts: true,
            calculation: -4105,
            last_find: None,
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
                HostObject::WorksheetFunction => "WorksheetFunction",
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

    fn is_worksheet_function(&self, object: &ObjectRef) -> bool {
        matches!(
            self.objects.get(object.handle as usize),
            Some(HostObject::WorksheetFunction)
        )
    }

    fn is_debug_console(&self, object: &ObjectRef) -> bool {
        matches!(
            self.objects.get(object.handle as usize),
            Some(HostObject::DebugConsole)
        )
    }

    /// `Debug.Print` wants values, so a cell prints what it holds. A range of
    /// several cells holds an array, which is a type mismatch rather than
    /// something to print, and an object with nothing scalar behind it likewise.
    fn printed_value(&self, value: &Value) -> Result<Value, String> {
        let Value::Object(object) = value else {
            return Ok(value.clone());
        };
        let Some(range) = self.range(object) else {
            return Err(format!(
                "Debug.Print has no value to print for a {} object",
                object.kind
            ));
        };
        let value = self.range_value(range)?;
        if matches!(value, Value::Array(_)) {
            return Err(
                "Debug.Print cannot print a range of several cells as one value".to_string(),
            );
        }
        Ok(value)
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

    /// `Worksheets.Add` puts a sheet in front of the active one unless told
    /// otherwise, makes it active, and hands it back.
    fn add_worksheet(&mut self, args: &[Value]) -> Result<Value, String> {
        let given = |index: usize| match args.get(index) {
            Some(Value::Missing) | None => None,
            Some(value) => Some(value),
        };
        if args.len() > 3 {
            return Err("Worksheets.Add takes Before, After and Count".to_string());
        }
        if given(2).is_some() {
            return Err("Worksheets.Add cannot add several sheets at once".to_string());
        }
        // Before and After name a sheet with the object itself, not its name.
        let placed = |host: &Self, value: &Value| match value {
            Value::Object(object) => host.worksheet(object).ok_or_else(|| {
                format!("Worksheets.Add expects a worksheet, not a {} object", object.kind)
            }),
            value => host.worksheet_from_value(value),
        };
        let before = given(0).map(|value| placed(self, value)).transpose()?;
        let after = given(1).map(|value| placed(self, value)).transpose()?;
        if before.is_some() && after.is_some() {
            return Err("Worksheets.Add takes Before or After, not both".to_string());
        }
        let at = match (before, after) {
            (Some(before), _) => before,
            (_, Some(after)) => after + 1,
            _ => self.active_sheet,
        };

        let template = &self.workbook.sheets[self.active_sheet];
        let sheet = Sheet {
            name: self.unused_sheet_name(),
            rows: Vec::new(),
            col_count: 0,
            col_widths: Vec::new(),
            default_col_width: template.default_col_width,
            default_row_height: template.default_row_height,
            merge_cells: Vec::new(),
            unsupported_elements: Vec::new(),
        };
        self.workbook.sheets.insert(at, sheet);
        self.active_sheet = at;
        Ok(self.object(HostObject::Worksheet(at)))
    }

    /// Excel numbers a new sheet from a counter that never goes back, so a
    /// workbook that has had sheets removed skips those numbers. This build
    /// takes the lowest number nothing is using, which agrees whenever no sheet
    /// has been removed.
    fn unused_sheet_name(&self) -> String {
        (1..)
            .map(|number| format!("Sheet{number}"))
            .find(|candidate| {
                !self
                    .workbook
                    .sheets
                    .iter()
                    .any(|sheet| sheet.name.eq_ignore_ascii_case(candidate))
            })
            .expect("a workbook cannot hold every possible sheet name")
    }

    /// Where a sheet being copied or moved should land. Excel takes Before or
    /// After, never both, and in the browser there is no second workbook to
    /// send it to when neither is given.
    fn worksheet_destination(&self, args: &[Value], name: &str) -> Result<usize, String> {
        let given = |index: usize| match args.get(index) {
            Some(Value::Missing) | None => None,
            Some(value) => Some(value),
        };
        let placed = |value: &Value| match value {
            Value::Object(object) => self.worksheet(object).ok_or_else(|| {
                format!(
                    "Worksheet.{name} expects a worksheet, not a {} object",
                    object.kind
                )
            }),
            value => self.worksheet_from_value(value),
        };
        match (given(0), given(1)) {
            (Some(_), Some(_)) => Err(format!(
                "Worksheet.{name} takes Before or After, not both"
            )),
            (Some(before), None) => placed(before),
            (None, Some(after)) => placed(after).map(|after| after + 1),
            (None, None) => Err(format!(
                "Worksheet.{name} needs Before or After; the browser has one workbook"
            )),
        }
    }

    fn copy_worksheet(&mut self, sheet: usize, args: &[Value]) -> Result<(), String> {
        let at = self.worksheet_destination(args, "Copy")?;
        let mut copy = self.workbook.sheets[sheet].clone();
        copy.name = self.unused_copy_name(&self.workbook.sheets[sheet].name);
        self.workbook.sheets.insert(at, copy);
        // Objects already handed out name sheets by position, so the ones past
        // the new sheet would now point at their neighbour.
        self.invalidate_worksheets_from(at);
        self.active_sheet = at;
        Ok(())
    }

    fn move_worksheet(&mut self, sheet: usize, args: &[Value]) -> Result<(), String> {
        let at = self.worksheet_destination(args, "Move")?;
        if at == sheet || at == sheet + 1 {
            return Ok(());
        }
        let moved = self.workbook.sheets.remove(sheet);
        // Taking it out shifts everything after it down one.
        let at = if at > sheet { at - 1 } else { at };
        self.workbook.sheets.insert(at, moved);
        self.invalidate_worksheets_from(0);
        self.active_sheet = at;
        Ok(())
    }

    /// Excel names a copy after the sheet it came from, numbered from two.
    fn unused_copy_name(&self, base: &str) -> String {
        (2..)
            .map(|number| format!("{base} ({number})"))
            .find(|candidate| {
                !self
                    .workbook
                    .sheets
                    .iter()
                    .any(|sheet| sheet.name.eq_ignore_ascii_case(candidate))
            })
            .expect("a workbook cannot hold every possible sheet name")
    }

    fn invalidate_worksheets_from(&mut self, first: usize) {
        self.objects.retain(|object| match object {
            HostObject::Worksheet(index) => *index < first,
            _ => true,
        });
    }

    fn delete_worksheet(&mut self, sheet: usize) -> Result<(), String> {
        if self.workbook.sheets.len() <= 1 {
            return Err("a workbook must keep at least one worksheet".to_string());
        }
        self.workbook.sheets.remove(sheet);
        // Objects already handed out name sheets by position, so the ones past
        // the hole would now point at their neighbour.
        self.invalidate_worksheets_from(sheet);
        if self.active_sheet >= self.workbook.sheets.len() {
            self.active_sheet = self.workbook.sheets.len() - 1;
        }
        Ok(())
    }

    fn rename_worksheet(&mut self, sheet: usize, name: &str) -> Result<(), String> {
        if name.trim().is_empty() {
            return Err("a worksheet name cannot be empty".to_string());
        }
        if self
            .workbook
            .sheets
            .iter()
            .enumerate()
            .any(|(index, held)| index != sheet && held.name.eq_ignore_ascii_case(name))
        {
            return Err(format!("another worksheet is already called {name}"));
        }
        self.workbook.sheets[sheet].name = name.to_string();
        Ok(())
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

    fn append_worksheet_function_values(
        &self,
        value: &Value,
        values: &mut Vec<Value>,
    ) -> Result<(), String> {
        match value {
            Value::Array(array) => {
                for value in &array.values {
                    self.append_worksheet_function_values(value, values)?;
                }
            }
            Value::Object(object) => {
                let Some(range) = self.range(object) else {
                    return Err(format!(
                        "WorksheetFunction cannot aggregate a {} object",
                        object.kind
                    ));
                };
                let value = self.range_value(range)?;
                self.append_worksheet_function_values(&value, values)?;
            }
            value => values.push(value.clone()),
        }
        Ok(())
    }

    fn lookup_table(&self, value: &Value, name: &str) -> Result<LookupTable, String> {
        match value {
            Value::Object(object) => {
                let Some(range) = self.range(object) else {
                    return Err(format!(
                        "WorksheetFunction.{name} cannot search a {} object",
                        object.kind
                    ));
                };
                Self::range_cell_count(range)?;
                Ok(LookupTable {
                    rows: (range.end_row - range.start_row + 1) as usize,
                    columns: (range.end_column - range.start_column + 1) as usize,
                    values: range
                        .addresses()
                        .map(|address| self.cell_value(address))
                        .collect(),
                })
            }
            Value::Array(array) => {
                let (rows, columns) = match array.dimensions.as_slice() {
                    [columns] => (1, columns.length),
                    [rows, columns] => (rows.length, columns.length),
                    _ => {
                        return Err(format!(
                            "WorksheetFunction.{name} needs a one- or two-dimensional array"
                        ))
                    }
                };
                Ok(LookupTable {
                    rows,
                    columns,
                    values: array.values.clone(),
                })
            }
            value => Ok(LookupTable {
                rows: 1,
                columns: 1,
                values: vec![value.clone()],
            }),
        }
    }

    fn worksheet_lookup(
        &self,
        name: &str,
        args: &[Value],
        orientation: LookupOrientation,
    ) -> Result<Value, String> {
        let (needle, table, index) = match args {
            [needle, table, index] | [needle, table, index, _] => (needle, table, index),
            _ => {
                return Err(format!(
                    "WorksheetFunction.{name} expects three or four arguments"
                ))
            }
        };
        let approximate = lookup_boolean_argument(args.get(3), name)?;
        let index = lookup_index_argument(index, name)?;
        let table = self.lookup_table(table, name)?;
        let (lanes, depth) = match orientation {
            LookupOrientation::Vertical => (table.rows, table.columns),
            LookupOrientation::Horizontal => (table.columns, table.rows),
        };
        if index > depth {
            return Err(format!(
                "WorksheetFunction.{name} index {index} is outside the {depth}-{} lookup table",
                orientation.depth_name()
            ));
        }
        let key_at = |lane: usize| match orientation {
            LookupOrientation::Vertical => table.get(lane, 0),
            LookupOrientation::Horizontal => table.get(0, lane),
        };
        let lane = if approximate {
            sorted_lookup_position(lanes, false, key_at, needle).map(|position| position - 1)
        } else {
            (0..lanes).find(|lane| lookup_exact_matches(&key_at(*lane), needle))
        };
        let Some(lane) = lane else {
            return Err(format!(
                "WorksheetFunction.{name} did not find a matching value"
            ));
        };
        Ok(match orientation {
            LookupOrientation::Vertical => table.get(lane, index - 1),
            LookupOrientation::Horizontal => table.get(index - 1, lane),
        })
    }

    /// `Index` over a cell range answers with a *reference*, so `Set` on the
    /// result and `Range.Address` both work and the result feeds back into
    /// other worksheet functions. Over a VBA array it answers with values.
    fn criteria_range(&self, value: &Value, name: &str) -> Result<CellRange, String> {
        let Value::Object(object) = value else {
            return Err(format!(
                "WorksheetFunction.{name} needs a cell range, not a value"
            ));
        };
        let Some(range) = self.range(object) else {
            return Err(format!(
                "WorksheetFunction.{name} cannot test a {} object",
                object.kind
            ));
        };
        Self::range_cell_count(range)?;
        Ok(range)
    }

    /// A criteria argument may itself be a cell, in which case Excel tests
    /// against that cell's value.
    fn criteria_value(&self, value: &Value) -> Result<Value, String> {
        match value {
            Value::Object(object) => match self.range(object) {
                Some(range) => self.range_value(range),
                None => Err(format!(
                    "WorksheetFunction cannot use a {} object as criteria",
                    object.kind
                )),
            },
            value => Ok(value.clone()),
        }
    }

    /// `SumIf` and `AverageIf` anchor their aggregated range at its top-left
    /// corner and stretch it to the shape of the range being tested, so
    /// `SumIf(A1:A4, ">15", C1)` aggregates `C1:C4`.
    fn stretched_range(
        &self,
        value: &Value,
        shape: CellRange,
        name: &str,
    ) -> Result<CellRange, String> {
        let source = self.criteria_range(value, name)?;
        let stretched = CellRange {
            sheet: source.sheet,
            start_row: source.start_row,
            start_column: source.start_column,
            end_row: source
                .start_row
                .saturating_add(shape.end_row - shape.start_row)
                .min(MAX_WORKSHEET_ROW),
            end_column: source
                .start_column
                .saturating_add(shape.end_column - shape.start_column)
                .min(MAX_WORKSHEET_COLUMN),
        };
        Self::range_cell_count(stretched)?;
        Ok(stretched)
    }

    fn worksheet_conditional(
        &self,
        name: &str,
        args: &[Value],
        kind: ConditionalKind,
    ) -> Result<Value, String> {
        let aggregates = kind != ConditionalKind::Count;
        if args.len() < 2 || args.len() > if aggregates { 3 } else { 2 } {
            return Err(format!(
                "WorksheetFunction.{name} expects {} arguments",
                if aggregates { "two or three" } else { "two" }
            ));
        }
        let tested = self.criteria_range(&args[0], name)?;
        let criteria = parse_criteria(&self.criteria_value(&args[1])?);
        let aggregated = match args.get(2) {
            Some(value) => self.stretched_range(value, tested, name)?,
            None => tested,
        };

        let mut matches = 0usize;
        let mut numbers = 0usize;
        let mut total = 0.0;
        for (tested, aggregated) in tested.addresses().zip(aggregated.addresses()) {
            if !criteria.matches(&self.cell_value(tested)) {
                continue;
            }
            matches += 1;
            if let Some(number) = criteria_number(&self.cell_value(aggregated)) {
                numbers += 1;
                total += number;
            }
        }
        conditional_result(name, kind, matches, numbers, total)
    }

    fn worksheet_conditional_set(
        &self,
        name: &str,
        args: &[Value],
        kind: ConditionalKind,
    ) -> Result<Value, String> {
        let (aggregated, pairs) = if kind == ConditionalKind::Count {
            (None, args)
        } else {
            let Some((aggregated, pairs)) = args.split_first() else {
                return Err(format!(
                    "WorksheetFunction.{name} expects a range to aggregate"
                ));
            };
            (Some(aggregated), pairs)
        };
        if pairs.len() < 2 {
            return Err(format!(
                "WorksheetFunction.{name} expects a range and criteria pair"
            ));
        }

        let mut tests = Vec::new();
        let mut shape = None;
        // Excel ignores a trailing range that has no criteria of its own.
        for pair in pairs.chunks(2) {
            let [tested, criteria] = pair else { continue };
            let tested = self.criteria_range(tested, name)?;
            let extent = (
                tested.end_row - tested.start_row,
                tested.end_column - tested.start_column,
            );
            match shape {
                None => shape = Some(extent),
                Some(shape) if shape == extent => {}
                Some(_) => {
                    return Err(format!(
                        "WorksheetFunction.{name} needs every criteria range to have one shape"
                    ))
                }
            }
            tests.push((tested, parse_criteria(&self.criteria_value(criteria)?)));
        }
        let Some((rows, columns)) = shape else {
            return Err(format!(
                "WorksheetFunction.{name} expects a range and criteria pair"
            ));
        };
        let aggregated = match aggregated {
            Some(value) => {
                let aggregated = self.criteria_range(value, name)?;
                if (
                    aggregated.end_row - aggregated.start_row,
                    aggregated.end_column - aggregated.start_column,
                ) != (rows, columns)
                {
                    return Err(format!(
                        "WorksheetFunction.{name} needs the aggregated range to match the criteria shape"
                    ));
                }
                Some(aggregated)
            }
            None => None,
        };

        let mut matches = 0usize;
        let mut numbers = 0usize;
        let mut total = 0.0;
        for row in 0..=rows {
            for column in 0..=columns {
                let matched = tests.iter().all(|(tested, criteria)| {
                    criteria.matches(&self.cell_value(CellAddress {
                        sheet: tested.sheet,
                        row: tested.start_row + row,
                        column: tested.start_column + column,
                    }))
                });
                if !matched {
                    continue;
                }
                matches += 1;
                let Some(aggregated) = aggregated else {
                    continue;
                };
                let value = self.cell_value(CellAddress {
                    sheet: aggregated.sheet,
                    row: aggregated.start_row + row,
                    column: aggregated.start_column + column,
                });
                if let Some(number) = criteria_number(&value) {
                    numbers += 1;
                    total += number;
                }
            }
        }
        conditional_result(name, kind, matches, numbers, total)
    }

    fn worksheet_index(&mut self, args: &[Value]) -> Result<Value, String> {
        let (array, row, column) = match args {
            [array, row] => (array, row, None),
            [array, row, column] => (array, row, Some(column)),
            _ => return Err("WorksheetFunction.Index expects two or three arguments".to_string()),
        };
        let row = index_argument(row, "row")?;
        let column = match column {
            None | Some(Value::Missing) => None,
            Some(column) => Some(index_argument(column, "column")?),
        };

        if let Value::Object(object) = array {
            if let Some(range) = self.range(object) {
                let rows = (range.end_row - range.start_row + 1) as usize;
                let columns = (range.end_column - range.start_column + 1) as usize;
                let (row, column) = index_selection(rows, columns, row, column, true)?;
                let selected = CellRange {
                    sheet: range.sheet,
                    start_row: range.start_row + row.saturating_sub(1) as u32,
                    end_row: if row == 0 {
                        range.end_row
                    } else {
                        range.start_row + row as u32 - 1
                    },
                    start_column: range.start_column + column.saturating_sub(1) as u32,
                    end_column: if column == 0 {
                        range.end_column
                    } else {
                        range.start_column + column as u32 - 1
                    },
                };
                return Ok(self.object(HostObject::Range(selected)));
            }
        }

        let table = self.lookup_table(array, "Index")?;
        let (row, column) = index_selection(table.rows, table.columns, row, column, false)?;
        let (first_row, last_row) = if row == 0 {
            (0, table.rows)
        } else {
            (row - 1, row)
        };
        let (first_column, last_column) = if column == 0 {
            (0, table.columns)
        } else {
            (column - 1, column)
        };
        let mut values = Vec::new();
        for row in first_row..last_row {
            for column in first_column..last_column {
                values.push(table.get(row, column));
            }
        }
        if values.len() == 1 {
            return Ok(values.remove(0));
        }
        let selected_rows = last_row - first_row;
        let selected_columns = last_column - first_column;
        let dimensions = if selected_rows > 1 && selected_columns > 1 {
            vec![
                ArrayDimension {
                    lower_bound: 1,
                    length: selected_rows,
                },
                ArrayDimension {
                    lower_bound: 1,
                    length: selected_columns,
                },
            ]
        } else {
            vec![ArrayDimension {
                lower_bound: 1,
                length: values.len(),
            }]
        };
        Ok(Value::Array(ArrayValue {
            dimensions,
            values,
            element_default: Box::new(Value::Empty),
            resizable: true,
        }))
    }

    fn worksheet_match(&self, args: &[Value]) -> Result<Value, String> {
        let (needle, array) = match args {
            [needle, array] | [needle, array, _] => (needle, array),
            _ => return Err("WorksheetFunction.Match expects two or three arguments".to_string()),
        };
        let match_type = match_type_argument(args.get(2))?;
        let table = self.lookup_table(array, "Match")?;
        if table.rows > 1 && table.columns > 1 {
            return Err(
                "WorksheetFunction.Match needs a single row or column to search".to_string(),
            );
        }
        let count = table.rows * table.columns;
        let value_at = |index: usize| table.get(index / table.columns, index % table.columns);
        let position = if match_type == 0 {
            (0..count)
                .find(|index| lookup_exact_matches(&value_at(*index), needle))
                .map(|index| index + 1)
        } else {
            sorted_lookup_position(count, match_type < 0, value_at, needle)
        };
        position
            .map(|position| Value::Integer(position as i64))
            .ok_or_else(|| "WorksheetFunction.Match did not find a matching value".to_string())
    }

    /// The worksheet functions that read their arguments by position rather than
    /// aggregating everything handed to them.
    ///
    /// `Round` here is the worksheet's, which sends a half away from zero, so
    /// 2.5 becomes 3 and -2.5 becomes -3. VBA's own `Round` sends a half to the
    /// even neighbour instead and leaves 2.5 at 2; both are right, in their own
    /// language.
    fn worksheet_arithmetic(
        &self,
        name: &str,
        args: &[Value],
    ) -> Result<Option<Value>, String> {
        let rounding = ["round", "roundup", "rounddown"]
            .iter()
            .find(|candidate| name.eq_ignore_ascii_case(candidate));
        if let Some(rounding) = rounding {
            let (value, digits) = match args {
                [value] => (value, None),
                [value, digits] => (value, Some(digits)),
                _ => return Err(format!("WorksheetFunction.{name} expects a number and digits")),
            };
            let value = worksheet_number(value, name)?;
            let digits = match digits {
                None | Some(Value::Missing) => 0.0,
                Some(digits) => worksheet_number(digits, name)?.trunc(),
            };
            let scale = 10_f64.powf(digits);
            let scaled = value * scale;
            let rounded = match *rounding {
                "roundup" => scaled.abs().ceil() * scaled.signum(),
                "rounddown" => scaled.abs().floor() * scaled.signum(),
                _ => scaled.abs().round() * scaled.signum(),
            };
            return Ok(Some(numeric_result(rounded / scale)));
        }

        if name.eq_ignore_ascii_case("power") {
            let [base, exponent] = args else {
                return Err("WorksheetFunction.Power expects a number and a power".to_string());
            };
            let result = worksheet_number(base, name)?.powf(worksheet_number(exponent, name)?);
            if !result.is_finite() {
                return Err("WorksheetFunction.Power has no answer for those".to_string());
            }
            return Ok(Some(numeric_result(result)));
        }

        if name.eq_ignore_ascii_case("trim") {
            let [value] = args else {
                return Err("WorksheetFunction.Trim expects one value".to_string());
            };
            // The worksheet's Trim squeezes runs of spaces as well as stripping
            // the ends, where VBA's only strips the ends.
            let text = find_value_text(&self.criteria_value(value)?);
            return Ok(Some(Value::String(
                text.split_whitespace().collect::<Vec<_>>().join(" "),
            )));
        }

        if name.eq_ignore_ascii_case("proper") {
            let [value] = args else {
                return Err("WorksheetFunction.Proper expects one value".to_string());
            };
            let text = find_value_text(&self.criteria_value(value)?);
            let mut proper = String::with_capacity(text.len());
            let mut starting = true;
            for character in text.chars() {
                if starting {
                    proper.extend(character.to_uppercase());
                } else {
                    proper.extend(character.to_lowercase());
                }
                // Anything that is not a letter starts a new word, so an
                // apostrophe or a digit capitalises what follows it.
                starting = !character.is_alphabetic();
            }
            return Ok(Some(Value::String(proper)));
        }

        let ranking = ["large", "small"]
            .iter()
            .find(|candidate| name.eq_ignore_ascii_case(candidate));
        if let Some(ranking) = ranking {
            let [values, rank] = args else {
                return Err(format!("WorksheetFunction.{name} expects values and a rank"));
            };
            let rank = worksheet_number(rank, name)?.trunc();
            let mut numbers = self.worksheet_numbers(values, name)?;
            if rank < 1.0 || rank > numbers.len() as f64 {
                return Err(format!(
                    "WorksheetFunction.{name} has no value ranked {rank}"
                ));
            }
            numbers.sort_by(|left, right| left.partial_cmp(right).unwrap_or(Ordering::Equal));
            let index = if *ranking == "large" {
                numbers.len() - rank as usize
            } else {
                rank as usize - 1
            };
            return Ok(Some(numeric_result(numbers[index])));
        }

        if name.eq_ignore_ascii_case("median") {
            let mut numbers = Vec::new();
            for value in args {
                numbers.append(&mut self.worksheet_numbers(value, name)?);
            }
            if numbers.is_empty() {
                return Err("WorksheetFunction.Median has no numeric values".to_string());
            }
            numbers.sort_by(|left, right| left.partial_cmp(right).unwrap_or(Ordering::Equal));
            let middle = numbers.len() / 2;
            let median = if numbers.len() % 2 == 0 {
                (numbers[middle - 1] + numbers[middle]) / 2.0
            } else {
                numbers[middle]
            };
            return Ok(Some(numeric_result(median)));
        }

        if name.eq_ignore_ascii_case("sumproduct") {
            if args.len() < 2 {
                return Err("WorksheetFunction.SumProduct expects two or more ranges".to_string());
            }
            let columns = args
                .iter()
                .map(|value| self.worksheet_numbers(value, name))
                .collect::<Result<Vec<_>, _>>()?;
            let length = columns[0].len();
            if columns.iter().any(|column| column.len() != length) {
                return Err(
                    "WorksheetFunction.SumProduct needs ranges of the same size".to_string(),
                );
            }
            let total = (0..length)
                .map(|index| columns.iter().map(|column| column[index]).product::<f64>())
                .sum::<f64>();
            return Ok(Some(numeric_result(total)));
        }

        Ok(None)
    }

    /// The numbers inside a range or array, with text and blanks passed over
    /// the way the ranking functions do.
    fn worksheet_numbers(&self, value: &Value, name: &str) -> Result<Vec<f64>, String> {
        let mut values = Vec::new();
        self.append_worksheet_function_values(value, &mut values)
            .map_err(|_| format!("WorksheetFunction.{name} cannot read those values"))?;
        Ok(values.iter().filter_map(criteria_number).collect())
    }

    fn worksheet_function(&mut self, name: &str, args: &[Value]) -> Result<Value, String> {
        if args.is_empty() {
            return Err(format!(
                "WorksheetFunction.{name} expects at least one argument"
            ));
        }
        if name.eq_ignore_ascii_case("vlookup") {
            return self.worksheet_lookup(name, args, LookupOrientation::Vertical);
        }
        if name.eq_ignore_ascii_case("hlookup") {
            return self.worksheet_lookup(name, args, LookupOrientation::Horizontal);
        }
        if name.eq_ignore_ascii_case("match") {
            return self.worksheet_match(args);
        }
        if name.eq_ignore_ascii_case("index") {
            return self.worksheet_index(args);
        }
        if let Some(result) = self.worksheet_arithmetic(name, args)? {
            return Ok(result);
        }
        for (conditional, kind) in [
            ("countif", ConditionalKind::Count),
            ("sumif", ConditionalKind::Sum),
            ("averageif", ConditionalKind::Average),
        ] {
            if name.eq_ignore_ascii_case(conditional) {
                return self.worksheet_conditional(name, args, kind);
            }
            if name.len() == conditional.len() + 1
                && name[..conditional.len()].eq_ignore_ascii_case(conditional)
                && name.ends_with(['s', 'S'])
            {
                return self.worksheet_conditional_set(name, args, kind);
            }
        }
        let mut values = Vec::new();
        for value in args {
            self.append_worksheet_function_values(value, &mut values)?;
        }
        if name.eq_ignore_ascii_case("counta") {
            let count = values
                .iter()
                .filter(|value| {
                    !matches!(
                        value,
                        Value::Empty | Value::Missing | Value::Nothing | Value::Null
                    )
                })
                .count();
            return Ok(Value::Integer(count as i64));
        }

        let mut numbers = Vec::new();
        for value in values {
            match value {
                Value::Integer(value) => numbers.push(value as f64),
                Value::Double(value) if value.is_finite() => numbers.push(value),
                Value::Error(value) => {
                    return Err(format!(
                        "WorksheetFunction.{name} encountered Error {value}"
                    ));
                }
                _ => {}
            }
        }
        if name.eq_ignore_ascii_case("count") {
            return Ok(Value::Integer(numbers.len() as i64));
        }

        let result = if name.eq_ignore_ascii_case("sum") {
            numbers.iter().sum()
        } else if name.eq_ignore_ascii_case("average") {
            if numbers.is_empty() {
                return Err("WorksheetFunction.Average has no numeric values".to_string());
            }
            numbers.iter().sum::<f64>() / numbers.len() as f64
        } else if name.eq_ignore_ascii_case("min") {
            numbers.into_iter().reduce(f64::min).unwrap_or(0.0)
        } else if name.eq_ignore_ascii_case("max") {
            numbers.into_iter().reduce(f64::max).unwrap_or(0.0)
        } else {
            return Err(format!(
                "WorksheetFunction.{name} is not supported in the browser"
            ));
        };
        Ok(numeric_result(result))
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

    fn find_in_range(&mut self, range: CellRange, args: &[Value]) -> Result<Value, String> {
        if args.is_empty() || args.len() > 9 {
            return Err("Range.Find expects between one and nine arguments".to_string());
        }
        Self::range_cell_count(range)?;
        let what = args
            .first()
            .filter(|value| !matches!(value, Value::Missing))
            .ok_or_else(|| "Range.Find What argument is required".to_string())?;
        let look_in = find_integer_argument(args.get(2), -4163, "LookIn")?;
        if !matches!(look_in, -4163 | -4123) {
            return Err(format!("unsupported Range.Find LookIn constant: {look_in}"));
        }
        let look_at = find_integer_argument(args.get(3), 2, "LookAt")?;
        if !matches!(look_at, 1 | 2) {
            return Err(format!("unsupported Range.Find LookAt constant: {look_at}"));
        }
        let search_order = find_integer_argument(args.get(4), 1, "SearchOrder")?;
        if !matches!(search_order, 1 | 2) {
            return Err(format!(
                "unsupported Range.Find SearchOrder constant: {search_order}"
            ));
        }
        let search_direction = find_integer_argument(args.get(5), 1, "SearchDirection")?;
        if !matches!(search_direction, 1 | 2) {
            return Err(format!(
                "unsupported Range.Find SearchDirection constant: {search_direction}"
            ));
        }
        let match_case = find_boolean_argument(args.get(6), false, "MatchCase")?;
        if find_boolean_argument(args.get(7), false, "MatchByte")? {
            return Err("Range.Find MatchByte:=True is not supported in the browser".to_string());
        }
        if find_boolean_argument(args.get(8), false, "SearchFormat")? {
            return Err(
                "Range.Find SearchFormat:=True is not supported in the browser".to_string(),
            );
        }

        let mut addresses = if search_order == 1 {
            range.addresses().collect::<Vec<_>>()
        } else {
            let mut addresses = Vec::with_capacity(Self::range_cell_count(range)?);
            for column in range.start_column..=range.end_column {
                for row in range.start_row..=range.end_row {
                    addresses.push(CellAddress {
                        sheet: range.sheet,
                        row,
                        column,
                    });
                }
            }
            addresses
        };
        let after = match args.get(1) {
            None | Some(Value::Missing) => addresses[0],
            Some(Value::Object(object)) => {
                let after = self
                    .range(object)
                    .filter(|range| range.is_single())
                    .ok_or_else(|| "Range.Find After must be a single cell".to_string())?;
                let address = after.addresses().next().unwrap();
                if !addresses.contains(&address) {
                    return Err("Range.Find After cell must be inside the search range".to_string());
                }
                address
            }
            _ => return Err("Range.Find After must be a single cell".to_string()),
        };
        let after_index = addresses
            .iter()
            .position(|address| *address == after)
            .unwrap();
        if search_direction == 1 {
            let address_count = addresses.len();
            addresses.rotate_left((after_index + 1) % address_count);
        } else {
            addresses.rotate_left(after_index);
            addresses.reverse();
        }
        let needle = find_value_text(what);
        let found = addresses.into_iter().find(|address| {
            let candidate = self.find_cell_text(*address, look_in);
            find_text_matches(&candidate, &needle, look_at == 1, match_case)
        });
        let result = found
            .map(|address| self.object(HostObject::Range(CellRange::single(address))))
            .unwrap_or(Value::Nothing);
        self.last_find = Some(FindState {
            range,
            args: args.to_vec(),
            last_found: found,
        });
        Ok(result)
    }

    fn find_again(&mut self, args: &[Value], search_direction: i64) -> Result<Value, String> {
        if args.len() > 1 {
            return Err("Range.FindNext and FindPrevious expect zero or one argument".to_string());
        }
        let state = self
            .last_find
            .clone()
            .ok_or_else(|| "Range.FindNext requires a preceding Range.Find call".to_string())?;
        let after = match args.first() {
            Some(Value::Object(object)) => Value::Object(object.clone()),
            Some(_) => {
                return Err(
                    "Range.FindNext and FindPrevious After must be a single cell".to_string(),
                )
            }
            None => match state.last_found {
                Some(address) => self.object(HostObject::Range(CellRange::single(address))),
                None => return Ok(Value::Nothing),
            },
        };
        let mut find_args = state.args;
        find_args.resize(6, Value::Missing);
        find_args[1] = after;
        find_args[5] = Value::Integer(search_direction);
        self.find_in_range(state.range, &find_args)
    }

    fn find_cell_text(&self, address: CellAddress, look_in: i64) -> String {
        if look_in == -4123 {
            if let Some(formula) = self
                .workbook
                .sheets
                .get(address.sheet)
                .and_then(|sheet| sheet.rows.iter().find(|row| row.index == address.row))
                .and_then(|row| row.cells.iter().find(|cell| cell.col == address.column))
                .and_then(|cell| cell.formula.as_deref())
            {
                return format!("={formula}");
            }
        }
        find_value_text(&self.cell_value(address))
    }

    fn replace_in_range(&mut self, range: CellRange, args: &[Value]) -> Result<Value, String> {
        if !(2..=8).contains(&args.len()) {
            return Err("Range.Replace expects between two and eight arguments".to_string());
        }
        Self::range_cell_count(range)?;
        let what = args
            .first()
            .filter(|value| !matches!(value, Value::Missing))
            .ok_or_else(|| "Range.Replace What argument is required".to_string())?;
        let replacement = args
            .get(1)
            .filter(|value| !matches!(value, Value::Missing))
            .ok_or_else(|| "Range.Replace Replacement argument is required".to_string())?;
        let look_at = find_integer_argument(args.get(2), 2, "LookAt")?;
        if !matches!(look_at, 1 | 2) {
            return Err(format!(
                "unsupported Range.Replace LookAt constant: {look_at}"
            ));
        }
        let search_order = find_integer_argument(args.get(3), 1, "SearchOrder")?;
        if !matches!(search_order, 1 | 2) {
            return Err(format!(
                "unsupported Range.Replace SearchOrder constant: {search_order}"
            ));
        }
        let match_case = find_boolean_argument(args.get(4), false, "MatchCase")?;
        if find_boolean_argument(args.get(5), false, "MatchByte")? {
            return Err(
                "Range.Replace MatchByte:=True is not supported in the browser".to_string(),
            );
        }
        if find_boolean_argument(args.get(6), false, "SearchFormat")? {
            return Err(
                "Range.Replace SearchFormat:=True is not supported in the browser".to_string(),
            );
        }
        if find_boolean_argument(args.get(7), false, "ReplaceFormat")? {
            return Err(
                "Range.Replace ReplaceFormat:=True is not supported in the browser".to_string(),
            );
        }

        let needle = find_value_text(what);
        let replacement_text = find_value_text(replacement);
        let mut changed = false;
        for address in range.addresses() {
            let formula = self
                .workbook
                .sheets
                .get(address.sheet)
                .and_then(|sheet| sheet.rows.iter().find(|row| row.index == address.row))
                .and_then(|row| row.cells.iter().find(|cell| cell.col == address.column))
                .and_then(|cell| cell.formula.as_deref())
                .map(|formula| format!("={formula}"));
            let candidate = formula
                .clone()
                .unwrap_or_else(|| find_value_text(&self.cell_value(address)));
            let Some(replaced) = replace_matching_text(
                &candidate,
                &needle,
                &replacement_text,
                look_at == 1,
                match_case,
            ) else {
                continue;
            };
            if formula.is_some() {
                if replaced.starts_with('=') {
                    self.set_cell_formula(address, replaced)?;
                } else {
                    self.set_cell_value(address, CellValue::String(replaced))?;
                }
            } else if look_at == 1 {
                self.set_cell_value(address, to_cell_value(replacement.clone())?)?;
            } else {
                self.set_cell_value(address, CellValue::String(replaced))?;
            }
            changed = true;
        }
        Ok(Value::Boolean(changed))
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
            [row] => (optional_integer_offset(row, 0, "row")?, 0),
            [row, column] => (
                optional_integer_offset(row, 0, "row")?,
                optional_integer_offset(column, 0, "column")?,
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
            [rows] => (
                optional_positive_index(rows, range.end_row - range.start_row + 1, "row size")?,
                current_columns,
            ),
            [rows, columns] => (
                optional_positive_index(rows, range.end_row - range.start_row + 1, "row size")?,
                optional_positive_index(columns, current_columns, "column size")?,
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

    /// Which way `Insert` and `Delete` move the cells around a range. Excel
    /// decides from the range's shape when the caller does not say: a range
    /// taller than it is wide moves sideways, anything else moves vertically.
    fn shift_direction(range: CellRange, args: &[Value], inserting: bool) -> Result<bool, String> {
        let sideways = match args {
            [] | [Value::Missing] => {
                (range.end_row - range.start_row) > (range.end_column - range.start_column)
            }
            [shift] => match shift {
                Value::Integer(-4121) if inserting => false,
                Value::Integer(-4161) if inserting => true,
                Value::Integer(-4162) if !inserting => false,
                Value::Integer(-4159) if !inserting => true,
                _ => {
                    return Err(format!(
                        "Range.{} does not understand that shift",
                        if inserting { "Insert" } else { "Delete" }
                    ))
                }
            },
            _ => {
                return Err(format!(
                    "Range.{} takes at most one shift",
                    if inserting { "Insert" } else { "Delete" }
                ))
            }
        };
        Ok(sideways)
    }

    fn insert_range(&mut self, range: CellRange, args: &[Value]) -> Result<Value, String> {
        let sideways = Self::shift_direction(range, args, true)?;
        self.shift_cells(range, sideways, true)?;
        Ok(Value::Empty)
    }

    fn delete_range(&mut self, range: CellRange, args: &[Value]) -> Result<Value, String> {
        let sideways = Self::shift_direction(range, args, false)?;
        self.shift_cells(range, sideways, false)?;
        Ok(Value::Empty)
    }

    /// Moves everything the range pushes ahead of it, drops what a removal
    /// covers, and rewrites the sheet's formulas to match.
    ///
    /// Only the rows or columns the range itself spans move, so inserting at
    /// `B2` leaves column C where it was.
    fn shift_cells(
        &mut self,
        range: CellRange,
        sideways: bool,
        inserting: bool,
    ) -> Result<(), String> {
        let axis = if sideways {
            ShiftAxis::Columns
        } else {
            ShiftAxis::Rows
        };
        let (at, span, across) = if sideways {
            (
                range.start_column + 1,
                range.end_column - range.start_column + 1,
                (range.start_row, range.end_row),
            )
        } else {
            (
                range.start_row,
                range.end_row - range.start_row + 1,
                (range.start_column + 1, range.end_column + 1),
            )
        };
        let count = if inserting {
            i64::from(span)
        } else {
            -i64::from(span)
        };

        let removed = if inserting { 0 } else { span };
        let past = at.saturating_add(removed);
        let moved = |value: u32| -> Option<u32> {
            if value < at {
                return Some(value);
            }
            if removed > 0 {
                if value < past {
                    return None;
                }
                return Some(value - removed);
            }
            Some(value.saturating_add(span))
        };
        // Only cells lying across the range's own width or height take part.
        let taking_part = |crossing: u32| crossing >= across.0 && crossing <= across.1;

        let Some(worksheet) = self.workbook.sheets.get_mut(range.sheet) else {
            return Err("worksheet is out of range".to_string());
        };
        if sideways {
            for row in &mut worksheet.rows {
                if !taking_part(row.index) {
                    continue;
                }
                row.cells.retain_mut(|cell| match moved(cell.col + 1) {
                    Some(column) => {
                        cell.col = column - 1;
                        true
                    }
                    None => false,
                });
            }
        } else {
            // A row only partly taking part has to be split, so cells move
            // between rows rather than rows moving whole.
            let mut carried: Vec<(u32, Cell)> = Vec::new();
            for row in &mut worksheet.rows {
                row.cells.retain_mut(|cell| {
                    if !taking_part(cell.col + 1) {
                        return true;
                    }
                    match moved(row.index) {
                        Some(index) if index == row.index => true,
                        Some(index) => {
                            carried.push((index, cell.clone()));
                            false
                        }
                        None => false,
                    }
                });
            }
            for (index, cell) in carried {
                let row = match worksheet.rows.iter().position(|row| row.index == index) {
                    Some(position) => &mut worksheet.rows[position],
                    None => {
                        worksheet.rows.push(Row {
                            index,
                            cells: Vec::new(),
                            height: None,
                        });
                        worksheet.rows.sort_by_key(|row| row.index);
                        worksheet
                            .rows
                            .iter_mut()
                            .find(|row| row.index == index)
                            .unwrap()
                    }
                };
                match row.cells.iter_mut().find(|held| held.col == cell.col) {
                    Some(held) => *held = cell,
                    None => {
                        row.cells.push(cell);
                        row.cells.sort_by_key(|cell| cell.col);
                    }
                }
            }
            worksheet.rows.retain(|row| !row.cells.is_empty());
        }

        worksheet.merge_cells.retain_mut(|merge| {
            let (near, far) = if sideways {
                (merge.start_row, merge.end_row)
            } else {
                (merge.start_col + 1, merge.end_col + 1)
            };
            if !taking_part(near) && !taking_part(far) {
                return true;
            }
            // A band reaching only part of the way across a merge cannot carry
            // it, so Excel takes the merge apart and moves the cells alone.
            if !taking_part(near) || !taking_part(far) {
                return false;
            }
            let (start, end) = if sideways {
                (merge.start_col + 1, merge.end_col + 1)
            } else {
                (merge.start_row, merge.end_row)
            };
            // A merge the removal swallows whole goes; one it clips shrinks.
            let start = moved(start).unwrap_or(at);
            let end = moved(end).unwrap_or_else(|| at.saturating_sub(1));
            if start > end {
                return false;
            }
            if sideways {
                merge.start_col = start - 1;
                merge.end_col = end - 1;
            } else {
                merge.start_row = start;
                merge.end_row = end;
            }
            true
        });

        // Every sheet's formulas are rewritten, not just this one: a formula on
        // another sheet naming this one follows the cells that moved.
        let moved_sheet = self.workbook.sheets[range.sheet].name.clone();
        let shift = ReferenceShift {
            axis,
            at,
            count,
            across,
            sheet: Some(&moved_sheet),
        };
        for worksheet in &mut self.workbook.sheets {
            for row in &mut worksheet.rows {
                for cell in &mut row.cells {
                    let Some(formula) = cell.formula.as_ref() else {
                        continue;
                    };
                    // A formula this build cannot read is left as the author
                    // wrote it rather than silently half-moved.
                    if let Ok(shifted) = oxicells_core::shift_formula_references(formula, &shift) {
                        cell.formula = Some(shifted);
                    }
                }
            }
        }
        Ok(())
    }

    /// Sorts the rows or columns of a range by up to three keys.
    ///
    /// Excel orders values by kind before value — numbers, then text, then
    /// Booleans — and leaves blanks at the end whichever way the sort runs.
    /// Text compares without regard to case, and equal values keep the order
    /// they were already in.
    fn sort_range(&mut self, range: CellRange, args: &[Value]) -> Result<(), String> {
        let given = |index: usize| match args.get(index) {
            Some(Value::Missing) | None => None,
            Some(value) => Some(value),
        };
        if given(9).is_some_and(|value| matches!(value, Value::Boolean(true))) {
            return Err("Range.Sort cannot tell case apart in the browser".to_string());
        }
        let sideways = match given(10) {
            None => false,
            Some(value) => match sort_number(value, "Orientation")? {
                1 => false,
                2 => true,
                other => return Err(format!("Range.Sort has no orientation {other}")),
            },
        };
        let header = match given(7) {
            None => false,
            Some(value) => match sort_number(value, "Header")? {
                2 => false,
                1 => true,
                0 => {
                    return Err(
                        "Range.Sort cannot guess whether a range has a header row".to_string()
                    )
                }
                other => return Err(format!("Range.Sort has no header setting {other}")),
            },
        };

        let mut keys = Vec::new();
        for (key, order) in [(0, 1), (2, 4), (5, 6)] {
            let Some(key) = given(key) else { continue };
            let Value::Object(object) = key else {
                return Err("Range.Sort takes a cell as a key".to_string());
            };
            let Some(key) = self.range(object) else {
                return Err("Range.Sort takes a cell as a key".to_string());
            };
            let lane = if sideways {
                key.start_row
            } else {
                key.start_column
            };
            let (first, last) = if sideways {
                (range.start_row, range.end_row)
            } else {
                (range.start_column, range.end_column)
            };
            if lane < first || lane > last {
                // Excel quietly sorts nothing here; saying so is more use.
                return Err("Range.Sort was given a key outside the range".to_string());
            }
            let descending = match given(order) {
                None => false,
                Some(value) => match sort_number(value, "Order")? {
                    1 => false,
                    2 => true,
                    other => return Err(format!("Range.Sort has no order {other}")),
                },
            };
            keys.push((lane, descending));
        }
        if keys.is_empty() {
            return Err("Range.Sort expects at least one key".to_string());
        }

        // Each line is one row of the range, or one column when sorting sideways.
        let (first, last) = if sideways {
            (range.start_column, range.end_column)
        } else {
            (range.start_row, range.end_row)
        };
        let first = if header { first + 1 } else { first };
        if first > last {
            return Ok(());
        }
        let mut lines: Vec<u32> = (first..=last).collect();
        let cell_at = |line: u32, lane: u32| {
            if sideways {
                CellAddress {
                    sheet: range.sheet,
                    row: lane,
                    column: line,
                }
            } else {
                CellAddress {
                    sheet: range.sheet,
                    row: line,
                    column: lane,
                }
            }
        };
        lines.sort_by(|left, right| {
            for (lane, descending) in &keys {
                let ordering = sort_compare(
                    &self.cell_value(cell_at(*left, *lane)),
                    &self.cell_value(cell_at(*right, *lane)),
                    *descending,
                );
                if ordering != Ordering::Equal {
                    return ordering;
                }
            }
            Ordering::Equal
        });
        // Read every line out before writing any back, since they swap places.
        let across = if sideways {
            (range.start_row, range.end_row)
        } else {
            (range.start_column, range.end_column)
        };
        let taken: Vec<Vec<Option<Cell>>> = lines
            .iter()
            .map(|line| {
                (across.0..=across.1)
                    .map(|lane| {
                        let address = cell_at(*line, lane);
                        self.workbook.sheets[address.sheet]
                            .rows
                            .iter()
                            .find(|row| row.index == address.row)
                            .and_then(|row| row.cells.iter().find(|cell| cell.col == address.column))
                            .cloned()
                    })
                    .collect()
            })
            .collect();

        for (offset, held) in taken.into_iter().enumerate() {
            let line = first + offset as u32;
            for (lane, cell) in (across.0..=across.1).zip(held) {
                let address = cell_at(line, lane);
                let sheet = &mut self.workbook.sheets[address.sheet];
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
                row.cells.retain(|held| held.col != address.column);
                if let Some(mut cell) = cell {
                    cell.col = address.column;
                    row.cells.push(cell);
                    row.cells.sort_by_key(|cell| cell.col);
                }
            }
        }
        Ok(())
    }

    fn fill_clipboard(&mut self, source: CellRange) -> Result<(), String> {
        Self::range_cell_count(source)?;
        let cells = source
            .addresses()
            .map(|address| {
                self.workbook.sheets[address.sheet]
                    .rows
                    .iter()
                    .find(|row| row.index == address.row)
                    .and_then(|row| row.cells.iter().find(|cell| cell.col == address.column))
                    .cloned()
            })
            .collect();
        self.clipboard = Some(Clipboard {
            cells,
            rows: source.end_row - source.start_row + 1,
            columns: source.end_column - source.start_column + 1,
            origin: CellAddress {
                sheet: source.sheet,
                row: source.start_row,
                column: source.start_column,
            },
        });
        Ok(())
    }

    /// Pastes what `Copy` set aside. `xlPasteValues` drops formulas and keeps
    /// the value each cell was holding, `xlPasteFormats` keeps only the styling,
    /// and the default brings everything, moving a formula's relative
    /// references by the distance it travelled.
    fn paste_special(&mut self, target: CellRange, args: &[Value]) -> Result<Value, String> {
        let given = |index: usize| match args.get(index) {
            Some(Value::Missing) | None => None,
            Some(value) => Some(value),
        };
        let kind = match given(0) {
            None => -4104,
            Some(value) => sort_number(value, "PasteSpecial")?,
        };
        let (values, formats) = match kind {
            -4104 => (true, true),
            -4163 => (true, false),
            -4122 => (false, true),
            other => return Err(format!("Range.PasteSpecial cannot paste {other}")),
        };
        if given(1).is_some() {
            return Err("Range.PasteSpecial cannot combine what it pastes".to_string());
        }
        if given(2).is_some_and(|value| matches!(value, Value::Boolean(true))) {
            return Err("Range.PasteSpecial cannot skip blanks in the browser".to_string());
        }
        let transpose = given(3).is_some_and(|value| matches!(value, Value::Boolean(true)));

        let Some(clipboard) = self.clipboard.take() else {
            return Err("Range.PasteSpecial has nothing to paste".to_string());
        };
        let (rows, columns) = if transpose {
            (clipboard.columns, clipboard.rows)
        } else {
            (clipboard.rows, clipboard.columns)
        };

        // A target larger than the block takes whole copies of it, and nothing
        // else: Excel refuses a target the block does not divide.
        let target_rows = target.end_row - target.start_row + 1;
        let target_columns = target.end_column - target.start_column + 1;
        let (down, across) = if target_rows == 1 && target_columns == 1 {
            (1, 1)
        } else if target_rows % rows == 0 && target_columns % columns == 0 {
            (target_rows / rows, target_columns / columns)
        } else {
            return Err(
                "Range.PasteSpecial needs a target the copied block fits into evenly".to_string(),
            );
        };

        for block_row in 0..down {
            for block_column in 0..across {
                for row in 0..rows {
                    for column in 0..columns {
                        let held = if transpose {
                            clipboard.cells.get((column * clipboard.columns + row) as usize)
                        } else {
                            clipboard.cells.get((row * clipboard.columns + column) as usize)
                        };
                        let address = CellAddress {
                            sheet: target.sheet,
                            row: target.start_row + block_row * rows + row,
                            column: target.start_column + block_column * columns + column,
                        };
                        // A formula moves by how far its own cell travelled,
                        // not by where the block's corner landed.
                        let came_from = CellAddress {
                            sheet: clipboard.origin.sheet,
                            row: clipboard.origin.row + if transpose { column } else { row },
                            column: clipboard.origin.column
                                + if transpose { row } else { column },
                        };
                        self.paste_cell(
                            address,
                            held.and_then(|cell| cell.as_ref()),
                            came_from,
                            values,
                            formats,
                            transpose,
                        )?;
                    }
                }
            }
        }
        self.clipboard = Some(clipboard);
        Ok(Value::Boolean(true))
    }

    fn paste_cell(
        &mut self,
        address: CellAddress,
        held: Option<&Cell>,
        came_from: CellAddress,
        values: bool,
        formats: bool,
        transpose: bool,
    ) -> Result<(), String> {
        let sheet = &mut self.workbook.sheets[address.sheet];
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
        if row.cells.iter().all(|cell| cell.col != address.column) {
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
            .unwrap();

        if formats {
            cell.style = held.map(|held| held.style.clone()).unwrap_or_default();
        }
        if !values {
            return Ok(());
        }
        let Some(held) = held else {
            cell.value = CellValue::Empty;
            cell.formula = None;
            return Ok(());
        };
        cell.value = held.value.clone();
        cell.formula = None;
        // Only a whole paste carries the formula, and it moves with the cell.
        if formats {
            if let Some(formula) = held.formula.as_ref() {
                if transpose {
                    // Which way Excel turns a transposed formula's references
                    // has not been measured, so it is not guessed at here.
                    return Err(
                        "Range.PasteSpecial cannot transpose a formula in the browser".to_string(),
                    );
                }
                let row_offset = i64::from(address.row) - i64::from(came_from.row);
                let column_offset = i64::from(address.column) - i64::from(came_from.column);
                cell.formula = Some(
                    oxicells_core::translate_formula_references(
                        formula,
                        row_offset,
                        column_offset,
                    )
                    .unwrap_or_else(|_| formula.clone()),
                );
                cell.value = CellValue::Empty;
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
        if args.is_empty() || matches!(args, [Value::Missing]) {
            self.fill_clipboard(source)?;
            return Ok(Value::Boolean(true));
        }
        let [Value::Object(destination)] = args else {
            return Err("Range.Copy expects one destination Range".to_string());
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
            if self.is_worksheet_function(receiver) {
                return self.worksheet_function(name, args).map(Some);
            }
            if self.is_debug_console(receiver) && name.eq_ignore_ascii_case("print") {
                let mut printed = Vec::with_capacity(args.len());
                for value in args {
                    printed.push(format_debug_value(&self.printed_value(value)?));
                }
                self.debug_output.push(printed.join("\t"));
                return Ok(Some(Value::Empty));
            }
            if let Some(sheet) = self.worksheet(receiver) {
                if name.eq_ignore_ascii_case("delete") {
                    return self.delete_worksheet(sheet).map(|()| Some(Value::Empty));
                }
                if name.eq_ignore_ascii_case("copy") {
                    return self.copy_worksheet(sheet, args).map(|()| Some(Value::Empty));
                }
                if name.eq_ignore_ascii_case("move") {
                    return self.move_worksheet(sheet, args).map(|()| Some(Value::Empty));
                }
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
            if self.is_worksheets(receiver) && name.eq_ignore_ascii_case("add") {
                return self.add_worksheet(args).map(Some);
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
                if name.eq_ignore_ascii_case("find") {
                    return self.find_in_range(range, args).map(Some);
                }
                if name.eq_ignore_ascii_case("findnext") {
                    return self.find_again(args, 1).map(Some);
                }
                if name.eq_ignore_ascii_case("findprevious") {
                    return self.find_again(args, 2).map(Some);
                }
                if name.eq_ignore_ascii_case("replace") {
                    return self.replace_in_range(range, args).map(Some);
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
                if name.eq_ignore_ascii_case("pastespecial") {
                    return self.paste_special(range, args).map(Some);
                }
                if name.eq_ignore_ascii_case("sort") {
                    return self.sort_range(range, args).map(|()| Some(Value::Empty));
                }
                if name.eq_ignore_ascii_case("insert") {
                    return self.insert_range(range, args).map(Some);
                }
                if name.eq_ignore_ascii_case("delete") {
                    return self.delete_range(range, args).map(Some);
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
        if name.eq_ignore_ascii_case("worksheetfunction") {
            if !args.is_empty() {
                return Err("WorksheetFunction does not accept arguments".to_string());
            }
            return Ok(Some(self.object(HostObject::WorksheetFunction)));
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

    fn call_named(
        &mut self,
        receiver: Option<&ObjectRef>,
        name: &str,
        args: &[Value],
        argument_names: &[Option<String>],
    ) -> Result<Option<Value>, String> {
        let parameters = if name.eq_ignore_ascii_case("find") {
            Some(
                &[
                    "What",
                    "After",
                    "LookIn",
                    "LookAt",
                    "SearchOrder",
                    "SearchDirection",
                    "MatchCase",
                    "MatchByte",
                    "SearchFormat",
                ][..],
            )
        } else if name.eq_ignore_ascii_case("findnext") || name.eq_ignore_ascii_case("findprevious")
        {
            Some(&["After"][..])
        } else if name.eq_ignore_ascii_case("replace") {
            Some(
                &[
                    "What",
                    "Replacement",
                    "LookAt",
                    "SearchOrder",
                    "MatchCase",
                    "MatchByte",
                    "SearchFormat",
                    "ReplaceFormat",
                ][..],
            )
        } else if name.eq_ignore_ascii_case("offset") {
            Some(&["RowOffset", "ColumnOffset"][..])
        } else if name.eq_ignore_ascii_case("resize") {
            Some(&["RowSize", "ColumnSize"][..])
        } else if name.eq_ignore_ascii_case("address") {
            Some(
                &[
                    "RowAbsolute",
                    "ColumnAbsolute",
                    "ReferenceStyle",
                    "External",
                    "RelativeTo",
                ][..],
            )
        } else if name.eq_ignore_ascii_case("add") {
            Some(&["Before", "After", "Count"][..])
        } else if name.eq_ignore_ascii_case("sort") {
            Some(
                &[
                    "Key1",
                    "Order1",
                    "Key2",
                    "Type",
                    "Order2",
                    "Key3",
                    "Order3",
                    "Header",
                    "OrderCustom",
                    "MatchCase",
                    "Orientation",
                    "SortMethod",
                    "DataOption1",
                    "DataOption2",
                    "DataOption3",
                ][..],
            )
        } else if name.eq_ignore_ascii_case("copy") {
            // A worksheet is copied somewhere among the sheets; a range is
            // copied onto other cells.
            match receiver.is_some_and(|receiver| self.worksheet(receiver).is_some()) {
                true => Some(&["Before", "After"][..]),
                false => Some(&["Destination"][..]),
            }
        } else if name.eq_ignore_ascii_case("move") {
            Some(&["Before", "After"][..])
        } else if name.eq_ignore_ascii_case("pastespecial") {
            Some(&["Paste", "Operation", "SkipBlanks", "Transpose"][..])
        } else if name.eq_ignore_ascii_case("merge") {
            Some(&["Across"][..])
        } else if name.eq_ignore_ascii_case("borders") {
            Some(&["Index"][..])
        } else if name.eq_ignore_ascii_case("range") {
            Some(&["Cell1", "Cell2"][..])
        } else if name.eq_ignore_ascii_case("cells") {
            Some(&["RowIndex", "ColumnIndex"][..])
        } else if name.eq_ignore_ascii_case("msgbox") {
            Some(&["Prompt", "Buttons", "Title", "HelpFile", "Context"][..])
        } else {
            None
        };
        if let Some(parameters) = parameters {
            let args = normalize_named_arguments(args, argument_names, parameters, name)?;
            return self.call(receiver, name, &args);
        }
        self.call(receiver, name, args)
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
            if name.eq_ignore_ascii_case("worksheetfunction") {
                return Ok(Some(self.object(HostObject::WorksheetFunction)));
            }
            if name.eq_ignore_ascii_case("screenupdating") {
                return Ok(Some(Value::Boolean(self.screen_updating)));
            }
            if name.eq_ignore_ascii_case("enableevents") {
                return Ok(Some(Value::Boolean(self.enable_events)));
            }
            if name.eq_ignore_ascii_case("displayalerts") {
                return Ok(Some(Value::Boolean(self.display_alerts)));
            }
            if name.eq_ignore_ascii_case("calculation") {
                return Ok(Some(Value::Integer(self.calculation)));
            }
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
            if name.eq_ignore_ascii_case("add") {
                return self.add_worksheet(&[]).map(Some);
            }
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
        if let Some(sheet) = self.worksheet(receiver) {
            if name.eq_ignore_ascii_case("name") {
                let Value::String(renamed) = &value else {
                    return Err("a worksheet name must be a String".to_string());
                };
                self.rename_worksheet(sheet, renamed)?;
                return Ok(true);
            }
        }
        if self.is_application(receiver) {
            if name.eq_ignore_ascii_case("screenupdating") {
                self.screen_updating = style_boolean(&value, "Application.ScreenUpdating")?;
                return Ok(true);
            }
            if name.eq_ignore_ascii_case("enableevents") {
                self.enable_events = style_boolean(&value, "Application.EnableEvents")?;
                return Ok(true);
            }
            if name.eq_ignore_ascii_case("displayalerts") {
                self.display_alerts = style_boolean(&value, "Application.DisplayAlerts")?;
                return Ok(true);
            }
            if name.eq_ignore_ascii_case("calculation") {
                self.calculation = application_calculation(&value)?;
                return Ok(true);
            }
            return Ok(false);
        }
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

fn application_calculation(value: &Value) -> Result<i64, String> {
    let value = match value {
        Value::Integer(value) => *value,
        Value::Double(value) if value.is_finite() && value.fract() == 0.0 => *value as i64,
        _ => {
            return Err("Application.Calculation must be an Excel calculation constant".to_string())
        }
    };
    match value {
        -4105 | -4135 | 2 => Ok(value),
        _ => Err(format!(
            "unsupported Application.Calculation constant: {value}"
        )),
    }
}

fn find_integer_argument(value: Option<&Value>, default: i64, label: &str) -> Result<i64, String> {
    match value {
        None | Some(Value::Missing) => Ok(default),
        Some(Value::Integer(value)) => Ok(*value),
        Some(Value::Double(value)) if value.is_finite() && value.fract() == 0.0 => {
            Ok(*value as i64)
        }
        _ => Err(format!("Range.Find {label} must be an Excel constant")),
    }
}

fn normalize_named_arguments(
    args: &[Value],
    argument_names: &[Option<String>],
    parameters: &[&str],
    method: &str,
) -> Result<Vec<Value>, String> {
    if args.len() != argument_names.len() {
        return Err(format!("Range.{method} received invalid argument metadata"));
    }
    if argument_names.iter().all(Option::is_none) {
        return Ok(args.to_vec());
    }
    let mut normalized = vec![Value::Missing; parameters.len()];
    let mut assigned = vec![false; parameters.len()];
    let mut positional_index = 0;
    let mut highest_index = None;
    for (value, argument_name) in args.iter().zip(argument_names) {
        let index = if let Some(argument_name) = argument_name {
            parameters
                .iter()
                .position(|parameter| parameter.eq_ignore_ascii_case(argument_name))
                .ok_or_else(|| format!("Range.{method} has no argument named {argument_name}"))?
        } else {
            while positional_index < assigned.len() && assigned[positional_index] {
                positional_index += 1;
            }
            if positional_index >= parameters.len() {
                return Err(format!("Range.{method} received too many arguments"));
            }
            let index = positional_index;
            positional_index += 1;
            index
        };
        if assigned[index] {
            return Err(format!(
                "Range.{method} argument {} was supplied more than once",
                parameters[index]
            ));
        }
        normalized[index] = value.clone();
        assigned[index] = true;
        highest_index = Some(highest_index.map_or(index, |highest: usize| highest.max(index)));
    }
    normalized.truncate(highest_index.map_or(0, |index| index + 1));
    Ok(normalized)
}

fn find_boolean_argument(
    value: Option<&Value>,
    default: bool,
    label: &str,
) -> Result<bool, String> {
    match value {
        None | Some(Value::Missing) => Ok(default),
        Some(value) => style_boolean(value, &format!("Range.Find {label}")),
    }
}

/// The three value classes Excel's lookups compare within. Values of different
/// classes never compare equal, so `Match` on a numeric column never matches a
/// text needle.
enum LookupKey<'a> {
    Number(f64),
    Text(&'a str),
    Boolean(bool),
}

fn lookup_key(value: &Value) -> Option<LookupKey<'_>> {
    match value {
        Value::Boolean(value) => Some(LookupKey::Boolean(*value)),
        Value::Integer(value) => Some(LookupKey::Number(*value as f64)),
        Value::Double(value) if value.is_finite() => Some(LookupKey::Number(*value)),
        Value::String(value) => Some(LookupKey::Text(value)),
        _ => None,
    }
}

fn lookup_compare(cell: &Value, needle: &Value) -> Option<Ordering> {
    match (lookup_key(cell)?, lookup_key(needle)?) {
        (LookupKey::Number(cell), LookupKey::Number(needle)) => cell.partial_cmp(&needle),
        (LookupKey::Text(cell), LookupKey::Text(needle)) => {
            Some(cell.to_lowercase().cmp(&needle.to_lowercase()))
        }
        (LookupKey::Boolean(cell), LookupKey::Boolean(needle)) => Some(cell.cmp(&needle)),
        _ => None,
    }
}

/// Exact lookup matching. A text needle is a wildcard pattern; every other
/// comparison is a plain equality within the needle's value class.
fn lookup_exact_matches(cell: &Value, needle: &Value) -> bool {
    match (lookup_key(cell), lookup_key(needle)) {
        (Some(LookupKey::Text(cell)), Some(LookupKey::Text(needle))) => {
            wildcard_matches(cell, needle)
        }
        _ => lookup_compare(cell, needle) == Some(Ordering::Equal),
    }
}

enum WildcardToken {
    AnyRun,
    AnyCharacter,
    Literal(char),
}

/// Excel's lookup wildcards: `*` runs, `?` single characters, and `~` escaping
/// the character after it. A trailing `~` contributes nothing at all.
fn wildcard_tokens(pattern: &str) -> Vec<WildcardToken> {
    let mut tokens = Vec::new();
    let mut characters = pattern.to_lowercase().chars().collect::<Vec<_>>().into_iter();
    while let Some(character) = characters.next() {
        match character {
            '~' => {
                if let Some(escaped) = characters.next() {
                    tokens.push(WildcardToken::Literal(escaped));
                }
            }
            '*' => tokens.push(WildcardToken::AnyRun),
            '?' => tokens.push(WildcardToken::AnyCharacter),
            character => tokens.push(WildcardToken::Literal(character)),
        }
    }
    tokens
}

fn wildcard_matches(candidate: &str, pattern: &str) -> bool {
    let candidate = candidate.to_lowercase().chars().collect::<Vec<_>>();
    let tokens = wildcard_tokens(pattern);
    let mut candidate_index = 0;
    let mut token_index = 0;
    let mut resume: Option<(usize, usize)> = None;
    while candidate_index < candidate.len() {
        let matched = match tokens.get(token_index) {
            Some(WildcardToken::AnyRun) => {
                token_index += 1;
                resume = Some((token_index, candidate_index));
                continue;
            }
            Some(WildcardToken::AnyCharacter) => true,
            Some(WildcardToken::Literal(expected)) => *expected == candidate[candidate_index],
            None => false,
        };
        if matched {
            token_index += 1;
            candidate_index += 1;
        } else if let Some((resume_token, resume_candidate)) = resume {
            token_index = resume_token;
            candidate_index = resume_candidate + 1;
            resume = Some((resume_token, candidate_index));
        } else {
            return false;
        }
    }
    tokens[token_index..]
        .iter()
        .all(|token| matches!(token, WildcardToken::AnyRun))
}

/// Excel's search over a table it assumes is sorted, derived from Excel 16 COM
/// measurements. It is a plain binary search that stops on the first equal
/// probe, then walks to the far end of that run of equal values: forward for an
/// ascending table, backward for a descending one. A descending search also
/// rejects a needle above the leading value, the ordering's assumed maximum.
///
/// Measured against Excel over 5,040 differential cases: the ascending search
/// agrees on every ordering, and the descending one agrees on every sorted
/// table. Only a descending search of a shuffled table can differ, which is
/// input the sorted contract does not define an answer for.
fn sorted_lookup_position(
    count: usize,
    descending: bool,
    value_at: impl Fn(usize) -> Value,
    needle: &Value,
) -> Option<usize> {
    if count == 0 {
        return None;
    }
    if descending {
        match lookup_compare(&value_at(0), needle) {
            Some(Ordering::Less) => return None,
            Some(Ordering::Equal) => return Some(1),
            _ => {}
        }
    }
    let mut low = 1;
    let mut high = count;
    let mut found = None;
    while low <= high {
        let middle = (low + high) / 2;
        match lookup_compare(&value_at(middle - 1), needle) {
            Some(Ordering::Equal) => {
                found = Some(middle);
                break;
            }
            Some(order) if (order == Ordering::Less) != descending => {
                found = Some(middle);
                low = middle + 1;
            }
            _ => high = middle - 1,
        }
    }
    let mut position = found?;
    if descending {
        while position > 1 && lookup_compare(&value_at(position - 2), needle) == Some(Ordering::Equal)
        {
            position -= 1;
        }
    } else {
        while position < count && lookup_compare(&value_at(position), needle) == Some(Ordering::Equal)
        {
            position += 1;
        }
    }
    Some(position)
}

fn lookup_index_argument(value: &Value, name: &str) -> Result<usize, String> {
    let number = match value {
        Value::Integer(value) => *value as f64,
        Value::Double(value) if value.is_finite() => *value,
        _ => return Err(format!("WorksheetFunction.{name} index must be numeric")),
    };
    let number = number.trunc();
    if !(1.0..=u32::MAX as f64).contains(&number) {
        return Err(format!(
            "WorksheetFunction.{name} index must be 1 or greater"
        ));
    }
    Ok(number as usize)
}

#[derive(Debug, Clone, Copy, PartialEq)]
enum ConditionalKind {
    Count,
    Sum,
    Average,
}

#[derive(Debug, Clone, Copy, PartialEq)]
enum CriteriaOperator {
    Equal,
    NotEqual,
    Less,
    LessOrEqual,
    Greater,
    GreaterOrEqual,
}

struct Criteria {
    operator: CriteriaOperator,
    /// `Value::Empty` for the bare `"="` and `"<>"` forms, which ask whether
    /// the cell is blank rather than comparing it to anything.
    operand: Value,
}

impl Criteria {
    fn matches(&self, cell: &Value) -> bool {
        let blank = matches!(cell, Value::Empty | Value::Missing);
        if matches!(self.operand, Value::Empty) {
            return match self.operator {
                CriteriaOperator::Equal => blank,
                CriteriaOperator::NotEqual => !blank,
                _ => false,
            };
        }
        match self.operator {
            CriteriaOperator::Equal => criteria_equal(cell, &self.operand, true),
            CriteriaOperator::NotEqual => !criteria_equal(cell, &self.operand, false),
            operator => match lookup_compare(cell, &self.operand) {
                Some(Ordering::Less) => matches!(
                    operator,
                    CriteriaOperator::Less | CriteriaOperator::LessOrEqual
                ),
                Some(Ordering::Equal) => matches!(
                    operator,
                    CriteriaOperator::LessOrEqual | CriteriaOperator::GreaterOrEqual
                ),
                Some(Ordering::Greater) => matches!(
                    operator,
                    CriteriaOperator::Greater | CriteriaOperator::GreaterOrEqual
                ),
                None => false,
            },
        }
    }
}

/// Splits a criteria argument into an operator and the value it compares
/// against. Only text carries an operator; any other value is compared for
/// equality as it stands.
fn parse_criteria(value: &Value) -> Criteria {
    let Value::String(text) = value else {
        return Criteria {
            operator: CriteriaOperator::Equal,
            operand: value.clone(),
        };
    };
    let text = text.trim();
    let (operator, rest) = if let Some(rest) = text.strip_prefix("<>") {
        (CriteriaOperator::NotEqual, rest)
    } else if let Some(rest) = text.strip_prefix(">=") {
        (CriteriaOperator::GreaterOrEqual, rest)
    } else if let Some(rest) = text.strip_prefix("<=") {
        (CriteriaOperator::LessOrEqual, rest)
    } else if let Some(rest) = text.strip_prefix('>') {
        (CriteriaOperator::Greater, rest)
    } else if let Some(rest) = text.strip_prefix('<') {
        (CriteriaOperator::Less, rest)
    } else if let Some(rest) = text.strip_prefix('=') {
        (CriteriaOperator::Equal, rest)
    } else {
        (CriteriaOperator::Equal, text)
    };
    Criteria {
        operator,
        operand: criteria_operand(rest.trim()),
    }
}

fn criteria_operand(text: &str) -> Value {
    if text.is_empty() {
        return Value::Empty;
    }
    if let Ok(number) = text.parse::<f64>() {
        if number.is_finite() {
            return Value::Double(number);
        }
    }
    if text.eq_ignore_ascii_case("true") {
        return Value::Boolean(true);
    }
    if text.eq_ignore_ascii_case("false") {
        return Value::Boolean(false);
    }
    Value::String(text.to_string())
}

/// Equality reads a cell holding text that spells a number as that number, so
/// `CountIf` over a column mixing `20` and `"20"` counts both. No other
/// operator coerces: `"<>20"` and `">=20"` see the text cell as text.
fn criteria_equal(cell: &Value, operand: &Value, coercing: bool) -> bool {
    if lookup_exact_matches(cell, operand) {
        return true;
    }
    if !coercing {
        return false;
    }
    match (cell, operand) {
        (Value::String(cell), Value::Double(operand)) => cell
            .trim()
            .parse::<f64>()
            .is_ok_and(|cell| cell == *operand),
        _ => false,
    }
}

fn criteria_number(value: &Value) -> Option<f64> {
    match value {
        Value::Integer(value) => Some(*value as f64),
        Value::Double(value) if value.is_finite() => Some(*value),
        _ => None,
    }
}

fn conditional_result(
    name: &str,
    kind: ConditionalKind,
    matches: usize,
    numbers: usize,
    total: f64,
) -> Result<Value, String> {
    match kind {
        ConditionalKind::Count => Ok(Value::Integer(matches as i64)),
        ConditionalKind::Sum => Ok(numeric_result(total)),
        ConditionalKind::Average if numbers == 0 => Err(format!(
            "WorksheetFunction.{name} has no matching numeric values"
        )),
        ConditionalKind::Average => Ok(numeric_result(total / numbers as f64)),
    }
}

fn numeric_result(value: f64) -> Value {
    if value.fract() == 0.0 && value >= i64::MIN as f64 && value <= i64::MAX as f64 {
        Value::Integer(value as i64)
    } else {
        Value::Double(value)
    }
}

fn worksheet_number(value: &Value, name: &str) -> Result<f64, String> {
    match value {
        Value::Integer(value) => Ok(*value as f64),
        Value::Double(value) if value.is_finite() => Ok(*value),
        Value::Boolean(value) => Ok(f64::from(*value)),
        Value::String(value) => value
            .trim()
            .parse::<f64>()
            .map_err(|_| format!("WorksheetFunction.{name} expects a number")),
        _ => Err(format!("WorksheetFunction.{name} expects a number")),
    }
}

fn sort_number(value: &Value, label: &str) -> Result<i64, String> {
    match value {
        Value::Integer(value) => Ok(*value),
        Value::Double(value) if value.is_finite() => Ok(value.trunc() as i64),
        _ => Err(format!("Range.Sort {label} must be numeric")),
    }
}

/// Excel's sort order: numbers first, then text, then Booleans, with blanks
/// last however the sort runs. Text ignores case.
fn sort_rank(value: &Value) -> u8 {
    match value {
        Value::Integer(_) | Value::Double(_) => 0,
        Value::String(_) => 1,
        Value::Boolean(_) => 2,
        _ => 3,
    }
}

fn sort_compare(left: &Value, right: &Value, descending: bool) -> Ordering {
    let (left_rank, right_rank) = (sort_rank(left), sort_rank(right));
    if left_rank == 3 || right_rank == 3 {
        // A blank sinks to the bottom whichever way the rest is going.
        return left_rank.cmp(&right_rank);
    }
    let ordering = if left_rank != right_rank {
        left_rank.cmp(&right_rank)
    } else {
        match (left, right) {
            (Value::String(left), Value::String(right)) => {
                left.to_lowercase().cmp(&right.to_lowercase())
            }
            (Value::Boolean(left), Value::Boolean(right)) => left.cmp(right),
            _ => match (number_of(left), number_of(right)) {
                (Some(left), Some(right)) => {
                    left.partial_cmp(&right).unwrap_or(Ordering::Equal)
                }
                _ => Ordering::Equal,
            },
        }
    };
    if descending {
        ordering.reverse()
    } else {
        ordering
    }
}

fn number_of(value: &Value) -> Option<f64> {
    match value {
        Value::Integer(value) => Some(*value as f64),
        Value::Double(value) if value.is_finite() => Some(*value),
        _ => None,
    }
}

fn index_argument(value: &Value, label: &str) -> Result<usize, String> {
    let number = match value {
        Value::Integer(value) => *value as f64,
        Value::Double(value) if value.is_finite() => *value,
        _ => return Err(format!("WorksheetFunction.Index {label} must be numeric")),
    };
    let number = number.trunc();
    if !(0.0..=u32::MAX as f64).contains(&number) {
        return Err(format!(
            "WorksheetFunction.Index {label} must be zero or greater"
        ));
    }
    Ok(number as usize)
}

/// Resolves `Index` arguments to a one-based row and column, where zero selects
/// a whole row or column. Omitting the column walks the length of a table that
/// has only one row or column; for anything wider Excel accepts the shorthand
/// from an array, taking a whole row, but rejects it from a cell reference.
fn index_selection(
    rows: usize,
    columns: usize,
    row: usize,
    column: Option<usize>,
    reference: bool,
) -> Result<(usize, usize), String> {
    let (row, column) = match column {
        Some(column) => (row, column),
        None if rows == 1 => (1, row),
        None if columns == 1 => (row, 1),
        None if reference => {
            return Err(
                "WorksheetFunction.Index needs a column for a two-dimensional reference".to_string(),
            )
        }
        None => (row, 0),
    };
    if row > rows || column > columns {
        return Err(format!(
            "WorksheetFunction.Index {row},{column} is outside a {rows}-by-{columns} array"
        ));
    }
    Ok((row, column))
}

fn lookup_boolean_argument(value: Option<&Value>, name: &str) -> Result<bool, String> {
    match value {
        None | Some(Value::Missing) => Ok(true),
        Some(Value::Boolean(value)) => Ok(*value),
        Some(Value::Integer(value)) => Ok(*value != 0),
        Some(Value::Double(value)) if value.is_finite() => Ok(*value != 0.0),
        Some(_) => Err(format!(
            "WorksheetFunction.{name} range lookup flag must be Boolean"
        )),
    }
}

fn match_type_argument(value: Option<&Value>) -> Result<i64, String> {
    let number = match value {
        None | Some(Value::Missing) => return Ok(1),
        Some(Value::Integer(value)) => *value as f64,
        Some(Value::Double(value)) if value.is_finite() => *value,
        Some(_) => return Err("WorksheetFunction.Match type must be numeric".to_string()),
    };
    let number = number.trunc();
    Ok(if number > 0.0 {
        1
    } else if number < 0.0 {
        -1
    } else {
        0
    })
}

fn find_value_text(value: &Value) -> String {
    match value {
        Value::Empty => String::new(),
        Value::Boolean(value) => if *value { "TRUE" } else { "FALSE" }.to_string(),
        Value::Integer(value) => value.to_string(),
        Value::Double(value) => value.to_string(),
        Value::Error(value) => format!("Error {value}"),
        Value::String(value) => value.clone(),
        Value::Null => "Null".to_string(),
        Value::Missing => String::new(),
        Value::Nothing => "Nothing".to_string(),
        Value::Array(_) => "<Array>".to_string(),
        Value::Object(value) => format!("<{}>", value.kind),
    }
}

fn find_text_matches(candidate: &str, needle: &str, whole: bool, match_case: bool) -> bool {
    if match_case {
        if whole {
            candidate == needle
        } else {
            candidate.contains(needle)
        }
    } else {
        let candidate = candidate.to_lowercase();
        let needle = needle.to_lowercase();
        if whole {
            candidate == needle
        } else {
            candidate.contains(&needle)
        }
    }
}

fn replace_matching_text(
    candidate: &str,
    needle: &str,
    replacement: &str,
    whole: bool,
    match_case: bool,
) -> Option<String> {
    if !find_text_matches(candidate, needle, whole, match_case) {
        return None;
    }
    if whole {
        return Some(replacement.to_string());
    }
    if needle.is_empty() {
        return None;
    }
    if match_case {
        return Some(candidate.replace(needle, replacement));
    }

    let mut result = String::with_capacity(candidate.len());
    let mut offset = 0;
    let mut replaced_any = false;
    while offset < candidate.len() {
        let Some(relative) = candidate[offset..].char_indices().find_map(|(index, _)| {
            let start = offset + index;
            let end = start.checked_add(needle.len())?;
            (candidate.is_char_boundary(end) && candidate[start..end].eq_ignore_ascii_case(needle))
                .then_some(index)
        }) else {
            result.push_str(&candidate[offset..]);
            break;
        };
        let start = offset + relative;
        result.push_str(&candidate[offset..start]);
        result.push_str(replacement);
        offset = start + needle.len();
        replaced_any = true;
    }
    replaced_any.then_some(result)
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
        "xlpasteall" => -4104,
        "xlpastevalues" => -4163,
        "xlpasteformats" => -4122,
        "xlascending" => 1,
        "xldescending" => 2,
        "xlyes" => 1,
        "xlno" => 2,
        "xlguess" => 0,
        "xltoptobottom" => 1,
        "xllefttoright" => 2,
        // The shift an Insert or Delete takes, which share these numbers.
        "xlshiftdown" => -4121,
        "xlshifttoright" => -4161,
        "xlshiftup" => -4162,
        "xlshifttoleft" => -4159,
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
        "xlcalculationautomatic" => -4105,
        "xlcalculationmanual" => -4135,
        "xlcalculationsemiautomatic" => 2,
        "xlformulas" => -4123,
        "xlvalues" => -4163,
        "xlwhole" => 1,
        "xlpart" => 2,
        "xlbyrows" => 1,
        "xlbycolumns" => 2,
        "xlnext" => 1,
        "xlprevious" => 2,
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

fn optional_positive_index(value: &Value, default: u32, label: &str) -> Result<u32, String> {
    if matches!(value, Value::Missing) {
        Ok(default)
    } else {
        positive_index(value, label)
    }
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

fn optional_integer_offset(value: &Value, default: i64, label: &str) -> Result<i64, String> {
    if matches!(value, Value::Missing) {
        Ok(default)
    } else {
        integer_offset(value, label)
    }
}

fn range_address_from_args(range: CellRange, args: &[Value]) -> Result<String, String> {
    let (row_absolute, column_absolute) = match args {
        [] => (true, true),
        [row_absolute] => (
            optional_boolean_argument(row_absolute, true, "row absolute")?,
            true,
        ),
        [row_absolute, column_absolute] => (
            optional_boolean_argument(row_absolute, true, "row absolute")?,
            optional_boolean_argument(column_absolute, true, "column absolute")?,
        ),
        _ => return Err("Range.Address supports up to two arguments".to_string()),
    };
    Ok(format_range_address(range, row_absolute, column_absolute))
}

fn optional_boolean_argument(value: &Value, default: bool, label: &str) -> Result<bool, String> {
    if matches!(value, Value::Missing) {
        Ok(default)
    } else {
        boolean_argument(value, label)
    }
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
        // Every number in a cell is a Double, whole or not, which is the type
        // VarType and TypeName report for one.
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
    fn vba_binds_sparse_named_range_dimensions() {
        let mut workbook = workbook();
        let module = parse_module(
            "Public Sub NamedDimensions()\n\
               Set horizontal = Range(\"B2:C3\").Offset(ColumnOffset:=2).Resize(ColumnSize:=3)\n\
               Set vertical = Range(\"B2:C3\").Offset(RowOffset:=2).Resize(RowSize:=1)\n\
               horizontal.Value = 5\n\
               vertical.Value = 7\n\
               Debug.Print horizontal.Address(False, False), vertical.Address(False, False)\n\
             End Sub\n",
        )
        .unwrap();
        let debug_output = {
            let mut host = WorkbookHost::new(&mut workbook, 0).unwrap();
            execute_with_host(&module, "NamedDimensions", vec![], &mut host).unwrap();
            host.take_debug_output()
        };

        assert_eq!(debug_output, vec!["D2:F3\tB4:C4".to_string()]);
        assert!(matches!(
            workbook.sheets[0]
                .rows
                .iter()
                .find(|row| row.index == 2)
                .unwrap()
                .cells
                .iter()
                .find(|cell| cell.col == 3)
                .unwrap()
                .value,
            CellValue::Number(5.0)
        ));
        assert!(matches!(
            workbook.sheets[0]
                .rows
                .iter()
                .find(|row| row.index == 4)
                .unwrap()
                .cells
                .iter()
                .find(|cell| cell.col == 1)
                .unwrap()
                .value,
            CellValue::Number(7.0)
        ));
    }

    #[test]
    fn vba_binds_named_cell_copy_address_and_merge_arguments() {
        let mut workbook = workbook();
        let module = parse_module(
            "Public Sub NamedRangeCalls()\n\
               Range(Cell1:=\"A1\").Value = 11\n\
               Set target = Cells(RowIndex:=3, ColumnIndex:=2)\n\
               Range(Cell1:=\"A1\").Copy Destination:=target\n\
               Range(Cell1:=\"D1\", Cell2:=\"E2\").Merge Across:=False\n\
               Range(Cell1:=\"D1:E2\").Borders(Index:=xlEdgeBottom).LineStyle = xlContinuous\n\
               MsgBox Prompt:=\"finished\", Title:=\"Named calls\"\n\
               Debug.Print target.Address(ColumnAbsolute:=False), target.Address(RowAbsolute:=False), target.Value, Range(\"D1:E2\").MergeCells, Range(\"D1:E2\").Borders(xlEdgeBottom).LineStyle\n\
             End Sub\n",
        )
        .unwrap();
        let (debug_output, messages) = {
            let mut host = WorkbookHost::new(&mut workbook, 0).unwrap();
            execute_with_host(&module, "NamedRangeCalls", vec![], &mut host).unwrap();
            (host.take_debug_output(), host.take_messages())
        };

        assert_eq!(debug_output, vec!["B$3\t$B3\t11\tTrue\t1".to_string()]);
        assert_eq!(messages.len(), 1);
        assert_eq!(messages[0].prompt, "finished");
        assert_eq!(messages[0].title, "Named calls");
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
    fn vba_aggregates_ranges_arrays_and_scalars_with_worksheet_functions() {
        let mut workbook = workbook();
        let module = parse_module(
            "Public Sub AggregateCells()\n\
               Dim values(1 To 2) As Variant\n\
               Range(\"A1\").Value = 10\n\
               Range(\"A2\").Value = 20\n\
               Range(\"A3\").Value = \"text\"\n\
               Range(\"B1\").Value = 5\n\
               values(1) = 2\n\
               values(2) = 3\n\
               Debug.Print Application.WorksheetFunction.Sum(Range(\"A1:B1\"), values)\n\
               Debug.Print WorksheetFunction.Average(Range(\"A1:A3\")), WorksheetFunction.Min(Range(\"A1:B2\")), WorksheetFunction.Max(Range(\"A1:B2\"))\n\
               Debug.Print WorksheetFunction.Count(Range(\"A1:A3\")), WorksheetFunction.CountA(Range(\"A1:A3\")), WorksheetFunction.Sum(0.5, 1)\n\
             End Sub\n",
        )
        .unwrap();
        let debug_output = {
            let mut host = WorkbookHost::new(&mut workbook, 0).unwrap();
            execute_with_host(&module, "AggregateCells", vec![], &mut host).unwrap();
            host.take_debug_output()
        };

        assert_eq!(
            debug_output,
            vec![
                "20".to_string(),
                "15\t5\t20".to_string(),
                "2\t3\t1.5".to_string(),
            ]
        );
    }

    /// Positions Excel 16 returned for `Match` over these tables. The ascending
    /// and shuffled tables share the same answers because Excel's search does
    /// not check that its input is sorted.
    #[test]
    fn worksheet_match_reproduces_excel_binary_search_positions() {
        let ascending = [10.0, 20.0, 30.0, 40.0, 50.0];
        let descending = [50.0, 40.0, 30.0, 20.0, 10.0];
        let shuffled = [10.0, 50.0, 20.0, 40.0, 30.0];
        let value_at = |table: &[f64; 5]| {
            let table = *table;
            move |index: usize| Value::Double(table[index])
        };

        for (needle, expected) in [
            (5.0, None),
            (10.0, Some(1)),
            (15.0, Some(1)),
            (20.0, Some(2)),
            (25.0, Some(2)),
            (30.0, Some(3)),
            (35.0, Some(3)),
            (40.0, Some(4)),
            (45.0, Some(4)),
            (50.0, Some(5)),
            (55.0, Some(5)),
        ] {
            assert_eq!(
                sorted_lookup_position(5, false, value_at(&ascending), &Value::Double(needle)),
                expected,
                "ascending needle {needle}"
            );
        }

        for (needle, expected) in [
            (25.0, None),
            (30.0, Some(3)),
            (35.0, Some(5)),
            (55.0, Some(5)),
        ] {
            assert_eq!(
                sorted_lookup_position(5, false, value_at(&descending), &Value::Double(needle)),
                expected,
                "descending needle {needle} searched as ascending"
            );
        }

        for (needle, expected) in [
            (5.0, None),
            (10.0, Some(1)),
            (20.0, Some(3)),
            (25.0, Some(3)),
            (40.0, Some(4)),
            (45.0, Some(5)),
        ] {
            assert_eq!(
                sorted_lookup_position(5, false, value_at(&shuffled), &Value::Double(needle)),
                expected,
                "shuffled needle {needle}"
            );
        }

        for (needle, expected) in [
            (5.0, Some(5)),
            (10.0, Some(5)),
            (15.0, Some(4)),
            (25.0, Some(3)),
            (45.0, Some(1)),
            (50.0, Some(1)),
            (55.0, None),
        ] {
            assert_eq!(
                sorted_lookup_position(5, true, value_at(&descending), &Value::Double(needle)),
                expected,
                "descending needle {needle}"
            );
        }

        // A descending search rejects anything above the leading value, so an
        // ascending table only answers below its own head.
        for (needle, expected) in [(5.0, Some(5)), (10.0, Some(1)), (15.0, None), (50.0, None)] {
            assert_eq!(
                sorted_lookup_position(5, true, value_at(&ascending), &Value::Double(needle)),
                expected,
                "ascending needle {needle} searched as descending"
            );
        }
    }

    /// Runs of equal values resolve to opposite ends: the last of the run when
    /// ascending, the first when descending.
    #[test]
    fn worksheet_match_walks_to_the_far_end_of_an_equal_run() {
        let ascending = [30.0, 5.0, 5.0, 40.0];
        let descending = [55.0, 40.0, 40.0, 15.0, 15.0, 5.0];
        assert_eq!(
            sorted_lookup_position(
                4,
                false,
                |index| Value::Double(ascending[index]),
                &Value::Double(5.0)
            ),
            Some(3)
        );
        assert_eq!(
            sorted_lookup_position(
                6,
                true,
                |index| Value::Double(descending[index]),
                &Value::Double(40.0)
            ),
            Some(2)
        );
    }

    #[test]
    fn lookup_wildcards_follow_excel_tilde_escaping() {
        assert!(wildcard_matches("axxb", "a*b"));
        assert!(wildcard_matches("a*b", "a~*b"));
        assert!(!wildcard_matches("axxb", "a~*b"));
        assert!(wildcard_matches("a*b", "A*B"));
        assert!(!wildcard_matches("ab", "a?b"));
        assert!(wildcard_matches("axb", "a?b"));
        assert!(wildcard_matches("~ab", "~~ab"));
        assert!(wildcard_matches("ab", "~ab"));
        assert!(wildcard_matches("ab", "ab~"));
        assert!(wildcard_matches("banana", "b*"));
        assert!(!wildcard_matches("cherry", "b*"));
    }

    #[test]
    fn lookups_never_match_across_value_classes() {
        assert!(lookup_exact_matches(
            &Value::Boolean(true),
            &Value::Boolean(true)
        ));
        assert!(!lookup_exact_matches(
            &Value::Boolean(true),
            &Value::Integer(1)
        ));
        assert!(!lookup_exact_matches(
            &Value::String("TRUE".to_string()),
            &Value::Boolean(true)
        ));
        assert!(!lookup_exact_matches(&Value::Empty, &Value::Integer(0)));
        assert!(lookup_exact_matches(
            &Value::String("Apple".to_string()),
            &Value::String("APPLE".to_string())
        ));
    }

    #[test]
    fn vba_looks_values_up_in_spreadsheet_tables() {
        let mut workbook = workbook();
        let module = parse_module(
            "Public Sub LookUpCells()\n\
               Range(\"A1\").Value = 10\n\
               Range(\"A2\").Value = 20\n\
               Range(\"A3\").Value = 30\n\
               Range(\"B1\").Value = \"ten\"\n\
               Range(\"B2\").Value = \"twenty\"\n\
               Range(\"B3\").Value = \"thirty\"\n\
               Debug.Print WorksheetFunction.VLookup(20, Range(\"A1:B3\"), 2, False)\n\
               Debug.Print WorksheetFunction.VLookup(25, Range(\"A1:B3\"), 2)\n\
               Debug.Print WorksheetFunction.VLookup(100, Range(\"A1:B3\"), 2, True)\n\
               Debug.Print WorksheetFunction.Match(20, Range(\"A1:A3\"), 0), WorksheetFunction.Match(25, Range(\"A1:A3\"))\n\
               Debug.Print Application.WorksheetFunction.HLookup(\"twenty\", Range(\"B2:B3\"), 2, False)\n\
             End Sub\n",
        )
        .unwrap();
        let debug_output = {
            let mut host = WorkbookHost::new(&mut workbook, 0).unwrap();
            execute_with_host(&module, "LookUpCells", vec![], &mut host).unwrap();
            host.take_debug_output()
        };

        assert_eq!(
            debug_output,
            vec![
                "twenty".to_string(),
                "twenty".to_string(),
                "thirty".to_string(),
                "2\t2".to_string(),
                "thirty".to_string(),
            ]
        );
    }

    #[test]
    fn vba_lookups_reject_what_excel_rejects() {
        let mut workbook = workbook();
        let module = parse_module(
            "Public Function LookUp(needle As Variant, column As Variant) As Variant\n\
               Range(\"A1\").Value = 10\n\
               Range(\"A2\").Value = 20\n\
               Range(\"A3\").Value = 30\n\
               LookUp = WorksheetFunction.VLookup(needle, Range(\"A1:B3\"), column, False)\n\
             End Function\n",
        )
        .unwrap();
        for (needle, column) in [
            (Value::Integer(25), Value::Integer(2)),
            (Value::Integer(10), Value::Integer(0)),
            (Value::Integer(10), Value::Integer(3)),
        ] {
            let mut host = WorkbookHost::new(&mut workbook, 0).unwrap();
            let error = execute_with_host(&module, "LookUp", vec![needle, column], &mut host)
                .expect_err("Excel raises rather than returning an error value");
            assert!(
                error.to_string().contains("VLookup"),
                "unexpected error: {error}"
            );
        }
    }

    /// A one-dimensional VBA array is a single row, so `Match` walks it while
    /// `VLookup` only ever sees its first column.
    #[test]
    fn vba_arrays_look_up_as_a_single_row() {
        let mut workbook = workbook();
        let module = parse_module(
            "Public Sub LookUpArray()\n\
               Dim values(1 To 3) As Variant\n\
               values(1) = 5\n\
               values(2) = 15\n\
               values(3) = 25\n\
               Debug.Print WorksheetFunction.Match(15, values, 0)\n\
               Debug.Print WorksheetFunction.HLookup(15, values, 1, False)\n\
               Debug.Print WorksheetFunction.VLookup(5, values, 1, False)\n\
             End Sub\n",
        )
        .unwrap();
        let debug_output = {
            let mut host = WorkbookHost::new(&mut workbook, 0).unwrap();
            execute_with_host(&module, "LookUpArray", vec![], &mut host).unwrap();
            host.take_debug_output()
        };

        assert_eq!(
            debug_output,
            vec!["2".to_string(), "15".to_string(), "5".to_string()]
        );
    }

    #[test]
    fn vba_indexes_cells_by_row_and_column() {
        let mut workbook = workbook();
        let module = parse_module(
            "Public Sub IndexCells()\n\
               Range(\"A1\").Value = 10\n\
               Range(\"A2\").Value = 20\n\
               Range(\"A3\").Value = 30\n\
               Range(\"B1\").Value = \"ten\"\n\
               Range(\"B2\").Value = \"twenty\"\n\
               Range(\"B3\").Value = \"thirty\"\n\
               Debug.Print WorksheetFunction.Index(Range(\"A1:B3\"), 2, 2).Value\n\
               Debug.Print WorksheetFunction.Index(Range(\"A1:A3\"), 2).Value\n\
               Debug.Print WorksheetFunction.Index(Range(\"B1:B3\"), WorksheetFunction.Match(20, Range(\"A1:A3\"), 0)).Value\n\
               Debug.Print WorksheetFunction.Index(Range(\"A1:B3\"), 2.9, 1).Value\n\
             End Sub\n",
        )
        .unwrap();
        let debug_output = {
            let mut host = WorkbookHost::new(&mut workbook, 0).unwrap();
            execute_with_host(&module, "IndexCells", vec![], &mut host).unwrap();
            host.take_debug_output()
        };

        assert_eq!(
            debug_output,
            vec![
                "twenty".to_string(),
                "20".to_string(),
                "twenty".to_string(),
                "20".to_string(),
            ]
        );
    }

    /// A zero row or column widens the answer to the whole column or row, and
    /// the reference it hands back is a real range: addressable, and countable
    /// by the aggregating functions.
    #[test]
    fn vba_indexes_whole_rows_and_columns_as_references() {
        let mut workbook = workbook();
        let module = parse_module(
            "Public Sub IndexReferences()\n\
               Range(\"A1\").Value = 10\n\
               Range(\"A2\").Value = 20\n\
               Range(\"A3\").Value = 30\n\
               Range(\"B1\").Value = \"ten\"\n\
               Range(\"B2\").Value = \"twenty\"\n\
               Range(\"B3\").Value = \"thirty\"\n\
               Debug.Print WorksheetFunction.Index(Range(\"A1:B3\"), 0, 2).Address(False, False)\n\
               Debug.Print WorksheetFunction.Index(Range(\"A1:B3\"), 2, 0).Address(False, False)\n\
               Debug.Print WorksheetFunction.Index(Range(\"A1:B3\"), 0, 0).Address(False, False)\n\
               Debug.Print WorksheetFunction.Sum(WorksheetFunction.Index(Range(\"A1:B3\"), 0, 1))\n\
               Debug.Print WorksheetFunction.Sum(WorksheetFunction.Index(Range(\"A1:B3\"), 2, 0))\n\
               Debug.Print WorksheetFunction.Count(WorksheetFunction.Index(Range(\"A1:B3\"), 0, 0))\n\
             End Sub\n",
        )
        .unwrap();
        let debug_output = {
            let mut host = WorkbookHost::new(&mut workbook, 0).unwrap();
            execute_with_host(&module, "IndexReferences", vec![], &mut host).unwrap();
            host.take_debug_output()
        };

        assert_eq!(
            debug_output,
            vec![
                "B1:B3".to_string(),
                "A2:B2".to_string(),
                "A1:B3".to_string(),
                "60".to_string(),
                "20".to_string(),
                "3".to_string(),
            ]
        );
    }

    #[test]
    fn vba_index_rejects_what_excel_rejects() {
        let mut workbook = workbook();
        let module = parse_module(
            "Public Function Pick(row As Variant, column As Variant) As Variant\n\
               Range(\"A1\").Value = 10\n\
               Range(\"B3\").Value = 30\n\
               If IsMissing(column) Then\n\
                 Set Pick = WorksheetFunction.Index(Range(\"A1:B3\"), row)\n\
               Else\n\
                 Set Pick = WorksheetFunction.Index(Range(\"A1:B3\"), row, column)\n\
               End If\n\
             End Function\n",
        )
        .unwrap();
        for args in [
            vec![Value::Integer(4), Value::Integer(1)],
            vec![Value::Integer(1), Value::Integer(3)],
            vec![Value::Integer(-1), Value::Integer(1)],
            vec![Value::Integer(1), Value::Missing],
        ] {
            let mut host = WorkbookHost::new(&mut workbook, 0).unwrap();
            let error = execute_with_host(&module, "Pick", args, &mut host)
                .expect_err("Excel raises rather than returning an error value");
            assert!(
                error.to_string().contains("Index"),
                "unexpected error: {error}"
            );
        }
    }

    /// Only an array accepts the one-argument shorthand on a two-dimensional
    /// table, where it takes a whole row; a cell reference rejects it.
    #[test]
    fn vba_indexes_arrays_by_value() {
        let mut workbook = workbook();
        let module = parse_module(
            "Public Sub IndexArray()\n\
               Dim flat(1 To 3) As Variant\n\
               Dim grid(1 To 2, 1 To 3) As Variant\n\
               flat(1) = 5\n\
               flat(2) = 15\n\
               flat(3) = 25\n\
               grid(1, 1) = 1\n\
               grid(1, 2) = 2\n\
               grid(1, 3) = 3\n\
               grid(2, 1) = 4\n\
               grid(2, 2) = 5\n\
               grid(2, 3) = 6\n\
               Debug.Print WorksheetFunction.Index(flat, 2), WorksheetFunction.Index(flat, 1, 2)\n\
               Debug.Print WorksheetFunction.Index(grid, 2, 3), WorksheetFunction.Index(grid, 1, 1)\n\
               Debug.Print WorksheetFunction.Sum(WorksheetFunction.Index(grid, 2))\n\
               Debug.Print WorksheetFunction.Sum(WorksheetFunction.Index(grid, 0, 2))\n\
             End Sub\n",
        )
        .unwrap();
        let debug_output = {
            let mut host = WorkbookHost::new(&mut workbook, 0).unwrap();
            execute_with_host(&module, "IndexArray", vec![], &mut host).unwrap();
            host.take_debug_output()
        };

        assert_eq!(
            debug_output,
            vec![
                "15\t15".to_string(),
                "6\t1".to_string(),
                "15".to_string(),
                "7".to_string(),
            ]
        );
    }

    #[test]
    fn vba_counts_and_sums_cells_that_meet_criteria() {
        let mut workbook = workbook();
        let module = parse_module(
            "Public Sub Conditionals()\n\
               Range(\"A1\").Value = 10\n\
               Range(\"A2\").Value = 20\n\
               Range(\"A3\").Value = 30\n\
               Range(\"A4\").Value = 20\n\
               Range(\"C1\").Value = 1\n\
               Range(\"C2\").Value = 2\n\
               Range(\"C3\").Value = 3\n\
               Range(\"C4\").Value = 4\n\
               Debug.Print WorksheetFunction.CountIf(Range(\"A1:A5\"), \">15\"), WorksheetFunction.CountIf(Range(\"A1:A5\"), 20)\n\
               Debug.Print WorksheetFunction.CountIf(Range(\"A1:A5\"), \"<>20\"), WorksheetFunction.CountIf(Range(\"A1:A5\"), \"<>\")\n\
               Debug.Print WorksheetFunction.SumIf(Range(\"A1:A4\"), \">15\"), WorksheetFunction.SumIf(Range(\"A1:A4\"), \">15\", Range(\"C1:C4\"))\n\
               Debug.Print WorksheetFunction.AverageIf(Range(\"A1:A4\"), \">15\"), WorksheetFunction.AverageIf(Range(\"A1:A4\"), \">15\", Range(\"C1:C4\"))\n\
             End Sub\n",
        )
        .unwrap();
        let debug_output = {
            let mut host = WorkbookHost::new(&mut workbook, 0).unwrap();
            execute_with_host(&module, "Conditionals", vec![], &mut host).unwrap();
            host.take_debug_output()
        };

        assert_eq!(
            debug_output,
            vec![
                "3\t2".to_string(),
                // The blank A5 is not equal to 20, but it is not non-blank.
                "3\t4".to_string(),
                "70\t9".to_string(),
                "23.333333333333332\t3".to_string(),
            ]
        );
    }

    /// `SumIf` anchors its aggregated range at the top-left corner and stretches
    /// it to the tested range's shape, so a single cell stands for a column.
    #[test]
    fn vba_stretches_a_sum_range_to_the_tested_shape() {
        let mut workbook = workbook();
        let module = parse_module(
            "Public Sub Stretch()\n\
               Range(\"A1\").Value = 10\n\
               Range(\"A2\").Value = 20\n\
               Range(\"A3\").Value = 30\n\
               Range(\"A4\").Value = 20\n\
               Range(\"C1\").Value = 1\n\
               Range(\"C2\").Value = 2\n\
               Range(\"C3\").Value = 3\n\
               Range(\"C4\").Value = 4\n\
               Range(\"C5\").Value = 5\n\
               Debug.Print WorksheetFunction.SumIf(Range(\"A1:A4\"), \">15\", Range(\"C1\"))\n\
               Debug.Print WorksheetFunction.SumIf(Range(\"A1:A4\"), \">15\", Range(\"C2:C3\"))\n\
             End Sub\n",
        )
        .unwrap();
        let debug_output = {
            let mut host = WorkbookHost::new(&mut workbook, 0).unwrap();
            execute_with_host(&module, "Stretch", vec![], &mut host).unwrap();
            host.take_debug_output()
        };

        assert_eq!(debug_output, vec!["9".to_string(), "12".to_string()]);
    }

    #[test]
    fn vba_counts_across_several_criteria_ranges() {
        let mut workbook = workbook();
        let module = parse_module(
            "Public Sub ManyCriteria()\n\
               Range(\"A1\").Value = 10\n\
               Range(\"A2\").Value = 20\n\
               Range(\"A3\").Value = 30\n\
               Range(\"A4\").Value = 20\n\
               Range(\"B1\").Value = \"Apple\"\n\
               Range(\"B2\").Value = \"banana\"\n\
               Range(\"B3\").Value = \"Cherry\"\n\
               Range(\"B4\").Value = \"apple\"\n\
               Range(\"C1\").Value = 1\n\
               Range(\"C2\").Value = 2\n\
               Range(\"C3\").Value = 3\n\
               Range(\"C4\").Value = 4\n\
               Debug.Print WorksheetFunction.CountIfs(Range(\"A1:A4\"), \">15\", Range(\"B1:B4\"), \"a*\")\n\
               Debug.Print WorksheetFunction.SumIfs(Range(\"C1:C4\"), Range(\"A1:A4\"), \">15\")\n\
               Debug.Print WorksheetFunction.SumIfs(Range(\"C1:C4\"), Range(\"A1:A4\"), \">15\", Range(\"B1:B4\"), \"a*\")\n\
               Debug.Print WorksheetFunction.AverageIfs(Range(\"C1:C4\"), Range(\"A1:A4\"), \">15\")\n\
             End Sub\n",
        )
        .unwrap();
        let debug_output = {
            let mut host = WorkbookHost::new(&mut workbook, 0).unwrap();
            execute_with_host(&module, "ManyCriteria", vec![], &mut host).unwrap();
            host.take_debug_output()
        };

        assert_eq!(
            debug_output,
            vec![
                "1".to_string(),
                "9".to_string(),
                "4".to_string(),
                "3".to_string(),
            ]
        );
    }

    /// Equality reads a cell spelling a number as that number; no other operator
    /// does, so the same text cell answers `=20` and `<>20` alike.
    #[test]
    fn criteria_only_coerce_numeric_text_for_equality() {
        let number = Value::Double(20.0);
        let spelled = Value::String("20".to_string());
        for criterion in ["20", "=20"] {
            let criteria = parse_criteria(&Value::String(criterion.to_string()));
            assert!(criteria.matches(&number), "{criterion} against a number");
            assert!(criteria.matches(&spelled), "{criterion} against text");
        }
        let criteria = parse_criteria(&Value::String("<>20".to_string()));
        assert!(!criteria.matches(&number));
        assert!(criteria.matches(&spelled));
        let criteria = parse_criteria(&Value::String(">=20".to_string()));
        assert!(criteria.matches(&number));
        assert!(!criteria.matches(&spelled));
    }

    #[test]
    fn criteria_read_operators_wildcards_and_blanks() {
        let blank = Value::Empty;
        let apple = Value::String("Apple".to_string());
        let truth = Value::Boolean(true);

        let criteria = parse_criteria(&Value::String(String::new()));
        assert!(criteria.matches(&blank));
        assert!(!criteria.matches(&apple));

        let criteria = parse_criteria(&Value::String("<>".to_string()));
        assert!(!criteria.matches(&blank));
        assert!(criteria.matches(&apple));

        // Wildcards survive negation, and stay case-insensitive.
        assert!(parse_criteria(&Value::String("a*".to_string())).matches(&apple));
        assert!(!parse_criteria(&Value::String("<>a*".to_string())).matches(&apple));
        assert!(parse_criteria(&Value::String("A*".to_string())).matches(&apple));

        // Spaces around an operator and its operand are ignored.
        assert!(parse_criteria(&Value::String("> 15".to_string())).matches(&Value::Integer(20)));
        assert!(parse_criteria(&Value::String(" 20".to_string())).matches(&Value::Integer(20)));

        // A criterion spelling a Boolean reads as one, and never as text.
        assert!(parse_criteria(&Value::String("TRUE".to_string())).matches(&truth));
        assert!(!parse_criteria(&Value::String("TRUE".to_string()))
            .matches(&Value::String("TRUE".to_string())));

        // Comparisons never reach across value classes.
        assert!(!parse_criteria(&Value::String(">15".to_string())).matches(&apple));
        assert!(!parse_criteria(&Value::String(">a".to_string())).matches(&Value::Integer(20)));
    }

    #[test]
    fn vba_conditionals_reject_what_excel_rejects() {
        let mut workbook = workbook();
        for source in [
            // An array is not a range Excel will test.
            "Public Function Bad() As Variant\n\
               Dim values(1 To 2) As Variant\n\
               values(1) = 10\n\
               Bad = WorksheetFunction.CountIf(values, \">5\")\n\
             End Function\n",
            // SumIfs demands the aggregated range match the criteria shape.
            "Public Function Bad() As Variant\n\
               Range(\"A1\").Value = 20\n\
               Bad = WorksheetFunction.SumIfs(Range(\"C1\"), Range(\"A1:A3\"), \">15\")\n\
             End Function\n",
            // Every criteria range must share one shape.
            "Public Function Bad() As Variant\n\
               Range(\"A1\").Value = 20\n\
               Bad = WorksheetFunction.CountIfs(Range(\"A1:A3\"), \">15\", Range(\"B1:B2\"), \"a*\")\n\
             End Function\n",
            // AverageIf has nothing to divide by when nothing matches.
            "Public Function Bad() As Variant\n\
               Range(\"A1\").Value = 20\n\
               Bad = WorksheetFunction.AverageIf(Range(\"A1:A3\"), \">100\")\n\
             End Function\n",
        ] {
            let module = parse_module(source).unwrap();
            let mut host = WorkbookHost::new(&mut workbook, 0).unwrap();
            execute_with_host(&module, "Bad", vec![], &mut host)
                .expect_err("Excel raises rather than returning an error value");
        }
    }

    #[test]
    fn vba_reads_a_cell_wherever_a_value_belongs() {
        let mut workbook = workbook();
        let module = parse_module(
            "Public Sub ImplicitValue()\n\
               Range(\"A1\").Value = 42\n\
               Range(\"A2\").Value = \"text\"\n\
               Debug.Print Range(\"A1\") + 1, \"x\" & Range(\"A1\"), Range(\"A1\") = 42\n\
               Debug.Print Range(\"A1\") > 40, -Range(\"A1\"), \"y\" & Range(\"A2\")\n\
               Debug.Print Range(\"A1\"), Range(\"A2\")\n\
               If Range(\"A1\") > 40 Then Debug.Print \"over\"\n\
             End Sub\n",
        )
        .unwrap();
        let debug_output = {
            let mut host = WorkbookHost::new(&mut workbook, 0).unwrap();
            execute_with_host(&module, "ImplicitValue", vec![], &mut host).unwrap();
            host.take_debug_output()
        };

        assert_eq!(
            debug_output,
            vec![
                "43\tx42\tTrue".to_string(),
                "True\t-42\tytext".to_string(),
                "42\ttext".to_string(),
                "over".to_string(),
            ]
        );
    }

    #[test]
    fn vba_cannot_print_a_range_of_several_cells() {
        let mut workbook = workbook();
        let module = parse_module(
            "Public Sub Bad()\n\
               Range(\"A1\").Value = 1\n\
               Range(\"A2\").Value = 2\n\
               Debug.Print Range(\"A1:A2\")\n\
             End Sub\n",
        )
        .unwrap();
        let mut host = WorkbookHost::new(&mut workbook, 0).unwrap();
        execute_with_host(&module, "Bad", vec![], &mut host)
            .expect_err("VBA reports a type mismatch");
    }

    /// A range covering several cells has an array for its value, which VBA
    /// refuses to treat as a scalar.
    #[test]
    fn vba_refuses_a_multi_cell_range_as_a_scalar() {
        let mut workbook = workbook();
        let module = parse_module(
            "Public Function Bad() As Variant\n\
               Range(\"A1\").Value = 1\n\
               Range(\"A2\").Value = 2\n\
               Bad = \"x\" & Range(\"A1:A2\")\n\
             End Function\n",
        )
        .unwrap();
        let mut host = WorkbookHost::new(&mut workbook, 0).unwrap();
        let error = execute_with_host(&module, "Bad", vec![], &mut host)
            .expect_err("VBA reports a type mismatch");
        assert!(
            error.to_string().contains("scalar"),
            "unexpected error: {error}"
        );
    }

    /// Excel keeps every number in a cell as a Double, whole or not, and that is
    /// the type VBA reports for it. Text, Booleans and blanks keep their own.
    #[test]
    fn a_cells_number_carries_the_double_type() {
        let mut workbook = workbook();
        let module = parse_module(
            "Public Sub CellTypes()\n\
               Range(\"A1\").Value = 42\n\
               Range(\"A2\").Value = 42.5\n\
               Range(\"A3\").Value = \"text\"\n\
               Range(\"A4\").Value = True\n\
               Debug.Print TypeName(Range(\"A1\").Value), TypeName(Range(\"A2\").Value)\n\
               Debug.Print TypeName(Range(\"A3\").Value), TypeName(Range(\"A4\").Value), TypeName(Range(\"A5\").Value)\n\
               Debug.Print VarType(Range(\"A1\")), VarType(Range(\"A3\")), VarType(Range(\"A5\"))\n\
               Debug.Print Range(\"A1\"), \"x\" & Range(\"A1\"), Range(\"A1\") + 1\n\
             End Sub\n",
        )
        .unwrap();
        let debug_output = {
            let mut host = WorkbookHost::new(&mut workbook, 0).unwrap();
            execute_with_host(&module, "CellTypes", vec![], &mut host).unwrap();
            host.take_debug_output()
        };

        assert_eq!(
            debug_output,
            vec![
                "Double\tDouble".to_string(),
                "String\tBoolean\tEmpty".to_string(),
                "5\t8\t0".to_string(),
                // A whole Double still reads back without a trailing zero.
                "42\tx42\t43".to_string(),
            ]
        );
    }

    /// Reads the top-left corner of a sheet the way the Excel probes printed it,
    /// so the expectations below are the grids Excel actually left behind.
    fn grid(workbook: &Workbook, rows: u32, columns: u32) -> String {
        (1..=rows)
            .map(|row| {
                (0..columns)
                    .map(|column| {
                        workbook.sheets[0]
                            .rows
                            .iter()
                            .find(|candidate| candidate.index == row)
                            .and_then(|found| found.cells.iter().find(|cell| cell.col == column))
                            .map(|cell| match &cell.value {
                                CellValue::String(value) => value.clone(),
                                value => value.display(),
                            })
                            .filter(|value| !value.is_empty())
                            .unwrap_or_else(|| ".".to_string())
                    })
                    .collect::<Vec<_>>()
                    .join(",")
            })
            .collect::<Vec<_>>()
            .join(" / ")
    }

    fn filled_grid() -> Workbook {
        let mut workbook = workbook();
        workbook.sheets[0].rows = (1..=4)
            .map(|row| Row {
                index: row,
                height: None,
                cells: (0..4)
                    .map(|column| Cell {
                        col: column,
                        value: CellValue::String(format!(
                            "{}{row}",
                            (b'A' + column as u8) as char
                        )),
                        style: CellStyle::default(),
                        formula: None,
                    })
                    .collect(),
            })
            .collect();
        workbook
    }

    fn run_on_grid(source: &str) -> Workbook {
        let mut workbook = filled_grid();
        let module = parse_module(source).unwrap();
        {
            let mut host = WorkbookHost::new(&mut workbook, 0).unwrap();
            execute_with_host(&module, "Act", vec![], &mut host).unwrap();
        }
        workbook
    }

    #[test]
    fn vba_inserts_and_deletes_whole_rows_and_columns() {
        let workbook = run_on_grid(
            "Public Sub Act()\n  Range(\"A2\").EntireRow.Insert\nEnd Sub\n",
        );
        assert_eq!(
            grid(&workbook, 4, 4),
            "A1,B1,C1,D1 / .,.,.,. / A2,B2,C2,D2 / A3,B3,C3,D3"
        );

        let workbook = run_on_grid(
            "Public Sub Act()\n  Range(\"B1\").EntireColumn.Insert\nEnd Sub\n",
        );
        assert_eq!(
            grid(&workbook, 4, 4),
            "A1,.,B1,C1 / A2,.,B2,C2 / A3,.,B3,C3 / A4,.,B4,C4"
        );

        let workbook = run_on_grid(
            "Public Sub Act()\n  Range(\"A2\").EntireRow.Delete\nEnd Sub\n",
        );
        assert_eq!(
            grid(&workbook, 4, 4),
            "A1,B1,C1,D1 / A3,B3,C3,D3 / A4,B4,C4,D4 / .,.,.,."
        );

        let workbook = run_on_grid(
            "Public Sub Act()\n  Range(\"B1\").EntireColumn.Delete\nEnd Sub\n",
        );
        assert_eq!(
            grid(&workbook, 4, 4),
            "A1,C1,D1,. / A2,C2,D2,. / A3,C3,D3,. / A4,C4,D4,."
        );

        // A band wider than one row moves everything below it that much further.
        let workbook = run_on_grid(
            "Public Sub Act()\n  Range(\"A2:A3\").EntireRow.Insert\nEnd Sub\n",
        );
        assert_eq!(
            grid(&workbook, 5, 2),
            "A1,B1 / .,. / .,. / A2,B2 / A3,B3"
        );
    }

    #[test]
    fn vba_moves_formulas_across_an_inserted_or_deleted_row() {
        let mut workbook = workbook();
        let module = parse_module(
            "Public Sub Act()\n\
               Range(\"A1\").Value = 10\n\
               Range(\"A2\").Value = 20\n\
               Range(\"B1\").Formula = \"=A1*2\"\n\
               Range(\"C1\").Formula = \"=SUM(A1:A2)\"\n\
               Range(\"D3\").Formula = \"=A2*2\"\n\
               Range(\"A1\").EntireRow.Insert\n\
             End Sub\n",
        )
        .unwrap();
        {
            let mut host = WorkbookHost::new(&mut workbook, 0).unwrap();
            execute_with_host(&module, "Act", vec![], &mut host).unwrap();
        }
        let formula = |row: u32, column: u32| {
            workbook.sheets[0]
                .rows
                .iter()
                .find(|candidate| candidate.index == row)
                .and_then(|found| found.cells.iter().find(|cell| cell.col == column))
                .and_then(|cell| cell.formula.clone())
                .unwrap_or_default()
        };
        // A sheet keeps a formula without its leading '=', so that is what the
        // shifted formula reads as too.
        assert_eq!(formula(2, 1), "A2*2");
        assert_eq!(formula(2, 2), "SUM(A2:A3)");
        assert_eq!(formula(4, 3), "A3*2");
    }

    #[test]
    fn a_formula_pointing_at_a_deleted_row_reads_ref() {
        let mut workbook = workbook();
        let module = parse_module(
            "Public Sub Act()\n\
               Range(\"A2\").Value = 20\n\
               Range(\"B1\").Formula = \"=A2*2\"\n\
               Range(\"A2\").EntireRow.Delete\n\
             End Sub\n",
        )
        .unwrap();
        {
            let mut host = WorkbookHost::new(&mut workbook, 0).unwrap();
            execute_with_host(&module, "Act", vec![], &mut host).unwrap();
        }
        let cell = workbook.sheets[0].rows[0]
            .cells
            .iter()
            .find(|cell| cell.col == 1)
            .expect("B1 still holds its formula");
        assert_eq!(cell.formula.as_deref(), Some("#REF!*2"));
    }

    /// A merge follows the rows under it, and one whose rows all go, goes too.
    #[test]
    fn merges_follow_an_inserted_or_deleted_row() {
        let mut workbook = filled_grid();
        workbook.sheets[0].merge_cells.push(MergeCell {
            start_row: 2,
            start_col: 1,
            end_row: 2,
            end_col: 2,
        });
        let module =
            parse_module("Public Sub Act()\n  Range(\"A1\").EntireRow.Insert\nEnd Sub\n").unwrap();
        {
            let mut host = WorkbookHost::new(&mut workbook, 0).unwrap();
            execute_with_host(&module, "Act", vec![], &mut host).unwrap();
        }
        let merge = &workbook.sheets[0].merge_cells[0];
        assert_eq!(
            (merge.start_row, merge.end_row, merge.start_col, merge.end_col),
            (3, 3, 1, 2)
        );

        let mut workbook = filled_grid();
        workbook.sheets[0].merge_cells.push(MergeCell {
            start_row: 2,
            start_col: 1,
            end_row: 2,
            end_col: 2,
        });
        let module =
            parse_module("Public Sub Act()\n  Range(\"A2\").EntireRow.Delete\nEnd Sub\n").unwrap();
        {
            let mut host = WorkbookHost::new(&mut workbook, 0).unwrap();
            execute_with_host(&module, "Act", vec![], &mut host).unwrap();
        }
        assert!(workbook.sheets[0].merge_cells.is_empty());
    }

    /// Every grid here is what Excel 16 left behind for the same call.
    #[test]
    fn vba_shifts_part_of_a_row_or_column() {
        // A single cell is as wide as it is tall, so it moves down.
        let workbook = run_on_grid("Public Sub Act()\n  Range(\"B2\").Insert\nEnd Sub\n");
        assert_eq!(
            grid(&workbook, 4, 4),
            "A1,B1,C1,D1 / A2,.,C2,D2 / A3,B2,C3,D3 / A4,B3,C4,D4"
        );

        // Taller than wide, so this one moves sideways instead.
        let workbook = run_on_grid("Public Sub Act()\n  Range(\"B2:B3\").Insert\nEnd Sub\n");
        assert_eq!(
            grid(&workbook, 4, 4),
            "A1,B1,C1,D1 / A2,.,B2,C2 / A3,.,B3,C3 / A4,B4,C4,D4"
        );

        // Wider than tall, so it moves down, carrying both columns.
        let workbook = run_on_grid("Public Sub Act()\n  Range(\"B2:C2\").Insert\nEnd Sub\n");
        assert_eq!(
            grid(&workbook, 4, 4),
            "A1,B1,C1,D1 / A2,.,.,D2 / A3,B2,C2,D3 / A4,B3,C3,D4"
        );

        let workbook = run_on_grid("Public Sub Act()\n  Range(\"B2\").Delete\nEnd Sub\n");
        assert_eq!(
            grid(&workbook, 4, 4),
            "A1,B1,C1,D1 / A2,B3,C2,D2 / A3,B4,C3,D3 / A4,.,C4,D4"
        );

        let workbook = run_on_grid("Public Sub Act()\n  Range(\"B2:B3\").Delete\nEnd Sub\n");
        assert_eq!(
            grid(&workbook, 4, 4),
            "A1,B1,C1,D1 / A2,C2,D2,. / A3,C3,D3,. / A4,B4,C4,D4"
        );

        let workbook = run_on_grid("Public Sub Act()\n  Range(\"B2:C3\").Delete\nEnd Sub\n");
        assert_eq!(
            grid(&workbook, 4, 4),
            "A1,B1,C1,D1 / A2,B4,C4,D2 / A3,.,.,D3 / A4,.,.,D4"
        );
    }

    /// An explicit shift overrides the shape's own leaning.
    #[test]
    fn an_explicit_shift_decides_the_direction() {
        let workbook =
            run_on_grid("Public Sub Act()\n  Range(\"B2\").Insert xlShiftToRight\nEnd Sub\n");
        assert_eq!(
            grid(&workbook, 4, 4),
            "A1,B1,C1,D1 / A2,.,B2,C2 / A3,B3,C3,D3 / A4,B4,C4,D4"
        );

        let workbook =
            run_on_grid("Public Sub Act()\n  Range(\"B2\").Delete xlShiftToLeft\nEnd Sub\n");
        assert_eq!(
            grid(&workbook, 4, 4),
            "A1,B1,C1,D1 / A2,C2,D2,. / A3,B3,C3,D3 / A4,B4,C4,D4"
        );
    }

    /// A partial shift moves only the formulas sharing its column, and leaves a
    /// range reaching past the band alone.
    #[test]
    fn a_partial_shift_moves_only_what_shares_its_column() {
        let mut workbook = workbook();
        let module = parse_module(
            "Public Sub Act()\n\
               Range(\"F1\").Formula = \"=B3*2\"\n\
               Range(\"F2\").Formula = \"=C3*2\"\n\
               Range(\"F3\").Formula = \"=SUM(B1:B4)\"\n\
               Range(\"F4\").Formula = \"=SUM(A1:C3)\"\n\
               Range(\"B2\").Insert xlShiftDown\n\
             End Sub\n",
        )
        .unwrap();
        {
            let mut host = WorkbookHost::new(&mut workbook, 0).unwrap();
            execute_with_host(&module, "Act", vec![], &mut host).unwrap();
        }
        let formula = |row: u32| {
            workbook.sheets[0]
                .rows
                .iter()
                .find(|candidate| candidate.index == row)
                .and_then(|found| found.cells.iter().find(|cell| cell.col == 5))
                .and_then(|cell| cell.formula.clone())
                .unwrap_or_default()
        };
        assert_eq!(formula(1), "B4*2");
        assert_eq!(formula(2), "C3*2");
        assert_eq!(formula(3), "SUM(B1:B5)");
        assert_eq!(formula(4), "SUM(A1:C3)");
    }

    #[test]
    fn vba_refuses_a_shift_that_does_not_belong(
    ) {
        let mut workbook = filled_grid();
        for source in [
            // xlShiftToLeft is a deletion's shift, not an insertion's.
            "Public Sub Act()\n  Range(\"B2\").Insert xlShiftToLeft\nEnd Sub\n",
            "Public Sub Act()\n  Range(\"B2\").Delete xlShiftDown\nEnd Sub\n",
        ] {
            let module = parse_module(source).unwrap();
            let mut host = WorkbookHost::new(&mut workbook, 0).unwrap();
            execute_with_host(&module, "Act", vec![], &mut host)
                .expect_err("that shift belongs to the other operation");
        }
    }

    /// A formula on one sheet follows rows inserted on the sheet it names.
    /// Measured against Excel with sheets called Data and Report.
    #[test]
    fn formulas_on_other_sheets_follow_the_rows_that_moved() {
        let mut workbook = workbook();
        workbook.sheets[0].name = "Data".to_string();
        let mut report = workbook.sheets[0].clone();
        report.name = "Report".to_string();
        report.rows = Vec::new();
        workbook.sheets.push(report);

        let module = parse_module(
            "Public Sub Act()\n\
               Worksheets(\"Report\").Range(\"A1\").Formula = \"=Data!A5*2\"\n\
               Worksheets(\"Report\").Range(\"A2\").Formula = \"=SUM(Data!A1:A6)\"\n\
               Worksheets(\"Report\").Range(\"A3\").Formula = \"=Report!A5*2\"\n\
               Worksheets(\"Data\").Range(\"A1\").EntireRow.Insert\n\
             End Sub\n",
        )
        .unwrap();
        {
            let mut host = WorkbookHost::new(&mut workbook, 0).unwrap();
            execute_with_host(&module, "Act", vec![], &mut host).unwrap();
        }

        let formula = |row: u32| {
            workbook.sheets[1]
                .rows
                .iter()
                .find(|candidate| candidate.index == row)
                .and_then(|found| found.cells.iter().find(|cell| cell.col == 0))
                .and_then(|cell| cell.formula.clone())
                .unwrap_or_default()
        };
        assert_eq!(formula(1), "Data!A6*2");
        assert_eq!(formula(2), "SUM(Data!A2:A7)");
        // Report's own rows never moved, so a reference to them stays.
        assert_eq!(formula(3), "Report!A5*2");
    }

    /// Renders the merges a sheet holds, in the address form the Excel probe
    /// printed, so the expectations below are what Excel actually left.
    fn merges(workbook: &Workbook) -> String {
        if workbook.sheets[0].merge_cells.is_empty() {
            return "none".to_string();
        }
        workbook.sheets[0]
            .merge_cells
            .iter()
            .map(|merge| {
                format!(
                    "{}{}:{}{}",
                    (b'A' + merge.start_col as u8) as char,
                    merge.start_row,
                    (b'A' + merge.end_col as u8) as char,
                    merge.end_row
                )
            })
            .collect::<Vec<_>>()
            .join(" ")
    }

    fn merged_grid(start_row: u32, start_col: u32, end_row: u32, end_col: u32) -> Workbook {
        let mut workbook = filled_grid();
        workbook.sheets[0].merge_cells.push(MergeCell {
            start_row,
            start_col,
            end_row,
            end_col,
        });
        workbook
    }

    fn act_on(mut workbook: Workbook, body: &str) -> Workbook {
        let module = parse_module(&format!("Public Sub Act()\n  {body}\nEnd Sub\n")).unwrap();
        {
            let mut host = WorkbookHost::new(&mut workbook, 0).unwrap();
            execute_with_host(&module, "Act", vec![], &mut host).unwrap();
        }
        workbook
    }

    /// A band carries a merge it covers and takes apart one it only half
    /// reaches. Every expectation is what Excel 16 left behind.
    #[test]
    fn a_merge_survives_only_a_band_that_covers_it() {
        // The band is column B alone, so the B2:C2 merge cannot come along.
        let workbook = act_on(
            merged_grid(2, 1, 2, 2),
            "Range(\"B2\").Insert xlShiftDown",
        );
        assert_eq!(merges(&workbook), "none");

        let workbook = act_on(merged_grid(2, 1, 2, 2), "Range(\"B2\").Delete xlShiftUp");
        assert_eq!(merges(&workbook), "none");

        // A band as wide as the merge carries it down whole.
        let workbook = act_on(
            merged_grid(2, 1, 2, 2),
            "Range(\"B2:C2\").Insert xlShiftDown",
        );
        assert_eq!(merges(&workbook), "B3:C3");

        // Taking one of its two rows leaves the merge a row shorter.
        let workbook = act_on(
            merged_grid(2, 1, 3, 2),
            "Range(\"B2:C2\").Delete xlShiftUp",
        );
        assert_eq!(merges(&workbook), "B2:C2");

        // A column band cannot carry a merge that spans two rows.
        let workbook = act_on(
            merged_grid(2, 1, 3, 1),
            "Range(\"B2\").Insert xlShiftToRight",
        );
        assert_eq!(merges(&workbook), "none");

        // Whole rows reach across everything, so these always carry.
        let workbook = act_on(merged_grid(2, 1, 2, 2), "Range(\"A2\").EntireRow.Delete");
        assert_eq!(merges(&workbook), "none");

        let workbook = act_on(merged_grid(2, 1, 3, 2), "Range(\"A2\").EntireRow.Delete");
        assert_eq!(merges(&workbook), "B2:C2");

        let workbook = act_on(merged_grid(2, 1, 3, 2), "Range(\"A1\").EntireRow.Insert");
        assert_eq!(merges(&workbook), "B3:C4");
    }

    /// A new sheet lands in front of the active one, becomes active, and is
    /// handed back. Measured against Excel with the first sheet named Base.
    #[test]
    fn vba_adds_a_worksheet_in_front_of_the_active_one() {
        let mut workbook = workbook();
        workbook.sheets[0].name = "Base".to_string();
        let module = parse_module(
            "Public Sub Act()\n\
               Dim added As Object\n\
               Set added = Worksheets.Add\n\
               Debug.Print added.Name, TypeName(added), ActiveSheet.Name\n\
             End Sub\n",
        )
        .unwrap();
        let debug_output = {
            let mut host = WorkbookHost::new(&mut workbook, 0).unwrap();
            execute_with_host(&module, "Act", vec![], &mut host).unwrap();
            host.take_debug_output()
        };

        assert_eq!(debug_output, vec!["Sheet1\tWorksheet\tSheet1".to_string()]);
        assert_eq!(
            workbook
                .sheets
                .iter()
                .map(|sheet| sheet.name.clone())
                .collect::<Vec<_>>(),
            vec!["Sheet1".to_string(), "Base".to_string()]
        );
    }

    #[test]
    fn vba_places_a_worksheet_before_or_after_another() {
        let mut workbook = workbook();
        workbook.sheets[0].name = "Base".to_string();
        let module = parse_module(
            "Public Sub Act()\n\
               Dim added As Object\n\
               Set added = Worksheets.Add(After:=Worksheets(\"Base\"))\n\
               added.Name = \"Tail\"\n\
               Set added = Worksheets.Add(Before:=Worksheets(\"Base\"))\n\
               added.Name = \"Head\"\n\
             End Sub\n",
        )
        .unwrap();
        {
            let mut host = WorkbookHost::new(&mut workbook, 0).unwrap();
            execute_with_host(&module, "Act", vec![], &mut host).unwrap();
        }
        assert_eq!(
            workbook
                .sheets
                .iter()
                .map(|sheet| sheet.name.clone())
                .collect::<Vec<_>>(),
            vec![
                "Head".to_string(),
                "Base".to_string(),
                "Tail".to_string()
            ]
        );
    }

    #[test]
    fn vba_deletes_a_worksheet_and_keeps_the_last_one() {
        let mut workbook = workbook();
        workbook.sheets[0].name = "Base".to_string();
        let module = parse_module(
            "Public Sub Act()\n\
               Dim added As Object\n\
               Set added = Worksheets.Add\n\
               added.Name = \"Second\"\n\
               Worksheets(\"Second\").Delete\n\
               Debug.Print Worksheets.Count, Worksheets(1).Name\n\
             End Sub\n",
        )
        .unwrap();
        let debug_output = {
            let mut host = WorkbookHost::new(&mut workbook, 0).unwrap();
            execute_with_host(&module, "Act", vec![], &mut host).unwrap();
            host.take_debug_output()
        };
        assert_eq!(debug_output, vec!["1\tBase".to_string()]);

        // Excel refuses to leave a workbook with no sheets at all.
        let module =
            parse_module("Public Sub Act()\n  Worksheets(\"Base\").Delete\nEnd Sub\n").unwrap();
        let mut host = WorkbookHost::new(&mut workbook, 0).unwrap();
        execute_with_host(&module, "Act", vec![], &mut host)
            .expect_err("the last worksheet cannot go");
    }

    #[test]
    fn vba_refuses_a_worksheet_name_another_sheet_holds() {
        let mut workbook = workbook();
        workbook.sheets[0].name = "Base".to_string();
        let module = parse_module(
            "Public Sub Act()\n\
               Dim added As Object\n\
               Set added = Worksheets.Add\n\
               added.Name = \"Base\"\n\
             End Sub\n",
        )
        .unwrap();
        let mut host = WorkbookHost::new(&mut workbook, 0).unwrap();
        execute_with_host(&module, "Act", vec![], &mut host)
            .expect_err("two sheets cannot share a name");
    }

    fn sorted_grid(fill: &str, call: &str, rows: u32, columns: u32) -> String {
        let mut workbook = workbook();
        let module = parse_module(&format!(
            "Public Sub Act()\n{fill}  {call}\nEnd Sub\n"
        ))
        .unwrap();
        {
            let mut host = WorkbookHost::new(&mut workbook, 0).unwrap();
            execute_with_host(&module, "Act", vec![], &mut host).unwrap();
        }
        grid(&workbook, rows, columns)
    }

    /// Each expectation is the grid Excel 16 left after the same call.
    #[test]
    fn vba_sorts_a_range_by_a_key_column() {
        let fill = "  Range(\"A1\").Value = 30\n  Range(\"B1\").Value = \"c\"\n\
                    \x20 Range(\"A2\").Value = 10\n  Range(\"B2\").Value = \"a\"\n\
                    \x20 Range(\"A3\").Value = 20\n  Range(\"B3\").Value = \"b\"\n";

        assert_eq!(
            sorted_grid(
                fill,
                "Range(\"A1:B3\").Sort Key1:=Range(\"A1\"), Order1:=xlAscending, Header:=xlNo",
                3,
                2
            ),
            "10,a / 20,b / 30,c"
        );
        assert_eq!(
            sorted_grid(
                fill,
                "Range(\"A1:B3\").Sort Key1:=Range(\"A1\"), Order1:=xlDescending, Header:=xlNo",
                3,
                2
            ),
            "30,c / 20,b / 10,a"
        );
        // The whole row travels with its key, whichever column the key is in.
        assert_eq!(
            sorted_grid(
                fill,
                "Range(\"A1:B3\").Sort Key1:=Range(\"B1\"), Order1:=xlDescending, Header:=xlNo",
                3,
                2
            ),
            "30,c / 20,b / 10,a"
        );
    }

    #[test]
    fn a_header_row_stays_where_it_is() {
        let fill = "  Range(\"A1\").Value = \"Name\"\n  Range(\"B1\").Value = \"Score\"\n\
                    \x20 Range(\"A2\").Value = \"c\"\n  Range(\"B2\").Value = 30\n\
                    \x20 Range(\"A3\").Value = \"a\"\n  Range(\"B3\").Value = 10\n\
                    \x20 Range(\"A4\").Value = \"b\"\n  Range(\"B4\").Value = 20\n";

        assert_eq!(
            sorted_grid(
                fill,
                "Range(\"A1:B4\").Sort Key1:=Range(\"A1\"), Order1:=xlAscending, Header:=xlYes",
                4,
                2
            ),
            "Name,Score / a,10 / b,20 / c,30"
        );
        // Left out, the header is just another row — Excel sorts it in.
        assert_eq!(
            sorted_grid(
                fill,
                "Range(\"A1:B4\").Sort Key1:=Range(\"A1\"), Order1:=xlAscending",
                4,
                2
            ),
            "a,10 / b,20 / c,30 / Name,Score"
        );
    }

    /// Numbers come first, then text, then Booleans, and a blank stays at the
    /// bottom whichever way the sort runs.
    #[test]
    fn a_sort_orders_values_by_kind_before_value() {
        let fill = "  Range(\"A1\").Value = 10\n  Range(\"A2\").Value = \"apple\"\n\
                    \x20 Range(\"A4\").Value = 2\n  Range(\"A5\").Value = True\n\
                    \x20 Range(\"A6\").Value = \"Banana\"\n";

        assert_eq!(
            sorted_grid(
                fill,
                "Range(\"A1:A6\").Sort Key1:=Range(\"A1\"), Order1:=xlAscending, Header:=xlNo",
                6,
                1
            ),
            // Excel shows a Boolean cell as TRUE; the probe read it back through
            // CStr, which spells it True. Only the order is being checked here.
            "2 / 10 / apple / Banana / TRUE / ."
        );
        assert_eq!(
            sorted_grid(
                fill,
                "Range(\"A1:A6\").Sort Key1:=Range(\"A1\"), Order1:=xlDescending, Header:=xlNo",
                6,
                1
            ),
            "TRUE / Banana / apple / 10 / 2 / ."
        );
    }

    /// Text ignores case, and values that compare equal keep the order they
    /// arrived in — b before B because b was already first.
    #[test]
    fn equal_values_keep_the_order_they_had() {
        let fill = "  Range(\"A1\").Value = \"b\"\n  Range(\"A2\").Value = \"A\"\n\
                    \x20 Range(\"A3\").Value = \"a\"\n  Range(\"A4\").Value = \"B\"\n";
        assert_eq!(
            sorted_grid(
                fill,
                "Range(\"A1:A4\").Sort Key1:=Range(\"A1\"), Order1:=xlAscending, Header:=xlNo",
                4,
                1
            ),
            "A / a / b / B"
        );
    }

    #[test]
    fn a_second_key_settles_a_tie() {
        let fill = "  Range(\"A1\").Value = 1\n  Range(\"B1\").Value = \"b\"\n\
                    \x20 Range(\"A2\").Value = 1\n  Range(\"B2\").Value = \"a\"\n\
                    \x20 Range(\"A3\").Value = 2\n  Range(\"B3\").Value = \"a\"\n";
        assert_eq!(
            sorted_grid(
                fill,
                "Range(\"A1:B3\").Sort Key1:=Range(\"A1\"), Key2:=Range(\"B1\"), Header:=xlNo",
                3,
                2
            ),
            "1,a / 1,b / 2,a"
        );
    }

    #[test]
    fn a_sort_can_run_left_to_right() {
        let fill = "  Range(\"A1\").Value = 3\n  Range(\"B1\").Value = 1\n  Range(\"C1\").Value = 2\n\
                    \x20 Range(\"A2\").Value = \"c\"\n  Range(\"B2\").Value = \"a\"\n  Range(\"C2\").Value = \"b\"\n";
        assert_eq!(
            sorted_grid(
                fill,
                "Range(\"A1:C2\").Sort Key1:=Range(\"A1\"), Header:=xlNo, Orientation:=xlLeftToRight",
                2,
                3
            ),
            "1,2,3 / a,b,c"
        );
    }

    #[test]
    fn vba_refuses_a_sort_it_cannot_carry_out() {
        for call in [
            // Excel quietly sorts nothing for a key outside the range.
            "Range(\"A1:A2\").Sort Key1:=Range(\"D1\"), Header:=xlNo",
            "Range(\"A1:A2\").Sort Key1:=Range(\"A1\"), Header:=xlGuess",
            "Range(\"A1:A2\").Sort Key1:=Range(\"A1\"), Header:=xlNo, MatchCase:=True",
            "Range(\"A1:A2\").Sort Header:=xlNo",
        ] {
            let mut workbook = workbook();
            let module = parse_module(&format!(
                "Public Sub Act()\n  Range(\"A1\").Value = 2\n  Range(\"A2\").Value = 1\n  {call}\nEnd Sub\n"
            ))
            .unwrap();
            let mut host = WorkbookHost::new(&mut workbook, 0).unwrap();
            execute_with_host(&module, "Act", vec![], &mut host)
                .unwrap_err();
        }
    }

    /// The worksheet's Round sends a half away from zero while VBA's sends it to
    /// the even neighbour. Both answers below came from Excel in one macro.
    #[test]
    fn the_worksheets_round_and_vbas_round_disagree_at_a_half() {
        let mut workbook = workbook();
        let module = parse_module(
            "Public Sub Act()\n\
               Debug.Print WorksheetFunction.Round(2.5, 0), WorksheetFunction.Round(3.5, 0), WorksheetFunction.Round(-2.5, 0)\n\
               Debug.Print Round(2.5, 0), Round(3.5, 0)\n\
               Debug.Print WorksheetFunction.Round(2.345, 2), WorksheetFunction.Round(25, -1)\n\
               Debug.Print WorksheetFunction.RoundUp(2.1, 0), WorksheetFunction.RoundUp(-2.1, 0)\n\
               Debug.Print WorksheetFunction.RoundDown(2.9, 0), WorksheetFunction.RoundDown(-2.9, 0)\n\
             End Sub\n",
        )
        .unwrap();
        let debug_output = {
            let mut host = WorkbookHost::new(&mut workbook, 0).unwrap();
            execute_with_host(&module, "Act", vec![], &mut host).unwrap();
            host.take_debug_output()
        };

        assert_eq!(
            debug_output,
            vec![
                "3\t4\t-3".to_string(),
                "2\t4".to_string(),
                "2.35\t30".to_string(),
                "3\t-3".to_string(),
                "2\t-2".to_string(),
            ]
        );
    }

    /// The worksheet's Trim squeezes runs of spaces; VBA's only strips the ends.
    #[test]
    fn the_worksheets_trim_squeezes_what_vbas_leaves() {
        let mut workbook = workbook();
        let module = parse_module(
            "Public Sub Act()\n\
               Debug.Print \"[\" & WorksheetFunction.Trim(\"  a   b  \") & \"]\"\n\
               Debug.Print \"[\" & Trim(\"  a   b  \") & \"]\"\n\
               Debug.Print WorksheetFunction.Proper(\"hello wORLD\"), WorksheetFunction.Proper(\"o'neil 3rd\")\n\
             End Sub\n",
        )
        .unwrap();
        let debug_output = {
            let mut host = WorkbookHost::new(&mut workbook, 0).unwrap();
            execute_with_host(&module, "Act", vec![], &mut host).unwrap();
            host.take_debug_output()
        };

        assert_eq!(
            debug_output,
            vec![
                "[a b]".to_string(),
                "[a   b]".to_string(),
                // A letter after anything that is not a letter starts a word, so
                // both the apostrophe and the digit capitalise what follows.
                "Hello World\tO'Neil 3Rd".to_string(),
            ]
        );
    }

    #[test]
    fn vba_ranks_and_averages_the_numbers_in_a_range() {
        let mut workbook = workbook();
        let module = parse_module(
            "Public Sub Act()\n\
               Range(\"A1\").Value = 30\n\
               Range(\"A2\").Value = 10\n\
               Range(\"A3\").Value = 20\n\
               Range(\"A4\").Value = 20\n\
               Range(\"C1\").Value = 5\n\
               Range(\"C3\").Value = \"text\"\n\
               Range(\"C4\").Value = 1\n\
               Debug.Print WorksheetFunction.Large(Range(\"A1:A4\"), 1), WorksheetFunction.Large(Range(\"A1:A4\"), 2), WorksheetFunction.Small(Range(\"A1:A4\"), 1)\n\
               Debug.Print WorksheetFunction.Median(Range(\"A1:A4\")), WorksheetFunction.Median(Range(\"A1:A3\"))\n\
               Debug.Print WorksheetFunction.Large(Range(\"C1:C4\"), 1), WorksheetFunction.Median(Range(\"C1:C4\"))\n\
               Debug.Print WorksheetFunction.Power(2, 10), WorksheetFunction.SumProduct(Range(\"A1:A2\"), Range(\"A3:A4\"))\n\
             End Sub\n",
        )
        .unwrap();
        let debug_output = {
            let mut host = WorkbookHost::new(&mut workbook, 0).unwrap();
            execute_with_host(&module, "Act", vec![], &mut host).unwrap();
            host.take_debug_output()
        };

        assert_eq!(
            debug_output,
            vec![
                "30\t20\t10".to_string(),
                "20\t20".to_string(),
                // Text and blanks are passed over, leaving 5 and 1.
                "5\t3".to_string(),
                "1024\t800".to_string(),
            ]
        );
    }

    #[test]
    fn a_rank_beyond_the_values_is_refused() {
        let mut workbook = workbook();
        let module = parse_module(
            "Public Function Act() As Variant\n\
               Range(\"A1\").Value = 30\n\
               Act = WorksheetFunction.Large(Range(\"A1:A4\"), 9)\n\
             End Function\n",
        )
        .unwrap();
        let mut host = WorkbookHost::new(&mut workbook, 0).unwrap();
        execute_with_host(&module, "Act", vec![], &mut host)
            .expect_err("Excel raises for a rank it cannot reach");
    }

    /// Reads a cell the way the Excel probe did: its value, its formula, and
    /// whether it came out bold.
    fn looked_at(workbook: &Workbook, row: u32, column: u32) -> String {
        let cell = workbook.sheets[0]
            .rows
            .iter()
            .find(|candidate| candidate.index == row)
            .and_then(|found| found.cells.iter().find(|cell| cell.col == column));
        match cell {
            None => "v= f= bold=False".to_string(),
            Some(cell) => format!(
                "v={} f={} bold={}",
                cell.value.display(),
                cell.formula.clone().unwrap_or_default(),
                cell.style.bold
            ),
        }
    }

    fn pasted(call: &str) -> Workbook {
        let mut workbook = workbook();
        let module = parse_module(&format!(
            "Public Sub Act()\n\
               Range(\"A1\").Value = 10\n\
               Range(\"A2\").Formula = \"=A1*2\"\n\
               Range(\"A1:A2\").Font.Bold = True\n\
               {call}\n\
             End Sub\n"
        ))
        .unwrap();
        {
            let mut host = WorkbookHost::new(&mut workbook, 0).unwrap();
            execute_with_host(&module, "Act", vec![], &mut host).unwrap();
        }
        workbook
    }

    /// A whole paste brings the formula along, moved by the distance it
    /// travelled. Measured against Excel.
    #[test]
    fn a_whole_paste_carries_the_formula_and_the_formatting() {
        let workbook = pasted(
            "Range(\"A1:A2\").Copy\n  Range(\"C1\").PasteSpecial xlPasteAll",
        );
        assert_eq!(looked_at(&workbook, 1, 2), "v=10 f= bold=true");
        assert_eq!(looked_at(&workbook, 2, 2), "v= f=C1*2 bold=true");

        // Leaving the argument out pastes everything too.
        let workbook = pasted("Range(\"A1:A2\").Copy\n  Range(\"C1\").PasteSpecial");
        assert_eq!(looked_at(&workbook, 2, 2), "v= f=C1*2 bold=true");
    }

    #[test]
    fn a_values_paste_drops_the_formula_and_the_formatting() {
        let workbook = pasted(
            "Range(\"A1:A2\").Copy\n  Range(\"C1\").PasteSpecial xlPasteValues",
        );
        assert_eq!(looked_at(&workbook, 1, 2), "v=10 f= bold=false");
        // The formula does not come; what the cell was holding does.
        assert_eq!(looked_at(&workbook, 2, 2), "v= f= bold=false");
    }

    #[test]
    fn a_formats_paste_leaves_the_value_alone() {
        let workbook = pasted(
            "Range(\"C1\").Value = 99\n  Range(\"A1:A2\").Copy\n  Range(\"C1\").PasteSpecial xlPasteFormats",
        );
        assert_eq!(looked_at(&workbook, 1, 2), "v=99 f= bold=true");
    }

    #[test]
    fn a_bigger_target_takes_whole_copies_of_the_block() {
        let workbook = pasted(
            "Range(\"A1:A2\").Copy\n  Range(\"C1:C4\").PasteSpecial xlPasteValues",
        );
        assert_eq!(looked_at(&workbook, 3, 2), "v=10 f= bold=false");

        let workbook = pasted(
            "Range(\"A1:A2\").Copy\n  Range(\"E1\").PasteSpecial xlPasteValues, , , True",
        );
        assert_eq!(looked_at(&workbook, 1, 4), "v=10 f= bold=false");
        assert_eq!(looked_at(&workbook, 1, 5), "v= f= bold=false");
    }

    #[test]
    fn vba_refuses_a_paste_it_cannot_carry_out() {
        for call in [
            // Nothing was copied first.
            "Range(\"C1\").PasteSpecial xlPasteValues",
            // The block does not divide the target evenly.
            "Range(\"A1:A2\").Copy\n  Range(\"C1:C3\").PasteSpecial xlPasteValues",
        ] {
            let mut workbook = workbook();
            let module = parse_module(&format!(
                "Public Sub Act()\n  Range(\"A1\").Value = 10\n  Range(\"A2\").Value = 20\n  {call}\nEnd Sub\n"
            ))
            .unwrap();
            let mut host = WorkbookHost::new(&mut workbook, 0).unwrap();
            execute_with_host(&module, "Act", vec![], &mut host).unwrap_err();
        }
    }

    fn sheet_names(workbook: &Workbook) -> String {
        workbook
            .sheets
            .iter()
            .map(|sheet| sheet.name.clone())
            .collect::<Vec<_>>()
            .join(",")
    }

    fn with_base(body: &str) -> Workbook {
        let mut workbook = workbook();
        workbook.sheets[0].name = "Base".to_string();
        let module = parse_module(&format!("Public Sub Act()\n{body}\nEnd Sub\n")).unwrap();
        {
            let mut host = WorkbookHost::new(&mut workbook, 0).unwrap();
            execute_with_host(&module, "Act", vec![], &mut host).unwrap();
        }
        workbook
    }

    /// A copy is named after the sheet it came from, numbered from two, and
    /// becomes the active sheet. Measured against Excel.
    #[test]
    fn vba_copies_a_worksheet_beside_another() {
        let workbook = with_base(
            "  Range(\"A1\").Value = 42\n\
             \x20 Range(\"A1\").Font.Bold = True\n\
             \x20 Worksheets(\"Base\").Copy After:=Worksheets(\"Base\")\n",
        );
        assert_eq!(sheet_names(&workbook), "Base,Base (2)");
        // The copy carries the cells and their formatting.
        let copied = &workbook.sheets[1].rows[0].cells[0];
        assert_eq!(copied.value.display(), "42");
        assert!(copied.style.bold);

        let workbook = with_base(
            "  Worksheets.Add.Name = \"Second\"\n\
             \x20 Worksheets(\"Base\").Copy Before:=Worksheets(\"Second\")\n",
        );
        assert_eq!(sheet_names(&workbook), "Base (2),Second,Base");
    }

    /// Copying twice takes the next unused number, while the position still
    /// follows the argument — so the later copy sits in front of the earlier.
    #[test]
    fn a_second_copy_takes_the_next_number() {
        let workbook = with_base(
            "  Worksheets(\"Base\").Copy After:=Worksheets(\"Base\")\n\
             \x20 Worksheets(\"Base\").Copy After:=Worksheets(\"Base\")\n",
        );
        assert_eq!(sheet_names(&workbook), "Base,Base (3),Base (2)");
    }

    #[test]
    fn vba_moves_a_worksheet_without_making_another() {
        let workbook = with_base(
            "  Worksheets.Add.Name = \"Second\"\n\
             \x20 Worksheets(\"Base\").Move After:=Worksheets(\"Second\")\n",
        );
        assert_eq!(sheet_names(&workbook), "Second,Base");
        assert_eq!(workbook.sheets.len(), 2);
    }

    #[test]
    fn vba_refuses_a_placement_it_cannot_make() {
        for body in [
            // Excel refuses Before and After together.
            "  Worksheets.Add.Name = \"Second\"\n\
             \x20 Worksheets(\"Base\").Copy Before:=Worksheets(\"Second\"), After:=Worksheets(\"Second\")\n",
            // With neither, Excel copies into a new workbook, which the browser
            // host has no room for.
            "  Worksheets(\"Base\").Copy\n",
            "  Worksheets(\"Base\").Move\n",
        ] {
            let mut workbook = workbook();
            workbook.sheets[0].name = "Base".to_string();
            let module = parse_module(&format!("Public Sub Act()\n{body}\nEnd Sub\n")).unwrap();
            let mut host = WorkbookHost::new(&mut workbook, 0).unwrap();
            execute_with_host(&module, "Act", vec![], &mut host).unwrap_err();
        }
    }

    #[test]
    fn vba_tracks_excel_application_settings() {
        let mut workbook = workbook();
        let module = parse_module(
            "Public Sub ConfigureApplication()\n\
               Application.ScreenUpdating = False\n\
               Application.EnableEvents = False\n\
               Application.DisplayAlerts = False\n\
               Application.Calculation = xlCalculationManual\n\
               Debug.Print Application.ScreenUpdating, Application.EnableEvents, Application.DisplayAlerts, Application.Calculation\n\
               Application.ScreenUpdating = True\n\
               Application.EnableEvents = True\n\
               Application.DisplayAlerts = True\n\
               Application.Calculation = xlCalculationAutomatic\n\
               Debug.Print Application.ScreenUpdating, Application.EnableEvents, Application.DisplayAlerts, Application.Calculation\n\
             End Sub\n",
        )
        .unwrap();
        let mut host = WorkbookHost::new(&mut workbook, 0).unwrap();
        execute_with_host(&module, "ConfigureApplication", vec![], &mut host).unwrap();

        assert!(host.screen_updating);
        assert!(host.enable_events);
        assert!(host.display_alerts);
        assert_eq!(host.calculation, -4105);
        assert_eq!(
            host.take_debug_output(),
            vec![
                "False\tFalse\tFalse\t-4135".to_string(),
                "True\tTrue\tTrue\t-4105".to_string(),
            ]
        );
    }

    #[test]
    fn vba_finds_values_and_formulas_in_ranges() {
        let mut workbook = workbook();
        let module = parse_module(
            "Public Sub FindCells()\n\
               Range(\"A1\").Value = \"Alpha\"\n\
               Range(\"B1\").Value = 42\n\
               Range(\"A2\").Value = \"needle in text\"\n\
               Range(\"B2\").Formula = \"=A1\"\n\
               Range(\"A3\").Value = \"Alpha\"\n\
               Set partial = Range(\"A1:B2\").Find(\"NEEDLE\")\n\
               Set exact = Range(\"A1:B2\").Find(42, , xlValues, xlWhole, xlByColumns, xlNext, False)\n\
               Set formula = Range(\"A1:B2\").Find(\"=A1\", , xlFormulas, xlWhole)\n\
               Set previous = Range(\"A1:B3\").Find(\"Alpha\", Range(\"A1\"), xlValues, xlWhole, xlByRows, xlPrevious, False)\n\
               Set missing = Range(\"A1:B2\").Find(\"absent\")\n\
               Set caseMiss = Range(\"A1:B2\").Find(\"alpha\", , xlValues, xlWhole, xlByRows, xlNext, True)\n\
               Debug.Print partial.Address(False, False), exact.Address(False, False), formula.Address(False, False), previous.Address(False, False), missing Is Nothing, caseMiss Is Nothing\n\
             End Sub\n",
        )
        .unwrap();
        let debug_output = {
            let mut host = WorkbookHost::new(&mut workbook, 0).unwrap();
            execute_with_host(&module, "FindCells", vec![], &mut host).unwrap();
            host.take_debug_output()
        };

        assert_eq!(debug_output, vec!["A2\tB1\tB2\tA3\tTrue\tTrue".to_string()]);
    }

    #[test]
    fn vba_continues_the_previous_range_search() {
        let mut workbook = workbook();
        let module = parse_module(
            "Public Sub ContinueSearch()\n\
               Range(\"A1\").Value = \"hit\"\n\
               Range(\"A3\").Value = \"hit\"\n\
               Set first = Range(\"A1:A3\").Find(What:=\"hit\", LookIn:=xlValues, LookAt:=xlWhole, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:=False)\n\
               Set following = Range(\"A1:A3\").FindNext(After:=first)\n\
               Set preceding = Range(\"A1:A3\").FindPrevious(following)\n\
               Set wrapped = Range(\"A1:A3\").FindNext()\n\
               Debug.Print first.Address(False, False), following.Address(False, False), preceding.Address(False, False), wrapped.Address(False, False)\n\
               Set missing = Range(\"A1:A3\").Find(\"absent\")\n\
               Set stillMissing = Range(\"A1:A3\").FindNext()\n\
               Debug.Print missing Is Nothing, stillMissing Is Nothing\n\
             End Sub\n",
        )
        .unwrap();
        let debug_output = {
            let mut host = WorkbookHost::new(&mut workbook, 0).unwrap();
            execute_with_host(&module, "ContinueSearch", vec![], &mut host).unwrap();
            host.take_debug_output()
        };

        assert_eq!(
            debug_output,
            vec!["A3\tA1\tA3\tA1".to_string(), "True\tTrue".to_string(),]
        );
    }

    #[test]
    fn vba_replaces_values_and_formula_text_in_ranges() {
        let mut workbook = workbook();
        let module = parse_module(
            "Public Sub ReplaceCells()\n\
               Range(\"A1\").Value = \"foo bar\"\n\
               Range(\"A2\").Value = \"FOO\"\n\
               Range(\"B1\").Value = 42\n\
               Range(\"B2\").Formula = \"=A1&\"\"foo\"\"\"\n\
               changed = Range(\"A1:B2\").Replace(What:=\"foo\", Replacement:=\"baz\", LookAt:=xlPart, SearchOrder:=xlByRows, MatchCase:=False)\n\
               exact = Range(\"B1\").Replace(What:=42, Replacement:=7, LookAt:=xlWhole)\n\
               missing = Range(\"A1:B2\").Replace(What:=\"absent\", Replacement:=\"unused\")\n\
               Debug.Print changed, exact, missing, Range(\"A1\").Value, Range(\"A2\").Value, Range(\"B1\").Value, Range(\"B2\").Formula\n\
             End Sub\n",
        )
        .unwrap();
        let debug_output = {
            let mut host = WorkbookHost::new(&mut workbook, 0).unwrap();
            execute_with_host(&module, "ReplaceCells", vec![], &mut host).unwrap();
            host.take_debug_output()
        };

        assert_eq!(
            debug_output,
            vec!["True\tTrue\tFalse\tbaz bar\tbaz\t7\t=A1&\"baz\"".to_string()]
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
