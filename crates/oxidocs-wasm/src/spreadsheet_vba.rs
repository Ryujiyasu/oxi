// SPDX-License-Identifier: MIT OR Apache-2.0

use oxicells_core::ir::{Cell, CellStyle, CellValue, Row, Workbook};
use oxivba_core::{execute_with_host, parse_module, ArrayValue, Host, ObjectRef, Value};
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
    RangeCollection(CellRange, RangeAxis),
    Worksheet(usize),
    Worksheets,
    Workbook,
    Application,
}

struct WorkbookHost<'a> {
    workbook: &'a mut Workbook,
    active_sheet: usize,
    objects: Vec<HostObject>,
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
            objects: Vec::new(),
        })
    }

    fn object(&mut self, object: HostObject) -> Value {
        let handle = self.objects.len() as u64;
        self.objects.push(object);
        Value::Object(ObjectRef {
            handle,
            kind: match object {
                HostObject::Range(_) => "Range",
                HostObject::RangeCollection(_, RangeAxis::Rows) => "Rows",
                HostObject::RangeCollection(_, RangeAxis::Columns) => "Columns",
                HostObject::Worksheet(_) => "Worksheet",
                HostObject::Worksheets => "Worksheets",
                HostObject::Workbook => "Workbook",
                HostObject::Application => "Application",
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
        let rows = u64::from(range.end_row - range.start_row) + 1;
        let columns = u64::from(range.end_column - range.start_column) + 1;
        let count = rows
            .checked_mul(columns)
            .ok_or_else(|| "cell range is too large".to_string())?;
        if count > 1_000_000 {
            return Err("cell range exceeds the 1,000,000-cell execution limit".to_string());
        }
        Ok(count as usize)
    }

    fn range_value(&self, range: CellRange) -> Result<Value, String> {
        Self::range_cell_count(range)?;
        if range.is_single() {
            return Ok(self.cell_value(range.addresses().next().unwrap()));
        }
        Ok(Value::Array(ArrayValue {
            lower_bound: 1,
            values: range
                .addresses()
                .map(|address| self.cell_value(address))
                .collect(),
            element_default: Box::new(Value::Empty),
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
            lower_bound: 1,
            values: range
                .addresses()
                .map(|address| self.cell_formula(address))
                .collect(),
            element_default: Box::new(Value::String(String::new())),
        }))
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
}

impl Host for WorkbookHost<'_> {
    fn call(
        &mut self,
        receiver: Option<&ObjectRef>,
        name: &str,
        args: &[Value],
    ) -> Result<Option<Value>, String> {
        if let Some(receiver) = receiver {
            if let Some(sheet) = self.worksheet(receiver) {
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
                    return Ok(Some(Value::Empty));
                }
                if name.eq_ignore_ascii_case("usedrange") {
                    if !args.is_empty() {
                        return Err("Worksheet.UsedRange does not accept arguments".to_string());
                    }
                    return self.used_range_object(sheet).map(Some);
                }
                return Ok(None);
            }
            if (self.is_workbook(receiver) || self.is_application(receiver))
                && (name.eq_ignore_ascii_case("worksheets") || name.eq_ignore_ascii_case("sheets"))
            {
                return self.worksheets_object_or_item(args).map(Some);
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
            }
            return Ok(None);
        }
        if args.is_empty() {
            if let Some(value) = excel_constant(name) {
                return Ok(Some(value));
            }
        }
        if name.eq_ignore_ascii_case("range") {
            return self.range_object(self.active_sheet, args).map(Some);
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
        if name.eq_ignore_ascii_case("usedrange") {
            if !args.is_empty() {
                return Err("UsedRange does not accept arguments".to_string());
            }
            return self.used_range_object(self.active_sheet).map(Some);
        }
        Ok(None)
    }

    fn get(&mut self, receiver: &ObjectRef, name: &str) -> Result<Option<Value>, String> {
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
        if name.eq_ignore_ascii_case("row") {
            return Ok(Some(Value::Integer(i64::from(range.start_row))));
        }
        if name.eq_ignore_ascii_case("column") {
            return Ok(Some(Value::Integer(i64::from(range.start_column) + 1)));
        }
        if name.eq_ignore_ascii_case("count") || name.eq_ignore_ascii_case("countlarge") {
            return Self::range_cell_count(range).map(|count| Some(Value::Integer(count as i64)));
        }
        if name.eq_ignore_ascii_case("address") {
            return Ok(Some(Value::String(format_range_address(range, true, true))));
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

fn cell_has_content(cell: &Cell) -> bool {
    cell.formula.is_some() || !matches!(&cell.value, CellValue::Empty)
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

fn excel_constant(name: &str) -> Option<Value> {
    let value = if name.eq_ignore_ascii_case("xlup") {
        -4162
    } else if name.eq_ignore_ascii_case("xldown") {
        -4121
    } else if name.eq_ignore_ascii_case("xltoleft") {
        -4159
    } else if name.eq_ignore_ascii_case("xltoright") {
        -4161
    } else {
        return None;
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
    let reference = reference.trim();
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
        CellValue::String(value) | CellValue::Error(value) => Value::String(value.clone()),
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
        Value::Boolean(value) => Ok(CellValue::Boolean(value)),
        Value::Integer(value) => Ok(CellValue::Number(value as f64)),
        Value::Double(value) => Ok(CellValue::Number(value)),
        Value::String(value) => Ok(CellValue::String(value)),
        Value::Array(_) => Err("a VBA array cannot be assigned to one cell".to_string()),
        Value::Object(_) => Err("a VBA object cannot be assigned to one cell".to_string()),
    }
}

fn to_formula(value: Value) -> Result<String, String> {
    match value {
        Value::Empty | Value::Null => Ok(String::new()),
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
    Null,
    Boolean(bool),
    Integer(i64),
    Double(f64),
    String(String),
    Array {
        lower_bound: i64,
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
            Value::Null => Self::Null,
            Value::Boolean(value) => Self::Boolean(value),
            Value::Integer(value) => Self::Integer(value),
            Value::Double(value) => Self::Double(value),
            Value::String(value) => Self::String(value),
            Value::Array(value) => Self::Array {
                lower_bound: value.lower_bound,
                values: value.values.into_iter().map(OutputValue::from).collect(),
            },
            Value::Object(value) => Self::Object {
                handle: value.handle,
                kind: value.kind,
            },
        }
    }
}

#[derive(Serialize)]
struct RunResult {
    workbook: Workbook,
    result: OutputValue,
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
    let result = execute_with_host(
        &module,
        procedure,
        args.into_iter().map(Value::from).collect(),
        &mut host,
    )
    .map_err(|error| JsError::new(&error.to_string()))?;
    serde_wasm_bindgen::to_value(&RunResult {
        workbook,
        result: result.into(),
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
               FillRange = total\n\
             End Function\n\
             Public Sub WriteArray()\n\
               Dim values(1 To 4) As Long\n\
               values(1) = 10\n\
               values(2) = 20\n\
               values(3) = 30\n\
               values(4) = 40\n\
               Range(\"C1:D2\").Value = values\n\
             End Sub\n",
        )
        .unwrap();
        {
            let mut host = WorkbookHost::new(&mut workbook, 0).unwrap();
            let mut runtime = oxivba_core::Runtime::new(&module).with_host(&mut host);
            assert_eq!(
                runtime.call("FillRange", vec![]).unwrap(),
                Value::Integer(20)
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
