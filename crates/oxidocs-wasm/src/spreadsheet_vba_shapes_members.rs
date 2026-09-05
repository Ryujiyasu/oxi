// SPDX-License-Identifier: MIT OR Apache-2.0

//! The members of the drawing objects: what `Shape`, `TextFrame`, `Chart`,
//! `Series` and the rest answer and take. Defaults and refusals as measured
//! (shapes.vba, shapes2.vba, 2026-09-05).

use super::shapes::*;
use super::*;

/// E_INVALIDARG, which is what `Shapes("nope")` and `Shapes(9)` raise.
const NO_SUCH_SHAPE: i64 = -2_147_024_809;

impl<'a> WorkbookHost<'a> {
    /// What `Selection` is: a ShapeRange when a macro has selected shapes,
    /// otherwise the cells.
    pub(super) fn selection_object(&mut self) -> Value {
        let live: Vec<u64> = self.shape_selection.iter().copied().filter(|id| self.shape(*id).is_ok()).collect();
        if live.is_empty() {
            return self.object(HostObject::Range(self.selection));
        }
        self.shape_range_object(live)
    }

    fn shape_range_object(&mut self, ids: Vec<u64>) -> Value {
        let index = self.shape_ranges.len();
        self.shape_ranges.push(ids);
        self.part(DrawingPart::ShapeRange(index))
    }

    fn shape_range_ids(&self, index: usize) -> Vec<u64> {
        self.shape_ranges.get(index).cloned().unwrap_or_default()
    }

    /// `ActiveChart`: the chart of the selected shape when it holds one, or
    /// the sheet's one chart when only one is there. Measured: after
    /// `AddChart2(...).Select`, ActiveChart is that chart. None -> 91.
    pub(super) fn active_chart_object(&mut self) -> Result<Value, String> {
        let selected: Option<u64> = self
            .shape_selection
            .iter()
            .copied()
            .find(|id| matches!(self.shape(*id).map(|s| s.kind.clone()), Ok(ShapeKind::Chart(_))));
        let chart = selected.or_else(|| {
            let charts = self.charts_on(self.active_sheet);
            (charts.len() == 1).then(|| charts[0])
        });
        match chart {
            Some(id) => Ok(self.part(DrawingPart::Chart(id))),
            None => Err(host_error(91, "there is no active chart")),
        }
    }

    pub(super) fn drawing_part(&self, object: &ObjectRef) -> Option<DrawingPart> {
        match self.objects.get(object.handle as usize) {
            Some(HostObject::Drawing(part)) => Some(*part),
            _ => None,
        }
    }

    fn part(&mut self, part: DrawingPart) -> Value {
        self.object(HostObject::Drawing(part))
    }

    /// The shapes on a sheet, in the order they were made -- which is how
    /// `Shapes(n)` counts them, deletions closing the gaps.
    fn shapes_on(&self, sheet: usize) -> Vec<u64> {
        self.shapes.iter().filter(|shape| shape.sheet == sheet).map(|shape| shape.id).collect()
    }

    fn charts_on(&self, sheet: usize) -> Vec<u64> {
        self.shapes
            .iter()
            .filter(|shape| shape.sheet == sheet && matches!(shape.kind, ShapeKind::Chart(_)))
            .map(|shape| shape.id)
            .collect()
    }

    /// `Shapes(n)` / `Shapes("name")` / `ChartObjects(...)`, among `listed`.
    fn pick_shape(&self, listed: &[u64], index: &Value) -> Result<u64, String> {
        match index {
            Value::String(name) => {
                let alias = localized_alias(name);
                listed
                    .iter()
                    .copied()
                    .find(|id| {
                        self.shape(*id).is_ok_and(|shape| {
                            shape.name.eq_ignore_ascii_case(name)
                                || alias.as_deref().is_some_and(|alias| shape.name.eq_ignore_ascii_case(alias))
                        })
                    })
                    .ok_or_else(|| host_error(NO_SUCH_SHAPE, format!("there is no shape called {name}")))
            }
            value => match any_whole_number(value) {
                Some(number) if number >= 1 && (number as usize) <= listed.len() => {
                    Ok(listed[number as usize - 1])
                }
                _ => Err(host_error(NO_SUCH_SHAPE, "there is no such shape")),
            },
        }
    }

    fn new_shape(&mut self, sheet: usize, kind: ShapeKind, base: &str) -> u64 {
        let number = self.next_shape_number(sheet);
        let id = self.next_shape_id;
        self.next_shape_id += 1;
        let record = ShapeRecord::blank(id, sheet, format!("{base} {number}"), kind);
        self.shapes.push(record);
        id
    }

    fn placed(args: &[Value], from: usize, what: &str) -> Result<(f64, f64, f64, f64), String> {
        let number = |index: usize| -> Result<f64, String> {
            args.get(index)
                .and_then(any_number)
                .ok_or_else(|| format!("{what} takes Left, Top, Width and Height"))
        };
        Ok((number(from)?, number(from + 1)?, number(from + 2)?, number(from + 3)?))
    }

    /// `Shapes.AddShape Type, Left, Top, Width, Height`.
    fn add_shape(&mut self, sheet: usize, args: &[Value]) -> Result<Value, String> {
        let kind = args
            .first()
            .and_then(any_whole_number)
            .ok_or_else(|| "Shapes.AddShape takes a type by number".to_string())?;
        // Measured: 99 is taken (Curved Up Ribbon), -1 is 1004.
        if !(1..=183).contains(&kind) {
            return Err(host_error(1004, "that is not a shape type"));
        }
        let (left, top, width, height) = Self::placed(args, 1, "Shapes.AddShape")?;
        let base = auto_shape_name(kind);
        let id = self.new_shape(sheet, ShapeKind::Auto(kind), base);
        self.place_shape(id, left, top, width, height)?;
        Ok(self.part(DrawingPart::Shape(id)))
    }

    /// `Shapes.AddTextbox Orientation, Left, Top, Width, Height`.
    fn add_textbox(&mut self, sheet: usize, args: &[Value]) -> Result<Value, String> {
        let (left, top, width, height) = Self::placed(args, 1, "Shapes.AddTextbox")?;
        let id = self.new_shape(sheet, ShapeKind::TextBox, "TextBox");
        self.place_shape(id, left, top, width, height)?;
        Ok(self.part(DrawingPart::Shape(id)))
    }

    /// `Shapes.AddLine BeginX, BeginY, EndX, EndY`: the box is the two
    /// points' bounds, and a line drawn back towards the top-left is flipped
    /// both ways -- measured on (110, 300) to (10, 250).
    fn add_line(&mut self, sheet: usize, args: &[Value]) -> Result<Value, String> {
        let (x1, y1, x2, y2) = Self::placed(args, 0, "Shapes.AddLine")?;
        let id = self.new_shape(sheet, ShapeKind::Line, "Straight Connector");
        self.place_shape(id, x1.min(x2), y1.min(y2), (x2 - x1).abs(), (y2 - y1).abs())?;
        let shape = self.shape_mut(id)?;
        shape.flip_h = x2 < x1;
        shape.flip_v = y2 < y1;
        Ok(self.part(DrawingPart::Shape(id)))
    }

    /// `ChartObjects.Add Left, Top, Width, Height` and
    /// `Shapes.AddChart2 Style, XlChartType, Left, Top, Width, Height`: an
    /// empty clustered column chart, without a title, with a legend.
    fn add_chart(&mut self, sheet: usize, left: f64, top: f64, width: f64, height: f64, chart_type: i64) -> Result<u64, String> {
        let chart = ChartRecord {
            chart_type,
            series: Vec::new(),
            has_title: false,
            title: String::new(),
            has_legend: true,
            legend_position: -4152,
            axes: [AxisRecord::default(), AxisRecord::default()],
            style: 201,
            title_auto: true,
        };
        let id = self.new_shape(sheet, ShapeKind::Chart(Box::new(chart)), "Chart");
        self.place_shape(id, left, top, width, height)?;
        Ok(id)
    }

    /// Put a shape where it was asked, and draw it. Measured: a negative
    /// Left or Width becomes 0 without a word.
    fn place_shape(&mut self, id: u64, left: f64, top: f64, width: f64, height: f64) -> Result<(), String> {
        let sheet = {
            let shape = self.shape_mut(id)?;
            shape.left = left.max(0.0);
            shape.top = top.max(0.0);
            shape.width = width.max(0.0);
            shape.height = height.max(0.0);
            shape.sheet
        };
        self.mirror_drawings(sheet);
        Ok(())
    }

    fn redraw(&mut self, id: u64) -> Result<(), String> {
        let sheet = self.shape(id)?.sheet;
        self.mirror_drawings(sheet);
        Ok(())
    }

    fn delete_shape(&mut self, id: u64) -> Result<(), String> {
        let sheet = self.shape(id)?.sheet;
        self.shapes.retain(|shape| shape.id != id);
        self.mirror_drawings(sheet);
        Ok(())
    }

    /// A copy of a shape, twelve points down and across, numbered on:
    /// measured, `Duplicate` of Rectangle 6 at (10, 300) is Rectangle 8 at
    /// (22, 312).
    fn duplicate_shape(&mut self, id: u64, at: Option<(f64, f64)>) -> Result<u64, String> {
        let source = self.shape(id)?.clone();
        let base = match &source.kind {
            ShapeKind::Auto(kind) => auto_shape_name(*kind).to_string(),
            ShapeKind::TextBox => "TextBox".to_string(),
            ShapeKind::Line => "Straight Connector".to_string(),
            ShapeKind::Chart(_) => "Chart".to_string(),
            ShapeKind::Picture => "Picture".to_string(),
            ShapeKind::Other => "Object".to_string(),
        };
        let number = self.next_shape_number(source.sheet);
        let new_id = self.next_shape_id;
        self.next_shape_id += 1;
        let (left, top) = at.unwrap_or((source.left + 12.0, source.top + 12.0));
        let copy = ShapeRecord { id: new_id, name: format!("{base} {number}"), left, top, ..source };
        self.shapes.push(copy);
        self.redraw(new_id)?;
        Ok(new_id)
    }

    /// `Worksheet.Paste` with a shape set aside: a copy lands with its
    /// top-left on the selection. Measured: pasted with D1 selected, the
    /// copy sits at Left 162, Top 0.
    pub(super) fn paste_shape(&mut self, sheet: usize) -> Result<Option<Value>, String> {
        let Some(id) = self.shape_clipboard else {
            return Ok(None);
        };
        if self.shape(id).is_err() {
            return Ok(None);
        }
        let at = self.selection.first();
        let left = (0..at.column).map(|c| f64::from(self.column_px(sheet, c))).sum::<f64>() * 0.75;
        let top = (1..at.row).map(|r| f64::from(self.row_px(sheet, r))).sum::<f64>() * 0.75;
        let source_sheet = self.shape(id)?.sheet;
        let new_id = self.duplicate_shape(id, Some((left, top)))?;
        if source_sheet != sheet {
            let old = source_sheet;
            self.shape_mut(new_id)?.sheet = sheet;
            self.mirror_drawings(old);
            self.mirror_drawings(sheet);
        }
        Ok(Some(Value::Boolean(true)))
    }

    /// `Chart.SetSourceData Source, PlotBy`: the block is read the way
    /// Excel reads it. By columns (the default), a first row of text is the
    /// series' names and a first column of text the categories; by rows the
    /// other way about. Measured: A1:B3 (k v / a 1 / b 2) gives one series
    /// `v` with categories a, b and formula
    /// `=SERIES(Sheet1!$B$1,Sheet1!$A$2:$A$3,Sheet1!$B$2:$B$3,1)`; by rows it
    /// gives `a` and `b` with `=SERIES(Sheet1!$A$2,Sheet1!$B$1,Sheet1!$B$2,1)`;
    /// B2:B3 alone gives `系列1` with `=SERIES(,,Sheet1!$B$2:$B$3,1)`.
    fn set_source_data(&mut self, id: u64, range: CellRange, by_rows: bool) -> Result<(), String> {
        self.settle(range);
        let sheet = range.sheet;
        let is_text = |host: &Self, row: u32, column: u32| {
            host.cell_here(sheet, row, column)
                .is_some_and(|cell| !matches!(cell.value, CellValue::Number(_) | CellValue::Empty))
        };
        let number_at = |host: &Self, row: u32, column: u32| -> f64 {
            match host.cell_here(sheet, row, column).map(|cell| cell.value) {
                Some(CellValue::Number(number)) => number,
                _ => 0.0,
            }
        };
        let text_at = |host: &Self, row: u32, column: u32| -> String {
            host.cell_here(sheet, row, column)
                .map(|cell| shown_text(&from_cell_value(&cell.value), cell.style.number_format.as_deref()))
                .unwrap_or_default()
        };
        let cell_ref = |host: &Self, row: u32, column: u32| -> String {
            host.address_of(CellRange::single(CellAddress { sheet, row, column }))
        };
        let block_ref = |host: &Self, r1: u32, c1: u32, r2: u32, c2: u32| -> String {
            host.address_of(CellRange { sheet, start_row: r1, start_column: c1, end_row: r2, end_column: c2 })
        };
        let (rows, columns) = (
            (range.start_row..=range.end_row).collect::<Vec<u32>>(),
            (range.start_column..=range.end_column).collect::<Vec<u32>>(),
        );
        let mut series = Vec::new();
        if !by_rows {
            let header = rows.len() > 1 && columns.iter().skip(1).any(|c| is_text(self, rows[0], *c))
                || (rows.len() > 1 && columns.len() == 1 && is_text(self, rows[0], columns[0]));
            let data_rows: Vec<u32> = if header { rows[1..].to_vec() } else { rows.clone() };
            let categories = columns.len() > 1 && data_rows.iter().any(|r| is_text(self, *r, columns[0]));
            let value_columns: Vec<u32> = if categories { columns[1..].to_vec() } else { columns.clone() };
            let (first, last) = (data_rows[0], data_rows[data_rows.len() - 1]);
            for column in value_columns {
                let name_ref = header.then(|| cell_ref(self, rows[0], column));
                series.push(SeriesRecord {
                    name: match &name_ref {
                        Some(_) => text_at(self, rows[0], column),
                        None => format!("系列{}", series.len() + 1),
                    },
                    name_ref,
                    values_ref: Some(block_ref(self, first, column, last, column)),
                    x_ref: categories.then(|| block_ref(self, first, columns[0], last, columns[0])),
                    values: data_rows.iter().map(|r| number_at(self, *r, column)).collect(),
                    xs: data_rows
                        .iter()
                        .enumerate()
                        .map(|(at, r)| if categories { text_at(self, *r, columns[0]) } else { (at + 1).to_string() })
                        .collect(),
                    color: None,
                    chart_type: None,
                    has_labels: false,
                });
            }
        } else {
            let header = columns.len() > 1 && rows.iter().skip(1).any(|r| is_text(self, *r, columns[0]));
            let data_columns: Vec<u32> = if header { columns[1..].to_vec() } else { columns.clone() };
            let categories = rows.len() > 1 && data_columns.iter().any(|c| is_text(self, rows[0], *c));
            let value_rows: Vec<u32> = if categories { rows[1..].to_vec() } else { rows.clone() };
            let (first, last) = (data_columns[0], data_columns[data_columns.len() - 1]);
            for row in value_rows {
                let name_ref = header.then(|| cell_ref(self, row, columns[0]));
                series.push(SeriesRecord {
                    name: match &name_ref {
                        Some(_) => text_at(self, row, columns[0]),
                        None => format!("系列{}", series.len() + 1),
                    },
                    name_ref,
                    values_ref: Some(block_ref(self, row, first, row, last)),
                    x_ref: categories.then(|| block_ref(self, rows[0], first, rows[0], last)),
                    values: data_columns.iter().map(|c| number_at(self, row, *c)).collect(),
                    xs: data_columns
                        .iter()
                        .enumerate()
                        .map(|(at, c)| if categories { text_at(self, rows[0], *c) } else { (at + 1).to_string() })
                        .collect(),
                    color: None,
                    chart_type: None,
                    has_labels: false,
                });
            }
        }
        let chart = self.chart_mut(id)?;
        chart.series = series;
        self.redraw(id)
    }

    /// `=SERIES(name, categories, values, order)`, the way Excel writes a
    /// series' formula: a name set as text is quoted, a part not given is
    /// left empty.
    fn series_formula(&self, id: u64, number: usize) -> Result<String, String> {
        let series = self.series(id, number)?;
        let name = match (&series.name_ref, series.name.is_empty()) {
            (Some(reference), _) => reference.clone(),
            (None, false) if !series.name.starts_with("系列") => format!("\"{}\"", series.name),
            _ => String::new(),
        };
        Ok(format!(
            "=SERIES({name},{},{},{number})",
            series.x_ref.clone().unwrap_or_default(),
            series.values_ref.clone().unwrap_or_default()
        ))
    }

    fn cells_of_range_object(&self, value: &Value) -> Option<CellRange> {
        match value {
            Value::Object(object) => self.range(object),
            _ => None,
        }
    }

    fn range_numbers(&mut self, range: CellRange) -> Vec<f64> {
        self.settle(range);
        range
            .addresses()
            .map(|at| match self.cell_here(at.sheet, at.row, at.column).map(|cell| cell.value) {
                Some(CellValue::Number(number)) => number,
                _ => 0.0,
            })
            .collect()
    }

    fn range_texts(&mut self, range: CellRange) -> Vec<String> {
        self.settle(range);
        range
            .addresses()
            .map(|at| {
                self.cell_here(at.sheet, at.row, at.column)
                    .map(|cell| shown_text(&from_cell_value(&cell.value), cell.style.number_format.as_deref()))
                    .unwrap_or_default()
            })
            .collect()
    }

    fn variant_row(values: Vec<Value>) -> Value {
        Value::Array(ArrayValue {
            dimensions: vec![ArrayDimension { lower_bound: 1, length: values.len() }],
            values,
            element_default: Box::new(Value::Empty),
            resizable: true,
        })
    }

    // ---- reading ------------------------------------------------------------

    pub(super) fn drawing_get(&mut self, part: DrawingPart, name: &str) -> Result<Option<Value>, String> {
        let lower = name.to_ascii_lowercase();
        match part {
            DrawingPart::Shapes(sheet) => Ok(match lower.as_str() {
                "count" => Some(Value::Integer(self.shapes_on(sheet).len() as i64)),
                "parent" => Some(self.object(HostObject::Worksheet(sheet))),
                _ => None,
            }),
            DrawingPart::ChartObjects(sheet) => Ok(match lower.as_str() {
                "count" => Some(Value::Integer(self.charts_on(sheet).len() as i64)),
                "parent" => Some(self.object(HostObject::Worksheet(sheet))),
                _ => None,
            }),
            DrawingPart::Shape(id) | DrawingPart::ChartObject(id) => self.shape_get(part, id, &lower),
            DrawingPart::Fill(id) => {
                let shape = self.shape(id)?.clone();
                Ok(match lower.as_str() {
                    "forecolor" => Some(self.part(DrawingPart::FillColor(id))),
                    "visible" => Some(mso(shape.fill_visible)),
                    "transparency" => Some(Value::Double(shape.transparency)),
                    _ => None,
                })
            }
            DrawingPart::FillColor(id) => {
                let shape = self.shape(id)?;
                Ok(match lower.as_str() {
                    "rgb" => Some(Value::Integer(shape.fill)),
                    "objectthemecolor" => Some(Value::Integer(shape.fill_theme.map_or(0, |theme| theme as i64))),
                    _ => None,
                })
            }
            DrawingPart::Line(id) => {
                let shape = self.shape(id)?.clone();
                Ok(match lower.as_str() {
                    "forecolor" => Some(self.part(DrawingPart::LineColor(id))),
                    "visible" => Some(mso(shape.line_visible)),
                    "weight" => Some(Value::Double(shape.line_weight)),
                    "dashstyle" => Some(Value::Integer(shape.dash)),
                    "endarrowheadstyle" => Some(Value::Integer(shape.arrow_end)),
                    "beginarrowheadstyle" => Some(Value::Integer(1)),
                    _ => None,
                })
            }
            DrawingPart::LineColor(id) => {
                let shape = self.shape(id)?;
                Ok(match lower.as_str() {
                    "rgb" => Some(Value::Integer(shape.line)),
                    "objectthemecolor" => Some(Value::Integer(shape.line_theme.map_or(0, |theme| theme as i64))),
                    _ => None,
                })
            }
            DrawingPart::TextFrame(id) => {
                let shape = self.shape(id)?.clone();
                Ok(match lower.as_str() {
                    "characters" => Some(self.part(DrawingPart::Characters(id, 1, None))),
                    "horizontalalignment" => Some(Value::Integer(shape.h_align)),
                    "verticalalignment" => Some(Value::Integer(shape.v_align)),
                    "autosize" => Some(Value::Boolean(shape.auto_size)),
                    "marginleft" => Some(Value::Double(shape.margins.0)),
                    "margintop" => Some(Value::Double(shape.margins.1)),
                    "marginright" => Some(Value::Double(shape.margins.2)),
                    "marginbottom" => Some(Value::Double(shape.margins.3)),
                    "parent" => Some(self.part(DrawingPart::Shape(id))),
                    _ => None,
                })
            }
            DrawingPart::TextFrame2(id) => {
                let shape = self.shape(id)?.clone();
                Ok(match lower.as_str() {
                    "textrange" => Some(self.part(DrawingPart::TextRange(id))),
                    "hastext" => Some(mso(!shape.text().is_empty())),
                    // msoAnchorTop 1, Middle 3, Bottom 4.
                    "verticalanchor" => Some(Value::Integer(match shape.v_align {
                        -4108 => 3,
                        -4107 => 4,
                        _ => 1,
                    })),
                    "wordwrap" => Some(mso(true)),
                    "autosize" => Some(Value::Integer(i64::from(shape.auto_size))),
                    "marginleft" => Some(Value::Double(shape.margins.0)),
                    "margintop" => Some(Value::Double(shape.margins.1)),
                    "marginright" => Some(Value::Double(shape.margins.2)),
                    "marginbottom" => Some(Value::Double(shape.margins.3)),
                    _ => None,
                })
            }
            DrawingPart::TextRange(id) => {
                let shape = self.shape(id)?.clone();
                Ok(match lower.as_str() {
                    "text" => Some(Value::String(shape.text())),
                    "font" => Some(self.part(DrawingPart::CharactersFont(id, 1, None))),
                    "paragraphformat" => Some(self.part(DrawingPart::ParagraphFormat(id))),
                    "characters" => Some(self.part(DrawingPart::Characters(id, 1, None))),
                    "paragraphs" => Some(self.part(DrawingPart::Paragraphs(id))),
                    "length" | "count" => Some(Value::Integer(shape.text().chars().count() as i64)),
                    _ => None,
                })
            }
            DrawingPart::ParagraphFormat(id) => {
                let shape = self.shape(id)?;
                Ok(match lower.as_str() {
                    // msoAlignLeft 1, Center 2, Right 3.
                    "alignment" => Some(Value::Integer(match shape.h_align {
                        -4108 => 2,
                        -4152 => 3,
                        _ => 1,
                    })),
                    _ => None,
                })
            }
            DrawingPart::Characters(id, start, length) => {
                let shape = self.shape(id)?.clone();
                let text: String = shape
                    .text()
                    .chars()
                    .skip(start.max(1) as usize - 1)
                    .take(length.map_or(usize::MAX, |length| length as usize))
                    .collect();
                Ok(match lower.as_str() {
                    "text" | "caption" => Some(Value::String(text)),
                    "count" => Some(Value::Integer(text.chars().count() as i64)),
                    "font" => Some(self.part(DrawingPart::CharactersFont(id, start, length))),
                    _ => None,
                })
            }
            DrawingPart::CharactersFont(id, start, length) => {
                let shape = self.shape(id)?.clone();
                let answer = |value: Option<Value>| Some(value.unwrap_or(Value::Null));
                Ok(match lower.as_str() {
                    "name" => answer(shape.uniform_style(start, length, |s| s.name.clone()).map(Value::String)),
                    "size" => answer(shape.uniform_style(start, length, |s| s.size).map(Value::Double)),
                    "bold" => answer(shape.uniform_style(start, length, |s| s.bold).map(Value::Boolean)),
                    "italic" => answer(shape.uniform_style(start, length, |s| s.italic).map(Value::Boolean)),
                    "underline" => answer(shape.uniform_style(start, length, |s| s.underline).map(|u| {
                        Value::Integer(if u { UNDERLINE_SINGLE } else { UNDERLINE_NONE })
                    })),
                    "color" => answer(shape.uniform_style(start, length, |s| s.color).map(Value::Integer)),
                    "colorindex" => answer(
                        shape.uniform_style(start, length, |s| s.color).map(|c| Value::Integer(nearest_palette_index(c))),
                    ),
                    // `TextRange.Font.Fill.ForeColor.RGB` reaches the colour
                    // the long way round.
                    "fill" => Some(self.part(DrawingPart::CharactersFont(id, start, length))),
                    "forecolor" => Some(self.part(DrawingPart::CharactersFont(id, start, length))),
                    "rgb" => answer(shape.uniform_style(start, length, |s| s.color).map(Value::Integer)),
                    // Measured: a shape's white text is msoThemeColorLight1
                    // (2), a box's black text Dark1 (1).
                    "objectthemecolor" => Some(Value::Integer(
                        if shape.uniform_style(start, length, |s| s.color) == Some(WHITE) { 2 } else { 1 },
                    )),
                    _ => None,
                })
            }
            DrawingPart::Chart(id) => self.chart_get(id, &lower),
            DrawingPart::SeriesCollection(id) => Ok(match lower.as_str() {
                "count" => Some(Value::Integer(self.chart(id)?.series.len() as i64)),
                "parent" => Some(self.part(DrawingPart::Chart(id))),
                _ => None,
            }),
            DrawingPart::Series(id, number) => {
                let series = self.series(id, number)?.clone();
                Ok(match lower.as_str() {
                    "name" => Some(Value::String(series.name)),
                    "formula" => Some(Value::String(self.series_formula(id, number)?)),
                    "values" => Some(Self::variant_row(series.values.iter().map(|v| Value::Double(*v)).collect())),
                    "xvalues" => Some(Self::variant_row(
                        series.xs.iter().map(|x| match x.parse::<f64>() {
                            Ok(number) => Value::Double(number),
                            Err(_) => Value::String(x.clone()),
                        }).collect(),
                    )),
                    "points" => Some(self.part(DrawingPart::Points(id, number))),
                    "hasdatalabels" => Some(Value::Boolean(series.has_labels)),
                    "datalabels" => Some(self.part(DrawingPart::DataLabels(id, number))),
                    "charttype" => Some(Value::Integer(series.chart_type.unwrap_or(self.chart(id)?.chart_type))),
                    "axisgroup" => Some(Value::Integer(1)),
                    "plotorder" => Some(Value::Integer(number as i64)),
                    "format" => Some(self.part(DrawingPart::SeriesFormat(id, number))),
                    "interior" => Some(self.part(DrawingPart::SeriesColor(id, number))),
                    "border" => Some(self.part(DrawingPart::SeriesColor(id, number))),
                    "parent" => Some(self.part(DrawingPart::Chart(id))),
                    _ => None,
                })
            }
            DrawingPart::SeriesFormat(id, number) => Ok(match lower.as_str() {
                "fill" => Some(self.part(DrawingPart::SeriesFill(id, number))),
                "line" => Some(self.part(DrawingPart::SeriesFill(id, number))),
                _ => None,
            }),
            DrawingPart::SeriesFill(id, number) => Ok(match lower.as_str() {
                "forecolor" => Some(self.part(DrawingPart::SeriesColor(id, number))),
                "visible" => Some(mso(true)),
                _ => None,
            }),
            DrawingPart::SeriesColor(id, number) => {
                let series = self.series(id, number)?;
                Ok(match lower.as_str() {
                    "rgb" | "color" => Some(Value::Integer(series.color.unwrap_or(THEME_COLOURS[ACCENT1_INDEX]))),
                    "colorindex" => Some(Value::Integer(nearest_palette_index(series.color.unwrap_or(THEME_COLOURS[ACCENT1_INDEX])))),
                    _ => None,
                })
            }
            DrawingPart::Points(id, number) => Ok(match lower.as_str() {
                "count" => Some(Value::Integer(self.series(id, number)?.values.len() as i64)),
                _ => None,
            }),
            DrawingPart::DataLabels(id, number) => {
                let series = self.series(id, number)?;
                Ok(match lower.as_str() {
                    "count" => Some(Value::Integer(if series.has_labels { series.values.len() as i64 } else { 0 })),
                    "showvalue" => Some(Value::Boolean(series.has_labels)),
                    _ => None,
                })
            }
            DrawingPart::ChartTitle(id) => {
                let chart = self.chart(id)?.clone();
                let shown = if chart.title.is_empty() {
                    match chart.series.as_slice() {
                        [only] => only.name.clone(),
                        _ => "グラフ タイトル".to_string(),
                    }
                } else {
                    chart.title
                };
                Ok(match lower.as_str() {
                    "text" | "caption" => Some(Value::String(shown)),
                    "characters" => Some(self.part(DrawingPart::ChartTitle(id))),
                    _ => None,
                })
            }
            DrawingPart::Legend(id) => Ok(match lower.as_str() {
                "position" => Some(Value::Integer(self.chart(id)?.legend_position)),
                _ => None,
            }),
            DrawingPart::Axes(id) => Ok(match lower.as_str() {
                "count" => Some(Value::Integer(2)),
                "parent" => Some(self.part(DrawingPart::Chart(id))),
                _ => None,
            }),
            DrawingPart::Axis(id, which) => {
                let chart = self.chart(id)?.clone();
                let axis = &chart.axes[usize::from(which == 2)];
                Ok(match lower.as_str() {
                    "hastitle" => Some(Value::Boolean(axis.has_title)),
                    "axistitle" => {
                        if !axis.has_title {
                            return Err(host_error(-2_147_467_259, "the axis has no title"));
                        }
                        Some(self.part(DrawingPart::AxisTitle(id, which)))
                    }
                    "minimumscale" => Some(Value::Double(axis.min.unwrap_or(0.0))),
                    "maximumscale" => Some(Value::Double(axis.max.unwrap_or(0.0))),
                    "minimumscaleisauto" => Some(Value::Boolean(axis.min.is_none())),
                    "maximumscaleisauto" => Some(Value::Boolean(axis.max.is_none())),
                    "categorynames" if which == 1 => Some(Self::variant_row(
                        chart.series.first().map(|s| s.xs.iter().map(|x| Value::String(x.clone())).collect()).unwrap_or_default(),
                    )),
                    "type" => Some(Value::Integer(which)),
                    _ => None,
                })
            }
            DrawingPart::AxisTitle(id, which) => {
                let chart = self.chart(id)?;
                Ok(match lower.as_str() {
                    "text" | "caption" => Some(Value::String(chart.axes[usize::from(which == 2)].title.clone())),
                    "characters" => Some(self.part(DrawingPart::AxisTitle(id, which))),
                    _ => None,
                })
            }
            // Measured: the chart area is the object's box less five points
            // each way.
            DrawingPart::ChartArea(id) => {
                let shape = self.shape(id)?;
                Ok(match lower.as_str() {
                    "width" => Some(Value::Double(shape.width - 5.0)),
                    "height" => Some(Value::Double(shape.height - 5.0)),
                    "format" => Some(self.part(DrawingPart::ChartArea(id))),
                    "fill" | "interior" | "forecolor" => Some(self.part(DrawingPart::ChartArea(id))),
                    "rgb" | "color" => Some(Value::Integer(WHITE)),
                    _ => None,
                })
            }
            // Measured on a 200 by 100 chart: the plot's inside left is 16.8.
            // The rest of the plot's frame is not measured and stands in.
            DrawingPart::PlotArea(id) => {
                let shape = self.shape(id)?;
                Ok(match lower.as_str() {
                    "format" | "fill" | "interior" | "forecolor" => Some(self.part(DrawingPart::PlotArea(id))),
                    "rgb" | "color" => Some(Value::Integer(WHITE)),
                    "insideleft" => Some(Value::Double(16.8025196850394)),
                    "insidetop" => Some(Value::Double(7.2)),
                    "insidewidth" => Some(Value::Double((shape.width - 40.0).max(0.0))),
                    "insideheight" => Some(Value::Double((shape.height - 40.0).max(0.0))),
                    "left" => Some(Value::Double(5.0)),
                    "top" => Some(Value::Double(5.0)),
                    "width" => Some(Value::Double((shape.width - 15.0).max(0.0))),
                    "height" => Some(Value::Double((shape.height - 15.0).max(0.0))),
                    _ => None,
                })
            }
            DrawingPart::Adjustments(id) => Ok(match lower.as_str() {
                "count" => Some(Value::Integer(self.shape(id)?.adjusts.len() as i64)),
                _ => None,
            }),
            // Measured: "a" & Chr(10) & "b" & vbCr & "c" is two paragraphs --
            // a line feed breaks a line, a carriage return a paragraph; and
            // both read back as line feeds.
            DrawingPart::Paragraphs(id) => Ok(match lower.as_str() {
                "count" => Some(Value::Integer(self.shape(id)?.paragraph_count as i64)),
                "text" => Some(Value::String(self.shape(id)?.text())),
                _ => None,
            }),
            DrawingPart::ShapeRange(index) => {
                let ids = self.shape_range_ids(index);
                match lower.as_str() {
                    "count" => Ok(Some(Value::Integer(ids.len() as i64))),
                    "shaperange" => Ok(Some(self.part(part))),
                    _ => match ids.first() {
                        Some(id) => self.drawing_get(DrawingPart::Shape(*id), name),
                        None => Ok(None),
                    },
                }
            }
        }
    }

    fn shape_get(&mut self, part: DrawingPart, id: u64, lower: &str) -> Result<Option<Value>, String> {
        let shape = self.shape(id)?.clone();
        Ok(match lower {
            "name" => Some(Value::String(shape.name)),
            "left" => Some(Value::Double(shape.left)),
            "top" => Some(Value::Double(shape.top)),
            "width" => Some(Value::Double(shape.width)),
            "height" => Some(Value::Double(shape.height)),
            "rotation" => Some(Value::Double(shape.rotation)),
            "visible" => Some(mso(shape.visible)),
            "placement" => Some(Value::Integer(shape.placement)),
            "onaction" => Some(Value::String(shape.on_action)),
            "alternativetext" => Some(Value::String(shape.alt_text)),
            "lockaspectratio" => Some(mso(shape.lock_aspect)),
            "horizontalflip" => Some(mso(shape.flip_h)),
            "verticalflip" => Some(mso(shape.flip_v)),
            "id" => Some(Value::Integer(shape.id as i64 + 1)),
            "zorderposition" => Some(Value::Integer(
                self.shapes_on(shape.sheet).iter().position(|held| *held == id).map_or(1, |at| at as i64 + 1),
            )),
            // msoAutoShape 1, msoChart 3, msoLine 9, msoPicture 13, msoTextBox 17.
            "type" => Some(Value::Integer(match shape.kind {
                ShapeKind::Auto(_) => 1,
                ShapeKind::TextBox => 17,
                ShapeKind::Line => 9,
                ShapeKind::Chart(_) => 3,
                ShapeKind::Picture => 13,
                ShapeKind::Other => 24,
            })),
            // Measured: a connector's AutoShapeType is msoShapeMixed, -2.
            "autoshapetype" => Some(Value::Integer(match shape.kind {
                ShapeKind::Auto(kind) => kind,
                ShapeKind::TextBox => 1,
                _ => -2,
            })),
            "haschart" => Some(mso(matches!(shape.kind, ShapeKind::Chart(_)))),
            "adjustments" => Some(self.part(DrawingPart::Adjustments(id))),
            "fill" => Some(self.part(DrawingPart::Fill(id))),
            "line" | "border" => Some(self.part(DrawingPart::Line(id))),
            "textframe" => Some(self.part(DrawingPart::TextFrame(id))),
            "textframe2" => Some(self.part(DrawingPart::TextFrame2(id))),
            "chart" => {
                if !matches!(shape.kind, ShapeKind::Chart(_)) {
                    return Err(host_error(1004, "the shape holds no chart"));
                }
                Some(self.part(DrawingPart::Chart(id)))
            }
            "topleftcell" => {
                let at = self.cell_under(shape.sheet, shape.left, shape.top);
                Some(self.object(HostObject::Range(CellRange::single(at))))
            }
            "bottomrightcell" => {
                let at = self.cell_under(shape.sheet, shape.left + shape.width, shape.top + shape.height);
                Some(self.object(HostObject::Range(CellRange::single(at))))
            }
            "parent" => Some(self.object(HostObject::Worksheet(shape.sheet))),
            "shaperange" => Some(self.part(part)),
            "index" => Some(Value::Integer(
                self.shapes_on(shape.sheet).iter().position(|held| *held == id).map_or(1, |at| at as i64 + 1),
            )),
            _ => None,
        })
    }

    fn chart_get(&mut self, id: u64, lower: &str) -> Result<Option<Value>, String> {
        let shape = self.shape(id)?.clone();
        let chart = self.chart(id)?.clone();
        Ok(match lower {
            "charttype" => Some(Value::Integer(chart.chart_type)),
            "hastitle" => Some(Value::Boolean(chart.has_title)),
            // Measured: `ChartTitle.Text` reads the one series' name on a
            // chart whose HasTitle is False.
            "charttitle" => Some(self.part(DrawingPart::ChartTitle(id))),
            "haslegend" => Some(Value::Boolean(chart.has_legend)),
            "legend" => Some(self.part(DrawingPart::Legend(id))),
            "seriescollection" => Some(self.part(DrawingPart::SeriesCollection(id))),
            "axes" => Some(self.part(DrawingPart::Axes(id))),
            "chartarea" => Some(self.part(DrawingPart::ChartArea(id))),
            "plotarea" => Some(self.part(DrawingPart::PlotArea(id))),
            "chartstyle" => Some(Value::Integer(chart.style)),
            // Measured: the chart is `Sheet1 グラフ 5` for the object `Chart 5`
            // -- the sheet, the Japanese word, and the number.
            "name" => Some(Value::String(format!(
                "{} グラフ {}",
                self.workbook.sheets[shape.sheet].name,
                shape.name.rsplit(' ').next().unwrap_or("")
            ))),
            "parent" => Some(self.part(DrawingPart::ChartObject(id))),
            "visible" => Some(mso(shape.visible)),
            _ => None,
        })
    }

    // ---- writing ------------------------------------------------------------

    pub(super) fn drawing_set(&mut self, part: DrawingPart, name: &str, value: Value) -> Result<bool, String> {
        let lower = name.to_ascii_lowercase();
        match part {
            DrawingPart::Shape(id) | DrawingPart::ChartObject(id) => {
                let taken = {
                    let shape = self.shape_mut(id)?;
                    match lower.as_str() {
                        "name" => {
                            let Value::String(named) = &value else {
                                return Err("Shape.Name takes text".to_string());
                            };
                            shape.name = named.clone();
                            true
                        }
                        "left" | "top" | "width" | "height" | "rotation" => {
                            let Some(number) = any_number(&value) else {
                                return Err(format!("Shape.{name} takes a number"));
                            };
                            match lower.as_str() {
                                "left" => shape.left = number.max(0.0),
                                "top" => shape.top = number.max(0.0),
                                "width" => shape.width = number.max(0.0),
                                "height" => shape.height = number.max(0.0),
                                _ => shape.rotation = number,
                            }
                            true
                        }
                        "visible" => {
                            shape.visible = mso_asked(&value, "Shape.Visible")?;
                            true
                        }
                        "placement" => {
                            shape.placement = any_whole_number(&value).unwrap_or(1);
                            true
                        }
                        "onaction" => {
                            shape.on_action = match &value {
                                Value::String(text) => text.clone(),
                                _ => String::new(),
                            };
                            true
                        }
                        "alternativetext" => {
                            shape.alt_text = match &value {
                                Value::String(text) => text.clone(),
                                _ => String::new(),
                            };
                            true
                        }
                        "lockaspectratio" => {
                            shape.lock_aspect = mso_asked(&value, "Shape.LockAspectRatio")?;
                            true
                        }
                        _ => false,
                    }
                };
                if taken {
                    self.redraw(id)?;
                }
                Ok(taken)
            }
            DrawingPart::Fill(id) => {
                let taken = {
                    let shape = self.shape_mut(id)?;
                    match lower.as_str() {
                        "visible" => {
                            shape.fill_visible = mso_asked(&value, "FillFormat.Visible")?;
                            true
                        }
                        "transparency" => {
                            shape.transparency = any_number(&value).unwrap_or(0.0).clamp(0.0, 1.0);
                            true
                        }
                        _ => false,
                    }
                };
                if taken {
                    self.redraw(id)?;
                }
                Ok(taken)
            }
            DrawingPart::FillColor(id) | DrawingPart::LineColor(id) => {
                let fill = matches!(part, DrawingPart::FillColor(_));
                let taken = {
                    let shape = self.shape_mut(id)?;
                    match lower.as_str() {
                        "rgb" => {
                            let Some(colour) = color_number(&value, "ColorFormat.RGB")? else {
                                return Ok(true);
                            };
                            if fill {
                                shape.fill = colour as i64;
                                shape.fill_theme = None;
                                shape.fill_visible = true;
                            } else {
                                shape.line = colour as i64;
                                shape.line_theme = None;
                                shape.line_visible = true;
                            }
                            true
                        }
                        "objectthemecolor" => {
                            let theme = theme_colour(&value)?;
                            if fill {
                                shape.fill = THEME_COLOURS[theme - 1];
                                shape.fill_theme = Some(theme);
                            } else {
                                shape.line = THEME_COLOURS[theme - 1];
                                shape.line_theme = Some(theme);
                            }
                            true
                        }
                        "schemecolor" => true,
                        _ => false,
                    }
                };
                if taken {
                    self.redraw(id)?;
                }
                Ok(taken)
            }
            DrawingPart::Line(id) => {
                let taken = {
                    let shape = self.shape_mut(id)?;
                    match lower.as_str() {
                        "visible" => {
                            shape.line_visible = mso_asked(&value, "LineFormat.Visible")?;
                            true
                        }
                        "weight" => {
                            shape.line_weight = any_number(&value).unwrap_or(0.75).max(0.0);
                            shape.line_visible = true;
                            true
                        }
                        "dashstyle" => {
                            shape.dash = any_whole_number(&value).unwrap_or(1);
                            true
                        }
                        "endarrowheadstyle" => {
                            shape.arrow_end = any_whole_number(&value).unwrap_or(1);
                            true
                        }
                        "beginarrowheadstyle" | "style" | "transparency" => true,
                        _ => false,
                    }
                };
                if taken {
                    self.redraw(id)?;
                }
                Ok(taken)
            }
            DrawingPart::TextFrame(id) => {
                let taken = {
                    let shape = self.shape_mut(id)?;
                    match lower.as_str() {
                        "horizontalalignment" => {
                            shape.h_align = any_whole_number(&value).unwrap_or(-4131);
                            true
                        }
                        "verticalalignment" => {
                            shape.v_align = any_whole_number(&value).unwrap_or(-4160);
                            true
                        }
                        "autosize" => {
                            shape.auto_size = style_face_boolean(&value, "TextFrame.AutoSize")?.unwrap_or(false);
                            if shape.auto_size {
                                shape.fit_to_text();
                            }
                            true
                        }
                        "marginleft" => { shape.margins.0 = any_number(&value).unwrap_or(7.2); true }
                        "margintop" => { shape.margins.1 = any_number(&value).unwrap_or(3.6); true }
                        "marginright" => { shape.margins.2 = any_number(&value).unwrap_or(7.2); true }
                        "marginbottom" => { shape.margins.3 = any_number(&value).unwrap_or(3.6); true }
                        _ => false,
                    }
                };
                if taken {
                    self.redraw(id)?;
                }
                Ok(taken)
            }
            DrawingPart::TextFrame2(id) => {
                let taken = {
                    let shape = self.shape_mut(id)?;
                    match lower.as_str() {
                        "verticalanchor" => {
                            shape.v_align = match any_whole_number(&value) {
                                Some(3) => -4108,
                                Some(4) => -4107,
                                _ => -4160,
                            };
                            true
                        }
                        "wordwrap" | "autosize" => true,
                        "marginleft" => { shape.margins.0 = any_number(&value).unwrap_or(7.2); true }
                        "margintop" => { shape.margins.1 = any_number(&value).unwrap_or(3.6); true }
                        "marginright" => { shape.margins.2 = any_number(&value).unwrap_or(7.2); true }
                        "marginbottom" => { shape.margins.3 = any_number(&value).unwrap_or(3.6); true }
                        _ => false,
                    }
                };
                if taken {
                    self.redraw(id)?;
                }
                Ok(taken)
            }
            DrawingPart::ParagraphFormat(id) => {
                if lower == "alignment" {
                    let shape = self.shape_mut(id)?;
                    shape.h_align = match any_whole_number(&value) {
                        Some(2) => -4108,
                        Some(3) => -4152,
                        _ => -4131,
                    };
                    self.redraw(id)?;
                    return Ok(true);
                }
                Ok(false)
            }
            DrawingPart::TextRange(id) | DrawingPart::Characters(id, ..) => {
                if lower == "text" || lower == "caption" {
                    let text = match &value {
                        Value::String(text) => text.clone(),
                        other => text_of_value(other),
                    };
                    self.shape_mut(id)?.set_text(&text);
                    self.redraw(id)?;
                    return Ok(true);
                }
                Ok(false)
            }
            DrawingPart::CharactersFont(id, start, length) => {
                let taken = {
                    let shape = self.shape_mut(id)?;
                    if shape.runs.is_empty() {
                        // Dress the text a macro will write: the style is kept
                        // on an empty run until the text arrives.
                        let style = shape.default_style();
                        shape.runs.push(ShapeRunRecord { text: String::new(), style });
                    }
                    let (first, last) = shape.split_runs(start, length);
                    let mut taken = true;
                    for run in &mut shape.runs[first..last] {
                        match lower.as_str() {
                            "name" => {
                                if let Value::String(named) = &value {
                                    run.style.name = named.clone();
                                }
                            }
                            "size" => run.style.size = any_number(&value).unwrap_or(run.style.size),
                            "bold" => run.style.bold = style_face_boolean(&value, "Font.Bold")?.unwrap_or(false),
                            "italic" => run.style.italic = style_face_boolean(&value, "Font.Italic")?.unwrap_or(false),
                            "underline" => {
                                run.style.underline = match &value {
                                    Value::Boolean(flag) => *flag,
                                    other => any_whole_number(other).is_some_and(|n| n != UNDERLINE_NONE && n != 0),
                                }
                            }
                            "color" | "rgb" => {
                                if let Some(colour) = color_number(&value, "Font.Color")? {
                                    run.style.color = colour as i64;
                                }
                            }
                            "colorindex" => {
                                if let Some(colour) = palette_choice(&value, COLOUR_AUTOMATIC, "Font.ColorIndex")? {
                                    run.style.color = colour_to_packed(Some(&colour)).unwrap_or(0);
                                }
                            }
                            "objectthemecolor" => {
                                let theme = theme_colour(&value)?;
                                run.style.color = THEME_COLOURS[theme - 1];
                            }
                            _ => taken = false,
                        }
                    }
                    taken
                };
                if taken {
                    self.redraw(id)?;
                }
                Ok(taken)
            }
            DrawingPart::Chart(id) => {
                let taken = {
                    let chart = self.chart_mut(id)?;
                    match lower.as_str() {
                        "charttype" => {
                            let Some(kind) = any_whole_number(&value) else {
                                return Err("Chart.ChartType takes an xlChartType".to_string());
                            };
                            chart.chart_type = kind;
                            for series in &mut chart.series {
                                series.chart_type = None;
                            }
                            if chart.title_auto && chart_is_pie(kind) && chart.series.len() == 1 {
                                chart.has_title = true;
                            }
                            chart.title_auto = false;
                            true
                        }
                        "hastitle" => {
                            chart.has_title = style_face_boolean(&value, "Chart.HasTitle")?.unwrap_or(false);
                            if chart.has_title && chart.title.is_empty() {
                                chart.title = match chart.series.as_slice() {
                                    [only] => only.name.clone(),
                                    _ => "グラフ タイトル".to_string(),
                                };
                            }
                            true
                        }
                        "haslegend" => {
                            chart.has_legend = style_face_boolean(&value, "Chart.HasLegend")?.unwrap_or(false);
                            true
                        }
                        "chartstyle" => {
                            chart.style = any_whole_number(&value).unwrap_or(chart.style);
                            true
                        }
                        "name" => true,
                        _ => false,
                    }
                };
                if taken {
                    self.redraw(id)?;
                }
                Ok(taken)
            }
            DrawingPart::Series(id, number) => {
                let taken = match lower.as_str() {
                    "name" => {
                        let series = self.series_mut(id, number)?;
                        match &value {
                            // `="S2"` is a name written as a formula.
                            Value::String(text) => {
                                let bare = text.strip_prefix('=').unwrap_or(text);
                                series.name = bare.trim_matches('"').to_string();
                                series.name_ref = None;
                            }
                            Value::Object(_) => {
                                if let Some(range) = self.cells_of_range_object(&value) {
                                    let text = self.range_texts(range).into_iter().next().unwrap_or_default();
                                    let reference = self.address_of(range);
                                    let series = self.series_mut(id, number)?;
                                    series.name = text;
                                    series.name_ref = Some(reference);
                                }
                            }
                            other => {
                                let text = text_of_value(other);
                                self.series_mut(id, number)?.name = text;
                            }
                        }
                        true
                    }
                    "values" => {
                        if let Some(range) = self.cells_of_range_object(&value) {
                            let numbers = self.range_numbers(range);
                            let reference = self.address_of(range);
                            let series = self.series_mut(id, number)?;
                            series.values = numbers;
                            series.values_ref = Some(reference);
                            if series.xs.len() != series.values.len() {
                                series.xs = (1..=series.values.len()).map(|n| n.to_string()).collect();
                            }
                        } else if let Value::Array(listed) = &value {
                            let series = self.series_mut(id, number)?;
                            series.values = listed.values.iter().map(|v| any_number(v).unwrap_or(0.0)).collect();
                            series.values_ref = None;
                        }
                        true
                    }
                    "xvalues" => {
                        if let Some(range) = self.cells_of_range_object(&value) {
                            let texts = self.range_texts(range);
                            let reference = self.address_of(range);
                            let series = self.series_mut(id, number)?;
                            series.xs = texts;
                            series.x_ref = Some(reference);
                        } else if let Value::Array(listed) = &value {
                            let series = self.series_mut(id, number)?;
                            series.xs = listed.values.iter().map(text_of_value).collect();
                            series.x_ref = None;
                        }
                        true
                    }
                    "hasdatalabels" => {
                        let flag = style_face_boolean(&value, "Series.HasDataLabels")?.unwrap_or(false);
                        self.series_mut(id, number)?.has_labels = flag;
                        true
                    }
                    // Measured: one series' ChartType is the chart's when it
                    // is the only one.
                    "charttype" => {
                        let kind = any_whole_number(&value).unwrap_or(51);
                        let single = self.chart(id)?.series.len() == 1;
                        self.series_mut(id, number)?.chart_type = Some(kind);
                        if single {
                            self.chart_mut(id)?.chart_type = kind;
                        }
                        true
                    }
                    "axisgroup" | "plotorder" | "smooth" | "markerstyle" | "markersize" => true,
                    _ => false,
                };
                if taken {
                    self.redraw(id)?;
                }
                Ok(taken)
            }
            DrawingPart::SeriesColor(id, number) => {
                if lower == "rgb" || lower == "color" {
                    if let Some(colour) = color_number(&value, "ColorFormat.RGB")? {
                        self.series_mut(id, number)?.color = Some(colour as i64);
                        self.redraw(id)?;
                    }
                    return Ok(true);
                }
                if lower == "colorindex" {
                    if let Some(colour) = palette_choice(&value, COLOUR_AUTOMATIC, "Interior.ColorIndex")? {
                        self.series_mut(id, number)?.color = colour_to_packed(Some(&colour));
                        self.redraw(id)?;
                    }
                    return Ok(true);
                }
                Ok(false)
            }
            DrawingPart::SeriesFill(..) | DrawingPart::SeriesFormat(..) => {
                Ok(matches!(lower.as_str(), "visible" | "transparency" | "weight"))
            }
            DrawingPart::DataLabels(id, number) => {
                if matches!(lower.as_str(), "showvalue" | "showcategoryname" | "showpercentage" | "position" | "numberformat") {
                    if lower == "showvalue" {
                        let flag = style_face_boolean(&value, "DataLabels.ShowValue")?.unwrap_or(false);
                        self.series_mut(id, number)?.has_labels = flag;
                        self.redraw(id)?;
                    }
                    return Ok(true);
                }
                Ok(false)
            }
            DrawingPart::ChartTitle(id) => {
                if lower == "text" || lower == "caption" {
                    let chart = self.chart_mut(id)?;
                    chart.title = text_of_value(&value);
                    chart.has_title = true;
                    self.redraw(id)?;
                    return Ok(true);
                }
                Ok(false)
            }
            DrawingPart::Legend(id) => {
                if lower == "position" {
                    self.chart_mut(id)?.legend_position = any_whole_number(&value).unwrap_or(-4152);
                    self.redraw(id)?;
                    return Ok(true);
                }
                Ok(matches!(lower.as_str(), "includeinlayout" | "font"))
            }
            DrawingPart::Axis(id, which) => {
                let at = usize::from(which == 2);
                let taken = {
                    let chart = self.chart_mut(id)?;
                    let axis = &mut chart.axes[at];
                    match lower.as_str() {
                        "hastitle" => {
                            axis.has_title = style_face_boolean(&value, "Axis.HasTitle")?.unwrap_or(false);
                            if axis.has_title && axis.title.is_empty() {
                                axis.title = "軸ラベル".to_string();
                            }
                            true
                        }
                        "minimumscale" => { axis.min = any_number(&value); true }
                        "maximumscale" => { axis.max = any_number(&value); true }
                        "minimumscaleisauto" => {
                            if style_face_boolean(&value, "Axis.MinimumScaleIsAuto")?.unwrap_or(false) {
                                axis.min = None;
                            }
                            true
                        }
                        "maximumscaleisauto" => {
                            if style_face_boolean(&value, "Axis.MaximumScaleIsAuto")?.unwrap_or(false) {
                                axis.max = None;
                            }
                            true
                        }
                        "majorunit" | "minorunit" | "hasmajorgridlines" | "hasminorgridlines" | "tickLabelPosition"
                        | "ticklabelposition" | "reverseplotorder" | "crosses" | "crossesat" | "majorunitisauto" => true,
                        _ => false,
                    }
                };
                if taken {
                    self.redraw(id)?;
                }
                Ok(taken)
            }
            DrawingPart::AxisTitle(id, which) => {
                if lower == "text" || lower == "caption" {
                    let chart = self.chart_mut(id)?;
                    chart.axes[usize::from(which == 2)].title = text_of_value(&value);
                    self.redraw(id)?;
                    return Ok(true);
                }
                Ok(false)
            }
            DrawingPart::ChartArea(_) | DrawingPart::PlotArea(_) => {
                Ok(matches!(lower.as_str(), "rgb" | "color" | "colorindex" | "visible" | "width" | "height" | "left" | "top" | "insideleft" | "insidetop" | "insidewidth" | "insideheight"))
            }
            DrawingPart::Adjustments(id) => {
                if lower == "item" {
                    return Ok(false);
                }
                let _ = id;
                Ok(false)
            }
            DrawingPart::Paragraphs(_) => Ok(false),
            DrawingPart::ShapeRange(index) => {
                let ids = self.shape_range_ids(index);
                let mut any = false;
                for id in ids {
                    if self.drawing_set(DrawingPart::Shape(id), name, value.clone())? {
                        any = true;
                    }
                }
                Ok(any)
            }
            DrawingPart::Shapes(_) | DrawingPart::ChartObjects(_) | DrawingPart::SeriesCollection(_) | DrawingPart::Points(..) | DrawingPart::Axes(_) => Ok(false),
        }
    }

    // ---- calling ------------------------------------------------------------

    pub(super) fn drawing_call(&mut self, part: DrawingPart, name: &str, args: &[Value]) -> Result<Option<Value>, String> {
        let lower = name.to_ascii_lowercase();
        match part {
            DrawingPart::Shapes(sheet) => match lower.as_str() {
                "item" | "_default" => {
                    let listed = self.shapes_on(sheet);
                    let index = args.first().ok_or_else(|| "Shapes.Item takes an index or a name".to_string())?;
                    let id = self.pick_shape(&listed, index)?;
                    Ok(Some(self.part(DrawingPart::Shape(id))))
                }
                "addshape" => self.add_shape(sheet, args).map(Some),
                "addtextbox" => self.add_textbox(sheet, args).map(Some),
                "addlabel" => self.add_textbox(sheet, args).map(Some),
                "addline" | "addconnector" => {
                    let args: Vec<Value> = if lower == "addconnector" { args.iter().skip(1).cloned().collect() } else { args.to_vec() };
                    self.add_line(sheet, &args).map(Some)
                }
                "addchart2" | "addchart" => {
                    // AddChart2 Style, XlChartType, Left, Top, Width, Height;
                    // AddChart XlChartType, Left, Top, Width, Height.
                    let skip = usize::from(lower == "addchart2");
                    let chart_type = args.get(skip).and_then(any_whole_number).unwrap_or(51);
                    let number = |index: usize| args.get(skip + 1 + index).and_then(any_number).unwrap_or(0.0);
                    let (left, top, width, height) = (number(0), number(1), number(2).max(1.0), number(3).max(1.0));
                    let (width, height) = if args.get(skip + 3).is_none() { (360.0, 216.0) } else { (width, height) };
                    let id = self.add_chart(sheet, left, top, width, height, chart_type)?;
                    Ok(Some(self.part(DrawingPart::Shape(id))))
                }
                "addpicture" => Err(host_error(1004, "the browser cannot read a picture from a path")),
                "range" => {
                    let listed = self.shapes_on(sheet);
                    let names: Vec<Value> = match args.first() {
                        Some(Value::Array(listed)) => listed.values.clone(),
                        Some(one) => vec![one.clone()],
                        None => return Err(host_error(1004, "Shapes.Range takes a name or an array of names")),
                    };
                    let mut ids = Vec::new();
                    for name in &names {
                        ids.push(self.pick_shape(&listed, name)?);
                    }
                    Ok(Some(self.shape_range_object(ids)))
                }
                "selectall" => {
                    self.shape_selection = self.shapes_on(sheet);
                    Ok(Some(Value::Empty))
                }
                "count" => Ok(Some(Value::Integer(self.shapes_on(sheet).len() as i64))),
                _ => Ok(None),
            },
            DrawingPart::ChartObjects(sheet) => match lower.as_str() {
                "item" | "_default" => {
                    let listed = self.charts_on(sheet);
                    let index = args.first().ok_or_else(|| "ChartObjects.Item takes an index or a name".to_string())?;
                    let id = self.pick_shape(&listed, index)?;
                    Ok(Some(self.part(DrawingPart::ChartObject(id))))
                }
                "add" => {
                    let (left, top, width, height) = Self::placed(args, 0, "ChartObjects.Add")?;
                    let id = self.add_chart(sheet, left, top, width, height, 51)?;
                    Ok(Some(self.part(DrawingPart::ChartObject(id))))
                }
                "delete" => {
                    for id in self.charts_on(sheet) {
                        self.delete_shape(id)?;
                    }
                    Ok(Some(Value::Empty))
                }
                "count" => Ok(Some(Value::Integer(self.charts_on(sheet).len() as i64))),
                _ => Ok(None),
            },
            DrawingPart::Shape(id) | DrawingPart::ChartObject(id) => match lower.as_str() {
                "delete" => {
                    self.delete_shape(id)?;
                    Ok(Some(Value::Empty))
                }
                "duplicate" => {
                    let new_id = self.duplicate_shape(id, None)?;
                    Ok(Some(self.part(DrawingPart::Shape(new_id))))
                }
                "copy" => {
                    self.shape_clipboard = Some(id);
                    self.clipboard = None;
                    self.pending_cut = None;
                    Ok(Some(Value::Boolean(true)))
                }
                "cut" => {
                    self.shape_clipboard = Some(id);
                    Ok(Some(Value::Boolean(true)))
                }
                "select" | "activate" => {
                    self.shape_selection = vec![id];
                    Ok(Some(Value::Boolean(true)))
                }
                "zorder" => Ok(Some(Value::Empty)),
                "incrementleft" | "incrementtop" | "incrementrotation" => {
                    let by = args.first().and_then(any_number).unwrap_or(0.0);
                    let shape = self.shape_mut(id)?;
                    match lower.as_str() {
                        "incrementleft" => shape.left = (shape.left + by).max(0.0),
                        "incrementtop" => shape.top = (shape.top + by).max(0.0),
                        _ => shape.rotation += by,
                    }
                    self.redraw(id)?;
                    Ok(Some(Value::Empty))
                }
                "scalewidth" | "scaleheight" => {
                    let factor = args.first().and_then(any_number).unwrap_or(1.0);
                    let shape = self.shape_mut(id)?;
                    if lower == "scalewidth" {
                        shape.width *= factor;
                    } else {
                        shape.height *= factor;
                    }
                    self.redraw(id)?;
                    Ok(Some(Value::Empty))
                }
                "bringtofront" | "sendtoback" => Ok(Some(Value::Empty)),
                "adjustments" => {
                    if args.is_empty() || matches!(args, [Value::Missing]) {
                        return Ok(Some(self.part(DrawingPart::Adjustments(id))));
                    }
                    self.drawing_call(DrawingPart::Adjustments(id), "item", args)
                }
                "textframe" => Ok(Some(self.part(DrawingPart::TextFrame(id)))),
                _ => Ok(None),
            },
            DrawingPart::TextFrame(id) | DrawingPart::TextRange(id) | DrawingPart::Characters(id, ..) => match lower.as_str() {
                "characters" => {
                    let start = args.first().and_then(any_whole_number).unwrap_or(1).max(1) as u32;
                    let length = args.get(1).and_then(any_whole_number).map(|n| n.max(0) as u32);
                    Ok(Some(self.part(DrawingPart::Characters(id, start, length))))
                }
                "delete" | "clear" => {
                    self.shape_mut(id)?.set_text("");
                    self.redraw(id)?;
                    Ok(Some(Value::Empty))
                }
                "insertafter" | "insertbefore" => {
                    let added = args.first().map(text_of_value).unwrap_or_default();
                    let shape = self.shape_mut(id)?;
                    let text = if lower == "insertafter" { format!("{}{added}", shape.text()) } else { format!("{added}{}", shape.text()) };
                    shape.set_text(&text);
                    self.redraw(id)?;
                    Ok(Some(Value::Empty))
                }
                _ => Ok(None),
            },
            DrawingPart::Chart(id) => match lower.as_str() {
                "setsourcedata" => {
                    let Some(range) = args.first().and_then(|value| self.cells_of_range_object(value)) else {
                        return Err(host_error(1004, "SetSourceData takes a Range"));
                    };
                    // xlRows 1, xlColumns 2.
                    let by_rows = args.get(1).and_then(any_whole_number) == Some(1);
                    self.set_source_data(id, range, by_rows)?;
                    Ok(Some(Value::Empty))
                }
                "axes" => {
                    let which = args.first().and_then(any_whole_number).unwrap_or(1);
                    if !matches!(which, 1 | 2) {
                        return Err(host_error(1004, "Axes takes xlCategory or xlValue"));
                    }
                    Ok(Some(self.part(DrawingPart::Axis(id, which))))
                }
                "seriescollection" => match args.first() {
                    None | Some(Value::Missing) => Ok(Some(self.part(DrawingPart::SeriesCollection(id)))),
                    Some(index) => {
                        let number = self.series_number(id, index)?;
                        Ok(Some(self.part(DrawingPart::Series(id, number))))
                    }
                },
                "hasaxis" => Ok(Some(Value::Boolean(true))),
                // Measured: ApplyLayout 1 gives the chart a title.
                "applylayout" => {
                    let chart = self.chart_mut(id)?;
                    chart.has_title = true;
                    if chart.title.is_empty() {
                        chart.title = "グラフ タイトル".to_string();
                    }
                    self.redraw(id)?;
                    Ok(Some(Value::Empty))
                }
                "export" => Err(host_error(1004, "the browser cannot write a file")),
                "delete" => {
                    self.delete_shape(id)?;
                    Ok(Some(Value::Empty))
                }
                "refresh" | "activate" | "select" | "deselect" | "setelement" | "applychartTemplate" | "applycharttemplate" | "clearToMatchStyle" | "cleartomatchstyle" => Ok(Some(Value::Empty)),
                "location" => Ok(Some(self.part(DrawingPart::Chart(id)))),
                _ => Ok(None),
            },
            DrawingPart::SeriesCollection(id) => match lower.as_str() {
                "item" | "_default" => {
                    let index = args.first().ok_or_else(|| "SeriesCollection takes an index".to_string())?;
                    let number = self.series_number(id, index)?;
                    Ok(Some(self.part(DrawingPart::Series(id, number))))
                }
                "newseries" => {
                    let chart = self.chart_mut(id)?;
                    let number = chart.series.len() + 1;
                    chart.series.push(SeriesRecord {
                        name: format!("系列{number}"),
                        name_ref: None,
                        values_ref: None,
                        x_ref: None,
                        values: Vec::new(),
                        xs: Vec::new(),
                        color: None,
                        chart_type: None,
                        has_labels: false,
                    });
                    self.redraw(id)?;
                    Ok(Some(self.part(DrawingPart::Series(id, number))))
                }
                "add" => {
                    if let Some(range) = args.first().and_then(|value| self.cells_of_range_object(value)) {
                        let numbers = self.range_numbers(range);
                        let reference = self.address_of(range);
                        let chart = self.chart_mut(id)?;
                        let number = chart.series.len() + 1;
                        chart.series.push(SeriesRecord {
                            name: format!("系列{number}"),
                            name_ref: None,
                            values_ref: Some(reference),
                            x_ref: None,
                            xs: (1..=numbers.len()).map(|n| n.to_string()).collect(),
                            values: numbers,
                            color: None,
                            chart_type: None,
                            has_labels: false,
                        });
                        self.redraw(id)?;
                    }
                    Ok(Some(Value::Empty))
                }
                "count" => Ok(Some(Value::Integer(self.chart(id)?.series.len() as i64))),
                _ => Ok(None),
            },
            DrawingPart::Series(id, number) => match lower.as_str() {
                "delete" => {
                    let chart = self.chart_mut(id)?;
                    if number >= 1 && number <= chart.series.len() {
                        chart.series.remove(number - 1);
                    }
                    self.redraw(id)?;
                    Ok(Some(Value::Empty))
                }
                "points" => match args.first() {
                    None | Some(Value::Missing) => Ok(Some(self.part(DrawingPart::Points(id, number)))),
                    Some(_) => Ok(Some(self.part(DrawingPart::Points(id, number)))),
                },
                "datalabels" => Ok(Some(self.part(DrawingPart::DataLabels(id, number)))),
                "applydatalabels" => {
                    self.series_mut(id, number)?.has_labels = true;
                    self.redraw(id)?;
                    Ok(Some(Value::Empty))
                }
                "select" => Ok(Some(Value::Boolean(true))),
                // `Values(2)` and `XValues(2)` index the arrays they answer.
                "values" | "xvalues" => {
                    let whole = self.drawing_get(DrawingPart::Series(id, number), &lower)?;
                    match (args.first().and_then(any_whole_number), whole) {
                        (Some(index), Some(Value::Array(array))) => Ok(Some(
                            array
                                .values
                                .get(index.max(1) as usize - 1)
                                .cloned()
                                .ok_or_else(|| host_error(9, "no such point"))?,
                        )),
                        (_, whole) => Ok(whole),
                    }
                }
                _ => Ok(None),
            },
            DrawingPart::Adjustments(id) => match lower.as_str() {
                "item" | "_default" => {
                    let index = args.first().and_then(any_whole_number).unwrap_or(1);
                    let shape = self.shape(id)?;
                    shape
                        .adjusts
                        .get(index.max(1) as usize - 1)
                        .map(|value| Some(Value::Double(*value)))
                        .ok_or_else(|| host_error(1004, "there is no such adjustment"))
                }
                _ => Ok(None),
            },
            DrawingPart::Paragraphs(id) => match lower.as_str() {
                "item" | "_default" => Ok(Some(self.part(DrawingPart::TextRange(id)))),
                _ => Ok(None),
            },
            DrawingPart::Axes(id) => match lower.as_str() {
                "item" | "_default" => {
                    let which = args.first().and_then(any_whole_number).unwrap_or(1);
                    if !matches!(which, 1 | 2) {
                        return Err(host_error(1004, "Axes takes xlCategory or xlValue"));
                    }
                    Ok(Some(self.part(DrawingPart::Axis(id, which))))
                }
                _ => Ok(None),
            },
            DrawingPart::Points(id, number) => match lower.as_str() {
                "item" | "_default" | "count" => Ok(Some(Value::Integer(self.series(id, number)?.values.len() as i64))),
                _ => Ok(None),
            },
            DrawingPart::ChartTitle(id) => match lower.as_str() {
                "characters" => Ok(Some(self.part(DrawingPart::ChartTitle(id)))),
                "select" | "delete" => Ok(Some(Value::Empty)),
                _ => Ok(None),
            },
            DrawingPart::AxisTitle(id, which) => match lower.as_str() {
                "characters" => Ok(Some(self.part(DrawingPart::AxisTitle(id, which)))),
                _ => Ok(None),
            },
            DrawingPart::ShapeRange(index) => {
                let ids = self.shape_range_ids(index);
                match lower.as_str() {
                    "select" => {
                        self.shape_selection = ids;
                        Ok(Some(Value::Boolean(true)))
                    }
                    "item" | "_default" => {
                        let at = args.first().and_then(any_whole_number).unwrap_or(1).max(1) as usize;
                        ids.get(at - 1)
                            .map(|id| Ok(Some(self.part(DrawingPart::Shape(*id)))))
                            .unwrap_or_else(|| Err(host_error(1004, "there is no such shape")))
                    }
                    "count" => Ok(Some(Value::Integer(ids.len() as i64))),
                    "shaperange" => Ok(Some(self.part(part))),
                    "delete" => {
                        for id in ids {
                            self.delete_shape(id)?;
                        }
                        self.shape_selection.clear();
                        Ok(Some(Value::Empty))
                    }
                    "group" => Ok(ids.first().map(|id| self.part(DrawingPart::Shape(*id)))),
                    _ => match ids.first() {
                        Some(id) => self.drawing_call(DrawingPart::Shape(*id), name, args),
                        None => Ok(None),
                    },
                }
            }
            DrawingPart::Legend(_) | DrawingPart::DataLabels(..) | DrawingPart::ChartArea(_) | DrawingPart::PlotArea(_) => match lower.as_str() {
                "select" | "delete" | "clearformats" => Ok(Some(Value::Empty)),
                _ => Ok(None),
            },
            _ => Ok(None),
        }
    }

    /// A series' `Values`, `XValues` or `Name` given a Range: the range is
    /// what is kept, and its cells are read now.
    pub(super) fn drawing_set_range(&mut self, part: DrawingPart, name: &str, range: CellRange) -> Result<Option<bool>, String> {
        let DrawingPart::Series(id, number) = part else {
            return Ok(None);
        };
        let reference = self.address_of(range);
        match name.to_ascii_lowercase().as_str() {
            "values" => {
                let numbers = self.range_numbers(range);
                let series = self.series_mut(id, number)?;
                series.values = numbers;
                series.values_ref = Some(reference);
                if series.xs.len() != series.values.len() {
                    series.xs = (1..=series.values.len()).map(|n| n.to_string()).collect();
                }
            }
            "xvalues" => {
                let texts = self.range_texts(range);
                let series = self.series_mut(id, number)?;
                series.xs = texts;
                series.x_ref = Some(reference);
            }
            "name" => {
                let text = self.range_texts(range).into_iter().next().unwrap_or_default();
                let series = self.series_mut(id, number)?;
                series.name = text;
                series.name_ref = Some(reference);
            }
            _ => return Ok(None),
        }
        self.redraw(id)?;
        Ok(Some(true))
    }

    fn series_number(&self, id: u64, index: &Value) -> Result<usize, String> {
        let chart = self.chart(id)?;
        match index {
            Value::String(name) => chart
                .series
                .iter()
                .position(|series| series.name.eq_ignore_ascii_case(name))
                .map(|at| at + 1)
                .ok_or_else(|| host_error(1004, format!("there is no series called {name}"))),
            value => match any_whole_number(value) {
                Some(number) if number >= 1 && (number as usize) <= chart.series.len() => Ok(number as usize),
                _ => Err(host_error(1004, "there is no such series")),
            },
        }
    }
}

/// The name Excel gives a preset shape, by `msoAutoShapeType`.
fn auto_shape_name(kind: i64) -> &'static str {
    auto_shape_label(kind)
}

fn chart_is_pie(kind: i64) -> bool {
    matches!(kind, 5 | 69 | 70 | 71 | -4102 | -4120 | 80)
}

/// Text for what a macro hands a text property that is not text.
fn text_of_value(value: &Value) -> String {
    match value {
        Value::String(text) => text.clone(),
        Value::Empty | Value::Null | Value::Missing | Value::Nothing => String::new(),
        Value::Boolean(flag) => if *flag { "True".to_string() } else { "False".to_string() },
        other => match any_number(other) {
            Some(number) if number.fract() == 0.0 => format!("{}", number as i64),
            Some(number) => number.to_string(),
            None => String::new(),
        },
    }
}

/// The theme table's index for accent1.
const ACCENT1_INDEX: usize = 4;
