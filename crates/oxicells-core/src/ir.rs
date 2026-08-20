// This Source Code Form is subject to the terms of the Mozilla Public
// License, v. 2.0. If a copy of the MPL was not distributed with this
// file, You can obtain one at https://mozilla.org/MPL/2.0/.

use serde::{Deserialize, Serialize};

#[derive(Debug, Clone, Default, Serialize, Deserialize)]
pub struct Workbook {
    pub sheets: Vec<Sheet>,
    /// The format a cell wears when it names none of its own: the first entry
    /// of the workbook's cell formats. A column width is stated in characters
    /// of this font, so drawing a sheet needs it.
    #[serde(default)]
    pub default_style: CellStyle,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct Sheet {
    pub name: String,
    pub rows: Vec<Row>,
    pub col_count: usize,
    pub col_widths: Vec<f32>,
    pub default_col_width: f32,
    pub default_row_height: f32,
    pub merge_cells: Vec<MergeCell>,
    /// Zero-based indices of the columns that are hidden. Columns have no
    /// record of their own, so this sits beside `col_widths`.
    #[serde(default, skip_serializing_if = "Vec::is_empty")]
    pub hidden_cols: Vec<u32>,
    /// The range the sheet declares it occupies, as `(start_row, start_col,
    /// end_row, end_col)` — rows one-based, columns zero-based. Excel hands
    /// this range over when asked for a picture of the sheet, and it can reach
    /// past the last cell that holds anything.
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub declared_range: Option<(u32, u32, u32, u32)>,
    /// The tables on the sheet, each dressed by a named style.
    #[serde(default, skip_serializing_if = "Vec::is_empty")]
    pub tables: Vec<Table>,
    /// The filter a sheet is under, if any.
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub auto_filter: Option<AutoFilter>,
    /// Unsupported elements found in this sheet (e.g. "Chart", "PivotTable", "Drawing")
    #[serde(default, skip_serializing_if = "Vec::is_empty")]
    pub unsupported_elements: Vec<String>,
}

/// A filter over a range, and what each filtered column is testing for.
#[derive(Debug, Clone, PartialEq, Serialize, Deserialize)]
pub struct AutoFilter {
    pub start_row: u32, // 1-based
    pub start_col: u32, // 0-based
    pub end_row: u32,   // 1-based
    pub end_col: u32,   // 0-based
    #[serde(default, skip_serializing_if = "Vec::is_empty")]
    pub columns: Vec<AutoFilterColumn>,
}

/// One column's test, written the way VBA states it: `"apple"`, `">15"`,
/// `"<>banana"`. Two entries are joined by `either`.
#[derive(Debug, Clone, PartialEq, Serialize, Deserialize)]
pub struct AutoFilterColumn {
    /// One-based column within the filtered range.
    pub field: u32,
    pub criteria: Vec<String>,
    /// True when the criteria are joined by xlOr rather than xlAnd.
    #[serde(default, skip_serializing_if = "std::ops::Not::not")]
    pub either: bool,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct MergeCell {
    pub start_row: u32, // 1-based
    pub start_col: u32, // 0-based
    pub end_row: u32,   // 1-based
    pub end_col: u32,   // 0-based
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct Row {
    pub index: u32,
    pub cells: Vec<Cell>,
    pub height: Option<f32>,
    /// A hidden row keeps everything it holds; only the display is affected.
    #[serde(default, skip_serializing_if = "std::ops::Not::not")]
    pub hidden: bool,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct Cell {
    pub col: u32,
    pub value: CellValue,
    pub style: CellStyle,
    /// Original formula string (e.g. "=SUM(A1:A3)"), if any
    #[serde(skip_serializing_if = "Option::is_none")]
    pub formula: Option<String>,
    /// The stretches the text is dressed in, when it is not all dressed alike.
    /// `value` still holds the whole of the text, so a reader that does not
    /// care about the dressing needs no change.
    #[serde(default, skip_serializing_if = "Vec::is_empty")]
    pub runs: Vec<TextRun>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub enum CellValue {
    Empty,
    String(String),
    Number(f64),
    Boolean(bool),
    Error(String),
}

impl CellValue {
    pub fn display(&self) -> String {
        match self {
            CellValue::Empty => String::new(),
            CellValue::String(s) => s.clone(),
            CellValue::Number(n) => {
                if *n == (*n as i64) as f64 {
                    format!("{}", *n as i64)
                } else {
                    format!("{}", n)
                }
            }
            CellValue::Boolean(b) => {
                if *b {
                    "TRUE".to_string()
                } else {
                    "FALSE".to_string()
                }
            }
            CellValue::Error(e) => e.clone(),
        }
    }
}

/// A range a sheet treats as a table, and the style it wears. Excel dresses
/// these itself: no cell inside carries the header's fill or the banding.
#[derive(Debug, Clone, PartialEq, Serialize, Deserialize)]
pub struct Table {
    pub start_row: u32, // 1-based
    pub start_col: u32, // 0-based
    pub end_row: u32,   // 1-based
    pub end_col: u32,   // 0-based
    /// The built-in style's name, e.g. "TableStyleMedium7".
    pub style: Option<String>,
    /// How many rows at the top are the header. Excel writes 1 unless told.
    pub header_rows: u32,
    /// Whether every other row is shaded.
    pub banded_rows: bool,
    /// The colour the style dresses the table in, as six hex digits. A built-in
    /// style takes it from the workbook's theme, so it is resolved on the way
    /// in rather than left as a name to look up later.
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub accent: Option<String>,
    /// The banded rows' fill, which is the accent under a tint.
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub band: Option<String>,
}

/// A stretch of a cell's text that is dressed differently from the rest of it.
/// A cell holds none of these when all of its text is dressed the same; when it
/// does hold them they cover the whole of the text, in order, and each field
/// left empty means "as the cell itself says".
#[derive(Debug, Clone, PartialEq, Serialize, Deserialize, Default)]
pub struct TextRun {
    pub text: String,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub size: Option<f32>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub font: Option<String>,
    #[serde(default, skip_serializing_if = "std::ops::Not::not")]
    pub bold: bool,
    #[serde(default, skip_serializing_if = "std::ops::Not::not")]
    pub italic: bool,
    #[serde(default, skip_serializing_if = "std::ops::Not::not")]
    pub underline: bool,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub color: Option<String>,
    /// "superscript" or "subscript" when the run is raised or lowered.
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub vert_align: Option<String>,
}

/// One edge of a cell, and how it is drawn.
#[derive(Debug, Clone, PartialEq, Serialize, Deserialize)]
pub struct BorderLine {
    /// The kind OOXML names: "thin", "medium", "thick", "hair", "dotted",
    /// "dashed", "double", and the rest. Each is drawn differently — a thick
    /// rule is three pixels where a thin one is a single pixel.
    pub style: String,
    /// The colour as six hex digits, when the edge names one.
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub color: Option<String>,
}

#[derive(Debug, Clone, PartialEq, Serialize, Deserialize, Default)]
pub struct CellStyle {
    pub bold: bool,
    pub italic: bool,
    #[serde(default, skip_serializing_if = "std::ops::Not::not")]
    pub underline: bool,
    pub font_size: Option<f32>,
    /// The typeface a cell asks for, when it names one.
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub font_name: Option<String>,
    pub font_color: Option<String>,
    pub bg_color: Option<String>,
    pub number_format: Option<String>,
    pub horizontal_align: Option<String>,
    /// Where the text sits within the cell's height: "top", "center", "bottom".
    /// Excel leaves a cell at the bottom when it says nothing.
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub vertical_align: Option<String>,
    /// True when the text breaks onto further lines rather than running past
    /// the cell's right edge.
    #[serde(default, skip_serializing_if = "std::ops::Not::not")]
    pub wrap_text: bool,
    pub border_top: Option<BorderLine>,
    pub border_bottom: Option<BorderLine>,
    pub border_left: Option<BorderLine>,
    pub border_right: Option<BorderLine>,
}
