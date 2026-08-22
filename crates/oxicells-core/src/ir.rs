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
    /// True when the sheet pins its default row height (customHeight="1").
    /// Without the pin Excel throws the stated number away and derives the
    /// height from the fonts the columns wear.
    #[serde(default, skip_serializing_if = "std::ops::Not::not")]
    pub default_row_custom: bool,
    /// The font each stretch of columns wears, as `(first, last, name, size)`
    /// with columns zero-based and both ends inside. A cell that states no
    /// format of its own is written in its column's font, and so is the
    /// blank space in every row — which is what a row's height is measured
    /// from when nothing in the row is taller.
    #[serde(default, skip_serializing_if = "Vec::is_empty")]
    pub col_fonts: Vec<(u32, u32, String, f32)>,
    /// The workbook's Normal font, worn by every column no `<col>` dresses.
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub normal_font: Option<(String, f32)>,
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
    /// What is drawn over the grid rather than in it: pictures and shapes.
    #[serde(default, skip_serializing_if = "Vec::is_empty")]
    pub drawings: Vec<Drawing>,
    /// The notes a sheet keeps pinned open. A comment Excel hides is not part
    /// of the picture; one the workbook pins is, and two of the conformance
    /// corpus's workbooks pin ninety between them.
    #[serde(default, skip_serializing_if = "Vec::is_empty")]
    pub comments: Vec<Comment>,
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
    /// True when the row's height is pinned (customHeight="1"). Without the
    /// pin the stored height is only a cache from the machine that wrote the
    /// file — Excel recomputes the height from the row's content on open.
    #[serde(default, skip_serializing_if = "std::ops::Not::not")]
    pub custom_height: bool,
    /// The font of the row's own format (`s=` with customFormat="1"). The
    /// row wears it across all its columns, and its default height becomes
    /// the row's base — which is how a row sinks below the sheet default
    /// without any cell saying a word.
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub style_font: Option<(String, f32)>,
    /// The row says a thick rule runs along its top or bottom edge. Excel
    /// keeps a pixel of room for each when it works a row's height out.
    #[serde(default, skip_serializing_if = "std::ops::Not::not")]
    pub thick_top: bool,
    #[serde(default, skip_serializing_if = "std::ops::Not::not")]
    pub thick_bottom: bool,
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
    /// The rule Excel draws along every row of a `TableStyleMedium` table: the
    /// accent again, under a lighter tint than the banding.
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub rule: Option<String>,
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

fn is_zero(value: &u32) -> bool {
    *value == 0
}

/// A corner of a drawing: the cell it hangs from, and how far into that cell
/// it sits. The offsets are EMU — 914400 to the inch, so 9525 to a pixel at
/// 96 dpi.
#[derive(Debug, Clone, Copy, PartialEq, Serialize, Deserialize)]
pub struct Anchor {
    /// Zero-based, like `Cell::col`.
    pub col: u32,
    pub col_off: i64,
    /// Zero-based, unlike `Row::index`, which is how the drawing part states it.
    pub row: u32,
    pub row_off: i64,
}

/// Something drawn over the grid rather than in it.
#[derive(Debug, Clone, PartialEq, Serialize, Deserialize)]
pub struct Drawing {
    /// Where its top-left corner hangs.
    pub from: Anchor,
    /// Where its bottom-right corner hangs, when the anchor names a second
    /// cell. A drawing anchored to one cell states its size instead.
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub to: Option<Anchor>,
    /// The size the anchor states outright, in EMU.
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub extent: Option<(i64, i64)>,
    pub kind: DrawingKind,
}

#[derive(Debug, Clone, PartialEq, Serialize, Deserialize)]
pub enum DrawingKind {
    /// A picture, holding the bytes of the part it names. They are left out of
    /// the serialised form: what reads the IR back is an editor that has the
    /// file beside it, and a base64 image in every sheet would dwarf the rest.
    Picture {
        #[serde(skip)]
        bytes: Vec<u8>,
    },
    /// A chart, whose picture is drawn from a part of its own.
    Chart(Chart),
    /// A shape Excel fills and rules. The corpus draws 2176 of its 2245
    /// shapes as a line, a rectangle or a rounded one.
    Shape(Shape),
    /// Anything else the drawing part holds — a group, a text box.
    Other,
}

/// A note pinned open on the sheet.
///
/// Its box hangs from two cells, the way a drawing's does. The VML the note
/// lives in states that twice over — once as a margin in points, cached from
/// wherever the file was last written, and once as `<x:Anchor>`, which is what
/// Excel lays it out from and what these corners hold.
#[derive(Debug, Clone, PartialEq, Serialize, Deserialize)]
pub struct Comment {
    pub from: Anchor,
    /// How big the box is, in points. The anchor states a second corner as
    /// well, but Excel sizes a note to its text and keeps the result in the
    /// style's `width`/`height` — which is what it shows.
    pub size: (f32, f32),
    /// The note's text, dressed the way a shape's is.
    pub text: ShapeText,
    /// Six hex digits; Excel's own note is FFFFE1.
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub fill: Option<String>,
}

/// A preset shape, with what it is painted and ruled with.
#[derive(Debug, Clone, PartialEq, Serialize, Deserialize)]
pub struct Shape {
    /// The preset geometry OOXML names: "rect", "line", "roundRect", …
    pub geometry: String,
    /// Six hex digits, when the shape is filled.
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub fill: Option<String>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub line: Option<ShapeLine>,
    /// A shape whose box is flipped: a line drawn from the other corner.
    #[serde(default, skip_serializing_if = "std::ops::Not::not")]
    pub flip_h: bool,
    #[serde(default, skip_serializing_if = "std::ops::Not::not")]
    pub flip_v: bool,
    /// What the shape says, when it says anything. 198 of the corpus's 2202
    /// shapes hold text, and they are the large ones: a banner across the top
    /// of a sheet, a heading over a table.
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub text: Option<ShapeText>,
}

/// The text a shape holds, and how it sits in the shape's box.
#[derive(Debug, Clone, PartialEq, Serialize, Deserialize)]
pub struct ShapeText {
    pub paragraphs: Vec<ShapeParagraph>,
    /// "t", "ctr" or "b" — where the block of lines sits in the box. Excel
    /// leaves a shape's text at the top.
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub anchor: Option<String>,
    /// Left, top, right and bottom, in EMU. Excel's defaults are 91440 and
    /// 45720 — a tenth and a twentieth of an inch.
    pub insets: (i64, i64, i64, i64),
    /// True when a line too long for the box breaks rather than running on.
    pub wrap: bool,
}

/// One paragraph of a shape's text, dressed the way its first run is.
#[derive(Debug, Clone, PartialEq, Serialize, Deserialize)]
pub struct ShapeParagraph {
    pub text: String,
    /// "l", "ctr", "r" or "just".
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub align: Option<String>,
    /// Points. A run states hundredths of one.
    pub size: f32,
    #[serde(default, skip_serializing_if = "std::ops::Not::not")]
    pub bold: bool,
    #[serde(default, skip_serializing_if = "std::ops::Not::not")]
    pub italic: bool,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub face: Option<String>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub color: Option<String>,
    /// The line pitch the paragraph pins, in points, when it pins one.
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub line_pitch: Option<f32>,
}

/// A graph drawn over the grid, as its own part states it.
///
/// Everything the picture needs is in the file: a chart caches the values it
/// plots beside the reference it plots them from, so nothing here waits on a
/// formula.
#[derive(Debug, Clone, PartialEq, Serialize, Deserialize, Default)]
pub struct Chart {
    /// "line", "bar", "pie" — the first plot the chart holds.
    pub kind: String,
    /// The rectangle the axes enclose, as fractions of the chart's own box,
    /// when the chart pins one. Left out, Excel places the plot itself.
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub plot: Option<Frame>,
    pub series: Vec<ChartSeries>,
    /// What the category axis is labelled with, shared by every series.
    #[serde(default, skip_serializing_if = "Vec::is_empty")]
    pub categories: Vec<String>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub value_axis: Option<ChartAxis>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub category_axis: Option<ChartAxis>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub legend: Option<Legend>,
    /// The chart's own background, and the plot rectangle's.
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub fill: Option<String>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub plot_fill: Option<String>,
}

/// A box stated as fractions of its parent's width and height.
#[derive(Debug, Clone, Copy, PartialEq, Serialize, Deserialize, Default)]
pub struct Frame {
    pub x: f64,
    pub y: f64,
    pub w: f64,
    pub h: f64,
}

#[derive(Debug, Clone, PartialEq, Serialize, Deserialize, Default)]
pub struct ChartSeries {
    pub name: String,
    /// One per category. A gap in the data is a `None`, and the line either
    /// breaks across it or steps over it, as the chart says.
    pub values: Vec<Option<f64>>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub line: Option<ShapeLine>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub marker: Option<Marker>,
    /// The points that dress differently from the rest of their series.
    #[serde(default, skip_serializing_if = "Vec::is_empty")]
    pub points: Vec<ChartPoint>,
    #[serde(default, skip_serializing_if = "Vec::is_empty")]
    pub labels: Vec<DataLabel>,
    /// How the series dresses its labels, where a label does not say for
    /// itself: points, a face, and "t", "r", "b", "l" or "ctr" for the side
    /// of the point the label is set against.
    pub label_size: f32,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub label_face: Option<String>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub label_pos: Option<String>,
}

/// One point of a series, where it is drawn unlike its neighbours.
#[derive(Debug, Clone, PartialEq, Serialize, Deserialize, Default)]
pub struct ChartPoint {
    pub index: u32,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub marker: Option<Marker>,
    /// How the line reaching this point is drawn, when the point restates it.
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub line: Option<ShapeLine>,
}

#[derive(Debug, Clone, PartialEq, Serialize, Deserialize, Default)]
pub struct Marker {
    /// "circle", "diamond", "square", "none", …
    pub symbol: String,
    /// Points across, as the chart states it. Excel's own default is 7.
    pub size: u32,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub fill: Option<String>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub line: Option<String>,
}

/// A number written beside the point it belongs to.
#[derive(Debug, Clone, PartialEq, Serialize, Deserialize, Default)]
pub struct DataLabel {
    pub index: u32,
    /// How far the label is nudged from where it would otherwise sit, as
    /// fractions of the chart's box.
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub offset: Option<(f64, f64)>,
    /// What it says, when the label says something other than the value.
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub text: Option<String>,
    pub size: f32,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub face: Option<String>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub number_format: Option<String>,
    /// Which side of the point the label is set against, when this one
    /// differs from the rest of its series.
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub position: Option<String>,
}

#[derive(Debug, Clone, PartialEq, Serialize, Deserialize, Default)]
pub struct ChartAxis {
    /// "b", "l", "r" or "t".
    pub position: String,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub min: Option<f64>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub max: Option<f64>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub major_unit: Option<f64>,
    /// "in", "out", "cross" or "none".
    pub major_tick: String,
    /// "nextTo", "low", "high" or "none".
    pub tick_labels: String,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub number_format: Option<String>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub line: Option<ShapeLine>,
    /// A gridline drawn across the plot at every major tick, when the axis
    /// asks for one.
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub major_gridline: Option<ShapeLine>,
    /// Points.
    pub size: f32,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub face: Option<String>,
    /// An axis the chart states but does not show.
    #[serde(default, skip_serializing_if = "std::ops::Not::not")]
    pub deleted: bool,
    /// "midCat" puts the first point on the axis itself; "between" puts it a
    /// half-step in.
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub cross_between: Option<String>,
}

#[derive(Debug, Clone, PartialEq, Serialize, Deserialize, Default)]
pub struct Legend {
    /// "r", "l", "t", "b" or "tr".
    pub position: String,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub frame: Option<Frame>,
    pub size: f32,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub face: Option<String>,
}

#[derive(Debug, Clone, PartialEq, Serialize, Deserialize)]
pub struct ShapeLine {
    /// Six hex digits.
    pub color: String,
    /// Width in EMU. Excel draws a line of its own accord at 9525 — one pixel
    /// at 96 dpi — when the shape states no width.
    pub width: i64,
    /// "dash", "dashDot", "sysDot" and the rest, when the rule is broken.
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub dash: Option<String>,
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
    /// How many levels the text is pushed in from the edge it is aligned to.
    /// 71 of the conformance corpus's 285 workbooks use one.
    #[serde(default, skip_serializing_if = "is_zero")]
    pub indent: u32,
    /// Where the text sits within the cell's height: "top", "center", "bottom".
    /// Excel leaves a cell at the bottom when it says nothing.
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub vertical_align: Option<String>,
    /// True when the text breaks onto further lines rather than running past
    /// the cell's right edge.
    #[serde(default, skip_serializing_if = "std::ops::Not::not")]
    pub wrap_text: bool,
    /// True when the text is made smaller until it fits the cell's width
    /// rather than running on or being clipped — 85 of the conformance
    /// corpus's 285 workbooks ask for it.
    #[serde(default, skip_serializing_if = "std::ops::Not::not")]
    pub shrink_to_fit: bool,
    /// True when the cell's characters are stacked one above the next rather
    /// than set in a line — `textRotation="255"`, which is how a Japanese
    /// form labels a narrow column. 771 of the 774 rotations in the
    /// conformance corpus are this one.
    #[serde(default, skip_serializing_if = "std::ops::Not::not")]
    pub stacked_text: bool,
    pub border_top: Option<BorderLine>,
    pub border_bottom: Option<BorderLine>,
    pub border_left: Option<BorderLine>,
    pub border_right: Option<BorderLine>,
    /// The rule drawn corner to corner, which a Japanese form uses to strike
    /// a cell out. The border says how it is drawn; the two flags say which
    /// way, and a cell can carry both.
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub border_diagonal: Option<BorderLine>,
    #[serde(default, skip_serializing_if = "std::ops::Not::not")]
    pub diagonal_up: bool,
    #[serde(default, skip_serializing_if = "std::ops::Not::not")]
    pub diagonal_down: bool,
}
