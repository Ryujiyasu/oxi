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
    /// The names the workbook gives to things, as `(name, what it refers to)`.
    /// A formula saying `SUM(sales)` means one of these, and without them it
    /// means nothing at all.
    ///
    /// Only the ones that belong to the whole workbook. A name can also be
    /// scoped to a single sheet, and two sheets are each entitled to mean
    /// something different by the same word — holding those here, where there
    /// is one name for the whole book, would answer some formulas with another
    /// sheet's range.
    #[serde(default, skip_serializing_if = "Vec::is_empty")]
    pub defined_names: Vec<(String, String)>,
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
    /// The first font in the workbook's list, with a theme scheme resolved
    /// the way every other font's is. One level of a cell's indent is three
    /// of this font's spaces — this one, and not the font the Normal style
    /// points at, which is a different entry in books openpyxl writes.
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub first_font: Option<(String, f32)>,
    /// How many rows are held at the top of the view while the rest scrolls,
    /// and how many columns at the left. Both zero unless the sheet says
    /// `<pane state="frozen">`, which counts them in cells; a plain split
    /// states its position in twips and is a different thing entirely.
    #[serde(default, skip_serializing_if = "is_zero")]
    pub frozen_rows: u32,
    #[serde(default, skip_serializing_if = "is_zero")]
    pub frozen_cols: u32,
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
    /// What a formula calls it. `tblNomina[[#This Row],[DATE]]` names this.
    #[serde(default, skip_serializing_if = "String::is_empty")]
    pub name: String,
    /// The heading of each column, left to right, so that a formula naming one
    /// can be told which column of `start_col..=end_col` it meant.
    #[serde(default, skip_serializing_if = "Vec::is_empty")]
    pub columns: Vec<String>,
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
    /// The rule the table draws round its whole self, which it states as an
    /// index into the workbook's differential formats rather than as a border
    /// on any cell. Fifteen of the corpus's workbooks carry one — the whole
    /// `procurement_contractor` family — and its bottom edge is a rule
    /// thirteen hundred pixels long that no cell mentions.
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub outline: Option<BorderLine>,
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
    /// Whether the run carried an `<rPr>` of its own.
    ///
    /// It is not the same question as whether any of the fields above are
    /// set. An `<rPr>` REPLACES the cell's font rather than overriding parts
    /// of it, so a run that carries one and does not say `<b/>` is NOT bold
    /// even in a bold cell — while a run with no `<rPr>` at all wears the
    /// cell's font whole. Both look identical in the fields above, so the
    /// fact that the element was there has to be kept.
    #[serde(default, skip_serializing_if = "std::ops::Not::not")]
    pub dressed: bool,
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

fn is_zero_i32(value: &i32) -> bool {
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
    /// Where it sits as fractions of its parent's box, for a shape that hangs
    /// on a chart rather than on the grid: a chart keeps the boxes it is
    /// annotated with in a part of its own, anchored that way.
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub frame: Option<Frame>,
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
    /// How big the box is, in points, as the style states it. It is what is
    /// used when the anchor's far corner names a cell the picture does not
    /// reach.
    pub size: (f32, f32),
    /// The anchor's far corner. Excel draws a note between its two corners:
    /// `002`'s note says `height:58pt` and Excel draws 78 pixels, which is
    /// what the anchor's rows come to. A note Excel wrote itself agrees —
    /// its far corner's offset is the width in pixels exactly.
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub to: Option<Anchor>,
    /// The note's text, dressed the way a shape's is.
    pub text: ShapeText,
    /// Six hex digits; Excel's own note is FFFFE1.
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub fill: Option<String>,
}

/// One step of a shape's own outline, in the space the path states.
#[derive(Debug, Clone, Copy, PartialEq, Serialize, Deserialize)]
pub enum PathStep {
    MoveTo(i64, i64),
    LineTo(i64, i64),
    /// Two control points and the point the curve ends on.
    CurveTo(i64, i64, i64, i64, i64, i64),
    Close,
}

/// A shape drawn from its own outline rather than from a preset.
///
/// The corpus holds sixteen of them across four workbooks — one path each,
/// 160 straight segments, 18 curves and 10 closes — and one of them is the
/// curly brace beside `002`'s notes, in the lowest-scoring workbook there is.
/// The points are stated in a space of the path's own, which is mapped onto
/// the box the anchors give the shape.
#[derive(Debug, Clone, PartialEq, Serialize, Deserialize)]
pub struct ShapePath {
    pub across: i64,
    pub down: i64,
    pub steps: Vec<PathStep>,
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
    /// What a preset's adjust handles are set to, by name — `adj`, `adj1`,
    /// `adj2` — in the hundred-thousandths a preset states them in. A preset
    /// left alone carries none and takes its own defaults.
    #[serde(default, skip_serializing_if = "Vec::is_empty")]
    pub adjusts: Vec<(String, i64)>,
    /// The outline the shape draws itself with, when it has one of its own
    /// rather than a preset's.
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub path: Option<ShapePath>,
    /// A shape whose box is flipped: a line drawn from the other corner.
    #[serde(default, skip_serializing_if = "std::ops::Not::not")]
    pub flip_h: bool,
    #[serde(default, skip_serializing_if = "std::ops::Not::not")]
    pub flip_v: bool,
    /// How far round the shape is turned, in sixtieth-thousandths of a degree,
    /// clockwise. The turn is about the box's centre, and the anchor states
    /// the box the turn LEAVES it in: a tall shape turned a quarter is hung
    /// from an anchor as wide as the shape was tall.
    #[serde(default, skip_serializing_if = "is_zero_i32")]
    pub rotation: i32,
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
    /// True when the body says `vertOverflow="clip"`. Excel then draws only
    /// the lines that fit the box and drops the rest; with the attribute
    /// absent, or saying `overflow`, every line is drawn and the block hangs
    /// out of the box. 256 of the corpus's shapes clip, 322 say nothing and
    /// 7 say overflow.
    #[serde(default, skip_serializing_if = "std::ops::Not::not")]
    pub clip: bool,
}

/// One piece of a shape's paragraph, when the paragraph is not written all
/// one way.
///
/// Only weight, underline and colour vary inside a shape's paragraph in the
/// corpus — 42 of its 826 shape paragraphs, across nine workbooks, and never
/// the size or the face — so the paragraph keeps the face and the size it is
/// laid out with and the runs carry the rest. The texts of the runs add up to
/// the paragraph's own text, the break a `<a:br/>` makes included.
#[derive(Debug, Clone, PartialEq, Serialize, Deserialize)]
pub struct ShapeRun {
    pub text: String,
    #[serde(default, skip_serializing_if = "std::ops::Not::not")]
    pub bold: bool,
    #[serde(default, skip_serializing_if = "std::ops::Not::not")]
    pub underline: bool,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub color: Option<String>,
}

/// One paragraph of a shape's text, dressed the way its first run is.
#[derive(Debug, Clone, PartialEq, Serialize, Deserialize)]
pub struct ShapeParagraph {
    pub text: String,
    /// The pieces the text is written in. Empty for a paragraph written all
    /// one way, which is the common case.
    #[serde(default, skip_serializing_if = "Vec::is_empty")]
    pub runs: Vec<ShapeRun>,
    /// "l", "ctr", "r" or "just".
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub align: Option<String>,
    /// Points. A run states hundredths of one.
    pub size: f32,
    #[serde(default, skip_serializing_if = "std::ops::Not::not")]
    pub bold: bool,
    #[serde(default, skip_serializing_if = "std::ops::Not::not")]
    pub italic: bool,
    // A shape's underline is NOT here on purpose. `u="sng"` sits on a run,
    // and a shape paragraph is dressed by its FIRST run — in `sanko_tool`
    // one paragraph holds an underlined heading, two `<a:br/>`s and then the
    // body, so wearing the heading's underline across the paragraph underlines
    // all of it and costs that workbook 0.0092. It needs per-run dressing of a
    // shape's line, the way a cell's runs are dressed, not a paragraph flag.
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub face: Option<String>,
    /// The charset the run states beside the face, when it states one.
    /// Excel's answer for a face this machine has not got turns on this and
    /// on nothing else — not the name, not the PANOSE the file carries.
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub charset: Option<i32>,
    /// The `pitchFamily` beside it, which chooses the METRICS Excel lays the
    /// substitute out with — a different question from which face it draws
    /// (SX101). It travels with the face the same way the charset does.
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub pitch_family: Option<i32>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub color: Option<String>,
    /// The line pitch the paragraph pins, in points, when it pins one.
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub line_pitch: Option<f32>,
    /// The share of the font's own pitch the paragraph asks for, when it
    /// states one as a percentage rather than outright — `glossary_05`'s
    /// flowchart sets every box at four fifths.
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub line_scale: Option<f32>,
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
    /// What the chart is annotated with: text boxes and rules kept in a part
    /// of their own and placed as fractions of the chart's box. The corpus's
    /// charts carry their footnotes there rather than in a cell.
    #[serde(default, skip_serializing_if = "Vec::is_empty")]
    pub shapes: Vec<Drawing>,
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
    /// What the rule wears at each end — "triangle", "arrow", "stealth",
    /// "oval", "diamond" — when it wears anything. 76 of the corpus's rules
    /// carry one, and `glossary_05`'s flowchart is drawn almost entirely out
    /// of them.
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub head_end: Option<String>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub tail_end: Option<String>,
    /// Six hex digits.
    pub color: String,
    /// Width in EMU. Excel draws a line of its own accord at 9525 — one pixel
    /// at 96 dpi — when the shape states no width.
    pub width: i64,
    /// "dash", "dashDot", "sysDot" and the rest, when the rule is broken.
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub dash: Option<String>,
    /// What the rule's ends are cut with: "rnd", "sq", or flat when it says
    /// nothing. It is not decoration — Excel lengthens every dash of a
    /// FLAT-capped or square-capped rule by a pixel when the rule is an odd
    /// number of pixels wide, and leaves a round-capped one alone. The two
    /// `dot` rules a workbook can hold are told apart by this and nothing
    /// else.
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub cap: Option<String>,
}

/// How a cell is drawn.
///
/// `#[serde(default)]` on the whole struct is load-bearing for the browser:
/// the editor builds a cell the moment someone types into an empty one, and it
/// has no style to give it. Requiring every field would mean the editor had to
/// know and restate the default of each — which is to say, hold a second copy
/// of this struct that would drift.
#[derive(Debug, Clone, PartialEq, Serialize, Deserialize, Default)]
#[serde(default)]
pub struct CellStyle {
    pub bold: bool,
    pub italic: bool,
    #[serde(default, skip_serializing_if = "std::ops::Not::not")]
    pub underline: bool,
    pub font_size: Option<f32>,
    /// The typeface a cell asks for, when it names one.
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub font_name: Option<String>,
    /// The charset the font record states beside the face, and the family.
    /// A workbook that asks for a face this machine has not got is answered
    /// from these two and not from the name, so both have to survive as far
    /// as the renderer (`cell_face_in_place`).
    #[serde(default)]
    pub font_charset: Option<i32>,
    #[serde(default)]
    pub font_family: Option<i32>,
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
