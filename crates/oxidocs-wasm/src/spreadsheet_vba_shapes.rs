// SPDX-License-Identifier: MIT OR Apache-2.0

//! What a macro draws over the grid: `Shapes.AddShape`, `AddTextbox`,
//! `AddLine`, `ChartObjects.Add`, and the members of what they make.
//!
//! Every shape is a record the host keeps in points -- what the macro reads
//! back has to be what it wrote, to the half point -- and the sheet's
//! `drawings` are rebuilt from the records whenever one changes, so the
//! renderer draws what the macro drew. Defaults were measured against Excel
//! 16.0 on a Japanese Office (shapes.vba, shapes2.vba, 2026-09-05).

use super::*;
use oxicells_core::ir::{
    Anchor, Chart, ChartAxis, ChartSeries, Drawing, DrawingKind, Legend, Shape, ShapeLine,
    ShapeParagraph, ShapeRun, ShapeText,
};

/// EMU to the pixel at 96 dpi.
const EMU_PER_PX: f64 = 9525.0;

/// The face a new shape's text is set in: the theme's minor Latin font.
const SHAPE_FACE: &str = "Aptos Narrow";

/// msoThemeColorAccent1 is 5; the packed colour is the theme table's.
const ACCENT1: usize = 5;
/// The darker accent1 a new shape's outline wears, and the gray a new text
/// box's does, as measured.
const SHAPE_OUTLINE: i64 = 3_351_556;
const TEXTBOX_OUTLINE: i64 = 12_369_084;

#[derive(Debug, Clone, PartialEq)]
pub(super) enum ShapeKind {
    /// msoAutoShape (Type 1), with its AutoShapeType.
    Auto(i64),
    /// msoTextBox (Type 17).
    TextBox,
    /// msoLine (Type 9): a straight connector.
    Line,
    /// msoChart (Type 3).
    Chart(Box<ChartRecord>),
    /// msoPicture (Type 13), read from the file and not redrawable here.
    Picture,
    /// Anything else the file held.
    Other,
}

#[derive(Debug, Clone, PartialEq)]
pub(super) struct TextStyle {
    pub name: String,
    pub size: f64,
    pub bold: bool,
    pub italic: bool,
    pub underline: bool,
    pub color: i64,
}

#[derive(Debug, Clone, PartialEq)]
pub(super) struct ShapeRunRecord {
    pub text: String,
    pub style: TextStyle,
}

#[derive(Debug, Clone, PartialEq)]
pub(super) struct ShapeRecord {
    pub id: u64,
    pub sheet: usize,
    pub name: String,
    pub kind: ShapeKind,
    /// Points, as the macro states them.
    pub left: f64,
    pub top: f64,
    pub width: f64,
    pub height: f64,
    pub rotation: f64,
    pub flip_h: bool,
    pub flip_v: bool,
    pub visible: bool,
    pub fill: i64,
    pub fill_theme: Option<usize>,
    pub fill_visible: bool,
    pub transparency: f64,
    pub line: i64,
    pub line_theme: Option<usize>,
    pub line_visible: bool,
    pub line_weight: f64,
    /// msoLineDashStyle, 1 solid.
    pub dash: i64,
    /// msoArrowheadStyle at the far end, 1 none.
    pub arrow_end: i64,
    pub runs: Vec<ShapeRunRecord>,
    /// xlLeft -4131, xlCenter -4108, xlRight -4152.
    pub h_align: i64,
    /// xlTop -4160, xlCenter -4108, xlBottom -4107.
    pub v_align: i64,
    pub margins: (f64, f64, f64, f64),
    pub auto_size: bool,
    /// The preset's adjust handles, as fractions -- a rounded rectangle's
    /// one corner, 0.16667 by default.
    pub adjusts: Vec<f64>,
    /// xlMoveAndSize 1, xlMove 2, xlFreeFloating 3.
    pub placement: i64,
    pub on_action: String,
    pub alt_text: String,
    pub lock_aspect: bool,
    /// How many paragraphs the text was written as.
    pub paragraph_count: usize,
    /// The drawing the file held, when this record stands for one, so the
    /// picture's bytes and everything else read from the file survive.
    pub original: Option<Drawing>,
}

#[derive(Debug, Clone, PartialEq)]
pub(super) struct ChartRecord {
    pub chart_type: i64,
    pub series: Vec<SeriesRecord>,
    pub has_title: bool,
    pub title: String,
    pub has_legend: bool,
    /// xlLegendPositionRight -4152, Bottom -4107, Left -4131, Top -4160.
    pub legend_position: i64,
    pub axes: [AxisRecord; 2],
    pub style: i64,
    /// Whether the title is still Excel's to decide. Measured: a one-series
    /// chart made a pie by its FIRST ChartType shows a title; one whose
    /// type was set to columns first, then a pie, does not.
    pub title_auto: bool,
}

#[derive(Debug, Clone, PartialEq)]
pub(super) struct SeriesRecord {
    /// The name as shown, and the reference it came from when it did.
    pub name: String,
    pub name_ref: Option<String>,
    pub values_ref: Option<String>,
    pub x_ref: Option<String>,
    pub values: Vec<f64>,
    pub xs: Vec<String>,
    pub color: Option<i64>,
    pub chart_type: Option<i64>,
    pub has_labels: bool,
}

#[derive(Debug, Clone, PartialEq, Default)]
pub(super) struct AxisRecord {
    pub has_title: bool,
    pub title: String,
    pub min: Option<f64>,
    pub max: Option<f64>,
}

/// One of the objects hung off a shape or a chart, by the shape's id.
#[derive(Debug, Clone, Copy, PartialEq)]
pub(super) enum DrawingPart {
    /// `Worksheet.Shapes`.
    Shapes(usize),
    Shape(u64),
    Fill(u64),
    FillColor(u64),
    Line(u64),
    LineColor(u64),
    TextFrame(u64),
    TextFrame2(u64),
    TextRange(u64),
    /// `TextFrame.Characters(Start, Length)`: one-based start, and the
    /// length, or None for the rest.
    Characters(u64, u32, Option<u32>),
    CharactersFont(u64, u32, Option<u32>),
    ParagraphFormat(u64),
    /// `Worksheet.ChartObjects`.
    ChartObjects(usize),
    ChartObject(u64),
    Chart(u64),
    SeriesCollection(u64),
    /// A series by its one-based number.
    Series(u64, usize),
    SeriesFormat(u64, usize),
    SeriesFill(u64, usize),
    SeriesColor(u64, usize),
    Points(u64, usize),
    DataLabels(u64, usize),
    ChartTitle(u64),
    Legend(u64),
    Axes(u64),
    /// xlCategory 1 or xlValue 2.
    Axis(u64, i64),
    AxisTitle(u64, i64),
    ChartArea(u64),
    PlotArea(u64),
    Adjustments(u64),
    Paragraphs(u64),
    /// `Selection.ShapeRange` / `Shapes.Range(Array(...))`: several shapes
    /// held as one, by their place in the host's `shape_ranges`.
    ShapeRange(usize),
}

impl DrawingPart {
    pub(super) fn kind_name(self) -> &'static str {
        match self {
            DrawingPart::Shapes(_) => "Shapes",
            DrawingPart::Shape(_) => "Shape",
            DrawingPart::Fill(_) | DrawingPart::SeriesFill(..) => "FillFormat",
            DrawingPart::FillColor(_) | DrawingPart::LineColor(_) | DrawingPart::SeriesColor(..) => {
                "ColorFormat"
            }
            DrawingPart::Line(_) => "LineFormat",
            DrawingPart::TextFrame(_) => "TextFrame",
            DrawingPart::TextFrame2(_) => "TextFrame2",
            DrawingPart::TextRange(_) => "TextRange2",
            DrawingPart::Characters(..) => "Characters",
            DrawingPart::CharactersFont(..) => "Font",
            DrawingPart::ParagraphFormat(_) => "ParagraphFormat2",
            DrawingPart::ChartObjects(_) => "ChartObjects",
            DrawingPart::ChartObject(_) => "ChartObject",
            DrawingPart::Chart(_) => "Chart",
            DrawingPart::SeriesCollection(_) => "SeriesCollection",
            DrawingPart::Series(..) => "Series",
            DrawingPart::SeriesFormat(..) => "ChartFormat",
            DrawingPart::Points(..) => "Points",
            DrawingPart::DataLabels(..) => "DataLabels",
            DrawingPart::ChartTitle(_) => "ChartTitle",
            DrawingPart::Legend(_) => "Legend",
            DrawingPart::Axes(_) => "Axes",
            DrawingPart::Axis(..) => "Axis",
            DrawingPart::AxisTitle(..) => "AxisTitle",
            DrawingPart::ChartArea(_) => "ChartArea",
            DrawingPart::PlotArea(_) => "PlotArea",
            DrawingPart::Adjustments(_) => "Adjustments",
            DrawingPart::Paragraphs(_) => "TextRange2",
            DrawingPart::ShapeRange(_) => "ShapeRange",
        }
    }

    /// The part after the sheets were renumbered; None when its sheet went.
    pub(super) fn renumbered(self, moved: &dyn Fn(usize) -> Option<usize>) -> Option<Self> {
        Some(match self {
            DrawingPart::Shapes(sheet) => DrawingPart::Shapes(moved(sheet)?),
            DrawingPart::ChartObjects(sheet) => DrawingPart::ChartObjects(moved(sheet)?),
            other => other,
        })
    }
}

/// The English name Excel gives a preset, and the OOXML geometry it draws
/// with, by `msoAutoShapeType`. The named ones were measured; the rest
/// follow Excel's own list.
pub(super) fn auto_shape(kind: i64) -> (&'static str, &'static str) {
    match kind {
        1 => ("Rectangle", "rect"),
        2 => ("Parallelogram", "parallelogram"),
        3 => ("Trapezoid", "trapezoid"),
        4 => ("Diamond", "diamond"),
        5 => ("Rounded Rectangle", "roundRect"),
        6 => ("Octagon", "octagon"),
        7 => ("Isosceles Triangle", "triangle"),
        8 => ("Right Triangle", "rtTriangle"),
        9 => ("Oval", "ellipse"),
        10 => ("Hexagon", "hexagon"),
        11 => ("Cross", "plus"),
        12 => ("Regular Pentagon", "pentagon"),
        13 => ("Can", "can"),
        14 => ("Cube", "cube"),
        15 => ("Bevel", "bevel"),
        16 => ("Folded Corner", "foldedCorner"),
        17 => ("Smiley Face", "smileyFace"),
        18 => ("Donut", "donut"),
        19 => ("\"No\" Symbol", "noSmoking"),
        20 => ("Block Arc", "blockArc"),
        21 => ("Heart", "heart"),
        22 => ("Lightning Bolt", "lightningBolt"),
        23 => ("Sun", "sun"),
        24 => ("Moon", "moon"),
        25 => ("Arc", "arc"),
        26 => ("Double Bracket", "bracketPair"),
        27 => ("Double Brace", "bracePair"),
        28 => ("Plaque", "plaque"),
        29 => ("Left Bracket", "leftBracket"),
        30 => ("Right Bracket", "rightBracket"),
        31 => ("Left Brace", "leftBrace"),
        32 => ("Right Brace", "rightBrace"),
        33 => ("Right Arrow", "rightArrow"),
        34 => ("Left Arrow", "leftArrow"),
        35 => ("Up Arrow", "upArrow"),
        36 => ("Down Arrow", "downArrow"),
        37 => ("Left-Right Arrow", "leftRightArrow"),
        38 => ("Up-Down Arrow", "upDownArrow"),
        39 => ("Quad Arrow", "quadArrow"),
        40 => ("Left-Right-Up Arrow", "leftRightUpArrow"),
        41 => ("Bent Arrow", "bentArrow"),
        42 => ("U-Turn Arrow", "uturnArrow"),
        43 => ("Left-Up Arrow", "leftUpArrow"),
        44 => ("Bent-Up Arrow", "bentUpArrow"),
        45 => ("Curved Right Arrow", "curvedRightArrow"),
        46 => ("Curved Left Arrow", "curvedLeftArrow"),
        47 => ("Curved Up Arrow", "curvedUpArrow"),
        48 => ("Curved Down Arrow", "curvedDownArrow"),
        49 => ("Striped Right Arrow", "stripedRightArrow"),
        50 => ("Notched Right Arrow", "notchedRightArrow"),
        51 => ("Pentagon", "homePlate"),
        52 => ("Chevron", "chevron"),
        53 => ("Right Arrow Callout", "rightArrowCallout"),
        54 => ("Left Arrow Callout", "leftArrowCallout"),
        55 => ("Up Arrow Callout", "upArrowCallout"),
        56 => ("Down Arrow Callout", "downArrowCallout"),
        57 => ("Left-Right Arrow Callout", "leftRightArrowCallout"),
        58 => ("Up-Down Arrow Callout", "upDownArrowCallout"),
        59 => ("Quad Arrow Callout", "quadArrowCallout"),
        60 => ("Circular Arrow", "circularArrow"),
        61 => ("Flowchart: Process", "flowChartProcess"),
        62 => ("Flowchart: Alternate Process", "flowChartAlternateProcess"),
        63 => ("Flowchart: Decision", "flowChartDecision"),
        64 => ("Flowchart: Data", "flowChartInputOutput"),
        65 => ("Flowchart: Predefined Process", "flowChartPredefinedProcess"),
        66 => ("Flowchart: Internal Storage", "flowChartInternalStorage"),
        67 => ("Flowchart: Document", "flowChartDocument"),
        68 => ("Flowchart: Multidocument", "flowChartMultidocument"),
        69 => ("Flowchart: Terminator", "flowChartTerminator"),
        70 => ("Flowchart: Preparation", "flowChartPreparation"),
        71 => ("Flowchart: Manual Input", "flowChartManualInput"),
        72 => ("Flowchart: Manual Operation", "flowChartManualOperation"),
        73 => ("Flowchart: Connector", "flowChartConnector"),
        74 => ("Flowchart: Off-page Connector", "flowChartOffpageConnector"),
        75 => ("Flowchart: Card", "flowChartPunchedCard"),
        76 => ("Flowchart: Punched Tape", "flowChartPunchedTape"),
        77 => ("Flowchart: Summing Junction", "flowChartSummingJunction"),
        78 => ("Flowchart: Or", "flowChartOr"),
        79 => ("Flowchart: Collate", "flowChartCollate"),
        80 => ("Flowchart: Sort", "flowChartSort"),
        81 => ("Flowchart: Extract", "flowChartExtract"),
        82 => ("Flowchart: Merge", "flowChartMerge"),
        83 => ("Flowchart: Stored Data", "flowChartOnlineStorage"),
        84 => ("Flowchart: Delay", "flowChartDelay"),
        85 => ("Flowchart: Sequential Access Storage", "flowChartMagneticTape"),
        86 => ("Flowchart: Magnetic Disk", "flowChartMagneticDisk"),
        87 => ("Flowchart: Direct Access Storage", "flowChartMagneticDrum"),
        88 => ("Flowchart: Display", "flowChartDisplay"),
        89 => ("Explosion 1", "irregularSeal1"),
        90 => ("Explosion 2", "irregularSeal2"),
        91 => ("4-Point Star", "star4"),
        92 => ("5-Point Star", "star5"),
        93 => ("8-Point Star", "star8"),
        94 => ("16-Point Star", "star16"),
        95 => ("24-Point Star", "star24"),
        96 => ("32-Point Star", "star32"),
        97 => ("Up Ribbon", "ribbon2"),
        98 => ("Down Ribbon", "ribbon"),
        99 => ("Curved Up Ribbon", "ellipseRibbon2"),
        100 => ("Curved Down Ribbon", "ellipseRibbon"),
        101 => ("Vertical Scroll", "verticalScroll"),
        102 => ("Horizontal Scroll", "horizontalScroll"),
        103 => ("Wave", "wave"),
        104 => ("Double Wave", "doubleWave"),
        105 => ("Rectangular Callout", "wedgeRectCallout"),
        106 => ("Rounded Rectangular Callout", "wedgeRoundRectCallout"),
        107 => ("Oval Callout", "wedgeEllipseCallout"),
        108 => ("Cloud Callout", "cloudCallout"),
        _ => ("AutoShape", "rect"),
    }
}

/// The OOXML dash a `msoLineDashStyle` draws.
pub(super) fn dash_name(style: i64) -> Option<&'static str> {
    Some(match style {
        1 => return None,
        2 => "sysDot",
        3 => "sysDash",
        4 => "dash",
        5 => "dashDot",
        6 => "lgDash",
        7 => "lgDashDot",
        8 => "lgDashDotDot",
        9 => "sysDashDot",
        10 => "sysDashDotDot",
        _ => return None,
    })
}

/// The OOXML arrowhead a `msoArrowheadStyle` draws.
pub(super) fn arrow_name(style: i64) -> Option<&'static str> {
    Some(match style {
        2 => "triangle",
        3 => "arrow",
        4 => "stealth",
        5 => "diamond",
        6 => "oval",
        _ => return None,
    })
}

/// "bar", "line", "pie", "area", "scatter", "doughnut" -- the chart part's
/// own word for an `xlChartType` -- and whether the bars stand up.
pub(super) fn chart_kind(chart_type: i64) -> &'static str {
    match chart_type {
        4 | 63 | 64 | 65 | 66 | 67 | 68 | -4101 => "line",
        5 | 69 | 70 | 71 | -4102 => "pie",
        -4120 | 80 => "doughnut",
        1 | 76 | 77 | 78 | -4098 => "area",
        -4169 | 72 | 73 | 74 | 75 => "scatter",
        -4152 | 15 | 91 | 92 | 93 => "radar",
        _ => "bar",
    }
}

pub(super) fn packed_hex(colour: i64) -> String {
    colour_from_packed(colour)
}

pub(super) fn mso(flag: bool) -> Value {
    Value::Integer(if flag { -1 } else { 0 })
}

pub(super) fn mso_asked(value: &Value, what: &str) -> Result<bool, String> {
    match value {
        Value::Boolean(flag) => Ok(*flag),
        value if any_number(value).is_some() => Ok(any_number(value).unwrap_or_default() != 0.0),
        _ => Err(format!("{what} takes msoTrue or msoFalse")),
    }
}

impl<'a> WorkbookHost<'a> {
    /// The records the file's own drawings stand for, so `Shapes.Count`
    /// counts them and a macro can move or delete them.
    pub(super) fn adopt_drawings(&mut self) {
        for sheet in 0..self.workbook.sheets.len() {
            let drawings = self.workbook.sheets[sheet].drawings.clone();
            for drawing in drawings {
                let (kind, base) = match &drawing.kind {
                    DrawingKind::Picture { .. } => (ShapeKind::Picture, "Picture"),
                    DrawingKind::Chart(chart) => (
                        ShapeKind::Chart(Box::new(ChartRecord {
                            chart_type: match chart.kind.as_str() {
                                "line" => 4,
                                "pie" => 5,
                                "doughnut" => -4120,
                                "area" => 1,
                                "scatter" => -4169,
                                _ => 51,
                            },
                            series: chart
                                .series
                                .iter()
                                .map(|series| SeriesRecord {
                                    name: series.name.clone(),
                                    name_ref: None,
                                    values_ref: None,
                                    x_ref: None,
                                    values: series.values.iter().map(|v| v.unwrap_or(0.0)).collect(),
                                    xs: chart.categories.clone(),
                                    color: None,
                                    chart_type: None,
                                    has_labels: !series.labels.is_empty(),
                                })
                                .collect(),
                            has_title: false,
                            title: String::new(),
                            has_legend: chart.legend.is_some(),
                            title_auto: false,
                            legend_position: -4152,
                            axes: [
                                AxisRecord {
                                    min: chart.category_axis.as_ref().and_then(|axis| axis.min),
                                    max: chart.category_axis.as_ref().and_then(|axis| axis.max),
                                    ..Default::default()
                                },
                                AxisRecord {
                                    min: chart.value_axis.as_ref().and_then(|axis| axis.min),
                                    max: chart.value_axis.as_ref().and_then(|axis| axis.max),
                                    ..Default::default()
                                },
                            ],
                            style: 201,
                        })),
                        "Chart",
                    ),
                    DrawingKind::Shape(shape) => match shape.geometry.as_str() {
                        "line" | "straightConnector1" => (ShapeKind::Line, "Straight Connector"),
                        geometry => {
                            let kind = (1..=108)
                                .find(|kind| auto_shape(*kind).1 == geometry)
                                .unwrap_or(1);
                            (ShapeKind::Auto(kind), auto_shape(kind).0)
                        }
                    },
                    DrawingKind::Other => (ShapeKind::Other, "Object"),
                };
                let (left, top, width, height) = self.points_of_drawing(sheet, &drawing);
                let number = self.next_shape_number(sheet);
                let id = self.next_shape_id;
                self.next_shape_id += 1;
                let mut record = ShapeRecord::blank(id, sheet, format!("{base} {number}"), kind);
                record.left = left;
                record.top = top;
                record.width = width;
                record.height = height;
                if let DrawingKind::Shape(shape) = &drawing.kind {
                    record.rotation = f64::from(shape.rotation) / 60_000.0;
                    record.flip_h = shape.flip_h;
                    record.flip_v = shape.flip_v;
                    if let Some(fill) = &shape.fill {
                        record.fill = colour_to_packed(Some(fill)).unwrap_or(record.fill);
                    } else {
                        record.fill_visible = false;
                    }
                    match &shape.line {
                        Some(line) => {
                            record.line = colour_to_packed(Some(&line.color)).unwrap_or(record.line);
                            record.line_weight = line.width as f64 / EMU_PER_PX * 0.75;
                        }
                        None => record.line_visible = false,
                    }
                    if let Some(text) = &shape.text {
                        record.runs = text
                            .paragraphs
                            .iter()
                            .enumerate()
                            .map(|(at, paragraph)| ShapeRunRecord {
                                text: if at + 1 < text.paragraphs.len() {
                                    format!("{}\n", paragraph.text)
                                } else {
                                    paragraph.text.clone()
                                },
                                style: TextStyle {
                                    name: paragraph.face.clone().unwrap_or_else(|| SHAPE_FACE.to_string()),
                                    size: f64::from(paragraph.size),
                                    bold: paragraph.bold,
                                    italic: paragraph.italic,
                                    underline: false,
                                    color: paragraph
                                        .color
                                        .as_deref()
                                        .and_then(|c| colour_to_packed(Some(c)))
                                        .unwrap_or(0),
                                },
                            })
                            .collect();
                    }
                }
                record.original = Some(drawing);
                self.shapes.push(record);
            }
        }
    }

    /// Where a drawing's box sits, in points, from its anchors.
    pub(super) fn points_of_drawing(&self, sheet: usize, drawing: &Drawing) -> (f64, f64, f64, f64) {
        let px_of = |anchor: &Anchor| -> (f64, f64) {
            let x: f64 = (0..anchor.col).map(|c| f64::from(self.column_px(sheet, c))).sum::<f64>()
                + anchor.col_off as f64 / EMU_PER_PX;
            let y: f64 = (1..=anchor.row).map(|r| f64::from(self.row_px(sheet, r))).sum::<f64>()
                + anchor.row_off as f64 / EMU_PER_PX;
            (x, y)
        };
        let (x, y) = px_of(&drawing.from);
        let (w, h) = match (&drawing.to, drawing.extent) {
            (Some(to), _) => {
                let (x2, y2) = px_of(to);
                ((x2 - x).max(0.0), (y2 - y).max(0.0))
            }
            (None, Some((cx, cy))) => (cx as f64 / EMU_PER_PX, cy as f64 / EMU_PER_PX),
            (None, None) => (0.0, 0.0),
        };
        (x * 0.75, y * 0.75, w * 0.75, h * 0.75)
    }

    /// The anchor of a point on the sheet, in EMU within its cell.
    pub(super) fn anchor_at(&self, sheet: usize, x_pt: f64, y_pt: f64) -> Anchor {
        let x_px = (x_pt / 0.75).max(0.0);
        let y_px = (y_pt / 0.75).max(0.0);
        let mut col = 0u32;
        let mut reached = 0.0f64;
        loop {
            let width = f64::from(self.column_px(sheet, col));
            if reached + width > x_px || col >= MAX_WORKSHEET_COLUMN {
                break;
            }
            reached += width;
            col += 1;
        }
        let col_off = ((x_px - reached) * EMU_PER_PX) as i64;
        let mut row = 1u32;
        let mut down = 0.0f64;
        loop {
            let height = f64::from(self.row_px(sheet, row));
            if down + height > y_px || row >= MAX_WORKSHEET_ROW {
                break;
            }
            down += height;
            row += 1;
        }
        let row_off = ((y_px - down) * EMU_PER_PX) as i64;
        // The drawing part counts rows from nought.
        Anchor { col, col_off, row: row - 1, row_off }
    }

    /// The cell a point of the shape falls in.
    pub(super) fn cell_under(&self, sheet: usize, x_pt: f64, y_pt: f64) -> CellAddress {
        let anchor = self.anchor_at(sheet, x_pt, y_pt);
        CellAddress { sheet, row: anchor.row + 1, column: anchor.col }
    }

    /// The number the next shape on a sheet takes. Excel counts every shape
    /// ever made on the sheet, charts included, and never goes back:
    /// measured, after Rectangle 1 .. Straight Connector 5 and a deletion
    /// the next is Rectangle 6, and the first chart after Rectangle 8 is
    /// Chart 9.
    pub(super) fn next_shape_number(&mut self, sheet: usize) -> u32 {
        let counter = self.shape_counters.entry(sheet).or_insert(0);
        *counter += 1;
        *counter
    }

    pub(super) fn shape(&self, id: u64) -> Result<&ShapeRecord, String> {
        self.shapes
            .iter()
            .find(|shape| shape.id == id)
            .ok_or_else(|| host_error(-2_147_024_809, "the shape has been deleted"))
    }

    pub(super) fn shape_mut(&mut self, id: u64) -> Result<&mut ShapeRecord, String> {
        self.shapes
            .iter_mut()
            .find(|shape| shape.id == id)
            .ok_or_else(|| host_error(-2_147_024_809, "the shape has been deleted"))
    }

    pub(super) fn chart(&self, id: u64) -> Result<&ChartRecord, String> {
        match &self.shape(id)?.kind {
            ShapeKind::Chart(chart) => Ok(chart),
            _ => Err(host_error(1004, "the shape holds no chart")),
        }
    }

    pub(super) fn chart_mut(&mut self, id: u64) -> Result<&mut ChartRecord, String> {
        match &mut self.shape_mut(id)?.kind {
            ShapeKind::Chart(chart) => Ok(chart),
            _ => Err(host_error(1004, "the shape holds no chart")),
        }
    }

    pub(super) fn series(&self, id: u64, number: usize) -> Result<&SeriesRecord, String> {
        self.chart(id)?
            .series
            .get(number.wrapping_sub(1))
            .ok_or_else(|| host_error(1004, "there is no such series"))
    }

    pub(super) fn series_mut(&mut self, id: u64, number: usize) -> Result<&mut SeriesRecord, String> {
        self.chart_mut(id)?
            .series
            .get_mut(number.wrapping_sub(1))
            .ok_or_else(|| host_error(1004, "there is no such series"))
    }

    /// The sheet's drawings, rebuilt from its records.
    pub(super) fn mirror_drawings(&mut self, sheet: usize) {
        let records: Vec<ShapeRecord> =
            self.shapes.iter().filter(|shape| shape.sheet == sheet).cloned().collect();
        let drawings: Vec<Drawing> = records
            .iter()
            .filter(|record| record.visible || record.original.is_some())
            .map(|record| self.drawing_of(record))
            .collect();
        if let Some(held) = self.workbook.sheets.get_mut(sheet) {
            held.drawings = drawings;
        }
    }

    pub(super) fn drawing_of(&self, record: &ShapeRecord) -> Drawing {
        let from = self.anchor_at(record.sheet, record.left, record.top);
        let to = self.anchor_at(record.sheet, record.left + record.width, record.top + record.height);
        let extent = Some((
            (record.width / 0.75 * EMU_PER_PX) as i64,
            (record.height / 0.75 * EMU_PER_PX) as i64,
        ));
        let kind = match &record.kind {
            ShapeKind::Picture | ShapeKind::Other => match &record.original {
                Some(original) => original.kind.clone(),
                None => DrawingKind::Other,
            },
            ShapeKind::Chart(chart) => DrawingKind::Chart(self.chart_of(chart)),
            ShapeKind::Auto(_) | ShapeKind::TextBox | ShapeKind::Line => {
                DrawingKind::Shape(self.shape_of(record))
            }
        };
        Drawing { from, to: Some(to), extent, kind, frame: None, grouped: false }
    }

    pub(super) fn shape_of(&self, record: &ShapeRecord) -> Shape {
        let geometry = match &record.kind {
            ShapeKind::Auto(kind) => auto_shape(*kind).1,
            ShapeKind::Line => "line",
            _ => "rect",
        };
        let line = record.line_visible.then(|| ShapeLine {
            head_end: None,
            tail_end: arrow_name(record.arrow_end).map(str::to_string),
            color: packed_hex(record.line),
            width: (record.line_weight / 0.75 * EMU_PER_PX) as i64,
            dash: dash_name(record.dash).map(str::to_string),
            cap: None,
        });
        let text = (!record.runs.is_empty() && !matches!(record.kind, ShapeKind::Line)).then(|| {
            let whole: String = record.runs.iter().map(|run| run.text.as_str()).collect();
            let first = record.runs.first().map(|run| run.style.clone());
            let paragraphs = whole
                .split('\n')
                .map(|line| ShapeParagraph {
                    text: line.to_string(),
                    runs: Vec::new(),
                    align: Some(match record.h_align {
                        -4108 => "ctr".to_string(),
                        -4152 => "r".to_string(),
                        _ => "l".to_string(),
                    }),
                    size: first.as_ref().map_or(11.0, |style| style.size as f32),
                    bold: first.as_ref().is_some_and(|style| style.bold),
                    italic: first.as_ref().is_some_and(|style| style.italic),
                    face: first.as_ref().map(|style| style.name.clone()),
                    charset: None,
                    pitch_family: None,
                    color: first.as_ref().map(|style| packed_hex(style.color)),
                    line_pitch: None,
                    line_scale: None,
                })
                .collect();
            let _ = ShapeRun { text: String::new(), bold: false, underline: false, color: None };
            ShapeText {
                paragraphs,
                anchor: Some(match record.v_align {
                    -4108 => "ctr".to_string(),
                    -4107 => "b".to_string(),
                    _ => "t".to_string(),
                }),
                insets: (
                    (record.margins.0 / 0.75 * EMU_PER_PX) as i64,
                    (record.margins.1 / 0.75 * EMU_PER_PX) as i64,
                    (record.margins.2 / 0.75 * EMU_PER_PX) as i64,
                    (record.margins.3 / 0.75 * EMU_PER_PX) as i64,
                ),
                wrap: true,
                clip: false,
            }
        });
        Shape {
            geometry: geometry.to_string(),
            fill: record.fill_visible.then(|| packed_hex(record.fill)),
            line,
            adjusts: record
                .adjusts
                .iter()
                .enumerate()
                .map(|(at, value)| {
                    let name = if record.adjusts.len() == 1 { "adj".to_string() } else { format!("adj{}", at + 1) };
                    (name, (value * 100_000.0).round() as i64)
                })
                .collect(),
            path: None,
            flip_h: record.flip_h,
            flip_v: record.flip_v,
            rotation: (record.rotation * 60_000.0) as i32,
            text,
        }
    }

    pub(super) fn chart_of(&self, chart: &ChartRecord) -> Chart {
        let categories = chart.series.first().map(|series| series.xs.clone()).unwrap_or_default();
        Chart {
            kind: chart_kind(chart.chart_type).to_string(),
            plot: None,
            series: chart
                .series
                .iter()
                .map(|series| ChartSeries {
                    name: series.name.clone(),
                    values: series.values.iter().map(|v| Some(*v)).collect(),
                    line: series.color.map(|colour| ShapeLine {
                        head_end: None,
                        tail_end: None,
                        color: packed_hex(colour),
                        width: 28_575,
                        dash: None,
                        cap: None,
                    }),
                    marker: None,
                    points: Vec::new(),
                    labels: Vec::new(),
                    label_size: 9.0,
                    label_face: None,
                    label_pos: None,
                })
                .collect(),
            categories,
            value_axis: Some(ChartAxis {
                position: "l".to_string(),
                min: chart.axes[1].min,
                max: chart.axes[1].max,
                major_unit: None,
                major_tick: "none".to_string(),
                tick_labels: "nextTo".to_string(),
                number_format: None,
                line: None,
                major_gridline: None,
                size: 9.0,
                face: None,
                deleted: false,
                cross_between: Some("between".to_string()),
            }),
            category_axis: Some(ChartAxis {
                position: "b".to_string(),
                min: None,
                max: None,
                major_unit: None,
                major_tick: "none".to_string(),
                tick_labels: "nextTo".to_string(),
                number_format: None,
                line: None,
                major_gridline: None,
                size: 9.0,
                face: None,
                deleted: false,
                cross_between: Some("between".to_string()),
            }),
            legend: chart.has_legend.then(|| Legend {
                position: match chart.legend_position {
                    -4107 => "b".to_string(),
                    -4131 => "l".to_string(),
                    -4160 => "t".to_string(),
                    2 => "tr".to_string(),
                    _ => "r".to_string(),
                },
                frame: None,
                size: 9.0,
                face: None,
            }),
            shapes: Vec::new(),
            fill: Some("FFFFFF".to_string()),
            plot_fill: None,
        }
    }
}

impl ShapeRecord {
    /// A record with a new shape's defaults, as measured: filled accent1 with
    /// a darker outline of 1.5 points, white 11-point text in the theme's
    /// minor face, left and top aligned, margins of a tenth and a twentieth
    /// of an inch, moved and sized with its cells.
    pub(super) fn blank(id: u64, sheet: usize, name: String, kind: ShapeKind) -> Self {
        let text_colour = match kind {
            ShapeKind::TextBox => 0,
            _ => WHITE,
        };
        let mut record = Self {
            id,
            sheet,
            name,
            kind: ShapeKind::Auto(1),
            left: 0.0,
            top: 0.0,
            width: 0.0,
            height: 0.0,
            rotation: 0.0,
            flip_h: false,
            flip_v: false,
            visible: true,
            fill: THEME_COLOURS[ACCENT1 - 1],
            fill_theme: Some(ACCENT1),
            fill_visible: true,
            transparency: 0.0,
            line: SHAPE_OUTLINE,
            line_theme: Some(ACCENT1),
            line_visible: true,
            line_weight: 1.5,
            dash: 1,
            arrow_end: 1,
            runs: Vec::new(),
            h_align: -4131,
            v_align: -4160,
            margins: (7.2, 3.6, 7.2, 3.6),
            auto_size: false,
            adjusts: Vec::new(),
            placement: 1,
            on_action: String::new(),
            alt_text: String::new(),
            lock_aspect: false,
            paragraph_count: 1,
            original: None,
        };
        record.kind = kind;
        if let ShapeKind::Auto(kind) = record.kind {
            record.adjusts = default_adjusts(kind);
        }
        match &record.kind {
            // Measured: a text box is filled white, outlined in a light gray
            // of three quarters of a point, and written in black.
            ShapeKind::TextBox => {
                record.fill = WHITE;
                record.fill_theme = None;
                record.line = TEXTBOX_OUTLINE;
                record.line_theme = None;
                record.line_weight = 0.75;
            }
            // A connector is drawn in accent1 itself.
            ShapeKind::Line => {
                record.line = THEME_COLOURS[ACCENT1 - 1];
                record.fill_visible = false;
            }
            ShapeKind::Chart(_) => {
                record.fill = WHITE;
                record.fill_theme = None;
                record.line = 14_277_081;
                record.line_theme = None;
                record.line_weight = 0.75;
            }
            _ => {}
        }
        let _ = text_colour;
        record
    }

    /// The style the text is written in where a macro has not said: the
    /// theme's minor face at 11 points, white on a shape and black in a box.
    pub(super) fn default_style(&self) -> TextStyle {
        TextStyle {
            name: SHAPE_FACE.to_string(),
            size: 11.0,
            bold: false,
            italic: false,
            underline: false,
            color: if matches!(self.kind, ShapeKind::TextBox) { 0 } else { WHITE },
        }
    }

    pub(super) fn text(&self) -> String {
        self.runs.iter().map(|run| run.text.as_str()).collect()
    }

    /// Replace the text, keeping the style of the first run.
    pub(super) fn set_text(&mut self, text: &str) {
        let style = self.runs.first().map(|run| run.style.clone()).unwrap_or_else(|| self.default_style());
        // Measured: a carriage return and a line feed both come back as a
        // line feed, and only the carriage returns cut paragraphs.
        let text = text.replace("\r\n", "\r");
        self.paragraph_count = text.split('\r').count().max(1);
        let text = text.replace('\r', "\n");
        self.runs = if text.is_empty() { Vec::new() } else { vec![ShapeRunRecord { text, style }] };
        if self.auto_size {
            self.fit_to_text();
        }
    }

    /// The box a text box shrinks to when it sizes itself: measured, "tb" in
    /// Aptos Narrow 11 comes to 23.54 by 20.83 -- a line of 1.24 em with
    /// the margins round it, and about 0.42 em per character across. A
    /// crude ruler, and a better one than leaving the box as it was.
    pub(super) fn fit_to_text(&mut self) {
        let style = self.runs.first().map(|run| run.style.clone()).unwrap_or_else(|| self.default_style());
        let text = self.text();
        let lines = text.split('\n').count().max(1) as f64;
        let widest = text.split('\n').map(|line| line.chars().count()).max().unwrap_or(0) as f64;
        self.height = lines * style.size * 1.239 + self.margins.1 + self.margins.3;
        self.width = widest * style.size * 0.4155 + self.margins.0 + self.margins.2;
    }

    /// The runs a stretch of characters is made of, split so the stretch's
    /// boundaries fall between runs.
    pub(super) fn split_runs(&mut self, start: u32, length: Option<u32>) -> (usize, usize) {
        let start = start.max(1) as usize - 1;
        let whole: usize = self.runs.iter().map(|run| run.text.chars().count()).sum();
        let end = match length {
            Some(length) => (start + length as usize).min(whole),
            None => whole,
        };
        for cut in [start, end] {
            let mut seen = 0usize;
            let mut at = 0usize;
            while at < self.runs.len() {
                let count = self.runs[at].text.chars().count();
                if cut > seen && cut < seen + count {
                    let (head, tail): (String, String) = {
                        let text = &self.runs[at].text;
                        let split = text.char_indices().nth(cut - seen).map(|(i, _)| i).unwrap_or(text.len());
                        (text[..split].to_string(), text[split..].to_string())
                    };
                    let style = self.runs[at].style.clone();
                    self.runs[at].text = head;
                    self.runs.insert(at + 1, ShapeRunRecord { text: tail, style });
                    break;
                }
                seen += count;
                at += 1;
            }
        }
        let mut seen = 0usize;
        let mut first = self.runs.len();
        let mut last = self.runs.len();
        for (at, run) in self.runs.iter().enumerate() {
            let count = run.text.chars().count();
            if seen >= start && first == self.runs.len() {
                first = at;
            }
            seen += count;
            if seen >= end {
                last = at + 1;
                break;
            }
        }
        (first.min(last), last)
    }

    /// One answer for a stretch of the text, or None where its runs disagree.
    pub(super) fn uniform_style<T: PartialEq + Clone>(
        &self,
        start: u32,
        length: Option<u32>,
        read: impl Fn(&TextStyle) -> T,
    ) -> Option<T> {
        if self.runs.is_empty() {
            return Some(read(&self.default_style()));
        }
        let start = start.max(1) as usize - 1;
        let whole: usize = self.runs.iter().map(|run| run.text.chars().count()).sum();
        let end = match length {
            Some(length) => (start + length as usize).min(whole),
            None => whole,
        };
        let mut seen = 0usize;
        let mut first: Option<T> = None;
        for run in &self.runs {
            let count = run.text.chars().count();
            let run_start = seen;
            seen += count;
            if seen <= start || run_start >= end.max(start + 1) {
                continue;
            }
            let value = read(&run.style);
            if first.as_ref().is_some_and(|held| *held != value) {
                return None;
            }
            first = Some(value);
        }
        first.or_else(|| Some(read(&self.runs[0].style)))
    }
}

/// The name Excel gives a preset shape, by `msoAutoShapeType`.
pub(super) fn auto_shape_label(kind: i64) -> &'static str {
    auto_shape(kind).0
}

/// The adjust handles a preset starts with. Measured: a rounded rectangle
/// has one, at 0.16667; a rectangle has none.
pub(super) fn default_adjusts(kind: i64) -> Vec<f64> {
    match kind {
        5 => vec![0.16667],
        _ => Vec::new(),
    }
}

/// The Japanese name a shape also answers to. Measured:
/// `Shapes("テキスト ボックス 2")` finds `TextBox 2` on an Office whose shapes
/// are named in English.
pub(super) fn localized_alias(name: &str) -> Option<String> {
    let (word, number) = name.rsplit_once(' ')?;
    let english = match word {
        "正方形/長方形" => "Rectangle",
        "テキスト ボックス" => "TextBox",
        "楕円" => "Oval",
        "四角形: 角を丸くする" | "角丸四角形" => "Rounded Rectangle",
        "直線コネクタ" | "直線" => "Straight Connector",
        "グラフ" => "Chart",
        "図" => "Picture",
        "ひし形" => "Diamond",
        "二等辺三角形" => "Isosceles Triangle",
        "右矢印" => "Right Arrow",
        "左矢印" => "Left Arrow",
        "上矢印" => "Up Arrow",
        "下矢印" => "Down Arrow",
        "星: 5 pt" | "星 5" => "5-Point Star",
        "フローチャート: 処理" => "Flowchart: Process",
        "フローチャート: 判断" => "Flowchart: Decision",
        "フローチャート: 端子" => "Flowchart: Terminator",
        _ => return None,
    };
    Some(format!("{english} {number}"))
}
