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

/// The generated note label accepts a signed decimal integer prefix.
/// These decimal families are measured against Excel; newer script digits and
/// supplementary-plane digits remain ordinary custom-name text.
fn generated_comment_alias_number(name: &str) -> Option<u32> {
    if !name.get(..8)?.eq_ignore_ascii_case("Comment ") { return None; }
    let suffix = name.get(8..)?.trim_start();
    let suffix = suffix.strip_prefix('+').unwrap_or(suffix);
    let zeroes = [0x0030, 0x0660, 0x06f0, 0x0966, 0x09e6, 0x0a66, 0x0ae6,
        0x0b66, 0x0c66, 0x0ce6, 0x0d66, 0x0e50, 0x0ed0, 0x0f20, 0x17e0, 0xff10];
    let (mut value, mut any) = (0u32, false);
    for ch in suffix.chars() {
        let Some(digit) = zeroes.iter().find_map(|zero| {
            let digit = (ch as u32).wrapping_sub(*zero);
            (digit < 10).then_some(digit)
        }) else { break };
        value = value.checked_mul(10)?.checked_add(digit)?;
        if value > i32::MAX as u32 { return None; }
        any = true;
    }
    any.then_some(value)
}

#[derive(Clone)]
pub(super) struct ShapeRangeRecord {
    pub ids: Vec<u64>,
    pub sheet: Option<usize>,
    /// A range containing notes uses the worksheet drawing collection's
    /// lookup errors even after its note members have been deleted.
    pub legacy_notes: bool,
    pub ordinary_range: bool,
    pub ordinary_slots: Vec<u64>,
    pub ordinary_initialized: bool,
}

/// The workbook allocator owns ID blocks even after their last shape is deleted.
/// DrawingML IDs are local to a part; legacy VML IDs use this separate space.
#[derive(Default)]
pub(super) struct DrawingIdentityAllocator {
    owners: std::collections::BTreeMap<u32, usize>,
    cursors: std::collections::HashMap<usize, u32>,
    last_block: u32,
}

impl DrawingIdentityAllocator {
    fn observe(&mut self, sheet: usize, blocks: &[u32], ids: impl Iterator<Item = u32>) {
        for block in blocks.iter().copied().filter(|block| *block != 0) {
            self.owners.insert(block, sheet);
            self.last_block = self.last_block.max(block);
            self.cursors.entry(sheet).or_insert(block * 1024);
        }
        for id in ids.filter(|id| *id >= 1024) {
            let block = id / 1024;
            self.owners.insert(block, sheet);
            self.last_block = self.last_block.max(block);
            let cursor = self.cursors.entry(sheet).or_insert(id);
            *cursor = (*cursor).max(id);
        }
    }

    fn reserve(&mut self, sheet: usize) -> u32 {
        if let Some((&block, _)) = self.owners.iter().find(|(_, owner)| **owner == sheet) {
            return block;
        }
        self.last_block += 1;
        self.owners.insert(self.last_block, sheet);
        self.cursors.insert(sheet, self.last_block * 1024);
        self.last_block
    }

    fn next(&mut self, sheet: usize) -> u32 {
        self.reserve(sheet);
        let next = self.cursors[&sheet] + 1;
        let block = next / 1024;
        let next = if self.owners.get(&block).is_some_and(|owner| *owner == sheet) {
            next
        } else {
            self.last_block += 1;
            self.owners.insert(self.last_block, sheet);
            self.last_block * 1024
        };
        self.cursors.insert(sheet, next);
        next
    }

    pub(super) fn renumber(&mut self, moved: &dyn Fn(usize) -> Option<usize>) {
        for owner in self.owners.values_mut() {
            if *owner != usize::MAX { *owner = moved(*owner).unwrap_or(usize::MAX); }
        }
        self.cursors = std::mem::take(&mut self.cursors).into_iter()
            .filter_map(|(sheet, cursor)| Some((moved(sheet)?, cursor))).collect();
    }
}
use super::shapes_members::text_of_value;
use oxicells_core::ir::{
    Anchor, Chart, ChartAxis, ChartSeries, Drawing, DrawingKind, GroupTransform, Legend, Shape, ShapeLine,
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
    pub color_theme: Option<usize>,
    pub color_brightness: f64,
    pub scheme_color: Option<u8>,
}

#[derive(Debug, Clone, PartialEq)]
pub(super) struct ShapeRunRecord {
    pub text: String,
    pub style: TextStyle,
    /// This run is the paragraph separator after an OOXML `<a:p>`. Kept
    /// separate from soft line breaks so character indices remain intact.
    pub paragraph_end: bool,
}

fn source_shape_run_at(paragraph: &ShapeParagraph, character: usize) -> Option<&ShapeRun> {
    let mut start = 0usize;
    paragraph.runs.iter().find(|run| {
        let end = start + run.text.chars().count();
        let contains = start <= character && character < end;
        start = end;
        contains
    })
}

#[derive(Debug, Clone, PartialEq)]
pub(super) struct ShapeRecord {
    pub id: u64,
    /// DrawingML/VML identity exposed by Shape.ID, independent of the arena handle.
    pub non_visual_id: u32,
    /// Source identity before copying/rebasing. New objects have no source.
    pub source_drawing_id: Option<u32>,
    /// Package sheet owning the original payload, retained across copies.
    pub source_sheet_origin: Option<u32>,
    pub sheet: usize,
    /// A legacy note's VML shape, which lives in worksheet comments rather
    /// than the normal drawing part.
    pub comment_at: Option<CellAddress>,
    pub name: String,
    /// Live default chart names are reassigned when Excel copies a worksheet;
    /// explicitly assigned names and literal imported names are retained.
    pub name_is_automatic: bool,
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
    pub theme_palette: [i64; 12],
    pub fill_scheme_color: Option<u8>,
    pub fill_brightness: f64,
    pub fill_visible: bool,
    pub transparency: f64,
    pub line: i64,
    pub line_theme: Option<usize>,
    pub line_scheme_color: Option<u8>,
    pub line_brightness: f64,
    pub line_shade: Option<f32>,
    pub line_back: i64,
    pub line_back_theme: Option<usize>,
    pub line_back_scheme_color: Option<u8>,
    pub line_back_brightness: f64,
    pub line_visible: bool,
    /// MsoPatternType; Excel answers -2 for a mixed/default line.
    pub line_pattern: i64,
    pub line_weight: f64,
    pub line_transparency: f64,
    /// msoLineDashStyle, 1 solid.
    pub dash: i64,
    /// msoArrowheadStyle at the start, 1 none.
    pub arrow_start: i64,
    /// msoArrowheadStyle at the far end, 1 none.
    pub arrow_end: i64,
    /// msoArrowheadLength/Width at each end; None retains DrawingML's default.
    pub arrow_start_length: Option<i64>,
    pub arrow_start_width: Option<i64>,
    pub arrow_end_length: Option<i64>,
    pub arrow_end_width: Option<i64>,
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
    pub alt_text_utf16: Option<Vec<u16>>,
    pub alt_text_modified: bool,
    pub lock_aspect: bool,
    /// The msoConnectorType a connector was added as (1 straight, 2 elbow,
    /// 3 curve), 0 for anything else -- AddLine's line included.
    pub connector: i64,
    /// How many paragraphs the text was written as.
    pub paragraph_count: usize,
    /// The drawing the file held, when this record stands for one, so the
    /// picture's bytes and everything else read from the file survive.
    pub original: Option<Drawing>,
    /// Source drawing name for a runtime OLE duplicate, needed by the package writer.
    pub ole_clone_of: Option<String>,
    /// For a grouped child, the DrawingML name of its parent group. A virtual
    /// group record stores the members in `group_children` instead.
    pub group_name: Option<String>,
    pub group_parent_id: Option<u64>,
    pub group_children: Vec<u64>,
    /// OLEObjects.Group retains its legacy English naming channel.
    pub group_legacy: bool,
    pub group_bounds: Option<(f64, f64, f64, f64)>,
    pub group_transform: Option<GroupTransform>,
    pub group_nested: bool,
}

#[derive(Debug, Clone, PartialEq)]
pub(super) struct ChartRecord {
    pub chart_type: i64,
    /// Keep absent automatic paint distinct from an explicitly assigned RGB.
    pub fill: Option<String>,
    pub plot_fill: Option<String>,
    pub plot_modified: bool,
    pub of_pie: Option<oxicells_core::ir::ChartOfPieOptions>,
    pub series: Vec<SeriesRecord>,
    pub source_series_order: Vec<usize>,
    /// Imported text labels belong to the chart axis, independently of series order.
    pub category_axis_labels: Vec<String>,
    pub category_axis_reference: Option<String>,
    pub has_title: bool,
    pub title: String,
    pub title_from_data: bool,
    pub title_default_is_explicit: bool,
    pub title_generation: u64,
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

impl ChartRecord {
    pub(super) fn data_title(&self) -> Option<&str> {
        if self.series.len() != 1 && !matches!(chart_kind(self.chart_type), "pie" | "ofPie") { return None; }
        let first = self.source_series_order.iter().find_map(|index| self.series.get(*index))
            .or_else(|| self.series.first())?;
        (!first.name_auto).then_some(first.name.as_str())
    }
}

#[derive(Debug, Clone, PartialEq)]
pub(super) struct SeriesRecord {
    /// The name as shown, and the reference it came from when it did.
    pub name: String,
    /// Automatic display labels are not literal name expressions.
    pub name_auto: bool,
    pub name_ref: Option<String>,
    pub values_ref: Option<String>,
    /// Explicitly assigned worksheet data is live; imported caches retain file-open behavior.
    pub values_live: bool,
    pub bubble_sizes: Option<Vec<Option<f64>>>,
    pub bubble_sizes_ref: Option<String>,
    pub shaded_bubbles: bool,
    pub x_ref: Option<String>,
    /// API values of explicitly assigned categories, before display-text conversion.
    pub typed_xs: Option<Vec<Value>>,
    /// An omitted category expression uses the chart family's default labels.
    pub implicit_categories: bool,
    /// Imported categories have completed the file-open normalization step.
    pub categories_imported: bool,
    pub category_values_filtered: bool,
    pub text_category_gaps: bool,
    pub values: Vec<f64>,
    /// Original sparse cache, retained until values are explicitly assigned.
    pub sparse_values: Option<Vec<Option<f64>>>,
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
    /// `Worksheet.OLEObjects`.
    OleObjects(usize),
    /// The `OLEObjects` object returned by `Duplicate`; Excel keeps it bound
    /// to the collection as it existed when `Duplicate` was called.
    OleObjectsSnapshot(usize, usize),
    Fill(u64),
    FillColor(u64),
    Line(u64),
    LineColor(u64),
    LineBackColor(u64),
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
    ChartGroups(u64),
    ChartGroup(u64),
    SeriesCollection(u64),
    /// A series by its one-based number.
    Series(u64, usize),
    SeriesFormat(u64, usize),
    SeriesFill(u64, usize),
    SeriesColor(u64, usize),
    Points(u64, usize),
    DataLabels(u64, usize),
    ChartTitle(u64, u64),
    Legend(u64),
    Axes(u64),
    /// xlCategory 1 or xlValue 2.
    Axis(u64, i64),
    AxisTitle(u64, i64),
    /// `TextFrame2.TextRange.Font`: the same dress as the characters' Font,
    /// answered the Office way (msoTriState, Singles).
    Font2(u64),
    /// One point of a series: the series' number and the point's, from 1.
    Point(u64, usize, usize),
    ChartArea(u64),
    PlotArea(u64),
    /// The boolean selects the plot area instead of the whole chart.
    ChartFormat(u64, bool),
    ChartFill(u64, bool),
    ChartColor(u64, bool),
    ChartInterior(u64, bool),
    Adjustments(u64),
    Paragraphs(u64),
    /// `Selection.ShapeRange` / `Shapes.Range(Array(...))`: several shapes
    /// held as one, by their place in the host's `shape_ranges`.
    ShapeRange(usize),
    /// The child collection of a grouped worksheet shape.
    GroupItems(u64),
    /// Image adjustments and crop state.
    PictureFormat(u64),
    /// The independent picture viewport/position object on PictureFormat.
    Crop(u64),
    /// OLE metadata exposed by Shape.OLEFormat.
    OleFormat(u64),
    /// The common Excel OLE wrapper returned by `OLEFormat.Object`.
    OleObject(u64),
    /// A grouped object returned by `OLEObjects.Group`.
    GroupObject(u64),
    OleBorder(u64),
    OleInterior(u64),
}

impl DrawingPart {
    pub(super) fn kind_name(self) -> &'static str {
        match self {
            DrawingPart::Shapes(_) => "Shapes",
            DrawingPart::OleObjects(_) | DrawingPart::OleObjectsSnapshot(..) => "OLEObjects",
            DrawingPart::Shape(_) => "Shape",
            DrawingPart::Fill(_) | DrawingPart::SeriesFill(..) | DrawingPart::ChartFill(..) => "FillFormat",
            DrawingPart::FillColor(_) | DrawingPart::LineColor(_) | DrawingPart::LineBackColor(_) | DrawingPart::SeriesColor(..) | DrawingPart::ChartColor(..) => {
                "ColorFormat"
            }
            DrawingPart::Line(_) => "LineFormat",
            DrawingPart::TextFrame(_) => "TextFrame",
            DrawingPart::TextFrame2(_) => "TextFrame2",
            DrawingPart::TextRange(_) => "TextRange2",
            DrawingPart::Characters(..) => "Characters",
            DrawingPart::CharactersFont(..) => "Font",
            DrawingPart::Font2(_) => "Font2",
            DrawingPart::ParagraphFormat(_) => "ParagraphFormat2",
            DrawingPart::ChartObjects(_) => "ChartObjects",
            DrawingPart::ChartObject(_) => "ChartObject",
            DrawingPart::Chart(_) => "Chart",
            DrawingPart::ChartGroups(_) => "ChartGroups",
            DrawingPart::ChartGroup(_) => "ChartGroup",
            DrawingPart::SeriesCollection(_) => "SeriesCollection",
            DrawingPart::Series(..) => "Series",
            DrawingPart::SeriesFormat(..) | DrawingPart::ChartFormat(..) => "ChartFormat",
            DrawingPart::ChartInterior(..) => "Interior",
            DrawingPart::Points(..) => "Points",
            DrawingPart::Point(..) => "Point",
            DrawingPart::DataLabels(..) => "DataLabels",
            DrawingPart::ChartTitle(..) => "ChartTitle",
            DrawingPart::Legend(_) => "Legend",
            DrawingPart::Axes(_) => "Axes",
            DrawingPart::Axis(..) => "Axis",
            DrawingPart::AxisTitle(..) => "AxisTitle",
            DrawingPart::ChartArea(_) => "ChartArea",
            DrawingPart::PlotArea(_) => "PlotArea",
            DrawingPart::Adjustments(_) => "Adjustments",
            DrawingPart::Paragraphs(_) => "TextRange2",
            DrawingPart::ShapeRange(_) => "ShapeRange",
            DrawingPart::GroupItems(_) => "GroupShapes",
            DrawingPart::PictureFormat(_) => "PictureFormat",
            DrawingPart::Crop(_) => "Crop",
            DrawingPart::OleFormat(_) => "OLEFormat",
            DrawingPart::OleObject(_) => "OLEObject",
            DrawingPart::GroupObject(_) => "GroupObject",
            DrawingPart::OleBorder(_) => "Border",
            DrawingPart::OleInterior(_) => "Interior",
        }
    }

    /// The part after the sheets were renumbered; None when its sheet went.
    pub(super) fn renumbered(self, moved: &dyn Fn(usize) -> Option<usize>) -> Option<Self> {
        Some(match self {
            DrawingPart::Shapes(sheet) => DrawingPart::Shapes(moved(sheet)?),
            DrawingPart::OleObjects(sheet) => DrawingPart::OleObjects(moved(sheet)?),
            DrawingPart::OleObjectsSnapshot(sheet, index) => DrawingPart::OleObjectsSnapshot(moved(sheet)?, index),
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

fn arrow_style(name: Option<&str>) -> i64 {
    match name {
        Some("triangle") => 2,
        Some("arrow") => 3,
        Some("stealth") => 4,
        Some("diamond") => 5,
        Some("oval") => 6,
        _ => 1,
    }
}

fn arrow_size(name: Option<&str>) -> Option<i64> {
    match name {
        Some("sm") => Some(1),
        Some("med") => Some(2),
        Some("lg") => Some(3),
        _ => None,
    }
}

fn arrow_size_name(size: i64) -> String {
    match size {
        1 => "sm",
        2 => "med",
        _ => "lg",
    }
    .to_string()
}

pub(super) fn chart_plot_options(kind: i64) -> oxicells_core::ir::ChartPlotOptions {
    use oxicells_core::ir::{ChartBarShape, ChartPlotOptions, ChartStacking};
    let mut options = ChartPlotOptions::default();
    options.three_dimensional = matches!(kind, 54..=56 | 60..=62 | 70 | 78 | 79 | 92..=112 | -4100 | -4101 | -4102 | -4098);
    options.depth_series = matches!(kind, -4100 | 98 | 105 | 112);
    options.bar_shape = match kind { 92..=98 => ChartBarShape::Cylinder, 99..=105 => ChartBarShape::Cone,
        106..=112 => ChartBarShape::Pyramid, _ => ChartBarShape::Box };
    options.explosion = matches!(kind, 69 | 70 | 80).then_some(25);
    options.horizontal = matches!(kind, 57..=62 | 95..=97 | 102..=104 | 109..=111);
    options.stacking = match kind {
        52 | 55 | 58 | 61 | 63 | 66 | 76 | 78 | 93 | 96 | 100 | 103 | 107 | 110 => ChartStacking::Stacked,
        53 | 56 | 59 | 62 | 64 | 67 | 77 | 79 | 94 | 97 | 101 | 104 | 108 | 111 => ChartStacking::Percent,
        _ => ChartStacking::None,
    };
    options.markers = match kind {
        4 | 63 | 64 | 73 | 75 | -4151 | 82 => Some(false),
        65..=67 | 72 | 74 | -4169 | 81 => Some(true),
        _ => None,
    };
    if matches!(kind, -4169 | 72..=75) { options.connected = Some(kind != -4169); }
    options.smooth = matches!(kind, 72 | 73);
    options.filled = kind == 82;
    if matches!(kind, 68 | 71) {
        options.of_pie = Some(oxicells_core::ir::ChartOfPieOptions {
            secondary: if kind == 71 { oxicells_core::ir::ChartSecondaryPlot::Bar }
                else { oxicells_core::ir::ChartSecondaryPlot::Pie },
            ..Default::default()
        });
    }
    options
}

pub(super) fn chart_type_of(chart: &Chart) -> i64 {
    use oxicells_core::ir::{ChartBarShape, ChartStacking};
    let options = &chart.plot_options;
    let level = match options.stacking { ChartStacking::None => 0, ChartStacking::Stacked => 1, ChartStacking::Percent => 2 };
    let markers = options.markers.unwrap_or(true);
    let exploded = options.explosion.unwrap_or(0) > 0;
    if options.three_dimensional {
        return match chart.kind.as_str() {
            "bar" => {
                let start = match options.bar_shape { ChartBarShape::Cylinder => 92, ChartBarShape::Cone => 99,
                    ChartBarShape::Pyramid => 106, ChartBarShape::Box => 54 };
                if options.depth_series { if start == 54 { -4100 } else { start + 6 } }
                else if options.horizontal { if start == 54 { 60 + level } else { start + 3 + level } }
                else { start + level }
            },
            "line" => -4101,
            "area" => match level { 0 => -4098, 1 => 78, _ => 79 },
            "pie" => if exploded { 70 } else { -4102 },
            _ => 51,
        };
    }
    match chart.kind.as_str() {
        "bar" => if options.horizontal { 57 + level } else { 51 + level },
        "line" => if markers { 65 + level } else { match level { 0 => 4, 1 => 63, _ => 64 } },
        "area" => match level { 0 => 1, 1 => 76, _ => 77 },
        "scatter" => if options.smooth { if markers { 72 } else { 73 } }
            else if options.connected.unwrap_or(true) { if markers { 74 } else { 75 } } else { -4169 },
        "radar" => if options.filled { 82 } else if markers { 81 } else { -4151 },
        "pie" => if exploded { 69 } else { 5 },
        "doughnut" => if exploded { 80 } else { -4120 },
        "bubble" => if chart.series.first().is_some_and(|s| s.shaded_bubbles) { 87 } else { 15 },
        "ofPie" => if options.of_pie.as_ref().is_some_and(|p| p.secondary == oxicells_core::ir::ChartSecondaryPlot::Bar) { 71 } else { 68 },
        _ => 51,
    }
}

/// "bar", "line", "pie", "area", "scatter", "doughnut" -- the chart part's
/// own word for an `xlChartType` -- and whether the bars stand up.
pub(super) fn chart_kind(chart_type: i64) -> &'static str {
    match chart_type {
        4 | 63 | 64 | 65 | 66 | 67 | -4101 => "line",
        5 | 69 | 70 | -4102 => "pie",
        68 | 71 => "ofPie",
        -4120 | 80 => "doughnut",
        1 | 76 | 77 | 78 | 79 | -4098 => "area",
        -4169 | 72 | 73 | 74 | 75 => "scatter",
        15 | 87 => "bubble",
        -4151 | 81 | 82 | -4152 | 91 => "radar",
        _ => "bar",
    }
}

pub(super) fn packed_hex(colour: i64) -> String {
    colour_from_packed(colour)
}

pub(super) fn line_pattern_name(value: i64) -> Option<&'static str> {
    const PATTERNS: [&str; 54] = [
        "pct5", "pct10", "pct20", "pct25", "pct30", "pct40", "pct50", "pct60", "pct70", "pct75", "pct80", "pct90",
        "dkHorz", "dkVert", "dkDnDiag", "dkUpDiag", "smCheck", "trellis", "ltHorz", "ltVert", "ltDnDiag", "ltUpDiag",
        "smGrid", "dotDmnd", "wdDnDiag", "wdUpDiag", "dashUpDiag", "dashDnDiag", "narVert", "narHorz", "dashVert",
        "dashHorz", "lgConfetti", "lgGrid", "horzBrick", "lgCheck", "smConfetti", "zigZag", "solidDmnd", "diagBrick",
        "openDmnd", "plaid", "sphere", "weave", "dotGrid", "divot", "shingle", "wave", "horz", "vert", "cross",
        "dnDiag", "upDiag", "diagCross",
    ];
    usize::try_from(value.checked_sub(1)?).ok().and_then(|index| PATTERNS.get(index).copied())
}

pub(super) fn line_pattern_value(value: &str) -> Option<i64> {
    (1..=54).find(|candidate| line_pattern_name(*candidate) == Some(value))
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
        self.adopt_drawing_identities();
        for sheet in 0..self.workbook.sheets.len() {
            let drawings = self.workbook.sheets[sheet].drawings.clone();
            for drawing in drawings {
                let (kind, base) = match &drawing.kind {
                    DrawingKind::Picture { .. } => (ShapeKind::Picture, "Picture"),
                    DrawingKind::Chart(chart) => (
                        ShapeKind::Chart(Box::new(ChartRecord {
                            chart_type: chart_type_of(chart),
                            fill: chart.fill.clone(),
                            plot_fill: chart.plot_fill.clone(),
                            plot_modified: false,
                            of_pie: chart.plot_options.of_pie.clone(),
                            category_axis_labels: chart.categories.clone(),
                            category_axis_reference: chart.category_reference.clone(),
                            source_series_order: chart.source_series_indices(),
                            series: chart
                                .series
                                .iter()
                                .enumerate()
                                .map(|(index, series)| {
                                    let text_categories = series.text_category_gaps || series.category_values.as_ref()
                                        .is_some_and(|values| values.iter().any(|value| matches!(value, oxicells_core::ir::ChartCategoryValue::Text(_))));
                                    SeriesRecord {
                                    name: if series.name.is_empty() && !series.explicit_empty_name && series.name_reference.is_none() {
                                        format!("系列{}", index + 1)
                                    } else { series.name.clone() },
                                    name_auto: series.name.is_empty() && !series.explicit_empty_name && series.name_reference.is_none(),
                                    categories_imported: true,
                                    category_values_filtered: series.category_values_filtered,
                                    text_category_gaps: series.text_category_gaps,
                                    name_ref: series.name_reference.clone(),
                                    values_ref: series.values_reference.clone(),
                                    values_live: false,
                                    bubble_sizes: series.bubble_sizes.clone(),
                                    bubble_sizes_ref: series.bubble_sizes_reference.clone(),
                                    shaded_bubbles: series.shaded_bubbles,
                                    x_ref: series.category_reference.clone(),
                                    implicit_categories: series.category_reference.is_none() && !series.text_category_gaps && series.category_values.as_ref()
                                        .map_or(chart.categories.is_empty(), Vec::is_empty),
                                    typed_xs: series.category_values.as_ref().filter(|xs| !xs.is_empty())
                                        .map(|xs| xs.iter().map(|x| match x {
                                            oxicells_core::ir::ChartCategoryValue::Number(n) => Value::Double(*n),
                                            oxicells_core::ir::ChartCategoryValue::Text(s) => Value::String(s.clone()),
                                            oxicells_core::ir::ChartCategoryValue::Blank => if text_categories { Value::String(String::new()) } else { Value::Empty },
                                        }).collect()).or_else(|| {
                                            (series.category_values.is_none() && !chart.categories.is_empty())
                                                .then(|| chart.categories.iter().cloned().map(Value::String).collect())
                                        }),
                                    values: series.values.iter().map(|v| v.unwrap_or(0.0)).collect(),
                                    sparse_values: Some(series.values.clone()),
                                    xs: series.category_values.as_ref().filter(|xs| !xs.is_empty())
                                        .map(|xs| xs.iter().map(|x| match x {
                                            oxicells_core::ir::ChartCategoryValue::Number(n) => text_of_value(&Value::Double(*n)),
                                            oxicells_core::ir::ChartCategoryValue::Text(s) => s.clone(),
                                            oxicells_core::ir::ChartCategoryValue::Blank => String::new(),
                                        }).collect()).unwrap_or_else(|| if !chart.categories.is_empty() {
                                            chart.categories.clone()
                                        } else { (1..=series.values.len()).map(|i| i.to_string()).collect() }),
                                    color: None,
                                    chart_type: None,
                                    has_labels: !series.labels.is_empty(),
                                }})
                                .collect(),
                            has_title: chart.title.is_some(),
                            title: chart.title.clone().unwrap_or_default(),
                            title_from_data: chart.title_from_data && chart.title == chart.title_from_data_cache,
                            title_default_is_explicit: chart.title_default_is_explicit && chart.title == chart.title_from_data_cache,
                            title_generation: 0,
                            has_legend: chart.legend.is_some(),
                            title_auto: false,
                            legend_position: match chart.legend.as_ref().map(|legend|legend.position.as_str()) {
                                Some("b")=>-4107,Some("l")=>-4131,Some("t")=>-4160,Some("tr")=>2,_=>-4152,
                            },
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
                    DrawingKind::Shape(shape) if shape.text_box => (ShapeKind::TextBox, "TextBox"),
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
                let is_ole = drawing.ole_prog_id.is_some();
                let number = if is_ole {
                    if let Some(number) = drawing.name.as_deref()
                        .and_then(|name| name.rsplit_once(' '))
                        .and_then(|(_, suffix)| suffix.parse::<u32>().ok())
                    {
                        let counter = self.ole_object_counters.entry(sheet).or_insert(0);
                        *counter = (*counter).max(number);
                        number
                    } else {
                        self.next_ole_object_number(sheet)
                    }
                } else {
                    self.next_shape_number(sheet)
                };
                let id = self.next_shape_id;
                self.next_shape_id += 1;
                let physical_id = drawing.non_visual_id.unwrap_or(number.saturating_add(1));
                let generated_number = physical_id.saturating_sub(1);
                let automatic = !is_ole && (drawing.name.is_none() || generated_shape_name(&kind, generated_number, true)
                    .is_some_and(|native| drawing.name.as_ref() == Some(&native)));
                let display = if automatic {
                    generated_shape_name(&kind, generated_number, false)
                        .unwrap_or_else(|| drawing.name.clone().unwrap_or_else(||format!("{base} {number}")))
                } else { drawing.name.clone().unwrap_or_else(||format!("{base} {number}")) };
                let mut record = ShapeRecord::blank(id, sheet, display, kind);
                if let Some(metadata) = &drawing.metadata {
                    record.alt_text = metadata.description.clone().unwrap_or_default();
                    record.alt_text_utf16 = metadata.utf16_description.clone();
                }
                record.non_visual_id = physical_id;
                record.source_drawing_id = Some(physical_id);
                record.source_sheet_origin = self.workbook.sheets[sheet].origin_id;
                if !is_ole {
                    let counter = self.shape_counters.entry(sheet).or_insert(0);
                    *counter = (*counter).max(record.non_visual_id.saturating_sub(1));
                }
                record.name_is_automatic = automatic;
                record.left = left;
                record.top = top;
                record.width = width;
                record.height = height;
                if let Some(placement) = drawing.ole_placement {
                    record.placement = placement;
                }
                record.group_name = drawing.group_name.clone();
                if let DrawingKind::Shape(shape) = &drawing.kind {
                    record.theme_palette = workbook_office_theme_palette(&self.workbook.theme_colors);
                    record.rotation = f64::from(shape.rotation) / 60_000.0;
                    record.flip_h = shape.flip_h;
                    record.flip_v = shape.flip_v;
                    if let Some(fill) = &shape.fill {
                        record.fill = colour_to_packed(Some(fill)).unwrap_or(record.fill);
                        record.fill_theme = shape.fill_theme;
                        record.fill_scheme_color = shape.fill_legacy_color_index;
                        record.fill_brightness = f64::from(shape.fill_brightness.unwrap_or(0.0));
                    } else {
                        record.fill_visible = false;
                    }
                    match &shape.line {
                        Some(line) => {
                            record.line = colour_to_packed(Some(&line.color)).unwrap_or(record.line);
                            record.line_theme = line.color_theme;
                            record.line_scheme_color = line.legacy_color_index;
                            record.line_brightness = f64::from(line.color_brightness.unwrap_or(0.0));
                            record.line_shade = line.color_shade;
                            record.line_pattern = line.pattern.as_deref().and_then(line_pattern_value).unwrap_or(-2);
                            record.line_back = line.back_color.as_deref().and_then(|color| colour_to_packed(Some(color))).unwrap_or(WHITE);
                            record.line_back_theme = line.back_color_theme;
                            record.line_back_scheme_color = line.back_legacy_color_index;
                            record.line_back_brightness = f64::from(line.back_color_brightness.unwrap_or(0.0));
                            record.line_weight = line.width as f64 / EMU_PER_PX * 0.75;
                            record.arrow_start = arrow_style(line.head_end.as_deref());
                            record.arrow_end = arrow_style(line.tail_end.as_deref());
                            record.arrow_start_length = arrow_size(line.head_end_length.as_deref());
                            record.arrow_start_width = arrow_size(line.head_end_width.as_deref());
                            record.arrow_end_length = arrow_size(line.tail_end_length.as_deref());
                            record.arrow_end_width = arrow_size(line.tail_end_width.as_deref());
                        }
                        None => record.line_visible = false,
                    }
                    if let Some(text) = &shape.text {
                        record.paragraph_count = text.paragraphs.len().max(1);
                        for (at, paragraph) in text.paragraphs.iter().enumerate() {
                            let base = TextStyle {
                                name: paragraph.face.clone().unwrap_or_else(|| SHAPE_FACE.to_string()),
                                size: f64::from(paragraph.size),
                                bold: paragraph.bold,
                                italic: paragraph.italic,
                                underline: false,
                                color: paragraph.color.as_deref().and_then(|c| colour_to_packed(Some(c))).unwrap_or(0),
                                color_theme: None,
                                color_brightness: 0.0,
                                scheme_color: None,
                            };
                            if paragraph.runs.is_empty() {
                                record.runs.push(ShapeRunRecord { text: paragraph.text.clone(), style: base.clone(), paragraph_end: false });
                            } else {
                                record.runs.extend(paragraph.runs.iter().map(|run| ShapeRunRecord {
                                    text: run.text.clone(),
                                    style: TextStyle {
                                        name: run.face.clone().or_else(|| paragraph.face.clone()).unwrap_or_else(|| SHAPE_FACE.to_string()),
                                        size: f64::from(run.size.unwrap_or(paragraph.size)),
                                        bold: run.bold,
                                        underline: run.underline,
                                        italic: run.italic.unwrap_or(paragraph.italic),
                                        color: run.color.as_deref().or(paragraph.color.as_deref())
                                            .and_then(|c| colour_to_packed(Some(c))).unwrap_or(base.color),
                                        color_theme: run.color_theme.or(base.color_theme),
                                        color_brightness: run.color_brightness.map(f64::from).unwrap_or(base.color_brightness),
                                        scheme_color: run.legacy_color_index,
                                    },
                                    paragraph_end: false,
                                }));
                            }
                            if at + 1 < text.paragraphs.len() {
                                let style = record.runs.last().map(|run| run.style.clone()).unwrap_or(base);
                                record.runs.push(ShapeRunRecord { text: "\n".into(), style, paragraph_end: true });
                            }
                        }
                    }
                }
                record.original = Some(drawing);
                record.on_action = record.original.as_ref().and_then(|drawing| drawing.ole_on_action.clone()).unwrap_or_default();
                self.shapes.push(record);
            }
            self.adopt_shape_groups(sheet);
        }
        self.adopt_comment_shapes();
    }

    fn adopt_shape_groups(&mut self, sheet: usize) {
        let mut names = Vec::<String>::new();
        for shape in self.shapes.iter().filter(|shape| shape.sheet == sheet) {
            if let Some(name) = &shape.group_name {
                if !names.iter().any(|held| held == name) {
                    names.push(name.clone());
                }
            }
        }
        for name in names {
            let children: Vec<u64> = self.shapes.iter()
                .filter(|shape| shape.sheet == sheet && shape.group_name.as_deref() == Some(&name))
                .map(|shape| shape.id)
                .collect();
            let Some(first_child) = children.first().and_then(|id| self.shape(*id).ok()).cloned() else { continue };
            let Some(mut group_drawing) = first_child.original.clone() else { continue };
            group_drawing.grouped = false;
            group_drawing.frame = None;
            group_drawing.group_name = None;
            group_drawing.group_non_visual_id = None;
            group_drawing.group_rotation = 0;
            let bounds = self.points_of_drawing(sheet, &group_drawing);
            let id = self.next_shape_id;
            self.next_shape_id += 1;
            let mut group = ShapeRecord::blank(id, sheet, name, ShapeKind::Other);
            group.non_visual_id = first_child.original.as_ref().and_then(|drawing| drawing.group_non_visual_id)
                .unwrap_or_else(|| self.shapes.iter().filter(|shape| shape.sheet == sheet)
                    .map(|shape| shape.non_visual_id).max().unwrap_or(1).saturating_add(1));
            group.name_is_automatic = group.name == generated_group_name(group.non_visual_id.saturating_sub(1), true);
            if group.name_is_automatic {
                group.name = generated_group_name(group.non_visual_id.saturating_sub(1), false);
            }
            let counter = self.shape_counters.entry(sheet).or_insert(0);
            *counter = (*counter).max(group.non_visual_id.saturating_sub(1));
            group.left = bounds.0;
            group.top = bounds.1;
            group.width = bounds.2;
            group.height = bounds.3;
            group.rotation = first_child.original.as_ref().map_or(0.0, |drawing| f64::from(drawing.group_rotation) / 60_000.0);
            group.group_children = children.clone();
            group.group_bounds = Some(bounds);
            group.group_transform = first_child.original.as_ref().and_then(|drawing| drawing.group_transform);
            group.group_nested = first_child.original.as_ref().is_some_and(|drawing| drawing.group_nested);
            for child in children {
                if let Ok(record) = self.shape_mut(child) {
                    record.group_parent_id = Some(id);
                }
            }
            // The flattened IR retains anchor order through the first child.
            // Put its synthetic parent there, rather than after later anchors.
            let position = self.shapes.iter().position(|shape| shape.id == first_child.id)
                .unwrap_or(self.shapes.len());
            self.shapes.insert(position, group);
        }
    }

    fn adopt_drawing_identities(&mut self) {
        for (sheet, held) in self.workbook.sheets.iter().enumerate() {
            self.drawing_identities.observe(sheet, &held.drawing_id_blocks,
                held.comments.iter().filter_map(|comment| comment.shape_id)
                    .chain(held.drawings.iter().filter(|drawing| drawing.ole_prog_id.is_some()).filter_map(|drawing| drawing.non_visual_id)));
        }
        let mut order: Vec<_> = (0..self.workbook.sheets.len()).collect();
        order.sort_by_key(|sheet| (*sheet != self.active_sheet, *sheet));
        for sheet in order {
            let held = &self.workbook.sheets[sheet];
            if held.drawings.is_empty() && held.comments.is_empty() { continue; }
            if !self.drawing_identities.owners.values().any(|owner| *owner == sheet) {
                // On import, ordinary DrawingML groups occupy the remaining
                // group slots, with the active sheet loaded first.
                let block = (1..).find(|block| !self.drawing_identities.owners.contains_key(block)).unwrap();
                self.drawing_identities.observe(sheet, &[block], std::iter::empty());
            }
        }
    }

    fn update_drawing_id_blocks(&mut self, sheet: usize) {
        self.workbook.sheets[sheet].drawing_id_blocks = self.drawing_identities.owners.iter()
            .filter_map(|(block, owner)| (*owner == sheet).then_some(*block)).collect();
    }

    pub(super) fn reserve_drawing_identity_group(&mut self, sheet: usize) {
        self.drawing_identities.reserve(sheet);
        self.update_drawing_id_blocks(sheet);
    }

    pub(super) fn next_legacy_shape_id(&mut self, sheet: usize) -> u32 {
        let id = self.drawing_identities.next(sheet);
        self.update_drawing_id_blocks(sheet);
        id
    }

    pub(super) fn comment_shape_number(&self, sheet: usize, id: u32) -> Option<u32> {
        let index = self.drawing_identities.owners.iter().filter(|(_, owner)| **owner == sheet)
            .position(|(block, _)| *block == id / 1024)?;
        Some(index as u32 * 1024 + id % 1024)
    }

    fn adopt_comment_shapes(&mut self) {
        for sheet in 0..self.workbook.sheets.len() {
            let comments = self.workbook.sheets[sheet].comments.clone();
            for comment in comments {
                if comment.author.strip_prefix("tc=").is_some_and(|id| self.workbook.sheets[sheet].threaded_comments.iter().any(|thread| thread.cell == comment.cell && thread.id.eq_ignore_ascii_case(id))) { continue; }
                let number = comment.shape_id.and_then(|id| self.comment_shape_number(sheet, id))
                    .unwrap_or_else(|| self.next_comment_shape_number(sheet));
                let counter = self.comment_shape_counters.entry(sheet).or_default();
                *counter = (*counter).max(number);
                let at = CellAddress { sheet, row: comment.cell.0 + 1, column: comment.cell.1 };
                let id = self.next_shape_id;
                self.next_shape_id += 1;
                let mut record = ShapeRecord::blank(id, sheet, format!("Comment {number}"), ShapeKind::Other);
                record.non_visual_id = comment.shape_id.unwrap_or_else(|| self.drawing_identities.next(sheet));
                record.name_is_automatic = comment.shape_name.is_none();
                if let Some(name) = &comment.shape_name { record.name = name.clone(); }
                record.comment_at = Some(at);
                record.left = self.anchor_point(sheet, comment.from.col, comment.from.col_off, comment.from.row, comment.from.row_off).0;
                record.top = self.anchor_point(sheet, comment.from.col, comment.from.col_off, comment.from.row, comment.from.row_off).1;
                record.width = f64::from(comment.size.0);
                record.height = f64::from(comment.size.1);
                record.visible = comment.visible;
                record.fill = colour_to_packed(comment.fill.as_deref()).unwrap_or(14_811_135);
                record.fill_theme = None;
                record.fill_visible = comment.fill_visible;
                record.transparency = f64::from(comment.fill_transparency);
                record.line = colour_to_packed(comment.line_color.as_deref()).unwrap_or(0);
                record.line_theme = None;
                record.line_visible = comment.line_visible;
                record.line_weight = f64::from(comment.line_weight);
                record.line_transparency = f64::from(comment.line_transparency);
                record.placement = 3;
                record.margins = (
                    comment.text.insets.0 as f64 / 12_700.0,
                    comment.text.insets.1 as f64 / 12_700.0,
                    comment.text.insets.2 as f64 / 12_700.0,
                    comment.text.insets.3 as f64 / 12_700.0,
                );
                for (index, paragraph) in comment.text.paragraphs.iter().enumerate() {
                    let base = TextStyle {
                        name: paragraph.face.clone().unwrap_or_else(|| SHAPE_FACE.to_string()),
                        size: f64::from(paragraph.size),
                        bold: paragraph.bold,
                        italic: paragraph.italic,
                        underline: false,
                        color: paragraph.color.as_deref().and_then(|color| colour_to_packed(Some(color))).unwrap_or(0),
                        color_theme: None,
                        color_brightness: 0.0,
                        scheme_color: None,
                    };
                    if paragraph.runs.is_empty() {
                        record.runs.push(ShapeRunRecord { text: paragraph.text.clone(), style: base.clone(), paragraph_end: false });
                    } else {
                        record.runs.extend(paragraph.runs.iter().map(|run| ShapeRunRecord {
                            text: run.text.clone(), paragraph_end: false,
                            style: TextStyle {
                                name: run.face.clone().unwrap_or_else(|| base.name.clone()),
                                size: f64::from(run.size.unwrap_or(paragraph.size)), bold: run.bold,
                                italic: run.italic.unwrap_or(paragraph.italic), underline: run.underline,
                                color: run.color.as_deref().and_then(|color| colour_to_packed(Some(color))).unwrap_or(base.color),
                                color_theme: run.color_theme, color_brightness: f64::from(run.color_brightness.unwrap_or(0.0)),
                                scheme_color: run.legacy_color_index,
                            },
                        }));
                    }
                    if index + 1 < comment.text.paragraphs.len() {
                        let style = record.runs.last().map(|run| run.style.clone()).unwrap_or(base);
                        record.runs.push(ShapeRunRecord { text: "\n".into(), style, paragraph_end: true });
                    }
                }
                record.paragraph_count = comment.text.paragraphs.len().max(1);
                self.comment_shape_aliases.entry(sheet).or_default().insert(number, id);
                self.shapes.push(record);
            }
        }
    }

    fn anchor_point(&self, sheet: usize, col: u32, col_off: i64, row: u32, row_off: i64) -> (f64, f64) {
        let x = (0..col).map(|column| f64::from(self.column_px(sheet, column))).sum::<f64>()
            + col_off as f64 / EMU_PER_PX;
        let y = (1..=row).map(|row| f64::from(self.row_px(sheet, row))).sum::<f64>()
            + row_off as f64 / EMU_PER_PX;
        (x * 0.75, y * 0.75)
    }

    pub(super) fn add_comment_shape(&mut self, at: CellAddress) {
        let existing = self.shapes.iter().any(|shape| shape.comment_at == Some(at));
        if existing { return; }
        let shape_id = self.next_legacy_shape_id(at.sheet);
        let number = self.next_comment_shape_number(at.sheet);
        let id = self.next_shape_id;
        self.next_shape_id += 1;
        let cell_left = (0..at.column).map(|column| f64::from(self.column_px(at.sheet, column))).sum::<f64>() * 0.75;
        let cell_top = (1..at.row).map(|row| f64::from(self.row_px(at.sheet, row))).sum::<f64>() * 0.75;
        let cell_width = f64::from(self.column_px(at.sheet, at.column)) * 0.75;
        let mut record = ShapeRecord::blank(id, at.sheet, format!("Comment {number}"), ShapeKind::Other);
        record.non_visual_id = shape_id;
        self.comment_shape_aliases.entry(at.sheet).or_default().insert(number, id);
        record.comment_at = Some(at);
        record.left = cell_left + cell_width + 11.25;
        record.top = (cell_top - 7.5).max(1.5);
        record.width = 96.0;
        record.height = 59.25;
        record.visible = false;
        record.fill = 14_811_135;
        record.fill_theme = None;
        record.fill_visible = true;
        record.line = 0;
        record.line_theme = None;
        record.line_visible = true;
        record.line_weight = 0.75;
        record.placement = 3;
        record.margins = (7.2, 3.6, 7.2, 3.6);
        if let Some(note) = self.notes.iter().find(|note| note.at == at) {
            let text = note.text.clone();
            if !text.is_empty() {
                record.runs = vec![ShapeRunRecord { text, paragraph_end: false, style: TextStyle {
                    name: SHAPE_FACE.to_string(), size: 9.0, bold: false, italic: false, underline: false, color: 0,
                    color_theme: None, color_brightness: 0.0, scheme_color: None,
                }}];
            }
        }
        self.shapes.push(record);
    }

    pub(super) fn comment_shape_alias_owner(&self, sheet: usize, name: &str) -> Option<u64> {
        let number = generated_comment_alias_number(name)?;
        self.comment_shape_aliases.get(&sheet)?.get(&number).copied()
    }

    pub(super) fn next_comment_shape_number(&mut self, sheet: usize) -> u32 {
        let counter = self.comment_shape_counters.entry(sheet).or_insert(0);
        *counter += 1;
        *counter
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
        let (x, y, w, h) = if drawing.grouped {
            drawing.frame.map_or((x, y, w, h), |frame| {
                let child_x = x + w * frame.x;
                let child_y = y + h * frame.y;
                let child_width = w * frame.w;
                let child_height = h * frame.h;
                if drawing.group_rotation == 0 {
                    (child_x, child_y, child_width, child_height)
                } else {
                    let angle = f64::from(drawing.group_rotation) / 60_000.0 * std::f64::consts::PI / 180.0;
                    let (sin, cos) = angle.sin_cos();
                    let group_center_x = x + w / 2.0;
                    let group_center_y = y + h / 2.0;
                    let child_center_x = child_x + child_width / 2.0;
                    let child_center_y = child_y + child_height / 2.0;
                    let dx = child_center_x - group_center_x;
                    let dy = child_center_y - group_center_y;
                    let center_x = group_center_x + dx * cos - dy * sin;
                    let center_y = group_center_y + dx * sin + dy * cos;
                    (center_x - child_width / 2.0, center_y - child_height / 2.0, child_width, child_height)
                }
            })
        } else {
            (x, y, w, h)
        };
        (x * 0.75, y * 0.75, w * 0.75, h * 0.75)
    }

    /// The anchor of a point on the sheet, in EMU within its cell.
    pub(super) fn anchor_at(&self, sheet: usize, x_pt: f64, y_pt: f64) -> Anchor {
        let x_px = x_pt / 0.75;
        let y_px = y_pt / 0.75;
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
        let col_off = ((x_px - reached) * EMU_PER_PX).round() as i64;
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
        let row_off = ((y_px - down) * EMU_PER_PX).round() as i64;
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

    pub(super) fn next_ole_object_number(&mut self, sheet: usize) -> u32 {
        let counter = self.ole_object_counters.entry(sheet).or_insert(0);
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
            .filter(|record| record.comment_at.is_none() && record.group_children.is_empty() && (record.visible || record.original.is_some()))
            .map(|record| self.drawing_of(record))
            .collect();
        if let Some(held) = self.workbook.sheets.get_mut(sheet) {
            held.drawings = drawings;
        }
        let comment_shapes: Vec<ShapeRecord> = records.into_iter().filter(|record| record.comment_at.is_some()).collect();
        for record in comment_shapes {
            let Some(at) = record.comment_at else { continue };
            if let Some(note) = self.notes.iter_mut().find(|note| note.at == at) {
                note.visible = record.visible;
                let text = record.text();
                if note.text != text { note.text = text; note.utf16_text = None; }
            }
            let from = self.anchor_at(sheet, record.left, record.top);
            let to = self.anchor_at(sheet, record.left + record.width, record.top + record.height);
            let comment_text = self.shape_of(&record).text;
            if let Some(comment) = self.workbook.sheets.get_mut(sheet)
                .and_then(|held| held.comments.iter_mut().find(|comment| comment.cell == (at.row - 1, at.column)))
            {
                comment.from = from;
                comment.to = Some(to);
                comment.size = (record.width as f32, record.height as f32);
                comment.visible = record.visible;
                comment.fill = Some(packed_hex(record.fill));
                comment.fill_visible = record.fill_visible;
                comment.fill_transparency = record.transparency as f32;
                // A fresh Excel note has an automatic/default line color
                // (`strokecolor="none [81]"`) that reads as RGB 0. Keeping
                // the IR value empty lets a read-only Shape access preserve
                // that state instead of turning it into an explicit black
                // color in the next SaveCopyAs.
                if record.line != 0 || comment.line_color.is_some() {
                    comment.line_color = Some(packed_hex(record.line));
                }
                comment.line_visible = record.line_visible;
                comment.line_weight = record.line_weight as f32;
                comment.line_transparency = record.line_transparency as f32;
                if let Some(text) = comment_text {
                    comment.text = text;
                }
            }
        }
    }

    /// Commit the session-local legacy note objects to the workbook IR before
    /// SaveCopyAs. Excel saves both the comment text part and its VML shape.
    pub(super) fn sync_legacy_notes_for_save(&mut self) {
        for sheet in 0..self.workbook.sheets.len() {
            let original = std::mem::take(&mut self.workbook.sheets[sheet].comments);
            let mut comments: Vec<_> = original.iter().filter(|comment| comment.author.strip_prefix("tc=").is_some_and(|id| self.threaded_notes.iter().any(|thread| !thread.deleted && thread.at.sheet == sheet && thread.at.row == comment.cell.0 + 1 && thread.at.column == comment.cell.1 && thread.id.eq_ignore_ascii_case(id)))).cloned().collect();
            for note in self.notes.iter().filter(|note| note.at.sheet == sheet) {
                let mut comment = original.iter()
                    .find(|comment| comment.cell == (note.at.row.saturating_sub(1), note.at.column))
                    .cloned()
                    .unwrap_or_else(|| oxicells_core::ir::Comment {
                        position: None,
                        shape_id: None,
                        shape_name: None,
                        author: note.author.clone(),
                        utf16_text: None,
                        from: Anchor { col: note.at.column, col_off: 0, row: note.at.row.saturating_sub(1), row_off: 0 },
                        size: (96.0, 59.25),
                        to: None,
                        text: ShapeText {
                            paragraphs: vec![ShapeParagraph {
                                text: String::new(), runs: Vec::new(), align: None, size: 9.0,
                                bold: false, italic: false, face: Some(SHAPE_FACE.into()), charset: None,
                                pitch_family: None, color: None, line_pitch: None, line_scale: None,
                            }],
                            anchor: Some("t".into()), insets: (91_440, 45_720, 91_440, 45_720), wrap: true, clip: true,
                        },
                        fill: Some("FFFFE1".into()),
                        fill_visible: true,
                        fill_transparency: 0.0,
                        line_color: None,
                        line_visible: true,
                        line_weight: 0.75,
                        line_transparency: 0.0,
                        cell: (note.at.row.saturating_sub(1), note.at.column),
                        visible: note.visible,
                    });
                comment.author = note.author.clone();
                comment.utf16_text = note.utf16_text.clone();
                comment.cell = (note.at.row.saturating_sub(1), note.at.column);
                comment.visible = note.visible;
                if let Some(record) = self.shapes.iter().find(|shape| shape.comment_at == Some(note.at)) {
                    if record.non_visual_id >= 1024 { comment.shape_id = Some(record.non_visual_id); }
                    comment.shape_name = (!record.name_is_automatic).then(|| record.name.clone());
                    let from = self.anchor_at(sheet, record.left, record.top);
                    let to = Some(self.anchor_at(sheet, record.left + record.width, record.top + record.height));
                    if comment.position.is_none() || comment.from != from || comment.to != to {
                        comment.position = Some((record.left, record.top));
                    }
                    comment.from = from;
                    comment.to = to;
                    comment.size = (record.width as f32, record.height as f32);
                    comment.fill = Some(packed_hex(record.fill));
                    comment.fill_visible = record.fill_visible;
                    comment.fill_transparency = record.transparency as f32;
                    if record.line != 0 || comment.line_color.is_some() {
                        comment.line_color = Some(packed_hex(record.line));
                    }
                    comment.line_visible = record.line_visible;
                    comment.line_weight = record.line_weight as f32;
                    comment.line_transparency = record.line_transparency as f32;
                    let baseline = (
                        comment.text.insets.0 as f64 / 12_700.0,
                        comment.text.insets.1 as f64 / 12_700.0,
                        comment.text.insets.2 as f64 / 12_700.0,
                        comment.text.insets.3 as f64 / 12_700.0,
                    );
                    if baseline != record.margins {
                        comment.text.insets = (
                            (record.margins.0 * 12_700.0).round() as i64,
                            (record.margins.1 * 12_700.0).round() as i64,
                            (record.margins.2 * 12_700.0).round() as i64,
                            (record.margins.3 * 12_700.0).round() as i64,
                        );
                    }
                }
                let held_text = comment.text.paragraphs.iter().map(|paragraph| paragraph.text.as_str()).collect::<Vec<_>>().join("\n");
                if held_text != note.text {
                    let shape_text = self.shapes.iter().find(|shape| shape.comment_at == Some(note.at) && shape.text() == note.text)
                        .and_then(|shape| self.shape_of(shape).text);
                    if let Some(shape_text) = shape_text {
                        comment.text.paragraphs = shape_text.paragraphs;
                    } else {
                    let style = comment.text.paragraphs.first().cloned().unwrap_or(ShapeParagraph {
                        text: String::new(), runs: Vec::new(), align: None, size: 9.0,
                        bold: false, italic: false, face: Some(SHAPE_FACE.into()), charset: None,
                        pitch_family: None, color: None, line_pitch: None, line_scale: None,
                    });
                    comment.text.paragraphs = note.text.split('\n').map(|line| {
                        let mut paragraph = style.clone();
                        paragraph.text = line.to_string();
                        paragraph.runs.clear();
                        paragraph
                    }).collect();
                    }
                }
                comments.push(comment);
            }
            comments.sort_by_key(|comment| comment.cell);
            self.workbook.sheets[sheet].comments = comments;
        }
    }

    /// Save/SaveAs establishes a new source package. SaveCopyAs keeps the
    /// existing source, so its copied objects remain additions for later saves.
    pub(super) fn rebase_saved_drawing_sources(&mut self, saved: &oxicells_core::ir::Workbook) {
        for record in &mut self.shapes {
            if record.comment_at.is_some() || record.original.as_ref().is_some_and(|drawing|drawing.ole_prog_id.is_some()) { continue; }
            let Some(sheet)=self.workbook.sheets.get(record.sheet)
                .and_then(|live|saved.sheets.iter().find(|sheet|sheet.name==live.name)) else { continue };
            let Some(drawing)=sheet.drawings.iter().find(|drawing|drawing.non_visual_id==Some(record.non_visual_id)) else { continue };
            record.source_drawing_id=drawing.non_visual_id;
            record.source_sheet_origin=sheet.origin_id;
            record.original=Some(drawing.clone());
        }
    }

    pub(super) fn drawing_of(&self, record: &ShapeRecord) -> Drawing {
        let grouped = record.group_parent_id.is_some();
        let original = grouped.then_some(record.original.as_ref()).flatten();
        let parent = record.group_parent_id.and_then(|id| self.shape(id).ok());
        let from = parent.map_or_else(
            || original.map_or_else(|| self.anchor_at(record.sheet, record.left, record.top), |drawing| drawing.from),
            |group| self.anchor_at(record.sheet, group.left, group.top),
        );
        let to = if let Some(group) = parent {
            Some(self.anchor_at(record.sheet, group.left + group.width, group.top + group.height))
        } else {
            Some(self.anchor_at(record.sheet, record.left + record.width, record.top + record.height))
        };
        let extent = if grouped {
            original.and_then(|drawing| drawing.extent)
        } else {
            Some((
            (record.width * 12700.0).round() as i64,
            (record.height * 12700.0).round() as i64,
            ))
        };
        let kind = match &record.kind {
            ShapeKind::Picture | ShapeKind::Other => match &record.original {
                Some(original) => original.kind.clone(),
                None => DrawingKind::Other,
            },
            ShapeKind::Chart(chart) => {
                let plot_modified=chart.plot_modified;
                let mut chart=self.chart_of(chart);
                if let Some(Drawing {kind:DrawingKind::Chart(original),..})=&record.original {
                    chart.origin_id=original.origin_id;
                    if !plot_modified { chart.kind=original.kind.clone();chart.plot_options=original.plot_options.clone(); }
                    chart.plot=original.plot.clone();
                    chart.shapes=original.shapes.clone();
                    for (series, previous) in chart.series.iter_mut().zip(&original.series) {
                        if !plot_modified { series.explosion=previous.explosion; }
                        if series.line.is_none() { series.line=previous.line.clone(); }
                        if chart.kind==original.kind && chart.plot_options.markers==original.plot_options.markers {
                            series.marker=previous.marker.clone();
                        }
                        series.points=previous.points.clone();
                        series.labels=previous.labels.clone();
                        series.label_size=previous.label_size;
                        series.label_face=previous.label_face.clone();
                        series.label_pos=previous.label_pos.clone();
                    }
                    for (axis, prior) in [(&mut chart.value_axis,&original.value_axis),(&mut chart.category_axis,&original.category_axis)] {
                        if let (Some(current),Some(prior))=(axis.as_mut(),prior.as_ref()) {
                            let min=current.min;let max=current.max;let position=current.position.clone();
                            *current=prior.clone();current.min=min;current.max=max;current.position=position;
                        }
                    }
                }
                DrawingKind::Chart(chart)
            },
            ShapeKind::Auto(_) | ShapeKind::TextBox | ShapeKind::Line => {
                DrawingKind::Shape(self.shape_of(record))
            }
        };
        Drawing {
            name: Some(if record.name_is_automatic && !record.original.as_ref().is_some_and(|drawing| drawing.ole_prog_id.is_some()) {
                generated_shape_name(&record.kind, record.non_visual_id.saturating_sub(1), true)
                    .unwrap_or_else(||record.name.clone())
            } else { record.name.clone() }),
            non_visual_id: Some(record.non_visual_id),
            metadata: record.drawing_metadata(),
            ole_prog_id: record.original.as_ref().and_then(|drawing| drawing.ole_prog_id.clone()),
            ole_source_name: record.ole_clone_of.clone(),
            ole_clone_position: record.ole_clone_of.as_ref().map(|_| (record.left, record.top)),
            ole_embedded_package: record.original.as_ref().and_then(|drawing| drawing.ole_embedded_package.clone()),
            ole_package_is_empty: record.original.as_ref().and_then(|drawing| drawing.ole_package_is_empty),
            ole_clone_visible: record.ole_clone_of.as_ref().map(|_| record.visible),
            ole_placement: record.original.as_ref().and_then(|drawing| drawing.ole_prog_id.as_ref().map(|_| record.placement)),
            ole_type: record.original.as_ref().and_then(|drawing| drawing.ole_type),
            ole_display_as_icon: record.original.as_ref().and_then(|drawing| drawing.ole_display_as_icon),
            ole_auto_load: record.original.as_ref().and_then(|drawing| drawing.ole_auto_load),
            ole_enabled: record.original.as_ref().and_then(|drawing| drawing.ole_enabled),
            ole_locked: record.original.as_ref().and_then(|drawing| drawing.ole_locked),
            ole_print_object: record.original.as_ref().and_then(|drawing| drawing.ole_print_object),
            ole_shadow: record.original.as_ref().and_then(|drawing| drawing.ole_shadow),
            ole_on_action: record.original.as_ref().and_then(|drawing| {
                if drawing.ole_prog_id.is_some()
                    && record.on_action != drawing.ole_on_action.as_deref().unwrap_or_default()
                {
                    Some(record.on_action.clone())
                } else {
                    drawing.ole_on_action.clone()
                }
            }),
            ole_border_color: record.original.as_ref().and_then(|drawing| drawing.ole_border_color.clone()),
            ole_border_weight: record.original.as_ref().and_then(|drawing| drawing.ole_border_weight),
            ole_interior_color: record.original.as_ref().and_then(|drawing| drawing.ole_interior_color.clone()),
            from,
            to,
            extent,
            kind,
            frame: grouped.then(|| record.original.as_ref().and_then(|drawing| drawing.frame)).flatten(),
            grouped,
            group_name: parent.map(|group| if group.name_is_automatic && !group.group_legacy {
                generated_group_name(group.non_visual_id.saturating_sub(1), true)
            } else { group.name.clone() }),
            group_non_visual_id: parent.map(|group| group.non_visual_id),
            group_rotation: if grouped { original.map_or(0, |drawing| drawing.group_rotation) } else { 0 },
            group_transform: parent.and_then(|group| group.group_transform),
            group_child_transform: grouped.then_some((
                (record.left / 0.75 * EMU_PER_PX).round() as i64,
                (record.top / 0.75 * EMU_PER_PX).round() as i64,
                (record.width / 0.75 * EMU_PER_PX).round() as i64,
                (record.height / 0.75 * EMU_PER_PX).round() as i64,
            )),
            group_nested: parent.is_some_and(|group| group.group_nested),
        }
    }

    pub(super) fn shape_of(&self, record: &ShapeRecord) -> Shape {
        let geometry = match &record.kind {
            ShapeKind::Auto(kind) => auto_shape(*kind).1,
            ShapeKind::Line => "line",
            _ => "rect",
        };
        let line = record.line_visible.then(|| ShapeLine {
            head_end: arrow_name(record.arrow_start).map(str::to_string),
            tail_end: arrow_name(record.arrow_end).map(str::to_string),
            head_end_length: record.arrow_start_length.map(arrow_size_name),
            head_end_width: record.arrow_start_width.map(arrow_size_name),
            tail_end_length: record.arrow_end_length.map(arrow_size_name),
            tail_end_width: record.arrow_end_width.map(arrow_size_name),
            color: packed_hex(record.line),
            color_theme: record.line_theme,
            legacy_color_index: record.line_scheme_color,
            color_brightness: (record.line_brightness != 0.0).then_some(record.line_brightness as f32),
            color_shade: record.line_shade,
            pattern: line_pattern_name(record.line_pattern).map(str::to_string),
            back_color: (record.line_pattern != -2).then(|| packed_hex(record.line_back)),
            back_color_theme: record.line_back_theme,
            back_legacy_color_index: record.line_back_scheme_color,
            back_color_brightness: (record.line_back_brightness != 0.0).then_some(record.line_back_brightness as f32),
            back_color_shade: None,
            width: (record.line_weight / 0.75 * EMU_PER_PX) as i64,
            dash: dash_name(record.dash).map(str::to_string),
            cap: None,
        });
        let text = (!record.runs.is_empty() && !matches!(record.kind, ShapeKind::Line)).then(|| {
            let mut output = record.original.as_ref().and_then(|drawing| match &drawing.kind {
                DrawingKind::Shape(shape) => shape.text.clone(),
                _ => None,
            }).filter(|text| text.paragraphs.len() == record.paragraph_count);
            if let Some(mut original) = output.take() {
                let mut run_index = 0usize;
                for paragraph in &mut original.paragraphs {
                    let source_paragraph = paragraph.clone();
                    let mut runs = Vec::new();
                    let mut paragraph_text = String::new();
                    while let Some(run) = record.runs.get(run_index) {
                        run_index += 1;
                        if run.paragraph_end { break; }
                        let position = paragraph_text.chars().count();
                        let source_run = source_shape_run_at(&source_paragraph, position);
                        paragraph_text.push_str(&run.text);
                        if !run.text.is_empty() {
                            let source_size = source_run.and_then(|source| source.size).unwrap_or(source_paragraph.size);
                            let source_face = source_run.and_then(|source| source.face.as_deref())
                                .or(source_paragraph.face.as_deref()).unwrap_or(SHAPE_FACE);
                            let source_italic = source_run.and_then(|source| source.italic).unwrap_or(source_paragraph.italic);
                            let source_color = source_run.and_then(|source| source.color.as_deref())
                                .or(source_paragraph.color.as_deref())
                                .and_then(|color| colour_to_packed(Some(color)));
                            let source_theme = source_run.and_then(|source| source.color_theme);
                            let source_brightness = source_run.and_then(|source| source.color_brightness).map(f64::from).unwrap_or(0.0);
                            let source_scheme = source_run.and_then(|source| source.legacy_color_index);
                            runs.push(ShapeRun {
                                text: run.text.clone(),
                                bold: run.style.bold,
                                italic: (run.style.italic != source_italic).then_some(run.style.italic)
                                    .or_else(|| source_run.and_then(|source| source.italic)),
                                underline: run.style.underline,
                                face: (run.style.name != source_face).then(|| run.style.name.clone())
                                    .or_else(|| source_run.and_then(|source| source.face.clone())),
                                size: ((run.style.size as f32) != source_size).then_some(run.style.size as f32)
                                    .or_else(|| source_run.and_then(|source| source.size)),
                                color: (source_color != Some(run.style.color)).then(|| packed_hex(run.style.color))
                                    .or_else(|| source_run.and_then(|source| source.color.clone())),
                                color_theme: if source_theme != run.style.color_theme {
                                    run.style.color_theme
                                } else {
                                    source_run.and_then(|source| source.color_theme)
                                },
                                color_brightness: if source_brightness != run.style.color_brightness {
                                    (run.style.color_brightness != 0.0).then_some(run.style.color_brightness as f32)
                                } else {
                                    source_run.and_then(|source| source.color_brightness)
                                },
                                legacy_color_index: if source_scheme != run.style.scheme_color {
                                    run.style.scheme_color
                                } else {
                                    source_scheme
                                },
                            });
                        }
                    }
                    paragraph.text = paragraph_text;
                    paragraph.runs = runs;
                }
                original.anchor = Some(match record.v_align {
                    -4108 => "ctr".to_string(),
                    -4107 => "b".to_string(),
                    _ => "t".to_string(),
                });
                original.insets = (
                    (record.margins.0 / 0.75 * EMU_PER_PX) as i64,
                    (record.margins.1 / 0.75 * EMU_PER_PX) as i64,
                    (record.margins.2 / 0.75 * EMU_PER_PX) as i64,
                    (record.margins.3 / 0.75 * EMU_PER_PX) as i64,
                );
                original
            } else {
                let mut paragraphs = Vec::new();
                let mut paragraph_text = String::new();
                let mut paragraph_runs = Vec::new();
                let mut first_style: Option<TextStyle> = None;
                let to_paragraph = |text: String, runs: Vec<ShapeRun>, style: Option<TextStyle>| {
                    let style = style.unwrap_or_else(|| record.runs.first().map(|run| run.style.clone()).unwrap_or_else(|| record.default_style()));
                    ShapeParagraph {
                        text, runs,
                        align: Some(match record.h_align { -4108 => "ctr".to_string(), -4152 => "r".to_string(), _ => "l".to_string() }),
                        size: style.size as f32, bold: style.bold, italic: style.italic, face: Some(style.name),
                        charset: None, pitch_family: None, color: Some(packed_hex(style.color)), line_pitch: None, line_scale: None,
                    }
                };
                for run in &record.runs {
                    if run.paragraph_end {
                        paragraphs.push(to_paragraph(std::mem::take(&mut paragraph_text), std::mem::take(&mut paragraph_runs), first_style.take()));
                    } else {
                        paragraph_text.push_str(&run.text);
                        if !run.text.is_empty() {
                            first_style.get_or_insert_with(|| run.style.clone());
                            paragraph_runs.push(ShapeRun { text: run.text.clone(), bold: run.style.bold,
                                italic: Some(run.style.italic), underline: run.style.underline,
                                face: Some(run.style.name.clone()), size: Some(run.style.size as f32),
                                color: Some(packed_hex(run.style.color)),
                                color_theme: run.style.color_theme,
                                color_brightness: (run.style.color_brightness != 0.0).then_some(run.style.color_brightness as f32),
                                legacy_color_index: run.style.scheme_color });
                        }
                    }
                }
                if paragraphs.len() < record.paragraph_count || !paragraph_text.is_empty() || !paragraph_runs.is_empty() {
                    paragraphs.push(to_paragraph(paragraph_text, paragraph_runs, first_style));
                }
                ShapeText {
                    paragraphs,
                    anchor: Some(match record.v_align { -4108 => "ctr".to_string(), -4107 => "b".to_string(), _ => "t".to_string() }),
                    insets: ((record.margins.0 / 0.75 * EMU_PER_PX) as i64, (record.margins.1 / 0.75 * EMU_PER_PX) as i64,
                        (record.margins.2 / 0.75 * EMU_PER_PX) as i64, (record.margins.3 / 0.75 * EMU_PER_PX) as i64),
                    wrap: true, clip: false,
                }
            }
        });
        Shape {
            transform_offset: if record.group_parent_id.is_some() || !record.group_children.is_empty()
                || record.original.as_ref().is_some_and(|drawing| drawing.ole_prog_id.is_some()) {
                record.original.as_ref().and_then(|drawing| match &drawing.kind {
                    DrawingKind::Shape(shape) => shape.transform_offset, _ => None,
                })
            } else {
                Some(((record.left * 12700.0).round() as i64, (record.top * 12700.0).round() as i64))
            },
            geometry: geometry.to_string(),
            text_box: matches!(record.kind, ShapeKind::TextBox),
            fill: record.fill_visible.then(|| packed_hex(record.fill)),
            fill_theme: record.fill_theme,
            fill_legacy_color_index: record.fill_scheme_color,
            fill_brightness: (record.fill_brightness != 0.0).then_some(record.fill_brightness as f32),
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
        let categories = if chart.series.iter().all(|series| series.categories_imported) {
            chart.category_axis_labels.clone()
        } else { chart.series.first().map(|series| series.xs.clone()).unwrap_or_default() };
        let title = if chart.has_title {
                Some(if chart.title_from_data || chart.title_default_is_explicit { chart.data_title().unwrap_or(&chart.title).to_string() }
                    else { chart.title.clone() })
            }
            else if chart.title_auto && chart.series.len() == 1 && chart.series[0].name_ref.is_some() {
                Some(chart.series[0].name.clone())
            } else { None };
        Chart {
            category_reference: if chart.series.iter().all(|series| series.categories_imported) {
                chart.category_axis_reference.clone()
            } else { chart.series.first().and_then(|series| series.x_ref.clone()) },
            origin_id: None,
            // A single named series selected at chart creation acquires its title
            // when Excel saves the chart, while the live HasTitle remains false.
            title_from_data: chart.title_from_data,
            title_default_is_explicit: chart.title_default_is_explicit,
            title_from_data_cache: if chart.title_from_data || chart.title_default_is_explicit { title.clone() } else { None },
            source_series_order: if chart.source_series_order.iter().copied().eq(0..chart.series.len()) {
                Vec::new()
            } else { chart.source_series_order.clone() },
            title,
            kind: chart_kind(chart.chart_type).to_string(),
            plot_options: {
                let mut options = chart_plot_options(chart.chart_type);
                options.of_pie = chart.of_pie.clone();
                options
            },
            plot: None,
            series: chart
                .series
                .iter()
                .map(|series| ChartSeries {
                    name: if series.name_auto { String::new() } else { series.name.clone() },
                    explicit_empty_name: !series.name_auto && series.name.is_empty() && series.name_ref.is_none(),
                    explosion: chart_plot_options(chart.chart_type).explosion,
                    name_reference: series.name_ref.clone(),
                    values_reference: series.values_ref.clone(),
                    bubble_sizes: series.bubble_sizes.clone(),
                    bubble_sizes_reference: series.bubble_sizes_ref.clone(),
                    shaded_bubbles: series.shaded_bubbles,
                    category_reference: series.x_ref.clone(),
                    category_values_filtered: series.category_values_filtered,
                    text_category_gaps: series.text_category_gaps,
                    category_values: Some(series.typed_xs.as_ref().map(|xs| xs.iter().map(|value| {
                        use oxicells_core::ir::ChartCategoryValue;
                        if value.string_units().is_some() { ChartCategoryValue::Text(text_of_value(value)) }
                        else if matches!(value, Value::Empty) { ChartCategoryValue::Blank }
                        else { text_of_value(value).parse::<f64>().map(ChartCategoryValue::Number)
                            .unwrap_or_else(|_| ChartCategoryValue::Text(text_of_value(value))) }
                    }).collect()).unwrap_or_else(|| if series.x_ref.is_some() {
                        series.xs.iter().map(|text| text.parse::<f64>().map(oxicells_core::ir::ChartCategoryValue::Number)
                            .unwrap_or_else(|_| oxicells_core::ir::ChartCategoryValue::Text(text.clone()))).collect()
                    } else { Vec::new() })),
                    values: series.sparse_values.clone().unwrap_or_else(|| series.values.iter().map(|v| Some(*v)).collect()),
                    line: series.color.map(|colour| ShapeLine {
                        head_end: None,
                        tail_end: None,
                        head_end_length: None,
                        head_end_width: None,
                        tail_end_length: None,
                        tail_end_width: None,
                        color: packed_hex(colour),
                        color_theme: None,
                        legacy_color_index: None,
                        color_brightness: None,
                        color_shade: None,
                        pattern: None,
                        back_color: None,
                        back_color_theme: None,
                        back_legacy_color_index: None,
                        back_color_brightness: None,
                        back_color_shade: None,
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
                position: if chart_plot_options(chart.chart_type).horizontal { "b" } else { "l" }.to_string(),
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
                position: if chart_plot_options(chart.chart_type).horizontal { "l" } else { "b" }.to_string(),
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
            fill: chart.fill.clone(),
            plot_fill: chart.plot_fill.clone(),
        }
    }
}

impl ShapeRecord {
    pub(super) fn alt_text_units(&self) -> Vec<u16> {
        self.alt_text_utf16.clone().unwrap_or_else(|| self.alt_text.encode_utf16().collect())
    }

    pub(super) fn set_alt_text_units(&mut self, mut units: Vec<u16>) {
        // This property uses a NUL-terminated Automation string. Other BSTR
        // properties retain embedded NULs and must not use this truncation.
        if let Some(end) = units.iter().position(|unit| *unit == 0) { units.truncate(end); }
        self.alt_text_utf16 = String::from_utf16(&units).is_err().then(|| units.clone());
        self.alt_text = String::from_utf16_lossy(&units);
        self.alt_text_modified = true;
    }

    pub(super) fn drawing_metadata(&self) -> Option<oxicells_core::ir::DrawingMetadata> {
        let original = self.original.as_ref().and_then(|drawing| drawing.metadata.clone());
        if !self.alt_text_modified { return original; }
        let mut metadata = original.unwrap_or_default();
        metadata.description = (!self.alt_text.is_empty()).then(|| self.alt_text.clone());
        metadata.utf16_description = self.alt_text_utf16.clone();
        (!metadata.is_empty()).then_some(metadata)
    }

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
            non_visual_id: u32::try_from(id.saturating_add(1)).unwrap_or(u32::MAX),
            source_drawing_id: None,
            source_sheet_origin: None,
            sheet,
            comment_at: None,
            name,
            name_is_automatic: true,
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
            theme_palette: workbook_office_theme_palette(&[]),
            fill_scheme_color: None,
            fill_brightness: 0.0,
            fill_visible: true,
            transparency: 0.0,
            line_transparency: 0.0,
            line: SHAPE_OUTLINE,
            line_theme: Some(ACCENT1),
            line_scheme_color: None,
            line_brightness: 0.0,
            line_shade: Some(0.15),
            line_back: WHITE,
            line_back_theme: None,
            line_back_scheme_color: None,
            line_back_brightness: 0.0,
            line_visible: true,
            line_pattern: -2,
            line_weight: 1.5,
            dash: 1,
            arrow_start: 1,
            arrow_end: 1,
            arrow_start_length: None,
            arrow_start_width: None,
            arrow_end_length: None,
            arrow_end_width: None,
            runs: Vec::new(),
            h_align: -4131,
            v_align: -4160,
            margins: (7.2, 3.6, 7.2, 3.6),
            auto_size: false,
            adjusts: Vec::new(),
            placement: 1,
            on_action: String::new(),
            alt_text: String::new(),
            alt_text_utf16: None,
            alt_text_modified: false,
            lock_aspect: false,
            paragraph_count: 1,
            original: None,
            ole_clone_of: None,
            connector: 0,
            group_name: None,
            group_parent_id: None,
            group_children: Vec::new(),
            group_legacy: false,
            group_bounds: None,
            group_transform: None,
            group_nested: false,
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
            ShapeKind::Chart(chart) => {
                record.fill = colour_to_packed(chart.fill.as_deref()).unwrap_or(WHITE);
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
            color_theme: Some(if matches!(self.kind, ShapeKind::TextBox) { 1 } else { 2 }),
            color_brightness: 0.0,
            scheme_color: None,
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
        self.runs = if text.is_empty() {
            Vec::new()
        } else {
            let lines = text.split('\n').collect::<Vec<_>>();
            let mut runs = Vec::with_capacity(lines.len() * 2 - 1);
            for (index, line) in lines.iter().enumerate() {
                runs.push(ShapeRunRecord { text: (*line).to_string(), style: style.clone(), paragraph_end: false });
                if index + 1 < lines.len() {
                    runs.push(ShapeRunRecord { text: "\n".into(), style: style.clone(), paragraph_end: true });
                }
            }
            runs
        };
        if self.auto_size {
            self.fit_to_text();
        }
    }

    /// Replace a Characters span while keeping the dress of the untouched runs.
    pub(super) fn replace_characters(&mut self, start: u32, length: Option<u32>, text: &str) {
        // Excel treats an explicit zero length as the rest of the text here.
        let length = length.filter(|length| *length != 0);
        let (first, last) = self.split_runs(start, length);
        let style = if self.comment_at.is_some() && first > 0 {
            self.runs[first - 1].style.clone()
        } else {
            self.runs.get(first).or_else(|| self.runs.last()).map(|run| run.style.clone()).unwrap_or_else(|| self.default_style())
        };
        let text = text.replace("\r\n", "\r").replace('\r', "\n");
        let mut replacement = Vec::new();
        let mut lines = text.split('\n').peekable();
        while let Some(line) = lines.next() {
            if !line.is_empty() { replacement.push(ShapeRunRecord { text: line.to_string(), style: style.clone(), paragraph_end: false }); }
            if lines.peek().is_some() { replacement.push(ShapeRunRecord { text: "\n".into(), style: style.clone(), paragraph_end: true }); }
        }
        self.runs.splice(first..last, replacement);
        self.paragraph_count = self.text().split('\n').count().max(1);
        if self.auto_size { self.fit_to_text(); }
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
                    let paragraph_end = self.runs[at].paragraph_end;
                    self.runs[at].text = head;
                    self.runs[at].paragraph_end = false;
                    self.runs.insert(at + 1, ShapeRunRecord { text: tail, style, paragraph_end });
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
pub(super) fn generated_shape_bases(kind: &ShapeKind) -> Option<(&'static str, &'static str)> {
    let english = match kind {
        ShapeKind::Auto(kind) if (1..=108).contains(kind) => auto_shape(*kind).0,
        ShapeKind::TextBox => "TextBox",
        ShapeKind::Line => "Straight Connector",
        ShapeKind::Chart(_) => "Chart",
        _ => return None,
    };
    let localized = match kind {
        ShapeKind::Auto(kind) => match kind {
        1 => "正方形/長方形",
        2 => "平行四辺形",
        3 => "台形",
        4 => "ひし形",
        5 => "四角形: 角を丸くする",
        6 => "八角形",
        7 => "二等辺三角形",
        8 => "直角三角形",
        9 => "楕円",
        10 => "六角形",
        11 => "十字形",
        12 => "五角形",
        13 => "円柱",
        14 => "直方体",
        15 => "四角形: 角度付き",
        16 => "四角形: メモ",
        17 => "スマイル",
        18 => "円: 塗りつぶしなし",
        19 => "\"禁止\"マーク",
        20 => "アーチ",
        21 => "ハート",
        22 => "稲妻",
        23 => "太陽",
        24 => "月",
        25 => "円弧",
        26 => "大かっこ",
        27 => "中かっこ",
        28 => "ブローチ",
        29 => "左大かっこ",
        30 => "右大かっこ",
        31 => "左中かっこ",
        32 => "右中かっこ",
        33 => "矢印: 右",
        34 => "矢印: 左",
        35 => "矢印: 上",
        36 => "矢印: 下",
        37 => "矢印: 左右",
        38 => "矢印: 上下",
        39 => "矢印: 四方向",
        40 => "矢印: 三方向",
        41 => "矢印: 折線",
        42 => "矢印: U ターン",
        43 => "矢印: 二方向",
        44 => "矢印: 上向き折線",
        45 => "矢印: 右カーブ",
        46 => "矢印: 左カーブ",
        47 => "矢印: 上カーブ",
        48 => "矢印: 下カーブ",
        49 => "矢印: ストライプ",
        50 => "矢印: V 字型",
        51 => "矢印: 五方向",
        52 => "矢印: 山形",
        53 => "吹き出し: 右矢印",
        54 => "吹き出し: 左矢印",
        55 => "吹き出し: 上矢印",
        56 => "吹き出し: 下矢印",
        57 => "吹き出し: 左右矢印",
        58 => "吹き出し: 上下矢印",
        59 => "吹き出し: 四方向矢印",
        60 => "矢印: 環状",
        61 => "フローチャート: 処理",
        62 => "フローチャート: 代替処理",
        63 => "フローチャート: 判断",
        64 => "フローチャート: データ",
        65 => "フローチャート: 定義済み処理",
        66 => "フローチャート: 内部記憶",
        67 => "フローチャート: 書類",
        68 => "フローチャート: 複数書類",
        69 => "フローチャート: 端子",
        70 => "フローチャート: 準備",
        71 => "フローチャート: 手操作入力",
        72 => "フローチャート: 手作業",
        73 => "フローチャート: 結合子",
        74 => "フローチャート: 他ページ結合子",
        75 => "フローチャート: カード",
        76 => "フローチャート: せん孔テープ",
        77 => "フローチャート: 和接合",
        78 => "フローチャート: 論理和",
        79 => "フローチャート: 照合",
        80 => "フローチャート: 分類",
        81 => "フローチャート: 抜出し",
        82 => "フローチャート: 組合せ",
        83 => "フローチャート: 記憶データ",
        84 => "フローチャート: 論理積ゲート",
        85 => "フローチャート: 順次アクセス記憶",
        86 => "フローチャート: 磁気ディスク",
        87 => "フローチャート: 直接アクセス記憶",
        88 => "フローチャート: 表示",
        89 => "爆発: 8 pt",
        90 => "爆発: 14 pt",
        91 => "星: 4 pt",
        92 => "星: 5 pt",
        93 => "星: 8 pt",
        94 => "星: 16 pt",
        95 => "星: 24 pt",
        96 => "星: 32 pt",
        97 => "リボン: 上に曲がる",
        98 => "リボン: 下に曲がる",
        99 => "リボン: カーブして上方向に曲がる",
        100 => "リボン: カーブして下方向に曲がる",
        101 => "スクロール: 縦",
        102 => "スクロール: 横",
        103 => "波線",
        104 => "小波",
        105 => "吹き出し: 四角形",
        106 => "吹き出し: 角を丸めた四角形",
        107 => "吹き出し: 円形",
        108 => "思考の吹き出し: 雲形",
            _ => return None,
        },
        ShapeKind::TextBox => "テキスト ボックス",
        ShapeKind::Line => "直線コネクタ",
        ShapeKind::Chart(_) => "グラフ",
        _ => return None,
    };
    Some((english, localized))
}

pub(super) fn generated_group_name(number: u32, localized: bool) -> String {
    format!("{} {number}", if localized { "\u{30b0}\u{30eb}\u{30fc}\u{30d7}\u{5316}" } else { "Group" })
}

pub(super) fn generated_shape_name(kind: &ShapeKind, number: u32, localized: bool) -> Option<String> {
    let (english, native) = generated_shape_bases(kind)?;
    Some(format!("{} {number}", if localized { native } else { english }))
}

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
    for kind in 1..=108 {
        if let Some((english, native)) = generated_shape_bases(&ShapeKind::Auto(kind)) {
            if word == native { return Some(format!("{english} {number}")); }
        }
    }
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
