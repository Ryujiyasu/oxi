// This Source Code Form is subject to the terms of the Mozilla Public
// License, v. 2.0. If a copy of the MPL was not distributed with this
// file, You can obtain one at https://mozilla.org/MPL/2.0/.

use std::collections::HashMap;

use serde::{Deserialize, Serialize};

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct Presentation {
    pub slides: Vec<Slide>,
    pub slide_width: f32,  // in points (default 960pt = 10 inches)
    pub slide_height: f32, // in points (default 540pt = 7.5 inches)
    /// Theme minor-font latin typeface (default font for body text / most shapes).
    /// From ppt/theme/themeN.xml <a:minorFont><a:latin typeface="..."/>. Default "Calibri".
    #[serde(default = "default_theme_font")]
    pub minor_font: String,
    /// Theme major-font latin typeface (title placeholders). Default "Calibri".
    #[serde(default = "default_theme_font")]
    pub major_font: String,
    /// Theme colour scheme (a:clrScheme from ppt/theme/themeN.xml): scheme slot
    /// name (dk1/dk2/lt1/lt2/tx1/tx2/accent1..accent6/hlink/folHlink) -> RGB hex
    /// (6-digit RRGGBB). `<a:schemeClr val="..."/>` references resolve through
    /// this map first; the built-in Office table is only the fallback when the
    /// theme part is absent (Spec #10).
    #[serde(default)]
    pub theme_colors: HashMap<String, String>,
    /// Slide master text styles (p:txStyles from slideMaster1.xml): the
    /// inherited marL/indent/bullet/spcBef per outline level for body / other
    /// (textbox) / title contexts (Spec #8). Placeholder body text uses
    /// `body`, plain textboxes use `other`, title placeholders use `title`.
    #[serde(default)]
    pub master_styles: MasterTxStyles,
}

/// Slide master text styles (p:txStyles). Each Vec is indexed by outline level
/// (0-based, a:lvlNpPr where N = level+1).
#[derive(Debug, Clone, Default, Serialize, Deserialize)]
pub struct MasterTxStyles {
    #[serde(default)]
    pub body: Vec<MasterStyleLevel>,
    #[serde(default)]
    pub other: Vec<MasterStyleLevel>,
    #[serde(default)]
    pub title: Vec<MasterStyleLevel>,
}

/// One outline-level master text style (a:lvlNpPr).
#[derive(Debug, Clone, Default, Serialize, Deserialize)]
pub struct MasterStyleLevel {
    /// a:lvlNpPr/@marL in points (left indent of the paragraph; default 0).
    #[serde(default)]
    pub mar_l: f32,
    /// a:lvlNpPr/@indent in points (first-line indent; negative = hanging).
    #[serde(default)]
    pub indent: f32,
    /// Inherited bullet marker (default Inherit = none unless a bu* child).
    #[serde(default)]
    pub bullet: SlideBullet,
    /// a:spcBef/a:spcPct val/100000 — space-before as a fraction of the line
    /// advance (e.g. 0.2 = 20000). None = no spcBef on this level.
    #[serde(default)]
    pub spc_bef_pct: Option<f32>,
    /// a:defRPr/@sz in points — the placeholder default font size for this
    /// outline level. None = no explicit size (engine default 18pt applies).
    /// Word render-truth (phfs probe, 2026-08): a body placeholder inherits
    /// the MASTER txStyles bodyStyle level size (layout txStyles is ignored);
    /// a title placeholder inherits master titleStyle. A run's explicit sz
    /// always wins.
    #[serde(default)]
    pub font_size: Option<f32>,
    /// a:lvlNpPr/@algn — the horizontal paragraph alignment inherited from the
    /// master txStyles level (Spec #6). None = not specified (paragraph-level
    /// Left applies). The master titleStyle lvl1pPr carries algn="ctr", which
    /// is what horizontally centres title placeholders (V4/P2 render-truth).
    #[serde(default)]
    pub algn: Option<SlideAlignment>,
}

pub fn default_l_ins() -> f32 {
    7.2
}
pub fn default_r_ins() -> f32 {
    7.2
}
pub fn default_t_ins() -> f32 {
    3.6
}
pub fn default_b_ins() -> f32 {
    3.6
}

pub fn default_theme_font() -> String {
    "Calibri".to_string()
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct Slide {
    pub index: usize,
    pub shapes: Vec<Shape>,
    pub background_color: Option<String>, // hex color e.g. "FFFFFF"
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct Shape {
    pub x: f32,      // position in points
    pub y: f32,
    pub width: f32,
    pub height: f32,
    pub rotation: f32, // rotation in degrees (0 = none), from a:xfrm/@rot (60000 = 1 deg)
    /// PresentationML shape type (a:prstGeom/@prst e.g. "rect", "ellipse", "roundRect", "chevron").
    /// None for picture / graphicFrame / plain textbox without a preset geometry.
    pub shape_type: Option<String>,
    /// Placeholder type (p:nvPr/p:ph/@type, "title" / "body" / "subTitle" / "obj"...).
    /// None for a plain (non-placeholder) shape. Title placeholders use the theme
    /// MAJOR font; everything else (incl. body placeholders) uses the MINOR font.
    #[serde(default)]
    pub ph_type: Option<String>,
    pub content: ShapeContent,
    pub fill_color: Option<String>,   // hex color for solid fill
    pub border_color: Option<String>, // hex color for outline
    pub border_width: Option<f32>,    // border width in points
    /// Text-area insets from a:bodyPr (lIns/rIns/tIns/bIns) in points.
    /// A placeholder with no bodyPr inset uses 7.2 / 7.2 / 3.6 / 3.6; a textbox
    /// carries its own insets (e.g. lIns=914400 EMU = 72pt). The left text
    /// position P0 = shape.x + l_ins (Spec #8).
    #[serde(default = "default_l_ins")]
    pub l_ins: f32,
    #[serde(default = "default_r_ins")]
    pub r_ins: f32,
    #[serde(default = "default_t_ins")]
    pub t_ins: f32,
    #[serde(default = "default_b_ins")]
    pub b_ins: f32,
    /// Vertical text-anchor (a:bodyPr/@anchor), resolved through the
    /// placeholder chain (slide shape's own bodyPr -> layout placeholder ->
    /// master placeholder). None = "t" (top). Spec #6 render-truth: the master
    /// TITLE placeholder carries anchor="ctr", which vertically centres title
    /// placeholders; layout3 title uses "t", layout8/9 "b".
    #[serde(default)]
    pub anchor: Option<String>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub enum ShapeContent {
    /// A preset-geometry AutoShape (a:prstGeom). May carry text (a:txBody).
    AutoShape {
        paragraphs: Vec<SlideParagraph>,
    },
    TextBox {
        paragraphs: Vec<SlideParagraph>,
    },
    /// A DrawingML table (a:graphicFrame -> a:graphicData uri=.../table -> a:tbl).
    Table {
        table: Table,
    },
    Image {
        data: Vec<u8>,
        content_type: Option<String>,
    },
    /// A DrawingML chart (a:graphicFrame -> a:graphicData uri=.../chart ->
    /// c:chart). Data is pulled from the embedded chart part (strCache /
    /// numCache), since the external xlsx workbook is not read.
    Chart {
        chart: Chart,
    },
    /// Unsupported element with type label (e.g. "SmartArt", "Chart", "OLE")
    Unsupported {
        element_type: String,
    },
    Placeholder, // shapes we can't parse yet
}

/// A DrawingML chart (c:chartSpace/c:chart/c:plotArea/<chartType>).
///
/// Word render-truth (chart1 repro, 2026-08, fitz get_drawings — the
/// decisive measurement for vector charts; text spans alone cannot see
/// bars/axes):
///   - shape frame = the a:xfrm of the graphicFrame (72,72,396,288pt).
///   - plot area   = (113.4,123.4,457,280.8) = frame insets for the value
///     axis labels (left), category names (bottom) and legend (top).
///     Word derives these from the axis text extent; for a clustered column
///     with 3 categories the measured inset is ~(41.4,51.4,0,40.7).
///   - bar width   = 40% of the category pitch (= plot_width / n_categories);
///     bar height  = value / max_value * plot_height (plot_height =
///     plot_bottom - plot_top); bar bottom = the X axis y.
///   - series colour = theme accent(i+1) for series i (accent1 #4F81BD,
///     accent2 #C0504D, accent3 #9BBB59 in the default Office theme).
///   - value axis = 0..max, evenly spaced labels (5 ticks for max 25 →
///     pitch = plot_height/5), Calibri 18pt right-aligned to plot_left.
///   - category names = Calibri 18pt centred under each bar (y = plot_bottom
///     + text height).
///   - legend (when `<c:legend>` is declared) = per-series swatch + series
///     name, RIGHT-aligned overlay: the swatch column sits at
///     legend_right - max_label_w - gap (legend_right = frame right - 10pt),
///     the block is vertically centred on the frame, and the plot area is
///     NOT shrunk (COM Legend.IncludeInLayout=False; chart_legend/chart_legend3
///     render-truth 2026-08-06).
#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct Chart {
    /// The chart type (c:pieChart / c:barChart / ... element name). "col" is
    /// the legacy default used when no chart-type element was matched (the
    /// pre-pie parser only handled c:barChart). "pie" selects the pie
    /// rendering path.
    #[serde(default = "default_chart_type")]
    pub chart_type: String,
    /// c:barChart/c:barDir@val — "col" (column/vertical) or "bar" (horizontal).
    #[serde(default = "default_chart_bar_dir")]
    pub bar_dir: String,
    /// c:barChart/c:grouping@val — "clustered", "stacked", "percentStacked".
    #[serde(default = "default_chart_grouping")]
    pub grouping: String,
    /// Series in document order (c:ser/c:idx). Series i renders with theme
    /// accent(i+1).
    pub series: Vec<ChartSeries>,
    /// Category labels from the first series' c:cat strCache (c:tx may be
    /// shared; category count == values count for a rectangular chart).
    pub categories: Vec<String>,
    /// True when the chart XML declares `<c:legend/>`. When true Word draws
    /// a legend (series names + accent-colour swatches) as an overlay on the
    /// right side of the plot area; when false (no <c:legend>) it is not
    /// drawn at all. None of the 4 chart1-4 probes declare one.
    #[serde(default)]
    pub has_legend: bool,
    /// True when the chart XML declares `<c:autoTitleDeleted val="1"/>`
    /// (self-closing). Word derives the chart's plot geometry AND whether the
    /// automatic series-name title is drawn from this flag: a pie with
    /// autoTitleDeleted=0 is "titled" — the series name is drawn at the
    /// frame top and the circle is shifted down; autoTitleDeleted=1 draws no
    /// title and the circle sits higher. Measured on chart_pie (0 → titled),
    /// chart_pie3 A/C/E (1 → untitled), chart_pie3 B/D/F (0 + <c:title> →
    /// titled) — the title element presence and this flag always agree for
    /// auto-titles, so this flag alone is the discriminator.
    #[serde(default)]
    pub auto_title_deleted: bool,
    /// True when a line chart declares `<c:marker val="1"/>` (self-closing
    /// child of <c:lineChart>) — LINE_MARKERS. When true Word draws a 6.96pt
    /// filled accent-colour circle at every data point (chart_line probe,
    /// all P0-P6 measured marker=1). A plain LINE chart (val="0" or absent)
    /// draws no markers — unmeasured, but the flag is read faithfully so the
    /// renderer can gate on it.
    #[serde(default)]
    pub marker: bool,
    /// True when the chart XML declares `<c:dLbls>` (data labels). When true
    /// Word draws a value label per data point, formatted by `number_format`
    /// and placed according to `datalabel_position` (chart_datalabel probe,
    /// measured on the Word PDF 2026-08-06). Absent <c:dLbls> draws none.
    #[serde(default)]
    pub has_data_labels: bool,
    /// c:dLbls/c:dLblPos@val — "outEnd" (default, above the bar),
    /// "ctr" (centre of the bar), "inEnd" (inside the bar top). Maps to the
    /// COM XL_LABEL_POSITION constants 2 / -4108 / 3. Measured on
    /// chart_datalabel S1/S3/S4 (outEnd/ctr/inEnd).
    #[serde(default = "default_datalabel_position")]
    pub datalabel_position: String,
    /// c:dLbls/c:numFmt@formatCode — e.g. "0.0%". Empty means General
    /// (a plain value). The renderer applies the format to the raw value
    /// (e.g. "0.0%" → value*100 with one decimal + "%", measured on
    /// chart_datalabel S2: 19.2 → "1920.0%").
    #[serde(default)]
    pub number_format: String,
    /// c:dLbls/c:showVal@val — draw the data value. (chart_datalabel S1-S5
    /// all declare showVal=1.)
    #[serde(default)]
    pub show_val: bool,
    /// c:dLbls/c:showCatName@val — draw the category name.
    #[serde(default)]
    pub show_cat_name: bool,
    /// c:dLbls/c:showSerName@val — draw the series name.
    #[serde(default)]
    pub show_ser_name: bool,
    /// c:dLbls/c:showPercent@val — draw the percentage (pie).
    #[serde(default)]
    pub show_percent: bool,
    /// c:dLbls/c:showLegendKey@val — draw the legend key swatch.
    #[serde(default)]
    pub show_legend_key: bool,
    /// c:dLbls/c:showBubbleSize@val — draw the bubble size.
    #[serde(default)]
    pub show_bubble_size: bool,
}

pub fn default_chart_type() -> String {
    "col".to_string()
}
pub fn default_chart_bar_dir() -> String {
    "col".to_string()
}
pub fn default_chart_grouping() -> String {
    "clustered".to_string()
}
pub fn default_datalabel_position() -> String {
    "outEnd".to_string()
}

/// One c:ser element: a named series of values (per category).
#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct ChartSeries {
    /// c:tx -> strCache "Series 1" (the legend entry).
    pub name: String,
    /// c:val -> numCache, one value per category.
    pub values: Vec<f64>,
}

/// A DrawingML table (a:tbl). Cell text is stored as paragraphs per cell so
/// the run/paragraph machinery is shared with shapes.
#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct Table {
    /// Column widths in points (a:tblGrid/a:gridCol/@w).
    pub col_widths: Vec<f32>,
    /// Row heights in points (a:tr/@h).
    pub row_heights: Vec<f32>,
    /// Cells in row-major order: rows[r][c].
    pub rows: Vec<Vec<TableCell>>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct TableCell {
    pub paragraphs: Vec<SlideParagraph>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct SlideParagraph {
    pub runs: Vec<SlideRun>,
    /// Horizontal alignment (a:pPr/@algn). None = not specified on the
    /// paragraph — the master txStyles level alignment (MasterStyleLevel.algn)
    /// applies at render time (Spec #6).
    #[serde(default)]
    pub alignment: Option<SlideAlignment>,
    /// Multiple line-spacing factor `n` from `a:lnSpc/a:spcPct` (val/100000).
    /// None = no explicit lnSpc -> PowerPoint uses its default single line
    /// (measured line advance = fs*1.2).  When Some(n), the measured line
    /// advance = fs*1.2*n (linear over n in [0.5, 3.0], verified spec4d).
    #[serde(default)]
    pub line_spacing: Option<f32>,
    /// Space before this paragraph, in points (`a:spcBef/a:spcPts` val/100).
    /// Added on top of the line advance (wave-1 measurement).
    #[serde(default)]
    pub space_before: Option<f32>,
    /// Space after this paragraph, in points (`a:spcAft/a:spcPts` val/100).
    #[serde(default)]
    pub space_after: Option<f32>,
    /// Outline level (a:pPr/@lvl, 0-based). Default 0.
    #[serde(default)]
    pub lvl: u32,
    /// Left indent (a:pPr/@marL) in points. None = not specified (the master
    /// txStyles level provides it).
    #[serde(default)]
    pub mar_l: Option<f32>,
    /// First-line indent (a:pPr/@indent) in points. None = not specified.
    #[serde(default)]
    pub indent: Option<f32>,
    /// Bullet marker spec (a:buChar / a:buNone / a:buAutoNum). Inherit = use
    /// the master txStyles level's bullet.
    #[serde(default)]
    pub bullet: SlideBullet,
}

/// Bullet marker specification for a paragraph (Spec #8, measured model).
#[derive(Debug, Clone, Default, Serialize, Deserialize)]
pub enum SlideBullet {
    /// No a:bu* child on the paragraph itself: inherit from the style chain
    /// (the master txStyles level for the shape's context).
    #[default]
    Inherit,
    /// a:buNone — no bullet marker is drawn (but indent geometry still applies).
    None,
    /// a:buChar — a literal character marker (e.g. "•", "–", "»").
    Char {
        ch: char,
        font: Option<String>, // a:buFont/@typeface
    },
    /// a:buAutoNum — an automatically numbered marker (Spec #11, derived
    /// 2026-08-06: kind-specific formats, per-(lvl, kind, startAt) counters,
    /// list split when startAt is present / changes).
    AutoNum {
        kind: String,
        /// a:buAutoNum/@startAt — the first marker value (a fresh list starts
        /// at this value; None = 1). A paragraph whose startAt differs from
        /// the previous paragraph's starts a new list.
        #[serde(default)]
        start_at: Option<u32>,
    },
}

#[derive(Debug, Clone, Copy, Default, Serialize, Deserialize)]
pub enum SlideAlignment {
    #[default]
    Left,
    Center,
    Right,
    Justify,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct SlideRun {
    pub text: String,
    pub font_size: Option<f32>,    // in points
    pub bold: bool,
    pub italic: bool,
    pub color: Option<String>,     // hex color
    pub font_family: Option<String>,
}
