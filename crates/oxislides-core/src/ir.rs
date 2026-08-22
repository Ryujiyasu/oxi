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
    /// Fonts the deck carries inside itself (`p:embeddedFontLst`), one entry
    /// per face. A consumer that can install them renders the deck's real
    /// typography instead of a substitute.
    #[serde(default)]
    pub embedded_fonts: Vec<EmbeddedFont>,
}

/// One face of an embedded font (`p:embeddedFontLst/p:embeddedFont`).
///
/// The bytes are the `.fntdata` part verbatim, which is **EOT** (Embedded
/// OpenType), not a bare TTF: all 262 parts in the dev corpus are EOT 2.2 with
/// `TTEMBED_TTCOMPRESSED`, i.e. MicroType Express compressed, so renaming the
/// part to .ttf produces nothing loadable. A Windows consumer hands them to
/// `TTLoadEmbeddedFont` (t2embed.dll), which is the same route PowerPoint
/// takes; measured 2026-08-17 on d28's Calistoga, after which GDI resolves the
/// face by name (473x102 for "Abraham Lincoln" at 60px, against 417x60 for the
/// MS PGothic fallback it got before).
///
/// 37 of the 40 dev decks embed fonts (262 parts), and most of the faces are
/// Google Fonts present neither on the system nor in Office's cloud cache, so
/// without this every run of those decks is drawn in a substitute face.
#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct EmbeddedFont {
    /// `p:font/@typeface` — the name runs refer to (e.g. "Montserrat SemiBold").
    pub typeface: String,
    /// Which child carried this part: `p:bold` / `p:boldItalic` set bold,
    /// `p:italic` / `p:boldItalic` set italic, `p:regular` sets neither.
    #[serde(default)]
    pub bold: bool,
    #[serde(default)]
    pub italic: bool,
    /// The EOT bytes.
    pub data: Vec<u8>,
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
    /// `a:defRPr/a:latin@typeface` -- the face this outline level asks for.
    /// The MASTER's placeholder SHAPE carries it: d15's `body` placeholder
    /// declares Barlow Light while every other level of the chain (layout,
    /// master txStyles, theme minor, presentation default) says Arial, and
    /// PowerPoint follows the placeholder. 188 layout/master placeholders in
    /// the dev corpus name a font this way.
    #[serde(default)]
    pub font_family: Option<String>,
    /// `a:defRPr/@i` -- the level asks for italic. d16's layout body level
    /// declares `<a:defRPr i="1" sz="3600"/>` and PowerPoint sets the whole
    /// quotation in italic; 18 levels across two dev decks declare one.
    #[serde(default)]
    pub italic: bool,
    /// `a:defRPr/a:highlight` -- the box this outline level paints behind its
    /// text. d35's master TITLE placeholder declares a white one, which is the
    /// slab behind "BIG CONCEPT"; 19 levels in that deck carry one and no other
    /// dev deck does.
    #[serde(default)]
    pub highlight: Option<String>,
    /// `a:lnSpc/a:spcPct` as a multiple (90000 -> 0.9). PowerPoint render-truth
    /// (d24 slide 1, 2026-08-18): the MASTER's title PLACEHOLDER carries
    /// `lnSpc 90%` while the master's `p:txStyles/p:titleStyle` says 100%, and
    /// the rendered pitch is 64.82pt on 60pt text = 1.2 x 0.9. The placeholder
    /// style wins over txStyles.
    #[serde(default)]
    pub line_spacing: Option<f32>,
    /// `a:defRPr/a:solidFill` — the level's text colour, theme colours already
    /// resolved. PowerPoint render-truth (d24 slide 1, 2026-08-18): the deck's
    /// master titleStyle carries no size or colour at all, while the LAYOUT's
    /// ctrTitle placeholder `a:lstStyle` carries `sz="6000"` and `lt1`, and
    /// PowerPoint draws the title 60pt white.
    #[serde(default)]
    pub color: Option<String>,
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

/// A picture turns with its shape unless `a:blipFill/@rotWithShape="0"`.
pub fn default_rot_with_shape() -> bool {
    true
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct Slide {
    pub index: usize,
    pub shapes: Vec<Shape>,
    pub background_color: Option<String>, // hex color e.g. "FFFFFF"
    /// Slide background gradient (`p:bg/p:bgPr/a:gradFill`), inherited from the
    /// layout / master exactly like `background_color`. The two are mutually
    /// exclusive: a gradFill has no single colour, so `background_color` stays
    /// None for a gradient slide and a consumer that cannot paint a ramp leaves
    /// the page as it was before gradients were parsed at all.
    #[serde(default)]
    pub background_gradient: Option<SlideGradient>,
    /// Slide background picture (`p:bg/p:bgPr/a:blipFill`), inherited from the
    /// layout / master like the other two. Mutually exclusive with them for the
    /// same reason: a picture fill has no single colour and no ramp.
    #[serde(default)]
    pub background_image: Option<SlideBackgroundImage>,
}

/// A slide background picture fill.
///
/// PowerPoint render-truth (dev corpus, 2026-08): the exported PDF places the
/// image at exactly the page rect -- `Rect(0, 0, 720, 405)` on d04 / d06 / d16 /
/// d19 -- with no soft mask, i.e. it is STRETCHED to the full page and fully
/// opaque. All 22 background fills in the corpus are the same degenerate shape,
/// `<a:blipFill><a:blip r:embed=".."><a:alphaModFix/></a:blip><a:stretch>
/// <a:fillRect/></a:stretch></a:blipFill>`: no `a:tile`, no `a:srcRect`, no
/// `alphaModFix/@amt`, no `a:fillRect` insets, no duotone. Those variants are
/// therefore NOT modelled -- there is nothing measured to model them from.
#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct SlideBackgroundImage {
    /// Raw bytes of the media part (PNG or JPEG in the corpus: 37 / 40).
    pub data: Vec<u8>,
    /// Content type guessed from the target extension, as for picture shapes.
    #[serde(default)]
    pub content_type: Option<String>,
}

/// One `a:gs` of a gradient ramp.
#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct SlideGradientStop {
    /// `a:gs/@pos` normalized to 0.0..1.0.
    pub pos: f32,
    /// RRGGBB, theme colours already resolved.
    pub color: String,
    /// `<a:alpha>` inside the stop's colour, 0.0..1.0 (1.0 = opaque).
    #[serde(default = "default_stop_alpha")]
    pub alpha: f32,
}

pub fn default_stop_alpha() -> f32 {
    1.0
}

/// A slide background gradient. Exactly one of `angle_deg` / `focus` is set:
/// `a:lin` gives the angle, `a:path path="circle"` gives the focus.
#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct SlideGradient {
    /// Stops in document order (PowerPoint writes them ascending by `pos`).
    pub stops: Vec<SlideGradientStop>,
    /// `a:lin/@ang` in degrees clockwise from the +x axis: 0 = left->right,
    /// 90 = top->bottom, 270 = bottom->top (PDF render-truth, probe B1/B2/B4).
    #[serde(default)]
    pub angle_deg: Option<f32>,
    /// `a:lin/@scaled="1"` makes the ramp direction 45-degree in NORMALIZED
    /// space, i.e. the axis vector is proportional to (1/width, 1/height)
    /// rather than a physical angle (probe B6).
    #[serde(default)]
    pub scaled: bool,
    /// `a:path path="circle"` focus as a fraction of the page (0..1 of width,
    /// 0..1 of height), derived from `a:fillToRect`. The ramp runs from this
    /// point (t=0) out to the FARTHEST page corner (t=1) -- measured on d04
    /// (centre, r=413.05 = the corner distance) and d15 (bottom-right corner
    /// focus, r=826.09 = the distance to the opposite corner).
    #[serde(default)]
    pub focus: Option<(f32, f32)>,
}

fn default_true() -> bool {
    true
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct Shape {
    pub x: f32,      // position in points
    pub y: f32,
    pub width: f32,
    pub height: f32,
    pub rotation: f32, // rotation in degrees (0 = none), from a:xfrm/@rot (60000 = 1 deg)
    /// a:xfrm/@flipH / @flipV. Mirrors the shape about its centre. For a
    /// connector (which is only a segment) a flip just selects the other
    /// diagonal of the box, so only flip_h != flip_v matters.
    #[serde(default)]
    pub flip_h: bool,
    #[serde(default)]
    pub flip_v: bool,
    /// PresentationML shape type (a:prstGeom/@prst e.g. "rect", "ellipse", "roundRect", "chevron").
    /// None for picture / graphicFrame / plain textbox without a preset geometry.
    pub shape_type: Option<String>,
    /// Explicit preset-geometry adjustment values from
    /// `a:prstGeom/a:avLst/a:gd` (normally on the 0..100000 DrawingML scale).
    /// Preset defaults are intentionally not materialized here; renderers apply
    /// the default for the selected `shape_type` when an entry is absent.
    #[serde(default)]
    pub adjustments: std::collections::HashMap<String, f32>,
    /// Placeholder type (p:nvPr/p:ph/@type, "title" / "body" / "subTitle" / "obj"...).
    /// None for a plain (non-placeholder) shape. Title placeholders use the theme
    /// MAJOR font; everything else (incl. body placeholders) uses the MINOR font.
    #[serde(default)]
    pub ph_type: Option<String>,
    pub content: ShapeContent,
    pub fill_color: Option<String>,   // hex color for solid fill
    /// Opacity of `fill_color`, from `<a:alpha val="N"/>` inside the solidFill
    /// (`N` is a percentage in thousandths, so 62010 = 62.01% opaque). `None`
    /// means the fill is opaque, which is what an absent `a:alpha` means.
    ///
    /// PowerPoint composites this straight source-over on sRGB bytes: its PDF
    /// carries `/ca .50196 /BM /Normal` in a `/DeviceRGB` transparency group
    /// for `val="50000"`, and a 10-arm probe over white / red / green backdrops
    /// (incl. stacked translucent rects) matches `a*src + (1-a)*dst` to within
    /// 2/255 -- the residual is PowerPoint quantising alpha to 8 bits.
    #[serde(default)]
    pub fill_alpha: Option<f32>,
    pub border_color: Option<String>, // hex color for outline
    pub border_width: Option<f32>,    // border width in points
    /// `a:ln/a:prstDash/@val` — "dash", "dot", "lgDashDot" and friends. None
    /// (or "solid") is an unbroken line.
    #[serde(default)]
    pub border_dash: Option<String>,
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
    /// `a:bodyPr/@wrap` -- false for `wrap="none"`, where PowerPoint lets the
    /// text run past the box instead of breaking it. No dev-corpus shape asks
    /// for it; every arm of the COM-built `embedsplit` probes does.
    #[serde(default = "default_true")]
    pub wrap_text: bool,
    /// `a:bodyPr/a:prstTxWarp/@prst` — WordArt text fitting. On an AUTOSHAPE
    /// PowerPoint stretches the text's ink box onto the shape box exactly; on
    /// a plain text box it changes nothing.
    #[serde(default)]
    pub text_warp: Option<String>,
    /// `a:blipFill/@rotWithShape` — whether the picture inside a shape turns
    /// with the shape's `@rot`. True (the default) for a `p:pic`, whose raster
    /// always follows the shape.
    ///
    /// PowerPoint render-truth (img_rotation probe, 2026-08-17): with
    /// rotWithShape="1" a 90-degree shape maps the source's bottom-left corner
    /// to the box's top-left (E4, identical to a rotated `p:pic` in E2); with
    /// "0" the raster stays upright while the shape still turns (E5). All 2141
    /// shape-level blipFills in the dev corpus declare "1", so the schema
    /// default is not exercised and is taken to be "rotate".
    #[serde(default = "default_rot_with_shape")]
    pub rot_with_shape: bool,
    /// `a:blipFill/a:blip/a:alphaModFix/@amt` — the picture's opacity, in
    /// percent-thousandths (7000 = 7%). None = the bare `<a:alphaModFix/>` the
    /// corpus writes 3087 times, which carries no attribute and means opaque.
    ///
    /// 35 shapes on 22 slides in 5 decks (d12/d15/d30/d32/d38) declare a real
    /// `amt`, from 1% to 80%. d32's title slide is a city map at **7%** -- a
    /// dark texture in PowerPoint, a stark white overlay when the attribute is
    /// ignored.
    #[serde(default)]
    pub image_alpha: Option<f32>,
    /// The text styles of the LAYOUT placeholder this shape inherits from
    /// (`a:lstStyle` on the layout's matching `p:sp`), indexed by outline
    /// level. This sits BETWEEN the run's own properties and the master
    /// txStyles: a run with no explicit `sz` takes the layout placeholder's,
    /// and only then the master's.
    ///
    /// It is NOT the layout's `p:txStyles`, which the phfs probe showed
    /// PowerPoint ignores -- a different element with a similar name.
    #[serde(default)]
    pub ph_levels: Vec<MasterStyleLevel>,
    /// `a:gradFill` on the shape itself. The dev corpus has 302 of these on
    /// 35 slides in 4 decks, plus 60 more on layout shapes -- d24's title
    /// slide is built entirely out of them, which is why it renders as a flat
    /// slab. The ramp model is the one already derived for slide backgrounds.
    #[serde(default)]
    pub gradient: Option<SlideGradient>,
    /// Custom geometry (`a:custGeom`) — the shape's outline as explicit paths
    /// instead of a named preset. None for a preset / picture / frame shape.
    ///
    /// Corpus census (dev 40 decks / 886 slides, 2026-08-17): 11470 custGeom
    /// shapes on 628 slides in 40/40 decks, i.e. every deck. The command
    /// vocabulary they actually use is exactly four elements -- lnTo 341943 /
    /// cubicBezTo 77615 / moveTo 21382 / close 20892 -- with `arcTo` and
    /// `quadBezTo` appearing ZERO times, so only those four are modelled and a
    /// path using anything else is refused whole (see `unsupported`).
    #[serde(default)]
    pub custom_geometry: Option<CustomGeometry>,
    /// Image source crop (a:blipFill/a:srcRect l,t,r,b), normalized to 0..1.
    /// Only meaningful for ShapeContent::Image. None = full source image.
    /// Word render-truth (01__Biology deck, 2026-08): a full-bleed background
    /// PNG crops the SOURCE (srcRect t=21.875% b=21.874%) so the cropped
    /// aspect matches the destination — keeping the image un-distorted.
    #[serde(default)]
    pub src_rect: Option<(f32, f32, f32, f32)>,
    /// Image destination insets (a:stretch/a:fillRect l,t,r,b), normalized to
    /// a fraction of the shape box (negative = expand beyond the box). Only
    /// meaningful for ShapeContent::Image. None = fill the shape box exactly.
    /// Word render-truth (01__Biology deck photo): fillRect t=b=-22.646%
    /// EXPANDS the destination vertically so the stretched source keeps its
    /// native aspect (portrait 977x1350 into a 297x282pt box).
    #[serde(default)]
    pub fill_rect: Option<(f32, f32, f32, f32)>,
}

/// `a:custGeom` — one or more closed/open outlines in their own coordinate
/// space. Every path declares the space it is drawn in (`w` x `h`), which the
/// consumer maps onto the shape box; the corpus always declares both.
#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct CustomGeometry {
    pub paths: Vec<GeomPath>,
    /// A path command outside the modelled vocabulary (`a:arcTo`,
    /// `a:quadBezTo`) was seen. The geometry is then incomplete, so a consumer
    /// must NOT draw it -- an outline missing one of its curves is worse ink
    /// than the rectangle it replaces.
    #[serde(default)]
    pub unsupported: bool,
}

/// One `a:path`.
#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct GeomPath {
    /// `a:path/@w` / `@h`, the local coordinate space. 0 = the path is already
    /// in the shape's own coordinates (ECMA-376 20.1.9.15).
    pub w: f32,
    pub h: f32,
    /// `a:path/@fill="none"` — stroke this subpath but do not fill it (450 of
    /// the corpus's 11470 geometries carry an explicit `fill`).
    #[serde(default)]
    pub fill_none: bool,
    pub commands: Vec<GeomCmd>,
}

/// A path command, in the path's local coordinate space.
#[derive(Debug, Clone, Serialize, Deserialize)]
pub enum GeomCmd {
    MoveTo(f32, f32),
    LineTo(f32, f32),
    /// Two control points followed by the end point.
    CubicTo(f32, f32, f32, f32, f32, f32),
    Close,
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
    /// c:doughnutChart/c:holeSize@val — the hole diameter as a percent of the
    /// outer diameter. Word render-truth (chart_doughnut 2026-08-09, 600dpi
    /// pixel scan of the ring): r_in / r_out == holeSize/100 exactly —
    /// 0.5010/0.5011/0.5013/0.5014 at holeSize 50 and 0.2510 at holeSize 25.
    #[serde(default = "default_chart_hole_size")]
    pub hole_size: f64,
    /// c:bubbleChart/c:bubbleScale@val — the bubble size as a PERCENT
    /// (default 100). Word render-truth (chart_bubble_size probe, 2026-08-10):
    /// the largest bubble's diameter saturates toward the available box,
    /// `d_max = avail * 3*scale / (3*scale + 1000)` — exact at 50 / 100 / 200
    /// / 300 (15.82 / 28.00 / 45.49 / 57.47 pt radius). At 100 that is the
    /// clean 3/13 of `avail`.
    #[serde(default = "default_chart_bubble_scale")]
    pub bubble_scale: f64,
    /// c:bubbleChart/c:sizeRepresents@val — "area" (default) or "w". Measured
    /// on the same probe: "w" gives radii in the ratio of the raw sizes
    /// (1:2:4), the default gives their square roots (1:1.414:2).
    #[serde(default = "default_chart_size_represents")]
    pub size_represents: String,
    /// c:stockChart/c:hiLowLines — when present Word draws one vertical
    /// rule per category spanning the MIN..MAX of every series at that
    /// category. Word render-truth (chart_stock K1, 2026-08-10): Q1
    /// High 24.0 / Low 18.2 renders 134.40..179.28 under the plain value
    /// mapping, black w=0.75, drawn UNDER the up/down bars.
    #[serde(default)]
    pub hi_low_lines: bool,
    /// c:stockChart/c:upDownBars — when present Word draws a box between
    /// the FIRST and LAST series' values (open..close). Measured on
    /// chart_stock K6/K7: white #F9F9F9 when the last value is above the
    /// first, dark #3F3F3F when below, both with a black w=0.75 outline.
    #[serde(default)]
    pub up_down_bars: bool,
    /// c:upDownBars/c:gapWidth@val — the gap between adjacent bars as a
    /// percent of the bar width. Word render-truth: the drawn width is
    /// `pitch / (1 + gapWidth/100)` — 34.32 at pitch 85.89 and 27.20 at
    /// pitch 68.00, both exact at the default 150.
    #[serde(default = "default_chart_updown_gap")]
    pub up_down_gap: f64,
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
    /// False only when the legend declares `<c:overlay val="0"/>`: the legend
    /// is NOT overlaid, so it takes a band on the right and the plot area
    /// shrinks into what remains. A bare `<c:legend/>` (python-pptx's default,
    /// no overlay child) is an OVERLAY — chart_pie2 p2/p3 keep the circle on
    /// the frame centre and the legend swatches (x0 379.55) sit ON TOP of the
    /// circle (right edge 402.84 / 388.23), while chart_doughnut (overlay=0)
    /// shifts the ring left. Measured 2026-08-09.
    #[serde(default = "default_chart_legend_overlay")]
    pub legend_overlay: bool,
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
    /// Text of an EXPLICIT `<c:title>` (c:title/c:tx/c:rich/a:p/a:r/a:t).
    /// python-pptx writes it when chart.has_title=True + chart_title text is
    /// set. When present Word draws THIS text as the chart title (Arial 18pt,
    /// regular, centred on the frame, baseline sy+24.43) and does NOT draw
    /// the automatic series-name title — chart_title / chart_title2
    /// render-truth 2026-08-07.
    #[serde(default)]
    pub explicit_title: Option<String>,
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
    "bar".to_string()
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
/// c:doughnutChart/c:holeSize@val — the hole diameter as a PERCENT of the
/// outer diameter. OOXML's implied default is 10; python-pptx writes 50.
pub fn default_chart_hole_size() -> f64 {
    50.0
}
/// `<c:bubbleScale>` defaults to 100 percent.
pub fn default_chart_bubble_scale() -> f64 {
    100.0
}
/// `<c:sizeRepresents>` defaults to "area" (the value is the bubble's AREA,
/// so the radius goes as its square root).
pub fn default_chart_size_represents() -> String {
    "area".to_string()
}
/// ECMA-376 default for `c:upDownBars/c:gapWidth`.
pub fn default_chart_updown_gap() -> f64 {
    150.0
}
/// A legend with no `<c:overlay val="0"/>` child is an overlay.
pub fn default_chart_legend_overlay() -> bool {
    true
}

/// One c:ser element: a named series of values (per category).
#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct ChartSeries {
    /// c:tx -> strCache "Series 1" (the legend entry).
    pub name: String,
    /// c:val -> numCache, one value per category.
    pub values: Vec<f64>,
    /// c:xVal -> numCache. Scatter (XY) series carry a NUMERIC x for every
    /// point instead of sharing the chart's categories; `values` then holds
    /// c:yVal. Empty for every category-based chart type.
    #[serde(default)]
    pub x_values: Vec<f64>,
    /// This series declares `<c:spPr><a:ln><a:noFill/>` — draw no connecting
    /// line. python-pptx writes it for XY_SCATTER (markers only).
    #[serde(default)]
    pub line_none: bool,
    /// This series declares `<c:marker><c:symbol val="none"/>` — draw no
    /// markers. python-pptx writes it for XY_SCATTER_LINES_NO_MARKERS.
    #[serde(default)]
    pub marker_none: bool,
    /// c:bubbleSize -> numCache. A bubble series carries a third number per
    /// point (x from c:xVal, y from c:yVal, size here). Empty for every other
    /// chart type.
    #[serde(default)]
    pub sizes: Vec<f64>,
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

/// One side of a cell's border (`a:tcPr/a:lnL|lnR|lnT|lnB`).
///
/// The corpus writes these explicitly on every cell, including the ones that
/// must NOT be drawn: an invisible edge is a `solidFill` whose colour carries
/// `<a:alpha val="0"/>`, not an absent element. A consumer that ignores the
/// alpha draws a full grid where PowerPoint draws a few rules.
#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct CellBorder {
    /// RRGGBB, theme colours already resolved.
    pub color: String,
    /// `@w` in points (9525 EMU = 0.75pt is what the corpus writes).
    pub width: f32,
    /// 0.0 = fully transparent (do not draw), 1.0 = opaque.
    #[serde(default = "default_border_alpha")]
    pub alpha: f32,
}

pub fn default_border_alpha() -> f32 {
    1.0
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct TableCell {
    pub paragraphs: Vec<SlideParagraph>,
    /// `a:tcPr/a:solidFill` — the cell's own fill (the corpus states the
    /// header and banded-row colours here rather than leaving them to the
    /// table style).
    #[serde(default)]
    pub fill_color: Option<String>,
    /// `<a:alpha>` inside that fill. The corpus leans on it hard: every cell of
    /// d19 slide 13 is `21355A` at **15.6%**, which is a pale wash over the
    /// page and a solid navy slab if the alpha is dropped.
    #[serde(default)]
    pub fill_alpha: Option<f32>,
    /// Left / right / top / bottom, in that order.
    #[serde(default)]
    pub borders: [Option<CellBorder>; 4],
    /// `a:tcPr/@marL|marR|marT|marB` in points (default 0.1 inch / 0.05 inch,
    /// i.e. 7.2 / 3.6pt, the same defaults a text body uses).
    #[serde(default = "default_l_ins")]
    pub mar_l: f32,
    #[serde(default = "default_r_ins")]
    pub mar_r: f32,
    #[serde(default = "default_t_ins")]
    pub mar_t: f32,
    #[serde(default = "default_b_ins")]
    pub mar_b: f32,
    /// `a:tcPr/@anchor` — vertical text anchor, as on a shape body.
    #[serde(default)]
    pub anchor: Option<String>,
    /// `a:tc/@gridSpan` — how many grid columns this cell occupies. The cells
    /// it swallows still exist in the row as `h_merge` continuations.
    #[serde(default = "default_grid_span")]
    pub grid_span: u32,
    /// `a:tc/@hMerge` — this cell is the continuation of the spanning cell to
    /// its left and is not drawn in its own right.
    #[serde(default)]
    pub h_merge: bool,
}

pub fn default_grid_span() -> u32 {
    1
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
    /// `a:endParaRPr/@sz` in points -- the size of the paragraph mark. It is
    /// what gives an EMPTY paragraph its line height: PowerPoint advances such
    /// a line by sz * 1.2 * lnSpc exactly (probe emptypara arms A-D: 7/10/24/40
    /// pt all land on the nose), and it wins over an rPr on a textless run
    /// (arm F: run 10pt + endParaRPr 40pt renders 40pt). Ignored when the
    /// paragraph has text -- then the runs govern.
    #[serde(default)]
    pub end_para_size: Option<f32>,
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
    /// `a:rPr/@u` -- any value but `none` underlines the run. PowerPoint draws
    /// the rule; 86 runs across 12 dev decks ask for one, most of them the
    /// hyperlinks a template puts in its instructions.
    #[serde(default)]
    pub underline: bool,
    pub color: Option<String>,     // hex color
    /// `<a:alpha>` inside the run's own `a:solidFill`, 0.0..1.0. d35's
    /// transition numerals are white at 26.9%, so the gradient behind shows
    /// through; drawn opaque they read as a solid slab.
    #[serde(default)]
    pub color_alpha: Option<f32>,
    /// `a:rPr/a:highlight` -- a filled box behind the run's glyphs, as hex.
    /// PowerPoint draws it the height of the LINE's font (hhea ascent plus
    /// descent) with its bottom on the line box's bottom, and as wide as the
    /// run's own advance including a trailing space (`highlight` probe,
    /// 2026-08-19). 65 runs across 19 slides in 8 of the 40 dev decks have one.
    #[serde(default)]
    pub highlight: Option<String>,
    pub font_family: Option<String>,
}
