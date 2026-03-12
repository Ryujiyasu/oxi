use serde::Serialize;
use std::collections::HashMap;

/// A parsed PDF document.
#[derive(Debug, Clone, Serialize)]
pub struct PdfDocument {
    /// PDF version (e.g. 1.7, 2.0).
    pub version: PdfVersion,
    /// Top-level document metadata from the Info dictionary.
    pub info: DocumentInfo,
    /// Pages in document order.
    pub pages: Vec<Page>,
    /// Named destinations, bookmarks, etc.
    pub outline: Vec<OutlineItem>,
    /// Embedded font data, keyed by font name (as used in TextSpan.font_name).
    /// When present, the font binary will be embedded in the PDF for portable rendering.
    #[serde(skip)]
    pub embedded_fonts: HashMap<String, EmbeddedFont>,
}

/// Embedded font data for PDF inclusion.
#[derive(Debug, Clone)]
pub struct EmbeddedFont {
    /// Raw font file data (CFF program, or TTF binary).
    pub data: Vec<u8>,
    /// Font format for correct /FontFile* reference.
    pub format: FontFormat,
    /// Unicode codepoint → Glyph ID mapping (from cmap table).
    /// Used by the writer to encode text as CID values matching the font's glyph indices.
    pub unicode_to_gid: HashMap<u32, u16>,
    /// CID → width in 1/1000 em units. Used for /W array in CIDFont dictionary.
    /// If empty, /DW 1000 is used for all glyphs (full-width).
    pub cid_widths: HashMap<u16, u16>,
    /// PostScript name for /BaseFont in PDF (e.g. "MS-Gothic").
    /// If None, the key name is used as-is.
    pub ps_name: Option<String>,
}

/// Font binary format.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum FontFormat {
    /// TrueType (.ttf) — embedded via /FontFile2
    TrueType,
    /// OpenType CFF (.otf) — embedded via /FontFile3 /Subtype /CIDFontType0C
    OpenTypeCff,
}

#[derive(Debug, Clone, Copy, PartialEq, Serialize)]
pub struct PdfVersion {
    pub major: u8,
    pub minor: u8,
}

impl PdfVersion {
    pub fn new(major: u8, minor: u8) -> Self {
        Self { major, minor }
    }
}

impl std::fmt::Display for PdfVersion {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        write!(f, "{}.{}", self.major, self.minor)
    }
}

/// Document-level metadata (from the Info dictionary).
#[derive(Debug, Clone, Default, Serialize)]
pub struct DocumentInfo {
    pub title: Option<String>,
    pub author: Option<String>,
    pub subject: Option<String>,
    pub keywords: Option<String>,
    pub creator: Option<String>,
    pub producer: Option<String>,
    pub creation_date: Option<String>,
    pub mod_date: Option<String>,
}

/// A single page.
#[derive(Debug, Clone, Serialize)]
pub struct Page {
    /// Page width in points (1/72 inch).
    pub width: f64,
    /// Page height in points.
    pub height: f64,
    /// Media box (the full page boundary).
    pub media_box: Rectangle,
    /// Crop box (visible area), defaults to media_box.
    pub crop_box: Option<Rectangle>,
    /// Content elements on this page.
    pub contents: Vec<ContentElement>,
    /// Page-level resources (fonts, images, etc.) are resolved during parsing.
    pub rotation: u16,
}

/// A rectangle defined by lower-left and upper-right corners (in points).
#[derive(Debug, Clone, Copy, Serialize)]
pub struct Rectangle {
    pub llx: f64,
    pub lly: f64,
    pub urx: f64,
    pub ury: f64,
}

impl Rectangle {
    pub fn width(&self) -> f64 {
        self.urx - self.llx
    }

    pub fn height(&self) -> f64 {
        self.ury - self.lly
    }
}

/// A content element on a page.
#[derive(Debug, Clone, Serialize)]
pub enum ContentElement {
    Text(TextSpan),
    Path(PathData),
    Image(ImageData),
    /// Set a clipping region (intersects with current clip).
    ClipPath(ClipPathData),
    /// Save graphics state (corresponds to PDF `q` operator).
    SaveState,
    /// Restore graphics state (corresponds to PDF `Q` operator).
    RestoreState,
}

/// A span of text with position and style.
#[derive(Debug, Clone, Serialize)]
pub struct TextSpan {
    pub x: f64,
    pub y: f64,
    pub text: String,
    pub font_name: String,
    pub font_size: f64,
    pub fill_color: Color,
    /// Extra character spacing in points (PDF Tc operator). 0.0 = default.
    pub character_spacing: f64,
}

/// Vector path data.
#[derive(Debug, Clone, Serialize)]
pub struct PathData {
    pub operations: Vec<PathOp>,
    pub stroke: Option<StrokeStyle>,
    pub fill: Option<Color>,
}

#[derive(Debug, Clone, Serialize)]
pub enum PathOp {
    MoveTo(f64, f64),
    LineTo(f64, f64),
    CurveTo(f64, f64, f64, f64, f64, f64),
    ClosePath,
}

#[derive(Debug, Clone, Serialize)]
pub struct StrokeStyle {
    pub color: Color,
    pub width: f64,
    pub line_cap: LineCap,
    pub line_join: LineJoin,
}

#[derive(Debug, Clone, Copy, Serialize)]
pub enum LineCap {
    Butt,
    Round,
    Square,
}

#[derive(Debug, Clone, Copy, Serialize)]
pub enum LineJoin {
    Miter,
    Round,
    Bevel,
}

/// A color value.
#[derive(Debug, Clone, Copy, Serialize)]
pub enum Color {
    Gray(f64),
    Rgb(f64, f64, f64),
    Cmyk(f64, f64, f64, f64),
}

/// A clipping path that constrains subsequent drawing.
#[derive(Debug, Clone, Serialize)]
pub struct ClipPathData {
    pub operations: Vec<PathOp>,
    /// True for even-odd rule (W*), false for non-zero winding (W).
    pub even_odd: bool,
}

/// An embedded image.
#[derive(Debug, Clone, Serialize)]
pub struct ImageData {
    pub x: f64,
    pub y: f64,
    pub width: f64,
    pub height: f64,
    pub data: Vec<u8>,
    pub color_space: ColorSpace,
    pub bits_per_component: u8,
}

#[derive(Debug, Clone, Copy, Serialize)]
pub enum ColorSpace {
    DeviceGray,
    DeviceRgb,
    DeviceCmyk,
}

/// A bookmark / outline entry.
#[derive(Debug, Clone, Serialize)]
pub struct OutlineItem {
    pub title: String,
    pub page_index: Option<usize>,
    pub children: Vec<OutlineItem>,
}
