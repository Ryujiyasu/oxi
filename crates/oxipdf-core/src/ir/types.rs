use serde::Serialize;

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
