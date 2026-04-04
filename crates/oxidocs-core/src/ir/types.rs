use serde::{Deserialize, Serialize};
use std::collections::HashMap;

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct Document {
    pub pages: Vec<Page>,
    pub styles: StyleSheet,
    pub metadata: DocumentMetadata,
    /// Comments referenced in the document
    #[serde(default, skip_serializing_if = "Vec::is_empty")]
    pub comments: Vec<Comment>,
    /// Compatibility: adjustLineHeightInTable (compat65).
    /// true = adjust line height in table cells (disable grid snap in cells).
    /// false (default) = table cells snap to document grid like normal paragraphs.
    #[serde(default)]
    pub adjust_line_height_in_table: bool,
    /// Default tab stop interval from w:settings/w:defaultTabStop (in points).
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub default_tab_stop: Option<f32>,
    /// Compatibility mode (from w:settings/w:compat/w:compatSetting w:name="compatibilityMode")
    /// 14=Word 2010, 15=Word 2013+. Affects table cell grid snap behavior.
    #[serde(default)]
    pub compat_mode: u32,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct Page {
    pub blocks: Vec<Block>,
    pub size: PageSize,
    pub margin: Margin,
    /// Document grid line pitch in points (from w:docGrid w:linePitch).
    /// When set with grid_type "lines" or "linesAndChars", line spacing
    /// snaps to multiples of this pitch.
    #[serde(default)]
    pub grid_line_pitch: Option<f32>,
    /// Character grid pitch in points (from w:docGrid w:charSpace for linesAndChars).
    /// When set, character widths are expanded to align to this grid.
    #[serde(default)]
    pub grid_char_pitch: Option<f32>,
    /// True when docGrid element exists but has NO type attribute.
    /// CJK 83/64 multiplier is NOT applied; COM-measured Single heights used instead.
    #[serde(default)]
    pub doc_grid_no_type: bool,
    /// Header content (paragraphs from header part)
    #[serde(default, skip_serializing_if = "Vec::is_empty")]
    pub header: Vec<Block>,
    /// Footer content (paragraphs from footer part)
    #[serde(default, skip_serializing_if = "Vec::is_empty")]
    pub footer: Vec<Block>,
    /// Header distance from page top edge in points (w:pgMar header attr)
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub header_distance: Option<f32>,
    /// Footer distance from page bottom edge in points (w:pgMar footer attr)
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub footer_distance: Option<f32>,
    /// Footnotes referenced in this page
    #[serde(default, skip_serializing_if = "Vec::is_empty")]
    pub footnotes: Vec<Footnote>,
    /// Endnotes referenced in this page
    #[serde(default, skip_serializing_if = "Vec::is_empty")]
    pub endnotes: Vec<Footnote>,
    /// Floating images (anchored, not inline)
    #[serde(default, skip_serializing_if = "Vec::is_empty")]
    pub floating_images: Vec<Image>,
    /// Text boxes
    #[serde(default, skip_serializing_if = "Vec::is_empty")]
    pub text_boxes: Vec<TextBox>,
    /// Geometric shapes (DrawingML / VML)
    #[serde(default, skip_serializing_if = "Vec::is_empty")]
    pub shapes: Vec<Shape>,
    /// Column layout for this section
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub columns: Option<ColumnLayout>,
    /// Page number format (e.g. "decimal", "lowerRoman", "upperRoman", "lowerLetter", "upperLetter")
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub page_number_format: Option<String>,
    /// Starting page number for this section
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub page_number_start: Option<u32>,
    /// Page borders
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub page_borders: Option<PageBorders>,
}

/// Page border definitions
#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct PageBorders {
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub top: Option<BorderDef>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub bottom: Option<BorderDef>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub left: Option<BorderDef>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub right: Option<BorderDef>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub enum Block {
    Paragraph(Paragraph),
    Table(Table),
    Image(Image),
    /// Placeholder for unsupported content (SmartArt, Chart, etc.)
    UnsupportedElement(UnsupportedElement),
}

/// Represents an unsupported OOXML element that was skipped during parsing
#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct UnsupportedElement {
    /// Type of unsupported element (e.g. "SmartArt", "Chart", "ActiveX")
    pub element_type: String,
    /// Optional fallback image data (base64)
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub fallback_image: Option<String>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct Paragraph {
    pub runs: Vec<Run>,
    pub style: ParagraphStyle,
    pub alignment: Alignment,
    /// Inline/anchor shapes attached to this paragraph (e.g. bracketPair)
    #[serde(default, skip_serializing_if = "Vec::is_empty")]
    pub shapes: Vec<Shape>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct Run {
    pub text: String,
    pub style: RunStyle,
    /// Hyperlink URL (external) or anchor (internal bookmark)
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub url: Option<String>,
    /// Footnote reference number (1-based)
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub footnote_ref: Option<u32>,
    /// Endnote reference number (1-based)
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub endnote_ref: Option<u32>,
    /// Comment IDs that start at this run
    #[serde(default, skip_serializing_if = "Vec::is_empty")]
    pub comment_range_start: Vec<String>,
    /// Comment IDs that end at this run
    #[serde(default, skip_serializing_if = "Vec::is_empty")]
    pub comment_range_end: Vec<String>,
    /// Tracked change info (insertion/deletion)
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub tracked_change: Option<TrackedChange>,
    /// Ruby (furigana) annotation
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub ruby: Option<Ruby>,
    /// Bookmark anchor name (from w:bookmarkStart w:name)
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub bookmark_name: Option<String>,
    /// Whether this run contains OMML math content
    #[serde(default, skip_serializing_if = "std::ops::Not::not")]
    pub is_math: bool,
    /// Field type for dynamic content substitution (PAGE, NUMPAGES, etc.)
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub field_type: Option<FieldType>,
}

/// Field types for dynamic content that gets resolved during layout
#[derive(Debug, Clone, Copy, PartialEq, Eq, Serialize, Deserialize)]
pub enum FieldType {
    /// Current page number (PAGE field)
    Page,
    /// Total number of pages (NUMPAGES field)
    NumPages,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct RunStyle {
    pub font_family: Option<String>,
    /// East Asian font family (w:rFonts eastAsia) for CJK characters
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub font_family_east_asia: Option<String>,
    pub font_size: Option<f32>,
    pub bold: bool,
    pub italic: bool,
    pub underline: bool,
    /// Underline style (e.g. "single", "double", "wave", "dash", "dotted", "thick")
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub underline_style: Option<String>,
    pub strikethrough: bool,
    /// Double strikethrough (w:dstrike)
    #[serde(default)]
    pub double_strikethrough: bool,
    pub color: Option<String>,
    pub highlight: Option<String>,
    pub vertical_align: Option<VerticalAlign>,
    /// Character spacing in points (w:spacing w:val in twips, converted to pt)
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub character_spacing: Option<f32>,
    /// Small capitals (w:smallCaps)
    #[serde(default)]
    pub small_caps: bool,
    /// All capitals (w:caps)
    #[serde(default)]
    pub all_caps: bool,
    /// Character-level shading/background color (w:shd fill, hex)
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub shading: Option<String>,
    /// Right-to-left text run (w:rtl)
    #[serde(default)]
    pub rtl: bool,
    /// Hidden text (w:vanish)
    #[serde(default)]
    pub vanish: bool,
    /// Text outline effect (w:outline)
    #[serde(default)]
    pub outline: bool,
    /// Text shadow effect (w:shadow)
    #[serde(default)]
    pub shadow: bool,
    /// Text emboss effect (w:emboss)
    #[serde(default)]
    pub emboss: bool,
    /// Text imprint/engrave effect (w:imprint)
    #[serde(default)]
    pub imprint: bool,
    /// Complex script font size in points (w:szCs, half-points / 2)
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub font_size_cs: Option<f32>,
    /// Complex script bold (w:bCs)
    #[serde(default)]
    pub bold_cs: bool,
    /// Complex script italic (w:iCs)
    #[serde(default)]
    pub italic_cs: bool,
    /// Character kerning threshold in points (w:kern, half-points / 2)
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub kern: Option<f32>,
    /// Fit text width in points (w:fitText, twips / 20)
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub fit_text: Option<f32>,
    /// East Asian layout: combine (kumimoji)
    #[serde(default)]
    pub combine: bool,
    /// East Asian layout: vertical-in-horizontal (tate-chu-yoko)
    #[serde(default)]
    pub vert_in_horz: bool,
}

impl Default for RunStyle {
    fn default() -> Self {
        Self {
            font_family: None,
            font_family_east_asia: None,
            font_size: None,
            bold: false,
            italic: false,
            underline: false,
            underline_style: None,
            strikethrough: false,
            double_strikethrough: false,
            color: None,
            highlight: None,
            vertical_align: None,
            character_spacing: None,
            small_caps: false,
            all_caps: false,
            shading: None,
            rtl: false,
            vanish: false,
            outline: false,
            shadow: false,
            emboss: false,
            imprint: false,
            font_size_cs: None,
            bold_cs: false,
            italic_cs: false,
            kern: None,
            fit_text: None,
            combine: false,
            vert_in_horz: false,
        }
    }
}

#[derive(Debug, Clone, Copy, Serialize, Deserialize, PartialEq)]
pub enum VerticalAlign {
    Baseline,
    Superscript,
    Subscript,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct Table {
    pub rows: Vec<TableRow>,
    pub style: TableStyle,
    /// Column widths from tblGrid/gridCol in points
    #[serde(default, skip_serializing_if = "Vec::is_empty")]
    pub grid_columns: Vec<f32>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct TableRow {
    pub cells: Vec<TableCell>,
    /// Row height in points (w:trHeight)
    #[serde(default)]
    pub height: Option<f32>,
    /// Height rule: "exact" or "atLeast" (default)
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub height_rule: Option<String>,
    /// Repeat as header row at top of each page (w:tblHeader)
    #[serde(default)]
    pub header: bool,
    /// Prevent row from breaking across pages (w:cantSplit)
    #[serde(default)]
    pub cant_split: bool,
    /// Number of grid columns to skip at start of row (w:gridBefore)
    #[serde(default)]
    pub grid_before: u32,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct TableCell {
    pub blocks: Vec<Block>,
    pub width: Option<f32>,
    /// Horizontal merge span (w:gridSpan), default 1
    #[serde(default = "default_one")]
    pub grid_span: u32,
    /// Vertical merge: "restart" starts a new merged cell, "continue" is merged into above
    #[serde(default)]
    pub v_merge: Option<String>,
    /// Cell shading/background color (hex)
    #[serde(default)]
    pub shading: Option<String>,
    /// Vertical alignment within cell
    #[serde(default)]
    pub v_align: Option<String>,
    /// Cell-specific borders (override table borders)
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub borders: Option<CellBorders>,
    /// Cell margins/padding in points
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub margins: Option<CellMargins>,
}

/// Cell border definitions
#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct CellBorders {
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub top: Option<BorderDef>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub bottom: Option<BorderDef>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub left: Option<BorderDef>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub right: Option<BorderDef>,
}

/// Cell margin/padding
#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct CellMargins {
    #[serde(default)]
    pub top: Option<f32>,
    #[serde(default)]
    pub bottom: Option<f32>,
    #[serde(default)]
    pub left: Option<f32>,
    #[serde(default)]
    pub right: Option<f32>,
}

fn default_one() -> u32 { 1 }

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct Image {
    pub data: Vec<u8>,
    pub width: f32,
    pub height: f32,
    pub alt_text: Option<String>,
    pub content_type: Option<String>,
    /// Floating position (None = inline image)
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub position: Option<FloatingPosition>,
    /// Text wrapping mode
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub wrap_type: Option<WrapType>,
    /// Crop percentages (a:srcRect) — top, right, bottom, left as 0-100%
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub crop: Option<ImageCrop>,
    /// Index of the anchor paragraph block (for paragraph-relative positioning)
    #[serde(default)]
    pub anchor_block_index: usize,
}

/// Image crop rectangle (percentages from each edge)
#[derive(Debug, Clone, Copy, Serialize, Deserialize)]
pub struct ImageCrop {
    #[serde(default)]
    pub top: f32,
    #[serde(default)]
    pub right: f32,
    #[serde(default)]
    pub bottom: f32,
    #[serde(default)]
    pub left: f32,
}

/// Position for a floating (anchored) element
#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct FloatingPosition {
    /// Horizontal offset in points from anchor
    pub x: f32,
    /// Vertical offset in points from anchor
    pub y: f32,
    /// Horizontal anchor reference (e.g. "column", "page", "margin")
    #[serde(default)]
    pub h_relative: Option<String>,
    /// Vertical anchor reference
    #[serde(default)]
    pub v_relative: Option<String>,
    /// Horizontal alignment (e.g. "left", "center", "right")
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub h_align: Option<String>,
    /// Vertical alignment (e.g. "top", "center", "bottom")
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub v_align: Option<String>,
}

/// Text wrapping mode for floating elements
#[derive(Debug, Clone, Copy, PartialEq, Eq, Serialize, Deserialize)]
pub enum WrapType {
    /// No wrapping (in front/behind text)
    None,
    /// Square wrapping
    Square,
    /// Tight wrapping
    Tight,
    /// Top and bottom only
    TopAndBottom,
}

/// A text box (from w:txbxContent or wps:txbx)
#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct TextBox {
    /// Content paragraphs inside the text box
    pub blocks: Vec<Block>,
    /// Width in points
    pub width: f32,
    /// Height in points
    pub height: f32,
    /// Position (floating)
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub position: Option<FloatingPosition>,
    /// Border style
    #[serde(default)]
    pub border: bool,
    /// Border stroke color (hex)
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub stroke_color: Option<String>,
    /// Border stroke width in points
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub stroke_width: Option<f32>,
    /// Background color (hex)
    #[serde(default)]
    pub fill: Option<String>,
    /// Index of the anchor block (paragraph) in page.blocks
    #[serde(default)]
    pub anchor_block_index: usize,
    /// Corner radius for rounded rectangles (in points)
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub corner_radius: Option<f32>,
    /// Text inset left (in points, default 7.2pt = 91440 EMU)
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub inset_left: Option<f32>,
    /// Text inset right (in points, default 7.2pt)
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub inset_right: Option<f32>,
    /// Text inset top (in points, default 3.6pt = 45720 EMU)
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub inset_top: Option<f32>,
    /// Text inset bottom (in points, default 3.6pt)
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub inset_bottom: Option<f32>,
    /// Wrap type for text wrapping around this text box
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub wrap_type: Option<WrapType>,
}

/// A geometric shape (DrawingML or VML)
#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct Shape {
    /// Shape type (e.g. "rect", "ellipse", "roundRect", "line", "arrow", etc.)
    pub shape_type: String,
    /// Width in points
    pub width: f32,
    /// Height in points
    pub height: f32,
    /// Position (floating)
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub position: Option<FloatingPosition>,
    /// Fill color (hex)
    #[serde(default)]
    pub fill: Option<String>,
    /// Outline/stroke color (hex)
    #[serde(default)]
    pub stroke_color: Option<String>,
    /// Outline width in points
    #[serde(default)]
    pub stroke_width: Option<f32>,
    /// Text content inside the shape (if any)
    #[serde(default, skip_serializing_if = "Vec::is_empty")]
    pub text_blocks: Vec<Block>,
    /// Rotation in degrees
    #[serde(default)]
    pub rotation: Option<f32>,
    /// Gradient fill stops (from a:gradFill)
    #[serde(default, skip_serializing_if = "Vec::is_empty")]
    pub gradient_stops: Vec<GradientStop>,
    /// Gradient fill angle in degrees
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub gradient_angle: Option<f32>,
    /// Index of the anchor paragraph block (for positioning)
    #[serde(default)]
    pub anchor_block_index: usize,
}

/// A gradient color stop
#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct GradientStop {
    /// Position as 0-100 percentage
    pub position: f32,
    /// Color hex
    pub color: String,
}

/// A comment annotation
#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct Comment {
    /// Comment ID
    pub id: String,
    /// Author name
    pub author: Option<String>,
    /// Date string
    pub date: Option<String>,
    /// Comment text paragraphs
    pub blocks: Vec<Block>,
}

/// A tracked change (insertion or deletion)
#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct TrackedChange {
    /// "insert" or "delete"
    pub change_type: String,
    /// Author of the change
    pub author: Option<String>,
    /// Date of the change
    pub date: Option<String>,
}

/// Ruby (furigana) annotation
#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct Ruby {
    /// Base text (the main character(s))
    pub base: String,
    /// Annotation text (furigana reading)
    pub text: String,
    /// Font size of the ruby text in points
    #[serde(default)]
    pub font_size: Option<f32>,
}

/// Column layout definition (from w:cols)
#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct ColumnLayout {
    /// Number of columns
    pub num: u32,
    /// Space between columns in points
    #[serde(default)]
    pub space: Option<f32>,
    /// Whether columns have equal width
    #[serde(default)]
    pub equal_width: bool,
    /// Individual column definitions (for unequal widths)
    #[serde(default, skip_serializing_if = "Vec::is_empty")]
    pub columns: Vec<ColumnDef>,
}

/// Individual column definition
#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct ColumnDef {
    /// Column width in points
    pub width: f32,
    /// Space after this column in points
    #[serde(default)]
    pub space: Option<f32>,
}

/// A tab stop definition
#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct TabStop {
    /// Position in points from the left margin
    pub position: f32,
    /// Alignment at the tab stop
    pub alignment: TabStopAlignment,
    /// Leader character
    #[serde(default)]
    pub leader: Option<String>,
}

#[derive(Debug, Clone, Copy, PartialEq, Eq, Serialize, Deserialize)]
pub enum TabStopAlignment {
    Left,
    Center,
    Right,
    Decimal,
}

impl Default for TabStopAlignment {
    fn default() -> Self {
        Self::Left
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq, Serialize, Deserialize)]
pub enum Alignment {
    Left,
    Center,
    Right,
    Justify,
    Distribute,
}

impl Default for Alignment {
    fn default() -> Self {
        Self::Left
    }
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct ParagraphStyle {
    pub heading_level: Option<u8>,
    /// Line spacing value. Interpretation depends on line_spacing_rule:
    /// - "auto": multiplier (w:line / 240, e.g. 1.15 for w:line="276")
    /// - "exact": fixed height in points (w:line / 20)
    /// - "atLeast": minimum height in points (w:line / 20)
    /// None means single spacing (1.0).
    pub line_spacing: Option<f32>,
    /// Line spacing rule: "auto" (default), "exact", or "atLeast"
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub line_spacing_rule: Option<String>,
    pub space_before: Option<f32>,
    pub space_after: Option<f32>,
    /// True when spacing was directly specified in paragraph's pPr (not just inherited from style).
    /// Word resets inherited spacing to Single/0 inside table cells.
    #[serde(default)]
    pub has_direct_spacing: bool,
    /// True when line_spacing was inherited from docDefaults pPrDefault (not from Normal style or direct).
    /// Word resets docDefaults lineSpacing to Single inside table cells but keeps Normal style's lineSpacing.
    #[serde(default)]
    pub line_spacing_from_doc_defaults: bool,
    /// w:spacing beforeLines — in 1/100 of a line (raw value from OOXML)
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub before_lines: Option<f32>,
    /// w:spacing afterLines — in 1/100 of a line (raw value from OOXML)
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub after_lines: Option<f32>,
    pub indent_left: Option<f32>,
    pub indent_right: Option<f32>,
    pub indent_first_line: Option<f32>,
    /// Default run style from style definition (font size, bold, etc.)
    pub default_run_style: Option<RunStyle>,
    /// Pre-resolved list marker text (e.g., "•", "1.", "a)")
    pub list_marker: Option<String>,
    /// Hanging indent for the list marker in points
    pub list_indent: Option<f32>,
    /// Suffix after list number: "tab" (default), "space", or "nothing"
    #[serde(default)]
    pub list_suff: Option<String>,
    /// Tab stop position for list numbering (in points)
    #[serde(default)]
    pub list_tab_stop: Option<f32>,
    /// Whether this paragraph snaps to the document grid (default: true).
    #[serde(default = "default_true")]
    pub snap_to_grid: bool,
    /// w:contextualSpacing: suppress space_before/after between paragraphs of the same style.
    #[serde(default)]
    pub contextual_spacing: bool,
    /// Style ID (e.g. "Normal", "Heading1") for contextual spacing comparison.
    #[serde(default)]
    pub style_id: Option<String>,
    /// Custom tab stops (w:tabs)
    #[serde(default, skip_serializing_if = "Vec::is_empty")]
    pub tab_stops: Vec<TabStop>,
    /// Paragraph background/shading color (hex from w:shd fill)
    #[serde(default)]
    pub shading: Option<String>,
    /// pPr/rPr: paragraph-level default run properties for empty paragraph height.
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub ppr_rpr: Option<RunStyle>,
    /// Page break before this paragraph (w:pageBreakBefore)
    #[serde(default)]
    pub page_break_before: bool,
    /// Paragraph borders
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub borders: Option<ParagraphBorders>,
    /// Keep with next paragraph on same page (w:keepNext)
    #[serde(default)]
    pub keep_next: bool,
    /// Keep all lines of this paragraph together (w:keepLines)
    #[serde(default)]
    pub keep_lines: bool,
    /// Widow/orphan control (w:widowControl, default true in Word)
    #[serde(default = "default_true")]
    pub widow_control: bool,
    /// Whether widowControl was explicitly set in XML (for docDefaults inheritance)
    #[serde(default, skip_serializing)]
    pub has_explicit_widow_control: bool,
    /// Auto space between East Asian and Western text (w:autoSpaceDE, default true)
    #[serde(default = "default_true")]
    pub auto_space_de: bool,
    /// Auto space between East Asian and numbers (w:autoSpaceDN, default true)
    #[serde(default = "default_true")]
    pub auto_space_dn: bool,
    /// Bidirectional text / RTL paragraph (w:bidi)
    #[serde(default)]
    pub bidi: bool,
    /// Numbering ID from style definition (w:numPr/w:numId)
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub num_id: Option<String>,
    /// Numbering indent level from style definition (w:numPr/w:ilvl)
    #[serde(default)]
    pub num_ilvl: u8,
}

/// Paragraph border definitions
#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct ParagraphBorders {
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub top: Option<BorderDef>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub bottom: Option<BorderDef>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub left: Option<BorderDef>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub right: Option<BorderDef>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub between: Option<BorderDef>,
}

/// A single border definition
#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct BorderDef {
    /// Border style (e.g. "single", "double", "dashed", "dotted", "thick")
    pub style: String,
    /// Width in points (w:sz is in 1/8 pt)
    pub width: f32,
    /// Color hex
    pub color: Option<String>,
    /// Distance from text in points (w:space)
    #[serde(default)]
    pub space: f32,
}

fn default_true() -> bool { true }

impl Default for ParagraphStyle {
    fn default() -> Self {
        Self {
            heading_level: None,
            line_spacing: None,
            line_spacing_rule: None,
            space_before: None,
            space_after: None,
            has_direct_spacing: false,
            line_spacing_from_doc_defaults: false,
            before_lines: None,
            after_lines: None,
            indent_left: None,
            indent_right: None,
            indent_first_line: None,
            default_run_style: None,
            list_marker: None,
            list_indent: None,
            list_suff: None,
            list_tab_stop: None,
            snap_to_grid: true,
            contextual_spacing: false,
            style_id: None,
            tab_stops: Vec::new(),
            shading: None,
            ppr_rpr: None,
            page_break_before: false,
            borders: None,
            keep_next: false,
            keep_lines: false,
            widow_control: true,
            has_explicit_widow_control: false,
            auto_space_de: true,
            auto_space_dn: true,
            bidi: false,
            num_id: None,
            num_ilvl: 0,
        }
    }
}

#[derive(Debug, Clone, Serialize, Deserialize, Default)]
pub struct TableStyle {
    pub border: bool,
    /// Whether the table has inside horizontal borders (insideH)
    #[serde(default)]
    pub has_inside_h: bool,
    /// Border color (hex), e.g. "000000"
    #[serde(default)]
    pub border_color: Option<String>,
    /// Border width in points (w:sz is in 1/8 pt)
    #[serde(default)]
    pub border_width: Option<f32>,
    /// Border style (e.g. "single", "double", "dashed")
    #[serde(default)]
    pub border_style: Option<String>,
    /// Table width in points (from w:tblW)
    #[serde(default)]
    pub width: Option<f32>,
    /// Table width type: "dxa" (fixed), "pct" (percentage), "auto"
    #[serde(default)]
    pub width_type: Option<String>,
    /// Table alignment (w:jc): "left", "center", "right"
    #[serde(default)]
    pub alignment: Option<String>,
    /// Table style ID reference (w:tblStyle)
    #[serde(default)]
    pub style_id: Option<String>,
    /// Table look flags (from w:tblLook)
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub tbl_look: Option<TableLook>,
    /// Table indent from left margin in points (w:tblInd)
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub indent: Option<f32>,
    /// Cell spacing in points (w:tblCellSpacing)
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub cell_spacing: Option<f32>,
    /// Table layout mode: "fixed" or "autofit" (w:tblLayout)
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub layout: Option<String>,
    /// Default cell margins in points (w:tblCellMar)
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub default_cell_margins: Option<CellMargins>,
    /// Paragraph properties from table style (pPr) — applied to cell paragraphs as fallback
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub para_style: Option<ParagraphStyle>,
    /// Paragraph alignment from table style pPr (jc) — applied as fallback
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub para_alignment: Option<Alignment>,
    /// Table floating position (w:tblpPr)
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub position: Option<TablePosition>,
}

/// Floating table position properties (w:tblpPr)
#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct TablePosition {
    /// Horizontal offset in points (w:tblpX)
    #[serde(default)]
    pub x: f32,
    /// Vertical offset in points (w:tblpY)
    #[serde(default)]
    pub y: f32,
    /// Horizontal anchor: "text", "margin", "page"
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub h_anchor: Option<String>,
    /// Vertical anchor: "text", "margin", "page"
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub v_anchor: Option<String>,
    /// Horizontal alignment spec: "left", "center", "right"
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub h_align: Option<String>,
    /// Distance from surrounding text (points)
    #[serde(default)]
    pub left_from_text: f32,
    #[serde(default)]
    pub right_from_text: f32,
    #[serde(default)]
    pub top_from_text: f32,
    #[serde(default)]
    pub bottom_from_text: f32,
}

/// Table look conditional formatting flags (w:tblLook)
#[derive(Debug, Clone, Copy, Serialize, Deserialize, Default)]
pub struct TableLook {
    /// Apply first row conditional style
    #[serde(default)]
    pub first_row: bool,
    /// Apply last row conditional style
    #[serde(default)]
    pub last_row: bool,
    /// Apply first column conditional style
    #[serde(default)]
    pub first_column: bool,
    /// Apply last column conditional style
    #[serde(default)]
    pub last_column: bool,
    /// Show horizontal banding (alternating row shading)
    #[serde(default)]
    pub banded_rows: bool,
    /// Show vertical banding (alternating column shading)
    #[serde(default)]
    pub banded_columns: bool,
    /// Row band size (number of rows per band, default 1)
    #[serde(default = "default_one")]
    pub row_band_size: u32,
    /// Column band size (number of columns per band, default 1)
    #[serde(default = "default_one")]
    pub col_band_size: u32,
}

#[derive(Debug, Clone, Copy, Serialize, Deserialize)]
pub struct PageSize {
    pub width: f32,
    pub height: f32,
}

impl Default for PageSize {
    fn default() -> Self {
        // A4 in points (210mm x 297mm)
        Self {
            width: 595.0,
            height: 842.0,
        }
    }
}

#[derive(Debug, Clone, Copy, Serialize, Deserialize)]
pub struct Margin {
    pub top: f32,
    pub bottom: f32,
    pub left: f32,
    pub right: f32,
}

impl Default for Margin {
    fn default() -> Self {
        // Word default margins in points (1 inch = 72pt)
        Self {
            top: 72.0,
            bottom: 72.0,
            left: 72.0,
            right: 72.0,
        }
    }
}

#[derive(Debug, Clone, Serialize, Deserialize, Default)]
pub struct StyleSheet {
    pub styles: HashMap<String, StyleDefinition>,
    /// Default run properties from w:docDefaults/w:rPrDefault
    pub doc_default_run_style: Option<RunStyle>,
    /// Default paragraph properties from w:docDefaults/w:pPrDefault
    pub doc_default_para_style: Option<ParagraphStyle>,
    /// Default paragraph alignment from w:docDefaults/w:pPrDefault/w:jc
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub doc_default_alignment: Option<Alignment>,
    /// Table style borders: style_id -> TableStyle (with border info from tblBorders)
    #[serde(default, skip_serializing_if = "HashMap::is_empty")]
    pub table_styles: HashMap<String, TableStyle>,
    /// Default paragraph style ID (w:type="paragraph" w:default="1")
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub default_paragraph_style_id: Option<String>,
}

/// A named style definition with inheritance
#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct StyleDefinition {
    /// Style ID
    pub style_id: String,
    /// Parent style ID (w:basedOn)
    #[serde(default)]
    pub based_on: Option<String>,
    /// Paragraph properties defined in this style
    pub paragraph: ParagraphStyle,
    /// Paragraph alignment from this style (w:jc)
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub alignment: Option<Alignment>,
    /// Whether inheritance has been resolved
    #[serde(skip, default)]
    pub resolved: bool,
}

#[derive(Debug, Clone, Serialize, Deserialize, Default)]
pub struct DocumentMetadata {
    pub title: Option<String>,
    pub author: Option<String>,
    pub description: Option<String>,
}

/// A footnote or endnote
#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct Footnote {
    /// Note number (1-based, matching the reference in the body)
    pub number: u32,
    /// Content paragraphs of the note
    pub blocks: Vec<Block>,
}
