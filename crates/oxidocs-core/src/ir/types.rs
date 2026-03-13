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
    /// Header content (paragraphs from header part)
    #[serde(default, skip_serializing_if = "Vec::is_empty")]
    pub header: Vec<Block>,
    /// Footer content (paragraphs from footer part)
    #[serde(default, skip_serializing_if = "Vec::is_empty")]
    pub footer: Vec<Block>,
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
    /// Column layout for this section
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub columns: Option<ColumnLayout>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub enum Block {
    Paragraph(Paragraph),
    Table(Table),
    Image(Image),
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct Paragraph {
    pub runs: Vec<Run>,
    pub style: ParagraphStyle,
    pub alignment: Alignment,
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
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct RunStyle {
    pub font_family: Option<String>,
    pub font_size: Option<f32>,
    pub bold: bool,
    pub italic: bool,
    pub underline: bool,
    pub strikethrough: bool,
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
}

impl Default for RunStyle {
    fn default() -> Self {
        Self {
            font_family: None,
            font_size: None,
            bold: false,
            italic: false,
            underline: false,
            strikethrough: false,
            color: None,
            highlight: None,
            vertical_align: None,
            character_spacing: None,
            small_caps: false,
            all_caps: false,
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
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct TableRow {
    pub cells: Vec<TableCell>,
    /// Row height in points (w:trHeight)
    #[serde(default)]
    pub height: Option<f32>,
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
    /// Background color (hex)
    #[serde(default)]
    pub fill: Option<String>,
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
}

impl Default for Alignment {
    fn default() -> Self {
        Self::Left
    }
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct ParagraphStyle {
    pub heading_level: Option<u8>,
    /// Line spacing multiplier (w:line / 240 for auto mode, e.g. 1.15 for w:line="276").
    /// None means single spacing (1.0).
    pub line_spacing: Option<f32>,
    pub space_before: Option<f32>,
    pub space_after: Option<f32>,
    pub indent_left: Option<f32>,
    pub indent_right: Option<f32>,
    pub indent_first_line: Option<f32>,
    /// Default run style from style definition (font size, bold, etc.)
    pub default_run_style: Option<RunStyle>,
    /// Pre-resolved list marker text (e.g., "•", "1.", "a)")
    pub list_marker: Option<String>,
    /// Hanging indent for the list marker in points
    pub list_indent: Option<f32>,
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
    /// Page break before this paragraph (w:pageBreakBefore)
    #[serde(default)]
    pub page_break_before: bool,
}

fn default_true() -> bool { true }

impl Default for ParagraphStyle {
    fn default() -> Self {
        Self {
            heading_level: None,
            line_spacing: None,
            space_before: None,
            space_after: None,
            indent_left: None,
            indent_right: None,
            indent_first_line: None,
            default_run_style: None,
            list_marker: None,
            list_indent: None,
            snap_to_grid: true,
            contextual_spacing: false,
            style_id: None,
            tab_stops: Vec::new(),
            shading: None,
            page_break_before: false,
        }
    }
}

#[derive(Debug, Clone, Serialize, Deserialize, Default)]
pub struct TableStyle {
    pub border: bool,
    /// Border color (hex), e.g. "000000"
    #[serde(default)]
    pub border_color: Option<String>,
    /// Border width in points (w:sz is in 1/8 pt)
    #[serde(default)]
    pub border_width: Option<f32>,
    /// Border style (e.g. "single", "double", "dashed")
    #[serde(default)]
    pub border_style: Option<String>,
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
    pub styles: HashMap<String, ParagraphStyle>,
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
