mod kinsoku;

use crate::font::{FontMetrics, FontMetricsRegistry};
use crate::ir::*;

/// Pre-allocated single-character strings to avoid heap allocation in hot loops.
const TAB_STRING: &str = "\t";
const SPACE_STRING: &str = " ";

/// Convert a char to a String with pre-sized buffer (avoids realloc for multi-byte chars).
#[inline]
fn char_to_string(ch: char) -> String {
    let mut s = String::with_capacity(ch.len_utf8());
    s.push(ch);
    s
}

/// Characters that allow a line break AFTER them (English punctuation).
/// Word treats these as breakable opportunities similar to spaces.
fn is_break_after(ch: char) -> bool {
    matches!(ch, '-' | '/' | '\\' | ')' | ']' | '}' | '>' | '!' | '?' | ';' | ':' | ',')
}

/// Result of layout: positioned elements across pages
pub struct LayoutResult {
    pub pages: Vec<LayoutPage>,
}

pub struct LayoutPage {
    pub width: f32,
    pub height: f32,
    pub elements: Vec<LayoutElement>,
}

pub struct LayoutElement {
    pub x: f32,
    pub y: f32,
    pub width: f32,
    pub height: f32,
    pub content: LayoutContent,
    /// Source paragraph index in the document body (for hit testing / editing)
    pub paragraph_index: Option<usize>,
    /// Source run index within the paragraph
    pub run_index: Option<usize>,
    /// Character offset within the run's text where this fragment starts
    pub char_offset: Option<usize>,
}

impl LayoutElement {
    /// Create a non-text element (border, shading, image, etc.) with no source indices.
    fn new(x: f32, y: f32, width: f32, height: f32, content: LayoutContent) -> Self {
        Self { x, y, width, height, content, paragraph_index: None, run_index: None, char_offset: None }
    }

    /// Create a text element with source location for hit testing.
    fn text(x: f32, y: f32, width: f32, height: f32, content: LayoutContent,
            para_idx: usize, run_idx: usize, char_offset: usize) -> Self {
        Self { x, y, width, height, content,
               paragraph_index: Some(para_idx), run_index: Some(run_idx), char_offset: Some(char_offset) }
    }
}

pub enum LayoutContent {
    Text {
        text: String,
        font_size: f32,
        font_family: Option<String>,
        bold: bool,
        italic: bool,
        underline: bool,
        underline_style: Option<String>,
        strikethrough: bool,
        color: Option<String>,
        highlight: Option<String>,
        field_type: Option<FieldType>,
        /// Pixel-snapped character spacing in points (0.0 = no extra spacing)
        character_spacing: f32,
    },
    Image {
        data: Vec<u8>,
        content_type: Option<String>,
    },
    TableBorder {
        x1: f32,
        y1: f32,
        x2: f32,
        y2: f32,
        color: Option<String>,
        width: f32,
    },
    CellShading {
        color: String,
    },
    /// A filled/stroked rectangle, optionally with rounded corners.
    BoxRect {
        fill: Option<String>,
        stroke_color: Option<String>,
        stroke_width: f32,
        corner_radius: f32,
    },
    /// A preset shape outline (e.g. bracketPair, brace, etc.)
    PresetShape {
        shape_type: String,
        stroke_color: Option<String>,
        stroke_width: f32,
    },
    /// Begin a clipping region. All subsequent elements until ClipEnd are clipped to this rect.
    ClipStart,
    /// End the current clipping region (restore graphics state).
    ClipEnd,
}

pub struct LayoutEngine {
    default_font_size: f32,
    default_font_family: Option<String>,
    registry: FontMetricsRegistry,
    /// Compatibility: adjustLineHeightInTable=true disables grid snap in table cells.
    adjust_line_height_in_table: bool,
    /// Document-level default tab stop interval (from w:settings/w:defaultTabStop)
    default_tab_stop: f32,
    /// Compatibility mode: 14=Word 2010 (table cells no grid snap), 15=Word 2013+ (grid snap)
    compat_mode: u32,
}

/// Word's default heading font sizes (in points)
fn heading_default_font_size(level: u8) -> f32 {
    // Word default heading sizes (half-points in styles.xml → points)
    match level {
        1 => 14.0,  // sz=28
        2 => 13.0,  // sz=26
        3 => 11.0,  // sz=22 (default body)
        4 => 11.0,
        _ => 11.0,
    }
}

/// Snap character spacing to pixel grid (DPI=96 fixed).
/// Character spacing pixel-snap: Word converts twips→pixels at 96 DPI
/// using round-to-nearest integer division, then back to points.
///
/// Derived from COM measurement: comparing Word's actual character
/// positions (Range.Information) against input spacing values.
/// Example: cs=-0.45pt → -9tw → round(-9*96/1440) = -1px → -0.75pt
fn snap_character_spacing(cs_pt: f32) -> f32 {
    let cs_twips = (cs_pt * 20.0).round() as i64;
    // round-to-nearest integer division (twips × 96 / 1440)
    // For positive: (a*b + c/2) / c
    // For negative: (a*b - c/2) / c, using floor division (not truncation)
    let numer = cs_twips * 96;
    let denom = 1440_i64;
    let cs_px = if numer >= 0 {
        (numer + denom / 2) / denom
    } else {
        // Floor division for negative: -((-numer + denom/2) / denom)
        -((-numer + denom / 2) / denom)
    };
    cs_px as f32 * 72.0 / 96.0
}

impl LayoutEngine {
    pub fn new() -> Self {
        Self {
            default_font_size: 11.0,
            default_font_family: None,
            registry: FontMetricsRegistry::load(),
            adjust_line_height_in_table: false,
            default_tab_stop: 36.0,
            compat_mode: 15,
        }
    }

    /// Create a LayoutEngine with document-specific defaults from docDefaults
    pub fn for_document(doc: &Document) -> Self {
        let default_font_size = doc.styles.doc_default_run_style
            .as_ref()
            .and_then(|s| s.font_size)
            .unwrap_or(11.0);
        let default_font_family = doc.styles.doc_default_run_style
            .as_ref()
            .and_then(|s| s.font_family.clone());
        Self {
            default_font_size,
            default_font_family,
            registry: FontMetricsRegistry::load(),
            adjust_line_height_in_table: doc.adjust_line_height_in_table,
            default_tab_stop: doc.default_tab_stop.unwrap_or(36.0),
            compat_mode: doc.compat_mode,
        }
    }

    pub fn layout(&self, doc: &Document) -> LayoutResult {
        let mut pages = Vec::new();

        for page in &doc.pages {
            let laid_out = self.layout_page(page);
            pages.extend(laid_out);
        }

        // Post-layout pass: substitute PAGE and NUMPAGES field placeholders
        let total_pages = pages.len();
        for (page_idx, page) in pages.iter_mut().enumerate() {
            for elem in &mut page.elements {
                if let LayoutContent::Text { text, field_type: Some(ft), .. } = &mut elem.content {
                    match ft {
                        FieldType::Page => *text = format!("{}", page_idx + 1),
                        FieldType::NumPages => *text = format!("{}", total_pages),
                    }
                }
            }
        }

        LayoutResult { pages }
    }

    /// Resolve font size for a run, considering paragraph style defaults and heading level
    fn resolve_font_size(&self, run_style: &RunStyle, para_style: &ParagraphStyle) -> f32 {
        if let Some(fs) = run_style.font_size {
            return fs;
        }
        if let Some(ref drs) = para_style.default_run_style {
            if let Some(fs) = drs.font_size {
                return fs;
            }
        }
        if let Some(level) = para_style.heading_level {
            return heading_default_font_size(level);
        }
        self.default_font_size
    }

    /// Resolve font family for a run.
    /// For CJK text, prefer font_family_east_asia over font_family.
    fn resolve_font_family<'a>(&'a self, run_style: &'a RunStyle, para_style: &'a ParagraphStyle) -> Option<&'a str> {
        if let Some(ref ff) = run_style.font_family {
            return Some(ff.as_str());
        }
        if let Some(ref drs) = para_style.default_run_style {
            if let Some(ref ff) = drs.font_family {
                return Some(ff.as_str());
            }
        }
        // Fallback to document default font (docDefaults rPrDefault)
        self.default_font_family.as_deref()
    }

    /// Resolve font family considering East Asian font for CJK characters.
    fn resolve_font_family_for_text<'a>(&'a self, text: &str, run_style: &'a RunStyle, para_style: &'a ParagraphStyle) -> Option<&'a str> {
        let has_cjk = text.chars().any(|c| kinsoku::is_cjk(c));
        if has_cjk {
            // Prefer East Asian font for CJK text
            if let Some(ref ff) = run_style.font_family_east_asia {
                return Some(ff.as_str());
            }
            if let Some(ref drs) = para_style.default_run_style {
                if let Some(ref ff) = drs.font_family_east_asia {
                    return Some(ff.as_str());
                }
            }
        }
        self.resolve_font_family(run_style, para_style)
    }

    /// Get font metrics for a run (uses registry with font-family resolution)
    fn metrics_for(&self, run_style: &RunStyle, para_style: &ParagraphStyle) -> &FontMetrics {
        match self.resolve_font_family(run_style, para_style) {
            Some(family) => self.registry.get(family),
            None => self.registry.default_metrics(),
        }
    }

    /// Get font metrics considering East Asian font for CJK text.
    fn metrics_for_text(&self, text: &str, run_style: &RunStyle, para_style: &ParagraphStyle) -> &FontMetrics {
        match self.resolve_font_family_for_text(text, run_style, para_style) {
            Some(family) => self.registry.get(family),
            None => self.registry.default_metrics(),
        }
    }

    /// Get font metrics for a single character, using East Asian font for CJK.
    fn metrics_for_char(&self, ch: char, run_style: &RunStyle, para_style: &ParagraphStyle) -> &FontMetrics {
        if kinsoku::is_cjk(ch) {
            if let Some(m) = self.metrics_for_cjk(run_style, para_style) {
                return m;
            }
        }
        self.metrics_for(run_style, para_style)
    }

    /// Get East Asian font metrics if an east-asia font family is specified.
    /// Returns None if no east-asia font is set (caller should fall back to latin metrics).
    fn metrics_for_cjk(&self, run_style: &RunStyle, para_style: &ParagraphStyle) -> Option<&FontMetrics> {
        if let Some(ref ff) = run_style.font_family_east_asia {
            return Some(self.registry.get(ff.as_str()));
        }
        if let Some(ref drs) = para_style.default_run_style {
            if let Some(ref ff) = drs.font_family_east_asia {
                return Some(self.registry.get(ff.as_str()));
            }
        }
        None
    }

    /// Resolve bold for a run, considering paragraph style defaults
    fn resolve_bold(&self, run_style: &RunStyle, para_style: &ParagraphStyle) -> bool {
        if run_style.bold {
            return true;
        }
        if let Some(ref drs) = para_style.default_run_style {
            if drs.bold {
                return true;
            }
        }
        if let Some(level) = para_style.heading_level {
            return level <= 2;
        }
        false
    }

    fn resolve_color<'a>(&self, run_style: &'a RunStyle, para_style: &'a ParagraphStyle) -> Option<&'a str> {
        if let Some(ref c) = run_style.color {
            return Some(c.as_str());
        }
        if let Some(ref drs) = para_style.default_run_style {
            if let Some(ref c) = drs.color {
                return Some(c.as_str());
            }
        }
        None
    }

    fn resolve_italic(&self, run_style: &RunStyle, para_style: &ParagraphStyle) -> bool {
        if run_style.italic {
            return true;
        }
        if let Some(ref drs) = para_style.default_run_style {
            if drs.italic {
                return true;
            }
        }
        false
    }

    /// Default font metrics for the document (uses docDefaults font if set, otherwise Calibri).
    fn doc_default_metrics(&self) -> &FontMetrics {
        match self.default_font_family.as_deref() {
            Some(ff) => self.registry.get(ff),
            None => self.registry.default_metrics(),
        }
    }

    fn layout_page(&self, page: &Page) -> Vec<LayoutPage> {
        let total_content_width = page.size.width - page.margin.left - page.margin.right;
        // COM-confirmed (2026-04-03, order_08): when header extends below margin.top,
        // body content starts below the header (header pushes body down).
        // header_distance + header_content_height = header_bottom.
        // start_y = max(margin.top, header_bottom)
        let header_bottom = if !page.header.is_empty() {
            let header_y = page.header_distance.unwrap_or(36.0);
            let mut hdr_h = 0.0_f32;
            for block in &page.header {
                if let Block::Paragraph(para) = block {
                    let fs = para.runs.first()
                        .and_then(|r| r.style.font_size)
                        .unwrap_or(self.default_font_size);
                    let metrics = para.runs.first()
                        .map(|r| self.metrics_for(&r.style, &para.style))
                        .unwrap_or_else(|| self.doc_default_metrics());
                    let lh = metrics.word_line_height(fs, 96.0);
                    hdr_h += lh;
                    hdr_h += para.style.space_after.unwrap_or(0.0);
                }
            }
            header_y + hdr_h
        } else {
            0.0
        };
        let start_y = page.margin.top.max(header_bottom);
        let content_height = page.size.height - start_y - page.margin.bottom;

        // Multi-column layout: compute column X positions and widths
        // COM-confirmed: col_x = margin + Σ(prev_width + prev_spacing)
        let num_columns = page.columns.as_ref().map(|c| c.num.max(1) as usize).unwrap_or(1);
        let mut col_x_positions: Vec<f32> = Vec::with_capacity(num_columns);
        let mut col_widths: Vec<f32> = Vec::with_capacity(num_columns);

        if num_columns > 1 {
            if let Some(ref cols) = page.columns {
                if !cols.columns.is_empty() {
                    // Unequal width columns: use explicit definitions
                    let mut x = page.margin.left;
                    for col_def in &cols.columns {
                        col_x_positions.push(x);
                        col_widths.push(col_def.width);
                        x += col_def.width + col_def.space.unwrap_or(0.0);
                    }
                } else {
                    // Equal width columns
                    let spacing = cols.space.unwrap_or(36.0); // default 36pt
                    let col_w = (total_content_width - spacing * (num_columns - 1) as f32) / num_columns as f32;
                    let mut x = page.margin.left;
                    for _ in 0..num_columns {
                        col_x_positions.push(x);
                        col_widths.push(col_w);
                        x += col_w + spacing;
                    }
                }
            }
        }
        if col_x_positions.is_empty() {
            col_x_positions.push(page.margin.left);
            col_widths.push(total_content_width);
        }

        let mut current_column: usize = 0;
        let mut start_x = col_x_positions[0];
        let mut content_width = col_widths[0];

        let grid_pitch = page.grid_line_pitch;
        let mut pages: Vec<LayoutPage> = Vec::new();
        let mut elements: Vec<LayoutElement> = Vec::new();
        let mut cursor_y = start_y;
        let mut prev_para_style_id: Option<String> = None;
        let mut prev_contextual_spacing: bool = false;
        let mut prev_space_after: f32 = 0.0;
        // Track Y position and layout page index for each block (for paragraph-relative TextBox positioning)
        let mut block_y_positions: Vec<f32> = Vec::with_capacity(page.blocks.len());
        let mut block_page_indices: Vec<usize> = Vec::with_capacity(page.blocks.len());
        let mut current_page_idx: usize = 0;


        for (block_idx, block) in page.blocks.iter().enumerate() {
            // wrapTopAndBottom: for inline TABLE blocks, push below overlapping TextBoxes
            // Skip for floating tables (tblpPr) as they have explicit positioning
            let is_floating_table = matches!(block, Block::Table(t) if t.style.position.is_some());
            if matches!(block, Block::Table(_)) && !is_floating_table {
                for tb in &page.text_boxes {
                    // Skip wrapNone text boxes (they don't affect text flow)
                    if tb.wrap_type == Some(crate::ir::WrapType::None) {
                        continue;
                    }
                    if tb.anchor_block_index < block_idx {
                        if let Some(ref pos) = tb.position {
                            let anchor_y = block_y_positions.get(tb.anchor_block_index).copied().unwrap_or(0.0);
                            let tb_top = match pos.v_relative.as_deref() {
                                Some("paragraph") | Some("line") => anchor_y + pos.y,
                                Some("margin") => page.margin.top + pos.y,
                                Some("page") => pos.y,
                                _ => anchor_y + pos.y,
                            };
                            let tb_bottom = tb_top + tb.height;
                            if cursor_y >= tb_top && cursor_y < tb_bottom {
                                cursor_y = tb_bottom;
                            }
                        }
                    }
                }
            }
            block_y_positions.push(cursor_y);
            block_page_indices.push(current_page_idx);
            match block {
                Block::Paragraph(para) => {
                    // pageBreakBefore: force a new page (not just next column)
                    if para.style.page_break_before && !elements.is_empty() {
                        pages.push(LayoutPage {
                            width: page.size.width,
                            height: page.size.height,
                            elements: std::mem::take(&mut elements),
                        });
                        cursor_y = start_y;
                        current_column = 0;
                        start_x = col_x_positions[0];
                        content_width = col_widths[0];
                        current_page_idx += 1;
                        *block_page_indices.last_mut().unwrap() = current_page_idx;
                        *block_y_positions.last_mut().unwrap() = cursor_y;
                    }

                    // keepLines: if doesn't fit, advance column or page
                    if para.style.keep_lines && !elements.is_empty() {
                        let est_h = self.estimate_para_height(para, content_width, grid_pitch, None);
                        let remaining = (start_y + content_height) - cursor_y;
                        if est_h > remaining && est_h <= content_height {
                            if num_columns > 1 && current_column + 1 < num_columns {
                                current_column += 1;
                                start_x = col_x_positions[current_column];
                                content_width = col_widths[current_column];
                                cursor_y = start_y;
                            } else {
                                pages.push(LayoutPage {
                                    width: page.size.width,
                                    height: page.size.height,
                                    elements: std::mem::take(&mut elements),
                                });
                                cursor_y = start_y;
                                current_column = 0;
                                start_x = col_x_positions[0];
                                content_width = col_widths[0];
                                current_page_idx += 1;
                            }
                            *block_page_indices.last_mut().unwrap() = current_page_idx;
                            *block_y_positions.last_mut().unwrap() = cursor_y;
                        }
                    }

                    // keepNext: advance column or page if pair doesn't fit
                    if para.style.keep_next && !elements.is_empty() {
                        if let Some(Block::Paragraph(next_para)) = page.blocks.get(block_idx + 1) {
                            let this_h = self.estimate_para_height(para, content_width, grid_pitch, None);
                            let next_h = self.estimate_para_height(next_para, content_width, grid_pitch, None);
                            let remaining = (start_y + content_height) - cursor_y;
                            if this_h + next_h > remaining && this_h + next_h <= content_height {
                                if num_columns > 1 && current_column + 1 < num_columns {
                                    current_column += 1;
                                    start_x = col_x_positions[current_column];
                                    content_width = col_widths[current_column];
                                    cursor_y = start_y;
                                } else {
                                    pages.push(LayoutPage {
                                        width: page.size.width,
                                        height: page.size.height,
                                        elements: std::mem::take(&mut elements),
                                    });
                                    cursor_y = start_y;
                                    current_column = 0;
                                    start_x = col_x_positions[0];
                                    content_width = col_widths[0];
                                    current_page_idx += 1;
                                }
                                *block_page_indices.last_mut().unwrap() = current_page_idx;
                                *block_y_positions.last_mut().unwrap() = cursor_y;
                            }
                        }
                    }

                    // Multi-column pre-check: advance column if paragraph won't fit
                    if num_columns > 1 {
                        let est_h = self.estimate_para_height(para, content_width, grid_pitch, None);
                        let remaining = (start_y + content_height) - cursor_y;
                        if est_h > remaining && est_h <= content_height {
                            if current_column + 1 < num_columns {
                                current_column += 1;
                                start_x = col_x_positions[current_column];
                                content_width = col_widths[current_column];
                                cursor_y = start_y;
                            } else {
                                pages.push(LayoutPage {
                                    width: page.size.width,
                                    height: page.size.height,
                                    elements: std::mem::take(&mut elements),
                                });
                                cursor_y = start_y;
                                current_column = 0;
                                start_x = col_x_positions[0];
                                content_width = col_widths[0];
                                current_page_idx += 1;
                            }
                            *block_page_indices.last_mut().unwrap() = current_page_idx;
                            *block_y_positions.last_mut().unwrap() = cursor_y;
                        }
                    }

                    let pages_before = pages.len();
                    let (para_elements, sa) = self.layout_paragraph(
                        para,
                        start_x,
                        &mut cursor_y,
                        content_width,
                        content_height,
                        start_y,
                        page,
                        &mut pages,
                        &mut elements,
                        grid_pitch,
                        prev_para_style_id.as_deref(), prev_contextual_spacing, false,
                        prev_space_after,
                        Some(block_idx),
                    );
                    prev_space_after = sa;
                    elements.extend(para_elements);
                    // Track page/column breaks that happened inside layout_paragraph
                    let pages_added = pages.len() - pages_before;
                    if pages_added > 0 {
                        // Multi-column: a "page break" inside layout_paragraph may actually
                        // be a column break. Check if we can advance to the next column.
                        if num_columns > 1 && current_column < num_columns - 1 {
                            // Move to next column instead of creating a new page.
                            // The page was already pushed by layout_paragraph — undo it
                            // by popping and re-merging elements.
                            // Actually, layout_paragraph already pushed the page.
                            // We update column state for subsequent blocks.
                            current_column += 1;
                            start_x = col_x_positions[current_column];
                            content_width = col_widths[current_column];
                            // cursor_y was already reset to start_y by layout_paragraph
                        } else if num_columns > 1 {
                            // All columns exhausted: reset to column 0 for new page
                            current_column = 0;
                            start_x = col_x_positions[0];
                            content_width = col_widths[0];
                        }
                        current_page_idx += pages_added;
                    }
                    prev_para_style_id = para.style.style_id.clone();
                    prev_contextual_spacing = para.style.contextual_spacing;
                }
                Block::Table(table) => {
                    // COM-confirmed: prev paragraph's space_after is always added before table
                    cursor_y += prev_space_after;
                    prev_space_after = 0.0;

                    let is_floating = table.style.position.is_some();
                    let saved_cursor_y = cursor_y;

                    // Floating table (tblpPr): position relative to anchor
                    if let Some(ref pos) = table.style.position {
                        cursor_y = match pos.v_anchor.as_deref() {
                            Some("page") => pos.y,
                            Some("margin") => start_y + pos.y,
                            _ => cursor_y + pos.y, // "text": offset from anchor para bottom
                        };
                    }
                    let pages_before = pages.len();
                    let table_elements = self.layout_table(
                        table,
                        start_x,
                        &mut cursor_y,
                        content_width,
                        grid_pitch,
                        page.grid_char_pitch,
                        start_y,
                        content_height,
                        page.size.width,
                        page.size.height,
                        &mut pages,
                        &mut elements,
                    );
                    elements.extend(table_elements);

                    if is_floating {
                        // Floating tables don't advance the text flow
                        cursor_y = saved_cursor_y;
                    } else {
                        let pages_added = pages.len() - pages_before;
                        if pages_added > 0 {
                            current_page_idx += pages_added;
                            *block_page_indices.last_mut().unwrap() = current_page_idx;
                            *block_y_positions.last_mut().unwrap() = cursor_y;
                            if num_columns > 1 {
                                current_column = 0;
                                start_x = col_x_positions[0];
                                content_width = col_widths[0];
                            }
                        }
                    }
                    prev_para_style_id = None;
                    prev_space_after = 0.0;
                }
                Block::Image(img) => {
                    if cursor_y + img.height > start_y + content_height {
                        if num_columns > 1 && current_column + 1 < num_columns {
                            current_column += 1;
                            start_x = col_x_positions[current_column];
                            content_width = col_widths[current_column];
                            cursor_y = start_y;
                        } else {
                            pages.push(LayoutPage {
                                width: page.size.width,
                                height: page.size.height,
                                elements: std::mem::take(&mut elements),
                            });
                            cursor_y = start_y;
                            current_column = 0;
                            start_x = col_x_positions[0];
                            content_width = col_widths[0];
                            current_page_idx += 1;
                        }
                        *block_page_indices.last_mut().unwrap() = current_page_idx;
                        *block_y_positions.last_mut().unwrap() = cursor_y;
                    }
                    elements.push(LayoutElement::new(start_x, cursor_y, img.width, img.height, LayoutContent::Image {
                            data: img.data.clone(),
                            content_type: img.content_type.clone(),
                    }));
                    cursor_y += img.height;
                    prev_para_style_id = None;
                }
                Block::UnsupportedElement(_) => {
                    // Skip unsupported elements in layout
                }
            }
        }

        // Final page
        pages.push(LayoutPage {
            width: page.size.width,
            height: page.size.height,
            elements,
        });

        // Layout text boxes and add to the correct layout page
        // The current_page_idx tracking tells us which layout page each anchor block ended up on
        for text_box in &page.text_boxes {
            let target_page = block_page_indices
                .get(text_box.anchor_block_index)
                .copied()
                .unwrap_or(0);
            let tb_elements = self.layout_text_box(text_box, page, &block_y_positions);
            if let Some(lp) = pages.get_mut(target_page) {
                lp.elements.extend(tb_elements);
            }
        }

        // Layout floating images and add to the correct layout page
        for img in &page.floating_images {
            if let Some(ref pos) = img.position {
                let (abs_x, abs_y) = self.resolve_floating_image_position(img, page, &block_y_positions);
                // Use the same page as the anchor block
                let target_page = block_page_indices
                    .get(img.anchor_block_index)
                    .copied()
                    .unwrap_or(0);
                let el = LayoutElement::new(abs_x, abs_y, img.width, img.height, LayoutContent::Image {
                        data: img.data.clone(),
                        content_type: img.content_type.clone(),
                });
                if let Some(lp) = pages.get_mut(target_page) {
                    lp.elements.push(el);
                } else if let Some(lp) = pages.last_mut() {
                    lp.elements.push(el);
                }
            } else {
                // No position info — treat as inline at end of last page
                if let Some(lp) = pages.last_mut() {
                    lp.elements.push(LayoutElement::new(start_x, 0.0, img.width, img.height, LayoutContent::Image {
                            data: img.data.clone(),
                            content_type: img.content_type.clone(),
                    }));
                }
            }
        }

        // Layout header/footer on each layout page
        // Header y = headerDistance (from page top edge), default 36pt (0.5in)
        // Footer y = pageHeight - footerDistance - footerContentHeight
        let header_y = page.header_distance.unwrap_or(36.0);
        let footer_dist = page.footer_distance.unwrap_or(36.0);
        let hdr_x = page.margin.left;
        let hdr_width = content_width;
        for (page_idx, lp) in pages.iter_mut().enumerate() {
            if !page.header.is_empty() {
                let mut cy = header_y;
                for block in &page.header {
                    if let Block::Paragraph(para) = block {
                        let (hdr_elements, _) = self.layout_paragraph(
                            para, hdr_x, &mut cy, hdr_width, page.size.height,
                            header_y, page, &mut Vec::new(), &mut Vec::new(),
                            grid_pitch, None, false,
                            false, 0.0, None,
                        );
                        lp.elements.extend(hdr_elements);
                    }
                }
            }
            if !page.footer.is_empty() {
                // Estimate footer content height first
                let mut footer_h: f32 = 0.0;
                for block in &page.footer {
                    if let Block::Paragraph(para) = block {
                        footer_h += self.estimate_para_height(para, hdr_width, grid_pitch, None);
                    }
                }
                let footer_top = page.size.height - footer_dist - footer_h;
                let mut cy = footer_top;
                for block in &page.footer {
                    if let Block::Paragraph(para) = block {
                        let (ftr_elements, _) = self.layout_paragraph(
                            para, hdr_x, &mut cy, hdr_width, page.size.height,
                            footer_top, page, &mut Vec::new(), &mut Vec::new(),
                            grid_pitch, None, false,
                            false, 0.0, None,
                        );
                        lp.elements.extend(ftr_elements);
                    }
                }
            }

            // Render shapes (e.g. bracketPair) positioned relative to anchor paragraph
            for shape in &page.shapes {
                if let Some(ref pos) = shape.position {
                    // Get anchor paragraph's Y position and page index
                    let anchor_y = block_y_positions.get(shape.anchor_block_index).copied().unwrap_or(start_y);
                    let anchor_page = block_page_indices.get(shape.anchor_block_index).copied().unwrap_or(0);

                    // Only render on the correct page
                    if anchor_page == page_idx {
                        // h_relative="column": x = margin_left + offset
                        // v_relative="paragraph": y = anchor_paragraph_y + offset
                        let sx = start_x + pos.x;
                        let sy = anchor_y + pos.y;
                        lp.elements.push(LayoutElement::new(sx, sy, shape.width, shape.height, LayoutContent::PresetShape {
                                shape_type: shape.shape_type.clone(),
                                stroke_color: shape.stroke_color.clone(),
                                stroke_width: shape.stroke_width.unwrap_or(0.75),
                        }));
                    }
                }
            }
        }

        pages
    }

    /// Resolve absolute (x, y) position for a text box based on its anchor references.
    fn resolve_textbox_position(&self, text_box: &TextBox, page: &Page, block_y_positions: &[f32]) -> (f32, f32) {
        let pos = match &text_box.position {
            Some(p) => p,
            None => return (page.margin.left, page.margin.top),
        };

        let content_width = page.size.width - page.margin.left - page.margin.right;

        // Horizontal: alignment takes precedence over offset
        let abs_x = if let Some(ref align) = pos.h_align {
            let ref_left;
            let ref_width;
            match pos.h_relative.as_deref() {
                Some("page") => { ref_left = 0.0; ref_width = page.size.width; }
                Some("margin") | Some("column") | _ => { ref_left = page.margin.left; ref_width = content_width; }
            }
            match align.as_str() {
                "left" => ref_left,
                "center" => ref_left + (ref_width - text_box.width) / 2.0,
                "right" => ref_left + ref_width - text_box.width,
                _ => ref_left,
            }
        } else {
            match pos.h_relative.as_deref() {
                Some("page") => pos.x,
                Some("margin") | Some("column") | Some("character") => page.margin.left + pos.x,
                Some("leftMarginArea") => pos.x,
                Some("rightMarginArea") => (page.size.width - page.margin.right) + pos.x,
                _ | None => page.margin.left + pos.x,
            }
        };

        // Vertical: paragraph-relative uses anchor block Y position
        let abs_y = if let Some(ref align) = pos.v_align {
            let ref_top;
            let ref_height;
            match pos.v_relative.as_deref() {
                Some("page") => { ref_top = 0.0; ref_height = page.size.height; }
                Some("margin") | _ => { ref_top = page.margin.top; ref_height = page.size.height - page.margin.top - page.margin.bottom; }
            }
            match align.as_str() {
                "top" => ref_top,
                "center" => ref_top + (ref_height - text_box.height) / 2.0,
                "bottom" => ref_top + ref_height - text_box.height,
                _ => ref_top,
            }
        } else {
            match pos.v_relative.as_deref() {
                Some("page") => pos.y,
                Some("paragraph") | Some("line") => {
                    let anchor_y = block_y_positions
                        .get(text_box.anchor_block_index)
                        .copied()
                        .unwrap_or(page.margin.top);
                    anchor_y + pos.y
                }
                Some("margin") => page.margin.top + pos.y,
                Some("topMarginArea") => pos.y,
                Some("bottomMarginArea") => (page.size.height - page.margin.bottom) + pos.y,
                _ | None => page.margin.top + pos.y,
            }
        };

        // Clamp TextBox to page boundaries (prevent overflow beyond page edge)
        let abs_y = if abs_y + text_box.height > page.size.height {
            (page.size.height - text_box.height).max(0.0)
        } else {
            abs_y
        };
        let abs_x = if abs_x + text_box.width > page.size.width {
            (page.size.width - text_box.width).max(0.0)
        } else {
            abs_x
        };

        (abs_x, abs_y)
    }

    /// Resolve absolute (x, y) position for a floating image.
    fn resolve_floating_image_position(&self, img: &Image, page: &Page, block_y_positions: &[f32]) -> (f32, f32) {
        let pos = match &img.position {
            Some(p) => p,
            None => return (page.margin.left, page.margin.top),
        };

        let content_width = page.size.width - page.margin.left - page.margin.right;

        let abs_x = if let Some(ref align) = pos.h_align {
            let (ref_left, ref_width) = match pos.h_relative.as_deref() {
                Some("page") => (0.0, page.size.width),
                Some("leftMargin") => (0.0, page.margin.left),
                Some("rightMargin") => (page.size.width - page.margin.right, page.margin.right),
                Some("margin") | Some("column") | _ => (page.margin.left, content_width),
            };
            match align.as_str() {
                "left" => ref_left,
                "center" => ref_left + (ref_width - img.width) / 2.0,
                "right" => ref_left + ref_width - img.width,
                _ => ref_left,
            }
        } else {
            match pos.h_relative.as_deref() {
                Some("page") => pos.x,
                Some("margin") | Some("column") => page.margin.left + pos.x,
                Some("leftMargin") | Some("leftMarginArea") => pos.x,
                Some("rightMargin") | Some("rightMarginArea") => (page.size.width - page.margin.right) + pos.x,
                _ => page.margin.left + pos.x,
            }
        };

        let abs_y = if let Some(ref align) = pos.v_align {
            let (ref_top, ref_height) = match pos.v_relative.as_deref() {
                Some("page") => (0.0, page.size.height),
                _ => (page.margin.top, page.size.height - page.margin.top - page.margin.bottom),
            };
            match align.as_str() {
                "top" => ref_top,
                "center" => ref_top + (ref_height - img.height) / 2.0,
                "bottom" => ref_top + ref_height - img.height,
                _ => ref_top,
            }
        } else {
            match pos.v_relative.as_deref() {
                Some("page") => pos.y,
                Some("paragraph") | Some("line") => {
                    let anchor_y = block_y_positions
                        .get(img.anchor_block_index)
                        .copied()
                        .unwrap_or(page.margin.top);
                    anchor_y + pos.y
                }
                Some("margin") => page.margin.top + pos.y,
                _ => page.margin.top + pos.y,
            }
        };

        // Clamp to page boundaries (floating images can extend into margins)
        let abs_y = if abs_y + img.height > page.size.height {
            (page.size.height - img.height).max(0.0)
        } else {
            abs_y
        };

        (abs_x, abs_y)
    }

    /// Layout a single text box: background, borders, and inner content.
    fn layout_text_box(&self, text_box: &TextBox, page: &Page, block_y_positions: &[f32]) -> Vec<LayoutElement> {
        let mut elements = Vec::new();

        // 1. Calculate absolute position
        let (abs_x, abs_y) = self.resolve_textbox_position(text_box, page, block_y_positions);

        // 2. Background fill + border as a single BoxRect (supports corner radius)
        let has_fill = text_box.fill.is_some();
        let has_border = text_box.border;
        if has_fill || has_border {
            let fill_hex = text_box.fill.as_ref().map(|f| {
                if f.starts_with('#') { f.clone() } else { format!("#{}", f) }
            });
            let cr = text_box.corner_radius.unwrap_or(0.0);
            elements.push(LayoutElement::new(abs_x, abs_y, text_box.width, text_box.height, LayoutContent::BoxRect {
                    fill: fill_hex,
                    stroke_color: if has_border {
                        text_box.stroke_color.as_ref()
                            .map(|c| if c.starts_with('#') { c.clone() } else { format!("#{}", c) })
                            .or_else(|| Some("#000000".to_string()))
                    } else { None },
                    stroke_width: if has_border { text_box.stroke_width.unwrap_or(1.0) } else { 0.0 },
                    corner_radius: cr,
            }));
        }

        // 3. Clip region — all TextBox content is clipped to the box boundary
        elements.push(LayoutElement::new(abs_x, abs_y, text_box.width, text_box.height, LayoutContent::ClipStart));

        // 4. Content layout within text box
        // Word default inset: L/R = 7.2pt (0.1in = 91440 EMU), T/B = 3.6pt (0.05in = 45720 EMU)
        let inset_l = text_box.inset_left.unwrap_or(7.2);
        let inset_r = text_box.inset_right.unwrap_or(7.2);
        let inset_t = text_box.inset_top.unwrap_or(3.6);
        let inset_b = text_box.inset_bottom.unwrap_or(3.6);
        let inner_x = abs_x + inset_l;
        let inner_width = (text_box.width - inset_l - inset_r).max(0.0);
        let inner_height = (text_box.height - inset_t - inset_b).max(0.0);
        let mut cursor_y = abs_y + inset_t;

        // We layout content inside the text box without page-breaking.
        // Use dummy page/elements vecs since we don't want page breaks inside text boxes.
        let mut dummy_pages: Vec<LayoutPage> = Vec::new();
        let mut dummy_elements: Vec<LayoutElement> = Vec::new();

        for block in &text_box.blocks {
            // Stop if we've exceeded the text box bounds
            if cursor_y > abs_y + text_box.height - inset_b {
                break;
            }

            match block {
                Block::Paragraph(para) => {
                    let clip_bottom = abs_y + text_box.height;
                    let (para_elements, _) = self.layout_paragraph(
                        para,
                        inner_x,
                        &mut cursor_y,
                        inner_width,
                        inner_height,
                        abs_y + inset_t,
                        page,
                        &mut dummy_pages,
                        &mut dummy_elements,
                        // TextBox grid snap: enabled for "lines" grid, disabled for "linesAndChars"
                        if page.grid_char_pitch.is_some() { None } else { page.grid_line_pitch },
                        None, false, // no prev style/contextual tracking
                        true, // in_textbox: suppress CJK compression
                        0.0, None,
                    );
                    // Word behavior: TextBox overflow text is not rendered.
                    // Filter: (1) Y overflow, (2) in dark-filled TextBox, skip text with no explicit color.
                    // Word PDF omits runs without color attribute inside colored TextBoxes —
                    // these are overflow text that would be black-on-dark and shouldn't be visible.
                    // Only apply to dark fills (not white/light backgrounds where black text is normal).
                    let has_dark_fill = text_box.fill.as_ref().map_or(false, |f| {
                        let hex = f.trim_start_matches('#');
                        if hex.len() >= 6 {
                            let r = u8::from_str_radix(&hex[0..2], 16).unwrap_or(255);
                            let g = u8::from_str_radix(&hex[2..4], 16).unwrap_or(255);
                            let b = u8::from_str_radix(&hex[4..6], 16).unwrap_or(255);
                            (r as u16 + g as u16 + b as u16) < 600
                        } else {
                            false
                        }
                    });
                    let accept_and_fix_color = |pe: &mut LayoutElement| -> bool {
                        if pe.y + pe.height > clip_bottom { return false; }
                        // Fix text color contrast in dark-filled TextBoxes:
                        // Theme colors not resolved → color=None/#000000 on dark background.
                        // Replace with white for readability.
                        if has_dark_fill {
                            if let LayoutContent::Text { ref mut color, .. } = pe.content {
                                match color.as_deref() {
                                    None | Some("#000000") | Some("000000") => {
                                        *color = Some("#FFFFFF".to_string());
                                    }
                                    _ => {}
                                }
                            }
                        }
                        true
                    };
                    for mut pe in para_elements {
                        if accept_and_fix_color(&mut pe) { elements.push(pe); }
                    }
                    for mut de in dummy_elements.drain(..) {
                        if accept_and_fix_color(&mut de) { elements.push(de); }
                    }
                }
                Block::Table(table) => {
                    // TextBox tables don't paginate — use large content_height
                    let mut tb_pages = Vec::new();
                    let mut tb_elems = Vec::new();
                    let table_elements = self.layout_table(
                        table,
                        inner_x,
                        &mut cursor_y,
                        inner_width,
                        None,
                        None,
                        0.0, 99999.0, 0.0, 99999.0,
                        &mut tb_pages, &mut tb_elems,
                    );
                    elements.extend(tb_elems);
                    elements.extend(table_elements);
                }
                Block::Image(img) => {
                    elements.push(LayoutElement::new(inner_x, cursor_y, img.width.min(inner_width), img.height, LayoutContent::Image {
                            data: img.data.clone(),
                            content_type: img.content_type.clone(),
                    }));
                    cursor_y += img.height;
                }
                Block::UnsupportedElement(_) => {}
            }
        }

        // Use specified height (no autoFit by default in Word).
        // Only shrink if content is smaller AND autoFit is explicitly enabled.
        let actual_height = text_box.height;

        // Patch background fill and clip elements with actual height
        for el in elements.iter_mut() {
            if el.x == abs_x && el.y == abs_y && el.height == text_box.height {
                el.height = actual_height;
            }
        }

        // End clip region
        elements.push(LayoutElement::new(abs_x, abs_y, text_box.width, actual_height, LayoutContent::ClipEnd));

        elements
    }

    #[allow(clippy::too_many_arguments)]
    fn layout_paragraph(
        &self,
        para: &Paragraph,
        start_x: f32,
        cursor_y: &mut f32,
        content_width: f32,
        content_height: f32,
        page_top: f32,
        page: &Page,
        pages: &mut Vec<LayoutPage>,
        current_elements: &mut Vec<LayoutElement>,
        grid_pitch: Option<f32>,
        prev_style_id: Option<&str>,
        prev_contextual_spacing: bool,
        #[allow(unused)] in_textbox: bool,
        prev_space_after: f32,
        body_para_index: Option<usize>,
    ) -> (Vec<LayoutElement>, f32) {
        let mut elements = Vec::new();

        // Apply paragraph spacing (space_before).
        // Word uses max(prev_space_after, space_before) — spacing collapse.
        let mut space_before = if let (Some(bl), Some(pitch)) = (para.style.before_lines, grid_pitch) {
            let raw = bl / 100.0 * pitch;
            if para.style.snap_to_grid && pitch > 0.0 {
                ((raw / pitch) + 0.5).floor() * pitch
            } else {
                raw
            }
        } else {
            para.style.space_before.unwrap_or(0.0)
        };

        // Spacing collapse: max(prev_sa, cur_sb) instead of prev_sa + cur_sb.
        // prev_space_after was NOT added to cursor_y by the caller.
        let collapsed_spacing = space_before.max(prev_space_after);

        // Contextual spacing: suppress spacing when EITHER paragraph has
        // contextualSpacing=true AND they share the same style (COM-confirmed).
        let mut effective_spacing = collapsed_spacing;
        if para.style.contextual_spacing || prev_contextual_spacing {
            if let (Some(cur_id), Some(prev_id)) = (para.style.style_id.as_deref(), prev_style_id) {
                if cur_id == prev_id {
                    effective_spacing = 0.0;
                }
            }
        }

        // Suppress space_before at the top of a page (page 2+).
        // COM-confirmed: page 1 preserves space_before (H1 sb=24 → y=96=72+24).
        // Page 2+ suppresses it.
        let is_page_2_plus = !pages.is_empty() || !current_elements.is_empty();
        if (*cursor_y - page_top).abs() < 0.01 && is_page_2_plus {
            effective_spacing = 0.0;
        }

        *cursor_y += effective_spacing;

        let indent_left = para.style.indent_left.unwrap_or(0.0);
        let indent_right = para.style.indent_right.unwrap_or(0.0);
        let first_line_indent = para.style.indent_first_line.unwrap_or(0.0);
        // COM-confirmed (2026-04-03): charGrid (linesAndChars) ignores paragraph indents
        // for line-break purposes. Text starts at margin and charsLine determines wrapping.
        // data_guideline: indent=12pt but x0=71 (margin), 38ch/line (=charsLine+1 kinsoku).
        let effective_char_pitch = if in_textbox { None } else { page.grid_char_pitch };
        let available_width = if effective_char_pitch.is_some() {
            content_width  // charGrid: ignore indents for wrapping
        } else {
            content_width - indent_left - indent_right
        };

        // Render list marker if present
        if let Some(ref marker) = para.style.list_marker {
            let default_style = RunStyle::default();
            let marker_style = para.runs.first().map(|r| &r.style).unwrap_or(&default_style);
            let mut marker_font_size = self.resolve_font_size(marker_style, &para.style);
            // Symbol font bullets (•/●) have large glyphs relative to em-square.
            // No font size adjustment needed — use the paragraph's font size directly.
            let marker_metrics = self.metrics_for(marker_style, &para.style);
            let marker_width: f32 = marker
                .chars()
                .map(|c| self.registry.char_width_pt_with_fallback(c, marker_font_size, marker_metrics))
                .sum();
            let list_indent = para.style.list_indent.unwrap_or(18.0);
            let marker_x = start_x + indent_left - list_indent;
            let line_height = self.line_height(marker_font_size, para.style.line_spacing, para.style.line_spacing_rule.as_deref(), marker_metrics, para.style.snap_to_grid, grid_pitch);

            // Determine marker text including suffix
            let suff = para.style.list_suff.as_deref().unwrap_or("tab");
            let marker_text = match suff {
                "space" => format!("{} ", marker),
                "nothing" => marker.clone(),
                // "tab" — marker text alone; tab stop handled by indent_left
                _ => {
                    // For tab suffix: if there's a tab_stop defined, use it to
                    // adjust text start position via indent_left. The marker sits
                    // at marker_x and text starts at indent_left (which should
                    // align with the tab stop).
                    marker.clone()
                }
            };

            // Page break check for marker
            if *cursor_y + line_height > page_top + content_height {
                pages.push(LayoutPage {
                    width: page.size.width,
                    height: page.size.height,
                    elements: std::mem::take(current_elements),
                });
                current_elements.extend(std::mem::take(&mut elements));
                elements = std::mem::take(current_elements);
                *cursor_y = page_top;
            }

            // Bullet markers are scaled up (2x) so adjust Y to align with text center
            let marker_y_offset = if marker.contains('\u{2022}') || marker.contains('\u{25CF}') {
                -marker_font_size * 0.15  // shift up slightly
            } else {
                0.0
            };
            elements.push(LayoutElement::new(marker_x, *cursor_y + marker_y_offset, marker_width, line_height, LayoutContent::Text {
                    text: marker_text,
                    font_size: marker_font_size,
                    font_family: None,
                    bold: false,
                    italic: false,
                    underline: false,
                    underline_style: None,
                    strikethrough: false,
                    color: None,
                    highlight: None,
                    field_type: None,
                    character_spacing: 0.0,
            }));
        }

        // Collect all text fragments with their styles, field types, and source indices
        let fragments: Vec<(&str, &RunStyle, Option<FieldType>, usize, usize)> = para
            .runs
            .iter()
            .enumerate()
            .map(|(i, r)| (r.text.as_str(), &r.style, r.field_type, i, 0usize))
            .collect();

        // Resolve font size for line breaking
        let default_style = RunStyle::default();
        let para_font_size = self.resolve_font_size(
            para.runs.first().map(|r| &r.style).unwrap_or(&default_style),
            &para.style,
        );

        // Line-break the text
        let effective_first_indent = if effective_char_pitch.is_some() { 0.0 } else { first_line_indent };
        let lines = self.break_into_lines(&fragments, available_width, effective_first_indent, &para.style, effective_char_pitch);

        // Widow/orphan control: pre-compute line heights for lookahead
        let line_heights: Vec<f32> = lines.iter().map(|line| {
            self.line_height_for_line(line, &para.style, para_font_size, para.style.snap_to_grid, grid_pitch)
        }).collect();

        for (line_idx, line) in lines.iter().enumerate() {
            let _first_style = line.fragments.first().map(|f| &f.style).unwrap_or(&default_style);
            let line_height = line_heights[line_idx];

            // Page break check with widow/orphan control
            // TextBox content: no page breaks, no widow/orphan. Overflow is clipped.
            let needs_page_break = if in_textbox { false } else {
                *cursor_y + line_height > page_top + content_height
            };

            // Widow/orphan: if this is line 0 (orphan) and there are 2+ lines,
            // check if only 1 line would fit on this page — if so, push the
            // entire paragraph to the next page.
            let widow_orphan_break = if !in_textbox && para.style.widow_control && lines.len() >= 2 {
                if line_idx == 0 && !needs_page_break {
                    // Orphan: check if the next line would overflow — that would leave
                    // only 1 line on this page. Push entire paragraph to next page.
                    let next_h = line_heights.get(1).copied().unwrap_or(0.0);
                    *cursor_y + line_height + next_h > page_top + content_height
                        && !current_elements.is_empty()
                } else if line_idx == lines.len() - 2 && !needs_page_break {
                    // Widow: if the last line would overflow to the next page alone,
                    // break BEFORE this line so at least 2 lines go to the next page.
                    let next_h = line_heights.get(line_idx + 1).copied().unwrap_or(0.0);
                    *cursor_y + line_height + next_h > page_top + content_height
                } else {
                    false
                }
            } else {
                false
            };

            if widow_orphan_break {
                // Push current page and move entire paragraph so far to next page
                pages.push(LayoutPage {
                    width: page.size.width,
                    height: page.size.height,
                    elements: std::mem::take(current_elements),
                });
                current_elements.extend(std::mem::take(&mut elements));
                elements = std::mem::take(current_elements);
                *cursor_y = page_top;
            } else if needs_page_break {
                // Mid-paragraph page break: keep already-laid-out lines on current page,
                // only the overflowing line (and subsequent) go to the next page.
                current_elements.extend(std::mem::take(&mut elements));
                pages.push(LayoutPage {
                    width: page.size.width,
                    height: page.size.height,
                    elements: std::mem::take(current_elements),
                });
                *cursor_y = page_top;
            }

            let extra_indent = if line_idx == 0 { first_line_indent } else { 0.0 };
            // For right-aligned text, firstLine indent reduces available width but
            // doesn't shift line_x (text is pushed from the right edge).
            let line_x = if para.alignment == Alignment::Right {
                start_x + indent_left
            } else {
                start_x + indent_left + extra_indent
            };

            // Alignment offset
            let line_text_width: f32 = line.fragments.iter().map(|f| f.width).sum();
            let is_last_line = line_idx == lines.len() - 1;
            let align_offset = match para.alignment {
                Alignment::Left => 0.0,
                Alignment::Center => {
                    // Word GDI: integer pixel division at 96dpi for center alignment
                    let slack_tw = ((available_width - extra_indent - line_text_width) * 20.0).round() as i32;
                    let center_tw = slack_tw / 2; // integer division (truncate)
                    center_tw as f32 / 20.0
                },
                Alignment::Right => available_width - extra_indent - line_text_width,
                Alignment::Justify => 0.0,
                // Distribute: when justification applies (multi-fragment lines), offset is 0
                // because slack is distributed across fragments. When justification can't
                // apply (single-fragment line), center the content.
                Alignment::Distribute => {
                    if line.fragments.len() > 1 {
                        0.0
                    } else {
                        let slack = available_width - extra_indent - line_text_width;
                        if slack > 0.0 { slack / 2.0 } else { 0.0 }
                    }
                }
            };

            // Justification (matches Word output, priority order):
            // 1. CJK punctuation compression (full-width -> half-width, 50% savings)
            // 2. Word-space expansion (distribute remaining slack at space characters)
            // Latin text: ONLY expand at word spaces, never between characters.
            // CJK text: compress punctuation first, then expand at inter-character gaps.

            let mut frag_width_adjustments: Vec<f32> = vec![0.0; line.fragments.len()];
            let mut frag_spacing_after: Vec<f32> = vec![0.0; line.fragments.len()];
            let mut justify_char_spacing: f32 = 0.0;

            let should_justify = !in_textbox
                && ((para.alignment == Alignment::Justify && !is_last_line)
                    || para.alignment == Alignment::Distribute);
            if should_justify && line.fragments.len() > 1 {
                let mut slack = available_width - extra_indent - line_text_width;

                // Phase 1: CJK punctuation compression (full-width -> half-width)
                // Only compress when the line overflows (slack < 0).
                // Matches Word output: TextBox content does NOT use punctuation compression.
                if slack < 0.0 && !in_textbox {
                    for (fi, frag) in line.fragments.iter().enumerate() {
                        for ch in frag.text.chars() {
                            if kinsoku::is_cjk_compressible(ch) {
                                let fs = frag.style.font_size.unwrap_or(para_font_size);
                                let fm = self.metrics_for(&frag.style, &para.style);
                                let char_w = self.registry.char_width_pt_with_fallback(ch, fs, fm);
                                let actual = char_w * 0.5;
                                frag_width_adjustments[fi] -= actual;
                                slack += actual; // reclaim freed space
                            }
                        }
                    }
                }

                // Phase 2: Distribute remaining slack at word spaces (only if slack > 0 after compression)
                if slack > 0.0 {
                    // Count ASCII word spaces only — CJK fullwidth spaces (U+3000) are NOT
                    // word boundaries for justify purposes.
                    let space_count = line.fragments.iter()
                        .enumerate()
                        .filter(|(i, f)| {
                            *i < line.fragments.len() - 1
                            && f.text.chars().all(|c| c == ' ')
                            && !f.text.is_empty()
                        })
                        .count();

                    if space_count > 0 {
                        let per_space = slack / space_count as f32;
                        for (fi, frag) in line.fragments.iter().enumerate() {
                            if fi < line.fragments.len() - 1
                                && frag.text.chars().all(|c| c == ' ')
                                && !frag.text.is_empty()
                            {
                                frag_spacing_after[fi] += per_space;
                            }
                        }
                    } else {
                        // No word spaces: distribute between CJK characters.
                        // Use character_spacing on each fragment so Canvas/PDF renderers
                        // apply per-character gap (not just fragment-level gap).
                        let total_chars: usize = line.fragments.iter()
                            .map(|f| f.text.chars().count())
                            .sum();
                        let has_cjk = line.fragments.iter()
                            .any(|f| f.text.chars().any(|c| kinsoku::is_cjk(c)));
                        if has_cjk && total_chars > 1 {
                            let char_gap_count = total_chars - 1;
                            let per_char_gap = slack / char_gap_count as f32;
                            // Distribute: fragment-boundary gaps via frag_spacing_after,
                            // internal gaps via frag_width_adjustments (for layout width),
                            // AND set justify_char_spacing for renderer to apply letterSpacing.
                            for fi in 0..line.fragments.len() {
                                let frag_chars = line.fragments[fi].text.chars().count();
                                if frag_chars > 1 {
                                    frag_width_adjustments[fi] += per_char_gap * (frag_chars - 1) as f32;
                                }
                                if fi < line.fragments.len() - 1 {
                                    frag_spacing_after[fi] += per_char_gap;
                                }
                            }
                            // Store per_char_gap for use in LayoutElement character_spacing
                            justify_char_spacing = per_char_gap;
                        }
                        // Pure Latin with no spaces: do NOT add inter-character spacing
                    }
                }
            }

            let mut x = line_x + align_offset;

            // Matches Word output: exact/atLeast line spacing places text at BOTTOM of line box.
            // Extra space goes above text (ascent increased, descent unchanged).
            let text_y_off = self.text_y_offset_for_line(line, &para.style, para_font_size, line_height, grid_pitch);

            // Compute max ascent across all fragments for baseline alignment.
            // All fragments in a line share the same baseline (matches Word output).
            let line_max_ascent: f32 = if line.fragments.is_empty() {
                // COM-confirmed: empty lines use paragraph font, not document default
                self.metrics_for(&RunStyle::default(), &para.style).word_ascent_pt(para_font_size)
            } else {
                line.fragments.iter().map(|f| {
                    let fs = f.style.font_size.unwrap_or(para_font_size);
                    self.metrics_for_text(&f.text, &f.style, &para.style).word_ascent_pt(fs)
                }).fold(0.0_f32, f32::max)
            };

            for (frag_idx, frag) in line.fragments.iter().enumerate() {
                let resolved_font_size = frag.style.font_size.unwrap_or(para_font_size);
                let resolved_bold = self.resolve_bold(&frag.style, &para.style);
                let adjusted_width = frag.width + frag_width_adjustments[frag_idx];

                // Per-fragment baseline alignment: shift fragments with smaller ascent
                // so all share the same baseline (y + frag_ascent = cursor_y + text_y_off + line_max_ascent)
                let frag_metrics = self.metrics_for_text(&frag.text, &frag.style, &para.style);
                let frag_ascent = frag_metrics.word_ascent_pt(resolved_font_size);
                let baseline_adjust = line_max_ascent - frag_ascent;

                let mut el = LayoutElement::new(x, *cursor_y + text_y_off + baseline_adjust, adjusted_width, line_height, LayoutContent::Text {
                        text: frag.text.clone(),
                        font_size: resolved_font_size,
                        font_family: self.resolve_font_family_for_text(&frag.text, &frag.style, &para.style)
                            .map(|s| s.to_string()),
                        bold: resolved_bold,
                        italic: self.resolve_italic(&frag.style, &para.style),
                        underline: frag.style.underline,
                        underline_style: frag.style.underline_style.clone(),
                        strikethrough: frag.style.strikethrough,
                        color: self.resolve_color(&frag.style, &para.style).map(|s| s.to_string()),
                        highlight: frag.style.highlight.clone(),
                        field_type: frag.field_type,
                        character_spacing: snap_character_spacing(frag.style.character_spacing.unwrap_or(0.0)) + justify_char_spacing,
                });
                if let Some(pi) = body_para_index {
                    el.paragraph_index = Some(pi);
                    el.run_index = Some(frag.run_index);
                    el.char_offset = Some(frag.char_offset);
                }
                elements.push(el);
                x += adjusted_width + frag_spacing_after[frag_idx];
            }

            *cursor_y += line_height;

            // Handle explicit page/column breaks after this line
            if line.break_type == LineBreakType::PageBreak || line.break_type == LineBreakType::ColumnBreak {
                // Push current page and start a new one
                pages.push(LayoutPage {
                    width: page.size.width,
                    height: page.size.height,
                    elements: std::mem::take(current_elements),
                });
                current_elements.extend(std::mem::take(&mut elements));
                elements = std::mem::take(current_elements);
                *cursor_y = page_top;
            }
        }

        let space_after = if let (Some(al), Some(pitch)) = (para.style.after_lines, grid_pitch) {
            let raw = al / 100.0 * pitch;
            if para.style.snap_to_grid && pitch > 0.0 {
                ((raw / pitch) + 0.5).floor() * pitch
            } else {
                raw
            }
        } else {
            para.style.space_after.unwrap_or(0.0)
        };
        // NOTE: space_after is NOT added to cursor_y here.
        // It will be collapsed with the next paragraph's space_before via max(sa, sb).

        // Paragraph borders (e.g., Title style bottom border)
        if let Some(ref borders) = para.style.borders {
            let para_top = elements.first().map(|e| e.y).unwrap_or(start_x);
            let para_bottom = *cursor_y;
            let border_x = start_x;
            let border_width = content_width;

            if let Some(ref bottom) = borders.bottom {
                let bw = bottom.width;
                let color = bottom.color.clone().unwrap_or_else(|| "000000".to_string());
                let border_y = para_bottom + bottom.space;
                elements.push(LayoutElement::new(border_x, border_y, border_width, bw.max(0.5), LayoutContent::CellShading {
                        color: format!("#{}", color),
                }));
                // Advance cursor past the border (space + width affect next paragraph Y)
                *cursor_y = border_y + bw.max(0.5);
            }
            if let Some(ref top) = borders.top {
                let bw = top.width;
                let color = top.color.clone().unwrap_or_else(|| "000000".to_string());
                let border_y = para_top - top.space - bw;
                elements.push(LayoutElement::new(border_x, border_y, border_width, bw.max(0.5), LayoutContent::CellShading {
                        color: format!("#{}", color),
                }));
            }
            // Left border
            if let Some(ref left) = borders.left {
                let bw = left.width;
                let color = left.color.clone().unwrap_or_else(|| "000000".to_string());
                let bx = border_x - left.space - bw;
                elements.push(LayoutElement::new(bx, para_top, bw.max(0.5), para_bottom - para_top, LayoutContent::CellShading {
                        color: format!("#{}", color),
                }));
            }
            // Right border
            if let Some(ref right) = borders.right {
                let bw = right.width;
                let color = right.color.clone().unwrap_or_else(|| "000000".to_string());
                let bx = border_x + border_width + right.space;
                elements.push(LayoutElement::new(bx, para_top, bw.max(0.5), para_bottom - para_top, LayoutContent::CellShading {
                        color: format!("#{}", color),
                }));
            }
        }

        (elements, space_after)
    }

    fn break_into_lines(
        &self,
        fragments: &[(&str, &RunStyle, Option<FieldType>, usize, usize)],
        available_width: f32,
        first_line_indent: f32,
        para_style: &ParagraphStyle,
        grid_char_pitch: Option<f32>,
    ) -> Vec<Line> {
        // Helper: convert pt to twips for Word-GDI-compatible integer comparison
        let pt_to_tw = |pt: f32| -> i32 { (pt * 20.0).round() as i32 };
        let available_tw = pt_to_tw(available_width);

        let mut lines = Vec::new();
        let mut current_line = Line { fragments: vec![], ..Default::default() };
        let mut current_width = first_line_indent;
        let mut current_grid_extra: f32 = 0.0; // charGrid extra for line-break

        // Word buffer spans across fragment boundaries so that a single word
        // split across two runs (e.g. "te" in Run1 + "st" in Run2) is kept
        // together for line-break decisions.
        let mut word = String::new();
        let mut word_width: f32 = 0.0;
        let mut word_grid_extra: f32 = 0.0; // charGrid extra width for line-break
        let mut word_style: Option<RunStyle> = None;
        let mut word_field_type: Option<FieldType> = None;
        let mut word_run_index: usize = 0;
        let mut word_char_offset: usize = 0;

        // Helper: flush the accumulated word into current_line, breaking if needed.
        macro_rules! flush_word {
            ($style:expr) => {
                if !word.is_empty() {
                    let ws = word_style.take().unwrap_or_else(|| $style.clone());
                    let wft = word_field_type.take();
                    if pt_to_tw(current_width + current_grid_extra + word_width + word_grid_extra) > available_tw && !current_line.fragments.is_empty() {
                        lines.push(std::mem::take(&mut current_line));
                        current_width = 0.0; current_grid_extra = 0.0;
                        current_grid_extra = 0.0;
                    }
                    current_line.fragments.push(LineFragment {
                        text: std::mem::take(&mut word),
                        width: word_width,
                        style: ws,
                        tab_alignment: None,
                        tab_position: None,
                        field_type: wft,
                        run_index: word_run_index,
                        char_offset: word_char_offset,
                    });
                    current_width += word_width;
                    current_grid_extra += word_grid_extra;
                    word_width = 0.0;
                    word_grid_extra = 0.0;
                }
            };
        }

        for &(text, style, frag_field_type, frag_run_index, frag_char_start) in fragments {
            let font_size = self.resolve_font_size(style, para_style);
            let mut char_pos_in_run = frag_char_start;

            let cs = snap_character_spacing(style.character_spacing.unwrap_or(0.0));

            // Pre-resolve font metrics and GDI width maps for this fragment.
            // Avoids repeated font family resolution and HashMap lookups per character.
            let latin_metrics = self.metrics_for(style, para_style);
            let cjk_metrics = self.metrics_for_cjk(style, para_style);
            let latin_gdi_map = self.registry.get_gdi_char_widths(&latin_metrics.family, font_size);
            let cjk_gdi_map = cjk_metrics.map(|m| self.registry.get_gdi_char_widths(&m.family, font_size)).flatten();

            for ch in text.chars() {
                let (char_metrics, gdi_map) = if kinsoku::is_cjk(ch) {
                    if let Some(cjk_m) = cjk_metrics {
                        (cjk_m, cjk_gdi_map)
                    } else {
                        (latin_metrics, latin_gdi_map)
                    }
                } else {
                    (latin_metrics, latin_gdi_map)
                };
                let mut char_width = self.registry.char_width_pt_with_gdi_map(ch, font_size, char_metrics, gdi_map) + cs;
                // charGrid: each character occupies 1 grid cell for wrapping.
                // For line-break, effective width = max(char_width, pitch).
                // When char_width > pitch, char overflows into next cell visually
                // but still counts as 1 cell for wrapping purposes.
                let char_grid_extra = if let Some(pitch) = grid_char_pitch {
                    if pitch > 0.0 && char_width > 0.0 && ch != ' ' && ch != '\t' && ch != '\n' {
                        // Effective cell width = pitch (1 cell). Extra = pitch - natural width.
                        // If char is wider than pitch, extra is 0 (char naturally fills the cell).
                        let effective_cell = pitch;
                        (effective_cell - char_width).max(0.0)
                    } else { 0.0 }
                } else { 0.0 };

                if ch == ' ' || ch == '\t' || ch == '\n' || ch == '\x0C' || ch == '\x0B' {
                    // Whitespace: flush word, then handle the whitespace
                    flush_word!(style);

                    if ch == '\n' || ch == '\x0C' || ch == '\x0B' {
                        // Set break type on the current line before pushing
                        let break_type = match ch {
                            '\x0C' => LineBreakType::PageBreak,
                            '\x0B' => LineBreakType::ColumnBreak,
                            _ => LineBreakType::Normal,
                        };
                        current_line.break_type = break_type;
                        lines.push(std::mem::take(&mut current_line));
                        current_width = 0.0; current_grid_extra = 0.0;
                    } else {
                        // Space or tab
                        if ch == '\t' {
                            // COM-confirmed: tab positions are absolute from left margin.
                            // current_width is relative to the indent start, so we add
                            // indent_left to get the absolute position from margin.
                            let indent_left = para_style.indent_left.unwrap_or(0.0);
                            let abs_pos = current_width + indent_left;
                            let (next_pos, tab_align) = if !para_style.tab_stops.is_empty() {
                                para_style.tab_stops.iter()
                                    .find(|ts| ts.position > abs_pos + 0.01)
                                    .map(|ts| (ts.position, ts.alignment))
                                    .unwrap_or_else(|| {
                                        let tab_stop = self.default_tab_stop;
                                        (((abs_pos / tab_stop).floor() + 1.0) * tab_stop, TabStopAlignment::Left)
                                    })
                            } else {
                                let tab_stop = self.default_tab_stop;
                                (((abs_pos / tab_stop).floor() + 1.0) * tab_stop, TabStopAlignment::Left)
                            };
                            // Convert absolute tab position back to relative width
                            let next_relative = next_pos - indent_left;
                            let w = (next_relative - current_width).max(char_width);
                            current_line.fragments.push(LineFragment {
                                text: TAB_STRING.to_owned(),
                                width: w,
                                style: style.clone(),
                                tab_alignment: Some(tab_align),
                                tab_position: Some(next_pos),
                                field_type: None,
                                run_index: frag_run_index,
                                char_offset: char_pos_in_run,
                            });
                            current_width += w;
                        } else {
                            // Regular space
                            current_line.fragments.push(LineFragment {
                                text: SPACE_STRING.to_owned(),
                                width: char_width,
                                style: style.clone(),
                                tab_alignment: None,
                                tab_position: None,
                                field_type: None,
                                run_index: frag_run_index,
                                char_offset: char_pos_in_run,
                            });
                            current_width += char_width; current_grid_extra += char_grid_extra;
                        }
                    }
                } else if is_break_after(ch) {
                    // Characters like '-', '/' that allow a line break AFTER them.
                    // Include them in the current word, flush, and allow a break.
                    if word_style.is_none() {
                        word_style = Some(style.clone());
                        word_field_type = frag_field_type;
                        word_run_index = frag_run_index;
                        word_char_offset = char_pos_in_run;
                    }
                    word.push(ch);
                    word_width += char_width; word_grid_extra += char_grid_extra;
                    flush_word!(style);
                } else if kinsoku::is_cjk(ch) {
                    // CJK characters can break at any point
                    // autoSpaceDE: add 2.5pt gap between Latin and CJK
                    let prev_is_latin = !word.is_empty() && word.chars().last().map_or(false, |c| c.is_ascii_alphanumeric());
                    let prev_frag_is_latin = if word.is_empty() {
                        current_line.fragments.last().map_or(false, |f| {
                            f.text.chars().last().map_or(false, |c| c.is_ascii_alphanumeric() || c.is_ascii_punctuation())
                        })
                    } else { false };
                    flush_word!(style);
                    if (prev_is_latin || prev_frag_is_latin) && para_style.auto_space_de {
                        current_width += 2.5; // COM-confirmed: 2.5pt auto space
                    }

                    if pt_to_tw(current_width + current_grid_extra + char_width + char_grid_extra) > available_tw && !current_line.fragments.is_empty() {
                        if kinsoku::is_line_start_prohibited(ch) && !current_line.fragments.is_empty() {
                            current_line.fragments.push(LineFragment {
                                text: char_to_string(ch),
                                width: char_width,
                                style: style.clone(),
                                tab_alignment: None,
                                tab_position: None,
                                field_type: frag_field_type,
                                run_index: frag_run_index,
                                char_offset: char_pos_in_run,
                            });
                            lines.push(std::mem::take(&mut current_line));
                            current_width = 0.0; current_grid_extra = 0.0;
                            continue;
                        }
                        lines.push(std::mem::take(&mut current_line));
                        current_width = 0.0; current_grid_extra = 0.0;
                    }

                    if kinsoku::is_line_end_prohibited(ch) {
                        current_line.fragments.push(LineFragment {
                            text: char_to_string(ch),
                            width: char_width,
                            style: style.clone(),
                            tab_alignment: None,
                            tab_position: None,
                            field_type: frag_field_type,
                            run_index: frag_run_index,
                            char_offset: char_pos_in_run,
                        });
                        current_width += char_width; current_grid_extra += char_grid_extra;
                        continue;
                    }

                    current_line.fragments.push(LineFragment {
                        text: char_to_string(ch),
                        width: char_width,
                        style: style.clone(),
                        tab_alignment: None,
                        tab_position: None,
                        field_type: frag_field_type,
                        run_index: frag_run_index,
                        char_offset: char_pos_in_run,
                    });
                    current_width += char_width; current_grid_extra += char_grid_extra;
                } else {
                    // Regular word character — accumulate
                    if word_style.is_none() {
                        word_style = Some(style.clone());
                        word_field_type = frag_field_type;
                        word_run_index = frag_run_index;
                        word_char_offset = char_pos_in_run;
                    }
                    word.push(ch);
                    word_width += char_width; word_grid_extra += char_grid_extra;
                }
                char_pos_in_run += 1; // character index (not byte offset) for JS compatibility
            }
            // Do NOT flush word here — it may continue in the next fragment
        }

        // Flush any remaining word after all fragments
        if !word.is_empty() {
            let ws = word_style.take().unwrap_or_else(|| {
                fragments.last().map(|f| f.1.clone()).unwrap_or_default()
            });
            let wft = word_field_type.take();
            if pt_to_tw(current_width + word_width) > available_tw && !current_line.fragments.is_empty() {
                lines.push(std::mem::take(&mut current_line));
                current_width = 0.0; current_grid_extra = 0.0;
            }
            current_line.fragments.push(LineFragment {
                text: word,
                width: word_width,
                style: ws,
                tab_alignment: None,
                tab_position: None,
                field_type: wft,
                run_index: word_run_index,
                char_offset: word_char_offset,
            });
            current_width += word_width;
        }

        // Flush last line
        if !current_line.fragments.is_empty() {
            lines.push(current_line);
        }

        // Ensure at least one empty line for empty paragraphs
        if lines.is_empty() {
            lines.push(Line { fragments: vec![], ..Default::default() });
        }

        // Post-process: adjust tab fragment widths for Center/Right/Decimal alignment.
        // ECMA-376 §17.3.1.38: Center tabs center the following segment on the tab position,
        // Right tabs right-align, Decimal tabs align at the decimal point.
        for line in &mut lines {
            let frag_count = line.fragments.len();
            let mut i = 0;
            while i < frag_count {
                if let Some(align) = line.fragments[i].tab_alignment {
                    if align == TabStopAlignment::Left {
                        i += 1;
                        continue;
                    }
                    let tab_pos = line.fragments[i].tab_position.unwrap_or(0.0);
                    // Measure the segment width after this tab until next tab or end of line
                    let mut segment_width: f32 = 0.0;
                    let mut decimal_offset: Option<f32> = None;
                    let mut j = i + 1;
                    while j < frag_count {
                        if line.fragments[j].tab_alignment.is_some() {
                            break;
                        }
                        if align == TabStopAlignment::Decimal && decimal_offset.is_none() {
                            // Find decimal point position within this fragment
                            let mut char_offset: f32 = 0.0;
                            let fs = line.fragments[j].style.font_size.unwrap_or(11.0);
                            let metrics = self.registry.default_metrics();
                            for ch in line.fragments[j].text.chars() {
                                if ch == '.' || ch == ',' {
                                    decimal_offset = Some(segment_width + char_offset);
                                    break;
                                }
                                char_offset += self.registry.char_width_pt_with_fallback(ch, fs, metrics);
                            }
                        }
                        segment_width += line.fragments[j].width;
                        j += 1;
                    }

                    // Calculate the desired tab width so the segment aligns correctly
                    // Current tab width advances cursor to tab_pos. We need to adjust it
                    // so the segment is positioned according to the alignment type.
                    let current_tab_width = line.fragments[i].width;
                    let adjustment = match align {
                        TabStopAlignment::Center => segment_width / 2.0,
                        TabStopAlignment::Right => segment_width,
                        TabStopAlignment::Decimal => decimal_offset.unwrap_or(segment_width),
                        TabStopAlignment::Left => 0.0,
                    };
                    // New tab width = original width - adjustment (shift left by adjustment)
                    let new_width = (current_tab_width - adjustment).max(0.0);
                    line.fragments[i].width = new_width;
                }
                i += 1;
            }
        }

        lines
    }

    /// Calculate line height considering:
    /// 1. Font metrics (base single-line height)
    /// 2. Paragraph default font minimum (from style/docDefaults)
    /// 3. Line spacing multiplier (w:line/240, e.g. 1.15 for default)
    /// 4. Document grid snapping (linePitch)
    ///
    /// Word determines line height as the max of the run font's height
    /// and the paragraph's default font height (from the style/theme).
    /// Then applies the spacing multiplier and optionally snaps to grid.
    fn line_height(
        &self,
        font_size: f32,
        line_spacing: Option<f32>,
        line_spacing_rule: Option<&str>,
        metrics: &FontMetrics,
        snap_to_grid: bool,
        grid_pitch: Option<f32>,
    ) -> f32 {
        self.line_height_inner(font_size, line_spacing, line_spacing_rule, metrics, snap_to_grid, grid_pitch, false)
    }

    fn line_height_inner(
        &self,
        font_size: f32,
        line_spacing: Option<f32>,
        line_spacing_rule: Option<&str>,
        metrics: &FontMetrics,
        snap_to_grid: bool,
        grid_pitch: Option<f32>,
        in_table_cell: bool,
    ) -> f32 {
        // For Single/auto spacing (no explicit line_spacing or factor=1.0),
        // try COM-measured lookup table first (most accurate, includes GDI hinting)
        let is_single = match (line_spacing_rule, line_spacing) {
            (Some("exact"), _) | (Some("atLeast"), _) => false,
            (_, Some(f)) if (f - 1.0).abs() > 0.01 => false,
            _ => true,
        };

        if is_single {
            if in_table_cell {
                // Table cells: use GDI table if available, otherwise word_line_height_table_cell.
                // Grid snap is applied below via the normal path (compat_mode dependent).
                // Don't early-return here — fall through to GDI table + grid snap logic.
            } else {
                if let Some(lh) = self.registry.com_line_height(
                    &metrics.family, font_size,
                    if snap_to_grid { grid_pitch } else { None }
                ) {
                    return lh;
                }
            }
        }

        // Use GDI tmHeight table if available (most accurate).
        // Falls back to formula-based calculation.
        let ppem = (font_size * 96.0 / 72.0).round() as u32;
        let base = if let Some((h_px, _a_px, _d_px)) = self.registry.gdi_height(&metrics.family, ppem) {
            // GDI table stores tmHeight (MulDiv-based ascent + descent).
            // Body paragraphs with COM lookup use that (more accurate).
            // Table cells: tmHeight only (no tmExternalLeading) — COM-confirmed.
            let gdi_height_pt = h_px as f32 * 72.0 / 96.0;
            if metrics.is_cjk_83_64_font() {
                let raw = gdi_height_pt * 83.0 / 64.0;
                (raw * 8.0).floor() / 8.0
            } else {
                gdi_height_pt
            }
        } else if in_table_cell && self.adjust_line_height_in_table {
            metrics.word_line_height_standard(font_size)
        } else if in_table_cell {
            metrics.word_line_height_table_cell(font_size)
        } else {
            metrics.word_line_height(font_size, 96.0)
        };

        match (line_spacing_rule, line_spacing) {
            (Some("exact"), Some(val)) => val,
            (Some("atLeast"), Some(val)) => {
                // COM-confirmed: atLeast = max(grid_snap(natural), specified)
                let snapped = if snap_to_grid {
                    if let Some(pitch) = grid_pitch {
                        if pitch > 0.0 {
                            (((base + pitch * 0.5) / pitch) + 0.5).floor().max(1.0) * pitch
                        } else { base }
                    } else { base }
                } else { base };
                snapped.max(val)
            }
            _ => {
                let spaced = match line_spacing {
                    Some(factor) => base * factor,
                    None => base,
                };
                // COM-confirmed (2026-04-03, gen2_023): grid snap is only applied when
                // lineSpacing is Single (factor=1.0) or unset. Multiple spacing (factor≠1.0)
                // does NOT get grid-snapped. MS Mincho 11pt line=276: gap=16.5pt (no snap),
                // NOT 18pt (with snap).
                let is_single = match line_spacing {
                    Some(f) => (f - 1.0).abs() < 0.001,
                    None => true,
                };
                if snap_to_grid && is_single {
                    if let Some(pitch) = grid_pitch {
                        if pitch > 0.0 {
                            return (((spaced + pitch * 0.5) / pitch) + 0.5).floor().max(1.0) * pitch;
                        }
                    }
                }
                // No grid snap: ceil to 10 twips (0.5pt) — Word internal line height.
                // COM-confirmed: 80/80 tests (5 fonts x 4 sizes x 4 spacings) all match.
                // Table cells use raw value (table row height has separate calculation).
                if !in_table_cell {
                    let tw = spaced * 20.0;
                    (tw / 10.0).ceil() * 10.0 / 20.0
                } else {
                    spaced
                }
            }
        }
    }

    /// Compute line height for a line with multiple runs using Word's algorithm:
    /// max(ascent across all runs) + max(descent across all runs).
    /// Uses EastAsia font metrics for CJK text (#2).
    fn line_height_for_line(
        &self,
        line: &Line,
        para_style: &ParagraphStyle,
        para_font_size: f32,
        snap_to_grid: bool,
        grid_pitch: Option<f32>,
    ) -> f32 {
        self.line_height_for_line_inner(line, para_style, para_font_size, snap_to_grid, grid_pitch, false)
    }

    fn line_height_for_line_inner(
        &self,
        line: &Line,
        para_style: &ParagraphStyle,
        para_font_size: f32,
        snap_to_grid: bool,
        grid_pitch: Option<f32>,
        in_table_cell: bool,
    ) -> f32 {
        let default_style = RunStyle::default();

        let mut max_ascent: f32 = 0.0;
        let mut max_descent: f32 = 0.0;

        // adjustLineHeightInTable=true: use standard height without CJK 83/64
        let use_standard = in_table_cell && self.adjust_line_height_in_table;

        if line.fragments.is_empty() {
            // Empty paragraph: use pPr/rPr font size if available (direct paragraph property),
            // otherwise fall back to paragraph style's default run style.
            // COM-confirmed: 3a4f P1 empty, pPr/rPr/sz=48 (24pt) → uses Century 24pt height.
            let font_size = para_style.ppr_rpr.as_ref()
                .and_then(|r| r.font_size)
                .unwrap_or(para_font_size);
            let rpr_ref = para_style.ppr_rpr.as_ref().cloned().unwrap_or_default();
            let metrics = self.metrics_for(&rpr_ref, para_style);
            if use_standard {
                let h = metrics.word_line_height_standard(font_size);
                max_ascent = h * metrics.win_ascent / (metrics.win_ascent + metrics.win_descent);
                max_descent = h - max_ascent;
            } else {
                max_ascent = metrics.word_ascent_pt(font_size);
                max_descent = metrics.word_descent_pt(font_size);
            }
        } else {
            for frag in &line.fragments {
                let font_size = frag.style.font_size.unwrap_or(para_font_size);
                let metrics = self.metrics_for_text(&frag.text, &frag.style, para_style);
                let (asc, des) = if use_standard {
                    let h = metrics.word_line_height_standard(font_size);
                    (h * metrics.win_ascent / (metrics.win_ascent + metrics.win_descent), h * metrics.win_descent / (metrics.win_ascent + metrics.win_descent))
                } else {
                    (metrics.word_ascent_pt(font_size), metrics.word_descent_pt(font_size))
                };
                if asc > max_ascent { max_ascent = asc; }
                if des > max_descent { max_descent = des; }
            }
        }

        let run_base = max_ascent + max_descent;

        // Word uses run font height only, no max with default.
        let base = run_base;

        // Apply line spacing rule
        let line_spacing = para_style.line_spacing;
        let line_spacing_rule = para_style.line_spacing_rule.as_deref();
        match (line_spacing_rule, line_spacing) {
            (Some("exact"), Some(val)) => val,
            (Some("atLeast"), Some(val)) => {
                // COM-confirmed: atLeast = max(grid_snap(natural), specified)
                let snapped = if para_style.snap_to_grid {
                    if let Some(pitch) = grid_pitch {
                        if pitch > 0.0 {
                            (((base + pitch * 0.5) / pitch) + 0.5).floor().max(1.0) * pitch
                        } else { base }
                    } else { base }
                } else { base };
                snapped.max(val)
            }
            _ => {
                let spaced = match line_spacing {
                    Some(factor) => base * factor,
                    None => base,
                };
                // COM-confirmed (2026-04-03): grid snap only for Single (factor=1.0) or unset.
                // Multiple spacing (factor≠1.0) does NOT get grid-snapped.
                let is_single_line = match line_spacing {
                    Some(f) => (f - 1.0).abs() < 0.001,
                    None => true,
                };
                if snap_to_grid && is_single_line {
                    if let Some(pitch) = grid_pitch {
                        if pitch > 0.0 {
                            return (((spaced + pitch * 0.5) / pitch) + 0.5).floor().max(1.0) * pitch;
                        }
                    }
                }
                // Ceil to 10 twips (0.5pt) — Word internal line height precision.
                // COM-confirmed: both empty and text paragraphs use ceil.
                // Meiryo 10.5pt: CJK 83/64=20.375pt → ceil→20.5pt for both.
                let tw = spaced * 20.0;
                (tw / 10.0).ceil() * 10.0 / 20.0
            }
        }
    }

    /// Compute the vertical offset to apply to text within a line for exact/atLeast spacing.
    /// Matches Word output: exact/atLeast place text at BOTTOM of line box (extra space above).
    /// Returns the offset from line-box top to where text should start.
    fn text_y_offset_for_line(
        &self,
        line: &Line,
        para_style: &ParagraphStyle,
        para_font_size: f32,
        line_height: f32,
        grid_pitch: Option<f32>,
    ) -> f32 {
        match (para_style.line_spacing_rule.as_deref(), para_style.line_spacing) {
            (Some("exact"), Some(_)) | (Some("atLeast"), Some(_)) => {
                // exact/atLeast: text at bottom of line box (extra space above).
                let mut max_ascent: f32 = 0.0;
                let mut max_descent: f32 = 0.0;
                if line.fragments.is_empty() {
                    let metrics = self.registry.default_metrics();
                    max_ascent = metrics.word_ascent_pt(para_font_size);
                    max_descent = metrics.word_descent_pt(para_font_size);
                } else {
                    for frag in &line.fragments {
                        let font_size = frag.style.font_size.unwrap_or(para_font_size);
                        let metrics = self.metrics_for_text(&frag.text, &frag.style, para_style);
                        let asc = metrics.word_ascent_pt(font_size);
                        let des = metrics.word_descent_pt(font_size);
                        if asc > max_ascent { max_ascent = asc; }
                        if des > max_descent { max_descent = des; }
                    }
                }
                let natural = max_ascent + max_descent;
                (line_height - natural).max(0.0)
            }
            _ => {
                // Grid-snapped lines: text is vertically centered within the grid cell.
                // COM-confirmed: P1 20pt in 35.7pt grid cell → 4.9pt offset above text.
                // Compute natural height and center within line_height.
                let mut max_ascent: f32 = 0.0;
                let mut max_descent: f32 = 0.0;
                if line.fragments.is_empty() {
                    let metrics = self.doc_default_metrics();
                    max_ascent = metrics.word_ascent_pt(para_font_size);
                    max_descent = metrics.word_descent_pt(para_font_size);
                } else {
                    for frag in &line.fragments {
                        let font_size = frag.style.font_size.unwrap_or(para_font_size);
                        let metrics = self.metrics_for_text(&frag.text, &frag.style, para_style);
                        let asc = metrics.word_ascent_pt(font_size);
                        let des = metrics.word_descent_pt(font_size);
                        if asc > max_ascent { max_ascent = asc; }
                        if des > max_descent { max_descent = des; }
                    }
                }
                // Only apply text_y_offset when document grid is active.
                // Without grid, GDI TextOutW places text at cursor_y (no offset).
                // COM-confirmed: test_line_height.docx all fonts show y = cursor_y (no offset).
                let has_grid = grid_pitch.map_or(false, |p| p > 0.0) && para_style.snap_to_grid;
                if has_grid {
                    // COM-confirmed (2026-04-04): text is vertically centered within the
                    // grid cell. The offset places text in the middle of the pitch.
                    // Measured: MS Gothic 12pt pitch=18 → offset=1.0 ≈ (18-16)/2 (via GDI)
                    //           Century 10.5pt pitch=18 → offset=2.5 ≈ (18-13)/2 (via GDI)
                    // Best approximation: (pitch - GDI_height_pt) / 2, floor to 0.5pt.
                    // Fallback when GDI table unavailable: (pitch - fontSize) / 2.
                    let font_size = if !line.fragments.is_empty() {
                        line.fragments[0].style.font_size.unwrap_or(para_font_size)
                    } else { para_font_size };
                    let pitch = grid_pitch.unwrap_or(0.0);
                    if pitch > 0.0 {
                        let natural = max_ascent + max_descent;
                        let raw = (pitch - natural).max(0.0) / 2.0;
                        // Round to 0.5pt (10 twips) — COM-confirmed best fit
                        (raw * 2.0 + 0.5).floor() / 2.0
                    } else {
                        0.0
                    }
                } else {
                    0.0
                }
            }
        }
    }

    fn layout_table(
        &self,
        table: &Table,
        start_x: f32,
        cursor_y: &mut f32,
        content_width: f32,
        grid_pitch: Option<f32>,
        grid_char_pitch: Option<f32>,
        page_top: f32,
        content_height: f32,
        page_width: f32,
        page_height: f32,
        pages: &mut Vec<LayoutPage>,
        current_elements: &mut Vec<LayoutElement>,
    ) -> Vec<LayoutElement> {
        let mut elements = Vec::new();

        // Resolve column widths from grid_columns, cell widths, or equal split
        let col_widths = self.resolve_table_col_widths(table, content_width);
        let table_width: f32 = col_widths.iter().sum();

        // Table positioning: tblpPr horizontal or inline alignment
        let table_x = if let Some(ref pos) = table.style.position {
            if let Some(ref h_align) = pos.h_align {
                let (ref_left, ref_width) = match pos.h_anchor.as_deref() {
                    Some("page") => (0.0, page_width),
                    _ => (start_x, content_width), // "margin" or "text"
                };
                match h_align.as_str() {
                    "center" => ref_left + (ref_width - table_width) / 2.0,
                    "right" => ref_left + ref_width - table_width,
                    _ => ref_left,
                }
            } else {
                match pos.h_anchor.as_deref() {
                    Some("page") => pos.x,
                    _ => start_x + pos.x,
                }
            }
        } else {
            match table.style.alignment.as_deref() {
                Some("center") => start_x + (content_width - table_width) / 2.0,
                Some("right") => start_x + content_width - table_width,
                _ => start_x + table.style.indent.unwrap_or(0.0),
            }
        };

        // Default cell padding from table style or OOXML default
        // COM-measured 2026-03-29: L/R=4.95pt (99tw), T/B=0pt
        let default_pad = &table.style.default_cell_margins;
        let default_pad_l = default_pad.as_ref().and_then(|m| m.left).unwrap_or(4.95);
        let default_pad_r = default_pad.as_ref().and_then(|m| m.right).unwrap_or(4.95);
        let default_pad_t = default_pad.as_ref().and_then(|m| m.top).unwrap_or(0.0);
        let default_pad_b = default_pad.as_ref().and_then(|m| m.bottom).unwrap_or(0.0);

        // Table cell grid snap: COM-confirmed always enabled regardless of compat mode.
        // adjustLineHeightInTable is always false (COM measurement of 151 documents).
        // Previous compat<15 check was incorrect — Word 2010 mode also grid-snaps table cells.
        let table_grid_pitch: Option<f32> = if self.adjust_line_height_in_table {
            None
        } else {
            grid_pitch
        };

        let num_rows = table.rows.len();
        for (row_idx, row) in table.rows.iter().enumerate() {
            let mut row_height: f32 = 0.0;

            // First pass: calculate row height
            let mut grid_idx = row.grid_before as usize;
            for cell in row.cells.iter() {
                let span = cell.grid_span.max(1) as usize;
                // vMerge="continue" cells don't contribute to row height
                // (their content is part of the vMerge="restart" cell above)
                if cell.v_merge.as_deref() == Some("continue") || cell.v_merge.as_deref() == Some("") {
                    grid_idx += span;
                    continue;
                }
                let cell_w: f32 = col_widths[grid_idx..grid_idx + span].iter().sum();
                let pad_l = cell.margins.as_ref().and_then(|m| m.left).unwrap_or(default_pad_l);
                let pad_r = cell.margins.as_ref().and_then(|m| m.right).unwrap_or(default_pad_r);
                let mut pad_t = cell.margins.as_ref().and_then(|m| m.top).unwrap_or(default_pad_t);
                let mut pad_b = cell.margins.as_ref().and_then(|m| m.bottom).unwrap_or(default_pad_b);
                // COM-confirmed (2026-04-02): only insideH border adds row height overhead.
                // Outer borders (top/bottom/left/right) do NOT affect row height.
                // The overhead equals the insideH border width, applied once per row (not per edge).
                let border_overhead = if table.style.has_inside_h {
                    table.style.border_width.unwrap_or(if table.style.border { 0.4 } else { 0.0 })
                } else {
                    0.0
                };
                // For line-wrapping estimation, use cell_w (not inner_w after padding)
                // Word allows text to extend into cell margins for wrapping purposes
                let inner_w = cell_w.max(0.0);
                let mut cell_content_h = pad_t;

                for block in &cell.blocks {
                    match block {
                        Block::Paragraph(para) => {
                            let para_h = self.estimate_para_height(para, inner_w, table_grid_pitch, table.style.para_style.as_ref());
                            cell_content_h += para_h;
                        }
                        Block::Table(nested) => {
                            // Estimate nested table height from rows
                            // COM-confirmed: nested table width = cell width - 2 × padding
                            let nested_w = (inner_w).max(0.0);
                            for nr in &nested.rows {
                                let mut nr_h = 0.0_f32;
                                for nc in &nr.cells {
                                    let mut nc_h = 0.0_f32;
                                    for nb in &nc.blocks {
                                        if let Block::Paragraph(np) = nb {
                                            nc_h += self.estimate_para_height(np, nested_w / 2.0, table_grid_pitch, nested.style.para_style.as_ref());
                                        }
                                    }
                                    nr_h = nr_h.max(nc_h);
                                }
                                if let Some(h) = nr.height {
                                    match nr.height_rule.as_deref() {
                                        Some("exact") => { nr_h = h; }
                                        Some("atLeast") => { nr_h = nr_h.max(h); }
                                        _ => {}
                                    }
                                }
                                cell_content_h += nr_h;
                            }
                        }
                        _ => {}
                    }
                }
                cell_content_h += pad_b;
                cell_content_h += border_overhead;

                row_height = row_height.max(cell_content_h);
                grid_idx += span;
            }

            // Grid snap row content height, then round to 0.5pt (10tw)
            // COM-confirmed: table row height = round_10tw(ceil(content / pitch) * pitch)
            // linesAndChars mode: Word does NOT grid-snap table row heights
            // (COM-measured: row heights are natural content height, not grid multiples)
            if let Some(pitch) = table_grid_pitch {
                if pitch > 0.0 && row_height > 0.0 && grid_char_pitch.is_none() {
                    let snapped = (row_height / pitch).ceil() * pitch;
                    // Round to 0.5pt (10 twips) — matches Word internal precision
                    row_height = (snapped * 2.0).round() / 2.0;
                }
            }

            // Apply trHeight constraint
            // rule=exact: fixed height; rule=atLeast: minimum height
            // COM-confirmed (2026-04-04): when hRule is absent but trHeight val is specified,
            // Word treats it as atLeast (COM reports HeightRule=1).
            // Note: actual rendered gap may be smaller than atLeast value when multi-cell
            // rows have cells with different content heights and spacing interactions.
            if let Some(h) = row.height {
                match row.height_rule.as_deref() {
                    Some("exact") => { row_height = h; }
                    Some("atLeast") | None => { row_height = row_height.max(h); }
                    _ => {} // explicit "auto" string: content determines height
                }
            }

            if row_height == 0.0 {
                let metrics = self.doc_default_metrics();
                row_height = self.line_height_inner(self.default_font_size, None, None, metrics, true, table_grid_pitch, true);
            }
            // Page break check: if this row won't fit, push current page and reset
            // Allow break if there are elements from previous rows OR from before the table
            let has_content = !elements.is_empty() || !current_elements.is_empty();
            let page_bottom = page_top + content_height;
            let row_overflows = *cursor_y + row_height > page_bottom;
            // Row splitting: when cantSplit=false (default) and the row overflows,
            // split it across pages rather than moving the entire row to the next page.
            // Word splits rows at the page boundary, keeping partial content on each page.
            // Only split single-cell rows in single-row tables (1x1 "box" tables).
            // Multi-row or multi-cell table splitting is too complex and can cause
            // regressions in page count for other documents.
            let is_single_cell_row = row.cells.len() == 1 && num_rows == 1;
            let needs_row_split = row_overflows && !row.cant_split && has_content && is_single_cell_row;
            if row_overflows && has_content && !needs_row_split {
                // Push all accumulated elements (including previous rows) to current page
                current_elements.extend(std::mem::take(&mut elements));
                pages.push(LayoutPage {
                    width: page_width,
                    height: page_height,
                    elements: std::mem::take(current_elements),
                });
                *cursor_y = page_top;
            }

            // Second pass: render cells
            // Track actual content height per cell for row_height correction
            let is_exact_row = row.height_rule.as_deref() == Some("exact");
            let mut max_actual_cell_h: f32 = row_height;
            let elements_before_row = elements.len();
            // Apply gridBefore: skip leading grid columns
            let mut cell_x = table_x + col_widths[..row.grid_before as usize].iter().sum::<f32>();
            let num_cells = row.cells.len();
            for (cell_idx, cell) in row.cells.iter().enumerate() {
                let span = cell.grid_span.max(1) as usize;
                // vMerge="continue" cells: skip content but still draw borders
                let is_vmerge_continue = cell.v_merge.as_deref() == Some("continue") || cell.v_merge.as_deref() == Some("");
                let grid_end = (col_widths.len()).min(
                    col_widths.iter().enumerate()
                        .scan(0.0f32, |acc, (i, w)| { *acc += w; Some((i, *acc)) })
                        .find(|(_, acc)| (*acc - cell_x + table_x).abs() < 0.1)
                        .map(|(i, _)| i + span)
                        .unwrap_or(col_widths.len())
                );
                // Calculate cell width from grid columns
                let cell_start_grid = col_widths.iter()
                    .scan(0.0f32, |acc, w| { let prev = *acc; *acc += w; Some(prev) })
                    .enumerate()
                    .find(|(_, acc)| (cell_x - table_x - acc).abs() < 0.5)
                    .map(|(i, _)| i)
                    .unwrap_or(0);
                let cell_end_grid = (cell_start_grid + span).min(col_widths.len());
                let cell_w: f32 = col_widths[cell_start_grid..cell_end_grid].iter().sum();

                let pad_l = cell.margins.as_ref().and_then(|m| m.left).unwrap_or(default_pad_l);
                let pad_r = cell.margins.as_ref().and_then(|m| m.right).unwrap_or(default_pad_r);
                let mut pad_t = cell.margins.as_ref().and_then(|m| m.top).unwrap_or(default_pad_t);
                let mut pad_b = cell.margins.as_ref().and_then(|m| m.bottom).unwrap_or(default_pad_b);

                // COM-confirmed (2026-04-02): border padding does NOT affect text positioning.
                // insideH overhead only affects row height (handled in first pass via border_overhead).
                // Text position within cell is determined by cell margins only.

                // Emit cell shading (background fill) before cell content
                if let Some(ref shading_color) = cell.shading {
                    if !shading_color.is_empty() && shading_color != "auto" {
                        let color_hex = if shading_color.starts_with('#') {
                            shading_color.clone()
                        } else {
                            format!("#{}", shading_color)
                        };
                        elements.push(LayoutElement::new(cell_x, *cursor_y, cell_w, row_height, LayoutContent::CellShading {
                                color: color_hex,
                        }));
                    }
                }

                // COM-confirmed: Word uses cell_w for text wrapping (text overflows into padding)
                let inner_w = cell_w.max(0.0);
                let mut cell_elements: Vec<LayoutElement> = Vec::new();
                let mut content_h: f32 = 0.0;

                // Layout blocks in document order (paragraphs and nested tables interleaved)
                let is_exact = row.height_rule.as_deref() == Some("exact");
                if !is_vmerge_continue {
                for block in &cell.blocks {
                // Clip content that overflows exact row height
                if is_exact && content_h + pad_t >= row_height {
                    break;
                }
                match block {
                Block::Table(nested) => {
                    // COM-confirmed: nested table width = outer cell width - 2 × padding
                    let nested_x = cell_x + pad_l;
                    let nested_content_w = (cell_w - pad_l - pad_r).max(0.0);
                    let mut nested_y = content_h;
                    let mut dummy_pages = Vec::new();
                    let mut dummy_elems = Vec::new();
                    let nested_elements = self.layout_table(
                        nested, nested_x, &mut nested_y, nested_content_w, table_grid_pitch,
                        grid_char_pitch,
                        0.0, 99999.0, 0.0, 99999.0,
                        &mut dummy_pages, &mut dummy_elems,
                    );
                    for elem in nested_elements {
                        cell_elements.push(elem);
                    }
                    content_h = nested_y;
                }
                Block::Paragraph(para) => {
                let para = para;
                    // Apply table style pPr as fallback (ECMA-376: table style pPr < paragraph style < direct)
                    // Word resets line spacing to Single and space_after to 0 for table cell
                    // paragraphs that inherit from Normal style (no direct spacing in pPr).
                    // COM-measured: Normal outside table = ls=13.80(1.15x) sa=10,
                    //               Normal inside table = ls=12.00(Single) sa=0.
                    let effective_line_spacing = para.style.line_spacing
                        .or_else(|| table.style.para_style.as_ref().and_then(|ps| ps.line_spacing));
                    let effective_line_rule = para.style.line_spacing_rule.as_deref()
                        .or_else(|| table.style.para_style.as_ref().and_then(|ps| ps.line_spacing_rule.as_deref()));
                    // COM-confirmed (2026-03-31): table cells inherit Normal style's lineSpacing.
                    // test_table_borders.docx: Cell(1,1) ls=13.80 (1.15x from Normal), NOT reset to Single.
                    // Only override with table style's lineSpacing if the table style explicitly sets it.
                    let style_has_explicit_rule = effective_line_rule == Some("exact") || effective_line_rule == Some("atLeast");
                    let should_reset = !para.style.has_direct_spacing && !style_has_explicit_rule;
                    let tbl_has_ls = table.style.para_style.as_ref().and_then(|ps| ps.line_spacing).is_some();
                    let (effective_line_spacing, effective_line_rule) = if tbl_has_ls && !para.style.has_direct_spacing {
                        let tbl_ls = table.style.para_style.as_ref().and_then(|ps| ps.line_spacing);
                        let tbl_lr = table.style.para_style.as_ref().and_then(|ps| ps.line_spacing_rule.as_deref());
                        (tbl_ls, tbl_lr)
                    } else {
                        (effective_line_spacing, effective_line_rule)
                    };
                    let effective_space_before = if should_reset {
                        0.0
                    } else if let (Some(bl), Some(pitch)) = (para.style.before_lines, table_grid_pitch) {
                        bl / 100.0 * pitch
                    } else {
                        para.style.space_before
                            .or_else(|| table.style.para_style.as_ref().and_then(|ps| ps.space_before))
                            .unwrap_or(0.0)
                    };
                    let effective_space_after = if should_reset {
                        None
                    } else {
                        para.style.space_after
                            .or_else(|| table.style.para_style.as_ref().and_then(|ps| ps.space_after))
                    };
                    content_h += effective_space_before;
                    let para_content_start_h = content_h;
                    {
                        // Paragraph indentation within cell (relative to cell content area)
                        let p_indent_left = para.style.indent_left.unwrap_or(0.0);
                        let p_indent_right = para.style.indent_right.unwrap_or(0.0);
                        let p_first_line_indent = para.style.indent_first_line.unwrap_or(0.0);
                        let wrap_w = (inner_w - p_indent_left - p_indent_right).max(0.0);
                        // Hanging indent (firstLineIndent < 0): first line starts further LEFT,
                        // so it has MORE available width, not less.
                        // first_line_left = indent_left + firstLineIndent (may be < indent_left)
                        // first_line_wrap = inner_w - first_line_left - indent_right
                        let first_line_wrap_w = if p_first_line_indent < 0.0 {
                            (inner_w - (p_indent_left + p_first_line_indent).max(0.0) - p_indent_right).max(0.0)
                        } else {
                            (wrap_w - p_first_line_indent).max(0.0)
                        };

                        // Collect runs into lines with greedy wrapping
                        // Tuple: (text, font_size, width, bold, italic, underline, underline_style, strikethrough, font_family, color, highlight, character_spacing)
                        let mut lines: Vec<Vec<(String, f32, f32, bool, bool, bool, Option<String>, bool, Option<String>, Option<String>, Option<String>, f32)>> = Vec::new();
                        let mut current_line: Vec<(String, f32, f32, bool, bool, bool, Option<String>, bool, Option<String>, Option<String>, Option<String>, f32)> = Vec::new();
                        let mut line_x: f32 = 0.0;
                        let mut is_first_line = true;

                        for run in &para.runs {
                            let font_size = self.resolve_font_size(&run.style, &para.style);
                            let bold = self.resolve_bold(&run.style, &para.style);
                            let font_family = self.resolve_font_family_for_text(&run.text, &run.style, &para.style)
                                .map(|s| s.to_string());

                            // Split text character by character for wrapping
                            let cs = snap_character_spacing(run.style.character_spacing.unwrap_or(0.0));
                            let mut buf = String::new();
                            let mut buf_w: f32 = 0.0;
                            for ch in run.text.chars() {
                                let cm = self.metrics_for_char(ch, &run.style, &para.style);
                                let cw = self.registry.char_width_pt_with_fallback(ch, font_size, cm) + cs;
                                let effective_wrap = if is_first_line { first_line_wrap_w } else { wrap_w };
                                // Trailing spaces don't trigger line wrapping (Word behavior)
                                let is_space = ch == ' ' || ch == '\u{3000}';
                                if !is_space && line_x + buf_w + cw > effective_wrap && !(current_line.is_empty() && buf.is_empty()) {
                                    // Kinsoku: line-start-prohibited chars (）。、etc.) stay on current line
                                    if kinsoku::is_line_start_prohibited(ch) {
                                        // Add to buffer and break AFTER this char
                                        buf.push(ch);
                                        buf_w += cw;
                                        if !buf.is_empty() {
                                            current_line.push((buf.clone(), font_size, buf_w, bold, run.style.italic, run.style.underline, run.style.underline_style.clone(), run.style.strikethrough, font_family.clone(), run.style.color.clone(), run.style.highlight.clone(), cs));
                                            buf.clear();
                                            buf_w = 0.0;
                                        }
                                        lines.push(std::mem::take(&mut current_line));
                                        line_x = 0.0;
                                        is_first_line = false;
                                        continue;
                                    }
                                    // Flush buffer to current line, then wrap
                                    if !buf.is_empty() {
                                        current_line.push((buf.clone(), font_size, buf_w, bold, run.style.italic, run.style.underline, run.style.underline_style.clone(), run.style.strikethrough, font_family.clone(), run.style.color.clone(), run.style.highlight.clone(), cs));
                                        buf.clear();
                                        buf_w = 0.0;
                                    }
                                    lines.push(std::mem::take(&mut current_line));
                                    line_x = 0.0;
                                    is_first_line = false;
                                }
                                buf.push(ch);
                                buf_w += cw;
                            }
                            if !buf.is_empty() {
                                current_line.push((buf, font_size, buf_w, bold, run.style.italic, run.style.underline, run.style.underline_style.clone(), run.style.strikethrough, font_family, run.style.color.clone(), run.style.highlight.clone(), cs));
                                line_x += buf_w;
                            }
                        }
                        if !current_line.is_empty() {
                            lines.push(current_line);
                        }

                        if lines.is_empty() {
                            let metrics = self.doc_default_metrics();
                            content_h += self.line_height_inner(self.default_font_size, effective_line_spacing, effective_line_rule, metrics, para.style.snap_to_grid, table_grid_pitch, true);
                        }

                        let total_lines = lines.len();
                        for (line_idx, line) in lines.iter().enumerate() {
                            // Clip content that overflows exact row height
                            if is_exact && content_h + pad_t >= row_height {
                                break;
                            }
                            // Line height = max of all runs in line (in_table_cell=true: no default font minimum)
                            let lh: f32 = line.iter().map(|(_text, fs, _, _, _, _, _, _, font_family, _, _, _)| {
                                let metrics = match font_family.as_deref() {
                                    Some(ff) => self.registry.get(ff),
                                    None => self.registry.default_metrics(),
                                };
                                self.line_height_inner(*fs, effective_line_spacing, effective_line_rule, metrics, para.style.snap_to_grid, table_grid_pitch, true)
                            }).fold(0.0_f32, f32::max);

                            // Paragraph indentation: first line uses indent_left + first_line_indent
                            let line_indent = p_indent_left + if line_idx == 0 { p_first_line_indent } else { 0.0 };

                            // Calculate line total width for alignment
                            let line_total_w: f32 = line.iter().map(|(_, _, tw, _, _, _, _, _, _, _, _, _)| tw).sum();
                            let effective_wrap = if line_idx == 0 { first_line_wrap_w } else { wrap_w };

                            // Justify: non-last lines for jc=both, all lines for distribute
                            let is_last_line = line_idx == total_lines - 1;
                            let should_justify = (para.alignment == Alignment::Justify && !is_last_line)
                                || para.alignment == Alignment::Distribute;

                            // Apply paragraph alignment within cell (wrap_w = available after indent)
                            let align_offset = if should_justify {
                                0.0
                            } else {
                                match para.alignment {
                                    Alignment::Center => (effective_wrap - line_total_w).max(0.0) / 2.0,
                                    Alignment::Right => (effective_wrap - line_total_w).max(0.0),
                                    _ => 0.0,
                                }
                            };

                            // Justify: CJK punctuation compression + space/gap distribution
                            let mut frag_width_adj: Vec<f32> = vec![0.0; line.len()];
                            let mut frag_spacing: Vec<f32> = vec![0.0; line.len()];
                            if should_justify && line.len() > 1 {
                                let mut slack = effective_wrap - line_total_w;

                                // Phase 1: CJK punctuation compression (only when overflowing)
                                if slack < 0.0 {
                                    for (fi, (text, fs, _, _, _, _, _, _, _, _, _, _)) in line.iter().enumerate() {
                                        for ch in text.chars() {
                                            if kinsoku::is_cjk_compressible(ch) {
                                                let fm = self.registry.default_metrics();
                                                let char_w = self.registry.char_width_pt_with_fallback(ch, *fs, fm);
                                                let savings = char_w * 0.5;
                                                frag_width_adj[fi] -= savings;
                                                slack += savings;
                                            }
                                        }
                                    }
                                }

                                // Phase 2: Distribute slack at word spaces, then CJK gaps
                                if slack > 0.0 {
                                    let space_count = line.iter()
                                        .enumerate()
                                        .filter(|(i, (text, _, _, _, _, _, _, _, _, _, _, _))| *i < line.len() - 1 && text.trim().is_empty())
                                        .count();
                                    if space_count > 0 {
                                        let per_space = slack / space_count as f32;
                                        for (fi, (text, _, _, _, _, _, _, _, _, _, _, _)) in line.iter().enumerate() {
                                            if fi < line.len() - 1 && text.trim().is_empty() {
                                                frag_spacing[fi] += per_space;
                                            }
                                        }
                                    } else {
                                        // No word spaces: distribute between CJK fragments
                                        let has_cjk = line.iter().any(|(text, _, _, _, _, _, _, _, _, _, _, _)| text.chars().any(|c| kinsoku::is_cjk(c)));
                                        if has_cjk {
                                            let gap_count = line.len() - 1;
                                            if gap_count > 0 {
                                                let per_gap = slack / gap_count as f32;
                                                for fi in 0..gap_count {
                                                    frag_spacing[fi] += per_gap;
                                                }
                                            }
                                        }
                                    }
                                }
                            }

                            let mut rx = 0.0_f32;
                            for (frag_idx, (text, fs, tw, bold, italic, underline, underline_style, strikethrough, font_family, color, highlight, cs)) in line.iter().enumerate() {
                                let adj_w = *tw + frag_width_adj[frag_idx];
                                cell_elements.push(LayoutElement::new(cell_x + pad_l + line_indent + align_offset + rx, content_h, adj_w, lh, LayoutContent::Text {
                                        text: text.clone(),
                                        font_size: *fs,
                                        font_family: font_family.clone(),
                                        bold: *bold,
                                        italic: *italic,
                                        underline: *underline,
                                        underline_style: underline_style.clone(),
                                        strikethrough: *strikethrough,
                                        color: color.clone(),
                                        highlight: highlight.clone(),
                                        character_spacing: *cs,
                                        field_type: None,
                                }));
                                rx += adj_w + frag_spacing[frag_idx];
                            }
                            content_h += lh;
                        }
                        content_h += effective_space_after.unwrap_or(0.0);

                        // Render shapes attached to this paragraph (e.g. bracketPair)
                        // pos.y = offset from paragraph start (Word COM confirmed)
                        for shape in &para.shapes {
                            if let Some(ref pos) = shape.position {
                                cell_elements.push(LayoutElement::new(cell_x + pad_l + pos.x, para_content_start_h + pos.y, shape.width, shape.height, LayoutContent::PresetShape {
                                        shape_type: shape.shape_type.clone(),
                                        stroke_color: shape.stroke_color.clone(),
                                        stroke_width: shape.stroke_width.unwrap_or(0.5),
                                }));
                            }
                        }
                    }
                }
                _ => {}
                } // match block
                } // for block
                } // if !is_vmerge_continue

                // Track actual cell height for row_height correction
                // Use max(y + height) from cell_elements as true content bottom,
                // since content_h may undercount due to spacing interactions
                if !is_vmerge_continue && !is_exact_row {
                    let max_elem_bottom = cell_elements.iter()
                        .map(|e| e.y + e.height)
                        .fold(0.0_f32, f32::max);
                    let actual = pad_t + max_elem_bottom.max(content_h) + pad_b;
                    if actual > max_actual_cell_h {
                        max_actual_cell_h = actual;
                    }
                }

                // Apply vAlign offset
                let v_offset = match cell.v_align.as_deref() {
                    Some("center") => ((row_height - pad_t - pad_b - content_h) / 2.0).max(0.0),
                    Some("bottom") => (row_height - pad_t - pad_b - content_h).max(0.0),
                    _ => 0.0, // top (default)
                };

                // Emit cell elements with absolute Y positions
                let dy = *cursor_y + pad_t + v_offset;
                for mut elem in cell_elements {
                    elem.y += dy;
                    // Also update y-coords inside TableBorder content (nested tables)
                    if let LayoutContent::TableBorder { ref mut y1, ref mut y2, .. } = elem.content {
                        *y1 += dy;
                        *y2 += dy;
                    }
                    elements.push(elem);
                }

                // Draw cell borders if table has borders OR cell has its own borders
                let has_cell_borders = cell.borders.as_ref().map_or(false, |b| {
                    b.top.is_some() || b.bottom.is_some() || b.left.is_some() || b.right.is_some()
                });
                if table.style.border || has_cell_borders {
                    let bx = cell_x;
                    let by = *cursor_y;

                    // Resolve border color and width from cell borders, falling back to table style
                    let resolve_border = |side: Option<&BorderDef>| -> (Option<String>, f32) {
                        if let Some(b) = side {
                            let c = b.color.as_ref().map(|c| {
                                if c.starts_with('#') { c.clone() } else { format!("#{}", c) }
                            });
                            (c, b.width)
                        } else if table.style.border {
                            // Table-level borders: use table style color, default to black
                            let c = Some(table.style.border_color.as_ref()
                                .map(|c| if c.starts_with('#') { c.clone() } else { format!("#{}", c) })
                                .unwrap_or_else(|| "#000000".to_string()));
                            (c, table.style.border_width.unwrap_or(0.4))
                        } else {
                            (None, 0.4)
                        }
                    };

                    let cell_borders = cell.borders.as_ref();
                    let (top_color, top_width) = resolve_border(cell_borders.and_then(|b| b.top.as_ref()));
                    let (bot_color, bot_width) = resolve_border(cell_borders.and_then(|b| b.bottom.as_ref()));
                    let (left_color, left_width) = resolve_border(cell_borders.and_then(|b| b.left.as_ref()));
                    let (right_color, right_width) = resolve_border(cell_borders.and_then(|b| b.right.as_ref()));

                    // When cells have their own borders (tcBorders), draw each side per cell.
                    // When using table-level borders, use collapsed model to avoid double-drawing.
                    let use_collapsed = table.style.border && !has_cell_borders;

                    // Top — skip for vMerge continue cells (internal to merged range)
                    if !is_vmerge_continue && top_color.is_some() && (!use_collapsed || row_idx == 0) {
                        elements.push(LayoutElement::new(bx, by, cell_w, 0.0, LayoutContent::TableBorder {
                                x1: bx, y1: by, x2: bx + cell_w, y2: by,
                                color: top_color, width: top_width,
                        }));
                    }
                    // Bottom — skip for vMerge continue cells unless next row is not continue
                    let next_is_continue = if row_idx + 1 < num_rows {
                        table.rows[row_idx + 1].cells.get(cell_idx)
                            .map_or(false, |nc| nc.v_merge.as_deref() == Some("continue") || nc.v_merge.as_deref() == Some(""))
                    } else {
                        false
                    };
                    if bot_color.is_some() && !next_is_continue {
                        elements.push(LayoutElement::new(bx, by + row_height, cell_w, 0.0, LayoutContent::TableBorder {
                                x1: bx, y1: by + row_height, x2: bx + cell_w, y2: by + row_height,
                                color: bot_color, width: bot_width,
                        }));
                    }
                    // Left
                    if left_color.is_some() && (!use_collapsed || cell_idx == 0) {
                        elements.push(LayoutElement::new(bx, by, 0.0, row_height, LayoutContent::TableBorder {
                                x1: bx, y1: by, x2: bx, y2: by + row_height,
                                color: left_color, width: left_width,
                        }));
                    }
                    // Right
                    if right_color.is_some() {
                        elements.push(LayoutElement::new(bx + cell_w, by, 0.0, row_height, LayoutContent::TableBorder {
                                x1: bx + cell_w, y1: by, x2: bx + cell_w, y2: by + row_height,
                                color: right_color, width: right_width,
                        }));
                    }
                }

                cell_x += cell_w;
            }

            // If actual content exceeds estimated row_height, fix border elements
            if max_actual_cell_h > row_height + 0.01 {
                let old_h = row_height;
                row_height = max_actual_cell_h;
                let by = *cursor_y;
                let old_bottom = by + old_h;
                let new_bottom = by + row_height;
                for elem in elements[elements_before_row..].iter_mut() {
                    match &mut elem.content {
                        LayoutContent::TableBorder { y1, y2, .. } => {
                            if (*y1 - old_bottom).abs() < 0.5 { *y1 = new_bottom; }
                            if (*y2 - old_bottom).abs() < 0.5 { *y2 = new_bottom; }
                        }
                        LayoutContent::CellShading { .. } => {
                            if (elem.height - old_h).abs() < 0.5 {
                                elem.height = row_height;
                            }
                        }
                        _ => {}
                    }
                }
            }

            // Row splitting across pages: when the row content extends beyond
            // the current page bottom, split elements between current and next page.
            // This handles single-cell rows with many paragraphs (e.g. list boxes).
            let row_bottom = *cursor_y + row_height;
            if row_bottom > page_bottom + 0.5 && !row.cant_split {
                let split_y = page_bottom;
                // Partition elements: those fitting on current page vs overflow
                let row_elements = elements.split_off(elements_before_row);
                let mut current_page_elems: Vec<LayoutElement> = Vec::new();
                let mut next_page_elems: Vec<LayoutElement> = Vec::new();

                for elem in row_elements {
                    let elem_top = elem.y;
                    match &elem.content {
                        LayoutContent::TableBorder { y1, y2, x1, x2, ref color, width } => {
                            // Horizontal borders: keep on their respective page
                            if (y1 - y2).abs() < 0.1 {
                                // Horizontal line
                                if *y1 <= split_y + 0.5 {
                                    current_page_elems.push(elem);
                                } else {
                                    // Shift to next page
                                    let shift = split_y - page_top;
                                    let mut e = elem;
                                    e.y -= shift;
                                    if let LayoutContent::TableBorder { ref mut y1, ref mut y2, .. } = e.content {
                                        *y1 -= shift;
                                        *y2 -= shift;
                                    }
                                    next_page_elems.push(e);
                                }
                            } else {
                                // Vertical border: split at page boundary
                                // Current page portion
                                let vy_top = *y1;
                                let vy_bot = *y2;
                                if vy_top < split_y {
                                    current_page_elems.push(LayoutElement::new(
                                        elem.x, elem.y, elem.width, split_y - vy_top,
                                        LayoutContent::TableBorder {
                                            x1: *x1, y1: vy_top, x2: *x2, y2: split_y,
                                            color: color.clone(), width: *width,
                                        },
                                    ));
                                }
                                // Next page portion
                                if vy_bot > split_y {
                                    let shift = split_y - page_top;
                                    let new_y1 = page_top;
                                    let new_y2 = vy_bot - shift;
                                    next_page_elems.push(LayoutElement::new(
                                        elem.x, new_y1, elem.width, new_y2 - new_y1,
                                        LayoutContent::TableBorder {
                                            x1: *x1, y1: new_y1, x2: *x2, y2: new_y2,
                                            color: color.clone(), width: *width,
                                        },
                                    ));
                                }
                            }
                        }
                        LayoutContent::CellShading { ref color } => {
                            // Split shading across pages
                            let shade_bottom = elem.y + elem.height;
                            if elem.y < split_y {
                                let clip_h = (split_y - elem.y).min(elem.height);
                                current_page_elems.push(LayoutElement::new(
                                    elem.x, elem.y, elem.width, clip_h,
                                    LayoutContent::CellShading { color: color.clone() },
                                ));
                            }
                            if shade_bottom > split_y {
                                let shift = split_y - page_top;
                                let new_y = (elem.y - shift).max(page_top);
                                let new_h = shade_bottom - shift - new_y;
                                next_page_elems.push(LayoutElement::new(
                                    elem.x, new_y, elem.width, new_h.max(0.0),
                                    LayoutContent::CellShading { color: color.clone() },
                                ));
                            }
                        }
                        _ => {
                            // Text and other elements
                            if elem_top < split_y {
                                current_page_elems.push(elem);
                            } else {
                                let shift = split_y - page_top;
                                let mut e = elem;
                                e.y -= shift;
                                next_page_elems.push(e);
                            }
                        }
                    }
                }

                // Add a closing horizontal border at split point on current page
                // and an opening horizontal border at page_top on next page
                // (to show table continuation)

                // Push current page elements
                elements.extend(current_page_elems);
                current_elements.extend(std::mem::take(&mut elements));
                pages.push(LayoutPage {
                    width: page_width,
                    height: page_height,
                    elements: std::mem::take(current_elements),
                });

                // Handle multi-page overflow: if next_page_elems still overflow,
                // keep splitting into additional pages.
                let mut remaining = next_page_elems;
                let mut current_split_base = page_top;
                loop {
                    // Find the maximum Y in remaining elements
                    let max_y = remaining.iter().map(|e| {
                        match &e.content {
                            LayoutContent::TableBorder { y2, .. } => *y2,
                            _ => e.y + e.height,
                        }
                    }).fold(0.0_f32, f32::max);

                    if max_y <= page_bottom + 0.5 {
                        // Everything fits on this page
                        break;
                    }

                    // Need another split at page_bottom
                    let next_split = page_bottom;
                    let mut this_page: Vec<LayoutElement> = Vec::new();
                    let mut overflow: Vec<LayoutElement> = Vec::new();

                    for elem in remaining {
                        let elem_top = elem.y;
                        match &elem.content {
                            LayoutContent::TableBorder { y1, y2, x1, x2, ref color, width } => {
                                if (y1 - y2).abs() < 0.1 {
                                    if *y1 <= next_split + 0.5 {
                                        this_page.push(elem);
                                    } else {
                                        let shift = next_split - page_top;
                                        let mut e = elem;
                                        e.y -= shift;
                                        if let LayoutContent::TableBorder { ref mut y1, ref mut y2, .. } = e.content {
                                            *y1 -= shift;
                                            *y2 -= shift;
                                        }
                                        overflow.push(e);
                                    }
                                } else {
                                    let vy_top = *y1;
                                    let vy_bot = *y2;
                                    if vy_top < next_split {
                                        this_page.push(LayoutElement::new(
                                            elem.x, elem.y, elem.width, next_split - vy_top,
                                            LayoutContent::TableBorder {
                                                x1: *x1, y1: vy_top, x2: *x2, y2: next_split,
                                                color: color.clone(), width: *width,
                                            },
                                        ));
                                    }
                                    if vy_bot > next_split {
                                        let shift = next_split - page_top;
                                        let new_y1 = page_top;
                                        let new_y2 = vy_bot - shift;
                                        overflow.push(LayoutElement::new(
                                            elem.x, new_y1, elem.width, new_y2 - new_y1,
                                            LayoutContent::TableBorder {
                                                x1: *x1, y1: new_y1, x2: *x2, y2: new_y2,
                                                color: color.clone(), width: *width,
                                            },
                                        ));
                                    }
                                }
                            }
                            LayoutContent::CellShading { ref color } => {
                                let shade_bottom = elem.y + elem.height;
                                if elem.y < next_split {
                                    let clip_h = (next_split - elem.y).min(elem.height);
                                    this_page.push(LayoutElement::new(
                                        elem.x, elem.y, elem.width, clip_h,
                                        LayoutContent::CellShading { color: color.clone() },
                                    ));
                                }
                                if shade_bottom > next_split {
                                    let shift = next_split - page_top;
                                    let new_y = (elem.y - shift).max(page_top);
                                    let new_h = shade_bottom - shift - new_y;
                                    overflow.push(LayoutElement::new(
                                        elem.x, new_y, elem.width, new_h.max(0.0),
                                        LayoutContent::CellShading { color: color.clone() },
                                    ));
                                }
                            }
                            _ => {
                                if elem_top < next_split {
                                    this_page.push(elem);
                                } else {
                                    let shift = next_split - page_top;
                                    let mut e = elem;
                                    e.y -= shift;
                                    overflow.push(e);
                                }
                            }
                        }
                    }

                    pages.push(LayoutPage {
                        width: page_width,
                        height: page_height,
                        elements: this_page,
                    });
                    remaining = overflow;
                }

                elements = remaining;
                let overflow_on_next = row_bottom - split_y;
                let pages_used = ((overflow_on_next) / content_height).floor() as usize;
                *cursor_y = page_top + overflow_on_next - (pages_used as f32 * content_height);
            } else {
                *cursor_y += row_height;
            }
        }

        elements
    }

    /// Resolve column widths for a table.
    /// Priority: grid_columns > cell widths > equal split.
    fn resolve_table_col_widths(&self, table: &Table, content_width: f32) -> Vec<f32> {
        // 1. Use grid_columns if available
        // When nested table overflows parent cell, Word keeps earlier columns
        // at their specified width and shrinks only the last column to fit.
        if !table.grid_columns.is_empty() {
            let total: f32 = table.grid_columns.iter().sum();
            let indent = table.style.indent.unwrap_or(0.0);
            let available = content_width - indent;
            // Floating tables (tblpPr) are not constrained by content_width
            let is_floating = table.style.position.is_some();
            if !is_floating && total > available && table.grid_columns.len() > 1 {
                let mut cols = table.grid_columns.clone();
                let prefix_sum: f32 = cols[..cols.len() - 1].iter().sum();
                let last = (available - prefix_sum).max(0.0);
                *cols.last_mut().unwrap() = last;
                return cols;
            }
            return table.grid_columns.clone();
        }

        // 2. Use cell widths from first row
        if let Some(first_row) = table.rows.first() {
            let cell_widths: Vec<f32> = first_row.cells.iter()
                .filter_map(|c| c.width)
                .collect();
            if cell_widths.len() == first_row.cells.len() && !cell_widths.is_empty() {
                return cell_widths;
            }
        }

        // 3. Use table style width
        if let Some(tw) = table.style.width {
            let num_cols = table.rows.first().map_or(1, |r| r.cells.len().max(1));
            return vec![tw / num_cols as f32; num_cols];
        }

        // 4. Equal split fallback
        let num_cols = table.rows.first().map_or(1, |r| r.cells.len().max(1));
        vec![content_width / num_cols as f32; num_cols]
    }

    /// Estimate paragraph height for table cell height calculation.
    fn estimate_para_height(&self, para: &Paragraph, available_width: f32, grid_pitch: Option<f32>, table_para_style: Option<&ParagraphStyle>) -> f32 {
        let mut height = 0.0;
        // Table cells snap to grid in default Word mode
        let snap = para.style.snap_to_grid;
        // COM-confirmed (2026-03-31): table cells inherit Normal style's lineSpacing.
        // Only override with table style if it explicitly defines lineSpacing.
        let raw_ls = para.style.line_spacing
            .or_else(|| table_para_style.and_then(|ps| ps.line_spacing));
        let raw_lr = para.style.line_spacing_rule.as_deref()
            .or_else(|| table_para_style.and_then(|ps| ps.line_spacing_rule.as_deref()));
        let style_has_explicit_rule = raw_lr == Some("exact") || raw_lr == Some("atLeast");
        let should_reset = !para.style.has_direct_spacing && !style_has_explicit_rule;
        let tbl_has_ls = table_para_style.and_then(|ps| ps.line_spacing).is_some();
        let (eff_ls, eff_lr): (Option<f32>, Option<&str>) = if tbl_has_ls && !para.style.has_direct_spacing {
            let tbl_ls = table_para_style.and_then(|ps| ps.line_spacing);
            let tbl_lr = table_para_style.and_then(|ps| ps.line_spacing_rule.as_deref());
            (tbl_ls, tbl_lr)
        } else {
            (raw_ls, raw_lr)
        };

        // estimate_para_height is called for table cell content.
        // COM-confirmed: table cells use no-grid line height (grid snap disabled inside cells).
        // Use COM table with grid_pitch=None to get no_grid value.
        if para.runs.is_empty() {
            // Use pPr/rPr font for empty paragraph height
            let empty_fs = para.style.ppr_rpr.as_ref()
                .and_then(|r| r.font_size)
                .unwrap_or(self.resolve_font_size(&RunStyle::default(), &para.style));
            let rpr_ref = para.style.ppr_rpr.as_ref().cloned().unwrap_or_default();
            let metrics = self.metrics_for(&rpr_ref, &para.style);
            let is_single_empty = eff_lr.is_none() || eff_lr == Some("auto");
            if is_single_empty {
                height += metrics.word_line_height_table_cell(empty_fs);
            } else {
                height += self.line_height_inner(empty_fs, eff_ls, eff_lr, metrics, false, None, true);
            }
        } else {
            let para_font_size = self.resolve_font_size(
                para.runs.first().map(|r| &r.style).unwrap_or(&RunStyle::default()),
                &para.style,
            );
            // Use break_into_lines for accurate line count (handles kinsoku, word break, etc.)
            let indent_l = para.style.indent_left.unwrap_or(0.0);
            let indent_r = para.style.indent_right.unwrap_or(0.0);
            let first_indent = para.style.indent_first_line.unwrap_or(0.0);
            let effective_width = (available_width - indent_l - indent_r).max(1.0);

            let fragments: Vec<(&str, &RunStyle, Option<FieldType>, usize, usize)> = para.runs.iter().enumerate()
                .map(|(ri, run)| (run.text.as_str(), &run.style, None, ri, 0))
                .collect();
            let lines = self.break_into_lines(&fragments, effective_width, first_indent, &para.style, None);
            let line_count = lines.len().max(1);

            let mut max_line_height: f32 = 0.0;
            for run in &para.runs {
                let font_size = self.resolve_font_size(&run.style, &para.style);
                let metrics = self.metrics_for_text(&run.text, &run.style, &para.style);
                let is_single_run = match (eff_lr, eff_ls) {
                    (Some("exact"), _) | (Some("atLeast"), _) => false,
                    (_, Some(f)) if (f - 1.0).abs() > 0.01 => false,
                    _ => true,
                };
                let lh = if is_single_run {
                    metrics.word_line_height_table_cell(font_size)
                } else {
                    self.line_height_inner(font_size, eff_ls, eff_lr, metrics, false, None, true)
                };
                if lh > max_line_height { max_line_height = lh; }
            }
            height += max_line_height * line_count as f32;
        }

        if should_reset {
            // Word resets inherited Normal-style spacing to 0 in table cells
            // but preserves style-defined exact/atLeast spacing
        } else {
            height += if let (Some(bl), Some(pitch)) = (para.style.before_lines, grid_pitch) {
                bl / 100.0 * pitch
            } else {
                para.style.space_before
                    .or_else(|| table_para_style.and_then(|ps| ps.space_before))
                    .unwrap_or(0.0)
            };
            height += para.style.space_after
                .or_else(|| table_para_style.and_then(|ps| ps.space_after))
                .unwrap_or(0.0);
        }
        height
    }
}

#[derive(Default)]
struct Line {
    fragments: Vec<LineFragment>,
    /// What kind of break follows this line (normal line break, page break, or column break)
    break_type: LineBreakType,
}

impl Default for LineBreakType {
    fn default() -> Self {
        Self::Normal
    }
}

#[derive(Clone)]
struct LineFragment {
    text: String,
    width: f32,
    style: RunStyle,
    /// For tab fragments: the alignment type of the tab stop they target.
    /// None for non-tab fragments.
    tab_alignment: Option<TabStopAlignment>,
    /// For tab fragments: the absolute position (from left margin) of the tab stop.
    tab_position: Option<f32>,
    /// Field type for dynamic content (PAGE, NUMPAGES)
    field_type: Option<FieldType>,
    /// Source run index within the paragraph (for editing support)
    run_index: usize,
    /// Source character byte offset within the run (for editing support)
    char_offset: usize,
}

/// Marker for page/column break after a line
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
enum LineBreakType {
    Normal,
    PageBreak,   // \x0C
    ColumnBreak, // \x0B
}

#[cfg(test)]
mod tests {
    #[allow(unused_imports)]
    use super::*;

    #[test]
    #[ignore]
    fn bench_layout_multi() {
        // Benchmark multiple documents to find the pattern
        let docs_dir = "../../tools/golden-test/documents/docx";
        let mut results = Vec::new();
        if let Ok(entries) = std::fs::read_dir(docs_dir) {
            for entry in entries.flatten() {
                let path = entry.path();
                if path.extension().map_or(false, |e| e == "docx") {
                    if let Ok(data) = std::fs::read(&path) {
                        if let Ok(doc) = crate::parse_docx(&data) {
                            let engine = LayoutEngine::for_document(&doc);
                            let _ = engine.layout(&doc); // warmup
                            let start = std::time::Instant::now();
                            let r = engine.layout(&doc);
                            let ms = start.elapsed().as_micros() as f64 / 1000.0;
                            let elems: usize = r.pages.iter().map(|p| p.elements.len()).sum();
                            results.push((path.file_name().unwrap().to_string_lossy().to_string(), ms, r.pages.len(), elems));
                        }
                    }
                }
            }
        }
        results.sort_by(|a, b| b.1.partial_cmp(&a.1).unwrap());
        println!("\nTop 10 slowest:");
        for (name, ms, pages, elems) in results.iter().take(10) {
            println!("  {:.1}ms  {}p {}el  {}", ms, pages, elems, name);
        }
        let total: f64 = results.iter().map(|r| r.1).sum();
        println!("Total: {:.0}ms for {} docs", total, results.len());
    }

    #[test]
    #[ignore]
    fn bench_layout_1ec_detail() {
        let data = std::fs::read("../../tools/golden-test/documents/docx/1ec1091177b1_006.docx")
            .expect("read docx");
        let doc = crate::parse_docx(&data).expect("parse");
        let engine = LayoutEngine::for_document(&doc);
        let _ = engine.layout(&doc); // warmup

        // Measure engine creation + layout
        let n = 20;
        let start = std::time::Instant::now();
        let mut result = None;
        for _ in 0..n {
            let eng = LayoutEngine::for_document(&doc);
            result = Some(eng.layout(&doc));
        }
        let full_ms = start.elapsed().as_micros() as f64 / 1000.0 / n as f64;

        // Measure layout only (engine reused)
        let start = std::time::Instant::now();
        for _ in 0..n {
            result = Some(engine.layout(&doc));
        }
        let layout_ms = start.elapsed().as_micros() as f64 / 1000.0 / n as f64;

        // Measure engine creation only
        let start = std::time::Instant::now();
        for _ in 0..n {
            let _eng = LayoutEngine::for_document(&doc);
        }
        let engine_ms = start.elapsed().as_micros() as f64 / 1000.0 / n as f64;

        let r = result.unwrap();
        let total_elems: usize = r.pages.iter().map(|p| p.elements.len()).sum();
        println!("Engine: {:.1}ms, Layout: {:.1}ms, Full: {:.1}ms, Pages: {}, Elements: {}", engine_ms, layout_ms, full_ms, r.pages.len(), total_elems);
    }

    #[test]
    #[ignore]
    fn bench_layout_1ec() {
        let data = std::fs::read("../../tools/golden-test/documents/docx/1ec1091177b1_006.docx")
            .expect("read docx");
        // Profile: parse vs layout
        let parse_start = std::time::Instant::now();
        let doc = crate::parse_docx(&data).expect("parse");
        let parse_ms = parse_start.elapsed().as_millis();
        println!("Parse: {}ms", parse_ms);

        // Count blocks, runs, characters
        let mut total_chars = 0usize;
        let mut total_runs = 0usize;
        let mut total_blocks = 0usize;
        let mut total_table_cells = 0usize;
        for page in &doc.pages {
            total_blocks += page.blocks.len();
            for block in &page.blocks {
                match block {
                    crate::ir::Block::Paragraph(p) => {
                        total_runs += p.runs.len();
                        for r in &p.runs { total_chars += r.text.len(); }
                    }
                    crate::ir::Block::Table(t) => {
                        for row in &t.rows {
                            for cell in &row.cells {
                                total_table_cells += 1;
                                for b in &cell.blocks {
                                    if let crate::ir::Block::Paragraph(p) = b {
                                        total_runs += p.runs.len();
                                        for r in &p.runs { total_chars += r.text.len(); }
                                    }
                                }
                            }
                        }
                    }
                    _ => {}
                }
            }
            // TextBox content
            for tb in &page.text_boxes {
                for b in &tb.blocks {
                    if let crate::ir::Block::Paragraph(p) = b {
                        total_runs += p.runs.len();
                        for r in &p.runs { total_chars += r.text.len(); }
                    }
                }
            }
        }
        println!("Doc: {} blocks, {} runs, {} chars, {} table_cells", total_blocks, total_runs, total_chars, total_table_cells);

        // Warmup
        let engine = LayoutEngine::for_document(&doc);
        let _ = engine.layout(&doc);
        // Measure
        let n = 10;
        let start = std::time::Instant::now();
        for _ in 0..n {
            let engine = LayoutEngine::for_document(&doc);
            let _ = engine.layout(&doc);
        }
        let elapsed = start.elapsed();
        println!("Layout: {:.1}ms avg ({} runs, {:.0}ms total)",
            elapsed.as_millis() as f64 / n as f64, n, elapsed.as_millis());
    }

    #[test]
    #[ignore] // debug only
    fn debug_1ec_y_positions() {
        let data = std::fs::read("../../tools/golden-test/documents/docx/1ec1091177b1_006.docx")
            .expect("read docx");
        let doc = crate::parse_docx(&data).expect("parse");

        // Print table structure in detail
        for page in &doc.pages {
            for (bi, block) in page.blocks.iter().enumerate() {
                if let crate::ir::Block::Table(t) = block {
                    println!("B{}: Table {}rows", bi, t.rows.len());
                    for (ri, row) in t.rows.iter().enumerate() {
                        let hr = row.height_rule.as_deref().unwrap_or("auto");
                        let hv = row.height.unwrap_or(0.0);
                        println!("  Row{}: h_spec={:.1} rule={} cells={}", ri, hv, hr, row.cells.len());
                        for (ci, cell) in row.cells.iter().enumerate() {
                            let pad = &cell.margins;
                            let pad_t = pad.as_ref().and_then(|m| m.top).unwrap_or(-1.0);
                            let pad_b = pad.as_ref().and_then(|m| m.bottom).unwrap_or(-1.0);
                            println!("    Cell{}: paras={} pad_t={:.1} pad_b={:.1} vmerge={:?}",
                                ci, cell.blocks.len(), pad_t, pad_b, cell.v_merge);
                            for (pi, blk) in cell.blocks.iter().enumerate() {
                                if let crate::ir::Block::Paragraph(p) = blk {
                                    let text: String = p.runs.iter().flat_map(|r| r.text.chars()).take(30).collect();
                                    let snap = p.style.snap_to_grid;
                                    let ls = p.style.line_spacing;
                                    let lr = p.style.line_spacing_rule.as_deref().unwrap_or("?");
                                    let sa = p.style.space_after.unwrap_or(0.0);
                                    let sb = p.style.space_before.unwrap_or(0.0);
                                    let font = p.runs.first().map(|r| r.style.font_family.as_deref().unwrap_or("?")).unwrap_or("?");
                                    let fsz = p.runs.first().map(|r| r.style.font_size.unwrap_or(0.0)).unwrap_or(0.0);
                                    println!("      P{}: snap={} ls={:?} lr={} sa={:.1} sb={:.1} font={}@{:.1} \"{}\"",
                                        pi, snap, ls, lr, sa, sb, font, fsz, text);
                                }
                            }
                        }
                    }
                }
            }
        }

        let engine = LayoutEngine::for_document(&doc);
        let result = engine.layout(&doc);
        println!("\nPages: {}", result.pages.len());
        // Show all elements with their Y positions grouped by row
        for (pi, lpage) in result.pages.iter().enumerate() {
            println!("--- Page {} ---", pi);
            let mut prev_y: f32 = -1.0;
            for el in &lpage.elements {
                match &el.content {
                    LayoutContent::Text { ref text, font_size, .. } => {
                        if (el.y - prev_y).abs() > 0.1 {
                            let snippet: String = text.chars().take(25).collect();
                            println!("  TEXT y={:.1} h={:.1} fs={:.1} \"{}\"", el.y, el.height, font_size, snippet);
                            prev_y = el.y;
                        }
                    }
                    _ => {}
                }
            }
        }
    }
}
