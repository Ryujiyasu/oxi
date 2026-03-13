mod kinsoku;

use crate::font::{FontMetrics, FontMetricsRegistry};
use crate::ir::*;

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
}

pub enum LayoutContent {
    Text {
        text: String,
        font_size: f32,
        font_family: Option<String>,
        bold: bool,
        italic: bool,
        underline: bool,
        strikethrough: bool,
        color: Option<String>,
        highlight: Option<String>,
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
    },
}

pub struct LayoutEngine {
    default_font_size: f32,
    registry: FontMetricsRegistry,
}

/// Word's default heading font sizes (in points)
fn heading_default_font_size(level: u8) -> f32 {
    match level {
        1 => 16.0,
        2 => 13.0,
        3 => 12.0,
        4 => 11.0,
        5 => 11.0,
        6 => 11.0,
        _ => 11.0,
    }
}

impl LayoutEngine {
    pub fn new() -> Self {
        Self {
            default_font_size: 11.0,
            registry: FontMetricsRegistry::load(),
        }
    }

    pub fn layout(&self, doc: &Document) -> LayoutResult {
        let mut pages = Vec::new();

        for page in &doc.pages {
            let laid_out = self.layout_page(page);
            pages.extend(laid_out);
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

    /// Resolve font family for a run
    fn resolve_font_family<'a>(&self, run_style: &'a RunStyle, para_style: &'a ParagraphStyle) -> Option<&'a str> {
        if let Some(ref ff) = run_style.font_family {
            return Some(ff.as_str());
        }
        if let Some(ref drs) = para_style.default_run_style {
            if let Some(ref ff) = drs.font_family {
                return Some(ff.as_str());
            }
        }
        None
    }

    /// Get font metrics for a run (uses registry with font-family resolution)
    fn metrics_for(&self, run_style: &RunStyle, para_style: &ParagraphStyle) -> &FontMetrics {
        match self.resolve_font_family(run_style, para_style) {
            Some(family) => self.registry.get(family),
            None => self.registry.default_metrics(),
        }
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

    fn layout_page(&self, page: &Page) -> Vec<LayoutPage> {
        let content_width = page.size.width - page.margin.left - page.margin.right;
        let content_height = page.size.height - page.margin.top - page.margin.bottom;
        let start_x = page.margin.left;
        let start_y = page.margin.top;

        let mut pages: Vec<LayoutPage> = Vec::new();
        let mut elements: Vec<LayoutElement> = Vec::new();
        let mut cursor_y = start_y;
        let mut prev_para_style_id: Option<String> = None;

        let grid_pitch = page.grid_line_pitch;

        for block in &page.blocks {
            match block {
                Block::Paragraph(para) => {
                    let para_elements = self.layout_paragraph(
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
                        prev_para_style_id.as_deref(),
                    );
                    elements.extend(para_elements);
                    prev_para_style_id = para.style.style_id.clone();
                }
                Block::Table(table) => {
                    let table_elements = self.layout_table(
                        table,
                        start_x,
                        &mut cursor_y,
                        content_width,
                        grid_pitch,
                    );
                    elements.extend(table_elements);
                }
                Block::Image(img) => {
                    if cursor_y + img.height > start_y + content_height {
                        pages.push(LayoutPage {
                            width: page.size.width,
                            height: page.size.height,
                            elements: std::mem::take(&mut elements),
                        });
                        cursor_y = start_y;
                    }
                    elements.push(LayoutElement {
                        x: start_x,
                        y: cursor_y,
                        width: img.width,
                        height: img.height,
                        content: LayoutContent::Image {
                            data: img.data.clone(),
                            content_type: img.content_type.clone(),
                        },
                    });
                    cursor_y += img.height;
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

        pages
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
    ) -> Vec<LayoutElement> {
        let mut elements = Vec::new();

        // Apply paragraph spacing (space_before)
        let mut space_before = para.style.space_before.unwrap_or(0.0);

        // Contextual spacing: suppress space_before when previous paragraph
        // has the same style and both have contextualSpacing enabled.
        if para.style.contextual_spacing {
            if let (Some(cur_id), Some(prev_id)) = (para.style.style_id.as_deref(), prev_style_id) {
                if cur_id == prev_id {
                    space_before = 0.0;
                }
            }
        }

        // Suppress space_before at the top of a page
        if (*cursor_y - page_top).abs() < 0.01 {
            space_before = 0.0;
        }

        *cursor_y += space_before;

        let indent_left = para.style.indent_left.unwrap_or(0.0);
        let indent_right = para.style.indent_right.unwrap_or(0.0);
        let first_line_indent = para.style.indent_first_line.unwrap_or(0.0);
        let available_width = content_width - indent_left - indent_right;

        // Render list marker if present
        if let Some(ref marker) = para.style.list_marker {
            let default_style = RunStyle::default();
            let marker_style = para.runs.first().map(|r| &r.style).unwrap_or(&default_style);
            let marker_font_size = self.resolve_font_size(marker_style, &para.style);
            let marker_metrics = self.metrics_for(marker_style, &para.style);
            let marker_width: f32 = marker
                .chars()
                .map(|c| marker_metrics.char_width_pt(c, marker_font_size))
                .sum();
            let list_indent = para.style.list_indent.unwrap_or(18.0);
            let marker_x = start_x + indent_left - list_indent;
            let line_height = self.line_height(marker_font_size, para.style.line_spacing, marker_metrics, para.style.snap_to_grid, grid_pitch);

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

            elements.push(LayoutElement {
                x: marker_x,
                y: *cursor_y,
                width: marker_width,
                height: line_height,
                content: LayoutContent::Text {
                    text: marker.clone(),
                    font_size: marker_font_size,
                    font_family: None,
                    bold: false,
                    italic: false,
                    underline: false,
                    strikethrough: false,
                    color: None,
                    highlight: None,
                },
            });
        }

        // Collect all text fragments with their styles
        let fragments: Vec<(&str, &RunStyle)> = para
            .runs
            .iter()
            .map(|r| (r.text.as_str(), &r.style))
            .collect();

        // Resolve font size for line breaking
        let default_style = RunStyle::default();
        let para_font_size = self.resolve_font_size(
            para.runs.first().map(|r| &r.style).unwrap_or(&default_style),
            &para.style,
        );

        // Line-break the text
        let lines = self.break_into_lines(&fragments, available_width, first_line_indent, &para.style);

        for (line_idx, line) in lines.iter().enumerate() {
            let first_style = line.fragments.first().map(|f| &f.style).unwrap_or(&default_style);
            let font_size = first_style.font_size.unwrap_or(para_font_size);
            let metrics = self.metrics_for(first_style, &para.style);
            let line_height = self.line_height(font_size, para.style.line_spacing, metrics, para.style.snap_to_grid, grid_pitch);

            // Page break check
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

            let extra_indent = if line_idx == 0 { first_line_indent } else { 0.0 };
            let line_x = start_x + indent_left + extra_indent;

            // Alignment offset
            let line_text_width: f32 = line.fragments.iter().map(|f| f.width).sum();
            let is_last_line = line_idx == lines.len() - 1;
            let align_offset = match para.alignment {
                Alignment::Left => 0.0,
                Alignment::Center => (available_width - extra_indent - line_text_width) / 2.0,
                Alignment::Right => available_width - extra_indent - line_text_width,
                Alignment::Justify => 0.0,
            };

            // For justify: distribute extra space between fragments (not on last line)
            let justify_extra = if para.alignment == Alignment::Justify
                && !is_last_line
                && line.fragments.len() > 1
            {
                let slack = available_width - extra_indent - line_text_width;
                if slack > 0.0 {
                    let gap_count = line
                        .fragments
                        .iter()
                        .filter(|f| f.text.trim().is_empty())
                        .count();
                    if gap_count > 0 {
                        slack / gap_count as f32
                    } else {
                        slack / (line.fragments.len() - 1) as f32
                    }
                } else {
                    0.0
                }
            } else {
                0.0
            };

            let mut x = line_x + align_offset;

            for (frag_idx, frag) in line.fragments.iter().enumerate() {
                let resolved_font_size = frag.style.font_size.unwrap_or(para_font_size);
                let resolved_bold = self.resolve_bold(&frag.style, &para.style);
                elements.push(LayoutElement {
                    x,
                    y: *cursor_y,
                    width: frag.width,
                    height: line_height,
                    content: LayoutContent::Text {
                        text: frag.text.clone(),
                        font_size: resolved_font_size,
                        font_family: frag.style.font_family.clone(),
                        bold: resolved_bold,
                        italic: frag.style.italic,
                        underline: frag.style.underline,
                        strikethrough: frag.style.strikethrough,
                        color: frag.style.color.clone(),
                        highlight: frag.style.highlight.clone(),
                    },
                });
                x += frag.width;

                // Add justify spacing
                if justify_extra > 0.0 && frag_idx < line.fragments.len() - 1 {
                    let has_space_gaps = line.fragments.iter().any(|f| f.text.trim().is_empty());
                    if has_space_gaps {
                        if frag.text.trim().is_empty() {
                            x += justify_extra;
                        }
                    } else {
                        x += justify_extra;
                    }
                }
            }

            *cursor_y += line_height;
        }

        let space_after = para.style.space_after.unwrap_or(0.0);
        *cursor_y += space_after;

        elements
    }

    fn break_into_lines(
        &self,
        fragments: &[(&str, &RunStyle)],
        available_width: f32,
        first_line_indent: f32,
        para_style: &ParagraphStyle,
    ) -> Vec<Line> {
        let mut lines = Vec::new();
        let mut current_line = Line { fragments: vec![] };
        let mut current_width = first_line_indent;
        let mut _is_first_line = true;

        for &(text, style) in fragments {
            let font_size = self.resolve_font_size(style, para_style);
            let metrics = self.metrics_for(style, para_style);

            let mut word = String::new();
            let mut word_width: f32 = 0.0;

            for ch in text.chars() {
                let char_width = metrics.char_width_pt(ch, font_size);

                if ch == ' ' || ch == '\n' {
                    // Flush current word
                    if !word.is_empty() {
                        if current_width + word_width > available_width && !current_line.fragments.is_empty() {
                            lines.push(std::mem::take(&mut current_line));
                            current_width = 0.0;
                            _is_first_line = false;
                        }
                        current_line.fragments.push(LineFragment {
                            text: word.clone(),
                            width: word_width,
                            style: style.clone(),
                        });
                        current_width += word_width;
                        word.clear();
                        word_width = 0.0;
                    }

                    if ch == '\n' {
                        lines.push(std::mem::take(&mut current_line));
                        current_width = 0.0;
                        _is_first_line = false;
                    } else {
                        let space_width = metrics.char_width_pt(' ', font_size);
                        current_line.fragments.push(LineFragment {
                            text: " ".to_string(),
                            width: space_width,
                            style: style.clone(),
                        });
                        current_width += space_width;
                    }
                } else if kinsoku::is_cjk(ch) {
                    // CJK characters can break at any point
                    if !word.is_empty() {
                        if current_width + word_width > available_width && !current_line.fragments.is_empty() {
                            lines.push(std::mem::take(&mut current_line));
                            current_width = 0.0;
                            _is_first_line = false;
                        }
                        current_line.fragments.push(LineFragment {
                            text: word.clone(),
                            width: word_width,
                            style: style.clone(),
                        });
                        current_width += word_width;
                        word.clear();
                        word_width = 0.0;
                    }

                    if current_width + char_width > available_width && !current_line.fragments.is_empty() {
                        if kinsoku::is_line_start_prohibited(ch) && !current_line.fragments.is_empty() {
                            current_line.fragments.push(LineFragment {
                                text: ch.to_string(),
                                width: char_width,
                                style: style.clone(),
                            });
                            lines.push(std::mem::take(&mut current_line));
                            current_width = 0.0;
                            _is_first_line = false;
                            continue;
                        }
                        lines.push(std::mem::take(&mut current_line));
                        current_width = 0.0;
                        _is_first_line = false;
                    }

                    if kinsoku::is_line_end_prohibited(ch) {
                        current_line.fragments.push(LineFragment {
                            text: ch.to_string(),
                            width: char_width,
                            style: style.clone(),
                        });
                        current_width += char_width;
                        continue;
                    }

                    current_line.fragments.push(LineFragment {
                        text: ch.to_string(),
                        width: char_width,
                        style: style.clone(),
                    });
                    current_width += char_width;
                } else {
                    word.push(ch);
                    word_width += char_width;
                }
            }

            // Flush remaining word
            if !word.is_empty() {
                if current_width + word_width > available_width && !current_line.fragments.is_empty() {
                    lines.push(std::mem::take(&mut current_line));
                    current_width = 0.0;
                }
                current_line.fragments.push(LineFragment {
                    text: word,
                    width: word_width,
                    style: style.clone(),
                });
                current_width += word_width;
            }
        }

        // Flush last line
        if !current_line.fragments.is_empty() {
            lines.push(current_line);
        }

        // Ensure at least one empty line for empty paragraphs
        if lines.is_empty() {
            lines.push(Line { fragments: vec![] });
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
        metrics: &FontMetrics,
        snap_to_grid: bool,
        grid_pitch: Option<f32>,
    ) -> f32 {
        // Word uses GDI TEXTMETRIC integer rounding internally.
        // DPI=150 gives best empirical match (avg 0.24pt error vs Word COM).
        const GDI_DPI: f32 = 150.0;

        // Run font's GDI-rounded height
        let run_base = metrics.gdi_line_height(font_size, GDI_DPI);

        // Default font height at the RUN's font size (not at 11pt!).
        // Word computes a minimum line height from the paragraph mark's font,
        // which inherits from docDefaults/theme. This minimum scales with the
        // run's font size, ensuring consistent line height even for fonts with
        // small vertical metrics (e.g. MS Gothic with ratio=1.0).
        let default_base = self.registry.default_metrics().gdi_line_height(font_size, GDI_DPI);

        let base = run_base.max(default_base);

        let spaced = match line_spacing {
            Some(factor) => base * factor,
            None => base,
        };

        // Apply grid snapping: round up to next multiple of grid pitch
        if snap_to_grid {
            if let Some(pitch) = grid_pitch {
                if pitch > 0.0 {
                    return (spaced / pitch).ceil() * pitch;
                }
            }
        }

        spaced
    }

    fn layout_table(
        &self,
        table: &Table,
        start_x: f32,
        cursor_y: &mut f32,
        content_width: f32,
        grid_pitch: Option<f32>,
    ) -> Vec<LayoutElement> {
        let mut elements = Vec::new();
        let num_cols = table.rows.first().map_or(1, |r| r.cells.len().max(1));
        let col_width = content_width / num_cols as f32;

        for row in &table.rows {
            let mut row_height: f32 = 0.0;

            // First pass: calculate row height
            for (_col_idx, cell) in row.cells.iter().enumerate() {
                let mut cell_y = *cursor_y;

                for block in &cell.blocks {
                    if let Block::Paragraph(para) = block {
                        for run in &para.runs {
                            let font_size = self.resolve_font_size(&run.style, &para.style);
                            let metrics = self.metrics_for(&run.style, &para.style);
                            let line_height = self.line_height(font_size, para.style.line_spacing, metrics, para.style.snap_to_grid, grid_pitch);
                            cell_y += line_height;
                        }
                        if para.runs.is_empty() {
                            let default_metrics = self.registry.default_metrics();
                            cell_y += self.line_height(self.default_font_size, None, default_metrics, true, grid_pitch);
                        }
                    }
                }

                row_height = row_height.max(cell_y - *cursor_y);
            }

            if row_height == 0.0 {
                let default_metrics = self.registry.default_metrics();
                row_height = self.line_height(self.default_font_size, None, default_metrics, true, grid_pitch);
            }

            // Second pass: render cells
            for (col_idx, cell) in row.cells.iter().enumerate() {
                let cell_x = start_x + col_idx as f32 * col_width;
                let mut text_y = *cursor_y;

                for block in &cell.blocks {
                    if let Block::Paragraph(para) = block {
                        for run in &para.runs {
                            let font_size = self.resolve_font_size(&run.style, &para.style);
                            let metrics = self.metrics_for(&run.style, &para.style);
                            let lh = self.line_height(font_size, para.style.line_spacing, metrics, para.style.snap_to_grid, grid_pitch);
                            let text_width = run
                                .text
                                .chars()
                                .map(|c| metrics.char_width_pt(c, font_size))
                                .sum();

                            elements.push(LayoutElement {
                                x: cell_x + 2.0,
                                y: text_y,
                                width: text_width,
                                height: lh,
                                content: LayoutContent::Text {
                                    text: run.text.clone(),
                                    font_size,
                                    font_family: run.style.font_family.clone(),
                                    bold: self.resolve_bold(&run.style, &para.style),
                                    italic: run.style.italic,
                                    underline: run.style.underline,
                                    strikethrough: run.style.strikethrough,
                                    color: run.style.color.clone(),
                                    highlight: run.style.highlight.clone(),
                                },
                            });
                            text_y += lh;
                        }
                    }
                }

                // Draw cell borders if table has borders
                if table.style.border {
                    let bx = cell_x;
                    let by = *cursor_y;
                    // Top
                    elements.push(LayoutElement {
                        x: bx, y: by, width: col_width, height: 0.0,
                        content: LayoutContent::TableBorder {
                            x1: bx, y1: by, x2: bx + col_width, y2: by,
                        },
                    });
                    // Bottom
                    elements.push(LayoutElement {
                        x: bx, y: by + row_height, width: col_width, height: 0.0,
                        content: LayoutContent::TableBorder {
                            x1: bx, y1: by + row_height, x2: bx + col_width, y2: by + row_height,
                        },
                    });
                    // Left
                    elements.push(LayoutElement {
                        x: bx, y: by, width: 0.0, height: row_height,
                        content: LayoutContent::TableBorder {
                            x1: bx, y1: by, x2: bx, y2: by + row_height,
                        },
                    });
                    // Right
                    elements.push(LayoutElement {
                        x: bx + col_width, y: by, width: 0.0, height: row_height,
                        content: LayoutContent::TableBorder {
                            x1: bx + col_width, y1: by, x2: bx + col_width, y2: by + row_height,
                        },
                    });
                }
            }

            *cursor_y += row_height;
        }

        elements
    }
}

#[derive(Default)]
struct Line {
    fragments: Vec<LineFragment>,
}

struct LineFragment {
    text: String,
    width: f32,
    style: RunStyle,
}
