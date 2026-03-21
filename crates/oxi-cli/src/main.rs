use std::collections::HashMap;
use std::fs;
use oxidocs_core::ir::{Block, Paragraph, Table, Alignment, TextBox, FloatingPosition};
use oxidocs_core::font::FontMetricsRegistry;
use oxipdf_core::ir::*;

fn main() {
    let args: Vec<String> = std::env::args().collect();

    match args.get(1).map(|s| s.as_str()) {
        Some("docx-to-pdf") => {
            let input = args.get(2).expect("Usage: oxi docx-to-pdf <input.docx> <output.pdf>");
            let output = args.get(3).expect("Usage: oxi docx-to-pdf <input.docx> <output.pdf>");
            docx_to_pdf(input, output);
        }
        _ => {
            eprintln!("Oxi CLI - Document processing toolkit");
            eprintln!();
            eprintln!("Usage:");
            eprintln!("  oxi docx-to-pdf <input.docx> <output.pdf>");
            std::process::exit(1);
        }
    }
}

fn docx_to_pdf(input: &str, output: &str) {
    let data = fs::read(input).expect("Failed to read input file");
    let doc = oxidocs_core::parse_docx(&data).expect("Failed to parse docx");

    eprintln!("Parsed: {} section(s), {} blocks total",
        doc.pages.len(),
        doc.pages.iter().map(|p| p.blocks.len()).sum::<usize>());

    let pdf_doc = doc_to_pdf(&doc);
    let pdf_bytes = oxipdf_core::write_pdf(&pdf_doc);

    fs::write(output, &pdf_bytes).expect("Failed to write output file");
    println!("Converted {} -> {} ({} bytes, {} pages)",
        input, output, pdf_bytes.len(), pdf_doc.pages.len());
}

// --- Direct Document IR → PDF conversion (no LayoutEngine) ---

struct PdfBuilder {
    pages: Vec<Page>,
    current_contents: Vec<ContentElement>,
    page_width: f64,
    page_height: f64,
    margin_top: f64,
    margin_bottom: f64,
    margin_left: f64,
    margin_right: f64,
    cursor_y: f64,
    /// Unicode → width in 1/1000 em units (from embedded font).
    /// Used for accurate text width estimation.
    font_widths: HashMap<u32, u16>,
}

impl PdfBuilder {
    /// Create a new PdfBuilder.
    /// Coordinate system: y=0 at TOP of page (top-down, matching writer's expectation).
    /// cursor_y starts at margin_top and increases downward.
    fn new(width: f64, height: f64, margin: (f64, f64, f64, f64), font_widths: HashMap<u32, u16>) -> Self {
        Self {
            pages: Vec::new(),
            current_contents: Vec::new(),
            page_width: width,
            page_height: height,
            margin_top: margin.0,
            margin_bottom: margin.1,
            margin_left: margin.2,
            margin_right: margin.3,
            cursor_y: margin.0,
            font_widths,
        }
    }

    /// Estimate text width using actual font metrics when available.
    fn text_width(&self, text: &str, font_size: f64) -> f64 {
        if self.font_widths.is_empty() {
            return estimate_text_width(text, font_size);
        }
        let mut w = 0.0;
        for ch in text.chars() {
            if let Some(&glyph_w) = self.font_widths.get(&(ch as u32)) {
                // glyph_w is in 1/1000 em units
                w += font_size * (glyph_w as f64 / 1000.0);
            } else {
                w += char_width(ch, font_size);
            }
        }
        w
    }

    /// Wrap text using actual font metrics.
    fn wrap_text_accurate(&self, text: &str, font_size: f64, max_width: f64) -> Vec<String> {
        if self.font_widths.is_empty() {
            return wrap_text(text, font_size, max_width);
        }
        let mut lines = Vec::new();
        let mut current_line = String::new();
        let mut current_width = 0.0;

        for ch in text.chars() {
            let ch_w = if let Some(&glyph_w) = self.font_widths.get(&(ch as u32)) {
                font_size * (glyph_w as f64 / 1000.0)
            } else {
                char_width(ch, font_size)
            };

            if current_width + ch_w > max_width && !current_line.is_empty() {
                lines.push(current_line);
                current_line = String::new();
                current_width = 0.0;
            }

            current_line.push(ch);
            current_width += ch_w;
        }

        if !current_line.is_empty() {
            lines.push(current_line);
        }
        if lines.is_empty() {
            lines.push(String::new());
        }
        lines
    }

    fn content_width(&self) -> f64 {
        self.page_width - self.margin_left - self.margin_right
    }

    fn max_y(&self) -> f64 {
        self.page_height - self.margin_bottom
    }

    fn needs_page_break(&self, needed_height: f64) -> bool {
        self.cursor_y + needed_height > self.max_y()
    }

    fn new_page(&mut self) {
        let page = Page {
            width: self.page_width,
            height: self.page_height,
            media_box: Rectangle {
                llx: 0.0, lly: 0.0,
                urx: self.page_width, ury: self.page_height,
            },
            crop_box: None,
            contents: std::mem::take(&mut self.current_contents),
            rotation: 0,
        };
        self.pages.push(page);
        self.cursor_y = self.margin_top;
    }

    fn finish(mut self) -> Vec<Page> {
        if !self.current_contents.is_empty() {
            let page = Page {
                width: self.page_width,
                height: self.page_height,
                media_box: Rectangle {
                    llx: 0.0, lly: 0.0,
                    urx: self.page_width, ury: self.page_height,
                },
                crop_box: None,
                contents: std::mem::take(&mut self.current_contents),
                rotation: 0,
            };
            self.pages.push(page);
        }
        self.pages
    }

    /// Add text at (x, y) where y is in top-down coordinates.
    fn add_text(&mut self, x: f64, y: f64, text: String, font_name: String, font_size: f64, color: Color) {
        self.current_contents.push(ContentElement::Text(TextSpan {
            x, y, text, font_name, font_size, fill_color: color, character_spacing: 0.0,
        }));
    }

    fn add_line(&mut self, x1: f64, y1: f64, x2: f64, y2: f64, width: f64, color: Color) {
        self.current_contents.push(ContentElement::Path(PathData {
            operations: vec![
                PathOp::MoveTo(x1, y1),
                PathOp::LineTo(x2, y2),
            ],
            stroke: Some(StrokeStyle {
                color,
                width,
                line_cap: LineCap::Butt,
                line_join: LineJoin::Miter,
            }),
            fill: None,
        }));
    }

    fn add_rect_fill(&mut self, x: f64, y: f64, w: f64, h: f64, color: Color) {
        self.current_contents.push(ContentElement::Path(PathData {
            operations: vec![
                PathOp::MoveTo(x, y),
                PathOp::LineTo(x + w, y),
                PathOp::LineTo(x + w, y + h),
                PathOp::LineTo(x, y + h),
                PathOp::ClosePath,
            ],
            stroke: None,
            fill: Some(color),
        }));
    }

    fn add_rounded_rect(&mut self, x: f64, y: f64, w: f64, h: f64, r: f64, fill: Option<Color>, stroke: Option<(f64, Color)>) {
        // Bézier control point factor for circular arcs: 4*(sqrt(2)-1)/3 ≈ 0.5523
        let k = 0.5523 * r;
        let ops = vec![
            PathOp::MoveTo(x + r, y),
            PathOp::LineTo(x + w - r, y),
            PathOp::CurveTo(x + w - r + k, y, x + w, y + r - k, x + w, y + r),
            PathOp::LineTo(x + w, y + h - r),
            PathOp::CurveTo(x + w, y + h - r + k, x + w - r + k, y + h, x + w - r, y + h),
            PathOp::LineTo(x + r, y + h),
            PathOp::CurveTo(x + r - k, y + h, x, y + h - r + k, x, y + h - r),
            PathOp::LineTo(x, y + r),
            PathOp::CurveTo(x, y + r - k, x + r - k, y, x + r, y),
            PathOp::ClosePath,
        ];
        self.current_contents.push(ContentElement::Path(PathData {
            operations: ops,
            stroke: stroke.map(|(sw, sc)| StrokeStyle {
                color: sc, width: sw, line_cap: LineCap::Butt, line_join: LineJoin::Round,
            }),
            fill,
        }));
    }

    fn add_rect_stroke(&mut self, x: f64, y: f64, w: f64, h: f64, line_w: f64, color: Color) {
        self.current_contents.push(ContentElement::Path(PathData {
            operations: vec![
                PathOp::MoveTo(x, y),
                PathOp::LineTo(x + w, y),
                PathOp::LineTo(x + w, y + h),
                PathOp::LineTo(x, y + h),
                PathOp::ClosePath,
            ],
            stroke: Some(StrokeStyle {
                color,
                width: line_w,
                line_cap: LineCap::Butt,
                line_join: LineJoin::Miter,
            }),
            fill: None,
        }));
    }
}

/// Render header/footer blocks into a page at specified y position.
fn render_header_footer_blocks(
    builder: &mut PdfBuilder,
    blocks: &[Block],
    y_start: f64,
    registry: &FontMetricsRegistry,
    doc_default_font: Option<&str>,
) {
    let saved_y = builder.cursor_y;
    builder.cursor_y = y_start;
    let ml = builder.margin_left;
    let cw = builder.content_width();

    for block in blocks {
        if let Block::Paragraph(para) = block {
            let text: String = para.runs.iter().map(|r| r.text.as_str()).collect();
            if text.trim().is_empty() { continue; }

            let (font_size, font_family, bold) = para_font_props(para);
            let font_name = resolve_font(font_family, bold);
            let color = para_color(para);
            let text_w = builder.text_width(&text, font_size);
            let x = match para.alignment {
                Alignment::Right => ml + cw - text_w,
                Alignment::Center => ml + (cw - text_w) / 2.0,
                _ => ml,
            };
            let lh = compute_line_height(
                registry, font_family, font_size as f32,
                para.style.line_spacing, para.style.line_spacing_rule.as_deref(),
                doc_default_font, 0.0,
            );
            builder.cursor_y += lh;
            add_text_split(builder, x, builder.cursor_y, text, font_name, font_size, color);
        }
    }
    builder.cursor_y = saved_y;
}

/// Render paragraph shading (background fill).
fn render_paragraph_shading(
    builder: &mut PdfBuilder,
    shading: &str,
    x: f64, y: f64, w: f64, h: f64,
) {
    if let Some(color) = parse_hex_color(shading) {
        builder.add_rect_fill(x, y - 1.0, w, h + 2.0, color);
    }
}

/// Render paragraph borders.
fn render_paragraph_borders(
    builder: &mut PdfBuilder,
    borders: &oxidocs_core::ir::ParagraphBorders,
    x: f64, y: f64, w: f64, h: f64,
) {
    let draw_border = |builder: &mut PdfBuilder, border: &Option<oxidocs_core::ir::BorderDef>, x1: f64, y1: f64, x2: f64, y2: f64| {
        if let Some(bd) = border {
            let bw = if bd.width > 0.0 { bd.width as f64 } else { 0.5 };
            let color = bd.color.as_deref()
                .and_then(parse_hex_color)
                .unwrap_or(Color::Gray(0.0));
            builder.add_line(x1, y1, x2, y2, bw, color);
        }
    };
    draw_border(builder, &borders.top, x, y, x + w, y);
    draw_border(builder, &borders.bottom, x, y + h, x + w, y + h);
    draw_border(builder, &borders.left, x, y, x, y + h);
    draw_border(builder, &borders.right, x + w, y, x + w, y + h);
}

/// Resolve TextBox absolute position (matches layout/mod.rs resolve_textbox_position).
fn resolve_textbox_position(
    text_box: &TextBox,
    section: &oxidocs_core::ir::Page,
    block_y_positions: &[f64],
) -> (f64, f64) {
    let pos = match &text_box.position {
        Some(p) => p,
        None => return (section.margin.left as f64, section.margin.top as f64),
    };

    let page_w = section.size.width as f64;
    let page_h = section.size.height as f64;
    let ml = section.margin.left as f64;
    let mr = section.margin.right as f64;
    let mt = section.margin.top as f64;
    let mb = section.margin.bottom as f64;
    let content_w = page_w - ml - mr;
    let tb_w = text_box.width as f64;
    let tb_h = text_box.height as f64;

    // Horizontal position
    let abs_x = if let Some(ref align) = pos.h_align {
        // Determine reference frame from h_relative
        let (ref_left, ref_w) = match pos.h_relative.as_deref() {
            Some("page") => (0.0, page_w),
            Some("margin") | Some("column") | None => (ml, content_w),
            Some("leftMarginArea") => (0.0, ml),
            Some("rightMarginArea") => (page_w - mr, mr),
            _ => (ml, content_w),
        };
        match align.as_str() {
            "left" => ref_left,
            "center" => ref_left + (ref_w - tb_w) / 2.0,
            "right" => ref_left + ref_w - tb_w,
            _ => ref_left + pos.x as f64,
        }
    } else {
        let offset = pos.x as f64;
        match pos.h_relative.as_deref() {
            Some("page") => offset,
            Some("margin") | Some("column") | None => ml + offset,
            Some("leftMarginArea") => offset,
            Some("rightMarginArea") => (page_w - mr) + offset,
            _ => ml + offset,
        }
    };

    // Vertical position
    let abs_y = if let Some(ref align) = pos.v_align {
        let (ref_top, ref_h) = match pos.v_relative.as_deref() {
            Some("page") => (0.0, page_h),
            Some("margin") | None => (mt, page_h - mt - mb),
            _ => (mt, page_h - mt - mb),
        };
        match align.as_str() {
            "top" => ref_top,
            "center" => ref_top + (ref_h - tb_h) / 2.0,
            "bottom" => ref_top + ref_h - tb_h,
            _ => ref_top + pos.y as f64,
        }
    } else {
        let offset = pos.y as f64;
        match pos.v_relative.as_deref() {
            Some("page") => offset,
            Some("margin") | None => mt + offset,
            Some("paragraph") | Some("line") => {
                let anchor_y = block_y_positions
                    .get(text_box.anchor_block_index)
                    .copied()
                    .unwrap_or(mt);
                anchor_y + offset
            }
            Some("topMarginArea") => offset,
            Some("bottomMarginArea") => (page_h - mb) + offset,
            _ => mt + offset,
        }
    };

    (abs_x, abs_y)
}

/// Render a TextBox: background fill, border, and inner content blocks.
fn render_textbox(
    builder: &mut PdfBuilder,
    text_box: &TextBox,
    section: &oxidocs_core::ir::Page,
    block_y_positions: &[f64],
    registry: &FontMetricsRegistry,
    doc_default_font: Option<&str>,
    doc_default_font_size: f32,
    grid_pitch_pt: f64,
) {
    let (abs_x, abs_y) = resolve_textbox_position(text_box, section, block_y_positions);
    let tb_w = text_box.width as f64;
    let tb_h = text_box.height as f64;

    let cr = text_box.corner_radius.map(|r| r as f64).unwrap_or(0.0);
    let has_rounded = cr > 0.0;

    // Background fill + Border
    if has_rounded {
        let fill_color = text_box.fill.as_ref().and_then(|h| parse_hex_color(h));
        let stroke = if text_box.border { Some((0.5, Color::Gray(0.0))) } else { None };
        if fill_color.is_some() || stroke.is_some() {
            builder.add_rounded_rect(abs_x, abs_y, tb_w, tb_h, cr, fill_color, stroke);
        }
    } else {
        if let Some(ref fill_hex) = text_box.fill {
            if let Some(color) = parse_hex_color(fill_hex) {
                builder.add_rect_fill(abs_x, abs_y, tb_w, tb_h, color);
            }
        }
        if text_box.border {
            let bc = Color::Gray(0.0);
            let bw = 0.5;
            builder.add_line(abs_x, abs_y, abs_x + tb_w, abs_y, bw, bc);
            builder.add_line(abs_x, abs_y + tb_h, abs_x + tb_w, abs_y + tb_h, bw, bc);
            builder.add_line(abs_x, abs_y, abs_x, abs_y + tb_h, bw, bc);
            builder.add_line(abs_x + tb_w, abs_y, abs_x + tb_w, abs_y + tb_h, bw, bc);
        }
    }

    // Inner content with inset (Word default: L/R=7.2pt, T/B=3.6pt)
    let inset_lr = 7.2;
    let inset_tb = 3.6;
    let inner_x = abs_x + inset_lr;
    let inner_w = tb_w - inset_lr * 2.0;
    let saved_y = builder.cursor_y;
    builder.cursor_y = abs_y + inset_tb;

    for block in &text_box.blocks {
        match block {
            Block::Paragraph(para) => {
                let gp = if para.style.snap_to_grid { grid_pitch_pt } else { 0.0 };
                render_paragraph_styled(builder, para, ParaRole::Body, inner_x, inner_w, registry, doc_default_font, doc_default_font_size, gp);
            }
            Block::Table(table) => {
                render_table(builder, table, inner_x, inner_w, registry, doc_default_font);
            }
            _ => {}
        }
    }

    builder.cursor_y = saved_y;
}

/// Render footnotes at the bottom of the current page.
fn render_footnotes(
    builder: &mut PdfBuilder,
    footnotes: &[oxidocs_core::ir::Footnote],
    registry: &FontMetricsRegistry,
    doc_default_font: Option<&str>,
) {
    if footnotes.is_empty() { return; }
    let ml = builder.margin_left;
    let cw = builder.content_width();
    let footnote_font_size = 8.0_f64;

    // Separator line
    let sep_y = builder.max_y() - (footnotes.len() as f64 * footnote_font_size * 1.4) - 10.0;
    builder.add_line(ml, sep_y, ml + cw * 0.3, sep_y, 0.5, Color::Gray(0.3));

    let mut fy = sep_y + 4.0;
    for note in footnotes {
        let text: String = note.blocks.iter().filter_map(|b| {
            if let Block::Paragraph(p) = b {
                Some(p.runs.iter().map(|r| r.text.as_str()).collect::<String>())
            } else { None }
        }).collect::<Vec<_>>().join(" ");

        let marker = format!("{}  ", note.number);
        let full = format!("{}{}", marker, text);
        add_text_split(builder, ml, fy, full, "Helvetica".to_string(), footnote_font_size, Color::Gray(0.2));
        fy += footnote_font_size * 1.4;
    }
}

/// Compute line height following Word's layout behavior.
///
/// Line spacing rules (ECMA-376 17.3.1.13 + Word COM measurements):
/// - auto (multiple): base_height * multiplier
/// - atLeast: max(base_height, specified_pts)
/// - exact: specified_pts exactly
///
/// base_height = max(run_line_height, default_font_line_height)
/// Compute line height following Word's layout behavior.
///
/// grid_pitch_pt: document grid pitch in points (default 18pt = 360 twips).
///   Set to 0.0 for table cells (DontSnapToGridInCell).
fn compute_line_height(
    registry: &FontMetricsRegistry,
    font_family: Option<&str>,
    font_size: f32,
    line_spacing: Option<f32>,
    line_spacing_rule: Option<&str>,
    doc_default_font: Option<&str>,
    grid_pitch_pt: f64,
) -> f64 {
    // grid_pitch_pt > 0 implies snap_to_grid=true
    let snap = grid_pitch_pt > 0.01;
    compute_line_height_with_default(registry, font_family, font_size, line_spacing, line_spacing_rule, doc_default_font, None, grid_pitch_pt, snap)
}

fn compute_line_height_with_default(
    registry: &FontMetricsRegistry,
    font_family: Option<&str>,
    font_size: f32,
    line_spacing: Option<f32>,
    line_spacing_rule: Option<&str>,
    doc_default_font: Option<&str>,
    doc_default_font_size: Option<f32>,
    grid_pitch_pt: f64,
    snap_to_grid: bool,
) -> f64 {
    let run_metrics = font_family
        .map(|f| registry.get(f))
        .unwrap_or(registry.default_metrics());
    let run_base = run_metrics.word_line_height(font_size, 96.0) as f64;

    // snap_to_grid=false uses natural font height only,
    // NO default font minimum. snap_to_grid=true applies max(run, default).
    let base_height = if snap_to_grid {
        let default_metrics = doc_default_font
            .map(|f| registry.get(f))
            .unwrap_or(registry.default_metrics());
        let default_fs = doc_default_font_size.unwrap_or(10.5);
        let default_base = default_metrics.word_line_height(default_fs, 96.0) as f64;
        run_base.max(default_base)
    } else {
        run_base
    };

    let raw = match line_spacing_rule {
        Some("exact") => {
            // Exact: use the specified value in points directly, no grid snap
            return line_spacing.unwrap_or(font_size) as f64;
        }
        Some("atLeast") => {
            let min_height = line_spacing.unwrap_or(0.0) as f64;
            base_height.max(min_height)
        }
        _ => {
            // Auto/multiple: base_height * multiplier
            let mult = line_spacing.unwrap_or(0.0) as f64;
            if mult > 0.01 && mult < 10.0 {
                base_height * mult
            } else {
                base_height // single spacing (1.0)
            }
        }
    };

    // Grid snap: ceil(raw / grid_pitch) * grid_pitch
    if grid_pitch_pt > 0.01 {
        (raw / grid_pitch_pt).ceil() * grid_pitch_pt
    } else {
        raw
    }
}

/// Estimate text width. CJK chars ≈ font_size, Latin width varies by character class.
fn estimate_text_width(text: &str, font_size: f64) -> f64 {
    let mut w = 0.0;
    for ch in text.chars() {
        if is_cjk(ch) {
            w += font_size;
        } else if ch == ' ' {
            w += font_size * 0.25;
        } else if ch == '\u{3000}' {
            // Ideographic space
            w += font_size;
        } else if ch.is_ascii_uppercase() {
            w += font_size * 0.60;
        } else if ch.is_ascii_lowercase() {
            w += font_size * 0.48;
        } else if ch.is_ascii_digit() {
            w += font_size * 0.50;
        } else if ch == '.' || ch == ',' || ch == ':' || ch == ';' || ch == '!' || ch == '\'' {
            w += font_size * 0.25;
        } else if ch == '(' || ch == ')' || ch == '[' || ch == ']' {
            w += font_size * 0.30;
        } else if ch == '—' || ch == '–' || ch == '〜' {
            w += font_size * 0.60;
        } else {
            w += font_size * 0.50;
        }
    }
    w
}

fn is_cjk(ch: char) -> bool {
    let c = ch as u32;
    // CJK Unified, Hiragana, Katakana, fullwidth punctuation, etc.
    (0x3000..=0x9FFF).contains(&c)
        || (0xF900..=0xFAFF).contains(&c)
        || (0xFF00..=0xFFEF).contains(&c)
        || (0x20000..=0x2FA1F).contains(&c)
        || ch == '①' || ch == '②' || ch == '③' || ch == '④' || ch == '⑤'
        || ch == '⑥' || ch == '⑦' || ch == '⑧' || ch == '⑨' || ch == '⑩'
        || ch == '→' || ch == '◆' || ch == '■' || ch == '●' || ch == '◎'
}

/// Break text into lines that fit within max_width
fn wrap_text(text: &str, font_size: f64, max_width: f64) -> Vec<String> {
    if text.is_empty() {
        return vec![String::new()];
    }

    let mut lines = Vec::new();
    let mut current_line = String::new();
    let mut current_width = 0.0;

    for ch in text.chars() {
        let ch_width = char_width(ch, font_size);

        if current_width + ch_width > max_width && !current_line.is_empty() {
            lines.push(current_line);
            current_line = String::new();
            current_width = 0.0;
        }

        current_line.push(ch);
        current_width += ch_width;
    }

    if !current_line.is_empty() {
        lines.push(current_line);
    }

    if lines.is_empty() {
        lines.push(String::new());
    }
    lines
}

/// Per-character width estimate for line wrapping
fn char_width(ch: char, font_size: f64) -> f64 {
    if is_cjk(ch) {
        font_size
    } else if ch == ' ' {
        font_size * 0.25
    } else if ch == '\u{3000}' {
        font_size
    } else if ch.is_ascii_uppercase() {
        font_size * 0.60
    } else if ch.is_ascii_lowercase() {
        font_size * 0.48
    } else if ch.is_ascii_digit() {
        font_size * 0.50
    } else if ch == '.' || ch == ',' || ch == ':' || ch == ';' || ch == '\'' {
        font_size * 0.25
    } else if ch == '(' || ch == ')' || ch == '[' || ch == ']' {
        font_size * 0.30
    } else {
        font_size * 0.50
    }
}

fn resolve_font(family: Option<&str>, bold: bool) -> String {
    let base = family.unwrap_or("Calibri");
    // CJK fonts: use bold variant when bold is requested
    let is_cjk = matches!(base,
        "ＭＳ ゴシック" | "MS ゴシック" | "MS Gothic" |
        "ＭＳ Ｐゴシック" | "MS PGothic" |
        "ＭＳ 明朝" | "MS 明朝" | "MS Mincho" |
        "ＭＳ Ｐ明朝" | "MS PMincho"
    ) || base.contains("Gothic") || base.contains("Mincho") || base.contains("游")
        || base.contains("ＭＳ") || base.contains("ゴシック") || base.contains("明朝")
        || base.contains("メイリオ") || base.contains("ヒラギノ");

    if is_cjk {
        if bold {
            return "OxiCJK-Bold".to_string();
        }
        return match base {
            "ＭＳ ゴシック" | "MS ゴシック" | "MS Gothic" => "MS-Gothic".to_string(),
            "ＭＳ Ｐゴシック" | "MS PGothic" => "MS-PGothic".to_string(),
            "ＭＳ 明朝" | "MS 明朝" | "MS Mincho" => "MS-Mincho".to_string(),
            "ＭＳ Ｐ明朝" | "MS PMincho" => "MS-PMincho".to_string(),
            _ => base.to_string(),
        };
    }
    // Use embedded Latin fonts for Calibri/Arial (matching LayoutEngine metrics)
    if bold {
        match base {
            "Calibri" | "Arial" | "Helvetica" => "OxiLatin-Bold".to_string(),
            "Cambria" | "Century" => "OxiCambria-Bold".to_string(),
            "Times New Roman" | "Times" => "Times-Bold".to_string(),
            "Courier New" | "Courier" => "Courier-Bold".to_string(),
            other => other.to_string(),
        }
    } else {
        match base {
            "Calibri" | "Arial" | "Helvetica" => "OxiLatin-Regular".to_string(),
            "Cambria" | "Century" => "OxiCambria-Regular".to_string(),
            "Times New Roman" => "Times-Roman".to_string(),
            "Courier New" => "Courier".to_string(),
            other => other.to_string(),
        }
    }
}

fn parse_hex_color(hex: &str) -> Option<Color> {
    let hex = hex.strip_prefix('#').unwrap_or(hex);
    if hex.len() != 6 { return None; }
    let r = u8::from_str_radix(&hex[0..2], 16).ok()? as f64 / 255.0;
    let g = u8::from_str_radix(&hex[2..4], 16).ok()? as f64 / 255.0;
    let b = u8::from_str_radix(&hex[4..6], 16).ok()? as f64 / 255.0;
    Some(Color::Rgb(r, g, b))
}

fn parse_highlight_color(name: &str) -> Option<Color> {
    match name {
        "yellow" => Some(Color::Rgb(1.0, 1.0, 0.0)),
        "green" => Some(Color::Rgb(0.0, 1.0, 0.0)),
        "cyan" => Some(Color::Rgb(0.0, 1.0, 1.0)),
        "magenta" => Some(Color::Rgb(1.0, 0.0, 1.0)),
        "blue" => Some(Color::Rgb(0.0, 0.0, 1.0)),
        "red" => Some(Color::Rgb(1.0, 0.0, 0.0)),
        "darkBlue" => Some(Color::Rgb(0.0, 0.0, 0.55)),
        "darkCyan" => Some(Color::Rgb(0.0, 0.55, 0.55)),
        "darkGreen" => Some(Color::Rgb(0.0, 0.39, 0.0)),
        "darkMagenta" => Some(Color::Rgb(0.55, 0.0, 0.55)),
        "darkRed" => Some(Color::Rgb(0.55, 0.0, 0.0)),
        "darkYellow" => Some(Color::Rgb(0.55, 0.55, 0.0)),
        "darkGray" | "darkGrey" => Some(Color::Rgb(0.66, 0.66, 0.66)),
        "lightGray" | "lightGrey" => Some(Color::Rgb(0.85, 0.85, 0.85)),
        "black" => Some(Color::Gray(0.0)),
        _ => parse_hex_color(name),
    }
}

/// Detect paragraph "role" from text patterns and run styling
#[derive(Debug, Clone, Copy, PartialEq)]
enum ParaRole {
    CoverTitle,      // First centered paragraph with large font
    CoverSubtitle,   // Second centered paragraph
    CoverAuthor,     // Third centered paragraph (author/date)
    SectionHeader,   // ① ② ③ ... ⑩ section headers
    SubHeader,       // "N-N." sub-headers
    SubSubHeader,    // Bold paragraph with spB>=200 and smaller font (topic intro)
    Body,            // Regular body text
    ListItem,        // Bulleted/numbered list
    CreditFooter,    // Last right-aligned "本提案書の..." line
}

fn classify_paragraph(para: &Paragraph, block_index: usize, _total_blocks: usize) -> ParaRole {
    let text: String = para.runs.iter().map(|r| r.text.as_str()).collect();
    let trimmed = text.trim();

    // Credit footer (last non-empty paragraph, right-aligned)
    if para.alignment == Alignment::Right && trimmed.contains("本提案書") {
        return ParaRole::CreditFooter;
    }

    // Cover page: first 3 centered paragraphs
    if para.alignment == Alignment::Center && block_index <= 2 {
        let font_size = para.runs.first()
            .and_then(|r| r.style.font_size)
            .unwrap_or(10.5);
        if block_index == 0 && font_size >= 14.0 {
            return ParaRole::CoverTitle;
        }
        if block_index == 1 {
            return ParaRole::CoverSubtitle;
        }
        if block_index == 2 {
            return ParaRole::CoverAuthor;
        }
    }

    // Section headers: starts with ① ② ③ etc.
    if trimmed.starts_with('①') || trimmed.starts_with('②') || trimmed.starts_with('③')
        || trimmed.starts_with('④') || trimmed.starts_with('⑤') || trimmed.starts_with('⑥')
        || trimmed.starts_with('⑦') || trimmed.starts_with('⑧') || trimmed.starts_with('⑨')
        || trimmed.starts_with('⑩')
    {
        return ParaRole::SectionHeader;
    }

    // Sub-headers: patterns like "1-1.", "2-2.", "3-1." etc.
    let is_sub_header = {
        let chars: Vec<char> = trimmed.chars().collect();
        chars.len() >= 4
            && chars[0].is_ascii_digit()
            && chars[1] == '-'
            && chars[2].is_ascii_digit()
            && chars[3] == '.'
    };
    if is_sub_header {
        return ParaRole::SubHeader;
    }

    // Sub-sub-headers: bold text with space_before >= 150, font size ~11pt, not too long
    let font_size = para.runs.first()
        .and_then(|r| r.style.font_size)
        .or(para.style.default_run_style.as_ref().and_then(|rs| rs.font_size))
        .unwrap_or(10.5);
    let is_bold = para.runs.first().map(|r| r.style.bold).unwrap_or(false);
    let space_before = para.style.space_before.unwrap_or(0.0);

    if is_bold && space_before >= 10.0 && font_size >= 10.5 && trimmed.len() < 120
        && !trimmed.starts_with('・') && !trimmed.starts_with('【')
        && para.alignment != Alignment::Center
        && para.style.list_marker.is_none()
    {
        return ParaRole::SubSubHeader;
    }

    // List items
    if para.style.list_marker.is_some() || trimmed.starts_with('・') || trimmed.starts_with('※') {
        return ParaRole::ListItem;
    }

    ParaRole::Body
}

fn doc_to_pdf(doc: &oxidocs_core::Document) -> PdfDocument {
    // Use LayoutEngine for layout (single source of truth with WASM path)
    let engine = oxidocs_core::layout::LayoutEngine::for_document(doc);
    let layout_result = engine.layout(doc);

    // Font metrics for accurate baseline calculation
    let registry = FontMetricsRegistry::load();

    // Build embedded fonts for CJK text
    let embedded_fonts = build_embedded_fonts(doc);

    // Convert LayoutResult pages → PDF pages
    let pages: Vec<Page> = layout_result.pages.iter().map(|lp| {
        let w = lp.width as f64;
        let h = lp.height as f64;
        let mut contents: Vec<ContentElement> = Vec::new();

        for elem in &lp.elements {
            let x = elem.x as f64;
            let y = elem.y as f64;
            let ew = elem.width as f64;
            let eh = elem.height as f64;

            match &elem.content {
                oxidocs_core::layout::LayoutContent::Text {
                    text, font_size, font_family, bold, italic,
                    underline, underline_style, strikethrough, color, highlight,
                    character_spacing, ..
                } => {
                    let fs = *font_size as f64;
                    let fill_color = color.as_deref()
                        .and_then(parse_hex_color)
                        .unwrap_or(Color::Gray(0.0));

                    // Highlight background
                    if let Some(ref hl) = highlight {
                        if let Some(hl_color) = parse_highlight_color(hl) {
                            contents.push(ContentElement::Path(PathData {
                                operations: vec![
                                    PathOp::MoveTo(x, y),
                                    PathOp::LineTo(x + ew, y),
                                    PathOp::LineTo(x + ew, y + eh),
                                    PathOp::LineTo(x, y + eh),
                                    PathOp::ClosePath,
                                ],
                                stroke: None,
                                fill: Some(hl_color),
                            }));
                        }
                    }

                    // Resolve font name for PDF (map to Type1 base fonts or CJK)
                    let font_name = resolve_font(font_family.as_deref(), *bold);

                    // Split text into ASCII (Type1) and CJK (embedded) segments
                    // When font_name IS a CJK font, use it for all characters (no splitting).
                    let is_cjk_font_name = font_name.contains("Gothic") || font_name.contains("Mincho") || font_name.contains("ゴシック") || font_name.contains("明朝")
                        || font_name.contains("游") || font_name.contains("ＭＳ")
                        || font_name.contains("ゴシック") || font_name.contains("明朝")
                        || font_name.contains("メイリオ") || font_name.contains("Meiryo");
                    let chars: Vec<char> = text.chars().collect();
                    let mut seg_x = x;
                    let mut segments: Vec<(String, String, f64, f64)> = Vec::new();

                    let mut raw_segments: Vec<(String, String, f64)> = Vec::new();
                    let mut total_est = 0.0_f64;
                    let mut i = 0;
                    while i < chars.len() {
                        let is_cjk = chars[i] as u32 > 0x7F;
                        let mut j = i + 1;
                        // If using a CJK font, don't split — keep everything in one segment
                        if !is_cjk_font_name {
                            while j < chars.len() && (chars[j] as u32 > 0x7F) == is_cjk {
                                j += 1;
                            }
                        } else {
                            j = chars.len(); // entire text as one segment
                        }
                        let seg_text: String = chars[i..j].iter().collect();
                        let seg_font = if is_cjk_font_name {
                            font_name.clone()
                        } else if is_cjk {
                            if *bold { "OxiCJK-Bold".to_string() } else { "OxiCJK-Regular".to_string() }
                        } else { font_name.clone() };
                        let seg_w: f64 = seg_text.chars()
                            .map(|c| char_width(c, fs))
                            .sum();
                        total_est += seg_w;
                        raw_segments.push((seg_text, seg_font, seg_w));
                        i = j;
                    }

                    // Second pass: scale segment widths to match LayoutEngine's elem.width
                    let scale = if total_est > 0.0 { ew / total_est } else { 1.0 };
                    for (seg_text, seg_font, seg_w) in raw_segments {
                        let scaled_w = seg_w * scale;
                        segments.push((seg_text, seg_font, seg_x, scaled_w));
                        seg_x += scaled_w;
                    }

                    // Baseline position: y is top of line box, PDF baseline = y + ascent
                    // Matches Word output: baseline = content_top + word_ascent_pt(fontSize)
                    let text_y = {
                        let metrics = font_family.as_deref()
                            .map(|ff| registry.get(ff))
                            .unwrap_or_else(|| registry.default_metrics());
                        y + metrics.word_ascent_pt(fs as f32) as f64
                    };

                    for (seg_text, seg_font, sx, _sw) in segments {
                        contents.push(ContentElement::Text(TextSpan {
                            x: sx, y: text_y,
                            text: seg_text, font_name: seg_font,
                            font_size: fs, fill_color,
                            character_spacing: *character_spacing as f64,
                        }));
                    }

                    // Underline
                    if *underline {
                        let is_double = underline_style.as_deref() == Some("double");
                        let ul_w = 0.5_f64.max(fs * 0.05);
                        if is_double {
                            // Double underline: two lines
                            let ul_y1 = text_y + fs * 0.10;
                            let ul_y2 = text_y + fs * 0.22;
                            for ul_y in [ul_y1, ul_y2] {
                                contents.push(ContentElement::Path(PathData {
                                    operations: vec![
                                        PathOp::MoveTo(x, ul_y),
                                        PathOp::LineTo(seg_x, ul_y),
                                    ],
                                    stroke: Some(StrokeStyle {
                                        color: fill_color, width: ul_w,
                                        line_cap: LineCap::Butt, line_join: LineJoin::Miter,
                                    }),
                                    fill: None,
                                }));
                            }
                        } else {
                            let ul_y = text_y + fs * 0.15;
                            contents.push(ContentElement::Path(PathData {
                                operations: vec![
                                    PathOp::MoveTo(x, ul_y),
                                    PathOp::LineTo(seg_x, ul_y),
                                ],
                                stroke: Some(StrokeStyle {
                                    color: fill_color, width: ul_w,
                                    line_cap: LineCap::Butt, line_join: LineJoin::Miter,
                                }),
                                fill: None,
                            }));
                        }
                    }

                    // Strikethrough
                    if *strikethrough {
                        let st_y = y + eh * 0.45;
                        let st_w = 0.5_f64.max(fs * 0.04);
                        contents.push(ContentElement::Path(PathData {
                            operations: vec![
                                PathOp::MoveTo(x, st_y),
                                PathOp::LineTo(seg_x, st_y),
                            ],
                            stroke: Some(StrokeStyle {
                                color: fill_color, width: st_w,
                                line_cap: LineCap::Butt, line_join: LineJoin::Miter,
                            }),
                            fill: None,
                        }));
                    }
                }

                oxidocs_core::layout::LayoutContent::TableBorder { x1, y1, x2, y2, color, width } => {
                    let bc = color.as_deref()
                        .and_then(parse_hex_color)
                        .unwrap_or(Color::Gray(0.0));
                    let bw = *width as f64;
                    contents.push(ContentElement::Path(PathData {
                        operations: vec![
                            PathOp::MoveTo(*x1 as f64, *y1 as f64),
                            PathOp::LineTo(*x2 as f64, *y2 as f64),
                        ],
                        stroke: Some(StrokeStyle {
                            color: bc, width: bw,
                            line_cap: LineCap::Butt, line_join: LineJoin::Miter,
                        }),
                        fill: None,
                    }));
                }

                oxidocs_core::layout::LayoutContent::CellShading { color: shade_color } => {
                    let sc = parse_hex_color(shade_color).unwrap_or(Color::Gray(0.9));
                    contents.push(ContentElement::Path(PathData {
                        operations: vec![
                            PathOp::MoveTo(x, y),
                            PathOp::LineTo(x + ew, y),
                            PathOp::LineTo(x + ew, y + eh),
                            PathOp::LineTo(x, y + eh),
                            PathOp::ClosePath,
                        ],
                        stroke: None,
                        fill: Some(sc),
                    }));
                }

                oxidocs_core::layout::LayoutContent::BoxRect { fill, stroke_color, stroke_width, corner_radius } => {
                    let fill_color = fill.as_deref().and_then(parse_hex_color);
                    let stroke = stroke_color.as_deref().and_then(parse_hex_color)
                        .map(|c| (*stroke_width as f64, c));
                    let cr = *corner_radius as f64;
                    if cr > 0.0 {
                        // Rounded rectangle
                        let k = 0.5523 * cr;
                        let (rx, ry, rw, rh) = (x, y, ew, eh);
                        let ops = vec![
                            PathOp::MoveTo(rx + cr, ry),
                            PathOp::LineTo(rx + rw - cr, ry),
                            PathOp::CurveTo(rx + rw - cr + k, ry, rx + rw, ry + cr - k, rx + rw, ry + cr),
                            PathOp::LineTo(rx + rw, ry + rh - cr),
                            PathOp::CurveTo(rx + rw, ry + rh - cr + k, rx + rw - cr + k, ry + rh, rx + rw - cr, ry + rh),
                            PathOp::LineTo(rx + cr, ry + rh),
                            PathOp::CurveTo(rx + cr - k, ry + rh, rx, ry + rh - cr + k, rx, ry + rh - cr),
                            PathOp::LineTo(rx, ry + cr),
                            PathOp::CurveTo(rx, ry + cr - k, rx + cr - k, ry, rx + cr, ry),
                            PathOp::ClosePath,
                        ];
                        contents.push(ContentElement::Path(PathData {
                            operations: ops,
                            stroke: stroke.map(|(sw, sc)| StrokeStyle {
                                color: sc, width: sw, line_cap: LineCap::Butt, line_join: LineJoin::Round,
                            }),
                            fill: fill_color,
                        }));
                    } else {
                        // Regular rectangle
                        let ops = vec![
                            PathOp::MoveTo(x, y),
                            PathOp::LineTo(x + ew, y),
                            PathOp::LineTo(x + ew, y + eh),
                            PathOp::LineTo(x, y + eh),
                            PathOp::ClosePath,
                        ];
                        contents.push(ContentElement::Path(PathData {
                            operations: ops,
                            stroke: stroke.map(|(sw, sc)| StrokeStyle {
                                color: sc, width: sw, line_cap: LineCap::Butt, line_join: LineJoin::Miter,
                            }),
                            fill: fill_color,
                        }));
                    }
                }

                oxidocs_core::layout::LayoutContent::Image { ref data, ref content_type } => {
                    // Decode image (PNG/JPEG/etc.) to raw RGB pixels
                    if !data.is_empty() {
                        if let Ok(img) = image::load_from_memory(data) {
                            let rgb = img.to_rgb8();
                            let (pw, ph) = rgb.dimensions();
                            let raw_pixels = rgb.into_raw();
                            contents.push(ContentElement::Image(ImageData {
                                x: elem.x as f64,
                                y: elem.y as f64,
                                width: elem.width as f64,
                                height: elem.height as f64,
                                data: raw_pixels,
                                color_space: ColorSpace::DeviceRgb,
                                bits_per_component: 8,
                                pixel_width: pw,
                                pixel_height: ph,
                            }));
                        }
                    }
                }
                oxidocs_core::layout::LayoutContent::ClipStart => {
                    // Pass coordinates in page-top-down system; the PDF writer's
                    // ClipPath handler flips Y via (page.height - y).
                    let clip_x = elem.x as f64;
                    let clip_y = elem.y as f64;
                    let clip_w = elem.width as f64;
                    let clip_h = elem.height as f64;
                    contents.push(ContentElement::SaveState);
                    contents.push(ContentElement::ClipPath(ClipPathData {
                        operations: vec![
                            PathOp::MoveTo(clip_x, clip_y),
                            PathOp::LineTo(clip_x + clip_w, clip_y),
                            PathOp::LineTo(clip_x + clip_w, clip_y + clip_h),
                            PathOp::LineTo(clip_x, clip_y + clip_h),
                            PathOp::ClosePath,
                        ],
                        even_odd: false,
                    }));
                }
                oxidocs_core::layout::LayoutContent::ClipEnd => {
                    contents.push(ContentElement::RestoreState);
                }
            }
        }

        Page {
            width: w,
            height: h,
            media_box: Rectangle { llx: 0.0, lly: 0.0, urx: w, ury: h },
            crop_box: None,
            contents,
            rotation: 0,
        }
    }).collect();

    let title = doc.pages.first()
        .and_then(|p| p.blocks.first())
        .and_then(|b| match b {
            Block::Paragraph(p) => {
                let text: String = p.runs.iter().map(|r| r.text.as_str()).collect();
                if !text.is_empty() { Some(text) } else { None }
            }
            _ => None,
        });

    PdfDocument {
        version: PdfVersion::new(1, 7),
        info: DocumentInfo {
            title,
            producer: Some("Oxi (oxipdf-core)".to_string()),
            ..Default::default()
        },
        pages,
        outline: Vec::new(),
        embedded_fonts,
    }
}

/// Render cover page elements with professional layout
fn render_cover_element(builder: &mut PdfBuilder, para: &Paragraph, role: ParaRole, x_offset: f64, available_width: f64) {
    let text: String = para.runs.iter().map(|r| r.text.as_str()).collect();
    if text.trim().is_empty() { return; }

    let font_family = para.runs.first()
        .and_then(|r| r.style.font_family.as_deref().or(r.style.font_family_east_asia.as_deref()));
    let color_hex = para.runs.first().and_then(|r| r.style.color.as_deref());
    let color = color_hex.and_then(parse_hex_color).unwrap_or(Color::Rgb(0.18, 0.25, 0.34));

    match role {
        ParaRole::CoverTitle => {
            // Auto-size: try 20pt, reduce until it fits on 1 line (or 2 max)
            let mut font_size = 20.0;
            loop {
                let lines = builder.wrap_text_accurate(&text, font_size, available_width);
                if lines.len() <= 1 || font_size <= 14.0 {
                    break;
                }
                font_size -= 0.5;
            }
            let font_name = resolve_font(font_family, true);
            let line_height = font_size * 1.6;

            // Position title at approximately 1/3 from top of page (top-down coords)
            let target_y = builder.page_height * 0.28;
            if builder.cursor_y < target_y {
                builder.cursor_y = target_y;
            }

            let lines = builder.wrap_text_accurate(&text, font_size, available_width);
            for line in &lines {
                let line_w = builder.text_width(line, font_size);
                let x = x_offset + (available_width - line_w) / 2.0;
                builder.add_text(x, builder.cursor_y, line.clone(), font_name.clone(), font_size, color);
                builder.cursor_y += line_height;
            }

            // Decorative line under title
            builder.cursor_y += 6.0;
            builder.add_line(
                x_offset + available_width * 0.2, builder.cursor_y,
                x_offset + available_width * 0.8, builder.cursor_y,
                1.5, color,
            );
            builder.cursor_y += 20.0;
        }
        ParaRole::CoverSubtitle => {
            let font_size = 12.0;
            let font_name = resolve_font(font_family, true);
            let line_w = builder.text_width(&text, font_size);
            let x = x_offset + (available_width - line_w) / 2.0;
            builder.add_text(x, builder.cursor_y, text, font_name, font_size, Color::Gray(0.3));
            builder.cursor_y += font_size * 3.0;
        }
        ParaRole::CoverAuthor => {
            let font_size = 11.0;
            let font_name = resolve_font(font_family, false);
            let line_w = builder.text_width(&text, font_size);
            let x = x_offset + (available_width - line_w) / 2.0;
            builder.add_text(x, builder.cursor_y, text, font_name, font_size, Color::Gray(0.25));
            builder.cursor_y += font_size * 2.0;
        }
        _ => {}
    }
}

/// Render a paragraph with role-based styling
fn render_paragraph_styled(builder: &mut PdfBuilder, para: &Paragraph, role: ParaRole, x_offset: f64, available_width: f64, registry: &FontMetricsRegistry, doc_default_font: Option<&str>, doc_default_font_size: f32, grid_pitch_pt: f64) {
    let full_text: String = para.runs.iter().map(|r| r.text.as_str()).collect();

    // Empty paragraph — add spacing based on default font metrics
    if full_text.trim().is_empty() && para.style.list_marker.is_none() {
        let default_fs = para.style.default_run_style.as_ref()
            .and_then(|rs| rs.font_size).unwrap_or(10.5);
        let default_ff = para.style.default_run_style.as_ref()
            .and_then(|rs| rs.font_family.as_deref());
        let lh = compute_line_height_with_default(
            registry, default_ff, default_fs,
            para.style.line_spacing, para.style.line_spacing_rule.as_deref(),
            doc_default_font, Some(doc_default_font_size), grid_pitch_pt, grid_pitch_pt > 0.01,
        );
        let gap = para.style.space_before.unwrap_or(0.0) as f64
            + lh
            + para.style.space_after.unwrap_or(0.0) as f64;
        builder.cursor_y += gap;
        return;
    }

    // Determine font size from run styling
    let default_size = para.runs.first()
        .and_then(|r| r.style.font_size)
        .or(para.style.default_run_style.as_ref().and_then(|rs| rs.font_size))
        .unwrap_or(10.5) as f64;

    let default_font = para.runs.first()
        .and_then(|r| r.style.font_family.as_deref().or(r.style.font_family_east_asia.as_deref()))
        .or(para.style.default_run_style.as_ref().and_then(|rs| rs.font_family.as_deref()));

    let _default_color_hex = para.runs.first()
        .and_then(|r| r.style.color.as_deref())
        .or(para.style.default_run_style.as_ref().and_then(|rs| rs.color.as_deref()));

    // Role-based styling overrides
    match role {
        ParaRole::SectionHeader => {
            render_section_header(builder, para, x_offset, available_width, registry, doc_default_font, grid_pitch_pt);
            return;
        }
        ParaRole::SubHeader => {
            render_sub_header(builder, para, x_offset, available_width, registry, doc_default_font, grid_pitch_pt);
            return;
        }
        ParaRole::SubSubHeader => {
            render_sub_sub_header(builder, para, x_offset, available_width, registry, doc_default_font, grid_pitch_pt);
            return;
        }
        ParaRole::CreditFooter => {
            render_credit_footer(builder, para, x_offset, available_width, registry, doc_default_font, grid_pitch_pt);
            return;
        }
        _ => {} // Body and ListItem handled below
    }

    // --- Regular body / list rendering ---
    let font_size = default_size;
    let line_height = compute_line_height_with_default(
        registry,
        default_font,
        font_size as f32,
        para.style.line_spacing,
        para.style.line_spacing_rule.as_deref(),
        doc_default_font,
        Some(doc_default_font_size),
        grid_pitch_pt,
        grid_pitch_pt > 0.01,
    );

    let space_before = para.style.space_before.unwrap_or(0.0) as f64;
    let space_after = para.style.space_after.unwrap_or(0.0) as f64;

    // Indentation
    let indent_left = para.style.indent_left.unwrap_or(0.0) as f64;
    let indent_first = para.style.indent_first_line.unwrap_or(0.0) as f64;

    builder.cursor_y += space_before;

    if builder.needs_page_break(line_height + 10.0) {
        builder.new_page();
    }

    // Paragraph shading (background fill)
    if let Some(ref shd) = para.style.shading {
        let lines_est = builder.wrap_text_accurate(&full_text, font_size, available_width - indent_left).len();
        let para_h = lines_est as f64 * line_height;
        render_paragraph_shading(builder, shd, x_offset, builder.cursor_y, available_width, para_h);
    }

    // Paragraph borders
    if let Some(ref borders) = para.style.borders {
        let lines_est = builder.wrap_text_accurate(&full_text, font_size, available_width - indent_left).len();
        let para_h = lines_est as f64 * line_height;
        render_paragraph_borders(builder, borders, x_offset, builder.cursor_y, available_width, para_h);
    }

    // List marker handling (both IR list_marker and text-based ・)
    let mut marker_offset = 0.0;
    let mut effective_indent_left = indent_left;

    if para.style.list_marker.is_some() || role == ParaRole::ListItem {
        if let Some(marker) = &para.style.list_marker {
            let marker_indent = para.style.list_indent.unwrap_or(18.0) as f64;
            marker_offset = marker_indent;
            let marker_x = x_offset + indent_left;
            let marker_font = resolve_font(default_font, false);
            // Marker baseline matches first text line (cursor_y + line_height)
            builder.add_text(
                marker_x, builder.cursor_y + line_height,
                marker.clone(), marker_font, font_size, Color::Gray(0.0),
            );
        } else if full_text.trim().starts_with('・') || full_text.trim().starts_with('※') {
            effective_indent_left = effective_indent_left.max(8.0);
        }
    }

    let text_x_start = x_offset + effective_indent_left + marker_offset;
    let text_width = (available_width - effective_indent_left - marker_offset).max(50.0);

    let para_default_bold = para.style.default_run_style.as_ref().map(|rs| rs.bold).unwrap_or(false);
    let para_default_ff = para.style.default_run_style.as_ref().and_then(|rs| rs.font_family.as_deref());
    let lines = wrap_runs_into_lines(&para.runs, font_size, text_width, indent_first, &builder.font_widths, para_default_bold, para_default_ff);

    for (line_idx, line_runs) in lines.iter().enumerate() {
        if builder.needs_page_break(line_height) {
            builder.new_page();
        }

        let line_x = if line_idx == 0 {
            text_x_start + indent_first
        } else {
            text_x_start
        };

        let line_width: f64 = line_runs.iter()
            .map(|lr| {
                let base_w = builder.text_width(&lr.text, lr.font_size);
                let spacing_w = lr.char_spacing * lr.text.chars().count() as f64;
                base_w + spacing_w
            })
            .sum();

        let align_offset = match para.alignment {
            Alignment::Center => (text_width - line_width) / 2.0,
            Alignment::Right => text_width - line_width,
            _ => 0.0,
        };

        // Advance cursor FIRST (Word places baseline at cursor + line_height)
        builder.cursor_y += line_height;
        let text_y = builder.cursor_y;

        let mut run_x = line_x + align_offset.max(0.0);

        // Distribute alignment: spread characters evenly across the line
        if para.alignment == Alignment::Distribute && line_width > 0.01 {
            let total_chars: usize = line_runs.iter().map(|lr| lr.text.chars().count()).sum();
            if total_chars > 1 {
                let slack = text_width - line_width;
                let extra_per_gap = slack / (total_chars - 1) as f64;
                for lr in line_runs {
                    for ch in lr.text.chars() {
                        let ch_str = ch.to_string();
                        builder.add_text(
                            run_x, text_y,
                            ch_str.clone(), lr.font_name.clone(), lr.font_size, lr.color,
                        );
                        run_x += builder.text_width(&ch_str, lr.font_size) + extra_per_gap;
                    }
                }
            } else {
                for lr in line_runs {
                    if !lr.text.is_empty() {
                        builder.add_text(
                            run_x, text_y,
                            lr.text.clone(), lr.font_name.clone(), lr.font_size, lr.color,
                        );
                        run_x += builder.text_width(&lr.text, lr.font_size);
                    }
                }
            }
        } else {
            for lr in line_runs {
                if !lr.text.is_empty() {
                    let run_start_x = run_x;
                    if lr.char_spacing.abs() > 0.01 {
                        // Render character by character with extra spacing
                        for ch in lr.text.chars() {
                            let ch_str = ch.to_string();
                            builder.add_text(
                                run_x, text_y,
                                ch_str.clone(), lr.font_name.clone(), lr.font_size, lr.color,
                            );
                            run_x += builder.text_width(&ch_str, lr.font_size) + lr.char_spacing;
                        }
                    } else {
                        builder.add_text(
                            run_x, text_y,
                            lr.text.clone(), lr.font_name.clone(), lr.font_size, lr.color,
                        );
                        run_x += builder.text_width(&lr.text, lr.font_size);
                    }
                    // Draw underline
                    if lr.underline {
                        let ul_y = text_y + lr.font_size * 0.15;
                        let ul_w = 0.5_f64.max(lr.font_size * 0.05);
                        let is_double = lr.underline_style.as_deref() == Some("double");
                        if is_double {
                            // Double underline: two lines with gap
                            let gap = ul_w * 2.0;
                            builder.add_line(run_start_x, ul_y - gap / 2.0, run_x, ul_y - gap / 2.0, ul_w, lr.color);
                            builder.add_line(run_start_x, ul_y + gap / 2.0, run_x, ul_y + gap / 2.0, ul_w, lr.color);
                        } else {
                            builder.add_line(run_start_x, ul_y, run_x, ul_y, ul_w, lr.color);
                        }
                    }
                }
            }
        }
    }

    builder.cursor_y += space_after;
}

/// Extract paragraph font properties
fn para_font_props(para: &Paragraph) -> (f64, Option<&str>, bool) {
    let font_size = para.runs.first()
        .and_then(|r| r.style.font_size)
        .or(para.style.default_run_style.as_ref().and_then(|rs| rs.font_size))
        .unwrap_or(10.5) as f64;
    let font_family = para.runs.first()
        .and_then(|r| r.style.font_family.as_deref().or(r.style.font_family_east_asia.as_deref()))
        .or(para.style.default_run_style.as_ref().and_then(|rs| rs.font_family.as_deref()));
    let bold = para.runs.first().map(|r| r.style.bold).unwrap_or(false)
        || para.style.default_run_style.as_ref().map(|rs| rs.bold).unwrap_or(false);
    (font_size, font_family, bold)
}

fn para_color(para: &Paragraph) -> Color {
    para.runs.first()
        .and_then(|r| r.style.color.as_deref())
        .or(para.style.default_run_style.as_ref().and_then(|rs| rs.color.as_deref()))
        .and_then(parse_hex_color)
        .unwrap_or(Color::Gray(0.0))
}

/// Render ① ② ③ section headers — uses actual paragraph properties
fn render_section_header(builder: &mut PdfBuilder, para: &Paragraph, x_offset: f64, available_width: f64, registry: &FontMetricsRegistry, doc_default_font: Option<&str>, grid_pitch_pt: f64) {
    let text: String = para.runs.iter().map(|r| r.text.as_str()).collect();
    let (font_size, font_family, bold) = para_font_props(para);
    let line_height = compute_line_height(
        registry, font_family, font_size as f32,
        para.style.line_spacing, para.style.line_spacing_rule.as_deref(),
        doc_default_font, grid_pitch_pt,
    );
    let space_before = para.style.space_before.unwrap_or(0.0) as f64;
    let space_after = para.style.space_after.unwrap_or(0.0) as f64;
    let font_name = resolve_font(font_family, bold);
    let color = para_color(para);

    builder.cursor_y += space_before;
    if builder.needs_page_break(line_height + 10.0) { builder.new_page(); }

    // Shading background if present
    if let Some(ref shd) = para.style.shading {
        if let Some(bg) = parse_hex_color(shd) {
            builder.add_rect_fill(x_offset, builder.cursor_y - 1.0, available_width, line_height + 2.0, bg);
        }
    }

    let indent_left = para.style.indent_left.unwrap_or(0.0) as f64;
    add_text_split(builder, x_offset + indent_left, builder.cursor_y, text, font_name, font_size, color);
    builder.cursor_y += line_height + space_after;
}

/// Render "N-N." sub-headers — uses actual paragraph properties
fn render_sub_header(builder: &mut PdfBuilder, para: &Paragraph, x_offset: f64, _available_width: f64, registry: &FontMetricsRegistry, doc_default_font: Option<&str>, grid_pitch_pt: f64) {
    let text: String = para.runs.iter().map(|r| r.text.as_str()).collect();
    let (font_size, font_family, bold) = para_font_props(para);
    let line_height = compute_line_height(
        registry, font_family, font_size as f32,
        para.style.line_spacing, para.style.line_spacing_rule.as_deref(),
        doc_default_font, grid_pitch_pt,
    );
    let space_before = para.style.space_before.unwrap_or(0.0) as f64;
    let space_after = para.style.space_after.unwrap_or(0.0) as f64;
    let font_name = resolve_font(font_family, bold);
    let color = para_color(para);

    builder.cursor_y += space_before;
    if builder.needs_page_break(line_height + 10.0) { builder.new_page(); }

    let indent_left = para.style.indent_left.unwrap_or(0.0) as f64;
    add_text_split(builder, x_offset + indent_left, builder.cursor_y, text, font_name, font_size, color);
    builder.cursor_y += line_height + space_after;
}

/// Render sub-sub-headers — uses actual paragraph properties
fn render_sub_sub_header(builder: &mut PdfBuilder, para: &Paragraph, x_offset: f64, available_width: f64, registry: &FontMetricsRegistry, doc_default_font: Option<&str>, grid_pitch_pt: f64) {
    let text: String = para.runs.iter().map(|r| r.text.as_str()).collect();
    let (font_size, font_family, bold) = para_font_props(para);
    let line_height = compute_line_height(
        registry, font_family, font_size as f32,
        para.style.line_spacing, para.style.line_spacing_rule.as_deref(),
        doc_default_font, grid_pitch_pt,
    );
    let space_before = para.style.space_before.unwrap_or(0.0) as f64;
    let space_after = para.style.space_after.unwrap_or(0.0) as f64;
    let font_name = resolve_font(font_family, bold);
    let color = para_color(para);

    builder.cursor_y += space_before;
    if builder.needs_page_break(line_height + 10.0) { builder.new_page(); }

    let indent_left = para.style.indent_left.unwrap_or(0.0) as f64;
    let lines = wrap_text(&text, font_size, available_width - indent_left);
    for line in &lines {
        add_text_split(builder, x_offset + indent_left, builder.cursor_y, line.clone(), font_name.clone(), font_size, color);
        builder.cursor_y += line_height;
    }
    builder.cursor_y += space_after;
}

/// Render the credit footer line — uses actual paragraph properties
fn render_credit_footer(builder: &mut PdfBuilder, para: &Paragraph, x_offset: f64, available_width: f64, registry: &FontMetricsRegistry, doc_default_font: Option<&str>, grid_pitch_pt: f64) {
    let text: String = para.runs.iter().map(|r| r.text.as_str()).collect();
    let (font_size, font_family, bold) = para_font_props(para);
    let line_height = compute_line_height(
        registry, font_family, font_size as f32,
        para.style.line_spacing, para.style.line_spacing_rule.as_deref(),
        doc_default_font, grid_pitch_pt,
    );
    let space_before = para.style.space_before.unwrap_or(0.0) as f64;
    let space_after = para.style.space_after.unwrap_or(0.0) as f64;
    let font_name = resolve_font(font_family, bold);
    let color = para_color(para);

    builder.cursor_y += space_before;

    let text_w = builder.text_width(&text, font_size);
    let x = match para.alignment {
        Alignment::Right => x_offset + available_width - text_w,
        Alignment::Center => x_offset + (available_width - text_w) / 2.0,
        _ => x_offset,
    };
    add_text_split(builder, x, builder.cursor_y, text, font_name, font_size, color);
    builder.cursor_y += line_height + space_after;
}

/// Resolve the CJK font name for non-ASCII text.
/// If font_name is a known CJK font (MS Gothic, MS Mincho, etc.), use it directly.
/// Otherwise, fall back to "OxiCJK-Regular"/"OxiCJK-Bold".
fn resolve_cjk_font(font_name: &str, bold: bool) -> String {
    // Check if the font_name is already a CJK font that we embed per-family
    if font_name.contains("Gothic") || font_name.contains("Mincho") || font_name.contains("ゴシック") || font_name.contains("明朝")
        || font_name.contains("游") || font_name.contains("ＭＳ")
        || font_name.contains("メイリオ") || font_name.contains("Meiryo")
    {
        return font_name.to_string();
    }
    if bold { "OxiCJK-Bold".to_string() } else { "OxiCJK-Regular".to_string() }
}

/// Emit text with proper ASCII/CJK font splitting.
/// ASCII characters use the given `font_name` (Type1), non-ASCII use the resolved CJK font.
/// When font_name IS a CJK font, all characters use that font (no splitting).
fn add_text_split(builder: &mut PdfBuilder, x: f64, y: f64, text: String, font_name: String, font_size: f64, color: Color) {
    // If font_name is a CJK font, use it for ALL characters (including ASCII)
    let is_cjk_font = font_name.contains("Gothic") || font_name.contains("Mincho") || font_name.contains("ゴシック") || font_name.contains("明朝")
        || font_name.contains("游") || font_name.contains("ＭＳ")
        || font_name.contains("メイリオ") || font_name.contains("Meiryo");
    if is_cjk_font {
        builder.add_text(x, y, text, font_name, font_size, color);
        return;
    }

    if !text.chars().any(|c| c as u32 > 0x7F) {
        // Pure ASCII — use the regular font directly
        builder.add_text(x, y, text, font_name, font_size, color);
        return;
    }

    let cjk_font = resolve_cjk_font(&font_name, font_name.contains("Bold"));

    // Split text into ASCII and non-ASCII segments
    let mut cur_x = x;
    let mut buf = String::new();
    let mut buf_is_cjk = false;

    for ch in text.chars() {
        let ch_is_cjk = ch as u32 > 0x7F;
        if !buf.is_empty() && ch_is_cjk != buf_is_cjk {
            let fn_for_buf = if buf_is_cjk { cjk_font.clone() } else { font_name.clone() };
            let w = builder.text_width(&buf, font_size);
            builder.add_text(cur_x, y, buf.clone(), fn_for_buf, font_size, color);
            cur_x += w;
            buf.clear();
        }
        buf_is_cjk = ch_is_cjk;
        buf.push(ch);
    }
    if !buf.is_empty() {
        let fn_for_buf = if buf_is_cjk { cjk_font } else { font_name };
        builder.add_text(cur_x, y, buf, fn_for_buf, font_size, color);
    }
}

/// A styled text fragment for a single line
struct LineRun {
    text: String,
    font_name: String,
    font_size: f64,
    color: Color,
    char_spacing: f64,
    underline: bool,
    underline_style: Option<String>,
}

/// Break runs into wrapped lines, preserving per-run styling.
/// Uses embedded font widths (font_widths) for accurate character width measurement
/// that matches the PDF renderer's actual glyph widths.
fn wrap_runs_into_lines(
    runs: &[oxidocs_core::ir::Run],
    default_font_size: f64,
    max_width: f64,
    first_line_indent: f64,
    font_widths: &HashMap<u32, u16>,
    default_bold: bool,
    default_font_family: Option<&str>,
) -> Vec<Vec<LineRun>> {
    let mut lines: Vec<Vec<LineRun>> = Vec::new();
    let mut current_line: Vec<LineRun> = Vec::new();
    let mut current_width = first_line_indent;
    let effective_max = max_width;

    for run in runs {
        let font_size = run.style.font_size.unwrap_or(default_font_size as f32) as f64;
        let bold = run.style.bold || default_bold;
        let font_family = run.style.font_family.as_deref()
            .or(run.style.font_family_east_asia.as_deref())
            .or(default_font_family);
        let font_name = resolve_font(font_family, bold);
        let color = run.style.color.as_deref()
            .and_then(parse_hex_color)
            .unwrap_or(Color::Gray(0.0));
        // Word doubles the character spacing (applies on both sides of each character)
        let char_spacing = run.style.character_spacing.unwrap_or(0.0) as f64 * 2.0;
        let underline = run.style.underline;
        let underline_style = run.style.underline_style.clone();

        // Process text character by character for line wrapping.
        // Split into ASCII (Type1 base font) and non-ASCII (CJK CIDFont) segments.
        let mut buf = String::new();
        let mut buf_width = 0.0;
        let mut buf_is_cjk = false;

        for ch in run.text.chars() {
            if ch == '\t' {
                // Tab: flush buffer, then advance to next tab stop
                if !buf.is_empty() {
                    let fn_for_buf = if buf_is_cjk { "OxiCJK-Regular".to_string() } else { font_name.clone() };
                    current_line.push(LineRun {
                        text: buf.clone(), font_name: fn_for_buf,
                        font_size, color, char_spacing, underline, underline_style: underline_style.clone(),
                    });
                    current_width += buf_width;
                    buf.clear();
                    buf_width = 0.0;
                }
                // Advance to next default tab stop (Word default: every 36pt = 0.5 inch)
                let tab_interval = 36.0;
                let next_tab = ((current_width / tab_interval).floor() + 1.0) * tab_interval;
                let tab_gap = (next_tab - current_width).max(font_size * 0.25);
                // Approximate tab with spaces
                let space_w = font_size * 0.25;
                let num_spaces = (tab_gap / space_w).round().max(1.0) as usize;
                let tab_text = " ".repeat(num_spaces);
                current_line.push(LineRun {
                    text: tab_text, font_name: font_name.clone(),
                    font_size, color, char_spacing: 0.0, underline: false, underline_style: None,
                });
                current_width = next_tab;
                continue;
            }
            if ch == '\n' || ch == '\r' {
                // Explicit line break
                if !buf.is_empty() {
                    let fn_for_buf = if buf_is_cjk { "OxiCJK-Regular".to_string() } else { font_name.clone() };
                    current_line.push(LineRun {
                        text: buf.clone(), font_name: fn_for_buf,
                        font_size, color, char_spacing, underline, underline_style: underline_style.clone(),
                    });
                    buf.clear();
                    buf_width = 0.0;
                }
                lines.push(std::mem::take(&mut current_line));
                current_width = 0.0;
                continue;
            }

            let ch_is_cjk = ch as u32 > 0x7F;

            // If script changes (ASCII↔CJK), flush buffer as a separate run
            if !buf.is_empty() && ch_is_cjk != buf_is_cjk {
                let fn_for_buf = if buf_is_cjk { "OxiCJK-Regular".to_string() } else { font_name.clone() };
                current_line.push(LineRun {
                    text: buf.clone(), font_name: fn_for_buf,
                    font_size, color, char_spacing, underline, underline_style: underline_style.clone(),
                });
                current_width += buf_width;
                buf.clear();
                buf_width = 0.0;
            }
            buf_is_cjk = ch_is_cjk;

            // Use embedded font widths when available, matching PDF renderer
            let ch_width = (if let Some(&glyph_w) = font_widths.get(&(ch as u32)) {
                font_size * (glyph_w as f64 / 1000.0)
            } else {
                char_width(ch, font_size)
            }) + char_spacing;

            if current_width + buf_width + ch_width > effective_max && !(current_line.is_empty() && buf.is_empty()) {
                // Wrap: flush buffer to current line, start new line
                if !buf.is_empty() {
                    let fn_for_buf = if buf_is_cjk { "OxiCJK-Regular".to_string() } else { font_name.clone() };
                    current_line.push(LineRun {
                        text: buf.clone(), font_name: fn_for_buf,
                        font_size, color, char_spacing, underline, underline_style: underline_style.clone(),
                    });
                    buf.clear();
                    buf_width = 0.0;
                }
                lines.push(std::mem::take(&mut current_line));
                current_width = 0.0;
            }

            buf.push(ch);
            buf_width += ch_width;
        }

        // Flush remaining buffer
        if !buf.is_empty() {
            let fn_for_buf = if buf_is_cjk { "OxiCJK-Regular".to_string() } else { font_name.clone() };
            current_line.push(LineRun {
                text: buf, font_name: fn_for_buf,
                font_size, color, char_spacing, underline, underline_style: underline_style.clone(),
            });
            current_width += buf_width;
        }
    }

    // Flush last line
    if !current_line.is_empty() {
        lines.push(current_line);
    }

    if lines.is_empty() {
        lines.push(Vec::new());
    }

    lines
}

fn render_table(builder: &mut PdfBuilder, table: &Table, x_offset: f64, available_width: f64, registry: &FontMetricsRegistry, doc_default_font: Option<&str>) {
    // Build per-grid-column widths from tblGrid (canonical), cell widths, or autofit.
    // Use actual widths from the docx — do NOT scale to available_width.
    let grid_col_widths: Vec<f64> = if !table.grid_columns.is_empty() {
        table.grid_columns.iter().map(|&w| w as f64).collect()
    } else if let Some(first_row) = table.rows.first() {
        let specified: Vec<Option<f64>> = first_row.cells.iter()
            .map(|c| c.width.map(|w| w as f64))
            .collect();
        let total_specified: f64 = specified.iter().filter_map(|w| *w).sum();
        let num_cols = first_row.cells.len();
        if total_specified > 0.0 {
            specified.iter().map(|w| w.unwrap_or(total_specified / num_cols as f64)).collect()
        } else {
            // AutoFit to Contents: column width = max text width + cell padding
            let pad_lr = 5.4 * 2.0; // default L+R cell padding (10.8pt)
            let min_col_w = 11.1; // empty cell minimum (padding only)
            let mut col_widths = vec![0.0_f64; num_cols];
            for row in &table.rows {
                for (ci, cell) in row.cells.iter().enumerate() {
                    if ci >= num_cols { break; }
                    let mut max_w = 0.0_f64;
                    for block in &cell.blocks {
                        if let Block::Paragraph(para) = block {
                            // Measure text width using font metrics registry
                            for run in &para.runs {
                                let fs = run.style.font_size.unwrap_or(10.5);
                                let font_family = run.style.font_family.as_deref()
                                    .or(doc_default_font)
                                    .unwrap_or("Calibri");
                                let metrics = registry.get(font_family);
                                let w: f64 = run.text.chars()
                                    .map(|c| metrics.char_width_pt(c, fs) as f64)
                                    .sum();
                                max_w += w;
                            }
                        }
                    }
                    col_widths[ci] = col_widths[ci].max(max_w + pad_lr);
                }
            }
            // Ensure minimum column width for empty cells
            for w in &mut col_widths {
                if *w < min_col_w { *w = min_col_w; }
            }
            col_widths
        }
    } else {
        return;
    };

    // Default cell margins from table style (Word default: 5.4pt left/right, 0 top/bottom)
    let cell_margin_l = table.style.default_cell_margins.as_ref()
        .and_then(|m| m.left).unwrap_or(5.4) as f64;
    let cell_margin_r = table.style.default_cell_margins.as_ref()
        .and_then(|m| m.right).unwrap_or(5.4) as f64;
    let cell_margin_t = table.style.default_cell_margins.as_ref()
        .and_then(|m| m.top).unwrap_or(0.0) as f64;
    let cell_margin_b = table.style.default_cell_margins.as_ref()
        .and_then(|m| m.bottom).unwrap_or(0.0) as f64;

    // Table alignment: shift x_offset based on alignment + indent
    let table_total_width: f64 = grid_col_widths.iter().sum();
    let table_x_offset = match table.style.alignment.as_deref() {
        Some("center") => x_offset + (available_width - table_total_width).max(0.0) / 2.0,
        Some("right") => x_offset + (available_width - table_total_width).max(0.0),
        _ => x_offset + table.style.indent.unwrap_or(0.0) as f64,
    };

    // Border style from document
    let border_color = table.style.border_color.as_deref()
        .and_then(parse_hex_color)
        .unwrap_or(Color::Gray(0.0)); // Word default: black borders
    let border_width = table.style.border_width.map(|w| w as f64).unwrap_or(0.5);

    let row_min_height = 14.0;

    for (_row_idx, row) in table.rows.iter().enumerate() {
        // Compute cell widths from grid columns + gridSpan
        // Handle rows where total gridSpan < grid_col count (vMerge or structural columns)
        let mut cell_widths = Vec::new();
        let total_span: usize = row.cells.iter().map(|c| c.grid_span.max(1) as usize).sum();
        let num_grid = grid_col_widths.len();
        // If row spans fewer columns than grid, skip leading columns
        let start_grid = if total_span < num_grid { num_grid - total_span } else { 0 };
        let mut grid_idx = start_grid;
        for cell in &row.cells {
            let span = cell.grid_span.max(1) as usize;
            let end = (grid_idx + span).min(num_grid);
            let cw: f64 = if grid_idx < num_grid {
                grid_col_widths[grid_idx..end].iter().sum()
            } else {
                // Fallback: use cell's own width if available
                cell.width.unwrap_or(0.0) as f64
            };
            // If computed grid width is tiny but cell has a specified width, use cell width
            let cw = if cw < 5.0 {
                cell.width.map(|w| w as f64).unwrap_or(cw).max(cw)
            } else {
                cw
            };
            cell_widths.push(cw);
            grid_idx = end;
            if grid_idx >= num_grid { break; }
        }

        // Calculate row height based on content (skip vMerge continue cells)
        let is_exact_height = row.height_rule.as_deref() == Some("exact");
        let mut row_height = row.height.map(|h| h as f64).unwrap_or(row_min_height);
        if !is_exact_height {
            // atLeast (default): max of specified height and content
            for (ci, cell) in row.cells.iter().enumerate() {
                if ci >= cell_widths.len() { break; }
                if cell.v_merge.as_deref() == Some("continue") { continue; }
                let content_w = (cell_widths[ci] - cell_margin_l - cell_margin_r).max(1.0);
                let cell_h = estimate_cell_height(cell, content_w, registry, doc_default_font);
                let needed = cell_h + cell_margin_t + cell_margin_b;
                if needed > row_height {
                    row_height = needed;
                }
            }
        }
        // exact: use specified height regardless of content

        if builder.needs_page_break(row_height + 2.0) {
            builder.new_page();
        }

        let row_top = builder.cursor_y;
        let row_bottom = row_top + row_height;

        // Draw cells — offset by any skipped leading grid columns
        let leading_width: f64 = grid_col_widths[..start_grid].iter().sum();
        let mut cell_x = table_x_offset + leading_width;
        for (ci, cell) in row.cells.iter().enumerate() {
            if ci >= cell_widths.len() { break; }
            let cw = cell_widths[ci];
            let is_vmerge_continue = cell.v_merge.as_deref() == Some("continue");

            // Cell shading from document
            if let Some(ref shade) = cell.shading {
                if let Some(bg_color) = parse_hex_color(shade) {
                    builder.add_rect_fill(cell_x, row_top, cw, row_height, bg_color);
                }
            }

            // Cell borders: use cell-specific borders if available, else table defaults
            if table.style.border || cell.borders.is_some() {
                let draw_cell_border = |builder: &mut PdfBuilder, bd: Option<&oxidocs_core::ir::BorderDef>, x1: f64, y1: f64, x2: f64, y2: f64, default_draw: bool| {
                    match bd {
                        Some(bd) if bd.style == "nil" || bd.style == "none" => {},
                        Some(bd) => {
                            let bw = if bd.width > 0.0 { bd.width as f64 } else { border_width };
                            let bc = bd.color.as_deref().and_then(parse_hex_color).unwrap_or(border_color);
                            builder.add_line(x1, y1, x2, y2, bw, bc);
                        },
                        None if default_draw => {
                            builder.add_line(x1, y1, x2, y2, border_width, border_color);
                        },
                        _ => {},
                    }
                };
                let cb = cell.borders.as_ref();
                let has_table_border = table.style.border;
                draw_cell_border(builder, cb.and_then(|b| b.top.as_ref()), cell_x, row_top, cell_x + cw, row_top, has_table_border);
                draw_cell_border(builder, cb.and_then(|b| b.bottom.as_ref()), cell_x, row_bottom, cell_x + cw, row_bottom, has_table_border);
                draw_cell_border(builder, cb.and_then(|b| b.left.as_ref()), cell_x, row_top, cell_x, row_bottom, has_table_border);
                draw_cell_border(builder, cb.and_then(|b| b.right.as_ref()), cell_x + cw, row_top, cell_x + cw, row_bottom, has_table_border);
            }

            // Skip content rendering for vertically merged continuation cells
            if is_vmerge_continue {
                cell_x += cw;
                continue;
            }

            // Cell text — use cell-specific margins if available, else table defaults
            let cm_l = cell.margins.as_ref().and_then(|m| m.left).unwrap_or(cell_margin_l as f32) as f64;
            let cm_r = cell.margins.as_ref().and_then(|m| m.right).unwrap_or(cell_margin_r as f32) as f64;
            let cm_t = cell.margins.as_ref().and_then(|m| m.top).unwrap_or(cell_margin_t as f32) as f64;
            let cell_content_x = cell_x + cm_l;
            let cell_content_width = (cw - cm_l - cm_r).max(1.0);
            let mut cell_text_y = row_top + cm_t;

            for block in &cell.blocks {
                if let Block::Paragraph(para) = block {
                    let (font_size, cell_font, bold) = para_font_props(para);
                    let line_height = compute_line_height(
                        registry, cell_font, font_size as f32,
                        para.style.line_spacing, para.style.line_spacing_rule.as_deref(),
                        doc_default_font, 0.0, // no grid snap in table cells
                    );

                    cell_text_y += para.style.space_before.unwrap_or(0.0) as f64;

                    // Clip content that overflows exact row height
                    if is_exact_height && cell_text_y >= row_bottom {
                        break;
                    }

                    let text: String = para.runs.iter().map(|r| r.text.as_str()).collect();
                    if text.trim().is_empty() {
                        cell_text_y += line_height;
                    } else {
                        // Use per-run wrapping with ASCII/CJK font splitting
                        let cell_default_bold = para.style.default_run_style.as_ref().map(|rs| rs.bold).unwrap_or(false);
                        let cell_default_ff = para.style.default_run_style.as_ref().and_then(|rs| rs.font_family.as_deref());
                        let cell_lines = wrap_runs_into_lines(&para.runs, font_size, cell_content_width, 0.0, &builder.font_widths, cell_default_bold, cell_default_ff);
                        let text_color = para_color(para);

                        for line_runs in &cell_lines {
                            let line_width: f64 = line_runs.iter()
                                .map(|lr| {
                                    let base_w = builder.text_width(&lr.text, lr.font_size);
                                    base_w + lr.char_spacing * lr.text.chars().count() as f64
                                })
                                .sum();
                            let align_off = match para.alignment {
                                Alignment::Center => (cell_content_width - line_width) / 2.0,
                                Alignment::Right => cell_content_width - line_width,
                                _ => 0.0,
                            };

                            // Advance cursor FIRST (baseline at cell_text_y + line_height)
                            cell_text_y += line_height;
                            // Clip text that overflows exact row height
                            if is_exact_height && cell_text_y > row_bottom {
                                break;
                            }
                            let mut run_x = cell_content_x + align_off.max(0.0);
                            for lr in line_runs {
                                if !lr.text.is_empty() {
                                    if lr.char_spacing.abs() > 0.01 {
                                        for ch in lr.text.chars() {
                                            let ch_str = ch.to_string();
                                            builder.add_text(
                                                run_x, cell_text_y,
                                                ch_str.clone(), lr.font_name.clone(), lr.font_size, text_color,
                                            );
                                            run_x += builder.text_width(&ch_str, lr.font_size) + lr.char_spacing;
                                        }
                                    } else {
                                        builder.add_text(
                                            run_x, cell_text_y,
                                            lr.text.clone(), lr.font_name.clone(), lr.font_size, text_color,
                                        );
                                        run_x += builder.text_width(&lr.text, lr.font_size);
                                    }
                                }
                            }
                        }
                    }
                    cell_text_y += para.style.space_after.unwrap_or(0.0) as f64;
                }
            }

            cell_x += cw;
        }

        builder.cursor_y = row_bottom;
    }
}

fn estimate_cell_height(cell: &oxidocs_core::ir::TableCell, content_width: f64, registry: &FontMetricsRegistry, doc_default_font: Option<&str>) -> f64 {
    let mut h = 0.0;
    for block in &cell.blocks {
        if let Block::Paragraph(para) = block {
            let (font_size, cell_font, _bold) = para_font_props(para);
            let line_height = compute_line_height(
                registry, cell_font, font_size as f32,
                para.style.line_spacing, para.style.line_spacing_rule.as_deref(),
                doc_default_font, 0.0, // no grid snap in table cells
            );
            h += para.style.space_before.unwrap_or(0.0) as f64;
            let text: String = para.runs.iter().map(|r| r.text.as_str()).collect();
            if text.trim().is_empty() {
                h += line_height;
            } else {
                let lines = wrap_text(&text, font_size, content_width);
                h += lines.len() as f64 * line_height;
            }
            h += para.style.space_after.unwrap_or(0.0) as f64;
        }
    }
    h.max(12.0)
}

// --- Font Embedding ---

/// System font paths to search for CJK fonts (Regular)
const CJK_FONT_PATHS: &[&str] = &[
    // Prefer MS Gothic (most common in Japanese government documents)
    "C:\\Windows\\Fonts\\msgothic.ttc",
    "C:\\Windows\\Fonts\\YuGothR.ttc",
    "C:\\Windows\\Fonts\\meiryo.ttc",
    "/usr/share/fonts/opentype/noto/NotoSansCJK-Regular.ttc",
    "/System/Library/Fonts/ヒラギノ角ゴシック W3.ttc",
    "/System/Library/Fonts/HiraginoSans-W3.ttc",
];

/// System font paths to search for CJK fonts (Bold)
const CJK_FONT_PATHS_BOLD: &[&str] = &[
    "C:\\Windows\\Fonts\\msgothic.ttc",
    "C:\\Windows\\Fonts\\YuGothB.ttc",
    "C:\\Windows\\Fonts\\meiryo.ttc",
    "/usr/share/fonts/opentype/noto/NotoSansCJK-Bold.ttc",
    "/System/Library/Fonts/ヒラギノ角ゴシック W6.ttc",
    "/System/Library/Fonts/HiraginoSans-W6.ttc",
];

/// System font paths for Latin fonts (Calibri)
const LATIN_FONT_PATHS: &[(&str, &str)] = &[
    ("C:\\Windows\\Fonts\\calibri.ttf", "C:\\Windows\\Fonts\\calibrib.ttf"),
    ("/usr/share/fonts/truetype/msttcorefonts/calibri.ttf", "/usr/share/fonts/truetype/msttcorefonts/calibrib.ttf"),
];

/// Path to Python3 for font subsetting (requires fonttools)
const PYTHON3: &str = "python3";

/// Build embedded fonts map from system fonts.
/// Detects CJK font usage in the document and embeds the font data.
fn build_embedded_fonts(doc: &oxidocs_core::Document) -> HashMap<String, EmbeddedFont> {
    let mut fonts = HashMap::new();

    // Collect all font names used in the document that likely need CJK
    let mut needs_cjk_font = false;
    let mut font_names_used: Vec<String> = Vec::new();

    for page in &doc.pages {
        for block in &page.blocks {
            collect_cjk_fonts_from_block(block, &mut needs_cjk_font, &mut font_names_used);
        }
    }

    // Also collect document-level default fonts (from rPrDefault / styles)
    if let Some(ref drs) = doc.styles.doc_default_run_style {
        if let Some(ref ff) = drs.font_family_east_asia {
            if !font_names_used.contains(ff) {
                font_names_used.push(ff.clone());
            }
        }
        if let Some(ref ff) = drs.font_family {
            if !font_names_used.contains(ff) {
                font_names_used.push(ff.clone());
            }
        }
    }

    // Embed Latin fonts (Calibri + Cambria) regardless of CJK usage
    let mut used_chars_latin: std::collections::BTreeSet<char> = std::collections::BTreeSet::new();
    for page in &doc.pages {
        for block in &page.blocks {
            collect_chars_from_block(block, &mut used_chars_latin);
        }
    }
    embed_latin_fonts(&used_chars_latin, &mut fonts);
    embed_cambria_fonts(&used_chars_latin, &mut fonts);

    if !needs_cjk_font {
        return fonts;
    }

    // Collect all unique characters used in the document (for font subsetting)
    let mut used_chars: std::collections::BTreeSet<char> = std::collections::BTreeSet::new();
    for page in &doc.pages {
        for block in &page.blocks {
            collect_chars_from_block(block, &mut used_chars);
        }
        // Also collect characters from TextBoxes
        for text_box in &page.text_boxes {
            for block in &text_box.blocks {
                collect_chars_from_block(block, &mut used_chars);
            }
        }
    }

    // Per-family CJK font embedding: map font families to system font files
    // Try system fonts first, fall back to bundled OxiGothic/OxiMincho (Noto-based, OFL)
    let exe_dir = std::env::current_exe().ok()
        .and_then(|p| p.parent().map(|d| d.to_path_buf()))
        .unwrap_or_default();
    let oxi_gothic = exe_dir.join("fonts").join("OxiGothic.ttf");
    let oxi_mincho = exe_dir.join("fonts").join("OxiMincho.ttf");
    let oxi_gothic_str = oxi_gothic.to_string_lossy().to_string();
    let oxi_mincho_str = oxi_mincho.to_string_lossy().to_string();

    let family_font_map: Vec<(&str, Vec<&str>)> = vec![
        ("MS-Gothic",  vec!["C:\\Windows\\Fonts\\msgothic.ttc", oxi_gothic_str.as_str()]),
        ("MS-PGothic", vec!["C:\\Windows\\Fonts\\msgothic.ttc", oxi_gothic_str.as_str()]),
        ("MS-Mincho",  vec!["C:\\Windows\\Fonts\\msmincho.ttc", oxi_mincho_str.as_str()]),
        ("MS-PMincho", vec!["C:\\Windows\\Fonts\\msmincho.ttc", oxi_mincho_str.as_str()]),
    ];

    // Embed per-family fonts: resolve original font names to PDF names, match against map
    let mut embedded_families: std::collections::HashSet<String> = std::collections::HashSet::new();
    for name in &font_names_used {
        let resolved = resolve_font(Some(name), false);
        for (family, paths) in &family_font_map {
            if resolved == *family {
                if !embedded_families.contains(&resolved) {
                    let label = resolved.replace('-', "_");
                    let found_path = paths.iter().find(|p| std::path::Path::new(p).exists());
                    if let Some(result) = found_path.and_then(|p| subset_font_file(p, &used_chars, &label)) {
                        eprintln!("Embedding CJK font '{}' ({}) from {} ({} bytes)", resolved, name, found_path.unwrap(), result.data.len());
                        fonts.insert(resolved.clone(), result);
                        embedded_families.insert(resolved.clone());
                    }
                }
                break;
            }
        }
    }

    // Try to find, subset, and load a default CJK font from the system (for OxiCJK-Regular)
    let font_result = match find_and_subset_cjk_font(&used_chars) {
        Some(result) => result,
        None => {
            eprintln!("Warning: No CJK system font found. PDF may not render Japanese text correctly.");
            return fonts;
        }
    };

    eprintln!("Embedding CJK font ({} bytes, {} glyph mappings)",
        font_result.data.len(), font_result.unicode_to_gid.len());

    // Map remaining CJK font names (not already per-family embedded) to default CJK font
    for name in &font_names_used {
        let resolved = resolve_font(Some(name), false);
        if !fonts.contains_key(&resolved) {
            fonts.insert(resolved.clone(), font_result.clone());
        }
        let resolved_bold = resolve_font(Some(name), true);
        if !fonts.contains_key(&resolved_bold) {
            fonts.insert(resolved_bold, font_result.clone());
        }
    }

    // Map the dedicated CJK font names used for non-ASCII text spans
    fonts.insert("OxiCJK-Regular".to_string(), font_result.clone());

    // Try to embed a Bold CJK font separately
    if let Some(bold_result) = find_and_subset_cjk_font_bold(&used_chars) {
        eprintln!("Embedding CJK Bold font ({} bytes, {} glyph mappings)",
            bold_result.data.len(), bold_result.unicode_to_gid.len());
        fonts.insert("OxiCJK-Bold".to_string(), bold_result);
    } else {
        // Fallback: use regular font for bold too
        fonts.insert("OxiCJK-Bold".to_string(), font_result.clone());
    }

    // Latin fonts already embedded above (before CJK check)

    fonts
}

/// Embed Cambria Regular and Bold for documents using Cambria/Century.
fn embed_cambria_fonts(used_chars: &std::collections::BTreeSet<char>, fonts: &mut HashMap<String, EmbeddedFont>) {
    let cambria_paths: &[(&str, &str)] = &[
        ("C:\\Windows\\Fonts\\cambria.ttc", "C:\\Windows\\Fonts\\cambriab.ttf"),
    ];
    let pair = cambria_paths.iter().find(|(r, _)| std::path::Path::new(r).exists());
    let (regular_path, bold_path) = match pair {
        Some(p) => p,
        None => return,
    };
    let latin_chars: std::collections::BTreeSet<char> = used_chars.iter()
        .filter(|c| (**c as u32) < 0x2000)
        .copied()
        .collect();
    if let Some(result) = subset_font_file(regular_path, &latin_chars, "cambria") {
        eprintln!("Embedding Cambria font ({} bytes, {} glyphs)", result.data.len(), result.unicode_to_gid.len());
        fonts.insert("OxiCambria-Regular".to_string(), result);
    }
    if let Some(result) = subset_font_file(bold_path, &latin_chars, "cambria-bold") {
        eprintln!("Embedding Cambria Bold font ({} bytes, {} glyphs)", result.data.len(), result.unicode_to_gid.len());
        fonts.insert("OxiCambria-Bold".to_string(), result);
    }
}

/// Embed Calibri Regular and Bold for Latin text rendering.
fn embed_latin_fonts(used_chars: &std::collections::BTreeSet<char>, fonts: &mut HashMap<String, EmbeddedFont>) {
    // Find available Latin font pair
    let latin_pair = LATIN_FONT_PATHS.iter().find(|(r, _)| std::path::Path::new(r).exists());
    let (regular_path, bold_path) = match latin_pair {
        Some(pair) => pair,
        None => { eprintln!("Warning: No Latin system font found"); return; }
    };

    // Filter to ASCII/Latin chars only (for smaller subset)
    let latin_chars: std::collections::BTreeSet<char> = used_chars.iter()
        .filter(|c| (**c as u32) < 0x2000)
        .copied()
        .collect();

    if let Some(result) = subset_font_file(regular_path, &latin_chars, "latin") {
        eprintln!("Embedding Latin font ({} bytes, {} glyphs)", result.data.len(), result.unicode_to_gid.len());
        fonts.insert("OxiLatin-Regular".to_string(), result);
    }
    if let Some(result) = subset_font_file(bold_path, &latin_chars, "latin-bold") {
        eprintln!("Embedding Latin Bold font ({} bytes, {} glyphs)", result.data.len(), result.unicode_to_gid.len());
        fonts.insert("OxiLatin-Bold".to_string(), result);
    }
}

/// Subset a single font file and return EmbeddedFont.
fn subset_font_file(font_path: &str, chars: &std::collections::BTreeSet<char>, label: &str) -> Option<EmbeddedFont> {
    let char_string: String = chars.iter()
        .filter(|c| { let cp = **c as u32; !(0xD800..=0xDFFF).contains(&cp) })
        .collect();

    let subset_path = format!("/tmp/oxi-font-subset-{}.otf", label);
    let cidmap_path = format!("/tmp/oxi-font-cidmap-{}.json", label);
    let widths_path = format!("/tmp/oxi-font-widths-{}.json", label);
    let font_path_escaped = font_path.replace('\\', "\\\\");
    let python_script = format!(
        r#"
import sys, json
sys.stdin.reconfigure(encoding='utf-8', errors='replace')
from fontTools.ttLib import TTCollection, TTFont
from fontTools.subset import Subsetter, Options

font_path = "{font_path_escaped}"
data = open(font_path, 'rb').read()

if data[:4] == b'ttcf':
    ttc = TTCollection(font_path)
    font = ttc[0]
else:
    font = TTFont(font_path)

chars = sys.stdin.read()
chars = ''.join(c for c in chars if ord(c) < 0xD800 or ord(c) > 0xDFFF)
opts = Options()
subsetter = Subsetter(options=opts)
subsetter.populate(text=chars)
subsetter.subset(font)

font.save("{subset_path}")

cmap = font.getBestCmap()
unicode_to_cid = {{}}
cid_widths = {{}}

upem = font['head'].unitsPerEm
hmtx = font['hmtx']

for unicode_val, glyph_name in cmap.items():
    cid = None
    if glyph_name.startswith('cid'):
        cid = int(glyph_name[3:])
    elif glyph_name == '.notdef':
        continue
    else:
        glyph_order = font.getGlyphOrder()
        try:
            cid = glyph_order.index(glyph_name)
        except ValueError:
            continue

    if cid is not None:
        unicode_to_cid[unicode_val] = cid
        if glyph_name in hmtx.metrics:
            advance, _ = hmtx.metrics[glyph_name]
            w = int(advance * 1000 / upem)
            cid_widths[cid] = w

with open("{cidmap_path}", 'w') as f:
    json.dump(unicode_to_cid, f)

with open("{widths_path}", 'w') as f:
    json.dump(cid_widths, f)

ps_name = None
for record in font['name'].names:
    if record.nameID == 6:
        try:
            ps_name = record.toUnicode()
            break
        except:
            pass
if ps_name:
    with open("{widths_path}.psname", 'w') as f:
        f.write(ps_name)

print(f"OK {{len(unicode_to_cid)}} mappings, {{len(cid_widths)}} widths")
"#
    );

    let result = std::process::Command::new(PYTHON3)
        .arg("-c")
        .arg(&python_script)
        .stdin(std::process::Stdio::piped())
        .stdout(std::process::Stdio::piped())
        .stderr(std::process::Stdio::piped())
        .spawn();

    match result {
        Ok(mut child) => {
            if let Some(ref mut stdin) = child.stdin {
                use std::io::Write;
                let _ = stdin.write_all(char_string.as_bytes());
            }
            let output = child.wait_with_output().ok()?;
            if !output.status.success() {
                eprintln!("Font subsetting ({}) failed: {}", label, String::from_utf8_lossy(&output.stderr));
                return None;
            }
            eprintln!("Font subset ({}): {}", label, String::from_utf8_lossy(&output.stdout).trim());
        }
        Err(e) => {
            eprintln!("Python not available for {} subsetting: {}", label, e);
            return None;
        }
    }

    let otf_data = fs::read(&subset_path).ok()?;
    let unicode_to_gid = if let Ok(json_str) = fs::read_to_string(&cidmap_path) {
        parse_cidmap_json(&json_str)
    } else {
        parse_cmap_table(&otf_data)
    };
    let cid_widths = if let Ok(json_str) = fs::read_to_string(&widths_path) {
        parse_cidwidths_json(&json_str)
    } else {
        HashMap::new()
    };

    // Read PostScript name from font subsetting output
    let ps_name = fs::read_to_string(format!("{}.psname", &widths_path)).ok()
        .map(|s| s.trim().to_string())
        .filter(|s| !s.is_empty());
    if let Some(ref psn) = ps_name {
        eprintln!("PostScript name: {}", psn);
    }

    let is_cff = otf_data.starts_with(b"OTTO") || has_cff_table(&otf_data);
    if is_cff {
        if let Some(cff) = extract_cff_from_otf(&otf_data) {
            return Some(EmbeddedFont { ps_name: ps_name.clone(),
                data: cff,
                format: FontFormat::OpenTypeCff,
                unicode_to_gid,
                cid_widths,
            });
        }
    }

    Some(EmbeddedFont { ps_name,
        data: otf_data,
        format: FontFormat::TrueType,
        unicode_to_gid,
        cid_widths,
    })
}

fn collect_chars_from_block(block: &Block, chars: &mut std::collections::BTreeSet<char>) {
    match block {
        Block::Paragraph(para) => {
            for run in &para.runs {
                for ch in run.text.chars() {
                    chars.insert(ch);
                }
            }
            if let Some(ref marker) = para.style.list_marker {
                for ch in marker.chars() {
                    chars.insert(ch);
                }
            }
        }
        Block::Table(table) => {
            for row in &table.rows {
                for cell in &row.cells {
                    for block in &cell.blocks {
                        collect_chars_from_block(block, chars);
                    }
                }
            }
        }
        _ => {}
    }
}

fn collect_cjk_fonts_from_block(block: &Block, needs_cjk: &mut bool, font_names: &mut Vec<String>) {
    match block {
        Block::Paragraph(para) => {
            for run in &para.runs {
                if run.text.chars().any(|c| c as u32 > 0x7F) {
                    *needs_cjk = true;
                }
                if let Some(ref family) = run.style.font_family {
                    if !font_names.contains(family) {
                        font_names.push(family.clone());
                    }
                }
                if let Some(ref family) = run.style.font_family_east_asia {
                    if !font_names.contains(family) {
                        font_names.push(family.clone());
                    }
                }
            }
            if let Some(ref drs) = para.style.default_run_style {
                if let Some(ref family) = drs.font_family {
                    if !font_names.contains(family) {
                        font_names.push(family.clone());
                    }
                }
            }
        }
        Block::Table(table) => {
            for row in &table.rows {
                for cell in &row.cells {
                    for block in &cell.blocks {
                        collect_cjk_fonts_from_block(block, needs_cjk, font_names);
                    }
                }
            }
        }
        _ => {}
    }
}

fn find_and_subset_cjk_font(used_chars: &std::collections::BTreeSet<char>) -> Option<EmbeddedFont> {
    let font_path = CJK_FONT_PATHS.iter().find(|p| std::path::Path::new(p).exists())?;
    eprintln!("Found CJK font: {}", font_path);
    subset_font_file(font_path, used_chars, "cjk")
}

fn find_and_subset_cjk_font_bold(used_chars: &std::collections::BTreeSet<char>) -> Option<EmbeddedFont> {
    let font_path = CJK_FONT_PATHS_BOLD.iter().find(|p| std::path::Path::new(p).exists())?;
    eprintln!("Found CJK Bold font: {}", font_path);

    let char_string: String = used_chars.iter()
        .filter(|c| { let cp = **c as u32; !(0xD800..=0xDFFF).contains(&cp) })
        .collect();

    let subset_path = "/tmp/oxi-font-subset-bold.otf";
    let cidmap_path = "/tmp/oxi-font-cidmap-bold.json";
    let widths_path = "/tmp/oxi-font-widths-bold.json";
    let font_path_escaped = font_path.replace('\\', "\\\\");
    let python_script = format!(
        r#"
import sys, json
sys.stdin.reconfigure(encoding='utf-8', errors='replace')
from fontTools.ttLib import TTCollection, TTFont
from fontTools.subset import Subsetter, Options

font_path = "{font_path_escaped}"
data = open(font_path, 'rb').read()

if data[:4] == b'ttcf':
    ttc = TTCollection(font_path)
    font = ttc[0]
else:
    font = TTFont(font_path)

chars = sys.stdin.read()
chars = ''.join(c for c in chars if ord(c) < 0xD800 or ord(c) > 0xDFFF)
opts = Options()
subsetter = Subsetter(options=opts)
subsetter.populate(text=chars)
subsetter.subset(font)

font.save("{subset_path}")

cmap = font.getBestCmap()
unicode_to_cid = {{}}
cid_widths = {{}}

upem = font['head'].unitsPerEm
hmtx = font['hmtx']

for unicode_val, glyph_name in cmap.items():
    cid = None
    if glyph_name.startswith('cid'):
        cid = int(glyph_name[3:])
    elif glyph_name == '.notdef':
        continue
    else:
        glyph_order = font.getGlyphOrder()
        try:
            cid = glyph_order.index(glyph_name)
        except ValueError:
            continue

    if cid is not None:
        unicode_to_cid[unicode_val] = cid
        if glyph_name in hmtx.metrics:
            advance, _ = hmtx.metrics[glyph_name]
            w = int(advance * 1000 / upem)
            cid_widths[cid] = w

with open("{cidmap_path}", 'w') as f:
    json.dump(unicode_to_cid, f)

with open("{widths_path}", 'w') as f:
    json.dump(cid_widths, f)

ps_name = None
for record in font['name'].names:
    if record.nameID == 6:
        try:
            ps_name = record.toUnicode()
            break
        except:
            pass
if ps_name:
    with open("{widths_path}.psname", 'w') as f:
        f.write(ps_name)

print(f"OK {{len(unicode_to_cid)}} mappings, {{len(cid_widths)}} widths")
"#
    );

    let result = std::process::Command::new(PYTHON3)
        .arg("-c")
        .arg(&python_script)
        .stdin(std::process::Stdio::piped())
        .stdout(std::process::Stdio::piped())
        .stderr(std::process::Stdio::piped())
        .spawn();

    match result {
        Ok(mut child) => {
            if let Some(ref mut stdin) = child.stdin {
                use std::io::Write;
                let _ = stdin.write_all(char_string.as_bytes());
            }
            let output = child.wait_with_output().ok()?;
            if !output.status.success() {
                eprintln!("Bold font subsetting failed: {}", String::from_utf8_lossy(&output.stderr));
                return None;
            }
            eprintln!("Font subset (Bold): {}", String::from_utf8_lossy(&output.stdout).trim());
        }
        Err(e) => {
            eprintln!("Python not available for bold font subsetting: {}", e);
            return None;
        }
    }

    let otf_data = fs::read(subset_path).ok()?;
    let unicode_to_gid = if let Ok(json_str) = fs::read_to_string(cidmap_path) {
        parse_cidmap_json(&json_str)
    } else {
        parse_cmap_table(&otf_data)
    };
    let cid_widths = if let Ok(json_str) = fs::read_to_string(widths_path) {
        parse_cidwidths_json(&json_str)
    } else {
        HashMap::new()
    };

    let is_cff = otf_data.starts_with(b"OTTO") || has_cff_table(&otf_data);
    if is_cff {
        if let Some(cff) = extract_cff_from_otf(&otf_data) {
            return Some(EmbeddedFont { ps_name: None,
                data: cff,
                format: FontFormat::OpenTypeCff,
                unicode_to_gid,
                cid_widths,
            });
        }
    }

    Some(EmbeddedFont { ps_name: None,
        data: otf_data,
        format: FontFormat::TrueType,
        unicode_to_gid,
        cid_widths,
    })
}

/// Fallback: load the full font without subsetting.
fn find_system_cjk_font_full() -> Option<EmbeddedFont> {
    find_system_cjk_font()
}

/// Find a CJK font on the system and extract it for PDF embedding (no subsetting).
fn find_system_cjk_font() -> Option<EmbeddedFont> {
    for path in CJK_FONT_PATHS {
        if let Ok(data) = fs::read(path) {
            eprintln!("Found CJK font: {}", path);

            // Get the standalone font data (extract from TTC if needed)
            let otf_data = if data.len() > 12 && &data[0..4] == b"ttcf" {
                extract_font_from_ttc(&data, 0)?
            } else {
                data
            };

            // Parse cmap table to get Unicode → GlyphID mapping
            let unicode_to_gid = parse_cmap_table(&otf_data);
            eprintln!("Parsed cmap: {} Unicode→GID entries", unicode_to_gid.len());

            // Determine format and extract appropriate data for embedding
            let is_cff = otf_data.starts_with(b"OTTO") || has_cff_table(&otf_data);
            if is_cff {
                if let Some(cff) = extract_cff_from_otf(&otf_data) {
                    return Some(EmbeddedFont { ps_name: None,
                        data: cff,
                        format: FontFormat::OpenTypeCff,
                        unicode_to_gid,
                        cid_widths: HashMap::new(),
                    });
                }
            }

            // TrueType
            return Some(EmbeddedFont { ps_name: None,
                data: otf_data,
                format: FontFormat::TrueType,
                unicode_to_gid,
                cid_widths: HashMap::new(),
            });
        }
    }
    None
}

/// Extract a single font from a TTC (TrueType Collection).
fn extract_font_from_ttc(data: &[u8], font_index: usize) -> Option<Vec<u8>> {
    if data.len() < 12 || &data[0..4] != b"ttcf" {
        return None;
    }
    let num_fonts = u32::from_be_bytes([data[8], data[9], data[10], data[11]]) as usize;
    if font_index >= num_fonts {
        return None;
    }

    let offset_pos = 12 + font_index * 4;
    if offset_pos + 4 > data.len() {
        return None;
    }
    let font_offset = u32::from_be_bytes([
        data[offset_pos], data[offset_pos + 1],
        data[offset_pos + 2], data[offset_pos + 3],
    ]) as usize;

    if font_offset + 12 > data.len() {
        return None;
    }

    // Read table directory at font_offset
    let num_tables = u16::from_be_bytes([data[font_offset + 4], data[font_offset + 5]]) as usize;
    let table_dir_start = font_offset + 12;

    // Collect all table records
    struct TableRecord {
        tag: [u8; 4],
        checksum: u32,
        offset: usize,
        length: usize,
    }

    let mut tables = Vec::new();
    for i in 0..num_tables {
        let rec_off = table_dir_start + i * 16;
        if rec_off + 16 > data.len() { return None; }
        let mut tag = [0u8; 4];
        tag.copy_from_slice(&data[rec_off..rec_off + 4]);
        let checksum = u32::from_be_bytes([
            data[rec_off + 4], data[rec_off + 5], data[rec_off + 6], data[rec_off + 7],
        ]);
        let offset = u32::from_be_bytes([
            data[rec_off + 8], data[rec_off + 9], data[rec_off + 10], data[rec_off + 11],
        ]) as usize;
        let length = u32::from_be_bytes([
            data[rec_off + 12], data[rec_off + 13], data[rec_off + 14], data[rec_off + 15],
        ]) as usize;
        tables.push(TableRecord { tag, checksum, offset, length });
    }

    // Build a standalone font file
    // OTF/TTF header: sfVersion(4) + numTables(2) + searchRange(2) + entrySelector(2) + rangeShift(2) = 12
    // Table directory: numTables * 16
    let header_size = 12 + num_tables * 16;
    let mut font = Vec::new();

    // Copy sfVersion from the original font
    font.extend_from_slice(&data[font_offset..font_offset + 4]);

    // numTables, searchRange, entrySelector, rangeShift
    font.extend_from_slice(&(num_tables as u16).to_be_bytes());

    // Calculate search params
    let mut search_range = 1u16;
    let mut entry_selector = 0u16;
    while search_range * 2 <= num_tables as u16 {
        search_range *= 2;
        entry_selector += 1;
    }
    search_range *= 16;
    let range_shift = (num_tables as u16) * 16 - search_range;
    font.extend_from_slice(&search_range.to_be_bytes());
    font.extend_from_slice(&entry_selector.to_be_bytes());
    font.extend_from_slice(&range_shift.to_be_bytes());

    // Calculate new offsets for each table
    let mut current_offset = header_size;
    let mut new_offsets = Vec::new();
    for t in &tables {
        new_offsets.push(current_offset);
        current_offset += (t.length + 3) & !3; // 4-byte align
    }

    // Write table directory
    for (i, t) in tables.iter().enumerate() {
        font.extend_from_slice(&t.tag);
        font.extend_from_slice(&t.checksum.to_be_bytes());
        font.extend_from_slice(&(new_offsets[i] as u32).to_be_bytes());
        font.extend_from_slice(&(t.length as u32).to_be_bytes());
    }

    // Write table data
    for t in &tables {
        if t.offset + t.length > data.len() { return None; }
        font.extend_from_slice(&data[t.offset..t.offset + t.length]);
        // Pad to 4-byte boundary
        while font.len() % 4 != 0 {
            font.push(0);
        }
    }

    Some(font)
}

/// Parse Unicode→CID mapping from JSON output by Python fonttools.
fn parse_cidmap_json(json_str: &str) -> HashMap<u32, u16> {
    let mut result = HashMap::new();
    // Simple JSON parsing: {"unicode": cid, ...}
    // Keys are stringified Unicode codepoints, values are CID numbers
    if let Ok(map) = serde_json::from_str::<HashMap<String, u64>>(json_str) {
        for (key, cid) in map {
            if let Ok(unicode) = key.parse::<u32>() {
                if cid <= 0xFFFF {
                    result.insert(unicode, cid as u16);
                }
            }
        }
    }
    result
}

/// Parse CID → width mapping from JSON output by Python fonttools.
fn parse_cidwidths_json(json_str: &str) -> HashMap<u16, u16> {
    let mut result = HashMap::new();
    if let Ok(map) = serde_json::from_str::<HashMap<String, u64>>(json_str) {
        for (key, width) in map {
            if let Ok(cid) = key.parse::<u16>() {
                if width <= 0xFFFF {
                    result.insert(cid, width as u16);
                }
            }
        }
    }
    result
}

/// Parse the cmap table from an OTF/TTF font to extract Unicode → GlyphID mappings.
fn parse_cmap_table(font_data: &[u8]) -> HashMap<u32, u16> {
    let mut result = HashMap::new();
    if font_data.len() < 12 { return result; }

    let num_tables = u16::from_be_bytes([font_data[4], font_data[5]]) as usize;

    // Find the cmap table
    let mut cmap_offset = 0usize;
    let mut cmap_length = 0usize;
    for i in 0..num_tables {
        let rec = 12 + i * 16;
        if rec + 16 > font_data.len() { break; }
        if &font_data[rec..rec + 4] == b"cmap" {
            cmap_offset = u32::from_be_bytes([
                font_data[rec + 8], font_data[rec + 9], font_data[rec + 10], font_data[rec + 11],
            ]) as usize;
            cmap_length = u32::from_be_bytes([
                font_data[rec + 12], font_data[rec + 13], font_data[rec + 14], font_data[rec + 15],
            ]) as usize;
            break;
        }
    }
    if cmap_offset == 0 || cmap_offset + 4 > font_data.len() { return result; }

    let cmap = &font_data[cmap_offset..font_data.len().min(cmap_offset + cmap_length)];
    if cmap.len() < 4 { return result; }

    let num_subtables = u16::from_be_bytes([cmap[2], cmap[3]]) as usize;

    // Prefer: Platform 3 Encoding 10 (Windows UCS-4, Format 12) > Platform 3 Encoding 1 (Windows BMP, Format 4)
    // > Platform 0 (Unicode)
    let mut best_offset = 0usize;
    let mut best_priority = 0u8;

    for i in 0..num_subtables {
        let rec = 4 + i * 8;
        if rec + 8 > cmap.len() { break; }
        let platform = u16::from_be_bytes([cmap[rec], cmap[rec + 1]]);
        let encoding = u16::from_be_bytes([cmap[rec + 2], cmap[rec + 3]]);
        let offset = u32::from_be_bytes([cmap[rec + 4], cmap[rec + 5], cmap[rec + 6], cmap[rec + 7]]) as usize;

        let priority = match (platform, encoding) {
            (3, 10) => 4, // Windows UCS-4 (best, supports all Unicode)
            (0, 4) => 3,  // Unicode full
            (3, 1) => 2,  // Windows BMP
            (0, 3) => 2,  // Unicode BMP
            (0, _) => 1,  // Any Unicode platform
            _ => 0,
        };

        if priority > best_priority {
            best_priority = priority;
            best_offset = offset;
        }
    }

    if best_offset == 0 || best_offset + 2 > cmap.len() { return result; }

    let subtable = &cmap[best_offset..];
    if subtable.len() < 2 { return result; }
    let format = u16::from_be_bytes([subtable[0], subtable[1]]);

    match format {
        4 => parse_cmap_format4(subtable, &mut result),
        12 => parse_cmap_format12(subtable, &mut result),
        _ => {
            eprintln!("Unsupported cmap format: {}", format);
        }
    }

    result
}

fn parse_cmap_format4(data: &[u8], result: &mut HashMap<u32, u16>) {
    if data.len() < 14 { return; }
    let seg_count = u16::from_be_bytes([data[6], data[7]]) as usize / 2;
    let header_size = 14;

    if data.len() < header_size + seg_count * 8 { return; }

    let end_codes = &data[header_size..];
    let start_codes = &data[header_size + seg_count * 2 + 2..]; // +2 for reservedPad
    let id_deltas = &data[header_size + seg_count * 4 + 2..];
    let id_range_offsets_start = header_size + seg_count * 6 + 2;
    let id_range_offsets = &data[id_range_offsets_start..];

    for seg in 0..seg_count {
        let off = seg * 2;
        if off + 2 > end_codes.len() || off + 2 > start_codes.len() { break; }
        let end_code = u16::from_be_bytes([end_codes[off], end_codes[off + 1]]);
        let start_code = u16::from_be_bytes([start_codes[off], start_codes[off + 1]]);
        if off + 2 > id_deltas.len() || off + 2 > id_range_offsets.len() { break; }
        let id_delta = i16::from_be_bytes([id_deltas[off], id_deltas[off + 1]]);
        let id_range_offset = u16::from_be_bytes([id_range_offsets[off], id_range_offsets[off + 1]]);

        if start_code == 0xFFFF { break; }

        for code in start_code..=end_code {
            let gid = if id_range_offset == 0 {
                (code as i32 + id_delta as i32) as u16
            } else {
                // idRangeOffset points into the glyphIdArray
                let glyph_idx_offset = id_range_offsets_start + off
                    + id_range_offset as usize
                    + (code - start_code) as usize * 2;
                if glyph_idx_offset + 2 <= data.len() {
                    let glyph_id = u16::from_be_bytes([data[glyph_idx_offset], data[glyph_idx_offset + 1]]);
                    if glyph_id == 0 {
                        0
                    } else {
                        (glyph_id as i32 + id_delta as i32) as u16
                    }
                } else {
                    0
                }
            };
            if gid != 0 {
                result.insert(code as u32, gid);
            }
        }
    }
}

fn parse_cmap_format12(data: &[u8], result: &mut HashMap<u32, u16>) {
    if data.len() < 16 { return; }
    let num_groups = u32::from_be_bytes([data[12], data[13], data[14], data[15]]) as usize;

    for i in 0..num_groups {
        let off = 16 + i * 12;
        if off + 12 > data.len() { break; }
        let start_code = u32::from_be_bytes([data[off], data[off + 1], data[off + 2], data[off + 3]]);
        let end_code = u32::from_be_bytes([data[off + 4], data[off + 5], data[off + 6], data[off + 7]]);
        let start_gid = u32::from_be_bytes([data[off + 8], data[off + 9], data[off + 10], data[off + 11]]);

        for code in start_code..=end_code {
            let gid = start_gid + (code - start_code);
            if gid != 0 && gid <= 0xFFFF {
                result.insert(code, gid as u16);
            }
        }
    }
}

fn has_cff_table(data: &[u8]) -> bool {
    if data.len() < 12 { return false; }
    let num_tables = u16::from_be_bytes([data[4], data[5]]) as usize;
    for i in 0..num_tables {
        let off = 12 + i * 16;
        if off + 4 > data.len() { return false; }
        if &data[off..off + 4] == b"CFF " {
            return true;
        }
    }
    false
}

/// Extract raw CFF data from an OTF font file.
fn extract_cff_from_otf(data: &[u8]) -> Option<Vec<u8>> {
    if data.len() < 12 { return None; }
    let num_tables = u16::from_be_bytes([data[4], data[5]]) as usize;
    for i in 0..num_tables {
        let rec_off = 12 + i * 16;
        if rec_off + 16 > data.len() { return None; }
        if &data[rec_off..rec_off + 4] == b"CFF " {
            let offset = u32::from_be_bytes([
                data[rec_off + 8], data[rec_off + 9], data[rec_off + 10], data[rec_off + 11],
            ]) as usize;
            let length = u32::from_be_bytes([
                data[rec_off + 12], data[rec_off + 13], data[rec_off + 14], data[rec_off + 15],
            ]) as usize;
            if offset + length <= data.len() {
                return Some(data[offset..offset + length].to_vec());
            }
        }
    }
    None
}
