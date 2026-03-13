use std::collections::HashMap;
use std::fs;
use oxidocs_core::ir::{Block, Paragraph, Table, Alignment};
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
            x, y, text, font_name, font_size, fill_color: color,
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
    let base = family.unwrap_or("Helvetica");
    // CJK fonts pass through as-is (CIDFont embedding handles them)
    if base.contains("Gothic") || base.contains("Mincho") || base.contains("游")
        || base.contains("ＭＳ") || base.contains("メイリオ") || base.contains("ヒラギノ")
    {
        return base.to_string();
    }
    if bold {
        match base {
            "Helvetica" => "Helvetica-Bold".to_string(),
            "Times New Roman" | "Times" => "Times-Bold".to_string(),
            "Courier New" | "Courier" => "Courier-Bold".to_string(),
            other => other.to_string(),
        }
    } else {
        match base {
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
    let page0 = doc.pages.first().expect("No pages in document");
    let w = page0.size.width as f64;
    let h = page0.size.height as f64;
    let m = (
        page0.margin.top as f64,
        page0.margin.bottom as f64,
        page0.margin.left as f64,
        page0.margin.right as f64,
    );

    // Build embedded fonts FIRST so we can use widths for accurate layout
    let embedded_fonts = build_embedded_fonts(&doc);

    // Build a Unicode → width lookup (in 1/1000 em units) from embedded font data
    let font_width_map: HashMap<u32, u16> = embedded_fonts.values().next()
        .map(|ef| {
            let mut map = HashMap::new();
            for (&unicode, &cid) in &ef.unicode_to_gid {
                if let Some(&w) = ef.cid_widths.get(&cid) {
                    map.insert(unicode, w);
                }
            }
            map
        })
        .unwrap_or_default();

    let mut builder = PdfBuilder::new(w, h, m, font_width_map);

    // Collect all blocks from all sections
    let all_blocks: Vec<&Block> = doc.pages.iter().flat_map(|p| &p.blocks).collect();
    let total_blocks = all_blocks.len();

    let ml = builder.margin_left;
    let cw = builder.content_width();

    let mut cover_done = false;
    for (i, block) in all_blocks.iter().enumerate() {
        match block {
            Block::Paragraph(para) => {
                let role = classify_paragraph(para, i, total_blocks);

                // Render cover page elements
                if !cover_done && matches!(role, ParaRole::CoverTitle | ParaRole::CoverSubtitle | ParaRole::CoverAuthor) {
                    render_cover_element(&mut builder, para, role, ml, cw);
                    // After author line, insert page break
                    if role == ParaRole::CoverAuthor {
                        // Add decorative line on cover
                        let line_y = builder.cursor_y + 20.0;
                        builder.add_line(
                            ml + cw * 0.2, line_y,
                            ml + cw * 0.8, line_y,
                            1.0, Color::Rgb(0.18, 0.25, 0.34),
                        );
                        builder.new_page();
                        cover_done = true;
                    }
                    continue;
                }
                cover_done = true;

                render_paragraph_styled(&mut builder, para, role, ml, cw);
            }
            Block::Table(table) => {
                cover_done = true;
                render_table(&mut builder, table, ml, cw);
            }
            Block::Image(_) | Block::UnsupportedElement(_) => {}
        }
    }

    // Add page numbers to all pages
    let pages = builder.finish();
    let _total_pages = pages.len();
    let pages: Vec<Page> = pages.into_iter().enumerate().map(|(i, mut page)| {
        // Add page number centered at bottom (top-down: large y = near bottom)
        let page_num = format!("- {} -", i + 1);
        let num_width = estimate_text_width(&page_num, 9.0);
        page.contents.push(ContentElement::Text(TextSpan {
            x: (w - num_width) / 2.0,
            y: h - m.1 * 0.5,
            text: page_num,
            font_name: "Helvetica".to_string(),
            font_size: 9.0,
            fill_color: Color::Gray(0.4),
        }));
        page
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
fn render_paragraph_styled(builder: &mut PdfBuilder, para: &Paragraph, role: ParaRole, x_offset: f64, available_width: f64) {
    let full_text: String = para.runs.iter().map(|r| r.text.as_str()).collect();

    // Empty paragraph — add minimal spacing
    if full_text.trim().is_empty() && para.style.list_marker.is_none() {
        let gap = para.style.space_before.unwrap_or(0.0) as f64
            + para.style.space_after.unwrap_or(4.0) as f64;
        builder.cursor_y += gap.max(6.0);
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
            render_section_header(builder, para, x_offset, available_width);
            return;
        }
        ParaRole::SubHeader => {
            render_sub_header(builder, para, x_offset, available_width);
            return;
        }
        ParaRole::SubSubHeader => {
            render_sub_sub_header(builder, para, x_offset, available_width);
            return;
        }
        ParaRole::CreditFooter => {
            render_credit_footer(builder, para, x_offset, available_width);
            return;
        }
        _ => {} // Body and ListItem handled below
    }

    // --- Regular body / list rendering ---
    let font_size = default_size;
    // Use line_spacing from IR if available
    let line_spacing_mult = para.style.line_spacing.unwrap_or(0.0) as f64;
    let line_height = if line_spacing_mult > 0.5 && line_spacing_mult < 5.0 {
        font_size * line_spacing_mult
    } else {
        font_size * 1.35
    };

    let space_before = para.style.space_before.unwrap_or(0.0) as f64;
    let space_after = para.style.space_after.unwrap_or(2.0) as f64;

    // Indentation
    let indent_left = para.style.indent_left.unwrap_or(0.0) as f64;
    let indent_first = para.style.indent_first_line.unwrap_or(0.0) as f64;

    builder.cursor_y += space_before;

    if builder.needs_page_break(line_height + 10.0) {
        builder.new_page();
    }

    // List marker handling (both IR list_marker and text-based ・)
    let mut marker_offset = 0.0;
    let mut effective_indent_left = indent_left;

    if role == ParaRole::ListItem {
        if let Some(marker) = &para.style.list_marker {
            let marker_indent = para.style.list_indent.unwrap_or(18.0) as f64;
            marker_offset = marker_indent;
            let marker_x = x_offset + indent_left;
            let marker_font = resolve_font(default_font, false);
            builder.add_text(
                marker_x, builder.cursor_y,
                marker.clone(), marker_font, font_size, Color::Gray(0.0),
            );
        } else if full_text.trim().starts_with('・') || full_text.trim().starts_with('※') {
            effective_indent_left = effective_indent_left.max(8.0);
        }
    }

    let text_x_start = x_offset + effective_indent_left + marker_offset;
    let text_width = (available_width - effective_indent_left - marker_offset).max(50.0);

    let lines = wrap_runs_into_lines(&para.runs, font_size, text_width, indent_first);

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
            .map(|lr| builder.text_width(&lr.text, lr.font_size))
            .sum();

        let align_offset = match para.alignment {
            Alignment::Center => (text_width - line_width) / 2.0,
            Alignment::Right => text_width - line_width,
            _ => 0.0,
        };

        let mut run_x = line_x + align_offset.max(0.0);
        let text_y = builder.cursor_y;

        for lr in line_runs {
            if !lr.text.is_empty() {
                builder.add_text(
                    run_x, text_y,
                    lr.text.clone(), lr.font_name.clone(), lr.font_size, lr.color,
                );
                run_x += builder.text_width(&lr.text, lr.font_size);
            }
        }

        builder.cursor_y += line_height;
    }

    builder.cursor_y += space_after;
}

/// Render ① ② ③ section headers with colored background bar
fn render_section_header(builder: &mut PdfBuilder, para: &Paragraph, x_offset: f64, available_width: f64) {
    let text: String = para.runs.iter().map(|r| r.text.as_str()).collect();
    let font_size = 14.0;
    let line_height = font_size * 1.6;
    let space_before = 20.0;
    let space_after = 10.0;

    let font_family = para.runs.first()
        .and_then(|r| r.style.font_family.as_deref().or(r.style.font_family_east_asia.as_deref()));
    let font_name = resolve_font(font_family, true);

    builder.cursor_y += space_before;

    if builder.needs_page_break(line_height + 30.0) {
        builder.new_page();
    }

    // Draw background bar (top-down: bar starts at cursor_y)
    let bar_height = line_height + 6.0;
    let bar_y = builder.cursor_y - 2.0;
    builder.add_rect_fill(x_offset, bar_y, available_width, bar_height, Color::Rgb(0.18, 0.25, 0.34));

    // Text in white on dark background
    let text_y = builder.cursor_y + 2.0;
    let text_x = x_offset + 10.0;
    builder.add_text(text_x, text_y, text, font_name, font_size, Color::Rgb(1.0, 1.0, 1.0));

    builder.cursor_y += bar_height + space_after;
}

/// Render "N-N." sub-headers with colored text and underline
fn render_sub_header(builder: &mut PdfBuilder, para: &Paragraph, x_offset: f64, available_width: f64) {
    let text: String = para.runs.iter().map(|r| r.text.as_str()).collect();
    let font_size = 12.0;
    let line_height = font_size * 1.5;
    let space_before = 14.0;
    let space_after = 6.0;

    let font_family = para.runs.first()
        .and_then(|r| r.style.font_family.as_deref().or(r.style.font_family_east_asia.as_deref()));
    let font_name = resolve_font(font_family, true);
    let color = Color::Rgb(0.18, 0.25, 0.34);

    builder.cursor_y += space_before;

    if builder.needs_page_break(line_height + 20.0) {
        builder.new_page();
    }

    builder.add_text(x_offset, builder.cursor_y, text, font_name, font_size, color);

    // Underline below text
    let underline_y = builder.cursor_y + line_height - 2.0;
    builder.add_line(x_offset, underline_y, x_offset + available_width, underline_y, 0.75, color);

    builder.cursor_y += line_height + space_after;
}

/// Render sub-sub-headers (bold topic introductions)
fn render_sub_sub_header(builder: &mut PdfBuilder, para: &Paragraph, x_offset: f64, available_width: f64) {
    let text: String = para.runs.iter().map(|r| r.text.as_str()).collect();
    let font_size = 11.0;
    let line_height = font_size * 1.5;
    let space_before = 10.0;
    let space_after = 4.0;

    let font_family = para.runs.first()
        .and_then(|r| r.style.font_family.as_deref().or(r.style.font_family_east_asia.as_deref()));
    let font_name = resolve_font(font_family, true);
    let color = Color::Rgb(0.18, 0.25, 0.34);

    builder.cursor_y += space_before;

    if builder.needs_page_break(line_height + 10.0) {
        builder.new_page();
    }

    // Small colored square bullet before text
    let sq_size = 6.0;
    let sq_y = builder.cursor_y + (font_size - sq_size) / 2.0;
    builder.add_rect_fill(x_offset, sq_y, sq_size, sq_size, color);

    let lines = wrap_text(&text, font_size, available_width - 12.0);
    for line in &lines {
        builder.add_text(x_offset + 10.0, builder.cursor_y, line.clone(), font_name.clone(), font_size, color);
        builder.cursor_y += line_height;
    }

    builder.cursor_y += space_after;
}

/// Render the credit footer line
fn render_credit_footer(builder: &mut PdfBuilder, para: &Paragraph, x_offset: f64, available_width: f64) {
    let text: String = para.runs.iter().map(|r| r.text.as_str()).collect();
    let font_size = 8.5;

    let font_family = para.runs.first()
        .and_then(|r| r.style.font_family.as_deref().or(r.style.font_family_east_asia.as_deref()));
    let font_name = resolve_font(font_family, false);

    builder.cursor_y += 20.0;

    // Thin line above
    builder.add_line(
        x_offset + available_width * 0.3, builder.cursor_y,
        x_offset + available_width * 0.7, builder.cursor_y,
        0.5, Color::Gray(0.6),
    );
    builder.cursor_y += 10.0;

    let text_w = builder.text_width(&text, font_size);
    let x = x_offset + available_width - text_w;
    builder.add_text(x, builder.cursor_y, text, font_name, font_size, Color::Gray(0.4));
    builder.cursor_y += font_size * 2.0;
}

/// A styled text fragment for a single line
struct LineRun {
    text: String,
    font_name: String,
    font_size: f64,
    color: Color,
}

/// Break runs into wrapped lines, preserving per-run styling
fn wrap_runs_into_lines(
    runs: &[oxidocs_core::ir::Run],
    default_font_size: f64,
    max_width: f64,
    first_line_indent: f64,
) -> Vec<Vec<LineRun>> {
    let mut lines: Vec<Vec<LineRun>> = Vec::new();
    let mut current_line: Vec<LineRun> = Vec::new();
    let mut current_width = first_line_indent;
    let effective_max = max_width;

    for run in runs {
        let font_size = run.style.font_size.unwrap_or(default_font_size as f32) as f64;
        let bold = run.style.bold;
        let font_family = run.style.font_family.as_deref()
            .or(run.style.font_family_east_asia.as_deref());
        let font_name = resolve_font(font_family, bold);
        let color = run.style.color.as_deref()
            .and_then(parse_hex_color)
            .unwrap_or(Color::Gray(0.0));

        // Process text character by character for line wrapping
        let mut buf = String::new();
        let mut buf_width = 0.0;

        for ch in run.text.chars() {
            if ch == '\n' || ch == '\r' {
                // Explicit line break
                if !buf.is_empty() {
                    current_line.push(LineRun {
                        text: buf.clone(), font_name: font_name.clone(),
                        font_size, color,
                    });
                    buf.clear();
                    buf_width = 0.0;
                }
                lines.push(std::mem::take(&mut current_line));
                current_width = 0.0;
                continue;
            }

            let ch_width = char_width(ch, font_size);

            if current_width + buf_width + ch_width > effective_max && !(current_line.is_empty() && buf.is_empty()) {
                // Wrap: flush buffer to current line, start new line
                if !buf.is_empty() {
                    current_line.push(LineRun {
                        text: buf.clone(), font_name: font_name.clone(),
                        font_size, color,
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
            current_line.push(LineRun {
                text: buf, font_name: font_name.clone(),
                font_size, color,
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

fn render_table(builder: &mut PdfBuilder, table: &Table, x_offset: f64, available_width: f64) {
    let num_cols = table.rows.first().map(|r| r.cells.len()).unwrap_or(0);
    if num_cols == 0 { return; }

    // Calculate column widths
    let col_widths: Vec<f64> = if let Some(first_row) = table.rows.first() {
        let specified: Vec<Option<f64>> = first_row.cells.iter()
            .map(|c| c.width.map(|w| w as f64))
            .collect();
        let total_specified: f64 = specified.iter().filter_map(|w| *w).sum();
        if total_specified > 0.0 {
            // Scale proportionally to fill available width
            let scale = available_width / total_specified;
            specified.iter().map(|w| w.map(|v| v * scale).unwrap_or(available_width / num_cols as f64)).collect()
        } else {
            vec![available_width / num_cols as f64; num_cols]
        }
    } else {
        vec![available_width / num_cols as f64; num_cols]
    };

    let cell_padding = 5.0;
    let row_min_height = 22.0;

    builder.cursor_y += 6.0; // gap before table

    // Detect if first row is a header (all cells centered)
    let first_row_is_header = table.rows.first().map(|r| {
        r.cells.iter().all(|c| {
            c.blocks.iter().all(|b| matches!(b, Block::Paragraph(p) if p.alignment == Alignment::Center))
        })
    }).unwrap_or(false);

    for (row_idx, row) in table.rows.iter().enumerate() {
        let is_header_row = row_idx == 0 && first_row_is_header;

        // Calculate row height based on content
        let mut row_height = row.height.map(|h| h as f64).unwrap_or(row_min_height);

        for (ci, cell) in row.cells.iter().enumerate() {
            if ci >= col_widths.len() { break; }
            let cell_w = col_widths[ci] - cell_padding * 2.0;
            let cell_h = estimate_cell_height(cell, cell_w);
            if cell_h + cell_padding * 2.0 > row_height {
                row_height = cell_h + cell_padding * 2.0;
            }
        }

        if builder.needs_page_break(row_height + 2.0) {
            builder.new_page();
        }

        let row_top = builder.cursor_y;
        let row_bottom = row_top + row_height;

        // Draw row background (top-down: row_top is smaller y, row_bottom is larger y)
        if is_header_row {
            builder.add_rect_fill(x_offset, row_top, available_width, row_height,
                Color::Rgb(0.18, 0.25, 0.34));
        } else if row_idx % 2 == 0 && !is_header_row {
            builder.add_rect_fill(x_offset, row_top, available_width, row_height,
                Color::Rgb(0.96, 0.96, 0.98));
        }

        // Draw cells
        let mut cell_x = x_offset;
        for (ci, cell) in row.cells.iter().enumerate() {
            if ci >= col_widths.len() { break; }
            let cw = col_widths[ci];

            if let Some(ref shade) = cell.shading {
                if let Some(bg_color) = parse_hex_color(shade) {
                    builder.add_rect_fill(cell_x, row_top, cw, row_height, bg_color);
                }
            }

            // Cell borders
            let border_color = Color::Rgb(0.75, 0.78, 0.82);
            let bw = 0.5;
            builder.add_line(cell_x, row_top, cell_x + cw, row_top, bw, border_color);
            builder.add_line(cell_x, row_bottom, cell_x + cw, row_bottom, bw, border_color);
            builder.add_line(cell_x, row_top, cell_x, row_bottom, bw, border_color);
            builder.add_line(cell_x + cw, row_top, cell_x + cw, row_bottom, bw, border_color);

            // Cell text with vertical centering
            let cell_content_x = cell_x + cell_padding;
            let cell_content_width = cw - cell_padding * 2.0;

            let total_text_h = cell.blocks.iter().map(|b| {
                if let Block::Paragraph(para) = b {
                    let fs = para.runs.first().and_then(|r| r.style.font_size).unwrap_or(9.0) as f64;
                    let text: String = para.runs.iter().map(|r| r.text.as_str()).collect();
                    let lines = wrap_text(&text, fs, cell_content_width);
                    lines.len() as f64 * fs * 1.4
                } else { 0.0 }
            }).sum::<f64>();

            let v_offset = ((row_height - total_text_h) / 2.0).max(cell_padding);
            let mut cell_text_y = row_top + v_offset;

            for block in &cell.blocks {
                if let Block::Paragraph(para) = block {
                    let font_size = para.runs.first()
                        .and_then(|r| r.style.font_size)
                        .or(para.style.default_run_style.as_ref().and_then(|rs| rs.font_size))
                        .unwrap_or(9.0) as f64;
                    let line_height = font_size * 1.4;

                    let text: String = para.runs.iter().map(|r| r.text.as_str()).collect();
                    let wrapped = wrap_text(&text, font_size, cell_content_width);

                    for line in &wrapped {
                        if line.is_empty() { continue; }
                        let font_family = para.runs.first()
                            .and_then(|r| r.style.font_family.as_deref()
                                .or(r.style.font_family_east_asia.as_deref()));
                        let bold = para.runs.first().map(|r| r.style.bold).unwrap_or(false)
                            || para.style.default_run_style.as_ref().map(|rs| rs.bold).unwrap_or(false);
                        let font_name = resolve_font(font_family, bold || is_header_row);

                        let text_color = if is_header_row {
                            Color::Rgb(1.0, 1.0, 1.0)
                        } else {
                            para.runs.first()
                                .and_then(|r| r.style.color.as_deref())
                                .and_then(parse_hex_color)
                                .unwrap_or(Color::Gray(0.1))
                        };

                        let line_w = builder.text_width(line, font_size);
                        let align_off = match para.alignment {
                            Alignment::Center => (cell_content_width - line_w) / 2.0,
                            Alignment::Right => cell_content_width - line_w,
                            _ => 0.0,
                        };

                        builder.add_text(
                            cell_content_x + align_off.max(0.0),
                            cell_text_y,
                            line.clone(), font_name, font_size, text_color,
                        );
                        cell_text_y += line_height;
                    }
                }
            }

            cell_x += cw;
        }

        builder.cursor_y = row_bottom;
    }

    builder.cursor_y += 8.0; // gap after table
}

fn estimate_cell_height(cell: &oxidocs_core::ir::TableCell, content_width: f64) -> f64 {
    let mut h = 0.0;
    for block in &cell.blocks {
        if let Block::Paragraph(para) = block {
            let font_size = para.style.default_run_style.as_ref()
                .and_then(|rs| rs.font_size)
                .or_else(|| para.runs.first().and_then(|r| r.style.font_size))
                .unwrap_or(9.0) as f64;
            let line_height = font_size * 1.4;
            let text: String = para.runs.iter().map(|r| r.text.as_str()).collect();
            let lines = wrap_text(&text, font_size, content_width);
            h += lines.len() as f64 * line_height;
        }
    }
    h.max(14.0)
}

// --- Font Embedding ---

/// System font paths to search for CJK fonts
const CJK_FONT_PATHS: &[&str] = &[
    "/usr/share/fonts/opentype/noto/NotoSansCJK-Regular.ttc",
    "/usr/share/fonts/opentype/noto/NotoSansCJK-Bold.ttc",
    "/System/Library/Fonts/ヒラギノ角ゴシック W3.ttc",
    "/System/Library/Fonts/HiraginoSans-W3.ttc",
    "C:\\Windows\\Fonts\\YuGothR.ttc",
    "C:\\Windows\\Fonts\\msgothic.ttc",
    "C:\\Windows\\Fonts\\meiryo.ttc",
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

    if !needs_cjk_font {
        return fonts;
    }

    // Collect all unique characters used in the document (for font subsetting)
    let mut used_chars: std::collections::BTreeSet<char> = std::collections::BTreeSet::new();
    for page in &doc.pages {
        for block in &page.blocks {
            collect_chars_from_block(block, &mut used_chars);
        }
    }

    // Try to find, subset, and load a CJK font from the system
    let font_result = match find_and_subset_cjk_font(&used_chars) {
        Some(result) => result,
        None => {
            eprintln!("Warning: No CJK system font found. PDF may not render Japanese text correctly.");
            return fonts;
        }
    };

    eprintln!("Embedding CJK font ({} bytes, {} glyph mappings)",
        font_result.data.len(), font_result.unicode_to_gid.len());

    // Map all CJK font names used in the document to this embedded font
    for name in &font_names_used {
        let resolved = resolve_font(Some(name), false);
        fonts.insert(resolved.clone(), font_result.clone());
        // Also insert bold variant
        let resolved_bold = resolve_font(Some(name), true);
        if resolved_bold != resolved {
            fonts.insert(resolved_bold, font_result.clone());
        }
    }

    fonts
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

/// Find a CJK font, subset it to only used characters, and prepare for PDF embedding.
fn find_and_subset_cjk_font(used_chars: &std::collections::BTreeSet<char>) -> Option<EmbeddedFont> {
    // Find a TTC/OTF font on the system
    let font_path = CJK_FONT_PATHS.iter().find(|p| std::path::Path::new(p).exists())?;
    eprintln!("Found CJK font: {}", font_path);

    // Build the character string for subsetting
    let char_string: String = used_chars.iter().collect();

    // Use Python fonttools to subset the font AND extract Unicode→CID mapping
    let subset_path = "/tmp/oxi-font-subset.otf";
    let cidmap_path = "/tmp/oxi-font-cidmap.json";
    let widths_path = "/tmp/oxi-font-widths.json";
    let python_script = format!(
        r#"
import sys, json
from fontTools.ttLib import TTCollection, TTFont
from fontTools.subset import Subsetter, Options

font_path = "{font_path}"
data = open(font_path, 'rb').read()

if data[:4] == b'ttcf':
    ttc = TTCollection(font_path)
    font = ttc[0]
else:
    font = TTFont(font_path)

chars = sys.stdin.read()
opts = Options()
subsetter = Subsetter(options=opts)
subsetter.populate(text=chars)
subsetter.subset(font)

font.save("{subset_path}")

# Extract Unicode → CID mapping and glyph widths
cmap = font.getBestCmap()
unicode_to_cid = {{}}
cid_widths = {{}}

# Get units per em for width normalization
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
        # Get advance width from hmtx table
        if glyph_name in hmtx.metrics:
            advance, _ = hmtx.metrics[glyph_name]
            # Normalize to 1000 units (PDF convention)
            w = int(advance * 1000 / upem)
            cid_widths[cid] = w

with open("{cidmap_path}", 'w') as f:
    json.dump(unicode_to_cid, f)

with open("{widths_path}", 'w') as f:
    json.dump(cid_widths, f)

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
                eprintln!("Font subsetting failed: {}", String::from_utf8_lossy(&output.stderr));
                return find_system_cjk_font_full();
            }
            eprintln!("Font subset: {}", String::from_utf8_lossy(&output.stdout).trim());
        }
        Err(e) => {
            eprintln!("Python not available for font subsetting: {}", e);
            return find_system_cjk_font_full();
        }
    }

    // Read the subset OTF
    let otf_data = fs::read(subset_path).ok()?;
    eprintln!("Subset font: {} bytes ({} KB)", otf_data.len(), otf_data.len() / 1024);

    // Read the CID mapping (from Python fonttools, which correctly resolves CID-keyed CFF names)
    let unicode_to_gid = if let Ok(json_str) = fs::read_to_string(cidmap_path) {
        parse_cidmap_json(&json_str)
    } else {
        // Fallback: parse cmap from binary (works for non-CID fonts)
        parse_cmap_table(&otf_data)
    };
    eprintln!("CID mapping: {} entries", unicode_to_gid.len());

    // Read glyph widths
    let cid_widths = if let Ok(json_str) = fs::read_to_string(widths_path) {
        parse_cidwidths_json(&json_str)
    } else {
        HashMap::new()
    };
    eprintln!("Glyph widths: {} entries", cid_widths.len());

    // Extract CFF data for embedding
    let is_cff = otf_data.starts_with(b"OTTO") || has_cff_table(&otf_data);
    if is_cff {
        if let Some(cff) = extract_cff_from_otf(&otf_data) {
            return Some(EmbeddedFont {
                data: cff,
                format: FontFormat::OpenTypeCff,
                unicode_to_gid,
                cid_widths,
            });
        }
    }

    Some(EmbeddedFont {
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
                    return Some(EmbeddedFont {
                        data: cff,
                        format: FontFormat::OpenTypeCff,
                        unicode_to_gid,
                        cid_widths: HashMap::new(),
                    });
                }
            }

            // TrueType
            return Some(EmbeddedFont {
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
