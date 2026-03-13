use std::collections::HashMap;
use std::io::{Cursor, Read};

use quick_xml::events::Event;
use quick_xml::reader::Reader;
use zip::ZipArchive;

use super::numbering::{parse_numbering, NumberingDefinitions};
use super::relationships::{parse_relationships, Relationship};
use super::styles::parse_styles;
use super::theme::{parse_theme, ThemeColors};
use super::ParseError;
use crate::ir::{*, VerticalAlign};

pub struct OoxmlParser {
    archive: ZipArchive<Cursor<Vec<u8>>>,
}

/// Context passed through parsing functions for resource resolution
struct ParseContext {
    /// Relationship ID -> Relationship mapping
    _rels: HashMap<String, Relationship>,
    /// Relationship ID -> binary data (images, etc.)
    media: HashMap<String, Vec<u8>>,
    /// Relationship ID -> content type (e.g., "image/png")
    media_types: HashMap<String, String>,
    /// Relationship ID -> hyperlink URL (external links)
    hyperlinks: HashMap<String, String>,
    /// Numbering definitions from word/numbering.xml
    numbering: NumberingDefinitions,
    /// Counters for numbered lists: (numId, ilvl) -> current count
    list_counters: std::cell::RefCell<HashMap<(String, u8), u32>>,
    /// Footnote ID -> paragraphs (from word/footnotes.xml)
    footnotes: HashMap<String, Vec<Block>>,
    /// Endnote ID -> paragraphs (from word/endnotes.xml)
    endnotes: HashMap<String, Vec<Block>>,
    /// Comment ID -> Comment (from word/comments.xml)
    comments: HashMap<String, Comment>,
    /// Theme colors from theme1.xml
    theme: ThemeColors,
}

impl OoxmlParser {
    pub fn new(data: &[u8]) -> Result<Self, ParseError> {
        let cursor = Cursor::new(data.to_vec());
        let archive = ZipArchive::new(cursor)?;
        Ok(Self { archive })
    }

    pub fn parse(mut self) -> Result<Document, ParseError> {
        let styles = self.parse_styles()?;
        let numbering = self.parse_numbering()?;
        let ctx = self.build_context(numbering)?;
        let sections = self.parse_document_xml(&ctx, &styles)?;
        let metadata = self.parse_metadata();

        let mut pages = Vec::new();
        for section in sections {
            let header = self.parse_header_footer_blocks(&section.properties.header_refs, &ctx, &styles);
            let footer = self.parse_header_footer_blocks(&section.properties.footer_refs, &ctx, &styles);

            // Collect referenced footnotes and endnotes for this section
            let mut footnotes_list = Vec::new();
            let mut endnotes_list = Vec::new();
            collect_note_refs(&section.blocks, &ctx, &mut footnotes_list, &mut endnotes_list);
            footnotes_list.sort_by_key(|f| f.number);
            endnotes_list.sort_by_key(|f| f.number);

            pages.push(Page {
                blocks: section.blocks,
                size: section.properties.page_size,
                margin: section.properties.margin,
                grid_line_pitch: section.properties.grid_line_pitch,
                header,
                footer,
                footnotes: footnotes_list,
                endnotes: endnotes_list,
                floating_images: section.floating_images,
                text_boxes: section.text_boxes,
                columns: section.properties.columns,
            });
        }

        // Collect all comments referenced in the document
        let all_comments: Vec<Comment> = ctx.comments.values().cloned().collect();

        Ok(Document {
            pages,
            styles,
            metadata,
            comments: all_comments,
        })
    }

    /// Parse header or footer XML parts referenced by relationship IDs
    fn parse_header_footer_blocks(
        &mut self,
        ref_ids: &[String],
        ctx: &ParseContext,
        styles: &StyleSheet,
    ) -> Vec<Block> {
        let mut blocks = Vec::new();
        for ref_id in ref_ids {
            // Look up the relationship target path
            let target = ctx._rels.get(ref_id)
                .map(|r| r.target.clone());
            if let Some(target) = target {
                let part_path = if target.starts_with('/') {
                    target[1..].to_string()
                } else {
                    format!("word/{}", target)
                };
                if let Ok(xml) = self.read_part(&part_path) {
                    if let Ok(parsed) = parse_header_footer_xml(&xml, ctx, styles) {
                        blocks.extend(parsed);
                    }
                }
            }
        }
        blocks
    }

    fn read_part(&mut self, name: &str) -> Result<String, ParseError> {
        let mut file = self
            .archive
            .by_name(name)
            .map_err(|_| ParseError::MissingPart(name.to_string()))?;
        let mut contents = String::new();
        file.read_to_string(&mut contents)?;
        Ok(contents)
    }

    fn read_binary_part(&mut self, name: &str) -> Result<Vec<u8>, ParseError> {
        let mut file = self
            .archive
            .by_name(name)
            .map_err(|_| ParseError::MissingPart(name.to_string()))?;
        let mut contents = Vec::new();
        file.read_to_end(&mut contents)?;
        Ok(contents)
    }

    fn parse_numbering(&mut self) -> Result<NumberingDefinitions, ParseError> {
        match self.read_part("word/numbering.xml") {
            Ok(xml) => parse_numbering(&xml),
            Err(ParseError::MissingPart(_)) => Ok(NumberingDefinitions::default()),
            Err(e) => Err(e),
        }
    }

    fn build_context(&mut self, numbering: NumberingDefinitions) -> Result<ParseContext, ParseError> {
        // Parse relationships
        let rels = match self.read_part("word/_rels/document.xml.rels") {
            Ok(xml) => parse_relationships(&xml)?,
            Err(ParseError::MissingPart(_)) => HashMap::new(),
            Err(e) => return Err(e),
        };

        // Pre-load media and hyperlinks from relationships
        let mut media = HashMap::new();
        let mut media_types = HashMap::new();
        let mut hyperlinks = HashMap::new();
        let image_rel_type = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image";
        let hyperlink_rel_type = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink";
        for (id, rel) in &rels {
            if rel.rel_type == image_rel_type {
                // Validate relationship target path against traversal attacks
                let path = match oxi_common::security::sanitize_rel_target("word", &rel.target) {
                    Ok(p) => p,
                    Err(_) => continue, // Skip suspicious paths
                };
                if let Ok(data) = self.read_binary_part(&path) {
                    // Detect content type from file extension
                    let ct = match rel.target.rsplit('.').next().map(|s| s.to_lowercase()).as_deref() {
                        Some("png") => "image/png",
                        Some("jpg") | Some("jpeg") => "image/jpeg",
                        Some("gif") => "image/gif",
                        Some("bmp") => "image/bmp",
                        Some("svg") => "image/svg+xml",
                        Some("tiff") | Some("tif") => "image/tiff",
                        Some("wmf") => "image/x-wmf",
                        Some("emf") => "image/x-emf",
                        _ => "application/octet-stream",
                    };
                    media_types.insert(id.clone(), ct.to_string());
                    media.insert(id.clone(), data);
                }
            } else if rel.rel_type == hyperlink_rel_type {
                hyperlinks.insert(id.clone(), rel.target.clone());
            }
        }

        // Parse footnotes
        let footnotes = match self.read_part("word/footnotes.xml") {
            Ok(xml) => parse_notes_xml(&xml)?,
            Err(ParseError::MissingPart(_)) => HashMap::new(),
            Err(e) => return Err(e),
        };

        // Parse endnotes
        let endnotes = match self.read_part("word/endnotes.xml") {
            Ok(xml) => parse_notes_xml(&xml)?,
            Err(ParseError::MissingPart(_)) => HashMap::new(),
            Err(e) => return Err(e),
        };

        // Parse comments
        let comments = match self.read_part("word/comments.xml") {
            Ok(xml) => parse_comments_xml(&xml)?,
            Err(ParseError::MissingPart(_)) => HashMap::new(),
            Err(e) => return Err(e),
        };

        // Parse theme colors
        let theme = match self.read_part("word/theme/theme1.xml") {
            Ok(xml) => parse_theme(&xml),
            Err(_) => ThemeColors::default(),
        };

        Ok(ParseContext {
            _rels: rels,
            media,
            media_types,
            hyperlinks,
            numbering,
            list_counters: std::cell::RefCell::new(HashMap::new()),
            footnotes,
            endnotes,
            comments,
            theme,
        })
    }

    fn parse_styles(&mut self) -> Result<StyleSheet, ParseError> {
        match self.read_part("word/styles.xml") {
            Ok(xml) => parse_styles(&xml),
            Err(ParseError::MissingPart(_)) => Ok(StyleSheet::default()),
            Err(e) => Err(e),
        }
    }

    fn parse_metadata(&self) -> DocumentMetadata {
        DocumentMetadata::default()
    }

    fn parse_document_xml(
        &mut self,
        ctx: &ParseContext,
        styles: &StyleSheet,
    ) -> Result<Vec<ParsedSection>, ParseError> {
        let xml = self.read_part("word/document.xml")?;
        parse_body(&xml, ctx, styles)
    }
}

/// A section: blocks + properties. Multiple sections make multiple pages.
struct ParsedSection {
    blocks: Vec<Block>,
    properties: SectionProperties,
    floating_images: Vec<Image>,
    text_boxes: Vec<TextBox>,
}

/// Parse the w:body content of document.xml into sections
fn parse_body(xml: &str, ctx: &ParseContext, styles: &StyleSheet) -> Result<Vec<ParsedSection>, ParseError> {
    let mut reader = Reader::from_str(xml);
    let mut sections: Vec<ParsedSection> = Vec::new();
    let mut current_blocks = Vec::new();
    let mut current_floating_images: Vec<Image> = Vec::new();
    let mut current_text_boxes: Vec<TextBox> = Vec::new();
    let mut final_sect_pr = None;
    let mut depth = 0;
    let mut in_body = false;

    loop {
        match reader.read_event()? {
            Event::Start(e) => {
                let local = local_name(e.name().as_ref());
                match local.as_str() {
                    "body" => {
                        in_body = true;
                        depth = 0;
                    }
                    "p" if in_body && depth == 0 => {
                        let (para, sect_break) = parse_paragraph(&mut reader, ctx, styles)?;
                        current_blocks.push(Block::Paragraph(para));
                        // If this paragraph contained a section break, start a new section
                        if let Some(sp) = sect_break {
                            sections.push(ParsedSection {
                                blocks: std::mem::take(&mut current_blocks),
                                properties: sp,
                                floating_images: std::mem::take(&mut current_floating_images),
                                text_boxes: std::mem::take(&mut current_text_boxes),
                            });
                        }
                    }
                    "tbl" if in_body && depth == 0 => {
                        let table = parse_table(&mut reader, ctx, styles)?;
                        current_blocks.push(Block::Table(table));
                    }
                    "sectPr" if in_body && depth == 0 => {
                        // Final section properties (for the last section)
                        final_sect_pr = Some(parse_section_properties(&mut reader)?);
                    }
                    _ if in_body => {
                        depth += 1;
                    }
                    _ => {}
                }
            }
            Event::End(e) => {
                let local = local_name(e.name().as_ref());
                if local == "body" {
                    in_body = false;
                } else if in_body && depth > 0 {
                    depth -= 1;
                }
            }
            Event::Eof => break,
            _ => {}
        }
    }

    // Remaining blocks form the last section
    let last_sp = final_sect_pr.unwrap_or(SectionProperties {
        page_size: PageSize::default(),
        margin: Margin::default(),
        grid_line_pitch: None,
        header_refs: Vec::new(),
        footer_refs: Vec::new(),
        columns: None,
    });
    sections.push(ParsedSection {
        blocks: current_blocks,
        properties: last_sp,
        floating_images: current_floating_images,
        text_boxes: current_text_boxes,
    });

    Ok(sections)
}

/// Parse a w:p element (paragraph).
/// Returns (Paragraph, optional SectionProperties if this paragraph ends a section).
fn parse_paragraph(reader: &mut Reader<&[u8]>, ctx: &ParseContext, styles: &StyleSheet) -> Result<(Paragraph, Option<SectionProperties>), ParseError> {
    let mut runs = Vec::new();
    let mut images = Vec::new();
    let mut style = ParagraphStyle::default();
    let mut alignment = Alignment::default();
    let mut style_id: Option<String> = None;
    let mut num_pr_ref: Option<NumPrRef> = None;
    let mut para_sect_pr: Option<SectionProperties> = None;
    let mut depth = 0;

    loop {
        match reader.read_event()? {
            Event::Start(e) => {
                let local = local_name(e.name().as_ref());
                match local.as_str() {
                    "pPr" if depth == 0 => {
                        let (s, a, sid, npr, spr) = parse_paragraph_properties(reader)?;
                        style = s;
                        alignment = a;
                        style_id = sid;
                        num_pr_ref = npr;
                        para_sect_pr = spr;
                    }
                    "r" if depth == 0 => {
                        let (run, img) = parse_run(reader, ctx, None)?;
                        runs.push(run);
                        if let Some(image) = img {
                            images.push(image);
                        }
                    }
                    "hyperlink" if depth == 0 => {
                        // w:hyperlink r:id="rIdN" or w:anchor="bookmarkName"
                        let mut link_url: Option<String> = None;
                        for attr in e.attributes().flatten() {
                            let key = std::str::from_utf8(attr.key.as_ref()).unwrap_or("");
                            let val = String::from_utf8_lossy(&attr.value).to_string();
                            if key == "r:id" || key.ends_with(":id") {
                                if let Some(url) = ctx.hyperlinks.get(&val) {
                                    link_url = Some(url.clone());
                                }
                            } else if key == "w:anchor" || key == "anchor" {
                                link_url = Some(format!("#{}", val));
                            }
                        }
                        let hyperlink_runs = parse_hyperlink_runs(reader, ctx, link_url)?;
                        runs.extend(hyperlink_runs);
                    }
                    // Track changes: inserted content
                    "ins" if depth == 0 => {
                        let mut author = None;
                        let mut date = None;
                        for attr in e.attributes().flatten() {
                            let key = local_name(attr.key.as_ref());
                            let val = String::from_utf8_lossy(&attr.value).to_string();
                            match key.as_str() {
                                "author" => author = Some(val),
                                "date" => date = Some(val),
                                _ => {}
                            }
                        }
                        let tc = TrackedChange { change_type: "insert".into(), author, date };
                        let tracked_runs = parse_tracked_change_runs(reader, ctx, "ins", tc)?;
                        runs.extend(tracked_runs);
                    }
                    // Track changes: deleted content
                    "del" if depth == 0 => {
                        let mut author = None;
                        let mut date = None;
                        for attr in e.attributes().flatten() {
                            let key = local_name(attr.key.as_ref());
                            let val = String::from_utf8_lossy(&attr.value).to_string();
                            match key.as_str() {
                                "author" => author = Some(val),
                                "date" => date = Some(val),
                                _ => {}
                            }
                        }
                        let tc = TrackedChange { change_type: "delete".into(), author, date };
                        let tracked_runs = parse_tracked_change_runs(reader, ctx, "del", tc)?;
                        runs.extend(tracked_runs);
                    }
                    _ => {
                        depth += 1;
                    }
                }
            }
            Event::Empty(e) => {
                if depth == 0 {
                    let local = local_name(e.name().as_ref());
                    match local.as_str() {
                        "commentRangeStart" => {
                            for attr in e.attributes().flatten() {
                                let key = local_name(attr.key.as_ref());
                                if key == "id" {
                                    let id = String::from_utf8_lossy(&attr.value).to_string();
                                    // Mark the next run as having a comment start
                                    if let Some(last_run) = runs.last_mut() {
                                        last_run.comment_range_start.push(id);
                                    }
                                }
                            }
                        }
                        "commentRangeEnd" => {
                            for attr in e.attributes().flatten() {
                                let key = local_name(attr.key.as_ref());
                                if key == "id" {
                                    let id = String::from_utf8_lossy(&attr.value).to_string();
                                    if let Some(last_run) = runs.last_mut() {
                                        last_run.comment_range_end.push(id);
                                    }
                                }
                            }
                        }
                        "bookmarkStart" | "bookmarkEnd" => {
                            // Parsed but not acted on yet — bookmarks tracked for future use
                        }
                        _ => {}
                    }
                }
            }
            Event::End(e) => {
                let local = local_name(e.name().as_ref());
                if local == "p" && depth == 0 {
                    break;
                }
                if depth > 0 {
                    depth -= 1;
                }
            }
            Event::Eof => break,
            _ => {}
        }
    }

    // Apply style inheritance from StyleSheet (basedOn already resolved)
    if let Some(ref sid) = style_id {
        if let Some(defined) = styles.styles.get(sid) {
            let ds = &defined.paragraph;
            if style.space_before.is_none() {
                style.space_before = ds.space_before;
            }
            if style.space_after.is_none() {
                style.space_after = ds.space_after;
            }
            if style.line_spacing.is_none() {
                style.line_spacing = ds.line_spacing;
            }
            if style.default_run_style.is_none() {
                style.default_run_style = ds.default_run_style.clone();
            }
            // Inherit keepNext, keepLines from style
            if ds.keep_next { style.keep_next = true; }
            if ds.keep_lines { style.keep_lines = true; }
        }
    }

    // Apply docDefaults fallback
    if style.default_run_style.is_none() {
        style.default_run_style = styles.doc_default_run_style.clone();
    }
    if let Some(ref doc_para) = styles.doc_default_para_style {
        if style.space_after.is_none() {
            style.space_after = doc_para.space_after;
        }
        if style.line_spacing.is_none() {
            style.line_spacing = doc_para.line_spacing;
        }
    }

    // Resolve list marker from numbering definitions
    if let Some(npr) = num_pr_ref {
        if !npr.num_id.is_empty() && npr.num_id != "0" {
            let (marker, indent) = ctx.numbering.resolve_marker(
                &npr.num_id,
                npr.ilvl,
                &mut ctx.list_counters.borrow_mut(),
            );
            style.list_marker = Some(marker);
            if let Some(ind) = indent {
                style.list_indent = Some(ind);
                // Set indent_left from numbering definition if not already set
                if style.indent_left.is_none() {
                    if let Some(left) = ctx.numbering.get_level_indent(&npr.num_id, npr.ilvl) {
                        style.indent_left = Some(left);
                    }
                }
            }
        }
    }

    // Store style ID for contextual spacing comparison
    style.style_id = style_id;

    // Append images as separate blocks after the paragraph runs
    // For now, store the first image inline with the paragraph
    let _ = images; // TODO: Better image-in-paragraph representation

    Ok((Paragraph {
        runs,
        style,
        alignment,
    }, para_sect_pr))
}

/// Numbering reference parsed from w:numPr
struct NumPrRef {
    num_id: String,
    ilvl: u8,
}

/// Parse w:pPr (paragraph properties).
/// Returns: (style, alignment, style_id, numPr, optional section properties for section break)
fn parse_paragraph_properties(
    reader: &mut Reader<&[u8]>,
) -> Result<(ParagraphStyle, Alignment, Option<String>, Option<NumPrRef>, Option<SectionProperties>), ParseError> {
    let mut style = ParagraphStyle::default();
    let mut alignment = Alignment::default();
    let mut style_id: Option<String> = None;
    let mut num_pr: Option<NumPrRef> = None;
    let mut sect_pr: Option<SectionProperties> = None;
    let mut depth = 0;

    loop {
        match reader.read_event()? {
            Event::Start(e) => {
                let local = local_name(e.name().as_ref());
                match local.as_str() {
                    "numPr" if depth == 0 => {
                        num_pr = Some(parse_num_pr(reader)?);
                    }
                    "spacing" if depth == 0 => {
                        for attr in e.attributes().flatten() {
                            let key = local_name(attr.key.as_ref());
                            let val = String::from_utf8_lossy(&attr.value);
                            match key.as_str() {
                                "before" => {
                                    style.space_before =
                                        val.parse::<f32>().ok().map(|v| v / 20.0);
                                }
                                "after" => {
                                    style.space_after =
                                        val.parse::<f32>().ok().map(|v| v / 20.0);
                                }
                                "line" => {
                                    style.line_spacing =
                                        val.parse::<f32>().ok().map(|v| v / 240.0);
                                }
                                _ => {}
                            }
                        }
                        depth += 1;
                    }
                    "tabs" if depth == 0 => {
                        style.tab_stops = parse_tab_stops(reader)?;
                    }
                    "pBdr" if depth == 0 => {
                        style.borders = Some(parse_paragraph_borders(reader)?);
                    }
                    "sectPr" if depth == 0 => {
                        sect_pr = Some(parse_section_properties(reader)?);
                    }
                    _ => {
                        depth += 1;
                    }
                }
            }
            Event::Empty(e) => {
                let local = local_name(e.name().as_ref());
                match local.as_str() {
                    "pStyle" => {
                        for attr in e.attributes().flatten() {
                            if local_name(attr.key.as_ref()) == "val" {
                                let val = String::from_utf8_lossy(&attr.value).to_string();
                                if val.starts_with("Heading") {
                                    style.heading_level =
                                        val.trim_start_matches("Heading").parse().ok();
                                }
                                style_id = Some(val);
                            }
                        }
                    }
                    "jc" => {
                        for attr in e.attributes().flatten() {
                            if local_name(attr.key.as_ref()) == "val" {
                                let val = String::from_utf8_lossy(&attr.value);
                                alignment = match val.as_ref() {
                                    "left" | "start" => Alignment::Left,
                                    "center" => Alignment::Center,
                                    "right" | "end" => Alignment::Right,
                                    "both" | "distribute" => Alignment::Justify,
                                    _ => Alignment::Left,
                                };
                            }
                        }
                    }
                    "snapToGrid" => {
                        for attr in e.attributes().flatten() {
                            if local_name(attr.key.as_ref()) == "val" {
                                let val = String::from_utf8_lossy(&attr.value);
                                style.snap_to_grid = val.as_ref() != "0" && val.as_ref() != "false";
                            }
                        }
                    }
                    "contextualSpacing" => {
                        // w:contextualSpacing: presence alone means true,
                        // or explicit val="1"/"true"
                        let mut enabled = true;
                        for attr in e.attributes().flatten() {
                            if local_name(attr.key.as_ref()) == "val" {
                                let val = String::from_utf8_lossy(&attr.value);
                                enabled = val.as_ref() != "0" && val.as_ref() != "false";
                            }
                        }
                        style.contextual_spacing = enabled;
                    }
                    "spacing" => {
                        for attr in e.attributes().flatten() {
                            let key = local_name(attr.key.as_ref());
                            let val = String::from_utf8_lossy(&attr.value);
                            match key.as_str() {
                                "before" => {
                                    style.space_before =
                                        val.parse::<f32>().ok().map(|v| v / 20.0);
                                }
                                "after" => {
                                    style.space_after =
                                        val.parse::<f32>().ok().map(|v| v / 20.0);
                                }
                                "line" => {
                                    style.line_spacing =
                                        val.parse::<f32>().ok().map(|v| v / 240.0);
                                }
                                _ => {}
                            }
                        }
                    }
                    "ind" => {
                        for attr in e.attributes().flatten() {
                            let key = local_name(attr.key.as_ref());
                            let val = String::from_utf8_lossy(&attr.value);
                            match key.as_str() {
                                "left" => {
                                    style.indent_left =
                                        val.parse::<f32>().ok().map(|v| v / 20.0);
                                }
                                "right" => {
                                    style.indent_right =
                                        val.parse::<f32>().ok().map(|v| v / 20.0);
                                }
                                "firstLine" => {
                                    style.indent_first_line =
                                        val.parse::<f32>().ok().map(|v| v / 20.0);
                                }
                                "hanging" => {
                                    // Hanging indent: negative first-line indent
                                    style.indent_first_line =
                                        val.parse::<f32>().ok().map(|v| -(v / 20.0));
                                }
                                _ => {}
                            }
                        }
                    }
                    "shd" => {
                        for attr in e.attributes().flatten() {
                            if local_name(attr.key.as_ref()) == "fill" {
                                let val = String::from_utf8_lossy(&attr.value).to_string();
                                if val != "auto" {
                                    style.shading = Some(val);
                                }
                            }
                        }
                    }
                    "pageBreakBefore" => {
                        style.page_break_before = true;
                    }
                    "keepNext" => {
                        style.keep_next = true;
                    }
                    "keepLines" => {
                        style.keep_lines = true;
                    }
                    "widowControl" => {
                        let mut enabled = true;
                        for attr in e.attributes().flatten() {
                            if local_name(attr.key.as_ref()) == "val" {
                                let val = String::from_utf8_lossy(&attr.value);
                                enabled = val.as_ref() != "0" && val.as_ref() != "false";
                            }
                        }
                        style.widow_control = enabled;
                    }
                    _ => {}
                }
            }
            Event::End(e) => {
                let local = local_name(e.name().as_ref());
                if local == "pPr" && depth == 0 {
                    break;
                }
                if depth > 0 {
                    depth -= 1;
                }
            }
            Event::Eof => break,
            _ => {}
        }
    }

    Ok((style, alignment, style_id, num_pr, sect_pr))
}

/// Parse w:numPr element
fn parse_num_pr(reader: &mut Reader<&[u8]>) -> Result<NumPrRef, ParseError> {
    let mut num_id = String::new();
    let mut ilvl: u8 = 0;
    let mut depth = 0;

    loop {
        match reader.read_event()? {
            Event::Start(_) => {
                depth += 1;
            }
            Event::Empty(e) => {
                let local = local_name(e.name().as_ref());
                match local.as_str() {
                    "ilvl" => {
                        for attr in e.attributes().flatten() {
                            if local_name(attr.key.as_ref()) == "val" {
                                let val = String::from_utf8_lossy(&attr.value);
                                ilvl = val.parse().unwrap_or(0);
                            }
                        }
                    }
                    "numId" => {
                        for attr in e.attributes().flatten() {
                            if local_name(attr.key.as_ref()) == "val" {
                                num_id = String::from_utf8_lossy(&attr.value).to_string();
                            }
                        }
                    }
                    _ => {}
                }
            }
            Event::End(e) => {
                let local = local_name(e.name().as_ref());
                if local == "numPr" && depth == 0 {
                    break;
                }
                if depth > 0 {
                    depth -= 1;
                }
            }
            Event::Eof => break,
            _ => {}
        }
    }

    Ok(NumPrRef { num_id, ilvl })
}

/// Parse w:pBdr element containing border children (top, bottom, left, right, between)
fn parse_paragraph_borders(reader: &mut Reader<&[u8]>) -> Result<ParagraphBorders, ParseError> {
    let mut borders = ParagraphBorders {
        top: None, bottom: None, left: None, right: None, between: None,
    };

    loop {
        match reader.read_event()? {
            Event::Empty(e) => {
                let local = local_name(e.name().as_ref());
                let bdr = parse_border_attrs(&e);
                match local.as_str() {
                    "top" => borders.top = bdr,
                    "bottom" => borders.bottom = bdr,
                    "left" | "start" => borders.left = bdr,
                    "right" | "end" => borders.right = bdr,
                    "between" => borders.between = bdr,
                    _ => {}
                }
            }
            Event::End(e) => {
                if local_name(e.name().as_ref()) == "pBdr" {
                    break;
                }
            }
            Event::Eof => break,
            _ => {}
        }
    }
    Ok(borders)
}

/// Parse border attributes from an element (w:sz, w:color, w:val)
fn parse_border_attrs(e: &quick_xml::events::BytesStart) -> Option<BorderDef> {
    let mut style = String::new();
    let mut width: f32 = 0.0;
    let mut color = None;

    for attr in e.attributes().flatten() {
        let key = local_name(attr.key.as_ref());
        let val = String::from_utf8_lossy(&attr.value).to_string();
        match key.as_str() {
            "val" => {
                if val == "none" || val == "nil" {
                    return None;
                }
                style = val;
            }
            "sz" => {
                width = val.parse::<f32>().unwrap_or(0.0) / 8.0;
            }
            "color" => {
                if val != "auto" {
                    color = Some(val);
                }
            }
            _ => {}
        }
    }

    if style.is_empty() {
        return None;
    }

    Some(BorderDef { style, width, color })
}

/// Parse w:tabs element containing w:tab children
fn parse_tab_stops(reader: &mut Reader<&[u8]>) -> Result<Vec<TabStop>, ParseError> {
    let mut stops = Vec::new();

    loop {
        match reader.read_event()? {
            Event::Empty(e) => {
                let local = local_name(e.name().as_ref());
                if local == "tab" {
                    let mut position: f32 = 0.0;
                    let mut alignment = TabStopAlignment::Left;
                    let mut leader: Option<String> = None;

                    for attr in e.attributes().flatten() {
                        let key = local_name(attr.key.as_ref());
                        let val = String::from_utf8_lossy(&attr.value);
                        match key.as_str() {
                            "pos" => {
                                // Position in twips (1/20 pt)
                                position = val.parse::<f32>().unwrap_or(0.0) / 20.0;
                            }
                            "val" => {
                                alignment = match val.as_ref() {
                                    "center" => TabStopAlignment::Center,
                                    "right" | "end" => TabStopAlignment::Right,
                                    "decimal" => TabStopAlignment::Decimal,
                                    _ => TabStopAlignment::Left,
                                };
                            }
                            "leader" => {
                                leader = match val.as_ref() {
                                    "none" => None,
                                    _ => Some(val.to_string()),
                                };
                            }
                            _ => {}
                        }
                    }

                    stops.push(TabStop {
                        position,
                        alignment,
                        leader,
                    });
                }
            }
            Event::End(e) => {
                let local = local_name(e.name().as_ref());
                if local == "tabs" {
                    break;
                }
            }
            Event::Eof => break,
            _ => {}
        }
    }

    // Sort by position
    stops.sort_by(|a, b| a.position.partial_cmp(&b.position).unwrap_or(std::cmp::Ordering::Equal));
    Ok(stops)
}

/// Parse a w:r element (run). Returns the Run, optionally an Image, and field info.
/// `url` is set when this run is inside a w:hyperlink element.
fn parse_run(reader: &mut Reader<&[u8]>, ctx: &ParseContext, url: Option<String>) -> Result<(Run, Option<Image>), ParseError> {
    let mut text = String::new();
    let mut style = RunStyle::default();
    let mut image = None;
    let mut depth = 0;
    let mut in_text = false;
    let mut in_instr_text = false;
    let mut instr_text = String::new();
    let mut footnote_ref: Option<u32> = None;
    let mut endnote_ref: Option<u32> = None;
    let mut ruby: Option<Ruby> = None;

    loop {
        match reader.read_event()? {
            Event::Start(e) => {
                let local = local_name(e.name().as_ref());
                match local.as_str() {
                    "rPr" if depth == 0 => {
                        style = parse_run_properties(reader, ctx)?;
                    }
                    "t" if depth == 0 => {
                        in_text = true;
                    }
                    "instrText" if depth == 0 => {
                        in_instr_text = true;
                    }
                    "drawing" if depth == 0 => {
                        image = parse_drawing(reader, ctx)?;
                    }
                    "ruby" if depth == 0 => {
                        ruby = Some(parse_ruby(reader)?);
                    }
                    _ => {
                        depth += 1;
                    }
                }
            }
            Event::Text(e) => {
                let content = e.unescape().unwrap_or_default();
                if in_text {
                    text.push_str(&content);
                } else if in_instr_text {
                    instr_text.push_str(&content);
                }
            }
            Event::End(e) => {
                let local = local_name(e.name().as_ref());
                if local == "t" {
                    in_text = false;
                } else if local == "instrText" {
                    in_instr_text = false;
                } else if local == "r" && depth == 0 {
                    break;
                } else if depth > 0 {
                    depth -= 1;
                }
            }
            Event::Empty(e) => {
                let local = local_name(e.name().as_ref());
                match local.as_str() {
                    "br" => {
                        let mut br_type = None;
                        for attr in e.attributes().flatten() {
                            if local_name(attr.key.as_ref()) == "type" {
                                br_type = Some(String::from_utf8_lossy(&attr.value).to_string());
                            }
                        }
                        match br_type.as_deref() {
                            Some("page") => text.push('\x0C'),   // form feed = page break
                            Some("column") => text.push('\x0B'), // vertical tab = column break
                            _ => text.push('\n'),                 // text wrap break
                        }
                    }
                    "tab" => text.push('\t'),
                    "fldChar" => {
                        // fldChar with fldCharType="separate" or "end" — no action needed
                    }
                    "footnoteReference" => {
                        for attr in e.attributes().flatten() {
                            let key = local_name(attr.key.as_ref());
                            if key == "id" {
                                let val = String::from_utf8_lossy(&attr.value);
                                if let Ok(id) = val.parse::<u32>() {
                                    if id > 0 { // Skip separator/continuation notes (id=0)
                                        footnote_ref = Some(id);
                                        text = format!("[{}]", id);
                                    }
                                }
                            }
                        }
                    }
                    "endnoteReference" => {
                        for attr in e.attributes().flatten() {
                            let key = local_name(attr.key.as_ref());
                            if key == "id" {
                                let val = String::from_utf8_lossy(&attr.value);
                                if let Ok(id) = val.parse::<u32>() {
                                    if id > 0 {
                                        endnote_ref = Some(id);
                                        text = format!("[{}]", id);
                                    }
                                }
                            }
                        }
                    }
                    _ => {}
                }
            }
            Event::Eof => break,
            _ => {}
        }
    }

    // If this run contains a field instruction, convert to placeholder text
    if !instr_text.is_empty() {
        let field = instr_text.trim();
        if field.contains("PAGE") && !field.contains("NUMPAGES") {
            text = "#".to_string();
        } else if field.contains("NUMPAGES") || field.contains("SECTIONPAGES") {
            text = "#".to_string();
        } else if field.contains("DATE") || field.contains("TIME") {
            text = field.to_string();
        }
    }

    // If ruby was parsed, use its base text as the run text
    if let Some(ref r) = ruby {
        if text.is_empty() {
            text = r.base.clone();
        }
    }

    Ok((Run {
        text, style, url, footnote_ref, endnote_ref,
        comment_range_start: Vec::new(),
        comment_range_end: Vec::new(),
        tracked_change: None,
        ruby,
    }, image))
}

/// Parse runs inside a w:hyperlink element
fn parse_hyperlink_runs(reader: &mut Reader<&[u8]>, ctx: &ParseContext, url: Option<String>) -> Result<Vec<Run>, ParseError> {
    let mut runs = Vec::new();
    let mut depth = 0;

    loop {
        match reader.read_event()? {
            Event::Start(e) => {
                let local = local_name(e.name().as_ref());
                if local == "r" && depth == 0 {
                    let (run, _img) = parse_run(reader, ctx, url.clone())?;
                    runs.push(run);
                } else {
                    depth += 1;
                }
            }
            Event::End(e) => {
                let local = local_name(e.name().as_ref());
                if local == "hyperlink" && depth == 0 {
                    break;
                }
                if depth > 0 {
                    depth -= 1;
                }
            }
            Event::Eof => break,
            _ => {}
        }
    }

    Ok(runs)
}

/// Parse word/footnotes.xml or word/endnotes.xml into a map of id -> blocks
fn parse_notes_xml(xml: &str) -> Result<HashMap<String, Vec<Block>>, ParseError> {
    let mut reader = Reader::from_str(xml);
    let mut notes: HashMap<String, Vec<Block>> = HashMap::new();
    let mut current_id: Option<String> = None;
    let mut current_blocks: Vec<Block> = Vec::new();
    let mut depth = 0;
    let mut in_note = false;

    // Create a minimal context for parsing (no media/hyperlinks in notes)
    let note_ctx = ParseContext {
        _rels: HashMap::new(),
        media: HashMap::new(),
        media_types: HashMap::new(),
        hyperlinks: HashMap::new(),
        numbering: super::numbering::NumberingDefinitions::default(),
        list_counters: std::cell::RefCell::new(HashMap::new()),
        footnotes: HashMap::new(),
        endnotes: HashMap::new(),
        comments: HashMap::new(),
        theme: ThemeColors::default(),
    };
    let empty_styles = StyleSheet::default();

    loop {
        match reader.read_event()? {
            Event::Start(e) => {
                let local = local_name(e.name().as_ref());
                match local.as_str() {
                    "footnote" | "endnote" if !in_note => {
                        in_note = true;
                        depth = 0;
                        current_blocks.clear();
                        current_id = None;
                        for attr in e.attributes().flatten() {
                            let key = local_name(attr.key.as_ref());
                            if key == "id" {
                                let val = String::from_utf8_lossy(&attr.value).to_string();
                                current_id = Some(val);
                            }
                        }
                    }
                    "p" if in_note && depth == 0 => {
                        let (para, _) = parse_paragraph(&mut reader, &note_ctx, &empty_styles)?;
                        current_blocks.push(Block::Paragraph(para));
                    }
                    _ if in_note => {
                        depth += 1;
                    }
                    _ => {}
                }
            }
            Event::End(e) => {
                let local = local_name(e.name().as_ref());
                match local.as_str() {
                    "footnote" | "endnote" if in_note && depth == 0 => {
                        if let Some(id) = current_id.take() {
                            // Skip separator notes (id 0 and -1)
                            if id != "0" && id != "-1" {
                                notes.insert(id, std::mem::take(&mut current_blocks));
                            }
                        }
                        in_note = false;
                    }
                    _ if in_note && depth > 0 => {
                        depth -= 1;
                    }
                    _ => {}
                }
            }
            Event::Eof => break,
            _ => {}
        }
    }

    Ok(notes)
}

/// Parse a w:drawing element to extract image info
fn parse_drawing(reader: &mut Reader<&[u8]>, ctx: &ParseContext) -> Result<Option<Image>, ParseError> {
    let mut width: f32 = 0.0;
    let mut height: f32 = 0.0;
    let mut alt_text = None;
    let mut rel_id = None;
    let mut depth = 0;
    // Floating image info
    let mut is_anchor = false;
    let mut pos_x: f32 = 0.0;
    let mut pos_y: f32 = 0.0;
    let mut h_relative: Option<String> = None;
    let mut v_relative: Option<String> = None;
    let mut wrap_type: Option<WrapType> = None;
    let mut in_pos_h = false;
    let mut in_pos_v = false;
    let mut in_pos_offset = false;

    loop {
        match reader.read_event()? {
            Event::Start(e) => {
                let local = local_name(e.name().as_ref());
                depth += 1;
                match local.as_str() {
                    "anchor" => {
                        is_anchor = true;
                    }
                    "docPr" => {
                        for attr in e.attributes().flatten() {
                            let key = std::str::from_utf8(attr.key.as_ref()).unwrap_or("");
                            if key == "descr" {
                                alt_text = Some(String::from_utf8_lossy(&attr.value).to_string());
                            }
                        }
                    }
                    "positionH" => {
                        in_pos_h = true;
                        for attr in e.attributes().flatten() {
                            let key = local_name(attr.key.as_ref());
                            if key == "relativeFrom" {
                                h_relative = Some(String::from_utf8_lossy(&attr.value).to_string());
                            }
                        }
                    }
                    "positionV" => {
                        in_pos_v = true;
                        for attr in e.attributes().flatten() {
                            let key = local_name(attr.key.as_ref());
                            if key == "relativeFrom" {
                                v_relative = Some(String::from_utf8_lossy(&attr.value).to_string());
                            }
                        }
                    }
                    "posOffset" => {
                        in_pos_offset = true;
                    }
                    _ => {}
                }
            }
            Event::Text(e) => {
                if in_pos_offset {
                    let content = e.unescape().unwrap_or_default();
                    if let Ok(v) = content.parse::<f32>() {
                        let pt = v / 12700.0; // EMU to points
                        if in_pos_h { pos_x = pt; }
                        else if in_pos_v { pos_y = pt; }
                    }
                }
            }
            Event::Empty(e) => {
                let local = local_name(e.name().as_ref());
                match local.as_str() {
                    "wrapNone" => { wrap_type = Some(WrapType::None); }
                    "wrapSquare" => { wrap_type = Some(WrapType::Square); }
                    "wrapTight" => { wrap_type = Some(WrapType::Tight); }
                    "wrapTopAndBottom" => { wrap_type = Some(WrapType::TopAndBottom); }
                    "extent" => {
                        // wp:extent cx/cy are in EMUs (English Metric Units)
                        // 1 inch = 914400 EMUs, 1 point = 12700 EMUs
                        for attr in e.attributes().flatten() {
                            let key = std::str::from_utf8(attr.key.as_ref()).unwrap_or("");
                            let val = String::from_utf8_lossy(&attr.value);
                            match key {
                                "cx" => {
                                    if let Ok(v) = val.parse::<f32>() {
                                        width = v / 12700.0; // EMU to points
                                    }
                                }
                                "cy" => {
                                    if let Ok(v) = val.parse::<f32>() {
                                        height = v / 12700.0;
                                    }
                                }
                                _ => {}
                            }
                        }
                    }
                    "blip" => {
                        // a:blip r:embed="rId1"
                        for attr in e.attributes().flatten() {
                            let key = std::str::from_utf8(attr.key.as_ref()).unwrap_or("");
                            if key == "r:embed" || key.ends_with(":embed") || key == "embed" {
                                rel_id = Some(String::from_utf8_lossy(&attr.value).to_string());
                            }
                        }
                    }
                    "ext" => {
                        // a:ext cx/cy fallback for size
                        for attr in e.attributes().flatten() {
                            let key = std::str::from_utf8(attr.key.as_ref()).unwrap_or("");
                            let val = String::from_utf8_lossy(&attr.value);
                            match key {
                                "cx" if width == 0.0 => {
                                    if let Ok(v) = val.parse::<f32>() {
                                        width = v / 12700.0;
                                    }
                                }
                                "cy" if height == 0.0 => {
                                    if let Ok(v) = val.parse::<f32>() {
                                        height = v / 12700.0;
                                    }
                                }
                                _ => {}
                            }
                        }
                    }
                    "docPr" => {
                        for attr in e.attributes().flatten() {
                            let key = std::str::from_utf8(attr.key.as_ref()).unwrap_or("");
                            if key == "descr" {
                                alt_text = Some(String::from_utf8_lossy(&attr.value).to_string());
                            }
                        }
                    }
                    _ => {}
                }
            }
            Event::End(e) => {
                let local = local_name(e.name().as_ref());
                match local.as_str() {
                    "positionH" => { in_pos_h = false; }
                    "positionV" => { in_pos_v = false; }
                    "posOffset" => { in_pos_offset = false; }
                    _ => {}
                }
                if local == "drawing" && depth == 0 {
                    break;
                }
                if depth > 0 {
                    depth -= 1;
                }
            }
            Event::Eof => break,
            _ => {}
        }
    }

    // Resolve the image data from the relationship
    if let Some(rid) = rel_id {
        let data = ctx.media.get(&rid).cloned().unwrap_or_default();
        let content_type = ctx.media_types.get(&rid).cloned();
        let position = if is_anchor {
            Some(FloatingPosition { x: pos_x, y: pos_y, h_relative, v_relative })
        } else {
            None
        };
        Ok(Some(Image {
            data,
            width,
            height,
            alt_text,
            content_type,
            position,
            wrap_type,
        }))
    } else {
        Ok(None)
    }
}

/// Parse w:rPr (run properties)
fn parse_run_properties(reader: &mut Reader<&[u8]>, ctx: &ParseContext) -> Result<RunStyle, ParseError> {
    let mut style = RunStyle::default();
    let mut depth = 0;

    loop {
        match reader.read_event()? {
            Event::Start(e) => {
                depth += 1;
                let local = local_name(e.name().as_ref());
                if local == "rFonts" {
                    for attr in e.attributes().flatten() {
                        let key = local_name(attr.key.as_ref());
                        if key == "ascii" || key == "hAnsi" {
                            style.font_family =
                                Some(String::from_utf8_lossy(&attr.value).to_string());
                        }
                    }
                }
            }
            Event::Empty(e) => {
                let local = local_name(e.name().as_ref());
                match local.as_str() {
                    "b" => style.bold = true,
                    "i" => style.italic = true,
                    "u" => {
                        style.underline = true;
                        for attr in e.attributes().flatten() {
                            if local_name(attr.key.as_ref()) == "val" {
                                let val = String::from_utf8_lossy(&attr.value).to_string();
                                if val == "none" {
                                    style.underline = false;
                                } else {
                                    style.underline_style = Some(val);
                                }
                            }
                        }
                    }
                    "strike" | "dstrike" => style.strikethrough = true,
                    "highlight" => {
                        for attr in e.attributes().flatten() {
                            if local_name(attr.key.as_ref()) == "val" {
                                style.highlight =
                                    Some(String::from_utf8_lossy(&attr.value).to_string());
                            }
                        }
                    }
                    "vertAlign" => {
                        for attr in e.attributes().flatten() {
                            if local_name(attr.key.as_ref()) == "val" {
                                let val = String::from_utf8_lossy(&attr.value);
                                style.vertical_align = match val.as_ref() {
                                    "superscript" => Some(VerticalAlign::Superscript),
                                    "subscript" => Some(VerticalAlign::Subscript),
                                    _ => Some(VerticalAlign::Baseline),
                                };
                            }
                        }
                    }
                    "sz" => {
                        for attr in e.attributes().flatten() {
                            if local_name(attr.key.as_ref()) == "val" {
                                let val = String::from_utf8_lossy(&attr.value);
                                // OOXML sz is in half-points
                                style.font_size = val.parse::<f32>().ok().map(|v| v / 2.0);
                            }
                        }
                    }
                    "rFonts" => {
                        for attr in e.attributes().flatten() {
                            let key = local_name(attr.key.as_ref());
                            if key == "ascii" || key == "hAnsi" {
                                style.font_family =
                                    Some(String::from_utf8_lossy(&attr.value).to_string());
                            }
                        }
                    }
                    "color" => {
                        let mut color_val = None;
                        let mut theme_color = None;
                        let mut theme_tint = None;
                        let mut theme_shade = None;
                        for attr in e.attributes().flatten() {
                            let key = local_name(attr.key.as_ref());
                            let val = String::from_utf8_lossy(&attr.value).to_string();
                            match key.as_str() {
                                "val" => color_val = Some(val),
                                "themeColor" => theme_color = Some(val),
                                "themeTint" => theme_tint = Some(val),
                                "themeShade" => theme_shade = Some(val),
                                _ => {}
                            }
                        }
                        // Resolve theme color if available
                        if let Some(ref tc) = theme_color {
                            if let Some(resolved) = ctx.theme.resolve(tc) {
                                let mut hex: String = resolved.clone();
                                // Apply tint/shade
                                if let Some(ref tint) = theme_tint {
                                    if let Ok(t) = u8::from_str_radix(tint, 16) {
                                        let tint_val = t as f64 / 255.0;
                                        hex = ThemeColors::apply_tint_shade(&hex, tint_val);
                                    }
                                }
                                if let Some(ref shade) = theme_shade {
                                    if let Ok(s) = u8::from_str_radix(shade, 16) {
                                        let shade_val = -(1.0 - s as f64 / 255.0);
                                        hex = ThemeColors::apply_tint_shade(&hex, shade_val);
                                    }
                                }
                                style.color = Some(hex);
                            } else if let Some(ref cv) = color_val {
                                if cv != "auto" {
                                    style.color = Some(cv.clone());
                                }
                            }
                        } else if let Some(ref cv) = color_val {
                            if cv != "auto" {
                                style.color = Some(cv.clone());
                            }
                        }
                    }
                    "spacing" => {
                        for attr in e.attributes().flatten() {
                            if local_name(attr.key.as_ref()) == "val" {
                                let val = String::from_utf8_lossy(&attr.value);
                                // w:spacing w:val is in twips (1/20 pt)
                                style.character_spacing =
                                    val.parse::<f32>().ok().map(|v| v / 20.0);
                            }
                        }
                    }
                    "smallCaps" => {
                        style.small_caps = true;
                    }
                    "caps" => {
                        style.all_caps = true;
                    }
                    _ => {}
                }
            }
            Event::End(e) => {
                let local = local_name(e.name().as_ref());
                if local == "rPr" && depth == 0 {
                    break;
                }
                if depth > 0 {
                    depth -= 1;
                }
            }
            Event::Eof => break,
            _ => {}
        }
    }

    Ok(style)
}

/// Parse a w:tbl element (table)
fn parse_table(reader: &mut Reader<&[u8]>, ctx: &ParseContext, styles: &StyleSheet) -> Result<Table, ParseError> {
    let mut rows = Vec::new();
    let mut style = TableStyle::default();
    let mut depth = 0;

    loop {
        match reader.read_event()? {
            Event::Start(e) => {
                let local = local_name(e.name().as_ref());
                match local.as_str() {
                    "tblPr" if depth == 0 => {
                        style = parse_table_properties(reader)?;
                    }
                    "tr" if depth == 0 => {
                        let row = parse_table_row(reader, ctx, styles)?;
                        rows.push(row);
                    }
                    _ => {
                        depth += 1;
                    }
                }
            }
            Event::End(e) => {
                let local = local_name(e.name().as_ref());
                if local == "tbl" && depth == 0 {
                    break;
                }
                if depth > 0 {
                    depth -= 1;
                }
            }
            Event::Eof => break,
            _ => {}
        }
    }

    Ok(Table { rows, style })
}

/// Parse w:tblPr (table properties)
fn parse_table_properties(reader: &mut Reader<&[u8]>) -> Result<TableStyle, ParseError> {
    let mut style = TableStyle::default();
    let mut depth = 0;
    let mut in_borders = false;

    loop {
        match reader.read_event()? {
            Event::Start(e) => {
                let local = local_name(e.name().as_ref());
                if local == "tblBorders" {
                    style.border = true;
                    in_borders = true;
                }
                depth += 1;
            }
            Event::End(e) => {
                let local = local_name(e.name().as_ref());
                if local == "tblBorders" {
                    in_borders = false;
                }
                if local == "tblPr" && depth == 0 {
                    break;
                }
                if depth > 0 {
                    depth -= 1;
                }
            }
            Event::Empty(e) => {
                let local = local_name(e.name().as_ref());
                match local.as_str() {
                    "top" | "left" | "bottom" | "right" | "insideH" | "insideV" | "start" | "end"
                        if in_borders || depth == 0 =>
                    {
                        style.border = true;
                        if style.border_color.is_none() {
                            for attr in e.attributes().flatten() {
                                let key = local_name(attr.key.as_ref());
                                let val = String::from_utf8_lossy(&attr.value);
                                match key.as_str() {
                                    "color" => {
                                        if val != "auto" {
                                            style.border_color = Some(val.to_string());
                                        }
                                    }
                                    "sz" => {
                                        style.border_width = val.parse::<f32>().ok().map(|v| v / 8.0);
                                    }
                                    "val" => {
                                        if val != "none" && val != "nil" {
                                            style.border_style = Some(val.to_string());
                                        }
                                    }
                                    _ => {}
                                }
                            }
                        }
                    }
                    "tblW" => {
                        let mut w_val = None;
                        let mut w_type = None;
                        for attr in e.attributes().flatten() {
                            let key = local_name(attr.key.as_ref());
                            let val = String::from_utf8_lossy(&attr.value).to_string();
                            match key.as_str() {
                                "w" => w_val = Some(val),
                                "type" => w_type = Some(val),
                                _ => {}
                            }
                        }
                        style.width_type = w_type.clone();
                        if let Some(ref wv) = w_val {
                            match w_type.as_deref() {
                                Some("dxa") => {
                                    style.width = wv.parse::<f32>().ok().map(|v| v / 20.0);
                                }
                                Some("pct") => {
                                    // Percentage stored as 50ths of a percent
                                    style.width = wv.parse::<f32>().ok().map(|v| v / 50.0);
                                }
                                _ => {}
                            }
                        }
                    }
                    "jc" => {
                        for attr in e.attributes().flatten() {
                            if local_name(attr.key.as_ref()) == "val" {
                                style.alignment = Some(String::from_utf8_lossy(&attr.value).to_string());
                            }
                        }
                    }
                    "tblStyle" => {
                        for attr in e.attributes().flatten() {
                            if local_name(attr.key.as_ref()) == "val" {
                                style.style_id = Some(String::from_utf8_lossy(&attr.value).to_string());
                            }
                        }
                    }
                    _ => {}
                }
            }
            Event::Eof => break,
            _ => {}
        }
    }

    Ok(style)
}

/// Parse a w:tr element (table row)
fn parse_table_row(reader: &mut Reader<&[u8]>, ctx: &ParseContext, styles: &StyleSheet) -> Result<TableRow, ParseError> {
    let mut cells = Vec::new();
    let mut height: Option<f32> = None;
    let mut depth = 0;

    loop {
        match reader.read_event()? {
            Event::Start(e) => {
                let local = local_name(e.name().as_ref());
                match local.as_str() {
                    "tc" if depth == 0 => {
                        let cell = parse_table_cell(reader, ctx, styles)?;
                        cells.push(cell);
                    }
                    _ => {
                        depth += 1;
                    }
                }
            }
            Event::Empty(e) => {
                let local = local_name(e.name().as_ref());
                if local == "trHeight" {
                    for attr in e.attributes().flatten() {
                        if local_name(attr.key.as_ref()) == "val" {
                            let val = String::from_utf8_lossy(&attr.value);
                            height = val.parse::<f32>().ok().map(|v| v / 20.0);
                        }
                    }
                }
            }
            Event::End(e) => {
                let local = local_name(e.name().as_ref());
                if local == "tr" && depth == 0 {
                    break;
                }
                if depth > 0 {
                    depth -= 1;
                }
            }
            Event::Eof => break,
            _ => {}
        }
    }

    Ok(TableRow { cells, height })
}

/// Parse a w:tc element (table cell)
fn parse_table_cell(reader: &mut Reader<&[u8]>, ctx: &ParseContext, styles: &StyleSheet) -> Result<TableCell, ParseError> {
    let mut blocks = Vec::new();
    let mut cell_props = CellProperties::default();
    let mut depth = 0;

    loop {
        match reader.read_event()? {
            Event::Start(e) => {
                let local = local_name(e.name().as_ref());
                match local.as_str() {
                    "p" if depth == 0 => {
                        let (para, _) = parse_paragraph(reader, ctx, styles)?;
                        blocks.push(Block::Paragraph(para));
                    }
                    "tcPr" if depth == 0 => {
                        cell_props = parse_cell_properties(reader)?;
                    }
                    _ => {
                        depth += 1;
                    }
                }
            }
            Event::End(e) => {
                let local = local_name(e.name().as_ref());
                if local == "tc" && depth == 0 {
                    break;
                }
                if depth > 0 {
                    depth -= 1;
                }
            }
            Event::Eof => break,
            _ => {}
        }
    }

    Ok(TableCell {
        blocks,
        width: cell_props.width,
        grid_span: cell_props.grid_span,
        v_merge: cell_props.v_merge,
        shading: cell_props.shading,
        v_align: cell_props.v_align,
        borders: cell_props.borders,
        margins: cell_props.margins,
    })
}

#[derive(Default)]
struct CellProperties {
    width: Option<f32>,
    grid_span: u32,
    v_merge: Option<String>,
    shading: Option<String>,
    v_align: Option<String>,
    borders: Option<CellBorders>,
    margins: Option<CellMargins>,
}

/// Parse w:tcPr (table cell properties)
fn parse_cell_properties(reader: &mut Reader<&[u8]>) -> Result<CellProperties, ParseError> {
    let mut props = CellProperties { grid_span: 1, ..Default::default() };
    let mut depth = 0;

    loop {
        match reader.read_event()? {
            Event::Start(e) => {
                let local = local_name(e.name().as_ref());
                match local.as_str() {
                    "vMerge" => {
                        let mut val = "continue".to_string();
                        for attr in e.attributes().flatten() {
                            if local_name(attr.key.as_ref()) == "val" {
                                val = String::from_utf8_lossy(&attr.value).to_string();
                            }
                        }
                        props.v_merge = Some(val);
                        depth += 1;
                    }
                    "tcBorders" if depth == 0 => {
                        props.borders = Some(parse_cell_borders(reader)?);
                    }
                    "tcMar" if depth == 0 => {
                        props.margins = Some(parse_cell_margins(reader)?);
                    }
                    _ => {
                        depth += 1;
                    }
                }
            }
            Event::Empty(e) => {
                let local = local_name(e.name().as_ref());
                match local.as_str() {
                    "tcW" => {
                        for attr in e.attributes().flatten() {
                            if local_name(attr.key.as_ref()) == "w" {
                                let val = String::from_utf8_lossy(&attr.value);
                                props.width = val.parse::<f32>().ok().map(|v| v / 20.0);
                            }
                        }
                    }
                    "gridSpan" => {
                        for attr in e.attributes().flatten() {
                            if local_name(attr.key.as_ref()) == "val" {
                                let val = String::from_utf8_lossy(&attr.value);
                                props.grid_span = val.parse().unwrap_or(1);
                            }
                        }
                    }
                    "vMerge" => {
                        let mut val = "continue".to_string();
                        for attr in e.attributes().flatten() {
                            if local_name(attr.key.as_ref()) == "val" {
                                val = String::from_utf8_lossy(&attr.value).to_string();
                            }
                        }
                        props.v_merge = Some(val);
                    }
                    "shd" => {
                        for attr in e.attributes().flatten() {
                            if local_name(attr.key.as_ref()) == "fill" {
                                let val = String::from_utf8_lossy(&attr.value).to_string();
                                if val != "auto" {
                                    props.shading = Some(val);
                                }
                            }
                        }
                    }
                    "vAlign" => {
                        for attr in e.attributes().flatten() {
                            if local_name(attr.key.as_ref()) == "val" {
                                props.v_align = Some(String::from_utf8_lossy(&attr.value).to_string());
                            }
                        }
                    }
                    _ => {}
                }
            }
            Event::End(e) => {
                let local = local_name(e.name().as_ref());
                if local == "tcPr" && depth == 0 {
                    break;
                }
                if depth > 0 {
                    depth -= 1;
                }
            }
            Event::Eof => break,
            _ => {}
        }
    }

    Ok(props)
}

/// Parse w:tcBorders
fn parse_cell_borders(reader: &mut Reader<&[u8]>) -> Result<CellBorders, ParseError> {
    let mut borders = CellBorders {
        top: None, bottom: None, left: None, right: None,
    };
    loop {
        match reader.read_event()? {
            Event::Empty(e) => {
                let local = local_name(e.name().as_ref());
                let bdr = parse_border_attrs(&e);
                match local.as_str() {
                    "top" => borders.top = bdr,
                    "bottom" => borders.bottom = bdr,
                    "left" | "start" => borders.left = bdr,
                    "right" | "end" => borders.right = bdr,
                    _ => {}
                }
            }
            Event::End(e) => {
                if local_name(e.name().as_ref()) == "tcBorders" { break; }
            }
            Event::Eof => break,
            _ => {}
        }
    }
    Ok(borders)
}

/// Parse w:tcMar
fn parse_cell_margins(reader: &mut Reader<&[u8]>) -> Result<CellMargins, ParseError> {
    let mut margins = CellMargins {
        top: None, bottom: None, left: None, right: None,
    };
    loop {
        match reader.read_event()? {
            Event::Empty(e) => {
                let local = local_name(e.name().as_ref());
                let val = e.attributes().flatten()
                    .find(|a| local_name(a.key.as_ref()) == "w")
                    .and_then(|a| {
                        String::from_utf8_lossy(&a.value).parse::<f32>().ok().map(|v| v / 20.0)
                    });
                match local.as_str() {
                    "top" => margins.top = val,
                    "bottom" => margins.bottom = val,
                    "left" | "start" => margins.left = val,
                    "right" | "end" => margins.right = val,
                    _ => {}
                }
            }
            Event::End(e) => {
                if local_name(e.name().as_ref()) == "tcMar" { break; }
            }
            Event::Eof => break,
            _ => {}
        }
    }
    Ok(margins)
}

/// Parsed section properties
struct SectionProperties {
    page_size: PageSize,
    margin: Margin,
    /// Document grid line pitch in points (from w:docGrid w:linePitch, twips/20)
    grid_line_pitch: Option<f32>,
    /// Reference IDs for header parts
    header_refs: Vec<String>,
    /// Reference IDs for footer parts
    footer_refs: Vec<String>,
    /// Column layout
    columns: Option<ColumnLayout>,
}

/// Parse w:sectPr (section properties - page size, margins, document grid)
fn parse_section_properties(
    reader: &mut Reader<&[u8]>,
) -> Result<SectionProperties, ParseError> {
    let mut page_size = PageSize::default();
    let mut margin = Margin::default();
    let mut grid_line_pitch: Option<f32> = None;
    let mut header_refs = Vec::new();
    let mut footer_refs = Vec::new();
    let mut columns: Option<ColumnLayout> = None;
    let mut depth = 0;

    loop {
        match reader.read_event()? {
            Event::Start(_) => {
                depth += 1;
            }
            Event::Empty(e) => {
                let local = local_name(e.name().as_ref());
                match local.as_str() {
                    "pgSz" => {
                        let mut orient = None;
                        for attr in e.attributes().flatten() {
                            let key = local_name(attr.key.as_ref());
                            let val = String::from_utf8_lossy(&attr.value);
                            match key.as_str() {
                                "w" => {
                                    if let Ok(v) = val.parse::<f32>() {
                                        page_size.width = v / 20.0;
                                    }
                                }
                                "h" => {
                                    if let Ok(v) = val.parse::<f32>() {
                                        page_size.height = v / 20.0;
                                    }
                                }
                                "orient" => orient = Some(val.to_string()),
                                _ => {}
                            }
                        }
                        // Landscape: ensure width > height
                        if orient.as_deref() == Some("landscape") && page_size.width < page_size.height {
                            std::mem::swap(&mut page_size.width, &mut page_size.height);
                        }
                    }
                    "pgMar" => {
                        for attr in e.attributes().flatten() {
                            let key = local_name(attr.key.as_ref());
                            let val = String::from_utf8_lossy(&attr.value);
                            match key.as_str() {
                                "top" => {
                                    if let Ok(v) = val.parse::<f32>() {
                                        margin.top = v / 20.0;
                                    }
                                }
                                "bottom" => {
                                    if let Ok(v) = val.parse::<f32>() {
                                        margin.bottom = v / 20.0;
                                    }
                                }
                                "left" => {
                                    if let Ok(v) = val.parse::<f32>() {
                                        margin.left = v / 20.0;
                                    }
                                }
                                "right" => {
                                    if let Ok(v) = val.parse::<f32>() {
                                        margin.right = v / 20.0;
                                    }
                                }
                                _ => {}
                            }
                        }
                    }
                    "docGrid" => {
                        let mut grid_type = String::new();
                        let mut line_pitch = 0u32;
                        for attr in e.attributes().flatten() {
                            let key = local_name(attr.key.as_ref());
                            let val = String::from_utf8_lossy(&attr.value);
                            match key.as_str() {
                                "type" => grid_type = val.to_string(),
                                "linePitch" => {
                                    line_pitch = val.parse().unwrap_or(0);
                                }
                                _ => {}
                            }
                        }
                        // Only apply grid for "lines" or "linesAndChars" types
                        if (grid_type == "lines" || grid_type == "linesAndChars")
                            && line_pitch > 0
                        {
                            grid_line_pitch = Some(line_pitch as f32 / 20.0);
                        }
                    }
                    "headerReference" => {
                        for attr in e.attributes().flatten() {
                            let key = local_name(attr.key.as_ref());
                            if key == "id" {
                                header_refs.push(String::from_utf8_lossy(&attr.value).to_string());
                            }
                        }
                    }
                    "footerReference" => {
                        for attr in e.attributes().flatten() {
                            let key = local_name(attr.key.as_ref());
                            if key == "id" {
                                footer_refs.push(String::from_utf8_lossy(&attr.value).to_string());
                            }
                        }
                    }
                    "cols" => {
                        let mut num = 1u32;
                        let mut space: Option<f32> = None;
                        let mut equal_width = true;
                        for attr in e.attributes().flatten() {
                            let key = local_name(attr.key.as_ref());
                            let val = String::from_utf8_lossy(&attr.value);
                            match key.as_str() {
                                "num" => { num = val.parse().unwrap_or(1); }
                                "space" => { space = val.parse::<f32>().ok().map(|v| v / 20.0); }
                                "equalWidth" => { equal_width = val.as_ref() != "0" && val.as_ref() != "false"; }
                                _ => {}
                            }
                        }
                        if num > 1 {
                            columns = Some(ColumnLayout { num, space, equal_width });
                        }
                    }
                    _ => {}
                }
            }
            Event::End(e) => {
                let local = local_name(e.name().as_ref());
                if local == "sectPr" && depth == 0 {
                    break;
                }
                if depth > 0 {
                    depth -= 1;
                }
            }
            Event::Eof => break,
            _ => {}
        }
    }

    Ok(SectionProperties {
        page_size,
        margin,
        grid_line_pitch,
        header_refs,
        footer_refs,
        columns,
    })
}

/// Parse a header or footer XML part (w:hdr or w:ftr element)
fn parse_header_footer_xml(xml: &str, ctx: &ParseContext, styles: &StyleSheet) -> Result<Vec<Block>, ParseError> {
    let mut reader = Reader::from_str(xml);
    let mut blocks = Vec::new();
    let mut depth = 0;
    let mut in_root = false;

    loop {
        match reader.read_event()? {
            Event::Start(e) => {
                let local = local_name(e.name().as_ref());
                match local.as_str() {
                    "hdr" | "ftr" => {
                        in_root = true;
                        depth = 0;
                    }
                    "p" if in_root && depth == 0 => {
                        let (para, _) = parse_paragraph(&mut reader, ctx, styles)?;
                        blocks.push(Block::Paragraph(para));
                    }
                    "tbl" if in_root && depth == 0 => {
                        let table = parse_table(&mut reader, ctx, styles)?;
                        blocks.push(Block::Table(table));
                    }
                    _ if in_root => {
                        depth += 1;
                    }
                    _ => {}
                }
            }
            Event::End(e) => {
                let local = local_name(e.name().as_ref());
                if local == "hdr" || local == "ftr" {
                    in_root = false;
                } else if in_root && depth > 0 {
                    depth -= 1;
                }
            }
            Event::Eof => break,
            _ => {}
        }
    }

    Ok(blocks)
}

/// Recursively collect footnote/endnote references from blocks
fn collect_note_refs(blocks: &[Block], ctx: &ParseContext, footnotes: &mut Vec<Footnote>, endnotes: &mut Vec<Footnote>) {
    for block in blocks {
        match block {
            Block::Paragraph(para) => {
                for run in &para.runs {
                    if let Some(fn_id) = run.footnote_ref {
                        let id_str = fn_id.to_string();
                        if let Some(note_blocks) = ctx.footnotes.get(&id_str) {
                            if !footnotes.iter().any(|f| f.number == fn_id) {
                                footnotes.push(Footnote {
                                    number: fn_id,
                                    blocks: note_blocks.clone(),
                                });
                            }
                        }
                    }
                    if let Some(en_id) = run.endnote_ref {
                        let id_str = en_id.to_string();
                        if let Some(note_blocks) = ctx.endnotes.get(&id_str) {
                            if !endnotes.iter().any(|f| f.number == en_id) {
                                endnotes.push(Footnote {
                                    number: en_id,
                                    blocks: note_blocks.clone(),
                                });
                            }
                        }
                    }
                }
            }
            Block::Table(table) => {
                for row in &table.rows {
                    for cell in &row.cells {
                        collect_note_refs(&cell.blocks, ctx, footnotes, endnotes);
                    }
                }
            }
            Block::Image(_) => {}
        }
    }
}

/// Parse runs inside w:ins or w:del (tracked changes)
fn parse_tracked_change_runs(
    reader: &mut Reader<&[u8]>,
    ctx: &ParseContext,
    end_tag: &str,
    tc: TrackedChange,
) -> Result<Vec<Run>, ParseError> {
    let mut runs = Vec::new();
    let mut depth = 0;

    loop {
        match reader.read_event()? {
            Event::Start(e) => {
                let local = local_name(e.name().as_ref());
                if local == "r" && depth == 0 {
                    let (mut run, _img) = parse_run(reader, ctx, None)?;
                    run.tracked_change = Some(tc.clone());
                    runs.push(run);
                } else {
                    depth += 1;
                }
            }
            Event::End(e) => {
                let local = local_name(e.name().as_ref());
                if local == end_tag && depth == 0 {
                    break;
                }
                if depth > 0 {
                    depth -= 1;
                }
            }
            Event::Eof => break,
            _ => {}
        }
    }

    Ok(runs)
}

/// Parse w:ruby element (furigana)
fn parse_ruby(reader: &mut Reader<&[u8]>) -> Result<Ruby, ParseError> {
    let mut base_text = String::new();
    let mut ruby_text = String::new();
    let mut ruby_font_size: Option<f32> = None;
    let mut depth = 0;
    let mut in_rt = false; // ruby text (annotation)
    let mut in_ruby_base = false; // base text
    let mut in_ruby_pr = false; // ruby properties
    let mut in_t = false;

    loop {
        match reader.read_event()? {
            Event::Start(e) => {
                let local = local_name(e.name().as_ref());
                match local.as_str() {
                    "rt" if depth == 0 => { in_rt = true; }
                    "rubyBase" if depth == 0 => { in_ruby_base = true; }
                    "rubyPr" if depth == 0 => { in_ruby_pr = true; }
                    "t" => { in_t = true; }
                    _ => {}
                }
                depth += 1;
            }
            Event::Text(e) => {
                if in_t {
                    let content = e.unescape().unwrap_or_default();
                    if in_rt {
                        ruby_text.push_str(&content);
                    } else if in_ruby_base {
                        base_text.push_str(&content);
                    }
                }
            }
            Event::Empty(e) => {
                let local = local_name(e.name().as_ref());
                if local == "sz" && in_ruby_pr {
                    for attr in e.attributes().flatten() {
                        if local_name(attr.key.as_ref()) == "val" {
                            let val = String::from_utf8_lossy(&attr.value);
                            ruby_font_size = val.parse::<f32>().ok().map(|v| v / 2.0);
                        }
                    }
                }
            }
            Event::End(e) => {
                let local = local_name(e.name().as_ref());
                match local.as_str() {
                    "ruby" if depth == 1 => {
                        depth -= 1;
                        break;
                    }
                    "rt" => { in_rt = false; }
                    "rubyBase" => { in_ruby_base = false; }
                    "rubyPr" => { in_ruby_pr = false; }
                    "t" => { in_t = false; }
                    _ => {}
                }
                if depth > 0 { depth -= 1; }
            }
            Event::Eof => break,
            _ => {}
        }
    }

    Ok(Ruby {
        base: base_text,
        text: ruby_text,
        font_size: ruby_font_size,
    })
}

/// Parse word/comments.xml
fn parse_comments_xml(xml: &str) -> Result<HashMap<String, Comment>, ParseError> {
    let mut reader = Reader::from_str(xml);
    let mut comments: HashMap<String, Comment> = HashMap::new();
    let mut depth = 0;
    let mut in_comment = false;
    let mut current_id = String::new();
    let mut current_author: Option<String> = None;
    let mut current_date: Option<String> = None;
    let mut current_blocks: Vec<Block> = Vec::new();

    let note_ctx = ParseContext {
        _rels: HashMap::new(),
        media: HashMap::new(),
        media_types: HashMap::new(),
        hyperlinks: HashMap::new(),
        numbering: super::numbering::NumberingDefinitions::default(),
        list_counters: std::cell::RefCell::new(HashMap::new()),
        footnotes: HashMap::new(),
        endnotes: HashMap::new(),
        comments: HashMap::new(),
        theme: ThemeColors::default(),
    };
    let empty_styles = StyleSheet::default();

    loop {
        match reader.read_event()? {
            Event::Start(e) => {
                let local = local_name(e.name().as_ref());
                match local.as_str() {
                    "comment" if !in_comment => {
                        in_comment = true;
                        depth = 0;
                        current_blocks.clear();
                        current_id.clear();
                        current_author = None;
                        current_date = None;
                        for attr in e.attributes().flatten() {
                            let key = local_name(attr.key.as_ref());
                            let val = String::from_utf8_lossy(&attr.value).to_string();
                            match key.as_str() {
                                "id" => current_id = val,
                                "author" => current_author = Some(val),
                                "date" => current_date = Some(val),
                                _ => {}
                            }
                        }
                    }
                    "p" if in_comment && depth == 0 => {
                        let (para, _) = parse_paragraph(&mut reader, &note_ctx, &empty_styles)?;
                        current_blocks.push(Block::Paragraph(para));
                    }
                    _ if in_comment => {
                        depth += 1;
                    }
                    _ => {}
                }
            }
            Event::End(e) => {
                let local = local_name(e.name().as_ref());
                if local == "comment" && in_comment && depth == 0 {
                    if !current_id.is_empty() {
                        comments.insert(current_id.clone(), Comment {
                            id: current_id.clone(),
                            author: current_author.take(),
                            date: current_date.take(),
                            blocks: std::mem::take(&mut current_blocks),
                        });
                    }
                    in_comment = false;
                } else if in_comment && depth > 0 {
                    depth -= 1;
                }
            }
            Event::Eof => break,
            _ => {}
        }
    }

    Ok(comments)
}

/// Extract local name from a potentially namespaced XML tag
fn local_name(name: &[u8]) -> String {
    let s = std::str::from_utf8(name).unwrap_or("");
    match s.rfind(':') {
        Some(pos) => s[pos + 1..].to_string(),
        None => s.to_string(),
    }
}
