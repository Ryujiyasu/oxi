use std::collections::HashMap;
use std::io::{Cursor, Read};

use quick_xml::events::Event;
use quick_xml::reader::Reader;
use zip::ZipArchive;

use super::numbering::{parse_numbering, NumberingDefinitions};
use super::relationships::{parse_relationships, Relationship};
use super::styles::parse_styles;
use super::ParseError;
use crate::ir::{*, VerticalAlign};

pub struct OoxmlParser {
    archive: ZipArchive<Cursor<Vec<u8>>>,
}

/// Context passed through parsing functions for resource resolution
struct ParseContext {
    /// Relationship ID -> Relationship mapping (reserved for future use)
    _rels: HashMap<String, Relationship>,
    /// Relationship ID -> binary data (images, etc.)
    media: HashMap<String, Vec<u8>>,
    /// Relationship ID -> content type (e.g., "image/png")
    media_types: HashMap<String, String>,
    /// Numbering definitions from word/numbering.xml
    numbering: NumberingDefinitions,
    /// Counters for numbered lists: (numId, ilvl) -> current count
    list_counters: std::cell::RefCell<HashMap<(String, u8), u32>>,
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
        let (blocks, sect_pr) = self.parse_document_xml(&ctx, &styles)?;
        let metadata = self.parse_metadata();

        let sect = sect_pr.unwrap_or(SectionProperties {
            page_size: PageSize::default(),
            margin: Margin::default(),
            grid_line_pitch: None,
        });

        Ok(Document {
            pages: vec![Page {
                blocks,
                size: sect.page_size,
                margin: sect.margin,
                grid_line_pitch: sect.grid_line_pitch,
            }],
            styles,
            metadata,
        })
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

        // Pre-load media referenced by image relationships
        let mut media = HashMap::new();
        let mut media_types = HashMap::new();
        let image_rel_type = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image";
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
            }
        }

        Ok(ParseContext {
            _rels: rels,
            media,
            media_types,
            numbering,
            list_counters: std::cell::RefCell::new(HashMap::new()),
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
    ) -> Result<(Vec<Block>, Option<SectionProperties>), ParseError> {
        let xml = self.read_part("word/document.xml")?;
        parse_body(&xml, ctx, styles)
    }
}

/// Parse the w:body content of document.xml
fn parse_body(xml: &str, ctx: &ParseContext, styles: &StyleSheet) -> Result<(Vec<Block>, Option<SectionProperties>), ParseError> {
    let mut reader = Reader::from_str(xml);
    let mut blocks = Vec::new();
    let mut sect_pr = None;
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
                        let para = parse_paragraph(&mut reader, ctx, styles)?;
                        blocks.push(Block::Paragraph(para));
                    }
                    "tbl" if in_body && depth == 0 => {
                        let table = parse_table(&mut reader, ctx, styles)?;
                        blocks.push(Block::Table(table));
                    }
                    "sectPr" if in_body && depth == 0 => {
                        sect_pr = Some(parse_section_properties(&mut reader)?);
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

    Ok((blocks, sect_pr))
}

/// Parse a w:p element (paragraph)
fn parse_paragraph(reader: &mut Reader<&[u8]>, ctx: &ParseContext, styles: &StyleSheet) -> Result<Paragraph, ParseError> {
    let mut runs = Vec::new();
    let mut images = Vec::new();
    let mut style = ParagraphStyle::default();
    let mut alignment = Alignment::default();
    let mut style_id: Option<String> = None;
    let mut num_pr_ref: Option<NumPrRef> = None;
    let mut depth = 0;

    loop {
        match reader.read_event()? {
            Event::Start(e) => {
                let local = local_name(e.name().as_ref());
                match local.as_str() {
                    "pPr" if depth == 0 => {
                        let (s, a, sid, npr) = parse_paragraph_properties(reader)?;
                        style = s;
                        alignment = a;
                        style_id = sid;
                        num_pr_ref = npr;
                    }
                    "r" if depth == 0 => {
                        let (run, img) = parse_run(reader, ctx)?;
                        runs.push(run);
                        if let Some(image) = img {
                            images.push(image);
                        }
                    }
                    _ => {
                        depth += 1;
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

    // Apply style inheritance from StyleSheet
    if let Some(ref sid) = style_id {
        if let Some(defined_style) = styles.styles.get(sid) {
            // Merge spacing from style definition (only if not explicitly set)
            if style.space_before.is_none() {
                style.space_before = defined_style.space_before;
            }
            if style.space_after.is_none() {
                style.space_after = defined_style.space_after;
            }
            if style.line_spacing.is_none() {
                style.line_spacing = defined_style.line_spacing;
            }
            // Carry over default run style from the style definition
            if style.default_run_style.is_none() {
                style.default_run_style = defined_style.default_run_style.clone();
            }
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

    Ok(Paragraph {
        runs,
        style,
        alignment,
    })
}

/// Numbering reference parsed from w:numPr
struct NumPrRef {
    num_id: String,
    ilvl: u8,
}

/// Parse w:pPr (paragraph properties)
fn parse_paragraph_properties(
    reader: &mut Reader<&[u8]>,
) -> Result<(ParagraphStyle, Alignment, Option<String>, Option<NumPrRef>), ParseError> {
    let mut style = ParagraphStyle::default();
    let mut alignment = Alignment::default();
    let mut style_id: Option<String> = None;
    let mut num_pr: Option<NumPrRef> = None;
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

    Ok((style, alignment, style_id, num_pr))
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

/// Parse a w:r element (run). Returns the Run and optionally an Image if a drawing was found.
fn parse_run(reader: &mut Reader<&[u8]>, ctx: &ParseContext) -> Result<(Run, Option<Image>), ParseError> {
    let mut text = String::new();
    let mut style = RunStyle::default();
    let mut image = None;
    let mut depth = 0;
    let mut in_text = false;

    loop {
        match reader.read_event()? {
            Event::Start(e) => {
                let local = local_name(e.name().as_ref());
                match local.as_str() {
                    "rPr" if depth == 0 => {
                        style = parse_run_properties(reader)?;
                    }
                    "t" if depth == 0 => {
                        in_text = true;
                    }
                    "drawing" if depth == 0 => {
                        image = parse_drawing(reader, ctx)?;
                    }
                    _ => {
                        depth += 1;
                    }
                }
            }
            Event::Text(e) => {
                if in_text {
                    text.push_str(&e.unescape().unwrap_or_default());
                }
            }
            Event::End(e) => {
                let local = local_name(e.name().as_ref());
                if local == "t" {
                    in_text = false;
                } else if local == "r" && depth == 0 {
                    break;
                } else if depth > 0 {
                    depth -= 1;
                }
            }
            Event::Empty(e) => {
                let local = local_name(e.name().as_ref());
                match local.as_str() {
                    "br" => text.push('\n'),
                    "tab" => text.push('\t'),
                    _ => {}
                }
            }
            Event::Eof => break,
            _ => {}
        }
    }

    Ok((Run { text, style }, image))
}

/// Parse a w:drawing element to extract image info
fn parse_drawing(reader: &mut Reader<&[u8]>, ctx: &ParseContext) -> Result<Option<Image>, ParseError> {
    let mut width: f32 = 0.0;
    let mut height: f32 = 0.0;
    let mut alt_text = None;
    let mut rel_id = None;
    let mut depth = 0;

    loop {
        match reader.read_event()? {
            Event::Start(e) => {
                let local = local_name(e.name().as_ref());
                depth += 1;
                match local.as_str() {
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
            Event::Empty(e) => {
                let local = local_name(e.name().as_ref());
                match local.as_str() {
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
        Ok(Some(Image {
            data,
            width,
            height,
            alt_text,
            content_type,
        }))
    } else {
        Ok(None)
    }
}

/// Parse w:rPr (run properties)
fn parse_run_properties(reader: &mut Reader<&[u8]>) -> Result<RunStyle, ParseError> {
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
                    "u" => style.underline = true,
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
                        for attr in e.attributes().flatten() {
                            if local_name(attr.key.as_ref()) == "val" {
                                style.color =
                                    Some(String::from_utf8_lossy(&attr.value).to_string());
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

    loop {
        match reader.read_event()? {
            Event::Start(e) => {
                let local = local_name(e.name().as_ref());
                if local == "tblBorders" {
                    style.border = true;
                }
                depth += 1;
            }
            Event::End(e) => {
                let local = local_name(e.name().as_ref());
                if local == "tblPr" && depth == 0 {
                    break;
                }
                if depth > 0 {
                    depth -= 1;
                }
            }
            Event::Empty(e) => {
                let local = local_name(e.name().as_ref());
                if matches!(
                    local.as_str(),
                    "top" | "left" | "bottom" | "right" | "insideH" | "insideV"
                ) {
                    style.border = true;
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

    Ok(TableRow { cells })
}

/// Parse a w:tc element (table cell)
fn parse_table_cell(reader: &mut Reader<&[u8]>, ctx: &ParseContext, styles: &StyleSheet) -> Result<TableCell, ParseError> {
    let mut blocks = Vec::new();
    let mut width = None;
    let mut depth = 0;

    loop {
        match reader.read_event()? {
            Event::Start(e) => {
                let local = local_name(e.name().as_ref());
                match local.as_str() {
                    "p" if depth == 0 => {
                        let para = parse_paragraph(reader, ctx, styles)?;
                        blocks.push(Block::Paragraph(para));
                    }
                    "tcPr" if depth == 0 => {
                        width = parse_cell_width(reader)?;
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

    Ok(TableCell { blocks, width })
}

/// Parse w:tcPr for cell width
fn parse_cell_width(reader: &mut Reader<&[u8]>) -> Result<Option<f32>, ParseError> {
    let mut width = None;
    let mut depth = 0;

    loop {
        match reader.read_event()? {
            Event::Start(_) => {
                depth += 1;
            }
            Event::Empty(e) => {
                let local = local_name(e.name().as_ref());
                if local == "tcW" {
                    for attr in e.attributes().flatten() {
                        if local_name(attr.key.as_ref()) == "w" {
                            let val = String::from_utf8_lossy(&attr.value);
                            width = val.parse::<f32>().ok().map(|v| v / 20.0);
                        }
                    }
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

    Ok(width)
}

/// Parsed section properties
struct SectionProperties {
    page_size: PageSize,
    margin: Margin,
    /// Document grid line pitch in points (from w:docGrid w:linePitch, twips/20)
    grid_line_pitch: Option<f32>,
}

/// Parse w:sectPr (section properties - page size, margins, document grid)
fn parse_section_properties(
    reader: &mut Reader<&[u8]>,
) -> Result<SectionProperties, ParseError> {
    let mut page_size = PageSize::default();
    let mut margin = Margin::default();
    let mut grid_line_pitch: Option<f32> = None;
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
                                _ => {}
                            }
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
    })
}

/// Extract local name from a potentially namespaced XML tag
fn local_name(name: &[u8]) -> String {
    let s = std::str::from_utf8(name).unwrap_or("");
    match s.rfind(':') {
        Some(pos) => s[pos + 1..].to_string(),
        None => s.to_string(),
    }
}
